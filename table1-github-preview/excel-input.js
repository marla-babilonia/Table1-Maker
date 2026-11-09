"use strict";

async function openExcelFile(bytes) {
    let zip;
    try {
        zip = await JSZip.loadAsync(bytes);
    } catch {
        throw new Error("This file could not be opened as an .xlsx workbook. Check that it is a valid, unencrypted Excel file.");
    }

    async function readXml(path, optional = false) {
        const entry = zip.file(path);
        if (!entry) {
            if (optional) return null;
            throw new Error("The Excel workbook is missing a required part.");
        }

        const text = await entry.async("string");
        if (/<!DOCTYPE/i.test(text)) throw new Error("Unsupported workbook XML.");
        const document = new DOMParser().parseFromString(text, "application/xml");
        if (document.getElementsByTagNameNS("*", "parsererror").length) {
            throw new Error("The Excel workbook contains invalid XML.");
        }
        return document.documentElement;
    }

    const children = (node, name) =>
        Array.from(node?.children ?? []).filter(child => child.localName === name);
    const child = (node, name) => children(node, name)[0];
    const attribute = (node, name) => node?.getAttribute(name);
    const richText = node => Array.from(node?.children ?? [])
        .map(part => part.localName === "t" ? part.textContent :
            part.localName === "r" ? child(part, "t")?.textContent ?? "" : "")
        .join("");

    const [workbook, relationships, strings, styles] = await Promise.all([
        readXml("xl/workbook.xml"),
        readXml("xl/_rels/workbook.xml.rels"),
        readXml("xl/sharedStrings.xml", true),
        readXml("xl/styles.xml", true),
    ]);

    const targets = new Map(children(relationships, "Relationship")
        .filter(rel => /\/worksheet$/.test(attribute(rel, "Type") ?? "") &&
            attribute(rel, "TargetMode") !== "External")
        .map(rel => [attribute(rel, "Id"), attribute(rel, "Target")]));

    const worksheets = children(child(workbook, "sheets"), "sheet").flatMap(sheet => {
        const id = Array.from(sheet.attributes).find(attr =>
            attr.localName === "id" && attr.namespaceURI?.endsWith("/relationships")
        )?.value;
        const target = targets.get(id);
        if (!target) return [];
        const path = target.startsWith("/") ? target.slice(1) : "xl/" + target;
        if (!/^xl\//.test(path) || path.split("/").includes("..")) {
            throw new Error("The workbook uses an unsupported worksheet path.");
        }
        return [{name: attribute(sheet, "name"), path}];
    });

    if (!worksheets.length) throw new Error("This workbook has no readable worksheets.");

    const sharedStrings = children(strings, "si").map(richText);
    const formats = new Map(children(child(styles, "numFmts"), "numFmt")
        .map(item => [Number(attribute(item, "numFmtId")), attribute(item, "formatCode")]));
    const cellStyles = children(child(styles, "cellXfs"), "xf");
    const date1904 = ["1", "true"].includes(attribute(child(workbook, "workbookPr"), "date1904"));

    function cellText(cell, reference) {
        const type = attribute(cell, "t");
        const value = child(cell, "v")?.textContent;
        if (type === "inlineStr") return richText(child(cell, "is"));
        if (type === "e") throw new Error(`Excel cell ${reference} contains an error (${value}). Correct it in Excel and try again.`);
        if (value === undefined || value === "") {
            if (child(cell, "f")) throw new Error(`Excel formula at ${reference} has no saved result. Recalculate and save the workbook in Excel.`);
            return "";
        }
        if (type === "s") {
            const text = sharedStrings[Number(value)];
            if (text === undefined) throw new Error(`Invalid shared text at ${reference}.`);
            return text;
        }
        if (type === "b") return value === "1" ? "TRUE" : "FALSE";
        if (type && type !== "n") return value;

        const style = cellStyles[Number(attribute(cell, "s") ?? 0)];
        const formatId = Number(attribute(style, "numFmtId") ?? 0);
        const format = formats.get(formatId) ?? "";
        if (/^0{2,}$/.test(format) && /^\d+$/.test(value)) {
            return value.padStart(format.length, "0");
        }

        const dateCode = format.replace(/"[^"]*"|\\.|\[[^\]]*\]/g, "");
        const isDate = (formatId >= 14 && formatId <= 22) ||
            (formatId >= 27 && formatId <= 36) ||
            (formatId >= 45 && formatId <= 47) ||
            (formatId >= 50 && formatId <= 58) ||
            /[ydhs]|m{2,}/i.test(dateCode) || /\[(?:h+|m+|s+)\]/i.test(format);
        if (!isDate) return value;

        const serial = Number(value);
        if (!Number.isFinite(serial)) return value;
        // Excel's 1900 calendar includes a fictitious leap day.
        if (!date1904 && Math.floor(serial) === 60) {
            throw new Error(`Excel cell ${reference} uses the fictitious date February 29, 1900.`);
        }
        const base = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31);
        const days = serial - (!date1904 && serial >= 61 ? 1 : 0);
        const date = new Date(base + Math.round(days * 86400000));
        if (!Number.isFinite(date.valueOf())) throw new Error(`Invalid Excel date at ${reference}.`);
        const iso = date.toISOString();
        const timeOnly = [18, 19, 20, 21, 45, 46, 47].includes(formatId) ||
            (/[hs]/i.test(dateCode) && !/[yd]/i.test(dateCode));
        if (timeOnly && serial >= 0 && serial < 1) return iso.slice(11, 19);
        return Number.isInteger(serial) && !timeOnly ? iso.slice(0, 10) : iso.replace(/\.000Z$/, "");
    }

    async function readSheet(index) {
        const sheet = worksheets[index];
        if (!sheet) throw new Error("Choose a worksheet first.");
        const xml = await readXml(sheet.path);
        const rows = [];
        let width = 0;
        for (const row of children(child(xml, "sheetData"), "row")) {
            const cells = new Map();
            let column = 0;
            for (const cell of children(row, "c")) {
                const ref = attribute(cell, "r");
                if (ref) {
                    const address = /^([A-Z]{1,3})[1-9]\d*$/i.exec(ref);
                    if (!address) throw new Error(`Invalid Excel cell address: ${ref}`);
                    column = [...address[1].toUpperCase()].reduce((n, letter) =>
                        n * 26 + letter.charCodeAt(0) - 64, 0) - 1;
                }
                if (column >= 16384) throw new Error("Invalid Excel column number.");
                const value = cellText(cell, ref ?? `in row ${attribute(row, "r") ?? "?"}`);
                if (value.trim()) {
                    cells.set(column, value);
                    width = Math.max(width, column + 1);
                }
                column++;
            }
            if (cells.size) rows.push(cells);
        }
        if (rows.length < 2) throw new Error("Include a header row and at least one data row.");
        if (width * rows.length > 2_000_000) {
            throw new Error("This sheet is too large for this version. Try a smaller worksheet.");
        }
        const names = Array.from({length: width}, (_, i) => (rows[0].get(i) ?? "").trim());
        if (names.some(name => !name) || new Set(names).size !== names.length) {
            throw new Error("Give every data column a unique, nonblank heading in the first nonblank row.");
        }
        const data = rows.slice(1).map(row => names.map((_, i) => row.get(i) ?? ""));
        return {names, data};
    }

    return {names: worksheets.map(sheet => sheet.name), readSheet};
}
