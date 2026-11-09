"use strict";

const NOTES = "Blank rows are excluded. Numerical summaries use nonmissing observations. Mean (sample SD); median [Q1, Q3]. Categorical percentages use records with a response for that variable. When a record can have multiple responses, it counts once per distinct response and percentages can sum to more than 100%.";
const SAMPLE = "record_id,age,score,group\n001,20,10,A\n002,30,20,B\n003,40,30,A\n004,50,40,B\n005,,50,A\n";

function parseCsv(input) {
    const text = input.replace(/^\uFEFF/, "");
    const records = [];
    let row = [], field = "", quoted = false, closed = false;
    function finishField() {
        row.push(field);
        field = "";
        closed = false;
    }
    function finishRow() {
        finishField();
        if (row.some(value => value.trim())) records.push(row);
        row = [];
    }
    for (let i = 0; i < text.length; i++) {
        const ch = text[i];
        if (quoted) {
            if (ch === '"' && text[i + 1] === '"') {
                field += '"';
                i++;
            } else if (ch === '"') {
                quoted = false;
                closed = true;
            } else {
                field += ch;
            }
        } else if (ch === ",") {
            finishField();
        } else if (ch === "\r" || ch === "\n") {
            finishRow();
            if (ch === "\r" && text[i + 1] === "\n") i++;
        } else if (ch === '"' && !field && !closed) {
            quoted = true;
        } else if (ch === '"' || closed) {
            throw new Error("Check the quotation marks in the CSV.");
        } else {
            field += ch;
        }
    }
    if (quoted) throw new Error("A quoted CSV field was left open.");
    finishRow();

    if (records.length < 2) throw new Error("Include a header and at least one data row.");
    const [header, ...data] = records;
    const names = header.map(value => value.trim());
    if (names.some(name => !name) || new Set(names).size !== names.length) {
        throw new Error("Give every column a unique, nonblank heading.");
    }
    if (data.some(record => record.length !== names.length)) {
        throw new Error("Every data row must have the same number of fields as the header.");
    }
    return {names, data};
}

function numeric(value) {
    const clean = value.trim();
    if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(clean)) return NaN;
    return Number(clean);
}

function inferMode(name, values) {
    if (/^(id|.*_id)$/i.test(name)) return "exclude";
    const observed = values.filter(value => value.trim()).map(numeric);
    return observed.length && observed.every(Number.isFinite) &&
        observed.some(value => value !== 0 && value !== 1) ? "mean" : "categorical";
}

function buildTable(dataset, modes) {
    const rows = [["Number of records", String(dataset.data.length), ""]];
    const stats = [["Variable", "n", "Missing", "Mean", "Sample SD", "Median", "Q1", "Q3", "IQR", "Min", "Max"]];
    const fmt = value => Number.isFinite(value) ? value.toFixed(2) : "NA";

    dataset.names.forEach((name, col) => {
        const mode = modes[col];
        if (mode === "exclude") return;
        const values = dataset.data.map(record => record[col]).filter(value => value.trim());
        const missing = dataset.data.length - values.length;
        if (mode === "categorical") {
            rows.push([name, "", "heading"]);
            const counts = new Map();
            for (const value of values) counts.set(value, (counts.get(value) || 0) + 1);
            for (const [value, count] of counts) {
                rows.push([value, `${count} (${(100 * count / values.length).toFixed(1)}%)`, "indent"]);
            }
            if (missing) rows.push(["Missing", String(missing), "indent"]);
            return;
        }
        const nums = values.map(numeric).sort((a, b) => a - b);
        if (nums.some(value => !Number.isFinite(value))) {
            throw new Error(`“${name}” contains nonnumeric values. Select Categorical or Exclude.`);
        }
        const n = nums.length;
        const mean = n ? nums.reduce((sum, v) => sum + v, 0) / n : NaN;
        const sd = n > 1 ? Math.sqrt(nums.reduce((sum, v) => sum + (v - mean) ** 2, 0) / (n - 1)) : NaN;
        const quantile = p => {
            if (!n) return NaN;
            const pos = p * (n - 1);
            const start = Math.floor(pos);
            return nums[start] + (nums[Math.ceil(pos)] - nums[start]) * (pos - start);
        };
        const q1 = quantile(.25), median = quantile(.5), q3 = quantile(.75);
        stats.push([name, n, missing, mean, sd, median, q1, q3, q3 - q1, nums[0], nums[n - 1]]);
        rows.push([name, mode === "mean" ? `${fmt(mean)} (${fmt(sd)})` :
            `${fmt(median)} [${fmt(q1)}, ${fmt(q3)}]`, ""]);
        if (missing) rows.push([`${name}: missing`, String(missing), "indent"]);
    });
    return {rows, stats};
}

const xml = value => String(value).replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, "")
    .replace(/[&<>]/g, ch => ({"&":"&amp;","<":"&lt;",">":"&gt;"}[ch]));

async function excelBytes(title, result) {
    const zip = new JSZip();
    const ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const pkg = "http://schemas.openxmlformats.org/package/2006/relationships";
    const types = "application/vnd.openxmlformats-officedocument.spreadsheetml";
    const showDetails = result.display.includeDetails;

    const detailsPart = showDetails
        ? `<Override PartName="/xl/worksheets/sheet2.xml" ContentType="${types}.worksheet+xml"/>`
        : "";

    const detailsSheet = showDetails
        ? '<sheet name="Complete Statistics" sheetId="2" r:id="rId2"/>'
        : "";

    const detailsRelationship = showDetails
        ? `<Relationship Id="rId2" Type="${rel}/worksheet" Target="worksheets/sheet2.xml"/>`
        : "";

    zip.file("[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="${types}.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="${types}.worksheet+xml"/>${detailsPart}<Override PartName="/xl/styles.xml" ContentType="${types}.styles+xml"/></Types>`);
    zip.file("_rels/.rels", `<Relationships xmlns="${pkg}"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
    zip.file("xl/workbook.xml", `<workbook xmlns="${ns}" xmlns:r="${rel}"><sheets><sheet name="Table 1" sheetId="1" r:id="rId1"/>${detailsSheet}</sheets></workbook>`);
    zip.file("xl/_rels/workbook.xml.rels", `<Relationships xmlns="${pkg}"><Relationship Id="rId1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/>${detailsRelationship}<Relationship Id="rId3" Type="${rel}/styles" Target="styles.xml"/></Relationships>`);
    zip.file("xl/styles.xml", `<styleSheet xmlns="${ns}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFEAF6FC"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="9"><xf/><xf fontId="1" fillId="2" applyFont="1" applyFill="1"/><xf fontId="1" applyFont="1"/><xf applyAlignment="1"><alignment indent="2"/></xf><xf fillId="2" applyFill="1"/><xf applyAlignment="1"><alignment indent="3"/></xf><xf applyAlignment="1"><alignment indent="1"/></xf><xf fontId="1" applyFont="1" applyAlignment="1"><alignment indent="1"/></xf><xf applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`);
    function excelColumn(index) {
        let name = "";
        for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) {
            name = String.fromCharCode(65 + (n - 1) % 26) + name;
        }
        return name;
    }
    function sheet(data, cols, freeze, footer) {
        const noteHeight = footer ? Math.min(409, 16 * footer.split(/\r?\n/).reduce(
            (count, line) => count + Math.max(1, Math.ceil(line.length / 54)), 0
        ) + 8) : 0;
        const mergedFooter = footer
            ? `<mergeCells count="1"><mergeCell ref="A${data.length}:${excelColumn(cols.at(-1)[1] - 1)}${data.length}"/></mergeCells>`
            : "";
        return `<worksheet xmlns="${ns}"><sheetViews><sheetView workbookViewId="0"><pane ySplit="${freeze}" topLeftCell="A${freeze + 1}" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>${cols.map(([min,max,width]) => `<col min="${min}" max="${max}" width="${width}" customWidth="1"/>`).join("")}</cols><sheetData>${data.map(([cells, style, firstOnly], i) => `<row r="${i + 1}"${footer && i === data.length - 1 ? ` ht="${noteHeight}" customHeight="1"` : ""}>${cells.map((value, j) => {
            const address = excelColumn(j) + (i + 1);
            const s = firstOnly && j ? 0 : style;
            if (Number.isFinite(value)) return `<c r="${address}" s="${s}"><v>${value}</v></c>`;
            return `<c r="${address}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${xml(value ?? "")}</t></is></c>`;
        }).join("")}</row>`).join("")}</sheetData>${mergedFooter}</worksheet>`;
    }
    const headingStyle = result.display.boldHeadings ? 2 : 0;
    const columnStyle = result.display.boldHeadings ? 1 : 4;
    const rows = [
        [[title], headingStyle],
        [result.headings, columnStyle],
        ...result.rows.map(([label, value, kind]) => [
            [label, ...value],
            kind === "heading" ? headingStyle
                : kind === "subheading" ? (result.display.boldHeadings ? 7 : 6)
                : kind === "deep-indent" ? 5 : kind === "indent" ? 3 : 0,
            true,
        ]),
    ];

    if (result.display.showNotes) rows.push([[result.display.note], 8]);

    zip.file("xl/worksheets/sheet1.xml", sheet(rows, [[1, 1, 40], [2, result.headings.length, 24]], 2,
        result.display.showNotes ? result.display.note : ""));

    if (showDetails) {
        const detailRows = result.details.map((row, index) => [
            row.map(value =>
                typeof value === "number" && !Number.isFinite(value) ? null : value
            ),
            index === 0 ? columnStyle : 0,
        ]);
        zip.file("xl/worksheets/sheet2.xml", sheet(detailRows, [
            [1, 3, 27], [4, 4, 16], [5, 5, 27],
            [6, 9, 16], [10, 10, 28], [11, 18, 15],
        ], 1));
    }

    return zip.generateAsync({type: "uint8array"});
}

async function pdfBytes(title, result) {
    const {PDFDocument, StandardFonts, rgb} = PDFLib;
    const doc = await PDFDocument.create();
    const normal = await doc.embedFont(StandardFonts.Helvetica);
    const headingFont = result.display.boldHeadings
        ? await doc.embedFont(StandardFonts.HelveticaBold) : normal;
    const ink = rgb(.09, .21, .29);

    function wrap(value, width, size, font) {
        const lines = [];
        let line = "";
        for (const char of String(value).replace(/\s+/g, " ")) {
            if (line && font.widthOfTextAtSize(line + char, size) > width) {
                lines.push(line);
                line = "";
            }
            line += char;
        }
        lines.push(line);
        return lines;
    }

    const strataCount = result.headings.length - 2;
    const panels = Math.max(1, Math.ceil(strataCount / 4));

    for (let panel = 0; panel < panels; panel++) {
        const start = panel * 4;
        const headings = [...result.headings.slice(0, 2),
            ...result.headings.slice(2 + start, 6 + start)];
        const wide = strataCount > 0;
        const [width, height] = wide ? [792, 612] : [612, 792];
        const margin = 40;
        const available = width - 2 * margin;
        const firstWidth = wide ? 192 : 300;
        const otherWidth = (available - firstWidth) / (headings.length - 1);
        const xFor = i => margin + (i ? firstWidth + (i - 1) * otherWidth : 0);
        let page, y;

        function draw(value, x, at, size, font) {
            page.drawText(value, {x, y: at, size, font, color: ink});
        }

        function newPage() {
            page = doc.addPage([width, height]);
            y = height - margin;
            for (const line of wrap(title, available, 17, headingFont)) {
                draw(line, margin, y, 17, headingFont);
                y -= 20;
            }
            if (panels > 1) {
                draw(`Stratum columns ${start + 1}–${Math.min(start + 4, strataCount)} of ${strataCount}`,
                    margin, y, 9, normal);
                y -= 15;
            }
            const lines = headings.map((label, i) =>
                wrap(label, (i ? otherWidth : firstWidth) - 12, 10, headingFont));
            const blockHeight = 12 * Math.max(...lines.map(item => item.length)) + 10;
            page.drawRectangle({x: margin, y: y + 5 - blockHeight,
                width: available, height: blockHeight, color: rgb(.92, .96, .99)});
            lines.forEach((items, i) => items.forEach((line, j) =>
                draw(line, xFor(i) + 6, y - 12 * j - 5, 10, headingFont)));
            y -= blockHeight + 9;
            if (result.display.showNotes) {
                draw(`Page ${doc.getPageCount()}`, width - 85, 24, 9, normal);
            }
        }

        newPage();
        for (const [label, values, kind] of result.rows) {
            const font = ["heading", "subheading"].includes(kind) ? headingFont : normal;
            const indent = {"subheading": 14, "indent": 28, "deep-indent": 42}[kind] ?? 0;
            const cells = [label, values[0], ...values.slice(1 + start, 5 + start)];
            const lines = cells.map((value, i) => wrap(value,
                (i ? otherWidth : firstWidth - indent) - 12, 10, i ? normal : font));
            for (let j = 0; j < Math.max(...lines.map(item => item.length)); j++) {
                if (y < 55) newPage();
                lines.forEach((items, i) =>
                    draw(items[j] ?? "", xFor(i) + 6 + (i ? 0 : indent), y, 10, i ? normal : font));
                y -= 13;
            }
            y -= 5;
        }

        if (result.display.showNotes) {
            y -= 10;
            for (const line of result.display.note.split(/\r?\n/).flatMap(line => wrap(line, available, 9, normal))) {
                if (y < 55) newPage();
                draw(line, margin, y, 9, normal);
                y -= 12;
            }
        }
    }
    return doc.save();
}

const EXTRA_STATS = [
    ["mean", "Mean"],
    ["min", "Minimum"],
    ["max", "Maximum"],
    ["median", "Median"],
    ["q1", "25th percentile"],
    ["q3", "75th percentile"],
    ["sd", "Sample SD"],
];

const SEPARATORS = {
    single: null,
    semicolon: ";",
    pipe: "|",
    comma: ",",
};

function parseStudyNumber(value, format) {
    let text = value.trim();

    if (format === "decimal-comma") {
        // Examples: 1,25 or -0,5. Grouped numbers need correction.
        if (!/^[+-]?(?:\d+(?:,\d+)?|,\d+)(?:e[+-]?\d+)?$/i.test(text)) {
            return NaN;
        }
        text = text.replace(",", ".");
    }

    if (format === "comma-thousands" && text.includes(",")) {
        // Validate group sizes before removing commas.
        if (!/^[+-]?\d{1,3}(?:,\d{3})+(?:\.\d+)?$/.test(text)) {
            return NaN;
        }
        text = text.replaceAll(",", "");
    }

    return numeric(text);
}

function numericReview(dataset, variable) {
    const values = [];
    const invalid = [];
    let missing = 0;

    dataset.data.forEach((record, rowIndex) => {
        const original = record[variable.column];
        const value = variable.corrections.get(dataset.recordIndices?.[rowIndex] ?? rowIndex) ?? original;

        if (!value.trim()) {
            missing++;
            return;
        }

        const number = parseStudyNumber(value, variable.numberFormat);

        if (Number.isFinite(number)) {
            values.push(number);
        } else {
            invalid.push({rowIndex, original, value});
        }
    });

    return {values, invalid, missing};
}

function categoryRecoder(variable, mode = variable.separator) {
    const separator = SEPARATORS[mode];
    const mapping = new Map(mode === variable.separator
        ? (variable.categoryMerges ?? []).flatMap(rule => rule.values.map(value => [value, rule.label]))
        : []);
    return cell => new Set((separator ? cell.split(separator) : [cell])
        .map(value => value.trim()).filter(Boolean)
        .map(value => mapping.get(value) ?? value));
}

function categoricalReview(dataset, variable) {
    const recode = categoryRecoder(variable);
    const counts = new Map();
    let denominator = 0;

    for (const record of dataset.data) {
        const responses = recode(record[variable.column]);

        if (!responses.size) continue;

        denominator++;

        for (const response of responses) {
            counts.set(response, (counts.get(response) ?? 0) + 1);
        }
    }

    return {
        counts,
        denominator,
        missing: dataset.data.length - denominator,
    };
}

function summarizeNumbers(values) {
    const ordered = [...values].sort((a, b) => a - b);
    const n = ordered.length;
    const mean = n ? ordered.reduce((sum, value) => sum + value, 0) / n : NaN;
    const sd = n > 1
        ? Math.sqrt(ordered.reduce((sum, value) => sum + (value - mean) ** 2, 0) / (n - 1))
        : NaN;

    function quantile(fraction) {
        if (!n) return NaN;
        const position = fraction * (n - 1);
        const lower = Math.floor(position);
        const upper = Math.ceil(position);
        return ordered[lower] + (ordered[upper] - ordered[lower]) * (position - lower);
    }

    const q1 = quantile(0.25);
    const q3 = quantile(0.75);

    return {
        n, mean, sd, q1, q3,
        median: quantile(0.5),
        iqr: q3 - q1,
        min: ordered[0] ?? NaN,
        max: ordered[n - 1] ?? NaN,
    };
}

function buildWizardTable(dataset, variables, display, reference) {
    const rows = [];
    const total = dataset.data.length;

    const details = [[
        "Source column", "Table label", "Type", "Category",
        "Total records", "Nonmissing", "Missing", "Category count",
        "Category % of nonmissing", "Mean", "Sample SD", "Median",
        "Q1", "Q3", "IQR", "Min", "Max",
    ]];

    const format = value =>
        Number.isFinite(value) ? value.toFixed(2) : "NA";

    if (display.showTotal) {
        rows.push(["Number of records", String(total), "", JSON.stringify(["total"])]);
    }

    for (const variable of variables) {
        const full = reference?.get(variable.column);
        const addRow = (label, value, kind, type, key = "") => {
            rows.push([label, value, kind, JSON.stringify([type, variable.column, key])]);
        };
        if (variable.type === "numeric") {
            const {values, invalid, missing} = numericReview(dataset, variable);

            if (invalid.length) {
                throw new Error(`Correct the unresolved values in ${variable.source}.`);
            }

            const s = summarizeNumbers(values);

            const summary = variable.summary === "median"
                ? `${format(s.median)} [${format(s.q1)}, ${format(s.q3)}]`
                : `${format(s.mean)} (${format(s.sd)})`;

            addRow(variable.label, summary, "", "numeric");

            if (display.includeDetails) {
                details.push([
                    variable.source, variable.label, "Numerical", "",
                    total, s.n, missing, null, null,
                    s.mean, s.sd, s.median, s.q1, s.q3, s.iqr, s.min, s.max,
                ]);
            }

            for (const [key, label] of EXTRA_STATS) {
                if (variable.extras.has(key)) {
                    addRow(label, format(s[key]), "indent", "statistic", key);
                }
            }

            if (display.showMissing && (full?.missing ?? missing)) {
                addRow("Missing", String(missing), "indent", "missing");
            }

            continue;
        }

        const {counts, denominator, missing} = categoricalReview(dataset, variable);

        if (variable.showHeader) {
            addRow(variable.label, "", "heading", "variable");
        }

        if (display.includeDetails && !(full?.counts ?? counts).size) {
            details.push([
                variable.source, variable.label, "Categorical", "",
                total, denominator, missing, null, null,
            ]);
        }

        let categoryIndent = variable.showHeader ? "indent" : "";
        for (const value of (full?.counts ?? counts).keys()) {
            const count = counts.get(value) ?? 0;
            const label = variable.levelLabels.get(value) ?? value;
            const percent = denominator ? 100 * count / denominator : NaN;

            if (display.includeDetails) {
                details.push([
                    variable.source, variable.label, "Categorical", label,
                    total, denominator, missing, count, percent,
                ]);
            }

            if (variable.hiddenLevels.has(value)) continue;

            const heading = variable.headingsBefore?.get(value);
            if (heading?.label.trim()) {
                addRow(heading.label.trim(), "", heading.kind, "heading", value);
                categoryIndent = heading.kind === "subheading" ? "deep-indent" : "indent";
            }

            addRow(
                label,
                `${count} (${Number.isFinite(percent) ? percent.toFixed(1) + "%" : "NA"})`,
                categoryIndent, "category", value
            );
        }

        if (display.showMissing && (full?.missing ?? missing)) {
            addRow(
                variable.showHeader ? "Missing" : `${variable.label}: missing`,
                String(missing),
                variable.showHeader ? "indent" : "", "missing"
            );
        }
    }

    if (!rows.length) {
        throw new Error("The table is empty. Include a category, show a heading, or turn on missing counts.");
    }

    return {rows, details, display: {...display}};
}

function buildStrata(dataset, variable, settings) {
    const groups = new Map();
    const add = (key, label) => {
        if (!groups.has(key)) groups.set(key, {key, label, indices: []});
        return groups.get(key);
    };
    const missing = () => add("missing", "(Missing)");

    if (settings.mode === "numeric") {
        const tokens = settings.cuts.split(";").map(text => text.trim());
        const cuts = tokens.map(text => parseStudyNumber(text, settings.numberFormat));
        if (cuts.some((value, i) => !Number.isFinite(value) || (i && value <= cuts[i - 1]))) {
            throw new Error("Enter one or more increasing cut points separated by semicolons.");
        }

        const bins = Array.from({length: cuts.length + 1}, (_, i) => {
            const label = i === 0 ? `< ${cuts[0]}`
                : i === cuts.length ? `${cuts[i - 1]} or higher`
                : `${cuts[i - 1]} to < ${cuts[i]}`;
            return add("bin:" + label, label);
        });

        dataset.data.forEach((record, i) => {
            const text = variable.corrections.get(i) ?? record[variable.column];
            if (!text.trim()) {
                missing().indices.push(i);
                return;
            }
            const value = parseStudyNumber(text, settings.numberFormat);
            if (!Number.isFinite(value)) {
                throw new Error(`Record ${i + 1} in ${variable.source} is nonnumeric. Change its number format, correct it in the numerical review, or stratify by category.`);
            }
            const upper = cuts.findIndex(cut => value < cut);
            bins[upper === -1 ? cuts.length : upper].indices.push(i);
        });
    } else {
        const recode = categoryRecoder(variable, settings.separator);
        dataset.data.forEach((record, i) => {
            const distinct = recode(record[variable.column]);
            if (!distinct.size) missing().indices.push(i);
            for (const value of distinct) add("category:" + value, value).indices.push(i);
        });
    }

    return [...groups.values()];
}

function buildStratifiedTable(dataset, variables, display, headings, strata, cohortMode) {
    function subset(indices) {
        return {names: dataset.names, data: indices.map(i => dataset.data[i]), recordIndices: indices};
    }
    const selected = cohortMode === "shown" && strata.length
        ? [...new Set(strata.flatMap(group => group.indices))].sort((a, b) => a - b)
        : null;
    const overall = selected ? subset(selected) : dataset;
    const reference = new Map(variables.map(variable => [
        variable.column,
        variable.type === "numeric"
            ? {missing: numericReview(dataset, variable).missing}
            : categoricalReview(dataset, variable),
    ]));

    const cohort = buildWizardTable(overall, variables, display, reference);
    const tables = strata.map(group =>
        buildWizardTable(subset(group.indices), variables, display, reference)
    );

    if (tables.some(table => table.rows.length !== cohort.rows.length)) {
        throw new Error("The group summaries could not be aligned.");
    }

    return {
        display: {...display},
        headings: [...headings, ...strata.map(group => group.label)],
        rows: cohort.rows.map(([label, value, kind, id], i) => [
            label, [value, ...tables.map(table => table.rows[i][1])], kind, id,
        ]),
        details: [
            ["Cohort", ...cohort.details[0]],
            ...[cohort, ...tables].flatMap((table, i) =>
                table.details.slice(1).map(row => [
                    i ? strata[i - 1].label : headings[1], ...row,
                ])
            ),
        ],
    };
}

function mergeCategoryValues(values) {
    const present = values.map(value => value.trim()).filter(Boolean);
    return present.map(value => /[+"\\]/.test(value) ? JSON.stringify(value) : value).join(" + ");
}

function orderTableRows(rows, order) {
    const rank = new Map(order.map((id, index) => [id, index]));
    const known = rows.filter(row => rank.has(row[3]))
        .sort((left, right) => rank.get(left[3]) - rank.get(right[3]));
    let next = 0;
    // New rows stay in their default positions while existing rows keep their relative order.
    return rows.map(row => rank.has(row[3]) ? known[next++] : row);
}
