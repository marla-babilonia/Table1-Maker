"use strict";

const $ = id => document.getElementById(id);

const STEPS = [
    ["Filter the cohort",
        "Keep all records or add filters. Every record must meet all filters to be included."],
    ["Identify ID columns",
        "Check any columns used as identifiers. Leave them unchecked if there are none."],
    ["Choose variables to include",
        "Select the columns you want to summarize and confirm their types."],
    ["Review numerical variables",
        "Choose a summary and any extra statistics. Review values the app cannot read as numbers."],
    ["Review categorical variables",
        "Choose how to interpret each cell, then select which categories to display."],
    ["Stratify the table",
        "Choose a grouping variable or keep one overall column. Excluded strata still count in All cohort."],
    ["Edit names",
        "Enter the table title and edit variable names or individual category labels."],
];

let dataset;
let sourceDataset;
let sourceName;
let filters = [];
let cohortIndices;
let variables = [];
let current;
let step = 0;
let title = "Table 1";
let loadVersion = 0;
let rowOrder = [];

const DEFAULT_DISPLAY = {
    showTotal: false,
    showMissing: false,
    showNotes: false,
    note: NOTES,
    boldHeadings: false,
    includeDetails: false,
};

let display = {...DEFAULT_DISPLAY};
let strataSettings = {column: ""};
let columnLabels = {characteristic: "Characteristic", overall: ""};

function element(tag, text = "", className = "") {
    const node = document.createElement(tag);
    node.textContent = text;
    if (className) node.className = className;
    return node;
}

function checkbox(parent, caption, checked, onChange) {
    const label = element("label", "", "check-row");
    const input = element("input");
    input.type = "checkbox";
    input.checked = checked;
    input.addEventListener("change", () => onChange(input.checked));
    label.append(input, document.createTextNode(caption));
    parent.append(label);
    return input;
}

function selectField(parent, caption, options, selected, onChange) {
    const label = element("label", caption, "wizard-field");
    const input = element("select");

    for (const [value, text] of options) {
        input.add(new Option(text, value));
    }

    input.value = selected;
    input.addEventListener("change", () => onChange(input.value));
    label.append(input);
    parent.append(label);
    return input;
}

function textField(parent, caption, value, onChange) {
    const label = element("label", caption, "wizard-field");
    const input = element("input");
    input.type = "text";
    input.value = value;
    input.addEventListener("input", () => onChange(input.value));
    label.append(input);
    parent.append(label);
    return input;
}

function card(parent, heading) {
    const box = element("div", "", "variable-card");
    box.append(element("h3", heading));
    parent.append(box);
    return box;
}

function canUseVariable(variable) {
    return !variable.identifier &&
        !variable.mergedFrom?.some(column => variables[column].identifier);
}

function included() {
    return variables.filter(variable => variable.included && canUseVariable(variable));
}

function makeVariable(source, column, type) {
    return {
        source, column, label: source, type,
        identifier: false, included: true,
        summary: "mean", numberFormat: "decimal-point", separator: "single",
        showHeader: true, extras: new Set(), hiddenLevels: new Set(),
        levelLabels: new Map(), corrections: new Map(), headingsBefore: new Map(), categoryMerges: [],
    };
}

function reset() {
    loadVersion++;
    rowOrder = [];
    $("row-order")?.remove();
    dataset = current = sourceDataset = cohortIndices = undefined;
    filters = [];
    variables = [];
    step = 0;
    title = "Table 1";
    display = {...DEFAULT_DISPLAY};
    strataSettings = {column: ""};
    columnLabels = {characteristic: "Characteristic", overall: ""};
    $("file").value = "";
    $("worksheet-picker").replaceChildren();
    $("status").textContent = "";
    $("error").textContent = "";
    $("step-error").textContent = "";
    $("step-content").replaceChildren();
    $("rows").replaceChildren();
    $("caption").textContent = "";
    $("settings").hidden = true;
    $("results").hidden = true;
    $("clear").hidden = true;
}

function loadData(text, name) {
    loadDataset(parseCsv(text), name);
}

function loadDataset(input, name) {
    sourceDataset = input;
    sourceName = name;
    $("clear").hidden = false;
    $("status").textContent = `${name} · ${input.data.length} records · ${input.names.length} columns`;
    renderStep();
}

function matchingCohort() {
    const tests = filters.map(rule => {
        if (rule.column === "") throw new Error("Choose a column for each filter.");
        const column = Number(rule.column);
        if (rule.mode === "nonempty") return row => Boolean(row[column].trim());
        if (rule.mode === "empty") return row => !row[column].trim();

        const wanted = new Set(rule.values.split(/\r?\n/).map(value => value.trim()).filter(Boolean));
        if (!wanted.size) throw new Error("Enter at least one value for each matching filter.");
        const separator = SEPARATORS[rule.separator];
        return row => {
            const cell = row[column];
            const answers = separator ? cell.split(separator) : [cell];
            return answers.some(answer => wanted.has(answer.trim()));
        };
    });
    const indices = [];
    sourceDataset.data.forEach((row, index) => {
        if (tests.every(test => test(row))) indices.push(index);
    });
    return indices;
}

function renderCohortFilter(parent) {
    const count = element("p");
    count.setAttribute("role", "status");
    parent.append(count);

    function updateCount() {
        try {
            count.textContent = `${matchingCohort().length} of ${sourceDataset.data.length} records match.`;
        } catch (error) {
            count.textContent = error.message;
        }
    }

    filters.forEach((rule, index) => {
        const box = card(parent, `Filter ${index + 1}`);
        selectField(box, "Column (including IDs)", [
            ["", "Choose a column"],
            ...sourceDataset.names.map((name, i) => [String(i), name]),
        ], rule.column, value => {
            rule.column = value;
            updateCount();
        });
        selectField(box, "Keep rows where the value", [
            ["match", "Matches any listed value"],
            ["nonempty", "Is not empty"],
            ["empty", "Is empty"],
        ], rule.mode, value => {
            rule.mode = value;
            renderStep();
        });
        if (rule.mode === "match") {
            const label = element("label", "Accepted values, one per line (case-sensitive)", "wizard-field");
            const input = element("textarea");
            input.rows = 4;
            input.value = rule.values;
            input.addEventListener("input", () => {
                rule.values = input.value;
                updateCount();
            });
            label.append(input);
            box.append(label);
            selectField(box, "How to read each cell", [
                ["single", "Match the whole cell"],
                ["comma", "Match any answer separated by commas"],
                ["semicolon", "Match any answer separated by semicolons"],
                ["pipe", "Match any answer separated by vertical bars"],
            ], rule.separator, value => {
                rule.separator = value;
                updateCount();
            });
        }
        const remove = element("button", "Remove filter");
        remove.type = "button";
        remove.addEventListener("click", () => {
            filters.splice(index, 1);
            renderStep();
        });
        box.append(remove);
    });

    const add = element("button", "Add filter");
    add.type = "button";
    add.addEventListener("click", () => {
        filters.push({column: "", mode: "match", values: "", separator: "single"});
        renderStep();
    });
    parent.append(add);
    if (cohortIndices) {
        parent.append(element("p", "Changing which records are included resets your table selections."));
    }
    updateCount();
}

function applyCohortFilter() {
    const indices = matchingCohort();
    if (!indices.length) throw new Error("No records match. Edit or remove a filter to continue.");
    if (cohortIndices?.length === indices.length && indices.every((value, i) => value === cohortIndices[i])) return;

    cohortIndices = indices;
    dataset = {
        names: [...sourceDataset.names],
        data: indices.map(index => [...sourceDataset.data[index]]),
    };
    rowOrder = [];
    title = "Table 1";
    display = {...DEFAULT_DISPLAY};
    strataSettings = {column: ""};
    columnLabels = {characteristic: "Characteristic", overall: ""};
    variables = dataset.names.map((source, column) => {
        const observed = dataset.data.map(row => row[column]);
        // Infer from values only. The user marks identifiers separately.
        const guessed = inferMode("", observed);

        return makeVariable(source, column, guessed === "mean" ? "numeric" : "categorical");
    });

    $("clear").hidden = false;
    $("status").textContent =
        `${sourceName} · ${dataset.data.length} of ${sourceDataset.data.length} records in cohort · ${variables.length} columns`;
}

function renderIdentifiers(parent) {
    for (const variable of variables.filter(item => !item.mergedFrom)) {
        checkbox(parent, variable.source, variable.identifier, checked => {
            variable.identifier = checked;
        });
    }
}

function renderMerge(parent, sources) {
    if (sources.length < 2) return;
    const box = card(parent, "Merge two columns");
    box.append(element("p",
        "Choose two columns and keep all values or select specific ones from each. The new column combines the retained values within each record."
    ));
    const choices = [["", "Choose a column"],
        ...sources.map(variable => [String(variable.column), variable.source])];
    const first = sourcePicker("First column");
    const second = sourcePicker("Second column");

    function sourcePicker(caption) {
        const block = element("div");
        box.append(block);
        const picker = selectField(block, caption, choices, "", renderValues);
        const options = element("div");
        block.append(options);
        let selected = null;

        function renderValues() {
            options.replaceChildren();
            selected = null;
            if (!picker.value) return;
            const variable = variables[Number(picker.value)];
            const {counts} = categoricalReview(dataset, variable);
            const levels = element("div");
            levels.hidden = true;
            selectField(options, `Values from ${variable.source}`, [
                ["all", "All values"], ["choose", "Choose values"],
            ], "all", mode => {
                levels.hidden = mode === "all";
                selected = mode === "all" ? null : new Set();
                if (selected) {
                    for (const [value, check] of checks) {
                        if (check.checked) selected.add(value);
                    }
                }
            });
            const checks = [...counts.keys()].map(value => [
                value, checkbox(levels, value, false, checked => {
                    if (!selected) return;
                    if (checked) selected.add(value);
                    else selected.delete(value);
                }),
            ]);
            options.append(levels);
        }

        return {picker, selected: () => selected};
    }
    const name = textField(box, "New variable name", "", () => {});
    const button = element("button", "Create merged variable");
    button.type = "button";
    box.append(button);

    button.addEventListener("click", () => {
        const label = name.value.trim();
        if (!first.picker.value || !second.picker.value || first.picker.value === second.picker.value) {
            $("step-error").textContent = "Choose two different columns.";
            return;
        }
        if (!label || variables.some(variable => variable.source.toLowerCase() === label.toLowerCase())) {
            $("step-error").textContent = "Give the merged variable a new, unique name.";
            return;
        }

        const selections = [first.selected(), second.selected()];
        if (selections.some(selection => selection && !selection.size)) {
            $("step-error").textContent = "Choose at least one value from each column, or select All values.";
            return;
        }
        const columns = [Number(first.picker.value), Number(second.picker.value)];
        const recoders = columns.map(i => categoryRecoder(variables[i]));
        const column = dataset.names.length;
        dataset.names.push(label);
        for (const record of dataset.data) {
            const values = columns.flatMap((source, i) =>
                [...recoders[i](record[source])].filter(value => !selections[i] || selections[i].has(value))
            );
            record.push(mergeCategoryValues(values));
        }
        variables.push({...makeVariable(label, column, "categorical"), mergedFrom: columns});
        renderStep();
    });
}

function renderIncluded(parent) {
    const available = variables.filter(canUseVariable);

    if (!available.length) {
        parent.append(element("p", "All columns are marked as identifiers. Go back to include a data column."));
        return;
    }

    for (const variable of available) {
        const row = element("div", "", "variable");
        const check = checkbox(row, variable.source, variable.included, checked => {
            variable.included = checked;
            type.disabled = !checked;
        });
        const type = selectField(row, `Type for ${variable.source}`, [
            ["numeric", "Numerical"],
            ["categorical", "Categorical"],
        ], variable.type, value => {
            variable.type = value;
        });

        type.disabled = !check.checked;
        parent.append(row);
    }

    renderMerge(parent, available.filter(variable => !variable.mergedFrom));
}

function renderNumerical(parent) {
    const numericVariables = included().filter(variable => variable.type === "numeric");

    if (!numericVariables.length) {
        parent.append(element("p", "There are no numerical variables to review."));
        return;
    }

    for (const variable of numericVariables) {
        const box = card(parent, variable.source);

        selectField(box, "Main summary", [
            ["mean", "Mean (SD)"],
            ["median", "Median [Q1, Q3]"],
        ], variable.summary, value => {
            variable.summary = value;
        });

        selectField(box, "Number format", [
            ["decimal-point", "Decimal point, e.g. 1.25"],
            ["decimal-comma", "Decimal comma, e.g. 1,25"],
            ["comma-thousands", "Comma thousands separator, e.g. 1,250.5"],
        ], variable.numberFormat, value => {
            variable.numberFormat = value;
            renderCorrections();
        });

        box.append(element("p", "Optional statistics shown as indented rows"));

        for (const [key, label] of EXTRA_STATS) {
            checkbox(box, label, variable.extras.has(key), checked => {
                if (checked) variable.extras.add(key);
                else variable.extras.delete(key);
            });
        }

        const corrections = element("div");
        box.append(corrections);
        renderCorrections();

        function renderCorrections() {
            corrections.replaceChildren();

            const {invalid} = numericReview(dataset, variable);
            const unresolved = new Set(invalid.map(item => item.rowIndex));

            // Keep edited rows visible so users can review or undo a fix.
            const toShow = [...new Set([
                ...unresolved,
                ...variable.corrections.keys(),
            ])].sort((a, b) => a - b);

            if (!toShow.length) return;

            corrections.append(element("p",
                `${invalid.length} unresolved value(s). Enter a number or explicitly mark the value as missing.`
            ));

            for (const rowIndex of toShow) {
                const row = element("div", "", "correction");
                const original = dataset.data[rowIndex][variable.column];
                const corrected = variable.corrections.get(rowIndex);
                const idText = variables
                    .filter(item => item.identifier)
                    .map(item => `${item.source} = ${dataset.data[rowIndex][item.column]}`)
                    .join(", ");
                const recordLabel = `Record ${rowIndex + 1}${idText ? ` (${idText})` : ""}`;

                row.append(element("p", `${recordLabel} · Original: ${original}`));

                const status = element("p", "", "small correction-status");
                status.setAttribute("role", "status");
                const accepted = !unresolved.has(rowIndex);

                function showCorrectionState(pendingEdit = false) {
                    const resolved = accepted && !pendingEdit;
                    row.classList.toggle("correction-resolved", resolved);
                    row.classList.toggle("correction-pending", !resolved);
                    status.textContent = pendingEdit ? "Changes not applied"
                        : !resolved ? "Needs review"
                        : corrected === "" ? "Marked missing" : "Correction applied";
                }

                const input = textField(row, "Replacement", corrected ?? original, value => {
                    showCorrectionState(value !== (corrected ?? original));
                });
                row.append(status);
                showCorrectionState();
                const apply = element("button", "Apply correction");
                const missing = element("button", "Mark missing");
                const undo = element("button", "Undo");
                apply.type = missing.type = undo.type = "button";
                undo.disabled = corrected === undefined;

                apply.addEventListener("click", () => {
                    const value = input.value.trim();

                    if (!value || !Number.isFinite(parseStudyNumber(value, variable.numberFormat))) {
                        $("step-error").textContent =
                            "Enter a valid number in the selected format, or click Mark missing.";
                        return;
                    }

                    variable.corrections.set(rowIndex, value);
                    $("step-error").textContent = "";
                    renderCorrections();
                });

                missing.addEventListener("click", () => {
                    variable.corrections.set(rowIndex, "");
                    $("step-error").textContent = "";
                    renderCorrections();
                });

                undo.addEventListener("click", () => {
                    variable.corrections.delete(rowIndex);
                    $("step-error").textContent = "";
                    renderCorrections();
                });

                const actions = element("div", "", "actions");
                actions.append(apply, missing, undo);
                row.append(actions);
                corrections.append(row);
            }
        }
    }
}

function renderCategorical(parent) {
    const categorical = included().filter(variable => variable.type === "categorical");

    if (!categorical.length) {
        parent.append(element("p", "There are no categorical variables to review."));
        return;
    }

    for (const variable of categorical) {
        const box = card(parent, variable.source);

        const separatorPicker = selectField(box, "How to read each cell", [
            ["single", "Keep each cell as one category"],
            ["semicolon", "Split on semicolon (;)"],
            ["pipe", "Split on vertical bar (|)"],
            ["comma", "Split on comma (,)"],
        ], variable.separator, value => {
            variable.separator = value;
            renderLevels();
        });

        checkbox(box, "Show the variable heading", variable.showHeader, checked => {
            variable.showHeader = checked;
        });

        const levels = element("div");
        box.append(levels);
        renderLevels();

        function renderLevels() {
            const wasEditing = levels.querySelector("details")?.open;
            levels.replaceChildren();
            separatorPicker.disabled = variable.categoryMerges.length > 0;

            const {counts, denominator} = categoricalReview(dataset, variable);

            if (!denominator) {
                levels.append(element("p", "This column has no nonmissing responses."));
                return;
            }

            levels.append(element("p", `Records with responses: ${denominator}`));

            for (const [value, count] of counts) {
                checkbox(levels, `${value} (${count})`, !variable.hiddenLevels.has(value), checked => {
                    if (checked) variable.hiddenLevels.delete(value);
                    else variable.hiddenLevels.add(value);
                });
            }

            renderValueMerge(levels, variable, renderLevels);
            if (wasEditing) levels.querySelector("details").open = true;
        }
    }
}

function renderStratification(parent) {
    const available = variables.filter(canUseVariable);
    if (strataSettings.column && !available.some(item => String(item.column) === strataSettings.column)) {
        strataSettings = {column: ""};
    }
    const choice = selectField(parent, "Stratify by", [
        ["", "No stratification"],
        ...available.map(variable => [String(variable.column), variable.source]),
    ], available.some(item => String(item.column) === strataSettings.column)
        ? strataSettings.column : "", value => {
        const variable = available.find(item => String(item.column) === value);
        strataSettings = {
            column: value,
            cohort: strataSettings.cohort ?? "all",
            mode: variable?.type ?? "categorical",
            numberFormat: variable?.numberFormat ?? "decimal-point",
            separator: variable?.separator ?? "single",
            cuts: "",
            excluded: new Set(["missing"]),
            labels: new Map(),
        };
        renderControls();
    });
    const controls = element("div");
    parent.append(controls);

    function renderControls() {
        controls.replaceChildren();
        if (!choice.value) return;
        const variable = variables[Number(choice.value)];
        selectField(controls, "Records in the cohort column", [
            ["all", "All records in the dataset"],
            ["shown", "Only records in the displayed strata"],
        ], strataSettings.cohort ?? "all", value => {
            strataSettings.cohort = value;
        });

        selectField(controls, "Grouping method", [
            ["categorical", "Categories"],
            ["numeric", "Numerical bins"],
        ], strataSettings.mode, value => {
            strataSettings.mode = value;
            renderControls();
        });

        if (strataSettings.mode === "numeric") {
            selectField(controls, "Number format", [
                ["decimal-point", "Decimal point"],
                ["decimal-comma", "Decimal comma"],
                ["comma-thousands", "Comma thousands separator"],
            ], strataSettings.numberFormat, value => {
                strataSettings.numberFormat = value;
                renderGroups();
            });
            const cuts = textField(controls,
                "Cut points separated by semicolons (e.g. 40; 65)", strataSettings.cuts,
                value => { strataSettings.cuts = value; });
            cuts.addEventListener("change", renderGroups);
            controls.append(element("p", "A value at a cut point goes into the higher bin."));
        } else {
            selectField(controls, "Read each cell as", [
                ["single", "One category per cell"],
                ["semicolon", "Separate responses on semicolon"],
                ["comma", "Separate responses on comma"],
                ["pipe", "Separate responses on vertical bar"],
            ], strataSettings.separator, value => {
                strataSettings.separator = value;
                renderGroups();
            });
            controls.append(element("p", "If you split responses, one record can appear in several stratum columns."));
        }

        const groups = element("div");
        controls.append(groups);
        renderGroups();

        function renderGroups() {
            groups.replaceChildren();
            let candidates;
            try {
                candidates = buildStrata(dataset, variable, strataSettings);
            } catch (error) {
                groups.append(element("p", error.message));
                return;
            }
            for (const group of candidates) {
                const box = card(groups, `${group.label} · ${group.indices.length} records`);
                checkbox(box, "Show this column", !strataSettings.excluded.has(group.key), show => {
                    if (show) strataSettings.excluded.delete(group.key);
                    else strataSettings.excluded.add(group.key);
                });
                textField(box, "Column heading", strataSettings.labels.get(group.key) ?? group.label,
                    label => { strataSettings.labels.set(group.key, label); });
            }
        }
    }

    renderControls();
}

function selectedStrata() {
    if (!strataSettings.column) return [];
    const variable = variables[Number(strataSettings.column)];
    if (!variable || !canUseVariable(variable)) {
        throw new Error("Choose a non-identifier column to stratify by.");
    }
    const groups = buildStrata(dataset, variable, strataSettings).filter(group =>
        !strataSettings.excluded.has(group.key)
    );
    if (!groups.length) throw new Error("Choose at least one stratum, or select No stratification.");
    if (groups.length > 16382) throw new Error("Select fewer strata to fit within Excel’s column limit.");
    return groups.map(group => ({
        ...group,
        label: (strataSettings.labels.get(group.key) ?? group.label).trim(),
    }));
}

function defaultCohortHeading() {
    if (!strataSettings.column) return "Overall";
    return strataSettings.cohort === "shown" ? "Selected strata" : "All cohort";
}

function tableHeadings() {
    return [columnLabels.characteristic.trim(), columnLabels.overall.trim() ||
        defaultCohortHeading()];
}

function validateHeadings(strata) {
    const labels = [...tableHeadings(), ...strata.map(group => group.label)];
    if (labels.some(label => !label)) throw new Error("Enter a name for each displayed column.");
    if (new Set(labels.map(label => label.toLowerCase())).size !== labels.length) {
        throw new Error("Use a different heading for each displayed column.");
    }
}

function renderValueMerge(parent, variable, refresh) {
    const editor = element("details");
    editor.append(element("summary", "Merge values in this column"));
    parent.append(editor);

    for (const rule of variable.categoryMerges) {
        const line = element("p", `${rule.values.join(", ")} → ${rule.label} `);
        const undo = element("button", "Undo merge");
        undo.type = "button";
        undo.addEventListener("click", () => {
            variable.categoryMerges = variable.categoryMerges.filter(item => item !== rule);
            refresh();
        });
        line.append(undo);
        editor.append(line);
    }

    const raw = categoricalReview(dataset, {...variable, categoryMerges: []}).counts;
    const used = new Set(variable.categoryMerges.flatMap(rule => rule.values));
    const eligible = [...raw.keys()].filter(value => !used.has(value));

    if (eligible.length < 2) return;
    editor.append(element("p", "Choose at least two values. Each record counts once in the new category."));
    const checks = eligible.map(value =>
        [value, checkbox(editor, value, false, () => {})]);
    const name = textField(editor, "Merged category name", "", () => {});
    const apply = element("button", "Merge selected values");
    apply.type = "button";
    const error = element("p");
    error.setAttribute("role", "alert");
    editor.append(apply, error);

    apply.addEventListener("click", () => {
        const values = checks.filter(([, check]) => check.checked).map(([value]) => value);
        const label = name.value.trim();
        if (values.length < 2 || !label) {
            error.textContent = "Select at least two values and enter a name.";
            return;
        }
        const existing = [...raw.keys(), ...variable.categoryMerges.map(rule => rule.label)];
        if (existing.includes(label)) {
            error.textContent = "Use a name that is different from the existing categories.";
            return;
        }
        variable.categoryMerges.push({values, label});
        variable.hiddenLevels.delete(label);
        refresh();
    });
}

function renderCategoryEditor(parent, variable, value) {
    const editor = element("details", "", "category-editor");
    const summary = element("summary");
    const heading = variable.headingsBefore.get(value) ?? {label: "", kind: "heading"};
    variable.headingsBefore.set(value, heading);

    function updateSummary() {
        const label = variable.levelLabels.get(value) ?? value;
        const headingText = heading.label.trim() ? ` · Heading: ${heading.label.trim()}` : "";
        summary.textContent = `Category ${value}: ${label}${headingText} · Edit`;
    }

    updateSummary();
    editor.append(summary);
    textField(editor, "Category display name", variable.levelLabels.get(value) ?? value, label => {
        variable.levelLabels.set(value, label);
        updateSummary();
    });
    textField(editor, "Optional heading before this category", heading.label, label => {
        heading.label = label;
        updateSummary();
    });
    selectField(editor, "Heading placement", [
        ["heading", "Heading"],
        ["subheading", "Indented subheading"],
    ], heading.kind, kind => { heading.kind = kind; });
    parent.append(editor);
}

function renderNames(parent) {
    textField(parent, "Table title", title, value => {
        title = value;
    });

    const headers = card(parent, "Table column names");
    textField(headers, "Characteristic column", columnLabels.characteristic,
        value => { columnLabels.characteristic = value; });
    const defaultOverall = defaultCohortHeading();
    textField(headers, "Cohort column", columnLabels.overall,
        value => { columnLabels.overall = value; }).placeholder = defaultOverall;

    const options = card(parent, "Table display and export");

    for (const [key, label] of [
        ["showTotal", "Show the number of records"],
        ["showMissing", "Show missing counts"],
        ["boldHeadings", "Bold the table title and headings"],
        ["includeDetails", "Include a complete statistics sheet in Excel"],
    ]) {
        checkbox(options, label, display[key], checked => {
            display[key] = checked;
        });
    }

    const noteLabel = element("label", "Footer text", "wizard-field");
    const note = element("textarea");
    note.rows = 5;
    note.value = display.note;
    note.addEventListener("input", () => { display.note = note.value; });
    noteLabel.append(note);
    checkbox(options, "Include the footer in the table and exports", display.showNotes, checked => {
        display.showNotes = checked;
    });
    options.append(noteLabel);

    for (const variable of included()) {
        const box = card(parent, variable.source);

        textField(box, "Variable display name", variable.label, value => {
            variable.label = value;
        });

        if (variable.type !== "categorical") continue;

        const {counts} = categoricalReview(dataset, variable);

        for (const value of counts.keys()) {
            if (variable.hiddenLevels.has(value)) continue;

            renderCategoryEditor(box, variable, value);
        }
    }
}

const RENDERERS = [
    renderCohortFilter,
    renderIdentifiers,
    renderIncluded,
    renderNumerical,
    renderCategorical,
    renderStratification,
    renderNames,
];

function renderStep() {
    current = undefined;
    $("results").hidden = true;
    $("settings").hidden = false;
    $("step-counter").textContent = `Step ${step + 1} of ${STEPS.length}`;
    $("step-title").textContent = STEPS[step][0];
    $("step-help").textContent = STEPS[step][1];
    $("step-error").textContent = "";
    $("step-content").replaceChildren();
    $("back").hidden = step === 0;
    $("next").textContent = step === STEPS.length - 1 ? "Generate preview" : "Continue";
    RENDERERS[step]($("step-content"));
    $("step-title").focus();
}

function validateStep() {
    if (step === 0) applyCohortFilter();
    if (step === 2 && !included().length) {
        throw new Error("Choose at least one variable to include.");
    }

    if (step === 3) {
        for (const variable of included().filter(item => item.type === "numeric")) {
            if (numericReview(dataset, variable).invalid.length) {
                throw new Error(`Resolve the highlighted values in ${variable.source}, or go back to change its type.`);
            }
        }
    }

    if (step === 4) {
        for (const variable of included().filter(item => item.type === "categorical")) {
            const {counts, missing} = categoricalReview(dataset, variable);
            const hasVisibleLevel = [...counts.keys()].some(value => !variable.hiddenLevels.has(value));

            if (!hasVisibleLevel && !missing) {
                throw new Error(`Select at least one category for ${variable.source}, or go back to exclude the variable.`);
            }
        }
    }

    if (step === 5) selectedStrata();

    if (step === 6) {
        validateHeadings(selectedStrata());
        if (!title.trim()) throw new Error("Enter a table title.");
        if (display.showNotes && !display.note.trim()) throw new Error("Enter footer text or turn the footer off.");

        for (const variable of included()) {
            if (!variable.label.trim()) {
                throw new Error(`Enter a display name for ${variable.source}.`);
            }

            if (variable.type !== "categorical") continue;

            const {counts} = categoricalReview(dataset, variable);
            const labels = new Set();

            for (const value of counts.keys()) {
                if (variable.hiddenLevels.has(value)) continue;

                const label = (variable.levelLabels.get(value) ?? value).trim();

                if (!label) throw new Error(`Enter a label for category ${value}.`);
                if (labels.has(label)) {
                    throw new Error(`Use distinct labels for the categories in ${variable.source}. Renaming does not combine counts.`);
                }

                labels.add(label);
            }
        }
    }
}

function preview() {
    const exportVariables = included().map(variable => ({
        ...variable,
        label: variable.label.trim(),
        levelLabels: new Map(
            [...variable.levelLabels].map(([value, label]) => [value, label.trim()])
        ),
    }));

    const strata = selectedStrata();
    const result = buildStratifiedTable(
        dataset, exportVariables, {...display, note: display.note.trim()},
        tableHeadings(), strata, strataSettings.cohort
    );

    $("results").classList.toggle(
        "bold-table-headings",
        result.display.boldHeadings
    );
    const defaultRows = result.rows;
    result.rows = orderTableRows(defaultRows, rowOrder);
    current = {title: title.trim(), result};
    renderRowOrder(result, defaultRows);
    $("caption").textContent = current.title;

    const header = $("results").querySelector("thead tr");
    header.replaceChildren(...result.headings.map(label => {
        const cell = element("th", label);
        cell.scope = "col";
        return cell;
    }));

    renderPreviewRows(result.rows);

    $("notes").textContent = result.display.showNotes ? result.display.note : "";
    $("notes").hidden = !result.display.showNotes;
    $("settings").hidden = true;
    $("results").hidden = false;
    $("results").scrollIntoView({block: "start"});
}

function renderPreviewRows(rows) {
    $("rows").replaceChildren(...rows.map(([label, values, kind]) => {
        const row = element("tr");
        const cell = element(["heading", "subheading"].includes(kind) ? "th" : "td", label, kind);
        row.append(cell, ...values.map(value => element("td", value)));
        return row;
    }));
}

function renderRowOrder(result, defaultRows) {
    let panel = $("row-order");
    if (!panel) {
        panel = element("details", "", "variable-card");
        panel.id = "row-order";
        $("results").querySelector(".scroll").before(panel);
    }
    panel.open = false;
    panel.replaceChildren(element("summary", "Reorder rows"));
    panel.append(element("p", "Move a row up or down. Headings and their rows move independently. Indentation stays as set."));
    const status = element("p", "", "small");
    status.setAttribute("role", "status");
    const list = element("ol");
    const resetOrder = element("button", "Restore default order");
    resetOrder.type = "button";
    panel.append(resetOrder, status, list);

    function apply(rows) {
        result.rows = rows;
        rowOrder = rows.map(row => row[3]);
        renderPreviewRows(rows);
        renderList();
    }

    function move(from, to) {
        if (to < 0 || to >= result.rows.length) return;
        const rows = [...result.rows];
        const [row] = rows.splice(from, 1);
        rows.splice(to, 0, row);
        apply(rows);
        status.textContent = `${row[0]} moved to position ${to + 1} of ${rows.length}.`;
        const buttons = [...list.children[to].querySelectorAll("button")];
        const preferred = buttons[to < from ? 0 : 1];
        (preferred.disabled ? buttons.find(button => !button.disabled) ?? resetOrder : preferred).focus();
    }

    function renderList() {
        list.replaceChildren(...result.rows.map(([label, , , id], i) => {
            const [type, column] = JSON.parse(id);
            const variable = variables[column];
            const context = variable && ["category", "statistic", "missing", "heading"].includes(type)
                ? ` (${variable.label})` : "";
            const item = element("li");
            const controls = element("div", "", "variable");
            controls.append(element("span", `${label}${context}`));
            const buttons = element("div", "", "actions");
            for (const [direction, offset] of [["Move up", -1], ["Move down", 1]]) {
                const button = element("button", direction);
                button.type = "button";
                button.disabled = i + offset < 0 || i + offset >= result.rows.length;
                button.setAttribute("aria-label", `${direction}: ${label}${context}`);
                button.addEventListener("click", () => move(i, i + offset));
                buttons.append(button);
            }
            controls.append(buttons);
            item.append(controls);
            return item;
        }));
    }

    resetOrder.addEventListener("click", () => {
        apply(defaultRows);
        rowOrder = [];
        status.textContent = "Default row order restored.";
    });
    renderList();
}

function chooseWorksheet(workbook, filename, version) {
    const picker = $("worksheet-picker");
    const select = selectField(picker, "Choose a worksheet", workbook.names.map(
        (name, index) => [String(index), name]
    ), "0", () => {});

    const button = element("button", "Use worksheet", "primary");
    button.type = "button";
    picker.append(button);

    button.addEventListener("click", async () => {
        const index = Number(select.value);
        select.disabled = button.disabled = true;
        $("error").textContent = "";

        try {
            const input = await workbook.readSheet(index);
            if (version !== loadVersion) return;
            picker.replaceChildren();
            loadDataset(input, `${filename} · ${workbook.names[index]}`);
        } catch (error) {
            if (version === loadVersion) $("error").textContent = error.message;
        } finally {
            select.disabled = button.disabled = false;
        }
    });

    $("clear").hidden = false;
    $("status").textContent = filename;
    select.focus();
}

$("file").addEventListener("change", async event => {
    const file = event.target.files[0];
    if (!file) return;

    reset();
    const version = loadVersion;

    try {
        if (!/\.(?:csv|xlsx)$/i.test(file.name)) {
            throw new Error("Choose a CSV or .xlsx file.");
        }

        const bytes = await file.arrayBuffer();
        if (version !== loadVersion) return;

        if (/\.xlsx$/i.test(file.name)) {
            const workbook = await openExcelFile(bytes);
            if (version !== loadVersion) return;

            if (workbook.names.length > 1) {
                chooseWorksheet(workbook, file.name, version);
            } else {
                const input = await workbook.readSheet(0);
                if (version === loadVersion) loadDataset(input, file.name);
            }
            return;
        }

        let text;
        try {
            text = new TextDecoder("utf-8", {fatal: true}).decode(bytes);
        } catch {
            text = new TextDecoder("windows-1252").decode(bytes);
        }

        loadData(text, file.name);
    } catch (error) {
        if (version === loadVersion) $("error").textContent = error.message;
    }
});

$("clear").addEventListener("click", reset);

$("back").addEventListener("click", () => {
    if (step > 0) {
        step--;
        renderStep();
    }
});

$("next").addEventListener("click", () => {
    try {
        validateStep();

        if (step === STEPS.length - 1) {
            preview();
        } else {
            step++;
            renderStep();
        }
    } catch (error) {
        $("step-error").textContent = error.message;
        $("step-error").scrollIntoView({block: "nearest"});
    }
});

$("edit").addEventListener("click", renderStep);

function attachDownload(id, extension, mime, generate) {
    $(id).addEventListener("click", async event => {
        const button = event.currentTarget;
        const snapshot = current;
        if (!snapshot) return;

        button.disabled = true;
        $("error").textContent = "";

        try {
            const bytes = await generate(snapshot.title, snapshot.result);
            if (current !== snapshot) return;

            const url = URL.createObjectURL(new Blob([bytes], {type: mime}));
            const link = element("a");
            const name = snapshot.title.replace(/[^\p{L}\p{N}._ -]+/gu, "_").trim();

            link.href = url;
            link.download = `${name || "table1"}.${extension}`;
            document.body.append(link);
            link.click();
            link.remove();
            setTimeout(() => URL.revokeObjectURL(url), 30000);
        } catch (error) {
            if (current === snapshot) {
                $("error").textContent = `Could not generate ${extension.toUpperCase()}: ${error.message}`;
                $("error").scrollIntoView({block: "nearest"});
            }
        } finally {
            button.disabled = false;
        }
    });
}

attachDownload(
    "excel",
    "xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    excelBytes
);

attachDownload("pdf", "pdf", "application/pdf", pdfBytes);
