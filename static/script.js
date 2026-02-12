(() => {
    // ===== Mode Tabs =====
    const modeTabs = document.querySelectorAll(".mode-tab");
    const modeContents = document.querySelectorAll(".mode-content");

    modeTabs.forEach(tab => {
        tab.addEventListener("click", () => {
            modeTabs.forEach(t => t.classList.remove("active"));
            modeContents.forEach(c => c.classList.remove("active"));
            tab.classList.add("active");
            document.getElementById("mode-" + tab.dataset.mode).classList.add("active");
        });
    });

    // ===== Merge Mode Elements =====
    const stepUpload = document.getElementById("step-upload");
    const stepConfig = document.getElementById("step-config");
    const stepPreview = document.getElementById("step-preview");

    const uploadForm = document.getElementById("upload-form");
    const file1Input = document.getElementById("file1");
    const file2Input = document.getElementById("file2");
    const name1 = document.getElementById("name1");
    const name2 = document.getElementById("name2");
    const conversionNotice = document.getElementById("conversion-notice");

    const sheet1Select = document.getElementById("sheet1");
    const sheet2Select = document.getElementById("sheet2");
    const sheetSelectGroup = document.getElementById("sheet-select-group");
    const joinOptions = document.getElementById("join-options");
    const joinColumn = document.getElementById("join-column");

    const btnMerge = document.getElementById("btn-merge");
    const btnBackUpload = document.getElementById("btn-back-upload");
    const btnBackConfig = document.getElementById("btn-back-config");
    const btnStartOver = document.getElementById("btn-start-over");

    const previewHead = document.getElementById("preview-head");
    const previewBody = document.getElementById("preview-body");
    const mergeInfo = document.getElementById("merge-info");
    const spinner = document.getElementById("spinner");

    // ===== Convert Mode Elements =====
    const convertForm = document.getElementById("convert-form");
    const convertFileInput = document.getElementById("convert-file");
    const convertName = document.getElementById("convert-name");
    const convertUploadSection = document.getElementById("convert-upload");
    const convertPreviewSection = document.getElementById("convert-preview");
    const convertHead = document.getElementById("convert-head");
    const convertBody = document.getElementById("convert-body");
    const convertInfo = document.getElementById("convert-info");
    const btnConvertAgain = document.getElementById("btn-convert-again");
    const convertSheetSelector = document.getElementById("convert-sheet-selector");
    const convertSheetSelect = document.getElementById("convert-sheet-select");

    let fileData = null; // stores sheet info returned by /upload

    // ===== Helpers =====
    function showStep(step) {
        [stepUpload, stepConfig, stepPreview].forEach(s => s.classList.remove("active"));
        step.classList.add("active");
    }

    function showConvertStep(step) {
        [convertUploadSection, convertPreviewSection].forEach(s => s.classList.remove("active"));
        step.classList.add("active");
    }

    function showSpinner() { spinner.style.display = "flex"; }
    function hideSpinner() { spinner.style.display = "none"; }

    function getSelectedMergeType() {
        return document.querySelector('input[name="merge_type"]:checked').value;
    }

    function populateSelect(select, items) {
        select.innerHTML = "";
        items.forEach(item => {
            const opt = document.createElement("option");
            opt.value = item;
            opt.textContent = item;
            select.appendChild(opt);
        });
    }

    function renderTable(headEl, bodyEl, data) {
        headEl.innerHTML = "";
        bodyEl.innerHTML = "";

        const tr = document.createElement("tr");
        data.columns.forEach(col => {
            const th = document.createElement("th");
            th.textContent = col;
            tr.appendChild(th);
        });
        headEl.appendChild(tr);

        data.rows.forEach(row => {
            const tr = document.createElement("tr");
            row.forEach(cell => {
                const td = document.createElement("td");
                td.textContent = cell;
                tr.appendChild(td);
            });
            bodyEl.appendChild(tr);
        });
    }

    // ===== Merge Mode: File Name Display =====
    file1Input.addEventListener("change", () => {
        name1.textContent = file1Input.files[0]?.name || "Choose File 1...";
        name1.classList.toggle("selected", !!file1Input.files[0]);
    });
    file2Input.addEventListener("change", () => {
        name2.textContent = file2Input.files[0]?.name || "Choose File 2...";
        name2.classList.toggle("selected", !!file2Input.files[0]);
    });

    // ===== Convert Mode: File Name Display =====
    convertFileInput.addEventListener("change", () => {
        convertName.textContent = convertFileInput.files[0]?.name || "Choose a file...";
        convertName.classList.toggle("selected", !!convertFileInput.files[0]);
    });

    // ===== Merge Mode: Upload =====
    uploadForm.addEventListener("submit", async (e) => {
        e.preventDefault();
        if (!file1Input.files[0] || !file2Input.files[0]) {
            alert("Please select two files.");
            return;
        }

        const fd = new FormData();
        fd.append("file1", file1Input.files[0]);
        fd.append("file2", file2Input.files[0]);

        showSpinner();
        try {
            const res = await fetch("/upload", { method: "POST", body: fd });
            const json = await res.json();
            if (!res.ok) {
                alert(json.error || "Upload failed.");
                return;
            }
            fileData = json;

            // Show conversion notice if files were auto-converted
            if (json.conversions && json.conversions.length > 0) {
                conversionNotice.innerHTML = "<strong>Auto-converted:</strong> " +
                    json.conversions.join(", ");
                conversionNotice.style.display = "block";
            } else {
                conversionNotice.style.display = "none";
            }

            setupConfigStep();
            showStep(stepConfig);
        } catch (err) {
            alert("Upload failed: " + err.message);
        } finally {
            hideSpinner();
        }
    });

    // Populate config step based on uploaded file info
    function setupConfigStep() {
        const sheets1 = Object.keys(fileData.file1_sheets);
        const sheets2 = Object.keys(fileData.file2_sheets);
        populateSelect(sheet1Select, sheets1);
        populateSelect(sheet2Select, sheets2);
        updateJoinColumns();
        handleMergeTypeChange();
    }

    // Update join column options when sheet selection changes
    function updateJoinColumns() {
        if (!fileData) return;
        const cols1 = fileData.file1_sheets[sheet1Select.value] || [];
        const cols2 = fileData.file2_sheets[sheet2Select.value] || [];
        const common = cols1.filter(c => cols2.includes(c));
        populateSelect(joinColumn, common.length ? common : ["(no common columns)"]);
    }

    sheet1Select.addEventListener("change", updateJoinColumns);
    sheet2Select.addEventListener("change", updateJoinColumns);

    // Show/hide options based on merge type
    function handleMergeTypeChange() {
        const type = getSelectedMergeType();
        joinOptions.style.display = type === "join" ? "flex" : "none";
        sheetSelectGroup.style.display = type === "sheets" ? "none" : "block";
    }

    document.querySelectorAll('input[name="merge_type"]').forEach(r =>
        r.addEventListener("change", handleMergeTypeChange)
    );

    // ===== Merge Mode: Merge =====
    btnMerge.addEventListener("click", async () => {
        const payload = {
            merge_type: getSelectedMergeType(),
            sheet1: sheet1Select.value,
            sheet2: sheet2Select.value,
            join_column: joinColumn.value,
            join_how: document.getElementById("join-how").value,
        };

        showSpinner();
        try {
            const res = await fetch("/merge", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(payload),
            });
            const json = await res.json();
            if (!res.ok) {
                alert(json.error || "Merge failed.");
                return;
            }
            renderPreview(json);
            showStep(stepPreview);
        } catch (err) {
            alert("Merge failed: " + err.message);
        } finally {
            hideSpinner();
        }
    });

    // Render merge preview table
    function renderPreview(data) {
        renderTable(previewHead, previewBody, data);

        let info = `Showing ${data.rows.length} of ${data.total_rows} rows.`;
        if (data.sheet_names) {
            info += ` Sheets: ${data.sheet_names.join(", ")}`;
        }
        mergeInfo.textContent = info;
    }

    // ===== Merge Mode: Navigation =====
    btnBackUpload.addEventListener("click", () => showStep(stepUpload));
    btnBackConfig.addEventListener("click", () => showStep(stepConfig));
    btnStartOver.addEventListener("click", () => {
        fileData = null;
        uploadForm.reset();
        name1.textContent = "Choose File 1...";
        name1.classList.remove("selected");
        name2.textContent = "Choose File 2...";
        name2.classList.remove("selected");
        conversionNotice.style.display = "none";
        showStep(stepUpload);
    });

    // ===== Convert Mode: Submit =====
    convertForm.addEventListener("submit", async (e) => {
        e.preventDefault();
        if (!convertFileInput.files[0]) {
            alert("Please select a file to convert.");
            return;
        }

        const fd = new FormData();
        fd.append("file", convertFileInput.files[0]);

        showSpinner();
        try {
            const res = await fetch("/convert", { method: "POST", body: fd });
            const json = await res.json();
            if (!res.ok) {
                alert(json.error || "Conversion failed.");
                return;
            }

            renderTable(convertHead, convertBody, json);

            // Show sheet selector if multiple tables were extracted
            if (json.sheet_names && json.sheet_names.length > 1) {
                populateSelect(convertSheetSelect, json.sheet_names);
                convertSheetSelector.style.display = "block";
                convertInfo.textContent =
                    `Converted "${json.original_name}" - ${json.sheet_names.length} tables found. Showing ${json.rows.length} of ${json.total_rows} rows.`;
            } else {
                convertSheetSelector.style.display = "none";
                convertInfo.textContent =
                    `Converted "${json.original_name}" - showing ${json.rows.length} of ${json.total_rows} rows.`;
            }
            showConvertStep(convertPreviewSection);
        } catch (err) {
            alert("Conversion failed: " + err.message);
        } finally {
            hideSpinner();
        }
    });

    // ===== Convert Mode: Sheet Switching =====
    convertSheetSelect.addEventListener("change", async () => {
        showSpinner();
        try {
            const res = await fetch("/convert-sheet?sheet=" + encodeURIComponent(convertSheetSelect.value));
            const json = await res.json();
            if (!res.ok) {
                alert(json.error || "Failed to load sheet.");
                return;
            }
            renderTable(convertHead, convertBody, json);
            convertInfo.textContent =
                `Viewing "${convertSheetSelect.value}" - showing ${json.rows.length} of ${json.total_rows} rows.`;
        } catch (err) {
            alert("Failed to load sheet: " + err.message);
        } finally {
            hideSpinner();
        }
    });

    // ===== Convert Mode: Navigation =====
    btnConvertAgain.addEventListener("click", () => {
        convertForm.reset();
        convertName.textContent = "Choose a file...";
        convertName.classList.remove("selected");
        convertSheetSelector.style.display = "none";
        showConvertStep(convertUploadSection);
    });
})();
