document.addEventListener('DOMContentLoaded', () => {
    const dropArea = document.getElementById('drop-area');
    const fileInput = document.getElementById('fileInput');
    const fileNameDisplay = document.getElementById('file-name');
    const processBtn = document.getElementById('processBtn');
    const statusMessage = document.getElementById('status-message');
    const markupPercentInput = document.getElementById('markupPercent');
    const flatMarkup20Input = document.getElementById('flatMarkup20');
    const flatMarkup40Input = document.getElementById('flatMarkup40');
    const flatMarkup40HCInput = document.getElementById('flatMarkup40HC');
    const flatMarkup45Input = document.getElementById('flatMarkup45');

    let selectedFile = null;

    // --- UI Interactions ---
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    ['dragenter', 'dragover'].forEach(eventName => {
        dropArea.addEventListener(eventName, () => dropArea.classList.add('highlight'), false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, () => dropArea.classList.remove('highlight'), false);
    });

    dropArea.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        if (dt.files.length) handleFile(dt.files[0]);
    });

    fileInput.addEventListener('change', function () {
        if (this.files.length) handleFile(this.files[0]);
    });

    function handleFile(file) {
        if (!file.name.match(/\.(xlsx|xls)$/im)) {
            showStatus('Please upload a valid Excel file (.xlsx or .xls)', 'error');
            return;
        }
        selectedFile = file;
        fileNameDisplay.textContent = file.name;
        processBtn.disabled = false;
        hideStatus();
    }

    function showStatus(text, type) {
        statusMessage.textContent = text;
        statusMessage.className = `status-message show ${type}`;
    }

    function hideStatus() {
        statusMessage.className = 'status-message hidden';
    }

    // --- Core Processing Logic ---
    processBtn.addEventListener('click', async () => {
        if (!selectedFile) return;

        const markupPercent = parseFloat(markupPercentInput.value) || 0;
        const flat20 = parseFloat(flatMarkup20Input.value) || 0;
        const flat40 = parseFloat(flatMarkup40Input.value) || 0;
        const flat40hc = parseFloat(flatMarkup40HCInput.value) || 0;
        const flat45 = parseFloat(flatMarkup45Input.value) || 0;

        processBtn.classList.add('loading');
        processBtn.disabled = true;
        showStatus('Analyzing and processing data...', 'info');

        try {
            const arrayBuffer = await readFileAsArrayBuffer(selectedFile);
            await new Promise(resolve => setTimeout(resolve, 50));

            const workbook = XLSX.read(arrayBuffer, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            let rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });

            if (rows.length < 2) throw new Error("The Excel sheet is empty or has no data rows.");

            // Helpers
            const getCol = (name) => rows[0].findIndex(h => typeof h === 'string' && h.trim().toUpperCase() === name.toUpperCase());
            const isBlank = (val) => val === null || val === undefined || String(val).trim() === '';

            // Ensure uniform row length
            const initialLen = rows[0].length;
            rows.forEach(r => { while(r.length < initialLen) r.push(null); });

            // Step 1: move INLAND ORIGIN to column B (index 1)
            let colN = getCol("INLAND ORIGIN");
            if (colN > -1) {
                for (let r = 0; r < rows.length; r++) {
                    let val = rows[r].splice(colN, 1)[0];
                    rows[r].splice(1, 0, val);
                }
            }

            // Step 2: move INLAND DEST to column E (index 4)
            let colQ = getCol("INLAND DEST");
            if (colQ > -1) {
                for (let r = 0; r < rows.length; r++) {
                    let val = rows[r].splice(colQ, 1)[0];
                    rows[r].splice(4, 0, val);
                }
            }

            // Get new indices for steps 3-6 (Dynamic so it accounts for array shifts properly)
            let colB = 1;
            let colC = 2;
            let colD = 3;
            let colE = 4;
            let colOrigRout = getCol("ORIGIN ROUTING");
            let colDestRout = getCol("DESTINATION ROUTING");

            for (let r = 1; r < rows.length; r++) {
                let row = rows[r];

                // Step 3
                if (isBlank(row[colB])) {
                    row[colB] = row[colC];
                    row[colC] = null;
                }

                // Step 4
                if (isBlank(row[colC]) && !isBlank(row[colB]) && String(row[colB]).toLowerCase().includes('ramp')) {
                    if (colOrigRout > -1) {
                        row[colC] = row[colOrigRout];
                        row[colOrigRout] = null;
                    }
                }

                // Step 5
                if (isBlank(row[colE])) {
                    row[colE] = row[colD];
                    row[colD] = null;
                }

                // Step 6
                if (isBlank(row[colD])) {
                    if (colDestRout > -1) {
                        row[colD] = row[colDestRout];
                        row[colDestRout] = null;
                    }
                }
            }

            // Step 7: Insert new 5 columns between column AC and AD (inserting at index 29)
            for (let r = 0; r < rows.length; r++) {
                if (r === 0) {
                    rows[r].splice(29, 0, "BOL", "Cntr. 20 ft", "Cntr. 40 ft", "Cntr. 40 ft HC", "Cntr. 45 ft HC");
                } else {
                    rows[r].splice(29, 0, null, null, null, null, null);
                }
            }

            // Step 8 & 9 & 10
            for (let r = 1; r < rows.length; r++) {
                let row = rows[r];

                // Step 10: Column F (index 5) gross total rate
                row[5] = extractNumber(row[5]);

                // Step 8: Clean price from Column AI (index 34) forward
                let bolSum = null;
                for (let c = 34; c < row.length; c++) {
                    let textVal = String(row[c] || '');
                    let cleanedStr = textVal.replace(/\s+/g, ' '); // Normalize spaces
                    if (cleanedStr.includes("Per BOL") && !cleanedStr.includes("Per BOL (Excluded)")) {
                        let num = extractNumber(textVal);
                        row[c] = num;
                        if (num !== null) bolSum = (bolSum || 0) + num;
                    } else {
                        row[c] = null;
                    }
                }

                // Step 9: Column AD (index 29) sum
                row[29] = bolSum;
            }

            // Step 11: Fill columns AE to AH from column F based on specific columns and group into 1 line.
            // Grouping by B(1), C(2), D(3), E(4), H(7), J(9), R(17), U(20)
            let groups = new Map();
            for (let r = 1; r < rows.length; r++) {
                let row = [...rows[r]];
                let keyArr = [
                    String(row[1] || '').trim(),  // B: INLAND ORIGIN
                    String(row[2] || '').trim(),  // C: BASE ORIGIN
                    String(row[3] || '').trim(),  // D: BASE DESTINATION
                    String(row[4] || '').trim(),  // E: INLAND DESTINATION
                    String(row[7] || '').trim(),  // H: EXPIRATION DATE
                    String(row[9] || '').trim(),  // J: CARRIER SCAC
                    String(row[17] || '').trim(), // R: COMMODITY GROUP
                    String(row[20] || '').trim()  // U: RATE TYPE
                ];
                let key = keyArr.join('|');

                if (!groups.has(key)) {
                    groups.set(key, { baseRow: row, cntrs: {} });
                }

                let g = groups.get(key);
                let unitType = String(row[8] || '').toUpperCase();
                let fVal = row[5];

                if (unitType.includes('20')) g.cntrs['20'] = fVal;
                else if (unitType.includes('40') && unitType.includes('HC')) g.cntrs['40HC'] = fVal;
                else if (unitType.includes('40')) g.cntrs['40'] = fVal;
                else if (unitType.includes('45')) g.cntrs['45'] = fVal;
            }

            let collapsedRows = [rows[0]];
            for (let g of groups.values()) {
                let r = g.baseRow;
                r[30] = g.cntrs['20'] || null;   // AE
                r[31] = g.cntrs['40'] || null;   // AF
                r[32] = g.cntrs['40HC'] || null; // AG
                r[33] = g.cntrs['45'] || null;   // AH
                collapsedRows.push(r);
            }
            rows = collapsedRows;

            // --- MainFreight Logic (Steps 12 to 17) ---
            let mfRows = processMainFreight(rows, markupPercent, flat20, flat40, flat40hc, flat45);

            // Export logic
            const newWorkbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWorkbook, XLSX.utils.aoa_to_sheet(rows), "Quote Cart");
            XLSX.utils.book_append_sheet(newWorkbook, XLSX.utils.aoa_to_sheet(mfRows), "MainFreight");

            const dateTimestamp = new Date().toISOString().replace(/[-:T]/g, '').slice(0, 14);
            const outFileName = `Processed_${selectedFile.name.replace(/\.[^/.]+$/, "")}_${dateTimestamp}.xlsx`;

            XLSX.writeFile(newWorkbook, outFileName);

            showStatus('Success! File processed and downloaded', 'success');
        } catch (error) {
            console.error(error);
            showStatus(`Processing Error: ${error.message}`, 'error');
        } finally {
            processBtn.classList.remove('loading');
            processBtn.disabled = false;
        }
    });

    function processMainFreight(qcRows, markupPct, flat20, flat40, flat40hc, flat45) {
        if (!qcRows || qcRows.length === 0) return [];

        // Step 13: Map specific columns sequentially
        // B(1), C(2), D(3), E(4), H(7), J(9), K(10), R(17), U(20), AC(28)
        // AD(29), AE(30), AF(31), AG(32), AH(33)
        // MainFreight starts with 15 base columns.
        const headerMap = [1, 2, 3, 4, 7, 9, 10, 17, 20, 28, 29, 30, 31, 32, 33];
        
        let mfRows = [];
        for (let r = 0; r < qcRows.length; r++) {
            let row = qcRows[r];
            let newRow = [];
            for (let idx of headerMap) {
                newRow.push(row[idx]);
            }
            mfRows.push(newRow);
        }

        // Steps 18 to 25: Column reorganization and formatting
        let transformedMfRows = [];
        for (let r = 0; r < mfRows.length; r++) {
            let oldRow = mfRows[r];
            let newRow = [];

            // Execute Steps 19 to 21 column reshuffling and Step 20 insertions
            newRow[0] = oldRow[0]; // Origin Name (from INLAND ORIGIN)
            newRow[1] = oldRow[0]; // POL (copied from Origin Name)
            newRow[2] = oldRow[3]; // Destination Name (from INLAND DEST at old index 3)
            newRow[3] = oldRow[3]; // POD (copied from Destination Name)
            newRow[4] = oldRow[1]; // Origin Via (from BASE ORIGIN at old index 1)
            newRow[5] = oldRow[2]; // Destination Via (from BASE DESTINATION at old index 2)

            // Shift the rest of the 15 columns (from old indices 4 to 14) over by 2 places
            for (let c = 4; c < oldRow.length; c++) {
                newRow[c + 2] = oldRow[c];
            }

            if (r === 0) {
                // Steps 18 and 20: Explicitly rename the headers
                newRow[0] = "Origin Name";
                newRow[1] = "POL";
                newRow[2] = "Destination Name";
                newRow[3] = "POD";
                newRow[4] = "Origin Via";
                newRow[5] = "Destination Via";
            } else {
                let originNameFull = String(newRow[0] || '');
                let destNameFull = String(newRow[2] || '');
                let originViaFull = String(newRow[4] || '');
                let destViaFull = String(newRow[5] || '');

                // Helper to extract 5-character UNLOC code
                const getUNLOC = (str) => {
                    let m = str.match(/\b[A-Z]{2}[A-Z0-9]{3}\b/);
                    return m ? m[0] : '';
                };

                // Helper to remove UNLOC code and clean up artifacts (Step 22 & 24)
                const stripUNLOC = (str) => {
                    let cleaned = str.replace(/\b[A-Z]{2}[A-Z0-9]{3}\b/g, '');
                    cleaned = cleaned.replace(/\(\s*\)/g, '').replace(/\[\s*\]/g, '');
                    cleaned = cleaned.replace(/^[-\s]+|[-\s]+$/g, '');
                    cleaned = cleaned.replace(/\s{2,}/g, ' ');

                    // Remove [PORT] and [RAMP] case-insensitively
                    cleaned = cleaned.replace(/\[PORT\]/gi, '').replace(/\[RAMP\]/gi, '');

                    // Remove data after the comma
                    let commaIdx = cleaned.indexOf(',');
                    if (commaIdx !== -1) {
                        cleaned = cleaned.substring(0, commaIdx);
                    }

                    return cleaned.trim();
                };

                // Step 23: extract UNLOC codes for POL and POD
                newRow[1] = getUNLOC(originNameFull);
                newRow[3] = getUNLOC(destNameFull);

                // Step 22 & 24: remove UNLOC codes, [PORT], [RAMP], and commas to retain city names etc.
                newRow[0] = stripUNLOC(originNameFull).toUpperCase();
                newRow[2] = stripUNLOC(destNameFull).toUpperCase();
                newRow[4] = stripUNLOC(originViaFull).toUpperCase();
                newRow[5] = stripUNLOC(destViaFull).toUpperCase();
            }
            transformedMfRows.push(newRow);
        }
        mfRows = transformedMfRows;

        // Note: mfRows lengths is now 17 (indices 0 to 16).
        
        // Step 14: Add headers P, Q, R, S (now indices 17 to 20 dynamically)
        mfRows[0].push("20FT + BOL", "40FT + BOL", "40HC + BOL", "45FT + BOL");

        // Step 16: Add headers T, U, V, W (now indices 21 to 24 dynamically)
        mfRows[0].push("20FT-TOTAL", "40FT-TOTAL", "40HC-TOTAL", "45FT-TOTAL");

        // Mathematical Multiplier (1 + Markup%)
        const m = 1 + (markupPct / 100);

        for (let r = 1; r < mfRows.length; r++) {
            let row = mfRows[r];
            
            // Map the parsed base values.
            // AD was index 29 in QC, mapped to mfRows old index 10. Now index 12 in shifted newRow.
            let bol = extractNumber(row[12]) || 0;

            // Container values from AE, AF, AG, AH -> old array index 11, 12, 13, 14. Now 13, 14, 15, 16.
            let c20 = extractNumber(row[13]);
            let c40 = extractNumber(row[14]);
            let c40hc = extractNumber(row[15]);
            let c45 = extractNumber(row[16]);

            // Step 15 calculations: Do not calculate if container is blank
            let p = c20 !== null ? bol + c20 : null;
            let q = c40 !== null ? bol + c40 : null;
            let hr = c40hc !== null ? bol + c40hc : null;
            let s = c45 !== null ? bol + c45 : null;

            row[17] = p;
            row[18] = q;
            row[19] = hr;
            row[20] = s;

            // Step 17 calculations
            // Formula is implemented as Base * (1 + Markup%)
            const calcTotal = (baseWithBol, flat) => {
                if (baseWithBol === null) return null;
                const base = baseWithBol + flat;
                return (base * (markupPct / 100)) + base;
            };

            row[21] = calcTotal(p, flat20);
            row[22] = calcTotal(q, flat40);
            row[23] = calcTotal(hr, flat40hc);
            row[24] = calcTotal(s, flat45);
        }

        return mfRows;
    }

    function readFileAsArrayBuffer(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(new Uint8Array(e.target.result));
            reader.onerror = () => reject(new Error('Failed to read the file.'));
            reader.readAsArrayBuffer(file);
        });
    }

    function extractNumber(val) {
        if (val === null || val === undefined || val === '') return null;
        if (typeof val === 'number') return val;
        const str = String(val).replace(/,/g, '');
        const numMatch = str.match(/-?\d+(?:\.\d+)?/);
        return numMatch ? parseFloat(numMatch[0]) : null;
    }
});
