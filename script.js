document.addEventListener('DOMContentLoaded', () => {
    const dropArea = document.getElementById('drop-area');
    const fileInput = document.getElementById('fileInput');
    const fileNameDisplay = document.getElementById('file-name');
    const processBtn = document.getElementById('processBtn');
    const statusMessage = document.getElementById('status-message');
    const markupPercentInput = document.getElementById('markupPercent');
    const flatMarkupInput = document.getElementById('flatMarkup');

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
        const files = dt.files;
        if (files.length) {
            handleFile(files[0]);
        }
    });

    fileInput.addEventListener('change', function () {
        if (this.files.length) {
            handleFile(this.files[0]);
        }
    });

    function handleFile(file) {
        if (!file.name.match(/\.(xlsx|xls)$/im)) {
            showStatus('Please upload a valid Excel file (.xlsx or .xls)', 'error');
            selectedFile = null;
            processBtn.disabled = true;
            fileNameDisplay.textContent = '';
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

        // Read Markup Values
        const markupPercent = parseFloat(markupPercentInput.value) || 0;
        const flatMarkup = parseFloat(flatMarkupInput.value) || 0;

        // UI Loading State
        processBtn.classList.add('loading');
        processBtn.disabled = true;
        showStatus('Analyzing and processing data...', 'info');

        try {
            // Read file async
            const arrayBuffer = await readFileAsArrayBuffer(selectedFile);

            // Process (Adding small timeout so UI render completes)
            await new Promise(resolve => setTimeout(resolve, 50));

            const workbook = XLSX.read(arrayBuffer, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            // Convert to Array of Arrays (AoA) for precise manipulation
            let rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });

            if (rows.length === 0) {
                throw new Error("The Excel sheet is empty.");
            }

            // Execute the original 6 transformation steps and new Steps 7-11
            processDataRowByRow(rows);

            // Create MainFreight sheet data (Steps 12-20)
            const mainFreightRows = createMainFreightSheet(rows, flatMarkup, markupPercent);

            // Convert back to worksheet
            const newWorksheet = XLSX.utils.aoa_to_sheet(rows);
            const newWorkbook = XLSX.utils.book_new();

            // Step 12 & 13: Rename first sheet to "Quote Cart" and add "MainFreight"
            XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Quote Cart");

            const mainFreightWorksheet = XLSX.utils.aoa_to_sheet(mainFreightRows);
            XLSX.utils.book_append_sheet(newWorkbook, mainFreightWorksheet, "MainFreight");

            // Generate output file name with datestamp
            const dateTimestamp = new Date().toISOString().replace(/[-:T]/g, '').slice(0, 14);
            const outFileName = `Processed_${selectedFile.name.replace(/\.[^/.]+$/, "")}_${dateTimestamp}.xlsx`;

            // Trigger file download
            XLSX.writeFile(newWorkbook, outFileName);

            showStatus('Success! File beautifully processed and downloaded.', 'success');
        } catch (error) {
            console.error(error);
            showStatus(`Processing Error: ${error.message}`, 'error');
        } finally {
            processBtn.classList.remove('loading');
            processBtn.disabled = false;
        }
    });

    function readFileAsArrayBuffer(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(new Uint8Array(e.target.result));
            reader.onerror = (e) => reject(new Error('Failed to read the file.'));
            reader.readAsArrayBuffer(file);
        });
    }

    function processDataRowByRow(rows) {
        if (rows.length < 2) return; // Need at least header and 1 data row

        const headerRow = rows[0];
        // Ensure every original row has elements to cover at least the header length.
        // This is necessary because empty trailing cells are sometimes omitted by SheetJS
        const originalColCount = headerRow.length;

        // 1. Array index manipulation maps perfectly to the user's manual column manipulations
        for (let r = 0; r < rows.length; r++) {
            let row = rows[r];
            while (row.length < originalColCount) row.push(null);

            // STEP 1: Cut Col I ("INLAND ORIGIN" origin index 8) and insert at Col B (index 1)
            let inlandOrig = row.splice(8, 1)[0];
            row.splice(1, 0, inlandOrig);

            // STEP 2: Cut Col L ("INLAND DEST" origin index 11) and insert at Col E (index 4)
            // Note: origin index 11 became index 10 when 8 was removed, but became 11 again when inserted at 1.
            let inlandDest = row.splice(11, 1)[0];
            row.splice(4, 0, inlandDest);
        }

        // The new absolute array indices for the remaining steps:
        const colB = 1;  // INLAND ORIGIN
        const colC = 2;  // BASE ORIGIN
        const colD = 3;  // BASE DESTINATION
        const colE = 4;  // INLAND DEST
        const colK = 10; // ORIGIN ROUTING (The data to pull for Step 4)
        const colL = 11; // DESTINATION ROUTING (The data to pull for Step 6)

        // Process actual data rows
        for (let r = 1; r < rows.length; r++) {
            let row = rows[r];

            // Helper to check for blanks
            const isBlank = (val) => val === null || val === undefined || String(val).trim() === '';

            // STEP 3: If B is blank, move C to B, and clear C
            if (isBlank(row[colB])) {
                row[colB] = row[colC];
                row[colC] = null;
            }

            // STEP 4: If C is blank AND B contains "Ramp", move K to C, and clear K
            if (isBlank(row[colC]) && !isBlank(row[colB]) && String(row[colB]).toLowerCase().includes('ramp')) {
                row[colC] = row[colK];
                row[colK] = null;
            }

            // STEP 5: If E is blank, move D to E, and clear D
            if (isBlank(row[colE])) {
                row[colE] = row[colD];
                row[colD] = null;
            }

            // STEP 6: If D is blank, move L to D, and clear L
            if (isBlank(row[colD])) {
                row[colD] = row[colL];
                row[colL] = null;
            }
        }

        // Add BOL Header (Step 10)
        rows[0].splice(14, 0, "BOL");

        // Execute Steps 7 to 11
        for (let r = 1; r < rows.length; r++) {
            let row = rows[r];

            // Insert null placeholder for BOL column to maintain column index integrity
            row.splice(14, 0, null);

            // STEP 7: Concatenated value in Column A (index 0)
            const b = String(row[1] || '').trim();
            const c = String(row[2] || '').trim();
            const d = String(row[3] || '').trim();
            const e = String(row[4] || '').trim();
            const g = String(row[6] || '').trim();
            const i = String(row[8] || '').trim();
            const m = String(row[12] || '').trim();
            const n = String(row[13] || '').trim();
            const h = String(row[7] || '').trim();
            row[0] = `${b}${c}${d}${e}${g}${i}${m}${n}${h}`;

            // STEP 9: Clean F (index 5)
            if (row[5] !== null && row[5] !== undefined) {
                row[5] = extractNumber(row[5]);
            }

            // STEP 8 & 11: Clean prices P to end (index 15 to end), and sum into BOL (index 14)
            let bolSum = null;
            for (let cIdx = 15; cIdx < row.length; cIdx++) {
                if (row[cIdx] !== null && row[cIdx] !== undefined) {
                    let cleanNum = extractNumber(row[cIdx]);
                    row[cIdx] = cleanNum; // Convert to true numeric
                    if (cleanNum !== null) {
                        bolSum = (bolSum || 0) + cleanNum;
                    }
                }
            }
            if (bolSum !== null) {
                row[14] = bolSum;
            }
        }
    }

    function extractNumber(val) {
        if (val === null || val === undefined) return null;
        if (typeof val === 'number') return val;
        const str = String(val).replace(/,/g, '');
        const numMatch = str.match(/-?\d+(?:\.\d+)?/);
        return numMatch ? parseFloat(numMatch[0]) : null;
    }

    function createMainFreightSheet(quoteCartRows, flatMarkup, markupPct) {
        if (!quoteCartRows || quoteCartRows.length === 0) return [];

        // Map columns A, B, C, D, E, G, I, M, O (which is BOL at index 14)
        const headerMap = [0, 1, 2, 3, 4, 6, 8, 12, 14];

        let mfRows = [];
        // Step 13: Copy mapped columns
        for (let r = 0; r < quoteCartRows.length; r++) {
            let row = quoteCartRows[r];
            let newRow = [];
            for (let idx of headerMap) {
                newRow.push(row[idx]);
            }
            mfRows.push(newRow);
        }

        // Step 14: Omit container texts in Column A
        const removeTexts = ["Cntr. 45 ft HC", "Cntr. 40 ft HC", "Cntr. 40 ft", "Cntr. 20 ft"];
        for (let r = 1; r < mfRows.length; r++) {
            if (typeof mfRows[r][0] === 'string') {
                let cellVal = mfRows[r][0];
                for (let rt of removeTexts) {
                    cellVal = cellVal.split(rt).join('');
                }
                mfRows[r][0] = cellVal.trim();
            }
        }

        // Step 15: Remove duplicate based on column A
        let uniqueMfRows = [mfRows[0]];
        let seenA = new Set();
        for (let r = 1; r < mfRows.length; r++) {
            let key = mfRows[r][0];
            if (!seenA.has(key)) {
                seenA.add(key);
                uniqueMfRows.push(mfRows[r]);
            }
        }
        mfRows = uniqueMfRows;

        // Step 16: Add headers for column J to M
        const colJ = 9, colK = 10, colL = 11, colM = 12;
        mfRows[0][colJ] = "Cntr. 20 ft";
        mfRows[0][colK] = "Cntr. 40 ft";
        mfRows[0][colL] = "Cntr. 40 ft HC";
        mfRows[0][colM] = "Cntr. 45 ft HC";

        // Step 19: Add headers in row 1 for column N to Q
        const colN = 13, colO = 14, colP = 15, colQ = 16;
        mfRows[0][colN] = "Total-Cntr. 20 ft";
        mfRows[0][colO] = "Total-Cntr. 40 ft";
        mfRows[0][colP] = "Total-Cntr. 40 ft HC";
        mfRows[0][colQ] = "Total-Cntr. 45 ft HC";

        // Pre-build a map for the 'VLOOKUP' simulation from Quote Cart
        let quoteCartLookup = new Map();
        for (let r = 1; r < quoteCartRows.length; r++) {
            let key = String(quoteCartRows[r][0] || '');
            let valF = quoteCartRows[r][5]; // Column F is index 5
            quoteCartLookup.set(key, valF);
        }

        // Step 17, 18, 20 logic
        const markupMult = 1 + (markupPct / 100);

        for (let r = 1; r < mfRows.length; r++) {
            let aVal = String(mfRows[r][0] || ''); // MainFreight Col A
            let bolVal = extractNumber(mfRows[r][8]) || 0; // mfRows Col I (index 8) is BOL mapped from Orig O (14)

            // J to M Vlookup simulation
            mfRows[r][colJ] = quoteCartLookup.get(aVal + "Cntr. 20 ft") ?? null;
            mfRows[r][colK] = quoteCartLookup.get(aVal + "Cntr. 40 ft") ?? null;
            mfRows[r][colL] = quoteCartLookup.get(aVal + "Cntr. 40 ft HC") ?? null;
            mfRows[r][colM] = quoteCartLookup.get(aVal + "Cntr. 45 ft HC") ?? null;

            // Step 20: Calculations
            const calcTotal = (cntrVal) => {
                if (cntrVal === null || cntrVal === undefined || String(cntrVal).trim() === '') return null;
                const base = bolVal + Number(cntrVal) + flatMarkup;
                return base * markupMult; // Which is base * (1 + markupPct/100)
            };

            mfRows[r][colN] = calcTotal(mfRows[r][colJ]);
            mfRows[r][colO] = calcTotal(mfRows[r][colK]);
            mfRows[r][colP] = calcTotal(mfRows[r][colL]);
            mfRows[r][colQ] = calcTotal(mfRows[r][colM]);
        }

        return mfRows;
    }
});
