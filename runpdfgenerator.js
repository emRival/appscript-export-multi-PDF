function runPDFGenerator(formData) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const folder = DriveApp.getFolderById(formData.folderId);
    const start = parseInt(formData.start);
    const end = parseInt(formData.end);
    const merge = formData.merge || false;

    for (let i = start; i <= end; i++) {
        // 1. Set Nomor Urut
        ss.getSheetByName(formData.numberRefSheet)
            .getRange(formData.numberRefCell)
            .setValue(i);

        SpreadsheetApp.flush();
        Utilities.sleep(500); // Jeda agar rumus update

        // 2. Ambil Nama Siswa
        let namaSiswa = ss
            .getSheetByName(formData.nameRefSheet)
            .getRange(formData.nameRefCell)
            .getDisplayValue()
            .trim()
            .replace(/[\/\\?%*:|"<>]/g, "_"); // Bersihkan karakter ilegal filename
        
        if (!namaSiswa) namaSiswa = "Tanpa_Nama";

        // === LOGIKA MERGE PDF ===
        if (merge) {
            const tempSS = SpreadsheetApp.create(`TEMP_PDF_${i}`);
            
            formData.sheetsData.forEach((sheetInfo, idx) => {
                const sourceSheet = ss.getSheetByName(sheetInfo.sheetName);
                if (!sourceSheet) return;

                let sourceRange;
                try {
                    sourceRange = sourceSheet.getRange(sheetInfo.range);
                } catch (err) { return; }

                const targetSheet = idx === 0
                    ? tempSS.getActiveSheet().setName(`Page ${idx + 1}`)
                    : tempSS.insertSheet(`Page ${idx + 1}`);

                copyFullVisualRange(sourceRange, targetSheet);
            });

            SpreadsheetApp.flush();
            
            // Gunakan Alias sheet pertama untuk nama file Merge (Opsional)
            const mainAlias = formData.sheetsData[0].alias 
                ? formData.sheetsData[0].alias 
                : "Laporan";
                
            const pdfName = `${mainAlias} - ${namaSiswa} - ${i}`;
            const pdf = exportSpreadsheetAsPDF(tempSS, pdfName);
            folder.createFile(pdf);
            DriveApp.getFileById(tempSS.getId()).setTrashed(true);

        } else { 
            // === LOGIKA SINGLE PDF (PER SHEET) ===
            const token = ScriptApp.getOAuthToken();
            const spreadsheetId = ss.getId();

            formData.sheetsData.forEach((sheetInfo) => {
                const sheet = ss.getSheetByName(sheetInfo.sheetName);
                if (!sheet) return;

                // --- PERBAIKAN DI SINI (LOGIKA ALIAS) ---
                let fileName;
                if (sheetInfo.alias && sheetInfo.alias.trim() !== "") {
                    // Jika Alias diisi: "Alias - Nama.pdf"
                    // (Nomor urut 'i' saya hilangkan agar lebih bersih, atau bisa ditambahkan jika perlu)
                    fileName = `${sheetInfo.alias} - ${namaSiswa}.pdf`; 
                } else {
                    // Default: "[SheetName] - Nama.pdf"
                    fileName = `[${sheetInfo.sheetName}] - ${namaSiswa}.pdf`;
                }
                // ----------------------------------------

                const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=pdf&size=A4&portrait=true&fitw=true&scale=4&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false&horizontal_alignment=CENTER&gid=${sheet.getSheetId()}&range=${encodeURIComponent(sheetInfo.range)}`;

                const response = UrlFetchApp.fetch(url, {
                    headers: { Authorization: "Bearer " + token },
                });

                folder.createFile(response.getBlob().setName(fileName));
                Utilities.sleep(1000); // Jeda anti-limit
            });
        }
    }
}

function copyFullVisualRange(sourceRange, targetSheet) {
    const data = sourceRange.getValues();

    const r = targetSheet.getRange(1, 1, data.length, data[0].length);

    r.setValues(data);

    r.setBackgrounds(sourceRange.getBackgrounds());

    r.setFontColors(sourceRange.getFontColors());

    r.setFontWeights(sourceRange.getFontWeights());

    r.setFontStyles(sourceRange.getFontStyles());

    r.setFontSizes(sourceRange.getFontSizes());

    r.setHorizontalAlignments(sourceRange.getHorizontalAlignments());

    r.setVerticalAlignments(sourceRange.getVerticalAlignments());

    r.setWraps(sourceRange.getWraps());

    r.setFontFamilies(sourceRange.getFontFamilies());

    // r.setBorder(true, true, true, true, true, true);

    for (let col = 1; col <= data[0].length; col++) {
        const width = sourceRange
            .getSheet()
            .getColumnWidth(sourceRange.getColumn() + col - 1);

        targetSheet.setColumnWidth(col, width);
    }

    for (let row = 1; row <= data.length; row++) {
        const height = sourceRange
            .getSheet()
            .getRowHeight(sourceRange.getRow() + row - 1);

        targetSheet.setRowHeight(row, height);
    }

    targetSheet.setFrozenRows(sourceRange.getSheet().getFrozenRows());

    targetSheet.setFrozenColumns(sourceRange.getSheet().getFrozenColumns());

    copyMergedRanges(sourceRange, targetSheet);
}

function copyMergedRanges(
    sourceRange,
    targetSheet,
    targetRowOffset = 0,
    targetColOffset = 0
) {
    if (!sourceRange || typeof sourceRange.getMergedRanges !== "function")
        return;

    let mergedRanges;

    try {
        mergedRanges = sourceRange.getMergedRanges();
    } catch (e) {
        Logger.log("⚠️ Gagal ambil merged ranges: " + e);

        return;
    }

    mergedRanges.forEach((range) => {
        try {
            const numRows = range.getNumRows();

            const numCols = range.getNumColumns();

            const rowOffset =
                range.getRow() - sourceRange.getRow() + 1 + targetRowOffset;

            const colOffset =
                range.getColumn() -
                sourceRange.getColumn() +
                1 +
                targetColOffset;

            targetSheet
                .getRange(rowOffset, colOffset, numRows, numCols)
                .merge();
        } catch (e) {
            Logger.log("⚠️ Gagal merge cell: " + e);
        }
    });
}

function exportSpreadsheetAsPDF(spreadsheet, filename) {
    const url = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/export?format=pdf&exportFormat=pdf&size=A4&portrait=true&fitw=true&scale=4&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false&horizontal_alignment=CENTER`;

    const token = ScriptApp.getOAuthToken();

    const response = UrlFetchApp.fetch(url, {
        headers: { Authorization: "Bearer " + token },
    });

    return response.getBlob().setName(filename + ".pdf");
}


function copyFullVisualRange(sourceRange, targetSheet) {
    const data = sourceRange.getValues();

    const r = targetSheet.getRange(1, 1, data.length, data[0].length);

    r.setValues(data);

    r.setBackgrounds(sourceRange.getBackgrounds());

    r.setFontColors(sourceRange.getFontColors());

    r.setFontWeights(sourceRange.getFontWeights());

    r.setFontStyles(sourceRange.getFontStyles());

    r.setFontSizes(sourceRange.getFontSizes());

    r.setHorizontalAlignments(sourceRange.getHorizontalAlignments());

    r.setVerticalAlignments(sourceRange.getVerticalAlignments());

    r.setWraps(sourceRange.getWraps());

    r.setFontFamilies(sourceRange.getFontFamilies());

    // r.setBorder(true, true, true, true, true, true);

    for (let col = 1; col <= data[0].length; col++) {
        const width = sourceRange
            .getSheet()
            .getColumnWidth(sourceRange.getColumn() + col - 1);

        targetSheet.setColumnWidth(col, width);
    }

    for (let row = 1; row <= data.length; row++) {
        const height = sourceRange
            .getSheet()
            .getRowHeight(sourceRange.getRow() + row - 1);

        targetSheet.setRowHeight(row, height);
    }

    targetSheet.setFrozenRows(sourceRange.getSheet().getFrozenRows());

    targetSheet.setFrozenColumns(sourceRange.getSheet().getFrozenColumns());

    copyMergedRanges(sourceRange, targetSheet);
}

function copyMergedRanges(
    sourceRange,
    targetSheet,
    targetRowOffset = 0,
    targetColOffset = 0
) {
    if (!sourceRange || typeof sourceRange.getMergedRanges !== "function")
        return;

    let mergedRanges;

    try {
        mergedRanges = sourceRange.getMergedRanges();
    } catch (e) {
        Logger.log("⚠️ Gagal ambil merged ranges: " + e);

        return;
    }

    mergedRanges.forEach((range) => {
        try {
            const numRows = range.getNumRows();

            const numCols = range.getNumColumns();

            const rowOffset =
                range.getRow() - sourceRange.getRow() + 1 + targetRowOffset;

            const colOffset =
                range.getColumn() -
                sourceRange.getColumn() +
                1 +
                targetColOffset;

            targetSheet
                .getRange(rowOffset, colOffset, numRows, numCols)
                .merge();
        } catch (e) {
            Logger.log("⚠️ Gagal merge cell: " + e);
        }
    });
}

function exportSpreadsheetAsPDF(spreadsheet, filename) {
    const url = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/export?format=pdf&exportFormat=pdf&size=A4&portrait=true&fitw=true&scale=4&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false&horizontal_alignment=CENTER`;

    const token = ScriptApp.getOAuthToken();

    const response = UrlFetchApp.fetch(url, {
        headers: { Authorization: "Bearer " + token },
    });

    return response.getBlob().setName(filename + ".pdf");
}

