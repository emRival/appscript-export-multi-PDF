function runPDFGenerator(formData) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const folder = DriveApp.getFolderById(formData.folderId);

    const start = parseInt(formData.start);

    const end = parseInt(formData.end);

    const merge = formData.merge || false;

    for (let i = start; i <= end; i++) {
        ss.getSheetByName(formData.numberRefSheet)
            .getRange(formData.numberRefCell)
            .setValue(i);

        SpreadsheetApp.flush();

        Utilities.sleep(500);

        const namaSiswa = ss
            .getSheetByName(formData.nameRefSheet)

            .getRange(formData.nameRefCell)

            .getDisplayValue()

            .trim()

            .replace(/\s+/g, "_");

        if (merge) {
            const tempSS = SpreadsheetApp.create(`TEMP_PDF_${i}`);

            formData.sheetsData.forEach((sheetInfo, idx) => {
                const sourceSheet = ss.getSheetByName(sheetInfo.sheetName);

                if (!sourceSheet) {
                    Logger.log(
                        `❌ Sheet tidak ditemukan: ${sheetInfo.sheetName}`
                    );

                    return;
                }

                let sourceRange;

                try {
                    sourceRange = sourceSheet.getRange(sheetInfo.range);
                } catch (err) {
                    Logger.log(
                        `❌ Gagal membaca range "${sheetInfo.range}" di "${sheetInfo.sheetName}": ${err}`
                    );

                    return;
                }

                if (!sourceRange.getNumRows() || !sourceRange.getNumColumns()) {
                    Logger.log(
                        `⚠️ Range kosong: ${sheetInfo.sheetName} - ${sheetInfo.range}`
                    );

                    return;
                }

                const targetSheet =
                    idx === 0
                        ? tempSS.getActiveSheet().setName(`Halaman ${idx + 1}`)
                        : tempSS.insertSheet(`Halaman ${idx + 1}`);

                copyFullVisualRange(sourceRange, targetSheet);
            });

            SpreadsheetApp.flush();

            const pdf = exportSpreadsheetAsPDF(
                tempSS,
                `Rapor - ${namaSiswa} - ${i}`
            );

            folder.createFile(pdf);

            DriveApp.getFileById(tempSS.getId()).setTrashed(true);
        } else {
            const token = ScriptApp.getOAuthToken();

            const spreadsheetId = ss.getId();

            formData.sheetsData.forEach((sheetInfo) => {
                const sheet = ss.getSheetByName(sheetInfo.sheetName);

                if (!sheet) return;

                const fileName = `[${sheetInfo.sheetName}] - ${namaSiswa} - ${i}.pdf`;

                const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=pdf&size=A4&portrait=true&fitw=true&scale=4&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false&horizontal_alignment=CENTER&gid=${sheet.getSheetId()}&range=${encodeURIComponent(
                    sheetInfo.range
                )}`;

                const response = UrlFetchApp.fetch(url, {
                    headers: { Authorization: "Bearer " + token },
                });

                folder.createFile(response.getBlob().setName(fileName));

                Utilities.sleep(1000);
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

