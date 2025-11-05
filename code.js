function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu("🤘🏻 TOOLS BOS SETU")
        .addItem("🖨️ Cetak PDF Siswa/Siswi", "showSidebar")
        .addToUi();
}

function showSidebar() {
    const html =
        HtmlService.createHtmlOutputFromFile("sidebar")
        .setTitle("Cetak Multi PDF Rapor");
    SpreadsheetApp.getUi().showSidebar(html);
}


