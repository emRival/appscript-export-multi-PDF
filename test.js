function testRunPDFGenerator() {
  const testData = {
    folderId: '1dXdiSeHzQAYLIwMs_z49hBWOESKy5xhO',
    nameRefSheet: '7A - Report 1',
    nameRefCell: 'D5',
    numberRefSheet: '7A - Report 1',
    numberRefCell: 'K1',
    start: 1,
    end: 1,
    merge: true,
    sheetsData: [
      { sheetName: '7A - Report 1', range: 'A1:I23' },
      { sheetName: '7A - Report 2', range: 'A1:G29' }
    ]
  };
  runPDFGenerator(testData);
}
// 7A - Report 1|A1:I23
// 7A - Report 2|A1:G29