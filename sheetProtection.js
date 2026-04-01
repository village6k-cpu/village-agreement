function protectSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var results = [];

  // 거래내역 시트 보호
  var txSheet = ss.getSheetByName('거래내역');
  if (txSheet) {
    var prots = txSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    for (var i = 0; i < prots.length; i++) prots[i].remove();
    var prot = txSheet.protect().setDescription('거래내역 수식 보호');
    var lastRow = Math.max(txSheet.getLastRow(), 1000);
    prot.setUnprotectedRanges([
      txSheet.getRange('A2:F' + lastRow),
      txSheet.getRange('H2:O' + lastRow)
    ]);
    prot.setWarningOnly(false);
    results.push('거래내역: 보호 완료 (G열 수식 + 헤더 보호, A-F/H-O 입력 허용)');
  }

  // 장비체크 시트 보호
  var eqSheet = ss.getSheetByName('장비체크');
  if (eqSheet) {
    var prots2 = eqSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    for (var i = 0; i < prots2.length; i++) prots2[i].remove();
    var prot2 = eqSheet.protect().setDescription('장비체크 수식 보호');
    var lastRow2 = Math.max(eqSheet.getLastRow(), 100);
    var lastCol2 = eqSheet.getLastColumn();
    var unprotected2 = [eqSheet.getRange('A2:G' + lastRow2)];
    if (lastCol2 > 9) unprotected2.push(eqSheet.getRange(2, 10, lastRow2 - 1, lastCol2 - 9));
    prot2.setUnprotectedRanges(unprotected2);
    prot2.setWarningOnly(false);
    results.push('장비체크: 보호 완료 (H-I열 수식 보호, A-G열 입력 허용)');
  }

  // 고객DB / 발령지DB 헤더 보호
  var dataSheets = ['고객DB', '발령지DB'];
  for (var d = 0; d < dataSheets.length; d++) {
    var dSheet = ss.getSheetByName(dataSheets[d]);
    if (!dSheet) continue;
    var prots3 = dSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    for (var i = 0; i < prots3.length; i++) prots3[i].remove();
    var prot3 = dSheet.protect().setDescription(dataSheets[d] + ' 헤더 보호');
    var lr = Math.max(dSheet.getLastRow(), 500);
    prot3.setUnprotectedRanges([dSheet.getRange(2, 1, lr - 1, dSheet.getLastColumn())]);
    prot3.setWarningOnly(false);
    results.push(dataSheets[d] + ': 헤더 보호 완료');
  }

  Logger.log(results.join('\n'));
}
