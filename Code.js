// ─── 설정값 ───────────────────────────────────────────────────
var PB_LINKID    = "BILL";
var PB_SECRETKEY = "tBVY2va4UNSEkt3dXtjYvvANlhHoUcu1/muEGU6trzQ=";
var PB_CORP_NUM  = "6902901012";
var INVOICER_CORP_NAME = "빌(BILL.)";
var INVOICER_CEO_NAME  = "최재형";

var BANK_INFO    = "우리은행 1005-404-109661 최재형(빌리지)";
var BANK_CODE    = "0020";
var ACCOUNT_NUM  = "1005404109661";

var SH_NAME           = "거래내역";
var MASTER_DB_NAME    = "발행처DB";
var CUSTOMER_DB_NAME  = "고객DB";
var CHECK_SHEET_NAME  = "장비체크";

var FOLDER_ID    = "12MifXuoED1jBYM-o2qwokb3-93jt9z0N";
var TEMPLATE_ID  = "18V9oJT8TiGBuaBOJQm-c433O5A0GCH66WvjD8OaVTRE";

var CONTRACT_AMOUNT_CELL  = "H47";
var CONTRACT_TEMPLATE_CODE = "026030001244";

var COL_DATE=1, COL_NAME=2, COL_PAYER_NAME=3, COL_TR_ID=4,
    COL_BOOKER_ID=5, COL_BUS_NAME=6, COL_BUS_ID=7, COL_AMOUNT=8,
    COL_PAY_METHOD=9, COL_DOC_TYPE=10, COL_ISSUE=11,
    COL_STATUS=12, COL_LINK=13, COL_NOTE=14, COL_MGTKEY=15;

// ─────────────────────────────────────────────────────────────

function normalizePhone(phone) {
  var digits = String(phone).replace(/[^0-9]/g, "");
  if (digits.length === 10) digits = "0" + digits;
  if (digits.length === 11) {
    return digits.substring(0,3) + "-" + digits.substring(3,7) + "-" + digits.substring(7,11);
  }
  return digits;
}

// ─── 발행처DB 중복 합치기 (시트 수정 시 자동 실행) ────────────
function mergeMasterDB(e) {
  if (e && e.source.getActiveSheet().getName() !== MASTER_DB_NAME) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MASTER_DB_NAME);
  var data = sheet.getDataRange().getValues();

  var map = {}, order = [];
  for (var i = 1; i < data.length; i++) {
    var bizNum = String(data[i][0]).replace(/-/g, "").trim();
    if (!bizNum) continue;
    if (!map[bizNum]) {
      map[bizNum] = { bizNum: data[i][0], name: data[i][1], owner: data[i][2], emails: [], address: data[i][4] || "" };
      order.push(bizNum);
    }
    var email = String(data[i][3]).trim();
    if (email && map[bizNum].emails.indexOf(email) === -1) map[bizNum].emails.push(email);
  }

  var header = data[0];
  var newData = [header];
  for (var j = 0; j < order.length; j++) {
    var item = map[order[j]];
    newData.push([item.bizNum, item.name, item.owner, item.emails.join(", "), item.address]);
  }
  sheet.clearContents();
  sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
}

// ─── 거래내역 → 장비체크 전체 동기화 (1회 실행) ───────────────
function syncAllToEquipCheck() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tradeSheet = ss.getSheetByName(SH_NAME);
  var checkSheet = ss.getSheetByName(CHECK_SHEET_NAME);
  if (!tradeSheet || !checkSheet) return;

  var tradeData = tradeSheet.getDataRange().getValues();
  var checkData = checkSheet.getDataRange().getValues();

  var checkMap = {};
  for (var i = 1; i < checkData.length; i++) {
    var id = String(checkData[i][0]).trim();
    if (id && id !== 'undefined') checkMap[id] = { row: i + 1, name: String(checkData[i][1]).trim() };
  }

  var lastRow = 1;
  for (var i = checkData.length - 1; i >= 1; i--) {
    if (String(checkData[i][0]).trim() && String(checkData[i][0]).trim() !== 'undefined') { lastRow = i + 1; break; }
  }

  var newRows = [], updateCount = 0;
  for (var i = 1; i < tradeData.length; i++) {
    var trID = String(tradeData[i][COL_TR_ID - 1]).trim();
    var clientName = String(tradeData[i][COL_NAME - 1]).trim();
    if (!trID || trID === 'undefined') continue;
    if (checkMap[trID]) {
      if (checkMap[trID].name !== clientName) { checkSheet.getRange(checkMap[trID].row, 2).setValue(clientName); updateCount++; }
    } else {
      newRows.push([trID, clientName, "", "", ""]);
    }
  }
  if (newRows.length > 0) checkSheet.getRange(lastRow + 1, 1, newRows.length, 5).setValues(newRows);
  Logger.log("장비체크 동기화: 추가 " + newRows.length + "건, 업데이트 " + updateCount + "건");
}

// ─── 계약서 금액 동기화 (5분 트리거) ─────────────────────────
function syncAmounts() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_NAME);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var linkUrl = String(data[i][COL_LINK - 1]);
    if (!linkUrl || linkUrl === "" || linkUrl === "undefined") continue;
    // 발행완료된 건은 금액 변경 불필요
    var issueStatus = String(data[i][COL_ISSUE - 1]);
    if (issueStatus === "발행완료") continue;
    try {
      var match = linkUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (!match) continue;
      var amount = SpreadsheetApp.openById(match[1]).getSheets()[0].getRange(CONTRACT_AMOUNT_CELL).getValue();
      if (typeof amount === "number" && amount > 0) {
        if (String(data[i][COL_PAY_METHOD - 1]) === "계좌이체(VAT별도)") amount = Math.round(amount / 1.1);
        var currentAmount = data[i][COL_AMOUNT - 1];
        if (currentAmount !== amount) sheet.getRange(i + 1, COL_AMOUNT).setValue(amount);
      }
    } catch (err) { Logger.log("행 " + (i + 1) + " 금액 동기화 실패: " + err.message); }
  }
}

// ─── 계약서 링크 자동 연결 (5분 트리거) ─────────────────────
function linkExistingContracts() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_NAME);
  var data = sheet.getDataRange().getValues();
  var folder = DriveApp.getFolderById(FOLDER_ID);

  // 파일 목록을 한 번만 가져와서 맵으로 변환 (거래ID → {url, id})
  var fileMap = {};
  var allFiles = folder.getFiles();
  while (allFiles.hasNext()) {
    var f = allFiles.next();
    if (f.getMimeType() !== "application/vnd.google-apps.spreadsheet") continue;
    var fName = f.getName();
    var fTrID = fName.split("_")[0];  // 파일명: "거래ID_고객명_계약서"
    if (fTrID && !fileMap[fTrID]) {
      fileMap[fTrID] = { url: f.getUrl(), id: f.getId() };
    }
  }

  for (var i = 1; i < data.length; i++) {
    var trID = String(data[i][COL_TR_ID - 1]).trim();
    var existingLink = String(data[i][COL_LINK - 1]).trim();
    if (!trID || trID === "undefined") continue;
    if (existingLink && existingLink !== "undefined") continue;

    var matched = fileMap[trID];
    if (matched) {
      sheet.getRange(i + 1, COL_LINK).setValue(matched.url);
      try {
        var amount = SpreadsheetApp.openById(matched.id).getSheets()[0].getRange(CONTRACT_AMOUNT_CELL).getValue();
        if (typeof amount === "number" && amount > 0) {
          if (String(data[i][COL_PAY_METHOD - 1]) === "계좌이체(VAT별도)") amount = Math.round(amount / 1.1);
          sheet.getRange(i + 1, COL_AMOUNT).setValue(amount);
        }
      } catch (err) { Logger.log("금액 동기화 실패 [행 " + (i + 1) + "]: " + err.message); }
    }
  }
}

// ─── 입금 매칭 (5분 트리거) ───────────────────────────────────
function matchDeposits() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_NAME);
  var data = sheet.getDataRange().getValues();
  var pendingRows = [];

  for (var i = 1; i < data.length; i++) {
    var status = String(data[i][COL_STATUS - 1]);
    var amount = data[i][COL_AMOUNT - 1];
    var payerName = String(data[i][COL_PAYER_NAME - 1]).trim();
    var matchName = (payerName === "" || payerName === "undefined") ? String(data[i][COL_NAME - 1]) : payerName;
    if (status !== "입금완료" && amount > 0 && matchName) pendingRows.push({ row: i + 1, name: matchName, amount: amount });
  }
  if (pendingRows.length === 0) return;

  try {
    var accessToken = getPopbillAccessToken(["member", "180"]);
    var today = Utilities.formatDate(new Date(), "GMT+9", "yyyyMMdd");
    var weekAgo = Utilities.formatDate(new Date(new Date().getTime() - 7 * 24 * 60 * 60 * 1000), "GMT+9", "yyyyMMdd");

    var jobRes = UrlFetchApp.fetch(
      "https://popbill.linkhub.co.kr/EasyFin/Bank/BankAccount?BankCode=" + BANK_CODE + "&AccountNumber=" + ACCOUNT_NUM + "&SDate=" + weekAgo + "&EDate=" + today,
      { method: "POST", headers: { Authorization: "Bearer " + accessToken, "Content-Type": "application/json" }, muteHttpExceptions: true }
    );
    if (jobRes.getResponseCode() !== 200) return;
    var jobID = JSON.parse(jobRes.getContentText()).jobID;

    var jobState = 0;
    for (var t = 0; t < 5; t++) {
      Utilities.sleep(2000);
      var stateData = JSON.parse(UrlFetchApp.fetch("https://popbill.linkhub.co.kr/EasyFin/Bank/" + jobID + "/State",
        { method: "GET", headers: { Authorization: "Bearer " + accessToken }, muteHttpExceptions: true }).getContentText());
      jobState = stateData.jobState;
      if (jobState == 3) break;
    }
    if (jobState !== 3) return;

    var searchRes = UrlFetchApp.fetch("https://popbill.linkhub.co.kr/EasyFin/Bank/" + jobID + "?Page=1&PerPage=100&InOut=1",
      { method: "GET", headers: { Authorization: "Bearer " + accessToken }, muteHttpExceptions: true });
    if (searchRes.getResponseCode() !== 200) return;

    var transactions = JSON.parse(searchRes.getContentText()).list;
    if (!transactions || transactions.length === 0) return;

    for (var r = 0; r < pendingRows.length; r++) {
      var pending = pendingRows[r];
      for (var tr = 0; tr < transactions.length; tr++) {
        var tx = transactions[tr];
        if (String(tx.remark1 || "").trim().indexOf(pending.name) !== -1 && Number(tx.accIn || 0) === pending.amount) {
          sheet.getRange(pending.row, COL_STATUS).setValue("입금완료");
          break;
        }
      }
    }
  } catch (err) { Logger.log("matchDeposits 오류: " + err.message); }
}

// ─── onChange 트리거 ───────────────────────────────────────────
function autoContract(e) {
  if (!e) return;
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  if (sheet.getName() !== SH_NAME) return;

  var row = range.getRow();
  var col = range.getColumn();
  var val = range.getDisplayValue();

  if (col == COL_TR_ID) {
    range.setValue(e.oldValue || '');
    SpreadsheetApp.getActiveSpreadsheet().toast('거래아이디는 자동 생성됩니다.', '입력 차단', 3);
    return;
  }

  try {
    // A열(날짜) 입력 → D열 거래ID 자동 생성
    if (col == COL_DATE && val != "") {
      var dateRaw = val.replace(/-/g, "");
      if (dateRaw.length == 6) {
        range.setValue("20" + dateRaw.substring(0,2) + "-" + dateRaw.substring(2,4) + "-" + dateRaw.substring(4,6));
        SpreadsheetApp.flush();
        var count = 0;
        var data = sheet.getDataRange().getValues();
        for (var i = 1; i < data.length; i++) {
          if (i + 1 === row) continue;
          var cellDate = data[i][COL_DATE - 1];
          var rowDate = cellDate instanceof Date
            ? Utilities.formatDate(cellDate, "Asia/Seoul", "yyMMdd")
            : String(cellDate).replace(/-/g, "").slice(-6);
          if (rowDate === dateRaw) count++;
        }
        sheet.getRange(row, COL_TR_ID).setValue(dateRaw + "-" + ("00" + (count + 1)).slice(-3));
      }
    }

    // B열(예약자명) 입력 → 고객DB에서 전화번호 자동완성 + 계약서 생성
    if (col == COL_NAME && val != "") {
      var dbSheet = e.source.getSheetByName(CUSTOMER_DB_NAME);
      if (dbSheet) {
        var dbData = dbSheet.getDataRange().getValues();
        var bookerID = "", matchCount = 0;
        for (var i = 1; i < dbData.length; i++) {
          if (String(dbData[i][1]) === val) { bookerID = String(dbData[i][0]); matchCount++; }
        }
        if (matchCount === 0) {
          sheet.getRange(row, COL_NOTE).setValue("고객DB에 없는 고객입니다. 고객DB에 먼저 등록해주세요.");
        } else if (matchCount > 1) {
          sheet.getRange(row, COL_NOTE).setValue("동명이인 확인 필요: 고객DB에서 정확한 이름을 확인해주세요.");
        } else {
          sheet.getRange(row, COL_BOOKER_ID).setValue(normalizePhone(bookerID));
          sheet.getRange(row, COL_NOTE).setValue("");
        }
      }

      Utilities.sleep(500);
      var trID = sheet.getRange(row, COL_TR_ID).getDisplayValue();
      var clientName = sheet.getRange(row, COL_NAME).getDisplayValue();
      if (trID && clientName && !sheet.getRange(row, COL_LINK).getValue()) {
        sheet.getRange(row, COL_NOTE).setValue("로봇: 계약서 생성 중...");
        var file = DriveApp.getFileById(TEMPLATE_ID).makeCopy(
          trID + "_" + clientName + "_계약서",
          DriveApp.getFolderById(FOLDER_ID)
        );
        sheet.getRange(row, COL_LINK).setValue(file.getUrl());
        sheet.getRange(row, COL_NOTE).setValue("로봇: 계약서 생성 완료");
      }
    }

    // I열(결제수단) → 자동 상태 설정
    if (col == COL_PAY_METHOD && (val == "카드결제" || val == "현금")) {
      sheet.getRange(row, COL_DOC_TYPE).setValue("미발행");
      sheet.getRange(row, COL_ISSUE).setValue("발행완료");
      sheet.getRange(row, COL_STATUS).setValue("입금완료");
    } else if (col == COL_PAY_METHOD && val == "계좌이체(VAT별도)") {
      sheet.getRange(row, COL_DOC_TYPE).setValue("미발행");
      sheet.getRange(row, COL_ISSUE).setValue("발행완료");
      sheet.getRange(row, COL_STATUS).setValue("미입금");
    }

    // K열(발행요청) → 증빙 발행
    if (col == COL_ISSUE && val == "발행요청") {
      if (sheet.getRange(row, COL_ISSUE).getValue() == "발행완료") {
        sheet.getRange(row, COL_NOTE).setValue("이미 발행완료된 건입니다.");
        return;
      }

      // 발행 직전 금액 동기화
      sheet.getRange(row, COL_NOTE).setValue("로봇: 금액 동기화 중...");
      var linkUrl = sheet.getRange(row, COL_LINK).getValue();
      if (linkUrl) {
        var matchLink = linkUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
        if (matchLink) {
          try {
            var latestAmount = SpreadsheetApp.openById(matchLink[1]).getSheets()[0].getRange(CONTRACT_AMOUNT_CELL).getValue();
            if (typeof latestAmount === "number" && latestAmount > 0) {
              if (sheet.getRange(row, COL_PAY_METHOD).getValue() === "계좌이체(VAT별도)") latestAmount = Math.round(latestAmount / 1.1);
              sheet.getRange(row, COL_AMOUNT).setValue(latestAmount);
            }
          } catch (syncErr) { Logger.log("금액 동기화 실패: " + syncErr.message); }
        }
      }

      var docType = sheet.getRange(row, COL_DOC_TYPE).getValue();
      sheet.getRange(row, COL_NOTE).setValue("로봇: 팝빌 통신 시도...");

      if (docType == "세금계산서") {
        requestTaxInvoice(row);
        sheet.getRange(row, COL_ISSUE).setValue("발행완료");
        sheet.getRange(row, COL_NOTE).setValue("로봇: 세금계산서 발행 성공!");
      } else if (docType == "현금영수증(전화번호)") {
        requestCashbill(row, "전화번호");
        sheet.getRange(row, COL_ISSUE).setValue("발행완료");
        sheet.getRange(row, COL_NOTE).setValue("로봇: 현금영수증(전화번호) 발행 성공!");
      } else if (docType == "현금영수증(사업자번호)") {
        requestCashbill(row, "사업자번호");
        sheet.getRange(row, COL_ISSUE).setValue("발행완료");
        sheet.getRange(row, COL_NOTE).setValue("로봇: 현금영수증(사업자번호) 발행 성공!");
      } else {
        throw new Error("J열 증빙유형을 선택해주세요.");
      }
    }

    // 장비체크 실시간 동기화
    var syncTrID = sheet.getRange(row, COL_TR_ID).getDisplayValue();
    var syncName = sheet.getRange(row, COL_NAME).getDisplayValue();
    if (syncTrID) {
      var chkSh = e.source.getSheetByName(CHECK_SHEET_NAME);
      if (chkSh) {
        var chkIDs = chkSh.getRange("A:A").getValues();
        var foundChkRow = -1;
        for (var ci = 1; ci < chkIDs.length; ci++) {
          if (String(chkIDs[ci][0]).trim() === syncTrID) { foundChkRow = ci + 1; break; }
        }
        if (foundChkRow > 0) {
          chkSh.getRange(foundChkRow, 2).setValue(syncName);
        } else if (syncName) {
          var chkLastRow = 1;
          for (var ci2 = chkIDs.length - 1; ci2 >= 1; ci2--) {
            if (String(chkIDs[ci2][0]).trim() && String(chkIDs[ci2][0]).trim() !== 'undefined') { chkLastRow = ci2 + 1; break; }
          }
          chkSh.getRange(chkLastRow + 1, 1, 1, 5).setValues([[syncTrID, syncName, "", "", ""]]);
        }
      }
    }
    SpreadsheetApp.flush();
  } catch (err) {
    sheet.getRange(row, COL_ISSUE).setValue("전송실패");
    sheet.getRange(row, COL_NOTE).setValue("실패: " + err.message);
  }
}

// ─── 세금계산서 발행 ──────────────────────────────────────────
function requestTaxInvoice(row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_NAME);
  var masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_DB_NAME);
  var busID = sheet.getRange(row, COL_BUS_ID).getDisplayValue().replace(/-/g, "");
  var amount = sheet.getRange(row, COL_AMOUNT).getValue();
  var tradeDate = sheet.getRange(row, COL_DATE).getDisplayValue().replace(/-/g, "");
  var trID = sheet.getRange(row, COL_TR_ID).getDisplayValue();

  var mData = masterSheet.getDataRange().getValues();
  var p = null;
  for (var i = 1; i < mData.length; i++) {
    if (String(mData[i][0]).replace(/-/g, "") == busID) {
      p = { id: busID, name: mData[i][1], owner: mData[i][2], email: mData[i][3] };
      break;
    }
  }
  if (!p) throw new Error("발행처DB에 사업자번호 없음: " + busID);

  var accessToken = getPopbillAccessToken(["member", "110"]);
  var tax = Math.floor(amount / 11);
  var purposeType = (sheet.getRange(row, COL_STATUS).getValue() == "입금완료") ? "영수" : "청구";
  var mgtKey = Utilities.formatDate(new Date(), "GMT+9", "yyyyMMddHHmmss") + "-" + row;

  var emails = p.email ? p.email.split(",").map(function(e) { return e.trim(); }) : [];

  var payload = {
    issueType: "정발행", taxType: "과세", chargeDirection: "정과금",
    writeDate: Utilities.formatDate(new Date(), "GMT+9", "yyyyMMdd"),
    purposeType: purposeType,
    invoicerMgtKey: mgtKey, invoicerCorpNum: PB_CORP_NUM,
    invoicerCorpName: INVOICER_CORP_NAME, invoicerCEOName: INVOICER_CEO_NAME,
    invoiceeType: "사업자", invoiceeCorpNum: p.id,
    invoiceeCorpName: p.name, invoiceeCEOName: p.owner,
    invoiceeEmail1: emails[0] || "", invoiceeEmail2: emails[1] || "",
    supplyCostTotal: String(amount - tax), taxTotal: String(tax), totalAmount: String(amount),
    detailList: [{
      serialNum: 1, purchaseDT: tradeDate, itemName: "렌탈 (" + trID + ")",
      qty: "1", unitCost: String(amount - tax), supplyCost: String(amount - tax),
      tax: String(tax), remark: BANK_INFO
    }]
  };

  var res = UrlFetchApp.fetch("https://popbill.linkhub.co.kr/Taxinvoice", {
    method: "POST",
    headers: { Authorization: "Bearer " + accessToken, "Content-Type": "application/json", "X-HTTP-Method-Override": "ISSUE" },
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) {
    var msg = ""; try { msg = JSON.parse(res.getContentText()).message; } catch(e2) { msg = res.getContentText(); }
    throw new Error("세금계산서 HTTP" + res.getResponseCode() + ": " + msg);
  }
  var result = JSON.parse(res.getContentText());
  if (result.code !== 1) throw new Error("세금계산서 발행실패: " + result.message);
  sheet.getRange(row, COL_MGTKEY).setValue(mgtKey);
  return result;
}

// ─── 현금영수증 발행 ──────────────────────────────────────────
function requestCashbill(row, idType) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_NAME);
  var amount = sheet.getRange(row, COL_AMOUNT).getValue();
  var customerName = sheet.getRange(row, COL_NAME).getDisplayValue();
  var tradeDate = sheet.getRange(row, COL_DATE).getDisplayValue();
  var trID = sheet.getRange(row, COL_TR_ID).getDisplayValue();

  var identityNum, tradeUsage;
  if (idType == "전화번호") {
    identityNum = normalizePhone(sheet.getRange(row, COL_BOOKER_ID).getDisplayValue()).replace(/-/g, "");
    tradeUsage = "소득공제용";
    if (!identityNum) throw new Error("E열 휴대폰번호가 비어있습니다.");
  } else {
    identityNum = sheet.getRange(row, COL_BUS_ID).getDisplayValue().replace(/-/g, "");
    tradeUsage = "지출증빙용";
    if (!identityNum) throw new Error("G열 사업자번호가 비어있습니다.");
  }

  var accessToken = getPopbillAccessToken(["member", "154"]);
  var tax = Math.floor(amount / 11);
  var mgtKey = Utilities.formatDate(new Date(), "GMT+9", "yyyyMMddHHmmss") + "-CB-" + row;

  var payload = {
    mgtKey: mgtKey, tradeType: "승인거래", tradeUsage: tradeUsage, taxationType: "과세",
    totalAmount: String(amount), supplyCost: String(amount - tax), tax: String(tax), serviceFee: "0",
    franchiseCorpNum: PB_CORP_NUM, franchiseCorpName: INVOICER_CORP_NAME, franchiseCEOName: INVOICER_CEO_NAME,
    identityNum: identityNum, customerName: customerName,
    itemName: tradeDate + " 렌탈 (" + trID + ") / " + customerName
  };

  var res = UrlFetchApp.fetch("https://popbill.linkhub.co.kr/Cashbill", {
    method: "POST",
    headers: { Authorization: "Bearer " + accessToken, "Content-Type": "application/json", "X-HTTP-Method-Override": "ISSUE" },
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) {
    var msg = ""; try { msg = JSON.parse(res.getContentText()).message; } catch(e2) { msg = res.getContentText(); }
    throw new Error("현금영수증 HTTP" + res.getResponseCode() + ": " + msg);
  }
  var result = JSON.parse(res.getContentText());
  if (result.code !== 1) throw new Error("현금영수증 발행실패: " + result.message);

  var cbPhone = sheet.getRange(row, COL_BOOKER_ID).getDisplayValue().replace(/-/g, "");
  if (cbPhone) sendAlimtalk(cbPhone, customerName, tradeDate);
  return result;
}

// ─── 팝빌 AccessToken 발급 ────────────────────────────────────
function getPopbillAccessToken(scope) {
  var serviceID = "POPBILL";
  var forwardedIP = "*";
  var apiVersion = "2.0";
  var resourceURI = "/" + serviceID + "/Token";

  var bodyStr = JSON.stringify({ access_id: PB_CORP_NUM, scope: scope });
  var bodyBytes = Utilities.newBlob(bodyStr, "UTF-8").getBytes();
  var md5 = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bodyBytes));
  var date = Utilities.formatDate(new Date(), "UTC", "yyyy-MM-dd'T'HH:mm:ss'Z'");
  var stringToSign = "POST\n" + md5 + "\n" + date + "\n" + forwardedIP + "\n" + apiVersion + "\n" + resourceURI;
  var rawKey = Utilities.base64Decode(PB_SECRETKEY);
  var signature = Utilities.base64Encode(
    Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, Utilities.newBlob(stringToSign, "UTF-8").getBytes(), rawKey)
  );

  var res = UrlFetchApp.fetch("https://auth.linkhub.co.kr" + resourceURI, {
    method: "POST",
    headers: {
      Authorization: "LINKHUB " + PB_LINKID + " " + signature,
      "x-lh-date": date, "x-lh-version": apiVersion, "x-lh-forwarded": forwardedIP,
      "Content-Type": "application/json; charset=utf-8"
    },
    payload: bodyStr, muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) throw new Error("토큰발급 실패: " + res.getContentText());
  return JSON.parse(res.getContentText()).session_token;
}

// ─── 세금계산서/현금영수증 발급 완료 알림톡 ─────────────────
function sendAlimtalk(phone, customerName, tradeDate) {
  try {
    var accessToken = getPopbillAccessToken(["member", "153"]);
    var dateParts = String(tradeDate).replace(/-/g, "").match(/(\d{4})(\d{2})(\d{2})/);
    var formattedDate = dateParts ? dateParts[1] + "년 " + dateParts[2] + "월 " + dateParts[3] + "일" : String(tradeDate);

    var content = '"' + customerName + ' 감독님, 안녕하세요! 빌리지 렌탈샵입니다. ' + formattedDate + ' 렌탈 이용에 따른 세금계산서/현금영수증이 발급되었습니다. 감사합니다!"';

    var res = UrlFetchApp.fetch("https://popbill.linkhub.co.kr/ATS", {
      method: "POST",
      headers: { Authorization: "Bearer " + accessToken, "Content-Type": "application/json" },
      payload: JSON.stringify({
        templateCode: "026030000752", snd: "01071261139", altSendType: "C",
        msgs: [{ rcv: phone, rcvnm: customerName, msg: content, altmsg: content }]
      }),
      muteHttpExceptions: true
    });

    var result = JSON.parse(res.getContentText());
    return !!(result.receiptNum || result.code === 1);
  } catch (err) {
    Logger.log("sendAlimtalk 오류: " + err.message);
    return false;
  }
}

// ─── 세금계산서 국세청 승인 확인 후 알림톡 (1시간 트리거) ────
function checkTaxInvoiceStatus() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_NAME);
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COL_DOC_TYPE - 1]) !== "세금계산서") continue;
    if (String(data[i][COL_ISSUE - 1]) !== "발행완료") continue;
    var mgtKey = String(data[i][COL_MGTKEY - 1]).trim();
    if (!mgtKey || mgtKey === "" || mgtKey === "알림톡발송완료") continue;

    try {
      var accessToken = getPopbillAccessToken(["member", "110"]);
      var statusRes = UrlFetchApp.fetch("https://popbill.linkhub.co.kr/Taxinvoice/SELL/" + mgtKey,
        { method: "GET", headers: { Authorization: "Bearer " + accessToken }, muteHttpExceptions: true });
      if (statusRes.getResponseCode() !== 200) continue;

      var statusData = JSON.parse(statusRes.getContentText());
      if (statusData.stateCode >= 300) {
        var phone = String(data[i][COL_BOOKER_ID - 1]).replace(/-/g, "");
        var customerName = String(data[i][COL_NAME - 1]);
        var tradeDate = data[i][COL_DATE - 1];
        if (tradeDate instanceof Date) tradeDate = Utilities.formatDate(new Date(tradeDate), "Asia/Seoul", "yyyy-MM-dd");
        if (phone) {
          var sent = sendAlimtalk(phone, customerName, tradeDate);
          if (sent) sheet.getRange(i + 1, COL_MGTKEY).setValue("알림톡발송완료");
        }
      }
    } catch (err) { Logger.log("세금계산서 상태 확인 오류 [행 " + (i+1) + "]: " + err.message); }
  }
}

// ─── 커스텀 메뉴 ─────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📩 빌리지")
    .addItem("견적서 발송 패널 열기", "openContractSidebar")
    .addSeparator()
    .addItem("📋 약관 동의 링크 복사", "copyAgreementLink")
    .addItem("📋 약관 동의 알림톡 발송", "sendAgreementAlimtalk")
    .addToUi();
}

function openContractSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("견적서 카톡 발송")
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ─── 사이드바: 선택 행 정보 ───────────────────────────────────
function getSelectedRowInfo() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_NAME);
  if (!sheet) return { error: "'" + SH_NAME + "' 시트를 찾을 수 없습니다." };

  var row = sheet.getActiveCell().getRow();
  if (row < 2) return { error: "발송할 행을 선택해주세요.\n(2행 이상의 데이터 행을 클릭)" };

  var data = sheet.getRange(row, 1, 1, COL_MGTKEY).getValues()[0];
  var clientName = String(data[COL_NAME - 1]).trim();
  var phone = normalizePhone(data[COL_BOOKER_ID - 1]);
  var linkUrl = String(data[COL_LINK - 1]).trim();
  var trID = String(data[COL_TR_ID - 1]).trim();
  var tradeDate = data[COL_DATE - 1];
  var amount = data[COL_AMOUNT - 1];

  if (tradeDate instanceof Date) tradeDate = Utilities.formatDate(tradeDate, "Asia/Seoul", "yyyy-MM-dd");

  var hasLink = linkUrl && linkUrl.indexOf("docs.google.com") !== -1;

  return {
    row: row,
    clientName: clientName || "(비어있음)",
    phone: phone || "(비어있음)",
    trID: trID || "(비어있음)",
    tradeDate: String(tradeDate) || "(비어있음)",
    amount: amount ? Number(amount).toLocaleString() + "원" : "(비어있음)",
    hasLink: hasLink,
    canSend: !!(clientName && phone && phone.length >= 10 && hasLink)
  };
}

// ─── 사이드바 / AppSheet: 견적서 발송 실행 ───────────────────
function executeSendContract(row) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch(e) { return { success: false, message: "Lock timeout" }; }
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_NAME);
    var data = sheet.getRange(row, 1, 1, COL_MGTKEY).getValues()[0];

    // 중복 발송 방지 (견적서 발송 중이거나 이미 완료된 경우만 차단)
    var currentNote = String(data[COL_NOTE - 1]).trim();
    if (currentNote.indexOf("로봇: 처리 시작") !== -1 ||
        currentNote.indexOf("로봇: 견적서") !== -1) {
      return { success: true, message: "이미 처리 중이거나 완료된 건입니다." };
    }
    sheet.getRange(row, COL_NOTE).setValue("로봇: 처리 시작...");
    SpreadsheetApp.flush();

    var clientName = String(data[COL_NAME - 1]).trim();
    var phone = normalizePhone(data[COL_BOOKER_ID - 1]).replace(/-/g, "");
    var linkUrl = String(data[COL_LINK - 1]).trim();
    var trID = String(data[COL_TR_ID - 1]).trim();
    var tradeDate = data[COL_DATE - 1];

    if (!clientName) throw new Error("B열(예약자명)이 비어있습니다.");
    if (!phone || phone.length < 10) throw new Error("E열(전화번호)이 비어있거나 잘못되었습니다.");
    if (!linkUrl || linkUrl.indexOf("docs.google.com") === -1) throw new Error("M열(계약서 링크)이 비어있습니다.");

    var fileIdMatch = linkUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!fileIdMatch) throw new Error("계약서 링크에서 파일 ID를 추출할 수 없습니다.");

    var pdfLink = convertSheetToPdf(fileIdMatch[1], trID, clientName);

    if (tradeDate instanceof Date) tradeDate = Utilities.formatDate(tradeDate, "Asia/Seoul", "yyyy-MM-dd");
    var dateParts = String(tradeDate).replace(/-/g, "").match(/(\d{4})(\d{2})(\d{2})/);
    var formattedDate = dateParts ? dateParts[1] + "년 " + dateParts[2] + "월 " + dateParts[3] + "일" : String(tradeDate);

    var sent = sendContractAlimtalk(phone, clientName, formattedDate, pdfLink);
    if (sent === "alimtalk") {
      sheet.getRange(row, COL_NOTE).setValue("로봇: 견적서 알림톡 발송 완료! ✅");
      SpreadsheetApp.flush();
      return { success: true, message: clientName + "님에게 견적서 알림톡 발송 완료!" };
    } else if (sent === "sms") {
      sheet.getRange(row, COL_NOTE).setValue("로봇: 견적서 SMS 대체발송 (카톡 미사용자) 📱");
      SpreadsheetApp.flush();
      return { success: true, message: clientName + "님에게 견적서 SMS 대체발송 완료!" };
    } else if (sent === "pending" || sent === true) {
      sheet.getRange(row, COL_NOTE).setValue("로봇: 견적서 발송 접수 완료 ⏳");
      SpreadsheetApp.flush();
      return { success: true, message: clientName + "님에게 견적서 발송 완료!" };
    } else {
      sheet.getRange(row, COL_NOTE).setValue("실패: 알림톡 발송 오류");
      SpreadsheetApp.flush();
      throw new Error("알림톡 발송 실패. Apps Script 로그를 확인하세요.");
    }
  } finally { lock.releaseLock(); }
}

// ─── 구글시트 → PDF 변환 → 드라이브 저장 (1페이지 맞춤) ─────
function convertSheetToPdf(spreadsheetId, trID, clientName) {
  var url = "https://docs.google.com/spreadsheets/d/" + spreadsheetId +
    "/export?format=pdf" +
    "&size=A4" +
    "&portrait=true" +
    "&scale=4" +          // 1페이지에 맞춤 (1=100%, 2=너비맞춤, 3=높이맞춤, 4=페이지맞춤)
    "&fitw=true" +
    "&top_margin=0.50" +
    "&bottom_margin=0.50" +
    "&left_margin=0.50" +
    "&right_margin=0.50" +
    "&sheetnames=false" +
    "&printtitle=false" +
    "&pagenumbers=false" +
    "&gridlines=false" +
    "&fzr=false";

  var blob = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  }).getBlob().setName(trID + "_" + clientName + "_견적서.pdf");

  var pdfFile = DriveApp.getFolderById(FOLDER_ID).createFile(blob);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return pdfFile.getUrl();
}

// ─── 견적서 알림톡 발송 ───────────────────────────────────────
function sendContractAlimtalk(phone, customerName, tradeDate, pdfLink) {
  try {
    var accessToken = getPopbillAccessToken(["member", "153"]);
    var content = "안녕하세요, " + customerName + " 감독님!\n\n빌리지 렌탈샵입니다. \n\n요청하신 견적서를 보내드립니다. \n\n아래 버튼을 눌러 확인해주세요 😊 \n\n감사합니다!\n\n(해당 견적 관련 메시지는 고객님의 알림 신청에 의해 발송되었습니다.)";

    var res = UrlFetchApp.fetch("https://popbill.linkhub.co.kr/ATS", {
      method: "POST",
      headers: { Authorization: "Bearer " + accessToken, "Content-Type": "application/json" },
      payload: JSON.stringify({
        templateCode: CONTRACT_TEMPLATE_CODE, snd: "01071261139", altSendType: "C",
        msgs: [{
          rcv: phone, rcvnm: customerName, msg: content,
          altmsg: "[빌리지] " + customerName + " 감독님, 견적서입니다.\n확인: " + pdfLink,
          btns: [{ n: "견적서 확인하기", t: "WL", u1: pdfLink, u2: pdfLink }]
        }]
      }),
      muteHttpExceptions: true
    });

    var resCode = res.getResponseCode();
    var resText = res.getContentText();
    Logger.log("견적서 알림톡 응답 [" + resCode + "]: " + resText);

    if (resCode !== 200) return false;

    var result = JSON.parse(resText);
    if (!result.receiptNum) return false;

    // 5초 후 전송결과 확인
    try {
      Utilities.sleep(5000);
      var delivery = checkAlimtalkDelivery(result.receiptNum, accessToken);
      if (delivery && delivery.msgs && delivery.msgs.length > 0) {
        var msg = delivery.msgs[0];
        if (msg.state == 3 && msg.result == 100) return "alimtalk";
        if (msg.state == 3) return "sms";
      }
    } catch (checkErr) { Logger.log("전송결과 확인 오류: " + checkErr.message); }

    return "pending";
  } catch (err) {
    Logger.log("sendContractAlimtalk 오류: " + err.message);
    return false;
  }
}

// ─── 알림톡 전송결과 조회 ─────────────────────────────────────
function checkAlimtalkDelivery(receiptNum, accessToken) {
  if (!accessToken) accessToken = getPopbillAccessToken(["member", "153"]);
  var res = UrlFetchApp.fetch("https://popbill.linkhub.co.kr/KakaoTalk/" + receiptNum, {
    method: "GET",
    headers: { Authorization: "Bearer " + accessToken, "Content-Type": "application/json" },
    muteHttpExceptions: true
  });
  return JSON.parse(res.getContentText());
}

// ─── 알림톡 발송 테스트 (GAS 편집기에서 직접 실행용) ─────────
// 실행 후 Apps Script 로그(Ctrl+Enter)에서 결과 확인
function testContractAlimtalk() {
  var TEST_ROW = 2; // 테스트할 행 번호로 변경
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_NAME);
  var data = sheet.getRange(TEST_ROW, 1, 1, COL_MGTKEY).getValues()[0];
  var phone = normalizePhone(data[COL_BOOKER_ID - 1]).replace(/-/g, "");
  var clientName = String(data[COL_NAME - 1]).trim();
  var tradeDate = data[COL_DATE - 1];
  if (tradeDate instanceof Date) tradeDate = Utilities.formatDate(tradeDate, "Asia/Seoul", "yyyy-MM-dd");
  var dateParts = String(tradeDate).replace(/-/g, "").match(/(\d{4})(\d{2})(\d{2})/);
  var formattedDate = dateParts ? dateParts[1] + "년 " + dateParts[2] + "월 " + dateParts[3] + "일" : String(tradeDate);

  Logger.log("발송 대상: " + clientName + " / " + phone + " / " + formattedDate);
  var result = sendContractAlimtalk(phone, clientName, formattedDate, "https://drive.google.com/file/d/test");
  Logger.log("결과: " + result);
}

// ─── AppSheet 요청 처리 (1분 트리거) ─────────────────────────
// 장비체크 시트 E열에 "견적서발송요청" 또는 "발행요청" 감지 후 처리
function processEquipCheckRequests() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var checkSheet = ss.getSheetByName(CHECK_SHEET_NAME);
  if (!checkSheet) return;

  var checkData = checkSheet.getDataRange().getValues();
  var tradeSheet = ss.getSheetByName(SH_NAME);

  var COL_CHECK_TRID = 0;
  var COL_CHECK_ACTION = 4;

  for (var i = 1; i < checkData.length; i++) {
    var action = String(checkData[i][COL_CHECK_ACTION]).trim();
    var trID = String(checkData[i][COL_CHECK_TRID]).trim();
    if (!trID || trID === "" || trID === "undefined") continue;

    if (action === "견적서발송요청") {
      var found = tradeSheet.getRange(1, COL_TR_ID, tradeSheet.getLastRow(), 1).createTextFinder(trID).matchEntireCell(true).findNext();
      var tradeRow = found ? found.getRow() : -1;
      if (tradeRow !== -1) {
        try {
          var result = executeSendContract(tradeRow);
          checkSheet.getRange(i + 1, COL_CHECK_ACTION + 1).setValue("견적서발송완료");
          Logger.log("견적서 발송 완료: " + trID + " - " + (result && result.message ? result.message : "OK"));
        } catch (err) {
          checkSheet.getRange(i + 1, COL_CHECK_ACTION + 1).setValue("발송오류: " + err.message);
          Logger.log("견적서 발송 오류: " + trID + " - " + err.message);
        }
      } else {
        checkSheet.getRange(i + 1, COL_CHECK_ACTION + 1).setValue("거래ID없음");
      }

    } else if (action === "발행요청") {
      var found2 = tradeSheet.getRange(1, COL_TR_ID, tradeSheet.getLastRow(), 1).createTextFinder(trID).matchEntireCell(true).findNext();
      var tradeRow2 = found2 ? found2.getRow() : -1;
      if (tradeRow2 !== -1) {
        try {
          // 이미 발행완료된 건 스킵
          if (tradeSheet.getRange(tradeRow2, COL_ISSUE).getValue() === "발행완료") {
            checkSheet.getRange(i + 1, COL_CHECK_ACTION + 1).setValue("이미발행완료");
            continue;
          }

          // 발행 직전 금액 동기화
          var linkUrl2 = tradeSheet.getRange(tradeRow2, COL_LINK).getValue();
          if (linkUrl2) {
            var matchLink2 = String(linkUrl2).match(/\/d\/([a-zA-Z0-9-_]+)/);
            if (matchLink2) {
              try {
                var latestAmt = SpreadsheetApp.openById(matchLink2[1]).getSheets()[0].getRange(CONTRACT_AMOUNT_CELL).getValue();
                if (typeof latestAmt === "number" && latestAmt > 0) {
                  if (tradeSheet.getRange(tradeRow2, COL_PAY_METHOD).getValue() === "계좌이체(VAT별도)") latestAmt = Math.round(latestAmt / 1.1);
                  tradeSheet.getRange(tradeRow2, COL_AMOUNT).setValue(latestAmt);
                }
              } catch (syncErr) { Logger.log("금액 동기화 실패: " + syncErr.message); }
            }
          }

          // 증빙 발행 실행
          var docType2 = tradeSheet.getRange(tradeRow2, COL_DOC_TYPE).getValue();
          if (docType2 === "세금계산서") {
            requestTaxInvoice(tradeRow2);
            tradeSheet.getRange(tradeRow2, COL_ISSUE).setValue("발행완료");
            tradeSheet.getRange(tradeRow2, COL_NOTE).setValue("로봇: 세금계산서 발행 성공!");
          } else if (docType2 === "현금영수증(전화번호)") {
            requestCashbill(tradeRow2, "전화번호");
            tradeSheet.getRange(tradeRow2, COL_ISSUE).setValue("발행완료");
            tradeSheet.getRange(tradeRow2, COL_NOTE).setValue("로봇: 현금영수증(전화번호) 발행 성공!");
          } else if (docType2 === "현금영수증(사업자번호)") {
            requestCashbill(tradeRow2, "사업자번호");
            tradeSheet.getRange(tradeRow2, COL_ISSUE).setValue("발행완료");
            tradeSheet.getRange(tradeRow2, COL_NOTE).setValue("로봇: 현금영수증(사업자번호) 발행 성공!");
          } else {
            tradeSheet.getRange(tradeRow2, COL_NOTE).setValue("J열 증빙유형을 먼저 선택해주세요.");
            checkSheet.getRange(i + 1, COL_CHECK_ACTION + 1).setValue("증빙유형미선택");
            continue;
          }
          checkSheet.getRange(i + 1, COL_CHECK_ACTION + 1).setValue("발행완료");
        } catch (err2) {
          tradeSheet.getRange(tradeRow2, COL_ISSUE).setValue("전송실패");
          tradeSheet.getRange(tradeRow2, COL_NOTE).setValue("실패: " + err2.message);
          checkSheet.getRange(i + 1, COL_CHECK_ACTION + 1).setValue("발행오류: " + err2.message);
        }
      }
    }
  }
}
