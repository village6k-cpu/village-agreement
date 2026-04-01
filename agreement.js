/**
 * ====================================================================
 * agreement.gs — 약관 동의 API (개고생2.0)
 * ====================================================================
 *
 * 약관 동의 페이지(GitHub Pages)에서 호출하는 API
 * - GET  ?action=info&id=거래ID  → 대여기간 + 계약서 링크 반환
 * - POST ?action=agree           → 동의 기록 저장
 *
 * 거래내역 P열: 약관동의일시
 * 약관동의기록 시트: 상세 로그 (거래ID, 동의일시, IP, UA, 약관버전)
 */

var COL_AGREEMENT = 16;  // P열: 약관동의일시
var AGREEMENT_VERSION = "v1.0";  // 약관 버전 관리

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 웹앱 엔드포인트
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function doGet(e) {
  var params = e ? e.parameter : {};
  var action = params.action || "";

  // CORS preflight 대응
  var output;

  try {
    if (action === "info") {
      output = getAgreementInfo(params.id);
    } else {
      output = { error: "action 파라미터 필요 (info)", available: ["info"] };
    }
  } catch (err) {
    output = { error: err.message };
  }

  return ContentService.createTextOutput(
    (params.callback ? params.callback + "(" : "") +
    JSON.stringify(output) +
    (params.callback ? ")" : "")
  ).setMimeType(
    params.callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON
  );
}

function doPost(e) {
  var output;
  try {
    var body = {};
    if (e && e.postData) {
      body = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      body = e.parameter;
    }

    if (body.action === "agree") {
      output = recordAgreement(body);
    } else {
      output = { error: "action 필요 (agree)" };
    }
  } catch (err) {
    output = { error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 거래 정보 조회
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function getAgreementInfo(tradeID) {
  if (!tradeID) return { error: "id 파라미터 필요" };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("거래내역");
  if (!sheet) return { error: "거래내역 시트 없음" };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { error: "데이터 없음" };

  // 거래ID(D열)로 행 찾기
  var data = sheet.getRange(2, 1, lastRow - 1, COL_AGREEMENT).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][COL_TR_ID - 1]).trim() === String(tradeID).trim()) {
      var date = data[i][COL_DATE - 1];
      var contractLink = data[i][COL_LINK - 1] || "";
      var existingAgreement = data[i][COL_AGREEMENT - 1] || "";
      var customerName = data[i][COL_NAME - 1] || "";

      // 날짜 포맷
      var dateStr = "";
      if (date instanceof Date) {
        dateStr = Utilities.formatDate(date, "Asia/Seoul", "yyyy.MM.dd");
      } else if (date) {
        dateStr = String(date);
      }

      return {
        status: "OK",
        tradeID: tradeID,
        customerName: customerName,
        date: dateStr,
        contractLink: contractLink,
        alreadyAgreed: existingAgreement ? true : false,
        agreedAt: existingAgreement ? String(existingAgreement) : null,
        version: AGREEMENT_VERSION
      };
    }
  }

  return { error: "거래ID를 찾을 수 없습니다: " + tradeID };
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 동의 기록 저장
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function recordAgreement(body) {
  var tradeID = body.id;
  var ip = body.ip || "";
  var userAgent = body.userAgent || "";

  if (!tradeID) return { error: "id 필요" };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("거래내역");
  if (!sheet) return { error: "거래내역 시트 없음" };

  var now = new Date();
  var nowStr = Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");

  // 거래ID(D열)로 행 찾기
  var lastRow = sheet.getLastRow();
  var ids = sheet.getRange(2, COL_TR_ID, lastRow - 1, 1).getValues();
  var targetRow = -1;

  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(tradeID).trim()) {
      targetRow = i + 2;
      break;
    }
  }

  if (targetRow < 0) return { error: "거래ID 없음: " + tradeID };

  // P열 헤더 확인
  var pHeader = sheet.getRange(1, COL_AGREEMENT).getValue();
  if (!pHeader || String(pHeader).trim() === "") {
    sheet.getRange(1, COL_AGREEMENT).setValue("약관동의일시");
    sheet.getRange(1, COL_AGREEMENT).setFontWeight("bold");
  }

  // 이미 동의한 건인지 확인
  var existing = sheet.getRange(targetRow, COL_AGREEMENT).getValue();
  if (existing) {
    return { status: "ALREADY_AGREED", message: "이미 동의 완료", agreedAt: String(existing) };
  }

  // ── 거래내역 P열에 동의일시 기록 ──
  sheet.getRange(targetRow, COL_AGREEMENT).setValue(nowStr);

  // ── 약관동의기록 시트에 상세 로그 ──
  var logSheet = ss.getSheetByName("약관동의기록");
  if (!logSheet) {
    logSheet = ss.insertSheet("약관동의기록");
    logSheet.getRange(1, 1, 1, 6).setValues([["거래ID", "동의일시", "IP주소", "기기정보", "약관버전", "비고"]]);
    logSheet.getRange(1, 1, 1, 6).setFontWeight("bold");
    logSheet.setFrozenRows(1);
    logSheet.setColumnWidths(1, 1, 120);
    logSheet.setColumnWidths(2, 1, 160);
    logSheet.setColumnWidths(3, 1, 130);
    logSheet.setColumnWidths(4, 1, 300);
    logSheet.setColumnWidths(5, 1, 80);
    logSheet.setColumnWidths(6, 1, 120);
  }

  var logRow = logSheet.getLastRow() + 1;
  logSheet.getRange(logRow, 1, 1, 6).setValues([[
    tradeID, nowStr, ip, userAgent, AGREEMENT_VERSION, ""
  ]]);

  return {
    status: "OK",
    message: "동의 기록 완료",
    tradeID: tradeID,
    agreedAt: nowStr
  };
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 약관 링크 생성 헬퍼
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 거래ID로 약관 동의 링크 생성
 * @param {string} tradeID
 * @param {string} agreementPageUrl - GitHub Pages 등 약관 페이지 기본 URL
 * @returns {string} 완성된 약관 동의 링크
 */
var AGREEMENT_PAGE_URL = "https://village6k-cpu.github.io/village-agreement/";

function generateAgreementLink(tradeID) {
  return AGREEMENT_PAGE_URL + "?id=" + encodeURIComponent(tradeID);
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 메뉴: 약관 링크 복사
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function copyAgreementLink() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("거래내역");
  if (!sheet) { ui.alert("❌ 거래내역 시트가 없습니다."); return; }

  var row = sheet.getActiveCell().getRow();
  if (row < 2) { ui.alert("❌ 발송할 행을 선택해주세요."); return; }

  var tradeID = sheet.getRange(row, COL_TR_ID).getValue();
  if (!tradeID) { ui.alert("❌ 선택한 행에 거래ID가 없습니다."); return; }

  var link = generateAgreementLink(tradeID);
  var customerName = sheet.getRange(row, COL_NAME).getValue() || "";

  ui.alert(
    "📋 약관 동의 링크\n\n" +
    "고객: " + customerName + "\n" +
    "거래ID: " + tradeID + "\n\n" +
    link + "\n\n" +
    "↑ 이 링크를 복사해서 카톡으로 보내세요."
  );
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 메뉴: 약관 동의 알림톡 발송
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function sendAgreementAlimtalk() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("거래내역");
  if (!sheet) { ui.alert("❌ 거래내역 시트가 없습니다."); return; }

  var row = sheet.getActiveCell().getRow();
  if (row < 2) { ui.alert("❌ 발송할 행을 선택해주세요."); return; }

  var tradeID = sheet.getRange(row, COL_TR_ID).getValue();
  var customerName = sheet.getRange(row, COL_NAME).getValue() || "";
  var phone = normalizePhone(sheet.getRange(row, COL_BOOKER_ID).getDisplayValue()).replace(/-/g, "");

  if (!tradeID) { ui.alert("❌ 거래ID가 없습니다."); return; }
  if (!phone || phone.length < 10) { ui.alert("❌ 전화번호가 없거나 올바르지 않습니다."); return; }

  // 이미 동의한 건인지 확인
  var existing = sheet.getRange(row, COL_AGREEMENT).getValue();
  if (existing) {
    ui.alert("ℹ️ 이미 동의 완료된 건입니다.\n동의일시: " + existing);
    return;
  }

  var link = generateAgreementLink(tradeID);

  var confirm = ui.alert(
    "📋 약관 동의 알림톡 발송",
    "고객: " + customerName + "\n" +
    "연락처: " + phone + "\n" +
    "거래ID: " + tradeID + "\n\n" +
    "약관 동의 링크를 알림톡으로 발송하시겠습니까?",
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  try {
    // 승인된 알림톡 템플릿과 정확히 일치해야 함
    var content = customerName + " 감독님, 안녕하세요.\n빌리지 렌탈샵입니다.\n\n" +
      "장비 대여에 앞서 약관 동의가 필요합니다.\n" +
      "아래 버튼을 눌러 약관을 확인하고 동의해주세요.\n\n" +
      "감사합니다.";

    var accessToken = getPopbillAccessToken(["member", "153"]);

    var response = UrlFetchApp.fetch("https://popbill.linkhub.co.kr/ATS", {
      method: "POST",
      headers: { Authorization: "Bearer " + accessToken, "Content-Type": "application/json" },
      payload: JSON.stringify({
        templateCode: "026040000008", snd: "01071261139", altSendType: "A",
        msgs: [{
          rcv: phone, rcvnm: customerName, msg: content,
          altmsg: "[빌리지] " + customerName + " 감독님, 약관 동의 요청입니다.\n" + link,
          btns: [{ n: "약관 확인 및 동의", t: "WL", u1: link, u2: link }]
        }]
      }),
      muteHttpExceptions: true
    });

    var result = JSON.parse(response.getContentText());
    Logger.log("약관 알림톡 발송 결과: " + response.getContentText());

    if (response.getResponseCode() === 200) {
      ui.alert("✅ 약관 동의 알림톡 발송 완료!\n\n" + customerName + "님에게 발송되었습니다.");
    } else {
      ui.alert("⚠️ 발송 실패: " + (result.message || response.getContentText()));
    }
  } catch (err) {
    ui.alert("❌ 발송 오류: " + err.message);
  }
}
