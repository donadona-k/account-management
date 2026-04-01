// ============================================================
// 社内アカウント申請・管理システム - Code.gs
// ============================================================

const CONFIG = {
  // シート名
  SHEET_REQUESTS:        "申請一覧",
  SHEET_ACCOUNTS:        "アカウント台帳",
  SHEET_DELETE_REQUESTS: "削除申請一覧",
  SHEET_SERVICES:        "サービスマスタ",

  // 申請ステータス
  STATUS_PENDING:  "申請中",
  STATUS_APPROVED: "承認",
  STATUS_REJECTED: "却下",

  // アカウント台帳のアカウントステータス
  ACCOUNT_STATUS_ACTIVE:  "運用中",
  ACCOUNT_STATUS_DEL_REQ: "削除申請中",
  ACCOUNT_STATUS_DELETED: "削除済み",

  // アカウント台帳の列インデックス（1始まり）
  ACC_COL_ACCOUNT_ID:   1,
  ACC_COL_SERVICE_ID:   2,  // サービスマスタとの紐付け
  ACC_COL_SERVICE_NAME: 3,
  ACC_COL_ACCOUNT_TYPE: 4,
  ACC_COL_USER_NAME:    5,
  ACC_COL_DEPT:         6,
  ACC_COL_EMAIL:        7,
  ACC_COL_START_DATE:   8,
  ACC_COL_PURPOSE:      9,
  ACC_COL_REQUEST_ID:   10,
  ACC_COL_CREATED_AT:   11,
  ACC_COL_STATUS:       12,
  ACC_COL_DEL_REQ_ID:   13,
  ACC_COL_DELETED_AT:   14,

  // 承認者メール（未設定時のフォールバック）
  APPROVER_EMAILS_FALLBACK: "admin@example.com",
};

// ============================================================
// 1. 初期セットアップ（初回のみ手動で実行）
// ============================================================
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _setupRequestSheet(ss);
  _setupAccountSheet(ss);
  _setupDeleteRequestSheet(ss);
  _setupServiceMasterSheet(ss);  // ServiceManager.gs に定義
  _setupOnEditTrigger();
  SpreadsheetApp.getUi().alert(
    "セットアップ完了！\n" +
    "「アカウント管理」メニュー → 管理画面から\n" +
    "サービスの追加とフォームIDの設定を行ってください。"
  );
}

function _setupRequestSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_REQUESTS);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_REQUESTS);

  const headers = [
    "申請ID", "タイムスタンプ", "申請者名", "部署", "申請者メール",
    "サービス名", "アカウント種別", "用途・理由", "利用開始希望日",
    "ステータス", "承認者コメント", "承認/却下日時", "台帳転記日時"
  ];
  _writeHeader(sheet, headers, "#1a73e8");

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([CONFIG.STATUS_PENDING, CONFIG.STATUS_APPROVED, CONFIG.STATUS_REJECTED])
    .setAllowInvalid(false).build();
  sheet.getRange(2, headers.indexOf("ステータス") + 1, 1000).setDataValidation(rule);

  sheet.setColumnWidth(1, 90); sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(8, 250); sheet.setColumnWidth(10, 80);
}

function _setupAccountSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_ACCOUNTS);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_ACCOUNTS);

  const headers = [
    "アカウントID", "サービスID", "サービス名", "アカウント種別",
    "利用者名", "部署", "メールアドレス", "利用開始日",
    "用途", "申請ID", "登録日時", "ステータス", "削除申請ID", "削除日時"
  ];
  _writeHeader(sheet, headers, "#0f9d58");

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([
      CONFIG.ACCOUNT_STATUS_ACTIVE,
      CONFIG.ACCOUNT_STATUS_DEL_REQ,
      CONFIG.ACCOUNT_STATUS_DELETED,
    ])
    .setAllowInvalid(false).build();
  sheet.getRange(2, CONFIG.ACC_COL_STATUS, 1000).setDataValidation(rule);

  sheet.setColumnWidth(1, 110); sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(9, 250); sheet.setColumnWidth(12, 100);
}

function _setupDeleteRequestSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_DELETE_REQUESTS);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_DELETE_REQUESTS);

  const headers = [
    "削除申請ID", "タイムスタンプ", "申請者名", "部署", "申請者メール",
    "アカウントID", "サービス名", "削除理由",
    "ステータス", "承認者コメント", "承認/却下日時"
  ];
  _writeHeader(sheet, headers, "#d93025");

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([CONFIG.STATUS_PENDING, CONFIG.STATUS_APPROVED, CONFIG.STATUS_REJECTED])
    .setAllowInvalid(false).build();
  sheet.getRange(2, headers.indexOf("ステータス") + 1, 1000).setDataValidation(rule);

  sheet.setColumnWidth(1, 120); sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(6, 120); sheet.setColumnWidth(8, 250); sheet.setColumnWidth(9, 80);
}

function _setupOnEditTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "onStatusChange") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("onStatusChange")
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
}

// ============================================================
// 2. 新規アカウント申請フォーム送信トリガー
// ============================================================
function onFormSubmit(e) {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const requestSheet = ss.getSheetByName(CONFIG.SHEET_REQUESTS);
  const responses    = e.namedValues;
  const requestId    = _generateId("REQ");

  const applicantName  = _val(responses["申請者名"]);
  const department     = _val(responses["部署"]);
  const applicantEmail = _val(responses["申請者メール（社内）"]);
  const serviceName    = _val(responses["サービス名"]);
  const accountType    = _val(responses["アカウント種別"]);
  const reason         = _val(responses["用途・理由"]);
  const startDate      = _val(responses["利用開始希望日"]);

  requestSheet.appendRow([
    requestId, new Date(), applicantName, department, applicantEmail,
    serviceName, accountType, reason, startDate,
    CONFIG.STATUS_PENDING, "", "", ""
  ]);

  _sendApproverNotification(requestId, applicantName, department, applicantEmail,
                             serviceName, accountType, reason, startDate);
  if (applicantEmail) {
    _sendApplicantAcknowledgement(applicantEmail, applicantName, requestId, serviceName);
  }
}

// ============================================================
// 3. 削除申請フォーム送信トリガー
//    ※ 削除フォーム専用のトリガーとして別途登録すること
// ============================================================
function onDeleteFormSubmit(e) {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const deleteSheet  = ss.getSheetByName(CONFIG.SHEET_DELETE_REQUESTS);
  const accountSheet = ss.getSheetByName(CONFIG.SHEET_ACCOUNTS);
  const responses    = e.namedValues;
  const deleteReqId  = _generateId("DEL");

  const applicantName  = _val(responses["申請者名"]);
  const department     = _val(responses["部署"]);
  const applicantEmail = _val(responses["申請者メール（社内）"]);
  const accountId      = _val(responses["アカウントID"]).trim();
  const deleteReason   = _val(responses["削除理由"]);

  const accountRow = _findAccountRow(accountSheet, accountId);
  if (!accountRow) {
    if (applicantEmail) {
      MailApp.sendEmail(
        applicantEmail,
        `【削除申請エラー】アカウントID "${accountId}" が見つかりません`,
        `${applicantName} 様\n\n指定されたアカウントID「${accountId}」は台帳に存在しないか、すでに削除済みです。\nアカウントIDを確認の上、再度申請してください。`
      );
    }
    return;
  }

  const serviceName = accountSheet.getRange(accountRow, CONFIG.ACC_COL_SERVICE_NAME).getValue();

  deleteSheet.appendRow([
    deleteReqId, new Date(), applicantName, department, applicantEmail,
    accountId, serviceName, deleteReason,
    CONFIG.STATUS_PENDING, "", ""
  ]);

  _markAccountAsPendingDeletion(accountSheet, accountRow, deleteReqId);

  _sendDeleteApproverNotification(deleteReqId, applicantName, department,
                                   applicantEmail, accountId, serviceName, deleteReason);
  if (applicantEmail) {
    MailApp.sendEmail(
      applicantEmail,
      `【削除申請受付】${serviceName} のアカウント削除申請を受け付けました`,
      `${applicantName} 様\n\n${serviceName}（アカウントID: ${accountId}）の削除申請を受け付けました。\n担当者が確認次第、結果をご連絡します。\n\n削除申請ID: ${deleteReqId}`
    );
  }
}

// ============================================================
// 4. ステータス変更トリガー（新規申請・削除申請 共通）
// ============================================================
function onStatusChange(e) {
  const sheet     = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const range     = e.range;
  const row       = range.getRow();
  if (row < 2) return;

  if (sheetName === CONFIG.SHEET_REQUESTS && range.getColumn() === 10) {
    _handleNewRequestStatusChange(sheet, row, range.getValue());
  }
  if (sheetName === CONFIG.SHEET_DELETE_REQUESTS && range.getColumn() === 9) {
    _handleDeleteRequestStatusChange(sheet, row, range.getValue());
  }
}

function _handleNewRequestStatusChange(sheet, row, newStatus) {
  const rowData = sheet.getRange(row, 1, 1, 13).getValues()[0];
  const [requestId, , applicantName, department, applicantEmail,
         serviceName, accountType, reason, startDate, , comment, , transferredAt] = rowData;

  if (newStatus === CONFIG.STATUS_APPROVED && !transferredAt) {
    // サービスマスタからサービスIDを取得して台帳に転記
    const svc = getServiceByName(serviceName);  // ServiceManager.gs
    _transferToAccountLedger(
      requestId, svc ? svc.id : "", serviceName,
      accountType, applicantName, department,
      applicantEmail, reason, startDate
    );
    sheet.getRange(row, 13).setValue(new Date());
    sheet.getRange(row, 12).setValue(new Date());
    if (applicantEmail) _sendApplicantResult(applicantEmail, applicantName, serviceName, true, comment);
  } else if (newStatus === CONFIG.STATUS_REJECTED) {
    sheet.getRange(row, 12).setValue(new Date());
    if (applicantEmail) _sendApplicantResult(applicantEmail, applicantName, serviceName, false, comment);
  }
}

function _handleDeleteRequestStatusChange(sheet, row, newStatus) {
  const rowData = sheet.getRange(row, 1, 1, 11).getValues()[0];
  const [deleteReqId, , applicantName, , applicantEmail,
         accountId, serviceName, , , comment] = rowData;

  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const accountSheet = ss.getSheetByName(CONFIG.SHEET_ACCOUNTS);
  const accountRow   = _findAccountRow(accountSheet, accountId);

  if (newStatus === CONFIG.STATUS_APPROVED) {
    sheet.getRange(row, 11).setValue(new Date());
    if (accountRow) _markAccountAsDeleted(accountSheet, accountRow);
    if (applicantEmail) {
      MailApp.sendEmail(applicantEmail,
        `【削除申請承認】${serviceName} のアカウントが削除されました`,
        `${applicantName} 様\n\n${serviceName}（アカウントID: ${accountId}）の削除申請が承認されました。\n${comment ? `コメント: ${comment}` : ""}`
      );
    }
  } else if (newStatus === CONFIG.STATUS_REJECTED) {
    sheet.getRange(row, 11).setValue(new Date());
    if (accountRow) _revertAccountStatus(accountSheet, accountRow);
    if (applicantEmail) {
      MailApp.sendEmail(applicantEmail,
        `【削除申請却下】${serviceName} の削除申請が却下されました`,
        `${applicantName} 様\n\n${serviceName}（アカウントID: ${accountId}）の削除申請が却下されました。\n${comment ? `コメント: ${comment}` : ""}`
      );
    }
  }
}

// ============================================================
// 5. アカウント台帳の操作
// ============================================================
function _transferToAccountLedger(requestId, serviceId, serviceName,
                                   accountType, userName, dept,
                                   email, reason, startDate) {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const accountSheet = ss.getSheetByName(CONFIG.SHEET_ACCOUNTS);
  const accountId    = _generateId("ACC");

  accountSheet.appendRow([
    accountId, serviceId, serviceName, accountType,
    userName, dept, email, startDate,
    reason, requestId, new Date(),
    CONFIG.ACCOUNT_STATUS_ACTIVE, "", ""
  ]);
}

function _findAccountRow(accountSheet, accountId) {
  const data = accountSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === accountId) return i + 1;
  }
  return null;
}

function _markAccountAsPendingDeletion(accountSheet, row, deleteReqId) {
  accountSheet.getRange(row, CONFIG.ACC_COL_STATUS).setValue(CONFIG.ACCOUNT_STATUS_DEL_REQ);
  accountSheet.getRange(row, CONFIG.ACC_COL_DEL_REQ_ID).setValue(deleteReqId);
  accountSheet.getRange(row, 1, 1, 14).setBackground("#fce8e6").setFontColor("#c5221f");
  accountSheet.getRange(row, CONFIG.ACC_COL_STATUS)
    .setBackground("#f28b82").setFontWeight("bold");
}

function _markAccountAsDeleted(accountSheet, row) {
  accountSheet.getRange(row, CONFIG.ACC_COL_STATUS).setValue(CONFIG.ACCOUNT_STATUS_DELETED);
  accountSheet.getRange(row, CONFIG.ACC_COL_DELETED_AT).setValue(new Date());
  const range = accountSheet.getRange(row, 1, 1, 14);
  range.setBackground("#f1f3f4").setFontColor("#80868b");
  accountSheet.getRange(row, 1, 1, CONFIG.ACC_COL_STATUS - 1)
    .setTextStyle(SpreadsheetApp.newTextStyle().setStrikethrough(true).build());
  accountSheet.getRange(row, CONFIG.ACC_COL_STATUS)
    .setBackground("#e8eaed").setFontWeight("bold")
    .setTextStyle(SpreadsheetApp.newTextStyle().setStrikethrough(false).build());
}

function _revertAccountStatus(accountSheet, row) {
  accountSheet.getRange(row, CONFIG.ACC_COL_STATUS).setValue(CONFIG.ACCOUNT_STATUS_ACTIVE);
  accountSheet.getRange(row, CONFIG.ACC_COL_DEL_REQ_ID).setValue("");
  const range = accountSheet.getRange(row, 1, 1, 14);
  range.setBackground(null).setFontColor(null);
  range.setTextStyle(SpreadsheetApp.newTextStyle().setStrikethrough(false).build());
  accountSheet.getRange(row, CONFIG.ACC_COL_STATUS)
    .setBackground("#e6f4ea").setFontColor("#137333").setFontWeight("bold");
}

// ============================================================
// 6. メール送信
// ============================================================
function _sendApproverNotification(requestId, name, dept, email,
                                    service, type, reason, startDate) {
  const subject = `【アカウント申請】${service} の申請が届きました（${requestId}）`;
  const body = `アカウント申請が届きました。スプレッドシートを確認して承認/却下してください。

■ 申請ID　　　: ${requestId}
■ 申請者　　　: ${name}（${dept}）
■ 申請者メール: ${email}
■ サービス名　: ${service}
■ 種別　　　　: ${type}
■ 用途・理由　: ${reason}
■ 利用開始希望: ${startDate}

▼ スプレッドシートを開く
${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;
  MailApp.sendEmail(_getApproverEmails(), subject, body);
}

function _sendDeleteApproverNotification(deleteReqId, name, dept, email,
                                          accountId, service, reason) {
  const subject = `【アカウント削除申請】${service} の削除申請が届きました（${deleteReqId}）`;
  const body = `アカウント削除申請が届きました。スプレッドシートを確認して承認/却下してください。

■ 削除申請ID　: ${deleteReqId}
■ 申請者　　　: ${name}（${dept}）
■ 申請者メール: ${email}
■ アカウントID: ${accountId}
■ サービス名　: ${service}
■ 削除理由　　: ${reason}

▼ スプレッドシートを開く（削除申請一覧シートを確認）
${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;
  MailApp.sendEmail(_getApproverEmails(), subject, body);
}

function _sendApplicantAcknowledgement(email, name, requestId, service) {
  MailApp.sendEmail(email,
    `【アカウント申請受付】${service} の申請を受け付けました`,
    `${name} 様\n\n${service} のアカウント申請を受け付けました。\n担当者が確認次第、結果をご連絡します。\n\n申請ID: ${requestId}`
  );
}

function _sendApplicantResult(email, name, service, approved, comment) {
  const result = approved ? "承認" : "却下";
  MailApp.sendEmail(email,
    `【アカウント申請結果】${service} の申請が${result}されました`,
    `${name} 様\n\n${service} のアカウント申請が${result}されました。\n${comment ? `コメント: ${comment}\n` : ""}${approved ? "アカウントの発行手続きを進めます。" : "ご不明な点は担当者までお問い合わせください。"}`
  );
}

// ============================================================
// 7. ユーティリティ
// ============================================================
function _getApproverEmails() {
  return PropertiesService.getScriptProperties().getProperty("APPROVER_EMAILS")
    || CONFIG.APPROVER_EMAILS_FALLBACK;
}

function _generateId(prefix) {
  const yymmdd = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyMMdd");
  const rand   = Math.random().toString(36).substring(2, 6).toUpperCase();
  return `${prefix}-${yymmdd}-${rand}`;
}

function _val(arr) {
  return arr && arr.length > 0 ? arr[0] : "";
}

function _writeHeader(sheet, headers, bgColor) {
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground(bgColor).setFontColor("#ffffff").setFontWeight("bold");
  sheet.setFrozenRows(1);
}
