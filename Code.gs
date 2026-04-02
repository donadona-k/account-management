// ============================================================
// 社内アカウント申請・管理システム - Code.gs
// ============================================================

const CONFIG = {
  // シート名
  SHEET_REQUESTS:        "申請一覧",
  SHEET_USERS:           "ユーザー台帳",
  SHEET_SERVICE_USAGE:   "サービス利用台帳",
  SHEET_DELETE_REQUESTS: "削除申請一覧",
  SHEET_SERVICES:        "サービスマスタ",

  // 申請ステータス
  STATUS_PENDING:  "申請中",
  STATUS_APPROVED: "承認",
  STATUS_REJECTED: "却下",

  // ユーザーステータス
  USER_STATUS_ACTIVE:  "有効",
  USER_STATUS_DELETED: "削除済み",

  // サービス利用ステータス
  ACCOUNT_STATUS_ACTIVE:  "運用中",
  ACCOUNT_STATUS_DEL_REQ: "削除申請中",
  ACCOUNT_STATUS_DELETED: "削除済み",

  // ユーザー台帳の列インデックス（1始まり）
  USR_COL_ID:         1,
  USR_COL_NAME:       2,
  USR_COL_DEPT:       3,
  USR_COL_EMAIL:      4,
  USR_COL_CREATED_AT: 5,
  USR_COL_STATUS:     6,

  // サービス利用台帳の列インデックス
  USG_COL_USAGE_ID:     1,
  USG_COL_USER_ID:      2,
  USG_COL_USER_NAME:    3,
  USG_COL_SERVICE_ID:   4,
  USG_COL_SERVICE_NAME: 5,
  USG_COL_ACCOUNT_TYPE: 6,
  USG_COL_START_DATE:   7,
  USG_COL_PURPOSE:      8,
  USG_COL_REQUEST_ID:   9,
  USG_COL_CREATED_AT:   10,
  USG_COL_STATUS:       11,
  USG_COL_DEL_REQ_ID:   12,
  USG_COL_DELETED_AT:   13,

  APPROVER_EMAILS_FALLBACK: "admin@example.com",
};

// ============================================================
// 1. 初期セットアップ（初回のみ手動で実行）
// ============================================================
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _setupRequestSheet(ss);
  _setupUserSheet(ss);
  _setupServiceUsageSheet(ss);
  _setupDeleteRequestSheet(ss);
  _setupServiceMasterSheet(ss);  // ServiceManager.gs に定義
  _setupOnEditTrigger();
  console.log("セットアップ完了！");
}

function _setupRequestSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_REQUESTS);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_REQUESTS);

  // 申請者と対象者を分離した新ヘッダー
  const headers = [
    "申請ID", "タイムスタンプ",
    "申請者名", "申請者メール",
    "対象者名", "対象者部署", "対象者メール",
    "サービス名", "アカウント種別", "用途・理由", "利用開始希望日",
    "ステータス", "承認者コメント", "承認/却下日時", "転記日時"
  ];
  _writeHeader(sheet, headers, "#1a73e8");

  const statusCol = headers.indexOf("ステータス") + 1; // 12列目
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([CONFIG.STATUS_PENDING, CONFIG.STATUS_APPROVED, CONFIG.STATUS_REJECTED])
    .setAllowInvalid(false).build();
  sheet.getRange(2, statusCol, 1000).setDataValidation(rule);

  sheet.setColumnWidth(1, 120); sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(10, 250); sheet.setColumnWidth(12, 80);
}

function _setupUserSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_USERS);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_USERS);

  const headers = ["ユーザーID", "氏名", "部署", "メールアドレス", "登録日時", "ステータス"];
  _writeHeader(sheet, headers, "#1565c0");

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([CONFIG.USER_STATUS_ACTIVE, CONFIG.USER_STATUS_DELETED])
    .setAllowInvalid(false).build();
  sheet.getRange(2, CONFIG.USR_COL_STATUS, 1000).setDataValidation(rule);

  sheet.setColumnWidth(1, 130); sheet.setColumnWidth(4, 200);
}

function _setupServiceUsageSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_SERVICE_USAGE);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_SERVICE_USAGE);

  const headers = [
    "利用ID", "ユーザーID", "氏名",
    "サービスID", "サービス名", "アカウント種別",
    "利用開始日", "用途", "申請ID", "登録日時",
    "ステータス", "削除申請ID", "削除日時"
  ];
  _writeHeader(sheet, headers, "#0f9d58");

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([
      CONFIG.ACCOUNT_STATUS_ACTIVE,
      CONFIG.ACCOUNT_STATUS_DEL_REQ,
      CONFIG.ACCOUNT_STATUS_DELETED,
    ])
    .setAllowInvalid(false).build();
  sheet.getRange(2, CONFIG.USG_COL_STATUS, 1000).setDataValidation(rule);

  sheet.setColumnWidth(1, 130); sheet.setColumnWidth(2, 130);
  sheet.setColumnWidth(8, 250); sheet.setColumnWidth(11, 100);
}

function _setupDeleteRequestSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_DELETE_REQUESTS);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_DELETE_REQUESTS);

  const headers = [
    "削除申請ID", "タイムスタンプ",
    "申請者名", "申請者メール",
    "対象者名", "対象者メール", "ユーザーID",
    "削除理由", "ステータス", "承認者コメント", "承認/却下日時"
  ];
  _writeHeader(sheet, headers, "#d93025");

  const statusCol = headers.indexOf("ステータス") + 1; // 9列目
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([CONFIG.STATUS_PENDING, CONFIG.STATUS_APPROVED, CONFIG.STATUS_REJECTED])
    .setAllowInvalid(false).build();
  sheet.getRange(2, statusCol, 1000).setDataValidation(rule);

  sheet.setColumnWidth(1, 130); sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(8, 250); sheet.setColumnWidth(9, 80);
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
  const responses    = _parseFormResponse(e);
  const requestId    = _generateId("REQ");

  const applicantName  = responses["申請者名"]           || "";
  const applicantEmail = responses["申請者メール（社内）"]   || "";
  const targetName     = responses["対象者名"]            || "";
  const targetDept     = responses["対象者部署"]           || "";
  const targetEmail    = responses["対象者メール（社内）"]   || "";
  const serviceName    = responses["サービス名"]           || "";
  const accountType    = responses["アカウント種別"]        || "";
  const reason         = responses["用途・理由"]           || "";
  const startDate      = responses["利用開始希望日"]        || "";

  requestSheet.appendRow([
    requestId, new Date(),
    applicantName, applicantEmail,
    targetName, targetDept, targetEmail,
    serviceName, accountType, reason, startDate,
    CONFIG.STATUS_PENDING, "", "", ""
  ]);

  _sendApproverNotification(requestId, applicantName, applicantEmail,
                             targetName, targetEmail,
                             serviceName, accountType, reason, startDate);
  if (applicantEmail) {
    _sendApplicantAcknowledgement(applicantEmail, applicantName, requestId, targetName, serviceName);
  }
}

// ============================================================
// 3. 削除申請フォーム送信トリガー
//    ※ 削除フォーム専用のトリガーとして別途登録すること
// ============================================================
function onDeleteFormSubmit(e) {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const deleteSheet = ss.getSheetByName(CONFIG.SHEET_DELETE_REQUESTS);
  const userSheet   = ss.getSheetByName(CONFIG.SHEET_USERS);
  const responses   = _parseFormResponse(e);
  const deleteReqId = _generateId("DEL");

  const applicantName  = responses["申請者名"]           || "";
  const applicantEmail = responses["申請者メール（社内）"]   || "";
  const targetEmail    = (responses["対象者メールアドレス"] || "").trim();
  const deleteReason   = responses["削除理由"]           || "";

  // ユーザー台帳から対象者を検索
  const userRow = _findUserRowByEmail(userSheet, targetEmail);
  if (!userRow) {
    if (applicantEmail) {
      MailApp.sendEmail(
        applicantEmail,
        `【削除申請エラー】メールアドレス "${targetEmail}" のユーザーが見つかりません`,
        `${applicantName} 様\n\n指定されたメールアドレス「${targetEmail}」のユーザーはユーザー台帳に存在しないか、すでに削除済みです。`
      );
    }
    return;
  }

  const userData   = userSheet.getRange(userRow, 1, 1, 6).getValues()[0];
  const userId     = String(userData[CONFIG.USR_COL_ID - 1]);
  const targetName = String(userData[CONFIG.USR_COL_NAME - 1]);

  deleteSheet.appendRow([
    deleteReqId, new Date(),
    applicantName, applicantEmail,
    targetName, targetEmail, userId,
    deleteReason, CONFIG.STATUS_PENDING, "", ""
  ]);

  // 紐づく全サービスを「削除申請中」に更新
  _markAllUserServicesAsPendingDeletion(userId, deleteReqId);

  _sendDeleteApproverNotification(deleteReqId, applicantName, applicantEmail,
                                   targetName, targetEmail, userId, deleteReason);
  if (applicantEmail) {
    MailApp.sendEmail(
      applicantEmail,
      `【削除申請受付】${targetName} のアカウント削除申請を受け付けました`,
      `${applicantName} 様\n\n${targetName}（${targetEmail}）のアカウント削除申請を受け付けました。\n紐づくすべてのサービスが削除申請中となりました。\n\n削除申請ID: ${deleteReqId}`
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

  // 申請一覧のステータス列 = 12列目
  if (sheetName === CONFIG.SHEET_REQUESTS && range.getColumn() === 12) {
    _handleNewRequestStatusChange(sheet, row, range.getValue());
  }
  // 削除申請一覧のステータス列 = 9列目
  if (sheetName === CONFIG.SHEET_DELETE_REQUESTS && range.getColumn() === 9) {
    _handleDeleteRequestStatusChange(sheet, row, range.getValue());
  }
}

function _handleNewRequestStatusChange(sheet, row, newStatus) {
  const rowData = sheet.getRange(row, 1, 1, 15).getValues()[0];
  const [requestId, , applicantName, applicantEmail,
         targetName, targetDept, targetEmail,
         serviceName, accountType, reason, startDate,
         , comment, , transferredAt] = rowData;

  if (newStatus === CONFIG.STATUS_APPROVED && !transferredAt) {
    const svc = getServiceByName(serviceName);  // ServiceManager.gs
    _registerServiceUsage(
      requestId, svc ? svc.id : "", serviceName,
      accountType, targetName, targetDept, targetEmail, reason, startDate
    );
    sheet.getRange(row, 15).setValue(new Date()); // 転記日時
    sheet.getRange(row, 14).setValue(new Date()); // 承認日時
    if (applicantEmail) {
      _sendApplicantResult(applicantEmail, applicantName, targetName, serviceName, true, comment);
    }
  } else if (newStatus === CONFIG.STATUS_REJECTED) {
    sheet.getRange(row, 14).setValue(new Date());
    if (applicantEmail) {
      _sendApplicantResult(applicantEmail, applicantName, targetName, serviceName, false, comment);
    }
  }
}

function _handleDeleteRequestStatusChange(sheet, row, newStatus) {
  const rowData = sheet.getRange(row, 1, 1, 11).getValues()[0];
  const [deleteReqId, , applicantName, applicantEmail,
         targetName, targetEmail, userId,
         , , comment] = rowData;

  if (newStatus === CONFIG.STATUS_APPROVED) {
    sheet.getRange(row, 11).setValue(new Date());
    _markAllUserServicesAsDeleted(userId);
    _markUserAsDeleted(userId);
    if (applicantEmail) {
      MailApp.sendEmail(applicantEmail,
        `【削除申請承認】${targetName} のアカウントが削除されました`,
        `${applicantName} 様\n\n${targetName}（${targetEmail}）のアカウント削除申請が承認されました。\n紐づくすべてのサービスが削除済みとなりました。\n${comment ? `コメント: ${comment}` : ""}`
      );
    }
  } else if (newStatus === CONFIG.STATUS_REJECTED) {
    sheet.getRange(row, 11).setValue(new Date());
    _revertAllUserServicesStatus(userId);
    if (applicantEmail) {
      MailApp.sendEmail(applicantEmail,
        `【削除申請却下】${targetName} の削除申請が却下されました`,
        `${applicantName} 様\n\n${targetName}（${targetEmail}）のアカウント削除申請が却下されました。\n${comment ? `コメント: ${comment}` : ""}`
      );
    }
  }
}

// ============================================================
// 5. ユーザー台帳 / サービス利用台帳の操作
// ============================================================

// 承認時：ユーザーを登録（既存なら再利用）してサービス利用を追記
function _registerServiceUsage(requestId, serviceId, serviceName,
                                accountType, targetName, targetDept,
                                targetEmail, reason, startDate) {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const usageSheet = ss.getSheetByName(CONFIG.SHEET_SERVICE_USAGE);
  const userId     = _findOrCreateUser(targetName, targetDept, targetEmail);
  const usageId    = _generateId("USG");

  usageSheet.appendRow([
    usageId, userId, targetName,
    serviceId, serviceName, accountType,
    startDate, reason, requestId, new Date(),
    CONFIG.ACCOUNT_STATUS_ACTIVE, "", ""
  ]);
}

// メールアドレスでユーザーを検索。未登録なら新規追加してIDを返す
function _findOrCreateUser(name, dept, email) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName(CONFIG.SHEET_USERS);
  const data      = userSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][CONFIG.USR_COL_EMAIL - 1]).trim() === email.trim()) {
      return String(data[i][CONFIG.USR_COL_ID - 1]);
    }
  }

  const userId = _generateId("USR");
  userSheet.appendRow([userId, name, dept, email, new Date(), CONFIG.USER_STATUS_ACTIVE]);
  return userId;
}

function _findUserRowByEmail(userSheet, email) {
  const data = userSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][CONFIG.USR_COL_EMAIL - 1]).trim() === email.trim()) {
      return i + 1;
    }
  }
  return null;
}

// ユーザーの全運用中サービスを「削除申請中」に（赤背景）
function _markAllUserServicesAsPendingDeletion(userId, deleteReqId) {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const usageSheet = ss.getSheetByName(CONFIG.SHEET_SERVICE_USAGE);
  const data       = usageSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][CONFIG.USG_COL_USER_ID - 1]) === userId &&
        data[i][CONFIG.USG_COL_STATUS - 1] === CONFIG.ACCOUNT_STATUS_ACTIVE) {
      const row = i + 1;
      usageSheet.getRange(row, CONFIG.USG_COL_STATUS).setValue(CONFIG.ACCOUNT_STATUS_DEL_REQ);
      usageSheet.getRange(row, CONFIG.USG_COL_DEL_REQ_ID).setValue(deleteReqId);
      usageSheet.getRange(row, 1, 1, 13).setBackground("#fce8e6").setFontColor("#c5221f");
      usageSheet.getRange(row, CONFIG.USG_COL_STATUS)
        .setBackground("#f28b82").setFontWeight("bold");
    }
  }
}

// ユーザーの全サービスを「削除済み」に（グレー＋取り消し線）
function _markAllUserServicesAsDeleted(userId) {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const usageSheet = ss.getSheetByName(CONFIG.SHEET_SERVICE_USAGE);
  const data       = usageSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][CONFIG.USG_COL_USER_ID - 1]) === userId &&
        data[i][CONFIG.USG_COL_STATUS - 1] !== CONFIG.ACCOUNT_STATUS_DELETED) {
      const row = i + 1;
      usageSheet.getRange(row, CONFIG.USG_COL_STATUS).setValue(CONFIG.ACCOUNT_STATUS_DELETED);
      usageSheet.getRange(row, CONFIG.USG_COL_DELETED_AT).setValue(new Date());
      usageSheet.getRange(row, 1, 1, 13).setBackground("#f1f3f4").setFontColor("#80868b");
      usageSheet.getRange(row, 1, 1, CONFIG.USG_COL_STATUS - 1)
        .setTextStyle(SpreadsheetApp.newTextStyle().setStrikethrough(true).build());
      usageSheet.getRange(row, CONFIG.USG_COL_STATUS)
        .setBackground("#e8eaed").setFontWeight("bold")
        .setTextStyle(SpreadsheetApp.newTextStyle().setStrikethrough(false).build());
    }
  }
}

// ユーザー自体を削除済みに
function _markUserAsDeleted(userId) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName(CONFIG.SHEET_USERS);
  const data      = userSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][CONFIG.USR_COL_ID - 1]) === userId) {
      const row = i + 1;
      userSheet.getRange(row, CONFIG.USR_COL_STATUS).setValue(CONFIG.USER_STATUS_DELETED);
      userSheet.getRange(row, 1, 1, 6).setBackground("#f1f3f4").setFontColor("#80868b");
      break;
    }
  }
}

// 削除申請が却下された場合、全サービスを「運用中」に戻す
function _revertAllUserServicesStatus(userId) {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const usageSheet = ss.getSheetByName(CONFIG.SHEET_SERVICE_USAGE);
  const data       = usageSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][CONFIG.USG_COL_USER_ID - 1]) === userId &&
        data[i][CONFIG.USG_COL_STATUS - 1] === CONFIG.ACCOUNT_STATUS_DEL_REQ) {
      const row = i + 1;
      usageSheet.getRange(row, CONFIG.USG_COL_STATUS).setValue(CONFIG.ACCOUNT_STATUS_ACTIVE);
      usageSheet.getRange(row, CONFIG.USG_COL_DEL_REQ_ID).setValue("");
      usageSheet.getRange(row, 1, 1, 13)
        .setBackground(null).setFontColor(null)
        .setTextStyle(SpreadsheetApp.newTextStyle().setStrikethrough(false).build());
      usageSheet.getRange(row, CONFIG.USG_COL_STATUS)
        .setBackground("#e6f4ea").setFontColor("#137333").setFontWeight("bold");
    }
  }
}

// ============================================================
// 6. メール送信
// ============================================================
function _sendApproverNotification(requestId, applicantName, applicantEmail,
                                    targetName, targetEmail,
                                    service, type, reason, startDate) {
  const subject = `【アカウント申請】${service} の申請が届きました（${requestId}）`;
  const body =
`アカウント申請が届きました。スプレッドシートを確認して承認/却下してください。

■ 申請ID　　　: ${requestId}
■ 申請者　　　: ${applicantName}（${applicantEmail}）
■ 対象者　　　: ${targetName}（${targetEmail}）
■ サービス名　: ${service}
■ 種別　　　　: ${type}
■ 用途・理由　: ${reason}
■ 利用開始希望: ${startDate}

▼ スプレッドシートを開く
${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;
  MailApp.sendEmail(_getApproverEmails(), subject, body);
}

function _sendDeleteApproverNotification(deleteReqId, applicantName, applicantEmail,
                                          targetName, targetEmail, userId, reason) {
  const subject = `【アカウント削除申請】${targetName} の削除申請が届きました（${deleteReqId}）`;
  const body =
`アカウント削除申請が届きました。スプレッドシートを確認して承認/却下してください。

■ 削除申請ID　: ${deleteReqId}
■ 申請者　　　: ${applicantName}（${applicantEmail}）
■ 対象者　　　: ${targetName}（${targetEmail}）
■ ユーザーID　: ${userId}
■ 削除理由　　: ${reason}
※ 承認すると対象者に紐づく全サービスが削除済みになります

▼ スプレッドシートを開く
${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;
  MailApp.sendEmail(_getApproverEmails(), subject, body);
}

function _sendApplicantAcknowledgement(email, applicantName, requestId, targetName, serviceName) {
  MailApp.sendEmail(email,
    `【アカウント申請受付】${targetName} の ${serviceName} 申請を受け付けました`,
    `${applicantName} 様\n\n${targetName} の ${serviceName} アカウント申請を受け付けました。\n担当者が確認次第、結果をご連絡します。\n\n申請ID: ${requestId}`
  );
}

function _sendApplicantResult(email, applicantName, targetName, service, approved, comment) {
  const result = approved ? "承認" : "却下";
  MailApp.sendEmail(email,
    `【アカウント申請結果】${targetName} の ${service} 申請が${result}されました`,
    `${applicantName} 様\n\n${targetName} の ${service} アカウント申請が${result}されました。\n${comment ? `コメント: ${comment}\n` : ""}${approved ? "アカウントの発行手続きを進めます。" : "ご不明な点は担当者までお問い合わせください。"}`
  );
}

// ============================================================
// 7. ユーティリティ
// ============================================================
function _getApproverEmails() {
  return PropertiesService.getScriptProperties().getProperty("APPROVER_EMAILS")
    || CONFIG.APPROVER_EMAILS_FALLBACK;
}

function _parseFormResponse(e) {
  const result = {};
  e.response.getItemResponses().forEach(r => {
    result[r.getItem().getTitle()] = r.getResponse();
  });
  return result;
}

function _generateId(prefix) {
  const yymmdd = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyMMdd");
  const rand   = Math.random().toString(36).substring(2, 6).toUpperCase();
  return `${prefix}-${yymmdd}-${rand}`;
}

function _writeHeader(sheet, headers, bgColor) {
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground(bgColor).setFontColor("#ffffff").setFontWeight("bold");
  sheet.setFrozenRows(1);
}
