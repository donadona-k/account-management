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
  USR_COL_EMP_NO:     2,  // 社員番号
  USR_COL_NAME:       3,
  USR_COL_DEPT:       4,
  USR_COL_EMAIL:      5,
  USR_COL_CATEGORY:   6,  // 区分（正社員／クルー／業務委託／その他）
  USR_COL_CREATED_AT: 7,
  USR_COL_STATUS:     8,

  USR_CATEGORIES: ["正社員", "クルー", "業務委託", "その他"],

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
// 社員マスタの列インデックス（別スプレッドシート）
// ============================================================
const EMP = {
  COL_EMP_NO:     1,
  COL_NAME:       2,
  COL_DEPT:       3,
  COL_EMAIL:      4,
  COL_CATEGORY:   5,
  COL_GENDER:     6,
  COL_BIRTH_DATE: 7,
  COL_JOIN_DATE:  8,
  COL_LEAVE_DATE: 9,
  COL_STATUS:     10,

  STATUSES:   ["入社前", "在籍中", "休職中", "退職済"],
  CATEGORIES: ["正社員", "クルー", "業務委託", "その他"],
  GENDERS:    ["男性", "女性", "その他"],

  SHEET_NAME: "社員マスタ",
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

  const headers = ["ユーザーID", "社員番号", "氏名", "部署", "メールアドレス", "区分", "登録日時", "ステータス"];
  _writeHeader(sheet, headers, "#1565c0");

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([CONFIG.USER_STATUS_ACTIVE, CONFIG.USER_STATUS_DELETED])
    .setAllowInvalid(false).build();
  sheet.getRange(2, CONFIG.USR_COL_STATUS, 1000).setDataValidation(statusRule);

  const categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.USR_CATEGORIES)
    .setAllowInvalid(false).build();
  sheet.getRange(2, CONFIG.USR_COL_CATEGORY, 1000).setDataValidation(categoryRule);

  sheet.setColumnWidth(1, 130); sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(5, 200); sheet.setColumnWidth(6, 100);
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

  // 新スキーマ：列8=削除種別、列9=削除対象利用ID、列10=削除理由、列11=ステータス
  const headers = [
    "削除申請ID", "タイムスタンプ",
    "申請者名", "申請者メール",
    "対象者名", "対象者メール", "ユーザーID",
    "削除種別", "削除対象利用ID",
    "削除理由", "ステータス", "承認者コメント", "承認/却下日時"
  ];
  _writeHeader(sheet, headers, "#d93025");

  const statusCol = headers.indexOf("ステータス") + 1; // 11列目
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([CONFIG.STATUS_PENDING, CONFIG.STATUS_APPROVED, CONFIG.STATUS_REJECTED])
    .setAllowInvalid(false).build();
  sheet.getRange(2, statusCol, 1000).setDataValidation(rule);

  sheet.setColumnWidth(1, 130); sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(9, 200); sheet.setColumnWidth(10, 250); sheet.setColumnWidth(11, 80);
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
// 2. Web App エントリポイント
// ============================================================
function doGet(e) {
  const form = e && e.parameter && e.parameter.form;
  if (form === "delete") {
    return HtmlService.createHtmlOutputFromFile("DeleteForm")
      .setTitle("アカウント削除申請")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutputFromFile("RequestForm")
    .setTitle("アカウント申請")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// 3. Web App から呼ばれる関数
// ============================================================

// ログインユーザーの情報を社員マスタから取得
// → {email, name, dept, found: boolean}
function getApplicantInfo() {
  const email = Session.getActiveUser().getEmail();
  const staff = _findStaffByEmail(email);
  if (!staff) return { email, name: "", dept: "", found: false };
  return {
    email,
    name:  String(staff[EMP.COL_NAME - 1]),
    dept:  String(staff[EMP.COL_DEPT - 1]),
    found: true,
  };
}

// 新規申請送信（複数サービス対応）
// data = {applicantName, applicantEmail, targetName, targetDept, targetEmail,
//         services: [{serviceId, serviceName, accountType, purpose, startDate}, ...]}
function submitNewRequest(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requestSheet = ss.getSheetByName(CONFIG.SHEET_REQUESTS);

  const { applicantName, applicantEmail, targetName, targetDept, targetEmail, services } = data;

  if (!services || services.length === 0) {
    return { ok: false, error: "サービスを1つ以上選択してください" };
  }

  const requestIds = [];
  services.forEach(svc => {
    const requestId = _generateId("REQ");
    requestIds.push(requestId);

    requestSheet.appendRow([
      requestId, new Date(),
      applicantName, applicantEmail,
      targetName, targetDept, targetEmail,
      svc.serviceName, svc.accountType, svc.purpose, svc.startDate,
      CONFIG.STATUS_PENDING, "", "", ""
    ]);

    const approverEmails = getApproverEmailsForService(svc.serviceName);
    _sendApproverNotification(requestId, applicantName, applicantEmail,
                               targetName, targetEmail,
                               svc.serviceName, svc.accountType, svc.purpose, svc.startDate,
                               approverEmails);
    if (applicantEmail) {
      _sendApplicantAcknowledgement(applicantEmail, applicantName, requestId, targetName, svc.serviceName);
    }
  });

  return { ok: true, requestIds };
}

// 削除申請送信（一部削除 / 全削除対応）
// data = {applicantName, applicantEmail, targetEmail, deleteReason, deleteAll, usageIds:[]}
function submitDeleteRequest(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const deleteSheet = ss.getSheetByName(CONFIG.SHEET_DELETE_REQUESTS);
  const userSheet   = ss.getSheetByName(CONFIG.SHEET_USERS);

  const { applicantName, applicantEmail, targetEmail, deleteReason, deleteAll, usageIds } = data;

  const userRow = _findUserRowByEmail(userSheet, targetEmail);
  if (!userRow) {
    return { ok: false, error: `メールアドレス "${targetEmail}" のユーザーが見つかりません` };
  }

  const userData   = userSheet.getRange(userRow, 1, 1, 7).getValues()[0];
  const userId     = String(userData[CONFIG.USR_COL_ID - 1]);
  const targetName = String(userData[CONFIG.USR_COL_NAME - 1]);

  const deleteReqId     = _generateId("DEL");
  const deleteType      = deleteAll ? "全削除" : "一部削除";
  const deleteTargetIds = deleteAll ? "ALL" : (usageIds || []).join(",");

  deleteSheet.appendRow([
    deleteReqId, new Date(),
    applicantName, applicantEmail,
    targetName, targetEmail, userId,
    deleteType, deleteTargetIds,
    deleteReason, CONFIG.STATUS_PENDING, "", ""
  ]);

  _markServicesAsPendingDeletion(userId, usageIds || [], deleteAll, deleteReqId);

  _sendDeleteApproverNotification(deleteReqId, applicantName, applicantEmail,
                                   targetName, targetEmail, userId, deleteReason);
  if (applicantEmail) {
    MailApp.sendEmail(
      applicantEmail,
      `【削除申請受付】${targetName} のアカウント削除申請を受け付けました`,
      `${applicantName} 様\n\n${targetName}（${targetEmail}）のアカウント削除申請を受け付けました。\n${deleteAll ? "紐づくすべてのサービス" : "選択したサービス"}が削除申請中となりました。\n\n削除申請ID: ${deleteReqId}`
    );
  }

  return { ok: true, deleteReqId };
}

// 対象者のメールで検索し、ユーザー情報 + 運用中サービス一覧を返す
// → {found, userId, name, dept, services:[{usageId, serviceName, accountType}]}
function lookupUserByEmail(email) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName(CONFIG.SHEET_USERS);
  const row       = _findUserRowByEmail(userSheet, email);
  if (!row) return { found: false };

  const userData = userSheet.getRange(row, 1, 1, 8).getValues()[0];
  const userId   = String(userData[CONFIG.USR_COL_ID - 1]);
  const name     = String(userData[CONFIG.USR_COL_NAME - 1]);
  const dept     = String(userData[CONFIG.USR_COL_DEPT - 1]);
  const status   = String(userData[CONFIG.USR_COL_STATUS - 1]);

  if (status === CONFIG.USER_STATUS_DELETED) {
    return { found: false, deleted: true };
  }

  const usageSheet = ss.getSheetByName(CONFIG.SHEET_SERVICE_USAGE);
  const usageData  = usageSheet.getDataRange().getValues();
  const services   = [];
  for (let i = 1; i < usageData.length; i++) {
    const uRow = usageData[i];
    if (String(uRow[CONFIG.USG_COL_USER_ID - 1]) === userId &&
        uRow[CONFIG.USG_COL_STATUS - 1] === CONFIG.ACCOUNT_STATUS_ACTIVE) {
      services.push({
        usageId:     String(uRow[CONFIG.USG_COL_USAGE_ID - 1]),
        serviceName: String(uRow[CONFIG.USG_COL_SERVICE_NAME - 1]),
        accountType: String(uRow[CONFIG.USG_COL_ACCOUNT_TYPE - 1]),
      });
    }
  }

  return { found: true, userId, name, dept, services };
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
  // 削除申請一覧のステータス列 = 11列目（新スキーマ）
  if (sheetName === CONFIG.SHEET_DELETE_REQUESTS && range.getColumn() === 11) {
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
  // 新スキーマ：13列（削除種別=8列目、削除対象利用ID=9列目、コメント=12列目、日時=13列目）
  const rowData = sheet.getRange(row, 1, 1, 13).getValues()[0];
  const [deleteReqId, , applicantName, applicantEmail,
         targetName, targetEmail, userId,
         deleteType, deleteTargetIds,
         , , comment] = rowData;

  const deleteAll  = deleteType === "全削除";
  const usageIds   = deleteAll
    ? []
    : String(deleteTargetIds).split(",").map(s => s.trim()).filter(Boolean);

  if (newStatus === CONFIG.STATUS_APPROVED) {
    sheet.getRange(row, 13).setValue(new Date());
    _markServicesAsDeleted(userId, usageIds, deleteAll);
    if (deleteAll) {
      _markUserAsDeleted(userId);
    } else {
      _updateUserStatusIfAllDeleted(userId);
    }
    if (applicantEmail) {
      MailApp.sendEmail(applicantEmail,
        `【削除申請承認】${targetName} のアカウントが削除されました`,
        `${applicantName} 様\n\n${targetName}（${targetEmail}）のアカウント削除申請が承認されました。\n${deleteAll ? "紐づくすべてのサービス" : "選択したサービス"}が削除済みとなりました。\n${comment ? `コメント: ${comment}` : ""}`
      );
    }
  } else if (newStatus === CONFIG.STATUS_REJECTED) {
    sheet.getRange(row, 13).setValue(new Date());
    _revertServicesStatus(userId, usageIds, deleteAll);
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
  userSheet.appendRow([userId, "", name, dept, email, "", new Date(), CONFIG.USER_STATUS_ACTIVE]);
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

// 指定サービス（usageIds）または全サービスを「削除申請中」に（赤背景）
function _markServicesAsPendingDeletion(userId, usageIds, deleteAll, deleteReqId) {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const usageSheet = ss.getSheetByName(CONFIG.SHEET_SERVICE_USAGE);
  const data       = usageSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][CONFIG.USG_COL_USER_ID - 1]) !== userId) continue;
    if (data[i][CONFIG.USG_COL_STATUS - 1] !== CONFIG.ACCOUNT_STATUS_ACTIVE) continue;

    const usageId = String(data[i][CONFIG.USG_COL_USAGE_ID - 1]);
    if (!deleteAll && !usageIds.includes(usageId)) continue;

    const row = i + 1;
    usageSheet.getRange(row, CONFIG.USG_COL_STATUS).setValue(CONFIG.ACCOUNT_STATUS_DEL_REQ);
    usageSheet.getRange(row, CONFIG.USG_COL_DEL_REQ_ID).setValue(deleteReqId);
    usageSheet.getRange(row, 1, 1, 13).setBackground("#fce8e6").setFontColor("#c5221f");
    usageSheet.getRange(row, CONFIG.USG_COL_STATUS)
      .setBackground("#f28b82").setFontWeight("bold");
  }
}

// 指定サービス（usageIds）または全サービスを「削除済み」に（グレー＋取り消し線）
function _markServicesAsDeleted(userId, usageIds, deleteAll) {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const usageSheet = ss.getSheetByName(CONFIG.SHEET_SERVICE_USAGE);
  const data       = usageSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][CONFIG.USG_COL_USER_ID - 1]) !== userId) continue;
    if (data[i][CONFIG.USG_COL_STATUS - 1] === CONFIG.ACCOUNT_STATUS_DELETED) continue;

    const usageId = String(data[i][CONFIG.USG_COL_USAGE_ID - 1]);
    if (!deleteAll && !usageIds.includes(usageId)) continue;

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

// ============================================================
// 社員マスタ参照ヘルパー
// ============================================================

// Script Properties の STAFF_MASTER_ID から社員マスタシートを取得
function _getStaffMasterSheet() {
  const id = PropertiesService.getScriptProperties().getProperty("STAFF_MASTER_ID");
  if (!id) return null;
  try {
    const ss = SpreadsheetApp.openById(id);
    return ss.getSheetByName(EMP.SHEET_NAME) || ss.getSheets()[0];
  } catch(e) {
    console.error("社員マスタを開けませんでした: " + e.message);
    return null;
  }
}

// メールアドレスで社員マスタを検索し、行データ配列を返す（見つからなければ null）
function _findStaffByEmail(email) {
  const sheet = _getStaffMasterSheet();
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][EMP.COL_EMAIL - 1]).trim() === email.trim()) {
      return data[i];
    }
  }
  return null;
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
      userSheet.getRange(row, 1, 1, 8).setBackground("#f1f3f4").setFontColor("#80868b");
      break;
    }
  }
}

// 削除申請が却下された場合、対象サービスを「運用中」に戻す
function _revertServicesStatus(userId, usageIds, deleteAll) {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const usageSheet = ss.getSheetByName(CONFIG.SHEET_SERVICE_USAGE);
  const data       = usageSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][CONFIG.USG_COL_USER_ID - 1]) !== userId) continue;
    if (data[i][CONFIG.USG_COL_STATUS - 1] !== CONFIG.ACCOUNT_STATUS_DEL_REQ) continue;

    const usageId = String(data[i][CONFIG.USG_COL_USAGE_ID - 1]);
    if (!deleteAll && !usageIds.includes(usageId)) continue;

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

// 一部削除承認後：残存する運用中サービスが0件ならユーザーも削除済みに
function _updateUserStatusIfAllDeleted(userId) {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const usageSheet = ss.getSheetByName(CONFIG.SHEET_SERVICE_USAGE);
  const data       = usageSheet.getDataRange().getValues();

  const hasActive = data.slice(1).some(row =>
    String(row[CONFIG.USG_COL_USER_ID - 1]) === userId &&
    row[CONFIG.USG_COL_STATUS - 1] === CONFIG.ACCOUNT_STATUS_ACTIVE
  );

  if (!hasActive) {
    _markUserAsDeleted(userId);
  }
}

// ============================================================
// 6. メール送信
// ============================================================
function _sendApproverNotification(requestId, applicantName, applicantEmail,
                                    targetName, targetEmail,
                                    service, type, reason, startDate,
                                    approverEmails) {
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
  MailApp.sendEmail(approverEmails || _getApproverEmails(), subject, body);
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
※ 承認すると対象サービスが削除済みになります

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

function _generateId(prefix) {
  const yymmdd = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyMMdd");
  const rand   = Math.random().toString(36).substring(2, 6).toUpperCase();
  return `${prefix}-${yymmdd}-${rand}`;
}

// ユーザー台帳に社員番号・区分列を追加するマイグレーション（既存データがある場合に一度だけ実行）
// 実行順序：社員番号（B列挿入）→ 区分（メールアドレスの右に挿入）の順で処理
function migrateUserSheetColumns() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName(CONFIG.SHEET_USERS);
  if (!userSheet) { console.log("ユーザー台帳シートが見つかりません"); return; }

  let headers = userSheet.getRange(1, 1, 1, userSheet.getLastColumn()).getValues()[0];

  // ① 社員番号（B列 = 2列目）
  if (headers[1] !== "社員番号") {
    userSheet.insertColumnBefore(2);
    userSheet.getRange(1, 2)
      .setValue("社員番号")
      .setBackground("#1565c0").setFontColor("#ffffff").setFontWeight("bold");
    userSheet.setColumnWidth(2, 100);
    console.log("社員番号列を追加しました");
    // 挿入後にヘッダーを再取得
    headers = userSheet.getRange(1, 1, 1, userSheet.getLastColumn()).getValues()[0];
  } else {
    console.log("社員番号列はすでに存在します");
  }

  // ② 区分（メールアドレスの右 = 6列目）
  if (!headers.includes("区分")) {
    const emailCol = headers.indexOf("メールアドレス") + 1; // 1始まり
    userSheet.insertColumnAfter(emailCol);
    const categoryCol = emailCol + 1;
    userSheet.getRange(1, categoryCol)
      .setValue("区分")
      .setBackground("#1565c0").setFontColor("#ffffff").setFontWeight("bold");
    userSheet.setColumnWidth(categoryCol, 100);

    // 区分のドロップダウンを設定
    const categoryRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.USR_CATEGORIES)
      .setAllowInvalid(false).build();
    userSheet.getRange(2, categoryCol, 1000).setDataValidation(categoryRule);

    console.log("区分列を追加しました");
  } else {
    console.log("区分列はすでに存在します");
  }
}

function _writeHeader(sheet, headers, bgColor) {
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground(bgColor).setFontColor("#ffffff").setFontWeight("bold");
  sheet.setFrozenRows(1);
}
