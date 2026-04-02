// ============================================================
// ServiceManager.gs
// サービスマスタ管理 / 管理画面 / フォーム連携 / 設定
// ============================================================

// サービスマスタの列インデックス（1始まり）
const SVC_COL = {
  ID:            1,
  NAME:          2,
  CATEGORY:      3,
  DESCRIPTION:   4,
  ACCOUNT_TYPES: 5,  // カンマ区切り文字列
  ENABLED:       6,
  CREATED_AT:    7,
  UPDATED_AT:    8,
};

// ============================================================
// フォームトリガーをスクリプトから登録（手動で一度だけ実行）
// ============================================================
function setupFormTriggers() {
  const props         = PropertiesService.getScriptProperties();
  const requestFormId = props.getProperty("REQUEST_FORM_ID");
  const deleteFormId  = props.getProperty("DELETE_FORM_ID");

  if (!requestFormId && !deleteFormId) {
    console.log("エラー: フォームIDが設定されていません。管理画面の設定タブで登録してください。");
    return;
  }

  // 既存のフォームトリガーを削除（重複防止）
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "onFormSubmit" ||
        t.getHandlerFunction() === "onDeleteFormSubmit") {
      ScriptApp.deleteTrigger(t);
    }
  });

  if (requestFormId) {
    ScriptApp.newTrigger("onFormSubmit")
      .forForm(requestFormId)
      .onFormSubmit()
      .create();
    console.log("新規申請フォームのトリガーを登録しました");
  }

  if (deleteFormId) {
    ScriptApp.newTrigger("onDeleteFormSubmit")
      .forForm(deleteFormId)
      .onFormSubmit()
      .create();
    console.log("削除申請フォームのトリガーを登録しました");
  }
}

// ============================================================
// メニュー & サイドバー
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("アカウント管理")
    .addItem("管理画面を開く", "openManagementSidebar")
    .addSeparator()
    .addItem("フォームの選択肢を同期", "syncServicesToForm")
    .addSeparator()
    .addItem("初期セットアップ（初回のみ）", "setup")
    .addToUi();
}

function openManagementSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("アカウント管理画面")
    .setWidth(440);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ============================================================
// サービス一覧取得（Sidebar から呼ばれる）
// ============================================================
function getServices() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_SERVICES);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  // サービス利用台帳から各サービスの運用中アカウント数を集計
  const usageSheet = ss.getSheetByName(CONFIG.SHEET_SERVICE_USAGE);
  const usageData  = usageSheet ? usageSheet.getDataRange().getValues().slice(1) : [];
  const countMap = {};
  usageData.forEach(row => {
    const svcId  = String(row[CONFIG.USG_COL_SERVICE_ID - 1]);
    const status = row[CONFIG.USG_COL_STATUS - 1];
    if (status === CONFIG.ACCOUNT_STATUS_ACTIVE) {
      countMap[svcId] = (countMap[svcId] || 0) + 1;
    }
  });

  return data.slice(1).map(row => ({
    id:           String(row[SVC_COL.ID - 1]),
    name:         String(row[SVC_COL.NAME - 1]),
    category:     String(row[SVC_COL.CATEGORY - 1] || ""),
    description:  String(row[SVC_COL.DESCRIPTION - 1] || ""),
    accountTypes: String(row[SVC_COL.ACCOUNT_TYPES - 1] || ""),
    enabled:      row[SVC_COL.ENABLED - 1] === true || row[SVC_COL.ENABLED - 1] === "TRUE",
    createdAt:    row[SVC_COL.CREATED_AT - 1]
                    ? Utilities.formatDate(new Date(row[SVC_COL.CREATED_AT - 1]), "Asia/Tokyo", "yyyy/MM/dd")
                    : "",
    activeCount:  countMap[String(row[SVC_COL.ID - 1])] || 0,
  }));
}

function getActiveServices() {
  return getServices().filter(s => s.enabled);
}

// サービス名でサービスを取得（Code.gs の onStatusChange から使用）
function getServiceByName(name) {
  return getServices().find(s => s.name === name) || null;
}

// ============================================================
// サービス追加・更新・削除（Sidebar から呼ばれる）
// ============================================================
function addService(params) {
  try {
    const { name, category, description, accountTypes, enabled } = params;
    if (!name) return { ok: false, error: "サービス名は必須です" };

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_SERVICES);
    if (!sheet) return { ok: false, error: "サービスマスタシートが見つかりません" };

    // 重複チェック
    const existing = getServices().find(s => s.name === name);
    if (existing) return { ok: false, error: `「${name}」はすでに登録されています` };

    const serviceId = _generateId("SVC");
    const now = new Date();
    sheet.appendRow([serviceId, name, category || "", description || "",
                     accountTypes || "", enabled !== false, now, now]);

    return { ok: true, serviceId };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function updateService(serviceId, params) {
  try {
    const { name, category, description, accountTypes, enabled } = params;
    if (!name) return { ok: false, error: "サービス名は必須です" };

    const row = _getServiceRow(serviceId);
    if (!row) return { ok: false, error: "サービスが見つかりません" };

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_SERVICES);

    // 名前の重複チェック（自分以外）
    const duplicate = getServices().find(s => s.name === name && s.id !== serviceId);
    if (duplicate) return { ok: false, error: `「${name}」は別のサービスで使用されています` };

    sheet.getRange(row, SVC_COL.NAME).setValue(name);
    sheet.getRange(row, SVC_COL.CATEGORY).setValue(category || "");
    sheet.getRange(row, SVC_COL.DESCRIPTION).setValue(description || "");
    sheet.getRange(row, SVC_COL.ACCOUNT_TYPES).setValue(accountTypes || "");
    sheet.getRange(row, SVC_COL.ENABLED).setValue(enabled !== false);
    sheet.getRange(row, SVC_COL.UPDATED_AT).setValue(new Date());

    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function toggleServiceEnabled(serviceId, enabled) {
  try {
    const row = _getServiceRow(serviceId);
    if (!row) return { ok: false, error: "サービスが見つかりません" };

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_SERVICES);
    sheet.getRange(row, SVC_COL.ENABLED).setValue(enabled);
    sheet.getRange(row, SVC_COL.UPDATED_AT).setValue(new Date());

    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function deleteService(serviceId) {
  try {
    // 運用中・削除申請中のアカウントが存在する場合は削除不可
    const accounts = getAccountsByService(serviceId);
    const blockers = accounts.filter(a =>
      a.status === CONFIG.ACCOUNT_STATUS_ACTIVE ||
      a.status === CONFIG.ACCOUNT_STATUS_DEL_REQ
    );
    if (blockers.length > 0) {
      return {
        ok: false,
        error: `削除できません。${blockers.length} 件の運用中／削除申請中アカウントが存在します。先にアカウントを削除済みにしてください。`
      };
    }

    const row = _getServiceRow(serviceId);
    if (!row) return { ok: false, error: "サービスが見つかりません" };

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_SERVICES);
    sheet.deleteRow(row);

    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ============================================================
// アカウント一覧（Sidebar から呼ばれる）
// ============================================================
function getAccountsByService(serviceId) {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const usageSheet = ss.getSheetByName(CONFIG.SHEET_SERVICE_USAGE);
  if (!usageSheet) return [];

  const data = usageSheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  return data.slice(1)
    .filter(row => serviceId === "all" || String(row[CONFIG.USG_COL_SERVICE_ID - 1]) === serviceId)
    .map(row => ({
      usageId:     String(row[CONFIG.USG_COL_USAGE_ID - 1]),
      userId:      String(row[CONFIG.USG_COL_USER_ID - 1]),
      userName:    String(row[CONFIG.USG_COL_USER_NAME - 1]),
      serviceId:   String(row[CONFIG.USG_COL_SERVICE_ID - 1]),
      serviceName: String(row[CONFIG.USG_COL_SERVICE_NAME - 1]),
      accountType: String(row[CONFIG.USG_COL_ACCOUNT_TYPE - 1]),
      startDate:   _fmtDate(row[CONFIG.USG_COL_START_DATE - 1]),
      purpose:     String(row[CONFIG.USG_COL_PURPOSE - 1]),
      requestId:   String(row[CONFIG.USG_COL_REQUEST_ID - 1]),
      createdAt:   _fmtDate(row[CONFIG.USG_COL_CREATED_AT - 1]),
      status:      String(row[CONFIG.USG_COL_STATUS - 1]),
      deletedAt:   _fmtDate(row[CONFIG.USG_COL_DELETED_AT - 1]),
    }));
}

function getAccountSummary() {
  return getServices().map(svc => {
    const accounts = getAccountsByService(svc.id);
    return {
      serviceId:       svc.id,
      serviceName:     svc.name,
      category:        svc.category,
      enabled:         svc.enabled,
      total:           accounts.length,
      active:          accounts.filter(a => a.status === CONFIG.ACCOUNT_STATUS_ACTIVE).length,
      pendingDeletion: accounts.filter(a => a.status === CONFIG.ACCOUNT_STATUS_DEL_REQ).length,
      deleted:         accounts.filter(a => a.status === CONFIG.ACCOUNT_STATUS_DELETED).length,
    };
  });
}

// ============================================================
// フォーム同期（サービス選択肢を Google Forms に反映）
// ============================================================
function syncServicesToForm() {
  const serviceNames = getActiveServices().map(s => s.name);
  if (serviceNames.length === 0) {
    return { ok: false, error: "有効なサービスが存在しません" };
  }

  const props         = PropertiesService.getScriptProperties();
  const requestFormId = props.getProperty("REQUEST_FORM_ID") || "";
  const deleteFormId  = props.getProperty("DELETE_FORM_ID")  || "";

  if (!requestFormId && !deleteFormId) {
    return { ok: false, error: "フォームIDが設定されていません。設定タブでフォームIDを登録してください。" };
  }

  let updatedCount = 0;
  const errors = [];

  [requestFormId, deleteFormId].filter(Boolean).forEach(formId => {
    try {
      const form  = FormApp.openById(formId);
      const items = form.getItems();
      items.forEach(item => {
        if (item.getTitle() === "サービス名") {
          const listItem = item.asListItem();
          listItem.setChoices(serviceNames.map(n => listItem.createChoice(n)));
          updatedCount++;
        }
      });
    } catch(e) {
      errors.push(`フォームID ${formId}: ${e.message}`);
    }
  });

  if (errors.length > 0) return { ok: false, error: errors.join("\n") };
  return { ok: true, updatedCount };
}

// ============================================================
// 設定（承認者メール・フォームID）
// ============================================================
function getSettings() {
  const props = PropertiesService.getScriptProperties();
  return {
    approverEmails: props.getProperty("APPROVER_EMAILS") || CONFIG.APPROVER_EMAILS_FALLBACK,
    requestFormId:  props.getProperty("REQUEST_FORM_ID") || "",
    deleteFormId:   props.getProperty("DELETE_FORM_ID")  || "",
  };
}

function saveSettings(settings) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperties({
      APPROVER_EMAILS: settings.approverEmails || "",
      REQUEST_FORM_ID: settings.requestFormId  || "",
      DELETE_FORM_ID:  settings.deleteFormId   || "",
    });
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ============================================================
// サービスマスタシートのセットアップ（setup() から呼ばれる）
// ============================================================
function _setupServiceMasterSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_SERVICES);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_SERVICES);

  const headers = [
    "サービスID", "サービス名", "カテゴリ",
    "説明", "アカウント種別", "有効", "登録日時", "更新日時"
  ];
  _writeHeader(sheet, headers, "#6200ea");

  // アカウント種別はテキスト形式（カンマ区切りで入力）
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 200);
  sheet.setColumnWidth(5, 160);
  sheet.setColumnWidth(6, 60);

  // サンプルデータを1件追加（Google Workspace）
  if (sheet.getLastRow() <= 1) {
    const now = new Date();
    sheet.appendRow([
      _generateId("SVC"), "Google Workspace", "コラボレーション",
      "Google の各種サービス（Gmail, Drive, Meet など）",
      "一般,管理者", true, now, now
    ]);
  }

  Logger.log("サービスマスタシート セットアップ完了");
}

// ============================================================
// 内部ヘルパー
// ============================================================
function _getServiceRow(serviceId) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_SERVICES);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][SVC_COL.ID - 1]) === serviceId) return i + 1;
  }
  return null;
}

function _fmtDate(val) {
  if (!val) return "";
  try { return Utilities.formatDate(new Date(val), "Asia/Tokyo", "yyyy/MM/dd"); }
  catch(e) { return String(val); }
}
