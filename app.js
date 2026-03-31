/**
 * 家計簿 — ローカルストレージ永続化
 * スキーマ: { version, entries: Entry[], fixedExpenses, expenseRecords? }
 * Entry: { id, kind, date (YYYY-MM-DD), amount (正の整数円), category, note }
 * ExpenseRecord: { id, date, memo, receiptDataUrl? } — 取引とは独立したメモ・領収書写真
 */

const STORAGE_KEY = "household-budget-v1";

/** Googleドライブへのアップロード用（ユーザーが GCP で発行したクライアント ID と保存先フォルダ） */
const GDRIVE_CLIENT_ID_KEY = "householdGdriveClientId";
const GDRIVE_FOLDER_ID_KEY = "householdGdriveFolderId";
/** ユーザー指定フォルダに JSON を作成するため（個人用 OAuth クライアント前提。サーバーは不要） */
const GDRIVE_OAUTH_SCOPE = "https://www.googleapis.com/auth/drive";

const INCOME_CATEGORIES = ["給与", "賞与", "副業", "投資", "案件", "その他（収入）"];
/** 案件登録タブ用（収入・note に案件名を格納） */
const PROJECT_INCOME_CATEGORY = "案件";
/** 当日クイック入力（食費と同様・カテゴリ固定） */
const SAISON_CATEGORY = "セゾン";
const MITSUI_CATEGORY = "三井住友カード";

/** 食費（PayPay）クイック入力で使うカテゴリ（変動内訳の PayPay（食費）と同一） */
const FOOD_CATEGORY = "PayPay（食費）";

/** 変動支出フォームのプルダウン項目（PayPay（食費）は食費（PayPay）専用タブのため含めない） */
const WATER_CATEGORY = "水道";
const VARIABLE_PAYMENT_LABELS = ["dカード", WATER_CATEGORY];

/** 変動内訳の備考（カテゴリ名は正規化後と一致） */
const VARIABLE_BREAKDOWN_REMARKS = /** @type {Record<string, string>} */ ({
  [FOOD_CATEGORY]: "毎月月末〆・翌27日引き落とし",
  [SAISON_CATEGORY]: "毎月10日〆・翌4日引き落とし",
  [MITSUI_CATEGORY]: "毎月15日〆・翌10日引き落とし",
  dカード: "月15日〆・翌10日引き落とし",
  [WATER_CATEGORY]: "奇数月に引き落とし",
});

const EXPENSE_CATEGORIES = [
  "食費",
  "日用品",
  "住居・光熱",
  "通信",
  "交通",
  "医療・保険",
  "教育",
  "娯楽",
  SAISON_CATEGORY,
  MITSUI_CATEGORY,
  FOOD_CATEGORY,
  ...VARIABLE_PAYMENT_LABELS,
  "その他（支出）",
];

/** 初回・旧データ向けの固定支出テンプレート（名前と月額のみ） */
const DEFAULT_FIXED_EXPENSE_TEMPLATE = [
  { name: "家賃", amount: 73490 },
  { name: "生命保険", amount: 3013 },
  { name: "ペイディ", amount: 11658 },
  { name: "ゼンロウサイ", amount: 6230 },
];

/** @typedef {'income'|'expense'} Kind */

/**
 * @typedef {object} Entry
 * @property {string} id
 * @property {Kind} kind
 * @property {string} date
 * @property {number} amount
 * @property {string} category
 * @property {string} note
 */

/**
 * @typedef {object} FixedExpenseItem
 * @property {string} id
 * @property {string} name
 * @property {number} amount
 */

/**
 * @typedef {object} ExpenseRecord
 * @property {string} id
 * @property {string} date YYYY-MM-DD
 * @property {string} memo
 * @property {string} [receiptDataUrl] data:image/jpeg;base64,...
 */

function createExpenseRecordId() {
  if (typeof crypto !== "undefined" && crypto.randomUUID) return `er-${crypto.randomUUID()}`;
  return `er-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
}

/** @param {unknown} raw */
function normalizeExpenseRecordsArray(raw) {
  if (!Array.isArray(raw)) return [];
  const out = /** @type {ExpenseRecord[]} */ ([]);
  for (const x of raw) {
    if (!x || typeof x !== "object") continue;
    const o = /** @type {Record<string, unknown>} */ (x);
    const id = typeof o.id === "string" && o.id.length ? o.id : createExpenseRecordId();
    const date = typeof o.date === "string" && /^\d{4}-\d{2}-\d{2}$/.test(o.date) ? o.date : "";
    const memo = typeof o.memo === "string" ? o.memo.slice(0, 2000) : "";
    const url = typeof o.receiptDataUrl === "string" ? o.receiptDataUrl.trim() : "";
    const receiptDataUrl = url.startsWith("data:image/") ? url : undefined;
    if (!date) continue;
    if (!memo && !receiptDataUrl) continue;
    out.push({ id, date, memo, ...(receiptDataUrl ? { receiptDataUrl } : {}) });
  }
  return out;
}

/** @param {{ expenseRecords?: unknown }} data */
function expenseRecordsFromStoredData(data) {
  if (data.expenseRecords === undefined) return [];
  return normalizeExpenseRecordsArray(data.expenseRecords);
}

function createFixedExpenseId() {
  if (typeof crypto !== "undefined" && crypto.randomUUID) return `fx-${crypto.randomUUID()}`;
  return `fx-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
}

/** @returns {FixedExpenseItem[]} */
function defaultFixedExpenses() {
  return DEFAULT_FIXED_EXPENSE_TEMPLATE.map((row) => ({
    id: createFixedExpenseId(),
    name: row.name,
    amount: row.amount,
  }));
}

/** @param {unknown} raw */
function normalizeFixedExpensesArray(raw) {
  if (!Array.isArray(raw)) return [];
  const out = /** @type {FixedExpenseItem[]} */ ([]);
  for (const x of raw) {
    if (!x || typeof x !== "object") continue;
    const o = /** @type {Record<string, unknown>} */ (x);
    const name = typeof o.name === "string" ? o.name.trim() : "";
    const amount = typeof o.amount === "number" && Number.isInteger(o.amount) && o.amount >= 0 ? o.amount : NaN;
    if (!name || !Number.isFinite(amount)) continue;
    const id = typeof o.id === "string" && o.id.length ? o.id : createFixedExpenseId();
    out.push({ id, name, amount });
  }
  return out;
}

/** @param {{ fixedExpenses?: unknown }} data */
function fixedExpensesFromStoredData(data) {
  if (data.fixedExpenses === undefined) return defaultFixedExpenses();
  return normalizeFixedExpensesArray(data.fixedExpenses);
}

/** @param {{ fixedExpenses: FixedExpenseItem[] }} s */
function fixedExpenseTotal(s) {
  return s.fixedExpenses.reduce((sum, i) => sum + i.amount, 0);
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return { version: 1, entries: [], fixedExpenses: defaultFixedExpenses(), expenseRecords: [] };
    }
    const data = JSON.parse(raw);
    if (!data || !Array.isArray(data.entries)) {
      return { version: 1, entries: [], fixedExpenses: defaultFixedExpenses(), expenseRecords: [] };
    }
    return {
      version: 1,
      entries: data.entries.filter(isValidEntry),
      fixedExpenses: fixedExpensesFromStoredData(data),
      expenseRecords: expenseRecordsFromStoredData(data),
    };
  } catch {
    return { version: 1, entries: [], fixedExpenses: defaultFixedExpenses(), expenseRecords: [] };
  }
}

/** @param {unknown} e */
function isValidEntry(e) {
  if (!e || typeof e !== "object") return false;
  const o = /** @type {Record<string, unknown>} */ (e);
  return (
    typeof o.id === "string" &&
    (o.kind === "income" || o.kind === "expense") &&
    typeof o.date === "string" &&
    typeof o.amount === "number" &&
    o.amount > 0 &&
    Number.isInteger(o.amount) &&
    typeof o.category === "string" &&
    typeof o.note === "string"
  );
}

/** @param {{ version: number, entries: Entry[], fixedExpenses: FixedExpenseItem[], expenseRecords: ExpenseRecord[] }} state */
function saveState(state) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch (e) {
    const name = e && typeof e === "object" && "name" in e ? String(/** @type {{ name?: string }} */ (e).name) : "";
    const msg = e instanceof Error ? e.message : String(e);
    if (name === "QuotaExceededError" || msg.includes("QuotaExceeded")) {
      alert(
        "ブラウザの保存容量の上限に達した可能性があります。古い経費の写真を削除するか、バックアップ後に不要なデータを減らしてください。"
      );
    } else {
      alert(`保存に失敗しました: ${msg}`);
    }
  }
}

function newId() {
  if (typeof crypto !== "undefined" && crypto.randomUUID) return crypto.randomUUID();
  return `e-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
}

const EXPENSE_RECEIPT_MAX_EDGE = 1400;
const EXPENSE_RECEIPT_JPEG_QUALITY = 0.82;

/**
 * 領収書写真を縮小して JPEG の data URL にする（容量抑制）。
 * @param {File} file
 * @returns {Promise<string>}
 */
function compressImageFileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    if (!file.type.startsWith("image/")) {
      reject(new Error("画像ファイルを選んでください"));
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = String(reader.result || "");
      const img = new Image();
      img.onload = () => {
        let w = img.naturalWidth;
        let h = img.naturalHeight;
        if (w < 1 || h < 1) {
          reject(new Error("画像のサイズを読み取れませんでした"));
          return;
        }
        const maxEdge = EXPENSE_RECEIPT_MAX_EDGE;
        if (w > maxEdge || h > maxEdge) {
          if (w >= h) {
            h = Math.round((h * maxEdge) / w);
            w = maxEdge;
          } else {
            w = Math.round((w * maxEdge) / h);
            h = maxEdge;
          }
        }
        const canvas = document.createElement("canvas");
        canvas.width = w;
        canvas.height = h;
        const ctx = canvas.getContext("2d");
        if (!ctx) {
          reject(new Error("画像の加工に失敗しました"));
          return;
        }
        ctx.drawImage(img, 0, 0, w, h);
        resolve(canvas.toDataURL("image/jpeg", EXPENSE_RECEIPT_JPEG_QUALITY));
      };
      img.onerror = () => reject(new Error("画像を読み込めませんでした（形式が未対応の可能性があります）"));
      img.src = dataUrl;
    };
    reader.onerror = () => reject(new Error("ファイルの読み込みに失敗しました"));
    reader.readAsDataURL(file);
  });
}

/** @param {string} yyyymm */
function parseMonthKey(yyyymm) {
  const [y, m] = yyyymm.split("-").map(Number);
  return new Date(y, m - 1, 1);
}

/** @param {Date} d */
function monthKeyFromDate(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  return `${y}-${m}`;
}

/** 月キー（YYYY-MM）に対応する月初日（ISO 日付） */
function firstDayOfMonthKey(monthKey) {
  return `${monthKey}-01`;
}

/** @param {string} monthKey YYYY-MM */
function formatMonthLongJaFromKey(monthKey) {
  const d = parseMonthKey(monthKey);
  return d.toLocaleDateString("ja-JP", { year: "numeric", month: "long" });
}

/** @param {string} monthKey YYYY-MM @param {number} delta */
function addMonthsToMonthKey(monthKey, delta) {
  const d = parseMonthKey(monthKey);
  d.setMonth(d.getMonth() + delta);
  return monthKeyFromDate(d);
}

/**
 * 引き落とし月（トップの表示月）が billingMonthKey のとき、その行に集計される利用の期間（カードルールの逆算）。
 * @param {string} billingMonthKey YYYY-MM
 * @param {string} categoryName 行のカテゴリ表示名
 * @returns {{ start: Date, end: Date }}
 */
function getUsagePeriodBoundsForBillingMonth(billingMonthKey, categoryName) {
  const norm = normalizeExpenseCategory(categoryName) || categoryName;

  if (norm === FOOD_CATEGORY) {
    const u = addMonthsToMonthKey(billingMonthKey, -1);
    const start = parseMonthKey(u);
    const end = new Date(start.getFullYear(), start.getMonth() + 1, 0);
    return { start, end };
  }
  if (norm === SAISON_CATEGORY) {
    const dStart = parseMonthKey(addMonthsToMonthKey(billingMonthKey, -2));
    const start = new Date(dStart.getFullYear(), dStart.getMonth(), 11);
    const dEnd = parseMonthKey(addMonthsToMonthKey(billingMonthKey, -1));
    const end = new Date(dEnd.getFullYear(), dEnd.getMonth(), 10);
    return { start, end };
  }
  if (norm === MITSUI_CATEGORY || norm === "dカード") {
    const dStart = parseMonthKey(addMonthsToMonthKey(billingMonthKey, -2));
    const start = new Date(dStart.getFullYear(), dStart.getMonth(), 16);
    const dEnd = parseMonthKey(addMonthsToMonthKey(billingMonthKey, -1));
    const end = new Date(dEnd.getFullYear(), dEnd.getMonth(), 15);
    return { start, end };
  }
  if (norm === WATER_CATEGORY) {
    const start = parseMonthKey(billingMonthKey);
    const end = new Date(start.getFullYear(), start.getMonth() + 1, 0);
    return { start, end };
  }
  const start = parseMonthKey(billingMonthKey);
  const end = new Date(start.getFullYear(), start.getMonth() + 1, 0);
  return { start, end };
}

/**
 * 支出を「トップ・内訳の表示月」に載せるための月キー（引き落とし月ベース）。
 * PayPay / セゾン / 三井住友 / dカード は画像の締日・引き落とし周期に合わせて利用日から算出する。
 * @param {Entry} entry
 * @returns {string} YYYY-MM
 */
function getExpenseBillingMonthKey(entry) {
  if (entry.kind !== "expense") return entry.date.slice(0, 7);
  const norm = normalizeExpenseCategory(entry.category);
  const parts = entry.date.split("-").map(Number);
  if (parts.length !== 3 || parts.some((n) => !Number.isFinite(n))) return entry.date.slice(0, 7);
  const [y, mo, day] = parts;
  const usageMonthKey = `${y}-${String(mo).padStart(2, "0")}`;

  /* PayPay（食費）: 当カレンダー月の利用 → 翌月27日引き落とし → 表示は「翌月」 */
  if (norm === FOOD_CATEGORY) {
    return addMonthsToMonthKey(usageMonthKey, 1);
  }

  /* セゾン: 10日締。1〜10日は当10日までの期間→翌月4日払い側の月が+1、11日以降は+2（例: 3/5→4月払い、3/15→5月払い） */
  if (norm === SAISON_CATEGORY) {
    if (day <= 10) return addMonthsToMonthKey(usageMonthKey, 1);
    return addMonthsToMonthKey(usageMonthKey, 2);
  }

  /* 三井住友・dカード: 15日締。1〜15日は+1、16日以降は+2（例: 3/5→4月10日払い、3/20→5月10日払い） */
  if (norm === MITSUI_CATEGORY || norm === "dカード") {
    if (day <= 15) return addMonthsToMonthKey(usageMonthKey, 1);
    return addMonthsToMonthKey(usageMonthKey, 2);
  }

  return usageMonthKey;
}

/** @param {number} n */
function formatYen(n) {
  return `¥${n.toLocaleString("ja-JP")}`;
}

/** カテゴリ名を表示・集計用に正規化 */
function normalizeExpenseCategory(category) {
  if (typeof category !== "string") return "";
  const t = category.trim();
  if (!t.length) return "";
  if (t === "PayPay") return "PayPay（食費）";
  if (t === "水道（奇数月に引き落とし）") return WATER_CATEGORY;
  return t;
}

/** @param {string} categoryName */
function variableBreakdownRemarkFor(categoryName) {
  const key = normalizeExpenseCategory(categoryName);
  return VARIABLE_BREAKDOWN_REMARKS[key] ?? "";
}

/** @param {Entry[]} entries @param {string} monthKey @param {Kind|'all'} filterKind */
function filterEntries(entries, monthKey, filterKind) {
  return entries.filter((e) => {
    if (e.kind === "income") {
      if (!e.date.startsWith(monthKey)) return false;
    } else {
      if (getExpenseBillingMonthKey(e) !== monthKey) return false;
    }
    if (filterKind === "all") return true;
    return e.kind === filterKind;
  });
}

/** 登録取引のみ集計（固定支出は含まない） @param {Entry[]} monthEntries */
function summarize(monthEntries) {
  let income = 0;
  let variableExpense = 0;
  for (const e of monthEntries) {
    if (e.kind === "income") income += e.amount;
    else variableExpense += e.amount;
  }
  return { income, variableExpense };
}

// --- DOM ---

const el = {
  sumSavingsRate: document.getElementById("sum-savings-rate"),
  sumIncome: document.getElementById("sum-income"),
  sumVariableExpense: document.getElementById("sum-variable-expense"),
  sumFixedExpense: document.getElementById("sum-fixed-expense"),
  sumTopFixedPlusVariable: document.getElementById("sum-top-fixed-plus-variable"),
  sumBalance: document.getElementById("sum-balance"),
  combinedLineFixed: document.getElementById("combined-line-fixed"),
  combinedLineVariable: document.getElementById("combined-line-variable"),
  sumFixedPlusVariable: document.getElementById("sum-fixed-plus-variable"),
  variableExpenseByCategory: document.getElementById("variable-expense-by-category"),
  variableBreakdownEmpty: document.getElementById("variable-breakdown-empty"),
  variableExpenseMonthTotal: document.getElementById("variable-expense-month-total"),
  projectIncomeList: document.getElementById("project-income-list"),
  projectIncomeEmpty: document.getElementById("project-income-empty"),
  projectIncomeMonthTotal: document.getElementById("project-income-month-total"),
  projectIncomeInputBreakdown: document.getElementById("project-income-input-breakdown"),
  projectIncomeInputBreakdownEmpty: document.getElementById("project-income-input-breakdown-empty"),
  formProjectIncome: document.getElementById("form-project-income"),
  projectRecordingDate: document.getElementById("project-recording-date"),
  projectIncomeName: document.getElementById("project-income-name"),
  projectIncomeAmount: document.getElementById("project-income-amount"),
  projectSumMonth: document.getElementById("project-sum-month"),
  fixedExpenseList: document.getElementById("fixed-expense-list"),
  fixedExpenseTotalLine: document.getElementById("fixed-expense-total-line"),
  fixedExpenseInputList: document.getElementById("fixed-expense-input-list"),
  fixedExpenseInputTotalLine: document.getElementById("fixed-expense-input-total-line"),
  formFixedExpenseInput: document.getElementById("form-fixed-expense-input"),
  fixedExpenseInputName: document.getElementById("fixed-expense-input-name"),
  fixedExpenseInputAmount: document.getElementById("fixed-expense-input-amount"),
  dialogFixedExpense: document.getElementById("dialog-fixed-expense"),
  formFixedExpense: document.getElementById("form-fixed-expense"),
  fixedExpenseEditId: document.getElementById("fixed-expense-edit-id"),
  fixedExpenseName: document.getElementById("fixed-expense-name"),
  fixedExpenseAmount: document.getElementById("fixed-expense-amount"),
  btnFixedExpenseCancel: document.getElementById("btn-fixed-expense-cancel"),
  btnFixedExpenseSave: document.getElementById("btn-fixed-expense-save"),
  monthLabel: document.getElementById("current-month-label"),
  monthToolbarLabelWrap: document.getElementById("month-toolbar-label-wrap"),
  btnMonthPicker: document.getElementById("btn-month-picker"),
  btnPrev: document.getElementById("btn-prev-month"),
  btnNext: document.getElementById("btn-next-month"),
  btnThisMonth: document.getElementById("btn-this-month"),
  form: document.getElementById("form-entry"),
  fieldDate: document.getElementById("field-date"),
  fieldAmount: document.getElementById("field-amount"),
  fieldCategory: document.getElementById("field-category"),
  fieldNote: document.getElementById("field-note"),
  filterKind: document.getElementById("filter-kind"),
  tbody: document.getElementById("entries-body"),
  empty: document.getElementById("empty-state"),
  btnExport: document.getElementById("btn-export"),
  btnExportGdrive: document.getElementById("btn-export-gdrive"),
  btnGdriveSettings: document.getElementById("btn-gdrive-settings"),
  dialogGdriveSettings: document.getElementById("dialog-gdrive-settings"),
  gdriveClientId: document.getElementById("gdrive-client-id"),
  gdriveFolderId: document.getElementById("gdrive-folder-id"),
  btnGdriveSettingsSave: document.getElementById("btn-gdrive-settings-save"),
  btnGdriveSettingsCancel: document.getElementById("btn-gdrive-settings-cancel"),
  btnPwaGuide: document.getElementById("btn-pwa-guide"),
  dialogPwaGuide: document.getElementById("dialog-pwa-guide"),
  btnPwaGuideClose: document.getElementById("btn-pwa-guide-close"),
  inputImport: document.getElementById("input-import"),
  formFood: document.getElementById("form-food"),
  foodRecordingDate: document.getElementById("food-recording-date"),
  foodAmount: document.getElementById("food-amount"),
  foodNote: document.getElementById("food-note"),
  foodSumToday: document.getElementById("food-sum-today"),
  foodSumMonth: document.getElementById("food-sum-month"),
  formSaison: document.getElementById("form-saison"),
  saisonRecordingDate: document.getElementById("saison-recording-date"),
  saisonAmount: document.getElementById("saison-amount"),
  saisonNote: document.getElementById("saison-note"),
  saisonSumToday: document.getElementById("saison-sum-today"),
  saisonSumMonth: document.getElementById("saison-sum-month"),
  formMitsui: document.getElementById("form-mitsui"),
  mitsuiRecordingDate: document.getElementById("mitsui-recording-date"),
  mitsuiAmount: document.getElementById("mitsui-amount"),
  mitsuiNote: document.getElementById("mitsui-note"),
  mitsuiSumToday: document.getElementById("mitsui-sum-today"),
  mitsuiSumMonth: document.getElementById("mitsui-sum-month"),
  formVariablePay: document.getElementById("form-variable-pay"),
  variablePayRecordingDate: document.getElementById("variable-pay-recording-date"),
  variablePayItem: document.getElementById("variable-pay-item"),
  variablePayAmount: document.getElementById("variable-pay-amount"),
  variablePayNote: document.getElementById("variable-pay-note"),
  foodUsagePeriod: document.getElementById("food-usage-period"),
  saisonUsagePeriod: document.getElementById("saison-usage-period"),
  mitsuiUsagePeriod: document.getElementById("mitsui-usage-period"),
  projectUsagePeriod: document.getElementById("project-usage-period"),
  variablePayUsagePeriod: document.getElementById("variable-pay-usage-period"),
  tradeUsagePeriod: document.getElementById("trade-usage-period"),
  editTargetBannerDesc: document.getElementById("edit-target-banner-desc"),
  editDate: document.getElementById("edit-date"),
  editAmount: document.getElementById("edit-amount"),
  editNote: document.getElementById("edit-note"),
  dialogBreakdownPick: document.getElementById("dialog-breakdown-pick"),
  breakdownPickStepList: document.getElementById("breakdown-pick-step-list"),
  breakdownPickStepEdit: document.getElementById("breakdown-pick-step-edit"),
  breakdownPickTitle: document.getElementById("breakdown-pick-title"),
  breakdownPickDesc: document.getElementById("breakdown-pick-desc"),
  breakdownPickList: document.getElementById("breakdown-pick-list"),
  breakdownPickEmpty: document.getElementById("breakdown-pick-empty"),
  btnBreakdownPickClose: document.getElementById("btn-breakdown-pick-close"),
  formInlineEdit: document.getElementById("form-inline-edit-entry"),
  inlineEditEntryId: document.getElementById("inline-edit-entry-id"),
  btnInlineEditBack: document.getElementById("btn-inline-edit-back"),
  btnInlineEditClose: document.getElementById("btn-inline-edit-close"),
  btnInlineEditSave: document.getElementById("btn-inline-edit-save"),
  formExpenseRecord: document.getElementById("form-expense-record"),
  expenseRecordDate: document.getElementById("expense-record-date"),
  expenseRecordMemo: document.getElementById("expense-record-memo"),
  expenseRecordPhotoBrowse: document.getElementById("expense-record-photo-browse"),
  expenseRecordPhotoCamera: document.getElementById("expense-record-photo-camera"),
  expenseRecordPhotoRoll: document.getElementById("expense-record-photo-roll"),
  expensePhotoPickStatus: document.getElementById("expense-photo-pick-status"),
  btnExpenseRecordSave: document.getElementById("btn-expense-record-save"),
  expenseRecordList: document.getElementById("expense-record-list"),
  expenseRecordEmpty: document.getElementById("expense-record-empty"),
};

/** @type {{ version: number, entries: Entry[], fixedExpenses: FixedExpenseItem[], expenseRecords: ExpenseRecord[] }} */
let state = loadState();

/** 表示中の月 YYYY-MM */
let viewMonthKey = monthKeyFromDate(new Date());

/** 取引選択ダイアログを削除後に再描画するための取得関数 */
let breakdownPickListRefetch = /** @type {null | (() => Entry[])} */ (null);

/** true のとき「一覧に戻る」で内訳一覧ステップへ戻る（取引一覧の編集から開いたときは false） */
let breakdownPickEditReturnToList = false;

function getSelectedKind() {
  const checked = document.querySelector('input[name="kind"]:checked');
  return checked && checked.value === "expense" ? "expense" : "income";
}

function refreshCategoryOptions() {
  const kind = getSelectedKind();
  const list = kind === "income" ? INCOME_CATEGORIES : EXPENSE_CATEGORIES;
  el.fieldCategory.innerHTML = list.map((c) => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join("");
}

function getEditKind() {
  const checked = document.querySelector('input[name="edit-kind"]:checked');
  return checked && checked.value === "expense" ? "expense" : "income";
}

/** @param {Entry} a @param {Entry} b */
function sortEntriesForEditSelect(a, b) {
  if (a.date !== b.date) return b.date.localeCompare(a.date);
  return b.id.localeCompare(a.id);
}

/** グループ化に失敗したときのフラット表示用 */
function formatEditFlatOptionLabel(entry) {
  const kindLabel = entry.kind === "income" ? "収入" : "支出";
  const cat = normalizeExpenseCategory(entry.category) || entry.category || "—";
  const note = (entry.note || "").trim();
  const noteShort = note.length > 26 ? `${note.slice(0, 26)}…` : note;
  const memoPart = noteShort ? ` · ${noteShort}` : "";
  return `${entry.date} ${kindLabel} ${cat} ${formatYen(entry.amount)}${memoPart}`;
}

function resetEditFormFields() {
  if (el.inlineEditEntryId) el.inlineEditEntryId.value = "";
  if (el.editDate) el.editDate.value = "";
  if (el.editAmount) el.editAmount.value = "";
  if (el.editNote) el.editNote.value = "";
  document.querySelectorAll('input[name="edit-kind"]').forEach((r) => {
    if (r instanceof HTMLInputElement) r.checked = r.value === "expense";
  });
  if (el.editTargetBannerDesc) el.editTargetBannerDesc.innerHTML = "";
}

/** @param {Entry} entry */
function buildEditTargetSummaryHtml(entry) {
  const kindLabel = entry.kind === "income" ? "収入" : "支出";
  const cat = escapeHtml(normalizeExpenseCategory(entry.category) || entry.category || "—");
  const note = (entry.note || "").trim();
  const noteBlock = note
    ? `<span class="edit-dialog__target-note">メモ: ${escapeHtml(note)}</span>`
    : `<span class="edit-dialog__target-note edit-dialog__target-note--empty">メモなし</span>`;
  return `<span class="edit-dialog__target-line">${escapeHtml(formatDateLongJa(entry.date))} · <strong>${kindLabel}</strong> · ${formatYen(entry.amount)} · ${cat}</span>${noteBlock}`;
}

/** @param {Entry} entry */
function loadEditFormFromEntry(entry) {
  if (el.editDate) el.editDate.value = entry.date;
  if (el.editAmount) el.editAmount.value = String(entry.amount);
  if (el.editNote) el.editNote.value = entry.note || "";
  document.querySelectorAll('input[name="edit-kind"]').forEach((r) => {
    if (r instanceof HTMLInputElement) r.checked = r.value === entry.kind;
  });
  if (el.editTargetBannerDesc) el.editTargetBannerDesc.innerHTML = buildEditTargetSummaryHtml(entry);
}

/** 一覧ステップを表示し、編集ステップを隠す */
function resetBreakdownPickToListStep() {
  if (el.breakdownPickStepEdit) el.breakdownPickStepEdit.hidden = true;
  if (el.breakdownPickStepList) el.breakdownPickStepList.hidden = false;
  if (el.btnInlineEditBack) el.btnInlineEditBack.hidden = true;
  breakdownPickEditReturnToList = false;
  resetEditFormFields();
  if (breakdownPickListRefetch) {
    renderBreakdownPickListBody(breakdownPickListRefetch());
  }
}

/**
 * @param {string} entryId
 * @param {boolean} returnToList true のとき「一覧に戻る」を表示（内訳から開いた取引一覧）
 */
function showInlineEditForEntry(entryId, returnToList) {
  const entry = state.entries.find((e) => e.id === entryId);
  if (!entry || !el.dialogBreakdownPick) return;
  breakdownPickEditReturnToList = returnToList;
  if (el.inlineEditEntryId) el.inlineEditEntryId.value = entryId;
  loadEditFormFromEntry(entry);
  if (el.breakdownPickStepList) el.breakdownPickStepList.hidden = true;
  if (el.breakdownPickStepEdit) el.breakdownPickStepEdit.hidden = false;
  if (el.btnInlineEditBack) el.btnInlineEditBack.hidden = !returnToList;
  if (!el.dialogBreakdownPick.open) el.dialogBreakdownPick.showModal();
  queueMicrotask(() => el.editAmount?.focus());
}

/** @param {string} entryId 取引一覧の「編集」など */
function openEditEntryDialog(entryId) {
  breakdownPickListRefetch = null;
  showInlineEditForEntry(entryId, false);
}

/** @param {Entry[]} entries */
function buildVariableBreakdownDetailInnerHtml(entries) {
  if (!entries.length) {
    return `<p class="variable-breakdown-list__detail-empty">この月（引き落とし月）に該当する取引はありません。</p>`;
  }
  const sorted = [...entries].sort(sortEntriesForEditSelect);
  return `<ul class="variable-breakdown-list__detail-list">
    ${sorted
      .map((e) => {
        const memo = (e.note || "").trim();
        const memoHtml = memo
          ? `<span class="variable-breakdown-list__detail-memo">${escapeHtml(memo)}</span>`
          : `<span class="variable-breakdown-list__detail-memo variable-breakdown-list__detail-memo--empty">（メモなし）</span>`;
        return `<li class="variable-breakdown-list__detail-item">
        <span class="variable-breakdown-list__detail-date">${escapeHtml(formatDateLongJa(e.date))}</span>
        <span class="variable-breakdown-list__detail-amount">${formatYen(e.amount)}</span>
        ${memoHtml}
      </li>`;
      })
      .join("")}
  </ul>`;
}

/** 変動支出（内訳）の行名に一致する、表示月（引き落とし月）に載る支出取引 */
function getExpenseEntriesForVariableRowName(rowName) {
  const monthAll = state.entries.filter((e) => {
    if (e.kind !== "expense") return false;
    return getExpenseBillingMonthKey(e) === viewMonthKey;
  });
  const targetNorm = normalizeExpenseCategory(rowName) || rowName;
  return monthAll
    .filter((e) => {
      if (e.kind !== "expense") return false;
      const n = normalizeExpenseCategory(e.category) || "（未分類）";
      return n === targetNorm;
    })
    .sort(sortEntriesForEditSelect);
}

/** 収入（案件）の表示名（集計キー）に一致する当月の案件収入 */
function getProjectIncomeEntriesForDisplayName(displayName) {
  const monthAll = state.entries.filter((e) => e.date.startsWith(viewMonthKey));
  return monthAll
    .filter(
      (e) =>
        e.kind === "income" &&
        e.category === PROJECT_INCOME_CATEGORY &&
        ((e.note || "").trim() || "（無題）") === displayName
    )
    .sort(sortEntriesForEditSelect);
}

/** @param {Entry[]} entries */
function renderBreakdownPickListBody(entries) {
  if (!el.breakdownPickEmpty || !el.breakdownPickList) return;
  if (!entries.length) {
    el.breakdownPickList.innerHTML = "";
    el.breakdownPickEmpty.hidden = false;
  } else {
    el.breakdownPickEmpty.hidden = true;
    el.breakdownPickList.innerHTML = entries
      .map(
        (e) => `<li class="breakdown-pick-list__item">
        <span class="breakdown-pick-list__line">${escapeHtml(formatEditFlatOptionLabel(e))}</span>
        <div class="breakdown-pick-list__actions">
          <button type="button" class="btn btn--small btn--primary breakdown-pick-list__edit" data-pick-edit-id="${escapeHtml(e.id)}">修正</button>
          <button type="button" class="btn btn--small btn--danger breakdown-pick-list__delete" data-pick-delete-id="${escapeHtml(e.id)}">削除</button>
        </div>
      </li>`
      )
      .join("");
  }
}

/** @param {string} title @param {string} desc @param {Entry[]} entries @param {(() => Entry[]) | null} [refetch] */
function showTransactionPickDialog(title, desc, entries, refetch = null) {
  if (!el.dialogBreakdownPick || !el.breakdownPickTitle) return;
  breakdownPickListRefetch = refetch;
  if (el.breakdownPickStepList) el.breakdownPickStepList.hidden = false;
  if (el.breakdownPickStepEdit) el.breakdownPickStepEdit.hidden = true;
  if (el.btnInlineEditBack) el.btnInlineEditBack.hidden = true;
  resetEditFormFields();
  el.breakdownPickTitle.textContent = title;
  if (el.breakdownPickDesc) {
    el.breakdownPickDesc.textContent = desc;
    el.breakdownPickDesc.hidden = !desc;
  }
  renderBreakdownPickListBody(entries);
  el.dialogBreakdownPick.showModal();
}

/** @param {string} encodedRowName encodeURIComponent 済みの内訳行名 */
function openVariableBreakdownRowForEdit(encodedRowName) {
  const rowName = decodeURIComponent(encodedRowName);
  const refetch = () => getExpenseEntriesForVariableRowName(rowName);
  showTransactionPickDialog(`「${rowName}」の取引を修正`, "修正・削除する取引を選んでください。", refetch(), refetch);
}

/** @param {string} encodedProjectName encodeURIComponent 済みの案件表示名 */
function openProjectIncomeRowForEdit(encodedProjectName) {
  const name = decodeURIComponent(encodedProjectName);
  const refetch = () => getProjectIncomeEntriesForDisplayName(name);
  showTransactionPickDialog(`「${name}」の案件収入を修正`, "修正・削除する取引を選んでください。", refetch(), refetch);
}

function initVariablePaySelect() {
  if (!el.variablePayItem) return;
  el.variablePayItem.innerHTML = VARIABLE_PAYMENT_LABELS.map(
    (c) => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`
  ).join("");
  if (!el.variablePayItem.dataset.usagePeriodListener) {
    el.variablePayItem.dataset.usagePeriodListener = "1";
    el.variablePayItem.addEventListener("change", () => {
      updateVariablePayUsagePeriodLine(new Date());
    });
  }
}

/** @param {string} s */
function escapeHtml(s) {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function isoDateFromDate(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

/**
 * カード種別ごとの利用明細期間（締日ルールと getExpenseBillingMonthKey の前提と一致）。
 * @param {string} norm normalizeExpenseCategory 済みのカテゴリ名
 * @param {Date} refDate 基準日（クイック入力は本日、取引追加は日付欄）
 */
function getUsagePeriodBoundsForCategoryNorm(norm, refDate) {
  const y = refDate.getFullYear();
  const m = refDate.getMonth();
  const d = refDate.getDate();

  if (norm === FOOD_CATEGORY) {
    return { start: new Date(y, m, 1), end: new Date(y, m + 1, 0) };
  }
  if (norm === SAISON_CATEGORY) {
    if (d <= 10) return { start: new Date(y, m - 1, 11), end: new Date(y, m, 10) };
    return { start: new Date(y, m, 11), end: new Date(y, m + 1, 10) };
  }
  if (norm === MITSUI_CATEGORY || norm === "dカード") {
    if (d <= 15) return { start: new Date(y, m - 1, 16), end: new Date(y, m, 15) };
    return { start: new Date(y, m, 16), end: new Date(y, m + 1, 15) };
  }
  return { start: new Date(y, m, 1), end: new Date(y, m + 1, 0) };
}

/** @param {Date} start @param {Date} end */
function formatUsagePeriodRangeJa(start, end) {
  return `${formatDateLongJa(isoDateFromDate(start))}〜${formatDateLongJa(isoDateFromDate(end))}`;
}

/** 食費・セゾン・三井・変動（項目別）・取引追加の利用期間表示を更新（本日の日付で再計算） */
function updateQuickInputUsagePeriodLabels(today = new Date()) {
  if (el.foodUsagePeriod) {
    const { start, end } = getUsagePeriodBoundsForCategoryNorm(FOOD_CATEGORY, today);
    el.foodUsagePeriod.textContent = `利用期間（今日基準・自動更新）: ${formatUsagePeriodRangeJa(start, end)}（当月カレンダー・月末〆）`;
  }
  if (el.saisonUsagePeriod) {
    const { start, end } = getUsagePeriodBoundsForCategoryNorm(SAISON_CATEGORY, today);
    el.saisonUsagePeriod.textContent = `利用期間（今日基準・自動更新）: ${formatUsagePeriodRangeJa(start, end)}（10日締の対象期間）`;
  }
  if (el.mitsuiUsagePeriod) {
    const { start, end } = getUsagePeriodBoundsForCategoryNorm(MITSUI_CATEGORY, today);
    el.mitsuiUsagePeriod.textContent = `利用期間（今日基準・自動更新）: ${formatUsagePeriodRangeJa(start, end)}（15日締の対象期間）`;
  }
  updateVariablePayUsagePeriodLine(today);
  updateTradeUsagePeriodLine();
}

/** @param {Date} [today] */
function updateVariablePayUsagePeriodLine(today = new Date()) {
  if (!el.variablePayUsagePeriod) return;
  const cat = el.variablePayItem?.value ?? "";
  if (cat === "dカード") {
    const { start, end } = getUsagePeriodBoundsForCategoryNorm("dカード", today);
    el.variablePayUsagePeriod.textContent = `利用期間（今日基準・自動更新）: ${formatUsagePeriodRangeJa(start, end)}（15日締・三井住友カードと同じ周期）`;
  } else if (cat === WATER_CATEGORY) {
    const mk = monthKeyFromDate(today);
    el.variablePayUsagePeriod.textContent = `計上月（記録日の暦月・自動更新）: ${formatMonthLongJaFromKey(mk)}（内訳は日付の月で集計。水道は奇数月引き落とし）`;
  } else {
    el.variablePayUsagePeriod.textContent = "";
  }
}

function refDateFromFieldDateOrToday() {
  const dateStr = el.fieldDate?.value;
  if (dateStr && /^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
    const [yy, mm, dd] = dateStr.split("-").map(Number);
    return new Date(yy, mm - 1, dd);
  }
  return new Date();
}

function updateTradeUsagePeriodLine() {
  if (!el.tradeUsagePeriod) return;
  if (getSelectedKind() !== "expense") {
    el.tradeUsagePeriod.textContent = `計上月: 上記の日付の属する暦月（収入はその月の集計に含まれます）`;
    return;
  }
  const norm = normalizeExpenseCategory(el.fieldCategory?.value ?? "");
  const ref = refDateFromFieldDateOrToday();
  const { start, end } = getUsagePeriodBoundsForCategoryNorm(norm, ref);
  const range = formatUsagePeriodRangeJa(start, end);
  if (norm === FOOD_CATEGORY) {
    el.tradeUsagePeriod.textContent = `利用期間（日付・カテゴリに連動）: ${range}（月末〆・翌27日引きの対象区間）`;
  } else if (norm === SAISON_CATEGORY) {
    el.tradeUsagePeriod.textContent = `利用期間（日付・カテゴリに連動）: ${range}（10日締・翌月4日引きの対象区間）`;
  } else if (norm === MITSUI_CATEGORY || norm === "dカード") {
    el.tradeUsagePeriod.textContent = `利用期間（日付・カテゴリに連動）: ${range}（15日締・翌月10日引きの対象区間）`;
  } else {
    el.tradeUsagePeriod.textContent = `利用の目安（日付の暦月）: ${range}（引き落とし月は内訳ルールに従い自動計算）`;
  }
}

const DATE_PICKER_WEEKDAYS = ["日", "月", "火", "水", "木", "金", "土"];

/** @type {{ input: HTMLInputElement, anchor: HTMLElement, y: number, m: number, onDoc: ((e: MouseEvent) => void) | null, onKey: ((e: KeyboardEvent) => void) | null } | null} */
let datePickerState = null;

/** @type {{ y: number, anchor: HTMLElement, onDoc: ((e: MouseEvent) => void) | null, onKey: ((e: KeyboardEvent) => void) | null } | null} */
let monthPickerState = null;

/** @param {number} y @param {number} month 1-12 @param {number} day */
function isoFromYmd(y, month, day) {
  return `${y}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function closeDatePicker() {
  const pop = document.getElementById("date-picker-popover");
  if (pop) pop.hidden = true;
  if (datePickerState?.onDoc) document.removeEventListener("mousedown", datePickerState.onDoc, true);
  if (datePickerState?.onKey) document.removeEventListener("keydown", datePickerState.onKey, true);
  datePickerState = null;
}

function closeMonthPicker() {
  const pop = document.getElementById("month-picker-popover");
  if (pop) pop.hidden = true;
  if (monthPickerState?.onDoc) document.removeEventListener("mousedown", monthPickerState.onDoc, true);
  if (monthPickerState?.onKey) document.removeEventListener("keydown", monthPickerState.onKey, true);
  monthPickerState = null;
}

/** @param {HTMLElement} pop @param {HTMLElement} anchor */
function positionDatePickerPopover(pop, anchor) {
  positionPopoverNearAnchor(pop, anchor, false);
}

/** @param {HTMLElement} pop @param {HTMLElement} anchor @param {boolean} [centerUnder] */
function positionPopoverNearAnchor(pop, anchor, centerUnder) {
  const r = anchor.getBoundingClientRect();
  const margin = 8;
  pop.style.position = "fixed";
  requestAnimationFrame(() => {
    const ph = pop.offsetHeight || 300;
    const pw = pop.offsetWidth || 280;
    let left = centerUnder ? r.left + (r.width - pw) / 2 : r.left;
    if (left + pw > window.innerWidth - margin) left = window.innerWidth - pw - margin;
    left = Math.max(margin, left);
    let top = r.bottom + 6;
    if (top + ph > window.innerHeight - margin) top = r.top - ph - 6;
    top = Math.max(margin, top);
    pop.style.left = `${left}px`;
    pop.style.top = `${top}px`;
  });
}

/** @param {HTMLElement} pop @param {HTMLElement} anchor */
function positionMonthPickerPopover(pop, anchor) {
  positionPopoverNearAnchor(pop, anchor, true);
}

/** @param {HTMLElement} pop */
function renderDatePickerDays(pop) {
  if (!datePickerState) return;
  const { y, m, input } = datePickerState;
  const first = new Date(y, m - 1, 1);
  const lastDay = new Date(y, m, 0).getDate();
  const startPad = first.getDay();
  const todayIso = isoDateFromDate(new Date());
  const selectedIso = input.value && /^\d{4}-\d{2}-\d{2}$/.test(input.value) ? input.value : "";

  let grid = "";
  for (const wd of DATE_PICKER_WEEKDAYS) {
    grid += `<div class="date-picker-popover__dow">${wd}</div>`;
  }
  for (let i = 0; i < startPad; i++) {
    grid += `<div class="date-picker-popover__cell--empty" aria-hidden="true"></div>`;
  }
  for (let d = 1; d <= lastDay; d++) {
    const iso = isoFromYmd(y, m, d);
    let cls = "date-picker-popover__day";
    if (iso === todayIso) cls += " date-picker-popover__day--today";
    if (iso === selectedIso) cls += " date-picker-popover__day--selected";
    grid += `<button type="button" class="${cls}" data-iso="${iso}">${d}</button>`;
  }

  const title = `${y}年${m}月`;
  pop.innerHTML = `<div class="date-picker-popover__header">
    <button type="button" class="date-picker-popover__nav" aria-label="前の月" data-dp-nav="prev">‹</button>
    <div class="date-picker-popover__title">${escapeHtml(title)}</div>
    <button type="button" class="date-picker-popover__nav" aria-label="次の月" data-dp-nav="next">›</button>
  </div>
  <div class="date-picker-popover__grid">${grid}</div>`;
}

/** @param {number} delta */
function shiftDatePickerMonth(delta) {
  if (!datePickerState) return;
  let y = datePickerState.y;
  let mo = datePickerState.m + delta;
  if (mo > 12) {
    mo = 1;
    y++;
  }
  if (mo < 1) {
    mo = 12;
    y--;
  }
  datePickerState.y = y;
  datePickerState.m = mo;
  const pop = document.getElementById("date-picker-popover");
  if (pop) {
    renderDatePickerDays(pop);
    positionDatePickerPopover(pop, datePickerState.anchor);
  }
}

function ensureDatePickerPopoverMounted() {
  let pop = document.getElementById("date-picker-popover");
  if (pop) return pop;
  pop = document.createElement("div");
  pop.id = "date-picker-popover";
  pop.className = "date-picker-popover";
  pop.hidden = true;
  pop.setAttribute("role", "dialog");
  pop.setAttribute("aria-label", "日付を選ぶ");
  document.body.appendChild(pop);
  pop.addEventListener("click", (ev) => {
    const t = ev.target;
    if (!(t instanceof HTMLElement)) return;
    const navEl = t.closest("[data-dp-nav]");
    if (navEl instanceof HTMLElement) {
      const dir = navEl.getAttribute("data-dp-nav");
      shiftDatePickerMonth(dir === "next" ? 1 : -1);
      return;
    }
    const dayBtn = t.closest(".date-picker-popover__day[data-iso]");
    if (dayBtn instanceof HTMLElement && datePickerState) {
      const iso = dayBtn.getAttribute("data-iso");
      if (iso) {
        datePickerState.input.value = iso;
        datePickerState.input.dispatchEvent(new Event("input", { bubbles: true }));
        datePickerState.input.dispatchEvent(new Event("change", { bubbles: true }));
      }
      closeDatePicker();
    }
  });
  return pop;
}

/** @param {HTMLElement} pop */
function renderMonthPickerMonths(pop) {
  if (!monthPickerState) return;
  const y = monthPickerState.y;
  const now = new Date();
  const cy = now.getFullYear();
  const cm = now.getMonth() + 1;
  const selParts = viewMonthKey.split("-").map(Number);
  const selY = selParts[0];
  const selM = selParts[1];

  let months = "";
  for (let mo = 1; mo <= 12; mo++) {
    let cls = "month-picker-popover__month";
    if (y === cy && mo === cm) cls += " month-picker-popover__month--this";
    if (y === selY && mo === selM) cls += " month-picker-popover__month--selected";
    months += `<button type="button" class="${cls}" data-month="${mo}">${mo}月</button>`;
  }

  pop.innerHTML = `<div class="month-picker-popover__header">
    <button type="button" class="month-picker-popover__nav" aria-label="前年" data-mp-year="prev">‹</button>
    <div class="month-picker-popover__title">${escapeHtml(String(y))}年</div>
    <button type="button" class="month-picker-popover__nav" aria-label="翌年" data-mp-year="next">›</button>
  </div>
  <div class="month-picker-popover__months">${months}</div>`;
}

/** @param {number} delta */
function shiftMonthPickerYear(delta) {
  if (!monthPickerState) return;
  monthPickerState.y += delta;
  const pop = document.getElementById("month-picker-popover");
  if (pop) {
    renderMonthPickerMonths(pop);
    positionMonthPickerPopover(pop, monthPickerState.anchor);
  }
}

function ensureMonthPickerPopoverMounted() {
  let pop = document.getElementById("month-picker-popover");
  if (pop) return pop;
  pop = document.createElement("div");
  pop.id = "month-picker-popover";
  pop.className = "month-picker-popover";
  pop.hidden = true;
  pop.setAttribute("role", "dialog");
  pop.setAttribute("aria-label", "表示する月を選ぶ");
  document.body.appendChild(pop);
  pop.addEventListener("click", (ev) => {
    const t = ev.target;
    if (!(t instanceof HTMLElement)) return;
    const navEl = t.closest("[data-mp-year]");
    if (navEl instanceof HTMLElement && monthPickerState) {
      const dir = navEl.getAttribute("data-mp-year");
      shiftMonthPickerYear(dir === "next" ? 1 : -1);
      return;
    }
    const mBtn = t.closest(".month-picker-popover__month[data-month]");
    if (mBtn instanceof HTMLElement && monthPickerState) {
      const m = Number(mBtn.getAttribute("data-month"));
      if (m >= 1 && m <= 12) {
        viewMonthKey = `${monthPickerState.y}-${String(m).padStart(2, "0")}`;
        closeMonthPicker();
        render();
      }
    }
  });
  return pop;
}

function openMonthPicker() {
  closeDatePicker();
  closeMonthPicker();
  const anchor = el.monthToolbarLabelWrap || el.btnMonthPicker;
  if (!anchor) return;
  const parts = viewMonthKey.split("-").map(Number);
  const y = parts[0];
  if (!Number.isFinite(y)) return;
  monthPickerState = { y, anchor, onDoc: null, onKey: null };
  const pop = ensureMonthPickerPopoverMounted();
  renderMonthPickerMonths(pop);
  pop.hidden = false;
  positionMonthPickerPopover(pop, anchor);
  monthPickerState.onDoc = (e) => {
    const t = e.target;
    if (!(t instanceof Node)) return;
    if (pop.contains(t) || anchor.contains(t)) return;
    closeMonthPicker();
  };
  monthPickerState.onKey = (e) => {
    if (e.key === "Escape") {
      e.preventDefault();
      closeMonthPicker();
    }
  };
  queueMicrotask(() => {
    if (!monthPickerState) return;
    document.addEventListener("mousedown", monthPickerState.onDoc, true);
    document.addEventListener("keydown", monthPickerState.onKey, true);
  });
}

/** @param {HTMLInputElement} input @param {HTMLElement} anchor */
function openDatePickerForInput(input, anchor) {
  closeDatePicker();
  closeMonthPicker();
  const pop = ensureDatePickerPopoverMounted();
  let y;
  let mo;
  if (input.value && /^\d{4}-\d{2}-\d{2}$/.test(input.value)) {
    const p = input.value.split("-").map(Number);
    y = p[0];
    mo = p[1];
  } else {
    const t = new Date();
    y = t.getFullYear();
    mo = t.getMonth() + 1;
  }
  datePickerState = { input, anchor, y, m: mo, onDoc: null, onKey: null };
  renderDatePickerDays(pop);
  pop.hidden = false;
  positionDatePickerPopover(pop, anchor);
  datePickerState.onDoc = (e) => {
    const t = e.target;
    if (!(t instanceof Node)) return;
    if (pop.contains(t) || anchor.contains(t)) return;
    closeDatePicker();
  };
  datePickerState.onKey = (e) => {
    if (e.key === "Escape") {
      e.preventDefault();
      closeDatePicker();
    }
  };
  queueMicrotask(() => {
    if (!datePickerState) return;
    document.addEventListener("mousedown", datePickerState.onDoc, true);
    document.addEventListener("keydown", datePickerState.onKey, true);
  });
}

function getPickedExpensePhotoFile() {
  return (
    el.expenseRecordPhotoCamera?.files?.[0] ??
    el.expenseRecordPhotoRoll?.files?.[0] ??
    el.expenseRecordPhotoBrowse?.files?.[0]
  );
}

function updateExpensePhotoPickStatus() {
  if (!el.expensePhotoPickStatus) return;
  const file = getPickedExpensePhotoFile();
  el.expensePhotoPickStatus.textContent = file ? `選択中: ${file.name || "画像 1 件"}` : "";
}

function clearExpensePhotoInputs() {
  if (el.expenseRecordPhotoBrowse) el.expenseRecordPhotoBrowse.value = "";
  if (el.expenseRecordPhotoCamera) el.expenseRecordPhotoCamera.value = "";
  if (el.expenseRecordPhotoRoll) el.expenseRecordPhotoRoll.value = "";
  updateExpensePhotoPickStatus();
}

function initExpensePhotoInputs() {
  [el.expenseRecordPhotoBrowse, el.expenseRecordPhotoCamera, el.expenseRecordPhotoRoll].forEach((inp) => {
    if (!inp) return;
    inp.addEventListener("change", () => {
      if (!inp.files?.length) {
        updateExpensePhotoPickStatus();
        return;
      }
      [el.expenseRecordPhotoBrowse, el.expenseRecordPhotoCamera, el.expenseRecordPhotoRoll].forEach((other) => {
        if (other && other !== inp) other.value = "";
      });
      updateExpensePhotoPickStatus();
    });
  });
}

function initCustomDatePickers() {
  document.querySelectorAll(".date-field").forEach((wrap) => {
    const input = wrap.querySelector(".date-field__input");
    const btn = wrap.querySelector(".date-field__calendar-btn");
    if (!(input instanceof HTMLInputElement) || !(btn instanceof HTMLButtonElement)) return;
    if (wrap.dataset.datePickerInit) return;
    wrap.dataset.datePickerInit = "1";
    const open = (e) => {
      e.preventDefault();
      openDatePickerForInput(input, wrap);
    };
    btn.addEventListener("click", open);
    input.addEventListener("click", open);
  });
}

/** @param {string} yyyymmdd */
function formatDateLongJa(yyyymmdd) {
  const parts = yyyymmdd.split("-");
  if (parts.length !== 3) return yyyymmdd;
  const y = Number(parts[0]);
  const m = Number(parts[1]);
  const d = Number(parts[2]);
  const dt = new Date(y, m - 1, d);
  return dt.toLocaleDateString("ja-JP", { year: "numeric", month: "long", day: "numeric", weekday: "short" });
}

function setDefaultDate() {
  el.fieldDate.value = isoDateFromDate(new Date());
}

/** @param {Entry[]} entries @param {string} category @param {string} yyyymmdd */
function sumExpenseCategoryDay(entries, category, yyyymmdd) {
  const norm = normalizeExpenseCategory(category);
  return entries
    .filter(
      (e) =>
        e.kind === "expense" &&
        normalizeExpenseCategory(e.category) === norm &&
        e.date === yyyymmdd
    )
    .reduce((s, e) => s + e.amount, 0);
}

/** @param {Entry[]} entries @param {string} category @param {string} monthKey 引き落とし月（表示月） */
function sumExpenseCategoryMonth(entries, category, monthKey) {
  const norm = normalizeExpenseCategory(category);
  return entries
    .filter(
      (e) =>
        e.kind === "expense" &&
        normalizeExpenseCategory(e.category) === norm &&
        getExpenseBillingMonthKey(e) === monthKey
    )
    .reduce((s, e) => s + e.amount, 0);
}

/** 既に表示月で絞った entries のスライス上でカテゴリ合計（内訳用） */
function sumExpenseCategoryInEntries(entries, category) {
  const norm = normalizeExpenseCategory(category);
  return entries
    .filter((e) => e.kind === "expense" && normalizeExpenseCategory(e.category) === norm)
    .reduce((s, e) => s + e.amount, 0);
}

/** @param {Entry[]} entries @param {string} category @param {string} monthKey */
function sumIncomeCategoryMonth(entries, category, monthKey) {
  return entries
    .filter((e) => e.kind === "income" && e.category === category && e.date.startsWith(monthKey))
    .reduce((s, e) => s + e.amount, 0);
}

/** 表示中の月の案件収入を案件名（note）で集計 @param {Entry[]} monthEntries */
function aggregateProjectIncomeByName(monthEntries) {
  const map = new Map();
  for (const e of monthEntries) {
    if (e.kind !== "income" || e.category !== PROJECT_INCOME_CATEGORY) continue;
    const name = (e.note || "").trim() || "（無題）";
    map.set(name, (map.get(name) ?? 0) + e.amount);
  }
  return [...map.entries()].sort((a, b) => {
    if (b[1] !== a[1]) return b[1] - a[1];
    return a[0].localeCompare(b[0], "ja");
  });
}

/** @param {Entry[]} monthEntries */
function renderProjectIncomeHome(monthEntries) {
  const pairs = aggregateProjectIncomeByName(monthEntries);
  const total = pairs.reduce((s, [, amt]) => s + amt, 0);
  if (el.projectIncomeList) {
    el.projectIncomeList.innerHTML = pairs
      .map(([name, amount]) => {
        const enc = escapeHtml(encodeURIComponent(name));
        const actionBtns =
          amount > 0
            ? `<div class="project-income-list__actions">
          <button type="button" class="btn btn--small btn--ghost project-income-list__edit" data-breakdown-project="${enc}" aria-label="この案件の取引を修正">修正</button>
          <button type="button" class="btn btn--small btn--danger project-income-list__delete" data-breakdown-project-delete="${enc}" aria-label="この案件の当月取引をすべて削除">削除</button>
        </div>`
            : "";
        return `<li class="fixed-list__item project-income-list__item">
        <span class="fixed-list__name">${escapeHtml(name)}</span>
        <div class="fixed-list__item-trailing">
          <span class="fixed-list__amount fixed-list__amount--income">${formatYen(amount)}</span>${actionBtns}
        </div>
      </li>`;
      })
      .join("");
  }
  if (el.projectIncomeEmpty) el.projectIncomeEmpty.hidden = pairs.length > 0;
  if (el.projectIncomeMonthTotal) el.projectIncomeMonthTotal.textContent = formatYen(total);
}

/** @param {Entry[]} monthEntries */
function renderProjectIncomeInputBreakdown(monthEntries) {
  const pairs = aggregateProjectIncomeByName(monthEntries);
  if (el.projectIncomeInputBreakdown) {
    el.projectIncomeInputBreakdown.innerHTML = pairs
      .map(([name, amount]) => {
        const enc = escapeHtml(encodeURIComponent(name));
        const actionBtns =
          amount > 0
            ? `<div class="project-income-list__actions">
          <button type="button" class="btn btn--small btn--ghost project-income-list__edit" data-breakdown-project="${enc}" aria-label="この案件の取引を修正">修正</button>
          <button type="button" class="btn btn--small btn--danger project-income-list__delete" data-breakdown-project-delete="${enc}" aria-label="この案件の当月取引をすべて削除">削除</button>
        </div>`
            : "";
        return `<li class="fixed-list__item project-income-list__item">
        <span class="fixed-list__name">${escapeHtml(name)}</span>
        <div class="fixed-list__item-trailing">
          <span class="fixed-list__amount fixed-list__amount--income">${formatYen(amount)}</span>${actionBtns}
        </div>
      </li>`;
      })
      .join("");
  }
  if (el.projectIncomeInputBreakdownEmpty) el.projectIncomeInputBreakdownEmpty.hidden = pairs.length > 0;
}

/** 変動内訳で常に行を出す項目（未入力は ¥0） */
function variableBreakdownFixedOrder() {
  return [
    FOOD_CATEGORY,
    SAISON_CATEGORY,
    MITSUI_CATEGORY,
    ...VARIABLE_PAYMENT_LABELS.filter((c) => c !== FOOD_CATEGORY),
  ];
}

function updateMonthLabel() {
  const d = parseMonthKey(viewMonthKey);
  el.monthLabel.dateTime = `${viewMonthKey}-01`;
  el.monthLabel.textContent = d.toLocaleDateString("ja-JP", { year: "numeric", month: "long" });
}

/**
 * @param {FixedExpenseItem[]} items
 * @param {"home"|"input"} where
 */
function fixedExpenseListInnerHtml(items, where) {
  const emptyMsg =
    where === "input"
      ? "まだ登録がありません。上のフォームから項目名と月額を登録できます。"
      : "項目がありません。入力ページの「固定支出」から登録してください。";
  if (!items.length) {
    return `<li class="fixed-list__item fixed-expense-list__item fixed-expense-list__item--empty">
      <span class="fixed-expense-list__empty-msg">${escapeHtml(emptyMsg)}</span>
    </li>`;
  }
  return items
    .map(
      (item) => `<li class="fixed-list__item fixed-expense-list__item">
      <span class="fixed-list__name">${escapeHtml(item.name)}</span>
      <div class="fixed-expense-list__trailing">
        <span class="fixed-list__amount fixed-expense-list__amount">${formatYen(item.amount)}</span>
        <div class="fixed-expense-list__actions">
          <button type="button" class="btn btn--small btn--ghost" data-fixed-edit-id="${escapeHtml(item.id)}">修正</button>
          <button type="button" class="btn btn--small btn--danger" data-fixed-delete-id="${escapeHtml(item.id)}">削除</button>
        </div>
      </div>
    </li>`
    )
    .join("");
}

function renderFixedExpenseList() {
  const items = state.fixedExpenses;
  const total = fixedExpenseTotal(state);
  const totalText = formatYen(total);

  if (el.fixedExpenseList) el.fixedExpenseList.innerHTML = fixedExpenseListInnerHtml(items, "home");
  if (el.fixedExpenseInputList) el.fixedExpenseInputList.innerHTML = fixedExpenseListInnerHtml(items, "input");

  if (el.fixedExpenseTotalLine) el.fixedExpenseTotalLine.textContent = totalText;
  if (el.fixedExpenseInputTotalLine) el.fixedExpenseInputTotalLine.textContent = totalText;
}

/** @param {string | null} editId 新規は null */
function openFixedExpenseDialog(editId) {
  if (!el.dialogFixedExpense) return;
  const titleEl = document.getElementById("fixed-expense-dialog-title");
  if (editId) {
    const item = state.fixedExpenses.find((x) => x.id === editId);
    if (!item) return;
    if (el.fixedExpenseEditId) el.fixedExpenseEditId.value = item.id;
    if (el.fixedExpenseName) el.fixedExpenseName.value = item.name;
    if (el.fixedExpenseAmount) el.fixedExpenseAmount.value = String(item.amount);
    if (titleEl) titleEl.textContent = "固定支出を修正";
  } else {
    if (el.fixedExpenseEditId) el.fixedExpenseEditId.value = "";
    if (el.fixedExpenseName) el.fixedExpenseName.value = "";
    if (el.fixedExpenseAmount) el.fixedExpenseAmount.value = "";
    if (titleEl) titleEl.textContent = "固定支出を追加";
  }
  el.dialogFixedExpense.showModal();
  queueMicrotask(() => el.fixedExpenseName?.focus());
}

function closeFixedExpenseDialog() {
  closeDatePicker();
  closeMonthPicker();
  el.dialogFixedExpense?.close();
}

function saveFixedExpenseFromForm() {
  const name = el.fixedExpenseName?.value.trim() ?? "";
  const amountRaw = el.fixedExpenseAmount?.value.trim() ?? "";
  const amount = Math.floor(Number(amountRaw));
  const editId = el.fixedExpenseEditId?.value ?? "";

  if (!name) {
    alert("項目名を入力してください");
    return;
  }
  if (!Number.isFinite(amount) || amount < 0) {
    alert("金額を0以上の整数で入力してください");
    return;
  }

  if (editId) {
    const idx = state.fixedExpenses.findIndex((x) => x.id === editId);
    if (idx < 0) {
      alert("項目が見つかりません");
      closeFixedExpenseDialog();
      return;
    }
    state.fixedExpenses[idx] = { ...state.fixedExpenses[idx], name, amount };
  } else {
    state.fixedExpenses.push({ id: createFixedExpenseId(), name, amount });
  }
  saveState(state);
  closeFixedExpenseDialog();
  render();
}

/**
 * 変動支出（内訳）の行を組み立てる。表示月の支出はすべてここで集計する（クイック入力・変動支出・取引を追加を区別しない）。
 * @param {Entry[]} monthEntries
 * @returns {Array<[string, number]>}
 */
function buildVariableBreakdownRows(monthEntries) {
  const fixedOrder = variableBreakdownFixedOrder();
  const fixedSet = new Set(fixedOrder.map((c) => normalizeExpenseCategory(c)));

  const totals = new Map();
  for (const e of monthEntries) {
    if (e.kind !== "expense") continue;
    const key = normalizeExpenseCategory(e.category) || "（未分類）";
    totals.set(key, (totals.get(key) ?? 0) + e.amount);
  }

  const rows = /** @type {Array<[string, number]>} */ ([]);
  for (const label of fixedOrder) {
    const norm = normalizeExpenseCategory(label);
    rows.push([label, totals.get(norm) ?? 0]);
  }

  const seenNorm = new Set(fixedOrder.map((c) => normalizeExpenseCategory(c)));
  for (const cat of EXPENSE_CATEGORIES) {
    const norm = normalizeExpenseCategory(cat);
    if (seenNorm.has(norm)) continue;
    seenNorm.add(norm);
    const amt = totals.get(norm) ?? 0;
    if (amt > 0) rows.push([cat, amt]);
  }

  const expenseNorms = new Set(EXPENSE_CATEGORIES.map((c) => normalizeExpenseCategory(c)));
  const orphans = [...totals.entries()]
    .filter(([normKey, amt]) => amt > 0 && !fixedSet.has(normKey) && !expenseNorms.has(normKey))
    .sort((a, b) => {
      if (b[1] !== a[1]) return b[1] - a[1];
      return a[0].localeCompare(b[0], "ja");
    });
  for (const [normKey, amt] of orphans) {
    rows.push([normKey, amt]);
  }

  return rows;
}

/** @param {Entry[]} monthEntries @param {number} variableExpenseTotal */
function renderVariableExpenseBreakdown(monthEntries, variableExpenseTotal) {
  if (!el.variableExpenseByCategory) return;

  if (variableExpenseTotal <= 0) {
    el.variableExpenseByCategory.innerHTML = "";
    if (el.variableBreakdownEmpty) el.variableBreakdownEmpty.hidden = false;
    if (el.variableExpenseMonthTotal) el.variableExpenseMonthTotal.textContent = formatYen(0);
    return;
  }

  const displayPairs = buildVariableBreakdownRows(monthEntries);

  if (el.variableBreakdownEmpty) el.variableBreakdownEmpty.hidden = true;
  el.variableExpenseByCategory.innerHTML = displayPairs
    .map(([name, amount], rowIndex) => {
      const remark = variableBreakdownRemarkFor(name);
      const remarkBlock = remark
        ? `<span class="variable-breakdown-list__remark">${escapeHtml(remark)}</span>`
        : "";
      const enc = escapeHtml(encodeURIComponent(name));
      const actionBtns =
        amount > 0
          ? `<div class="variable-breakdown-list__actions">
        <button type="button" class="btn btn--small btn--ghost variable-breakdown-list__edit" data-breakdown-variable="${enc}" aria-label="この項目の取引を修正">修正</button>
        <button type="button" class="btn btn--small btn--danger variable-breakdown-list__delete" data-breakdown-variable-delete="${enc}" aria-label="この項目の表示月（引き落とし月）に含まれる取引をすべて削除">削除</button>
      </div>`
          : "";
      const periodBounds = getUsagePeriodBoundsForBillingMonth(viewMonthKey, name);
      const periodRange = formatUsagePeriodRangeJa(periodBounds.start, periodBounds.end);
      const periodCol = `<div class="variable-breakdown-list__period-col">
        <span class="variable-breakdown-list__period-k">利用期間</span>
        <span class="variable-breakdown-list__period-v">${escapeHtml(periodRange)}</span>
      </div>`;
      const rowEntries = getExpenseEntriesForVariableRowName(name);
      const detailId = `variable-breakdown-detail-${rowIndex}`;
      return `<li class="fixed-list__item variable-breakdown-list__item" aria-expanded="false" aria-controls="${detailId}">
      <div class="variable-breakdown-list__label-col">
        <span class="fixed-list__name">${escapeHtml(name)}</span>${remarkBlock ? `\n        ${remarkBlock}` : ""}
      </div>
      ${periodCol}
      <div class="variable-breakdown-list__trailing">
        <span class="fixed-list__amount">${formatYen(amount)}</span>${actionBtns}
      </div>
      <div class="variable-breakdown-list__detail" id="${detailId}" hidden>
        <p class="variable-breakdown-list__detail-title">内訳（日付・金額・メモ）</p>
        ${buildVariableBreakdownDetailInnerHtml(rowEntries)}
      </div>
    </li>`;
    })
    .join("");
  if (el.variableExpenseMonthTotal) el.variableExpenseMonthTotal.textContent = formatYen(variableExpenseTotal);
}

/** @param {string} recordId */
async function downloadExpenseReceiptImage(recordId) {
  const records = Array.isArray(state.expenseRecords) ? state.expenseRecords : [];
  const r = records.find((x) => x.id === recordId);
  if (!r?.receiptDataUrl) return;

  const safeId = r.id.replace(/[^a-zA-Z0-9._-]/g, "_").slice(0, 48);
  const filename = `receipt-${r.date}-${safeId}.jpg`;

  try {
    const res = await fetch(r.receiptDataUrl);
    const blob = await res.blob();
    const objUrl = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = objUrl;
    a.download = filename;
    a.rel = "noopener";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(objUrl);
  } catch {
    try {
      const a = document.createElement("a");
      a.href = r.receiptDataUrl;
      a.download = filename;
      a.rel = "noopener";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    } catch {
      window.open(r.receiptDataUrl, "_blank", "noopener,noreferrer");
    }
  }
}

function renderExpenseRecords() {
  if (!el.expenseRecordList || !el.expenseRecordEmpty) return;

  const records = Array.isArray(state.expenseRecords) ? state.expenseRecords : [];
  const list = records.filter((r) => r.date.startsWith(viewMonthKey)).sort((a, b) => {
    if (a.date !== b.date) return b.date.localeCompare(a.date);
    return b.id.localeCompare(a.id);
  });

  if (!list.length) {
    el.expenseRecordList.innerHTML = "";
    el.expenseRecordEmpty.hidden = false;
    return;
  }

  el.expenseRecordEmpty.hidden = true;
  el.expenseRecordList.innerHTML = list
    .map((r) => {
      const memoHtml = r.memo.trim()
        ? `<p class="expense-record-card__memo">${escapeHtml(r.memo.trim())}</p>`
        : `<p class="expense-record-card__memo expense-record-card__memo--empty">（メモなし）</p>`;
      const imgHtml = r.receiptDataUrl
        ? `<div class="expense-record-card__photo-block">
        <div class="expense-record-card__photo-toolbar">
          <button type="button" class="btn btn--small btn--ghost" data-expense-record-download="${escapeHtml(r.id)}" aria-label="領収書の写真をダウンロード">ダウンロード</button>
        </div>
        <div class="expense-record-card__photo-wrap">
        <img class="expense-record-card__photo" src="${escapeHtml(r.receiptDataUrl)}" alt="領収書の画像" loading="lazy" />
      </div>
      </div>`
        : `<p class="expense-record-card__no-photo">（写真なし）</p>`;
      return `<li class="expense-record-card">
      <div class="expense-record-card__head">
        <time class="expense-record-card__date" datetime="${escapeHtml(r.date)}">${escapeHtml(formatDateLongJa(r.date))}</time>
        <button type="button" class="btn btn--small btn--danger" data-expense-record-delete="${escapeHtml(r.id)}">削除</button>
      </div>
      ${memoHtml}
      ${imgHtml}
    </li>`;
    })
    .join("");
}

function render() {
  updateMonthLabel();
  const filterKind = /** @type {Kind|'all'} */ (el.filterKind.value);
  const monthAll = state.entries.filter((e) => {
    if (e.kind === "income") return e.date.startsWith(viewMonthKey);
    return getExpenseBillingMonthKey(e) === viewMonthKey;
  });
  const { income, variableExpense } = summarize(monthAll);
  const fixedTotal = fixedExpenseTotal(state);
  const balance = income - variableExpense - fixedTotal;

  el.sumIncome.textContent = formatYen(income);
  el.sumVariableExpense.textContent = formatYen(variableExpense);
  el.sumFixedExpense.textContent = formatYen(fixedTotal);
  el.sumBalance.textContent = formatYen(balance);
  el.sumBalance.style.color = balance >= 0 ? "var(--income)" : "var(--expense)";

  if (el.sumSavingsRate) {
    if (income <= 0) {
      el.sumSavingsRate.textContent = "—";
      el.sumSavingsRate.style.color = "var(--muted)";
    } else {
      const ratePct = (balance / income) * 100;
      el.sumSavingsRate.textContent = `${ratePct.toFixed(1)}%`;
      el.sumSavingsRate.style.color = ratePct >= 0 ? "var(--income)" : "var(--expense)";
    }
  }

  const fixedPlusVariable = fixedTotal + variableExpense;
  if (el.combinedLineFixed) el.combinedLineFixed.textContent = formatYen(fixedTotal);
  if (el.combinedLineVariable) el.combinedLineVariable.textContent = formatYen(variableExpense);
  if (el.sumFixedPlusVariable) el.sumFixedPlusVariable.textContent = formatYen(fixedPlusVariable);
  if (el.sumTopFixedPlusVariable) el.sumTopFixedPlusVariable.textContent = formatYen(fixedPlusVariable);

  renderProjectIncomeHome(monthAll);
  renderProjectIncomeInputBreakdown(monthAll);
  renderFixedExpenseList();
  renderVariableExpenseBreakdown(monthAll, variableExpense);
  renderExpenseRecords();

  const todayIso = isoDateFromDate(new Date());
  if (el.foodRecordingDate) {
    el.foodRecordingDate.dateTime = todayIso;
    el.foodRecordingDate.textContent = formatDateLongJa(todayIso);
  }
  if (el.variablePayRecordingDate) {
    el.variablePayRecordingDate.dateTime = todayIso;
    el.variablePayRecordingDate.textContent = formatDateLongJa(todayIso);
  }
  if (el.foodSumToday) el.foodSumToday.textContent = formatYen(sumExpenseCategoryDay(state.entries, FOOD_CATEGORY, todayIso));
  if (el.foodSumMonth) el.foodSumMonth.textContent = formatYen(sumExpenseCategoryMonth(state.entries, FOOD_CATEGORY, viewMonthKey));
  if (el.saisonRecordingDate) {
    el.saisonRecordingDate.dateTime = todayIso;
    el.saisonRecordingDate.textContent = formatDateLongJa(todayIso);
  }
  if (el.mitsuiRecordingDate) {
    el.mitsuiRecordingDate.dateTime = todayIso;
    el.mitsuiRecordingDate.textContent = formatDateLongJa(todayIso);
  }
  if (el.saisonSumToday) el.saisonSumToday.textContent = formatYen(sumExpenseCategoryDay(state.entries, SAISON_CATEGORY, todayIso));
  if (el.saisonSumMonth) el.saisonSumMonth.textContent = formatYen(sumExpenseCategoryMonth(state.entries, SAISON_CATEGORY, viewMonthKey));
  if (el.mitsuiSumToday) el.mitsuiSumToday.textContent = formatYen(sumExpenseCategoryDay(state.entries, MITSUI_CATEGORY, todayIso));
  if (el.mitsuiSumMonth) el.mitsuiSumMonth.textContent = formatYen(sumExpenseCategoryMonth(state.entries, MITSUI_CATEGORY, viewMonthKey));
  if (el.projectRecordingDate) {
    el.projectRecordingDate.dateTime = firstDayOfMonthKey(viewMonthKey);
    el.projectRecordingDate.textContent = formatMonthLongJaFromKey(viewMonthKey);
  }
  if (el.projectSumMonth) {
    el.projectSumMonth.textContent = formatYen(sumIncomeCategoryMonth(state.entries, PROJECT_INCOME_CATEGORY, viewMonthKey));
  }
  if (el.projectUsagePeriod) {
    el.projectUsagePeriod.textContent = `利用・計上月（表示中の月と連動・月切替で自動更新）: ${formatMonthLongJaFromKey(viewMonthKey)}`;
  }

  updateQuickInputUsagePeriodLabels(new Date());

  const rows = filterEntries(state.entries, viewMonthKey, filterKind).sort((a, b) => {
    if (a.date !== b.date) return b.date.localeCompare(a.date);
    return b.id.localeCompare(a.id);
  });

  el.tbody.innerHTML = rows
    .map((e) => {
      const kindLabel = e.kind === "income" ? "収入" : "支出";
      const badgeClass = e.kind === "income" ? "badge--income" : "badge--expense";
      const amountClass = e.kind === "income" ? "var(--income)" : "var(--expense)";
      return `<tr>
        <td>${escapeHtml(e.date)}</td>
        <td><span class="badge ${badgeClass}">${kindLabel}</span></td>
        <td>${escapeHtml(normalizeExpenseCategory(e.category) || e.category || "—")}</td>
        <td class="table__num" style="color:${amountClass}">${formatYen(e.amount)}</td>
        <td>${escapeHtml(e.note || "—")}</td>
        <td class="table__actions">
          <button type="button" class="btn btn--small btn--ghost btn--edit-row" data-action="edit" data-id="${escapeHtml(e.id)}">編集</button>
          <button type="button" class="btn btn--danger btn--small" data-action="delete" data-id="${escapeHtml(e.id)}">削除</button>
        </td>
      </tr>`;
    })
    .join("");

  el.empty.hidden = rows.length > 0;
}

document.querySelectorAll('input[name="kind"]').forEach((radio) => {
  radio.addEventListener("change", () => {
    refreshCategoryOptions();
    updateTradeUsagePeriodLine();
  });
});

el.fieldCategory?.addEventListener("change", () => {
  updateTradeUsagePeriodLine();
});

el.fieldDate?.addEventListener("change", () => {
  updateTradeUsagePeriodLine();
});
el.fieldDate?.addEventListener("input", () => {
  updateTradeUsagePeriodLine();
});

el.btnPrev.addEventListener("click", () => {
  const d = parseMonthKey(viewMonthKey);
  d.setMonth(d.getMonth() - 1);
  viewMonthKey = monthKeyFromDate(d);
  render();
});

el.btnNext.addEventListener("click", () => {
  const d = parseMonthKey(viewMonthKey);
  d.setMonth(d.getMonth() + 1);
  viewMonthKey = monthKeyFromDate(d);
  render();
});

el.btnThisMonth.addEventListener("click", () => {
  viewMonthKey = monthKeyFromDate(new Date());
  render();
});

if (el.btnMonthPicker) {
  el.btnMonthPicker.addEventListener("click", (e) => {
    e.preventDefault();
    openMonthPicker();
  });
}

/**
 * @param {object} cfg
 * @param {HTMLFormElement | null} cfg.form
 * @param {HTMLInputElement | null | undefined} cfg.amountInput
 * @param {HTMLInputElement | null | undefined} cfg.noteInput
 * @param {string} cfg.category
 * @param {string} cfg.alertInvalidAmount
 */
function bindDailyExpenseSubmit(cfg) {
  const { form, amountInput, noteInput, category, alertInvalidAmount } = cfg;
  if (!form) return;
  const submitBtn = /** @type {HTMLButtonElement | null} */ (form.querySelector("button.btn--primary"));
  const run = () => {
    const date = isoDateFromDate(new Date());
    const amountRaw = amountInput?.value.trim() ?? "";
    const amount = Math.floor(Number(amountRaw));
    const note = noteInput?.value.trim() ?? "";

    if (!Number.isFinite(amount) || amount < 1) {
      alert(alertInvalidAmount);
      return;
    }

    const entry = /** @type {Entry} */ ({
      id: newId(),
      kind: "expense",
      date,
      amount,
      category,
      note,
    });

    state.entries.push(entry);
    saveState(state);
    if (amountInput) amountInput.value = "";
    if (noteInput) noteInput.value = "";
    render();
    amountInput?.focus();
  };
  submitBtn?.addEventListener("click", () => {
    run();
  });
  form.addEventListener("submit", (ev) => {
    ev.preventDefault();
  });
}

bindDailyExpenseSubmit({
  form: el.formFood,
  amountInput: el.foodAmount,
  noteInput: el.foodNote,
  category: FOOD_CATEGORY,
  alertInvalidAmount: "食費（PayPay）の金額を正しく入力してください",
});

bindDailyExpenseSubmit({
  form: el.formSaison,
  amountInput: el.saisonAmount,
  noteInput: el.saisonNote,
  category: SAISON_CATEGORY,
  alertInvalidAmount: "セゾンの金額を正しく入力してください",
});

bindDailyExpenseSubmit({
  form: el.formMitsui,
  amountInput: el.mitsuiAmount,
  noteInput: el.mitsuiNote,
  category: MITSUI_CATEGORY,
  alertInvalidAmount: "三井住友カードの金額を正しく入力してください",
});

/**
 * @param {object} cfg
 * @param {HTMLFormElement | null} cfg.form
 * @param {HTMLInputElement | null | undefined} cfg.amountInput
 * @param {HTMLInputElement | null | undefined} cfg.nameInput
 */
function bindDailyProjectIncomeSubmit(cfg) {
  const { form, amountInput, nameInput } = cfg;
  if (!form) return;
  const submitBtn = /** @type {HTMLButtonElement | null} */ (form.querySelector("button.btn--primary"));
  const run = () => {
    const date = firstDayOfMonthKey(viewMonthKey);
    const amountRaw = amountInput?.value.trim() ?? "";
    const amount = Math.floor(Number(amountRaw));
    const projectName = nameInput?.value.trim() ?? "";

    if (!projectName) {
      alert("案件名を入力してください");
      return;
    }
    if (!Number.isFinite(amount) || amount < 1) {
      alert("金額を正しく入力してください");
      return;
    }

    state.entries.push(
      /** @type {Entry} */ ({
        id: newId(),
        kind: "income",
        date,
        amount,
        category: PROJECT_INCOME_CATEGORY,
        note: projectName,
      })
    );
    saveState(state);
    if (amountInput) amountInput.value = "";
    if (nameInput) nameInput.value = "";
    render();
    nameInput?.focus();
  };
  submitBtn?.addEventListener("click", () => {
    run();
  });
  form.addEventListener("submit", (ev) => {
    ev.preventDefault();
  });
}

bindDailyProjectIncomeSubmit({
  form: el.formProjectIncome,
  amountInput: el.projectIncomeAmount,
  nameInput: el.projectIncomeName,
});

if (el.formVariablePay) {
  const variablePaySubmitBtn = /** @type {HTMLButtonElement | null} */ (
    el.formVariablePay.querySelector("button.btn--primary")
  );
  const runVariablePay = () => {
    const date = isoDateFromDate(new Date());
    const category = el.variablePayItem?.value ?? "";
    const amountRaw = el.variablePayAmount?.value.trim() ?? "";
    const amount = Math.floor(Number(amountRaw));
    const note = el.variablePayNote?.value.trim() ?? "";

    if (!category || !Number.isFinite(amount) || amount < 1) {
      alert("項目と金額を正しく入力してください");
      return;
    }

    const entry = /** @type {Entry} */ ({
      id: newId(),
      kind: "expense",
      date,
      amount,
      category,
      note,
    });

    state.entries.push(entry);
    saveState(state);
    if (el.variablePayAmount) el.variablePayAmount.value = "";
    if (el.variablePayNote) el.variablePayNote.value = "";
    render();
    el.variablePayAmount?.focus();
  };
  variablePaySubmitBtn?.addEventListener("click", () => {
    runVariablePay();
  });
  el.formVariablePay.addEventListener("submit", (ev) => {
    ev.preventDefault();
  });
}

if (el.formFixedExpenseInput) {
  const fixedExpenseInputSubmitBtn = /** @type {HTMLButtonElement | null} */ (
    el.formFixedExpenseInput.querySelector("button.btn--primary")
  );
  const runFixedExpenseInputAdd = () => {
    const name = el.fixedExpenseInputName?.value.trim() ?? "";
    const amountRaw = el.fixedExpenseInputAmount?.value.trim() ?? "";
    const amount = Math.floor(Number(amountRaw));

    if (!name) {
      alert("項目名を入力してください");
      return;
    }
    if (!Number.isFinite(amount) || amount < 0) {
      alert("金額を0以上の整数で入力してください");
      return;
    }

    state.fixedExpenses.push({ id: createFixedExpenseId(), name, amount });
    saveState(state);
    if (el.fixedExpenseInputAmount) el.fixedExpenseInputAmount.value = "";
    if (el.fixedExpenseInputName) el.fixedExpenseInputName.value = "";
    render();
    el.fixedExpenseInputName?.focus();
  };
  fixedExpenseInputSubmitBtn?.addEventListener("click", () => {
    runFixedExpenseInputAdd();
  });
  el.formFixedExpenseInput.addEventListener("submit", (ev) => {
    ev.preventDefault();
  });
}

if (el.form) {
  const tradeSubmitBtn = /** @type {HTMLButtonElement | null} */ (el.form.querySelector("button.btn--primary"));
  const runTradeEntry = () => {
    const kind = getSelectedKind();
    const date = el.fieldDate.value;
    const amountRaw = el.fieldAmount.value.trim();
    const amount = Math.floor(Number(amountRaw));
    const category = el.fieldCategory.value;
    const note = el.fieldNote.value.trim();

    if (!date || !category || !Number.isFinite(amount) || amount < 1) {
      alert("日付・金額・カテゴリを正しく入力してください");
      return;
    }

    const entry = /** @type {Entry} */ ({
      id: newId(),
      kind,
      date,
      amount,
      category,
      note,
    });

    state.entries.push(entry);
    saveState(state);
    el.fieldAmount.value = "";
    el.fieldNote.value = "";
    render();
  };
  tradeSubmitBtn?.addEventListener("click", () => {
    runTradeEntry();
  });
  el.form.addEventListener("submit", (ev) => {
    ev.preventDefault();
  });
}

el.filterKind.addEventListener("change", render);

el.tbody.addEventListener("click", (ev) => {
  const t = ev.target;
  if (!(t instanceof HTMLElement)) return;
  const btn = t.closest("button[data-action][data-id]");
  if (!btn) return;
  const id = btn.getAttribute("data-id");
  const action = btn.getAttribute("data-action");
  if (!id || !action) return;
  if (action === "edit") {
    openEditEntryDialog(id);
    return;
  }
  if (action === "delete") {
    if (!confirm("この取引を削除しますか")) return;
    state.entries = state.entries.filter((e) => e.id !== id);
    saveState(state);
    render();
  }
});

el.variableExpenseByCategory?.addEventListener("click", (ev) => {
  const t = ev.target instanceof Element ? ev.target : null;
  if (!t) return;

  const delBtn = t.closest("[data-breakdown-variable-delete]");
  if (delBtn instanceof HTMLButtonElement) {
    const enc = delBtn.getAttribute("data-breakdown-variable-delete");
    if (!enc) return;
    const rowName = decodeURIComponent(enc);
    const list = getExpenseEntriesForVariableRowName(rowName);
    if (!list.length) return;
    if (!confirm(`「${rowName}」の表示月（引き落とし月）に含まれる取引 ${list.length} 件をすべて削除しますか？`)) return;
    const ids = new Set(list.map((e) => e.id));
    state.entries = state.entries.filter((e) => !ids.has(e.id));
    saveState(state);
    render();
    return;
  }

  const editBtn = t.closest("[data-breakdown-variable]");
  if (editBtn instanceof HTMLButtonElement) {
    const enc = editBtn.getAttribute("data-breakdown-variable");
    if (!enc) return;
    openVariableBreakdownRowForEdit(enc);
    return;
  }

  if (t.closest(".variable-breakdown-list__actions")) return;
  if (t.closest(".variable-breakdown-list__detail")) return;

  const item = t.closest(".variable-breakdown-list__item");
  if (!item) return;
  const detail = item.querySelector(".variable-breakdown-list__detail");
  if (!(detail instanceof HTMLElement)) return;

  detail.hidden = !detail.hidden;
  item.setAttribute("aria-expanded", detail.hidden ? "false" : "true");
  item.classList.toggle("variable-breakdown-list__item--open", !detail.hidden);
});

function onProjectIncomeBreakdownClick(ev) {
  const delBtn = ev.target instanceof Element ? ev.target.closest("[data-breakdown-project-delete]") : null;
  if (delBtn instanceof HTMLButtonElement) {
    const enc = delBtn.getAttribute("data-breakdown-project-delete");
    if (!enc) return;
    const name = decodeURIComponent(enc);
    const list = getProjectIncomeEntriesForDisplayName(name);
    if (!list.length) return;
    if (!confirm(`「${name}」の当月の案件収入 ${list.length} 件をすべて削除しますか？`)) return;
    const ids = new Set(list.map((e) => e.id));
    state.entries = state.entries.filter((e) => !ids.has(e.id));
    saveState(state);
    render();
    return;
  }
  const btn = ev.target instanceof Element ? ev.target.closest("[data-breakdown-project]") : null;
  if (!(btn instanceof HTMLButtonElement)) return;
  const enc = btn.getAttribute("data-breakdown-project");
  if (!enc) return;
  openProjectIncomeRowForEdit(enc);
}

el.projectIncomeList?.addEventListener("click", onProjectIncomeBreakdownClick);
el.projectIncomeInputBreakdown?.addEventListener("click", onProjectIncomeBreakdownClick);

function onFixedExpenseListClick(ev) {
  const delBtn = ev.target instanceof Element ? ev.target.closest("[data-fixed-delete-id]") : null;
  if (delBtn instanceof HTMLButtonElement) {
    const id = delBtn.getAttribute("data-fixed-delete-id");
    if (!id || !confirm("この固定支出項目を削除しますか")) return;
    state.fixedExpenses = state.fixedExpenses.filter((x) => x.id !== id);
    saveState(state);
    render();
    return;
  }
  const editBtn = ev.target instanceof Element ? ev.target.closest("[data-fixed-edit-id]") : null;
  if (editBtn instanceof HTMLButtonElement) {
    const id = editBtn.getAttribute("data-fixed-edit-id");
    if (id) openFixedExpenseDialog(id);
  }
}

el.fixedExpenseList?.addEventListener("click", onFixedExpenseListClick);
el.fixedExpenseInputList?.addEventListener("click", onFixedExpenseListClick);

el.btnFixedExpenseCancel?.addEventListener("click", () => {
  closeFixedExpenseDialog();
});

el.btnFixedExpenseSave?.addEventListener("click", (e) => {
  e.preventDefault();
  saveFixedExpenseFromForm();
});

el.formFixedExpense?.addEventListener("submit", (e) => {
  e.preventDefault();
});

if (el.dialogFixedExpense) {
  el.dialogFixedExpense.addEventListener("click", (ev) => {
    if (ev.target === el.dialogFixedExpense) closeFixedExpenseDialog();
  });
}

if (el.breakdownPickList) {
  el.breakdownPickList.addEventListener("click", (ev) => {
    const delBtn = ev.target instanceof Element ? ev.target.closest("[data-pick-delete-id]") : null;
    if (delBtn instanceof HTMLButtonElement) {
      const id = delBtn.getAttribute("data-pick-delete-id");
      if (!id || !confirm("この取引を削除しますか")) return;
      state.entries = state.entries.filter((e) => e.id !== id);
      saveState(state);
      render();
      if (breakdownPickListRefetch) {
        const next = breakdownPickListRefetch();
        renderBreakdownPickListBody(next);
        if (!next.length) el.dialogBreakdownPick?.close();
      } else {
        el.dialogBreakdownPick?.close();
      }
      return;
    }
    const btn = ev.target instanceof Element ? ev.target.closest("[data-pick-edit-id]") : null;
    if (!(btn instanceof HTMLButtonElement)) return;
    const id = btn.getAttribute("data-pick-edit-id");
    if (!id) return;
    showInlineEditForEntry(id, true);
  });
}

if (el.btnBreakdownPickClose && el.dialogBreakdownPick) {
  el.btnBreakdownPickClose.addEventListener("click", () => el.dialogBreakdownPick.close());
}

if (el.dialogBreakdownPick) {
  el.dialogBreakdownPick.addEventListener("close", () => {
    closeDatePicker();
    closeMonthPicker();
    breakdownPickListRefetch = null;
    breakdownPickEditReturnToList = false;
    if (el.breakdownPickStepList) el.breakdownPickStepList.hidden = false;
    if (el.breakdownPickStepEdit) el.breakdownPickStepEdit.hidden = true;
    if (el.btnInlineEditBack) el.btnInlineEditBack.hidden = true;
    resetEditFormFields();
  });
  el.dialogBreakdownPick.addEventListener("click", (ev) => {
    if (ev.target === el.dialogBreakdownPick) el.dialogBreakdownPick.close();
  });
}

if (el.formInlineEdit) {
  const runInlineEditSave = () => {
    const id = el.inlineEditEntryId?.value ?? "";
    const kind = getEditKind();
    const date = el.editDate?.value ?? "";
    const amountRaw = el.editAmount?.value.trim() ?? "";
    const amount = Math.floor(Number(amountRaw));
    const note = el.editNote?.value.trim() ?? "";

    if (!id) {
      alert("取引を特定できません");
      return;
    }
    if (!date || !Number.isFinite(amount) || amount < 1) {
      alert("日付・金額を正しく入力してください");
      return;
    }

    const idx = state.entries.findIndex((e) => e.id === id);
    if (idx < 0) {
      alert("取引が見つかりません");
      el.dialogBreakdownPick?.close();
      return;
    }

    const prev = state.entries[idx];
    let category = prev.category;
    if (kind !== prev.kind) {
      category = kind === "income" ? "その他（収入）" : "その他（支出）";
    }
    state.entries[idx] = {
      ...prev,
      kind,
      date,
      amount,
      category,
      note,
    };
    saveState(state);
    el.dialogBreakdownPick?.close();
    render();
  };
  el.btnInlineEditSave?.addEventListener("click", () => {
    runInlineEditSave();
  });
  el.formInlineEdit.addEventListener("submit", (ev) => {
    ev.preventDefault();
  });
}

if (el.btnInlineEditBack) {
  el.btnInlineEditBack.addEventListener("click", () => {
    resetBreakdownPickToListStep();
  });
}

if (el.btnInlineEditClose) {
  el.btnInlineEditClose.addEventListener("click", () => {
    el.dialogBreakdownPick?.close();
  });
}

/**
 * @param {string} src
 * @returns {Promise<void>}
 */
function loadScript(src) {
  return new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = src;
    s.async = true;
    s.onload = () => resolve();
    s.onerror = () => reject(new Error("script load failed"));
    document.head.appendChild(s);
  });
}

function ensureGoogleIdentityLoaded() {
  if (window.google?.accounts?.oauth2) return Promise.resolve();
  return loadScript("https://accounts.google.com/gsi/client");
}

/** @param {string} raw */
function normalizeDriveFolderId(raw) {
  const s = raw.trim();
  const m = s.match(/folders\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  return s;
}

/**
 * @param {string} accessToken
 * @param {string} folderId
 * @param {string} jsonText
 */
async function uploadJsonToGoogleDrive(accessToken, folderId, jsonText) {
  const boundary = `household_budget_${Math.random().toString(36).slice(2)}_${Date.now()}`;
  const metadata = {
    name: `household-budget-backup-${viewMonthKey}-${Date.now()}.json`,
    parents: [folderId],
  };
  const delimiter = `\r\n--${boundary}\r\n`;
  const close = `\r\n--${boundary}--`;
  const body =
    delimiter +
    "Content-Type: application/json; charset=UTF-8\r\n\r\n" +
    JSON.stringify(metadata) +
    delimiter +
    "Content-Type: application/json; charset=UTF-8\r\n\r\n" +
    jsonText +
    close;

  const res = await fetch(
    "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink",
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": `multipart/related; boundary=${boundary}`,
      },
      body,
    }
  );
  const text = await res.text();
  if (!res.ok) {
    let msg = text;
    try {
      const j = JSON.parse(text);
      if (j.error?.message) msg = j.error.message;
    } catch {
      /* 生テキストのまま */
    }
    throw new Error(msg || `HTTP ${res.status}`);
  }
  return JSON.parse(text);
}

function openGdriveSettingsDialog() {
  if (!el.dialogGdriveSettings) return;
  if (el.gdriveClientId) el.gdriveClientId.value = localStorage.getItem(GDRIVE_CLIENT_ID_KEY) || "";
  if (el.gdriveFolderId) el.gdriveFolderId.value = localStorage.getItem(GDRIVE_FOLDER_ID_KEY) || "";
  el.dialogGdriveSettings.showModal();
}

function saveGdriveSettingsFromForm() {
  const clientId = el.gdriveClientId?.value.trim() ?? "";
  const folderRaw = el.gdriveFolderId?.value.trim() ?? "";
  const folderId = normalizeDriveFolderId(folderRaw);
  if (!clientId) {
    alert("OAuth 2.0 クライアント ID を入力してください");
    return;
  }
  if (!folderId) {
    alert("保存先フォルダ ID を入力してください");
    return;
  }
  localStorage.setItem(GDRIVE_CLIENT_ID_KEY, clientId);
  localStorage.setItem(GDRIVE_FOLDER_ID_KEY, folderId);
  el.dialogGdriveSettings?.close();
}

function getGdriveConfig() {
  const clientId = (localStorage.getItem(GDRIVE_CLIENT_ID_KEY) || "").trim();
  const folderId = normalizeDriveFolderId(localStorage.getItem(GDRIVE_FOLDER_ID_KEY) || "");
  return { clientId, folderId };
}

function startBackupToGoogleDrive() {
  if (!window.isSecureContext || location.protocol === "file:") {
    alert("Googleドライブへの保存は、https または http://localhost で開いたときのみ利用できます");
    return;
  }
  const { clientId, folderId } = getGdriveConfig();
  if (!clientId || !folderId) {
    alert("先に「ドライブ設定」でクライアント ID と保存先フォルダ ID を保存してください");
    openGdriveSettingsDialog();
    return;
  }

  ensureGoogleIdentityLoaded()
    .then(() => {
      const g = window.google;
      if (!g?.accounts?.oauth2) {
        throw new Error("Google の認証ライブラリを読み込めませんでした");
      }
      const jsonText = JSON.stringify(state, null, 2);
      const tokenClient = g.accounts.oauth2.initTokenClient({
        client_id: clientId,
        scope: GDRIVE_OAUTH_SCOPE,
        callback: (resp) => {
          if (resp.error) {
            alert(resp.error_description || resp.error || "Google の認証に失敗しました");
            return;
          }
          if (!resp.access_token) {
            alert("アクセストークンを取得できませんでした");
            return;
          }
          uploadJsonToGoogleDrive(resp.access_token, folderId, jsonText)
            .then((data) => {
              const link = data.webViewLink ? `\n\nブラウザで開く: ${data.webViewLink}` : "";
              alert(`Googleドライブに保存しました。${link}`);
            })
            .catch((err) => {
              alert(
                `アップロードに失敗しました。\n${err.message || String(err)}\n\nDrive API が有効か、フォルダ ID・クライアント ID・JavaScript 生成元をご確認ください。`
              );
            });
        },
      });
      tokenClient.requestAccessToken({ prompt: "" });
    })
    .catch(() => {
      alert("Google のスクリプトを読み込めませんでした。ネットワークをご確認ください");
    });
}

if (el.btnExportGdrive) {
  el.btnExportGdrive.addEventListener("click", () => {
    startBackupToGoogleDrive();
  });
}

if (el.btnGdriveSettings) {
  el.btnGdriveSettings.addEventListener("click", () => {
    openGdriveSettingsDialog();
  });
}

if (el.btnGdriveSettingsSave) {
  el.btnGdriveSettingsSave.addEventListener("click", () => {
    saveGdriveSettingsFromForm();
  });
}

if (el.btnGdriveSettingsCancel) {
  el.btnGdriveSettingsCancel.addEventListener("click", () => {
    el.dialogGdriveSettings?.close();
  });
}

if (el.dialogGdriveSettings) {
  el.dialogGdriveSettings.addEventListener("click", (ev) => {
    if (ev.target === el.dialogGdriveSettings) el.dialogGdriveSettings.close();
  });
}

el.btnExport.addEventListener("click", () => {
  const blob = new Blob([JSON.stringify(state, null, 2)], { type: "application/json" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `household-budget-backup-${viewMonthKey}.json`;
  a.click();
  URL.revokeObjectURL(a.href);
});

el.inputImport.addEventListener("change", () => {
  const file = el.inputImport.files?.[0];
  el.inputImport.value = "";
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const data = JSON.parse(String(reader.result));
      if (!data || !Array.isArray(data.entries)) {
        alert("無効なバックアップ形式です");
        return;
      }
      const entries = data.entries.filter(isValidEntry);
      const fixedExpenses =
        data.fixedExpenses !== undefined ? normalizeFixedExpensesArray(data.fixedExpenses) : defaultFixedExpenses();
      const expenseRecords =
        data.expenseRecords !== undefined ? normalizeExpenseRecordsArray(data.expenseRecords) : [];
      if (
        !confirm(
          `バックアップから取引 ${entries.length} 件・固定支出 ${fixedExpenses.length} 件・経費（メモ・領収書）${expenseRecords.length} 件を読み込み、現在のデータと置き換えますか`
        )
      )
        return;
      state = { version: 1, entries, fixedExpenses, expenseRecords };
      saveState(state);
      render();
    } catch {
      alert("ファイルの読み込みに失敗しました");
    }
  };
  reader.readAsText(file);
});

if (el.btnExpenseRecordSave) {
  el.btnExpenseRecordSave.addEventListener("click", async () => {
    const date = el.expenseRecordDate?.value ?? "";
    const memo = el.expenseRecordMemo?.value.trim() ?? "";
    const file = getPickedExpensePhotoFile();
    let receiptDataUrl = "";
    if (file) {
      try {
        receiptDataUrl = await compressImageFileToDataUrl(file);
        if (receiptDataUrl.length > 2_500_000) {
          alert("画像を縮小してもデータが大きすぎます。別の画像を試してください");
          return;
        }
      } catch (err) {
        alert(err instanceof Error ? err.message : String(err));
        return;
      }
    }
    if (!date || !/^\d{4}-\d{2}-\d{2}$/.test(date)) {
      alert("日付を正しく選択してください");
      return;
    }
    if (!memo && !receiptDataUrl) {
      alert("メモを入力するか、領収書の写真を選択してください");
      return;
    }
    if (!Array.isArray(state.expenseRecords)) state.expenseRecords = [];
    state.expenseRecords.push({
      id: createExpenseRecordId(),
      date,
      memo,
      ...(receiptDataUrl ? { receiptDataUrl } : {}),
    });
    saveState(state);
    if (el.expenseRecordMemo) el.expenseRecordMemo.value = "";
    clearExpensePhotoInputs();
    if (el.expenseRecordDate) el.expenseRecordDate.value = isoDateFromDate(new Date());
    render();
  });
}

el.formExpenseRecord?.addEventListener("submit", (ev) => {
  ev.preventDefault();
});

el.expenseRecordList?.addEventListener("click", (ev) => {
  const t = ev.target instanceof Element ? ev.target : null;
  if (!t) return;

  const dlBtn = t.closest("[data-expense-record-download]");
  if (dlBtn instanceof HTMLButtonElement) {
    const id = dlBtn.getAttribute("data-expense-record-download");
    if (id) void downloadExpenseReceiptImage(id);
    return;
  }

  const btn = t.closest("[data-expense-record-delete]");
  if (!(btn instanceof HTMLButtonElement)) return;
  const id = btn.getAttribute("data-expense-record-delete");
  if (!id || !confirm("この経費記録を削除しますか（メモ・写真も消えます）")) return;
  if (!Array.isArray(state.expenseRecords)) state.expenseRecords = [];
  state.expenseRecords = state.expenseRecords.filter((r) => r.id !== id);
  saveState(state);
  render();
});

document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "visible") render();
});

/** @typedef {"food"|"saison"|"mitsui"|"project"|"variable"|"fixed"|"trade"|"list"} InputTab */

const VALID_INPUT_TABS = /** @type {InputTab[]} */ ([
  "food",
  "saison",
  "mitsui",
  "project",
  "variable",
  "fixed",
  "trade",
  "list",
]);
const DEFAULT_INPUT_TAB = /** @type {InputTab} */ ("food");

function parseRoute() {
  const raw = (location.hash || "#/").replace(/^#\/?/, "").trim();
  if (!raw || raw === "/") return { page: /** @type {"home"} */ ("home"), tab: /** @type {InputTab | null} */ (null) };
  const parts = raw.split("/").filter(Boolean);
  if (parts[0] === "expenses") {
    return { page: /** @type {"expenses"} */ ("expenses"), tab: /** @type {InputTab | null} */ (null) };
  }
  if (parts[0] === "input") {
    const tab = VALID_INPUT_TABS.includes(/** @type {InputTab} */ (parts[1])) ? /** @type {InputTab} */ (parts[1]) : DEFAULT_INPUT_TAB;
    return { page: /** @type {"input"} */ ("input"), tab };
  }
  return { page: "home", tab: null };
}

function normalizeHashRoute() {
  const raw = (location.hash || "").replace(/^#\/?/, "").trim();
  if (!raw || raw === "/") {
    if (location.hash !== "#/" && location.hash !== "" && location.hash !== "#") {
      history.replaceState(null, "", "#/");
    }
    return;
  }
  const parts = raw.split("/").filter(Boolean);
  if (parts[0] === "expenses" && parts.length > 1) {
    history.replaceState(null, "", "#/expenses");
    return;
  }
  if (parts[0] === "input" && (!parts[1] || !VALID_INPUT_TABS.includes(/** @type {InputTab} */ (parts[1])))) {
    history.replaceState(null, "", `#/input/${DEFAULT_INPUT_TAB}`);
  }
}

/** @param {"home"|"input"|"expenses"} page @param {InputTab | null} tab */
function setPageVisibility(page, tab) {
  const home = document.getElementById("page-home");
  const input = document.getElementById("page-input");
  const expenses = document.getElementById("page-expenses");
  if (!home || !input || !expenses) return;

  home.toggleAttribute("hidden", page !== "home");
  input.toggleAttribute("hidden", page !== "input");
  expenses.toggleAttribute("hidden", page !== "expenses");

  home.setAttribute("aria-hidden", page === "home" ? "false" : "true");
  input.setAttribute("aria-hidden", page === "input" ? "false" : "true");
  expenses.setAttribute("aria-hidden", page === "expenses" ? "false" : "true");

  if (page === "input") {
    const activeTab = tab || DEFAULT_INPUT_TAB;
    VALID_INPUT_TABS.forEach((key) => {
      const panel = document.getElementById(`input-panel-${key}`);
      if (panel) panel.toggleAttribute("hidden", key !== activeTab);
    });
  }

  document.querySelectorAll(".app-nav__link").forEach((link) => {
    const route = link.getAttribute("data-route");
    const match = route === page;
    if (match) link.setAttribute("aria-current", "page");
    else link.removeAttribute("aria-current");
  });

  const activeTab = tab || DEFAULT_INPUT_TAB;
  document.querySelectorAll(".input-tabs__tab").forEach((tabEl) => {
    const key = tabEl.getAttribute("data-input-tab");
    const on = page === "input" && key === activeTab;
    if (on) {
      tabEl.setAttribute("aria-current", "page");
      tabEl.setAttribute("aria-selected", "true");
    } else {
      tabEl.removeAttribute("aria-current");
      tabEl.setAttribute("aria-selected", "false");
    }
  });
}

function refreshScrollRevealForRoute() {
  const reduced = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  document.querySelectorAll(".scroll-reveal").forEach((el) => el.classList.remove("is-visible"));
  if (reduced) {
    document.querySelectorAll(".scroll-reveal").forEach((el) => el.classList.add("is-visible"));
    return;
  }
  document.querySelectorAll(".app-shell.scroll-reveal").forEach((el) => el.classList.add("is-visible"));
  requestAnimationFrame(() => {
    const home = document.getElementById("page-home");
    const input = document.getElementById("page-input");
    const expenses = document.getElementById("page-expenses");
    const visible =
      home && !home.hasAttribute("hidden")
        ? home
        : input && !input.hasAttribute("hidden")
          ? input
          : expenses && !expenses.hasAttribute("hidden")
            ? expenses
            : null;
    visible?.querySelectorAll(".scroll-reveal").forEach((el) => el.classList.add("is-visible"));
  });
}

function applyRoute() {
  normalizeHashRoute();
  const { page, tab } = parseRoute();
  setPageVisibility(page, tab);
  if (page === "expenses" && el.expenseRecordDate && !el.expenseRecordDate.value) {
    el.expenseRecordDate.value = isoDateFromDate(new Date());
  }
  refreshScrollRevealForRoute();
  window.scrollTo(0, 0);
}

/** PWA: localhost / HTTPS のときだけ登録（file:// では不可） */
function registerServiceWorker() {
  if (!("serviceWorker" in navigator)) return;
  const { protocol, hostname } = window.location;
  const isLocal =
    hostname === "localhost" || hostname === "127.0.0.1" || hostname === "[::1]";
  if (protocol !== "https:" && !isLocal) return;
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("./sw.js", { scope: "./" }).catch(() => {});
  });
}

/** すでにホーム画面追加済み（スタンドアロン）か */
function isStandaloneDisplayMode() {
  if (window.matchMedia("(display-mode: standalone)").matches) return true;
  if (window.matchMedia("(display-mode: window-controls-overlay)").matches) return true;
  const nav = /** @type {{ standalone?: boolean }} */ (navigator);
  return nav.standalone === true;
}

/** iPhone / iPad（ホーム画面追加の案内が必要な環境） */
function isAppleMobileLike() {
  const ua = navigator.userAgent || "";
  if (/iPad|iPhone|iPod/.test(ua)) return true;
  return navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1;
}

function initPwaGuideDialog() {
  const open = () => {
    el.dialogPwaGuide?.showModal();
  };
  const close = () => {
    el.dialogPwaGuide?.close();
  };
  el.btnPwaGuide?.addEventListener("click", open);
  el.btnPwaGuideClose?.addEventListener("click", close);
  el.dialogPwaGuide?.addEventListener("click", (ev) => {
    if (ev.target === el.dialogPwaGuide) close();
  });
}

/**
 * Chrome/Edge/Android: beforeinstallprompt。
 * iOS 系: イベントは来ないため、ホーム画面追加の案内バナーを表示する。
 */
function initPwaInstallBanner() {
  const banner = document.getElementById("pwa-install-banner");
  const textEl = document.getElementById("pwa-install-banner-text");
  const btnDo = document.getElementById("pwa-install-do");
  const btnDismiss = document.getElementById("pwa-install-dismiss");
  if (!banner || !textEl || !btnDismiss) return;

  if (isStandaloneDisplayMode()) return;

  const dismissKey = "pwa-install-banner-dismissed";
  if (sessionStorage.getItem(dismissKey) === "1") return;

  const { protocol, hostname } = window.location;
  const isLocal =
    hostname === "localhost" || hostname === "127.0.0.1" || hostname === "[::1]";
  if (protocol !== "https:" && !isLocal) return;

  /** @type {{ prompt: () => Promise<void>; userChoice: Promise<{ outcome: string }> } | null} */
  let deferredPrompt = null;
  let appleHintShown = false;

  function hideBanner(dismissed) {
    banner.hidden = true;
    if (dismissed) sessionStorage.setItem(dismissKey, "1");
  }

  btnDismiss.addEventListener("click", () => hideBanner(true));

  function showAppleInstallHint() {
    if (appleHintShown) return;
    appleHintShown = true;
    textEl.textContent =
      "ホーム画面に追加すると、アプリのように全画面で開けます。Safari の「共有」→「ホーム画面に追加」から追加してください（詳しくはヘッダーの「アプリで使う」）。";
    if (btnDo) btnDo.hidden = true;
    banner.hidden = false;
  }

  if (isAppleMobileLike()) {
    showAppleInstallHint();
  }

  window.addEventListener("beforeinstallprompt", (e) => {
    e.preventDefault();
    deferredPrompt = /** @type {{ prompt: () => Promise<void>; userChoice: Promise<{ outcome: string }> }} */ (e);
    textEl.textContent =
      "ホーム画面に追加すると、アプリのように全画面で開け、オフラインでも使いやすくなります（対応ブラウザのみ）。";
    if (btnDo) btnDo.hidden = false;
    banner.hidden = false;
  });

  if (btnDo) {
    btnDo.addEventListener("click", async () => {
      if (!deferredPrompt) return;
      deferredPrompt.prompt();
      await deferredPrompt.userChoice.catch(() => {});
      deferredPrompt = null;
      hideBanner(true);
    });
  }

  window.addEventListener("appinstalled", () => hideBanner(true));
}

registerServiceWorker();
initPwaGuideDialog();
initPwaInstallBanner();

window.addEventListener("hashchange", applyRoute);

refreshCategoryOptions();
initVariablePaySelect();
setDefaultDate();
initCustomDatePickers();
initExpensePhotoInputs();
render();
if (!location.hash || location.hash === "#") {
  history.replaceState(null, "", "#/");
}
applyRoute();
