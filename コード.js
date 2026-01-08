const SHEET_NAMES = {
  settings: "設定",
  records: "記録",
};

const SETTINGS_KEYS = {
  defaultWeightKg: "default_weight_kg",
  dailyTargetKcal: "daily_target_kcal",
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("カロリー管理")
    .addItem("サイドバーを開く", "showSidebar")
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("ランニングカロリー管理");
  SpreadsheetApp.getUi().showSidebar(html);
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("ランニングカロリー管理");
}

function getAppData() {
  const sheets = initSheets_();
  const settings = getSettings_(sheets.settings);
  const today = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );
  const todaySummary = getDailySummary_(sheets.records, today);
  return { settings: settings, today: today, todaySummary: todaySummary };
}

function saveSettings(payload) {
  const sheets = initSheets_();
  const defaultWeightKg = toNumber_(payload.defaultWeightKg);
  const dailyTargetKcal = toNumber_(payload.dailyTargetKcal);
  upsertSetting_(sheets.settings, SETTINGS_KEYS.defaultWeightKg, defaultWeightKg);
  upsertSetting_(sheets.settings, SETTINGS_KEYS.dailyTargetKcal, dailyTargetKcal);
  return getSettings_(sheets.settings);
}

function saveRunRecord(payload) {
  const sheets = initSheets_();
  const date = normalizeDate_(payload.date);
  const distanceKm = toNumber_(payload.distanceKm);
  const durationMin = toNumber_(payload.durationMin);
  const weightKg = toNumber_(payload.weightKg);
  const memo = payload.memo ? String(payload.memo) : "";

  // 距離か体重が未入力のときは0として扱う
  const calories = Math.round(
    calculateCalories_(distanceKm, durationMin, weightKg)
  );

  sheets.records.appendRow([
    date,
    distanceKm || "",
    durationMin || "",
    weightKg || "",
    calories || "",
    memo,
    new Date(),
  ]);

  const dateKey = Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );
  const summary = getDailySummary_(sheets.records, dateKey);
  return { calories: calories, summary: summary };
}

function initSheets_() {
  // 既存シートがない場合は作成する
  const settings = ensureSheet_(SHEET_NAMES.settings, ["キー", "値"]);
  const records = ensureSheet_(SHEET_NAMES.records, [
    "日付",
    "距離(km)",
    "時間(分)",
    "体重(kg)",
    "推定消費(kcal)",
    "メモ",
    "登録日時",
  ]);

  ensureDefaultSettings_(settings);
  return { settings: settings, records: records };
}

function ensureSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

function ensureDefaultSettings_(sheet) {
  const current = getSettings_(sheet);
  upsertSetting_(sheet, SETTINGS_KEYS.defaultWeightKg, current.defaultWeightKg);
  upsertSetting_(sheet, SETTINGS_KEYS.dailyTargetKcal, current.dailyTargetKcal);
}

function getSettings_(sheet) {
  const defaults = {
    defaultWeightKg: 60,
    dailyTargetKcal: 2000,
  };

  const lastRow = sheet.getLastRow();
  const values =
    lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 2).getValues() : [];

  const map = {};
  values.forEach((row) => {
    const key = row[0];
    if (key) {
      map[key] = row[1];
    }
  });

  return {
    defaultWeightKg:
      toNumber_(map[SETTINGS_KEYS.defaultWeightKg]) || defaults.defaultWeightKg,
    dailyTargetKcal:
      toNumber_(map[SETTINGS_KEYS.dailyTargetKcal]) ||
      defaults.dailyTargetKcal,
  };
}

function upsertSetting_(sheet, key, value) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    sheet.getRange(2, 1, 1, 2).setValues([[key, value]]);
    return;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === key) {
      sheet.getRange(i + 2, 2).setValue(value);
      return;
    }
  }

  sheet.getRange(lastRow + 1, 1, 1, 2).setValues([[key, value]]);
}

function getDailySummary_(sheet, dateKey) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { date: dateKey, totalCalories: 0, count: 0 };
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const tz = Session.getScriptTimeZone();
  let total = 0;
  let count = 0;

  values.forEach((row) => {
    const cell = row[0];
    if (!cell) {
      return;
    }
    const date = cell instanceof Date ? cell : new Date(cell);
    const key = Utilities.formatDate(date, tz, "yyyy-MM-dd");
    if (key === dateKey) {
      total += toNumber_(row[4]);
      count++;
    }
  });

  return { date: dateKey, totalCalories: Math.round(total), count: count };
}

function calculateCalories_(distanceKm, durationMin, weightKg) {
  if (!distanceKm || !weightKg) {
    return 0;
  }

  if (durationMin && durationMin > 0) {
    const hours = durationMin / 60;
    const speed = distanceKm / hours;
    const met = resolveMet_(speed);
    return met * weightKg * hours;
  }

  return 1.036 * weightKg * distanceKm;
}

function resolveMet_(speedKmh) {
  if (speedKmh < 6.4) return 6.0;
  if (speedKmh < 8.0) return 8.3;
  if (speedKmh < 9.7) return 9.8;
  if (speedKmh < 11.3) return 11.0;
  if (speedKmh < 12.9) return 11.8;
  if (speedKmh < 14.5) return 12.8;
  if (speedKmh < 16.1) return 14.5;
  return 16.0;
}

function normalizeDate_(value) {
  if (!value) {
    return new Date();
  }
  if (value instanceof Date) {
    return value;
  }
  const str = String(value);
  if (/^\\d{4}-\\d{2}-\\d{2}$/.test(str)) {
    return new Date(str + "T00:00:00");
  }
  const date = new Date(str);
  return isNaN(date.getTime()) ? new Date() : date;
}

function toNumber_(value) {
  const num = Number(value);
  return isNaN(num) ? 0 : num;
}
