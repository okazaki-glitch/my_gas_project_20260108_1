const SHEET_NAMES = {
  settings: "設定",
  records: "記録",
};

const SETTINGS_KEYS = {
  defaultWeightKg: "default_weight_kg",
  dailyTargetKcal: "daily_target_kcal",
  gender: "gender",
  age: "age",
  monthlyGoalKg: "monthly_goal_kg",
  activityLevel: "activity_level",
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
  const monthlySummary = getMonthlySummary_(
    sheets.records,
    settings,
    new Date()
  );
  return {
    settings: settings,
    today: today,
    todaySummary: todaySummary,
    monthlySummary: monthlySummary,
  };
}

function saveSettings(payload) {
  const sheets = initSheets_();
  const defaultWeightKg = toNumber_(payload.defaultWeightKg);
  const dailyTargetKcal = toNumber_(payload.dailyTargetKcal);
  const gender = normalizeGender_(payload.gender);
  const age = toNumber_(payload.age);
  const monthlyGoalKg = toNumber_(payload.monthlyGoalKg);
  const activityLevel = normalizeActivityLevel_(payload.activityLevel);
  upsertSetting_(sheets.settings, SETTINGS_KEYS.defaultWeightKg, defaultWeightKg);
  upsertSetting_(sheets.settings, SETTINGS_KEYS.dailyTargetKcal, dailyTargetKcal);
  upsertSetting_(sheets.settings, SETTINGS_KEYS.gender, gender);
  upsertSetting_(sheets.settings, SETTINGS_KEYS.age, age);
  upsertSetting_(sheets.settings, SETTINGS_KEYS.monthlyGoalKg, monthlyGoalKg);
  upsertSetting_(sheets.settings, SETTINGS_KEYS.activityLevel, activityLevel);
  const settings = getSettings_(sheets.settings);
  const monthlySummary = getMonthlySummary_(
    sheets.records,
    settings,
    new Date()
  );
  return { settings: settings, monthlySummary: monthlySummary };
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
  const settings = getSettings_(sheets.settings);
  const monthlySummary = getMonthlySummary_(
    sheets.records,
    settings,
    new Date()
  );
  return { calories: calories, summary: summary, monthlySummary: monthlySummary };
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
  upsertSetting_(sheet, SETTINGS_KEYS.gender, current.gender);
  upsertSetting_(sheet, SETTINGS_KEYS.age, current.age);
  upsertSetting_(sheet, SETTINGS_KEYS.monthlyGoalKg, current.monthlyGoalKg);
  upsertSetting_(sheet, SETTINGS_KEYS.activityLevel, current.activityLevel);
}

function getSettings_(sheet) {
  const defaults = {
    defaultWeightKg: 60,
    dailyTargetKcal: 2000,
    gender: "male",
    age: 30,
    monthlyGoalKg: -1,
    activityLevel: "medium",
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
    defaultWeightKg: coalesceNumber_(
      map[SETTINGS_KEYS.defaultWeightKg],
      defaults.defaultWeightKg
    ),
    dailyTargetKcal: coalesceNumber_(
      map[SETTINGS_KEYS.dailyTargetKcal],
      defaults.dailyTargetKcal
    ),
    gender: map[SETTINGS_KEYS.gender] || defaults.gender,
    age: coalesceNumber_(map[SETTINGS_KEYS.age], defaults.age),
    monthlyGoalKg: coalesceNumber_(
      map[SETTINGS_KEYS.monthlyGoalKg],
      defaults.monthlyGoalKg
    ),
    activityLevel: normalizeActivityLevel_(
      map[SETTINGS_KEYS.activityLevel] || defaults.activityLevel
    ),
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

function getMonthlySummary_(sheet, settings, baseDate) {
  const date = baseDate instanceof Date ? baseDate : new Date();
  const tz = Session.getScriptTimeZone();
  const monthKey = Utilities.formatDate(date, tz, "yyyy-MM");
  const year = date.getFullYear();
  const month = date.getMonth();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const daysElapsed = Math.min(date.getDate(), daysInMonth);

  const bmrPerDay = estimateBmr_(settings.gender, settings.age);
  const activityFactor = resolveActivityFactor_(settings.activityLevel);
  const energyPerDay = Math.round(bmrPerDay * activityFactor);
  const energyTotal = Math.round(energyPerDay * daysElapsed);
  const bmrTotal = Math.round(bmrPerDay * daysElapsed);
  const runningTotal = Math.round(getMonthlyRunningTotal_(sheet, year, month));
  const totalBurn = Math.round(energyTotal + runningTotal);
  const targetIntakeTotal = Math.round(
    toNumber_(settings.dailyTargetKcal) * daysElapsed
  );
  const deficit = Math.round(totalBurn - targetIntakeTotal);

  const goalKg = toNumber_(settings.monthlyGoalKg);
  const targetAmount = Math.round(Math.abs(goalKg) * 7500);
  let goalType = "maintain";
  if (goalKg < 0) {
    goalType = "deficit";
  } else if (goalKg > 0) {
    goalType = "surplus";
  }
  const progressAmount = Math.min(runningTotal, targetAmount);
  let targetTotalBurn = targetIntakeTotal;
  if (goalType === "deficit") {
    targetTotalBurn = targetIntakeTotal + targetAmount;
  } else if (goalType === "surplus") {
    targetTotalBurn = Math.max(targetIntakeTotal - targetAmount, 0);
  }
  const remaining = Math.max(targetAmount - runningTotal, 0);

  return {
    monthKey: monthKey,
    daysInMonth: daysInMonth,
    daysElapsed: daysElapsed,
    activityLevel: settings.activityLevel,
    activityFactor: activityFactor,
    bmrPerDay: Math.round(bmrPerDay),
    bmrTotal: bmrTotal,
    energyPerDay: energyPerDay,
    energyTotal: energyTotal,
    runningTotal: runningTotal,
    totalBurn: totalBurn,
    targetTotalBurn: targetTotalBurn,
    targetIntakeTotal: targetIntakeTotal,
    deficit: deficit,
    goalKg: goalKg,
    goalType: goalType,
    targetAmount: targetAmount,
    progressAmount: progressAmount,
    remaining: remaining,
  };
}

function getMonthlyRunningTotal_(sheet, year, month) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return 0;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  let total = 0;
  values.forEach((row) => {
    const cell = row[0];
    if (!cell) {
      return;
    }
    const date = cell instanceof Date ? cell : new Date(cell);
    if (isNaN(date.getTime())) {
      return;
    }
    if (date.getFullYear() === year && date.getMonth() === month) {
      total += toNumber_(row[4]);
    }
  });
  return total;
}

function estimateBmr_(gender, age) {
  const normalized = normalizeGender_(gender);
  const value = toNumber_(age);
  if (!normalized || !value) {
    return 0;
  }

  // 年齢・性別からの簡易平均推定（個人差を考慮しない）
  const table = {
    male: [
      { min: 0, max: 17, kcal: 1350 },
      { min: 18, max: 29, kcal: 1530 },
      { min: 30, max: 49, kcal: 1500 },
      { min: 50, max: 69, kcal: 1400 },
      { min: 70, max: 120, kcal: 1280 },
    ],
    female: [
      { min: 0, max: 17, kcal: 1250 },
      { min: 18, max: 29, kcal: 1210 },
      { min: 30, max: 49, kcal: 1170 },
      { min: 50, max: 69, kcal: 1110 },
      { min: 70, max: 120, kcal: 1010 },
    ],
  };

  const list = table[normalized];
  for (let i = 0; i < list.length; i++) {
    if (value >= list[i].min && value <= list[i].max) {
      return list[i].kcal;
    }
  }
  return list[list.length - 1].kcal;
}

function normalizeGender_(value) {
  const str = value ? String(value).toLowerCase() : "";
  if (str === "male" || str === "m" || str === "男性") {
    return "male";
  }
  if (str === "female" || str === "f" || str === "女性") {
    return "female";
  }
  return "";
}

function normalizeActivityLevel_(value) {
  const str = value ? String(value).toLowerCase() : "";
  if (
    str === "low" ||
    str === "medium" ||
    str === "high" ||
    str === "低い" ||
    str === "標準" ||
    str === "高い"
  ) {
    if (str === "低い") return "low";
    if (str === "高い") return "high";
    if (str === "標準") return "medium";
    return str;
  }
  return "medium";
}

function resolveActivityFactor_(level) {
  // 活動レベルの係数（目安）
  const normalized = normalizeActivityLevel_(level);
  if (normalized === "low") return 1.2;
  if (normalized === "high") return 1.75;
  return 1.55;
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

function coalesceNumber_(value, fallback) {
  if (value === "" || value === null || value === undefined) {
    return fallback;
  }
  const num = Number(value);
  return Number.isFinite(num) ? num : fallback;
}
