/**
 * Рыболовный дневник -> Google Sheets
 * Версия 7: снасть, способ ловли, место и приманки для каждой записи улова.
 */

const SPREADSHEET_ID = "1H7H2AUwtfeqYaaWE0LCZq4QFm7o5Rp7M8eYg70uJLkM";

function doGet() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    return jsonResponse_({
      ok: true,
      service: "Fishing Day Sheets Sync",
      spreadsheet: ss.getName(),
      spreadsheetId: SPREADSHEET_ID
    });
  } catch (error) {
    return jsonResponse_({ok: false, error: String(error)});
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    if (!e || !e.postData || !e.postData.contents) throw new Error("Запрос не содержит данных.");

    const payload = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const years = Array.isArray(payload.years) ? payload.years : [];

    writeSummary_(ss, years);
    writeYearSheets_(ss, years);
    writeCatches_(ss, years);
    SpreadsheetApp.flush();

    return jsonResponse_({
      ok: true,
      message: "Данные успешно записаны",
      yearsProcessed: years.length,
      syncedAt: new Date().toISOString()
    });
  } catch (error) {
    console.error(error);
    return jsonResponse_({ok: false, error: String(error), syncedAt: new Date().toISOString()});
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function writeSummary_(ss, years) {
  const sheet = getOrCreateSheet_(ss, "Сводка");
  prepareSheet_(sheet);

  const headers = [[
    "Год", "Дней с рыбалкой", "Всего поймано", "Видов рыбы",
    "Лучший день", "Улов в лучший день", "Последняя синхронизация"
  ]];

  const rows = years.map(function(y) {
    const s = y.summary || {};
    return [
      y.year || "", number_(s.fishingDays), number_(s.total), number_(s.species),
      s.bestDay ? parseDate_(s.bestDay) : "", number_(s.bestTotal), new Date()
    ];
  });

  sheet.getRange(1,1,1,headers[0].length).setValues(headers);
  if (rows.length) {
    sheet.getRange(2,1,rows.length,headers[0].length).setValues(rows);
    sheet.getRange(2,5,rows.length,1).setNumberFormat("dd.mm.yyyy");
    sheet.getRange(2,7,rows.length,1).setNumberFormat("dd.mm.yyyy hh:mm");
  }
  formatSheet_(sheet, headers[0].length);
}

function writeYearSheets_(ss, years) {
  years.forEach(function(y) {
    const sheet = getOrCreateSheet_(ss, String(y.year));
    prepareSheet_(sheet);

    const headers = [[
      "Дата", "День недели", "Рыбалка", "Количество", "Результат дня", "Комментарий",
      "Температура, °C", "Ощущается, °C", "Погода",
      "Направление ветра", "Ветер, °", "Скорость ветра, км/ч", "Порывы ветра, км/ч",
      "Влажность, %", "Давление, hPa", "Давление, мм рт. ст.",
      "Осадки, мм", "Облачность, %"
    ]];

    const rows = (Array.isArray(y.rows) ? y.rows : []).map(function(r) {
      return [
        parseDate_(r.date), r.weekday || "", r.fishing ? "Да" : "Нет", number_(r.total),
        r.result || "", r.comment || "",
        value_(r.temperature), value_(r.apparentTemperature), r.weather || "",
        r.windDirection || "", value_(r.windDegrees), value_(r.windSpeed), value_(r.windGusts),
        value_(r.humidity), value_(r.pressureHpa), value_(r.pressureMm),
        value_(r.precipitation), value_(r.cloudCover)
      ];
    });

    sheet.getRange(1,1,1,headers[0].length).setValues(headers);
    if (rows.length) {
      sheet.getRange(2,1,rows.length,headers[0].length).setValues(rows);
      sheet.getRange(2,1,rows.length,1).setNumberFormat("dd.mm.yyyy");
    }
    formatSheet_(sheet, headers[0].length);

    if (rows.length) {
      const f = sheet.getFilter();
      if (f) f.remove();
      sheet.getRange(1,1,rows.length+1,headers[0].length).createFilter();
    }
  });
}

function writeCatches_(ss, years) {
  const sheet = getOrCreateSheet_(ss, "Улов");
  prepareSheet_(sheet);
  const headers = [["Год", "Дата", "Время", "Рыба", "Количество", "Снасть", "Способ ловли", "Место ловли", "Приманка"]];
  const rows = [];

  years.forEach(function(y) {
    (Array.isArray(y.catches) ? y.catches : []).forEach(function(c) {
      rows.push([y.year || "", parseDate_(c.date), c.time || "", c.fish || "", number_(c.qty), c.tackle || "", c.method || "", c.place || "", Array.isArray(c.baits) ? c.baits.join(", ") : (c.baits || "")]);
    });
  });

  rows.sort(function(a,b) {
    const da = a[1] instanceof Date ? a[1].getTime() : 0;
    const db = b[1] instanceof Date ? b[1].getTime() : 0;
    return da !== db ? da-db : String(a[2]).localeCompare(String(b[2]));
  });

  sheet.getRange(1,1,1,headers[0].length).setValues(headers);
  if (rows.length) {
    sheet.getRange(2,1,rows.length,headers[0].length).setValues(rows);
    sheet.getRange(2,2,rows.length,1).setNumberFormat("dd.mm.yyyy");
  }
  formatSheet_(sheet, headers[0].length);

  if (rows.length) {
    const f = sheet.getFilter();
    if (f) f.remove();
    sheet.getRange(1,1,rows.length+1,headers[0].length).createFilter();
  }
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function prepareSheet_(sheet) {
  const filter = sheet.getFilter();
  if (filter) filter.remove();
  sheet.clear();
  sheet.clearConditionalFormatRules();
  sheet.setFrozenRows(0);
}

function formatSheet_(sheet, columnCount) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const range = sheet.getRange(1,1,lastRow,columnCount);
  range.setFontFamily("Arial");
  sheet.getRange(1,1,1,columnCount)
    .setBackground("#F4DC4F")
    .setFontColor("#000000")
    .setFontWeight("bold");
  sheet.setFrozenRows(1);
  range.setVerticalAlignment("middle");
  sheet.autoResizeColumns(1,columnCount);
  for (let c=1;c<=columnCount;c++) {
    if (sheet.getColumnWidth(c)>350) sheet.setColumnWidth(c,350);
  }
}

function parseDate_(isoDate) {
  if (!isoDate) return "";
  const p = String(isoDate).split("-").map(Number);
  if (p.length!==3 || !p[0] || !p[1] || !p[2]) return isoDate;
  return new Date(p[0],p[1]-1,p[2]);
}

function value_(v) {
  return v === "" || v === null || typeof v === "undefined" ? "" : v;
}

function number_(v) {
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

function jsonResponse_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function testConnection() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  console.log("Подключение успешно. Таблица: " + ss.getName());
}
