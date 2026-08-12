/**
 * Fishing Day MVP -> Google Sheets
 *
 * Как использовать:
 * 1. Создайте Google Таблицу.
 * 2. Откройте Расширения -> Apps Script.
 * 3. Замените содержимое Code.gs этим файлом.
 * 4. Deploy -> New deployment -> Web app.
 * 5. Execute as: Me.
 * 6. Who has access: Anyone.
 * 7. Скопируйте URL, который заканчивается на /exec, и вставьте его на сайте.
 */

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ok: true, service: "Fishing Day Sheets Sync"}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("Откройте Apps Script из нужной Google Таблицы: Расширения -> Apps Script.");

    writeSummary_(ss, payload.years || []);
    writeYearSheets_(ss, payload.years || []);
    writeCatches_(ss, payload.years || []);

    return ContentService
      .createTextOutput(JSON.stringify({ok: true, syncedAt: new Date().toISOString()}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ok: false, error: String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function writeSummary_(ss, years) {
  const sheet = getOrCreateSheet_(ss, "Сводка");
  sheet.clear();

  const headers = [["Год","Дней с рыбалкой","Всего поймано","Видов рыбы","Лучший день","Улов в лучший день","Последняя синхронизация"]];
  const rows = years.map(y => [
    y.year,
    y.summary.fishingDays,
    y.summary.total,
    y.summary.species,
    y.summary.bestDay || "",
    y.summary.bestTotal || 0,
    new Date()
  ]);

  sheet.getRange(1,1,1,headers[0].length).setValues(headers);
  if (rows.length) sheet.getRange(2,1,rows.length,headers[0].length).setValues(rows);
  formatSheet_(sheet, headers[0].length);
}

function writeYearSheets_(ss, years) {
  years.forEach(y => {
    const name = String(y.year);
    const sheet = getOrCreateSheet_(ss, name);
    sheet.clear();

    const headers = [[
      "Дата","День недели","Рыбалка","Количество","Результат дня",
      "Температура, °C","Ощущается, °C","Погода",
      "Направление ветра","Ветер, °","Скорость ветра, км/ч","Порывы, км/ч",
      "Влажность, %","Давление, hPa","Давление, мм рт. ст.",
      "Осадки, мм","Облачность, %"
    ]];

    const rows = y.rows.map(r => [
      parseDate_(r.date), r.weekday, r.fishing ? "Да" : "Нет", r.total, r.result,
      value_(r.temperature), value_(r.apparentTemperature), r.weather,
      r.windDirection, value_(r.windDegrees), value_(r.windSpeed), value_(r.windGusts),
      value_(r.humidity), value_(r.pressureHpa), value_(r.pressureMm),
      value_(r.precipitation), value_(r.cloudCover)
    ]);

    sheet.getRange(1,1,1,headers[0].length).setValues(headers);
    if (rows.length) {
      sheet.getRange(2,1,rows.length,headers[0].length).setValues(rows);
      sheet.getRange(2,1,rows.length,1).setNumberFormat("dd.mm.yyyy");
    }
    formatSheet_(sheet, headers[0].length);
  });
}

function writeCatches_(ss, years) {
  const sheet = getOrCreateSheet_(ss, "Улов");
  sheet.clear();

  const headers = [["Год","Дата","Время","Рыба","Количество"]];
  const rows = [];
  years.forEach(y => {
    (y.catches || []).forEach(c => rows.push([
      y.year, parseDate_(c.date), c.time, c.fish, c.qty
    ]));
  });

  rows.sort((a,b) => a[1] - b[1] || String(a[2]).localeCompare(String(b[2])));

  sheet.getRange(1,1,1,headers[0].length).setValues(headers);
  if (rows.length) {
    sheet.getRange(2,1,rows.length,headers[0].length).setValues(rows);
    sheet.getRange(2,2,rows.length,1).setNumberFormat("dd.mm.yyyy");
  }
  formatSheet_(sheet, headers[0].length);
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function value_(v) {
  return v === "" || v === null || typeof v === "undefined" ? "" : v;
}

function parseDate_(iso) {
  if (!iso) return "";
  const p = iso.split("-").map(Number);
  return new Date(p[0], p[1]-1, p[2]);
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
  sheet.autoResizeColumns(1, columnCount);
  sheet.getDataRange().setVerticalAlignment("middle");
}
