/**
 * Рыболовный дневник -> Google Sheets
 * Версия 16: накопительная синхронизация + окно корректировки 7 дней назад.
 *
 * Логика:
 * 1. Все новые даты после последней записи добавляются в таблицу.
 *    Поэтому если синхронизацию не запускали несколько дней, промежуток не потеряется.
 * 2. Сегодня + 7 предыдущих календарных дней разрешено корректировать.
 * 3. Даты старше этого окна НЕ перезаписываются.
 * 4. Лист "Улов" работает по той же логике:
 *    новые даты добавляются, а улов в окне корректировки пересобирается.
 * 5. Будущие строки, оставшиеся от старых версий, удаляются.
 */

const SPREADSHEET_ID = "1H7H2AUwtfeqYaaWE0LCZq4QFm7o5Rp7M8eYg70uJLkM";

// Сегодня + семь предыдущих дней.
// Например, 12 августа можно исправлять 5–12 августа включительно.
const CORRECTION_DAYS_BACK = 7;

const YEAR_HEADERS = [
  "Дата", "День недели", "Рыбалка", "Количество", "Результат дня", "Комментарий",
  "Температура, °C", "Ощущается, °C", "Погода",
  "Направление ветра", "Ветер, °", "Скорость ветра, км/ч", "Порывы ветра, км/ч",
  "Влажность, %", "Давление, hPa", "Давление, мм рт. ст.",
  "Осадки, мм", "Облачность, %",
  "Город погоды",

  // Данные из блока «Добавить улов».
  // Добавлены в конец, чтобы не сдвигать старые исторические столбцы.
  "Город улова",
  "Рыба",
  "Количество по записям",
  "Время улова",
  "Снасть",
  "Способ ловли",
  "Место ловли",
  "Приманка"
];

const CATCH_HEADERS = [
  "Год", "Дата", "Время", "Рыба", "Количество",
  "Снасть", "Способ ловли", "Место ловли", "Приманка",
  "Город"
];

const SUMMARY_HEADERS = [
  "Год", "Дней с рыбалкой", "Всего поймано", "Видов рыбы",
  "Лучший день", "Улов в лучший день", "Последняя синхронизация"
];


function doGet() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const tz = ss.getSpreadsheetTimeZone();
    const todayIso = dateToIso_(new Date(), tz);
    const correctionStartIso = shiftIsoDate_(todayIso, -CORRECTION_DAYS_BACK, tz);

    return jsonResponse_({
      ok: true,
      service: "Fishing Day Sheets Sync",
      version: 16,
      mode: "incremental-with-7-day-correction",
      spreadsheet: ss.getName(),
      spreadsheetId: SPREADSHEET_ID,
      today: todayIso,
      correctionFrom: correctionStartIso,
      correctionTo: todayIso
    });

  } catch (error) {
    return jsonResponse_({
      ok: false,
      error: String(error)
    });
  }
}


function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("Запрос не содержит данных.");
    }

    const payload = JSON.parse(e.postData.contents);
    const years = Array.isArray(payload.years) ? payload.years : [];

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const tz = ss.getSpreadsheetTimeZone();

    // Берём календарную дату с сайта пользователя.
    // Это устраняет расхождение, если часовой пояс самой Google Таблицы
    // отличается от часового пояса, в котором открыт рыболовный дневник.
    const todayIso = isIsoDate_(payload.clientToday)
      ? payload.clientToday
      : dateToIso_(new Date(), tz);

    const correctionStartIso = shiftIsoDate_(
      todayIso,
      -CORRECTION_DAYS_BACK,
      tz
    );

    const syncInfo = writeYearSheets_(
      ss,
      years,
      todayIso,
      correctionStartIso
    );

    // Дополнительная адресная синхронизация комментариев.
    // Обновляем колонку "Комментарий" по дате во всём разрешённом
    // окне корректировки, в том числе если в листе случайно есть
    // несколько строк одной даты после старых версий.
    writeCommentsForCorrectionWindow_(
      ss,
      years,
      correctionStartIso,
      todayIso
    );

    // Явно обновляем колонку "Город погоды" для разрешённого
    // 7-дневного окна корректировки.
    writeWeatherCityForCorrectionWindow_(
      ss,
      years,
      correctionStartIso,
      todayIso
    );

    // Обновляем поля «Добавить улов» прямо в строке дня
    // за сегодня и предыдущие 7 дней.
    writeCatchFieldsForCorrectionWindow_(
      ss,
      years,
      correctionStartIso,
      todayIso
    );

    writeCatches_(
      ss,
      years,
      syncInfo,
      todayIso,
      correctionStartIso
    );

    updateSummary_(
      ss,
      years,
      syncInfo
    );

    SpreadsheetApp.flush();

    return jsonResponse_({
      ok: true,
      message: "Синхронизация завершена",
      correctionFrom: correctionStartIso,
      correctionTo: todayIso,
      result: Object.keys(syncInfo).map(function(year) {
        return {
          year: Number(year),
          newDatesAdded: syncInfo[year].newDates.length,
          correctionDatesUpdated: syncInfo[year].correctionDates.length,
          lastDate: syncInfo[year].lastDate || ""
        };
      }),
      syncedAt: new Date().toISOString()
    });

  } catch (error) {
    console.error(error);

    return jsonResponse_({
      ok: false,
      error: String(error),
      syncedAt: new Date().toISOString()
    });

  } finally {
    try {
      lock.releaseLock();
    } catch (ignore) {}
  }
}


/**
 * Годовые листы.
 *
 * 1. Добавляем все новые даты после последней записи.
 * 2. Отдельно обновляем существующие строки в окне:
 *    correctionStartIso ... todayIso.
 * 3. Всё старше correctionStartIso не трогаем.
 */
function writeYearSheets_(ss, years, todayIso, correctionStartIso) {
  const tz = ss.getSpreadsheetTimeZone();
  const info = {};

  years.forEach(function(yearData) {
    const year = String(yearData.year);
    const sheet = getOrCreateSheet_(ss, year);

    ensureHeaders_(sheet, YEAR_HEADERS);

    // Переход со старых версий: убираем заранее созданное будущее.
    removeFutureRows_(
      sheet,
      1,
      todayIso,
      tz
    );

    const sourceRows = Array.isArray(yearData.rows)
      ? yearData.rows.slice()
      : [];

    sourceRows.sort(function(a, b) {
      return String(a.date).localeCompare(String(b.date));
    });

    const lastBefore = getLastDateInfo_(
      sheet,
      1,
      tz
    );

    const lastIsoBefore = lastBefore
      ? lastBefore.iso
      : "";

    // Все даты после последней записи должны быть добавлены,
    // даже если пользователь не синхронизировал таблицу несколько недель.
    const newRows = sourceRows.filter(function(r) {
      if (!r.date) return false;
      if (r.date > todayIso) return false;

      return !lastIsoBefore ||
        r.date > lastIsoBefore;
    });

    const newDates = [];

    newRows.forEach(function(r) {
      const existing = findDateRow_(
        sheet,
        1,
        r.date,
        tz
      );

      if (existing) {
        return;
      }

      const row = sheet.getLastRow() + 1;

      sheet
        .getRange(
          row,
          1,
          1,
          YEAR_HEADERS.length
        )
        .setValues([
          yearRowToValues_(r)
        ]);

      sheet
        .getRange(row, 1)
        .setNumberFormat("dd.mm.yyyy");

      formatDataRows_(
        sheet,
        row,
        1,
        YEAR_HEADERS.length
      );

      newDates.push(r.date);
    });


    // Последние 7 дней + сегодня можно исправлять.
    const correctionRows = sourceRows.filter(function(r) {
      return r.date &&
        r.date >= correctionStartIso &&
        r.date <= todayIso;
    });

    const correctionDates = [];

    correctionRows.forEach(function(r) {
      let existing = findDateRow_(
        sheet,
        1,
        r.date,
        tz
      );

      // Если внутри окна почему-то есть пропуск,
      // создаём недостающую строку.
      if (!existing) {
        const row = sheet.getLastRow() + 1;

        sheet
          .getRange(
            row,
            1,
            1,
            YEAR_HEADERS.length
          )
          .setValues([
            yearRowToValues_(r)
          ]);

        sheet
          .getRange(row, 1)
          .setNumberFormat("dd.mm.yyyy");

        formatDataRows_(
          sheet,
          row,
          1,
          YEAR_HEADERS.length
        );

        existing = {
          row: row,
          iso: r.date
        };

        if (newDates.indexOf(r.date) === -1) {
          newDates.push(r.date);
        }

      } else {
        // Перезаписывать разрешено только в окне корректировки.
        sheet
          .getRange(
            existing.row,
            1,
            1,
            YEAR_HEADERS.length
          )
          .setValues([
            yearRowToValues_(r)
          ]);

        sheet
          .getRange(existing.row, 1)
          .setNumberFormat("dd.mm.yyyy");

        formatDataRows_(
          sheet,
          existing.row,
          1,
          YEAR_HEADERS.length
        );
      }

      correctionDates.push(r.date);
    });


    formatHeader_(
      sheet,
      YEAR_HEADERS.length
    );

    ensureFilter_(
      sheet,
      YEAR_HEADERS.length
    );

    const lastAfter = getLastDateInfo_(
      sheet,
      1,
      tz
    );

    info[year] = {
      newDates: unique_(newDates),
      correctionDates: unique_(correctionDates),
      lastDate: lastAfter
        ? lastAfter.iso
        : ""
    };
  });

  return info;
}


function yearRowToValues_(r) {
  return [
    parseDate_(r.date),
    r.weekday || "",
    r.fishing ? "Да" : "Нет",
    number_(r.total),
    r.result || "",
    r.comment || "",

    value_(r.temperature),
    value_(r.apparentTemperature),
    r.weather || "",

    r.windDirection || "",
    value_(r.windDegrees),
    value_(r.windSpeed),
    value_(r.windGusts),

    value_(r.humidity),
    value_(r.pressureHpa),
    value_(r.pressureMm),

    value_(r.precipitation),
    value_(r.cloudCover),

    // Пишем именно город, по которому реально получены погодные данные.
    // Если погода не была загружена, значение остаётся пустым.
    r.weatherCity || "",

    // Все поля из формы «Добавить улов».
    r.catchCities || "",
    r.catchFish || "",
    r.catchQuantities || "",
    r.catchTimes || "",
    r.catchTackles || "",
    r.catchMethods || "",
    r.catchPlaces || "",
    r.catchBaits || ""
  ];
}


/**
 * Явно синхронизирует комментарий выбранных дат.
 *
 * Почему это отдельная функция:
 * - комментарий является полем дня, а не улова;
 * - старые версии могли оставить дубликаты строк одной даты;
 * - обновляем все совпавшие строки в разрешённом окне;
 * - если строки нет, её создаёт основной writeYearSheets_().
 */
function writeCommentsForCorrectionWindow_(
  ss,
  years,
  correctionStartIso,
  todayIso
) {
  const tz = ss.getSpreadsheetTimeZone();

  years.forEach(function(yearData) {
    const sheet = ss.getSheetByName(
      String(yearData.year)
    );

    if (!sheet || sheet.getLastRow() < 2) {
      return;
    }

    const sourceRows = Array.isArray(yearData.rows)
      ? yearData.rows
      : [];

    const commentByDate = {};

    sourceRows.forEach(function(row) {
      if (
        row.date &&
        row.date >= correctionStartIso &&
        row.date <= todayIso
      ) {
        commentByDate[row.date] =
          row.comment === null ||
          typeof row.comment === "undefined"
            ? ""
            : String(row.comment);
      }
    });

    const lastRow = sheet.getLastRow();

    const sheetDates = sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getValues();

    sheetDates.forEach(function(row, index) {
      const iso = cellDateToIso_(
        row[0],
        tz
      );

      if (
        iso &&
        Object.prototype.hasOwnProperty.call(
          commentByDate,
          iso
        )
      ) {
        // Колонка 6 = "Комментарий"
        sheet
          .getRange(
            index + 2,
            6
          )
          .setValue(
            commentByDate[iso]
          )
          .setFontFamily("Arial");
      }
    });
  });
}


/**
 * Явно записывает город, по которому получена погода.
 *
 * Колонка "Город погоды" находится в последнем столбце YEAR_HEADERS.
 * Значение берётся из weatherCity, сформированного сайтом из weather.city
 * и weather.region. Поэтому город формы не подменяет фактический город погоды.
 */
function writeWeatherCityForCorrectionWindow_(
  ss,
  years,
  correctionStartIso,
  todayIso
) {
  const tz = ss.getSpreadsheetTimeZone();
  const cityColumn = YEAR_HEADERS.indexOf("Город погоды") + 1;

  if (cityColumn <= 0) {
    return;
  }

  years.forEach(function(yearData) {
    const sheet = ss.getSheetByName(
      String(yearData.year)
    );

    if (!sheet || sheet.getLastRow() < 2) {
      return;
    }

    const sourceRows = Array.isArray(yearData.rows)
      ? yearData.rows
      : [];

    const cityByDate = {};

    sourceRows.forEach(function(row) {
      if (
        row.date &&
        row.date >= correctionStartIso &&
        row.date <= todayIso
      ) {
        cityByDate[row.date] =
          row.weatherCity === null ||
          typeof row.weatherCity === "undefined"
            ? ""
            : String(row.weatherCity);
      }
    });

    const lastRow = sheet.getLastRow();

    const sheetDates = sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getValues();

    sheetDates.forEach(function(row, index) {
      const iso = cellDateToIso_(
        row[0],
        tz
      );

      if (
        iso &&
        Object.prototype.hasOwnProperty.call(
          cityByDate,
          iso
        )
      ) {
        sheet
          .getRange(
            index + 2,
            cityColumn
          )
          .setValue(
            cityByDate[iso]
          )
          .setFontFamily("Arial");
      }
    });
  });
}


/**
 * Записывает данные формы «Добавить улов» прямо в годовой лист.
 *
 * На годовом листе одна строка = один день, поэтому при нескольких
 * записях улова значения объединяются через " | ".
 * Детальная структура "одна поимка = одна строка" сохраняется на листе "Улов".
 */
function writeCatchFieldsForCorrectionWindow_(
  ss,
  years,
  correctionStartIso,
  todayIso
) {
  const tz = ss.getSpreadsheetTimeZone();

  const columns = {
    catchCities: YEAR_HEADERS.indexOf("Город улова") + 1,
    catchFish: YEAR_HEADERS.indexOf("Рыба") + 1,
    catchQuantities: YEAR_HEADERS.indexOf("Количество по записям") + 1,
    catchTimes: YEAR_HEADERS.indexOf("Время улова") + 1,
    catchTackles: YEAR_HEADERS.indexOf("Снасть") + 1,
    catchMethods: YEAR_HEADERS.indexOf("Способ ловли") + 1,
    catchPlaces: YEAR_HEADERS.indexOf("Место ловли") + 1,
    catchBaits: YEAR_HEADERS.indexOf("Приманка") + 1
  };

  years.forEach(function(yearData) {
    const sheet = ss.getSheetByName(
      String(yearData.year)
    );

    if (!sheet || sheet.getLastRow() < 2) {
      return;
    }

    const sourceRows = Array.isArray(yearData.rows)
      ? yearData.rows
      : [];

    const byDate = {};

    sourceRows.forEach(function(row) {
      if (
        row.date &&
        row.date >= correctionStartIso &&
        row.date <= todayIso
      ) {
        byDate[row.date] = {
          catchCities: row.catchCities || "",
          catchFish: row.catchFish || "",
          catchQuantities: row.catchQuantities || "",
          catchTimes: row.catchTimes || "",
          catchTackles: row.catchTackles || "",
          catchMethods: row.catchMethods || "",
          catchPlaces: row.catchPlaces || "",
          catchBaits: row.catchBaits || ""
        };
      }
    });

    const sheetDates = sheet
      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        1
      )
      .getValues();

    sheetDates.forEach(function(dateRow, index) {
      const iso = cellDateToIso_(
        dateRow[0],
        tz
      );

      if (
        !iso ||
        !Object.prototype.hasOwnProperty.call(
          byDate,
          iso
        )
      ) {
        return;
      }

      const values = byDate[iso];

      Object.keys(columns).forEach(function(key) {
        const column = columns[key];

        if (column > 0) {
          sheet
            .getRange(
              index + 2,
              column
            )
            .setValue(
              values[key]
            )
            .setFontFamily("Arial");
        }
      });
    });
  });
}


/**
 * Лист "Улов".
 *
 * Для последних 7 дней + сегодня:
 * - удаляем старую детализацию;
 * - записываем актуальные данные с сайта заново.
 *
 * Для совершенно новых дат:
 * - просто добавляем улов.
 *
 * Всё, что старше окна корректировки и уже находится
 * в Google Таблице, остаётся неизменным.
 */
function writeCatches_(
  ss,
  years,
  syncInfo,
  todayIso,
  correctionStartIso
) {
  const tz = ss.getSpreadsheetTimeZone();

  const sheet = getOrCreateSheet_(
    ss,
    "Улов"
  );

  ensureHeaders_(
    sheet,
    CATCH_HEADERS
  );

  // Разрешаем пересобрать улов только за последние 7 дней + сегодня.
  deleteCatchRowsForRange_(
    sheet,
    correctionStartIso,
    todayIso,
    tz
  );

  const rowsToAppend = [];

  years.forEach(function(yearData) {
    const year = String(yearData.year);

    const newDates = syncInfo[year]
      ? syncInfo[year].newDates
      : [];

    const allowedDates = new Set(
      newDates
    );

    // В окне корректировки разрешаем актуальную запись независимо
    // от того, была ли эта дата "новой".
    dateRangeIso_(
      correctionStartIso,
      todayIso,
      tz
    ).forEach(function(date) {
      allowedDates.add(date);
    });

    const catches = Array.isArray(yearData.catches)
      ? yearData.catches
      : [];

    catches.forEach(function(c) {
      if (!c.date) return;
      if (c.date > todayIso) return;
      if (!allowedDates.has(c.date)) return;

      rowsToAppend.push([
        yearData.year || "",
        parseDate_(c.date),
        c.time || "",
        c.fish || "",
        number_(c.qty),
        c.tackle || "",
        c.method || "",
        c.place || "",
        Array.isArray(c.baits)
          ? c.baits.join(", ")
          : (c.baits || ""),
        c.city || ""
      ]);
    });
  });


  rowsToAppend.sort(function(a, b) {
    const da = a[1] instanceof Date
      ? a[1].getTime()
      : 0;

    const db = b[1] instanceof Date
      ? b[1].getTime()
      : 0;

    if (da !== db) {
      return da - db;
    }

    return String(a[2])
      .localeCompare(
        String(b[2])
      );
  });


  if (rowsToAppend.length) {
    const startRow = sheet.getLastRow() + 1;

    sheet
      .getRange(
        startRow,
        1,
        rowsToAppend.length,
        CATCH_HEADERS.length
      )
      .setValues(
        rowsToAppend
      );

    sheet
      .getRange(
        startRow,
        2,
        rowsToAppend.length,
        1
      )
      .setNumberFormat(
        "dd.mm.yyyy"
      );

    formatDataRows_(
      sheet,
      startRow,
      rowsToAppend.length,
      CATCH_HEADERS.length
    );
  }


  formatHeader_(
    sheet,
    CATCH_HEADERS.length
  );

  ensureFilter_(
    sheet,
    CATCH_HEADERS.length
  );
}


/**
 * Сводка пересчитывается по фактическим данным,
 * уже находящимся в Google Таблице.
 *
 * Старые строки при этом не меняются —
 * меняется только итоговая сводка года.
 */
function updateSummary_(ss, years, syncInfo) {
  const sheet = getOrCreateSheet_(
    ss,
    "Сводка"
  );

  ensureHeaders_(
    sheet,
    SUMMARY_HEADERS
  );

  years.forEach(function(yearData) {
    const year = String(yearData.year);

    const changed = syncInfo[year] &&
      (
        syncInfo[year].newDates.length > 0 ||
        syncInfo[year].correctionDates.length > 0
      );

    const existingRow = findSummaryYearRow_(
      sheet,
      Number(yearData.year)
    );

    if (existingRow && !changed) {
      return;
    }

    const summary = calculateSummaryFromSheets_(
      ss,
      Number(yearData.year)
    );

    const values = [[
      Number(yearData.year),
      summary.fishingDays,
      summary.total,
      summary.species,
      summary.bestDay
        ? parseDate_(summary.bestDay)
        : "",
      summary.bestTotal,
      new Date()
    ]];


    if (existingRow) {
      sheet
        .getRange(
          existingRow,
          1,
          1,
          SUMMARY_HEADERS.length
        )
        .setValues(values);

      sheet
        .getRange(existingRow, 5)
        .setNumberFormat("dd.mm.yyyy");

      sheet
        .getRange(existingRow, 7)
        .setNumberFormat("dd.mm.yyyy hh:mm");

      formatDataRows_(
        sheet,
        existingRow,
        1,
        SUMMARY_HEADERS.length
      );

    } else {
      const row = sheet.getLastRow() + 1;

      sheet
        .getRange(
          row,
          1,
          1,
          SUMMARY_HEADERS.length
        )
        .setValues(values);

      sheet
        .getRange(row, 5)
        .setNumberFormat("dd.mm.yyyy");

      sheet
        .getRange(row, 7)
        .setNumberFormat("dd.mm.yyyy hh:mm");

      formatDataRows_(
        sheet,
        row,
        1,
        SUMMARY_HEADERS.length
      );
    }
  });


  formatHeader_(
    sheet,
    SUMMARY_HEADERS.length
  );
}


function calculateSummaryFromSheets_(ss, year) {
  const yearSheet = ss.getSheetByName(
    String(year)
  );

  let fishingDays = 0;
  let total = 0;
  let bestTotal = 0;
  let bestDay = "";

  if (
    yearSheet &&
    yearSheet.getLastRow() >= 2
  ) {
    const data = yearSheet
      .getRange(
        2,
        1,
        yearSheet.getLastRow() - 1,
        4
      )
      .getValues();

    data.forEach(function(row) {
      const date = row[0];

      const fishing =
        String(row[2]).toLowerCase() === "да";

      const qty = number_(row[3]);

      if (fishing) {
        fishingDays++;
      }

      total += qty;

      if (qty > bestTotal) {
        bestTotal = qty;

        bestDay =
          date instanceof Date
            ? dateToIso_(
                date,
                ss.getSpreadsheetTimeZone()
              )
            : "";
      }
    });
  }


  const species = new Set();

  const catchSheet = ss.getSheetByName(
    "Улов"
  );

  if (
    catchSheet &&
    catchSheet.getLastRow() >= 2
  ) {
    const catches = catchSheet
      .getRange(
        2,
        1,
        catchSheet.getLastRow() - 1,
        4
      )
      .getValues();

    catches.forEach(function(row) {
      if (
        Number(row[0]) === Number(year) &&
        row[3]
      ) {
        species.add(
          String(row[3])
        );
      }
    });
  }


  return {
    fishingDays: fishingDays,
    total: total,
    species: species.size,
    bestDay: bestDay,
    bestTotal: bestTotal
  };
}


/**
 * Удаляем строки листа "Улов" только внутри разрешённого окна.
 */
function deleteCatchRowsForRange_(
  sheet,
  fromIso,
  toIso,
  timeZone
) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return;
  }

  const dates = sheet
    .getRange(
      2,
      2,
      lastRow - 1,
      1
    )
    .getValues();

  // Снизу вверх, чтобы номера строк не сдвигались.
  for (
    let i = dates.length - 1;
    i >= 0;
    i--
  ) {
    const iso = cellDateToIso_(
      dates[i][0],
      timeZone
    );

    if (
      iso &&
      iso >= fromIso &&
      iso <= toIso
    ) {
      sheet.deleteRow(i + 2);
    }
  }
}


/**
 * Удаляет только будущие строки,
 * созданные старыми версиями приложения.
 */
function removeFutureRows_(
  sheet,
  dateColumn,
  todayIso,
  timeZone
) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return;
  }

  const values = sheet
    .getRange(
      2,
      dateColumn,
      lastRow - 1,
      1
    )
    .getValues();

  const rowsToDelete = [];

  values.forEach(function(row, index) {
    const iso = cellDateToIso_(
      row[0],
      timeZone
    );

    if (
      iso &&
      iso > todayIso
    ) {
      rowsToDelete.push(
        index + 2
      );
    }
  });

  // Удаляем снизу вверх.
  rowsToDelete
    .reverse()
    .forEach(function(rowNumber) {
      sheet.deleteRow(rowNumber);
    });
}


function getLastDateInfo_(
  sheet,
  dateColumn,
  timeZone
) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return null;
  }

  const values = sheet
    .getRange(
      2,
      dateColumn,
      lastRow - 1,
      1
    )
    .getValues();

  let best = null;

  values.forEach(function(row, index) {
    const iso = cellDateToIso_(
      row[0],
      timeZone
    );

    if (
      iso &&
      (
        !best ||
        iso > best.iso
      )
    ) {
      best = {
        iso: iso,
        row: index + 2
      };
    }
  });

  return best;
}


function findDateRow_(
  sheet,
  dateColumn,
  isoDate,
  timeZone
) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return null;
  }

  const values = sheet
    .getRange(
      2,
      dateColumn,
      lastRow - 1,
      1
    )
    .getValues();

  for (
    let i = 0;
    i < values.length;
    i++
  ) {
    if (
      cellDateToIso_(
        values[i][0],
        timeZone
      ) === isoDate
    ) {
      return {
        row: i + 2,
        iso: isoDate
      };
    }
  }

  return null;
}


function findSummaryYearRow_(
  sheet,
  year
) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return null;
  }

  const values = sheet
    .getRange(
      2,
      1,
      lastRow - 1,
      1
    )
    .getValues();

  for (
    let i = 0;
    i < values.length;
    i++
  ) {
    if (
      Number(values[i][0]) ===
      Number(year)
    ) {
      return i + 2;
    }
  }

  return null;
}


function ensureHeaders_(
  sheet,
  headers
) {
  sheet
    .getRange(
      1,
      1,
      1,
      headers.length
    )
    .setValues([
      headers
    ]);

  formatHeader_(
    sheet,
    headers.length
  );
}


function formatHeader_(
  sheet,
  columnCount
) {
  sheet
    .getRange(
      1,
      1,
      1,
      columnCount
    )
    .setFontFamily("Arial")
    .setBackground("#F4DC4F")
    .setFontColor("#000000")
    .setFontWeight("bold")
    .setVerticalAlignment("middle");

  sheet.setFrozenRows(1);
}


function formatDataRows_(
  sheet,
  startRow,
  rowCount,
  columnCount
) {
  if (rowCount <= 0) {
    return;
  }

  sheet
    .getRange(
      startRow,
      1,
      rowCount,
      columnCount
    )
    .setFontFamily("Arial")
    .setVerticalAlignment("middle");
}


function ensureFilter_(
  sheet,
  columnCount
) {
  const existing =
    sheet.getFilter();

  if (existing) {
    existing.remove();
  }

  const lastRow =
    sheet.getLastRow();

  if (lastRow >= 2) {
    sheet
      .getRange(
        1,
        1,
        lastRow,
        columnCount
      )
      .createFilter();
  }
}


function getOrCreateSheet_(
  ss,
  name
) {
  return (
    ss.getSheetByName(name) ||
    ss.insertSheet(name)
  );
}


function parseDate_(isoDate) {
  if (!isoDate) {
    return "";
  }

  const parts = String(isoDate)
    .split("-")
    .map(Number);

  if (
    parts.length !== 3 ||
    !parts[0] ||
    !parts[1] ||
    !parts[2]
  ) {
    return isoDate;
  }

  // Полдень уменьшает риск сдвига даты
  // из-за часового пояса.
  return new Date(
    parts[0],
    parts[1] - 1,
    parts[2],
    12,
    0,
    0
  );
}


function cellDateToIso_(
  value,
  timeZone
) {
  if (
    value instanceof Date &&
    !isNaN(value.getTime())
  ) {
    return dateToIso_(
      value,
      timeZone
    );
  }

  const text =
    String(value || "")
      .trim();

  if (
    /^\d{4}-\d{2}-\d{2}$/.test(text)
  ) {
    return text;
  }

  const match = text.match(
    /^(\d{1,2})[./](\d{1,2})[./](\d{4})$/
  );

  if (match) {
    return [
      match[3],
      String(match[2]).padStart(2, "0"),
      String(match[1]).padStart(2, "0")
    ].join("-");
  }

  return "";
}


function isIsoDate_(value) {
  return /^\d{4}-\d{2}-\d{2}$/.test(
    String(value || "")
  );
}


function dateToIso_(
  date,
  timeZone
) {
  return Utilities.formatDate(
    date,
    timeZone ||
      Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );
}


/**
 * Сдвигает ISO-дату на указанное число календарных дней.
 */
function shiftIsoDate_(
  isoDate,
  days,
  timeZone
) {
  const parts = isoDate
    .split("-")
    .map(Number);

  const date = new Date(
    parts[0],
    parts[1] - 1,
    parts[2],
    12,
    0,
    0
  );

  date.setDate(
    date.getDate() + days
  );

  return dateToIso_(
    date,
    timeZone
  );
}


/**
 * Возвращает все ISO-даты в диапазоне включительно.
 */
function dateRangeIso_(
  fromIso,
  toIso,
  timeZone
) {
  const result = [];

  let current = fromIso;

  while (
    current <= toIso
  ) {
    result.push(current);

    current = shiftIsoDate_(
      current,
      1,
      timeZone
    );
  }

  return result;
}


function unique_(array) {
  return Array.from(
    new Set(array)
  );
}


function value_(value) {
  if (
    value === "" ||
    value === null ||
    typeof value === "undefined"
  ) {
    return "";
  }

  return value;
}


function number_(value) {
  const number = Number(value);

  return isNaN(number)
    ? 0
    : number;
}


function jsonResponse_(data) {
  return ContentService
    .createTextOutput(
      JSON.stringify(data)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}


/**
 * Проверка подключения и окна корректировки.
 */
function testConnection() {
  const ss =
    SpreadsheetApp.openById(
      SPREADSHEET_ID
    );

  const tz =
    ss.getSpreadsheetTimeZone();

  const today =
    dateToIso_(
      new Date(),
      tz
    );

  const from =
    shiftIsoDate_(
      today,
      -CORRECTION_DAYS_BACK,
      tz
    );

  console.log(
    "Подключение успешно."
  );

  console.log(
    "Таблица: " +
    ss.getName()
  );

  console.log(
    "Окно корректировки: " +
    from +
    " — " +
    today
  );
}
