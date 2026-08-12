/**
 * Рыболовный дневник -> Google Sheets
 * Версия 19.
 *
 * Главная задача версии:
 * привести старые, новые и отредактированные строки к ОДНОЙ структуре.
 *
 * Что делает скрипт:
 * - один раз создаёт скрытые резервные копии рабочих листов;
 * - мигрирует старые схемы колонок в текущую каноническую;
 * - не меняет смысл исторических значений;
 * - выравнивает форматирование всех строк;
 * - новые даты добавляет после последней записи;
 * - сегодня + 7 предыдущих дней разрешает редактировать;
 * - более старые даты после миграции не перезаписывает;
 * - данные формы "Добавить улов" сохраняет и в листе года, и в листе "Улов".
 */

const SPREADSHEET_ID = "1H7H2AUwtfeqYaaWE0LCZq4QFm7o5Rp7M8eYg70uJLkM";
const CORRECTION_DAYS_BACK = 7;
const MIGRATION_PROPERTY = "FISHING_DAY_UNIFIED_FORMAT_V19";

const YEAR_HEADERS = [
  "Дата",
  "День недели",
  "Рыбалка",
  "Количество",
  "Результат дня",
  "Комментарий",

  "Температура, °C",
  "Ощущается, °C",
  "Погода",
  "Направление ветра",
  "Ветер, °",
  "Скорость ветра, км/ч",
  "Порывы ветра, км/ч",
  "Влажность, %",
  "Давление, hPa",
  "Давление, мм рт. ст.",
  "Осадки, мм",
  "Облачность, %",
  "Город погоды",

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
  "Год",
  "Дата",
  "Время",
  "Рыба",
  "Количество",
  "Снасть",
  "Способ ловли",
  "Место ловли",
  "Приманка",
  "Город"
];

const SUMMARY_HEADERS = [
  "Год",
  "Дней с рыбалкой",
  "Всего поймано",
  "Видов рыбы",
  "Лучший день",
  "Улов в лучший день",
  "Последняя синхронизация"
];

const YEAR_ALIASES = {
  "Дата": ["Дата"],
  "День недели": ["День недели"],
  "Рыбалка": ["Рыбалка"],
  "Количество": ["Количество"],
  "Результат дня": ["Результат дня"],
  "Комментарий": ["Комментарий"],

  "Температура, °C": ["Температура, °C", "Температура °C"],
  "Ощущается, °C": ["Ощущается, °C", "Ощущается °C"],
  "Погода": ["Погода"],
  "Направление ветра": ["Направление ветра"],
  "Ветер, °": ["Ветер, °", "Ветер °"],
  "Скорость ветра, км/ч": ["Скорость ветра, км/ч", "Скорость ветра км/ч"],
  "Порывы ветра, км/ч": ["Порывы ветра, км/ч", "Порывы ветра км/ч", "Порывы км/ч"],
  "Влажность, %": ["Влажность, %", "Влажность %"],
  "Давление, hPa": ["Давление, hPa", "Давление hPa"],
  "Давление, мм рт. ст.": [
    "Давление, мм рт. ст.",
    "Давление мм рт. ст.",
    "Давление мм рт ст"
  ],
  "Осадки, мм": ["Осадки, мм", "Осадки мм"],
  "Облачность, %": ["Облачность, %", "Облачность %"],
  "Город погоды": ["Город погоды"],

  "Город улова": ["Город улова"],
  "Рыба": ["Рыба"],
  "Количество по записям": ["Количество по записям"],
  "Время улова": ["Время улова"],
  "Снасть": ["Снасть"],
  "Способ ловли": ["Способ ловли"],
  "Место ловли": ["Место ловли"],
  "Приманка": ["Приманка"]
};

const CATCH_ALIASES = {
  "Год": ["Год"],
  "Дата": ["Дата"],
  "Время": ["Время"],
  "Рыба": ["Рыба"],
  "Количество": ["Количество"],
  "Снасть": ["Снасть"],
  "Способ ловли": ["Способ ловли"],
  "Место ловли": ["Место ловли"],
  "Приманка": ["Приманка"],
  "Город": ["Город"]
};


function doGet() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const tz = ss.getSpreadsheetTimeZone();
    const todayIso = dateToIso_(new Date(), tz);

    return jsonResponse_({
      ok: true,
      service: "Fishing Day Sheets Sync",
      version: 19,
      schema: "unified-v19",
      spreadsheet: ss.getName(),
      today: todayIso,
      correctionFrom: shiftIsoDate_(
        todayIso,
        -CORRECTION_DAYS_BACK,
        tz
      )
    });

  } catch (error) {
    return jsonResponse_({
      ok: false,
      version: 19,
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
    const years = Array.isArray(payload.years)
      ? payload.years
      : [];

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const tz = ss.getSpreadsheetTimeZone();

    const todayIso = isIsoDate_(payload.clientToday)
      ? payload.clientToday
      : dateToIso_(new Date(), tz);

    const correctionStartIso = shiftIsoDate_(
      todayIso,
      -CORRECTION_DAYS_BACK,
      tz
    );

    // Первая синхронизация версии 19:
    // резервная копия + миграция старой структуры.
    ensureUnifiedMigrationV19_(ss);

    // На каждой синхронизации выполняется лёгкая проверка схемы.
    normalizeWorkingSheets_(ss);

    const syncInfo = writeYearSheets_(
      ss,
      years,
      todayIso,
      correctionStartIso
    );

    writeCatches_(
      ss,
      years,
      syncInfo,
      todayIso,
      correctionStartIso
    );

    // Годовые поля улова формируем из детального листа "Улов".
    // Так старые и новые строки выглядят одинаково.
    rebuildYearCatchColumnsFromCatchSheet_(ss);

    updateSummary_(
      ss,
      years,
      syncInfo
    );

    applyUnifiedFormatting_(ss);
    SpreadsheetApp.flush();

    return jsonResponse_({
      ok: true,
      version: 19,
      schema: "unified-v19",
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
    console.error(
      "SYNC ERROR V19: " +
      String(error)
    );

    return jsonResponse_({
      ok: false,
      version: 19,
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
 * Выполняется автоматически один раз.
 * Создаёт резервные копии и приводит старые листы к единой схеме.
 */
function ensureUnifiedMigrationV19_(ss) {
  const properties = PropertiesService.getScriptProperties();

  if (
    properties.getProperty(
      MIGRATION_PROPERTY
    ) === "done"
  ) {
    return;
  }

  createMigrationBackups_(ss);

  const yearSheets = getYearSheets_(ss);

  yearSheets.forEach(function(sheet) {
    normalizeYearSheet_(sheet, ss);
  });

  normalizeCatchSheet_(
    getOrCreateSheet_(ss, "Улов"),
    ss
  );

  normalizeSummarySheet_(
    getOrCreateSheet_(ss, "Сводка")
  );

  rebuildYearCatchColumnsFromCatchSheet_(ss);
  applyUnifiedFormatting_(ss);

  properties.setProperty(
    MIGRATION_PROPERTY,
    "done"
  );
}


/**
 * Ручная функция.
 * Можно один раз выбрать её сверху в Apps Script и нажать "Выполнить".
 *
 * Она:
 * - создаёт резервные копии, если их ещё нет;
 * - заново приводит рабочие листы к канонической структуре;
 * - не удаляет исходные резервные копии.
 */
function migrateGoogleSheetsToUnifiedFormat() {
  const ss = SpreadsheetApp.openById(
    SPREADSHEET_ID
  );

  createMigrationBackups_(ss);

  getYearSheets_(ss).forEach(function(sheet) {
    normalizeYearSheet_(sheet, ss);
  });

  normalizeCatchSheet_(
    getOrCreateSheet_(ss, "Улов"),
    ss
  );

  normalizeSummarySheet_(
    getOrCreateSheet_(ss, "Сводка")
  );

  rebuildYearCatchColumnsFromCatchSheet_(ss);
  applyUnifiedFormatting_(ss);

  PropertiesService
    .getScriptProperties()
    .setProperty(
      MIGRATION_PROPERTY,
      "done"
    );

  SpreadsheetApp.flush();

  console.log(
    "Миграция в единый формат v19 завершена."
  );
}


/**
 * Если нужно повторно прогнать миграцию после ручных изменений,
 * запустите эту функцию, а затем migrateGoogleSheetsToUnifiedFormat().
 */
function resetUnifiedMigrationFlag() {
  PropertiesService
    .getScriptProperties()
    .deleteProperty(
      MIGRATION_PROPERTY
    );
}


/**
 * Резервные копии создаются только один раз и скрываются.
 */
function createMigrationBackups_(ss) {
  const names = ["Улов", "Сводка"];

  getYearSheets_(ss).forEach(function(sheet) {
    names.push(
      sheet.getName()
    );
  });

  names.forEach(function(name) {
    const source = ss.getSheetByName(name);

    if (!source) {
      return;
    }

    const backupName =
      "_backup_v18_" + name;

    if (
      ss.getSheetByName(
        backupName
      )
    ) {
      return;
    }

    const backup = source
      .copyTo(ss)
      .setName(
        backupName
      );

    try {
      backup.hideSheet();
    } catch (ignore) {}
  });
}


/**
 * Проверка структуры при каждой синхронизации.
 * Никакие резервные копии здесь уже не создаются.
 */
function normalizeWorkingSheets_(ss) {
  getYearSheets_(ss).forEach(function(sheet) {
    normalizeYearSheet_(sheet, ss);
  });

  normalizeCatchSheet_(
    getOrCreateSheet_(ss, "Улов"),
    ss
  );

  normalizeSummarySheet_(
    getOrCreateSheet_(ss, "Сводка")
  );
}


/**
 * Миграция годового листа.
 *
 * Значения раскладываются по названиям колонок, а не по старым номерам.
 * Это исправляет ситуации, когда старые строки были записаны по одной
 * схеме, а заголовок позднее был изменён.
 */
function normalizeYearSheet_(sheet, ss) {
  const lastRow = Math.max(
    sheet.getLastRow(),
    1
  );

  const lastColumn = Math.max(
    sheet.getLastColumn(),
    1
  );

  ensureSheetCapacity_(
    sheet,
    lastRow,
    YEAR_HEADERS.length
  );

  if (lastRow === 1) {
    sheet
      .getRange(
        1,
        1,
        1,
        YEAR_HEADERS.length
      )
      .setValues([
        YEAR_HEADERS
      ]);

    return;
  }

  const raw = sheet
    .getRange(
      1,
      1,
      lastRow,
      Math.max(
        lastColumn,
        YEAR_HEADERS.length
      )
    )
    .getValues();

  const oldHeaders = raw[0];
  const indexMap = buildCanonicalIndex_(
    oldHeaders,
    YEAR_HEADERS,
    YEAR_ALIASES
  );

  const canonicalRows = [];

  for (
    let r = 1;
    r < raw.length;
    r++
  ) {
    const source = raw[r];

    if (
      source.every(
        isEmptyValue_
      )
    ) {
      continue;
    }

    let row = YEAR_HEADERS.map(function(header) {
      const index = indexMap[header];

      return index >= 0
        ? source[index]
        : "";
    });

    // Если заголовок уже был перезаписан новой версией,
    // но сама строка осталась из самой ранней схемы без "Комментария",
    // распознаём характерный сдвиг погодных колонок и исправляем его.
    row = repairLegacyNoCommentShift_(
      row
    );

    canonicalRows.push(
      row
    );
  }

  const mergedRows = mergeYearRowsByDate_(
    canonicalRows,
    ss.getSpreadsheetTimeZone()
  );

  const clearColumns = Math.max(
    lastColumn,
    YEAR_HEADERS.length
  );

  sheet
    .getRange(
      1,
      1,
      Math.max(
        lastRow,
        mergedRows.length + 1
      ),
      clearColumns
    )
    .clearContent();

  ensureSheetCapacity_(
    sheet,
    mergedRows.length + 1,
    YEAR_HEADERS.length
  );

  sheet
    .getRange(
      1,
      1,
      1,
      YEAR_HEADERS.length
    )
    .setValues([
      YEAR_HEADERS
    ]);

  if (mergedRows.length) {
    sheet
      .getRange(
        2,
        1,
        mergedRows.length,
        YEAR_HEADERS.length
      )
      .setValues(
        mergedRows
      );
  }
}


/**
 * Старейшая схема не содержала колонку "Комментарий".
 * Если заголовок уже был заменён поздней версией, температура могла
 * визуально оказаться под "Комментарий". Этот сдвиг исправляется здесь.
 */
function repairLegacyNoCommentShift_(row) {
  const comment = row[5];
  const temperature = row[6];
  const apparent = row[7];

  const looksShifted =
    isNumberLike_(comment) &&
    isNumberLike_(temperature) &&
    typeof apparent === "string" &&
    apparent !== "" &&
    !isNumberLike_(apparent);

  if (!looksShifted) {
    return row;
  }

  const fixed = row.slice();

  // Старые F:Q -> новые G:R.
  for (
    let target = 17;
    target >= 6;
    target--
  ) {
    fixed[target] =
      row[target - 1];
  }

  fixed[5] = "";

  return fixed;
}


/**
 * Если после старых версий появились дубли даты,
 * оставляем одну строку и объединяем непустые значения.
 * Более поздняя строка имеет приоритет.
 */
function mergeYearRowsByDate_(
  rows,
  timeZone
) {
  const byDate = {};
  const withoutDate = [];

  rows.forEach(function(row) {
    const iso = cellDateToIso_(
      row[0],
      timeZone
    );

    if (!iso) {
      withoutDate.push(row);
      return;
    }

    if (!byDate[iso]) {
      byDate[iso] =
        row.slice();
      return;
    }

    const target = byDate[iso];

    row.forEach(function(value, index) {
      if (!isEmptyValue_(value)) {
        target[index] = value;
      }
    });
  });

  const datedRows = Object.keys(
    byDate
  )
    .sort()
    .map(function(key) {
      return byDate[key];
    });

  return datedRows.concat(
    withoutDate
  );
}


/**
 * Приведение листа "Улов" к одной схеме.
 * Старые версии имели 5, 9 или 10 колонок — значения сохраняются.
 */
function normalizeCatchSheet_(sheet, ss) {
  const lastRow = Math.max(
    sheet.getLastRow(),
    1
  );

  const lastColumn = Math.max(
    sheet.getLastColumn(),
    1
  );

  ensureSheetCapacity_(
    sheet,
    lastRow,
    CATCH_HEADERS.length
  );

  if (lastRow === 1) {
    sheet
      .getRange(
        1,
        1,
        1,
        CATCH_HEADERS.length
      )
      .setValues([
        CATCH_HEADERS
      ]);

    return;
  }

  const raw = sheet
    .getRange(
      1,
      1,
      lastRow,
      Math.max(
        lastColumn,
        CATCH_HEADERS.length
      )
    )
    .getValues();

  const indexMap = buildCanonicalIndex_(
    raw[0],
    CATCH_HEADERS,
    CATCH_ALIASES
  );

  const rows = [];

  for (
    let r = 1;
    r < raw.length;
    r++
  ) {
    const source = raw[r];

    if (
      source.every(
        isEmptyValue_
      )
    ) {
      continue;
    }

    const row = CATCH_HEADERS.map(function(header) {
      const index = indexMap[header];

      return index >= 0
        ? source[index]
        : "";
    });

    rows.push(row);
  }

  rows.sort(function(a, b) {
    const aIso = cellDateToIso_(
      a[1],
      ss.getSpreadsheetTimeZone()
    );

    const bIso = cellDateToIso_(
      b[1],
      ss.getSpreadsheetTimeZone()
    );

    if (aIso !== bIso) {
      return aIso.localeCompare(
        bIso
      );
    }

    return String(a[2] || "")
      .localeCompare(
        String(b[2] || "")
      );
  });

  sheet
    .getRange(
      1,
      1,
      Math.max(
        lastRow,
        rows.length + 1
      ),
      Math.max(
        lastColumn,
        CATCH_HEADERS.length
      )
    )
    .clearContent();

  ensureSheetCapacity_(
    sheet,
    rows.length + 1,
    CATCH_HEADERS.length
  );

  sheet
    .getRange(
      1,
      1,
      1,
      CATCH_HEADERS.length
    )
    .setValues([
      CATCH_HEADERS
    ]);

  if (rows.length) {
    sheet
      .getRange(
        2,
        1,
        rows.length,
        CATCH_HEADERS.length
      )
      .setValues(rows);
  }
}


function normalizeSummarySheet_(sheet) {
  ensureSheetCapacity_(
    sheet,
    Math.max(
      sheet.getLastRow(),
      2
    ),
    SUMMARY_HEADERS.length
  );

  sheet
    .getRange(
      1,
      1,
      1,
      SUMMARY_HEADERS.length
    )
    .setValues([
      SUMMARY_HEADERS
    ]);
}


/**
 * Находит старую колонку по названию/алиасу.
 */
function buildCanonicalIndex_(
  oldHeaders,
  canonicalHeaders,
  aliases
) {
  const normalizedOld = {};

  oldHeaders.forEach(function(header, index) {
    const key = normalizeHeader_(
      header
    );

    if (
      key &&
      typeof normalizedOld[key] === "undefined"
    ) {
      normalizedOld[key] =
        index;
    }
  });

  const result = {};

  canonicalHeaders.forEach(function(header) {
    const candidates =
      aliases[header] ||
      [header];

    let found = -1;

    candidates.some(function(candidate) {
      const key = normalizeHeader_(
        candidate
      );

      if (
        typeof normalizedOld[key] !==
        "undefined"
      ) {
        found =
          normalizedOld[key];

        return true;
      }

      return false;
    });

    result[header] =
      found;
  });

  return result;
}


function normalizeHeader_(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/ё/g, "е")
    .replace(
      /[^a-zа-я0-9]+/gi,
      ""
    );
}


/**
 * Добавление новых дат + обновление только разрешённого окна.
 */
function writeYearSheets_(
  ss,
  years,
  todayIso,
  correctionStartIso
) {
  const tz = ss.getSpreadsheetTimeZone();
  const info = {};

  years.forEach(function(yearData) {
    const year = String(
      yearData.year
    );

    const sheet = getOrCreateSheet_(
      ss,
      year
    );

    normalizeYearSheet_(
      sheet,
      ss
    );

    removeFutureRows_(
      sheet,
      todayIso,
      tz
    );

    const dateRows = buildDateRowMap_(
      sheet,
      tz
    );

    const lastDate = getLastDateFromMap_(
      dateRows
    );

    const sourceRows = Array.isArray(
      yearData.rows
    )
      ? yearData.rows.slice()
      : [];

    sourceRows.sort(function(a, b) {
      return String(a.date)
        .localeCompare(
          String(b.date)
        );
    });

    const newDates = [];
    const correctionDates = [];

    sourceRows.forEach(function(sourceRow) {
      const date = sourceRow.date;

      if (
        !date ||
        date > todayIso
      ) {
        return;
      }

      const existingRow =
        dateRows[date];

      const canCorrect =
        date >= correctionStartIso &&
        date <= todayIso;

      if (existingRow) {
        if (!canCorrect) {
          return;
        }

        sheet
          .getRange(
            existingRow,
            1,
            1,
            YEAR_HEADERS.length
          )
          .setValues([
            yearRowToValues_(
              sourceRow
            )
          ]);

        correctionDates.push(
          date
        );

        return;
      }

      // Новые даты добавляются только после последней существующей даты.
      // Если таблица пустая — добавляем всё до сегодня.
      if (
        !lastDate ||
        date > lastDate
      ) {
        const row =
          sheet.getLastRow() + 1;

        ensureSheetCapacity_(
          sheet,
          row,
          YEAR_HEADERS.length
        );

        sheet
          .getRange(
            row,
            1,
            1,
            YEAR_HEADERS.length
          )
          .setValues([
            yearRowToValues_(
              sourceRow
            )
          ]);

        dateRows[date] =
          row;

        newDates.push(
          date
        );
      }
    });

    formatYearSheet_(
      sheet
    );

    info[year] = {
      newDates: unique_(newDates),
      correctionDates: unique_(correctionDates),
      lastDate:
        getLastDateFromMap_(
          buildDateRowMap_(
            sheet,
            tz
          )
        )
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
    r.weatherCity || "",

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
 * Детальный лист "Улов".
 */
function writeCatches_(
  ss,
  years,
  syncInfo,
  todayIso,
  correctionStartIso
) {
  const sheet = getOrCreateSheet_(
    ss,
    "Улов"
  );

  normalizeCatchSheet_(
    sheet,
    ss
  );

  const tz =
    ss.getSpreadsheetTimeZone();

  // Пересобираем только разрешённое окно корректировки.
  deleteCatchRowsForRange_(
    sheet,
    correctionStartIso,
    todayIso,
    tz
  );

  const rowsToAppend = [];

  years.forEach(function(yearData) {
    const year = String(
      yearData.year
    );

    const allowedDates = new Set(
      syncInfo[year]
        ? syncInfo[year].newDates
        : []
    );

    dateRangeIso_(
      correctionStartIso,
      todayIso,
      tz
    ).forEach(function(date) {
      allowedDates.add(date);
    });

    const catches = Array.isArray(
      yearData.catches
    )
      ? yearData.catches
      : [];

    catches.forEach(function(c) {
      if (
        !c.date ||
        c.date > todayIso ||
        !allowedDates.has(c.date)
      ) {
        return;
      }

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
    const da = cellDateToIso_(
      a[1],
      tz
    );

    const db = cellDateToIso_(
      b[1],
      tz
    );

    if (da !== db) {
      return da.localeCompare(db);
    }

    return String(a[2] || "")
      .localeCompare(
        String(b[2] || "")
      );
  });

  if (rowsToAppend.length) {
    const startRow =
      sheet.getLastRow() + 1;

    ensureSheetCapacity_(
      sheet,
      startRow +
        rowsToAppend.length - 1,
      CATCH_HEADERS.length
    );

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
  }

  formatCatchSheet_(
    sheet
  );
}


/**
 * Заполняет колонки "Город улова" ... "Приманка"
 * на годовых листах из детального листа "Улов".
 *
 * Благодаря этому исторические и новые строки имеют одну структуру.
 */
function rebuildYearCatchColumnsFromCatchSheet_(ss) {
  const catchSheet = ss.getSheetByName(
    "Улов"
  );

  if (
    !catchSheet ||
    catchSheet.getLastRow() < 2
  ) {
    return;
  }

  normalizeCatchSheet_(
    catchSheet,
    ss
  );

  const tz =
    ss.getSpreadsheetTimeZone();

  const catchValues = catchSheet
    .getRange(
      2,
      1,
      catchSheet.getLastRow() - 1,
      CATCH_HEADERS.length
    )
    .getValues();

  const groups = {};

  catchValues.forEach(function(row) {
    const year = String(
      row[0] || ""
    );

    const date = cellDateToIso_(
      row[1],
      tz
    );

    if (
      !year ||
      !date
    ) {
      return;
    }

    const key =
      year + "|" + date;

    if (!groups[key]) {
      groups[key] = [];
    }

    groups[key].push(row);
  });

  getYearSheets_(ss).forEach(function(sheet) {
    const year =
      sheet.getName();

    normalizeYearSheet_(
      sheet,
      ss
    );

    if (
      sheet.getLastRow() < 2
    ) {
      return;
    }

    const dates = sheet
      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        1
      )
      .getValues();

    dates.forEach(function(dateRow, index) {
      const date = cellDateToIso_(
        dateRow[0],
        tz
      );

      if (!date) {
        return;
      }

      const catches =
        groups[
          year + "|" + date
        ];

      // Если детальных исторических записей нет,
      // уже существующие значения годового листа не затираем.
      if (
        !catches ||
        !catches.length
      ) {
        return;
      }

      const joined = [
        joinNonEmpty_(
          catches.map(
            function(c) {
              return c[9];
            }
          )
        ),
        joinNonEmpty_(
          catches.map(
            function(c) {
              return c[3];
            }
          )
        ),
        joinNonEmpty_(
          catches.map(
            function(c) {
              return c[4];
            }
          )
        ),
        joinNonEmpty_(
          catches.map(
            function(c) {
              return c[2];
            }
          )
        ),
        joinNonEmpty_(
          catches.map(
            function(c) {
              return c[5];
            }
          )
        ),
        joinNonEmpty_(
          catches.map(
            function(c) {
              return c[6];
            }
          )
        ),
        joinNonEmpty_(
          catches.map(
            function(c) {
              return c[7];
            }
          )
        ),
        joinNonEmpty_(
          catches.map(
            function(c) {
              return c[8];
            }
          )
        )
      ];

      // T:AA = колонки 20-27.
      sheet
        .getRange(
          index + 2,
          20,
          1,
          8
        )
        .setValues([
          joined
        ]);
    });

    formatYearSheet_(
      sheet
    );
  });
}


function joinNonEmpty_(values) {
  return values
    .filter(function(value) {
      return !isEmptyValue_(
        value
      );
    })
    .map(function(value) {
      return String(value);
    })
    .join(" | ");
}


/**
 * Сводка считается уже по приведённым к одной схеме листам.
 */
function updateSummary_(
  ss,
  years,
  syncInfo
) {
  const sheet = getOrCreateSheet_(
    ss,
    "Сводка"
  );

  normalizeSummarySheet_(
    sheet
  );

  years.forEach(function(yearData) {
    const year = String(
      yearData.year
    );

    const changed =
      syncInfo[year] &&
      (
        syncInfo[year].newDates.length > 0 ||
        syncInfo[year].correctionDates.length > 0
      );

    const existingRow =
      findSummaryYearRow_(
        sheet,
        Number(
          yearData.year
        )
      );

    if (
      existingRow &&
      !changed
    ) {
      return;
    }

    const summary =
      calculateSummaryFromSheets_(
        ss,
        Number(
          yearData.year
        )
      );

    const values = [[
      Number(yearData.year),
      summary.fishingDays,
      summary.total,
      summary.species,
      summary.bestDay
        ? parseDate_(
            summary.bestDay
          )
        : "",
      summary.bestTotal,
      new Date()
    ]];

    const row =
      existingRow ||
      sheet.getLastRow() + 1;

    ensureSheetCapacity_(
      sheet,
      row,
      SUMMARY_HEADERS.length
    );

    sheet
      .getRange(
        row,
        1,
        1,
        SUMMARY_HEADERS.length
      )
      .setValues(values);
  });

  formatSummarySheet_(
    sheet
  );
}


function calculateSummaryFromSheets_(
  ss,
  year
) {
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
      const fishing =
        String(row[2] || "")
          .toLowerCase() === "да";

      const qty =
        number_(row[3]);

      if (fishing) {
        fishingDays++;
      }

      total += qty;

      if (qty > bestTotal) {
        bestTotal = qty;

        bestDay =
          cellDateToIso_(
            row[0],
            ss.getSpreadsheetTimeZone()
          );
      }
    });
  }

  const species =
    new Set();

  const catchSheet =
    ss.getSheetByName(
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
        Number(row[0]) ===
          Number(year) &&
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
 * Единое оформление старых и новых строк.
 */
function applyUnifiedFormatting_(ss) {
  getYearSheets_(ss).forEach(function(sheet) {
    formatYearSheet_(
      sheet
    );
  });

  const catchSheet =
    ss.getSheetByName(
      "Улов"
    );

  if (catchSheet) {
    formatCatchSheet_(
      catchSheet
    );
  }

  const summarySheet =
    ss.getSheetByName(
      "Сводка"
    );

  if (summarySheet) {
    formatSummarySheet_(
      summarySheet
    );
  }
}


function formatYearSheet_(sheet) {
  const lastRow = Math.max(
    sheet.getLastRow(),
    1
  );

  ensureSheetCapacity_(
    sheet,
    lastRow,
    YEAR_HEADERS.length
  );

  const range = sheet
    .getRange(
      1,
      1,
      lastRow,
      YEAR_HEADERS.length
    );

  range
    .setFontFamily("Arial")
    .setFontColor("#000000")
    .setVerticalAlignment("middle");

  if (lastRow > 1) {
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        YEAR_HEADERS.length
      )
      .setBackground("#ffffff")
      .setFontWeight("normal");

    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .setNumberFormat(
        "dd.mm.yyyy"
      );

    sheet
      .getRange(
        2,
        5,
        lastRow - 1,
        YEAR_HEADERS.length - 4
      )
      .setWrap(true);
  }

  formatHeader_(
    sheet,
    YEAR_HEADERS.length
  );

  setYearColumnWidths_(
    sheet
  );

  ensureFilter_(
    sheet,
    YEAR_HEADERS.length
  );
}


function formatCatchSheet_(sheet) {
  const lastRow = Math.max(
    sheet.getLastRow(),
    1
  );

  ensureSheetCapacity_(
    sheet,
    lastRow,
    CATCH_HEADERS.length
  );

  sheet
    .getRange(
      1,
      1,
      lastRow,
      CATCH_HEADERS.length
    )
    .setFontFamily("Arial")
    .setFontColor("#000000")
    .setVerticalAlignment("middle");

  if (lastRow > 1) {
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        CATCH_HEADERS.length
      )
      .setBackground("#ffffff")
      .setFontWeight("normal")
      .setWrap(true);

    sheet
      .getRange(
        2,
        2,
        lastRow - 1,
        1
      )
      .setNumberFormat(
        "dd.mm.yyyy"
      );
  }

  formatHeader_(
    sheet,
    CATCH_HEADERS.length
  );

  const widths = [
    70, 95, 75, 120, 90,
    150, 130, 160, 180, 150
  ];

  widths.forEach(function(width, index) {
    sheet.setColumnWidth(
      index + 1,
      width
    );
  });

  ensureFilter_(
    sheet,
    CATCH_HEADERS.length
  );
}


function formatSummarySheet_(sheet) {
  const lastRow = Math.max(
    sheet.getLastRow(),
    1
  );

  ensureSheetCapacity_(
    sheet,
    lastRow,
    SUMMARY_HEADERS.length
  );

  sheet
    .getRange(
      1,
      1,
      lastRow,
      SUMMARY_HEADERS.length
    )
    .setFontFamily("Arial")
    .setFontColor("#000000")
    .setVerticalAlignment("middle");

  if (lastRow > 1) {
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        SUMMARY_HEADERS.length
      )
      .setBackground("#ffffff")
      .setFontWeight("normal");

    sheet
      .getRange(
        2,
        5,
        lastRow - 1,
        1
      )
      .setNumberFormat(
        "dd.mm.yyyy"
      );

    sheet
      .getRange(
        2,
        7,
        lastRow - 1,
        1
      )
      .setNumberFormat(
        "dd.mm.yyyy hh:mm"
      );
  }

  formatHeader_(
    sheet,
    SUMMARY_HEADERS.length
  );

  const widths = [
    75, 130, 120, 110,
    110, 140, 170
  ];

  widths.forEach(function(width, index) {
    sheet.setColumnWidth(
      index + 1,
      width
    );
  });
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
    .setVerticalAlignment("middle")
    .setWrap(true);

  sheet.setFrozenRows(1);
}


function setYearColumnWidths_(sheet) {
  const widths = [
    95, 105, 80, 90, 220, 240,
    105, 105, 140, 130, 80, 120, 120,
    95, 105, 125, 90, 95, 190,
    150, 140, 145, 120, 150, 140, 170, 190
  ];

  widths.forEach(function(width, index) {
    sheet.setColumnWidth(
      index + 1,
      width
    );
  });
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

  if (sheet.getLastRow() >= 2) {
    sheet
      .getRange(
        1,
        1,
        sheet.getLastRow(),
        columnCount
      )
      .createFilter();
  }
}


function ensureSheetCapacity_(
  sheet,
  requiredRows,
  requiredColumns
) {
  const maxColumns =
    sheet.getMaxColumns();

  if (
    maxColumns <
    requiredColumns
  ) {
    sheet.insertColumnsAfter(
      maxColumns,
      requiredColumns -
        maxColumns
    );
  }

  const maxRows =
    sheet.getMaxRows();

  if (
    maxRows <
    requiredRows
  ) {
    sheet.insertRowsAfter(
      maxRows,
      requiredRows -
        maxRows
    );
  }
}


function getYearSheets_(ss) {
  return ss
    .getSheets()
    .filter(function(sheet) {
      return /^\d{4}$/.test(
        sheet.getName()
      );
    });
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


function buildDateRowMap_(
  sheet,
  timeZone
) {
  const map = {};

  if (
    sheet.getLastRow() < 2
  ) {
    return map;
  }

  const values = sheet
    .getRange(
      2,
      1,
      sheet.getLastRow() - 1,
      1
    )
    .getValues();

  values.forEach(function(row, index) {
    const iso =
      cellDateToIso_(
        row[0],
        timeZone
      );

    if (iso) {
      map[iso] =
        index + 2;
    }
  });

  return map;
}


function getLastDateFromMap_(
  map
) {
  const dates =
    Object.keys(map)
      .sort();

  return dates.length
    ? dates[
        dates.length - 1
      ]
    : "";
}


function removeFutureRows_(
  sheet,
  todayIso,
  timeZone
) {
  if (
    sheet.getLastRow() < 2
  ) {
    return;
  }

  const dates = sheet
    .getRange(
      2,
      1,
      sheet.getLastRow() - 1,
      1
    )
    .getValues();

  for (
    let index =
      dates.length - 1;
    index >= 0;
    index--
  ) {
    const iso =
      cellDateToIso_(
        dates[index][0],
        timeZone
      );

    if (
      iso &&
      iso > todayIso
    ) {
      sheet.deleteRow(
        index + 2
      );
    }
  }
}


function deleteCatchRowsForRange_(
  sheet,
  fromIso,
  toIso,
  timeZone
) {
  if (
    sheet.getLastRow() < 2
  ) {
    return;
  }

  const dates = sheet
    .getRange(
      2,
      2,
      sheet.getLastRow() - 1,
      1
    )
    .getValues();

  for (
    let index =
      dates.length - 1;
    index >= 0;
    index--
  ) {
    const iso =
      cellDateToIso_(
        dates[index][0],
        timeZone
      );

    if (
      iso &&
      iso >= fromIso &&
      iso <= toIso
    ) {
      sheet.deleteRow(
        index + 2
      );
    }
  }
}


function findSummaryYearRow_(
  sheet,
  year
) {
  if (
    sheet.getLastRow() < 2
  ) {
    return null;
  }

  const values = sheet
    .getRange(
      2,
      1,
      sheet.getLastRow() - 1,
      1
    )
    .getValues();

  for (
    let index = 0;
    index < values.length;
    index++
  ) {
    if (
      Number(values[index][0]) ===
      Number(year)
    ) {
      return index + 2;
    }
  }

  return null;
}


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
    result.push(
      current
    );

    current =
      shiftIsoDate_(
        current,
        1,
        timeZone
      );
  }

  return result;
}


function shiftIsoDate_(
  isoDate,
  days,
  timeZone
) {
  const parts = String(
    isoDate
  )
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
    date.getDate() +
      days
  );

  return dateToIso_(
    date,
    timeZone
  );
}


function parseDate_(
  isoDate
) {
  if (!isoDate) {
    return "";
  }

  const parts = String(
    isoDate
  )
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
    !isNaN(
      value.getTime()
    )
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
    /^\d{4}-\d{2}-\d{2}$/.test(
      text
    )
  ) {
    return text;
  }

  const match =
    text.match(
      /^(\d{1,2})[./](\d{1,2})[./](\d{4})$/
    );

  if (match) {
    return [
      match[3],
      String(match[2])
        .padStart(2, "0"),
      String(match[1])
        .padStart(2, "0")
    ].join("-");
  }

  return "";
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


function isIsoDate_(
  value
) {
  return /^\d{4}-\d{2}-\d{2}$/
    .test(
      String(value || "")
    );
}


function isEmptyValue_(
  value
) {
  return (
    value === "" ||
    value === null ||
    typeof value ===
      "undefined"
  );
}


function isNumberLike_(
  value
) {
  if (
    value === "" ||
    value === null ||
    typeof value ===
      "undefined"
  ) {
    return false;
  }

  return !isNaN(
    Number(value)
  );
}


function value_(
  value
) {
  return isEmptyValue_(
    value
  )
    ? ""
    : value;
}


function number_(
  value
) {
  const number =
    Number(value);

  return isNaN(number)
    ? 0
    : number;
}


function unique_(
  values
) {
  return Array.from(
    new Set(values)
  );
}


function jsonResponse_(
  data
) {
  return ContentService
    .createTextOutput(
      JSON.stringify(data)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}


function testConnection() {
  const ss = SpreadsheetApp.openById(
    SPREADSHEET_ID
  );

  console.log(
    "Подключение успешно."
  );

  console.log(
    "Таблица: " +
      ss.getName()
  );

  console.log(
    "Версия схемы: unified-v19"
  );
}
