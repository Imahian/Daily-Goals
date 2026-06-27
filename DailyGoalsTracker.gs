// ============================================================
//  DAILY GOALS TRACKER - Google Apps Script  v4
//  All comments and messages in English
// ============================================================

// ------------------------------------------------------------
// 0. GLOBAL CONFIG
// ------------------------------------------------------------
const CFG = {
  SHEET_SCHEDULE : "Schedule",
  SHEET_DAILY    : "Daily",
  SHEET_METRICS  : "Metrics",

  RANGE_ROUTINE   : "Routine_enc",
  RANGE_DAILY     : "Daily_enc",
  RANGE_REFERENCE : "Reference_enc",
  RANGE_AVERAGE   : "Average_enc",

  COL_ACTIVITY   : "Activity",
  COL_RANGE      : "Range",
  COL_CHECK      : "Check",
  COL_DATE       : "Date",
  COL_EFFICIENCY : "Efficiency",
  COL_REF_ICON   : "Icon",
  COL_REF_CAT    : "Category",

  CALENDAR_ID : "primary",
  EVENT_COLOR : CalendarApp.EventColor.CYAN,
  TOAST_SEC   : 4,

  // Row in the sheet where the progress bar lives (and any formula below Daily)
  // The script NEVER touches this row or anything below it.
  // Update this number if you move the progress bar.
  DAILY_PROTECTED_FROM_ROW : 17,
};

// ------------------------------------------------------------
// TOAST HELPERS
// ------------------------------------------------------------
function toast(msg, title, sec) {
  SpreadsheetApp.getActiveSpreadsheet()
    .toast(msg, title || "Daily Tracker", sec || CFG.TOAST_SEC);
}
function toastOk(msg)   { toast("OK: "    + msg, "OK");       }
function toastInfo(msg) { toast("INFO: "  + msg, "Info");     }
function toastWarn(msg) { toast("AVISO: " + msg, "Aviso", 5); }
function toastErr(msg)  { toast("ERROR: " + msg, "Error", 8); }

// ------------------------------------------------------------
// 1. TABLE UTILITIES
// ------------------------------------------------------------
function getTableMeta(ss, namedRange) {
  var nr = ss.getRangeByName(namedRange);
  if (!nr) throw new Error('Rango con nombre "' + namedRange + '" no encontrado.');
  var sheet     = nr.getSheet();
  var headerRow = nr.getRow();
  var startCol  = nr.getColumn();
  var numCols   = nr.getNumColumns();
  var headers   = nr.getValues()[0];
  var headerMap = {};
  headers.forEach(function(h, i) {
    if (String(h).trim() !== "") headerMap[String(h).trim()] = i;
  });
  return { sheet: sheet, headerRow: headerRow, dataStartRow: headerRow + 1,
           startCol: startCol, numCols: numCols, headers: headers, headerMap: headerMap };
}

function getTableData(ss, namedRange) {
  var meta    = getTableMeta(ss, namedRange);
  var lastRow = meta.sheet.getLastRow();
  if (lastRow < meta.dataStartRow) return { meta: meta, data: [] };
  var numRows = lastRow - meta.dataStartRow + 1;
  var data    = meta.sheet
    .getRange(meta.dataStartRow, meta.startCol, numRows, meta.numCols)
    .getValues();
  return { meta: meta, data: data };
}

function colIdx(headerMap, name) {
  var key = Object.keys(headerMap).find(function(k) {
    return k.toLowerCase() === name.toLowerCase();
  });
  if (key === undefined) throw new Error('Column "' + name + '" not found.');
  return headerMap[key];
}

// ------------------------------------------------------------
// 2. READ REFERENCE TABLE
// ------------------------------------------------------------
function getReference(ss) {
  var res  = getTableData(ss, CFG.RANGE_REFERENCE);
  var meta = res.meta;
  var data = res.data;
  var iIcon = colIdx(meta.headerMap, CFG.COL_REF_ICON);
  var iCat  = colIdx(meta.headerMap, CFG.COL_REF_CAT);
  return data
    .filter(function(row) { return String(row[iIcon]).trim() !== ""; })
    .map(function(row) {
      return { icon: String(row[iIcon]).trim(), category: String(row[iCat]).trim() };
    });
}

// ------------------------------------------------------------
// 3. PARSE ROUTINE -> BLOCKS
// ------------------------------------------------------------
function todayDayName() {
  return ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"][new Date().getDay()];
}

function addMinutes(timeStr, mins) {
  var parts = timeStr.split(":");
  var h = parseInt(parts[0]);
  var m = parseInt(parts[1]);
  var t = h * 60 + m + mins;
  return pad2(Math.floor(t / 60) % 24) + ":" + pad2(t % 60);
}

function pad2(n) { return String(n).padStart(2, "0"); }

function isClockEmoji(val) {
  var s  = String(val).trim();
  if (!s) return false;
  var cp = s.codePointAt(0);
  return cp >= 0x1F550 && cp <= 0x1F567;
}

/**
 * Builds a time map for every row in the Routine table.
 * Counts only valid clock emojis (ignores decorators like the sun emoji).
 * Each valid clock emoji = one 30-min slot starting at 00:00.
 * Returns an array parallel to data with "HH:MM" or null.
 */
function buildTimeMap(data, iTime) {
  var slotIndex = 0;
  return data.map(function(row) {
    var val = row[iTime];
    if (isClockEmoji(val)) {
      var totalMin = slotIndex * 30;
      slotIndex++;
      return pad2(Math.floor(totalMin / 60) % 24) + ":" + pad2(totalMin % 60);
    }
    if (val instanceof Date) {
      return pad2(val.getHours()) + ":" + pad2(val.getMinutes());
    }
    var s = String(val).trim();
    if (/^\d{1,2}:\d{2}/.test(s)) {
      var hp = s.split(":");
      return pad2(parseInt(hp[0])) + ":" + pad2(parseInt(hp[1]));
    }
    var num = parseFloat(s);
    if (!isNaN(num) && num >= 0 && num < 1) {
      var t = Math.round(num * 24 * 60);
      return pad2(Math.floor(t / 60)) + ":" + pad2(t % 60);
    }
    return null;
  });
}

/**
 * Lee Routine y agrupa slots consecutivos del mismo emoji en bloques.
 * End = inicio del siguiente slot de reloj diferente (hora exacta).
 * Emojis no en Reference (sol, luna, etc.) se ignoran.
 */
function getDayBlocks(ss, dayName) {
  var res  = getTableData(ss, CFG.RANGE_ROUTINE);
  var meta = res.meta;
  var data = res.data;

  var refs   = getReference(ss);
  var refMap = {};
  refs.forEach(function(r) { refMap[r.icon] = r.category; });

  var iTime = colIdx(meta.headerMap, "Time");
  var iDay  = colIdx(meta.headerMap, dayName);

  var timeMap = buildTimeMap(data, iTime);

  var slots = data
    .map(function(row, i) {
      return { time: timeMap[i], icon: String(row[iDay]).trim() };
    })
    .filter(function(s) { return s.time !== null; });

  var blocks = [];
  var i = 0;
  while (i < slots.length) {
    var icon = slots[i].icon;
    if (!icon || !refMap[icon]) { i++; continue; }

    var start = slots[i].time;
    var j = i + 1;
    while (j < slots.length && slots[j].icon === icon) j++;

    var end = (j < slots.length) ? slots[j].time : addMinutes(slots[j-1].time, 30);

    blocks.push({
      icon     : icon,
      category : refMap[icon],
      start    : start,
      end      : end,
      rangeStr : start + " - " + end
    });
    i = j;
  }
  return blocks;
}

// ------------------------------------------------------------
// 4. POPULATE DAILY TABLE
// ------------------------------------------------------------

/**
 * Returns how many data rows the Daily table currently has.
 * NEVER counts past CFG.DAILY_PROTECTED_FROM_ROW - 1
 * to guarantee external formulas are never detected or touched.
 */
function getDailyRowCount(sheet, meta) {
  var maxRow    = CFG.DAILY_PROTECTED_FROM_ROW - 1; // row 16 if progress bar is on row 17
  var maxRows   = maxRow - meta.dataStartRow + 1;   // = 11 with header on row 5
  var actAbsCol = meta.startCol + colIdx(meta.headerMap, CFG.COL_ACTIVITY);
  var count = 0;
  for (var r = 0; r < maxRows; r++) {
    var val = sheet.getRange(meta.dataStartRow + r, actAbsCol).getValue();
    if (val === "" || val === null) break;
    count++;
  }
  return count;
}

function populateDailyTable(ss) {
  toastInfo("Reading Routine table...");

  var day    = todayDayName();
  var blocks = getDayBlocks(ss, day);
  toastInfo("Day: " + day + " - " + blocks.length + " block(s)");

  var meta        = getTableMeta(ss, CFG.RANGE_DAILY);
  var sheet       = meta.sheet;
  var iAct        = colIdx(meta.headerMap, CFG.COL_ACTIVITY);
  var iRange      = colIdx(meta.headerMap, CFG.COL_RANGE);
  var iCheck      = colIdx(meta.headerMap, CFG.COL_CHECK);
  var actAbsCol   = meta.startCol + iAct;
  var rangeAbsCol = meta.startCol + iRange;
  var checkAbsCol = meta.startCol + iCheck;

  // Max rows available before the progress bar
  var maxAllowed = CFG.DAILY_PROTECTED_FROM_ROW - meta.dataStartRow; // = 11

  // Count current rows (never past the protected zone)
  var currentRows = getDailyRowCount(sheet, meta);
  toastInfo("Current rows: " + currentRows + " | New: " + blocks.length + " | Max: " + maxAllowed);

  // Warn if blocks exceed available space
  if (blocks.length > maxAllowed) {
    toastWarn("There are " + blocks.length + " blocks but only room for " + maxAllowed +
      ". Move the progress bar (B" + CFG.DAILY_PROTECTED_FROM_ROW + ") further down and update CFG.DAILY_PROTECTED_FROM_ROW.");
    blocks = blocks.slice(0, maxAllowed); // truncate to avoid overwriting
  }

  // ── Limpiar exactamente las filas anteriores de la tabla ──────
  // Only clears table columns, only up to currentRows.
  // NEVER touches CFG.DAILY_PROTECTED_FROM_ROW or beyond.
  if (currentRows > 0) {
    var clearArea = sheet.getRange(
      meta.dataStartRow, meta.startCol, currentRows, meta.numCols
    );
    clearArea.clearDataValidations();
    clearArea.clearContent();
    toastInfo("Cleared " + currentRows + " previous row(s)");
  }

  if (blocks.length === 0) {
    toastWarn("No blocks for today. Check Routine and Reference.");
    return;
  }

  // ── Escribir exactamente blocks.length filas ──────────────────
  var actValues   = blocks.map(function(b) { return [b.icon + " " + b.category]; });
  var rangeValues = blocks.map(function(b) { return [b.rangeStr]; });

  sheet.getRange(meta.dataStartRow, actAbsCol,   blocks.length, 1).setValues(actValues);
  sheet.getRange(meta.dataStartRow, rangeAbsCol, blocks.length, 1).setValues(rangeValues);

  // ── Resetear checkboxes a false ───────────────────────────────
  for (var r = 0; r < blocks.length; r++) {
    sheet.getRange(meta.dataStartRow + r, checkAbsCol).setValue(false);
  }

  SpreadsheetApp.flush();
  toastOk("Daily ready - " + blocks.length + " activities for " + day);
}

// ------------------------------------------------------------
// 5. GOOGLE CALENDAR / TASKS - CREATE TODAY'S EVENTS
// ------------------------------------------------------------
function syncCalendarEvents(ss) {
  toastInfo("Creating Tasks in Google Calendar...");

  var blocks = getDayBlocks(ss, todayDayName());
  if (blocks.length === 0) {
    toastWarn("No blocks - no Tasks created");
    return;
  }

  var today = new Date(); today.setHours(0, 0, 0, 0);

  // Try Tasks API first
  try {
    var taskLists = Tasks.Tasklists.list({ maxResults: 1 });
    var listId    = taskLists.items[0].id;

    // Delete previous tracker tasks
    var existing = Tasks.Tasks.list(listId, { showCompleted: true, showHidden: true });
    if (existing.items) {
      existing.items
        .filter(function(t) { return (t.notes || "").indexOf("DailyTracker") === 0; })
        .forEach(function(t) { try { Tasks.Tasks.remove(listId, t.id); } catch(e) {} });
    }
    toastInfo("Previous tasks removed. Creating new ones...");

    var created = 0;
    blocks.forEach(function(b) {
      try {
        var parts = b.start.split(":");
        var due = new Date(today);
        due.setHours(parseInt(parts[0]), parseInt(parts[1]), 0, 0);
        Tasks.Tasks.insert({
          title : b.icon + " " + b.category + "  " + b.rangeStr,
          notes : "DailyTracker - " + b.rangeStr,
          due   : due.toISOString()
        }, listId);
        created++;
      } catch(e) {
        Logger.log("Task insert error: " + e.message);
      }
    });
    toastOk(created + " Task(s) created in Google Tasks");

  } catch (tasksErr) {
    // Fallback: Calendar events
    toastWarn("Tasks API unavailable. Falling back to Calendar events...");
    Logger.log("Tasks API error: " + tasksErr.message);

    var cal;
    try {
      cal = CalendarApp.getCalendarById(CFG.CALENDAR_ID);
      if (!cal) throw new Error("null");
    } catch(e) { cal = CalendarApp.getDefaultCalendar(); }

    var tomorrow = new Date(today); tomorrow.setDate(today.getDate() + 1);
    cal.getEvents(today, tomorrow)
      .filter(function(ev) { return ev.getTag("dailyTracker") === "true"; })
      .forEach(function(ev) { try { ev.deleteEvent(); } catch(e) {} });

    var created2 = 0;
    blocks.forEach(function(b) {
      try {
        var sp = b.start.split(":"); var ep = b.end.split(":");
        var startDt = new Date(today); startDt.setHours(parseInt(sp[0]), parseInt(sp[1]), 0, 0);
        var endDt   = new Date(today); endDt.setHours(parseInt(ep[0]), parseInt(ep[1]), 0, 0);
        if (endDt <= startDt) endDt.setDate(endDt.getDate() + 1);
        var ev = cal.createEvent(b.icon + " " + b.category, startDt, endDt,
          { description: "DailyTracker - " + b.rangeStr });
        ev.setTag("dailyTracker", "true");
        ev.setColor(CFG.EVENT_COLOR);
        created2++;
      } catch(e) { toastWarn("Error creando evento: " + e.message); }
    });
    toastOk(created2 + " Calendar event(s) created");
  }
}

function clearPreviousTasks() {
  try {
    var taskLists = Tasks.Tasklists.list({ maxResults: 1 });
    var listId    = taskLists.items[0].id;
    var existing  = Tasks.Tasks.list(listId, { showCompleted: true, showHidden: true });
    if (!existing.items) return;
    var removed = 0;
    existing.items
      .filter(function(t) { return (t.notes || "").indexOf("DailyTracker") === 0; })
      .forEach(function(t) { try { Tasks.Tasks.remove(listId, t.id); removed++; } catch(e) {} });
    toastInfo(removed + " task(s) from previous day removed");
  } catch(e) {
    try {
      var cal = CalendarApp.getCalendarById(CFG.CALENDAR_ID) || CalendarApp.getDefaultCalendar();
      var yesterday = new Date(); yesterday.setDate(yesterday.getDate()-1); yesterday.setHours(0,0,0,0);
      var today2    = new Date(); today2.setHours(0,0,0,0);
      cal.getEvents(yesterday, today2)
        .filter(function(ev) { return ev.getTag("dailyTracker") === "true"; })
        .forEach(function(ev) { try { ev.deleteEvent(); } catch(e2) {} });
    } catch(e2) {}
  }
}

// ------------------------------------------------------------
// 6. SYNC DAILY CHECKBOX -> CALENDAR
// ------------------------------------------------------------
function syncCheckboxToCalendar(ss) {
  toastInfo("Syncing checkboxes with Calendar...");

  var res  = getTableData(ss, CFG.RANGE_DAILY);
  var meta = res.meta;
  var data = res.data;

  var iAct   = colIdx(meta.headerMap, CFG.COL_ACTIVITY);
  var iRange = colIdx(meta.headerMap, CFG.COL_RANGE);
  var iCheck = colIdx(meta.headerMap, CFG.COL_CHECK);

  if (data.length === 0) { toastWarn("Daily is empty - nothing to sync"); return; }

  var updated = 0;

  // Try Tasks API first
  var usedTasks = false;
  try {
    var taskLists = Tasks.Tasklists.list({ maxResults: 1 });
    var listId    = taskLists.items[0].id;
    var taskData  = Tasks.Tasks.list(listId, { showCompleted: true, showHidden: true });
    var tasks     = (taskData.items || []).filter(function(t) {
      return (t.notes || "").indexOf("DailyTracker") === 0;
    });

    if (tasks.length > 0) {
      usedTasks = true;
      data.forEach(function(row) {
        var rangeStr = String(row[iRange]).trim();
        var checked  = row[iCheck] === true;
        var task = tasks.find(function(t) { return (t.notes || "").indexOf(rangeStr) >= 0; });
        if (!task) return;
        var taskDone = task.status === "completed";
        if (checked && !taskDone) {
          Tasks.Tasks.patch({ status: "completed", completed: new Date().toISOString() }, listId, task.id);
          updated++;
        } else if (!checked && taskDone) {
          Tasks.Tasks.patch({ status: "needsAction", completed: null }, listId, task.id);
          updated++;
        }
      });
      Logger.log("syncCheckboxToCalendar Tasks API: " + updated + " updated");
    }
  } catch(tasksErr) {
    Logger.log("Tasks API error en syncCheckbox: " + tasksErr.message);
  }

  // Fallback: Calendar events
  if (!usedTasks) {
    try {
      var cal;
      try {
        cal = CalendarApp.getCalendarById(CFG.CALENDAR_ID);
        if (!cal) throw new Error();
      } catch(e) { cal = CalendarApp.getDefaultCalendar(); }

      var today    = new Date(); today.setHours(0, 0, 0, 0);
      var tomorrow = new Date(today); tomorrow.setDate(today.getDate() + 1);
      var events   = cal.getEvents(today, tomorrow)
        .filter(function(ev) { return ev.getTag("dailyTracker") === "true"; });

      Logger.log("syncCheckboxToCalendar: " + events.length + " eventos encontrados");

      data.forEach(function(row) {
        var activity = String(row[iAct]).trim();
        var checked  = row[iCheck] === true;
        var rangeStr = String(row[iRange]).trim();
        var match = events.find(function(ev) {
          return (ev.getDescription() || "").indexOf(rangeStr) >= 0;
        });
        if (!match) { Logger.log("  not found: " + rangeStr); return; }
        var done = match.getTitle().indexOf("OK ") === 0;
        if (checked && !done)  { match.setTitle("OK " + activity); updated++; }
        if (!checked && done)  { match.setTitle(activity);          updated++; }
      });
      Logger.log("syncCheckboxToCalendar Calendar: " + updated + " updated");
    } catch(calErr) {
      toastErr("Calendar error: " + calErr.message);
      Logger.log("syncCheckboxToCalendar Calendar error: " + calErr.stack);
    }
  }

  toastOk(updated + " updated in Calendar");
}

// ------------------------------------------------------------
// 6b. SYNC CALENDAR -> DAILY (polling every 5 min)
// ------------------------------------------------------------
function syncCalendarToDaily(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();

  var res  = getTableData(ss, CFG.RANGE_DAILY);
  var meta = res.meta;
  var data = res.data;
  if (data.length === 0) {
    Logger.log("syncCalendarToDaily: Daily is empty, nothing to sync");
    return;
  }

  var iRange      = colIdx(meta.headerMap, CFG.COL_RANGE);
  var iCheck      = colIdx(meta.headerMap, CFG.COL_CHECK);
  var checkAbsCol = meta.startCol + iCheck;
  var updated     = 0;
  var usedTasks   = false;

  // --- Tasks API: search across ALL task lists ---
  try {
    var allTasks  = [];
    var taskLists = Tasks.Tasklists.list({ maxResults: 20 });

    (taskLists.items || []).forEach(function(list) {
      try {
        var taskData = Tasks.Tasks.list(list.id, {
          showCompleted : true,
          showHidden    : true,
          maxResults    : 100
        });
        (taskData.items || []).forEach(function(t) {
          if ((t.notes || "").indexOf("DailyTracker") === 0) {
            t._listId = list.id;  // store listId for patching
            allTasks.push(t);
          }
        });
      } catch(e) { Logger.log("Error leyendo lista " + list.id + ": " + e.message); }
    });

    Logger.log("syncCalendarToDaily: Tasks API found " + allTasks.length + " tracker task(s)");
    allTasks.forEach(function(t) {
      Logger.log("  task: status=" + t.status + " notes=" + t.notes);
    });

    if (allTasks.length > 0) {
      usedTasks = true;
      data.forEach(function(row, idx) {
        var rangeStr   = String(row[iRange]).trim();
        var curChecked = row[iCheck] === true;
        var task = allTasks.find(function(t) {
          return (t.notes || "").indexOf(rangeStr) >= 0;
        });
        if (!task) {
          Logger.log("  no match for rangeStr: " + rangeStr);
          return;
        }
        var taskDone = task.status === "completed";
        Logger.log("  rangeStr:" + rangeStr + " taskDone:" + taskDone + " curChecked:" + curChecked);
        if (taskDone !== curChecked) {
          meta.sheet.getRange(meta.dataStartRow + idx, checkAbsCol).setValue(taskDone);
          updated++;
        }
      });
    }

    if (updated > 0) {
      SpreadsheetApp.flush();
      toastInfo("Calendar->Daily: " + updated + " checkbox(es) updated");
    }
    Logger.log("syncCalendarToDaily Tasks API: " + updated + " updated");

  } catch(tasksErr) {
    Logger.log("syncCalendarToDaily Tasks API error: " + tasksErr.message);
  }

  // --- Fallback: Calendar events ---
  if (!usedTasks) {
    try {
      var cal;
      try {
        cal = CalendarApp.getCalendarById(CFG.CALENDAR_ID);
        if (!cal) throw new Error("null");
      } catch(e) { cal = CalendarApp.getDefaultCalendar(); }

      var today    = new Date(); today.setHours(0, 0, 0, 0);
      var tomorrow = new Date(today); tomorrow.setDate(today.getDate() + 1);
      var events   = cal.getEvents(today, tomorrow)
        .filter(function(ev) { return ev.getTag("dailyTracker") === "true"; });

      Logger.log("syncCalendarToDaily fallback: " + events.length + " events found");

      data.forEach(function(row, idx) {
        var rangeStr   = String(row[iRange]).trim();
        var curChecked = row[iCheck] === true;
        var ev = events.find(function(e) { return (e.getDescription() || "").indexOf(rangeStr) >= 0; });
        if (!ev) return;
        var evDone = ev.getTitle().indexOf("OK ") === 0;
        if (evDone !== curChecked) {
          meta.sheet.getRange(meta.dataStartRow + idx, checkAbsCol).setValue(evDone);
          updated++;
        }
      });

      if (updated > 0) {
        toastInfo("Calendar->Daily: " + updated + " checkbox(es) synced");
      }
    } catch(e2) {
      Logger.log("syncCalendarToDaily fallback error: " + e2.message);
    }
  }
}

// ------------------------------------------------------------
// 7. AVERAGE COLUMNS - AUTO-SYNC WITH REFERENCE
// ------------------------------------------------------------
function ensureAverageColumns(ss) {
  toastInfo("Checking Average columns...");

  var refs    = getReference(ss);
  var avgMeta = getTableMeta(ss, CFG.RANGE_AVERAGE);
  var sheet   = avgMeta.sheet;
  var hRow    = avgMeta.headerRow;
  var hStart  = avgMeta.startCol;

  // Read current headers
  var currentCols    = Math.max(sheet.getLastColumn() - hStart + 1, 2);
  var currentHeaders = sheet.getRange(hRow, hStart, 1, currentCols).getValues()[0]
    .map(function(v) { return String(v).trim(); });

  // Find absolute column of Efficiency (last real column of the table)
  var iEff      = currentHeaders.findIndex(function(h) { return h.toLowerCase() === CFG.COL_EFFICIENCY.toLowerCase(); });
  var effAbsCol = hStart + (iEff >= 0 ? iEff : currentHeaders.length - 1); // absolute Efficiency colmna absoluta en la hoja

  // Icons already in the header (between Date and Efficiency)
  var iDate         = currentHeaders.findIndex(function(h) { return h.toLowerCase() === CFG.COL_DATE.toLowerCase(); });
  var iconStartIdx  = iDate >= 0 ? iDate + 1 : 1;
  var iconEndIdx    = iEff  >= 0 ? iEff      : currentHeaders.length - 1;
  var existingIcons = currentHeaders.slice(iconStartIdx, iconEndIdx);

  // Detect new icons not yet present (in Reference order)
  var refIcons = refs.map(function(r) { return r.icon; });
  var newIcons = refIcons.filter(function(ic) { return existingIcons.indexOf(ic) === -1; });

  if (newIcons.length === 0) {
    toastOk("Average up to date - no new columns");
    return;
  }

  // For each new icon: insert a physical column BEFORE Efficiency
  // and write the header. Done left-to-right; effAbsCol tracks the
  // shifting column index of Efficiency after each insert.
  // Since we always insert before Efficiency, the final order matches
  // Reference order (new icons appended after existing ones).
  newIcons.forEach(function(ic) {
    // effAbsCol se actualiza con cada insercion porque insertamos antes de ella
    sheet.insertColumnBefore(effAbsCol);
    // El nuevo icono va en la columna recien insertada (effAbsCol, que se desplazo +1)
    sheet.getRange(hRow, effAbsCol).setValue(ic);
    // effAbsCol increments because Efficiency shifted one column to the right
    // after each insert:
    effAbsCol = effAbsCol + 1;
  });

  SpreadsheetApp.flush();
  toastOk(newIcons.length + " new column(s) inserted in Average: " + newIcons.join(" "));
}

// ------------------------------------------------------------
// 8. NIGHTLY SNAPSHOT -> AVERAGE
// ------------------------------------------------------------
function saveDailySnapshot(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();

  toastInfo("Starting snapshot...");
  ensureAverageColumns(ss);

  var refs     = getReference(ss);
  var refIcons = refs.map(function(r) { return r.icon; });

  var res   = getTableData(ss, CFG.RANGE_DAILY);
  var dMeta = res.meta;
  var dData = res.data;

  var iAct   = colIdx(dMeta.headerMap, CFG.COL_ACTIVITY);
  var iCheck = colIdx(dMeta.headerMap, CFG.COL_CHECK);

  toastInfo("Counting " + dData.length + " activitie(s) in Daily...");

  var count = {};
  refIcons.forEach(function(ic) { count[ic] = { done: 0, total: 0 }; });

  dData.forEach(function(row) {
    var actStr  = String(row[iAct]).trim();
    var checked = row[iCheck] === true;
    var icon    = refIcons.find(function(ic) { return actStr.indexOf(ic) === 0; }) || null;
    if (icon) { count[icon].total++; if (checked) count[icon].done++; }
  });

  var totalTasks = 0, doneTasks = 0;
  refIcons.forEach(function(ic) { totalTasks += count[ic].total; doneTasks += count[ic].done; });
  var effPct = totalTasks > 0 ? Math.round((doneTasks / totalTasks) * 100) : 0;
  toastInfo("Efficiency: " + effPct + "% (" + doneTasks + "/" + totalTasks + ")");

  var avgMeta  = getTableMeta(ss, CFG.RANGE_AVERAGE);
  var avgSheet = avgMeta.sheet;
  var insertAt = avgMeta.dataStartRow;

  // Insert new row just below header (most recent date on top)
  avgSheet.insertRowBefore(insertAt);

  var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  var rowValues = [dateStr].concat(
    refIcons.map(function(ic) {
      return count[ic].total === 0 ? false : (count[ic].done === count[ic].total);
    })
  ).concat([effPct + "%"]);

  avgSheet.getRange(insertAt, avgMeta.startCol, 1, rowValues.length).setValues([rowValues]);
  SpreadsheetApp.flush();
  toastOk("Snapshot saved -> row " + insertAt + " | " + dateStr + " | " + effPct + "%");

  // Clear Daily table after saving (ready for the new day)
  // Only clears content — does NOT delete physical rows to preserve B17
  toastInfo("Clearing Daily table...");
  var dailyMeta   = getTableMeta(ss, CFG.RANGE_DAILY);
  var dSheet      = dailyMeta.sheet;
  var currentRows = getDailyRowCount(dSheet, dailyMeta);

  if (currentRows > 0) {
    dSheet.getRange(dailyMeta.dataStartRow, dailyMeta.startCol,
      currentRows, dailyMeta.numCols)
      .clearDataValidations()
      .clearContent();
  }

  SpreadsheetApp.flush();
  toastOk("Daily cleared and ready for the new day");
}

// ------------------------------------------------------------
// 9. MAIN FUNCTION - iniciarSistema()
// ------------------------------------------------------------
function iniciarSistema() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var h  = new Date().getHours();
  var m  = new Date().getMinutes();
  toast("System starting...", "Daily Tracker", 3);

  try {
    if (h === 0 && m <= 5) {
      toastInfo("Mode: Morning setup");
      clearPreviousTasks();
      Utilities.sleep(500);
      populateDailyTable(ss);
      Utilities.sleep(800);
      syncCalendarEvents(ss);
      toastOk("Morning setup complete - good day!");

    } else if (h === 23 && m >= 50) {
      toastInfo("Mode: Nightly snapshot");
      saveDailySnapshot(ss);
      toastOk("Snapshot saved - good night!");

    } else {
      toastInfo("Mode: Full flow");
      clearPreviousTasks();
      Utilities.sleep(500);
      populateDailyTable(ss);
      Utilities.sleep(800);
      syncCalendarEvents(ss);
      Utilities.sleep(800);
      ensureAverageColumns(ss);
      toastOk("Full flow executed without errors");
    }
  } catch (err) {
    toastErr("iniciarSistema: " + err.message);
    Logger.log("iniciarSistema ERROR:\n" + err.stack);
  }
}

// ------------------------------------------------------------
// 10. onEdit - Reference auto-sync + Daily checkbox -> Calendar
// ------------------------------------------------------------
function onEditTrigger(e) {
  try {
    var ss      = e.source;
    var sheet   = e.range.getSheet();
    var editRow = e.range.getRow();
    var editCol = e.range.getColumn();

    // A) Edicion en Reference -> actualizar Average y Daily
    var refMeta      = getTableMeta(ss, CFG.RANGE_REFERENCE);
    var refColStart  = refMeta.startCol;
    var refColEnd    = refMeta.startCol + refMeta.numCols - 1;
    var refRowStart  = refMeta.dataStartRow;

    if (sheet.getName() === refMeta.sheet.getName() &&
        editRow >= refRowStart &&
        editCol >= refColStart &&
        editCol <= refColEnd) {

      Logger.log("onEdit: Reference edited row:" + editRow + " col:" + editCol);
      toast("Reference edited - updating...", "Sync", 3);
      Utilities.sleep(600);
      ensureAverageColumns(ss);
      Utilities.sleep(600);
      populateDailyTable(ss);
      toastOk("Average and Daily updated");
      return;
    }

    // B) Checkbox en Daily -> sync Calendar
    var dailyMeta   = getTableMeta(ss, CFG.RANGE_DAILY);
    var checkAbsCol = dailyMeta.startCol + colIdx(dailyMeta.headerMap, CFG.COL_CHECK);

    if (sheet.getName() === dailyMeta.sheet.getName() &&
        editCol === checkAbsCol &&
        editRow >= dailyMeta.dataStartRow) {

      var checked = e.range.getValue() === true;
      Logger.log("onEdit: Daily checkbox row:" + editRow + " checked:" + checked);
      toast(checked ? "Checking in Calendar..." : "Unchecking in Calendar...", "Sync", 3);
      syncCheckboxToCalendar(ss);
    }

  } catch (err) {
    toastErr("onEdit: " + err.message);
    Logger.log("onEditTrigger ERROR:\n" + err.stack);
  }
}

// ------------------------------------------------------------
// 11. TRIGGER INSTALLATION (run ONCE)
// ------------------------------------------------------------
function installTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  toastInfo("Removing previous triggers...");
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });

  // 1. Setup matutino 00:01
  ScriptApp.newTrigger("iniciarSistema").timeBased().everyDays(1).atHour(0).nearMinute(1).create();
  // 2. Snapshot nocturno 23:50
  ScriptApp.newTrigger("iniciarSistema").timeBased().everyDays(1).atHour(23).nearMinute(50).create();
  // 3. onEdit instalable
  ScriptApp.newTrigger("onEditTrigger").forSpreadsheet(ss).onEdit().create();
  // 4. Polling Calendar->Daily cada 5 min
  ScriptApp.newTrigger("syncCalendarToDaily").timeBased().everyMinutes(5).create();

  toast("4 triggers installed. System active!", "Setup", 6);
  Logger.log("Triggers installed: iniciarSistema x2, onEditTrigger, syncCalendarToDaily");
}

// ------------------------------------------------------------
// 12. CUSTOM MENU
// ------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Daily Tracker")
    .addItem("Start system (full flow)",             "iniciarSistema")
    .addSeparator()
    .addItem("Only: load today's Daily",            "runOnlyDaily")
    .addItem("Only: sync Google Calendar",          "runOnlyCalendar")
    .addItem("Only: Daily -> Calendar (checkboxes)","runOnlyCheckbox")
    .addItem("Only: Calendar -> Daily (polling)",   "runSyncCalToDaily")
    .addItem("Simulate 23:55 snapshot -> Average",  "runOnlySnapshot")
    .addItem("Debug snapshot (check Logger)",       "runDebugSnapshot")
    .addSeparator()
    .addItem("Diagnostic (toasts + Logger)",        "runDiagnostic")
    .addItem("Install triggers (run once)",         "installTrigger")
    .addToUi();
}

function runOnlyDaily()      { populateDailyTable(SpreadsheetApp.getActiveSpreadsheet()); }
function runOnlyCalendar()   { syncCalendarEvents(SpreadsheetApp.getActiveSpreadsheet()); }
function runOnlyCheckbox()   { syncCheckboxToCalendar(SpreadsheetApp.getActiveSpreadsheet()); }
function runSyncCalToDaily() { syncCalendarToDaily(SpreadsheetApp.getActiveSpreadsheet()); }
function runOnlySnapshot()   { saveDailySnapshot(SpreadsheetApp.getActiveSpreadsheet()); }

// ------------------------------------------------------------
// 13. DEBUG SNAPSHOT
// ------------------------------------------------------------
function runDebugSnapshot() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var log = [];

  var refs     = getReference(ss);
  var refIcons = refs.map(function(r) { return r.icon; });
  log.push("=== REFERENCE (" + refs.length + " icons) ===");
  refs.forEach(function(r) { log.push('  "' + r.icon + '" -> "' + r.category + '"'); });

  var res  = getTableData(ss, CFG.RANGE_DAILY);
  var meta = res.meta;
  var data = res.data;
  log.push("=== DAILY (" + data.length + " rows) ===");

  if (data.length === 0) {
    log.push("  EMPTY - run Start system first");
  } else {
    var iAct   = colIdx(meta.headerMap, CFG.COL_ACTIVITY);
    var iRange = colIdx(meta.headerMap, CFG.COL_RANGE);
    var iCheck = colIdx(meta.headerMap, CFG.COL_CHECK);

    var count = {};
    refIcons.forEach(function(ic) { count[ic] = { done: 0, total: 0 }; });

    data.forEach(function(row, i) {
      var act     = String(row[iAct]).trim();
      var range   = String(row[iRange]).trim();
      var checked = row[iCheck] === true;
      var icon    = refIcons.find(function(ic) { return act.indexOf(ic) === 0; }) || "?";
      log.push("  fila " + (i+1) + ": " + act + " | " + range + " | checked:" + checked + " | icon:" + icon);
      if (icon !== "?" && count[icon]) {
        count[icon].total++;
        if (checked) count[icon].done++;
      }
    });

    log.push("=== COUNT ===");
    var total = 0, done = 0;
    refIcons.forEach(function(ic) {
      log.push("  " + ic + ": " + count[ic].done + "/" + count[ic].total);
      total += count[ic].total;
      done  += count[ic].done;
    });
    var eff = total > 0 ? Math.round((done/total)*100) : 0;
    log.push("  Efficiency: " + eff + "% (" + done + "/" + total + ")");
  }

  try {
    var avgMeta = getTableMeta(ss, CFG.RANGE_AVERAGE);
    log.push("=== AVERAGE_ENC ===");
    log.push("  hoja:" + avgMeta.sheet.getName() +
             " headerRow:" + avgMeta.headerRow +
             " dataStartRow:" + avgMeta.dataStartRow);
    log.push("  headers: " + JSON.stringify(avgMeta.headers));
  } catch(e) {
    log.push("  ERROR Average: " + e.message);
  }

  log.forEach(function(l) { Logger.log(l); });
  toast(data.length > 0
    ? "Daily: " + data.length + " row(s). Check Logger (Ctrl+Enter)"
    : "Daily EMPTY - run Start system first",
    "Debug Snapshot", 6);
}

// ------------------------------------------------------------
// 14. DIAGNOSTIC
// ------------------------------------------------------------
function runDiagnostic() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var log = [];

  var rangeNames = [CFG.RANGE_ROUTINE, CFG.RANGE_DAILY, CFG.RANGE_REFERENCE, CFG.RANGE_AVERAGE];
  log.push("=== NAMED RANGES ===");
  rangeNames.forEach(function(name) {
    var nr = ss.getRangeByName(name);
    if (!nr) {
      log.push('  NO EXISTE: "' + name + '"');
    } else {
      log.push('  OK: "' + name + '" -> hoja:"' + nr.getSheet().getName() +
               '" fila:' + nr.getRow() + ' col:' + nr.getColumn() + ' cols:' + nr.getNumColumns());
      log.push('     Headers: ' + JSON.stringify(nr.getValues()[0]));
    }
  });

  log.push("=== REFERENCE ===");
  var refs = [];
  try {
    refs = getReference(ss);
    if (refs.length === 0) {
      log.push("  EMPTY");
    } else {
      refs.forEach(function(r) {
        log.push('  "' + r.icon + '" (U+' + r.icon.codePointAt(0).toString(16).toUpperCase() + ') -> "' + r.category + '"');
      });
    }
  } catch(e) { log.push("  ERROR: " + e.message); }

  log.push("=== TODAY'S BLOCKS ===");
  var blockCount = 0;
  try {
    var blocks = getDayBlocks(ss, todayDayName());
    blockCount = blocks.length;
    if (blocks.length === 0) {
      log.push("  No blocks");
    } else {
      blocks.forEach(function(b) {
        log.push("  " + b.icon + " " + b.category + " -> " + b.rangeStr);
      });
    }
  } catch(e) { log.push("  ERROR: " + e.message); }

  log.forEach(function(l) { Logger.log(l); });

  var rangesOk = rangeNames.filter(function(n) { return ss.getRangeByName(n); }).length;
  toast("Ranges OK: " + rangesOk + "/4 | Reference: " + refs.length + " icon(s)", "Diagnostic 1/2", 5);
  Utilities.sleep(5500);
  if (blockCount === 0) {
    toast("No blocks today (" + todayDayName() + "). Check Logger.", "Diagnostic 2/2", 7);
  } else {
    toast(blockCount + " block(s) for today - ready!", "Diagnostic 2/2", 5);
  }
}
