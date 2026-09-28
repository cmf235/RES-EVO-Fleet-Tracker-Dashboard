function buildCalendar() {
  try {


  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var cal = ss.getSheetByName('Calendar');
  if (!cal) {
    cal = ss.insertSheet('Calendar');
  } else {
    // Clear content and formatting but preserve drawings
    var lr = cal.getMaxRows();
    var lc = cal.getMaxColumns();
    cal.clear();
    if (lr > 0 && lc > 0) {
      cal.getRange(1, 1, lr, lc).clearFormat();
      cal.getRange(1, 1, lr, lc).breakApart();
    }
    if (lr > 300) { cal.deleteRows(301, lr - 300); } else if (lr < 300) { cal.insertRowsAfter(lr, 300 - lr); }
    if (lc > 26) cal.deleteColumns(27, lc - 26);
  }
  ss.setActiveSheet(cal);

  // Dynamic 3-month window from today
  var today = new Date();
  today.setHours(0,0,0,0);
  var endDate = new Date(today);
  endDate.setMonth(endDate.getMonth() + 3);

  var events = collectEvents(ss);
  Logger.log("EVTS:" + JSON.stringify(events.map(function(e){return {l:e.label,t:e.type,s:String(e.startDate),e:String(e.endDate)};})));

  // Color palette - modern professional
  var C = {
    // Event colors
    rdBg:      '#C62828',  // deep red
    rdFg:      '#FFFFFF',
    rdlBg:     '#E57373',  // medium red/rose
    rdlFg:     '#FFFFFF',
    ruBg:      '#1B5E20',  // deep green
    ruFg:      '#FFFFFF',
    rulBg:     '#66BB6A',  // medium green
    rulFg:     '#FFFFFF',
        roBg:       '#7B1FA2',  // purple
            roFg:       '#FFFFFF',
    miBg:       '#E65100' ,  // deep orange
    miFg:       '#FFFFFF' ,

    // Chrome
    titleBg:   '#0D1B2A',  // near black navy
    titleFg:   '#FFFFFF',
    subtitleBg:'#1A2F45',
    subtitleFg:'#90CAF9',
    legendBg:  '#F0F4F8',
    legendFg:  '#37474F',
    monthBg:   '#1565C0',  // strong blue
    monthFg:   '#FFFFFF',
    dowBg:     '#1976D2',  // lighter blue
    dowFg:     '#FFFFFF',
    wkBg:      '#E3EAF4',
    wkFg:      '#1565C0',
    todayBg:   '#FFF9C4',
    todayFg:   '#E65100',
    todayBorder:'#F57F17',
    weekendBg: '#ECEFF1',
    weekendFg: '#607D8B',
    dayBg:     '#FAFAFA',
    dayFg:     '#37474F',
    emptyBg:   '#F5F5F5',
    outBg:     '#E0E0E0',
    outFg:     '#BDBDBD',
    border:    '#CFD8DC',
    spacer:    '#E8EDF2'
  };

  var ICONS = {
    rigdown:     'ARROW_DOWN  RIG DOWN',
    rigdown_log: 'TRUCK_R  RD LOGISTICS',
    rigup:       'ARROW_UP  RIG UP',
    rigup_log:   'TRUCK_L  RU LOGISTICS',
        roustabout:   'WRENCH  ROUSTABOUT'
  };

  var DAYS = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'];

  // Column setup: col 1 = week label, cols 2-8 = Mon-Sun
  cal.setColumnWidth(1, 82);
  for (var d = 0; d < 7; d++) {
    cal.setColumnWidth(d + 2, 152);
  }

  var ROW = 1;

  // ── ROW 1: Title ──────────────────────────────────────────────
  cal.setRowHeight(ROW, 52);
  var tr = cal.getRange(ROW, 1, 1, 8);
  tr.merge();
  tr.setValue('EVOLUTION SERVICE CALENDAR');
  tr.setBackground(C.titleBg);
  tr.setFontColor(C.titleFg);
  tr.setFontSize(17);
  tr.setFontWeight('bold');
  tr.setFontFamily('Arial');
  tr.setHorizontalAlignment('center');
  tr.setVerticalAlignment('middle');
  ROW++;

  // ── ROW 2: Subtitle ──────────────────────────────────────────
  cal.setRowHeight(ROW, 22);
  var sr = cal.getRange(ROW, 1, 1, 8);
  sr.merge();
  sr.setValue(formatDateStr(today) + '  —  ' + formatDateStr(endDate) + '     |     Refreshed: ' + formatDateStr(today));
  sr.setBackground(C.subtitleBg);
  sr.setFontColor(C.subtitleFg);
  sr.setFontSize(9);
  sr.setFontFamily('Arial');
  sr.setHorizontalAlignment('center');
  sr.setVerticalAlignment('middle');
  ROW++;

  // ── ROW 3: Legend ─────────────────────────────────────────────
  cal.setRowHeight(ROW, 28);
  var legItems = [
    { col:1, text:'LEGEND',           bg:C.legendBg, fg:'#546E7A', bold:true, sz:8 },
    { col:2, text:'⬇️  RIG DOWN',      bg:C.rdBg,  fg:C.rdFg,  bold:true, sz:9 },
    { col:3, text:'🚚  RD LOGISTICS',bg:C.rdlBg, fg:C.rdlFg, bold:false,sz:9 },
    { col:4, text:'⬆️  RIG UP',        bg:C.ruBg,  fg:C.ruFg,  bold:true, sz:9 },
    { col:5, text:'🚛  RU LOGISTICS',bg:C.rulBg, fg:C.rulFg, bold:false,sz:9 },
        { col:6, text:'\uD83D\uDD27  ROUSTABOUT',bg:C.roBg, fg:C.roFg, bold:false,sz:9 },
    { col:7, text:'\u2699\uFE0F MISC SERVICES', bg:C.miBg, fg:C.miFg, bold:false, sz:9 },
    { col:8, text:'\u2B50 Today',          bg:'#FFF9C4', fg:'#E65100',bold:false,sz:9 }
  ];
  legItems.forEach(function(it) {
    var lc = cal.getRange(ROW, it.col);
    lc.setValue(it.text);
    lc.setBackground(it.bg);
    lc.setFontColor(it.fg);
    lc.setFontWeight(it.bold ? 'bold' : 'normal');
    lc.setFontSize(it.sz);
    lc.setFontFamily('Arial');
    lc.setHorizontalAlignment('center');
    lc.setVerticalAlignment('middle');
  });
  ROW++;

  // ── THIN SEPARATOR ────────────────────────────────────────────
  cal.setRowHeight(ROW, 4);
  cal.getRange(ROW, 1, 1, 8).setBackground(C.titleBg);
  ROW++;

  // Freeze top 4 rows (title + subtitle + legend + separator)
  cal.setFrozenRows(4);

  // ── MONTH BLOCKS ──────────────────────────────────────────────
  var monthNames = ['January','February','March','April','May','June',
                    'July','August','September','October','November','December'];

  // Iterate month by month
  var cursorMonth = new Date(today.getFullYear(), today.getMonth(), 1);
  var endMonth    = new Date(endDate.getFullYear(), endDate.getMonth(), 1);

  while (cursorMonth <= endMonth) {
    var yr  = cursorMonth.getFullYear();
    var mo  = cursorMonth.getMonth();
    var daysInMonth = new Date(yr, mo + 1, 0).getDate();

    // Month header
    cal.setRowHeight(ROW, 36);
    var mh = cal.getRange(ROW, 1, 1, 8);
    mh.merge();
    mh.setValue(monthNames[mo] + '  ' + yr);
    mh.setBackground(C.monthBg);
    mh.setFontColor(C.monthFg);
    mh.setFontSize(13);
    mh.setFontWeight('bold');
    mh.setFontFamily('Arial');
    mh.setHorizontalAlignment('center');
    mh.setVerticalAlignment('middle');
    ROW++;

    // Day-of-week header
    cal.setRowHeight(ROW, 22);
    var blank = cal.getRange(ROW, 1);
    blank.setValue('');
    blank.setBackground(C.dowBg);
    DAYS.forEach(function(dn, idx) {
      var hc = cal.getRange(ROW, idx + 2);
      hc.setValue(dn);
      hc.setBackground(C.dowBg);
      hc.setFontColor(C.dowFg);
      hc.setFontWeight('bold');
      hc.setFontSize(10);
      hc.setFontFamily('Arial');
      hc.setHorizontalAlignment('center');
      hc.setVerticalAlignment('middle');
    });
    ROW++;

    // Build weeks - only days that belong to THIS month
    // Find first Monday on or before the 1st of this month
    var firstDay = new Date(yr, mo, 1);
    var fdow = firstDay.getDay(); // 0=Sun
    var weekStart = new Date(firstDay);
    var backDays = (fdow === 0) ? 6 : fdow - 1;
    weekStart.setDate(weekStart.getDate() - backDays);

    while (true) {
      // Check if this week has any days in this month
      var weekEnd = new Date(weekStart);
      weekEnd.setDate(weekEnd.getDate() + 6);
      var lastOfMonth = new Date(yr, mo, daysInMonth);

      if (weekStart > lastOfMonth) break;

      // Pre-scan: count max events per day this week
      var maxEvts = 3; // minimum 3 event rows
      for (var scanDi = 0; scanDi < 7; scanDi++) {
        var scanDate = new Date(weekStart);
        scanDate.setDate(scanDate.getDate() + scanDi);
        scanDate.setHours(0,0,0,0);
        var scanCount = events.filter(function(e) {
          return dayWithinRange(scanDate, e.startDate, e.endDate);
        }).length;
        if (scanCount > maxEvts) maxEvts = scanCount;
      }
      var weekRows = 1 + maxEvts; // 1 date row + N event rows

      // Row heights: date row + dynamic event rows
      cal.setRowHeight(ROW,     18);
      for (var rh = 0; rh < maxEvts; rh++) {
        cal.setRowHeight(ROW + 1 + rh, 30);
      }

      // Week label (col A, spans 4 rows)
      // Show first day of THIS month in this week
      var wkLabelDay = (weekStart.getMonth() === mo) ? weekStart.getDate() : 1;
      var wkLbl = cal.getRange(ROW, 1, weekRows, 1);
      wkLbl.merge();
      wkLbl.setValue((mo+1) + '/' + wkLabelDay);
      wkLbl.setBackground(C.wkBg);
      wkLbl.setFontColor(C.wkFg);
      wkLbl.setFontWeight('bold');
      wkLbl.setFontSize(8);
      wkLbl.setFontFamily('Arial');
      wkLbl.setHorizontalAlignment('center');
      wkLbl.setVerticalAlignment('middle');
      wkLbl.setWrap(true);

      // Fill 7 columns (Mon-Sun)
      for (var di = 0; di < 7; di++) {
        var dayDate = new Date(weekStart);
        dayDate.setDate(dayDate.getDate() + di);

        var inThisMonth = (dayDate.getMonth() === mo && dayDate.getFullYear() === yr);
        var inRange = dayDate >= today && dayDate <= endDate;
        var isToday = sameDay(dayDate, today);
        var isWeekend = (di === 5 || di === 6);
        var col = di + 2;

        // Date number cell
        var dc = cal.getRange(ROW, col);
        dc.setFontSize(9);
        dc.setFontFamily('Arial');
        dc.setFontWeight('bold');
        dc.setHorizontalAlignment('right');
        dc.setVerticalAlignment('middle');

        if (!inThisMonth) {
          // Grey out days not in this month
          dc.setValue('');
          dc.setBackground(C.outBg);
          dc.setFontColor(C.outFg);
          for (var er = 0; er < maxEvts; er++) {
            cal.getRange(ROW + 1 + er, col).setBackground(C.outBg);
          }
        } else {
          dc.setValue(dayDate.getDate());
          if (isToday) {
            dc.setBackground(C.todayBg);
            dc.setFontColor(C.todayFg);
          } else if (isWeekend) {
            dc.setBackground(C.weekendBg);
            dc.setFontColor(C.weekendFg);
          } else {
            dc.setBackground(C.dayBg);
            dc.setFontColor(C.dayFg);
          }

      var dayEvs = events.filter(function(e) { return dayWithinRange(dayDate, e.startDate, e.endDate); });

          for (var er = 0; er < maxEvts; er++) {
            var evc = cal.getRange(ROW + 1 + er, col);
            evc.setFontSize(8);
            evc.setFontFamily('Arial');
            evc.setHorizontalAlignment('center');
            evc.setVerticalAlignment('middle');
            evc.setWrap(true);

            if (er < dayEvs.length) {
              var ev = dayEvs[er];
              var icon = '';
              var bgc = '', fgc = '', bold = false;
              if (ev.type === 'rigdown') {
                icon = '⬇️'; bgc = C.rdBg;  fgc = C.rdFg;  bold = true;
              } else if (ev.type === 'rigdown_log') {
                icon = '🚚'; bgc = C.rdlBg; fgc = C.rdlFg; bold = false;
              } else if (ev.type === 'rigup') {
                icon = '⬆️'; bgc = C.ruBg;  fgc = C.ruFg;  bold = true;
              } else if (ev.type === 'roustabout') {
                              icon = '\uD83D\uDD27'; bgc = C.roBg;  fgc = C.roFg;  bold = false;
              } else if (ev.type === 'misc') {
            icon = '\u2699\uFE0F'; bgc = C.miBg; fgc = C.miFg; bold = false;
          } else {
                icon = '🚛'; bgc = C.rulBg; fgc = C.rulFg; bold = false;
              }
              evc.setValue(icon + '  ' + ev.label);
              evc.setBackground(bgc);
              evc.setFontColor(fgc);
              evc.setFontWeight(bold ? 'bold' : 'normal');
            } else {
              evc.setBackground(isWeekend ? C.weekendBg : C.emptyBg);
            }
          }
        }
      }

      // Outer border for this week row
      cal.getRange(ROW, 1, weekRows, 8)
         .setBorder(true, true, true, true, false, false, C.border, SpreadsheetApp.BorderStyle.SOLID);
      // Inner vertical borders between day columns
      cal.getRange(ROW, 2, weekRows, 7)
         .setBorder(null, null, null, null, true, null, C.border, SpreadsheetApp.BorderStyle.SOLID);

      ROW += weekRows;

      // Move to next week
      weekStart.setDate(weekStart.getDate() + 7);
    }

    // Spacer between months
    cal.setRowHeight(ROW, 10);
    cal.getRange(ROW, 1, 1, 8).setBackground(C.spacer);
    ROW++;

    // Advance to next month
    cursorMonth.setMonth(cursorMonth.getMonth() + 1);
  }

  cal.setTabColor('#1565C0');
  Logger.log('Done. Events: ' + events.length);
  events.forEach(function(e) { Logger.log(e.type + ' | ' + e.label + ' | ' + Utilities.formatDate(e.startDate, 'America/New_York', 'MM/dd/yyyy') + ' - ' + Utilities.formatDate(e.endDate, 'America/New_York', 'MM/dd/yyyy')); });
  } catch (_bcErr) { notifyError_('buildCalendar', _bcErr); throw _bcErr; }
}

/* ===== C-3: single source of truth — panel cell map (shared by buildDashboardData + collectEvents). Change a cell here and BOTH the feed and the in-sheet calendar follow. ===== */
var PANEL_MAP = {
  rigUp: {
    location:'J3', pad:'O3', date:'L4', time:'L4', logisticsDate:'S4', logisticsTime:'S4',
    dayShift:'J5', nightShift:'O5', hotel:'T5',
    monolineIron:'J6', pumpdownIron:'J7', lowPressureIron:'J8', flowbackIron:'J9',
    highPressureHoses:'J10', rigMats:'J11', pullCables:'J12',
    loadTransports:'J14', loadQty:'X14', unloadTransports:'J15', unloadQty:'X15', semiHotshot:'J16', semiQty:'X16',
    roustabout:{ start:'M19', startTime:'M19', end:'T19', dayShift:'J20', dayQty:'X20', nightShift:'J21', nightQty:'X21' },
    misc:{ start:'M24', startTime:'M24', end:'T24', desc:'M23' }
  },
  rigDown: {
    location:'Z3', pad:'AE3', date:'AB4', time:'AB4', logisticsDate:'AI4', logisticsTime:'AI4',
    dayShift:'Z5', nightShift:'AE5', hotel:'AJ5',
    monolineIron:'Z6', pumpdownIron:'Z7', lowPressureIron:'Z8', flowbackIron:'Z9',
    highPressureHoses:'Z10', rigMats:'Z11', pullCables:'Z12',
    loadTransports:'Z14', loadQty:'AN14', unloadTransports:'Z15', unloadQty:'AN15', semiHotshot:'Z16', semiQty:'AN16',
    roustabout:{ start:'AC19', startTime:'AC19', end:'AJ19', dayShift:'Z20', dayQty:'AN20', nightShift:'Z21', nightQty:'AN21' },
    misc:{ start:'AC24', startTime:'AC24', end:'AJ24', desc:'AC23' }
  }
};

function collectEvents(ss) {
  var events = [];

  function parseDate(val) {
    if (!val) return null;
    if (val instanceof Date) {
      var y = val.getFullYear();
      if (y < 2024 || y > 2030) return null;
      var d = new Date(val); d.setHours(0,0,0,0); return d;
    }
    var s = String(val).trim();
    if (!s || s === '?' || s.indexOf('00/00') !== -1) return null;
    var m = s.match(/(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})/);
    if (!m) return null;
    var mo = parseInt(m[1]) - 1, da = parseInt(m[2]), yr = parseInt(m[3]);
    if (yr < 100) yr += 2000;
    if (yr < 2024 || yr > 2030) return null;
    return new Date(yr, mo, da);
  }

  function lbl(sh, num, opCell, padCell) {
    var op  = String(sh.getRange(opCell).getValue()).trim()  || '?';
    var pad = String(sh.getRange(padCell).getValue()).trim() || '?';
    return 'F#' + num + ' ' + op + ' | ' + pad;
  }

  function addRange(label, type, sh, startCell, endCell) {
    var rawVal = sh.getRange(startCell).getValue(); Logger.log("ADD:" + type + " cell=" + startCell + " raw=[" + rawVal + "] type=" + typeof rawVal + " isDate=" + (rawVal instanceof Date));
    var sd = parseDate(sh.getRange(startCell).getValue());
    var ed = endCell ? parseDate(sh.getRange(endCell).getValue()) : null;
    if (!sd && !ed) return;
    if (sd && !ed) ed = sd;
    if (!sd && ed) sd = ed;
    if (ed < sd) { var t = sd; sd = ed; ed = t; }
    events.push({ label: label, type: type, startDate: sd, endDate: ed });
  }

  var fleets = {
    '1':  'FLEET # 1 - Lindsay Riley',
    '6':  'FLEET # 6 - Brandon Dock',
    '9':  'FLEET # 9 - Eric Cronberg',
    '10': 'FLEET # 10 - Corey Miller',
    '13': 'FLEET # 13 - Ryan Poorman',
    '15': 'FLEET # 15 - Jeremy Bower'
  };

  Object.keys(fleets).forEach(function(num) {
    var sh = ss.getSheetByName(fleets[num]);
    if (!sh) return;

    var cur = lbl(sh, num, PANEL_MAP.rigDown.location, PANEL_MAP.rigDown.pad);
    addRange(cur, 'rigdown',      sh, PANEL_MAP.rigDown.date, null);
    addRange(cur, 'rigdown_log',  sh, PANEL_MAP.rigDown.logisticsDate, null);
    addRange(cur, 'roustabout',  sh, PANEL_MAP.rigDown.roustabout.start, PANEL_MAP.rigDown.roustabout.end);
    addRange(cur, 'misc',        sh, PANEL_MAP.rigDown.misc.start, PANEL_MAP.rigDown.misc.end);

    var up = lbl(sh, num, PANEL_MAP.rigUp.location, PANEL_MAP.rigUp.pad);
    addRange(up,  'rigup',        sh, PANEL_MAP.rigUp.date, null);
    addRange(up,  'rigup_log',    sh, PANEL_MAP.rigUp.logisticsDate, null);
    addRange(up,  'roustabout',  sh, PANEL_MAP.rigUp.roustabout.start, PANEL_MAP.rigUp.roustabout.end);
    addRange(up,  'misc',        sh, PANEL_MAP.rigUp.misc.start, PANEL_MAP.rigUp.misc.end);
  });

  return events;
}

function dayWithinRange(d, s, e) {
  if (!d || !s || !e) return false;
  var dDay = new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
  var sDay = new Date(s.getFullYear(), s.getMonth(), s.getDate()).getTime();
  var eDay = new Date(e.getFullYear(), e.getMonth(), e.getDate()).getTime();
  return dDay >= sDay && dDay <= eDay;
}

function sameDay(a, b) {
  return a.getFullYear() === b.getFullYear() &&
         a.getMonth()    === b.getMonth()    &&
         a.getDate()     === b.getDate();
}

function formatDateStr(d) {
  var m = ['January','February','March','April','May','June',
           'July','August','September','October','November','December'];
  return m[d.getMonth()] + ' ' + d.getDate() + ', ' + d.getFullYear();
}

function buildCalendarMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('Evolution Calendar')
    .addItem('Refresh Calendar', 'buildCalendar')
    .addToUi();
}

function formatFleetSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets().filter(function(s) { return s.getName().indexOf('FLEET #') === 0; });
  var dtFmt = 'M/d/yyyy h:mm:ss AM/PM';
  sheets.forEach(function(sheet) {
    for (var col = 10; col <= 24; col++) sheet.setColumnWidth(col, 36);
    for (var col = 26; col <= 40; col++) sheet.setColumnWidth(col, 36);
    var range = sheet.getDataRange();
    var values = range.getValues();
    for (var r = 0; r < values.length; r++) {
      for (var c = 0; c < values[r].length; c++) {
        if (values[r][c] instanceof Date) {
          sheet.getRange(r + 1, c + 1).setNumberFormat(dtFmt);
        }
      }
    }
  });
}

/**
 * Simple trigger: auto-format date cells on fleet tabs.
 * Whenever a user types a date/time value into any cell on a
 * FLEET # tab, this sets the display format to M/d/yyyy h:mm:ss AM/PM.
 */
function onEdit(e) {
  try { CacheService.getScriptCache().remove('dashData'); } catch (_ce0) {}
  /* Claude: enforce standard date format on fleet tabs */
  try { var _e = arguments[0];
    if (_e && _e.range) {
      var _sh = _e.range.getSheet();
      if (/FLEET\s*#\s*\d+/i.test(_sh.getName())) {
        var _rg = _e.range, _vl = _rg.getValues(), _ft = _rg.getNumberFormats(), _hit = false;
        for (var _i = 0; _i < _vl.length; _i++) for (var _j = 0; _j < _vl[_i].length; _j++) {
          if (_vl[_i][_j] instanceof Date) { _ft[_i][_j] = "m/d/yy hh:mm AM/PM"; _hit = true; }
        }
        if (_hit) _rg.setNumberFormats(_ft);
      }
    }
  } catch (_er) {}

  var sheet = e.source.getActiveSheet();
  if (sheet.getName().indexOf("FLEET #") !== 0) return;
  var range = e.range;
  var value = range.getValue();
  if (value instanceof Date) {
    range.setNumberFormat("M/d/yyyy h:mm:ss AM/PM");
  }
}


/**
 * Creates a time-driven trigger that runs buildCalendar nightly at midnight.
 * Run this function ONCE manually to set up the trigger.
 * To remove: go to Edit > Current project triggers and delete it.
 */
function createAutoRefreshTrigger() {
  // Delete any existing buildCalendar triggers to avoid duplicates
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'buildCalendar') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // Create the nightly (midnight) trigger
  ScriptApp.newTrigger('buildCalendar')
    .timeBased()
    .atHour(0)
      .everyDays(1)
    .create();
  Logger.log('Auto-refresh trigger created: buildCalendar will run nightly at midnight.');
}

/**
 * Removes the auto-refresh trigger.
 */
function removeAutoRefreshTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'buildCalendar') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  Logger.log('Removed ' + removed + ' trigger(s).');
}


function diagnoseMergedCells() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
    var fleets = {
        '1':  'FLEET # 1 - Lindsay Riley',
            '6':  'FLEET # 6 - Brandon Dock',
                '9':  'FLEET # 9 - Eric Cronberg',
                    '10': 'FLEET # 10 - Corey Miller',
                        '13': 'FLEET # 13 - Ryan Poorman',
                            '15': 'FLEET # 15 - Jeremy Bower'
                              };
                                var checks = [
                                    {name:'cur_roust_start', cell:'M19'},
                                        {name:'cur_roust_end',   cell:'T19'},
                                            {name:'cur_misc_start',  cell:'M24'},
                                                {name:'cur_misc_end',    cell:'T24'},
                                                    {name:'up_roust_start',  cell:'AC19'},
                                                        {name:'up_roust_end',    cell:'AJ19'},
                                                            {name:'up_misc_start',   cell:'AC24'},
                                                                {name:'up_misc_end',     cell:'AJ24'}
                                                                  ];
                                                                    Object.keys(fleets).forEach(function(num) {
                                                                        var sh = ss.getSheetByName(fleets[num]);
                                                                            if (!sh) { Logger.log('Fleet #' + num + ': SHEET NOT FOUND'); return; }
                                                                                Logger.log('=== Fleet #' + num + ' ===');
                                                                                    checks.forEach(function(c) {
                                                                                          var r = sh.getRange(c.cell);
                                                                                                var val = r.getValue();
                                                                                                      var merges = r.getMergedRanges();
                                                                                                            if (merges.length === 0) {
                                                                                                                    Logger.log('  ' + c.name + ' (' + c.cell + '): not merged, val=[' + val + ']');
                                                                                                                          } else {
                                                                                                                                  var m = merges[0];
                                                                                                                                          var primaryCol = m.getColumn();
                                                                                                                                                  var primaryRow = m.getRow();
                                                                                                                                                          var primaryA1  = String.fromCharCode(64 + primaryCol) + primaryRow;
                                                                                                                                                                  if (primaryCol > 26) primaryA1 = String.fromCharCode(64 + Math.floor((primaryCol-1)/26)) + String.fromCharCode(65 + (primaryCol-1)%26) + primaryRow;
                                                                                                                                                                          var isCurrent  = (r.getColumn() === primaryCol && r.getRow() === primaryRow);
                                                                                                                                                                                  Logger.log('  ' + c.name + ' (' + c.cell + '): merged=' + m.getA1Notation() + ' primary=' + primaryA1 + ' OK=' + isCurrent + ' val=[' + val + ']');
                                                                                                                                                                                        }
                                                                                                                                                                                            });
                                                                                                                                                                                              });
                                                                                                                                                                                              }
                                                                                                                                                                                              

// ============================================================
// WEB APP: doGet() — JSON API endpoint for the live dashboard
// Deploy as: Execute as Me / Anyone can access
// ============================================================
function doGet(e) {
  // Debug: return raw row5 values
  if (e && e.parameter && e.parameter.debug === 'row5') {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheets().find(function(s){ return s.getName().indexOf('Lindsay Riley') >= 0; });
    var row5 = {};
    ['J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ'].forEach(function(col){
      row5[col+'5'] = sh.getRange(col+'5').getValue();
    });
    return ContentService.createTextOutput(JSON.stringify(row5)).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var _cache = CacheService.getScriptCache();
    var _bypass = e && e.parameter && (e.parameter.fresh === '1' || e.parameter.nocache === '1');
    var _json = _bypass ? null : _cache.get('dashData');
    if (!_json) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var data = buildDashboardData(ss);
      _json = JSON.stringify(data);
      try { if (_json && _json.length < 90000) _cache.put('dashData', _json, 90); } catch (_ce) {}
    }
    var output = ContentService.createTextOutput(_json);
    output.setMimeType(ContentService.MimeType.JSON);
    return output;
  } catch (err) {
    notifyError_('doGet', err);
    var output = ContentService.createTextOutput(JSON.stringify({ error: err.toString(), lastUpdated: new Date().toISOString() }));
    output.setMimeType(ContentService.MimeType.JSON);
    return output;
  }
}

function notifyError_(w,e){}   function buildDashboardData(ss) {
  var fleets = {
    '1':  'FLEET # 1 - Lindsay Riley',
    '6':  'FLEET # 6 - Brandon Dock',
    '9':  'FLEET # 9 - Eric Cronberg',
    '10': 'FLEET # 10 - Corey Miller',
    '13': 'FLEET # 13 - Ryan Poorman',
    '15': 'FLEET # 15 - Jeremy Bower'
  };

  function fmtDate(val) {
    if (!val) return null;
    if (val instanceof Date) {
      var y = val.getFullYear();
      if (y < 2024 || y > 2035) return null;
      return Utilities.formatDate(val, 'America/New_York', 'yyyy-MM-dd');
    }
    var s = String(val).trim();
    if (!s || s === '?' || s.indexOf('00/00') !== -1) return null;
    return s;
  }

  function safeStr(sh, cell) {
    try { return String(_getCell(sh, cell)).trim() || ''; } catch (e2) { return ''; }
  }

  function safeDate(sh, cell) {
    try { return fmtDate(_getCell(sh, cell)); } catch (e2) { return null; }
  }
function fmtTime(val) {
  if (!val || !(val instanceof Date)) return '';
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  return Utilities.formatDate(val, tz, 'h:mm a');
}
function safeTime(sh, cell) { try { return fmtTime(_getCell(sh, cell)); } catch(e) { return ''; } }
function safeNum(sh, cell) { try { var v = _getCell(sh, cell); return (v === '' || v === null || v === undefined) ? 0 : Number(v); } catch(e) { return 0; } }
function safeBool(sh, cell) { try { return _getCell(sh, cell) === true; } catch(e) { return false; } }


  function buildPanel(sh, M) {
    return {
      location:          safeStr(sh, M.location),
      pad:               safeStr(sh, M.pad),
      date:              safeDate(sh, M.date),
      time:              safeTime(sh, M.time),
      logisticsDate:     safeDate(sh, M.logisticsDate),
      logisticsTime:     safeTime(sh, M.logisticsTime),
      dayShift:          safeBool(sh, M.dayShift),
      nightShift:        safeBool(sh, M.nightShift),
      hotel:             safeBool(sh, M.hotel),
      monolineIron:      safeBool(sh, M.monolineIron),
      pumpdownIron:      safeBool(sh, M.pumpdownIron),
      lowPressureIron:   safeBool(sh, M.lowPressureIron),
      flowbackIron:      safeBool(sh, M.flowbackIron),
      highPressureHoses: safeBool(sh, M.highPressureHoses),
      rigMats:           safeBool(sh, M.rigMats),
      pullCables:        safeBool(sh, M.pullCables),
      loadTransports:    safeBool(sh, M.loadTransports), loadQty: safeNum(sh, M.loadQty),
      unloadTransports:  safeBool(sh, M.unloadTransports), unloadQty: safeNum(sh, M.unloadQty),
      semiHotshot:       safeBool(sh, M.semiHotshot), semiQty: safeNum(sh, M.semiQty),
      roustabout: { start: safeDate(sh, M.roustabout.start), startTime: safeTime(sh, M.roustabout.startTime), end: safeDate(sh, M.roustabout.end), dayShift: safeBool(sh, M.roustabout.dayShift), dayQty: safeNum(sh, M.roustabout.dayQty), nightShift: safeBool(sh, M.roustabout.nightShift), nightQty: safeNum(sh, M.roustabout.nightQty) },
      misc: { start: safeDate(sh, M.misc.start), startTime: safeTime(sh, M.misc.startTime), end: safeDate(sh, M.misc.end), desc: safeStr(sh, M.misc.desc) }
    };
  }

  var result = { lastUpdated: new Date().toISOString(), fleets: [] };

  Object.keys(fleets).forEach(function (num) {
    var sh = ss.getSheetByName(fleets[num]);
    if (!sh) return;
    var shortName = fleets[num].replace('FLEET # ' + num + ' - ', '');
    result.fleets.push({ gid: sh.getSheetId(),
      num: num,
      name: fleets[num],
      shortName: shortName,
      rigUp: buildPanel(sh, PANEL_MAP.rigUp),
      rigDown: buildPanel(sh, PANEL_MAP.rigDown)
    });
    attachSheetExtras_(result.fleets[result.fleets.length - 1], sh);
  });

  return result;
}


function testCellValues() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheets().find(s => s.getName().includes('Lindsay Riley'));
  if (!sh) { Logger.log('Sheet not found'); return; }
  Logger.log('J5 (Day Shift): ' + sh.getRange('J5').getValue());
  Logger.log('N5 (Night Shift): ' + sh.getRange('N5').getValue());
  Logger.log('T5 (Hotel): ' + sh.getRange('T5').getValue());
  Logger.log('Z5 (RU Day Shift): ' + sh.getRange('Z5').getValue());
  Logger.log('AD5 (RU Night Shift): ' + sh.getRange('AD5').getValue());
  Logger.log('AH5 (RU Hotel?): ' + sh.getRange('AH5').getValue());
  Logger.log('AJ5 (RU Hotel?): ' + sh.getRange('AJ5').getValue());
  Logger.log('L4 (RD Date+Time): ' + sh.getRange('L4').getValue());
}

/* ===== C-1: batch cell reads (per-request, per-sheet value cache) ===== */
var __cellCache = {};
function _a1(a){ var mm = String(a).match(/^([A-Z]+)(\d+)$/); if(!mm) return null; var col=0; for(var i=0;i<mm[1].length;i++){ col = col*26 + (mm[1].charCodeAt(i)-64); } return { r: parseInt(mm[2],10), c: col }; }
function _getCell(sh, cell){ var rc=_a1(cell); if(!rc) return sh.getRange(cell).getValue(); var id=sh.getSheetId(); var V=__cellCache[id]; if(!V){ V=sh.getRange(1,1,30,45).getValues(); __cellCache[id]=V; } if(rc.r>=1 && rc.r<=V.length && rc.c>=1 && rc.c<=V[0].length){ return V[rc.r-1][rc.c-1]; } return sh.getRange(cell).getValue(); }
