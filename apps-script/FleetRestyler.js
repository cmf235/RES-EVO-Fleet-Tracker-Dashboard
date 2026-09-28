/**
 * FleetRestyler.gs — Clean & Minimal redesign for all fleet tabs
 * Run restyleAllFleets() to apply.
 */

function restyleAllFleets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fleetNames = [
    'FLEET # 1 - Brandon Dock',
    'FLEET # 6 - Lindsay Riley',
    'FLEET # 9 - Eric Cronberg',
    'FLEET # 10 - Corey Miller',
    'FLEET # 13 - Ryan Poorman',
    'FLEET # 15 - Jeremy Bower'
  ];

  fleetNames.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) { Logger.log('SKIP: ' + name); return; }
    Logger.log('Restyling: ' + name);
    restyleFleetSheet(sh);
    applyFleetDateFormats(sh);
  });

  // Also restyle the Calendar tab
  restyleCalendarTab(ss.getSheetByName('Calendar'));

  Logger.log('Done — all fleet tabs restyled.');
}


/* ─────────────────────────── PALETTE ─────────────────────────── */
var P = {
  // Primary
  dark:       '#1E293B',
  darkFg:     '#FFFFFF',
  mid:        '#334155',
  midFg:      '#FFFFFF',
  accent:     '#3B82F6',
  accentFg:   '#FFFFFF',

  // Surfaces
  surface:    '#FFFFFF',
  surfaceAlt: '#F8FAFC',
  muted:      '#F1F5F9',

  // Borders
  border:     '#E2E8F0',
  borderDark: '#CBD5E1',

  // Text
  textPrimary:'#0F172A',
  textSecondary:'#475569',
  textMuted:  '#94A3B8',

  // Status / functional
  green:      '#059669',
  greenLight: '#D1FAE5',
  greenBg:    '#ECFDF5',
  red:        '#DC2626',
  redLight:   '#FEE2E2',
  redBg:      '#FEF2F2',
  yellow:     '#D97706',
  yellowLight:'#FEF3C7',
  yellowBg:   '#FFFBEB',
  purple:     '#7C3AED',
  purpleBg:   '#F5F3FF',
  orange:     '#EA580C',
  orangeBg:   '#FFF7ED',
};


/* ─────────── MAIN FLEET RESTYLE FUNCTION ─────────── */
function restyleFleetSheet(sh) {
  var maxRow = 30;
  var maxCol = 46;
  
  // First, break ALL existing merges in the sheet
  var merges = sh.getRange(1, 1, maxRow, maxCol).getMergedRanges();
  merges.forEach(function(m) { m.breakApart(); });
  
  var all = sh.getRange(1, 1, maxRow, maxCol);
  all.setBackground(P.surface);
  all.setFontFamily('Roboto');
  all.setFontColor(P.textPrimary);
  all.setFontSize(10);
  all.setFontWeight('normal');
  all.setVerticalAlignment('middle');
  all.setBorder(false, false, false, false, false, false);

  // Column widths
  sh.setColumnWidth(1, 30);
  sh.setColumnWidth(2, 75);
  sh.setColumnWidth(3, 75);
  sh.setColumnWidth(4, 75);
  sh.setColumnWidth(5, 75);
  sh.setColumnWidth(6, 75);
  sh.setColumnWidth(7, 30);
  sh.setColumnWidth(8, 8);
  sh.setColumnWidth(9, 8);
  for (var c = 10; c <= 24; c++) sh.setColumnWidth(c, 50);
  sh.setColumnWidth(25, 8);
  for (var c = 26; c <= 40; c++) sh.setColumnWidth(c, 50);

  // Row heights
  sh.setRowHeight(1, 28);
  sh.setRowHeight(2, 4);
  sh.setRowHeight(3, 32);
  sh.setRowHeight(4, 28);
  for (var r = 5; r <= 24; r++) sh.setRowHeight(r, 24);

  styleFleetInfo(sh);
  stylePadPanel(sh, 10, 'down');
  stylePadPanel(sh, 26, 'up');
  styleChecklist(sh);

  sh.setFrozenRows(0);
  sh.setFrozenColumns(0);
}


/* ────────────── FLEET INFO (left panel) ────────────── */
function styleFleetInfo(sh) {
  var r1 = sh.getRange('A1:G1');
  r1.merge().setValue('FLEET').setBackground(P.dark).setFontColor(P.darkFg)
    .setFontSize(9).setFontWeight('bold').setHorizontalAlignment('center');

  sh.getRange('A2:G2').merge().setBackground(P.surface);

  var r3 = sh.getRange('A3:G3');
  r3.setBackground(P.accent).setFontColor(P.accentFg)
    .setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');

  for (var r = 4; r <= 6; r++) {
    var row = sh.getRange(r, 1, 1, 7);
    row.setBackground(P.surfaceAlt).setFontColor(P.textPrimary).setFontSize(10);
    row.setBorder(false, false, true, false, false, false, P.border, SpreadsheetApp.BorderStyle.SOLID);
  }
  sh.getRange('A4:G4').setFontWeight('bold').setFontSize(12);

  sh.getRange('A7:G7').setBackground(P.surface);

  var notes = sh.getRange('A8:G17');
  notes.setBackground(P.surfaceAlt).setFontColor(P.textSecondary).setFontSize(9)
    .setVerticalAlignment('top').setWrap(true);
  notes.setBorder(true, true, true, true, false, false, P.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('A8').setFontWeight('bold').setFontColor(P.textPrimary);
}


/* ────────────── PAD PANEL (Current or Upcoming) ────────────── */
function stylePadPanel(sh, startCol, direction) {
  var w = 15;
  var isDown = (direction === 'down');
  var headerColor = isDown ? P.red : P.green;
  var headerLight = isDown ? P.redLight : P.greenLight;
  var headerLabel = isDown ? 'CURRENT PAD  \u00b7  RIG DOWN' : 'UPCOMING PAD  \u00b7  RIG UP';

  var h1 = sh.getRange(1, startCol, 1, w);
  h1.merge().setValue(headerLabel)
    .setBackground(P.dark).setFontColor(P.darkFg)
    .setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('center');

  var h2 = sh.getRange(2, startCol, 1, w);
  h2.merge().setBackground(headerColor).setValue('');

  var opRange = sh.getRange(3, startCol, 1, 5);
  opRange.setBackground(P.muted).setFontColor(P.textPrimary)
    .setFontSize(11).setFontWeight('bold').setHorizontalAlignment('center');
  var padRange = sh.getRange(3, startCol + 5, 1, 5);
  padRange.setBackground(P.muted).setFontColor(P.textPrimary)
    .setFontSize(11).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(3, startCol + 10, 1, 5).setBackground(P.muted);

  var dateRow = sh.getRange(4, startCol, 1, w);
  dateRow.setBackground(headerLight).setFontSize(10).setFontWeight('bold')
    .setHorizontalAlignment('center').setFontColor(headerColor);

  sh.getRange(5, startCol, 1, w).setBackground(P.surface).setFontSize(9).setFontColor(P.textSecondary);

  for (var r = 6; r <= 12; r++) {
    var bg = (r % 2 === 0) ? P.surfaceAlt : P.surface;
    sh.getRange(r, startCol, 1, w).setBackground(bg).setFontSize(10).setFontColor(P.textPrimary);
  }

  sh.getRange(13, startCol, 1, w).setBackground(P.mid).setFontColor(P.midFg)
    .setFontSize(9).setFontWeight('bold').setHorizontalAlignment('left');

  for (var r = 14; r <= 16; r++) {
    sh.getRange(r, startCol, 1, w).setBackground(P.surface).setFontSize(10).setFontColor(P.textPrimary)
      .setBorder(false, false, true, false, false, false, P.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(r, startCol + w - 1).setFontWeight('bold').setFontColor(P.accent).setFontSize(11)
      .setHorizontalAlignment('center');
  }

  sh.getRange(17, startCol, 1, w).setBackground(P.mid).setFontColor(P.midFg)
    .setFontSize(9).setFontWeight('bold').setHorizontalAlignment('left');

  sh.getRange(18, startCol, 1, w).setBackground(P.purpleBg).setFontColor(P.purple)
    .setFontWeight('bold').setFontSize(10);
  sh.getRange(19, startCol, 1, w).setBackground(P.purpleBg).setFontSize(9).setFontColor(P.textSecondary);

  for (var r = 20; r <= 21; r++) {
    sh.getRange(r, startCol, 1, w).setBackground(P.surface).setFontSize(10).setFontColor(P.textPrimary)
      .setBorder(false, false, true, false, false, false, P.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(r, startCol + w - 1).setFontWeight('bold').setFontColor(P.purple).setFontSize(11)
      .setHorizontalAlignment('center');
  }

  sh.getRange(22, startCol, 1, w).setBackground(P.mid).setFontColor(P.midFg)
    .setFontSize(9).setFontWeight('bold').setHorizontalAlignment('left');

  sh.getRange(23, startCol, 1, w).setBackground(P.orangeBg).setFontColor(P.orange)
    .setFontWeight('bold').setFontSize(10);
  sh.getRange(24, startCol, 1, w).setBackground(P.orangeBg).setFontSize(9).setFontColor(P.textSecondary);

  sh.getRange(1, startCol, 24, w)
    .setBorder(true, true, true, true, false, false, P.borderDark, SpreadsheetApp.BorderStyle.SOLID);
}


/* ──────── UPCOMING PAD CHECKLIST ──────── */
function styleChecklist(sh) {
  for (var r = 26; r <= 30; r++) {
    sh.getRange(r, 26, 1, 15).setBackground(P.surface).setFontSize(10).setFontColor(P.textPrimary)
      .setBorder(false, false, true, false, false, false, P.border, SpreadsheetApp.BorderStyle.SOLID);
  }
}


/* ────────────── CALENDAR TAB RESTYLE ────────────── */
function restyleCalendarTab(cal) {
  if (!cal) return;
  var title = cal.getRange(1, 1, 1, 8);
  title.setBackground(P.dark).setFontColor(P.darkFg);
  var sub = cal.getRange(2, 1, 1, 8);
  sub.setBackground('#334155').setFontColor('#CBD5E1');
  cal.setTabColor(P.accent);
}

/* ---------- DATE / TIME FORMAT ENFORCEMENT ---------- */

var FLEET_DATE_CELLS = [
  'M4','T4',           // RIG DOWN / RD LOGISTICS
  'N19','T19',         // Current ROUSTABOUT start/end
  'M24','T24',         // Current MISC start/end
  'AC4','AJ4',         // RIG UP / RU LOGISTICS
  'AD19','AJ19',       // Upcoming ROUSTABOUT start/end
  'AC24','AJ24'        // Upcoming MISC start/end
];

var FLEET_DT_FORMAT = 'M/d/yyyy h:mm:ss';

function applyFleetDateFormats(sh) {
  if (!sh) return;
  FLEET_DATE_CELLS.forEach(function(a1) {
    var rng = sh.getRange(a1);
    rng.setNumberFormat(FLEET_DT_FORMAT);
  });
}

function applyAllFleetDateFormats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fleetNames = [
    'FLEET # 1 - Brandon Dock',
    'FLEET # 6 - Lindsay Riley',
    'FLEET # 9 - Eric Cronberg',
    'FLEET # 10 - Corey Miller',
    'FLEET # 13 - Ryan Poorman',
    'FLEET # 15 - Jeremy Bower'
  ];
  fleetNames.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    applyFleetDateFormats(sh);
  });
  Logger.log('Applied date formats to all fleet tabs.');
}

/**
 * Installable onEdit trigger: whenever a user edits any of the date cells
 * on a fleet tab, reapply the standard M/d/yyyy h:mm:ss format so the
 * display is consistent regardless of how they typed the value.
 */
function onEditFleetDateFormat(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    var name = sh.getName();
    if (name.indexOf('FLEET #') !== 0) return;
    var editedA1 = e.range.getA1Notation();
    // Reapply format on the full known list (cheap) - also ensures any newly
    // typed date in adjacent merged cells snaps to the right format.
    applyFleetDateFormats(sh);
  } catch (err) {
    Logger.log('onEditFleetDateFormat error: ' + err);
  }
}

function installFleetDateFormatTrigger() {
  var ss = SpreadsheetApp.getActive();
  // Remove any existing triggers for this handler to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'onEditFleetDateFormat') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onEditFleetDateFormat')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  Logger.log('Installed onEdit trigger for fleet date formatting.');
}
