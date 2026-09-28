/**
 * LayoutCopier.gs — Copy Fleet #1 page layout to all other fleet tabs.
 * Copies: merged cells, column widths, row heights.
 * Run copyFleet1Layout() to apply.
 */

function copyFleet1Layout() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var source = ss.getSheetByName('FLEET # 1 - Brandon Dock');
  if (!source) { Logger.log('ERROR: Fleet #1 not found'); return; }

  var targets = [
    'FLEET # 6 - Lindsay Riley',
    'FLEET # 9 - Eric Cronberg',
    'FLEET # 10 - Corey Miller',
    'FLEET # 13 - Ryan Poorman',
    'FLEET # 15 - Jeremy Bower'
  ];

  // 1. Read Fleet #1 layout info
  var maxRow = source.getMaxRows();
  var maxCol = source.getMaxColumns();
  Logger.log('Source: ' + maxRow + ' rows x ' + maxCol + ' cols');

  // Get merged ranges from source
  var srcMerges = source.getRange(1, 1, maxRow, maxCol).getMergedRanges();
  var mergeSpecs = [];
  for (var i = 0; i < srcMerges.length; i++) {
    var m = srcMerges[i];
    mergeSpecs.push({
      row: m.getRow(),
      col: m.getColumn(),
      numRows: m.getNumRows(),
      numCols: m.getNumColumns()
    });
  }
  Logger.log('Found ' + mergeSpecs.length + ' merged ranges in Fleet #1');

  // Get column widths from source
  var colWidths = [];
  for (var c = 1; c <= maxCol; c++) {
    colWidths.push(source.getColumnWidth(c));
  }

  // Get row heights from source
  var rowHeights = [];
  for (var r = 1; r <= maxRow; r++) {
    rowHeights.push(source.getRowHeight(r));
  }

  // 2. Apply to each target
  targets.forEach(function(name) {
    var tgt = ss.getSheetByName(name);
    if (!tgt) { Logger.log('SKIP: ' + name); return; }
    Logger.log('Copying layout to: ' + name);

    var tgtMaxRow = tgt.getMaxRows();
    var tgtMaxCol = tgt.getMaxColumns();

    // Ensure target has enough rows and columns
    if (tgtMaxRow < maxRow) {
      tgt.insertRowsAfter(tgtMaxRow, maxRow - tgtMaxRow);
      tgtMaxRow = maxRow;
    }
    if (tgtMaxCol < maxCol) {
      tgt.insertColumnsAfter(tgtMaxCol, maxCol - tgtMaxCol);
      tgtMaxCol = maxCol;
    }

    // Break apart all existing merges in target
    var existingMerges = tgt.getRange(1, 1, tgtMaxRow, tgtMaxCol).getMergedRanges();
    existingMerges.forEach(function(em) { em.breakApart(); });
    Logger.log('  Broke apart ' + existingMerges.length + ' existing merges');

    // Set column widths
    for (var c = 1; c <= maxCol; c++) {
      tgt.setColumnWidth(c, colWidths[c - 1]);
    }

    // Set row heights (up to source maxRow)
    for (var r = 1; r <= maxRow; r++) {
      tgt.setRowHeight(r, rowHeights[r - 1]);
    }

    // Apply merged ranges
    for (var i = 0; i < mergeSpecs.length; i++) {
      var ms = mergeSpecs[i];
      try {
        tgt.getRange(ms.row, ms.col, ms.numRows, ms.numCols).merge();
      } catch (e) {
        Logger.log('  Merge error at R' + ms.row + 'C' + ms.col + ': ' + e.message);
      }
    }
    Logger.log('  Applied ' + mergeSpecs.length + ' merges');
  });

  Logger.log('Done — layout copied from Fleet #1 to all other fleet tabs.');
}