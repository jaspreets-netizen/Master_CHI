// ════════════════════════════════════════════════════════
// DATA READER
// Reads the "Complete CHI Data" sheet and trend sheets,
// returns structured objects consumed by cu-sync.gs.
// ════════════════════════════════════════════════════════

/**
 * Read a trend tab and return { siteName_lowercase → value } for the last
 * COMPLETE month — defined as the last column where ≥50 % of sites have a
 * value ≥ 1.  Falls back to the last column with any data if no complete
 * month is found.
 */
function readTrendSheetComplete_(sheetName) {
  var MIN_COVERAGE = 0.5;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(sheetName);
  if (!sh) { Logger.log('Sheet not found: ' + sheetName); return { map: {}, monthLabel: '' }; }
  var data = sh.getDataRange().getValues();
  if (data.length < 3) return { map: {}, monthLabel: '' };

  // Count site rows (rows 2+ with a non-empty name in col 0)
  var siteCount = 0;
  for (var r = 2; r < data.length; r++) { if (String(data[r][0]).trim()) siteCount++; }
  if (siteCount === 0) return { map: {}, monthLabel: '' };

  // Scan right to left; pick the first column where ≥ MIN_COVERAGE sites have val ≥ 1
  var bestCol = -1, bestLabel = '';
  for (var c = data[1].length - 1; c >= 2; c--) {
    var goodCount = 0;
    for (var r = 2; r < data.length; r++) {
      if (!String(data[r][0]).trim()) continue;
      var v = parseFloat(data[r][c]);
      if (!isNaN(v) && v >= 1) goodCount++;
    }
    if (goodCount / siteCount >= MIN_COVERAGE) {
      bestCol = c; bestLabel = String(data[1][c]); break;
    }
  }

  // Fallback: last non-empty column
  if (bestCol < 0) {
    Logger.log(sheetName + ': no complete month found, using latest with any data');
    for (var c = data[1].length - 1; c >= 2; c--) {
      for (var r = 2; r < data.length; r++) {
        var v = parseFloat(data[r][c]);
        if (!isNaN(v) && v > 0) { bestCol = c; bestLabel = String(data[1][c]); break; }
      }
      if (bestCol >= 0) break;
    }
  }

  Logger.log(sheetName + ' → using column ' + bestCol + ' (' + bestLabel + ')');
  var result = {};
  for (var r = 2; r < data.length; r++) {
    var name = String(data[r][0]).trim();
    if (!name) continue;
    var v = parseFloat(data[r][bestCol]);
    if (!isNaN(v) && v > 0) result[name.toLowerCase()] = v;
  }
  return { map: result, monthLabel: bestLabel };
}

/**
 * Return { siteName_lowercase → cemName } from the Activation sheet.
 */
function readCemNames_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Activation');
  if (!sh) return {};
  var data = sh.getRange(3, 1, 100, 3).getValues();
  var result = {};
  for (var i = 0; i < data.length; i++) {
    var name = String(data[i][1]).trim();
    var cem  = String(data[i][2]).trim();
    if (name) result[name.toLowerCase()] = cem;
  }
  return result;
}
