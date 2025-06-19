function onEdit(e) {
  if (e.range.getSheet().getName() === "Keyword Research") {
    refreshBorders();
  }
}

function refreshBorders() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Keyword Research");
  if (!sheet) return;
  
  var maxRows = sheet.getMaxRows();
  var maxCols = sheet.getMaxColumns();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;  // only header
  
  // 1) Clear every border on the sheet
  sheet.getRange(1, 1, maxRows, maxCols)
       .setBorder(false, false, false, false, false, false);
  
  // 2) Read D2:D(lastRow) into a simple array
  var colD = sheet
    .getRange(2, 4, lastRow - 1, 1)
    .getValues()
    .map(function(r){ return r[0]; });
  
  // 3) Prepare three lists of A1-notated ranges
  var both      = [];
  var topOnly   = [];
  var bottomOnly= [];
  var endCol    = columnToLetter(maxCols);
  
  colD.forEach(function(curr, i) {
    var prev = (i > 0             ? colD[i - 1]       : null);
    var next = (i < colD.length-1 ? colD[i + 1]       : null);
    var row  = i + 2;  // because colD[0] is sheet row 2
    var a1   = "A" + row + ":" + endCol + row;
    
    var isStart = (curr !== prev);
    var isEnd   = (curr !== next);
    
    if (isStart && isEnd) {
      both.push(a1);
    } else if (isStart) {
      topOnly.push(a1);
    } else if (isEnd) {
      bottomOnly.push(a1);
    }
  });
  
  // 4) Apply borders in three batch calls
  if (both.length) {
    sheet.getRangeList(both)
         .setBorder(true, false, true, false, false, false,
                    "black", SpreadsheetApp.BorderStyle.SOLID);
  }
  if (topOnly.length) {
    sheet.getRangeList(topOnly)
         .setBorder(true, false, false, false, false, false,
                    "black", SpreadsheetApp.BorderStyle.SOLID);
  }
  if (bottomOnly.length) {
    sheet.getRangeList(bottomOnly)
         .setBorder(false, false, true, false, false, false,
                    "black", SpreadsheetApp.BorderStyle.SOLID);
  }
}

// Helper: 1 → A, 27 → AA, etc.
function columnToLetter(col) {
  var s = "";
  while (col > 0) {
    var m = (col - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    col = Math.floor((col - 1) / 26);
  }
  return s;
}
