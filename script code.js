function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Sheet1") || ss.getSheets()[0];
    var data = JSON.parse(e.postData.contents);
    
    // Standardize incoming parameters from the ESP32 payload
    var operation = (data.operation || "push").toLowerCase().trim();
    var rawWeight = parseFloat(data.final_weight) || 0;
    var durationSec = parseFloat(data.duration) || 0; 
    
    var lastRow = sheet.getLastRow();
    
    // Initialize blank spreadsheet cleanly if required
    if (lastRow === 0) {
      sheet.appendRow(["Batch 1", "Weight Data", "Batch Duration"]);
      lastRow = 1;
    }
    
    // --- MATHEMATICAL BATCH COUNTING ---
    var allData = sheet.getRange(1, 1, lastRow, 1).getValues();
    var summaryCount = 0;
    for (var i = 0; i < allData.length; i++) {
      if (String(allData[i][0]).indexOf("Summary") !== -1) {
        summaryCount++;
      }
    }
    var currentBatchNum = summaryCount + 1;

    // --- OPERATION A: INGEST INDIVIDUAL WEIGHT DATA ---
    if (operation === "push") {
      sheet.appendRow(["Batch " + currentBatchNum, rawWeight, ""]);
      return ContentService.createTextOutput("Logged entry to Batch " + currentBatchNum);
    } 
    
    // --- OPERATION B: TERMINATE AND SUMMARIZE BATCH TIME ---
    else if (operation === "stop") {
      var lastCellText = String(allData[lastRow - 1][0]);
      
      // Debounce protection blocks double summaries if button bounces
      if (lastCellText.indexOf("Summary") === -1) {
        var mins = Math.floor(durationSec / 60);
        var secs = Math.floor(durationSec % 60);
        var durationString = mins + "m " + secs + "s";
        
        // 1. Commit batch summary row
        sheet.appendRow(["Batch " + currentBatchNum + " Summary", "Done", durationString]);
        
        // 2. Append blank spacer break line
        sheet.appendRow([]); 
        
        // 3. Mount upcoming header tag markers for subsequent operations
        var nextBatchNum = currentBatchNum + 1;
        sheet.appendRow(["Batch " + nextBatchNum, "Weight Data", "Batch Duration"]);
        
        return ContentService.createTextOutput("Closed Batch " + currentBatchNum + " processing clock.");
      }
      return ContentService.createTextOutput("System already stopped.");
    }
    
  } catch(err) {
    return ContentService.createTextOutput("Backend Error: " + err.message);
  }
}

function doGet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var data = sheet.getDataRange().getValues();
  var jsonOutput = [];
  
  for (var i = 0; i < data.length; i++) {
    var colA = String(data[i][0]).trim();
    var colB = String(data[i][1]).trim();
    var colC = String(data[i][2]).trim();
    
    if (!colA) continue;
    
    // Isolate batch summary structural properties
    if (colA.indexOf("Summary") !== -1) {
      jsonOutput.push({
        type: "summary",
        batch: colA.split("Summary")[0].trim(),
        runtime: colC
      });
      continue;
    }
    
    if (colB === "Weight Data") continue;
    
    // Maps standalone weight instances sequentially
    var weightValue = parseFloat(colB);
    if (colA.indexOf("Batch") === 0 && !isNaN(weightValue)) {
      jsonOutput.push({ 
        type: "data",
        time: new Date().toLocaleDateString() + " (Logged)",
        val: weightValue,
        batch: colA
      });
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify(jsonOutput))
                       .setMimeType(ContentService.MimeType.JSON);
}