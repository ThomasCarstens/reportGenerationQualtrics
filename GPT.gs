
function setSheetName(milliseconds, startupName, apiTriggered) {
    // Create a new sheet for the radar graph
  var sheetId = new Date(milliseconds);

  if (apiTriggered == 1){
    /* API Code. */
    sheetName = startupName + " | "+sheetId.toUTCString();
  } else {
    /* Test purposes (without API) */
    sheetName = "Valorio | "+sheetId.toUTCString();
  }
  return sheetName

}

function generateRadarGraph(sessionID, startup, apiTriggered) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ALL") //DEPENDENCY : Name of Sheet to get Data from
  var dataRange = sheet.getDataRange();
  var lastRow = dataRange.getLastRow();
  var lastColumn = dataRange.getLastColumn();
  var data = dataRange.getValues();
  
  var startupNameColumnIndex = getColumnIndexByName('Token');          
  
  // GROUP ROWS with the specified token (sessionID)
  var startupRows = [];
  for (var i = 1; i < lastRow; i++) {
    console.log(data[i][startupNameColumnIndex - 1])                             
    if (data[i][startupNameColumnIndex - 1] == sessionID) {
      startupRows.push(i);
      console.log('Found a Row for '+sessionID)
    }
  }
  
  // AVERAGE COLUMNS based on rows with the specified token
  var averages = [];
  for (var j = 6; j <= 19; j++) {

    var sum = 0;
    var totalCoachResponses= startupRows.length
    for (var k = 0; k < startupRows.length; k++) {
      cellValue = data[startupRows[k]][j - 1]   //one column taking only the grouped rows
      sum += parseFloat(cellValue)              // sum up and move to next in column 
    }

    averages.push(sum /  startupRows.length);
    /* txa: This works but it is problematic because it assumes that no answer is a 0, but if the user mistakingly puts nothing, the resulting 0 skews our average... */
  }

  // Create a new sheet for the radar graph
  milliseconds = Date.now()

  var source = SpreadsheetApp.getActiveSpreadsheet();
  var template = source.getSheetByName('18.03 Template'); //DEPENDENCY : Name of Sheet to copy from
  var sheetName = setSheetName(milliseconds, startup, apiTriggered)
  var radarSheet = template.copyTo(source).setName(sheetName);

  ///////////////////////////////
  // FOR SUCH ROWS: Put values into the final sheet 
  var coach_values = [];
  
  var totalCoachResponses= startupRows.length
  for (var k = 0; k < startupRows.length; k++) {
    coach_values = []
    for (var j = 6; j <= 26; j++) {
      cellValue = data[startupRows[k]][j - 1]   //one column taking only the grouped rows
      // Prepare data for radar graph
      coach_values.push(cellValue);
    }


      var chartData = [
        [coach_values[0], coach_values[7],  'Comment: ', coach_values[14]],
        [coach_values[1], coach_values[8],  'Comment: ', coach_values[15]],
        [coach_values[2], coach_values[9], 'Comment: ', coach_values[16]],
        [coach_values[3], coach_values[10], 'Comment: ', coach_values[17]],
        [coach_values[4], coach_values[11],'Comment: ',coach_values[18]],
        [coach_values[5], coach_values[12],'Comment: ',coach_values[19]],
        [coach_values[6], coach_values[13],'Comment: ',coach_values[20]],
      ];
      var startRow = 19 + (k * 10); // D3 for the first row, G3 for the second, etc.
      var coach_range = radarSheet.getRange(startRow, 1, 7, 4); // 7 rows, 3 columns
      coach_range.setValues(chartData);
      
      /* this will fill all the coaches and overflow if there are more than 5 coaches - unlikely scenario.*/
  }
  radarSheet.autoResizeColumn(4);



  
  // Prepare data for radar graph
  var chartData = [
    ['Category',              'Maturity', 'Potential'],
    ['Customer',              averages[0], averages[7]],
    ['Technology',            averages[1], averages[8]],
    ['Business',              averages[2], averages[9]],
    ['Team',                  averages[3], averages[10]],
    ['Funding',               averages[4], averages[11]],
    ['Intellecutal Property', averages[5], averages[12]],
    ['Sustainability',        averages[6], averages[13]],

  ];
  
  // Add chart data to sheet
  radarSheet.getRange('A5:C12').setValues(chartData);
  
  // Insert radar chart
  var chart = radarSheet.newChart()
      .setChartType(Charts.ChartType.RADAR)
      .setOption('title', 'Maturity Level for ' +startup) // could add date , or .setOption('subtitle', ventureLab)
      .addRange(radarSheet.getRange('A2:B9'))
      .setPosition(15, 10, 0, 0)
      .build();

  
  radarSheet.insertChart(chart);

  // graph2range = getNonAdjascentRange(radarSheet)
  
  // Insert radar chart
  var chart = radarSheet.newChart()
      .setChartType(Charts.ChartType.RADAR)
      .setOption('title', 'Potential Level for ' +startup) // could add date , or .setOption('subtitle', ventureLab)
      .addRange(radarSheet.getRange('A2:C9'))     
      // .removeRange(radarSheet.getRange('B2:B9')) // this is not working, maybe the graph is updating too fast...            //URGENT
      .setPosition(40, 10, 0, 0)
      .build();
  radarSheet.insertChart(chart);
  // removeDataFromPotentialGraph() 

  return sheetName
}

function getColumnIndexByName(name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ALL")
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] == name) {
      return i + 1;
    }
  }
  return -1; // Return -1 if column not found
}

function getNonAdjascentRange(radarSheet) {
  const values3D = radarSheet
    .getRangeList(['A5:A12', 'C5:C12']) // DEPENDENCY: Range for Graph 2: Maturity graph
    .getRanges()
    .map(range => range.getValues());
  const valuesJoin = [];
  values3D.forEach((columnValues, columnIndex) => {
    columnValues.forEach((rowValues, rowIndex) => {
      valuesJoin[rowIndex][columnIndex] = rowValues[0];
    });
  });
  console.log(valuesJoin);
  return valuesJoin
}

