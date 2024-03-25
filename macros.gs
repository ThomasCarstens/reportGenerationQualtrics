function testMove() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('E6').activate();
  var sheet = spreadsheet.getActiveSheet();
  var charts = sheet.getCharts();
  var chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .setChartType(Charts.ChartType.RADAR)
  .addRange(spreadsheet.getRange('A1:B15'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_ROWS)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', true)
  .setOption('subtitle', 'VL Sustainability/Circular')
  .setOption('title', 'Startup Radar Graph')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#1a1a1a')
  .setOption('subtitleTextStyle.color', '#999999')
  .setOption('titleTextStyle.color', '#757575')
  .setOption('annotations.total.textStyle.color', '#808080')
  .setOption('height', 556)
  .setOption('width', 900)
  .setPosition(1, 5, 47, 6)
  .build();
  sheet.insertChart(chart);
};

function formatGraph15_02(startup, ventureLab, date) {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C3').activate();
  var sheet = spreadsheet.getActiveSheet();
  var charts = sheet.getCharts();
  var chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .setChartType(Charts.ChartType.RADAR)
  .addRange(spreadsheet.getRange('A1:B15'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_ROWS)
  .setOption('useFirstColumnAsDomain', true)
  .setPosition(10, 6, 14, 13)
  .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .setChartType(Charts.ChartType.RADAR)
  .addRange(spreadsheet.getRange('A1:B15'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_ROWS)
  .setOption('useFirstColumnAsDomain', true)
  .setOption('height', 504)
  .setOption('width', 816)
  .setPosition(4, 6, 14, 5)
  .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .setChartType(Charts.ChartType.RADAR)
  .addRange(spreadsheet.getRange('A1:B15'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_ROWS)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', true)
  .setOption('subtitle', 'AMunich')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#1a1a1a')
  .setOption('subtitleTextStyle.color', '#999999')
  .setOption('annotations.total.textStyle.color', '#808080')
  .setOption('height', 504)
  .setOption('width', 816)
  .setPosition(4, 6, 14, 5)
  .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .setChartType(Charts.ChartType.RADAR)
  .addRange(spreadsheet.getRange('A1:B15'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_ROWS)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', true)
  .setOption('subtitle', 'AMunich on ')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#1a1a1a')
  .setOption('subtitleTextStyle.color', '#999999')
  .setOption('annotations.total.textStyle.color', '#808080')
  .setOption('height', 504)
  .setOption('width', 816)
  .setPosition(4, 6, 14, 5)
  .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .setChartType(Charts.ChartType.RADAR)
  .addRange(spreadsheet.getRange('A1:B15'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_ROWS)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', true)
  .setOption('subtitle', 'AMunich on 23.02.202')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#1a1a1a')
  .setOption('subtitleTextStyle.color', '#999999')
  .setOption('annotations.total.textStyle.color', '#808080')
  .setOption('height', 504)
  .setOption('width', 816)
  .setPosition(4, 6, 14, 5)
  .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .setChartType(Charts.ChartType.RADAR)
  .addRange(spreadsheet.getRange('A1:B15'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_ROWS)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', true)
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#1a1a1a')
  .setOption('subtitleTextStyle.color', '#999999')
  .setOption('titleTextStyle.color', '#757575')
  .setOption('annotations.total.textStyle.color', '#808080')
  .setOption('height', 504)
  .setOption('width', 816)
  .setPosition(4, 6, 14, 5)
  .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .setChartType(Charts.ChartType.RADAR)
  .addRange(spreadsheet.getRange('A1:B15'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_ROWS)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', true)
  .setOption('subtitle', ventureLab)
  .setOption('title', startup+ ' on '+ date)
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#1a1a1a')
  .setOption('subtitleTextStyle.color', '#999999')
  .setOption('titleTextStyle.color', '#757575')
  .setOption('annotations.total.textStyle.color', '#808080')
  .setOption('height', 504)
  .setOption('width', 816)
  .setPosition(4, 6, 14, 5)
  .build();
  sheet.insertChart(chart);
  spreadsheet.getActiveSheet().setHiddenGridlines(true);
};

function formatGraph23_02Design(startup, ventureLab, sheetDate, date) {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  var sheet = spreadsheet.getActiveSheet();
  var charts = sheet.getCharts();
  var chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .setChartType(Charts.ChartType.RADAR)
  .addRange(spreadsheet.getRange('A2:C9'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_ROWS)
  .setOption('useFirstColumnAsDomain', true)
  .setPosition(16, 14, 30, 12)
  .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[0];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .setChartType(Charts.ChartType.RADAR)
  .addRange(spreadsheet.getRange('A2:B9'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_ROWS)
  .setOption('useFirstColumnAsDomain', true)
  .setPosition(16, 7, 70, 12)
  .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[0];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .setChartType(Charts.ChartType.RADAR)
  .addRange(spreadsheet.getRange('A2:B9'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_ROWS)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', true)
  .setOption('title', 'Maturity')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#1a1a1a')
  .setOption('titleTextStyle.color', '#757575')
  .setOption('annotations.total.textStyle.color', '#808080')
  .setPosition(16, 7, 70, 12)
  .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[0];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .setChartType(Charts.ChartType.RADAR)
  .addRange(spreadsheet.getRange('A2:B9'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_ROWS)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', true)
  .setOption('title', 'Maturity')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#1a1a1a')
  .setOption('titleTextStyle.color', '#757575')
  .setOption('annotations.total.textStyle.color', '#808080')
  .setOption('vAxes.0.viewWindow.min', 0)
  .setPosition(16, 7, 70, 12)
  .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[0];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .setChartType(Charts.ChartType.RADAR)
  .addRange(spreadsheet.getRange('A2:B9'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_ROWS)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', true)
  .setOption('title', 'Maturity')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#1a1a1a')
  .setOption('titleTextStyle.color', '#757575')
  .setOption('annotations.total.textStyle.color', '#808080')
  .setOption('vAxes.0.viewWindow.max', 7)
  .setOption('vAxes.0.viewWindow.min', 0)
  .setPosition(16, 7, 70, 12)
  .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .setChartType(Charts.ChartType.RADAR)
  .addRange(spreadsheet.getRange('A2:C9'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_ROWS)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', true)
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#1a1a1a')
  .setOption('titleTextStyle.color', '#757575')
  .setOption('annotations.total.textStyle.color', '#808080')
  .setOption('vAxes.0.viewWindow.min', 0)
  .setPosition(16, 14, 30, 12)
  .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .setChartType(Charts.ChartType.RADAR)
  .addRange(spreadsheet.getRange('A2:C9'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_ROWS)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', true)
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#1a1a1a')
  .setOption('titleTextStyle.color', '#757575')
  .setOption('annotations.total.textStyle.color', '#808080')
  .setOption('vAxes.0.viewWindow.max', 5)
  .setOption('vAxes.0.viewWindow.min', 0)
  .setPosition(16, 14, 30, 12)
  .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .setChartType(Charts.ChartType.RADAR)
  .addRange(spreadsheet.getRange('A2:C9'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_ROWS)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', true)
  .setOption('title', 'Potential')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#1a1a1a')
  .setOption('titleTextStyle.color', '#757575')
  .setOption('annotations.total.textStyle.color', '#808080')
  .setOption('vAxes.0.viewWindow.max', 5)
  .setOption('vAxes.0.viewWindow.min', 0)
  .setPosition(16, 14, 30, 12)
  .build();
  sheet.insertChart(chart);
  spreadsheet.getActiveSheet().setHiddenGridlines(true);



  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Template'), true);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(sheetDate), true);



  spreadsheet.getRange('D1').activate();
  spreadsheet.getRange('\'Template\'!A1:R12').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('E3:F9').activate();
  spreadsheet.getRange('B3:C9').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('A:C').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('C1'));
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
  spreadsheet.getRange('D:D').activate();
  spreadsheet.getRange('36:1000').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('D36'));

  var sheet = spreadsheet.getActiveSheet();
  var activeRange = spreadsheet.getActiveRange();
  var startRow = activeRange.getRow();
  var numRows = activeRange.getNumRows();

  // Ensure we don't try to delete more rows than exist from the start position to the end of the sheet
  var lastRow = sheet.getLastRow();
  if (startRow + numRows - 1 > lastRow) {
    numRows = lastRow - startRow + 1;
  }

  // Only attempt to delete if there are rows to delete
  if (numRows > 0) {
    sheet.deleteRows(startRow, numRows);
  } else {
    // Log an error or handle the case where there are no rows to delete
    console.log("No rows to delete or range out of bounds");
  }


  // spreadsheet.getActiveSheet().deleteRows(startRow, numRows);
  spreadsheet.getRange('X5').activate();
};

function removeDataFromPotentialGraph() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C15').activate();
  var sheet = spreadsheet.getActiveSheet();
  var charts = sheet.getCharts();
  var chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .setChartType(Charts.ChartType.RADAR)
  .addRange(spreadsheet.getRange('A2:C9'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_ROWS)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', true)
  .setOption('title', 'Potential Level for undefined')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#1a1a1a')
  .setOption('titleTextStyle.color', '#757575')
  .setOption('annotations.total.textStyle.color', '#808080')
  .setPosition(12, 12, 0, 0)
  .build();
  sheet.insertChart(chart);
};
