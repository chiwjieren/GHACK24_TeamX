// Function to update the Income summary and pie chart 
function updateIncome() {
  var sheetId = '1B1yXKBZdzFUs3R55SUYDNfTrW0f7Er0bTQgh0N_Nxeo';
  var spreadsheet = SpreadsheetApp.openById(sheetId);
  var incomeSheet = spreadsheet.getSheetByName('Income');
  var incomeSummarySheet = spreadsheet.getSheetByName('Income Summary');
  
  if (!incomeSheet || !incomeSummarySheet) {
    Logger.log('Income or Income Summary sheet not found.');
    return;
  }
  
  var incomeData = incomeSheet.getDataRange().getValues();
  
  if (incomeData.length <= 1) {
    Logger.log('Income sheet is empty or has insufficient rows.');
    return;
  }
  
  var incomeSummary = summarizeDataSheet(incomeData, 'Income');
  var totalIncome = calculateTotal(incomeSummary);
  
  createOrUpdateSummarySheet(incomeSummarySheet, incomeSummary, 'Income Summary');
  createOrUpdatePieChart(incomeSummarySheet, 'Income Pie Chart', totalIncome);
}
// Function to update the Expenses summary and pie chart
function updateExpenses() {
  var sheetId = '1B1yXKBZdzFUs3R55SUYDNfTrW0f7Er0bTQgh0N_Nxeo';
  var spreadsheet = SpreadsheetApp.openById(sheetId);
  var expensesSheet = spreadsheet.getSheetByName('Expenses');
  var expensesSummarySheet = spreadsheet.getSheetByName('Expenses Summary');
  
  if (!expensesSheet || !expensesSummarySheet) {
    Logger.log('Expenses or Expenses Summary sheet not found.');
    return;
  }
  
  var expensesData = expensesSheet.getDataRange().getValues();
  
  if (expensesData.length <= 1) {
    Logger.log('Expenses sheet is empty or has insufficient rows.');
    return;
  }
  
  var expensesSummary = summarizeDataSheet(expensesData, 'Expenses');
  var totalExpenses = calculateTotal(expensesSummary);
  
  createOrUpdateSummarySheet(expensesSummarySheet, expensesSummary, 'Expenses Summary');
  createOrUpdatePieChart(expensesSummarySheet, 'Expenses Pie Chart', totalExpenses);
}
// Function to update the Daily Summary sheet
function updateDailySummary() {
  var sheetId = '1B1yXKBZdzFUs3R55SUYDNfTrW0f7Er0bTQgh0N_Nxeo';
  var spreadsheet = SpreadsheetApp.openById(sheetId);
  var incomeSummarySheet = spreadsheet.getSheetByName('Income Summary');
  var expensesSummarySheet = spreadsheet.getSheetByName('Expenses Summary');
  var dailySummarySheet = spreadsheet.getSheetByName('Daily Summary');
  
  if (!incomeSummarySheet || !expensesSummarySheet || !dailySummarySheet) {
    Logger.log('Income Summary, Expenses Summary, or Daily Summary sheet not found.');
    return;
  }
  
  var incomeSummaryData = incomeSummarySheet.getDataRange().getValues();
  var expensesSummaryData = expensesSummarySheet.getDataRange().getValues();
  
  var totalIncome = calculateTotal(incomeSummaryData.slice(1));
  var totalExpenses = calculateTotal(expensesSummaryData.slice(1));
  var profitLoss = totalIncome - totalExpenses;
  
  dailySummarySheet.clear(); // Clear existing data
  
  dailySummarySheet.appendRow(['Metric', 'Amount']);
  dailySummarySheet.appendRow(['Total Income', totalIncome]);
  dailySummarySheet.appendRow(['Total Expenses', totalExpenses]);
  var profitLossRow = ['Profit/Loss', profitLoss];
  dailySummarySheet.appendRow(profitLossRow);
  
  // Format the Profit/Loss cell
  var profitLossCell = dailySummarySheet.getRange('B4');
  if (profitLoss > 0) {
    profitLossCell.setBackground('green');
  } else {
    profitLossCell.setBackground('red');
  }
  // Add title and color code
  var titleCell = dailySummarySheet.getRange('A1');
  titleCell.setValue('Daily Financial Summary').setFontSize(14).setFontWeight('bold');
  dailySummarySheet.getRange('A1:B1').setBackground('#f0f0f0'); // Light gray background for header
  createOrUpdateColumnChart(dailySummarySheet, 'Daily Profit/Loss');
}
// Function to update the Monthly Summary sheet
function updateMonthlySummary() {
  var sheetId = '1B1yXKBZdzFUs3R55SUYDNfTrW0f7Er0bTQgh0N_Nxeo';
  var spreadsheet = SpreadsheetApp.openById(sheetId);
  var incomeSummarySheet = spreadsheet.getSheetByName('Income Summary');
  var expensesSummarySheet = spreadsheet.getSheetByName('Expenses Summary');
  var monthlySummarySheet = spreadsheet.getSheetByName('Monthly Summary');
  
  if (!incomeSummarySheet || !expensesSummarySheet || !monthlySummarySheet) {
    Logger.log('Income Summary, Expenses Summary, or Monthly Summary sheet not found.');
    return;
  }
  
  var incomeSummaryData = incomeSummarySheet.getDataRange().getValues();
  var expensesSummaryData = expensesSummarySheet.getDataRange().getValues();
  
  var totalIncome = calculateTotal(incomeSummaryData.slice(1));
  var totalExpenses = calculateTotal(expensesSummaryData.slice(1));
  var profitLoss = totalIncome - totalExpenses;
  
  monthlySummarySheet.clear(); // Clear existing data
  
  monthlySummarySheet.appendRow(['Metric', 'Amount']);
  monthlySummarySheet.appendRow(['Total Income', totalIncome]);
  monthlySummarySheet.appendRow(['Total Expenses', totalExpenses]);
  var profitLossRow = ['Profit/Loss', profitLoss];
  monthlySummarySheet.appendRow(profitLossRow);
  
  // Format the Profit/Loss cell
  var profitLossCell = monthlySummarySheet.getRange('B4');
  if (profitLoss > 0) {
    profitLossCell.setBackground('green');
  } else {
    profitLossCell.setBackground('red');
  }
  // Add title and color code
  var titleCell = monthlySummarySheet.getRange('A1');
  titleCell.setValue('Monthly Financial Summary').setFontSize(14).setFontWeight('bold');
  monthlySummarySheet.getRange('A1:B1').setBackground('#d9ead3'); // Light green background for header
  createOrUpdateColumnChart(monthlySummarySheet, 'Monthly Profit/Loss');
}
// Helper function to summarize data from a sheet
function summarizeDataSheet(data, type) {
  var summary = {};
  var headers = data[0].map(header => header.toString().trim().toUpperCase());
  var categoryIndex = headers.indexOf('CATEGORY');
  var amountIndex = headers.indexOf('AMOUNT');
  
  if (categoryIndex === -1 || amountIndex === -1) {
    Logger.log('CATEGORY or AMOUNT column not found.');
    return [];
  }
  
  for (var i = 1; i < data.length; i++) {
    var category = data[i][categoryIndex].toString().toUpperCase();
    var amount = parseFloat(data[i][amountIndex]);
    
    if (!summary[category]) {
      summary[category] = 0;
    }
    if (!isNaN(amount)) {
      summary[category] += amount;
    }
  }
  
  var summaryArray = [];
  for (var category in summary) {
    summaryArray.push([category, summary[category]]);
  }
  
  return summaryArray;
}
// Helper function to calculate the total amount
function calculateTotal(summaryData) {
  return summaryData.reduce((total, row) => total + row[1], 0);
}
// Helper function to create or update the summary sheet
function createOrUpdateSummarySheet(sheet, summaryData, title) {
  sheet.clear(); // Clear existing data
  
  sheet.appendRow(['CATEGORY', 'AMOUNT']);
  if (summaryData.length > 0) {
    sheet.getRange(2, 1, summaryData.length, 2).setValues(summaryData);
  }
  
  // Add title
  var titleCell = sheet.getRange('A1');
  titleCell.setValue(title).setFontSize(14).setFontWeight('bold');
  sheet.getRange('A1:B1').setBackground('#ffff00'); // Yellow background for header
}
// Helper function to create or update a pie chart
function createOrUpdatePieChart(sheet, chartTitle, total) {
  clearCharts(sheet);
  
  var chartBuilder = sheet.newChart();
  chartBuilder.setChartType(Charts.ChartType.PIE);
  chartBuilder.addRange(sheet.getRange('A2:B' + sheet.getLastRow()));
  chartBuilder.setOption('title', chartTitle + ' - Total: ' + total.toFixed(2));
  chartBuilder.setPosition(5, 5, 0, 0);
  
  sheet.insertChart(chartBuilder.build());
}
// Helper function to create or update a column chart
function createOrUpdateColumnChart(sheet, chartTitle) {
  clearCharts(sheet);
  
  var chartBuilder = sheet.newChart();
  chartBuilder.setChartType(Charts.ChartType.COLUMN);
  chartBuilder.addRange(sheet.getRange('A2:B4'));
  chartBuilder.setOption('title', chartTitle);
  chartBuilder.setPosition(7, 1, 0, 0);
  
  sheet.insertChart(chartBuilder.build());
}
// Helper function to clear existing charts
function clearCharts(sheet) {
  var charts = sheet.getCharts();
  charts.forEach(chart => sheet.removeChart(chart));
}
// Main function to handle end-of-day processing
function endOfDayProcessing() {
  var docId = createDailySummaryDoc();
  sendEmailWithDoc(docId, 'Daily Financial Summary');
  clearSummarySheets();
}
// Main function to handle end-of-month processing
function endOfMonthProcessing() {
  var docId = createMonthlySummaryDoc();
  sendEmailWithDoc(docId, 'Monthly Financial Summary');
}
// Function to create a Google Doc with daily summaries and charts
function createDailySummaryDoc() {
  var sheetId = '1B1yXKBZdzFUs3R55SUYDNfTrW0f7Er0bTQgh0N_Nxeo';
  var spreadsheet = SpreadsheetApp.openById(sheetId);
  var incomeSummarySheet = spreadsheet.getSheetByName('Income Summary');
  var expensesSummarySheet = spreadsheet.getSheetByName('Expenses Summary');
  var dailySummarySheet = spreadsheet.getSheetByName('Daily Summary');
  
  var today = new Date();
  var dateString = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MM-dd-yyyy');
  var docTitle = 'Daily Financial Summary - ' + dateString;
  var doc = DocumentApp.create(docTitle);
  var body = doc.getBody();
  
  body.appendParagraph(docTitle)
      .setHeading(DocumentApp.ParagraphHeading.TITLE)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .setFontSize(16)
      .setBold(true)
      .setForegroundColor('#003366'); // Dark blue
  appendSheetDataToDoc(body, incomeSummarySheet, 'Income Summary');
  appendSheetDataToDoc(body, expensesSummarySheet, 'Expenses Summary');
  appendSheetDataToDoc(body, dailySummarySheet, 'Daily Summary');
  return doc.getId();
}
// Function to update the Monthly Summary sheet
function updateMonthlySummary() {
  var sheetId = '1B1yXKBZdzFUs3R55SUYDNfTrW0f7Er0bTQgh0N_Nxeo';
  var spreadsheet = SpreadsheetApp.openById(sheetId);
  var incomeSummarySheet = spreadsheet.getSheetByName('Income Summary');
  var expensesSummarySheet = spreadsheet.getSheetByName('Expenses Summary');
  var monthlySummarySheet = spreadsheet.getSheetByName('Monthly Summary');
  
  if (!incomeSummarySheet || !expensesSummarySheet || !monthlySummarySheet) {
    Logger.log('Income Summary, Expenses Summary, or Monthly Summary sheet not found.');
    return;
  }
  
  var incomeSummaryData = incomeSummarySheet.getDataRange().getValues();
  var expensesSummaryData = expensesSummarySheet.getDataRange().getValues();
  
  var totalIncome = calculateTotal(incomeSummaryData.slice(1));
  var totalExpenses = calculateTotal(expensesSummaryData.slice(1));
  var profitLoss = totalIncome - totalExpenses;
  
  monthlySummarySheet.clear(); // Clear existing data
  
  monthlySummarySheet.appendRow(['Metric', 'Amount']);
  monthlySummarySheet.appendRow(['Total Income', totalIncome]);
  monthlySummarySheet.appendRow(['Total Expenses', totalExpenses]);
  var profitLossRow = ['Profit/Loss', profitLoss];
  monthlySummarySheet.appendRow(profitLossRow);
  
  // Format the Profit/Loss cell
  var profitLossCell = monthlySummarySheet.getRange('B4');
  if (profitLoss > 0) {
    profitLossCell.setBackground('green');
  } else {
    profitLossCell.setBackground('red');
  }
  // Add title and color code
  var titleCell = monthlySummarySheet.getRange('A1');
  titleCell.setValue('Monthly Financial Summary').setFontSize(14).setFontWeight('bold');
  monthlySummarySheet.getRange('A1:B1').setBackground('#d9ead3'); // Light green background for header
  createOrUpdateColumnChart(monthlySummarySheet, 'Monthly Profit/Loss');
}

// Helper function to calculate the total amount
function calculateTotal(summaryData) {
  return summaryData.reduce((total, row) => total + row[1], 0);
}

// Helper function to create or update a column chart
function createOrUpdateColumnChart(sheet, chartTitle) {
  clearCharts(sheet);
  
  var chartBuilder = sheet.newChart();
  chartBuilder.setChartType(Charts.ChartType.COLUMN);
  chartBuilder.addRange(sheet.getRange('A2:B4'));
  chartBuilder.setOption('title', chartTitle);
  chartBuilder.setPosition(7, 1, 0, 0);
  
  sheet.insertChart(chartBuilder.build());
}

// Helper function to clear existing charts
function clearCharts(sheet) {
  var charts = sheet.getCharts();
  charts.forEach(chart => sheet.removeChart(chart));
}

// Function to create a Google Doc with monthly summaries and charts
function createMonthlySummaryDoc() {
  var sheetId = '1B1yXKBZdzFUs3R55SUYDNfTrW0f7Er0bTQgh0N_Nxeo';
  var spreadsheet = SpreadsheetApp.openById(sheetId);
  var monthlySummarySheet = spreadsheet.getSheetByName('Monthly Summary');
  
  var today = new Date();
  var dateString = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MM-yyyy');
  var docTitle = 'Monthly Financial Summary - ' + dateString;
  var doc = DocumentApp.create(docTitle);
  var body = doc.getBody();
  
  body.appendParagraph(docTitle)
      .setHeading(DocumentApp.ParagraphHeading.TITLE)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .setFontSize(16)
      .setBold(true)
      .setForegroundColor('#003366'); // Dark blue
  appendSheetDataToDoc(body, monthlySummarySheet, 'Monthly Summary');
  return doc.getId();
}

// Add sheet data and charts to Google Doc
function appendSheetDataToDoc(body, sheet, title) {
  body.appendPageBreak(); // Add a page break before each section
  body.appendParagraph(title)
      .setHeading(DocumentApp.ParagraphHeading.HEADING1)
      .setForegroundColor('#333333'); // Dark grey
  var data = sheet.getRange('A1:B' + sheet.getLastRow()).getValues();
  if (data.length > 1) {
    var table = body.appendTable();
    data.forEach(function(row) {
      var tableRow = table.appendTableRow();
      tableRow.appendTableCell(row[0] ? row[0].toString() : '');
      tableRow.appendTableCell(row[1] ? row[1].toString() : '');
    });
  } else {
    body.appendParagraph('No data available for ' + title);
  }
  var charts = sheet.getCharts();
  charts.forEach(function(chart) {
    var blob = chart.getAs('image/png');
    body.appendImage(blob);
  });
}

// Send email with the Google Doc link
function sendEmailWithDoc(docId, summaryType) {
  var docUrl = 'https://docs.google.com/document/d/' + docId;
  var today = new Date();
  var dateString = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MM-dd-yyyy');
  var emailAddress = 'Jerzltaz@gmail.com'; // Replace with recipient's email address
  var subject = summaryType + ' - ' + dateString;
  var body = 'Please find the ' + summaryType + ' for ' + dateString + ' at the following link:\n\n' + docUrl;
  MailApp.sendEmail(emailAddress, subject, body);
}

// Function to clear the data from summary sheets but keep headers and charts
function clearSummarySheets() {
  var sheetId = '1B1yXKBZdzFUs3R55SUYDNfTrW0f7Er0bTQgh0N_Nxeo';
  var spreadsheet = SpreadsheetApp.openById(sheetId);
  
  var sheets = ['Income Summary', 'Expenses Summary', 'Daily Summary'];
  sheets.forEach(function(sheetName) {
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) {
      clearSheetData(sheet);
    } else {
      Logger.log(sheetName + ' sheet not found.');
    }
  });
}

// Helper function to clear data but keep headers and charts
function clearSheetData(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange('A2:B' + lastRow).clearContent(); // Clear data while keeping headers
  }
}

// Main function to handle end-of-month processing
function endOfMonthProcessing() {
  var docId = createMonthlySummaryDoc();
  sendEmailWithDoc(docId, 'Monthly Financial Summary');
  clearMonthlySummarySheet();
}

// Create a time-driven trigger to execute the end-of-month processing
function createMonthlyTrigger() {
  ScriptApp.newTrigger('endOfMonthProcessing')
    .timeBased()
    .onMonthDay(1)
    .atHour(0)
    .create();
}

// Function to clear the data from the Monthly Summary sheet but keep headers and charts
function clearMonthlySummarySheet() {
  var sheetId = '1B1yXKBZdzFUs3R55SUYDNfTrW0f7Er0bTQgh0N_Nxeo';
  var spreadsheet = SpreadsheetApp.openById(sheetId);
  var monthlySummarySheet = spreadsheet.getSheetByName('Monthly Summary');
  
  if (monthlySummarySheet) {
    clearSheetData(monthlySummarySheet);
  } else {
    Logger.log('Monthly Summary sheet not found.');
  }
}

// Helper function to clear data but keep headers and charts
function clearSheetData(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange('A2:B' + lastRow).clearContent(); // Clear data while keeping headers
  }
}

