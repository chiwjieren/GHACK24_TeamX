# 📈 Financial Summary Automation 📊( SMARTKIRA )

Welcome to the **Financial Summary Automation** project! This tool simplifies financial tracking and reporting by automating daily and monthly summaries, complete with charts and email notifications.

## 🌟 Features

- **📅 Daily Reports**: Automated daily financial summaries with detailed breakdowns of income, expenses, and profit/loss.
- **📆 Monthly Overviews**: Comprehensive monthly summaries with visual charts.
- **💡 Real-time Insights**: Data and charts are updated with the latest information.
- **🧹 Data Management**: Old data is cleared automatically while retaining essential headers and charts.

## 🛠 Setup Instructions

### 1. Google Sheets Setup 📝

Create a Google Spreadsheet with the following sheets:

- **Income**: Track your income entries here. (create a google forms and link the sheets to this sheet rename = Income)
- **Expenses**: Log your expenses here.(create a google forms and link the sheets to this sheet rename = Expenses)
- the forms contains (Category, Remarks, Amount)
- **Income Summary**: Summarize income with a pie chart.
- **Expenses Summary**: Summarize expenses with a pie chart.
- **Daily Summary**: Daily overview with a column chart.
- **Monthly Summary**: Monthly overview with a column chart.

### 2. Integrate with Google Apps Script 💻

1. Open your Google Spreadsheet.
2. Go to `Extensions > Apps Script`.
3. Paste the provided script into the script editor.
4. Update the `sheetId` variable with your spreadsheet ID.
5. Set up time-driven triggers for automatic daily and monthly processing.

### 3. Grant Permissions 🔐

Authorize the script to access your Google Sheets and send emails on your behalf.

## 🚀 Functions Overview

- **`updateIncome()`**: Refreshes the Income Summary and updates the pie chart.
- **`updateExpenses()`**: Refreshes the Expenses Summary and updates the pie chart.
- **`updateDailySummary()`**: Updates the Daily Summary with a column chart.
- **`updateMonthlySummary()`**: Updates the Monthly Summary with a column chart.
- **`createDailySummaryDoc()`**: Creates a Google Doc with daily summaries and charts.
- **`createMonthlySummaryDoc()`**: Creates a Google Doc with monthly summaries and charts.
- **`sendEmailWithDoc(docId, summaryType)`**: Sends an email with a link to the Google Doc.
- **`clearSummarySheets()`**: Clears data in summary sheets while retaining headers and charts.
- **`clearMonthlySummarySheet()`**: Clears data in the Monthly Summary sheet, keeping headers and charts.

## 🔄 Usage

- **Daily Processing**: Run `endOfDayProcessing()` to generate and email the daily financial report.
- **Monthly Processing**: Run `endOfMonthProcessing()` to generate and email the monthly financial report.
- **Scheduling**: Use `createDailyTrigger()` and `createMonthlyTrigger()` to automate daily and monthly reports.

## 🤝 Contributing

We welcome contributions to improve this project! Feel free to submit a pull request or open an issue with suggestions.

## 📬 Contact

For questions or feedback, reach out to [Chiw Jie Ren] at [Jerzltaz@gmail.com].

---

### 🖼️ Screenshots

**Income Sheet:**

![Screenshot 2024-07-26 135043](https://github.com/user-attachments/assets/f0eb3470-54f3-4dbc-9565-7cf65bca5fd8)

**Income Summary Sheet:**

![Screenshot 2024-07-26 135053](https://github.com/user-attachments/assets/a6c8d18b-8934-4e60-b99a-b204da6d6323)

**Expenses Sheet:**

![Screenshot 2024-07-26 135014](https://github.com/user-attachments/assets/62cbc89b-b279-4127-b141-2586c136d588)

**Expenses Summary Sheet:**

![Screenshot 2024-07-26 135027](https://github.com/user-attachments/assets/2bec3607-61ea-4fc6-895f-7109ea2797bc)

**Daily Summary sheets:**

![Screenshot 2024-07-26 135103](https://github.com/user-attachments/assets/fae331fd-e15d-4e57-a815-69e321b595e1)

**Monthly Summary sheets:**

![Screenshot 2024-07-26 135119](https://github.com/user-attachments/assets/0622d313-423d-4398-9caa-dee6b2de6e0b)

**Daily Financial Summary Example:**

![Screenshot 2024-07-26 135807](https://github.com/user-attachments/assets/640ca649-1d73-4ae0-9fd9-a66f607afbf8)

**Monthly Financial Overview Example:**

![Screenshot 2024-07-26 135930](https://github.com/user-attachments/assets/c2789057-52f7-4208-87ab-79b3ffdcafcb)

