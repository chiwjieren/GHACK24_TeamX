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

- **Income**: Track your income entries here.
- **Expenses**: Log your expenses here.
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

## 📝 License

This project is licensed under the MIT License. For details, see the [LICENSE](LICENSE) file.

## 📬 Contact

For questions or feedback, reach out to [Your Name] at [Your Email Address].

---

### 🖼️ Screenshots

**Income Sheet:**

![Income Summary](https://via.placeholder.com/400x200?text=Income+Summary+Sheet)

**Income Summary Sheet:**

**Expenses Summary Sheet:**

![Expenses Summary](https://via.placeholder.com/400x200?text=Expenses+Summary+Sheet)

**Daily Summary sheets:**

![Daily Summary Doc](https://via.placeholder.com/400x200?text=Daily+Summary+Document)

**Daily Financial Summary Example:**

![Daily Summary](https://via.placeholder.com/400x200?text=Daily+Summary+Chart)

**Monthly Financial Overview Example:**

![Monthly Summary](https://via.placeholder.com/400x200?text=Monthly+Summary+Chart)
