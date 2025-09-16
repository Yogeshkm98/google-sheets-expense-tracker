# Google Sheets Personal Expense Tracker

A powerful Google Apps Script that automatically scans your Gmail for bank and credit card transaction alerts and logs them into a Google Sheet. It creates monthly and category-based summaries to give you a clear overview of your finances.



## ✨ Features

- **Automated Scanning:** Runs automatically to scan recent emails for new transactions.
- **Multi-Bank Support:** Pre-configured patterns for major Indian banks (HDFC, ICICI, SBI, Axis, IOB) and easily extendable.
- **Intelligent Data Extraction:** Reliably extracts key information like amount, merchant name (e.g., from `Info: BOOKMYSHOW`), transaction type (Debit/Credit), card number, and UPI details.
- **Advanced Duplicate Detection:** Prevents logging the same transaction twice using a combination of MessageID and transaction reference numbers.
- **Automatic Categorization:** Assigns categories like 'Food & Dining', 'Shopping', and 'Utilities' based on merchant names.
- **Dashboard Summaries:** Automatically generates 'Monthly Summary' and 'Category Analysis' sheets to visualize spending habits.
- **Email Notifications:** Sends you a summary email after new transactions have been processed.

## 🚀 Setup Instructions

1.  **Create a new Google Sheet:** Go to [sheets.new](https://sheets.new).
2.  **Open the Apps Script Editor:** In your new sheet, go to `Extensions` > `Apps Script`.
3.  **Paste the Code:**
    * Delete any placeholder code in the `Code.gs` file.
    * Copy the entire content of the `Code.gs` file from this repository and paste it into the editor.
4.  **Configure the Script:**
    * Update the `CONFIG` object at the top of the script:
        * `EMAIL_ACCOUNTS`: Add the email addresses you want to scan.
        * `DAYS_TO_SCAN`: Set how many days back the script should look for emails on each run (e.g., `2` for the last 48 hours).
5.  **Save the Project:** Click the floppy disk icon 💾 to save. Give your project a name when prompted (e.g., "Expense Tracker").
6.  **Run the Initial Setup:**
    * From the function dropdown list, select `scanAllExpenseEmails`.
    * Click the **"Run"** button.
    * **Authorization:** A pop-up will appear asking for permissions. This is required for the script to read your emails and write to your sheet. Click "Review permissions", select your Google account, and on the "unsafe" screen, click "Advanced" and then "Go to (your project name)".
7.  **Set Up Automated Trigger:**
    * To make the script run automatically, select `setupAutomatedTrigger` from the function dropdown and click **"Run"**.
    * This will set up a trigger to run the `scanAllExpenseEmails` function automatically every 6 hours.

Your expense tracker is now live! New transactions will be added to the 'All Transactions' sheet automatically.

## ⚙️ Configuration

All configuration is done in the `CONFIG` object at the top of the `Code.gs` file.

```javascript
const CONFIG = {
  MAIN_SHEET: 'All Transactions',
  SUMMARY_SHEET: 'Monthly Summary',
  CATEGORY_SHEET: 'Category Analysis',
  EMAIL_ACCOUNTS: [
    'your-email@gmail.com',
    'another-email@gmail.com'
  ],
  DAYS_TO_SCAN: 1, // How many days back to scan for emails
  MIN_AMOUNT: 1,   // Ignore transactions below this amount
  DEBUG: true,     // Set to false to reduce logging
};
