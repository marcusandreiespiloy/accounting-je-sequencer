# Google Sheets JE Sequencer

A Google Apps Script utility for accountants to automate Journal Entry (JE) numbering based on transaction groupings.

## 📌 Purpose
When processing large financial datasets (like manpower costs or meal allowances), multiple rows often belong to a single Journal Entry. This script eliminates manual numbering by grouping rows based on:
1. **Check Voucher (CV) Numbers**
2. **Line Memos/Descriptions**

It ensures that as long as the Voucher and Memo remain the same, the JE number stays locked. Once either changes, the sequence increments.

## 🚀 How It Works
- **Smart Detection:** Monitors changes in specific columns to decide when to "step" the JE number.
- **Prefix Support:** Handles alphanumeric formats (e.g., "JE100" -> "JE101").
- **Safety Limits:** Includes a "Stop JE" prompt to prevent overwriting existing data or exceeding a budget range.

## 🛠️ Setup
1. Open your Google Sheet.
2. Go to **Extensions > Apps Script**.
3. Paste the code from `Code.gs`.
4. Adjust the `CONFIG` object at the top of the script if your columns differ:
   - `COLUMN_CV`: The column containing your Voucher/Reference numbers.
   - `COLUMN_JE`: The target column where JE numbers should be written.
   - `COLUMN_MEMO`: The column containing line descriptions.

## 📖 Usage
Run the function `updateTemplateByCVAndMemo`. You will be prompted for:
1. The row to start processing.
2. The starting JE number (e.g., `JE300`).
3. The maximum JE number allowed (to prevent run-on errors).
