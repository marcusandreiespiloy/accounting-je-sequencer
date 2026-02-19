/**
 * Google Apps Script: Financial Journal Entry (JE) Sequencer
 * Main Purpose: Automates JE numbering based on changes in Voucher Numbers (CV) or Memos.
 * Useful for: Accountants locking in manpower costs before moving to variable allowances.
 */

const CONFIG = {
  SHEET_NAME: "TEMPLATE",
  COLUMN_CV: 2,    // Column C (0-indexed)
  COLUMN_JE: 7,    // Column H (0-indexed)
  COLUMN_MEMO: 13, // Column N (0-indexed)
  DEFAULT_START_ROW: 3
};

function updateTemplateByCVAndMemo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const tempSheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!tempSheet) {
    ui.alert(`Error: '${CONFIG.SHEET_NAME}' sheet not found.`);
    return;
  }

  // 1. User Inputs
  const startRowInput = ui.prompt("Start Row", "Enter the Row Number to start (e.g., 3):", ui.ButtonSet.OK).getResponseText();
  const startRowNum = parseInt(startRowInput) || CONFIG.DEFAULT_START_ROW;

  const startJEInput = ui.prompt("Start JE", "Enter the Starting JE (e.g., JE300):", ui.ButtonSet.OK).getResponseText();
  const jePrefix = startJEInput.replace(/[0-9]/g, '');
  const startNumber = parseInt(startJEInput.replace(/\D/g, ''));

  const stopJEInput = ui.prompt("Stop JE", "Enter the LAST JE Number allowed:", ui.ButtonSet.OK).getResponseText();
  const stopNumber = parseInt(stopJEInput.replace(/\D/g, ''));

  // 2. Data Processing
  const lastRow = tempSheet.getLastRow();
  if (lastRow < startRowNum) return;
  
  const rangeHeight = lastRow - startRowNum + 1;
  const data = tempSheet.getRange(startRowNum, 1, rangeHeight, 14).getValues();
  
  let updatedJEValues = [];
  let currentJENum = startNumber;
  let lastCV = "";
  let lastMemo = "";

  for (let i = 0; i < data.length; i++) {
    let currentCV = (data[i][CONFIG.COLUMN_CV] || "").toString().trim();
    let currentMemo = (data[i][CONFIG.COLUMN_MEMO] || "").toString().trim();

    if (!currentCV && !currentMemo) {
      updatedJEValues.push([""]);
      continue;
    }

    // Logic: Increment JE if CV or Memo changes from the previous row
    if (lastCV !== "" && (currentCV !== lastCV || currentMemo !== lastMemo)) {
      currentJENum++; 
    }

    if (currentJENum > stopNumber) {
      ui.alert(`Stop limit reached at Row ${startRowNum + i}.`);
      break; 
    }

    updatedJEValues.push([jePrefix + currentJENum]);
    lastCV = currentCV;
    lastMemo = currentMemo;
  }

  // 3. Write Output
  if (updatedJEValues.length > 0) {
    tempSheet.getRange(startRowNum, CONFIG.COLUMN_JE + 1, updatedJEValues.length, 1).setValues(updatedJEValues);
    ui.alert(`Sequence Complete!\nLast JE used: ${jePrefix}${currentJENum}`);
  }
}
