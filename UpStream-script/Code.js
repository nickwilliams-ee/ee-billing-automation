/**
 * Main function to extract Google Doc/Sheet data and write values to Column G in the 'Example' sheet.
 * Also creates or replaces a new tab named "Client Name - Project Name - Contract Date" highlighted in green.
 */
function extractSoWData() {
  const DOC_NAME = 'Billing_Addendum_TEMPLATE-DO-NOT-EDIT';
  const SHEET_NAME = 'Example';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    const sheetNames = ss.getSheets().map(s => s.getName());
    throw new Error(`Sheet "${SHEET_NAME}" not found. Available sheets: ${sheetNames.join(', ')}`);
  }

  // Get the parent folder of the active spreadsheet
  const ssFile = DriveApp.getFileById(ss.getId());
  const parents = ssFile.getParents();
  if (!parents.hasNext()) {
    throw new Error('Spreadsheet parent folder could not be determined.');
  }
  const folder = parents.next();

  // Find a native Google Sheet or Google Doc file in the folder
  const altName = DOC_NAME.replace(/\.pdf$/i, '');

  let docFile = null;

  let files = folder.getFilesByName(altName);
  while (files.hasNext()) {
    const f = files.next();
    const mime = f.getMimeType();
    if (mime === MimeType.GOOGLE_SHEETS || mime === MimeType.GOOGLE_DOCS) {
      docFile = f;
      break;
    }
  }

  if (!docFile) {
    files = folder.getFilesByName(DOC_NAME);
    while (files.hasNext()) {
      const f = files.next();
      const mime = f.getMimeType();
      if (mime === MimeType.GOOGLE_SHEETS || mime === MimeType.GOOGLE_DOCS) {
        docFile = f;
        break;
      }
    }
  }

  if (!docFile) {
    throw new Error(`File "${altName}" was not found in folder "${folder.getName()}". Please ensure the file is saved in Google Sheets or Google Docs format.`);
  }

  console.log(`Processing file: ${docFile.getName()}...`);

  // Extract text directly from the Google Sheet or Doc
  const docText = readDocText(docFile);

  // Read labels from Column F
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    console.warn('No label rows found in sheet.');
    return;
  }

  const rangeF = sheet.getRange(1, 6, lastRow, 1);
  const labels = rangeF.getValues();

  // Process each row in Column F
  const results = [];
  let clientName = '';
  let projectName = '';
  let contractDate = '';

  for (let i = 0; i < labels.length; i++) {
    const label = String(labels[i][0]).trim();
    if (!label) {
      results.push(['']);
      continue;
    }

    const extractedValue = findValueForLabel(docText, label);
    results.push([extractedValue]);

    // Store key metadata fields for sheet creation
    const cleanedLabel = label.toLowerCase().replace(/:\s*$/, '');
    if (cleanedLabel === 'client name') clientName = extractedValue;
    if (cleanedLabel === 'project name') projectName = extractedValue;
    if (cleanedLabel === 'contract date') contractDate = extractedValue;
  }

  // Fallback to explicit cells in Column G if labels were missing
  if (!clientName) clientName = String(sheet.getRange('G24').getDisplayValue()).trim();
  if (!projectName) projectName = String(sheet.getRange('G25').getDisplayValue()).trim();
  if (!contractDate) contractDate = String(sheet.getRange('G26').getDisplayValue()).trim() || String(sheet.getRange('G2').getDisplayValue()).trim();

  // Write the extracted values to Column G in Example sheet
  sheet.getRange(1, 7, results.length, 1).setValues(results);
  if (contractDate) {
    sheet.getRange('G26').setValue(contractDate);
  }
  console.log('Data extraction completed and written to Column G.');

  // Create or replace tab named "Client Name - Project Name - Contract Date"
  if (clientName && projectName && contractDate) {
    let tabName = clientName + ' - ' + projectName + ' - ' + contractDate;
    // Sanitize sheet name for Google Sheets invalid characters: \ / ? * : [ ] and max length
    tabName = tabName.replace(/[\\/?*:[\]]/g, '').substring(0, 100).trim();

    // Check if tab already exists; if so, delete to replace
    const existingSheet = ss.getSheetByName(tabName);
    if (existingSheet) {
      ss.deleteSheet(existingSheet);
    }

    // Copy Example sheet to preserve original
    const newSheet = sheet.copyTo(ss);
    newSheet.setName(tabName);
    newSheet.setTabColor('#34a853'); // Green tab color
    ss.setActiveSheet(newSheet);
    ss.moveActiveSheet(sheet.getIndex() + 1); // Move directly to the right of Example sheet
    console.log('Created new active tab: "' + tabName + '" highlighted in green directly after Example sheet.');
  } else {
    console.warn('Could not create new tab: Client Name, Project Name, or Contract Date is missing.');
  }
}

/**
 * Reads plain text content directly from a Google Doc or Google Sheet file.
 * @param {GoogleAppsScript.Drive.File} docFile
 * @return {string} Extracted text content
 */
function readDocText(docFile) {
  const mimeType = docFile.getMimeType();
  if (mimeType === MimeType.GOOGLE_SHEETS) {
    const sourceSs = SpreadsheetApp.openById(docFile.getId());
    return sourceSs.getSheets()
      .map(s => s.getDataRange().getValues().map(row => row.join(' ')).join('\n'))
      .join('\n');
  }
  if (mimeType === MimeType.GOOGLE_DOCS) {
    const doc = DocumentApp.openById(docFile.getId());
    return doc.getBody().getText();
  }
  throw new Error(`File "${docFile.getName()}" is not a Google Sheet or Google Doc (MIME type: ${mimeType}).`);
}

/**
 * Parses extracted text to find the value associated with a label.
 * @param {string} fullText
 * @param {string} label
 * @return {string}
 */
function findValueForLabel(fullText, label) {
  const cleanedLabel = label.replace(/:\s*$/, '').trim();
  if (!cleanedLabel) return '';

  // Special pattern matching for "Client Name"
  if (cleanedLabel.toLowerCase() === 'client name') {
    const clientRegex = /between\s+Equal\s+Experts\s+Australia\s+Pty\s+Limited\s*(?:\([^)]*\))?\s+and\s+(.*?)(?:\s*\(A[BC]N|\.|$)/i;
    const match = fullText.match(clientRegex);
    if (match && match[1]) {
      return match[1].trim();
    }
  }

  const escapedLabel = cleanedLabel.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const regex = new RegExp(escapedLabel + '\\s*:?\\s*(.*)', 'i');
  const lines = fullText.split('\n');

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    const match = line.match(regex);
    if (match) {
      let val = match[1].trim();
      // If value is not on the same line, check the immediate next line
      if (!val && i + 1 < lines.length) {
        val = lines[i + 1].trim();
      }
      return val;
    }
  }
  return '';
}

/**
 * Adds a custom menu to Google Sheets toolbar on document open.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Billing Automation')
    .addItem('Show Right Panel', 'showRightPanel')
    .addItem('Extract SoW Data', 'extractSoWData')
    .addToUi();
}

/**
 * Opens the right panel (sidebar) in Google Sheets.
 */
function showRightPanel() {
  const template = HtmlService.createTemplateFromFile('DiagramSidebar');
  template.webAppUrl = getWebAppUrl();
  const html = template.evaluate()
    .setTitle('Billing Automation');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Helper to fetch the published Web App URL.
 * @return {string}
 */
function getWebAppUrl() {
  // [ACTION_REQUIRED]: Update with your published web app URL
  const WEB_APP_URL_OVERRIDE = 'https://script.google.com/a/macros/equalexperts.com/s/AKfycbzj6KxlyvPbnTUWBrcqVDaPhUW5PuOdQ70a0EwKjw/exec';

  const scriptUrl = ScriptApp.getService().getUrl();
  const resolvedUrl = WEB_APP_URL_OVERRIDE || scriptUrl;

  if (WEB_APP_URL_OVERRIDE) {
    console.log('getWebAppUrl resolved via explicit override URL: ' + WEB_APP_URL_OVERRIDE);
  } else if (scriptUrl) {
    console.log('getWebAppUrl resolved via ScriptApp service URL: ' + scriptUrl);
  } else {
    console.warn('ScriptApp.getService().getUrl() returned empty/null.');
  }

  return resolvedUrl;
}

/**
 * Web App entry point that serves the Index HTML interface.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('System HLD Architecture')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Optional helper function to supply dynamic system configuration data.
 */
function getSystemComponents() {
  return {
    database: "Google Sheet (Active)",
    triggers: "Time-driven (Every 1 hour)"
  };
}
