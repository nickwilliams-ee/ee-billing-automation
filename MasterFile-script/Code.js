function onOpen() {
  // Create a custom menu to allow users to manually delete sheets or run reports
  SpreadsheetApp.getUi()
    .createMenu('Billing Tools')
    .addItem('Show Sidebar Panel', 'showSidebar')
    .addItem('Show Architecture Diagram', 'showArchitectureDiagram')
    .addSeparator()
    .addItem('Delete Open Items Sheet', 'deleteOpenItemsSheet')
    .addItem('Output Approved Rows by Client & Project', 'outputApprovedRowsByClient')
    .addToUi();

  // Clear all filters across sheets on open
  clearAllFilters();
}

/**
 * Clears existing filters across all sheets in the active spreadsheet.
 */
function clearAllFilters() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    try {
      const filter = sheets[i].getFilter();
      if (filter !== null) {
        filter.remove();
      }
    } catch (e) {
      console.warn(`Could not clear filter on sheet "${sheets[i].getName()}": ${e.message}`);
    }
  }
}

/**
 * Opens a persistent sidebar panel containing action controls.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Billing Tools');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Opens a modal dialog displaying the system architecture diagram.
 */
function showArchitectureDiagram() {
  const html = HtmlService.createHtmlOutputFromFile('Architecture')
    .setWidth(1024)
    .setHeight(768)
    .setTitle('System Architecture Diagram');
  SpreadsheetApp.getUi().showModalDialog(html, 'System Architecture Diagram');
}

/**
 * Runs the full filter reset and data output process.
 * Can be triggered from the custom menu or via a sheet button.
 */
function runEntireProcess() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Billing Report"); // <-- Change to your tab name
  if (!sheet) return;

  // 1. Clear any existing filters to reset the data view
  if (sheet.getFilter() != null) {
    sheet.getFilter().remove();
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return; // Stop if the sheet is empty or only has headers
  
  // 2. Fetch allowed client names from "AU Billing Client" / "AU Billing Clients" tab
  const clientSheet = ss.getSheetByName("AU Billing Client") || ss.getSheetByName("AU Billing Clients");
  const clientValues = clientSheet && clientSheet.getLastRow() >= 1
    ? clientSheet.getRange(1, 1, clientSheet.getLastRow(), 1).getValues().flat().map(v => String(v).trim()).filter(Boolean)
    : [];

  const range = sheet.getRange("A1:V" + lastRow); // <-- Update "V" to match your last column letter
  const filter = range.createFilter();

  // Helper to calculate hidden values and set filter criteria per column
  function applyColumnFilter(colIndex, allowedList) {
    const cellValues = sheet.getRange(2, colIndex, lastRow - 1, 1).getValues();
    const valuesToHide = [];
    const allowedSet = new Set(allowedList.map(v => String(v).trim().toUpperCase()));

    for (let i = 0; i < cellValues.length; i++) {
      const currentVal = String(cellValues[i][0] ?? "");
      if (!allowedSet.has(currentVal.trim().toUpperCase()) && valuesToHide.indexOf(currentVal) === -1) {
        valuesToHide.push(currentVal);
      }
    }

    const criteria = SpreadsheetApp.newFilterCriteria()
      .setHiddenValues(valuesToHide)
      .build();

    filter.setColumnFilterCriteria(colIndex, criteria);
  }

  // 3. Apply criteria to Column 7 (Client - Column G)
  applyColumnFilter(7, clientValues);

  // Run output functions for Open Items, Cancelled, and Pending Approval status tabs
  outputOpenRows(clientValues);
  outputCancelledRows(clientValues);
  outputPendingApprovalRows(clientValues);

  // Run outputApprovedRowsByClient to create individual Google Sheet files by Client & Project
  outputApprovedRowsByClient(clientValues);
}

/**
 * Copies all rows with status "OPEN" or "DRAFT" matching specified filter criteria from "Billing Report"
 * into a new sheet named "Open Items".
 */
function outputOpenRows(clientValues) {
  outputStatusRows(clientValues, "Open Items", ["OPEN", "DRAFT"]);
}

/**
 * Copies all rows with status "CANCELLED" matching specified filter criteria from "Billing Report"
 * into a new sheet named "Cancelled".
 */
function outputCancelledRows(clientValues) {
  outputStatusRows(clientValues, "Cancelled", ["CANCELLED"]);
}

/**
 * Copies all rows with status "PENDING_APPROVAL" matching specified filter criteria from "Billing Report"
 * into a new sheet named "Pending Approval".
 */
function outputPendingApprovalRows(clientValues) {
  outputStatusRows(clientValues, "Pending Approval", ["PENDING_APPROVAL"]);
}

/**
 * Generic helper to extract rows from "Billing Report" matching target statuses and client criteria,
 * and output them into a designated tab in the active spreadsheet.
 */
function outputStatusRows(clientValues, targetSheetName, targetStatuses) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Display persistent status toast in the bottom right corner
  ss.toast(`Processing ${targetSheetName}...`, "Status", -1);

  const sourceSheet = ss.getSheetByName("Billing Report");
  if (!sourceSheet) {
    ss.toast("Error: Source sheet not found.", "Status", 5);
    throw new Error('Source sheet "Billing Report" was not found.');
  }

  const data = sourceSheet.getDataRange().getValues();
  if (data.length <= 1) {
    console.log("No data found in the source sheet.");
    ss.toast("No data found in Billing Report.", "Status", 5);
    return;
  }

  // Fetch allowed clients if not passed or if an invalid argument type was passed
  if (!clientValues || !Array.isArray(clientValues) || clientValues.length === 0) {
    const clientSheet = ss.getSheetByName("AU Billing Client") || ss.getSheetByName("AU Billing Clients");
    clientValues = clientSheet && clientSheet.getLastRow() >= 1
      ? clientSheet.getRange(1, 1, clientSheet.getLastRow(), 1).getValues().flat().map(v => String(v).trim()).filter(Boolean)
      : [];
  }

  const header = data[0];
  const matchingRows = [header];

  const allowedClients = clientValues || [];
  const clientSet = new Set(allowedClients.map(v => String(v).trim().toUpperCase()));
  const statusSet = new Set(targetStatuses.map(s => String(s).trim().toUpperCase()));
  const targetSupplier = "EQUAL EXPERTS AUSTRALIA PTY LTD";

  // Column C = 0-based index 2 (Status)
  // Column G = 0-based index 6 (Client)
  // Column J = 0-based index 9 (Supplier Legal Entity)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const clientValue = String(row[6] || "").trim().toUpperCase();

    if (clientSet.size > 0 && !clientSet.has(clientValue)) continue;

    const supplierValue = String(row[9] || "").trim().toUpperCase();
    if (supplierValue !== targetSupplier) continue;

    const statusValue = String(row[2] || "").trim().toUpperCase();
    if (statusSet.has(statusValue)) {
      matchingRows.push(row);
    }
  }

  let targetSheet = ss.getSheetByName(targetSheetName);
  
  if (!targetSheet) {
    targetSheet = ss.insertSheet(targetSheetName);
  } else {
    targetSheet.clear();
  }

  if (matchingRows.length > 1) {
    targetSheet.getRange(1, 1, matchingRows.length, matchingRows[0].length).setValues(matchingRows);
    console.log(`Successfully output ${matchingRows.length - 1} records to "${targetSheetName}".`);
  } else {
    // Output headers even if no matching rows exist
    targetSheet.getRange(1, 1, 1, header.length).setValues([header]);
    console.log(`No rows with status matching [${targetStatuses.join(", ")}] were found for "${targetSheetName}".`);
  }

  // Update status toast on completion
  ss.toast(`${targetSheetName} processing completed!`, "Status", 5);
}

/**
 * Filters rows with status "APPROVED" matching Client criteria,
 * and outputs specified columns into Google Sheet files in the "SubFile-Output" folder.
 * For "Equal Experts" client, all projects are combined into a single file.
 */
function outputApprovedRowsByClient(clientValues) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Display persistent status toast in the bottom right corner
  ss.toast("Processing approved rows by Client & Project...", "Status", -1);

  // Locate the "SubFile-Output" folder in Google Drive
  const folderIterator = DriveApp.getFoldersByName("SubFile-Output");
  if (!folderIterator.hasNext()) {
    ss.toast("Error: Folder 'SubFile-Output' not found in Drive.", "Status", 5);
    throw new Error('Folder "SubFile-Output" was not found in Google Drive. Please ensure it exists.');
  }
  const targetFolder = folderIterator.next();

  // Locate the "Archive" subfolder inside "SubFile-Output"
  const archiveIterator = targetFolder.getFoldersByName("Archive");
  if (!archiveIterator.hasNext()) {
    ss.toast("Error: Subfolder 'Archive' not found in SubFile-Output.", "Status", 5);
    throw new Error('Subfolder "Archive" was not found inside "SubFile-Output". Please ensure it exists.');
  }
  const archiveFolder = archiveIterator.next();

  // Move existing files in the "SubFile-Output" folder to the "Archive" subfolder with a timestamp appended
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  const existingFiles = targetFolder.getFiles();
  while (existingFiles.hasNext()) {
    const file = existingFiles.next();
    const originalName = file.getName();
    file.setName(`${originalName}_${timestamp}`);
    file.moveTo(archiveFolder);
    ss.toast(`Archived "${originalName}"`, "Archiving", 3);
  }

  const sourceSheet = ss.getSheetByName("Billing Report");
  if (!sourceSheet) {
    ss.toast("Error: Source sheet not found.", "Status", 5);
    throw new Error('Source sheet "Billing Report" was not found.');
  }

  const data = sourceSheet.getDataRange().getValues();
  if (data.length <= 1) {
    console.log("No data found in the source sheet.");
    ss.toast("No data found in Billing Report.", "Status", 5);
    return;
  }

  if (!clientValues || !Array.isArray(clientValues) || clientValues.length === 0) {
    const clientSheet = ss.getSheetByName("AU Billing Client") || ss.getSheetByName("AU Billing Clients");
    clientValues = clientSheet && clientSheet.getLastRow() >= 1
      ? clientSheet.getRange(1, 1, clientSheet.getLastRow(), 1).getValues().flat().map(v => String(v).trim()).filter(Boolean)
      : [];
  }

  const allowedClients = clientValues || [];
  const clientSet = new Set(allowedClients.map(v => String(v).trim().toUpperCase()));
  const targetSupplier = "EQUAL EXPERTS AUSTRALIA PTY LTD";

  // Define requested output columns and their 0-based indices from "Billing Report"
  const targetColumnIndices = [1, 2, 6, 7, 8, 9, 11, 12, 13, 14, 16, 17, 18, 19, 20];
  const targetHeaders = [
    "ID", "Status", "Client", "Engagement", "Project",
    "Supplier Legal Entity", "Timesheet Start", "Timesheet End",
    "Consultant Name", "Unit Type", "Number of Days",
    "Bill Rate", "Bill Rate Currency", "Total Billed", "Approved By"
  ];

  // Group approved rows in memory by Client and Project (or combine all projects for "Equal Experts")
  const groups = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const clientValue = String(row[6] || "").trim().toUpperCase(); // Column G (Client)
    
    // Check filter criteria in memory
    if (clientSet.size > 0 && !clientSet.has(clientValue)) continue;

    const supplierValue = String(row[9] || "").trim().toUpperCase(); // Column J (Supplier Legal Entity)
    if (supplierValue !== targetSupplier) continue;

    const status = String(row[2] || "").trim().toUpperCase();
    if (status === "APPROVED") {
      const client = String(row[6] || "").trim() || "Unspecified Client";
      const project = String(row[8] || "").trim() || "Unspecified Project";
      
      const isEqualExperts = client.toUpperCase() === "EQUAL EXPERTS" || client.toUpperCase() === "EQUAL EXPERTS AUSTRALIA PTY LTD";
      const effectiveProject = isEqualExperts ? "All Projects" : project;
      const groupKey = isEqualExperts ? "EQUAL EXPERTS" : `${client}___${project}`;
      
      if (!groups[groupKey]) {
        groups[groupKey] = {
          client: client,
          project: effectiveProject,
          rows: []
        };
      }
      
      // Extract only the required target columns for this row
      const mappedRow = targetColumnIndices.map(colIdx => row[colIdx]);
      groups[groupKey].rows.push(mappedRow);
    }
  }

  const groupKeys = Object.keys(groups);
  if (groupKeys.length === 0) {
    console.log('No rows with status "APPROVED" matching the filter criteria were found.');
    ss.toast("No approved rows found matching criteria.", "Status", 5);
    try {
      SpreadsheetApp.getUi().alert("Process Completed.\n\nNo approved rows found matching the criteria.");
    } catch (e) {
      // Ignore UI alert error if running outside active UI context
    }
    return;
  }

  const createdLogs = [];
  const totalGroups = groupKeys.length;

  groupKeys.forEach((key, index) => {
    const group = groups[key];
    const clientName = group.client;
    const projectName = group.project;
    const rows = group.rows;

    // Infer Month and Year from the latest Timesheet End date (mapped index 7)
    let maxEndDate = null;
    rows.forEach(r => {
      const endVal = r[7];
      if (endVal) {
        const d = (endVal instanceof Date) ? endVal : new Date(String(endVal).replace(/-/g, '/'));
        if (!isNaN(d.getTime())) {
          if (!maxEndDate || d > maxEndDate) {
            maxEndDate = d;
          }
        }
      }
    });

    const targetDate = maxEndDate || new Date();
    const yearMonth = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), "yyyy.MM");
    const monthYear = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), "MMMM yyyy");

    // Construct file name: YYYY.MM Client Name Billing - Project Name - full month spelt out with yyyy
    const rawFileName = `${yearMonth} ${clientName} Billing - ${projectName} - ${monthYear}`;
    const safeFileName = rawFileName.replace(/[\\\/:\?\*\[\]]/g, "_").substring(0, 150);

    // Create a new standalone Google Sheet file in Drive and move it into SubFile-Output
    const newSs = SpreadsheetApp.create(safeFileName);
    const newFile = DriveApp.getFileById(newSs.getId());
    newFile.moveTo(targetFolder);

    const targetSheet = newSs.getActiveSheet();
    targetSheet.setName("Billing Output");

    // Sort data rows by Consultant Name (mapped index 8)
    rows.sort((a, b) => {
      const consultantA = String(a[8] || "");
      const consultantB = String(b[8] || "");
      return consultantA.localeCompare(consultantB);
    });

    const titleText = `${yearMonth} ${clientName} Billing - ${projectName} - ${monthYear}`;

    // Row 1: Header title in bold, larger text (14pt)
    const titleCell = targetSheet.getRange(1, 1);
    titleCell.setValue(titleText);
    titleCell.setFontWeight("bold");
    titleCell.setFontSize(14);

    // Row 3: Table headers starting on Row 3 with dark blue highlight and white text
    const headerRange = targetSheet.getRange(3, 1, 1, targetHeaders.length);
    headerRange.setValues([targetHeaders]);
    headerRange.setBackground("#0b5394");
    headerRange.setFontColor("#ffffff");
    headerRange.setFontWeight("bold");

    // Row 4 onwards: Output data rows
    if (rows.length > 0) {
      targetSheet.getRange(4, 1, rows.length, targetHeaders.length).setValues(rows);
    }

    // Calculate last data row index (data starts at Row 4)
    const lastDataRow = 3 + rows.length;
    const summaryRow = lastDataRow + 2;

    // Output SUM formulas for Column K (Number of Days) and Column N (Total Billed), and "AUD" for Column M
    const kSummaryCell = targetSheet.getRange(summaryRow, 11);
    const mSummaryCell = targetSheet.getRange(summaryRow, 13);
    const nSummaryCell = targetSheet.getRange(summaryRow, 14);

    kSummaryCell.setFormula('=SUM(K4:K' + lastDataRow + ')').setFontColor('#ff0000').setFontWeight('bold');
    mSummaryCell.setValue('AUD').setFontColor('#ff0000').setFontWeight('bold');
    nSummaryCell.setFormula('=SUM(N4:N' + lastDataRow + ')').setFontColor('#ff0000').setFontWeight('bold');

    const logEntry = `Successfully created file "${safeFileName}" with ${rows.length} rows in SubFile-Output.`;
    console.log(logEntry);
    createdLogs.push(logEntry);

    // Live status notification as each file is processed
    ss.toast(`Created [${index + 1}/${totalGroups}]: "${safeFileName}" (${rows.length} rows)`, "Progress", 3);
  });

  // Update status toast on completion
  ss.toast("Process completed!", "Status", 5);

  // Show pop-up alert dialog with all accumulated logs
  const completionMessage = "Process Completed.\n\nLogs:\n" + createdLogs.join("\n");
  try {
    SpreadsheetApp.getUi().alert(completionMessage);
  } catch (e) {
    console.log(completionMessage);
  }
}

// ⚠️ Transmits spreadsheet data to external Slack webhooks - verify webhook endpoints before executing.
/**
 * Reads all items from both "Open Items" and "Pending Approval" sheets and posts structured JSON objects
 * formatted specifically for Slack Workflow Builder schema endpoints.
 */
function sendOpenItemsToSlack() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Sending Open & Pending Items to Slack...", "Status", -1);

  const sheetNames = ["Open Items", "Pending Approval"];
  const allRows = [];

  sheetNames.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      if (data.length > 1) {
        allRows.push(...data.slice(1));
      }
    } else {
      console.warn(`Sheet '${sheetName}' not found.`);
    }
  });

  if (allRows.length === 0) {
    ss.toast("No open or pending items found to send.", "Status", 5);
    return;
  }

  // Webhook URLs are stored in Script Properties (Project Settings > Script Properties), not in code,
  // so they are never committed to the repo.
  const props = PropertiesService.getScriptProperties();
  const SLACK_WEBHOOK_URL_1 = props.getProperty('SLACK_WEBHOOK_URL_1');

  // To add Person 2: create a SLACK_WEBHOOK_URL_2 Script Property and add it to recipients below.
  //const SLACK_WEBHOOK_URL_2 = props.getProperty('SLACK_WEBHOOK_URL_2');

  if (!SLACK_WEBHOOK_URL_1) {
    ss.toast("Slack webhook not configured. Add SLACK_WEBHOOK_URL_1 in Project Settings > Script Properties.", "Error", 10);
    return;
  }

  const recipients = [
    { name: "Person 1", url: SLACK_WEBHOOK_URL_1 }
  ];

  // Send structured JSON variables expected by Slack Webhooks / Workflow Builder
  allRows.forEach((row, idx) => {
    const id = String(row[1] || row[0] || idx + 1).trim(); // Fallback to Column A (Unique ID) if Column B (ID) is empty
    const status = String(row[2] || "N/A").trim(); // Column C (Status)
    const consultantName = String(row[13] || "N/A").trim(); // Column N (Consultant Name)

    const payloadObj = {
      "Consultant-Name": consultantName,
      "Status": status
    };

    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payloadObj),
      muteHttpExceptions: true
    };

    recipients.forEach(recipient => {
      if (recipient.url && !recipient.url.includes("INSERT_SLACK_WEBHOOK")) {
        try {
          const response = UrlFetchApp.fetch(recipient.url, options);
          const code = response.getResponseCode();
          if (code >= 200 && code < 300) {
            console.log(`Successfully sent item (${id}) for ${consultantName} [Status: ${status}] to ${recipient.name}. HTTP ${code}`);
          } else {
            console.warn(`Failed to send item (${id}) to ${recipient.name}. HTTP ${code}: ${response.getContentText()}`);
          }
        } catch (e) {
          console.error(`Error sending Slack message to ${recipient.name}: ${e.message}`);
        }
      }
    });
  });

  ss.toast("Items sent to Slack!", "Status", 5);
}

/**
 * Deletes the "Open Items" sheet if it exists.
 */
function deleteOpenItemsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Deleting Open Items sheet...", "Status", -1);

  const sheet = ss.getSheetByName("Open Items");
  if (sheet) {
    ss.deleteSheet(sheet);
    console.log('Sheet "Open Items" deleted successfully.');
    ss.toast("Sheet \"Open Items\" deleted successfully.", "Status", 5);
  } else {
    console.log('Sheet "Open Items" does not exist.');
    ss.toast("Sheet \"Open Items\" does not exist.", "Status", 5);
  }
}
