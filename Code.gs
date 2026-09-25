// Add this at the very top of Code.gs
// We extract JUST the ID from the URL: 1bELSUI7xnIyqNxUPMOFTU8uRc9sGsIhkEre6LkVenmI
const SHEET_ID = '1bELSUI7xnIyqNxUPMOFTU8uRc9sGsIhkEre6LkVenmI'; 

let _cachedSpreadsheet = null;
function getSpreadsheet() {
  if (!_cachedSpreadsheet) {
    _cachedSpreadsheet = SpreadsheetApp.openById(SHEET_ID);
  }
  return _cachedSpreadsheet;
}

// 1. SERVE THE WEB APP
// This is the required function that runs when someone visits your web app URL.
function doGet(e) {
  // Evaluates the Index.html file so it can render on the screen
  var htmlOutput = HtmlService.createTemplateFromFile('Index').evaluate();
  
  // Sets the title of the browser tab and makes it mobile responsive
  htmlOutput.setTitle('BPM Maintenance Portal')
            .addMetaTag('viewport', 'width=device-width, initial-scale=1');
            
  return htmlOutput;
}

// 1.5 INCLUDE HTML FILES
// Helper function to include HTML files in the main Index.html file
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// 2. GET USER ROLE (Passed to frontend on load)
function getUserRole() {
  // Use both Active and Effective user methods to guarantee we capture an email
  var userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'Developer';
  var ss = getSpreadsheet();
  var empSheet = ss.getSheetByName('dim_Employees');
  
  var userInfo = {
    email: userEmail,
    role: 'admin', // Default to admin so you don't get locked out
    employeeId: 'EMP-001',
    name: userEmail // CRITICAL: Ensures the name is NEVER blank!
  };
  
  if (empSheet) {
    var lastRow = empSheet.getLastRow();
    if (lastRow > 1) {
      var data = empSheet.getRange(2, 1, lastRow - 1, 7).getValues();
      for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var emailInSheet = (row[2] || '').toString().trim().toLowerCase();
        
        if (emailInSheet && emailInSheet === userEmail.toLowerCase()) {
          userInfo.employeeId = row[0]; // Column A: Employee_ID
          userInfo.name = row[1] || userEmail;       // Column B: Full_Name
          var roleInSheet = (row[4] || '').toString().trim().toLowerCase(); // Column E: Role
          if (roleInSheet.indexOf('owner') > -1 || roleInSheet.indexOf('executive') > -1) {
              userInfo.role = 'owner';
          } else if (roleInSheet.indexOf('admin') > -1) {
              userInfo.role = 'admin';
          } else {
              userInfo.role = 'employee';
          }
          break;
        }
      }
    }
  }
  
  return userInfo;
}

// 2.5 FETCH DYNAMIC DROPDOWNS
// This pulls live data from your "Properties" and "Tasks" sheets for the autocomplete boxes.
function getDropdownData() {
  var ss = getSpreadsheet();
  
  var propSheet = ss.getSheetByName('Properties');
  var taskSheet = ss.getSheetByName('Tasks');
  
  var properties = [];
  var tasks = [];
  
  // Try to load properties (Assuming names are in Column A, starting row 2)
  if (propSheet) {
    var lastRow = propSheet.getLastRow();
    if (lastRow > 1) {
      var propData = propSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      // Flatten the 2D array and remove empty rows
      properties = propData.map(function(row) { return row[0]; }).filter(String);
    }
  } else {
    // Fallback default list if the 'Properties' sheet hasn't been created yet
    properties = ["Alpha Phi", "4638 B", "645 Ber", "Company Office"];
  }
  
  // Try to load tasks (Assuming names are in Column A, starting row 2)
  if (taskSheet) {
    var lastTaskRow = taskSheet.getLastRow();
    if (lastTaskRow > 1) {
      var taskData = taskSheet.getRange(2, 1, lastTaskRow - 1, 1).getValues();
      tasks = taskData.map(function(row) { return row[0]; }).filter(String);
    }
  } else {
    // Fallback default list if the 'Tasks' sheet hasn't been created yet
    tasks = ["Handyman", "Plumbing", "Painting", "Lawn Care", "Materials Run"];
  }
  
  return { properties: properties, tasks: tasks };
}

// Helper to get or create Google Drive folder for receipts
function getOrCreateReceiptsFolder() {
  // [ACTION_REQUIRED]: Update with preferred folder name if different
  var folderName = "BPM Material Receipts";
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty('RECEIPTS_FOLDER_ID');

  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (e) {
      // Folder might have been deleted, proceed to search by name
    }
  }

  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    var folder = folders.next();
    props.setProperty('RECEIPTS_FOLDER_ID', folder.getId());
    return folder;
  }

  var newFolder = DriveApp.createFolder(folderName);
  props.setProperty('RECEIPTS_FOLDER_ID', newFolder.getId());
  return newFolder;
}

// 3. SUBMIT WORK LOG TO SHEETS
// This is called from the HTML file via google.script.run.submitWorkLog(payload)
function submitWorkLog(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = getSpreadsheet();
    
    // Generate a unique ID for this specific log entry
    var workLogId = Utilities.getUuid();
    var timestamp = new Date();
    
    // Dynamically look up active user's Employee ID
    var activeUser = getUserRole();
    var employeeId = activeUser.employeeId || "EMP-001";
    
    // --- WRITE TO LABOR LOGS SHEET ---
    var laborSheet = ss.getSheetByName('fact_Work_Logs');
    if (!laborSheet) throw new Error("Could not find sheet: fact_Work_Logs");
    
    // Append the main labor entry
    laborSheet.appendRow([
      workLogId,              // Work_Log_ID
      timestamp,              // Timestamp
      payload.date,           // Date_Completed
      employeeId,             // Employee_ID
      payload.propertyId,     // Property_ID
      payload.taskId,         // Task_ID
      payload.hours,          // Hours_Worked
      payload.notes,          // Work_Notes
      "Pending",              // Payroll_Status (Approval workflow starts as Pending)
      "Unbilled",             // Billing_Status
      payload.chargeTarget || "Owner" // Charge_Target
    ]);
    
    // --- WRITE TO MATERIALS SHEET (If materials were added) ---
    if (payload.materials && payload.materials.length > 0) {
      var materialSheet = ss.getSheetByName('fact_Material_Expenses');
      if (!materialSheet) throw new Error("Could not find sheet: fact_Material_Expenses");
      
      var receiptsFolder = null;

      for (var i = 0; i < payload.materials.length; i++) {
        var item = payload.materials[i];
        var receiptUrl = "";

        // Handle File Upload to Drive if provided
        if (item.fileData && item.fileName) {
          try {
            if (!receiptsFolder) {
              receiptsFolder = getOrCreateReceiptsFolder();
            }
            var base64Content = item.fileData;
            if (base64Content.indexOf(',') > -1) {
              base64Content = base64Content.split(',')[1];
            }
            var bytes = Utilities.base64Decode(base64Content);
            var blob = Utilities.newBlob(bytes, item.fileType || 'application/octet-stream', item.fileName);
            var driveFile = receiptsFolder.createFile(blob);
            driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            receiptUrl = driveFile.getUrl();
          } catch (fileErr) {
            console.error("Error saving receipt photo: " + fileErr.message);
          }
        }

        materialSheet.appendRow([
          Utilities.getUuid(),    // Expense_ID
          workLogId,              // Work_Log_ID
          payload.propertyId,     // Property_ID
          item.vendor,            // Vendor_Name
          item.description,       // Item_Description
          item.cost,              // Cost
          receiptUrl              // Receipt_URL
        ]);
      }
    }
    
    return { success: true, message: "Log saved successfully" };
    
  } catch (error) {
    return { success: false, error: error.message };
  } finally {
    lock.releaseLock();
  }
}

// 13. SETTLEMENT CALCULATIONS

function generateOwnerSettlement(propertyId, monthYear, managementFeePercent) {
  try {
    var ss = getSpreadsheet();
    if (managementFeePercent === undefined) managementFeePercent = 8.0;

    // Determine start and end dates from monthYear ("YYYY-MM")
    var parts = monthYear.split('-');
    var year = parseInt(parts[0], 10);
    var month = parseInt(parts[1], 10) - 1; // 0-indexed month

    var startDate = new Date(year, month, 1);
    var endDate = new Date(year, month + 1, 0, 23, 59, 59); // Last day of month

    var startDateStr = Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var endDateStr = Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    // 1. Get Leases for Property
    var leaseSheet = ss.getSheetByName('fact_Leases');
    var propertyLeaseIds = {};
    var leaseManagementFees = {};
    if (leaseSheet && leaseSheet.getLastRow() > 1) {
      var leaseData = leaseSheet.getRange(2, 1, leaseSheet.getLastRow() - 1, leaseSheet.getLastColumn()).getValues();
      for (var i = 0; i < leaseData.length; i++) {
        if (leaseData[i][1] === propertyId) {
          propertyLeaseIds[leaseData[i][0]] = true;
          var mFee = parseFloat(leaseData[i][7]);
          leaseManagementFees[leaseData[i][0]] = !isNaN(mFee) ? mFee : managementFeePercent;
        }
      }
    }

    // 2. Total Rent Collected & 3. Management Fee Calculation
    var ledgerSheet = ss.getSheetByName('fact_Rent_Ledger');
    var totalRentCollected = 0;
    var managementFee = 0;
    var rentLedgerItems = [];
    if (ledgerSheet && ledgerSheet.getLastRow() > 1) {
      var ledgerData = ledgerSheet.getRange(2, 1, ledgerSheet.getLastRow() - 1, 7).getValues();
      for (var j = 0; j < ledgerData.length; j++) {
        var l_leaseId = ledgerData[j][1];
        var l_date = ledgerData[j][3];
        var l_type = ledgerData[j][4];
        var l_payment = parseFloat(ledgerData[j][6]) || 0;

        if (propertyLeaseIds[l_leaseId] && l_type === 'Rent' && l_payment > 0) {
          var logDate = new Date(l_date);
          if (logDate >= startDate && logDate <= endDate) {
            totalRentCollected += l_payment;
            var feePercent = leaseManagementFees[l_leaseId] !== undefined ? leaseManagementFees[l_leaseId] : managementFeePercent;
            managementFee += l_payment * (feePercent / 100);
            rentLedgerItems.push({
              date: Utilities.formatDate(logDate, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
              amount: l_payment,
              leaseId: l_leaseId,
              feePercent: feePercent
            });
          }
        }
      }
    }

    // 4. Maintenance Billed to Owner
    var adminData = getAdminPayrollData(startDateStr, endDateStr, 'All');
    if (!adminData.success) throw new Error(adminData.error);

    var filteredLogs = adminData.workLogs.filter(function(log) {
      return log.propertyId === propertyId && log.chargeTarget === 'Owner';
    });

    var ownerWorkLogIds = {};
    filteredLogs.forEach(function(log) { ownerWorkLogIds[log.workLogId] = true; });

    var filteredMaterials = adminData.materials.filter(function(mat) {
      if (mat.propertyId !== propertyId) return false;
      if (mat.workLogId && !ownerWorkLogIds[mat.workLogId]) return false;
      return true;
    });

    var totalLaborCost = filteredLogs.reduce(function(sum, log) {
      return sum + (log.billableAmount !== undefined && log.billableAmount > 0 ? log.billableAmount : log.grossPay);
    }, 0);
    var totalMaterialsCost = filteredMaterials.reduce(function(sum, mat) { return sum + mat.cost; }, 0);
    var totalMaintenance = totalLaborCost + totalMaterialsCost;

    // 5. Net Payout
    var netPayout = totalRentCollected - managementFee - totalMaintenance;

    return {
      success: true,
      settlement: {
        propertyId: propertyId,
        monthYear: monthYear,
        totalRentCollected: totalRentCollected,
        rentLedgerItems: rentLedgerItems,
        managementFeePercent: managementFeePercent,
        managementFeeAmount: managementFee,
        laborItems: filteredLogs,
        materialItems: filteredMaterials,
        totalMaintenanceBilled: totalMaintenance,
        netPayout: netPayout
      }
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function generateSecurityDepositSettlement(leaseId) {
  try {
    var ss = getSpreadsheet();

    // 1. Get Lease Information
    var leaseSheet = ss.getSheetByName('fact_Leases');
    if (!leaseSheet) throw new Error("Could not find sheet: fact_Leases");

    var leaseData = leaseSheet.getRange(2, 1, leaseSheet.getLastRow() - 1, 7).getValues();
    var lease = null;
    for (var i = 0; i < leaseData.length; i++) {
      if (leaseData[i][0] === leaseId) {
        lease = {
          propertyId: leaseData[i][1],
          startDate: new Date(leaseData[i][2]),
          endDate: new Date(leaseData[i][3]),
          securityDeposit: parseFloat(leaseData[i][5]) || 0
        };
        break;
      }
    }

    if (!lease) throw new Error("Lease not found for ID: " + leaseId);

    // 2. Unpaid Rent/Late Fees
    var balanceData = calculateLeaseBalance(leaseId);
    if (!balanceData.success) throw new Error("Error calculating balance: " + balanceData.error);

    var rawUnpaidRent = balanceData.balance - (balanceData.lateFeesIncurred || 0);
    var unpaidRent = Math.max(0, rawUnpaidRent);
    var lateFeesIncurred = balanceData.lateFeesIncurred || 0;

    // 3. Damages / Maintenance (Target = Security Deposit)
    var extendedEndDate = new Date(lease.endDate);
    extendedEndDate.setDate(extendedEndDate.getDate() + 30); // 30 days after lease

    var startDateStr = Utilities.formatDate(lease.startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var endDateStr = Utilities.formatDate(extendedEndDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    var adminData = getAdminPayrollData(startDateStr, endDateStr, 'All');
    if (!adminData.success) throw new Error(adminData.error);

    var filteredLogs = adminData.workLogs.filter(function(log) {
      return log.propertyId === lease.propertyId && log.chargeTarget === 'Security Deposit';
    });

    var depositWorkLogIds = {};
    filteredLogs.forEach(function(log) { depositWorkLogIds[log.workLogId] = true; });

    var filteredMaterials = adminData.materials.filter(function(mat) {
      if (mat.propertyId !== lease.propertyId) return false;
      if (mat.workLogId && !depositWorkLogIds[mat.workLogId]) return false;
      return true;
    });

    var totalLaborCost = filteredLogs.reduce(function(sum, log) {
      return sum + (log.billableAmount !== undefined && log.billableAmount > 0 ? log.billableAmount : log.grossPay);
    }, 0);
    var totalMaterialsCost = filteredMaterials.reduce(function(sum, mat) { return sum + mat.cost; }, 0);
    var totalDamages = totalLaborCost + totalMaterialsCost;

    // 4. Net Refund
    var netRefund = lease.securityDeposit - unpaidRent - totalDamages - lateFeesIncurred;

    return {
      success: true,
      settlement: {
        leaseId: leaseId,
        propertyId: lease.propertyId,
        securityDeposit: lease.securityDeposit,
        unpaidRent: unpaidRent,
        lateFeesIncurred: lateFeesIncurred,
        laborItems: filteredLogs,
        materialItems: filteredMaterials,
        totalDamages: totalDamages,
        netRefund: netRefund
      }
    };

  } catch (err) {
    return { success: false, error: err.message };
  }
}

// 3.5 GET EMPLOYEE WORK LOGS & DAILY TOTALS
function exportOwnerSettlementPDF(propertyId, monthYear, managementFeePercent) {
  try {
    var report = generateOwnerSettlement(propertyId, monthYear, managementFeePercent);
    if (!report.success) throw new Error(report.error);
    var st = report.settlement;

    var docTitle = "Owner Settlement - " + propertyId + " - " + monthYear;
    var doc = DocumentApp.create(docTitle);
    var body = doc.getBody();

    var titlePara = body.appendParagraph("MONTHLY OWNER SETTLEMENT");
    titlePara.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    titlePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    body.appendParagraph("Property: " + st.propertyId);
    body.appendParagraph("Month/Year: " + st.monthYear);
    body.appendParagraph("Generated: " + new Date().toLocaleDateString());
    body.appendParagraph("");

    var rentHeader = body.appendParagraph("Rent Collected");
    rentHeader.setHeading(DocumentApp.ParagraphHeading.HEADING2);

    if (st.rentLedgerItems && st.rentLedgerItems.length > 0) {
      var rentTable = [["Date", "Lease ID", "Amount"]];
      st.rentLedgerItems.forEach(function(item) {
        rentTable.push([item.date, item.leaseId, "$" + item.amount.toFixed(2)]);
      });
      body.appendTable(rentTable);
    } else {
      body.appendParagraph("No rent collected this period.");
    }
    body.appendParagraph("Total Rent Collected: $" + st.totalRentCollected.toFixed(2));
    body.appendParagraph("");

    var feeHeader = body.appendParagraph("Management Fee");
    feeHeader.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    if (st.rentLedgerItems && st.rentLedgerItems.length > 0) {
      body.appendParagraph("Fee Amount: $" + st.managementFeeAmount.toFixed(2) + " (Calculated per lease)");
    } else {
      body.appendParagraph("Fee Amount: $" + st.managementFeeAmount.toFixed(2));
    }
    body.appendParagraph("");

    var maintHeader = body.appendParagraph("Maintenance & Repairs");
    maintHeader.setHeading(DocumentApp.ParagraphHeading.HEADING2);

    if (st.laborItems.length > 0 || st.materialItems.length > 0) {
       body.appendParagraph("Labor:");
       st.laborItems.forEach(function(log) {
         var amt = (log.billableAmount !== undefined && log.billableAmount > 0) ? log.billableAmount : log.grossPay;
         body.appendParagraph("- " + log.dateCompleted + ": " + log.taskId + " (" + log.hoursWorked + " hrs) = $" + amt.toFixed(2));
       });
       body.appendParagraph("Materials:");
       st.materialItems.forEach(function(mat) {
         body.appendParagraph("- " + mat.vendor + ": " + mat.description + " = $" + mat.cost.toFixed(2));
       });
    } else {
      body.appendParagraph("No maintenance billed this period.");
    }
    body.appendParagraph("Total Maintenance: $" + st.totalMaintenanceBilled.toFixed(2));
    body.appendParagraph("");

    var payoutPara = body.appendParagraph("NET PAYOUT TO OWNER: $" + st.netPayout.toFixed(2));
    payoutPara.setHeading(DocumentApp.ParagraphHeading.HEADING3);

    doc.saveAndClose();

    var docFile = DriveApp.getFileById(doc.getId());
    var pdfBlob = docFile.getAs('application/pdf');
    pdfBlob.setName(docTitle + ".pdf");
    var pdfFile = DriveApp.createFile(pdfBlob);
    var base64Pdf = Utilities.base64Encode(pdfBlob.getBytes());
    docFile.setTrashed(true);

    return {
      success: true,
      docUrl: docFile.getUrl(),
      pdfUrl: pdfFile.getUrl(),
      fileName: docTitle + ".pdf",
      base64Pdf: base64Pdf
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function exportSecurityDepositSettlementPDF(leaseId, waivedLateFees) {
  try {
    var report = generateSecurityDepositSettlement(leaseId);
    if (!report.success) throw new Error(report.error);
    var st = report.settlement;

    var waived = parseFloat(waivedLateFees) || 0;
    if (waived > st.lateFeesIncurred) {
        waived = st.lateFeesIncurred;
    }
    var netLateFees = st.lateFeesIncurred - waived;
    var finalNetRefund = st.securityDeposit - st.unpaidRent - netLateFees - st.totalDamages;

    var docTitle = "Security Deposit Settlement - " + leaseId;
    var doc = DocumentApp.create(docTitle);
    var body = doc.getBody();

    var titlePara = body.appendParagraph("SECURITY DEPOSIT SETTLEMENT");
    titlePara.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    titlePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    body.appendParagraph("Lease ID: " + st.leaseId);
    body.appendParagraph("Property: " + st.propertyId);
    body.appendParagraph("Generated: " + new Date().toLocaleDateString());
    body.appendParagraph("");

    var depHeader = body.appendParagraph("Security Deposit");
    depHeader.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph("Initial Deposit Held: $" + st.securityDeposit.toFixed(2));
    body.appendParagraph("");

    var rentHeader = body.appendParagraph("Unpaid Rent & Late Fees");
    rentHeader.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph("Unpaid Rent: $" + st.unpaidRent.toFixed(2));
    body.appendParagraph("Late Fees Incurred: $" + st.lateFeesIncurred.toFixed(2));
    if (waived > 0) {
        body.appendParagraph("Late Fees Waived (Courtesy): -$" + waived.toFixed(2));
    }
    body.appendParagraph("Net Late Fees Charged: $" + netLateFees.toFixed(2));
    body.appendParagraph("");

    var maintHeader = body.appendParagraph("Damages & Maintenance Deductions");
    maintHeader.setHeading(DocumentApp.ParagraphHeading.HEADING2);

    if (st.laborItems.length > 0 || st.materialItems.length > 0) {
       body.appendParagraph("Labor Deductions:");
       st.laborItems.forEach(function(log) {
         var amt = (log.billableAmount !== undefined && log.billableAmount > 0) ? log.billableAmount : log.grossPay;
         body.appendParagraph("- " + log.dateCompleted + ": " + log.taskId + " (" + log.workNotes + ") = $" + amt.toFixed(2));
       });
       body.appendParagraph("Material Deductions:");
       st.materialItems.forEach(function(mat) {
         body.appendParagraph("- " + mat.vendor + ": " + mat.description + " = $" + mat.cost.toFixed(2));
       });
    } else {
      body.appendParagraph("No damages recorded.");
    }
    body.appendParagraph("Total Damages Deducted: $" + st.totalDamages.toFixed(2));
    body.appendParagraph("");

    var refPara = body.appendParagraph("NET REFUND TO TENANT: $" + finalNetRefund.toFixed(2));
    refPara.setHeading(DocumentApp.ParagraphHeading.HEADING3);

    doc.saveAndClose();

    var docFile = DriveApp.getFileById(doc.getId());
    var pdfBlob = docFile.getAs('application/pdf');
    pdfBlob.setName(docTitle + ".pdf");
    var pdfFile = DriveApp.createFile(pdfBlob);
    var base64Pdf = Utilities.base64Encode(pdfBlob.getBytes());
    docFile.setTrashed(true);

    return {
      success: true,
      docUrl: docFile.getUrl(),
      pdfUrl: pdfFile.getUrl(),
      fileName: docTitle + ".pdf",
      base64Pdf: base64Pdf
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// 3.5 GET EMPLOYEE WORK LOGS & DAILY TOTALS
function getEmployeeWorkLogs() {
  try {
    var ss = getSpreadsheet();
    var activeUser = getUserRole();
    var employeeId = activeUser.employeeId || "EMP-001";

    var laborSheet = ss.getSheetByName('fact_Work_Logs');
    var punches = [];
    var dailyTotals = {};

    if (laborSheet && laborSheet.getLastRow() > 1) {
      var laborData = laborSheet.getRange(2, 1, laborSheet.getLastRow() - 1, 11).getValues();
      for (var i = 0; i < laborData.length; i++) {
        var row = laborData[i];
        var empId = (row[3] || '').toString().trim();

        if (empId === employeeId) {
          var rawDate = row[2];
          var dateStr = '';
          if (rawDate instanceof Date) {
            dateStr = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          } else if (rawDate) {
            dateStr = row[2].toString().trim();
          }

          var hours = parseFloat(row[6]) || 0;
          var status = (row[8] || 'Pending').toString().trim();

          punches.push({
            workLogId: row[0],
            timestamp: row[1] ? new Date(row[1]).toISOString() : '',
            dateCompleted: dateStr,
            propertyId: row[4],
            taskId: row[5],
            hoursWorked: hours,
            workNotes: row[7],
            payrollStatus: status,
            billingStatus: (row[9] || 'Unbilled').toString().trim()
          });

          if (dateStr) {
            if (!dailyTotals[dateStr]) {
              dailyTotals[dateStr] = 0;
            }
            dailyTotals[dateStr] += hours;
          }
        }
      }
    }

    punches.sort(function(a, b) {
      return new Date(b.dateCompleted) - new Date(a.dateCompleted);
    });

    var dailySummary = Object.keys(dailyTotals).map(function(d) {
      return { date: d, totalHours: dailyTotals[d] };
    }).sort(function(a, b) {
      return new Date(b.date) - new Date(a.date);
    });

    return {
      success: true,
      punches: punches,
      dailySummary: dailySummary,
      employeeName: activeUser.name || employeeId
    };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// 3.6 UPDATE EXISTING WORK LOG
function updateWorkLog(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // Wait up to 10 seconds for other processes to finish
    var ss = getSpreadsheet();
    var laborSheet = ss.getSheetByName('fact_Work_Logs');
    if (!laborSheet) throw new Error("Could not find sheet: fact_Work_Logs");

    var lastRow = laborSheet.getLastRow();
    if (lastRow <= 1) throw new Error("No work logs found to update.");

    var range = laborSheet.getRange(2, 1, lastRow - 1, 11);
    var values = range.getValues();

    var updated = false;
    var rowToUpdate = -1;
    for (var i = 0; i < values.length; i++) {
      var id = (values[i][0] || '').toString().trim();
      if (id === payload.workLogId) {
        var payrollStatus = (values[i][8] || '').toString().trim().toLowerCase();
        if (payrollStatus === 'paid') {
          throw new Error("Cannot edit a punch that has already been marked as Paid.");
        }

        rowToUpdate = i + 2;
        break;
      }
    }

    if (rowToUpdate === -1) throw new Error("Work log entry not found.");

    // Update only the specific row to prevent lost updates
    var updatedRowRange = laborSheet.getRange(rowToUpdate, 3, 1, 6); // Col 3 (Date) to Col 8 (Notes)
    var updatedRowData = [[
      payload.date,       // Date_Completed
      values[rowToUpdate - 2][3], // Preserve Employee_ID (Col 4, index 3 in values)
      payload.propertyId, // Property_ID
      payload.taskId,     // Task_ID
      payload.hours,      // Hours_Worked
      payload.notes       // Work_Notes
    ]];
    updatedRowRange.setValues(updatedRowData);

    if (payload.chargeTarget !== undefined) {
      laborSheet.getRange(rowToUpdate, 11, 1, 1).setValue(payload.chargeTarget);
    }

    return { success: true, message: "Punch updated successfully" };
  } catch (error) {
    return { success: false, error: error.message };
  } finally {
    lock.releaseLock();
  }
}

// 3.7 WORK LOG APPROVAL WORKFLOW FUNCTION
function updateApprovalStatus(workLogId, newStatus) {
  var roleCheck = getUserRole();
  if (roleCheck.role !== 'admin' && roleCheck.role !== 'owner') {
    return { success: false, error: "Unauthorized access." };
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = getSpreadsheet();
    var laborSheet = ss.getSheetByName('fact_Work_Logs');
    if (!laborSheet) throw new Error("Could not find sheet: fact_Work_Logs");

    var lastRow = laborSheet.getLastRow();
    if (lastRow <= 1) throw new Error("No work logs found.");

    var range = laborSheet.getRange(2, 1, lastRow - 1, 1); // Only fetch the ID column for faster search
    var values = range.getValues();

    var rowToUpdate = -1;
    for (var i = 0; i < values.length; i++) {
      var id = (values[i][0] || '').toString().trim();
      if (id === workLogId) {
        rowToUpdate = i + 2;
        break;
      }
    }

    if (rowToUpdate === -1) throw new Error("Work log entry not found.");

    // Update only the specific cell
    laborSheet.getRange(rowToUpdate, 9).setValue(newStatus); // Column I: Payroll_Status / Approval Status

    return { success: true, message: "Work log status updated to " + newStatus };
  } catch (error) {
    return { success: false, error: error.message };
  } finally {
    lock.releaseLock();
  }
}

// 3.8 UPDATE CUSTOM BILLABLE AMOUNT
function updateChargeAmount(workLogId, amount) {
  var roleCheck = getUserRole();
  if (roleCheck.role !== 'admin' && roleCheck.role !== 'owner') {
    return { success: false, error: "Unauthorized. Admin or Owner access required." };
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = getSpreadsheet();
    var laborSheet = ss.getSheetByName('fact_Work_Logs');
    if (!laborSheet) throw new Error("Could not find sheet: fact_Work_Logs");

    var lastRow = laborSheet.getLastRow();
    if (lastRow <= 1) throw new Error("No work logs found.");

    var range = laborSheet.getRange(2, 1, lastRow - 1, 1);
    var values = range.getValues();

    var rowToUpdate = -1;
    for (var i = 0; i < values.length; i++) {
      var id = (values[i][0] || '').toString().trim();
      if (id === workLogId) {
        rowToUpdate = i + 2;
        break;
      }
    }

    if (rowToUpdate === -1) {
      // Also check archive sheet just in case, though editing archived logs is typically not expected
      var archiveSheet = ss.getSheetByName('archive_fact_Work_Logs');
      if (archiveSheet) {
        var aLastRow = archiveSheet.getLastRow();
        if (aLastRow > 1) {
          var aRange = archiveSheet.getRange(2, 1, aLastRow - 1, 1);
          var aValues = aRange.getValues();
          for (var j = 0; j < aValues.length; j++) {
            var aId = (aValues[j][0] || '').toString().trim();
            if (aId === workLogId) {
              rowToUpdate = j + 2;
              laborSheet = archiveSheet;
              break;
            }
          }
        }
      }
    }

    if (rowToUpdate === -1) throw new Error("Work log entry not found.");

    // Update only the specific cell (Column L / 12)
    laborSheet.getRange(rowToUpdate, 12).setValue(amount !== null && amount !== undefined && amount !== '' ? amount : '');

    return { success: true, message: "Billable amount updated successfully." };
  } catch (error) {
    return { success: false, error: error.message };
  } finally {
    lock.releaseLock();
  }
}

// 4. ADMIN PAYROLL & BILLING CALCULATIONS WITH DATE FILTERING
// Fetches records from fact_Work_Logs, dim_Employees, and fact_Material_Expenses with optional start/end date filtering.
function getAdminPayrollData(startDateStr, endDateStr, statusFilter) {
  try {
    var ss = getSpreadsheet();
    
    var startDate = parseLocalDate(startDateStr, false);
    var endDate = parseLocalDate(endDateStr, true);

    var cutoffDate = new Date();
    cutoffDate.setMonth(cutoffDate.getMonth() - 3);

    var pullArchive = false;
    if (!startDate || startDate < cutoffDate) {
      pullArchive = true;
    }

    // Build Employee Lookup Dictionary
    var empSheet = ss.getSheetByName('dim_Employees');
    var employeeMap = {};
    if (empSheet && empSheet.getLastRow() > 1) {
      var empData = empSheet.getRange(2, 1, empSheet.getLastRow() - 1, 7).getValues();
      empData.forEach(function(row) {
        var empId = (row[0] || '').toString().trim();
        if (empId) {
          employeeMap[empId] = {
            id: empId,
            name: row[1] || 'Unknown Employee',
            email: row[2] || '',
            payRate: parseFloat(row[3]) || 0,
            role: row[4] || 'Field Crew',
            status: row[6] || 'Active'
          };
        }
      });
    }

    // Build Billing Rates Dictionary
    var ratesSheet = ss.getSheetByName('dim_Billing_Rates');
    var rateMap = {};
    if (ratesSheet && ratesSheet.getLastRow() > 1) {
      var ratesData = ratesSheet.getRange(2, 1, ratesSheet.getLastRow() - 1, 6).getValues();
      ratesData.forEach(function(row) {
        var taskId = (row[1] || '').toString().trim();
        var empId = (row[3] || '').toString().trim();
        var rate = parseFloat(row[5]) || 0;
        if (taskId && empId) {
          var key = taskId + '_' + empId;
          rateMap[key] = rate;
        }
      });
    }

    // Read Work Logs
    var laborSheet = ss.getSheetByName('fact_Work_Logs');
    var archiveLaborSheet = pullArchive ? ss.getSheetByName('archive_fact_Work_Logs') : null;

    var workLogs = [];
    var laborData = [];

    if (laborSheet && laborSheet.getLastRow() > 1) {
      laborData = laborData.concat(laborSheet.getRange(2, 1, laborSheet.getLastRow() - 1, 12).getValues());
    }

    if (archiveLaborSheet && archiveLaborSheet.getLastRow() > 1) {
      laborData = laborData.concat(archiveLaborSheet.getRange(2, 1, archiveLaborSheet.getLastRow() - 1, 12).getValues());
    }

    if (laborData.length > 0) {
      laborData.forEach(function(row) {
        var empId = (row[3] || '').toString().trim();
        var taskId = (row[5] || '').toString().trim();
        var empInfo = employeeMap[empId] || { name: empId || 'Unassigned', payRate: 0 };
        var hours = parseFloat(row[6]) || 0;
        var payRate = empInfo.payRate;
        var grossPay = hours * payRate;

        // Calculate Billable Amount
        var billableRate = 0;
        if (rateMap[taskId + '_' + empId] !== undefined) {
          billableRate = rateMap[taskId + '_' + empId];
        } else if (rateMap[taskId + '_DEFAULT'] !== undefined) {
          billableRate = rateMap[taskId + '_DEFAULT'];
        }
        var originalBillableAmount = hours * billableRate;

        var customBillableAmountRaw = row[11];
        var billableAmount = originalBillableAmount;
        if (customBillableAmountRaw !== undefined && customBillableAmountRaw !== null && customBillableAmountRaw !== '') {
          var parsedCustom = parseFloat(customBillableAmountRaw);
          if (!isNaN(parsedCustom)) {
            billableAmount = parsedCustom;
          }
        }

        var rawDate = row[2];
        var logDate = rawDate ? new Date(rawDate) : null;
        var dateFormatted = logDate ? Utilities.formatDate(logDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';

        // Apply Date Range Filter
        if (startDate && logDate && logDate < startDate) return;
        if (endDate && logDate && logDate > endDate) return;

        var status = (row[8] || 'Pending').toString().trim();
        if (statusFilter && statusFilter !== 'All' && status.toLowerCase() !== statusFilter.toLowerCase()) return;

        workLogs.push({
          workLogId: row[0],
          timestamp: row[1] ? new Date(row[1]).toISOString() : '',
          dateCompleted: dateFormatted || (row[2] ? row[2].toString() : ''),
          employeeId: empId,
          employeeName: empInfo.name,
          payRate: payRate,
          propertyId: row[4],
          taskId: row[5],
          hoursWorked: hours,
          workNotes: row[7],
          payrollStatus: status,
          billingStatus: (row[9] || 'Unbilled').toString().trim(),
          chargeTarget: (row[10] || 'Owner').toString().trim(),
          grossPay: grossPay,
          originalBillableAmount: originalBillableAmount,
          billableAmount: billableAmount
        });
      });
    }

    // Read Material Expenses
    var matSheet = ss.getSheetByName('fact_Material_Expenses');
    var archiveMatSheet = pullArchive ? ss.getSheetByName('archive_fact_Material_Expenses') : null;

    var materials = [];
    var matData = [];

    if (matSheet && matSheet.getLastRow() > 1) {
      matData = matData.concat(matSheet.getRange(2, 1, matSheet.getLastRow() - 1, 7).getValues());
    }

    if (archiveMatSheet && archiveMatSheet.getLastRow() > 1) {
      matData = matData.concat(archiveMatSheet.getRange(2, 1, archiveMatSheet.getLastRow() - 1, 7).getValues());
    }

    if (matData.length > 0) {
      matData.forEach(function(row) {
        materials.push({
          expenseId: row[0],
          workLogId: row[1],
          propertyId: row[2],
          vendor: row[3],
          description: row[4],
          cost: parseFloat(row[5]) || 0,
          receiptUrl: row[6] || ''
        });
      });
    }

    // Aggregates for Payroll by Employee
    var employeeSummaries = {};
    Object.keys(employeeMap).forEach(function(id) {
      employeeSummaries[id] = {
        employeeId: id,
        name: employeeMap[id].name,
        payRate: employeeMap[id].payRate,
        totalHours: 0,
        pendingHours: 0,
        approvedHours: 0,
        totalGrossPay: 0,
        pendingGrossPay: 0,
        approvedGrossPay: 0,
        logCount: 0
      };
    });

    var pendingPayrollHours = 0;
    var pendingPayrollGrossPay = 0;
    var approvedPayrollHours = 0;
    var approvedPayrollGrossPay = 0;
    var totalUnbilledLaborCost = 0;
    var totalUnbilledMaterials = 0;

    workLogs.forEach(function(log) {
      if (!employeeSummaries[log.employeeId]) {
        employeeSummaries[log.employeeId] = {
          employeeId: log.employeeId,
          name: log.employeeName,
          payRate: log.payRate,
          totalHours: 0,
          pendingHours: 0,
          approvedHours: 0,
          totalGrossPay: 0,
          pendingGrossPay: 0,
          approvedGrossPay: 0,
          logCount: 0
        };
      }
      var emp = employeeSummaries[log.employeeId];
      emp.totalHours += log.hoursWorked;
      emp.totalGrossPay += log.grossPay;
      emp.logCount++;

      var statusLower = log.payrollStatus.toLowerCase();
      if (statusLower === 'pending' || statusLower === 'submitted') {
        emp.pendingHours += log.hoursWorked;
        emp.pendingGrossPay += log.grossPay;
        pendingPayrollHours += log.hoursWorked;
        pendingPayrollGrossPay += log.grossPay;
      } else if (statusLower === 'approved') {
        emp.approvedHours += log.hoursWorked;
        emp.approvedGrossPay += log.grossPay;
        approvedPayrollHours += log.hoursWorked;
        approvedPayrollGrossPay += log.grossPay;
      }

      if (log.billingStatus.toLowerCase() === 'unbilled') {
        totalUnbilledLaborCost += log.grossPay;
      }
    });

    // Filter materials to match filtered work logs
    var validWorkLogIds = {};
    workLogs.forEach(function(log) { validWorkLogIds[log.workLogId] = true; });
    
    var filteredMaterials = materials.filter(function(mat) {
      return !mat.workLogId || validWorkLogIds[mat.workLogId];
    });

    // Aggregates for Billing by Property
    var propertySummaries = {};
    workLogs.forEach(function(log) {
      var propId = log.propertyId || 'Unassigned';
      if (!propertySummaries[propId]) {
        propertySummaries[propId] = {
          propertyId: propId,
          unbilledHours: 0,
          unbilledLaborCost: 0,
          unbilledMaterialsCost: 0,
          totalUnbilled: 0,
          logCount: 0
        };
      }
      if (log.billingStatus.toLowerCase() === 'unbilled') {
        propertySummaries[propId].unbilledHours += log.hoursWorked;
        propertySummaries[propId].unbilledLaborCost += log.grossPay;
        propertySummaries[propId].logCount++;
      }
    });

    filteredMaterials.forEach(function(mat) {
      var propId = mat.propertyId || 'Unassigned';
      if (!propertySummaries[propId]) {
        propertySummaries[propId] = {
          propertyId: propId,
          unbilledHours: 0,
          unbilledLaborCost: 0,
          unbilledMaterialsCost: 0,
          totalUnbilled: 0,
          logCount: 0
        };
      }
      propertySummaries[propId].unbilledMaterialsCost += mat.cost;
      totalUnbilledMaterials += mat.cost;
    });

    Object.keys(propertySummaries).forEach(function(propId) {
      var p = propertySummaries[propId];
      p.totalUnbilled = p.unbilledLaborCost + p.unbilledMaterialsCost;
    });

    return {
      success: true,
      metrics: {
        pendingPayrollHours: pendingPayrollHours,
        pendingPayrollGrossPay: pendingPayrollGrossPay,
        approvedPayrollHours: approvedPayrollHours,
        approvedPayrollGrossPay: approvedPayrollGrossPay,
        unbilledMaterialsCost: totalUnbilledMaterials,
        unbilledLaborCost: totalUnbilledLaborCost,
        totalUnbilledInvoiceAmount: totalUnbilledLaborCost + totalUnbilledMaterials,
        totalLogs: workLogs.length
      },
      employeeSummaries: Object.values(employeeSummaries),
      propertySummaries: Object.values(propertySummaries),
      workLogs: workLogs,
      materials: filteredMaterials
    };

  } catch (error) {
    return { success: false, error: error.message };
  }
}

// 5. UPDATE PAYROLL STATUS TO PAID
function markPayrollPaid(payload) {
  var roleCheck = getUserRole();
  if (roleCheck.role !== 'admin' && roleCheck.role !== 'owner') {
    return { success: false, error: "Unauthorized access." };
  }

  var employeeId = payload.employeeId;
  var startDateStr = payload.startDate;
  var endDateStr = payload.endDate;
  var paymentMethod = payload.paymentMethod;
  var amount = payload.amount;
  var employeeName = payload.employeeName || "";

  var startDate = parseLocalDate(startDateStr, false);
  var endDate = parseLocalDate(endDateStr, true);

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = getSpreadsheet();
    var laborSheet = ss.getSheetByName('fact_Work_Logs');
    if (!laborSheet) throw new Error("Could not find sheet: fact_Work_Logs");

    var lastRow = laborSheet.getLastRow();
    var updated = false;

    if (lastRow > 1) {
      var fullDataRange = laborSheet.getRange(2, 1, lastRow - 1, laborSheet.getLastColumn());
      var fullDataValues = fullDataRange.getValues();
      var payrollStatusColIdx = 8; // 0-indexed column 9 (Payroll_Status)
      var empIdColIdx = 3;         // 0-indexed column 4 (Employee_ID)
      var dateColIdx = 2;          // 0-indexed column 3 (Date_Completed)

      for (var i = 0; i < fullDataValues.length; i++) {
        var row = fullDataValues[i];
        var empIdInRow = (row[empIdColIdx] || '').toString().trim();
        var status = (row[payrollStatusColIdx] || '').toString().trim().toLowerCase();

        var rawDate = row[dateColIdx];
        var logDate = rawDate ? new Date(rawDate) : null;

        // Check date bounds
        var inDateRange = true;
        if (startDate && logDate && logDate < startDate) inDateRange = false;
        if (endDate && logDate && logDate > endDate) inDateRange = false;

        if (inDateRange && (!employeeId || empIdInRow === employeeId) && (status === 'pending' || status === 'approved' || status === 'submitted')) {
          fullDataValues[i][payrollStatusColIdx] = 'Paid';
          updated = true;
        }
      }

      if (updated) {
        fullDataRange.setValues(fullDataValues);
      }
    }

    // Add to payroll history if anything was updated (or even if it wasn't, record the payment)
    var historySheet = ss.getSheetByName('fact_Payroll_History');
    if (!historySheet) {
        historySheet = ss.insertSheet('fact_Payroll_History');
        historySheet.appendRow([
            'Payroll_ID',
            'Employee_ID',
            'Employee_Name',
            'Start_Date',
            'End_Date',
            'Amount_Paid',
            'Payment_Method',
            'Timestamp'
        ]);
        historySheet.getRange(1, 1, 1, 8).setFontWeight('bold');
        historySheet.setFrozenRows(1);
    }

    var payrollId = "PAY-" + Utilities.getUuid().substring(0, 8).toUpperCase();
    historySheet.appendRow([
        payrollId,
        employeeId,
        employeeName,
        startDateStr,
        endDateStr,
        amount,
        paymentMethod,
        new Date()
    ]);

    return { success: true, message: 'Payroll marked as Paid successfully.' };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

// 5.5. GET PAYROLL HISTORY
function getPayrollHistory() {
    try {
        var ss = getSpreadsheet();
        var historySheet = ss.getSheetByName('fact_Payroll_History');

        if (!historySheet) {
            return { success: true, history: [] }; // Return empty if sheet doesn't exist yet
        }

        var lastRow = historySheet.getLastRow();
        if (lastRow <= 1) {
             return { success: true, history: [] };
        }

        var data = historySheet.getRange(2, 1, lastRow - 1, 8).getValues();
        var history = data.map(function(row) {
            return {
                payrollId: row[0],
                employeeId: row[1],
                employeeName: row[2],
                startDate: row[3],
                endDate: row[4],
                amountPaid: parseFloat(row[5]) || 0,
                paymentMethod: row[6],
                timestamp: row[7] ? new Date(row[7]).toISOString() : ''
            };
        });

        // Sort by timestamp descending
        history.sort(function(a, b) {
            return new Date(b.timestamp) - new Date(a.timestamp);
        });

        return { success: true, history: history };
    } catch (err) {
        return { success: false, error: err.message };
    }
}

// 6. GENERATE ITEMIZED INVOICE REPORT
function generateInvoiceReport(propertyId) {
  try {
    var adminData = getAdminPayrollData();
    if (!adminData.success) throw new Error(adminData.error);

    var filteredLogs = adminData.workLogs.filter(function(log) {
      return log.propertyId === propertyId && log.billingStatus.toLowerCase() === 'unbilled';
    });

    var filteredMaterials = adminData.materials.filter(function(mat) {
      return mat.propertyId === propertyId;
    });

    var totalLabor = filteredLogs.reduce(function(sum, log) { return sum + log.grossPay; }, 0);
    var totalHours = filteredLogs.reduce(function(sum, log) { return sum + log.hoursWorked; }, 0);
    var totalMaterials = filteredMaterials.reduce(function(sum, mat) { return sum + mat.cost; }, 0);

    return {
      success: true,
      invoice: {
        propertyId: propertyId,
        invoiceDate: new Date().toLocaleDateString(),
        laborItems: filteredLogs,
        materialItems: filteredMaterials,
        totalHours: totalHours,
        totalLaborCost: totalLabor,
        totalMaterialsCost: totalMaterials,
        grandTotal: totalLabor + totalMaterials
      }
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// 7. MARK PROPERTY INVOICE AS BILLED
function markPropertyBilled(propertyId) {
  var roleCheck = getUserRole();
  if (roleCheck.role !== 'admin' && roleCheck.role !== 'owner') {
    return { success: false, error: "Unauthorized access." };
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = getSpreadsheet();
    var laborSheet = ss.getSheetByName('fact_Work_Logs');
    
    if (laborSheet && laborSheet.getLastRow() > 1) {
      var lastRow = laborSheet.getLastRow();

      var range = laborSheet.getRange(2, 10, lastRow - 1, 1); // Get only the Billing_Status column
      var values = range.getValues();

      var propRange = laborSheet.getRange(2, 5, lastRow - 1, 1); // Get Property_ID column
      var propValues = propRange.getValues();

      var updated = false;

      for (var i = 0; i < values.length; i++) {
        var propIdInRow = (propValues[i][0] || '').toString().trim();
        var status = (values[i][0] || '').toString().trim().toLowerCase();

        if (propIdInRow === propertyId && status === 'unbilled') {
          values[i][0] = 'Billed';
          updated = true;
        }
      }

      // Perform one bulk update on the single column
      if (updated) {
        range.setValues(values);
      }
    }
    return { success: true, message: 'Property billing status updated to Billed.' };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

// 8. EXPORT INVOICE TO GOOGLE DOC & PDF
function exportInvoicePDF(propertyId) {
  try {
    var report = generateInvoiceReport(propertyId);
    if (!report.success) throw new Error(report.error);
    var inv = report.invoice;

    var sanitizedDate = inv.invoiceDate.replace(/\//g, '-');
    var docTitle = "Invoice - " + propertyId + " - " + sanitizedDate;
    
    // Create Google Doc
    var doc = DocumentApp.create(docTitle);
    var body = doc.getBody();

    // Document Title
    var titlePara = body.appendParagraph("BPM MAINTENANCE INVOICE");
    titlePara.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    titlePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    body.appendParagraph("Property: " + inv.propertyId);
    body.appendParagraph("Invoice Date: " + inv.invoiceDate);
    body.appendParagraph(""); // Blank spacer

    // Section 1: Labor Table
    var laborHeader = body.appendParagraph("Itemized Labor Entries");
    laborHeader.setHeading(DocumentApp.ParagraphHeading.HEADING2);

    if (inv.laborItems && inv.laborItems.length > 0) {
      var laborTableData = [["Date", "Worker", "Task / Description", "Hours", "Cost"]];
      inv.laborItems.forEach(function(item) {
        laborTableData.push([
          item.dateCompleted,
          item.employeeName,
          item.taskId + " - " + item.workNotes,
          item.hoursWorked.toString(),
          "$" + item.grossPay.toFixed(2)
        ]);
      });
      body.appendTable(laborTableData);
    } else {
      body.appendParagraph("No unbilled labor entries.");
    }

    body.appendParagraph(""); // Blank spacer

    // Section 2: Material Expenses
    var matHeader = body.appendParagraph("Material Expenses");
    matHeader.setHeading(DocumentApp.ParagraphHeading.HEADING2);

    if (inv.materialItems && inv.materialItems.length > 0) {
      var matTableData = [["Vendor", "Description", "Cost", "Receipt File"]];
      inv.materialItems.forEach(function(mat) {
        matTableData.push([
          mat.vendor,
          mat.description,
          "$" + mat.cost.toFixed(2),
          mat.receiptUrl ? mat.receiptUrl : "N/A"
        ]);
      });
      body.appendTable(matTableData);
    } else {
      body.appendParagraph("No unbilled material expenses.");
    }

    body.appendParagraph(""); // Blank spacer

    // Section 3: Summary Totals
    var summaryHeader = body.appendParagraph("Summary Totals");
    summaryHeader.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    
    body.appendParagraph("Total Labor Cost: $" + inv.totalLaborCost.toFixed(2));
    body.appendParagraph("Total Materials Cost: $" + inv.totalMaterialsCost.toFixed(2));
    
    var grandTotalPara = body.appendParagraph("Grand Invoice Total: $" + inv.grandTotal.toFixed(2));
    grandTotalPara.setHeading(DocumentApp.ParagraphHeading.HEADING3);

    doc.saveAndClose();

    // Convert Doc to PDF Blob and Save File in Drive
    var docFile = DriveApp.getFileById(doc.getId());
    var pdfBlob = docFile.getAs('application/pdf');
    pdfBlob.setName(docTitle + ".pdf");
    var pdfFile = DriveApp.createFile(pdfBlob);

    // Encode PDF bytes to base64 for browser client download
    var base64Pdf = Utilities.base64Encode(pdfBlob.getBytes());

    // Trash the temporary doc
    docFile.setTrashed(true);

    return {
      success: true,
      docUrl: docFile.getUrl(), // Might be inaccessible since trashed, but leaving for legacy reasons
      pdfUrl: pdfFile.getUrl(),
      fileName: docTitle + ".pdf",
      base64Pdf: base64Pdf
    };

  } catch (err) {
    return { success: false, error: err.message };
  }
}

// 9. ADD NEW LEASE AND MULTIPLE TENANTS
function addLease(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = getSpreadsheet();

    var leaseSheet = ss.getSheetByName('fact_Leases');
    if (!leaseSheet) throw new Error("Could not find sheet: fact_Leases");

    var tenantSheet = ss.getSheetByName('dim_Tenants');
    if (!tenantSheet) throw new Error("Could not find sheet: dim_Tenants");

    var leaseId = "LSE-" + Utilities.getUuid().substring(0, 8).toUpperCase();

    // Add lease
    leaseSheet.appendRow([
      leaseId,
      payload.propertyId,
      payload.startDate,
      payload.endDate,
      payload.monthlyRent,
      payload.securityDeposit,
      "Active",
      payload.managementFeePercent !== undefined ? payload.managementFeePercent : 8.0,
      payload.lateFeePerDay !== undefined ? payload.lateFeePerDay : 5.0,
      payload.gracePeriodDays !== undefined ? payload.gracePeriodDays : 5
    ]);

    // Add tenants
    if (payload.tenants && payload.tenants.length > 0) {
      var rentPortion = payload.monthlyRent / payload.tenants.length;
      payload.tenants.forEach(function(tenant) {
        var tenantId = "TNT-" + Utilities.getUuid().substring(0, 8).toUpperCase();
        tenantSheet.appendRow([
          tenantId,
          leaseId,
          tenant.name,
          tenant.contactInfo,
          tenant.rentPortion ? tenant.rentPortion : rentPortion
        ]);
      });
    }

    return { success: true, message: "Lease and tenants added successfully", leaseId: leaseId };
  } catch (error) {
    return { success: false, error: error.message };
  } finally {
    lock.releaseLock();
  }
}

// 10. RECORD RENT LEDGER ENTRY (CHARGE OR PAYMENT)
function recordLedgerEntry(payload) {
  var roleCheck = getUserRole();
  if (roleCheck.role !== 'admin' && roleCheck.role !== 'owner') {
    return { success: false, error: "Unauthorized access." };
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = getSpreadsheet();

    var ledgerSheet = ss.getSheetByName('fact_Rent_Ledger');
    if (!ledgerSheet) throw new Error("Could not find sheet: fact_Rent_Ledger");

    var ledgerId = "LDG-" + Utilities.getUuid().substring(0, 8).toUpperCase();

    ledgerSheet.appendRow([
      ledgerId,
      payload.leaseId,
      payload.tenantId,
      payload.date,
      payload.transactionType,
      payload.chargeAmount || 0,
      payload.paymentAmount || 0,
      payload.paymentMode || "",
      payload.notes || "",
      payload.rentMonth !== undefined ? payload.rentMonth : "",
      payload.rentYear !== undefined ? payload.rentYear : ""
    ]);

    return { success: true, message: "Ledger entry recorded successfully" };
  } catch (error) {
    return { success: false, error: error.message };
  } finally {
    lock.releaseLock();
  }
}

// 11. FETCH LEASES AND THEIR TENANTS (For UI Dropdowns)
function getLeasesAndTenants() {
  try {
    var ss = getSpreadsheet();

    var leaseSheet = ss.getSheetByName('fact_Leases');
    var tenantSheet = ss.getSheetByName('dim_Tenants');

    var leases = [];
    var tenants = [];

    if (leaseSheet && leaseSheet.getLastRow() > 1) {
      var leaseData = leaseSheet.getRange(2, 1, leaseSheet.getLastRow() - 1, leaseSheet.getLastColumn()).getValues();
      leases = leaseData.map(function(row) {
        return {
          leaseId: row[0],
          propertyId: row[1],
          startDate: row[2],
          endDate: row[3],
          monthlyRent: row[4],
          securityDeposit: row[5],
          status: row[6],
          managementFeePercent: parseFloat(row[7]) || 8.0,
          lateFeePerDay: parseFloat(row[8]) || 5.0,
          gracePeriodDays: parseInt(row[9]) || 5
        };
      });
    }

    if (tenantSheet && tenantSheet.getLastRow() > 1) {
      var tenantData = tenantSheet.getRange(2, 1, tenantSheet.getLastRow() - 1, 5).getValues();
      tenants = tenantData.map(function(row) {
        return {
          tenantId: row[0],
          leaseId: row[1],
          name: row[2],
          contactInfo: row[3],
          rentPortion: row[4]
        };
      });
    }

    return { success: true, leases: leases, tenants: tenants };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// 12. CALCULATE OUTSTANDING BALANCE FOR A LEASE
function calculateLeaseBalance(leaseId) {
  try {
    var ss = getSpreadsheet();

    var leaseSheet = ss.getSheetByName('fact_Leases');
    if (!leaseSheet) throw new Error("Could not find sheet: fact_Leases");

    var ledgerSheet = ss.getSheetByName('fact_Rent_Ledger');

    var leaseData = leaseSheet.getRange(2, 1, leaseSheet.getLastRow() - 1, 7).getValues();
    var lease = null;
    for (var i = 0; i < leaseData.length; i++) {
      if (leaseData[i][0] === leaseId) {
        lease = {
          startDate: parseLocalDate(leaseData[i][2], false),
          monthlyRent: parseFloat(leaseData[i][4]) || 0
        };
        break;
      }
    }

    if (!lease) throw new Error("Lease not found");

    // Calculate full months passed
    var currentDate = new Date();
    var monthsPassed = 0;
    if (currentDate > lease.startDate) {
        var d1 = lease.startDate;
        var d2 = currentDate;
        var months = (d2.getFullYear() - d1.getFullYear()) * 12;
        months -= d1.getMonth();
        months += d2.getMonth();
        // If current day is less than start day, a full month hasn't passed yet for the current month
        if (d2.getDate() < d1.getDate()) {
            months--;
        }
        // At least 1 month of expected rent if the lease has started
        monthsPassed = Math.max(1, months + 1); // +1 because first month rent is expected at start
    }

    var expectedRent = monthsPassed * lease.monthlyRent;
    var totalCharges = 0;
    var totalPayments = 0;

    if (ledgerSheet && ledgerSheet.getLastRow() > 1) {
      var ledgerData = ledgerSheet.getRange(2, 1, ledgerSheet.getLastRow() - 1, 7).getValues();
      for (var j = 0; j < ledgerData.length; j++) {
        if (ledgerData[j][1] === leaseId) {
          totalCharges += parseFloat(ledgerData[j][5]) || 0;
          totalPayments += parseFloat(ledgerData[j][6]) || 0;
        }
      }
    }

    // Balance = (Expected Rent from start to now) + (Additional Charges like Late Fees) - (Payments)
    var balance = expectedRent + totalCharges - totalPayments;

    return {
      success: true,
      balance: balance,
      expectedRent: expectedRent,
      totalCharges: totalCharges,
      totalPayments: totalPayments,
      monthsPassed: monthsPassed
    };

  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * Creates a custom menu item in Google Sheets when opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('BPM App')
    .addItem('Setup Database Sheets', 'setupDatabase')
    .addItem('Install Weekly Archiving Trigger', 'createWeeklyArchiveTrigger')
    .addToUi();
}

function createWeeklyArchiveTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'archiveOldData') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('archiveOldData')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(1)
    .create();

  SpreadsheetApp.getUi().alert('Weekly archiving trigger successfully installed! It will run every Sunday at 1 AM.');
}

function archiveOldData() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // 30 seconds

    var ss = getSpreadsheet();

    // 1. Calculate the cutoff date (3 months ago)
    var cutoffDate = new Date();
    cutoffDate.setMonth(cutoffDate.getMonth() - 3);

    var laborSheet = ss.getSheetByName('fact_Work_Logs');
    var matSheet = ss.getSheetByName('fact_Material_Expenses');
    var archiveLaborSheet = ss.getSheetByName('archive_fact_Work_Logs');
    var archiveMatSheet = ss.getSheetByName('archive_fact_Material_Expenses');

    if (!laborSheet || !matSheet || !archiveLaborSheet || !archiveMatSheet) {
      console.error("One or more required sheets for archiving are missing.");
      return;
    }

    var laborLastRow = laborSheet.getLastRow();
    var matLastRow = matSheet.getLastRow();

    if (laborLastRow <= 1) {
      return; // Nothing to archive
    }

    var laborRange = laborSheet.getRange(2, 1, laborLastRow - 1, laborSheet.getLastColumn());
    var laborData = laborRange.getValues();

    var matData = [];
    if (matLastRow > 1) {
      matData = matSheet.getRange(2, 1, matLastRow - 1, matSheet.getLastColumn()).getValues();
    }

    var archivedLaborData = [];
    var activeLaborData = [];
    var archivedWorkLogIds = {};

    // 2. Identify work logs to archive
    for (var i = 0; i < laborData.length; i++) {
      var row = laborData[i];
      var logDate = new Date(row[2]); // Date_Completed
      var payrollStatus = (row[8] || '').toString().trim().toLowerCase();
      var billingStatus = (row[9] || '').toString().trim().toLowerCase();

      if (logDate < cutoffDate && payrollStatus === 'paid' && billingStatus === 'paid') {
        archivedLaborData.push(row);
        archivedWorkLogIds[row[0]] = true; // Store Work_Log_ID
      } else {
        activeLaborData.push(row);
      }
    }

    // 3. Identify materials to archive based on archived work log IDs
    var archivedMatData = [];
    var activeMatData = [];

    for (var j = 0; j < matData.length; j++) {
      var mRow = matData[j];
      var wLogId = mRow[1]; // Work_Log_ID

      if (archivedWorkLogIds[wLogId]) {
        archivedMatData.push(mRow);
      } else {
        activeMatData.push(mRow);
      }
    }

    // 4. Append to archive sheets using bulk setValues
    if (archivedLaborData.length > 0) {
      var archLaborLastRow = archiveLaborSheet.getLastRow();
      archiveLaborSheet.getRange(archLaborLastRow + 1, 1, archivedLaborData.length, archivedLaborData[0].length).setValues(archivedLaborData);
    }

    if (archivedMatData.length > 0) {
      var archMatLastRow = archiveMatSheet.getLastRow();
      archiveMatSheet.getRange(archMatLastRow + 1, 1, archivedMatData.length, archivedMatData[0].length).setValues(archivedMatData);
    }

    // 5. Rewrite active data to original sheets to remove archived rows efficiently
    if (archivedLaborData.length > 0 || archivedMatData.length > 0) {
      // Clear content but preserve headers (row 2 and below)
      if (laborLastRow > 1) {
        laborSheet.getRange(2, 1, laborLastRow - 1, laborSheet.getLastColumn()).clearContent();
      }
      if (matLastRow > 1) {
        matSheet.getRange(2, 1, matLastRow - 1, matSheet.getLastColumn()).clearContent();
      }

      // Write back active data
      if (activeLaborData.length > 0) {
        laborSheet.getRange(2, 1, activeLaborData.length, activeLaborData[0].length).setValues(activeLaborData);
      }
      if (activeMatData.length > 0) {
        matSheet.getRange(2, 1, activeMatData.length, activeMatData[0].length).setValues(activeMatData);
      }
    }

  } catch (error) {
    console.error("Error during archiving: " + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Automatically creates all required database sheets, header rows, and initial default dropdown values.
 */
function setupDatabase() {
  var ss = getSpreadsheet();
  
  var schema = [
    {
      name: 'fact_Work_Logs',
      headers: [
        'Work_Log_ID',
        'Timestamp',
        'Date_Completed',
        'Employee_ID',
        'Property_ID',
        'Task_ID',
        'Hours_Worked',
        'Work_Notes',
        'Payroll_Status',
        'Billing_Status',
        'Charge_Target',
        'Charge_Amount_Custom'
      ]
    },
    {
      name: 'archive_fact_Work_Logs',
      headers: [
        'Work_Log_ID',
        'Timestamp',
        'Date_Completed',
        'Employee_ID',
        'Property_ID',
        'Task_ID',
        'Hours_Worked',
        'Work_Notes',
        'Payroll_Status',
        'Billing_Status',
        'Charge_Target',
        'Charge_Amount_Custom'
      ]
    },
    {
      name: 'dim_Billing_Rates',
      headers: [
        'Rate_ID',
        'Task_ID',
        'Task_Name',
        'Employee_ID',
        'Employee_Name',
        'Owner_Hourly_Bill_Rate'
      ]
    },
    {
      name: 'fact_Material_Expenses',
      headers: [
        'Expense_ID',
        'Work_Log_ID',
        'Property_ID',
        'Vendor_Name',
        'Item_Description',
        'Cost',
        'Receipt_URL'
      ]
    },
    {
      name: 'archive_fact_Material_Expenses',
      headers: [
        'Expense_ID',
        'Work_Log_ID',
        'Property_ID',
        'Vendor_Name',
        'Item_Description',
        'Cost',
        'Receipt_URL'
      ]
    },
    {
      name: 'dim_Employees',
      headers: [
        'Employee_ID',
        'Full_Name',
        'Email',
        'Hourly_Pay_Rate',
        'Role',
        'Phone_Number',
        'Status'
      ],
      defaults: [
        ['EMP-001', 'Brian Gross', 'BGross19@gmail.com', 35.00, 'Admin', '555-0199', 'Active'],
        ['EMP-002', 'Jason R.', 'jason@example.com', 25.00, 'Field Crew', '555-0123', 'Active']
      ]
    },
    {
      name: 'Properties',
      headers: ['Property_Name', 'Owner_Company', 'Address'],
      defaults: [
        ['Alpha Phi', 'Greek Housing LLC', '123 Alpha St'],
        ['4638 B', 'Beta Properties', '4638 B Ave'],
        ['645 Ber', 'Beta Properties', '645 Ber St'],
        ['Company Office', 'Internal', 'Main Office']
      ]
    },
    {
      name: 'Tasks',
      headers: ['Task_Name'],
      defaults: [['Handyman'], ['Plumbing'], ['Painting'], ['Lawn Care'], ['Materials Run']]
    },
    {
      name: 'fact_Leases',
      headers: [
        'Lease_ID',
        'Property_ID',
        'Start_Date',
        'End_Date',
        'Monthly_Rent',
        'Security_Deposit',
        'Status',
        'Management_Fee_Percent',
        'Late_Fee_Per_Day',
        'Grace_Period_Days'
      ]
    },
    {
      name: 'dim_Tenants',
      headers: [
        'Tenant_ID',
        'Lease_ID',
        'Name',
        'Contact_Info',
        'Rent_Portion'
      ]
    },
    {
      name: 'fact_Rent_Ledger',
      headers: [
        'Ledger_ID',
        'Lease_ID',
        'Tenant_ID',
        'Date',
        'Transaction_Type',
        'Charge_Amount',
        'Payment_Amount',
        'Payment_Mode',
        'Notes',
        'Rent_Month',
        'Rent_Year'
      ]
    },
    {
      name: 'fact_Payroll_History',
      headers: [
        'Payroll_ID',
        'Employee_ID',
        'Employee_Name',
        'Start_Date',
        'End_Date',
        'Amount_Paid',
        'Payment_Method',
        'Timestamp'
      ]
    }
  ];

  var newlyCreated = 0;

  schema.forEach(function(item) {
    var sheet = ss.getSheetByName(item.name);
    
    if (!sheet) {
      sheet = ss.insertSheet(item.name);
      newlyCreated++;
    }
    
    // If headers are missing, populate headers and formatting
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(item.headers);
      sheet.getRange(1, 1, 1, item.headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
      
      // Populate defaults if provided
      if (item.defaults && item.defaults.length > 0) {
        var numCols = item.defaults[0].length;
        sheet.getRange(2, 1, item.defaults.length, numCols).setValues(item.defaults);
      }
    }
  });

  SpreadsheetApp.getUi().alert(
    'Database setup complete!\n\nChecked ' + schema.length + ' required sheets. Created ' + newlyCreated + ' missing sheet(s).'
  );
}

// 12. BILLING RATES CRUD
function getBillingRates() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('dim_Billing_Rates');
    if (!sheet) return [];

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    return data.map(function(row) {
      return {
        rateId: row[0],
        taskId: row[1],
        taskName: row[2],
        employeeId: row[3],
        employeeName: row[4],
        rate: row[5]
      };
    });
  } catch (error) {
    Logger.log("Error getting billing rates: " + error.message);
    return [];
  }
}

function addBillingRate(payload) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('dim_Billing_Rates');
    if (!sheet) throw new Error("Sheet dim_Billing_Rates not found");

    var rateId = Utilities.getUuid();
    sheet.appendRow([
      rateId,
      payload.taskId,
      payload.taskName || payload.taskId,
      payload.employeeId,
      payload.employeeName || payload.employeeId,
      payload.rate
    ]);
    return { success: true, message: "Rate added" };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function updateBillingRate(payload) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('dim_Billing_Rates');
    if (!sheet) throw new Error("Sheet not found");

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) throw new Error("No rates to update");

    var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === payload.rateId) {
        var row = i + 2;
        sheet.getRange(row, 2, 1, 5).setValues([[
          payload.taskId,
          payload.taskName || payload.taskId,
          payload.employeeId,
          payload.employeeName || payload.employeeId,
          payload.rate
        ]]);
        return { success: true, message: "Rate updated" };
      }
    }
    throw new Error("Rate ID not found");
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function deleteBillingRate(rateId) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('dim_Billing_Rates');
    if (!sheet) throw new Error("Sheet not found");

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) throw new Error("No rates to delete");

    var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === rateId) {
        sheet.deleteRow(i + 2);
        return { success: true, message: "Rate deleted" };
      }
    }
    throw new Error("Rate ID not found");
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// 13. GET EXECUTIVE DASHBOARD DATA
function getExecutiveDashboardData(startDateStr, endDateStr) {
  try {
    var ss = getSpreadsheet();

    var startDate = startDateStr ? parseLocalDate(startDateStr, false) : new Date(new Date().getFullYear(), new Date().getMonth(), 1);
    var endDate = endDateStr ? parseLocalDate(endDateStr, true) : new Date(new Date().getFullYear(), new Date().getMonth() + 1, 0, 23, 59, 59);

    var cutoffDate = new Date();
    cutoffDate.setMonth(cutoffDate.getMonth() - 3);

    var pullArchive = startDate < cutoffDate;

    // 1. RENT METRICS
    var leaseSheet = ss.getSheetByName('fact_Leases');
    var ledgerSheet = ss.getSheetByName('fact_Rent_Ledger');

    var totalRentBilledPeriod = 0;
    var totalRentCollectedPeriod = 0;
    var totalRentOutstanding = 0;

    var arRent = [];

    var leases = [];
    if (leaseSheet && leaseSheet.getLastRow() > 1) {
      var leaseData = leaseSheet.getRange(2, 1, leaseSheet.getLastRow() - 1, 7).getValues();
      leases = leaseData.map(function(row) {
        return {
          leaseId: row[0],
          propertyId: row[1],
          startDate: new Date(row[2]),
          endDate: new Date(row[3]),
          monthlyRent: parseFloat(row[4]) || 0,
          status: row[6]
        };
      });
    }

    leases.forEach(function(lease) {
      var expectedInPeriod = 0;
      var d = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
      while (d <= endDate) {
        var endOfThisMonth = new Date(d.getFullYear(), d.getMonth() + 1, 0);
        if (lease.startDate <= endOfThisMonth && lease.endDate >= d) {
           expectedInPeriod += lease.monthlyRent;
        }
        d.setMonth(d.getMonth() + 1);
      }
      totalRentBilledPeriod += expectedInPeriod;
    });

    if (ledgerSheet && ledgerSheet.getLastRow() > 1) {
      var ledgerData = ledgerSheet.getRange(2, 1, ledgerSheet.getLastRow() - 1, 7).getValues();
      ledgerData.forEach(function(row) {
        var lDate = new Date(row[3]);
        if (lDate >= startDate && lDate <= endDate) {
           totalRentCollectedPeriod += parseFloat(row[6]) || 0;
        }
      });
    }

    leases.forEach(function(lease) {
       var balanceData = calculateLeaseBalance(lease.leaseId);
       if (balanceData.success && balanceData.balance > 0) {
           totalRentOutstanding += balanceData.balance;

           var today = new Date();
           var daysOverdue = 0;
           var monthsBehind = Math.floor(balanceData.balance / lease.monthlyRent);

           if (monthsBehind > 0 || today.getDate() > 5) {
               if (monthsBehind === 0 && today.getDate() > 5) {
                   daysOverdue = today.getDate() - 5;
               } else if (monthsBehind > 0) {
                   daysOverdue = (monthsBehind * 30) + (today.getDate() > 5 ? today.getDate() - 5 : 0);
               }
           }

           if (balanceData.balance > 0 && daysOverdue > 0) {
               arRent.push({
                   type: 'Rent',
                   id: lease.leaseId,
                   propertyId: lease.propertyId,
                   amount: balanceData.balance,
                   daysOverdue: daysOverdue
               });
           }
       }
    });

    // 2. MAINTENANCE METRICS
    var totalMaintBilledPeriod = 0;
    var totalMaintCollectedPeriod = 0;
    var totalMaintOutstanding = 0;
    var arMaint = [];

    var laborSheet = ss.getSheetByName('fact_Work_Logs');
    var archiveLaborSheet = pullArchive ? ss.getSheetByName('archive_fact_Work_Logs') : null;

    var matSheet = ss.getSheetByName('fact_Material_Expenses');
    var archiveMatSheet = pullArchive ? ss.getSheetByName('archive_fact_Material_Expenses') : null;

    var ratesSheet = ss.getSheetByName('dim_Billing_Rates');

    var rateMap = {};
    if (ratesSheet && ratesSheet.getLastRow() > 1) {
      var ratesData = ratesSheet.getRange(2, 1, ratesSheet.getLastRow() - 1, 6).getValues();
      ratesData.forEach(function(row) {
        var taskId = (row[1] || '').toString().trim();
        var empId = (row[3] || '').toString().trim();
        var rate = parseFloat(row[5]) || 0;
        if (taskId && empId) { rateMap[taskId + '_' + empId] = rate; }
      });
    }

    var laborData = [];
    if (laborSheet && laborSheet.getLastRow() > 1) {
      laborData = laborData.concat(laborSheet.getRange(2, 1, laborSheet.getLastRow() - 1, 11).getValues());
    }
    if (archiveLaborSheet && archiveLaborSheet.getLastRow() > 1) {
      laborData = laborData.concat(archiveLaborSheet.getRange(2, 1, archiveLaborSheet.getLastRow() - 1, 11).getValues());
    }

    var matData = [];
    if (matSheet && matSheet.getLastRow() > 1) {
      matData = matData.concat(matSheet.getRange(2, 1, matSheet.getLastRow() - 1, 7).getValues());
    }
    if (archiveMatSheet && archiveMatSheet.getLastRow() > 1) {
      matData = matData.concat(archiveMatSheet.getRange(2, 1, archiveMatSheet.getLastRow() - 1, 7).getValues());
    }

    var logs = [];
    var materials = [];

    if (laborData.length > 0) {
      laborData.forEach(function(row) {
        var empId = (row[3] || '').toString().trim();
        var taskId = (row[5] || '').toString().trim();
        var hours = parseFloat(row[6]) || 0;
        var logDate = new Date(row[2]);
        var billingStatus = (row[9] || '').toString().trim();
        var chargeTarget = (row[10] || 'Owner').toString().trim();

        var billableRate = rateMap[taskId + '_' + empId] !== undefined ? rateMap[taskId + '_' + empId] : (rateMap[taskId + '_DEFAULT'] !== undefined ? rateMap[taskId + '_DEFAULT'] : 0);
        var billableAmount = hours * billableRate;

        if (chargeTarget === 'Owner') {
          logs.push({
            workLogId: row[0],
            date: logDate,
            propertyId: row[4],
            billableAmount: billableAmount,
            billingStatus: billingStatus
          });
        }
      });
    }

    if (matData.length > 0) {
      matData.forEach(function(row) {
        materials.push({
          expenseId: row[0],
          workLogId: row[1],
          propertyId: row[2],
          cost: parseFloat(row[5]) || 0
        });
      });
    }

    var propertyArMap = {};

    logs.forEach(function(log) {
       var matCostForLog = materials.filter(function(m) { return m.workLogId === log.workLogId; }).reduce(function(sum, m) { return sum + m.cost; }, 0);
       var totalCost = log.billableAmount + matCostForLog;

       if (log.date >= startDate && log.date <= endDate) {
           if (log.billingStatus === 'Billed' || log.billingStatus === 'Paid') {
               totalMaintBilledPeriod += totalCost;
           }
           if (log.billingStatus === 'Paid') {
               totalMaintCollectedPeriod += totalCost;
           }
       }

       if (log.billingStatus === 'Billed') {
           totalMaintOutstanding += totalCost;

           var today = new Date();
           var daysOverdue = Math.floor((today - log.date) / (1000 * 60 * 60 * 24));

           if (!propertyArMap[log.propertyId]) {
               propertyArMap[log.propertyId] = { amount: 0, daysOverdue: 0 };
           }
           propertyArMap[log.propertyId].amount += totalCost;
           propertyArMap[log.propertyId].daysOverdue = Math.max(propertyArMap[log.propertyId].daysOverdue, daysOverdue);
       }
    });

    for (var prop in propertyArMap) {
        arMaint.push({
            type: 'Maintenance',
            id: 'INV-' + prop,
            propertyId: prop,
            amount: propertyArMap[prop].amount,
            daysOverdue: propertyArMap[prop].daysOverdue
        });
    }

    // 3. OVERALL CASH FLOW
    var totalCashIn = totalRentCollectedPeriod + totalMaintCollectedPeriod;
    var totalCashOut = 0;
    var totalPayrollPaidPeriod = 0;

    var empSheet = ss.getSheetByName('dim_Employees');
    var employeeMap = {};
    if (empSheet && empSheet.getLastRow() > 1) {
      var empData = empSheet.getRange(2, 1, empSheet.getLastRow() - 1, 7).getValues();
      empData.forEach(function(row) {
        employeeMap[row[0]] = parseFloat(row[3]) || 0;
      });
    }

    if (laborData.length > 0) {
      laborData.forEach(function(row) {
        var empId = (row[3] || '').toString().trim();
        var hours = parseFloat(row[6]) || 0;
        var logDate = new Date(row[2]);
        var payrollStatus = (row[8] || '').toString().trim();

        var grossPay = hours * (employeeMap[empId] || 0);

        if (logDate >= startDate && logDate <= endDate && payrollStatus === 'Paid') {
            totalPayrollPaidPeriod += grossPay;
        }
      });
    }

    var totalMaterialsPeriod = 0;
    if (matData.length > 0) {
        var workLogDates = {};
        if (laborData.length > 0) {
            laborData.forEach(function(row) {
                workLogDates[row[0]] = new Date(row[2]);
            });
        }

        materials.forEach(function(mat) {
            var mDate = workLogDates[mat.workLogId];
            if (mDate && mDate >= startDate && mDate <= endDate) {
                totalMaterialsPeriod += mat.cost;
            }
        });
    }

    totalCashOut = totalPayrollPaidPeriod + totalMaterialsPeriod;
    var netProfit = totalCashIn - totalCashOut;

    var arTable = arRent.concat(arMaint);
    arTable.sort(function(a, b) {
        return b.daysOverdue - a.daysOverdue;
    });

    return {
      success: true,
      data: {
          rent: {
              billed: totalRentBilledPeriod,
              collected: totalRentCollectedPeriod,
              outstanding: totalRentOutstanding
          },
          maintenance: {
              billed: totalMaintBilledPeriod,
              collected: totalMaintCollectedPeriod,
              outstanding: totalMaintOutstanding
          },
          financials: {
              cashIn: totalCashIn,
              cashOut: totalCashOut,
              netProfit: netProfit
          },
          accountsReceivable: arTable
      }
    };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// 14. MARK PROPERTY INVOICE AS PAID
function markPropertyPaid(propertyId) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = getSpreadsheet();
    var laborSheet = ss.getSheetByName('fact_Work_Logs');

    if (laborSheet && laborSheet.getLastRow() > 1) {
      var lastRow = laborSheet.getLastRow();

      var range = laborSheet.getRange(2, 10, lastRow - 1, 1);
      var values = range.getValues();

      var propRange = laborSheet.getRange(2, 5, lastRow - 1, 1);
      var propValues = propRange.getValues();

      var updated = false;

      for (var i = 0; i < values.length; i++) {
        var propIdInRow = (propValues[i][0] || '').toString().trim();
        var status = (values[i][0] || '').toString().trim().toLowerCase();

        if (propIdInRow === propertyId && status === 'billed') {
          values[i][0] = 'Paid';
          updated = true;
        }
      }

      if (updated) {
        range.setValues(values);
      }
    }
    return { success: true, message: 'Property billing status updated to Paid.' };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}
// 17. GET RENT ROLL DATA FOR RENT TRACKING
function getRentRollData(year) {
  try {
    var ss = getSpreadsheet();
    var propSheet = ss.getSheetByName('Properties');
    var leaseSheet = ss.getSheetByName('fact_Leases');
    var tenantSheet = ss.getSheetByName('dim_Tenants');
    var ledgerSheet = ss.getSheetByName('fact_Rent_Ledger');

    var props = [];
    if (propSheet && propSheet.getLastRow() > 1) {
      var pData = propSheet.getRange(2, 1, propSheet.getLastRow() - 1, 3).getValues();
      pData.forEach(function(row) {
        props.push({ name: row[0], owner: row[1] || "", address: row[2] || "" });
      });
    }

    var leases = [];
    if (leaseSheet && leaseSheet.getLastRow() > 1) {
      var lData = leaseSheet.getRange(2, 1, leaseSheet.getLastRow() - 1, 7).getValues();
      lData.forEach(function(row) {
        leases.push({ leaseId: row[0], propertyId: row[1], startDate: new Date(row[2]), endDate: new Date(row[3]), monthlyRent: row[4], status: row[6] });
      });
    }

    var tenants = [];
    if (tenantSheet && tenantSheet.getLastRow() > 1) {
      var tData = tenantSheet.getRange(2, 1, tenantSheet.getLastRow() - 1, 5).getValues();
      tData.forEach(function(row) {
        tenants.push({ tenantId: row[0], leaseId: row[1], name: row[2], rentPortion: row[4] });
      });
    }

    var ledgers = [];
    if (ledgerSheet && ledgerSheet.getLastRow() > 1) {
      var ldData = ledgerSheet.getRange(2, 1, ledgerSheet.getLastRow() - 1, 11).getValues();
      ldData.forEach(function(row) {
        if (row[4] === 'Rent' && (parseFloat(row[6]) > 0)) {
          var rentM = row[9] !== "" ? parseInt(row[9]) : new Date(row[3]).getMonth();
          var rentY = row[10] !== "" ? parseInt(row[10]) : new Date(row[3]).getFullYear();
          if (rentY == year) {
            ledgers.push({ leaseId: row[1], tenantId: row[2], date: new Date(row[3]), rentMonth: rentM, rentYear: rentY, amount: parseFloat(row[6]) });
          }
        }
      });
    }

    var rentRoll = [];

    // Build rows for each tenant in active leases for the year
    tenants.forEach(function(tenant) {
      var lease = leases.find(function(l) { return l.leaseId === tenant.leaseId; });
      if (!lease) return;

      // Check if lease is active in the given year
      if (lease.startDate.getFullYear() > year || lease.endDate.getFullYear() < year) return;

      var prop = props.find(function(p) { return p.name === lease.propertyId; });
      var propName = prop ? prop.name : lease.propertyId;
      var propOwner = prop ? prop.owner : "";
      var propAddress = prop ? prop.address : "";

      var row = {
        leaseId: lease.leaseId,
        tenantId: tenant.tenantId,
        property: propName,
        owner: propOwner,
        address: propAddress,
        tenantName: tenant.name,
        expectedRent: tenant.rentPortion,
        payments: { 0:0, 1:0, 2:0, 3:0, 4:0, 5:0, 6:0, 7:0, 8:0, 9:0, 10:0, 11:0 },
        activeMonths: { 0:false, 1:false, 2:false, 3:false, 4:false, 5:false, 6:false, 7:false, 8:false, 9:false, 10:false, 11:false }
      };

      // Figure out which months the lease is active
      for (var m = 0; m < 12; m++) {
        var startOfM = new Date(year, m, 1);
        var endOfM = new Date(year, m + 1, 0);
        if (startOfM <= lease.endDate && endOfM >= lease.startDate) {
          row.activeMonths[m] = true;
        }
      }

      // Sum payments by month
      var tLedgers = ledgers.filter(function(ld) { return ld.tenantId === tenant.tenantId && ld.leaseId === tenant.leaseId; });
      tLedgers.forEach(function(ld) {
        var m = ld.rentMonth;
        if (m >= 0 && m < 12) {
          row.payments[m] += ld.amount;
        }
      });

      rentRoll.push(row);
    });

    return { success: true, data: rentRoll };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// Properties CRUD

function _isAdminOrOwner() {
  var activeUser = getUserRole();
  return activeUser.role === 'admin' || activeUser.role === 'owner';
}

function getProperties() {
  if (!_isAdminOrOwner()) return [];
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Properties');
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    return data.map(function(row) {
      return { propertyName: row[0], ownerCompany: row[1], address: row[2] };
    });
  } catch (error) {
    return [];
  }
}

function addProperty(payload) {
  if (!_isAdminOrOwner()) return { success: false, error: 'Unauthorized' };
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Properties');
    if (!sheet) throw new Error("Properties sheet not found");
    sheet.appendRow([payload.propertyName, payload.ownerCompany || "", payload.address || ""]);
    return { success: true, message: "Property added" };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function updateProperty(payload) {
  if (!_isAdminOrOwner()) return { success: false, error: 'Unauthorized' };
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Properties');
    if (!sheet) throw new Error("Properties sheet not found");
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) throw new Error("No properties to update");
    var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === payload.oldPropertyName) {
        sheet.getRange(i + 2, 1, 1, 3).setValues([[payload.propertyName, payload.ownerCompany || "", payload.address || ""]]);
        return { success: true, message: "Property updated" };
      }
    }
    throw new Error("Property not found");
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function deleteProperty(propertyName) {
  if (!_isAdminOrOwner()) return { success: false, error: 'Unauthorized' };
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Properties');
    if (!sheet) throw new Error("Properties sheet not found");
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) throw new Error("No properties to delete");
    var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === propertyName) {
        sheet.deleteRow(i + 2);
        return { success: true, message: "Property deleted" };
      }
    }
    throw new Error("Property not found");
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// Employees CRUD
function getEmployees() {
  if (!_isAdminOrOwner()) return [];
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('dim_Employees');
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    return data.map(function(row) {
      return {
        employeeId: row[0],
        fullName: row[1],
        email: row[2],
        hourlyPayRate: row[3],
        role: row[4],
        phoneNumber: row[5],
        status: row[6]
      };
    });
  } catch (error) {
    return [];
  }
}

function addEmployee(payload) {
  if (!_isAdminOrOwner()) return { success: false, error: 'Unauthorized' };
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('dim_Employees');
    if (!sheet) throw new Error("dim_Employees sheet not found");
    var newId = payload.employeeId || ("EMP-" + Utilities.getUuid().substring(0, 5).toUpperCase());
    sheet.appendRow([
      newId,
      payload.fullName || "",
      payload.email || "",
      payload.hourlyPayRate || 0,
      payload.role || "Field Crew",
      payload.phoneNumber || "",
      payload.status || "Active"
    ]);
    return { success: true, message: "Employee added" };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function updateEmployee(payload) {
  if (!_isAdminOrOwner()) return { success: false, error: 'Unauthorized' };
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('dim_Employees');
    if (!sheet) throw new Error("dim_Employees sheet not found");
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) throw new Error("No employees to update");
    var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === payload.employeeId) {
        sheet.getRange(i + 2, 1, 1, 7).setValues([[
          payload.employeeId,
          payload.fullName || "",
          payload.email || "",
          payload.hourlyPayRate || 0,
          payload.role || "Field Crew",
          payload.phoneNumber || "",
          payload.status || "Active"
        ]]);
        return { success: true, message: "Employee updated" };
      }
    }
    throw new Error("Employee not found");
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function deleteEmployee(employeeId) {
  if (!_isAdminOrOwner()) return { success: false, error: 'Unauthorized' };
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('dim_Employees');
    if (!sheet) throw new Error("dim_Employees sheet not found");
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) throw new Error("No employees to delete");
    var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === employeeId) {
        sheet.deleteRow(i + 2);
        return { success: true, message: "Employee deleted" };
      }
    }
    throw new Error("Employee not found");
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// Tasks CRUD
function getTasks() {
  if (!_isAdminOrOwner()) return [];
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Tasks');
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    return data.map(function(row) {
      return { taskName: row[0] };
    });
  } catch (error) {
    return [];
  }
}

function addTask(payload) {
  if (!_isAdminOrOwner()) return { success: false, error: 'Unauthorized' };
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Tasks');
    if (!sheet) throw new Error("Tasks sheet not found");
    sheet.appendRow([payload.taskName]);
    return { success: true, message: "Task added" };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function updateTask(payload) {
  if (!_isAdminOrOwner()) return { success: false, error: 'Unauthorized' };
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Tasks');
    if (!sheet) throw new Error("Tasks sheet not found");
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) throw new Error("No tasks to update");
    var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === payload.oldTaskName) {
        sheet.getRange(i + 2, 1, 1, 1).setValues([[payload.taskName]]);
        return { success: true, message: "Task updated" };
      }
    }
    throw new Error("Task not found");
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function deleteTask(taskName) {
  if (!_isAdminOrOwner()) return { success: false, error: 'Unauthorized' };
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Tasks');
    if (!sheet) throw new Error("Tasks sheet not found");
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) throw new Error("No tasks to delete");
    var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === taskName) {
        sheet.deleteRow(i + 2);
        return { success: true, message: "Task deleted" };
      }
    }
    throw new Error("Task not found");
  } catch (error) {
    return { success: false, error: error.message };
  }
}


// ==========================================
// BULK ADD FUNCTIONS
// ==========================================

function addBillingRatesBulk(payloads) {
  if (!_isAdminOrOwner()) return { success: false, error: 'Unauthorized' };
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('dim_Billing_Rates');
    if (!sheet) throw new Error("Sheet dim_Billing_Rates not found");

    if (!payloads || payloads.length === 0) return { success: true, message: "No data to add" };

    var rows = [];
    for (var i = 0; i < payloads.length; i++) {
      var p = payloads[i];
      var rateId = Utilities.getUuid();
      rows.push([
        rateId,
        p.taskId,
        p.taskName || p.taskId,
        p.employeeId,
        p.employeeName || p.employeeId,
        p.rate
      ]);
    }

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }
    return { success: true, message: rows.length + " Rates added" };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function addPropertiesBulk(payloads) {
  if (!_isAdminOrOwner()) return { success: false, error: 'Unauthorized' };
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Properties');
    if (!sheet) throw new Error("Properties sheet not found");

    if (!payloads || payloads.length === 0) return { success: true, message: "No data to add" };

    var rows = [];
    for (var i = 0; i < payloads.length; i++) {
      var p = payloads[i];
      rows.push([
        p.propertyName,
        p.ownerCompany || "",
        p.address || ""
      ]);
    }

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }
    return { success: true, message: rows.length + " Properties added" };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function addEmployeesBulk(payloads) {
  if (!_isAdminOrOwner()) return { success: false, error: 'Unauthorized' };
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('dim_Employees');
    if (!sheet) throw new Error("dim_Employees sheet not found");

    if (!payloads || payloads.length === 0) return { success: true, message: "No data to add" };

    var rows = [];
    for (var i = 0; i < payloads.length; i++) {
      var p = payloads[i];
      var newId = p.employeeId || ("EMP-" + Utilities.getUuid().substring(0, 5).toUpperCase());
      rows.push([
        newId,
        p.fullName || "",
        p.email || "",
        p.hourlyPayRate || 0,
        p.role || "Field Crew",
        p.phoneNumber || "",
        p.status || "Active"
      ]);
    }

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }
    return { success: true, message: rows.length + " Employees added" };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function addTasksBulk(payloads) {
  if (!_isAdminOrOwner()) return { success: false, error: 'Unauthorized' };
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('dim_Task_Categories');
    if (!sheet) throw new Error("dim_Task_Categories sheet not found");

    if (!payloads || payloads.length === 0) return { success: true, message: "No data to add" };

    var rows = [];
    for (var i = 0; i < payloads.length; i++) {
      var p = payloads[i];
      rows.push([
        p.taskName,
        p.description || "",
        p.isBillable ? "Yes" : "No"
      ]);
    }

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }
    return { success: true, message: rows.length + " Tasks added" };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// HELPER: Robust Local Date Parsing
// Avoids UTC shift issues by constructing dates explicitly in the script's timezone.
function parseLocalDate(dateString, isEndOfDay) {
  if (!dateString) return null;
  // If it's already a Date object, return it or adjust its time.
  if (dateString instanceof Date || Object.prototype.toString.call(dateString) === '[object Date]') {
    if (isEndOfDay) {
        return new Date(dateString.getFullYear(), dateString.getMonth(), dateString.getDate(), 23, 59, 59);
    }
    return new Date(dateString.getFullYear(), dateString.getMonth(), dateString.getDate());
  }

  var str = dateString.toString();
  if (str.indexOf('T') !== -1) {
     str = str.split('T')[0];
  }

  var parts = str.split('-');
  if (parts.length === 3) {
      var year = parseInt(parts[0], 10);
      var month = parseInt(parts[1], 10) - 1;
      var day = parseInt(parts[2], 10);
      if (isEndOfDay) {
        return new Date(year, month, day, 23, 59, 59);
      }
      return new Date(year, month, day);
  }

  // Fallback if format is unrecognized
  var fallbackDate = new Date(dateString);
  if (isEndOfDay && fallbackDate && !isNaN(fallbackDate.getTime())) {
     fallbackDate.setHours(23, 59, 59, 999);
  }
  return fallbackDate;
}
