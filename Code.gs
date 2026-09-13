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
// Checks the active Google account against 'dim_Employees' sheet to return user details.
function getUserRole() {
  var userEmail = Session.getActiveUser().getEmail();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var empSheet = ss.getSheetByName('dim_Employees');
  
  var userInfo = {
    email: userEmail,
    role: 'employee',
    employeeId: '',
    name: ''
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
          userInfo.name = row[1];       // Column B: Full_Name
          var roleInSheet = (row[4] || '').toString().trim().toLowerCase(); // Column E: Role
          userInfo.role = (roleInSheet.indexOf('admin') > -1) ? 'admin' : 'employee';
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
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
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
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
  }
}

// 3.5 GET EMPLOYEE WORK LOGS & DAILY TOTALS
function getEmployeeWorkLogs() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
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
    var ss = SpreadsheetApp.getActiveSpreadsheet();
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
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
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

// 4. ADMIN PAYROLL & BILLING CALCULATIONS WITH DATE FILTERING
// Fetches records from fact_Work_Logs, dim_Employees, and fact_Material_Expenses with optional start/end date filtering.
function getAdminPayrollData(startDateStr, endDateStr, statusFilter) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var startDate = startDateStr ? new Date(startDateStr + 'T00:00:00') : null;
    var endDate = endDateStr ? new Date(endDateStr + 'T23:59:59') : null;

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
    var workLogs = [];
    if (laborSheet && laborSheet.getLastRow() > 1) {
      // Adjusted to read 11 columns to include Charge_Target
      var laborData = laborSheet.getRange(2, 1, laborSheet.getLastRow() - 1, 11).getValues();
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
        var billableAmount = hours * billableRate;

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
          billableAmount: billableAmount
        });
      });
    }

    // Read Material Expenses
    var matSheet = ss.getSheetByName('fact_Material_Expenses');
    var materials = [];
    if (matSheet && matSheet.getLastRow() > 1) {
      var matData = matSheet.getRange(2, 1, matSheet.getLastRow() - 1, 7).getValues();
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
function markPayrollPaid(employeeId) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var laborSheet = ss.getSheetByName('fact_Work_Logs');
    if (!laborSheet) throw new Error("Could not find sheet: fact_Work_Logs");

    var lastRow = laborSheet.getLastRow();

    if (lastRow > 1) {
      var range = laborSheet.getRange(2, 9, lastRow - 1, 1); // Get only the Payroll_Status column
      var values = range.getValues();

      var empRange = laborSheet.getRange(2, 4, lastRow - 1, 1); // Get Employee_ID column
      var empValues = empRange.getValues();

      var updated = false;

      for (var i = 0; i < values.length; i++) {
        var empIdInRow = (empValues[i][0] || '').toString().trim();
        var status = (values[i][0] || '').toString().trim().toLowerCase();

        if ((!employeeId || empIdInRow === employeeId) && (status === 'pending' || status === 'approved')) {
          values[i][0] = 'Paid';
          updated = true;
        }
      }

      // Perform one bulk update on the single column
      if (updated) {
        range.setValues(values);
      }
    }
    return { success: true, message: 'Payroll marked as Paid successfully.' };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
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
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
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
    var ss = SpreadsheetApp.getActiveSpreadsheet();

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
      "Active"
    ]);

    // Add tenants
    if (payload.tenants && payload.tenants.length > 0) {
      payload.tenants.forEach(function(tenant) {
        var tenantId = "TNT-" + Utilities.getUuid().substring(0, 8).toUpperCase();
        tenantSheet.appendRow([
          tenantId,
          leaseId,
          tenant.name,
          tenant.contactInfo
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
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

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
      payload.paymentAmount || 0
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
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var leaseSheet = ss.getSheetByName('fact_Leases');
    var tenantSheet = ss.getSheetByName('dim_Tenants');

    var leases = [];
    var tenants = [];

    if (leaseSheet && leaseSheet.getLastRow() > 1) {
      var leaseData = leaseSheet.getRange(2, 1, leaseSheet.getLastRow() - 1, 7).getValues();
      leases = leaseData.map(function(row) {
        return {
          leaseId: row[0],
          propertyId: row[1],
          startDate: row[2],
          endDate: row[3],
          monthlyRent: row[4],
          securityDeposit: row[5],
          status: row[6]
        };
      });
    }

    if (tenantSheet && tenantSheet.getLastRow() > 1) {
      var tenantData = tenantSheet.getRange(2, 1, tenantSheet.getLastRow() - 1, 4).getValues();
      tenants = tenantData.map(function(row) {
        return {
          tenantId: row[0],
          leaseId: row[1],
          name: row[2],
          contactInfo: row[3]
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
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var leaseSheet = ss.getSheetByName('fact_Leases');
    if (!leaseSheet) throw new Error("Could not find sheet: fact_Leases");

    var ledgerSheet = ss.getSheetByName('fact_Rent_Ledger');

    var leaseData = leaseSheet.getRange(2, 1, leaseSheet.getLastRow() - 1, 7).getValues();
    var lease = null;
    for (var i = 0; i < leaseData.length; i++) {
      if (leaseData[i][0] === leaseId) {
        lease = {
          startDate: new Date(leaseData[i][2]),
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
    .addToUi();
}

/**
 * Automatically creates all required database sheets, header rows, and initial default dropdown values.
 */
function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
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
        'Charge_Target'
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
      headers: ['Property_Name'],
      defaults: [['Alpha Phi'], ['4638 B'], ['645 Ber'], ['Company Office']]
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
        'Status'
      ]
    },
    {
      name: 'dim_Tenants',
      headers: [
        'Tenant_ID',
        'Lease_ID',
        'Name',
        'Contact_Info'
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
        'Payment_Amount'
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
    var ss = SpreadsheetApp.getActiveSpreadsheet();
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
    var ss = SpreadsheetApp.getActiveSpreadsheet();
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
    var ss = SpreadsheetApp.getActiveSpreadsheet();
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
    var ss = SpreadsheetApp.getActiveSpreadsheet();
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
