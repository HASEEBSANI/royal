

// ============== Configuration ==============
var USERS_SHEET = 'Users';
var LOGS_SHEET = 'Logs';
var PARTIES_SHEET = 'Parties';
var CASH_BOOK_SHEET = 'CashBook';
var DAYBOOK_SHEET_NAME = 'DayBook';
var PARTY_LEDGER_SHEET = 'PartyLedger';
var DAILY_SUMMARY_SHEET = 'DailySummary';
var VOUCHER_SEQUENCE_SHEET = 'VoucherSequence';
var SETTINGS_SHEET = 'Settings';
var ASSETS_FOLDER_NAME = 'ASSETS';

// ============== Main Web App Entry Point ==============
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index');

  return template
    .evaluate()
    .setTitle('Daily Book Cash & Party Ledger System')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============== Initialize Sheets ==============
function initializeSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create Users sheet if not exists
  var usersSheet = ss.getSheetByName(USERS_SHEET);
  if (!usersSheet) {
    usersSheet = ss.insertSheet(USERS_SHEET);
    usersSheet.appendRow(['Username', 'Email', 'Password', 'Role', 'Status', 'ProfileImage', 'ThemeMode', 'CustomColors', 'CreatedAt', 'CreatedBy', 'UpdatedAt', 'UpdatedBy']);
    usersSheet.getRange(1, 1, 1, 12).setBackground('#001f3f').setFontColor('white').setFontWeight('bold');

    // Add default admin user
    var timestamp = new Date().toISOString();
    var defaultLogo = 'https://drive.google.com/file/d/165GgME6_H3EYDwA-eqUIqDNzHIq3KFy0/view';
    usersSheet.appendRow(['admin', 'admin@example.com', 'admin123', 'Admin', 'Active', defaultLogo, 'light', '', timestamp, 'System', timestamp, 'System']);

    // Add default regular user
    usersSheet.appendRow(['user', 'user@example.com', 'user123', 'User', 'Active', defaultLogo, 'light', '', timestamp, 'System', timestamp, 'System']);
  }

  // Create Logs sheet if not exists
  var logsSheet = ss.getSheetByName(LOGS_SHEET);
  if (!logsSheet) {
    logsSheet = ss.insertSheet(LOGS_SHEET);
    logsSheet.appendRow(['Timestamp', 'User', 'Action', 'Details']);
    logsSheet.getRange(1, 1, 1, 4).setBackground('#001f3f').setFontColor('white').setFontWeight('bold');
  }

  // Create Parties sheet if not exists
  var partiesSheet = ss.getSheetByName(PARTIES_SHEET);
  if (!partiesSheet) {
    partiesSheet = ss.insertSheet(PARTIES_SHEET);
    partiesSheet.appendRow(['PartyID', 'PartyName', 'PartyType', 'ContactNumber', 'Address', 'OpeningBalance', 'BalanceType', 'CurrentBalance', 'IsActive', 'CreatedAt', 'CreatedBy', 'UpdatedAt', 'UpdatedBy']);
    partiesSheet.getRange(1, 1, 1, 13).setBackground('#001f3f').setFontColor('white').setFontWeight('bold');
  }

  // Create Cash Book sheet if not exists
  var cashBookSheet = ss.getSheetByName(CASH_BOOK_SHEET);
  if (!cashBookSheet) {
    cashBookSheet = ss.insertSheet(CASH_BOOK_SHEET);
    cashBookSheet.appendRow(['CashID', 'VoucherNo', 'TransactionDate', 'Particulars', 'CashIn', 'CashOut', 'EntryType', 'PartyID', 'PartyName', 'PaymentMode', 'ReferenceNo', 'RunningBalance', 'LedgerLinked', 'CreatedBy', 'CreatedAt', 'UpdatedAt']);
    cashBookSheet.getRange(1, 1, 1, 16).setBackground('#001f3f').setFontColor('white').setFontWeight('bold');
  }

  // Create Party Ledger sheet if not exists
  var partyLedgerSheet = ss.getSheetByName(PARTY_LEDGER_SHEET);
  if (!partyLedgerSheet) {
    partyLedgerSheet = ss.insertSheet(PARTY_LEDGER_SHEET);
    partyLedgerSheet.appendRow(['LedgerID', 'TransactionDate', 'PartyID', 'PartyName', 'InvoiceNo', 'DebitAmount', 'CreditAmount', 'TransactionType', 'PaymentMode', 'Narration', 'ReferenceNo', 'RunningBalance', 'CashBookID', 'IsCashEntry', 'CreatedBy', 'CreatedAt', 'UpdatedAt']);
    partyLedgerSheet.getRange(1, 1, 1, 17).setBackground('#001f3f').setFontColor('white').setFontWeight('bold');
  }

  // Create Daily Summary sheet if not exists
  var dailySummarySheet = ss.getSheetByName(DAILY_SUMMARY_SHEET);
  if (!dailySummarySheet) {
    dailySummarySheet = ss.insertSheet(DAILY_SUMMARY_SHEET);
    dailySummarySheet.appendRow(['SummaryID', 'SummaryDate', 'OpeningBalance', 'TotalCashIn', 'TotalCashOut', 'ClosingBalance', 'TotalTransactions', 'CreatedAt']);
    dailySummarySheet.getRange(1, 1, 1, 8).setBackground('#001f3f').setFontColor('white').setFontWeight('bold');
  }

  // Create Voucher Sequence sheet if not exists
  var voucherSequenceSheet = ss.getSheetByName(VOUCHER_SEQUENCE_SHEET);
  if (!voucherSequenceSheet) {
    voucherSequenceSheet = ss.insertSheet(VOUCHER_SEQUENCE_SHEET);
    voucherSequenceSheet.appendRow(['SequenceID', 'VoucherDate', 'LastNumber', 'CreatedAt']);
    voucherSequenceSheet.getRange(1, 1, 1, 4).setBackground('#001f3f').setFontColor('white').setFontWeight('bold');
  }

  // Create Settings sheet if not exists
  var settingsSheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SETTINGS_SHEET);
    settingsSheet.appendRow(['SettingKey', 'SettingValue', 'UpdatedAt', 'UpdatedBy']);
    settingsSheet.getRange(1, 1, 1, 4).setBackground('#001f3f').setFontColor('white').setFontWeight('bold');

    // Add default settings
    var timestamp = new Date().toISOString();
    settingsSheet.appendRow(['currency', 'PKR', timestamp, 'System']);
    settingsSheet.appendRow(['currencySymbol', 'Rs.', timestamp, 'System']);
    settingsSheet.appendRow(['systemName', 'Cash Book System', timestamp, 'System']);
  }

  return { success: true, message: 'Sheets initialized successfully!' };
}

// ============== Authentication ==============
function authenticateUser(username, password) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET);

    if (!sheet) {
      return { success: false, message: 'Users sheet not found. Please initialize the system first.' };
    }

    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === username) {
        // Check if user is active
        if (data[i][4] !== 'Active') {
          addLog(username, 'Login Failed', 'Account is inactive');
          return { success: false, message: 'Account is inactive. Please contact administrator.' };
        }

        // Verify password (plain text comparison)
        if (password === data[i][2]) {
          addLog(username, 'Login Success', 'User logged in successfully');
          return {
            success: true,
            username: data[i][0],
            email: data[i][1],
            role: data[i][3],
            profileImage: data[i][5] || '',
            themeMode: data[i][6] || 'light',
            customColors: data[i][7] || ''
          };
        } else {
          addLog(username, 'Login Failed', 'Invalid password');
          return { success: false, message: 'Invalid password' };
        }
      }
    }

    addLog(username, 'Login Failed', 'Username not found');
    return { success: false, message: 'Username not found' };

  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// ============== Get All Users (Admin Only) ==============
function getAllUsers() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET);

    if (!sheet) {
      return { success: false, message: 'Users sheet not found' };
    }

    var data = sheet.getDataRange().getValues();
    var users = [];

    for (var i = 1; i < data.length; i++) {
      var createdAtValue = data[i][8];  // CreatedAt at column 9
      var updatedAtValue = data[i][10]; // UpdatedAt at column 11

      // Convert dates to ISO format if needed
      if (createdAtValue && !(typeof createdAtValue === 'string' && createdAtValue.includes('T'))) {
        createdAtValue = new Date(createdAtValue).toISOString();
      }
      if (updatedAtValue && !(typeof updatedAtValue === 'string' && updatedAtValue.includes('T'))) {
        updatedAtValue = new Date(updatedAtValue).toISOString();
      }

      users.push({
        Username: data[i][0],
        Email: data[i][1],
        Role: data[i][3],
        Status: data[i][4],
        CreatedAt: createdAtValue,
        CreatedBy: data[i][9],  // CreatedBy at column 10
        UpdatedAt: updatedAtValue,
        UpdatedBy: data[i][11]   // UpdatedBy at column 12
      });
    }

    return { success: true, data: users };

  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// ============== Add User (Admin Only) ==============
function addUser(userData, currentUser) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET);

    if (!sheet) {
      return { success: false, message: 'Users sheet not found' };
    }

    // Check if username already exists
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === userData.Username) {
        return { success: false, message: 'Username already exists' };
      }
    }

    // Store plain password
    var timestamp = new Date().toISOString();
    var defaultLogo = 'https://drive.google.com/file/d/165GgME6_H3EYDwA-eqUIqDNzHIq3KFy0/view';

    // Add new user
    sheet.appendRow([
      userData.Username,
      userData.Email,
      userData.Password,
      userData.Role || 'User',
      userData.Status || 'Active',
      defaultLogo,        // ProfileImage
      'light',            // ThemeMode
      '',                 // CustomColors
      timestamp,          // CreatedAt
      currentUser,        // CreatedBy
      timestamp,          // UpdatedAt
      currentUser         // UpdatedBy
    ]);

    addLog(currentUser, 'User Added', 'Added user: ' + userData.Username);

    return { success: true, message: 'User added successfully!' };

  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// ============== Update User (Admin Only) ==============
function updateUser(username, userData, currentUser) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET);

    if (!sheet) {
      return { success: false, message: 'Users sheet not found' };
    }

    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === username) {
        var timestamp = new Date().toISOString();

        // Update user data
        sheet.getRange(i + 1, 2).setValue(userData.Email);

        // Only update password if provided (plain text)
        if (userData.Password && userData.Password.trim() !== '') {
          sheet.getRange(i + 1, 3).setValue(userData.Password);
        }

        sheet.getRange(i + 1, 4).setValue(userData.Role);
        sheet.getRange(i + 1, 5).setValue(userData.Status);
        sheet.getRange(i + 1, 11).setValue(timestamp); // UpdatedAt
        sheet.getRange(i + 1, 12).setValue(currentUser); // UpdatedBy

        addLog(currentUser, 'User Updated', 'Updated user: ' + username);

        return { success: true, message: 'User updated successfully!' };
      }
    }

    return { success: false, message: 'User not found' };

  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// ============== Delete User (Admin Only) ==============
function deleteUser(username, currentUser) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET);

    if (!sheet) {
      return { success: false, message: 'Users sheet not found' };
    }

    // Prevent deleting yourself
    if (username === currentUser) {
      return { success: false, message: 'You cannot delete your own account' };
    }

    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === username) {
        sheet.deleteRow(i + 1);
        addLog(currentUser, 'User Deleted', 'Deleted user: ' + username);
        return { success: true, message: 'User deleted successfully!' };
      }
    }

    return { success: false, message: 'User not found' };

  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// ============== Update My Account (User Profile) ==============
function updateMyAccount(username, formData) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET);

    if (!sheet) {
      return { success: false, message: 'Users sheet not found' };
    }

    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === username) {
        // Verify current password (plain text comparison)
        if (formData.CurrentPassword !== data[i][2]) {
          return { success: false, message: 'Current password is incorrect' };
        }

        var timestamp = new Date().toISOString();

        // Update email
        sheet.getRange(i + 1, 2).setValue(formData.Email);

        // Update password if new password is provided (plain text)
        if (formData.NewPassword && formData.NewPassword.trim() !== '') {
          sheet.getRange(i + 1, 3).setValue(formData.NewPassword);
        }

        // Update timestamp
        sheet.getRange(i + 1, 11).setValue(timestamp); // UpdatedAt
        sheet.getRange(i + 1, 12).setValue(username); // UpdatedBy

        addLog(username, 'Profile Updated', 'Updated own profile');

        return { success: true, message: 'Account updated successfully!' };
      }
    }

    return { success: false, message: 'User not found' };

  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// ============== Get Dashboard Stats (Admin Only) ==============
function getDashboardStats() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ===== CASH BOOK STATS =====
    var cashBookSheet = ss.getSheetByName(CASH_BOOK_SHEET);
    var totalCashBalance = 0;
    var todayCashIn = 0;
    var todayCashOut = 0;
    var monthCashIn = 0;
    var monthCashOut = 0;
    var totalTransactions = 0;
    var lastTransactionDate = null;

    if (cashBookSheet) {
      var cashData = cashBookSheet.getDataRange().getValues();
      totalTransactions = cashData.length - 1; // Exclude header

      // Get latest running balance (last entry)
      if (cashData.length > 1) {
        totalCashBalance = parseFloat(cashData[cashData.length - 1][11]) || 0; // RunningBalance column
        lastTransactionDate = new Date(cashData[cashData.length - 1][2]); // Last transaction date
      }

      // Calculate today's and month's cash in/out
      var today = new Date();
      var todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      var currentMonth = today.getMonth();
      var currentYear = today.getFullYear();

      for (var i = 1; i < cashData.length; i++) {
        var transactionDate = new Date(cashData[i][2]); // TransactionDate
        var transactionDateStr = Utilities.formatDate(transactionDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        var transMonth = transactionDate.getMonth();
        var transYear = transactionDate.getFullYear();

        // Today's totals
        if (transactionDateStr === todayStr) {
          todayCashIn += parseFloat(cashData[i][4]) || 0; // CashIn
          todayCashOut += parseFloat(cashData[i][5]) || 0; // CashOut
        }

        // Month-to-date totals
        if (transMonth === currentMonth && transYear === currentYear) {
          monthCashIn += parseFloat(cashData[i][4]) || 0;
          monthCashOut += parseFloat(cashData[i][5]) || 0;
        }
      }
    }

    // ===== PARTIES STATS =====
    var partiesSheet = ss.getSheetByName(PARTIES_SHEET);
    var totalReceivable = 0;
    var totalPayable = 0;
    var totalParties = 0;
    var activeParties = 0;
    var topReceivables = [];
    var topPayables = [];

    if (partiesSheet) {
      var partiesData = partiesSheet.getDataRange().getValues();
      totalParties = partiesData.length - 1; // Exclude header

      var receivableList = [];
      var payableList = [];

      for (var i = 1; i < partiesData.length; i++) {
        var currentBalance = parseFloat(partiesData[i][7]) || 0; // CurrentBalance
        var isActive = partiesData[i][9]; // IsActive
        var partyName = partiesData[i][1]; // PartyName

        if (isActive === true || isActive === 'TRUE') {
          activeParties++;

          if (currentBalance > 0) {
            // Receivable (customer owes us)
            totalReceivable += currentBalance;
            receivableList.push({
              name: partyName,
              amount: currentBalance,
              type: 'Receivable'
            });
          } else if (currentBalance < 0) {
            // Payable (we owe supplier)
            totalPayable += Math.abs(currentBalance);
            payableList.push({
              name: partyName,
              amount: Math.abs(currentBalance),
              type: 'Payable'
            });
          }
        }
      }

      // Sort and get top 5 receivables
      receivableList.sort(function(a, b) { return b.amount - a.amount; });
      topReceivables = receivableList.slice(0, 5);

      // Sort and get top 5 payables
      payableList.sort(function(a, b) { return b.amount - a.amount; });
      topPayables = payableList.slice(0, 5);
    }

    // ===== USERS STATS =====
    var usersSheet = ss.getSheetByName(USERS_SHEET);
    var totalUsers = 0;
    var activeUsers = 0;

    if (usersSheet) {
      var usersData = usersSheet.getDataRange().getValues();
      totalUsers = usersData.length - 1;

      for (var i = 1; i < usersData.length; i++) {
        if (usersData[i][4] === 'Active') {
          activeUsers++;
        }
      }
    }

    return {
      success: true,
      data: {
        // Cash Stats
        totalCashBalance: totalCashBalance,
        todayCashIn: todayCashIn,
        todayCashOut: todayCashOut,
        monthCashIn: monthCashIn,
        monthCashOut: monthCashOut,
        totalTransactions: totalTransactions,
        lastTransactionDate: lastTransactionDate ? lastTransactionDate.toISOString() : null,

        // Party Stats
        totalReceivable: totalReceivable,
        totalPayable: totalPayable,
        totalParties: totalParties,
        activeParties: activeParties,
        topReceivables: topReceivables,
        topPayables: topPayables,

        // User Stats
        totalUsers: totalUsers,
        activeUsers: activeUsers,

        // Net Position
        netPosition: totalCashBalance + totalReceivable - totalPayable
      }
    };

  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// ============== Parties Management ==============

// Get all parties
function getAllParties(currentUser, role) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTIES_SHEET);

    if (!sheet) {
      return { success: false, message: 'Parties sheet not found' };
    }

    var data = sheet.getDataRange().getValues();
    var parties = [];

    for (var i = 1; i < data.length; i++) {
      var createdAtValue = data[i][9];  // CreatedAt at column 10
      var updatedAtValue = data[i][11]; // UpdatedAt at column 12

      // Convert dates to ISO format if needed
      if (createdAtValue && !(typeof createdAtValue === 'string' && createdAtValue.includes('T'))) {
        createdAtValue = new Date(createdAtValue).toISOString();
      }
      if (updatedAtValue && !(typeof updatedAtValue === 'string' && updatedAtValue.includes('T'))) {
        updatedAtValue = new Date(updatedAtValue).toISOString();
      }

      parties.push({
        PartyID: data[i][0],
        PartyName: data[i][1],
        PartyType: data[i][2],
        ContactNumber: data[i][3] || '',
        Address: data[i][4] || '',
        OpeningBalance: data[i][5] || 0,
        BalanceType: data[i][6],
        CurrentBalance: data[i][7] || 0,
        IsActive: data[i][8],
        CreatedAt: createdAtValue,
        CreatedBy: data[i][10],
        UpdatedAt: updatedAtValue,
        UpdatedBy: data[i][12]
      });
    }

    addLog(currentUser, 'Parties Viewed', 'Viewed parties list');

    return { success: true, data: parties };

  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// Add new party
function addParty(partyData, currentUser) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTIES_SHEET);

    if (!sheet) {
      return { success: false, message: 'Parties sheet not found' };
    }

    // Check if party name already exists
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1].toLowerCase() === partyData.PartyName.toLowerCase()) {
        return { success: false, message: 'Party name already exists' };
      }
    }

    // Generate new Party ID
    var newId = data.length; // Simple incremental ID

    var timestamp = new Date().toISOString();

    // Calculate current balance
    var currentBalance = parseFloat(partyData.OpeningBalance) || 0;

    // Add new party
    sheet.appendRow([
      newId,                                    // PartyID
      partyData.PartyName,                      // PartyName
      partyData.PartyType || 'Customer',        // PartyType
      partyData.ContactNumber || '',            // ContactNumber
      partyData.Address || '',                  // Address
      parseFloat(partyData.OpeningBalance) || 0,// OpeningBalance
      partyData.BalanceType || 'Debit',         // BalanceType
      currentBalance,                           // CurrentBalance
      partyData.IsActive !== false,             // IsActive (default true)
      timestamp,                                // CreatedAt
      currentUser,                              // CreatedBy
      timestamp,                                // UpdatedAt
      currentUser                               // UpdatedBy
    ]);

    addLog(currentUser, 'Party Added', 'Added party: ' + partyData.PartyName);

    return { success: true, message: 'Party added successfully!' };

  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// Update party
function updateParty(partyId, partyData, currentUser) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTIES_SHEET);

    if (!sheet) {
      return { success: false, message: 'Parties sheet not found' };
    }

    var data = sheet.getDataRange().getValues();

    // Check if party name already exists (excluding current party)
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] !== partyId && data[i][1].toLowerCase() === partyData.PartyName.toLowerCase()) {
        return { success: false, message: 'Party name already exists' };
      }
    }

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === partyId) {
        var timestamp = new Date().toISOString();

        // Update party data
        sheet.getRange(i + 1, 2).setValue(partyData.PartyName);
        sheet.getRange(i + 1, 3).setValue(partyData.PartyType);
        sheet.getRange(i + 1, 4).setValue(partyData.ContactNumber || '');
        sheet.getRange(i + 1, 5).setValue(partyData.Address || '');
        sheet.getRange(i + 1, 6).setValue(parseFloat(partyData.OpeningBalance) || 0);
        sheet.getRange(i + 1, 7).setValue(partyData.BalanceType);
        sheet.getRange(i + 1, 8).setValue(parseFloat(partyData.CurrentBalance) || 0);
        sheet.getRange(i + 1, 9).setValue(partyData.IsActive !== false);
        sheet.getRange(i + 1, 12).setValue(timestamp); // UpdatedAt
        sheet.getRange(i + 1, 13).setValue(currentUser); // UpdatedBy

        addLog(currentUser, 'Party Updated', 'Updated party: ' + partyData.PartyName);

        return { success: true, message: 'Party updated successfully!' };
      }
    }

    return { success: false, message: 'Party not found' };

  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// Delete party
function deleteParty(partyId, currentUser) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTIES_SHEET);

    if (!sheet) {
      return { success: false, message: 'Parties sheet not found' };
    }

    // CRITICAL FIX: Check for existing references before deletion
    var cashBookSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CASH_BOOK_SHEET);
    var partyLedgerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTY_LEDGER_SHEET);

    // FIX: Convert PartyID to number for type-safe comparisons
    var partyIdNum = parseInt(partyId, 10);

    var hasReferences = false;
    var referenceCount = 0;

    // Check CashBook references
    if (cashBookSheet) {
      var cashBookData = cashBookSheet.getDataRange().getValues();
      for (var j = 1; j < cashBookData.length; j++) {
        if (parseInt(cashBookData[j][7], 10) === partyIdNum) { // PartyID column
          hasReferences = true;
          referenceCount++;
        }
      }
    }

    // Check PartyLedger references
    if (partyLedgerSheet) {
      var ledgerData = partyLedgerSheet.getDataRange().getValues();
      for (var k = 1; k < ledgerData.length; k++) {
        if (parseInt(ledgerData[k][2], 10) === partyIdNum) { // PartyID column
          hasReferences = true;
          referenceCount++;
        }
      }
    }

    // Prevent deletion if references exist
    if (hasReferences) {
      addLog(currentUser, 'Party Delete Failed', 'Cannot delete party with ' + referenceCount + ' transaction(s)');
      return {
        success: false,
        message: 'Cannot delete party with existing transactions (' + referenceCount + ' found). Please delete all related transactions first or mark party as inactive instead.'
      };
    }

    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (parseInt(data[i][0], 10) === partyIdNum) {
        var partyName = data[i][1];
        sheet.deleteRow(i + 1);
        addLog(currentUser, 'Party Deleted', 'Deleted party: ' + partyName);
        return { success: true, message: 'Party deleted successfully!' };
      }
    }

    return { success: false, message: 'Party not found' };

  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// ============== Cash Book Management ==============

// Generate voucher number using sequence table (VC-YYYYMMDD-XXX)
function generateVoucherNumber(date) {
  try {
    var sequenceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VOUCHER_SEQUENCE_SHEET);

    if (!sequenceSheet) {
      var dateStr = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'yyyyMMdd');
      return 'VC-' + dateStr + '-001';
    }

    var transactionDate = new Date(date);
    var dateStr = Utilities.formatDate(transactionDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var dateStrFormatted = Utilities.formatDate(transactionDate, Session.getScriptTimeZone(), 'yyyyMMdd');

    var data = sequenceSheet.getDataRange().getValues();
    var rowIndex = -1;
    var lastNumber = 0;

    // Find existing sequence for this date
    for (var i = 1; i < data.length; i++) {
      var seqDateStr = Utilities.formatDate(new Date(data[i][1]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (seqDateStr === dateStr) {
        rowIndex = i;
        lastNumber = parseInt(data[i][2]) || 0;
        break;
      }
    }

    var newNumber = lastNumber + 1;

    if (rowIndex !== -1) {
      // Update existing sequence
      sequenceSheet.getRange(rowIndex + 1, 3).setValue(newNumber);
    } else {
      // Create new sequence entry
      var newId = data.length > 1 ? sequenceSheet.getRange(data.length, 1).getValue() + 1 : 1;
      var timestamp = new Date().toISOString();
      sequenceSheet.appendRow([
        newId,
        transactionDate.toISOString(),
        newNumber,
        timestamp
      ]);
    }

    var voucherNumber = 'VC-' + dateStrFormatted + '-' + ('000' + newNumber).slice(-3);
    return voucherNumber;

  } catch (error) {
    Logger.log('Error generating voucher number: ' + error.toString());
    var dateStr = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'yyyyMMdd');
    return 'VC-' + dateStr + '-001';
  }
}

// Get all cash book entries
function getCashBookEntries(currentUser, role) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CASH_BOOK_SHEET);

    if (!sheet) {
      return { success: false, message: 'Cash Book sheet not found' };
    }

    var data = sheet.getDataRange().getValues();
    var entries = [];

    for (var i = 1; i < data.length; i++) {
      var transactionDateValue = data[i][2];  // TransactionDate
      var createdAtValue = data[i][14];       // CreatedAt
      var updatedAtValue = data[i][15];       // UpdatedAt

      // Convert dates to ISO format if needed
      if (transactionDateValue && !(typeof transactionDateValue === 'string' && transactionDateValue.includes('T'))) {
        transactionDateValue = new Date(transactionDateValue).toISOString();
      }
      if (createdAtValue && !(typeof createdAtValue === 'string' && createdAtValue.includes('T'))) {
        createdAtValue = new Date(createdAtValue).toISOString();
      }
      if (updatedAtValue && !(typeof updatedAtValue === 'string' && updatedAtValue.includes('T'))) {
        updatedAtValue = new Date(updatedAtValue).toISOString();
      }

      entries.push({
        CashID: data[i][0],
        VoucherNo: data[i][1],
        TransactionDate: transactionDateValue,
        Particulars: data[i][3],
        CashIn: data[i][4] || 0,
        CashOut: data[i][5] || 0,
        EntryType: data[i][6],
        PartyID: data[i][7] || null,
        PartyName: data[i][8] || '',
        PaymentMode: data[i][9],
        ReferenceNo: data[i][10] || '',
        RunningBalance: data[i][11] || 0,
        LedgerLinked: data[i][12] || false,
        CreatedBy: data[i][13],
        CreatedAt: createdAtValue,
        UpdatedAt: updatedAtValue
      });
    }

    addLog(currentUser, 'Cash Book Viewed', 'Viewed cash book entries');

    return { success: true, data: entries };

  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// Add new cash book entry
function addCashBookEntry(entryData, currentUser) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CASH_BOOK_SHEET);

    if (!sheet) {
      return { success: false, message: 'Cash Book sheet not found' };
    }

    var data = sheet.getDataRange().getValues();

    // Generate new Cash ID
    var newId = data.length; // Simple incremental ID

    // Generate voucher number
    var voucherNo = generateVoucherNumber(entryData.TransactionDate);

    var timestamp = new Date().toISOString();
    var transactionDate = new Date(entryData.TransactionDate).toISOString();

    // Calculate running balance
    var previousBalance = 0;
    if (data.length > 1) {
      previousBalance = parseFloat(data[data.length - 1][11]) || 0; // Last running balance
    }

    var cashIn = parseFloat(entryData.CashIn) || 0;
    var cashOut = parseFloat(entryData.CashOut) || 0;
    var runningBalance = previousBalance + cashIn - cashOut;

    // Get party name if party selected and handle dual entry
    var partyName = '';
    var partyId = entryData.PartyID || null;
    var cashBookId = newId;

    if (partyId && (entryData.EntryType === 'Customer' || entryData.EntryType === 'Supplier')) {
      var partiesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTIES_SHEET);
      var ledgerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTY_LEDGER_SHEET);

      if (partiesSheet && ledgerSheet) {
        var partiesData = partiesSheet.getDataRange().getValues();
        // FIX: Convert PartyID to number for type-safe comparison
        var partyIdNum = parseInt(partyId, 10);

        for (var i = 1; i < partiesData.length; i++) {
          if (parseInt(partiesData[i][0], 10) === partyIdNum) {
            partyName = partiesData[i][1];

            // ===== DUAL ENTRY LOGIC =====
            // Determine debit/credit amounts for party ledger
            var debitAmount = 0;
            var creditAmount = 0;
            var transactionType = '';
            var narration = '';

            if (entryData.EntryType === 'Customer') {
              if (cashIn > 0) {
                // Customer paid us - CREDIT entry (reduces receivable)
                creditAmount = cashIn;
                debitAmount = 0;
                transactionType = 'Receipt';
                narration = 'Cash payment received';
              } else if (cashOut > 0) {
                // Refund to customer - DEBIT entry (increases receivable)
                debitAmount = cashOut;
                creditAmount = 0;
                transactionType = 'Refund';
                narration = 'Cash refund given';
              }
            } else if (entryData.EntryType === 'Supplier') {
              if (cashOut > 0) {
                // Paid to supplier - DEBIT entry (reduces payable - which is negative)
                debitAmount = cashOut;
                creditAmount = 0;
                transactionType = 'Payment';
                narration = 'Cash payment made';
              } else if (cashIn > 0) {
                // Refund from supplier - CREDIT entry (increases payable)
                creditAmount = cashIn;
                debitAmount = 0;
                transactionType = 'Refund';
                narration = 'Cash refund received';
              }
            }

            // Create party ledger entry
            if (debitAmount > 0 || creditAmount > 0) {
              var ledgerData = ledgerSheet.getDataRange().getValues();
              var newLedgerId = ledgerData.length; // Simple incremental ID

              // Calculate running balance for party ledger
              var partyRunningBalance = 0;
              var hasExistingEntries = false;

              // Get last running balance for this party
              for (var j = ledgerData.length - 1; j >= 1; j--) {
                if (parseInt(ledgerData[j][2], 10) === partyIdNum) {
                  partyRunningBalance = parseFloat(ledgerData[j][11]) || 0;
                  hasExistingEntries = true;
                  break;
                }
              }

              // If no previous entries exist, use opening balance from parties table
              if (!hasExistingEntries) {
                var openingBalance = parseFloat(partiesData[i][5]) || 0; // OpeningBalance
                partyRunningBalance = openingBalance;
              }

              // Calculate new running balance: Balance + Debit - Credit
              var newPartyRunningBalance = partyRunningBalance + debitAmount - creditAmount;

              // Insert party ledger entry
              ledgerSheet.appendRow([
                newLedgerId,                          // LedgerID
                transactionDate,                      // TransactionDate
                partyId,                              // PartyID
                partyName,                            // PartyName
                entryData.InvoiceNo || '',            // InvoiceNo
                debitAmount,                          // DebitAmount
                creditAmount,                         // CreditAmount
                transactionType,                      // TransactionType
                entryData.PaymentMode || 'Cash',      // PaymentMode
                narration,                            // Narration
                entryData.ReferenceNo || '',          // ReferenceNo
                newPartyRunningBalance,               // RunningBalance
                cashBookId,                           // CashBookID (link)
                true,                                 // IsCashEntry
                currentUser,                          // CreatedBy
                timestamp,                            // CreatedAt
                timestamp                             // UpdatedAt
              ]);

              // FIX: Update party's current balance by syncing with ledger
              syncPartyBalanceFromLedger(partyId);
            }

            break;
          }
        }
      }
    }

    // Add new cash book entry
    sheet.appendRow([
      newId,                                    // CashID
      voucherNo,                                // VoucherNo
      transactionDate,                          // TransactionDate
      entryData.Particulars,                    // Particulars
      cashIn,                                   // CashIn
      cashOut,                                  // CashOut
      entryData.EntryType || 'Cash',            // EntryType
      partyId,                                  // PartyID
      partyName,                                // PartyName
      entryData.PaymentMode || 'Cash',          // PaymentMode
      entryData.ReferenceNo || '',              // ReferenceNo
      runningBalance,                           // RunningBalance
      partyId ? true : false,                   // LedgerLinked
      currentUser,                              // CreatedBy
      timestamp,                                // CreatedAt
      timestamp                                 // UpdatedAt
    ]);

    addLog(currentUser, 'Cash Entry Added', 'Added cash entry: ' + voucherNo);

    return { success: true, message: 'Cash entry added successfully!', voucherNo: voucherNo };

  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// Update cash book entry
function updateCashBookEntry(cashId, entryData, currentUser) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CASH_BOOK_SHEET);

    if (!sheet) {
      return { success: false, message: 'Cash Book sheet not found' };
    }

    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === cashId) {
        var timestamp = new Date().toISOString();
        var transactionDate = new Date(entryData.TransactionDate).toISOString();

        // Store old values for party ledger update
        var oldCashIn = parseFloat(data[i][4]) || 0;
        var oldCashOut = parseFloat(data[i][5]) || 0;
        var oldEntryType = data[i][6];
        var oldPartyId = data[i][7];

        // Get party name if party selected
        var partyName = '';
        var partyId = entryData.PartyID || null;

        if (partyId && (entryData.EntryType === 'Customer' || entryData.EntryType === 'Supplier')) {
          var partiesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTIES_SHEET);
          if (partiesSheet) {
            var partiesData = partiesSheet.getDataRange().getValues();
            // FIX: Convert PartyID to number for type-safe comparison
            var partyIdNum = parseInt(partyId, 10);

            for (var j = 1; j < partiesData.length; j++) {
              if (parseInt(partiesData[j][0], 10) === partyIdNum) {
                partyName = partiesData[j][1];
                break;
              }
            }
          }
        }

        var newCashIn = parseFloat(entryData.CashIn) || 0;
        var newCashOut = parseFloat(entryData.CashOut) || 0;

        // Update corresponding party ledger entry if exists
        if (oldPartyId && (oldEntryType === 'Customer' || oldEntryType === 'Supplier')) {
          var ledgerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTY_LEDGER_SHEET);
          if (ledgerSheet) {
            var ledgerData = ledgerSheet.getDataRange().getValues();
            // Find ledger entry with matching CashBookID
            for (var k = 1; k < ledgerData.length; k++) {
              if (ledgerData[k][12] === cashId) { // CashBookID column
                // Calculate new debit/credit based on new amounts and entry type
                var newDebit = 0;
                var newCredit = 0;

                if (entryData.EntryType === 'Customer') {
                  if (newCashIn > 0) {
                    newCredit = newCashIn; // Customer payment received
                  } else if (newCashOut > 0) {
                    newDebit = newCashOut; // Customer refund
                  }
                } else if (entryData.EntryType === 'Supplier') {
                  if (newCashOut > 0) {
                    newDebit = newCashOut; // Supplier payment made
                  } else if (newCashIn > 0) {
                    newCredit = newCashIn; // Supplier refund received
                  }
                }

                // Update ledger entry
                ledgerSheet.getRange(k + 1, 2).setValue(transactionDate); // TransactionDate
                ledgerSheet.getRange(k + 1, 6).setValue(newDebit); // DebitAmount
                ledgerSheet.getRange(k + 1, 7).setValue(newCredit); // CreditAmount
                ledgerSheet.getRange(k + 1, 17).setValue(timestamp); // UpdatedAt

                // Recalculate party running balances
                recalculatePartyRunningBalances(oldPartyId);

                // FIX: Sync party balance from ledger
                syncPartyBalanceFromLedger(oldPartyId);
                break;
              }
            }
          }
        }

        // Update cash book entry data
        sheet.getRange(i + 1, 3).setValue(transactionDate);            // TransactionDate
        sheet.getRange(i + 1, 4).setValue(entryData.Particulars);      // Particulars
        sheet.getRange(i + 1, 5).setValue(newCashIn);   // CashIn
        sheet.getRange(i + 1, 6).setValue(newCashOut);  // CashOut
        sheet.getRange(i + 1, 7).setValue(entryData.EntryType);        // EntryType
        sheet.getRange(i + 1, 8).setValue(partyId);                    // PartyID
        sheet.getRange(i + 1, 9).setValue(partyName);                  // PartyName
        sheet.getRange(i + 1, 10).setValue(entryData.PaymentMode);     // PaymentMode
        sheet.getRange(i + 1, 11).setValue(entryData.ReferenceNo || ''); // ReferenceNo
        sheet.getRange(i + 1, 16).setValue(timestamp);                 // UpdatedAt

        // Recalculate cash book running balances
        recalculateRunningBalances();

        addLog(currentUser, 'Cash Entry Updated', 'Updated cash entry: ' + data[i][1]);

        return { success: true, message: 'Cash entry updated successfully!' };
      }
    }

    return { success: false, message: 'Cash entry not found' };

  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// Delete cash book entry
function deleteCashBookEntry(cashId, currentUser) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CASH_BOOK_SHEET);

    if (!sheet) {
      return { success: false, message: 'Cash Book sheet not found' };
    }

    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === cashId) {
        var voucherNo = data[i][1];
        var partyId = data[i][7]; // PartyID column

        // Delete corresponding party ledger entry if exists
        if (partyId) {
          var ledgerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTY_LEDGER_SHEET);
          if (ledgerSheet) {
            var ledgerData = ledgerSheet.getDataRange().getValues();
            // Find ledger entry with matching CashBookID
            for (var j = ledgerData.length - 1; j >= 1; j--) {
              if (ledgerData[j][12] === cashId) { // CashBookID column
                var ledgerDebit = parseFloat(ledgerData[j][5]) || 0;
                var ledgerCredit = parseFloat(ledgerData[j][6]) || 0;

                // Delete ledger row
                ledgerSheet.deleteRow(j + 1);

                // Recalculate party running balances
                recalculatePartyRunningBalances(partyId);

                // FIX: Sync party balance from ledger
                syncPartyBalanceFromLedger(partyId);
                break;
              }
            }
          }
        }

        // Delete cash book entry
        sheet.deleteRow(i + 1);

        // Recalculate running balances
        recalculateRunningBalances();

        addLog(currentUser, 'Cash Entry Deleted', 'Deleted cash entry: ' + voucherNo);
        return { success: true, message: 'Cash entry deleted successfully!' };
      }
    }

    return { success: false, message: 'Cash entry not found' };

  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// Recalculate running balances for all entries
function recalculateRunningBalances() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CASH_BOOK_SHEET);

    if (!sheet) return;

    var data = sheet.getDataRange().getValues();

    // CRITICAL FIX: Sort entries by date before calculating running balance
    var entries = [];
    for (var i = 1; i < data.length; i++) {
      entries.push({
        rowIndex: i,
        date: new Date(data[i][2]), // TransactionDate
        cashIn: parseFloat(data[i][4]) || 0,
        cashOut: parseFloat(data[i][5]) || 0,
        cashId: data[i][0] // Keep track for reference
      });
    }

    // Sort by date (oldest first)
    entries.sort(function(a, b) {
      return a.date - b.date;
    });

    // Recalculate running balances in chronological order
    var runningBalance = 0;
    for (var j = 0; j < entries.length; j++) {
      runningBalance = runningBalance + entries[j].cashIn - entries[j].cashOut;
      sheet.getRange(entries[j].rowIndex + 1, 12).setValue(runningBalance); // Update RunningBalance
    }

  } catch (error) {
    Logger.log('Error recalculating balances: ' + error.toString());
  }
}

// ============== Party Ledger Functions ==============

// Get all party ledger entries (optionally filter by party)
function getPartyLedgerEntries(partyId, currentUser, role) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTY_LEDGER_SHEET);

    if (!sheet) {
      return { success: false, message: 'Party Ledger sheet not found' };
    }

    var data = sheet.getDataRange().getValues();
    var ledgerEntries = [];

    for (var i = 1; i < data.length; i++) {
      var entry = {
        LedgerID: data[i][0],
        TransactionDate: data[i][1],
        PartyID: data[i][2],
        PartyName: data[i][3],
        InvoiceNo: data[i][4],
        DebitAmount: data[i][5],
        CreditAmount: data[i][6],
        TransactionType: data[i][7],
        PaymentMode: data[i][8],
        Narration: data[i][9],
        ReferenceNo: data[i][10],
        RunningBalance: data[i][11],
        CashBookID: data[i][12],
        IsCashEntry: data[i][13],
        CreatedBy: data[i][14],
        CreatedAt: data[i][15],
        UpdatedAt: data[i][16]
      };

      // Filter by party if partyId is provided
      if (!partyId || partyId === '' || entry.PartyID === partyId) {
        ledgerEntries.push(entry);
      }
    }

    return { success: true, data: ledgerEntries };

  } catch (error) {
    return { success: false, message: 'Error fetching ledger entries: ' + error.toString() };
  }
}

// Add new party ledger entry
function addPartyLedgerEntry(entryData, currentUser) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTY_LEDGER_SHEET);

    if (!sheet) {
      return { success: false, message: 'Party Ledger sheet not found' };
    }

    // Get ledger data (reused for invoice check and running balance calculation)
    var data = sheet.getDataRange().getValues();

    // FIX: Normalize invoice number to uppercase
    if (entryData.InvoiceNo && entryData.InvoiceNo.trim() !== '') {
      entryData.InvoiceNo = entryData.InvoiceNo.trim().toUpperCase();
    }

    // ===== DUPLICATE INVOICE CHECK =====
    // Check if invoice number already exists (if provided)
    if (entryData.InvoiceNo && entryData.InvoiceNo.trim() !== '') {
      for (var invoiceIdx = 1; invoiceIdx < data.length; invoiceIdx++) {
        var existingInvoice = data[invoiceIdx][4]; // InvoiceNo column
        if (existingInvoice && existingInvoice.toString().trim().toUpperCase() === entryData.InvoiceNo) {
          return {
            success: false,
            message: 'Duplicate Invoice Number: "' + entryData.InvoiceNo + '" already exists!'
          };
        }
      }
    }

    // Get next ID
    var newId = data.length; // Simple incremental ID based on row count

    // FIX: Get and validate party name (with type conversion)
    var partiesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTIES_SHEET);
    if (!partiesSheet) {
      return { success: false, message: 'Parties sheet not found' };
    }

    var partiesData = partiesSheet.getDataRange().getValues();
    var partyName = '';
    var partyFound = false;

    // Convert PartyID to number for comparison (handles both string and number inputs)
    var partyIdToFind = parseInt(entryData.PartyID, 10);

    for (var partyIdx = 1; partyIdx < partiesData.length; partyIdx++) {
      var sheetPartyId = parseInt(partiesData[partyIdx][0], 10);
      if (sheetPartyId === partyIdToFind) {
        partyName = partiesData[partyIdx][1];
        partyFound = true;
        break;
      }
    }

    // Validate party exists
    if (!partyFound || !partyName) {
      return {
        success: false,
        message: 'Invalid Party: Party ID "' + entryData.PartyID + '" not found. Please select a valid party.'
      };
    }

    var timestamp = new Date().toISOString();
    var transactionDate = new Date(entryData.TransactionDate).toISOString();

    // Calculate running balance for this party (reusing data from above)
    var runningBalance = 0;
    var hasExistingEntries = false;

    // Check for existing entries and get last running balance
    for (var ledgerIdx = data.length - 1; ledgerIdx >= 1; ledgerIdx--) {
      if (data[ledgerIdx][2] === entryData.PartyID) {
        runningBalance = parseFloat(data[ledgerIdx][11]) || 0; // Get last RunningBalance
        hasExistingEntries = true;
        break;
      }
    }

    // If no previous entries exist, use opening balance from parties table
    if (!hasExistingEntries) {
      for (var openingIdx = 1; openingIdx < partiesData.length; openingIdx++) {
        if (partiesData[openingIdx][0] === entryData.PartyID) {
          runningBalance = parseFloat(partiesData[openingIdx][5]) || 0; // OpeningBalance
          break;
        }
      }
    }

    // Add new entry's debit and credit to calculate new running balance
    var debit = parseFloat(entryData.DebitAmount) || 0;
    var credit = parseFloat(entryData.CreditAmount) || 0;
    runningBalance = runningBalance + debit - credit;

    sheet.appendRow([
      newId,
      transactionDate,
      entryData.PartyID,
      partyName,
      entryData.InvoiceNo || '',
      debit,
      credit,
      entryData.TransactionType,
      entryData.PaymentMode,
      entryData.Narration || '',
      entryData.ReferenceNo || '',
      runningBalance,
      entryData.CashBookID || '',
      entryData.IsCashEntry || false,
      currentUser || 'Admin',
      timestamp,
      timestamp
    ]);

    // FIX: Sync party's current balance from ledger
    syncPartyBalanceFromLedger(entryData.PartyID);

    // Log activity
    addLog(currentUser || 'Admin', 'Add Ledger Entry', 'Added ledger entry #' + newId + ' for party: ' + partyName);

    return { success: true, message: 'Ledger entry added successfully', ledgerId: newId };

  } catch (error) {
    return { success: false, message: 'Error adding ledger entry: ' + error.toString() };
  }
}

// Update party ledger entry
function updatePartyLedgerEntry(ledgerId, entryData, currentUser) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTY_LEDGER_SHEET);

    if (!sheet) {
      return { success: false, message: 'Party Ledger sheet not found' };
    }

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    var oldDebit = 0;
    var oldCredit = 0;
    var oldPartyId = 0;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === ledgerId) {
        rowIndex = i;
        oldDebit = parseFloat(data[i][5]) || 0;
        oldCredit = parseFloat(data[i][6]) || 0;
        oldPartyId = data[i][2];
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: 'Ledger entry not found' };
    }

    // FIX: Normalize invoice number to uppercase
    if (entryData.InvoiceNo && entryData.InvoiceNo.trim() !== '') {
      entryData.InvoiceNo = entryData.InvoiceNo.trim().toUpperCase();
    }

    // FIX: Get and validate party name (with type conversion)
    var partiesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTIES_SHEET);
    if (!partiesSheet) {
      return { success: false, message: 'Parties sheet not found' };
    }

    var partiesData = partiesSheet.getDataRange().getValues();
    var partyName = '';
    var partyFound = false;

    // Convert PartyID to number for comparison (handles both string and number inputs)
    var partyIdToFind = parseInt(entryData.PartyID, 10);

    for (var i = 1; i < partiesData.length; i++) {
      var sheetPartyId = parseInt(partiesData[i][0], 10);
      if (sheetPartyId === partyIdToFind) {
        partyName = partiesData[i][1];
        partyFound = true;
        break;
      }
    }

    // Validate party exists
    if (!partyFound || !partyName) {
      return {
        success: false,
        message: 'Invalid Party: Party ID "' + entryData.PartyID + '" not found. Please select a valid party.'
      };
    }

    var timestamp = new Date().toISOString();
    var transactionDate = new Date(entryData.TransactionDate).toISOString();
    var debit = parseFloat(entryData.DebitAmount) || 0;
    var credit = parseFloat(entryData.CreditAmount) || 0;

    // Update the row
    sheet.getRange(rowIndex + 1, 2).setValue(transactionDate);
    sheet.getRange(rowIndex + 1, 3).setValue(entryData.PartyID);
    sheet.getRange(rowIndex + 1, 4).setValue(partyName);
    sheet.getRange(rowIndex + 1, 5).setValue(entryData.InvoiceNo || '');
    sheet.getRange(rowIndex + 1, 6).setValue(debit);
    sheet.getRange(rowIndex + 1, 7).setValue(credit);
    sheet.getRange(rowIndex + 1, 8).setValue(entryData.TransactionType);
    sheet.getRange(rowIndex + 1, 9).setValue(entryData.PaymentMode);
    sheet.getRange(rowIndex + 1, 10).setValue(entryData.Narration || '');
    sheet.getRange(rowIndex + 1, 11).setValue(entryData.ReferenceNo || '');
    sheet.getRange(rowIndex + 1, 17).setValue(timestamp);

    // Reverse old party balance update if party changed
    if (oldPartyId !== entryData.PartyID) {
      updatePartyBalance(oldPartyId, -oldDebit, -oldCredit, currentUser || 'Admin');
      updatePartyBalance(entryData.PartyID, debit, credit, currentUser || 'Admin');
    } else {
      // Update with the difference
      updatePartyBalance(entryData.PartyID, debit - oldDebit, credit - oldCredit, currentUser || 'Admin');
    }

    // Recalculate running balances for affected parties
    recalculatePartyRunningBalances(entryData.PartyID);
    if (oldPartyId !== entryData.PartyID) {
      recalculatePartyRunningBalances(oldPartyId);
    }

    // FIX: Sync party balances from ledger
    syncPartyBalanceFromLedger(entryData.PartyID);
    if (oldPartyId !== entryData.PartyID) {
      syncPartyBalanceFromLedger(oldPartyId);
    }

    // Log activity
    addLog(currentUser || 'Admin', 'Update Ledger Entry', 'Updated ledger entry #' + ledgerId);

    return { success: true, message: 'Ledger entry updated successfully' };

  } catch (error) {
    return { success: false, message: 'Error updating ledger entry: ' + error.toString() };
  }
}

// Delete party ledger entry
function deletePartyLedgerEntry(ledgerId, currentUser) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTY_LEDGER_SHEET);

    if (!sheet) {
      return { success: false, message: 'Party Ledger sheet not found' };
    }

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    var debit = 0;
    var credit = 0;
    var partyId = 0;
    var partyName = '';

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === ledgerId) {
        rowIndex = i;
        partyId = data[i][2];
        partyName = data[i][3];
        debit = parseFloat(data[i][5]) || 0;
        credit = parseFloat(data[i][6]) || 0;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: 'Ledger entry not found' };
    }

    // Delete the row
    sheet.deleteRow(rowIndex + 1);

    // Recalculate running balances for this party
    recalculatePartyRunningBalances(partyId);

    // FIX: Sync party balance from ledger after deletion
    syncPartyBalanceFromLedger(partyId);

    // Log activity
    addLog(currentUser || 'Admin', 'Delete Ledger Entry', 'Deleted ledger entry #' + ledgerId + ' for party: ' + partyName);

    return { success: true, message: 'Ledger entry deleted successfully' };

  } catch (error) {
    return { success: false, message: 'Error deleting ledger entry: ' + error.toString() };
  }
}

// Recalculate running balances for a specific party
function recalculatePartyRunningBalances(partyId) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTY_LEDGER_SHEET);

    if (!sheet) return;

    var data = sheet.getDataRange().getValues();

    // Fetch party's opening balance from parties table
    var partiesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTIES_SHEET);
    var partiesData = partiesSheet.getDataRange().getValues();
    var runningBalance = 0;

    for (var j = 1; j < partiesData.length; j++) {
      if (partiesData[j][0] === partyId) {
        runningBalance = parseFloat(partiesData[j][5]) || 0; // OpeningBalance column
        break;
      }
    }

    // Sort entries by date for this party
    var partyEntries = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][2] === partyId) {
        partyEntries.push({
          rowIndex: i,
          date: new Date(data[i][1]),
          debit: parseFloat(data[i][5]) || 0,
          credit: parseFloat(data[i][6]) || 0
        });
      }
    }

    // Sort by date
    partyEntries.sort(function(a, b) {
      return a.date - b.date;
    });

    // Recalculate running balances starting from opening balance
    for (var i = 0; i < partyEntries.length; i++) {
      runningBalance = runningBalance + partyEntries[i].debit - partyEntries[i].credit;
      sheet.getRange(partyEntries[i].rowIndex + 1, 12).setValue(runningBalance);
    }

  } catch (error) {
    Logger.log('Error recalculating party running balances: ' + error.toString());
  }
}

// Update party's current balance
function updatePartyBalance(partyId, debitChange, creditChange, updatedBy) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTIES_SHEET);

    if (!sheet) return;

    var data = sheet.getDataRange().getValues();
    var timestamp = new Date().toISOString();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === partyId) {
        var currentBalance = parseFloat(data[i][7]) || 0;
        var balanceChange = debitChange - creditChange;
        var newBalance = currentBalance + balanceChange;

        sheet.getRange(i + 1, 8).setValue(newBalance); // Update CurrentBalance
        sheet.getRange(i + 1, 12).setValue(timestamp); // Update UpdatedAt (Column 12, Index 11)
        sheet.getRange(i + 1, 13).setValue(updatedBy || 'System'); // Update UpdatedBy (Column 13, Index 12)
        break;
      }
    }

  } catch (error) {
    Logger.log('Error updating party balance: ' + error.toString());
  }
}

// FIX: Sync party CurrentBalance with latest ledger RunningBalance
// This ensures consistency between Parties and PartyLedger sheets
function syncPartyBalanceFromLedger(partyId) {
  try {
    var partiesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTIES_SHEET);
    var ledgerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTY_LEDGER_SHEET);

    if (!partiesSheet || !ledgerSheet) return;

    var ledgerData = ledgerSheet.getDataRange().getValues();
    var partiesData = partiesSheet.getDataRange().getValues();

    // Find the latest running balance for this party from ledger
    var latestRunningBalance = null;
    var latestDate = null;

    for (var i = 1; i < ledgerData.length; i++) {
      if (ledgerData[i][2] === partyId) { // PartyID column
        var entryDate = new Date(ledgerData[i][1]); // TransactionDate
        var runningBalance = parseFloat(ledgerData[i][11]) || 0; // RunningBalance

        if (latestDate === null || entryDate >= latestDate) {
          latestDate = entryDate;
          latestRunningBalance = runningBalance;
        }
      }
    }

    // If no ledger entries found, use opening balance
    if (latestRunningBalance === null) {
      for (var j = 1; j < partiesData.length; j++) {
        if (partiesData[j][0] === partyId) {
          latestRunningBalance = parseFloat(partiesData[j][5]) || 0; // OpeningBalance
          break;
        }
      }
    }

    // Update CurrentBalance in Parties sheet
    for (var k = 1; k < partiesData.length; k++) {
      if (partiesData[k][0] === partyId) {
        var timestamp = new Date().toISOString();
        partiesSheet.getRange(k + 1, 8).setValue(latestRunningBalance); // CurrentBalance
        partiesSheet.getRange(k + 1, 12).setValue(timestamp); // UpdatedAt
        break;
      }
    }

  } catch (error) {
    Logger.log('Error syncing party balance from ledger: ' + error.toString());
  }
}

// ============== RECONCILIATION FUNCTIONS ==============

// Reconcile all party balances and fix missing party names
function reconcileAllParties(currentUser) {
  try {
    var partiesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTIES_SHEET);
    var ledgerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTY_LEDGER_SHEET);

    if (!partiesSheet || !ledgerSheet) {
      return { success: false, message: 'Required sheets not found' };
    }

    var partiesData = partiesSheet.getDataRange().getValues();
    var ledgerData = ledgerSheet.getDataRange().getValues();

    var reconciliationReport = [];
    var totalPartiesProcessed = 0;
    var totalBalancesFixed = 0;
    var totalNamesFixed = 0;

    // Process each party
    for (var i = 1; i < partiesData.length; i++) {
      var partyId = partiesData[i][0];
      var partyName = partiesData[i][1];
      var currentBalance = parseFloat(partiesData[i][7]) || 0;
      var openingBalance = parseFloat(partiesData[i][5]) || 0;

      totalPartiesProcessed++;

      // ===== FIX MISSING PARTY NAMES IN LEDGER =====
      var namesFixed = 0;
      for (var j = 1; j < ledgerData.length; j++) {
        if (ledgerData[j][2] === partyId) {
          var ledgerPartyName = ledgerData[j][3];
          // If party name is empty or doesn't match, update it
          if (!ledgerPartyName || ledgerPartyName.trim() === '' || ledgerPartyName !== partyName) {
            ledgerSheet.getRange(j + 1, 4).setValue(partyName); // PartyName column
            namesFixed++;
            totalNamesFixed++;
          }
        }
      }

      // ===== RECALCULATE PARTY RUNNING BALANCES =====
      recalculatePartyRunningBalances(partyId);

      // ===== CALCULATE CORRECT BALANCE FROM LEDGER =====
      var calculatedBalance = openingBalance;
      var transactionCount = 0;

      // Collect all transactions for this party
      var partyTransactions = [];
      for (var k = 1; k < ledgerData.length; k++) {
        if (ledgerData[k][2] === partyId) {
          partyTransactions.push({
            date: new Date(ledgerData[k][1]),
            debit: parseFloat(ledgerData[k][5]) || 0,
            credit: parseFloat(ledgerData[k][6]) || 0
          });
        }
      }

      // Sort by date
      partyTransactions.sort(function(a, b) {
        return a.date - b.date;
      });

      // Calculate final balance
      for (var t = 0; t < partyTransactions.length; t++) {
        calculatedBalance += partyTransactions[t].debit - partyTransactions[t].credit;
        transactionCount++;
      }

      // ===== UPDATE CURRENT BALANCE IF DIFFERENT =====
      var balanceFixed = false;
      if (Math.abs(currentBalance - calculatedBalance) > 0.01) { // Allow for small floating point differences
        partiesSheet.getRange(i + 1, 8).setValue(calculatedBalance); // CurrentBalance
        partiesSheet.getRange(i + 1, 12).setValue(new Date().toISOString()); // UpdatedAt
        partiesSheet.getRange(i + 1, 13).setValue(currentUser || 'System'); // UpdatedBy
        balanceFixed = true;
        totalBalancesFixed++;
      }

      // Add to report if anything was fixed
      if (balanceFixed || namesFixed > 0) {
        reconciliationReport.push({
          PartyID: partyId,
          PartyName: partyName,
          OldBalance: currentBalance,
          NewBalance: calculatedBalance,
          BalanceDifference: calculatedBalance - currentBalance,
          TransactionCount: transactionCount,
          NamesFixed: namesFixed,
          Status: balanceFixed ? 'Balance Updated' : 'Names Fixed'
        });
      }
    }

    // Log the reconciliation
    addLog(currentUser || 'System', 'Reconciliation Complete',
      'Processed ' + totalPartiesProcessed + ' parties. Fixed ' + totalBalancesFixed + ' balances and ' + totalNamesFixed + ' party names.');

    return {
      success: true,
      message: 'Reconciliation completed successfully!',
      summary: {
        totalPartiesProcessed: totalPartiesProcessed,
        totalBalancesFixed: totalBalancesFixed,
        totalNamesFixed: totalNamesFixed
      },
      details: reconciliationReport
    };

  } catch (error) {
    return { success: false, message: 'Reconciliation error: ' + error.toString() };
  }
}

// Reconcile single party
function reconcileSingleParty(partyId, currentUser) {
  try {
    var partiesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTIES_SHEET);
    var ledgerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTY_LEDGER_SHEET);

    if (!partiesSheet || !ledgerSheet) {
      return { success: false, message: 'Required sheets not found' };
    }

    var partiesData = partiesSheet.getDataRange().getValues();
    var ledgerData = ledgerSheet.getDataRange().getValues();

    // Find party
    var partyIndex = -1;
    var partyName = '';
    var openingBalance = 0;

    for (var i = 1; i < partiesData.length; i++) {
      if (partiesData[i][0] === partyId) {
        partyIndex = i;
        partyName = partiesData[i][1];
        openingBalance = parseFloat(partiesData[i][5]) || 0;
        break;
      }
    }

    if (partyIndex === -1) {
      return { success: false, message: 'Party not found' };
    }

    // Fix missing party names in ledger
    var namesFixed = 0;
    for (var j = 1; j < ledgerData.length; j++) {
      if (ledgerData[j][2] === partyId) {
        var ledgerPartyName = ledgerData[j][3];
        if (!ledgerPartyName || ledgerPartyName.trim() === '') {
          ledgerSheet.getRange(j + 1, 4).setValue(partyName);
          namesFixed++;
        }
      }
    }

    // Recalculate running balances
    recalculatePartyRunningBalances(partyId);

    // Sync balance from ledger
    syncPartyBalanceFromLedger(partyId);

    addLog(currentUser || 'System', 'Party Reconciled',
      'Reconciled party: ' + partyName + ' (Fixed ' + namesFixed + ' names)');

    return {
      success: true,
      message: 'Party reconciled successfully! Fixed ' + namesFixed + ' party name(s).',
      partyName: partyName,
      namesFixed: namesFixed
    };

  } catch (error) {
    return { success: false, message: 'Reconciliation error: ' + error.toString() };
  }
}

// Get party transaction history with date filters
function getPartyTransactionHistory(partyId, dateFilter) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTY_LEDGER_SHEET);

    if (!sheet) {
      return { success: false, message: 'Party Ledger sheet not found' };
    }

    var data = sheet.getDataRange().getValues();
    var transactions = [];
    var totalCredit = 0;
    var totalDebit = 0;

    // Calculate date range based on filter
    var now = new Date();
    var filterDate = new Date();

    switch(dateFilter) {
      case '30days':
        filterDate.setDate(now.getDate() - 30);
        break;
      case '6months':
        filterDate.setMonth(now.getMonth() - 6);
        break;
      case '1year':
        filterDate.setFullYear(now.getFullYear() - 1);
        break;
      case 'lifetime':
        filterDate = new Date(0); // Epoch time (include all)
        break;
      default:
        filterDate.setDate(now.getDate() - 30); // Default to 30 days
    }

    // Loop through all ledger entries
    for (var i = 1; i < data.length; i++) {
      var entry = {
        LedgerID: data[i][0],
        TransactionDate: data[i][1],
        PartyID: data[i][2],
        PartyName: data[i][3],
        InvoiceNo: data[i][4] || 'N/A',
        DebitAmount: parseFloat(data[i][5]) || 0,
        CreditAmount: parseFloat(data[i][6]) || 0,
        TransactionType: data[i][7] || 'N/A',
        PaymentMode: data[i][8] || 'N/A',
        Narration: data[i][9] || '',
        ReferenceNo: data[i][10] || '',
        RunningBalance: parseFloat(data[i][11]) || 0,
        CashBookID: data[i][12] || '',
        IsCashEntry: data[i][13] || false,
        CreatedBy: data[i][14] || '',
        CreatedAt: data[i][15] || '',
        UpdatedAt: data[i][16] || ''
      };

      // Filter by partyId
      if (entry.PartyID === partyId) {
        // Parse transaction date
        var transactionDate = new Date(entry.TransactionDate);

        // Filter by date range
        if (transactionDate >= filterDate) {
          // Convert dates to ISO format for frontend
          entry.TransactionDate = new Date(entry.TransactionDate).toISOString();
          entry.CreatedAt = entry.CreatedAt ? new Date(entry.CreatedAt).toISOString() : 'N/A';
          entry.UpdatedAt = entry.UpdatedAt ? new Date(entry.UpdatedAt).toISOString() : 'N/A';

          transactions.push(entry);

          // Calculate totals
          totalCredit += entry.CreditAmount;
          totalDebit += entry.DebitAmount;
        }
      }
    }

    // Calculate net amount (Debit - Credit for receivables/payables)
    var totalAmount = totalDebit - totalCredit;

    return {
      success: true,
      data: transactions,
      totalCredit: totalCredit,
      totalDebit: totalDebit,
      totalAmount: totalAmount
    };

  } catch (error) {
    return { success: false, message: 'Error fetching transaction history: ' + error.toString() };
  }
}

// ============== Daily Summary Functions ==============

// Get all daily summaries
function getDailySummaries(currentUser, role) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DAILY_SUMMARY_SHEET);

    if (!sheet) {
      return { success: false, message: 'Daily Summary sheet not found' };
    }

    var data = sheet.getDataRange().getValues();
    var summaries = [];

    for (var i = 1; i < data.length; i++) {
      summaries.push({
        SummaryID: data[i][0],
        SummaryDate: data[i][1],
        OpeningBalance: data[i][2],
        TotalCashIn: data[i][3],
        TotalCashOut: data[i][4],
        ClosingBalance: data[i][5],
        TotalTransactions: data[i][6],
        CreatedAt: data[i][7]
      });
    }

    return { success: true, data: summaries };

  } catch (error) {
    return { success: false, message: 'Error fetching daily summaries: ' + error.toString() };
  }
}

// Generate daily summary for a specific date
function generateDailySummary(date, currentUser) {
  try {
    var summarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DAILY_SUMMARY_SHEET);
    var cashBookSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CASH_BOOK_SHEET);

    if (!summarySheet || !cashBookSheet) {
      return { success: false, message: 'Required sheets not found' };
    }

    var targetDate = new Date(date);
    var dateStr = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    // Check if summary already exists for this date
    var summaryData = summarySheet.getDataRange().getValues();
    var existingRowIndex = -1;

    for (var i = 1; i < summaryData.length; i++) {
      var existingDateStr = Utilities.formatDate(new Date(summaryData[i][1]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (existingDateStr === dateStr) {
        existingRowIndex = i;
        break;
      }
    }

    // Get cash book data for this date
    var cashBookData = cashBookSheet.getDataRange().getValues();
    var totalCashIn = 0;
    var totalCashOut = 0;
    var transactionCount = 0;
    var openingBalance = 0;

    // Find opening balance (previous day's closing or first entry before this date)
    for (var i = 1; i < cashBookData.length; i++) {
      var transactionDate = new Date(cashBookData[i][2]);
      var transDateStr = Utilities.formatDate(transactionDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');

      if (transDateStr < dateStr) {
        // Get the running balance from previous transactions
        openingBalance = parseFloat(cashBookData[i][11]) || 0;
      } else if (transDateStr === dateStr) {
        // This is a transaction for the target date
        var cashIn = parseFloat(cashBookData[i][4]) || 0;
        var cashOut = parseFloat(cashBookData[i][5]) || 0;
        totalCashIn += cashIn;
        totalCashOut += cashOut;
        transactionCount++;
      }
    }

    var closingBalance = openingBalance + totalCashIn - totalCashOut;
    var timestamp = new Date().toISOString();

    if (existingRowIndex !== -1) {
      // Update existing summary
      summarySheet.getRange(existingRowIndex + 1, 3).setValue(openingBalance);
      summarySheet.getRange(existingRowIndex + 1, 4).setValue(totalCashIn);
      summarySheet.getRange(existingRowIndex + 1, 5).setValue(totalCashOut);
      summarySheet.getRange(existingRowIndex + 1, 6).setValue(closingBalance);
      summarySheet.getRange(existingRowIndex + 1, 7).setValue(transactionCount);
      summarySheet.getRange(existingRowIndex + 1, 8).setValue(timestamp);

      return { success: true, message: 'Daily summary updated for ' + dateStr };
    } else {
      // Create new summary
      var newId = summaryData.length > 1 ? summarySheet.getRange(summaryData.length, 1).getValue() + 1 : 1;

      summarySheet.appendRow([
        newId,
        targetDate.toISOString(),
        openingBalance,
        totalCashIn,
        totalCashOut,
        closingBalance,
        transactionCount,
        timestamp
      ]);

      return { success: true, message: 'Daily summary generated for ' + dateStr };
    }

  } catch (error) {
    return { success: false, message: 'Error generating daily summary: ' + error.toString() };
  }
}

// Regenerate all daily summaries from cash book data
function regenerateAllSummaries(currentUser) {
  try {
    var summarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DAILY_SUMMARY_SHEET);
    var cashBookSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CASH_BOOK_SHEET);

    if (!summarySheet || !cashBookSheet) {
      return { success: false, message: 'Required sheets not found' };
    }

    // Clear existing summaries (keep headers)
    var lastRow = summarySheet.getLastRow();
    if (lastRow > 1) {
      summarySheet.deleteRows(2, lastRow - 1);
    }

    // Get all cash book entries
    var cashBookData = cashBookSheet.getDataRange().getValues();

    // Group transactions by date
    var dateGroups = {};

    for (var i = 1; i < cashBookData.length; i++) {
      var transactionDate = new Date(cashBookData[i][2]);
      var dateStr = Utilities.formatDate(transactionDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');

      if (!dateGroups[dateStr]) {
        dateGroups[dateStr] = {
          cashIn: 0,
          cashOut: 0,
          count: 0
        };
      }

      var cashIn = parseFloat(cashBookData[i][4]) || 0;
      var cashOut = parseFloat(cashBookData[i][5]) || 0;
      dateGroups[dateStr].cashIn += cashIn;
      dateGroups[dateStr].cashOut += cashOut;
      dateGroups[dateStr].count++;
    }

    // Sort dates
    var sortedDates = Object.keys(dateGroups).sort();
    var runningBalance = 0;
    var summaryId = 1;

    // Generate summaries for each date
    for (var j = 0; j < sortedDates.length; j++) {
      var dateStr = sortedDates[j];
      var dayData = dateGroups[dateStr];

      var openingBalance = runningBalance;
      var closingBalance = openingBalance + dayData.cashIn - dayData.cashOut;
      runningBalance = closingBalance;

      var summaryDate = new Date(dateStr);
      var timestamp = new Date().toISOString();

      summarySheet.appendRow([
        summaryId++,
        summaryDate.toISOString(),
        openingBalance,
        dayData.cashIn,
        dayData.cashOut,
        closingBalance,
        dayData.count,
        timestamp
      ]);
    }

    // Log activity
    addLog(currentUser || 'Admin', 'Regenerate Summaries', 'Regenerated ' + sortedDates.length + ' daily summaries');

    return {
      success: true,
      message: 'Successfully regenerated ' + sortedDates.length + ' daily summaries',
      count: sortedDates.length
    };

  } catch (error) {
    return { success: false, message: 'Error regenerating summaries: ' + error.toString() };
  }
}

// Delete daily summary
function deleteDailySummary(summaryId, currentUser) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DAILY_SUMMARY_SHEET);

    if (!sheet) {
      return { success: false, message: 'Daily Summary sheet not found' };
    }

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    var summaryDate = '';

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === summaryId) {
        rowIndex = i;
        summaryDate = Utilities.formatDate(new Date(data[i][1]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: 'Summary not found' };
    }

    // Delete the row
    sheet.deleteRow(rowIndex + 1);

    // Log activity
    addLog(currentUser || 'Admin', 'Delete Daily Summary', 'Deleted summary for date: ' + summaryDate);

    return { success: true, message: 'Daily summary deleted successfully' };

  } catch (error) {
    return { success: false, message: 'Error deleting summary: ' + error.toString() };
  }
}

// ============== Voucher Sequence Functions ==============

// Get all voucher sequences
function getVoucherSequences(currentUser, role) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VOUCHER_SEQUENCE_SHEET);

    if (!sheet) {
      return { success: false, message: 'Voucher Sequence sheet not found' };
    }

    var data = sheet.getDataRange().getValues();
    var sequences = [];

    for (var i = 1; i < data.length; i++) {
      sequences.push({
        SequenceID: data[i][0],
        VoucherDate: data[i][1],
        LastNumber: data[i][2],
        CreatedAt: data[i][3]
      });
    }

    return { success: true, data: sequences };

  } catch (error) {
    return { success: false, message: 'Error fetching voucher sequences: ' + error.toString() };
  }
}

// Reset voucher sequence for a specific date
function resetVoucherSequence(sequenceId, currentUser) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VOUCHER_SEQUENCE_SHEET);

    if (!sheet) {
      return { success: false, message: 'Voucher Sequence sheet not found' };
    }

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    var voucherDate = '';

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === sequenceId) {
        rowIndex = i;
        voucherDate = Utilities.formatDate(new Date(data[i][1]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: 'Sequence not found' };
    }

    // Reset last number to 0
    sheet.getRange(rowIndex + 1, 3).setValue(0);

    // Log activity
    addLog(currentUser || 'Admin', 'Reset Voucher Sequence', 'Reset sequence for date: ' + voucherDate);

    return { success: true, message: 'Voucher sequence reset successfully for ' + voucherDate };

  } catch (error) {
    return { success: false, message: 'Error resetting sequence: ' + error.toString() };
  }
}

// Delete voucher sequence
function deleteVoucherSequence(sequenceId, currentUser) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VOUCHER_SEQUENCE_SHEET);

    if (!sheet) {
      return { success: false, message: 'Voucher Sequence sheet not found' };
    }

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    var voucherDate = '';

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === sequenceId) {
        rowIndex = i;
        voucherDate = Utilities.formatDate(new Date(data[i][1]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: 'Sequence not found' };
    }

    // Delete the row
    sheet.deleteRow(rowIndex + 1);

    // Log activity
    addLog(currentUser || 'Admin', 'Delete Voucher Sequence', 'Deleted sequence for date: ' + voucherDate);

    return { success: true, message: 'Voucher sequence deleted successfully' };

  } catch (error) {
    return { success: false, message: 'Error deleting sequence: ' + error.toString() };
  }
}

// Get next voucher number for a specific date (read-only, doesn't increment)
function getNextVoucherNumber(date) {
  try {
    var sequenceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VOUCHER_SEQUENCE_SHEET);

    if (!sequenceSheet) {
      var dateStr = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'yyyyMMdd');
      return { success: true, voucherNumber: 'VC-' + dateStr + '-001', nextNumber: 1 };
    }

    var transactionDate = new Date(date);
    var dateStr = Utilities.formatDate(transactionDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var dateStrFormatted = Utilities.formatDate(transactionDate, Session.getScriptTimeZone(), 'yyyyMMdd');

    var data = sequenceSheet.getDataRange().getValues();
    var lastNumber = 0;

    // Find existing sequence for this date
    for (var i = 1; i < data.length; i++) {
      var seqDateStr = Utilities.formatDate(new Date(data[i][1]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (seqDateStr === dateStr) {
        lastNumber = parseInt(data[i][2]) || 0;
        break;
      }
    }

    var nextNumber = lastNumber + 1;
    var voucherNumber = 'VC-' + dateStrFormatted + '-' + ('000' + nextNumber).slice(-3);

    return {
      success: true,
      voucherNumber: voucherNumber,
      nextNumber: nextNumber,
      lastNumber: lastNumber
    };

  } catch (error) {
    return { success: false, message: 'Error getting next voucher number: ' + error.toString() };
  }
}

// ============== System Settings Functions (Admin Only) ==============

// Get system settings
function getSystemSettings() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET);

    if (!sheet) {
      return {
        success: true,
        data: {
          currency: 'PKR',
          currencySymbol: 'Rs.',
          systemName: 'Cash Book System'
        }
      };
    }

    var data = sheet.getDataRange().getValues();
    var settings = {};

    for (var i = 1; i < data.length; i++) {
      var key = data[i][0];
      var value = data[i][1];
      settings[key] = value;
    }

    return { success: true, data: settings };

  } catch (error) {
    return { success: false, message: 'Error fetching settings: ' + error.toString() };
  }
}

// Update system settings (Admin only)
function updateSystemSettings(settingsData, currentUser, role) {
  try {
    // Check if user is admin
    if (role !== 'Admin') {
      return { success: false, message: 'Access Denied: Only Admin can update system settings!' };
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET);

    if (!sheet) {
      return { success: false, message: 'Settings sheet not found' };
    }

    var data = sheet.getDataRange().getValues();
    var timestamp = new Date().toISOString();

    // Update each setting
    for (var key in settingsData) {
      var value = settingsData[key];
      var rowIndex = -1;

      // Find existing setting row
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] === key) {
          rowIndex = i;
          break;
        }
      }

      if (rowIndex !== -1) {
        // Update existing setting
        sheet.getRange(rowIndex + 1, 2).setValue(value); // SettingValue
        sheet.getRange(rowIndex + 1, 3).setValue(timestamp); // UpdatedAt
        sheet.getRange(rowIndex + 1, 4).setValue(currentUser); // UpdatedBy
      } else {
        // Add new setting
        sheet.appendRow([key, value, timestamp, currentUser]);
      }
    }

    // Log activity
    addLog(currentUser, 'Update Settings', 'Updated system settings');

    return { success: true, message: 'Settings updated successfully!' };

  } catch (error) {
    return { success: false, message: 'Error updating settings: ' + error.toString() };
  }
}

// ============== File Upload Functions ==============

// Get or create ASSETS folder
function getAssetsFolder() {
  try {
    var folders = DriveApp.getFoldersByName(ASSETS_FOLDER_NAME);

    if (folders.hasNext()) {
      return folders.next();
    } else {
      // Create ASSETS folder
      var folder = DriveApp.createFolder(ASSETS_FOLDER_NAME);
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return folder;
    }
  } catch (error) {
    Logger.log('Error getting ASSETS folder: ' + error.toString());
    return null;
  }
}

// Upload profile image
function uploadProfileImage(base64Data, filename, username) {
  try {
    // Get ASSETS folder
    var folder = getAssetsFolder();
    if (!folder) {
      return { success: false, message: 'Failed to create ASSETS folder' };
    }

    // Remove data URL prefix if present
    var base64String = base64Data.split(',')[1] || base64Data;

    // Decode base64
    var blob = Utilities.newBlob(Utilities.base64Decode(base64String), 'image/jpeg', filename);

    // Upload file
    var file = folder.createFile(blob);
    file.setName(username + '_' + new Date().getTime() + '_' + filename);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Get file ID and create Google Drive URL
    var fileId = file.getId();
    var fileUrl = 'https://lh3.google.com/u/0/d/' + fileId;

    addLog(username, 'Profile Image Uploaded', 'Uploaded profile image: ' + file.getName());

    return {
      success: true,
      fileId: fileId,
      fileUrl: fileUrl,
      fileName: file.getName()
    };

  } catch (error) {
    return { success: false, message: 'Upload error: ' + error.toString() };
  }
}

// Update user settings (profile image, theme, colors)
function updateUserSettings(username, settings) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET);

    if (!sheet) {
      return { success: false, message: 'Users sheet not found' };
    }

    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === username) {
        var timestamp = new Date().toISOString();

        // Update settings
        if (settings.profileImage !== undefined) {
          sheet.getRange(i + 1, 6).setValue(settings.profileImage); // ProfileImage column
        }

        if (settings.themeMode !== undefined) {
          sheet.getRange(i + 1, 7).setValue(settings.themeMode); // ThemeMode column
        }

        if (settings.customColors !== undefined) {
          sheet.getRange(i + 1, 8).setValue(settings.customColors); // CustomColors column
        }

        // Update timestamp
        sheet.getRange(i + 1, 11).setValue(timestamp); // UpdatedAt
        sheet.getRange(i + 1, 12).setValue(username); // UpdatedBy

        addLog(username, 'Settings Updated', 'Updated user settings');

        return { success: true, message: 'Settings updated successfully!' };
      }
    }

    return { success: false, message: 'User not found' };

  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// Get user settings
function getUserSettings(username) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET);

    if (!sheet) {
      return { success: false, message: 'Users sheet not found' };
    }

    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === username) {
        return {
          success: true,
          settings: {
            profileImage: data[i][5] || '',
            themeMode: data[i][6] || 'light',
            customColors: data[i][7] || ''
          }
        };
      }
    }

    return { success: false, message: 'User not found' };

  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// ============== Helper Functions ==============

// Add log entry
function addLog(user, action, details) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOGS_SHEET);

    if (!sheet) {
      return;
    }

    var timestamp = new Date().toISOString();
    sheet.appendRow([timestamp, user, action, details]);

  } catch (error) {
    Logger.log('Error adding log: ' + error.toString());
  }
}

// ============== Setup Function ==============
// Run this function once to initialize the sheets
function setup() {
  return initializeSheets();
}

// ============== Setup Demo Data Function ==============
// Run this function to populate the system with demo/sample data for testing
function setupDemoData() {
  try {
    // First, initialize sheets
    initializeSheets();

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var usersSheet = ss.getSheetByName(USERS_SHEET);
    var logsSheet = ss.getSheetByName(LOGS_SHEET);

    // Clear existing data (keep headers)
    var lastRowUsers = usersSheet.getLastRow();
    if (lastRowUsers > 1) {
      usersSheet.deleteRows(2, lastRowUsers - 1);
    }

    var lastRowLogs = logsSheet.getLastRow();
    if (lastRowLogs > 1) {
      logsSheet.deleteRows(2, lastRowLogs - 1);
    }

    var timestamp = new Date().toISOString();

    // ============== Demo Users ==============
    var demoUsers = [
      // Admin Users
      {
        username: 'admin',
        email: 'admin@demo.com',
        password: 'admin123',
        role: 'Admin',
        status: 'Active'
      },
      {
        username: 'admin1',
        email: 'admin1@demo.com',
        password: 'demo123',
        role: 'Admin',
        status: 'Active'
      },
      {
        username: 'admin2',
        email: 'admin2@demo.com',
        password: 'demo123',
        role: 'Admin',
        status: 'Active'
      },

      // Regular Users - Active
      {
        username: 'user',
        email: 'user@demo.com',
        password: 'user123',
        role: 'User',
        status: 'Active'
      },
      {
        username: 'user1',
        email: 'user1@demo.com',
        password: 'demo123',
        role: 'User',
        status: 'Active'
      },
      {
        username: 'user2',
        email: 'user2@demo.com',
        password: 'demo123',
        role: 'User',
        status: 'Active'
      },
      {
        username: 'user3',
        email: 'user3@demo.com',
        password: 'demo123',
        role: 'User',
        status: 'Active'
      },
      {
        username: 'user4',
        email: 'user4@demo.com',
        password: 'demo123',
        role: 'User',
        status: 'Active'
      },
      {
        username: 'user5',
        email: 'user5@demo.com',
        password: 'demo123',
        role: 'User',
        status: 'Active'
      },
      {
        username: 'user6',
        email: 'user6@demo.com',
        password: 'demo123',
        role: 'User',
        status: 'Active'
      },

      // Inactive Users
      {
        username: 'user7',
        email: 'user7@demo.com',
        password: 'demo123',
        role: 'User',
        status: 'Inactive'
      },
      {
        username: 'user8',
        email: 'user8@demo.com',
        password: 'demo123',
        role: 'User',
        status: 'Inactive'
      },
      {
        username: 'admin3',
        email: 'admin3@demo.com',
        password: 'demo123',
        role: 'Admin',
        status: 'Inactive'
      }
    ];

    // Add demo users to sheet
    var defaultLogo = 'https://drive.google.com/file/d/165GgME6_H3EYDwA-eqUIqDNzHIq3KFy0/view?usp=drive_link';

    for (var i = 0; i < demoUsers.length; i++) {
      var user = demoUsers[i];

      // Create timestamp with slight variations for realistic data
      var userTimestamp = new Date(Date.now() - (demoUsers.length - i) * 86400000).toISOString(); // Each user created 1 day apart

      usersSheet.appendRow([
        user.username,
        user.email,
        user.password,
        user.role,
        user.status,
        defaultLogo,        // ProfileImage
        'light',            // ThemeMode
        '',                 // CustomColors
        userTimestamp,
        'System',
        userTimestamp,
        'System'
      ]);
    }

    // ============== Demo Logs ==============
    var demoLogs = [
      { user: 'admin', action: 'System Setup', details: 'Demo data initialized' },
      { user: 'admin', action: 'Login Success', details: 'Administrator logged in' },
      { user: 'admin1', action: 'Login Success', details: 'User logged in successfully' },
      { user: 'admin', action: 'User Added', details: 'Added user: user1' },
      { user: 'admin', action: 'User Added', details: 'Added user: user2' },
      { user: 'user', action: 'Login Success', details: 'User logged in successfully' },
      { user: 'user1', action: 'Login Success', details: 'User logged in successfully' },
      { user: 'admin', action: 'User Updated', details: 'Updated user: user7' },
      { user: 'admin2', action: 'Login Success', details: 'User logged in successfully' },
      { user: 'user2', action: 'Profile Updated', details: 'Updated own profile' },
      { user: 'admin', action: 'User Updated', details: 'Changed status for user: user8' },
      { user: 'user3', action: 'Login Success', details: 'User logged in successfully' },
      { user: 'user4', action: 'Profile Updated', details: 'Updated email address' },
      { user: 'admin', action: 'User Deleted', details: 'Deleted test user account' },
      { user: 'admin1', action: 'User Added', details: 'Added user: user5' },
      { user: 'user6', action: 'Login Success', details: 'User logged in successfully' },
      { user: 'user5', action: 'Profile Updated', details: 'Changed password' },
      { user: 'admin', action: 'Dashboard Viewed', details: 'Viewed dashboard statistics' },
      { user: 'admin2', action: 'User Updated', details: 'Updated user: admin3' },
      { user: 'user1', action: 'Login Failed', details: 'Invalid password attempt' }
    ];

    // Add demo logs with timestamps spread over the last 7 days
    for (var j = 0; j < demoLogs.length; j++) {
      var log = demoLogs[j];
      var logTimestamp = new Date(Date.now() - (demoLogs.length - j) * 7200000).toISOString(); // Each log 2 hours apart

      logsSheet.appendRow([
        logTimestamp,
        log.user,
        log.action,
        log.details
      ]);
    }

    // ============== Demo Parties ==============
    var partiesSheet = ss.getSheetByName(PARTIES_SHEET);

    // Clear existing parties data (keep headers)
    var lastRowParties = partiesSheet.getLastRow();
    if (lastRowParties > 1) {
      partiesSheet.deleteRows(2, lastRowParties - 1);
    }

    var demoParties = [
      // Customers
      {
        name: 'Customer 1',
        type: 'Customer',
        contact: '0300-1111111',
        address: 'Address 1',
        openingBalance: 5000.00,
        balanceType: 'Debit',
        isActive: true
      },
      {
        name: 'Customer 2',
        type: 'Customer',
        contact: '0300-2222222',
        address: 'Address 2',
        openingBalance: 12500.50,
        balanceType: 'Debit',
        isActive: true
      },
      {
        name: 'Customer 3',
        type: 'Customer',
        contact: '0300-3333333',
        address: 'Address 3',
        openingBalance: 8750.00,
        balanceType: 'Debit',
        isActive: true
      },
      {
        name: 'Customer 4',
        type: 'Customer',
        contact: '0300-4444444',
        address: 'Address 4',
        openingBalance: 0,
        balanceType: 'Debit',
        isActive: true
      },
      {
        name: 'Customer 5',
        type: 'Customer',
        contact: '0300-5555555',
        address: 'Address 5',
        openingBalance: 15000.00,
        balanceType: 'Credit',
        isActive: false
      },

      // Suppliers
      {
        name: 'Supplier 1',
        type: 'Supplier',
        contact: '0301-1111111',
        address: 'Address 6',
        openingBalance: 7500.00,
        balanceType: 'Credit',
        isActive: true
      },
      {
        name: 'Supplier 2',
        type: 'Supplier',
        contact: '0301-2222222',
        address: 'Address 7',
        openingBalance: 10250.75,
        balanceType: 'Credit',
        isActive: true
      },
      {
        name: 'Supplier 3',
        type: 'Supplier',
        contact: '0301-3333333',
        address: 'Address 8',
        openingBalance: 5600.00,
        balanceType: 'Credit',
        isActive: true
      },
      {
        name: 'Supplier 4',
        type: 'Supplier',
        contact: '0301-4444444',
        address: 'Address 9',
        openingBalance: 0,
        balanceType: 'Credit',
        isActive: true
      },
      {
        name: 'Supplier 5',
        type: 'Supplier',
        contact: '0301-5555555',
        address: 'Address 10',
        openingBalance: 3200.00,
        balanceType: 'Credit',
        isActive: false
      },

      // Both (Customer & Supplier)
      {
        name: 'Party 1',
        type: 'Both',
        contact: '0302-1111111',
        address: 'Address 11',
        openingBalance: 2500.00,
        balanceType: 'Debit',
        isActive: true
      },
      {
        name: 'Party 2',
        type: 'Both',
        contact: '0302-2222222',
        address: 'Address 12',
        openingBalance: 4750.50,
        balanceType: 'Credit',
        isActive: true
      },
      {
        name: 'Party 3',
        type: 'Both',
        contact: '0302-3333333',
        address: 'Address 13',
        openingBalance: 0,
        balanceType: 'Debit',
        isActive: true
      }
    ];

    // Add demo parties to sheet
    for (var k = 0; k < demoParties.length; k++) {
      var party = demoParties[k];
      var partyId = k + 1;

      // Create timestamp with variations
      var partyTimestamp = new Date(Date.now() - (demoParties.length - k) * 172800000).toISOString(); // Each party created 2 days apart

      var currentBalance = parseFloat(party.openingBalance) || 0;

      partiesSheet.appendRow([
        partyId,                  // PartyID
        party.name,               // PartyName
        party.type,               // PartyType
        party.contact,            // ContactNumber
        party.address,            // Address
        party.openingBalance,     // OpeningBalance
        party.balanceType,        // BalanceType
        currentBalance,           // CurrentBalance
        party.isActive,           // IsActive
        partyTimestamp,           // CreatedAt
        'System',                 // CreatedBy
        partyTimestamp,           // UpdatedAt
        'System'                  // UpdatedBy
      ]);
    }

    // ============== Demo Cash Book Entries ==============
    var cashBookSheet = ss.getSheetByName(CASH_BOOK_SHEET);

    // Clear existing cash book data (keep headers)
    var lastRowCashBook = cashBookSheet.getLastRow();
    if (lastRowCashBook > 1) {
      cashBookSheet.deleteRows(2, lastRowCashBook - 1);
    }

    var demoCashEntries = [
      // Opening Balance
      {
        date: new Date(Date.now() - 30 * 86400000),
        particulars: 'Opening Balance - Cash in Hand',
        cashIn: 50000.00,
        cashOut: 0,
        entryType: 'Cash',
        partyId: null,
        paymentMode: 'Cash',
        referenceNo: ''
      },

      // Customer Payments
      {
        date: new Date(Date.now() - 28 * 86400000),
        particulars: 'Payment received from ABC Corporation',
        cashIn: 5000.00,
        cashOut: 0,
        entryType: 'Customer',
        partyId: 1, // ABC Corporation
        paymentMode: 'Cash',
        referenceNo: ''
      },
      {
        date: new Date(Date.now() - 25 * 86400000),
        particulars: 'Payment received from XYZ Enterprises via bank transfer',
        cashIn: 12500.50,
        cashOut: 0,
        entryType: 'Customer',
        partyId: 2, // XYZ Enterprises
        paymentMode: 'Bank',
        referenceNo: 'TXN2025001234'
      },
      {
        date: new Date(Date.now() - 20 * 86400000),
        particulars: 'Payment received from Global Tech Solutions - Cheque',
        cashIn: 8750.00,
        cashOut: 0,
        entryType: 'Customer',
        partyId: 3, // Global Tech Solutions
        paymentMode: 'Cheque',
        referenceNo: 'CHQ-458921'
      },

      // Supplier Payments
      {
        date: new Date(Date.now() - 27 * 86400000),
        particulars: 'Payment made to Supreme Suppliers Inc',
        cashIn: 0,
        cashOut: 7500.00,
        entryType: 'Supplier',
        partyId: 6, // Supreme Suppliers Inc
        paymentMode: 'Bank',
        referenceNo: 'TXN2025001567'
      },
      {
        date: new Date(Date.now() - 22 * 86400000),
        particulars: 'Payment to Quality Materials Co via online transfer',
        cashIn: 0,
        cashOut: 10250.75,
        entryType: 'Supplier',
        partyId: 7, // Quality Materials Co
        paymentMode: 'Online',
        referenceNo: 'OLT-789654'
      },
      {
        date: new Date(Date.now() - 18 * 86400000),
        particulars: 'Payment to Bulk Distributors LLC',
        cashIn: 0,
        cashOut: 5600.00,
        entryType: 'Supplier',
        partyId: 8, // Bulk Distributors LLC
        paymentMode: 'Cheque',
        referenceNo: 'CHQ-112233'
      },

      // General Cash Transactions
      {
        date: new Date(Date.now() - 26 * 86400000),
        particulars: 'Office rent payment',
        cashIn: 0,
        cashOut: 3500.00,
        entryType: 'Cash',
        partyId: null,
        paymentMode: 'Cash',
        referenceNo: ''
      },
      {
        date: new Date(Date.now() - 24 * 86400000),
        particulars: 'Electricity bill payment',
        cashIn: 0,
        cashOut: 850.00,
        entryType: 'Cash',
        partyId: null,
        paymentMode: 'Online',
        referenceNo: 'BILL-2025-JAN-001'
      },
      {
        date: new Date(Date.now() - 23 * 86400000),
        particulars: 'Office supplies purchase',
        cashIn: 0,
        cashOut: 1250.00,
        entryType: 'Cash',
        partyId: null,
        paymentMode: 'Cash',
        referenceNo: ''
      },
      {
        date: new Date(Date.now() - 21 * 86400000),
        particulars: 'Employee salary advance',
        cashIn: 0,
        cashOut: 5000.00,
        entryType: 'Cash',
        partyId: null,
        paymentMode: 'Cash',
        referenceNo: ''
      },
      {
        date: new Date(Date.now() - 19 * 86400000),
        particulars: 'Fuel and maintenance expenses',
        cashIn: 0,
        cashOut: 2300.00,
        entryType: 'Cash',
        partyId: null,
        paymentMode: 'Cash',
        referenceNo: ''
      },
      {
        date: new Date(Date.now() - 17 * 86400000),
        particulars: 'Internet and phone bill',
        cashIn: 0,
        cashOut: 650.00,
        entryType: 'Cash',
        partyId: null,
        paymentMode: 'Bank',
        referenceNo: 'AUTO-DEBIT-5678'
      },
      {
        date: new Date(Date.now() - 15 * 86400000),
        particulars: 'Miscellaneous income',
        cashIn: 1500.00,
        cashOut: 0,
        entryType: 'Cash',
        partyId: null,
        paymentMode: 'Cash',
        referenceNo: ''
      },
      {
        date: new Date(Date.now() - 12 * 86400000),
        particulars: 'Stationery and printing',
        cashIn: 0,
        cashOut: 890.00,
        entryType: 'Cash',
        partyId: null,
        paymentMode: 'Cash',
        referenceNo: ''
      },
      {
        date: new Date(Date.now() - 10 * 86400000),
        particulars: 'Payment from Universal Trading Co',
        cashIn: 2500.00,
        cashOut: 0,
        entryType: 'Customer',
        partyId: 11, // Universal Trading Co (Both type)
        paymentMode: 'Bank',
        referenceNo: 'TXN2025002345'
      },
      {
        date: new Date(Date.now() - 8 * 86400000),
        particulars: 'Payment to Mega Wholesale Group',
        cashIn: 0,
        cashOut: 4750.50,
        entryType: 'Supplier',
        partyId: 12, // Mega Wholesale Group (Both type)
        paymentMode: 'Cheque',
        referenceNo: 'CHQ-998877'
      },
      {
        date: new Date(Date.now() - 6 * 86400000),
        particulars: 'Transportation charges',
        cashIn: 0,
        cashOut: 1200.00,
        entryType: 'Cash',
        partyId: null,
        paymentMode: 'Cash',
        referenceNo: ''
      },
      {
        date: new Date(Date.now() - 4 * 86400000),
        particulars: 'Staff refreshment and welfare',
        cashIn: 0,
        cashOut: 750.00,
        entryType: 'Cash',
        partyId: null,
        paymentMode: 'Cash',
        referenceNo: ''
      },
      {
        date: new Date(Date.now() - 2 * 86400000),
        particulars: 'Cash sales income',
        cashIn: 3500.00,
        cashOut: 0,
        entryType: 'Cash',
        partyId: null,
        paymentMode: 'Cash',
        referenceNo: ''
      }
    ];

    // Add demo cash book entries
    var runningBalance = 0;

    for (var m = 0; m < demoCashEntries.length; m++) {
      var entry = demoCashEntries[m];
      var cashId = m + 1;

      var entryTimestamp = entry.date.toISOString();
      var voucherNo = 'VC-' + Utilities.formatDate(entry.date, Session.getScriptTimeZone(), 'yyyyMMdd') + '-' + ('000' + cashId).slice(-3);

      var cashIn = parseFloat(entry.cashIn) || 0;
      var cashOut = parseFloat(entry.cashOut) || 0;
      runningBalance = runningBalance + cashIn - cashOut;

      // Get party name if applicable
      var partyName = '';
      if (entry.partyId) {
        var partyRow = partiesSheet.getRange(entry.partyId + 1, 2).getValue();
        partyName = partyRow || '';
      }

      cashBookSheet.appendRow([
        cashId,                   // CashID
        voucherNo,                // VoucherNo
        entryTimestamp,           // TransactionDate
        entry.particulars,        // Particulars
        cashIn,                   // CashIn
        cashOut,                  // CashOut
        entry.entryType,          // EntryType
        entry.partyId || '',      // PartyID
        partyName,                // PartyName
        entry.paymentMode,        // PaymentMode
        entry.referenceNo,        // ReferenceNo
        runningBalance,           // RunningBalance
        entry.partyId ? true : false,  // LedgerLinked
        'System',                 // CreatedBy
        entryTimestamp,           // CreatedAt
        entryTimestamp            // UpdatedAt
      ]);
    }

    // ============== Demo Party Ledger Entries ==============
    var partyLedgerSheet = ss.getSheetByName(PARTY_LEDGER_SHEET);

    // Clear existing party ledger data (keep headers)
    var lastRowLedger = partyLedgerSheet.getLastRow();
    if (lastRowLedger > 1) {
      partyLedgerSheet.deleteRows(2, lastRowLedger - 1);
    }

    var demoLedgerEntries = [
      // Opening balances
      {
        date: new Date(Date.now() - 30 * 86400000),
        partyId: 1, // ABC Corporation
        invoiceNo: '',
        debit: 15000.00,
        credit: 0,
        transactionType: 'Opening',
        paymentMode: 'Credit',
        narration: 'Opening balance - Amount receivable',
        referenceNo: '',
        cashBookId: null,
        isCashEntry: false
      },
      {
        date: new Date(Date.now() - 30 * 86400000),
        partyId: 6, // Supreme Suppliers Inc
        invoiceNo: '',
        debit: 0,
        credit: 20000.00,
        transactionType: 'Opening',
        paymentMode: 'Credit',
        narration: 'Opening balance - Amount payable',
        referenceNo: '',
        cashBookId: null,
        isCashEntry: false
      },

      // Sales transactions
      {
        date: new Date(Date.now() - 28 * 86400000),
        partyId: 1, // ABC Corporation
        invoiceNo: 'INV-2025-001',
        debit: 25000.00,
        credit: 0,
        transactionType: 'Sale',
        paymentMode: 'Credit',
        narration: 'Sale of goods as per invoice INV-2025-001',
        referenceNo: '',
        cashBookId: null,
        isCashEntry: false
      },
      {
        date: new Date(Date.now() - 26 * 86400000),
        partyId: 2, // XYZ Enterprises
        invoiceNo: 'INV-2025-002',
        debit: 18500.00,
        credit: 0,
        transactionType: 'Sale',
        paymentMode: 'Credit',
        narration: 'Sale of products - Invoice INV-2025-002',
        referenceNo: '',
        cashBookId: null,
        isCashEntry: false
      },
      {
        date: new Date(Date.now() - 24 * 86400000),
        partyId: 3, // Global Tech Solutions
        invoiceNo: 'INV-2025-003',
        debit: 32500.50,
        credit: 0,
        transactionType: 'Sale',
        paymentMode: 'Credit',
        narration: 'Goods sold on credit terms',
        referenceNo: '',
        cashBookId: null,
        isCashEntry: false
      },

      // Receipt transactions
      {
        date: new Date(Date.now() - 23 * 86400000),
        partyId: 1, // ABC Corporation
        invoiceNo: '',
        debit: 0,
        credit: 5000.00,
        transactionType: 'Receipt',
        paymentMode: 'Cash',
        narration: 'Cash received against outstanding balance',
        referenceNo: '',
        cashBookId: 2,
        isCashEntry: true
      },
      {
        date: new Date(Date.now() - 21 * 86400000),
        partyId: 2, // XYZ Enterprises
        invoiceNo: '',
        debit: 0,
        credit: 12500.50,
        transactionType: 'Receipt',
        paymentMode: 'Bank',
        narration: 'Payment received via bank transfer',
        referenceNo: 'TXN2025001234',
        cashBookId: 3,
        isCashEntry: true
      },
      {
        date: new Date(Date.now() - 19 * 86400000),
        partyId: 3, // Global Tech Solutions
        invoiceNo: '',
        debit: 0,
        credit: 8750.00,
        transactionType: 'Receipt',
        paymentMode: 'Cheque',
        narration: 'Cheque received and deposited',
        referenceNo: 'CHQ-458921',
        cashBookId: 4,
        isCashEntry: true
      },

      // Purchase transactions
      {
        date: new Date(Date.now() - 27 * 86400000),
        partyId: 6, // Supreme Suppliers Inc
        invoiceNo: 'PINV-5001',
        debit: 0,
        credit: 15000.00,
        transactionType: 'Purchase',
        paymentMode: 'Credit',
        narration: 'Purchased raw materials on credit',
        referenceNo: '',
        cashBookId: null,
        isCashEntry: false
      },
      {
        date: new Date(Date.now() - 25 * 86400000),
        partyId: 7, // Quality Materials Co
        invoiceNo: 'PINV-5002',
        debit: 0,
        credit: 22500.75,
        transactionType: 'Purchase',
        paymentMode: 'Credit',
        narration: 'Materials purchased as per invoice PINV-5002',
        referenceNo: '',
        cashBookId: null,
        isCashEntry: false
      },
      {
        date: new Date(Date.now() - 23 * 86400000),
        partyId: 8, // Bulk Distributors LLC
        invoiceNo: 'PINV-5003',
        debit: 0,
        credit: 18750.00,
        transactionType: 'Purchase',
        paymentMode: 'Credit',
        narration: 'Bulk purchase of inventory items',
        referenceNo: '',
        cashBookId: null,
        isCashEntry: false
      },

      // Payment transactions
      {
        date: new Date(Date.now() - 22 * 86400000),
        partyId: 6, // Supreme Suppliers Inc
        invoiceNo: '',
        debit: 7500.00,
        credit: 0,
        transactionType: 'Payment',
        paymentMode: 'Bank',
        narration: 'Payment made via bank transfer',
        referenceNo: 'TXN2025001567',
        cashBookId: 5,
        isCashEntry: true
      },
      {
        date: new Date(Date.now() - 20 * 86400000),
        partyId: 7, // Quality Materials Co
        invoiceNo: '',
        debit: 10250.75,
        credit: 0,
        transactionType: 'Payment',
        paymentMode: 'Online',
        narration: 'Online payment against invoice',
        referenceNo: 'OLT-789654',
        cashBookId: 6,
        isCashEntry: true
      },
      {
        date: new Date(Date.now() - 18 * 86400000),
        partyId: 8, // Bulk Distributors LLC
        invoiceNo: '',
        debit: 5600.00,
        credit: 0,
        transactionType: 'Payment',
        paymentMode: 'Cash',
        narration: 'Cash payment made',
        referenceNo: '',
        cashBookId: 7,
        isCashEntry: true
      },

      // Additional mixed transactions
      {
        date: new Date(Date.now() - 17 * 86400000),
        partyId: 4, // Metro Retail Chain
        invoiceNo: 'INV-2025-004',
        debit: 45000.00,
        credit: 0,
        transactionType: 'Sale',
        paymentMode: 'Credit',
        narration: 'Bulk order delivered - Invoice INV-2025-004',
        referenceNo: '',
        cashBookId: null,
        isCashEntry: false
      },
      {
        date: new Date(Date.now() - 15 * 86400000),
        partyId: 5, // Diamond Services Ltd
        invoiceNo: 'INV-2025-005',
        debit: 28750.25,
        credit: 0,
        transactionType: 'Sale',
        paymentMode: 'Credit',
        narration: 'Service charges billed',
        referenceNo: '',
        cashBookId: null,
        isCashEntry: false
      },
      {
        date: new Date(Date.now() - 14 * 86400000),
        partyId: 4, // Metro Retail Chain
        invoiceNo: '',
        debit: 0,
        credit: 15000.00,
        transactionType: 'Receipt',
        paymentMode: 'Bank',
        narration: 'Partial payment received',
        referenceNo: 'TXN2025002345',
        cashBookId: 8,
        isCashEntry: true
      },
      {
        date: new Date(Date.now() - 12 * 86400000),
        partyId: 9, // Premium Goods Trading
        invoiceNo: 'PINV-5004',
        debit: 0,
        credit: 35000.00,
        transactionType: 'Purchase',
        paymentMode: 'Credit',
        narration: 'Premium quality goods purchased',
        referenceNo: '',
        cashBookId: null,
        isCashEntry: false
      },
      {
        date: new Date(Date.now() - 10 * 86400000),
        partyId: 11, // Nationwide Distributors
        invoiceNo: 'INV-2025-006',
        debit: 52500.00,
        credit: 0,
        transactionType: 'Sale',
        paymentMode: 'Credit',
        narration: 'Large order shipment completed',
        referenceNo: '',
        cashBookId: null,
        isCashEntry: false
      },
      {
        date: new Date(Date.now() - 8 * 86400000),
        partyId: 12, // Eastern Supply Co
        invoiceNo: 'PINV-5005',
        debit: 0,
        credit: 19500.50,
        transactionType: 'Purchase',
        paymentMode: 'Credit',
        narration: 'Monthly supply order received',
        referenceNo: '',
        cashBookId: null,
        isCashEntry: false
      },
      {
        date: new Date(Date.now() - 6 * 86400000),
        partyId: 13, // Coastal Traders Inc
        invoiceNo: 'INV-2025-007',
        debit: 38250.75,
        credit: 0,
        transactionType: 'Sale',
        paymentMode: 'Credit',
        narration: 'Export order invoiced',
        referenceNo: '',
        cashBookId: null,
        isCashEntry: false
      }
    ];

    // Add demo party ledger entries with running balances
    var partyBalances = {}; // Track running balance per party

    for (var n = 0; n < demoLedgerEntries.length; n++) {
      var ledgerEntry = demoLedgerEntries[n];
      var ledgerId = n + 1;

      var ledgerTimestamp = ledgerEntry.date.toISOString();

      // Get party name
      var ledgerPartyName = '';
      if (ledgerEntry.partyId) {
        var ledgerPartyRow = partiesSheet.getRange(ledgerEntry.partyId + 1, 2).getValue();
        ledgerPartyName = ledgerPartyRow || '';
      }

      // Calculate running balance for this party
      if (!partyBalances[ledgerEntry.partyId]) {
        partyBalances[ledgerEntry.partyId] = 0;
      }

      var debitAmount = parseFloat(ledgerEntry.debit) || 0;
      var creditAmount = parseFloat(ledgerEntry.credit) || 0;
      partyBalances[ledgerEntry.partyId] = partyBalances[ledgerEntry.partyId] + debitAmount - creditAmount;

      partyLedgerSheet.appendRow([
        ledgerId,                      // LedgerID
        ledgerTimestamp,               // TransactionDate
        ledgerEntry.partyId,           // PartyID
        ledgerPartyName,               // PartyName
        ledgerEntry.invoiceNo,         // InvoiceNo
        debitAmount,                   // DebitAmount
        creditAmount,                  // CreditAmount
        ledgerEntry.transactionType,   // TransactionType
        ledgerEntry.paymentMode,       // PaymentMode
        ledgerEntry.narration,         // Narration
        ledgerEntry.referenceNo,       // ReferenceNo
        partyBalances[ledgerEntry.partyId], // RunningBalance
        ledgerEntry.cashBookId || '',  // CashBookID
        ledgerEntry.isCashEntry,       // IsCashEntry
        'System',                      // CreatedBy
        ledgerTimestamp,               // CreatedAt
        ledgerTimestamp                // UpdatedAt
      ]);
    }

    // Format sheets
    usersSheet.autoResizeColumns(1, 9);
    logsSheet.autoResizeColumns(1, 4);
    partiesSheet.autoResizeColumns(1, 13);
    cashBookSheet.autoResizeColumns(1, 16);
    partyLedgerSheet.autoResizeColumns(1, 17);

    // ============== Generate Daily Summaries from Cash Book ==============
    var summaryResult = regenerateAllSummaries('System');
    var dailySummarySheet = ss.getSheetByName(DAILY_SUMMARY_SHEET);
    if (dailySummarySheet) {
      dailySummarySheet.autoResizeColumns(1, 8);
    }

    // Return success with summary
    return {
      success: true,
      message: 'Demo data setup complete!',
      summary: {
        totalUsers: demoUsers.length,
        adminUsers: demoUsers.filter(function(u) { return u.role === 'Admin'; }).length,
        regularUsers: demoUsers.filter(function(u) { return u.role === 'User'; }).length,
        activeUsers: demoUsers.filter(function(u) { return u.status === 'Active'; }).length,
        inactiveUsers: demoUsers.filter(function(u) { return u.status === 'Inactive'; }).length,
        totalLogs: demoLogs.length,
        totalParties: demoParties.length,
        customers: demoParties.filter(function(p) { return p.type === 'Customer'; }).length,
        suppliers: demoParties.filter(function(p) { return p.type === 'Supplier'; }).length,
        both: demoParties.filter(function(p) { return p.type === 'Both'; }).length,
        totalCashEntries: demoCashEntries.length,
        totalCashIn: demoCashEntries.reduce(function(sum, e) { return sum + e.cashIn; }, 0).toFixed(2),
        totalCashOut: demoCashEntries.reduce(function(sum, e) { return sum + e.cashOut; }, 0).toFixed(2),
        currentBalance: runningBalance.toFixed(2),
        totalLedgerEntries: demoLedgerEntries.length,
        totalDebit: demoLedgerEntries.reduce(function(sum, e) { return sum + e.debit; }, 0).toFixed(2),
        totalCredit: demoLedgerEntries.reduce(function(sum, e) { return sum + e.credit; }, 0).toFixed(2),
        totalDailySummaries: summaryResult.count || 0
      }
    };

  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// ==========================================
// DAYBOOK BACKEND FUNCTIONS
// ==========================================
function getDaybookSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DAYBOOK_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DAYBOOK_SHEET_NAME);
    // Updated header to "Total Sale"
    sheet.appendRow(['ID', 'Date', 'Cash IN', 'Online IN', 'Expense', 'Total Sale']);
    sheet.getRange("A1:F1").setBackground('#001f3f').setFontColor('white').setFontWeight('bold');
  }
  return sheet;
}

function getDaybookEntries() {
  const sheet = getDaybookSheet();
  const data = sheet.getDataRange().getValues();
  const entries = [];
  
  for (let i = 1; i < data.length; i++) {
    if(data[i][0]) { 
      let dateVal = data[i][1];
      if (dateVal instanceof Date) {
        dateVal = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }

      entries.push({
        id: data[i][0].toString(),
        date: dateVal,
        cash: data[i][2],
        online: data[i][3],
        expenses: data[i][4],
        totalSale: data[i][5] // Including the total from the sheet
      });
    }
  }
  return entries.reverse();
}

function addDaybookEntry(formData) {
  const sheet = getDaybookSheet();
  const id = new Date().getTime().toString(); 
  
  const cash = parseFloat(formData.cash) || 0;
  const online = parseFloat(formData.online) || 0;
  const expense = parseFloat(formData.expenses) || 0;
  
  // FIXED MATH: Total Sale = Cash + Online + Expense
  const totalSale = cash + online + expense; 
  
  sheet.appendRow([id, formData.date, cash, online, expense, totalSale]);
  return id; 
}

function updateDaybookEntry(id, formData) {
  const sheet = getDaybookSheet();
  const data = sheet.getDataRange().getValues();
  
  const cash = parseFloat(formData.cash) || 0;
  const online = parseFloat(formData.online) || 0;
  const expense = parseFloat(formData.expenses) || 0;
  
  // FIXED MATH: Total Sale = Cash + Online + Expense
  const totalSale = cash + online + expense; 
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === id.toString()) {
      // Update starting from column 2 (Date) through column 6 (Total Sale)
      sheet.getRange(i + 1, 2, 1, 5).setValues([[formData.date, cash, online, expense, totalSale]]);
      return true;
    }
  }
  return false;
}

function deleteDaybookEntry(id) {
  const sheet = getDaybookSheet();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === id.toString()) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}