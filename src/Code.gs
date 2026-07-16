/**
 * Packing Slip Software - Backend Script
 * Google Apps Script backend
 */

const SHEETS = {
  SETTINGS: 'Settings',
  CLIENTS: 'Clients',
  PACKING_SLIPS: 'PackingSlips',
  USERS: 'Users',
  ORDERS: 'Orders'
};

const COUNTERS = {
  CLIENT: 'clientCounter',
  SLIP: 'slipCounter',
  USER: 'userCounter',
  ORDER: 'orderCounter'
};

// Invoices are no longer first-class records. They live as comma-separated
// "Invoice Numbers" string lists on Orders (column 8) and free-floating on
// PackingSlips (column 11 + column 16). Helpers below normalize both.
function normalizeInvoiceNumberList(raw) {
  if (raw === null || raw === undefined) return [];
  var seen = {};
  var out = [];
  String(raw).split(',').forEach(function(part) {
    var s = String(part || '').trim();
    if (s && !seen[s]) {
      seen[s] = true;
      out.push(s);
    }
  });
  return out;
}

function joinInvoiceNumberList(arr) {
  if (!arr || !arr.length) return '';
  return arr.map(function(s) { return String(s).trim(); }).filter(Boolean).join(', ');
}

/* ------------------------------------------------------------------
 * Licensing & Access Control
 * ------------------------------------------------------------------ */

function checkLicense() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetId = ss.getId();
  var props = PropertiesService.getScriptProperties();
  var blockedSheetsStr = props.getProperty('BLOCKED_SHEETS');
  
  // Log license check attempts in script logs
  try {
    var owner = ss.getOwner() ? ss.getOwner().getEmail() : 'Unknown Owner';
    var sheetName = ss.getName();
    Logger.log("License check for Sheet ID: " + sheetId + " | Name: " + sheetName + " | Owner: " + owner);
  } catch (e) {
    Logger.log("License logging error: " + e.message);
  }
  
  // If the admin hasn't set up the blocked list yet, default to active (everyone runs free)
  if (blockedSheetsStr === null || blockedSheetsStr.trim() === "") {
    return true;
  }
  
  var blockedSheets = blockedSheetsStr.split(',').map(function(s) { return s.trim(); });
  // If this Sheet ID is in the blocked list, access is denied
  if (blockedSheets.indexOf(sheetId) !== -1) {
    return false;
  }
  
  return true;
}

function checkLicenseOrThrow() {
  if (!checkLicense()) {
    throw new Error("License Inactive: Your Bifro Packing Slip license is inactive. Please contact support to reactivate or downgrade to the offline version.");
  }
}

function getLicenseExpiredHtml() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ownerEmail = 'Unknown';
  try {
    var owner = ss.getOwner();
    if (owner) ownerEmail = owner.getEmail();
  } catch (e) {
    ownerEmail = 'Unknown';
  }

  return HtmlService.createHtmlOutput(
    "<!DOCTYPE html><html><head><title>License Expired</title>" +
    "<meta name='viewport' content='width=device-width, initial-scale=1.0'>" +
    "<style>" +
    "body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; background-color: #0f172a; color: #f8fafc; display: flex; flex-direction: column; align-items: center; justify-content: center; min-height: 100vh; margin: 0; padding: 24px; box-sizing: border-box; text-align: center; }" +
    ".card { background-color: #1e293b; padding: 40px; border-radius: 16px; border: 1px solid #334155; max-width: 500px; box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.3), 0 8px 10px -6px rgba(0, 0, 0, 0.3); }" +
    "h1 { color: #f59e0b; font-size: 1.8rem; margin: 0 0 16px 0; font-weight: 700; letter-spacing: -0.025em; }" +
    "p { color: #94a3b8; font-size: 1rem; line-height: 1.6; margin: 0 0 24px 0; }" +
    ".contact-btn { display: inline-block; background-color: #f59e0b; color: #0f172a; font-weight: 600; text-decoration: none; padding: 12px 24px; border-radius: 8px; font-size: 0.95rem; transition: background-color 0.2s; }" +
    ".contact-btn:hover { background-color: #d97706; }" +
    "span.meta { display: block; font-family: monospace; font-size: 0.75rem; color: #64748b; margin-top: 16px; word-break: break-all; }" +
    "</style></head><body>" +
    "<div class='card'>" +
    "<h1>Bifro Packing Slip App</h1>" +
    "<h2>License Expired</h2>" +
    "<p>Your automatic updates subscription is inactive. To continue using the software, please contact us to reactivate your plan or request to have the self-contained, offline version installed on your sheet.</p>" +
    "<a href='mailto:support@bifro.com' class='contact-btn'>Contact Support</a>" +
    "<span class='meta' style='margin-top: 32px;'>Sheet ID: " + ss.getId() + "</span>" +
    "<span class='meta'>Sheet Owner: " + ownerEmail + "</span>" +
    "</div></body></html>"
  ).setTitle("Bifro - License Expired");
}

/* ------------------------------------------------------------------
 * Helpers
 * ------------------------------------------------------------------ */

function getSheet(name, create) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var sheet = null;
    var searchName = name.toLowerCase().trim();
    
    for (var i = 0; i < sheets.length; i++) {
      var sName = sheets[i].getName().toLowerCase().trim();
      if (sName === searchName) {
        sheet = sheets[i];
        // Rename it to the canonical name if it's slightly off
        if (sheets[i].getName() !== name) {
          try {
            sheets[i].setName(name);
          } catch(e) {}
        }
        break;
      }
    }
    
    if (!sheet && create !== false) {
      try {
        sheet = ss.insertSheet(name);
        initSheet(sheet, name);
      } catch (e) {
        try {
          sheet = ss.getSheetByName(name) || ss.insertSheet();
          if (sheet.getName() !== name) {
            sheet.setName(name);
          }
          initSheet(sheet, name);
        } catch(err) {
          sheet = null;
        }
      }
    }
    return sheet;
  } catch (err) {
    Logger.log("getSheet error for name " + name + ": " + err.message);
    return null;
  }
}

function initSheet(sheet, name) {
  if (name === SHEETS.SETTINGS) {
    sheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]).setFontWeight('bold');
    var defaults = [
      ['CompanyName', 'My Company'],
      ['Address', ''],
      ['Phone', '']
    ];
    if (sheet.getLastRow() < 2) sheet.getRange(2, 1, defaults.length, 2).setValues(defaults);
  } else if (name === SHEETS.CLIENTS) {
    sheet.getRange(1, 1, 1, 6).setValues([['Client ID', 'Client Name', 'Company Name', 'Phone', 'Address', 'Created At']]).setFontWeight('bold');
  } else if (name === SHEETS.PACKING_SLIPS) {
    sheet.getRange(1, 1, 1, 11).setValues([['Slip Code', 'Date', 'Client ID', 'Client Name', 'Client Company', 'Client Phone', 'Client Address', 'Items JSON', 'Total Quantity', 'Notes', 'Created At']]).setFontWeight('bold');
  } else if (name === SHEETS.USERS) {
    sheet.getRange(1, 1, 1, 5).setValues([['User ID', 'Name', 'Email', 'Role', 'Created At']]).setFontWeight('bold');
  } else if (name === SHEETS.ORDERS) {
    sheet.getRange(1, 1, 1, 8).setValues([['Order ID', 'Date', 'Client ID', 'Client Name', 'Items JSON', 'Status', 'Created At', 'Invoice Numbers']]).setFontWeight('bold');
  }
}

/**
 * Generates the next sequential ID thread-safely using LockService.
 */
function getNextId(key, prefix) {
  var lock = LockService.getScriptLock();
  try {
    // Wait for up to 15 seconds to acquire lock
    lock.waitLock(15000);
  } catch (e) {
    throw new Error('Timeout acquiring lock to generate unique ID. Please try again.');
  }

  try {
    var props = PropertiesService.getScriptProperties();
    var propKey = 'counter_' + key;
    var val = props.getProperty(propKey);
    var num = val ? parseInt(val, 10) : 0;
    num += 1;
    props.setProperty(propKey, num.toString());
    return prefix + String(num).padStart(4, '0');
  } finally {
    // Make sure we release the lock even if code fails
    lock.releaseLock();
  }
}

function getSettingsSheetValues() {
  var sheet = getSheet(SHEETS.SETTINGS);
  var values = sheet.getDataRange().getValues();
  var settings = {};
  for (var i = 1; i < values.length; i++) {
    settings[values[i][0]] = values[i][1];
  }
  return settings;
}

function getUserFirstName(email) {
  if (!email) return 'Admin';
  var username = email.split('@')[0].toLowerCase();
  var parts = username.split(/[\._-]/);
  if (parts.length > 0 && parts[0]) {
    var name = parts[0];
    return name.charAt(0).toUpperCase() + name.slice(1);
  }
  return 'Admin';
}

function getCurrentUserEmail() {
  var email = '';
  try {
    email = Session.getActiveUser().getEmail();
    if (!email) {
      email = Session.getEffectiveUser().getEmail();
    }
  } catch (e) {
    email = '';
  }
  return email;
}

function getUserAccess() {
  var userEmail = '';
  try {
    userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      userEmail = Session.getEffectiveUser().getEmail();
    }
  } catch (e) {
    userEmail = '';
  }

  var authUrl = '';
  try {
    authUrl = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL).getAuthorizationUrl();
  } catch (e) {
    Logger.log('Could not get auth URL: ' + e.message);
  }

  // Load company settings if available to render uploader brand on access screens
  var companyName = 'Packing Slip Software';
  var logoBase64 = '';
  try {
    var settings = getSettingsSheetValues();
    if (settings.CompanyName) companyName = settings.CompanyName;
    if (settings.LogoBase64) logoBase64 = settings.LogoBase64;
  } catch (e) {
    Logger.log('Could not load company settings for access block: ' + e.message);
  }

  if (!userEmail) {
    return { loggedIn: false, email: '', name: 'Guest', access: false, role: 'Viewer', authUrl: authUrl, companyName: companyName, logoBase64: logoBase64, message: 'Please sign in with your Google account to continue.' };
  }

  var firstName = getUserFirstName(userEmail);
  
  // 1. Check spreadsheet access
  var ss;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    // Force a read operation to trigger Google's permission check immediately
    ss.getSheets();
  } catch (e) {
    return { loggedIn: true, email: userEmail, name: firstName, access: false, role: 'Viewer', authUrl: authUrl, companyName: companyName, logoBase64: logoBase64, message: 'You do not have access to this Google Sheet. Details: ' + e.toString() + '. Ask the owner to share it with you.' };
  }

  // 2. Ensure Users sheet is initialized
  var usersSheet = getSheet(SHEETS.USERS);
  var usersData = usersSheet.getDataRange().getValues();
  
  // If the sheet only has headers (empty user list), register current user as default Admin
  if (usersData.length <= 1) {
    var ownerName = firstName + ' ' + (userEmail.split('@')[0].split(/[\._-]/)[1] || 'Admin');
    usersSheet.appendRow(['U-0001', ownerName, userEmail, 'Admin', new Date().toISOString()]);
    usersData = usersSheet.getDataRange().getValues(); // refresh data
  }

  // 3. Match user email in Users list
  var authorized = false;
  var userRole = 'Viewer';
  var displayName = firstName;

  for (var i = 1; i < usersData.length; i++) {
    var emailInSheet = String(usersData[i][2]).trim().toLowerCase();
    if (emailInSheet === userEmail.toLowerCase()) {
      authorized = true;
      displayName = usersData[i][1];
      userRole = usersData[i][3];
      break;
    }
  }

  if (!authorized) {
    return { 
      loggedIn: true, 
      email: userEmail, 
      name: displayName, 
      access: false, 
      role: 'Viewer', 
      authUrl: authUrl,
      companyName: companyName,
      logoBase64: logoBase64,
      message: 'Access Restricted. Your Google Account (' + userEmail + ') is not registered in this application. Please ask an Admin to register you.' 
    };
  }

  return { loggedIn: true, email: userEmail, name: displayName, access: true, role: userRole, companyName: companyName, logoBase64: logoBase64, message: '' };
}

function getUsers() {
  checkLicenseOrThrow();
  var callerAccess = getUserAccess();
  if (!callerAccess.access || callerAccess.role !== 'Admin') {
    throw new Error('Unauthorized access: Only Admins can manage users.');
  }

  var sheet = getSheet(SHEETS.USERS);
  var values = sheet.getDataRange().getValues();
  var users = [];
  for (var i = 1; i < values.length; i++) {
    if (values[i][0]) {
      users.push({
        id: values[i][0],
        name: values[i][1],
        email: values[i][2],
        role: values[i][3],
        createdAt: values[i][4]
      });
    }
  }
  return users;
}

function saveUser(user) {
  checkLicenseOrThrow();
  var callerAccess = getUserAccess();
  if (!callerAccess.access || callerAccess.role !== 'Admin') {
    throw new Error('Unauthorized access: Only Admins can manage users.');
  }

  if (!user || !user.name || !user.name.trim() || !user.email || !user.email.trim()) {
    throw new Error('User Name and Email are required.');
  }

  user.email = user.email.trim().toLowerCase();
  var sheet = getSheet(SHEETS.USERS);
  var data = sheet.getDataRange().getValues();

  var shareErrors = [];

  // Try programmatically sharing spreadsheet and script project with the user
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.addEditor(user.email);
  } catch (e) {
    shareErrors.push('Sheet sharing failed: ' + e.message);
  }

  try {
    var scriptId = ScriptApp.getScriptId();
    var scriptFile = DriveApp.getFileById(scriptId);
    if (user.role === 'Admin') {
      scriptFile.addEditor(user.email);
    } else {
      scriptFile.addViewer(user.email);
    }
  } catch (e) {
    shareErrors.push('Script sharing failed: ' + e.message);
  }

  if (user.id) {
    // Update existing user details
    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === user.id) {
        sheet.getRange(i + 1, 2, 1, 3).setValues([[
          user.name,
          user.email,
          user.role || 'User'
        ]]);
        found = true;
        break;
      }
    }
    if (!found) {
      throw new Error('User ID ' + user.id + ' not found.');
    }
  } else {
    // Check duplicate email
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2]).trim().toLowerCase() === user.email) {
        throw new Error('User with email ' + user.email + ' already exists.');
      }
    }
    // Create new user record
    user.id = getNextId(COUNTERS.USER, 'U-');
    user.createdAt = new Date().toISOString();
    sheet.appendRow([user.id, user.name, user.email, user.role || 'User', user.createdAt]);
  }
  
  user.shareWarnings = shareErrors;
  return user;
}

function deleteUser(userId) {
  checkLicenseOrThrow();
  var callerAccess = getUserAccess();
  if (!callerAccess.access || callerAccess.role !== 'Admin') {
    throw new Error('Unauthorized access: Only Admins can manage users.');
  }

  var sheet = getSheet(SHEETS.USERS);
  var data = sheet.getDataRange().getValues();
  
  var targetEmail = '';
  var foundIndex = -1;

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      targetEmail = data[i][2];
      foundIndex = i + 1;
      break;
    }
  }

  if (foundIndex === -1) {
    throw new Error('User ID ' + userId + ' not found.');
  }

  if (targetEmail.toLowerCase() === callerAccess.email.toLowerCase()) {
    throw new Error('You cannot delete your own administrator account.');
  }

  // Attempt to revoke permissions from spreadsheet and script
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.removeEditor(targetEmail);
  } catch (e) {
    Logger.log('Could not revoke spreadsheet editor: ' + e.message);
  }

  try {
    var scriptId = ScriptApp.getScriptId();
    var scriptFile = DriveApp.getFileById(scriptId);
    scriptFile.removeEditor(targetEmail);
    scriptFile.removeViewer(targetEmail);
  } catch (e) {
    Logger.log('Could not revoke script access: ' + e.message);
  }

  sheet.deleteRow(foundIndex);
  return true;
}

/* ------------------------------------------------------------------
 * Web app entry point
 * ------------------------------------------------------------------ */

function doGet(e) {
  if (!checkLicense()) {
    return getLicenseExpiredHtml();
  }

  // Serve the PWA manifest from a real URL so browsers recognize it.
  if (e && e.parameter && e.parameter.manifest === 'true') {
    return serveManifest();
  }

  var settings = getSettingsSheetValues();
  var companyName = settings.CompanyName || 'Packing Slip Software';
  var access = getUserAccess();
  var template = HtmlService.createTemplateFromFile('index');
  template.access = access;
  return template.evaluate()
    .setTitle(companyName + ' - Packing Slips')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

function serveManifest() {
  var settings = getSettingsSheetValues();
  var companyName = settings.CompanyName || 'Packing Slip Software';
  // SVG icon with a package layout emoji, works at any size.
  var svgIcon = "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 512 512'%3E%3Crect width='512' height='512' rx='80' fill='%230f172a'/%3E%3Ctext x='256' y='340' font-size='280' text-anchor='middle' fill='%23f59e0b'%3E%F0%9F%9B%8F%EF%B8%8F%3C/text%3E%3C/svg%3E";
  var manifest = {
    name: companyName,
    short_name: companyName.split(' ')[0] || 'PackingSlip',
    start_url: getScriptUrl(),
    display: 'standalone',
    background_color: '#0f172a',
    theme_color: '#0f172a',
    description: 'Create, print, and track professional packing slips for ' + companyName + '.',
    icons: [
      { src: svgIcon, sizes: '192x192', type: 'image/svg+xml' },
      { src: svgIcon, sizes: '512x512', type: 'image/svg+xml' }
    ]
  };
  return ContentService.createTextOutput(JSON.stringify(manifest))
    .setMimeType(ContentService.MimeType.JSON);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ------------------------------------------------------------------
 * Exposed API
 * ------------------------------------------------------------------ */

function getAccess() {
  checkLicenseOrThrow();
  return getUserAccess();
}

function getSettings() {
  checkLicenseOrThrow();
  return getSettingsSheetValues();
}

function saveSettings(settings) {
  checkLicenseOrThrow();
  if (!settings || !settings.CompanyName || !settings.CompanyName.trim()) {
    throw new Error('Company Name is required.');
  }

  var sheet = getSheet(SHEETS.SETTINGS);
  var values = sheet.getDataRange().getValues();
  var existing = {};
  for (var i = 1; i < values.length; i++) {
    existing[values[i][0]] = i + 1;
  }
  for (var key in settings) {
    if (existing[key]) {
      sheet.getRange(existing[key], 2).setValue(settings[key]);
    } else {
      sheet.appendRow([key, settings[key]]);
    }
  }
  return getSettingsSheetValues();
}

function getClients() {
  checkLicenseOrThrow();
  var sheet = getSheet(SHEETS.CLIENTS);
  var values = sheet.getDataRange().getValues();
  var clients = [];
  for (var i = 1; i < values.length; i++) {
    if (values[i][0]) {
      clients.push({
        id: values[i][0],
        name: values[i][1],
        company: values[i][2],
        phone: values[i][3],
        address: values[i][4],
        createdAt: values[i][5]
      });
    }
  }
  return clients;
}

function saveClient(client) {
  checkLicenseOrThrow();
  if (!client || !client.name || !client.name.trim()) {
    throw new Error('Client Name is required.');
  }

  var sheet = getSheet(SHEETS.CLIENTS);
  
  if (client.id) {
    // Update existing client details
    var data = sheet.getDataRange().getValues();
    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === client.id) {
        // Update Columns: Name, Company, Phone, Address (columns 2 to 5, which is 1-indexed range of columns)
        sheet.getRange(i + 1, 2, 1, 4).setValues([[
          client.name,
          client.company || '',
          client.phone || '',
          client.address || ''
        ]]);
        found = true;
        break;
      }
    }
    if (!found) {
      throw new Error('Client ID ' + client.id + ' not found for update.');
    }
  } else {
    // Create new client
    client.id = getNextId(COUNTERS.CLIENT, 'C-');
    client.createdAt = new Date().toISOString();
    sheet.appendRow([client.id, client.name, client.company || '', client.phone || '', client.address || '', client.createdAt]);
  }
  return client;
}

function normalizeDate(d) {
  if (!d) return '';
  if (typeof d === 'string') return d;
  if (d instanceof Date) {
    var year = d.getFullYear();
    var month = String(d.getMonth() + 1).padStart(2, '0');
    var day = String(d.getDate()).padStart(2, '0');
    return year + '-' + month + '-' + day;
  }
  return String(d);
}

function ensurePackingSlipsHeader(sheet) {
  var headerLen = sheet.getLastColumn();
  var headerRow = headerLen > 0 ? sheet.getRange(1, 1, 1, headerLen).getValues()[0] : [];
  var desired = ['Slip Code', 'Date', 'Client ID', 'Client Name', 'Client Company', 'Client Phone', 'Client Address', 'Items JSON', 'Total Quantity', 'Notes', 'Invoice Numbers', 'Created At', 'Created By', 'Edited By', 'Linked Orders', 'Invoice Numbers (Legacy)'];
  var col11 = headerRow.length >= 11 ? String(headerRow[10] || '').trim() : '';
  var col16 = headerRow.length >= 16 ? String(headerRow[15] || '').trim() : '';
  // Migrate legacy header names independently so we don't accidentally rely
  // on column 11 to deduce column 16's state.
  if (col11 === 'Invoice Number') {
    sheet.getRange(1, 11).setValue('Invoice Numbers');
  }
  if (col16 === 'Linked Invoices') {
    sheet.getRange(1, 16).setValue('Invoice Numbers (Legacy)');
  }
  // Ensure header is at least 16 columns wide with the right headers.
  if (headerRow.length < 16) {
    sheet.getRange(1, 1, 1, desired.length).setValues([desired]).setFontWeight('bold');
  } else if (headerRow[14] !== 'Linked Orders' || headerRow[15] !== 'Invoice Numbers (Legacy)') {
    sheet.getRange(1, 1, 1, desired.length).setValues([desired]).setFontWeight('bold');
  }
}

/** Repairs any rows where the legacy "Invoice Number" column has raw data and the
 *  new "Invoice Numbers" column is empty – essentially a no-op safety net since
 *  the header migration keeps the existing data in place and the read path merges
 *  both columns at load time.
 */
function ensureOrdersHeader(sheet) {
  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var desired = ['Order ID', 'Date', 'Client ID', 'Client Name', 'Items JSON', 'Status', 'Created At', 'Invoice Numbers'];
  if (header.length < 8 || header[7] !== 'Invoice Numbers') {
    sheet.getRange(1, 1, 1, desired.length).setValues([desired]).setFontWeight('bold');
  }
}

function getPackingSlips() {
  checkLicenseOrThrow();
  var sheet = getSheet(SHEETS.PACKING_SLIPS);
  ensurePackingSlipsHeader(sheet);
  var values = sheet.getDataRange().getValues();
  var slips = [];
  for (var i = 1; i < values.length; i++) {
    if (values[i][0]) {
      var items = [];
      try {
        items = JSON.parse(values[i][7] || '[]');
      } catch (e) {
        items = [];
      }
      // Merge free-floating invoice numbers from the primary column and
      // the legacy "Invoice Numbers (Legacy)" column (was "Linked Invoices").
      var invoiceNumbers = normalizeInvoiceNumberList(values[i][10]).concat(
        normalizeInvoiceNumberList(values[i][15])
      );
      slips.push({
        slipCode: values[i][0],
        date: normalizeDate(values[i][1]),
        clientId: values[i][2],
        clientName: values[i][3],
        clientCompany: values[i][4],
        clientPhone: values[i][5],
        clientAddress: values[i][6],
        items: items,
        totalQuantity: values[i][8],
        notes: values[i][9],
        invoiceNumber: joinInvoiceNumberList(invoiceNumbers),
        invoiceNumbers: invoiceNumbers,
        createdAt: values[i][11],
        createdBy: values[i][12] || '',
        editedBy: values[i][13] || '',
        linkedOrders: values[i][14] || '',
        linkedInvoices: values[i][15] || ''
      });
    }
  }
  return slips;
}

function savePackingSlip(slip) {
  checkLicenseOrThrow();
  if (!slip) {
    throw new Error('Packing slip data is empty.');
  }
  if (!slip.clientId || !slip.clientName) {
    throw new Error('Client selection is required.');
  }
  if (!slip.date) {
    throw new Error('Date is required.');
  }
  if (!Array.isArray(slip.items) || slip.items.length === 0) {
    throw new Error('Packing slip must contain at least one item.');
  }
  
  // Validate items on the backend
  var totalQty = 0;
  for (var i = 0; i < slip.items.length; i++) {
    var item = slip.items[i];
    if (!item.name || !item.name.trim()) {
      throw new Error('Product name in item ' + (i + 1) + ' is blank.');
    }
    var qty = Number(item.quantity);
    if (isNaN(qty) || qty <= 0) {
      throw new Error('Quantity for product "' + item.name + '" must be greater than 0.');
    }
    totalQty += qty;
  }
  slip.totalQuantity = totalQty;

  var sheet = getSheet(SHEETS.PACKING_SLIPS);
  var currentUser = getCurrentUserEmail();
  var originalCreatedBy = '';

  if (slip.slipCode) {
    // If it's an edit, find the old active row(s) and rename their slipCode to archive them
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === slip.slipCode) {
        originalCreatedBy = data[i][12] || currentUser;
        sheet.getRange(i + 1, 1).setValue(slip.slipCode + '-rev' + Date.now());
      }
    }
  } else {
    // Brand new slip
    slip.slipCode = getNextId(COUNTERS.SLIP, 'PS-');
  }
  slip.createdAt = new Date().toISOString();

  // Preserve original creator on edits; use current user for new slips
  slip.createdBy = originalCreatedBy || currentUser;
  slip.editedBy = currentUser;

  ensurePackingSlipsHeader(sheet);

  var orderStr = Array.isArray(slip.linkedOrders) ? slip.linkedOrders.join(', ') : (slip.linkedOrders || '');
  // Accept either free-floating invoice numbers (array) or a legacy single string.
  var invoiceNumbers = [];
  if (Array.isArray(slip.invoiceNumbers)) {
    invoiceNumbers = slip.invoiceNumbers.map(function(s) { return String(s).trim(); }).filter(Boolean);
  } else if (Array.isArray(slip.linkedInvoices) && slip.linkedInvoices.length) {
    invoiceNumbers = slip.linkedInvoices.map(function(s) { return String(s).trim(); }).filter(Boolean);
  }
  if (slip.invoiceNumber && typeof slip.invoiceNumber === 'string') {
    slip.invoiceNumber.split(',').forEach(function(part) {
      var s = String(part || '').trim();
      if (s && invoiceNumbers.indexOf(s) === -1) invoiceNumbers.push(s);
    });
  }
  var invoiceNumbersPrimary = joinInvoiceNumberList(invoiceNumbers);

  sheet.appendRow([
    slip.slipCode,
    slip.date,
    slip.clientId,
    slip.clientName,
    slip.clientCompany || '',
    slip.clientPhone || '',
    slip.clientAddress || '',
    JSON.stringify(slip.items),
    slip.totalQuantity,
    slip.notes || '',
    invoiceNumbersPrimary,
    slip.createdAt,
    slip.createdBy,
    slip.editedBy,
    orderStr,
    invoiceNumbersPrimary // mirror into legacy column for backward compatibility
  ]);

  // Reflect the value we persisted onto the returned object so the client stays in sync.
  slip.invoiceNumber = invoiceNumbersPrimary;
  slip.invoiceNumbers = invoiceNumbers;
  slip.linkedInvoices = invoiceNumbersPrimary;

  // Recalculate statuses for any linked orders (invoices live inside orders now).
  try {
    recalculateLinkedStatuses(slip.linkedOrders);
  } catch (e) {
    Logger.log("Recalculate order statuses failed: " + e.message);
  }

  return slip;
}

function getSuggestions(clientId) {
  checkLicenseOrThrow();
  if (!clientId) return [];
  var suggestions = {};

  // 1. Get from past packing slips
  var slips = getPackingSlips();
  for (var i = 0; i < slips.length; i++) {
    if (slips[i].clientId === clientId && slips[i].items) {
      for (var j = 0; j < slips[i].items.length; j++) {
        var name = String(slips[i].items[j].name || '').trim();
        if (name) suggestions[name] = true;
      }
    }
  }

  // 2. Get from client's Orders (and any associated invoice-number lists)
  try {
    var ordersList = getOrders();
    for (var i = 0; i < ordersList.length; i++) {
      if (ordersList[i].clientId === clientId && ordersList[i].items) {
        for (var j = 0; j < ordersList[i].items.length; j++) {
          var name = String(ordersList[i].items[j].name || '').trim();
          if (name) suggestions[name] = true;
        }
      }
    }
  } catch(e) {}

  return Object.keys(suggestions).sort();
}

/**
 * Returns a deduplicated list of invoice numbers recorded against any order or
 * any packing slip for a given client. Used by the slip-form picker so users
 * can quickly pick a known invoice number without re-typing it.
 */
function getClientInvoiceNumbers(clientId) {
  checkLicenseOrThrow();
  var out = [];
  var seen = {};
  if (!clientId) return out;

  var pushOne = function(val) {
    var s = String(val || '').trim();
    if (s && !seen[s]) {
      seen[s] = true;
      out.push(s);
    }
  };

  try {
    var ordersList = getOrders();
    ordersList.forEach(function(o) {
      if (o.clientId === clientId) {
        (o.invoiceNumbers || []).forEach(pushOne);
      }
    });
  } catch (e) {}

  try {
    var slipsList = getPackingSlips();
    slipsList.forEach(function(s) {
      if (s.clientId === clientId) {
        (s.invoiceNumbers || []).forEach(pushOne);
      }
    });
  } catch (e) {}

  return out.sort();
}

function duplicateSlip(slipCode) {
  checkLicenseOrThrow();
  if (!slipCode) return null;
  var slips = getPackingSlips();
  for (var i = 0; i < slips.length; i++) {
    if (slips[i].slipCode === slipCode) {
      var copy = JSON.parse(JSON.stringify(slips[i]));
      copy.slipCode = '';
      copy.createdAt = '';
      return copy;
    }
  }
  return null;
}

/* ------------------------------------------------------------------
 * Orders Database Operations
 * ------------------------------------------------------------------ */

function getOrders() {
  try {
    checkLicenseOrThrow();
    var sheet = getSheet(SHEETS.ORDERS);
    if (!sheet) {
      Logger.log("Orders sheet could not be opened or created.");
      return [];
    }
    ensureOrdersHeader(sheet);
    var values = sheet.getDataRange().getValues();
    var orders = [];
    for (var i = 1; i < values.length; i++) {
      if (values[i][0]) {
        var items = [];
        try {
          items = JSON.parse(values[i][4] || '[]');
        } catch (e) {
          items = [];
        }
        orders.push({
          id: values[i][0],
          date: normalizeDate(values[i][1]),
          clientId: values[i][2],
          clientName: values[i][3],
          items: items,
          status: values[i][5] || 'Pending',
          createdAt: values[i][6],
          invoiceNumbers: normalizeInvoiceNumberList(values[i][7])
        });
      }
    }
    return orders;
  } catch (err) {
    Logger.log("getOrders error: " + err.message);
    return [];
  }
}

function calculateStatusFromItems(items, defaultStatus) {
  if (!items || !Array.isArray(items) || !items.length) return defaultStatus;
  var allDelivered = true;
  var anyDelivered = false;
  items.forEach(function(item) {
    var qty = Number(item.quantity || 0);
    var shipped = Number(item.shipped || 0);
    if (shipped < qty) {
      allDelivered = false;
    }
    if (shipped > 0) {
      anyDelivered = true;
    }
  });
  if (allDelivered) return 'Completed';
  if (anyDelivered) return 'Partially Delivered';
  return defaultStatus;
}

function saveOrder(order) {
  checkLicenseOrThrow();
  if (!order) throw new Error('Order data is empty.');
  if (!order.clientId || !order.clientName) throw new Error('Client selection is required.');

  var sheet = getSheet(SHEETS.ORDERS);
  if (!sheet) throw new Error('Could not access or create the "Orders" sheet tab. Please make sure you have edit access to this spreadsheet.');

  ensureOrdersHeader(sheet);
  var values = sheet.getDataRange().getValues();
  order.createdAt = order.createdAt || new Date().toISOString();
  order.status = calculateStatusFromItems(order.items, 'Pending');

  // Normalize invoiceNumbers: accept array or comma-separated string for client compatibility.
  var normalizedInvoices = [];
  if (Array.isArray(order.invoiceNumbers)) {
    normalizedInvoices = order.invoiceNumbers.map(function(s) { return String(s).trim(); }).filter(Boolean);
  } else if (typeof order.invoiceNumbers === 'string') {
    normalizedInvoices = normalizeInvoiceNumberList(order.invoiceNumbers);
  }
  if (typeof order.invoiceNumbersRaw === 'string') {
    normalizeInvoiceNumberList(order.invoiceNumbersRaw).forEach(function(s) {
      if (normalizedInvoices.indexOf(s) === -1) normalizedInvoices.push(s);
    });
  }
  for (var n = 0; n < normalizedInvoices.length; n++) {
    if (!normalizedInvoices[n]) {
      normalizedInvoices.splice(n, 1);
      n--;
    }
  }
  order.invoiceNumbers = normalizedInvoices;
  var invoiceNumbersStr = joinInvoiceNumberList(normalizedInvoices);

  if (order.id) {
    // Update existing order
    var found = false;
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] === order.id) {
        sheet.getRange(i + 1, 2, 1, 6).setValues([[
          order.date || '',
          order.clientId,
          order.clientName,
          JSON.stringify(order.items || []),
          order.status,
          order.createdAt
        ]]);
        sheet.getRange(i + 1, 8).setValue(invoiceNumbersStr);
        found = true;
        break;
      }
    }
    if (!found) throw new Error('Order ID ' + order.id + ' not found for update.');
  } else {
    // Create new order
    if (order.customId) {
      // Check for uniqueness
      for (var i = 1; i < values.length; i++) {
        if (String(values[i][0]).toLowerCase().trim() === String(order.customId).toLowerCase().trim()) {
          throw new Error('Order Number ' + order.customId + ' already exists. Please choose a unique number.');
        }
      }
      order.id = order.customId;
    } else {
      order.id = getNextId(COUNTERS.ORDER, 'OR-');
    }

    sheet.appendRow([
      order.id,
      order.date || '',
      order.clientId,
      order.clientName,
      JSON.stringify(order.items || []),
      order.status,
      order.createdAt,
      invoiceNumbersStr
    ]);
  }
  return order;
}

function deleteOrder(orderId) {
  checkLicenseOrThrow();
  if (!orderId) return false;
  var sheet = getSheet(SHEETS.ORDERS);
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === orderId) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

/* ------------------------------------------------------------------
 * Invoices are a property of Orders and Packing Slips now, not a
 * standalone entity. The Invoices sheet is therefore no longer
 * initialised and dedicated invoice functions have been removed.
 * ------------------------------------------------------------------ */

/* ------------------------------------------------------------------
 * Balance & Linkage Statistics
 * ------------------------------------------------------------------ */

function getLinkedDeliveryStats(clientId, orderIds, excludeSlipCode) {
  checkLicenseOrThrow();

  // Clean IDs
  var targetOrders = [];
  if (orderIds) {
    if (Array.isArray(orderIds)) targetOrders = orderIds;
    else targetOrders = orderIds.split(',').map(function(s) { return s.trim(); }).filter(Boolean);
  }

  var stats = {}; // { productName: { ordered: 0, delivered: 0, description: '', orderId: '', orderDate: '' } }
  var excludeSet = {};
  if (excludeSlipCode) {
    String(excludeSlipCode).split(',').forEach(function (s) {
      var t = String(s || '').trim();
      if (t) excludeSet[t] = true;
    });
  }

  // 1. Gather Ordered quantities (and copy the canonical description / order date)
  if (targetOrders.length > 0) {
    var dbOrders = getOrders();
    dbOrders.forEach(function(order) {
      if (targetOrders.indexOf(order.id) !== -1) {
        (order.items || []).forEach(function(item) {
          var name = String(item.name || '').trim();
          if (name) {
            if (!stats[name]) stats[name] = { ordered: 0, delivered: 0, description: '', orderId: order.id, orderDate: order.date || '' };
            stats[name].ordered += Number(item.quantity || 0);
            if (!stats[name].description && item.description) {
              stats[name].description = String(item.description);
            }
          }
        });
      }
    });
  }

  // 2. Sum up Delivered quantities from all active packing slips.
  //    Slips in `excludeSet` (typically the one currently being edited) are
  //    skipped so their own qty does not double-count against row hints.
  var dbSlips = getPackingSlips();
  dbSlips.forEach(function(slip) {
    if (slip.slipCode.indexOf('-rev') !== -1) return;
    if (excludeSet[slip.slipCode]) return;
    (slip.items || []).forEach(function(item) {
      var name = String(item.name || '').trim();
      if (name) {
        var matchesOrder = item.orderId && targetOrders.indexOf(item.orderId) !== -1;
        if (matchesOrder) {
          if (!stats[name]) stats[name] = { ordered: 0, delivered: 0, description: '', orderId: item.orderId, orderDate: '' };
          stats[name].delivered += Number(item.quantity || 0);
          if (!stats[name].description && item.description) {
            stats[name].description = String(item.description);
          }
        }
      }
    });
  });

  return stats;
}

/**
 * Returns the linked order details (id + date) for a slip. Powers the
 * Order Date vs Dispatch Date display on the printed/preview packing slip.
 */
function getLinkedOrderMetadata(orderIds) {
  checkLicenseOrThrow();
  var ids = [];
  if (orderIds) {
    if (Array.isArray(orderIds)) ids = orderIds;
    else ids = String(orderIds || '').split(',').map(function (s) { return String(s || '').trim(); }).filter(Boolean);
  }
  if (!ids.length) return [];
  var dbOrders = getOrders();
  var out = [];
  ids.forEach(function (id) {
    var found = null;
    for (var i = 0; i < dbOrders.length; i++) {
      if (dbOrders[i].id === id) { found = dbOrders[i]; break; }
    }
    if (found) {
      out.push({ id: found.id, date: found.date || '', clientId: found.clientId, clientName: found.clientName });
    } else {
      out.push({ id: id, date: '', clientId: '', clientName: '' });
    }
  });
  return out;
}

/**
 * Recalculates order statuses (Pending / Partially Delivered / Completed)
 * after a packing slip has been saved. Called from savePackingSlip so that
 * order statuses stay in sync with delivered quantities on linked slips.
 */
function recalculateLinkedStatuses(orderIds) {
  var targetOrders = [];
  if (orderIds) {
    if (Array.isArray(orderIds)) targetOrders = orderIds;
    else targetOrders = orderIds.split(',').map(function(s) { return s.trim(); }).filter(Boolean);
  }
  if (!targetOrders.length) return;

  var dbSlips = getPackingSlips().filter(function(s) { return s.slipCode.indexOf('-rev') === -1; });

  var orderSheet = getSheet(SHEETS.ORDERS);
  if (!orderSheet) return;
  var orderValues = orderSheet.getDataRange().getValues();

  targetOrders.forEach(function(orderId) {
    var foundIndex = -1;
    var orderObj = null;
    for (var i = 1; i < orderValues.length; i++) {
      if (orderValues[i][0] === orderId) {
        foundIndex = i + 1;
        try {
          orderObj = {
            items: JSON.parse(orderValues[i][4] || '[]'),
            status: orderValues[i][5]
          };
        } catch (e) {}
        break;
      }
    }

    var orderItems = orderObj && Array.isArray(orderObj.items) ? orderObj.items : [];
    if (foundIndex === -1 || !orderObj || !orderItems.length) return;

    var allDelivered = true;
    var anyDelivered = false;

    orderItems.forEach(function(orderItem) {
      var itemName = String(orderItem.name || '').trim();
      var orderedQty = Number(orderItem.quantity || 0);
      var deliveredQty = 0;

      dbSlips.forEach(function(slip) {
        (slip.items || []).forEach(function(slipItem) {
          if (String(slipItem.name || '').trim() === itemName && slipItem.orderId === orderId) {
            deliveredQty += Number(slipItem.quantity || 0);
          }
        });
      });

      orderItem.shipped = deliveredQty;

      if (deliveredQty < orderedQty) {
        allDelivered = false;
      }
      if (deliveredQty > 0) {
        anyDelivered = true;
      }
    });

    var newStatus = 'Pending';
    if (allDelivered) {
      newStatus = 'Completed';
    } else if (anyDelivered) {
      newStatus = 'Partially Delivered';
    }

    orderSheet.getRange(foundIndex, 5, 1, 2).setValues([[
      JSON.stringify(orderItems),
      newStatus
    ]]);
  });
}

function runCentral(functionName, args) {
  if (typeof globalThis[functionName] === 'function') {
    return globalThis[functionName].apply(null, args || []);
  }
  throw new Error("Function " + functionName + " is not defined in Central Library.");
}
