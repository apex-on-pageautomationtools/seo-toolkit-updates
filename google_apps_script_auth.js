/*
 * SEO Toolkit Pro — Auth & License Server + GSC Backend
 * Deploy as Google Apps Script Web App (Execute as: Me, Access: Anyone)
 * OAuth2 Library ID: 1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF
 *
 * Google Sheet columns (Sheet name: "users"):
 *   A: Email  B: Password  C: Device ID  D: Approved  E: Name
 *   F: Created Date  G: Last Login  H: Notes  I: Admin (TRUE/FALSE)
 *
 * Sheet name: "config":
 *   A1: min_version   B1: 3.2
 *
 * GSC Tabs:
 *   GSC_Config:   Domain | GSC Email | Account Key | Access Level | Last Synced | GSC Property
 *   GSC_Accounts: Account Key | GSC Email | Report Recipient
 *   GSC_Settings: Setting | Value | Description
 *   Auth_Status:  Email | Account Key | Auth URL | Status
 *   CrawlHistory: BatchId | UserId | UserName | UrlCount | Results | CreatedAt
 *   Log_YYYY_MM:  Monthly audit log
 *
 * Required Script Properties: CLIENT_ID, CLIENT_SECRET
 * (SHEET_ID is hardcoded below; override via Script Properties if needed)
 */

var SHEET_ID = '1etkOCwKPWpnJ0wSEKERWSxNGj5x6hteswTkRv1ex2y0';
var ADMIN_KEY = '44bXCd9yaa#xMR!E';

function _getProp(name) {
  return PropertiesService.getScriptProperties().getProperty(name);
}
function _getSheetId() { return _getProp('SHEET_ID') || SHEET_ID; }
function _getClientId() { return _getProp('CLIENT_ID'); }
function _getClientSecret() { return _getProp('CLIENT_SECRET'); }

var CACHE = CacheService.getScriptCache();

/* =====================================================
   ROUTING — doGet / doPost
   ===================================================== */
function doGet(e) {
  var params = e ? (e.parameter || {}) : {};
  var action = params.action || '';

  // Auth actions (query-string based)
  if (action === 'login') return _login(params);
  if (action === 'change_password') return _changePassword(params);
  if (action === 'version_check') return _versionCheck(params);
  if (action === 'admin_approve') return _adminApprove(params);
  if (action === 'admin_reject') return _adminReject(params);
  if (action === 'admin_list') return _adminList(params);
  if (action === 'admin_add_user') return _adminAddUser(params);
  if (action === 'admin_change_password') return _adminChangePassword(params);
  if (action === 'admin_update_mac') return _adminUpdateMac(params);
  if (action === 'admin_remove_user') return _adminRemoveUser(params);

  // GSC read actions
  if (action === 'get_config') return _json(getConfigForExtension());
  if (action === 'get_settings') return handleGetSettings();
  if (action === 'ping') return _json({ status: 'ok', timestamp: new Date().toISOString() });
  if (action === 'get_crawl_history') return handleGetCrawlHistory(params);
  if (action === 'get_auth_status') return _json(handleGetAuthStatus());
  if (action === 'get_accounts') return _json(handleGetAccounts());

  return _json({ status: 'SEO Toolkit Pro — Active' });
}

function doPost(e) {
  try {
    var contentType = e.postData ? e.postData.type : '';
    // JSON body (GSC actions)
    if (contentType && contentType.indexOf('json') !== -1) {
      var data = JSON.parse(e.postData.contents);
      var action = (data.action || '').toString();

      if (action === 'send_alert') return _json(handleSendAlert(data));
      if (action === 'save_settings') return _json(handleSaveSettings(data.settings));
      if (action === 'validate_batch') return _json(handleValidateBatch(data));
      if (action === 'inspect_single') return _json(handleInspectSingle(data));
      if (action === 'save_crawl_batch') return _json(handleSaveCrawlBatch(data));
      if (action === 'sync_domains') { syncDomainsFromGSC(); return _json({ success: true }); }
      if (action === 'generate_auth_urls') { generateAuthSheet(); return _json({ success: true }); }
      if (action === 'recheck_auth') return _json(recheckAuthStatus());

      return _json({ error: 'Unknown action: ' + action });
    }
    // Query-string fallback (auth actions via POST)
    return doGet(e);
  } catch (err) {
    return _json({ error: err.message });
  }
}

/* =====================================================
   AUTH — Login, User Management
   ===================================================== */
// An approved user may log in from up to this many of their OWN devices - Device
// ID (column C) is a comma-separated list, not a single value. First login on a new
// device auto-claims a free slot; once all slots are used, a new device is rejected
// until an admin frees one (Admin -> Settings -> Device ID, or admin_update_mac).
var MAX_DEVICES_PER_USER = 3;

function _parseDeviceList(raw) {
  return (raw || '').toString().split(',').map(function(m) { return m.trim().toUpperCase(); })
    .filter(function(m) { return m; });
}

function _login(p) {
  var email = (p.email || '').trim().toLowerCase();
  var password = p.password || '';
  var mac = (p.mac || '').trim().toUpperCase();

  if (!email || !password || !mac) return _json({error: 'Missing email, password, or Device ID'});

  var sheet = _getSheet('users');
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var rowEmail = (data[i][0] || '').toString().trim().toLowerCase();
    var rowPass = (data[i][1] || '').toString();
    var devices = _parseDeviceList(data[i][2]);
    var rowApproved = data[i][3];

    if (rowEmail === email && rowPass === password) {
      var known = devices.indexOf(mac) !== -1;
      if (known || devices.length < MAX_DEVICES_PER_USER) {
        if (!known) {
          devices.push(mac);
          sheet.getRange(i + 1, 3).setValue(devices.join(', '));
        }
        sheet.getRange(i + 1, 7).setValue(new Date());

        if (rowApproved === true || rowApproved === 'TRUE' || rowApproved === 'Yes') {
          var rowAdmin = data[i][8];
          var isAdmin = (rowAdmin === true || rowAdmin === 'TRUE' || rowAdmin === 'Yes');
          return _json({status: 'approved', email: email, mac: mac, is_admin: isAdmin});
        } else {
          return _json({status: 'pending', email: email, mac: mac,
                        message: 'Your account is pending admin approval. Contact your administrator.'});
        }
      } else {
        return _json({status: 'mac_mismatch', email: email, registered_mac: devices.join(', '), current_mac: mac,
                      message: 'This account is already registered on ' + MAX_DEVICES_PER_USER +
                        ' device(s). Contact admin to free up a device.'});
      }
    }
  }
  return _json({status: 'invalid', message: 'Invalid email or password.'});
}

function _adminAddUser(p) {
  if (p.admin_key !== ADMIN_KEY) return _json({error: 'Unauthorized'});
  var email = (p.email || '').trim().toLowerCase();
  var password = p.password || '';
  var mac = (p.mac || '').trim().toUpperCase();
  var name = p.name || '';
  if (!email || !password) return _json({error: 'Email and password required'});

  var sheet = _getSheet('users');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().trim().toLowerCase() === email) {
      return _json({error: 'User already exists: ' + email});
    }
  }
  sheet.appendRow([email, password, mac, true, name, new Date(), '', 'Added by admin']);
  return _json({status: 'added', email: email, mac: mac});
}

function _changePassword(p) {
  // Self-service password change: verify the user's CURRENT password, then set the new one.
  var email = (p.email || '').trim().toLowerCase();
  var oldPass = p.old_password || '';
  var newPass = p.new_password || '';
  if (!email || !oldPass || !newPass) return _json({error: 'Email, old_password, and new_password required'});
  if (newPass.length < 6) return _json({status: 'error', message: 'New password must be at least 6 characters.'});

  var sheet = _getSheet('users');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().trim().toLowerCase() === email) {
      if ((data[i][1] || '').toString() !== oldPass) {
        return _json({status: 'invalid', message: 'Current password is incorrect.'});
      }
      sheet.getRange(i + 1, 2).setValue(newPass);
      return _json({status: 'password_changed', email: email, message: 'Password changed successfully.'});
    }
  }
  return _json({error: 'User not found'});
}

function _adminChangePassword(p) {
  if (p.admin_key !== ADMIN_KEY) return _json({error: 'Unauthorized'});
  var email = (p.email || '').trim().toLowerCase();
  var newPass = p.new_password || '';
  if (!email || !newPass) return _json({error: 'Email and new_password required'});

  var sheet = _getSheet('users');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().trim().toLowerCase() === email) {
      sheet.getRange(i + 1, 2).setValue(newPass);
      return _json({status: 'password_changed', email: email});
    }
  }
  return _json({error: 'User not found'});
}

function _adminUpdateMac(p) {
  if (p.admin_key !== ADMIN_KEY) return _json({error: 'Unauthorized'});
  var email = (p.email || '').trim().toLowerCase();
  var mac = (p.mac || '').trim().toUpperCase();
  if (!email) return _json({error: 'Email required'});

  var sheet = _getSheet('users');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().trim().toLowerCase() === email) {
      sheet.getRange(i + 1, 3).setValue(mac);
      return _json({status: 'mac_updated', email: email, mac: mac});
    }
  }
  return _json({error: 'User not found'});
}

function _adminRemoveUser(p) {
  if (p.admin_key !== ADMIN_KEY) return _json({error: 'Unauthorized'});
  var email = (p.email || '').trim().toLowerCase();
  if (!email) return _json({error: 'Email required'});

  var sheet = _getSheet('users');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().trim().toLowerCase() === email) {
      sheet.deleteRow(i + 1);
      return _json({status: 'removed', email: email});
    }
  }
  return _json({error: 'User not found'});
}

function _versionCheck(p) {
  var version = p.version || '0';
  var configSheet = _getSheet('config');
  var minVersion = configSheet.getRange('B1').getValue().toString();
  var allowed = _compareVersions(version, minVersion) >= 0;
  return _json({allowed: allowed, min_version: minVersion, your_version: version,
                message: allowed ? 'OK' : 'Your version is outdated. Please update to v' + minVersion + ' or later.'});
}

function _adminApprove(p) {
  if (p.admin_key !== ADMIN_KEY) return _json({error: 'Unauthorized'});
  var email = (p.email || '').trim().toLowerCase();
  var sheet = _getSheet('users');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().trim().toLowerCase() === email) {
      sheet.getRange(i + 1, 4).setValue(true);
      return _json({status: 'approved', email: email});
    }
  }
  return _json({error: 'User not found'});
}

function _adminReject(p) {
  if (p.admin_key !== ADMIN_KEY) return _json({error: 'Unauthorized'});
  var email = (p.email || '').trim().toLowerCase();
  var sheet = _getSheet('users');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().trim().toLowerCase() === email) {
      sheet.getRange(i + 1, 4).setValue(false);
      return _json({status: 'rejected', email: email});
    }
  }
  return _json({error: 'User not found'});
}

function _adminList(p) {
  if (p.admin_key !== ADMIN_KEY) return _json({error: 'Unauthorized'});
  var sheet = _getSheet('users');
  var data = sheet.getDataRange().getValues();
  var users = [];
  for (var i = 1; i < data.length; i++) {
    users.push({
      email: data[i][0], mac: data[i][2], approved: data[i][3],
      name: data[i][4], registered: data[i][5], last_login: data[i][6],
      notes: data[i][7],
      is_admin: data[i][8] === true || data[i][8] === 'TRUE' || data[i][8] === 'Yes'
    });
  }
  return _json({users: users});
}

/* =====================================================
   GSC — OAuth2 per account
   ===================================================== */
function getOAuthService(accountKey) {
  return OAuth2.createService(accountKey)
    .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/auth')
    .setTokenUrl('https://oauth2.googleapis.com/token')
    .setClientId(_getClientId())
    .setClientSecret(_getClientSecret())
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getScriptProperties())
    .setScope(['https://www.googleapis.com/auth/webmasters.readonly'])
    .setParam('access_type', 'offline')
    .setParam('prompt', 'consent')
    .setParam('login_hint', getEmailForKey(accountKey));
}

function authCallback(request) {
  var accountKey = request.parameter.accountKey;
  var service = getOAuthService(accountKey);
  var authorized = service.handleCallback(request);
  return HtmlService.createHtmlOutput(
    authorized
      ? 'Account authorised: ' + accountKey + '. You can close this tab.'
      : 'Authorisation failed for: ' + accountKey
  );
}

/* =====================================================
   GSC — Lookups
   ===================================================== */
function getEmailForKey(accountKey) {
  try {
    var acct = getAccountsMap()[accountKey];
    if (acct && acct.email) return acct.email;
  } catch (e) {}
  var ss = SpreadsheetApp.openById(_getSheetId());
  var config = ss.getSheetByName('GSC_Config');
  if (!config) return null;
  var data = config.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][2] && data[i][2].toString().trim() === accountKey) {
      return data[i][1].toString().trim();
    }
  }
  return null;
}

function getAccountKeyForDomain(domain) {
  var ss = SpreadsheetApp.openById(_getSheetId());
  var config = ss.getSheetByName('GSC_Config');
  if (!config) return null;
  var data = config.getDataRange().getValues();
  var target = domain.toString().trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === target) {
      return data[i][2] ? data[i][2].toString().trim() : null;
    }
  }
  return null;
}

function extractDomainFromUrl(url) {
  try {
    var match = url.match(/^https?:\/\/([^\/]+)/i);
    if (!match) return null;
    return match[1].toLowerCase().replace(/^www\./, '');
  } catch (e) { return null; }
}

/* =====================================================
   GSC — Property resolution + caching
   ===================================================== */
function resolveSiteProperty(domain, accessToken, sampleUrl) {
  var response = UrlFetchApp.fetch('https://www.googleapis.com/webmasters/v3/sites', {
    method: 'GET', headers: { 'Authorization': 'Bearer ' + accessToken }, muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) {
    return { error: 'Could not list GSC properties (HTTP ' + response.getResponseCode() + ')' };
  }
  var entries = (JSON.parse(response.getContentText()).siteEntry) || [];
  var bareHost = function(s) {
    return s.toLowerCase().replace(/^sc-domain:/, '').replace(/^https?:\/\//, '').replace(/^www\./, '').replace(/\/.*$/, '');
  };
  var target = bareHost(domain);

  for (var i = 0; i < entries.length; i++) {
    if (entries[i].siteUrl.indexOf('sc-domain:') === 0 && bareHost(entries[i].siteUrl) === target)
      return { siteUrl: entries[i].siteUrl, type: 'domain' };
  }
  var hostPrefixes = [];
  var sample = sampleUrl ? sampleUrl.toLowerCase() : '';
  for (var j = 0; j < entries.length; j++) {
    var su = entries[j].siteUrl;
    if (su.indexOf('sc-domain:') !== 0 && bareHost(su) === target) {
      hostPrefixes.push(su);
      if (sample && sample.indexOf(su.toLowerCase()) === 0) return { siteUrl: su, type: 'prefix' };
    }
  }
  if (hostPrefixes.length > 0) return { mismatch: true, available: hostPrefixes };
  return { notFound: true };
}

function getCachedSiteProperty(domain, accountKey, accessToken, sampleUrl) {
  var origin = (sampleUrl.match(/^https?:\/\/[^\/]+/i) || [''])[0].toLowerCase();
  var cacheKey = 'prop_' + accountKey + '_' + origin;
  var cached = CACHE.get(cacheKey);
  if (cached) { try { return JSON.parse(cached); } catch (e) {} }
  var prop = resolveSiteProperty(domain, accessToken, sampleUrl);
  if (prop.siteUrl) CACHE.put(cacheKey, JSON.stringify(prop), 1800);
  return prop;
}

function inspectUrl(inspectionUrl, siteUrl, accessToken) {
  var response = UrlFetchApp.fetch(
    'https://searchconsole.googleapis.com/v1/urlInspection/index:inspect',
    {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + accessToken, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ inspectionUrl: inspectionUrl, siteUrl: siteUrl }),
      muteHttpExceptions: true
    }
  );
  var body = {};
  try { body = JSON.parse(response.getContentText() || '{}'); } catch (e) {}
  return { code: response.getResponseCode(), body: body };
}

/* =====================================================
   GSC — Batch progress tracking
   ===================================================== */
function getBatchProgress(batchId) {
  var raw = CACHE.get('batch_' + batchId);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function setBatchProgress(batchId, data) {
  CACHE.put('batch_' + batchId, JSON.stringify(data), 21600);
}

function incrementBatchProcessed(batchId) {
  var progress = getBatchProgress(batchId);
  if (!progress) return null;
  progress.processed = (progress.processed || 0) + 1;
  setBatchProgress(batchId, progress);
  return progress;
}

/* =====================================================
   GSC — Settings tab helpers
   ===================================================== */
function getSettings() {
  var ss = SpreadsheetApp.openById(_getSheetId());
  var settingsSheet = ss.getSheetByName('GSC_Settings');
  if (!settingsSheet) settingsSheet = createSettingsTab();
  var data = settingsSheet.getDataRange().getValues();
  var settings = {};
  for (var i = 1; i < data.length; i++) {
    var key = data[i][0] ? data[i][0].toString().trim() : '';
    var value = data[i][1] !== undefined && data[i][1] !== null ? data[i][1].toString().trim() : '';
    if (key) settings[key] = value;
  }
  return settings;
}

function getNotificationEmails() {
  var settings = getSettings();
  var raw = settings['Notification Emails'] || '';
  return raw.split(',').map(function(e) { return e.trim(); }).filter(function(e) { return e.length > 0 && e.indexOf('@') !== -1; });
}

function createSettingsTab() {
  var ss = SpreadsheetApp.openById(_getSheetId());
  var settingsSheet = ss.getSheetByName('GSC_Settings');
  if (settingsSheet) return settingsSheet;

  settingsSheet = ss.insertSheet('GSC_Settings');
  var rows = [
    ['Setting', 'Value', 'Description'],
    ['Notification Emails', '', 'Comma-separated fallback emails for sync summary.'],
    ['Send Weekly Summary', 'Yes', 'Yes or No.'],
    ['Alert Email TO', '', 'Comma-separated primary recipients for audit alerts.'],
    ['Alert Email CC', '', 'Comma-separated CC recipients for audit alerts.'],
  ];
  settingsSheet.getRange(1, 1, rows.length, 3).setValues(rows);
  settingsSheet.getRange('A1:C1').setFontWeight('bold').setBackground('#1e1e2e').setFontColor('#00ff88');
  settingsSheet.setColumnWidth(1, 200);
  settingsSheet.setColumnWidth(2, 400);
  settingsSheet.setColumnWidth(3, 500);
  settingsSheet.setFrozenRows(1);
  return settingsSheet;
}

function handleGetSettings() {
  var settings = getSettings();
  return _json({
    success: true,
    settings: {
      notificationEmails: settings['Notification Emails'] || '',
      sendWeeklySummary: settings['Send Weekly Summary'] || 'Yes',
      alertEmailTo: settings['Alert Email TO'] || '',
      alertEmailCc: settings['Alert Email CC'] || ''
    }
  });
}

function handleSaveSettings(settings) {
  if (!settings) return { success: false, error: 'No settings provided' };
  var ss = SpreadsheetApp.openById(_getSheetId());
  var sheet = ss.getSheetByName('GSC_Settings');
  if (!sheet) sheet = createSettingsTab();

  var data = sheet.getDataRange().getValues();
  var keyMap = {
    'Notification Emails': 'notificationEmails',
    'Send Weekly Summary': 'sendWeeklySummary',
    'Alert Email TO': 'alertEmailTo',
    'Alert Email CC': 'alertEmailCc'
  };
  for (var i = 1; i < data.length; i++) {
    var settingName = String(data[i][0]).trim();
    var jsKey = keyMap[settingName];
    if (jsKey && settings[jsKey] !== undefined) {
      sheet.getRange(i + 1, 2).setValue(settings[jsKey]);
    }
  }
  return { success: true };
}

/* =====================================================
   GSC — Accounts registry
   ===================================================== */
function getAccountsTab() {
  var ss = SpreadsheetApp.openById(_getSheetId());
  var sh = ss.getSheetByName('GSC_Accounts');
  if (!sh) {
    sh = ss.insertSheet('GSC_Accounts');
    sh.getRange(1, 1, 1, 3).setValues([['Account Key', 'GSC Email', 'Report Recipient']])
      .setFontWeight('bold').setBackground('#1e1e2e').setFontColor('#00ff88');
    sh.setColumnWidth(1, 260); sh.setColumnWidth(2, 280); sh.setColumnWidth(3, 320);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getAccountsMap() {
  var cached = CACHE.get('accounts_map_v1');
  if (cached) { try { return JSON.parse(cached); } catch (e) {} }
  var sh = getAccountsTab();
  var data = sh.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var key = data[i][0] ? data[i][0].toString().trim() : '';
    if (!key) continue;
    map[key] = { row: i + 1, key: key,
      email: data[i][1] ? data[i][1].toString().trim() : '',
      recipient: data[i][2] ? data[i][2].toString().trim() : '' };
  }
  CACHE.put('accounts_map_v1', JSON.stringify(map), 300);
  return map;
}

function syncAccountsRegistry() {
  var ss = SpreadsheetApp.openById(_getSheetId());
  var config = ss.getSheetByName('GSC_Config');
  if (!config) throw new Error('GSC_Config tab missing.');
  var cfg = config.getDataRange().getValues();
  var distinct = {};
  for (var i = 1; i < cfg.length; i++) {
    var email = cfg[i][1] ? cfg[i][1].toString().trim() : '';
    var key = cfg[i][2] ? cfg[i][2].toString().trim() : '';
    if (!key) continue;
    if (!distinct[key]) distinct[key] = email;
  }
  var sh = getAccountsTab();
  CACHE.remove('accounts_map_v1');
  var existing = getAccountsMap();
  var newRows = [];
  Object.keys(distinct).forEach(function(key) {
    if (existing[key]) {
      if (!existing[key].email && distinct[key]) sh.getRange(existing[key].row, 2).setValue(distinct[key]);
      return;
    }
    newRows.push([key, distinct[key] || '', '']);
  });
  if (newRows.length > 0) sh.getRange(sh.getLastRow() + 1, 1, newRows.length, 3).setValues(newRows);
  CACHE.remove('accounts_map_v1');
  return newRows.length;
}

function handleGetAccounts() {
  var map = getAccountsMap();
  var accounts = [];
  Object.keys(map).forEach(function(key) {
    var acct = map[key];
    var service = getOAuthService(key);
    accounts.push({
      key: key,
      email: acct.email,
      recipient: acct.recipient,
      authorized: service.hasAccess()
    });
  });
  return { success: true, accounts: accounts };
}

/* =====================================================
   GSC — Auth status helpers
   ===================================================== */
function generateAuthSheet() {
  var ss = SpreadsheetApp.openById(_getSheetId());
  var config = ss.getSheetByName('GSC_Config');
  if (!config) return;
  var data = config.getDataRange().getValues();
  var seen = {};
  var accounts = [];
  for (var i = 1; i < data.length; i++) {
    var email = data[i][1] ? data[i][1].toString().trim() : '';
    var key = data[i][2] ? data[i][2].toString().trim() : '';
    if (!email || !key || seen[key]) continue;
    seen[key] = true;
    accounts.push({ email: email, key: key });
  }
  var authSheet = ss.getSheetByName('Auth_Status');
  if (authSheet) { authSheet.clear(); } else { authSheet = ss.insertSheet('Auth_Status'); }
  authSheet.getRange(1, 1, 1, 4).setValues([['Email', 'Account Key', 'Auth URL', 'Status']]);
  authSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#1e1e2e').setFontColor('#00ff88');

  var rows = [];
  accounts.forEach(function(a) {
    var service = getOAuthService(a.key);
    if (service.hasAccess()) {
      rows.push([a.email, a.key, '', 'Authorised']);
    } else {
      rows.push([a.email, a.key, service.getAuthorizationUrl({ accountKey: a.key }), 'Pending']);
    }
  });
  if (rows.length > 0) authSheet.getRange(2, 1, rows.length, 4).setValues(rows);
  authSheet.setColumnWidth(1, 280); authSheet.setColumnWidth(2, 240);
  authSheet.setColumnWidth(3, 520); authSheet.setColumnWidth(4, 110);
  authSheet.setFrozenRows(1);
}

function recheckAuthStatus() {
  var ss = SpreadsheetApp.openById(_getSheetId());
  var config = ss.getSheetByName('GSC_Config');
  if (!config) return { done: 0, pending: 0, pendingKeys: [] };
  var data = config.getDataRange().getValues();
  var seen = {};
  var keys = [];
  for (var i = 1; i < data.length; i++) {
    var key = data[i][2] ? data[i][2].toString().trim() : '';
    if (!key || seen[key]) continue;
    seen[key] = true;
    keys.push(key);
  }
  var result = {};
  var done = 0, pending = 0;
  var pendingKeys = [];
  keys.forEach(function(key) {
    var ok = getOAuthService(key).hasAccess();
    result[key] = ok;
    if (ok) done++; else { pending++; pendingKeys.push(key); }
  });
  return { done: done, pending: pending, pendingKeys: pendingKeys };
}

function handleGetAuthStatus() {
  var ss = SpreadsheetApp.openById(_getSheetId());
  var authSheet = ss.getSheetByName('Auth_Status');
  if (!authSheet) return { success: true, accounts: [] };
  var data = authSheet.getDataRange().getValues();
  var accounts = [];
  for (var i = 1; i < data.length; i++) {
    accounts.push({
      email: data[i][0] || '',
      key: data[i][1] || '',
      authUrl: data[i][2] || '',
      status: data[i][3] || ''
    });
  }
  return { success: true, accounts: accounts };
}

/* =====================================================
   GSC — Config for extension / domain mapping
   ===================================================== */
function getConfigForExtension() {
  try {
    var ss = SpreadsheetApp.openById(_getSheetId());
    var config = ss.getSheetByName('GSC_Config');
    if (!config) return { success: false, error: 'GSC_Config tab not found' };
    var data = config.getDataRange().getValues();
    if (data.length < 2) return { success: true, mapping: {}, domains: [] };

    var accessRank = {
      'Full (domain)': 100, 'Full (URL prefix)': 90,
      'Owner (domain)': 80, 'Owner (URL prefix)': 70,
      'Restricted (domain)': 60, 'Restricted (URL prefix)': 50
    };
    var best = {};
    for (var i = 1; i < data.length; i++) {
      var domain = data[i][0] ? data[i][0].toString().trim().toLowerCase() : '';
      var email = data[i][1] ? data[i][1].toString().trim().toLowerCase() : '';
      var accountKey = data[i][2] ? data[i][2].toString().trim() : '';
      var accessLevel = data[i][3] ? data[i][3].toString().trim() : '';
      if (!domain || !email || domain === '_setup') continue;
      if (accessLevel.indexOf('Removed') !== -1 || accessLevel.indexOf('Unverified') !== -1) continue;
      var rank = accessRank[accessLevel] || 0;
      if (!best[domain] || rank > best[domain].rank) {
        best[domain] = { email: email, accountKey: accountKey, accessLevel: accessLevel, rank: rank };
      }
    }
    var mapping = {};
    Object.keys(best).forEach(function(d) {
      var b = best[d];
      mapping[d] = { email: b.email, accountKey: b.accountKey, accessLevel: b.accessLevel };
    });
    return { success: true, mapping: mapping, domains: Object.keys(mapping).sort(), total: Object.keys(mapping).length };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/* =====================================================
   GSC — Validate batch (crawl tracker)
   ===================================================== */
function handleValidateBatch(data) {
  var urls = Array.isArray(data.urls) ? data.urls : [];
  if (urls.length === 0) return { success: false, error: 'No URLs provided.' };
  if (urls.length > 100) return { success: false, error: 'Maximum 100 URLs allowed.' };

  var items = urls.map(function(url) {
    var cleanUrl = url.toString().trim();
    if (!cleanUrl) return { url: cleanUrl, ready: false, error: 'Empty URL' };
    var domain = extractDomainFromUrl(cleanUrl);
    if (!domain) return { url: cleanUrl, ready: false, error: 'Invalid URL format' };
    var accountKey = getAccountKeyForDomain(domain);
    if (!accountKey) return { url: cleanUrl, domain: domain, ready: false, error: 'Domain not in master list.' };
    var service = getOAuthService(accountKey);
    if (!service.hasAccess()) return { url: cleanUrl, domain: domain, accountKey: accountKey, ready: false, error: 'Account not authorised.' };
    return { url: cleanUrl, domain: domain, accountKey: accountKey, ready: true };
  });

  var readyCount = items.filter(function(i) { return i.ready; }).length;
  var uniqueDomains = [];
  var seenDomains = {};
  items.forEach(function(i) {
    if (i.domain && !seenDomains[i.domain]) { seenDomains[i.domain] = true; uniqueDomains.push(i.domain); }
  });

  var logRow = logRequest({
    domain: uniqueDomains.join(', ').substring(0, 250),
    urlCount: items.length,
    status: 'Batch Validated',
    notes: readyCount + ' ready, ' + (items.length - readyCount) + ' will fail'
  });

  var batchId = Utilities.getUuid();
  setBatchProgress(batchId, { logRow: logRow, readyCount: readyCount, processed: 0 });
  return { success: true, batchId: batchId, items: items, totalUrls: items.length, readyCount: readyCount };
}

/* =====================================================
   GSC — Inspect single URL (crawl tracker)
   ===================================================== */
function handleInspectSingle(data) {
  var url = (data.url || '').toString().trim();
  var domain = (data.domain || '').toString().trim();
  var accountKey = (data.accountKey || '').toString().trim();
  var batchId = (data.batchId || '').toString().trim();
  if (!url || !domain || !accountKey) return { success: false, error: 'Missing url, domain, or accountKey' };

  var service = getOAuthService(accountKey);
  if (!service.hasAccess()) return { success: false, url: url, error: 'Account not authorised.' };
  var accessToken = service.getAccessToken();
  var acctEmail = getEmailForKey(accountKey) || accountKey;

  var prop = getCachedSiteProperty(domain, accountKey, accessToken, url);
  if (prop.error) return { success: false, url: url, error: prop.error };
  if (prop.notFound) return { success: false, url: url, error: 'No GSC property for ' + domain + ' on ' + acctEmail + '.' };
  if (prop.mismatch) return { success: false, url: url, error: 'URL version not covered by ' + acctEmail + '. Available: ' + prop.available.join(', ') };

  var res = inspectUrl(url, prop.siteUrl, accessToken);
  if (res.code !== 200) {
    var msg = (res.body.error && res.body.error.message) || ('HTTP ' + res.code);
    return { success: false, url: url, error: 'GSC API ' + res.code + ' on ' + prop.siteUrl + ': ' + msg };
  }

  var r = res.body.inspectionResult && res.body.inspectionResult.indexStatusResult;
  var crawlTime = r && r.lastCrawlTime;
  var coverage = r && r.coverageState;
  var verdict = r && r.verdict;

  if (!crawlTime) {
    var reason = coverage || (verdict ? ('verdict ' + verdict) : 'no crawl data returned');
    if (batchId) { var p = incrementBatchProcessed(batchId); if (p && p.logRow) updateLogProcessed(p.logRow, p.processed); }
    return { success: true, url: url, domain: domain, accountKey: accountKey,
             lastCrawlDate: 'Never Crawled', indexStatus: reason, property: prop.siteUrl };
  }

  var formatted = Utilities.formatDate(new Date(crawlTime), 'Asia/Kolkata', 'dd MMM yyyy');
  if (batchId) { var p2 = incrementBatchProcessed(batchId); if (p2 && p2.logRow) updateLogProcessed(p2.logRow, p2.processed); }
  return { success: true, url: url, domain: domain, accountKey: accountKey,
           lastCrawlDate: formatted, indexStatus: coverage || 'Unknown', property: prop.siteUrl };
}

/* =====================================================
   GSC — Alert email handler
   ===================================================== */
function handleSendAlert(payload) {
  var domain = payload.domain || 'Unknown domain';
  var issues = payload.issues || [];
  var timestamp = payload.timestamp || new Date().toISOString();
  var accountEmail = payload.accountEmail || '';
  var extraEmail = payload.extraEmail || '';

  if (issues.length === 0) return { success: true, sent: false, reason: 'No issues to report' };

  var settings = getSettings();
  var toRaw = settings['Alert Email TO'] || '';
  var ccRaw = settings['Alert Email CC'] || '';
  var toList = toRaw.split(',').map(function(e) { return e.trim(); }).filter(function(e) { return e.includes('@'); });
  var ccList = ccRaw.split(',').map(function(e) { return e.trim(); }).filter(function(e) { return e.includes('@'); });

  if (extraEmail && extraEmail.includes('@') && toList.indexOf(extraEmail) === -1 && ccList.indexOf(extraEmail) === -1) {
    ccList.push(extraEmail);
  }
  if (toList.length === 0) return { success: false, error: 'No valid TO recipients configured in GSC_Settings' };

  var subject = 'GSC Audit Alert: Issues detected on ' + domain;
  var body = 'GSC Audit Alert\n' + '='.repeat(50) + '\n\nDomain: ' + domain + '\nAudit Time: ' + new Date(timestamp).toLocaleString() + '\n';
  if (accountEmail) body += 'GSC Account: ' + accountEmail + '\n';
  body += '\nIssues Found (' + issues.length + '):\n' + '-'.repeat(40) + '\n';
  issues.forEach(function(issue, i) { body += '\n' + (i + 1) + '. ' + issue.type + '\n   ' + issue.detail + '\n'; });

  var html = '<div style="font-family:Arial,sans-serif;max-width:620px;margin:0 auto">';
  html += '<div style="background:#0D3D4A;padding:22px 28px;border-radius:6px 6px 0 0">';
  html += '<h2 style="color:#C9A84C;margin:0;font-size:20px">GSC Audit Alert</h2>';
  html += '<p style="color:#fff;margin:6px 0 0;font-size:13px">Issues detected during automated audit</p></div>';
  html += '<div style="background:#fff;padding:24px 28px;border:1px solid #dee2e6">';
  html += '<table style="width:100%;border-collapse:collapse;margin-bottom:20px">';
  html += '<tr style="background:#f8f9fa"><td style="padding:10px 14px;font-weight:bold;color:#555;width:130px;border:1px solid #dee2e6">Domain</td>';
  html += '<td style="padding:10px 14px;font-size:17px;font-weight:bold;color:#0D3D4A;border:1px solid #dee2e6">' + domain + '</td></tr>';
  if (accountEmail) {
    html += '<tr><td style="padding:10px 14px;font-weight:bold;color:#555;border:1px solid #dee2e6">GSC Account</td>';
    html += '<td style="padding:10px 14px;color:#333;border:1px solid #dee2e6;font-family:monospace">' + accountEmail + '</td></tr>';
  }
  html += '<tr style="background:#f8f9fa"><td style="padding:10px 14px;font-weight:bold;color:#555;border:1px solid #dee2e6">Issues Found</td>';
  html += '<td style="padding:10px 14px;color:#c0392b;font-weight:bold;border:1px solid #dee2e6">' + issues.length + ' issue(s)</td></tr></table>';
  html += '<h3 style="color:#0D3D4A;border-bottom:2px solid #C9A84C;padding-bottom:8px;margin-top:24px">Issue Details</h3>';
  issues.forEach(function(issue, i) {
    var borderColor = issue.type.indexOf('Warning') !== -1 ? '#e67e22' : '#e74c3c';
    html += '<div style="background:#fff8f8;border-left:4px solid ' + borderColor + ';padding:14px 16px;margin-bottom:14px;border-radius:0 4px 4px 0">';
    html += '<strong style="color:' + borderColor + ';font-size:14px">' + (i + 1) + '. ' + issue.type + '</strong>';
    html += '<div style="margin-top:8px;color:#555;font-size:13px">' + issue.detail.replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>') + '</div></div>';
  });
  html += '</div></div>';

  try {
    var options = { htmlBody: html };
    if (ccList.length > 0) options.cc = ccList.join(',');
    GmailApp.sendEmail(toList.join(','), subject, body, options);
    return { success: true, sent: true, to: toList, cc: ccList, issueCount: issues.length };
  } catch (err) {
    return { success: false, error: 'GmailApp error: ' + err.message };
  }
}

/* =====================================================
   GSC — Monthly audit log
   ===================================================== */
function getOrCreateLogSheet() {
  var ss = SpreadsheetApp.openById(_getSheetId());
  var monthTag = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy_MM');
  var tabName = 'Log_' + monthTag;
  var logSheet = ss.getSheetByName(tabName);
  if (!logSheet) {
    logSheet = ss.insertSheet(tabName);
    var headers = ['Timestamp (IST)', 'Domain(s)', 'Account Key', 'URL Count', 'Status', 'Processed', 'Error / Notes'];
    logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    logSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1e1e2e').setFontColor('#00ff88');
    logSheet.setFrozenRows(1);
  }
  return logSheet;
}

function logRequest(entry) {
  try {
    var logSheet = getOrCreateLogSheet();
    var timestamp = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd MMM yyyy, HH:mm:ss');
    logSheet.appendRow([
      timestamp, entry.domain || '', entry.accountKey || '',
      entry.urlCount || 0, entry.status || '', entry.processed || 0, entry.notes || ''
    ]);
    return logSheet.getLastRow();
  } catch (err) { return null; }
}

function updateLogProcessed(rowNumber, processedCount) {
  if (!rowNumber) return;
  try { getOrCreateLogSheet().getRange(rowNumber, 6).setValue(processedCount); } catch (err) {}
}

/* =====================================================
   GSC — Crawl history
   ===================================================== */
function handleSaveCrawlBatch(data) {
  var ss = SpreadsheetApp.openById(_getSheetId());
  var sheet = ss.getSheetByName('CrawlHistory');
  if (!sheet) {
    sheet = ss.insertSheet('CrawlHistory');
    sheet.appendRow(['BatchId', 'UserId', 'UserName', 'UrlCount', 'Results', 'CreatedAt']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  var batchId = Utilities.getUuid();
  sheet.appendRow([batchId, data.userId || '', data.userName || '', data.urlCount || 0, JSON.stringify(data.results || []), new Date().toISOString()]);
  return { success: true, batchId: batchId };
}

function handleGetCrawlHistory(params) {
  var ss = SpreadsheetApp.openById(_getSheetId());
  var sheet = ss.getSheetByName('CrawlHistory');
  if (!sheet) return _json({ success: true, batches: [] });
  var data = sheet.getDataRange().getValues();
  var userId = params.userId || '';
  var role = params.role || '';
  var all = [];
  for (var i = data.length - 1; i >= 1; i--) {
    var row = { id: data[i][0], userId: data[i][1], userName: data[i][2],
      urlCount: Number(data[i][3]) || 0, results: String(data[i][4]), createdAt: String(data[i][5]) };
    if (role === 'admin' || row.userId === userId) all.push(row);
    if (all.length >= 5) break;
  }
  return _json({ success: true, batches: all });
}

/* =====================================================
   GSC — Sync domains from GSC
   ===================================================== */
function syncDomainsFromGSC() {
  var ss = SpreadsheetApp.openById(_getSheetId());
  try { syncAccountsRegistry(); } catch (e) {}

  var authSheet = ss.getSheetByName('Auth_Status');
  if (!authSheet) { generateAuthSheet(); authSheet = ss.getSheetByName('Auth_Status'); }
  if (!authSheet) return;

  var authData = authSheet.getRange(2, 1, Math.max(authSheet.getLastRow() - 1, 1), 4).getValues();
  var accounts = [];
  authData.forEach(function(row) {
    var email = row[0] ? row[0].toString().trim() : '';
    var key = row[1] ? row[1].toString().trim() : '';
    var status = row[3] ? row[3].toString().trim() : '';
    if (email && key && status === 'Authorised') accounts.push({ email: email, key: key });
  });
  if (accounts.length === 0) return;

  var config = ss.getSheetByName('GSC_Config');
  if (!config) {
    config = ss.insertSheet('GSC_Config');
    config.getRange(1, 1, 1, 6).setValues([['Domain', 'GSC Email', 'Account Key', 'Access Level', 'Last Synced', 'GSC Property']]);
    config.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#1e1e2e').setFontColor('#00ff88');
  }

  var runDateShort = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd MMM yyyy');
  var existingData = config.getDataRange().getValues();
  var byProp = {};
  for (var i = 1; i < existingData.length; i++) {
    var key = existingData[i][2] ? existingData[i][2].toString().trim() : '';
    var prop = existingData[i][5] ? existingData[i][5].toString().trim() : '';
    if (key && prop) byProp[prop.toLowerCase() + '||' + key] = { row: i + 1 };
  }

  var bareHost = function(s) {
    return s.toLowerCase().replace(/^sc-domain:/, '').replace(/^https?:\/\//, '').replace(/^www\./, '').replace(/\/.*$/, '');
  };
  var labelFor = function(permissionLevel, isDomainProperty) {
    var nice = ({ siteOwner: 'Owner', siteFullUser: 'Full', siteRestrictedUser: 'Restricted' })[permissionLevel] || permissionLevel;
    return nice + (isDomainProperty ? ' (domain)' : ' (URL prefix)');
  };

  var newRows = [];
  var seenRows = {};

  for (var a = 0; a < accounts.length; a++) {
    var account = accounts[a];
    var service = getOAuthService(account.key);
    if (!service.hasAccess()) continue;

    var response;
    try {
      response = UrlFetchApp.fetch('https://www.googleapis.com/webmasters/v3/sites', {
        headers: { 'Authorization': 'Bearer ' + service.getAccessToken() }, muteHttpExceptions: true
      });
    } catch (e) { continue; }
    if (response.getResponseCode() !== 200) continue;

    var entries = (JSON.parse(response.getContentText()).siteEntry) || [];
    entries.forEach(function(entry) {
      var siteUrl = entry.siteUrl;
      var domain = bareHost(siteUrl);
      var isDomainProperty = siteUrl.indexOf('sc-domain:') === 0;
      if (entry.permissionLevel === 'siteUnverifiedUser') return;
      var accessLevel = labelFor(entry.permissionLevel, isDomainProperty);

      var rec = byProp[siteUrl.toLowerCase() + '||' + account.key];
      if (rec) {
        seenRows[rec.row] = true;
        config.getRange(rec.row, 4).setValue(accessLevel);
        config.getRange(rec.row, 5).setValue(runDateShort);
      } else {
        newRows.push([domain, account.email, account.key, accessLevel, runDateShort, siteUrl]);
      }
    });
    Utilities.sleep(100);
  }

  if (newRows.length > 0) {
    config.getRange(config.getLastRow() + 1, 1, newRows.length, 6).setValues(newRows);
  }
}

/* =====================================================
   GSC — Sheet menu
   ===================================================== */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SEO Toolkit')
    .addItem('Sync domains from GSC', 'syncDomainsFromGSC')
    .addItem('Generate auth URLs', 'generateAuthSheet')
    .addItem('Check auth status', 'recheckAuthStatusMenu')
    .addItem('Sync accounts registry', 'syncAccountsRegistryMenu')
    .addSeparator()
    .addItem('Open Settings tab', 'openSettingsTab')
    .addToUi();
}

function recheckAuthStatusMenu() {
  var r = recheckAuthStatus();
  SpreadsheetApp.getUi().alert('Authorised: ' + r.done + ' | Pending: ' + r.pending);
}

function syncAccountsRegistryMenu() {
  var added = syncAccountsRegistry();
  SpreadsheetApp.getUi().alert(added + ' new account(s) added to GSC_Accounts.');
}

function openSettingsTab() {
  var ss = SpreadsheetApp.openById(_getSheetId());
  var sheet = ss.getSheetByName('GSC_Settings');
  if (!sheet) sheet = createSettingsTab();
  sheet.activate();
}

/* =====================================================
   UTILITY
   ===================================================== */
function _getSheet(name) {
  var ss = SpreadsheetApp.openById(_getSheetId());
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === 'users') {
      sheet.appendRow(['Email', 'Password', 'Device ID', 'Approved', 'Name', 'Registered', 'Last Login', 'Notes', 'Admin']);
    } else if (name === 'config') {
      sheet.appendRow(['min_version', '3.2']);
    }
  }
  return sheet;
}

function _compareVersions(a, b) {
  var pa = a.toString().split('.').map(Number);
  var pb = b.toString().split('.').map(Number);
  for (var i = 0; i < Math.max(pa.length, pb.length); i++) {
    var na = pa[i] || 0;
    var nb = pb[i] || 0;
    if (na > nb) return 1;
    if (na < nb) return -1;
  }
  return 0;
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
