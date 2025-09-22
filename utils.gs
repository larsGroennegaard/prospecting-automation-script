function cfg_(key) {
  // First, check the secure, private storage for the current user.
  const userProperties = PropertiesService.getUserProperties();
  const secretValue = userProperties.getProperty(key);
  if (secretValue) {
    return secretValue;
  }

  // If not found in secure storage, fall back to the 'Config' sheet.
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG_SHEET);
  if (!sh) throw new Error(`Missing sheet: ${CFG_SHEET}`);
  const last = sh.getLastRow();
  if (last < 2) throw new Error(`Config sheet needs key/value rows`);
  const rows = sh.getRange(2,1,last-1,2).getValues();
  const map = new Map(rows.filter(r => r[0]).map(r => [String(r[0]).trim(), String(r[1]||'').trim()]));
  
  const val = map.get(key);
  // Do not throw an error if the key isn't found in the sheet, 
  // as it might be an optional key intended for secure storage only.
  if (val) {
    return val;
  }
  
  // If the key is not in UserProperties and not in the sheet, it's truly missing.
  throw new Error(`Missing Config value for key: ${key}. Please set it in the 'Config' sheet or via the Admin menu.`);
}

function getContentLibrary_() {
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName('Content Library');
    if (!sh || sh.getLastRow() < 2) {
      console.log('Content Library sheet is missing or empty. Skipping.');
      return 'No content available.';
    }
    
    const values = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
    const library = values.map(row => ({
      url: row[0],
      title: row[1],
      description: row[2],
      target_persona: row[3]
    }));
    
    return JSON.stringify(library, null, 2); // Return as a formatted JSON string
  } catch (e) {
    console.error(`Error reading Content Library: ${e.message}`);
    return 'Error reading content library.';
  }
}

function getCfgList_(key) {
  try {
    const v = cfg_(key);
    if (!v) return [];
    return String(v).split(',').map(s => s.trim()).filter(Boolean);
  } catch (_) { return []; }
}

function httpJson_(url, options) {
  const resp = UrlFetchApp.fetch(url, Object.assign({ muteHttpExceptions:true }, options||{}));
  const code = resp.getResponseCode();
  const body = resp.getContentText();
  if (code < 200 || code >= 300) throw new Error(`HTTP ${code}: ${body}`);
  return JSON.parse(body);
}

function ensureAccountsHeader_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(ACC_SHEET);
  if (!sh) throw new Error(`Missing sheet: ${ACC_SHEET}`);

  // New header with the account story column
  const expected = ['selected', 'hubspot_company_id', 'company_name', 'domain', 'signals_last_7_days', 'signals_last_30_days', 'hubspot_owner_email', 'account_story_30_days'];

  if (sh.getLastRow() === 0) {
    sh.appendRow(expected);
  } else {
    const headers = sh.getRange(1, 1, 1, expected.length).getValues()[0].map(String);
    const ok = expected.every((h, i) => (headers[i] || '').toLowerCase() === h);
    if (!ok) throw new Error(`First row of ${ACC_SHEET} must be: ${expected.join(' | ')}`);
  }
  return sh;
}

function ensureContactsHeader_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CON_SHEET);
  if (!sh) throw new Error(`Missing sheet: ${CON_SHEET}`);
  
  // Final header with 20 columns, including assignment columns
  const expected = [
    'selected', 'company_domain', 'contact_name', 'title', 'stage', 'email', 
    'apollo_contact_id', 'hubspot_contact_id', 'contact_story_30_days', 'gem_subject', 'gem_body', 'status', 
    'email_1_subject', 'email_1_body', 'email_2_subject', 'email_2_body', 
    'email_3_subject', 'email_3_body', 'assigned_sending_email', 'assigned_sender_name'
  ];

  if (sh.getLastRow() === 0) {
    sh.appendRow(expected);
    sh.getRange("A2:A").insertCheckboxes(); 
  } else {
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    if (headers.length < expected.length) {
       sh.getRange(1, 1, 1, expected.length).setValues([expected]);
    }
    const ok = expected.every((h, i) => (headers[i] || '').toLowerCase() === h.toLowerCase());
    if (!ok) throw new Error(`First row of ${CON_SHEET} must be: ${expected.join(' | ')}`);
  }
  return sh;
}