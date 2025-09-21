/**** ===== UTILS ===== ****/
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