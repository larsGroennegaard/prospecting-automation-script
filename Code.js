/**** ===== SHEET NAMES ===== ****/
const CFG_SHEET = 'Config';
const ACC_SHEET = 'Accounts';
const CON_SHEET = 'Contacts';
const MAILBOX_MAP_SHEET = 'Mailbox Mapping';
const CONTENT_LIB_SHEET = 'Content Library'; // Add this line


/**** ===== MENU ===== ****/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // First, create the submenu for Admin tasks
  const adminMenu = ui.createMenu('Admin')
      .addItem('Set API Keys', 'setupApiKeys');

  // Then, create the main menu and add all items, including the submenu
  ui.createMenu('Prospecting')
    .addItem('1) Fetch HS Companies (marked)', 'hsFetchMarkedCompaniesToSheet')
    .addItem('2) Find Contacts in Apollo', 'apolloFindContactsForAccounts')
    .addSeparator()
    .addItem('Enrich: Get Account Stories from BQ', 'enrichFromDataWarehouse')
    .addItem('Enrich: Get Contact Journeys from BQ', 'enrichContactsFromBigQuery')
    .addSeparator()
    .addItem('3) Generate AI Messages', 'generateAiMessages')
    .addItem('4) Push AI Messages to Apollo', 'apolloPushMessages')
    .addItem('5) Add Contacts to Apollo Sequence', 'apolloAddContactsToSequence')
    .addSeparator()
    .addSubMenu(adminMenu) // Add the Admin menu as a submenu
    .addToUi();
}

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

/**
 * Reads the 'Content Library' sheet and formats it as a JSON string for the AI prompt.
 * @returns {string} A JSON string representing the content library, or an empty string.
 */
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
  
  // Final header with 18 columns, including 'selected'
  const expected = [
    'selected', 'company_domain', 'contact_name', 'title', 'stage', 'email', 
    'apollo_contact_id', 'hubspot_contact_id', 'contact_story_30_days', 'gem_subject', 'gem_body', 'status', 
    'email_1_subject', 'email_1_body', 'email_2_subject', 'email_2_body', 
    'email_3_subject', 'email_3_body'
  ];

  if (sh.getLastRow() === 0) {
    sh.appendRow(expected);
    // Automatically insert checkboxes for any new rows in the 'selected' column
    sh.getRange("A2:A").insertCheckboxes(); 
  } else {
    const headers = sh.getRange(1, 1, 1, expected.length).getValues()[0].map(String);
    const ok = expected.every((h, i) => (headers[i] || '').toLowerCase() === h.toLowerCase());
    if (!ok) throw new Error(`First row of ${CON_SHEET} must be: ${expected.join(' | ')}`);
  }
  return sh;
}

function apolloGet_(url) {
  const apiKey = cfg_('APOLLO_API_KEY');
  const headers = { 'x-api-key': apiKey, 'Content-Type': 'application/json' };
  const options = { method: 'get', headers: headers, muteHttpExceptions: true };
  const resp = UrlFetchApp.fetch(url, options);
  const code = resp.getResponseCode();
  const txt = resp.getContentText();
  console.log(`GET ${url} → HTTP ${code}`);
  if (code < 200 || code >= 300) throw new Error(`Apollo API Error ${code}: ${txt}`);
  return JSON.parse(txt);
}

function apolloPost_(url, payload) {
  const apiKey = cfg_('APOLLO_API_KEY');
  const headers = { 'x-api-key': apiKey, 'Content-Type': 'application/json' };
  const options = { method: 'post', headers: headers, payload: JSON.stringify(payload), muteHttpExceptions: true };
  const resp = UrlFetchApp.fetch(url, options);
  const code = resp.getResponseCode();
  const txt = resp.getContentText();
  console.log(`POST ${url} → HTTP ${code}, Body: ${txt.slice(0, 500)}`);
  if (code < 200 || code >= 300) throw new Error(`Apollo API Error ${code}: ${txt}`);
  return JSON.parse(txt);
}

function apolloPut_(url, payload) {
  const apiKey = cfg_('APOLLO_API_KEY');
  const headers = { 'x-api-key': apiKey, 'Content-Type': 'application/json' };
  const options = { method: 'put', headers: headers, payload: JSON.stringify(payload), muteHttpExceptions: true };
  const resp = UrlFetchApp.fetch(url, options);
  const code = resp.getResponseCode();
  const txt = resp.getContentText();
  console.log(`PUT ${url} → HTTP ${code}, Body: ${txt.slice(0, 500)}`);
  if (code < 200 || code >= 300) throw new Error(`Apollo API Error ${code}: ${txt}`);
  return JSON.parse(txt);
}

function getApolloStageMap_() {
  try {
    const url = 'https://api.apollo.io/v1/contact_stages';
    const apiKey = cfg_('APOLLO_API_KEY');
    // **CORRECTED**: This header requires Content-Type to work.
    const headers = { 'x-api-key': apiKey, 'Content-Type': 'application/json' };
    const data = httpJson_(url, { method: 'get', headers: headers });
    const stageMap = new Map();
    if (data.contact_stages && Array.isArray(data.contact_stages)) {
      data.contact_stages.forEach(stage => {
        if (stage.id && stage.name) stageMap.set(stage.id, stage.name);
      });
    }
    console.log(`Successfully fetched and mapped ${stageMap.size} Apollo contact stages.`);
    return stageMap;
  } catch (e) {
    console.error(`Could not fetch Apollo contact stages: ${e.message}`);
    return new Map();
  }
}

function getApolloMailboxMap_() {
  try {
    const url = 'https://api.apollo.io/v1/email_accounts';
    const data = apolloGet_(url);
    const mailboxMap = new Map();
    if (data.email_accounts && Array.isArray(data.email_accounts)) {
      data.email_accounts.forEach(acc => {
        if (acc.id && acc.email) mailboxMap.set(acc.email.toLowerCase(), acc.id);
      });
    }
    console.log(`Successfully fetched ${mailboxMap.size} Apollo mailboxes.`);
    return mailboxMap;
  } catch (e) {
    console.error(`Could not fetch Apollo mailboxes: ${e.message}`);
    return new Map();
  }
}

function apolloEnrichPerson_(personId) {
  const url = 'https://api.apollo.io/v1/people/enrich';
  const payload = { id: personId };
  console.log(`Enriching person ID: ${personId}`);
  const apiKey = cfg_('APOLLO_API_KEY');
  const headers = { 'x-api-key': apiKey, 'Content-Type': 'application/json' };
  return httpJson_(url, { method: 'post', headers: headers, payload: JSON.stringify(payload) });
}

function apolloSaveAsContact_(personId) {
  const url = 'https://api.apollo.io/v1/contacts';
  const payload = { person_ids: [personId] };
  console.log(`Saving person ID as contact: ${personId}`);
  return apolloPost_(url, payload);
}

/**** ===== 1) HUBSPOT: FETCH MARKED COMPANIES ===== ****/
function hsFetchMarkedCompaniesToSheet() {
  const token = cfg_('HUBSPOT_TOKEN');
  const prop  = cfg_('HUBSPOT_COMPANY_PROP');
  const acc   = ensureAccountsHeader_();

  const last = acc.getLastRow();
  const existing = new Map();
  if (last >= 2) {
    const vals = acc.getRange(2, 1, last - 1, 4).getValues();
    vals.forEach((r, i) => {
      const id = String(r[1] || '').trim();
      if (id) existing.set(id, i + 2);
    });
  }

  const headers = { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' };
  const url = 'https://api.hubapi.com/crm/v3/objects/companies/search';
  const baseBody = {
    filterGroups: [{ filters: [{ propertyName: prop, operator: 'EQ', value: 'true' }] }],
    properties: ['name', 'domain', 'signals_last_7_days', 'signals_last_30_days', 'hubspot_owner_id'],
    limit: 100
  };

  let after;
  const allCompanies = [];
  while (true) {
    const body = Object.assign({}, baseBody, after ? { after } : {});
    const res = httpJson_(url, { method: 'post', payload: JSON.stringify(body), headers });
    if (res.results && res.results.length > 0) allCompanies.push(...res.results);
    if (res.paging && res.paging.next && res.paging.next.after) after = res.paging.next.after;
    else break;
  }
  console.log(`Fetched a total of ${allCompanies.length} companies.`);

  const ownerIds = [...new Set(allCompanies.map(c => c.properties.hubspot_owner_id).filter(Boolean))];
  const ownerMap = new Map();
  if (ownerIds.length > 0) {
    console.log(`Found ${ownerIds.length} unique owners. Fetching their emails...`);
    for (const ownerId of ownerIds) {
      try {
        const ownerUrl = `https://api.hubapi.com/crm/v3/owners/${ownerId}`;
        const ownerRes = httpJson_(ownerUrl, { method: 'get', headers: headers });
        if (ownerRes && ownerRes.email) ownerMap.set(ownerRes.id, ownerRes.email);
      } catch (e) {
        console.error(`Could not fetch details for owner ID ${ownerId}: ${e.message}`);
      }
    }
  }

  let syncedCount = 0;
  allCompanies.forEach(c => {
    const props = c.properties || {};
    const id = c.id;
    const name = props.name || '';
    const domain = props.domain || '';
    const signals7 = props.signals_last_7_days || 0;
    const signals30 = props.signals_last_30_days || 0;
    const ownerId = props.hubspot_owner_id || '';
    const ownerEmail = ownerMap.get(ownerId) || '';
    const rowData = [true, id, name, domain, signals7, signals30, ownerEmail];
    if (existing.has(id)) {
      const row = existing.get(id);
      acc.getRange(row, 1, 1, 7).setValues([rowData]);
    } else {
      acc.appendRow(rowData);
    }
    syncedCount++;
  });
  SpreadsheetApp.getActive().toast(`HubSpot companies synced: ${syncedCount}`);
}

/**** ===== 2) APOLLO: FIND CONTACTS ===== ****/
// Replace this entire function
// Replace this entire function
function apolloFindContactsForAccounts() {
  const ui = SpreadsheetApp.getUi();

  const accSh = ensureAccountsHeader_();
  const conSh = ensureContactsHeader_();

  // --- FIX #1: Use the correct endpoints ---
  const SEARCH_URL = 'https://api.apollo.io/v1/mixed_people/search';
  const ORGS_SEARCH_URL = 'https://api.apollo.io/v1/accounts/search'; // Changed from /organizations/search
  
  const stageMap = getApolloStageMap_();
  
  const allowedStages = getCfgList_('APOLLO_ALLOWED_STAGES').map(s => s.toLowerCase());
  allowedStages.push('');

  const accRows = accSh.getRange(2, 1, Math.max(0, accSh.getLastRow() - 1), 4).getValues();
  const accounts = accRows
    .filter(r => r[0] === true)
    .map(r => ({ domain: String(r[3] || '').toLowerCase().trim(), name: String(r[2] || '').trim() }))
    .filter(a => a.domain);

  if (!accounts.length) {
    ui.alert('No selected accounts with domains.');
    return;
  }

  const existing = new Set();
  if (conSh.getLastRow() >= 2) {
    const vals = conSh.getRange(2, 1, conSh.getLastRow() - 1, 7).getValues();
    vals.forEach(r => {
      const domain = String(r[1] || '').toLowerCase();
      const email = String(r[5] || '').toLowerCase();
      const apolloId = String(r[6] || '').toLowerCase();
      if (domain && (email || apolloId)) existing.add(`${domain}::${email || apolloId}`);
    });
  }

  const headers = { 'x-api-key': cfg_('APOLLO_API_KEY'), 'Content-Type': 'application/json' };
  
  const fTitles = getCfgList_('APOLLO_TITLES');
  const includeSimilarTitles = (() => {
    try { return String(cfg_('APOLLO_INCLUDE_SIMILAR_TITLES')).toLowerCase() === 'true'; } 
    catch (e) { return false; }
  })();

  function buildSearchPayload(base) {
    const payload = Object.assign({
      page: 1, per_page: 100,
      person_titles: fTitles.length ? fTitles : undefined,
    }, base);
    
    // --- FIX #2: Use the correct parameter name for domain search ---
    if (payload.q_organization_domains) {
      payload.q_organization_domains_list = payload.q_organization_domains;
      delete payload.q_organization_domains;
    }
    
    if (payload.person_titles && includeSimilarTitles) payload.include_similar_titles = true;
    return payload;
  }

  // Helper functions processResults and appendRows remain the same as your current version
  function processResults(results, domain) {
    const rows = [];
    const placeholderEmail = 'email_not_unlocked@domain.com';
    for (const p of results) {
      const apolloId = String(p.id || '').trim();
      let email = String(p.email || '').toLowerCase();
      const name = [p.first_name, p.last_name].filter(Boolean).join(' ').trim() || 'N/A';
      const isContact = p.is_apollo_contact;
      const status = isContact ? 'from_apollo_contact' : 'from_apollo_person';
      if (!email && !apolloId) continue;
      if (email === placeholderEmail) email = '';
      const title = p.title || '';
      const stageId = p.contact_stage_id || '';
      const stage = stageId ? (stageMap.get(stageId) || stageId) : '';
      const hubspotId = p.hubspot_vid || '';
      const key = `${domain}::${email || apolloId}`;
      if (existing.has(key)) continue;
      const originalEmail = String(p.email || '').toLowerCase();
      rows.push([true, domain, name, title, stage, originalEmail, apolloId, hubspotId, '', '', '', status, '', '', '', '', '', '']);
      existing.add(key);
    }
    return rows;
  }

  let totalAppended = 0;
  for (const acc of accounts) {
    const { domain, name } = acc;
    console.log(`\n\n=== Starting Apollo search for company: "${name}" (domain: "${domain}") ===`);
    
    let allPotentialContacts = [];
    let searchPayloadBase = { q_organization_domains: [domain] }; // Start with old key, buildSearchPayload will fix it

    let page = 1;
    while (true) {
      try {
        const payload = buildSearchPayload({ ...searchPayloadBase, page });
        const data = httpJson_(SEARCH_URL, { method: 'post', headers, payload: JSON.stringify(payload) });
        
        const contacts = (data.contacts || []).map(c => ({...c, is_apollo_contact: true}));
        const people = (data.people || []).map(p => ({...p, is_apollo_contact: false}));
        const results = contacts.concat(people);
        if (results.length > 0) allPotentialContacts.push(...results);
        
        const pg = data.pagination || {};
        if (pg.page && pg.total_pages && pg.page < pg.total_pages) {
          page++;
        } else {
          if (allPotentialContacts.length === 0 && !searchPayloadBase.organization_ids) {
            console.log(`No results via domain search. Trying to find organization ID...`);
            let org = null;
            try {
              // Note: The org search here doesn't use the domain, but the name, as per your successful curl.
              // We'll prioritize the name search.
              if (name) {
                 const orgsByNameData = httpJson_(ORGS_SEARCH_URL, { method: 'post', headers, payload: JSON.stringify({ q_organization_name: name }) });
                 org = orgsByNameData.accounts && orgsByNameData.accounts[0]; // Look for 'accounts' array
              }
            } catch(e) { console.error(`Error while finding Organization ID: ${e.message}`); }
            
            if (org) {
              console.log(`Found Org match: id=${org.id}, name="${org.name}". Restarting search with Org ID.`);
              searchPayloadBase = { organization_ids: [org.id] };
              page = 1;
              continue;
            }
          }
          break;
        }
      } catch (e) { console.error(e); break; }
    }
    
    // The rest of the function (filtering, sorting, enriching, appending) remains the same.
    console.log(`Found ${allPotentialContacts.length} total potential contacts for ${domain}.`);

    const stageFilteredContacts = allPotentialContacts.filter(p => {
      const stageId = p.contact_stage_id || '';
      const stageName = stageId ? (stageMap.get(stageId) || '') : '';
      return allowedStages.includes(stageName.toLowerCase());
    });
    console.log(`After stage filtering: ${stageFilteredContacts.length} contacts remain.`);

    const placeholderEmail = 'email_not_unlocked@domain.com';
    const sortedContacts = stageFilteredContacts.sort((a, b) => {
      const aHasEmail = a.email && a.email !== placeholderEmail;
      const bHasEmail = b.email && b.email !== placeholderEmail;
      return aHasEmail === bHasEmail ? 0 : aHasEmail ? -1 : 1;
    });

    const rowsToAdd = [];
    for (const p of sortedContacts) {
      if (rowsToAdd.length >= 5) break;
      let finalPerson = p;
      let email = String(p.email || '').toLowerCase();
      let isNowAContact = p.is_apollo_contact;
      if (!email || email === placeholderEmail) {
        try {
          const enrichedData = apolloEnrichPerson_(p.id);
          if (enrichedData.person && enrichedData.person.email) {
            finalPerson = { ...p, ...enrichedData.person };
            try { apolloSaveAsContact_(finalPerson.id); isNowAContact = true; } 
            catch (saveError) { /* ignore */ }
          } else { continue; }
        } catch (e) { continue; }
      }
      
      const finalEmail = String(finalPerson.email || '').toLowerCase();
      if (!finalEmail || finalEmail === placeholderEmail) continue;
      
      const processedRows = processResults([ {...finalPerson, is_apollo_contact: isNowAContact} ], domain);
      if (processedRows.length > 0) rowsToAdd.push(...processedRows);
    }

    if (rowsToAdd.length > 0) {
      conSh.getRange(conSh.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
      totalAppended += rowsToAdd.length;
    }
  }
  ui.alert(`Apollo contact search complete. Appended a total of ${totalAppended} new contacts.`);
}

/**** ===== 3) GEMINI: GENERATE AI MESSAGES ===== ****/
// Replace this entire function
// Replace your entire generateAiMessages function with this
function generateAiMessages() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const conSh = ss.getSheetByName(CON_SHEET);
  if (conSh.getLastRow() < 2) { ui.alert('No contacts to process.'); return; }

  const allContacts = conSh.getRange(2, 1, conSh.getLastRow() - 1, conSh.getLastColumn()).getValues();
  const selectedContacts = allContacts.filter(row => row[0] === true);

  if (selectedContacts.length === 0) {
    ui.alert('No contacts are selected. Please check the boxes in column A for the contacts you want to process.');
    return;
  }
  
  const response = ui.alert('Generate AI Messages?', `This will process the ${selectedContacts.length} selected contacts. Continue?`, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  const accSh = ss.getSheetByName(ACC_SHEET);
  const mapSh = ss.getSheetByName(MAILBOX_MAP_SHEET);
  
  // Get all config values at the start
  const sequencePromptTemplate = cfg_('EMAIL_SEQUENCE_PROMPT');
  const myCompanyName = cfg_('MY_COMPANY_NAME');
  const myValueProp = cfg_('MY_VALUE_PROPOSITION');
  // --- NEW: Get positioning and use case data ---
  const dreamdataPositioning = (() => { try { return cfg_('DREAMDATA_POSITIONING'); } catch(e) { return 'A B2B Attribution Platform.'; } })();
  const dreamdataUseCases = (() => { try { return cfg_('DREAMDATA_USE_CASES'); } catch(e) { return 'No use cases provided.'; } })();
  const contentLibraryJson = getContentLibrary_();


  const ownerDetailsMap = new Map();
  if (mapSh && mapSh.getLastRow() > 1) {
    const mappingValues = mapSh.getRange(2, 1, mapSh.getLastRow() - 1, 3).getValues();
    mappingValues.forEach(row => {
      const ownerEmail = String(row[0]).toLowerCase().trim();
      const senderName = String(row[2]).trim();
      if (ownerEmail && senderName) ownerDetailsMap.set(ownerEmail, { senderName: senderName });
    });
  }

  const companyData = new Map();
  if (accSh.getLastRow() > 1) {
    const accValues = accSh.getRange(2, 1, accSh.getLastRow() - 1, 8).getValues();
    accValues.forEach(row => {
      const domain = String(row[3]).toLowerCase().trim();
      if (domain) {
        companyData.set(domain, {
          company_name: row[2],
          signals_last_7_days: row[4],
          signals_last_30_days: row[5],
          owner_email: String(row[6]).toLowerCase().trim(),
          account_story: row[7] || 'No specific account journey data available.',
        });
      }
    });
  }

  let processedCount = 0;
  allContacts.forEach((row, index) => {
    const isSelected = row[0];
    const subjectCell = row[12];
    
    if (isSelected && subjectCell === '') {
      const placeholders = {
        '{contact_name}': row[2],
        '{title}': row[3],
        '{stage}': row[4],
        '{company_domain}': row[1],
        '{contact_story_30_days}': row[8] || 'No specific contact journey data available.',
        '{my_company_name}': myCompanyName,
        '{my_value_proposition}': myValueProp,
        '{email_sender}': '',
        '{content_library}': contentLibraryJson,
        // --- NEW: Add the new placeholders ---
        '{dreamdata_positioning}': dreamdataPositioning,
        '{dreamdata_use_cases}': dreamdataUseCases,
      };

      const companyInfo = companyData.get(String(row[1]).toLowerCase().trim()) || {};
      placeholders['{company_name}'] = companyInfo.company_name || row[1];
      placeholders['{signals_last_7_days}'] = companyInfo.signals_last_7_days || 0;
      placeholders['{signals_last_30_days}'] = companyInfo.signals_last_30_days || 0;
      placeholders['{account_story_30_days}'] = companyInfo.account_story || 'No specific account journey data available.';
      
      if (companyInfo.owner_email) {
        const ownerDetails = ownerDetailsMap.get(companyInfo.owner_email);
        if (ownerDetails && ownerDetails.senderName) placeholders['{email_sender}'] = ownerDetails.senderName;
      }

      let finalPrompt = sequencePromptTemplate;
      for (const key in placeholders) {
        finalPrompt = finalPrompt.replace(new RegExp(key, 'g'), placeholders[key]);
      }
      
      console.log(`--> Final prompt sent to AI:\n${finalPrompt}`);
      const jsonResponseString = geminiGenerate_(finalPrompt);
      
      let outputs = ['Error', 'Error', 'Error', 'Error', 'Error', 'Error'];
      try {
        const jsonMatch = jsonResponseString.match(/\[[\s\S]*\]|{[\s\S]*}/);
        if (!jsonMatch) throw new Error("No JSON found");
        const sequenceArray = JSON.parse(jsonMatch[0]);
        if (Array.isArray(sequenceArray) && sequenceArray.length === 3 && sequenceArray[0].subject) {
          outputs = [
            sequenceArray[0].subject, sequenceArray[0].body,
            sequenceArray[1].subject, sequenceArray[1].body,
            sequenceArray[2].subject, sequenceArray[2].body,
          ];
        } else { outputs[0] = 'Error: Invalid JSON structure'; outputs[1] = jsonResponseString; }
      } catch (e) { outputs[0] = 'Error: Could not parse AI response'; outputs[1] = jsonResponseString; }
      
      conSh.getRange(index + 2, 13, 1, 6).setValues([outputs]);
      processedCount++;
    }
  });
  ui.alert(`AI sequence generation complete. Processed ${processedCount} selected contacts.`);
}

function geminiGenerate_(prompt) {
  const apiKey = cfg_('GEMINI_API_KEY');
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${apiKey}`;

  const payload = {
    contents: [{
      parts: [{ text: prompt }]
    }],
    generationConfig: {
      temperature: 0.7,
      topK: 1,
      topP: 1,
      maxOutputTokens: 8192,
      // **THE FIX**: Enforce JSON output from the API
      response_mime_type: "application/json",
    },
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseBody = response.getContentText();
    const json = JSON.parse(responseBody);
    
    if (json.candidates && json.candidates[0].content.parts[0].text) {
      return json.candidates[0].content.parts[0].text.trim();
    } else {
      console.error('Gemini API Error: Invalid response structure.', responseBody);
      return `Error: ${json.error ? json.error.message : 'Invalid response'}`;
    }
  } catch (e) {
    console.error('Gemini API call failed.', e);
    return `Error: ${e.message}`;
  }
}

/**** ===== 4) APOLLO: PUSH AI MESSAGES ===== ****/
function apolloPushMessages() {
  const ui = SpreadsheetApp.getUi();
  const conSh = SpreadsheetApp.getActive().getSheetByName(CON_SHEET);
  if (conSh.getLastRow() < 2) { ui.alert('No contacts to process.'); return; }

  const allContacts = conSh.getRange(2, 1, conSh.getLastRow() - 1, conSh.getLastColumn()).getValues();
  const selectedContacts = allContacts.filter(row => row[0] === true);

  if (selectedContacts.length === 0) {
    ui.alert('No contacts are selected.');
    return;
  }

  const response = ui.alert('Push AI Messages to Apollo?', `This will update ${selectedContacts.length} selected contacts in Apollo. Continue?`, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  const fieldIds = {
    subject1: cfg_('APOLLO_FIELD_ID_EMAIL_1_SUBJECT'), body1: cfg_('APOLLO_FIELD_ID_EMAIL_1_BODY'),
    subject2: cfg_('APOLLO_FIELD_ID_EMAIL_2_SUBJECT'), body2: cfg_('APOLLO_FIELD_ID_EMAIL_2_BODY'),
    subject3: cfg_('APOLLO_FIELD_ID_EMAIL_3_SUBJECT'), body3: cfg_('APOLLO_FIELD_ID_EMAIL_3_BODY'),
  };

  let processedCount = 0, errorCount = 0, skippedCount = 0;
  allContacts.forEach((row, index) => {
    const isSelected = row[0];       // col A
    const apolloContactId = row[6];  // col G
    const status = row[11];          // col L
    const email1Subject = row[12];   // col M

    if (isSelected && apolloContactId && email1Subject && !status.includes('apollo_pushed') && status.includes('from_apollo_contact')) {
      const payload = { typed_custom_fields: {} };
      payload.typed_custom_fields[fieldIds.subject1] = row[12];
      payload.typed_custom_fields[fieldIds.body1] = row[13];
      payload.typed_custom_fields[fieldIds.subject2] = row[14];
      payload.typed_custom_fields[fieldIds.body2] = row[15];
      payload.typed_custom_fields[fieldIds.subject3] = row[16];
      payload.typed_custom_fields[fieldIds.body3] = row[17];
      try {
        const url = `https://api.apollo.io/v1/contacts/${apolloContactId}`;
        apolloPut_(url, payload);
        const newStatus = status ? `${status};apollo_pushed` : 'apollo_pushed';
        conSh.getRange(index + 2, 12).setValue(newStatus);
        processedCount++;
      } catch (e) {
        const newStatus = status ? `${status};apollo_push_failed` : 'apollo_push_failed';
        conSh.getRange(index + 2, 12).setValue(newStatus);
        errorCount++;
      }
      SpreadsheetApp.flush();
    } else if (isSelected && apolloContactId && email1Subject && !status.includes('apollo_pushed')) {
      skippedCount++;
      const newStatus = status ? `${status};apollo_push_skipped(person)` : 'apollo_push_skipped(person)';
      conSh.getRange(index + 2, 12).setValue(newStatus);
    }
  });
  ui.alert(`Push to Apollo complete.\n\nUpdated: ${processedCount}\nFailed: ${errorCount}\nSkipped: ${skippedCount}`);
}

/**** ===== 5) APOLLO: ADD CONTACTS TO SEQUENCE ===== ****/
function apolloAddContactsToSequence() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accSh = ss.getSheetByName(ACC_SHEET);
  const conSh = ss.getSheetByName(CON_SHEET);
  const mapSh = ss.getSheetByName(MAILBOX_MAP_SHEET);

  if (conSh.getLastRow() < 2) { ui.alert('No contacts to process.'); return; }
  
  const allContacts = conSh.getRange(2, 1, conSh.getLastRow() - 1, conSh.getLastColumn()).getValues();
  const selectedContacts = allContacts.filter(row => row[0] === true);

  if (selectedContacts.length === 0) {
    ui.alert('No contacts are selected.');
    return;
  }
  
  const response = ui.alert('Add Contacts to Apollo Sequence?', `This will add ${selectedContacts.length} selected contacts to the sequence. Continue?`, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  if (!mapSh || mapSh.getLastRow() < 2) { /* ... */ } // This part is correct
  const apolloMailboxMap = getApolloMailboxMap_();
  const sequenceId = cfg_('APOLLO_SEQUENCE_ID');
  if (apolloMailboxMap.size === 0) { /* ... */ } // This part is correct

  const ownerToMailboxMap = new Map();
  const mappingValues = mapSh.getRange(2, 1, mapSh.getLastRow() - 1, 2).getValues();
  mappingValues.forEach(row => {
    const ownerEmail = String(row[0]).toLowerCase().trim();
    const sendingEmail = String(row[1]).toLowerCase().trim();
    if (ownerEmail && sendingEmail) ownerToMailboxMap.set(ownerEmail, sendingEmail);
  });
  
  const companyToOwnerMap = new Map();
  if (accSh.getLastRow() > 1) { /* ... */ } // This part is correct

  const contactsBySender = new Map();
  allContacts.forEach((row, index) => {
    const isSelected = row[0];       // col A
    const status = row[11];          // col L
    const apolloContactId = row[6];  // col G
    const domain = String(row[1]).toLowerCase().trim();

    if (isSelected && apolloContactId && status.includes('from_apollo_contact') && !status.includes('apollo_sequenced')) {
      const ownerEmail = companyToOwnerMap.get(domain);
      if (ownerEmail) {
        const sendingEmail = ownerToMailboxMap.get(ownerEmail);
        if (sendingEmail) {
          if (!contactsBySender.has(sendingEmail)) contactsBySender.set(sendingEmail, []);
          contactsBySender.get(sendingEmail).push({ id: apolloContactId, rowIndex: index + 2 });
        } else {
          const currentStatus = conSh.getRange(index + 2, 12).getValue();
          conSh.getRange(index + 2, 12).setValue(`${currentStatus};sequence_failed(no_mapping)`);
        }
      }
    }
  });

  let totalSuccess = 0, totalFailed = 0;
  for (const [sendingEmail, contacts] of contactsBySender.entries()) {
    const mailboxId = apolloMailboxMap.get(sendingEmail);
    if (!mailboxId) {
      contacts.forEach(c => {
        const currentStatus = conSh.getRange(c.rowIndex, 12).getValue();
        conSh.getRange(c.rowIndex, 12).setValue(`${currentStatus};sequence_failed(no_mailbox)`);
      });
      continue;
    }
    
    const contactIds = contacts.map(c => c.id);
    const url = `https://api.apollo.io/v1/emailer_campaigns/${sequenceId}/add_contact_ids`;
    const payload = { contact_ids: contactIds, send_email_from_email_account_id: mailboxId, emailer_campaign_id: sequenceId };
    try {
      apolloPost_(url, payload);
      contacts.forEach(c => {
        const currentStatus = conSh.getRange(c.rowIndex, 12).getValue();
        conSh.getRange(c.rowIndex, 12).setValue(`${currentStatus};apollo_sequenced`);
      });
      totalSuccess += contacts.length;
    } catch (e) {
      contacts.forEach(c => {
        const currentStatus = conSh.getRange(c.rowIndex, 12).getValue();
        conSh.getRange(c.rowIndex, 12).setValue(`${currentStatus};sequence_failed(api_error)`);
      });
    }
    SpreadsheetApp.flush();
  }
  ui.alert(`Sequence enrollment complete.\n\nAdded: ${totalSuccess}\nFailed/Skipped: ${totalFailed}`);
}

/**** ===== 6) BIGQUERY: ENRICH ACCOUNTS ===== ****/

/**
 * For selected accounts, runs a BigQuery query to get a 30-day event story.
 */
function enrichFromDataWarehouse() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive(); // Use ss for toast messages
  const response = ui.alert(
    'Enrich Accounts from BigQuery?',
    'This will run a query for each selected account in the "Accounts" sheet. This will incur BigQuery costs. Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const accSh = ss.getSheetByName(ACC_SHEET);
  if (accSh.getLastRow() < 2) {
    ui.alert('No accounts to process.');
    return;
  }

  const projectId = cfg_('GCP_PROJECT_ID');
  const queryTemplate = cfg_('BQ_ACCOUNT_STORY_QUERY');

  const headers = accSh.getRange(1, 1, 1, accSh.getLastColumn()).getValues()[0];
  const storyColumnIndex = headers.map(h => h.toLowerCase()).indexOf('account_story_30_days');
  if (storyColumnIndex === -1) {
    ui.alert('Error: Could not find the "account_story_30_days" column.');
    return;
  }

  const accRange = accSh.getRange(2, 1, accSh.getLastRow() - 1, storyColumnIndex + 1);
  const accValues = accRange.getValues();
  let processedCount = 0;

  // --- FIX: Use a toast message instead of a modal dialog ---
  ss.toast('Enriching accounts from BigQuery... Please wait.', 'Processing...', -1);

  accValues.forEach((row, index) => {
    const isSelected = row[0];
    const hubspotId = row[1];

    if (isSelected && hubspotId) {
      console.log(`Processing company ID: ${hubspotId}`);
      const companyIdForQuery = `hubspot-${hubspotId}`;
      
      const job = {
        configuration: {
          query: {
            query: queryTemplate,
            useLegacySql: false,
            queryParameters: [{
              name: 'companyId',
              parameterType: { type: 'STRING' },
              parameterValue: { value: companyIdForQuery }
            }]
          }
        }
      };

      try {
        let queryJob = BigQuery.Jobs.insert(job, projectId);
        const jobId = queryJob.jobReference.jobId;
        
        while (queryJob.status.state !== 'DONE') {
          Utilities.sleep(500);
          queryJob = BigQuery.Jobs.get(projectId, jobId);
        }

        if (queryJob.status.errorResult) {
          throw new Error(`BigQuery job failed: ${JSON.stringify(queryJob.status.errors)}`);
        }

        const results = BigQuery.Jobs.getQueryResults(projectId, jobId);
        let story = 'No events found in the last 30 days.';
        if (results.rows && results.rows.length > 0) {
          story = results.rows[0].f[0].v;
        }

        accSh.getRange(index + 2, storyColumnIndex + 1).setValue(story);
        processedCount++;

      } catch (e) {
        console.error(`Failed to enrich company ${hubspotId}: ${e.message}`);
        accSh.getRange(index + 2, storyColumnIndex + 1).setValue(`Error: ${e.message}`);
      }
    }
  });
  
  // --- FIX: Display a final toast which dismisses the "Processing" one ---
  ss.toast(`Enrichment complete. Processed ${processedCount} accounts.`, 'Success!', 5);
}

/**
 * For selected contacts, runs a BigQuery query to get a 30-day event journey.
 */
function enrichContactsFromBigQuery() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive(); // Use ss for toast messages
  const response = ui.alert(
    'Enrich Contacts from BigQuery?',
    'This will run a query for each selected contact in the "Contacts" sheet. This will incur BigQuery costs. Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const conSh = ss.getSheetByName(CON_SHEET);
  if (conSh.getLastRow() < 2) {
    ui.alert('No contacts to process.');
    return;
  }

  const projectId = cfg_('GCP_PROJECT_ID');
  const queryTemplate = cfg_('BQ_CONTACT_STORY_QUERY');

  const headers = conSh.getRange(1, 1, 1, conSh.getLastColumn()).getValues()[0];
  const hsIdColIdx = headers.map(h => h.toLowerCase()).indexOf('hubspot_contact_id');
  const storyColIdx = headers.map(h => h.toLowerCase()).indexOf('contact_story_30_days');
  const selectedColIdx = headers.map(h => h.toLowerCase()).indexOf('selected');

  if (hsIdColIdx === -1 || storyColIdx === -1 || selectedColIdx === -1) {
    ui.alert('Error: Could not find one of the required columns: selected, hubspot_contact_id, contact_story_30_days.');
    return;
  }
  
  const conRange = conSh.getRange(2, 1, conSh.getLastRow() - 1, conSh.getLastColumn());
  const conValues = conRange.getValues();
  let processedCount = 0;

  // --- FIX: Use a toast message instead of a modal dialog ---
  ss.toast('Enriching contacts from BigQuery... Please wait.', 'Processing...', -1);

  conValues.forEach((row, index) => {
    const isSelected = row[selectedColIdx];
    const hubspotId = row[hsIdColIdx];

    if (isSelected && hubspotId) {
      console.log(`Processing HubSpot Contact ID: ${hubspotId}`);
      
      const job = {
        configuration: {
          query: {
            query: queryTemplate,
            useLegacySql: false,
            queryParameters: [{
              name: 'hubspotContactId',
              parameterType: { type: 'STRING' },
              parameterValue: { value: hubspotId }
            }]
          }
        }
      };

      try {
        let queryJob = BigQuery.Jobs.insert(job, projectId);
        const jobId = queryJob.jobReference.jobId;
        
        while (queryJob.status.state !== 'DONE') {
          Utilities.sleep(500);
          queryJob = BigQuery.Jobs.get(projectId, jobId);
        }

        if (queryJob.status.errorResult) {
          throw new Error(`BigQuery job failed: ${JSON.stringify(queryJob.status.errors)}`);
        }

        const results = BigQuery.Jobs.getQueryResults(projectId, jobId);
        let story = 'No events found in the last 30 days.';
        if (results.rows && results.rows.length > 0) {
          story = results.rows[0].f[0].v;
        }

        conSh.getRange(index + 2, storyColIdx + 1).setValue(story);
        processedCount++;

      } catch (e) {
        console.error(`Failed to enrich contact ${hubspotId}: ${e.message}`);
        conSh.getRange(index + 2, storyColIdx + 1).setValue(`Error: ${e.message}`);
      }
    }
  });

  // --- FIX: Display a final toast which dismisses the "Processing" one ---
  ss.toast(`Enrichment complete. Processed ${processedCount} contacts.`, 'Success!', 5);
}

/**** ===== SETUP UTILITIES ===== ****/

/**
 * Ensures the Content Library sheet has the correct header.
 */
function ensureContentLibraryHeader_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONTENT_LIB_SHEET);
  if (!sh) throw new Error(`Missing sheet: ${CONTENT_LIB_SHEET}`);
  const expected = ["URL", "Title", "Description / Key Takeaway", "Target Persona"];
  if (sh.getLastRow() > 0) sh.clear(); // Clear the sheet before adding headers
  sh.appendRow(expected);
  SpreadsheetApp.flush();
  return sh;
}

/**
 * Main function to orchestrate the blog cataloging process.
 */
function cl_buildLibrary() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Rebuild Content Library?',
    'This will clear the "Content Library" sheet and rebuild it by scraping and analyzing your blog posts. This can take several minutes. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  const sheet = ensureContentLibraryHeader_();

  // You can add more blog index pages here if needed
  const blogIndexPages = [
    "https://dreamdata.io/blog",
    "https://dreamdata.io/blog?offset=1744801352826"
  ];

  try {
    const allPostUrls = cl_getAllPostUrls_(blogIndexPages);
    Logger.log(`Found ${allPostUrls.length} unique blog post URLs.`);

    if (allPostUrls.length === 0) {
      throw new Error("Could not find any blog post URLs. The website's design might have changed.");
    }

    for (const postUrl of allPostUrls) {
      Logger.log(`Processing: ${postUrl}`);
      const html = UrlFetchApp.fetch(postUrl, { muteHttpExceptions: true }).getContentText();
      const contentData = cl_getPostContentAndTitle_(html);
      
      if (contentData && contentData.content.length > 50) {
        const analysis = cl_analyzeContentWithGemini_(contentData.title, contentData.content);
        if (analysis) {
          const [analyzedTitle, description, persona] = analysis;
          sheet.appendRow([postUrl, contentData.title, description, persona]);
        }
      } else {
         Logger.log(`Skipping ${postUrl} due to insufficient content.`);
      }
      Utilities.sleep(1000); // Pause to be respectful to the server
    }

    ui.alert("Success!", `The Content Library has been created with ${allPostUrls.length} posts.`, ui.ButtonSet.OK);

  } catch (e) {
    ui.alert("An Error Occurred", e.message, ui.ButtonSet.OK);
    Logger.log(e);
  }
}

/**
 * Scrapes blog index pages to find all individual post URLs.
 * @param {string[]} urls - An array of blog index page URLs.
 * @returns {string[]} A unique array of blog post URLs.
 */
function cl_getAllPostUrls_(urls) {
  const postUrls = new Set();
  const regex = /<a href="(\/blog\/[^"]+)"/g;

  urls.forEach(url => {
    try {
      const html = UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText();
      let match;
      while ((match = regex.exec(html)) !== null) {
        if (!match[1].includes('/category/')) {
          postUrls.add("https://dreamdata.io" + match[1]);
        }
      }
    } catch (e) {
      Logger.log(`Failed to fetch or parse URL: ${url}. Error: ${e.message}`);
    }
  });
  return Array.from(postUrls);
}

/**
 * Extracts the title and clean text content from a blog post's HTML.
 * @param {string} html - The HTML content of a blog post.
 * @returns {Object|null} An object with title and content, or null.
 */
function cl_getPostContentAndTitle_(html) {
  try {
    const titleMatch = html.match(/<h1[^>]*>([\s\S]*?)<\/h1>/);
    const title = titleMatch ? titleMatch[1].trim().replace(/&nbsp;/g, ' ') : "Title not found";

    const contentMatch = html.match(/<div class="blog-item-content-wrapper"[\s\S]*?>([\s\S]*?)<\/section>/);
    let content = "Content not found";
    if (contentMatch) {
      content = contentMatch[1].replace(/<[^>]*>/g, ' ').replace(/\s\s+/g, ' ').trim();
    }
    
    return { title, content };
  } catch (e) {
    Logger.log(`Failed to get content. Error: ${e.message}`);
    return null;
  }
}

/**
 * Sends content to the Gemini API for analysis and returns a parsed CSV row.
 * @param {string} title - The title of the blog post.
 * @param {string} content - The text content of the blog post.
 * @returns {string[]|null} An array of [title, description, persona].
 */
function cl_analyzeContentWithGemini_(title, content) {
  const apiKey = cfg_('GEMINI_API_KEY'); // Using the cfg_ utility
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${apiKey}`;
  const truncatedContent = content.substring(0, 15000);

  const prompt = `
    You are an expert B2B Content Analyst for Dreamdata, a B2B GTM Attribution Platform.
    TARGET PERSONAS: CMO / VP Marketing, VP Demand Generation, Head of Marketing Ops, Head of Performance Marketing.

    ANALYZE THE FOLLOWING BLOG POST:
    Title: "${title}"
    Content: "${truncatedContent}"

    YOUR TASK:
    1.  Description / Key Takeaway: Write a single, crisp sentence describing the key problem the article solves.
    2.  Target Persona: Identify the primary target persona from the list above.

    OUTPUT FORMAT:
    You MUST respond with a single line of CSV with THREE fields, each enclosed in double quotes: "Title","Description / Key Takeaway","Target Persona".
    The title in your output MUST MATCH the input title exactly.
  `;

  const payload = { 
    "contents": [{ "parts": [{ "text": prompt }] }],
    "generationConfig": { "maxOutputTokens": 512 }
  };
  const options = { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(payload), 'muteHttpExceptions': true };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    const jsonResponse = JSON.parse(responseText);
    
    if (jsonResponse.candidates && jsonResponse.candidates[0].content.parts[0].text) {
      let result = jsonResponse.candidates[0].content.parts[0].text.trim();
      // Simple CSV parser for "field1","field2","field3"
      return result.split('","').map(s => s.replace(/"/g, ''));
    } else {
      Logger.log(`Invalid API response for title "${title}": ${responseText}`);
      return [title, "Analysis Failed: Invalid API response", "N/A"];
    }
  } catch (e) {
    Logger.log(`API call failed for title "${title}". Error: ${e.message}`);
    return [title, `Analysis Failed: ${e.message}`, "N/A"];
  }
}

/**** ===== ADMIN & SETUP ===== ****/

/**
 * A one-time setup function for the script owner to securely store their API keys.
 */
// Replace this entire function
function setupApiKeys() {
  const ui = SpreadsheetApp.getUi();
  const userProperties = PropertiesService.getUserProperties();

  const keysToSet = ['HUBSPOT_TOKEN', 'APOLLO_API_KEY', 'GEMINI_API_KEY'];
  
  for (const key of keysToSet) {
    const response = ui.prompt(`Set API Key`, `Please enter your value for:\n\n${key}`, ui.ButtonSet.OK_CANCEL);
    
    if (response.getSelectedButton() == ui.Button.OK) {
      const value = response.getResponseText().trim();
      if (value) {
        userProperties.setProperty(key, value);
        // --- FIX: Added the required ui.ButtonSet.OK parameter ---
        ui.alert('Success', `The key for "${key}" has been set securely.`, ui.ButtonSet.OK);
      } else {
        // --- FIX: Added the required ui.ButtonSet.OK parameter ---
        ui.alert('Skipped', `No value entered for "${key}". It was not set.`, ui.ButtonSet.OK);
      }
    } else {
      // This alert call was also missing the button parameter.
      ui.alert('Cancelled', 'The setup process was cancelled.', ui.ButtonSet.OK);
      return;
    }
  }
  // --- FIX: Added the required ui.ButtonSet.OK parameter ---
  ui.alert('Setup Complete', 'All API keys have been processed. You can now safely remove them from your "Config" sheet.', ui.ButtonSet.OK);
}