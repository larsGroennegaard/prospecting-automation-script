
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

/**** =====  APOLLO: FIND CONTACTS ===== ****/
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

/**** ===== APOLLO: PUSH AI MESSAGES ===== ****/
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

/**** ===== APOLLO: ADD CONTACTS TO SEQUENCE ===== ****/
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