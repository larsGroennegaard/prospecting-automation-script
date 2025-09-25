
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

function apolloAddContactsToSequence() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const conSh = ss.getSheetByName(CON_SHEET);

  if (conSh.getLastRow() < 2) { ui.alert('No contacts to process.'); return; }
  
  // --- NEW: Get column index for assigned email ---
  const headers = conSh.getRange(1, 1, 1, conSh.getLastColumn()).getValues()[0];
  const assignedEmailColIdx = headers.indexOf('assigned_sending_email');
  if (assignedEmailColIdx === -1) {
    ui.alert('Error: Could not find the "assigned_sending_email" column.');
    return;
  }

  const allContacts = conSh.getRange(2, 1, conSh.getLastRow() - 1, conSh.getLastColumn()).getValues();
  const selectedContacts = allContacts.filter(row => row[0] === true);

  if (selectedContacts.length === 0) { ui.alert('No contacts are selected.'); return; }
  
  const response = ui.alert('Add Contacts to Apollo Sequence?', `This will add ${selectedContacts.length} selected contacts to the sequence. Continue?`, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  const apolloMailboxMap = getApolloMailboxMap_();
  const sequenceId = cfg_('APOLLO_SEQUENCE_ID');
  if (apolloMailboxMap.size === 0) { ui.alert('Could not fetch Apollo mailboxes. Check API key and permissions.'); return; }

  // --- NEW: Simplified grouping logic ---
  const contactsBySender = new Map();
  allContacts.forEach((row, index) => {
    const isSelected = row[0];
    const status = row[11];
    const apolloContactId = row[6];
    const assignedEmail = row[assignedEmailColIdx]; // Read from the new column

    if (isSelected && apolloContactId && assignedEmail && status.includes('from_apollo_contact') && !status.includes('apollo_sequenced')) {
      if (!contactsBySender.has(assignedEmail)) {
        contactsBySender.set(assignedEmail, []);
      }
      contactsBySender.get(assignedEmail).push({ id: apolloContactId, rowIndex: index + 2 });
    }
  });

  let totalSuccess = 0, totalFailed = 0;
  for (const [sendingEmail, contacts] of contactsBySender.entries()) {
    const mailboxId = apolloMailboxMap.get(sendingEmail.toLowerCase());
    if (!mailboxId) {
      contacts.forEach(c => {
        const currentStatus = conSh.getRange(c.rowIndex, 12).getValue();
        conSh.getRange(c.rowIndex, 12).setValue(`${currentStatus};sequence_failed(no_mailbox)`);
        totalFailed++;
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
        totalFailed++;
      });
    }
    SpreadsheetApp.flush();
  }
  ui.alert(`Sequence enrollment complete.\n\nAdded: ${totalSuccess}\nFailed/Skipped: ${totalFailed}`);
}



function apolloFindContactsForAccounts() {
  const ui = SpreadsheetApp.getUi();

  const accSh = ensureAccountsHeader_();
  const conSh = ensureContactsHeader_();

  const SEARCH_URL = 'https://api.apollo.io/v1/mixed_people/search';
  const ORGS_SEARCH_URL = 'https://api.apollo.io/v1/accounts/search';
  
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
      const email = String(r[5] || '').toLowerCase();
      if (email) existing.add(email);
    });
  }

  const headers = { 'x-api-key': cfg_('APOLLO_API_KEY'), 'Content-Type': 'application/json' };
  
  const fTitles = getCfgList_('APOLLO_TITLES');
  const includeSimilarTitles = String(cfg_('APOLLO_INCLUDE_SIMILAR_TITLES')).toLowerCase() === 'true';

  function buildSearchPayload(base) {
    const payload = Object.assign({
      page: 1, per_page: 100,
      person_titles: fTitles.length ? fTitles : undefined,
    }, base);
    
    // =================================================================
    // THIS IS THE FIX: Ensure the correct parameter name is used.
    if (payload.organization_domains) {
      payload.q_organization_domains = payload.organization_domains;
      delete payload.organization_domains;
    }
    // =================================================================
    
    if (payload.person_titles && includeSimilarTitles) payload.include_similar_titles = true;
    return payload;
  }

  function processResults(people, domain) {
    const rows = [];
    for (const p of people) {
      let email = String(p.email || '').toLowerCase();
      if (!email || existing.has(email)) continue;
      
      const contactId = p.contact ? p.contact.id : null;
      const personId = p.id;

      rows.push([
        true, // selected
        domain, // company_domain
        p.name || 'N/A', // contact_name
        p.title || '', // title
        p.contact && p.contact.contact_stage_id ? (stageMap.get(p.contact.contact_stage_id) || '') : '', // stage
        email, // email
        contactId, // apollo_contact_id
        personId, // apollo_person_id
        p.contact ? p.contact.hubspot_vid : '', // hubspot_contact_id
      ]);
      existing.add(email);
    }
    return rows;
  }

  let totalAppended = 0;
  for (const acc of accounts) {
    const { domain, name } = acc;
    console.log(`\n\n=== Starting Apollo search for company: "${name}" (domain: "${domain}") ===`);
    
    let allPotentialPeople = [];
    let searchPayloadBase = { organization_domains: [domain] };

    let page = 1;
    while (true) {
      try {
        const payload = buildSearchPayload({ ...searchPayloadBase, page });
        const data = httpJson_(SEARCH_URL, { method: 'post', headers, payload: JSON.stringify(payload) });
        
        const people = (data.people || []);
        if (people.length > 0) allPotentialPeople.push(...people);
        
        const pg = data.pagination || {};
        if (pg.page && pg.total_pages && pg.page < pg.total_pages) {
          page++;
        } else {
          if (allPotentialPeople.length === 0 && !searchPayloadBase.organization_ids) {
            console.log(`No results via domain search for ${domain}. Trying by name: ${name}`);
            let org = null;
            try {
              if (name) {
                 const orgsData = httpJson_(ORGS_SEARCH_URL, { method: 'post', headers, payload: JSON.stringify({ q_organization_name: name }) });
                 org = orgsData.accounts && orgsData.accounts[0];
              }
            } catch(e) { console.error(`Error while finding Organization ID: ${e.message}`); }
            
            if (org) {
              console.log(`Found Org match: id=${org.id}. Restarting search with Org ID.`);
              searchPayloadBase = { organization_ids: [org.id] };
              allPotentialPeople = [];
              page = 1;
              continue;
            }
          }
          break;
        }
      } catch (e) { console.error(e); break; }
    }
    
    console.log(`Found ${allPotentialPeople.length} total potential people for ${domain}.`);

    const stageFiltered = allPotentialPeople.filter(p => {
        const stageId = p.contact ? p.contact.contact_stage_id : null;
        const stageName = stageId ? (stageMap.get(stageId) || '') : '';
        return allowedStages.includes(stageName.toLowerCase());
    });

    const sorted = stageFiltered.sort((a, b) => (a.email ? -1 : 1));

    const finalPeople = [];
    for (const p of sorted) {
      if (finalPeople.length >= 5) break;
      if (p.email && !existing.has(p.email.toLowerCase())) {
        finalPeople.push(p);
      } else if (!p.email) {
        try {
          const enriched = apolloEnrichPerson_(p.id);
          if (enriched.person && enriched.person.email && !existing.has(enriched.person.email.toLowerCase())) {
            finalPeople.push(enriched.person);
          }
        } catch (e) { /* ignore enrich error */ }
      }
    }
    
    const personIdsToSave = finalPeople.filter(p => !p.contact).map(p => p.id);
    if (personIdsToSave.length > 0) {
        try { apolloSaveAsContact_(personIdsToSave); }
        catch (e) { console.error("Failed to save some people as contacts", e); }
    }

    const rowsToAdd = processResults(finalPeople, domain);
    if (rowsToAdd.length > 0) {
      conSh.getRange(conSh.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
      totalAppended += rowsToAdd.length;
    }
  }
  ui.alert(`Apollo contact search complete. Appended a total of ${totalAppended} new contacts.`);
}