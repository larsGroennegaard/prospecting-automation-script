/**
 * @OnlyCurrentDoc
 */

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
    const isSelected = row[0];        // col A
    const apolloContactId = row[6];   // col G
    const status = row[11] || '';     // col L
    const email1Subject = row[12];    // col M

    // --- THE FIX: Updated the condition to work with the new status ---
    if (isSelected && apolloContactId && email1Subject && !status.includes('apollo_pushed')) {
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
    } else if (isSelected && status.includes('apollo_pushed')) {
      skippedCount++;
    }
  });
  ui.alert(`Push to Apollo complete.\n\nUpdated: ${processedCount}\nFailed: ${errorCount}\nSkipped (already pushed): ${skippedCount}`);
}


function apolloAddContactsToSequence() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const conSh = ss.getSheetByName(CON_SHEET);

  if (conSh.getLastRow() < 2) { ui.alert('No contacts to process.'); return; }
  
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

  const contactsBySender = new Map();
  allContacts.forEach((row, index) => {
    const isSelected = row[0];
    const status = row[11] || '';
    const apolloContactId = row[6];
    const assignedEmail = row[assignedEmailColIdx]; 

    // --- THE FIX: Updated the condition to work with the new workflow ---
    if (isSelected && apolloContactId && assignedEmail && !status.includes('apollo_sequenced')) {
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