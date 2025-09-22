/**
 * @file rotator.gs
 * @description Contains the logic for assigning rotating sender personas to contacts.
 */

/**
 * Assigns a sending email and sender name to each selected contact in a round-robin fashion
 * based on the personas defined in the 'Mailbox Mapping' sheet.
 */
function assignSendersForRotation() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const conSh = ss.getSheetByName(CON_SHEET);
  const accSh = ss.getSheetByName(ACC_SHEET);
  const mapSh = ss.getSheetByName(MAILBOX_MAP_SHEET);

  if (conSh.getLastRow() < 2) {
    ui.alert('No contacts to process.');
    return;
  }

  const allContacts = conSh.getRange(2, 1, conSh.getLastRow() - 1, conSh.getLastColumn()).getValues();
  const selectedContactsIndices = [];
  allContacts.forEach((row, index) => {
    if (row[0] === true) {
      selectedContactsIndices.push(index);
    }
  });

  if (selectedContactsIndices.length === 0) {
    ui.alert('No contacts are selected.');
    return;
  }
  
  const response = ui.alert('Assign Senders?', `This will assign a sending persona to the ${selectedContactsIndices.length} selected contacts. This will overwrite any existing assignments. Continue?`, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  // 1. Build the Persona Map from the Mailbox Mapping sheet
  // The map will look like: { 'owner_email': [ {email: 'sender@...', name: 'Sender'}, ... ] }
  const ownerPersonaMap = new Map();
  if (mapSh && mapSh.getLastRow() > 1) {
    const mappingValues = mapSh.getRange(2, 1, mapSh.getLastRow() - 1, 3).getValues();
    mappingValues.forEach(row => {
      const ownerEmail = String(row[0]).toLowerCase().trim();
      const sendingEmail = String(row[1]).toLowerCase().trim();
      const senderName = String(row[2]).trim();
      if (ownerEmail && sendingEmail && senderName) {
        if (!ownerPersonaMap.has(ownerEmail)) {
          ownerPersonaMap.set(ownerEmail, []);
        }
        ownerPersonaMap.get(ownerEmail).push({ email: sendingEmail, name: senderName });
      }
    });
  }

  if (ownerPersonaMap.size === 0) {
    ui.alert('Could not find any valid entries in the "Mailbox Mapping" sheet.');
    return;
  }

  // 2. Build the Company to Owner Map
  const companyToOwnerMap = new Map();
  if (accSh && accSh.getLastRow() > 1) {
    const accValues = accSh.getRange(2, 1, accSh.getLastRow() - 1, 7).getValues();
    accValues.forEach(row => {
      const domain = String(row[3]).toLowerCase().trim();
      const ownerEmail = String(row[6]).toLowerCase().trim();
      if (domain && ownerEmail) {
        companyToOwnerMap.set(domain, ownerEmail);
      }
    });
  }

  // 3. Group selected contacts by their HubSpot Owner
  const contactsByOwner = new Map();
  selectedContactsIndices.forEach(index => {
    const contactRow = allContacts[index];
    const domain = String(contactRow[1]).toLowerCase().trim();
    const ownerEmail = companyToOwnerMap.get(domain);
    if (ownerEmail) {
      if (!contactsByOwner.has(ownerEmail)) {
        contactsByOwner.set(ownerEmail, []);
      }
      contactsByOwner.get(ownerEmail).push(index); // Store the row index
    }
  });

  // Get the column indices for the new columns
  const headers = conSh.getRange(1, 1, 1, conSh.getLastColumn()).getValues()[0];
  const assignedEmailColIdx = headers.indexOf('assigned_sending_email');
  const assignedNameColIdx = headers.indexOf('assigned_sender_name');

  if (assignedEmailColIdx === -1 || assignedNameColIdx === -1) {
      ui.alert('Error: Could not find the required columns "assigned_sending_email" and "assigned_sender_name" in the "Contacts" sheet. Please ensure they exist.');
      return;
  }

  let assignmentsMade = 0;

  // 4. Iterate through each owner's contacts and assign personas
  for (const [ownerEmail, contactIndices] of contactsByOwner.entries()) {
    const personas = ownerPersonaMap.get(ownerEmail);
    if (personas && personas.length > 0) {
      contactIndices.forEach((contactIndex, i) => {
        const personaIndex = i % personas.length; // Round-robin logic
        const assignedPersona = personas[personaIndex];
        
        // Write the assigned email and name to the sheet
        conSh.getRange(contactIndex + 2, assignedEmailColIdx + 1).setValue(assignedPersona.email);
        conSh.getRange(contactIndex + 2, assignedNameColIdx + 1).setValue(assignedPersona.name);
        assignmentsMade++;
      });
    }
  }

  ui.alert('Assignment Complete', `Assigned sender personas to ${assignmentsMade} contacts.`, ui.ButtonSet.OK);
}