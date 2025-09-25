/**
 * @file enrichment.gs
 * @description Contains functions for enriching contacts and accounts with external data.
 */

/**
 * For selected contacts, calls the Apollo /people/enrich API to get their
 * professional headline and pastes it into the 'contact_summary' column.
 */
function enrichContactsFromApollo() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const conSh = ss.getSheetByName(CON_SHEET);

  if (conSh.getLastRow() < 2) {
    ui.alert('No contacts to process.');
    return;
  }

  const headers = conSh.getRange(1, 1, 1, conSh.getLastColumn()).getValues()[0];
  // TINY CHANGE 1: Look for the 'apollo_person_id' column
  const personIdColIdx = headers.indexOf('apollo_person_id');
  const summaryColIdx = headers.indexOf('contact_summary');

  if (personIdColIdx === -1 || summaryColIdx === -1) {
    ui.alert('Error: Could not find required columns "apollo_person_id" or "contact_summary".');
    return;
  }

  const allContacts = conSh.getRange(2, 1, conSh.getLastRow() - 1, conSh.getLastColumn()).getValues();
  const selectedContacts = [];
  allContacts.forEach((row, index) => {
    if (row[0] === true) {
      selectedContacts.push({
        personId: row[personIdColIdx], // TINY CHANGE 2: Get the ID from the correct column
        rowIndex: index + 2
      });
    }
  });

  if (selectedContacts.length === 0) {
    ui.alert('No contacts are selected.');
    return;
  }

  const response = ui.alert('Enrich Contacts from Apollo?', `This will enrich ${selectedContacts.length} selected contacts. Continue?`, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  let processedCount = 0;
  let errorCount = 0;
  ss.toast('Enriching contacts from Apollo... Please wait.', 'Processing...', -1);

  selectedContacts.forEach(contact => {
    if (contact.personId) { // TINY CHANGE 3: Check for and use the personId
      try {
        const url = 'https://api.apollo.io/v1/people/enrich';
        const payload = { id: contact.personId };
        const data = apolloPost_(url, payload);

        const headline = data.person && data.person.headline ? data.person.headline : '';

        conSh.getRange(contact.rowIndex, summaryColIdx + 1).setValue(headline);
        processedCount++;
      } catch (e) {
        console.error(`Failed to enrich contact with Person ID ${contact.personId}: ${e.message}`);
        conSh.getRange(contact.rowIndex, summaryColIdx + 1).setValue(`Error`);
        errorCount++;
      }
    }
  });

  ss.toast(`Enrichment complete. Processed: ${processedCount}, Failed: ${errorCount}`, 'Success!', 5);
}