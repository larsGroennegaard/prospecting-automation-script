
function enrichContactsFromApollo() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const conSh = ss.getSheetByName(CON_SHEET);

  if (conSh.getLastRow() < 2) {
    ui.alert('No contacts to process.');
    return;
  }

  const headers = conSh.getRange(1, 1, 1, conSh.getLastColumn()).getValues()[0];
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
        personId: row[personIdColIdx],
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
    if (contact.personId) {
      try {
        const url = 'https://api.apollo.io/v1/people/enrich';
        const payload = { id: contact.personId };
        const data = apolloPost_(url, payload);

        if (data.person) {
          // Create the JSON object with all the rich context
          const enrichmentData = {
            headline: data.person.headline || '',
            employment_history: data.person.employment_history || [],
            company_summary: data.person.organization ? (data.person.organization.short_description || '') : '',
            company_technologies: data.person.organization ? (data.person.organization.technology_names || []) : []
          };

          // Convert the object to a JSON string and save it to the sheet
          const jsonString = JSON.stringify(enrichmentData, null, 2); // The '2' makes it nicely formatted
          conSh.getRange(contact.rowIndex, summaryColIdx + 1).setValue(jsonString);
          processedCount++;
        }

      } catch (e) {
        console.error(`Failed to enrich contact with Person ID ${contact.personId}: ${e.message}`);
        conSh.getRange(contact.rowIndex, summaryColIdx + 1).setValue(`Error`);
        errorCount++;
      }
    }
  });

  ss.toast(`Enrichment complete. Processed: ${processedCount}, Failed: ${errorCount}`, 'Success!', 5);
}