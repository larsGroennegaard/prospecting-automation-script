
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