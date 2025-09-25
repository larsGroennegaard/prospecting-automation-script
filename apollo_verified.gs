/**
 * @OnlyCurrentDoc
 *
 * The above comment directs App Script to limit the scope of file access for this script to the Spreadsheet
 * this script is container-bound to. It does not use any other Google Docs scopes.
 */

// A cache to store person details we've already looked up in this run, to avoid redundant API calls.
const personCache = CacheService.getScriptCache();

/**
 * Finds contacts from Apollo for the selected accounts in the 'Accounts' sheet,
 * verifies they are still employed at the company, and adds the verified contacts to the 'Contacts' sheet.
 * This function loops through pages of search results and performs verification until the desired
 * number of contacts is found for each company.
 */
function apolloFindAndVerifyContactsForAccounts() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accSh = ss.getSheetByName(ACC_SHEET);
  const conSh = ss.getSheetByName(CON_SHEET);

  if (accSh.getLastRow() < 2) {
    ui.alert('No accounts to process.');
    return;
  }

  const selectedAccounts = accSh.getRange(2, 1, accSh.getLastRow() - 1, accSh.getLastColumn()).getValues().filter(row => row[0] === true);

  if (selectedAccounts.length === 0) {
    ui.alert('No accounts are selected. Please check the boxes in column A for the accounts you want to process.');
    return;
  }

  const contactsPerCompanyInput = ui.prompt('Enter the number of verified contacts to find per company:', '5', ui.ButtonSet.OK_CANCEL);
  if (contactsPerCompanyInput.getSelectedButton() !== ui.Button.OK) return;
  const contactsPerCompany = parseInt(contactsPerCompanyInput.getResponseText(), 10);
  if (isNaN(contactsPerCompany) || contactsPerCompany <= 0) {
    ui.alert('Invalid number. Please enter a positive integer.');
    return;
  }

  const response = ui.alert('Find & Verify Contacts?', `This will search for up to ${contactsPerCompany} VERIFIED contacts for the ${selectedAccounts.length} selected accounts. This may take a while and consume more API credits. Continue?`, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  const titles = cfg_('APOLLO_TITLES').split(',').map(t => t.trim());
  let totalContactsAdded = 0;

  for (const account of selectedAccounts) {
    const hubspotCompanyId = account[1];
    const companyName = account[2];
    const companyDomain = account[3];
    const hubspotOwnerEmail = account[6];
    
    let verifiedContactsForThisAccount = [];
    let page = 1;
    const maxPagesToSearch = 5; // To prevent infinite loops

    SpreadsheetApp.getActiveSpreadsheet().toast(`Searching for contacts at ${companyName}...`);

    while (verifiedContactsForThisAccount.length < contactsPerCompany && page <= maxPagesToSearch) {
      const searchPayload = {
        q_organization_domains: [companyDomain],
        person_titles: titles,
        page: page,
        per_page: 20 // Over-sample by fetching more results
      };

      const searchResult = apolloSearcher_(searchPayload);
      if (!searchResult || searchResult.people.length === 0) {
        console.log(`No more results found for ${companyName} on page ${page}.`);
        break; // Exit if no more people are found
      }

      for (const person of searchResult.people) {
        // Check if we already found enough contacts for this account
        if (verifiedContactsForThisAccount.length >= contactsPerCompany) {
          break;
        }

        // --- VERIFICATION STEP ---
        // Use a cache to avoid re-fetching the same person's details
        let personDetails = JSON.parse(personCache.get(person.id));
        if (!personDetails) {
          personDetails = apolloPersonById_(person.id);
          if (personDetails) {
             personCache.put(person.id, JSON.stringify(personDetails), 3600); // Cache for 1 hour
          }
        }
       
        if (personDetails && personDetails.organization && personDetails.organization.name) {
          const currentCompanyName = personDetails.organization.name;
          // Simple normalization for comparison
          if (normalizeString_(currentCompanyName) === normalizeString_(companyName)) {
            const contactRow = [
              false, // selected
              companyDomain,
              personDetails.name || '',
              personDetails.title || '',
              'Warm-Up', // Default stage
              personDetails.email || '',
              personDetails.id, // apollo_contact_id is the person id
              personDetails.id, // apollo_person_id
              '', // hubspot_contact_id
              '', // contact_story_30_days
              '', // contact_summary
              'Verified', // status
            ];
            verifiedContactsForThisAccount.push(contactRow);
          } else {
            console.log(`Verification failed for ${person.name}. Expected: ${companyName}, Found: ${currentCompanyName}`);
          }
        }
      }
      page++;
    }

    if (verifiedContactsForThisAccount.length > 0) {
      conSh.getRange(conSh.getLastRow() + 1, 1, verifiedContactsForThisAccount.length, 12)
           .setValues(verifiedContactsForThisAccount);
      totalContactsAdded += verifiedContactsForThisAccount.length;
      console.log(`Added ${verifiedContactsForThisAccount.length} verified contacts for ${companyName}.`);
    } else {
      console.log(`Could not find any verified contacts for ${companyName}.`);
    }
  }

  ui.alert('Process Complete', `Added a total of ${totalContactsAdded} verified contacts for the ${selectedAccounts.length} selected accounts.`, ui.ButtonSet.OK);
}

/**
 * Helper function to get full person details by Apollo ID.
 * This is crucial for verification. Leverages the /v1/people/{id} endpoint.
 * @param {string} personId The Apollo Person ID.
 * @return {Object|null} The person object from Apollo or null if not found.
 */
function apolloPersonById_(personId) {
  const apiKey = cfg_('APOLLO_API_KEY');
  const url = `https://api.apollo.io/v1/people/${personId}?api_key=${apiKey}`;
  const options = {
    method: 'get',
    contentType: 'application/json',
    muteHttpExceptions: true,
  };
  try {
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 200) {
      return JSON.parse(response.getContentText());
    }
    console.error(`Failed to fetch person ${personId}. Status: ${response.getResponseCode()}`);
    return null;
  } catch (e) {
    console.error(`Error in apolloPersonById_ for ID ${personId}: ${e.message}`);
    return null;
  }
}

/**
 * Normalizes company names for more reliable comparison.
 * Converts to lowercase and removes common business suffixes.
 * @param {string} name The company name.
 * @return {string} The normalized name.
 */
function normalizeString_(name) {
  if (!name) return '';
  return name.toLowerCase()
             .replace(/,|\.|inc|ltd|llc|co|corp|corporation/g, '')
             .trim();
}