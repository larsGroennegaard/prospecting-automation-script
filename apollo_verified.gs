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
 * INCLUDES DETAILED LOGGING.
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
  
  console.log(`--- Starting 'Find & Verify' Process ---`);
  console.log(`Accounts selected: ${selectedAccounts.length}. Contacts requested per account: ${contactsPerCompany}.`);

  const response = ui.alert('Find & Verify Contacts?', `This will search for up to ${contactsPerCompany} VERIFIED contacts for the ${selectedAccounts.length} selected accounts. This may take a while and consume more API credits. Continue?`, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) {
    console.log('User cancelled the operation.');
    return;
  }

  const titles = cfg_('APOLLO_TITLES').split(',').map(t => t.trim());
  let totalContactsAdded = 0;

  for (const account of selectedAccounts) {
    const companyName = account[2];
    const companyDomain = account[3];
    
    console.log(`\n[${companyName}] ==> Starting search.`);
    let verifiedContactsForThisAccount = [];
    let page = 1;
    const maxPagesToSearch = 5; // To prevent infinite loops

    SpreadsheetApp.getActiveSpreadsheet().toast(`Searching for contacts at ${companyName}...`);

    while (verifiedContactsForThisAccount.length < contactsPerCompany && page <= maxPagesToSearch) {
      console.log(`[${companyName}] Fetching page ${page} of potential contacts.`);
      const searchPayload = {
        q_organization_domains: [companyDomain],
        person_titles: titles,
        page: page,
        per_page: 20 // Over-sample by fetching more results
      };

      const searchResult = apolloSearcher_(searchPayload);
      if (!searchResult || searchResult.people.length === 0) {
        console.log(`[${companyName}] No more results found on page ${page}.`);
        break; // Exit if no more people are found
      }
      
      console.log(`[${companyName}] Found ${searchResult.people.length} potential contacts on page ${page}. Starting verification...`);

      for (const person of searchResult.people) {
        if (verifiedContactsForThisAccount.length >= contactsPerCompany) {
          console.log(`[${companyName}] Quota of ${contactsPerCompany} met. Halting search for this company.`);
          break;
        }

        // --- VERIFICATION STEP ---
        let personDetails = JSON.parse(personCache.get(person.id));
        if (personDetails) {
          console.log(`[${companyName}]   - Verifying ${person.name} (ID: ${person.id}) from cache.`);
        } else {
          console.log(`[${companyName}]   - Verifying ${person.name} (ID: ${person.id}) via API call.`);
          personDetails = apolloPersonById_(person.id);
          if (personDetails) {
             personCache.put(person.id, JSON.stringify(personDetails), 3600); // Cache for 1 hour
          }
        }
       
        if (personDetails && personDetails.organization && personDetails.organization.name) {
          const currentCompanyName = personDetails.organization.name;
          if (normalizeString_(currentCompanyName) === normalizeString_(companyName)) {
            console.log(`[${companyName}]     -> SUCCESS: ${person.name} is verified at ${currentCompanyName}.`);
            const contactRow = [
              false, // selected
              companyDomain, personDetails.name || '', personDetails.title || '',
              'Warm-Up', personDetails.email || '', personDetails.id,
              personDetails.id, '', '', '', 'Verified',
            ];
            verifiedContactsForThisAccount.push(contactRow);
          } else {
            console.warn(`[${companyName}]     -> FAILED: Company mismatch for ${person.name}. Expected: "${companyName}", Found: "${currentCompanyName}".`);
          }
        } else {
          console.error(`[${companyName}]     -> ERROR: Could not retrieve valid person details for ID ${person.id}.`);
        }
      }
      page++;
    }

    if (verifiedContactsForThisAccount.length > 0) {
      conSh.getRange(conSh.getLastRow() + 1, 1, verifiedContactsForThisAccount.length, 12).setValues(verifiedContactsForThisAccount);
      totalContactsAdded += verifiedContactsForThisAccount.length;
      console.log(`[${companyName}] <== Process finished. Wrote ${verifiedContactsForThisAccount.length} verified contacts to the sheet.`);
    } else {
      console.log(`[${companyName}] <== Process finished. No verified contacts found to write.`);
    }
  }

  console.log(`\n--- 'Find & Verify' Process Complete ---`);
  console.log(`Added a total of ${totalContactsAdded} verified contacts for the ${selectedAccounts.length} selected accounts.`);
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
    console.error(`Failed to fetch person ${personId}. Status: ${response.getResponseCode()}, Response: ${response.getContentText()}`);
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