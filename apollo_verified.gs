/**
 * @OnlyCurrentDoc
 */

// A cache to store person and organization details we've already looked up in this run.
const personCache = CacheService.getScriptCache();
const orgCache = CacheService.getScriptCache();

/**
 * Finds contacts, verifies employment, retrieves email, intelligently saves them as contacts,
 * and adds the final, actionable contacts to the 'Contacts' sheet.
 */
function apolloFindAndVerifyContactsForAccounts() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accSh = ss.getSheetByName(ACC_SHEET);
  const conSh = ss.getSheetByName(CON_SHEET);

  if (accSh.getLastRow() < 2) { ui.alert('No accounts to process.'); return; }

  const selectedAccounts = accSh.getRange(2, 1, accSh.getLastRow() - 1, accSh.getLastColumn()).getValues().filter(row => row[0] === true);
  if (selectedAccounts.length === 0) { ui.alert('No accounts are selected.'); return; }

  const contactsPerCompanyInput = ui.prompt('Enter the number of contacts (from your Persona) to find per company:', '5', ui.ButtonSet.OK_CANCEL);
  if (contactsPerCompanyInput.getSelectedButton() !== ui.Button.OK) return;
  const contactsPerCompany = parseInt(contactsPerCompanyInput.getResponseText(), 10);
  if (isNaN(contactsPerCompany) || contactsPerCompany <= 0) { ui.alert('Invalid number.'); return; }

  console.log(`--- Starting 'Find & Verify' Process using Persona ---`);
  console.log(`Accounts selected: ${selectedAccounts.length}. Contacts with email requested: ${contactsPerCompany}.`);

  const response = ui.alert('Find, Verify & Save Contacts?', `This will find, verify, get emails, and SAVE contacts to your Apollo instance. Continue?`, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) { console.log('User cancelled the operation.'); return; }

  const personaId = cfg_('APOLLO_PERSONA_ID');
  if (!personaId) {
    ui.alert('Error: APOLLO_PERSONA_ID is not defined in your Config sheet.');
    return;
  }
  console.log(`Using Persona ID: ${personaId}`);
  let totalContactsAdded = 0;

  for (const account of selectedAccounts) {
    const companyName = account[2];
    const companyDomain = account[3];
    console.log(`\n[${companyName}] ==> Starting process.`);
    SpreadsheetApp.getActiveSpreadsheet().toast(`Processing ${companyName}...`);

    const orgId = getApolloOrgId_(companyDomain);
    if (!orgId) {
      console.error(`[${companyName}] Could not find an Apollo Organization ID for domain "${companyDomain}". Skipping.`);
      continue;
    }
    console.log(`[${companyName}] Found Organization ID: ${orgId}`);

    let verifiedContactsForThisAccount = [];
    let page = 1;
    const maxPagesToSearch = 10;

    while (verifiedContactsForThisAccount.length < contactsPerCompany && page <= maxPagesToSearch) {
      console.log(`[${companyName}] Fetching page ${page} of potential contacts.`);

      const searchPayload = {
        organization_ids: [orgId],
        q_person_persona_ids: [personaId],
        page: page,
        per_page: 25
      };

      const searchResult = apolloSearcher_(searchPayload);
      if ((!searchResult.people || searchResult.people.length === 0) && (!searchResult.contacts || searchResult.contacts.length === 0)) {
        console.log(`[${companyName}] No more results found on page ${page}.`);
        break;
      }
      
      // --- THE FIX: Differentiate between existing contacts and new people ---
      const contactsFound = (searchResult.contacts || []).map(c => ({ ...c, is_already_contact: true }));
      const peopleFound = (searchResult.people || []).map(p => ({ ...p, is_already_contact: false }));
      const candidates = contactsFound.concat(peopleFound);
      
      console.log(`[${companyName}] Found ${candidates.length} potential contacts. Starting verification...`);

      for (const candidate of candidates) {
        const personId = candidate.person_id || candidate.id;
        if (!personId) continue;
        if (verifiedContactsForThisAccount.length >= contactsPerCompany) {
          console.log(`[${companyName}] Quota of ${contactsPerCompany} met.`);
          break;
        }

        let personDetails = JSON.parse(personCache.get(personId));
        if (!personDetails) {
          personDetails = apolloPersonById_(personId);
          if (personDetails) personCache.put(personId, JSON.stringify(personDetails), 3600);
        }

        if (personDetails && personDetails.person && personDetails.person.organization && personDetails.person.organization.name) {
          const currentPerson = personDetails.person;
          const currentCompanyName = currentPerson.organization.name;

          if (normalizeString_(currentCompanyName) === normalizeString_(companyName)) {
            console.log(`[${companyName}]   - SUCCESS: ${currentPerson.name} is verified.`);
            const email = apolloEnrichAndGetEmail_(currentPerson.id);

            if (email) {
              console.log(`[${companyName}]     - SUCCESS: Found email for ${currentPerson.name}: ${email}`);

              // --- THE FIX: Smarter save logic ---
              let contactId = null;
              if (candidate.is_already_contact) {
                  contactId = candidate.id;
                  console.log(`[${companyName}]       -> SUCCESS: Person is already a contact with ID: ${contactId}`);
              } else {
                  const newContact = apolloSavePersonAsContact_(personId);
                  if (newContact && newContact.id) {
                      contactId = newContact.id;
                      console.log(`[${companyName}]       -> SUCCESS: Saved as new Contact with ID: ${contactId}`);
                  } else {
                      console.error(`[${companyName}]       -> FAILED: Could not save ${currentPerson.name} as a contact.`);
                  }
              }

              if (contactId) {
                  const contactRow = [
                    false, companyDomain, currentPerson.name || '', currentPerson.title || '',
                    'Warm-Up', email, contactId, personId,
                    '', '', '', 'Contact Saved'
                  ];
                  verifiedContactsForThisAccount.push(contactRow);
              }

            } else {
              console.log(`[${companyName}]     - FAILED: No email found for ${currentPerson.name}. Discarding.`);
            }
          } else {
            console.warn(`[${companyName}]   - FAILED: Company mismatch for ${currentPerson.name}. Expected: "${companyName}", Found: "${currentCompanyName}".`);
          }
        } else {
          console.error(`[${companyName}]   - ERROR: Could not retrieve valid details for ID ${personId}.`);
        }
      }
      page++;
    }

    if (verifiedContactsForThisAccount.length > 0) {
      conSh.getRange(conSh.getLastRow() + 1, 1, verifiedContactsForThisAccount.length, 12).setValues(verifiedContactsForThisAccount);
      totalContactsAdded += verifiedContactsForThisAccount.length;
      console.log(`[${companyName}] <== Wrote ${verifiedContactsForThisAccount.length} verified contacts.`);
    } else {
      console.log(`[${companyName}] <== No verified contacts were found to write.`);
    }
  }

  console.log(`\n--- 'Find & Verify' Process Complete ---`);
  ui.alert('Process Complete', `Added a total of ${totalContactsAdded} new contacts.`, ui.ButtonSet.OK);
}

// =================================================================================================
// HELPER FUNCTIONS (No changes were made to the helpers)
// =================================================================================================

function apolloSavePersonAsContact_(personId) {
  const apiKey = cfg_('APOLLO_API_KEY');
  const url = 'https://api.apollo.io/v1/contacts';
  const payload = {
    api_key: apiKey,
    person_ids: [personId]
  };
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Cache-Control': 'no-cache' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  console.log(`--> Apollo Save Contact Call for Person ID: ${personId}`);

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseBody = response.getContentText();
    console.log(`<-- Apollo Save Contact Response (Status: ${response.getResponseCode()})`);

    if (response.getResponseCode() === 200 || response.getResponseCode() === 201) {
      const data = JSON.parse(responseBody);
      return data.contacts && data.contacts.length > 0 ? data.contacts[0] : null;
    }
    return null;
  } catch (e) {
    console.error(`Error in apolloSavePersonAsContact_: ${e.message}`);
    return null;
  }
}

function apolloEnrichAndGetEmail_(personId) {
  const apiKey = cfg_('APOLLO_API_KEY');
  const url = 'https://api.apollo.io/v1/people/enrich';
  const payload = {
    api_key: apiKey,
    id: personId
  };
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Cache-Control': 'no-cache' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const logPayload = { id: personId, api_key: 'REDACTED' };
  console.log(`--> Apollo Enrich Call:\nURL: ${url}\nPayload: ${JSON.stringify(logPayload, null, 2)}`);

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseBody = response.getContentText();
    console.log(`<-- Apollo Enrich Response (Status: ${response.getResponseCode()})`);

    if (response.getResponseCode() === 200) {
      const data = JSON.parse(responseBody);
      return data.person && data.person.email ? data.person.email : null;
    }
    return null;
  } catch (e) {
    console.error(`Error in apolloEnrichAndGetEmail_: ${e.message}`);
    return null;
  }
}

function getApolloOrgId_(domain) {
  let cachedId = orgCache.get(domain);
  if (cachedId) {
    console.log(`[${domain}] Org ID found in cache.`);
    return cachedId;
  }
  
  const apiKey = cfg_('APOLLO_API_KEY');
  const url = `https://api.apollo.io/v1/organizations/enrich?api_key=${apiKey}&domain=${domain}`;
  const options = {
    method: 'get',
    contentType: 'application/json',
    headers: { 'Cache-Control': 'no-cache' },
    muteHttpExceptions: true,
  };

  console.log(`--> Apollo Org ID Call:\nURL: ${url.replace(apiKey, 'REDACTED')}`);
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseBody = response.getContentText();
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(responseBody);
      if (data.organization && data.organization.id) {
        const orgId = data.organization.id;
        console.log(`<-- Apollo Org ID Response (Status: 200)`);
        orgCache.put(domain, orgId, 21600);
        return orgId;
      }
    }
    console.error(`<-- Apollo Org ID Response (Status: ${response.getResponseCode()}):\n${responseBody}`);
    return null;
  } catch (e) {
    console.error(`Error in getApolloOrgId_: ${e.message}`);
    return null;
  }
}

function apolloSearcher_(payload) {
  const apiKey = cfg_('APOLLO_API_KEY');
  payload.api_key = apiKey;
  const url = 'https://api.apollo.io/v1/mixed_people/search';
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Cache-Control': 'no-cache' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const logPayload = JSON.parse(JSON.stringify(payload));
  logPayload.api_key = 'REDACTED';
  console.log(`--> Apollo Search Call:\nPayload: ${JSON.stringify(logPayload, null, 2)}`);

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseBody = response.getContentText();
    if (response.getResponseCode() === 200) {
       console.log(`<-- Apollo Search Response (Status: 200)`);
       return JSON.parse(responseBody);
    }
    console.error(`<-- Apollo Search Response (Status: ${response.getResponseCode()}):\n${responseBody}`);
    return null;
  } catch (e) {
    console.error(`Error in apolloSearcher_: ${e.message}`);
    return null;
  }
}

function apolloPersonById_(personId) {
  const apiKey = cfg_('APOLLO_API_KEY');
  const url = `https://api.apollo.io/v1/people/${personId}?api_key=${apiKey}`;
  const options = {
    method: 'get',
    contentType: 'application/json',
    headers: { 'Cache-Control': 'no-cache' },
    muteHttpExceptions: true,
  };

  const logUrl = url.replace(apiKey, 'REDACTED');
  console.log(`--> Apollo Person Call:\nURL: ${logUrl}`);

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseBody = response.getContentText();
    if (response.getResponseCode() === 200) {
      console.log(`<-- Apollo Person Response (Status: 200) for ID ${personId}`);
      return JSON.parse(responseBody);
    }
     console.error(`<-- Apollo Person Response (Status: ${response.getResponseCode()}) for ID ${personId}:\n${responseBody}`);
    return null;
  } catch (e) {
    console.error(`Error in apolloPersonById_ for ID ${personId}: ${e.message}`);
    return null;
  }
}

function normalizeString_(name) {
  if (!name) return '';
  return name.toLowerCase().replace(/,|\.|inc|ltd|llc|co|corp|corporation/g, '').trim();
}