/**** ===== SHEET NAMES ===== ****/
const CFG_SHEET = 'Config';
const ACC_SHEET = 'Accounts';
const CON_SHEET = 'Contacts';
const MAILBOX_MAP_SHEET = 'Mailbox Mapping';
const CONTENT_LIB_SHEET = 'Content Library'; // Add this line


/**** ===== MENU ===== ****/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // First, create the submenu for Admin tasks
  const adminMenu = ui.createMenu('Admin')
      .addItem('Set API Keys', 'setupApiKeys');

  // Then, create the main menu and add all items, including the submenu
  ui.createMenu('Prospecting')
    .addItem('1) Fetch HS Companies (marked)', 'hsFetchMarkedCompaniesToSheet')
    .addItem('2) Find Contacts in Apollo', 'apolloFindContactsForAccounts')
    .addSeparator()
    .addItem('Enrich: Get Account Stories from BQ', 'enrichFromDataWarehouse')
    .addItem('Enrich: Get Contact Journeys from BQ', 'enrichContactsFromBigQuery')
    .addSeparator()
    .addItem('3) Generate AI Messages', 'generateAiMessages')
    .addItem('4) Push AI Messages to Apollo', 'apolloPushMessages')
    .addItem('5) Add Contacts to Apollo Sequence', 'apolloAddContactsToSequence')
    .addSeparator()
    .addSubMenu(adminMenu) // Add the Admin menu as a submenu
    .addToUi();
}












/**** ===== 6) BIGQUERY: ENRICH ACCOUNTS ===== ****/

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

/**** ===== SETUP UTILITIES ===== ****/

/**
 * Ensures the Content Library sheet has the correct header.
 */
function ensureContentLibraryHeader_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONTENT_LIB_SHEET);
  if (!sh) throw new Error(`Missing sheet: ${CONTENT_LIB_SHEET}`);
  const expected = ["URL", "Title", "Description / Key Takeaway", "Target Persona"];
  if (sh.getLastRow() > 0) sh.clear(); // Clear the sheet before adding headers
  sh.appendRow(expected);
  SpreadsheetApp.flush();
  return sh;
}

/**
 * Main function to orchestrate the blog cataloging process.
 */
function cl_buildLibrary() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Rebuild Content Library?',
    'This will clear the "Content Library" sheet and rebuild it by scraping and analyzing your blog posts. This can take several minutes. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  const sheet = ensureContentLibraryHeader_();

  // You can add more blog index pages here if needed
  const blogIndexPages = [
    "https://dreamdata.io/blog",
    "https://dreamdata.io/blog?offset=1744801352826"
  ];

  try {
    const allPostUrls = cl_getAllPostUrls_(blogIndexPages);
    Logger.log(`Found ${allPostUrls.length} unique blog post URLs.`);

    if (allPostUrls.length === 0) {
      throw new Error("Could not find any blog post URLs. The website's design might have changed.");
    }

    for (const postUrl of allPostUrls) {
      Logger.log(`Processing: ${postUrl}`);
      const html = UrlFetchApp.fetch(postUrl, { muteHttpExceptions: true }).getContentText();
      const contentData = cl_getPostContentAndTitle_(html);
      
      if (contentData && contentData.content.length > 50) {
        const analysis = cl_analyzeContentWithGemini_(contentData.title, contentData.content);
        if (analysis) {
          const [analyzedTitle, description, persona] = analysis;
          sheet.appendRow([postUrl, contentData.title, description, persona]);
        }
      } else {
         Logger.log(`Skipping ${postUrl} due to insufficient content.`);
      }
      Utilities.sleep(1000); // Pause to be respectful to the server
    }

    ui.alert("Success!", `The Content Library has been created with ${allPostUrls.length} posts.`, ui.ButtonSet.OK);

  } catch (e) {
    ui.alert("An Error Occurred", e.message, ui.ButtonSet.OK);
    Logger.log(e);
  }
}

/**
 * Scrapes blog index pages to find all individual post URLs.
 * @param {string[]} urls - An array of blog index page URLs.
 * @returns {string[]} A unique array of blog post URLs.
 */
function cl_getAllPostUrls_(urls) {
  const postUrls = new Set();
  const regex = /<a href="(\/blog\/[^"]+)"/g;

  urls.forEach(url => {
    try {
      const html = UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText();
      let match;
      while ((match = regex.exec(html)) !== null) {
        if (!match[1].includes('/category/')) {
          postUrls.add("https://dreamdata.io" + match[1]);
        }
      }
    } catch (e) {
      Logger.log(`Failed to fetch or parse URL: ${url}. Error: ${e.message}`);
    }
  });
  return Array.from(postUrls);
}

/**
 * Extracts the title and clean text content from a blog post's HTML.
 * @param {string} html - The HTML content of a blog post.
 * @returns {Object|null} An object with title and content, or null.
 */
function cl_getPostContentAndTitle_(html) {
  try {
    const titleMatch = html.match(/<h1[^>]*>([\s\S]*?)<\/h1>/);
    const title = titleMatch ? titleMatch[1].trim().replace(/&nbsp;/g, ' ') : "Title not found";

    const contentMatch = html.match(/<div class="blog-item-content-wrapper"[\s\S]*?>([\s\S]*?)<\/section>/);
    let content = "Content not found";
    if (contentMatch) {
      content = contentMatch[1].replace(/<[^>]*>/g, ' ').replace(/\s\s+/g, ' ').trim();
    }
    
    return { title, content };
  } catch (e) {
    Logger.log(`Failed to get content. Error: ${e.message}`);
    return null;
  }
}

/**
 * Sends content to the Gemini API for analysis and returns a parsed CSV row.
 * @param {string} title - The title of the blog post.
 * @param {string} content - The text content of the blog post.
 * @returns {string[]|null} An array of [title, description, persona].
 */
function cl_analyzeContentWithGemini_(title, content) {
  const apiKey = cfg_('GEMINI_API_KEY'); // Using the cfg_ utility
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${apiKey}`;
  const truncatedContent = content.substring(0, 15000);

  const prompt = `
    You are an expert B2B Content Analyst for Dreamdata, a B2B GTM Attribution Platform.
    TARGET PERSONAS: CMO / VP Marketing, VP Demand Generation, Head of Marketing Ops, Head of Performance Marketing.

    ANALYZE THE FOLLOWING BLOG POST:
    Title: "${title}"
    Content: "${truncatedContent}"

    YOUR TASK:
    1.  Description / Key Takeaway: Write a single, crisp sentence describing the key problem the article solves.
    2.  Target Persona: Identify the primary target persona from the list above.

    OUTPUT FORMAT:
    You MUST respond with a single line of CSV with THREE fields, each enclosed in double quotes: "Title","Description / Key Takeaway","Target Persona".
    The title in your output MUST MATCH the input title exactly.
  `;

  const payload = { 
    "contents": [{ "parts": [{ "text": prompt }] }],
    "generationConfig": { "maxOutputTokens": 512 }
  };
  const options = { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(payload), 'muteHttpExceptions': true };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    const jsonResponse = JSON.parse(responseText);
    
    if (jsonResponse.candidates && jsonResponse.candidates[0].content.parts[0].text) {
      let result = jsonResponse.candidates[0].content.parts[0].text.trim();
      // Simple CSV parser for "field1","field2","field3"
      return result.split('","').map(s => s.replace(/"/g, ''));
    } else {
      Logger.log(`Invalid API response for title "${title}": ${responseText}`);
      return [title, "Analysis Failed: Invalid API response", "N/A"];
    }
  } catch (e) {
    Logger.log(`API call failed for title "${title}". Error: ${e.message}`);
    return [title, `Analysis Failed: ${e.message}`, "N/A"];
  }
}

/**** ===== ADMIN & SETUP ===== ****/

/**
 * A one-time setup function for the script owner to securely store their API keys.
 */
// Replace this entire function
function setupApiKeys() {
  const ui = SpreadsheetApp.getUi();
  const userProperties = PropertiesService.getUserProperties();

  const keysToSet = ['HUBSPOT_TOKEN', 'APOLLO_API_KEY', 'GEMINI_API_KEY'];
  
  for (const key of keysToSet) {
    const response = ui.prompt(`Set API Key`, `Please enter your value for:\n\n${key}`, ui.ButtonSet.OK_CANCEL);
    
    if (response.getSelectedButton() == ui.Button.OK) {
      const value = response.getResponseText().trim();
      if (value) {
        userProperties.setProperty(key, value);
        // --- FIX: Added the required ui.ButtonSet.OK parameter ---
        ui.alert('Success', `The key for "${key}" has been set securely.`, ui.ButtonSet.OK);
      } else {
        // --- FIX: Added the required ui.ButtonSet.OK parameter ---
        ui.alert('Skipped', `No value entered for "${key}". It was not set.`, ui.ButtonSet.OK);
      }
    } else {
      // This alert call was also missing the button parameter.
      ui.alert('Cancelled', 'The setup process was cancelled.', ui.ButtonSet.OK);
      return;
    }
  }
  // --- FIX: Added the required ui.ButtonSet.OK parameter ---
  ui.alert('Setup Complete', 'All API keys have been processed. You can now safely remove them from your "Config" sheet.', ui.ButtonSet.OK);
}