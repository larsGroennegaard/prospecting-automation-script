function generateAiMessages() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const conSh = ss.getSheetByName(CON_SHEET);
  if (conSh.getLastRow() < 2) { ui.alert('No contacts to process.'); return; }

  // --- NEW: Get column indices ---
  const headers = conSh.getRange(1, 1, 1, conSh.getLastColumn()).getValues()[0];
  const assignedSenderNameColIdx = headers.indexOf('assigned_sender_name');
  if (assignedSenderNameColIdx === -1) {
    ui.alert('Error: Could not find the "assigned_sender_name" column.');
    return;
  }
  
  const allContacts = conSh.getRange(2, 1, conSh.getLastRow() - 1, conSh.getLastColumn()).getValues();
  const selectedContacts = allContacts.filter(row => row[0] === true);

  if (selectedContacts.length === 0) {
    ui.alert('No contacts are selected. Please check the boxes in column A for the contacts you want to process.');
    return;
  }
  
  const response = ui.alert('Generate AI Messages?', `This will process the ${selectedContacts.length} selected contacts. Continue?`, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  const accSh = ss.getSheetByName(ACC_SHEET);
  
  const sequencePromptTemplate = cfg_('EMAIL_SEQUENCE_PROMPT');
  const myCompanyName = cfg_('MY_COMPANY_NAME');
  const myValueProp = cfg_('MY_VALUE_PROPOSITION');
  const dreamdataPositioning = (() => { try { return cfg_('DREAMDATA_POSITIONING'); } catch(e) { return 'A B2B Attribution Platform.'; } })();
  const dreamdataUseCases = (() => { try { return cfg_('DREAMDATA_USE_CASES'); } catch(e) { return 'No use cases provided.'; } })();
  const contentLibraryJson = getContentLibrary_();

  const companyData = new Map();
  if (accSh.getLastRow() > 1) {
    const accValues = accSh.getRange(2, 1, accSh.getLastRow() - 1, 8).getValues();
    accValues.forEach(row => {
      const domain = String(row[3]).toLowerCase().trim();
      if (domain) {
        companyData.set(domain, {
          company_name: row[2],
          signals_last_7_days: row[4],
          signals_last_30_days: row[5],
          owner_email: String(row[6]).toLowerCase().trim(),
          account_story: row[7] || 'No specific account journey data available.',
        });
      }
    });
  }

  let processedCount = 0;
  allContacts.forEach((row, index) => {
    const isSelected = row[0];
    const subjectCell = row[12];
    
    // --- NEW: Read the assigned sender name directly from the row ---
    const assignedSenderName = row[assignedSenderNameColIdx];

    if (isSelected && subjectCell === '' && assignedSenderName) { // Only process if a sender has been assigned
      const placeholders = {
        '{contact_name}': row[2],
        '{title}': row[3],
        '{stage}': row[4],
        '{company_domain}': row[1],
        '{contact_story_30_days}': row[8] || 'No specific contact journey data available.',
        '{my_company_name}': myCompanyName,
        '{my_value_proposition}': myValueProp,
        '{email_sender}': assignedSenderName, // <-- USE THE ASSIGNED NAME
        '{content_library}': contentLibraryJson,
        '{dreamdata_positioning}': dreamdataPositioning,
        '{dreamdata_use_cases}': dreamdataUseCases,
      };

      const companyInfo = companyData.get(String(row[1]).toLowerCase().trim()) || {};
      placeholders['{company_name}'] = companyInfo.company_name || row[1];
      placeholders['{signals_last_7_days}'] = companyInfo.signals_last_7_days || 0;
      placeholders['{signals_last_30_days}'] = companyInfo.signals_last_30_days || 0;
      placeholders['{account_story_30_days}'] = companyInfo.account_story || 'No specific account journey data available.';
      
      let finalPrompt = sequencePromptTemplate;
      for (const key in placeholders) {
        finalPrompt = finalPrompt.replace(new RegExp(key, 'g'), placeholders[key]);
      }
      
      const jsonResponseString = geminiGenerate_(finalPrompt);
      
      let outputs = ['Error', 'Error', 'Error', 'Error', 'Error', 'Error'];
      try {
        const jsonMatch = jsonResponseString.match(/\[[\s\S]*\]|{[\s\S]*}/);
        if (!jsonMatch) throw new Error("No JSON found");
        const sequenceArray = JSON.parse(jsonMatch[0]);
        if (Array.isArray(sequenceArray) && sequenceArray.length === 3 && sequenceArray[0].subject) {
          outputs = [
            sequenceArray[0].subject, sequenceArray[0].body,
            sequenceArray[1].subject, sequenceArray[1].body,
            sequenceArray[2].subject, sequenceArray[2].body,
          ];
        } else { outputs[0] = 'Error: Invalid JSON structure'; outputs[1] = jsonResponseString; }
      } catch (e) { outputs[0] = 'Error: Could not parse AI response'; outputs[1] = jsonResponseString; }
      
      conSh.getRange(index + 2, 13, 1, 6).setValues([outputs]);
      processedCount++;
    } else if (isSelected && subjectCell === '' && !assignedSenderName) {
        // Optional: Mark rows that were skipped because they had no sender assigned
        conSh.getRange(index + 2, 12).setValue('skipped_no_sender_assigned');
    }
  });
  ui.alert(`AI sequence generation complete. Processed ${processedCount} selected contacts.`);
}

function geminiGenerate_(prompt) {
  const apiKey = cfg_('GEMINI_API_KEY');
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${apiKey}`;

  const payload = {
    contents: [{
      parts: [{ text: prompt }]
    }],
    generationConfig: {
      temperature: 0.7,
      topK: 1,
      topP: 1,
      maxOutputTokens: 8192,
      // **THE FIX**: Enforce JSON output from the API
      response_mime_type: "application/json",
    },
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseBody = response.getContentText();
    const json = JSON.parse(responseBody);
    
    if (json.candidates && json.candidates[0].content.parts[0].text) {
      return json.candidates[0].content.parts[0].text.trim();
    } else {
      console.error('Gemini API Error: Invalid response structure.', responseBody);
      return `Error: ${json.error ? json.error.message : 'Invalid response'}`;
    }
  } catch (e) {
    console.error('Gemini API call failed.', e);
    return `Error: ${e.message}`;
  }
}