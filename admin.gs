/**
 * A one-time setup function for the script owner to securely store their API keys.
 */
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
  ui.alert('Setup Complete', 'All API keys have been processed. You can now safely remove them from your "Config" sheet.', ui.ButtonSet.OK);
}