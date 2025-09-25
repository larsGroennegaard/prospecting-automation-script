/**** ===== SHEET NAMES ===== ****/
const CFG_SHEET = 'Config';
const ACC_SHEET = 'Accounts';
const CON_SHEET = 'Contacts';
const MAILBOX_MAP_SHEET = 'Mailbox Mapping';
const CONTENT_LIB_SHEET = 'Content Library'; // Add this line


function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  const adminMenu = ui.createMenu('Admin')
      .addItem('Set API Keys', 'setupApiKeys')
      .addItem('Rebuild Content Library', 'cl_buildLibrary');

ui.createMenu('Prospecting')
    .addItem('1) Fetch HS Companies (marked)', 'hsFetchMarkedCompaniesToSheet')
    .addItem('2) Find Contacts in Apollo', 'apolloFindContactsForAccounts')
    .addItem('2.5) Assign Senders for Rotation', 'assignSendersForRotation')
    .addSeparator()
    .addItem('Enrich: Get Account Stories from BQ', 'enrichFromDataWarehouse')
    .addItem('Enrich: Get Contact Journeys from BQ', 'enrichContactsFromBigQuery')
    .addItem('Enrich: Get Contact Bios from Apollo', 'enrichContactsFromApollo') // <-- NEW ITEM
    .addSeparator()
    .addItem('3) Generate AI Messages', 'generateAiMessages')
    .addItem('4) Push AI Messages to Apollo', 'apolloPushMessages')
    .addItem('5) Add Contacts to Apollo Sequence', 'apolloAddContactsToSequence')
    .addSeparator()
    .addSubMenu(adminMenu)
    .addToUi();
}


