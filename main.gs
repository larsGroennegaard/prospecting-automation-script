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
      .addItem('Set API Keys', 'setupApiKeys')
      .addItem('Rebuild Content Library', 'cl_buildLibrary'); 


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