function hsFetchMarkedCompaniesToSheet() {
  const token = cfg_('HUBSPOT_TOKEN');
  const prop  = cfg_('HUBSPOT_COMPANY_PROP');
  const acc   = ensureAccountsHeader_();

  const last = acc.getLastRow();
  const existing = new Map();
  if (last >= 2) {
    const vals = acc.getRange(2, 1, last - 1, 4).getValues();
    vals.forEach((r, i) => {
      const id = String(r[1] || '').trim();
      if (id) existing.set(id, i + 2);
    });
  }

  const headers = { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' };
  const url = 'https://api.hubapi.com/crm/v3/objects/companies/search';
  const baseBody = {
    filterGroups: [{ filters: [{ propertyName: prop, operator: 'EQ', value: 'true' }] }],
    properties: ['name', 'domain', 'signals_last_7_days', 'signals_last_30_days', 'hubspot_owner_id'],
    limit: 100
  };

  let after;
  const allCompanies = [];
  while (true) {
    const body = Object.assign({}, baseBody, after ? { after } : {});
    const res = httpJson_(url, { method: 'post', payload: JSON.stringify(body), headers });
    if (res.results && res.results.length > 0) allCompanies.push(...res.results);
    if (res.paging && res.paging.next && res.paging.next.after) after = res.paging.next.after;
    else break;
  }
  console.log(`Fetched a total of ${allCompanies.length} companies.`);

  const ownerIds = [...new Set(allCompanies.map(c => c.properties.hubspot_owner_id).filter(Boolean))];
  const ownerMap = new Map();
  if (ownerIds.length > 0) {
    console.log(`Found ${ownerIds.length} unique owners. Fetching their emails...`);
    for (const ownerId of ownerIds) {
      try {
        const ownerUrl = `https://api.hubapi.com/crm/v3/owners/${ownerId}`;
        const ownerRes = httpJson_(ownerUrl, { method: 'get', headers: headers });
        if (ownerRes && ownerRes.email) ownerMap.set(ownerRes.id, ownerRes.email);
      } catch (e) {
        console.error(`Could not fetch details for owner ID ${ownerId}: ${e.message}`);
      }
    }
  }

  let syncedCount = 0;
  allCompanies.forEach(c => {
    const props = c.properties || {};
    const id = c.id;
    const name = props.name || '';
    const domain = props.domain || '';
    const signals7 = props.signals_last_7_days || 0;
    const signals30 = props.signals_last_30_days || 0;
    const ownerId = props.hubspot_owner_id || '';
    const ownerEmail = ownerMap.get(ownerId) || '';
    const rowData = [true, id, name, domain, signals7, signals30, ownerEmail];
    if (existing.has(id)) {
      const row = existing.get(id);
      acc.getRange(row, 1, 1, 7).setValues([rowData]);
    } else {
      acc.appendRow(rowData);
    }
    syncedCount++;
  });
  SpreadsheetApp.getActive().toast(`HubSpot companies synced: ${syncedCount}`);
}