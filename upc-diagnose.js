(async () => {
  const web = (window._spPageContextInfo && window._spPageContextInfo.webAbsoluteUrl) || location.origin;
  const LIST = 'UPC_ProvisioningJobs';
  const api = `${web}/_api/web/lists/getByTitle('${LIST}')`;

  const REQUIRED = [
    'Id', 'Title', 'JobType', 'Status', 'PayloadJson', 'StepsJson', 'ScheduledFor',
    'CorrelationId', 'BatchId', 'TargetUpn', 'TargetUserId', 'ApprovalsJson',
    'RunningInstanceId', 'RunningSince'
  ];

  const get = async (url, label) => {
    const started = performance.now();
    try {
      const res = await fetch(url, { headers: { Accept: 'application/json;odata=nometadata' }, credentials: 'include' });
      const ms = Math.round(performance.now() - started);
      const text = await res.text();
      let body;
      try { body = JSON.parse(text); } catch { body = text.slice(0, 400); }
      console.log(`%c[${label}] ${res.status} in ${ms}ms`, res.ok ? 'color:green' : 'color:red', body);
      return { ok: res.ok, status: res.status, body };
    } catch (e) {
      const ms = Math.round(performance.now() - started);
      console.log(`%c[${label}] NETWORK FAILURE in ${ms}ms`, 'color:red', e);
      return { ok: false, status: 0, body: String(e) };
    }
  };

  console.log('%c=== UPC diagnostic ===', 'font-weight:bold;font-size:14px');
  console.log('web:', web);

  const list = await get(`${api}?$select=Title,ItemCount`, '1. list exists');
  if (!list.ok) {
    console.log('%cSTOP: the list itself is not readable. Everything else will fail.', 'color:red;font-weight:bold');
    return;
  }

  const fields = await get(
    `${api}/fields?$select=InternalName&$top=500`,
    '2. columns'
  );
  if (fields.ok) {
    const present = new Set((fields.body.value || []).map((f) => f.InternalName));
    const missing = REQUIRED.filter((c) => !present.has(c));
    if (missing.length) {
      console.log('%cMISSING COLUMNS: ' + missing.join(', '), 'color:red;font-weight:bold');
      console.log('%c-> Run "Provision lists" from the web part properties to add them.', 'color:orange;font-weight:bold');
    } else {
      console.log('%cAll required columns present.', 'color:green;font-weight:bold');
    }
  }

  await get(`${api}/items?$select=Id&$top=1`, '3. simplest query');

  const summarySelect = 'Id,Title,JobType,Status,CorrelationId,BatchId,TargetUpn,ScheduledFor,Created,Modified,RunningSince,RequestedBy/Title,ApprovedBy/Title';
  await get(
    `${api}/items?$select=${summarySelect}&$expand=RequestedBy,ApprovedBy&$orderby=Id desc&$top=500`,
    '4. EXACT query the dashboard runs'
  );

  await get(`${web}/_api/web/lists/getByTitle('UPC_Roles')/items?$select=Title,MemberGroupId&$top=50`, '5. UPC_Roles (drives your tabs)');

  console.log('%c=== done ===', 'font-weight:bold');
})();
