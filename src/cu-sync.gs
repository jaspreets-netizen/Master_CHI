// ════════════════════════════════════════════════════════
// CLICKUP SYNC v3
// setupFieldsOnTestingList  — adds/checks custom fields, stores IDs
// updateTestingScorecard    — pushes latest complete-month data to ClickUp
// pushToClickUp             — first-run list + task creation
// diagnoseTestingList       — dumps raw API state for debugging
// ════════════════════════════════════════════════════════

// ── Field name registry ──
var FIELD_NAMES = [
  'CEM Name',
  'CHI Score',
  'Performance Value',
  'Experience Value',
  'Business Value',
  'Solution KPIs',
  'Uptime',
  'MTBF / MTTR',
  'Frowns vs Smiles',
  'Sentiment',
  'Trust',
  'Throughput Blueprint',
  'Outcome Metric',
  'Move the Needle',
  'RAG Status'
];

// ── Field ID persistence ──

function loadFieldIds_() {
  var stored = PropertiesService.getScriptProperties().getProperty('testing_field_ids');
  if (!stored) return null;
  return JSON.parse(stored);
}

function saveFieldIds_(fieldIds, ragOptions) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('testing_field_ids',   JSON.stringify(fieldIds));
  props.setProperty('testing_rag_options', JSON.stringify(ragOptions));
}

function loadRagOptions_() {
  var stored = PropertiesService.getScriptProperties().getProperty('testing_rag_options');
  if (!stored) return {};
  return JSON.parse(stored);
}

function clearStoredIds_() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty('testing_list_id');
  props.deleteProperty('testing_field_ids');
  props.deleteProperty('testing_rag_options');
  Logger.log('Cleared stored IDs.');
}

// ── Name normalisation and matching ──

function normHard_(s) {
  return String(s).toLowerCase().replace(/[^a-z0-9]/g, '');
}

function normLight_(s) {
  return String(s).toLowerCase().replace(/['\-\.]/g, '').replace(/\s+/g, ' ').trim();
}

/**
 * Match a ClickUp task name against the keys in our data map.
 * Three passes: exact hard-norm → substring → token overlap.
 */
function matchSite_(taskName, dataKeys) {
  var tn = normHard_(taskName);

  // Pass 1: exact
  for (var i = 0; i < dataKeys.length; i++) {
    if (normHard_(dataKeys[i]) === tn) return dataKeys[i];
  }
  // Pass 2: one contains the other
  for (var i = 0; i < dataKeys.length; i++) {
    var dk = normHard_(dataKeys[i]);
    if (tn.length >= 4 && dk.length >= 4) {
      if (tn.indexOf(dk) >= 0 || dk.indexOf(tn) >= 0) return dataKeys[i];
    }
  }
  // Pass 3: all tokens of the shorter name appear in the longer
  var tnTokens = normLight_(taskName).split(' ');
  for (var i = 0; i < dataKeys.length; i++) {
    var dkTokens  = normLight_(dataKeys[i]).split(' ');
    var shorter   = tnTokens.length <= dkTokens.length ? tnTokens : dkTokens;
    var longer    = tnTokens.length <= dkTokens.length ? dkTokens : tnTokens;
    var longerStr = longer.join(' ');
    var allMatch  = true;
    for (var t = 0; t < shorter.length; t++) {
      if (shorter[t].length >= 3 && longerStr.indexOf(shorter[t]) < 0) { allMatch = false; break; }
    }
    if (allMatch && shorter.length >= 2) return dataKeys[i];
  }
  return null;
}

// ── Helpers ──

function ragKey_(chiScore) {
  if (chiScore >= 7) return 'Green';
  if (chiScore >= 5) return 'Amber';
  return 'Red';
}

function setField_(taskId, fieldId, value) {
  if (!fieldId || value === null || value === undefined || value === '') return;
  try {
    Utilities.sleep(180);
    cuFetch_('POST', '/task/' + taskId + '/field/' + fieldId, { value: value });
  } catch(e) {
    Logger.log('setField failed task=' + taskId + ' field=' + fieldId + ': ' + e.message);
  }
}

// ════════════════════════════════════════════════════════
// SETUP: ADD FIELDS TO EXISTING TESTING LIST
// Run once (or if fields are missing).
// Menu: ⚙ Master CHI → 🔧 Setup fields on Testing list
// ════════════════════════════════════════════════════════

function setupFieldsOnTestingList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (CU_TOKEN === 'paste_your_api_token_here') { ss.toast('Paste your API token first.', '❌'); return; }
  ss.toast('Finding Testing CHI Scorecard list...', '⏳');

  var list;
  try { list = findTestingList_(); } catch(e) { ss.toast(e.message, '❌'); return; }
  var listId = list.id;
  Logger.log('Setting up fields on list: ' + listId);

  // Get existing fields to avoid duplicates
  ss.toast('Checking existing fields...', '⏳');
  var existingFieldIds = {}, existingRagOptions = {};
  try {
    var existing   = cuFetch_('GET', '/list/' + listId + '/field');
    var rawFields  = existing.fields || existing.data || [];
    for (var i = 0; i < rawFields.length; i++) {
      var f = rawFields[i];
      if (f.name) existingFieldIds[f.name] = f.id;
      if (f.name === 'RAG Status' && f.type_config && f.type_config.options) {
        for (var j = 0; j < f.type_config.options.length; j++) {
          existingRagOptions[f.type_config.options[j].name] = f.type_config.options[j].id;
        }
      }
    }
    Logger.log('Existing fields: ' + JSON.stringify(Object.keys(existingFieldIds)));
  } catch(e) {
    Logger.log('Could not read existing fields: ' + e.message);
  }

  var fieldDefs = [
    { name: 'CEM Name',             type: 'text'      },
    { name: 'CHI Score',            type: 'number'    },
    { name: 'Performance Value',    type: 'number'    },
    { name: 'Experience Value',     type: 'number'    },
    { name: 'Business Value',       type: 'number'    },
    { name: 'Solution KPIs',        type: 'number'    },
    { name: 'Uptime',               type: 'number'    },
    { name: 'MTBF / MTTR',          type: 'number'    },
    { name: 'Frowns vs Smiles',     type: 'number'    },
    { name: 'Sentiment',            type: 'number'    },
    { name: 'Trust',                type: 'number'    },
    { name: 'Throughput Blueprint', type: 'number'    },
    { name: 'Outcome Metric',       type: 'number'    },
    { name: 'Move the Needle',      type: 'number'    },
    { name: 'RAG Status',           type: 'drop_down',
      type_config: { options: [
        { name: 'Green', color: '#548235' },
        { name: 'Amber', color: '#BF8F00' },
        { name: 'Red',   color: '#FF0000' }
      ]}
    }
  ];

  var fieldIds = {}, ragOptions = existingRagOptions;
  for (var key in existingFieldIds) fieldIds[key] = existingFieldIds[key];

  var created = 0, skipped = 0, failed = 0;
  ss.toast('Creating missing custom fields...', '⏳');

  for (var i = 0; i < fieldDefs.length; i++) {
    var def = fieldDefs[i];
    if (existingFieldIds[def.name]) {
      Logger.log('Already exists: ' + def.name + ' → ' + existingFieldIds[def.name]);
      skipped++; continue;
    }
    Utilities.sleep(400);
    var payload = { name: def.name, type: def.type };
    if (def.type_config) payload.type_config = def.type_config;
    try {
      var f = cuFetch_('POST', '/list/' + listId + '/field', payload);
      if (f && f.id) {
        fieldIds[def.name] = f.id;
        if (def.name === 'RAG Status' && f.type_config && f.type_config.options) {
          for (var j = 0; j < f.type_config.options.length; j++) {
            ragOptions[f.type_config.options[j].name] = f.type_config.options[j].id;
          }
        }
        Logger.log('Created: ' + def.name + ' → ' + f.id);
        created++;
      } else {
        Logger.log('No ID returned for: ' + def.name + ' → ' + JSON.stringify(f));
        failed++;
      }
    } catch(e) {
      Logger.log('Failed: ' + def.name + ' → ' + e.message);
      failed++;
    }
  }

  saveFieldIds_(fieldIds, ragOptions);
  var summary = [
    '✅ Field setup complete',
    'Created: ' + created,
    'Already existed: ' + skipped,
    'Failed: ' + failed,
    'Field IDs stored in Script Properties.',
    'Run "Update Testing Scorecard" next.'
  ].join('\n');
  ss.toast(summary, '✅');
  Logger.log('Field IDs: '   + JSON.stringify(fieldIds));
  Logger.log('RAG Options: ' + JSON.stringify(ragOptions));
}

// ════════════════════════════════════════════════════════
// MAIN UPDATE
// Menu: ⚙ Master CHI → 🔄 Update Testing Scorecard
// ════════════════════════════════════════════════════════

function updateTestingScorecard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (CU_TOKEN === 'paste_your_api_token_here') { ss.toast('Paste your API token first.', '❌'); return; }

  var fids = loadFieldIds_();
  if (!fids || !fids['CHI Score']) {
    ss.toast('Field IDs not set up. Run "Setup fields on Testing list" first.', '❌');
    return;
  }
  var ragOpts = loadRagOptions_();
  Logger.log('Loaded field IDs: '   + JSON.stringify(fids));
  Logger.log('Loaded RAG options: ' + JSON.stringify(ragOpts));

  ss.toast('Reading scores from trend sheets...', '⏳');
  var chiResult  = readTrendSheetComplete_('CHI Trend');
  var perfResult = readTrendSheetComplete_('Performance Value Trend');
  var expResult  = readTrendSheetComplete_('Experience Value Trend');
  var bizResult  = readTrendSheetComplete_('Business Value Trend');
  var cemMap     = readCemNames_();

  var chiMap  = chiResult.map,  perfMap = perfResult.map;
  var expMap  = expResult.map,  bizMap  = bizResult.map;
  Logger.log('Month — CHI: ' + chiResult.monthLabel + ', Perf: ' + perfResult.monthLabel);
  Logger.log('Sites with CHI data: ' + Object.keys(chiMap).length);
  var allKeys = Object.keys(chiMap);

  ss.toast('Fetching tasks...', '⏳');
  var list  = findTestingList_();
  var tasks = getAllTasks_(list.id);
  Logger.log('Tasks: ' + tasks.length);

  ss.toast('Updating ' + tasks.length + ' tasks...', '⏳');
  var updated = 0, skipped = 0;

  for (var t = 0; t < tasks.length; t++) {
    var task     = tasks[t];
    var matchKey = matchSite_(task.name, allKeys);
    if (!matchKey) {
      Logger.log('No match: "' + task.name + '"');
      skipped++; continue;
    }

    var chi  = chiMap[matchKey]  || null;
    var perf = perfMap[matchKey] || null;
    var exp  = expMap[matchKey]  || null;
    var biz  = bizMap[matchKey]  || null;
    var cem  = cemMap[matchKey]  || null;
    Logger.log('→ ' + task.name + ' [' + matchKey + '] CHI=' + chi);

    setField_(task.id, fids['CHI Score'],          chi);
    setField_(task.id, fids['Performance Value'],   perf);
    setField_(task.id, fids['Experience Value'],    exp);
    setField_(task.id, fids['Business Value'],      biz);
    setField_(task.id, fids['CEM Name'],            cem);

    if (chi !== null && fids['RAG Status'] && ragOpts) {
      var rk    = ragKey_(chi);
      var optId = ragOpts[rk];
      if (optId) setField_(task.id, fids['RAG Status'], optId);
      else Logger.log('  RAG option not found for: ' + rk);
    }
    updated++;
    Utilities.sleep(300);
  }

  ss.toast('✅ ' + updated + ' tasks updated, ' + skipped + ' skipped.\nMonth: ' + chiResult.monthLabel, '✅ Done');
  Logger.log('=== DONE: ' + updated + ' updated, ' + skipped + ' skipped ===');
}

// ════════════════════════════════════════════════════════
// DIAGNOSTIC
// Menu: ⚙ Master CHI → 🔍 Diagnose Testing list fields
// Then check: Apps Script → View → Logs
// ════════════════════════════════════════════════════════

function diagnoseTestingList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (CU_TOKEN === 'paste_your_api_token_here') { ss.toast('Paste your API token first.', '❌'); return; }
  ss.toast('Running diagnostics — check Logs when done...', '⏳');

  var list;
  try { list = findTestingList_(); } catch(e) { ss.toast(e.message, '❌'); return; }

  Logger.log('=== LIST ===');
  Logger.log('Name: ' + list.name + '  ID: ' + list.id);
  Logger.log('=== CUSTOM FIELDS (raw API response) ===');
  Logger.log(JSON.stringify(cuFetch_('GET', '/list/' + list.id + '/field'), null, 2));
  Logger.log('=== STORED FIELD IDs ===');
  Logger.log(JSON.stringify(loadFieldIds_()));
  Logger.log('=== STORED RAG OPTIONS ===');
  Logger.log(JSON.stringify(loadRagOptions_()));
  Logger.log('=== TASK NAMES (first 10) ===');
  var tasks = getAllTasks_(list.id);
  for (var i = 0; i < Math.min(10, tasks.length); i++) {
    Logger.log(tasks[i].name + '  id=' + tasks[i].id);
  }
  ss.toast('Diagnostics complete — open View → Logs to see results.', '✅');
}

// ════════════════════════════════════════════════════════
// INITIAL CREATION (first run only)
// Menu: ⚙ Master CHI → 🔗 Create Testing Scorecard list
// ════════════════════════════════════════════════════════

function pushToClickUp() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (CU_TOKEN === 'paste_your_api_token_here') { ss.toast('Paste your API token first.', '❌'); return; }
  ss.toast('Creating Testing CHI Scorecard list...', '⏳');

  var folder;
  try { folder = findFolder_('Health & Growth'); } catch(e) { ss.toast(e.message, '❌'); return; }

  var list = cuFetch_('POST', '/folder/' + folder.id + '/list', { name: 'Testing CHI Scorecard' });
  if (!list || !list.id) { ss.toast('List creation failed.', '❌'); return; }

  PropertiesService.getScriptProperties().setProperty('testing_list_id', list.id);
  Logger.log('Created list: ' + list.id);

  ss.toast('Creating tasks...', '⏳');
  var sites = getActiveSites_(), created = 0;
  for (var i = 0; i < sites.length; i++) {
    Utilities.sleep(350);
    try { cuFetch_('POST', '/list/' + list.id + '/task', { name: sites[i].name }); created++; }
    catch(e) { Logger.log('Task failed: ' + sites[i].name); }
  }
  ss.toast('✅ List created with ' + created + ' tasks.\nNow run "Setup fields on Testing list".', '✅');
}
