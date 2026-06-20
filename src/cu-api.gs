// ════════════════════════════════════════════════════════
// CLICKUP API HELPERS
// ════════════════════════════════════════════════════════

function cuFetch_(method, endpoint, payload) {
  var options = {
    method: method,
    headers: { 'Authorization': CU_TOKEN, 'Content-Type': 'application/json' },
    muteHttpExceptions: true
  };
  if (payload) options.payload = JSON.stringify(payload);
  var res  = UrlFetchApp.fetch('https://api.clickup.com/api/v2' + endpoint, options);
  var code = res.getResponseCode();
  var text = res.getContentText();
  if (code < 200 || code >= 300) throw new Error('ClickUp ' + code + ': ' + text.substring(0, 200));
  return JSON.parse(text);
}

function findFolder_(folderName) {
  var spaces = cuFetch_('GET', '/team/' + CU_TEAM_ID + '/space?archived=false').spaces;
  for (var i = 0; i < spaces.length; i++) {
    var folders = cuFetch_('GET', '/space/' + spaces[i].id + '/folder?archived=false').folders;
    for (var j = 0; j < folders.length; j++) {
      if (folders[j].name === folderName) return folders[j];
    }
  }
  throw new Error('Folder "' + folderName + '" not found.');
}

function findTestingList_() {
  var props    = PropertiesService.getScriptProperties();
  var storedId = props.getProperty('testing_list_id');
  if (storedId) {
    try {
      var list = cuFetch_('GET', '/list/' + storedId);
      if (list && list.id) return list;
    } catch(e) { /* list may have been deleted, fall through */ }
  }
  var folder = findFolder_('Health & Growth');
  var lists  = cuFetch_('GET', '/folder/' + folder.id + '/list').lists;
  for (var i = 0; i < lists.length; i++) {
    if (lists[i].name === 'Testing CHI Scorecard') {
      props.setProperty('testing_list_id', lists[i].id);
      return lists[i];
    }
  }
  throw new Error('"Testing CHI Scorecard" list not found. Run "Create Testing Scorecard" first.');
}

function getAllTasks_(listId) {
  var tasks = [], page = 0;
  while (true) {
    var res = cuFetch_('GET', '/list/' + listId + '/task?page=' + page + '&include_closed=false');
    if (!res.tasks || res.tasks.length === 0) break;
    tasks = tasks.concat(res.tasks);
    if (res.last_page) break;
    page++;
    Utilities.sleep(200);
  }
  return tasks;
}
