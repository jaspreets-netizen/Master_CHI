/**
 * Master CHI Scorecard v2.3 + ClickUp Sync v3
 *
 * API token stored in Apps Script → User Properties (per-user, not shared)
 * Property name: CU_TOKEN  (set via clickupSaveToken)
 *
 * RUN ORDER:
 *   1. Build all trend sheets  (⚙ Master CHI → 📊 Build all trend sheets)
 *   2. Sync to ClickUp         (⚙ Master CHI → 🔄 Sync to ClickUp now, or run
 *                               clickupSyncNow). One call mirrors the sheet to
 *                               BOTH lists: adds new site tasks, pushes scorecard
 *                               + history, and deletes tasks for removed sites.
 *   3. Enable auto-sync        (function dropdown → clickupSetupDailySync; runs
 *                               clickupSyncNow daily at 3 AM New York)
 *
 * TEST lists:       Rough CHI Scorecard / Rough CHI History (ids hardcoded below)
 * PRODUCTION lists: CHI Scorecard / CHI History — hard-blocked via CU_PROD_LIST_IDS
 */

// ═══ CONFIG ═══
var CU_TOKEN     = PropertiesService.getUserProperties().getProperty('CU_TOKEN');
var CU_TEAM_ID   = '90161459573';
var CU_FOLDER_ID = '90169480684';  // Customer Success → Health & Growth
var CU_LIST_NAME = 'Rough CHI Scorecard';   // current name (informational only — targeting is by ID)
var CU_LIST_ID   = '901615509118';   // Scorecard list — NUMERIC id (STABLE across renames)
// CHI History list — month-by-month CHI Score matrix (one number field per calendar month)
var CU_HISTORY_LIST_NAME = 'Rough CHI History';   // current name (informational only)
var CU_HISTORY_LIST_ID   = '901615510712';   // History list — NUMERIC id (STABLE across renames)
// Safety: numeric ids of PRODUCTION lists to NEVER write to. Targeting is by the IDs
// above (which don't change when you rename a list), so renaming the Rough lists is safe.
// These are the real production lists — a hard block against accidental misconfiguration.
var CU_PROD_LIST_IDS = ['901614388515','901614779350'];  // CHI Scorecard (prod), CHI History (prod)

// Field UUIDs for Rough CHI Scorecard — CONFIRMED against the live list 2026-06-22
// (read from /list/901615509118/field; duplicates excluded — primary instances only)
var CU_FIELD_IDS = {
  'CEM Name':             'bb8fb32c-f667-4f72-8aa3-282f6c3aaae1',
  'CHI Score':            '9e077747-0061-43b1-b0f0-790889fff41d',
  'Performance Value':    '82334a33-03cb-4d5f-ad3e-1469faf85aee',
  'Solution KPIs':        '3edd45c5-744e-45f3-bf36-da6f915c6da7',
  'Uptime':               'e95fc07b-a4e1-433f-a9b8-f7d20547aca2',
  'MTBF / MTTR':          '3a2818f2-9de4-4a14-ac92-07984e5f33e5',
  'Experience Value':     '110f56cc-6cd1-4aa0-93ed-8d109b4c17d9',
  'Frowns vs Smiles':     '60c4f5dc-49e4-41a1-9b4c-f2216386d4c0',
  'Sentiment':            '61bb0afc-6197-4f94-849d-8ce57085fdfd',
  'Trust':                '6803010c-7c09-4962-9547-c53989c527f0',
  'Business Value':       '864eb038-86c7-41a8-855e-539b6113c791',
  'Throughput Blueprint': 'ec4cbb39-852a-4ae2-9c7f-fa136534efb9',
  'Outcome Metric':       'b87534cc-30ad-44ab-9fd4-9203244972b0',
  'Move the Needle':      'f911da72-28d9-45bc-b5bb-d38a76b0ed2e',
  'Rag Status':           '0da35124-38e4-4604-b1ac-f1be25d57860'
};
// Dropdown option UUIDs for the Rag Status field — CONFIRMED 2026-06-22
var CU_RAG_OPTIONS = {
  'Green': '0e6151c4-1beb-46a2-9cc0-b909db09f130',
  'Amber': 'ff68b87b-cb58-4677-8276-9a8e654a08bb',
  'Red':   '5c32adab-145a-4394-9a23-33723dff797f'
};

var CLR = {DB:'#1F4E79',W:'#FFFFFF',P:'#2E75B6',E:'#548235',B:'#BF8F00',
  GY:'#EFEFEF',ME:'#DCEEFB',LBL:'#D9D9D9',SHE:'#B6D7A8',EXG:'#E2EFDA',BZG:'#FFF2CC'};
var MN_AB = {2:'Feb',3:'Mar',4:'Apr',5:'May',6:'Jun',7:'Jul',8:'Aug',9:'Sep',10:'Oct',11:'Nov',12:'Dec'};

// ════════════════════════════════════════════════════════
// API TOKEN SETUP (run once — only you can see User Properties)
// ════════════════════════════════════════════════════════
/**
 * HOW TO USE:
 * 1. Replace PASTE_YOUR_TOKEN_HERE below with your actual ClickUp API token
 * 2. Run this function once from the editor
 * 3. Remove the token value (put back PASTE_YOUR_TOKEN_HERE)
 * Token is saved to User Properties — editors of this sheet cannot read it.
 */
function clickupSaveToken() {
  var token = 'PASTE_YOUR_TOKEN_HERE';
  if (token === 'PASTE_YOUR_TOKEN_HERE') {
    Logger.log('Step 1: Replace PASTE_YOUR_TOKEN_HERE with your real token, then run again.');
    return;
  }
  PropertiesService.getUserProperties().setProperty('CU_TOKEN', token);
  Logger.log('Done — token saved to User Properties. Now remove the token from the code (put back PASTE_YOUR_TOKEN_HERE).');
}


// ════════════════════════════════════════════════════════
// MENU
// ════════════════════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚙ Master CHI')
    .addItem('📊 Build all trend sheets', 'buildAllTrends')
    .addItem('🔄 Sync to ClickUp now', 'clickupSyncNow')
    .addToUi();
}

// One-click full MIRROR sync of the Google Sheet to both ClickUp lists. Each site is bound
// to its task by a stable id stored in the hidden CU_LINK_SHEET tab (keyed by sheet id, not
// in the registry), so renames are mirrored in place (no duplicates / lost history):
// 1) link every active site to its task (stored id, else backfill by name, else create),
// 2) rename + push scorecard + history data, 3) delete orphan tasks whose id no site claims.
// The human registry is never written to. Also runs every 24h.
function clickupSyncNow(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  if(!CU_TOKEN){ss.toast('API token not set. Run clickupSaveToken first.','❌');return;}
  var sites;try{sites=getActiveSites_();}catch(e){ss.toast(e.message,'❌');return;}
  var rm=reportingMonth_();
  ss.toast('Syncing to ClickUp (scorecard + history)...','⏳');
  var addedSet={},removedSet={},scChanges=[],histChanges=[];
  // Each list is a true one-way MIRROR keyed by a stable task id stored in the hidden
  // CU_LINK_SHEET tab (keyed by each site's sheet id):
  //   link (use stored id / backfill by name / create) → rename + update fields → prune by id.
  try{
    var scList=findTestingList_();
    var scLink=ensureTaskLinks_(scList,'scorecard',sites);
    scLink.added.forEach(function(n){addedSet[n]=true;});
    clickupUpdateScorecard(scChanges,scList,scLink);
    pruneListById_(scList,scLink.idToSite,scLink.tasks).forEach(function(n){removedSet[n]=true;});
  }catch(e){Logger.log('scorecard sync: '+e.message);}
  try{
    var hList=findHistoryList_();
    var hLink=ensureTaskLinks_(hList,'history',sites);
    hLink.added.forEach(function(n){addedSet[n]=true;});
    clickupUpdateHistory(histChanges,hList,hLink);
    pruneListById_(hList,hLink.idToSite,hLink.tasks).forEach(function(n){removedSet[n]=true;});
  }catch(e){Logger.log('history sync: '+e.message);}

  var sections=[
    {header:'SITES ADDED',                         lines:Object.keys(addedSet)},
    {header:'SITES REMOVED',                       lines:Object.keys(removedSet)},
    {header:'CHI SCORECARD — values changed',      lines:scChanges},
    {header:'CHI SCORECARD HISTORY — values changed',lines:histChanges}
  ];
  var total=sections.reduce(function(s,sec){return s+sec.lines.length;},0);
  ss.toast('✅ Sync complete — '+total+' change(s).','✅');
  showSyncReport_('ClickUp Sync — '+rm.friendly,sections);
}
// Show a copyable summary grouped into clear categories. Uses a modal dialog
// (selectable/copyable); on the background trigger there's no UI, so it just logs.
function showSyncReport_(title,sections){
  var out=[title,''],total=0;
  for(var i=0;i<sections.length;i++){
    var sec=sections[i];
    if(!sec.lines||!sec.lines.length)continue;
    total+=sec.lines.length;
    out.push('▌ '+sec.header+' ('+sec.lines.length+')');
    for(var j=0;j<sec.lines.length;j++)out.push('   • '+sec.lines[j]);
    out.push('');
  }
  out.push(total===0?'No changes — both lists were already up to date.':'Total: '+total+' change(s).');
  var text=out.join('\n');
  Logger.log(text);
  try{
    var safe=text.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
    var html=HtmlService.createHtmlOutput(
      '<div style="font-family:Arial,sans-serif;margin:4px">'+
      '<textarea readonly onclick="this.select()" style="width:100%;height:340px;font-family:monospace;font-size:12px;white-space:pre;box-sizing:border-box">'+safe+'</textarea>'+
      '<p style="font-size:11px;color:#666;margin:6px 2px 0">Click in the box, then Ctrl/Cmd-A to select all and Ctrl/Cmd-C to copy.</p>'+
      '</div>').setWidth(560).setHeight(440);
    SpreadsheetApp.getUi().showModalDialog(html,title);
  }catch(e){/* no UI context (background trigger) — already logged above */}
}

// ════════════════════════════════════════════════════════
// MASTER CHI SCORECARD v2.3 — CORE (do not modify)
// ════════════════════════════════════════════════════════
function extractSheetId_(url){var m=String(url).match(/\/d\/([a-zA-Z0-9_-]+)/);return m?m[1]:'';}
function colLetter_(n){var s='';while(n>0){n--;s=String.fromCharCode(65+(n%26))+s;n=Math.floor(n/26);}return s;}
// Read one fixed Dashboard cell from a site. The Dashboard layout is code-defined, so we
// address the exact cell — column = the month/biweek dashCol (both start Feb 2026 at B),
// row = the metric's fixed Dashboard row — instead of label-matching column A. This makes
// the Master immune to label renames/typos on the site side. A blank cell reads blank
// (= "that month isn't published yet"); only a real access failure shows the ⚠ message.
// IFNA absorbs IMPORTRANGE's transient "Loading…" (#N/A) state shown right after a rebuild,
// so the ⚠ doesn't flicker on already-connected sites; a genuine access failure is #REF!
// (not #N/A) and still falls through to the ⚠ message.
function impFormula_(sid,dashCol,row){
  var cell='Dashboard!'+colLetter_(dashCol)+row;
  var imp='IMPORTRANGE("https://docs.google.com/spreadsheets/d/'+sid+'","'+cell+'")';
  return '=IFERROR(IFNA('+imp+',""),"⚠ no access — grant access in column E")';
}
// The Activation sheet is a hand-maintained registry (# | Site | CEM | URL | Allow | Status,
// plus any manual columns). There is intentionally no code that builds or rebuilds it — that
// would risk wiping live site data — so it is created and edited by hand.
function getActiveSites_(){
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Activation');
  if(!sh)throw new Error('Activation sheet not found — it is maintained manually.');
  // The Activation sheet is human-maintained (# | Site | CEM | URL | Allow | Status, plus any
  // manual columns such as a Salesforce System ID). The Master only needs the site name, CEM
  // and the sheet id parsed from the URL — it does NOT read, write, or depend on the
  // Salesforce column. ClickUp task ids live in the hidden CU_LINK_SHEET tab (never in this
  // sheet), so manual columns here are never touched by the sync.
  var sites=[],vals=sh.getRange(3,1,100,4).getValues();
  for(var i=0;i<vals.length;i++){
    var name=String(vals[i][1]).trim(),cem=String(vals[i][2]).trim(),url=String(vals[i][3]).trim();
    if(name&&url){var sid=extractSheetId_(url);if(sid)sites.push({name:name,cem:cem,sid:sid});}}
  return sites;
}
function getMonthColumns_(){
  var today=new Date(),curYr=today.getFullYear(),curMn=today.getMonth()+1;
  var endMn=curMn+1,endYr=curYr;if(endMn>12){endMn=1;endYr++;}
  var cols=[],yr=2026,mn=2;
  while(yr<endYr||(yr===endYr&&mn<=endMn)){
    if(MN_AB[mn])cols.push({label:MN_AB[mn].toUpperCase()+' '+(yr%100),dashCol:cols.length+2});
    mn++;if(mn>12){mn=1;yr++;}if(mn===1)continue;}
  return cols;
}
function getBiweekColumns_(){
  var defs=[{mn:2,bw:[[5,6],[7,8]]},{mn:3,bw:[[9,10],[11,12],[13,14]]},{mn:4,bw:[[15,16],[17,18]]},
    {mn:5,bw:[[19,20],[21,22]]},{mn:6,bw:[[23,24],[25,26]]},{mn:7,bw:[[27,28],[29,30]]},
    {mn:8,bw:[[31,32],[33,34],[35,36]]},{mn:9,bw:[[37,38],[39,40]]},{mn:10,bw:[[41,42],[43,44]]},
    {mn:11,bw:[[45,46],[47,48]]},{mn:12,bw:[[49,50],[51,52]]}];
  var today=new Date(),curYr=today.getFullYear(),curMn=today.getMonth()+1;
  var endMn=curMn+1,endYr=curYr;if(endMn>12){endMn=1;endYr++;}
  var cols=[],idx=0,yr=2026,mn=2;
  while(yr<endYr||(yr===endYr&&mn<=endMn)){var def=null;for(var i=0;i<defs.length;i++)if(defs[i].mn===mn){def=defs[i];break;}
    if(def){for(var b=0;b<def.bw.length;b++){cols.push({label:'Wk '+def.bw[b][0]+','+def.bw[b][1],dashCol:idx+2});idx++;}}
    mn++;if(mn>12){mn=1;yr++;}if(mn===1)continue;}
  return cols;
}
function buildTrendTab_(tabName,row,columns,titleSuffix,headerColor){
  var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss.getSheetByName(tabName);if(sh)ss.deleteSheet(sh);
  sh=ss.insertSheet(tabName);
  var sites=getActiveSites_(),nc=columns.length,ns=sites.length,totalCols=nc+1;
  if(sh.getMaxColumns()<totalCols)sh.insertColumnsAfter(sh.getMaxColumns(),totalCols-sh.getMaxColumns());
  if(sh.getMaxRows()<ns+25)sh.insertRowsAfter(sh.getMaxRows(),ns+25-sh.getMaxRows());
  sh.getRange(1,1,sh.getMaxRows(),totalCols).setFontFamily('Arial').setFontSize(11);
  sh.setColumnWidth(1,180);for(var c=2;c<=totalCols;c++)sh.setColumnWidth(c,75);
  sh.getRange(1,1).setValue(titleSuffix+' Trend').setFontWeight('bold').setFontSize(13).setFontColor(headerColor);
  sh.getRange(2,1).setValue('Site').setBackground(CLR.DB).setFontColor(CLR.W).setFontWeight('bold');
  for(var i=0;i<nc;i++)sh.getRange(2,2+i).setValue(columns[i].label).setBackground(CLR.DB).setFontColor(CLR.W).setFontWeight('bold').setHorizontalAlignment('center').setFontSize(9);
  for(var s=0;s<ns;s++){var r=3+s,site=sites[s];
    sh.getRange(r,1).setValue(site.name).setFontWeight('bold');
    var rowFormulas=[];
    for(var i=0;i<nc;i++)rowFormulas.push(impFormula_(site.sid,columns[i].dashCol,row));
    if(nc>0)sh.getRange(r,2,1,nc).setFormulas([rowFormulas]).setNumberFormat('0.0').setHorizontalAlignment('center');}
  if(ns>0&&nc>0){var dataRange=sh.getRange(3,2,Math.max(ns,1),nc);
    sh.setConditionalFormatRules(sh.getConditionalFormatRules().concat([
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(5).setBackground('#F4CCCC').setRanges([dataRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(5,6.99).setBackground('#FFF2CC').setRanges([dataRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(7).setBackground('#D9EAD3').setRanges([dataRange]).build()]));}
  sh.setFrozenRows(2);sh.setFrozenColumns(1);
  return ns+' sites × '+nc+' periods';
}
// Each metric carries its FIXED Dashboard row (see the Dashboard layout). Months/biweeks
// are the columns; the row is constant per metric. Blank spacer rows have no row.
function completeBlockDef_(){
  return [
    {label:'CHI Score',         kind:'chi',                row:3},
    {label:'',                  kind:'blank'},
    {label:'Performance Value', kind:'parent',color:CLR.P, row:5},
    {label:'Solution KPIs',     kind:'child', color:CLR.P, row:6},
    {label:'Uptime',            kind:'child', color:CLR.P, row:7},
    {label:'MTBF / MTTR',       kind:'child', color:CLR.P, row:8},
    {label:'',                  kind:'blank'},
    {label:'Experience Value',  kind:'parent',color:CLR.E, row:10},
    {label:'Frown vs Smile',    kind:'child', color:CLR.E, row:11},
    {label:'Sentiment',         kind:'child', color:CLR.E, row:12},
    {label:'Trust',             kind:'child', color:CLR.E, row:13},
    {label:'',                  kind:'blank'},
    {label:'Business Value',    kind:'parent',color:CLR.B, row:15},
    {label:'Throughput Blueprint', kind:'child', color:CLR.B, row:16},
    {label:'Outcome Metrics',   kind:'child', color:CLR.B, row:17},
    {label:'Move the Needle',   kind:'child', color:CLR.B, row:18}
  ];
}
function buildCompleteCHIData_(){
  var ss=SpreadsheetApp.getActiveSpreadsheet(),name='Complete CHI Data';
  var sh=ss.getSheetByName(name);if(sh)ss.deleteSheet(sh);
  sh=ss.insertSheet(name);
  var sites=getActiveSites_(),mCols=getMonthColumns_(),nc=mCols.length,ns=sites.length;
  var block=completeBlockDef_(),bl=block.length,stride=bl+2,totalCols=2+nc;
  var lastRow=2+(ns>0?ns*stride:1);
  if(sh.getMaxColumns()<totalCols)sh.insertColumnsAfter(sh.getMaxColumns(),totalCols-sh.getMaxColumns());
  if(sh.getMaxRows()<lastRow+5)sh.insertRowsAfter(sh.getMaxRows(),lastRow+5-sh.getMaxRows());
  sh.getRange(1,1,sh.getMaxRows(),totalCols).setFontFamily('Arial').setFontSize(11);
  sh.setColumnWidth(1,140);sh.setColumnWidth(2,160);for(var c=3;c<=totalCols;c++)sh.setColumnWidth(c,75);
  sh.getRange(1,1).setValue('Complete CHI Data — Site-wise Monthly Breakdown').setFontWeight('bold').setFontSize(13).setFontColor(CLR.DB);
  var hdr=['Site','Score'];for(var i=0;i<nc;i++)hdr.push(mCols[i].label);
  sh.getRange(2,1,1,totalCols).setValues([hdr]).setBackground(CLR.DB).setFontColor(CLR.W).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(2,1).setHorizontalAlignment('left');
  for(var s=0;s<ns;s++){
    var site=sites[s],top=3+s*stride;
    sh.getRange(top,1,bl,1).merge();
    sh.getRange(top,1).setValue(site.name).setFontWeight('bold').setFontColor(CLR.DB)
      .setVerticalAlignment('middle').setHorizontalAlignment('center').setWrap(true).setBackground(CLR.GY);
    var labelsCol=[],valGrid=[];
    for(var j=0;j<bl;j++){
      labelsCol.push([block[j].label]);
      var row=[];
      if(block[j].kind==='blank'){for(var i=0;i<nc;i++)row.push('');}
      else{for(var i=0;i<nc;i++)row.push(impFormula_(site.sid,mCols[i].dashCol,block[j].row));}
      valGrid.push(row);}
    sh.getRange(top,2,bl,1).setValues(labelsCol);
    if(nc>0){sh.getRange(top,3,bl,nc).setFormulas(valGrid).setNumberFormat('0.0').setHorizontalAlignment('center');}
    for(var j=0;j<bl;j++){var rr=top+j,d=block[j];
      if(d.kind==='chi'){sh.getRange(rr,2).setFontWeight('bold').setFontColor(CLR.DB);}
      else if(d.kind==='parent'){sh.getRange(rr,2).setFontWeight('bold').setFontColor(d.color);}
      else if(d.kind==='child'){sh.getRange(rr,2).setFontColor(d.color).setFontStyle('italic');}}}
  if(ns>0&&nc>0){var dataRange=sh.getRange(3,3,ns*stride,nc);
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(5).setBackground('#F4CCCC').setRanges([dataRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(5,6.99).setBackground('#FFF2CC').setRanges([dataRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(7).setBackground('#D9EAD3').setRanges([dataRange]).build()]);}
  sh.setFrozenRows(2);sh.setFrozenColumns(2);
  return ns+' sites × '+nc+' months';
}
function buildAllTrends(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();ss.toast('Building trend sheets...','⏳');
  var mCols=getMonthColumns_(),bwCols=getBiweekColumns_(),r=[];
  r.push('CHI: '+buildTrendTab_('CHI Trend',3,mCols,'CHI Score',CLR.DB));
  r.push('Perf: '+buildTrendTab_('Performance Value Trend',5,mCols,'Performance Value',CLR.P));
  r.push('Exp: '+buildTrendTab_('Experience Value Trend',10,mCols,'Experience Value',CLR.E));
  r.push('Biz: '+buildTrendTab_('Business Value Trend',15,mCols,'Business Value',CLR.B));
  r.push('Frown: '+buildTrendTab_('Frown vs Smile Trend',22,bwCols,'Frown vs Smile',CLR.E));
  r.push('Complete: '+buildCompleteCHIData_());
  ss.toast(r.join('\n'),'✅ All trend sheets built');
}

// ════════════════════════════════════════════════════════
// CLICKUP API HELPERS
// ════════════════════════════════════════════════════════
function cuFetch_(method,endpoint,payload){
  var options={method:method,headers:{'Authorization':CU_TOKEN,'Content-Type':'application/json'},muteHttpExceptions:true};
  if(payload)options.payload=JSON.stringify(payload);
  var res=UrlFetchApp.fetch('https://api.clickup.com/api/v2'+endpoint,options);
  var code=res.getResponseCode(),text=res.getContentText();
  if(code<200||code>=300)throw new Error('ClickUp '+code+': '+text.substring(0,200));
  return text?JSON.parse(text):{};
}
// Resolve a custom-field id by its exact name (used for the 'Month' field, which
// lives on the list but is not in the hardcoded CU_FIELD_IDS map).
function resolveFieldIdByName_(listId,name){
  try{
    var fields=(cuFetch_('GET','/list/'+listId+'/field').fields)||[];
    for(var i=0;i<fields.length;i++)if(String(fields[i].name).trim()===name)return fields[i].id;
  }catch(e){Logger.log('resolveFieldIdByName_ failed: '+e.message);}
  return null;
}
// Fetch a list by id; HARD-REFUSE only if the id is on the production denylist.
// (ID-based, not name-based, so renaming the Rough lists never breaks or misfires.)
function fetchListGuarded_(listId){
  for(var i=0;i<CU_PROD_LIST_IDS.length;i++)
    if(String(CU_PROD_LIST_IDS[i])===String(listId))
      throw new Error('REFUSING: list '+listId+' is flagged as PRODUCTION (CU_PROD_LIST_IDS).');
  var list=cuFetch_('GET','/list/'+listId);
  if(!list||!list.id)throw new Error('List '+listId+' not found via API.');
  return list;
}
function findTestingList_(){return fetchListGuarded_(CU_LIST_ID);}
function findHistoryList_(){return fetchListGuarded_(CU_HISTORY_LIST_ID);}

function getAllTasks_(listId){
  var tasks=[],page=0;
  while(true){
    var res=cuFetch_('GET','/list/'+listId+'/task?page='+page+'&include_closed=false');
    if(!res.tasks||res.tasks.length===0)break;
    tasks=tasks.concat(res.tasks);if(res.last_page)break;page++;Utilities.sleep(200);}
  return tasks;
}

// ════════════════════════════════════════════════════════
// DATA READER — reads trend sheets, finds last complete month
// ════════════════════════════════════════════════════════
// Reporting month = previous calendar month (today − 1 month), in the project timezone.
//   short    → matches the Complete CHI Data column header, e.g. 'MAY 26'
//   friendly → written into the ClickUp Month field, e.g. 'May 2026'
function reportingMonth_() {
  var ABBR = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var FULL = ['January','February','March','April','May','June','July','August',
              'September','October','November','December'];
  var now = new Date();
  var d = new Date(now.getFullYear(), now.getMonth() - 1, 1); // first day of previous month
  return {
    short:    (ABBR[d.getMonth()] + ' ' + (d.getFullYear() % 100)).toUpperCase(),
    friendly: FULL[d.getMonth()] + ' ' + d.getFullYear()
  };
}

// Reads the hidden 'Complete CHI Data' sheet for the PREVIOUS calendar month only.
// monthLabel is always returned (even if the column is empty/absent) so the ClickUp
// Month field reflects the reporting month; sites with no data return an empty {} so
// their fields get blanked downstream.
function readCompleteChiData_() {
  var rm = reportingMonth_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Complete CHI Data');
  if (!sh) { Logger.log('Complete CHI Data sheet not found — run Build All Trends first.'); return {data:{}, monthLabel: rm.friendly}; }
  var data = sh.getDataRange().getValues();
  if (data.length < 3) return {data:{}, monthLabel: rm.friendly};
  var header = data[1];

  var targetCol = -1;
  for (var c = 2; c < header.length; c++) {
    if (monthHeaderToFriendly_(header[c]) === rm.friendly) { targetCol = c; break; }
  }
  if (targetCol < 0) {
    Logger.log('Complete CHI Data: reporting-month column for "' + rm.friendly + '" not found — board will show blank.');
    return {data:{}, monthLabel: rm.friendly};
  }
  Logger.log('Complete CHI Data → reporting month ' + rm.friendly + ' (column ' + targetCol + ')');

  var result = {}, currentSite = null;
  for (var r = 2; r < data.length; r++) {
    var siteName = String(data[r][0]).trim();
    var label    = String(data[r][1]).trim();
    if (siteName) currentSite = siteName.toLowerCase();
    if (!currentSite || !label) continue;
    if (!result[currentSite]) result[currentSite] = {};   // ensure every site has an entry (so blanks get cleared)
    var numVal = parseFloat(data[r][targetCol]);
    if (!isNaN(numVal) && numVal > 0) result[currentSite][label] = numVal;
  }
  return {data: result, monthLabel: rm.friendly};
}

// Converts a Complete CHI Data column header to the friendly month label 'May 2026'.
// Handles BOTH a date-typed cell (Sheets auto-parses 'MAY 26' into a Date) and the
// literal string 'MAY 26'. Returns null for non-month headers.
function monthHeaderToFriendly_(h){
  var FULL=['January','February','March','April','May','June','July','August',
            'September','October','November','December'];
  if(Object.prototype.toString.call(h)==='[object Date]'){
    if(isNaN(h.getTime()))return null;
    return FULL[h.getMonth()]+' '+h.getFullYear();
  }
  var ABBR=['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  var parts=String(h).toUpperCase().trim().split(/\s+/);
  if(parts.length<2)return null;
  var idx=ABBR.indexOf(parts[0]);if(idx<0)return null;
  var yy=parseInt(parts[1],10);if(isNaN(yy))return null;
  return FULL[idx]+' '+(yy<100?2000+yy:yy);
}

// Reads the CHI Score for every site across ALL month columns of 'Complete CHI Data'.
// Returns { data:{ siteKey:{ 'May 2026':8.12, ... } }, months:['February 2026', ...] }.
function readChiHistory_(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName('Complete CHI Data');
  if(!sh){Logger.log('Complete CHI Data sheet not found — run Build All Trends first.');return {data:{},months:[]};}
  var data=sh.getDataRange().getValues();
  if(data.length<3)return {data:{},months:[]};
  var header=data[1];
  var monthCols=[];
  for(var c=2;c<header.length;c++){
    var fr=monthHeaderToFriendly_(header[c]);   // handles date-typed and string headers
    if(fr)monthCols.push({col:c,friendly:fr});
  }
  var result={},currentSite=null;
  for(var r=2;r<data.length;r++){
    var siteName=String(data[r][0]).trim();
    var label=String(data[r][1]).trim();
    if(siteName)currentSite=siteName.toLowerCase();
    if(!currentSite)continue;
    if(!result[currentSite])result[currentSite]={};
    if(label==='CHI Score'){
      for(var m=0;m<monthCols.length;m++){
        var v=parseFloat(data[r][monthCols[m].col]);
        if(!isNaN(v)&&v>0)result[currentSite][monthCols[m].friendly]=v;
      }
    }
  }
  return {data:result,months:monthCols.map(function(x){return x.friendly;})};
}

// ════════════════════════════════════════════════════════
// CLICKUP SYNC v3
// ════════════════════════════════════════════════════════
// [ ClickUp field name (key in CU_FIELD_IDS) , source label in 'Complete CHI Data' ]
// Note the label differences handled here: sheet 'Frown vs Smile' → field 'Frowns vs
// Smiles'; 'Outcome Metrics' → 'Outcome Metric'. ('Throughput Blueprint' matches on both
// sides.) Source labels are the row labels written into 'Complete CHI Data' column B by
// completeBlockDef_ — they no longer depend on each site's Dashboard wording, since the
// Master now reads Dashboard cells by fixed position. Rag Status is derived from CHI Score.
var METRIC_MAP = [
  ['CHI Score',            'CHI Score'],
  ['Performance Value',    'Performance Value'],
  ['Solution KPIs',        'Solution KPIs'],
  ['Uptime',               'Uptime'],
  ['MTBF / MTTR',          'MTBF / MTTR'],
  ['Experience Value',     'Experience Value'],
  ['Frowns vs Smiles',     'Frown vs Smile'],
  ['Sentiment',            'Sentiment'],
  ['Trust',                'Trust'],
  ['Business Value',       'Business Value'],
  ['Throughput Blueprint', 'Throughput Blueprint'],
  ['Outcome Metric',       'Outcome Metrics'],
  ['Move the Needle',      'Move the Needle']
];

function normHard_(s){return String(s).toLowerCase().replace(/[^a-z0-9]/g,'');}
function normLight_(s){return String(s).toLowerCase().replace(/['\-\.]/g,'').replace(/\s+/g,' ').trim();}
function matchSite_(taskName,dataKeys){
  var tn=normHard_(taskName);
  for(var i=0;i<dataKeys.length;i++){if(normHard_(dataKeys[i])===tn)return dataKeys[i];}
  for(var i=0;i<dataKeys.length;i++){var dk=normHard_(dataKeys[i]);
    if(tn.length>=4&&dk.length>=4&&(tn.indexOf(dk)>=0||dk.indexOf(tn)>=0))return dataKeys[i];}
  var tnTokens=normLight_(taskName).split(' ');
  for(var i=0;i<dataKeys.length;i++){
    var dkTokens=normLight_(dataKeys[i]).split(' ');
    var shorter=tnTokens.length<=dkTokens.length?tnTokens:dkTokens;
    var longerStr=(tnTokens.length<=dkTokens.length?dkTokens:tnTokens).join(' ');
    var allMatch=true;
    for(var t=0;t<shorter.length;t++){if(shorter[t].length>=3&&longerStr.indexOf(shorter[t])<0){allMatch=false;break;}}
    if(allMatch&&shorter.length>=2)return dataKeys[i];}
  return null;
}
function ragKey_(chiScore){return chiScore>=7?'Green':chiScore>=5?'Amber':'Red';}
function setField_(taskId,fieldId,value){
  if(!fieldId||value===null||value===undefined||value==='')return;
  try{Utilities.sleep(150);cuFetch_('POST','/task/'+taskId+'/field/'+fieldId,{value:value});}
  catch(e){Logger.log('setField failed task='+taskId+' field='+fieldId+': '+e.message);}
}
function clearField_(taskId,fieldId){
  if(!fieldId)return;
  try{Utilities.sleep(150);cuFetch_('DELETE','/task/'+taskId+'/field/'+fieldId);}
  catch(e){Logger.log('clearField failed task='+taskId+' field='+fieldId+': '+e.message);}
}
// Loose equality so we can skip rewriting unchanged values (ClickUp returns numbers
// as strings, etc.). Returns true when a and b represent the same value.
function valuesEqual_(a,b){
  if(a===b)return true;
  if(String(a)===String(b))return true;
  var na=parseFloat(a),nb=parseFloat(b);
  return (!isNaN(na)&&!isNaN(nb)&&na===nb);
}
// Mirror one field to a target value, using the task's CURRENT raw value (curRaw) to
// avoid needless API calls: write only when changed, DELETE only when clearing a value
// that exists. This is what keeps routine syncs fast.
function applyField_(taskId,fieldId,value,curRaw){
  if(!fieldId)return null;
  var newEmpty=(value===null||value===undefined||value==='');
  var curEmpty=(curRaw===undefined||curRaw===null||curRaw==='');
  if(newEmpty){if(!curEmpty){clearField_(taskId,fieldId);return 'clear';}return null;}
  if(!curEmpty&&valuesEqual_(curRaw,value))return null;   // unchanged → skip
  setField_(taskId,fieldId,value);
  return curEmpty?'set':'update';
}
// Build a map { optionId → orderindex } for a drop_down field, so we can tell whether
// a task's dropdown already shows the right option (ClickUp stores the value as the
// option's orderindex, not its id).
function dropdownIndexById_(listId,fieldId){
  var map={};
  try{
    var fields=(cuFetch_('GET','/list/'+listId+'/field').fields)||[];
    for(var i=0;i<fields.length;i++){
      var f=fields[i];
      if(f.id===fieldId&&f.type_config&&f.type_config.options){
        var opts=f.type_config.options;
        for(var j=0;j<opts.length;j++)
          map[opts[j].id]=(opts[j].orderindex!==undefined&&opts[j].orderindex!==null)?opts[j].orderindex:j;
      }
    }
  }catch(e){Logger.log('dropdownIndexById_ failed: '+e.message);}
  return map;
}
// Mirror a drop_down field with change-detection. curRaw is the task's current value
// (an orderindex, or sometimes the option id); targetIdx is the orderindex of optionId.
function applyDropdown_(taskId,fieldId,optionId,targetIdx,curRaw){
  if(!fieldId)return null;
  var curEmpty=(curRaw===undefined||curRaw===null||curRaw==='');
  if(optionId===null||optionId===undefined){if(!curEmpty){clearField_(taskId,fieldId);return 'clear';}return null;}
  // Already correct? Match either the option id or its orderindex.
  if(!curEmpty&&(String(curRaw)===String(optionId)||(targetIdx!==undefined&&targetIdx!==null&&Number(curRaw)===Number(targetIdx))))return null;
  setField_(taskId,fieldId,optionId);
  return curEmpty?'set':'update';
}

// Hidden tab that stores the ClickUp task links (one row per site, keyed by sheet id). It is
// the ONLY place task ids are persisted — the human Activation registry is never written to.
var CU_LINK_SHEET = '_ClickUp Sync';   // hidden; "do not edit" bookkeeping for the sync
var CU_LINK_COL = {scorecard:3, history:4};   // role → column in CU_LINK_SHEET (C / D)
// Fetch (creating + hiding if needed) the hidden link tab. Columns:
//   A Site Sheet ID (key) | B Site Name | C Scorecard Task ID | D History Task ID
function getLinkSheet_(){
  var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss.getSheetByName(CU_LINK_SHEET);
  if(!sh){
    sh=ss.insertSheet(CU_LINK_SHEET);
    sh.getRange(1,1,1,4).setValues([['Site Sheet ID','Site Name','Scorecard Task ID','History Task ID']])
      .setFontWeight('bold').setBackground(CLR.DB).setFontColor(CLR.W);
    sh.getRange(1,1).setNote('Auto-managed by the ClickUp sync — do not edit or delete.');
    sh.setFrozenRows(1);
  }
  try{sh.hideSheet();}catch(e){}
  return sh;
}
// Read the link tab into { sid: {row, name, scorecard, history} }.
function readLinkRows_(sh){
  var map={},last=sh.getLastRow();
  if(last<2)return map;
  var vals=sh.getRange(2,1,last-1,4).getValues();
  for(var i=0;i<vals.length;i++){
    var sid=String(vals[i][0]).trim();if(!sid)continue;
    map[sid]={row:2+i,name:String(vals[i][1]).trim(),scorecard:String(vals[i][2]).trim(),history:String(vals[i][3]).trim()};
  }
  return map;
}
// Link every active site to a stable ClickUp task on `list` for `role` ('scorecard'|'history'),
// keyed by the task id stored in the hidden CU_LINK_SHEET (by the site's sheet id). Per site:
//   1) stored id that still exists  → use it directly
//   2) else backfill once by fuzzy NAME-matching an existing (unclaimed) task → store its id
//   3) else create a new task                                                → store its id
// The id is written to the hidden tab (never the registry) so a later site rename never
// orphans the task. Returns { idToSite:{taskId→site}, added:[names], tasks:[all task objs] }.
function ensureTaskLinks_(list,role,sites){
  var sh=getLinkSheet_(),linkMap=readLinkRows_(sh),col=CU_LINK_COL[role];
  var appendAt=sh.getLastRow()+1;
  var tasks=getAllTasks_(list.id),byId={};
  for(var i=0;i<tasks.length;i++)byId[tasks[i].id]=tasks[i];
  var added=[],idToSite={},claimed={};
  // Ensure a row exists for this site's sheet id (append if new), keep its name current.
  function rowFor(site){
    var e=linkMap[site.sid];
    if(!e){sh.getRange(appendAt,1,1,2).setValues([[site.sid,site.name]]);e={row:appendAt,name:site.name,scorecard:'',history:''};linkMap[site.sid]=e;appendAt++;}
    else if(e.name!==site.name){sh.getRange(e.row,2).setValue(site.name);e.name=site.name;}
    return e;
  }
  function store(site,taskId){var e=rowFor(site);sh.getRange(e.row,col).setValue(taskId);e[role]=taskId;idToSite[taskId]=site;claimed[taskId]=true;}
  for(var s=0;s<sites.length;s++){
    var site=sites[s],e=linkMap[site.sid],storedId=e?e[role]:'';
    // 1) stored id still valid → use directly
    if(storedId&&byId[storedId]&&!claimed[storedId]){idToSite[storedId]=site;claimed[storedId]=true;rowFor(site);continue;}
    // 2) backfill: first unclaimed existing task whose name matches this site
    var matched=null;
    for(var ti=0;ti<tasks.length;ti++){
      if(claimed[tasks[ti].id])continue;
      if(matchSite_(tasks[ti].name,[site.name])){matched=tasks[ti];break;}}
    if(matched){store(site,matched.id);continue;}
    // 3) create a new task and store the returned id
    Utilities.sleep(350);
    try{
      var created=cuFetch_('POST','/list/'+list.id+'/task',{name:site.name});
      byId[created.id]=created;tasks.push(created);
      store(site,created.id);added.push(site.name);
    }catch(e2){Logger.log('Failed to create task for "'+site.name+'" on list '+list.id+' — '+e2.message);}
  }
  return {idToSite:idToSite,added:added,tasks:tasks};
}
// Prune by ID: delete any task whose id is claimed by no active site (idToSite). SAFETY:
// refuses to prune when nothing is claimed (e.g. Activation unreadable → zero sites), so a
// bad read can never wipe the list. `tasks` reuses the snapshot from ensureTaskLinks_.
function pruneListById_(list,idToSite,tasks){
  var claimed=idToSite||{};
  if(Object.keys(claimed).length===0){Logger.log('Prune skipped for list '+list.id+' — no claimed tasks (safety guard).');return [];}
  var all=tasks||getAllTasks_(list.id),deleted=[];
  for(var i=0;i<all.length;i++){
    if(claimed[all[i].id])continue;   // claimed by an active site → keep
    Utilities.sleep(200);
    try{cuFetch_('DELETE','/task/'+all[i].id);Logger.log('Pruned orphan task: "'+all[i].name+'" id='+all[i].id);deleted.push(all[i].name);}
    catch(e){Logger.log('Prune failed: "'+all[i].name+'" — '+e.message);}
  }
  return deleted;
}
// Rename a task in place to mirror the current site name (PUT /task/{id}). Change-detected
// so we only call the API when the name actually differs. Returns 'Name' when it renamed.
function renameTaskIfNeeded_(task,taskId,siteName){
  if(!task||String(task.name)===String(siteName))return null;
  try{Utilities.sleep(150);cuFetch_('PUT','/task/'+taskId,{name:siteName});return 'Name';}
  catch(e){Logger.log('rename failed task='+taskId+': '+e.message);return null;}
}

function clickupUpdateScorecard(report,list,link){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  if(!CU_TOKEN){ss.toast('API token not set. Run clickupSaveToken first.','❌');return;}
  // Allow standalone calls (e.g. a legacy trigger): build the list + id-links ourselves.
  if(!list||!link){
    var sites0;try{sites0=getActiveSites_();}catch(e){ss.toast(e.message,'❌');return;}
    try{list=findTestingList_();}catch(e){ss.toast(e.message,'❌');return;}
    link=ensureTaskLinks_(list,'scorecard',sites0);
  }

  ss.toast('Reading Complete CHI Data...','⏳');
  var chiData=readCompleteChiData_();
  if(!chiData.data||Object.keys(chiData.data).length===0){
    ss.toast('No data found. Run buildAllTrends first.','❌');return;}

  // CU_FIELD_IDS / CU_RAG_OPTIONS are confirmed against the live list — use them directly.
  var FIELD_IDS=CU_FIELD_IDS, RAG_OPTS=CU_RAG_OPTIONS;
  var monthFieldId=resolveFieldIdByName_(list.id,'Month');
  if(!monthFieldId)Logger.log('WARN: "Month" field not found on list — month label will not be written.');
  var ragId=FIELD_IDS['Rag Status'];
  var ragIdxById=dropdownIndexById_(list.id,ragId);   // optionId → orderindex, for change-detection

  var byId={};for(var i=0;i<link.tasks.length;i++)byId[link.tasks[i].id]=link.tasks[i];
  var ids=Object.keys(link.idToSite);
  Logger.log('Scorecard tasks linked: '+ids.length+'  Reporting month: '+chiData.monthLabel);
  ss.toast('Updating '+ids.length+' tasks...','⏳');
  var updated=0;
  for(var k=0;k<ids.length;k++){
    var taskId=ids[k],site=link.idToSite[taskId],task=byId[taskId];
    var d=chiData.data[site.name.trim().toLowerCase()]||{};   // data key = site name, lowercased
    var chi=d['CHI Score'];if(chi===undefined||chi==='')chi=null;

    // Snapshot each field's CURRENT raw value, so we only write what actually changed.
    var cur={},cfs=(task&&task.custom_fields)||[];
    for(var cf=0;cf<cfs.length;cf++){var fv=cfs[cf];cur[fv.id]=fv.value;}

    // Track what actually changed for this task (incl. an in-place rename).
    var changed=[];
    var renamed=renameTaskIfNeeded_(task,taskId,site.name);if(renamed)changed.push(renamed);
    if(applyField_(taskId,monthFieldId,chiData.monthLabel,cur[monthFieldId]))changed.push('Month');

    // Metric fields — set when present, blank when missing.
    for(var mi=0;mi<METRIC_MAP.length;mi++){
      var fid=FIELD_IDS[METRIC_MAP[mi][0]];if(!fid)continue;
      var val=d[METRIC_MAP[mi][1]];if(val===undefined)val=null;
      if(applyField_(taskId,fid,val,cur[fid]))changed.push(METRIC_MAP[mi][0]);}

    // RAG status derived from CHI (still written to ClickUp via change-detection, but
    // intentionally NOT listed in the change report).
    if(chi!==null){
      var optId=RAG_OPTS[ragKey_(chi)];
      if(optId)applyDropdown_(taskId,ragId,optId,ragIdxById[optId],cur[ragId]);
      else Logger.log('  RAG option not found for: '+ragKey_(chi));
    }else applyDropdown_(taskId,ragId,null,null,cur[ragId]);

    if(changed.length){if(report)report.push(site.name+': '+changed.join(', '));updated++;}
    Logger.log('→ '+site.name+'  CHI='+chi+(changed.length?'  changed: '+changed.join(','):'  (no change)'));
  }
  ss.toast('✅ Scorecard: '+updated+' task(s) changed.\nMonth: '+chiData.monthLabel,'✅ Done');
  Logger.log('=== SCORECARD DONE: '+updated+' changed ===');
}

function clickupAddMissingTasks(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  if(!CU_TOKEN){ss.toast('API token not set. Run clickupSaveToken first.','❌');return;}
  ss.toast('Linking sites to '+CU_LIST_NAME+'...','⏳');
  var list,sites;try{list=findTestingList_();sites=getActiveSites_();}catch(e){ss.toast(e.message,'❌');return;}
  var r=ensureTaskLinks_(list,'scorecard',sites);
  ss.toast('✅ Added: '+r.added.length+' new tasks\nLinked total: '+Object.keys(r.idToSite).length,'✅');
  Logger.log('Scorecard tasks — added: '+r.added.length+', linked: '+Object.keys(r.idToSite).length);
}
function clickupAddMissingHistoryTasks(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  if(!CU_TOKEN){ss.toast('API token not set. Run clickupSaveToken first.','❌');return;}
  ss.toast('Linking sites to '+CU_HISTORY_LIST_NAME+'...','⏳');
  var list,sites;try{list=findHistoryList_();sites=getActiveSites_();}catch(e){ss.toast(e.message,'❌');return;}
  var r=ensureTaskLinks_(list,'history',sites);
  ss.toast('✅ Added: '+r.added.length+' new tasks\nLinked total: '+Object.keys(r.idToSite).length,'✅');
  Logger.log('History tasks — added: '+r.added.length+', linked: '+Object.keys(r.idToSite).length);
}

function clickupUpdateHistory(report,list,link){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  if(!CU_TOKEN){ss.toast('API token not set. Run clickupSaveToken first.','❌');return;}
  // Allow standalone calls: build the list + id-links ourselves.
  if(!list||!link){
    var sites0;try{sites0=getActiveSites_();}catch(e){ss.toast(e.message,'❌');return;}
    try{list=findHistoryList_();}catch(e){ss.toast(e.message,'❌');return;}
    link=ensureTaskLinks_(list,'history',sites0);
  }
  ss.toast('Reading CHI history...','⏳');
  var hist=readChiHistory_();
  if(!hist.data||Object.keys(hist.data).length===0){ss.toast('No data found. Run buildAllTrends first.','❌');return;}

  // Resolve month-field name → id from the live list.
  var fieldMap={},rawFields=(cuFetch_('GET','/list/'+list.id+'/field').fields)||[];
  for(var i=0;i<rawFields.length;i++)if(rawFields[i].name)fieldMap[String(rawFields[i].name).trim()]=rawFields[i].id;
  var missingFields=[];
  for(var m=0;m<hist.months.length;m++)if(!fieldMap[hist.months[m]])missingFields.push(hist.months[m]);
  if(missingFields.length)Logger.log('History fields not found on list (skipped): '+missingFields.join(', '));

  var byId={};for(var i=0;i<link.tasks.length;i++)byId[link.tasks[i].id]=link.tasks[i];
  var ids=Object.keys(link.idToSite);
  Logger.log('History tasks linked: '+ids.length+'  months: '+hist.months.join(', '));
  ss.toast('Updating '+ids.length+' history tasks...','⏳');
  var updated=0;
  for(var k=0;k<ids.length;k++){
    var taskId=ids[k],site=link.idToSite[taskId],task=byId[taskId];
    var d=hist.data[site.name.trim().toLowerCase()]||{};
    var cur={},cfs=(task&&task.custom_fields)||[];
    for(var cf=0;cf<cfs.length;cf++){var fv=cfs[cf];cur[fv.id]=fv.value;}
    var changedM=[];
    var renamed=renameTaskIfNeeded_(task,taskId,site.name);if(renamed)changedM.push(renamed);
    for(var mm=0;mm<hist.months.length;mm++){
      var fname=hist.months[mm],fid=fieldMap[fname];
      if(!fid)continue;
      var val=d[fname];if(val===undefined)val=null;
      if(applyField_(taskId,fid,val,cur[fid]))changedM.push(fname);
    }
    if(changedM.length){if(report)report.push(site.name+': '+changedM.join(', '));updated++;}
    Logger.log('→ '+site.name+(changedM.length?' history changed: '+changedM.join(','):' history (no change)'));
  }
  ss.toast('✅ History: '+updated+' task(s) changed.','✅ Done');
  Logger.log('=== HISTORY DONE: '+updated+' changed ===');
}

function clickupDiagnose(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  if(!CU_TOKEN){ss.toast('API token not set. Run clickupSaveToken first.','❌');return;}
  ss.toast('Running diagnostics...','⏳');
  var list;try{list=findTestingList_();}catch(e){ss.toast(e.message,'❌');return;}
  Logger.log('=== LIST ===\nName: '+list.name+'  ID: '+list.id);
  Logger.log('=== CUSTOM FIELDS ===\n'+JSON.stringify(cuFetch_('GET','/list/'+list.id+'/field'),null,2));
  Logger.log('=== TASK NAMES (first 10) ===');
  var tasks=getAllTasks_(list.id);
  for(var i=0;i<Math.min(10,tasks.length);i++)Logger.log(tasks[i].name+'  id='+tasks[i].id);
  ss.toast('Diagnostics complete — open View → Logs.','✅');
}

// ════════════════════════════════════════════════════════
// 24-HOUR AUTO-SYNC
// ════════════════════════════════════════════════════════
var DAILY_SYNC_HANDLERS=['clickupDailySync','clickupUpdateScorecard'];  // current + legacy
// Daily trigger target — same full sync as the menu's "Sync to ClickUp now".
function clickupDailySync(){clickupSyncNow();}
function clickupSetupDailySync(){
  var triggers=ScriptApp.getProjectTriggers();
  for(var i=0;i<triggers.length;i++)
    if(DAILY_SYNC_HANDLERS.indexOf(triggers[i].getHandlerFunction())>=0)ScriptApp.deleteTrigger(triggers[i]);
  ScriptApp.newTrigger('clickupDailySync').timeBased().everyDays(1).atHour(3).create();
  SpreadsheetApp.getActiveSpreadsheet().toast('✅ Auto-sync active — scorecard + history daily at 3 AM (New York).','✅');
}
function clickupStopDailySync(){
  var triggers=ScriptApp.getProjectTriggers(),removed=0;
  for(var i=0;i<triggers.length;i++)
    if(DAILY_SYNC_HANDLERS.indexOf(triggers[i].getHandlerFunction())>=0){ScriptApp.deleteTrigger(triggers[i]);removed++;}
  SpreadsheetApp.getActiveSpreadsheet().toast(removed>0?'✅ Auto-sync stopped.':'No trigger found.',removed>0?'✅':'ℹ️');
}
