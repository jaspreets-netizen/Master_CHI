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
var DASH_RANGE = '"Dashboard!A1:AZ80"';

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

// One-click full MIRROR sync of the Google Sheet to both ClickUp lists:
//   1) add a task for every active site, 2) push scorecard + history data,
//   3) delete orphan tasks for sites removed from the sheet. Also runs every 24h.
function clickupSyncNow(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  if(!CU_TOKEN){ss.toast('API token not set. Run clickupSaveToken first.','❌');return;}
  var sites;try{sites=getActiveSites_();}catch(e){ss.toast(e.message,'❌');return;}
  var siteKeys=sites.map(function(s){return s.name.toLowerCase();});
  var rm=reportingMonth_();
  ss.toast('Syncing to ClickUp (scorecard + history)...','⏳');
  var addedSet={},removedSet={},scChanges=[],histChanges=[];
  // 1) ensure tasks exist for every active site (a site lands on BOTH lists → count once)
  try{addMissingTasksToList_(findTestingList_()).added.forEach(function(n){addedSet[n]=true;});}catch(e){Logger.log('add scorecard tasks: '+e.message);}
  try{addMissingTasksToList_(findHistoryList_()).added.forEach(function(n){addedSet[n]=true;});}catch(e){Logger.log('add history tasks: '+e.message);}
  // 2) push data (each appends its per-site change lines)
  clickupUpdateScorecard(scChanges);
  clickupUpdateHistory(histChanges);
  // 3) prune orphan tasks (sites removed from the sheet) — true mirror
  try{pruneListToActiveSites_(findTestingList_(),siteKeys).forEach(function(n){removedSet[n]=true;});}catch(e){Logger.log('prune scorecard: '+e.message);}
  try{pruneListToActiveSites_(findHistoryList_(),siteKeys).forEach(function(n){removedSet[n]=true;});}catch(e){Logger.log('prune history: '+e.message);}

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
function impFormula_(sid,dashCol,labels){
  var imp='IMPORTRANGE("https://docs.google.com/spreadsheets/d/'+sid+'",'+DASH_RANGE+')';
  var lookup='ARRAYFORMULA(TRIM(SUBSTITUTE(INDEX('+imp+',0,1),CHAR(160)," ")))';
  var expr='""';
  for(var i=labels.length-1;i>=0;i--){
    expr='IFERROR(INDEX('+imp+',MATCH("'+labels[i]+'",'+lookup+',0),'+dashCol+'),'+expr+')';}
  return '='+expr;
}
function buildMasterActivation(){
  var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss.getSheetByName('Activation');if(sh)ss.deleteSheet(sh);
  sh=ss.insertSheet('Activation',0);
  if(sh.getMaxColumns()<6)sh.insertColumnsAfter(sh.getMaxColumns(),6-sh.getMaxColumns());
  if(sh.getMaxRows()<110)sh.insertRowsAfter(sh.getMaxRows(),110-sh.getMaxRows());
  sh.setColumnWidth(1,30);sh.setColumnWidth(2,180);sh.setColumnWidth(3,160);sh.setColumnWidth(4,500);sh.setColumnWidth(5,200);sh.setColumnWidth(6,200);
  sh.getRange(1,1,sh.getMaxRows(),6).setFontFamily('Arial').setFontSize(11);
  sh.getRange(1,1).setValue('Master CHI Scorecard — Site Registry').setFontWeight('bold').setFontSize(14).setFontColor(CLR.DB);
  sh.getRange(2,1,1,6).setValues([['#','Site Name','CEM Name','CHI Sheet URL','Allow Access','Connection Status']]).setBackground(CLR.DB).setFontColor(CLR.W).setFontWeight('bold');
  sh.getRange(3,1).setValue(1).setHorizontalAlignment('center');
  sh.getRange(3,2).setValue("Dillard's").setBackground(CLR.ME);
  sh.getRange(3,3).setValue('Jaspreet').setBackground(CLR.ME);
  sh.getRange(3,4).setValue('https://docs.google.com/spreadsheets/d/1FtBti0RsDJwuig1z71kdQL6rq6fKRFXUBHD4SF6d_RE/edit').setBackground(CLR.ME).setFontSize(9).setWrap(true);
  sh.getRange(3,5).setFormula('=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1FtBti0RsDJwuig1z71kdQL6rq6fKRFXUBHD4SF6d_RE","Dashboard!A3")');
  sh.getRange(3,6).setFormula('=IF(D3="","",IF(ISTEXT(E3),"✅ Connected","⚠ Click E3 → Allow Access"))');
  sh.setRowHeight(3,30);
  for(var i=1;i<100;i++){var r=3+i;
    sh.getRange(r,1).setValue(i+1).setHorizontalAlignment('center').setFontColor('#ccc');
    sh.getRange(r,2).setBackground(CLR.ME);sh.getRange(r,3).setBackground(CLR.ME);sh.getRange(r,4).setBackground(CLR.ME);
    sh.getRange(r,5).setFormula('=IF(D'+r+'="","",IMPORTRANGE(D'+r+',"Dashboard!A3"))');
    sh.getRange(r,6).setFormula('=IF(D'+r+'="","",IF(ISTEXT(E'+r+'),"✅ Connected","⚠ Click E'+r+' → Allow Access"))');
  }
  sh.getRange(104,1).setValue('Instructions').setFontWeight('bold').setFontColor(CLR.DB);
  sh.getRange(105,1).setValue('1. Enter site name in column B, CEM name in column C, and paste the full Google Sheet URL in column D').setFontSize(10).setFontColor('#666');
  sh.getRange(106,1).setValue('2. Column E will show an IMPORTRANGE formula — click it and grant access when prompted').setFontSize(10).setFontColor('#666');
  sh.getRange(107,1).setValue('3. Column F shows ✅ Connected once access is granted').setFontSize(10).setFontColor('#666');
  sh.getRange(108,1).setValue('4. Run ⚙ Master CHI → Build all trend sheets to generate trend tabs').setFontSize(10).setFontColor('#666');
  sh.getRange(109,1).setValue('5. Values update automatically as site sheets are populated — no refresh needed').setFontSize(10).setFontColor('#666');
  sh.setFrozenRows(2);
  ss.toast('Activation built with 100 site slots + CEM Name column.','Done');
}
function getActiveSites_(){
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Activation');
  if(!sh)throw new Error('Run Build Activation first.');
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
function buildTrendTab_(tabName,labels,columns,titleSuffix,headerColor){
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
    for(var i=0;i<nc;i++)rowFormulas.push(impFormula_(site.sid,columns[i].dashCol,labels));
    if(nc>0)sh.getRange(r,2,1,nc).setFormulas([rowFormulas]).setNumberFormat('0.0').setHorizontalAlignment('center');}
  if(ns>0&&nc>0){var dataRange=sh.getRange(3,2,Math.max(ns,1),nc);
    sh.setConditionalFormatRules(sh.getConditionalFormatRules().concat([
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(5).setBackground('#F4CCCC').setRanges([dataRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(5,6.99).setBackground('#FFF2CC').setRanges([dataRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(7).setBackground('#D9EAD3').setRanges([dataRange]).build()]));}
  sh.setFrozenRows(2);sh.setFrozenColumns(1);
  return ns+' sites × '+nc+' periods';
}
function completeBlockDef_(){
  return [
    {label:'CHI Score',         kind:'chi'},
    {label:'',                  kind:'blank'},
    {label:'Performance Value', kind:'parent',color:CLR.P},
    {label:'Solution KPIs',     kind:'child', color:CLR.P},
    {label:'Uptime',            kind:'child', color:CLR.P},
    {label:'MTBF / MTTR',       kind:'child', color:CLR.P},
    {label:'',                  kind:'blank'},
    {label:'Experience Value',  kind:'parent',color:CLR.E},
    {label:'Frown vs Smile',    kind:'child', color:CLR.E},
    {label:'Sentiment',         kind:'child', color:CLR.E},
    {label:'Trust',             kind:'child', color:CLR.E},
    {label:'',                  kind:'blank'},
    {label:'Business Value',    kind:'parent',color:CLR.B},
    {label:'Throughput Blueprint', kind:'child', color:CLR.B, alts:['Thruput Blueprint']},
    {label:'Outcome Metrics',   kind:'child', color:CLR.B},
    {label:'Move the Needle',   kind:'child', color:CLR.B}
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
      else{for(var i=0;i<nc;i++)row.push(impFormula_(site.sid,mCols[i].dashCol,[block[j].label].concat(block[j].alts||[])));}
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
  r.push('CHI: '+buildTrendTab_('CHI Trend',['CHI Score'],mCols,'CHI Score',CLR.DB));
  r.push('Perf: '+buildTrendTab_('Performance Value Trend',['Performance Value'],mCols,'Performance Value',CLR.P));
  r.push('Exp: '+buildTrendTab_('Experience Value Trend',['Experience Value'],mCols,'Experience Value',CLR.E));
  r.push('Biz: '+buildTrendTab_('Business Value Trend',['Business Value'],mCols,'Business Value',CLR.B));
  r.push('Frown: '+buildTrendTab_('Frown vs Smile Trend',['Frown vs Smile (2.1 score)','Frown Vs Smile Score'],bwCols,'Frown vs Smile',CLR.E));
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
// Smiles'; 'Outcome Metrics' → 'Outcome Metric'. Rag Status is derived from CHI Score
// separately. ('Throughput Blueprint' now matches on both sides after the rename; the
// sheet still tolerates the old 'Thruput Blueprint' spelling via completeBlockDef_ alts.)
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

function clickupUpdateScorecard(report){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  if(!CU_TOKEN){ss.toast('API token not set. Run clickupSaveToken first.','❌');return;}

  ss.toast('Reading Complete CHI Data...','⏳');
  var chiData=readCompleteChiData_();
  if(!chiData.data||Object.keys(chiData.data).length===0){
    ss.toast('No data found. Run buildAllTrends first.','❌');return;}

  var allKeys=Object.keys(chiData.data);

  // CU_FIELD_IDS / CU_RAG_OPTIONS are confirmed against the live list — use them directly.
  var FIELD_IDS=CU_FIELD_IDS, RAG_OPTS=CU_RAG_OPTIONS;

  ss.toast('Fetching tasks...','⏳');
  var list=findTestingList_(),tasks=getAllTasks_(list.id);
  var monthFieldId=resolveFieldIdByName_(list.id,'Month');
  if(!monthFieldId)Logger.log('WARN: "Month" field not found on list — month label will not be written.');
  var ragId=FIELD_IDS['Rag Status'];
  var ragIdxById=dropdownIndexById_(list.id,ragId);   // optionId → orderindex, for change-detection
  Logger.log('Tasks: '+tasks.length+'  Reporting month: '+chiData.monthLabel);

  ss.toast('Updating '+tasks.length+' tasks...','⏳');
  var updated=0,skipped=0;
  for(var t=0;t<tasks.length;t++){
    var task=tasks[t],matchKey=matchSite_(task.name,allKeys);
    if(!matchKey){Logger.log('No match: "'+task.name+'"');skipped++;continue;}
    var d=chiData.data[matchKey]||{};
    var chi=d['CHI Score'];if(chi===undefined||chi==='')chi=null;

    // Snapshot each field's CURRENT raw value, so we only write what actually changed.
    var cur={},cfs=task.custom_fields||[];
    for(var cf=0;cf<cfs.length;cf++){var fv=cfs[cf];cur[fv.id]=fv.value;}

    // Track which fields actually changed for this task.
    var changed=[];
    if(applyField_(task.id,monthFieldId,chiData.monthLabel,cur[monthFieldId]))changed.push('Month');

    // Metric fields — set when present, blank when missing.
    for(var mi=0;mi<METRIC_MAP.length;mi++){
      var fid=FIELD_IDS[METRIC_MAP[mi][0]];if(!fid)continue;
      var val=d[METRIC_MAP[mi][1]];if(val===undefined)val=null;
      if(applyField_(task.id,fid,val,cur[fid]))changed.push(METRIC_MAP[mi][0]);}

    // RAG status derived from CHI (still written to ClickUp via change-detection, but
    // intentionally NOT listed in the change report).
    if(chi!==null){
      var optId=RAG_OPTS[ragKey_(chi)];
      if(optId)applyDropdown_(task.id,ragId,optId,ragIdxById[optId],cur[ragId]);
      else Logger.log('  RAG option not found for: '+ragKey_(chi));
    }else applyDropdown_(task.id,ragId,null,null,cur[ragId]);

    if(changed.length){if(report)report.push(task.name+': '+changed.join(', '));updated++;}
    Logger.log('→ '+task.name+'  CHI='+chi+(changed.length?'  changed: '+changed.join(','):'  (no change)'));
  }
  ss.toast('✅ Scorecard: '+updated+' task(s) changed, '+skipped+' unmatched.\nMonth: '+chiData.monthLabel,'✅ Done');
  Logger.log('=== SCORECARD DONE: '+updated+' changed, '+skipped+' unmatched ===');
}

function addMissingTasksToList_(list){
  var tasks=getAllTasks_(list.id),existing={};
  for(var i=0;i<tasks.length;i++) existing[normHard_(tasks[i].name)]=true;
  var sites=getActiveSites_(),added=[],skipped=0;
  for(var i=0;i<sites.length;i++){
    if(existing[normHard_(sites[i].name)]){skipped++;continue;}
    Utilities.sleep(350);
    try{cuFetch_('POST','/list/'+list.id+'/task',{name:sites[i].name});added.push(sites[i].name);}
    catch(e){Logger.log('Failed to add: '+sites[i].name+' — '+e.message);}
  }
  return {added:added,skipped:skipped};
}
// Mirror deletions: permanently delete any task whose name matches no active site.
// SAFETY: refuses to prune when siteKeys is empty (e.g. Activation sheet unreadable),
// so a bad read can never wipe the whole list. Returns the list of deleted task names.
function pruneListToActiveSites_(list,siteKeys){
  if(!siteKeys||siteKeys.length===0){Logger.log('Prune skipped for list '+list.id+' — no active sites (safety guard).');return [];}
  var tasks=getAllTasks_(list.id),deleted=[];
  for(var i=0;i<tasks.length;i++){
    if(matchSite_(tasks[i].name,siteKeys))continue;   // matches an active site → keep
    Utilities.sleep(200);
    try{cuFetch_('DELETE','/task/'+tasks[i].id);Logger.log('Pruned orphan task: "'+tasks[i].name+'" id='+tasks[i].id);deleted.push(tasks[i].name);}
    catch(e){Logger.log('Prune failed: "'+tasks[i].name+'" — '+e.message);}
  }
  return deleted;
}
function clickupAddMissingTasks(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  if(!CU_TOKEN){ss.toast('API token not set. Run clickupSaveToken first.','❌');return;}
  ss.toast('Adding missing sites to '+CU_LIST_NAME+'...','⏳');
  var list;try{list=findTestingList_();}catch(e){ss.toast(e.message,'❌');return;}
  var r=addMissingTasksToList_(list);
  ss.toast('✅ Added: '+r.added.length+' new tasks\nAlready existed: '+r.skipped,'✅');
  Logger.log('Scorecard tasks — added: '+r.added.length+', skipped: '+r.skipped);
}
function clickupAddMissingHistoryTasks(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  if(!CU_TOKEN){ss.toast('API token not set. Run clickupSaveToken first.','❌');return;}
  ss.toast('Adding missing sites to '+CU_HISTORY_LIST_NAME+'...','⏳');
  var list;try{list=findHistoryList_();}catch(e){ss.toast(e.message,'❌');return;}
  var r=addMissingTasksToList_(list);
  ss.toast('✅ Added: '+r.added.length+' new tasks\nAlready existed: '+r.skipped,'✅');
  Logger.log('History tasks — added: '+r.added.length+', skipped: '+r.skipped);
}

function clickupUpdateHistory(report){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  if(!CU_TOKEN){ss.toast('API token not set. Run clickupSaveToken first.','❌');return;}
  ss.toast('Reading CHI history...','⏳');
  var hist=readChiHistory_();
  if(!hist.data||Object.keys(hist.data).length===0){ss.toast('No data found. Run buildAllTrends first.','❌');return;}
  var allKeys=Object.keys(hist.data);

  ss.toast('Fetching history tasks...','⏳');
  var list;try{list=findHistoryList_();}catch(e){ss.toast(e.message,'❌');return;}
  // Resolve month-field name → id from the live list.
  var fieldMap={},rawFields=(cuFetch_('GET','/list/'+list.id+'/field').fields)||[];
  for(var i=0;i<rawFields.length;i++)if(rawFields[i].name)fieldMap[String(rawFields[i].name).trim()]=rawFields[i].id;
  var missingFields=[];
  for(var m=0;m<hist.months.length;m++)if(!fieldMap[hist.months[m]])missingFields.push(hist.months[m]);
  if(missingFields.length)Logger.log('History fields not found on list (skipped): '+missingFields.join(', '));

  var tasks=getAllTasks_(list.id);
  Logger.log('History tasks: '+tasks.length+'  months: '+hist.months.join(', '));
  ss.toast('Updating '+tasks.length+' history tasks...','⏳');
  var updated=0,skipped=0;
  for(var t=0;t<tasks.length;t++){
    var task=tasks[t],matchKey=matchSite_(task.name,allKeys);
    if(!matchKey){Logger.log('No match: "'+task.name+'"');skipped++;continue;}
    var d=hist.data[matchKey]||{};
    var cur={},cfs=task.custom_fields||[];
    for(var cf=0;cf<cfs.length;cf++){var fv=cfs[cf];cur[fv.id]=fv.value;}
    var changedM=[];
    for(var mm=0;mm<hist.months.length;mm++){
      var fname=hist.months[mm],fid=fieldMap[fname];
      if(!fid)continue;
      var val=d[fname];if(val===undefined)val=null;
      if(applyField_(task.id,fid,val,cur[fid]))changedM.push(fname);
    }
    if(changedM.length){if(report)report.push(task.name+': '+changedM.join(', '));updated++;}
    Logger.log('→ '+task.name+(changedM.length?' history changed: '+changedM.join(','):' history (no change)'));
  }
  ss.toast('✅ History: '+updated+' task(s) changed, '+skipped+' unmatched.','✅ Done');
  Logger.log('=== HISTORY DONE: '+updated+' changed, '+skipped+' unmatched ===');
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
