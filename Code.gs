/**
 * Master CHI Scorecard v2.3 + ClickUp Sync v3
 *
 * API token stored in Apps Script → User Properties (per-user, not shared)
 * Property name: CU_TOKEN  (set via clickupSaveToken)
 *
 * RUN ORDER (first time):
 *   1. Build all trend sheets  (⚙ Master CHI → 📊 Build all trend sheets)
 *   2. Add missing tasks       (function dropdown → clickupAddMissingTasks)
 *   3. Sync data               (function dropdown → clickupUpdateScorecard)
 *   4. Enable auto-sync        (function dropdown → clickupSetupDailySync)
 *
 * TEST list: Rough CHI Scorecard (numeric ID hardcoded below)
 * PRODUCTION list: CHI Scorecard — NEVER TOUCH
 */

// ═══ CONFIG ═══
var CU_TOKEN     = PropertiesService.getUserProperties().getProperty('CU_TOKEN');
var CU_TEAM_ID   = '90161459573';
var CU_FOLDER_ID = '90169480684';  // Customer Success → Health & Growth
var CU_LIST_NAME = 'Rough CHI Scorecard';   // TEST list — NEVER touch 'CHI Scorecard' (production)
var CU_LIST_ID   = '901615509118';   // Rough CHI Scorecard — NUMERIC list id (from URL /v/l/li/901615509118)
var CU_PROD_LIST_NAME = 'CHI Scorecard';   // PRODUCTION — findTestingList_ HARD-REFUSES to operate on this

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
    .addToUi();
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
    {label:'Thruput Blueprint', kind:'child', color:CLR.B},
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
      else{for(var i=0;i<nc;i++)row.push(impFormula_(site.sid,mCols[i].dashCol,[block[j].label]));}
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
  sh.hideSheet();
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
  return JSON.parse(text);
}
function findTestingList_(){
  var list=cuFetch_('GET','/list/'+CU_LIST_ID);
  if(!list||!list.id)throw new Error('List '+CU_LIST_ID+' ('+CU_LIST_NAME+') not found via API.');
  // HARD GUARD: never operate on the production list, no matter what id is configured.
  var resolvedName=String(list.name||'').trim().toLowerCase();
  if(resolvedName===CU_PROD_LIST_NAME.toLowerCase())
    throw new Error('REFUSING: list '+list.id+' is the PRODUCTION list "'+list.name+'". '+
      'Point CU_LIST_ID at the test list "'+CU_LIST_NAME+'" before running.');
  return list;
}
function clickupShowStructure(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  if(!CU_TOKEN){ss.toast('API token not set. Run clickupSaveToken first.','❌');return;}
  var spaces=cuFetch_('GET','/team/'+CU_TEAM_ID+'/space?archived=false').spaces;
  Logger.log('=== CLICKUP WORKSPACE STRUCTURE ===');
  for(var i=0;i<spaces.length;i++){
    var sp=spaces[i];
    Logger.log('SPACE: "'+sp.name+'"  id='+sp.id);
    var folders=cuFetch_('GET','/space/'+sp.id+'/folder?archived=false').folders;
    for(var j=0;j<folders.length;j++){
      var fo=folders[j];
      Logger.log('  FOLDER: "'+fo.name+'"  id='+fo.id);
      var lists=cuFetch_('GET','/folder/'+fo.id+'/list').lists;
      for(var k=0;k<lists.length;k++)
        Logger.log('    LIST: "'+lists[k].name+'"  id='+lists[k].id);
    }
    var spaceLists=cuFetch_('GET','/space/'+sp.id+'/list?archived=false').lists||[];
    for(var j=0;j<spaceLists.length;j++)
      Logger.log('  LIST (no folder): "'+spaceLists[j].name+'"  id='+spaceLists[j].id);
  }
  Logger.log('=== END ===');
  ss.toast('Done — open Apps Script → View → Logs','✅');
}

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
function readTrendSheetComplete_(sheetName){
  var MIN_COVERAGE=0.5;
  var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss.getSheetByName(sheetName);
  if(!sh){Logger.log('Sheet not found: '+sheetName);return{map:{},monthLabel:''};}
  var data=sh.getDataRange().getValues();
  if(data.length<3)return{map:{},monthLabel:''};
  var siteCount=0;
  for(var r=2;r<data.length;r++){if(String(data[r][0]).trim())siteCount++;}
  if(siteCount===0)return{map:{},monthLabel:''};
  var bestCol=-1,bestLabel='';
  for(var c=data[1].length-1;c>=2;c--){
    var goodCount=0;
    for(var r=2;r<data.length;r++){
      if(!String(data[r][0]).trim())continue;
      var v=parseFloat(data[r][c]);if(!isNaN(v)&&v>=1)goodCount++;}
    if(goodCount/siteCount>=MIN_COVERAGE){bestCol=c;bestLabel=String(data[1][c]);break;}}
  if(bestCol<0){
    Logger.log(sheetName+': no complete month found, using latest with any data');
    for(var c=data[1].length-1;c>=2;c--){
      for(var r=2;r<data.length;r++){var v=parseFloat(data[r][c]);if(!isNaN(v)&&v>0){bestCol=c;bestLabel=String(data[1][c]);break;}}
      if(bestCol>=0)break;}}
  Logger.log(sheetName+' → using column '+bestCol+' ('+bestLabel+')');
  var result={};
  for(var r=2;r<data.length;r++){
    var name=String(data[r][0]).trim();if(!name)continue;
    var v=parseFloat(data[r][bestCol]);if(!isNaN(v)&&v>0)result[name.toLowerCase()]=v;}
  return{map:result,monthLabel:bestLabel};
}
function readCemNames_(){
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Activation');
  if(!sh)return{};
  var data=sh.getRange(3,1,100,3).getValues(),result={};
  for(var i=0;i<data.length;i++){
    var name=String(data[i][1]).trim(),cem=String(data[i][2]).trim();
    if(name)result[name.toLowerCase()]=cem;}
  return result;
}

function readCompleteChiData_() {
  var MIN_COVERAGE = 0.5;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Complete CHI Data');
  if (!sh) { Logger.log('Complete CHI Data sheet not found — run Build All Trends first.'); return {data:{}, monthLabel:''}; }
  var data = sh.getDataRange().getValues();
  if (data.length < 3) return {data:{}, monthLabel:''};
  var header = data[1];

  var chiRows = [];
  for (var r = 2; r < data.length; r++) {
    if (String(data[r][1]).trim() === 'CHI Score') chiRows.push(r);
  }
  var bestCol = -1, bestLabel = '';
  for (var c = header.length - 1; c >= 2; c--) {
    var good = 0;
    for (var ri = 0; ri < chiRows.length; ri++) {
      var v = parseFloat(data[chiRows[ri]][c]);
      if (!isNaN(v) && v >= 1) good++;
    }
    if (chiRows.length > 0 && good / chiRows.length >= MIN_COVERAGE) {
      bestCol = c; bestLabel = String(header[c]); break;
    }
  }
  if (bestCol < 0) {
    for (var c = header.length - 1; c >= 2; c--) {
      if (chiRows.length > 0) {
        var v = parseFloat(data[chiRows[0]][c]);
        if (!isNaN(v) && v > 0) { bestCol = c; bestLabel = String(header[c]); break; }
      }
    }
  }
  if (bestCol < 0) return {data:{}, monthLabel:''};
  Logger.log('Complete CHI Data → column ' + bestCol + ' (' + bestLabel + ')');

  var result = {}, currentSite = null;
  for (var r = 2; r < data.length; r++) {
    var siteName = String(data[r][0]).trim();
    var label    = String(data[r][1]).trim();
    if (siteName) currentSite = siteName.toLowerCase();
    if (!currentSite || !label) continue;
    if (!result[currentSite]) result[currentSite] = {};
    var numVal = parseFloat(data[r][bestCol]);
    if (!isNaN(numVal) && numVal > 0) result[currentSite][label] = numVal;
  }
  return {data: result, monthLabel: bestLabel};
}

// ════════════════════════════════════════════════════════
// CLICKUP SYNC v3
// ════════════════════════════════════════════════════════
var FIELD_NAMES = [
  'Rag Status',
  'Performance Value',
  'Solution KPIs',
  'MTBF / MTTR',
  'Frowns vs Smiles',
  'Sentiment',
  'Trust',
  'Throughput Blueprint',
  'Outcome Metric',
  'Move the Needle'
];

function loadFieldIds_(){
  var stored=PropertiesService.getScriptProperties().getProperty('testing_field_ids');
  return stored?JSON.parse(stored):null;
}
function saveFieldIds_(fieldIds,ragOptions){
  var props=PropertiesService.getScriptProperties();
  props.setProperty('testing_field_ids',JSON.stringify(fieldIds));
  props.setProperty('testing_rag_options',JSON.stringify(ragOptions));
}
function loadRagOptions_(){
  var stored=PropertiesService.getScriptProperties().getProperty('testing_rag_options');
  return stored?JSON.parse(stored):{};
}
function clearStoredIds_(){
  var props=PropertiesService.getScriptProperties();
  props.deleteProperty('testing_list_id');props.deleteProperty('testing_field_ids');props.deleteProperty('testing_rag_options');
  Logger.log('Cleared stored IDs.');
}
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
  try{Utilities.sleep(180);cuFetch_('POST','/task/'+taskId+'/field/'+fieldId,{value:value});}
  catch(e){Logger.log('setField failed task='+taskId+' field='+fieldId+': '+e.message);}
}

function clickupSetupFields(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  if(!CU_TOKEN){ss.toast('API token not set. Run clickupSaveToken first.','❌');return;}
  ss.toast('Finding '+CU_LIST_NAME+' list...','⏳');
  var list;try{list=findTestingList_();}catch(e){ss.toast(e.message,'❌');return;}
  var listId=list.id;
  Logger.log('Setting up fields on list: '+listId);
  ss.toast('Checking existing fields...','⏳');
  var existingFieldIds={},existingRagOptions={};
  try{
    var rawFields=(cuFetch_('GET','/list/'+listId+'/field').fields)||[];
    for(var i=0;i<rawFields.length;i++){
      var f=rawFields[i];if(f.name)existingFieldIds[f.name]=f.id;
      if(f.name==='Rag Status'&&f.type_config&&f.type_config.options){
        for(var j=0;j<f.type_config.options.length;j++)
          existingRagOptions[f.type_config.options[j].name]=f.type_config.options[j].id;}}
    Logger.log('Existing fields: '+JSON.stringify(Object.keys(existingFieldIds)));
  }catch(e){Logger.log('Could not read existing fields: '+e.message);}
  var fieldDefs=[
    {name:'CEM Name',type:'text'},{name:'CHI Score',type:'number'},
    {name:'Performance Value',type:'number'},{name:'Experience Value',type:'number'},
    {name:'Business Value',type:'number'},{name:'Solution KPIs',type:'number'},
    {name:'Uptime',type:'number'},{name:'MTBF / MTTR',type:'number'},
    {name:'Frowns vs Smiles',type:'number'},{name:'Sentiment',type:'number'},
    {name:'Trust',type:'number'},{name:'Throughput Blueprint',type:'number'},
    {name:'Outcome Metric',type:'number'},{name:'Move the Needle',type:'number'},
    {name:'Rag Status',type:'drop_down',type_config:{options:[
      {name:'Green',color:'#548235'},{name:'Amber',color:'#BF8F00'},{name:'Red',color:'#FF0000'}]}}
  ];
  var fieldIds={},ragOptions=existingRagOptions;
  for(var key in existingFieldIds)fieldIds[key]=existingFieldIds[key];
  var created=0,skipped=0,failed=0;
  ss.toast('Creating missing custom fields...','⏳');
  for(var i=0;i<fieldDefs.length;i++){
    var def=fieldDefs[i];
    if(existingFieldIds[def.name]){Logger.log('Already exists: '+def.name);skipped++;continue;}
    Utilities.sleep(400);
    var payload={name:def.name,type:def.type};if(def.type_config)payload.type_config=def.type_config;
    try{
      var f=cuFetch_('POST','/list/'+listId+'/field',payload);
      if(f&&f.id){
        fieldIds[def.name]=f.id;
        if(def.name==='Rag Status'&&f.type_config&&f.type_config.options)
          for(var j=0;j<f.type_config.options.length;j++)
            ragOptions[f.type_config.options[j].name]=f.type_config.options[j].id;
        Logger.log('Created: '+def.name+' → '+f.id);created++;
      }else{Logger.log('No ID returned for: '+def.name);failed++;}
    }catch(e){Logger.log('Failed: '+def.name+' → '+e.message);failed++;}}
  saveFieldIds_(fieldIds,ragOptions);
  ss.toast('✅ Field setup complete\nCreated: '+created+'\nAlready existed: '+skipped+'\nFailed: '+failed,'✅');
  Logger.log('Field IDs: '+JSON.stringify(fieldIds));
}

function clickupReadFieldIds(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  if(!CU_TOKEN){ss.toast('API token not set. Run clickupSaveToken first.','❌');return;}
  ss.toast('Reading field IDs from ClickUp...','⏳');
  var list;try{list=findTestingList_();}catch(e){ss.toast(e.message,'❌');return;}
  var rawFields;
  try{rawFields=(cuFetch_('GET','/list/'+list.id+'/field').fields)||[];}
  catch(e){ss.toast('Could not read fields: '+e.message,'❌');return;}
  var fieldIds={},ragOptions={},found=[],missing=[];
  for(var i=0;i<rawFields.length;i++){
    var f=rawFields[i];
    if(!f.name)continue;
    fieldIds[f.name]=f.id;
    if(f.name==='Rag Status'&&f.type_config&&f.type_config.options){
      for(var j=0;j<f.type_config.options.length;j++)
        ragOptions[f.type_config.options[j].name]=f.type_config.options[j].id;}
  }
  for(var k=0;k<FIELD_NAMES.length;k++){
    if(fieldIds[FIELD_NAMES[k]])found.push(FIELD_NAMES[k]);
    else missing.push(FIELD_NAMES[k]);}
  saveFieldIds_(fieldIds,ragOptions);
  Logger.log('Found: '+JSON.stringify(found));
  Logger.log('Missing: '+JSON.stringify(missing));
  Logger.log('RAG options: '+JSON.stringify(ragOptions));
  if(missing.length===0){
    ss.toast('✅ All '+FIELD_NAMES.length+' field IDs stored.','✅');
  } else {
    ss.toast('⚠ Still missing: '+missing.join(', '),'⚠');
  }
}

function clickupUpdateScorecard(){
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
  Logger.log('Tasks: '+tasks.length+'  Month: '+chiData.monthLabel);

  ss.toast('Updating '+tasks.length+' tasks...','⏳');
  var updated=0,skipped=0;
  for(var t=0;t<tasks.length;t++){
    var task=tasks[t],matchKey=matchSite_(task.name,allKeys);
    if(!matchKey){Logger.log('No match: "'+task.name+'"');skipped++;continue;}
    var d=chiData.data[matchKey]||{};
    var chi=d['CHI Score']||null;
    Logger.log('→ '+task.name+' CHI='+chi);

    setField_(task.id,FIELD_IDS['CHI Score'],            d['CHI Score']||null);
    setField_(task.id,FIELD_IDS['Performance Value'],    d['Performance Value']||null);
    setField_(task.id,FIELD_IDS['Solution KPIs'],        d['Solution KPIs']||null);
    setField_(task.id,FIELD_IDS['Uptime'],               d['Uptime']||null);
    setField_(task.id,FIELD_IDS['MTBF / MTTR'],          d['MTBF / MTTR']||null);
    setField_(task.id,FIELD_IDS['Experience Value'],     d['Experience Value']||null);
    setField_(task.id,FIELD_IDS['Frowns vs Smiles'],     d['Frown vs Smile']||null);
    setField_(task.id,FIELD_IDS['Sentiment'],            d['Sentiment']||null);
    setField_(task.id,FIELD_IDS['Trust'],                d['Trust']||null);
    setField_(task.id,FIELD_IDS['Business Value'],        d['Business Value']||null);
    setField_(task.id,FIELD_IDS['Throughput Blueprint'], d['Thruput Blueprint']||null);
    setField_(task.id,FIELD_IDS['Outcome Metric'],       d['Outcome Metrics']||null);
    setField_(task.id,FIELD_IDS['Move the Needle'],      d['Move the Needle']||null);

    if(chi!==null){
      var optId=RAG_OPTS[ragKey_(chi)];
      if(optId)setField_(task.id,FIELD_IDS['Rag Status'],optId);
      else Logger.log('  RAG option not found for: '+ragKey_(chi));}
    updated++;Utilities.sleep(300);}
  ss.toast('✅ '+updated+' updated, '+skipped+' skipped.\nMonth: '+chiData.monthLabel,'✅ Done');
  Logger.log('=== DONE: '+updated+' updated, '+skipped+' skipped ===');
}

function clickupAddMissingTasks(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  if(!CU_TOKEN){ss.toast('API token not set. Run clickupSaveToken first.','❌');return;}
  ss.toast('Fetching existing tasks...','⏳');
  var list;try{list=findTestingList_();}catch(e){ss.toast(e.message,'❌');return;}
  var tasks=getAllTasks_(list.id);
  var existing={};
  for(var i=0;i<tasks.length;i++) existing[normHard_(tasks[i].name)]=true;
  var sites=getActiveSites_();
  var added=0,skipped=0;
  ss.toast('Adding missing sites...','⏳');
  for(var i=0;i<sites.length;i++){
    if(existing[normHard_(sites[i].name)]){skipped++;continue;}
    Utilities.sleep(350);
    try{cuFetch_('POST','/list/'+list.id+'/task',{name:sites[i].name});added++;}
    catch(e){Logger.log('Failed to add: '+sites[i].name+' — '+e.message);}
  }
  ss.toast('✅ Added: '+added+' new tasks\nAlready existed: '+skipped,'✅');
  Logger.log('Added: '+added+', skipped: '+skipped);
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

function clickupCreateList(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  if(!CU_TOKEN){ss.toast('API token not set. Run clickupSaveToken first.','❌');return;}
  ss.toast('Creating '+CU_LIST_NAME+' list...','⏳');
  var list=cuFetch_('POST','/folder/'+CU_FOLDER_ID+'/list',{name:CU_LIST_NAME});
  if(!list||!list.id){ss.toast('List creation failed.','❌');return;}
  Logger.log('Created list: '+list.id);
  ss.toast('Creating tasks...','⏳');
  var sites=getActiveSites_(),created=0;
  for(var i=0;i<sites.length;i++){
    Utilities.sleep(350);
    try{cuFetch_('POST','/list/'+list.id+'/task',{name:sites[i].name});created++;}
    catch(e){Logger.log('Task failed: '+sites[i].name);}}
  ss.toast('✅ List created with '+created+' tasks.','✅');
}

// ════════════════════════════════════════════════════════
// 24-HOUR AUTO-SYNC
// ════════════════════════════════════════════════════════
function clickupSetupDailySync(){
  var triggers=ScriptApp.getProjectTriggers();
  for(var i=0;i<triggers.length;i++)
    if(triggers[i].getHandlerFunction()==='clickupUpdateScorecard')ScriptApp.deleteTrigger(triggers[i]);
  ScriptApp.newTrigger('clickupUpdateScorecard').timeBased().everyDays(1).atHour(3).create();
  SpreadsheetApp.getActiveSpreadsheet().toast('✅ Auto-sync active — runs daily at 3 AM.','✅');
}
function clickupStopDailySync(){
  var triggers=ScriptApp.getProjectTriggers(),removed=0;
  for(var i=0;i<triggers.length;i++)
    if(triggers[i].getHandlerFunction()==='clickupUpdateScorecard'){ScriptApp.deleteTrigger(triggers[i]);removed++;}
  SpreadsheetApp.getActiveSpreadsheet().toast(removed>0?'✅ Auto-sync stopped.':'No trigger found.',removed>0?'✅':'ℹ️');
}
