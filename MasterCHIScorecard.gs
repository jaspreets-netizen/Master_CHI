/**
 * Master CHI Scorecard v2.1
 * Changes from v2.0:
 *   - "Build Activation" removed from top menu (function still exists in script)
 *   - CEM Name column added (col C) between Site Name and CHI Sheet URL
 *   - Columns shifted: URL→D, Allow Access→E, Connection Status→F
 *   - 6 columns total (was 5)
 */

var CLR={DB:'#1F4E79',W:'#FFFFFF',P:'#2E75B6',E:'#548235',B:'#BF8F00',
  GY:'#EFEFEF',ME:'#DCEEFB',LBL:'#D9D9D9',SHE:'#B6D7A8'};
var MN_AB={2:'Feb',3:'Mar',4:'Apr',5:'May',6:'Jun',7:'Jul',8:'Aug',9:'Sep',10:'Oct',11:'Nov',12:'Dec'};

// ═══ MENU — Build Activation removed from top bar ═══
function onOpen(){SpreadsheetApp.getUi().createMenu('⚙ Master CHI')
  .addItem('📊 Build all trend sheets','buildAllTrends')
  .addToUi();}

function extractSheetId_(url){var m=String(url).match(/\/d\/([a-zA-Z0-9_-]+)/);return m?m[1]:'';}
function colLetter_(n){var s='';while(n>0){n--;s=String.fromCharCode(65+(n%26))+s;n=Math.floor(n/26);}return s;}

// ═══ ACTIVATION — now with CEM Name column (C) ═══
function buildMasterActivation(){
  var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss.getSheetByName('Activation');if(sh)ss.deleteSheet(sh);
  sh=ss.insertSheet('Activation',0);
  if(sh.getMaxColumns()<6)sh.insertColumnsAfter(sh.getMaxColumns(),6-sh.getMaxColumns());
  if(sh.getMaxRows()<110)sh.insertRowsAfter(sh.getMaxRows(),110-sh.getMaxRows());
  sh.setColumnWidth(1,30);sh.setColumnWidth(2,180);sh.setColumnWidth(3,160);sh.setColumnWidth(4,500);sh.setColumnWidth(5,200);sh.setColumnWidth(6,200);
  sh.getRange(1,1,sh.getMaxRows(),6).setFontFamily('Arial').setFontSize(11);
  sh.getRange(1,1).setValue('Master CHI Scorecard — Site Registry').setFontWeight('bold').setFontSize(14).setFontColor(CLR.DB);
  // Header row: #, Site Name, CEM Name, CHI Sheet URL, Allow Access, Connection Status
  sh.getRange(2,1,1,6).setValues([['#','Site Name','CEM Name','CHI Sheet URL','Allow Access','Connection Status']]).setBackground(CLR.DB).setFontColor(CLR.W).setFontWeight('bold');

  // Pre-fill Dillard's
  sh.getRange(3,1).setValue(1).setHorizontalAlignment('center');
  sh.getRange(3,2).setValue("Dillard's").setBackground(CLR.ME);
  sh.getRange(3,3).setValue('Jaspreet').setBackground(CLR.ME);
  sh.getRange(3,4).setValue('https://docs.google.com/spreadsheets/d/1FtBti0RsDJwuig1z71kdQL6rq6fKRFXUBHD4SF6d_RE/edit').setBackground(CLR.ME).setFontSize(9).setWrap(true);
  // E3: IMPORTRANGE formula to test connection
  sh.getRange(3,5).setFormula('=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1FtBti0RsDJwuig1z71kdQL6rq6fKRFXUBHD4SF6d_RE","Dashboard!A3")');
  // F3: Connection status based on E3
  sh.getRange(3,6).setFormula('=IF(D3="","",IF(ISTEXT(E3),"✅ Connected","⚠ Click E3 → Allow Access"))');
  sh.setRowHeight(3,30);

  // Rows 4-102: empty slots (100 total)
  for(var i=1;i<100;i++){var r=3+i;
    sh.getRange(r,1).setValue(i+1).setHorizontalAlignment('center').setFontColor('#ccc');
    sh.getRange(r,2).setBackground(CLR.ME);
    sh.getRange(r,3).setBackground(CLR.ME);
    sh.getRange(r,4).setBackground(CLR.ME);
    // E: auto IMPORTRANGE when D has URL
    sh.getRange(r,5).setFormula('=IF(D'+r+'="","",IMPORTRANGE(D'+r+',"Dashboard!A3"))');
    // F: connection status
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

// ═══ GET ACTIVE SITES — updated column positions (URL now col D) ═══
function getActiveSites_(){
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Activation');
  if(!sh)throw new Error('Run Build Activation first.');
  var sites=[],vals=sh.getRange(3,1,100,4).getValues();  // cols A-D now
  for(var i=0;i<vals.length;i++){
    var name=String(vals[i][1]).trim(),cem=String(vals[i][2]).trim(),url=String(vals[i][3]).trim();
    if(name&&url){var sid=extractSheetId_(url);if(sid)sites.push({name:name,cem:cem,sid:sid});}}
  return sites;
}

// ═══ MONTH/BIWEEK COLUMNS — UNCHANGED ═══
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

// ═══ BUILD TREND TAB — UNCHANGED ═══
function buildTrendTab_(tabName,dashRow,columns,titleSuffix,headerColor){
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
    for(var i=0;i<nc;i++){var cl=colLetter_(columns[i].dashCol);
      sh.getRange(r,2+i).setFormula('=IFERROR(IMPORTRANGE("https://docs.google.com/spreadsheets/d/'+site.sid+'","Dashboard!'+cl+dashRow+'"),"")').setNumberFormat('0.0').setHorizontalAlignment('center');}}
  if(ns>0&&nc>0){var dataRange=sh.getRange(3,2,Math.max(ns,1),nc);
    sh.setConditionalFormatRules(sh.getConditionalFormatRules().concat([
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(5).setBackground('#F4CCCC').setRanges([dataRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(5,6.99).setBackground('#FFF2CC').setRanges([dataRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(7).setBackground('#D9EAD3').setRanges([dataRange]).build()]));}
  sh.setFrozenRows(2);sh.setFrozenColumns(1);
  return ns+' sites × '+nc+' periods';
}

// ═══ BUILD ALL — UNCHANGED ═══
function buildAllTrends(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();ss.toast('Building trend sheets...','⏳');
  var mCols=getMonthColumns_(),bwCols=getBiweekColumns_(),r=[];
  r.push('CHI: '+buildTrendTab_('CHI Trend',3,mCols,'CHI Score',CLR.DB));
  r.push('Perf: '+buildTrendTab_('Performance Value Trend',4,mCols,'Performance Value',CLR.P));
  r.push('Exp: '+buildTrendTab_('Experience Value Trend',5,mCols,'Experience Value',CLR.E));
  r.push('Biz: '+buildTrendTab_('Business Value Trend',6,mCols,'Business Value',CLR.B));
  r.push('Frown: '+buildTrendTab_('Frown vs Smile Trend',9,bwCols,'Frown vs Smile',CLR.E));
  ss.toast(r.join('\n'),'✅ All trend sheets built');
}
