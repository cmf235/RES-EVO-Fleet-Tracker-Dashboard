/* ============================================================================
   Dashboard write-back: edit Rig Up / Rig Down on the dashboard -> Google Sheet
   ----------------------------------------------------------------------------
   PASTE this at the BOTTOM of the "FleetSheetExtras.gs" file in the
   "Generate PO Request" Apps Script project (the same project behind the
   dashboard feed), then redeploy:
       Deploy -> Manage deployments -> the active Web app deployment
       (AKfycbxzRn4...QMWHCqMUg = the dashboard's API_URL) -> pencil (Edit)
       -> Version: New version -> Deploy -> Done.
   The /exec URL stays the same, so no dashboard change is needed for the URL.

   It adds a POST handler (doPost) the dashboard calls to save edits. It writes
   each field to the exact same cells buildDashboardData reads (PANEL_MAP), so
   the round-trip is consistent, combines date+time into the shared datetime
   cells, writes checkboxes as booleans, and clears the 90s feed cache
   ('dashData') so edits show on the next refresh.
   No existing code is modified — this only ADDS functions.
   ============================================================================ */

function doPost(e){
  var out={ok:false};
  try{
    var body=JSON.parse((e&&e.postData&&e.postData.contents)||'{}');
    if(body.action==='saveFleetPanel'){ out=saveFleetPanel_(body); }
    else { out={ok:false,error:'unknown action'}; }
  }catch(err){ out={ok:false,error:String(err)}; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// Field -> cell map. rigUp = cols J-X; rigDown mirrors it +16 cols (Z-AN).
// Matches PANEL_MAP in CalendarBuilder.gs. date+time share one cell (a datetime),
// as do logistics and the roustabout/misc "start" fields.
var WRITE_MAP_={
  rigUp:{location:'J3',pad:'O3',dateTime:'L4',logisticsDateTime:'S4',dayShift:'J5',nightShift:'O5',hotel:'T5',
    monolineIron:'J6',pumpdownIron:'J7',lowPressureIron:'J8',flowbackIron:'J9',highPressureHoses:'J10',rigMats:'J11',pullCables:'J12',
    loadTransports:'J14',loadQty:'X14',unloadTransports:'J15',unloadQty:'X15',semiHotshot:'J16',semiQty:'X16',
    ro_startTime:'M19',ro_end:'T19',ro_dayShift:'J20',ro_dayQty:'X20',ro_nightShift:'J21',ro_nightQty:'X21',
    misc_startTime:'M24',misc_end:'T24',misc_desc:'M23'},
  rigDown:{location:'Z3',pad:'AE3',dateTime:'AB4',logisticsDateTime:'AI4',dayShift:'Z5',nightShift:'AE5',hotel:'AJ5',
    monolineIron:'Z6',pumpdownIron:'Z7',lowPressureIron:'Z8',flowbackIron:'Z9',highPressureHoses:'Z10',rigMats:'Z11',pullCables:'Z12',
    loadTransports:'Z14',loadQty:'AN14',unloadTransports:'Z15',unloadQty:'AN15',semiHotshot:'Z16',semiQty:'AN16',
    ro_startTime:'AC19',ro_end:'AJ19',ro_dayShift:'Z20',ro_dayQty:'AN20',ro_nightShift:'Z21',ro_nightQty:'AN21',
    misc_startTime:'AC24',misc_end:'AJ24',misc_desc:'AC23'}
};

function findFleetSheet_(num){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sheets=ss.getSheets(), t=String(num);
  var pres=['FLEET#'+t,'FLEET'+t];
  for(var i=0;i<sheets.length;i++){
    var nm=sheets[i].getName().toUpperCase().split(' ').join('');
    for(var p=0;p<pres.length;p++){
      if(nm.indexOf(pres[p])===0){ var nx=nm.charAt(pres[p].length); if(nx<'0'||nx>'9') return sheets[i]; }
    }
  }
  return null;
}

// Combine a date ('YYYY-MM-DD' or 'M/D/YYYY') + time ('HH:MM' or 'h:mm AM/PM')
// into a Date for the shared datetime cells. Blank date -> '' (clears the cell).
function comboDT_(dateStr,timeStr){
  dateStr=String(dateStr==null?'':dateStr).trim(); if(!dateStr) return '';
  var y,mo,da,parts;
  if(dateStr.indexOf('-')>0){ parts=dateStr.split('-'); y=+parts[0];mo=+parts[1];da=+parts[2]; }
  else if(dateStr.indexOf('/')>0){ parts=dateStr.split('/'); mo=+parts[0];da=+parts[1];y=+parts[2]; if(y<100)y+=2000; }
  else { return dateStr; }
  if(!y||!mo||!da) return dateStr;
  var hh=0,mi=0,t=String(timeStr==null?'':timeStr).trim();
  if(t.indexOf(':')>0){ var tp=t.split(':'); hh=parseInt(tp[0],10)||0; mi=parseInt(tp[1],10)||0;
    var up=t.toUpperCase(); if(up.indexOf('PM')>-1&&hh<12)hh+=12; if(up.indexOf('AM')>-1&&hh===12)hh=0; }
  return new Date(y,mo-1,da,hh,mi,0);
}

function saveFleetPanel_(body){
  var num=String(body.fleet==null?'':body.fleet).trim(), side=body.side, data=body.data||{};
  if(!num||isNaN(Number(num))) return {ok:false,error:'bad fleet'};
  if(side!=='rigUp'&&side!=='rigDown') return {ok:false,error:'bad side'};
  var sh=findFleetSheet_(num); if(!sh) return {ok:false,error:'fleet tab not found for '+num};
  var M=WRITE_MAP_[side], writes=[];
  function S(c,v){ if(c) writes.push([c, v==null?'':String(v)]); }
  function B(c,v){ if(c) writes.push([c, v===true||v==='true'||v===1||v==='Yes']); }
  function N(c,v){ if(c){ var n=Number(v); writes.push([c, isNaN(n)?0:n]); } }
  if('location' in data) S(M.location,data.location);
  if('pad' in data) S(M.pad,data.pad);
  if(('date' in data)||('time' in data)) writes.push([M.dateTime, comboDT_(data.date,data.time)]);
  if(('logisticsDate' in data)||('logisticsTime' in data)) writes.push([M.logisticsDateTime, comboDT_(data.logisticsDate,data.logisticsTime)]);
  var bk=['dayShift','nightShift','hotel','monolineIron','pumpdownIron','lowPressureIron','flowbackIron','highPressureHoses','rigMats','pullCables','loadTransports','unloadTransports','semiHotshot'];
  for(var bi=0;bi<bk.length;bi++){ if(bk[bi] in data) B(M[bk[bi]],data[bk[bi]]); }
  var nk=['loadQty','unloadQty','semiQty'];
  for(var ni=0;ni<nk.length;ni++){ if(nk[ni] in data) N(M[nk[ni]],data[nk[ni]]); }
  var ro=data.roustabout||{};
  if(('start' in ro)||('startTime' in ro)) writes.push([M.ro_startTime, comboDT_(ro.start,ro.startTime)]);
  if('end' in ro) writes.push([M.ro_end, comboDT_(ro.end,'')]);
  if('dayShift' in ro) B(M.ro_dayShift,ro.dayShift);
  if('dayQty' in ro) N(M.ro_dayQty,ro.dayQty);
  if('nightShift' in ro) B(M.ro_nightShift,ro.nightShift);
  if('nightQty' in ro) N(M.ro_nightQty,ro.nightQty);
  var mc=data.misc||{};
  if(('start' in mc)||('startTime' in mc)) writes.push([M.misc_startTime, comboDT_(mc.start,mc.startTime)]);
  if('end' in mc) writes.push([M.misc_end, comboDT_(mc.end,'')]);
  if('desc' in mc) S(M.misc_desc,mc.desc);
  for(var i=0;i<writes.length;i++){ sh.getRange(writes[i][0]).setValue(writes[i][1]); }
  try{ CacheService.getScriptCache().remove('dashData'); }catch(e){}
  return {ok:true, fleet:num, side:side, wrote:writes.length};
}
