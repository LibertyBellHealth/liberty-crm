// ── CONFIGURATION ─────────────────────────────────────────────
function toast(msg, type, duration){
  type = type || 'success'; duration = duration || 3000;
  var c = document.getElementById('toastContainer');
  var t = document.createElement('div');
  t.className = 'toast ' + type;
  t.textContent = msg;
  c.appendChild(t);
  requestAnimationFrame(function(){ requestAnimationFrame(function(){ t.classList.add('show'); }); });
  setTimeout(function(){ t.classList.remove('show'); setTimeout(function(){ if(t.parentNode)t.parentNode.removeChild(t); }, 300); }, duration);
}

var API_BASE    = 'https://liberty-crm-api-cyb3dkhnd2e7a3cy.centralus-01.azurewebsites.net/api';
var API_APP_ID  = '0c1627c1-c186-4e46-b919-e4a12f2f3952'; // Easy Auth app registration
var _apiToken   = null; // cached Bearer token, refreshed automatically
var SP_CLIENT_ID = '63828fd5-e676-4dd7-bfaa-0055fdb9b3c7';
var SP_TENANT_ID = '12be0d3c-3e63-429f-bf46-1a2f746aa25f';
var REDIRECT_URI  = 'https://polite-pebble-039f4a010.7.azurestaticapps.net';

var ALLOWED_USERS = [
  'tommy@mybellcare.com',
  'paul@mybellcare.com',
  'rob@mybellcare.com'
];

function aiTrack(name, props) {
  try {
    var ai = window.appInsights;
    if (ai && ai.trackEvent) {
      var acc = msalInstance && msalInstance.getAllAccounts && msalInstance.getAllAccounts();
      var user = (acc && acc.length) ? acc[0].username : 'unknown';
      ai.trackEvent({ name: name }, Object.assign({ user: user, site: 'crm' }, props || {}));
    }
  } catch (e) { /* silent */ }
}

var _sessionTimer = null;
function resetSessionTimer() {
  clearTimeout(_sessionTimer);
  _sessionTimer = setTimeout(function () {
    aiTrack('SessionTimeout', { reason: '15min-inactivity' });
    clearCRMStorage();
    signOut();
  }, 15 * 60 * 1000);
}
['click','keydown','mousemove','touchstart'].forEach(function(ev){
  document.addEventListener(ev, resetSessionTimer, true);
});

function clearCRMStorage() {
  ['crm_preview','crm_carriers','crm_settings','crm_custom_meds',
   'crm_agents','crm_lead_sources','crm_renewals','crm_plan_types',
   'crm_project_codes','crm_display_name','crm_todos','crm_recent'].forEach(function(k){
    localStorage.removeItem(k);
  });
}

var msalInstance=null,spToken=null,clients=[],editingId=null,csvHeaders=[],csvData=[],currentReportData=[],carriers=[];

// ── MSAL authentication ────────────────────────────────────────
function initMSAL(){
  var config={auth:{clientId:SP_CLIENT_ID,authority:'https://login.microsoftonline.com/'+SP_TENANT_ID,redirectUri:REDIRECT_URI},cache:{cacheLocation:'localStorage',storeAuthStateInCookie:true}};
  msalInstance=new msal.PublicClientApplication(config);
  msalInstance.initialize().then(function(){
    msalInstance.handleRedirectPromise().then(function(resp){
      if(resp&&resp.account){onSignedIn(resp.account);return;}
      var accounts=msalInstance.getAllAccounts();
      if(accounts.length>0){onSignedIn(accounts[0]);}
      else{showAuthScreen();}
    }).catch(function(){showAuthScreen();});
  });
}
function showAuthScreen(){document.getElementById('authScreen').style.display='flex';document.getElementById('mainApp').style.display='none';}
function onSignedIn(account){
  var email=(account&&(account.username||account.name||'')).toLowerCase();
  if(!ALLOWED_USERS.map(function(u){return u.toLowerCase();}).includes(email)){
    alert('Access denied. Your account ('+email+') is not authorized for this application.');
    msalInstance.logoutRedirect({redirectUri:REDIRECT_URI});
    return;
  }
  document.getElementById('authScreen').style.display='none';
  document.getElementById('mainApp').style.display='flex';
  document.getElementById('userEmail').textContent=email;
  aiTrack('UserSignIn',{email:email});
  resetSessionTimer();
  refreshApiToken().then(function(){ loadClients(); });
}
var API_SCOPE='api://'+API_APP_ID+'/user_impersonation';
// Only Graph scopes in loginRedirect — API_SCOPE from different resource causes 400 on token endpoint.
// refreshApiToken() acquires API token silently after login (admin consent already granted).
function signIn(){msalInstance.loginRedirect({scopes:['openid','profile'],redirectUri:REDIRECT_URI});}
function signOut(){
  aiTrack('UserSignOut',{});
  clearCRMStorage();
  clearTimeout(_sessionTimer);
  _apiToken=null;
  msalInstance.logoutRedirect({redirectUri:REDIRECT_URI});
}

// ── API HELPERS ────────────────────────────────────────────────
function apiHeaders(){var h={'Content-Type':'application/json'};if(_apiToken)h['Authorization']='Bearer '+_apiToken;return h;}
function authUploadHeaders(){return _apiToken?{'Authorization':'Bearer '+_apiToken}:{};}
async function refreshApiToken(){
  if(!msalInstance)return;
  var accounts=msalInstance.getAllAccounts();if(!accounts.length)return;
  try{
    var res=await msalInstance.acquireTokenSilent({scopes:[API_SCOPE],account:accounts[0]});
    _apiToken=res.accessToken;
    var ttl=res.expiresOn?(res.expiresOn.getTime()-Date.now()-600000):3000000;
    setTimeout(refreshApiToken,Math.max(ttl,60000));
  }catch(e){
    console.warn('API token silent refresh failed, opening consent popup:',e);
    try{
      var r2=await msalInstance.acquireTokenPopup({scopes:[API_SCOPE]});
      _apiToken=r2.accessToken;
      var ttl2=r2.expiresOn?(r2.expiresOn.getTime()-Date.now()-600000):3000000;
      setTimeout(refreshApiToken,Math.max(ttl2,60000));
    }catch(e2){console.warn('API token popup also failed:',e2);}
  }
}

// ── DATA MAPPING (form fields ↔ DB columns) ────────────────────
function clientToDbRow(d){
  return {
    first_name:d.f_firstName, middle_initial:d.f_mi, last_name:d.f_lastName,
    dob:d.f_dob, age:parseInt(d.f_age)||null, gender:d.f_gender,
    ssn:d.f_ssn, relation:d.f_relation, marital_status:d.f_marital,
    tobacco:d.f_tobacco, height:d.f_height, weight:d.f_weight, insured:d.f_insured,
    phone:d.f_phone, alt_phone:d.f_altPhone, email:d.f_email, email2:d.f_email2,
    res_address:d.f_resAddress, res_zip:d.f_resZip, res_city:d.f_resCity,
    res_state:d.f_resSt, res_county:d.f_resCounty,
    same_address:d.f_sameAddress?1:0,
    bill_address:d.f_billAddress, bill_zip:d.f_billZip, bill_city:d.f_billCity,
    bill_state:d.f_billSt, bill_county:d.f_billCounty,
    plan_name:d.f_planName, plan_type:d.f_planType, plan_carrier:d.f_planCarrier, plan_network:d.f_type,
    plan_level:d.f_level, deductible:d.f_deductible, comoop:d.f_comoop,
    total_premium:d.f_totalPremium, subsidy:d.f_subsidy, premium:d.f_premium, app_fee:d.f_appFee,
    has_medicare:d.f_hasMedicare?1:0,
    medicare_num:d.f_medicareNum, medicare_a_eff:d.f_medicareA, medicare_b_eff:d.f_medicareB,
    has_medicaid:d.f_hasMedicaid?1:0, medicaid_num:d.f_medicaid, medicaid_eff:d.f_medicaidEff,
    waive_dental:d.f_waiveDental?1:0, total_monthly:d.f_totalMonthly,
    health_pay_date:d.f_healthPayDate, health_effective:d.f_healthEffective,
    ancil_pay_date:d.f_ancilPayDate, ancil_effective:d.f_ancilEffective,
    dental_pay_date:d.f_dentalPayDate, dental_effective:d.f_dentalEffective,
    total_first_month:d.f_totalFirstMonth,
    primary_employer:d.f_primaryEmployer, primary_income:d.f_primaryIncome,
    spouse_employer:d.f_spouseEmployer, spouse_income:d.f_spouseIncome,
    other_income1:d.f_otherIncome1, other_income_amt1:d.f_otherIncomeAmt1,
    other_income2:d.f_otherIncome2, other_income_amt2:d.f_otherIncomeAmt2,
    other_income3:d.f_otherIncome3, other_income_amt3:d.f_otherIncomeAmt3,
    total_income:d.f_totalIncome,
    emergency_name:d.f_emergencyName, emergency_relation:d.f_emergencyRelation,
    emergency_phone:d.f_emergencyPhone,
    bank_name:d.f_bankName, account_type:d.f_accountType,
    routing:d.f_routing, account_num:d.f_account, account_name:d.f_accountName,
    card_type:d.f_cardType, card_number:d.f_cardNumber, card_exp:d.f_cardExp, cvv:d.f_cvv,
    agent:d.f_agent, submitted_by:d.f_submittedBy, application_date:d.f_date,
    lead_source:d.f_leadSource, lead_date:d.f_leadDate, renewed:d.f_renewed,
    mothers_maiden:d.f_mothersMaiden, notes:d.f_notes,
    members_json:JSON.stringify(d.members||[]),
    doctors_json:JSON.stringify(d.doctors||[]),
    medications_json:JSON.stringify(d.meds||[]),
    ancil_plans_json:JSON.stringify(d.ancilPlans||[]),
    homecare_client_id:parseInt(d.f_homecareClientId)||null,
  };
}
function dbRowToClient(row){
  var c={
    _id:row.id,
    f_firstName:row.first_name, f_mi:row.middle_initial, f_lastName:row.last_name,
    f_dob:row.dob, f_age:row.age, f_gender:row.gender, f_ssn:row.ssn,
    f_relation:row.relation, f_marital:row.marital_status,
    f_tobacco:row.tobacco, f_height:row.height, f_weight:row.weight, f_insured:row.insured,
    f_phone:row.phone, f_altPhone:row.alt_phone, f_email:row.email, f_email2:row.email2,
    f_resAddress:row.res_address, f_resZip:row.res_zip, f_resCity:row.res_city,
    f_resSt:row.res_state, f_resCounty:row.res_county,
    f_sameAddress:!!row.same_address,
    f_billAddress:row.bill_address, f_billZip:row.bill_zip, f_billCity:row.bill_city,
    f_billSt:row.bill_state, f_billCounty:row.bill_county,
    f_planName:row.plan_name, f_planType:row.plan_type, f_planCarrier:row.plan_carrier, f_type:row.plan_network,
    f_level:row.plan_level, f_deductible:row.deductible, f_comoop:row.comoop,
    f_totalPremium:row.total_premium, f_subsidy:row.subsidy, f_premium:row.premium,
    f_appFee:row.app_fee,
    f_hasMedicare:!!row.has_medicare, f_medicareNum:row.medicare_num,
    f_medicareA:row.medicare_a_eff, f_medicareB:row.medicare_b_eff,
    f_hasMedicaid:!!row.has_medicaid, f_medicaid:row.medicaid_num, f_medicaidEff:row.medicaid_eff,
    f_waiveDental:!!row.waive_dental, f_totalMonthly:row.total_monthly,
    f_healthPayDate:row.health_pay_date, f_healthEffective:row.health_effective,
    f_ancilPayDate:row.ancil_pay_date, f_ancilEffective:row.ancil_effective,
    f_dentalPayDate:row.dental_pay_date, f_dentalEffective:row.dental_effective,
    f_totalFirstMonth:row.total_first_month,
    f_primaryEmployer:row.primary_employer, f_primaryIncome:row.primary_income,
    f_spouseEmployer:row.spouse_employer, f_spouseIncome:row.spouse_income,
    f_otherIncome1:row.other_income1, f_otherIncomeAmt1:row.other_income_amt1,
    f_otherIncome2:row.other_income2, f_otherIncomeAmt2:row.other_income_amt2,
    f_otherIncome3:row.other_income3, f_otherIncomeAmt3:row.other_income_amt3,
    f_totalIncome:row.total_income,
    f_emergencyName:row.emergency_name, f_emergencyRelation:row.emergency_relation,
    f_emergencyPhone:row.emergency_phone,
    f_bankName:row.bank_name, f_accountType:row.account_type,
    f_routing:row.routing, f_account:row.account_num, f_accountName:row.account_name,
    f_cardType:row.card_type, f_cardNumber:row.card_number, f_cardExp:row.card_exp,
    f_cvv:row.cvv,
    f_agent:row.agent, f_submittedBy:row.submitted_by, f_date:row.application_date,
    f_leadSource:row.lead_source, f_leadDate:row.lead_date, f_renewed:row.renewed,
    f_mothersMaiden:row.mothers_maiden, f_notes:row.notes,
    f_homecareClientId:row.homecare_client_id||'',
  };
  try{c.members=JSON.parse(row.members_json||'[]');}catch(e){c.members=[];}
  try{c.doctors=JSON.parse(row.doctors_json||'[]');}catch(e){c.doctors=[];}
  try{c.meds=JSON.parse(row.medications_json||'[]');}catch(e){c.meds=[];}
  try{c.ancilPlans=JSON.parse(row.ancil_plans_json||'[]');}catch(e){c.ancilPlans=[];}
  return c;
}

function loadClients(){
  fetch(API_BASE+'/health-clients',{headers:apiHeaders()})
  .then(function(r){return r.json();})
  .then(function(data){
    clients=data.map(function(row){return dbRowToClient(row);});
    // Seed the filtered-results cache so sort/pagination work before any filter is applied
    _clientFilteredCache=clients.slice();
    _lastClientsFetch=Date.now();
    // If any filter is currently active, re-run it against the new data so state is preserved
    var anyFilter=false;
    ['searchInput','searchSSN4','filterAgent','filterPlan','filterSpecial','filterRenewed','filterLeadSource'].forEach(function(id){
      var el=document.getElementById(id);if(el&&el.value)anyFilter=true;
    });
    if(anyFilter)filterClients();else renderClientTable(clients);
    renderReportCards();renderReminderBanner();refreshReferrerDatalist();
  }).catch(function(e){console.error('Load error:',e);});
}
function saveClientAPI(data,id){
  var body=clientToDbRow(data);
  if(id)body.id=id;
  return fetch(API_BASE+'/health-clients',{method:'POST',headers:apiHeaders(),body:JSON.stringify(body)})
    .then(function(r){return r.json();});
}
function deleteClientAPI(id){
  return fetch(API_BASE+'/health-clients/'+id,{method:'DELETE',headers:apiHeaders()});
}

var VIEWS=['viewClients','viewNew','viewForm','viewImport','viewReports','viewCarriers','viewAdvSearch','viewTodo','viewSettings','viewRecent'];
function showView(v){
  // Guard against leaving the edit form with unsaved changes
  var form=document.getElementById('viewForm');
  if(form&&form.style.display!=='none'&&_formDirty&&v!=='form_edit'){
    guardUnsavedChanges(function(){_doShowView(v);});
    return;
  }
  _doShowView(v);
}
function _doShowView(v){
  VIEWS.forEach(function(id){document.getElementById(id).style.display='none';});
  document.querySelectorAll('.nav-btn').forEach(function(b){b.classList.remove('active');});
  var map={clients:'viewClients',new:'viewNew',form_edit:'viewForm',import:'viewImport',reports:'viewReports',carriers:'viewCarriers',advSearch:'viewAdvSearch',todo:'viewTodo',settings:'viewSettings',recent:'viewRecent'};
  var navMap={clients:'navClients',new:'navNew',reports:'navReports',carriers:'navCarriers',advSearch:'navAdvSearch',todo:'navTodo',settings:'navSettings',recent:'navRecent'};
  var vid=map[v];if(vid)document.getElementById(vid).style.display='block';
  var nid=navMap[v];if(nid)document.getElementById(nid).classList.add('active');
  if(vid)document.querySelector('.main').scrollTop=0;
  if(v==='clients')maybeRevalidateClients();
  if(v==='carriers')renderCarriers();
  if(v==='form_edit'||v==='new')setTimeout(wireCopyableFields,50);
  if(v==='todo')renderTodos();
  if(v==='settings'){renderSettings();populateDefaultAgentSelect();}
  if(v==='advSearch')populateAdvSearchCarriers();
  if(v==='recent')renderRecentRecords();
}

function startNewApp(type){
  if(type==='life'){toast('Life App coming soon!','info');return;}
  try{clearForm();}catch(e){console.log('clearForm error:',e);}
  editingId=null;
  try{addDoctorRow();}catch(e){}
  try{addMedRow();}catch(e){}
  try{loadCarriersToSelect();}catch(e){}
  document.getElementById('formTitle').textContent='New Health Application';
  document.getElementById('deleteBtn').style.display='none';
  document.getElementById('deleteBtn2').style.display='none';
  showView('form_edit');
}
function fmtToday(){var d=new Date();return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');}

// ── Client-table sort / pagination / resize state (mirrors Home Care CRM pattern) ──
var _clientSort={key:'name',dir:'asc'};
var _clientColDefaults={name:14,dob:8,phone:11,email:16,planType:9,planName:12,premium:8,agent:10,renewed:7,homeCare:8};
var _clientColumnWidths=null;
var _clientPage=1;
var _clientPageSize=25;
(function loadClientPrefs(){
  try{
    var w=localStorage.getItem('lch_client_col_widths');if(w)_clientColumnWidths=JSON.parse(w);
    var ps=localStorage.getItem('lch_client_page_size');
    if(ps==='all')_clientPageSize=Infinity;else if(ps){var n=parseInt(ps);if(n>0)_clientPageSize=n;}
  }catch(e){}
})();
// Filtered-results cache — sort + pagination re-render from this without re-filtering.
var _clientFilteredCache=[];
function sortClientBy(key){if(_clientSort.key===key)_clientSort.dir=_clientSort.dir==='asc'?'desc':'asc';else{_clientSort.key=key;_clientSort.dir='asc';}renderClientTable(_clientFilteredCache.slice());}
function _clientSortCompare(a,b){
  var key=_clientSort.key,dir=_clientSort.dir==='asc'?1:-1;
  function val(k,c){
    if(k==='name')return ((c.f_firstName||'')+' '+(c.f_lastName||'')).toLowerCase();
    if(k==='dob'){if(!c.f_dob)return 0;var d=new Date(c.f_dob);return isNaN(d)?0:d.getTime();}
    if(k==='phone')return (c.f_phone||'').toLowerCase();
    if(k==='email')return (c.f_email||'').toLowerCase();
    if(k==='planType')return (c.f_planType||'').toLowerCase();
    if(k==='planName')return (c.f_planName||'').toLowerCase();
    if(k==='premium')return parseFloat(c.f_premium||0)||0;
    if(k==='agent')return (c.f_agent||'').toLowerCase();
    if(k==='renewed')return (c.f_renewed||'').toLowerCase();
    if(k==='homeCare')return (c.homecare_client_id||c.f_homecareClientId)?1:0;
    return 0;
  }
  var va=val(key,a),vb=val(key,b);
  if(typeof va==='string')return va.localeCompare(vb)*dir;
  return (va<vb?-1:va>vb?1:0)*dir;
}
function applyClientColWidths(){
  var headers=document.querySelectorAll('#clientTable thead th[data-col]');
  headers.forEach(function(th){
    var col=th.dataset.col;
    var w=(_clientColumnWidths&&_clientColumnWidths[col])||_clientColDefaults[col];
    if(w)th.style.width=w+'%';
  });
}
function startColResize(e,col){
  e.preventDefault();e.stopPropagation();
  var th=e.target.closest('th');if(!th)return;
  var startX=e.pageX,startWidth=th.offsetWidth,tableWidth=th.closest('table').offsetWidth;
  document.body.style.cursor='col-resize';document.body.style.userSelect='none';
  function onMove(ev){var d=ev.pageX-startX;var nx=Math.max(60,startWidth+d);var pct=(nx/tableWidth)*100;th.style.width=pct+'%';if(!_clientColumnWidths)_clientColumnWidths={};_clientColumnWidths[col]=pct;}
  function onUp(){document.removeEventListener('mousemove',onMove);document.removeEventListener('mouseup',onUp);document.body.style.cursor='';document.body.style.userSelect='';try{localStorage.setItem('lch_client_col_widths',JSON.stringify(_clientColumnWidths||{}));}catch(e){}}
  document.addEventListener('mousemove',onMove);document.addEventListener('mouseup',onUp);
}
function clientPageSizeChange(){var s=document.getElementById('clientPageSize');var v=s.value;if(v==='all')_clientPageSize=Infinity;else _clientPageSize=parseInt(v)||25;_clientPage=1;try{localStorage.setItem('lch_client_page_size',v);}catch(e){}renderClientTable(_clientFilteredCache.slice());}
function goToClientPage(n){_clientPage=n;renderClientTable(_clientFilteredCache.slice());}
function toggleKebab(e){e.stopPropagation();var m=document.getElementById('clientKebabMenu');if(m)m.classList.toggle('open');}
function closeKebab(){var m=document.getElementById('clientKebabMenu');if(m)m.classList.remove('open');}
document.addEventListener('click',function(e){var w=document.querySelector('.kebab-wrap');if(w&&!w.contains(e.target))closeKebab();});
function toggleSidebar(){
  var sb=document.getElementById('sidebar');if(!sb)return;
  sb.classList.toggle('collapsed');
  try{localStorage.setItem('lch_sidebar_collapsed',sb.classList.contains('collapsed')?'1':'0');}catch(e){}
}
(function restoreSidebar(){try{if(localStorage.getItem('lch_sidebar_collapsed')==='1'){var s=document.getElementById('sidebar');if(s)s.classList.add('collapsed');}}catch(e){}})();

/* Reminder banner — surfaces client renewal + age-milestone actions.
   Called after every loadClients() so numbers stay fresh. */
function renderReminderBanner(){
  var el=document.getElementById('reminderBanner');if(!el)return;
  if(!clients||!clients.length){el.style.display='none';return;}
  var yr=new Date().getFullYear();
  var now=new Date();
  var in60=new Date();in60.setDate(in60.getDate()+60);
  var renewingSoon=0,turning65=0,turning26=0;
  clients.forEach(function(c){
    if(c.f_renewed==='Not Renewed')renewingSoon++;
    if(c.f_dob){
      var by=parseInt((c.f_dob.split('-')[0]||c.f_dob.split('/')[2]||'0'));
      var age=yr-by;
      if(age===64||age===65)turning65++;
      if(age===25||age===26)turning26++;
    }
  });
  var items=[];
  if(renewingSoon)items.push('<span class="rb-item"><a onclick="quickFilterRenewed(\'Not Renewed\')">'+renewingSoon+' client'+(renewingSoon===1?'':'s')+' not renewed</a></span>');
  if(turning65)items.push('<span class="rb-item"><a onclick="quickFilterSpecial(\'turning65\')">'+turning65+' turning 65 this year</a></span>');
  if(turning26)items.push('<span class="rb-item"><a onclick="quickFilterSpecial(\'turning26\')">'+turning26+' turning 26 this year</a></span>');
  if(items.length){
    el.className='reminder-banner active';
    el.innerHTML='<span class="rb-label">Attention:</span> '+items.join(' ');
  } else {
    el.className='reminder-banner calm';
    el.innerHTML='<span class="rb-label">All caught up.</span> No renewals or age milestones need attention.';
  }
  el.style.display='flex';
}
function quickFilterRenewed(v){var s=document.getElementById('filterRenewed');if(s){s.value=v;filterClients();}}
function quickFilterSpecial(v){var s=document.getElementById('filterSpecial');if(s){s.value=v;filterClients();}}

/* SWR-style revalidation — refetch clients when user returns to Clients view
   if data is older than 30s. Silent — no spinner unless empty. */
var _lastClientsFetch=0;
function maybeRevalidateClients(){
  var age=Date.now()-_lastClientsFetch;
  if(age>30000&&spToken){loadClients();}
}
document.addEventListener('visibilitychange',function(){
  if(document.visibilityState==='visible'){
    var cur=document.getElementById('viewClients');
    if(cur&&cur.style.display!=='none')maybeRevalidateClients();
  }
});

function renderClientTable(data){
  var tbody=document.getElementById('clientTableBody');tbody.innerHTML='';

  // Sort the incoming filtered data using the current column-sort state
  data.sort(_clientSortCompare);

  // Reflect sort arrows on the headers
  var headers=document.querySelectorAll('#clientTable thead th.sortable');
  headers.forEach(function(h){h.classList.remove('sort-asc','sort-desc');if(h.dataset.sortkey===_clientSort.key)h.classList.add(_clientSort.dir==='asc'?'sort-asc':'sort-desc');});
  applyClientColWidths();

  // Pagination — slice to current page
  var totalMatched=data.length;
  var totalPages=Math.max(1,Math.ceil(totalMatched/_clientPageSize));
  if(_clientPage>totalPages)_clientPage=totalPages;
  var startIdx=(_clientPage-1)*_clientPageSize;
  var endIdx=Math.min(startIdx+_clientPageSize,totalMatched);
  var visible=data.slice(startIdx,endIdx);

  // Count + page controls in the footer
  var countEl=document.getElementById('clientCount');
  if(countEl){
    if(!totalMatched)countEl.textContent='';
    else if(_clientPageSize===Infinity||totalMatched<=_clientPageSize)countEl.textContent=totalMatched+' client'+(totalMatched===1?'':'s');
    else countEl.textContent='Showing '+(startIdx+1)+'–'+endIdx+' of '+totalMatched;
  }
  var pageControls=document.getElementById('clientPageControls');
  if(pageControls){
    if(totalPages<=1)pageControls.innerHTML='';
    else {
      var html='<button class="pg-btn" onclick="goToClientPage('+(_clientPage-1)+')"'+(_clientPage===1?' disabled':'')+'>‹</button>';
      var startPg=Math.max(1,_clientPage-2),endPg=Math.min(totalPages,startPg+4);
      if(endPg-startPg<4)startPg=Math.max(1,endPg-4);
      for(var p=startPg;p<=endPg;p++)html+='<button class="pg-btn'+(p===_clientPage?' active':'')+'" onclick="goToClientPage('+p+')">'+p+'</button>';
      html+='<button class="pg-btn" onclick="goToClientPage('+(_clientPage+1)+')"'+(_clientPage===totalPages?' disabled':'')+'>›</button>';
      pageControls.innerHTML=html;
    }
  }

  if(!totalMatched){document.getElementById('emptyState').style.display='block';return;}
  document.getElementById('emptyState').style.display='none';
  visible.forEach(function(c){
    var tr=document.createElement('tr');
    var bc='badge-blue';
    if(c.f_planType==='Medicare')bc='badge-green';
    else if(c.f_planType==='Medicaid')bc='badge-orange';
    else if(c.f_planType==='Short Term')bc='badge-red';
    var badge=c.f_planType?'<span class="badge '+bc+'">'+c.f_planType+'</span>':'';
    var phone=c.f_phone||'';var email=c.f_email||'';
    // Renewed status → dot + plain text (matches Home Care's status pattern)
    var renewedRaw=(c.f_renewed||'').trim();
    var renewedDotClass=renewedRaw==='2026 Renewed'||renewedRaw.indexOf('Renewed')===0?'renewed':(renewedRaw==='Not Renewed'?'notrenewed':'');
    var renewedCell=renewedRaw?'<span class="status-inline"><span class="status-dot '+renewedDotClass+'"></span>'+renewedRaw+'</span>':'';
    tr.innerHTML='<td><input type="checkbox" class="row-cb" data-id="'+c._id+'" onchange="updateBulkBtn()"></td>'+
      '<td><span class="client-name-link link-plain" onclick="editClient(\''+c._id+'\')">'+(c.f_firstName||'')+' '+(c.f_lastName||'')+'</span></td>'+
      '<td>'+(c.f_dob||'')+'</td>'+
      '<td>'+phone+(phone?'<a href="tel:'+phone+'" class="icon-btn" title="Call">'+SVG_PHONE+'</a><button class="icon-btn" onclick="copyText(\''+phone.replace(/'/g,"\\'")+'\',this)" title="Copy">'+SVG_COPY+'</button>':'')+'</td>'+
      '<td>'+email+(email?'<a href="mailto:'+email+'" class="icon-btn" title="Email">'+SVG_MAIL+'</a><button class="icon-btn" onclick="copyText(\''+email.replace(/'/g,"\\'")+'\',this)" title="Copy">'+SVG_COPY+'</button>':'')+'</td>'+
      '<td>'+badge+'</td>'+
      '<td>'+(c.f_planName||'')+'</td>'+
      '<td>'+(c.f_premium?'$'+c.f_premium:'')+'</td>'+
      '<td>'+(c.f_agent||'')+'</td>'+
      '<td>'+renewedCell+'</td>'+
'';
    tbody.appendChild(tr);
  });

  // Sync the page-size selector with saved value
  var sel=document.getElementById('clientPageSize');
  if(sel){var savedPs=_clientPageSize===Infinity?'all':String(_clientPageSize);if(sel.value!==savedPs)sel.value=savedPs;}
}
function filterClients(){
  var q=document.getElementById('searchInput').value.toLowerCase();
  var ssn4=document.getElementById('searchSSN4').value;
  var agent=document.getElementById('filterAgent').value;
  var plan=document.getElementById('filterPlan').value;
  var special=document.getElementById('filterSpecial').value;
  var renewed=document.getElementById('filterRenewed').value;
  var leadSource=document.getElementById('filterLeadSource').value;
  var yr=new Date().getFullYear();
  var filtered=clients.filter(function(c){
    if(q){var nm=((c.f_firstName||'')+' '+(c.f_lastName||'')).toLowerCase();var ph=(c.f_phone||'').toLowerCase();var em=(c.f_email||'').toLowerCase();if(!nm.includes(q)&&!ph.includes(q)&&!em.includes(q))return false;}
    if(ssn4&&c.f_ssn){var l4=(c.f_ssn||'').replace(/\D/g,'').slice(-4);if(l4!==ssn4)return false;}
    var norm=function(v){return (v||'').toString().trim().toLowerCase();};
    if(agent&&norm(c.f_agent)!==norm(agent))return false;
    if(plan&&norm(c.f_planType)!==norm(plan))return false;
    if(renewed&&norm(c.f_renewed)!==norm(renewed))return false;
    if(leadSource){
      // 'Referral' filter matches both plain 'Referral' and packed 'Referral: <name>'
      var cls=norm(c.f_leadSource),fls=norm(leadSource);
      var match=cls===fls||(fls==='referral'&&cls.indexOf('referral')===0);
      if(!match)return false;
    }
    if(special&&c.f_dob){
      var by=c.f_dob.split('/')[2]||c.f_dob.split('-')[0];var age=yr-parseInt(by);
      if(special==='turning65'&&age!==64&&age!==65)return false;
      if(special==='turning64'&&age!==63&&age!==64)return false;
      if(special==='turning26'&&age!==25&&age!==26)return false;
    } else if(special){return false;}
    return true;
  });
  // Cache the filtered set so sort/pagination don't have to re-run all filter predicates
  _clientFilteredCache=filtered;
  // Any filter change resets to page 1 — user asked for narrower results, show from the top
  _clientPage=1;
  renderClientTable(filtered);
}
function clearFilters(){
  document.getElementById('searchInput').value='';document.getElementById('searchSSN4').value='';
  ['filterAgent','filterPlan','filterSpecial','filterRenewed','filterLeadSource'].forEach(function(id){document.getElementById(id).value='';});
  _clientFilteredCache=clients.slice();
  _clientPage=1;
  renderClientTable(_clientFilteredCache);
}

var FIELDS=['firstName','mi','lastName','relation','marital','gender','tobacco','height','weight','insured','ssn','dob','age',
  'hasMedicare','medicareNum','medicareA','medicareB','hasMedicaid','medicaid','medicaidEff','mothersMaiden','homecareClientId',
  'planName','planType','planCarrier','deductible','comoop','type','level','totalPremium','subsidy','premium','appFee',
  'waiveDental','totalMonthly',
  'resAddress','resZip','resCity','resSt','resCounty',
  'billAddress','billZip','billCity','billSt','billCounty',
  'phone','altPhone','email','email2',
  'emergencyName','emergencyRelation','emergencyPhone',
  'bankName','accountType','routing','account','accountName','cardType','cardNumber','cardExp','cvv',
  'healthPayDate','healthEffective','ancilPayDate','ancilEffective','dentalPayDate','dentalEffective','totalFirstMonth',
  'primaryEmployer','primaryIncome','spouseEmployer','spouseIncome',
  'otherIncome1','otherIncomeAmt1','otherIncome2','otherIncomeAmt2','otherIncome3','otherIncomeAmt3','totalIncome',
  'agent','submittedBy','date','leadSource','leadDate','renewed','notes'];

function clearForm(){
  clearFormDirty();
  FIELDS.forEach(function(f){
    var el=document.getElementById('f_'+f);if(!el)return;
    if(el.type==='checkbox')el.checked=false;else el.value='';
  });
  ['membersContainer','doctorsContainer','medsContainer','ancilContainer','otherIncomeContainer'].forEach(function(id){
    var el=document.getElementById(id);if(el)el.innerHTML='';
  });
  updateOtherIncomeAddBtn();
  var ag=document.getElementById('f_agent');if(ag)ag.value=localStorage.getItem('crm_default_agent')||'Thomas Jaboro';
  var sa=document.getElementById('f_diffMailing');if(sa)sa.checked=false;
  var ms=document.getElementById('mailingAddressSection');if(ms)ms.style.display='none';
  var at=document.getElementById('f_addressType');if(at){at.value='Mailing';updateMailingTitle();}
  var rb=document.getElementById('f_referredBy');if(rb){rb.value='';rb.style.borderColor='';rb.title='';}_referrerPicked=false;
  var rbf=document.getElementById('referredByField');if(rbf)rbf.style.display='none';
  var mf=document.getElementById('medicareFields');if(mf)mf.style.display='none';
  var mcd=document.getElementById('medicaidFields');if(mcd)mcd.style.display='none';
  var wd=document.getElementById('waiveDentalField');if(wd)wd.style.display='none';
  var mc=document.getElementById('memberCount');if(mc)mc.textContent='0';
  var td=document.getElementById('totalMonthlyDisplay');if(td)td.textContent='$0.00';
  var ts=document.getElementById('f_totalMonthlyShow');if(ts)ts.value='';
  var rc=document.getElementById('f_resCounty');if(rc)rc.innerHTML='<option value=""></option>';
  var bc=document.getElementById('f_billCounty');if(bc)bc.innerHTML='<option value=""></option>';
}
function getFormData(){
  var data={};
  FIELDS.forEach(function(f){var el=document.getElementById('f_'+f);if(!el)return;if(el.type==='checkbox')data['f_'+f]=el.checked;else data['f_'+f]=el.value;});
  data.members=[];
  document.querySelectorAll('.member-row-data').forEach(function(row){
    var m={};var hasData=false;
    row.querySelectorAll('[data-field]').forEach(function(el){m[el.dataset.field]=el.value;if(el.value)hasData=true;});
    if(hasData)data.members.push(m);
  });
  data.doctors=[];document.querySelectorAll('.doctor-row-data').forEach(function(row){var d={};row.querySelectorAll('[data-field]').forEach(function(el){d[el.dataset.field]=el.value;});data.doctors.push(d);});
  data.meds=[];document.querySelectorAll('.med-row-data').forEach(function(row){var m={};row.querySelectorAll('[data-field]').forEach(function(el){m[el.dataset.field]=el.value;});data.meds.push(m);});
  data.ancilPlans=[];document.querySelectorAll('.ancil-row-data').forEach(function(row){var a={};row.querySelectorAll('[data-field]').forEach(function(el){a[el.dataset.field]=el.value;});data.ancilPlans.push(a);});
  // Dynamic other-income rows → pack into the 3 legacy DB columns (max 3 enforced in UI)
  var oiRows=document.querySelectorAll('#otherIncomeContainer .oi-row');
  for(var i=0;i<3;i++){
    var row=oiRows[i];
    data['f_otherIncome'+(i+1)]=row?row.querySelector('.oi-src').value:'';
    data['f_otherIncomeAmt'+(i+1)]=row?row.querySelector('.oi-amt').value:'';
  }
  // Referral packing: if lead source is a referral variant and a name is given,
  // save as "Referral: <name>" in the existing lead_source column (no schema change).
  var refByEl=document.getElementById('f_referredBy');
  var refBy=refByEl?refByEl.value.trim():'';
  var ls=(data.f_leadSource||'').trim();
  if(/^referral$/i.test(ls)&&refBy){
    // Canonicalize spelling so "sarah smith" saves as "Sarah Smith" if that
    // matches an existing entry — dedupes typos across clients.
    var canon=canonicalReferrerName(refBy);
    data.f_leadSource='Referral: '+(canon||refBy);
  }
  // Address type packing: prefix bill_address with [BILL] or [BOTH] when the user
  // marked the extra address as billing or both. Mailing (default) gets no prefix
  // for cleanliness. Only applies if the address is actually populated.
  var atEl=document.getElementById('f_addressType');
  var at=atEl?atEl.value:'Mailing';
  if(data.f_billAddress&&data.f_billAddress.trim()&&at&&at!=='Mailing'){
    data.f_billAddress='['+at.toUpperCase()+'] '+data.f_billAddress.replace(/^\[(BILL|BOTH|MAIL(?:ING)?)\]\s*/i,'');
  } else if(data.f_billAddress){
    // Strip any stale prefix so switching back to Mailing removes it
    data.f_billAddress=data.f_billAddress.replace(/^\[(BILL|BOTH|MAIL(?:ING)?)\]\s*/i,'');
  }
  return data;
}
function setFormData(data){
  FIELDS.forEach(function(f){var el=document.getElementById('f_'+f);if(!el||data['f_'+f]===undefined)return;if(el.type==='checkbox')el.checked=data['f_'+f]===true||data['f_'+f]==='true';else el.value=data['f_'+f]||'';});
  document.getElementById('membersContainer').innerHTML='';
  if(data.members&&data.members.length)data.members.forEach(function(m){addMemberRow(m);});
  document.getElementById('doctorsContainer').innerHTML='';
  if(data.doctors&&data.doctors.length)data.doctors.forEach(function(d){addDoctorRow(d);});else addDoctorRow();
  document.getElementById('medsContainer').innerHTML='';
  if(data.meds&&data.meds.length)data.meds.forEach(function(m){addMedRow(m);});else addMedRow();
  document.getElementById('ancilContainer').innerHTML='';
  if(data.ancilPlans&&data.ancilPlans.length)data.ancilPlans.forEach(function(a){addAncilRow(a);});
  var oiC=document.getElementById('otherIncomeContainer');
  if(oiC){oiC.innerHTML='';for(var i=1;i<=3;i++){var s=data['f_otherIncome'+i],a=data['f_otherIncomeAmt'+i];if(s||a)addOtherIncomeRow({source:s||'',amount:a||''});}updateOtherIncomeAddBtn();}
  var mcChecked=data.f_hasMedicare===true||data.f_hasMedicare==='true'||data.f_hasMedicare==='1';
  var mcdChecked=data.f_hasMedicaid===true||data.f_hasMedicaid==='true'||data.f_hasMedicaid==='1';
  document.getElementById('f_hasMedicare').checked=mcChecked;
  document.getElementById('f_hasMedicaid').checked=mcdChecked;
  document.getElementById('medicareFields').style.display=mcChecked?'block':'none';
  document.getElementById('medicaidFields').style.display=mcdChecked?'block':'none';
  // Auto-reveal mailing address section if the client actually has one saved
  var hasBill=!!(data.f_billAddress||data.f_billZip||data.f_billCity);
  var diffCb=document.getElementById('f_diffMailing');
  if(diffCb){diffCb.checked=hasBill;document.getElementById('mailingAddressSection').style.display=hasBill?'block':'none';}
  // Unpack address type prefix: '[BILL] 123 Main' → type='Billing', address='123 Main'
  var atEl=document.getElementById('f_addressType'),billInput=document.getElementById('f_billAddress');
  if(atEl){
    var raw=data.f_billAddress||'',m=raw.match(/^\[(BILL|BOTH|MAIL(?:ING)?)\]\s*(.*)$/i);
    if(m){
      var tag=m[1].toUpperCase();
      atEl.value=tag==='BILL'?'Billing':(tag==='BOTH'?'Both':'Mailing');
      if(billInput)billInput.value=m[2];
    } else {
      atEl.value='Mailing';
    }
    updateMailingTitle();
  }
  // Unpack packed referral: "Referral: John Smith" → leadSource="Referral", referredBy="John Smith"
  var lsEl=document.getElementById('f_leadSource'),rbEl=document.getElementById('f_referredBy');
  if(lsEl&&rbEl){
    var packed=(data.f_leadSource||'').match(/^referral\s*:\s*(.+)$/i);
    if(packed){lsEl.value='Referral';rbEl.value=packed[1].trim();_referrerPicked=true;rbEl.style.borderColor='';}
    else{rbEl.value='';_referrerPicked=false;rbEl.style.borderColor='';}
    toggleReferredBy();
  }
  if(data.f_resZip)restoreCounty(data.f_resZip,'res',data.f_resCounty);
  if(data.f_billZip)restoreCounty(data.f_billZip,'bill',data.f_billCounty);
  updateMemberCount();checkWaiveDental();calcTotalMonthly();
  if(editingId)renderClientTodos(editingId);
}
function editClient(id){
  var c=clients.find(function(x){return String(x._id)===String(id);});
  if(!c){toast('Could not find client record. Please refresh and try again.','error');return;}
  aiTrack('ClientRecordOpened',{clientId:id,clientName:(c.f_firstName||'')+' '+(c.f_lastName||'')});
  trackRecentRecord(id,c);
  editingId=id;
  try{clearForm();}catch(e){console.log('clearForm err:',e);}
  try{loadCarriersToSelect();}catch(e){}
  try{setFormData(c);}catch(e){console.log('setFormData err:',e);}
  clearFormDirty();
  document.getElementById('formTitle').textContent=(c.f_firstName||'')+' '+(c.f_lastName||'');
  document.getElementById('deleteBtn').style.display='inline-block';
  document.getElementById('deleteBtn2').style.display='inline-block';
  showView('form_edit');
  // Inject document upload section
  var formCard=document.querySelector('#viewForm .form-card');
  var existing=document.getElementById('clientDocsSection');
  if(existing)existing.remove();
  var docSec=document.createElement('div');
  docSec.className='form-section';
  docSec.id='clientDocsSection';
  var actions=formCard.querySelector('.form-actions');
  formCard.insertBefore(docSec,actions);
  loadClientDocs(id);
}
function saveClient(onSuccess){
  var data=getFormData();
  if(!data.f_firstName&&!data.f_lastName){toast('Please enter at least a first or last name.','error');return;}
  var isNew=!editingId;
  saveClientAPI(data,editingId).then(function(){
    aiTrack(isNew?'ClientCreated':'ClientUpdated',{clientName:(data.f_firstName||'')+' '+(data.f_lastName||''),clientId:editingId||'new'});
    clearFormDirty();
    loadClients();
    if(typeof onSuccess==='function')onSuccess();
    else showView('clients');
    toast('Client saved!','success');
  }).catch(function(e){toast('Error: '+e,'error');});
}
function deleteClient(){
  if(!editingId)return;
  var c=clients.find(function(x){return String(x._id)===String(editingId);});
  var name=c?((c.f_firstName||'')+' '+(c.f_lastName||'')).trim():'this client';
  showConfirm('Delete '+(name||'this client')+'? This cannot be undone.',function(){
    deleteClientAPI(editingId).then(function(){
      aiTrack('ClientDeleted',{clientId:editingId,clientName:name||editingId});
      loadClients();showView('clients');
    });
  },{title:'Delete Client',okText:'Delete'});
}

function fmtMoney(el){
  var v=el.value.replace(/[^0-9.]/g,'');
  el.value=v;
}
function fmtMoneyBlur(el){
  var v=parseFloat(el.value.replace(/[^0-9.]/g,''));
  if(!isNaN(v))el.value='$'+v.toFixed(2);
  else el.value='';
}
function formatPhone(el){var v=el.value.replace(/\D/g,'');if(v.length>=10)el.value='('+v.substr(0,3)+') '+v.substr(3,3)+'-'+v.substr(6,4);}
function formatDate(el){var v=el.value.replace(/\D/g,'');if(v.length>4)v=v.substr(0,2)+'/'+v.substr(2,2)+'/'+v.substr(4,4);else if(v.length>2)v=v.substr(0,2)+'/'+v.substr(2);el.value=v;}
function formatSSN(el){var v=el.value.replace(/\D/g,'').substr(0,9);if(v.length>5)el.value=v.substr(0,3)+'-'+v.substr(3,2)+'-'+v.substr(5,4);else if(v.length>3)el.value=v.substr(0,3)+'-'+v.substr(3);else el.value=v;}
function formatMedicare(el){var v=el.value.toUpperCase().replace(/[^A-Z0-9]/g,'').substr(0,11);if(v.length>7)el.value=v.substr(0,4)+'-'+v.substr(4,3)+'-'+v.substr(7,4);else if(v.length>4)el.value=v.substr(0,4)+'-'+v.substr(4);else el.value=v;}
function calcAge(){var v=document.getElementById('f_dob').value;if(!v)return;var d=new Date(v);var a=Math.floor((new Date()-d)/31557600000);if(!isNaN(a)&&a>=0&&a<120)document.getElementById('f_age').value=a;}
function calcMemberAge(dobEl,ageId){var v=dobEl.value;if(!v)return;var d=new Date(v);var a=Math.floor((new Date()-d)/31557600000);var el=document.getElementById(ageId);if(el&&!isNaN(a)&&a>=0&&a<120)el.value=a;}
function toggleReveal(id,btn){var el=document.getElementById(id);if(el.type==='password'){el.type='text';btn.textContent='Hide';}else{el.type='password';btn.textContent='Show';}}
function focusReveal(el){el.type='text';}
function blurReveal(el){el.type='password';}
function copyField(id,btn){var el=document.getElementById(id);var t=el.type;el.type='text';navigator.clipboard.writeText(el.value);el.type=t;if(btn){btn.classList.add('copied');var o=btn.innerHTML;btn.innerHTML=SVG_CHECK;setTimeout(function(){btn.classList.remove('copied');btn.innerHTML=o;},1200);}}
var SVG_PHONE='<svg viewBox="0 0 24 24"><path d="M22 16.9v3a2 2 0 0 1-2.2 2 19.8 19.8 0 0 1-8.6-3.1 19.5 19.5 0 0 1-6-6A19.8 19.8 0 0 1 2.1 4.2 2 2 0 0 1 4.1 2h3a2 2 0 0 1 2 1.7c.1.9.3 1.8.6 2.6a2 2 0 0 1-.5 2.1L8 9.6a16 16 0 0 0 6 6l1.2-1.2a2 2 0 0 1 2.1-.5c.8.3 1.7.5 2.6.6a2 2 0 0 1 1.7 2z"/></svg>';
var SVG_MAIL='<svg viewBox="0 0 24 24"><rect x="2" y="4" width="20" height="16" rx="2"/><polyline points="22 6 12 13 2 6"/></svg>';
var SVG_COPY='<svg viewBox="0 0 24 24"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>';
var SVG_CHECK='<svg viewBox="0 0 24 24"><polyline points="20 6 9 17 4 12"/></svg>';
function copyText(txt,btn){
  navigator.clipboard.writeText(txt);
  if(btn){var orig=btn.innerHTML;btn.classList.add('copied');btn.innerHTML=SVG_CHECK;setTimeout(function(){btn.classList.remove('copied');btn.innerHTML=orig;},1200);}
}
function toggleMedicare(){document.getElementById('medicareFields').style.display=document.getElementById('f_hasMedicare').checked?'block':'none';}
function toggleMedicaid(){document.getElementById('medicaidFields').style.display=document.getElementById('f_hasMedicaid').checked?'block':'none';}
function toggleMailingAddress(){
  var cb=document.getElementById('f_diffMailing');
  var sec=document.getElementById('mailingAddressSection');
  if(sec&&cb)sec.style.display=cb.checked?'block':'none';
}
function updateMailingTitle(){
  var t=(document.getElementById('f_addressType')||{}).value||'Mailing';
  var el=document.getElementById('mailingAddressTitleText');
  if(el)el.textContent=t==='Both'?'Mailing & Billing Address':t+' Address';
}
function calcTotalIncome(){
  var t=0;
  ['f_primaryIncome','f_spouseIncome'].forEach(function(id){var el=document.getElementById(id);if(el)t+=parseFloat((el.value||'').replace(/[^0-9.]/g,''))||0;});
  document.querySelectorAll('#otherIncomeContainer .oi-amt').forEach(function(el){t+=parseFloat((el.value||'').replace(/[^0-9.]/g,''))||0;});
  var out=document.getElementById('f_totalIncome');if(out)out.value=t>0?'$'+t.toLocaleString():'';
}
var OTHER_INCOME_MAX=3;
function addOtherIncomeRow(data){
  var c=document.getElementById('otherIncomeContainer');if(!c)return;
  var rows=c.querySelectorAll('.oi-row');
  if(rows.length>=OTHER_INCOME_MAX){toast('Maximum '+OTHER_INCOME_MAX+' other income sources','info');return;}
  data=data||{};
  var row=document.createElement('div');
  row.className='oi-row fg';
  row.style.cssText='grid-template-columns:2fr 0.8fr 30px;gap:8px;margin-bottom:8px;align-items:end;';
  row.innerHTML='<div class="field"><label>Other Income Source</label><input class="oi-src" value="'+(data.source||'').replace(/"/g,'&quot;')+'"></div>'+
    '<div class="field"><label>Income</label><input class="oi-amt" placeholder="$" value="'+(data.amount||'').replace(/"/g,'&quot;')+'" oninput="fmtMoney(this);calcTotalIncome()" onblur="fmtMoneyBlur(this);calcTotalIncome()"></div>'+
    '<button type="button" class="icon-btn" onclick="removeOtherIncomeRow(this)" title="Remove"><svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="1.8" fill="none" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button>';
  c.appendChild(row);
  updateOtherIncomeAddBtn();
  calcTotalIncome();
}
function removeOtherIncomeRow(btn){
  var row=btn.closest('.oi-row');if(row)row.remove();
  updateOtherIncomeAddBtn();
  calcTotalIncome();
}
/* In-app confirm modal — replaces browser confirm() dialogs.
   Usage: showConfirm('Delete this?', function(){...}, {title:'Delete', okText:'Delete'}) */
function showConfirm(message,onOk,opts){
  opts=opts||{};
  var modal=document.getElementById('confirmModal');if(!modal)return;
  document.getElementById('confirmTitle').textContent=opts.title||'Confirm';
  document.getElementById('confirmMessage').textContent=message||'Are you sure?';
  var okBtn=document.getElementById('confirmOkBtn'),cancelBtn=document.getElementById('confirmCancelBtn'),extraBtn=document.getElementById('confirmExtraBtn');
  okBtn.textContent=opts.okText||'Confirm';
  cancelBtn.textContent=opts.cancelText||'Cancel';
  okBtn.className='btn '+(opts.danger===false?'btn-blue':'btn-red');
  if(opts.extraText){
    extraBtn.style.display='';
    extraBtn.textContent=opts.extraText;
    extraBtn.className='btn '+(opts.extraClass||'btn-green');
  } else {
    extraBtn.style.display='none';
  }
  // Replace handlers freshly each open so old callbacks don't stack
  var newOk=okBtn.cloneNode(true),newCancel=cancelBtn.cloneNode(true),newExtra=extraBtn.cloneNode(true);
  okBtn.parentNode.replaceChild(newOk,okBtn);
  cancelBtn.parentNode.replaceChild(newCancel,cancelBtn);
  extraBtn.parentNode.replaceChild(newExtra,extraBtn);
  var close=function(){modal.style.display='none';};
  newOk.addEventListener('click',function(){close();if(typeof onOk==='function')onOk();});
  newCancel.addEventListener('click',function(){close();if(typeof opts.onCancel==='function')opts.onCancel();});
  newExtra.addEventListener('click',function(){close();if(typeof opts.onExtra==='function')opts.onExtra();});
  modal.style.display='flex';
}
/* Unsaved-changes tracking on the client edit form.
   Set dirty on any input inside #viewForm; cleared by clearForm/setFormData/saveClient. */
var _formDirty=false;
function markFormDirty(){_formDirty=true;}
function clearFormDirty(){_formDirty=false;}
document.addEventListener('DOMContentLoaded',function(){
  var f=document.getElementById('viewForm');if(!f)return;
  f.addEventListener('input',markFormDirty,true);
  f.addEventListener('change',markFormDirty,true);
});
/* Called by showView() and Cancel button before actually navigating away from viewForm.
   Returns true if we should proceed immediately; otherwise shows a modal and calls proceed() later. */
function guardUnsavedChanges(proceed){
  var cur=document.getElementById('viewForm');
  if(!cur||cur.style.display==='none'||!_formDirty){proceed();return;}
  showConfirm('You have unsaved changes. Save them before leaving?',
    function(){_formDirty=false;proceed();},
    {
      title:'Unsaved Changes',
      okText:'Discard & Leave',
      cancelText:'Stay',
      extraText:'Save & Leave',
      onExtra:function(){
        try{saveClient(function(){proceed();});}catch(e){_formDirty=false;proceed();}
      }
    });
}
/* Warn on tab-close only when the form is dirty */
window.addEventListener('beforeunload',function(e){
  if(_formDirty){e.preventDefault();e.returnValue='';}
});
/* Row-removal helper used by generated rows so onclick attributes stay short */
function confirmRemoveRow(el,message,after){
  showConfirm(message||'Remove this row?',function(){
    var row=el.closest('.member-row-data')||el.closest('.doctor-row-data')||el.closest('.med-row-data')||el.closest('.ancil-row-wrap')||el.parentNode;
    if(row&&row.remove)row.remove();
    if(typeof after==='function')after();
  },{title:'Remove',okText:'Remove'});
}
function openMedsPasteModal(){var m=document.getElementById('medsPasteModal');if(m){document.getElementById('medsPasteInput').value='';m.style.display='flex';setTimeout(function(){document.getElementById('medsPasteInput').focus();},50);}}
function closeMedsPasteModal(){var m=document.getElementById('medsPasteModal');if(m)m.style.display='none';}
/* Parse one line into {name, mg, frequency}. Handles common formats:
     Metformin 500mg BID
     Lisinopril 10 mg once daily
     Atorvastatin - 20mg - at bedtime
     Just a plain name (no dose) → name only */
function parseMedLine(line){
  var s=line.replace(/[–—]/g,'-').replace(/\s+/g,' ').trim();
  if(!s)return null;
  // Extract dose: number (optionally decimal) with OPTIONAL unit (mg/mcg/g/units/iu/ml).
  // Unit optional handles inputs like "Metformin 30 twice" — 30 → dose, "twice" → frequency.
  var doseMatch=s.match(/\b(\d+(?:\.\d+)?)\s*(mg|mcg|g|units?|iu|ml)?\b/i);
  if(!doseMatch)return{name:s.replace(/^-+|-+$/g,'').trim(),mg:'',frequency:''};
  var unit=(doseMatch[2]||'mg').toLowerCase();
  var mg=doseMatch[1]+unit;
  var before=s.slice(0,doseMatch.index).trim().replace(/[-,]+$/,'').trim();
  var after=s.slice(doseMatch.index+doseMatch[0].length).trim().replace(/^[-,]+/,'').trim();
  // If no name found before the number (line starts with number), skip — likely not a med
  if(!before)return{name:s,mg:'',frequency:''};
  return{name:before,mg:mg,frequency:after};
}
function importPastedMeds(){
  var txt=document.getElementById('medsPasteInput').value||'';
  var lines=txt.split(/\r?\n/).map(function(l){return l.trim();}).filter(Boolean);
  if(!lines.length){closeMedsPasteModal();return;}
  // Wipe out any empty starter row before importing
  var c=document.getElementById('medsContainer');
  if(c){Array.from(c.querySelectorAll('.med-row-data')).forEach(function(row){
    var vals=Array.from(row.querySelectorAll('[data-field]')).map(function(el){return el.value.trim();}).join('');
    if(!vals)row.remove();
  });}
  var added=0;
  lines.forEach(function(line){var m=parseMedLine(line);if(m&&(m.name||m.mg)){addMedRow(m);added++;}});
  closeMedsPasteModal();
  toast('Imported '+added+' medication'+(added===1?'':'s'),'success');
}
function openDocsPasteModal(){var m=document.getElementById('docsPasteModal');if(m){document.getElementById('docsPasteInput').value='';m.style.display='flex';setTimeout(function(){document.getElementById('docsPasteInput').focus();},50);}}
function closeDocsPasteModal(){var m=document.getElementById('docsPasteModal');if(m)m.style.display='none';}
/* Parse one line into {name, specialty}. First token before a dash/comma/pipe/tab is name; rest joined is specialty. */
function parseDoctorLine(line){
  var s=line.replace(/[–—]/g,'-').replace(/\s+/g,' ').trim();
  if(!s)return null;
  var m=s.match(/^(.+?)\s*(?:[-,|\t])\s*(.+)$/);
  if(!m)return{name:s,specialty:''};
  return{name:m[1].trim(),specialty:m[2].trim()};
}
function importPastedDoctors(){
  var txt=document.getElementById('docsPasteInput').value||'';
  var lines=txt.split(/\r?\n/).map(function(l){return l.trim();}).filter(Boolean);
  if(!lines.length){closeDocsPasteModal();return;}
  var c=document.getElementById('doctorsContainer');
  if(c){Array.from(c.querySelectorAll('.doctor-row-data')).forEach(function(row){
    var vals=Array.from(row.querySelectorAll('[data-field]')).map(function(el){return el.value.trim();}).join('');
    if(!vals)row.remove();
  });}
  var added=0;
  lines.forEach(function(line){var d=parseDoctorLine(line);if(d&&d.name){addDoctorRow(d);added++;}});
  closeDocsPasteModal();
  toast('Imported '+added+' doctor'+(added===1?'':'s'),'success');
}
function copyAllDoctors(btn){
  var rows=document.querySelectorAll('#doctorsContainer .doctor-row-data');
  var lines=[];
  rows.forEach(function(row){
    var name=(row.querySelector('[data-field="name"]')||{}).value||'';
    var spec=(row.querySelector('[data-field="specialty"]')||{}).value||'';
    var parts=[name,spec].filter(function(p){return p&&p.trim();});
    if(parts.length)lines.push(parts.join(' - '));
  });
  if(!lines.length){toast('No doctors to copy','info');return;}
  navigator.clipboard.writeText(lines.join('\n'));
  if(btn){var o=btn.textContent;btn.textContent='Copied '+lines.length+'!';setTimeout(function(){btn.textContent=o;},1400);}
}
function copyAllMeds(btn){
  var rows=document.querySelectorAll('#medsContainer .med-row-data');
  var lines=[];
  rows.forEach(function(row){
    var name=(row.querySelector('[data-field="name"]')||{}).value||'';
    var mg=(row.querySelector('[data-field="mg"]')||{}).value||'';
    var freq=(row.querySelector('[data-field="frequency"]')||{}).value||'';
    var parts=[name,mg,freq].filter(function(p){return p&&p.trim();});
    if(parts.length)lines.push(parts.join(' '));
  });
  if(!lines.length){toast('No medications to copy','info');return;}
  navigator.clipboard.writeText(lines.join('\n'));
  if(btn){var o=btn.textContent;btn.textContent='Copied '+lines.length+'!';setTimeout(function(){btn.textContent=o;},1400);}
}
function updateOtherIncomeAddBtn(){
  var c=document.getElementById('otherIncomeContainer');var b=document.getElementById('addOtherIncomeBtn');
  if(!c||!b)return;
  b.style.display=c.querySelectorAll('.oi-row').length>=OTHER_INCOME_MAX?'none':'';
}
function calcTotalMonthly(){
  var pEl=document.getElementById('f_premium');
  var h=pEl?parseFloat((pEl.value||'').replace(/[^0-9.]/g,''))||0:0;
  var a=0;document.querySelectorAll('.ancil-row-data').forEach(function(row){var p=row.querySelector('[data-field="premium"]');a+=parseFloat((p&&p.value||'').replace(/[^0-9.]/g,''))||0;});
  var totalMonthly=h+a;
  var disp=document.getElementById('totalMonthlyDisplay');if(disp)disp.textContent='$'+totalMonthly.toFixed(2);
  var hid=document.getElementById('f_totalMonthly');if(hid)hid.value=totalMonthly>0?'$'+totalMonthly.toFixed(2):'';
}
function checkWaiveDental(){/* no-op — waive dental is now inline in the Dental ancillary row */}
/* Credit card formatting: Amex (starts with 3) → 4-6-5; everything else → 4-4-4-4 */
function formatCardNumber(el){
  var v=(el.value||'').replace(/\D/g,'');
  var isAmex=v.charAt(0)==='3';
  if(isAmex){v=v.slice(0,15);var m=v.match(/^(\d{0,4})(\d{0,6})(\d{0,5})$/);el.value=m?[m[1],m[2],m[3]].filter(Boolean).join(' '):v;}
  else{v=v.slice(0,16);var m2=v.match(/^(\d{0,4})(\d{0,4})(\d{0,4})(\d{0,4})$/);el.value=m2?[m2[1],m2[2],m2[3],m2[4]].filter(Boolean).join(' '):v;}
}
function formatCardExp(el){
  var v=(el.value||'').replace(/\D/g,'').slice(0,4);
  if(v.length>=3)v=v.slice(0,2)+'/'+v.slice(2);
  el.value=v;
}
/* Detect card brand from first digit; only auto-select if Card Type is still blank
   so a manual override sticks. 3→Amex, 4→Visa, 5→Mastercard, 6→Discover. */
function autoDetectCardType(el){
  var first=(el.value||'').replace(/\D/g,'').charAt(0);
  var brand={'3':'Amex','4':'Visa','5':'Mastercard','6':'Discover'}[first];
  if(!brand)return;
  var sel=document.getElementById('f_cardType');
  if(sel&&!sel.value)sel.value=brand;
}
/* ROUTING_LOOKUP is loaded from routing-lookup.js — full FedACH directory
   (18k+ US banks) generated from Moov's open-source github.com/moov-io/fed
   dataset which mirrors the Federal Reserve FedACH Participants Directory. */
/* Look up bank name from ABA routing number. Only populates if empty so a manual
   entry isn't overwritten. If not in the local table, does nothing silently. */
function lookupBankFromRouting(rn){
  var digits=(rn||'').replace(/\D/g,'');
  if(digits.length!==9)return;
  var bankInput=document.getElementById('f_bankName');
  if(!bankInput||bankInput.value.trim())return;
  var name=ROUTING_LOOKUP[digits];
  if(name){
    bankInput.value=name;
    markFormDirty();
    toast('Bank identified: '+name,'success');
  }
}
/* Clamp date inputs so a 5+ digit year gets trimmed to 4 (browsers accept 6-digit years).
   Also enforces 1900–2099 to catch obvious typos. */
function clampDate(el){
  var v=el.value||'';var m=v.match(/^(\d{4,})-(\d{2})-(\d{2})$/);
  if(!m)return;
  var y=parseInt(m[1]);
  if(m[1].length>4||y<1900||y>2099){
    y=Math.max(1900,Math.min(2099,parseInt(m[1].slice(0,4))||2000));
    el.value=String(y).padStart(4,'0')+'-'+m[2]+'-'+m[3];
  }
}
/* Carrier autocomplete — filters `carriers` array by prefix + substring match.
   Populates the dropdown under the carrier input; click fills the field. */
function carrierAC(el){
  var q=(el.value||'').toLowerCase().trim();
  var list=document.getElementById('carrierACList');if(!list)return;
  if(!q){list.style.display='none';return;}
  var matches=carriers.filter(function(c){return (c.name||'').toLowerCase().indexOf(q)!==-1;}).slice(0,10);
  if(!matches.length){list.style.display='none';return;}
  // Rank: exact prefix first, then substring
  matches.sort(function(a,b){
    var ap=a.name.toLowerCase().indexOf(q)===0?0:1,bp=b.name.toLowerCase().indexOf(q)===0?0:1;
    return ap-bp;
  });
  list.innerHTML=matches.map(function(c){
    return '<div onmousedown="document.getElementById(\'f_planCarrier\').value=\''+(c.name||'').replace(/'/g,"\\'")+'\';document.getElementById(\'carrierACList\').style.display=\'none\';markFormDirty();">'+(c.name||'')+'</div>';
  }).join('');
  list.style.display='block';
}
/* Add hover-only copy button to every text/email/tel/password input. Runs after
   any dynamic content is inserted. Skips: no id, existing copyables, hidden inputs. */
function wireCopyableFields(){
  document.querySelectorAll('#viewForm input:not([type="checkbox"]):not([type="date"]):not([type="hidden"]):not([data-copyable])').forEach(function(inp){
    if(!inp.id)return;
    // Skip inputs that already sit inside a reveal-wrap (they have their own copy button)
    if(inp.parentNode&&inp.parentNode.classList&&inp.parentNode.classList.contains('reveal-wrap'))return;
    inp.setAttribute('data-copyable','1');
    var field=inp.closest('.field');if(!field)return;
    field.classList.add('field-copyable');
    var btn=document.createElement('button');
    btn.type='button';btn.className='copy-hover-btn';btn.title='Copy';
    btn.innerHTML='<svg viewBox="0 0 24 24" width="12" height="12" stroke="currentColor" stroke-width="1.8" fill="none" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>';
    btn.addEventListener('click',function(e){e.preventDefault();e.stopPropagation();if(!inp.value)return;navigator.clipboard.writeText(inp.value);btn.classList.add('copied');btn.innerHTML='<svg viewBox="0 0 24 24" width="12" height="12" stroke="currentColor" stroke-width="1.8" fill="none" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>';setTimeout(function(){btn.classList.remove('copied');btn.innerHTML='<svg viewBox="0 0 24 24" width="12" height="12" stroke="currentColor" stroke-width="1.8" fill="none" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>';},1200);});
    field.appendChild(btn);
  });
}
/* Install a delegated input listener that auto-clamps all date fields */
document.addEventListener('input',function(e){if(e.target&&e.target.type==='date')clampDate(e.target);});
/* Show Referred By field only when the lead source contains 'referral' */
function toggleReferredBy(){
  var v=(document.getElementById('f_leadSource')||{}).value||'';
  var field=document.getElementById('referredByField');
  if(field)field.style.display=/referral/i.test(v)?'':'none';
}
/* Build the referrer typeahead datalist from unique referrers already on file
   so the same person's name doesn't get typed 5 different ways. */
/* Backwards-compat stub — old callsite calls refreshReferrerDatalist() but the
   real work is now inside referrerAC() (built on demand from current data). */
function refreshReferrerDatalist(){}
/* Build a case-insensitive index of existing canonical referrer names + counts
   from the current clients array. Canonical spelling is the most common one. */
function buildReferrerIndex(){
  var buckets={}; // lowerName -> {name, count, spellings:{spelling:count}}
  (clients||[]).forEach(function(c){
    var m=(c.f_leadSource||'').match(/^referral\s*:\s*(.+)$/i);
    var raw=m?m[1].trim():'';
    if(!raw)return;
    var key=raw.toLowerCase();
    if(!buckets[key])buckets[key]={name:raw,count:0,spellings:{}};
    buckets[key].count++;
    buckets[key].spellings[raw]=(buckets[key].spellings[raw]||0)+1;
    // Canonical = most-used exact spelling
    var top=Object.keys(buckets[key].spellings).sort(function(a,b){return buckets[key].spellings[b]-buckets[key].spellings[a];})[0];
    buckets[key].name=top;
  });
  return buckets;
}
/* Look up canonical spelling for a typed name (case-insensitive). Returns null
   if no existing referrer matches — caller can decide to prompt "+ Add new". */
function canonicalReferrerName(typed){
  var t=(typed||'').trim().toLowerCase();
  if(!t)return null;
  var idx=buildReferrerIndex();
  return idx[t]?idx[t].name:null;
}
/* Autocomplete dropdown for the Referred By field. Shows existing referrers
   with counts; last item is always "+ Add new" if the typed text doesn't
   exactly match an existing entry. */
var _referrerPicked=false; // did the user explicitly pick or add this session?
function referrerAC(el){
  var list=document.getElementById('referrerPicker');if(!list)return;
  var q=(el.value||'').trim();var qLow=q.toLowerCase();
  var idx=buildReferrerIndex();
  var entries=Object.keys(idx).map(function(k){return idx[k];})
    .sort(function(a,b){return b.count-a.count;});
  var matches=q?entries.filter(function(e){return e.name.toLowerCase().indexOf(qLow)!==-1;}):entries;
  var exact=idx[qLow];
  var html='';
  matches.slice(0,8).forEach(function(e){
    var canonical=e.name.replace(/"/g,'&quot;');
    html+='<div onmousedown="pickReferrer(\''+canonical.replace(/'/g,"\\'")+'\')" style="display:flex;justify-content:space-between;">'+
      '<span>'+e.name+'</span>'+
      '<span style="color:var(--text-muted);font-size:11px;">'+e.count+' referral'+(e.count===1?'':'s')+'</span></div>';
  });
  if(q&&!exact){
    html+='<div onmousedown="addReferrer(\''+q.replace(/'/g,"\\'")+'\')" style="border-top:1px solid var(--border);color:var(--accent);font-weight:600;">'+
      '+ Add new referrer: "'+q+'"</div>';
  } else if(!q&&!entries.length){
    html='<div style="color:var(--text-muted);cursor:default;">Type a name to add your first referrer</div>';
  }
  list.innerHTML=html;
  list.style.display=html?'block':'none';
  updateReferrerFieldState(el,!!exact||!q);
}
function updateReferrerFieldState(el,valid){
  // Amber outline when the typed name isn't a canonical existing entry AND
  // the user hasn't explicitly picked / added — nudges them to pick, not
  // silently create a phantom referrer.
  if(!el)return;
  if(valid||_referrerPicked){el.style.borderColor='';el.title='';}
  else{el.style.borderColor='var(--dot-warning)';el.title='This is a new name — click "+ Add new referrer" in the dropdown to add them, or pick an existing one.';}
}
function pickReferrer(name){
  var el=document.getElementById('f_referredBy');if(!el)return;
  el.value=name;_referrerPicked=true;markFormDirty();
  document.getElementById('referrerPicker').style.display='none';
  updateReferrerFieldState(el,true);
}
function addReferrer(name){
  // Explicit add — user confirmed this is a genuinely new person, not a typo
  var el=document.getElementById('f_referredBy');if(!el)return;
  el.value=name.trim();_referrerPicked=true;markFormDirty();
  document.getElementById('referrerPicker').style.display='none';
  updateReferrerFieldState(el,true);
  toast('New referrer added: '+name.trim(),'success');
}
function referrerBlur(){
  var list=document.getElementById('referrerPicker');if(list)list.style.display='none';
  var el=document.getElementById('f_referredBy');if(!el)return;
  var typed=(el.value||'').trim();if(!typed){updateReferrerFieldState(el,true);return;}
  // Auto-canonicalize on blur — if the typed name matches an existing entry
  // case-insensitively, snap to that canonical spelling.
  var canonical=canonicalReferrerName(typed);
  if(canonical&&canonical!==typed){el.value=canonical;_referrerPicked=true;markFormDirty();}
  updateReferrerFieldState(el,!!canonical||_referrerPicked);
}
/* Count how many times each referrer appears — used by the top-referrers report card */
function getReferrerCounts(){
  var idx=buildReferrerIndex();
  return Object.keys(idx).map(function(k){return{name:idx[k].name,count:idx[k].count};})
    .sort(function(a,b){return b.count-a.count;});
}
function updateMemberCount(){document.getElementById('memberCount').textContent=document.querySelectorAll('.member-row-data').length+1;}
function populateCountySel(sel,counties,savedVal){
  sel.innerHTML='';
  if(!counties||!counties.length){var o=document.createElement('option');o.value='';o.textContent='';sel.appendChild(o);return;}
  if(counties.length>1){var b=document.createElement('option');b.value='';b.textContent='';sel.appendChild(b);}
  counties.forEach(function(cn){var o=document.createElement('option');o.value=cn;o.textContent=cn;if(savedVal&&cn===savedVal)o.selected=true;sel.appendChild(o);});
  if(counties.length===1&&!savedVal)sel.options[0].selected=true;
}
function fetchCountyByLatLon(lat,lon,prefix,savedVal){
  var sel=document.getElementById('f_'+prefix+'County');
  fetch('https://geo.fcc.gov/api/census/area?lat='+lat+'&lon='+lon+'&format=json').then(function(r){return r.json();}).then(function(data){
    var counties=[];var seen={};
    if(data&&data.results&&data.results.length>0){data.results.forEach(function(result){if(result.county_name&&!seen[result.county_name]){seen[result.county_name]=true;counties.push(result.county_name);}});}
    if(counties.length>0){populateCountySel(sel,counties,savedVal);}
    else{sel.innerHTML='<option value=""></option>';}
  }).catch(function(){sel.innerHTML='<option value=""></option>';});
}
function lookupZip(el,prefix){
  var zip=el.value.replace(/\D/g,'');if(zip.length!==5)return;
  fetch('https://api.zippopotam.us/us/'+zip).then(function(r){return r.json();}).then(function(data){
    if(!data.places||!data.places.length)return;
    document.getElementById('f_'+prefix+'City').value=data.places[0]['place name']||'';
    document.getElementById('f_'+prefix+'St').value=data.places[0]['state abbreviation']||'';
    fetchCountiesForPlaces(data.places,prefix,null);
  }).catch(function(){});
}
function restoreCounty(zip,prefix,saved){
  var z=(zip||'').replace(/\D/g,'');if(z.length!==5)return;
  fetch('https://api.zippopotam.us/us/'+z).then(function(r){return r.json();}).then(function(data){
    if(!data.places||!data.places.length)return;
    fetchCountiesForPlaces(data.places,prefix,saved);
  }).catch(function(){});
}
/* Query FCC for EVERY place in a zip (a single zip can span multiple counties),
   dedupe, then present all as options. If only one county is found, it's picked;
   otherwise the user gets a proper dropdown to choose. */
function fetchCountiesForPlaces(places,prefix,saved){
  var sel=document.getElementById('f_'+prefix+'County');if(!sel)return;
  Promise.all(places.map(function(p){
    if(!p.latitude||!p.longitude)return Promise.resolve([]);
    return fetch('https://geo.fcc.gov/api/census/area?lat='+p.latitude+'&lon='+p.longitude+'&format=json')
      .then(function(r){return r.json();})
      .then(function(d){var out=[];if(d&&d.results){d.results.forEach(function(r){if(r.county_name)out.push(r.county_name);});}return out;})
      .catch(function(){return [];});
  })).then(function(arrs){
    var seen={},counties=[];
    arrs.forEach(function(a){a.forEach(function(c){if(!seen[c]){seen[c]=true;counties.push(c);}});});
    counties.sort();
    if(counties.length)populateCountySel(sel,counties,saved);
    else sel.innerHTML='<option value=""></option>';
  });
}

function addMemberRow(data){
  var uid='m'+Date.now()+Math.floor(Math.random()*1000);
  var div=document.createElement('div');div.className='member-row-data member-row-compact';
  div.style.cssText='display:grid;grid-template-columns:1fr 0.25fr 1fr 0.7fr 0.5fr 0.5fr 0.5fr 0.4fr 0.4fr 0.85fr 0.4fr 1fr 0.5fr 30px;gap:5px;align-items:end;margin-bottom:6px;';
  div.innerHTML=
    mk('First Name','firstName',data)+mkS('MI','mi',data)+mk('Last Name','lastName',data)+
    mkSel('Relation','relation',['','Spouse','Child','Mother','Father','Other'],data)+
    mkSel('Married','married',['','Yes','No'],data)+
    mkSel('Gender','gender',['','M','F'],data)+
    mkSel('Tobacco','tobacco',['','Yes','No'],data)+
    '<div class="field"><label>Height</label><input data-field="height" value="'+(data&&data.height||'')+'" placeholder="5\'10&quot;" oninput="fmtHeight(this)"></div>'+
    '<div class="field"><label>Weight</label><input data-field="weight" value="'+(data&&data.weight||'')+'" placeholder="lbs"></div>'+
    '<div class="field"><label>DOB</label><input type="date" data-field="dob" id="'+uid+'_dob" value="'+(data&&data.dob||'')+'" onchange="calcMemberAge(this,\''+uid+'_age\')"></div>'+
    '<div class="field"><label>Age</label><input data-field="age" id="'+uid+'_age" readonly style="background:#f9f9f9;" value="'+(data&&data.age||'')+'"></div>'+
    '<div class="field"><label>SSN</label><input data-field="ssn" id="'+uid+'_ssn" type="password" placeholder="XXX-XX-XXXX" value="'+(data&&data.ssn||'')+'" oninput="formatSSN(this)" onfocus="focusReveal(this)" onblur="blurReveal(this)" maxlength="11"></div>'+
    mkSel('Insured','insured',['','Yes','No'],data)+
    '<button type="button" class="icon-btn" onclick="confirmRemoveRow(this,\'Remove this household member?\',updateMemberCount)" title="Remove"><svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="1.8" fill="none" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button>';
  document.getElementById('membersContainer').appendChild(div);
  updateMemberCount();
}
function mk(lbl,field,data){return '<div class="field"><label>'+lbl+'</label><input data-field="'+field+'" value="'+(data&&data[field]||'')+'"></div>';}
function mkS(lbl,field,data){return '<div class="field"><label>'+lbl+'</label><input data-field="'+field+'" value="'+(data&&data[field]||'')+'" maxlength="1"></div>';}
function mkSmall(lbl,field,data){
  var extra='';
  if(field==='height')extra=' oninput="fmtHeight(this)"';
  return '<div class="field"><label>'+lbl+'</label><input data-field="'+field+'" value="'+(data&&data[field]||'')+'"'+extra+'></div>';
}
function mkC(lbl,field,data){return '<div class="field"><label>'+lbl+'</label><input data-field="'+field+'" value="'+(data&&data[field]||'')+'"></div>';}
function mkSelC(lbl,field,opts,data){var val=data&&data[field]||'';var options=opts.map(function(o){return '<option'+(o===val?' selected':'')+'>'+o+'</option>';}).join('');return '<div class="field"><label>'+lbl+'</label><select data-field="'+field+'">'+options+'</select></div>';}
function mkSel(lbl,field,opts,data){
  var val=data&&data[field]||'';
  var options=opts.map(function(o){return '<option'+(o===val?' selected':'')+'>'+o+'</option>';}).join('');
  return '<div class="field"><label>'+lbl+'</label><select data-field="'+field+'">'+options+'</select></div>';
}
function addDoctorRow(data){
  var div=document.createElement('div');div.className='doctor-row-data';
  div.style.cssText='display:grid;grid-template-columns:2fr 1fr 30px;gap:6px;align-items:end;margin-bottom:6px;';
  div.innerHTML='<div class="field"><label>Doctor Name</label><input data-field="name" value="'+(data&&data.name||'')+'"></div>'+
    '<div class="field"><label>Specialty / Phone</label><input data-field="specialty" value="'+(data&&data.specialty||'')+'"></div>'+
    '<button type="button" class="icon-btn" onclick="confirmRemoveRow(this,\'Remove this doctor?\')" title="Remove"><svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="1.8" fill="none" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button>';
  document.getElementById('doctorsContainer').appendChild(div);
}

var _customMeds=[];
function loadCustomMeds(){try{_customMeds=JSON.parse(localStorage.getItem('crm_custom_meds')||'[]');}catch(e){_customMeds=[];}}
function saveCustomMed(name){if(!name)return;var n=name.trim();if(!n)return;if(MED_LIST.indexOf(n)===-1&&_customMeds.indexOf(n)===-1){_customMeds.push(n);_customMeds.sort();localStorage.setItem('crm_custom_meds',JSON.stringify(_customMeds));}}
loadCustomMeds();

var MED_LIST=['Abilify','Acarbose','Accolate','Accupril','Aciphex','Actonel','Actos','Adderall','Adderall XR','Advair','Advair Diskus','Aggrenox','Aldactone','Alendronate','Albuterol','Aleve','Allopurinol','Alprazolam','Altace','Ambien','Ambien CR','Amlodipine','Amoxicillin','Amoxicillin-Clavulanate','Amphetamine','Anastrozole','Androgel','Apixaban','Aripiprazole','Aspirin','Atenolol','Atomoxetine','Atorvastatin','Ativan','Augmentin','Azithromycin','Baclofen','Basaglar','Benadryl','Benazepril','Benicar','Benzonatate','Bisoprolol','Brilinta','Breo','Budesonide','Buprenorphine','Bupropion','Buspirone','Byetta','Bydureon','Caduet','Calcitriol','Carbamazepine','Carbidopa-Levodopa','Carvedilol','Celebrex','Celexa','Cephalexin','Cetirizine','Chantix','Cialis','Ciprofloxacin','Citalopram','Clindamycin','Clobetasol','Clonazepam','Clonidine','Clopidogrel','Colchicine','Colcrys','Combivent','Concerta','Coreg','Coreg CR','Coumadin','Cozaar','Crestor','Cyclobenzaprine','Cymbalta','Dapagliflozin','Dexamethasone','Dexilant','Dextroamphetamine','Diazepam','Diclofenac','Digoxin','Diltiazem','Diphenhydramine','Donepezil','Doxazosin','Doxycycline','Dulaglutide','Duloxetine','Dupixent','Effexor','Effexor XR','Eliquis','Empagliflozin','Enalapril','Entresto','Epidiolex','Escitalopram','Esomeprazole','Estradiol','Evista','Ezetimibe','Famotidine','Farxiga','Fentanyl','Ferrous Sulfate','Fexofenadine','Finasteride','Flagyl','Flexeril','Flomax','Flovent','Fluconazole','Fluoxetine','Fluticasone','Fluticasone-Salmeterol','Folic Acid','Fosamax','Furosemide','Gabapentin','Glimepiride','Glipizide','Glucophage','Glucotrol','Glyburide','Humalog','Humulin','Humulin N','Humulin R','Hydrochlorothiazide','Hydrocodone','Hydrocodone-Acetaminophen','Hydrocortisone','Hydroxychloroquine','Hydroxyzine','Ibuprofen','Invega','Invokamet','Invokana','Ipratropium','Irbesartan','Isosorbide','Janumet','Januvia','Jardiance','Juvisync','Ketamine','Klonopin','Lamictal','Lamotrigine','Lansoprazole','Lantus','Lantus SoloStar','Latuda','Levemir','Levofloxacin','Levothyroxine','Lexapro','Linagliptin','Linzess','Liraglutide','Lisinopril','Lisinopril-HCTZ','Lithium','Lopressor','Loratadine','Lorazepam','Losartan','Lovastatin','Lozol','Lyrica','Mavyret','Medroxyprogesterone','Meloxicam','Metformin','Metformin ER','Methocarbamol','Methylphenidate','Methylprednisolone','Metoprolol','Metoprolol Succinate','Metoprolol Tartrate','Metronidazole','Mirtazapine','Monjaro','Montelukast','Morphine','Mounjaro','Mucinex','Naproxen','Neurontin','Nexium','Nifedipine','Nitrofurantoin','Nitroglycerin','Norco','Nortriptyline','Novolin','Novolog','Novolog FlexPen','Nuvaring','Olmesartan','Omeprazole','Ondansetron','Oseltamivir','Ozempic','Oxycodone','Oxycodone-Acetaminophen','Oxycontin','Pantoprazole','Paroxetine','Paxil','Penicillin','Percocet','Phenergan','Phentermine','Plavix','Potassium Chloride','Pradaxa','Pravastatin','Prednisone','Pregabalin','Premarin','Prilosec','Pristiq','Proair','Prolia','Promethazine','Propranolol','Protonix','Provigil','Prozac','Quetiapine','Ramipril','Ranexa','Ranitidine','Reclipsen','Renvela','Repaglinide','Restasis','Rexulti','Risperidone','Ritalin','Rivaroxaban','Rosiglitazone','Rosuvastatin','Rybelsus','Saxenda','Semaglutide','Senna','Seroquel','Sertraline','Simvastatin','Singulair','Sitagliptin','Skyrizi','Solifenacin','Spironolactone','Strattera','Sulfamethoxazole','Sumatriptan','Symbicort','Synthroid','Tacrolimus','Tamsulosin','Temazepam','Testosterone','Tiotropium','Tizanidine','Topamax','Topiramate','Torsemide','Toujeo','Tramadol','Tradjenta','Trazodone','Tresiba','Trulicity','Valacyclovir','Valium','Valsartan','Venlafaxine','Ventolin','Vesicare','Viberzi','Victoza','Viibryd','Vimpat','Vitamin B12','Vitamin D','Voltaren','Vraylar','Warfarin','Wegovy','Wellbutrin','Xanax','Xarelto','Xifaxan','Xolair','Zestril','Zetia','Ziprasidone','Zofran','Zoloft','Zolpidem','Zopiclone','Zyprexa','Zyrtec'];
function getAllMeds(){return MED_LIST.concat(_customMeds).sort(function(a,b){return a.toLowerCase()<b.toLowerCase()?-1:1;});}
function addMedRow(data){
  var div=document.createElement('div');div.className='med-row-data';
  div.style.cssText='display:grid;grid-template-columns:2fr 80px 1fr auto 30px;gap:6px;align-items:end;margin-bottom:6px;position:relative;';
  div.innerHTML='<div class="field autocomplete-wrap"><label>Medication Name</label><input data-field="name" placeholder="Start typing..." value="'+(data&&data.name||'')+'" oninput="medAC(this)" onblur="medBlur(this)" autocomplete="off"><div class="autocomplete-list"></div></div>'+
    '<div class="field"><label>Mg</label><input data-field="mg" value="'+(data&&data.mg||'')+'"></div>'+
    '<div class="field"><label>Frequency</label><input data-field="frequency" value="'+(data&&data.frequency||'')+'"></div>'+
    '<button type="button" class="icon-btn" onclick="saveMedFromRow(this)" title="Save this medication to your library"><svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="1.8" fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M19 21l-7-5-7 5V5a2 2 0 0 1 2-2h10a2 2 0 0 1 2 2z"/></svg></button>'+
    '<button type="button" class="icon-btn" onclick="confirmRemoveRow(this,\'Remove this medication?\')" title="Remove"><svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="1.8" fill="none" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button>';
  document.getElementById('medsContainer').appendChild(div);
}
function saveMedFromRow(btn){
  var row=btn.closest('.med-row-data');
  var nameEl=row.querySelector('[data-field="name"]');
  var name=(nameEl&&nameEl.value||'').trim();
  if(!name){toast('Please enter a medication name first.','error');return;}
  if(MED_LIST.indexOf(name)!==-1||_customMeds.indexOf(name)!==-1){toast('"'+name+'" is already in your library','info');return;}
  saveCustomMed(name);
  toast('"'+name+'" saved to your library','success');
  var orig=btn.innerHTML;
  btn.classList.add('copied');
  // Filled bookmark = saved state
  btn.innerHTML='<svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="1.8" fill="currentColor" stroke-linecap="round" stroke-linejoin="round"><path d="M19 21l-7-5-7 5V5a2 2 0 0 1 2-2h10a2 2 0 0 1 2 2z"/></svg>';
  setTimeout(function(){btn.classList.remove('copied');btn.innerHTML=orig;},1600);
}
function medAC(el){
  var v=el.value.toLowerCase();var list=el.parentNode.querySelector('.autocomplete-list');
  if(!v||v.length<2){list.style.display='none';return;}
  var allMeds=getAllMeds();
  var m=allMeds.filter(function(x){return x.toLowerCase().startsWith(v);}).slice(0,8);
  if(!m.length){list.style.display='none';return;}
  list.innerHTML='';m.forEach(function(med){var d=document.createElement('div');d.textContent=med;var isCustom=_customMeds.indexOf(med)!==-1;if(isCustom)d.style.cssText='color:#065f46;font-style:italic;';d.addEventListener('mousedown',function(e){e.preventDefault();el.value=med;list.style.display='none';});list.appendChild(d);});
  list.style.display='block';
}
function medBlur(el){setTimeout(function(){var list=el.parentNode.querySelector('.autocomplete-list');if(list)list.style.display='none';},200);}
document.addEventListener('click',function(e){if(!e.target.closest('.autocomplete-wrap'))document.querySelectorAll('.autocomplete-list').forEach(function(l){l.style.display='none';});});
var EMAIL_DOMAINS=['gmail.com','yahoo.com','icloud.com','outlook.com','hotmail.com','aol.com','comcast.net','att.net','verizon.net','me.com','msn.com','live.com'];
function emailAC(el){
  var v=el.value;var list=el.parentNode.querySelector('.autocomplete-list');
  var atIdx=v.lastIndexOf('@');
  if(atIdx<0){list.style.display='none';return;}
  var local=v.substring(0,atIdx+1);var typed=v.substring(atIdx+1).toLowerCase();
  var matches=EMAIL_DOMAINS.filter(function(d){return d.startsWith(typed);}).slice(0,6);
  if(!matches.length){list.style.display='none';return;}
  list.innerHTML='';
  matches.forEach(function(d){var div=document.createElement('div');div.textContent=local+d;div.addEventListener('mousedown',function(e){e.preventDefault();el.value=local+d;list.style.display='none';});list.appendChild(div);});
  list.style.display='block';
}

function addAncilRow(data){
  var type=data?data.type:document.getElementById('ancilTypeSelect').value;
  if(!type){toast('Please select a plan type first.','error');return;}
  document.getElementById('ancilTypeSelect').value='';
  data=data||{};
  var wrap=document.createElement('div');wrap.className='ancil-row-wrap';
  wrap.style.cssText='padding:8px 10px;border:1px solid var(--border);border-radius:6px;background:var(--surface-alt);margin-bottom:8px;';
  var main=document.createElement('div');main.className='ancil-row-data';
  main.style.cssText='display:grid;grid-template-columns:0.7fr 1.4fr 0.7fr 1fr 1fr 30px;gap:6px;align-items:end;';
  main.innerHTML='<div class="field"><label>Type</label><input data-field="type" value="'+type+'" readonly style="background:var(--surface);"></div>'+
    '<div class="field"><label>Plan Name</label><input data-field="planName" value="'+(data.planName||'').replace(/"/g,'&quot;')+'"></div>'+
    '<div class="field"><label>Premium</label><input data-field="premium" placeholder="$" value="'+(data.premium||'').replace(/"/g,'&quot;')+'" oninput="fmtMoney(this);calcTotalMonthly()" onblur="fmtMoneyBlur(this);calcTotalMonthly()"></div>'+
    '<div class="field"><label>Pay Date</label><input type="date" data-field="payDate" value="'+(data.payDate||'')+'"></div>'+
    '<div class="field"><label>Effective</label><input type="date" data-field="effective" value="'+(data.effective||'')+'"></div>'+
    '<button type="button" class="icon-btn" onclick="confirmRemoveRow(this,\'Remove this ancillary plan?\',calcTotalMonthly)" title="Remove"><svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="1.8" fill="none" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button>';
  wrap.appendChild(main);
  // Dental-only extras: Waive Dental checkbox + previous-carrier fields
  if(type==='Dental'){
    var extras=document.createElement('div');
    extras.className='ancil-dental-extras';
    extras.style.cssText='margin-top:8px;padding-top:8px;border-top:1px dashed var(--border);';
    var waived=data.waiveDental?'checked':'';
    var prevDisp=data.waiveDental?'':'none';
    extras.innerHTML='<label style="display:inline-flex;align-items:center;gap:6px;font-size:12px;color:var(--text);cursor:pointer;">'+
      '<input type="checkbox" data-field="waiveDental" '+waived+' onchange="toggleWaiveDentalFields(this)"> Waive Dental (client has existing dental coverage)</label>'+
      '<div class="dental-prev-fields fg g2" style="gap:8px;margin-top:8px;display:'+(data.waiveDental?'grid':'none')+';">'+
        '<div class="field"><label>Previous Dental Carrier</label><input data-field="prevCarrier" value="'+(data.prevCarrier||'').replace(/"/g,'&quot;')+'"></div>'+
        '<div class="field"><label>Previous Member #</label><input data-field="prevMemberNum" value="'+(data.prevMemberNum||'').replace(/"/g,'&quot;')+'"></div>'+
      '</div>';
    wrap.appendChild(extras);
  }
  document.getElementById('ancilContainer').appendChild(wrap);
  calcTotalMonthly();
}
function toggleWaiveDentalFields(cb){
  var extras=cb.closest('.ancil-dental-extras');if(!extras)return;
  var fields=extras.querySelector('.dental-prev-fields');if(fields)fields.style.display=cb.checked?'grid':'none';
}

function toggleSelectAll(cb){document.querySelectorAll('.row-cb').forEach(function(c){c.checked=cb.checked;});updateBulkBtn();}
function updateBulkBtn(){var n=document.querySelectorAll('.row-cb:checked').length;var btn=document.getElementById('bulkDeleteBtn');btn.style.display=n>0?'inline-block':'none';btn.textContent='Delete Selected ('+n+')';}
function bulkDelete(){
  var ids=Array.from(document.querySelectorAll('.row-cb:checked')).map(function(cb){return cb.dataset.id;});
  if(!ids.length)return;
  showConfirm('Delete '+ids.length+' client'+(ids.length===1?'':'s')+'? This cannot be undone.',function(){
    Promise.all(ids.map(function(id){return deleteClientAPI(id);})).then(function(){loadClients();document.getElementById('bulkDeleteBtn').style.display='none';});
  },{title:'Delete Clients',okText:'Delete '+ids.length});
}

function renderReportCards(){
  var yr=new Date().getFullYear();
  function age(c){if(!c.f_dob)return null;var y=c.f_dob.split('/')[2]||c.f_dob.split('-')[0];return yr-parseInt(y);}
  var cards=[
    {num:clients.length,label:'Total Clients',filter:'all'},
    {num:clients.filter(function(c){var a=age(c);return a===64||a===65;}).length,label:'Turning 65 This Year',filter:'turning65'},
    {num:clients.filter(function(c){var a=age(c);return a===25||a===26;}).length,label:'Turning 26 This Year',filter:'turning26'},
    {num:clients.filter(function(c){return c.f_planType==='ACA/Marketplace';}).length,label:'ACA/Marketplace',filter:'aca'},
    {num:clients.filter(function(c){return c.f_planType==='Medicare';}).length,label:'Medicare',filter:'medicare'},
    {num:'$'+Math.round(clients.reduce(function(s,c){return s+(parseFloat((c.f_premium||'').replace(/[^0-9.]/g,''))||0);},0)).toLocaleString(),label:'Total Monthly Premium',filter:'all'}
  ];
  var container=document.getElementById('reportCards');container.innerHTML='';
  cards.forEach(function(card){var div=document.createElement('div');div.className='report-card';div.innerHTML='<div class="num">'+card.num+'</div><div class="lbl">'+card.label+'</div>';div.addEventListener('click',function(){runReport(card.filter,card.label);});container.appendChild(div);});
  renderTopReferrersCard();
}
/* Show a small stacked list of top referrers with counts.
   Sits next to the report cards so it's visible on the Reports page. */
function renderTopReferrersCard(){
  var container=document.getElementById('reportCards');if(!container)return;
  var refs=getReferrerCounts();
  if(!refs.length)return;
  var top=refs.slice(0,5);
  var div=document.createElement('div');
  div.className='report-card';
  div.style.cssText='min-width:200px;';
  var items=top.map(function(r){
    return '<div style="display:flex;justify-content:space-between;font-size:12px;padding:3px 0;border-bottom:1px solid #f0f3f7;"><span>'+r.name+'</span><strong style="color:var(--accent);">'+r.count+'</strong></div>';
  }).join('');
  div.innerHTML='<div class="lbl" style="margin-bottom:6px;font-weight:700;color:var(--text);text-transform:uppercase;letter-spacing:0.4px;">Top Referrers</div>'+items+
    (refs.length>5?'<div style="font-size:10px;color:var(--text-muted);margin-top:4px;">+ '+(refs.length-5)+' more</div>':'');
  container.appendChild(div);
}
function runReport(filter,title){
  var yr=new Date().getFullYear();function age(c){if(!c.f_dob)return null;var y=c.f_dob.split('/')[2]||c.f_dob.split('-')[0];return yr-parseInt(y);}
  var data=clients;
  if(filter==='turning65')data=clients.filter(function(c){var a=age(c);return a===64||a===65;});
  else if(filter==='turning26')data=clients.filter(function(c){var a=age(c);return a===25||a===26;});
  else if(filter==='aca')data=clients.filter(function(c){return c.f_planType==='ACA/Marketplace';});
  else if(filter==='medicare')data=clients.filter(function(c){return c.f_planType==='Medicare';});
  currentReportData=data;
  document.getElementById('reportTitle').textContent=title+' ('+data.length+')';
  document.getElementById('reportTableHead').innerHTML='<tr><th>Name</th><th>DOB</th><th>Phone</th><th>Email</th><th>Plan</th><th>Premium</th><th>Agent</th></tr>';
  var tbody=document.getElementById('reportTableBody');tbody.innerHTML='';
  data.forEach(function(c){var tr=document.createElement('tr');tr.innerHTML='<td>'+(c.f_firstName||'')+' '+(c.f_lastName||'')+'</td><td>'+(c.f_dob||'')+'</td><td>'+(c.f_phone||'')+'</td><td>'+(c.f_email||'')+'</td><td>'+(c.f_planType||'')+'</td><td>'+(c.f_premium?'$'+c.f_premium:'')+'</td><td>'+(c.f_agent||'')+'</td>';tbody.appendChild(tr);});
  document.getElementById('reportResult').style.display='block';
}
function exportReportExcel(){
  var rows=[['Name','DOB','Phone','Email','Plan Type','Plan Name','Premium','Agent']];
  currentReportData.forEach(function(c){rows.push([(c.f_firstName||'')+' '+(c.f_lastName||''),c.f_dob||'',c.f_phone||'',c.f_email||'',c.f_planType||'',c.f_planName||'',c.f_premium||'',c.f_agent||'']);});
  dlXLSX(rows,'report.xlsx');
}
function exportExcel(){
  var rows=[['First Name','Last Name','DOB','Phone','Email','Plan Type','Plan Name','Premium','Subsidy','Agent','Lead Source','Renewed']];
  clients.forEach(function(c){rows.push([c.f_firstName||'',c.f_lastName||'',c.f_dob||'',c.f_phone||'',c.f_email||'',c.f_planType||'',c.f_planName||'',c.f_premium||'',c.f_subsidy||'',c.f_agent||'',c.f_leadSource||'',c.f_renewed||'']);});
  dlXLSX(rows,'clients.xlsx');
}
function dlXLSX(rows,filename){
  var ws=XLSX.utils.aoa_to_sheet(rows);
  var wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Data');
  XLSX.writeFile(wb,filename);
}

var CRM_IMPORT_FIELDS=[
  {key:'f_firstName',label:'First Name'},{key:'f_lastName',label:'Last Name'},
  {key:'f_dob',label:'Date of Birth'},{key:'f_gender',label:'Gender'},{key:'f_ssn',label:'SSN'},
  {key:'f_phone',label:'Phone'},{key:'f_email',label:'Email'},{key:'f_resAddress',label:'Address'},
  {key:'f_resCity',label:'City'},{key:'f_resSt',label:'State'},{key:'f_resZip',label:'Zip'},
  {key:'f_planName',label:'Plan Name'},{key:'f_planType',label:'Plan Type'},{key:'f_premium',label:'Premium'},
  {key:'f_subsidy',label:'Subsidy'},{key:'f_agent',label:'Agent'},{key:'f_leadSource',label:'Lead Source'},
  {key:'f_healthEffective',label:'Health Effective'},{key:'f_totalMonthly',label:'Total Monthly'},
  {key:'f_medicareNum',label:'Medicare #'},{key:'f_medicaid',label:'Medicaid #'},{key:'f_notes',label:'Notes'}
];
function handleCSV(event){
  var file=event.target.files[0];if(!file)return;
  var reader=new FileReader();
  reader.onload=function(e){
    var lines=e.target.result.split('\n').filter(function(l){return l.trim();});
    csvHeaders=lines[0].split(',').map(function(h){return h.trim().replace(/"/g,'');});
    csvData=lines.slice(1).map(function(line){var vals=line.split(',').map(function(v){return v.trim().replace(/"/g,'');});var obj={};csvHeaders.forEach(function(h,i){obj[h]=vals[i]||'';});return obj;});
    var tbody=document.getElementById('mappingBody');tbody.innerHTML='';
    CRM_IMPORT_FIELDS.forEach(function(f){
      var tr=document.createElement('tr');
      var opts='<option value="">-- Skip --</option>'+csvHeaders.map(function(h){var match=h.toLowerCase().replace(/[^a-z]/g,'').includes(f.label.toLowerCase().replace(/[^a-z]/g,'').substr(0,4));return '<option value="'+h+'" '+(match?'selected':'')+'>'+h+'</option>';}).join('');
      tr.innerHTML='<td>'+f.label+'</td><td><select id="map_'+f.key+'">'+opts+'</select></td>';
      tbody.appendChild(tr);
    });
    document.getElementById('mappingSection').style.display='block';
  };
  reader.readAsText(file);
}
function importClients(){
  var imported=0;
  var promises=csvData.map(function(row){
    var data={};CRM_IMPORT_FIELDS.forEach(function(f){var col=document.getElementById('map_'+f.key);if(col&&col.value&&row[col.value]!==undefined)data[f.key]=row[col.value];});
    data.f_agent=data.f_agent||'Thomas Jaboro';
    return saveClientAPI(data,null).then(function(){imported++;});
  });
  Promise.all(promises).then(function(){document.getElementById('importStatus').textContent='Imported '+imported+' clients!';loadClients();}).catch(function(e){document.getElementById('importStatus').textContent='Error: '+e;});
}

// CARRIER MANAGEMENT
function loadCarriers(){
  var saved=localStorage.getItem('crmCarriers');
  carriers=saved?JSON.parse(saved):[];
}
function saveCarriers(){
  localStorage.setItem('crmCarriers',JSON.stringify(carriers));
}
function addCarrier(){
  var name=prompt('Enter carrier name:');
  if(!name)return;
  carriers.push({name:name,contact:'',phone:'',email:''});
  saveCarriers();
  renderCarriers();
}
function renderCarriers(){
  var container=document.getElementById('carrierList');
  if(!container)return;
  // One-time cleanup: purge any auto-seeded carriers (marked by an `availability`
  // array) that don't have manually-entered contact info. User-added carriers
  // stay put.
  if(!localStorage.getItem('crm_carriers_seed_purged')){
    carriers=(carriers||[]).filter(function(c){
      var wasSeeded=Array.isArray(c.availability);
      var hasUserData=(c.contact||'').trim()||(c.phone||'').trim()||(c.email||'').trim();
      return !wasSeeded||hasUserData;
    });
    saveCarriers();
    localStorage.removeItem('crm_carriers_seeded');
    localStorage.removeItem('crm_carriers_seed_version');
    localStorage.setItem('crm_carriers_seed_purged','1');
  }
  container.innerHTML='';
  if(carriers.length===0){
    container.innerHTML='<p style="text-align:center;padding:20px;color:#999;">No carriers yet. Click "+ Add Carrier" to add one.</p>';
    return;
  }
  carriers.forEach(function(c,idx){
    var div=document.createElement('div');
    div.style.cssText='display:grid;grid-template-columns:1.5fr 1fr 0.8fr 30px;gap:8px;align-items:end;margin-bottom:8px;';
    div.innerHTML='<div class="field"><label>Carrier Name</label><input value="'+(c.name||'').replace(/"/g,'&quot;')+'" oninput="carriers['+idx+'].name=this.value;saveCarriers();"></div>'+
      '<div class="field"><label>Contact Person</label><input value="'+(c.contact||'').replace(/"/g,'&quot;')+'" oninput="carriers['+idx+'].contact=this.value;saveCarriers();"></div>'+
      '<div class="field"><label>Phone</label><input value="'+(c.phone||'').replace(/"/g,'&quot;')+'" oninput="carriers['+idx+'].phone=this.value;saveCarriers();"></div>'+
      '<button type="button" class="icon-btn" onclick="confirmRemoveCarrier('+idx+')" title="Remove"><svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="1.8" fill="none" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button>';
    container.appendChild(div);
  });
}
function confirmRemoveCarrier(idx){
  var c=carriers[idx];if(!c)return;
  showConfirm('Remove carrier "'+c.name+'"?',function(){
    carriers.splice(idx,1);saveCarriers();renderCarriers();
  },{title:'Remove Carrier',okText:'Remove'});
}

function loadCarriersToSelect(){
  var sel=document.getElementById('f_planCarrier');
  if(sel){
    sel.innerHTML='<option value="">Select</option>';
    carriers.forEach(function(c){
      var opt=document.createElement('option');
      opt.value=c.name;
      opt.textContent=c.name;
      sel.appendChild(opt);
    });
  }
}

// CUSTOM REPORT BUILDER

// ===================== HEIGHT FORMATTER =====================
function fmtHeight(el){
  var raw=el.value.replace(/[^0-9]/g,'');
  if(!raw){return;}
  if(raw.length===1){el.value=raw;}
  else if(raw.length===2){
    // Could be feet only (e.g. "50" → "5'0") or partial
    el.value=raw[0]+"'"+raw[1];
  } else if(raw.length===3){
    el.value=raw[0]+"'"+raw.slice(1);
  } else if(raw.length>=4){
    el.value=raw[0]+"'"+raw.slice(1,3)+'"';
  }
}

// ===================== RECENT RECORDS =====================
var _recentRecords=[];
function loadRecentRecords(){try{_recentRecords=JSON.parse(localStorage.getItem('crm_recent')||'[]');}catch(e){_recentRecords=[];}}
function saveRecentRecords(){localStorage.setItem('crm_recent',JSON.stringify(_recentRecords));}
loadRecentRecords();
function trackRecentRecord(id,c){
  var name=((c.f_firstName||'')+' '+(c.f_lastName||'')).trim()||'Unknown';
  _recentRecords=_recentRecords.filter(function(r){return String(r.id)!==String(id);});
  _recentRecords.unshift({id:id,name:name,planType:c.f_planType||'',agent:c.f_agent||'',accessed:new Date().toISOString()});
  if(_recentRecords.length>20)_recentRecords=_recentRecords.slice(0,20);
  saveRecentRecords();
}
function renderRecentRecords(){
  var el=document.getElementById('recentRecordsList');
  var empty=document.getElementById('recentRecordsEmpty');
  if(!el)return;
  if(!_recentRecords.length){el.innerHTML='';empty.style.display='block';return;}
  empty.style.display='none';
  el.innerHTML='';
  _recentRecords.forEach(function(r,i){
    var when=new Date(r.accessed);
    var whenStr=when.toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'})+' '+when.toLocaleTimeString('en-US',{hour:'numeric',minute:'2-digit'});
    var div=document.createElement('div');
    div.style.cssText='display:flex;align-items:center;gap:12px;padding:10px 12px;border-bottom:1px solid #f0f0f0;cursor:pointer;transition:background 0.1s;';
    div.onmouseover=function(){this.style.background='#f0f4f8';};
    div.onmouseout=function(){this.style.background='';};
    div.innerHTML=
      '<div style="width:24px;height:24px;border-radius:50%;background:#1a3a5c;color:#fff;font-size:11px;font-weight:bold;display:flex;align-items:center;justify-content:center;flex-shrink:0;">'+(i+1)+'</div>'+
      '<div style="flex:1;">'+
        '<div style="font-weight:600;font-size:13px;color:#1a3a5c;">'+r.name+'</div>'+
        '<div style="font-size:11px;color:#666;margin-top:2px;">'+(r.planType?'<span style="background:#dbeafe;color:#1a3a5c;padding:1px 6px;border-radius:8px;font-size:10px;margin-right:6px;">'+r.planType+'</span>':'')+(r.agent||'')+'</div>'+
      '</div>'+
      '<div style="font-size:10px;color:#999;white-space:nowrap;">'+whenStr+'</div>'+
      '<button class="btn btn-blue" style="padding:4px 10px;font-size:11px;">Open</button>';
    div.querySelector('.btn').addEventListener('click',function(e){e.stopPropagation();editClient(r.id);});
    div.addEventListener('click',function(){editClient(r.id);});
    el.appendChild(div);
  });
}
function clearRecentRecords(){
  showConfirm('Clear recent records history?',function(){
    _recentRecords=[];saveRecentRecords();renderRecentRecords();
  },{title:'Clear History',okText:'Clear'});
}

// ===================== TODO CLIENT AUTOCOMPLETE =====================
function todoClientAC(el){
  var v=el.value.toLowerCase().trim();
  var list=document.getElementById('todoClientList');
  document.getElementById('todoClientId').value='';
  if(!v||v.length<2){list.style.display='none';return;}
  var matches=clients.filter(function(c){
    var name=((c.f_firstName||'')+' '+(c.f_lastName||'')).toLowerCase();
    return name.includes(v);
  }).slice(0,8);
  if(!matches.length){list.style.display='none';return;}
  list.innerHTML='';
  matches.forEach(function(c){
    var name=((c.f_firstName||'')+' '+(c.f_lastName||'')).trim();
    var d=document.createElement('div');
    d.innerHTML='<strong>'+name+'</strong><span style="font-size:10px;color:#666;margin-left:6px;">'+(c.f_planType||'')+'</span>';
    d.style.cssText='padding:6px 8px;cursor:pointer;';
    d.addEventListener('mousedown',function(e){
      e.preventDefault();
      el.value=name;
      document.getElementById('todoClientId').value=c._id;
      list.style.display='none';
    });
    list.appendChild(d);
  });
  list.style.display='block';
}

// ===================== TODO - updated saveTodo with client link =====================
function saveTodo(){
  var task=document.getElementById('todoTaskInput').value.trim();
  if(!task)return;
  var due=document.getElementById('todoDueInput').value;
  var priority=document.getElementById('todoPriorityInput').value;
  var clientId=document.getElementById('todoClientId').value||'';
  var clientName=document.getElementById('todoClientInput').value.trim()||'';
  _todos.unshift({id:Date.now(),task:task,due:due,priority:priority,done:false,created:new Date().toISOString(),clientId:clientId,clientName:clientId?clientName:''});
  saveTodos();
  document.getElementById('todoAddSection').style.display='none';
  document.getElementById('todoClientInput').value='';
  document.getElementById('todoClientId').value='';
  renderTodos();
}

// ===================== ADVANCED SEARCH - add create date filter =====================
function runAdvSearch(){
  var fn=(document.getElementById('as_firstName').value||'').toLowerCase().trim();
  var ln=(document.getElementById('as_lastName').value||'').toLowerCase().trim();
  var dobStart=document.getElementById('as_dobStart').value;
  var dobEnd=document.getElementById('as_dobEnd').value;
  var agent=document.getElementById('as_agent').value;
  var gender=document.getElementById('as_gender').value;
  var tobacco=document.getElementById('as_tobacco').value;
  var marital=document.getElementById('as_marital').value;
  var planType=document.getElementById('as_planType').value;
  var carrier=document.getElementById('as_carrier').value;
  var level=document.getElementById('as_level').value;
  var pType=document.getElementById('as_type').value;
  var effStart=document.getElementById('as_effStart').value;
  var effEnd=document.getElementById('as_effEnd').value;
  var premMin=parseFloat(document.getElementById('as_premMin').value)||0;
  var premMax=parseFloat(document.getElementById('as_premMax').value)||Infinity;
  var state=(document.getElementById('as_state').value||'').toUpperCase().trim();
  var zip=(document.getElementById('as_zip').value||'').trim();
  var city=(document.getElementById('as_city').value||'').toLowerCase().trim();
  var county=(document.getElementById('as_county').value||'').toLowerCase().trim();
  var email=(document.getElementById('as_email').value||'').toLowerCase().trim();
  var leadSource=document.getElementById('as_leadSource').value;
  var renewed=document.getElementById('as_renewed').value;
  var leadStart=document.getElementById('as_leadStart').value;
  var leadEnd=document.getElementById('as_leadEnd').value;
  var createStart=document.getElementById('as_createStart').value;
  var createEnd=document.getElementById('as_createEnd').value;
  var submittedBy=(document.getElementById('as_submittedBy').value||'').toLowerCase().trim();
  var ageGroup=document.getElementById('as_ageGroup').value;
  var medicare=document.getElementById('as_medicare').value;
  var medicaid=document.getElementById('as_medicaid').value;
  var medication=(document.getElementById('as_medication').value||'').toLowerCase().trim();

  var results=clients.filter(function(c){
    if(fn&&!(c.f_firstName||'').toLowerCase().includes(fn))return false;
    if(ln&&!(c.f_lastName||'').toLowerCase().includes(ln))return false;
    if(agent&&c.f_agent!==agent)return false;
    if(gender&&c.f_gender!==gender)return false;
    if(tobacco&&c.f_tobacco!==tobacco)return false;
    if(marital&&c.f_marital!==marital)return false;
    if(planType&&c.f_planType!==planType)return false;
    if(carrier&&c.f_planCarrier!==carrier)return false;
    if(level&&c.f_level!==level)return false;
    if(pType&&c.f_type!==pType)return false;
    if(state&&(c.f_resSt||'').toUpperCase()!==state)return false;
    if(zip&&(c.f_resZip||'')!==zip)return false;
    if(city&&!(c.f_resCity||'').toLowerCase().includes(city))return false;
    if(county&&!(c.f_resCounty||'').toLowerCase().includes(county))return false;
    if(email&&!(c.f_email||'').toLowerCase().includes(email))return false;
    if(leadSource&&c.f_leadSource!==leadSource)return false;
    if(renewed&&c.f_renewed!==renewed)return false;
    if(submittedBy&&!(c.f_submittedBy||'').toLowerCase().includes(submittedBy))return false;
    if(medicare==='yes'&&!c.f_hasMedicare)return false;
    if(medicare==='no'&&c.f_hasMedicare)return false;
    if(medicaid==='yes'&&!c.f_hasMedicaid)return false;
    if(medicaid==='no'&&c.f_hasMedicaid)return false;
    if(premMin>0||premMax<Infinity){var prem=parseFloat((c.f_premium||'').replace(/[^0-9.]/g,''))||0;if(prem<premMin||prem>premMax)return false;}
    if(dobStart&&c.f_dob&&c.f_dob<dobStart)return false;
    if(dobEnd&&c.f_dob&&c.f_dob>dobEnd)return false;
    if(effStart&&c.f_healthEffective&&c.f_healthEffective<effStart)return false;
    if(effEnd&&c.f_healthEffective&&c.f_healthEffective>effEnd)return false;
    if(leadStart&&c.f_leadDate&&c.f_leadDate<leadStart)return false;
    if(leadEnd&&c.f_leadDate&&c.f_leadDate>leadEnd)return false;
    if(createStart&&c.f_date&&c.f_date<createStart)return false;
    if(createEnd&&c.f_date&&c.f_date>createEnd)return false;
    if(ageGroup){
      var age=calcClientAge(c);if(age===null)return false;
      if(ageGroup==='turning65'&&!(age===64||age===65))return false;
      if(ageGroup==='turning64'&&age!==63&&age!==64)return false;
      if(ageGroup==='turning26'&&!(age===25||age===26))return false;
      if(ageGroup==='under18'&&age>=18)return false;
      if(ageGroup==='18to26'&&(age<18||age>26))return false;
      if(ageGroup==='26to64'&&(age<26||age>64))return false;
      if(ageGroup==='over64'&&age<=64)return false;
    }
    if(medication){var meds=c.meds||[];var found=meds.some(function(m){return(m.name||'').toLowerCase().includes(medication);});if(!found)return false;}
    return true;
  });

  _advSearchResults=results;
  var tbody=document.getElementById('advSearchBody');tbody.innerHTML='';
  document.getElementById('advSearchResults').style.display=results.length>0?'block':'none';
  document.getElementById('advSearchEmpty').style.display=results.length===0?'block':'none';
  document.getElementById('advSearchCount').textContent=results.length+' client'+(results.length!==1?'s':'')+' found';
  document.getElementById('advSearchExportBtn').style.display=results.length>0?'inline-block':'none';
  results.forEach(function(c){
    var tr=document.createElement('tr');
    var age=calcClientAge(c);
    tr.innerHTML='<td><span class="client-name-link" onclick="editClient(\''+c._id+'\')">'+((c.f_firstName||'')+' '+(c.f_lastName||'')).trim()+'</span></td>'+
      '<td>'+(c.f_dob||'')+'</td><td>'+(age!==null?age:'')+'</td><td>'+(c.f_phone||'')+'</td><td>'+(c.f_email||'')+'</td>'+
      '<td>'+(c.f_planType||'')+'</td><td>'+(c.f_planCarrier||'')+'</td><td>'+(c.f_premium||'')+'</td>'+
      '<td>'+(c.f_agent||'')+'</td><td>'+(c.f_resSt||'')+'</td><td>'+(c.f_renewed||'')+'</td>';
    tbody.appendChild(tr);
  });
}

// ===================== SETTINGS EXTRAS =====================
var _settingsPlanTypes=[];
var _settingsProjectCodes=[];
function loadSettingsExtras(){
  try{_settingsPlanTypes=JSON.parse(localStorage.getItem('crm_plan_types')||'["ACA/Marketplace","Medicare","Medicaid","Short Term","Ancillary"]');}catch(e){_settingsPlanTypes=["ACA/Marketplace","Medicare","Medicaid","Short Term","Ancillary"];}
  try{_settingsProjectCodes=JSON.parse(localStorage.getItem('crm_project_codes')||'[]');}catch(e){_settingsProjectCodes=[];}
}
loadSettingsExtras();
function populateDefaultAgentSelect(){
  var sel=document.getElementById('settingsDefaultAgent');if(!sel)return;
  var cur=localStorage.getItem('crm_default_agent')||_settingsAgents[0]||'';
  sel.innerHTML='';
  _settingsAgents.forEach(function(a){var o=document.createElement('option');o.value=a;o.textContent=a;if(a===cur)o.selected=true;sel.appendChild(o);});
}
function saveDefaultAgent(){
  var sel=document.getElementById('settingsDefaultAgent');if(!sel)return;
  localStorage.setItem('crm_default_agent',sel.value);
}
function addPlanTypeSetting(){
  var v=document.getElementById('newPlanTypeInput').value.trim();if(!v)return;
  if(_settingsPlanTypes.indexOf(v)!==-1){toast('Already exists.','info');return;}
  _settingsPlanTypes.push(v);localStorage.setItem('crm_plan_types',JSON.stringify(_settingsPlanTypes));
  document.getElementById('newPlanTypeInput').value='';renderSettings();
}
function removePlanTypeSetting(i){
  showConfirm('Remove this plan type?',function(){
    _settingsPlanTypes.splice(i,1);localStorage.setItem('crm_plan_types',JSON.stringify(_settingsPlanTypes));renderSettings();
  },{title:'Remove',okText:'Remove'});
}
function addProjectCodeSetting(){
  var v=document.getElementById('newProjectCodeInput').value.trim();if(!v)return;
  if(_settingsProjectCodes.indexOf(v)!==-1){toast('Already exists.','info');return;}
  _settingsProjectCodes.push(v);localStorage.setItem('crm_project_codes',JSON.stringify(_settingsProjectCodes));
  document.getElementById('newProjectCodeInput').value='';renderSettings();
}
function removeProjectCodeSetting(i){
  showConfirm('Remove this code?',function(){
    _settingsProjectCodes.splice(i,1);localStorage.setItem('crm_project_codes',JSON.stringify(_settingsProjectCodes));renderSettings();
  },{title:'Remove',okText:'Remove'});
}
function exportFullBackup(){
  if(!clients.length){toast('No clients to export.','error');return;}
  var rows=[['First Name','Last Name','DOB','Phone','Email','Plan Type','Plan Name','Carrier','Premium','Subsidy','Total Monthly','App Fee','Agent','Lead Source','Renewed','State','City','ZIP','County','Medicare','Medicaid','Notes','App Date']];
  clients.forEach(function(c){
    rows.push([c.f_firstName||'',c.f_lastName||'',c.f_dob||'',c.f_phone||'',c.f_email||'',c.f_planType||'',c.f_planName||'',c.f_planCarrier||'',c.f_premium||'',c.f_subsidy||'',c.f_totalMonthly||'',c.f_appFee||'',c.f_agent||'',c.f_leadSource||'',c.f_renewed||'',c.f_resSt||'',c.f_resCity||'',c.f_resZip||'',c.f_resCounty||'',c.f_hasMedicare?'Yes':'No',c.f_hasMedicaid?'Yes':'No',c.f_notes||'',c.f_date||'']);
  });
  dlXLSX(rows,'liberty_crm_backup_'+new Date().toISOString().split('T')[0]+'.xlsx');
}
function clearPreviewData(){
  showConfirm('This will permanently delete all client data saved in Preview Mode. Are you sure?',function(){
    localStorage.removeItem('crm_preview');
    clients=[];renderClientTable(clients);renderReportCards();
    toast('Preview data cleared.','info');
  },{title:'Clear Preview Data',okText:'Delete All'});
}

// ===================== DOCUMENT UPLOAD =====================
var _clientDocs=[];
function loadClientDocs(clientId){
  var sec=document.getElementById('clientDocsSection');
  if(!sec)return;
  sec.innerHTML='<div class="form-section-title">&#128196; Client Documents</div><p style="font-size:11px;color:#999;">Loading...</p>';
  fetch(API_BASE+'/documents?clientType=health&clientId='+clientId,{headers:apiHeaders()})
  .then(function(r){return r.json();})
  .then(function(docs){_clientDocs=docs||[];renderClientDocs(clientId,docs);})
  .catch(function(){_clientDocs=[];renderClientDocs(clientId,[]);});
}
function renderClientDocs(clientId,docs){
  var sec=document.getElementById('clientDocsSection');if(!sec)return;
  sec.innerHTML='<div class="form-section-title">&#128196; Client Documents <span style="font-size:10px;font-weight:normal;color:#999;">(SSN card, driver\'s license, etc.)</span></div>';
  if(docs&&docs.length){
    var ul=document.createElement('div');ul.style.cssText='margin-bottom:10px;';
    docs.forEach(function(d){
      var row=document.createElement('div');
      row.style.cssText='display:flex;align-items:center;gap:8px;padding:5px 0;border-bottom:1px solid #f0f0f0;font-size:12px;';
      var ext=(d.name||'').split('.').pop().toLowerCase();
      var icon=(['jpg','jpeg','png','gif','webp'].indexOf(ext)>=0)?'&#128247;':(['pdf'].indexOf(ext)>=0)?'&#128196;':'&#128196;';
      var kb=d.size?Math.round(d.size/1024)+'KB':'';
      row.innerHTML=icon+' <a href="'+d.url+'" target="_blank" style="flex:1;color:#1a3a5c;text-decoration:none;word-break:break-all;">'+d.name+'</a>'+
        '<span style="color:#999;font-size:10px;">'+kb+'</span>'+
        '<button class="btn btn-red" style="padding:2px 8px;font-size:10px;" onclick="deleteClientDoc(\''+clientId+'\',\''+encodeURIComponent(d.name)+'\')">✕</button>';
      ul.appendChild(row);
    });
    sec.appendChild(ul);
  } else {
    var emp=document.createElement('p');emp.style.cssText='font-size:11px;color:#999;margin-bottom:8px;';emp.textContent='No documents uploaded yet.';sec.appendChild(emp);
  }
  var upRow=document.createElement('div');upRow.style.cssText='display:flex;align-items:center;gap:8px;margin-top:6px;flex-wrap:wrap;';
  upRow.innerHTML='<input type="file" id="docFileInput" accept=".pdf,.jpg,.jpeg,.png,.doc,.docx" style="font-size:11px;flex:1;min-width:0;" onchange="uploadClientDoc(\''+clientId+'\')" multiple>'+
    '<span id="docUploadStatus" style="font-size:11px;color:#666;"></span>';
  sec.appendChild(upRow);
}
function uploadClientDoc(clientId){
  var input=document.getElementById('docFileInput');
  if(!input||!input.files||!input.files.length)return;
  var status=document.getElementById('docUploadStatus');
  var files=Array.from(input.files);
  status.textContent='Uploading '+files.length+' file(s)...';
  var promises=files.map(function(file){
    var fd=new FormData();
    fd.append('file',file);
    fd.append('clientType','health');
    fd.append('clientId',clientId);
    return fetch(API_BASE+'/documents',{method:'POST',headers:authUploadHeaders(),body:fd});
  });
  var fileNames=files.map(function(f){return f.name;}).join(', ');
  Promise.all(promises)
  .then(function(){
    aiTrack('DocumentUploaded',{clientType:'health',clientId:clientId,files:fileNames});
    status.textContent='';input.value='';loadClientDocs(clientId);
  })
  .catch(function(e){status.textContent='Upload failed: '+e;});
}
function deleteClientDoc(clientId,encodedName){
  showConfirm('Delete this document?',function(){
    fetch(API_BASE+'/documents?clientType=health&clientId='+clientId+'&name='+encodedName,{method:'DELETE',headers:apiHeaders()})
    .then(function(){loadClientDocs(clientId);})
    .catch(function(e){toast('Delete failed: '+e,'error');});
  },{title:'Delete Document',okText:'Delete'});
}

// initMSAL called below
var _addrTimer={};
function addrAC(el,prefix){
  var v=el.value.trim();
  var listId=prefix+'AddrList';
  var list=document.getElementById(listId);
  if(!list)return;
  if(v.length<3){list.style.display='none';return;}
  clearTimeout(_addrTimer[prefix]);
  _addrTimer[prefix]=setTimeout(function(){
    fetch('https://nominatim.openstreetmap.org/search?format=json&addressdetails=1&limit=6&countrycodes=us&q='+encodeURIComponent(v),{headers:{'Accept-Language':'en-US,en'}})
    .then(function(r){return r.json();})
    .then(function(results){
      list.innerHTML='';
      if(!results||!results.length){list.style.display='none';return;}
      results.forEach(function(r){
        var addr=r.address||{};
        var street=(addr.house_number?addr.house_number+' ':'')+( addr.road||addr.pedestrian||'');
        var city=addr.city||addr.town||addr.village||addr.hamlet||'';
        var state=addr.state||'';var zip=addr.postcode||'';
        if(!street)return;
        var d=document.createElement('div');
        d.innerHTML='<div class="addr-main">'+street+'</div><div class="addr-sub">'+city+(state?', '+state:'')+(zip?' '+zip:'')+'</div>';
        d.addEventListener('mousedown',function(e){
          e.preventDefault();
          el.value=street;
          list.style.display='none';
          if(zip){
            var zipEl=document.getElementById('f_'+prefix+'Zip');
            var cityEl=document.getElementById('f_'+prefix+'City');
            var stEl=document.getElementById('f_'+prefix+'St');
            if(zipEl)zipEl.value=zip;
            if(cityEl)cityEl.value=city;
            if(stEl)stEl.value=addr.state_code||addr['ISO3166-2-lvl4']||state;
            if(zip.replace(/\D/g,'').length===5)restoreCounty(zip,prefix,addr.county||'');
          }
        });
        list.appendChild(d);
      });
      if(list.children.length>0)list.style.display='block';else list.style.display='none';
    }).catch(function(){list.style.display='none';});
  },350);
}

// ===================== ADVANCED SEARCH (canonical - with create date) =====================
var _advSearchResults=[];
function populateAdvSearchCarriers(){
  var sel=document.getElementById('as_carrier');if(!sel)return;
  sel.innerHTML='<option value="">All</option>';
  carriers.forEach(function(c){var o=document.createElement('option');o.value=c.name;o.textContent=c.name;sel.appendChild(o);});
}
function calcClientAge(c){
  if(!c.f_dob)return null;
  var y=c.f_dob.split(/[-/]/);
  var yr=y[0].length===4?parseInt(y[0]):parseInt(y[2]);
  if(isNaN(yr))return null;
  return new Date().getFullYear()-yr;
}
function clearAdvSearch(){
  ['as_firstName','as_lastName','as_dobStart','as_dobEnd','as_effStart','as_effEnd','as_leadStart','as_leadEnd','as_createStart','as_createEnd','as_premMin','as_premMax','as_state','as_zip','as_city','as_county','as_email','as_medication','as_submittedBy'].forEach(function(id){var el=document.getElementById(id);if(el)el.value='';});
  ['as_agent','as_gender','as_tobacco','as_marital','as_planType','as_carrier','as_level','as_type','as_leadSource','as_renewed','as_ageGroup','as_medicare','as_medicaid'].forEach(function(id){var el=document.getElementById(id);if(el)el.value='';});
  document.getElementById('advSearchResults').style.display='none';
  document.getElementById('advSearchEmpty').style.display='none';
  document.getElementById('advSearchExportBtn').style.display='none';
  _advSearchResults=[];
}
/* Column list — order also determines Excel column order.
   Defaults reflect the most-used report shape. */
var REPORT_FIELDS=[
  ['name','Name',true],['dob','Birth Date',false],['age','Age',false],['gender','Gender',false],['marital','Marital',false],['tobacco','Tobacco',false],
  ['phone','Phone',true],['altPhone','Alt Phone',false],['email','Email',true],['email2','Email 2',false],
  ['address','Address',false],['city','City',false],['state','State',false],['zip','Zip',false],['county','County',false],
  ['planType','Plan Type',true],['planName','Plan Name',false],['carrier','Carrier',false],['level','Plan Level',false],['type','Network',false],['deductible','Deductible',false],
  ['premium','Premium',true],['subsidy','Subsidy',false],['totalMonthly','Total Monthly',false],['healthEffective','Health Effective',false],
  ['medicare','Medicare',false],['medicaid','Medicaid',false],
  ['agent','Agent',true],['leadSource','Lead Source',false],['leadDate','Lead Date',false],['renewed','Renewed',false],
  ['bankName','Bank Name',false],['accountType','Account Type',false],['routing','Routing',false],['account','Account Number',false],['accountName','Name on Account',false],
  ['cardType','Card Type',false],['cardNumber','Card Number',false],['cardExp','Card Exp',false],
  ['primaryEmployer','Primary Employer',false],['primaryIncome','Primary Income',false],['spouseEmployer','Spouse Employer',false],['spouseIncome','Spouse Income',false],
  ['otherIncome1','Other Income 1',false],['otherIncomeAmt1','Other Income Amt 1',false],['otherIncome2','Other Income 2',false],['otherIncomeAmt2','Other Income Amt 2',false],
  ['totalIncome','Total Income',false],['notes','Notes',false]
];
function reportFieldValue(c,f){
  switch(f){
    case 'name':return (c.f_firstName||'')+' '+(c.f_lastName||'');
    case 'age':return calcClientAge(c)||'';
    case 'address':return c.f_resAddress||'';
    case 'city':return c.f_resCity||'';
    case 'state':return c.f_resSt||'';
    case 'zip':return c.f_resZip||'';
    case 'county':return c.f_resCounty||'';
    case 'carrier':return c.f_planCarrier||'';
    case 'medicare':return c.f_hasMedicare?'Yes':'No';
    case 'medicaid':return c.f_hasMedicaid?'Yes':'No';
    default:return c['f_'+f]||'';
  }
}
function loadReportColumns(){
  try{
    var saved=JSON.parse(localStorage.getItem('crm_report_columns')||'null');
    if(saved){REPORT_FIELDS.forEach(function(f){var el=document.getElementById('rpt_'+f[0]);if(el)el.checked=saved.indexOf(f[0])!==-1;});}
  }catch(e){}
}
function saveReportColumns(){
  var picked=REPORT_FIELDS.filter(function(f){var el=document.getElementById('rpt_'+f[0]);return el&&el.checked;}).map(function(f){return f[0];});
  try{localStorage.setItem('crm_report_columns',JSON.stringify(picked));}catch(e){}
}
function toggleAdvColumns(){
  var body=document.getElementById('advColumnsBody'),tog=document.getElementById('advColumnsToggle');
  if(!body)return;
  var open=body.style.display!=='none';
  body.style.display=open?'none':'block';
  if(tog)tog.textContent=open?'(click to expand)':'(click to collapse)';
  if(!open)loadReportColumns();
}
function advColumnsSelectAll(v){REPORT_FIELDS.forEach(function(f){var el=document.getElementById('rpt_'+f[0]);if(el)el.checked=v;});saveReportColumns();}
function advColumnsResetDefault(){REPORT_FIELDS.forEach(function(f){var el=document.getElementById('rpt_'+f[0]);if(el)el.checked=f[2];});saveReportColumns();}
document.addEventListener('change',function(e){if(e.target&&e.target.id&&e.target.id.indexOf('rpt_')===0)saveReportColumns();});

function exportAdvSearchExcel(){
  if(!_advSearchResults.length)return;
  var chosen=REPORT_FIELDS.filter(function(f){var el=document.getElementById('rpt_'+f[0]);return el?el.checked:f[2];});
  if(!chosen.length){toast('Select at least one column in "Columns to Include"','error');return;}
  var rows=[chosen.map(function(f){return f[1];})];
  _advSearchResults.forEach(function(c){
    rows.push(chosen.map(function(f){return reportFieldValue(c,f[0]);}));
  });
  dlXLSX(rows,'advanced_search_results.xlsx');
}

// ===================== TO-DO LIST =====================
var _todos=[];
var _todoFilter='all';
function loadTodos(){try{_todos=JSON.parse(localStorage.getItem('crm_todos')||'[]');}catch(e){_todos=[];}}
function saveTodos(){localStorage.setItem('crm_todos',JSON.stringify(_todos));}
loadTodos();
function openAddTodo(){
  var s=document.getElementById('todoAddSection');
  s.style.display='block';
  document.getElementById('todoTaskInput').focus();
  document.getElementById('todoTaskInput').value='';
  document.getElementById('todoDueInput').value='';
  document.getElementById('todoPriorityInput').value='normal';
  document.getElementById('todoClientInput').value='';
  document.getElementById('todoClientId').value='';
}
function toggleTodo(id){
  var t=_todos.find(function(x){return x.id===id;});
  if(t)t.done=!t.done;
  saveTodos();renderTodos();
}
function deleteTodo(id){
  showConfirm('Delete this task?',function(){
    _todos=_todos.filter(function(x){return x.id!==id;});
    saveTodos();renderTodos();
  },{title:'Delete Task',okText:'Delete'});
}
function clearCompletedTodos(){
  var done=_todos.filter(function(x){return x.done;}).length;
  if(!done){toast('No completed tasks to clear.','info');return;}
  showConfirm('Remove '+done+' completed task'+(done!==1?'s':'')+'?',function(){
    _todos=_todos.filter(function(x){return !x.done;});
    saveTodos();renderTodos();
  },{title:'Clear Completed',okText:'Remove'});
}
function setTodoFilter(f){
  _todoFilter=f;
  ['All','Pending','High','Done'].forEach(function(n){var btn=document.getElementById('todoFilter'+n);if(btn)btn.style.cssText='font-size:11px;padding:4px 10px;';});
  var active=document.getElementById('todoFilter'+f.charAt(0).toUpperCase()+f.slice(1));
  if(active)active.style.cssText='font-size:11px;padding:4px 10px;background:#1a3a5c;color:#fff;border-color:#1a3a5c;';
  renderTodos();
}
function renderTodos(){
  var container=document.getElementById('todoListContainer');if(!container)return;
  var filtered=_todos.filter(function(t){
    if(_todoFilter==='pending')return !t.done;
    if(_todoFilter==='high')return t.priority==='high'&&!t.done;
    if(_todoFilter==='done')return t.done;
    return true;
  });
  var today=new Date().toISOString().split('T')[0];
  document.getElementById('todoEmpty').style.display=filtered.length===0?'block':'none';
  container.innerHTML='';
  filtered.forEach(function(t){
    var overdue=t.due&&t.due<today&&!t.done;
    var dueSoon=t.due&&t.due===today&&!t.done;
    var div=document.createElement('div');
    div.style.cssText='display:flex;align-items:center;gap:10px;padding:10px 12px;border-bottom:1px solid #f0f0f0;'+(t.done?'opacity:0.55;':'');
    var prioColor=t.priority==='high'?'#dc3545':t.priority==='low'?'#6c757d':'#1a3a5c';
    var prioLabel=t.priority==='high'?'🔴 High':t.priority==='low'?'🔵 Low':'⚪ Normal';
    var dueStr='';
    if(t.due){
      if(overdue)dueStr='<span style="color:#dc3545;font-size:10px;font-weight:bold;margin-left:6px;">⚠ Overdue: '+t.due+'</span>';
      else if(dueSoon)dueStr='<span style="color:#e67e22;font-size:10px;font-weight:bold;margin-left:6px;">⏰ Due Today</span>';
      else dueStr='<span style="color:#666;font-size:10px;margin-left:6px;">Due: '+t.due+'</span>';
    }
    var clientLink='';
    if(t.clientId&&t.clientName){
      clientLink='<span onclick="editClient(\''+t.clientId+'\')" style="font-size:10px;background:#dbeafe;color:#1a3a5c;padding:2px 8px;border-radius:10px;cursor:pointer;margin-left:6px;font-weight:600;" title="Open client record">&#128101; '+t.clientName+'</span>';
    }
    div.innerHTML=
      '<input type="checkbox" '+(t.done?'checked':'')+' style="width:16px;height:16px;cursor:pointer;flex-shrink:0;" onchange="toggleTodo('+t.id+')">'+
      '<div style="flex:1;">'+
        '<span style="font-size:13px;'+(t.done?'text-decoration:line-through;color:#999;':'')+'">'+(t.task)+'</span>'+dueStr+clientLink+
      '</div>'+
      '<span style="font-size:10px;color:'+prioColor+';font-weight:600;white-space:nowrap;">'+prioLabel+'</span>'+
      '<button class="btn btn-red" style="padding:2px 7px;font-size:11px;" onclick="deleteTodo('+t.id+')">✕</button>';
    container.appendChild(div);
  });
}

// Show client tasks on client form (called after setFormData)
function renderClientTodos(clientId){
  var existing=document.getElementById('clientTodoSection');
  if(existing)existing.remove();
  if(!clientId)return;
  var tasks=_todos.filter(function(t){return String(t.clientId)===String(clientId);});
  var formCard=document.querySelector('#viewForm .form-card');
  if(!formCard)return;
  var sec=document.createElement('div');sec.className='form-section';sec.id='clientTodoSection';
  var pendingCount=tasks.filter(function(t){return !t.done;}).length;
  sec.innerHTML='<div class="form-section-title">&#9989; Tasks for this Client <span style="font-weight:normal;color:#999;">('+tasks.length+' total, '+pendingCount+' pending)</span></div>';
  if(!tasks.length){
    sec.innerHTML+='<p style="font-size:11px;color:#999;">No tasks linked to this client. <span style="color:#1a3a5c;cursor:pointer;text-decoration:underline;" onclick="openAddTodoForClient(\''+clientId+'\')">Add one</span></p>';
  } else {
    var ul=document.createElement('div');
    tasks.forEach(function(t){
      var today=new Date().toISOString().split('T')[0];
      var overdue=t.due&&t.due<today&&!t.done;
      var row=document.createElement('div');
      row.style.cssText='display:flex;align-items:center;gap:8px;padding:5px 0;border-bottom:1px solid #f5f5f5;font-size:12px;';
      row.innerHTML='<input type="checkbox" '+(t.done?'checked':'')+' onchange="toggleTodo('+t.id+');renderClientTodos(\''+clientId+'\')" style="cursor:pointer;">'+
        '<span style="flex:1;'+(t.done?'text-decoration:line-through;color:#999;':'')+'">'+(t.task)+'</span>'+
        (t.due?'<span style="font-size:10px;color:'+(overdue?'#dc3545':'#666')+';">'+t.due+'</span>':'')+
        (t.priority==='high'?'<span style="font-size:10px;color:#dc3545;font-weight:bold;">High</span>':'')+
        '<button class="btn btn-red" style="padding:1px 5px;font-size:10px;" onclick="deleteTodo('+t.id+');renderClientTodos(\''+clientId+'\')">✕</button>';
      ul.appendChild(row);
    });
    sec.appendChild(ul);
    var addBtn=document.createElement('button');
    addBtn.className='btn add-row-btn';addBtn.style.marginTop='8px';
    addBtn.textContent='+ Add Task';
    addBtn.onclick=function(){openAddTodoForClient(clientId);};
    sec.appendChild(addBtn);
  }
  // Insert before the Notes section (last form-section before form-actions)
  var actions=formCard.querySelector('.form-actions');
  formCard.insertBefore(sec,actions);
}
function openAddTodoForClient(clientId){
  var c=clients.find(function(x){return String(x._id)===String(clientId);});
  var name=c?((c.f_firstName||'')+' '+(c.f_lastName||'')).trim():'';
  showView('todo');
  openAddTodo();
  document.getElementById('todoClientId').value=clientId;
  document.getElementById('todoClientInput').value=name;
}

// ===================== SETTINGS =====================
var _settingsAgents=[];
var _settingsLeadSources=[];
var _settingsRenewals=[];
function loadSettings(){
  try{_settingsAgents=JSON.parse(localStorage.getItem('crm_agents')||'["Thomas Jaboro","Paul Jaboro Jr."]');}catch(e){_settingsAgents=["Thomas Jaboro","Paul Jaboro Jr."];}
  try{_settingsLeadSources=JSON.parse(localStorage.getItem('crm_lead_sources')||'["Insurance Quotes","Datalot","Smart Financial","Referral","Other"]');}catch(e){_settingsLeadSources=["Insurance Quotes","Datalot","Smart Financial","Referral","Other"];}
  try{_settingsRenewals=JSON.parse(localStorage.getItem('crm_renewals')||'["2026 Renewed","Not Renewed"]');}catch(e){_settingsRenewals=["2026 Renewed","Not Renewed"];}
}
function saveSettings(){
  localStorage.setItem('crm_agents',JSON.stringify(_settingsAgents));
  localStorage.setItem('crm_lead_sources',JSON.stringify(_settingsLeadSources));
  localStorage.setItem('crm_renewals',JSON.stringify(_settingsRenewals));
  applySettingsToDropdowns();
}
loadSettings();
function applySettingsToDropdowns(){
  function repopSel(id,items,blank){var sel=document.getElementById(id);if(!sel||sel.tagName!=='SELECT')return;var cur=sel.value;sel.innerHTML=blank||'';items.forEach(function(v){var o=document.createElement('option');o.value=v;o.textContent=v;sel.appendChild(o);});sel.value=cur;}
  function repopDatalist(id,items){var dl=document.getElementById(id);if(!dl)return;dl.innerHTML='';items.forEach(function(v){var o=document.createElement('option');o.value=v;dl.appendChild(o);});}
  repopSel('f_agent',_settingsAgents,'');
  repopDatalist('leadSourceOptions',_settingsLeadSources);
  refreshReferrerDatalist();
  repopSel('as_leadSource',_settingsLeadSources,'<option value="">All</option>');
  repopSel('f_renewed',_settingsRenewals,'<option value=""></option>');
  repopSel('as_renewed',_settingsRenewals,'<option value="">All</option>');
  repopSel('filterAgent',_settingsAgents,'<option value="">All Agents</option>');
  repopSel('filterLeadSource',_settingsLeadSources,'<option value="">All Lead Sources</option>');
  repopSel('filterRenewed',_settingsRenewals,'<option value="">All Renewal Status</option>');
}
function renderSettingsList(arr,containerId,removeFunc){
  var el=document.getElementById(containerId);if(!el)return;
  el.innerHTML='';
  if(!arr.length){el.innerHTML='<p style="font-size:11px;color:#999;margin-bottom:6px;">None added yet.</p>';return;}
  arr.forEach(function(item,i){
    var div=document.createElement('div');
    div.style.cssText='display:flex;align-items:center;gap:8px;padding:5px 0;border-bottom:1px solid #f0f0f0;';
    div.innerHTML='<span style="flex:1;font-size:12px;">'+item+'</span>'+
      '<button class="btn btn-red" style="padding:2px 8px;font-size:10px;" onclick="'+removeFunc+'('+i+')">Remove</button>';
    el.appendChild(div);
  });
}
function renderSettings(){
  renderSettingsList(_settingsAgents,'agentSettingsList','removeAgentSetting');
  renderSettingsList(_settingsLeadSources,'leadSourceSettingsList','removeLeadSourceSetting');
  renderSettingsList(_customMeds,'customMedSettingsList','removeCustomMedSetting');
  renderSettingsList(_settingsRenewals,'renewalSettingsList','removeRenewalSetting');
  renderSettingsList(_settingsPlanTypes,'planTypeSettingsList','removePlanTypeSetting');
  renderSettingsList(_settingsProjectCodes,'projectCodeSettingsList','removeProjectCodeSetting');
  var nameEl=document.getElementById('settingsCrmName');
  if(nameEl)nameEl.value=localStorage.getItem('crm_display_name')||'Liberty Bell Health';
}
function addAgentSetting(){var v=document.getElementById('newAgentInput').value.trim();if(!v)return;if(_settingsAgents.indexOf(v)!==-1){toast('Agent already exists.','info');return;}_settingsAgents.push(v);document.getElementById('newAgentInput').value='';saveSettings();renderSettings();populateDefaultAgentSelect();}
function removeAgentSetting(i){showConfirm('Remove this agent?',function(){_settingsAgents.splice(i,1);saveSettings();renderSettings();populateDefaultAgentSelect();},{title:'Remove',okText:'Remove'});}
function addLeadSourceSetting(){var v=document.getElementById('newLeadSourceInput').value.trim();if(!v)return;if(_settingsLeadSources.indexOf(v)!==-1){toast('Lead source already exists.','info');return;}_settingsLeadSources.push(v);document.getElementById('newLeadSourceInput').value='';saveSettings();renderSettings();}
function removeLeadSourceSetting(i){showConfirm('Remove this lead source?',function(){_settingsLeadSources.splice(i,1);saveSettings();renderSettings();},{title:'Remove',okText:'Remove'});}
function addCustomMedSetting(){var v=document.getElementById('newCustomMedInput').value.trim();if(!v)return;saveCustomMed(v);document.getElementById('newCustomMedInput').value='';renderSettings();}
function removeCustomMedSetting(i){showConfirm('Remove this medication?',function(){_customMeds.splice(i,1);localStorage.setItem('crm_custom_meds',JSON.stringify(_customMeds));renderSettings();},{title:'Remove',okText:'Remove'});}
function addRenewalSetting(){var v=document.getElementById('newRenewalInput').value.trim();if(!v)return;if(_settingsRenewals.indexOf(v)!==-1){toast('Already exists.','info');return;}_settingsRenewals.push(v);document.getElementById('newRenewalInput').value='';saveSettings();renderSettings();}
function removeRenewalSetting(i){showConfirm('Remove this renewal option?',function(){_settingsRenewals.splice(i,1);saveSettings();renderSettings();},{title:'Remove',okText:'Remove'});}
function saveCrmName(){var v=document.getElementById('settingsCrmName').value.trim();if(!v)return;localStorage.setItem('crm_display_name',v);document.querySelector('.sidebar .logo').childNodes[0].textContent=v;toast('CRM name updated!','success');}

try{initMSAL();loadCarriers();loadSettingsExtras();applySettingsToDropdowns();}catch(e){console.log('MSAL error:',e);showAuthScreen();}
