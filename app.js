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

// Health CRM is Paul + Tommy only. Rob has Home Care access but must NOT reach
// Health PHI — the backend enforces this too via checkApiKey(req,'health').
var ALLOWED_USERS = [
  'tommy@mybellcare.com',
  'paul@mybellcare.com'
];

/* HTML-escape for interpolating client data into markup. Covers element text AND
   quoted attribute values. Client data reaches the DOM from paste-import and CSV,
   so it is never safe to concatenate raw. Use textContent where practical instead. */
var _fullRecordFailed=false; // true when GET /health-clients/{id} failed — blocks save
// Optimistic-concurrency token for the record currently open in the form. Read from
// GET /health-clients/{id} and sent back on save, so the server can refuse a write that
// would silently overwrite someone else's newer edit (409). Null = unconditional write,
// which is what a brand-new record and any pre-upgrade caller does.
var _rowVersion=null;
function escHtml(v){
  return String(v==null?'':v)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}
/* For a value going inside a single-quoted JS string inside an HTML attribute
   (e.g. onclick="f('...')"). Backslash MUST be escaped first or the quote escape
   can be neutralised. Prefer data-* attributes + listeners over this. */
function escJsAttr(v){
  return escHtml(String(v==null?'':v).replace(/\\/g,'\\\\').replace(/'/g,"\\'"));
}

// Telemetry leaves our control, so PHI must never reach it. Every call site here passes ids
// only — but that is discipline, not enforcement, and one careless aiTrack('X',{clientName:…})
// regresses it silently. Drop any property whose KEY looks like PHI, whatever its value.
// `user` is deliberately exempt: it is workforce identity (also set via
// setAuthenticatedUserContext) and is what makes an event attributable.
var _AI_PHI_KEY = /name|client|carrier|plan|email|phone|address|street|city|zip|county|dob|ssn|medicare|medicaid|routing|account|card|premium|income|referrer|employer|notes|file/i;
function _aiScrub(props) {
  var out = {};
  if (!props) return out;
  try {
    for (var k in props) {
      if (!Object.prototype.hasOwnProperty.call(props, k)) continue;
      if (!_AI_PHI_KEY.test(k)) out[k] = props[k];
    }
  } catch (e) { return {}; }
  return out;
}
function aiTrack(name, props) {
  try {
    var ai = window.appInsights;
    if (ai && ai.trackEvent) {
      var acc = msalInstance && msalInstance.getAllAccounts && msalInstance.getAllAccounts();
      var user = (acc && acc.length) ? acc[0].username : 'unknown';
      ai.trackEvent({ name: name }, Object.assign({ user: user, site: 'crm' }, _aiScrub(props)));
    }
  } catch (e) { /* silent */ }
}

/* HIPAA-oriented session inactivity: warn at 43 min, sign out at 45 min.

   Only GENUINE interaction counts. mousemove and scroll used to reset this, which
   meant any cursor drift — or a trackpad nudge, or an animated element scrolling —
   kept an unattended session alive indefinitely, defeating the point. click /
   keydown / touchstart require a person.

   The timers are armed from refreshApiToken().then(...), NOT at sign-in: this
   function no-ops while _apiToken is null, and at sign-in time the token has not
   resolved yet, so arming it there silently did nothing. */
var _sessionTimer=null, _sessionWarnTimer=null;
var SESSION_TIMEOUT_MS = 45 * 60 * 1000;
var SESSION_WARN_MS    = 43 * 60 * 1000;
function resetSessionTimer(){
  clearTimeout(_sessionTimer);clearTimeout(_sessionWarnTimer);
  if(!_apiToken)return; // only enforce once signed in
  _sessionWarnTimer=setTimeout(function(){
    try{toast('Session expiring in 2 minutes. Click anywhere to stay signed in.','info');}catch(e){}
  },SESSION_WARN_MS);
  _sessionTimer=setTimeout(function(){
    aiTrack('SessionTimeout',{reason:'inactivity'});
    try{toast('Signed out due to inactivity','error');}catch(e){}
    signOut();
  },SESSION_TIMEOUT_MS);
}
['click','keydown','touchstart'].forEach(function(ev){
  document.addEventListener(ev,resetSessionTimer,{passive:true,capture:true});
});

// Wipe every crm_* / crmCarriers / lch_* key EXCEPT an explicit whitelist of non-PHI
// *preferences*, so a key added later is covered automatically instead of silently surviving
// sign-out. The previous fixed list was already wrong: it named 'crm_carriers' and
// 'crm_settings', NEITHER of which is ever written, while the real key `crmCarriers` (written
// at saveCarriers) survived — and carrier names are operator-editable free text. It also
// missed crm_default_agent and the crm_carriers_seed_* flags. crm_report_columns is a list of
// COLUMN NAMES with no client data in it, and it has no server copy — it belongs with the other
// per-device preferences, not in the wipe. Everything else here is mirrored to AppSettings and
// comes back from the server at the next sign-in.
// Home Care hit the same bug and replaced the same pattern; see its clearPHIFromStorage.
// crm_todos is a TEMPORARY exception and the only one that is not a preference. It is the legacy
// task list, and until migrateLegacyTodos has handed it to the server it is the ONLY copy — wiping
// it here would silently destroy the agent's tasks on the first sign-out or tab close after this
// change. Nothing writes the key any more, and the migration deletes it the moment the server
// accepts every row, so it disappears on its own at the next successful sign-in.
var _KEEP_ON_SIGNOUT = /(_col_widths|_page_size|_collapsed)$|^crm_display_name$|^crm_report_columns$|^crm_todos$/;
function clearCRMStorage() {
  try {
    Object.keys(localStorage)
      .filter(function(k){
        return (k.indexOf('crm_') === 0 || k.indexOf('crmCarriers') === 0 || k.indexOf('lch_') === 0)
               && !_KEEP_ON_SIGNOUT.test(k);
      })
      .forEach(function(k){ try{ localStorage.removeItem(k); }catch(e){} });
  } catch (e) { /* storage unavailable — nothing to wipe */ }
  // PHI also sits in MEMORY. `clients` holds the whole roster (names, DOB, addresses, last-4),
  // and the open form holds DECRYPTED ssn / card / routing / account. Sign-out today navigates
  // away via logoutRedirect, which takes both with it — but that is incidental, not a control,
  // and it does not hold if the redirect is slow or cancelled.
  try { clients = []; carriers = []; } catch (e) {}
  try {
    document.querySelectorAll('#viewForm input, #viewForm textarea').forEach(function(el){
      if (el.type !== 'checkbox' && el.type !== 'radio' && el.type !== 'button') el.value = '';
    });
  } catch (e) {}
}

var msalInstance=null,_apiTokenTimer=null,clients=[],editingId=null,csvHeaders=[],csvData=[],currentReportData=[],carriers=[];

// ── MSAL authentication ────────────────────────────────────────
function initMSAL(){
  var config={auth:{clientId:SP_CLIENT_ID,authority:'https://login.microsoftonline.com/'+SP_TENANT_ID,redirectUri:REDIRECT_URI},cache:{cacheLocation:'localStorage',storeAuthStateInCookie:true}};
  msalInstance=new msal.PublicClientApplication(config);
  msalInstance.initialize().then(function(){
    msalInstance.handleRedirectPromise().then(function(resp){
      if(resp&&resp.account){onSignedIn(resp.account);return;}
      var accounts=msalInstance.getAllAccounts();
      if(accounts.length>0){
        // Already signed in on this device — attempt silent Microsoft SSO
        setAuthMsg('Signing you in…');
        onSignedIn(accounts[0]);
      } else {
        // No cached account — reveal the Sign In button
        showAuthScreen();
      }
    }).catch(function(){showAuthScreen();});
  });
}
/* Show the login wall with the Sign In button. While MSAL is checking silent
   auth we leave the button hidden and show 'Verifying authentication…'. */
function setAuthMsg(text){var m=document.getElementById('authScreenMsg');if(m)m.textContent=text||'';}
function showAuthScreen(){
  var scr=document.getElementById('authScreen'),btn=document.getElementById('authScreenBtn');
  if(scr)scr.style.display='flex';
  document.getElementById('mainApp').style.display='none';
  setAuthMsg('Sign in with your Microsoft 365 account to continue');
  if(btn)btn.style.display='inline-block';
}
function onSignedIn(account){
  var email=(account&&(account.username||account.name||'')).toLowerCase();
  if(!ALLOWED_USERS.map(function(u){return u.toLowerCase();}).includes(email)){
    // Show the denial inline on the loginWall so the user sees why before auto sign-out.
    document.getElementById('authScreen').style.display='flex';
    document.getElementById('mainApp').style.display='none';
    setAuthMsg('');
    var btn=document.getElementById('authScreenBtn');if(btn)btn.style.display='none';
    var scr=document.getElementById('authScreen');
    var existing=document.getElementById('authDeniedNote');if(existing)existing.remove();
    var note=document.createElement('div');
    note.id='authDeniedNote';
    note.style.cssText='max-width:420px;text-align:center;padding:16px 20px;background:rgba(160,32,32,0.15);border:1px solid rgba(160,32,32,0.4);border-radius:8px;color:#fff;font-size:13px;line-height:1.6;';
    note.innerHTML='<div style="color:#ff8080;font-weight:700;font-size:14px;margin-bottom:6px;">Access denied</div>'+
      'The account <strong>'+escHtml(email)+'</strong> is not authorized for this application. Signing you out…';
    scr.appendChild(note);
    try{toast('Access denied for '+email,'error');}catch(e){}
    // §164.308(a)(1)(ii)(D) — an unauthorized access attempt is a reviewable event. It existed
    // only as a toast that vanished in 3 seconds.
    try{logActivity('auth','ACCESS DENIED for '+email);}catch(e){}
    setTimeout(function(){msalInstance.logoutRedirect({redirectUri:REDIRECT_URI});},4000);
    return;
  }
  document.getElementById('authScreen').style.display='none';
  document.getElementById('mainApp').style.display='flex';
  document.getElementById('userEmail').textContent=email;
  aiTrack('UserSignIn',{email:email});
  // Tag every subsequent App Insights event with this user so we can correlate
  // errors / usage to a specific person when debugging.
  try{window._aiUser=email;if(window.appInsights&&window.appInsights.setAuthenticatedUserContext)window.appInsights.setAuthenticatedUserContext(email);}catch(e){}
  refreshApiToken().then(function(){
    try{loadSettingsAPI();}catch(e){} resetSessionTimer(); loadClients(); routeFromHash();
    // Migrate first, THEN load: anything still on this device has to reach the server before the
    // server's list replaces the in-memory one, or it is lost.
    try{migrateLegacyTodos().then(loadTasksAPI);}catch(e){} });
}
/* Hash-based deep links so URLs like /#/client/<id> open that client directly.
   Enables bookmarking and sharing a link to a specific record. */
function routeFromHash(){
  var h=(window.location.hash||'').replace(/^#/,'');
  var m=h.match(/^\/client\/(.+)$/);
  if(!m)return;
  var id=decodeURIComponent(m[1]);
  // Wait a tick for clients to be in memory if we just loaded them
  var tryOpen=function(attempts){
    var c=(clients||[]).find(function(x){return String(x._id)===String(id);});
    if(c){editClient(id);}
    else if(attempts>0){setTimeout(function(){tryOpen(attempts-1);},250);}
  };
  tryOpen(8);
}
window.addEventListener('hashchange',routeFromHash);
var API_SCOPE='api://'+API_APP_ID+'/user_impersonation';
// Only Graph scopes in loginRedirect — API_SCOPE from different resource causes 400 on token endpoint.
// refreshApiToken() acquires API token silently after login (admin consent already granted).
function signIn(){msalInstance.loginRedirect({scopes:['openid','profile'],redirectUri:REDIRECT_URI});}
function signOut(){
  aiTrack('UserSignOut',{});
  // keepalive on the audit POST lets it survive the logout redirect below.
  try{logActivity('auth','Signed out');}catch(e){}
  clearCRMStorage();
  clearTimeout(_sessionTimer);clearTimeout(_sessionWarnTimer);
  clearTimeout(_apiTokenTimer);_apiTokenTimer=null;
  _apiToken=null;
  msalInstance.logoutRedirect({redirectUri:REDIRECT_URI});
}

// ── API HELPERS ────────────────────────────────────────────────
/* ── AUDIT (HIPAA §164.312(b)) ────────────────────────────────────────────────
   This app recorded NOTHING. Opening a client pulls the DECRYPTED ssn, card, routing and
   account number into the page, and none of that left a trace of who looked or when.
   The backend endpoints already existed and were reachable; only the caller was missing.

   scope:'health' is REQUIRED — without it these rows land in Liberty Home Care's AuditLog.
   The two companies are separate legal entities and their audit trails are separate tables.

   `who` / `actor` are set SERVER-SIDE from the authenticated identity and ignored from the
   body, so the trail cannot be forged by a caller. */
function currentUserEmail(){
  try{
    var acc=msalInstance&&msalInstance.getAllAccounts&&msalInstance.getAllAccounts();
    return (acc&&acc.length?acc[0].username:null)||'Unknown';
  }catch(e){return 'Unknown';}
}
// MUST match what documents.js writes for clientType=health:
//   LTRIM(RTRIM(ISNULL(first_name,'') + ' ' + ISNULL(last_name,'')))
// client_name is the ONLY lookup key the audit search uses, so a mismatch splits one client's
// history into two views that can never see each other.
function _auditName(c){
  if(!c)return '';
  return String((c.f_firstName||'')+' '+(c.f_lastName||'')).trim();
}
/* The audit trail lives SERVER-SIDE (POST /audit, read back by loadClientAudit). There used to be
   a parallel localStorage mirror here holding the patient name on 200 entries deep. Nothing ever
   read it — loadClientAudit has always gone to the API — so it was pure standing PHI on the disk of
   every browser that had ever opened a record, in plain violation of the no-PHI-in-localStorage
   rule that crm_recent and crm_todos were both already fixed for. Removing the writes does nothing
   for the browsers that already have months of it, so purge on startup too. */
function purgeLegacyAuditLog(){try{localStorage.removeItem('crm_audit');}catch(e){}}
function _postAuditRecord(body){
  // A failure here must be VISIBLE. A silently-dropped audit row is worse than a failed save:
  // the action still happened, and nothing records it.
  if(!_apiToken){
    toast('Not signed in — this action was NOT written to the audit log.','error',15000);
    return;
  }
  fetch(API_BASE+'/audit',{
    method:'POST',headers:apiHeaders(),
    body:JSON.stringify(Object.assign({scope:'health'},body)),
    keepalive:true   // survives a sign-out redirect or tab close
  }).then(function(r){
    if(!r.ok)toast('Audit log write failed ('+r.status+') — this action was not recorded.','error',15000);
  }).catch(function(){
    toast('Audit log write failed — this action was not recorded.','error',15000);
  });
}
function addAuditEntry(clientName,action){
  _postAuditRecord({event_type:'audit',client_name:clientName,action:action});
}
// Global (non-client) events — exports, sign-in/out, denied access. client_name:'' keeps them
// out of any one client's tab while still being in the trail.
function logActivity(type,text){ _postAuditRecord({event_type:type,client_name:'',action:text}); }

/* ── SETTINGS SYNC ────────────────────────────────────────────────────────────
   These were localStorage-only, so two agents on two machines saw DIFFERENT agent lists,
   lead sources, plan types and carriers — and everything was lost on sign-out, since the
   sign-out wipe (correctly) clears crm_*. Now mirrored to AppSettings under scope 'health'.

   scope:'health' is required. AppSettings is keyed on (scope, setting_key) precisely so the
   two companies can each have an 'agents' key without one silently overwriting the other.

   `raw` marks the two values stored as bare strings rather than JSON. */
var _SYNCED_SETTINGS=[
  ['crm_agents','agents',false],
  ['crm_lead_sources','lead_sources',false],
  ['crm_renewals','renewals',false],
  ['crm_plan_types','plan_types',false],
  ['crm_project_codes','project_codes',false],
  ['crm_custom_meds','custom_meds',false],
  ['crmCarriers','carriers',false],
  ['crm_default_agent','default_agent',true],
  ['crm_display_name','display_name',true]
];
var _settingsPushTimer=null,_settingsPushInFlight=0;
// Write locally FIRST so the UI never waits on the network, then push. Debounced because the
// settings screens fire several writes in a row (add an agent -> re-render -> save again).
function _syncedSetItem(k,v){
  try{localStorage.setItem(k,v);}catch(e){}
  clearTimeout(_settingsPushTimer);
  _settingsPushTimer=setTimeout(_pushSettings,800);
}
/* True once the server's copy has been read at least once this session — including a read that
   came back empty, which is a legitimate first run. Until then this device's settings are NOT
   authoritative and must never be pushed.

   This matters because the sign-out / tab-close wipe clears every crm_* key, so a session starts
   with nothing and depends on loadSettingsAPI to refill it. If that fetch fails, loadSettings()
   has already fallen back to hardcoded defaults and carriers to an empty array — and the old
   comment on the catch below, "the local values stand, and a later write pushes them", then meant
   pushing DEFAULTS over the real list. One added agent would have replaced the server's agents
   with the two built-ins plus that one; one added carrier would have replaced every carrier. */
var _settingsLoaded=false;
function _pushSettings(){
  if(!_apiToken)return;
  if(!_settingsLoaded){
    toast('Settings are not syncing — the saved list could not be loaded. This change stays on this device only. Reload to try again.','error',15000);
    return;
  }
  var body={scope:'health'};
  _SYNCED_SETTINGS.forEach(function(e){
    var raw=localStorage.getItem(e[0]);
    if(raw===null)return;
    if(e[2]){body[e[1]]=raw;return;}
    try{body[e[1]]=JSON.parse(raw);}catch(err){body[e[1]]=raw;}
  });
  _settingsPushInFlight++;
  fetch(API_BASE+'/settings',{method:'POST',headers:apiHeaders(),body:JSON.stringify(body),keepalive:true})
    .then(_apiOk)
    .catch(function(e){
      // Loud, because the local value LOOKS saved. Silence here is how two machines drift apart.
      toast('Settings saved on this device but did NOT sync: '+((e&&e.message)||e),'error',10000);
    })
    .then(function(){_settingsPushInFlight=Math.max(0,_settingsPushInFlight-1);});
}
function loadSettingsAPI(){
  if(!_apiToken)return Promise.resolve();
  return fetch(API_BASE+'/settings?scope=health',{headers:apiHeaders()})
    .then(_apiOk).then(function(r){return r.json();})
    .then(function(remote){
      // The read succeeded, so this device is reconciled with the server and may push again.
      // Set before the early returns below: "the server has nothing yet" is a successful read on
      // a first run, and pushing local values up from there is exactly right.
      _settingsLoaded=true;
      // A local save is in flight — applying the server's older copy now would revert what the
      // user just typed. Home Care hit exactly this.
      if(_settingsPushInFlight)return;
      if(!remote||typeof remote!=='object')return;
      var changed=false;
      _SYNCED_SETTINGS.forEach(function(e){
        if(!Object.prototype.hasOwnProperty.call(remote,e[1]))return;
        var v=remote[e[1]];
        try{localStorage.setItem(e[0],e[2]?String(v):JSON.stringify(v));changed=true;}catch(err){}
      });
      if(!changed)return;
      // Re-read into memory and repaint, or the running page keeps the pre-sync values.
      try{loadSettings();}catch(err){}
      try{loadCarriers();}catch(err){}
      try{loadCustomMeds();}catch(err){}
      try{applySettingsToDropdowns();}catch(err){}
      try{if(document.getElementById('viewSettings').style.display!=='none')renderSettings();}catch(err){}
    })
    .catch(function(e){
      // Deliberately leaves _settingsLoaded false, so nothing this device holds can overwrite the
      // server copy it never managed to read. Loud, because the lists on screen are now defaults
      // rather than this agency's real ones, and that is not otherwise obvious.
      toast('Could not load your saved settings ('+((e&&e.message)||'network error')+'). Agent, carrier and plan lists may be incomplete, and changes will not sync until you reload.','error',15000);
    });
}

function apiHeaders(){var h={'Content-Type':'application/json'};if(_apiToken)h['Authorization']='Bearer '+_apiToken;return h;}
function authUploadHeaders(){return _apiToken?{'Authorization':'Bearer '+_apiToken}:{};}
async function refreshApiToken(){
  if(!msalInstance)return;
  var accounts=msalInstance.getAllAccounts();if(!accounts.length)return;
  try{
    var res=await msalInstance.acquireTokenSilent({scopes:[API_SCOPE],account:accounts[0]});
    _apiToken=res.accessToken;
    var ttl=res.expiresOn?(res.expiresOn.getTime()-Date.now()-600000):3000000;
    // Keep the handle: an uncancelled refresh timer silently re-acquires a token after
    // sign-out, and resetSessionTimer (gated on _apiToken) then re-arms — so the HIPAA idle
    // sign-out would not actually hold if the logout redirect were slow or cancelled.
    clearTimeout(_apiTokenTimer);
    _apiTokenTimer=setTimeout(refreshApiToken,Math.max(ttl,60000));
  }catch(e){
    console.warn('API token silent refresh failed, opening consent popup:',e);
    try{
      var r2=await msalInstance.acquireTokenPopup({scopes:[API_SCOPE]});
      _apiToken=r2.accessToken;
      var ttl2=r2.expiresOn?(r2.expiresOn.getTime()-Date.now()-600000):3000000;
      clearTimeout(_apiTokenTimer);
      _apiTokenTimer=setTimeout(refreshApiToken,Math.max(ttl2,60000));
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
    phone:d.f_phone, phone_ext:d.f_phoneExt, alt_phone:d.f_altPhone, alt_phone_ext:d.f_altPhoneExt, email:d.f_email, email2:d.f_email2,
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
    // undefined, not 0, when the form did not supply it. The backend filters an update to the
    // columns actually present in the body (`body[f] !== undefined`) precisely so a save cannot
    // touch what it was not given, and JSON.stringify drops undefined keys — but a default of 0
    // is a real value, so it sailed through and cleared the flag on every save. There is no
    // #f_waiveDental input any more (the flag moved into the Dental ancillary row), so the form
    // NEVER supplies it: this column could only ever be 0. false still sends 0, an explicit clear.
    waive_dental:d.f_waiveDental===undefined?undefined:(d.f_waiveDental?1:0), total_monthly:d.f_totalMonthly,
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
    card_type:d.f_cardType, card_number:d.f_cardNumber, card_exp:d.f_cardExp,
    agent:d.f_agent, submitted_by:d.f_submittedBy, application_date:d.f_date,
    lead_source:d.f_leadSource, lead_date:d.f_leadDate, renewed:d.f_renewed,
    mothers_maiden:d.f_mothersMaiden, notes:d.f_notes,
    members_json:JSON.stringify(d.members||[]),
    doctors_json:JSON.stringify(d.doctors||[]),
    medications_json:JSON.stringify(d.meds||[]),
    ancil_plans_json:JSON.stringify(d.ancilPlans||[]),
    // Same trap, worse consequence: `parseInt(undefined)||null` is null, a real value, so every
    // save unlinked the Home Care record. There is no #f_homecareClientId input, so the form can
    // never supply it — which is also why the client list's "homeCare" sort key can only ever
    // read 0. An empty string still sends null, so an intentional unlink still works.
    homecare_client_id:d.f_homecareClientId===undefined?undefined:(parseInt(d.f_homecareClientId)||null),
  };
}
function dbRowToClient(row){
  var c={
    _id:row.id,
    f_firstName:row.first_name, f_mi:row.middle_initial, f_lastName:row.last_name,
    f_dob:row.dob, f_age:row.age, f_gender:row.gender, f_ssn:row.ssn,
    // list rows carry only masked last-4s; the full values arrive via GET /{id}
    f_ssnLast4:row.ssn_last4, f_cardLast4:row.card_last4,
    f_relation:row.relation, f_marital:row.marital_status,
    f_tobacco:row.tobacco, f_height:row.height, f_weight:row.weight, f_insured:row.insured,
    f_phone:row.phone, f_phoneExt:row.phone_ext, f_altPhone:row.alt_phone, f_altPhoneExt:row.alt_phone_ext, f_email:row.email, f_email2:row.email2,
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
  .then(_apiOk)
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
    renderSidebarRecent(); // names only resolvable once clients are in memory
  }).catch(function(e){
    // Console-only was not enough. This runs after every save and delete, so a failure here left
    // the agent looking at a roster that silently no longer matched the server — and on a 500 the
    // error body is valid JSON, so `data.map` threw and even the render never ran. The roster
    // already in memory is deliberately left alone: stale-but-labelled beats blank.
    console.error('Load error:',e);
    toast('Could not refresh the client list ('+(e&&e.message||'network error')+'). What you see may be out of date.','error',10000);
  });
}
/* Reject on a non-2xx instead of sailing through it.
   The backend's 500 path returns {error:'...'} — VALID JSON — so a bare .then(r=>r.json())
   RESOLVED on failure. saveClient then took its success branch: cleared the dirty flag and
   called loadClients(), which repopulated the form from the server and DISCARDED the user's
   edit, behind a green "Client saved!" toast. Surface the server's own message when it sends
   one. This is the shape editClient already used; it was the only status check in the file. */
function _apiOk(r){
  if(r.ok)return r;
  return r.json().catch(function(){return null;}).then(function(body){
    var err=new Error((body&&body.error)||('HTTP '+r.status));
    // Carry the status so callers can distinguish a lost-update conflict (409) from a
    // generic failure — they need different handling, not just a different message.
    err.status=r.status;
    if(body&&body.current_version)err.currentVersion=body.current_version;
    throw err;
  });
}
function saveClientAPI(data,id){
  var body=clientToDbRow(data);
  if(id)body.id=id;
  // Only on an update, and only when we actually read a version — otherwise omit it and the
  // server writes unconditionally, which is the correct behaviour for a new record.
  if(id&&_rowVersion)body.expected_version=_rowVersion;
  return fetch(API_BASE+'/health-clients',{method:'POST',headers:apiHeaders(),body:JSON.stringify(body)})
    .then(_apiOk).then(function(r){return r.json();});
}
function deleteClientAPI(id){
  return fetch(API_BASE+'/health-clients/'+id,{method:'DELETE',headers:apiHeaders()})
    .then(_apiOk);
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
  // Some views (new, recent) no longer have a sidebar button — guard against null.
  var nid=navMap[v];var nel=nid&&document.getElementById(nid);if(nel)nel.classList.add('active');
  if(vid)document.querySelector('.main').scrollTop=0;
  // Clear the deep-link hash when leaving the edit form so URL matches the view
  if(v!=='form_edit'&&window.location.hash){try{history.replaceState(null,'','#');}catch(e){window.location.hash='';}}
  if(v==='clients')maybeRevalidateClients();
  if(v==='carriers')renderCarriers();
  if(v==='form_edit'||v==='new')setTimeout(wireCopyableFields,50);
  if(v==='todo')renderTodos();
  if(v==='settings'){renderSettings();populateDefaultAgentSelect();}
  if(v==='advSearch')populateAdvSearchCarriers();
  if(v==='recent')renderRecentRecords();
}

function openDiscoveryPasteModal(){var m=document.getElementById('discoveryPasteModal');if(m){document.getElementById('discoveryPasteInput').value='';m.style.display='flex';setTimeout(function(){document.getElementById('discoveryPasteInput').focus();},50);}}
function closeDiscoveryPasteModal(){var m=document.getElementById('discoveryPasteModal');if(m)m.style.display='none';}

/* Parse a filled discovery template into a client + members structure.
   Recognizes labels ending in "-" or ":". Blank values are skipped. Spouse
   and Child blocks only create members when they have any data. Ignores
   fields the user chose to skip (SEP, HA/S/C, D/V, Preexisting Conditions,
   Hospital, Rx (handled elsewhere), etc). */
/* Every label parseDiscoveryText recognises. Used to distinguish a real "Label- value" line from
   a LIST ITEM that merely contains a dash — "Metformin 500mg - twice daily", "Dr. Smith -
   Cardiology". Both match the label pattern, and inside a collector both are data, not labels.
   Arrays + indexOf rather than an object, so a value like "constructor" can't test truthy. */
var _DISCOVERY_LABELS=['rx','medications','primary dr','primary doctor','specialist','dr visits',
  'lead source','phone number','phone','inquiry date','lead date','time','name','email','state',
  'zip','county','household income','date of birth','gender','m/f'];
var _DISCOVERY_NORM_LABELS=['dob','age','mf','tobacco','height','weight'];
function _isDiscoveryLabel(label,norm){
  return _DISCOVERY_LABELS.indexOf(label)!==-1||_DISCOVERY_NORM_LABELS.indexOf(norm)!==-1;
}
function parseDiscoveryText(text){
  var lines=(text||'').split(/\r?\n/);
  var section='primary'; // primary | spouse | child1 | child2 | child3 | child4
  var data={};        // top-level primary values (f_firstName etc.)
  var members={};      // section -> {firstName,mi,lastName,relation,...}
  var meds=[]; // parsed from Rx: (single-line or multi-line)
  var doctors=[]; // parsed from Primary Dr / Specialist / Dr Visits lines
  var notesExtras=[];

  function setPrimary(field,val){data['f_'+field]=val;}
  function setMember(sec,rel,field,val){
    if(!members[sec])members[sec]={relation:rel};
    members[sec][field]=val;
  }
  function addMed(str){
    var s=(str||'').trim();if(!s)return;
    _splitMedLine(s).forEach(function(part){
      var p=parseMedLine(part);
      if(p&&(p.name||p.mg))meds.push(p);
    });
  }
  function addDoc(name,specialty){
    var n=(name||'').trim();if(!n)return;
    // parseDoctorLine already splits "Dr. Smith - Cardiology"; the section ('Primary'/'Specialist')
    // is the fallback for a line that names no specialty of its own.
    var p=parseDoctorLine(n)||{name:n,specialty:''};
    doctors.push({name:p.name||n,specialty:p.specialty||specialty||''});
  }

  // Multi-line collector: after a label like "Rx-" (empty value), subsequent
  // unlabeled non-blank lines belong to that field until the next label.
  var collecting=null; // 'rx' | 'primary_dr' | 'specialist' | null

  for(var i=0;i<lines.length;i++){
    var raw=lines[i]; var line=raw.trim();
    if(!line){collecting=null;continue;} // blank line ends multi-line mode
    // Section headers first (bare labels, not "field-value" lines)
    if(/^spouse\s*[-:]?\s*$/i.test(line)){section='spouse';collecting=null;continue;}
    var childMatch=line.match(/^child\s*(\d)\s*[-:]?\s*$/i);
    if(childMatch){section='child'+childMatch[1];collecting=null;continue;}
    // Parse "Label- value" or "Label: value"
    var m=line.match(/^([^\-:]+?)\s*[-:]\s*(.*)$/);
    if(!m){
      // Unlabeled continuation of a multi-line field
      if(collecting==='rx')addMed(line);
      else if(collecting==='primary_dr')addDoc(line,'Primary');
      else if(collecting==='specialist')addDoc(line,'Specialist');
      continue;
    }
    var label=m[1].trim().toLowerCase(),val=m[2].trim();

    // Multi-line trigger labels — accept either same-line values OR set collecting mode for following lines
    if(label==='rx'||label==='medications'){
      collecting='rx';
      if(val){_splitMedLine(val).forEach(function(ln){addMed(ln);});}
      continue;
    }
    if(label==='primary dr'||label==='primary doctor'){
      collecting='primary_dr';
      if(val)addDoc(val,'Primary');
      continue;
    }
    if(label==='specialist'){
      collecting='specialist';
      if(val)addDoc(val,'Specialist');
      continue;
    }
    if(label==='dr visits'){
      // Just informational; drop into notes if there's a value
      collecting=null;
      if(val)notesExtras.push('Dr Visits: '+val);
      continue;
    }
    // Inside a collector, a line whose "label" is not one we recognise is a LIST ITEM that happens
    // to contain a dash or colon, not a label. It used to be parsed as a label, dropped for
    // matching nothing — and, because the next statement cleared the collector, every remaining
    // line was dropped too. One "Metformin 500mg - twice daily" silently emptied the whole list.
    if(collecting&&!_isDiscoveryLabel(label,label.replace(/[^\w]/g,''))){
      if(collecting==='rx')addMed(line);
      else if(collecting==='primary_dr')addDoc(line,'Primary');
      else if(collecting==='specialist')addDoc(line,'Specialist');
      continue;
    }
    // Any other label line resets the multi-line collector
    collecting=null;
    if(!val)continue; // skip blanks after non-multi-line labels

    // Primary-only single-value fields
    if(section==='primary'){
      if(label==='lead source'){setPrimary('leadSource',val);continue;}
      if(label==='phone number'||label==='phone'){setPrimary('phone',formatPhoneStr(val));continue;}
      if(label==='inquiry date'||label==='lead date'){setPrimary('leadDate',parseDateStr(val));continue;}
      if(label==='time'){notesExtras.push('Inquiry time: '+val);continue;}
      if(label==='name'){
        var parts=val.split(/\s+/);
        setPrimary('firstName',parts.shift()||'');
        if(parts.length){var last=parts.pop();setPrimary('lastName',last);if(parts.length)setPrimary('mi',parts.join(' ').charAt(0));}
        continue;
      }
      if(label==='email'){setPrimary('email',val);continue;}
      if(label==='state'){setPrimary('resSt',val.toUpperCase().slice(0,2));continue;}
      if(label==='zip'){setPrimary('resZip',val.replace(/\D/g,'').slice(0,5));continue;}
      if(label==='county'){setPrimary('resCounty',val);continue;}
      if(label==='household income'){setPrimary('primaryIncome',val.replace(/[^\d.]/g,''));continue;}
    }

    // Shared "person" fields — always apply to whichever section we're in
    var normLabel=label.replace(/[^\w]/g,''); // "m/f" → "mf"
    var isPrimaryScope=(section==='primary');
    var mapField=null;
    if(normLabel==='dob'||label==='date of birth')mapField='dob';
    else if(normLabel==='age')mapField='age';
    else if(normLabel==='mf'||label==='gender'||label==='m/f')mapField='gender';
    else if(normLabel==='tobacco')mapField='tobacco';
    else if(normLabel==='height')mapField='height';
    else if(normLabel==='weight')mapField='weight';
    if(!mapField)continue;
    // Normalize DOB to yyyy-mm-dd; normalize M/F to M or F
    if(mapField==='dob')val=parseDateStr(val);
    if(mapField==='gender'){var g=val.toUpperCase().charAt(0);val=(g==='M'||g==='F')?g:'';if(!val)continue;}
    if(mapField==='tobacco'){var t=val.toLowerCase();val=/^y|1|smok/.test(t)?'Yes':(/^n|0/.test(t)?'No':'');if(!val)continue;}

    if(isPrimaryScope){setPrimary(mapField,val);}
    else if(section==='spouse'){setMember('spouse','Spouse',mapField,val);}
    else if(section.indexOf('child')===0){setMember(section,'Child',mapField,val);}
  }

  // Build members array in order, filtering out any that ended up empty
  var memberList=[];
  ['spouse','child1','child2','child3','child4'].forEach(function(sec){
    var mem=members[sec];if(!mem)return;
    var hasData=Object.keys(mem).some(function(k){return k!=='relation'&&(mem[k]||'').toString().trim();});
    if(hasData)memberList.push(mem);
  });

  return {data:data,members:memberList,meds:meds,doctors:doctors,notesExtras:notesExtras};
}
function parseDateStr(s){
  // Accepts YYYY-MM-DD, MM/DD/YYYY, MM/DD/YY, M-D-YY, etc. Returns YYYY-MM-DD or original.
  if(!s)return s;var t=s.trim();
  var iso=t.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if(iso)return iso[1]+'-'+String(iso[2]).padStart(2,'0')+'-'+String(iso[3]).padStart(2,'0');
  var us=t.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})$/);
  if(us){var y=us[3];if(y.length===2)y=(parseInt(y)>30?'19':'20')+y;return y+'-'+String(us[1]).padStart(2,'0')+'-'+String(us[2]).padStart(2,'0');}
  return t;
}
function formatPhoneStr(s){var d=(s||'').replace(/\D/g,'').slice(0,10);if(d.length!==10)return s;return '('+d.slice(0,3)+') '+d.slice(3,6)+'-'+d.slice(6);}

function importDiscoveryPaste(){
  var text=document.getElementById('discoveryPasteInput').value||'';
  if(!text.trim()){closeDiscoveryPasteModal();return;}
  var parsed=parseDiscoveryText(text);
  closeDiscoveryPasteModal();
  // Boot up a new health app (clears form, opens edit view)
  startNewApp('health');
  // Populate — done in a timeout to let the view + starter rows render first
  setTimeout(function(){
    Object.keys(parsed.data).forEach(function(k){
      var el=document.getElementById(k);if(!el)return;
      if(el.type==='checkbox')el.checked=!!parsed.data[k];else el.value=parsed.data[k];
    });
    // Members: reset container (startNewApp doesn't add a starter member row)
    var memC=document.getElementById('membersContainer');if(memC)memC.innerHTML='';
    parsed.members.forEach(function(m){addMemberRow(m);});
    // Doctors / meds: startNewApp added one empty starter row each — wipe first
    var docC=document.getElementById('doctorsContainer');if(docC)docC.innerHTML='';
    parsed.doctors.forEach(function(d){addDoctorRow(d);});
    if(!parsed.doctors.length)addDoctorRow(); // keep an empty starter
    var medC=document.getElementById('medsContainer');if(medC)medC.innerHTML='';
    parsed.meds.forEach(function(m){addMedRow(m);});
    if(!parsed.meds.length)addMedRow();
    // Extras (like inquiry time) → append to notes
    if(parsed.notesExtras.length){
      var n=document.getElementById('f_notes');if(n)n.value=(n.value?n.value+'\n\n':'')+parsed.notesExtras.join('\n');
    }
    // Trigger age recalc + zip lookup so derived fields fill in
    if(parsed.data.f_dob){try{calcAge();}catch(e){}}
    if(parsed.data.f_resZip){try{lookupZip(document.getElementById('f_resZip'),'res');}catch(e){}}
    markFormDirty();
    updateMemberCount();
    toast('Imported '+parsed.members.length+' member'+(parsed.members.length===1?'':'s')+' + '+parsed.meds.length+' med'+(parsed.meds.length===1?'':'s')+' + '+parsed.doctors.length+' doctor'+(parsed.doctors.length===1?'':'s'),'success');
  },80);
}

function startNewApp(type){
  if(type==='life'){toast('Life App coming soon!','info');return;}
  try{clearForm();}catch(e){console.log('clearForm error:',e);}
  editingId=null;
  _rowVersion=null;
  _fullRecordFailed=false;
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
/* Off-canvas sidebar for phones. Separate from toggleSidebar(), which is the DESKTOP
   collapse-to-44px behaviour and is persisted; this one is transient and never stored, because
   a drawer left "open" across loads would cover the app on the next visit. */
function toggleMobileSidebar(force){
  var sb=document.getElementById('sidebar');if(!sb)return;
  var open=(force===undefined)?!sb.classList.contains('mobile-open'):!!force;
  sb.classList.toggle('mobile-open',open);
  var bd=document.getElementById('sbBackdrop');
  if(open&&!bd){
    bd=document.createElement('div');
    bd.className='sb-backdrop';bd.id='sbBackdrop';
    bd.addEventListener('click',function(){toggleMobileSidebar(false);});
    document.body.appendChild(bd);
  }else if(!open&&bd){ bd.remove(); }
  var btn=document.getElementById('sbMobileToggle');
  if(btn)btn.setAttribute('aria-expanded',open?'true':'false');
}
// Tapping a nav item should navigate AND close the drawer — leaving it open over the view the
// user just asked for is the classic mobile-nav mistake.
document.addEventListener('click',function(e){
  var btn=e.target&&e.target.closest?e.target.closest('.nav-btn'):null;
  if(btn)toggleMobileSidebar(false);
},true);
// Escape closes it, and a resize back to desktop must not leave a stranded backdrop.
document.addEventListener('keydown',function(e){
  if(e.key==='Escape'&&document.getElementById('sbBackdrop'))toggleMobileSidebar(false);
});
window.addEventListener('resize',function(){
  if(window.innerWidth>720&&document.getElementById('sbBackdrop'))toggleMobileSidebar(false);
});

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
  // Was `spToken`, which is declared nowhere and assigned nowhere — permanently undefined,
  // so this whole SWR-on-tab-return path was dead and the list never refreshed.
  if(age>30000&&_apiToken){loadClients();}
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
    var badge=c.f_planType?'<span class="badge '+bc+'">'+escHtml(c.f_planType)+'</span>':'';
    var phone=c.f_phone||'';var email=c.f_email||'';
    // Renewed status → dot + plain text (matches Home Care's status pattern)
    var renewedRaw=(c.f_renewed||'').trim();
    var renewedDotClass=renewedRaw==='2026 Renewed'||renewedRaw.indexOf('Renewed')===0?'renewed':(renewedRaw==='Not Renewed'?'notrenewed':'');
    var renewedCell=renewedRaw?'<span class="status-inline"><span class="status-dot '+renewedDotClass+'"></span>'+escHtml(renewedRaw)+'</span>':'';
    tr.innerHTML='<td><input type="checkbox" class="row-cb" data-id="'+escHtml(c._id)+'" onchange="updateBulkBtn()"></td>'+
      '<td><span class="client-name-link link-plain" onclick="editClient(\''+escJsAttr(c._id)+'\')">'+escHtml(((c.f_firstName||'')+' '+(c.f_lastName||'')).trim())+'</span></td>'+
      '<td>'+escHtml(c.f_dob)+'</td>'+
      '<td>'+escHtml(phone)+(phone?'<a href="tel:'+escHtml(phone)+'" class="icon-btn" title="Call">'+SVG_PHONE+'</a><button class="icon-btn" onclick="copyText(\''+escJsAttr(phone)+'\',this)" title="Copy">'+SVG_COPY+'</button>':'')+'</td>'+
      '<td>'+escHtml(email)+(email?'<a href="mailto:'+escHtml(email)+'" class="icon-btn" title="Email">'+SVG_MAIL+'</a><button class="icon-btn" onclick="copyText(\''+escJsAttr(email)+'\',this)" title="Copy">'+SVG_COPY+'</button>':'')+'</td>'+
      '<td>'+badge+'</td>'+
      '<td>'+escHtml(c.f_planName)+'</td>'+
      '<td>'+(c.f_premium?'$'+escHtml(c.f_premium):'')+'</td>'+
      '<td>'+escHtml(c.f_agent)+'</td>'+
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
    if(ssn4){var l4=(c.f_ssnLast4||c.f_ssn||'').replace(/\D/g,'').slice(-4);if(l4!==ssn4)return false;}
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
  'phone','phoneExt','altPhone','altPhoneExt','email','email2',
  'emergencyName','emergencyRelation','emergencyPhone',
  // No 'cvv' here, and no #f_cvv input: PCI-DSS prohibits storing the card verification value
  // after authorization, and the HealthClients.cvv column was dropped on 2026-09-03. Adding
  // either one back would silently start capturing it again — the tests pin this.
  'bankName','accountType','routing','account','accountName','cardType','cardNumber','cardExp',
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
  syncExtVisibility(); // fields are now empty, so this re-hides both Ext boxes
  ['membersContainer','doctorsContainer','medsContainer','ancilContainer','otherIncomeContainer'].forEach(function(id){
    var el=document.getElementById(id);if(el)el.innerHTML='';
  });
  updateOtherIncomeAddBtn();
  var ag=document.getElementById('f_agent');if(ag)ag.value=localStorage.getItem('crm_default_agent')||'Thomas Jaboro';
  var sa=document.getElementById('f_diffMailing');if(sa)sa.checked=false;
  var ms=document.getElementById('mailingAddressSection');if(ms)ms.style.display='none';
  var at=document.getElementById('f_addressType');if(at){at.value='Mailing';updateMailingTitle();}
  var rb=document.getElementById('f_referredBy');if(rb){rb.value='';rb.style.borderColor='';rb.title='';}_referrerPicked=false;
  var st=document.getElementById('f_status');if(st)st.value='Active';
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
  // Status packing: strip any existing [STATUS:X] prefix from notes, then
  // prepend the current status. Kept invisible in the UI (extracted on load).
  var stEl=document.getElementById('f_status');
  var status=stEl?(stEl.value||'Active'):'Active';
  var notes=(data.f_notes||'').replace(/^\[STATUS:[^\]]+\]\n?/i,'');
  data.f_notes='[STATUS:'+status+']\n'+notes;
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
  syncExtVisibility(); // reveal the Ext box only when this client actually has one
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
  // Unpack packed status prefix in notes and drive the Status dropdown.
  // Old clients without a prefix default to Active.
  var stEl=document.getElementById('f_status');
  if(stEl){
    var rawNotes=data.f_notes||'';
    var stMatch=rawNotes.match(/^\[STATUS:([^\]]+)\]\n?/i);
    if(stMatch){
      stEl.value=stMatch[1].trim();
      // Strip the prefix from the visible notes field
      var notesInput=document.getElementById('f_notes');
      if(notesInput)notesInput.value=rawNotes.replace(/^\[STATUS:[^\]]+\]\n?/i,'');
    } else {
      stEl.value='Active';
    }
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
  aiTrack('ClientRecordOpened',{clientId:id}); // no PHI in telemetry — id only
  // The detail fetch below returns the DECRYPTED ssn, card, routing and account number, so
  // opening a record IS an ePHI access and has to be recorded as one.
  addAuditEntry(_auditName(c),'Client record opened');
  trackRecentRecord(id,c);
  // Deep-link URL so bookmarking / copy-link opens this client next time
  try{if(('#/client/'+id)!==window.location.hash)window.location.hash='/client/'+id;}catch(e){}
  editingId=id;
  _fullRecordFailed=false;
  _rowVersion=null; // cleared until the full record arrives with the real token
  try{clearForm();}catch(e){console.log('clearForm err:',e);}
  try{loadCarriersToSelect();}catch(e){}
  // Populate from the list first so the form isn't blank while the fetch is in flight,
  // then overwrite with the full record. The list deliberately omits SSN / card / bank,
  // so we MUST fetch the real values before any save — otherwise the save would write
  // the masked placeholders back over the real data.
  try{setFormData(c);}catch(e){console.log('setFormData err:',e);}
  clearFormDirty();
  fetch(API_BASE+'/health-clients/'+encodeURIComponent(id),{headers:apiHeaders()})
    .then(function(r){if(!r.ok)throw new Error('HTTP '+r.status);return r.json();})
    .then(function(row){
      // Ignore if the user navigated to a different record while this was loading
      if(String(editingId)!==String(id))return;
      _rowVersion=row.row_version_hex||null;
      try{setFormData(dbRowToClient(row));clearFormDirty();}catch(e){console.log('setFormData(full) err:',e);}
    })
    .catch(function(){
      if(String(editingId)!==String(id))return;
      // Block saving rather than risk writing blanks over real SSN / card / bank values.
      _fullRecordFailed=true;
      toast('Could not load the full record. Sensitive fields are hidden — saving is disabled until you reload.','error');
    });
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

  var exAudit=document.getElementById('clientAuditSection');
  if(exAudit)exAudit.remove();
  var audSec=document.createElement('div');
  audSec.className='form-section';
  audSec.id='clientAuditSection';
  formCard.insertBefore(audSec,actions);
  loadClientAudit(_auditName(c));
}

/* Per-client access history. Reads the SAME table the writes above go to, and the same one
   documents.js already writes Health document access into — so this shows real history
   immediately, before any of the new write sites have fired even once. */
function loadClientAudit(clientName){
  var sec=document.getElementById('clientAuditSection');
  if(!sec)return;
  sec.innerHTML='<div class="form-section-title">&#128274; Access History</div>'+
    '<p style="font-size:11px;color:#999;">Loading…</p>';
  fetch(API_BASE+'/audit/search',{
    method:'POST',headers:apiHeaders(),
    body:JSON.stringify({scope:'health',client:clientName,limit:100})
  })
  .then(_apiOk).then(function(r){return r.json();})
  .then(function(rows){renderClientAudit(clientName,rows||[]);})
  .catch(function(e){
    // Same rule as documents: a failed load must not read as "nothing ever happened".
    sec.innerHTML='<div class="form-section-title">&#128274; Access History</div>'+
      '<p style="font-size:12px;color:#b00;">Could not load the access history ('+
      escHtml((e&&e.message)||'error')+'). '+
      '<a href="#" onclick="loadClientAudit('+escJsAttr(JSON.stringify(clientName))+');return false;">Retry</a></p>';
  });
}
function renderClientAudit(clientName,rows){
  var sec=document.getElementById('clientAuditSection');if(!sec)return;
  var h='<div class="form-section-title">&#128274; Access History '+
        '<span style="font-size:10px;font-weight:normal;color:#999;">(who opened, changed or exported this record)</span></div>';
  if(!rows.length){
    h+='<p style="font-size:11px;color:#999;">No recorded activity for this client yet.</p>';
    sec.innerHTML=h; return;
  }
  h+='<div style="max-height:260px;overflow:auto;border:1px solid #eee;border-radius:6px;">';
  rows.forEach(function(r){
    var when=r.created_at?new Date(r.created_at).toLocaleString():'';
    var destructive=/DELETED/.test(r.action||'');
    h+='<div style="display:flex;gap:10px;padding:6px 10px;border-bottom:1px solid #f5f5f5;font-size:12px;'+
       (destructive?'background:#fff5f5;':'')+'">'+
       '<span style="flex:1;'+(destructive?'color:#b00;font-weight:600;':'')+'">'+escHtml(r.action||'')+'</span>'+
       '<span style="color:#666;white-space:nowrap;">'+escHtml(r.who||'')+'</span>'+
       '<span style="color:#999;white-space:nowrap;">'+escHtml(when)+'</span>'+
       '</div>';
  });
  h+='</div>';
  if(rows.length>=100)h+='<p style="font-size:10px;color:#999;margin-top:4px;">Showing the 100 most recent entries.</p>';
  sec.innerHTML=h;
}
function saveClient(onSuccess){
  var data=getFormData();
  if(!data.f_firstName&&!data.f_lastName){toast('Please enter at least a first or last name.','error');return;}
  // If the full record never loaded, the sensitive fields on screen are blank rather
  // than real. Saving would write those blanks over the stored SSN / card / bank values.
  if(_fullRecordFailed){toast('This record did not fully load. Reload the page before saving to avoid overwriting sensitive fields.','error');return;}
  var isNew=!editingId;
  saveClientAPI(data,editingId).then(function(res){
    // Keep the token current so a second save in the same sitting isn't rejected as stale.
    if(res&&res.row_version)_rowVersion=res.row_version;
    aiTrack(isNew?'ClientCreated':'ClientUpdated',{clientId:editingId||'new'}); // no PHI in telemetry
    addAuditEntry(_auditName(data),isNew?'Client record created':'Profile information updated');
    clearFormDirty();
    loadClients();
    if(typeof onSuccess==='function')onSuccess();
    else showView('clients');
    toast('Client saved!','success');
  }).catch(function(e){
    if(e&&e.status===409){
      // Someone else saved this record first. Deliberately do NOT clear the dirty flag, do NOT
      // reload, and do NOT navigate — the user's edit is still on screen and reloading here is
      // exactly what would discard it. Long toast because this needs a decision, not a glance.
      toast(e.message,'error',15000);
      return;
    }
    if(e&&e.status===404){
      toast('This client no longer exists — it was deleted elsewhere. Your changes were NOT saved.','error',15000);
      return;
    }
    toast('Could not save: '+((e&&e.message)||e),'error',10000);
  });
}
function deleteClient(){
  if(!editingId)return;
  // Pin the target BEFORE the dialog opens. The confirm is not modal to the browser: `hashchange`
  // fires on the Back button with no click on the page, and routeFromHash → editClient reassigns
  // editingId behind the open dialog. Reading editingId in the callback read it at OK-press time,
  // so a dialog naming one client could delete a different one — and file the audit row under the
  // name of the record that survived.
  var id=editingId;
  var c=clients.find(function(x){return String(x._id)===String(id);});
  var name=c?((c.f_firstName||'')+' '+(c.f_lastName||'')).trim():'this client';
  showConfirm('Delete '+(name||'this client')+'? This cannot be undone.',function(){
    // Logged BEFORE the delete: afterwards the record is gone and the name no longer resolves,
    // so the row would be filed under an id nobody can search for.
    addAuditEntry(name,'CLIENT RECORD DELETED by '+currentUserEmail());
    deleteClientAPI(id).then(function(){
      aiTrack('ClientDeleted',{clientId:id}); // no PHI in telemetry
      loadClients();showView('clients');
    }).catch(function(e){
      // Previously uncaught: a failed delete still navigated away as if it had worked.
      toast('Could not delete: '+(e&&e.message||e),'error');
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
/* Phone extension — always collapsed on load; only shown when the user clicks.
   The toggle itself surfaces the stored value ("Ext 4021") so a saved extension
   is still discoverable while the input stays hidden. Stored in its own column
   (phone_ext / alt_phone_ext) rather than packed into the phone string, which
   would break formatPhone(), the tel: links and phone search. */
function extLabel(w){
  var inp=document.getElementById('f_'+w+'Ext');
  var btn=document.getElementById('extToggle_'+w);
  if(!btn)return;
  // Always a bare "+" — no label text, no value shown.
  btn.textContent='+';
  btn.classList.toggle('active', !!(inp&&inp.style.display!=='none'));
}
function toggleExt(which){
  var inp=document.getElementById('f_'+which+'Ext');
  if(!inp)return;
  var showing=inp.style.display!=='none';
  inp.style.display=showing?'none':'';   // collapses even when it holds a value
  extLabel(which);
  if(!showing)inp.focus();
}
/* Called after the form is populated or cleared: keep both Ext inputs hidden
   regardless of content, and refresh the toggle labels. */
document.addEventListener('input',function(e){
  if(!e.target||!e.target.id)return;
  if(e.target.id==='f_phoneExt')extLabel('phone');
  else if(e.target.id==='f_altPhoneExt')extLabel('altPhone');
});
function syncExtVisibility(){
  ['phone','altPhone'].forEach(function(w){
    var inp=document.getElementById('f_'+w+'Ext');
    if(!inp)return;
    inp.style.display='none';
    extLabel(w);
  });
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
function copyField(id,btn){
  var el=document.getElementById(id);var t=el.type;el.type='text';
  // Card number / exp: strip formatting spaces + slashes so it pastes cleanly
  // into other forms that expect raw digits.
  var val=el.value;
  if(id==='f_cardNumber')val=val.replace(/\s+/g,'');
  navigator.clipboard.writeText(val);
  el.type=t;
  if(btn){btn.classList.add('copied');var o=btn.innerHTML;btn.innerHTML=SVG_CHECK;setTimeout(function(){btn.classList.remove('copied');btn.innerHTML=o;},1200);}
}
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
  row.innerHTML='<div class="field"><label>Other Income Source</label><input class="oi-src" value="'+escHtml(data.source||'')+'"></div>'+
    '<div class="field"><label>Income</label><input class="oi-amt" placeholder="$" value="'+escHtml(data.amount||'')+'" oninput="fmtMoney(this);calcTotalIncome()" onblur="fmtMoneyBlur(this);calcTotalIncome()"></div>'+
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
/* Remove one entry from a settings list by VALUE. Every caller reads an index at click time and
   acts on it after the dialog closes, and these lists can shift underneath an open dialog (a
   settings pull, or the same list edited in another tab). Splicing the stale index removed a
   neighbour instead. Returns false if the entry is already gone. */
function _removeByValue(arr,value){
  if(value===undefined)return false;
  var i=arr.indexOf(value);
  if(i<0)return false;
  arr.splice(i,1);
  return true;
}
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
/* HIPAA: closing the tab is the common way this app is left — far more common than pressing Sign
   out — and it used to leave everything clearCRMStorage clears sitting in the browser profile
   until the next explicit sign-out. Internal navigation is hash-based and does not fire pagehide,
   so this only runs on a real unload. Skipped when the page is going into bfcache, since it may
   be restored. Everything here is re-fetched from the server on the next visit. */
window.addEventListener('pagehide',function(e){
  if(e&&e.persisted)return;
  try{clearCRMStorage();}catch(_){}
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
/* Strip an ordinary list marker from the front of a pasted line: "1.", "2)", "-", "•", "*".
   A marker is only a marker when something follows it, and a NUMBER only counts when it is
   followed by . ) or ] — so "5-HTP 100mg" and "10mg Lipitor" keep their leading digits. */
function _stripListMarker(s){
  return s.replace(/^\s*(?:\d+\s*[.)\]]|[-•*\u2022])\s+/,'').trim();
}
/* Split a pasted line that holds MORE THAN ONE medication.

   "Metformin 500mg, Lisinopril 10mg" is two medications; "Metformin 500mg, twice daily" is one
   medication and its frequency. Same punctuation, opposite meanings. The tell is whether what
   follows the separator carries a dose of its own, so only split when every following part does.
   Splitting on every comma turned frequencies into medications; splitting on none absorbed a
   whole medication into the previous one's frequency field. Both were happening, in different
   places, on the same input. */
function _splitMedLine(line){
  var parts=String(line||'').split(/\s*[;,]\s*/).filter(function(p){return p.trim();});
  if(parts.length<2)return [String(line||'')];
  var DOSE=/\d+(?:\.\d+)?\s*(mg|mcg|g|units?|iu|ml)\b/i;
  return parts.slice(1).every(function(p){return DOSE.test(p);}) ? parts : [String(line||'')];
}
function parseMedLine(line){
  var s=_stripListMarker(String(line||'').replace(/[–—]/g,'-').replace(/\s+/g,' ').trim());
  if(!s)return null;
  // Extract the dose. A number WITH a unit wins outright, wherever it sits in the line: taking
  // the first number instead meant a name containing one ate the dose — "5-HTP 100mg" parsed as
  // the name "5-HTP 100mg" with no dose, and so did every numbered list line before the marker
  // above was stripped. Only when no unit appears anywhere do we fall back to a bare number, which
  // is what makes "Metformin 30 twice" work (30 → dose, "twice" → frequency).
  var doseMatch=s.match(/\b(\d+(?:\.\d+)?)\s*(mg|mcg|g|units?|iu|ml)\b/i)
              || s.match(/\b(\d+(?:\.\d+)?)\s*(mg|mcg|g|units?|iu|ml)?\b/i);
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
  lines.forEach(function(line){
    _splitMedLine(line).forEach(function(part){
      var m=parseMedLine(part);
      if(m&&(m.name||m.mg)){addMedRow(m);added++;}
    });
  });
  closeMedsPasteModal();
  toast('Imported '+added+' medication'+(added===1?'':'s'),'success');
}
function openDocsPasteModal(){var m=document.getElementById('docsPasteModal');if(m){document.getElementById('docsPasteInput').value='';m.style.display='flex';setTimeout(function(){document.getElementById('docsPasteInput').focus();},50);}}
function closeDocsPasteModal(){var m=document.getElementById('docsPasteModal');if(m)m.style.display='none';}
/* Parse one line into {name, specialty}. First token before a dash/comma/pipe/tab is name; rest joined is specialty. */
function parseDoctorLine(line){
  var s=_stripListMarker(String(line||'').replace(/[–—]/g,'-').replace(/\s+/g,' ').trim());
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
/* Auto-fill helper: only overwrite an existing value if we set it ourselves
   (tracked via data-autofilled). If the user manually typed/changed it, leave
   it alone. Called whenever the source field changes. */
function _isSafeToAutofill(el){return el&&(!el.value||el.dataset.autofilled==='1');}
function _setAutofilled(el,value){if(!el)return;el.value=value;el.dataset.autofilled='1';markFormDirty();}
/* When the user manually edits an auto-filled field, drop the autofilled flag
   so future auto-fills won't overwrite their edit. */
document.addEventListener('input',function(e){var t=e.target;if(t&&t.dataset&&t.dataset.autofilled==='1'&&t.id!=='f_routing'&&t.id!=='f_cardNumber')delete t.dataset.autofilled;});
document.addEventListener('change',function(e){var t=e.target;if(t&&t.dataset&&t.dataset.autofilled==='1'&&t.id!=='f_routing'&&t.id!=='f_cardNumber')delete t.dataset.autofilled;});

function autoDetectCardType(el){
  var first=(el.value||'').replace(/\D/g,'').charAt(0);
  var brand={'3':'Amex','4':'Visa','5':'Mastercard','6':'Discover'}[first];
  var sel=document.getElementById('f_cardType');
  if(!sel)return;
  // No first digit yet? Clear our previous auto-fill so the field goes empty again.
  if(!first){if(sel.dataset.autofilled==='1'){sel.value='';delete sel.dataset.autofilled;}return;}
  if(!brand)return;
  if(_isSafeToAutofill(sel))_setAutofilled(sel,brand);
}
/* ROUTING_LOOKUP is loaded from routing-lookup.js — full FedACH directory
   (18k+ US banks) generated from Moov's open-source github.com/moov-io/fed
   dataset which mirrors the Federal Reserve FedACH Participants Directory. */
/* Look up bank name from ABA routing number. Overwrites a previously auto-filled
   bank name when the routing changes, but leaves manually-typed names alone. */
function lookupBankFromRouting(rn){
  var digits=(rn||'').replace(/\D/g,'');
  var bankInput=document.getElementById('f_bankName');
  if(!bankInput)return;
  // Routing cleared or partial → clear our previous auto-fill (if the user
  // hadn't manually replaced it).
  if(digits.length<9){if(bankInput.dataset.autofilled==='1'){bankInput.value='';delete bankInput.dataset.autofilled;}return;}
  var name=ROUTING_LOOKUP[digits];if(!name)return;
  if(_isSafeToAutofill(bankInput)){
    _setAutofilled(bankInput,name);
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
    // STORED XSS: carrier names are operator-editable free text persisted in localStorage, and
    // both the element text and the JS-string attribute were unescaped. The ad-hoc
    // .replace(/'/g,"\\'") escaped neither a backslash nor '<' nor '"'. escJsAttr handles the
    // attribute case correctly (the HTML parser decodes entities BEFORE the JS runs, which is
    // why escHtml alone is not enough there). Matches referrerAC, which already did this right.
    return '<div onmousedown="document.getElementById(\'f_planCarrier\').value=\''+escJsAttr(c.name||'')+'\';document.getElementById(\'carrierACList\').style.display=\'none\';markFormDirty();">'+escHtml(c.name||'')+'</div>';
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
    html+='<div onmousedown="pickReferrer(\''+escJsAttr(e.name)+'\')" style="display:flex;justify-content:space-between;">'+
      '<span>'+escHtml(e.name)+'</span>'+
      '<span style="color:var(--text-muted);font-size:11px;">'+e.count+' referral'+(e.count===1?'':'s')+'</span></div>';
  });
  if(q&&!exact){
    html+='<div onmousedown="addReferrer(\''+escJsAttr(q)+'\')" style="border-top:1px solid var(--border);color:var(--accent);font-weight:600;">'+
      '+ Add new referrer: "'+escHtml(q)+'"</div>';
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
/* Resolve a zip to its county list, preferring the bundled ZIP_COUNTIES dataset
   (accurate multi-county coverage), falling back to FCC lat/lon lookup when the
   zip isn't in the bundle (~9k unlisted zips like some PO-Box-only). */
function bundledCounties(zip){
  if(typeof ZIP_COUNTIES==='undefined')return null;
  var v=ZIP_COUNTIES[zip];if(!v)return null;
  return v.split('|').map(function(s){return s.trim();}).filter(Boolean);
}
function lookupZip(el,prefix){
  var zip=el.value.replace(/\D/g,'');if(zip.length!==5)return;
  fetch('https://api.zippopotam.us/us/'+zip).then(function(r){return r.json();}).then(function(data){
    if(!data.places||!data.places.length)return;
    document.getElementById('f_'+prefix+'City').value=data.places[0]['place name']||'';
    document.getElementById('f_'+prefix+'St').value=data.places[0]['state abbreviation']||'';
    var sel=document.getElementById('f_'+prefix+'County');
    var bundled=bundledCounties(zip);
    if(bundled&&bundled.length&&sel){populateCountySel(sel,bundled,null);}
    else{fetchCountiesForPlaces(data.places,prefix,null);}
  }).catch(function(){});
}
function restoreCounty(zip,prefix,saved){
  var z=(zip||'').replace(/\D/g,'');if(z.length!==5)return;
  var sel=document.getElementById('f_'+prefix+'County');
  var bundled=bundledCounties(z);
  if(bundled&&bundled.length&&sel){populateCountySel(sel,bundled,saved);return;}
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
    '<div class="field"><label>Height</label><input data-field="height" value="'+escHtml(data&&data.height||'')+'" placeholder="5\'10&quot;" oninput="fmtHeight(this)"></div>'+
    '<div class="field"><label>Weight</label><input data-field="weight" value="'+escHtml(data&&data.weight||'')+'" placeholder="lbs"></div>'+
    '<div class="field"><label>DOB</label><input type="date" data-field="dob" id="'+uid+'_dob" value="'+escHtml(data&&data.dob||'')+'" onchange="calcMemberAge(this,\''+uid+'_age\')"></div>'+
    '<div class="field"><label>Age</label><input data-field="age" id="'+uid+'_age" readonly style="background:#f9f9f9;" value="'+escHtml(data&&data.age||'')+'"></div>'+
    '<div class="field"><label>SSN</label><input data-field="ssn" id="'+uid+'_ssn" type="password" placeholder="XXX-XX-XXXX" value="'+escHtml(data&&data.ssn||'')+'" oninput="formatSSN(this)" onfocus="focusReveal(this)" onblur="blurReveal(this)" maxlength="11"></div>'+
    mkSel('Insured','insured',['','Yes','No'],data)+
    '<button type="button" class="icon-btn" onclick="confirmRemoveRow(this,\'Remove this household member?\',updateMemberCount)" title="Remove"><svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="1.8" fill="none" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button>';
  document.getElementById('membersContainer').appendChild(div);
  updateMemberCount();
}
function mk(lbl,field,data){return '<div class="field"><label>'+lbl+'</label><input data-field="'+field+'" value="'+escHtml(data&&data[field]||'')+'"></div>';}
function mkS(lbl,field,data){return '<div class="field"><label>'+lbl+'</label><input data-field="'+field+'" value="'+escHtml(data&&data[field]||'')+'" maxlength="1"></div>';}
function mkSmall(lbl,field,data){
  var extra='';
  if(field==='height')extra=' oninput="fmtHeight(this)"';
  return '<div class="field"><label>'+lbl+'</label><input data-field="'+field+'" value="'+escHtml(data&&data[field]||'')+'"'+extra+'></div>';
}
function mkC(lbl,field,data){return '<div class="field"><label>'+lbl+'</label><input data-field="'+field+'" value="'+escHtml(data&&data[field]||'')+'"></div>';}
function mkSelC(lbl,field,opts,data){var val=data&&data[field]||'';var options=opts.map(function(o){return '<option'+(o===val?' selected':'')+'>'+o+'</option>';}).join('');return '<div class="field"><label>'+lbl+'</label><select data-field="'+field+'">'+options+'</select></div>';}
function mkSel(lbl,field,opts,data){
  var val=data&&data[field]||'';
  var options=opts.map(function(o){return '<option'+(o===val?' selected':'')+'>'+o+'</option>';}).join('');
  return '<div class="field"><label>'+lbl+'</label><select data-field="'+field+'">'+options+'</select></div>';
}
function addDoctorRow(data){
  var div=document.createElement('div');div.className='doctor-row-data';
  div.style.cssText='display:grid;grid-template-columns:2fr 1fr 30px;gap:6px;align-items:end;margin-bottom:6px;';
  div.innerHTML='<div class="field"><label>Doctor Name</label><input data-field="name" value="'+escHtml(data&&data.name||'')+'"></div>'+
    '<div class="field"><label>Specialty / Phone</label><input data-field="specialty" value="'+escHtml(data&&data.specialty||'')+'"></div>'+
    '<button type="button" class="icon-btn" onclick="confirmRemoveRow(this,\'Remove this doctor?\')" title="Remove"><svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="1.8" fill="none" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button>';
  document.getElementById('doctorsContainer').appendChild(div);
}

var _customMeds=[];
function loadCustomMeds(){try{_customMeds=JSON.parse(localStorage.getItem('crm_custom_meds')||'[]');}catch(e){_customMeds=[];}}
function saveCustomMed(name){if(!name)return;var n=name.trim();if(!n)return;if(MED_LIST.indexOf(n)===-1&&_customMeds.indexOf(n)===-1){_customMeds.push(n);_customMeds.sort();_syncedSetItem('crm_custom_meds',JSON.stringify(_customMeds));}}
loadCustomMeds();

var MED_LIST=['Abilify','Acarbose','Accolate','Accupril','Aciphex','Actonel','Actos','Adderall','Adderall XR','Advair','Advair Diskus','Aggrenox','Aldactone','Alendronate','Albuterol','Aleve','Allopurinol','Alprazolam','Altace','Ambien','Ambien CR','Amlodipine','Amoxicillin','Amoxicillin-Clavulanate','Amphetamine','Anastrozole','Androgel','Apixaban','Aripiprazole','Aspirin','Atenolol','Atomoxetine','Atorvastatin','Ativan','Augmentin','Azithromycin','Baclofen','Basaglar','Benadryl','Benazepril','Benicar','Benzonatate','Bisoprolol','Brilinta','Breo','Budesonide','Buprenorphine','Bupropion','Buspirone','Byetta','Bydureon','Caduet','Calcitriol','Carbamazepine','Carbidopa-Levodopa','Carvedilol','Celebrex','Celexa','Cephalexin','Cetirizine','Chantix','Cialis','Ciprofloxacin','Citalopram','Clindamycin','Clobetasol','Clonazepam','Clonidine','Clopidogrel','Colchicine','Colcrys','Combivent','Concerta','Coreg','Coreg CR','Coumadin','Cozaar','Crestor','Cyclobenzaprine','Cymbalta','Dapagliflozin','Dexamethasone','Dexilant','Dextroamphetamine','Diazepam','Diclofenac','Digoxin','Diltiazem','Diphenhydramine','Donepezil','Doxazosin','Doxycycline','Dulaglutide','Duloxetine','Dupixent','Effexor','Effexor XR','Eliquis','Empagliflozin','Enalapril','Entresto','Epidiolex','Escitalopram','Esomeprazole','Estradiol','Evista','Ezetimibe','Famotidine','Farxiga','Fentanyl','Ferrous Sulfate','Fexofenadine','Finasteride','Flagyl','Flexeril','Flomax','Flovent','Fluconazole','Fluoxetine','Fluticasone','Fluticasone-Salmeterol','Folic Acid','Fosamax','Furosemide','Gabapentin','Glimepiride','Glipizide','Glucophage','Glucotrol','Glyburide','Humalog','Humulin','Humulin N','Humulin R','Hydrochlorothiazide','Hydrocodone','Hydrocodone-Acetaminophen','Hydrocortisone','Hydroxychloroquine','Hydroxyzine','Ibuprofen','Invega','Invokamet','Invokana','Ipratropium','Irbesartan','Isosorbide','Janumet','Januvia','Jardiance','Juvisync','Ketamine','Klonopin','Lamictal','Lamotrigine','Lansoprazole','Lantus','Lantus SoloStar','Latuda','Levemir','Levofloxacin','Levothyroxine','Lexapro','Linagliptin','Linzess','Liraglutide','Lisinopril','Lisinopril-HCTZ','Lithium','Lopressor','Loratadine','Lorazepam','Losartan','Lovastatin','Lozol','Lyrica','Mavyret','Medroxyprogesterone','Meloxicam','Metformin','Metformin ER','Methocarbamol','Methylphenidate','Methylprednisolone','Metoprolol','Metoprolol Succinate','Metoprolol Tartrate','Metronidazole','Mirtazapine','Monjaro','Montelukast','Morphine','Mounjaro','Mucinex','Naproxen','Neurontin','Nexium','Nifedipine','Nitrofurantoin','Nitroglycerin','Norco','Nortriptyline','Novolin','Novolog','Novolog FlexPen','Nuvaring','Olmesartan','Omeprazole','Ondansetron','Oseltamivir','Ozempic','Oxycodone','Oxycodone-Acetaminophen','Oxycontin','Pantoprazole','Paroxetine','Paxil','Penicillin','Percocet','Phenergan','Phentermine','Plavix','Potassium Chloride','Pradaxa','Pravastatin','Prednisone','Pregabalin','Premarin','Prilosec','Pristiq','Proair','Prolia','Promethazine','Propranolol','Protonix','Provigil','Prozac','Quetiapine','Ramipril','Ranexa','Ranitidine','Reclipsen','Renvela','Repaglinide','Restasis','Rexulti','Risperidone','Ritalin','Rivaroxaban','Rosiglitazone','Rosuvastatin','Rybelsus','Saxenda','Semaglutide','Senna','Seroquel','Sertraline','Simvastatin','Singulair','Sitagliptin','Skyrizi','Solifenacin','Spironolactone','Strattera','Sulfamethoxazole','Sumatriptan','Symbicort','Synthroid','Tacrolimus','Tamsulosin','Temazepam','Testosterone','Tiotropium','Tizanidine','Topamax','Topiramate','Torsemide','Toujeo','Tramadol','Tradjenta','Trazodone','Tresiba','Trulicity','Valacyclovir','Valium','Valsartan','Venlafaxine','Ventolin','Vesicare','Viberzi','Victoza','Viibryd','Vimpat','Vitamin B12','Vitamin D','Voltaren','Vraylar','Warfarin','Wegovy','Wellbutrin','Xanax','Xarelto','Xifaxan','Xolair','Zestril','Zetia','Ziprasidone','Zofran','Zoloft','Zolpidem','Zopiclone','Zyprexa','Zyrtec'];
function getAllMeds(){return MED_LIST.concat(_customMeds).sort(function(a,b){return a.toLowerCase()<b.toLowerCase()?-1:1;});}
function addMedRow(data){
  var div=document.createElement('div');div.className='med-row-data';
  div.style.cssText='display:grid;grid-template-columns:2fr 80px 1fr auto 30px;gap:6px;align-items:end;margin-bottom:6px;position:relative;';
  div.innerHTML='<div class="field autocomplete-wrap"><label>Medication Name</label><input data-field="name" placeholder="Start typing..." value="'+escHtml(data&&data.name||'')+'" oninput="medAC(this)" onblur="medBlur(this)" autocomplete="off"><div class="autocomplete-list"></div></div>'+
    '<div class="field"><label>Mg</label><input data-field="mg" value="'+escHtml(data&&data.mg||'')+'"></div>'+
    '<div class="field"><label>Frequency</label><input data-field="frequency" value="'+escHtml(data&&data.frequency||'')+'"></div>'+
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
    '<div class="field"><label>Plan Name</label><input data-field="planName" value="'+escHtml(data.planName||'')+'"></div>'+
    '<div class="field"><label>Premium</label><input data-field="premium" placeholder="$" value="'+escHtml(data.premium||'')+'" oninput="fmtMoney(this);calcTotalMonthly()" onblur="fmtMoneyBlur(this);calcTotalMonthly()"></div>'+
    '<div class="field"><label>Pay Date</label><input type="date" data-field="payDate" value="'+escHtml(data.payDate||'')+'"></div>'+
    '<div class="field"><label>Effective</label><input type="date" data-field="effective" value="'+escHtml(data.effective||'')+'"></div>'+
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
        '<div class="field"><label>Previous Dental Carrier</label><input data-field="prevCarrier" value="'+escHtml(data.prevCarrier||'')+'"></div>'+
        '<div class="field"><label>Previous Member #</label><input data-field="prevMemberNum" value="'+escHtml(data.prevMemberNum||'')+'"></div>'+
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
    // Per-client rows before the deletes, for the same reason as the single delete.
    ids.forEach(function(id){
      var c=clients.find(function(x){return String(x._id)===String(id);});
      addAuditEntry(_auditName(c),'CLIENT RECORD DELETED by '+currentUserEmail()+' (bulk action)');
    });
    logActivity('delete',ids.length+' client records deleted in one bulk action by '+currentUserEmail());
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
    return '<div style="display:flex;justify-content:space-between;font-size:12px;padding:3px 0;border-bottom:1px solid #f0f3f7;"><span>'+escHtml(r.name)+'</span><strong style="color:var(--accent);">'+escHtml(r.count)+'</strong></div>';
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
  data.forEach(function(c){var tr=document.createElement('tr');tr.innerHTML='<td>'+escHtml(((c.f_firstName||'')+' '+(c.f_lastName||'')).trim())+'</td><td>'+escHtml(c.f_dob)+'</td><td>'+escHtml(c.f_phone)+'</td><td>'+escHtml(c.f_email)+'</td><td>'+escHtml(c.f_planType)+'</td><td>'+(c.f_premium?'$'+escHtml(c.f_premium):'')+'</td><td>'+escHtml(c.f_agent)+'</td>';tbody.appendChild(tr);});
  document.getElementById('reportResult').style.display='block';
}
function exportReportExcel(){
  // §164.528: a bulk extract of client records leaving the application is a disclosure and
  // needs a record of its own. exportAdvSearchExcel can include routing / account / card
  // columns, and had no trace at all.
  try{logActivity('export','report export downloaded by '+currentUserEmail());}catch(e){}
  var rows=[['Name','DOB','Phone','Email','Plan Type','Plan Name','Premium','Agent']];
  currentReportData.forEach(function(c){rows.push([(c.f_firstName||'')+' '+(c.f_lastName||''),c.f_dob||'',c.f_phone||'',c.f_email||'',c.f_planType||'',c.f_planName||'',c.f_premium||'',c.f_agent||'']);});
  dlXLSX(rows,'report.xlsx');
}
function exportExcel(){
  // §164.528: this sends the WHOLE roster — names, DOB, phone, email, plan, premium — out of the
  // application. The other three export paths recorded that; this one did not, so the largest
  // extract of the four was the only one leaving no trace.
  try{logActivity('export','client list export downloaded by '+currentUserEmail());}catch(e){}
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
  {key:'f_dob',label:'Date of Birth'},{key:'f_gender',label:'Gender'},
  // Full SSN is no longer returned by the list endpoint, so reports export the last 4 only.
  {key:'f_ssnLast4',label:'SSN (last 4)'},
  {key:'f_phone',label:'Phone'},{key:'f_email',label:'Email'},{key:'f_resAddress',label:'Address'},
  {key:'f_resCity',label:'City'},{key:'f_resSt',label:'State'},{key:'f_resZip',label:'Zip'},
  {key:'f_planName',label:'Plan Name'},{key:'f_planType',label:'Plan Type'},{key:'f_premium',label:'Premium'},
  {key:'f_subsidy',label:'Subsidy'},{key:'f_agent',label:'Agent'},{key:'f_leadSource',label:'Lead Source'},
  {key:'f_healthEffective',label:'Health Effective'},{key:'f_totalMonthly',label:'Total Monthly'},
  {key:'f_medicareNum',label:'Medicare #'},{key:'f_medicaid',label:'Medicaid #'},{key:'f_notes',label:'Notes'}
];
/* Split a CSV into headers + row objects, honouring quoted fields.

   The old parser was `line.split(',')` after `text.split('\n')`, with every quote character
   stripped afterwards. Any quoted field containing a comma — "123 Main St, Apt 4", or a notes
   field with a comma in it, both of which any real export produces — shifted EVERY column after
   it by one. The address became "123 Main St" and "Apt 4" was written into the next column, so
   an import silently wrote wrong values into patient records with no error anywhere.

   Handles: quoted commas, "" escapes, newlines inside quoted fields, CRLF and Excel's BOM. */
function parseCSV(text){
  text=String(text==null?'':text);
  var rows=[],row=[],field='',inQuotes=false;
  for(var i=0;i<text.length;i++){
    var ch=text.charAt(i);
    if(inQuotes){
      if(ch!=='"'){field+=ch;continue;}          // \r and \n inside quotes are DATA, not structure
      if(text.charAt(i+1)==='"'){field+='"';i++;} // "" is a literal quote
      else inQuotes=false;
    } else if(ch==='"'){inQuotes=true;}
    else if(ch===','){row.push(field);field='';}
    else if(ch==='\n'){row.push(field);field='';rows.push(row);row=[];}
    else if(ch!=='\r'){field+=ch;}               // bare \r only ever precedes the \n
  }
  row.push(field);rows.push(row);
  // trim() also removes Excel's leading BOM (U+FEFF is whitespace to trim), which is why there
  // is no separate BOM strip — without this the first header name would be '\uFEFFFirst Name'
  // and every lookup against it would miss.
  rows=rows.map(function(r){return r.map(function(v){return v.trim();});})
           .filter(function(r){return r.some(function(v){return v!=='';});});
  if(!rows.length)return {headers:[],rows:[]};
  var headers=rows[0];
  return {
    headers:headers,
    rows:rows.slice(1).map(function(vals){
      var obj={};headers.forEach(function(h,i){obj[h]=vals[i]||'';});return obj;
    })
  };
}
/* Column-mapping table. The option labels are the CSV's OWN header names — text from a file
   someone was sent — and they used to be concatenated into markup raw, so a header like
   <img src=x onerror=...> became a real element. Built as nodes now; nothing is parsed as HTML. */
function renderCsvMapping(headers){
  var tbody=document.getElementById('mappingBody');if(!tbody)return;
  tbody.innerHTML='';
  CRM_IMPORT_FIELDS.forEach(function(f){
    var tr=document.createElement('tr');
    var labelTd=document.createElement('td');labelTd.textContent=f.label;
    var selTd=document.createElement('td');
    var sel=document.createElement('select');sel.id='map_'+f.key;
    var skip=document.createElement('option');skip.value='';skip.textContent='-- Skip --';
    sel.appendChild(skip);
    var want=f.label.toLowerCase().replace(/[^a-z]/g,'').substr(0,4);
    (headers||[]).forEach(function(h){
      var o=document.createElement('option');
      o.value=h;o.textContent=h;
      if(String(h).toLowerCase().replace(/[^a-z]/g,'').includes(want))o.selected=true;
      sel.appendChild(o);
    });
    selTd.appendChild(sel);
    tr.appendChild(labelTd);tr.appendChild(selTd);
    tbody.appendChild(tr);
  });
}
function handleCSV(event){
  var file=event.target.files[0];if(!file)return;
  var reader=new FileReader();
  reader.onload=function(e){
    var parsed=parseCSV(e.target.result);
    csvHeaders=parsed.headers;csvData=parsed.rows;
    if(!csvHeaders.length){toast('That file had no readable rows.','error');return;}
    renderCsvMapping(csvHeaders);
    document.getElementById('mappingSection').style.display='block';
  };
  reader.onerror=function(){toast('Could not read that file.','error');};
  reader.readAsText(file);
}
function importClients(){
  // `.then(function(){imported++;})` counted every row as imported, because saveClientAPI
  // never rejected — a CSV where every row 500'd still reported "Imported 40 clients!".
  // Failures are counted separately, and the per-row rejection handler means one bad row
  // no longer aborts the rows after it.
  var imported=0,failed=0,firstError='';
  var promises=csvData.map(function(row){
    var data={};CRM_IMPORT_FIELDS.forEach(function(f){var col=document.getElementById('map_'+f.key);if(col&&col.value&&row[col.value]!==undefined)data[f.key]=row[col.value];});
    data.f_agent=data.f_agent||'Thomas Jaboro';
    return saveClientAPI(data,null).then(function(){imported++;},function(e){
      failed++;if(!firstError)firstError=String(e&&e.message||e);
    });
  });
  Promise.all(promises).then(function(){
    var el=document.getElementById('importStatus');
    if(failed){
      el.textContent='Imported '+imported+', FAILED '+failed+(firstError?(' — first error: '+firstError):'');
      toast(failed+' row'+(failed===1?'':'s')+' failed to import.','error');
    }else{
      el.textContent='Imported '+imported+' clients!';
    }
    loadClients();
  });
}

// CARRIER MANAGEMENT
function loadCarriers(){
  var saved=localStorage.getItem('crmCarriers');
  carriers=saved?JSON.parse(saved):[];
}
function saveCarriers(){
  _syncedSetItem('crmCarriers',JSON.stringify(carriers));
}
function addCarrier(){
  showPrompt('Add Carrier','Carrier name:','',function(name){
    if(!name||!name.trim())return;
    carriers.push({name:name.trim(),contact:'',phone:'',email:''});
    saveCarriers();
    renderCarriers();
  });
}
/* Universal quick-search modal (Cmd+K / Ctrl+K). Fuzzy-searches clients by name,
   phone, email, DOB, plan carrier. Keyboard nav + Enter to open. */
var _quickSearchIdx=0;
var _topSearchIdx=0;

// Shared matcher for both search surfaces (modal + permanent top bar).
function _searchClients(q){
  q=(q||'').trim().toLowerCase();
  return (clients||[]).filter(function(c){
    if(!q)return true;
    var hay=[
      (c.f_firstName||'')+' '+(c.f_lastName||''),
      c.f_phone||'',c.f_email||'',c.f_dob||'',
      c.f_planCarrier||'',c.f_planType||'',c.f_agent||''
    ].join(' ').toLowerCase();
    return hay.indexOf(q)!==-1;
  }).slice(0,10);
}
// Shared result-row markup. hoverJs runs on mouseenter to move the active index.
function _searchRowsHtml(results,activeIdx,hoverFn,pickFn){
  return results.map(function(c,i){
    var name=((c.f_firstName||'')+' '+(c.f_lastName||'')).trim()||'Unnamed';
    var sub=[c.f_phone,c.f_email,c.f_dob].filter(Boolean).join(' · ');
    var meta=[c.f_planType,c.f_planCarrier,c.f_agent].filter(Boolean).join(' · ');
    return '<div class="qs-row" data-id="'+escHtml(c._id)+'" onmouseenter="'+hoverFn+'('+i+')" onmousedown="'+pickFn+'()" style="padding:9px 14px;cursor:pointer;border-bottom:1px solid #f0f3f7;'+(i===activeIdx?'background:var(--accent-tint);':'')+'">'+
      '<div style="font-weight:600;font-size:13px;color:var(--text);">'+escHtml(name)+'</div>'+
      (sub?'<div style="font-size:11px;color:var(--text-muted);margin-top:1px;">'+escHtml(sub)+'</div>':'')+
      (meta?'<div style="font-size:11px;color:var(--text-subtle);margin-top:1px;">'+escHtml(meta)+'</div>':'')+
      '</div>';
  }).join('');
}

/* ---- Modal surface (sidebar Search button + Cmd/Ctrl+K) ---- */
function openQuickSearch(){
  var m=document.getElementById('quickSearchModal');if(!m)return;
  m.style.display='flex';
  var inp=document.getElementById('quickSearchInput');
  inp.value='';_quickSearchIdx=0;
  renderQuickSearch();
  setTimeout(function(){inp.focus();},30);
}
function closeQuickSearch(){var m=document.getElementById('quickSearchModal');if(m)m.style.display='none';}
function _qsHover(i){_quickSearchIdx=i;renderQuickSearch();}
function renderQuickSearch(){
  var q=document.getElementById('quickSearchInput').value||'';
  var box=document.getElementById('quickSearchResults');if(!box)return;
  var results=_searchClients(q);
  if(_quickSearchIdx>=results.length)_quickSearchIdx=Math.max(0,results.length-1);
  if(!results.length){
    box.innerHTML='<div style="padding:24px;text-align:center;color:var(--text-muted);font-size:13px;">'+(q.trim()?'No clients match "'+escHtml(q.trim())+'"':'Start typing to search '+(clients.length||0)+' clients')+'</div>';
    return;
  }
  box.innerHTML=_searchRowsHtml(results,_quickSearchIdx,'_qsHover','pickQuickSearch');
}
function pickQuickSearch(){
  var c=_searchClients(document.getElementById('quickSearchInput').value)[_quickSearchIdx];
  if(!c)return;
  closeQuickSearch();
  editClient(c._id);
}

/* ---- Permanent top-bar surface (always visible above every view) ---- */
function closeTopSearch(){
  var box=document.getElementById('topSearchResults');
  if(box)box.style.display='none';
}
function _tsHover(i){_topSearchIdx=i;renderTopSearch();}
function renderTopSearch(){
  var inp=document.getElementById('topSearchInput');
  var box=document.getElementById('topSearchResults');
  if(!inp||!box)return;
  var q=(inp.value||'').trim();
  // Only drop the panel once the user has typed — an empty bar stays quiet.
  if(!q){box.style.display='none';return;}
  var results=_searchClients(q);
  if(_topSearchIdx>=results.length)_topSearchIdx=Math.max(0,results.length-1);
  box.style.display='block';
  box.innerHTML=results.length
    ? _searchRowsHtml(results,_topSearchIdx,'_tsHover','pickTopSearch')
    : '<div style="padding:20px;text-align:center;color:var(--text-muted);font-size:13px;">No clients match "'+escHtml(q)+'"</div>';
}
function pickTopSearch(){
  var inp=document.getElementById('topSearchInput');
  var c=_searchClients(inp.value)[_topSearchIdx];
  if(!c)return;
  inp.value='';_topSearchIdx=0;
  closeTopSearch();inp.blur();
  editClient(c._id);
}
function topSearchKey(e){
  var box=document.getElementById('topSearchResults');
  var open=box&&box.style.display!=='none';
  if(e.key==='Escape'){
    e.preventDefault();
    if(open){closeTopSearch();}else{e.target.value='';e.target.blur();}
    return;
  }
  if(!open)return;
  if(e.key==='ArrowDown'){e.preventDefault();_topSearchIdx++;renderTopSearch();}
  else if(e.key==='ArrowUp'){e.preventDefault();_topSearchIdx=Math.max(0,_topSearchIdx-1);renderTopSearch();}
  else if(e.key==='Enter'){e.preventDefault();pickTopSearch();}
}
// Clicking anywhere outside the top-bar search closes its dropdown.
document.addEventListener('mousedown',function(e){
  var wrap=document.querySelector('.app-topbar .ts-wrap');
  if(wrap&&!wrap.contains(e.target))closeTopSearch();
});

// Global keybindings: Cmd/Ctrl+K opens the modal, Esc closes, arrows navigate
document.addEventListener('keydown',function(e){
  if((e.metaKey||e.ctrlKey)&&e.key.toLowerCase()==='k'&&_apiToken){
    e.preventDefault();openQuickSearch();return;
  }
  var m=document.getElementById('quickSearchModal');
  if(!m||m.style.display==='none')return;
  if(e.key==='Escape'){e.preventDefault();closeQuickSearch();}
  else if(e.key==='ArrowDown'){e.preventDefault();_quickSearchIdx++;renderQuickSearch();}
  else if(e.key==='ArrowUp'){e.preventDefault();_quickSearchIdx=Math.max(0,_quickSearchIdx-1);renderQuickSearch();}
  else if(e.key==='Enter'){e.preventDefault();pickQuickSearch();}
});
// Click outside the modal-card closes it
document.getElementById('quickSearchModal')&&document.getElementById('quickSearchModal').addEventListener('mousedown',function(e){if(e.target.id==='quickSearchModal')closeQuickSearch();});

/* In-app prompt modal — replaces native prompt() dialogs.
   Usage: showPrompt('Title','Label:','default val',function(val){...}, {okText:'Add'}) */
function showPrompt(title,message,defaultVal,onOk,opts){
  opts=opts||{};
  var modal=document.getElementById('promptModal');if(!modal)return;
  document.getElementById('promptTitle').textContent=title||'Enter value';
  document.getElementById('promptMessage').textContent=message||'';
  var input=document.getElementById('promptInput');
  input.value=defaultVal||'';
  input.placeholder=opts.placeholder||'';
  var okBtn=document.getElementById('promptOkBtn'),cancelBtn=document.getElementById('promptCancelBtn');
  okBtn.textContent=opts.okText||'OK';
  cancelBtn.textContent=opts.cancelText||'Cancel';
  // Replace handlers freshly each open so old callbacks don't stack
  var newOk=okBtn.cloneNode(true),newCancel=cancelBtn.cloneNode(true);
  okBtn.parentNode.replaceChild(newOk,okBtn);
  cancelBtn.parentNode.replaceChild(newCancel,cancelBtn);
  var close=function(){modal.style.display='none';input.onkeydown=null;};
  var submit=function(){var v=input.value;close();if(typeof onOk==='function')onOk(v);};
  newOk.addEventListener('click',submit);
  newCancel.addEventListener('click',function(){close();if(typeof opts.onCancel==='function')opts.onCancel();});
  input.onkeydown=function(e){if(e.key==='Enter'){e.preventDefault();submit();}else if(e.key==='Escape'){close();}};
  modal.style.display='flex';
  setTimeout(function(){input.focus();input.select();},50);
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
    div.innerHTML='<div class="field"><label>Carrier Name</label><input value="'+escHtml(c.name||'')+'" oninput="carriers['+idx+'].name=this.value;saveCarriers();"></div>'+
      '<div class="field"><label>Contact Person</label><input value="'+escHtml(c.contact||'')+'" oninput="carriers['+idx+'].contact=this.value;saveCarriers();"></div>'+
      '<div class="field"><label>Phone</label><input value="'+escHtml(c.phone||'')+'" oninput="carriers['+idx+'].phone=this.value;saveCarriers();"></div>'+
      '<button type="button" class="icon-btn" onclick="confirmRemoveCarrier('+idx+')" title="Remove"><svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="1.8" fill="none" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button>';
    container.appendChild(div);
  });
}
function confirmRemoveCarrier(idx){
  var c=carriers[idx];if(!c)return;
  showConfirm('Remove carrier "'+c.name+'"?',function(){
    // By identity, not by the position it held when the dialog opened.
    var j=carriers.indexOf(c);if(j<0)return;
    carriers.splice(j,1);saveCarriers();renderCarriers();
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
/* PHI-free recent records: persist only client IDs + access timestamps in
   localStorage. Names / plan / agent are resolved from the in-memory clients
   array at render time. Nothing sensitive ever hits disk. */
var _recentRecords=[]; // [{id, accessed}]
function loadRecentRecords(){
  try{
    var raw=JSON.parse(localStorage.getItem('crm_recent')||'[]');
    // Migration: old entries stored {id, name, planType, agent, accessed}. Strip everything but id + accessed.
    _recentRecords=raw.map(function(r){return {id:r.id,accessed:r.accessed||new Date().toISOString()};}).filter(function(r){return r.id;});
    // If any migration was needed, re-persist the sanitized shape immediately so old PHI is gone.
    if(raw.length&&raw[0]&&(raw[0].name||raw[0].planType||raw[0].agent))saveRecentRecords();
  }catch(e){_recentRecords=[];}
}
function saveRecentRecords(){localStorage.setItem('crm_recent',JSON.stringify(_recentRecords));}
loadRecentRecords();
function trackRecentRecord(id,_ignored){
  _recentRecords=_recentRecords.filter(function(r){return String(r.id)!==String(id);});
  _recentRecords.unshift({id:id,accessed:new Date().toISOString()});
  if(_recentRecords.length>20)_recentRecords=_recentRecords.slice(0,20);
  saveRecentRecords();
  renderSidebarRecent();
}
function renderRecentRecords(){
  var el=document.getElementById('recentRecordsList');
  var empty=document.getElementById('recentRecordsEmpty');
  if(!el)return;
  if(!_recentRecords.length){el.innerHTML='';empty.style.display='block';return;}
  empty.style.display='none';
  el.innerHTML='';
  _recentRecords.forEach(function(r,i){
    // Resolve display fields from the in-memory clients array; if a client was
    // deleted since it was last opened, show a muted placeholder and skip the click.
    var c=(clients||[]).find(function(x){return String(x._id)===String(r.id);});
    var name=c?((c.f_firstName||'')+' '+(c.f_lastName||'')).trim()||'Unnamed':'(client not found)';
    var planType=c?c.f_planType||'':'';
    var agent=c?c.f_agent||'':'';
    var when=new Date(r.accessed);
    var whenStr=when.toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'})+' '+when.toLocaleTimeString('en-US',{hour:'numeric',minute:'2-digit'});
    var div=document.createElement('div');
    div.style.cssText='display:flex;align-items:center;gap:12px;padding:10px 12px;border-bottom:1px solid #f0f0f0;'+(c?'cursor:pointer;transition:background 0.1s;':'opacity:0.55;');
    if(c){div.onmouseover=function(){this.style.background='#f0f4f8';};div.onmouseout=function(){this.style.background='';};}
    div.innerHTML=
      '<div style="width:24px;height:24px;border-radius:50%;background:#1a3a5c;color:#fff;font-size:11px;font-weight:bold;display:flex;align-items:center;justify-content:center;flex-shrink:0;">'+(i+1)+'</div>'+
      '<div style="flex:1;">'+
        '<div style="font-weight:600;font-size:13px;color:#1a3a5c;">'+escHtml(name)+'</div>'+
        '<div style="font-size:11px;color:#666;margin-top:2px;">'+(planType?'<span style="background:#dbeafe;color:#1a3a5c;padding:1px 6px;border-radius:8px;font-size:10px;margin-right:6px;">'+escHtml(planType)+'</span>':'')+escHtml(agent||'')+'</div>'+
      '</div>'+
      '<div style="font-size:10px;color:#999;white-space:nowrap;">'+whenStr+'</div>'+
      (c?'<button class="btn btn-blue" style="padding:4px 10px;font-size:11px;">Open</button>':'');
    if(c){
      div.querySelector('.btn').addEventListener('click',function(e){e.stopPropagation();editClient(r.id);});
      div.addEventListener('click',function(){editClient(r.id);});
    }
    el.appendChild(div);
  });
}
/* Sidebar Recent Records — last 5 openable clients, pinned above Sign Out.
   Same PHI rule as the full list: only {id, accessed} is persisted, names are
   resolved from the in-memory clients array here. Entries whose client no longer
   exists are skipped rather than shown as placeholders (no room in the sidebar). */
function renderSidebarRecent(){
  var el=document.getElementById('sbRecentList');if(!el)return;
  el.innerHTML='';
  var resolved=[];
  for(var i=0;i<_recentRecords.length&&resolved.length<5;i++){
    var r=_recentRecords[i];
    var c=(clients||[]).find(function(x){return String(x._id)===String(r.id);});
    if(c)resolved.push(c);
  }
  if(!resolved.length){
    var d=document.createElement('div');
    d.className='sb-recent-empty';
    d.textContent='No records opened yet';
    el.appendChild(d);
    return;
  }
  resolved.forEach(function(c){
    var name=((c.f_firstName||'')+' '+(c.f_lastName||'')).trim()||'Unnamed';
    var btn=document.createElement('button');
    btn.className='sb-recent-item';
    btn.title=name;
    var sq=document.createElement('span');sq.className='sb-recent-sq';
    var lbl=document.createElement('span');lbl.className='sb-recent-name';
    lbl.textContent=name; // textContent, not innerHTML — client names are user data
    btn.appendChild(sq);btn.appendChild(lbl);
    btn.addEventListener('click',function(){editClient(c._id);});
    el.appendChild(btn);
  });
}
function toggleSbRecent(){
  var box=document.getElementById('sbRecent');if(!box)return;
  var collapsed=box.classList.toggle('collapsed');
  localStorage.setItem('crm_sb_recent_collapsed',collapsed?'1':'0');
}
if(localStorage.getItem('crm_sb_recent_collapsed')==='1'){
  document.getElementById('sbRecent')&&document.getElementById('sbRecent').classList.add('collapsed');
}
function clearRecentRecords(){
  showConfirm('Clear recent records history?',function(){
    _recentRecords=[];saveRecentRecords();renderRecentRecords();renderSidebarRecent();
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
    d.innerHTML='<strong>'+escHtml(name)+'</strong><span style="font-size:10px;color:#666;margin-left:6px;">'+escHtml(c.f_planType)+'</span>';
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
  // clientName deliberately NOT persisted — resolved from `clients` in renderTodos
  var t={id:Date.now(),task:task,due:due,priority:priority,done:false,created:new Date().toISOString(),clientId:clientId};
  _todos.unshift(t);
  document.getElementById('todoAddSection').style.display='none';
  document.getElementById('todoClientInput').value='';
  document.getElementById('todoClientId').value='';
  renderTodos();
  // Show it immediately, then tell the truth about whether it actually saved. Nothing is kept on
  // this device any more, so a task that never reached the server is gone at the next load —
  // the agent has to be told that while the text is still on screen to re-enter.
  saveTaskAPI(t).then(function(){renderTodos();}).catch(function(e){
    t._unsaved=true;renderTodos();
    toast('Task NOT saved: '+((e&&e.message)||e)+'. It will disappear when you reload — please re-enter it.','error',15000);
  });
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
    tr.innerHTML='<td><span class="client-name-link" onclick="editClient(\''+escJsAttr(c._id)+'\')">'+escHtml(((c.f_firstName||'')+' '+(c.f_lastName||'')).trim())+'</span></td>'+
      '<td>'+escHtml(c.f_dob)+'</td><td>'+(age!==null?age:'')+'</td><td>'+escHtml(c.f_phone)+'</td><td>'+escHtml(c.f_email)+'</td>'+
      '<td>'+escHtml(c.f_planType)+'</td><td>'+escHtml(c.f_planCarrier)+'</td><td>'+escHtml(c.f_premium)+'</td>'+
      '<td>'+escHtml(c.f_agent)+'</td><td>'+escHtml(c.f_resSt)+'</td><td>'+escHtml(c.f_renewed)+'</td>';
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
  _syncedSetItem('crm_default_agent',sel.value);
}
function addPlanTypeSetting(){
  var v=document.getElementById('newPlanTypeInput').value.trim();if(!v)return;
  if(_settingsPlanTypes.indexOf(v)!==-1){toast('Already exists.','info');return;}
  _settingsPlanTypes.push(v);_syncedSetItem('crm_plan_types',JSON.stringify(_settingsPlanTypes));
  document.getElementById('newPlanTypeInput').value='';renderSettings();
}
function removePlanTypeSetting(i){
  var v=_settingsPlanTypes[i];
  showConfirm('Remove this plan type?',function(){
    if(!_removeByValue(_settingsPlanTypes,v))return;
    _syncedSetItem('crm_plan_types',JSON.stringify(_settingsPlanTypes));renderSettings();
  },{title:'Remove',okText:'Remove'});
}
function addProjectCodeSetting(){
  var v=document.getElementById('newProjectCodeInput').value.trim();if(!v)return;
  if(_settingsProjectCodes.indexOf(v)!==-1){toast('Already exists.','info');return;}
  _settingsProjectCodes.push(v);_syncedSetItem('crm_project_codes',JSON.stringify(_settingsProjectCodes));
  document.getElementById('newProjectCodeInput').value='';renderSettings();
}
function removeProjectCodeSetting(i){
  var v=_settingsProjectCodes[i];
  showConfirm('Remove this code?',function(){
    if(!_removeByValue(_settingsProjectCodes,v))return;
    _syncedSetItem('crm_project_codes',JSON.stringify(_settingsProjectCodes));renderSettings();
  },{title:'Remove',okText:'Remove'});
}
function exportFullBackup(){
  // §164.528: a bulk extract of client records leaving the application is a disclosure and
  // needs a record of its own. exportAdvSearchExcel can include routing / account / card
  // columns, and had no trace at all.
  try{logActivity('export','full backup export downloaded by '+currentUserEmail());}catch(e){}
  if(!clients.length){toast('No clients to export.','error');return;}
  var rows=[['First Name','Last Name','DOB','Phone','Email','Plan Type','Plan Name','Carrier','Premium','Subsidy','Total Monthly','App Fee','Agent','Lead Source','Renewed','State','City','ZIP','County','Medicare','Medicaid','Notes','App Date']];
  clients.forEach(function(c){
    rows.push([c.f_firstName||'',c.f_lastName||'',c.f_dob||'',c.f_phone||'',c.f_email||'',c.f_planType||'',c.f_planName||'',c.f_planCarrier||'',c.f_premium||'',c.f_subsidy||'',c.f_totalMonthly||'',c.f_appFee||'',c.f_agent||'',c.f_leadSource||'',c.f_renewed||'',c.f_resSt||'',c.f_resCity||'',c.f_resZip||'',c.f_resCounty||'',c.f_hasMedicare?'Yes':'No',c.f_hasMedicaid?'Yes':'No',c.f_notes||'',c.f_date||'']);
  });
  // fmtToday(), not toISOString(): the latter is UTC, so a backup taken after 8pm in Michigan
  // was filed under tomorrow's date — a disclosure record dated to the wrong day.
  dlXLSX(rows,'liberty_crm_backup_'+fmtToday()+'.xlsx');
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
  .catch(function(){
    // A FAILED load must never render as "no documents yet" — indistinguishable from genuinely
    // empty, and the agent then re-uploads a document that is already there.
    _clientDocs=[];
    var sec=document.getElementById('clientDocsSection');
    if(sec)sec.innerHTML='<div class="form-section-title">&#128196; Client Documents</div>'+
      '<p style="font-size:12px;color:#b00;">Could not load documents. '+
      '<a href="#" onclick="loadClientDocs(\''+String(clientId).replace(/[^A-Za-z0-9_-]/g,'')+'\');return false;">Retry</a> '+
      '— do not re-upload until this loads, the files may already be here.</p>';
  });
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
      // Only emit an https URL — never let a stored value inject a javascript: scheme.
      var safeUrl=/^https:\/\//i.test(d.url||'')?escHtml(d.url):'';
      row.innerHTML=icon+' <a href="'+safeUrl+'" target="_blank" rel="noopener noreferrer" style="flex:1;color:#1a3a5c;text-decoration:none;word-break:break-all;">'+escHtml(d.name)+'</a>'+
        '<span style="color:#999;font-size:10px;">'+kb+'</span>';
      // The filename is data from outside this app — it arrives on a file someone was sent, and the
      // backend only strips path separators and whitespace from it. It used to be interpolated into
      // an onclick="deleteClientDoc('id','name')" attribute via encodeURIComponent, which leaves
      // ' . ( ) untouched — enough for a name like  x'.concat(deleteClient())).concat('y.pdf  to
      // close the string literal and run. A real element with a real listener never parses it as code.
      var delBtn=document.createElement('button');
      delBtn.className='btn btn-red';
      delBtn.style.cssText='padding:2px 8px;font-size:10px;';
      delBtn.textContent='✕';
      delBtn.title='Delete document';
      delBtn.addEventListener('click',function(){deleteClientDoc(clientId,encodeURIComponent(d.name));});
      row.appendChild(delBtn);
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
    // Resolve to a per-file verdict rather than letting one rejection abandon the whole batch —
    // with Promise.all a single failure hid the fate of every other file in the selection.
    return fetch(API_BASE+'/documents',{method:'POST',headers:authUploadHeaders(),body:fd})
      .then(function(r){
        if(r.ok)return {ok:true,name:file.name};
        return r.json().catch(function(){return null;}).then(function(b){
          return {ok:false,name:file.name,msg:(b&&b.error)||('HTTP '+r.status)};
        });
      })
      .catch(function(e){return {ok:false,name:file.name,msg:(e&&e.message)||'network error'};});
  });
  Promise.all(promises)
  .then(function(results){
    var failed=results.filter(function(x){return !x.ok;});
    var okCount=results.length-failed.length;
    // fileCount was `fileNames.length` on a JOINED STRING — it has always reported the character
    // count of the concatenated filenames, not how many files went up.
    if(okCount)aiTrack('DocumentUploaded',{clientType:'health',clientId:clientId,fileCount:okCount}); // filenames can carry PHI
    if(failed.length){
      // Leave the file input populated. Clearing it is the app's only "done" signal, and the agent
      // needs to see which files still have to go up.
      status.textContent=failed.length+' of '+results.length+' upload(s) FAILED — not saved.';
      toast('Upload failed: '+failed.map(function(f){return f.name;}).join(', ')+
        ' ('+failed[0].msg+'). '+(failed.length===results.length?'Nothing was saved.':'The rest were saved.'),'error',15000);
    } else {
      status.textContent='';input.value='';
    }
    if(okCount)loadClientDocs(clientId);
  });
}
function deleteClientDoc(clientId,encodedName){
  showConfirm('Delete this document?',function(){
    // The name goes in the BODY, not the query string: a document filename routinely embeds the
    // patient's name, and URLs end up in access logs and request telemetry. The Home Care frontend
    // moved to this shape when the backend added it; this one was left on the legacy ?name= path.
    fetch(API_BASE+'/documents?clientType=health&clientId='+clientId,
      {method:'DELETE',headers:apiHeaders(),body:JSON.stringify({name:decodeURIComponent(encodedName)})})
    .then(_apiOk)
    // Only refresh on success. Refreshing either way made a refused delete look like one that
    // worked and simply left the file in place.
    .then(function(){loadClientDocs(clientId);})
    .catch(function(e){toast('Delete failed: '+((e&&e.message)||e)+' — the document is still there.','error',10000);});
  },{title:'Delete Document',okText:'Delete'});
}

// initMSAL called below
/* Street-address autocomplete REMOVED 2026-07-18. It sent partial patient street
   addresses to nominatim.openstreetmap.org — a third party with no BAA, in a URL
   query string. City/state/county still fill from the ZIP field (lookupZip), which
   uses the bundled ZIP_COUNTIES dataset. Do not reintroduce a client-side geocoder;
   proxy through our own API if this is ever wanted again. */

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
  // §164.528: a bulk extract of client records leaving the application is a disclosure and
  // needs a record of its own. exportAdvSearchExcel can include routing / account / card
  // columns, and had no trace at all.
  try{logActivity('export','advanced-search export downloaded by '+currentUserEmail());}catch(e){}
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
/* Same PHI rule as recent-records: only clientId is persisted, the name is
   resolved from the in-memory clients array at render time. Migration below
   strips clientName from any entry saved before this fix. */
/* ── TASKS API (source='health') ──────────────────────────────────────────────
   Tasks used to live in localStorage and nowhere else, which was wrong twice over.

   1. Task text is operator free-text — "call Jane about her Medicaid renewal" — so the list
      was PHI sitting at rest in the browser profile, the same rule crm_recent, crm_todos'
      clientName and the audit mirror were all already fixed for.
   2. clearCRMStorage wipes every crm_* key on sign-out, so the list was DESTROYED on every
      sign-out, and existed on one device only.

   The Tasks table has been there the whole time, discriminated by `source` and authorized per
   entity (a Home Care user cannot read, edit, move or delete a health row). Home Care keeps a
   localStorage cache on top of it; this app deliberately does not — the roster is already
   memory-only here, and tasks now follow it.

   `client_name` carries the client's ID for health rows, never the name: the UI resolves names
   from the in-memory roster at render time, and nothing needs the name in the column. */
function _taskToBody(t){
  return {
    id:t.dbId||undefined,
    text:t.task||'',
    done:t.done?1:0,
    due:t.due||null,
    client:t.clientId?String(t.clientId):'',   // ID, not name — see above
    priority:t.priority||'normal',
    source:'health'
  };
}
function saveTaskAPI(t){
  return fetch(API_BASE+'/tasks',{method:'POST',headers:apiHeaders(),body:JSON.stringify(_taskToBody(t))})
    .then(_apiOk).then(function(r){return r.json();})
    .then(function(res){
      if(!t.dbId&&res&&res.id)t.dbId=res.id;
      t._unsaved=false;
      return res;
    });
}
function deleteTaskAPI(dbId){
  if(!dbId)return Promise.resolve();
  return fetch(API_BASE+'/tasks/'+encodeURIComponent(dbId),{method:'DELETE',headers:apiHeaders()}).then(_apiOk);
}
function loadTasksAPI(){
  if(!_apiToken)return Promise.resolve();
  return fetch(API_BASE+'/tasks?source=health',{headers:apiHeaders()})
    .then(_apiOk).then(function(r){return r.json();})
    .then(function(rows){
      _todos=(Array.isArray(rows)?rows:[]).map(function(t){
        return {
          id:t.id, dbId:t.id,
          task:t.task_text||'',
          done:!!t.done,
          due:t.due_date?String(t.due_date).split('T')[0]:'',
          priority:t.priority||'normal',
          clientId:t.client_name||'',
          created:t.created_at||''
        };
      });
      // Newest first, matching the order the unshift-based local list produced.
      _todos.sort(function(a,b){return (b.dbId||0)-(a.dbId||0);});
      renderTodos();
    })
    .catch(function(e){
      // Never render an empty task list on a failed load — indistinguishable from having none.
      toast('Could not load tasks: '+((e&&e.message)||e)+'. This list may be incomplete.','error',10000);
    });
}
/* One-time move of whatever is still in localStorage on this device up to the server, BEFORE the
   key is dropped. Without this, shipping the change would delete every task the agent had. The key
   is only removed once every row it held has been accepted, so a failed migration retries next
   sign-in rather than losing the list. */
function migrateLegacyTodos(){
  var legacy=[];
  try{legacy=JSON.parse(localStorage.getItem('crm_todos')||'[]');}catch(e){legacy=[];}
  if(!Array.isArray(legacy)||!legacy.length){
    try{localStorage.removeItem('crm_todos');}catch(e){}
    return Promise.resolve();
  }
  return Promise.all(legacy.map(function(t){
    if(!t)return Promise.resolve(true);
    // Drop any dbId a legacy row somehow carries — these have never been on the server.
    return saveTaskAPI({task:t.task,done:t.done,due:t.due,priority:t.priority,clientId:t.clientId})
      .then(function(){return true;}).catch(function(){return false;});
  })).then(function(results){
    if(results.every(Boolean)){
      try{localStorage.removeItem('crm_todos');}catch(e){}
    } else {
      toast('Some tasks could not be moved to the server and are still only on this device. They will retry at next sign-in.','error',15000);
    }
  });
}
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
  if(!t)return;
  t.done=!t.done;
  renderTodos();
  saveTaskAPI(t).catch(function(e){
    // Put the tick back. A checkbox that stays ticked after a failed save is a lie the agent
    // acts on — the task reappears as outstanding at the next load with no explanation.
    t.done=!t.done;renderTodos();
    toast('Could not update that task: '+((e&&e.message)||e),'error',10000);
  });
}
function deleteTodo(id){
  var t=_todos.find(function(x){return x.id===id;});
  if(!t)return;
  showConfirm('Delete this task?',function(){
    // Remove locally only once the server has accepted it, or the task returns at the next load.
    deleteTaskAPI(t.dbId).then(function(){
      _todos=_todos.filter(function(x){return x!==t;});
      renderTodos();
    }).catch(function(e){
      toast('Could not delete that task: '+((e&&e.message)||e)+' — it is still there.','error',10000);
    });
  },{title:'Delete Task',okText:'Delete'});
}
function clearCompletedTodos(){
  // Pin the actual tasks now, not a count to re-derive after the dialog: the list can change
  // while it is open, and this used to re-filter _todos at OK-press time.
  var doneTasks=_todos.filter(function(x){return x.done;});
  if(!doneTasks.length){toast('No completed tasks to clear.','info');return;}
  showConfirm('Remove '+doneTasks.length+' completed task'+(doneTasks.length!==1?'s':'')+'?',function(){
    Promise.all(doneTasks.map(function(t){
      return deleteTaskAPI(t.dbId).then(function(){return t;}).catch(function(){return null;});
    })).then(function(results){
      var removed=results.filter(Boolean);
      _todos=_todos.filter(function(x){return removed.indexOf(x)<0;});
      renderTodos();
      var failed=results.length-removed.length;
      if(failed)toast(failed+' task'+(failed===1?'':'s')+' could not be removed and are still here.','error',10000);
    });
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
  // The due date comes from an <input type="date">, whose value is the LOCAL calendar day.
  // Comparing it against a UTC "today" made every evening wrong: from 8pm Michigan time until
  // midnight, toISOString() already reports tomorrow, so a task due today rendered as Overdue
  // and one due tomorrow rendered as Due Today.
  var today=fmtToday();
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
    if(t.clientId){
      // Resolve the name from memory — it is deliberately not stored in localStorage.
      var tc=(clients||[]).find(function(x){return String(x._id)===String(t.clientId);});
      var tcName=tc?(((tc.f_firstName||'')+' '+(tc.f_lastName||'')).trim()||'Unnamed'):'';
      if(tcName)clientLink='<span onclick="editClient(\''+escJsAttr(t.clientId)+'\')" style="font-size:10px;background:#dbeafe;color:#1a3a5c;padding:2px 8px;border-radius:10px;cursor:pointer;margin-left:6px;font-weight:600;" title="Open client record">&#128101; '+escHtml(tcName)+'</span>';
    }
    // A task the server never accepted is not coming back after a reload. Say so on the row
    // itself, not only in a toast that has already gone.
    var unsavedTag=t._unsaved?'<span style="font-size:10px;background:#fde2e2;color:#a01c1c;padding:2px 8px;border-radius:10px;margin-left:6px;font-weight:600;" title="This task was not saved to the server and will disappear when you reload.">NOT SAVED</span>':'';
    div.innerHTML=
      '<input type="checkbox" '+(t.done?'checked':'')+' style="width:16px;height:16px;cursor:pointer;flex-shrink:0;" onchange="toggleTodo('+t.id+')">'+
      '<div style="flex:1;">'+
        '<span style="font-size:13px;'+(t.done?'text-decoration:line-through;color:#999;':'')+'">'+escHtml(t.task)+'</span>'+dueStr+clientLink+unsavedTag+
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
      var today=fmtToday(); // local calendar day — see renderTodos
      var overdue=t.due&&t.due<today&&!t.done;
      var row=document.createElement('div');
      row.style.cssText='display:flex;align-items:center;gap:8px;padding:5px 0;border-bottom:1px solid #f5f5f5;font-size:12px;';
      row.innerHTML='<input type="checkbox" '+(t.done?'checked':'')+' onchange="toggleTodo('+t.id+');renderClientTodos(\''+clientId+'\')" style="cursor:pointer;">'+
        '<span style="flex:1;'+(t.done?'text-decoration:line-through;color:#999;':'')+'">'+escHtml(t.task)+'</span>'+
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
  _syncedSetItem('crm_agents',JSON.stringify(_settingsAgents));
  _syncedSetItem('crm_lead_sources',JSON.stringify(_settingsLeadSources));
  _syncedSetItem('crm_renewals',JSON.stringify(_settingsRenewals));
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
    // STORED XSS: `item` is operator-typed free text (agents, lead sources, custom meds,
    // renewal statuses, plan types, project codes) persisted to localStorage and re-rendered
    // on every Settings visit. Six entry points through this one function.
    div.innerHTML='<span style="flex:1;font-size:12px;">'+escHtml(item)+'</span>'+
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
function removeAgentSetting(i){var v=_settingsAgents[i];showConfirm('Remove this agent?',function(){if(!_removeByValue(_settingsAgents,v))return;saveSettings();renderSettings();populateDefaultAgentSelect();},{title:'Remove',okText:'Remove'});}
function addLeadSourceSetting(){var v=document.getElementById('newLeadSourceInput').value.trim();if(!v)return;if(_settingsLeadSources.indexOf(v)!==-1){toast('Lead source already exists.','info');return;}_settingsLeadSources.push(v);document.getElementById('newLeadSourceInput').value='';saveSettings();renderSettings();}
function removeLeadSourceSetting(i){var v=_settingsLeadSources[i];showConfirm('Remove this lead source?',function(){if(!_removeByValue(_settingsLeadSources,v))return;saveSettings();renderSettings();},{title:'Remove',okText:'Remove'});}
function addCustomMedSetting(){var v=document.getElementById('newCustomMedInput').value.trim();if(!v)return;saveCustomMed(v);document.getElementById('newCustomMedInput').value='';renderSettings();}
function removeCustomMedSetting(i){var v=_customMeds[i];showConfirm('Remove this medication?',function(){if(!_removeByValue(_customMeds,v))return;_syncedSetItem('crm_custom_meds',JSON.stringify(_customMeds));renderSettings();},{title:'Remove',okText:'Remove'});}
function addRenewalSetting(){var v=document.getElementById('newRenewalInput').value.trim();if(!v)return;if(_settingsRenewals.indexOf(v)!==-1){toast('Already exists.','info');return;}_settingsRenewals.push(v);document.getElementById('newRenewalInput').value='';saveSettings();renderSettings();}
function removeRenewalSetting(i){var v=_settingsRenewals[i];showConfirm('Remove this renewal option?',function(){if(!_removeByValue(_settingsRenewals,v))return;saveSettings();renderSettings();},{title:'Remove',okText:'Remove'});}
function saveCrmName(){var v=document.getElementById('settingsCrmName').value.trim();if(!v)return;_syncedSetItem('crm_display_name',v);document.querySelector('.sidebar .logo').childNodes[0].textContent=v;toast('CRM name updated!','success');}

purgeLegacyAuditLog();
try{initMSAL();loadCarriers();loadSettingsExtras();applySettingsToDropdowns();}catch(e){console.log('MSAL error:',e);showAuthScreen();}
