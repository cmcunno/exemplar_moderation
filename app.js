
/* Moderation Hub for Microsoft Teams.
   Sign-in: Microsoft Entra ID (nested app authentication inside Teams, normal sign-in in a browser).
   Data: SharePoint lists and a document library on the site named in config.js, reached through Microsoft Graph. */
const CONFIG = window.CONFIG || {};
const LEVELS = ['KS3','KS4','KS5'];
const DEFAULT_SUBJECTS = ['Biology','Chemistry','Physics','General Science'];
const ROLE_LABEL = {staff:'Teacher', hos:'Head of Subject', cc:'Curriculum Coordinator'};
const LIB = { mammoth:'lib/mammoth.browser.min.js', pdf:'lib/pdf.min.js', pdfWorker:'lib/pdf.worker.min.js' };
const LISTS = {staff:'Moderation Staff', questions:'Moderation Questions', marks:'Moderation Marks', settings:'Moderation Settings', files:'Moderation Files'};
const GRAPH = 'https://graph.microsoft.com/v1.0';
const SCOPES = ['User.Read','Sites.ReadWrite.All'];

const S = {
  phase:'starting', error:'', inTeams:false, me:null, myEmails:new Set(),
  staff:[], questions:[], markItems:[], markIndex:{}, settingsId:null, config:{tolerance:10, subjects:DEFAULT_SUBJECTS.slice()},
  lastSync:null, missingLists:[], busy:false,
  viewAs:'', route:'dashboard', qf:{level:'',subject:'',status:'',q:'',sort:'new'}, tf:{subject:'',level:''},
};
Object.defineProperty(S,'canEdit',{get:()=>isLead()});
Object.defineProperty(S,'canWrite',{get:()=>true});
Object.defineProperty(S,'assets',{get:()=>!!(Store.driveId && isLead())});

/* ---------- utilities ---------- */
const $ = (s,r=document)=>r.querySelector(s);
const esc = v => String(v ?? '').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const signed = (v,d=1) => (v==null||!isFinite(v)) ? '—' : (v>0.05?'+':v<-0.05?'−':'')+Math.abs(v).toFixed(d)+'%';
const fmtN = (v,d=1) => (v==null||!isFinite(v)) ? '—' : (Math.round(v*10**d)/10**d).toString();
const mean = a => a.length ? a.reduce((x,y)=>x+y,0)/a.length : null;
const sd = a => { if(a.length<2) return 0; const m=mean(a); return Math.sqrt(a.reduce((s,x)=>s+(x-m)**2,0)/(a.length-1)); };
const todayStr = () => new Date().toISOString().slice(0,10);
const daysUntil = d => d ? Math.round((new Date(d+'T23:59:59') - new Date())/86400000) : null;
const fmtDate = d => d ? new Date(d+'T12:00:00').toLocaleDateString('en-GB',{day:'numeric',month:'short'}) : '';
const hasRef = q => q.referenceMark!==null && q.referenceMark!==undefined && q.referenceMark!=='' && isFinite(q.referenceMark);

function toast(msg, err=false){
  const t=document.createElement('div'); t.className='toast'+(err?' err':''); t.textContent=msg;
  $('#toast-root').replaceChildren(t); setTimeout(()=>t.remove(), err?6000:3200);
}
function dbErr(e){
  const st=e&&e.status;
  if(st===401||st===403) return 'SharePoint refused this change. Coordinators and heads of subject must be owners of the Science team; teachers need Contribute on the Moderation Marks list (see the setup guide).';
  if(st===404) return 'That item no longer exists. The page has been refreshed.';
  if(st===429||st===503) return 'SharePoint is busy. Wait a few seconds and try again.';
  return 'The change was not saved: '+((e&&e.message)||'unknown error')+'.';
}
const loaded = {};
function loadScript(src){
  if(!loaded[src]) loaded[src]=new Promise((res,rej)=>{const s=document.createElement('script');s.src=src;s.onload=res;s.onerror=()=>{delete loaded[src];rej(new Error('Could not load '+src))};document.head.appendChild(s)});
  return loaded[src];
}

/* ---------- identity & roles ---------- */
const lc = v => String(v||'').trim().toLowerCase();
const meStaff = () => S.staff.find(s=>s.email && S.myEmails.has(lc(s.email))) || null;
const isAdminEmail = () => (CONFIG.adminEmails||[]).some(e=>S.myEmails.has(lc(e)));
function myRole(){ if(isAdminEmail()) return 'cc'; const s=meStaff(); return s ? s.role : null; }
const isCC = () => myRole()==='cc';
const isHoS = () => myRole()==='hos';
const isLead = () => isCC()||isHoS();
function scopeSubjects(){ if(isCC()) return S.config.subjects; if(isHoS()) return (meStaff()?.subjects||[]); return []; }
function teamStaff(){
  if(isCC()) return S.staff;
  if(isHoS()){ const subs=scopeSubjects(); return S.staff.filter(s=>(s.subjects||[]).some(x=>subs.includes(x))); }
  return [];
}
function targetStaff(){ if(S.viewAs && isLead()) return S.staff.find(s=>s.id===S.viewAs)||null; return meStaff(); }
const staffById = id => S.staff.find(s=>s.id===id);
const isMe = s => !!(s && meStaff() && s.id===meStaff().id);
function staffMarks(s){ return (s && S.markIndex[s.id]) || {}; }
function assignedTo(q,s){ return (s.levels||[]).includes(q.level) && (s.subjects||[]).includes(q.subject); }
function expectedMarkers(q){ return S.staff.filter(s=>assignedTo(q,s)); }
function canMarkFor(s){ if(!s) return false; if(isMe(s)) return true; return isCC(); }
/* A mark counts only if the person it belongs to created it, or a coordinator / head of subject entered it for them. */
function buildMarkIndex(){
  const leads=new Set([...(CONFIG.adminEmails||[]).map(lc), ...S.staff.filter(x=>x.role!=='staff').map(x=>lc(x.email)).filter(Boolean)]);
  const idx={};
  for(const m of [...S.markItems].sort((a,b)=>(a._mod||'').localeCompare(b._mod||''))){
    const st=staffById(m.staffId); if(!st || !isFinite(m.mark)) continue;
    const by=m._by; if(by && by!==lc(st.email) && !leads.has(by)) continue;
    (idx[st.id]=idx[st.id]||{})[m.questionId]={mark:+m.mark, note:m.note||'', at:m.at||m._mod, itemId:m.id};
  }
  S.markIndex=idx;
}

/* ---------- statistics ---------- */
function questionMarks(q){
  const out=[]; for(const s of S.staff){ const m=staffMarks(s)[q.id]; if(m && isFinite(m.mark)) out.push({s, mark:+m.mark, note:m.note||'', at:m.at}); } return out;
}
function qStats(q){
  const qm=questionMarks(q), ms=qm.map(x=>x.mark), max=+q.maxMarks||1;
  return {n:ms.length, expected:expectedMarkers(q).length, mean:mean(ms), sd:sd(ms), sdPct:sd(ms)/max*100, min:ms.length?Math.min(...ms):null, max:ms.length?Math.max(...ms):null, marks:qm};
}
/* One row per (marker, question) with at least one colleague to compare against.
   dev = (my mark − mean of the other markers) ÷ max marks × 100. Positive = lenient, negative = severe. */
function devRows(){
  const rows=[];
  for(const q of S.questions){
    const qm=questionMarks(q); if(qm.length<2) continue;
    const max=+q.maxMarks||1, total=qm.reduce((a,x)=>a+x.mark,0);
    for(const x of qm){
      const others=(total-x.mark)/(qm.length-1);
      rows.push({staffId:x.s.id, q, mark:x.mark, others, dev:(x.mark-others)/max*100, ref: hasRef(q) ? (x.mark-(+q.referenceMark))/max*100 : null});
    }
  }
  return rows;
}
function summarise(rows){
  const d=rows.map(r=>r.dev), refs=rows.filter(r=>r.ref!=null).map(r=>r.ref);
  return {n:rows.length, mean:mean(d), abs:mean(d.map(Math.abs)), ref:mean(refs), refN:refs.length};
}
function verdict(v){
  const tol=+S.config.tolerance||10;
  if(v==null) return '<span class="chip">Not enough data</span>';
  if(Math.abs(v)<=tol) return '<span class="chip ok">Within tolerance</span>';
  return v>0 ? '<span class="chip len">Lenient</span>' : '<span class="chip sev">Severe</span>';
}
function devBar(v){
  const tol=+S.config.tolerance||10, range=Math.max(25,tol*2.5);
  if(v==null) return `<div class="dev"><div class="dev-track"><div class="axis"></div></div><span class="dev-val muted">—</span></div>`;
  const pos=v=>50+Math.max(-1,Math.min(1,v/range))*50;
  const col=Math.abs(v)<=tol?'var(--ok)':v>0?'var(--lenient)':'var(--severe)';
  return `<div class="dev" title="Centre line = colleagues' average. Shaded band = ±${tol}% tolerance."><div class="dev-track"><div class="axis"></div><div class="band" style="left:${pos(-tol)}%;right:${100-pos(tol)}%"></div><div class="mid"></div><div class="dot" style="left:${pos(v)}%;background:${col}"></div></div><span class="dev-val" style="color:${col}">${signed(v)}</span></div>`;
}
function heatCell(sum){
  if(!sum || !sum.n) return `<span class="heat" style="color:var(--muted)">—<small>no data</small></span>`;
  const tol=+S.config.tolerance||10, v=sum.mean, strength=Math.min(1,Math.abs(v)/(tol*2.5));
  const base = Math.abs(v)<=tol ? 'var(--ok)' : v>0 ? 'var(--lenient)' : 'var(--severe)';
  const pct = Math.round(14+strength*36);
  return `<span class="heat" style="background:color-mix(in srgb, ${base} ${pct}%, var(--surface));color:var(--fg)">${signed(v)}<small>${sum.n} script${sum.n==1?'':'s'}</small></span>`;
}

/* ---------- sign-in ---------- */
let pca=null;
function configured(){ return CONFIG.clientId && !/PASTE/i.test(CONFIG.clientId) && CONFIG.tenantId && !/PASTE/i.test(CONFIG.tenantId) && CONFIG.siteUrl && !/YOURSCHOOL/i.test(CONFIG.siteUrl); }
async function initAuth(){
  try{
    await microsoftTeams.app.initialize(); S.inTeams=true;
    const ctx=await microsoftTeams.app.getContext();
    applyTeamsTheme(ctx.app && ctx.app.theme);
    microsoftTeams.app.registerOnThemeChangeHandler(applyTeamsTheme);
    S.loginHint=(ctx.user && (ctx.user.loginHint||ctx.user.userPrincipalName))||'';
  }catch(e){ S.inTeams=false; }
  const cfg={auth:{clientId:CONFIG.clientId, authority:'https://login.microsoftonline.com/'+CONFIG.tenantId, redirectUri:location.origin+location.pathname}, cache:{cacheLocation:'localStorage'}};
  pca = S.inTeams ? await msal.createNestablePublicClientApplication(cfg) : await msal.PublicClientApplication.createPublicClientApplication(cfg);
  if(!S.inTeams){ const r=await pca.handleRedirectPromise(); if(r&&r.account) pca.setActiveAccount(r.account); }
  let acct=pca.getActiveAccount()||pca.getAllAccounts()[0];
  if(!acct && S.inTeams){ try{ const r=await pca.ssoSilent({scopes:SCOPES, loginHint:S.loginHint}); acct=r.account; }catch(e){} }
  if(acct) pca.setActiveAccount(acct);
  return !!acct;
}
function applyTeamsTheme(t){ document.documentElement.setAttribute('data-theme', t==='dark'||t==='contrast' ? 'dark' : 'light'); }
async function signIn(){
  try{
    if(S.inTeams){ const r=await pca.acquireTokenPopup({scopes:SCOPES, loginHint:S.loginHint}); pca.setActiveAccount(r.account); start(); }
    else await pca.loginRedirect({scopes:SCOPES});
  }catch(e){ toast('Sign-in did not finish: '+(e.errorMessage||e.message||e),true); }
}
async function getToken(scopes=SCOPES){
  const account=pca.getActiveAccount();
  try{ return (await pca.acquireTokenSilent({scopes, account})).accessToken; }
  catch(e){
    if(S.inTeams){ const r=await pca.acquireTokenPopup({scopes, loginHint:S.loginHint}); pca.setActiveAccount(r.account); return r.accessToken; }
    await pca.acquireTokenRedirect({scopes}); throw new Error('Redirecting to sign in');
  }
}

/* ---------- Microsoft Graph ---------- */
async function graph(path, o={}){
  const token=await getToken(o.scopes);
  const headers={Authorization:'Bearer '+token, ...(o.json?{'Content-Type':'application/json'}:{}), ...(o.headers||{})};
  for(let attempt=0; attempt<3; attempt++){
    const r=await fetch(path.startsWith('http')?path:GRAPH+path,{method:o.method||'GET',headers,body:o.json?JSON.stringify(o.json):o.body});
    if(r.status===429||r.status===503){ await new Promise(res=>setTimeout(res,(+r.headers.get('Retry-After')||2)*1000)); continue; }
    if(r.status===204) return null;
    if(!r.ok){ let m=''; try{ m=(await r.json()).error.message; }catch(e){} const err=new Error(m||r.statusText); err.status=r.status; throw err; }
    return r.json();
  }
  const err=new Error('SharePoint is busy'); err.status=429; throw err;
}
async function graphAll(path){ let out=[], next=path; while(next){ const r=await graph(next); out=out.concat(r.value||[]); next=r['@odata.nextLink']||null; } return out; }

/* ---------- SharePoint store: each record is one list item, its fields kept as JSON in the Data column ---------- */
const Store = {
  siteId:null, ids:{}, driveId:null, urls:{},
  async init(){
    const u=new URL(CONFIG.siteUrl);
    const site=await graph(`/sites/${u.hostname}:${u.pathname.replace(/\/+$/,'')}?$select=id,displayName,webUrl`);
    this.siteId=site.id; this.siteWebUrl=site.webUrl;
    const lists=await graphAll(`/sites/${this.siteId}/lists?$select=id,displayName&$top=200`);
    this.ids={};
    for(const [k,name] of Object.entries(LISTS)){ const l=lists.find(x=>x.displayName===name); if(l) this.ids[k]=l.id; }
    if(this.ids.files){ this.driveId=(await graph(`/sites/${this.siteId}/lists/${this.ids.files}/drive?$select=id`)).id; }
    return Object.keys(LISTS).filter(k=>!this.ids[k]);
  },
  async createLists(missing){
    for(const k of missing){
      const body = k==='files' ? {displayName:LISTS[k], list:{template:'documentLibrary'}}
        : {displayName:LISTS[k], columns:[{name:'Data', text:{allowMultipleLines:true, textType:'plain', linesForEditing:6}}], list:{template:'genericList'}};
      await graph(`/sites/${this.siteId}/lists`,{method:'POST', json:body, scopes:['Sites.Manage.All']});
    }
  },
  async items(k){
    const rows=await graphAll(`/sites/${this.siteId}/lists/${this.ids[k]}/items?$select=id,createdBy,lastModifiedDateTime&$expand=fields($select=Title,Data)&$top=500`);
    return rows.map(r=>{ let d={}; try{ d=JSON.parse((r.fields&&r.fields.Data)||'{}'); }catch(e){}
      return {...d, id:r.id, _by:lc(r.createdBy&&r.createdBy.user&&r.createdBy.user.email), _mod:r.lastModifiedDateTime}; });
  },
  async create(k,title,data){ return (await graph(`/sites/${this.siteId}/lists/${this.ids[k]}/items`,{method:'POST',json:{fields:{Title:String(title||'').slice(0,255),Data:JSON.stringify(data)}}})).id; },
  async update(k,id,title,data){ await graph(`/sites/${this.siteId}/lists/${this.ids[k]}/items/${id}/fields`,{method:'PATCH',json:{Title:String(title||'').slice(0,255),Data:JSON.stringify(data)}}); },
  async remove(k,id){ await graph(`/sites/${this.siteId}/lists/${this.ids[k]}/items/${id}`,{method:'DELETE'}); },
  async upload(file){
    const safe=(file.name||'file').replace(/[^\w.\- ]+/g,'_');
    const path=`/drives/${this.driveId}/root:/${encodeURIComponent(Date.now()+'-'+safe)}`;
    if(file.size<=4*1024*1024) return graph(path+':/content',{method:'PUT',body:file,headers:{'Content-Type':file.type||'application/octet-stream'}});
    const sess=await graph(path+':/createUploadSession',{method:'POST',json:{item:{'@microsoft.graph.conflictBehavior':'rename'}}});
    const chunk=320*1024*16; let res;
    for(let start=0; start<file.size; start+=chunk){
      const end=Math.min(start+chunk,file.size);
      const r=await fetch(sess.uploadUrl,{method:'PUT',headers:{'Content-Range':`bytes ${start}-${end-1}/${file.size}`},body:file.slice(start,end)});
      if(!r.ok){ const e=new Error('Upload failed'); e.status=r.status; throw e; }
      res=await r.json().catch(()=>null);
    }
    return res;
  },
  async download(id){
    const it=await graph(`/drives/${this.driveId}/items/${id}?$select=id,@microsoft.graph.downloadUrl`);
    const r=await fetch(it['@microsoft.graph.downloadUrl']); if(!r.ok) throw new Error('Download failed'); return r.blob();
  },
  fileUrl(id){ return this.urls[id] || (this.urls[id]=this.download(id).then(b=>URL.createObjectURL(b)).catch(e=>{ delete this.urls[id]; throw e; })); },
  async fileBuffer(id){ return (await this.download(id)).arrayBuffer(); },
};
function hydrateFiles(root){
  (root||document).querySelectorAll('img[data-file]:not([src])').forEach(img=>{
    Store.fileUrl(img.dataset.file).then(u=>{ img.src=u; }).catch(()=>{ img.replaceWith(Object.assign(document.createElement('span'),{className:'small muted',textContent:'Image could not be loaded'})); });
  });
}

/* ---------- loading and writing ---------- */
const REC_INTERNAL=['id','_by','_mod'];
function recData(r){ const d={...r}; REC_INTERNAL.forEach(k=>delete d[k]); return d; }
function titleFor(kind,d){
  if(kind==='staff') return d.name; if(kind==='questions') return d.title; if(kind==='settings') return 'Settings';
  if(kind==='marks') return `${staffById(d.staffId)?.name||'Staff'} · ${S.questions.find(q=>q.id===d.questionId)?.title||'Question'}`;
  return '';
}
const arr = kind => kind==='marks' ? S.markItems : S[kind];
async function dbCreate(kind,data){ return Store.create(kind,titleFor(kind,data),data); }
async function dbUpdate(kind,id,patch){ const cur=(arr(kind)||[]).find(x=>x.id===id); const data={...(cur?recData(cur):{}),...patch}; await Store.update(kind,id,titleFor(kind,data),data); }
async function dbDelete(kind,id){ await Store.remove(kind,id); }

async function refresh(){
  if(S.phase!=='ready' || S.busy) return;
  S.busy=true;
  try{
    const [staff,questions,marks,settings]=await Promise.all(['staff','questions','marks','settings'].map(k=>Store.items(k)));
    S.staff=staff.sort((a,b)=>(a.name||'').localeCompare(b.name||''));
    S.questions=questions; S.markItems=marks;
    const st=settings[0]; S.settingsId=st?st.id:null;
    S.config={tolerance:(st&&+st.tolerance)||10, subjects:(st&&st.subjects&&st.subjects.length)?st.subjects:DEFAULT_SUBJECTS.slice()};
    buildMarkIndex(); S.lastSync=new Date();
  }catch(e){ toast('Could not refresh from SharePoint: '+(e.message||e),true); }
  S.busy=false; render();
}
async function start(){
  S.phase='loading'; render();
  try{
    const me=await graph('/me?$select=displayName,mail,userPrincipalName');
    S.me={name:me.displayName||'', email:me.mail||me.userPrincipalName||''};
    S.myEmails=new Set([lc(me.mail),lc(me.userPrincipalName)].filter(Boolean));
    S.missingLists=await Store.init();
    if(S.missingLists.length){ S.phase='setup'; render(); return; }
    S.phase='ready'; await refresh();
  }catch(e){ S.phase='error'; S.error=(e.status===403||e.status===404)?'This account can’t open the SharePoint site in config.js. Check the address is right and that you’ve been added to the site.':(e.message||String(e)); render(); }
}
async function boot(){
  render();
  if(!configured()){ S.phase='unconfigured'; render(); return; }
  try{ const ok=await initAuth(); if(!ok){ S.phase='signin'; render(); return; } }
  catch(e){ S.phase='error'; S.error='Sign-in could not start: '+(e.errorMessage||e.message||e); render(); return; }
  await start();
  setInterval(()=>{ if(document.visibilityState==='visible' && !$('#modal-root').innerHTML) refresh(); }, 60000);
  document.addEventListener('visibilitychange',()=>{ if(document.visibilityState==='visible') refresh(); });
}
async function createLists(){
  const btn=$('#create-lists'); if(btn){ btn.disabled=true; btn.textContent='Creating lists…'; }
  try{ await Store.createLists(S.missingLists); toast('SharePoint lists created'); await start(); }
  catch(e){ toast(dbErr(e),true); if(btn){ btn.disabled=false; btn.textContent='Create the SharePoint lists'; } }
}

async function saveMark(s,qid,mark,note){
  const cur=staffMarks(s)[qid];
  const data={staffId:s.id, questionId:qid, mark, note:note||'', at:new Date().toISOString(), by:S.me.email};
  const item=cur && S.markItems.find(m=>m.id===cur.itemId);
  /* teachers can only edit items they created; if a coordinator entered this mark, the teacher's change is saved as a newer item */
  if(item && (isLead() || item._by===lc(S.me.email) || !item._by)) await dbUpdate('marks',item.id,data);
  else await dbCreate('marks',data);
}
async function removeMark(s,qid){
  const mine=S.markItems.filter(m=>m.staffId===s.id && m.questionId===qid);
  for(const m of mine){ if(isLead() || m._by===lc(S.me.email)) await dbDelete('marks',m.id); }
}

/* ---------- rendering ---------- */
function navItems(){
  const items=[{r:'dashboard',l:'My marking'}];
  if(isLead()){ items.push({r:'questions',l:'Questions'}); items.push({r:'team',l:'Team performance'}); }
  if(isLead()) items.push({r:'staff',l:'Staff & settings'});
  return items;
}
function todoFor(s){
  if(!s) return [];
  const mine=staffMarks(s);
  return S.questions.filter(q=>q.status!=='closed' && assignedTo(q,s) && !mine[q.id])
    .sort((a,b)=>(a.deadline||'9999').localeCompare(b.deadline||'9999'));
}
function phaseScreen(){
  const card=(t,b)=>`<div class="page-head"><div><div class="eyebrow">Moderation Hub</div><h1>${t}</h1></div></div><div class="card">${b}</div>`;
  switch(S.phase){
    case 'starting': case 'loading': return `<div class="loading">Connecting to SharePoint…</div>`;
    case 'unconfigured': return card('Setup isn’t finished',`<p>Open <b>config.js</b> and fill in the app’s client ID, tenant ID, SharePoint site address and the curriculum coordinator’s email. The setup guide explains where each value comes from.</p>`);
    case 'signin': return card('Sign in to continue',`<p>Sign in with your school Microsoft account to see your marking.</p><div><button class="btn primary" data-act="signin">Sign in with Microsoft</button></div>`);
    case 'error': return card('Something needs fixing',`<p>${esc(S.error)}</p><div class="btn-row"><button class="btn primary" data-act="retry">Try again</button></div>`);
    case 'setup': return isAdminEmail()
      ? card('Create the SharePoint lists',`<p>The site is connected. The app now needs its lists: ${S.missingLists.map(k=>'<b>'+esc(LISTS[k])+'</b>').join(', ')}.</p><div><button class="btn primary" id="create-lists" data-act="create-lists">Create the SharePoint lists</button></div><p class="small muted">After this, follow step 7 of the setup guide to set permissions on the new lists.</p>`)
      : card('Not set up yet',`<p>Your curriculum coordinator hasn’t finished setting up the Moderation Hub. Please check back later.</p>`);
  }
  return '';
}
function render(){
  const ae=document.activeElement, aeId=ae&&ae.id, selS=ae&&ae.selectionStart, selE=ae&&ae.selectionEnd;
  if(S.phase!=='ready'){ $('#nav').innerHTML=''; $('#who').innerHTML=whoHtml(); $('#main').innerHTML=phaseScreen(); return; }
  const allowed=navItems().map(i=>i.r);
  const h=(location.hash||'').slice(1);
  S.route = allowed.includes(h) ? h : 'dashboard';
  const todoN = meStaff()? todoFor(meStaff()).length : 0;
  $('#nav').innerHTML = navItems().map(i=>`<a href="#${i.r}" ${S.route===i.r?'aria-current="page"':''}>${i.l}${i.r==='dashboard'&&todoN?`<span class="count">${todoN}</span>`:''}</a>`).join('');
  $('#who').innerHTML = whoHtml();
  $('#main').innerHTML = ({dashboard:viewDashboard, questions:viewQuestions, team:viewTeam, staff:viewStaff}[S.route])();
  hydrateFiles($('#main'));
  if(aeId){ const el=document.getElementById(aeId); if(el&&el!==document.activeElement){ el.focus(); try{ if(selS!=null) el.setSelectionRange(selS,selE);}catch(e){} } }
}
function whoHtml(){
  if(!S.me) return '';
  const role=myRole(), name=S.me.name || meStaff()?.name || 'You';
  const initials=name.split(/\s+/).filter(Boolean).map(w=>w[0]).slice(0,2).join('').toUpperCase();
  let h=`<div class="who-row"><span class="avatar" aria-hidden="true">${esc(initials)}</span><div style="min-width:0"><div class="who-name">${esc(name)}</div><div class="small muted">${role?ROLE_LABEL[role]:'Not on the staff list'}</div></div></div>`;
  if(S.phase==='ready' && isLead()){
    const opts=teamStaff().filter(s=>!isMe(s));
    h+=`<div class="view-as"><label class="small muted" for="viewas">View a dashboard as</label><select id="viewas" data-change="viewas"><option value="">Myself</option>${opts.map(s=>`<option value="${s.id}" ${S.viewAs===s.id?'selected':''}>${esc(s.name)}</option>`).join('')}</select></div>`;
  }
  if(S.phase==='ready') h+=`<div class="sync small muted"><span>${S.lastSync?'Updated '+S.lastSync.toLocaleTimeString('en-GB',{hour:'2-digit',minute:'2-digit'}):''}</span><button class="linkish small" data-act="refresh">Refresh</button></div>`;
  return h;
}
/* ---- My marking dashboard ---- */
function viewDashboard(){
  const s=targetStaff(), viewingOther = S.viewAs && s && s.id!==meStaff()?.id;
  if(!s){
    if(isCC()) return onboarding();
    return `<div class="page-head"><div><div class="eyebrow">My marking</div><h1>You're not on the staff list yet</h1><p class="sub">Your curriculum coordinator needs to add you to the staff list before marking tasks appear here.</p></div></div>
      <div class="card"><p>Ask your curriculum coordinator to add you in <b>Staff &amp; settings</b> using your school email address (<b>${esc(S.me?.email||'')}</b>). Your list of scripts to mark will then appear here.</p></div>`;
  }
  const todo=todoFor(s), mine=staffMarks(s);
  const rows=devRows().filter(r=>r.staffId===s.id), all=summarise(rows);
  const markedQs=S.questions.filter(q=>mine[q.id]);
  const assignedAll=S.questions.filter(q=>assignedTo(q,s));
  const overdue=todo.filter(q=>q.deadline && daysUntil(q.deadline)<0).length;
  let h=`<div class="page-head"><div><div class="eyebrow">${viewingOther?'Viewing as':'My marking'}</div><h1>${esc(s.name)}</h1>
    <p class="sub">${ROLE_LABEL[s.role]||'Teacher'} · ${(s.levels||[]).map(l=>`<span class="chip ks">${l}</span>`).join(' ')} ${(s.subjects||[]).map(x=>`<span class="chip">${esc(x)}</span>`).join(' ')}</p></div>
    ${viewingOther?`<button class="btn" data-act="viewas-clear">Back to my dashboard</button>`:''}</div>`;
  h+=`<div class="tiles">
    <div class="tile"><span class="v">${todo.length}</span><span class="l">Scripts to mark${overdue?` · <b style="color:var(--danger)">${overdue} overdue</b>`:''}</span></div>
    <div class="tile"><span class="v">${markedQs.length}<span class="muted" style="font-size:1rem"> / ${assignedAll.length}</span></span><span class="l">Moderation scripts marked</span></div>
    <div class="tile"><span class="v" style="color:${all.mean==null?'var(--fg)':Math.abs(all.mean)<=S.config.tolerance?'var(--ok)':all.mean>0?'var(--lenient)':'var(--severe)'}">${signed(all.mean)}</span><span class="l">Overall variation from colleagues</span></div>
    <div class="tile"><span class="v">${all.abs==null?'—':fmtN(all.abs)+'%'}</span><span class="l">Average size of difference</span></div>
  </div>`;
  h+=`<div class="grid-2"><section class="card"><div class="card-head"><h2>To do</h2><span class="small muted">${todo.length?'Earliest deadline first':''}</span></div>${todoList(todo, s)}</section>`;
  h+=`<section class="card"><div class="card-head"><h2>Marked recently</h2></div>${recentList(s, mine)}</section></div>`;
  if(!viewingOther && isLead()) h+=leadTasks();
  h+=`<section class="card"><div class="card-head"><div><h2>My marking compared with colleagues</h2><p class="small muted">Each script you mark is compared with the average of the other markers on the same script, as a % of the marks available. Positive = more generous, negative = harsher. Tolerance: ±${S.config.tolerance}%.</p></div></div>${variationTable(rows)}</section>`;
  return h;
}
function todoList(todo, s){
  if(!todo.length) return `<div class="empty"><b>Nothing to mark</b><span class="small">New moderation scripts for ${esc((s.levels||[]).join(', ')||'your key stages')} appear here when a head of subject adds them.</span></div>`;
  return `<div class="todo">${todo.map(q=>{ const d=daysUntil(q.deadline); const due = q.deadline ? (d<0?`<span class="chip danger">Overdue · ${fmtDate(q.deadline)}</span>`: d<=3?`<span class="chip warn">Due ${d===0?'today':fmtDate(q.deadline)}</span>`:`<span class="chip">Due ${fmtDate(q.deadline)}</span>`) : '';
    return `<div class="todo-item ${d!=null&&d<0?'overdue':''}"><span class="box" aria-hidden="true"></span><div class="body"><span class="t">${esc(q.title)}</span><span class="meta"><span class="chip ks">${q.level}</span><span class="chip">${esc(q.subject)}</span>${q.topic?`<span class="small muted">${esc(q.topic)}</span>`:''}<span class="small muted">${q.maxMarks} marks</span>${due}</span></div><button class="btn primary sm" data-act="mark" data-id="${q.id}">${canMarkFor(s)?'Mark':'Open'}</button></div>`;}).join('')}</div>`;
}
function recentList(s, mine){
  const items=S.questions.filter(q=>mine[q.id]).sort((a,b)=>(mine[b.id].at||'').localeCompare(mine[a.id].at||'')).slice(0,6);
  if(!items.length) return `<div class="empty"><b>No marks yet</b><span class="small">After you mark a script you’ll see how your mark compares with the team.</span></div>`;
  return `<div class="table-wrap"><table><thead><tr><th>Script</th><th class="r">Yours</th><th class="r">Team avg</th><th class="r">Agreed</th><th></th></tr></thead><tbody>${items.map(q=>{const st=qStats(q);return `<tr><td><button class="linkish" data-act="mark" data-id="${q.id}">${esc(q.title)}</button><div class="small muted">${q.level} · ${esc(q.subject)}</div></td><td class="r num">${fmtN(mine[q.id].mark)}/${q.maxMarks}</td><td class="r num">${st.n>1?fmtN(st.mean):'—'}</td><td class="r num">${hasRef(q)?q.referenceMark:'—'}</td><td>${q.status==='closed'?'<span class="chip">Closed</span>':''}</td></tr>`;}).join('')}</tbody></table></div>`;
}
function variationTable(rows){
  if(!rows.length) return `<div class="empty"><b>No comparisons yet</b><span class="small">Variation appears once you and at least one colleague have marked the same script.</span></div>`;
  let body='';
  for(const lv of LEVELS){
    const lr=rows.filter(r=>r.q.level===lv); if(!lr.length) continue;
    const ls=summarise(lr);
    body+=`<tr class="group"><td>${lv} overall</td><td class="r num">${ls.n}</td><td>${devBar(ls.mean)}</td><td class="r num">${fmtN(ls.abs)}%</td><td class="r num">${signed(ls.ref)}</td><td>${verdict(ls.mean)}</td></tr>`;
    const subs=[...new Set(lr.map(r=>r.q.subject))].sort();
    for(const sub of subs){ const ss=summarise(lr.filter(r=>r.q.subject===sub));
      body+=`<tr><td style="padding-left:22px">${esc(sub)}</td><td class="r num">${ss.n}</td><td>${devBar(ss.mean)}</td><td class="r num">${fmtN(ss.abs)}%</td><td class="r num">${signed(ss.ref)}</td><td>${verdict(ss.mean)}</td></tr>`; }
  }
  return `<div class="table-wrap"><table><thead><tr><th>Key stage / subject</th><th class="r">Scripts</th><th>Variation from colleagues</th><th class="r">Avg difference</th><th class="r">vs agreed mark</th><th>Judgement</th></tr></thead><tbody>${body}</tbody></table></div>
  <div class="legend"><span><i style="background:var(--ok)"></i>Within ±${S.config.tolerance}%</span><span><i style="background:var(--lenient)"></i>Lenient (marks higher)</span><span><i style="background:var(--severe)"></i>Severe (marks lower)</span></div>`;
}
function leadTasks(){
  const scopeQ=S.questions.filter(q=>scopeSubjects().includes(q.subject));
  const team=teamStaff();
  const needRef=scopeQ.filter(q=>!hasRef(q) && qStats(q).n>=2);
  const overdueStaff=team.filter(s=>todoFor(s).some(q=>q.deadline&&daysUntil(q.deadline)<0));
  const unlinked=isCC()?S.staff.filter(s=>!s.email && !s.demo):[];
  const wide=scopeQ.map(q=>({q,st:qStats(q)})).filter(x=>x.st.n>=2 && x.st.sdPct>S.config.tolerance);
  const items=[];
  if(needRef.length) items.push(`<b>${needRef.length}</b> marked script${needRef.length>1?'s have':' has'} no agreed mark yet. <a href="#questions">Set agreed marks</a>`);
  if(wide.length) items.push(`<b>${wide.length}</b> script${wide.length>1?'s show':' shows'} a spread wider than tolerance — worth discussing at standardisation. <a href="#team">See team performance</a>`);
  if(overdueStaff.length) items.push(`<b>${overdueStaff.length}</b> colleague${overdueStaff.length>1?'s have':' has'} overdue marking: ${overdueStaff.slice(0,5).map(s=>esc(s.name)).join(', ')}${overdueStaff.length>5?'…':''}`);
  if(unlinked.length) items.push(`<b>${unlinked.length}</b> staff ${unlinked.length>1?'entries have':'entry has'} no email address, so they can’t see their tasks. <a href="#staff">Add emails</a>`);
  return `<section class="card"><div class="card-head"><h2>Moderation tasks for ${isCC()?'the department':'your team'}</h2></div>${items.length?`<div class="todo">${items.map(i=>`<div class="todo-item"><span class="box" aria-hidden="true"></span><div class="body">${i}</div></div>`).join('')}</div>`:`<div class="empty small">No outstanding moderation tasks.</div>`}</section>`;
}
function onboarding(){
  return `<div class="page-head"><div><div class="eyebrow">Curriculum coordinator</div><h1>Set up moderation for your department</h1><p class="sub">Add your team, give heads of subject access, then upload the scripts you want everyone to mark. Variation appears as soon as two people have marked the same script.</p></div></div>
  <div class="steps"><div class="step"><b>1 · Add staff</b><span class="small muted">Name, school email, role, key stages and subjects. Their to-do list appears when they open the app in Teams.</span><a class="btn" href="#staff">Open staff list</a></div>
  <div class="step"><b>2 · Upload scripts</b><span class="small muted">Paste or upload the question, mark scheme and student answer. Classify by key stage and subject.</span><a class="btn" href="#questions">Add questions</a></div>
  <div class="step"><b>3 · Review the team</b><span class="small muted">See who marks leniently or severely, by key stage and subject, against the agreed mark.</span><a class="btn" href="#team">Team performance</a></div></div>
  <div class="card"><div class="card-head"><div><h2>Want to look around first?</h2><p class="small muted">Load a sample team of six teachers with eight marked scripts. Everything is labelled “sample” and can be removed in one click from Staff &amp; settings.</p></div><button class="btn primary" data-act="demo-load">Load sample team</button></div></div>
  <div class="card"><p class="small muted">To mark scripts yourself, add yourself to the staff list with your own email address.</p></div>`;
}

/* ---- Questions ---- */
function viewQuestions(){
  const subs=scopeSubjects();
  let h=`<div class="page-head"><div><div class="eyebrow">${isCC()?'All subjects':'Your subjects: '+esc(subs.join(', '))}</div><h1>Moderation questions</h1><p class="sub">Each entry is one question, its mark scheme and a student’s answer for the team to mark blind.</p></div>
    <div class="btn-row"><button class="btn" data-act="bulk" ${S.canEdit?'':'disabled'}>Bulk import</button><button class="btn primary" data-act="q-new" ${S.canEdit?'':'disabled'}>Add question</button></div></div>`;
  if(!S.canEdit) h+=`<div class="banner">Adding and editing questions needs Editor access to this page. Ask the page owner to make you an Editor.</div>`;
  h+=`<section class="card"><div class="toolbar">
    <input type="search" id="qf-q" placeholder="Search titles and topics" value="${esc(S.qf.q)}" data-input="qf" data-key="q" aria-label="Search">
    <select id="qf-level" data-change="qf" data-key="level" aria-label="Key stage"><option value="">All key stages</option>${LEVELS.map(l=>`<option ${S.qf.level===l?'selected':''}>${l}</option>`).join('')}</select>
    <select id="qf-subject" data-change="qf" data-key="subject" aria-label="Subject"><option value="">All subjects</option>${subs.map(x=>`<option ${S.qf.subject===x?'selected':''}>${esc(x)}</option>`).join('')}</select>
    <select id="qf-status" data-change="qf" data-key="status" aria-label="Status"><option value="">Open and closed</option><option value="open" ${S.qf.status==='open'?'selected':''}>Open</option><option value="closed" ${S.qf.status==='closed'?'selected':''}>Closed</option></select>
    <select id="qf-sort" data-change="qf" data-key="sort" aria-label="Sort"><option value="new" ${S.qf.sort==='new'?'selected':''}>Newest first</option><option value="title" ${S.qf.sort==='title'?'selected':''}>Title A–Z</option><option value="deadline" ${S.qf.sort==='deadline'?'selected':''}>Deadline</option><option value="spread" ${S.qf.sort==='spread'?'selected':''}>Widest spread</option></select>
  </div><div id="qlist">${questionList()}</div></section>`;
  return h;
}
function questionList(){
  const subs=scopeSubjects(), f=S.qf, term=f.q.trim().toLowerCase();
  let qs=S.questions.filter(q=>subs.includes(q.subject) && (!f.level||q.level===f.level) && (!f.subject||q.subject===f.subject) && (!f.status||(q.status||'open')===f.status) && (!term||((q.title||'')+' '+(q.topic||'')).toLowerCase().includes(term)));
  const st=new Map(qs.map(q=>[q.id,qStats(q)]));
  const sorters={new:(a,b)=>(b.createdAt||'').localeCompare(a.createdAt||''), title:(a,b)=>(a.title||'').localeCompare(b.title||''), deadline:(a,b)=>(a.deadline||'9999').localeCompare(b.deadline||'9999'), spread:(a,b)=>st.get(b.id).sdPct-st.get(a.id).sdPct};
  qs.sort(sorters[f.sort]||sorters.new);
  if(!S.questions.length) return `<div class="empty"><b>No questions yet</b><span class="small">Add a question by pasting text or uploading a Word, PDF or text file, or bulk-import a spreadsheet of questions.</span><button class="btn primary" data-act="q-new">Add the first question</button></div>`;
  if(!qs.length) return `<div class="empty small">No questions match these filters.</div>`;
  return `<div class="table-wrap"><table><thead><tr><th>Title</th><th>Class</th><th class="r">Max</th><th>Deadline</th><th class="r">Marked</th><th class="r">Mean</th><th class="r">Spread (SD)</th><th class="r">Agreed</th><th></th></tr></thead><tbody>${qs.map(q=>{const s=st.get(q.id); const tol=S.config.tolerance;
    return `<tr><td><button class="linkish" data-act="q-results" data-id="${q.id}">${esc(q.title)}</button>${q.topic?`<div class="small muted">${esc(q.topic)}</div>`:''}${q.demo?' <span class="chip">sample</span>':''}</td>
    <td><span class="chip ks">${q.level}</span> <span class="chip">${esc(q.subject)}</span></td><td class="r num">${q.maxMarks}</td>
    <td>${q.deadline?fmtDate(q.deadline):'<span class="muted">—</span>'} ${(q.status||'open')==='closed'?'<span class="chip">Closed</span>':''}</td>
    <td class="r num">${s.n}/${s.expected}</td><td class="r num">${s.n?fmtN(s.mean):'—'}</td>
    <td class="r num" style="color:${s.n>=2&&s.sdPct>tol?'var(--danger)':'inherit'}">${s.n>=2?fmtN(s.sdPct)+'%':'—'}</td><td class="r num">${hasRef(q)?q.referenceMark:'—'}</td>
    <td><div class="btn-row" style="justify-content:flex-end;flex-wrap:nowrap"><button class="btn sm" data-act="q-edit" data-id="${q.id}">Edit</button><button class="btn sm" data-act="q-toggle" data-id="${q.id}">${(q.status||'open')==='closed'?'Reopen':'Close'}</button><button class="btn sm danger" data-act="q-del" data-id="${q.id}" data-confirm="Confirm delete">Delete</button></div></td></tr>`;}).join('')}</tbody></table></div>`;
}

/* ---- Team performance ---- */
function viewTeam(){
  const subs=scopeSubjects(), team=teamStaff();
  const subF=S.tf.subject, lvF=S.tf.level;
  const all=devRows().filter(r=>subs.includes(r.q.subject) && (!subF||r.q.subject===subF) && (!lvF||r.q.level===lvF));
  const qsScope=S.questions.filter(q=>subs.includes(q.subject)&&(!subF||q.subject===subF)&&(!lvF||q.level===lvF));
  const tol=S.config.tolerance;
  let expected=0, done=0;
  for(const s of team){ for(const q of qsScope){ if(assignedTo(q,s)){expected++; if(staffMarks(s)[q.id]) done++;} } }
  const perStaff=team.map(s=>({s, sum:summarise(all.filter(r=>r.staffId===s.id))}));
  const outside=perStaff.filter(x=>x.sum.n && Math.abs(x.sum.mean)>tol);
  const spreads=qsScope.map(q=>({q,st:qStats(q)})).filter(x=>x.st.n>=2).sort((a,b)=>b.st.sdPct-a.st.sdPct);
  const teamSpread=mean(spreads.map(x=>x.st.sdPct));
  const levelsShown = lvF?[lvF]:LEVELS;
  let h=`<div class="page-head"><div><div class="eyebrow">${isCC()?'Whole department':'Your team · '+esc(subs.join(', '))}</div><h1>Team performance</h1><p class="sub">How each marker compares with colleagues on the same scripts. Click a name to open their dashboard.</p></div>
  <div class="toolbar"><select id="tf-subject" data-change="tf" data-key="subject" aria-label="Subject"><option value="">All subjects</option>${subs.map(x=>`<option ${subF===x?'selected':''}>${esc(x)}</option>`).join('')}</select>
  <select id="tf-level" data-change="tf" data-key="level" aria-label="Key stage"><option value="">All key stages</option>${LEVELS.map(l=>`<option ${lvF===l?'selected':''}>${l}</option>`).join('')}</select></div></div>`;
  h+=`<div class="tiles"><div class="tile"><span class="v">${expected?Math.round(done/expected*100)+'%':'—'}</span><span class="l">Marking complete (${done} of ${expected})</span></div>
  <div class="tile"><span class="v" style="color:${outside.length?'var(--danger)':'var(--ok)'}">${outside.length}</span><span class="l">Markers outside ±${tol}% tolerance</span></div>
  <div class="tile"><span class="v">${teamSpread==null?'—':fmtN(teamSpread)+'%'}</span><span class="l">Average spread per script (SD)</span></div>
  <div class="tile"><span class="v">${spreads.length}</span><span class="l">Scripts with 2+ markers</span></div></div>`;
  if(!team.length){ h+=`<div class="card"><div class="empty"><b>No team members yet</b><span class="small">${isCC()?'Add staff in Staff &amp; settings.':'Ask your curriculum coordinator to add staff to your subjects.'}</span></div></div>`; return h; }
  h+=`<section class="card"><div class="card-head"><div><h2>Variation by marker and key stage</h2><p class="small muted">Average of (marker’s mark − colleagues’ average) ÷ max marks.</p></div>
    <div class="legend"><span><i style="background:var(--ok)"></i>Within tolerance</span><span><i style="background:var(--lenient)"></i>Lenient</span><span><i style="background:var(--severe)"></i>Severe</span></div></div>
    <div class="table-wrap"><table><thead><tr><th>Marker</th><th>Role</th>${levelsShown.map(l=>`<th style="text-align:center">${l}</th>`).join('')}<th style="text-align:center">Overall</th><th class="r">Avg difference</th><th class="r">vs agreed</th><th class="r">Completed</th></tr></thead><tbody>
    ${perStaff.sort((a,b)=>Math.abs(b.sum.mean||0)-Math.abs(a.sum.mean||0)).map(({s,sum})=>{
      let ex=0,dn=0; for(const q of qsScope){ if(assignedTo(q,s)){ex++; if(staffMarks(s)[q.id]) dn++;} }
      return `<tr><td><button class="linkish" data-act="viewas" data-id="${s.id}">${esc(s.name)}</button>${s.demo?' <span class="chip">sample</span>':''}</td><td class="small muted">${ROLE_LABEL[s.role]||''}</td>
      ${levelsShown.map(l=>`<td style="text-align:center">${(s.levels||[]).includes(l)?heatCell(summarise(all.filter(r=>r.staffId===s.id&&r.q.level===l))):'<span class="muted small">not assigned</span>'}</td>`).join('')}
      <td style="text-align:center">${heatCell(sum)}</td><td class="r num">${sum.n?fmtN(sum.abs)+'%':'—'}</td><td class="r num">${signed(sum.ref)}</td><td class="r num">${ex?`${dn}/${ex}`:'—'}</td></tr>`;}).join('')}
    </tbody></table></div></section>`;
  h+=`<div class="grid-2"><section class="card"><div class="card-head"><h2>By subject</h2></div><div class="table-wrap"><table><thead><tr><th>Subject</th><th class="r">Comparisons</th><th class="r">Avg difference</th><th class="r">Markers outside tolerance</th></tr></thead><tbody>
    ${(subF?[subF]:subs).map(sub=>{const r=all.filter(x=>x.q.subject===sub); const ss=summarise(r); const out=team.filter(s=>{const m=summarise(r.filter(x=>x.staffId===s.id));return m.n&&Math.abs(m.mean)>tol;}).length;
      return `<tr><td>${esc(sub)}</td><td class="r num">${ss.n}</td><td class="r num">${ss.n?fmtN(ss.abs)+'%':'—'}</td><td class="r num" style="color:${out?'var(--danger)':'inherit'}">${ss.n?out:'—'}</td></tr>`;}).join('')}
    </tbody></table></div></section>
    <section class="card"><div class="card-head"><h2>Least consistent scripts</h2><span class="small muted">Spread as SD ÷ max marks</span></div>
    ${spreads.length?`<div class="table-wrap"><table><thead><tr><th>Script</th><th class="r">Markers</th><th class="r">Range</th><th class="r">Spread</th></tr></thead><tbody>${spreads.slice(0,8).map(({q,st})=>`<tr><td><button class="linkish" data-act="q-results" data-id="${q.id}">${esc(q.title)}</button><div class="small muted">${q.level} · ${esc(q.subject)}</div></td><td class="r num">${st.n}</td><td class="r num">${fmtN(st.min)}–${fmtN(st.max)} / ${q.maxMarks}</td><td class="r num" style="color:${st.sdPct>tol?'var(--danger)':'inherit'}">${fmtN(st.sdPct)}%</td></tr>`).join('')}</tbody></table></div>`:`<div class="empty small">Appears once two markers have marked the same script.</div>`}
    </section></div>`;
  return h;
}

/* ---- Staff & settings (CC: everyone; HoS: their subjects + unassigned staff) ---- */
function manageableStaff(){
  if(isCC()) return S.staff;
  const subs=scopeSubjects();
  return S.staff.filter(s=>!(s.subjects||[]).length || (s.subjects||[]).some(x=>subs.includes(x)));
}
const canEditStaffRec = s => isCC() || s.role!=='cc';
const editableSubjects = () => isCC() ? S.config.subjects : scopeSubjects();
function viewStaff(){
  const list=manageableStaff(), dis=S.canEdit?'':'disabled';
  let h=`<div class="page-head"><div><div class="eyebrow">${isCC()?'Curriculum coordinator':'Head of Subject · '+esc(scopeSubjects().join(', '))}</div><h1>Staff &amp; settings</h1><p class="sub">Add, edit or remove staff and assign them to key stages and subjects. Teachers see the scripts that match both.${isCC()?'':' You can manage staff in your subjects and anyone not yet given a subject.'}</p></div><button class="btn primary" data-act="staff-new" ${dis}>Add staff member</button></div>`;
  if(!S.canEdit) h+=`<div class="banner">Changing staff, settings and questions needs Editor access to this page. Ask the page owner to make you an Editor.</div>`;
  h+=`<section class="card"><div class="card-head"><h2>Staff (${list.length})</h2></div>`;
  if(!list.length) h+=`<div class="empty"><b>No staff yet</b><span class="small">Add each teacher with their role, key stages and subjects.</span><button class="btn primary" data-act="staff-new" ${dis}>Add staff member</button></div>`;
  else h+=`<div class="table-wrap"><table><thead><tr><th>Name</th><th>Role</th><th>Key stages</th><th>Subjects</th><th>Email</th><th class="r">To mark</th><th></th></tr></thead><tbody>${list.map(s=>`<tr>
    <td><b>${esc(s.name)}</b>${s.demo?' <span class="chip">sample</span>':''}</td><td>${ROLE_LABEL[s.role]||'Teacher'}</td>
    <td>${(s.levels||[]).map(l=>`<span class="chip ks">${l}</span>`).join(' ')}</td><td class="small">${esc((s.subjects||[]).join(', '))||'<span class="chip warn">No subject</span>'}</td>
    <td class="small">${s.email?esc(s.email):'<span class="chip warn">No email</span>'}</td>
    <td class="r num">${todoFor(s).length}</td>
    <td>${canEditStaffRec(s)?`<div class="btn-row" style="justify-content:flex-end;flex-wrap:nowrap"><button class="btn sm" data-act="staff-edit" data-id="${s.id}" ${dis}>Edit</button><button class="btn sm danger" data-act="staff-del" data-id="${s.id}" data-confirm="Confirm remove" ${dis}>Remove</button></div>`:'<span class="small muted">Curriculum coordinator only</span>'}</td></tr>`).join('')}</tbody></table></div>`;
  h+=`</section><div class="grid-2"><section class="card"><h2>Moderation settings</h2>
    <div class="field"><label for="set-tol">Tolerance (± % of marks available)</label><input id="set-tol" type="number" min="1" max="50" step="0.5" value="${S.config.tolerance}"><span class="hint">Markers whose average variation is inside this band are “within tolerance”. Many departments use 5–10%.</span></div>
    <div class="field"><label for="set-subs">Subjects</label><input id="set-subs" value="${esc(S.config.subjects.join(', '))}"><span class="hint">Separate with commas. Renaming a subject does not update questions already filed under the old name.</span></div>
    <div><button class="btn primary" data-act="save-settings" ${dis}>Save settings</button></div></section>
    <section class="card"><h2>Who can do what</h2>
    <div class="table-wrap"><table><thead><tr><th>Role</th><th>SharePoint access</th><th>Can</th></tr></thead><tbody>
      <tr><td>Curriculum coordinator</td><td>Team <b>owner</b></td><td class="small">Everything: all staff, settings, questions and team results</td></tr>
      <tr><td>Head of Subject</td><td>Team <b>owner</b></td><td class="small">Staff, settings, questions and team results for their subjects</td></tr>
      <tr><td>Teacher</td><td>Team <b>member</b></td><td class="small">Mark scripts and see their own dashboard only</td></tr></tbody></table></div>
    <p class="small muted">The role shown in this app comes from the staff list. What each person can change is enforced by their place in the Science team (owner or member). Keep the two in step.</p>
    ${isCC()?`<h3>Sample data</h3><p class="small muted">A sample team of six teachers with eight marked scripts across KS3–KS5. Sample entries are labelled and can be removed without touching your real data.</p>
    <div class="btn-row"><button class="btn" data-act="demo-load" ${dis}>Load sample team</button><button class="btn danger" data-act="demo-clear" data-confirm="Confirm removal" ${dis}>Remove sample data</button></div>`:''}</section></div>`;
  return h;
}

/* ---------- modals ---------- */
function openModal(html, wide){
  $('#modal-root').innerHTML=`<div class="overlay" data-act="overlay"><div class="modal" role="dialog" aria-modal="true" ${wide?'style="width:min(1040px,100%)"':''}>${html}</div></div>`;
  document.body.style.overflow='hidden';
  const f=$('#modal-root [autofocus]')||$('#modal-root input,#modal-root textarea,#modal-root select'); if(f) f.focus();
}
function closeModal(){ $('#modal-root').innerHTML=''; document.body.style.overflow=''; }
const modalHead = (eyebrow,title)=>`<div class="modal-head"><div><div class="eyebrow">${eyebrow}</div><h2>${title}</h2></div><button class="xbtn" data-act="modal-close" aria-label="Close">×</button></div>`;
const FIELD_KEY={'qe-question':'question','qe-ms':'markScheme','qe-resp':'response','qe-quick':'question'};
const FIELD_NAME={question:'Question',markScheme:'Mark scheme',response:'Student response'};
function docSections(q){
  const att=q.attachments||[];
  const sec=(title,text,field,cls)=>{ const a=att.filter(x=>x.field===field); if(!text&&!a.length) return ''; return `<div class="doc ${cls}"><h3>${title}</h3>${text?`<div class="txt">${esc(text)}</div>`:''}${a.length?attachInner(a):''}</div>`; };
  return sec('Question',q.questionText,'question','')+sec('Student response',q.response,'response','')+sec('Mark scheme',q.markScheme,'markScheme','ms')+attachHtml(att.filter(x=>!x.field));
}
function attachHtml(list){
  if(!list||!list.length) return '';
  return `<div class="doc"><h3>Attachments</h3>${attachInner(list)}</div>`;
}
function attachInner(list){
  return `<div class="attach">${list.map(a=>(a.type||'').startsWith('image/')?`<span style="display:flex;flex-direction:column;gap:4px;max-width:220px;font-size:.8rem"><img data-file="${esc(a.id)}" alt="${esc(a.name)}"><span>${esc(a.name)}</span></span>`:`<a href="${esc(a.webUrl||'#')}" target="_blank" rel="noopener" class="btn sm">Open ${esc(a.name)}</a>`).join('')}</div>`;
}

/* Mark a script: full-screen split viewer */
const VIEW={split:50, z:{left:1,right:1}};
try{ const v=JSON.parse(localStorage.getItem('mm-view')||'{}'); if(v.split) VIEW.split=v.split; if(v.z) VIEW.z={left:+v.z.left||1,right:+v.z.right||1}; }catch(e){}
function saveView(){ try{ localStorage.setItem('mm-view',JSON.stringify(VIEW)); }catch(e){} }
function paneMedia(q, field){
  return (q.attachments||[]).filter(a=>(a.field||'')===field).map(a=>{
    if((a.type||'').startsWith('image/')) return `<img data-file="${esc(a.id)}" alt="${esc(a.name)}" data-fit>`;
    if(/pdf/i.test(a.type||'')||/\.pdf$/i.test(a.name||'')) return `<div class="pdf-pages" data-pdf="${esc(a.id)}" data-name="${esc(a.name)}" data-web="${esc(a.webUrl||'')}"><span class="small muted">Loading ${esc(a.name)}…</span></div>`;
    return `<a class="btn sm" href="${esc(a.webUrl||'#')}" target="_blank" rel="noopener">Open ${esc(a.name)}</a>`;
  }).join('');
}
function paneSec(q,label,text,field){
  const media=paneMedia(q,field); if(!text && !media) return '';
  return `<section class="psec"><div class="sec-label">${label}</div>${text?`<div class="txt">${esc(text)}</div>`:''}${media}</section>`;
}
const zoomBar = pane => `<div class="zoom" role="group" aria-label="Zoom"><button type="button" data-act="zoom" data-pane="${pane}" data-d="-1" aria-label="Zoom out">−</button><button type="button" class="zv" data-act="zoom" data-pane="${pane}" data-d="0" title="Reset to 100%">${Math.round(VIEW.z[pane]*100)}%</button><button type="button" data-act="zoom" data-pane="${pane}" data-d="1" aria-label="Zoom in">+</button></div>`;
function openMark(qid){
  const q=S.questions.find(x=>x.id===qid), s=targetStaff(); if(!q) return;
  const can=canMarkFor(s) && ((q.status||'open')!=='closed' || isCC());
  const existing=s?staffMarks(s)[qid]:null;
  const onBehalf = s && !isMe(s);
  const left = paneSec(q,'Question',q.questionText,'question') + paneSec(q,'Student response',q.response,'response') + paneSec(q,'Other files','','') || '<p class="muted">No question text has been added.</p>';
  const right = paneSec(q,'Mark scheme',q.markScheme,'markScheme') || '<p class="muted">No mark scheme has been added to this question.</p>';
  $('#modal-root').innerHTML=`<div class="viewer" id="viewer" role="dialog" aria-modal="true" aria-label="${esc(q.title)}">
    <div class="v-head"><div style="min-width:0"><div class="eyebrow">${q.level} · ${esc(q.subject)}${q.topic?' · '+esc(q.topic):''}</div><h2>${esc(q.title)} <span class="muted num" style="font-size:.85rem">[${q.maxMarks} marks]</span></h2></div>
      <div class="btn-row"><div class="seg" role="group" aria-label="Pane widths"><button type="button" data-act="split-set" data-v="70">Question wider</button><button type="button" data-act="split-set" data-v="50">Even</button><button type="button" data-act="split-set" data-v="30">Mark scheme wider</button></div><button class="btn" data-act="modal-close">Close</button></div></div>
    <div class="v-split" id="v-split" style="--split:${VIEW.split}%">
      <section class="pane left" aria-label="Question paper"><div class="pane-bar"><h3>Question paper</h3>${zoomBar('left')}</div><div class="pane-scroll" id="ps-left"><div class="pane-content" id="pc-left">${left}</div></div></section>
      <div class="divider" id="v-div" role="separator" tabindex="0" aria-label="Drag to resize. Arrow keys also work; double-click to reset." aria-valuemin="15" aria-valuemax="85" aria-valuenow="${Math.round(VIEW.split)}" title="Drag to resize · double-click to reset"></div>
      <section class="pane right ms" aria-label="Mark scheme"><div class="pane-bar"><h3>Mark scheme</h3>${zoomBar('right')}</div><div class="pane-scroll" id="ps-right"><div class="pane-content" id="pc-right">${right}</div></div></section>
    </div>
    <div class="v-foot">
      ${can?`<form class="mark-box" id="mark-form">
        <div class="field"><label for="mk-val">${onBehalf?'Mark for '+esc(s.name):'Your mark'}</label><div style="display:flex;align-items:baseline;gap:6px"><input id="mk-val" type="number" min="0" max="${q.maxMarks}" step="0.5" required value="${existing?existing.mark:''}"><span class="num muted">/ ${q.maxMarks}</span></div></div>
        <div class="field" style="flex:1;min-width:200px"><label for="mk-note">Rationale (optional)</label><input id="mk-note" maxlength="400" placeholder="e.g. Credited M1 for correct formula; no A1 as unit missing" value="${esc(existing?.note||'')}"></div>
        <div class="btn-row"><button class="btn primary" type="submit">${existing?'Update mark':'Submit mark'}</button>${existing?`<button class="btn danger" type="button" data-act="mark-del" data-id="${q.id}" data-confirm="Confirm withdraw">Withdraw</button>`:''}</div>
      </form>${existing?'':'<p class="small muted">Marking is blind: you’ll see how colleagues marked this script after you submit.</p>'}`
      : `<div class="banner info">${(q.status||'open')==='closed'?'This script is closed for marking.':'You can view this script but not mark it here.'}</div>`}
      <div id="mark-result">${existing?markResult(q,s):''}</div>
    </div></div>`;
  document.body.style.overflow='hidden';
  for(const p of ['left','right']) setZoom(p, VIEW.z[p], null, true);
  document.querySelectorAll('#viewer img[data-fit]').forEach(img=>{ const fit=()=>{ img.style.width=(img.naturalWidth/15)+'em'; }; img.complete&&img.naturalWidth?fit():img.addEventListener('load',fit); });
  wireViewer();
  hydrateFiles($('#viewer'));
  renderPdfs($('#viewer'));
  const form=$('#mark-form');
  if(form) form.addEventListener('submit', async e=>{
    e.preventDefault(); const v=parseFloat($('#mk-val').value);
    if(!isFinite(v)||v<0||v>q.maxMarks){ toast(`Enter a mark between 0 and ${q.maxMarks}.`,true); return; }
    const btn=form.querySelector('[type=submit]'); btn.disabled=true;
    try{ await saveMark(s,q.id,v,$('#mk-note').value.trim()); await refresh(); toast('Mark saved'); const r=$('#mark-result'); if(r) r.innerHTML=markResult(q,s); btn.textContent='Update mark'; }
    catch(err){ toast(dbErr(err),true); }
    btn.disabled=false;
  });
}
/* Zoom scales the pane's content width and type together, so text, screenshots and PDF pages all magnify and the pane scrolls to pan. */
function setZoom(pane, z, anchor, initial){
  z=Math.max(0.5,Math.min(4,Math.round(z*100)/100));
  const sc=$('#ps-'+pane), pc=$('#pc-'+pane); if(!sc||!pc) return;
  const k=z/(initial?z:VIEW.z[pane]);
  const cx=anchor?anchor[0]:0, cy=anchor?anchor[1]:sc.clientHeight/3;
  const x=sc.scrollLeft+cx, y=sc.scrollTop+cy;
  VIEW.z[pane]=z; pc.style.width=(z*100)+'%'; pc.style.fontSize=(z*15)+'px';
  if(!initial){ sc.scrollLeft=x*k-cx; sc.scrollTop=y*k-cy; saveView(); }
  const lab=document.querySelector(`#viewer [data-pane="${pane}"][data-d="0"]`); if(lab) lab.textContent=Math.round(z*100)+'%';
}
function setSplit(p){
  VIEW.split=Math.max(15,Math.min(85,p));
  const el=$('#v-split'); if(el) el.style.setProperty('--split',VIEW.split+'%');
  const d=$('#v-div'); if(d) d.setAttribute('aria-valuenow',Math.round(VIEW.split));
}
function wireViewer(){
  const div=$('#v-div'), split=$('#v-split'), viewer=$('#viewer');
  const stacked=()=>matchMedia('(max-width:760px)').matches;
  div.addEventListener('pointerdown',e=>{ div.setPointerCapture(e.pointerId); div.classList.add('drag'); viewer.classList.add('dragging'); e.preventDefault(); });
  div.addEventListener('pointermove',e=>{ if(!div.hasPointerCapture(e.pointerId)) return; const r=split.getBoundingClientRect(); setSplit(stacked()?(e.clientY-r.top)/r.height*100:(e.clientX-r.left)/r.width*100); });
  const end=()=>{ div.classList.remove('drag'); viewer.classList.remove('dragging'); saveView(); };
  div.addEventListener('pointerup',end); div.addEventListener('pointercancel',end);
  div.addEventListener('dblclick',()=>{ setSplit(50); saveView(); });
  div.addEventListener('keydown',e=>{ const st=e.shiftKey?10:3;
    if(e.key==='ArrowLeft'||e.key==='ArrowUp'){ setSplit(VIEW.split-st); saveView(); e.preventDefault(); }
    else if(e.key==='ArrowRight'||e.key==='ArrowDown'){ setSplit(VIEW.split+st); saveView(); e.preventDefault(); }
    else if(e.key==='Home'){ setSplit(50); saveView(); e.preventDefault(); } });
  for(const pane of ['left','right']){
    const sc=$('#ps-'+pane), pc=$('#pc-'+pane);
    /* text re-wraps to the visible pane width at any zoom; images and PDF pages magnify beyond it */
    const fitText=()=>pc.style.setProperty('--pw',sc.clientWidth+'px'); fitText();
    if(window.ResizeObserver) new ResizeObserver(fitText).observe(sc);
    sc.addEventListener('wheel',e=>{ if(!e.ctrlKey&&!e.metaKey) return; e.preventDefault(); const r=sc.getBoundingClientRect(); setZoom(pane, VIEW.z[pane]*(e.deltaY<0?1.1:1/1.1), [e.clientX-r.left,e.clientY-r.top]); },{passive:false});
    let pinch=null;
    const dist=t=>Math.hypot(t[0].clientX-t[1].clientX,t[0].clientY-t[1].clientY);
    sc.addEventListener('touchstart',e=>{ if(e.touches.length===2){ pinch={d:dist(e.touches),z:VIEW.z[pane]}; } },{passive:true});
    sc.addEventListener('touchmove',e=>{ if(pinch&&e.touches.length===2){ e.preventDefault(); const r=sc.getBoundingClientRect(); const mx=(e.touches[0].clientX+e.touches[1].clientX)/2-r.left, my=(e.touches[0].clientY+e.touches[1].clientY)/2-r.top; setZoom(pane, pinch.z*dist(e.touches)/pinch.d, [mx,my]); } },{passive:false});
    sc.addEventListener('touchend',e=>{ if(e.touches.length<2) pinch=null; });
  }
}
async function renderPdfs(root){
  const boxes=[...root.querySelectorAll('[data-pdf]')]; if(!boxes.length) return;
  const fallback=b=>{ b.innerHTML=b.dataset.web?`<a class="btn sm" href="${esc(b.dataset.web)}" target="_blank" rel="noopener">Open ${esc(b.dataset.name||'PDF')}</a>`:'<span class="small muted">This PDF could not be shown.</span>'; };
  try{ await loadScript(LIB.pdfWorker); await loadScript(LIB.pdf); }catch(e){ boxes.forEach(fallback); return; }
  const lib=window.pdfjsLib; lib.GlobalWorkerOptions.workerSrc=LIB.pdfWorker;
  for(const b of boxes){
    try{
      const buf=await Store.fileBuffer(b.dataset.pdf);
      const pdf=await lib.getDocument({data:new Uint8Array(buf)}).promise; b.innerHTML='';
      for(let i=1;i<=pdf.numPages;i++){
        if(!b.isConnected) return;
        const page=await pdf.getPage(i), base=page.getViewport({scale:1});
        const css=Math.min(1100, base.width*1.5), vp=page.getViewport({scale:(css/base.width)*2});
        const c=document.createElement('canvas'); c.width=Math.round(vp.width); c.height=Math.round(vp.height); c.style.width=(css/15)+'em'; c.setAttribute('role','img'); c.setAttribute('aria-label',(b.dataset.name||'PDF')+' page '+i);
        await page.render({canvasContext:c.getContext('2d'),viewport:vp}).promise; b.appendChild(c);
      }
    }catch(e){ fallback(b); }
  }
}
function markResult(q,s){
  const qm=questionMarks(q), mine=qm.find(x=>x.s.id===s.id); if(!mine) return '';
  const others=qm.filter(x=>x.s.id!==s.id), max=+q.maxMarks||1;
  const pos=v=>(v/max*100);
  const counts={}; let pts='';
  for(const x of others){ const k=x.mark; counts[k]=(counts[k]||0)+1; pts+=`<span class="pt" style="left:${pos(k)}%;top:${26-(counts[k]-1)*13}px" title="Colleague: ${k}"></span>`; }
  pts+=`<span class="pt me" style="left:${pos(mine.mark)}%;top:24px" title="You: ${mine.mark}"></span>`;
  const ticks=[]; const stepT=max<=10?1:max<=20?2:5; for(let i=0;i<=max;i+=stepT) ticks.push(i);
  const othersMean=mean(others.map(x=>x.mark));
  const dev=others.length?(mine.mark-othersMean)/max*100:null;
  return `<div class="doc"><h3>How this script was marked</h3><div style="padding:12px 14px;display:flex;flex-direction:column;gap:10px">
    <div class="strip">${ticks.map(t=>`<span class="tick" style="left:${pos(t)}%"></span><span class="tl" style="left:${pos(t)}%">${t}</span>`).join('')}<div class="axis"></div>${hasRef(q)?`<span class="agreed" style="left:${pos(+q.referenceMark)}%"><span>Agreed ${q.referenceMark}</span></span>`:''}${pts}</div>
    <div class="btn-row small" style="gap:16px"><span><b class="num">${fmtN(mine.mark)}</b> ${isMe(s)?'your mark':'their mark'}</span><span><b class="num">${others.length?fmtN(othersMean):'—'}</b> colleagues’ average (${others.length} marker${others.length===1?'':'s'})</span>${hasRef(q)?`<span><b class="num">${q.referenceMark}</b> agreed mark</span>`:''}<span>Variation <b class="num">${signed(dev)}</b></span>${verdict(dev)}</div>
    ${!others.length?'<p class="small muted">No colleague has marked this yet. Check back later.</p>':''}
  </div></div>`;
}

/* Question results (leads) */
function openResults(qid){
  const q=S.questions.find(x=>x.id===qid); if(!q) return; const st=qStats(q), max=+q.maxMarks||1;
  const total=st.marks.reduce((a,x)=>a+x.mark,0);
  const missing=expectedMarkers(q).filter(s=>!staffMarks(s)[q.id]);
  openModal(`${modalHead(`${q.level} · ${esc(q.subject)}${q.topic?' · '+esc(q.topic):''}`, esc(q.title))}
  <div class="modal-body">
    <div class="tiles"><div class="tile"><span class="v">${st.n}/${st.expected}</span><span class="l">Marked</span></div><div class="tile"><span class="v">${st.n?fmtN(st.mean):'—'}<span class="muted" style="font-size:1rem"> / ${q.maxMarks}</span></span><span class="l">Mean mark</span></div><div class="tile"><span class="v">${st.n?fmtN(st.min)+'–'+fmtN(st.max):'—'}</span><span class="l">Range</span></div><div class="tile"><span class="v">${st.n>=2?fmtN(st.sdPct)+'%':'—'}</span><span class="l">Spread (SD ÷ max)</span></div></div>
    <div class="table-wrap"><table><thead><tr><th>Marker</th><th class="r">Mark</th><th>vs colleagues</th><th class="r">vs agreed</th><th>Rationale</th></tr></thead><tbody>
    ${st.marks.length?st.marks.sort((a,b)=>b.mark-a.mark).map(x=>{const o=st.n>1?(total-x.mark)/(st.n-1):null; const d=o==null?null:(x.mark-o)/max*100; return `<tr><td>${esc(x.s.name)}</td><td class="r num">${fmtN(x.mark)}</td><td>${devBar(d)}</td><td class="r num">${hasRef(q)?signed((x.mark-q.referenceMark)/max*100):'—'}</td><td class="small">${esc(x.note)}</td></tr>`;}).join(''):`<tr><td colspan="5" class="muted">No marks yet.</td></tr>`}
    </tbody></table></div>
    ${missing.length?`<p class="small"><b>Still to mark:</b> ${missing.map(s=>esc(s.name)).join(', ')}</p>`:''}
    <form class="mark-box" id="ref-form"><div class="field"><label for="ref-val">Agreed (definitive) mark</label><div style="display:flex;align-items:baseline;gap:6px"><input id="ref-val" type="number" min="0" max="${q.maxMarks}" step="0.5" value="${hasRef(q)?q.referenceMark:''}"><span class="num muted">/ ${q.maxMarks}</span></div></div>
      <p class="small muted" style="flex:1;min-width:200px">Set after standardisation. Teachers see it only once they have submitted their own mark.</p><button class="btn primary" type="submit">Save agreed mark</button></form>
    <details><summary class="small" style="cursor:pointer">Show question, response and mark scheme</summary><div style="display:flex;flex-direction:column;gap:12px;margin-top:12px">
      ${docSections(q)}</div></details>
  </div><div class="modal-foot"><button class="btn" data-act="mark" data-id="${q.id}">Open side-by-side view</button><button class="btn" data-act="q-edit" data-id="${q.id}">Edit question</button><button class="btn" data-act="modal-close">Close</button></div>`, true);
  hydrateFiles($('#modal-root'));
  $('#ref-form').addEventListener('submit', async e=>{
    e.preventDefault(); const raw=$('#ref-val').value.trim(); const v=raw===''?null:parseFloat(raw);
    if(v!=null && (!isFinite(v)||v<0||v>q.maxMarks)){ toast(`Enter a mark between 0 and ${q.maxMarks}, or leave it blank.`,true); return; }
    try{ await dbUpdate('questions',q.id,{referenceMark:v}); await refresh(); toast(v==null?'Agreed mark cleared':'Agreed mark saved'); }catch(err){ toast(dbErr(err),true); }
  });
}

/* Question editor */
let draftAttach=[];
function openQuestionEditor(qid){
  const q=qid?S.questions.find(x=>x.id===qid):null, subs=scopeSubjects();
  draftAttach=(q?.attachments||[]).slice();
  const ta=(id,label,val,ph)=>`<div class="field span-4"><div class="field-tools"><label for="${id}">${label}</label><span class="btn-row"><span class="drop-hint">Paste with <kbd>Ctrl</kbd>+<kbd>V</kbd>, drop a file, or</span><label class="btn sm" style="cursor:pointer">Upload file<input type="file" hidden data-import="${id}" accept=".txt,.md,.csv,.docx,.pdf,image/*"></label></span></div><textarea id="${id}" data-drop="${id}" placeholder="${ph}">${esc(val||'')}</textarea><div class="thumbs" data-thumbs="${FIELD_KEY[id]}"></div></div>`;
  openModal(`${modalHead(q?'Edit question':'New moderation question', q?esc(q.title):'Add a script for the team to mark')}
  <form id="q-form"><div class="modal-body">
    <div class="paste-zone">
      <label for="qe-quick"><b>Quick paste.</b> Copy the question and mark scheme from Word, a PDF or an exam board document, click in this box and press <kbd>Ctrl</kbd>+<kbd>V</kbd> (<kbd>⌘</kbd>+<kbd>V</kbd> on a Mac).</label>
      <textarea id="qe-quick" data-paste="quick" placeholder="Paste here. Put a line saying “Mark scheme” before the mark scheme, and “Student response” before the student’s answer, and each part goes into the right box."></textarea>
      <div class="btn-row"><button type="button" class="btn primary sm" data-act="quick-split">Sort into fields</button><span class="small muted" id="qe-quick-msg">Tables, superscripts and subscripts (cm³, CO₂) are kept. Pasted screenshots appear in the question for staff to view.</span></div>
    </div>
    <div class="form-grid">
      <div class="field span-2"><label for="qe-title">Title</label><input id="qe-title" required maxlength="140" value="${esc(q?.title||'')}" placeholder="e.g. Y11 Mock P1 Q04 – Osmosis 6-marker" autofocus></div>
      <div class="field"><label for="qe-level">Key stage</label><select id="qe-level" required>${LEVELS.map(l=>`<option ${q?.level===l?'selected':''}>${l}</option>`).join('')}</select></div>
      <div class="field"><label for="qe-subject">Subject</label><select id="qe-subject" required>${subs.map(x=>`<option ${q?.subject===x?'selected':''}>${esc(x)}</option>`).join('')}</select></div>
      <div class="field span-2"><label for="qe-topic">Topic or paper <span class="muted">(optional)</span></label><input id="qe-topic" maxlength="120" value="${esc(q?.topic||'')}" placeholder="e.g. Paper 1 · Cell biology"></div>
      <div class="field"><label for="qe-max">Max marks</label><input id="qe-max" type="number" min="1" max="200" step="1" required value="${q?.maxMarks||''}"></div>
      <div class="field"><label for="qe-deadline">Marking deadline</label><input id="qe-deadline" type="date" value="${esc(q?.deadline||'')}"></div>
      ${ta('qe-question','Question',q?.questionText,'Paste the question text here')}
      ${ta('qe-ms','Mark scheme',q?.markScheme,'Paste the mark scheme, including indicative content and level descriptors')}
      ${ta('qe-resp','Student response',q?.response,'Paste or type the student’s answer. Scanned scripts can be uploaded as images below.')}
      <div class="field span-2"><label for="qe-ref">Agreed mark <span class="muted">(optional)</span></label><input id="qe-ref" type="number" min="0" step="0.5" value="${q&&hasRef(q)?q.referenceMark:''}"><span class="hint">Hidden from teachers until they submit their own mark.</span></div>
      <div class="field span-2"><span class="lbl">Other files</span>${S.assets?`<label class="btn sm" style="cursor:pointer;align-self:flex-start">Add scans, images or PDF<input type="file" hidden id="qe-attach" multiple accept="image/*,.pdf"></label>`:`<span class="hint">Attachments are not available with your access level. Paste the text instead.</span>`}<div class="attach-list" id="qe-attach-list"></div></div>
    </div>
    <p class="small muted">Pasted screenshots, images and PDFs are shown in full beside the text when staff mark. Word (.docx) and text files are converted to text in the box.</p>
  </div><div class="modal-foot"><button class="btn" type="button" data-act="modal-close">Cancel</button><button class="btn primary" type="submit">${q?'Save changes':'Add question'}</button></div></form>`, true);
  renderAttachList();
  $('#q-form').addEventListener('submit', async e=>{
    e.preventDefault();
    const max=parseInt($('#qe-max').value,10), refRaw=$('#qe-ref').value.trim(), ref=refRaw===''?null:parseFloat(refRaw);
    if(!(max>0)){ toast('Enter the maximum marks for this question.',true); return; }
    if(ref!=null && (ref<0||ref>max)){ toast('The agreed mark must be between 0 and the maximum.',true); return; }
    const data={title:$('#qe-title').value.trim(), level:$('#qe-level').value, subject:$('#qe-subject').value, topic:$('#qe-topic').value.trim(), maxMarks:max, deadline:$('#qe-deadline').value||'', questionText:$('#qe-question').value, markScheme:$('#qe-ms').value, response:$('#qe-resp').value, referenceMark:ref, attachments:draftAttach};
    if(!data.title){ toast('Give the question a title so it can be found later.',true); return; }
    try{
      if(q){ await dbUpdate('questions',q.id,data); toast('Question updated'); }
      else { await dbCreate('questions',{...data, status:'open', createdAt:new Date().toISOString(), createdBy:S.me.email}); toast('Question added. It now appears on matching teachers’ to-do lists.'); }
      closeModal(); refresh();
    }catch(err){ toast(dbErr(err),true); }
  });
}
function renderAttachList(){
  const thumb=(a,i)=>`<div class="thumb">${(a.type||'').startsWith('image/')?`<img data-file="${esc(a.id)}" alt="">`:`<div class="pdf">${/pdf/i.test(a.type||a.name)?'PDF':'File'}</div>`}<span title="${esc(a.name)}">${esc(a.name)}</span><button type="button" data-act="att-del" data-i="${i}" aria-label="Remove ${esc(a.name)}">×</button></div>`;
  document.querySelectorAll('[data-thumbs]').forEach(box=>{ box.innerHTML=draftAttach.map((a,i)=>a.field===box.dataset.thumbs?thumb(a,i):'').join(''); });
  const el=$('#qe-attach-list'); if(el) el.innerHTML=`<div class="thumbs">${draftAttach.map((a,i)=>!a.field?thumb(a,i):'').join('')}</div>`;
  hydrateFiles($('#modal-root'));
}
async function uploadAttachment(file, fieldId){
  if(!S.assets) return false;
  try{
    toast('Uploading '+file.name+'…');
    const r=await Store.upload(file);
    const type=(r.file&&r.file.mimeType)||file.type||(/\.pdf$/i.test(file.name)?'application/pdf':'');
    if(type.startsWith('image/')) Store.urls[r.id]=Promise.resolve(URL.createObjectURL(file));
    draftAttach.push({id:r.id,name:file.name,type,webUrl:r.webUrl||'',field:FIELD_KEY[fieldId]||''}); renderAttachList(); return true;
  }catch(e){ toast('Could not upload '+file.name+'. '+dbErr(e),true); return false; }
}
async function fileToText(file){
  const n=file.name.toLowerCase();
  if(n.endsWith('.docx')){ await loadScript(LIB.mammoth); const r=await window.mammoth.extractRawText({arrayBuffer:await file.arrayBuffer()}); return r.value.trim(); }
  if(n.endsWith('.pdf')){
    await loadScript(LIB.pdfWorker); await loadScript(LIB.pdf);
    const lib=window.pdfjsLib; lib.GlobalWorkerOptions.workerSrc=LIB.pdfWorker;
    const pdf=await lib.getDocument({data:new Uint8Array(await file.arrayBuffer())}).promise; const out=[];
    for(let i=1;i<=pdf.numPages;i++){ const c=await (await pdf.getPage(i)).getTextContent(); out.push(c.items.map(it=>it.str+(it.hasEOL?'\n':' ')).join('')); }
    return out.join('\n\n').replace(/[ \t]+\n/g,'\n').trim();
  }
  if(n.endsWith('.doc')) throw new Error('Older .doc files can’t be read here. Save as .docx or copy and paste the text.');
  if(file.type.startsWith('text/')||/\.(txt|md|csv|tsv)$/.test(n)) return (await file.text()).trim();
  return null;
}
async function importInto(targetId, file){
  const ta=document.getElementById(targetId); if(!ta) return;
  const isImg=file.type.startsWith('image/'), isPdf=file.name.toLowerCase().endsWith('.pdf');
  try{
    if(isImg){ if(await uploadAttachment(file,targetId)) toast(file.name+' added. Staff will see it in full when marking.'); else if(!S.assets) toast('Images need attachment access. Type or paste the text instead.',true); return; }
    if(isPdf && S.assets){ if(await uploadAttachment(file,targetId)) toast(file.name+' added. Staff will see its pages in full when marking.'); return; }
    toast('Reading '+file.name+'…');
    const text=await fileToText(file);
    if(text==null){ toast('That file type can’t be read. Use .docx, .pdf, .txt or paste the text.',true); return; }
    ta.value = ta.value ? ta.value.replace(/\s+$/,'')+'\n\n'+text : text;
    toast(text?'Imported text from '+file.name:'No text found in '+file.name+' (it may be a scan). Attach it instead.', !text);
  }catch(e){ toast((e&&e.message)||'Could not read that file.',true); }
}

/* Clipboard paste: browsers never let a page read the clipboard by itself, so this works on the viewer's own Ctrl/⌘+V. */
const SUP={'0':'⁰','1':'¹','2':'²','3':'³','4':'⁴','5':'⁵','6':'⁶','7':'⁷','8':'⁸','9':'⁹','+':'⁺','-':'⁻','−':'⁻'};
const SUB={'0':'₀','1':'₁','2':'₂','3':'₃','4':'₄','5':'₅','6':'₆','7':'₇','8':'₈','9':'₉','+':'₊','-':'₋'};
function htmlToText(html){
  const doc=new DOMParser().parseFromString(html,'text/html');
  doc.querySelectorAll('style,script,meta,link,title,xml').forEach(n=>n.remove());
  const BLOCK=/^(P|DIV|H[1-6]|LI|TR|BLOCKQUOTE|PRE|SECTION|ARTICLE|UL|OL|TABLE)$/;
  let out='';
  const walk=n=>{
    if(n.nodeType===3){ out+=n.nodeValue.replace(/\s+/g,' '); return; }
    if(n.nodeType!==1) return;
    const t=n.tagName;
    if(t==='BR'){ out+='\n'; return; }
    if(t==='IMG') return;
    if(t==='SUP'||t==='SUB'){ const map=t==='SUP'?SUP:SUB, txt=n.textContent.trim(); out+=[...txt].every(c=>map[c])?[...txt].map(c=>map[c]).join(''):(t==='SUP'?'^':'_')+txt; return; }
    if(t==='TD'||t==='TH'){ for(const c of n.childNodes) walk(c); out+=' | '; return; }
    if(t==='LI') out+='\n• '; else if(BLOCK.test(t)) out+='\n';
    for(const c of n.childNodes) walk(c);
    if(t==='TR') out=out.replace(/ \| $/,'');
    if(BLOCK.test(t)) out+='\n';
  };
  walk(doc.body);
  return out.split('\n').map(l=>l.replace(/[ \t ]+/g,' ').trim()).join('\n').replace(/\n{3,}/g,'\n\n').trim();
}
document.addEventListener('paste', async e=>{
  const ta=e.target.closest && e.target.closest('textarea[data-drop],textarea[data-paste]'); if(!ta) return;
  const cd=e.clipboardData; if(!cd) return;
  const text=cd.getData('text/plain')||'', html=cd.getData('text/html')||'';
  const imgs=[...(cd.files||[])].filter(f=>f.type.startsWith('image/'));
  const fieldId=ta.id;
  if(imgs.length && !text.trim()){
    e.preventDefault();
    if(!S.assets){ toast('Pasted images can’t be saved with your access level. Paste the text instead.',true); return; }
    const label=FIELD_NAME[FIELD_KEY[fieldId]]||'Question';
    for(const [i,f] of imgs.entries()){
      const stamp=new Date().toLocaleTimeString('en-GB',{hour:'2-digit',minute:'2-digit'});
      const named=new File([f],`Pasted ${label.toLowerCase()} ${stamp}${imgs.length>1?' ('+(i+1)+')':''}.png`,{type:f.type});
      if(await uploadAttachment(named,fieldId)) toast('Screenshot added to '+label.toLowerCase());
    }
    return;
  }
  if(html && /<(table|sup|sub)\b/i.test(html)){
    const t=htmlToText(html);
    if(t){ e.preventDefault(); ta.setRangeText(t,ta.selectionStart,ta.selectionEnd,'end'); }
  }
});
function quickSplit(){
  const raw=$('#qe-quick').value;
  if(!raw.trim()){ toast('Paste the question and mark scheme into the quick paste box first.',true); return; }
  const HEADS=[
    {id:'qe-ms', re:/^\s*(mark(ing)?\s*scheme|marking\s*guidance|indicative\s*content|ms)\s*[:\-–]?\s*(.*)$/i},
    {id:'qe-resp', re:/^\s*((student|candidate|pupil)('?s)?\s*(response|answer|work)|response|exemplar\s*(response|answer))\s*[:\-–]?\s*(.*)$/i},
    {id:'qe-question', re:/^\s*(question)\s*[:\-–]?\s*(.*)$/i},
  ];
  const buf={'qe-question':[],'qe-ms':[],'qe-resp':[]}; let cur='qe-question', found=0;
  for(const line of raw.split('\n')){
    let hit=null;
    if(!line.includes('|') && line.trim().length<=80){ for(const h of HEADS){ const m=line.match(h.re); if(m){ const rest=m[m.length-1]; if(line.trim().length<=40 || /[:\-–]/.test(line)){ hit={id:h.id,rest}; break; } } } }
    if(hit){ cur=hit.id; found++; if(hit.rest&&hit.rest.trim()) buf[cur].push(hit.rest.trim()); continue; }
    buf[cur].push(line);
  }
  const filled=[];
  for(const [id,lines] of Object.entries(buf)){ const t=lines.join('\n').replace(/^\n+|\n+$/g,'').trim(); if(t){ document.getElementById(id).value=t; filled.push({'qe-question':'Question','qe-ms':'Mark scheme','qe-resp':'Student response'}[id]); } }
  const qText=buf['qe-question'].join('\n');
  const mm=[...qText.matchAll(/[\[(]\s*(\d{1,3})\s*marks?\s*[\])]/gi)].pop() || [...qText.matchAll(/total\s*(?:of\s*)?(\d{1,3})\s*marks?/gi)].pop();
  if(mm && !$('#qe-max').value) $('#qe-max').value=mm[1];
  const first=qText.split('\n').map(l=>l.trim()).find(Boolean);
  if(first && !$('#qe-title').value.trim()) $('#qe-title').value=first.slice(0,80);
  $('#qe-quick-msg').textContent = found
    ? `Filled ${filled.join(', ')}${mm?' and max marks':''}. Check the boxes below, then add the key stage and subject.`
    : 'No “Mark scheme” heading found, so everything went into Question. Add a line saying Mark scheme before it and sort again, or paste the mark scheme into its own box.';
  $('#qe-question').scrollIntoView({behavior:'smooth',block:'center'});
}

/* Bulk import */
const BULK_COLS={title:['title','name'],level:['level','key stage','keystage','ks'],subject:['subject'],topic:['topic','paper','category'],maxMarks:['max','max marks','marks','total'],questionText:['question','question text'],markScheme:['mark scheme','markscheme','ms'],response:['response','student response','answer','script'],referenceMark:['agreed','agreed mark','reference','reference mark'],deadline:['deadline','due','due date']};
function parseDelimited(text){
  const delim = (text.split('\n')[0].includes('\t')) ? '\t' : ',';
  const rows=[]; let row=[], cell='', q=false;
  for(let i=0;i<text.length;i++){ const c=text[i];
    if(q){ if(c==='"'){ if(text[i+1]==='"'){cell+='"';i++;} else q=false; } else cell+=c; }
    else if(c==='"') q=true; else if(c===delim){row.push(cell);cell='';} else if(c==='\n'||c==='\r'){ if(c==='\r'&&text[i+1]==='\n') i++; row.push(cell); rows.push(row); row=[]; cell=''; } else cell+=c; }
  if(cell!==''||row.length){row.push(cell);rows.push(row);}
  return rows.filter(r=>r.some(x=>x.trim()!==''));
}
function normLevel(v){ const m=String(v||'').toUpperCase().replace(/\s/g,'').match(/(\d)/); return m?('KS'+m[1]):''; }
function parseBulk(text){
  const rows=parseDelimited(text); if(rows.length<2) return {items:[],error:'Include a header row and at least one question row.'};
  const head=rows[0].map(h=>h.trim().toLowerCase()); const idx={};
  for(const [k,al] of Object.entries(BULK_COLS)){ const i=head.findIndex(h=>al.includes(h)); if(i>=0) idx[k]=i; }
  const missing=['title','level','subject','maxMarks'].filter(k=>idx[k]==null);
  if(missing.length) return {items:[],error:'Missing column(s): '+missing.map(k=>BULK_COLS[k][0]).join(', ')+'.'};
  const subs=scopeSubjects();
  const items=rows.slice(1).map(r=>{ const g=k=>idx[k]!=null?(r[idx[k]]||'').trim():'';
    const it={title:g('title'), level:normLevel(g('level')), subject:subs.find(s=>s.toLowerCase()===g('subject').toLowerCase())||g('subject'), topic:g('topic'), maxMarks:parseInt(g('maxMarks'),10), questionText:g('questionText'), markScheme:g('markScheme'), response:g('response'), referenceMark:g('referenceMark')===''?null:parseFloat(g('referenceMark')), deadline:/^\d{4}-\d{2}-\d{2}$/.test(g('deadline'))?g('deadline'):''};
    const errs=[]; if(!it.title) errs.push('no title'); if(!LEVELS.includes(it.level)) errs.push('key stage'); if(!subs.includes(it.subject)) errs.push('subject not in '+(isCC()?'list':'your subjects')); if(!(it.maxMarks>0)) errs.push('max marks');
    it.err=errs.join(', '); return it; });
  return {items};
}
let bulkItems=[];
function openBulk(){
  bulkItems=[];
  openModal(`${modalHead('Bulk import','Import several questions at once')}
  <div class="modal-body">
    <p class="small">Copy rows from Excel or Google Sheets and paste below, or upload a .csv file. The first row must be headers: <span class="mono small">title, level, subject, max marks</span> are required; <span class="mono small">topic, question, mark scheme, response, agreed mark, deadline (YYYY-MM-DD)</span> are optional.</p>
    <div class="field"><div class="field-tools"><label for="bulk-text">Rows</label><label class="btn sm" style="cursor:pointer">Upload .csv<input type="file" hidden id="bulk-file" accept=".csv,.tsv,.txt"></label></div>
    <textarea id="bulk-text" style="min-height:150px;font-family:var(--f-mono);font-size:.82rem" placeholder="title&#9;level&#9;subject&#9;max marks&#9;mark scheme&#10;Respiration 4-marker&#9;KS4&#9;Biology&#9;4&#9;..."></textarea></div>
    <div id="bulk-preview"></div>
  </div><div class="modal-foot"><button class="btn" data-act="modal-close">Cancel</button><button class="btn primary" data-act="bulk-go" id="bulk-go" disabled>Import</button></div>`, true);
  const upd=()=>{ const t=$('#bulk-text').value; if(!t.trim()){ $('#bulk-preview').innerHTML=''; $('#bulk-go').disabled=true; return; }
    const r=parseBulk(t); bulkItems=r.items.filter(i=>!i.err);
    $('#bulk-preview').innerHTML = r.error?`<div class="banner">${esc(r.error)}</div>`:`<div class="table-wrap"><table><thead><tr><th>Title</th><th>KS</th><th>Subject</th><th class="r">Max</th><th>Check</th></tr></thead><tbody>${r.items.map(i=>`<tr><td>${esc(i.title)}</td><td>${esc(i.level)}</td><td>${esc(i.subject)}</td><td class="r num">${isFinite(i.maxMarks)?i.maxMarks:''}</td><td>${i.err?`<span class="chip danger">Skip: ${esc(i.err)}</span>`:'<span class="chip ok">Ready</span>'}</td></tr>`).join('')}</tbody></table></div>`;
    $('#bulk-go').disabled=!bulkItems.length; $('#bulk-go').textContent=bulkItems.length?`Import ${bulkItems.length} question${bulkItems.length>1?'s':''}`:'Import'; };
  $('#bulk-text').addEventListener('input',upd);
  $('#bulk-file').addEventListener('change',async e=>{ const f=e.target.files[0]; if(f){ $('#bulk-text').value=await f.text(); upd(); } });
}
async function runBulk(){
  const btn=$('#bulk-go'); btn.disabled=true; let ok=0;
  for(const it of bulkItems){ const {err,...data}=it; try{ await dbCreate('questions',{...data, attachments:[], status:'open', createdAt:new Date().toISOString(), createdBy:S.me.email}); ok++; }catch(e){ toast(dbErr(e),true); break; } }
  if(ok){ toast(`Imported ${ok} question${ok>1?'s':''}`); closeModal(); refresh(); } else btn.disabled=false;
}

/* Staff editor */
function openStaffEditor(id){
  const s=id?staffById(id):null; if(s && !canEditStaffRec(s)) return;
  openModal(`${modalHead(s?'Edit staff member':'Add staff member', s?esc(s.name):'New team member')}
  <form id="s-form"><div class="modal-body">
    <div class="form-grid"><div class="field span-2"><label for="se-name">Name</label><input id="se-name" required maxlength="80" value="${esc(s?.name||'')}" placeholder="e.g. Ms R. Okafor" autofocus></div>
    <div class="field span-2"><label for="se-email">School email</label><input id="se-email" type="email" maxlength="120" value="${esc(s?.email||'')}" placeholder="r.okafor@yourschool.org.uk"><span class="hint">The address they sign in to Teams with. This is how the app recognises them.</span></div>
    <div class="field span-2"><label for="se-role">Role</label><select id="se-role">${Object.entries(ROLE_LABEL).filter(([k])=>isCC()||k!=='cc').map(([k,v])=>`<option value="${k}" ${(s?.role||'staff')===k?'selected':''}>${v}</option>`).join('')}</select><span class="hint">Heads of subject see questions and team results for their subjects. Curriculum coordinators see everything.</span></div>
    <div class="field span-4"><span class="lbl">Key stages</span><div class="checks">${LEVELS.map(l=>`<label><input type="checkbox" name="se-lv" value="${l}" ${(s?.levels||[]).includes(l)?'checked':''}>${l}</label>`).join('')}</div></div>
    <div class="field span-4"><span class="lbl">Subjects</span><div class="checks">${editableSubjects().map(x=>`<label><input type="checkbox" name="se-sub" value="${esc(x)}" ${(s?.subjects||[]).includes(x)?'checked':''}>${esc(x)}</label>`).join('')}</div><span class="hint">A teacher is asked to mark every script that matches one of their key stages and one of their subjects.</span></div></div>
    <p class="small muted">In Teams, heads of subject and coordinators must be <b>owners</b> of the Science team, and teachers <b>members</b>.</p>
  </div><div class="modal-foot"><button class="btn" type="button" data-act="modal-close">Cancel</button><button class="btn primary" type="submit">${s?'Save changes':'Add staff member'}</button></div></form>`);
  $('#s-form').addEventListener('submit', async e=>{
    e.preventDefault();
    const email=lc($('#se-email').value);
    const data={name:$('#se-name').value.trim(), email, role:$('#se-role').value, levels:[...document.querySelectorAll('input[name=se-lv]:checked')].map(x=>x.value), subjects:[...new Set([...(s?.subjects||[]).filter(x=>!editableSubjects().includes(x)), ...[...document.querySelectorAll('input[name=se-sub]:checked')].map(x=>x.value)])], demo:s?.demo||false};
    if(!data.name){ toast('Enter a name.',true); return; }
    if(email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)){ toast('That email address doesn’t look right.',true); return; }
    const clash=email && S.staff.find(x=>lc(x.email)===email && x.id!==s?.id);
    if(clash){ toast('That email is already used by '+clash.name+'.',true); return; }
    try{ if(s) await dbUpdate('staff',s.id,data); else await dbCreate('staff',{...data, createdAt:new Date().toISOString()}); toast(s?'Staff member updated':'Staff member added'); closeModal(); refresh(); }
    catch(err){ toast(dbErr(err),true); }
  });
}

/* ---------- sample data ---------- */
const DEMO_STAFF=[
  {k:'a',name:'A. Patel (sample)',role:'hos',levels:['KS4','KS5'],subjects:['Biology']},
  {k:'b',name:'J. Morgan (sample)',role:'staff',levels:['KS3','KS4'],subjects:['Biology','General Science']},
  {k:'c',name:'L. Chen (sample)',role:'staff',levels:['KS3','KS4','KS5'],subjects:['Biology','Chemistry']},
  {k:'d',name:'R. Okafor (sample)',role:'hos',levels:['KS4','KS5'],subjects:['Chemistry']},
  {k:'e',name:'S. Davies (sample)',role:'staff',levels:['KS3','KS4'],subjects:['Chemistry','Physics','General Science']},
  {k:'f',name:'T. Williams (sample)',role:'staff',levels:['KS3','KS4','KS5'],subjects:['Physics','General Science']},
];
const BIAS={a:0,b:.12,c:-.03,d:.02,e:-.13,f:.05};
const DEMO_Q=[
  {k:1,title:'Y11 Mock P1 Q04 – Osmosis in potato cylinders',level:'KS4',subject:'Biology',topic:'Paper 1 · Cell biology',maxMarks:6,ref:4,q:'Describe how you would investigate the effect of sugar concentration on the mass of potato cylinders. [6 marks]',ms:'Level 3 (5–6): a coherent method with dependent/independent variables, control variables and repeat measurements.\nLevel 2 (3–4): a method with some relevant detail; may lack control variables.\nLevel 1 (1–2): simple relevant statements.\nIndicative content: cut cylinders of equal length with a cork borer; blot dry and weigh; place in range of concentrations (0–1.0 M); leave for a set time (e.g. 30 min); blot and reweigh; calculate % change in mass; repeat and average.',r:'I would cut the potato into pieces the same size and weigh them. Then put them in different sugar solutions like 0.2, 0.4, 0.6 for 20 minutes. Then dry them and weigh again and see if they get heavier or lighter. I would keep the temperature the same.'},
  {k:2,title:'Y10 Respiration 4-marker',level:'KS4',subject:'Biology',topic:'Paper 2 · Bioenergetics',maxMarks:4,ref:null,q:'Compare aerobic and anaerobic respiration in muscle cells. [4 marks]',ms:'1 mark each, max 4: aerobic uses oxygen / anaerobic does not; aerobic produces CO2 and water, anaerobic produces lactic acid; aerobic releases more energy; anaerobic creates an oxygen debt.',r:'Aerobic needs oxygen and makes carbon dioxide and water. Anaerobic doesn’t need oxygen and makes lactic acid which causes cramp.'},
  {k:3,title:'A-level Paper 2 – Gene expression 9-marker',level:'KS5',subject:'Biology',topic:'Paper 2 · Control of gene expression',maxMarks:9,ref:6,q:'Explain how oestrogen stimulates transcription of a target gene. [9 marks]',ms:'Level 3 (7–9) detailed, logically sequenced; Level 2 (4–6) mostly correct with omissions; Level 1 (1–3) fragmented.\nIndicative: lipid-soluble, diffuses through membrane; binds receptor on transcription factor; changes shape; TF released from inhibitor; enters nucleus; binds promoter; RNA polymerase binds; mRNA produced.',r:'Oestrogen is a steroid so it goes through the membrane. It binds to a receptor which is part of a transcription factor and makes it change shape so it can bind to DNA at the promoter region. This lets RNA polymerase start transcription and mRNA is made.'},
  {k:4,title:'Y8 Particle model – states of matter',level:'KS3',subject:'General Science',topic:'Y8 end of unit test',maxMarks:4,ref:3,q:'Use the particle model to explain why a gas can be compressed but a solid cannot. [4 marks]',ms:'Gas particles far apart (1); spaces between them so can be pushed closer (1); solid particles touching/closely packed (1); no space to move into (1).',r:'In a gas the particles are spread out so you can squash them together. In a solid they are packed tight.'},
  {k:5,title:'Y11 Titration calculation',level:'KS4',subject:'Chemistry',topic:'Paper 2 · Quantitative chemistry',maxMarks:5,ref:3,q:'25.0 cm³ of NaOH was neutralised by 21.40 cm³ of 0.100 mol/dm³ HCl. Calculate the concentration of NaOH. [5 marks]',ms:'moles HCl = 0.1 × 21.40/1000 = 2.14×10⁻³ (1); ratio 1:1 (1); moles NaOH = 2.14×10⁻³ (1); conc = 2.14×10⁻³ / 0.025 (1); = 0.0856 mol/dm³ (1).',r:'moles = 0.1 x 21.4 = 2.14\nNaOH 1:1 so 2.14\n2.14 / 25 = 0.0856 mol/dm3'},
  {k:6,title:'A-level Organic mechanism – electrophilic addition',level:'KS5',subject:'Chemistry',topic:'Paper 2 · Alkenes',maxMarks:4,ref:null,q:'Outline the mechanism for the reaction of propene with HBr. [4 marks]',ms:'Curly arrow from C=C to H (1); arrow from H–Br bond to Br (1); correct carbocation (1); arrow from Br⁻ lone pair to C+ (1).',r:'Arrows drawn from double bond to H, and from bond to Br. Secondary carbocation shown. Br- attacks carbocation. (Lone pair not shown on Br-.)'},
  {k:7,title:'Y10 Specific heat capacity required practical',level:'KS4',subject:'Physics',topic:'Paper 1 · Energy',maxMarks:6,ref:4,q:'Describe a method to determine the specific heat capacity of a metal block. [6 marks]',ms:'Level-marked 6-marker. Indicative: measure mass; insulate block; heater and thermometer in holes; measure initial temp; switch on with joulemeter (or V, I, t); record energy and temperature rise; c = E/(mΔθ).',r:'Weigh the block, put the heater in and a thermometer. Turn on for 10 minutes and use a joulemeter. Record the temp at the start and end. Use E = mcΔθ to work out c. Wrap it in insulation to stop heat loss.'},
  {k:8,title:'Y9 Forces – resultant force',level:'KS3',subject:'Physics',topic:'Y9 assessment 2',maxMarks:3,ref:2,q:'A car has a driving force of 3000 N and air resistance of 1200 N. Describe and explain its motion. [3 marks]',ms:'Resultant 1800 N (1); forwards (1); so car accelerates (1).',r:'The forces are unbalanced so it speeds up.'},
];
function demoMark(qk, sk, max, ref){
  const seed=(qk*31+sk.charCodeAt(0)*17)%97/97; const base=ref!=null?ref:max*0.55;
  let m=base + BIAS[sk]*max + (seed-0.5)*max*0.18; return Math.max(0,Math.min(max,Math.round(m)));
}
async function loadDemo(){
  if(!isCC()){ toast('Only the curriculum coordinator can load sample data.',true); return; }
  if(S.staff.some(s=>s.demo)){ toast('Sample data is already loaded.'); return; }
  try{
    const sid={}, qid={}; let n=0;
    const step=()=>toast(`Loading sample team… ${++n}`);
    for(const d of DEMO_STAFF){ sid[d.k]=await dbCreate('staff',{name:d.name,email:'',role:d.role,levels:d.levels,subjects:d.subjects.filter(x=>S.config.subjects.includes(x)),demo:true,createdAt:new Date().toISOString()}); step(); }
    S.staff=S.staff.concat(DEMO_STAFF.map(d=>({id:sid[d.k],name:d.name})));
    const t0=Date.now();
    for(const q of DEMO_Q){
      const dl=new Date(t0+(q.k%3===0?-2:q.k*3)*86400000).toISOString().slice(0,10);
      qid[q.k]=await dbCreate('questions',{title:q.title,level:q.level,subject:q.subject,topic:q.topic,maxMarks:q.maxMarks,deadline:dl,questionText:q.q,markScheme:q.ms,response:q.r,referenceMark:q.ref,attachments:[],status:'open',demo:true,createdAt:new Date(t0-q.k*3600000).toISOString(),createdBy:S.me.email});
      S.questions=S.questions.concat([{id:qid[q.k],title:q.title}]); step();
    }
    for(const q of DEMO_Q){
      for(const [i,d] of DEMO_STAFF.entries()){
        if(!d.levels.includes(q.level)||!d.subjects.includes(q.subject)||(q.k+i)%7===0) continue;
        await dbCreate('marks',{staffId:sid[d.k],questionId:qid[q.k],mark:demoMark(q.k,d.k,q.maxMarks,q.ref),note:'',at:new Date(t0-(i+q.k)*1800000).toISOString(),by:S.me.email,demo:true}); step();
      }
    }
    await refresh(); toast('Sample team loaded. Try “View a dashboard as” in the sidebar.');
  }catch(e){ toast(dbErr(e),true); refresh(); }
}
async function clearDemo(){
  try{
    const demoStaff=new Set(S.staff.filter(s=>s.demo).map(s=>s.id));
    for(const m of S.markItems.filter(m=>m.demo||demoStaff.has(m.staffId))) await dbDelete('marks',m.id);
    for(const id of demoStaff) await dbDelete('staff',id);
    for(const q of S.questions.filter(q=>q.demo)) await dbDelete('questions',q.id);
    S.viewAs=''; await refresh(); toast('Sample data removed');
  }catch(e){ toast(dbErr(e),true); refresh(); }
}

/* ---------- events ---------- */
document.addEventListener('click', async e=>{
  const el=e.target.closest('[data-act]'); if(!el) return;
  const act=el.dataset.act, id=el.dataset.id;
  if(act==='overlay'){ if(e.target===el) closeModal(); return; }
  if(el.dataset.confirm && !el.dataset.armed){ el.dataset.armed='1'; const old=el.textContent; el.textContent=el.dataset.confirm; setTimeout(()=>{ if(el.isConnected){ delete el.dataset.armed; el.textContent=old; } },4000); return; }
  switch(act){
    case 'modal-close': closeModal(); break;
    case 'mark': openMark(id); break;
    case 'mark-del': { const s=targetStaff(); try{ await removeMark(s,id); toast('Mark withdrawn'); closeModal(); refresh(); }catch(err){ toast(dbErr(err),true); } break; }
    case 'q-new': openQuestionEditor(null); break;
    case 'q-edit': openQuestionEditor(id); break;
    case 'q-results': openResults(id); break;
    case 'q-toggle': { const q=S.questions.find(x=>x.id===id); try{ await dbUpdate('questions',id,{status:(q.status||'open')==='closed'?'open':'closed'}); await refresh(); }catch(err){ toast(dbErr(err),true); } break; }
    case 'q-del': try{ await dbDelete('questions',id); for(const m of S.markItems.filter(m=>m.questionId===id)) await dbDelete('marks',m.id); await refresh(); toast('Question deleted'); }catch(err){ toast(dbErr(err),true); } break;
    case 'bulk': openBulk(); break;
    case 'bulk-go': runBulk(); break;
    case 'quick-split': quickSplit(); break;
    case 'zoom': { const p=el.dataset.pane, d=+el.dataset.d; setZoom(p, d===0?1:VIEW.z[p]*(d>0?1.25:0.8)); break; }
    case 'split-set': setSplit(+el.dataset.v); saveView(); break;
    case 'att-del': draftAttach.splice(+el.dataset.i,1); renderAttachList(); break;
    case 'staff-new': openStaffEditor(null); break;
    case 'staff-edit': openStaffEditor(id); break;
    case 'staff-del': if(!canEditStaffRec(staffById(id)||{})) break; try{ await dbDelete('staff',id); await refresh(); toast('Staff member removed'); }catch(err){ toast(dbErr(err),true); } break;
    case 'signin': signIn(); break;
    case 'retry': location.reload(); break;
    case 'refresh': refresh(); break;
    case 'create-lists': createLists(); break;
    case 'viewas': S.viewAs=id; location.hash='#dashboard'; render(); window.scrollTo(0,0); break;
    case 'viewas-clear': S.viewAs=''; render(); break;
    case 'demo-load': loadDemo(); break;
    case 'demo-clear': clearDemo(); break;
    case 'save-settings': {
      const tol=parseFloat($('#set-tol').value), subs=$('#set-subs').value.split(',').map(x=>x.trim()).filter(Boolean);
      if(!(tol>0&&tol<=50)){ toast('Tolerance must be between 1 and 50%.',true); break; }
      if(!subs.length){ toast('List at least one subject.',true); break; }
      try{ const d={tolerance:tol,subjects:[...new Set(subs)]}; if(S.settingsId) await Store.update('settings',S.settingsId,'Settings',d); else await Store.create('settings','Settings',d); await refresh(); toast('Settings saved'); }catch(err){ toast(dbErr(err),true); }
      break; }
  }
});
document.addEventListener('change', e=>{
  const el=e.target;
  if(el.dataset.change==='viewas'){ S.viewAs=el.value; location.hash='#dashboard'; render(); return; }
  if(el.dataset.change==='qf'){ S.qf[el.dataset.key]=el.value; const l=$('#qlist'); if(l) l.innerHTML=questionList(); return; }
  if(el.dataset.change==='tf'){ S.tf[el.dataset.key]=el.value; render(); return; }
  if(el.dataset.import && el.files[0]){ importInto(el.dataset.import, el.files[0]); el.value=''; return; }
  if(el.id==='qe-attach'){ [...el.files].forEach(f=>uploadAttachment(f).then(ok=>ok&&toast(f.name+' attached'))); el.value=''; }
});
document.addEventListener('input', e=>{
  const el=e.target; if(el.dataset.input==='qf'){ S.qf[el.dataset.key]=el.value; const l=$('#qlist'); if(l) l.innerHTML=questionList(); }
});
document.addEventListener('dragover', e=>{ const t=e.target.closest && e.target.closest('[data-drop]'); if(t){ e.preventDefault(); t.classList.add('dragover'); } });
document.addEventListener('dragleave', e=>{ const t=e.target.closest && e.target.closest('[data-drop]'); if(t) t.classList.remove('dragover'); });
document.addEventListener('drop', e=>{ const t=e.target.closest && e.target.closest('[data-drop]'); if(t && e.dataTransfer.files.length){ e.preventDefault(); t.classList.remove('dragover'); importInto(t.dataset.drop, e.dataTransfer.files[0]); } });
document.addEventListener('keydown', e=>{ if(e.key==='Escape' && $('#modal-root').innerHTML) closeModal(); });
window.addEventListener('hashchange', ()=>{ render(); window.scrollTo(0,0); });
boot();
