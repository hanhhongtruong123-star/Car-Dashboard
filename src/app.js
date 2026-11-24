(function(){
'use strict';

/* =====================================================================
   CONFIG – assumptions taken from the Excel "Analysis" sheet (Oct 2026)
   ===================================================================== */
const CFG = {
  // Grab is hidden while the focus is Vegas Car. Owner/editors switch it in the Admin panel (saved for everyone on the hosted page).
  features: { grab: false },
  hostTeams: ['HM','CM','IM','BD','GM','Non-Assigned'],
  teamDrivers: 18,            // Analysis C39
  dedicatedDrivers: ['35'],   // boss's dedicated driver, excluded from the guest pool
  offDays: 3, cycleDays: 14,  // min off days per 2 weeks
  maxDrive: 6,                // productive driving h / driver / shift (8h − 2h break) — estimate when no roster
  breakH: 2,                  // break deducted from a rostered shift (Driver Roster 'Break' column overrides)
  buffer: 0.125,              // pre-trip buffer per trip (h) = 7.5 min
  fleet: 9,                   // active guest fleet (vehicles)
  retiredCars: ['Benz S400','Benz S500'], // no longer in use (owner, Nov 2026): kept in history and cost, left out of fleet KPIs
  internalCars: ['Bus'],      // internal use only: left out of fleet KPIs
  bannedTiers: ['X'],         // tier code X = banned guest
  grabMonthlyLimit: 300,      // USD: a guest whose Grab cost passes this in one calendar month is shown in red
  serviceWindow: 16,          // vehicle service window (h/day)
  targetUtil: 0.65,           // target vehicle utilization
  cutoffMin: 330,             // shift-day cut-off 05:30
  targetTiers: ['V','V+','I+'],   // = Premium Level
  areas: ['Ho Chi Minh','Binh Duong','Dong Nai','Long An','Others'], // main areas by CITY column (airport trips are in Ho Chi Minh)
  levels: [['VIP Level',['ONE','ONE+','EC']],['Premium Level',['I+','V','V+']],['Potential Level',['P','I']]], // tiers listed low -> high
  tierRank: ['P','I','I+','V','V+','ONE','ONE+','EC'], // lowest -> highest; Trip mix donut runs clockwise in this order
  tierOrder: ['EC','ONE+','ONE','V+','V','I+','I','P','X'],
  shifts: [['Morning 06:00–14:00',360,840],['Afternoon 14:00–22:00',840,1320],['Night 22:00–06:00',1320,1800]],
  cappedTasks: ['Maintenance Car','Take Car Back'], // driver hrs capped at 0.5h (same rule as source sheet)
  costCats: ['Maintenance','Depreciation','Oil & Gasoline','Toll Fee','Insurance','Miscellaneous'],
};
const STORE_KEY = 'carServiceDash.v1';
const WD = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'];
const MON = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

/* ============================ utils ============================ */
const $ = s => document.querySelector(s);
const esc = s => String(s==null?'':s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const sum = (a,f) => a.reduce((s,x)=>s+(f?f(x):x),0);
const uniq = a => [...new Set(a)];
const groupBy = (a,f) => { const m=new Map(); for(const x of a){const k=f(x); if(!m.has(k)) m.set(k,[]); m.get(k).push(x);} return m; };
const countBy = (a,f) => { const m=new Map(); for(const x of a){const k=f(x); m.set(k,(m.get(k)||0)+1);} return m; };
const topKey = m => { let b=null,bv=-1; for(const [k,v] of m) if(v>bv){b=k;bv=v;} return b; };
const pad = n => String(n).padStart(2,'0');
const fmt = {
  int: v => v==null||!isFinite(v)?'–':Math.round(v).toLocaleString('en-US'),
  d1: v => v==null||!isFinite(v)?'–':v.toLocaleString('en-US',{minimumFractionDigits:1,maximumFractionDigits:1}),
  d2: v => v==null||!isFinite(v)?'–':v.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2}),
  pct: v => v==null||!isFinite(v)?'–':(v*100).toLocaleString('en-US',{maximumFractionDigits:1})+'%',
  usd: v => v==null||!isFinite(v)?'–':'$'+v.toLocaleString('en-US',{maximumFractionDigits:v<100?2:0}),
};

// dates as 'YYYY-MM-DD' strings, arithmetic in UTC to avoid TZ drift
const isoOf = d => d.getUTCFullYear()+'-'+pad(d.getUTCMonth()+1)+'-'+pad(d.getUTCDate());
const dOf = s => new Date(Date.UTC(+s.slice(0,4), +s.slice(5,7)-1, +s.slice(8,10)));
const addDays = (s,n) => { const d=dOf(s); d.setUTCDate(d.getUTCDate()+n); return isoOf(d); };
const diffDays = (a,b) => Math.round((dOf(b)-dOf(a))/864e5);
const weekday = s => (dOf(s).getUTCDay()+6)%7; // 0 = Mon
const serialToIso = n => isoOf(new Date(Date.UTC(1899,11,30) + Math.round(n*864e5)));
const fmtDay = s => +s.slice(8,10)+' '+MON[+s.slice(5,7)-1]+' '+s.slice(0,4);

/* ======================= data normalisation ======================= */
// One code path for the built-in seed and for uploaded files:
// both are "array of arrays" (SheetJS header:1 / raw shape).
const norm = h => String(h==null?'':h).toLowerCase().replace(/[^a-z0-9]/g,'');
function toIso(v){
  if(v==null||v==='') return null;
  if(v instanceof Date) return isoOf(new Date(Date.UTC(v.getFullYear(),v.getMonth(),v.getDate())));
  if(typeof v==='number') return v>20000 ? serialToIso(v) : null;
  const s=String(v).trim();
  let m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/); if(m) return m[1]+'-'+pad(m[2])+'-'+pad(m[3]);
  m=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); if(m) return m[3]+'-'+pad(m[2])+'-'+pad(m[1]); // dd/mm/yyyy
  return null;
}
function toMin(v){ // time of day or duration -> minutes
  if(v==null||v==='') return null;
  if(v instanceof Date) return v.getHours()*60+v.getMinutes();
  if(typeof v==='number') return Math.round((v%1===v||v<1?v:v%1)*1440);
  const m=String(v).trim().match(/^(\d{1,2}):(\d{2})/); return m? +m[1]*60+ +m[2] : null;
}
function findHeader(aoa, must){
  for(let i=0;i<Math.min(aoa.length,12);i++){
    const row=(aoa[i]||[]).map(norm);
    if(must.every(k=>row.includes(k))) return i;
  }
  return -1;
}
function parseTrips(aoa, src){
  const hi=findHeader(aoa,['guestid','date','depart']); if(hi<0) return null;
  const H=(aoa[hi]||[]).map(norm), col=(...ks)=>{ for(const k of ks){ const i=H.indexOf(k); if(i>=0) return i; } return -1; };
  const c={gid:col('guestid'),type:col('type'),tier:col('customertier','level','tier'),nat:col('nationality','nation'),date:col('date'),dep:col('depart'),arr:col('arrival'),
    occ:col('occupancy'),hr:col('hour','hours'),km:col('distance','distancekm'),arr_by:col('arrangby'),hname:col('confirm','hostname'),drv:col('employeeid'),car:col('brand'),dest:col('destination'),city:col('city'),
    team:H.findIndex(h=>h==='hostteam'||h==='team'||h==='arrangebyteam'||h==='arrangbyteam'),
    cost:H.findIndex(h=>h==='cost'||h==='costusd'||h==='amountusd'||h==='fare'||h==='fareusd'||h==='grab'||h==='grabusd')};
  const out=[];
  for(let r=hi+1;r<aoa.length;r++){
    const row=aoa[r]; if(!row) continue;
    const d=toIso(row[c.date]); if(!d) continue;
    const dep=toMin(row[c.dep]); if(dep==null) continue;
    let dur=c.occ>=0?toMin(row[c.occ]):null;
    if(dur==null && c.arr>=0){ const a=toMin(row[c.arr]); if(a!=null){ dur=a-dep; if(dur<0) dur+=1440; } }
    if(dur==null && c.hr>=0 && row[c.hr]!=null && row[c.hr]!=='' && isFinite(+row[c.hr])) dur=Math.round(+row[c.hr]*60); // Grab sheet: trip length in hours
    if(dur==null) dur=0;
    const gid=String(row[c.gid]==null?'':row[c.gid]).trim();
    const type=String(row[c.type]||'Guest').trim();
    const dest=c.dest>=0?String(row[c.dest]||'').trim():'';
    const capped = CFG.cappedTasks.includes(gid) || dest==='Garage';
    const h=dur/60;
    out.push({
      src, gid, type: /internal/i.test(type)?'Internal':'Guest',
      tier: c.tier>=0&&row[c.tier]!=null&&row[c.tier]!==''?String(row[c.tier]).trim():null,
      nat: c.nat>=0&&row[c.nat]!=null&&row[c.nat]!==''?String(row[c.nat]).trim():null,
      d, dep, dur, km: +row[c.km]||0,
      drv: src==='V'&&c.drv>=0&&row[c.drv]!=null?String(row[c.drv]).trim():null,
      team: c.team>=0&&row[c.team]!=null&&row[c.team]!==''?String(row[c.team]).trim():null,
      host: c.arr_by>=0&&row[c.arr_by]!=null&&row[c.arr_by]!==''?String(row[c.arr_by]).trim():null, // Arrangby = host who booked the trip
      hname: c.hname>=0&&row[c.hname]!=null&&String(row[c.hname]).trim()?String(row[c.hname]).trim():null, // Confirm = that host's name
      car: src==='V'? (c.car>=0&&row[c.car]?String(row[c.car]).trim():'Unknown') : 'Grab',
      dest: dest||'Unknown', city: c.city>=0&&row[c.city]?String(row[c.city]).trim():'Unknown',
      cost: c.cost>=0? (+row[c.cost]||0) : null,
      dh: capped? Math.min(h,0.5) : h,
    });
  }
  return out;
}
// Car Cost sheet: Date, Vehicle, Category, Amount (USD) — one row per expense. A legacy "Month" column (YYYY-MM) is still read as the 1st of that month.
// Also accepts the VG Car Cost file layout: Date | Car Brand | Cost Type | Amount (Usd).
const COST_COLS = {date:['date','month'], car:['vehicle','carbrand','car'], cat:['category','costtype'], amt:['amountusd','amount']};
function parseCosts(aoa){
  let hj=-1, ci=null;
  for(let i=0;i<Math.min(aoa.length,12)&&hj<0;i++){
    const H=(aoa[i]||[]).map(norm), pick=ks=>{ for(const k of ks){ const j=H.indexOf(k); if(j>=0) return j; } return -1; };
    const m={date:pick(COST_COLS.date),car:pick(COST_COLS.car),cat:pick(COST_COLS.cat),amt:pick(COST_COLS.amt)};
    if(Object.values(m).every(j=>j>=0)){ hj=i; ci=m; }
  }
  if(hj<0) return null;
  const out=[];
  for(let r=hj+1;r<aoa.length;r++){
    const row=aoa[r]; if(!row) continue;
    const v=row[ci.date]; let iso=toIso(v);
    if(!iso && v!=null){ const mm=String(v).match(/^(\d{4})-(\d{1,2})$/); if(mm) iso=mm[1]+'-'+pad(mm[2])+'-01'; }
    const a=+row[ci.amt]; if(!iso||row[ci.amt]==null||row[ci.amt]===''||!isFinite(a)||!row[ci.car]) continue;
    out.push({d:iso, m:iso.slice(0,7), car:String(row[ci.car]).trim(), cat:String(row[ci.cat]||'Miscellaneous').trim(), amt:a});
  }
  return out;
}
const costDay = c => c.d || (c.m+'-01'); // older stored rows carry only the month
// Driver Roster sheet: Date | EmployeeId | Start | End [| Break (h)] [| Status]. Blank Start/End or Status Off/Leave = day off.
// A shift whose End is not after Start runs past midnight (22:00–06:00 ends the next morning).
function parseRoster(aoa){
  const hi=findHeader(aoa,['date','employeeid','start','end']); if(hi<0) return null;
  const H=aoa[hi].map(norm), ci=k=>H.indexOf(k), bi=H.findIndex(h=>h==='break'||h==='breakh'||h==='breakhours'), si=ci('status');
  const out=[];
  for(let r=hi+1;r<aoa.length;r++){
    const row=aoa[r]; if(!row) continue;
    const d=toIso(row[ci('date')]), drv=row[ci('employeeid')]; if(!d||drv==null||drv==='') continue;
    const st=si>=0&&row[si]!=null?String(row[si]).trim():'';
    let s=toMin(row[ci('start')]), e=toMin(row[ci('end')]);
    const off=/off|leave|nghỉ|nghi|absent/i.test(st)||s==null||e==null;
    if(!off&&e<=s) e+=1440;
    const brk=bi>=0&&row[bi]!==''&&row[bi]!=null&&isFinite(+row[bi])?+row[bi]:CFG.breakH;
    out.push({d, drv:String(drv).trim(), s:off?null:s, e:off?null:e, brk, off});
  }
  return out;
}
// Drivers sheet: EmployeeId | Name (or Driver Name)
function parseDriverNames(aoa){
  let hi=findHeader(aoa,['employeeid','name']); if(hi<0) hi=findHeader(aoa,['employeeid','drivername']); if(hi<0) return null;
  const H=aoa[hi].map(norm); if(H.includes('date')||H.includes('guestid')) return null;
  const ni=H.includes('name')?H.indexOf('name'):H.indexOf('drivername'), out={};
  for(let r=hi+1;r<aoa.length;r++){ const row=aoa[r]; if(!row) continue; const id=row[H.indexOf('employeeid')], nm=row[ni];
    if(id!=null&&id!==''&&nm!=null&&String(nm).trim()) out[String(id).trim()]=String(nm).trim(); }
  return out;
}
function parseWorkbookSheets(sheets){ // sheets: [{name, aoa}]
  const res={trips:[],costs:[],roster:[],drivers:{},log:[]};
  for(const s of sheets){
    const isGrab=/grab/i.test(s.name);
    const t=parseTrips(s.aoa, isGrab?'G':'V');
    if(t&&t.length){ res.trips.push(...t); { const ds=t.map(x=>x.d).sort(); res.log.push(`✓ "${s.name}": ${t.length} ${isGrab?'Grab':'Vegas Car'} trips (${ds[0]} … ${ds[ds.length-1]})`); }; continue; }
    const ro=parseRoster(s.aoa);
    if(ro&&ro.length){ res.roster.push(...ro); const ds=ro.map(x=>x.d).sort(); res.log.push(`✓ "${s.name}": ${ro.length} driver shift rows (${ds[0]} … ${ds[ds.length-1]})`); continue; }
    const c=parseCosts(s.aoa);
    if(c&&c.length){ res.costs.push(...c); res.log.push(`✓ "${s.name}": ${c.length} car-cost rows`); continue; }
    const dn=parseDriverNames(s.aoa);
    if(dn&&Object.keys(dn).length){ Object.assign(res.drivers,dn); res.log.push(`✓ "${s.name}": ${Object.keys(dn).length} driver names`); continue; }
  }
  return res;
}

/* ============================ state ============================ */
const S = { trips:[], costs:[], roster:[], drivers:{}, source:'built-in', f:{gran:'month',period:null,compare:'prev',service:'all',ttype:'all'}, mod:'overview', extra:{} };

function loadSeed(){
  const raw=JSON.parse(document.getElementById('seed').textContent);
  const r=parseWorkbookSheets(raw.sheets);
  S.trips=r.trips; S.costs=r.costs; S.roster=r.roster; S.drivers=r.drivers; S.source=raw.label||'built-in'; indexRoster();
}
function saveLocal(){ try{ localStorage.setItem(STORE_KEY, JSON.stringify({trips:S.trips,costs:S.costs,roster:S.roster,drivers:S.drivers,source:S.source})); }catch(e){} }
function loadLocal(){ try{ const v=JSON.parse(localStorage.getItem(STORE_KEY)||'null'); if(v&&Array.isArray(v.trips)&&v.trips.length){ S.trips=v.trips; S.costs=v.costs||[]; S.roster=v.roster||[]; S.drivers=v.drivers||{}; S.source=v.source||'uploaded'; indexRoster(); return true; } }catch(e){} return false; }

function mergeData(r){
  // new trips replace existing trips of the same service on the same dates; costs replace same month+vehicle+category
  for(const src of ['V','G']){
    const nd=new Set(r.trips.filter(t=>t.src===src).map(t=>t.d));
    if(nd.size) S.trips=S.trips.filter(t=>!(t.src===src&&nd.has(t.d)));
  }
  S.trips.push(...r.trips);
  const ck=c=>costDay(c)+'|'+c.car+'|'+c.cat, nk=new Set(r.costs.map(ck));
  S.costs=S.costs.filter(c=>!nk.has(ck(c))).concat(r.costs);
  const rk=x=>x.drv+'|'+x.d, rn=new Set((r.roster||[]).map(rk));          // a re-uploaded driver-day replaces the old one
  S.roster=S.roster.filter(x=>!rn.has(rk(x))).concat(r.roster||[]);
  S.drivers=Object.assign({},S.drivers,r.drivers||{});
  S.trips.sort((a,b)=>a.d<b.d?1:a.d>b.d?-1:b.dep-a.dep);
  indexRoster();
}
/* ---- roster helpers ---- */
let ROSTER_IDX=new Map(), ROSTER_DAYS=new Set();
function indexRoster(){ ROSTER_IDX=new Map(); ROSTER_DAYS=new Set();
  for(const r of S.roster){ ROSTER_IDX.set(r.drv+'|'+r.d,r); ROSTER_DAYS.add(r.d); }
  // the cost sheet may spell a car differently from the trip log (SEDONA / Sedona): use the trip-log spelling
  const cars=new Map(); for(const t of S.trips) if(t.src==='V'&&!cars.has(t.car.toLowerCase())) cars.set(t.car.toLowerCase(),t.car);
  for(const c of S.costs){ const k=cars.get(c.car.toLowerCase()); if(k) c.car=k; }
  // guard (security team) = Arrangby; the name comes from the Confirm column and applies to every month the ID appears in
  GUARD_NAMES=new Map(); for(const t of S.trips) if(t.host&&t.hname&&!GUARD_NAMES.has(t.host)) GUARD_NAMES.set(t.host,t.hname);
}
let GUARD_NAMES=new Map();
const guardName = id => GUARD_NAMES.has(id)? `${GUARD_NAMES.get(id)} (${id})` : 'Guard '+id;
// PLACEHOLDER (owner, Nov 2026): real host-team data comes later. Until then every Vegas guest trip is given a random host team
// (stable: the same trip always lands in the same team). A "Host Team" column in an uploaded trip log replaces it automatically.
const HOST_PLACEHOLDER_NOTE = 'Placeholder data: each Vegas Car guest trip is assigned to a host team at random until the real host-team data is uploaded.';
const driverName = id => { const n=S.drivers&&S.drivers[id]; if(!n) return id!=null?'Driver '+id:'–';
  return Object.entries(S.drivers).some(([k,v])=>k!==String(id)&&String(v).toLowerCase()===n.toLowerCase())? n+' ('+id+')' : n; }; // two drivers share a name: add the ID
const shiftHours = r => r.off?0:(r.e-r.s)/60;
const availHours = r => r.off?0:Math.max(0,shiftHours(r)-(r.brk==null?CFG.breakH:r.brk));
// the rostered shift a trip belongs to: same day or a night shift that started the day before (30-min grace either side)
function rosterFor(t){
  for(const [d,off] of [[t.d,0],[addDays(t.d,-1),1440]]){ const r=ROSTER_IDX.get(t.drv+'|'+d);
    if(r&&!r.off&&t.dep+off>=r.s-30&&t.dep+off<=r.e+30) return r; }
  return null;
}
const hasRosterDay = (drv,d) => ROSTER_IDX.has(drv+'|'+d)||ROSTER_IDX.has(drv+'|'+addDays(d,-1));


/* ===================== shared uploads (artifact db) =====================
   On the hosted page, uploads by the owner/editors are saved in the artifact's shared store
   (collection "uploads") so every signed-in viewer sees them. Docs:
   trips-<src>-<YYYY-MM>-<i>  {kind:'trips', src, month, i, n, rows:[packed trip], at}
   cost-<YYYY-MM>             {kind:'cost', month, rows:[[vehicle, category, amount]], at}
   Shared rows replace built-in rows for the same service+date / cost month. */
const SHARED = { db:null, canWrite:false, docs:new Map(), ready:false, seed:null, armed:null };
const PACK = ['gid','type','tier','nat','d','dep','dur','km','host','drv','car','dest','city','cost','dh','team','hname']; // later fields appended last: older rows simply lack them
const pack = t => PACK.map(k=>t[k]==null?null:t[k]);
const unpack = (row,src) => { const t={src}; PACK.forEach((k,i)=>{ t[k]=row[i]; }); return t; };
const byNewest = (a,b)=>a.d<b.d?1:a.d>b.d?-1:b.dep-a.dep;
function applyShared(){
  const trips=[], costs=[], roster=[], names={}, dates=new Set();
  for(const d of SHARED.docs.values()){
    if(d.kind==='trips') for(const r of d.rows||[]){ const t=unpack(r,d.src); trips.push(t); dates.add(d.src+'|'+t.d); }
    else if(d.kind==='cost') for(const r of d.rows||[]) costs.push({d:r[3]||d.month+'-01',m:d.month,car:r[0],cat:r[1],amt:r[2]});
    else if(d.kind==='roster') for(const r of d.rows||[]) roster.push({drv:r[0],d:r[1],s:r[2],e:r[3],brk:r[4],off:!!r[5]});
    else if(d.kind==='drivers') Object.assign(names,d.map||{});
  }
  const cm=new Set(costs.map(c=>c.m));
  S.trips=SHARED.seed.trips.filter(t=>!dates.has(t.src+'|'+t.d)).concat(trips).sort(byNewest);
  S.costs=SHARED.seed.costs.filter(c=>!cm.has(c.m)).concat(costs);
  const rk=new Set(roster.map(r=>r.drv+'|'+r.d));
  S.roster=SHARED.seed.roster.filter(r=>!rk.has(r.drv+'|'+r.d)).concat(roster);
  S.drivers=Object.assign({},SHARED.seed.drivers,names);
  indexRoster();
  S.source=SHARED.docs.size?'built-in data + shared uploads':SHARED.seed.source;
}
async function shareUpload(r){
  const db=SHARED.db, at=new Date().toISOString();
  for(const [k,ts] of groupBy(r.trips,t=>t.src+'|'+t.d.slice(0,7))){
    const [src,month]=k.split('|');
    const prev=[...SHARED.docs.entries()].filter(([,d])=>d.kind==='trips'&&d.src===src&&d.month===month).sort((a,b)=>a[1].i-b[1].i);
    const newDates=new Set(ts.map(t=>t.d));
    const rows=prev.flatMap(([,d])=>d.rows||[]).filter(row=>!newDates.has(row[4])).concat(ts.map(pack));
    const parts=[]; for(let i=0;i<rows.length;i+=500) parts.push(rows.slice(i,i+500)); // ~60 KB per doc (limit 256 KiB)
    for(let i=0;i<parts.length;i++) await db.doc(`uploads/trips-${src}-${month}-${i}`).set({kind:'trips',src,month,i,n:parts.length,rows:parts[i],at});
    for(const [id,d] of prev) if(d.i>=parts.length) await db.doc('uploads/'+id).delete();
  }
  for(const [month,cs] of groupBy(r.costs,c=>c.m)){
    const prev=SHARED.docs.get('cost-'+month), nk=new Set(cs.map(c=>costDay(c)+'|'+c.car+'|'+c.cat));
    const rows=(prev?(prev.rows||[]).filter(x=>!nk.has((x[3]||month+'-01')+'|'+x[0]+'|'+x[1])):[]).concat(cs.map(c=>[c.car,c.cat,c.amt,costDay(c)]));
    await db.doc('uploads/cost-'+month).set({kind:'cost',month,rows,at});
  }
  for(const [month,rs] of groupBy(r.roster||[],x=>x.d.slice(0,7))){
    const prev=SHARED.docs.get('roster-'+month), nk=new Set(rs.map(x=>x.drv+'|'+x.d));
    const rows=(prev?(prev.rows||[]).filter(x=>!nk.has(x[0]+'|'+x[1])):[]).concat(rs.map(x=>[x.drv,x.d,x.s,x.e,x.brk,x.off?1:0]));
    await db.doc('uploads/roster-'+month).set({kind:'roster',month,rows,at});
  }
  if(r.drivers&&Object.keys(r.drivers).length){
    const prev=SHARED.docs.get('drivers'); await db.doc('uploads/drivers').set({kind:'drivers',map:Object.assign({},prev&&prev.map||{},r.drivers),at});
  }
}
function sharedGroups(){
  const g=new Map();
  for(const [id,d] of SHARED.docs){ const key=d.kind==='trips'?'trips|'+d.src+'|'+d.month:d.kind==='drivers'?'drivers':d.kind+'|'+d.month;
    if(!g.has(key)) g.set(key,{key,kind:d.kind,src:d.src,month:d.month,ids:[],count:0,at:d.at});
    const e=g.get(key); e.ids.push(id); e.count+=d.kind==='drivers'?Object.keys(d.map||{}).length:(d.rows||[]).length; if(d.at>e.at) e.at=d.at; }
  return [...g.values()].sort((a,b)=>(a.month||'9999')<(b.month||'9999')?1:-1);
}
function renderSharedList(){
  const el=$('#sharedList'); if(!el) return;
  if(!SHARED.db||!SHARED.canWrite){ el.innerHTML=''; return; }
  const gs=sharedGroups();
  el.innerHTML=`<div class="card-h" style="margin:14px 0 6px"><div><h3>Shared uploads</h3><div class="desc">Saved for everyone who opens this link (signed in to claude.ai)</div></div></div>`+
    (gs.length?gs.map(e=>`<div class="stat-row"><span class="l">${e.kind==='drivers'?'Driver names':e.kind==='cost'?'Car cost':e.kind==='roster'?'Driver roster':(e.src==='G'?'Grab':'Vegas Car')+' trips'}${e.month?' · '+esc(periodLabel('month',e.month)):''}</span>
      <span class="v num" style="display:flex;gap:10px;align-items:center">${fmt.int(e.count)} ${e.kind==='trips'?'trips':e.kind==='drivers'?'drivers':e.kind==='roster'?'shifts':'rows'}
      <button type="button" class="btn" data-unshare="${esc(e.key)}" style="padding:4px 10px">${SHARED.armed===e.key?'Click again to remove':'Remove'}</button></span></div>`).join('')
      :`<p class="muted" style="font-size:12px;margin:0">No shared uploads yet. The dashboard shows the built-in data (September – October 2026).</p>`);
}
document.addEventListener('click', async e=>{
  const b=e.target.closest&&e.target.closest('[data-unshare]'); if(!b||!SHARED.db) return;
  const key=b.dataset.unshare;
  if(SHARED.armed!==key){ SHARED.armed=key; renderSharedList(); return; } // two-step remove (no confirm dialogs in the viewer)
  SHARED.armed=null; const g=sharedGroups().find(x=>x.key===key); if(!g) return;
  b.disabled=true; b.textContent='Removing…';
  try{ for(const id of g.ids) await SHARED.db.doc('uploads/'+id).delete(); $('#log').textContent='Removed. Everyone now sees the dashboard without that upload.'; }
  catch(err){ $('#log').textContent='✗ Could not remove ('+(err&&err.code||'error')+'). Please try again.'; renderSharedList(); }
});
async function initShared(){
  if(!window.__ARTIFACT__||!window.claude||!window.claude.use) return;
  const db=await window.claude.use('db'); if(!db) return;          // signed-out visitors: built-in data only
  const user=await window.claude.use('user');
  SHARED.db=db; SHARED.canWrite=user?!!(await user.canEdit()):false; // write rule = admin (owner / editors)
  S.adminOK=SHARED.canWrite; renderNav();
  $('#btnReset').hidden=true;
  const tip=document.querySelector('#modal p.muted'); if(tip) tip.textContent='Uploads are saved for everyone who opens this link. Uploading the same month again replaces it.';
  $('#uploadNote').textContent='Upload the monthly Excel report. It is saved for everyone who opens this link.';
  const timer=setTimeout(()=>{ if(!SHARED.ready) $('#stamp').textContent+=' · shared data not reachable yet'; },10000);
  db.doc('settings/features').onSnapshot(d=>{ if(d.exists) applyGrab(!!(d.data()||{}).grab); }, err=>console.warn('settings:', err&&err.code));
  db.collection('uploads').onSnapshot(snap=>{
    SHARED.docs=new Map(snap.docs.map(d=>[d.id,d.data()]));
    if(!snap.metadata.fromCache){ SHARED.ready=true; clearTimeout(timer); }
    applyShared(); render(); renderSharedList();
  }, err=>{ clearTimeout(timer); $('#stamp').textContent+=' · shared data unavailable'; console.warn('shared uploads:', err&&err.code); });
  renderSharedList();
}

/* ======================= periods & context ======================= */
const dataBounds = () => { let mn=null,mx=null; for(const t of S.trips){ if(!mn||t.d<mn) mn=t.d; if(!mx||t.d>mx) mx=t.d; } return [mn,mx]; };
function periodKey(gran,iso){
  const y=iso.slice(0,4), m=+iso.slice(5,7);
  return gran==='day'?iso : gran==='month'?iso.slice(0,7) : gran==='quarter'? y+'-Q'+(Math.floor((m-1)/3)+1) : y;
}
function periodRange(gran,key){
  if(gran==='day') return [key,key];
  if(gran==='month'){ const s=key+'-01'; const y=+key.slice(0,4), m=+key.slice(5,7); return [s, isoOf(new Date(Date.UTC(y,m,0)))]; }
  if(gran==='quarter'){ const y=+key.slice(0,4), q=+key.slice(6); return [isoOf(new Date(Date.UTC(y,(q-1)*3,1))), isoOf(new Date(Date.UTC(y,q*3,0)))]; }
  return [key+'-01-01', key+'-12-31'];
}
function shiftKey(gran,key,dir,yoy){
  const [a]=periodRange(gran,key);
  if(yoy){ const y=+a.slice(0,4)-1; return periodKey(gran, y+a.slice(4)); }
  if(gran==='day') return addDays(key,dir);
  const d=dOf(a);
  if(gran==='month') d.setUTCMonth(d.getUTCMonth()+dir);
  else if(gran==='quarter') d.setUTCMonth(d.getUTCMonth()+3*dir);
  else d.setUTCFullYear(d.getUTCFullYear()+dir);
  return periodKey(gran,isoOf(d));
}
function periodLabel(gran,key){
  if(!key) return '–';
  if(gran==='day') return fmtDay(key);
  if(gran==='month') return MON[+key.slice(5,7)-1]+' '+key.slice(0,4);
  if(gran==='quarter') return key.slice(5)+' '+key.slice(0,4);
  return key;
}
function listPeriods(gran){
  const keys=new Set(S.trips.map(t=>periodKey(gran,t.d)));
  for(const c of S.costs) keys.add(periodKey(gran,costDay(c)));
  return [...keys].sort().reverse();
}
// Customers is guest-only by definition: the Trip type filter is pinned to Guest there (the user's own choice is kept for other modules)
const effTtype = () => S.mod==='customers' ? 'Guest' : (S.f.ttype||'all');
function makeCtx(key){
  const gran=S.f.gran, [a,b]=periodRange(gran,key), [mn,mx]=dataBounds();
  const lo=a>mn?a:mn, hi=b<mx?b:mx;
  const all=S.trips.filter(t=>t.d>=a&&t.d<=b);
  const costs=S.costs.filter(c=>costDay(c)>=a&&costDay(c)<=b);
  const days = mn&&hi>=lo ? diffDays(lo,hi)+1 : 0;
  const svc=S.f.service, tt=effTtype(), typeOk=t=>tt==='all'||t.type===tt;
  const vegasAll=svc==='G'?[]:all.filter(t=>t.src==='V');
  return { key, gran, a, b, lo, hi, days, label:periodLabel(gran,key), costs,
    all, trips: all.filter(t=>(svc==='all'||t.src===svc)&&typeOk(t)),
    vegas: vegasAll.filter(typeOk), vegasAll, typeOk, // VG Cars, Drivers and capacity always use every Vegas trip
    grab: svc==='V'?[]:all.filter(t=>t.src==='G'&&typeOk(t)),
    hasData: all.length>0 || costs.length>0 };
}
function ctxPair(){
  const cur=makeCtx(S.f.period);
  if(S.f.compare==='none') return {cur, cmp:null, cmpLabel:null};
  const k=shiftKey(S.f.gran,S.f.period,-1,S.f.compare==='yoy');
  const cmp=makeCtx(k);
  return {cur, cmp: cmp.hasData?cmp:null, cmpLabel: periodLabel(S.f.gran,k)};
}

/* ======================= analytics helpers ======================= */
const shiftDay = t => { const r=t.drv&&ROSTER_IDX.size?rosterFor(t):null; return r? r.d : (t.dep<CFG.cutoffMin ? addDays(t.d,-1) : t.d); };
const isPool = t => t.src==='V' && !CFG.dedicatedDrivers.includes(t.drv);
const hrs = t => t.dur/60;
// minutes of each trip falling in each clock hour; cb(hourIndexFromTripDate, minutes)
function eachHourOverlap(start, len, cb){
  const end=start+len;
  for(let h=Math.floor(start/60); h*60<end; h++){
    const o=Math.min(end,(h+1)*60)-Math.max(start,h*60);
    if(o>0) cb(h,o);
  }
}
function overlap(s,e,a,b){ return Math.max(0,Math.min(e,b)-Math.max(s,a)); }

// Driver capacity per day: rostered pool drivers' available driving hours on days that have a roster, the estimate on days that don't.
function rosterCapacity(x){
  const estOnDuty=(CFG.teamDrivers-CFG.dedicatedDrivers.length)*(CFG.cycleDays-CFG.offDays)/CFG.cycleDays;
  let duty=0, hrs=0, rosterDays=0;
  if(x.days) for(let d=x.lo; d<=x.hi; d=addDays(d,1)){
    if(ROSTER_DAYS.has(d)){ rosterDays++; const rs=S.roster.filter(r=>r.d===d&&!r.off&&!CFG.dedicatedDrivers.includes(r.drv)); duty+=rs.length; hrs+=sum(rs,availHours); }
    else { duty+=estOnDuty; hrs+=estOnDuty*CFG.maxDrive; }
  }
  return {onDuty:x.days?duty/x.days:estOnDuty, driverCap:x.days?hrs/x.days:estOnDuty*CFG.maxDrive, rosterDays};
}
function capacity(x){
  const vegas=x.vegasAll, days=x.days;
  if(!vegas.length||!days) return null;
  const pool=vegas.filter(isPool), ded=vegas.filter(t=>!isPool(t));
  const rc=rosterCapacity(x), onDuty=rc.onDuty;
  const tripsDay=pool.length/days, avgDur=pool.length?sum(pool,t=>t.dh)/pool.length:0;
  const vehHrs=sum(pool,t=>t.dh)/days, used=vehHrs+tripsDay*CFG.buffer, bufferH=tripsDay*CFG.buffer;
  const driverCap=rc.driverCap, dedHrs=sum(ded,t=>t.dh)/days;
  const vehicleCap=CFG.fleet*CFG.serviceWindow*CFG.targetUtil-dedHrs;
  const binding=Math.min(driverCap,vehicleCap), spare=Math.max(0,binding-used);
  const add=Math.floor(spare/(avgDur+CFG.buffer));
  return {onDuty,tripsDay,avgDur,vehHrs,used,bufferH,driverCap,vehicleCap,binding,constraint:driverCap<=vehicleCap?'Drivers':'Vehicles',spare,add,max:Math.round(tripsDay+add),usedPct:used/binding,
    rosterDays:rc.rosterDays, basis:rc.rosterDays===0?'estimate':rc.rosterDays>=days?'roster':'mixed'};
}
function hourProfile(trips, days, keyFn){ // avg minutes on road per clock hour, plus distinct keys (cars / drivers) per day-hour
  const mins=new Array(24).fill(0), sets=new Map();
  for(const t of trips){
    eachHourOverlap(t.dep,t.dur,(h,o)=>{
      mins[h%24]+=o;
      if(keyFn){ const k=addDays(t.d,Math.floor(h/24))+'|'+(h%24); if(!sets.has(k)) sets.set(k,new Set()); sets.get(k).add(keyFn(t)); }
    });
  }
  const avgOnRoad=mins.map(m=>m/60/days);
  let avgDistinct=null,maxDistinct=null;
  if(keyFn){
    avgDistinct=new Array(24).fill(0); maxDistinct=new Array(24).fill(0);
    for(const [k,set] of sets){ const h=+k.split('|')[1]; avgDistinct[h]+=set.size; maxDistinct[h]=Math.max(maxDistinct[h],set.size); }
    avgDistinct=avgDistinct.map(v=>v/days);
  }
  return {avgOnRoad,avgDistinct,maxDistinct};
}
function trendBuckets(ctx, trips, valFn){
  valFn=valFn||(()=>1);
  const g=ctx.gran; let labels=[], keys=[], keyOf;
  if(g==='day'){ keys=[...Array(24).keys()]; labels=keys.map(h=>pad(h)+':00'); keyOf=t=>Math.floor(t.dep/60); }
  else if(g==='month'){ for(let d=ctx.a; d<=ctx.b; d=addDays(d,1)){ keys.push(d); labels.push(+d.slice(8)); } keyOf=t=>t.d; }
  else if(g==='quarter'){ // weeks starting Monday
    let w=addDays(ctx.a,-weekday(ctx.a)); while(w<=ctx.b){ keys.push(w); labels.push('Wk '+fmtDay(w).slice(0,6)); w=addDays(w,7); }
    keyOf=t=>addDays(t.d,-weekday(t.d));
  } else { for(let m=1;m<=12;m++){ keys.push(ctx.a.slice(0,4)+'-'+pad(m)); labels.push(MON[m-1]); } keyOf=t=>t.d.slice(0,7); }
  const idx=new Map(keys.map((k,i)=>[k,i])), vals=new Array(keys.length).fill(0);
  for(const t of trips){ const i=idx.get(keyOf(t)); if(i!=null) vals[i]+=valFn(t); }
  const unit={day:'hour of day',month:'day',quarter:'week',year:'month'}[g];
  return {labels,vals,unit};
}
function weekdayCount(ctx){ const c=new Array(7).fill(0); if(!ctx.days) return c; const [mn,mx]=dataBounds(); const lo=ctx.a>mn?ctx.a:mn, hi=ctx.b<mx?ctx.b:mx; for(let d=lo; d<=hi; d=addDays(d,1)) c[weekday(d)]++; return c; }
const areaOf = t => CFG.areas.includes(t.city)&&t.city!=='Others'?t.city:'Others'; // any CITY not listed in CFG.areas goes to Others
const levelOf = tier => { if(CFG.bannedTiers.includes(tier)) return 'Banned Guest'; const l=CFG.levels.find(([,ts])=>ts.includes(tier)); return l?l[0]:'Other'; };
const tierGroup = t => t.type==='Internal'?'Internal Group':levelOf(t.tier);
const LEVELS = ['VIP Level','Premium Level','Potential Level','Internal Group'];
const levelColor = (name,P) => name==='Other'||name==='Banned Guest'?P.neutral:name==='Potential Level'?P.potential:P.s[[6,0,-1,4][LEVELS.indexOf(name)]];
const tierColor = (tier,P) => { const i=CFG.levels.findIndex(l=>l[1].includes(tier)); if(i<0) return P.neutral; const sh=P.lv[i], j=CFG.levels[i][1].indexOf(tier); return sh[Math.max(0,sh.length-1-j)]; }; // higher tier = darker shade
const levelNames = trips => LEVELS.concat(['Banned Guest','Other'].filter(n=>trips.some(t=>tierGroup(t)===n)));
const sortTiers = arr => arr.sort((a,b)=>{ const ia=CFG.tierOrder.indexOf(a), ib=CFG.tierOrder.indexOf(b); return (ia<0?99:ia)-(ib<0?99:ib); });
function topN(map, n){ const arr=[...map].sort((a,b)=>b[1]-a[1]); if(arr.length<=n) return arr; const head=arr.slice(0,n); head.push(['Other',sum(arr.slice(n),x=>x[1])]); return head; }

/* ============================ colors ============================ */
function css(v){ return getComputedStyle(document.documentElement).getPropertyValue(v).trim(); }
function palette(){
  return { s:[1,2,3,4,5,6,7,8].map(i=>css('--s'+i)), neutral:css('--neutral-mark'), potential:css('--level-potential'), lv:['--lv-vip','--lv-prem','--lv-pot'].map(v=>css(v).split(',').map(x=>x.trim())), text:css('--text-secondary'), muted:css('--text-muted'), grid:css('--grid'),
    seq:css('--seq').split(',').map(x=>x.trim()), surface:css('--glass-strong'), primary:css('--text-primary') };
}
// fixed colour per entity so a filter never repaints survivors
const ENTITY_COLOR = {};
function colorFor(domain, name, P){
  if(name==='Other'||name==='Others'||name==='In yard') return P.neutral;
  const k=domain+'|'+name; if(ENTITY_COLOR[k]==null){ const used=Object.keys(ENTITY_COLOR).filter(x=>x.startsWith(domain+'|')).length; ENTITY_COLOR[k]=used%8; }
  return P.s[ENTITY_COLOR[k]];
}
function seqColor(v,max,P){ if(!v||!max) return null; const i=Math.min(P.seq.length-1, Math.max(0, Math.round((v/max)*(P.seq.length-1)))); return {bg:P.seq[i], fg: i>=6?'#fff':'#0d1b33'}; }

/* ============================ charts ============================ */
const CHARTS=[];
function destroyCharts(){ while(CHARTS.length) CHARTS.pop().destroy(); }
function baseOptions(P, o){
  o=o||{};
  const valFmt=o.fmt||fmt.int, horizontal=!!o.horizontal;
  const valueAxis={ beginAtZero:true, stacked:!!o.stacked, grid:{color:P.grid, drawTicks:false}, border:{display:false},
    ticks:{color:P.muted, padding:6, callback:v=>o.percent?v+'%':(Number.isInteger(v)&&!o.tickFmt?fmt.int(v):(o.tickFmt||valFmt)(v))}, max:o.percent?100:o.max,
    title: o.yTitle?{display:true,text:o.yTitle,color:P.muted,font:{size:11,weight:'600'}}:undefined };
  const catAxis={ stacked:!!o.stacked, grid:{display:false}, border:{color:P.grid}, ticks:{color:P.muted, autoSkip:true, maxRotation:0, font:{size:11}} };
  return {
    responsive:true, maintainAspectRatio:false, indexAxis: horizontal?'y':'x',
    // stacked charts (owner): the tooltip shows only the segment under the pointer, not every segment of the bar
    interaction:(m=>({mode:m, intersect:m==='point', axis: horizontal?'y':'x'}))(o.mode||(o.stacked?'point':'index')),
    animation:{duration:350},
    plugins:{
      legend:{display:!!o.legend, position:'bottom', labels:{color:P.text, usePointStyle:true, pointStyle:'rectRounded', boxWidth:10, boxHeight:10, padding:14, font:{size:12}}},
      tooltip:{backgroundColor:P.surface, titleColor:P.primary, bodyColor:P.primary, borderColor:P.grid, borderWidth:1, padding:10, cornerRadius:10, usePointStyle:true, boxPadding:4,
        callbacks:{label:c=>' '+(c.dataset.label?c.dataset.label+': ':'')+(o.percent?fmt.d1(c.parsed[horizontal?'x':'y'])+'%':valFmt(c.parsed[horizontal?'x':'y']))+(o.unit?' '+o.unit:'')}}
    },
    scales: horizontal? {x:valueAxis, y:catAxis} : {x:catAxis, y:valueAxis},
  };
}
function barGradient(c, horizontal){ // scriptable fill: 45% tint at the baseline -> full colour at the bar end
  return ctx=>{ const a=ctx.chart.chartArea; if(!a||!/^#[0-9a-f]{6}$/i.test(c)) return c;
    const g=horizontal?ctx.chart.ctx.createLinearGradient(a.left,0,a.right,0):ctx.chart.ctx.createLinearGradient(0,a.bottom,0,a.top);
    g.addColorStop(0,c+'73'); g.addColorStop(1,c); return g; };
}
function mkChart(id, type, labels, datasets, o){
  const el=document.getElementById(id); if(!el) return;
  const P=palette(); o=o||{};
  datasets.forEach((d,i)=>{
    const c=d.color||P.s[i], tp=d.type||type; // a dataset may set its own type (bar + line combo)
    if(tp==='bar'){ Object.assign(d,{backgroundColor:o.stacked?c:barGradient(c,!!o.horizontal), borderColor:'transparent', borderRadius:o.stacked?0:4, borderSkipped:o.stacked?false:'start', maxBarThickness:o.thick||28, categoryPercentage:.8, barPercentage:.9}); if(o.stacked){ d.borderWidth={top:o.horizontal?0:1,right:o.horizontal?1:0,bottom:0,left:0}; d.borderColor=P.surface; } }
    if(tp==='line'){ Object.assign(d,{borderColor:c, backgroundColor:c, borderWidth:2, pointRadius:0, pointHoverRadius:5, pointHoverBorderWidth:2, pointHoverBorderColor:P.surface, tension:.35, cubicInterpolationMode:'monotone', fill:d.area?'origin':false, borderDash:d.dash||[]}); if(d.area){ const g=el.getContext('2d').createLinearGradient(0,0,0,el.parentNode.clientHeight||300); g.addColorStop(0,c+'40'); g.addColorStop(1,c+'00'); d.backgroundColor=g; } if(d.points){ Object.assign(d,{pointRadius:3,pointBackgroundColor:c,pointBorderColor:P.surface,pointBorderWidth:1.5}); } }
  });
  const opts = type==='doughnut' ? {
      responsive:true, maintainAspectRatio:false, cutout:'72%',
      plugins:{legend:{display:false}, tooltip:{backgroundColor:P.surface,titleColor:P.primary,bodyColor:P.primary,borderColor:P.grid,borderWidth:1,padding:10,cornerRadius:10,
        callbacks:{label:c=>{ const tot=sum(c.dataset.data); return ' '+c.label+': '+(o.fmt||fmt.int)(c.parsed)+(o.unit?' '+o.unit:'')+' ('+fmt.pct(c.parsed/tot)+')'; }}}}
    } : baseOptions(P,o);
  if(type==='doughnut') datasets.forEach(d=>{ d.borderColor=P.surface; d.borderWidth=2; d.hoverOffset=4; d.borderRadius=4; });
  const plugins=[];
  if(type==='doughnut'&&o.pctLabels){
    // Labels sit in a column left/right of the ring, spread apart so they never overlap, each tied to its slice by a thin leader line.
    // Narrow charts stack 'Name' over '00.0%' so they need less width.
    const named=o.pctLabels==='name';
    opts.layout={padding:named?(c=>c.chart.width<440?{left:62,right:62,top:24,bottom:24}:{left:96,right:96,top:22,bottom:22}):34};
    plugins.push({id:'pctLabels',afterDatasetsDraw(ch){
      const ctx=ch.ctx, ds=ch.data.datasets[0], tot=sum(ds.data), narrow=named&&ch.width<440;
      if(!tot) return;
      const meta=ch.getDatasetMeta(0).data, gap=narrow?26:15, top=6, bottom=ch.height-6;
      const items=[];
      meta.forEach((arc,i)=>{ const v=ds.data[i]; if(!v||v/tot<0.015) return; // slivers under 1.5%: tooltip only
        const a=(arc.startAngle+arc.endAngle)/2, R=arc.outerRadius;
        items.push({i,v,a,cx:arc.x,cy:arc.y,R,side:Math.cos(a)>=0?1:-1,y:arc.y+Math.sin(a)*(R+14)}); });
      for(const side of [1,-1]){
        const col=items.filter(t=>t.side===side).sort((p,q)=>p.y-q.y);
        for(let k=0;k<col.length;k++){ const lo=k?col[k-1].y+gap:top+gap/2; if(col[k].y<lo) col[k].y=lo; }      // push down
        for(let k=col.length-1;k>=0;k--){ const hi=k<col.length-1?col[k+1].y-gap:bottom-gap/2; if(col[k].y>hi) col[k].y=hi; } // pull up if past the bottom
      }
      ctx.save(); ctx.font=(named?'700 11px':'700 12px')+' "Plus Jakarta Sans",system-ui,sans-serif'; ctx.textBaseline='middle';
      for(const t of items){
        const ex=t.cx+Math.cos(t.a)*(t.R+2), ey=t.cy+Math.sin(t.a)*(t.R+2);           // on the ring edge
        const kx=t.cx+Math.cos(t.a)*(t.R+10), ky=t.cy+Math.sin(t.a)*(t.R+10);         // elbow
        const lx=t.cx+t.side*(t.R+18);                                                 // label column
        ctx.strokeStyle=P.muted; ctx.globalAlpha=.55; ctx.lineWidth=1;
        ctx.beginPath(); ctx.moveTo(ex,ey); ctx.lineTo(kx,ky); ctx.lineTo(lx-t.side*3,t.y); ctx.stroke(); ctx.globalAlpha=1;
        ctx.fillStyle=P.primary; ctx.textAlign=t.side>0?'left':'right';
        const pct=fmt.pct(t.v/tot), name=named?ch.data.labels[t.i]:'';
        if(narrow){ ctx.fillText(name,lx,t.y-6.5); ctx.fillText(pct,lx,t.y+6.5); }
        else ctx.fillText((name?name+' ':'')+pct,lx,t.y);
      }
      ctx.restore(); }});
  }
  CHARTS.push(new Chart(el,{type,data:{labels,datasets},options:opts,plugins}));
}

// Guest trips by Nationality: top 9 nationalities + Other, as share of guest trips (shared by Overview and Customers)
// donut: top 7 nationalities + Other (8 slices = the categorical palette), largest first, clockwise.
// Wide cards label the slices around the ring; narrow cards (list:true) show a clean ring plus a legend list underneath.
function natChart(id, guests, P, opt){
  const nat=topN(countBy(guests,t=>t.nat||'Unknown'),7), tot=guests.length, list=opt&&opt.list;
  mkChart(id,'doughnut',nat.map(x=>x[0]),[{data:nat.map(x=>x[1]),backgroundColor:nat.map(x=>colorFor('nat',x[0],P))}],list?{unit:'guest trips'}:{unit:'guest trips',pctLabels:'name'});
  const el=document.getElementById(id+'-legend');
  if(el) el.innerHTML=nat.map(([n,v])=>`<div class="lg-row"><i class="dot" style="background:${colorFor('nat',n,P)}"></i><span class="lg-name">${esc(n)}</span><span class="lg-pct num">${fmt.pct(tot?v/tot:null)}</span><span class="lg-n num">${fmt.int(v)}</span></div>`).join('');
}
// Trips by main area (city) — shared by Overview (all trips) and Customers (guest trips)
function areaChart(id, trips, P){
  const ar=countBy(trips,areaOf);
  mkChart(id,'bar',CFG.areas,[{label:'Trips',data:CFG.areas.map(a=>ar.get(a)||0)}],{horizontal:true,unit:'trips',thick:40});
}
// Value that reveals the guests behind it on hover/tap, as a small table (top 15 by trips)
function hovList(valueHtml, trips, title){
  if(!trips||!trips.length) return `<span class="v">${valueHtml}</span>`;
  const rows=[...groupBy(trips,t=>t.gid)].map(([gid,ts])=>({gid,n:ts.length,tier:ts[0].tier||'–',nat:ts[0].nat||'–'})).sort((a,b)=>b.n-a.n);
  const body=rows.slice(0,15).map(r=>`<tr><td>${esc(r.gid)}</td><td>${esc(r.tier)}</td><td>${esc(r.nat)}</td><td class="r">${fmt.int(r.n)}</td></tr>`).join('');
  const html=`<div class="tt-h"><b>${esc(title)}</b><span>${fmt.int(rows.length)} guest${rows.length>1?'s':''} · ${fmt.int(trips.length)} trips</span></div>
    <table class="tt"><thead><tr><th>Guest ID</th><th>Tier</th><th>Nationality</th><th class="r">Trips</th></tr></thead><tbody>${body}</tbody></table>
    ${rows.length>15?`<div class="tt-f">Top 15 shown · ${fmt.int(rows.length-15)} more guests</div>`:''}`;
  return `<span class="v hov" tabindex="0" data-tiphtml="${esc(html)}">${valueHtml}</span>`;
}
const natBody = (id, guests, h=330, list=false) => `<div style="height:${h}px;position:relative">${canvas(id).replace('class="chart "',`class="chart" style="height:${h}px"`)}
  <div style="position:absolute;inset:0;display:grid;place-items:center;pointer-events:none;text-align:center"><div><div class="big num" style="font-size:26px">${fmt.int(guests)}</div><div class="muted" style="font-size:11px">guest trips</div></div></div></div>${list?`<div class="legend-list" id="${id}-legend"></div>`:''}`;

/* ============================ UI atoms ============================ */
const ICON = {
  home:'<path d="M3 11l9-8 9 8"/><path d="M5 10v10h14V10"/>',
  car:'<path d="M5 17h14v-5l-2-5H7l-2 5z"/><circle cx="7.5" cy="17.5" r="1.5"/><circle cx="16.5" cy="17.5" r="1.5"/><path d="M5 12h14"/>',
  users:'<circle cx="9" cy="8" r="3.2"/><path d="M3 20c0-3.3 2.7-6 6-6s6 2.7 6 6"/><circle cx="17" cy="9" r="2.5"/><path d="M15.5 14.2c3 .2 5.5 2.6 5.5 5.8"/>',
  wheel:'<circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="2.2"/><path d="M12 14.2V21M9.9 11.4 3.4 9.8M14.1 11.4l6.5-1.6"/>',
  wrench:'<path d="M14.7 6.3a4 4 0 0 0-5.4 5.4L3 18l3 3 6.3-6.3a4 4 0 0 0 5.4-5.4l-2.5 2.5-2.5-.5-.5-2.5z"/>',
  phone:'<rect x="6" y="2.5" width="12" height="19" rx="2.5"/><path d="M10 18h4"/><path d="M9 8.5h6l-1 3h-4z"/>',
  info:'<circle cx="12" cy="12" r="9"/><path d="M12 11v6"/><path d="M12 7.5v.5"/>',
  admin:'<path d="M4 6h9M17 6h3M4 12h3M11 12h9M4 18h11M19 18h1"/><circle cx="15" cy="6" r="2"/><circle cx="9" cy="12" r="2"/><circle cx="17" cy="18" r="2"/>',
  route:'<circle cx="6" cy="19" r="2"/><circle cx="18" cy="5" r="2"/><path d="M8 19h8a3 3 0 0 0 0-6H8a3 3 0 0 1 0-6h8"/>',
  clock:'<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
  chart:'<path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/>',
  gauge:'<path d="M4 18a8 8 0 1 1 16 0"/><path d="M12 18l4-6"/>',
  star:'<path d="M12 3l2.7 5.6 6.1.9-4.4 4.3 1 6.1L12 17l-5.4 2.9 1-6.1L3.2 9.5l6.1-.9z"/>',
  money:'<rect x="2.5" y="6" width="19" height="12" rx="2"/><circle cx="12" cy="12" r="2.6"/>',
  globe:'<circle cx="12" cy="12" r="9"/><path d="M3 12h18M12 3c3 3.5 3 14.5 0 18M12 3c-3 3.5-3 14.5 0 18"/>',
  plus:'<path d="M12 5v14M5 12h14"/>',
  calendar:'<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/>',
};
const svg = (n,sw) => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="${sw||2}" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICON[n]}</svg>`;

function deltaHtml(cur, prev, pc){
  if(pc.cmp==null) return pc.cmpLabel? `<span class="muted">No data for ${esc(pc.cmpLabel)}</span>` : '';
  if(prev==null||!isFinite(prev)||cur==null||!isFinite(cur)) return `<span class="muted">vs ${esc(pc.cmpLabel)}: –</span>`;
  if(prev===0) return `<span class="muted">vs ${esc(pc.cmpLabel)}: new</span>`;
  const ch=(cur-prev)/Math.abs(prev), arrow=ch>0.0005?'↑':ch<-0.0005?'↓':'→';
  return `<b>${arrow} ${fmt.pct(Math.abs(ch))}</b> vs ${esc(pc.cmpLabel)}`;
}
// KPI definitions are evaluated on current AND comparison context -> one code path for every module
function kpiRow(defs, pc, cls){
  return `<section class="grid kpis ${cls||''}">${defs.map((k,i)=>{
    const v=k.get(pc.cur), p=pc.cmp?k.get(pc.cmp):null;
    return `<div class="glass kpi"><div class="chip" style="background:var(--chip-${(i%4)+1})">${svg(k.icon)}</div>
      <div class="k-body"><div class="k-label">${esc(k.label)}</div>
      <div class="k-val num${typeof v==='string'?' txt':''}">${(k.fmt||fmt.int)(v)}${k.unit?`<small>${k.unit}</small>`:''}</div>
      <div class="k-delta">${k.noCompare?'':deltaHtml(v,p,pc)}</div>${k.note?`<div class="k-note">${k.note(pc.cur)}</div>`:''}</div></div>`;
  }).join('')}</section>`;
}
function card(title, desc, body, opts){
  opts=opts||{};
  return `<div class="glass card ${opts.cls||''}" ${opts.id?`id="${opts.id}"`:''}><div class="card-h"><div><h3>${esc(title)}</h3>${desc?`<div class="desc">${desc}</div>`:''}</div>${opts.right?`<div class="card-r">${opts.right}</div>`:''}</div>${body}</div>`;
}
const canvas = (id,cls) => `<div class="chart ${cls||''}"><canvas id="${id}" role="img" aria-label="chart"></canvas></div>`;
function table(cols, rows, opts){
  opts=opts||{};
  const th=cols.map(c=>`<th class="${c.r?'r':''}">${esc(c.label)}</th>`).join('');
  const tr=rows.map(r=>`<tr class="${opts.flag&&opts.flag(r)?'flag':''} ${opts.rowCls?opts.rowCls(r):''}">${cols.map(c=>`<td class="${c.r?'r':''}">${c.html?c.html(r):esc(c.f?c.f(r[c.k],r):r[c.k])}</td>`).join('')}</tr>`).join('');
  const tf=opts.foot?`<tfoot><tr>${cols.map(c=>`<td class="${c.r?'r':''}">${opts.foot[c.k]!=null?esc(c.f?c.f(opts.foot[c.k],opts.foot):opts.foot[c.k]):''}</td>`).join('')}</tr></tfoot>`:'';
  return `<div class="tbl-wrap" style="${opts.maxH?`max-height:${opts.maxH}px`:''}"><table class="tbl"><thead><tr>${th}</tr></thead><tbody>${tr||`<tr><td colspan="${cols.length}" class="muted">No data</td></tr>`}</tbody>${tf}</table></div>`;
}
function heatmap(rowLabels, colLabels, m, f, unit){
  const P=palette(); const max=Math.max(0,...m.flat());
  const head=`<tr><th></th>${colLabels.map(c=>`<th>${esc(c)}</th>`).join('')}</tr>`;
  const body=rowLabels.map((r,i)=>`<tr><th class="rl">${esc(r)}</th>${m[i].map((v,j)=>{ const c=seqColor(v,max,P);
    const tip=`${esc(r)} · ${esc(colLabels[j])}: ${f(v)} ${unit}`;
    return `<td data-tip="${tip}" aria-label="${tip}" style="background:${c?c.bg:'var(--heat-empty)'}"></td>`; } /* colour only; value on hover */).join('')}</tr>`).join('');
  return `<div class="tbl-wrap"><table class="heat">${head}${body}</table></div>
    <div class="legend-scale">0 <span class="bar" style="background:linear-gradient(90deg,${P.seq.join(',')})"></span> ${f(max)} ${esc(unit)}</div>`;
}
function emptyModule(icon, title, text, extra){
  return `<div class="glass empty"><div class="ico">${svg(icon)}</div><h3>${esc(title)}</h3><p>${text}</p>${extra||''}</div>`;
}
function allTripsNote(){ return (S.f.ttype||'all')!=='all' ? `<div style="margin:-4px 0 8px"><span class="badge">Uses all trips (guest + internal)</span></div>` : ''; }
function vegasOnlyNote(){ return (S.f.service==='all' ? `<span class="badge">Vegas Car only</span>` : '') + ((S.f.ttype||'all')!=='all' ? ` <span class="badge">All trip types (filter not applied)</span>` : ''); }

/* ============================ MODULES ============================ */
let MODULES = [];
function buildModules(){ MODULES = [
  {id:'overview', name:'Overview', icon:'home', render:renderOverview},
  {id:'customers', name:'Customers', icon:'users', render:renderCustomers},
  {id:'vehicles', name:'VG Cars', icon:'car', render:renderVehicles},
  {id:'drivers', name:'Drivers', icon:'wheel', render:renderDrivers},
  {id:'carcost', name:'Car Cost', icon:'wrench', render:renderCarCost},
  ...(CFG.features.grab?[{id:'grabcost', name:'Grab Cost', icon:'phone', render:renderGrabCost}]:[]),
  {id:'about', name:'Data & Definitions', icon:'info', render:renderAbout},
]; }
buildModules();

/* ---------- Overview ---------- */
function renderOverview(pc, after){
  const c=pc.cur;
  const guestsOf=x=>uniq(x.trips.filter(t=>t.type==='Guest').map(t=>t.gid)).length;
  // cost KPIs take the 3rd and 7th tiles (total cost, cost per km) (owner request); Total cost = Vegas Car cost (+ Grab fares when Grab is switched on)
  const costTotalDef = CFG.features.grab ? {label:'Total cost', icon:'money', fmt:fmt.usd, get:x=>{const v=S.f.service==='G'||!x.costs.length?null:sum(x.costs,r=>r.amt), g=x.grab.filter(t=>t.cost!=null), gc=g.length?sum(g,t=>t.cost):null; return v==null&&gc==null?null:(v||0)+(gc||0);}, note:()=>'Vegas Car + Grab'} :
    {label:'Total cost', icon:'money', fmt:fmt.usd, get:x=>x.costs.length?sum(x.costs,r=>r.amt):null, note:x=>x.costs.length?'Vegas Car fleet cost':'no car-cost data yet'};
  // owner request: the 7th tile shows cost per km (cost ÷ km driven by the trips that carry that cost)
  const costPerKmDef = CFG.features.grab ? {label:'Cost per km', icon:'money', fmt:fmt.usd, get:x=>{const v=S.f.service==='G'||!x.costs.length?null:sum(x.costs,r=>r.amt), g=x.grab.filter(t=>t.cost!=null), gc=g.length?sum(g,t=>t.cost):null;
      if(v==null&&gc==null) return null; const km=(v!=null?sum(x.vegasAll,t=>t.km):0)+(gc!=null?sum(x.grab,t=>t.km):0); return km?((v||0)+(gc||0))/km:null;}, note:()=>'total cost ÷ km driven'} :
    {label:'Cost per km', icon:'money', fmt:fmt.usd, get:x=>{const km=sum(x.vegasAll,t=>t.km); return x.costs.length&&km?sum(x.costs,r=>r.amt)/km:null;}, note:()=>'car cost ÷ km driven by Vegas Car'};
  const defs=[
    {label:'Total trips', icon:'route', get:x=>x.trips.length, note:x=>`${fmt.int(x.trips.filter(t=>t.type==='Guest').length)} guest · ${fmt.int(x.trips.filter(t=>t.type==='Internal').length)} internal`},
    {label:'Guests served', icon:'users', get:guestsOf, note:()=>'distinct guest IDs'},
    costTotalDef,
    {label:'Avg trips / day', icon:'route', fmt:fmt.d1, get:x=>x.days?x.trips.length/x.days:null, note:x=>`${x.days} day${x.days===1?'':'s'} with data`},
    {label:'Total distance', icon:'globe', unit:'km', get:x=>sum(x.trips,t=>t.km)},
    {label:'Total trip hours', icon:'clock', unit:'h', get:x=>sum(x.trips,hrs), note:()=>'incl. empty return to base'},
    costPerKmDef,
    {label:'Avg km / trip', icon:'globe', fmt:fmt.d1, unit:'km', get:x=>x.trips.length?sum(x.trips,t=>t.km)/x.trips.length:null},
  ];
  const cap=capacity(c);
  const grp=countBy(c.trips,tierGroup), grpNames=levelNames(c.trips);
  const P=palette();
  const svc=countBy(c.all,t=>t.src);
  const capCard = cap ? card('Vegas Car capacity',`Driver pool capacity (excl. boss driver) · ${cap.basis==='roster'?'actual driver roster':cap.basis==='mixed'?`actual roster on ${cap.rosterDays} of ${c.days} days`:'estimated roster'}`, `${allTripsNote()}
      <div style="display:flex;align-items:baseline;gap:10px"><span class="big num">${fmt.pct(cap.usedPct)}</span><span class="muted">capacity used</span></div>
      <div class="meter" style="margin:10px 0 6px" title="Used ${fmt.d1(cap.used)}h · spare ${fmt.d1(cap.spare)}h per day">
        <span style="width:${cap.used/cap.binding*100}%;background:${P.s[0]}"></span><span style="flex:1;background:${P.neutral}"></span></div>
      <div class="stat-row"><span class="l"><i class="dot" style="background:${P.s[0]}"></i>Hours used / day</span><span class="v num">${fmt.d1(cap.used)} h</span></div>
      <div class="stat-row"><span class="l"><i class="dot" style="background:${P.neutral}"></i>Spare capacity / day</span><span class="v num">${fmt.d1(cap.spare)} h</span></div>
      <div class="stat-row"><span class="l">Additional trips possible / day</span><span class="v num">+${fmt.int(cap.add)}</span></div>
      <div class="stat-row"><span class="l">Max sustainable trips / day</span><span class="v num">≈ ${fmt.int(cap.max)}</span></div>
      <div class="stat-row"><span class="l">Binding constraint</span><span class="v">${cap.constraint}</span></div>`)
    : card('Vegas Car capacity','', `<div class="ph">No Vegas Car trips in this selection</div>`);
  const mixCard = card('Trip mix', 'Service and customer group', `
      <div style="height:270px;position:relative">${canvas('ovDonut','short').replace('class="chart short"','class="chart" style="height:270px"')}
        <div style="position:absolute;inset:0;display:grid;place-items:center;pointer-events:none;text-align:center"><div><div class="big num" style="font-size:26px">${fmt.int(c.trips.length)}</div><div class="muted" style="font-size:11px">trips</div></div></div></div>
      ${grpNames.map(n=>`<div class="stat-row"><span class="l"><i class="dot" style="background:${levelColor(n,P)}"></i>${n}<span class="muted" style="font-size:11px">${n==='Internal Group'||n==='Other'?'':n==='Banned Guest'?CFG.bannedTiers.join(', '):(CFG.levels.find(l=>l[0]===n)[1].join(', '))}</span></span><span class="v num">${fmt.int(grp.get(n)||0)} <span style="display:inline-block;min-width:52px;text-align:right;color:var(--accent)">${fmt.pct(c.trips.length?(grp.get(n)||0)/c.trips.length:null)}</span></span></div>`).join('')}
      ${CFG.features.grab?`<div class="stat-row"><span class="l">Vegas Car / Grab trips</span><span class="v num">${fmt.int(svc.get('V')||0)} / ${S.trips.some(t=>t.src==='G')?fmt.int(svc.get('G')||0):'<span class="muted">no Grab data</span>'}</span></div>`:''}`);

  const html = kpiRow(defs,pc) + `
    <section class="grid g-2-1">
      ${card('Trips over time', `Trips per ${trendBuckets(c,[]).unit} · ${esc(c.label)}`, canvas('ovTrend','tall'), {cls:'fill', right:`<span class="badge">${esc(S.f.service==='all'?'All services':S.f.service==='V'?'Vegas Car':'Grab')}</span>`})}
      <div class="stack">${mixCard}</div>
    </section>
    <section class="grid g-3">
      ${card('Weekly activity','Avg trips per weekday', canvas('ovWeek'))}
      ${card('Departures by hour','Avg trips starting in each clock hour, per day', canvas('ovHour'))}
      ${capCard}
    </section>
    <section class="grid g-2">
      ${card('Trips by area','Main areas by city · Ho Chi Minh includes airport trips · Others = OutSide, Long An, Tay Ninh, Vung Tau…', canvas('ovDest','tall'))}
      ${card('Guest trips by Nationality','Share of guest trips by nationality', natBody('ovNat',c.trips.filter(t=>t.type==='Guest').length))}
    </section>`;
  after.push(()=>{
    const tb=trendBuckets(c,c.trips), P=palette();
    const pts=tb.labels.length<=31; // markers on daily/weekly/monthly points; hourly view stays a clean line
    const datasets=[{label:'Trips',data:tb.vals,color:P.s[6],area:true,points:pts}];
    if(S.f.service==='all' && c.grab.length){ const v=trendBuckets(c,c.vegas), g=trendBuckets(c,c.grab); datasets.splice(0,1,{label:'Vegas Car',data:v.vals,color:P.s[0],area:true,points:pts},{label:'Grab',data:g.vals,color:P.s[2],area:true,points:pts}); }
    mkChart('ovTrend','line',tb.labels,datasets,{legend:datasets.length>1,unit:'trips'});
    // donut: one slice per tier (shades of its level colour); the table below stays grouped by level
    const tierKey=t=>t.type==='Internal'?'Internal':(t.tier||'Unknown'), tc=countBy(c.trips,tierKey);
    const known=CFG.tierRank;
    const keys=[...known,'Internal',...[...tc.keys()].filter(k=>k!=='Internal'&&!known.includes(k))].filter(k=>tc.get(k));
    mkChart('ovDonut','doughnut',keys,[{data:keys.map(k=>tc.get(k)),backgroundColor:keys.map(k=>k==='Internal'?levelColor('Internal Group',P):tierColor(k,P))}],{unit:'trips',pctLabels:'name'});
    const wc=weekdayCount(c), wt=new Array(7).fill(0); c.trips.forEach(t=>wt[weekday(t.d)]++);
    mkChart('ovWeek','line',WD,[{label:'Avg trips',data:wt.map((v,i)=>wc[i]?v/wc[i]:0),color:P.s[6],area:true,points:true}],{fmt:fmt.d1,unit:'trips/day'});
    const hd=new Array(24).fill(0); c.trips.forEach(t=>hd[Math.floor(t.dep/60)%24]++);
    mkChart('ovHour','line',[...Array(24).keys()].map(h=>pad(h)),[{label:'Departures / day',data:hd.map(v=>c.days?v/c.days:0),color:P.s[6]}],{fmt:fmt.d1,unit:'trips'});
    areaChart('ovDest',c.trips,P);
    natChart('ovNat',c.trips.filter(t=>t.type==='Guest'),P);
  });
  return html;
}

/* ---------- Vehicles ---------- */
// guest fleet = every Vegas car except retired cars, internal-use cars and unnamed rows
const inFleet = car => car!=='Others'&&car!=='Unknown'&&!CFG.retiredCars.includes(car)&&!CFG.internalCars.includes(car);
const carTag = car => CFG.retiredCars.includes(car)?'retired':CFG.internalCars.includes(car)?'internal use':'';
function vehicleStats(ctx){
  const by=groupBy(ctx.vegasAll,t=>t.car);
  return [...by].map(([car,ts])=>{
    const perDay=groupBy(ts,t=>t.d), h=sum(ts,hrs);
    const maxDay=Math.max(...[...perDay.values()].map(a=>sum(a,hrs)));
    const avgDay=ctx.days?h/ctx.days:0;
    return {car, trips:ts.length, hours:h, km:sum(ts,t=>t.km), avgMin:sum(ts,t=>t.dur)/ts.length, avgKm:sum(ts,t=>t.km)/ts.length,
      daysUsed:perDay.size, avgDay, maxDay, util:avgDay/CFG.serviceWindow, yard:Math.max(0,24-avgDay),
      mainTier: topKey(countBy(ts.filter(t=>t.type==='Guest'),t=>t.tier||'–'))||'–', mainDest: topKey(countBy(ts,t=>t.city))};
  }).sort((a,b)=>b.trips-a.trips);
}
function renderVehicles(pc, after){
  if(S.f.service==='G') return emptyModule('car','VG Cars analysis covers Vegas Car only','Grab trips do not use the Vegas fleet. Switch the service filter to <b>All</b> or <b>Vegas Car</b>.');
  const c=pc.cur; if(!c.vegasAll.length) return emptyModule('car','No Vegas Car trips in this period','Choose another period or upload data.');
  const fleetTrips=x=>x.vegasAll.filter(t=>inFleet(t.car)), activeCars=x=>uniq(fleetTrips(x).map(t=>t.car)).length;
  const peak=x=>{ if(!x.vegasAll.length) return null; const hp=hourProfile(x.vegasAll,x.days,t=>t.car); return Math.max(...hp.avgDistinct); };
  const defs=[
    {label:'Active vehicles', icon:'car', get:activeCars, note:()=>`guest fleet · model: ${CFG.fleet} cars`},
    {label:'Hours on the road', icon:'clock', unit:'h', get:x=>sum(x.vegasAll,hrs)},
    {label:'Avg hours / vehicle / day', icon:'gauge', fmt:fmt.d1, unit:'h', get:x=>{const n=activeCars(x); return n&&x.days?sum(fleetTrips(x),hrs)/n/x.days:null;}, note:()=>'guest fleet'},
    {label:'Fleet utilization', icon:'chart', fmt:fmt.pct, get:x=>{const n=activeCars(x); return n&&x.days?sum(fleetTrips(x),hrs)/n/x.days/CFG.serviceWindow:null;}, note:()=>`guest fleet · share of ${CFG.serviceWindow}h service window`},
    {label:'Peak cars in use', icon:'star', fmt:fmt.d1, get:peak, note:()=>'avg distinct cars, busiest hour'},
    {label:'Total km', icon:'globe', unit:'km', get:x=>sum(x.vegasAll,t=>t.km)},
  ];
  const vs=vehicleStats(c);
  const cars=vs.map(v=>v.car);
  const html = `<div class="section-title">VG Cars ${vegasOnlyNote()}</div>` + kpiRow(defs,pc,'six') + `
    ${card('Vehicle scorecard','Per vehicle in '+esc(c.label)+' · in-yard = 24h − avg hours on the road per day', table([
      {k:'car',label:'Vehicle',html:r=>r.car==='Total'?'Total':esc(r.car)+(carTag(r.car)?` <span class="muted" style="font-size:11px">· ${carTag(r.car)}</span>`:'')},{k:'trips',label:'Trips',r:1,f:fmt.int},{k:'hours',label:'Hours',r:1,f:fmt.d1},{k:'km',label:'Km',r:1,f:fmt.int},
      {k:'avgMin',label:'Avg min/trip',r:1,f:fmt.int},{k:'avgKm',label:'Avg km/trip',r:1,f:fmt.d1},{k:'daysUsed',label:'Days used',r:1,f:fmt.int},
      {k:'avgDay',label:'Avg h/day',r:1,f:fmt.d1},{k:'maxDay',label:'Max h in a day',r:1,f:fmt.d1},{k:'util',label:'Util. (16h)',r:1,f:fmt.pct},
      {k:'yard',label:'In yard h/day',r:1,f:fmt.d1},{k:'mainTier',label:'Top tier'},{k:'mainDest',label:'Top city / province'}], vs,
      {foot:{car:'Total',trips:sum(vs,v=>v.trips),hours:sum(vs,v=>v.hours),km:sum(vs,v=>v.km)}}))}
    <section class="grid g-2">
      ${card('On the road vs in the yard','Average hours per day (24h)', canvas('vhRoad','tall'))}
      ${card('Where each vehicle goes','Share of trips by main area', canvas('vhArea','tall'))}
    </section>
    <section class="grid g-2">
      ${card('Customer levels by vehicle','Share of trips by customer level', canvas('vhTier','tall'))}
      ${card('Cars in use by hour of day','Distinct cars on the road · avg vs busiest day', canvas('vhHour','tall'))}
    </section>
    ${card('Hours on the road by time slot','Total hours per vehicle in each clock hour · '+esc(c.label), `<div id="vhHeat"></div>`)}`;
  after.push(()=>{
    const P=palette();
    mkChart('vhRoad','bar',cars,[{label:'On the road',data:vs.map(v=>v.avgDay),color:P.s[0]},{label:'In yard',data:vs.map(v=>v.yard),color:P.neutral}],{horizontal:true,stacked:true,legend:true,fmt:fmt.d1,unit:'h/day',max:24});
    const pctData=(keys,fn)=>keys.map(k=>({label:k,color:colorFor('area',k,P),data:cars.map(car=>{const ts=c.vegasAll.filter(t=>t.car===car); return ts.length?ts.filter(t=>fn(t)===k).length/ts.length*100:0;})}));
    mkChart('vhArea','bar',cars,pctData(CFG.areas,areaOf),{horizontal:true,stacked:true,percent:true,legend:true});
    const groups=levelNames(c.vegasAll);
    const gd=groups.map(g=>({label:g,color:levelColor(g,P),data:cars.map(car=>{const ts=c.vegasAll.filter(t=>t.car===car); return ts.length?ts.filter(t=>tierGroup(t)===g).length/ts.length*100:0;})}));
    mkChart('vhTier','bar',cars,gd,{horizontal:true,stacked:true,percent:true,legend:true});
    const hp=hourProfile(c.vegasAll,c.days,t=>t.car);
    mkChart('vhHour','line',[...Array(24).keys()].map(h=>pad(h)+':00'),[{label:'Avg distinct cars',data:hp.avgDistinct,color:P.s[0]},{label:'Max on busiest day',data:hp.maxDistinct,color:P.s[1],dash:[5,4]},{label:'Practical ceiling (fleet × 65%)',data:new Array(24).fill(CFG.fleet*CFG.targetUtil),color:P.muted,dash:[2,3]}],{legend:true,fmt:fmt.d1,unit:'cars'});
    const m=cars.map(car=>{ const row=new Array(24).fill(0); c.vegasAll.filter(t=>t.car===car).forEach(t=>eachHourOverlap(t.dep,t.dur,(h,o)=>row[h%24]+=o/60)); return row; });
    $('#vhHeat').innerHTML=heatmap(cars,[...Array(24).keys()].map(pad),m,fmt.d1,'h');
  });
  return html;
}

/* ---------- Customers ---------- */
const TOPN_OPTS=[5,10,20,50,100];
const topN_ = (key,def) => { const v=S.extra[key]; return v==='all'?Infinity:(+v||def); };
const topNSel = (key,def) => { const cur=S.extra[key]||String(def);
  return `<select class="sel" data-extra="${key}" aria-label="Number of rows">${TOPN_OPTS.map(n=>`<option value="${n}" ${cur===String(n)?'selected':''}>Top ${n}</option>`).join('')}<option value="all" ${cur==='all'?'selected':''}>All</option></select>`; };
const topTitle = (key,def,what) => { const n=topN_(key,def); return (n===Infinity?'All ':'Top '+n+' ')+what; };
function guestTable(rows, hide){ // hide: column keys left out of this table
  return table([
    {k:'rank',label:'#',r:1},{k:'gid',label:'Guest ID'},{k:'tier',label:'Tier'},{k:'nat',label:'Nationality'},
    {k:'trips',label:'Guest Trips',r:1,f:fmt.int},{k:'km',label:'Total KM',r:1,f:fmt.int},{k:'hours',label:'Total Hours',r:1,f:fmt.d1},{k:'car',label:'Most used vehicle'},{k:'dest',label:'Top city / province'}].filter(c=>!(hide||[]).includes(c.k)), rows,{maxH:520});
}
function guestRows(trips, by, n){
  const g=groupBy(trips,t=>t.gid);
  const rows=[...g].map(([gid,ts])=>({gid, tier:topKey(countBy(ts,t=>t.tier||'–')), nat:topKey(countBy(ts,t=>t.nat||'–')), trips:ts.length, km:sum(ts,t=>t.km), hours:sum(ts,hrs), car:topKey(countBy(ts,t=>t.car)), dest:topKey(countBy(ts,t=>t.city))}));
  rows.sort((a,b)=>b[by]-a[by]||b.trips-a.trips); return rows.slice(0,n).map((r,i)=>Object.assign(r,{rank:i+1}));
}
// Host team = team that booked the trip. Guest trips on Vegas Car only.
// team of a trip: real team (Host Team column, or an Arrangby value that is a team name), else a stable random team (placeholder)
const realTeamOf = t => { const v=(t.team||'').trim(), h=(t.host||'').trim();
  const hit=CFG.hostTeams.find(x=>x.toLowerCase()===v.toLowerCase())||CFG.hostTeams.find(x=>x.toLowerCase()===h.toLowerCase()); return hit||null; };
const SIM_TEAMS = CFG.hostTeams.filter(x=>x!=='Non-Assigned');
const simTeam = t => { let h=2166136261; for(const ch of t.gid+'|'+t.d+'|'+t.dep+'|'+t.car) h=Math.imul(h^ch.charCodeAt(0),16777619); return SIM_TEAMS[(h>>>0)%SIM_TEAMS.length]; };
const teamOf = t => realTeamOf(t) || simTeam(t);
function hostSection(g){
  const vg=g.filter(t=>t.src==='V');
  if(!vg.length) return '';
  const placeholder=vg.some(t=>!realTeamOf(t));
  const top2=ts=>[...countBy(ts,t=>t.tier||'–')].sort((a,b)=>b[1]-a[1]).slice(0,2).map(([k,n])=>`${k} (${fmt.int(n)})`).join(', ');
  const stats=ts=>({trips:ts.length, guests:uniq(ts.map(t=>t.gid)).length, km:sum(ts,t=>t.km), hours:sum(ts,hrs), tiers:top2(ts)});
  const byTeam=groupBy(vg,teamOf);
  const rows=CFG.hostTeams.filter(tm=>byTeam.has(tm)).map(team=>Object.assign({team},stats(byTeam.get(team))));
  const foot=Object.assign({team:'Total'},stats(vg));
  let sel=S.extra.custHost||'__all';
  if(!(sel==='__all'||(sel.startsWith('team:')&&byTeam.has(sel.slice(5))))) sel='__all';
  const selTrips=sel==='__all'?vg:byTeam.get(sel.slice(5));
  const selName=sel==='__all'?'all host teams':'team '+sel.slice(5);
  const note=placeholder?`<span class="badge" title="${esc(HOST_PLACEHOLDER_NOTE)}">Placeholder data</span>`:'';
  const pnote=placeholder?`<p class="muted" style="font-size:11.5px;margin:10px 0 0">⚠ ${esc(HOST_PLACEHOLDER_NOTE)}</p>`:'';
  const options=`<option value="__all" ${sel==='__all'?'selected':''}>All host teams · ${fmt.int(vg.length)} trips</option>`+
    rows.map(r=>`<option value="team:${esc(r.team)}" ${sel==='team:'+r.team?'selected':''}>Team ${esc(r.team)} · ${fmt.int(r.trips)} trips</option>`).join('');
  return `
    ${card('Trips Booked by Host Team',placeholder?'Vegas Car guest trips by host team (simulated)':'Vegas Car guest trips booked by each host team', table([
      {k:'team',label:'Host team'},{k:'trips',label:'Guest Trips',r:1,f:fmt.int},{k:'guests',label:'Guests',r:1,f:fmt.int},
      {k:'km',label:'Total KM',r:1,f:fmt.int},{k:'hours',label:'Total Hours',r:1,f:fmt.d1},{k:'tiers',label:'Top Tiers'}], rows, {foot,maxH:560})+pnote, {right:note})}
    ${card('Guests by host team',`All ${fmt.int(uniq(selTrips.map(t=>t.gid)).length)} guests served by ${esc(selName)} · ${fmt.int(selTrips.length)} guest trips · ranked by trips`, guestTable(guestRows(selTrips,'trips',Infinity),['nat'])+(placeholder?`<p class="muted" style="font-size:11.5px;margin:10px 0 0">⚠ Placeholder: host teams are assigned at random.</p>`:''),
      {right:`<select class="sel" data-extra="custHost" aria-label="Host team">${options}</select>`})}
    ${guardSection(vg)}`;
}
// Guests by guard: guard = Arrangby (security team member who took the booking and dispatched the car)
function guardSection(vg){
  const byGuard=groupBy(vg.filter(t=>t.host),t=>t.host); if(!byGuard.size) return '';
  const all=[...byGuard.values()].flat(), guards=[...byGuard].sort((a,b)=>b[1].length-a[1].length);
  let sel=S.extra.custGuard||'__all'; if(sel!=='__all'&&!byGuard.has(sel)) sel='__all';
  const ts=sel==='__all'?all:byGuard.get(sel), who=sel==='__all'?'all guards':guardName(sel);
  const options=`<option value="__all" ${sel==='__all'?'selected':''}>All guards · ${fmt.int(all.length)} trips</option>`+
    guards.map(([id,x])=>`<option value="${esc(id)}" ${sel===id?'selected':''}>${esc(guardName(id))} · ${fmt.int(x.length)} trips</option>`).join('');
  return card('Guests by guard',`Guard = Arrangby (security team) · all ${fmt.int(uniq(ts.map(t=>t.gid)).length)} guests handled by ${esc(who)} · ${fmt.int(ts.length)} guest trips · ranked by trips`,
    guestTable(guestRows(ts,'trips',Infinity),['nat']), {right:`<select class="sel" data-extra="custGuard" aria-label="Guard">${options}</select>`});
}
function renderCustomers(pc, after){
  const c=pc.cur; const G=x=>x.trips.filter(t=>t.type==='Guest');
  if(!c.trips.length) return emptyModule('users','No trips in this selection','Choose another period or filter.');
  const defs=[
    {label:'Guests served', icon:'users', get:x=>uniq(G(x).map(t=>t.gid)).length, note:()=>'distinct guest IDs'},
    {label:'Guest trips', icon:'route', get:x=>G(x).length},
    {label:'Avg trips / guest', icon:'route', fmt:fmt.d1, get:x=>{const g=uniq(G(x).map(t=>t.gid)).length; return g?G(x).length/g:null;}},
    {label:'Avg km / guest trip', icon:'globe', fmt:fmt.d1, unit:'km', get:x=>G(x).length?sum(G(x),t=>t.km)/G(x).length:null},
    {label:'Premium Level share', icon:'gauge', fmt:fmt.pct, get:x=>G(x).length?G(x).filter(t=>CFG.targetTiers.includes(t.tier)).length/G(x).length:null, note:x=>`${fmt.int(uniq(G(x).filter(t=>CFG.targetTiers.includes(t.tier)).map(t=>t.gid)).length)} Premium guests (I+, V, V+)`},
    {label:'Avg hours / guest trip', icon:'clock', fmt:fmt.d2, unit:'h', get:x=>G(x).length?sum(G(x),hrs)/G(x).length:null, note:x=>G(x).length?`≈ ${fmt.int(sum(G(x),t=>t.dur)/G(x).length)} min per trip`:''},
  ];
  const g=G(c);
  const tiers=groupBy(g,t=>t.tier||'Unknown'), tk=sortTiers([...tiers.keys()]);
  const tierRows=tk.map(k=>{const ts=tiers.get(k), u=uniq(ts.map(t=>t.gid)).length; return {tier:k, group:levelOf(k).replace(' Level',''), guests:u, trips:ts.length, share:ts.length/g.length, tpg:ts.length/u, km:sum(ts,t=>t.km), avgKm:sum(ts,t=>t.km)/ts.length, topNat:topKey(countBy(ts,t=>t.nat||'–'))};});
  const byTpg=[...tierRows].sort((a,b)=>b.tpg-a.tpg);
  const cars=uniq(c.trips.map(t=>t.car)).sort();
  const selCar=S.extra.custCar&&cars.includes(S.extra.custCar)?S.extra.custCar:cars[0];
  const topBy=S.extra.custTopBy||'trips';

  const html = kpiRow(defs,pc,'six') + `
    <section class="grid g-3">
      ${card('Guest trips by tier','Number of guest trips per membership tier', canvas('cuTier'))}
      ${card('Guest trips by Nationality','Share of guest trips by nationality', natBody('cuNat',g.length,210,true))}
      ${card('Guest trips by Destination',CFG.areas.join(', ').replace('Ho Chi Minh','Ho Chi Minh (incl. airport)'), canvas('cuCity'))}
    </section>
    <section class="grid g-2-1">
      ${card('Usage intensity by tier','Who uses the service most / least — sorted by trips per guest', table([
        {k:'tier',label:'Tier'},{k:'group',label:'Tier Group',html:r=>`<span class="pill">${r.group}</span>`},{k:'guests',label:'Guests',r:1,f:fmt.int},{k:'trips',label:'Trips',r:1,f:fmt.int},
        {k:'share',label:'Share',r:1,f:fmt.pct},{k:'tpg',label:'Trips/guest',r:1,f:fmt.d1},{k:'avgKm',label:'Avg km/trip',r:1,f:fmt.d1},{k:'topNat',label:'Top nationality'}], byTpg,
        {foot:(()=>{const u=uniq(g.map(t=>t.gid)).length; return {tier:'Total',guests:u,trips:g.length,share:g.length?1:null,tpg:u?g.length/u:null,avgKm:g.length?sum(g,t=>t.km)/g.length:null,topNat:topKey(countBy(g,t=>t.nat||'–'))};})()}))}
      ${card('Most & least active tiers','By trips per guest · hover a value to see the guests in that group', `
        <div class="stat-row"><span class="l">Most active tier</span>${hovList(byTpg[0]?`${esc(byTpg[0].tier)} · ${fmt.d1(byTpg[0].tpg)} trips/guest`:'–', byTpg[0]&&g.filter(t=>t.tier===byTpg[0].tier), byTpg[0]&&'Tier '+byTpg[0].tier)}</div>
        <div class="stat-row"><span class="l">Least active tier</span>${(()=>{const r=byTpg[byTpg.length-1]; return hovList(r?`${esc(r.tier)} · ${fmt.d1(r.tpg)} trips/guest`:'–', r&&g.filter(t=>t.tier===r.tier), r&&'Tier '+r.tier);})()}</div>
        <div class="stat-row"><span class="l">Largest tier by trips</span>${(()=>{const r=tierRows.slice().sort((a,b)=>b.trips-a.trips)[0]; return hovList(r?esc(r.tier):'–', r&&g.filter(t=>t.tier===r.tier), r&&'Tier '+r.tier);})()}</div>
        <div class="stat-row"><span class="l">Top nationality</span>${(()=>{const n=topKey(countBy(g,t=>t.nat||'–')); return hovList(n?esc(n):'–', n&&g.filter(t=>(t.nat||'–')===n), n&&'Nationality '+n);})()}</div>
        <div class="stat-row"><span class="l">Repeat guests (≥ 5 trips)</span>${(()=>{const c=countBy(g,t=>t.gid); return hovList(fmt.int([...c.values()].filter(v=>v>=5).length), g.filter(t=>c.get(t.gid)>=5), 'Repeat guests (≥ 5 trips)');})()}</div>
        <div class="stat-row"><span class="l">One-time guests</span>${(()=>{const c=countBy(g,t=>t.gid); return hovList(fmt.int([...c.values()].filter(v=>v===1).length), g.filter(t=>c.get(t.gid)===1), 'One-time guests');})()}</div>`)}
    </section>
    ${card('Peak usage time','Guest trips starting by weekday × hour · '+esc(c.label), `<div id="cuHeat"></div>`)}
    ${card('Guest tiers by destination','Guest trips to each main area, split by customer tier · distinct guests next to each area · hover a colour to see the trips of that tier', canvas('cuArea','short')+(S.extra.destTable!=='show'?'':
      table([{k:'area',label:'Destination'},...tk.map(k=>({k:'t_'+k,label:k,r:1,f:v=>v?fmt.int(v):'–'})),{k:'total',label:'Total trips',r:1,f:fmt.int},{k:'guests',label:'Guests',r:1,f:fmt.int}],
        CFG.areas.map(a=>{ const ts=g.filter(t=>areaOf(t)===a), o={area:a,total:ts.length,guests:uniq(ts.map(t=>t.gid)).length}; tk.forEach(k=>o['t_'+k]=ts.filter(t=>(t.tier||'Unknown')===k).length); return o; }).filter(r=>r.total),
        {foot:Object.assign({area:'Total',total:g.length,guests:uniq(g.map(t=>t.gid)).length},...tk.map(k=>({['t_'+k]:tiers.get(k).length})))})),
      {right:`<div class="seg" data-extra="destTable"><button type="button" class="${S.extra.destTable!=='show'?'on':''}" data-v="hide">Hide table</button><button type="button" class="${S.extra.destTable==='show'?'on':''}" data-v="show">Show table</button></div>`})}
    ${card('Km per guest by tier','Average km per guest vs the guest with the most km in each tier', canvas('cuAreaKm','tall'))}
    <section class="grid g-2">
      ${card(topTitle('custTopN',50,'guests'), 'Ranked across all vehicles', guestTable(guestRows(g,topBy,topN_('custTopN',50))), {right:topNSel('custTopN',50)+`<div class="seg" data-extra="custTopBy">${[['trips','By trips'],['km','By km'],['hours','By hours']].map(([k,l])=>`<button type="button" class="${topBy===k?'on':''}" data-v="${k}">${l}</button>`).join('')}</div>`})}
      ${card(topTitle('custCarN',10,'guests per vehicle'), 'Ranked by trips', guestTable(guestRows(g.filter(t=>t.car===selCar),'trips',topN_('custCarN',10))), {right:topNSel('custCarN',10)+`<select class="sel" data-extra="custCar" aria-label="Vehicle">${cars.map(x=>`<option ${x===selCar?'selected':''}>${esc(x)}</option>`).join('')}</select>`})}
    </section>
    ${hostSection(g)}
`;
  after.push(()=>{
    const P=palette();
    mkChart('cuTier','bar',tk,[{label:'Guest trips',data:tierRows.map(r=>r.trips),color:P.s[0]}],{unit:'trips'});
    natChart('cuNat',g,P,{list:true});
    areaChart('cuCity',g,P);
    // one bar per destination area, stacked by tier (e.g. Binh Duong: 100 I, 50 V, 10 EC …)
    const areaGuests=a=>uniq(g.filter(t=>areaOf(t)===a).map(t=>t.gid)).length;
    mkChart('cuArea','bar',CFG.areas.map(a=>`${a} · ${fmt.int(areaGuests(a))} guests`),tk.map(k=>({label:k,color:tierColor(k,P),data:CFG.areas.map(a=>tiers.get(k).filter(t=>areaOf(t)===a).length)})),{horizontal:true,stacked:true,legend:true,unit:'trips'});
    // per tier: km of each guest (sum of their trips in that tier) -> average across guests, and the highest single guest
    const kmPerGuest=tk.map(k=>[...groupBy(tiers.get(k),t=>t.gid).values()].map(ts=>sum(ts,t=>t.km)));
    mkChart('cuAreaKm','bar',tk,[{label:'Avg km per guest',color:P.s[0],order:1,data:kmPerGuest.map(v=>v.length?sum(v)/v.length:0)},
      {type:'line',label:'Max km (top guest)',color:P.s[1],order:0,points:true,data:kmPerGuest.map(v=>v.length?Math.max(...v):0)}],{legend:true,fmt:fmt.int,unit:'km'});
    const m=WD.map(()=>new Array(24).fill(0)); g.forEach(t=>m[weekday(t.d)][Math.floor(t.dep/60)%24]++);
    $('#cuHeat').innerHTML=heatmap(WD,[...Array(24).keys()].map(pad),m,fmt.int,'trips');
  });
  return html;
}

/* ---------- Drivers ---------- */
function driverStats(ctx){
  const by=groupBy(ctx.vegasAll.filter(t=>t.drv),t=>t.drv);
  const rosterIn=S.roster.filter(r=>r.d>=ctx.a&&r.d<=ctx.b);
  const ids=new Set([...by.keys(), ...rosterIn.filter(r=>!r.off).map(r=>r.drv)]); // rostered drivers show up even with no trips
  return [...ids].map(drv=>{
    const ts=by.get(drv)||[];
    const rs=rosterIn.filter(r=>r.drv===drv), work=rs.filter(r=>!r.off);
    const sd=groupBy(ts,shiftDay), dh=sum(ts,t=>t.dh), g=ts.filter(t=>t.type==='Guest');
    // days on duty: rostered working days, plus shift-days with trips on dates this driver has no roster row
    const rosterDates=new Set(work.map(r=>r.d)), extra=[...sd.keys()].filter(d=>!rosterDates.has(d)&&!ROSTER_IDX.has(drv+'|'+d));
    const shiftDays=work.length+extra.length;
    const avail=sum(work,availHours)+extra.length*CFG.maxDrive;
    const offShift=ts.filter(t=>hasRosterDay(drv,t.d)&&!rosterFor(t)).length;
    const maxSd=sd.size?Math.max(...[...sd.values()].map(a=>sum(a,t=>t.dh))):0;
    return {drv, name:driverName(drv), dedicated:CFG.dedicatedDrivers.includes(drv), shiftDays, rostered:work.length, offDays:rs.length-work.length,
      shiftH:sum(work,shiftHours), avail, trips:ts.length, tps:shiftDays?ts.length/shiftDays:null, hours:dh, avgSd:shiftDays?dh/shiftDays:null, maxSd,
      use:avail?dh/avail:null, offShift, km:sum(ts,t=>t.km), tier:topKey(countBy(g,t=>t.tier||'–'))||'–', nat:topKey(countBy(g,t=>t.nat||'–'))||'–', internal:ts.length?ts.filter(t=>t.type==='Internal').length/ts.length:null};
  }).sort((a,b)=>(a.dedicated-b.dedicated)||b.hours-a.hours);
}
function renderDrivers(pc, after){
  if(S.f.service==='G') return emptyModule('wheel','Driver analysis covers Vegas drivers only','Grab trips have no Vegas driver. Switch the service filter to <b>All</b> or <b>Vegas Car</b>.');
  const c=pc.cur; if(!c.vegasAll.length) return emptyModule('wheel','No Vegas Car trips in this period','Choose another period or upload data.');
  const poolStats=x=>driverStats(x).filter(d=>!d.dedicated);
  const defs=[
    {label:'Active pool drivers', icon:'users', get:x=>poolStats(x).length, note:()=>`team ${CFG.teamDrivers} incl. ${CFG.dedicatedDrivers.length} dedicated`},
    {label:'Avg shift-days worked', icon:'calendar', fmt:fmt.d1, get:x=>{const p=poolStats(x); return p.length?sum(p,d=>d.shiftDays)/p.length:null;}, note:x=>capacity(x)&&capacity(x).basis!=='estimate'?'per pool driver · rostered days':'per pool driver · days with trips'},
    {label:'Avg driving h / shift-day', icon:'clock', fmt:fmt.d2, unit:'h', get:x=>{const p=poolStats(x).filter(d=>d.shiftDays); return p.length?sum(p,d=>d.avgSd)/p.length:null;}},
    {label:'Use of available driving time', icon:'gauge', fmt:fmt.pct, get:x=>{const p=poolStats(x), av=sum(p,d=>d.avail); return av?sum(p,d=>d.hours)/av:null;}, note:x=>{const k=capacity(x); return !k?'':k.basis==='roster'?'from actual shifts':k.basis==='mixed'?`actual shifts on ${k.rosterDays} of ${x.days} days`:'estimated (6h per shift-day)';}},
    {label:'Capacity used', icon:'chart', fmt:fmt.pct, get:x=>{const k=capacity(x); return k&&k.usedPct;}},
    {label:'Additional trips / day', icon:'route', get:x=>{const k=capacity(x); return k&&k.add;}, note:x=>{const k=capacity(x); return k?`max sustainable ≈ ${fmt.int(k.max)} trips/day`:'';}},
  ];
  const ds=driverStats(c), pool=c.vegasAll.filter(isPool); // staffing & capacity use all trips
  const view=S.extra.drvView||'weekday';
  const cap=capacity(c);
  // shift workload (pool drivers)
  const shiftRows=CFG.shifts.map(([name,a,b])=>{
    let dep=0, vh=0;
    pool.forEach(t=>{ if((t.dep>=a&&t.dep<b)||(t.dep+1440>=a&&t.dep+1440<b)) dep++; const e=t.dep+t.dh*60; vh+=(overlap(t.dep,e,a,b)+overlap(t.dep,e,a-1440,b-1440))/60; });
    const depD=dep/c.days, vhD=vh/c.days, wl=vhD+depD*CFG.buffer; return {name,dep:depD,vh:vhD,wl,min:wl/CFG.maxDrive};
  });
  const totWl=sum(shiftRows,r=>r.wl); shiftRows.forEach(r=>r.sugg=cap?cap.onDuty*r.wl/totWl:null);
  // with a roster: average pool drivers rostered at the middle of each shift window, on days that have a roster
  const rDays=[]; for(let d=c.lo; c.days&&d<=c.hi; d=addDays(d,1)) if(ROSTER_DAYS.has(d)) rDays.push(d);
  const hasRoster=rDays.length>0;
  if(hasRoster){ const work=S.roster.filter(r=>!r.off&&!CFG.dedicatedDrivers.includes(r.drv));
    shiftRows.forEach((row,i)=>{ const [,a,b]=CFG.shifts[i], mid=(a+b)/2; let n=0;
      for(const d of rDays){ n+=work.filter(r=>{ const off=diffDays(d,r.d)*1440; return r.s+off<=mid&&mid<r.e+off; }).length; }
      shiftRows[i].rost=n/rDays.length; shiftRows[i].gap=shiftRows[i].rost-shiftRows[i].min; }); }
  const anyRoster=ds.some(d=>d.rostered||d.offDays);
  const basisBadge=`<span class="badge">${!cap||cap.basis==='estimate'?'Shifts: estimated (no roster uploaded)':cap.basis==='roster'?'Shifts: actual roster':`Shifts: actual roster on ${cap.rosterDays} of ${c.days} days`}</span>`;
  const html = `<div class="section-title">Drivers ${vegasOnlyNote()} ${basisBadge}</div>` + kpiRow(defs,pc,'six') + `
    ${card('Driver scorecard',(anyRoster?'Shift-days = rostered working days (plus days with trips but no roster row) · use = driving h ÷ (shift length − break) · off-shift = trips outside the rostered shift':'Shift-day = trips before 05:30 count to the previous day\'s shift · use = driving h ÷ 6h per shift-day')+' · highlighted row = dedicated boss driver (excluded from pool averages)', table([
      {k:'name',label:'Driver',html:r=>`${esc(r.name)} <span class="muted" style="font-size:11px">#${esc(r.drv)}</span>`+(r.dedicated?' <span class="pill">Dedicated</span>':'')},
      {k:'shiftDays',label:'Shift-days',r:1,f:fmt.int},...(anyRoster?[{k:'offDays',label:'Days off',r:1,f:fmt.int},{k:'shiftH',label:'Shift h',r:1,f:fmt.d1}]:[]),
      {k:'trips',label:'Trips',r:1,f:fmt.int},{k:'tps',label:'Trips/shift-day',r:1,f:fmt.d1},
      {k:'hours',label:'Driving h',r:1,f:fmt.d1},{k:'avgSd',label:'Avg h/shift-day',r:1,f:fmt.d2},{k:'maxSd',label:'Max h/shift-day',r:1,f:fmt.d1},
      {k:'use',label:anyRoster?'Use of avail. h':'Use of 6h',r:1,f:fmt.pct},...(anyRoster?[{k:'offShift',label:'Off-shift trips',r:1,f:fmt.int}]:[]),
      {k:'km',label:'Km',r:1,f:fmt.int},{k:'tier',label:'Main tier'},{k:'nat',label:'Main nationality'},{k:'internal',label:'Internal %',r:1,f:fmt.pct}],
      ds,{flag:r=>r.dedicated,maxH:560}))}
    ${card('Driving hours per driver', view==='weekday'?'Avg driving hours per weekday (hours ÷ number of that weekday in period)':'Total driving hours in each clock hour · '+esc(c.label), `<div id="drHeat"></div>`,
      {right:`<div class="seg" data-extra="drvView"><button type="button" class="${view==='weekday'?'on':''}" data-v="weekday">By weekday</button><button type="button" class="${view==='hour'?'on':''}" data-v="hour">By time slot</button></div>`})}
    <section class="grid g-2">
      ${card('Avg drivers on the road by hour','Pool drivers (excl. boss driver)', canvas('drHour','tall'))}
      ${card('Drivers on duty per shift-day',hasRoster?'Rostered drivers vs drivers with ≥ 1 trip (excl. boss driver)':'Drivers with ≥ 1 trip (excl. boss driver)', canvas('drDuty','tall'))}
    </section>
    <section class="grid ${hasRoster?'':'g-2-1'}">
      ${card('Workload by shift & suggested staffing','Pool drivers · per day · workload = driving hours + pre-trip buffer', table([
        {k:'name',label:'Shift'},{k:'dep',label:'Departures/day',r:1,f:fmt.d1},{k:'vh',label:'Driving h/day',r:1,f:fmt.d1},{k:'wl',label:'Workload h/day',r:1,f:fmt.d1},
        {k:'min',label:'Min. drivers',r:1,f:fmt.d1},{k:'sugg',label:'Suggested on duty',r:1,f:fmt.d1},
        ...(hasRoster?[{k:'rost',label:'Rostered on duty',r:1,f:fmt.d1},{k:'gap',label:'Rostered − needed',r:1,html:r=>`<span style="font-weight:700;color:${r.gap<0?'var(--s8)':'inherit'}">${r.gap>0?'+':''}${fmt.d1(r.gap)}</span>`}]:[])], shiftRows,
        {foot:{name:'Total',dep:sum(shiftRows,r=>r.dep),vh:sum(shiftRows,r=>r.vh),wl:totWl,min:sum(shiftRows,r=>r.min),sugg:cap&&cap.onDuty}}) + `<div class="chart short" style="margin-top:14px"><canvas id="drShift"></canvas></div>`)}
      ${card('Daily driver capacity','Hours per day · binding constraint: '+(cap?cap.constraint:'–'), `<div class="chart short"><canvas id="drCap"></canvas></div>
        ${cap?[['Hours used',cap.used,0],['Spare',cap.spare,-1]].map(([l,v,i])=>`<div class="stat-row"><span class="l"><i class="dot" style="background:${i<0?palette().neutral:palette().s[i]}"></i>${l}</span><span class="v num">${fmt.d1(v)} h</span></div>`).join(''):''}
        <div class="stat-row"><span class="l">Drivers on duty / day (${cap&&cap.basis!=='estimate'?'actual roster':'estimated roster'})</span><span class="v num">${fmt.d1(cap&&cap.onDuty)}</span></div>
        <div class="stat-row"><span class="l">Driver capacity</span><span class="v num">${fmt.d1(cap&&cap.driverCap)} h</span></div>
        <div class="stat-row"><span class="l">Vehicle capacity</span><span class="v num">${fmt.d1(cap&&cap.vehicleCap)} h</span></div>`)}
    </section>`;
  after.push(()=>{
    const P=palette();
    const drvs=ds.map(d=>d.name+(d.dedicated?' (ded.)':''));
    let m;
    if(view==='weekday'){ const wc=weekdayCount(c); m=ds.map(d=>{const row=new Array(7).fill(0); c.vegasAll.filter(t=>t.drv===d.drv).forEach(t=>row[weekday(shiftDay(t))]+=t.dh); return row.map((v,i)=>wc[i]?v/wc[i]:0);}); $('#drHeat').innerHTML=heatmap(drvs,WD,m,fmt.d1,'h/day'); }
    else { m=ds.map(d=>{const row=new Array(24).fill(0); c.vegasAll.filter(t=>t.drv===d.drv).forEach(t=>eachHourOverlap(t.dep,t.dh*60,(h,o)=>row[h%24]+=o/60)); return row;}); $('#drHeat').innerHTML=heatmap(drvs,[...Array(24).keys()].map(pad),m,fmt.d1,'h'); }
    const hp=hourProfile(pool,c.days,t=>t.drv);
    mkChart('drHour','line',[...Array(24).keys()].map(h=>pad(h)+':00'),[{label:'Avg drivers on the road',data:hp.avgOnRoad,color:P.s[6]},{label:'Max distinct drivers (busiest day)',data:hp.maxDistinct,color:P.s[1],dash:[5,4]}],{legend:true,fmt:fmt.d1,unit:'drivers'});
    const duty=groupBy(pool,shiftDay); const dk=[...duty.keys()].filter(d=>d>=c.a&&d<=c.b).sort();
    const dutyDays=uniq(dk.concat(rDays)).sort();
    const dsets=[{label:'Drivers with trips',data:dutyDays.map(d=>duty.has(d)?uniq(duty.get(d).map(t=>t.drv)).length:0),color:P.s[0]}];
    if(hasRoster) dsets.unshift({label:'Rostered',data:dutyDays.map(d=>ROSTER_DAYS.has(d)?S.roster.filter(r=>r.d===d&&!r.off&&!CFG.dedicatedDrivers.includes(r.drv)).length:null),color:P.s[6]});
    mkChart('drDuty','bar',dutyDays.map(d=>c.gran==='month'?+d.slice(8):fmtDay(d).slice(0,6)),dsets,{unit:'drivers',legend:hasRoster});
    mkChart('drShift','bar',shiftRows.map(r=>r.name.split(' ')[0]),[{label:'Min. drivers needed',data:shiftRows.map(r=>r.min),color:P.s[0]},{label:'Suggested on duty',data:shiftRows.map(r=>r.sugg),color:P.s[1]},...(hasRoster?[{label:'Rostered on duty',data:shiftRows.map(r=>r.rost),color:P.s[6]}]:[])],{legend:true,fmt:fmt.d1,unit:'drivers'});
    if(cap) mkChart('drCap','doughnut',['Hours used','Spare'],[{data:[cap.used,cap.spare],backgroundColor:[P.s[0],P.neutral]}],{fmt:fmt.d1,unit:'h'});
  });
  return html;
}

/* ---------- Car Cost ---------- */
const ROSTER_SCHEMA = `<div class="schema">${table([{k:'a',label:'Date'},{k:'b',label:'EmployeeId'},{k:'c',label:'Start'},{k:'d',label:'End'},{k:'e',label:'Break (h)'},{k:'f',label:'Status'}],
  [{a:'2026-10-05',b:'78',c:'06:00',d:'14:00',e:'2',f:''},{a:'2026-10-05',b:'74',c:'22:00',d:'06:00',e:'2',f:''},{a:'2026-10-05',b:'49',c:'',d:'',e:'',f:'Off'}])}
  <p class="muted" style="font-size:12px;margin-top:8px">Sheet <code>Driver Roster</code>: one row per driver per day. An End earlier than Start runs past midnight. Break is optional (default ${CFG.breakH}h); available driving time = shift length − break. Leave Start/End blank or write Off / Leave in Status for a day off. Sheet <code>Drivers</code>: <code>EmployeeId | Name</code> to show names instead of IDs.</p></div>`;
const COST_SCHEMA = `<div class="schema">${table([{k:'a',label:'Date'},{k:'b',label:'Vehicle'},{k:'c',label:'Category'},{k:'d',label:'Amount (USD)',r:1}],
  [{a:'2026-10-05',b:'Benz V250',c:'Maintenance',d:'1,250'},{a:'2026-10-31',b:'Benz V250',c:'Depreciation',d:'2,100'},{a:'2026-10-12',b:'KIA Carnival (84)',c:'Oil & Gasoline',d:'64'}])}
  <p class="muted" style="font-size:12px;margin-top:8px">Add a sheet named <code>Car Cost</code> with these 4 columns: one row per expense (date × vehicle × category). Dates as Excel dates or YYYY-MM-DD; a cost can then be viewed by day, month, quarter or year. Categories: ${CFG.costCats.join(', ')}. Vehicle names must match the <code>Brand</code> column of the trip log (upper / lower case does not matter). The VG Car Cost file layout (<code>Date, Car Brand, Cost Type, Amount (Usd)</code>) is also accepted.</p></div>`;
function renderCarCost(pc, after){
  const c=pc.cur;
  if(!S.costs.length){
    return emptyModule('wrench','Car cost data not available yet',`This module shows cost by category (${CFG.costCats.join(', ')}) for each vehicle, cost per trip / per km, and the monthly / quarterly trend. It fills in automatically once a <b>Car Cost</b> sheet is uploaded via <b>Update data</b>.`, COST_SCHEMA)
      + `<section class="grid g-3 ghost-grid" aria-hidden="true">${['Cost by category per vehicle','Monthly cost trend','Cost per trip & per km'].map(t=>card(t,'Awaiting data','<div class="ph">Preview of layout</div>')).join('')}</section>`;
  }
  const defs=[
    {label:'Total car cost', icon:'money', fmt:fmt.usd, get:x=>sum(x.costs,r=>r.amt)},
    {label:'Cost per trip', icon:'money', fmt:fmt.usd, get:x=>{const n=x.all.filter(t=>t.src==='V').length; return n?sum(x.costs,r=>r.amt)/n:null;}},
    {label:'Cost per km', icon:'globe', fmt:fmt.usd, get:x=>{const k=sum(x.all.filter(t=>t.src==='V'),t=>t.km); return k?sum(x.costs,r=>r.amt)/k:null;}},
    {label:'Largest category', icon:'chart', fmt:v=>v||'–', noCompare:true, get:x=>{const m=new Map(); x.costs.forEach(r=>m.set(r.cat,(m.get(r.cat)||0)+r.amt)); return topKey(m);}},
  ];
  if(!c.costs.length) return kpiRow(defs,pc) + emptyModule('wrench','No car cost rows for '+c.label,'Choose a period covered by the cost sheet.');
  const cats=uniq(c.costs.map(r=>r.cat)), cars=uniq(c.costs.map(r=>r.car)).sort();
  const months=uniq(c.costs.map(r=>r.m)).sort();
  const vrows=cars.map(car=>{ const rs=c.costs.filter(r=>r.car===car), ts=c.all.filter(t=>t.src==='V'&&t.car===car), tot=sum(rs,r=>r.amt), o={car,total:tot,trips:ts.length,cpt:ts.length?tot/ts.length:null,cpk:sum(ts,t=>t.km)?tot/sum(ts,t=>t.km):null}; cats.forEach(k=>o[k]=sum(rs.filter(r=>r.cat===k),r=>r.amt)); return o; }).sort((a,b)=>b.total-a.total);
  const html = kpiRow(defs,pc) + `
    <section class="grid g-2">
      ${card('Cost by category per vehicle','USD · '+esc(c.label), canvas('ccCar','tall'))}
      ${card('Cost trend by month','USD, stacked by category', canvas('ccTrend','tall'))}
    </section>
    ${card('Vehicle cost table','USD', table([{k:'car',label:'Vehicle'},...cats.map(k=>({k,label:k,r:1,f:fmt.usd})),{k:'total',label:'Total',r:1,f:fmt.usd},{k:'trips',label:'Trips',r:1,f:fmt.int},{k:'cpt',label:'Cost/trip',r:1,f:fmt.usd},{k:'cpk',label:'Cost/km',r:1,f:fmt.usd}], vrows))}`;
  after.push(()=>{
    const P=palette(); const topCats=cats.slice(0,8);
    mkChart('ccCar','bar',cars,topCats.map(k=>({label:k,color:colorFor('cat',k,P),data:cars.map(car=>sum(c.costs.filter(r=>r.car===car&&r.cat===k),r=>r.amt))})),{horizontal:true,stacked:true,legend:true,fmt:fmt.usd});
    mkChart('ccTrend','bar',months.map(m=>MON[+m.slice(5)-1]+' '+m.slice(2,4)),topCats.map(k=>({label:k,color:colorFor('cat',k,P),data:months.map(m=>sum(c.costs.filter(r=>r.m===m&&r.cat===k),r=>r.amt))})),{stacked:true,legend:true,fmt:fmt.usd});
  });
  return html;
}

/* ---------- Grab Cost ---------- */
const GRAB_SCHEMA = `<div class="schema"><p class="muted" style="font-size:12px">Add a sheet whose name contains <code>Grab</code> (e.g. <code>Grab.2610</code>) using the same column names as the Vegas trip log — <code>GuestID, Type, Customer Tier, Nationality, Date, Depart, Arrival, Occupancy, Distance, Destination, CITY</code> — plus a <code>Cost (USD)</code> column. Driver / vehicle columns are not needed.</p></div>`;
function renderGrabCost(pc, after){
  const c=pc.cur; const hasGrab=S.trips.some(t=>t.src==='G');
  if(!hasGrab) return emptyModule('phone','Grab data not available yet','This module will mirror the Vegas Car analysis for Grab rides — trips, guests, tiers, km, time — plus the cost of every ride (cost per trip, per km, per tier, top guests by cost). There is no driver analysis for Grab. Grab trips will also flow into Overview and Customers when the service filter is <b>All</b> or <b>Grab</b>.', GRAB_SCHEMA)
    + `<section class="grid g-3 ghost-grid" aria-hidden="true">${['Grab cost trend','Cost by customer tier','Top guests by Grab cost'].map(t=>card(t,'Awaiting data','<div class="ph">Preview of layout</div>')).join('')}</section>`;
  const gr=x=>x.all.filter(t=>t.src==='G'&&x.typeOk(t));
  const defs=[
    {label:'Grab trips', icon:'route', get:x=>gr(x).length},
    {label:'Guests using Grab', icon:'users', get:x=>uniq(gr(x).filter(t=>t.type==='Guest').map(t=>t.gid)).length},
    {label:'Grab cost', icon:'money', fmt:fmt.usd, get:x=>sum(gr(x),t=>t.cost||0)},
    {label:'Avg cost / trip', icon:'money', fmt:fmt.usd, get:x=>gr(x).length?sum(gr(x),t=>t.cost||0)/gr(x).length:null},
    {label:'Cost / km', icon:'globe', fmt:fmt.usd, get:x=>{const k=sum(gr(x),t=>t.km); return k?sum(gr(x),t=>t.cost||0)/k:null;}},
    {label:'Grab share of all trips', icon:'chart', fmt:fmt.pct, get:x=>{const a=x.all.filter(x.typeOk); return a.length?gr(x).length/a.length:null;}},
  ];
  const g=gr(c);
  if(!g.length) return kpiRow(defs,pc,'six') + emptyModule('phone','No Grab trips in '+c.label,'Choose another period.');
  const tiers=groupBy(g,t=>t.type==='Internal'?'Internal':(t.tier||'Unknown')), tk=sortTiers([...tiers.keys()]);
  const top=[...groupBy(g,t=>t.gid)].map(([gid,ts])=>({gid,guest:ts[0].type==='Guest',tier:ts[0].tier||'–',trips:ts.length,km:sum(ts,t=>t.km),cost:sum(ts,t=>t.cost||0),
    maxMonth:Math.max(...[...groupBy(ts,t=>t.d.slice(0,7)).values()].map(a=>sum(a,t=>t.cost||0)))})).sort((a,b)=>b.cost-a.cost).slice(0,topN_('grabTopN',20));
  const over=r=>r.guest&&r.maxMonth>CFG.grabMonthlyLimit, nOver=top.filter(over).length;
  const html = kpiRow(defs,pc,'six') + `
    <section class="grid g-2">
      ${card('Grab cost over time','USD per '+trendBuckets(c,[]).unit, canvas('grTrend','tall'))}
      ${card('Grab cost by customer tier','USD', canvas('grTier','tall'))}
    </section>
    ${card(topTitle('grabTopN',20,'guests by Grab cost'),`Red = guest with more than ${fmt.usd(CFG.grabMonthlyLimit)} of Grab in one calendar month${nOver?` · ${nOver} in this list`:''}`, table([{k:'gid',label:'Guest ID'},{k:'tier',label:'Tier'},{k:'trips',label:'Trips',r:1,f:fmt.int},{k:'km',label:'Km',r:1,f:fmt.int},{k:'cost',label:'Cost',r:1,f:fmt.usd},{k:'maxMonth',label:'Highest month',r:1,f:fmt.usd}],top,{rowCls:r=>over(r)?'over':''}),{right:topNSel('grabTopN',20)})}`;
  after.push(()=>{
    const P=palette(); const tb=trendBuckets(c,g,t=>t.cost||0);
    mkChart('grTrend','bar',tb.labels,[{label:'Grab cost',data:tb.vals,color:P.s[2]}],{fmt:fmt.usd});
    mkChart('grTier','bar',tk,[{label:'Grab cost',data:tk.map(k=>sum(tiers.get(k),t=>t.cost||0)),color:P.s[2]}],{fmt:fmt.usd});
  });
  return html;
}

/* ---------- Data & definitions ---------- */
const DEFS = [
  ['Trip','One row in the trip log = one car movement from departure to arrival.'],
  ['Trip duration (hours)','Time from guest boarding until the driver returns the car to the car park (includes the empty return to base).'],
  ['Driving hours','Trip duration, capped at 0.5h for "Maintenance Car", "Take Car Back" and trips to "Garage" (same rule as the source sheet). Used for driver and capacity metrics.'],
  ['Calendar day vs shift-day','Calendar day = date in the log. Shift-day: trips departing before 05:30 belong to the previous day\'s shift, so a night shift counts as one day.'],
  ['Days with data','Calendar days in the selected period that fall inside the data range. All "per day" averages divide by this number.'],
  ['Pool drivers','All Vegas drivers except the dedicated boss driver (ID 35). Pool averages and capacity exclude driver 35; the driver is still shown in the scorecard (highlighted).'],
  ['Main areas','Ho Chi Minh (all HCMC destinations, airport included), Binh Duong, Dong Nai, Long An, and Others (OutSide, Tay Ninh, Vung Tau and other provinces). Based on the CITY column.'],
  ['Host team','Each Vegas Car guest trip is booked by a host team (HM, CM, IM, BD, GM). The team comes from a "Host Team" column in the trip log. Until that data is uploaded, every trip is assigned to a team at random (placeholder, marked "Placeholder data"); the same trip always stays in the same team.'],
  ['Driver roster','When a Driver Roster is uploaded, a driver\'s working days, available driving time (shift length − break) and the trips that fall inside each shift come from the roster: drivers on duty and driver capacity use the real roster on those days, rostered days without trips count, and trips outside the rostered shift are counted as off-shift. Days without a roster keep the estimate (13.4 drivers on duty × 6h, shift-day cut-off 05:30).'],
  ['Customer levels','VIP Level = ONE, ONE+, EC. Premium Level = I+, V, V+ (the tiers OD wants to extend the service to). Potential Level = P, I. Internal Group = staff and operational trips. X = banned guest (shown as Banned Guest, grey). Any other tier code shows as Other. Premium share is calculated on guest trips (internal excluded).'],
  ['Guest vs internal','Guest = customer trips. Internal = staff, boss and operational trips (fuel, maintenance, food…). Customer metrics use guest trips only; internal usage is shown as a separate group.'],
  ['Capacity used','(Driving hours + pre-trip buffer per trip) ÷ binding capacity. Binding capacity = lower of driver capacity and vehicle capacity.'],
  ['Additional trips / day','Spare hours ÷ (avg driving hours per trip + pre-trip buffer), rounded down.'],
  ['Cars in use (distinct)','Number of different cars on the road in a clock hour on a day (a car doing two trips in the hour counts once), averaged over days with data.'],
  ['In yard','24h − average hours on the road per day for that vehicle.'],
  ['Utilization (16h)','Average hours on the road per day ÷ 16h service window.'],
  ['Comparison','"Previous period" = the period just before (previous day / month / quarter / year). "Same period last year" = the same day / month / quarter one year earlier. Shown under each KPI when data for that period exists.'],
  ['Trip type filter',`All / Guest / Internal applies to Overview${CFG.features.grab?' and Grab Cost (combined with the Service filter)':''}. Customers always counts guest trips only (the filter is pinned to Guest there). VG Cars and Drivers always use all trips, and capacity figures (capacity used, additional trips/day) always use all trips, because internal trips use the same drivers and cars. Car Cost is not split by trip type.`],
  ['Service filter','All / Vegas Car / Grab applies to Overview and Customers. VG Cars and Drivers always use Vegas Car trips (Grab has no fleet or Vegas drivers). Car Cost follows the Vegas fleet; Grab Cost shows Grab rides.'],
];
function renderAbout(){
  const [mn,mx]=dataBounds(); const v=S.trips.filter(t=>t.src==='V').length, g=S.trips.filter(t=>t.src==='G').length;
  const months=uniq(S.trips.map(t=>t.d.slice(0,7))).sort();
  return `<section class="grid g-2">
    ${card('Data loaded','Source: '+esc(S.source), `
      <div class="stat-row"><span class="l">Date range</span><span class="v">${mn?fmtDay(mn)+' – '+fmtDay(mx):'–'}</span></div>
      <div class="stat-row"><span class="l">Months</span><span class="v">${months.map(m=>periodLabel('month',m)).join(', ')||'–'}</span></div>
      <div class="stat-row"><span class="l">Vegas Car trips</span><span class="v num">${fmt.int(v)}</span></div>
      ${CFG.features.grab?`<div class="stat-row"><span class="l">Grab trips</span><span class="v num">${g?fmt.int(g):'<span class="muted">not loaded</span>'}</span></div>`:''}
      <div class="stat-row"><span class="l">Car cost rows</span><span class="v num">${S.costs.length?fmt.int(S.costs.length):'<span class="muted">not loaded</span>'}</span></div>`)}
    ${card('Model assumptions','From the Excel Analysis sheet (validate with Ops)', `
      ${[['Drivers in team',CFG.teamDrivers],['Dedicated (boss) driver ID',CFG.dedicatedDrivers.join(', ')],['Off days per 2 weeks',CFG.offDays],['Productive driving h / shift',CFG.maxDrive+'h (8h shift − 2h break)'],
         ['Pre-trip buffer',CFG.buffer*60+' min'],['Active fleet',CFG.fleet+' vehicles'],['Service window',CFG.serviceWindow+'h / day'],['Target vehicle utilization',fmt.pct(CFG.targetUtil)],['Shift-day cut-off','05:30']]
        .map(([l,v])=>`<div class="stat-row"><span class="l">${l}</span><span class="v">${esc(v)}</span></div>`).join('')}`)}
  </section>
  ${card('Definitions','How every metric is calculated', `<dl class="def">${DEFS.filter(([t])=>CFG.features.grab||t!=='Service filter').map(([t,d])=>`<dt>${esc(t)}</dt><dd>${esc(d)}</dd>`).join('')}</dl>`)}
  ${card('How to update data','Monthly routine', `<ol style="margin:0;padding-left:18px;color:var(--text-secondary)">
    <li>Click <b>Update data</b> and drop the monthly Excel file (same column layout as the monthly VG Car data sheet). Several files or sheets can be loaded at once; uploading the same dates again replaces them.</li>
    <li>Optional: add a ${CFG.features.grab?'<code>Grab…</code> sheet and a ':''}<code>Car Cost</code> sheet (format below) to unlock the cost figures.</li>
    <li>Click <b>Save dashboard as HTML</b> and share the new file with OD. It opens on any phone or computer, with no installation needed.</li></ol>
    <h4 style="margin:16px 0 0">Car Cost sheet</h4>${COST_SCHEMA}${CFG.features.grab?`<h4 style="margin:16px 0 0">Grab sheet</h4>${GRAB_SCHEMA}`:''}`)}`;
}

/* ============================ shell render ============================ */
function renderNav(){
  $('#nav').innerHTML = `<div class="logo" aria-hidden="true">V</div>` + MODULES.map((m,i)=>
    `${i===MODULES.length-1?'<div class="nav-spacer"></div>':''}<button type="button" class="nav-btn ${S.mod===m.id?'active':''}" data-mod="${m.id}" aria-label="${m.name}" ${S.mod===m.id?'aria-current="page"':''}>${svg(m.icon)}<span class="tip">${m.name}</span></button>`).join('') +
    (S.adminOK?`<button type="button" class="nav-btn" data-admin aria-label="Admin">${svg('admin')}<span class="tip">Admin</span></button>`:'');
}
function renderFilters(){
  const periods=listPeriods(S.f.gran);
  if(!periods.includes(S.f.period)) S.f.period=periods[0]||null;
  const seg=(key,opts,cur,locked)=>`<div class="seg" role="group" ${locked?`title="${locked}"`:''}>${opts.map(([v,l])=>`<button type="button" class="${(cur??S.f[key])===v?'on':''}" data-f="${key}" data-v="${v}" ${locked?'disabled':''}>${l}</button>`).join('')}</div>`;
  $('#filters').innerHTML = `
    <div class="f-group"><span class="f-label">View by</span>${seg('gran',[['day','Day'],['month','Month'],['quarter','Quarter'],['year','Year']])}</div>
    <div class="f-group"><span class="f-label">Period</span><select class="sel" id="fPeriod" aria-label="Period">${periods.map(p=>`<option value="${p}" ${p===S.f.period?'selected':''}>${periodLabel(S.f.gran,p)}</option>`).join('')}</select></div>
    <div class="f-group"><span class="f-label">Compare with</span><select class="sel" id="fCompare" aria-label="Compare with">
      ${[['prev','Previous period'],['yoy','Same period last year'],['none','No comparison']].map(([v,l])=>`<option value="${v}" ${S.f.compare===v?'selected':''}>${l}</option>`).join('')}</select></div>
    ${CFG.features.grab?`<div class="f-group"><span class="f-label">Service</span>${seg('service',[['all','All'],['V','Vegas Car'],['G','Grab']])}</div>`:''}
    <div class="f-group"><span class="f-label">Trip type</span>${seg('ttype',[['all','All'],['Guest','Guest'],['Internal','Internal']],effTtype(),S.mod==='customers'?'Customers counts guest trips only':null)}</div>`;
}
function render(){
  destroyCharts();
  renderNav(); renderFilters();
  const mod=MODULES.find(m=>m.id===S.mod)||MODULES[0];
  const after=[];
  let html;
  if(!S.f.period) html=emptyModule('info','No data loaded','Click <b>Update data</b> to upload the Excel report.');
  else { const pc=ctxPair(); html=mod.render(pc,after); }
  $('#title').textContent = mod.id==='overview'?'Car Service Dashboard':mod.name;
  $('#hello').textContent = S.f.period ? `${mod.id==='overview'?'Welcome back! Here\'s your overview for':'Showing'} ${periodLabel(S.f.gran,S.f.period)}` : '';
  const [mn,mx]=dataBounds(); $('#stamp').textContent = mn?`Data: ${fmtDay(mn)} – ${fmtDay(mx)}`:'';
  $('#app').innerHTML = `<div style="display:flex;flex-direction:column;gap:18px">${html}</div>`;
  after.forEach(fn=>fn());
  try{ localStorage.setItem(STORE_KEY+'.ui', JSON.stringify({f:S.f,mod:S.mod,extra:S.extra})); }catch(e){}
}

/* ============================ events ============================ */
document.addEventListener('click', e=>{
  const nb=e.target.closest('[data-mod]'); if(nb){ S.mod=nb.dataset.mod; location.hash=S.mod; render(); window.scrollTo({top:0}); return; }
  const fb=e.target.closest('[data-f]'); if(fb){ S.f[fb.dataset.f]=fb.dataset.v; render(); return; }
  const xb=e.target.closest('[data-extra] [data-v]'); if(xb){ S.extra[xb.closest('[data-extra]').dataset.extra]=xb.dataset.v; render(); return; }
});
document.addEventListener('change', e=>{
  if(e.target.id==='fPeriod'){ S.f.period=e.target.value; render(); }
  else if(e.target.id==='fCompare'){ S.f.compare=e.target.value; render(); }
  else if(e.target.dataset.extra){ S.extra[e.target.dataset.extra]=e.target.value; render(); }
});
// heatmap tooltip
const tip=$('#tip');
document.addEventListener('mousemove', e=>{
  const t=e.target.closest&&e.target.closest('[data-tip],[data-tiphtml]');
  if(!t){ tip.style.display='none'; return; }
  fillTip(t);
  const x=Math.min(e.clientX+14, window.innerWidth-tip.offsetWidth-8), y=Math.min(e.clientY+14, window.innerHeight-tip.offsetHeight-8);
  tip.style.left=x+'px'; tip.style.top=y+'px';
});
// keyboard focus and taps (phones have no hover) show the same tip under the element
function fillTip(el){ const rich=el.dataset.tiphtml!=null; tip.classList.toggle('rich',rich); if(rich) tip.innerHTML=el.dataset.tiphtml; else tip.textContent=el.dataset.tip; tip.style.display='block'; }
function tipAt(el){ fillTip(el); const r=el.getBoundingClientRect();
  tip.style.left=Math.max(8,Math.min(r.left, window.innerWidth-tip.offsetWidth-8))+'px';
  tip.style.top=Math.min(r.bottom+6, window.innerHeight-tip.offsetHeight-8)+'px'; }
document.addEventListener('focusin', e=>{ const t=e.target.closest&&e.target.closest('.hov[data-tiphtml]'); if(t) tipAt(t); });
document.addEventListener('focusout', ()=>{ tip.style.display='none'; });
document.addEventListener('click', e=>{ const t=e.target.closest&&e.target.closest('.hov[data-tiphtml]'); if(t) tipAt(t); });
window.addEventListener('hashchange',()=>{ const h=location.hash.slice(1); if(MODULES.some(m=>m.id===h)&&h!==S.mod){ S.mod=h; render(); } });
if(window.matchMedia) matchMedia('(prefers-color-scheme: dark)').addEventListener('change',()=>render());

/* ---- upload modal ---- */
const modal=$('#modal'), log=$('#log');
function openAdmin(){ modal.classList.add('open'); log.textContent=''; $('#grabToggle').checked=!!CFG.features.grab;
  $('#grabScope').textContent=SHARED.db?'Grab Cost module and Service filter · applies to everyone who opens this link':'Grab Cost module and Service filter · saved in this browser'; }
document.addEventListener('click',e=>{ if(e.target.closest&&e.target.closest('[data-admin]')) openAdmin(); });
$('#mClose').onclick=()=>modal.classList.remove('open');
modal.addEventListener('click',e=>{ if(e.target===modal) modal.classList.remove('open'); });
document.addEventListener('keydown',e=>{ if(e.key==='Escape') modal.classList.remove('open'); });
$('#schemaHelp').innerHTML = `<p class="muted" style="font-size:12px">Trip log: any sheet with columns <code>GuestID, Type, Customer Tier</code> (or <code>Level</code>), <code>Nationality</code> (or <code>Nation</code>), <code>Date, Depart, Arrival, Occupancy, Distance, Arrangby, EmployeeId, Brand, Destination, CITY</code>. Optional: <code>Confirm</code> = name of the Arrangby host. Driver names: a sheet <code>EmployeeId | Name</code> (or <code>Driver Name</code>).</p><b>Driver Roster</b>${ROSTER_SCHEMA}<b>Car Cost</b>${COST_SCHEMA}${CFG.features.grab?`<b>Grab</b>${GRAB_SCHEMA}`:''}`;
async function handleFiles(files){
  const sheets=[]; log.textContent='Reading…';
  for(const f of files){
    try{
      const wb=XLSX.read(await f.arrayBuffer(),{type:'array'});
      wb.SheetNames.forEach(n=>sheets.push({name:n, aoa:XLSX.utils.sheet_to_json(wb.Sheets[n],{header:1,raw:true,blankrows:false})}));
    }catch(err){ log.textContent=`✗ Could not read ${f.name}: ${err.message}`; return; }
  }
  const r=parseWorkbookSheets(sheets);
  if(!r.trips.length&&!r.costs.length&&!r.roster.length&&!Object.keys(r.drivers).length){ log.textContent='✗ No trip log or car-cost sheet recognised. Check the column names (see "Accepted sheet formats").'; return; }
  if(SHARED.db&&SHARED.canWrite){
    log.textContent=r.log.join('\n')+'\n\nSaving for everyone…';
    try{ await shareUpload(r); }
    catch(err){ log.textContent=r.log.join('\n')+'\n\n✗ Could not save the shared data ('+(err&&err.code||'error')+'). '+(err&&err.code==='quota_exceeded'?'Storage is full: remove an older upload below, then try again.':'Please try again.'); return; }
    const latest=r.trips.length? r.trips.map(t=>t.d).sort().pop() : null;
    if(latest){ S.f.gran='month'; S.f.period=latest.slice(0,7); }
    log.textContent=r.log.join('\n')+'\n\nSaved. Everyone who opens this link now sees this data.';
    return; // the shared snapshot re-renders the dashboard
  }
  mergeData(r); S.source='uploaded '+new Date().toLocaleDateString('en-GB'); saveLocal();
  const latest=r.trips.length? r.trips.map(t=>t.d).sort().pop() : null;
  if(latest){ S.f.gran='month'; S.f.period=latest.slice(0,7); }
  log.textContent=r.log.join('\n')+'\n\nDashboard updated.';
  render();
}
const drop=$('#drop'), fileIn=$('#file');
fileIn.onchange=()=>{ if(fileIn.files.length) handleFiles([...fileIn.files]); fileIn.value=''; };
['dragenter','dragover'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add('over');}));
['dragleave','drop'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove('over');}));
drop.addEventListener('drop',e=>{ if(e.dataTransfer.files.length) handleFiles([...e.dataTransfer.files]); });
$('#btnReset').onclick=()=>{ try{ localStorage.removeItem(STORE_KEY); }catch(e){} loadSeed(); S.f.period=null; log.textContent='Reset to built-in data.'; render(); };
$('#btnSave').onclick=()=>{
  // Re-embed current data in the seed as trip-log rows, so the saved file goes through the same parser
  const hdr=['GuestID','Type','Customer Tier','Nationality','Date','Depart','Occupancy','Distance','Arrangby','Confirm','Host Team','EmployeeId','Brand','Destination','CITY','Cost (USD)'];
  const toSerial=iso=>(dOf(iso)-Date.UTC(1899,11,30))/864e5;
  const rows=src=>[hdr,...S.trips.filter(t=>t.src===src).map(t=>[t.gid,t.type,t.tier,t.nat,toSerial(t.d),t.dep/1440,t.dur/1440,t.km,t.host,t.hname,t.team,t.drv,t.car==='Grab'?null:t.car,t.dest,t.city,t.cost])];
  const sheets=[{name:'Vegas trips',aoa:rows('V')}];
  if(S.trips.some(t=>t.src==='G')) sheets.push({name:'Grab trips',aoa:rows('G')});
  if(S.roster.length) sheets.push({name:'Driver Roster',aoa:[['Date','EmployeeId','Start','End','Break','Status'],...S.roster.map(r=>[(dOf(r.d)-Date.UTC(1899,11,30))/864e5,r.drv,r.off?null:r.s/1440,r.off?null:(r.e%1440)/1440,r.brk,r.off?'Off':''])]});
  if(Object.keys(S.drivers).length) sheets.push({name:'Drivers',aoa:[['EmployeeId','Name'],...Object.entries(S.drivers)]});
  if(S.costs.length) sheets.push({name:'Car Cost',aoa:[['Date','Vehicle','Category','Amount (USD)'],...S.costs.map(c=>[costDay(c),c.car,c.cat,c.amt])]});
  const clone=document.documentElement.cloneNode(true);
  clone.querySelector('#app').innerHTML=''; clone.querySelector('#nav').innerHTML=''; clone.querySelector('#filters').innerHTML='';
  clone.querySelector('#modal').classList.remove('open'); clone.querySelector('#log').textContent=''; clone.querySelector('#tip').removeAttribute('style');
  clone.querySelector('#schemaHelp').innerHTML='';
  clone.querySelector('#seed').textContent=JSON.stringify({label:S.source.replace(/^uploaded/,'saved'),sheets}).replace(/</g,'\\u003c');
  const blob=new Blob(['<!doctype html>\n'+clone.outerHTML],{type:'text/html'});
  const [,mx]=dataBounds(); const a=document.createElement('a');
  a.href=URL.createObjectURL(blob); a.download=`Car_Service_Dashboard_${(mx||'').slice(0,7)}.html`; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(()=>URL.revokeObjectURL(a.href),2000);
};

/* ---- Grab on/off (Admin panel) ---- */
function applyGrab(on){
  if(!!CFG.features.grab===!!on) return;
  CFG.features.grab=!!on; buildModules();
  S.f.service=on?'all':'V';
  if(!MODULES.some(m=>m.id===S.mod)) S.mod='overview';
  render();
}
$('#grabToggle').addEventListener('change', async e=>{
  const on=e.target.checked; applyGrab(on);
  if(SHARED.db&&SHARED.canWrite){
    try{ await SHARED.db.doc('settings/features').set({grab:on,at:new Date().toISOString()}); log.textContent=on?'Grab data is now shown to everyone.':'Grab data is now hidden for everyone.'; }
    catch(err){ log.textContent='✗ Could not save the setting ('+(err&&err.code||'error')+'). Please try again.'; e.target.checked=!on; applyGrab(!on); }
  } else { try{ localStorage.setItem(STORE_KEY+'.grab',on?'1':'0'); }catch(_){} log.textContent=on?'Grab data shown in this browser.':'Grab data hidden in this browser.'; }
});
/* ============================ boot ============================ */
if(window.__ARTIFACT__){ // hosted page: browser downloads are blocked, so offer upload only
  $('#btnSave').hidden=true;
  const t=document.querySelector('#modal p.muted'); if(t) t.textContent='Uploaded data is remembered in this browser only. To update the shared page for OD, send the new Excel file to the dashboard owner.';
}
if(!window.__ARTIFACT__){ S.adminOK=true; try{ if(localStorage.getItem(STORE_KEY+'.grab')==='1'){ CFG.features.grab=true; buildModules(); } }catch(_){} }
loadSeed();
SHARED.seed={trips:S.trips.slice(),costs:S.costs.slice(),roster:S.roster.slice(),drivers:Object.assign({},S.drivers),source:S.source};
if(!window.__ARTIFACT__) loadLocal(); // offline file keeps per-browser uploads; hosted page uses the shared store
try{ const ui=JSON.parse(localStorage.getItem(STORE_KEY+'.ui')||'null'); if(ui){ Object.assign(S.f,ui.f||{}); S.mod=ui.mod||S.mod; S.extra=ui.extra||{}; } }catch(e){}
const h=location.hash.slice(1); if(MODULES.some(m=>m.id===h)) S.mod=h;
if(!MODULES.some(m=>m.id===S.mod)) S.mod='overview';
if(!CFG.features.grab) S.f.service='V'; // Vegas Car only while Grab is hidden
Chart.defaults.font.family='"Plus Jakarta Sans",system-ui,sans-serif';
render();
initShared();
})();
