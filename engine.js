/* NECD: exact lab formats; browser-local parsing, numeric copy & OOXML export. */
'use strict';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const state={year:2025,template:null,files:[],records:[],submissions:[],manualSites:{},audit:[],category:'air',method:'file',manualType:'O3',busy:false};
const $ = id => document.getElementById(id);
const escapeHtml = s => String(s??'').replace(/[&<>"']/g,x=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[x]));
const normalized = s => String(s??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim().replace(/\s+/g,' ');
const numeric = v => (v===null || v===undefined || String(v).trim()==='') ? null :
  (typeof v === 'number' ? (Number.isFinite(v)?v:null) : (Number.isFinite(Number(String(v).trim().replace(',','.')))?Number(String(v).trim().replace(',','.')):null));
const mean = a => {const x=a.filter(v=>v!==null && Number.isFinite(v)); return x.length?x.reduce((s,v)=>s+v,0)/x.length:null;};
const cv = a => {const x=a.filter(v=>v!==null && Number.isFinite(v)); if(x.length<2)return null;const m=mean(x); return m===0?null:Math.sqrt(x.reduce((s,v)=>s+(v-m)**2,0)/(x.length-1))/Math.abs(m)*100;};
const asExcelDate = (iso,styleId)=>({__excelDate:(Date.parse(iso+'T00:00:00Z')-Date.UTC(1899,11,30))/86400000,__styleId:styleId});
const colNumber = s => [...s].reduce((a,c)=>a*26+c.charCodeAt(0)-64,0);
const colName = n => {let s='';while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26);}return s;};
const cellParts = ref => {const m=/^([A-Z]+)(\d+)$/.exec(ref);return m?{col:colNumber(m[1]),row:+m[2]}:null;};
const round = (n,k=4) => Number.isFinite(n)?String(Number(n.toFixed(k))):'';
function parseXml(s){const d=new DOMParser().parseFromString(s,'application/xml');if(d.getElementsByTagName('parsererror').length)throw new Error('XML Excel invalid');return d;}
function children(el,name){return [...el.childNodes].filter(n=>n.nodeType===1 && (!name||n.localName===name));}
function first(el,name){return children(el,name)[0]||null;}
function readCell(c,shared){
 const type=c.getAttribute('t')||'';const value=first(c,'v')?.textContent??'';
 if(type==='s')return shared[Number(value)]??'';
 if(type==='inlineStr')return first(c,'is')?.textContent??'';
 if(type==='b')return value==='1';
 if(type==='str'||type==='e')return value;
 if(value==='')return '';
 return Number.isFinite(Number(value)) ? Number(value) : value;
}
function resolveWorkbookPath(target){let p=target.replace(/^\//,'');if(p.startsWith('xl/'))return p;return 'xl/'+p.replace(/^\.\//,'');}
async function workbookFromBuffer(buf){
 const zip=await JSZip.loadAsync(buf);
 const w=zip.file('xl/workbook.xml');const r=zip.file('xl/_rels/workbook.xml.rels');
 if(!w||!r)throw new Error('Fișierul nu este un registru XLSX valid.');
 const doc=parseXml(await w.async('string')), relDoc=parseXml(await r.async('string'));
 const relMap={};for(const e of relDoc.getElementsByTagName('*'))if(e.localName==='Relationship')relMap[e.getAttribute('Id')]=resolveWorkbookPath(e.getAttribute('Target')||'');
 const paths={};for(const e of doc.getElementsByTagName('*'))if(e.localName==='sheet')paths[e.getAttribute('name')]=relMap[e.getAttributeNS(REL,'id')||e.getAttribute('r:id')];
 let shared=[];
 if(zip.file('xl/sharedStrings.xml')){
  const sh=parseXml(await zip.file('xl/sharedStrings.xml').async('string'));
  shared=[...sh.getElementsByTagNameNS(NS,'si')].map(e=>e.textContent??'');
 }
 return {zip,paths,shared,sheetCache:{}};
}
async function loadSheet(wb,name){
 if(wb.sheetCache[name])return wb.sheetCache[name];
 const path=wb.paths[name];if(!path || !wb.zip.file(path))throw new Error(`Lipsește foaia ${name}.`);
 const doc=parseXml(await wb.zip.file(path).async('string'));
 const rows=new Map(), data=doc.getElementsByTagNameNS(NS,'sheetData')[0];
 for(const row of children(data,'row')){
  const n=Number(row.getAttribute('r'));const obj={};
  for(const c of children(row,'c')){
   const ref=c.getAttribute('r');if(ref)obj[ref.replace(/\d+$/,'')]=readCell(c,wb.shared);
  }
  rows.set(n,obj);
 }
 const output={path,doc,rows};wb.sheetCache[name]=output;return output;
}
function parseDate(v){
 if(typeof v==='number'&&Number.isFinite(v) && v>1000 && v<100000){return new Date(Date.UTC(1899,11,30)+Math.floor(v)*86400000).toISOString().slice(0,10);}
 const s=String(v??'').trim();let m;
 if((m=/^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s)))return validDate(+m[1],+m[2],+m[3]);
 if((m=/^(\d{1,2})[.\/\-](\d{1,2})[.\/\-](\d{2,4})$/.exec(s))){const yr=+m[3]<100?2000+(+m[3]):+m[3];return validDate(yr,+m[2],+m[1]);}
 return null;
}
function validDate(y,m,d){if(y<1900||y>2150)return null;const dt=new Date(Date.UTC(y,m-1,d));return dt.getUTCFullYear()===y && dt.getUTCMonth()===m-1 && dt.getUTCDate()===d?dt.toISOString().slice(0,10):null;}
function dateDMY(s){return s?`${s.slice(8,10)}.${s.slice(5,7)}.${s.slice(0,4)}`:'';}
function dateTimeMs(date,time){if(!date)return null;const t=numeric(time);return Date.parse(date+'T00:00:00Z')+(t!==null?t*86400000:0);}
function issue(severity,text){return {severity,text};}
function csvRows(input){
 input=input.replace(/^\uFEFF/,'');const sep=(input.split(/\r?\n/,1)[0].match(/;/g)||[]).length>(input.split(/\r?\n/,1)[0].match(/,/g)||[]).length?';':',';
 let rows=[],row=[],field='',quoted=false;
 for(let i=0;i<input.length;i++){
  const c=input[i];
  if(c==='"'){if(quoted && input[i+1]==='"'){field+='"';i++;}else quoted=!quoted;}
  else if(c===sep && !quoted){row.push(field);field='';}
  else if((c==='\n'||c==='\r')&&!quoted){if(c==='\r'&&input[i+1]==='\n')i++;row.push(field);if(row.some(x=>x!==''))rows.push(row);row=[];field='';}
  else field+=c;
 }
 if(field!==''||row.length){row.push(field);if(row.some(x=>x!==''))rows.push(row);}
 return rows;
}
function recordBase(file,sheet,row,kind,site,year){return {id:`${file.id}/${sheet}/${row}`,fileId:file.id,filename:file.name,sheet,sourceRow:row,kind,originSite:String(site??'').trim(),year:numeric(year),metadataIssues:[]};}
async function parseAir(wb,file){
 const rec=[];
 for(const [sheet,parameter,repcols,finalCol,cvCol] of [['Ozon','O3',['R','S'],'T','V'],['NO2','NO2',['Q','R'],null,'S'],['NH3','NH3',['M','N'],null,'O']]){
  const sh=await loadSheet(wb,sheet);const h=sh.rows.get(1)||{};
  const expect=sheet==='Ozon'?'Conc O3':sheet==='NO2'?'NO2':'Conc NH3';
  if(!Object.values(h).some(x=>String(x).includes(expect)))throw Error(`Foaia ${sheet} nu are structura de laborator recunoscută.`);
  for(const [i,r] of sh.rows){if(i<2||!String(r.B??'').trim())continue;
   const x=recordBase(file,sheet,i,'air',r.B,r.A);
   Object.assign(x,{parameter,exposure:r.C,start:parseDate(r.F),end:parseDate(r.H),startTime:r.G,endTime:r.I,
       hours:numeric(r.J),replicates:repcols.map(c=>numeric(r[c])),sourceAverage:finalCol?numeric(r[finalCol]):null,
       sourceCV:numeric(r[cvCol]),primitive:[numeric(r.K),numeric(r.L)],filterCodes:[r.D,r.E]});
   rec.push(x);
  }
 }
 return rec;
}
async function parseIons(wb,file){
 const sh=await loadSheet(wb,'Sheet1');const h=sh.rows.get(1)||{};
 if(normalized(h.A)!=='ssi'||normalized(h.B)!=='nivel'||normalized(h.G)!=='nnh4')throw Error('Structură ionică neașteptată (SSI, Nivel, NNH4).');
 const rec=[];
 for(const [row,r] of sh.rows){if(row<2||!String(r.A??'').trim())continue;
  const x=recordBase(file,'Sheet1',row,'ions',r.A,null);
  x.level=String(r.B??'').trim(); x.ions={Ca:numeric(r.C),Mg:numeric(r.D),Na:numeric(r.E),K:numeric(r.F),NH4:numeric(r.G),SO4:numeric(r.H),NO3:numeric(r.I),Cl:numeric(r.J)};
  x.invalidNumbers=[];for(const [k,c] of Object.entries({Ca:'C',Mg:'D',Na:'E',K:'F',NH4:'G',SO4:'H',NO3:'I',Cl:'J'}))if(String(r[c]??'').trim()!==''&&numeric(r[c])===null)x.invalidNumbers.push(k);
  rec.push(x);
 }
 return rec;
}
function parseNutrition(file,text){
 const rows=csvRows(text);if(!rows.length)throw Error('CSV gol');
 const headers=rows[0].map(x=>normalized(x));const map={};headers.forEach((h,i)=>map[h]=i);
 for(const field of ['ssi','specia','an','n','p','ca','mg','k','c'])if(map[field]===undefined)throw Error(`CSV nu conține coloana ${field}.`);
 const rec=[];
 for(let row=1;row<rows.length;row++){
  const r=rows[row];if(!r.some(x=>x.trim()))continue;
  const get=k=>r[map[k]]??'';
  const x=recordBase(file,'CSV',row+1,'nutrition',get('ssi'),get('an'));
  x.species=String(get('specia')).trim();x.nutrients={C:numeric(get('c')),N:numeric(get('n')),P:numeric(get('p')),Ca:numeric(get('ca')),Mg:numeric(get('mg')),K:numeric(get('k'))};
  x.micronutrients={Mn:numeric(get('mn')),Zn:numeric(get('zn')),Fe:numeric(get('fe')),Cu:numeric(get('cu'))};
  x.invalidNumbers=[];for(const k of ['c','n','p','ca','mg','k','mn','zn','fe','cu'])if(map[k]!==undefined&&String(get(k)).trim()!==''&&numeric(get(k))===null)x.invalidNumbers.push(k);
  rec.push(x);
 }
 return rec;
}

function xmlCell(doc,col,row,value,old){
 const c=doc.createElementNS(NS,'c');c.setAttribute('r',col+row);
 const prevStyle=value&&typeof value==='object'&&'__excelDate' in value?String(value.__styleId):old?.getAttribute('s');
 if(prevStyle!==null && prevStyle!==undefined)c.setAttribute('s',prevStyle);
 if(value && typeof value==='object' && '__excelDate' in value){
  const v=doc.createElementNS(NS,'v');v.textContent=String(value.__excelDate);c.appendChild(v);
 }else if(typeof value==='number' && Number.isFinite(value)){
  const v=doc.createElementNS(NS,'v');v.textContent=String(value);c.appendChild(v);
 }else{
  c.setAttribute('t','inlineStr');const is=doc.createElementNS(NS,'is'), t=doc.createElementNS(NS,'t');
  const s=String(value??'');if(/^\s|\s$/.test(s))t.setAttribute('xml:space','preserve');t.textContent=s;is.appendChild(t);c.appendChild(is);
 }
 return c;
}
function editWorksheetXml(xml,updates){
 if(!updates.size)return xml;
 const doc=parseXml(xml), sheetData=doc.getElementsByTagNameNS(NS,'sheetData')[0];
 if(!sheetData)throw Error('SheetData lipsă din șablon.');
 const rowMap=new Map();for(const r of children(sheetData,'row'))rowMap.set(+r.getAttribute('r'),r);
 let maxRow=0,maxCol=0;
 for(const [rowNum,colVals] of [...updates.entries()].sort((a,b)=>a[0]-b[0])){
  let row=rowMap.get(rowNum);
  if(!row){row=doc.createElementNS(NS,'row');row.setAttribute('r',String(rowNum));
   const next=[...rowMap.keys()].filter(v=>v>rowNum).sort((a,b)=>a-b)[0];
   if(next)sheetData.insertBefore(row,rowMap.get(next));else sheetData.appendChild(row);rowMap.set(rowNum,row);
  }
  const existing=new Map(children(row,'c').map(c=>[cellParts(c.getAttribute('r')).col,c]));
  for(const [col,value] of Object.entries(colVals)){
   if(value===undefined||value===null||value==='')continue;
   const num=colNumber(col);maxCol=Math.max(maxCol,num);maxRow=Math.max(maxRow,rowNum);
   const newC=xmlCell(doc,col,rowNum,value,existing.get(num));const old=existing.get(num);if(old){row.replaceChild(newC,old);existing.set(num,newC);}
   else{
    const next=[...existing.keys()].filter(x=>x>num).sort((a,b)=>a-b)[0];
    if(next)row.insertBefore(newC,existing.get(next));else row.appendChild(newC);existing.set(num,newC);
   }
  }
 }
 const dim=doc.getElementsByTagNameNS(NS,'dimension')[0];
 if(dim){let orig=dim.getAttribute('ref')||'A1';let end=orig.split(':').pop();let part=cellParts(end);
  if(part){maxRow=Math.max(maxRow,part.row);maxCol=Math.max(maxCol,part.col);dim.setAttribute('ref',`A1:${colName(maxCol)}${maxRow}`);}}
 return new XMLSerializer().serializeToString(doc);
}
async function ensureDateStyle(wb){
 const file=wb.zip.file('xl/styles.xml');if(!file)throw Error('Nu se găsește xl/styles.xml');
 const doc=parseXml(await file.async('string')),root=doc.documentElement;
 let numFmts=children(root,'numFmts')[0];
 if(!numFmts){numFmts=doc.createElementNS(NS,'numFmts');root.insertBefore(numFmts,root.firstChild);}
 let format=[...children(numFmts,'numFmt')].find(n=>/dd[.\/-]mm[.\/-]yyyy/i.test(n.getAttribute('formatCode')||''));
 if(!format){
  const max=Math.max(163,...children(numFmts,'numFmt').map(n=>Number(n.getAttribute('numFmtId'))||0));
  format=doc.createElementNS(NS,'numFmt');format.setAttribute('numFmtId',String(max+1));format.setAttribute('formatCode','dd.mm.yyyy');numFmts.appendChild(format);
  numFmts.setAttribute('count',String(children(numFmts,'numFmt').length));
 }
 const xfs=children(root,'cellXfs')[0];if(!xfs)throw Error('Nu se găsește cellXfs');
 let style=children(xfs,'xf').find(n=>n.getAttribute('numFmtId')===format.getAttribute('numFmtId'));
 let index;
 if(style)index=children(xfs,'xf').indexOf(style);
 else{
  style=doc.createElementNS(NS,'xf');
  for(const [k,v] of Object.entries({numFmtId:format.getAttribute('numFmtId'),fontId:'0',fillId:'0',borderId:'0',xfId:'0',applyNumberFormat:'1'}))style.setAttribute(k,v);
  index=children(xfs,'xf').length;xfs.appendChild(style);xfs.setAttribute('count',String(index+1));
 }
 wb.zip.file('xl/styles.xml',new XMLSerializer().serializeToString(doc));
 return index;
}

const CATEGORIES=[
 {id:'O3',label:'Ozon (O₃)',kind:'air'},
 {id:'NO2',label:'Dioxid de azot (NO₂)',kind:'air'},
 {id:'NH3',label:'Amoniac (NH₃)',kind:'air'},
 {id:'soil',label:'Soluția solului · ioni',kind:'ions'},
 {id:'collector',label:'Colectori TL / SC · ioni',kind:'ions'},
 {id:'nutrition',label:'Nutriție foliară',kind:'nutrition'}
];
const LAB_CATEGORIES=[
 {id:'air',label:'Concentrații atmosferice · O₃, NO₂, NH₃',subtitle:'Fișiere MI_FINAL / LM_FINAL (.xlsx)'},
 {id:'ions',label:'Concentrații ionice · sol și colectori',subtitle:'Fișier Conc_ioni (.xlsx)'},
 {id:'nutrition',label:'Nutriție foliară',subtitle:'Fișier Nutritie (.csv)'}
];
const NUTRIENTS=['C','N','P','Ca','Mg','K','Mn','Zn','Fe','Cu'];
const IONS=['Ca','Mg','Na','K','NH4','SO4','NO3','Cl'];
function groupOf(x){if(x.kind==='air')return x.parameter;if(x.kind==='nutrition')return 'nutrition';if(x.kind==='ions')return /^(tl|sc)$/.test(normalized(x.level))?'collector':'soil';return 'unknown';}
function fmt(n){return Number.isFinite(n)?String(Number(n.toPrecision(9))):'—';}
function clockDate(x){return x?dateDMY(x):'—';}
function html(s){return escapeHtml(s);}
function notice(msg,type='info'){const el=$('message');el.textContent=msg;el.dataset.level=type;}
function now(){return new Date().toISOString();}
function log(action,description){state.audit.push({when:now(),action,description});if(state.audit.length>2000)state.audit=state.audit.slice(-2000);}
function isWellFormed(v){return numeric(v)!==null;}
function station(raw){
 const rawName=String(raw??'').trim();const key=normalized(rawName);
 const override=state.manualSites[key];
 if(override&&state.template)return state.template.sites.find(s=>s.code===override)||null;
 if(!state.template||!key)return null;
 const aliases={stefanesti:'stefanesti',mihaesti:'mihaesti',baragan:'baraganu',braganu:'baraganu',
  demacusa:'demacusa','poiana stanii':'poiana stanii','podul cu flori':'podu cu flori'};
 const candidates=state.template.sites.filter(s=>s.norm===(aliases[key]||key));
 if(candidates.length===1)return candidates[0];
 const partial=state.template.sites.filter(s=>s.norm.endsWith(' '+key)||s.norm.startsWith(key+' '));
 return partial.length===1?partial[0]:null;
}
function diagnose(x){
 const warnings=[];const grp=groupOf(x);const site=station(x.originSite);let fields={},extra={};
 const year=Number.isFinite(x.year)&&x.year>=1900&&x.year<=2100?x.year:state.year;
 if(!site)warnings.push(`Stația „${x.originSite||'nespecificată'}” nu are cod confirmat în șablon; se păstrează denumirea originală.`);
 if(Number.isFinite(x.year)&&x.year!==state.year)warnings.push(`An ${x.year}, diferit de anul raportului (${state.year}).`);
 let start=x.start||null,end=x.end||null;
 if(grp==='O3'||grp==='NO2'||grp==='NH3'){
  let value=null;
  if(grp==='O3' && x.sourceAverage!==null && Number.isFinite(x.sourceAverage))value=x.sourceAverage;
  else value=mean((x.replicates||[]).filter(Number.isFinite));
  if(value===null)return {x,group:grp,site,year,start,end,warnings:[...warnings,'Nu există nicio concentrație numerică finală.'],status:'invalid',fields,extra};
  if(value<0)warnings.push('Valoare negativă; verifică rezultatul furnizat de laborator.');
  if(value>1000)warnings.push('Concentrație neobișnuit de mare; verifică introducerea.');
  if(!start||!end)warnings.push('Lipsește una dintre datele intervalului de expunere.');
  if(start&&end&&start>end)warnings.push('Datele intervalului sunt inversate.');
  if(grp!=='O3'&&(!x.replicates||x.replicates.filter(Number.isFinite).length===1))warnings.push('O singură replică numerică finală disponibilă.');
  fields[{O3:'DK',NO2:'DI',NH3:'DH'}[grp]]=value;
 }else if(grp==='nutrition'){
  const species=({molid:'Picea abies',fag:'Fagus sylvatica',gorun:'Quercus petraea'})[normalized(x.species)]||x.species||'';
  if(!species)warnings.push('Specia lipsește.');
  if(normalized(x.species)==='stejar')warnings.push('Denumirea „Stejar” este preluată exact din laborator, fără identificarea speciei latine.');
  fields.CQ=species;
  for(const [p,col] of Object.entries({C:'CR',N:'CS',P:'CT',Ca:'CU',Mg:'CV',K:'CW'})){
    const v=x.nutrients?.[p];if(v!==null&&Number.isFinite(v))fields[col]=v;
    else warnings.push(`Valoare lipsă/numeric invalid pentru ${p}.`);
  }
  for(const [p,col] of Object.entries({Mn:'T',Zn:'U',Fe:'V',Cu:'W'})){
    const v=x.micronutrients?.[p];if(v!==null&&Number.isFinite(v))extra[col]=v;
  }
  if(!Object.keys(fields).some(k=>k!=='CQ'))return {x,group:grp,site,year,start,end,warnings:[...warnings,'Nu există valori numerice de nutriție.'],status:'invalid',fields,extra};
  if(Object.values(fields).some(v=>typeof v==='number'&&v<0)||Object.values(extra).some(v=>v<0))warnings.push('Conține valori negative.');
  if(x.invalidNumbers?.length)warnings.push(`Valori nenumerice: ${x.invalidNumbers.join(', ')}.`);
  if(Object.keys(extra).length)extra.S=species;
 }else if(grp==='soil'||grp==='collector'){
  if(grp==='soil'){
   const m=/^(\d+)\s*cm$/.exec(normalized(x.level));
   if(m)fields.W=Number(m[1]);else warnings.push(`Adâncimea „${x.level}” nu este de tipul 30 cm, 60 cm etc.`);
   for(const [p,col] of Object.entries({NH4:'BG',NO3:'BH',SO4:'BI',Cl:'BJ',Ca:'BP',Mg:'BQ',Na:'BR',K:'BS'})){
    const v=x.ions?.[p];if(v!==null&&Number.isFinite(v))fields[col]=v;
   }
  }else{
   extra.K=/^tl$/.test(normalized(x.level))?'TL':'SC';
   for(const [p,col] of Object.entries({Ca:'L',Mg:'M',Na:'N',K:'O',NH4:'P',SO4:'Q',NO3:'R',Cl:'S'})){
    const v=x.ions?.[p];if(v!==null&&Number.isFinite(v))extra[col]=v;
   }
  }
  if(!IONS.some(p=>Number.isFinite(x.ions?.[p])))return {x,group:grp,site,year,start,end,warnings:[...warnings,'Nu există concentrații ionice numerice.'],status:'invalid',fields,extra};
  if(IONS.some(p=>Number.isFinite(x.ions?.[p])&&x.ions[p]<0))warnings.push('Conține concentrații negative.');
  if(x.invalidNumbers?.length)warnings.push(`Valori nenumerice: ${x.invalidNumbers.join(', ')}.`);
  // Source Conc_ioni has neither a reference date nor an interval. Use the reporting year, with no invented dates.
 }else return {x,group:grp,site,year,start,end,warnings:['Tip de analiză necunoscut'],status:'invalid',fields,extra};
 return {x,group:grp,site,year,start,end,warnings,status:warnings.length?'warning':'ok',fields,extra};
}
function evalAll(){
 const items=state.records.map(diagnose).filter(e=>e.year===state.year);const seen=new Map();
 for(const z of items){if(z.status==='invalid')continue;
  const fingerprint=[z.group,z.site?.code||z.x.originSite,z.year,z.start,z.end,z.x.kind==='ions'?z.x.level:'',z.x.kind==='nutrition'?z.x.species:''].join('|');
  const prev=seen.get(fingerprint);
  if(prev&&JSON.stringify(prev.fields)===JSON.stringify(z.fields)&&JSON.stringify(prev.extra)===JSON.stringify(z.extra))z.warnings.push('Posibil duplicat identic cu o altă înregistrare; este păstrat separat.');
  else if(prev)z.warnings.push('Există altă valoare pentru aceeași stație/perioadă; ambele sunt păstrate.');
  seen.set(fingerprint,z);
  if(z.warnings.length)z.status='warning';
 }
 return items;
}
async function loadPrivateTemplate(blob){
 const buf=await blob.arrayBuffer(),wb=await workbookFromBuffer(buf);
 for(const sh of ['(1) Contact','(2) Sites','(4) Terrestrial','(5) Additional'])if(!wb.paths[sh])throw Error(`Șablonul nu are foaia ${sh}.`);
 const sitesSheet=await loadSheet(wb,'(2) Sites'),terr=await loadSheet(wb,'(4) Terrestrial');
 if(terr.rows.get(9)?.DH!=='NH3 concentration'||terr.rows.get(9)?.DK!=='O3 concentration')throw Error('Versiune de șablon NECD incompatibilă.');
 const sites=[];for(const [row,r] of sitesSheet.rows)if(row>=13&&r.C&&r.D)sites.push({row,code:String(r.C).trim(),name:String(r.D).trim(),norm:normalized(r.D)});
 if(!sites.length)throw Error('Șablonul nu conține lista stațiilor.');
 state.template={name:'NECD art9 Reporting_06_26.xlsx',buffer:buf,sites};render();
}
async function importFile(f){
 const buf=await f.arrayBuffer(),id=await hashBuffer(buf);
 if(state.files.some(x=>x.id===id))throw Error('Acest fișier a fost deja importat (conținut identic).');
 const file={id,name:f.name,when:now(),size:f.size,sourceCategory:state.category};
 let results=[];
 if(/\.csv$/i.test(f.name)){if(state.category!=='nutrition')throw Error('Un CSV de nutriție trebuie importat la Nutriție foliară.');results=parseNutrition(file,new TextDecoder('utf-8').decode(buf));}
 else if(/\.xlsx$/i.test(f.name)){
  const wb=await workbookFromBuffer(buf);
  if(state.category==='air'){
    if(!['Ozon','NO2','NH3'].every(s=>s in wb.paths))throw Error('Pentru aer aștept foile originale Ozon, NO2 și NH3.');
    results=await parseAir(wb,file);
  }else if(state.category==='ions'){
    if(!('Sheet1' in wb.paths))throw Error('Aștept fișierul original Conc_ioni cu Sheet1.');
    results=await parseIons(wb,file);
  }else throw Error('Pentru nutriție aștept CSV-ul original de laborator.');
 }else throw Error('Fișier neacceptat: folosește XLSX sau CSV în formatul laboratorului.');
 if(!results.length)throw Error('Nu am găsit înregistrări în fișier.');
 for(const x of results)if(x.kind==='ions'&&!Number.isFinite(x.year))x.year=state.year;
 file.count=results.length;
 file.categories=[...new Set(results.map(groupOf))];
 return {file,results};
}
async function hashBuffer(buffer){
 if(!globalThis.crypto?.subtle){const a=new Uint8Array(buffer);let h=2166136261;for(const b of a)h=Math.imul(h^b,16777619);return 'fnv-'+(h>>>0).toString(16)+'-'+a.length;}
 const d=await crypto.subtle.digest('SHA-256',buffer);
 return [...new Uint8Array(d)].map(x=>x.toString(16).padStart(2,'0')).join('');
}
function fieldsNumeric(names,fd){const out={};for(const k of names){const v=String(fd.get(k)||'').trim();out[k]=numeric(v);}return out;}
function manualRecord(form){
 const fd=new FormData(form),type=state.manualType,raw=String(fd.get('site')||'').trim(),yr=numeric(fd.get('year'))||state.year;
 if(!raw)throw Error('Completează stația.');
 const record={id:`manual/${globalThis.crypto?.randomUUID?.()||('fallback-'+Date.now()+'-'+Math.random().toString(36).slice(2))}`,fileId:null,filename:'Introducere manuală',sheet:'Manual',sourceRow:0,
   originSite:raw,year:yr,kind:'',manual:true,createdAt:now()};
 if(['O3','NO2','NH3'].includes(type)){
  record.kind='air';record.parameter=type;record.start=parseDate(fd.get('start'));record.end=parseDate(fd.get('end'));
  record.replicates=[numeric(fd.get('rep1')),numeric(fd.get('rep2'))];
  record.sourceAverage=type==='O3'?numeric(fd.get('avg')):null;
  if(record.sourceAverage===null&&!record.replicates.some(Number.isFinite))throw Error('Introdu o valoare numerică pentru concentrație.');
 }else if(type==='nutrition'){
  record.kind='nutrition';record.species=String(fd.get('species')||'').trim();
  record.nutrients=fieldsNumeric(['C','N','P','Ca','Mg','K'],fd);
  record.micronutrients=fieldsNumeric(['Mn','Zn','Fe','Cu'],fd);
  record.invalidNumbers=[];
  for(const k of NUTRIENTS){const v=String(fd.get(k)||'').trim();if(v&&!Number.isFinite(numeric(v)))record.invalidNumbers.push(k);}
 }else if(type==='soil'||type==='collector'){
  record.kind='ions';record.level=type==='soil'?String(fd.get('level')||'30 cm'):String(fd.get('level')||'TL');
  record.ions=fieldsNumeric(IONS,fd);record.invalidNumbers=[];
  for(const k of IONS){const v=String(fd.get(k)||'').trim();if(v&&!Number.isFinite(numeric(v)))record.invalidNumbers.push(k);}
 }
 return record;
}
function cellRow(colvals,z,dateId){
 const row={B:'RO',C:z.site?.code||z.x.originSite,D:z.x.originSite,F:z.year};
 if(z.start)row.H=asExcelDate(z.start,dateId);
 if(z.end)row.I=asExcelDate(z.end,dateId);
 return {...row,...colvals};
}
const apply=(map,row,fields)=>{if(!map.has(row))map.set(row,{});Object.assign(map.get(row),fields);};
async function makeReport(){
 if(!state.template)throw Error('Șablonul Excel NECD din Google Drive nu este disponibil.');
 const entries=evalAll().filter(e=>e.status!=='invalid');
 if(!entries.length){
  const otherYears=[...new Set(state.records.map(x=>Number(x.year)).filter(y=>Number.isInteger(y)&&y!==state.year))].sort();
  const invalid=evalAll();
  if(invalid.length)throw Error('Înregistrările pentru anul '+state.year+' nu conțin valori numerice valide. Verifică atenționările QA/QC.');
  throw Error('Nu există înregistrări pentru anul '+state.year+'.'+(otherYears.length?' Există date pentru: '+otherYears.join(', ')+'.':''));
 }
 const wb=await workbookFromBuffer(state.template.buffer);
 const terrestrial=await loadSheet(wb,'(4) Terrestrial'),additional=await loadSheet(wb,'(5) Additional'),contact=await loadSheet(wb,'(1) Contact'),sites=await loadSheet(wb,'(2) Sites');
 const dateId=entries.some(e=>e.start||e.end)?await ensureDateStyle(wb):null;
 const maxOccupied=(sheet)=>Math.max(13,...[...sheet.rows.entries()].filter(([r,vals])=>r>13&&Object.values(vals).some(v=>v!==''&&v!==null&&v!==undefined)).map(([r])=>r));
 let rt=maxOccupied(terrestrial)+1,ra=maxOccupied(additional)+1;
 const T=new Map(),A=new Map(),SITE=new Map();
 const cols={K:'Tip colector (TL/SC)',L:'Ca (laborator)',M:'Mg (laborator)',N:'Na (laborator)',O:'K (laborator)',P:'NH4 (laborator)',Q:'SO4 (laborator)',R:'NO3 (laborator)',S:'Cl (laborator)',T:'Mn foliar (laborator)',U:'Zn foliar (laborator)',V:'Fe foliar (laborator)',W:'Cu foliar (laborator)',X:'Specie foliară (laborator)'};
 for(const [col,title] of Object.entries(cols))apply(A,9,{[col]:title});
 // No fabricated unit labels or scientific conversion. Values copied from laboratory sources.
 for(const z of entries){
  if(z.group==='collector')apply(A,ra++,{...cellRow({},z,dateId),J:'Terrestrial',...z.extra});
  else {
   apply(T,rt++,cellRow(z.fields,z,dateId));
   if(z.group==='nutrition'&&Object.keys(z.extra).length){const extra={...z.extra};if('S' in extra){extra.X=extra.S;delete extra.S;}apply(A,ra++,{...cellRow({},z,dateId),J:'Terrestrial',...extra});}
  }
  if(z.site){const ff={O3:'FM',NO2:'FK',NH3:'FJ',nutrition:'ES',soil:'DI'}[z.group];
   // Only mark capability of the site when the respective data was actually submitted.
   if(ff)apply(SITE,z.site.row,{[ff]:'Yes'});
  }
 }
 // Skip filling a false date interval if the source did not supply one. Use the selected report year.
 for(const [sheet,updates] of [[terrestrial,T],[additional,A],[sites,SITE]])if(updates.size){
  wb.zip.file(sheet.path,editWorksheetXml(await wb.zip.file(sheet.path).async('string'),updates));
 }
 const contactUpdates=new Map();apply(contactUpdates,7,{D:state.year});
 wb.zip.file(contact.path,editWorksheetXml(await wb.zip.file(contact.path).async('string'),contactUpdates));
 const blob=await wb.zip.generateAsync({type:'blob',compression:'DEFLATE',compressionOptions:{level:6}});
 return {blob,rows:rt-maxOccupied(terrestrial)-1,extra:ra-maxOccupied(additional)-1,count:entries.length,excluded:evalAll().length-entries.length};
}
function auditCsv(){
 const rows=[['status','category','source_file','sheet','row','site','site_code','year','start','end','warnings']];
 for(const z of evalAll())rows.push([z.status,z.group,z.x.filename,z.x.sheet,z.x.sourceRow,z.x.originSite,z.site?.code||'',z.year,z.start||'',z.end||'',z.warnings.join(' | ')]);
 rows.push([]);rows.push(['time','action','description']);for(const a of state.audit)rows.push([a.when,a.action,a.description]);
 const text=rows.map(r=>r.map(v=>'"'+String(v??'').replace(/"/g,'""')+'"').join(',')).join('\r\n');
 return new Blob(['\ufeff'+text],{type:'text/csv;charset=utf-8'});
}
function download(blob,name){const u=URL.createObjectURL(blob),a=document.createElement('a');a.href=u;a.download=name;document.body.append(a);a.click();setTimeout(()=>{URL.revokeObjectURL(u);a.remove();},1200);}
function serializeWorkspace(){return {schema:3,year:state.year,files:state.files,records:state.records,submissions:state.submissions,audit:state.audit,manualSites:state.manualSites,updatedAt:now()};}
function restoreWorkspace(data){
 if(!data||data.schema!==3)throw Error('Structură de date veche. Creează un repository nou sau migrează manual datele existente.');
 state.year=Number(data.year)||2025;state.files=data.files||[];state.records=data.records||[];
 state.submissions=data.submissions||[];state.audit=data.audit||[];state.manualSites=data.manualSites||{};
 // Recover a previous submission whose year differs from the old default report year.
 const years=[...new Set(state.records.map(x=>Number(x.year)).filter(y=>Number.isInteger(y)&&y>=2000&&y<=2100))];
 if(!state.records.some(x=>Number(x.year)===state.year)&&years.length===1)state.year=years[0];
 render();
}
function badgeData(){
 const entries=evalAll();return CATEGORIES.map(cat=>{
  const matches=entries.filter(e=>e.group===cat.id),active=matches.filter(e=>e.status!=='invalid');
  return {...cat,number:active.length,warnings:active.filter(e=>e.status==='warning').length,invalid:matches.filter(e=>e.status==='invalid').length,
   status:active.length?(active.some(e=>e.status==='warning')?'warning':'ok'):'missing'};
 });
}
function render(){
 if(!$('traffic'))return;
 $('traffic').innerHTML=badgeData().map(c=>`<div class="traffic-item"><i class="dot ${c.status}"></i><div><b>${html(c.label)}</b><small>${c.number?`${c.number} înregistrări${c.warnings?' · '+c.warnings+' cu atenționări':''}`:'Nedepus'}${c.invalid?' · '+c.invalid+' fără valoare':''}</small></div></div>`).join('');
 $('projectYear').value=state.year;
 $('stationSuggestions').innerHTML=(state.template?.sites||[]).map(s=>`<option value="${html(s.name)}"></option>`).join('');
 $('sourceCount').textContent=`${state.files.length} fișiere · ${state.records.length} înregistrări`;
 const yrSubmissions=state.submissions.filter(s=>s.year===state.year);
 const otherYears=[...new Set(state.records.map(x=>Number(x.year)).filter(y=>Number.isInteger(y)&&y!==state.year))].sort();
 $('savedInfo').textContent=yrSubmissions.length?`Ultima depunere: ${new Date(yrSubmissions.at(-1).time).toLocaleString('ro-RO')} · ${yrSubmissions.at(-1).what}`:(otherYears.length?`Nicio analiză pentru ${state.year}. Date disponibile pentru: ${otherYears.join(', ')}.`:'Nicio analiză depusă pentru acest an.');
 const d=evalAll();
 $('qualitySummary').textContent=`${d.filter(x=>x.status==='ok').length} fără atenționări · ${d.filter(x=>x.status==='warning').length} cu atenționări · ${d.filter(x=>x.status==='invalid').length} fără valori numerice`;
 $('warnings').innerHTML=d.filter(x=>x.warnings.length).slice(0,80).map(z=>`<div class="warning-line"><span>${html(z.x.originSite)} · ${html(z.group)} (${html(z.x.filename)})</span><small>${html(z.warnings.join(' • '))}</small></div>`).join('')||'<p>Nu există atenționări.</p>';
 $('importList').innerHTML=state.submissions.filter(s=>s.year===state.year).slice().reverse().slice(0,12).map(s=>`<div class="submission"><b>${html(s.what)}</b><small>${new Date(s.time).toLocaleString('ro-RO')} · ${s.count} înregistrări</small></div>`).join('')||'<p class="muted">Nu au fost salvate analize.</p>';
 const methods=$('entryMethod');if(methods)methods.value=state.method;
 const category=$('category');if(category)category.value=state.category;
 $('entryFile').hidden=state.method!=='file';$('entryManual').hidden=state.method!=='manual';
 $('labFile').accept=state.category==='nutrition'?'.csv':'.xlsx';
 $('fileHint').textContent=LAB_CATEGORIES.find(x=>x.id===state.category)?.subtitle||'';
 $('exportBtn').disabled=!state.template||!d.some(x=>x.status!=='invalid')||!window.NECD_CLOUD?.isConnected();
 $('csvBtn').disabled=!d.length;
}
function manualLayout(){
 const subtype=$('manualType');
 let opts=state.category==='air'?CATEGORIES.slice(0,3):state.category==='ions'?CATEGORIES.slice(3,5):CATEGORIES.slice(5);
 subtype.innerHTML=opts.map(o=>`<option value="${o.id}">${html(o.label)}</option>`).join('');
 state.manualType=opts[0].id;subtype.value=state.manualType;drawManualFields();
}
function inputField(name,label,placeholder='',type='text',required=false){return `<label><span>${label}</span><input name="${name}" ${type==='date'?'type="date"':'type="text"'} placeholder="${html(placeholder)}" ${required?'required':''} autocomplete="off"><small class="input-error" hidden></small></label>`;}
function drawManualFields(){
 let f='';const type=state.manualType;
 if(['O3','NO2','NH3'].includes(type)){
  f+=`<p class="muted">Concentrații finale în µg/m³. Pentru O₃ poți introduce direct valoarea finală (Avg). Pentru NO₂ și NH₃ introdu valorile finale ale replicilor.</p>`;
  f+=inputField('start','Data expunerii','','date')+inputField('end','Data colectării','','date');
  if(type==='O3')f+=inputField('avg','Concentrația finală O₃ (Avg)','48,79');
  f+=inputField('rep1','Concentrație finală, replica 1','0,792')+inputField('rep2','Concentrație finală, replica 2','0,838');
 }else if(type==='nutrition'){
  f+=inputField('species','Specie','Molid / Fag / Gorun / Stejar', 'text',true);
  for(const n of NUTRIENTS)f+=inputField(n,n,'12,25');
 }else if(type==='soil'||type==='collector'){
  f+=type==='collector'?`<label><span>Tip colector</span><select name="level"><option value="TL">TL</option><option value="SC">SC</option></select></label>`:inputField('level','Adâncime / nivel','30 cm');
  for(const n of IONS)f+=inputField(n,n,'0,56');
 }
 $('manualFields').innerHTML=f;
 const first=$('manualFields').querySelector('[name="level"]');if(first&&type==='soil')first.value='30 cm';
 // Basic numeric-format checks while typing, without scientific gating.
 for(const el of $('manualFields').querySelectorAll('input[type=text]')){
  if(['species','level'].includes(el.name))continue;
  el.addEventListener('input',()=>{
   const good=!el.value.trim()||numeric(el.value)!==null;
   el.classList.toggle('invalid',!good);
   const msg=el.parentElement.querySelector('.input-error');if(msg){msg.hidden=good;msg.textContent=good?'':'Folosește o valoare numerică (punct sau virgulă).';}
  });
 }
}
async function saveSubmission(records,file=null){
 if(!window.NECD_CLOUD?.isConnected())throw Error('Google Drive nu este conectat.');
 if(!records.length)throw Error('Nu există înregistrări de salvat.');
 // Validate the new observations before writing to Drive. An invalid or out-of-year
 // submission must not be silently archived and then fail during report generation.
 const inspected=records.map(diagnose);
 const valid=inspected.filter(x=>x.status!=='invalid');
 if(!valid.length){
  const issues=[...new Set(inspected.flatMap(x=>x.warnings))];
  throw Error('Analiza nu conține valori care pot fi incluse în raport. '+(issues.join(' ')||'Verifică valorile numerice introduse.'));
 }
 if(!valid.some(x=>x.year===state.year)){
  // Manual entries and laboratory files may refer to a different reporting year.
  // Select the year represented by the new data instead of producing an empty report.
  const counts=new Map();
  for(const x of valid)counts.set(x.year,(counts.get(x.year)||0)+1);
  state.year=[...counts].sort((a,b)=>b[1]-a[1])[0][0];
  render();
 }
 if(file)await window.NECD_CLOUD.archiveSource(file,records[0].fileId);
 const entry={time:now(),what:file?file.name:CATEGORIES.find(c=>c.id===state.manualType).label,count:records.length,category:state.category,mode:file?'file':'manual',year:state.year};
 if(file)state.files.push({id:records[0].fileId,name:file.name,when:entry.time,count:records.length,kind:state.category});
 state.records.push(...records);state.submissions.push(entry);log('submission',`${entry.what} (${records.length})`);
 // Check the selected year's actual report data before persisting the workspace.
 if(!evalAll().some(x=>x.status!=='invalid'))
  throw Error('Nu există valori valide pentru anul '+state.year+'. Selectează anul corect al analizei.');
 await window.NECD_CLOUD.saveWorkspace();
 let report;
 try{
  report=await makeReport();
  await window.NECD_CLOUD.saveReport(report.blob,auditCsv(),state.year);
 }catch(error){
  throw Error('Înregistrarea a fost salvată în Google Drive, dar raportul Excel nu a putut fi actualizat: '+error.message);
 }
 render();notice(`Salvat: ${entry.what}. Raportul NECD ${state.year} a fost actualizat (${report.count} înregistrări).`,'ok');
}
async function handleFile(){
 const f=$('labFile').files?.[0];if(!f)throw Error('Selectează un fișier.');
 const parsed=await importFile(f);
 // Do not alter input files or pre-aggregate their records: each source row is retained.
 await saveSubmission(parsed.results,f);$('labFile').value='';
}
async function handleManual(ev){ev.preventDefault();const form=ev.currentTarget;const rec=manualRecord(form);await saveSubmission([rec]);form.reset();render();drawManualFields();}
function setBusy(b){state.busy=b;for(const btn of ['importBtn','manualSubmit','exportBtn'])if($(btn))$(btn).disabled=b;}
async function safely(task){
 setBusy(true);try{await task();}catch(e){console.error(e);notice(e.message||String(e),'error');}finally{setBusy(false);render();}
}
function setup(){
 $('projectYear').addEventListener('change',ev=>{
  const n=Number(ev.target.value);if(n>=2000&&n<=2100){state.year=n;log('year',String(n));render();}
 });
 $('category').innerHTML=LAB_CATEGORIES.map(c=>`<option value="${c.id}">${html(c.label)}</option>`).join('');
 $('category').addEventListener('change',ev=>{state.category=ev.target.value;manualLayout();render();});
 $('entryMethod').addEventListener('change',ev=>{state.method=ev.target.value;render();});
 $('manualType').addEventListener('change',ev=>{state.manualType=ev.target.value;drawManualFields();});
 $('importBtn').addEventListener('click',()=>safely(handleFile));
 $('manualForm').addEventListener('submit',ev=>{ev.preventDefault();safely(()=>handleManual(ev));});
 $('exportBtn').addEventListener('click',()=>safely(async()=>{const result=await makeReport();download(result.blob,`NECD_Art9_${state.year}.xlsx`);notice('Raportul NECD a fost descărcat.','ok');}));
 $('csvBtn').addEventListener('click',()=>download(auditCsv(),'NECD_QAQC.csv'));
 $('detailsBtn').addEventListener('click',()=>{$('details').hidden=!$('details').hidden;});
 $('siteMapToggle').addEventListener('click',()=>{const pane=$('siteMapPanel');pane.hidden=!pane.hidden;if(!pane.hidden)drawMapping();});
 $('siteMapPanel').addEventListener('change',ev=>{if(!ev.target.dataset.site)return;
  const raw=ev.target.dataset.site;state.manualSites[normalized(raw)]=ev.target.value;render();window.NECD_CLOUD?.saveWorkspace().catch(e=>notice('Asocierea stației nu a fost salvată: '+e.message,'error'));
 });
 manualLayout();render();
}
function drawMapping(){
 const missing=[...new Set(state.records.filter(x=>!station(x.originSite)).map(x=>x.originSite))];
 $('siteMapPanel').innerHTML=missing.length?missing.map(raw=>`<label><span>${html(raw)}</span><select data-site="${html(raw)}"><option value="">Nicio asociere (păstrează denumirea)</option>${(state.template?.sites||[]).map(s=>`<option value="${html(s.code)}" ${state.manualSites[normalized(raw)]===s.code?'selected':''}>${html(s.code)} · ${html(s.name)}</option>`).join('')}</select></label>`).join(''):'<p>Toate stațiile au cod identificat.</p>';
}
window.NECD_APP={state,loadPrivateTemplate,serializeWorkspace,restoreWorkspace,render,download,notice,makeReport};
document.addEventListener('DOMContentLoaded',setup);
