/* GitHub private storage for static app. Contents API and fine-grained PAT (restricted to data repo). */
'use strict';
(()=>{
 const $=id=>document.getElementById(id);
 const TOKEN_KEY='necd-v3-session-token';
 const FIXED_OWNER='INCDSNEC', FIXED_REPO='necd-reporting-data';
 const TEMPLATE='NECD art9 Reporting_06_26.xlsx';
 let owner=FIXED_OWNER,repo=FIXED_REPO,token='',isOnline=false,workspaceSha=null;
 let queue=Promise.resolve();
 const encode=path=>path.split('/').map(encodeURIComponent).join('/');
 const url=path=>`https://api.github.com/repos/${encodeURIComponent(owner)}/${encodeURIComponent(repo)}/contents/${encode(path)}`;
 function status(s,good=false){$('cloudState').textContent=s;$('cloudState').dataset.ok=good?'true':'false';}
 async function req(address,options={}){
  const response=await fetch(address,{...options,headers:{Accept:'application/vnd.github+json',Authorization:`Bearer ${token}`,...options.headers},cache:'no-store'});
  if(response.status===404&&options.optional)return null;
  if(!response.ok){let detail='';try{detail=(await response.json()).message||''}catch{}
   throw Error(`GitHub (${response.status}): ${detail||'Conexiune sau permisiune invalidă.'}`);
  }
  return response.json();
 }
 function decode(str){const b=atob(str.replace(/\s/g,''));return Uint8Array.from(b,ch=>ch.charCodeAt(0));}
 function b64(arr){const bytes=arr instanceof Uint8Array?arr:new Uint8Array(arr);let result='';for(let i=0;i<bytes.length;i+=8192)result+=String.fromCharCode(...bytes.subarray(i,i+8192));return btoa(result);}
 async function getFile(path,optional=false){
  const meta=await req(url(path),{optional});if(!meta)return null;
  if(meta.type!=='file')throw Error(`Nu există un fișier valid la ${path}.`);
  if(meta.content&&meta.encoding==='base64')return {bytes:decode(meta.content),sha:meta.sha};
  const raw=await req(`https://api.github.com/repos/${encodeURIComponent(owner)}/${encodeURIComponent(repo)}/git/blobs/${encodeURIComponent(meta.sha)}`);
  if(raw.content&&raw.encoding==='base64')return {bytes:decode(raw.content),sha:meta.sha};
  throw Error('Fișierul este prea mare sau indisponibil prin API.');
 }
 async function put(path,data,oldSha=null,description='actualizare'){
  const buffer=data instanceof Blob?await data.arrayBuffer():data;
  const binary=typeof buffer==='string'?new TextEncoder().encode(buffer):buffer;
  const body={message:`NECD: ${description}`,content:b64(binary)};
  if(oldSha)body.sha=oldSha;
  return req(url(path),{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});
 }
 function saveSerial(work){const p=queue.then(work);queue=p.catch(()=>{});return p;}
 async function fetchWorkspace(){
  const item=await getFile('workspace.json',true);
  if(item){const obj=JSON.parse(new TextDecoder().decode(item.bytes));window.NECD_APP.restoreWorkspace(obj);workspaceSha=item.sha;}
  else{workspaceSha=null;window.NECD_APP.restoreWorkspace({schema:3,year:2025,records:[],files:[],submissions:[],audit:[],manualSites:{}});}
 }
 async function connect(ev){
  ev?.preventDefault();
  const t=$('gitToken').value.trim()||sessionStorage.getItem(TOKEN_KEY)||'';
  if(!t)throw Error('Introdu tokenul de acces.');
  token=t;status('Conectare…');
  try{
   const info=await req(`https://api.github.com/repos/${encodeURIComponent(owner)}/${encodeURIComponent(repo)}`);
   if(!info.private)throw Error('Repository-ul pentru date trebuie să fie privat.');
   const template=await getFile(TEMPLATE);
   await window.NECD_APP.loadPrivateTemplate(new Blob([template.bytes]));
   await fetchWorkspace();
   isOnline=true;sessionStorage.setItem(TOKEN_KEY,token);
   $('gitToken').value='';$('connectForm').hidden=true;$('online').hidden=false;
   status('Sincronizat · GitHub privat',true);window.NECD_APP.render();
  }catch(e){isOnline=false;status('Conexiune nereușită: '+e.message);throw e;}
 }
 async function saveWorkspace(){
  if(!isOnline)throw Error('Nu ești conectat la spațiul privat.');
  return saveSerial(async()=>{
   status('Salvez progresul…');
   const serialized=JSON.stringify(window.NECD_APP.serializeWorkspace());
   try{
    const reply=await put('workspace.json',serialized,workspaceSha,'depunere analiză');
    workspaceSha=reply.content?.sha;
    status('Progres salvat · GitHub privat',true);
   }catch(e){status('Nu s-a salvat! '+e.message);throw e;}
  });
 }
 async function archiveSource(file,id){
  if(!isOnline)throw Error('Conexiune absentă.');
  const extension=/\.csv$/i.test(file.name)?'csv':'xlsx',path=`sources/${id}.${extension}`;
  if(!(await getFile(path,true)))await put(path,file,null,`fișier original ${file.name}`);
 }
 async function saveReport(blob,audit,year){
  if(!isOnline)throw Error('Conexiune absentă.');
  return saveSerial(async()=>{
   status('Actualizez raportul online…');
   const main=`reports/NECD_Art9_${year}.xlsx`,csv=`reports/NECD_Art9_${year}_QAQC.csv`;
   const prior=await getFile(main,true),priorCsv=await getFile(csv,true);
   await put(main,blob,prior?.sha,`raport NECD ${year}`);
   await put(csv,audit,priorCsv?.sha,`QAQC ${year}`);
   status('Raport și progres salvate · GitHub privat',true);
  });
 }
 async function retrieveReport(){
  if(!isOnline)throw Error('Conectează-te la GitHub.');
  const path=`reports/NECD_Art9_${window.NECD_APP.state.year}.xlsx`;
  const result=await getFile(path,true);if(!result)throw Error('Raportul nu a fost salvat încă.');
  window.NECD_APP.download(new Blob([result.bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),path.split('/').at(-1));
 }
 async function refresh(){
  if(!isOnline)return;
  if(window.NECD_APP.state.busy)throw Error('O depunere este în curs.');
  await fetchWorkspace();status('Date actualizate de pe GitHub',true);window.NECD_APP.render();
 }
 function disconnect(){
  token='';isOnline=false;workspaceSha=null;sessionStorage.removeItem(TOKEN_KEY);
  $('connectForm').hidden=false;$('online').hidden=true;window.NECD_APP.state.template=null;
  window.NECD_APP.restoreWorkspace({schema:3,year:2025,records:[],files:[],submissions:[],audit:[],manualSites:{}});
  status('Neconectat');
 }
 function setup(){
  $('connectForm').addEventListener('submit',ev=>connect(ev).catch(e=>window.NECD_APP.notice(e.message,'error')));
  $('refreshBtn').addEventListener('click',()=>refresh().catch(e=>window.NECD_APP.notice(e.message,'error')));
  $('disconnectBtn').addEventListener('click',disconnect);
  $('onlineReportBtn').addEventListener('click',()=>retrieveReport().catch(e=>window.NECD_APP.notice(e.message,'error')));
  const remembered=sessionStorage.getItem(TOKEN_KEY);
  if(remembered){$('gitToken').value=remembered;connect().catch(e=>window.NECD_APP.notice(e.message,'error'));}
 }
 window.NECD_CLOUD={isConnected:()=>isOnline,saveWorkspace,archiveSource,saveReport,refresh};
 document.addEventListener('DOMContentLoaded',setup);
})();
