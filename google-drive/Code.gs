/**
 * NECD Submission — Google Drive storage bridge.
 * Paste this entire file into a new Google Apps Script project (Code.gs).
 * Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone.
 * Upload the original "NECD art9 Reporting_06_26.xlsx" to your Google Drive once.
 * The script automatically creates its "NECD Submission" storage folder.
 */
const NECD_FOLDER = 'NECD Submission';
const NECD_TEMPLATE = 'NECD art9 Reporting_06_26.xlsx';
const NECD_ALLOWED_ORIGIN = 'https://incdsnec.github.io';

function doGet() {
  const page = `<!doctype html><html><head><meta charset="utf-8"></head><body>
<script>
window.addEventListener('message',function(event){
  if(event.origin!==${JSON.stringify(NECD_ALLOWED_ORIGIN)})return;
  const data=event.data;
  if(!data||data.channel!=='necd-drive-request'||!data.id)return;
  const reply=function(result,error){
    event.source.postMessage({channel:'necd-drive-response',id:data.id,result:result,error:error||null},event.origin);
  };
  google.script.run.withSuccessHandler(function(result){reply(result,null)})
    .withFailureHandler(function(error){reply(null,error&&error.message?error.message:String(error))})
    .necdDispatch(data.action,data.payload||{});
});
window.parent.postMessage({channel:'necd-drive-ready'},'*');
</script></body></html>`;
  return HtmlService.createHtmlOutput(page)
    .setTitle('NECD Drive bridge')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function necdFolder() {
  const it = DriveApp.getFoldersByName(NECD_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(NECD_FOLDER);
}

function necdSubfolder(name) {
  const root=necdFolder(),it=root.getFoldersByName(name);
  return it.hasNext()?it.next():root.createFolder(name);
}

function necdFile(folder,name) {
  const it=folder.getFilesByName(name);
  return it.hasNext()?it.next():null;
}

function necdSaveText(folder,name,text,mime) {
  const file=necdFile(folder,name);
  if(file){file.setContent(String(text));return;}
  folder.createFile(name,String(text),mime||'text/plain');
}

function necdSaveBase64(folder,name,base64,mime) {
  const existing=necdFile(folder,name);
  const blob=Utilities.newBlob(Utilities.base64Decode(base64),mime,name);
  if(existing)existing.setTrashed(true);
  folder.createFile(blob);
}

function necdTemplate() {
  const own=necdFile(necdFolder(),NECD_TEMPLATE);
  if(own)return own;
  const it=DriveApp.getFilesByName(NECD_TEMPLATE);
  if(it.hasNext())return it.next();
  throw new Error('Lipsește șablonul NECD art9 Reporting_06_26.xlsx din Google Drive. Încarcă-l o singură dată.');
}

function necdDispatch(action,data) {
  if(action==='load'){
    const file=necdFile(necdFolder(),'workspace.json');
    return {
      workspace:file?file.getBlob().getDataAsString('UTF-8'):null,
      template:Utilities.base64Encode(necdTemplate().getBlob().getBytes())
    };
  }
  if(action==='report'){
    const file=necdFile(necdSubfolder('reports'),'NECD_Art9_'+Number(data.year)+'.xlsx');
    if(!file)throw new Error('Raportul pentru acest an nu a fost încă salvat.');
    return {base64:Utilities.base64Encode(file.getBlob().getBytes())};
  }
  const lock=LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    if(action==='saveWorkspace'){
      necdSaveText(necdFolder(),'workspace.json',data.json,'application/json');
      return {ok:true};
    }
    if(action==='saveSource'){
      if(!/^[A-Za-z0-9_-]+$/.test(data.id))throw new Error('ID fișier invalid.');
      const ext=data.extension==='csv'?'csv':'xlsx';
      const name=data.id+'.'+ext,folder=necdSubfolder('sources');
      if(!necdFile(folder,name))necdSaveBase64(folder,name,data.base64,ext==='csv'?'text/csv':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      return {ok:true};
    }
    if(action==='saveReport'){
      const year=Number(data.year);
      if(!Number.isInteger(year)||year<2000||year>2100)throw new Error('An invalid.');
      const folder=necdSubfolder('reports');
      necdSaveBase64(folder,'NECD_Art9_'+year+'.xlsx',data.base64,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      necdSaveText(folder,'NECD_Art9_'+year+'_QAQC.csv',data.audit,'text/csv');
      return {ok:true};
    }
    throw new Error('Operație necunoscută: '+action);
  }finally{lock.releaseLock();}
}
