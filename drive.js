// User-initiated export only. Access tokens and the GIS client live in memory, never in backups.
import {documentTabs,selectedTab,appendAnchor,verifyProtected,dailyTablePlan,structureRequests,newTable,mergeRequests,contentRequests,assertAppendOnly,namedMarker,logicalTableCells,masterHeadingTemplate,headingRangeText} from './docs-append.js';
export const CLIENT_ID='659642111191-16q64t5mbd32hba7um1gvqd3vmc4tn9p.apps.googleusercontent.com';
const SCOPE='https://www.googleapis.com/auth/drive';
const DOCX='application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const GOOGLE_DOC='application/vnd.google-apps.document';
const FOLDER='application/vnd.google-apps.folder';
let client,token=null,expiresAt=0,loading,authorizing;

export class DriveError extends Error{constructor(kind,message){super(message);this.name='DriveError';this.kind=kind;}}
export function parseDriveId(value){
 const input=value.trim();if(!input)return null; // My Drive root
 if(/^[\w-]{10,}$/.test(input))return input;
 try{const url=new URL(input);if(url.protocol!=='https:'||!['drive.google.com','docs.google.com'].includes(url.hostname))throw Error();
  const match=url.pathname.match(/\/(?:folders|document\/d)\/([\w-]+)/);if(match&&match[1].length>=10)return match[1];
  const id=url.searchParams.get('id');if(id&&/^[\w-]{10,}$/.test(id))return id;
 }catch{}throw new DriveError('input','請輸入有效的 Google Drive 資料夾 ID 或網址。');
}
function gis(){
 if(globalThis.google?.accounts?.oauth2)return Promise.resolve(globalThis.google.accounts.oauth2);
 if(!loading)loading=new Promise((resolve,reject)=>{
  const script=document.createElement('script');script.src='https://accounts.google.com/gsi/client';script.async=true;script.defer=true;
  const timer=setTimeout(()=>{
   script.remove();loading=null;
   reject(new DriveError('network','載入 Google 登入元件逾時（10 秒），請檢查網路。本機 Word 匯出不受影響。'));
  },10000);
  script.onload=()=>{
   clearTimeout(timer);
   globalThis.google?.accounts?.oauth2?resolve(globalThis.google.accounts.oauth2):reject(new DriveError('auth','Google 登入元件未能載入。'));
  };
  script.onerror=()=>{
   clearTimeout(timer);
   script.remove();loading=null;
   reject(new DriveError('network','無法載入 Google 登入元件，請檢查網路或瀏覽器封鎖設定。'));
  };
  document.head.append(script);
 }).catch(e=>{loading=null;throw e;});
 return loading;
}
export async function authorize(forceSelect=false){
 if(!forceSelect&&token&&Date.now()<expiresAt-60000)return token;
 if(authorizing)return authorizing;
 authorizing=authorizeFresh(forceSelect?'select_account':'').finally(()=>{authorizing=null;});
 return authorizing;
}
async function authorizeFresh(promptMode='select_account'){
 const oauth=await gis();
 return new Promise((resolve,reject)=>{
  if(!client)client=oauth.initTokenClient({client_id:CLIENT_ID,scope:SCOPE,callback:()=>{}});
  client.callback=result=>{
   if(result.error){token=null;reject(new DriveError('auth',`Google 授權未完成：${result.error_description||result.error}`));return;}
   token=result.access_token;expiresAt=Date.now()+(Number(result.expires_in)||3600)*1000;resolve(token);
  };
  client.error_callback=error=>reject(new DriveError('auth',`Google 登入中斷：${error.type||'視窗關閉或遭封鎖'}`));
  try{client.requestAccessToken({prompt:promptMode});}catch(e){reject(new DriveError('auth',`無法開啟 Google 授權：${e.message}`));}
 });
}
export function hasToken(){return !!token&&Date.now()<expiresAt-60000;}
async function api(url,options={}){
 const access=await authorize();let response;
 try{response=await fetch(url,{...options,headers:{...options.headers,Authorization:`Bearer ${access}`}});}
 catch{throw new DriveError('network',options.method==='POST'?'網路未回應；上傳結果未知。請先到 Drive 檢查是否已建立文件，再決定是否重試。':'網路未回應；請檢查連線後重試。');}
 if(!response.ok){if(response.status===401){token=null;expiresAt=0;throw new DriveError('auth','Google 授權已過期，請重新登入後重試。');}
  let detail='';try{detail=(await response.json()).error?.message||'';}catch{}
  throw new DriveError('api',`Google API 回報 ${response.status}${detail?'：'+detail:''}。`);
 }
 try{return await response.json();}catch{throw new DriveError('api','Google API 回應格式不正確；請先檢查 Drive 再重試。');}
}
export async function listSharedDrives(){
 const drives=[];let pageToken;const seen=new Set();
 do{
  const params=new URLSearchParams({pageSize:'100',fields:'nextPageToken,drives(id,name)'});
  if(pageToken)params.set('pageToken',pageToken);
  const data=await api(`https://www.googleapis.com/drive/v3/drives?${params}`);
  drives.push(...(data.drives||[]));pageToken=data.nextPageToken;
  if(pageToken&&seen.has(pageToken))throw new DriveError('api','共用雲端硬碟分頁回應重複，請重新讀取。');
  if(pageToken)seen.add(pageToken);
 }while(pageToken);
 return drives.sort((a,b)=>a.name.localeCompare(b.name));
}
export async function listSubfolders(parentId,{shared=false,driveId=null}={}){
 const parent=parentId||'root';
 if(!/^[\w-]+$/.test(parent))throw new DriveError('input','資料夾 ID 格式無效。');
 const location=shared?'sharedWithMe = true':`'${parent}' in parents`;
 const q=`${location} and mimeType = '${FOLDER}' and trashed = false`;
 const files=[];let pageToken;const seen=new Set();
 do{
  const params=new URLSearchParams({q,fields:'nextPageToken,files(id,name,capabilities(canAddChildren))',pageSize:'100',orderBy:'name',supportsAllDrives:'true',includeItemsFromAllDrives:'true'});
  if(driveId){
   if(!/^[\w-]+$/.test(driveId))throw new DriveError('input','共用雲端硬碟 ID 格式無效。');
   params.set('corpora','drive');params.set('driveId',driveId);
  }
  if(pageToken)params.set('pageToken',pageToken);
  const data=await api(`https://www.googleapis.com/drive/v3/files?${params}`);
  files.push(...(data.files||[]));pageToken=data.nextPageToken;
  if(pageToken&&seen.has(pageToken))throw new DriveError('api','Drive 分頁回應重複，請重新讀取。');
  if(pageToken)seen.add(pageToken);
 }while(pageToken);
 return files;
}

export async function createFolder(name,parentId=null){
 if(!name||!name.trim())throw new DriveError('input','資料夾名稱不可為空');
 const meta={name:name.trim(),mimeType:FOLDER};
 if(parentId)meta.parents=[parentId];
 const res=await api('https://www.googleapis.com/drive/v3/files?fields=id,name,mimeType&supportsAllDrives=true',{
  method:'POST',
  headers:{'Content-Type':'application/json; charset=UTF-8'},
  body:JSON.stringify(meta)
 });
 return res;
}

export async function ensureDailyFolder(parentId,day){
 const folderName=day;
 const existing=await listSubfolders(parentId);
 const found=existing.find(f=>f.name===folderName);
 if(found)return found;
 return await createFolder(folderName,parentId);
}

export async function verifyFolder(id){
 try{
  const file=await api(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(id)}?fields=id,mimeType,name,capabilities(canAddChildren)&supportsAllDrives=true`);
  if(file.mimeType!==FOLDER)throw new DriveError('input','指定的 ID 不是資料夾。');
  if(file.capabilities?.canAddChildren!==true)throw new DriveError('api','沒有權限在此資料夾新增文件。');
  return file;
 }catch(e){if(e.kind==='api'&&/回報 (403|404)/.test(e.message))throw new DriveError('api','無法存取此資料夾（403/404）。請確認目前 Google 帳號擁有資料夾的編輯權限，並已允許 App 存取 Google Drive。');throw e;}
}
export function validateGoogleDocName(value){
 if(typeof value!=='string'||!value.trim())throw new DriveError('input','請輸入獨立文件名稱。');
 const name=value.trim();
 if(name.length>128||/[\u0000-\u001f\u007f]/u.test(name))throw new DriveError('input','文件名稱請使用 128 字元內的單行文字，不可包含控制字元。');
 return name;
}
export async function uploadDayDoc(day,docBlob,folderId=null,{name:customName}={}){
 if(!/^\d{4}-\d{2}-\d{2}$/.test(day))throw new DriveError('input','工作日格式無效。');
 const name=customName===undefined?`施工監看紀錄-${day}`:validateGoogleDocName(customName);
 if(folderId)await verifyFolder(folderId);
 const boundary=`fieldPhotoNotebook_${crypto.randomUUID().replaceAll('-','')}`;
 const meta={name,mimeType:GOOGLE_DOC};if(folderId)meta.parents=[folderId];
 const body=new Blob([`--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(meta)}\r\n--${boundary}\r\nContent-Type: ${DOCX}\r\n\r\n`,docBlob,`\r\n--${boundary}--\r\n`]);
 const file=await api('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,mimeType,webViewLink,parents&supportsAllDrives=true',{
  method:'POST',headers:{'Content-Type':`multipart/related; boundary=${boundary}`},body
 });
 if(!file.id||file.mimeType!==GOOGLE_DOC)throw new DriveError('api','Drive 未確認轉換成原生 Google Doc；請先檢查 Drive，避免重複上傳。');
 return {...file,webViewLink:file.webViewLink||`https://docs.google.com/document/d/${encodeURIComponent(file.id)}/edit`};
}

export function formatPhotoFileName(photo){
 const day=photo.workDay||'unknown-day';
 const num=Number.isSafeInteger(photo.photoNumber)?String(photo.photoNumber).padStart(3,'0'):'000';
 const uuidPart=(photo.id||'').split('-')[0]||crypto.randomUUID().slice(0,8);
 const ext=photo.mime==='image/png'?'.png':photo.mime==='image/webp'?'.webp':'.jpg';
 return `${day}_${num}_${uuidPart}${ext}`;
}

export async function uploadOriginalPhoto(photo,blob,folderId=null,onProgress=null){
 if(!blob)throw new DriveError('input','缺少原始影像檔案');
 if(folderId)await verifyFolder(folderId);
 const fileName=formatPhotoFileName(photo);
 const boundary=`fieldPhotoNotebook_${crypto.randomUUID().replaceAll('-','')}`;
 const meta={
  name:fileName,
  mimeType:photo.mime||blob.type||'image/jpeg',
  appProperties:{
   photoId:photo.id,
   workDay:photo.workDay||'',
   photoNumber:String(photo.photoNumber||'')
  }
 };
 if(folderId)meta.parents=[folderId];
 const body=new Blob([
  `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(meta)}\r\n--${boundary}\r\nContent-Type: ${meta.mimeType}\r\n\r\n`,
  blob,
  `\r\n--${boundary}--\r\n`
 ]);
 const res=await api('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,mimeType,webViewLink,parents&supportsAllDrives=true',{
  method:'POST',headers:{'Content-Type':`multipart/related; boundary=${boundary}`},body
 });
 if(!res.id)throw new DriveError('api','Drive 上傳未確認成功；請檢查 Drive。');
 return res;
}

async function docsApi(url,options={}){
 const controller=new AbortController(),timer=setTimeout(()=>controller.abort(),60000);
 try{return await api(url,{...options,signal:controller.signal});}finally{clearTimeout(timer);}
}
export async function readGoogleDocument(id){
 if(!/^[\w-]{10,}$/.test(id||''))throw new DriveError('input','Google 文件 ID 無效。');
 return docsApi(`https://docs.googleapis.com/v1/documents/${encodeURIComponent(id)}?includeTabsContent=true`);
}
export async function inspectMaster(value){
 const id=parseDriveId(value);if(!id)throw new DriveError('input','請貼上 Google 文件副本網址。');
 const file=await api(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(id)}?supportsAllDrives=true&fields=id,name,mimeType,parents,capabilities(canEdit)`);
 if(file.mimeType!==GOOGLE_DOC)throw new DriveError('input','這仍是 Word 或其他檔案。請使用原生 Google 文件副本，不會修改原 Word 母檔。');
 if(file.capabilities?.canEdit!==true)throw new DriveError('api','目前 Google 帳號沒有此文件的編輯權限。');
 const document=await readGoogleDocument(id),tabs=documentTabs(document);
 if(!tabs.length)throw new DriveError('api','Google 未回傳可操作的文件分頁。');
 return {file,document,tabs:tabs.map(t=>({id:t.id,title:t.title}))};
}
export async function appendMonitoringRecord({masterId,tabId,sourceId,day,fingerprint,rowLayout,onProgress=()=>{},onJournal}){
 if(!/^\d{4}-\d{2}-\d{2}$/.test(day)||!/^[a-f0-9]{64}$/.test(fingerprint))throw new DriveError('input','追加識別資料無效。');
 if(typeof onJournal!=='function')throw new DriveError('input','追加操作必須能保存本機操作紀錄。');
 if(masterId===sourceId)throw new DriveError('input','當日文件不可與母檔相同。');
 const checked=await inspectMaster(masterId),before=checked.document;
 selectedTab(before,tabId);
 const source=await readGoogleDocument(sourceId),plan=dailyTablePlan(source,{rowLayout});
 const suffix=`${day}_${fingerprint}`,pending=`fpn_pending_${suffix}`,done=`fpn_done_${suffix}`;
 if(namedMarker(before,tabId,pending)||namedMarker(before,tabId,done))throw new DriveError('duplicate','母檔已有相同內容的追加標記。請開啟文件核對，不會重複追加。');
 const heading=masterHeadingTemplate(before,tabId);
 const anchor=appendAnchor(before,tabId),tail=selectedTab(before,tabId).data.body.content.at(-1);
 // A trailing empty placeholder may inherit newline styling; never exclude non-empty old text.
 const emptyTail=tail.paragraph.elements.every(e=>e.textRun)&&tail.paragraph.elements.map(e=>e.textRun.content||'').join('')==='\n';
 const boundary=emptyTail?tail.startIndex:anchor+1;
 let current=before;
 const write=async(phase,requests,floor)=>{
  assertAppendOnly(requests,tabId,floor,anchor+1);
  verifyProtected(before,current,tabId,boundary);
  // Persist BEFORE sending. A lost response may already have committed; never auto-retry.
  onJournal({phase,masterId,tabId,sourceId,day,fingerprint});onProgress(phase);
  await docsApi(`https://docs.googleapis.com/v1/documents/${encodeURIComponent(masterId)}:batchUpdate`,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({requests,writeControl:{requiredRevisionId:current.revisionId}})});
  current=await readGoogleDocument(masterId);verifyProtected(before,current,tabId,boundary);
 };
 await write('建立追加表格',structureRequests(anchor,tabId,day,pending,plan.rows.length,heading),anchor);
 let table=newTable(current,tabId,anchor),merges=mergeRequests(table,tabId,plan);
 if(merges.length){await write('合併新表格儲存格',merges,table.startIndex);table=newTable(current,tabId,anchor);}
 await write('填入文字與照片',contentRequests(table,tabId,plan,done,anchor,day,heading),table.startIndex);
 if(!namedMarker(current,tabId,done))throw new DriveError('api','追加完成標記未確認；請核對母檔，不會自動重試。');
 const expectedHeading=(heading?.text||`\n監看紀錄 ${day}\n`).slice(1);
 if(headingRangeText(current,tabId,anchor+1,anchor+1+expectedHeading.length)!==expectedHeading)throw new DriveError('api','新增標題讀回不符，請核對副本；不自動重試。');
 const result=newTable(current,tabId,anchor).table;
 let expectedText='',actualText='',expectedImages=0,actualImages=0;
 for(const row of plan.rows)for(const cell of row){expectedText+=cell.segments.filter(s=>s.kind==='text').map(s=>s.text).join('')+'\n';expectedImages+=cell.segments.filter(s=>s.kind==='image').length;}
 for(const [i,row] of result.tableRows.entries())for(const {cell} of logicalTableCells(row,plan.columns,plan.rows[i].map(c=>c.span)))for(const node of cell.content)for(const e of node.paragraph?.elements||[]){actualText+=e.textRun?.content||'';if(e.inlineObjectElement)actualImages++;}
 if(expectedText!==actualText||expectedImages!==actualImages)throw new DriveError('api','新增文字或照片數量核對不符，請檢查母檔；未刪除任何內容。');
 onJournal({phase:'verified',masterId,tabId,sourceId,day,fingerprint});
 return {name:checked.file.name,webViewLink:`https://docs.google.com/document/d/${encodeURIComponent(masterId)}/edit`,images:actualImages,rows:result.tableRows.length};
}

