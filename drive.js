// User-initiated export only. Access tokens and the GIS client live in memory, never in backups.
export const CLIENT_ID='659642111191-11dq6vitc6b2nltejs79p35ggb7e2a1k.apps.googleusercontent.com';
const SCOPE='https://www.googleapis.com/auth/drive.file';
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
export async function authorize(){
 if(token&&Date.now()<expiresAt-60000)return token;
 if(authorizing)return authorizing;
 authorizing=authorizeFresh().finally(()=>{authorizing=null;});
 return authorizing;
}
async function authorizeFresh(){
 const oauth=await gis();
 return new Promise((resolve,reject)=>{
  if(!client)client=oauth.initTokenClient({client_id:CLIENT_ID,scope:SCOPE,callback:()=>{}});
  client.callback=result=>{
   if(result.error){token=null;reject(new DriveError('auth',`Google 授權未完成：${result.error_description||result.error}`));return;}
   token=result.access_token;expiresAt=Date.now()+(Number(result.expires_in)||3600)*1000;resolve(token);
  };
  client.error_callback=error=>reject(new DriveError('auth',`Google 登入中斷：${error.type||'視窗關閉或遭封鎖'}`));
  try{client.requestAccessToken({prompt:token?'':'consent'});}catch(e){reject(new DriveError('auth',`無法開啟 Google 授權：${e.message}`));}
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
export async function verifyFolder(id){
 if(!id)return null;
 try{
  const file=await api(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(id)}?fields=id,mimeType,name,capabilities(canAddChildren)&supportsAllDrives=true`);
  if(file.mimeType!==FOLDER)throw new DriveError('input','指定的 ID 不是資料夾。');
  if(file.capabilities?.canAddChildren!==true)throw new DriveError('api','沒有權限在此資料夾新增文件。');
  return file;
 }catch(e){if(e.kind==='api'&&/回報 (403|404)/.test(e.message))throw new DriveError('api','無法存取此資料夾（403/404）。貼上連結不會授予 drive.file 權限；請選擇此應用建立的資料夾，或等待支援 Google Picker 授權。');throw e;}
}
export async function uploadDayDoc(day,docBlob,folderId=null){
 if(!/^\d{4}-\d{2}-\d{2}$/.test(day))throw new DriveError('input','工作日格式無效。');
 if(folderId)await verifyFolder(folderId);
 const boundary=`fieldPhotoNotebook_${crypto.randomUUID().replaceAll('-','')}`;
 const meta={name:`施工監看紀錄-${day}`,mimeType:GOOGLE_DOC};if(folderId)meta.parents=[folderId];
 const body=new Blob([`--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(meta)}\r\n--${boundary}\r\nContent-Type: ${DOCX}\r\n\r\n`,docBlob,`\r\n--${boundary}--\r\n`]);
 const file=await api('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,mimeType,webViewLink,parents',{
  method:'POST',headers:{'Content-Type':`multipart/related; boundary=${boundary}`},body
 });
 if(!file.id||file.mimeType!==GOOGLE_DOC)throw new DriveError('api','Drive 未確認轉換成原生 Google Doc；請先檢查 Drive，避免重複上傳。');
 return {...file,webViewLink:file.webViewLink||`https://docs.google.com/document/d/${encodeURIComponent(file.id)}/edit`};
}
