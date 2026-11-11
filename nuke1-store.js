// Separate database: never migrates or touches photo/Shanhua data. Not in TAR.
let opening;
function openFiles(){
 if(!opening)opening=new Promise((resolve,reject)=>{
  const request=indexedDB.open('field-photo-notebook-nuke1-v1',1);
  request.onupgradeneeded=()=>request.result.createObjectStore('files');
  request.onerror=()=>reject(request.error);request.onblocked=()=>reject(Error('核一簽名資料庫被其他分頁占用'));
  request.onsuccess=()=>{const db=request.result;db.onversionchange=()=>{db.close();opening=null;};resolve(db);};
 }).catch(error=>{opening=null;throw error;});
 return opening;
}
function validateKey(key){if(!['signature:monitor','signature:leader'].includes(key))throw Error('簽名索引無效');}
async function transaction(key,mode,action){
 validateKey(key);const db=await openFiles();return new Promise((resolve,reject)=>{
  const tx=db.transaction('files',mode),request=action(tx.objectStore('files'));let value;
  request.onsuccess=()=>{value=request.result;};tx.oncomplete=()=>resolve(value??null);
  tx.onerror=tx.onabort=()=>reject(tx.error||Error('簽名儲存未完成'));
 });
}
export const getNuke1File=key=>transaction(key,'readonly',store=>store.get(key));
export const removeNuke1File=key=>transaction(key,'readwrite',store=>store.delete(key));
export async function saveNuke1File(key,blob){
 validateKey(key);if(!(blob instanceof Blob)||!blob.type.startsWith('image/'))throw Error('請選擇圖片檔案');
 const url=URL.createObjectURL(blob),image=new Image();
 try{
  image.src=url;await image.decode();const scale=Math.min(1,800/Math.max(image.naturalWidth,image.naturalHeight));
  const canvas=document.createElement('canvas');canvas.width=Math.max(1,Math.round(image.naturalWidth*scale));canvas.height=Math.max(1,Math.round(image.naturalHeight*scale));
  const context=canvas.getContext('2d');if(!context)throw Error('無法建立簽名縮圖');context.drawImage(image,0,0,canvas.width,canvas.height);
  const resized=await new Promise(resolve=>canvas.toBlob(resolve,'image/png'));if(!resized)throw Error('無法轉換簽名圖片');
  const value={blob:resized,width:canvas.width,height:canvas.height};await transaction(key,'readwrite',store=>store.put(value,key));return value;
 }finally{URL.revokeObjectURL(url);}
}
