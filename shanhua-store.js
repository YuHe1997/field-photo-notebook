// Isolated from the photo notebook database. No migrations of existing user data.
let opening;
function openFiles(){
 if(!opening)opening=new Promise((resolve,reject)=>{
  const request=indexedDB.open('field-photo-notebook-shanhua-v1',1);
  request.onupgradeneeded=()=>request.result.createObjectStore('files');
  request.onerror=()=>reject(request.error);
  request.onblocked=()=>reject(Error('善化圖片資料庫被其他分頁占用'));
  request.onsuccess=()=>{const db=request.result;db.onversionchange=()=>{db.close();opening=null;};resolve(db);};
 }).catch(error=>{opening=null;throw error;});
 return opening;
}
function validateKey(key){if(typeof key!=='string'||!(/^(signature:(monitor|leader)|map:\d{4}-\d{2}-\d{2})$/).test(key))throw Error('圖片索引無效');}
async function transaction(key,mode,action){
 validateKey(key);const db=await openFiles();return new Promise((resolve,reject)=>{
  const tx=db.transaction('files',mode),request=action(tx.objectStore('files'));let value;
  request.onsuccess=()=>{value=request.result;};
  tx.oncomplete=()=>resolve(value??null);tx.onerror=tx.onabort=()=>reject(tx.error||Error('圖片儲存未完成'));
 });
}
export const getShanhuaFile=key=>transaction(key,'readonly',store=>store.get(key));
export const removeShanhuaFile=key=>transaction(key,'readwrite',store=>store.delete(key));
export async function resizeShanhuaImage(blob,{signature=false}={}){
 if(!(blob instanceof Blob)||!blob.type.startsWith('image/'))throw Error('請選擇圖片檔案');
 const url=URL.createObjectURL(blob),image=new Image();
 try{
  image.src=url;await image.decode();
  const scale=Math.min(1,(signature?800:1600)/Math.max(image.naturalWidth,image.naturalHeight));
  const canvas=document.createElement('canvas');canvas.width=Math.max(1,Math.round(image.naturalWidth*scale));canvas.height=Math.max(1,Math.round(image.naturalHeight*scale));
  const context=canvas.getContext('2d');if(!context)throw Error('無法建立縮圖');
  if(!signature){context.fillStyle='#fff';context.fillRect(0,0,canvas.width,canvas.height);}
  context.drawImage(image,0,0,canvas.width,canvas.height);
  const result=await new Promise(resolve=>canvas.toBlob(resolve,signature?'image/png':'image/jpeg',.9));
  if(!result)throw Error('無法轉換圖片');
  return {blob:result,width:canvas.width,height:canvas.height};
 }finally{URL.revokeObjectURL(url);}
}
export async function saveShanhuaFile(key,blob){
 validateKey(key);const image=await resizeShanhuaImage(blob,{signature:key.startsWith('signature:')});
 await transaction(key,'readwrite',store=>store.put(image,key));return image;
}
