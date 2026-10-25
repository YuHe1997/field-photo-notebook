import {workDayOf,validDay,validateGroup,validCrop,normalizeCrop,rotateCrop} from './core.js';
import {compareGroups} from './ordering.js';
let connection;
export async function open(){if(connection)return connection;connection=await new Promise((resolve,reject)=>{const r=indexedDB.open('field-photo-notebook-v1',1);r.onupgradeneeded=()=>{for(const n of ['photos','assets','groups'])r.result.createObjectStore(n,{keyPath:'id'});};r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);r.onblocked=()=>reject(Error('請關閉其他使用本工具的分頁'));});return connection;}
export async function transaction(stores,action){const db=await open();return new Promise((resolve,reject)=>{const tx=db.transaction(stores,'readwrite');tx.oncomplete=()=>resolve();tx.onerror=()=>reject(tx.error);tx.onabort=()=>reject(tx.failure||tx.error||Error('交易已中止'));try{action(tx);}catch(e){tx.abort();reject(e);}});}
function checked(tx,fn){return ()=>{try{fn();}catch(e){tx.failure=e;tx.abort();}};}
export async function all(store){const db=await open();return new Promise((resolve,reject)=>{const tx=db.transaction(store),r=tx.objectStore(store).getAll();let result;r.onsuccess=()=>result=r.result;tx.oncomplete=()=>resolve(result);tx.onerror=()=>reject(tx.error);tx.onabort=()=>reject(tx.error||Error('讀取交易已中止'));});}
const COUNTERS='field-photo-notebook-number-counters';
export function numberCounters(){try{const data=JSON.parse(localStorage.getItem(COUNTERS)||'{}');return Object.fromEntries(Object.entries(data).filter(([day,n])=>validDay(day)&&Number.isSafeInteger(n)&&n>0));}catch{return {};}}
export function mergeCounters(incoming){const out=numberCounters();for(const [day,n] of Object.entries(incoming??{}))if(validDay(day)&&Number.isSafeInteger(n)&&n>0)out[day]=Math.max(out[day]||0,n);localStorage.setItem(COUNTERS,JSON.stringify(out));return out;}
function countersFor(records){const counters=numberCounters();for(const p of records)if(Number.isSafeInteger(p.photoNumber)&&p.photoNumber>0){const day=workDayOf(p);counters[day]=Math.max(counters[day]||0,p.photoNumber);}return counters;}
// Additive metadata only: no DB version change, original Blob/title/notes edits or renumbering.
export const ensurePhotoIdentities=()=>transaction(['photos'],tx=>{const store=tx.objectStore('photos'),r=store.getAll();r.onsuccess=checked(tx,()=>{const records=r.result.sort((a,b)=>(a.buttonPressedAt??a.capturedAt??a.importedAt??0)-(b.buttonPressedAt??b.capturedAt??b.importedAt??0)||a.id.localeCompare(b.id)),counters=countersFor(records);for(const p of records){const workDay=workDayOf(p);let photoNumber=p.photoNumber;if(!Number.isSafeInteger(photoNumber)||photoNumber<1)photoNumber=counters[workDay]=(counters[workDay]||0)+1;if(p.workDay!==workDay||p.photoNumber!==photoNumber)store.put({...p,workDay,photoNumber});}mergeCounters(counters);});});
export const savePhoto=(photo,blob)=>transaction(['photos','assets'],tx=>{const store=tx.objectStore('photos'),r=store.getAll();r.onsuccess=checked(tx,()=>{const counters=countersFor(r.result),workDay=workDayOf(photo),photoNumber=(counters[workDay]||0)+1;mergeCounters({...counters,[workDay]:photoNumber});store.add({...photo,workDay,photoNumber});tx.objectStore('assets').add({id:photo.id,original:blob});});});
const groupInDay=(g,day,photos)=>g.workDay===day||g.photoIds.some(id=>photos.some(p=>p.id===id&&workDayOf(p)===day));
// Content saves read the current rank in the same transaction. Legacy mixed-day
// groups remain editable; no splitting, schema upgrade or original-byte writes.
export const saveGroup=(g,{existing=false}={})=>transaction(['photos','groups'],tx=>{
 const p=tx.objectStore('photos').getAll(),store=tx.objectStore('groups'),r=store.getAll();
 r.onsuccess=checked(tx,()=>{
  validateGroup(g,new Set(p.result.map(x=>x.id)));
  const previous=r.result.find(x=>x.id===g.id);
  if(existing&&!previous)throw Error('紀錄集已被其他分頁移除；未重新建立');
  if(previous){const next={...previous,...g};delete next.sortOrder;if(Object.hasOwn(previous,'sortOrder'))next.sortOrder=previous.sortOrder;store.put(next);}
  else{
   const day=g.workDay??workDayOf(p.result.find(x=>x.id===g.photoIds[0]));
   const ordered=r.result.filter(x=>groupInDay(x,day,p.result)).sort(compareGroups);
   ordered.forEach((x,i)=>store.put({...x,sortOrder:i+1}));store.add({...g,sortOrder:ordered.length+1});
  }
 });
});
export const saveGroupOrder=(day,orderedIds)=>transaction(['photos','groups'],tx=>{
 if(!validDay(day)||!Array.isArray(orderedIds)||new Set(orderedIds).size!==orderedIds.length)throw Error('排序 ID 重複或日期無效');
 const p=tx.objectStore('photos').getAll(),store=tx.objectStore('groups'),r=store.getAll();
 r.onsuccess=checked(tx,()=>{
  const current=r.result.filter(g=>groupInDay(g,day,p.result));
  if(current.length!==orderedIds.length||orderedIds.some(id=>!current.some(g=>g.id===id)))throw Error('同日紀錄集已變更，請重新載入後排序');
  orderedIds.forEach((id,i)=>store.put({...current.find(g=>g.id===id),sortOrder:i+1}));
 });
});
export const removeGroup=id=>transaction(['groups'],tx=>tx.objectStore('groups').delete(id));
// A saved crop window turns with the photo, so the same part stays framed.
export const setPhotoRotation=(id,rotation,delta=0)=>transaction(['photos'],tx=>{if(![0,90,180,270].includes(rotation))throw Error('旋轉角度無效');const store=tx.objectStore('photos'),r=store.get(id);r.onsuccess=checked(tx,()=>{if(!r.result)throw Error('照片已不存在，請重新整理');const next={...r.result,displayRotation:rotation};if(validCrop(next.crop)){const c=delta?normalizeCrop(rotateCrop(next.crop,delta)):next.crop;if(c)next.crop=c;else delete next.crop;}store.put(next);});});
// Non-destructive 4:3 crop window (normalised to the displayed photo). null removes it. Original bytes are never touched.
export const setPhotoCrop=(id,crop)=>transaction(['photos'],tx=>{const value=crop===null?null:normalizeCrop(crop);if(crop!==null&&!value)throw Error('裁切範圍無效');const store=tx.objectStore('photos'),r=store.get(id);r.onsuccess=checked(tx,()=>{if(!r.result)throw Error('照片已不存在，請重新整理');const next={...r.result};if(value)next.crop=value;else delete next.crop;store.put(next);});});
// Deletion is permanent but only for ungrouped originals. Check references INSIDE the same
// transaction to protect a group another tab may have saved after the confirmation.
export const deletePhotos=ids=>transaction(['photos','assets','groups'],tx=>{if(!ids.length||new Set(ids).size!==ids.length)throw Error('請選取照片');const r=tx.objectStore('groups').getAll();r.onsuccess=checked(tx,()=>{if(r.result.some(g=>g.photoIds.some(id=>ids.includes(id))))throw Error('照片已加入紀錄集，請先移出紀錄集再刪除');const p=tx.objectStore('photos').getAll();p.onsuccess=checked(tx,()=>{if(ids.some(id=>!p.result.some(x=>x.id===id)))throw Error('照片已變更，請重新整理');for(const id of ids){tx.objectStore('photos').delete(id);tx.objectStore('assets').delete(id);}});});});
// Every read is queued synchronously in one transaction, including original blobs and groups.
export async function snapshot(){const db=await open();return new Promise((resolve,reject)=>{const result={startedAt:Date.now()};const tx=db.transaction(['photos','assets','groups'],'readonly');for(const name of ['photos','assets','groups']){const request=tx.objectStore(name).getAll();request.onsuccess=()=>result[name]=request.result;}tx.oncomplete=()=>resolve({...result,completedAt:Date.now()});tx.onerror=()=>reject(tx.error||Error('快照讀取失敗'));tx.onabort=()=>reject(tx.error||Error('快照交易已中止'));});}
