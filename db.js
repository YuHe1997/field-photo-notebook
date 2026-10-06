let connection;
export async function open(){if(connection)return connection;connection=await new Promise((resolve,reject)=>{const r=indexedDB.open('field-photo-notebook-v1',1);r.onupgradeneeded=()=>{for(const n of ['photos','assets','groups'])r.result.createObjectStore(n,{keyPath:'id'});};r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);r.onblocked=()=>reject(Error('請關閉其他使用本工具的分頁'));});return connection;}
export async function transaction(stores,action){const db=await open();return new Promise((resolve,reject)=>{const tx=db.transaction(stores,'readwrite');tx.oncomplete=()=>resolve();tx.onerror=()=>reject(tx.error);tx.onabort=()=>reject(tx.error||Error('交易已中止'));try{action(tx);}catch(e){tx.abort();reject(e);}});}
export async function all(store){const db=await open();return new Promise((resolve,reject)=>{const tx=db.transaction(store),r=tx.objectStore(store).getAll();let result;r.onsuccess=()=>result=r.result;tx.oncomplete=()=>resolve(result);tx.onerror=()=>reject(tx.error);tx.onabort=()=>reject(tx.error||Error('讀取交易已中止'));});}
export const savePhoto=(photo,blob)=>transaction(['photos','assets'],tx=>{tx.objectStore('photos').add(photo);tx.objectStore('assets').add({id:photo.id,original:blob});});
export const saveGroup=g=>transaction(['groups'],tx=>tx.objectStore('groups').put(g));
export const removeGroup=id=>transaction(['groups'],tx=>tx.objectStore('groups').delete(id));

// Queue every read synchronously in one transaction, including blobs and groups.
// Overlapping writers (including other tabs) are serialized by IndexedDB.
export async function snapshot(){
  const db=await open();
  return new Promise((resolve,reject)=>{
    const result={startedAt:Date.now()};
    const tx=db.transaction(['photos','assets','groups'],'readonly');
    for(const name of ['photos','assets','groups']){
      const request=tx.objectStore(name).getAll();
      request.onsuccess=()=>{result[name]=request.result;};
    }
    tx.oncomplete=()=>resolve({...result,completedAt:Date.now()});
    tx.onerror=()=>reject(tx.error||Error('快照讀取失敗'));
    tx.onabort=()=>reject(tx.error||Error('快照交易已中止'));
  });
}
