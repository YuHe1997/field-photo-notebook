import * as db from './db.js';
import {readingAt,validateGroup,move,digest,tar,untar} from './core.js';
const $=id=>document.getElementById(id),say=s=>$('status').textContent=s;
let photos=[],assets=[],groups=[],selected=new Set(),multi=false,edit=null,editBaseline=null,detailId=null,urls={gallery:[],groups:[],detail:[],order:[]},stream=null,watch=null,gps=null,heading=null,sensorState='尚未啟用',active='camera',starting=false,busy=false,cameraSession=0;
const text=(tag,s)=>{const e=document.createElement(tag);e.textContent=s;return e;};
async function run(fn){try{await fn();}catch(e){say('操作／儲存失敗：'+e.message+'。資料未宣稱成功，請重試或先備份。');}}
async function load(){({photos,assets,groups}=await db.snapshot());photos.sort((a,b)=>photoTime(a)-photoTime(b)||a.id.localeCompare(b.id));render();}
const photoTime=p=>p.buttonPressedAt??p.capturedAt??p.importedAt??0;
const photoLabel=p=>p?`照片 ${String(photos.indexOf(p)+1).padStart(2,'0')}`:'照片未取得';
function releaseImages(pool){urls[pool].forEach(URL.revokeObjectURL);urls[pool]=[];}
function photoImage(id,pool){const a=assets.find(a=>a.id===id),p=photos.find(p=>p.id===id);if(!a)return text('p','原始影像未取得');const im=document.createElement('img');im.loading='lazy';im.alt=photoLabel(p);im.src=URL.createObjectURL(a.original);urls[pool].push(im.src);return im;}
function photoInfo(p){const box=text('div','');box.className='photo-info';if(!p)return box;const time=photoTime(p);box.append(text('p',`${p.source==='camera'?'按下拍照':'匯入'}時間：${time?new Date(time).toLocaleString():'未取得'}${p.source==='camera'?'（非精確曝光時間）':''}`));const l=p.location;box.append(text('p',l?`拍攝位置：${l.latitude.toFixed(6)}, ${l.longitude.toFixed(6)}；定位誤差約 ±${Math.round(l.accuracyM)} 公尺；${l.freshnessAtAcquisitionCompleted??l.freshness??'新鮮度未知'}`:'拍攝位置：未取得（匯入照片不補填目前位置）'));const h=p.heading;box.append(text('p',h?`裝置羅盤參考：${h.degrees.toFixed(1)}°；${h.freshnessAtAcquisitionCompleted??h.freshness??'新鮮度未知'}；${h.quality??'精度未保證'}；鏡頭朝向未確認`:'拍攝方向：未取得；鏡頭朝向未確認'));return box;}
function editorValue(){return edit?JSON.stringify({photoIds:edit.photoIds,title:$('title').value,notes:$('notes').value,engineering:$('engineering').value}):null;}
function editorDirty(){return !!edit&&editorValue()!==editBaseline;}
function canDiscardEditor(){return !editorDirty()||confirm('記事尚未儲存，確定放棄這次編輯？');}
window.addEventListener('beforeunload',e=>{if(editorDirty()||pending.length||busy){e.preventDefault();e.returnValue='';}});
function renderDetail(){releaseImages('detail');$('detail').replaceChildren();const p=photos.find(p=>p.id===detailId);if(!p)return;$('detail').append(text('h3',photoLabel(p)),photoImage(p.id,'detail'),photoInfo(p));const note=text('button','為這張照片記事');note.onclick=()=>openEditor(null,[p.id]);const data=text('details','');data.append(text('summary','技術資料（座標未寫入原照片 EXIF）'),text('pre',JSON.stringify(p,null,2)));$('detail').append(note,data);}
function stop(){cameraSession++;stream?.getTracks().forEach(t=>t.stop());stream=null;$('video').srcObject=null;$('shutter').disabled=true;if(watch!==null)navigator.geolocation.clearWatch(watch);watch=null;gps=null;heading=null;window.removeEventListener('deviceorientation',onOrientation);window.removeEventListener('deviceorientationabsolute',onOrientation);sensorState='感測已停止';sensors();}
function sensors(){ $('sensor').textContent=`${sensorState}；GPS：${gps?`${gps.latitude.toFixed(6)}, ${gps.longitude.toFixed(6)} ±${gps.accuracyM}m；${Date.now()-gps.readingAt<=10000?'新鮮':'過期'}；${gps.accuracyM>20?'約略':'精度非保證'}`:'未取得'}；裝置羅盤：${heading?`${heading.degrees.toFixed(1)}° ${Date.now()-heading.readingAt<=2000?'新鮮':'過期'}（僅供參考）`:'未取得絕對方位'}；鏡頭朝向：無法可靠判定` ;}
setInterval(sensors,1000);
function tab(name){if(name!=='camera')stop();active=name;for(const n of ['camera','photos','groups'])$(n).hidden=n!==name;document.querySelectorAll('[data-tab]').forEach(b=>b.setAttribute('aria-pressed',b.dataset.tab===name));}
document.querySelectorAll('[data-tab]').forEach(b=>b.onclick=()=>tab(b.dataset.tab));
$('start').onclick=()=>run(async()=>{
  if(starting)return;
  starting=true;
  let acquired=null;
  try{
    stop();const session=cameraSession;
    if(!isSecureContext)throw Error('相機需要 HTTPS 或 localhost');
    acquired=await navigator.mediaDevices.getUserMedia({video:{facingMode:{ideal:'environment'}},audio:false});
    if(session!==cameraSession||active!=='camera'||document.hidden){acquired.getTracks().forEach(t=>t.stop());return;}
    stream=acquired;$('video').srcObject=acquired;
    await $('video').play();
    // play() can resolve after stop, navigation or backgrounding.
    if(session!==cameraSession||active!=='camera'||document.hidden){acquired.getTracks().forEach(t=>t.stop());return;}
    $('shutter').disabled=false;
    sensorState='實際鏡頭：'+(acquired.getVideoTracks()[0].getSettings().facingMode||'未知');
    if(navigator.geolocation)watch=navigator.geolocation.watchPosition(p=>{
      if(session!==cameraSession)return;
      const c=p.coords;
      if(Number.isFinite(c.latitude)&&Number.isFinite(c.longitude)&&Number.isFinite(c.accuracy))gps={latitude:c.latitude,longitude:c.longitude,accuracyM:c.accuracy,readingAt:p.timestamp,receivedAt:Date.now(),source:'browser-geolocation',precision:c.accuracy>20?'約略':'非保證'};
    },e=>{if(session!==cameraSession)return;gps=null;sensorState='定位未取得：'+e.message;},{enableHighAccuracy:true,maximumAge:0,timeout:15000});
  }catch(e){
    acquired?.getTracks().forEach(t=>t.stop());
    stop();
    throw e;
  }finally{starting=false;}
});
$('stop').onclick=stop;document.addEventListener('visibilitychange',()=>{if(document.hidden)stop();});window.addEventListener('pagehide',stop);
function onOrientation(e){let degrees,source;if(Number.isFinite(e.webkitCompassHeading)&&e.webkitCompassHeading>=0){degrees=e.webkitCompassHeading;source='webkitCompassHeading';}else if(e.absolute===true&&Number.isFinite(e.alpha)){degrees=(360-e.alpha)%360;source='absolute-alpha-device-reference';}else{sensorState='相對旋轉不是絕對羅盤';return;}heading={degrees,source,readingAt:Date.now(),timestampMeaning:'瀏覽器收到讀值時間（非硬體時間）',accuracy:Number.isFinite(e.webkitCompassAccuracy)?e.webkitCompassAccuracy:null,quality:Number.isFinite(e.webkitCompassAccuracy)&&e.webkitCompassAccuracy<0?'平台回報不可靠':'精度未保證／僅供參考',northReference:'平台絕對參考；真北／磁北未保證',alpha:e.alpha,beta:e.beta,gamma:e.gamma,screenAngle:screen.orientation?.angle??window.orientation??null,cameraBearing:null,cameraTransformEvidence:'未驗證裝置／螢幕／鏡頭轉換'};}
$('orientation').onclick=()=>{if(active!=='camera'){say('請先切換拍攝');return;}const E=window.DeviceOrientationEvent; if(!E){say('不支援方向感測');return;}const permission=typeof E.requestPermission==='function'?E.requestPermission():Promise.resolve('granted');run(async()=>{if(await permission!=='granted')throw Error('方向權限被拒絕');if(active!=='camera'||document.hidden)return;window.addEventListener('deviceorientation',onOrientation);window.addEventListener('deviceorientationabsolute',onOrientation);sensorState='方向監聽已啟用';});};
let pending=[];
async function commitPending(){while(pending.length){const item=pending[0];say('正在儲存原始影像…');await db.savePhoto(item.photo,item.blob);pending.shift();}await load();say('已儲存於本機（未上傳）');}
const retry=text('button','重試未完成照片儲存');retry.onclick=()=>run(commitPending);$('camera').append(retry);
function buttonReading(reading,time,limit){
  const sample=readingAt(reading,time,limit);
  if(!sample)return null;
  const {ageAtShutterMs,...rest}=sample;
  return {...rest,ageAtButtonPressMs:ageAtShutterMs,freshnessReference:'按下拍照按鈕時間；不是曝光時間'};
}
function completionReading(reading,time,limit){
  return reading?{...reading,ageAtAcquisitionCompletedMs:time-reading.readingAt,freshnessAtAcquisitionCompleted:time-reading.readingAt<=limit?'新鮮':'過期'}:null;
}
$('shutter').onclick=()=>run(async()=>{
  if(busy||!stream)return;
  // Timestamp the actual gesture, before any queued storage work.
  const buttonPressedAt=Date.now(),session=cameraSession,track=stream.getVideoTracks()[0];
  const location=buttonReading(gps,buttonPressedAt,10000),direction=buttonReading(heading,buttonPressedAt,2000);
  busy=true;
  try{
    if(pending.length)await commitPending();
    if(session!==cameraSession||track.readyState==='ended')throw Error('相機已停止，請重新開啟後拍攝');
    const acquisitionStartedAt=Date.now();let blob,method,frameCopiedAt=null;
    try{
      if(!window.ImageCapture)throw Error();
      blob=await new ImageCapture(track).takePhoto();method='ImageCapture.takePhoto';
    }catch(e){
      if(session!==cameraSession||track.readyState==='ended')throw Error('相機已停止，未擷取影像');
      const v=$('video');if(!v.videoWidth)throw Error('相機影格尚未就緒');
      const c=document.createElement('canvas');c.width=v.videoWidth;c.height=v.videoHeight;
      c.getContext('2d').drawImage(v,0,0);frameCopiedAt=Date.now();
      blob=await new Promise(r=>c.toBlob(r,'image/jpeg',.92));method='瀏覽器影格擷取（非 RAW／感光元件原檔）';
    }
    const acquisitionCompletedAt=Date.now();
    if(!blob)throw Error('未取得影像');
    const id=crypto.randomUUID();
    pending.push({blob,photo:{id,source:'camera',capturedAt:null,buttonPressedAt,acquisitionStartedAt,acquisitionCompletedAt,frameCopiedAt,
      timingMeaning:'曝光時間未知；按鈕／呼叫開始／Blob 取得／影格複製均為瀏覽器裝置時鐘，非硬體曝光時間',
      importedAt:null,originalName:id+'.jpg',mime:blob.type,captureMethod:method,
      location:completionReading(location,acquisitionCompletedAt,10000),heading:completionReading(direction,acquisitionCompletedAt,2000),
      cameraFacing:track.getSettings().facingMode??'unknown',cameraBearing:null}});
    await commitPending();
  }finally{busy=false;}
});
$('import').onchange=()=>run(async()=>{for(const file of $('import').files){if(!file.type.startsWith('image/'))throw Error('請選取影像');pending.push({blob:file,photo:{id:crypto.randomUUID(),source:'gallery',originalName:file.name,mime:file.type,capturedAt:null,importedAt:Date.now(),fileLastModifiedAt:file.lastModified,location:null,heading:null,sourceExif:null,captureMethod:'原始檔案匯入；未解析 EXIF'}});}await commitPending();$('import').value='';});
function choose(id){if(multi){selected.has(id)?selected.delete(id):selected.add(id);}else selected=new Set([id]);detailId=id;render();if(!multi&&!edit)openEditor(null,[id]);}
function render(){releaseImages('gallery');releaseImages('groups');$('gallery').replaceChildren();const assigned=new Set(groups.flatMap(g=>g.photoIds));for(const p of photos.filter(p=>!assigned.has(p.id)||selected.has(p.id))){const card=text('div','');card.className='card';const b=text('button',`${selected.has(p.id)?'☑':'□'} ${photoLabel(p)} · ${p.source==='camera'?'拍攝':'相簿匯入'}`);b.setAttribute('aria-pressed',selected.has(p.id));b.prepend(photoImage(p.id,'gallery'));let timer,long=false;b.onpointerdown=e=>{if(e.button!==0)return;long=false;timer=setTimeout(()=>{long=true;multi=true;selected.add(p.id);render();},550);};b.onpointerup=b.onpointerleave=b.onpointercancel=()=>clearTimeout(timer);b.onclick=()=>{if(long){long=false;return;}choose(p.id);};card.append(b);$('gallery').append(card);}$('multi').textContent=multi?'結束多選':'選取多張';$('multi').setAttribute('aria-pressed',multi);$('selectionCount').textContent=`已選 ${selected.size} 張；待整理 ${photos.filter(p=>!assigned.has(p.id)).length} 張`;$('groupList').replaceChildren();if(!groups.length)$('groupList').append(text('p','尚無紀錄集。請到待整理照片選取照片並記事。'));for(const g of groups){const a=text('article','');a.append(text('h3',g.title),text('p',`${g.photoIds.length} 張 · 更新 ${new Date(g.updatedAt).toLocaleString()}`));const strip=text('div','');strip.className='group-photos';g.photoIds.forEach((id,i)=>{const p=photos.find(p=>p.id===id),figure=text('figure','');figure.append(photoImage(id,'groups'),text('figcaption',`${i+1}. ${photoLabel(p)}`),photoInfo(p));strip.append(figure);});a.append(strip);for(const [label,value] of [['工程描述',g.engineering],['共同備註',g.notes]]){const note=text('p',`${label}：${value||'未填寫'}`);note.className='note-text';a.append(note);}const e=text('button','編輯／排序照片');e.onclick=()=>openEditor(g);const d=text('button','移除紀錄集（保留原始照片）');d.onclick=()=>run(async()=>{if(edit?.id===g.id){say('請先關閉或儲存此紀錄集的編輯，再移除。');return;}if(confirm('只移除此紀錄集，原始照片仍保留？')){await db.removeGroup(g.id);await load();say('已移除紀錄集，照片仍保留');}});a.append(e,d);$('groupList').append(a);}renderDetail();if(edit)renderOrder();}
$('multi').onclick=()=>{multi=!multi;if(!multi)selected=new Set([...selected].slice(0,1));render();};
function openEditor(g,ids=[...selected]){if(!canDiscardEditor())return;edit=structuredClone(g||{id:crypto.randomUUID(),title:'',notes:'',engineering:'',photoIds:ids,createdAt:Date.now()});$('title').value=edit.title;$('notes').value=edit.notes;$('engineering').value=edit.engineering;editBaseline=editorValue();$('editor').hidden=false;renderOrder();$('editor').scrollIntoView({block:'start'});$('title').focus();}
function renderOrder(){releaseImages('order');$('order').replaceChildren();edit.photoIds.forEach((id,i)=>{const p=photos.find(p=>p.id===id),row=text('article','');row.append(text('h3',`${i+1}. ${photoLabel(p)}`),photoImage(id,'order'),photoInfo(p));for(const [label,delta] of [['上移',-1],['下移',1]]){const b=text('button',label);b.disabled=i+delta<0||i+delta>=edit.photoIds.length;b.onclick=()=>{edit.photoIds=move(edit.photoIds,i,delta);renderOrder();};row.append(b);}const r=text('button','從此組移除');r.onclick=()=>{edit.photoIds.splice(i,1);renderOrder();};row.append(r);$('order').append(row);});}
$('compose').onclick=()=>{if(!selected.size){say('請先選照片');return;}openEditor();};$('addSelected').onclick=()=>{edit.photoIds=[...new Set([...edit.photoIds,...selected])];renderOrder();};$('cancel').onclick=()=>{if(!canDiscardEditor())return;$('editor').hidden=true;edit=null;editBaseline=null;releaseImages('order');};$('save').onclick=()=>run(async()=>{if(!edit)return;if(!$('title').value.trim()){say('請填寫紀錄集名稱，再儲存。');$('title').focus();return;}const g={...edit,title:$('title').value,notes:$('notes').value,engineering:$('engineering').value,updatedAt:Date.now()};validateGroup(g,new Set(photos.map(p=>p.id)));await db.saveGroup(g);$('editor').hidden=true;edit=null;editBaseline=null;releaseImages('order');selected.clear();detailId=null;await load();say('紀錄集已儲存');tab('groups');});
$('backup').onclick=()=>run(async()=>{say('正在建立本機備份（交易快照；不含未提交資料）…');const snapshot=await db.snapshot();const entries=[],records=[];for(const p of snapshot.photos){const blob=snapshot.assets.find(a=>a.id===p.id)?.original;if(!blob)throw Error('原始影像缺失');const path=`originals/${p.id}.bin`;entries.push([path,blob]);records.push({...p,path,size:blob.size,sha256:await digest(blob)});}for(const g of snapshot.groups)validateGroup(g,new Set(snapshot.photos.map(p=>p.id)));const manifest={schema:1,createdAt:Date.now(),snapshotStartedAt:snapshot.startedAt,snapshotCompletedAt:snapshot.completedAt,completeness:'單一唯讀交易中的已提交資料；不含未儲存表單／待存照片／快照之後提交資料',photos:records,groups:snapshot.groups};entries.push(['manifest.json',new Blob([JSON.stringify(manifest,null,2)],{type:'application/json'})],['notes.txt',new Blob([snapshot.groups.map(g=>`${g.title}\n工程描述：${g.engineering}\n共同備註：${g.notes}\n順序：${g.photoIds.join(', ')}\n`).join('\n')])],['README.txt',new Blob(['原始影像在 originals/；檔名與 MIME 在 manifest.json。GPS 是裝置位置；方向不是可靠鏡頭朝向。相機讀值以按鈕時間為基準，不是曝光資料。舊版 capturedAt / ageAtShutterMs 亦是按鈕時間／年齡。沒有 EXIF 嵌入或雲端同步。'])]);const archive=await tar(entries),url=URL.createObjectURL(archive),a=document.createElement('a');a.href=url;a.download=`field-notebook-${Date.now()}.tar`;a.click();setTimeout(()=>URL.revokeObjectURL(url),60000);say('已交給瀏覽器下載；請確認下載完成（無法驗證使用者是否取消）');});
$('restore').onchange=()=>run(async()=>{const file=$('restore').files[0];if(!file)return;say('驗證備份中…');const entries=await untar(file),m=JSON.parse(await entries.get('manifest.json')?.text());if(m.schema!==1||!Array.isArray(m.photos)||!Array.isArray(m.groups))throw Error('不支援的備份');const ids=new Set(),gids=new Set(),staged=[];for(const p of m.photos){if(typeof p.id!=='string'||!/^[-a-zA-Z0-9]+$/.test(p.id)||ids.has(p.id)||photos.some(x=>x.id===p.id)||p.path!==`originals/${p.id}.bin`||typeof p.originalName!=='string'||!['camera','gallery'].includes(p.source)||typeof p.mime!=='string'||!p.mime.startsWith('image/'))throw Error('照片格式／ID 重複或與本機衝突');ids.add(p.id);const b=entries.get(p.path);if(!b||b.size!==p.size||await digest(b)!==p.sha256)throw Error('影像大小或雜湊不符');const {path,size,sha256,...photo}=p;staged.push({photo,blob:new Blob([b],{type:p.mime})});}for(const g of m.groups){if(typeof g.id!=='string'||typeof g.title!=='string'||typeof g.notes!=='string'||typeof g.engineering!=='string'||gids.has(g.id)||groups.some(x=>x.id===g.id))throw Error('紀錄集格式／ID 衝突');gids.add(g.id);validateGroup(g,ids);}if(!confirm(`已驗證 ${staged.length} 張照片、${m.groups.length} 組，匯入本機？`))return;await db.transaction(['photos','assets','groups'],tx=>{for(const x of staged){tx.objectStore('photos').add(x.photo);tx.objectStore('assets').add({id:x.photo.id,original:x.blob});}for(const g of m.groups)tx.objectStore('groups').add(g);});await load();say('備份已完整還原');$('restore').value='';});
$('persist').onclick=()=>run(async()=>{const ok=await navigator.storage?.persist?.();say(ok?'瀏覽器允許持續儲存，仍須備份':'未取得持續儲存保證，請備份');});
run(async()=>{await load();const s=await navigator.storage?.estimate?.();$('storage').textContent=s?`使用約 ${(s.usage/1048576).toFixed(1)} MB / 配額約 ${(s.quota/1048576).toFixed(0)} MB（可能變動）`:'無法取得配額';});
if('serviceWorker' in navigator&&isSecureContext)run(async()=>{await navigator.serviceWorker.register('./sw.js',{scope:'./'});await navigator.serviceWorker.ready;const update=()=>{$('offline').textContent=navigator.serviceWorker.controller?'離線殼已受控制；照片只存 IndexedDB，仍需備份。':'離線殼已安裝，重新載入後確認控制。';};update();navigator.serviceWorker.addEventListener('controllerchange',update);});
