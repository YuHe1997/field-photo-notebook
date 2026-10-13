import * as db from './db.js';
import {authorize,hasToken,parseDriveId,uploadDayDoc,uploadOriginalPhoto,listSubfolders,createFolder,ensureDailyFolder} from './drive.js';
import {preparePhotoGroups,photoTableDocx,tablePlan,monitoringDocx} from './docx.js';
import {
  MON_FIELDS,
  monKey,
  genericMonTemplate,
  MonitoringDraftManager,
  formatMonDate,
  safeGetLocalStorage
} from './mon-draft.js';
import {
  parseWorkScope,
  synchronizeBothStrata,
  getNextMainNumber,
  getNextChildNumber,
  insertTextAtCaret,
  handleWorkScopeKeyDown
} from './mon-numbering.js';
import {compactPhotoLines,directionLabel,readingAt,validateGroup,move,digest,tar,untar,localDay,validDay,workDayOf,numberedPhoto,numberedGroup,previewRotation,zoomCapability,snapZoom,cropRect,focusCapabilities,mapPreviewTapToSensor,deriveCameraBearing} from './core.js';
const $=id=>document.getElementById(id);
let overlayErrorText=null;
const say=s=>{
  $('status').textContent=s;
  if($('cameraOverlay')&&!$('cameraOverlay').hidden){
    overlayErrorText=s;
    if($('cameraOverlayStatus')){$('cameraOverlayStatus').textContent=s;$('cameraOverlayStatus').classList.add('is-notice');}
  }
};
let photos=[],assets=[],groups=[],selected=new Set(),multi=false,edit=null,editBaseline=null,urls={gallery:[],groups:[],order:[]},stream=null,watch=null,gps=null,heading=null,sensorState='尚未啟用',active='camera',starting=false,busy=false,rotating=false,cameraSession=0,zoomState=null,focusState=null,lastFocusedElement=null,directionStatus='尚未啟用',gpsStatus='尚未啟用',activeSessionQueue=Promise.resolve(),constraintApplying=false,commitPromise=null,activeStartupSession=0,sessionReady=false;
const text=(tag,s)=>{const e=document.createElement(tag);e.textContent=s;return e;};
async function run(fn){try{await fn();}catch(e){say('操作／儲存失敗：'+e.message+'。資料未宣稱成功，請重試或先備份。');}}
const DAY_KEY='field-photo-notebook-work-days';
function readDays(){try{const s=safeGetLocalStorage();const d=JSON.parse(s?.getItem(DAY_KEY)||'{}');return {active:validDay(d.active)?d.active:null,days:Array.isArray(d.days)?d.days.filter(validDay):[]};}catch{return {active:null,days:[]};}}
let dayState=readDays(),activeDay=dayState.active;
let monDisplayedDay=null,daySwitchGeneration=0,dayTransitioning=false;
function displayedDay(){return monDisplayedDay||$('monForm')?.dataset?.displayedDay||activeDay;}
const previewCache=new Map(),rotations=new Map();
function rememberDay(day){const next={active:day,days:[...new Set([...dayState.days,day])].sort()};try{safeGetLocalStorage()?.setItem(DAY_KEY,JSON.stringify(next));}catch{}dayState=next;activeDay=day;}
async function load(){await db.ensurePhotoIdentities();({photos,assets,groups}=await db.snapshot());photos.sort((a,b)=>photoTime(a)-photoTime(b)||(a.photoNumber??0)-(b.photoNumber??0)||a.id.localeCompare(b.id));if(!activeDay)rememberDay(photos.length?workDayOf(photos.at(-1)):localDay());if(!monDisplayedDay)monDisplayedDay=activeDay;if($('monForm')&&monDisplayedDay)$('monForm').dataset.displayedDay=monDisplayedDay;render();}
const photoTime=p=>p.buttonPressedAt??p.capturedAt??p.importedAt??0;
const photoLabel=p=>p?numberedPhoto(p):'照片未取得';
const groupTitle=ids=>numberedGroup(ids,photos);
const dayPhotos=()=>photos.filter(p=>workDayOf(p)===activeDay);
const groupInDay=g=>g.workDay===activeDay||g.photoIds.some(id=>photos.some(p=>p.id===id&&workDayOf(p)===activeDay));
function releaseImages(pool){urls[pool].forEach(URL.revokeObjectURL);urls[pool]=[];}
function photoImage(id,pool){const a=assets.find(a=>a.id===id),p=photos.find(p=>p.id===id);if(!a)return text('p','原始影像未取得');const im=document.createElement('img');im.loading='lazy';im.alt=photoLabel(p);const key=`${id}:${p.displayRotation??'auto'}`;if(previewCache.has(key)){im.src=previewCache.get(key);im.dataset.previewReady='true';im.dataset.rotation=p.displayRotation??rotations.get(id)??0;return im;}im.onerror=()=>{im.dataset.previewReady='true';im.dataset.previewError='true';im.alt=photoLabel(p)+'（原圖無法解碼，未刪除）';};im.src=URL.createObjectURL(a.original);urls[pool].push(im.src);im.onload=()=>{im.onload=null;try{const w=im.naturalWidth,h=im.naturalHeight,r=previewRotation(p,w,h);rotations.set(id,r);const swapped=r%180!==0,ow=swapped?h:w,oh=swapped?w:h,scale=Math.min(1,960/Math.max(ow,oh)),c=document.createElement('canvas');c.width=Math.max(1,Math.round(ow*scale));c.height=Math.max(1,Math.round(oh*scale));const ctx=c.getContext('2d');ctx.translate(c.width/2,c.height/2);ctx.rotate(r*Math.PI/180);ctx.drawImage(im,-w*scale/2,-h*scale/2,w*scale,h*scale);const result=c.toDataURL('image/jpeg',.9);if(previewCache.size>=60)previewCache.delete(previewCache.keys().next().value);previewCache.set(key,result);if(im.isConnected)im.src=result;im.dataset.rotation=r;im.dataset.previewReady='true';}catch{im.dataset.previewReady='true';say('此照片無法產生旋轉預覽；原圖仍保留。');}};return im;}
function photoInfo(p){
  const box=text('div','');
  box.className='photo-info';
  for(const line of compactPhotoLines(p,0).slice(1))box.append(text('p',line));
  if(p?.heading?.degrees!==undefined||p?.cameraBearing!==undefined||p?.cameraBearingEvidence){
    const details=text('div','');
    details.className='camera-direction-details';
    if(p?.heading?.degrees!==undefined&&p?.heading?.degrees!==null){
      details.append(text('span',`裝置羅盤參考：${directionLabel(p.heading.degrees)}（${p.heading.degrees.toFixed(1)}°） `));
    }
    if(p?.cameraBearing!==undefined&&p?.cameraBearing!==null){
      details.append(text('span',`鏡頭朝向估計：${directionLabel(p.cameraBearing)}（${p.cameraBearing.toFixed(1)}°，${p.cameraBearingEvidence||'推估'}）`));
    }else if(p?.cameraBearing===null&&p?.cameraBearingEvidence){
      details.append(text('span',`鏡頭朝向：未確認（${p.cameraBearingEvidence}）`));
    }
    box.append(details);
  }
  return box;
}
function editorValue(){return edit?JSON.stringify({photoIds:edit.photoIds,engineering:$('engineering').value}):null;}
function editorDirty(){return !!edit&&editorValue()!==editBaseline;}
function canDiscardEditor(){return !editorDirty()||confirm('記事尚未儲存，確定放棄這次編輯？');}
window.addEventListener('beforeunload',e=>{if(editorDirty()||pending.length||busy||rotating||monDraftManager.hasAnyUnsaved()){e.preventDefault();e.returnValue='';}});
function isolateBackground(inert){
  const bgEls=[
    $('cameraGuide'),
    document.querySelector('header'),
    document.querySelector('nav'),
    $('status'),
    $('globalRetry'),
    $('photos'),
    $('groups'),
    $('monitoring'),
    $('editor')
  ];
  if($('camera')){
    bgEls.push(...$('camera').querySelectorAll(':scope > button, :scope > div:not(#cameraOverlay), :scope > p:not(#sensor)'));
  }
  for(const el of bgEls){
    if(el){
      if(inert)el.setAttribute('inert','');
      else el.removeAttribute('inert');
    }
  }
}
function stop(){
  cameraSession++;
  activeStartupSession=0;
  sessionReady=false;
  starting=false;
  activeSessionQueue=Promise.resolve();
  constraintApplying=false;
  overlayErrorText=null;
  isolateBackground(false);
  stream?.getTracks().forEach(t=>t.stop());
  stream=null;
  $('video').srcObject=null;
  if(watch!==null)navigator.geolocation.clearWatch(watch);
  watch=null;gps=null;heading=null;
  zoomState?.cancelFrame?.();zoomState=null;
  focusState=null;
  if($('focusReticle'))$('focusReticle').hidden=true;
  if($('focusStatus'))$('focusStatus').textContent='';
  if($('cameraOverlayStatus'))$('cameraOverlayStatus').textContent='';
  if($('cameraOverlay'))$('cameraOverlay').hidden=true;
  document.body.classList.remove('camera-overlay-active');
  directionStatus=gpsStatus='已停止';
  window.removeEventListener('deviceorientation',onOrientation);
  window.removeEventListener('deviceorientationabsolute',onOrientation);
  sensorState='感測已停止';
  zoomUI();sensors();
}
async function closeCamera(){
  stop();tab('photos');
  const returnBtn=$('multi')||document.querySelector('[data-tab="photos"]');
  if(returnBtn?.focus){try{returnBtn.focus();}catch{}}
  if(pending.length){
    try{await commitPending();}
    catch(err){
      say('相機已關閉，但有 '+pending.length+' 張照片尚未完成儲存：'+err.message+'。照片保留於待存清單，請點「重試未完成照片儲存」。');
      updateRetryUI();
    }
  }
}
function sensors(){
  const now=Date.now(),fresh=heading?now-heading.readingAt<=2000:false;
  let bearingText;
  if(!fresh){
    bearingText='無法可靠判定（讀值已過期，無即時姿態）';
  }else if(heading?.cameraBearing!==null&&heading?.cameraBearing!==undefined){
    bearingText=`估計 ${directionLabel(heading.cameraBearing)}（${heading.cameraBearing.toFixed(1)}°，${heading.cameraBearingEvidence}）`;
  }else{
    bearingText=`無法可靠判定${heading?.cameraBearingEvidence?`（${heading.cameraBearingEvidence}）`:''}`;
  }
  $('sensor').textContent=`${sensorState}；GPS 狀態：${gpsStatus}；方向狀態：${directionStatus}；GPS：${gps?`${gps.latitude.toFixed(6)}, ${gps.longitude.toFixed(6)} ±${gps.accuracyM}m；${now-gps.readingAt<=10000?'新鮮':'過期'}；${gps.accuracyM>20?'約略':'精度非保證'}`:'未取得'}；裝置羅盤：${heading?`裝置參考推估：${directionLabel(heading.degrees)}；${heading.degrees.toFixed(1)}°；${heading.quality}；${fresh?'新鮮':'過期'}（僅供參考）`:'未取得絕對方位'}；鏡頭朝向：${bearingText}`;
  if($('cameraOverlayStatus')){
    $('cameraOverlayStatus').classList.toggle('is-notice',Boolean(overlayErrorText));
    if(overlayErrorText){
      $('cameraOverlayStatus').textContent=overlayErrorText;
    }else if(starting){
      $('cameraOverlayStatus').textContent='正在啟動相機…';
    }else if(!fresh){
      $('cameraOverlayStatus').textContent=heading?`裝置：${directionLabel(heading.degrees)}（讀值過期）`:'等待方位';
    }else if(heading?.cameraBearing!==null&&heading?.cameraBearing!==undefined){
      $('cameraOverlayStatus').textContent=`鏡頭估計：${directionLabel(heading.cameraBearing)}（${heading.cameraBearing.toFixed(1)}°）`;
    }else if(heading){
      $('cameraOverlayStatus').textContent='鏡頭方向未確認';
    }else{
      $('cameraOverlayStatus').textContent='等待方位';
    }
  }
}
setInterval(sensors,1000);
function tab(name){if(name==='editor'&&!edit)return;if(name!=='camera')stop();active=name;for(const n of ['camera','photos','groups','monitoring'])$(n).hidden=n!==name;$('editor').hidden=name!=='editor'||!edit;$('editingTab').hidden=$('appendToEditor').hidden=!edit;document.querySelectorAll('[data-tab]').forEach(b=>b.setAttribute('aria-pressed',b.dataset.tab===name));if(name==='editor'){renderOrder();$('editor').scrollIntoView({block:'start'});const h=$('editor').querySelector('h2');h.tabIndex=-1;h.focus({preventScroll:true});}if(name==='monitoring')updateMonUI();}
window.tab=tab;
document.querySelectorAll('[data-tab]').forEach(b=>b.onclick=()=>tab(b.dataset.tab));
function currentCamera(session){return session===cameraSession&&active==='camera'&&!document.hidden;}
function beginDirection(session){const E=window.DeviceOrientationEvent;if(!E){directionStatus='不支援方向感測';sensors();return;}directionStatus='等待方向授權';let permission;try{permission=typeof E.requestPermission==='function'?E.requestPermission(true):Promise.resolve('granted');}catch(e){permission=Promise.reject(e);}Promise.resolve(permission).then(result=>{if(!currentCamera(session))return;if(result!=='granted'){directionStatus='方向權限被拒絕';sensors();return;}window.addEventListener('deviceorientation',onOrientation);window.addEventListener('deviceorientationabsolute',onOrientation);directionStatus='已啟用，等待絕對方位讀值';sensors();},e=>{if(currentCamera(session)){directionStatus='方向授權失敗：'+e.message;sensors();}});}
function setCameraDetails(open){
  $('cameraDetails').hidden=!open;
  $('cameraInfo').setAttribute('aria-expanded',String(open));
}
$('cameraInfo').onclick=()=>setCameraDetails($('cameraDetails').hidden);
$('cameraDetailsClose').onclick=()=>{setCameraDetails(false);$('cameraInfo').focus({preventScroll:true});};
function previewDimensions(){
  const v=$('video'),preview=$('cameraPreview'),vp=$('cameraViewport');
  if(!v||!preview||!vp)return;
  const vw=v.videoWidth,vh=v.videoHeight;
  if(!vw||!vh)return;
  const aspect=vw/vh;
  preview.style.setProperty('--camera-aspect',aspect);
  vp.style.setProperty('--camera-aspect',aspect);
  const stage=preview.getBoundingClientRect();
  if(!stage.width||!stage.height)return;
  let fitW,fitH;
  if(stage.width/stage.height>aspect){fitH=stage.height;fitW=stage.height*aspect;}
  else{fitW=stage.width;fitH=stage.width/aspect;}
  vp.style.width=Math.floor(fitW)+'px';
  vp.style.height=Math.floor(fitH)+'px';
}
if(window.ResizeObserver&&$('cameraPreview')){new ResizeObserver(previewDimensions).observe($('cameraPreview'));}
$('video').addEventListener('loadedmetadata',previewDimensions);$('video').addEventListener('resize',previewDimensions);
function queueTrackConstraint(session,op){
  if(!currentCamera(session))return Promise.resolve();
  constraintApplying=true;zoomUI();
  activeSessionQueue=activeSessionQueue.catch(()=>{}).then(async()=>{
    if(!currentCamera(session))return;
    const track=stream?.getVideoTracks()[0];
    if(!track||track.readyState==='ended')return;
    await op(track);
  }).finally(()=>{
    if(session===cameraSession){constraintApplying=false;zoomUI();}
  });
  return activeSessionQueue;
}
function zoomUI(){const z=zoomState,disabled=!z||z.applying||constraintApplying||busy;for(const id of ['zoom','zoomIn','zoomOut'])$(id).disabled=disabled;$('zoom').disabled=!z||busy;$('start').disabled=starting||busy;$('shutter').disabled=!stream||starting||busy||!sessionReady||!!z?.applying||constraintApplying||z?.factor===null||z?.frameReady===false;if(!z){$('zoomValue').textContent='開啟相機後可縮放';$('zoomNote').textContent='支援時使用鏡頭 API；否則使用數位裁切，畫質會降低。';$('video').style.transform='none';return;}$('zoom').min=z.range.min;$('zoom').max=z.range.max;$('zoom').step=z.range.step;$('zoom').value=z.requested??z.factor??z.range.min;$('zoomValue').textContent=z.applying||constraintApplying?'正在套用縮放…':z.factor===null?'倍率未確認，請重新開啟':`${Number(z.factor.toFixed(2))}×`;$('zoomNote').textContent=z.mode==='camera-api'?'鏡頭 API 縮放（不保證光學）；變焦拍攝保存目前影格 JPEG，解析度依手機。':'數位裁切（非光學）：預覽與新照片一同裁切，放大越多解析度越低。';$('video').style.transform=z.mode==='browser-digital-crop'?`scale(${z.factor})`:'none';if(z.factor===null)for(const id of ['zoom','zoomIn','zoomOut'])$(id).disabled=true;}
function initZoom(track,session,settings){let cap=null;try{cap=track.getCapabilities?.();}catch{}const range=zoomCapability(cap,settings.zoom);zoomState={track,session,mode:range?'camera-api':'browser-digital-crop',range:range??{min:1,max:4,step:.1},factor:range?settings.zoom:1,applying:false,requested:null,frameReady:true};previewDimensions();zoomUI();}
function nextZoomFrame(z){const v=$('video');return new Promise((resolve,reject)=>{let frame,tick,timer,finished=false;const done=error=>{if(finished)return;finished=true;clearTimeout(timer);if(frame!==undefined)v.cancelVideoFrameCallback?.(frame);if(tick!==undefined)cancelAnimationFrame(tick);z.cancelFrame=null;error?reject(error):resolve();};z.cancelFrame=()=>done(Error('相機已停止'));timer=setTimeout(()=>done(Error('縮放後影格未更新，請重試或重新開啟')),4000);if(v.requestVideoFrameCallback)frame=v.requestVideoFrameCallback(()=>done());else{const before=v.currentTime;const check=()=>{if(v.currentTime>before)done();else tick=requestAnimationFrame(check);};tick=requestAnimationFrame(check);}});}
async function requestZoom(value){
  const z=zoomState;if(!z||busy||z.factor===null)return;
  z.requested=snapZoom(value,z.range);if(z.applying)return;
  z.applying=true;zoomUI();
  try{
    while(z.requested!==null&&zoomState===z&&currentCamera(z.session)){
      const next=z.requested;z.requested=null;
      if(z.mode==='camera-api'){
        await queueTrackConstraint(z.session,async track=>{
          await track.applyConstraints({advanced:[{zoom:next}]});
          if(zoomState!==z||!currentCamera(z.session))return;
          const actual=track.getSettings().zoom;
          if(!Number.isFinite(actual)||actual<z.range.min||actual>z.range.max){
            z.factor=null;throw Error('手機未回報可確認的倍率');
          }
          z.factor=actual;z.frameReady=false;
          await nextZoomFrame(z);
          if(zoomState!==z||!currentCamera(z.session))return;
          z.frameReady=true;
          if(Math.abs(actual-next)>Math.max(1e-6,z.range.step/2))throw Error('手機未套用要求倍率，顯示仍以實際讀值為準');
        });
      }else{
        z.factor=next;
        $('video').style.transform=`scale(${z.factor})`;
      }
    }
  }catch(e){
    if(zoomState===z&&currentCamera(z.session)){
      if(z.mode==='camera-api'){
        try{
          const actual=z.track.getSettings().zoom;
          const valid=Number.isFinite(actual)&&actual>=z.range.min&&actual<=z.range.max;
          if(!valid||actual!==z.factor)z.frameReady=false;
          z.factor=valid?actual:null;
        }catch{z.factor=null;z.frameReady=false;}
      }
      z.requested=null;say('縮放未成功：'+e.message+'；未宣稱已放大，可重新開啟相機。');
    }
  }finally{z.applying=false;if(zoomState===z)zoomUI();}
}
$('zoom').oninput=()=>run(()=>requestZoom(Number($('zoom').value)));$('zoomIn').onclick=()=>run(()=>requestZoom((zoomState?.factor??1)+Math.max(zoomState?.range.step??.1,.5)));$('zoomOut').onclick=()=>run(()=>requestZoom((zoomState?.factor??1)-Math.max(zoomState?.range.step??.1,.5)));
$('start').onclick=()=>run(async()=>{
  if(starting||busy||active!=='camera')return;
  stop();
  const session=cameraSession;
  starting=true;
  activeStartupSession=session;
  sessionReady=false;
  zoomUI();
  let acquired=null;
  try{
    if(!isSecureContext)throw Error('相機需要 HTTPS 或 localhost');
    // Synchronously open overlay and isolate background during user click gesture before any await!
    if($('cameraOverlay')){$('cameraOverlay').hidden=false;setCameraDetails(false);}
    document.body.classList.add('camera-overlay-active');
    isolateBackground(true);
    overlayErrorText=null;
    if($('cameraOverlayStatus'))$('cameraOverlayStatus').textContent='正在啟動相機…';
    $('cameraClose').focus();
    // iOS needs requestPermission directly inside this click, before any await.
    beginDirection(session);
    acquired=await navigator.mediaDevices.getUserMedia({video:{facingMode:{ideal:'environment'},aspectRatio:{ideal:4/3},width:{ideal:1600},height:{ideal:1200}},audio:false});
    if(!currentCamera(session)){acquired.getTracks().forEach(t=>t.stop());return;}
    stream=acquired;$('video').srcObject=acquired;
    await $('video').play();
    if(!currentCamera(session)){acquired.getTracks().forEach(t=>t.stop());return;}
    previewDimensions();
    const track=acquired.getVideoTracks()[0],settings=track.getSettings();
    sensorState='實際鏡頭：'+(settings.facingMode||'未知');initZoom(track,session,settings);initFocus(track,session);
    gpsStatus=navigator.geolocation?'等待定位授權／讀值':'不支援定位';
    if(navigator.geolocation)watch=navigator.geolocation.watchPosition(p=>{
      if(!currentCamera(session))return;const c=p.coords;
      if(Number.isFinite(c.latitude)&&Number.isFinite(c.longitude)&&Number.isFinite(c.accuracy)){gps={latitude:c.latitude,longitude:c.longitude,accuracyM:c.accuracy,readingAt:p.timestamp,receivedAt:Date.now(),source:'browser-geolocation',precision:c.accuracy>20?'約略':'非保證'};gpsStatus='已取得定位';}sensors();
    },e=>{if(!currentCamera(session))return;gps=null;gpsStatus='定位未取得：'+e.message;sensors();},{enableHighAccuracy:true,maximumAge:0,timeout:15000});sensors();
    sessionReady=true;
    sensors();
    if(!$('shutter').disabled)$('shutter').focus();
  }catch(e){
    acquired?.getTracks().forEach(t=>t.stop());
    if(session===cameraSession){stop();throw e;}
  }finally{
    if(session===activeStartupSession){
      starting=false;
      activeStartupSession=0;
      zoomUI();
    }
  }
});
$('stop').onclick=stop;
if($('cameraClose'))$('cameraClose').onclick=()=>run(closeCamera);
window.addEventListener('keydown',e=>{
  if(active!=='camera'||$('cameraOverlay')?.hidden)return;
  if(e.key==='Escape'){e.preventDefault();run(closeCamera);return;}
  if(e.key==='Tab'){
    const overlay=$('cameraOverlay');
    const focusables=[...overlay.querySelectorAll('button:not(:disabled), input:not(:disabled), [tabindex="0"]')].filter(el=>el.offsetParent!==null);
    if(!focusables.length)return;
    const first=focusables[0],last=focusables[focusables.length-1];
    if(e.shiftKey&&document.activeElement===first){e.preventDefault();last.focus();}
    else if(!e.shiftKey&&document.activeElement===last){e.preventDefault();first.focus();}
  }
});
document.addEventListener('visibilitychange',()=>{if(document.hidden)stop();});window.addEventListener('pagehide',stop);
function initFocus(track,session){let sup={},cap={};try{sup=navigator.mediaDevices?.getSupportedConstraints?.()??{};}catch{}try{cap=track.getCapabilities?.()??{};}catch{}const caps=focusCapabilities(cap,sup);focusState={track,session,caps,sup,applying:false,pendingTap:null};if($('focusStatus'))$('focusStatus').textContent=caps.summary;}
async function handlePreviewTap(e){
  const fs=focusState;if(!fs||!stream||busy||!currentCamera(fs.session))return;
  const viewport=$('cameraViewport'),video=$('video');if(!viewport||!video||!video.videoWidth)return;
  const vpRect=viewport.getBoundingClientRect(),reticle=$('focusReticle'),status=$('focusStatus');
  const inVp=e.clientX>=vpRect.left&&e.clientX<=vpRect.right&&e.clientY>=vpRect.top&&e.clientY<=vpRect.bottom;
  if(!inVp){if(status)status.textContent='點按位置在黑邊（letterbox）區域，請點按鏡頭影像區域';return;}
  const tapX=e.clientX-vpRect.left,tapY=e.clientY-vpRect.top,z=zoomState;
  const mapping=mapPreviewTapToSensor(tapX,tapY,vpRect.width,vpRect.height,video.videoWidth,video.videoHeight,z?.mode,z?.factor);
  if(!mapping.inBounds){if(status)status.textContent=mapping.reason;return;}
  if(reticle){reticle.style.left=`${tapX}px`;reticle.style.top=`${tapY}px`;reticle.className='focus-reticle requesting';reticle.hidden=false;}
  if(!fs.caps.canPointFocus){
    if(reticle)reticle.className='focus-reticle unsupported';
    if(status)status.textContent='此裝置不支援點對焦座標，維持自動模式';
    setTimeout(()=>{if(reticle&&focusState===fs)reticle.hidden=true;},1500);
    return;
  }
  fs.pendingTap=mapping;if(fs.applying)return;fs.applying=true;
  try{
    while(fs.pendingTap&&focusState===fs&&currentCamera(fs.session)){
      const next=fs.pendingTap;fs.pendingTap=null;
      if(status)status.textContent='正在要求點對焦…';
      try{
        await queueTrackConstraint(fs.session,async track=>{
          if(!currentCamera(fs.session))return;
          const advanced=[];
          const curZoom=(zoomState?.mode==='camera-api'&&Number.isFinite(zoomState.factor))?zoomState.factor:null;
          if(curZoom)advanced.push({zoom:curZoom});
          if(fs.caps.pointsOfInterest)advanced.push({pointsOfInterest:[{x:next.sensorX,y:next.sensorY}]});
          const reqFocusMode=fs.sup?.focusMode?(fs.caps.focusModes.includes('single-shot')?'single-shot':(fs.caps.focusModes.includes('continuous')?'continuous':null)):null;
          const reqExposureMode=fs.sup?.exposureMode?(fs.caps.exposureModes.includes('continuous')?'continuous':null):null;
          if(reqFocusMode)advanced.push({focusMode:reqFocusMode});
          if(reqExposureMode)advanced.push({exposureMode:reqExposureMode});
          await track.applyConstraints({advanced});
          if(!currentCamera(fs.session))return;
          let s={};try{s=track.getSettings();}catch{}
          if(zoomState?.mode==='camera-api'){
            const reportedZoom=s.zoom;
            const isZoomValid=Number.isFinite(reportedZoom)&&reportedZoom>=(zoomState.range.min-1e-4)&&reportedZoom<=(zoomState.range.max+1e-4);
            if(!isZoomValid){
              zoomState.factor=null;
              zoomState.frameReady=false;
              zoomUI();
            }else{
              zoomState.factor=reportedZoom;
              zoomState.frameReady=false;
              zoomUI();
              await nextZoomFrame(zoomState);
              if(!currentCamera(fs.session))return;
              zoomState.frameReady=true;
              zoomUI();
            }
          }else{
            zoomState.frameReady=false;
            zoomUI();
            await nextZoomFrame(zoomState);
            if(!currentCamera(fs.session))return;
            zoomState.frameReady=true;
            zoomUI();
          }
          const poiReported=Array.isArray(s.pointsOfInterest)&&s.pointsOfInterest.length>0;
          let poiMatched=false;
          if(poiReported){
            const read=s.pointsOfInterest[0];
            poiMatched=Number.isFinite(read?.x)&&Number.isFinite(read?.y)&&Math.abs(read.x-next.sensorX)<=0.05&&Math.abs(read.y-next.sensorY)<=0.05;
          }
          let coordText;
          if(poiMatched){
            coordText='點對焦已送出，裝置回讀座標相符';
          }else if(poiReported){
            coordText='點對焦已送出（裝置回讀座標與請求不符，未確認命中）';
          }else{
            coordText='點對焦已送出（裝置未回讀座標，未確認命中）';
          }
          let modeNotes=[];
          if(reqFocusMode){
            if(!s.focusMode){modeNotes.push('對焦模式未回讀');}
            else if(s.focusMode!==reqFocusMode){modeNotes.push('對焦模式不符');}
            else{modeNotes.push('對焦模式相符');}
          }
          if(reqExposureMode){
            if(!s.exposureMode){modeNotes.push('曝光模式未回讀');}
            else if(s.exposureMode!==reqExposureMode){modeNotes.push('曝光模式不符');}
            else{modeNotes.push('曝光模式相符');}
          }
          const modeStr=modeNotes.length?`（${modeNotes.join('、')}）`:'';
          if(status)status.textContent=coordText+(modeStr?`；${modeStr}`:'');
          if(reticle){
            const allModesConfirmed=!modeNotes.some(n=>n.includes('未')||n.includes('不符'));
            reticle.className=(poiMatched&&allModesConfirmed)?'focus-reticle confirmed':'focus-reticle unconfirmed';
          }
        });
      }catch(err){
        if(!currentCamera(fs.session))return;
        if(status)status.textContent='對焦請求未被接受：'+err.message+'；維持既有模式';
        if(reticle)reticle.className='focus-reticle failed';
      }
    }
  }finally{fs.applying=false;setTimeout(()=>{if(reticle&&focusState===fs)reticle.hidden=true;},2000);}
}
$('cameraPreview').onclick=e=>run(()=>handlePreviewTap(e));
function onOrientation(e){
  let degrees,source;
  const isAbsolute=e.absolute===true||e.type==='deviceorientationabsolute';
  const hasValidAlpha=Number.isFinite(e.alpha);
  const hasAnyReading=hasValidAlpha||Number.isFinite(e.beta)||Number.isFinite(e.gamma);
  if(e.isTrusted&&!hasAnyReading&&!Number.isFinite(e.webkitCompassHeading)){
    if(!heading)directionStatus='等待絕對方位讀值';
    return;
  }
  if(Number.isFinite(e.webkitCompassHeading)&&e.webkitCompassHeading>=0){
    degrees=e.webkitCompassHeading;source='webkitCompassHeading';
  }else if(isAbsolute&&hasValidAlpha){
    degrees=(360-e.alpha)%360;source='absolute-alpha-device-reference';
  }else if(!isAbsolute&&hasValidAlpha){
    directionStatus='相對旋轉不是絕對羅盤';
    if(heading){
      heading.cameraBearing=null;
      heading.cameraBearingLabel=null;
      heading.cameraBearingAlgorithm=null;
      heading.cameraBearingEvidence='最新感測為相對旋轉，無有效即時姿態';
      heading.cameraTransformEvidence=heading.cameraBearingEvidence;
    }
    sensors();
    return;
  }else{
    directionStatus='方位讀值缺值或非數值';
    if(heading){
      heading.cameraBearing=null;
      heading.cameraBearingLabel=null;
      heading.cameraBearingAlgorithm=null;
      heading.cameraBearingEvidence='最新方位讀值缺值或非數值，無有效即時姿態';
      heading.cameraTransformEvidence=heading.cameraBearingEvidence;
    }
    sensors();
    return;
  }
  directionStatus='已取得裝置參考（需驗證）';
  const track=stream?.getVideoTracks()[0],facing=track?.getSettings().facingMode??'unknown';
  const derived=deriveCameraBearing({alpha:e.alpha,beta:e.beta,gamma:e.gamma,source,absolute:e.absolute===true,accuracy:Number.isFinite(e.webkitCompassAccuracy)?e.webkitCompassAccuracy:null,quality:Number.isFinite(e.webkitCompassAccuracy)&&e.webkitCompassAccuracy<0?'平台回報不可靠':'精度未保證／僅供參考',facingMode:facing,screenAngle:screen.orientation?.angle??window.orientation??null,readingAt:Date.now()});
  heading={degrees,source,readingAt:Date.now(),timestampMeaning:'瀏覽器收到讀值時間（非硬體時間）',accuracy:Number.isFinite(e.webkitCompassAccuracy)?e.webkitCompassAccuracy:null,quality:Number.isFinite(e.webkitCompassAccuracy)&&e.webkitCompassAccuracy<0?'平台回報不可靠':'精度未保證／僅供參考',northReference:'平台絕對參考；真北／磁北未保證',alpha:e.alpha,beta:e.beta,gamma:e.gamma,screenAngle:screen.orientation?.angle??window.orientation??null,cameraBearing:derived.degrees,cameraBearingLabel:derived.label,cameraBearingAlgorithm:derived.algorithm,cameraBearingEvidence:derived.degrees!==null?derived.evidence:derived.reason,cameraTransformEvidence:derived.degrees!==null?derived.evidence:derived.reason};
  sensors();
}
let pending=[];
function updateRetryUI(){const has=pending.length>0;retry.hidden=!has;const o=$('overlayRetry');if(o)o.hidden=!has;const g=$('globalRetry');if(g)g.hidden=!has;}
async function commitPending(){
  if(commitPromise)return commitPromise;
  commitPromise=(async()=>{
    while(pending.length){
      const item=pending[0];
      say('正在儲存原始影像…');
      try{
        await db.savePhoto(item.photo,item.blob);
        pending.shift();
      }catch(err){
        say('儲存原始影像失敗：'+err.message+'；照片保留於待存清單，請點「重試未完成照片儲存」。');
        updateRetryUI();
        throw err;
      }
    }
    await load();
    say('已儲存於本機（未上傳）');
    updateRetryUI();
  })().finally(()=>{commitPromise=null;});
  return commitPromise;
}
const retry=text('button','重試未完成照片儲存');retry.onclick=()=>run(commitPending);$('camera').append(retry);
if($('overlayRetry'))$('overlayRetry').onclick=()=>run(commitPending);
if($('globalRetry'))$('globalRetry').onclick=()=>run(commitPending);
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
  if(constraintApplying){await activeSessionQueue.catch(()=>{});}
  if(busy||!stream||zoomState?.applying||zoomState?.factor===null||zoomState?.frameReady===false)return;
  // Timestamp the actual gesture, before any queued storage work.
  const buttonPressedAt=Date.now(),session=cameraSession,track=stream.getVideoTracks()[0],workDay=activeDay,screenAngle=screen.orientation?.angle??window.orientation??null;
  const location=buttonReading(gps,buttonPressedAt,10000),direction=buttonReading(heading,buttonPressedAt,2000);
  const missing=[];if(!location||location.freshness==='過期')missing.push('GPS 未取得或過期');if(!direction||direction.freshness==='過期'||direction.accuracy<0)missing.push('方向未取得、過期或不可靠');if(missing.length&&!confirm(missing.join('；')+'。建議等待讀值／重新開啟。仍要拍照並保存缺值或未確認資料？'))return;
  const z=zoomState,captureSettings=track.getSettings(),reportedZoom=Number.isFinite(captureSettings.zoom)?captureSettings.zoom:null,zoom={mode:z?.mode??'unknown',factor:z?.mode==='camera-api'?reportedZoom:(z?.factor??null),reportedTrackZoom:reportedZoom,meaning:'API 倍率非光學認證；數位倍率為目前影格的中央裁切；曝光時間未知'},useZoomFrame=zoom.factor!==1;
  if(z?.mode==='camera-api'&&!zoomCapability({zoom:z.range},reportedZoom)){z.factor=null;z.frameReady=false;zoomUI();throw Error('拍照時相機未回報有效倍率，請重新開啟');}
  const currentFacing=captureSettings.facingMode??'unknown';
  const isFacingRear=currentFacing==='environment';
  const isDirectionFresh=direction&&direction.freshness==='新鮮';
  const isBearingValid=isDirectionFresh&&isFacingRear&&direction.cameraBearing!==null;
  const cBearing=isBearingValid?direction.cameraBearing:null;
  const cLabel=isBearingValid?direction.cameraBearingLabel:null;
  const cAlg=isBearingValid?direction.cameraBearingAlgorithm:null;
  let cEvidence;
  if(!isFacingRear){
    cEvidence=`目前鏡頭為 ${currentFacing}，僅後鏡頭（environment）支援鏡頭朝向估計`;
  }else if(!direction||!Number.isFinite(direction.readingAt)){
    cEvidence='未取得方位讀值，無即時鏡頭朝向';
  }else if(direction&&direction.cameraBearing===null){
    cEvidence=direction.cameraBearingEvidence||'姿態未提供有效即時方位';
  }else if(!isDirectionFresh){
    cEvidence='方位讀值已過期，無法提供即時鏡頭朝向估計';
  }else{
    cEvidence=direction?.cameraBearingEvidence??null;
  }
  busy=true;zoomUI();overlayErrorText=null;
  try{
    if(pending.length)await commitPending();
    if(session!==cameraSession||track.readyState==='ended')throw Error('相機已停止，請重新開啟後拍攝');
    const acquisitionStartedAt=Date.now();let blob,method,frameCopiedAt=null,outputWidth=null,outputHeight=null;
    try{
      if(useZoomFrame||!window.ImageCapture)throw Error();
      blob=await new ImageCapture(track).takePhoto();method='ImageCapture.takePhoto';
    }catch(e){
      if(session!==cameraSession||track.readyState==='ended')throw Error('相機已停止，未擷取影像');
      const v=$('video');if(!v.videoWidth)throw Error('相機影格尚未就緒');
      const crop=cropRect(v.videoWidth,v.videoHeight,zoom.mode==='browser-digital-crop'?zoom.factor:1),c=document.createElement('canvas');c.width=crop.width;c.height=crop.height;
      c.getContext('2d').drawImage(v,crop.sx,crop.sy,crop.sw,crop.sh,0,0,c.width,c.height);frameCopiedAt=Date.now();outputWidth=c.width;outputHeight=c.height;
      blob=await new Promise(r=>c.toBlob(r,'image/jpeg',.92));method=useZoomFrame?'變焦瀏覽器影格擷取（數位／API 縮放，非 RAW）':'瀏覽器影格擷取（非 RAW／感光元件原檔）';
    }
    const acquisitionCompletedAt=Date.now();
    if(!blob)throw Error('未取得影像');
    const acquisitionAge=direction?(acquisitionCompletedAt-direction.readingAt):Infinity;
    const isAcquisitionFresh=isBearingValid&&acquisitionAge<=2000;
    const finalBearing=isAcquisitionFresh?cBearing:null;
    const finalLabel=isAcquisitionFresh?cLabel:null;
    const finalAlg=isAcquisitionFresh?cAlg:null;
    let finalEvidence;
    if(!isFacingRear){
      finalEvidence=cEvidence;
    }else if(!direction||!Number.isFinite(direction.readingAt)){
      finalEvidence='未取得方位讀值，無即時鏡頭朝向';
    }else if(direction&&direction.freshness==='過期'){
      finalEvidence='方位讀值於按鈕前已過期，無法提供即時鏡頭朝向估計';
    }else if(direction&&direction.cameraBearing===null){
      finalEvidence=direction.cameraBearingEvidence||'姿態未提供有效即時方位';
    }else if(!isAcquisitionFresh){
      finalEvidence=`影格取得完成時方位已過期（${Math.round(acquisitionAge/1000)} 秒），無即時鏡頭朝向`;
    }else{
      finalEvidence=cEvidence;
    }
    const headingCompleted=completionReading(direction,acquisitionCompletedAt,2000);
    if(headingCompleted){
      headingCompleted.cameraBearing=finalBearing;
      headingCompleted.cameraBearingLabel=finalLabel;
      headingCompleted.cameraBearingAlgorithm=finalAlg;
      headingCompleted.cameraBearingEvidence=finalEvidence;
      headingCompleted.cameraTransformEvidence=finalEvidence;
    }
    const id=crypto.randomUUID();
    pending.push({blob,photo:{id,source:'camera',capturedAt:null,buttonPressedAt,acquisitionStartedAt,acquisitionCompletedAt,frameCopiedAt,
      timingMeaning:'曝光時間未知；按鈕／呼叫開始／Blob 取得／影格複製均為瀏覽器裝置時鐘，非硬體曝光時間',
      importedAt:null,originalName:id+(blob.type==='image/png'?'.png':blob.type==='image/webp'?'.webp':'.jpg'),mime:blob.type,captureMethod:method,
      location:completionReading(location,acquisitionCompletedAt,10000),heading:headingCompleted,
      cameraFacing:currentFacing,cameraBearing:finalBearing,cameraBearingLabel:finalLabel,cameraBearingAlgorithm:finalAlg,cameraBearingEvidence:finalEvidence,cameraTransformEvidence:finalEvidence??'未驗證裝置／螢幕／鏡頭轉換',workDay,screenAngle,zoom,outputWidth,outputHeight}});
    await commitPending();
  }finally{busy=false;zoomUI();}
});
$('import').onchange=()=>run(async()=>{const workDay=activeDay;for(const file of $('import').files){if(!file.type.startsWith('image/'))throw Error('請選取影像');pending.push({blob:file,photo:{id:crypto.randomUUID(),source:'gallery',originalName:file.name,mime:file.type,capturedAt:null,importedAt:Date.now(),fileLastModifiedAt:file.lastModified,location:null,heading:null,sourceExif:null,workDay,captureMethod:'原始檔案匯入；未解析 EXIF，預覽由瀏覽器解碼方向' }});}await commitPending();$('import').value='';});
function choose(id){if(multi){selected.has(id)?selected.delete(id):selected.add(id);}else selected=selected.has(id)?new Set():new Set([id]);render();}
function render(){$('editingTab').hidden=$('appendToEditor').hidden=!edit;releaseImages('gallery');releaseImages('groups');$('gallery').replaceChildren();const days=[...new Set([...dayState.days,...photos.map(workDayOf),activeDay])].filter(validDay).sort();$('workDay').replaceChildren();for(const d of days){const o=text('option',`${d}（${photos.filter(p=>workDayOf(p)===d).length} 張）`);o.value=d;$('workDay').append(o);}$('workDay').value=activeDay;$('currentDay').textContent=`工作日：${activeDay}；編號按工作日獨立且不因刪除重新編排`;for(const id of selected)if(!dayPhotos().some(p=>p.id===id))selected.delete(id);const assigned=new Set(groups.flatMap(g=>g.photoIds));for(const p of dayPhotos().filter(p=>!assigned.has(p.id)||selected.has(p.id))){const card=text('div','');card.className='card';const b=text('button',`${selected.has(p.id)?'☑':'□'} ${photoLabel(p)} · ${p.source==='camera'?'拍攝':'相簿匯入'}`);b.setAttribute('aria-pressed',selected.has(p.id));b.prepend(photoImage(p.id,'gallery'));let timer,long=false;b.onpointerdown=e=>{if(e.button!==0)return;long=false;timer=setTimeout(()=>{long=true;multi=true;selected.add(p.id);render();},550);};b.onpointerup=b.onpointerleave=b.onpointercancel=()=>clearTimeout(timer);b.onclick=()=>{if(long){long=false;return;}choose(p.id);};card.append(b);$('gallery').append(card);}$('multi').textContent=multi?'結束多選':'選取多張';$('multi').setAttribute('aria-pressed',multi);$('selectionCount').textContent=`已選 ${selected.size} 張；本日待整理 ${dayPhotos().filter(p=>!assigned.has(p.id)).length} 張`;$('appendToEditor').disabled=$('deleteSelected').disabled=$('rotateLeft').disabled=$('rotateRight').disabled=!selected.size;$('groupList').replaceChildren();if(!groups.some(groupInDay))$('groupList').append(text('p','本工作日尚無紀錄集。請到待整理照片選取照片並記事。'));for(const g of groups.filter(groupInDay)){const a=text('article','');a.append(text('h3',groupTitle(g.photoIds)),text('p',`${g.photoIds.length} 張 · 更新 ${new Date(g.updatedAt).toLocaleString()}`));const strip=text('div','');strip.className='group-photos';g.photoIds.forEach((id,i)=>{const p=photos.find(p=>p.id===id),figure=text('figure','');figure.append(photoImage(id,'groups'),text('figcaption',compactPhotoLines(p,i)[0]));strip.append(figure);});a.append(strip);g.photoIds.forEach((id,i)=>{const row=text('div','');row.append(text('strong',compactPhotoLines(photos.find(p=>p.id===id),i)[0]),photoInfo(photos.find(p=>p.id===id)));a.append(row);});for(const [label,value] of [['工程描述',g.engineering]]){const note=text('p',`${label}：${value||'未填寫'}`);note.className='note-text';a.append(note);}const e=text('button','編輯／排序照片');e.onclick=()=>openEditor(g);const d=text('button','移除紀錄集（保留原始照片）');d.onclick=()=>run(async()=>{if(edit?.id===g.id){say('請先關閉或儲存此紀錄集的編輯，再移除。');return;}if(confirm('只移除此紀錄集，原始照片仍保留？')){await db.removeGroup(g.id);await load();say('已移除紀錄集，照片仍保留');}});a.append(e,d);$('groupList').append(a);}if(edit&&active==='editor')renderOrder();}
$('multi').onclick=()=>{multi=!multi;if(!multi)selected=new Set([...selected].slice(0,1));render();};
function openEditor(g,ids=[...selected]){if(!canDiscardEditor())return;edit=structuredClone(g||{id:crypto.randomUUID(),title:'',notes:'',engineering:'',photoIds:ids,createdAt:Date.now(),workDay:activeDay});$('title').value=groupTitle(edit.photoIds);$('engineering').value=edit.engineering;editBaseline=editorValue();tab('editor');}
function renderOrder(focus){$('title').value=groupTitle(edit.photoIds);releaseImages('order');$('order').replaceChildren();$('photoInfo').replaceChildren();const previews=text('div','');previews.className='editor-previews';previews.dataset.count=edit.photoIds.length;edit.photoIds.forEach((id,i)=>{const figure=text('figure','');figure.append(photoImage(id,'order'),text('figcaption',compactPhotoLines(photos.find(p=>p.id===id),i)[0]));previews.append(figure);});$('order').append(previews);edit.photoIds.forEach((id,i)=>{const p=photos.find(p=>p.id===id),row=text('article',''),head=text('div',''),controls=text('div','');head.className='photo-row-head';controls.className='photo-row-controls';head.append(text('h3',compactPhotoLines(p,i)[0]));for(const [label,delta] of [['上移',-1],['下移',1]]){const b=text('button',delta<0?'↑':'↓');b.title=label;b.setAttribute('aria-label',`${i+1}. ${photoLabel(p)} ${label}`);b.dataset.photoId=id;b.dataset.action=label;b.disabled=i+delta<0||i+delta>=edit.photoIds.length;b.onclick=()=>{edit.photoIds=move(edit.photoIds,i,delta);renderOrder({id,action:label});say(`${photoLabel(p)} 已${label}至第 ${edit.photoIds.indexOf(id)+1} 張`);};controls.append(b);}const r=text('button','×');r.title='從此組移除';r.setAttribute('aria-label',`${i+1}. ${photoLabel(p)} 從此組移除`);r.dataset.photoId=id;r.dataset.action='從此組移除';r.onclick=()=>{edit.photoIds.splice(i,1);renderOrder({id:edit.photoIds[Math.min(i,edit.photoIds.length-1)],action:'從此組移除'});say(`${photoLabel(p)} 已從此組移除，原始照片仍保留`);};controls.append(r);head.append(controls);row.append(head,photoInfo(p));$('photoInfo').append(row);});if(focus){const controls=[...$('photoInfo').querySelectorAll('button')].filter(b=>b.dataset.photoId===focus.id&&!b.disabled);(controls.find(b=>b.dataset.action===focus.action)||controls[0]||$('addSelected')).focus();}}
$('compose').onclick=()=>{if(!selected.size){say('請先選照片');return;}openEditor();};$('appendToEditor').onclick=()=>{if(!edit)return;const day=edit.workDay??workDayOf(photos.find(p=>p.id===edit.photoIds[0])??{});if([...selected].some(id=>workDayOf(photos.find(p=>p.id===id))!==day)){say('請勿將不同工作日的照片加入同一記事。');return;}edit.photoIds=[...new Set([...edit.photoIds,...selected])];tab('editor');};$('addSelected').onclick=()=>{if(edit){tab('photos');say('請選取照片，再按「加入目前記事並返回」。');}};$('cancel').onclick=()=>{if(!canDiscardEditor())return;const returnTab=groups.some(g=>g.id===edit?.id)?'groups':'photos';$('editor').hidden=true;edit=null;editBaseline=null;releaseImages('order');render();tab(returnTab);};$('save').onclick=()=>run(async()=>{if(!edit)return;const title=groupTitle(edit.photoIds),g={...edit,title,notes:edit.notes??'',engineering:$('engineering').value,updatedAt:Date.now()};if(edit.title&&edit.title!==title&&!edit.legacyTitle)g.legacyTitle=edit.title;validateGroup(g,new Set(photos.map(p=>p.id)));await db.saveGroup(g);$('editor').hidden=true;edit=null;editBaseline=null;releaseImages('order');selected.clear();await load();say('紀錄集已儲存');tab('groups');});
$('backup').onclick=()=>run(async()=>{say('正在建立本機備份（交易快照；不含未提交資料）…');const snapshot=await db.snapshot();const entries=[],records=[];for(const p of snapshot.photos){const blob=snapshot.assets.find(a=>a.id===p.id)?.original;if(!blob)throw Error('原始影像缺失');const path=`originals/${p.id}.bin`;entries.push([path,blob]);records.push({...p,path,size:blob.size,sha256:await digest(blob)});}for(const g of snapshot.groups)validateGroup(g,new Set(snapshot.photos.map(p=>p.id)));const manifest={schema:1,createdAt:Date.now(),snapshotStartedAt:snapshot.startedAt,snapshotCompletedAt:snapshot.completedAt,completeness:'單一唯讀交易中的已提交資料；不含未儲存表單／待存照片／快照之後提交資料',photos:records,groups:snapshot.groups,workDays:dayState.days,activeWorkDay:activeDay??localDay(),numberCounters:db.numberCounters()};entries.push(['manifest.json',new Blob([JSON.stringify(manifest,null,2)],{type:'application/json'})],['notes.txt',new Blob([snapshot.groups.map(g=>`${g.title}\n工作日：${g.workDay??'舊版請參照各照片 workDay'}\n工程描述：${g.engineering}\n${g.notes?'舊版備註（保留）：'+g.notes+'\n':''}${g.legacyTitle?'舊版名稱（保留）：'+g.legacyTitle+'\n':''}順序：${g.photoIds.join(', ')}\n`).join('\n')])],['README.txt',new Blob(['原始影像在 originals/；檔名與 MIME 在 manifest.json。GPS 是裝置位置；方向不是可靠鏡頭朝向。相機讀值以按鈕時間為基準，不是曝光資料。舊版 capturedAt / ageAtShutterMs 亦是按鈕時間／年齡。沒有 EXIF 嵌入或雲端同步。'])]);const archive=await tar(entries),url=URL.createObjectURL(archive),a=document.createElement('a');a.href=url;a.download=`field-notebook-${Date.now()}.tar`;a.click();setTimeout(()=>URL.revokeObjectURL(url),60000);say('已交給瀏覽器下載；請確認下載完成（無法驗證使用者是否取消）');});
$('restore').onchange=()=>run(async()=>{const file=$('restore').files[0];if(!file)return;say('驗證備份中…');const entries=await untar(file),m=JSON.parse(await entries.get('manifest.json')?.text());if(m.schema!==1||!Array.isArray(m.photos)||!Array.isArray(m.groups))throw Error('不支援的備份');if(m.workDays!==undefined&&(!Array.isArray(m.workDays)||m.workDays.some(d=>!validDay(d))))throw Error('工作日格式無效');if(m.activeWorkDay!==undefined&&!validDay(m.activeWorkDay))throw Error('目前工作日格式無效');if(m.numberCounters!==undefined&&(typeof m.numberCounters!=='object'||m.numberCounters===null||Array.isArray(m.numberCounters)||Object.entries(m.numberCounters).some(([d,n])=>!validDay(d)||!Number.isSafeInteger(n)||n<1)))throw Error('編號計數格式無效');const incomingNumbers=new Set();const ids=new Set(),gids=new Set(),staged=[];for(const p of m.photos){if(p.workDay!==undefined&&!validDay(p.workDay)||p.photoNumber!==undefined&&(!Number.isSafeInteger(p.photoNumber)||p.photoNumber<1)||p.displayRotation!==undefined&&![0,90,180,270].includes(p.displayRotation))throw Error('照片工作日／編號／旋轉格式無效');if(p.workDay&&p.photoNumber){const numberKey=p.workDay+':'+p.photoNumber;if(incomingNumbers.has(numberKey)||photos.some(x=>x.workDay===p.workDay&&x.photoNumber===p.photoNumber))throw Error('工作日照片編號重複或與本機衝突');incomingNumbers.add(numberKey);}if(typeof p.id!=='string'||!/^[-a-zA-Z0-9]+$/.test(p.id)||ids.has(p.id)||photos.some(x=>x.id===p.id)||p.path!==`originals/${p.id}.bin`||typeof p.originalName!=='string'||!['camera','gallery'].includes(p.source)||typeof p.mime!=='string'||!p.mime.startsWith('image/'))throw Error('照片格式／ID 重複或與本機衝突');ids.add(p.id);const b=entries.get(p.path);if(!b||b.size!==p.size||await digest(b)!==p.sha256)throw Error('影像大小或雜湊不符');const {path,size,sha256,...photo}=p;staged.push({photo,blob:new Blob([b],{type:p.mime})});}for(const g of m.groups){if(typeof g.id!=='string'||typeof g.title!=='string'||typeof g.notes!=='string'||typeof g.engineering!=='string'||gids.has(g.id)||groups.some(x=>x.id===g.id))throw Error('紀錄集格式／ID 衝突');gids.add(g.id);validateGroup(g,ids);}if(!confirm(`已驗證 ${staged.length} 張照片、${m.groups.length} 組，匯入本機？`))return;await db.transaction(['photos','assets','groups'],tx=>{for(const x of staged){tx.objectStore('photos').add(x.photo);tx.objectStore('assets').add({id:x.photo.id,original:x.blob});}for(const g of m.groups)tx.objectStore('groups').add(g);});db.mergeCounters(m.numberCounters);for(const day of m.workDays??[])if(validDay(day))dayState.days.push(day);dayState.days=[...new Set(dayState.days)];activeDay=validDay(m.activeWorkDay)?m.activeWorkDay:null;if(activeDay)rememberDay(activeDay);await load();say('備份已完整還原（原圖與顯示旋轉分開保存）');$('restore').value='';});
async function changeDay(day){
  if(!validDay(day))throw Error('日期請使用有效 YYYY-MM-DD');
  if(busy||pending.length||rotating)throw Error('請先完成照片儲存或旋轉再切換工作日');
  if(!canDiscardEditor()){$('workDay').value=activeDay;return;}
  stop();

  // 1. Bind form saves to currently displayed day: flush pending edits before transition
  const prevDisplayed = monDisplayedDay;
  if(prevDisplayed){
    monDraftManager.saveDraft(prevDisplayed, readMonInputs(true));
  }

  // 2. Protect transition and handle rapid switches with generation token
  const currentSwitch = ++daySwitchGeneration;
  dayTransitioning = true;

  edit=null;
  editBaseline=null;
  $('editor').hidden=true;
  releaseImages('order');
  selected.clear();
  multi=false;
  if(active==='editor')tab('photos');

  rememberDay(day);

  try {
    await load();
  } finally {
    if(currentSwitch !== daySwitchGeneration){
      return;
    }
    dayTransitioning = false;
  }

  monDisplayedDay = day;
  if($('monForm')) $('monForm').dataset.displayedDay = day;
  updateMonUI();
  say(`已切換工作日 ${day}，其他日期照片仍保留`);
}
$('workDay').onchange=()=>run(async()=>{try{await changeDay($('workDay').value);}catch(e){$('workDay').value=activeDay;throw e;}});
$('newDay').onclick=()=>run(async()=>{const day=prompt('新工作日日期（YYYY-MM-DD）。只分開整理，不清除照片，也不改拍照時間。',localDay());if(day===null)return;await changeDay(day.trim());});
$('deleteSelected').onclick=()=>run(async()=>{if(edit&&(editorDirty()||groups.some(g=>g.id===edit.id))){say('請先儲存或關閉記事，再刪除照片。');return;}const ids=[...selected];if(!ids.length)return;if(groups.some(g=>g.photoIds.some(id=>ids.includes(id))))throw Error('已有記事的照片請先移出紀錄集再刪除');if(!confirm(`永久刪除 ${ids.length} 張照片及原圖？無法復原，建議先下載完整備份。`))return;await db.deletePhotos(ids);if(edit){edit=null;editBaseline=null;$('editor').hidden=true;releaseImages('order');}for(const id of ids){rotations.delete(id);for(const k of previewCache.keys())if(k.startsWith(id+':'))previewCache.delete(k);}selected.clear();await load();say(`已永久刪除 ${ids.length} 張照片；其他照片編號不變`);});
async function rotateSelected(delta){if(rotating||!selected.size)return;rotating=true;$('rotateLeft').disabled=$('rotateRight').disabled=true;say('正在調整顯示方向…');try{for(const id of [...selected]){const p=photos.find(p=>p.id===id);let r=p.displayRotation??rotations.get(id);if(r===undefined){const a=assets.find(a=>a.id===id);const url=URL.createObjectURL(a.original);try{const im=new Image();im.src=url;await im.decode();r=previewRotation(p,im.naturalWidth,im.naturalHeight);}finally{URL.revokeObjectURL(url);}}await db.setPhotoRotation(id,(r+delta+360)%360);for(const k of previewCache.keys())if(k.startsWith(id+':'))previewCache.delete(k);}await load();say('已儲存顯示方向，原始影像位元組未改動');}finally{rotating=false;$('rotateLeft').disabled=$('rotateRight').disabled=!selected.size;}}
$('rotateLeft').onclick=()=>run(()=>rotateSelected(-90));$('rotateRight').onclick=()=>run(()=>rotateSelected(90));
let exportSnapshot=null,exportDay=null,exportOrder=[],exportSelected=new Set(),exportBlob=null,exportWorking=false,exportGeneration=0,exportController=null,exportUrls=[];
function resetExportPreview(){exportBlob=null;exportUrls.forEach(URL.revokeObjectURL);exportUrls=[];$('exportPreview').replaceChildren();$('downloadWord').disabled=true;}
function exportControls(){const ids=exportOrder.filter(id=>exportSelected.has(id));for(const el of $('exportRows').querySelectorAll('input,button'))el.disabled=exportWorking||el.dataset.edge==='true';$('exportAll').disabled=$('exportNone').disabled=exportWorking;$('prepareWord').disabled=exportWorking||!ids.length;$('downloadWord').disabled=exportWorking||!exportBlob;$('exportSelection').textContent=`已選 ${ids.length} 組、${ids.reduce((n,id)=>n+(exportSnapshot?.groups.find(g=>g.id===id)?.photoIds.length??0),0)} 張照片`;} 
function renderExportChoices(focus){$('exportRows').replaceChildren();exportOrder.forEach((id,i)=>{const g=exportSnapshot.groups.find(g=>g.id===id),name=numberedGroup(g.photoIds,exportSnapshot.photos),row=text('li',''),label=text('label',''),check=document.createElement('input');check.type='checkbox';check.checked=exportSelected.has(id);check.dataset.groupId=id;check.setAttribute('aria-label',`選取 ${name}`);check.onchange=()=>{check.checked?exportSelected.add(id):exportSelected.delete(id);resetExportPreview();$('exportStatus').textContent='選取已改變，請重新製作預覽。';exportControls();};label.append(check,text('span',`${i+1}. ${name} · ${g.photoIds.length} 張`));row.append(label,text('p',g.engineering.length>180?g.engineering.slice(0,180)+'…':g.engineering||'工程描述未填寫'));const actions=text('div','');actions.className='actions';for(const [name,delta] of [['上移',-1],['下移',1]]){const b=text('button',name);b.dataset.groupId=id;b.dataset.action=name;b.dataset.edge=String(i+delta<0||i+delta>=exportOrder.length);b.setAttribute('aria-label',`第 ${i+1} 組 ${numberedGroup(g.photoIds,exportSnapshot.photos)} ${name}`);b.onclick=()=>{exportOrder=move(exportOrder,i,delta);resetExportPreview();$('exportStatus').textContent='匯出順序已改變，請重新製作預覽。';renderExportChoices({id,action:name});};actions.append(b);}row.append(actions);$('exportRows').append(row);});exportControls();if(focus){const buttons=[...$('exportRows').querySelectorAll('button')].filter(b=>b.dataset.groupId===focus.id&&!b.disabled);(buttons.find(b=>b.dataset.action===focus.action)||buttons[0]||$('prepareWord')).focus();}}
$('exportWord').onclick=()=>run(async()=>{if($('exportDialog').open)return;if(!validDay(activeDay)){say('工作日尚未載入，請稍候再開啟匯出。');return;}if(editorDirty()||pending.length||busy||rotating){say('請先儲存或關閉未完成的記事／照片，再匯出已儲存紀錄集。');return;}const day=activeDay,snapshot=await db.snapshot();if($('exportDialog').open||day!==activeDay||$('groups').hidden||editorDirty()||pending.length||busy||rotating)return;const byId=new Map(snapshot.photos.map(p=>[p.id,p]));const included=snapshot.groups.filter(g=>g.workDay===day||g.photoIds.some(id=>byId.has(id)&&workDayOf(byId.get(id))===day)).sort((a,b)=>(a.createdAt??a.updatedAt??0)-(b.createdAt??b.updatedAt??0)||a.id.localeCompare(b.id));if(!included.length){say('本工作日沒有已儲存紀錄集可匯出。');return;}exportGeneration++;exportSnapshot=snapshot;exportDay=day;exportOrder=included.map(g=>g.id);exportSelected=new Set(exportOrder);exportWorking=false;resetExportPreview();$('exportStatus').textContent='請確認選取與順序，再按「預覽並製作表格」。';$('exportSnapshotInfo').textContent=`工作日 ${day} · 已儲存快照 ${new Date(snapshot.completedAt).toLocaleString()}；不修改原圖或記事。`;renderExportChoices();$('exportDialog').showModal();});
$('exportAll').onclick=()=>{exportSelected=new Set(exportOrder);resetExportPreview();renderExportChoices();};$('exportNone').onclick=()=>{exportSelected.clear();resetExportPreview();renderExportChoices();};
$('closeExport').onclick=()=>$('exportDialog').close();$('exportDialog').addEventListener('close',()=>{exportGeneration++;exportController?.abort();exportController=null;exportWorking=false;resetExportPreview();exportSnapshot=null;exportOrder=[];exportSelected.clear();$('exportRows').replaceChildren();});
$('prepareWord').onclick=async()=>{if(exportWorking||!$('exportDialog').open||!exportSnapshot)return;const token=++exportGeneration;exportController?.abort();exportController=new AbortController();const signal=exportController.signal,ids=exportOrder.filter(id=>exportSelected.has(id));exportWorking=true;resetExportPreview();exportControls();try{const prepared=await preparePhotoGroups(exportSnapshot,ids,{signal,onProgress:(n,total)=>{if(token===exportGeneration)$('exportStatus').textContent=`正在本機整理照片 ${n}/${total}，不會上傳…`;}});if(token!==exportGeneration||!$('exportDialog').open)return;const blob=await photoTableDocx(prepared);if(token!==exportGeneration||!$('exportDialog').open)return;const table=document.createElement('table');table.className='docx-preview-table';const body=document.createElement('tbody');for(const row of tablePlan(prepared)){const tr=document.createElement('tr');if(row.kind==='description'){const td=document.createElement('td');td.colSpan=2;const p=text('p',row.engineering);p.className='note-text';td.append(p);tr.append(td);}else for(let i=0;i<2;i++){const td=document.createElement('td'),photo=row.photos[i];if(photo){const im=document.createElement('img'),url=URL.createObjectURL(photo.blob);exportUrls.push(url);im.src=url;im.alt=photoLabel(exportSnapshot.photos.find(p=>p.id===photo.photoId));td.append(im);}else{const blank=text('div','');blank.className='empty-export-photo';blank.setAttribute('aria-label','本列無第二張照片');td.append(blank);}tr.append(td);}body.append(tr);}table.append(body);$('exportPreview').append(text('p','以下為連續表格預覽；Word 會依頁面長度自動跨頁。'),table);exportBlob=blob;const low=prepared.flatMap(g=>g.photos).filter(p=>Math.max(p.width,p.height)<800).length;$('exportStatus').textContent=`表格已製作完成，${prepared.length} 組／${prepared.reduce((n,g)=>n+g.photos.length,0)} 張；可下載可編輯 Word。${low?'其中 '+low+' 張解析度較低，列印放大可能模糊。':''}`;}catch(e){if(token===exportGeneration&&e.name!=='AbortError'){resetExportPreview();$('exportStatus').textContent='匯出未完成：'+e.message+'；原始資料未改動。';}}finally{if(token===exportGeneration){exportWorking=false;exportControls();}}};
$('downloadWord').onclick=()=>{if(!exportBlob||exportWorking)return;const url=URL.createObjectURL(exportBlob),a=document.createElement('a');a.href=url;a.download=`照片表格-${exportDay}.docx`;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);$('exportStatus').textContent='已交給瀏覽器下載 Word；請確認檔案存在並用 Word／相容軟體開啟。這不是完整原圖／中繼資料備份。';};
window.addEventListener('pagehide',()=>exportController?.abort());
const monDraftManager = new MonitoringDraftManager(safeGetLocalStorage());
function readMonInputs(preserveWhitespace = true){
 return Object.fromEntries(Object.entries(MON_FIELDS).map(([key, id]) => [key, preserveWhitespace ? $(id).value : $(id).value.trim()]));
}
function fillMonInputs(data){
 for(const [key, id] of Object.entries(MON_FIELDS)) $(id).value = data?.[key] !== undefined && data?.[key] !== null ? data[key] : '';
}
function loadMonForm(day){
 return monDraftManager.getDraft(day || displayedDay());
}
function saveMonForm(day, data){
 return monDraftManager.saveDraft(day || displayedDay(), data);
}
function updateMonStatusUI(){
 const d = displayedDay();
 const st = monDraftManager.getStatus(d);
 $('monStatus').textContent = st ? st.status : '';
 const isUnsaved = monDraftManager.hasUnsaved(d);
 const isCorrupt = monDraftManager.hasCorrupt(d);
 const isConflict = monDraftManager.hasConflict(d);
 const isReadBlocked = monDraftManager.isReadBlocked(d);
 $('monRetrySave').hidden = !isUnsaved || isCorrupt;
 $('monExportDraft').hidden = !isUnsaved && !isCorrupt && !isConflict && !isReadBlocked;
 $('monOverwriteCorrupt').hidden = !isCorrupt;
}
function syncStrataBoth(notifyIfNoChange = false){
 const d = displayedDay();
 const scope = $('monWorkScope').value;
 const summary = $('monStrataSummary').value;
 const desc = $('monStrata').value;
 const res = synchronizeBothStrata(scope, summary, desc);

 if(res.warnings.length > 0){
  $('monSyncNotice').textContent = res.warnings.join('；');
  $('monSyncNotice').hidden = false;
 }else{
  $('monSyncNotice').textContent = '';
  $('monSyncNotice').hidden = true;
 }

 if(res.changed){
  $('monStrataSummary').value = res.strataSummary;
  $('monStrata').value = res.strataDesc;
  if(!dayTransitioning){
   monDraftManager.saveDraft(d, readMonInputs(true));
   updateMonStatusUI();
  }
  if(notifyIfNoChange) say('已同步地層欄位編號。');
 }else if(notifyIfNoChange){
  say('地層欄位編號與作業項目已一致，無須變更。');
 }
}

function updateMonUI(){
 const d = displayedDay();
 fillMonInputs(null);
 const data = monDraftManager.getDraft(d);
 if(data){
  fillMonInputs(data);
 }else{
  $('monDate').value = formatMonDate(d);
 }
 updateMonStatusUI();
 const scope = $('monWorkScope').value;
 const summary = $('monStrataSummary').value;
 const desc = $('monStrata').value;
 const check = synchronizeBothStrata(scope, summary, desc);
 if(check.warnings.length > 0){
  $('monSyncNotice').textContent = check.warnings.join('；');
  $('monSyncNotice').hidden = false;
 }else{
  $('monSyncNotice').textContent = '';
  $('monSyncNotice').hidden = true;
 }
}
$('monForm').addEventListener('input', (e) => {
 const d = displayedDay();
 if(dayTransitioning || !d) return;
 monDraftManager.saveDraft(d, readMonInputs(true));
 updateMonStatusUI();
 if(e.target === $('monWorkScope')){
  const parsed = parseWorkScope($('monWorkScope').value);
  if(parsed.warnings.length > 0){
   $('monSyncNotice').textContent = parsed.warnings.join('；');
   $('monSyncNotice').hidden = false;
  }else{
   $('monSyncNotice').textContent = '';
   $('monSyncNotice').hidden = true;
  }
 }
});
let isTransientOffer = false;

$('monWorkScope').addEventListener('focus', () => {
 if($('monWorkScope').value === ''){
  isTransientOffer = true;
  $('monWorkScope').value = '1. ';
  $('monWorkScope').setSelectionRange(3, 3);
 }
});
$('monWorkScope').addEventListener('input', () => {
 isTransientOffer = false;
});
$('monWorkScope').addEventListener('blur', () => {
 if(isTransientOffer && $('monWorkScope').value === '1. '){
  isTransientOffer = false;
  $('monWorkScope').value = '';
  return;
 }
 isTransientOffer = false;
 syncStrataBoth(false);
});
$('monWorkScope').addEventListener('change', () => {
 syncStrataBoth(false);
});
$('monWorkScope').addEventListener('keydown', (e) => {
 handleWorkScopeKeyDown(e, $('monWorkScope'));
});
$('addMainItem').onclick = () => {
 const ta = $('monWorkScope');
 ta.focus();
 if(isTransientOffer && ta.value === '1. '){
  isTransientOffer = false;
  ta.setSelectionRange(3, 3);
  ta.dispatchEvent(new Event('input', {bubbles: true}));
  return;
 }
 isTransientOffer = false;
 const val = ta.value;

 if(!val.trim() || val.trim() === '1.'){
  ta.value = '';
  insertTextAtCaret(ta, '1. ');
  ta.dispatchEvent(new Event('input', {bubbles: true}));
  return;
 }
 const nextMain = getNextMainNumber(val);
 const selStart = ta.selectionStart ?? val.length;
 const selEnd = ta.selectionEnd ?? val.length;
 const before = val.slice(0, selStart);
 const after = val.slice(selEnd);

 const needPrefix = before.length > 0 && !before.endsWith('\n');
 const prefix = needPrefix ? '\n' : '';
 const needSuffix = after.length > 0 && !after.startsWith('\n');
 const suffix = needSuffix ? '\n' : '';

 const itemText = `${prefix}${nextMain}. `;
 const insertText = `${itemText}${suffix}`;
 insertTextAtCaret(ta, insertText);

 if(suffix.length > 0){
  const newCaret = selStart + itemText.length;
  ta.setSelectionRange(newCaret, newCaret);
 }
 ta.dispatchEvent(new Event('input', {bubbles: true}));
 syncStrataBoth(false);
};
$('addChildItem').onclick = () => {
 const ta = $('monWorkScope');
 const val = ta.value;
 ta.focus();

 const selStart = ta.selectionStart ?? val.length;
 const selEnd = ta.selectionEnd ?? val.length;

 if(!val.trim()){
  insertTextAtCaret(ta, '1.(1) ');
  ta.dispatchEvent(new Event('input', {bubbles: true}));
  return;
 }
 const nextChild = getNextChildNumber(val, selStart);
 const before = val.slice(0, selStart);
 const after = val.slice(selEnd);

 const needPrefix = before.length > 0 && !before.endsWith('\n');
 const prefix = needPrefix ? '\n' : '';
 const needSuffix = after.length > 0 && !after.startsWith('\n');
 const suffix = needSuffix ? '\n' : '';

 const itemText = `${prefix}(${nextChild}) `;
 const insertText = `${itemText}${suffix}`;
 insertTextAtCaret(ta, insertText);

 if(suffix.length > 0){
  const newCaret = selStart + itemText.length;
  ta.setSelectionRange(newCaret, newCaret);
 }
 ta.dispatchEvent(new Event('input', {bubbles: true}));
 syncStrataBoth(false);
};
$('syncStrata').onclick = () => {
 syncStrataBoth(true);
};
$('monRetrySave').onclick = () => {
 const d = displayedDay();
 if(monDraftManager.isReadBlocked(d)){
  monDraftManager.retryRead(d);
 }else if(monDraftManager.hasConflict(d)){
  monDraftManager.forceOverwriteConflict(d);
 }else{
  monDraftManager.saveDraft(d, readMonInputs(true));
 }
 updateMonStatusUI();
};
$('monExportDraft').onclick = () => {
 const d = displayedDay();
 const exp = monDraftManager.exportCopy(d);
 if(!exp){
  say('無可匯出之草稿文字');
  return;
 }
 const textContent = exp.combinedText || (exp.type === 'corrupt_raw' ? exp.raw : (exp.json || JSON.stringify(exp.data, null, 2)));
 const blob = new Blob([textContent], {type: 'text/plain;charset=utf-8'});
 const url = URL.createObjectURL(blob), a = document.createElement('a');
 a.href = url;
 a.download = `監看草稿備份-${d}.txt`;
 document.body.append(a);
 a.click();
 a.remove();
 setTimeout(() => URL.revokeObjectURL(url), 60000);
 say('已下載草稿文字複本（含原始異常字串及記憶體變更），請確認存檔。');
};
$('monOverwriteCorrupt').onclick = () => {
 const d = displayedDay();
 if(!confirm('確認要覆寫異常草稿？此動作無法復原。')){
  return;
 }
 monDraftManager.forceOverwriteCorrupt(d);
 updateMonStatusUI();
};
$('fillMonTemplate').onclick = () => {
 const d = displayedDay();
 const currentInputs = readMonInputs(true);
 const tpl = genericMonTemplate(d);
 const res = monDraftManager.applyTemplate(d, tpl, currentInputs);
 fillMonInputs(res.filled);
 syncStrataBoth(false);
 monDraftManager.setStatus(d, {
  status: res.statusMsg,
  isError: !res.saveResult.success,
  isUnsaved: !res.saveResult.success
 });
 updateMonStatusUI();
};
// The offline download and Drive export use the same committed snapshot preparation.
async function prepareMonitoringDay(){
 const day=displayedDay(),formData=readMonInputs(),snapshot=await db.snapshot(),byId=new Map(snapshot.photos.map(p=>[p.id,p]));
 const included=snapshot.groups.filter(g=>g.workDay===day||g.photoIds.some(id=>byId.has(id)&&workDayOf(byId.get(id))===day)).sort((a,b)=>(a.createdAt??a.updatedAt??0)-(b.createdAt??b.updatedAt??0)||a.id.localeCompare(b.id));
 const prepared=included.length?await preparePhotoGroups(snapshot,included.map(g=>g.id)):[];
 return {day,docBlob:await monitoringDocx(formData,prepared)};
}
let monitoringWorking=false;
function monitoringControls(){ $('exportMonDocx').disabled=$('uploadDrive').disabled=monitoringWorking; }
$('exportMonDocx').onclick=async()=>{
 if(monitoringWorking)return;monitoringWorking=true;monitoringControls();
 try{
  $('monStatus').textContent='正在處理照片並生成完整 Word 監看紀錄…';
  const {day,docBlob}=await prepareMonitoringDay();
  const url=URL.createObjectURL(docBlob),a=document.createElement('a');
  a.href=url;a.download=`${loadMonForm(day)?.date||day}_施工監看紀錄.docx`;
  document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);
  $('monStatus').textContent='已交給瀏覽器下載完整監看紀錄 Word (.docx)；請確認下載完成。';
 }catch(e){$('monStatus').textContent='匯出失敗：'+e.message;}
 finally{monitoringWorking=false;monitoringControls();}
};
$('connectDrive').onclick=async()=>{
 $('connectDrive').disabled=true;
 try{await authorize();$('driveAccount').textContent='已取得 Google Drive 授權（僅本頁有效；過期時須重新授權）。';}
 catch(e){$('driveAccount').textContent='授權失敗：'+e.message;}
 finally{$('connectDrive').disabled=false;}
};

let uploadingPhotos=false;
const PHOTO_FOLDER_KEY='field-photo-drive-folder-target';
const savedFolder=safeGetLocalStorage()?.getItem(PHOTO_FOLDER_KEY);
if(savedFolder&&$('photoDriveFolder'))$('photoDriveFolder').value=savedFolder;

$('photoDriveFolder')?.addEventListener('change',()=>{
 try{safeGetLocalStorage()?.setItem(PHOTO_FOLDER_KEY,$('photoDriveFolder').value.trim());}catch{}
});

// Folder browser interactive logic
let currentBrowsingFolderId=null;
async function refreshFolderList(parentId){
 const listEl=$('folderList'),statusEl=$('folderBrowserStatus');
 if(!listEl)return;
 listEl.replaceChildren();
 if(statusEl)statusEl.textContent='正在讀取資料夾清單…';
 try{
  const folders=await listSubfolders(parentId);
  if(!folders.length){
   if(statusEl)statusEl.textContent='此資料夾內目前沒有子資料夾。';
   return;
  }
  for(const f of folders){
   const li=document.createElement('li');
   li.style.cssText='display:flex;justify-content:space-between;align-items:center;padding:8px 12px;border-bottom:1px solid #eef1eb';
   const nameSpan=document.createElement('span');
   nameSpan.textContent=f.name;
   nameSpan.style.cssText='font-weight:500;overflow-wrap:anywhere';
   const btn=document.createElement('button');
   btn.type='button';
   btn.textContent='選取此資料夾';
   btn.style.cssText='min-height:36px;padding:4px 8px;font-size:0.85rem';
   btn.onclick=()=>{
    $('photoDriveFolder').value=f.id;
    try{safeGetLocalStorage()?.setItem(PHOTO_FOLDER_KEY,f.id);}catch{}
    $('folderBrowserDialog').close();
    if($('photoDriveStatus'))$('photoDriveStatus').textContent=`已選擇目標資料夾：${f.name}`;
   };
   li.append(nameSpan,btn);
   listEl.append(li);
  }
  if(statusEl)statusEl.textContent=`共找到 ${folders.length} 個子資料夾。點選「選取此資料夾」以設定。`;
 }catch(e){
  if(statusEl)statusEl.textContent='讀取失敗：'+e.message;
 }
}

$('browseDriveFolders')?.addEventListener('click',async()=>{
 const inputVal=$('photoDriveFolder').value.trim();
 let folderId=null;
 try{folderId=parseDriveId(inputVal);}catch(e){
  if($('photoDriveStatus'))$('photoDriveStatus').textContent=e.message;
  return;
 }
 currentBrowsingFolderId=folderId;
 $('folderBrowserDialog')?.showModal();
 const currentNameEl=$('folderBrowserCurrent');
 if(currentNameEl)currentNameEl.textContent=folderId?`母資料夾 ID：${folderId}`:'根目錄';
 if($('newSubfolderName'))$('newSubfolderName').value=activeDay||'';
 try{
  await authorize();
  await refreshFolderList(currentBrowsingFolderId);
 }catch(e){
  if($('folderBrowserStatus'))$('folderBrowserStatus').textContent='授權失敗：'+e.message;
 }
});

$('closeFolderBrowser')?.addEventListener('click',()=>{
 $('folderBrowserDialog')?.close();
});

$('createNewSubfolder')?.addEventListener('click',async()=>{
 const nameInput=$('newSubfolderName');
 const name=nameInput?.value.trim();
 if(!name){if($('folderBrowserStatus'))$('folderBrowserStatus').textContent='請輸入新資料夾名稱';return;}
 const btn=$('createNewSubfolder');
 btn.disabled=true;
 try{
  if($('folderBrowserStatus'))$('folderBrowserStatus').textContent=`正在建立「${name}」…`;
  const created=await createFolder(name,currentBrowsingFolderId);
  if($('folderBrowserStatus'))$('folderBrowserStatus').textContent=`已建立資料夾：${created.name}！已自動為您選取。`;
  $('photoDriveFolder').value=created.id;
  try{safeGetLocalStorage()?.setItem(PHOTO_FOLDER_KEY,created.id);}catch{}
  await refreshFolderList(currentBrowsingFolderId);
 }catch(e){
  if($('folderBrowserStatus'))$('folderBrowserStatus').textContent='建立失敗：'+e.message;
 }finally{
  btn.disabled=false;
 }
});

$('uploadPhotosToDrive')?.addEventListener('click',async()=>{
 if(uploadingPhotos)return;
 const statusEl=$('photoDriveStatus');
 let folderId=null;
 try{folderId=parseDriveId($('photoDriveFolder').value);}catch(e){if(statusEl)statusEl.textContent=e.message;return;}
 const targets=dayPhotos().filter(p=>p.source==='camera');
 const uploadList=selected.size?targets.filter(p=>selected.has(p.id)):targets;
 if(!uploadList.length){if(statusEl)statusEl.textContent='目前工作日沒有可上傳的相機照片。';return;}
 uploadingPhotos=true;$('uploadPhotosToDrive').disabled=true;
 try{
  await authorize();
  let targetFolderId=folderId;
  const autoDaily=$('autoDailySubfolder')?.checked;
  if(autoDaily&&targetFolderId){
   if(statusEl)statusEl.textContent=`正在檢查/建立工作日 (${activeDay}) 子資料夾…`;
   const dailyFolder=await ensureDailyFolder(targetFolderId,activeDay);
   targetFolderId=dailyFolder.id;
  }
  let successCount=0;
  for(let i=0;i<uploadList.length;i++){
   const p=uploadList[i];
   if(statusEl)statusEl.textContent=`正在上傳照片 (${i+1}/${uploadList.length})：${photoLabel(p)}…`;
   const asset=assets.find(a=>a.id===p.id);
   if(!asset||!asset.original)continue;
   await uploadOriginalPhoto(p,asset.original,targetFolderId);
   successCount++;
  }
  if(statusEl)statusEl.textContent=`成功上傳 ${successCount} 張原始照片至 Google Drive！本機照片仍保留。`;
 }catch(e){
  if(statusEl)statusEl.textContent=`上傳失敗：${e.message}。已成功項目不受影響。`;
 }finally{
  uploadingPhotos=false;$('uploadPhotosToDrive').disabled=false;
 }
});

$('uploadDrive').onclick=async()=>{
 if(monitoringWorking)return;
 let folderId;
 try{folderId=parseDriveId($('driveFolder').value);}catch(e){$('driveStatus').textContent=e.message;return;}
 monitoringWorking=true;monitoringControls();$('driveResult').replaceChildren();
 try{
  // Request consent in direct response to the user's tap, before expensive async generation.
  await authorize();$('driveAccount').textContent='已取得 Google Drive 授權（僅本頁有效）。';
  $('driveStatus').textContent='正在本機生成當日文件並上傳；請勿重複點擊…';
  const {day,docBlob}=await prepareMonitoringDay();
  const file=await uploadDayDoc(day,docBlob,folderId);
  const link=document.createElement('a');link.href=file.webViewLink;link.target='_blank';link.rel='noopener noreferrer';link.textContent='Open Google Doc（開啟 Google 文件）';
  $('driveResult').append(link);
  $('driveStatus').textContent=`已上傳 ${file.name||day} 並轉為 Google Doc。轉換可能改變版面，請開啟檢查。此操作不會同步本機資料。`;
 }catch(e){$('driveStatus').textContent=`${e.kind==='auth'?'授權':e.kind==='network'?'網路':e.kind==='api'?'Drive API':'匯出'}失敗：${e.message}`;if(!hasToken())$('driveAccount').textContent='尚未授權或授權已過期；請重新登入。';}
 finally{monitoringWorking=false;monitoringControls();}
};
$('persist').onclick=()=>run(async()=>{const ok=await navigator.storage?.persist?.();say(ok?'瀏覽器允許持續儲存，仍須備份':'未取得持續儲存保證，請備份');});
run(async()=>{await load();const s=await navigator.storage?.estimate?.();$('storage').textContent=s?`使用約 ${(s.usage/1048576).toFixed(1)} MB / 配額約 ${(s.quota/1048576).toFixed(0)} MB（可能變動）`:'無法取得配額';});
window.addEventListener('storage',e=>{
  if(e.key&&e.key.startsWith('field-mon-form-')){
    const day=e.key.replace('field-mon-form-','');
    if(displayedDay()===day){
      if(!monDraftManager.hasUnsaved(day)){
        monDraftManager.clearCache(day);
        updateMonUI();
      }else{
        monDraftManager.checkExternalChange(day);
        updateMonStatusUI();
      }
    }
  }
});
if('serviceWorker' in navigator&&isSecureContext)run(async()=>{await navigator.serviceWorker.register('./sw.js',{scope:'./'});await navigator.serviceWorker.ready;const update=()=>{$('offline').textContent=navigator.serviceWorker.controller?'離線殼已受控制；照片只存 IndexedDB，仍需備份。':'離線殼已安裝，重新載入後確認控制。';};update();navigator.serviceWorker.addEventListener('controllerchange',update);});
