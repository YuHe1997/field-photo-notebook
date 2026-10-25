import * as db from './db.js';
import {compareGroups} from './ordering.js';
import * as docsAppendInfo from './docs-append.js';
import {authorize,hasToken,parseDriveId,uploadDayDoc,validateGoogleDocName,uploadOriginalPhoto,listSubfolders,listMonitoringDocuments,listSharedDrives,verifyFolder,createFolder,ensureDailyFolder,inspectMaster,appendMonitoringRecord,rocDayName} from './drive.js';
import {preparePhotoGroups,photoTableDocx,tablePlan,monitoringDocx,monitoringRowLayout} from './docx.js';
import {
  MON_FIELDS,
  monKey,
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
import {validCrop,normalizeCrop,rotateCrop,cropSummary,compactCoordinates,compactPhotoLines,directionLabel,readingAt,validateGroup,move,digest,tar,untar,localDay,validDay,workDayOf,numberedPhoto,numberedGroup,cameraAutoRotation,previewRotation,zoomCapability,snapZoom,cropRect,focusCapabilities,mapPreviewTapToSensor,deriveCameraBearing} from './core.js';
import {WorkPhrasePicker,readPhraseMetadata,photoEngineeringPhrases,readSavedScopePhrases,phraseInsertion} from './work-phrases.js';
import {renderPhotoCanvas} from './photo-render.js';
import {openCropDialog} from './crop-ui.js';
import {AREA_DATA} from './area-data.js';
import {createAreaMatcher,areaReferenceText} from './area-reference.js';
let areaMatcher;try{areaMatcher=createAreaMatcher(AREA_DATA);}catch{areaMatcher=()=>({status:'unavailable',reason:'dataset-invalid',candidates:[],datasetVersion:AREA_DATA?.version});}
const areaReference=p=>{try{return areaMatcher(p);}catch{return {status:'unavailable',reason:'dataset-invalid',candidates:[]};}};
const $=id=>document.getElementById(id);
const preventCameraPageZoom=e=>{if(document.body.classList.contains('camera-overlay-active'))e.preventDefault();};
for(const type of ['gesturestart','gesturechange','gestureend'])document.addEventListener(type,preventCameraPageZoom,{passive:false});
let overlayErrorText=null;
const say=s=>{
  $('status').textContent=s;
  $('status').hidden=!s;
  if($('cameraOverlay')&&!$('cameraOverlay').hidden){
    overlayErrorText=s;
    if($('cameraOverlayStatus')){$('cameraOverlayStatus').textContent=s;$('cameraOverlayStatus').classList.add('is-notice');}
  }
};
let photos=[],assets=[],groups=[],selected=new Set(),multi=false,edit=null,editBaseline=null,urls={gallery:[],groups:[],order:[],reference:[]},stream=null,watch=null,gps=null,heading=null,deviceTilt=null,sensorState='尚未啟用',active='camera',starting=false,busy=false,rotating=false,cameraSession=0,zoomState=null,focusState=null,lastFocusedElement=null,directionStatus='尚未啟用',gpsStatus='尚未啟用',activeSessionQueue=Promise.resolve(),constraintApplying=false,commitPromise=null,activeStartupSession=0,sessionReady=false;
const text=(tag,s)=>{const e=document.createElement(tag);e.textContent=s;return e;};
async function run(fn){try{await fn();}catch(e){say('操作／儲存失敗：'+e.message+'。資料未宣稱成功，請重試或先備份。');}}
const DAY_KEY='field-photo-notebook-work-days';
function readDays(){try{const s=safeGetLocalStorage();const d=JSON.parse(s?.getItem(DAY_KEY)||'{}');return {active:validDay(d.active)?d.active:null,days:Array.isArray(d.days)?d.days.filter(validDay):[]};}catch{return {active:null,days:[]};}}
let dayState=readDays(),activeDay=dayState.active;
let monDisplayedDay=null,daySwitchGeneration=0,dayTransitioning=false,loadGeneration=0;
let photoPhrasePicker=null,scopePhrasePicker=null;
let coordinateAlbum=null,coordinateAlbumGeneration=0,coordinateReturn=null;
function closePhrasePickers(){photoPhrasePicker?.invalidate();scopePhrasePicker?.invalidate();}
function displayedDay(){return monDisplayedDay||$('monForm')?.dataset?.displayedDay||activeDay;}
const previewCache=new Map(),rotations=new Map();
function rememberDay(day){const next={active:day,days:[...new Set([...dayState.days,day])].sort()};try{safeGetLocalStorage()?.setItem(DAY_KEY,JSON.stringify(next));}catch{}dayState=next;activeDay=day;}
async function load(){const generation=++loadGeneration;await db.ensurePhotoIdentities();const snapshot=await db.snapshot();if(generation!==loadGeneration)return;({photos,assets,groups}=snapshot);groups.sort(compareGroups);photos.sort((a,b)=>photoTime(a)-photoTime(b)||(a.photoNumber??0)-(b.photoNumber??0)||a.id.localeCompare(b.id));if(!activeDay)rememberDay(photos.length?workDayOf(photos.at(-1)):localDay());if(!monDisplayedDay)monDisplayedDay=activeDay;if($('monForm')&&monDisplayedDay)$('monForm').dataset.displayedDay=monDisplayedDay;render();}
const photoTime=p=>p.buttonPressedAt??p.capturedAt??p.importedAt??0;
const photoLabel=p=>p?numberedPhoto(p):'照片未取得';
const groupTitle=ids=>numberedGroup(ids,photos);
const dayPhotos=()=>photos.filter(p=>workDayOf(p)===activeDay);
const groupInDay=g=>g.workDay===activeDay||g.photoIds.some(id=>photos.some(p=>p.id===id&&workDayOf(p)===activeDay));
function releaseImages(pool){urls[pool].forEach(URL.revokeObjectURL);urls[pool]=[];}
function photoImage(id,pool,source={assets,photos}){const a=source.assets.find(a=>a.id===id),p=source.photos.find(p=>p.id===id);if(!a)return text('p','原始影像未取得');const im=document.createElement('img');im.loading='lazy';im.alt=photoLabel(p);const key=`${id}:${p.displayRotation??'auto'}:${validCrop(p.crop)?[p.crop.x,p.crop.y,p.crop.w].join(','):'full'}`;if(previewCache.has(key)){im.src=previewCache.get(key);im.dataset.previewReady='true';im.dataset.rotation=p.displayRotation??rotations.get(id)??0;return im;}im.onerror=()=>{im.dataset.previewReady='true';im.dataset.previewError='true';im.alt=photoLabel(p)+'（原圖無法解碼，未刪除）';};im.src=URL.createObjectURL(a.original);urls[pool].push(im.src);im.onload=()=>{im.onload=null;try{const {canvas:c,rotation:r}=renderPhotoCanvas(im,p,{maxSide:960});rotations.set(id,r);const result=c.toDataURL('image/jpeg',.9);if(previewCache.size>=60)previewCache.delete(previewCache.keys().next().value);previewCache.set(key,result);if(im.isConnected)im.src=result;im.dataset.rotation=r;im.dataset.previewReady='true';}catch{im.dataset.previewReady='true';say('此照片無法產生旋轉預覽；原圖仍保留。');}};return im;}
function photoAreaInfo(p,{collapsed=false}={}){
  const result=areaReference(p),box=text('div','');box.className='photo-area-reference';
  const reference=text('p',areaReferenceText(result));
  if(collapsed){const fold=text('details','');fold.className='area-summary-details';fold.append(text('summary','區域參考'),reference);box.append(fold);}else box.append(reference);
  const details=text('details','');details.append(text('summary','區域判讀依據（暫定）'));
  details.append(text('p',`區界版本：${result.datasetVersion||'未取得'}。依照片已保存定位及目前版本回算，換版後提示可能改變；不改原始 GPS。`));
  if(result.accuracyM!==null&&Number.isFinite(result.accuracyM))details.append(text('p',`GPS 回報精度：約 ${Math.ceil(result.accuracyM)} m；暫採圖界緩衝 ${result.mapBufferM} m。兩者均非總誤差保證。`));
  if(result.candidates?.length)details.append(text('p',result.candidates.map(c=>`${c.name}：${c.inside?'座標在標註內':`距標註邊緣約 ${Math.ceil(c.distanceM)} m`}`).join('；')));
  details.append(text('p','圖名、區界與套圖位置仍待確認。GPS 是拍攝者位置，不代表照片中的工程或被攝物位置；不自動填寫工程描述。'));
  box.append(details);return box;
}
function photoInfo(p){
  const box=text('div','');
  box.className='photo-info';
  for(const line of compactPhotoLines(p,0).slice(1))box.append(text('p',line));
  box.append(photoAreaInfo(p,{collapsed:true}));
  if(p?.heading?.degrees!==undefined||p?.cameraBearing!==undefined||p?.cameraBearingEvidence){
    const details=text('details','');
    details.className='camera-direction-details';
    details.append(text('summary','裝置羅盤參考'));
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
  watch=null;gps=null;heading=null;deviceTilt=null;
  zoomState?.cancelFrame?.();zoomState=null;
  focusState=null;
  if($('focusReticle'))$('focusReticle').hidden=true;
  if($('focusStatus'))$('focusStatus').textContent='';
  if($('cameraOverlayStatus'))$('cameraOverlayStatus').textContent='';
  if($('cameraOverlay'))$('cameraOverlay').hidden=true;
  leaveCameraFullscreen();lastLayoutKey='';previewRot=0;
  document.body.classList.remove('camera-overlay-active');
  directionStatus=gpsStatus='已停止';
  window.removeEventListener('deviceorientation',onOrientation);
  window.removeEventListener('deviceorientationabsolute',onOrientation);
  sensorState='感測已停止';
  zoomUI();sensors();
}
async function closeCamera(){
  // Leave full screen first: the browser will not move focus to the page while the overlay is still the full-screen element.
  await exitFullscreenAndWait();
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
if(/iPhone|iPod/.test(navigator.userAgent)&&!standaloneApp()&&$('fullscreenHint'))$('fullscreenHint').hidden=false;
function tab(name){if(name==='editor'&&!edit)return;closeCoordinateAlbum(false);closePhrasePickers();if(name!=='camera')stop();active=name;for(const n of ['camera','photos','groups','monitoring'])$(n).hidden=n!==name;$('editor').hidden=name!=='editor'||!edit;$('editingTab').hidden=true;$('appendToEditor').hidden=$('returnToEditor').hidden=!edit;$('compose').hidden=!!edit;document.querySelectorAll('[data-tab]').forEach(b=>b.setAttribute('aria-pressed',b.dataset.tab===name));if(name==='editor'){renderOrder();$('editor').scrollIntoView({block:'start'});const h=$('editor').querySelector('h2');h.tabIndex=-1;h.focus({preventScroll:true});}if(name==='monitoring')updateMonUI();}
window.tab=tab;
document.querySelectorAll('[data-tab]').forEach(b=>{b.onclick=()=>tab(b.dataset.tab);b.setAttribute('aria-pressed',String(b.dataset.tab===active));});
function currentCamera(session){return session===cameraSession&&active==='camera'&&!document.hidden;}
function beginDirection(session){const E=window.DeviceOrientationEvent;if(!E){directionStatus='不支援方向感測';sensors();return;}directionStatus='等待方向授權';let permission;try{permission=typeof E.requestPermission==='function'?E.requestPermission(true):Promise.resolve('granted');}catch(e){permission=Promise.reject(e);}Promise.resolve(permission).then(result=>{if(!currentCamera(session))return;if(result!=='granted'){directionStatus='方向權限被拒絕';sensors();return;}window.addEventListener('deviceorientation',onOrientation);window.addEventListener('deviceorientationabsolute',onOrientation);directionStatus='已啟用，等待絕對方位讀值';sensors();},e=>{if(currentCamera(session)){directionStatus='方向授權失敗：'+e.message;sensors();}});}
function setCameraDetails(open){
  $('cameraDetails').hidden=!open;
  $('cameraInfo').setAttribute('aria-expanded',String(open));
}
$('cameraInfo').onclick=()=>setCameraDetails($('cameraDetails').hidden);
$('cameraDetailsClose').onclick=()=>{setCameraDetails(false);$('cameraInfo').focus({preventScroll:true});};
// The picture the user sees must match the saved photo: turn the preview by the same gravity-based rotation that the
// capture applies, and keep re-measuring so a rotation of the phone (or the browser bars) can never leave it stretched.
let previewRot=0,lastLayoutKey='';
function desiredPreviewRotation(){
  const v=$('video');if(!v?.videoWidth||!v.videoHeight)return 0;
  const tilt=deviceTilt&&Date.now()-deviceTilt.readingAt<=2000?deviceTilt:null;
  if(!tilt)return 0;
  const b=tilt.beta*Math.PI/180,g=tilt.gamma*Math.PI/180;
  // Phone lying flat: gravity says nothing about "up", keep what we had instead of flipping.
  if(Math.hypot(-Math.cos(b)*Math.sin(g),Math.sin(b))<0.25)return previewRot;
  return cameraAutoRotation({width:v.videoWidth,height:v.videoHeight,screenAngle:tilt.screenAngle??screen.orientation?.angle??window.orientation??null,beta:tilt.beta,gamma:tilt.gamma});
}
function previewDimensions(){
  const v=$('video'),preview=$('cameraPreview'),vp=$('cameraViewport'),rotor=$('cameraRotor');
  if(!v||!preview||!vp)return;
  const vw=v.videoWidth,vh=v.videoHeight;
  if(!vw||!vh)return;
  const rot=desiredPreviewRotation(),turned=rot%180!==0,aspect=turned?vh/vw:vw/vh;
  const stage=preview.getBoundingClientRect();
  if(!stage.width||!stage.height)return;
  const key=[vw,vh,rot,Math.round(stage.width),Math.round(stage.height)].join('x');
  if(key===lastLayoutKey)return;
  lastLayoutKey=key;previewRot=rot;
  preview.style.setProperty('--camera-aspect',aspect);
  vp.style.setProperty('--camera-aspect',aspect);
  let fitW,fitH;
  if(stage.width/stage.height>aspect){fitH=stage.height;fitW=stage.height*aspect;}
  else{fitW=stage.width;fitH=stage.width/aspect;}
  const boxW=Math.floor(fitW),boxH=Math.floor(fitH);
  vp.style.width=boxW+'px';
  vp.style.height=boxH+'px';
  if(rotor){rotor.style.setProperty('--vw',(turned?boxH:boxW)+'px');rotor.style.setProperty('--vh',(turned?boxW:boxH)+'px');rotor.style.setProperty('--rot',rot+'deg');}
}
// iOS often fires no video "resize" when the phone is turned: also poll while the camera is open.
setInterval(()=>{if(!$('cameraOverlay').hidden)previewDimensions();},350);
for(const type of ['resize','orientationchange'])window.addEventListener(type,()=>{lastLayoutKey='';previewDimensions();setTimeout(()=>{lastLayoutKey='';previewDimensions();},350);});
window.visualViewport?.addEventListener('resize',()=>{lastLayoutKey='';previewDimensions();});
screen.orientation?.addEventListener?.('change',()=>{lastLayoutKey='';previewDimensions();setTimeout(()=>{lastLayoutKey='';previewDimensions();},350);});
// A tap lands on the turned picture; the camera wants coordinates in its own (unturned) frame.
function unrotatePoint(x,y,rot){return rot===90?{x:y,y:1-x}:rot===180?{x:1-x,y:1-y}:rot===270?{x:1-y,y:x}:{x,y};}
// Real full screen where the browser allows it (Android Chrome, iPad); portrait lock where the OS allows it.
// iPhone Safari offers neither: there the "add to Home Screen" app is the only way to lose the address bar.
function lockPortrait(){try{screen.orientation?.lock?.('portrait')?.catch?.(()=>{});}catch{}}
let fullscreenWanted=false;
function enterCameraFullscreen(){
  const el=$('cameraOverlay'),request=el.requestFullscreen||el.webkitRequestFullscreen;
  fullscreenWanted=true;
  if(request&&!document.fullscreenElement&&!document.webkitFullscreenElement){
    try{
      const p=request.call(el,{navigationUI:'hide'});
      // The request is asynchronous: if the camera was refused/closed meanwhile, do not stay stuck in full screen.
      p?.then?.(()=>{if(fullscreenWanted)lockPortrait();else leaveCameraFullscreen();},()=>{});p?.catch?.(()=>{});
    }catch{}
  }
  lockPortrait();
}
function leaveCameraFullscreen(){
  fullscreenWanted=false;
  try{if(document.fullscreenElement||document.webkitFullscreenElement){const p=(document.exitFullscreen||document.webkitExitFullscreen).call(document);p?.catch?.(()=>{});}}catch{}
  try{screen.orientation?.unlock?.();}catch{}
}
async function exitFullscreenAndWait(){
  if(!(document.fullscreenElement||document.webkitFullscreenElement))return;
  try{await Promise.race([(document.exitFullscreen||document.webkitExitFullscreen).call(document),new Promise(r=>setTimeout(r,500))]);}catch{}
}
function standaloneApp(){return navigator.standalone===true||!!window.matchMedia?.('(display-mode: standalone)').matches;}
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
    // Full screen consumes the click's activation, so it must come AFTER the (iOS) sensor permission request.
    enterCameraFullscreen();
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
  const turned=previewRot%180!==0,mapping=mapPreviewTapToSensor(tapX,tapY,vpRect.width,vpRect.height,turned?video.videoHeight:video.videoWidth,turned?video.videoWidth:video.videoHeight,z?.mode,z?.factor);
  if(!mapping.inBounds){if(status)status.textContent=mapping.reason;return;}
  {const raw=unrotatePoint(mapping.sensorX,mapping.sensorY,previewRot);mapping.sensorX=raw.x;mapping.sensorY=raw.y;}
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
  if(Number.isFinite(e.beta)&&Number.isFinite(e.gamma)){
    deviceTilt={beta:e.beta,gamma:e.gamma,readingAt:Date.now(),screenAngle:screen.orientation?.angle??window.orientation??null};
    if(desiredPreviewRotation()!==previewRot)previewDimensions();
  }else if(!e.isTrusted){
    deviceTilt=null;
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
    let lastSavedPhoto=null;
    while(pending.length){
      const item=pending[0];
      say('正在儲存原始影像…');
      try{
        await db.savePhoto(item.photo,item.blob);
        pending.shift();
        lastSavedPhoto=item.photo;
      }catch(err){
        say('儲存原始影像失敗：'+err.message+'；照片保留於待存清單，請點「重試未完成照片儲存」。');
        updateRetryUI();
        throw err;
      }
    }
    await load();
    say('已儲存於本機（未上傳）'+(lastSavedPhoto?.source==='camera'?'；'+areaReferenceText(areaReference(lastSavedPhoto)):''));
    updateRetryUI();
  })().finally(()=>{commitPromise=null;});
  return commitPromise;
}
const retry=text('button','重試未完成照片儲存');retry.hidden=true;retry.onclick=()=>run(commitPending);$('camera').append(retry);
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
    const v=$('video'),tilt=buttonReading(deviceTilt,buttonPressedAt,10000)??direction;
    const captureRotation=v?.videoWidth&&v?.videoHeight?cameraAutoRotation({width:v.videoWidth,height:v.videoHeight,screenAngle,beta:tilt?.beta,gamma:tilt?.gamma}):0;
    try{
      if(useZoomFrame||captureRotation!==0||!window.ImageCapture)throw Error();
      blob=await new ImageCapture(track).takePhoto();method='ImageCapture.takePhoto';
    }catch(e){
      if(session!==cameraSession||track.readyState==='ended')throw Error('相機已停止，未擷取影像');
      if(!v.videoWidth)throw Error('相機影格尚未就緒');
      const crop=cropRect(v.videoWidth,v.videoHeight,zoom.mode==='browser-digital-crop'?zoom.factor:1),swapped=captureRotation%180!==0,c=document.createElement('canvas');c.width=swapped?crop.height:crop.width;c.height=swapped?crop.width:crop.height;
      const ctx=c.getContext('2d');
      if(captureRotation){ctx.translate(c.width/2,c.height/2);ctx.rotate(captureRotation*Math.PI/180);ctx.drawImage(v,crop.sx,crop.sy,crop.sw,crop.sh,-crop.width/2,-crop.height/2,crop.width,crop.height);}
      else ctx.drawImage(v,crop.sx,crop.sy,crop.sw,crop.sh,0,0,c.width,c.height);
      frameCopiedAt=Date.now();outputWidth=c.width;outputHeight=c.height;
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
      cameraFacing:currentFacing,cameraBearing:finalBearing,cameraBearingLabel:finalLabel,cameraBearingAlgorithm:finalAlg,cameraBearingEvidence:finalEvidence,cameraTransformEvidence:finalEvidence??'未驗證裝置／螢幕／鏡頭轉換',workDay,screenAngle,zoom,outputWidth,outputHeight,displayRotation:0}});
    await commitPending();
  }finally{busy=false;zoomUI();}
});
$('import').onchange=()=>run(async()=>{const workDay=activeDay;for(const file of $('import').files){if(!file.type.startsWith('image/'))throw Error('請選取影像');pending.push({blob:file,photo:{id:crypto.randomUUID(),source:'gallery',originalName:file.name,mime:file.type,capturedAt:null,importedAt:Date.now(),fileLastModifiedAt:file.lastModified,location:null,heading:null,sourceExif:null,workDay,captureMethod:'原始檔案匯入；未解析 EXIF，預覽由瀏覽器解碼方向',...(file.type!=='image/png'?{displayRotation:0}:{})}});}await commitPending();$('import').value='';});
function choose(id){if(multi){selected.has(id)?selected.delete(id):selected.add(id);}else selected=selected.has(id)?new Set():new Set([id]);render();}
function render(){$('editingTab').hidden=true;$('appendToEditor').hidden=!edit;releaseImages('gallery');releaseImages('groups');$('gallery').replaceChildren();const days=[...new Set([...dayState.days,...photos.map(workDayOf),activeDay])].filter(validDay).sort();$('workDay').replaceChildren();for(const d of days){const o=text('option',`${d}（${photos.filter(p=>workDayOf(p)===d).length} 張）`);o.value=d;$('workDay').append(o);}$('workDay').value=activeDay;const assignedIds=new Set(groups.flatMap(g=>g.photoIds));$('currentDay').replaceChildren(...[activeDay,`拍攝 ${dayPhotos().filter(p=>p.source==='camera').length} 張`,`相簿 ${dayPhotos().filter(p=>p.source==='gallery').length} 張`,`待整理 ${dayPhotos().filter(p=>!assignedIds.has(p.id)).length} 張`,`紀錄集 ${groups.filter(groupInDay).length} 組`].map(label=>text('span',label)));for(const id of selected)if(!dayPhotos().some(p=>p.id===id))selected.delete(id);const assigned=new Set(groups.flatMap(g=>g.photoIds));for(const p of dayPhotos().filter(p=>!assigned.has(p.id)||selected.has(p.id))){const card=text('div','');card.className='card';const b=text('button',`${selected.has(p.id)?'已選 · ':''}${photoLabel(p)} · ${p.source==='camera'?'拍攝':'相簿'}${validCrop(p.crop)?' · 已裁切':''}`);b.setAttribute('aria-pressed',selected.has(p.id));b.prepend(photoImage(p.id,'gallery'));let timer,long=false;b.onpointerdown=e=>{if(e.button!==0)return;long=false;timer=setTimeout(()=>{long=true;multi=true;selected.add(p.id);render();},550);};b.onpointerup=b.onpointerleave=b.onpointercancel=()=>clearTimeout(timer);b.onclick=()=>{if(long){long=false;return;}choose(p.id);};const coordinates=text('p',compactCoordinates(p.location));coordinates.className='gallery-coordinates';card.append(b,photoAreaInfo(p,{collapsed:true}),coordinates);$('gallery').append(card);}$('multi').textContent=multi?'結束多選':'選取多張';$('multi').setAttribute('aria-pressed',multi);$('selectionCount').textContent=`已選 ${selected.size} 張；本日待整理 ${dayPhotos().filter(p=>!assigned.has(p.id)).length} 張`;$('appendToEditor').disabled=$('deleteSelected').disabled=$('rotateLeft').disabled=$('rotateRight').disabled=$('compose').disabled=!selected.size;$('cropSelected').disabled=selected.size!==1;$('returnToEditor').hidden=!edit;$('compose').hidden=!!edit;$('groupList').replaceChildren();if(!groups.some(groupInDay))$('groupList').append(text('p','本工作日尚無紀錄集。請到待整理照片選取照片並記事。'));const ordered=groups.filter(groupInDay).sort(compareGroups);$('toggleGroupOrder').disabled=ordered.length<2;$('groupOrderHint').textContent=ordered.length<2?'至少需要 2 組才能調整順序':'即時儲存 · 所有參考與匯出沿用此順序';for(const [index,g] of ordered.entries()){const a=text('article','');a.dataset.groupId=g.id;const orderTools=text('div','');orderTools.className='group-order-controls';orderTools.hidden=!groupOrdering;for(const [label,delta] of [['↑ 上移',-1],['↓ 下移',1]]){const b=text('button',label);b.dataset.direction=String(delta);b.disabled=index+delta<0||index+delta>=ordered.length;b.setAttribute('aria-label',`${groupTitle(g.photoIds)} ${label}`);b.onclick=()=>run(async()=>{await db.saveGroupOrder(activeDay,move(ordered.map(x=>x.id),index,delta));await load();say('紀錄集順序已儲存');});orderTools.append(b);}a.append(orderTools,text('h3',`${index+1}. ${groupTitle(g.photoIds)}`),text('p',`${g.photoIds.length} 張 · 更新 ${new Date(g.updatedAt).toLocaleString()}`));const strip=text('div','');strip.className='group-photos';g.photoIds.forEach((id,i)=>{const p=photos.find(p=>p.id===id),figure=text('figure','');figure.append(photoImage(id,'groups'),text('figcaption',compactPhotoLines(p,i)[0]));strip.append(figure);});a.append(strip);const info=text('details','');info.append(text('summary','照片資訊與區域判讀'));g.photoIds.forEach((id,i)=>{const row=text('div','');row.append(text('strong',compactPhotoLines(photos.find(p=>p.id===id),i)[0]),photoInfo(photos.find(p=>p.id===id)));info.append(row);});a.append(info);for(const [label,value] of [['工程描述',g.engineering]]){const note=text('p',`${label}：${value||'未填寫'}`);note.className='note-text';a.append(note);}const e=text('button','編輯／排序照片');e.onclick=()=>openEditor(g);const d=text('button','移除紀錄集（保留原始照片）');d.className='danger';d.onclick=()=>run(async()=>{if(edit?.id===g.id){say('請先關閉或儲存此紀錄集的編輯，再移除。');return;}if(confirm('只移除此紀錄集，原始照片仍保留？')){await db.removeGroup(g.id);await load();say('已移除紀錄集，照片仍保留');}});a.append(e,d);$('groupList').append(a);}if(edit&&active==='editor')renderOrder();}
$('multi').onclick=()=>{multi=!multi;if(!multi)selected=new Set([...selected].slice(0,1));render();};
let groupOrdering=false,editingExisting=false;
$('toggleGroupOrder').onclick=()=>{groupOrdering=!groupOrdering;$('toggleGroupOrder').setAttribute('aria-pressed',String(groupOrdering));render();};
$('returnToEditor').onclick=()=>tab('editor');
function openEditor(g,ids=[...selected]){if(!canDiscardEditor())return;editingExisting=!!g;edit=structuredClone(g||{id:crypto.randomUUID(),title:'',notes:'',engineering:'',photoIds:ids,createdAt:Date.now(),workDay:activeDay});$('title').value=groupTitle(edit.photoIds);$('engineering').value=edit.engineering;editBaseline=editorValue();tab('editor');}
function renderOrder(focus){$('title').value=groupTitle(edit.photoIds);releaseImages('order');$('order').replaceChildren();$('photoInfo').replaceChildren();const previews=text('div','');previews.className='editor-previews';previews.dataset.count=edit.photoIds.length;edit.photoIds.forEach((id,i)=>{const figure=text('figure','');figure.append(photoImage(id,'order'),text('figcaption',compactPhotoLines(photos.find(p=>p.id===id),i)[0]));previews.append(figure);});$('order').append(previews);edit.photoIds.forEach((id,i)=>{const p=photos.find(p=>p.id===id),row=text('article',''),head=text('div',''),controls=text('div','');head.className='photo-row-head';controls.className='photo-row-controls';head.append(text('h3',compactPhotoLines(p,i)[0]));for(const [label,delta] of [['上移',-1],['下移',1]]){const b=text('button',delta<0?'↑':'↓');b.title=label;b.setAttribute('aria-label',`${i+1}. ${photoLabel(p)} ${label}`);b.dataset.photoId=id;b.dataset.action=label;b.disabled=i+delta<0||i+delta>=edit.photoIds.length;b.onclick=()=>{edit.photoIds=move(edit.photoIds,i,delta);renderOrder({id,action:label});say(`${photoLabel(p)} 已${label}至第 ${edit.photoIds.indexOf(id)+1} 張`);};controls.append(b);}const cr=text('button','裁切');cr.title='裁切照片（保持原比例）';cr.setAttribute('aria-label',`${i+1}. ${photoLabel(p)} 裁切`);cr.dataset.photoId=id;cr.dataset.action='裁切';cr.onclick=()=>run(()=>cropPhoto(id));controls.append(cr);const r=text('button','×');r.title='從此組移除';r.setAttribute('aria-label',`${i+1}. ${photoLabel(p)} 從此組移除`);r.dataset.photoId=id;r.dataset.action='從此組移除';r.onclick=()=>{edit.photoIds.splice(i,1);renderOrder({id:edit.photoIds[Math.min(i,edit.photoIds.length-1)],action:'從此組移除'});say(`${photoLabel(p)} 已從此組移除，原始照片仍保留`);};controls.append(r);head.append(controls);row.append(head,photoInfo(p));$('photoInfo').append(row);});if(focus){const controls=[...$('photoInfo').querySelectorAll('button')].filter(b=>b.dataset.photoId===focus.id&&!b.disabled);(controls.find(b=>b.dataset.action===focus.action)||controls[0]||$('addSelected')).focus();}}
$('compose').onclick=()=>{if(!selected.size){say('請先選照片');return;}openEditor();};$('appendToEditor').onclick=()=>{if(!edit)return;const day=edit.workDay??workDayOf(photos.find(p=>p.id===edit.photoIds[0])??{});if([...selected].some(id=>workDayOf(photos.find(p=>p.id===id))!==day)){say('請勿將不同工作日的照片加入同一記事。');return;}edit.photoIds=[...new Set([...edit.photoIds,...selected])];tab('editor');};$('addSelected').onclick=()=>{if(edit){tab('photos');say('請選取照片，再按「加入目前記事並返回」。');}};$('cancel').onclick=()=>{if(!canDiscardEditor())return;const returnTab=groups.some(g=>g.id===edit?.id)?'groups':'photos';$('editor').hidden=true;edit=null;editBaseline=null;releaseImages('order');render();tab(returnTab);};$('save').onclick=()=>run(async()=>{if(!edit)return;const title=groupTitle(edit.photoIds),g={...edit,title,notes:edit.notes??'',engineering:$('engineering').value,updatedAt:Date.now()};if(edit.title&&edit.title!==title&&!edit.legacyTitle)g.legacyTitle=edit.title;validateGroup(g,new Set(photos.map(p=>p.id)));await db.saveGroup(g,{existing:editingExisting});$('editor').hidden=true;edit=null;editBaseline=null;releaseImages('order');selected.clear();await load();say('紀錄集已儲存');tab('groups');});
$('backup').onclick=()=>run(async()=>{say('正在建立本機備份（交易快照；不含未提交資料）…');const snapshot=await db.snapshot();const entries=[],records=[];for(const p of snapshot.photos){const blob=snapshot.assets.find(a=>a.id===p.id)?.original;if(!blob)throw Error('原始影像缺失');const path=`originals/${p.id}.bin`;entries.push([path,blob]);records.push({...p,path,size:blob.size,sha256:await digest(blob)});}for(const g of snapshot.groups)validateGroup(g,new Set(snapshot.photos.map(p=>p.id)));const manifest={schema:1,createdAt:Date.now(),snapshotStartedAt:snapshot.startedAt,snapshotCompletedAt:snapshot.completedAt,completeness:'單一唯讀交易中的已提交資料；不含未儲存表單／待存照片／快照之後提交資料',photos:records,groups:snapshot.groups,workDays:dayState.days,activeWorkDay:activeDay??localDay(),numberCounters:db.numberCounters()};entries.push(['area-reference.json',new Blob([JSON.stringify({schema:1,calculatedAt:Date.now(),meaning:'暫定區界及本次備份時計算的參考；還原照片後仍依 App 當前區界回算，不自動匯入或信任此區界',dataset:AREA_DATA,photos:snapshot.photos.map(p=>({photoId:p.id,reference:areaReference(p)}))},null,2)],{type:'application/json'})]);entries.push(['manifest.json',new Blob([JSON.stringify(manifest,null,2)],{type:'application/json'})],['notes.txt',new Blob([snapshot.groups.map(g=>`${g.title}\n工作日：${g.workDay??'舊版請參照各照片 workDay'}\n工程描述：${g.engineering}\n${g.notes?'舊版備註（保留）：'+g.notes+'\n':''}${g.legacyTitle?'舊版名稱（保留）：'+g.legacyTitle+'\n':''}順序：${g.photoIds.join(', ')}\n`).join('\n')])],['README.txt',new Blob(['原始影像在 originals/；檔名與 MIME 在 manifest.json。GPS 是裝置位置；方向不是可靠鏡頭朝向。相機讀值以按鈕時間為基準，不是曝光資料。舊版 capturedAt / ageAtShutterMs 亦是按鈕時間／年齡。沒有 EXIF 嵌入或雲端同步。'])]);const archive=await tar(entries),url=URL.createObjectURL(archive),a=document.createElement('a');a.href=url;a.download=`field-notebook-${Date.now()}.tar`;a.click();setTimeout(()=>URL.revokeObjectURL(url),60000);say('已交給瀏覽器下載；請確認下載完成（無法驗證使用者是否取消）');});
$('restore').onchange=()=>run(async()=>{const file=$('restore').files[0];if(!file)return;say('驗證備份中…');const entries=await untar(file),m=JSON.parse(await entries.get('manifest.json')?.text());if(m.schema!==1||!Array.isArray(m.photos)||!Array.isArray(m.groups))throw Error('不支援的備份');if(m.workDays!==undefined&&(!Array.isArray(m.workDays)||m.workDays.some(d=>!validDay(d))))throw Error('工作日格式無效');if(m.activeWorkDay!==undefined&&!validDay(m.activeWorkDay))throw Error('目前工作日格式無效');if(m.numberCounters!==undefined&&(typeof m.numberCounters!=='object'||m.numberCounters===null||Array.isArray(m.numberCounters)||Object.entries(m.numberCounters).some(([d,n])=>!validDay(d)||!Number.isSafeInteger(n)||n<1)))throw Error('編號計數格式無效');const incomingNumbers=new Set();const ids=new Set(),gids=new Set(),staged=[];for(const p of m.photos){if(p.workDay!==undefined&&!validDay(p.workDay)||p.photoNumber!==undefined&&(!Number.isSafeInteger(p.photoNumber)||p.photoNumber<1)||p.displayRotation!==undefined&&![0,90,180,270].includes(p.displayRotation))throw Error('照片工作日／編號／旋轉格式無效');if(p.crop!==undefined&&!validCrop(p.crop))throw Error('照片裁切資料格式無效');if(p.workDay&&p.photoNumber){const numberKey=p.workDay+':'+p.photoNumber;if(incomingNumbers.has(numberKey)||photos.some(x=>x.workDay===p.workDay&&x.photoNumber===p.photoNumber))throw Error('工作日照片編號重複或與本機衝突');incomingNumbers.add(numberKey);}if(typeof p.id!=='string'||!/^[-a-zA-Z0-9]+$/.test(p.id)||ids.has(p.id)||photos.some(x=>x.id===p.id)||p.path!==`originals/${p.id}.bin`||typeof p.originalName!=='string'||!['camera','gallery'].includes(p.source)||typeof p.mime!=='string'||!p.mime.startsWith('image/'))throw Error('照片格式／ID 重複或與本機衝突');ids.add(p.id);const b=entries.get(p.path);if(!b||b.size!==p.size||await digest(b)!==p.sha256)throw Error('影像大小或雜湊不符');const {path,size,sha256,...photo}=p;staged.push({photo,blob:new Blob([b],{type:p.mime})});}for(const g of m.groups){if(typeof g.id!=='string'||typeof g.title!=='string'||typeof g.notes!=='string'||typeof g.engineering!=='string'||gids.has(g.id)||groups.some(x=>x.id===g.id))throw Error('紀錄集格式／ID 衝突');gids.add(g.id);validateGroup(g,ids);}if(!confirm(`已驗證 ${staged.length} 張照片、${m.groups.length} 組，匯入本機？`))return;await db.transaction(['photos','assets','groups'],tx=>{for(const x of staged){tx.objectStore('photos').add(x.photo);tx.objectStore('assets').add({id:x.photo.id,original:x.blob});}for(const g of m.groups)tx.objectStore('groups').add(g);});db.mergeCounters(m.numberCounters);for(const day of m.workDays??[])if(validDay(day))dayState.days.push(day);dayState.days=[...new Set(dayState.days)];activeDay=validDay(m.activeWorkDay)?m.activeWorkDay:null;if(activeDay)rememberDay(activeDay);await load();say('備份已完整還原（原圖與顯示旋轉分開保存）');$('restore').value='';});
async function changeDay(day){
  if(!validDay(day))throw Error('日期請使用有效 YYYY-MM-DD');
  if(busy||pending.length||rotating)throw Error('請先完成照片儲存或旋轉再切換工作日');
  if(!canDiscardEditor()){$('workDay').value=activeDay;return;}
  stop();closeCoordinateAlbum(false);closePhrasePickers();

  // 1. Bind form saves to currently displayed day: flush pending edits before transition
  const prevDisplayed = monDisplayedDay;
  if(prevDisplayed&&$('monForm').dataset.hydratedDay===prevDisplayed){
    monDraftManager.saveDraft(prevDisplayed, readMonInputs(true));
  }

  // 2. Protect transition and handle rapid switches with generation token
  const currentSwitch = ++daySwitchGeneration;
  dayTransitioning = true;monitoringControls();

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
    dayTransitioning = false;monitoringControls();
  }

  monDisplayedDay = day;
  if($('monForm')) $('monForm').dataset.displayedDay = day;
  updateMonUI();monitoringControls();
  say(`已切換工作日 ${day}，其他日期照片仍保留`);
}
$('workDay').onchange=()=>run(async()=>{try{await changeDay($('workDay').value);}catch(e){$('workDay').value=activeDay;throw e;}});
$('newDay').onclick=()=>run(async()=>{const day=prompt('新工作日日期（YYYY-MM-DD）。只分開整理，不清除照片，也不改拍照時間。',localDay());if(day===null)return;await changeDay(day.trim());});
$('deleteSelected').onclick=()=>run(async()=>{if(edit&&(editorDirty()||groups.some(g=>g.id===edit.id))){say('請先儲存或關閉記事，再刪除照片。');return;}const ids=[...selected];if(!ids.length)return;if(groups.some(g=>g.photoIds.some(id=>ids.includes(id))))throw Error('已有記事的照片請先移出紀錄集再刪除');if(!confirm(`永久刪除 ${ids.length} 張照片及原圖？無法復原，建議先下載完整備份。`))return;await db.deletePhotos(ids);if(edit){edit=null;editBaseline=null;$('editor').hidden=true;releaseImages('order');}for(const id of ids){rotations.delete(id);for(const k of previewCache.keys())if(k.startsWith(id+':'))previewCache.delete(k);}selected.clear();await load();say(`已永久刪除 ${ids.length} 張照片；其他照片編號不變`);});
async function rotateSelected(delta){if(rotating||!selected.size)return;rotating=true;$('rotateLeft').disabled=$('rotateRight').disabled=true;say('正在調整顯示方向…');try{for(const id of [...selected]){const p=photos.find(p=>p.id===id);let r=p.displayRotation??rotations.get(id);if(r===undefined){const a=assets.find(a=>a.id===id);const url=URL.createObjectURL(a.original);try{const im=new Image();im.src=url;await im.decode();r=previewRotation(p,im.naturalWidth,im.naturalHeight);}finally{URL.revokeObjectURL(url);}}await db.setPhotoRotation(id,(r+delta+360)%360,delta);for(const k of previewCache.keys())if(k.startsWith(id+':'))previewCache.delete(k);}await load();say('已儲存顯示方向，原始影像位元組未改動');}finally{rotating=false;$('rotateLeft').disabled=$('rotateRight').disabled=!selected.size;}}
$('rotateLeft').onclick=()=>run(()=>rotateSelected(-90));$('rotateRight').onclick=()=>run(()=>rotateSelected(90));
// Simple crop: a window that keeps the photo's own 4:3 shape; shrink it inwards and slide it to leave things out.
async function cropPhoto(id){
  const p=photos.find(x=>x.id===id),a=assets.find(x=>x.id===id);if(!p||!a)throw Error('找不到照片，請重新整理。');
  if(rotating)return;
  const result=await openCropDialog({photo:p,original:a.original});
  if(result===undefined)return;
  await db.setPhotoCrop(id,result);for(const k of previewCache.keys())if(k.startsWith(id+':'))previewCache.delete(k);
  await load();say(result?'已儲存裁切（'+cropSummary(result)+'）；原始影像位元組未改動。':'已還原為完整照片。');
}
$('cropSelected').onclick=()=>run(()=>{if(selected.size!==1){say('請只選一張照片再裁切。');return;}return cropPhoto([...selected][0]);});
let exportSnapshot=null,exportDay=null,exportOrder=[],exportSelected=new Set(),exportBlob=null,exportWorking=false,exportGeneration=0,exportController=null,exportUrls=[];
function resetExportPreview(){exportBlob=null;exportUrls.forEach(URL.revokeObjectURL);exportUrls=[];$('exportPreview').replaceChildren();$('downloadWord').disabled=true;}
function exportControls(){const ids=exportOrder.filter(id=>exportSelected.has(id));for(const el of $('exportRows').querySelectorAll('input,button'))el.disabled=exportWorking||el.dataset.edge==='true';$('exportAll').disabled=$('exportNone').disabled=exportWorking;$('prepareWord').disabled=exportWorking||!ids.length;$('downloadWord').disabled=exportWorking||!exportBlob;$('exportSelection').textContent=`已選 ${ids.length} 組、${ids.reduce((n,id)=>n+(exportSnapshot?.groups.find(g=>g.id===id)?.photoIds.length??0),0)} 張照片`;} 
function renderExportChoices(focus){$('exportRows').replaceChildren();exportOrder.forEach((id,i)=>{const g=exportSnapshot.groups.find(g=>g.id===id),name=numberedGroup(g.photoIds,exportSnapshot.photos),row=text('li',''),label=text('label',''),check=document.createElement('input');check.type='checkbox';check.checked=exportSelected.has(id);check.dataset.groupId=id;check.setAttribute('aria-label',`選取 ${name}`);check.onchange=()=>{check.checked?exportSelected.add(id):exportSelected.delete(id);resetExportPreview();$('exportStatus').textContent='選取已改變，請重新製作預覽。';exportControls();};label.append(check,text('span',`${i+1}. ${name} · ${g.photoIds.length} 張`));row.append(label,text('p',g.engineering.length>180?g.engineering.slice(0,180)+'…':g.engineering||'工程描述未填寫'));const actions=text('div','');actions.className='actions';for(const [name,delta] of [['上移',-1],['下移',1]]){const b=text('button',name);b.dataset.groupId=id;b.dataset.action=name;b.dataset.edge=String(i+delta<0||i+delta>=exportOrder.length);b.setAttribute('aria-label',`第 ${i+1} 組 ${numberedGroup(g.photoIds,exportSnapshot.photos)} ${name}`);b.onclick=()=>{exportOrder=move(exportOrder,i,delta);resetExportPreview();$('exportStatus').textContent='匯出順序已改變，請重新製作預覽。';renderExportChoices({id,action:name});};actions.append(b);}row.append(actions);$('exportRows').append(row);});exportControls();if(focus){const buttons=[...$('exportRows').querySelectorAll('button')].filter(b=>b.dataset.groupId===focus.id&&!b.disabled);(buttons.find(b=>b.dataset.action===focus.action)||buttons[0]||$('prepareWord')).focus();}}
$('exportWord').onclick=()=>run(async()=>{if($('exportDialog').open)return;if(!validDay(activeDay)){say('工作日尚未載入，請稍候再開啟匯出。');return;}if(editorDirty()||pending.length||busy||rotating){say('請先儲存或關閉未完成的記事／照片，再匯出已儲存紀錄集。');return;}const day=activeDay,snapshot=await db.snapshot();if($('exportDialog').open||day!==activeDay||$('monitoring').hidden||editorDirty()||pending.length||busy||rotating)return;const byId=new Map(snapshot.photos.map(p=>[p.id,p]));const included=snapshot.groups.filter(g=>g.workDay===day||g.photoIds.some(id=>byId.has(id)&&workDayOf(byId.get(id))===day)).sort(compareGroups);if(!included.length){say('本工作日沒有已儲存紀錄集可匯出。');return;}exportGeneration++;exportSnapshot=snapshot;exportDay=day;exportOrder=included.map(g=>g.id);exportSelected=new Set(exportOrder);exportWorking=false;resetExportPreview();$('exportStatus').textContent='請確認選取與順序，再按「預覽並製作表格」。';$('exportSnapshotInfo').textContent=`工作日 ${day} · 已儲存快照 ${new Date(snapshot.completedAt).toLocaleString()}；不修改原圖或記事。`;renderExportChoices();$('exportDialog').showModal();});
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
 $('monOverwriteConflict').hidden = !isConflict;
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
 if(dayTransitioning)return;
 const d = displayedDay();
 if(!validDay(d))return;
 $('monForm').dataset.hydratedDay=d;
 refreshStandaloneOptions();refreshTemplateRestore();
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
let isTransientOffer = false,suppressScopeOffer=false;

$('monWorkScope').addEventListener('focus', () => {
 if(!suppressScopeOffer&&$('monWorkScope').value === ''){
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
function closeCoordinateAlbum(restore=true){
 const back=coordinateReturn;coordinateReturn=null;coordinateAlbum=null;coordinateAlbumGeneration++;
 const dialog=$('coordinateAlbumDialog');if(dialog?.open)dialog.close();
 releaseImages('reference');$('coordinateAlbumList')?.replaceChildren();$('coordinateAlbumPhoto')?.replaceChildren();$('coordinateAlbumInfo')?.replaceChildren();
 if(restore&&back&&!dayTransitioning&&back.day===displayedDay()&&!$('monitoring').hidden){const ta=$('monWorkScope');suppressScopeOffer=true;try{ta.focus({preventScroll:true});if(ta.value===back.value){ta.setSelectionRange(back.start,back.end,back.direction);ta.scrollTop=back.scrollTop;}}finally{suppressScopeOffer=false;}window.scrollTo(back.x,back.y);}
}
function coordinateAlbumEntries(snapshot,day,mode){
 const dayPhotos=snapshot.photos.filter(p=>(validDay(p.workDay)||Number.isFinite(p.buttonPressedAt??p.capturedAt??p.importedAt??p.createdAt))&&workDayOf(p)===day),byId=new Map(dayPhotos.map(p=>[p.id,p]));
 if(mode==='pending'){const assigned=new Set(snapshot.groups.flatMap(g=>g.photoIds));return dayPhotos.filter(p=>!assigned.has(p.id)).sort((a,b)=>(a.photoNumber??0)-(b.photoNumber??0)).map(p=>({photo:p,source:'待整理照片'}));}
 const entries=[];for(const g of [...snapshot.groups].sort(compareGroups)){if(g.workDay!=null&&g.workDay!==day)continue;const ids=g.photoIds.filter(id=>byId.has(id));if(!ids.length)continue;const source=numberedGroup(ids,snapshot.photos);for(const id of ids)entries.push({photo:byId.get(id),source,engineering:g.engineering||''});}return entries;
}
function renderCoordinateAlbum(index=-1){
 const state=coordinateAlbum;if(!state?.snapshot||!$('coordinateAlbumDialog').open)return;
 releaseImages('reference');for(const id of ['coordinateAlbumList','coordinateAlbumPhoto','coordinateAlbumInfo'])$(id).replaceChildren();
 state.index=index;$('coordinateAlbumStatus').textContent='';$('coordinateAlbumList').hidden=index>=0;$('coordinateAlbumDetail').hidden=index<0;
 $('coordinatePending').setAttribute('aria-pressed',state.mode==='pending');$('coordinateGroups').setAttribute('aria-pressed',state.mode==='groups');
 if(index<0){
  if(!state.entries.length){$('coordinateAlbumStatus').textContent=state.mode==='pending'?'本工作日沒有待整理照片':'本工作日沒有紀錄集照片';return;}
  for(const [i,entry] of state.entries.entries()){const button=text('button','');button.type='button';button.className='coordinate-thumbnail';button.append(photoImage(entry.photo.id,'reference',state.snapshot),text('strong',photoLabel(entry.photo)),text('small',entry.source));if(entry.engineering){const note=text('small',entry.engineering);note.className='coordinate-source-note';button.append(note);}button.onclick=()=>{state.listScroll=$('coordinateAlbumList').scrollTop;renderCoordinateAlbum(i);$('coordinateAlbumValue').focus({preventScroll:true});};$('coordinateAlbumList').append(button);}
  $('coordinateAlbumList').scrollTop=state.listScroll||0;return;
 }
 const entry=state.entries[index];if(!entry){renderCoordinateAlbum();return;}
 const image=photoImage(entry.photo.id,'reference',state.snapshot);image.classList.add('coordinate-image');
 let touch=null;image.onpointerdown=e=>{touch=e.pointerType==='touch'?{x:e.clientX,y:e.clientY}:null;};image.onpointercancel=()=>{touch=null;};image.onpointerup=e=>{if(!touch)return;const dx=e.clientX-touch.x,dy=e.clientY-touch.y;touch=null;if(Math.abs(dx)>45&&Math.abs(dy)<60){const next=index+(dx<0?1:-1);if(next>=0&&next<state.entries.length)renderCoordinateAlbum(next);}};
 $('coordinateAlbumPhoto').append(text('h3',`${photoLabel(entry.photo)} · ${index+1}/${state.entries.length}`),image,text('p',entry.source));
 if(entry.engineering){const note=text('p',`工程描述：${entry.engineering}`);note.className='note-text';$('coordinateAlbumPhoto').append(note);}
 const coordinates=compactCoordinates(entry.photo.location);$('coordinateAlbumValue').value=coordinates;$('copyPhotoCoordinates').disabled=coordinates==='座標未取得';
 const info=photoInfo(entry.photo);info.querySelector(':scope > p:nth-child(2)')?.remove();$('coordinateAlbumInfo').append(info);$('coordinatePrevious').disabled=index===0;$('coordinateNext').disabled=index===state.entries.length-1;
}
$('viewPhotoCoordinates').onclick=async()=>{
 const context=phraseContext('monitoring');if(!context||$('coordinateAlbumDialog').open)return;
 const ta=$('monWorkScope');coordinateReturn={day:context.day,value:ta.value,start:ta.selectionStart,end:ta.selectionEnd,direction:ta.selectionDirection,scrollTop:ta.scrollTop,x:window.scrollX,y:window.scrollY};
 const generation=++coordinateAlbumGeneration;coordinateAlbum={snapshot:null,day:context.day,mode:'pending',entries:[],index:-1};
 $('coordinateAlbumDay').textContent=`工作日：${context.day}`;$('coordinateAlbumStatus').textContent='讀取中…';$('coordinateAlbumDetail').hidden=true;$('coordinateAlbumList').hidden=false;$('coordinateAlbumDialog').showModal();
 try{const snapshot=await db.snapshot();if(generation!==coordinateAlbumGeneration)return;if(context.key!==phraseContext('monitoring')?.key){closeCoordinateAlbum(false);return;}coordinateAlbum.snapshot=snapshot;coordinateAlbum.entries=coordinateAlbumEntries(snapshot,context.day,'pending');renderCoordinateAlbum();}
 catch(e){if(generation===coordinateAlbumGeneration)$('coordinateAlbumStatus').textContent='無法讀取照片：'+e.message;}
};
$('closeCoordinateAlbum').onclick=()=>closeCoordinateAlbum();
$('coordinateAlbumDialog').addEventListener('cancel',e=>{e.preventDefault();closeCoordinateAlbum();});
$('coordinateAlbumDialog').addEventListener('close',()=>{if(!$('coordinateAlbumDialog').open&&coordinateReturn)closeCoordinateAlbum();});
for(const [id,mode] of [['coordinatePending','pending'],['coordinateGroups','groups']])$(id).onclick=()=>{if(!coordinateAlbum?.snapshot)return;coordinateAlbum.mode=mode;coordinateAlbum.listScroll=0;coordinateAlbum.entries=coordinateAlbumEntries(coordinateAlbum.snapshot,coordinateAlbum.day,mode);renderCoordinateAlbum();};
$('coordinateAlbumBack').onclick=()=>renderCoordinateAlbum();
$('coordinatePrevious').onclick=()=>{if(coordinateAlbum?.index>0)renderCoordinateAlbum(coordinateAlbum.index-1);};
$('coordinateNext').onclick=()=>{if(coordinateAlbum&&coordinateAlbum.index+1<coordinateAlbum.entries.length)renderCoordinateAlbum(coordinateAlbum.index+1);};
$('coordinateAlbumDialog').addEventListener('keydown',e=>{if(!coordinateAlbum||coordinateAlbum.index<0||e.target.matches('input,textarea')||e.ctrlKey||e.metaKey||e.altKey)return;if(e.key==='ArrowLeft'){e.preventDefault();$('coordinatePrevious').click();}if(e.key==='ArrowRight'){e.preventDefault();$('coordinateNext').click();}});
$('copyPhotoCoordinates').onclick=async()=>{
 const state=coordinateAlbum,generation=coordinateAlbumGeneration,index=state?.index,entry=state?.entries[index];if(!entry||!$('coordinateAlbumDialog').open||state.day!==displayedDay())return;
 const value=compactCoordinates(entry.photo.location);if(value==='座標未取得')return;
 try{await navigator.clipboard.writeText(value);if(generation===coordinateAlbumGeneration&&coordinateAlbum?.index===index&&coordinateAlbum.entries[index]===entry)$('coordinateAlbumStatus').textContent='已複製座標';}
 catch{if(generation!==coordinateAlbumGeneration||coordinateAlbum?.index!==index||coordinateAlbum.entries[index]!==entry)return;$('coordinateAlbumStatus').textContent='無法自動複製，請長按或選取座標複製';$('coordinateAlbumValue').focus();$('coordinateAlbumValue').select();}
};
function phraseContext(kind){
 const day=displayedDay();
 if(dayTransitioning||!validDay(day)||day!==activeDay||day!==$('workDay').value||$('monForm').dataset.displayedDay!==day||monitoringWorking||uploadingPhotos)return null;
 if(kind==='monitoring'){if($('monitoring').hidden)return null;return {day,key:`monitoring:${day}:${daySwitchGeneration}`};}
 if(!edit||$('editor').hidden||(edit.workDay??workDayOf(photos.find(p=>p.id===edit.photoIds[0])??{}))!==day||!edit.photoIds.every(id=>photos.some(p=>p.id===id&&workDayOf(p)===day)))return null;
 return {day,key:JSON.stringify(['engineering',day,daySwitchGeneration,edit.id,edit.photoIds])};
}
function applyWorkPhrase(target,phrase,before,scope){
 const plan=phraseInsertion(before.value,before.start,before.end,phrase,{scope});
 suppressScopeOffer=true;
 try{target.focus();target.setSelectionRange(plan.start,plan.end);insertTextAtCaret(target,plan.text);}
 finally{suppressScopeOffer=false;}
 if(scope)isTransientOffer=false;
 target.dispatchEvent(new Event('input',{bubbles:true}));return plan;
}
function restoreWorkPhrase(target,x,scope){
 suppressScopeOffer=true;
 try{target.focus();target.setSelectionRange(x.start,x.start+x.length);insertTextAtCaret(target,'');target.setSelectionRange(x.before.start,x.before.end);}
 finally{suppressScopeOffer=false;}
 if(scope)isTransientOffer=false;target.dispatchEvent(new Event('input',{bubbles:true}));
}
function phrasePickerElements(prefix,panel,target){return {panel:$(panel),search:$(prefix+'PhraseSearch'),list:$(prefix+'PhraseList'),status:$(prefix+'PhraseStatus'),refresh:$('refresh'+(prefix==='photo'?'Photo':'Scope')+'Phrases'),undo:$('undo'+(prefix==='photo'?'Photo':'Scope')+'Phrase'),target:$(target)};}
photoPhrasePicker=new WorkPhrasePicker({
 ...phrasePickerElements('photo','photoPhrasePanel','monWorkScope'),getContext:()=>phraseContext('monitoring'),
 loadSource:async c=>{const snapshot=await readPhraseMetadata(await db.open()),choices=photoEngineeringPhrases(snapshot,c.day);return {choices,revision:JSON.stringify(choices.map(p=>[p.text,p.revision])),emptyMessage:'本工作日尚無工程描述片語' };},
 apply:(p,b)=>applyWorkPhrase($('monWorkScope'),p,b,true),restore:x=>restoreWorkPhrase($('monWorkScope'),x,true),
 successMessage:()=>monDraftManager.hasUnsaved(displayedDay())?{text:'已插入，但草稿尚未儲存。',quiet:false}:{text:'已插入',quiet:true}
});
scopePhrasePicker=new WorkPhrasePicker({
 ...phrasePickerElements('scope','scopePhrasePanel','engineering'),getContext:()=>phraseContext('engineering'),
 loadSource:async c=>{if(monDraftManager.hasUnsaved(c.day)||monDraftManager.hasCorrupt(c.day)||monDraftManager.hasConflict(c.day)||monDraftManager.isReadBlocked(c.day))throw Error('本工作日監看草稿尚未安全儲存，請先重試儲存或匯出備份；不使用舊草稿冒充最新內容');return {...readSavedScopePhrases(safeGetLocalStorage(),c.day),emptyMessage:'本工作日尚無作業子項目片語' };},
 apply:(p,b)=>applyWorkPhrase($('engineering'),p,b,false),restore:x=>restoreWorkPhrase($('engineering'),x,false),
 successMessage:()=>({text:'已插入，請按「儲存紀錄集」。',quiet:true})
});
$('syncStrata').onclick = () => {
 syncStrataBoth(true);
};
$('monRetrySave').onclick = () => {
 const d = displayedDay();
 if(monDraftManager.isReadBlocked(d)){
  monDraftManager.retryRead(d);
 }else{
  monDraftManager.saveDraft(d, readMonInputs(true));
 }
 updateMonStatusUI();
};
$('monOverwriteConflict').onclick = async () => {
 const d = displayedDay();
 if(!monDraftManager.hasConflict(d))return;
 const expectedRaw=monDraftManager.getConflict(d).externalRaw;
 if(!confirm('其他分頁已更新或刪除草稿。建議先匯出備份。確認以本頁記憶體內容覆寫外部草稿？取消不變更任何資料。'))return;
 // Let storage notifications queued during the modal confirmation reach this tab
 // before comparing the exact external revision the user agreed to replace.
 $('monOverwriteConflict').disabled=true;
 try{
  await new Promise(resolve=>setTimeout(resolve,0));
  monDraftManager.forceOverwriteConflict(d,expectedRaw);
  updateMonStatusUI();
 }finally{$('monOverwriteConflict').disabled=dayTransitioning;}
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
function templateBackup(day){
 const storage=safeGetLocalStorage(),id=storage?.getItem(`field-mon-template-latest-${day}`);if(!id)return null;
 if(!id.startsWith(`field-mon-template-backup-${day}-`))throw Error('表單備份索引異常，未還原。');
 const backup=JSON.parse(storage.getItem(id)||'null');if(backup?.day!==day||!Object.keys(MON_FIELDS).every(k=>typeof backup.fields?.[k]==='string'))throw Error('表單備份不完整，未還原。');return backup;
}
function refreshTemplateRestore(){try{$('restoreMonTemplate').hidden=!templateBackup(displayedDay());}catch{$('restoreMonTemplate').hidden=false;}}
function preserveTemplateForm(day,fields){
 if(!monDraftManager.saveDraft(day,fields).success)throw Error('目前表單未能安全儲存，先處理草稿問題；未取代欄位。');
 const storage=safeGetLocalStorage();if(!storage)throw Error('無法保存套用前的表單備份，未取代欄位。');
 const id=`field-mon-template-backup-${day}-${crypto.randomUUID()}`,backup={day,fields,rawDraft:storage.getItem(monKey(day)),createdAt:Date.now()};
 storage.setItem(id,JSON.stringify(backup));storage.setItem(`field-mon-template-latest-${day}`,id);refreshTemplateRestore();
}
$('fillMonTemplate').onclick=async()=>{
 if(monitoringWorking||uploadingPhotos||dayTransitioning)return;
 if(!masterReady()){$('monStatus').textContent='請先登入、選擇並檢查本月母表。';return;}
 const day=displayedDay(),generation=daySwitchGeneration,current=readMonInputs(true),input=$('masterDocUrl').value.trim(),preferred=$('masterTab').value||null;
 monitoringWorking=true;monitoringControls();
 try{
  if(docsAppendInfo.DOCS_APPEND_BUILD!=='date-order-v2')throw Error('請先安全更新 App，不要清除網站資料。');
  if(day!==$('workDay').value)throw Error('工作日尚未切換完成，未套用。');
  if(!validDay(day)||!input)throw Error('請先填入 Google 母表副本網址；多分頁文件請先檢查並選擇分頁。');
  $('monStatus').textContent='正在唯讀取得母表最後一個監看表格…';
  const result=await inspectMaster(input);acceptMasterCheck(result,input,preferred);
  const source=docsAppendInfo.lastMonitoringTemplate(result.document,$('masterTab').value||null);
  if(generation!==daySwitchGeneration||day!==displayedDay()||JSON.stringify(current)!==JSON.stringify(readMonInputs(true)))throw Error('讀取期間工作日或表單已變更，未取代欄位，請重新操作。');
  if(!confirm(`套用「${result.file.name}」最後表格（${source.date}）的文字欄位？\n目前表單將先保存備份，日期保留「${current.date}」，照片不帶入。`)){updateMonStatusUI();return;}
  preserveTemplateForm(day,current);
  const saved=monDraftManager.saveDraft(day,{...source.fields,date:current.date});fillMonInputs(saved.data);
  monDraftManager.setStatus(day,{status:saved.success?`已套用母表末筆 ${source.date}；監看日期保留。可在「Word 備份與其他操作」還原套用前表單。`:'已帶入文字，但草稿尚未儲存成功；套用前表單備份已保留，請先備份／處理儲存問題。',isError:!saved.success,isUnsaved:!saved.success});updateMonStatusUI();
 }catch(e){updateMonStatusUI();$('monStatus').textContent='套用未完成：'+e.message;say('套用未完成：'+e.message);}
 finally{monitoringWorking=false;monitoringControls();}
};
$('restoreMonTemplate').onclick=()=>{
 if(monitoringWorking||uploadingPhotos||dayTransitioning)return;
 try{const day=displayedDay(),backup=templateBackup(day);if(!backup)throw Error('沒有此工作日的表單備份。');if(!confirm('還原上一份表單備份？目前表單也會先另存備份；不改動照片或母表。'))return;preserveTemplateForm(day,readMonInputs(true));const saved=monDraftManager.saveDraft(day,backup.fields);fillMonInputs(saved.data);updateMonStatusUI();}catch(e){say('還原未完成：'+e.message);}
};
// The offline download and Drive export use the same committed snapshot preparation.
async function prepareMonitoringDay({day=displayedDay(),formData=readMonInputs()}={}){
 const snapshot=await db.snapshot(),byId=new Map(snapshot.photos.map(p=>[p.id,p]));
 const included=snapshot.groups.filter(g=>g.workDay===day||g.photoIds.some(id=>byId.has(id)&&workDayOf(byId.get(id))===day)).sort(compareGroups);
 const prepared=included.length?await preparePhotoGroups(snapshot,included.map(g=>g.id)):[];
 return {day,docBlob:await monitoringDocx(formData,prepared),rowLayout:monitoringRowLayout(prepared)};
}
let monitoringWorking=false;
const standaloneNames=new Map();
// A standalone Google Doc copy of the day's record is ALWAYS kept (it is also the source of the master append).
const defaultStandaloneName=day=>validDay(day)?`施工監看紀錄-${rocDayName(day)}`:'施工監看紀錄';
function refreshStandaloneOptions(){
 const day=displayedDay()||'',input=$('masterStandaloneName');
 if(input.dataset.workDay!==day){input.value=standaloneNames.has(day)?standaloneNames.get(day):defaultStandaloneName(day);input.dataset.workDay=day;}
 const busy=monitoringWorking||uploadingPhotos||dayTransitioning;
 input.disabled=$('browseMonitoringDriveFolders').disabled=$('standaloneWhere').disabled=busy;
 refreshDailyFolderLabels();
}
// Remember where the day's original photos went: the standalone record copy is kept in the same folder.
const dayFolderKey=day=>`field-photo-day-folder-${day}`,dayDocKey=day=>`field-standalone-doc-${day}`;
function readJson(key){try{const raw=safeGetLocalStorage()?.getItem(key);return raw?JSON.parse(raw):null;}catch{return null;}}
function writeJson(key,value){try{safeGetLocalStorage()?.setItem(key,JSON.stringify(value));}catch{}}
function refreshDailyFolderLabels(){
 const day=displayedDay(),name=validDay(day)?rocDayName(day):'';
 if($('dailyFolderName'))$('dailyFolderName').textContent=name?`（${name}）`:'';
 if($('standaloneWhereHint')){const known=readJson(dayFolderKey(day));$('standaloneWhereHint').textContent=known?`今天的照片資料夾：${known.name}`:'今天尚未上傳照片；將改用下方指定的資料夾，或母檔所在資料夾。';}
}
$('masterStandaloneName').addEventListener('input',()=>standaloneNames.set(displayedDay()||'',$('masterStandaloneName').value));
function showDailySource(sourceId){
 const link=document.createElement('a');link.href=`https://docs.google.com/document/d/${encodeURIComponent(sourceId)}/edit`;link.target='_blank';link.rel='noopener noreferrer';link.textContent='開啟當日獨立 Google 文件';$('driveResult').replaceChildren(link);
}
function monitoringControls(){
 const busy=monitoringWorking||uploadingPhotos||dayTransitioning;
 for(const id of [...Object.values(MON_FIELDS),'addMainItem','addChildItem','syncStrata','monRetrySave','monOverwriteConflict','monOverwriteCorrupt'])$(id).disabled=dayTransitioning;
 for(const id of ['viewPhotoCoordinates','exportMonDocx','uploadDrive','uploadPhotosToDrive','browseMonitoringDriveFolders','browseDriveFolders','driveFolder','photoDriveFolder','monitoringAutoDailySubfolder','autoDailySubfolder','connectDrive','switchGoogleAccount','masterDocUrl','masterTab','checkMaster','photoAlsoDoc','fillMonTemplate','restoreMonTemplate','loginMonthMaster','browseMasterDocuments'])$(id).disabled=busy;
 refreshStandaloneOptions();
 $('fillMonTemplate').disabled=busy||!masterReady();$('fillDisabledReason').hidden=masterReady();$('finishMasterSetup').disabled=busy||!masterReady();refreshMasterSelection();
 $('appendMaster').disabled=busy||!masterReady()||!validDay(displayedDay())||displayedDay()!==$('workDay').value;$('openMaster').disabled=busy||!checkedMaster;
}
$('exportMonDocx').onclick=async()=>{
 if(monitoringWorking||uploadingPhotos)return;monitoringWorking=true;monitoringControls();
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
 if(monitoringWorking||uploadingPhotos)return;monitoringWorking=true;monitoringControls();$('driveAccount').hidden=false;
 try{await authorize();$('driveAccount').textContent='已取得 Google Drive 授權（僅本頁有效；過期時須重新授權）。';}
 catch(e){$('driveAccount').textContent='授權失敗：'+e.message;$('driveStatus').textContent='授權失敗：'+e.message;}
 finally{monitoringWorking=false;monitoringControls();}
};

let uploadingPhotos=false;
function readOptionalSetting(key){try{return safeGetLocalStorage()?.getItem(key)||null;}catch{return null;}}
const ALSO_DOC_KEY='field-photo-also-doc';
if(readOptionalSetting(ALSO_DOC_KEY)==='0')$('photoAlsoDoc').checked=false;
$('photoAlsoDoc').addEventListener('change',()=>{try{safeGetLocalStorage()?.setItem(ALSO_DOC_KEY,$('photoAlsoDoc').checked?'1':'0');}catch{}});
const PHOTO_FOLDER_KEY='field-photo-drive-folder-target';
const savedFolder=readOptionalSetting(PHOTO_FOLDER_KEY);
if(savedFolder&&$('photoDriveFolder'))$('photoDriveFolder').value=savedFolder;

$('photoDriveFolder')?.addEventListener('change',()=>{
 try{safeGetLocalStorage()?.setItem(PHOTO_FOLDER_KEY,$('photoDriveFolder').value.trim());}catch{}
});

const MONITORING_FOLDER_KEY='field-monitoring-drive-folder-target';
const savedMonitoringFolder=readOptionalSetting(MONITORING_FOLDER_KEY);
if(savedMonitoringFolder)$('driveFolder').value=savedMonitoringFolder;
$('driveFolder').addEventListener('change',()=>{
 try{safeGetLocalStorage()?.setItem(MONITORING_FOLDER_KEY,$('driveFolder').value.trim());}catch{}
});

// Keep the shared modal outside hidden tab sections, so monitoring can open it too.
document.body.append($('folderBrowserDialog'));
// One browser, with explicit and independent photo/document destinations.
let folderDestination='photos';
function folderTarget(){return folderDestination==='monitoring'
 ? {input:$('driveFolder'),path:$('monitoringDrivePath'),status:$('driveStatus'),key:MONITORING_FOLDER_KEY}
 : {input:$('photoDriveFolder'),path:$('photoDrivePath'),status:$('photoDriveStatus'),key:PHOTO_FOLDER_KEY};}
function savedFolderPath(key,id){
 try{const stored=JSON.parse(safeGetLocalStorage()?.getItem(key+'-path')||'null');return stored?.id===id&&typeof stored.path==='string'?stored.path:'';}catch{return '';}
}
$('photoDrivePath').value=savedFolderPath(PHOTO_FOLDER_KEY,savedFolder)|| (savedFolder?'已儲存的資料夾（請重新選擇）':'');
$('monitoringDrivePath').value=savedFolderPath(MONITORING_FOLDER_KEY,savedMonitoringFolder)|| (savedMonitoringFolder?'已儲存的資料夾（請重新選擇）':'');
let currentBrowsingFolderId=null,currentBrowsingFolderName='';
let browsingShared=false,browsingDriveId=null,browsingDriveList=false,folderHistory=[],folderRequest=0;
async function selectDriveFolder(id,name){
 if(folderDestination==='master')return;
 try{
  if(folderDestination==='monitoring'&&!id){$('folderBrowserStatus').textContent='請先進入或建立一個具體資料夾，再選取。';return;}
  if(id)await verifyFolder(id);
  const target=folderTarget();
  target.input.value=id||'';
  const parts=folderHistory.map(p=>p.name);
  parts.push(currentBrowsingFolderName);
  if(id!==currentBrowsingFolderId)parts.push(name);
  const path=parts.filter(Boolean).join(' / ');
  target.path.value=path;
  try{safeGetLocalStorage()?.setItem(target.key,id||'');safeGetLocalStorage()?.setItem(target.key+'-path',JSON.stringify({id:id||'',path}));}catch{}
  $('folderBrowserDialog').close();
  target.status.textContent=`已選擇：${name}`;
 }catch(e){$('folderBrowserStatus').textContent='無法選取：'+e.message;}
}
async function showDriveFolder(id,name,shared=false,push=false,driveId=null,driveList=false){
 if(push)folderHistory.push({id:currentBrowsingFolderId,name:currentBrowsingFolderName,shared:browsingShared,driveId:browsingDriveId,driveList:browsingDriveList});
 currentBrowsingFolderId=id;browsingShared=shared;browsingDriveId=driveId;browsingDriveList=driveList;
 currentBrowsingFolderName=name;
 $('folderBrowserCurrent').textContent=[...folderHistory.map(p=>p.name),name].filter(Boolean).join(' / ');
 $('folderUp').disabled=!folderHistory.length;
 $('folderSelectCurrent').hidden=$('folderCreateControls').hidden=folderDestination==='master';
 $('folderSelectCurrent').disabled=shared||driveList||folderDestination==='master';
 $('createNewSubfolder').disabled=shared||driveList||folderDestination==='master';
 await refreshFolderList(id);
}
async function refreshFolderList(parentId){
 const request=++folderRequest;const shared=browsingShared,driveId=browsingDriveId,driveList=browsingDriveList,pickingMaster=folderDestination==='master';
 const listEl=$('folderList'),statusEl=$('folderBrowserStatus');
 if(!listEl)return;
 listEl.replaceChildren();
 if(statusEl)statusEl.textContent='正在讀取資料夾清單…';
 try{
  const [folders,documents]=await Promise.all([driveList?listSharedDrives():listSubfolders(parentId,{shared,driveId}),pickingMaster&&!driveList?listMonitoringDocuments(parentId,{driveId}):Promise.resolve([])]);
  if(request!==folderRequest)return;
  if(!folders.length&&!documents.length){
   if(statusEl)statusEl.textContent=driveList?'目前登入帳號沒有可列出的共用雲端硬碟；請確認使用公司帳號且已加入硬碟成員。':shared?'目前帳號沒有可列出的共用資料夾；可貼上已獲授權的資料夾連結。':pickingMaster?'此資料夾內沒有原生 Google 文件或子資料夾；請返回上層，或關閉後貼上母表副本網址。':'此資料夾內目前沒有子資料夾；可選取目前資料夾或返回上層。';
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
   btn.disabled=f.capabilities?.canAddChildren!==true;
   if(btn.disabled)btn.textContent='無新增權限';
   btn.onclick=()=>selectDriveFolder(f.id,f.name);
   const enter=document.createElement('button');
   enter.type='button';enter.textContent='進入';
   enter.onclick=()=>showDriveFolder(f.id,f.name,false,true,driveList?f.id:driveId);
   li.append(nameSpan,enter);
   if(!driveList&&!pickingMaster)li.append(btn);
   listEl.append(li);
  }
  for(const file of documents){
   const li=document.createElement('li'),name=document.createElement('span'),choose=document.createElement('button');li.className='master-document-choice';name.textContent=file.name;choose.type='button';choose.textContent=file.capabilities?.canEdit===true?'選擇母表':'無編輯權限';choose.disabled=file.capabilities?.canEdit!==true;
   choose.onclick=()=>{if(request!==folderRequest||!$('masterSetupDialog').open||!/^[-\w]+$/.test(file.id))return;$('masterDocUrl').value=`https://docs.google.com/document/d/${file.id}/edit`;$('masterDocUrl').dispatchEvent(new Event('input'));$('masterCheckStatus').textContent=`已選擇：${file.name}。請按「檢查母表」。`;$('folderBrowserDialog').close();++folderRequest;};
   li.append(name,choose);listEl.append(li);
  }
  if(statusEl)statusEl.textContent=driveList?`共找到 ${folders.length} 個共用雲端硬碟。按「進入」選擇其中的專案資料夾。`:pickingMaster?`共 ${folders.length} 個子資料夾、${documents.length} 份 Google 文件。只列出原生 Google 文件，請選擇母表副本，再檢查。`:`共找到 ${folders.length} 個資料夾。可進入子資料夾，或選取有新增權限的資料夾。`;
 }catch(e){
  if(request===folderRequest&&statusEl)statusEl.textContent='讀取失敗：'+e.message;
 }
}

$('switchGoogleAccount')?.addEventListener('click',async()=>{
 if(monitoringWorking||uploadingPhotos)return;monitoringWorking=true;monitoringControls();$('driveAccount').hidden=false;$('driveAccount').textContent='正在開啟 Google 帳號選擇視窗…';
 const statusEl=$('photoDriveStatus');
 if(statusEl)statusEl.textContent='正在開啟 Google 帳號選擇視窗…';
 try{
  await authorize(true);
  if(statusEl)statusEl.textContent='已成功切換 Google 帳號！';$('driveAccount').textContent='已切換帳號，請重新檢查母檔與獨立文件資料夾權限。';
 }catch(e){
  if(statusEl)statusEl.textContent='帳號切換未完成：'+e.message;$('driveAccount').textContent='帳號切換未完成：'+e.message;
 }finally{monitoringWorking=false;monitoringControls();}
});

async function openDriveFolderBrowser(destination){
 if(monitoringWorking||uploadingPhotos)return;
 folderDestination=destination;
 const target=destination==='master'?null:folderTarget();
 const inputVal=target?.input.value.trim()||'';
 let folderId=null;
 try{folderId=parseDriveId(inputVal);}catch(e){
  if(target)target.status.textContent=e.message;
  return;
 }
 folderHistory=[];
 const session=++folderRequest;$('folderList').replaceChildren();$('folderBrowserCurrent').textContent='';
 $('folderSelectCurrent').hidden=$('folderCreateControls').hidden=destination==='master';
 $('folderBrowserHeading').textContent=destination==='master'?'選擇本月 Google 母表副本':destination==='monitoring'?'選擇監看紀錄表的 Drive 資料夾':'選擇照片的 Drive 資料夾';
 $('folderBrowserDialog')?.showModal();
 if($('newSubfolderName'))$('newSubfolderName').value=validDay(displayedDay())?rocDayName(displayedDay()):'';
 try{
  await authorize();
  if(session!==folderRequest||!$('folderBrowserDialog').open)return;
  if(folderId){
   const folder=await verifyFolder(folderId);
   if(session!==folderRequest||!$('folderBrowserDialog').open)return;
   await showDriveFolder(folderId,savedFolderPath(target.key,folderId)||folder.name||'目標資料夾');
  }else{
   await showDriveFolder(null,'共用雲端硬碟',false,false,null,true);
  }
 }catch(e){
  if(session===folderRequest&&$('folderBrowserDialog').open)$('folderBrowserStatus').textContent='授權失敗：'+e.message;
 }
}
$('browseMasterDocuments').addEventListener('click',()=>openDriveFolderBrowser('master'));
$('browseDriveFolders')?.addEventListener('click',()=>openDriveFolderBrowser('photos'));
$('browseMonitoringDriveFolders')?.addEventListener('click',()=>openDriveFolderBrowser('monitoring'));

$('folderMyDrive')?.addEventListener('click',()=>{folderHistory=[];showDriveFolder(null,'我的雲端硬碟');});
$('folderShared')?.addEventListener('click',()=>{folderHistory=[];showDriveFolder(null,'與我共用',true);});
$('folderSharedDrives')?.addEventListener('click',()=>{folderHistory=[];showDriveFolder(null,'共用雲端硬碟',false,false,null,true);});
$('folderUp')?.addEventListener('click',()=>{
 const previous=folderHistory.pop();
 if(previous)showDriveFolder(previous.id,previous.name,previous.shared,false,previous.driveId,previous.driveList);
});
$('folderSelectCurrent')?.addEventListener('click',()=>{
 if(folderDestination!=='master'&&!browsingShared&&!browsingDriveList)selectDriveFolder(currentBrowsingFolderId,currentBrowsingFolderName);
});

$('closeFolderBrowser')?.addEventListener('click',()=>{
 $('folderBrowserDialog')?.close();++folderRequest;
});
$('folderBrowserDialog').addEventListener('cancel',()=>{++folderRequest;});

$('createNewSubfolder')?.addEventListener('click',async()=>{
 if(folderDestination==='master')return;
 const nameInput=$('newSubfolderName');
 const name=nameInput?.value.trim();
 if(!name){if($('folderBrowserStatus'))$('folderBrowserStatus').textContent='請輸入新資料夾名稱';return;}
 const btn=$('createNewSubfolder');
 btn.disabled=true;
 try{
  if($('folderBrowserStatus'))$('folderBrowserStatus').textContent=`正在建立「${name}」…`;
  const created=await createFolder(name,currentBrowsingFolderId);
  await selectDriveFolder(created.id,created.name);
 }catch(e){
  if($('folderBrowserStatus'))$('folderBrowserStatus').textContent='建立失敗：'+e.message;
 }finally{
  btn.disabled=false;
 }
});

// Where the standalone record copy goes: today's photo folder by default (then a chosen folder, then the master's folder).
function resolveStandalonePlacement(day,target){
 const name=validateGoogleDocName($('masterStandaloneName').value||defaultStandaloneName(day)),where=$('standaloneWhere').value,known=readJson(dayFolderKey(day));
 let folderId=null,label='';
 if(where==='photos'&&known?.id){folderId=known.id;label=`照片當日資料夾「${known.name}」`;}
 else{
  const chosen=parseDriveId($('driveFolder').value);
  if(chosen){folderId=chosen;label=`指定資料夾「${$('monitoringDrivePath').value||'已選資料夾'}」`;}
  else if(where==='chosen')throw Error('請先選擇獨立文件資料夾，或改用「照片當日資料夾」。');
 }
 if(folderId)return {placement:{custom:true,folderId,name},placementLabel:`獨立文件副本：${name}\n存放於：${label}`};
 return {placement:{custom:false,folderId:target.parent,name},placementLabel:`獨立文件副本：${name}\n存放於：${target.parent?'母檔所在資料夾':'我的雲端硬碟'}（今天尚未上傳照片，也未指定資料夾）`};
}
async function blobFingerprint(blob){return [...new Uint8Array(await crypto.subtle.digest('SHA-256',await blob.arrayBuffer()))].map(n=>n.toString(16).padStart(2,'0')).join('');}
// Uploaded "original" photos keep full size and are only turned upright; the crop is applied only when asked for.
async function preparePhotoForDrive(photo,original,{applyCrop=false}={}){
 const hasExplicitRotation=[90,180,270].includes(photo.displayRotation);
 const needsAutoCameraFix=photo.source==='camera'&&photo.displayRotation===undefined&&photo.mime==='image/jpeg';
 const cropIt=applyCrop&&validCrop(photo.crop);
 if(!hasExplicitRotation&&!needsAutoCameraFix&&!cropIt)return original;
 const url=URL.createObjectURL(original);let canvas;
 try{
  const im=new Image();im.src=url;await im.decode();
  const w=im.naturalWidth,h=im.naturalHeight,r=previewRotation(photo,w,h);
  if((!r||r%360===0)&&!cropIt)return original;
  canvas=renderPhotoCanvas(im,photo,{maxSide:Infinity,applyCrop:cropIt}).canvas;
  const mime=photo.mime==='image/png'?'image/png':photo.mime==='image/webp'?'image/webp':'image/jpeg';
  return (await new Promise(resolve=>canvas.toBlob(resolve,mime,0.92)))||original;
 }finally{
  URL.revokeObjectURL(url);if(canvas)canvas.width=canvas.height=1;
 }
}
// "同時建立獨立文件": the day's monitoring record as a Google Doc next to the photos. Identical content is never duplicated.
async function createStandaloneCopy(day,folderId,statusEl){
 if(statusEl)statusEl.textContent='照片已上傳，正在建立獨立監看紀錄文件…';
 const name=validateGoogleDocName($('masterStandaloneName').value||defaultStandaloneName(day));
 const {docBlob}=await prepareMonitoringDay({day,formData:readMonInputs()}),fingerprint=await blobFingerprint(docBlob),saved=readJson(dayDocKey(day));
 if(saved?.docId&&saved.fingerprint===fingerprint&&saved.folderId===folderId)return `；獨立文件「${saved.name}」內容未變，沿用既有文件。`;
 const finalName=saved?.docId&&saved.name===name?validateGoogleDocName(`${name}（更新 ${new Date().toTimeString().slice(0,5).replace(':','')}）`):name;
 const file=await uploadDayDoc(day,docBlob,folderId,{name:finalName});
 writeJson(dayDocKey(day),{docId:file.id,folderId,name:finalName,fingerprint,at:Date.now()});showDailySource(file.id);
 return `；已建立獨立文件「${finalName}」。`;
}
$('uploadPhotosToDrive')?.addEventListener('click',async()=>{
 if(uploadingPhotos||monitoringWorking)return;
 const statusEl=$('photoDriveStatus');
 let folderId=null;
 try{folderId=parseDriveId($('photoDriveFolder').value);}catch(e){if(statusEl)statusEl.textContent=e.message;return;}
 const batchDay=displayedDay();
 const autoDaily=$('autoDailySubfolder').checked;
 const uploadList=photos.filter(p=>workDayOf(p)===batchDay&&(p.source==='camera'||p.source==='gallery')).map(p=>({photo:p,original:assets.find(a=>a.id===p.id)?.original}));
 if(!uploadList.length){if(statusEl)statusEl.textContent='目前工作日沒有可上傳的照片。';return;}
 if(uploadList.some(item=>!item.original)){if(statusEl)statusEl.textContent='原始照片缺失，未開始上傳。請先檢查本機資料。';return;}
 uploadingPhotos=true;monitoringControls();
 try{
  await authorize();
  let targetFolderId=folderId,targetFolderName=$('photoDrivePath')?.value.split('/').pop().trim()||'照片資料夾';
  if(autoDaily&&targetFolderId){
   if(statusEl)statusEl.textContent=`正在檢查/建立工作日子資料夾 ${rocDayName(batchDay)}…`;
   const dailyFolder=await ensureDailyFolder(targetFolderId,batchDay);
   targetFolderId=dailyFolder.id;targetFolderName=dailyFolder.name;
  }
  let successCount=0;
  for(let i=0;i<uploadList.length;i++){
   const {photo:p,original}=uploadList[i];
   if(statusEl)statusEl.textContent=`正在上傳照片 (${i+1}/${uploadList.length})：${photoLabel(p)}…`;
   const uploadBlob=await preparePhotoForDrive(p,original,{applyCrop:$('uploadCropped')?.checked===true});
   await uploadOriginalPhoto(p,uploadBlob,targetFolderId);
   successCount++;
  }
  if(targetFolderId)writeJson(dayFolderKey(batchDay),{id:targetFolderId,name:targetFolderName,at:Date.now()});
  let docNote='';
  if($('photoAlsoDoc')?.checked&&targetFolderId){
   try{docNote=await createStandaloneCopy(batchDay,targetFolderId,statusEl);}
   catch(e){docNote=`；但獨立文件未建立：${e.message}`;}
  }
  if(statusEl)statusEl.textContent=`成功上傳 ${successCount} 張原始照片至 Google Drive（${targetFolderName}）！本機照片仍保留。${docNote}`;
 }catch(e){
  if(statusEl)statusEl.textContent=`上傳失敗：${e.message}。已成功項目不受影響。`;
 }finally{
  uploadingPhotos=false;monitoringControls();
 }
});

$('uploadDrive').onclick=async()=>{
 if(monitoringWorking||uploadingPhotos)return;
 let folderId;
 try{folderId=parseDriveId($('driveFolder').value);}catch(e){$('driveStatus').textContent=e.message;return;}
 const batch={day:displayedDay(),formData:readMonInputs()};
 const autoDaily=$('monitoringAutoDailySubfolder').checked;
 monitoringWorking=true;monitoringControls();$('driveResult').replaceChildren();
 try{
  // Freeze workday, form and destination before authorization yields.
  await authorize();$('driveAccount').textContent='已取得 Google Drive 授權（僅本頁有效）。';
  let targetFolderId=folderId;
  if(autoDaily&&folderId){
   await verifyFolder(folderId);
   $('driveStatus').textContent=`正在檢查/建立工作日 ${batch.day} 的子資料夾…`;
   targetFolderId=(await ensureDailyFolder(folderId,batch.day)).id;
  }
  $('driveStatus').textContent=`正在生成 ${batch.day} 的監看紀錄表並上傳；請勿重複點擊…`;
  const {day,docBlob}=await prepareMonitoringDay(batch);
  const file=await uploadDayDoc(day,docBlob,targetFolderId);
  const link=document.createElement('a');link.href=file.webViewLink;link.target='_blank';link.rel='noopener noreferrer';link.textContent='Open Google Doc（開啟 Google 文件）';
  $('driveResult').append(link);
  $('driveStatus').textContent=`已上傳 ${file.name||day} 並轉為 Google Doc。轉換可能改變版面，請開啟檢查。此操作不會同步本機資料。`;
 }catch(e){$('driveStatus').textContent=`${e.kind==='auth'?'授權':e.kind==='network'?'網路':e.kind==='api'?'Drive API':'匯出'}失敗：${e.message}`;if(!hasToken())$('driveAccount').textContent='尚未授權或授權已過期；請重新登入。';}
 finally{monitoringWorking=false;monitoringControls();}
};
// Master append is explicit, copy-first, and uses the SAME GIS client as other App exports.
let checkedMaster=null;
if($('masterBuild'))$('masterBuild').textContent=`App v39 / 追加核心：${docsAppendInfo.DOCS_APPEND_BUILD||'舊版，請使用下方更新連結'}`;
$('masterUpdate')?.addEventListener('click',event=>{if(monitoringWorking||uploadingPhotos){event.preventDefault();$('masterStatus').textContent='請先等待目前輸出完成，再更新程式。';}});
const MASTER_TARGET_KEY='field-monitoring-master-target',MASTER_JOURNAL_KEY='field-monitoring-master-journal-v1';
try{$('masterDocUrl').value=safeGetLocalStorage()?.getItem(MASTER_TARGET_KEY)||'';}catch{}
function masterReady(){return !!checkedMaster&&$('masterDocUrl').value.trim()===checkedMaster.input&&checkedMaster.tabs.some(t=>(t.id||'')===$('masterTab').value);}
function refreshMasterSelection(){
 const tab=checkedMaster?.tabs.find(t=>(t.id||'')===$('masterTab').value);
 $('masterSelectionStatus').textContent=masterReady()?`已檢查：${checkedMaster.file.name}${checkedMaster.tabs.length>1?`／${tab.title}`:''}`:checkedMaster?'請先選擇母表分頁，再完成設定。':'請先登入、選擇並檢查本月母表。';
}
$('loginMonthMaster').onclick=async()=>{
 if(monitoringWorking||uploadingPhotos||dayTransitioning)return;
 $('masterSetupDialog').showModal();await $('connectDrive').onclick();
};
$('closeMasterSetup').onclick=()=>$('masterSetupDialog').close();
$('finishMasterSetup').onclick=()=>{if(!masterReady()||monitoringWorking||uploadingPhotos||dayTransitioning)return;$('masterSetupDialog').close();$('fillMonTemplate').focus();};
function masterJournal(){const storage=safeGetLocalStorage();if(!storage)throw Error('無法保存本機操作紀錄，為避免重複追加，暫不寫入母檔。');const raw=storage.getItem(MASTER_JOURNAL_KEY);const data=raw?JSON.parse(raw):{};if(!data||Array.isArray(data)||typeof data!=='object')throw Error('操作紀錄異常，請先核對母檔；不要清除網站資料。');return {storage,data};}
function saveMasterJournal(key,entry){const {storage,data}=masterJournal();data[key]=entry;storage.setItem(MASTER_JOURNAL_KEY,JSON.stringify(data));}
$('masterDocUrl').addEventListener('input',()=>{checkedMaster=null;$('masterCheckStatus').textContent='';monitoringControls();});
function masterTitleNotice(){
 if(!masterReady())return '請先選擇追加分頁。';
 const heading=docsAppendInfo.masterHeadingTemplate?.(checkedMaster.document,$('masterTab').value||null);
 return heading?'將沿用母檔三行標題與格式，不另加日期標題。':'未辨識到母檔三行標題，仍使用原本日期標題；請先確認範本。';
}
$('masterTab').addEventListener('change',()=>{try{$('masterCheckStatus').textContent=masterTitleNotice();}catch(e){$('masterCheckStatus').textContent=e.message;}monitoringControls();});
$('switchGoogleAccount').addEventListener('click',()=>{checkedMaster=null;$('masterCheckStatus').textContent='帳號切換後請重新檢查母表。';monitoringControls();});
$('openMaster').onclick=()=>{if(checkedMaster)window.open(`https://docs.google.com/document/d/${encodeURIComponent(checkedMaster.file.id)}/edit`,'_blank','noopener,noreferrer');};
function acceptMasterCheck(result,input,preferred=null){
 checkedMaster={...result,input};$('masterTab').replaceChildren();
 if(result.tabs.length>1){const option=document.createElement('option');option.value='';option.textContent='請選擇要追加的分頁';$('masterTab').append(option);}
 for(const tab of result.tabs){const option=document.createElement('option');option.value=tab.id||'';option.textContent=tab.title;$('masterTab').append(option);}
 if(preferred&&result.tabs.some(t=>t.id===preferred))$('masterTab').value=preferred;
 $('masterTabLabel').hidden=result.tabs.length<=1;try{safeGetLocalStorage()?.setItem(MASTER_TARGET_KEY,input);}catch{}
}
$('checkMaster').onclick=async()=>{
 if(monitoringWorking||uploadingPhotos)return;
 const input=$('masterDocUrl').value.trim();checkedMaster=null;monitoringWorking=true;monitoringControls();
 try{
  $('masterCheckStatus').textContent='正在唯讀檢查格式、編輯權限與分頁…';
  const result=await inspectMaster(input);acceptMasterCheck(result,input);
  $('masterCheckStatus').textContent=`可追加：${result.file.name}。${masterTitleNotice()}尚未修改任何內容。`;
 }catch(e){checkedMaster=null;$('masterCheckStatus').textContent=`檢查失敗：${e.message}（未修改文件）`;}
 finally{monitoringWorking=false;monitoringControls();}
};
$('appendMaster').onclick=async()=>{
 if(monitoringWorking||uploadingPhotos||dayTransitioning||!masterReady())return;
 if(docsAppendInfo.DOCS_APPEND_BUILD!=='date-order-v2'){$('masterStatus').textContent='追加核心仍是舊版，尚未寫入。請使用「更新 App 程式」連結，不要清除網站資料。';return;}
 if(!navigator.locks?.request){$('masterStatus').textContent='此瀏覽器沒有多分頁防重複鎖，暫不追加；當日 Word 與獨立 Google 文件仍可匯出。';return;}
 const target={id:checkedMaster.file.id,name:checkedMaster.file.name,tabId:$('masterTab').value||null,parent:checkedMaster.file.parents?.[0]||null};
 const batch={day:displayedDay(),formData:readMonInputs()};
 let placement,placementLabel,recordDate;
 try{
  recordDate=docsAppendInfo.parseMonitoringDate(batch.formData.date);
  if(batch.day!==$('workDay').value)throw Error('工作日尚未切換完成，請稍候。');
  if(!validDay(batch.day))throw Error('工作日尚未載入，請稍候再操作。');
  refreshStandaloneOptions();
  ({placement,placementLabel}=resolveStandalonePlacement(batch.day,target));
 }catch(e){$('masterStatus').textContent=`尚未寫入：${e.message}`;return;}
 if(!confirm(`依監看日期 ${recordDate} 插入「${target.name}」。\n${recordDate!==batch.day?`注意：目前工作日為 ${batch.day}，照片仍取該工作日；補登前請確認工作日正確。\n`:''}${placementLabel}\n同日已有紀錄時會再詢問是否取代；其餘原有內容不刪除、不覆寫。是否繼續？`))return;
 let masterWriteAttempted=false;
 monitoringWorking=true;monitoringControls();$('driveResult').replaceChildren();
 try{
  await authorize();
  await navigator.locks.request(`field-master-${target.id}-${target.tabId||''}`,{ifAvailable:true},async lock=>{
   if(!lock)throw Error('另一個分頁正在操作這份母檔，請稍後檢查。');
   $('masterStatus').textContent='正在凍結當日表單與照片快照…';
   const {day,docBlob,rowLayout}=await prepareMonitoringDay(batch);
   const digest=await crypto.subtle.digest('SHA-256',await docBlob.arrayBuffer());
   const fingerprint=[...new Uint8Array(digest)].map(n=>n.toString(16).padStart(2,'0')).join('');
   const key=`${target.id}:${target.tabId||''}:${day}:${fingerprint}`,{data}=masterJournal(),previous=data[key];
   for(const record of Object.values(data))if(record.masterId===target.id&&(record.tabId||null)===target.tabId&&!['verified','replaced','superseded','source-ready','source-uploading'].includes(record.phase))throw Error('此母檔有未確認完成的操作。請先開啟母檔核對；不自動重試，也不要清除網站資料。');
   if(previous&&!['source-ready','superseded'].includes(previous.phase))throw Error(['verified','replaced'].includes(previous.phase)?'相同內容已完成追加，不會重複加入。':'先前上傳結果未確認，請核對 Drive；不自動重新上傳。');
   const fresh=await inspectMaster(target.id);
   let replace=false;
   if(docsAppendInfo.monitoringRecords(fresh.document,target.tabId).some(r=>r.date===recordDate)){
    if(!confirm(`母檔已有 ${recordDate} 的紀錄。\n\n是否取代？\n新紀錄會先加入並讀回核對；確認無誤後，才刪除該日期的舊紀錄。\n舊紀錄上若有人手動加過內容，也會一併被取代。`))throw Error('已取消取代；母檔未修改。');
    replace=true;
   }
   docsAppendInfo.chronologicalPlacement(fresh.document,target.tabId,recordDate,{replace});
   let sourceId=previous?.sourceId;
   if(previous){
    if(typeof sourceId!=='string'||!sourceId)throw Error('既有來源的操作紀錄不完整，請先核對，不重新上傳。');
    showDailySource(sourceId);
    // The source document that was already created keeps its recorded location (which may now resolve differently) and name;
    // a name typed since then is refused rather than silently ignored: it is never renamed or uploaded again.
    if(previous.sourcePlacement){
     const kept=previous.sourcePlacement.name;
     if(kept!==placement.name&&!kept.startsWith(placement.name+'（'))throw Error('已有保留的來源文件，名稱與本次設定不一致。請先開啟核對並恢復原名稱；不移動、改名或重複上傳。');
     placement=previous.sourcePlacement;
    }
   }else if(!sourceId){
    const saved=readJson(dayDocKey(day));
    // The photo upload step may already have created the identical standalone copy: reuse, never duplicate.
    if(saved?.docId&&saved.fingerprint===fingerprint&&placement.custom&&saved.folderId===placement.folderId){sourceId=saved.docId;placement={...placement,name:saved.name};}
    else if(replace)placement={...placement,name:validateGoogleDocName(`${placement.name}（取代版 ${new Date().toTimeString().slice(0,5).replace(':','')}）`)};
   }
   const journalPlacement=previous?previous.sourcePlacement:placement;
   if(sourceId&&!previous){saveMasterJournal(key,{phase:'source-ready',masterId:target.id,tabId:target.tabId,sourceId,day,fingerprint,sourcePlacement:placement});showDailySource(sourceId);}
   if(!sourceId){
    if(placement.custom)await verifyFolder(placement.folderId);
    saveMasterJournal(key,{phase:'source-uploading',masterId:target.id,tabId:target.tabId,day,fingerprint,sourcePlacement:placement});
    $('masterStatus').textContent=`正在建立獨立文件「${placement.name}」…`;
    sourceId=(await uploadDayDoc(day,docBlob,placement.folderId,{name:placement.name})).id;
    showDailySource(sourceId);
    saveMasterJournal(key,{phase:'source-ready',masterId:target.id,tabId:target.tabId,sourceId,day,fingerprint,sourcePlacement:placement});
   }
   showDailySource(sourceId);
   const result=await appendMonitoringRecord({masterId:target.id,tabId:target.tabId,sourceId,day,recordDate,fingerprint,rowLayout,replace,onProgress:phase=>{$('masterStatus').textContent=`${phase}…請勿關閉頁面。`;},onJournal:entry=>{saveMasterJournal(key,{...entry,...(journalPlacement?{sourcePlacement:journalPlacement}:{})});masterWriteAttempted=true;}});
   if(result.replaced){const {storage,data:all}=masterJournal();for(const [k,v] of Object.entries(all))if(k!==key&&v.masterId===target.id&&(v.tabId||null)===target.tabId&&v.recordDate===recordDate&&v.phase==='verified')all[k]={...v,phase:'superseded'};storage.setItem(MASTER_JOURNAL_KEY,JSON.stringify(all));}
   $('masterStatus').textContent=result.replaced?`已取代監看日期 ${recordDate} 的紀錄（工作日 ${day}）：${result.rows} 列、${result.images} 張照片。新紀錄核對後才刪除舊紀錄；請開啟副本核對版面。`:`已追加監看日期 ${recordDate}（工作日 ${day}）：${result.rows} 列、${result.images} 張照片。原有內容已比對保留；請開啟副本核對版面。`;
  });
 }catch(e){$('masterStatus').textContent=`已停止：${e.message}${masterWriteAttempted?' 如已建立新表格，會保留現況供核對，不刪除、不自動重試。':' 母檔尚未開始寫入，不需刪除任何表格。'}`;}
 finally{monitoringWorking=false;monitoringControls();}
};
$('persist').onclick=()=>run(async()=>{const ok=await navigator.storage?.persist?.();say(ok?'瀏覽器允許持續儲存，仍須備份':'未取得持續儲存保證，請備份');});
run(async()=>{await load();const s=await navigator.storage?.estimate?.();$('storage').textContent=s?`使用約 ${(s.usage/1048576).toFixed(1)} MB / 配額約 ${(s.quota/1048576).toFixed(0)} MB（可能變動）`:'無法取得配額';});
window.addEventListener('storage',e=>{
  for(const day of e.key===null ? [...new Set([displayedDay(),...monDraftManager.getUnsavedDays()])] : e.key?.startsWith('field-mon-form-') ? [e.key.replace('field-mon-form-','')] : []){
    if(!validDay(day))continue;
    if(displayedDay()===day){
      if(!monDraftManager.hasUnsaved(day)){
        monDraftManager.clearCache(day);
        updateMonUI();
      }else{
        monDraftManager.checkExternalChange(day);
        updateMonStatusUI();
      }
    }else if(monDraftManager.hasUnsaved(day))monDraftManager.checkExternalChange(day);
  }
});
// Home-screen access: guidance on iOS; a native prompt only when the browser offers it.
let homeScreenPrompt=null,homeScreenInstalled=false;
const homeScreenMode=window.matchMedia('(display-mode: standalone)');
const homeScreenIOS=/iPhone|iPad|iPod/.test(navigator.userAgent)||(navigator.platform==='MacIntel'&&navigator.maxTouchPoints>1);
const homeScreenAndroid=/Android/.test(navigator.userAgent);
const homeScreenEmbedded=/Line\/|\bLINE\b|FBAN|FBAV|Instagram|MicroMessenger/i.test(navigator.userAgent);
function updateHomeScreenUI(){
 const installed=homeScreenInstalled||homeScreenMode.matches||navigator.standalone===true;
 $('homeScreenHelp').hidden=installed;
 $('installHomeScreen').hidden=installed||homeScreenIOS||homeScreenEmbedded||!homeScreenPrompt;
 $('homeScreenInstructions').textContent=homeScreenEmbedded
  ? '請先複製網址，用 iPhone 的 Safari 或 Android 的 Chrome 開啟；不要在 LINE 等內建瀏覽器內加入。'
  :homeScreenIOS
   ? '在 Safari 點「分享」→「加入主畫面」。若有「作為網頁 App 打開」，可開啟後按「加入」。分享按鈕的位置依 iOS 版本不同；網站不能替你按下加入。'
   :homeScreenPrompt
    ? '點下方「安裝至主畫面」，再於瀏覽器的確認視窗選擇安裝。'
    :homeScreenAndroid
     ? '在 Chrome 的 ⋮ 選單選「安裝應用程式」或「加入主畫面」（名稱依版本不同）。沒有安裝按鈕時，也可由這個選單加入捷徑。'
     : '手機請用 Safari／Chrome 開啟此網址，從分享／選單加入主畫面。電腦 Chrome 可使用網址列的安裝圖示或選單中的安裝功能。';
}
window.addEventListener('beforeinstallprompt',event=>{
 event.preventDefault();homeScreenPrompt=event;updateHomeScreenUI();
});
window.addEventListener('appinstalled',()=>{
 homeScreenInstalled=true;homeScreenPrompt=null;updateHomeScreenUI();
 $('homeScreenStatus').textContent='瀏覽器回報已安裝。請從主畫面圖示開啟，並確認原有資料。';
});
homeScreenMode.addEventListener?.('change',updateHomeScreenUI);
$('homeScreenUrl').value=new URL('./',document.baseURI).href;
$('homeScreenHelp').onclick=()=>{updateHomeScreenUI();$('homeScreenStatus').textContent='';$('homeScreenDialog').showModal();};
$('closeHomeScreen').onclick=()=>$('homeScreenDialog').close();
$('homeScreenBackup').onclick=()=>$('backup').click();
$('homeScreenMonitoring').onclick=()=>{$('homeScreenDialog').close();tab('monitoring');$('monitoringOtherExport').open=true;$('exportMonDocx').scrollIntoView({block:'center'});};
$('copyHomeScreenUrl').onclick=async()=>{
 try{
  if(!navigator.clipboard?.writeText)throw Error('clipboard unavailable');
  await navigator.clipboard.writeText($('homeScreenUrl').value);
  $('homeScreenStatus').textContent='網址已複製。';
 }catch{$('homeScreenUrl').focus();$('homeScreenUrl').select();$('homeScreenStatus').textContent='無法自動複製，請長按／選取上方網址手動複製。';}
};
$('installHomeScreen').onclick=async()=>{
 const prompt=homeScreenPrompt;if(!prompt)return;
 homeScreenPrompt=null;$('installHomeScreen').disabled=true;
 try{
  await prompt.prompt();
  const choice=await prompt.userChoice;
  $('homeScreenStatus').textContent=choice?.outcome==='accepted'
   ? '已送出安裝確認；請等待瀏覽器完成，並從主畫面確認圖示與資料。'
   : '未安裝，資料沒有變更；之後仍可從瀏覽器選單加入主畫面。';
 }catch{$('homeScreenStatus').textContent='瀏覽器未能開啟安裝視窗；請從分享／選單加入主畫面。';}
 finally{$('installHomeScreen').disabled=false;updateHomeScreenUI();}
};
updateHomeScreenUI();
if('serviceWorker' in navigator&&isSecureContext)run(async()=>{await navigator.serviceWorker.register('./sw.js',{scope:'./'});await navigator.serviceWorker.ready;const update=()=>{$('offline').textContent=navigator.serviceWorker.controller?'離線殼已受控制；照片只存 IndexedDB，仍需備份。':'離線殼已安裝，重新載入後確認控制。';};update();navigator.serviceWorker.addEventListener('controllerchange',update);});
