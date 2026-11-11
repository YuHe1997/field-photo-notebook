import * as db from './db.js';
import {emptyShanhuaForm,validateShanhuaForm,rocDate,SHANHUA_PROJECT_NAME,SHANHUA_CARRY_FIELDS,previousWorkContent,previousShanhuaForm} from './shanhua-form.js';
import {shanhuaDocx} from './shanhua-docx.js';
import {getShanhuaFile,saveShanhuaFile,removeShanhuaFile} from './shanhua-store.js';
import {PROJECTS,DEFAULT_PROJECT,projectOf,projectInfo,projectKey,daysKey,monitoringDraftKey,backupRecord,restoreProjectRecords,numberIdentity} from './projects.js';
const PROJECT_KEY='field-photo-notebook-project';
let currentProject=(()=>{try{return projectOf({projectId:safeGetLocalStorage()?.getItem(PROJECT_KEY)});}catch{return DEFAULT_PROJECT;}})();
const isNanke=()=>currentProject===DEFAULT_PROJECT;
const projectSnapshot=s=>({...s,photos:s.photos.filter(p=>projectOf(p)===currentProject),groups:s.groups.filter(g=>projectOf(g)===currentProject)});
import {compareGroups} from './ordering.js';
import * as docsAppendInfo from './docs-append.js';
import {authorize,hasToken,parseDriveId,uploadDayDoc,validateGoogleDocName,uploadOriginalPhoto,listSubfolders,listMonitoringDocuments,listSharedDrives,verifyFolder,createFolder,ensureDailyFolder,inspectMaster,appendMonitoringRecord,rocDayName} from './drive.js';
import {preparePhotoGroups,photoTableDocx,tablePlan,monitoringDocx,monitoringRowLayout} from './docx.js';
import {
  MON_FIELDS,
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
import {validCrop,normalizeCrop,rotateCrop,cropSummary,compactCoordinates,compactPhotoLines,photoCameraBearing,directionLabel,readingAt,validateGroup,move,digest,tar,untar,localDay,validDay,workDayOf,numberedPhoto,numberedGroup,cameraAutoRotation,cameraPreviewRotation,cameraFrameTurn,skyEdgeFromTilt,previewRotation,zoomCapability,snapZoom,cropRect,focusCapabilities,mapPreviewTapToSensor,deriveCameraBearing} from './core.js';
import {WorkPhrasePicker,readPhraseMetadata,photoEngineeringPhrases,readSavedScopePhrases,phraseInsertion} from './work-phrases.js';
import {renderPhotoCanvas} from './photo-render.js';
import {openCropDialog} from './crop-ui.js';
import {isHeif,heifToJpeg,jpegName,canDecode,describeFormat} from './heic.js';
import {importFiles} from './photo-import.js';
import {AREA_DATA} from './area-data.js';
import {createAreaMatcher,areaReferenceText} from './area-reference.js';
let areaMatcher;try{areaMatcher=createAreaMatcher(AREA_DATA);}catch{areaMatcher=()=>({status:'unavailable',reason:'dataset-invalid',candidates:[],datasetVersion:AREA_DATA?.version});}
const areaReference=p=>{if(!projectInfo(projectOf(p)).areaReference)return null;try{return areaMatcher(p);}catch{return {status:'unavailable',reason:'dataset-invalid',candidates:[]};}};
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
const dayKey=()=>daysKey(currentProject);
const monKey=day=>monitoringDraftKey(day,currentProject);
function readDays(projectId=currentProject){try{const s=safeGetLocalStorage();const d=JSON.parse(s?.getItem(daysKey(projectId))||'{}');return {active:validDay(d.active)?d.active:null,days:Array.isArray(d.days)?d.days.filter(validDay):[]};}catch{return {active:null,days:[]};}}
let dayState=readDays(),activeDay=dayState.active;
let monDisplayedDay=null,daySwitchGeneration=0,dayTransitioning=false,loadGeneration=0;
let photoPhrasePicker=null,scopePhrasePicker=null;
let coordinateAlbum=null,coordinateAlbumGeneration=0,coordinateReturn=null;
function closePhrasePickers(){photoPhrasePicker?.invalidate();scopePhrasePicker?.invalidate();}
function displayedDay(){return monDisplayedDay||$('monForm')?.dataset?.displayedDay||activeDay;}
const previewCache=new Map(),rotations=new Map();
function rememberDay(day){const next={active:day,days:[...new Set([...dayState.days,day])].sort()};try{safeGetLocalStorage()?.setItem(dayKey(),JSON.stringify(next));}catch{}dayState=next;activeDay=day;}
async function load(){const generation=++loadGeneration;await db.ensurePhotoIdentities();const snapshot=await db.snapshot();if(generation!==loadGeneration)return;({photos,assets,groups}=projectSnapshot(snapshot));groups.sort(compareGroups);photos.sort((a,b)=>photoTime(a)-photoTime(b)||(a.photoNumber??0)-(b.photoNumber??0)||a.id.localeCompare(b.id));if(!activeDay)rememberDay(photos.length?workDayOf(photos.at(-1)):localDay());if(!monDisplayedDay)monDisplayedDay=activeDay;if($('monForm')&&monDisplayedDay)$('monForm').dataset.displayedDay=monDisplayedDay;render();}
const photoTime=p=>p.buttonPressedAt??p.capturedAt??p.importedAt??0;
const photoLabel=p=>p?numberedPhoto(p):'照片未取得';
const groupTitle=ids=>numberedGroup(ids,photos);
const dayPhotos=()=>photos.filter(p=>workDayOf(p)===activeDay);
const groupInDay=g=>g.workDay===activeDay||g.photoIds.some(id=>photos.some(p=>p.id===id&&workDayOf(p)===activeDay));
function releaseImages(pool){urls[pool].forEach(URL.revokeObjectURL);urls[pool]=[];}
function photoImage(id,pool,source={assets,photos}){const a=source.assets.find(a=>a.id===id),p=source.photos.find(p=>p.id===id);if(!a)return text('p','原始影像未取得');const im=document.createElement('img');im.loading='lazy';im.alt=photoLabel(p);const key=`${id}:${p.displayRotation??'auto'}:${validCrop(p.crop)?[p.crop.x,p.crop.y,p.crop.w].join(','):'full'}`;if(previewCache.has(key)){im.src=previewCache.get(key);im.dataset.previewReady='true';im.dataset.rotation=p.displayRotation??rotations.get(id)??0;return im;}im.onerror=()=>{im.dataset.previewReady='true';im.dataset.previewError='true';im.alt=photoLabel(p)+'（原圖無法解碼，未刪除）';$('repairPhotos').hidden=false;};im.src=URL.createObjectURL(a.original);urls[pool].push(im.src);im.onload=()=>{im.onload=null;try{const {canvas:c,rotation:r}=renderPhotoCanvas(im,p,{maxSide:960});rotations.set(id,r);const result=c.toDataURL('image/jpeg',.9);if(previewCache.size>=60)previewCache.delete(previewCache.keys().next().value);previewCache.set(key,result);if(im.isConnected)im.src=result;im.dataset.rotation=r;im.dataset.previewReady='true';}catch{im.dataset.previewReady='true';say('此照片無法產生旋轉預覽；原圖仍保留。');}};return im;}
function photoAreaInfo(p,{collapsed=false}={}){
  if(!projectInfo(projectOf(p)).areaReference)return document.createDocumentFragment();
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
      details.append(text('span',`裝置羅盤原始讀值：${p.heading.degrees.toFixed(1)}° `));
    }
    const bearing=photoCameraBearing(p);
    if(bearing.degrees!==null){
      details.append(text('span',`鏡頭朝向估計：${bearing.label}（${bearing.degrees.toFixed(1)}°，${bearing.method==='recorded'?(p.cameraBearingEvidence||'推估'):bearing.method}；需驗證）`));
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
  leaveCameraFullscreen();lastLayoutKey='';frameKey='';previewRot=0;skyEdge=null;
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
// Shown in 資訊: lets a real-phone test tell which way the platform delivers the camera picture.
function orientationDiagnostics(){
  const v=$('video'),sky=currentSky(),edge={0:'上緣',90:'右緣',180:'下緣',270:'左緣'};
  if(!v?.videoWidth||!v.videoHeight)return '相機影格未就緒';
  return `影格 ${v.videoWidth}×${v.videoHeight}；視窗${windowPortrait()?'直向':'橫向'}；朝天：${sky===null?'未取得':edge[sky]}；存檔轉 ${uprightRotation()}°；預覽轉 ${previewRot}°；畫面轉 ${frameTurnNow()}°`;
}
function sensors(){
  const now=Date.now(),fresh=heading?now-heading.readingAt<=2000:false;
  const bearing=photoCameraBearing({heading,cameraBearing:heading?.cameraBearing});
  let bearingText;
  if(!fresh){
    bearingText='無法可靠判定（讀值已過期，無即時姿態）';
  }else if(bearing.degrees!==null){
    bearingText=`估計 ${bearing.label}（${bearing.degrees.toFixed(1)}°，需驗證）`;
  }else{
    bearingText=`無法可靠判定${heading?.cameraBearingEvidence?`（${heading.cameraBearingEvidence}）`:''}`;
  }
  $('sensor').textContent=`${sensorState}；GPS 狀態：${gpsStatus}；方向狀態：${directionStatus}；GPS：${gps?`${gps.latitude.toFixed(6)}, ${gps.longitude.toFixed(6)} ±${gps.accuracyM}m；${now-gps.readingAt<=10000?'新鮮':'過期'}；${gps.accuracyM>20?'約略':'精度非保證'}`:'未取得'}；裝置羅盤：${heading?`裝置參考推估：${bearing.label}；${bearing.degrees?.toFixed(1)}°；${heading.quality}；${fresh?'新鮮':'過期'}（僅供參考）`:'未取得絕對方位'}；鏡頭朝向：${bearingText}；轉正診斷：${orientationDiagnostics()}`;
  if($('cameraOverlayStatus')){
    $('cameraOverlayStatus').classList.toggle('is-notice',Boolean(overlayErrorText));
    if(overlayErrorText){
      $('cameraOverlayStatus').textContent=overlayErrorText;
    }else if(starting){
      $('cameraOverlayStatus').textContent='正在啟動相機…';
    }else if(!fresh){
      $('cameraOverlayStatus').textContent=heading?`方向：${bearing.label}（讀值過期，需驗證）`:'等待方位';
    }else if(bearing.degrees!==null){
      $('cameraOverlayStatus').textContent=`鏡頭估計：${bearing.label}（${bearing.degrees.toFixed(1)}°，需驗證）`;
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
let previewRot=0,lastLayoutKey='',frameKey='',skyEdge=null;
// Which phone edge points at the sky; null while there is no fresh sensor reading.
function currentSky(){return deviceTilt&&Date.now()-deviceTilt.readingAt<=2000&&skyEdge!==null?skyEdge:null;}
// Like a camera with buttons, the camera screen stays fixed to the phone. When the system rotates the screen along with the
// phone (landscape window) the whole camera frame is turned back, so every control stays exactly where it was in portrait.
function windowPortrait(){const overlay=$('cameraOverlay');return !overlay||overlay.clientHeight>=overlay.clientWidth;}
function wantedFrameTurn(){
  if(!$('cameraOverlay'))return 0;
  return cameraFrameTurn({skyEdge:currentSky(),screenAngle:screen.orientation?.angle??window.orientation??null,windowPortrait:windowPortrait()});
}
const frameTurnNow=()=>Number($('cameraFrame')?.dataset.turn)||0;
// Degrees clockwise that turn the live picture upright (sky at the top); the shutter saves the photo with this turn.
function uprightRotation(screenAngle=screen.orientation?.angle??window.orientation??null,tilt=null){
  const v=$('video');if(!v?.videoWidth||!v.videoHeight)return 0;
  return cameraAutoRotation({width:v.videoWidth,height:v.videoHeight,screenAngle,beta:tilt?.beta,gamma:tilt?.gamma,skyEdge:currentSky(),windowPortrait:windowPortrait()});
}
// The camera screen is fixed to the phone, but the picture inside it is what the lens sees, the right way up in the world,
// exactly like the saved photo. iPhone already levels the camera picture by itself (it turns wide when the phone is held
// sideways, even with portrait lock); this works out the turn either way from the saved-photo rotation and the sky edge.
function desiredPreviewRotation(){
  const v=$('video');if(!v?.videoWidth||!v.videoHeight)return 0;
  const sky=currentSky();
  if(sky===null){const turn=frameTurnNow();return v.videoWidth>v.videoHeight&&turn?(360-turn)%360:0;}
  return cameraPreviewRotation({captureRotation:uprightRotation(),skyEdge:sky});
}
function updateCameraFrame(){
  const overlay=$('cameraOverlay'),frame=$('cameraFrame');if(!overlay||!frame||overlay.hidden)return;
  const W=overlay.clientWidth,H=overlay.clientHeight;if(!W||!H)return;
  const turn=wantedFrameTurn(),fw=turn?H:W,fh=turn?W:H;
  const key=[W,H,turn].join('x');if(key===frameKey)return;frameKey=key;
  frame.classList.toggle('is-turned',turn!==0);frame.dataset.turn=String(turn);
  frame.style.setProperty('--fw',fw+'px');frame.style.setProperty('--fh',fh+'px');frame.style.setProperty('--fr',turn+'deg');
}
function previewDimensions(){
  updateCameraFrame();
  const v=$('video'),preview=$('cameraPreview'),vp=$('cameraViewport'),rotor=$('cameraRotor');
  if(!v||!preview||!vp)return;
  const vw=v.videoWidth,vh=v.videoHeight;
  if(!vw||!vh)return;
  const rot=desiredPreviewRotation(),turned=rot%180!==0,aspect=turned?vh/vw:vw/vh;
  // Layout size, not the on-screen box: the camera screen itself may be turned.
  const stage={width:preview.offsetWidth,height:preview.offsetHeight};
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
    frameKey='';updateCameraFrame();
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
  // The tap is on the screen; the picture lives in the (possibly turned) camera screen: bring the tap into that frame.
  const turnNow=frameTurnNow(),theta=turnNow*Math.PI/180,dx=e.clientX-(vpRect.left+vpRect.width/2),dy=e.clientY-(vpRect.top+vpRect.height/2);
  const vpW=viewport.offsetWidth,vpH=viewport.offsetHeight,tapX=vpW/2+dx*Math.cos(theta)+dy*Math.sin(theta),tapY=vpH/2-dx*Math.sin(theta)+dy*Math.cos(theta);
  const inVp=tapX>=0&&tapX<=vpW&&tapY>=0&&tapY<=vpH;
  if(!inVp){if(status)status.textContent='點按位置在黑邊（letterbox）區域，請點按鏡頭影像區域';return;}
  const z=zoomState;
  const turned=previewRot%180!==0,mapping=mapPreviewTapToSensor(tapX,tapY,vpW,vpH,turned?video.videoHeight:video.videoWidth,turned?video.videoWidth:video.videoHeight,z?.mode,z?.factor);
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
    skyEdge=skyEdgeFromTilt({beta:e.beta,gamma:e.gamma,previous:skyEdge});
    if(wantedFrameTurn()!==frameTurnNow()||desiredPreviewRotation()!==previewRot)previewDimensions();
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
  if(source==='webkitCompassHeading'&&facing==='environment'){
    const bearing=photoCameraBearing({heading:{degrees,source,beta:e.beta,gamma:e.gamma}});
    const sky=skyEdgeFromTilt({beta:e.beta,gamma:e.gamma});
    Object.assign(derived,{degrees:bearing.degrees,label:bearing.label,algorithm:bearing.method,evidence:`webkitCompassHeading＋朝天邊 ${sky??'未知'}°${sky===null?'（平放／姿態缺值，沿用羅盤）':''}；需驗證`});
  }
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
    say('已儲存於本機（未上傳）'+(lastSavedPhoto?.source==='camera'&&projectInfo(projectOf(lastSavedPhoto)).areaReference?'；'+areaReferenceText(areaReference(lastSavedPhoto)):''));
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
    const captureRotation=uprightRotation(screenAngle,tilt);
    // Always save the live picture itself (what the viewfinder shows). ImageCapture.takePhoto is not used: on iPhone it
    // returns a picture fixed to the phone, so a photo taken with the phone sideways came out turned with the phone.
    {
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
      cameraFacing:currentFacing,cameraBearing:finalBearing,cameraBearingLabel:finalLabel,cameraBearingAlgorithm:finalAlg,cameraBearingEvidence:finalEvidence,cameraTransformEvidence:finalEvidence??'未驗證裝置／螢幕／鏡頭轉換',workDay,projectId:currentProject,screenAngle,zoom,outputWidth,outputHeight,displayRotation:0}});
    await commitPending();
  }finally{busy=false;zoomUI();}
});
$('import').onchange=()=>run(async()=>{
  if(pending.length)throw Error('請先點「重試未完成照片儲存」，再選取照片。');
  const workDay=activeDay,projectId=currentProject,files=[...$('import').files];
  if(!files.length)return;
  busy=true;
  try{
    const result=await importFiles(files,{workDay,isHeif,canDecode,heifToJpeg,describeFormat,jpegName,
      newId:()=>crypto.randomUUID(),now:()=>Date.now(),progress:say,
      save:async item=>{
        item.photo.projectId=projectId;
        pending.push(item);
        updateRetryUI();
        await db.savePhoto(item.photo,item.blob);
        pending.shift();
        updateRetryUI();
      }});
    await load();
    say(result.message);
  }finally{busy=false;$('import').value='';updateRetryUI();}
});
// Photos stored before v44 (or mislabeled HEIC) cannot be shown; offer to convert them in place.
$('repairPhotos').onclick=()=>run(async()=>{const bad=[];for(const p of photos){const a=assets.find(a=>a.id===p.id);if(a&&!await canDecode(a.original))bad.push([p,a.original]);}
  if(!bad.length){$('repairPhotos').hidden=true;say('所有照片都能正常顯示');return;}
  if(!confirm(`有 ${bad.length} 張照片無法顯示。要把其中的 HEIC／HEIF 轉成 JPEG 嗎？（轉換後本機改存 JPEG，不保留原檔 EXIF）`))return;
  let fixed=0;const failed=[];for(const [i,[p,blob]] of bad.entries()){say(`正在修復照片（${i+1}／${bad.length}）…`);try{const {blob:jpeg,decoder}=await heifToJpeg(blob);await db.replaceOriginal(p.id,jpeg,{originalName:jpegName(p.originalName),mime:'image/jpeg',convertedFrom:{name:p.originalName,mime:p.mime,size:blob.size,decoder,repairedAt:Date.now()},displayRotation:0,captureMethod:`${p.captureMethod??''}；已修復：HEIC／HEIF 轉為 JPEG（${decoder}）`});for(const k of [...previewCache.keys()])if(k.startsWith(p.id+':'))previewCache.delete(k);fixed++;}catch(e){failed.push(`${photoLabel(p)}（${p.mime||'未標類型'}，${await describeFormat(blob)}）`);}}
  await load();$('repairPhotos').hidden=!failed.length;say(`已修復 ${fixed} 張照片`+(failed.length?`；${failed.length} 張仍無法解碼：${failed.join('、')}`:''));});
function choose(id){if(multi){selected.has(id)?selected.delete(id):selected.add(id);}else selected=selected.has(id)?new Set():new Set([id]);render();}
function render(){renderProjectUI();$('editingTab').hidden=true;$('appendToEditor').hidden=!edit;releaseImages('gallery');releaseImages('groups');$('gallery').replaceChildren();const days=[...new Set([...dayState.days,...photos.map(workDayOf),activeDay])].filter(validDay).sort();$('workDay').replaceChildren();for(const d of days){const o=text('option',`${d}（${photos.filter(p=>workDayOf(p)===d).length} 張）`);o.value=d;$('workDay').append(o);}$('workDay').value=activeDay;const assignedIds=new Set(groups.flatMap(g=>g.photoIds));$('currentDay').replaceChildren(...[activeDay,`拍攝 ${dayPhotos().filter(p=>p.source==='camera').length} 張`,`相簿 ${dayPhotos().filter(p=>p.source==='gallery').length} 張`,`待整理 ${dayPhotos().filter(p=>!assignedIds.has(p.id)).length} 張`,`紀錄集 ${groups.filter(groupInDay).length} 組`].map(label=>text('span',label)));for(const id of selected)if(!dayPhotos().some(p=>p.id===id))selected.delete(id);const assigned=new Set(groups.flatMap(g=>g.photoIds));for(const p of dayPhotos().filter(p=>!assigned.has(p.id)||selected.has(p.id))){const card=text('div','');card.className='card';const b=text('button',`${selected.has(p.id)?'已選 · ':''}${photoLabel(p)} · ${p.source==='camera'?'拍攝':'相簿'}${validCrop(p.crop)?' · 已裁切':''}`);b.setAttribute('aria-pressed',selected.has(p.id));b.prepend(photoImage(p.id,'gallery'));let timer,long=false;b.onpointerdown=e=>{if(e.button!==0)return;long=false;timer=setTimeout(()=>{long=true;multi=true;selected.add(p.id);render();},550);};b.onpointerup=b.onpointerleave=b.onpointercancel=()=>clearTimeout(timer);b.onclick=()=>{if(long){long=false;return;}choose(p.id);};const coordinates=text('p',compactCoordinates(p.location));coordinates.className='gallery-coordinates';card.append(b,photoAreaInfo(p,{collapsed:true}),coordinates,text('p',compactPhotoLines(p,0)[3]));$('gallery').append(card);}$('multi').textContent=multi?'結束多選':'選取多張';$('multi').setAttribute('aria-pressed',multi);$('selectionCount').textContent=`已選 ${selected.size} 張；本日待整理 ${dayPhotos().filter(p=>!assigned.has(p.id)).length} 張`;$('appendToEditor').disabled=$('deleteSelected').disabled=$('rotateLeft').disabled=$('rotateRight').disabled=$('compose').disabled=!selected.size;$('cropSelected').disabled=selected.size!==1;$('returnToEditor').hidden=!edit;$('compose').hidden=!!edit;$('groupList').replaceChildren();if(!groups.some(groupInDay))$('groupList').append(text('p','本工作日尚無紀錄集。請到待整理照片選取照片並記事。'));const ordered=groups.filter(groupInDay).sort(compareGroups);$('toggleGroupOrder').disabled=ordered.length<2;$('groupOrderHint').textContent=ordered.length<2?'至少需要 2 組才能調整順序':'即時儲存 · 所有參考與匯出沿用此順序';for(const [index,g] of ordered.entries()){const a=text('article','');a.dataset.groupId=g.id;const orderTools=text('div','');orderTools.className='group-order-controls';orderTools.hidden=!groupOrdering;for(const [label,delta] of [['↑ 上移',-1],['↓ 下移',1]]){const b=text('button',label);b.dataset.direction=String(delta);b.disabled=index+delta<0||index+delta>=ordered.length;b.setAttribute('aria-label',`${groupTitle(g.photoIds)} ${label}`);b.onclick=()=>run(async()=>{await db.saveGroupOrder(activeDay,move(ordered.map(x=>x.id),index,delta),currentProject);await load();say('紀錄集順序已儲存');});orderTools.append(b);}a.append(orderTools,text('h3',`${index+1}. ${groupTitle(g.photoIds)}`),text('p',`${g.photoIds.length} 張 · 更新 ${new Date(g.updatedAt).toLocaleString()}`));const strip=text('div','');strip.className='group-photos';g.photoIds.forEach((id,i)=>{const p=photos.find(p=>p.id===id),figure=text('figure','');figure.append(photoImage(id,'groups'),text('figcaption',compactPhotoLines(p,i)[0]));strip.append(figure);});a.append(strip);const info=text('details','');info.append(text('summary',projectInfo(currentProject).areaReference?'照片資訊與區域判讀':'照片資訊'));g.photoIds.forEach((id,i)=>{const row=text('div','');row.append(text('strong',compactPhotoLines(photos.find(p=>p.id===id),i)[0]),photoInfo(photos.find(p=>p.id===id)));info.append(row);});a.append(info);for(const [label,value] of [['工程描述',g.engineering]]){const note=text('p',`${label}：${value||'未填寫'}`);note.className='note-text';a.append(note);}const e=text('button','編輯／排序照片');e.onclick=()=>openEditor(g);const d=text('button','移除紀錄集（保留原始照片）');d.className='danger';d.onclick=()=>run(async()=>{if(edit?.id===g.id){say('請先關閉或儲存此紀錄集的編輯，再移除。');return;}if(confirm('只移除此紀錄集，原始照片仍保留？')){await db.removeGroup(g.id);await load();say('已移除紀錄集，照片仍保留');}});a.append(e,d);$('groupList').append(a);}if(edit&&active==='editor')renderOrder();}
$('multi').onclick=()=>{multi=!multi;if(!multi)selected=new Set([...selected].slice(0,1));render();};
let groupOrdering=false,editingExisting=false;
$('toggleGroupOrder').onclick=()=>{groupOrdering=!groupOrdering;$('toggleGroupOrder').setAttribute('aria-pressed',String(groupOrdering));render();};
$('returnToEditor').onclick=()=>tab('editor');
function openEditor(g,ids=[...selected]){if(!canDiscardEditor())return;editingExisting=!!g;edit=structuredClone(g||{id:crypto.randomUUID(),title:'',notes:'',engineering:'',photoIds:ids,createdAt:Date.now(),workDay:activeDay,projectId:currentProject});$('title').value=groupTitle(edit.photoIds);$('engineering').value=edit.engineering;editBaseline=editorValue();tab('editor');}
function renderOrder(focus){$('title').value=groupTitle(edit.photoIds);releaseImages('order');$('order').replaceChildren();$('photoInfo').replaceChildren();const previews=text('div','');previews.className='editor-previews';previews.dataset.count=edit.photoIds.length;edit.photoIds.forEach((id,i)=>{const figure=text('figure','');figure.append(photoImage(id,'order'),text('figcaption',compactPhotoLines(photos.find(p=>p.id===id),i)[0]));previews.append(figure);});$('order').append(previews);edit.photoIds.forEach((id,i)=>{const p=photos.find(p=>p.id===id),row=text('article',''),head=text('div',''),controls=text('div','');head.className='photo-row-head';controls.className='photo-row-controls';head.append(text('h3',compactPhotoLines(p,i)[0]));for(const [label,delta] of [['上移',-1],['下移',1]]){const b=text('button',delta<0?'↑':'↓');b.title=label;b.setAttribute('aria-label',`${i+1}. ${photoLabel(p)} ${label}`);b.dataset.photoId=id;b.dataset.action=label;b.disabled=i+delta<0||i+delta>=edit.photoIds.length;b.onclick=()=>{edit.photoIds=move(edit.photoIds,i,delta);renderOrder({id,action:label});say(`${photoLabel(p)} 已${label}至第 ${edit.photoIds.indexOf(id)+1} 張`);};controls.append(b);}const cr=text('button','裁切');cr.title='裁切照片（保持原比例）';cr.setAttribute('aria-label',`${i+1}. ${photoLabel(p)} 裁切`);cr.dataset.photoId=id;cr.dataset.action='裁切';cr.onclick=()=>run(()=>cropPhoto(id));controls.append(cr);const r=text('button','×');r.title='從此組移除';r.setAttribute('aria-label',`${i+1}. ${photoLabel(p)} 從此組移除`);r.dataset.photoId=id;r.dataset.action='從此組移除';r.onclick=()=>{edit.photoIds.splice(i,1);renderOrder({id:edit.photoIds[Math.min(i,edit.photoIds.length-1)],action:'從此組移除'});say(`${photoLabel(p)} 已從此組移除，原始照片仍保留`);};controls.append(r);head.append(controls);row.append(head,photoInfo(p));$('photoInfo').append(row);});if(focus){const controls=[...$('photoInfo').querySelectorAll('button')].filter(b=>b.dataset.photoId===focus.id&&!b.disabled);(controls.find(b=>b.dataset.action===focus.action)||controls[0]||$('addSelected')).focus();}}
$('compose').onclick=()=>{if(!selected.size){say('請先選照片');return;}openEditor();};$('appendToEditor').onclick=()=>{if(!edit)return;const day=edit.workDay??workDayOf(photos.find(p=>p.id===edit.photoIds[0])??{});if([...selected].some(id=>workDayOf(photos.find(p=>p.id===id))!==day)){say('請勿將不同工作日的照片加入同一記事。');return;}edit.photoIds=[...new Set([...edit.photoIds,...selected])];tab('editor');};$('addSelected').onclick=()=>{if(edit){tab('photos');say('請選取照片，再按「加入目前記事並返回」。');}};$('cancel').onclick=()=>{if(!canDiscardEditor())return;const returnTab=groups.some(g=>g.id===edit?.id)?'groups':'photos';$('editor').hidden=true;edit=null;editBaseline=null;releaseImages('order');render();tab(returnTab);};$('save').onclick=()=>run(async()=>{if(!edit)return;const title=groupTitle(edit.photoIds),g={...edit,title,notes:edit.notes??'',engineering:$('engineering').value,updatedAt:Date.now()};if(edit.title&&edit.title!==title&&!edit.legacyTitle)g.legacyTitle=edit.title;validateGroup(g,new Set(photos.map(p=>p.id)));await db.saveGroup(g,{existing:editingExisting});$('editor').hidden=true;edit=null;editBaseline=null;releaseImages('order');selected.clear();await load();say('紀錄集已儲存');tab('groups');});
$('backup').onclick=()=>run(async()=>{say('正在建立本機備份（交易快照；不含未提交資料）…');const snapshot=await db.snapshot();const entries=[],records=[];for(const p of snapshot.photos){const blob=snapshot.assets.find(a=>a.id===p.id)?.original;if(!blob)throw Error('原始影像缺失');const path=`originals/${p.id}.bin`;entries.push([path,blob]);records.push({...backupRecord(p),path,size:blob.size,sha256:await digest(blob)});}for(const g of snapshot.groups)validateGroup(g,new Set(snapshot.photos.map(p=>p.id)));const manifest={schema:1,createdAt:Date.now(),snapshotStartedAt:snapshot.startedAt,snapshotCompletedAt:snapshot.completedAt,completeness:'單一唯讀交易中的已提交資料；不含未儲存表單／待存照片／快照之後提交資料',photos:records,groups:snapshot.groups.map(backupRecord),workDays:readDays(DEFAULT_PROJECT).days,activeWorkDay:readDays(DEFAULT_PROJECT).active??localDay(),numberCounters:db.numberCounters(),activeProject:currentProject,projectStates:Object.fromEntries(PROJECTS.map(p=>[p.id,{...readDays(p.id),numberCounters:db.numberCounters(p.id)}]))};entries.push(['area-reference.json',new Blob([JSON.stringify({schema:1,calculatedAt:Date.now(),meaning:'暫定區界及本次備份時計算的參考；還原照片後仍依 App 當前區界回算，不自動匯入或信任此區界',dataset:AREA_DATA,photos:snapshot.photos.map(p=>({photoId:p.id,reference:areaReference(p)}))},null,2)],{type:'application/json'})]);entries.push(['manifest.json',new Blob([JSON.stringify(manifest,null,2)],{type:'application/json'})],['notes.txt',new Blob([snapshot.groups.map(g=>`${g.title}\n工作日：${g.workDay??'舊版請參照各照片 workDay'}\n工程描述：${g.engineering}\n${g.notes?'舊版備註（保留）：'+g.notes+'\n':''}${g.legacyTitle?'舊版名稱（保留）：'+g.legacyTitle+'\n':''}順序：${g.photoIds.join(', ')}\n`).join('\n')])],['README.txt',new Blob(['原始影像在 originals/；檔名與 MIME 在 manifest.json。GPS 是裝置位置；方向不是可靠鏡頭朝向。相機讀值以按鈕時間為基準，不是曝光資料。舊版 capturedAt / ageAtShutterMs 亦是按鈕時間／年齡。沒有 EXIF 嵌入或雲端同步。'])]);const archive=await tar(entries),url=URL.createObjectURL(archive),a=document.createElement('a');a.href=url;a.download=`field-notebook-${Date.now()}.tar`;a.click();setTimeout(()=>URL.revokeObjectURL(url),60000);say('已交給瀏覽器下載；請確認下載完成（無法驗證使用者是否取消）');});
$('restore').onchange=()=>run(async()=>{const file=$('restore').files[0];if(!file)return;say('驗證備份中…');const entries=await untar(file),m=JSON.parse(await entries.get('manifest.json')?.text()),local=await db.snapshot();if(m.schema!==1||!Array.isArray(m.photos)||!Array.isArray(m.groups))throw Error('不支援的備份');Object.assign(m,restoreProjectRecords(m));if(m.projectStates!==undefined){if(!m.projectStates||typeof m.projectStates!=='object'||Array.isArray(m.projectStates))throw Error('專案狀態格式無效');for(const [id,s] of Object.entries(m.projectStates)){if(!PROJECTS.some(p=>p.id===id)||!s||!Array.isArray(s.days)||s.days.some(d=>!validDay(d))||(s.active!==null&&!validDay(s.active))||!s.numberCounters||typeof s.numberCounters!=='object'||Array.isArray(s.numberCounters)||Object.entries(s.numberCounters).some(([d,n])=>!validDay(d)||!Number.isSafeInteger(n)||n<1))throw Error('專案工作日／計數格式無效');}}if(m.workDays!==undefined&&(!Array.isArray(m.workDays)||m.workDays.some(d=>!validDay(d))))throw Error('工作日格式無效');if(m.activeWorkDay!==undefined&&!validDay(m.activeWorkDay))throw Error('目前工作日格式無效');if(m.numberCounters!==undefined&&(typeof m.numberCounters!=='object'||m.numberCounters===null||Array.isArray(m.numberCounters)||Object.entries(m.numberCounters).some(([d,n])=>!validDay(d)||!Number.isSafeInteger(n)||n<1)))throw Error('編號計數格式無效');const incomingNumbers=new Set();const ids=new Set(),gids=new Set(),staged=[];for(const p of m.photos){if(p.workDay!==undefined&&!validDay(p.workDay)||p.photoNumber!==undefined&&(!Number.isSafeInteger(p.photoNumber)||p.photoNumber<1)||p.displayRotation!==undefined&&![0,90,180,270].includes(p.displayRotation))throw Error('照片工作日／編號／旋轉格式無效');if(p.crop!==undefined&&!validCrop(p.crop))throw Error('照片裁切資料格式無效');if(p.workDay&&p.photoNumber){const numberKey=numberIdentity(p);if(incomingNumbers.has(numberKey)||local.photos.some(x=>numberIdentity(x)===numberKey))throw Error('工作日照片編號重複或與本機衝突');incomingNumbers.add(numberKey);}if(typeof p.id!=='string'||!/^[-a-zA-Z0-9]+$/.test(p.id)||ids.has(p.id)||local.photos.some(x=>x.id===p.id)||p.path!==`originals/${p.id}.bin`||typeof p.originalName!=='string'||!['camera','gallery'].includes(p.source)||typeof p.mime!=='string'||!p.mime.startsWith('image/'))throw Error('照片格式／ID 重複或與本機衝突');ids.add(p.id);const b=entries.get(p.path);if(!b||b.size!==p.size||await digest(b)!==p.sha256)throw Error('影像大小或雜湊不符');const {path,size,sha256,...photo}=p;staged.push({photo,blob:new Blob([b],{type:p.mime})});}for(const g of m.groups){if(typeof g.id!=='string'||typeof g.title!=='string'||typeof g.notes!=='string'||typeof g.engineering!=='string'||gids.has(g.id)||local.groups.some(x=>x.id===g.id))throw Error('紀錄集格式／ID 衝突');gids.add(g.id);validateGroup(g,ids);}if(!confirm(`已驗證 ${staged.length} 張照片、${m.groups.length} 組，匯入本機？`))return;await db.transaction(['photos','assets','groups'],tx=>{for(const x of staged){tx.objectStore('photos').add(x.photo);tx.objectStore('assets').add({id:x.photo.id,original:x.blob});}for(const g of m.groups)tx.objectStore('groups').add(g);});db.mergeCounters(m.numberCounters);const states=m.projectStates??{[DEFAULT_PROJECT]:{days:m.workDays??[],active:m.activeWorkDay??null,numberCounters:m.numberCounters??{}}};for(const [id,s] of Object.entries(states)){db.mergeCounters(s.numberCounters,id);const old=readDays(id);safeGetLocalStorage()?.setItem(daysKey(id),JSON.stringify({days:[...new Set([...old.days,...s.days])].sort(),active:s.active??old.active}));}dayState=readDays();activeDay=dayState.active;monDisplayedDay=null;await load();say('備份已完整還原（原圖與顯示旋轉分開保存）');$('restore').value='';});
const SHANHUA_ID='shanhua-1350pscp',SHANHUA_REMEMBER_KEY=projectKey('field-shanhua-remembered',SHANHUA_ID);
const shanhuaRememberFields=SHANHUA_CARRY_FIELDS;
let shanhuaCarried=[];
let shanhuaDay=null,shanhuaWorking=false,shanhuaReady=false,shanhuaGeneration=0,shanhuaDirty=false;
const shanhuaUrls=new Map();
function buildShanhuaUI(){
 const box=text('div','');box.id='shanhuaMonitoring';box.className='shanhua-form';
 box.append(text('h2','善化監看紀錄表'));
 const form=document.createElement('form');form.id='shanhuaForm';form.noValidate=true;form.onsubmit=e=>e.preventDefault();
 function tag(parent,key){const badge=text('span','沿用');badge.className='carried-tag';badge.dataset.carried=key;badge.hidden=true;parent.append(badge);}
 function card(number,title){const section=text('section','');section.className='step-card';const heading=text('h3','');const no=text('span',number);no.className='step-no';heading.append(no,document.createTextNode('　'+title));section.append(heading);box.append(section);return section;}
 function group(parent,title,key){const set=document.createElement('fieldset'),legend=text('legend',title);if(key)tag(legend,key);set.append(legend);parent.append(set);return set;}
 function field(parent,path,label,{type='text',badge=true,value=null}={}){
  const wrap=text('label',''),caption=text('span',label),input=document.createElement(type==='textarea'?'textarea':'input');
  input.id='shanhua-'+path.replaceAll('.','-')+(value?'-'+value:'');input.dataset.shanhuaField=path;
  if(type==='textarea')input.rows=3;else input.type=type;
  if(value){input.value=value;input.name='shanhua-'+path;}
  if(badge&&SHANHUA_CARRY_FIELDS.concat('report').includes(path))tag(caption,path);
  if(type==='checkbox'||type==='radio'){wrap.className='shanhua-check';wrap.append(input,caption);}else wrap.append(caption,input);
  parent.append(wrap);return input;
 }
 const reference=card('01','參考前一天紀錄');reference.id='shanhuaReference';
 const previousStatus=text('p','');previousStatus.id='shanhuaPrevStatus';
 const view=document.createElement('details');view.id='shanhuaPrevView';view.append(text('summary','查看前一天紀錄表'));
 const preview=text('div','');preview.id='shanhuaPrevFields';view.append(preview);
 const apply=text('button','套用前一天紀錄表');apply.id='shanhuaApplyPrev';apply.type='button';apply.disabled=true;
 apply.onclick=()=>runShanhua(async()=>{
  const found=previousShanhuaForm(shanhuaDay,readShanhuaDrafts());if(!found)return;
  const current=readShanhuaInputs(),complete=[...SHANHUA_CARRY_FIELDS,'report'];
  if((current.workContent||complete.some(key=>!current.carried.includes(key)))&&!confirm('今日紀錄已有修改，要套用 '+rocDate(found.day)+' 的紀錄表？'))return;
  const next={...found.form,day:current.day,docDate:current.docDate,carried:complete};
  fillShanhuaInputs(next);shanhuaCarried=[...complete];updateShanhuaPresentation();saveShanhuaDraft();
 });
 reference.append(previousStatus,view,apply);
 const heading=text('h3','');const no=text('span','02');no.className='step-no';heading.append(no,document.createTextNode('　填寫今日紀錄'));box.append(heading,form);
 const basics=group(form,'基本資料');
 const date=document.createElement('input');date.id='shanhuaDay';date.type='text';date.readOnly=true;const dateLabel=text('label','監看日期');dateLabel.append(date);basics.append(dateLabel);
 field(basics,'docDate','文件製作日期',{type:'date'});
 const name=document.createElement('input');name.type='text';name.value=SHANHUA_PROJECT_NAME;name.readOnly=true;const nameLabel=text('label','工程計畫名稱');nameLabel.append(name);basics.append(nameLabel);
 const details=document.createElement('details');details.id='shanhuaBasics';details.append(text('summary','沿用上次　修改'));basics.append(details);
 for(const [path,label] of [['owner','開發業主'],['contractor','施工單位'],['location','地點'],['placeName','小地名'],['site','相關遺址']])field(details,path,label);
 const today=form,work=group(today,'現地觀察與作業狀況');
 field(work,'workContent','今日工程內容',{type:'textarea'}).placeholder='例：今日繼續下挖3公尺並設置新管線…';
 const previous=text('button','帶入昨天內容');previous.id='shanhuaPrevWork';previous.type='button';work.append(previous);
 previous.onclick=()=>runShanhua(async()=>{
  const found=previousWorkContent(shanhuaDay,readShanhuaDrafts());if(!found){$('shanhuaStatus').textContent='沒有更早的工程內容可帶入';return;}
  const input=$('shanhua-workContent');if(input.value&&!confirm('工程內容已有文字，要以 '+found.day+' 的內容覆蓋？'))return;
  input.value=found.workContent;input.dispatchEvent(new Event('input',{bubbles:true}));
 });
 const strata=group(today,'地層堆積');field(strata,'stratigraphy','地層堆積狀況（土質、土色、堆積型態）',{type:'textarea'});
 const checks=group(strata,'地層勾選','landform noExcavation disturbedFill naturalLayer culturalLayer');
 const landform=field(checks,'landform.checked','地貌',{type:'checkbox'}).parentElement;
 const landformText=document.createElement('input');landformText.type='text';landformText.id='shanhua-landform-text';landformText.dataset.shanhuaField='landform.text';landformText.setAttribute('aria-label','地貌說明');landform.append(landformText);
 for(const [path,label] of [['noExcavation','無下挖'],['disturbedFill','擾亂填覆地層'],['naturalLayer','原堆積自然地層']])field(checks,path,label,{type:'checkbox',badge:false});
 field(checks,'culturalLayer.checked','原堆積文化層（測繪地層圖）',{type:'checkbox'});
 const cultureFields=text('div','');cultureFields.id='shanhuaCultureFields';cultureFields.className='shanhua-small-grid';checks.append(cultureFields);
 for(const [path,label] of [['layers','層數'],['thickness','厚度（公分）'],['from','起深（公分）'],['to','迄深（公分）']])field(cultureFields,'culturalLayer.'+path,label);
 const artifacts=group(today,'遺物出土狀況','artifacts');
 for(const [value,label] of [['none','無遺物出土'],['sparse','零星遺物出土'],['dense','遺物密集出土']])field(artifacts,'artifacts.status',label,{type:'radio',value});
 const artifactFields=text('div','');artifactFields.id='shanhuaArtifactFields';artifacts.append(artifactFields);field(artifactFields,'artifacts.position','位置');field(artifactFields,'artifacts.depth','深度');
 for(const [path,label] of [['features','遺跡、遺物'],['report','陳報資訊']])field(group(today,label),path,label,{type:'textarea'});
 const suggestion=document.createElement('details');suggestion.append(text('summary','相關建議（沿用，通常不用改）'));today.append(suggestion);field(suggestion,'suggestion','相關建議',{type:'textarea'});
 form.addEventListener('input',e=>{if(!shanhuaReady)return;const key=e.target.dataset.shanhuaField?.split('.')[0];shanhuaCarried=shanhuaCarried.filter(field=>field!==key);updateShanhuaPresentation();shanhuaDirty=true;try{saveShanhuaDraft();}catch(e){$('shanhuaStatus').textContent='草稿未儲存：'+e.message;}});
 const assets=group(today,'位置圖與簽名');field(assets,'mapCaption','位置圖說明',{type:'textarea'});
 for(const [kind,label] of [['map','施工與監看位置圖'],['monitor','監看人簽名'],['leader','計畫負責人簽名']]){
  const section=text('div','');section.className='shanhua-image';
  const input=document.createElement('input');input.type='file';input.accept='image/*';input.id='shanhuaImage-'+kind;
  input.className='shanhua-file-input';
  const wrap=text('label',label);wrap.htmlFor=input.id;
  const choose=text('button','選擇圖片');choose.type='button';choose.id='shanhuaChoose-'+kind;choose.onclick=()=>input.click();
  const image=document.createElement('img');image.id='shanhuaPreview-'+kind;image.alt=label;image.hidden=true;
  const remove=text('button','移除'+label);remove.type='button';remove.id='shanhuaRemove-'+kind;remove.hidden=true;
  input.onchange=()=>runShanhua(async()=>{const file=input.files[0];if(!file)return;const key=shanhuaImageKey(kind);await saveShanhuaFile(key,file);await refreshShanhuaImages();$('shanhuaStatus').textContent=label+'已儲存於本機';input.value='';});
  remove.onclick=()=>runShanhua(async()=>{await removeShanhuaFile(shanhuaImageKey(kind));await refreshShanhuaImages();$('shanhuaStatus').textContent=label+'已移除';});
  section.append(wrap,input,choose,remove,image);if(kind==='map'){const hint=text('p','沿用上次的位置圖');hint.id='shanhuaMapHint';hint.hidden=true;section.append(hint);}assets.append(section);
 }
 const note=text('p','簽名與位置圖只存在本機瀏覽器，換手機需重新上傳。草稿文字也只存在本機 localStorage，不在 TAR 備份內；簽名與位置圖亦不在 TAR 備份內。');note.className='export-note';assets.append(note);
 const output=card('03','輸出今日紀錄表');output.id='shanhuaOutput';const cloud=document.createElement('details');cloud.id='shanhuaCloud';cloud.append(text('summary','上傳到 Google 雲端（轉成 Google 文件）'));
 const folder=document.createElement('input');folder.type='hidden';folder.id='shanhuaFolder';
 const path=document.createElement('input');path.id='shanhuaFolderPath';path.readOnly=true;path.placeholder='尚未選擇資料夾';
 const folderLabel=text('label','善化 Google 雲端資料夾');folderLabel.append(path);
 const browse=text('button','選擇善化雲端資料夾');browse.id='shanhuaBrowseFolder';browse.type='button';browse.onclick=()=>{if(!hasToken()){$('shanhuaStatus').textContent='尚未登入 Google，請先在 Google 雲端設定連線帳號後再選擇資料夾。';return;}openDriveFolderBrowser('shanhua');};
 const download=text('button','下載 Word');download.id='shanhuaDownload';download.type='button';
 download.onclick=()=>runShanhua(async()=>{const {blob,name}=await prepareShanhuaDoc();const url=URL.createObjectURL(blob),link=document.createElement('a');link.href=url;link.download=name+'.docx';document.body.append(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);$('shanhuaStatus').textContent='已交給瀏覽器下載 Word，請確認檔案。';});
 const upload=text('button','上傳到 Google 雲端（轉成 Google 文件）');upload.id='shanhuaUpload';upload.type='button';
 upload.onclick=()=>runShanhua(async()=>{if(!hasToken())throw Error('尚未登入 Google，請先在 Google 雲端設定連線帳號');const folderId=parseDriveId(folder.value);if(!folderId)throw Error('請先選擇善化雲端資料夾');const day=shanhuaDay,{blob,name}=await prepareShanhuaDoc();await uploadDayDoc(day,blob,folderId,{name});$('shanhuaStatus').textContent='已上傳 Google 文件：'+name;});
 const connect=text('button','連線 Google 帳號');connect.id='shanhuaConnect';connect.type='button';connect.onclick=()=>runShanhua(async()=>{await authorize();$('shanhuaStatus').textContent='Google 已連線，請選擇善化雲端資料夾';});
 cloud.append(folder,folderLabel,connect,browse,upload);
 const status=text('p','');status.id='shanhuaStatus';status.setAttribute('role','status');
 const reset=text('button','以新草稿覆寫');reset.id='shanhuaResetDraft';reset.type='button';reset.hidden=true;
 reset.onclick=()=>{if(!confirm('確定以新草稿覆寫此日無法讀取的草稿？原草稿將被取代。'))return;try{const storage=safeGetLocalStorage();if(!storage)throw Error('本機儲存空間無法使用');let remembered={};try{remembered=JSON.parse(storage.getItem(SHANHUA_REMEMBER_KEY)||'{}');}catch{ /* Invalid remembered JSON must not block an explicit reset. */ }const fresh=emptyShanhuaForm(shanhuaDay,remembered);storage.setItem(monitoringDraftKey(shanhuaDay,SHANHUA_ID),JSON.stringify(fresh));shanhuaDay=null;hydrateShanhuaUI();}catch(e){status.textContent='未完成：'+e.message;}};
 output.append(download,cloud,status,reset);
 const other=document.createElement('details');other.id='shanhuaOtherExport';other.append(text('summary','Word / TAR 進階備份'));const tools=text('div','');tools.className='actions';other.append(tools);box.append(other);return box;
}
function readShanhuaDrafts(){
 const storage=safeGetLocalStorage(),drafts={},prefix=monitoringDraftKey('',SHANHUA_ID);
 if(storage)for(let i=0;i<storage.length;i++){const key=storage.key(i);if(key?.startsWith(prefix)){try{drafts[key.slice(prefix.length)]=JSON.parse(storage.getItem(key));}catch{ /* Preserve unreadable drafts. */ }}}
 return drafts;
}
function fillShanhuaInputs(form){
 for(const input of $('shanhuaForm').querySelectorAll('[data-shanhua-field]')){
  const value=input.dataset.shanhuaField.split('.').reduce((o,k)=>o[k],form);
  if(input.type==='checkbox')input.checked=value;else if(input.type==='radio')input.checked=input.value===value;else input.value=value;
 }
}
function updateShanhuaPrevious(){
 const found=previousShanhuaForm(shanhuaDay,readShanhuaDrafts());
 $('shanhuaPrevStatus').textContent=found?'前一天紀錄表：'+rocDate(found.day):'沒有更早的善化紀錄表草稿（只讀本機）';
 $('shanhuaApplyPrev').disabled=!shanhuaReady||!found;
 const preview=$('shanhuaPrevFields');preview.replaceChildren();
 if(!found)return;
 const form=found.form;
 preview.append(text('p','監看日期：'+rocDate(form.day)),text('p','文件製作日期：'+rocDate(form.docDate)),text('p','工程計畫名稱：'+SHANHUA_PROJECT_NAME));
 const labels={owner:'開發業主',contractor:'施工單位',location:'地點',placeName:'小地名',site:'相關遺址',workContent:'工程內容',stratigraphy:'地層堆積狀況',landform:'地貌',noExcavation:'無下挖',disturbedFill:'擾亂填覆地層',naturalLayer:'原堆積自然地層',culturalLayer:'原堆積文化層',artifacts:'遺物出土狀況',features:'遺跡、遺物',report:'陳報資訊',suggestion:'相關建議',mapCaption:'位置圖說明'};
 const nested={checked:'勾選',text:'說明',layers:'層數',thickness:'厚度（公分）',from:'起深（公分）',to:'迄深（公分）',status:'出土狀態',position:'位置',depth:'深度'};
 const display=value=>typeof value==='boolean'?(value?'✓':'—'):(value||'—');
 for(const [key,label] of Object.entries(labels)){
  const value=form[key];
  if(value&&typeof value==='object')for(const [part,item] of Object.entries(value))preview.append(text('p',label+'／'+nested[part]+'：'+(part==='status'?{none:'無遺物出土',sparse:'零星遺物出土',dense:'遺物密集出土'}[item]:display(item))));
  else preview.append(text('p',label+'：'+display(value)));
 }
}
function shanhuaControls(){return document.querySelectorAll('#shanhuaForm input,#shanhuaForm textarea,#shanhuaForm button,#shanhuaOutput input,#shanhuaOutput button,#shanhuaApplyPrev');}
function readShanhuaInputs(){
 if(!shanhuaReady||!shanhuaDay)throw Error('善化草稿尚未載入，未改寫資料');
 const form=emptyShanhuaForm(shanhuaDay);
 for(const input of $('shanhuaForm').querySelectorAll('[data-shanhua-field]')){
  const parts=input.dataset.shanhuaField.split('.'),key=parts.pop();let target=form;
  if(input.type==='radio'&&!input.checked)continue;
  for(const part of parts)target=target[part];target[key]=input.type==='checkbox'?input.checked:input.value;
 }
 form.carried=[...shanhuaCarried];return validateShanhuaForm(form);
}
function updateShanhuaPresentation(){
 for(const tag of $('shanhuaForm').querySelectorAll('[data-carried]'))tag.hidden=!tag.dataset.carried.split(/\s+/).some(key=>shanhuaCarried.includes(key));
 $('shanhuaBasics').querySelector('summary').textContent='沿用上次：'+['contractor','location','placeName','site'].map(key=>$('shanhua-'+key).value).join('／')+'　修改';
 $('shanhuaCultureFields').hidden=!$('shanhua-culturalLayer-checked').checked;
 $('shanhuaArtifactFields').hidden=$('shanhua-artifacts-status-none').checked;
}
function saveShanhuaDraft(){
 if(!shanhuaDay)return;
 const form=readShanhuaInputs(),storage=safeGetLocalStorage();if(!storage)throw Error('本機儲存空間無法使用');
 storage.setItem(monitoringDraftKey(shanhuaDay,SHANHUA_ID),JSON.stringify(form));
 storage.setItem(SHANHUA_REMEMBER_KEY,JSON.stringify(Object.fromEntries(shanhuaRememberFields.map(key=>[key,form[key]]))));
 shanhuaDirty=false;$('shanhuaStatus').textContent='草稿已自動儲存於本機';return form;
}
function hydrateShanhuaUI(){
 if(!validDay(activeDay)||shanhuaDay===activeDay)return;
 shanhuaDay=activeDay;shanhuaReady=false;shanhuaDirty=false;
 $('shanhuaDay').value=rocDate(shanhuaDay);
 try{
  const storage=safeGetLocalStorage(),raw=storage?.getItem(monitoringDraftKey(shanhuaDay,SHANHUA_ID));
  const form=raw?validateShanhuaForm(JSON.parse(raw)):emptyShanhuaForm(shanhuaDay,JSON.parse(storage?.getItem(SHANHUA_REMEMBER_KEY)||'{}'));
  if(form.day!==shanhuaDay)throw Error('草稿工作日不符，未覆寫');
  fillShanhuaInputs(form);
  shanhuaCarried=[...(form.carried||[])];updateShanhuaPresentation();
  $('shanhuaResetDraft').hidden=true;
  for(const input of shanhuaControls())input.disabled=false;
  shanhuaReady=true;updateShanhuaPrevious();
  const key=projectKey(MONITORING_FOLDER_KEY,SHANHUA_ID),id=readOptionalSetting(key)||'';
  $('shanhuaFolder').value=id;$('shanhuaFolderPath').value=savedFolderPath(key,id)||(id?'已儲存的資料夾（請重新選擇）':'');
  $('shanhuaStatus').textContent=raw?'已載入本機草稿':'新草稿；標示「沿用」的欄位來自上次或預設，工程內容請自行填寫';
  refreshShanhuaImages().catch(e=>{$('shanhuaStatus').textContent='圖片載入失敗：'+e.message;});
 }catch(e){
  shanhuaCarried=[];updateShanhuaPresentation();
  for(const input of shanhuaControls())input.disabled=true;
  updateShanhuaPrevious();
  $('shanhuaStatus').textContent='此日草稿無法讀取，已保留原資料；可按『以新草稿覆寫』重新開始';
  $('shanhuaResetDraft').hidden=false;$('shanhuaResetDraft').disabled=false;
 }
}
const shanhuaImageKey=kind=>kind==='map'?'map:current':'signature:'+kind;
async function refreshShanhuaImages(){
 const generation=++shanhuaGeneration,day=shanhuaDay;
 const images=await Promise.all(['map','monitor','leader'].map(kind=>getShanhuaFile(shanhuaImageKey(kind))));
 if(generation!==shanhuaGeneration||day!==shanhuaDay)return;
 ['map','monitor','leader'].forEach((kind,index)=>{
  const old=shanhuaUrls.get(kind);if(old)URL.revokeObjectURL(old);shanhuaUrls.delete(kind);
  const image=$('shanhuaPreview-'+kind),value=images[index];image.hidden=$('shanhuaRemove-'+kind).hidden=!value;
  $('shanhuaChoose-'+kind).textContent=value?'更換':'選擇圖片';if(kind==='map')$('shanhuaMapHint').hidden=!value;
  image.removeAttribute('src');if(value){const url=URL.createObjectURL(value.blob);shanhuaUrls.set(kind,url);image.src=url;}
 });
}
async function runShanhua(action){
 if(shanhuaWorking||isNanke()||!shanhuaReady)return;
 shanhuaWorking=true;for(const input of shanhuaControls())input.disabled=true;
 try{await action();}catch(e){$('shanhuaStatus').textContent='未完成：'+e.message;}
 finally{shanhuaWorking=false;for(const input of shanhuaControls())input.disabled=!shanhuaReady;updateShanhuaPrevious();}
}
async function prepareShanhuaDoc(){
 const form=saveShanhuaDraft(),[map,monitorSignature,leaderSignature]=await Promise.all([getShanhuaFile('map:current'),getShanhuaFile('signature:monitor'),getShanhuaFile('signature:leader')]);
 return {blob:await shanhuaDocx(form,{map,monitorSignature,leaderSignature}),name:'監看紀錄表_'+rocDate(form.day)};
}
window.addEventListener('beforeunload',e=>{if(shanhuaDirty||shanhuaWorking){e.preventDefault();e.returnValue='';}});
let projectUIReady=false,sharedToolSlots=[],nankeAlsoDoc=true;
function renderProjectUI(){
 if(!projectUIReady){
  projectUIReady=true;nankeAlsoDoc=$('photoAlsoDoc').checked;
  for(const p of PROJECTS){const option=text('option',p.shortName);option.value=p.id;option.title=p.name;$('project').append(option);}
  const wrapper=document.createElement('div');wrapper.id='nankeMonitoring';wrapper.append(...$('monitoring').childNodes);$('monitoring').append(wrapper,buildShanhuaUI());
  const shared=[...['exportWord','backup','persist','storage','photoDriveExport'].map($),$('restore').parentElement];
  sharedToolSlots=shared.map(node=>{const slot=document.createComment('shared project photo tool');node.before(slot);return {node,slot};});
  // Defense in depth: hidden Nanke controls must not dispatch for Shanhua.
  document.addEventListener('click',e=>{if(!isNanke()&&e.target.closest('#nankeMonitoring,#masterSetupDialog,#scopePhrasePanel')){e.preventDefault();e.stopImmediatePropagation();}},true);
 }
 $('project').value=currentProject;$('projectName').textContent=projectInfo(currentProject).name;
 $('areaDatasetNote').hidden=!projectInfo(currentProject).areaReference;
 $('nankeMonitoring').hidden=!isNanke();$('shanhuaMonitoring').hidden=isNanke();
 if(!isNanke())hydrateShanhuaUI();
 $('scopePhrasePanel').hidden=!isNanke();
 for(const {node,slot} of sharedToolSlots){
  if(isNanke())slot.after(node);
  else if(node.id==='photoDriveExport')$('shanhuaOtherExport').before(node);
  else $('shanhuaOtherExport').querySelector('.actions').append(node);
 }
 $('photoAlsoDoc').parentElement.hidden=!isNanke();
 if(!isNanke())$('photoAlsoDoc').checked=false;
}
async function changeProject(id){
 if(id===currentProject)return;
 if(!isNanke()&&shanhuaReady)saveShanhuaDraft();
 if(shanhuaWorking)throw Error('請先完成善化圖片儲存或輸出');
 if(busy||pending.length||rotating||monitoringWorking||uploadingPhotos||dayTransitioning||exportWorking)throw Error('請先完成照片儲存、旋轉或輸出，再切換專案');
 if(editorDirty())throw Error('記事尚未儲存，請先儲存或關閉編輯再切換專案');
 if(monDraftManager.hasAnyUnsaved())throw Error('監看草稿尚未安全儲存，請先處理再切換專案');
 if(isNanke()&&monDisplayedDay&&$('monForm').dataset.hydratedDay===monDisplayedDay){const saved=monDraftManager.saveDraft(monDisplayedDay,readMonInputs(true));if(!saved.success)throw Error('監看草稿儲存未完成，未切換專案');}
 stop();closeCoordinateAlbum(false);closePhrasePickers();
 for(const dialog of document.querySelectorAll('dialog[open]'))dialog.close();
 if(isNanke())nankeAlsoDoc=$('photoAlsoDoc').checked;
 const nextProject=projectOf({projectId:id});localStorage.setItem(PROJECT_KEY,nextProject);
 dayTransitioning=true;$('project').disabled=$('workDay').disabled=$('newDay').disabled=true;
 try{
 currentProject=nextProject;
 dayState=readDays();activeDay=dayState.active;monDisplayedDay=null;
 delete $('monForm').dataset.hydratedDay;
 edit=null;editBaseline=null;selected.clear();multi=false;releaseImages('order');
 $('photoAlsoDoc').checked=isNanke()?nankeAlsoDoc:false;
 $('photoDriveFolder').value=readOptionalSetting(projectKey(PHOTO_FOLDER_KEY,currentProject))||'';$('photoDrivePath').value=savedFolderPath(projectKey(PHOTO_FOLDER_KEY,currentProject),$('photoDriveFolder').value)||($('photoDriveFolder').value?'已儲存的資料夾（請重新選擇）':'');
 await load();if(active==='editor')tab('photos');else tab(active);
 say(`已切換專案：${projectInfo(currentProject).shortName}`);
 }finally{dayTransitioning=false;$('project').disabled=$('workDay').disabled=$('newDay').disabled=false;}
 if(isNanke())updateMonUI();monitoringControls();
}
$('project').onchange=()=>run(async()=>{try{await changeProject($('project').value);}finally{$('project').value=currentProject;}});
async function changeDay(day){
  if(!validDay(day))throw Error('日期請使用有效 YYYY-MM-DD');
  if(!isNanke()&&shanhuaReady)saveShanhuaDraft();
  if(shanhuaWorking)throw Error('請先完成善化圖片儲存或輸出');
  if(busy||pending.length||rotating)throw Error('請先完成照片儲存或旋轉再切換工作日');
  if(!canDiscardEditor()){$('workDay').value=activeDay;return;}
  stop();closeCoordinateAlbum(false);closePhrasePickers();

  // 1. Bind form saves to currently displayed day: flush pending edits before transition
  const prevDisplayed = monDisplayedDay;
  if(isNanke()&&prevDisplayed&&$('monForm').dataset.hydratedDay===prevDisplayed){
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
$('exportWord').onclick=()=>run(async()=>{if($('exportDialog').open)return;if(!validDay(activeDay)){say('工作日尚未載入，請稍候再開啟匯出。');return;}if(editorDirty()||pending.length||busy||rotating){say('請先儲存或關閉未完成的記事／照片，再匯出已儲存紀錄集。');return;}const day=activeDay,projectId=currentProject,snapshot=projectSnapshot(await db.snapshot());if(projectId!==currentProject||$('exportDialog').open||day!==activeDay||(isNanke()?$('monitoring').hidden:$('groups').hidden)||editorDirty()||pending.length||busy||rotating)return;const byId=new Map(snapshot.photos.map(p=>[p.id,p]));const included=snapshot.groups.filter(g=>g.workDay===day||g.photoIds.some(id=>byId.has(id)&&workDayOf(byId.get(id))===day)).sort(compareGroups);if(!included.length){say('本工作日沒有已儲存紀錄集可匯出。');return;}exportGeneration++;exportSnapshot=snapshot;exportDay=day;exportOrder=included.map(g=>g.id);exportSelected=new Set(exportOrder);exportWorking=false;resetExportPreview();$('exportStatus').textContent='請確認選取與順序，再按「預覽並製作表格」。';$('exportSnapshotInfo').textContent=`工作日 ${day} · 已儲存快照 ${new Date(snapshot.completedAt).toLocaleString()}；不修改原圖或記事。`;renderExportChoices();$('exportDialog').showModal();});
$('exportAll').onclick=()=>{exportSelected=new Set(exportOrder);resetExportPreview();renderExportChoices();};$('exportNone').onclick=()=>{exportSelected.clear();resetExportPreview();renderExportChoices();};
$('closeExport').onclick=()=>$('exportDialog').close();$('exportDialog').addEventListener('close',()=>{exportGeneration++;exportController?.abort();exportController=null;exportWorking=false;resetExportPreview();exportSnapshot=null;exportOrder=[];exportSelected.clear();$('exportRows').replaceChildren();});
$('prepareWord').onclick=async()=>{if(exportWorking||!$('exportDialog').open||!exportSnapshot)return;const token=++exportGeneration;exportController?.abort();exportController=new AbortController();const signal=exportController.signal,ids=exportOrder.filter(id=>exportSelected.has(id));exportWorking=true;resetExportPreview();exportControls();try{const prepared=await preparePhotoGroups(exportSnapshot,ids,{signal,onProgress:(n,total)=>{if(token===exportGeneration)$('exportStatus').textContent=`正在本機整理照片 ${n}/${total}，不會上傳…`;}});if(token!==exportGeneration||!$('exportDialog').open)return;const blob=await photoTableDocx(prepared);if(token!==exportGeneration||!$('exportDialog').open)return;const table=document.createElement('table');table.className='docx-preview-table';const body=document.createElement('tbody');for(const row of tablePlan(prepared)){const tr=document.createElement('tr');if(row.kind==='description'){const td=document.createElement('td');td.colSpan=2;const p=text('p',row.engineering);p.className='note-text';td.append(p);tr.append(td);}else for(let i=0;i<2;i++){const td=document.createElement('td'),photo=row.photos[i];if(photo){const im=document.createElement('img'),url=URL.createObjectURL(photo.blob);exportUrls.push(url);im.src=url;im.alt=photoLabel(exportSnapshot.photos.find(p=>p.id===photo.photoId));td.append(im);}else{const blank=text('div','');blank.className='empty-export-photo';blank.setAttribute('aria-label','本列無第二張照片');td.append(blank);}tr.append(td);}body.append(tr);}table.append(body);$('exportPreview').append(text('p','以下為連續表格預覽；Word 會依頁面長度自動跨頁。'),table);exportBlob=blob;const low=prepared.flatMap(g=>g.photos).filter(p=>Math.max(p.width,p.height)<800).length;$('exportStatus').textContent=`表格已製作完成，${prepared.length} 組／${prepared.reduce((n,g)=>n+g.photos.length,0)} 張；可下載可編輯 Word。${low?'其中 '+low+' 張解析度較低，列印放大可能模糊。':''}`;}catch(e){if(token===exportGeneration&&e.name!=='AbortError'){resetExportPreview();$('exportStatus').textContent='匯出未完成：'+e.message+'；原始資料未改動。';}}finally{if(token===exportGeneration){exportWorking=false;exportControls();}}};
$('downloadWord').onclick=()=>{if(!exportBlob||exportWorking)return;const url=URL.createObjectURL(exportBlob),a=document.createElement('a');a.href=url;a.download=`照片表格-${exportDay}.docx`;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);$('exportStatus').textContent='已交給瀏覽器下載 Word；請確認檔案存在並用 Word／相容軟體開啟。這不是完整原圖／中繼資料備份。';};
window.addEventListener('pagehide',()=>exportController?.abort());
const draftManagers=new Map();
function projectDraftManager(){if(!draftManagers.has(currentProject)){const storage=safeGetLocalStorage(),id=currentProject;const key=k=>k.startsWith('field-mon-form-')?monitoringDraftKey(k.slice('field-mon-form-'.length),id):projectKey(k,id);draftManagers.set(id,new MonitoringDraftManager(storage?{getItem:k=>storage.getItem(key(k)),setItem:(k,v)=>storage.setItem(key(k),v),removeItem:k=>storage.removeItem(key(k))}:null));}return draftManagers.get(currentProject);}
const monDraftManager=new Proxy({},{get:(_,name)=>{const manager=projectDraftManager();return typeof manager[name]==='function'?manager[name].bind(manager):manager[name];}});
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
 if(!isNanke()||dayTransitioning)return;
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
 if(!isNanke() || dayTransitioning || !d) return;
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
 try{const snapshot=projectSnapshot(await db.snapshot());if(generation!==coordinateAlbumGeneration)return;if(context.key!==phraseContext('monitoring')?.key){closeCoordinateAlbum(false);return;}coordinateAlbum.snapshot=snapshot;coordinateAlbum.entries=coordinateAlbumEntries(snapshot,context.day,'pending');renderCoordinateAlbum();}
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
 loadSource:async c=>{const snapshot=projectSnapshot(await readPhraseMetadata(await db.open())),choices=photoEngineeringPhrases(snapshot,c.day);return {choices,revision:JSON.stringify(choices.map(p=>[p.text,p.revision])),emptyMessage:'本工作日尚無工程描述片語' };},
 apply:(p,b)=>applyWorkPhrase($('monWorkScope'),p,b,true),restore:x=>restoreWorkPhrase($('monWorkScope'),x,true),
 successMessage:()=>monDraftManager.hasUnsaved(displayedDay())?{text:'已插入，但草稿尚未儲存。',quiet:false}:{text:'已插入',quiet:true}
});
scopePhrasePicker=new WorkPhrasePicker({
 ...phrasePickerElements('scope','scopePhrasePanel','engineering'),getContext:()=>phraseContext('engineering'),
 loadSource:async c=>{if(!isNanke())return {choices:[],emptyMessage:'善化監看紀錄表將在下一版提供'};if(monDraftManager.hasUnsaved(c.day)||monDraftManager.hasCorrupt(c.day)||monDraftManager.hasConflict(c.day)||monDraftManager.isReadBlocked(c.day))throw Error('本工作日監看草稿尚未安全儲存，請先重試儲存或匯出備份；不使用舊草稿冒充最新內容');return {...readSavedScopePhrases(safeGetLocalStorage(),c.day),emptyMessage:'本工作日尚無作業子項目片語' };},
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
 if(!isNanke())throw Error('善化監看紀錄表將在下一版提供');
 const snapshot=projectSnapshot(await db.snapshot()),byId=new Map(snapshot.photos.map(p=>[p.id,p]));
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
const dayFolderKey=day=>projectKey(`field-photo-day-folder-${day}`,currentProject),dayDocKey=day=>projectKey(`field-standalone-doc-${day}`,currentProject);
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
const savedFolder=readOptionalSetting(projectKey(PHOTO_FOLDER_KEY,currentProject));
if(savedFolder&&$('photoDriveFolder'))$('photoDriveFolder').value=savedFolder;

$('photoDriveFolder')?.addEventListener('change',()=>{
 try{safeGetLocalStorage()?.setItem(projectKey(PHOTO_FOLDER_KEY,currentProject),$('photoDriveFolder').value.trim());}catch{}
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
function folderTarget(){if(folderDestination==='shanhua')return {input:$('shanhuaFolder'),path:$('shanhuaFolderPath'),status:$('shanhuaStatus'),key:projectKey(MONITORING_FOLDER_KEY,'shanhua-1350pscp')};return folderDestination==='monitoring'
 ? {input:$('driveFolder'),path:$('monitoringDrivePath'),status:$('driveStatus'),key:MONITORING_FOLDER_KEY}
 : {input:$('photoDriveFolder'),path:$('photoDrivePath'),status:$('photoDriveStatus'),key:projectKey(PHOTO_FOLDER_KEY,currentProject)};}
function savedFolderPath(key,id){
 try{const stored=JSON.parse(safeGetLocalStorage()?.getItem(key+'-path')||'null');return stored?.id===id&&typeof stored.path==='string'?stored.path:'';}catch{return '';}
}
$('photoDrivePath').value=savedFolderPath(projectKey(PHOTO_FOLDER_KEY,currentProject),savedFolder)|| (savedFolder?'已儲存的資料夾（請重新選擇）':'');
$('monitoringDrivePath').value=savedFolderPath(MONITORING_FOLDER_KEY,savedMonitoringFolder)|| (savedMonitoringFolder?'已儲存的資料夾（請重新選擇）':'');
let currentBrowsingFolderId=null,currentBrowsingFolderName='';
let browsingShared=false,browsingDriveId=null,browsingDriveList=false,folderHistory=[],folderRequest=0;
async function selectDriveFolder(id,name){
 if(folderDestination==='master')return;
 try{
  if(['monitoring','shanhua'].includes(folderDestination)&&!id){$('folderBrowserStatus').textContent='請先進入或建立一個具體資料夾，再選取。';return;}
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
 $('folderBrowserHeading').textContent=destination==='master'?'選擇本月 Google 母表副本':['monitoring','shanhua'].includes(destination)?'選擇監看紀錄表的 Drive 資料夾':'選擇照片的 Drive 資料夾';
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
  if(isNanke()&&$('photoAlsoDoc')?.checked&&targetFolderId){
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
if($('masterBuild'))$('masterBuild').textContent=`App v45 / 追加核心：${docsAppendInfo.DOCS_APPEND_BUILD||'舊版，請使用下方更新連結'}`;
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
  for(const day of e.key===null ? [...new Set([displayedDay(),...monDraftManager.getUnsavedDays()])] : e.key?.startsWith(monitoringDraftKey('',currentProject)) ? [e.key.slice(monitoringDraftKey('',currentProject).length)] : []){
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
$('homeScreenMonitoring').onclick=()=>{$('homeScreenDialog').close();tab('monitoring');if(!isNanke())return;$('monitoringOtherExport').open=true;$('exportMonDocx').scrollIntoView({block:'center'});};
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
