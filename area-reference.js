// Pure, offline reference geometry. Does not read sensors/storage or change photo records.
export const AREA_ENGINE_BUILD='area-reference-v1';
const finite=Number.isFinite,rad=Math.PI/180;
const dist=(a,b)=>Math.hypot(a[0]-b[0],a[1]-b[1]);
const orient=(a,b,c)=>(b[0]-a[0])*(c[1]-a[1])-(b[1]-a[1])*(c[0]-a[0]);
function on(a,b,p){return Math.abs(orient(a,b,p))<1e-8&&p[0]>=Math.min(a[0],b[0])-1e-8&&p[0]<=Math.max(a[0],b[0])+1e-8&&p[1]>=Math.min(a[1],b[1])-1e-8&&p[1]<=Math.max(a[1],b[1])+1e-8;}
function intersects(a,b,c,d){return orient(a,b,c)*orient(a,b,d)<0&&orient(c,d,a)*orient(c,d,b)<0||on(a,b,c)||on(a,b,d)||on(c,d,a)||on(c,d,b);}
export function simpleRing(p){if(!Array.isArray(p)||p.length<3||p.length>10000)return false;let area=0;for(let i=0;i<p.length;i++){const a=p[i],b=p[(i+1)%p.length],prev=p[(i+p.length-1)%p.length];if(!Array.isArray(a)||a.length!==2||!a.every(finite)||!Array.isArray(b)||b.length!==2||!b.every(finite)||!Array.isArray(prev)||prev.length!==2||!prev.every(finite))return false;if(dist(a,b)<.01)return false;if(Math.abs(orient(prev,a,b))<1e-8&&(prev[0]-a[0])*(b[0]-a[0])+(prev[1]-a[1])*(b[1]-a[1])>0)return false;area+=a[0]*b[1]-a[1]*b[0];for(let j=i+1;j<p.length;j++){if(j===i+1||i===0&&j===p.length-1)continue;const c=p[j],d=p[(j+1)%p.length];if(!Array.isArray(c)||!Array.isArray(d)||intersects(a,b,c,d))return false;}}return Math.abs(area)>1e-6;}
export function validateAreaDataset(d){
 if(d?.schema!=='field-area-reference-v1'||typeof d.version!=='string'||!d.version||typeof d.sourceSha256!=='string'||!/^[a-f0-9]{64}$/.test(d.sourceSha256)||d.status!=='provisional')throw Error('區界資料版本不完整');
 const b=d.bounds,p=d.policy;if(!b||!['north','south','east','west','rotation'].every(k=>finite(b[k]))||b.rotation!==0||b.north<=b.south||b.east<=b.west||b.north>90||b.south< -90||b.east>180||b.west< -180||b.north-b.south>.1||b.east-b.west>.1)throw Error('區界定位不受支援，需重新校核');
 if(!finite(d.width)||!finite(d.height)||d.width<=0||d.height<=0||!p||!['mapBufferM','maxAccuracyM','nearestLimitM','maxReadingAgeMs'].every(k=>finite(p[k])&&p[k]>=0)||p.mapBufferM>200||p.maxAccuracyM>500||p.nearestLimitM>1000||p.maxReadingAgeMs>10000)throw Error('區界容差設定無效');
 if(!Array.isArray(d.regions)||!d.regions.length||d.regions.length>1000)throw Error('沒有可用區域');const ids=new Set();
 for(const r of d.regions){if(!r||typeof r.id!=='string'||!r.id||ids.has(r.id)||typeof r.name!=='string'||!r.name.trim()||r.name.length>100||/[\u0000-\u001f]/u.test(r.name)||!simpleRing(r.points)||r.points.some(p=>p[0]<0||p[0]>d.width||p[1]<0||p[1]>d.height))throw Error('區界多邊形不完整或交叉');ids.add(r.id);}return true;
}
export function pixelToLonLat(point,d){return [d.bounds.west+point[0]/d.width*(d.bounds.east-d.bounds.west),d.bounds.north-point[1]/d.height*(d.bounds.north-d.bounds.south)];}
// WGS84 local linear metric, anchored at the small project centre; no map tiles/network.
export function localMetric(bounds){const lon=(bounds.east+bounds.west)/2,lat=(bounds.north+bounds.south)/2,phi=lat*rad,e2=6.6943799901413165e-3,a=6378137,w=Math.sqrt(1-e2*Math.sin(phi)**2),mx=a/w*Math.cos(phi)*rad,my=a*(1-e2)/w**3*rad;return p=>[(p[0]-lon)*mx,(p[1]-lat)*my];}
export function segmentDistance(p,a,b){const dx=b[0]-a[0],dy=b[1]-a[1],n=dx*dx+dy*dy,t=n?Math.max(0,Math.min(1,((p[0]-a[0])*dx+(p[1]-a[1])*dy)/n)):0;return Math.hypot(p[0]-a[0]-t*dx,p[1]-a[1]-t*dy);}
export function ringDistance(p,ring){let inside=false,distanceM=Infinity;for(let i=0,j=ring.length-1;i<ring.length;j=i++){const a=ring[j],b=ring[i];distanceM=Math.min(distanceM,segmentDistance(p,a,b));if((a[1]>p[1])!==(b[1]>p[1])&&p[0]<(b[0]-a[0])*(p[1]-a[1])/(b[1]-a[1])+a[0])inside=!inside;}return {inside:inside||distanceM<1e-7,distanceM};}
export function createAreaMatcher(dataset){
 validateAreaDataset(dataset);const d=JSON.parse(JSON.stringify(dataset)),metric=localMetric(d.bounds),regions=d.regions.map(r=>({...r,ring:r.points.map(p=>metric(pixelToLonLat(p,d)))}));
 return photo=>{
  const base={engine:AREA_ENGINE_BUILD,datasetVersion:d.version,sourceSha256:d.sourceSha256,provisional:true,status:'unavailable',reason:null,candidates:[],accuracyM:null,mapBufferM:d.policy.mapBufferM,uncertaintyRadiusM:null,positionMeaning:'拍攝者位置，不是被攝物；依已保存定位及目前暫定區界回算'};
  const stop=reason=>({...base,reason});
  if(photo?.source!=='camera')return stop('not-camera');const l=photo.location;
  if(!l||!finite(l.latitude)||!finite(l.longitude)||Math.abs(l.latitude)>90||Math.abs(l.longitude)>180)return stop('no-valid-gps');
  if(!finite(l.accuracyM)||l.accuracyM<0)return stop('unknown-accuracy');base.accuracyM=l.accuracyM;if(l.accuracyM>d.policy.maxAccuracyM)return stop('low-accuracy');
  if(l.freshness==='過期'||l.freshnessAtAcquisitionCompleted==='過期')return stop('stale-gps');
  let age=finite(photo.buttonPressedAt)&&finite(l.readingAt)?photo.buttonPressedAt-l.readingAt:finite(l.ageAtButtonPressMs)?l.ageAtButtonPressMs:finite(l.ageAtShutterMs)?l.ageAtShutterMs:null;
  if(age===null||age<0)return stop('unknown-gps-time');if(age>d.policy.maxReadingAgeMs)return stop('stale-gps');
  if(finite(photo.acquisitionCompletedAt)&&finite(l.readingAt)){const endAge=photo.acquisitionCompletedAt-l.readingAt;if(endAge<0||endAge>d.policy.maxReadingAgeMs)return stop('stale-gps');}
  if(finite(l.ageAtAcquisitionCompletedMs)&&(l.ageAtAcquisitionCompletedMs<0||l.ageAtAcquisitionCompletedMs>d.policy.maxReadingAgeMs))return stop('stale-gps');
  const p=metric([l.longitude,l.latitude]),merged=new Map(),radius=l.accuracyM+d.policy.mapBufferM;
  for(const r of regions){const s=ringDistance(p,r.ring),hit=s.inside||s.distanceM<=radius+1e-7,existing=merged.get(r.name);if(!existing)merged.set(r.name,{name:r.name,regionIds:[r.id],inside:s.inside,distanceM:s.inside?0:s.distanceM,boundaryDistanceM:s.distanceM,hit});else{existing.regionIds.push(r.id);existing.inside ||=s.inside;existing.distanceM=Math.min(existing.distanceM,s.inside?0:s.distanceM);existing.boundaryDistanceM=Math.min(existing.boundaryDistanceM,s.distanceM);existing.hit ||=hit;}}
  const all=[...merged.values()].sort((a,b)=>Number(b.inside)-Number(a.inside)||a.distanceM-b.distanceM||a.name.localeCompare(b.name,'zh-Hant'));
  let candidates=all.filter(r=>r.hit),status;
  if(candidates.length){status=candidates.length===1&&candidates[0].inside&&candidates[0].boundaryDistanceM>radius?'within-reference':'boundary';}
  else{candidates=all.filter(r=>r.distanceM<=d.policy.nearestLimitM);status=candidates.length?'near-outside':'outside';}
  return {...base,status,reason:null,uncertaintyRadiusM:radius,candidates:candidates.map(({hit,...r})=>({...r,distanceM:Math.round(r.distanceM*10)/10,boundaryDistanceM:Math.round(r.boundaryDistanceM*10)/10}))};
 };
}
export function areaReferenceText(r){
 if(!r)return '區域參考未取得';const reasons={'not-camera':'匯入照片沒有拍照當時 GPS，不判定','no-valid-gps':'GPS 未取得，不判定','unknown-accuracy':'GPS 精度未知，不判定','low-accuracy':'GPS 精度不足，不判定','stale-gps':'拍照定位過期，不判定','unknown-gps-time':'無法確認拍照定位時效，不判定','dataset-invalid':'區界資料檢查未通過，不判定'};
 if(r.status==='unavailable')return '區域參考：'+(reasons[r.reason]||'資訊不足，不判定');
 const names=r.candidates.slice(0,5).map(c=>c.name).join('／')+(r.candidates.length>5?` 等 ${r.candidates.length} 區`:'');
 if(r.status==='within-reference')return `區域參考：接近 ${names}（座標位於標註內；暫定、需確認）`;
 if(r.status==='boundary')return `區域參考：${names} 邊界／誤差範圍附近，請確認（暫定）`;
 if(r.status==='near-outside')return `區域參考：接近 ${names}，但在標註區外（暫定、需確認）`;
 return '區域參考：不在目前標註區域附近，不硬歸類';
}
