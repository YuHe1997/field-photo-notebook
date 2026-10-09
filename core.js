export function readingAt(reading,time,limit){return reading && reading.readingAt<=time ? {...reading,ageAtShutterMs:time-reading.readingAt,freshness:time-reading.readingAt<=limit?'新鮮':'過期'}:null;}
export function validateGroup(g,photos){if(!g.id||!g.title.trim()||!Array.isArray(g.photoIds)||!g.photoIds.length||new Set(g.photoIds).size!==g.photoIds.length||g.photoIds.some(id=>!photos.has(id)))throw Error('紀錄集資料或照片參照無效');return g;}
export function move(ids,index,delta){const out=[...ids],dest=index+delta;if(dest>=0&&dest<out.length)[out[index],out[dest]]=[out[dest],out[index]];return out;}
export const digest=async b=>Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',await b.arrayBuffer()))).map(n=>n.toString(16).padStart(2,'0')).join('');
// POSIX ustar, uncompressed: interoperable archive without external services/libraries.
export async function tar(entries){const parts=[];for(const [name,blob] of entries){if(name.length>99)throw Error('路徑過長');const h=new Uint8Array(512),put=(at,s)=>h.set(new TextEncoder().encode(s),at);put(0,name);put(100,'0000644\0');put(108,'0000000\0');put(116,'0000000\0');put(124,blob.size.toString(8).padStart(11,'0')+'\0');put(136,'00000000000\0');h.fill(32,148,156);put(156,'0');put(257,'ustar\0');put(263,'00');put(148,h.reduce((a,b)=>a+b,0).toString(8).padStart(6,'0')+'\0 ');parts.push(h,blob,new Uint8Array((512-blob.size%512)%512));}parts.push(new Uint8Array(1024));return new Blob(parts,{type:'application/x-tar'});}
export async function untar(blob){const a=new Uint8Array(await blob.arrayBuffer()),out=new Map(),dec=new TextDecoder();let p=0;while(p+512<=a.length){const h=a.slice(p,p+512);if(h.every(x=>!x))return out;const str=(s,e)=>dec.decode(h.slice(s,e)).split('\0')[0];const name=str(0,100),size=parseInt(str(124,136),8),sum=parseInt(str(148,156),8);h.fill(32,148,156);if(h.reduce((x,y)=>x+y,0)!==sum||!Number.isSafeInteger(size)||size<0||p+512+size>a.length||!name||name.includes('..')||name.startsWith('/')||out.has(name)||![0,48].includes(a[p+156]))throw Error('備份封存格式無效');out.set(name,new Blob([a.slice(p+512,p+512+size)]));p+=512+Math.ceil(size/512)*512;}throw Error('備份結尾不完整');}

export function localDay(time=Date.now()){const d=new Date(time);if(!Number.isFinite(d.getTime()))return null;return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;}
export function validDay(day){if(typeof day!=='string'||!/^\d{4}-\d{2}-\d{2}$/.test(day))return false;const [y,m,d]=day.split('-').map(Number),date=new Date(y,m-1,d);return date.getFullYear()===y&&date.getMonth()===m-1&&date.getDate()===d;}
// Extract the first valid calendar day, preferring a record's title over its content.
// Returns a normalized YYYY-MM-DD day, or null if no valid day is present.
export function extractMonitoringDate(record){
 const fields=typeof record==='string'?[record]:[record?.title,record?.content];
 for(const field of fields){
  if(typeof field!=='string')continue;
  const dates=/(?<!\d)(\d{4})(?:年(\d{1,2})月(\d{1,2})日|-(\d{1,2})-(\d{1,2}))(?!\d)/g;
  for(const match of field.matchAll(dates)){
   const day=`${match[1]}-${(match[2]??match[4]).padStart(2,'0')}-${(match[3]??match[5]).padStart(2,'0')}`;
   if(validDay(day))return day;
  }
 }
 return null;
}
// Dates must be valid normalized days. Lexical order then equals calendar order.
export function compareMonitoringDates(a,b){
 if(!validDay(a)||!validDay(b))throw new TypeError('請使用有效 YYYY-MM-DD 日期比較');
 return a<b?-1:a>b?1:0;
}
// Existing records must be chronological; equal dates retain their original order.
// Unknown-date records are kept after dated records rather than guessed into a day.
export function monitoringInsertionIndex(records,day){
 if(!validDay(day))throw new TypeError('請使用有效 YYYY-MM-DD 插入日期');
 for(let i=0;i<records.length;i++){
  const existing=extractMonitoringDate(records[i]);
  if(existing===null||compareMonitoringDates(existing,day)>0)return i;
 }
 return records.length;
}
export function workDayOf(p){return validDay(p.workDay)?p.workDay:localDay(p.buttonPressedAt??p.capturedAt??p.importedAt??p.createdAt??Date.now());}
export function numberedPhoto(p){return `照片 ${String(p.photoNumber??'?').padStart(2,'0')}`;}
export function numberedGroup(ids,photos){return '照片 '+ids.map(id=>String(photos.find(p=>p.id===id)?.photoNumber??'?').padStart(2,'0')).join('、');}
export function previewRotation(p,width,height){if([0,90,180,270].includes(p.displayRotation))return p.displayRotation;const angle=p.screenAngle??p.heading?.screenAngle,base=height<=width?0:(angle===270||angle===-90?270:90);return (base+180)%360;}
export function photoMinute(p){const time=p.buttonPressedAt??p.capturedAt??p.importedAt;if(!Number.isFinite(time))return '時間未取得';const d=new Date(time);if(!Number.isFinite(d.getTime()))return '時間未取得';const two=n=>String(n).padStart(2,'0');return `${d.getFullYear()}/${two(d.getMonth()+1)}/${two(d.getDate())} ${two(d.getHours())}:${two(d.getMinutes())}${p.source==='gallery'?'(匯入)':''}`;}
export function compactCoordinates(l){if(!l||!Number.isFinite(l.latitude)||!Number.isFinite(l.longitude)||Math.abs(l.latitude)>90||Math.abs(l.longitude)>180)return '座標未取得';const f=n=>n.toFixed(7).replace(/\.?0+$/,'');return `${f(l.latitude)},${f(l.longitude)}`;}
export function compactDirection(h){const label=directionLabel(h?.degrees);if(!label)return '方向未取得';const notes=[];if(h.accuracy<0||(typeof h.quality==='string'&&h.quality.includes('不可靠')))notes.push('讀值不可靠');if((h.freshnessAtAcquisitionCompleted??h.freshness)==='過期')notes.push('讀值過期');notes.push('需驗證');return `${label}(${notes.join('，')})`;}
export function compactPhotoLines(p,index){return [`${index+1}.照片${p?.photoNumber??'?'}`,p?photoMinute(p):'時間未取得',compactCoordinates(p?.location),compactDirection(p?.heading)];}

export function zoomCapability(cap,reported){const z=cap?.zoom;if(!z||!Number.isFinite(z.min)||!Number.isFinite(z.max)||z.min<=0||z.max<=z.min||!Number.isFinite(reported)||reported<z.min||reported>z.max)return null;return {min:z.min,max:z.max,step:Number.isFinite(z.step)&&z.step>0?z.step:.1};}
export function snapZoom(value,range){if(!Number.isFinite(value))throw Error('縮放倍率無效');const n=Math.max(range.min,Math.min(range.max,value)),snapped=range.min+Math.round((n-range.min)/range.step)*range.step;return Number(Math.max(range.min,Math.min(range.max,snapped)).toFixed(6));}
export function cropRect(width,height,factor){if(!Number.isFinite(width)||!Number.isFinite(height)||width<=0||height<=0||!Number.isFinite(factor)||factor<1)throw Error('相機影格或數位倍率尚未就緒');const sw=width/factor,sh=height/factor;return {sx:(width-sw)/2,sy:(height-sh)/2,sw,sh,width:Math.max(1,Math.floor(sw)),height:Math.max(1,Math.floor(sh))};}

// Half-open sectors: [337.5, 22.5) is north; labels describe opposite → reference.
export function directionLabel(degrees){
  if(!Number.isFinite(degrees))return null;
  const labels=['北','東北','東','東南','南','西南','西','西北'];
  const sector=Math.floor((((degrees%360)+360)%360+22.5)/45)%8;
  return `${labels[(sector+4)%8]}向${labels[sector]}`;
}
export function headingSummary(h){
  const label=directionLabel(h?.degrees);
  if(!label)return '拍攝方向：未取得；鏡頭朝向未確認';
  const unreliable=h.accuracy<0||h.quality?.includes('不可靠');
  const fresh=h.freshnessAtAcquisitionCompleted??h.freshness??'新鮮度未知';
  return `裝置參考推估：${label}；鏡頭朝向未確認；${unreliable?'平台回報不可靠；':''}${h.quality??'精度未保證'}；記錄時${fresh}（按鈕／取得讀值，非曝光方向）`;
}
