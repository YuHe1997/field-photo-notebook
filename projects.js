import {validDay,workDayOf} from './core.js';
export const DEFAULT_PROJECT='nanke-chiayi-2';
export const PROJECTS=Object.freeze([
 Object.freeze({id:DEFAULT_PROJECT,name:'南部科學園區嘉義園區二期考古遺址監看計畫',shortName:'南科嘉義二期',areaReference:true,monitoringFormat:'nanke'}),
 Object.freeze({id:'shanhua-1350pscp',name:'臺南市善化區光復路至文昌路1350PSCP汰換管線工程施工監看計畫',shortName:'善化 1350PSCP 汰換管線',areaReference:false,monitoringFormat:'shanhua'}),
 Object.freeze({id:'nuke1-monitoring',name:'核能一廠三號低放射性廢棄物貯存庫暨土石方暫置場興建工程',shortName:'核一監看案',areaReference:false,monitoringFormat:'nuke1'})
]);
// Unknown identifiers fail closed to the legacy project, never to Shanhua.
export const projectOf=record=>PROJECTS.some(p=>p.id===record?.projectId)?record.projectId:DEFAULT_PROJECT;
export const projectInfo=id=>PROJECTS.find(p=>p.id===id)||PROJECTS[0];
export const projectKey=(base,id=DEFAULT_PROJECT)=>id===DEFAULT_PROJECT?base:`${base}-${projectOf({projectId:id})}`;
export const counterKey=id=>projectKey('field-photo-notebook-number-counters',id);
export const daysKey=id=>projectKey('field-photo-notebook-work-days',id);
export const monitoringDraftKey=(day,id=DEFAULT_PROJECT)=>id===DEFAULT_PROJECT?`field-mon-form-${day}`:`field-mon-form-${projectOf({projectId:id})}-${day}`;
export function readCounters(storage,id=DEFAULT_PROJECT){try{const data=JSON.parse(storage.getItem(counterKey(id))||'{}');return Object.fromEntries(Object.entries(data).filter(([d,n])=>validDay(d)&&Number.isSafeInteger(n)&&n>0));}catch{return {};}}
export function mergeProjectCounters(storage,incoming,id=DEFAULT_PROJECT){const out=readCounters(storage,id);for(const [d,n] of Object.entries(incoming??{}))if(validDay(d)&&Number.isSafeInteger(n)&&n>0)out[d]=Math.max(out[d]||0,n);storage.setItem(counterKey(id),JSON.stringify(out));return out;}
export function countersForRecords(storage,records,id=DEFAULT_PROJECT){const out=readCounters(storage,id);for(const p of records)if(projectOf(p)===id&&Number.isSafeInteger(p.photoNumber)&&p.photoNumber>0){const d=workDayOf(p);out[d]=Math.max(out[d]||0,p.photoNumber);}return out;}
export const numberIdentity=p=>`${projectOf(p)}:${workDayOf(p)}:${p.photoNumber}`;
export const backupRecord=p=>({...p,projectId:projectOf(p)});
export function restoreProjectRecords(manifest){const photos=manifest.photos.map(backupRecord),groups=manifest.groups.map(backupRecord);const numbers=new Set();for(const p of photos)if(p.photoNumber){const key=numberIdentity(p);if(numbers.has(key))throw Error('專案工作日照片編號重複');numbers.add(key);}const byId=new Map(photos.map(p=>[p.id,p]));for(const g of groups)if(g.photoIds.some(id=>byId.has(id)&&projectOf(byId.get(id))!==projectOf(g)))throw Error('紀錄集不可跨專案');return {photos,groups};}
