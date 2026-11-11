import {validDay} from './core.js';

export const NUKE1_ID='nuke1-monitoring';
export const NUKE1_PROJECT_NAME='核能一廠三號低放射性廢棄物貯存庫暨土石方暫置場興建工程';
export const NUKE1_REMARK='遺物清冊應併監看工作日誌繳交。';
export const NUKE1_FIELDS=Object.freeze(['contractor','location','site','workScope','soil','color','strata','remains','artifacts','features','advice','otherEvents']);
export const NUKE1_LABELS=Object.freeze({contractor:'2.施工單位',location:'3.地點',site:'4.考古遺址',workScope:'5.作業項目、範圍與深度',soil:'6.土質',color:'7.土色',strata:'8.地層堆積狀況（描述、繪圖或照相）',remains:'9.考古遺留出土狀況（描述、繪圖或照相）',artifacts:'10.遺物（類別、編號與描述）',features:'11.遺跡（類別、編號與描述）',advice:'12.施工或搶救發掘之建議',otherEvents:'13.其他重要紀事（土方運入來源與運出處置之地點）'});
export const NUKE1_DEFAULTS=Object.freeze({contractor:'榮金營造股份有限公司',location:'',site:'',workScope:'',soil:'',color:'',strata:'',remains:'無',artifacts:'無',features:'無',advice:'未發現考古遺物及文化層。施工現場配合度佳，建議持續監看，監看過程中如發現任何涉文化資產標的，將依《文化資產保存法》第33條、第57條、第77條、第88條規定辦理。若有發現疑似考古遺址文化資產時，將即通報主管機關依照「文化資產保存法施行細則第27條及27之1條」辦理。',otherEvents:''});
function checkDay(day){if(!validDay(day))throw Error('日期無效，請使用 YYYY-MM-DD');}
export function emptyNuke1Form(day){checkDay(day);return {day,...NUKE1_DEFAULTS};}
export function validateNuke1Form(form){
 if(!form||typeof form!=='object'||Array.isArray(form))throw Error('核一表單資料格式無效');
 checkDay(form.day);const result={day:form.day};
 for(const key of NUKE1_FIELDS){if(typeof form[key]!=='string')throw Error('核一表單欄位格式無效');result[key]=form[key];}
 return result;
}
export function previousNuke1Value(day,drafts,key){
 checkDay(day);if(!NUKE1_FIELDS.includes(key))throw Error('核一欄位無效');let previous=null;
 for(const [date,value] of Object.entries(drafts||{})){
  try{checkDay(date);const form=validateNuke1Form(value);if(form.day===date&&date<day&&form[key].trim()&&(!previous||date>previous.day))previous={day:date,value:form[key]};}catch{ /* Preserve and skip invalid drafts. */ }
 }
 return previous;
}
export function nuke1Date(day){checkDay(day);return `${Number(day.slice(0,4))-1911}.${day.slice(5,7)}.${day.slice(8,10)}`;}
export function appendCoordinates(location,{latitude,longitude}={}){
 if(typeof location!=='string'||!Number.isFinite(latitude)||!Number.isFinite(longitude)||Math.abs(latitude)>90||Math.abs(longitude)>180)throw Error('座標無效');
 const f=n=>n.toFixed(7).replace(/\.?0+$/,'');const prefix=location.trimEnd();
 return `${prefix}${prefix?'\n':''}（${f(latitude)}, ${f(longitude)}）\nWGS84系統`;
}
