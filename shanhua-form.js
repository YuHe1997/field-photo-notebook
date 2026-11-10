import {localDay} from './core.js';

export const SHANHUA_PROJECT_NAME='臺南市善化區光復路至文昌路1350PSCP汰換管線工程施工監看計畫';
export const SHANHUA_STRATA='地表至地下1.2公尺有道路結構（0～1.2公尺為回填，上面級配厚度從0.5～1.2公尺均有，下面回填土，兩者之中均有些許磚頭和塑膠麻布袋碎片混雜其中），地表下1.2～3公尺為乾淨的黃褐色砂質壤土。';
export const SHANHUA_CARRY_FIELDS=Object.freeze(['owner','contractor','location','placeName','site','stratigraphy','landform','noExcavation','disturbedFill','naturalLayer','culturalLayer','artifacts','features','suggestion','mapCaption']);
export const SHANHUA_DEFAULTS=Object.freeze({
 day:'',docDate:'',owner:'台灣自來水股份有限公司第六區管理處',contractor:'長源工程有限公司',
 location:'善化區南關里文昌路至三民路',placeName:'牛墟',site:'大洲里埔遺址',workContent:'',stratigraphy:SHANHUA_STRATA,
 landform:Object.freeze({checked:false,text:''}),noExcavation:false,disturbedFill:true,naturalLayer:true,
 culturalLayer:Object.freeze({checked:false,layers:'',thickness:'',from:'',to:''}),
 artifacts:Object.freeze({status:'none',position:'',depth:''}),features:'無',
 suggestion:'建議持續監看，監看過程中如發現任何涉文化資產標的，將依《文化資產保存法》第33條、第57條、第77條、第88條規定辦理。若有發現疑似考古遺址文化資產時，將即通報主管機關依照「文化資產保存法施行細則第27條及27之1條」辦理。',
 report:'無。',mapCaption:'善美大道339號'
});
export function rocDate(day){
 if(typeof day!=='string'||!/^\d{4}-\d{2}-\d{2}$/.test(day))throw Error('日期格式無效');
 const d=new Date(day+'T00:00:00Z');
 if(!Number.isFinite(d.getTime())||d.toISOString().slice(0,10)!==day)throw Error('日期無效');
 return `${Number(day.slice(0,4))-1911}年${day.slice(5,7)}月${day.slice(8,10)}日`;
}
export function emptyShanhuaForm(day,remembered={}){
 const form=structuredClone(SHANHUA_DEFAULTS);
 for(const key of SHANHUA_CARRY_FIELDS){
  const value=remembered?.[key],fallback=form[key];
  if(fallback&&typeof fallback==='object'){
   if(value&&typeof value==='object'&&!Array.isArray(value))for(const field of Object.keys(fallback)){
    if(typeof value[field]===typeof fallback[field]&&(key!=='artifacts'||field!=='status'||['none','sparse','dense'].includes(value[field])))form[key][field]=value[field];
   }
  }else if(typeof value===typeof fallback)form[key]=value;
 }
 form.carried=[...SHANHUA_CARRY_FIELDS,'report'];
 form.day=day;form.docDate=localDay();return validateShanhuaForm(form);
}
export function validateShanhuaForm(form){
 const fail=()=>{throw Error('善化表單資料格式無效');};
 if(!form||typeof form!=='object'||Array.isArray(form))fail();
 for(const key of ['day','docDate','owner','contractor','location','placeName','site','workContent','stratigraphy','features','suggestion','report','mapCaption'])if(typeof form[key]!=='string')fail();
 rocDate(form.day);rocDate(form.docDate);
 for(const key of ['noExcavation','disturbedFill','naturalLayer'])if(typeof form[key]!=='boolean')fail();
 for(const [key,strings] of [['landform',['text']],['culturalLayer',['layers','thickness','from','to']],['artifacts',['status','position','depth']]]){
  const value=form[key];if(!value||typeof value!=='object'||Array.isArray(value))fail();
  for(const field of strings)if(typeof value[field]!=='string')fail();
  if(key!=='artifacts'&&typeof value.checked!=='boolean')fail();
 }
 if(!['none','sparse','dense'].includes(form.artifacts.status))fail();
 if(form.carried!==undefined&&(!Array.isArray(form.carried)||!form.carried.every(key=>SHANHUA_CARRY_FIELDS.includes(key)||key==='report')))fail();
 return form;
}
export function previousWorkContent(day,drafts){
 rocDate(day);
 let previous=null;
 for(const [date,form] of Object.entries(drafts||{})){
  try{rocDate(date);}catch{continue;}
  if(date<day&&typeof form?.workContent==='string'&&form.workContent.trim()&&(!previous||date>previous.day))previous={day:date,workContent:form.workContent};
 }
 return previous;
}
