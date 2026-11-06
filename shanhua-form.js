import {localDay} from './core.js';

export const SHANHUA_PROJECT_NAME='臺南市善化區光復路至文昌路1350PSCP汰換管線工程施工監看計畫';
const rememberedFields=['owner','contractor','location','placeName','site','suggestion'];
export const SHANHUA_DEFAULTS=Object.freeze({
 day:'',docDate:'',owner:'台灣自來水股份有限公司第六區管理處',contractor:'長源工程有限公司',
 location:'',placeName:'',site:'',workContent:'',stratigraphy:'',
 landform:Object.freeze({checked:false,text:''}),noExcavation:false,disturbedFill:false,naturalLayer:false,
 culturalLayer:Object.freeze({checked:false,layers:'',thickness:'',from:'',to:''}),
 artifacts:Object.freeze({status:'none',position:'',depth:''}),features:'無',
 suggestion:'建議持續監看，監看過程中如發現任何涉文化資產標的，將依《文化資產保存法》第33條、第57條、第77條、第88條規定辦理。若有發現疑似考古遺址文化資產時，將即通報主管機關依照「文化資產保存法施行細則第27條及27之1條」辦理。',
 report:'',mapCaption:''
});
export function rocDate(day){
 if(typeof day!=='string'||!/^\d{4}-\d{2}-\d{2}$/.test(day))throw Error('日期格式無效');
 const d=new Date(day+'T00:00:00Z');
 if(!Number.isFinite(d.getTime())||d.toISOString().slice(0,10)!==day)throw Error('日期無效');
 return `${Number(day.slice(0,4))-1911}年${day.slice(5,7)}月${day.slice(8,10)}日`;
}
export function emptyShanhuaForm(day,remembered={}){
 const form=structuredClone(SHANHUA_DEFAULTS);
 for(const key of rememberedFields)if(typeof remembered?.[key]==='string')form[key]=remembered[key];
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
 return form;
}
