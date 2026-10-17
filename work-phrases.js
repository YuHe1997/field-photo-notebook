import {validDay,workDayOf,numberedGroup} from './core.js';
import {normalizeFullwidth} from './mon-numbering.js';

// No phrase bank or new storage: suggestions are derived from committed local sources.
export function readPhraseMetadata(connection){
 return new Promise((resolve,reject)=>{const data={},tx=connection.transaction(['groups','photos'],'readonly');for(const store of ['groups','photos']){const r=tx.objectStore(store).getAll();r.onsuccess=()=>{data[store]=r.result;};}tx.oncomplete=()=>resolve(data);tx.onerror=()=>reject(tx.error||Error('讀取紀錄集失敗'));tx.onabort=()=>reject(tx.error||Error('紀錄集讀取已中止'));});
}
function combine(entries){
 const map=new Map();for(const e of entries){const text=e.text.trim().replace(/\r\n?/g,'\n');if(!text)continue;let p=map.get(text);if(!p){p={text,labels:[],origins:[]};map.set(text,p);}if(!p.labels.includes(e.label))p.labels.push(e.label);p.origins.push(e.origin);}
 return [...map.values()].map(p=>({...p,label:p.labels.join('、'),revision:JSON.stringify(p.origins)}));
}
export function photoEngineeringPhrases(snapshot,day){
 if(!validDay(day))return [];const photos=new Map((snapshot?.photos||[]).map(p=>[p.id,p])),entries=[];
 for(const g of [...(snapshot?.groups||[])].sort((a,b)=>(a.createdAt??a.updatedAt??0)-(b.createdAt??b.updatedAt??0)||String(a.id).localeCompare(String(b.id)))){
  if(typeof g.id!=='string'||typeof g.engineering!=='string'||!Array.isArray(g.photoIds)||!g.photoIds.length||new Set(g.photoIds).size!==g.photoIds.length)continue;
  // An explicit conflicting group day or mixed/missing photo days is not guessed away.
  if(g.workDay!=null&&g.workDay!==day||!g.photoIds.every(id=>{const p=photos.get(id);return p&&(validDay(p.workDay)||Number.isFinite(p.buttonPressedAt??p.capturedAt??p.importedAt??p.createdAt))&&workDayOf(p)===day;}))continue;
  entries.push({text:g.engineering,label:numberedGroup(g.photoIds,snapshot.photos),origin:{id:g.id,day,photoIds:g.photoIds,raw:g.engineering}});
 }
 return combine(entries);
}
const compound=/^([ \t]*[0-9０-９]+[ \t]*[.．]?[ \t]*[（(][0-9０-９]+[）)][ \t]*[:：]?[ \t]*)(.*)$/;
const child=/^([ \t]*[（(][0-9０-９]+[）)][ \t]*[:：]?[ \t]*)(.*)$/;
const main=/^[ \t]*([0-9０-９]+)(?:[.．](?![0-9０-９])|[:：])[ \t]*(.*)$/;
export function workScopePhrases(value){
 if(typeof value!=='string')return [];const entries=[];let active=null,parent=null;
 const flush=()=>{if(active)entries.push({text:active.lines.join('\n'),label:active.label,origin:{line:active.line,label:active.label}});active=null;};
 value.replace(/\r\n?/g,'\n').split('\n').forEach((line,index)=>{
  const comp=line.match(compound),sub=line.match(child);
  if(comp||sub){flush();const m=comp||sub,norm=normalizeFullwidth(m[1]).trim().replace(/[:：]$/u,'').trim();if(comp){parent=norm.match(/^([0-9]+)/)?.[1]??null;}active={line:index,label:comp?norm:parent?`${parent}.${norm}`:norm,lines:[m[2]]};return;}
  const heading=line.match(main);if(heading){flush();parent=normalizeFullwidth(heading[1]);return;}
  if(active)active.lines.push(line);
 });flush();return combine(entries);
}
export function readSavedScopePhrases(storage,day){
 if(!validDay(day)||!storage)throw Error('無法讀取本工作日已儲存草稿');const raw=storage.getItem('field-mon-form-'+day);
 if(raw===null)return {choices:[],revision:null};let d;try{d=JSON.parse(raw);}catch{throw Error('監看草稿格式異常，未讀取或覆寫；請先處理草稿問題');}
 if(!d||Array.isArray(d)||typeof d!=='object'||d.workScope!==undefined&&typeof d.workScope!=='string')throw Error('監看草稿格式異常，未讀取或覆寫');
 return {choices:workScopePhrases(d.workScope||''),revision:raw};
}
export function phraseInsertion(value,start,end,phrase,{scope=false}={}){
 if(typeof value!=='string'||typeof phrase!=='string'||!phrase.trim())throw Error('片語內容為空，未插入');
 if(!Number.isInteger(start)||!Number.isInteger(end)||start<0||end<start||end>value.length)throw Error('游標位置已變更，請重新選擇');
 const text=phrase.trim().replace(/\r\n?/g,'\n');if(start===end&&value.slice(0,start).endsWith(text))throw Error('游標前已有相同片語，未重複插入');
 if(!scope)return {start,end:start,text};
 if(!value.trim())return {start,end:start,text:'1.(1) '+text};
 const lineStart=start===0?0:value.lastIndexOf('\n',start-1)+1,next=value.indexOf('\n',start),lineEnd=next<0?value.length:next,line=value.slice(lineStart,lineEnd);
 const match=line.match(compound)||line.match(child);
 if(match){const contentStart=lineStart+match[1].length;if(end>lineEnd||end>start&&start<contentStart)throw Error('選取範圍包含編號或多行，請先點在子項目文字位置');return {start:Math.max(start,contentStart),end:Math.max(start,contentStart),text};}
 if(main.test(value)&&!value.includes('\n')&&!value.match(main)[2].trim())return {start:value.length,end:value.length,text:'\n(1) '+text};
 // Continuation text belongs to the preceding child, unless a parent heading intervenes.
 let inChild=false;for(const previous of value.slice(0,lineStart).split('\n')){if(compound.test(previous)||child.test(previous))inChild=true;else if(main.test(previous))inChild=false;}
 if(inChild&&end<=lineEnd)return {start,end:start,text};
 throw Error('請先按「新增子項目」，或把游標放在要填寫的子項目文字位置，再選片語');
}

// UI adapter. Every load is user-initiated, and sources are re-read before applying.
export class WorkPhrasePicker{
 constructor({panel,search,list,status,refresh,undo,target,getContext,loadSource,apply,restore,successMessage}){
  Object.assign(this,{panel,search,list,status,refresh,undo,target,getContext,loadSource,apply,restore,successMessage});this.sequence=0;this.busy=false;this.source=null;this.lastInsert=null;this.composing=false;
  panel.addEventListener('toggle',()=>{if(panel.open)this.reload();else this.invalidate(false);});refresh.onclick=()=>this.reload();search.addEventListener('input',()=>this.render());search.addEventListener('keydown',e=>{if(e.key==='Enter')e.preventDefault();});undo.onclick=()=>this.undoInsert();
  target.addEventListener('compositionstart',()=>{this.composing=true;});target.addEventListener('compositionend',()=>{this.composing=false;});target.addEventListener('input',()=>this.updateUndo());
 }
 contextMatches(c){const now=this.getContext();return !!c&&!!now&&c.key===now.key;}
 message(s){this.status.textContent=s;}
 invalidate(close=true){this.sequence++;this.source=null;this.lastInsert=null;this.busy=false;this.list.replaceChildren();this.updateUndo();if(close)this.panel.open=false;}
 async reload(){
  const seq=++this.sequence,c=this.getContext();this.busy=false;this.source=null;this.updateUndo();this.list.replaceChildren();if(!c){this.message('請先完成工作日切換或開啟記事。');return;}this.message('正在讀取本工作日已儲存內容…');
  try{const source=await this.loadSource(c);if(seq!==this.sequence||!this.panel.open||!this.contextMatches(c))return;this.source={...source,context:c};this.message(source.choices.length?`工作日 ${c.day} · ${source.choices.length} 段片語；相同文字合併顯示。`:source.emptyMessage||'本工作日尚無已儲存的片語來源。');this.render();}
  catch(e){if(seq===this.sequence&&this.contextMatches(c)){this.message('片語讀取未完成：'+e.message);this.list.replaceChildren();}}
 }
 render(){
  this.list.replaceChildren();if(!this.source)return;const q=this.search.value.trim().toLocaleLowerCase(),all=this.source.choices.filter(p=>(p.text+' '+p.label).toLocaleLowerCase().includes(q));
  for(const choice of all.slice(0,50)){const button=document.createElement('button');button.type='button';button.className='phrase-choice';button.disabled=this.busy;const body=document.createElement('span'),origin=document.createElement('small');body.textContent=choice.text.length>180?choice.text.slice(0,180)+'…':choice.text;origin.textContent=choice.label+(choice.text.length>180?' · 長段落，點選帶入全文':'');button.append(body,origin);button.title=choice.text;button.onclick=e=>{if(e.detail>1)return;this.pick(choice);};this.list.append(button);}
  if(all.length>50||!all.length){const p=document.createElement('p');p.textContent=all.length>50?`符合 ${all.length} 段，目前顯示前 50 段；可輸入關鍵字縮小範圍。`:'沒有符合搜尋的片語。';this.list.append(p);}
 }
 capture(){return {value:this.target.value,start:this.target.selectionStart,end:this.target.selectionEnd};}
 unchanged(t){return this.target.value===t.value&&this.target.selectionStart===t.start&&this.target.selectionEnd===t.end;}
 async pick(choice){
  if(this.busy||!this.source)return;if(this.composing){this.message('請先完成輸入法組字，再選片語。');return;}const source=this.source,c=source.context,seq=this.sequence,t=this.capture();if(!this.contextMatches(c)){this.message('工作日或記事已變更，請重新整理片語。');return;}
  this.busy=true;this.render();this.message('確認來源仍有效…');
  try{
   const fresh=await this.loadSource(c);if(seq!==this.sequence||!this.panel.open||!this.contextMatches(c))return;
   if(fresh.revision!==source.revision||!fresh.choices.some(p=>p.text===choice.text&&p.revision===choice.revision)){await this.reload();if(this.contextMatches(c))this.message('來源已變更，片語已更新；未插入，請重新選取。');return;}
   if(this.composing||!this.unchanged(t))throw Error('文字或游標已變更，未插入；請重新選擇');
   const plan=this.apply(choice.text,t);this.lastInsert={context:c,before:t,after:this.target.value,start:plan.start,length:this.target.value.length-t.value.length};this.updateUndo();this.message(this.successMessage());
  }catch(e){if(seq===this.sequence&&this.contextMatches(c))this.message('未插入：'+e.message);}
  finally{if(seq===this.sequence){this.busy=false;this.updateUndo();this.render();}}
 }
 updateUndo(){const x=this.lastInsert;this.undo.hidden=!x;this.undo.disabled=!x||this.busy||!this.contextMatches(x.context)||this.target.value!==x.after;}
 undoInsert(){
  const x=this.lastInsert;if(!x||this.busy||this.composing)return;if(!this.contextMatches(x.context)||this.target.value!==x.after){this.message('已有後續編輯或已切換工作日，未復原；請用欄位原生復原或手動調整。');return;}
  try{this.restore(x);this.lastInsert=null;this.updateUndo();this.message('已復原本次片語插入；來源內容未改動。');}catch(e){this.message('復原未完成：'+e.message);}
 }
}
