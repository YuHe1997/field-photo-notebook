// Pure append-only planning. Never deletes/replaces old document content or changes global styles.
export function documentTabs(document){
 const out=[];function walk(tabs){for(const tab of tabs||[]){if(tab.documentTab)out.push({id:tab.tabProperties?.tabId,title:tab.tabProperties?.title||'文件',data:tab.documentTab});walk(tab.childTabs);}}
 walk(document.tabs);
 if(!out.length&&document.body)out.push({id:null,title:'文件',data:document});
 return out;
}
export function selectedTab(document,id){const tabs=documentTabs(document);const found=id?tabs.find(t=>t.id===id):tabs.length===1?tabs[0]:null;if(!found)throw Error('請明確選擇母檔分頁。');return found;}
function stable(value){if(Array.isArray(value))return value.map(stable);if(value&&typeof value==='object')return Object.fromEntries(Object.keys(value).sort().filter(k=>k!=='contentUri').map(k=>[k,stable(value[k])]));return value;}
export function protectedSnapshot(document,tabId,boundary){
 const tab=selectedTab(document,tabId),tabs=documentTabs(document);
 return JSON.stringify(stable({documentStyle:document.documentStyle,namedStyles:document.namedStyles,otherTabs:tabs.filter(t=>t!==tab&&t.id!==tab.id),tabStyle:tab.data.documentStyle,tabNamedStyles:tab.data.namedStyles,suggestedNamedStylesChanges:tab.data.suggestedNamedStylesChanges,headers:tab.data.headers,footers:tab.data.footers,footnotes:tab.data.footnotes,lists:tab.data.lists,positionedObjects:tab.data.positionedObjects,body:tab.data.body.content.filter(e=>e.endIndex<=boundary),inlineObjects:tab.data.inlineObjects}));
}
// Snapshot only pre-existing inline objects when checking; new objects do not invalidate the prefix.
export function verifyProtected(before,after,tabId,boundary){
 const oldTab=selectedTab(before,tabId),newTab=selectedTab(after,tabId);
 const filtered=structuredClone(after),filteredTab=selectedTab(filtered,tabId);
 if(oldTab.data.inlineObjects){filteredTab.data.inlineObjects=Object.fromEntries(Object.keys(oldTab.data.inlineObjects).map(id=>{if(!newTab.data.inlineObjects?.[id])throw Error('原有照片物件缺失，停止後續寫入。');return [id,newTab.data.inlineObjects[id]];}));}
 else delete filteredTab.data.inlineObjects;
 if(protectedSnapshot(before,tabId,boundary)!==protectedSnapshot(filtered,tabId,boundary))throw Error('原有內容與讀取基線不同，停止後續寫入；不自動刪除或覆寫。');
}
export function appendAnchor(document,tabId){
 const body=selectedTab(document,tabId).data.body?.content;
 const tail=body?.at(-1);
 if(!document.revisionId||!tail?.paragraph||!Number.isInteger(tail.endIndex))throw Error('無法取得母檔版本或安全追加位置。');
 return tail.endIndex-1;
}
export function namedMarker(document,tabId,name){const ranges=selectedTab(document,tabId).data.namedRanges||document.namedRanges||{};return Object.values(ranges).some(group=>group.name===name||group.namedRanges?.some(r=>r.name===name));}
// Some Docs import/merge responses are sparse or retain zero-span covered cells.
// Normalize only against the known DOCX grid (or explicit positive spans), never drop content.
export function logicalTableCells(row,columns,expected){
 const raw=row.tableCells||[],cells=[];
 const isEmpty=cell=>(cell.content||[]).every(n=>n.paragraph&&!n.paragraph.positionedObjectIds?.length&&(n.paragraph.elements||[]).every(e=>e.textRun&&!e.textRun.content?.trim()));
 for(const cell of raw){
  const style=cell.tableCellStyle||{};
  if((style.rowSpan??1)>1)throw Error('目前不支援跨列合併，未追加。');
  if(![0,1].includes(style.rowSpan??1)||style.columnSpan!==undefined&&(!Number.isInteger(style.columnSpan)||style.columnSpan<0||style.columnSpan>columns))throw Error('儲存格跨度資料無效，未追加。');
  if((style.columnSpan===0||style.rowSpan===0)&&!(expected&&raw.length===expected.length)){
   if(!isEmpty(cell))throw Error('合併覆蓋儲存格內仍有內容，未追加，避免遺漏。');
   continue;
  }
  cells.push({cell,span:style.columnSpan??1});
 }
 if(expected?.length===1&&expected[0]===columns&&cells.length===2&&cells[0].span===columns&&isEmpty(cells[1].cell))cells.pop();
 if(expected){
  if(!Array.isArray(expected)||expected.some(s=>s!==1&&s!==2)||expected.reduce((a,b)=>a+b,0)!==columns||expected.length!==cells.length)throw Error(`轉換後表格與當日範本不符（有效儲存格 ${cells.length}、預期 ${expected?.length??'?'}），未追加。`);
  cells.forEach((entry,i)=>{
   // A sole full-width cell is unambiguous only with the original generated DOCX grid.
   if(entry.span!==expected[i]&&entry.span!==0&&!(cells.length===1&&entry.span===1&&expected[i]===columns))throw Error('轉換後合併欄位與當日範本不符，未追加。');
   entry.span=expected[i];
  });
 }
 if(cells.some(e=>!Number.isInteger(e.span)||e.span<1)||cells.reduce((n,e)=>n+e.span,0)!==columns)throw Error(`表格合併結構不完整（${cells.length} 格，欄跨度 ${cells.map(e=>e.span).join('+')}，預期 ${columns}），未追加。`);
 return cells;
}
export function dailyTablePlan(document,{rowLayout}={}){
 const tabs=documentTabs(document);if(tabs.length!==1)throw Error('當日文件必須只有一個分頁。');
 const data=tabs[0].data,body=data.body?.content||[],tables=body.filter(n=>n.table);
 if(tables.length!==1)throw Error('當日文件不是單一監看紀錄表，未追加。');
 for(const e of body){if(!e.table&&!e.paragraph&&!e.sectionBreak)throw Error('當日表格以外有不支援結構，未追加。');if(e.paragraph&&e.paragraph.elements?.some(r=>!r.textRun||r.textRun.content?.trim()))throw Error('當日表格以外有內容，未追加，避免遺漏。');}
 if(Object.keys(data.headers||{}).length||Object.keys(data.footers||{}).length||Object.keys(data.footnotes||{}).length)throw Error('當日文件有頁首頁尾或註腳，未追加，避免遺漏。');
 const source=tables[0].table,columns=source.columns;
 if(columns!==2||!source.tableRows?.length||source.tableRows.length>1000)throw Error('當日表格欄數／列數不支援。');
 if(rowLayout&&rowLayout.length!==source.tableRows.length)throw Error('轉換後表格列數與當日範本不符，未追加。');
 const rows=source.tableRows.map((row,rowIndex)=>{
  let column=0;
  let normalized;
  try{normalized=logicalTableCells(row,columns,rowLayout?.[rowIndex]);}catch(e){throw Error(`第 ${rowIndex+1} 列：${e.message}`);}
  const cells=normalized.map(({cell,span})=>{
   const segments=[],paragraphs=[];let offset=0;
   for(const e of cell.content||[]){
    if(!e.paragraph||e.paragraph.bullet)throw Error('儲存格有巢狀表格／項目清單等不支援結構，未追加。');
    const start=offset;
    for(const run of e.paragraph.elements||[]){
     if(run.textRun){const text=run.textRun.content||'';if(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ue000-\uf8ff]/u.test(text))throw Error('文字含 Docs API 會移除的字元，未追加。');segments.push({kind:'text',text,style:run.textRun.textStyle||{},offset});offset+=text.length;}
     else if(run.inlineObjectElement){
      const object=data.inlineObjects?.[run.inlineObjectElement.inlineObjectId]?.inlineObjectProperties?.embeddedObject;
      const uri=object?.imageProperties?.contentUri;
      if(!uri||!uri.startsWith('https://')||uri.length>2000)throw Error('照片無法取得可插入的暫時網址，未追加。');
      const w=object.size?.width?.magnitude,h=object.size?.height?.magnitude;
      if(!w||!h||object.size.width.unit!=='PT'||object.size.height.unit!=='PT')throw Error('照片尺寸不支援，未追加。');
      const scale=Math.min(1,180/w);segments.push({kind:'image',uri,size:{width:{magnitude:w*scale,unit:'PT'},height:{magnitude:h*scale,unit:'PT'}},offset});offset++;
     }else throw Error('當日文件有不支援的內容，未追加。');
    }
    paragraphs.push({start,end:offset,style:e.paragraph.paragraphStyle||{}});
   }
   const last=segments.at(-1);
   if(last?.kind!=='text'||!last.text.endsWith('\n'))throw Error('儲存格結尾格式不支援，未追加。');
   last.text=last.text.slice(0,-1);offset--;paragraphs.at(-1).end--;
   const result={column,span,segments,paragraphs,length:offset,style:cell.tableCellStyle||{}};column+=span;return result;
  });
  if(column!==columns)throw Error('表格合併結構不完整，未追加。');return cells;
 });
 return {columns,rows,rowStyles:source.tableRows.map(row=>row.tableRowStyle||{})};
}
const loc=(index,tabId)=>({index,...(tabId?{tabId}:{})});
export function structureRequests(anchor,tabId,day,pending,rows){
 const text=`\n監看紀錄 ${day}\n`;
 return [{insertText:{location:loc(anchor,tabId),text}},{updateTextStyle:{range:{startIndex:anchor+1,endIndex:anchor+text.length-1,...(tabId?{tabId}:{})},textStyle:{fontSize:{magnitude:12,unit:'PT'},bold:true},fields:'fontSize,bold'}},{createNamedRange:{name:pending,range:{startIndex:anchor+1,endIndex:anchor+text.length,...(tabId?{tabId}:{})}}},{insertTable:{rows,columns:2,endOfSegmentLocation:tabId?{tabId}:{}}}];
}
export function newTable(document,tabId,anchor){const tables=selectedTab(document,tabId).data.body.content.filter(n=>n.table&&n.startIndex>anchor);if(tables.length!==1)throw Error('新增表格無法唯一辨識，停止，不再寫入。');return tables[0];}
export function mergeRequests(table,tabId,plan){return plan.rows.flatMap((row,rowIndex)=>row.filter(c=>c.span>1).map(c=>({mergeTableCells:{tableRange:{tableCellLocation:{tableStartLocation:loc(table.startIndex,tabId),rowIndex,columnIndex:c.column},rowSpan:1,columnSpan:c.span}}})));}
function pick(object,keys){return Object.fromEntries(keys.filter(k=>object[k]!==undefined).map(k=>[k,object[k]]));}
export function contentRequests(table,tabId,plan,done,anchor,day){
 const flat=[];
 if(table.table.tableRows.length!==plan.rows.length)throw Error('新增列數不符，停止。');
 plan.rows.forEach((row,i)=>{
  const cells=logicalTableCells(table.table.tableRows[i],plan.columns,row.map(c=>c.span));
  row.forEach((source,j)=>{const cell=cells[j].cell,index=cell.content?.[0]?.startIndex;if(!Number.isInteger(index))throw Error('新增儲存格格式不符，停止。');flat.push({source,cell,index});});
 });
 // Descending indices and reversed segments avoid shifting the unfilled cells.
 const requests=[];
 for(const {source,index} of [...flat].sort((a,b)=>b.index-a.index))for(const segment of [...source.segments].reverse()){
  if(segment.kind==='image')requests.push({insertInlineImage:{location:loc(index,tabId),uri:segment.uri,objectSize:segment.size}});
  else if(segment.text)requests.push({insertText:{location:loc(index,tabId),text:segment.text}});
 }
 let shift=0;
 const range=(startIndex,endIndex)=>({startIndex,endIndex,...(tabId?{tabId}:{})});
 for(const {source,cell,index} of [...flat].sort((a,b)=>a.index-b.index)){
  const start=index+shift,end=start+source.length;
  if(source.length){
   requests.push({updateTextStyle:{range:range(start,end),textStyle:{fontSize:{magnitude:12,unit:'PT'},bold:false,italic:false,underline:false,strikethrough:false,foregroundColor:{color:{rgbColor:{red:0,green:0,blue:0}}}},fields:'fontSize,bold,italic,underline,strikethrough,foregroundColor'}});
   for(const segment of source.segments)if(segment.kind==='text'&&segment.text){const style=pick(segment.style,['bold','italic','underline','strikethrough','fontSize','weightedFontFamily','foregroundColor','backgroundColor','baselineOffset']);if(Object.keys(style).length)requests.push({updateTextStyle:{range:range(start+segment.offset,start+segment.offset+segment.text.length),textStyle:style,fields:Object.keys(style).join(',')}});}
  }
  for(const paragraph of source.paragraphs){const style=pick(paragraph.style,['alignment','lineSpacing','spaceAbove','spaceBelow','indentStart','indentEnd','indentFirstLine','keepWithNext','avoidWidowAndOrphan']);if(Object.keys(style).length)requests.push({updateParagraphStyle:{range:range(start+paragraph.start,start+Math.max(paragraph.end,paragraph.start+1)),paragraphStyle:style,fields:Object.keys(style).join(',')}});}
  const style=pick(source.style,['backgroundColor','borderTop','borderBottom','borderLeft','borderRight','paddingTop','paddingBottom','paddingLeft','paddingRight','contentAlignment']);
  if(Object.keys(style).length)requests.push({updateTableCellStyle:{tableRange:{tableCellLocation:{tableStartLocation:loc(table.startIndex,tabId),rowIndex:table.table.tableRows.findIndex(r=>r.tableCells.includes(cell)),columnIndex:source.column},rowSpan:1,columnSpan:source.span},tableCellStyle:style,fields:Object.keys(style).join(',')}});
  shift+=source.length;
 }
 for(const [i,style] of (plan.rowStyles||[]).entries())if(style.minRowHeight)requests.push({updateTableRowStyle:{tableStartLocation:loc(table.startIndex,tabId),rowIndices:[i],tableRowStyle:{minRowHeight:style.minRowHeight},fields:'minRowHeight'}});
 const headingLength=`\n監看紀錄 ${day}\n`.length;
 requests.push({createNamedRange:{name:done,range:range(anchor+1,anchor+headingLength)}});
 if(requests.length>9000)throw Error('當日內容過多，停止追加。');return requests;
}
export function assertAppendOnly(requests,tabId,floor,markerFloor=floor){
 for(const r of requests){
  const keys=Object.keys(r);if(keys.length!==1||!['insertText','insertTable','insertInlineImage','createNamedRange','mergeTableCells','updateTextStyle','updateParagraphStyle','updateTableCellStyle','updateTableRowStyle'].includes(keys[0]))throw Error('拒絕非追加指令。');
  const v=r[keys[0]],where=v.location||v.range||v.tableStartLocation||v.tableRange?.tableCellLocation?.tableStartLocation;
  if(where){if((where.tabId||null)!==(tabId||null))throw Error('拒絕跨分頁寫入。');const start=where.index??where.startIndex;if(!Number.isInteger(start)||start<(keys[0]==='createNamedRange'?markerFloor:floor))throw Error('拒絕修改原有內容。');}
  else if(keys[0]!=='insertTable'||(v.endOfSegmentLocation?.tabId||null)!==(tabId||null))throw Error('寫入位置不明。');
  if(v.fields==='*')throw Error('拒絕全域樣式更新。');
 }
}
