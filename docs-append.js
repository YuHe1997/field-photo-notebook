// Insert-only planning. Existing records are never deleted, replaced or globally restyled.
export const DOCS_APPEND_BUILD='date-order-v1';
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
// Reuse the last recognizable daily-log title in the selected tab, not a company name in code.
export function masterHeadingTemplate(document,tabId){
 const data=selectedTab(document,tabId).data,body=data.body?.content||[];
 const textOf=node=>node?.paragraph&&(node.paragraph.elements||[]).every(e=>e.textRun)?node.paragraph.elements.map(e=>e.textRun.content||'').join(''):null;
 const linesOf=text=>text.replace(/\u000b/g,'\n').trim().split('\n').map(line=>line.trim()).filter(Boolean);
 for(let end=body.length-1;end>=0;end--){
  const last=textOf(body[end]);if(!last||linesOf(last).at(-1)!=='監看工作日誌')continue;
  // A title must belong to a following table; do not copy an unrelated note elsewhere.
  let tableFollows=null;for(let i=end+1;i<Math.min(body.length,end+6);i++){if(body[i].table){tableFollows=body[i];break;}const text=textOf(body[i]);if(text===null||text.trim())break;}
  if(!tableFollows)continue;
  const nodes=[];let lineCount=0;
  for(let i=end;i>=Math.max(0,end-5);i--){const text=textOf(body[i]);if(text===null)break;if(!text.trim())continue;nodes.unshift(body[i]);lineCount+=linesOf(text).length;if(lineCount>=3)break;}
  const all=nodes.map(textOf).join(''),lines=linesOf(all);
  if(lines.length!==3||lines[2]!=='監看工作日誌'||!lines[1].includes('監看計畫')||all.length>1500)continue;
  const stripped=all.match(/[\u0000-\u0008\u000c-\u001f\ue000-\uf8ff]/u);if(stripped)throw Error('母檔標題有不支援的字元，未追加。');
  const named=(data.namedStyles||document.namedStyles)?.styles||[];
  const normal=named.find(style=>style.namedStyleType==='NORMAL_TEXT')||{};
  const paragraphs=nodes.map(node=>{
   const p=node.paragraph,style=p.paragraphStyle||{},base=named.find(s=>s.namedStyleType===style.namedStyleType)||normal;
   return {text:textOf(node),paragraphStyle:{...normal.paragraphStyle,...base.paragraphStyle,...style},textStyle:{fontSize:{magnitude:12,unit:'PT'},bold:false,italic:false,underline:false,strikethrough:false,foregroundColor:{color:{rgbColor:{red:0,green:0,blue:0}}},...normal.textStyle,...base.textStyle},runs:p.elements.map(e=>({text:e.textRun.content||'',style:e.textRun.textStyle||{}}))};
  });
  return {text:'\n'+all,paragraphs,lines,startIndex:nodes[0].startIndex,tableStartIndex:tableFollows.startIndex};
 }
 return null;
}
export function headingRangeText(document,tabId,start,end){
 let result='';for(const node of selectedTab(document,tabId).data.body.content){if(!node.paragraph)continue;let cursor=node.startIndex;for(const run of node.paragraph.elements||[]){if(!run.textRun)continue;const text=run.textRun.content||'',at=run.startIndex??cursor,last=at+text.length;if(last>start&&at<end)result+=text.slice(Math.max(0,start-at),Math.min(text.length,end-at));cursor=last;}}
 return result;
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
     if(run.textRun){
      const text=run.textRun.content||'';
      // Docs preserves U+000B soft line breaks. Its documented stripped ranges start at U+000C,
      // not U+000B. Keep soft breaks verbatim so UTF-16 indices and paragraph styles stay intact.
      const removed=text.match(/[\u0000-\u0008\u000c-\u001f\ue000-\uf8ff]/u);
      if(removed){const code='U+'+removed[0].codePointAt(0).toString(16).toUpperCase().padStart(4,'0');throw Error(`第 ${rowIndex+1} 列第 ${column+1} 欄含 Docs API 會移除的字元（${code}），未追加；原文字保留。`);}
      segments.push({kind:'text',text,style:run.textRun.textStyle||{},offset});offset+=text.length;
     }
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
export function structureRequests(anchor,tabId,day,pending,rows,heading=null,placement=null){
 const text=heading?.text||`\n監看紀錄 ${day}\n`,requests=[{insertText:{location:loc(anchor,tabId),text}}];
 const range=(startIndex,endIndex)=>({startIndex,endIndex,...(tabId?{tabId}:{})});
 if(heading){
  let offset=1;
  for(const paragraph of heading.paragraphs){
   const start=anchor+offset,end=start+paragraph.text.length;
   const pStyle=pick(paragraph.paragraphStyle,['namedStyleType','alignment','lineSpacing','spaceAbove','spaceBelow','indentStart','indentEnd','indentFirstLine','keepWithNext','avoidWidowAndOrphan','pageBreakBefore']);
   if(!pStyle.alignment)pStyle.alignment='CENTER';
   requests.push({updateParagraphStyle:{range:range(start,end),paragraphStyle:pStyle,fields:Object.keys(pStyle).join(',')}});
   const textStyle=pick(paragraph.textStyle,['bold','italic','underline','strikethrough','fontSize','weightedFontFamily','foregroundColor','backgroundColor','baselineOffset']);
   requests.push({updateTextStyle:{range:range(start,end),textStyle,fields:Object.keys(textStyle).join(',')}});
   let runOffset=0;for(const run of paragraph.runs){const style=pick(run.style,['bold','italic','underline','strikethrough','fontSize','weightedFontFamily','foregroundColor','backgroundColor','baselineOffset']);if(run.text.length&&Object.keys(style).length)requests.push({updateTextStyle:{range:range(start+runOffset,start+runOffset+run.text.length),textStyle:style,fields:Object.keys(style).join(',')}});runOffset+=run.text.length;}
   offset+=paragraph.text.length;
  }
 }else requests.push({updateTextStyle:{range:range(anchor+1,anchor+text.length-1),textStyle:{fontSize:{magnitude:12,unit:'PT'},bold:true},fields:'fontSize,bold'}});
 requests.push({createNamedRange:{name:pending,range:range(anchor+1,anchor+text.length)}},{insertTable:{rows,columns:2,...(placement?.mode==='before'?{location:loc(anchor+text.length,tabId)}:{endOfSegmentLocation:tabId?{tabId}:{}})}});
 return requests;
}
export function newTable(document,tabId,anchor){const tables=selectedTab(document,tabId).data.body.content.filter(n=>n.table&&n.startIndex>anchor);if(tables.length!==1)throw Error('新增表格無法唯一辨識，停止，不再寫入。');return tables[0];}
export function mergeRequests(table,tabId,plan){return plan.rows.flatMap((row,rowIndex)=>row.filter(c=>c.span>1).map(c=>({mergeTableCells:{tableRange:{tableCellLocation:{tableStartLocation:loc(table.startIndex,tabId),rowIndex,columnIndex:c.column},rowSpan:1,columnSpan:c.span}}})));}
function pick(object,keys){return Object.fromEntries(keys.filter(k=>object[k]!==undefined).map(k=>[k,object[k]]));}
export function contentRequests(table,tabId,plan,done,anchor,day,heading=null){
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
 const headingLength=(heading?.text||`\n監看紀錄 ${day}\n`).length;
 requests.push({createNamedRange:{name:done,range:range(anchor+1,anchor+headingLength)}});
 if(requests.length>9000)throw Error('當日內容過多，停止追加。');return requests;
}
export function assertAppendOnly(requests,tabId,floor,markerFloor=floor){
 for(const r of requests){
  const keys=Object.keys(r);if(keys.length!==1||!['insertText','insertTable','insertInlineImage','createNamedRange','mergeTableCells','updateTextStyle','updateParagraphStyle','updateTableCellStyle','updateTableRowStyle'].includes(keys[0]))throw Error('拒絕非追加指令。');
  const v=r[keys[0]],where=v.location||v.range||v.tableStartLocation||v.tableRange?.tableCellLocation?.tableStartLocation;
  if(where){if((where.tabId||null)!==(tabId||null))throw Error('拒絕跨分頁寫入。');const start=where.index??where.startIndex;if(!Number.isInteger(start)||start<(keys[0]==='createNamedRange'?markerFloor:floor))throw Error('拒絕修改原有內容。');}
  else if(keys[0]!=='insertTable'||(v.endOfSegmentLocation?.tabId||null)!==(tabId||null))throw Error('寫入位置不明。');
  if(v.fields?.split(',').some(f=>f.trim()==='*'))throw Error('拒絕全域樣式更新。');
 }
}

// Dates are explicit calendar dates, never inferred from upload time, document title or workday.
export function parseMonitoringDate(value){
 if(typeof value!=='string')throw Error('請填寫可辨識的監看日期。');
 const text=value.normalize('NFKC').trim();
 const m=text.match(/^(民國\s*)?(\d{1,4})\s*(?:年|[-/.])\s*(\d{1,2})\s*(?:月|[-/.])\s*(\d{1,2})\s*日?(?:\s*(?:(?:星期|週)[一二三四五六日天]|\([一二三四五六日天]\)))?$/u);
 if(!m)throw Error('監看日期無法辨識；請使用西元／民國年月日，不依上傳時間猜日期。');
 let year=Number(m[2]);if(m[1]||m[2].length<=3){if(year<1||year>300)throw Error('民國年格式無效。');year+=1911;}
 const month=Number(m[3]),day=Number(m[4]),date=new Date(Date.UTC(year,month-1,day));
 if(year<1912||year>9999||date.getUTCFullYear()!==year||date.getUTCMonth()+1!==month||date.getUTCDate()!==day)throw Error('監看日期不是有效日期。');
 return `${year}-${String(month).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
}
function cellText(cell,{strict=false}={}){
 let text='';for(const node of cell.content||[]){if(!node.paragraph){if(strict)throw Error('文字欄位含巢狀表格，暫不套用。');continue;}
 for(const run of node.paragraph.elements||[]){if(run.textRun)text+=run.textRun.content||'';else if(strict)throw Error('文字欄位含非文字內容，暫不套用。');}}
 return text.replace(/\n$/,'');
}
const labelPattern=label=>new RegExp('^\\s*'+[...label].map(c=>c.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')).join('\\s*')+'\\s*[:：][ \\t]?','u');
function fieldValue(cell,label){const text=cellText(cell,{strict:true}),m=text.match(labelPattern(label));if(!m)throw Error(`母表最後表格缺少「${label}」欄位，未套用。`);return text.slice(m[0].length).replace(/\u000b/g,'\n');}
export function lastMonitoringTemplate(document,tabId){
 const tables=selectedTab(document,tabId).data.body?.content?.filter(n=>n.table)||[],node=tables.at(-1);
 if(!node||node.table.columns!==2||node.table.tableRows?.length<7)throw Error('母表最後表格不是完整的兩欄監看紀錄，未套用。');
 const shapes=[[1,1],[1,1],[2],[2],[1,1],[1,1],[1,1]],rows=shapes.map((shape,i)=>logicalTableCells(node.table.tableRows[i],2,shape).map(c=>c.cell));
 const fields={projectName:fieldValue(rows[0][0],'工程名稱'),date:fieldValue(rows[0][1],'日期'),constructionLocation:fieldValue(rows[1][0],'施工地點'),siteName:fieldValue(rows[1][1],'考古遺址'),watcherCount:fieldValue(rows[2][0],'監看人數'),workScope:fieldValue(rows[3][0],'作業項目、範圍與深度'),strataSummary:fieldValue(rows[4][0],'地層堆積'),strataDesc:cellText(rows[4][1],{strict:true}).replace(/\u000b/g,'\n'),culturalArtifacts:fieldValue(rows[5][0],'文化遺物'),culturalFeatures:fieldValue(rows[5][1],'文化遺跡'),otherEvents:fieldValue(rows[6][0],'其他重要紀事'),dispositionAdvice:fieldValue(rows[6][1],'文資處置建議')};
 return {fields,date:parseMonitoringDate(fields.date),tableStartIndex:node.startIndex};
}
export function sourceMonitoringDate(document){
 const tabs=documentTabs(document);if(tabs.length!==1)throw Error('來源分頁不唯一，未寫入。');
 const table=tabs[0].data.body?.content?.find(n=>n.table)?.table,dates=(table?.tableRows||[]).slice(0,7).flatMap(r=>r.tableCells||[]).filter(c=>labelPattern('日期').test(cellText(c)));
 if(dates.length!==1)throw Error('來源監看日期無法唯一辨識，未寫入。');return parseMonitoringDate(fieldValue(dates[0],'日期'));
}
export function monitoringRecords(document,tabId){
 const body=selectedTab(document,tabId).data.body?.content;if(!Array.isArray(body))throw Error('無法讀取母表內容。');
 const records=[];
 for(const node of body){if(!node.table)continue;
  const first=node.table.tableRows?.[0]?.tableCells||[],dates=first.filter(c=>labelPattern('日期').test(cellText(c))),projects=first.filter(c=>labelPattern('工程名稱').test(cellText(c)));
  if(!dates.length&&!projects.length)continue; // Front matter and continuation/photo tables stay with their existing record.
  if(dates.length!==1||projects.length!==1)throw Error('母表有工程／日期欄位不完整的表格，請先核對。');
  if(!Number.isInteger(node.startIndex)||!Number.isInteger(node.endIndex))throw Error('母表表格位置不完整。');
  records.push({date:parseMonitoringDate(fieldValue(dates[0],'日期')),tableStartIndex:node.startIndex,tableEndIndex:node.endIndex});
 }
 return records;
}
function rangeNames(document,tabId){
 const names=new Set();for(const groups of [document.namedRanges,selectedTab(document,tabId).data.namedRanges])for(const group of Object.values(groups||{})){if(group.name)names.add(group.name);for(const r of group.namedRanges||[])if(r.name)names.add(r.name);}return names;
}
export function chronologicalPlacement(document,tabId,value){
 if(typeof document.revisionId!=='string'||!document.revisionId)throw Error('無法取得母表版本，未插入。');
 const date=parseMonitoringDate(value),body=selectedTab(document,tabId).data.body?.content||[],names=rangeNames(document,tabId);
 for(const name of names)if(name.startsWith('fpn_pending_')&&!names.has(name.replace('fpn_pending_','fpn_done_')))throw Error('母表有未確認完成的操作標記，請先核對；不再新增。');
 let previousEnd=0;for(const node of body){const start=node.startIndex??(node.sectionBreak?0:NaN);if(!Number.isInteger(start)||!Number.isInteger(node.endIndex)||node.endIndex<=start||start<previousEnd)throw Error('母表段落索引不完整或重疊，未插入。');previousEnd=node.endIndex;}
 const records=monitoringRecords(document,tabId);
 if(records.some(r=>r.date===date))throw Error(`母表已有 ${date} 的紀錄，停止，避免同日重複；不覆寫。`);
 // Find a boundary that places this record after EVERY earlier date and before EVERY later date.
 // Existing records are never reordered. Historic disorder is accepted only if this slot is unique.
 const slots=[];for(let i=0;i<=records.length;i++)if(records.slice(0,i).every(r=>r.date<date)&&records.slice(i).every(r=>r.date>date))slots.push(i);
 if(slots.length!==1)throw Error('母表日期次序無法唯一界定插入位置，請先核對；不搬動原有紀錄。');
 const next=records[slots[0]];
 if(!next){return {mode:'end',date,anchor:appendAnchor(document,tabId),nextDate:null};}
 const tableIndex=body.findIndex(n=>n.startIndex===next.tableStartIndex&&n.table),prefix=structuredClone(document);selectedTab(prefix,tabId).data.body.content=structuredClone(body.slice(0,tableIndex+1));
 const title=masterHeadingTemplate(prefix,tabId);let index=tableIndex-1;
 if(title?.tableStartIndex===next.tableStartIndex)index=body.findIndex(n=>n.startIndex===title.startIndex);
 else if(!body[index]?.paragraph||!(body[index].paragraph.elements||[]).every(r=>r.textRun)||(!/^\s*監看紀錄\s+\d{4}-\d{2}-\d{2}\s*$/u.test(cellParagraphText(body[index]))&&cellParagraphText(body[index]).trim()))throw Error('較晚日期之前沒有安全的標題／空白段落，未插入。');
 if(index>0&&body[index-1].paragraph&&cellParagraphText(body[index-1])==='\n')index--;
 const node=body[index];if(!node?.paragraph||node.paragraph.bullet||node.paragraph.positionedObjectIds?.length||!Number.isInteger(node.startIndex)||node.startIndex<1||!Number.isInteger(node.endIndex)||node.endIndex<=node.startIndex)throw Error('無法取得較晚紀錄前的安全段落位置。');
 return {mode:'before',date,anchor:node.startIndex,nextDate:next.date,nextTableStartIndex:next.tableStartIndex};
}
function cellParagraphText(node){return node?.paragraph&&(node.paragraph.elements||[]).every(r=>r.textRun)?node.paragraph.elements.map(r=>r.textRun.content||'').join(''):'';}
function shiftBodyIndices(value,delta){if(Array.isArray(value))return value.map(v=>shiftBodyIndices(v,delta));if(value&&typeof value==='object')return Object.fromEntries(Object.entries(value).map(([k,v])=>[k,(k==='startIndex'||k==='endIndex')&&Number.isInteger(v)?v+delta:shiftBodyIndices(v,delta)]));return value;}
export function insertionExtent(before,after,tabId,placement){
 const a=selectedTab(before,tabId).data.body.content,b=selectedTab(after,tabId).data.body.content,delta=b.at(-1)?.endIndex-a.at(-1)?.endIndex;
 if(!Number.isInteger(delta)||delta<=0)throw Error('新增範圍無法確認，停止後續寫入。');return {start:placement.anchor,end:placement.anchor+delta,delta};
}
export function verifyInsertionProtected(before,after,tabId,placement){
 if(placement.mode==='end'){
  const tail=selectedTab(before,tabId).data.body.content.at(-1),empty=cellParagraphText(tail)==='\n'&&!tail.paragraph.positionedObjectIds?.length;return verifyProtected(before,after,tabId,empty?tail.startIndex:placement.anchor+1);
 }
 const {start,end,delta}=insertionExtent(before,after,tabId,placement),clone=structuredClone(after),body=selectedTab(clone,tabId).data.body.content;
 if(body.some(n=>(n.startIndex<start&&n.endIndex>start)||(n.startIndex<end&&n.endIndex>end)))throw Error('新增範圍跨到原有段落，停止後續寫入。');
 selectedTab(clone,tabId).data.body.content=body.filter(n=>n.endIndex<=start).concat(body.filter(n=>n.startIndex>=end).map(n=>shiftBodyIndices(n,-delta)));
 verifyProtected(before,clone,tabId,Infinity);
}
export function insertedRecordTable(before,after,tabId,placement){
 if(placement.mode==='end')return newTable(after,tabId,placement.anchor);
 const {start,end}=insertionExtent(before,after,tabId,placement),tables=selectedTab(after,tabId).data.body.content.filter(n=>n.table&&n.startIndex>start&&n.endIndex<=end);
 if(tables.length!==1)throw Error('新增日期表格無法唯一辨識，停止後續寫入。');return tables[0];
}
export function assertRecordInsertion(requests,tabId,{placement,headingText,table=null,plan=null}){
 const anchor=placement.anchor,titleEnd=anchor+headingText.length;
 assertAppendOnly(requests,tabId,table?table.startIndex:anchor,anchor+1);
 const max=table?table.endIndex+plan.rows.flat().reduce((n,c)=>n+c.length,0):titleEnd;
 for(const request of requests){const [kind]=Object.keys(request),v=request[kind];
  if(kind==='insertTable'){
   if(table||v.rows!==plan.rows.length||v.columns!==2)throw Error('拒絕新增範圍外的表格指令。');
   if(placement.mode==='before'?(v.location?.index!==titleEnd||v.endOfSegmentLocation):(!v.endOfSegmentLocation||v.location))throw Error('拒絕不明表格插入位置。');continue;
  }
  if(kind==='createNamedRange'){if(v.range.startIndex!==anchor+1||v.range.endIndex!==titleEnd)throw Error('拒絕範圍外的完成標記。');continue;}
  if(!table&&kind==='insertText'){if(v.location.index!==anchor||v.text!==headingText)throw Error('拒絕標題範圍外的文字插入。');continue;}
  const range=v.range;if(range){if(!Number.isInteger(range.endIndex)||range.startIndex<(table?table.startIndex:anchor+1)||range.endIndex>max||range.endIndex<=range.startIndex)throw Error('拒絕修改原有紀錄範圍。');}
  else if(v.location){if(!table||v.location.index<table.startIndex||v.location.index>=table.endIndex)throw Error('拒絕在原有紀錄插入內容。');}
  else{const at=v.tableStartLocation||v.tableRange?.tableCellLocation?.tableStartLocation;if(!table||at?.index!==table.startIndex)throw Error('拒絕修改原有表格。');}
 }
}
