import {zipStore,xmlText} from './docx.js';
import {validDay,workDayOf} from './core.js';
import {validateNuke1Form,nuke1Date,NUKE1_PROJECT_NAME,NUKE1_REMARK,NUKE1_LABELS} from './nuke1-form.js';
const XML='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL='http://schemas.openxmlformats.org/package/2006/relationships';
const namespaces=`xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"`;
const font='<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="標楷體"/><w:sz w:val="24"/></w:rPr>';
function paragraph(value){return `<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r>${font}${value.replace(/\r\n?/g,'\n').split('\n').map(s=>`<w:t xml:space="preserve">${xmlText(s)}</w:t>`).join('<w:br/>')}</w:r></w:p>`;}
// Separate from monitoring draft keys: choices cannot become a prior-day form.
export const nuke1PlacementKey=day=>{nuke1Date(day);return 'field-nuke1-photo-placement-v1-'+day;};
export function nuke1Placements(groups,saved={}){
 return Object.fromEntries(groups.map(g=>[g.id,['8','14','omit'].includes(saved?.[g.id])?saved[g.id]:'14']));
}
// Which Form item a record set goes to: its own `nuke1Slot`; else the Release 54 per-day choice ('8' or the old 'strata' → 8); else 14.
export const NUKE1_SLOTS=Object.freeze(['8','14']);
export function nuke1SlotOf(group,legacy={}){
 if(NUKE1_SLOTS.includes(group?.nuke1Slot))return group.nuke1Slot;
 const old=legacy?.[group?.id];
 return old==='8'||old==='strata'?'8':'14';
}
// The day a record set belongs to: its own workDay, else the day of its first photo (same rule as db.saveGroup).
export function nuke1GroupDay(group,photos=[]){
 if(validDay(group?.workDay))return group.workDay;
 const first=photos.find(p=>p.id===group?.photoIds?.[0]);
 return first?workDayOf(first):null;
}
// The single rule for editor, 02 preview and Word: the Release 54 fallback is looked up by the group's own day, never the selected day.
export function nuke1SlotOfGroup(group,photos,legacyOf){
 const day=nuke1GroupDay(group,photos);
 return nuke1SlotOf(group,day?legacyOf(day):{});
}
export function splitNuke1Groups(groups,placements){
 const choices=nuke1Placements(groups,placements);
 return {strataGroups:groups.filter(g=>choices[g.id]==='8'),relatedGroups:groups.filter(g=>choices[g.id]==='14')};
}
export async function nuke1Docx(form,{strataGroups=[],relatedGroups=[],monitorSignature=null,leaderSignature=null}={}){
 const f=validateNuke1Form(form),media=[],relations=[];
 function picture(image,signature=false){
  if(!(image?.blob instanceof Blob)||!['image/jpeg','image/png'].includes(image.blob.type)||!Number.isFinite(image.width)||image.width<=0||!Number.isFinite(image.height)||image.height<=0)throw Error('圖片格式無效');
  const id=media.length+1,ext=image.blob.type==='image/png'?'png':'jpg',scale=Math.min((signature?1004400:2766154)/image.width,(signature?752400:2074616)/image.height),cx=Math.round(image.width*scale),cy=Math.round(image.height*scale);
  media.push([`word/media/image${id}.${ext}`,image.blob]);relations.push(`<Relationship Id="image${id}" Type="${R}/image" Target="media/image${id}.${ext}"/>`);
  return `<w:r>${font}<w:drawing><wp:inline><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${id}" name="image${id}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="image${id}.${ext}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="image${id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
 }
 const imageParagraph=images=>`<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${images.join('')}</w:p>`;
 function groupContent(groups){
  if(!Array.isArray(groups))throw Error('紀錄集格式無效');
  return groups.map(g=>{
   if(!g||typeof g.engineering!=='string'||!Array.isArray(g.photos))throw Error('紀錄集格式無效');
   let content=g.engineering.trim()?paragraph(g.engineering):'';
   for(let i=0;i<g.photos.length;i+=2)content+=imageParagraph(g.photos.slice(i,i+2).map(p=>picture(p)));
   return content;
  }).join('');
 }
 const widths=[3402,1266,1853,3544];
 function cell(span,content,start=0){return `<w:tc><w:tcPr><w:tcW w:w="${widths.slice(start,start+span).reduce((a,b)=>a+b,0)}" w:type="dxa"/><w:gridSpan w:val="${span}"/></w:tcPr>${content}</w:tc>`;}
 const row=cells=>`<w:tr>${cells.join('')}</w:tr>`;
 function field(key,span=4,start=0){
  const label=NUKE1_LABELS[key]+(['contractor','location','site','workScope','soil','color','advice','otherEvents'].includes(key)?'：':''),separate=['soil','color','remains','artifacts','features','advice'].includes(key);
  return cell(span,separate?paragraph(label)+(f[key]?paragraph(f[key]):''):paragraph(label+f[key]),start);
 }
 const signature=image=>image===null?paragraph(''):imageParagraph([picture(image,true)]);
 const rows=[
  row([cell(4,paragraph('1.工程名稱：'+NUKE1_PROJECT_NAME))]),
  row([field('contractor')]),
  row([field('location',2),field('site',2,2)]),
  row([field('workScope')]),
  row([field('soil',2),field('color',2,2)]),
  row([cell(4,paragraph(NUKE1_LABELS.strata)+(f.strata?paragraph(f.strata):'')+groupContent(strataGroups))]),
  ...['remains','artifacts','features','advice','otherEvents'].map(key=>row([field(key)])),
  row([cell(4,paragraph('14.相關圖示或照片：')+groupContent(relatedGroups))]),
  row([cell(4,paragraph('15.備註：'+NUKE1_REMARK))]),
  row([cell(1,paragraph('監看日期')),cell(2,paragraph('監看人員簽署'),1),cell(1,paragraph('計畫負責人簽署'),3)]),
  row([cell(1,paragraph(nuke1Date(f.day))),cell(2,signature(monitorSignature),1),cell(1,signature(leaderSignature),3)])
 ];
 const borders=['top','left','bottom','right','insideH','insideV'].map(e=>`<w:${e} w:val="single" w:sz="4" w:color="000000"/>`).join('');
 const document=XML+`<w:document ${namespaces}><w:body><w:tbl><w:tblPr><w:tblW w:w="10065" w:type="dxa"/><w:tblInd w:w="-5" w:type="dxa"/><w:tblBorders>${borders}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>${widths.map(w=>`<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${rows.join('')}</w:tbl><w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="851" w:right="849" w:bottom="568" w:left="993" w:header="851" w:footer="340"/></w:sectPr></w:body></w:document>`;
 const types=XML+`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;
 return zipStore([['[Content_Types].xml',new Blob([types])],['_rels/.rels',new Blob([XML+`<Relationships xmlns="${REL}"><Relationship Id="document" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`])],['word/document.xml',new Blob([document])],['word/_rels/document.xml.rels',new Blob([XML+`<Relationships xmlns="${REL}">${relations.join('')}</Relationships>`])],...media]);
}
