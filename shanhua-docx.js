import {zipStore,xmlText} from './docx.js';
import {validateShanhuaForm,rocDate,SHANHUA_PROJECT_NAME} from './shanhua-form.js';
const XML='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL='http://schemas.openxmlformats.org/package/2006/relationships';
const namespaces=`xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"`;
const font='<w:rPr><w:rFonts w:ascii="DFKai-SB" w:hAnsi="DFKai-SB" w:eastAsia="DFKai-SB"/><w:sz w:val="24"/><w:lang w:val="zh-TW" w:eastAsia="zh-TW"/></w:rPr>';
function paragraph(value){return `<w:p><w:r>${font}${value.replace(/\r\n?/g,'\n').split('\n').map(s=>`<w:t xml:space="preserve">${xmlText(s)}</w:t>`).join('<w:br/>')}</w:r></w:p>`;}
function cell(span,value,merge='',image=''){return `<w:tc><w:tcPr><w:gridSpan w:val="${span}"/>${merge}</w:tcPr>${paragraph(value)}${image}</w:tc>`;}
const row=cells=>`<w:tr>${cells.join('')}</w:tr>`;
function table(columns,rows){const borders=['top','left','bottom','right','insideH','insideV'].map(e=>`<w:${e} w:val="single" w:sz="4"/>`).join('');return `<w:tbl><w:tblPr><w:tblW w:w="9638" w:type="dxa"/><w:tblBorders>${borders}</w:tblBorders><w:tblCellMar><w:top w:w="72" w:type="dxa"/><w:bottom w:w="72" w:type="dxa"/><w:left w:w="72" w:type="dxa"/><w:right w:w="72" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>${Array.from({length:columns},()=>`<w:gridCol w:w="${Math.floor(9638/columns)}"/>`).join('')}</w:tblGrid>${rows.join('')}</w:tbl>`;}
const mark=value=>value?'▓':'⬜';
const blank=value=>value||'＿＿';
export async function shanhuaDocx(form,{map=null,monitorSignature=null,leaderSignature=null}={}){
 const f=validateShanhuaForm(form),media=[],relations=[`<Relationship Id="header" Type="${R}/header" Target="header1.xml"/>`];
 function picture(image,kind){
  if(image===null)return '';
  if(!(image?.blob instanceof Blob)||!['image/jpeg','image/png'].includes(image.blob.type)||!Number.isFinite(image.width)||image.width<=0||!Number.isFinite(image.height)||image.height<=0)throw Error('圖片格式無效');
  const id=media.length+1,ext=image.blob.type==='image/png'?'png':'jpg';
  // Allow for both 72-twip cell margins. A monitor signature has one of
  // four grid columns; the leader has two. Wide images must not widen the table.
  const cellWidth=kind==='map'?9638:Math.floor(9638/4)*(kind==='leader'?2:1);
  const availableWidth=(cellWidth-144)*635;
  const scale=kind==='map'?availableWidth/image.width:Math.min(540000/image.height,availableWidth/image.width);
  const cx=Math.round(image.width*scale),cy=Math.round(image.height*scale);
  media.push([`word/media/image${id}.${ext}`,image.blob]);relations.push(`<Relationship Id="image${id}" Type="${R}/image" Target="media/image${id}.${ext}"/>`);
  return `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${id}" name="${kind}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="image${id}.${ext}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="image${id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
 }
 const c=f.culturalLayer,a=f.artifacts;
 const strata=`地層堆積狀況（土質、土色、堆積型態）\n${f.stratigraphy}\n${mark(f.landform.checked)}地貌${blank(f.landform.text)}\n${mark(f.noExcavation)}無下挖　${mark(f.disturbedFill)}擾亂填覆地層　${mark(f.naturalLayer)}原堆積自然地層\n${mark(c.checked)}原堆積文化層（測繪地層圖）共 ${blank(c.layers)} 層 ${blank(c.thickness)} 公分厚，約地表下 ${blank(c.from)} 公分至 ${blank(c.to)} 公分`;
 const artifacts=`遺物出土狀況\n${mark(a.status==='none')}無遺物出土\n${mark(a.status==='sparse')}零星遺物出土（位置：${a.status==='sparse'?blank(a.position):'＿＿'} 深度：${a.status==='sparse'?blank(a.depth):'＿＿'}）\n${mark(a.status==='dense')}遺物密集出土（位置：${a.status==='dense'?blank(a.position):'＿＿'} 深度：${a.status==='dense'?blank(a.depth):'＿＿'}）`;
 const rows=[
  row([cell(2,'工程計畫'),cell(3,SHANHUA_PROJECT_NAME)]),
  row([cell(2,'開發業主'),cell(1,f.owner),cell(1,'施工單位'),cell(1,f.contractor)]),
  row([cell(1,'地點'),cell(2,f.location),cell(1,'監看日期'),cell(1,rocDate(f.day))]),
  row([cell(1,'小地名'),cell(2,f.placeName),cell(1,'相關遺址'),cell(1,f.site)]),
  row([cell(2,'工程內容（作業項目、範圍與深度）'),cell(3,f.workContent)]),
  row([cell(2,'地層堆積','<w:vMerge w:val="restart"/>'),cell(3,strata)]),
  row([cell(2,'','<w:vMerge/>'),cell(3,artifacts)]),
  row([cell(2,'遺跡、遺物'),cell(3,f.features)]),
  row([cell(2,'相關建議'),cell(3,f.suggestion)]),
  row([cell(2,'陳報資訊'),cell(3,f.report)]),
  row([cell(5,'施工與監看位置圖')]),
  row([cell(5,f.mapCaption,'',picture(map,'map'))])
 ];
 const signatures=[row([cell(4,'監看照片及過程說明')]),row([cell(1,'監看人簽署','',picture(monitorSignature,'monitor')),cell(2,'計畫負責人簽署','',picture(leaderSignature,'leader')),cell(1,`文件製作日期\n${rocDate(f.docDate)}`)])];
 const document=XML+`<w:document ${namespaces}><w:body>${paragraph('監看紀錄表_'+rocDate(f.day))}${table(5,rows)}${paragraph('')}${table(4,signatures)}${paragraph('')}<w:sectPr><w:headerReference w:type="default" r:id="header"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567"/></w:sectPr></w:body></w:document>`;
 const types=XML+`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`;
 return zipStore([
  ['[Content_Types].xml',new Blob([types])],
  ['_rels/.rels',new Blob([XML+`<Relationships xmlns="${REL}"><Relationship Id="document" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`])],
  ['word/document.xml',new Blob([document])],
  ['word/header1.xml',new Blob([XML+`<w:hdr xmlns:w="${W}">${paragraph(rocDate(f.day)+'　_P')}</w:hdr>`])],
  ['word/_rels/document.xml.rels',new Blob([XML+`<Relationships xmlns="${REL}">${relations.join('')}</Relationships>`])],...media
 ]);
}
