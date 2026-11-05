// HEIC/HEIF (Samsung "高效率照片", iPhone) import: convert to JPEG once, at import time, so every later step
// (preview, crop, Word/Docs export, Drive upload, backup restore on another phone) gets a format all browsers decode.
// Try the browser's own decoder first (Safari); otherwise load the bundled libheif WebAssembly decoder (Chrome/Android).
const HEIF_TYPES=new Set(['image/heic','image/heif','image/heic-sequence','image/heif-sequence']);
const HEIF_BRANDS=new Set(['heic','heix','heim','heis','hevc','hevx','heif']);
const DECODER_URL=new URL('./vendor/libheif/libheif-bundle.esm.js',import.meta.url).href;

// Pure check on the first bytes: ISO-BMFF "ftyp" box whose major brand (or, for the generic mif1/msf1, any compatible brand) is HEIF/HEVC.
// AVIF also uses ftyp/mif1 but browsers decode it, so it is not treated as HEIF.
export function heifBrand(bytes){
  if(!bytes||bytes.length<12)return false;
  const tag=i=>String.fromCharCode(bytes[i],bytes[i+1],bytes[i+2],bytes[i+3]);
  if(tag(4)!=='ftyp')return false;
  const major=tag(8);if(HEIF_BRANDS.has(major))return true;
  if(major!=='mif1'&&major!=='msf1')return false;
  const size=Math.min(bytes.length,(bytes[0]<<24|bytes[1]<<16|bytes[2]<<8|bytes[3])>>>0);
  for(let i=16;i+4<=size;i+=4)if(HEIF_BRANDS.has(tag(i)))return true;
  return false;
}
export async function isHeif(file){
  if(HEIF_TYPES.has((file.type||'').toLowerCase())||/\.(heic|heif|hif)$/i.test(file.name||''))return true;
  // Always look at the bytes: some Android pickers label HEIC as image/jpeg, .jpg or an empty type.
  return heifBrand(new Uint8Array(await file.slice(0,64).arrayBuffer()));
}
// Short, value-free description of a file's real format, for error messages (e.g. "ftyp heic", "JPEG").
export async function describeFormat(file){
  const b=new Uint8Array(await file.slice(0,32).arrayBuffer()),s=(i,n)=>String.fromCharCode(...b.slice(i,i+n));
  if(b[0]===0xff&&b[1]===0xd8)return 'JPEG';if(s(1,3)==='PNG')return 'PNG';if(s(0,4)==='RIFF'&&s(8,4)==='WEBP')return 'WebP';
  if(s(4,4)==='ftyp')return 'ftyp '+s(8,4).replace(/[^\x20-\x7e]/g,'?');if(s(0,2)==='II'||s(0,2)==='MM')return 'TIFF/RAW';
  return b.length?'未知格式':'空檔案';
}
// True when this browser itself can decode the image.
export async function canDecode(blob){
  const url=URL.createObjectURL(blob),im=new Image();
  try{im.src=url;await im.decode();return im.naturalWidth>0;}catch{return false;}finally{im.src='';URL.revokeObjectURL(url);}
}
export function jpegName(name){return String(name||'photo').replace(/\.(heic|heif|hif|jpe?g)$/i,'')+'.jpg';}

function canvasToJpeg(canvas){
  return new Promise((resolve,reject)=>canvas.toBlob(b=>b?resolve(b):reject(Error('無法轉存為 JPEG（照片可能過大）')),'image/jpeg',.92));
}
async function nativeDecode(file){
  const url=URL.createObjectURL(file);
  try{
    const im=new Image();im.src=url;await im.decode();
    if(!im.naturalWidth)throw Error('empty');
    const c=document.createElement('canvas');c.width=im.naturalWidth;c.height=im.naturalHeight;
    const ctx=c.getContext('2d');if(!ctx)throw Error('無法建立影像畫布');ctx.drawImage(im,0,0);
    return c;
  }finally{URL.revokeObjectURL(url);}
}
let libheifPromise=null;
function loadLibheif(){
  libheifPromise??=import(DECODER_URL).then(m=>m.default()).catch(e=>{libheifPromise=null;throw e;});
  return libheifPromise;
}
async function wasmDecode(file){
  let libheif;
  try{libheif=await loadLibheif();}catch{throw Error('HEIC 解碼程式載入失敗；請連線後重新開啟 App 再匯入');}
  const decoder=new libheif.HeifDecoder(),images=decoder.decode(new Uint8Array(await file.arrayBuffer()));
  try{
    // libheif returns the top-level images; the first one is the photo (rotation/mirroring from the file is already applied).
    const image=images[0];if(!image)throw Error('HEIC 檔內沒有影像');
    const width=image.get_width(),height=image.get_height();
    const c=document.createElement('canvas');c.width=width;c.height=height;
    const ctx=c.getContext('2d');if(!ctx)throw Error('無法建立影像畫布（照片可能過大）');
    const data=ctx.createImageData(width,height);
    await new Promise((resolve,reject)=>image.display(data,out=>out?resolve():reject(Error('HEIC 解碼失敗'))));
    ctx.putImageData(data,0,0);
    return c;
  }finally{for(const im of images)im.free?.();if(decoder.decoder)libheif.heif_context_free?.(decoder.decoder);decoder.decoder=null;}
}
// Returns {blob, decoder} where blob is a JPEG of the upright photo.
export async function heifToJpeg(file){
  let canvas,decoder='瀏覽器內建解碼';
  try{canvas=await nativeDecode(file);}catch{canvas=await wasmDecode(file);decoder='libheif（App 內建）';}
  const blob=await canvasToJpeg(canvas);canvas.width=canvas.height=0;
  return {blob,decoder};
}
