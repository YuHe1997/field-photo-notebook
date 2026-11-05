// Sequential, DOM/storage-independent gallery import. Read only a small header for common images.
export function sniffImageType(bytes){
  const starts=signature=>signature.every((value,i)=>bytes[i]===value);
  if(starts([0xff,0xd8,0xff]))return 'image/jpeg';
  if(starts([0x89,0x50,0x4e,0x47,0x0d,0x0a,0x1a,0x0a]))return 'image/png';
  if(bytes.length>=12&&starts([0x52,0x49,0x46,0x46])&&[0x57,0x45,0x42,0x50].every((value,i)=>bytes[8+i]===value))return 'image/webp';
  return null;
}

export async function importFiles(files,{workDay,isHeif,sniffImageType:sniff=sniffImageType,canDecode,heifToJpeg,describeFormat,jpegName,save,progress=()=>{},newId,now}){
  let saved=0,converted=0;
  const failed=[],total=files.length;
  for(const [i,file] of files.entries()){
    progress(`正在加入照片（${i+1}／${total}）…`);
    const base={id:newId(),source:'gallery',capturedAt:null,importedAt:now(),fileLastModifiedAt:file.lastModified,location:null,heading:null,sourceExif:null,workDay};
    let item,conversionError,convertedItem=false;
    const heif=await isHeif(file);
    if(heif){
      progress(`正在轉換照片為 JPEG（${i+1}／${total}）…`);
      try{
        const {blob,decoder}=await heifToJpeg(file);
        item={blob,photo:{...base,originalName:jpegName(file.name),mime:'image/jpeg',convertedFrom:{name:file.name,mime:file.type||'image/heif',size:file.size,decoder},captureMethod:`HEIC／HEIF 匯入，已轉為 JPEG（${decoder}）；未保留 EXIF，影像已依檔案方向轉正`,displayRotation:0}};
        convertedItem=true;
      }catch(e){conversionError=e;}
    }else{
      const detected=sniff(new Uint8Array(await file.slice(0,12).arrayBuffer()));
      if(detected||await canDecode(file)){
        const mime=detected||file.type;
        item={blob:file,photo:{...base,originalName:file.name,mime,captureMethod:'原始檔案匯入；未解析 EXIF，預覽由瀏覽器解碼方向',...(mime!=='image/png'?{displayRotation:0}:{})}};
      }
    }
    if(!item){
      failed.push(`${file.name}（${file.type||'未標類型'}，${await describeFormat(file)}${conversionError?'；'+conversionError.message:''}）`);
      continue;
    }
    // Do not read the next file until this individual write has completed.
    try{await save(item);}catch(error){
      const unprocessed=total-i-1;
      return {saved,total,converted,failed,storageError:error,failedIndex:i+1,unprocessed,message:`已存 ${saved}／${total} 張；第 ${i+1} 張儲存失敗：${error.message}；其餘 ${unprocessed} 張未處理，請重試後再選取剩下的照片`};
    }
    saved++;if(convertedItem)converted++;
  }
  let message=`已加入 ${saved} 張，已儲存於本機（未上傳）`;
  if(converted)message+=`；${converted} 張 HEIC 已轉為 JPEG`;
  if(failed.length)message+=`；已匯入 ${saved} 張；${failed.length} 張無法解碼，未加入：${failed.join('、')}`;
  return {saved,total,converted,failed,message};
}
