// One place that turns a stored original into a displayable/exportable canvas: applies the saved display rotation
// and (optionally) the saved crop. Shared by previews, Word/Docs export and Drive upload.
import {previewRotation,validCrop} from './core.js';
export function renderPhotoCanvas(im,photo,{maxSide=960,applyCrop=true,background=null}={}){
  const w=im.naturalWidth||im.width,h=im.naturalHeight||im.height,r=previewRotation(photo,w,h),swapped=r%180!==0,ow=swapped?h:w,oh=swapped?w:h;
  const crop=applyCrop&&validCrop(photo?.crop)?photo.crop:null,cw=ow*(crop?crop.w:1),ch=oh*(crop?crop.h:1),scale=Math.min(1,maxSide/Math.max(cw,ch));
  const canvas=document.createElement('canvas');canvas.width=Math.max(1,Math.round(cw*scale));canvas.height=Math.max(1,Math.round(ch*scale));
  const ctx=canvas.getContext('2d');if(!ctx)throw Error('無法建立影像畫布');
  if(background){ctx.fillStyle=background;ctx.fillRect(0,0,canvas.width,canvas.height);}
  if(crop)ctx.translate(ow*scale/2-crop.x*ow*scale,oh*scale/2-crop.y*oh*scale);else ctx.translate(canvas.width/2,canvas.height/2);
  ctx.rotate(r*Math.PI/180);ctx.drawImage(im,-w*scale/2,-h*scale/2,w*scale,h*scale);
  return {canvas,rotation:r,width:canvas.width,height:canvas.height,fullWidth:ow,fullHeight:oh};
}
