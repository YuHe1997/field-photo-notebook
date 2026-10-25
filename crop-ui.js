// Simple crop dialog. The crop window always keeps the photo's own shape (4:3 stays 4:3): one slider shrinks it
// towards the centre, dragging (or the arrow buttons) slides it, so a bystander at an edge can be left out.
import {renderPhotoCanvas} from './photo-render.js';
import {validCrop,cropWindow,moveCrop,normalizeCrop,cropSummary,MIN_CROP} from './core.js';
const $=id=>document.getElementById(id);
// Resolves to a crop object, null (use the whole photo) or undefined (cancelled).
export async function openCropDialog({photo,original}){
  const dialog=$('cropDialog'),canvas=$('cropCanvas'),slider=$('cropScale'),url=URL.createObjectURL(original),im=new Image();
  let base;
  try{im.src=url;await im.decode();base=renderPhotoCanvas(im,{...photo,crop:null},{maxSide:900,applyCrop:false});}
  finally{URL.revokeObjectURL(url);im.src='';}
  canvas.width=base.width;canvas.height=base.height;
  const ctx=canvas.getContext('2d'),maxShrink=Math.round((1-MIN_CROP)*100);
  let crop=validCrop(photo.crop)?{...photo.crop}:cropWindow(1);
  slider.max=String(maxShrink);
  const draw=()=>{
    const W=canvas.width,H=canvas.height,x=crop.x*W,y=crop.y*H,w=crop.w*W,h=crop.h*H;
    ctx.clearRect(0,0,W,H);ctx.drawImage(base.canvas,0,0);
    ctx.fillStyle='rgba(0,0,0,.55)';ctx.fillRect(0,0,W,y);ctx.fillRect(0,y+h,W,H-y-h);ctx.fillRect(0,y,x,h);ctx.fillRect(x+w,y,W-x-w,h);
    ctx.strokeStyle='rgba(255,255,255,.35)';ctx.lineWidth=1;ctx.beginPath();for(const t of [1/3,2/3]){ctx.moveTo(x+w*t,y);ctx.lineTo(x+w*t,y+h);ctx.moveTo(x,y+h*t);ctx.lineTo(x+w,y+h*t);}ctx.stroke();
    ctx.strokeStyle='#ffd54a';ctx.lineWidth=Math.max(2,W/300);ctx.strokeRect(x,y,w,h);
    const shrink=Math.round((1-crop.w)*100);slider.value=String(shrink);
    $('cropAmount').textContent=`${shrink}%`;slider.setAttribute('aria-valuetext',`往內縮 ${shrink}%`);
    $('cropDetail').textContent=cropSummary(normalizeCrop(crop),base.fullWidth,base.fullHeight);
  };
  const setScale=value=>{crop=cropWindow(1-value/100,crop.x+crop.w/2,crop.y+crop.h/2);draw();};
  let drag=null;
  const point=e=>{const r=canvas.getBoundingClientRect();return {x:(e.clientX-r.left)/r.width,y:(e.clientY-r.top)/r.height};};
  const onDown=e=>{drag={at:point(e),start:{...crop}};canvas.setPointerCapture?.(e.pointerId);e.preventDefault();};
  const onMove=e=>{if(!drag)return;const p=point(e);crop=moveCrop(drag.start,p.x-drag.at.x,p.y-drag.at.y);draw();};
  const onUp=()=>{drag=null;};
  const nudge=(dx,dy)=>()=>{crop=moveCrop(crop,dx,dy);draw();};
  const handlers=[[canvas,'pointerdown',onDown],[canvas,'pointermove',onMove],[canvas,'pointerup',onUp],[canvas,'pointercancel',onUp],
    [slider,'input',()=>setScale(Number(slider.value))],[$('cropLeft'),'click',nudge(-.01,0)],[$('cropRight'),'click',nudge(.01,0)],[$('cropUp'),'click',nudge(0,-.01)],[$('cropDown'),'click',nudge(0,.01)],
    [$('cropReset'),'click',()=>{crop=cropWindow(1);draw();}]];
  for(const [el,type,fn] of handlers)el.addEventListener(type,fn);
  draw();
  return new Promise(resolve=>{
    const finish=value=>{
      for(const [el,type,fn] of handlers)el.removeEventListener(type,fn);
      $('cropSave').onclick=$('cropCancel').onclick=null;dialog.removeEventListener('cancel',onCancel);
      if(dialog.open)dialog.close();canvas.width=canvas.height=1;resolve(value);
    };
    const onCancel=e=>{e.preventDefault();finish(undefined);};
    dialog.addEventListener('cancel',onCancel);
    $('cropCancel').onclick=()=>finish(undefined);
    $('cropSave').onclick=()=>finish(normalizeCrop(crop));
    dialog.showModal();slider.focus();
  });
}
