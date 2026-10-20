// Preview loader: runs the released app.js with the redesign's small edits applied in memory.
// The service worker pins ../app.js to its released hash, so any drift fails loudly.
const root=new URL('../',import.meta.url).href,here=new URL('./',import.meta.url).href;
let src=await (await fetch(new URL('../app.js',import.meta.url))).text();
const edits=[
 ["document.querySelectorAll('[data-tab]').forEach(b=>{b.onclick=()=>tab(b.dataset.tab);","document.querySelectorAll('[data-tab]').forEach(b=>{b.onclick=()=>{tab(b.dataset.tab);if(b.dataset.tab!=='editor')window.scrollTo(0,0);};"],
 ["li.style.cssText='display:flex;justify-content:space-between;align-items:center;padding:8px 12px;border-bottom:1px solid #eef1eb';","li.className='folder-row';"],
 ["   nameSpan.style.cssText='font-weight:500;overflow-wrap:anywhere';\n",""],
 ["   btn.style.cssText='min-height:36px;padding:4px 8px;font-size:0.85rem';\n",""],
 ["import {AREA_DATA} from './area-data.js';\n","import {AREA_DATA} from './area-data.js';\nimport {initFormFit} from './form-fit.js';\n"],
 ["const $=id=>document.getElementById(id);\n","const $=id=>document.getElementById(id);\ninitFormFit();\n"],
 ["$('closeCoordinateAlbum').onclick=()=>closeCoordinateAlbum();","$('viewPhotoCoordinatesPhrase').onclick=$('viewPhotoCoordinates').onclick;\n$('closeCoordinateAlbum').onclick=()=>closeCoordinateAlbum();"],
 ["for(const id of ['viewPhotoCoordinates',","for(const id of ['viewPhotoCoordinates','viewPhotoCoordinatesPhrase',"]
];
// The phrase panel gets its own copy of the reference-photo button (same dialog, same handler).
const refBtn=document.getElementById('viewPhotoCoordinates'),refreshBtn=document.getElementById('refreshPhotoPhrases');
if(refBtn&&refreshBtn&&!document.getElementById('viewPhotoCoordinatesPhrase')){const copy=refBtn.cloneNode(true);copy.id='viewPhotoCoordinatesPhrase';refreshBtn.after(copy);}
for(const [from,to] of edits){if(!src.includes(from))throw Error('preview patch failed: '+from.slice(0,48));src=src.replace(from,()=>to);}
src=src.replace(/from '\.\/([\w-]+\.js)'/g,(_,file)=>`from '${file==='form-fit.js'?here:root}${file}'`);
await import(URL.createObjectURL(new Blob([src],{type:'text/javascript'})));
