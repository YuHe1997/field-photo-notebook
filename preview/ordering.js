// Finite additive ranks precede legacy rows; milliseconds never share the rank scale.
export function compareGroups(a,b){
 const ar=Number.isFinite(a?.sortOrder),br=Number.isFinite(b?.sortOrder);
 if(ar!==br)return ar?-1:1;
 if(ar&&a.sortOrder!==b.sortOrder)return a.sortOrder-b.sortOrder;
 const time=g=>Number.isFinite(g?.createdAt)?g.createdAt:Number.isFinite(g?.updatedAt)?g.updatedAt:0;
 return time(a)-time(b)||String(a?.id??'').localeCompare(String(b?.id??''));
}
