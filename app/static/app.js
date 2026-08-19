const oldFile=document.querySelector('#old-file'),newFile=document.querySelector('#new-file');
const analyse=document.querySelector('#analyse'),generate=document.querySelector('#generate');
let analysis=null;
oldFile.addEventListener('change',()=>document.querySelector('#old-name').textContent=oldFile.files[0]?.name||'Word .docx, up to 20 MB');
newFile.addEventListener('change',()=>document.querySelector('#new-name').textContent=newFile.files[0]?.name||'Use a completed example if available');
function setBusy(button,busy,label){button.disabled=busy;button.textContent=busy?label:button.dataset.label}
analyse.dataset.label=analyse.textContent;generate.dataset.label=generate.textContent;
analyse.addEventListener('click',async()=>{
  const error=document.querySelector('#error');error.textContent='';
  if(!oldFile.files[0]||!newFile.files[0]){error.textContent='Choose both Word documents first.';return}
  const form=new FormData();form.append('old_document',oldFile.files[0]);form.append('new_document',newFile.files[0]);
  setBusy(analyse,true,'Analysing…');
  try{const response=await fetch('/api/analyse',{method:'POST',body:form});const data=await response.json();if(!response.ok)throw new Error(data.detail||'Analysis failed.');analysis=data;renderReview();}
  catch(e){error.textContent=e.message}finally{setBusy(analyse,false)}
});
function humanise(value){return value.replaceAll('_',' ').replace(/\b\w/g,c=>c.toUpperCase())}
function renderReview(){
  document.querySelector('#detected').innerHTML=analysis.detected.map(x=>`<span class="chip">${humanise(x)} detected</span>`).join('')||'<span class="missing">No common named fields detected</span>';
  const options=analysis.old_fields.map(f=>`<option value="${f.id}">${f.label}</option>`).join('');
  document.querySelector('#mapping-body').innerHTML=analysis.mappings.map(m=>`<tr data-target="${m.target_id}"><td><strong>${m.target_label}</strong></td><td><select><option value="">Not mapped / enter manually</option>${options}</select></td><td><div class="preview"></div><textarea placeholder="Missing information — enter a value manually if required"></textarea></td></tr>`).join('');
  analysis.mappings.forEach(m=>{const row=document.querySelector(`[data-target="${m.target_id}"]`),select=row.querySelector('select'),preview=row.querySelector('.preview');select.value=m.source_id||'';const update=()=>{const f=analysis.old_fields.find(x=>x.id===select.value);preview.textContent=f?f.value.slice(0,280):'No source selected';preview.className=f?'preview':'preview missing'};select.addEventListener('change',update);update()});
  document.querySelector('#review').hidden=false;document.querySelector('#review').scrollIntoView({behavior:'smooth'});
}
generate.addEventListener('click',async()=>{
  const error=document.querySelector('#generate-error');error.textContent='';const mappings={},manual={};
  document.querySelectorAll('#mapping-body tr').forEach(row=>{const id=row.dataset.target;mappings[id]=row.querySelector('select').value||null;const value=row.querySelector('textarea').value.trim();if(value)manual[id]=value});
  const form=new FormData();form.append('session_id',analysis.session_id);form.append('mappings',JSON.stringify(mappings));form.append('manual_values',JSON.stringify(manual));setBusy(generate,true,'Generating…');
  try{const response=await fetch('/api/generate',{method:'POST',body:form});if(!response.ok){const data=await response.json();throw new Error(data.detail||'Generation failed.')}const blob=await response.blob(),url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download='converted-assignment-brief.docx';a.click();URL.revokeObjectURL(url);generate.textContent='Assignment successfully converted — download again';generate.dataset.label=generate.textContent;}
  catch(e){error.textContent=e.message}finally{setBusy(generate,false)}
});

