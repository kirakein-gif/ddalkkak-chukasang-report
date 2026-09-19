function init(){
  setTimeout(()=>{if(!window.XLSX||!window.ExcelJS||!window.JSZip)$('#dependencyWarning').classList.remove('hidden');},2500);
  $('#fileInput').addEventListener('change',e=>e.target.files[0]&&handleFile(e.target.files[0]));
  const dz=$('#dropZone');['dragenter','dragover'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.add('drag')}));['dragleave','drop'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.remove('drag')}));dz.addEventListener('drop',e=>{const f=e.dataTransfer.files[0];if(f)handleFile(f)});
  $('#monthSelect').addEventListener('change',e=>{state.selectedMonth=e.target.value;renderStats();renderExpense();renderCard();renderGift();});
  $('#resetBtn').addEventListener('click',()=>{state.fileName='';state.sourceRows=[];state.expense=[];state.card=[];state.gift=[];state.months=[];state.selectedMonth='';$('#workArea').classList.add('hidden');$('#fileInput').value='';window.scrollTo({top:0,behavior:'smooth'});});
  $$('.tab').forEach(btn=>btn.addEventListener('click',()=>{$$('.tab').forEach(x=>x.classList.toggle('active',x===btn));$$('.panel').forEach(p=>p.classList.toggle('active',p.id===`panel-${btn.dataset.tab}`));}));
  $('#expenseBoard').addEventListener('change',onEdit);$('#expenseBoard').addEventListener('input',onEdit);$('#giftTable').addEventListener('change',onEdit);
  $$('.preview-btn').forEach(b=>b.addEventListener('click',()=>preview(b.dataset.preview)));$('#closePreview').addEventListener('click',()=>$('#previewDialog').close());
  $$('[data-download]').forEach(b=>b.addEventListener('click',()=>downloadOne(b.dataset.download)));$('#downloadAllBtn').addEventListener('click',downloadAll);
}

document.addEventListener('DOMContentLoaded',init);
