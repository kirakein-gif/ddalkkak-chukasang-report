function init(){
  setTimeout(()=>{if(!window.XLSX||!window.ExcelJS||!window.JSZip)$('#dependencyWarning').classList.remove('hidden');},2500);
  $('#fileInput').addEventListener('click',e=>{e.currentTarget.value='';});
  $('#fileInput').addEventListener('change',e=>e.target.files[0]&&handleFile(e.target.files[0]));
  const dz=$('#dropZone');['dragenter','dragover'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.add('drag')}));['dragleave','drop'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.remove('drag')}));dz.addEventListener('drop',e=>{const f=e.dataTransfer.files[0];if(f)handleFile(f)});
  const statTabs=document.querySelector('.stat-tab-grid');
  if(statTabs){
    statTabs.addEventListener('click',e=>{
      const btn=e.target.closest('.stat-tab'); if(!btn||btn.classList.contains('active')) return;
      const work=$('#workArea');
      const beforeHeight=work.offsetHeight;
      const beforeY=window.scrollY;
      work.style.minHeight=Math.max(parseFloat(work.dataset.lockHeight||0),beforeHeight)+'px';
      work.dataset.lockHeight=String(Math.max(parseFloat(work.dataset.lockHeight||0),beforeHeight));
      document.querySelectorAll('.stat-tab').forEach(x=>x.classList.toggle('active',x===btn));
      document.querySelectorAll('.panel').forEach(p=>p.classList.toggle('active',p.id===`panel-${btn.dataset.tab}`));
      requestAnimationFrame(()=>{
        const grown=Math.max(parseFloat(work.dataset.lockHeight||0),work.offsetHeight);
        work.dataset.lockHeight=String(grown); work.style.minHeight=grown+'px';
        window.scrollTo({top:beforeY,left:0,behavior:'auto'});
      });
    });
  }
  $('#expenseBoard').addEventListener('change',onEdit);$('#expenseBoard').addEventListener('input',onEdit);$('#expenseBoard').addEventListener('click',onExpenseAction);$('#giftTable').addEventListener('change',onEdit);
  $$('.preview-btn').forEach(b=>b.addEventListener('click',()=>preview(b.dataset.preview)));$('#closePreview').addEventListener('click',()=>$('#previewDialog').close());
  document.querySelectorAll('[data-download]').forEach(b=>b.addEventListener('click',()=>downloadOne(b.dataset.download)));$('#downloadAllBtn').addEventListener('click',downloadAll);
  $('#versionCheckBtn').addEventListener('click',e=>{ if(e.currentTarget.dataset.updateAvailable==='1') reloadLatestVersion(); else checkLatestVersion(true); });
  setTimeout(()=>checkLatestVersion(false),700);
}

document.addEventListener('DOMContentLoaded',init);
