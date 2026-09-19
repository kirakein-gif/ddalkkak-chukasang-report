'use strict';

const EXPENSE_GROUPS = ['회의ㆍ간담회','경조사','물품구입','위문ㆍ격려ㆍ직원사기진작','각종회비'];
const CARD_GROUPS = ['건당 50만원이상 업무추진비','건당 100만원 이상 지출건 중 업무추진비 성격 이외의 경비'];
const state = { fileName:'', sourceRows:[], expense:[], card:[], gift:[], months:[], selectedMonth:'' };
let expenseSortables = [];

const $ = (s, root=document) => root.querySelector(s);
const $$ = (s, root=document) => [...root.querySelectorAll(s)];
const won = n => Number(n || 0).toLocaleString('ko-KR') + '원';
const num = n => Number(n || 0).toLocaleString('ko-KR');
const esc = s => String(s ?? '').replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));

function toast(msg,actionLabel='',action=null){
  const el=$('#toast'); el.innerHTML='';
  const span=document.createElement('span'); span.textContent=msg; el.appendChild(span);
  if(actionLabel&&typeof action==='function'){
    const btn=document.createElement('button'); btn.type='button'; btn.className='toast-action'; btn.textContent=actionLabel;
    btn.addEventListener('click',()=>{ action(); el.classList.remove('show'); }); el.appendChild(btn);
  }
  el.classList.add('show'); clearTimeout(toast.t); toast.t=setTimeout(()=>el.classList.remove('show'),action?4500:2400);
}
function parseAmount(v){ if(typeof v==='number' && Number.isFinite(v)) return v; const n=Number(String(v??'').replace(/[원,\s]/g,'')); return Number.isFinite(n)?n:0; }
function normalizeText(v){ return String(v ?? '').trim(); }
function normalizeDate(v){
  if(v instanceof Date && !isNaN(v)) return new Date(v.getFullYear(),v.getMonth(),v.getDate());
  if(typeof v==='number' && window.XLSX?.SSF){ const d=XLSX.SSF.parse_date_code(v); if(d) return new Date(d.y,d.m-1,d.d); }
  const s=normalizeText(v); if(!s) return null;
  const m=s.match(/(20\d{2})[.\-/년\s]+(\d{1,2})[.\-/월\s]+(\d{1,2})/);
  if(m) return new Date(+m[1],+m[2]-1,+m[3]);
  const d=new Date(s); return isNaN(d)?null:d;
}
function dateText(d){ return d ? `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}` : ''; }
function monthKey(d){ return d ? `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}` : ''; }
function monthLabel(k){ if(!k) return ''; const [y,m]=k.split('-'); return `${y}년 ${Number(m)}월`; }
function sameMonth(d,k){ return !k || monthKey(d)===k; }
function sortDate(a,b){ return (a.date?.getTime()||0)-(b.date?.getTime()||0); }

function classifyExpense(detail){
  const t=normalizeText(detail).toLowerCase();
  if(t.includes('위문')||t.includes('격려')||t.includes('사기진작')) return '위문ㆍ격려ㆍ직원사기진작';
  if(t.includes('경조사')||t.includes('축의금')) return '경조사';
  if(t.includes('구입')) return '물품구입';
  if(t.includes('회의')||t.includes('협의회')||t.includes('간담회')) return '회의ㆍ간담회';
  return '각종회비';
}
function normalizePayment(v){ const t=normalizeText(v); if(t.includes('카드')) return '법인카드'; if(t.includes('이체')) return '계좌이체'; return t; }
function classifyCard(statItem){ const t=normalizeText(statItem).toLowerCase(); return (t.includes('업무추진비')||t.includes('협의회')||t.includes('간담회')) ? CARD_GROUPS[0] : CARD_GROUPS[1]; }
function maskKoreanName(name){
  const t=normalizeText(name);
  if(!/^[가-힣]{2,4}$/.test(t)) return '';
  const compound=['남궁','황보','제갈','선우','사공','서문','독고'];
  const surname=compound.find(x=>t.startsWith(x) && t.length>x.length) || t[0];
  return surname + '○'.repeat(Math.max(1,t.length-surname.length));
}
function publicVendor(row){
  const raw=normalizeText(row.vendor);
  if(row.category!=='경조사') return raw;
  if(row.privacyMode==='raw') return raw;
  if(row.privacyMode==='staff') return raw ? '해당교직원' : '';
  const masked=maskKoreanName(raw);
  if(masked) return masked;
  if(raw && row.payment==='계좌이체') return '해당교직원';
  return raw;
}
function publicText(text,row){
  const raw=normalizeText(row.vendor), replacement=publicVendor(row);
  let out=normalizeText(text);
  if(row.category==='경조사' && raw && replacement && raw!==replacement) out=out.split(raw).join(replacement);
  return out;
}
function inferUnitPrice(detail){
  const t=normalizeText(detail).replace(/\s/g,'');
  if(/5천원|5,000원|5000원/.test(t)) return 5000;
  if(/1만원|10,000원|10000원/.test(t)) return 10000;
  if(/3만원|30,000원|30000원/.test(t)) return 30000;
  if(/5만원|50,000원|50000원/.test(t)) return 50000;
  if(/10만원|100,000원|100000원/.test(t)) return 100000;
  return 10000;
}
function inferQuantity(detail){ const m=normalizeText(detail).match(/(\d{1,6})\s*(매|장|개)/); return m?Number(m[1]):null; }

function findHeaderRow(rows){
  const required=['일자','제목','원인행위액','목','원가통계비목','수령인','지급방법'];
  for(let i=0;i<Math.min(rows.length,20);i++){
    const cells=(rows[i]||[]).map(normalizeText);
    if(required.every(h=>cells.includes(h))) return i;
  }
  return -1;
}
function headerMap(header){
  const cells=header.map(normalizeText);
  const find=(...names)=>{ for(const name of names){ const i=cells.indexOf(name); if(i>=0)return i; } return -1; };
  return {
    date:find('일자'), detail:find('제목','세부내역'), amount:find('원인행위액','금액'),
    expenseType:find('목'), statItem:find('원가통계비목'), vendor:find('수령인','거래처','장소'),
    payment:find('지급방법','결제방법')
  };
}
function extractData(rows){
  const expense=[], card=[], gift=[];
  const headerIndex=findHeaderRow(rows);
  if(headerIndex<0) throw new Error('K-에듀파인 예산거래처별실적 형식을 찾지 못했습니다. “일자·제목·원인행위액·목·원가통계비목·수령인·지급방법” 열을 확인해 주세요.');
  const col=headerMap(rows[headerIndex]||[]);
  let dataRows=0;
  for(let i=headerIndex+1;i<rows.length;i++){
    const r=rows[i]||[]; if(r.every(v=>normalizeText(v)==='')) continue;
    const detail=normalizeText(r[col.detail]);
    if(detail==='합계') continue;
    dataRows++;
    const date=normalizeDate(r[col.date]); const amount=parseAmount(r[col.amount]);
    const expenseType=normalizeText(r[col.expenseType]); const statItem=normalizeText(r[col.statItem]);
    const vendor=normalizeText(r[col.vendor]); const payment=normalizeText(r[col.payment]);
    if(statItem==='일반업무추진비') expense.push({id:`e${i}`,category:classifyExpense(detail),date,detail,target:'',vendor,amount,payment:normalizePayment(payment),privacyMode:'auto',excluded:false});
    if(payment.toLowerCase().includes('카드')){
      const isBusiness=expenseType==='업무추진비'; const eligible=(isBusiness&&amount>=500000)||(!isBusiness&&amount>=1000000);
      if(eligible) card.push({id:`c${i}`,category:classifyCard(statItem),date,detail,amount,statItem,expenseType});
    }
    if(detail.toLowerCase().includes('상품권')){
      const unitPrice=inferUnitPrice(detail); const parsedQty=inferQuantity(detail); const qty=parsedQty ?? Math.max(1,Math.ceil(amount/unitPrice));
      gift.push({id:`g${i}`,date,detail,vendor,unitPrice,quantity:qty,amount,qtyConfirmed:parsedQty!==null});
    }
  }
  return {expense,card,gift,headerIndex,dataRows};
}

function giftDerived(g){ const face=(Number(g.unitPrice)||0)*(Number(g.quantity)||0); const rate=face>0?(face-(Number(g.amount)||0))/face:0; return {face,rate}; }
function filtered(list){ return list.filter(x=>sameMonth(x.date,state.selectedMonth)); }
function current(){ return {expense:filtered(state.expense).filter(x=>!x.excluded),card:filtered(state.card),gift:filtered(state.gift)}; }

async function handleFile(file){
  if(!window.XLSX){ $('#dependencyWarning').classList.remove('hidden'); return; }
  try{
    const buf=await file.arrayBuffer();
    const wb=XLSX.read(buf,{type:'array',cellDates:true,raw:true});
    const ws=wb.Sheets[wb.SheetNames[0]];
    const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:true});
    if(rows.length<2) throw new Error('데이터 행이 없습니다.');
    const out=extractData(rows);
    state.fileName=file.name; state.sourceRows=rows; state.expense=out.expense; state.card=out.card; state.gift=out.gift;
    const months=[...new Set([...out.expense,...out.card,...out.gift].map(x=>monthKey(x.date)).filter(Boolean))].sort();
    state.months=months; state.selectedMonth=months[0]||'';
    renderAll();
    $('#workArea').classList.remove('hidden');
    $('#fileName').textContent=file.name; $('#sourceInfo').textContent=`${out.dataRows}개 원본 거래 · K-에듀파인 형식 확인`;
    window.scrollTo({top:$('#workArea').offsetTop-12,behavior:'smooth'}); toast('원본 분석이 완료되었습니다.');
  }catch(err){ console.error(err); alert(`파일을 읽지 못했습니다.\n${err.message||err}`); }
}

function renderMonthSelect(){
  const sel=$('#monthSelect'); sel.innerHTML='';
  if(!state.months.length){ sel.innerHTML='<option value="">날짜 없음</option>'; return; }
  state.months.forEach(k=>{ const o=document.createElement('option'); o.value=k; o.textContent=monthLabel(k); o.selected=k===state.selectedMonth; sel.appendChild(o); });
  const warn=$('#monthWarning');
  if(state.months.length>1){ warn.textContent=`원본에 ${state.months.length}개월 데이터가 섞여 있습니다. 기준월을 선택하면 해당 월만 보고서에 반영됩니다.`; warn.classList.remove('hidden'); }
  else warn.classList.add('hidden');
}
function renderStats(){
  const c=current(), review=c.gift.filter(g=>!g.qtyConfirmed).length;
  $('#expenseCount').textContent=`${c.expense.length}건`; $('#cardCount').textContent=`${c.card.length}건`; $('#giftCount').textContent=`${c.gift.length}건`; $('#reviewCount').textContent=`${review}건`;
  $('#expenseBadge').textContent=c.expense.length; $('#cardBadge').textContent=c.card.length; $('#giftBadge').textContent=c.gift.length;
}
function renderExpenseBoard(){
  const board=$('#expenseBoard'); if(!board) return;
  expenseSortables.forEach(x=>{ try{x.destroy();}catch(_){} }); expenseSortables=[];

  const monthRows=filtered(state.expense);
  const list=monthRows.filter(x=>!x.excluded);
  const excluded=monthRows.filter(x=>x.excluded).sort(sortDate);
  const total=list.reduce((s,x)=>s+x.amount,0);

  const groupsHtml=EXPENSE_GROUPS.map((g,idx)=>{
    const arr=list.filter(x=>x.category===g).sort(sortDate);
    const amount=arr.reduce((s,x)=>s+x.amount,0);
    const ratio=total?((amount/total)*100).toFixed(1):'0.0';

    const items=arr.map(x=>{
      const privacy=x.category==='경조사'
        ? `<div class="expense-field privacy-field">
             <label>개인정보 공개</label>
             <select class="cell-select compact" data-id="${x.id}" data-field="privacyMode">
               <option value="auto" ${x.privacyMode==='auto'?'selected':''}>자동 보호 · ${esc(publicVendor(x)||'-')}</option>
               <option value="staff" ${x.privacyMode==='staff'?'selected':''}>해당교직원</option>
               <option value="raw" ${x.privacyMode==='raw'?'selected':''}>원문 유지</option>
             </select>
           </div>`
        : '';

      return `
        <article class="expense-drag-item" data-id="${x.id}">
          <button type="button" class="drag-handle" aria-label="분류 이동" title="잡아서 위·아래로 이동">☰</button>
          <div class="expense-row-content">
            <div class="expense-main-line">
              <span class="expense-date">${dateText(x.date)}</span>
              <span class="expense-detail">${esc(x.detail)}</span>
              <strong class="expense-amount">${won(x.amount)}</strong>
            </div>
            <div class="expense-sub-line">
              <div class="expense-field target-field">
                <label>집행대상자</label>
                <input class="cell-input compact" type="text" placeholder="집행대상자 입력" value="${esc(x.target)}" data-id="${x.id}" data-field="target">
              </div>
              <div class="expense-field vendor-field">
                <label>장소/수령인${x.category==='경조사'?'(원본)':''}</label>
                <span>${esc(x.vendor)||'-'}</span>
              </div>
              <div class="expense-field payment-field">
                <label>결재방법</label>
                <span>${esc(x.payment)||'-'}</span>
              </div>
              ${privacy}
            </div>
          </div>
          <details class="item-menu">
            <summary aria-label="항목 메뉴" title="항목 메뉴">⋮</summary>
            <div class="item-menu-pop">
              <button type="button" data-expense-action="exclude" data-id="${x.id}">보고서에서 제외</button>
              <small>오추출·실수 수정용</small>
            </div>
          </details>
        </article>`;
    }).join('');

    return `
      <section class="expense-drop-group group-${idx}">
        <header>
          <div class="expense-group-name"><span class="order-no">${idx+1}</span><strong>${esc(g)}</strong></div>
          <div class="expense-group-summary"><span>${arr.length}건</span><span>${won(amount)}</span><b>${ratio}%</b></div>
        </header>
        <div class="expense-dropzone" data-category="${esc(g)}">${items || '<div class="empty-drop">이 구분의 내역이 없습니다 · 여기로 끌어다 놓을 수 있습니다</div>'}</div>
      </section>`;
  }).join('');

  const excludedHtml=excluded.length ? `
    <details class="excluded-panel">
      <summary>제외된 항목 <b>${excluded.length}건</b><span>원본은 변경되지 않습니다</span></summary>
      <div class="excluded-list">
        ${excluded.map(x=>`
          <div class="excluded-item">
            <div>
              <span>${dateText(x.date)}</span>
              <strong>${esc(x.detail)}</strong>
              <small>${esc(x.category)} · ${won(x.amount)}</small>
            </div>
            <button type="button" class="restore-btn" data-expense-action="restore" data-id="${x.id}">복원</button>
          </div>`).join('')}
      </div>
    </details>` : '';

  board.innerHTML=groupsHtml+excludedHtml;

  if(!window.Sortable) return;
  Array.from(board.querySelectorAll('.expense-dropzone')).forEach(zone=>{
    expenseSortables.push(new Sortable(zone,{
      group:'expense-categories',
      animation:160,
      handle:'.drag-handle',
      delay:180,
      delayOnTouchOnly:true,
      touchStartThreshold:4,
      fallbackTolerance:4,
      forceFallback:true,
      fallbackOnBody:true,
      ghostClass:'drag-ghost',
      chosenClass:'drag-chosen',
      dragClass:'drag-active',
      onEnd(evt){
        const id=evt.item?.dataset?.id, category=evt.to?.dataset?.category;
        const row=state.expense.find(x=>x.id===id);
        if(row && category){
          const changed=row.category!==category;
          row.category=category;
          renderStats();
          renderExpense();
          if(changed) toast(`“${category}”으로 분류를 변경했습니다.`);
        }
      }
    }));
  });
}

function renderExpense(){
  renderExpenseBoard();
}
function renderCard(){
  const tbody=$('#cardTable tbody'); tbody.innerHTML='';
  filtered(state.card).forEach(r=>{ const tr=document.createElement('tr'); tr.innerHTML=`<td><span class="category-badge">${esc(r.category)}</span></td><td class="center">${dateText(r.date)}</td><td>${esc(r.detail)}</td><td class="amount">${won(r.amount)}</td><td>${esc(r.statItem)}</td>`; tbody.appendChild(tr); });
  if(!tbody.children.length) tbody.innerHTML='<tr><td colspan="5" class="center">해당 월의 법인카드 공개대상 내역이 없습니다.</td></tr>';
}
function renderGift(){
  const tbody=$('#giftTable tbody'); tbody.innerHTML='';
  filtered(state.gift).forEach(r=>{ const d=giftDerived(r); const tr=document.createElement('tr'); tr.innerHTML=`<td class="center"><span class="status-badge ${r.qtyConfirmed?'ok':'review'}">${r.qtyConfirmed?'자동확인':'확인 필요'}</span></td><td class="center">${dateText(r.date)}</td><td>${esc(r.detail)}</td><td>${esc(r.vendor)}</td><td class="editable"><select class="cell-select gift-edit" data-id="${r.id}" data-field="unitPrice">${[5000,10000,30000,50000,100000].map(v=>`<option value="${v}" ${v===r.unitPrice?'selected':''}>${won(v)}</option>`).join('')}</select></td><td class="editable"><input class="cell-input gift-edit" type="number" min="1" step="1" value="${r.quantity}" data-id="${r.id}" data-field="quantity"></td><td class="amount">${won(d.face)}</td><td class="amount">${won(r.amount)}</td><td class="amount">${(d.rate*100).toFixed(2)}%</td>`; tbody.appendChild(tr); });
  if(!tbody.children.length) tbody.innerHTML='<tr><td colspan="9" class="center">해당 월의 상품권 내역이 없습니다.</td></tr>';
}
function renderAll(){ renderMonthSelect(); renderStats(); renderExpense(); renderCard(); renderGift(); }

function onExpenseAction(e){
  const btn=e.target.closest('[data-expense-action]'); if(!btn) return;
  const row=state.expense.find(x=>x.id===btn.dataset.id); if(!row) return;
  const action=btn.dataset.expenseAction;
  if(action==='exclude'){
    row.excluded=true; renderStats(); renderExpense();
    toast('보고서에서 제외했습니다.','되돌리기',()=>{row.excluded=false;renderStats();renderExpense();});
  }else if(action==='restore'){
    row.excluded=false; renderStats(); renderExpense(); toast('제외한 항목을 복원했습니다.');
  }
}

function onEdit(e){
  const el=e.target, id=el.dataset.id, field=el.dataset.field; if(!id||!field) return;
  let row;
  if(id.startsWith('e')) row=state.expense.find(x=>x.id===id); else if(id.startsWith('g')) row=state.gift.find(x=>x.id===id);
  if(!row) return;
  if(field==='quantity'||field==='unitPrice'){ row[field]=Number(el.value)||0; if(field==='quantity') row.qtyConfirmed=true; renderGift(); renderStats(); }
  else { row[field]=el.value; if(id.startsWith('e') && (field==='category'||field==='privacyMode')) { if(field==='category') renderStats(); renderExpense(); } }
}

function grouped(list, groups){
  const map=new Map(groups.map(g=>[g,[]])); list.forEach(x=>{ if(map.has(x.category)) map.get(x.category).push(x); }); map.forEach(arr=>arr.sort(sortDate)); return map;
}
function preview(type){
  const c=current(); let title='', html='';
  if(type==='expense'){
    title=`업무추진비 월별 집행내역 (${monthLabel(state.selectedMonth)})`;
    const gp=grouped(c.expense,EXPENSE_GROUPS); const total=c.expense.reduce((s,x)=>s+x.amount,0);
    html=`<h2>${esc(title)}</h2><h4>▣ 내역별 현황</h4><table><thead><tr><th>구분</th><th>건수</th><th>금액</th><th>구성비</th></tr></thead><tbody>${EXPENSE_GROUPS.map(g=>{const a=gp.get(g);const amt=a.reduce((s,x)=>s+x.amount,0);return `<tr><td>${esc(g)}</td><td>${a.length}건</td><td>${num(amt)}</td><td>${total?((amt/total)*100).toFixed(1):'0.0'}%</td></tr>`}).join('')}</tbody></table><h4>▣ 세부 집행내역</h4><table><thead><tr><th>구분</th><th>집행일자</th><th>세부내역</th><th>집행대상자</th><th>장소</th><th>집행금액</th><th>결재방법</th></tr></thead><tbody>${c.expense.sort((a,b)=>EXPENSE_GROUPS.indexOf(a.category)-EXPENSE_GROUPS.indexOf(b.category)||sortDate(a,b)).map(x=>`<tr><td>${esc(x.category)}</td><td>${dateText(x.date)}</td><td>${esc(publicText(x.detail,x))}</td><td>${esc(publicText(x.target,x))}</td><td>${esc(publicVendor(x))}</td><td>${num(x.amount)}</td><td>${esc(x.payment)}</td></tr>`).join('')}</tbody></table>`;
  }else if(type==='card'){
    title=`법인카드 사용내역 (${monthLabel(state.selectedMonth)})`;
    html=`<h2>${esc(title)}</h2><h4>▣ 공개대상 : 업무추진비(50만원 이상) 기타(100만원 이상)</h4><table><thead><tr><th>구분</th><th>사용일자</th><th>사용내역</th><th>금액</th></tr></thead><tbody>${c.card.sort((a,b)=>CARD_GROUPS.indexOf(a.category)-CARD_GROUPS.indexOf(b.category)||sortDate(a,b)).map(x=>`<tr><td>${esc(x.category)}</td><td>${dateText(x.date)}</td><td>${esc(x.detail)}</td><td>${num(x.amount)}</td></tr>`).join('')}</tbody></table>`;
  }else{
    title=`상품권 구매 및 사용내역 (${monthLabel(state.selectedMonth)})`;
    html=`<h2>${esc(title)}</h2><table><thead><tr><th>구매일자</th><th>구매(사용)용도</th><th>구매처</th><th>총구매수량</th><th>총구매금액(할인전)</th><th>결제금액(할인후)</th><th>할인율</th></tr></thead><tbody>${c.gift.map(x=>{const d=giftDerived(x);return `<tr><td>${dateText(x.date)}</td><td>${esc(x.detail)}</td><td>${esc(x.vendor)}</td><td>${x.quantity}</td><td>${num(d.face)}</td><td>${num(x.amount)}</td><td>${(d.rate*100).toFixed(2)}%</td></tr>`}).join('')}</tbody></table>`;
  }
  $('#previewTitle').textContent=title; $('#previewContent').innerHTML=html; $('#previewDialog').showModal();
}

