/* =========================================================================
   전화요금 계정분할 자동화 — app.js (lazy-billsplit)
   ⚠ 계산 로직(normKey/parseAmount/매핑 매칭/합산/정렬/엑셀 출력 서식) 섹션은
      "CALC CORE"로 표시되어 있습니다. 이 부분은 절대 수정하지 마세요.
      화면 공통 기능(Toast/에러배너/확인모달/Stepper/Progress/Storage/DnD/표복사)은
      lazy-office 저장소의 shared/app-core.js를 CDN으로 불러와 사용합니다
      (window.AppCore). index.html의 <script> 로드 순서를 바꾸지 마세요.
   ========================================================================= */

const STORAGE_KEY = 'phoneBillMapping_v1';
const PERSIST_FLAG_KEY = 'phoneBillMapping_persist_v1';

let mapping = [];       // [{dept, deptName, name, phone, account, owner}]
let rawWorkbook = null;
let lastResult = null;  // {rows, unmatched, otherDeptSkipped, memo}
let persistEnabled = true;
let sortState = { col: null, dir: 1 };
let officialDeptTotals = new Map(); // dept -> {amount, sheet, deptName} — from per-department subtotal sheets (검증용, 계산 로직과 무관한 신규 부가기능)

const { el, toast, error: errorBanner, dialog: dialogUi, stepper, progress: progressUi, storage, dropzone, table: tableUtil } = AppCore;
const showToast = toast.show;
const showError = errorBanner.show;
const setStep = stepper.set;
const setProgress = progressUi.set;

/* ================= CALC CORE : helpers (동일, 수정 금지) ================= */
function normKey(v){
  if(v===null||v===undefined) return '';
  if(typeof v==='number'){
    return Number.isFinite(v) ? String(Number.isInteger(v)?v:v) : '';
  }
  let s=String(v).trim();
  if(/^\d+\.0+$/.test(s)) s = s.replace(/\.0+$/,'');
  return s;
}
function parseAmount(v){
  if(typeof v==='number') return v;
  if(v===null||v===undefined) return 0;
  const s=String(v).replace(/,/g,'').trim();
  const n=parseFloat(s);
  return isNaN(n)?0:n;
}
function fmt(n){ return Number(n||0).toLocaleString('ko-KR'); }
function headerNorm(h){ return String(h||'').replace(/\s/g,''); }
/* ================= /CALC CORE ================= */

/* ---------- 계정 매핑 저장/로딩 ---------- */
function loadPersistFlag(){
  persistEnabled = storage.loadPersistFlag(PERSIST_FLAG_KEY, 'persist-toggle');
}
document.getElementById('persist-toggle').addEventListener('change', (e)=>{
  persistEnabled = e.target.checked;
  storage.savePersistFlag(PERSIST_FLAG_KEY, persistEnabled);
  if(persistEnabled){
    storage.save(STORAGE_KEY, mapping, true);
    showToast('이 PC에 저장하도록 설정했습니다.', 'success');
  }else{
    showToast('이 PC에 저장하지 않습니다. (현재 세션에서만 사용됩니다)');
  }
});

function loadMapping(){
  mapping = storage.load(STORAGE_KEY, []);
  renderMapping();
}
function saveMapping(){
  storage.save(STORAGE_KEY, mapping, persistEnabled);
  const status = document.getElementById('mapping-status');
  status.innerHTML = '';
  const msg = persistEnabled
    ? `현재 ${mapping.length}명이 등록되어 있습니다. (브라우저에 자동 저장됨)`
    : `현재 ${mapping.length}명이 등록되어 있습니다. (저장 옵션이 꺼져있어 이 세션에서만 유지됩니다)`;
  status.appendChild(el('div','status ok', msg));
}

function renderMapping(){
  const tbody = document.getElementById('mapping-tbody');
  const filterText = (document.getElementById('mapping-search').value || '').trim().toLowerCase();
  tbody.innerHTML = '';
  mapping.forEach((row, idx)=>{
    if(filterText){
      const hay = [row.name, row.phone, row.dept].join(' ').toLowerCase();
      if(!hay.includes(filterText)) return;
    }
    const tr = document.createElement('tr');
    const fields = ['dept','deptName','name','phone','account','owner'];
    const fieldLabels = {dept:'부서번호',deptName:'부서명',name:'성명',phone:'전화번호',account:'계정번호',owner:'계정책임자'};
    fields.forEach(f=>{
      const td = document.createElement('td');
      const input = document.createElement('input');
      input.value = row[f] ?? '';
      input.setAttribute('aria-label', fieldLabels[f]);
      input.addEventListener('change', ()=>{ mapping[idx][f] = input.value.trim(); saveMapping(); });
      td.appendChild(input);
      tr.appendChild(td);
    });
    const tdDel = document.createElement('td');
    const delBtn = el('button','btn btn-outline','✕');
    delBtn.type = 'button';
    delBtn.setAttribute('aria-label', `${row.name || row.phone || idx+1}번째 행 삭제`);
    delBtn.style.padding = '4px 8px';
    delBtn.addEventListener('click', ()=>{ mapping.splice(idx,1); renderMapping(); saveMapping(); });
    tdDel.appendChild(delBtn);
    tr.appendChild(tdDel);
    tbody.appendChild(tr);
  });
  saveMapping();
}
document.getElementById('mapping-search').addEventListener('input', renderMapping);

document.getElementById('add-row-btn').addEventListener('click', ()=>{
  mapping.push({dept:'',deptName:'',name:'',phone:'',account:'',owner:''});
  renderMapping();
  showToast('행이 추가되었습니다.', 'success');
});
document.getElementById('clear-mapping-btn').addEventListener('click', async ()=>{
  const ok = await dialogUi.confirm({
    title:'전체 삭제',
    message:'등록된 계정 매핑을 전부 삭제할까요? 되돌릴 수 없습니다.',
    confirmText:'삭제', cancelText:'취소', danger:true
  });
  if(ok){
    mapping = [];
    renderMapping();
    showToast('전체 삭제되었습니다.', 'success');
  }
});
document.getElementById('export-mapping-btn').addEventListener('click', ()=>{
  const header = ['부서번호','부서명','성명','전화번호','계정번호','계정책임자'];
  const aoa = [header, ...mapping.map(r=>[r.dept,r.deptName,r.name,r.phone,r.account,r.owner])];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [{wch:10},{wch:22},{wch:10},{wch:10},{wch:14},{wch:12}];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '계정매핑');
  XLSX.writeFile(wb, `계정매핑_백업_${new Date().toISOString().slice(0,10)}.xlsx`);
  showToast('백업 파일을 다운로드했습니다.', 'success');
});

document.getElementById('guideBadge').addEventListener('keydown', (e)=>{
  if(e.key==='Enter' || e.key===' '){ e.preventDefault(); e.target.click(); }
});
document.getElementById('guideBadge').addEventListener('click', ()=>{
  dialogUi.info({
    title:'사용법',
    bodyHtml:`<ul>
      <li><b>계정 매핑</b>: [계정 매핑 양식 다운로드]로 빈 엑셀을 받아 채운 뒤 업로드하면 자동으로 표에 들어옵니다. 등록한 내용은 이 브라우저에 저장되어 다음 달에도 그대로 남아있습니다 (다른 컴퓨터/브라우저에서는 다시 불러와야 해요).</li>
      <li><b>부서별 전화 사용내역 업로드</b>: 매달 연구전략실(시설관리실)에서 배포하는 원본 파일(부서별 전화 사용 내역)을 업로드하고 [분할 실행]을 누르면 계정번호별로 자동 합산됩니다.</li>
      <li><b>계정 매핑 양식</b>: 연구전략실(시설관리실) 배포 원본파일의 번호별요금 시트에서 부서번호/부서명/성명/전화번호를 복붙하신 뒤 계정번호, 계정책임자를 입력하시면 편해요.</li>
      <li>매핑 키는 부서번호 + 전화번호 조합입니다 (이름은 보지 않습니다). 매핑이 안 되는 번호가 있으면 별도 카드로 보여드려요.</li>
    </ul>`
  });
});

document.getElementById('download-template-btn').addEventListener('click', ()=>{
  const header = ['부서번호','부서명','성명','전화번호','계정번호','계정책임자'];
  const example = ['5400','국방안전융합연구본부','홍길동','1234','25HR1234','홍길동'];
  const aoa = [header, example];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [{wch:10},{wch:22},{wch:10},{wch:10},{wch:14},{wch:12}];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '계정매핑');
  XLSX.writeFile(wb, `계정매핑_표준양식.xlsx`);
  showToast('표준 양식을 다운로드했습니다.', 'success');
});

/* ---------- 계정 매핑 파일 업로드 (표준 양식 인식) ---------- */
function handleMappingFile(file){
  if(!file) return;
  const reader = new FileReader();
  reader.onload = async (ev)=>{
    let wb;
    try{
      wb = XLSX.read(new Uint8Array(ev.target.result), {type:'array'});
    }catch(err){
      showError('엑셀 파일을 읽지 못했습니다. 파일이 손상되지 않았는지 확인해주세요.');
      return;
    }
    let sheetName = null, headers = null;
    for(const sn of wb.SheetNames){
      const aoaCheck = XLSX.utils.sheet_to_json(wb.Sheets[sn], {header:1, defval:''});
      if(aoaCheck.length===0) continue;
      const h = aoaCheck[0].map(headerNorm);
      if(h.includes('부서번호') && h.includes('전화번호') && h.includes('계정번호')){
        sheetName = sn; headers = h; break;
      }
    }
    if(!sheetName){
      await dialogUi.info({
        title:'파일 형식을 확인해주세요',
        bodyHtml:'필요한 열(부서번호/전화번호/계정번호)이 있는 시트를 찾지 못했습니다. "계정 매핑 양식 다운로드"로 받은 양식 그대로 채워서 올려주세요.',
        icon:'warning'
      });
      return;
    }
    const sheet = wb.Sheets[sheetName];
    const aoa = XLSX.utils.sheet_to_json(sheet, {header:1, defval:''});
    const colIdx = (names)=> headers.findIndex(h=>names.includes(h));
    const iDept = colIdx(['부서번호']);
    const iDeptName = colIdx(['부서명']);
    const iName = colIdx(['성명']);
    const iPhone = colIdx(['전화번호']);
    const iAccount = colIdx(['계정번호']);
    const iOwner = colIdx(['계정책임자']);
    const imported = [];
    for(let r=1; r<aoa.length; r++){
      const row = aoa[r];
      const dept = normKey(row[iDept]);
      const phone = normKey(row[iPhone]);
      if(!dept && !phone) continue;
      imported.push({
        dept, deptName: iDeptName>=0?String(row[iDeptName]||'').trim():'',
        name: iName>=0?String(row[iName]||'').trim():'',
        phone, account: iAccount>=0?String(row[iAccount]||'').trim():'',
        owner: iOwner>=0?String(row[iOwner]||'').trim():''
      });
    }
    if(imported.length===0){ showError('불러올 데이터가 없습니다.'); return; }
    if(mapping.length>0){
      const choice = await dialogUi.confirm({
        title:'계정 매핑 불러오기',
        message:`${imported.length}건을 불러옵니다. 기존에 등록된 ${mapping.length}건을 어떻게 할까요?`,
        confirmText:'이어서 추가', cancelText:'새로 덮어쓰기', neutralText:'취소'
      });
      if(choice === null) return;
      mapping = choice ? mapping.concat(imported) : imported;
    } else {
      mapping = imported;
    }
    renderMapping();
    setStep(1);
    showToast(`${imported.length}건을 불러왔습니다.`, 'success');
  };
  reader.readAsArrayBuffer(file);
}
document.getElementById('import-mapping-file').addEventListener('change', (e)=>{
  handleMappingFile(e.target.files[0]);
  e.target.value = '';
});

/* ---------- [신규] 부서별 공식 합계 스캔 (검증용, 계산 로직과 무관한 부가 기능) ----------
   원본 파일에는 "1.○○연구소", "2.△△연구소"처럼 상위 그룹별로 나뉜 시트가 있고,
   그 안에 부서번호별 전화요금 소계가 들어있습니다. 이 값을 부서번호별로 모아
   우리가 계산한 분할 합계와 나중에 비교합니다. */
function buildOfficialDeptTotals(workbook){
  const map = new Map(); // dept -> {amount, sheet, deptName}
  const skipSheets = new Set(['번호별요금','Module1']);
  workbook.SheetNames.forEach(sn=>{
    if(skipSheets.has(sn)) return;
    let aoa;
    try{ aoa = XLSX.utils.sheet_to_json(workbook.Sheets[sn], {header:1, defval:''}); }
    catch(e){ return; }
    // 첫 10행 안에서 "부서번호" + "전화요금"이 함께 있는 헤더 행을 찾는다
    let headerRow = -1, iDept = -1, iName = -1, iFee = -1;
    for(let r=0; r<Math.min(10, aoa.length); r++){
      const norm = (aoa[r]||[]).map(headerNorm);
      const dIdx = norm.findIndex(h=>h==='부서번호');
      const nIdx = norm.findIndex(h=>h==='사용부서' || h==='부서명');
      const fIdx = norm.findIndex(h=>h.indexOf('전화요금')===0);
      if(dIdx>=0 && fIdx>=0){ headerRow=r; iDept=dIdx; iName=nIdx; iFee=fIdx; break; }
    }
    if(headerRow<0) return;
    for(let r=headerRow+1; r<aoa.length; r++){
      const row = aoa[r];
      const dept = normKey(row[iDept]);
      if(!dept) continue;
      const amount = parseAmount(row[iFee]);
      if(!amount) continue; // 소계/합계 행(예: "계") 등은 부서번호가 없거나 숫자가 아니라 자연히 걸러짐
      if(!map.has(dept)){
        map.set(dept, {amount, sheet: sn, deptName: iName>=0?String(row[iName]||'').trim():''});
      }
    }
  });
  return map;
}

/* ---------- 전화요금 원본 파일 업로드 ---------- */
function handleRawFile(file){
  if(!file) return;
  setProgress('ready');
  const reader = new FileReader();
  reader.onload = (ev)=>{
    setProgress('analyzing');
    setTimeout(()=>{
      try{
        rawWorkbook = XLSX.read(new Uint8Array(ev.target.result), {type:'array'});
      }catch(err){
        showError('엑셀 파일을 읽지 못했습니다. 파일이 손상되지 않았는지 확인해주세요.');
        setProgress('idle');
        return;
      }
      officialDeptTotals = buildOfficialDeptTotals(rawWorkbook);
      let target = rawWorkbook.SheetNames.find(n=>headerNorm(n)==='번호별요금');
      const status = document.getElementById('raw-status');
      status.innerHTML = '';
      if(!target){
        rawWorkbook = null;
        dialogUi.info({
          title:'시트를 찾지 못했어요',
          bodyHtml:'"번호별요금" 시트를 자동으로 찾지 못했습니다. 원본 파일 형식이 맞는지 확인 후 다시 올려주세요.',
          icon:'warning'
        });
        setProgress('idle');
        return;
      }
      status.appendChild(el('div','status ok', `[번호별요금] 시트를 찾았습니다. (${file.name})`));
      // 적요(비고) 월 자동 추출
      // 파일명 형식이 보통 "2026년 6월(5월 사용분) 부서별 전화사용 내역"처럼
      // 청구월(6월)과 사용월(5월 사용분, 괄호 안)이 함께 있으므로,
      // 괄호 바로 앞에 오는 월(=청구월)을 우선으로 사용합니다.
      let monthGuess = '';
      const fname = file.name || '';
      let mName = fname.match(/(\d{4})년\s*(\d{1,2})월\s*[\(（]/);
      if(!mName) mName = fname.match(/(\d{4})년\s*(\d{1,2})월/);
      if(mName){
        monthGuess = `${mName[2]}월 전화요금`;
      }else{
        // 파일명에서 못 찾으면 시트 내용에서 찾은 월(대개 사용월)로 대체
        outer:
        for(const sn of rawWorkbook.SheetNames){
          const rows = XLSX.utils.sheet_to_json(rawWorkbook.Sheets[sn], {header:1, defval:''});
          for(const row of rows){
            for(const cell of row){
              const m = String(cell).match(/(\d{4})년\s*(\d{1,2})월/);
              if(m){ monthGuess = `${m[2]}월 전화요금`; break outer; }
            }
          }
        }
      }
      document.getElementById('memo-input').value = monthGuess || `${new Date().getMonth()+1}월 전화요금`;
      document.getElementById('run-split-btn').disabled = false;
      setProgress('done');
      setStep(2);
      showToast('분석 완료', 'success');
    }, 250); // 분석중 단계를 눈으로 확인할 수 있도록 짧은 지연
  };
  reader.readAsArrayBuffer(file);
}
document.getElementById('import-raw').addEventListener('change', (e)=>{
  handleRawFile(e.target.files[0]);
  e.target.value = '';
});

/* ================= CALC CORE : 분할 실행 (동일, 수정 금지) ================= */
document.getElementById('run-split-btn').addEventListener('click', ()=>{
  if(!rawWorkbook) return;
  if(mapping.length===0){
    showError('먼저 [① 계정 매핑]에서 계정 정보를 등록해주세요.');
    return;
  }
  const sheetName = document.getElementById('sheet-select').value || rawWorkbook.SheetNames.find(n=>headerNorm(n)==='번호별요금');
  const sheet = rawWorkbook.Sheets[sheetName];
  const aoa = XLSX.utils.sheet_to_json(sheet, {header:1, defval:''});
  if(aoa.length<2){ showError('선택한 시트에 데이터가 없습니다.'); return; }
  const headers = aoa[0].map(headerNorm);
  const colIdx = (names)=> headers.findIndex(h=>names.includes(h));
  const iDept = colIdx(['부서번호']);
  const iDeptName = colIdx(['부서명']);
  const iName = colIdx(['성명']);
  const iPhone = colIdx(['전화번호']);
  const iFee = colIdx(['전화요금']);
  if(iDept<0 || iPhone<0 || iFee<0){
    showError('원본 시트에서 부서번호/전화번호/전화요금 열을 찾지 못했습니다. 원본 파일 형식이 맞는지 확인해주세요.');
    return;
  }

  // build mapping dict: dept-phone -> {account, deptName, owner} (first occurrence wins, same as original)
  const mapDict = new Map();
  mapping.forEach(m=>{
    if(!m.dept && !m.phone) return;
    const key = normKey(m.dept)+'-'+normKey(m.phone);
    if(!mapDict.has(key)){
      mapDict.set(key, {account: (m.account||'').trim(), deptName: (m.deptName||'').trim(), owner:(m.owner||'').trim()});
    }
  });

  const knownDepts = new Set([...mapDict.keys()].map(k=>k.split('-')[0]).filter(Boolean));

  const sumMap = new Map();   // deptNo-account -> amount
  const infoMap = new Map();  // deptNo-account -> {deptName, owner}
  const unmatched = [];       // only rows whose 부서번호 is already known (likely a real registration gap)
  let otherDeptSkipped = 0;   // rows belonging to departments we don't track at all — expected, not shown

  for(let r=1; r<aoa.length; r++){
    const row = aoa[r];
    const dept = normKey(row[iDept]);
    const phone = normKey(row[iPhone]);
    if(!dept && !phone) continue;
    const key = dept+'-'+phone;
    const amount = parseAmount(row[iFee]);
    if(mapDict.has(key)){
      const info = mapDict.get(key);
      const sumKey = dept+'-'+info.account;
      sumMap.set(sumKey, (sumMap.get(sumKey)||0) + amount);
      if(!infoMap.has(sumKey)) infoMap.set(sumKey, {deptName: info.deptName, owner: info.owner});
    }else if(knownDepts.has(dept)){
      unmatched.push({
        dept, deptName: iDeptName>=0?String(row[iDeptName]||'').trim():'',
        name: iName>=0?String(row[iName]||'').trim():'',
        phone, amount
      });
    }else{
      otherDeptSkipped++;
    }
  }

  // build output rows sorted by numeric dept, then account
  const outRows = [];
  for(const [key, amount] of sumMap.entries()){
    const [dept, account] = key.split('-');
    const info = infoMap.get(key) || {deptName:'', owner:''};
    outRows.push({dept, account, deptName: info.deptName, owner: info.owner, amount});
  }
  outRows.sort((a,b)=>{
    const da = parseFloat(a.dept), db = parseFloat(b.dept);
    if(!isNaN(da) && !isNaN(db) && da!==db) return da-db;
    if(a.dept !== b.dept) return a.dept.localeCompare(b.dept);
    return a.account.localeCompare(b.account);
  });

  const memo = document.getElementById('memo-input').value.trim() || `${new Date().getMonth()+1}월 전화요금`;
  const costName = document.getElementById('cost-name-input').value.trim() || '공공요금';
  lastResult = {rows: outRows, unmatched, otherDeptSkipped, memo, costName};
  sortState = { col: null, dir: 1 };
  renderResult();
  setStep(3);
  showToast('분할 완료', 'success');
});
/* ================= /CALC CORE ================= */

function sortedRowsForDisplay(rows){
  if(!sortState.col || sortState.col === 'dept' || sortState.col === 'deptName'){
    return rows; // 기본 정렬(부서번호 순) — 부서별 묶음 보기 유지
  }
  const copy = rows.slice();
  copy.sort((a,b)=>{
    let av = a[sortState.col], bv = b[sortState.col];
    if(sortState.col === 'amount'){ av = Number(av); bv = Number(bv); }
    else { av = String(av||''); bv = String(bv||''); }
    if(av < bv) return -1*sortState.dir;
    if(av > bv) return 1*sortState.dir;
    return 0;
  });
  return copy;
}

function renderResult(){
  const {rows, unmatched, otherDeptSkipped, memo, costName} = lastResult;
  document.getElementById('result-panel').style.display = 'block';

  // stats
  const total = rows.reduce((s,r)=>s+r.amount,0);
  const unmatchedTotal = unmatched.reduce((s,u)=>s+u.amount,0);
  const stat = document.getElementById('stat-cards');
  stat.innerHTML = '';
  const cards = [
    ['총 전화요금(분할분)', fmt(total)+' 원', false],
    ['분할된 계정 수', rows.length+' 건', false],
    ['등록 누락 의심', unmatched.length+' 건', unmatched.length>0],
    ['등록 누락 추정 금액', fmt(unmatchedTotal)+' 원', unmatchedTotal>0],
  ];
  cards.forEach(([l,n,warn])=>{
    const c = el('div','card-mini'+(warn?' warn':''));
    c.appendChild(el('div','n',n));
    c.appendChild(el('div','l',l));
    stat.appendChild(c);
  });

  const tbody = document.getElementById('result-tbody');
  tbody.innerHTML = '';

  if(!sortState.col || sortState.col==='dept' || sortState.col==='deptName'){
    // 기본: 부서별 rowSpan 묶음 보기 (원본과 동일)
    let i = 0;
    while(i < rows.length){
      let j = i;
      while(j < rows.length && rows[j].dept === rows[i].dept) j++;
      const groupLen = j - i;
      for(let k=i; k<j; k++){
        const r = rows[k];
        const tr = document.createElement('tr');
        if(k===i){
          const tdName = el('td', undefined, r.deptName);
          tdName.rowSpan = groupLen;
          const tdDept = el('td', undefined, r.dept);
          tdDept.rowSpan = groupLen;
          tr.appendChild(tdName);
          tr.appendChild(tdDept);
        }
        tr.appendChild(el('td', undefined, r.account));
        tr.appendChild(el('td', undefined, r.owner));
        tr.appendChild(el('td','num', fmt(r.amount)));
        tr.appendChild(el('td', undefined, costName));
        tr.appendChild(el('td', undefined, memo));
        tbody.appendChild(tr);
      }
      i = j;
    }
  }else{
    // 다른 컬럼 정렬: 묶음 없이 전체 컬럼 반복 표시
    const display = sortedRowsForDisplay(rows);
    display.forEach(r=>{
      const tr = document.createElement('tr');
      tr.appendChild(el('td', undefined, r.deptName));
      tr.appendChild(el('td', undefined, r.dept));
      tr.appendChild(el('td', undefined, r.account));
      tr.appendChild(el('td', undefined, r.owner));
      tr.appendChild(el('td','num', fmt(r.amount)));
      tr.appendChild(el('td', undefined, costName));
      tr.appendChild(el('td', undefined, memo));
      tbody.appendChild(tr);
    });
  }

  // unmatched
  const uPanel = document.getElementById('unmatched-panel');
  const uBody = document.getElementById('unmatched-tbody');
  const uDesc = document.getElementById('unmatched-desc');
  uBody.innerHTML = '';
  uDesc.textContent = `이미 계정 매핑에 등록된 부서인데, 이 번호만 매핑에 없어서 이번 분할에서 제외되었습니다. 신규 입사자나 번호 변경을 놓쳤을 가능성이 있으니 확인 후 [① 계정 매핑]에 추가해주세요. (다른 부서 소속 번호 ${otherDeptSkipped.toLocaleString('ko-KR')}건은 우리 부서 대상이 아니라 정상적으로 제외되었습니다.)`;
  if(unmatched.length>0){
    uPanel.style.display = 'block';
    unmatched.forEach(u=>{
      const tr = document.createElement('tr');
      tr.appendChild(el('td', undefined, u.dept));
      tr.appendChild(el('td', undefined, u.deptName));
      tr.appendChild(el('td', undefined, u.name));
      tr.appendChild(el('td', undefined, u.phone));
      tr.appendChild(el('td','num', fmt(u.amount)));
      uBody.appendChild(tr);
    });
  }else{
    uPanel.style.display = 'none';
  }

  renderDeptValidation(rows, unmatched);
}

/* ---------- [신규] 부서별 합계 검증 렌더링 (계산 로직과 무관한 부가 기능) ---------- */
function renderDeptValidation(rows, unmatched){
  const panel = document.getElementById('validation-panel');
  const tbody = document.getElementById('validation-tbody');
  tbody.innerHTML = '';

  // 우리 쪽 부서별 합계 = "실제로 계정에 분할된 금액"만 집계합니다 (등록 누락분은 제외).
  // → 이렇게 해야 [① 계정 매핑] 등록이 빠진 경우에도 여기서 불일치로 잡아낼 수 있습니다.
  //   (등록 누락분까지 더해버리면 항상 원본 합계와 맞아떨어져서 누락을 놓치게 됩니다.)
  const matchedTotalByDept = new Map();
  rows.forEach(r=>{
    matchedTotalByDept.set(r.dept, (matchedTotalByDept.get(r.dept)||0) + r.amount);
  });
  const unmatchedTotalByDept = new Map();
  unmatched.forEach(u=>{
    unmatchedTotalByDept.set(u.dept, (unmatchedTotalByDept.get(u.dept)||0) + u.amount);
  });
  const deptSet = new Set([...matchedTotalByDept.keys(), ...unmatchedTotalByDept.keys()]);

  const depts = Array.from(deptSet).sort((a,b)=>{
    const da = parseFloat(a), db = parseFloat(b);
    if(!isNaN(da) && !isNaN(db)) return da-db;
    return a.localeCompare(b);
  });

  if(depts.length===0){ panel.style.display = 'none'; return; }
  panel.style.display = 'block';

  let uncheckable = 0;
  depts.forEach(dept=>{
    const ours = matchedTotalByDept.get(dept) || 0;
    const official = officialDeptTotals.get(dept);
    const tr = document.createElement('tr');
    tr.appendChild(el('td', undefined, dept));

    if(!official){
      uncheckable++;
      tr.appendChild(el('td', undefined, '-'));
      tr.appendChild(el('td','num', fmt(ours)));
      tr.appendChild(el('td','num', '-'));
      tr.appendChild(el('td','num', '-'));
      const pill = el('span','status-pill unknown','확인 불가');
      const td = document.createElement('td');
      td.appendChild(pill);
      tr.appendChild(td);
    }else{
      const diff = official.amount - ours;
      tr.appendChild(el('td', undefined, official.deptName || ''));
      tr.appendChild(el('td','num', fmt(ours)));
      tr.appendChild(el('td','num', fmt(official.amount)));
      tr.appendChild(el('td','num', fmt(diff)));
      const td = document.createElement('td');
      const pill = el('span', 'status-pill '+(diff===0?'ok':'mismatch'), diff===0 ? '일치' : '불일치');
      td.appendChild(pill);
      tr.appendChild(td);
      if(diff!==0) tr.classList.add('row-mismatch');
    }
    tbody.appendChild(tr);
  });

  const desc = document.getElementById('validation-desc');
  const mismatchCount = depts.filter(d=>{
    const o = officialDeptTotals.get(d);
    return o && (o.amount - (matchedTotalByDept.get(d)||0)) !== 0;
  }).length;
  let msg = '업로드한 원본 파일 안의 부서별 소계 시트(예: "1.○○연구소")에 있는 부서 전체 전화요금과, 우리가 실행한 분할 결과의 부서별 합계가 일치하는지 확인합니다. ("우리 계산 합계"는 실제로 계정에 분할된 금액만 반영하며, 등록 누락 항목은 포함하지 않습니다.)';
  if(mismatchCount>0){
    msg += ` ⚠ ${mismatchCount}개 부서에서 금액이 일치하지 않습니다 — 아래 [⑤ 계정 매핑이 안 된 항목]을 확인해주세요.`;
  }else if(uncheckable>0){
    msg += ` (${uncheckable}개 부서는 원본 파일에서 대응하는 소계 시트를 찾지 못해 비교하지 못했습니다.)`;
  }else{
    msg += ' ✔ 모든 부서 합계가 일치합니다.';
  }
  desc.textContent = msg;
}

/* 표 헤더 클릭 정렬 */
document.querySelectorAll('#result-table th[data-sort]').forEach(th=>{
  th.addEventListener('click', ()=>{
    const col = th.dataset.sort;
    if(!lastResult) return;
    if(sortState.col === col){ sortState.dir *= -1; } else { sortState.col = col; sortState.dir = 1; }
    renderResult();
  });
});

/* ================= CALC CORE : 엑셀 다운로드 (출력 서식 동일, 수정 금지) ================= */
document.getElementById('download-btn').addEventListener('click', ()=>{
  if(!lastResult) return;
  const {rows, memo, costName} = lastResult;
  const header = ['부서명','부서번호','계정번호','계정책임자','전화요금','종비용명','적요'];
  const aoa = [header];
  const merges = [];
  let i = 0;
  while(i < rows.length){
    let j = i;
    while(j < rows.length && rows[j].dept === rows[i].dept) j++;
    for(let k=i;k<j;k++){
      const r = rows[k];
      aoa.push([k===i?r.deptName:'', k===i?r.dept:'', r.account, r.owner, r.amount, costName, memo]);
    }
    if(j-i > 1){
      const startRow = i+1, endRow = j; // +1 for header row offset in sheet coords
      merges.push({s:{r:startRow,c:0}, e:{r:endRow,c:0}});
      merges.push({s:{r:startRow,c:1}, e:{r:endRow,c:1}});
    }
    i = j;
  }
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!merges'] = merges;
  ws['!cols'] = [{wch:22},{wch:10},{wch:14},{wch:12},{wch:12},{wch:10},{wch:14}];
  // number format on 전화요금 column (E, idx4)
  for(let r=1;r<aoa.length;r++){
    const addr = XLSX.utils.encode_cell({r,c:4});
    if(ws[addr]){ ws[addr].t='n'; ws[addr].z='#,##0'; }
  }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '출력결과_전화요금');
  XLSX.writeFile(wb, `전화요금분할결과_${memo.replace(/\s/g,'')}.xlsx`);
  setStep(4);
  showToast('다운로드 완료', 'success');
});
/* ================= /CALC CORE ================= */

/* ---------- 표 복사하기 (한글/워드 등에 "표" 형태로 붙여넣기 가능) ---------- */
document.getElementById('copy-btn').addEventListener('click', ()=>{
  if(!lastResult){ showError('먼저 분할을 실행해주세요.'); return; }
  const {rows, memo, costName} = lastResult;

  // 붙여넣을 문서(한글/워드)가 표로 정확히 인식하도록, 화면 표를 그대로 복사하는 대신
  // 아주 단순한 <table><tr><td> 구조로 직접 새로 만듭니다. (thead/tbody, id, class 등 없이 —
  // 이런 부가 태그·계산된 스타일이 많으면 일부 에디터가 "표 안에 표"로 잘못 해석하는 경우가 있습니다.)
  const cellCss = 'border:1px solid #333333;padding:6px 10px;font-family:맑은 고딕,Malgun Gothic,sans-serif;font-size:12px;';
  const thCss = cellCss + 'background-color:#F1F5FA;font-weight:bold;text-align:center;';
  const tdCss = cellCss + 'text-align:center;';
  const tdCssRight = cellCss + 'text-align:right;';
  const esc = (s)=> String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');

  const headerLabels = ['부서명','부서번호','계정번호','계정책임자','전화요금','종비용명','적요'];
  let html = '<table style="border-collapse:collapse;">';
  html += '<tr>' + headerLabels.map(h=>`<td style="${thCss}">${esc(h)}</td>`).join('') + '</tr>';

  if(!sortState.col || sortState.col==='dept' || sortState.col==='deptName'){
    // 기본 보기: 화면과 동일하게 부서별로 셀 병합
    let i = 0;
    while(i < rows.length){
      let j = i;
      while(j < rows.length && rows[j].dept === rows[i].dept) j++;
      const groupLen = j - i;
      for(let k=i; k<j; k++){
        const r = rows[k];
        html += '<tr>';
        if(k===i){
          html += `<td rowspan="${groupLen}" style="${tdCss}">${esc(r.deptName)}</td>`;
          html += `<td rowspan="${groupLen}" style="${tdCss}">${esc(r.dept)}</td>`;
        }
        html += `<td style="${tdCss}">${esc(r.account)}</td>`;
        html += `<td style="${tdCss}">${esc(r.owner)}</td>`;
        html += `<td style="${tdCssRight}">${esc(fmt(r.amount))}</td>`;
        html += `<td style="${tdCss}">${esc(costName)}</td>`;
        html += `<td style="${tdCss}">${esc(memo)}</td>`;
        html += '</tr>';
      }
      i = j;
    }
  }else{
    // 다른 컬럼으로 정렬된 화면: 병합 없이 매 행 전체 표시 (화면과 동일)
    sortedRowsForDisplay(rows).forEach(r=>{
      html += '<tr>' +
        `<td style="${tdCss}">${esc(r.deptName)}</td>` +
        `<td style="${tdCss}">${esc(r.dept)}</td>` +
        `<td style="${tdCss}">${esc(r.account)}</td>` +
        `<td style="${tdCss}">${esc(r.owner)}</td>` +
        `<td style="${tdCssRight}">${esc(fmt(r.amount))}</td>` +
        `<td style="${tdCss}">${esc(costName)}</td>` +
        `<td style="${tdCss}">${esc(memo)}</td>` +
      '</tr>';
    });
  }
  html += '</table>';

  // 텍스트만 지원하는 경우를 위한 대체용 TSV (엑셀에도 표로 잘 들어갑니다)
  const display = sortedRowsForDisplay(rows);
  const tsvLines = [headerLabels.join('\t')];
  display.forEach(r=>{
    tsvLines.push([r.deptName, r.dept, r.account, r.owner, r.amount, costName, memo].join('\t'));
  });
  const text = tsvLines.join('\n');

  const doneAsTable = () => showToast('표를 복사했습니다. 붙여넣을 때 "셀 안에 표로 넣기"를 선택해주세요.', 'success');
  const doneAsText = () => showToast('표 서식 없이 텍스트로 복사되었습니다. (이 브라우저는 표 복사 방식을 지원하지 않아요)');
  const fail = () => showError('복사에 실패했습니다. 결과를 다운로드한 뒤 엑셀에서 복사해주세요.');

  tableUtil.copyRich(html, text, doneAsTable, doneAsText, fail);
});

/* ---------- Drag & Drop ---------- */
dropzone.attach('mapping-dropzone', handleMappingFile);
dropzone.attach('raw-dropzone', handleRawFile);

/* ---------- init ---------- */
if (typeof XLSX === 'undefined') {
  showError('xlsx.full.min.js 라이브러리 파일을 찾을 수 없습니다. shared/vendor 폴더 안에 xlsx.full.min.js 파일이 있는지 확인해주세요.');
}
loadPersistFlag();
loadMapping();
setStep(1);
setProgress('idle');
