const SLOT_INFO = {
  monthly: ['업무자용 파일', '상세 출결 CSV 또는 엑셀 · 여러 파일 업로드 가능', '▦'],
  parent: ['학부모 나이스 신고서', '결석·지각·조퇴·결과 신고서', '家'],
  experience: ['체험학습', '체험학습 신청·결과 자료', '旅'],
  recognized: ['인정 지각·조퇴·결과', '인정 처리 근거 자료', '認'],
  official: ['공문', '대회·행사 등 관련 공문', '文'],
  manager: ['학급별 출결현황 파일', '대조 기준이 되는 학급별 출결현황 PDF 또는 엑셀', '合']
};
const newClassState = () => ({ ...Object.fromEntries(Object.keys(SLOT_INFO).map(k => [k, []])), manual: [] });
const classes = Object.fromEntries(Array.from({ length: 10 }, (_, i) => [i + 1, newClassState()]));
let activeClass = 5, lastResult = null, currentFilter = 'all';
const $ = id => document.getElementById(id);
const localProcessing = ['127.0.0.1', 'localhost', '::1'].includes(location.hostname);

function renderClasses() {
  $('classButtons').innerHTML = Array.from({ length: 10 }, (_, i) => i + 1).map(n => {
    const count = Object.values(classes[n]).flat().length;
    return `<button type="button" data-class="${n}" class="${n === activeClass ? 'active ' : ''}${count ? 'has-files' : ''}">${n}반 <span>●</span></button>`;
  }).join('');
  document.querySelectorAll('[data-class]').forEach(btn => btn.onclick = () => { activeClass = Number(btn.dataset.class); lastResult = null; render(); });
}

function slotHtml(key, baseline = false) {
  const [name, help, symbol] = SLOT_INFO[key], files = classes[activeClass][key];
  const accept = key === 'manager' ? '.xlsx,.xls,.pdf' : key === 'monthly' ? '.xlsx,.xls,.csv' : '.xlsx,.xls,.pdf,.hwp,.hwpx,.odt';
  return `<div class="file-slot ${baseline ? 'baseline' : ''}" data-slot="${key}" tabindex="0" role="button" aria-label="${name} 파일 올리기">
    <input type="file" accept="${accept}" ${baseline ? '' : 'multiple'} hidden>
    <div class="slot-top"><span class="slot-symbol">${symbol}</span><span><span class="slot-name">${name}</span><span class="slot-help">${help}</span></span><span class="slot-action">+ 파일 추가</span></div>
    <div class="file-pills">${files.map((f, i) => `<span class="file-pill" title="${escapeHtml(f.name)}">${escapeHtml(f.name)}<button type="button" data-remove="${key}:${i}" aria-label="삭제">×</button></span>`).join('')}</div>
  </div>`;
}

function render() {
  renderClasses(); $('classTitle').textContent = `${activeClass}반`; $('classCrumb').textContent = `${activeClass}반`;
  $('baselineSlot').innerHTML = slotHtml('manager', true);
  $('evidenceSlots').innerHTML = ['monthly', 'parent', 'experience', 'recognized', 'official'].map(k => slotHtml(k)).join('');
  renderManual();
  bindSlots();
  const count = Object.values(classes[activeClass]).flat().length;
  $('fileCount').textContent = count; $('auditBtn').disabled = classes[activeClass].manager.length === 0;
  $('results').hidden = !lastResult; if (lastResult) renderResults();
}

function renderManual() {
  const rows = classes[activeClass].manual;
  $('manualList').innerHTML = rows.length ? rows.map((r, i) => `<div class="manual-row"><b>${escapeHtml(r.name)}</b><span>${escapeHtml(r.no ? `${r.no}번` : '번호 미입력')}</span><span>${escapeHtml(r.date)}</span><span>${escapeHtml(AttendanceAudit.LABELS[r.category] || r.category)}</span><span class="manual-note">${escapeHtml(r.file || '')}</span><button type="button" data-manual-remove="${i}" aria-label="수기 기록 삭제">×</button></div>`).join('') : '<p class="manual-empty">등록된 수기 기록이 없습니다.</p>';
  document.querySelectorAll('[data-manual-remove]').forEach(btn => btn.onclick = () => { classes[activeClass].manual.splice(Number(btn.dataset.manualRemove), 1); lastResult = null; render(); });
}

function addManual() {
  const no = $('manualNo').value.trim(), name = $('manualName').value.trim(), date = $('manualDate').value, category = $('manualCategory').value, file = $('manualMemo').value.trim();
  if (!name || !date || !category) { alert('학생 이름, 날짜, 출결 구분을 입력해 주세요.'); return; }
  classes[activeClass].manual.push({ no, name, date, category, file });
  ['manualNo', 'manualName', 'manualDate', 'manualCategory', 'manualMemo'].forEach(id => { $(id).value = ''; });
  lastResult = null; render();
}

function bindSlots() {
  document.querySelectorAll('.file-slot').forEach(slot => {
    const input = slot.querySelector('input');
    slot.onclick = e => { if (!e.target.matches('[data-remove]')) input.click(); };
    slot.onkeydown = e => { if (e.key === 'Enter') input.click(); };
    input.onchange = () => addFiles(slot.dataset.slot, [...input.files]);
  });
  document.querySelectorAll('[data-remove]').forEach(btn => btn.onclick = e => {
    e.stopPropagation(); const [slot, index] = btn.dataset.remove.split(':'); classes[activeClass][slot].splice(Number(index), 1); lastResult = null; render();
  });
}

function addFiles(slot, files) {
  const valid = files.filter(f => baselineFileAllowed(slot, f.name));
  if (slot === 'manager') classes[activeClass][slot] = valid.slice(-1); else classes[activeClass][slot].push(...valid);
  lastResult = null; render();
}

async function auditNow() {
  $('auditBtn').disabled = true; $('auditBtn').innerHTML = '<span>…</span> 점검하는 중';
  const inputs = [];
  for (const [slot, files] of Object.entries(classes[activeClass])) {
    if (slot === 'manual') { files.forEach(row => inputs.push({ ...row, slot, classNo: activeClass, manual: true, name: row.file || '수기 기록' })); continue; }
    for (const file of files) {
      if (slot === 'manager' && /\.pdf$/i.test(file.name)) inputs.push({ slot, classNo: activeClass, name: file.name, pdfText: await extractPdfText(file) });
      else if (/\.odt$/i.test(file.name)) inputs.push({ slot, classNo: activeClass, name: file.name, odtContent: await extractOdtContent(file) });
      else inputs.push({ slot, classNo: activeClass, name: file.name, data: await file.arrayBuffer() });
    }
  }
  lastResult = AttendanceAudit.audit(inputs, XLSX); currentFilter = 'all';
  $('auditBtn').innerHTML = '<span>✓</span> 다시 점검하기'; $('auditBtn').disabled = false;
  $('results').hidden = false; renderResults(); $('results').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

async function extractPdfText(file) {
  const pdfjs = await import('./pdf.min.mjs');
  pdfjs.GlobalWorkerOptions.workerSrc = './pdf.worker.min.mjs';
  const pdf = await pdfjs.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
  const lines = [];
  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber++) {
    const page = await pdf.getPage(pageNumber), content = await page.getTextContent(), byY = new Map();
    for (const item of content.items) {
      if (!item.str || !item.str.trim()) continue;
      const y = Math.round(item.transform[5] * 2) / 2;
      if (!byY.has(y)) byY.set(y, []);
      byY.get(y).push({ x: item.transform[4], text: item.str.trim() });
    }
    lines.push(...[...byY.entries()].sort((a, b) => b[0] - a[0]).map(([, items]) => items.sort((a, b) => a.x - b.x).map(item => item.text).join(' ')));
  }
  return lines.join('\n');
}

async function extractOdtContent(file) {
  if (!localProcessing) throw new Error('개인정보 보호를 위해 배포 사이트에서는 PDF·ODT 파일을 전송하지 않습니다. 이 기능은 컴퓨터에서 실행한 프로그램에서 사용해 주세요.');
  const response = await fetch('/api/odt-content', { method: 'POST', headers: { 'Content-Type': 'application/vnd.oasis.opendocument.text' }, body: file });
  if (!response.ok) throw new Error((await response.json().catch(() => ({}))).error || 'ODT를 읽지 못했습니다.');
  return response.text();
}

function renderResults() {
  const s = lastResult.stats;
  $('summaryCards').innerHTML = `<div class="summary-card"><small>업무자용 상세 기록</small><b>${s.monthly}건</b></div><div class="summary-card"><small>대조 신고 기록</small><b>${s.evidence}건</b></div><div class="summary-card error"><small>발견된 오류</small><b>${s.errors}건</b></div><div class="summary-card ${s.errors ? '' : 'ok'}"><small>확인 필요</small><b>${s.warnings}건</b></div>`;
  document.querySelectorAll('.filter').forEach(btn => btn.classList.toggle('active', btn.dataset.filter === currentFilter));
  const issues = lastResult.issues.filter(i => currentFilter === 'all' || i.severity === currentFilter);
  $('issueList').innerHTML = issues.length ? issues.map(i => `<article class="issue ${i.severity}"><span class="badge">${i.severity === 'error' ? '오류' : '확인 필요'}</span><div><strong>${escapeHtml(i.message)}</strong><p>${escapeHtml(i.detail)}</p></div><code>${escapeHtml(i.code)}</code></article>`).join('') : '<div class="empty"><b>오류가 없습니다</b>올린 자료가 기준 파일과 일치합니다.</div>';
}

function autoSlot(file) {
  const n = file.name;
  if (/학급별|업무.?담당자|담당자.*출결|download/i.test(n)) return 'manager';
  if (/월별.*출결|출결.*현황/i.test(n)) return 'monthly';
  if (/지각|조퇴|결과/i.test(n)) return 'recognized';
  if (/결석신고|학부모|나이스/i.test(n)) return 'parent';
  if (/체험/i.test(n)) return 'experience';
  if (/공문|대회|행사|참가|출석인정|장관|시행/i.test(n)) return 'official';
  return 'parent';
}
function baselineFileAllowed(slot, name) { return slot === 'manager' ? /\.(xlsx?|pdf)$/i.test(name) : slot === 'monthly' ? /\.(xlsx?|csv)$/i.test(name) : /\.(xlsx?|pdf|hwp|hwpx|odt)$/i.test(name); }
function addAuto(files) { for (const f of files.filter(f => /\.(xlsx?|csv|pdf|hwp|hwpx|odt)$/i.test(f.name))) addFiles(autoSlot(f), [f]); }
function escapeHtml(v) { const d = document.createElement('div'); d.textContent = v || ''; return d.innerHTML; }

$('browseAll').onclick = () => $('allFiles').click(); $('allFiles').onchange = () => addAuto([...$('allFiles').files]);
$('auditBtn').onclick = auditNow;
$('resetBtn').onclick = () => { classes[activeClass] = newClassState(); lastResult = null; render(); };
$('addManualBtn').onclick = addManual;
document.querySelectorAll('.filter').forEach(btn => btn.onclick = () => { currentFilter = btn.dataset.filter; renderResults(); });
$('dropZone').ondragover = e => { e.preventDefault(); $('dropZone').classList.add('drag'); };
$('dropZone').ondragleave = () => $('dropZone').classList.remove('drag');
$('dropZone').ondrop = e => { e.preventDefault(); $('dropZone').classList.remove('drag'); addAuto([...e.dataTransfer.files]); };
$('csvBtn').onclick = () => {
  const rows = [['수준','오류유형','내용','상세'], ...lastResult.issues.map(i => [i.severity === 'error' ? '오류' : '확인 필요', i.code, i.message, i.detail])];
  const csv = '\ufeff' + rows.map(r => r.map(v => `"${String(v || '').replace(/"/g, '""')}"`).join(',')).join('\r\n');
  const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' })); a.download = `2학년_${activeClass}반_출결오류.csv`; a.click(); URL.revokeObjectURL(a.href);
};
render();
