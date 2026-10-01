(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.AttendanceAudit = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  const EVENT_TYPES = ['결석', '지각', '조퇴', '결과'];
  const REASONS = ['질병', '미인정', '기타', '인정'];
  const LABELS = Object.fromEntries(EVENT_TYPES.flatMap(e => REASONS.map(r => [`${e}|${r}`, `${r}${e}`])));

  const clean = value => value == null ? '' : String(value).replace(/\s+/g, '').trim();
  const number = value => Number(clean(value).replace(/[^0-9.-]/g, '')) || 0;
  const date = value => {
    const m = String(value == null ? '' : value).match(/(20\d{2})[.\-/년\s]+(\d{1,2})[.\-/월\s]+(\d{1,2})/);
    return m ? `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}` : '';
  };
  const category = value => {
    const v = clean(value).replace('출석인정', '인정').replace('인정결석', '인정결석');
    const event = EVENT_TYPES.find(x => v.includes(x));
    let reason = REASONS.find(x => v.includes(x));
    if (!reason && v.includes('인정')) reason = '인정';
    return event && reason ? `${event}|${reason}` : '';
  };
  const rowKey = row => `${row.no || row.name}|${row.date}|${row.category}`;
  const PARENT_NICE_REQUIRED = new Set(['결석|질병', '결석|인정', '조퇴|인정', '지각|인정']);
  const requiresParentNice = row => PARENT_NICE_REQUIRED.has(row.category);

  function rowsOf(workbook) {
    return workbook.SheetNames.flatMap(name => {
      const sheet = workbook.Sheets[name];
      return [{ name, rows: root.XLSX ? root.XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: false }) : [] }];
    });
  }

  function workbookRows(workbook, XLSX) {
    return workbook.SheetNames.map(name => ({ name, rows: XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: null, raw: false }) }));
  }

  function detect(workbook, XLSX) {
    const sheets = workbookRows(workbook, XLSX);
    const all = sheets.flatMap(s => s.rows).flat().map(clean);
    if (all.includes('일자') && all.includes('출결구분')) return 'monthly';
    if (all.some(v => v.includes('학급별출결현황')) || (all.includes('수업일수') && all.includes('결석총계'))) return 'summary';
    if (all.includes('결석기간') && all.includes('결석신고구분')) return 'absence';
    if (all.includes('신고종류') && all.some(v => v.includes('신고기간'))) return 'movement';
    if (all.includes('신청서체험기간') && all.includes('성명')) return 'experience';
    return 'generic';
  }

  function findHeader(rows, required) {
    return rows.findIndex(row => required.every(word => row.some(cell => clean(cell).includes(word))));
  }

  function parseMonthly(workbook, XLSX) {
    const records = [], warnings = [];
    for (const { name: sheet, rows } of workbookRows(workbook, XLSX)) {
      const h = findHeader(rows, ['일자', '성명', '출결구분']);
      if (h < 0) continue;
      const head = rows[h].map(clean);
      const idx = key => head.findIndex(v => v.includes(key));
      const cols = { date: idx('일자'), no: idx('번호'), name: idx('성명'), type: idx('출결구분'), period: idx('결시교시'), reason: idx('사유') };
      let currentNo = '', currentName = '', currentType = '';
      for (let i = h + 1; i < rows.length; i++) {
        const r = rows[i];
        const d = date(r[cols.date]);
        if (!d) continue;
        if (clean(r[cols.no])) currentNo = clean(r[cols.no]);
        if (clean(r[cols.name])) currentName = clean(r[cols.name]);
        if (clean(r[cols.type])) currentType = clean(r[cols.type]);
        const cat = category(currentType);
        const rec = { date: d, no: currentNo, name: currentName, category: cat, type: currentType, period: clean(r[cols.period]), reason: clean(r[cols.reason]), sheet, row: i + 1 };
        if (!currentNo || !currentName) warnings.push({ severity: 'error', code: '기준파일형식', message: `${d} 행에 학생 번호 또는 이름이 없습니다.`, detail: `${sheet} ${i + 1}행` });
        else if (!cat) warnings.push({ severity: 'error', code: '기준파일형식', message: `${currentName} ${d}의 출결구분을 해석할 수 없습니다.`, detail: currentType || '출결구분 빈칸' });
        else records.push(rec);
      }
    }
    return { records, warnings };
  }

  function parseSummary(workbook, XLSX) {
    const records = [];
    for (const { rows } of workbookRows(workbook, XLSX)) {
      const h = findHeader(rows, ['번호', '성명', '결석', '지각', '조퇴', '결과']);
      if (h < 0 || !rows[h + 1]) continue;
      const top = rows[h].map(clean), sub = rows[h + 1].map(clean);
      const mapping = {};
      let event = '';
      for (let c = 0; c < Math.max(top.length, sub.length); c++) {
        if (EVENT_TYPES.includes(top[c])) event = top[c];
        else if (top[c] && !EVENT_TYPES.includes(top[c])) event = '';
        if (event && REASONS.includes(sub[c])) mapping[c] = `${event}|${sub[c]}`;
      }
      for (let i = h + 2; i < rows.length; i++) {
        const r = rows[i], no = clean(r[0]), name = clean(r[1]);
        if (!/^\d+$/.test(no) || !name) continue;
        const counts = Object.fromEntries(Object.keys(LABELS).map(k => [k, 0]));
        for (const [c, cat] of Object.entries(mapping)) counts[cat] = number(r[Number(c)]);
        records.push({ no, name, counts });
      }
    }
    return records;
  }

  function parseSummaryPdfText(text) {
    const records = [];
    for (const line of String(text || '').split(/\r?\n/)) {
      const tokens = line.trim().split(/\s+/);
      if (!/^\d+$/.test(tokens[0]) || tokens.length < 19 || !tokens[1]) continue;
      const values = tokens.slice(3).filter(v => /^\d+$/.test(v)).map(Number);
      if (values.length < 16) continue;
      const counts = Object.fromEntries(Object.keys(LABELS).map(k => [k, 0]));
      EVENT_TYPES.forEach((event, eventIndex) => REASONS.forEach((reason, reasonIndex) => {
        counts[`${event}|${reason}`] = values[eventIndex * REASONS.length + reasonIndex] || 0;
      }));
      records.push({ no: tokens[0], name: tokens[1], counts });
    }
    return records;
  }

  function expandRange(text) {
    const matches = String(text || '').match(/20\d{2}[.\-/]\d{1,2}[.\-/]\d{1,2}/g) || [];
    const start = date(matches[0]), end = date(matches[1] || matches[0]);
    if (!start) return [];
    const [sy, sm, sd] = start.split('-').map(Number), [ey, em, ed] = end.split('-').map(Number);
    const out = [], d = new Date(Date.UTC(sy, sm - 1, sd)), limit = new Date(Date.UTC(ey, em - 1, ed));
    while (d <= limit && out.length < 370) { out.push(d.toISOString().slice(0, 10)); d.setUTCDate(d.getUTCDate() + 1); }
    return out;
  }
  const schoolDates = dates => dates.filter(value => { const day = new Date(`${value}T00:00:00Z`).getUTCDay(); return day !== 0 && day !== 6; });

  function parseAbsence(workbook, XLSX) {
    const records = [];
    for (const { rows } of workbookRows(workbook, XLSX)) {
      const h = findHeader(rows, ['성명', '결석기간', '결석신고구분']);
      if (h < 0) continue;
      const head = rows[h].map(clean), ix = key => head.findIndex(v => v.includes(key));
      for (let i = h + 1; i < rows.length; i++) {
        const r = rows[i], name = clean(r[ix('성명')]), no = clean(r[ix('번호')]), cat = category(r[ix('결석신고구분')]);
        if (!name || !cat) continue;
        for (const d of expandRange(r[ix('결석기간')])) records.push({ no, name, date: d, category: cat, sourceRow: i + 1 });
      }
    }
    return records;
  }

  function parseMovement(workbook, XLSX) {
    const records = [];
    for (const { rows } of workbookRows(workbook, XLSX)) {
      const h = findHeader(rows, ['신고종류', '성명', '신고구분']);
      if (h < 0) continue;
      const head = rows[h].map(clean), ix = key => head.findIndex(v => v.includes(key));
      const dateCol = ix('접수기간') >= 0 ? ix('접수기간') : ix('신청일자');
      for (let i = h + 1; i < rows.length; i++) {
        const r = rows[i], name = clean(r[ix('성명')]), no = clean(r[ix('번호')]);
        const d = date(r[dateCol]), cat = category(`${r[ix('신고구분')] || ''}${r[ix('신고종류')] || ''}`);
        if (name && d && cat) records.push({ no, name, date: d, category: cat, sourceRow: i + 1 });
      }
    }
    return records;
  }

  function parseExperience(workbook, XLSX) {
    const records = [];
    for (const { rows } of workbookRows(workbook, XLSX)) {
      const h = findHeader(rows, ['성명', '신청서체험기간']);
      if (h < 0) continue;
      const head = rows[h].map(clean), ix = key => head.findIndex(v => v.includes(key));
      for (let i = h + 1; i < rows.length; i++) {
        const r = rows[i], name = clean(r[ix('성명')]), no = clean(r[ix('번호')]);
        if (!name) continue;
        for (const d of schoolDates(expandRange(r[ix('신청서체험기간')]))) records.push({ no, name, date: d, category: '결석|인정', sourceRow: i + 1 });
      }
    }
    return records;
  }

  const decodeXml = value => String(value || '').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n)));
  const xmlText = value => clean(decodeXml(String(value || '').replace(/<[^>]+>/g, ' ')));
  const odtPlainText = xml => decodeXml(String(xml || '').replace(/<text:p\b[^>]*>/g, '\n').replace(/<text:tab\s*\/>/g, ' ').replace(/<[^>]+>/g, ' ')).replace(/[ \t]+/g, ' ').replace(/ *\n */g, '\n');
  function odtRows(xml) {
    const rows = [];
    const rowPattern = /<table:table-row\b[^>]*>([\s\S]*?)<\/table:table-row>/g;
    let row;
    while ((row = rowPattern.exec(String(xml || '')))) {
      const cells = [], cellPattern = /<table:table-cell\b[^>]*>([\s\S]*?)<\/table:table-cell>|<table:covered-table-cell\b[^>]*\/>/g;
      let cell;
      while ((cell = cellPattern.exec(row[1]))) cells.push(xmlText(cell[1] || ''));
      if (cells.length) rows.push(cells);
    }
    return rows;
  }
  function parseOfficialOdt(xml, targetClass) {
    const text = odtPlainText(xml);
    if (!/출석\s*인정\s*결석|출석인정결석/.test(text)) return [];
    const yearMatch = text.match(/(20\d{2})\s*[.년]/), year = yearMatch ? Number(yearMatch[1]) : new Date().getFullYear();
    const periodMatch = text.match(/출석\s*인정\s*결석\s*일수\s*:\s*([^\n]+)/) || text.match(/출석인정결석\s*일수\s*:\s*([^\n]+)/);
    if (!periodMatch) return [];
    const period = periodMatch[1];
    const start = period.match(/(?:(20\d{2})\s*[.년]\s*)?(\d{1,2})\s*[.월]\s*(\d{1,2})\s*일?/);
    const rest = start ? period.slice((start.index || 0) + start[0].length) : '';
    const end = rest.match(/(?:~|∼|-)\s*(?:(20\d{2})\s*[.년]\s*)?(?:(\d{1,2})\s*[.월]\s*)?(\d{1,2})\s*일?/);
    if (!start || !end) return [];
    const sy = Number(start[1] || year), sm = Number(start[2]), sd = Number(start[3]);
    const ey = Number(end[1] || sy), em = Number(end[2] || sm), ed = Number(end[3]);
    const dateRange = schoolDates(expandRange(`${sy}-${sm}-${sd} ~ ${ey}-${em}-${ed}`));
    const studentMatch = text.match(/참가\s*학생\s*:\s*([\s\S]*?)(?=\n[가-힣]\.\s|$)/);
    if (!studentMatch) return [];
    const records = [], studentPattern = /([가-힣]{2,4})\s*\((\d{4,6})\)/g;
    let student;
    while ((student = studentPattern.exec(studentMatch[1]))) {
      const code = student[2], classNo = Number(code.slice(1, -2));
      if (targetClass && classNo !== Number(targetClass)) continue;
      const no = String(Number(code.slice(-2)));
      for (const date of dateRange) records.push({ no, name: student[1], date, category: '결석|인정' });
    }
    return records;
  }
  function parseOdtContent(xml, targetClass) {
    const rows = odtRows(xml);
    const absenceHeader = findHeader(rows, ['성명', '결석기간', '결석신고구분']);
    if (absenceHeader >= 0) {
      const head = rows[absenceHeader].map(clean), ix = key => head.findIndex(v => v.includes(key)), records = [];
      for (const r of rows.slice(absenceHeader + 1)) {
        const name = clean(r[ix('성명')]), no = clean(r[ix('번호')]), cat = category(r[ix('결석신고구분')]);
        if (name && cat) for (const d of expandRange(r[ix('결석기간')])) records.push({ no, name, date: d, category: cat });
      }
      return { type: 'absence', records };
    }
    const movementHeader = findHeader(rows, ['신고종류', '성명', '신고구분']);
    if (movementHeader >= 0) {
      const head = rows[movementHeader].map(clean), ix = key => head.findIndex(v => v.includes(key)), dateCol = ix('접수기간') >= 0 ? ix('접수기간') : ix('신청일자'), records = [];
      for (const r of rows.slice(movementHeader + 1)) {
        const name = clean(r[ix('성명')]), no = clean(r[ix('번호')]), d = date(r[dateCol]), cat = category(`${r[ix('신고구분')] || ''}${r[ix('신고종류')] || ''}`);
        if (name && d && cat) records.push({ no, name, date: d, category: cat });
      }
      return { type: 'movement', records };
    }
    const experienceHeader = findHeader(rows, ['성명', '신청서체험기간']);
    if (experienceHeader >= 0) {
      const head = rows[experienceHeader].map(clean), ix = key => head.findIndex(v => v.includes(key)), records = [];
      for (const r of rows.slice(experienceHeader + 1)) {
        const name = clean(r[ix('성명')]), no = clean(r[ix('번호')]);
        if (name) for (const d of schoolDates(expandRange(r[ix('신청서체험기간')]))) records.push({ no, name, date: d, category: '결석|인정' });
      }
      return { type: 'experience', records };
    }
    const officialRecords = parseOfficialOdt(xml, targetClass);
    if (officialRecords.length) return { type: 'official', records: officialRecords };
    return { type: 'generic', records: [] };
  }

  function audit(inputs, XLSX) {
    const issues = [], parsed = { monthly: [], summary: [], evidence: [] }, detected = [];
    for (const input of inputs) {
      if (input.odtContent) {
        const result = parseOdtContent(input.odtContent, input.classNo);
        detected.push({ name: input.name, slot: input.slot, type: `odt-${result.type}` });
        if (result.records.length) parsed.evidence.push(...result.records.map(r => ({ ...r, file: input.name })));
        else issues.push({ severity: 'warning', code: '문서확인', message: `${input.name}에서 자동 대조할 출결 표를 찾지 못했습니다.`, detail: 'ODT의 표에 성명·날짜·출결 구분이 있는지 확인해 주세요.' });
        continue;
      }
      if (input.pdfText) {
        const records = parseSummaryPdfText(input.pdfText);
        detected.push({ name: input.name, slot: input.slot, type: records.length ? 'summary-pdf' : 'pdf-unreadable' });
        if (input.slot !== 'manager' || !records.length) issues.push({ severity: 'error', code: '기준파일형식', message: `${input.name}에서 출결업무 담당자 기준 집계를 읽지 못했습니다.`, detail: '학급별출결현황 표가 포함된 PDF인지 확인해 주세요.' });
        else parsed.summary.push(...records);
        continue;
      }
      if (input.manual) {
        const record = { no: clean(input.no), name: clean(input.name), date: date(input.date), category: clean(input.category), file: clean(input.file) || '수기 기록' };
        detected.push({ name: record.file, slot: input.slot || 'manual', type: 'manual' });
        if (!record.name || !record.date || !LABELS[record.category]) {
          issues.push({ severity: 'error', code: '수기기록형식', message: '수기 기록에 학생 이름·날짜·출결 구분을 모두 입력해 주세요.', detail: record.file });
        } else parsed.evidence.push(record);
        continue;
      }
      if (!/\.(xlsx?|csv)$/i.test(input.name)) {
        detected.push({ name: input.name, slot: input.slot, type: 'document' });
        issues.push({ severity: 'warning', code: '문서확인', message: `${input.name}은 업로드되었지만 자동 표 대조 대상은 아닙니다.`, detail: 'PDF/HWP/HWPX/ODT 공문·체험학습 자료는 사람이 내용을 확인해 주세요.' });
        continue;
      }
      let wb;
      try { wb = XLSX.read(input.data, { type: 'array', cellDates: true }); }
      catch (e) { issues.push({ severity: 'error', code: '파일읽기', message: `${input.name} 파일을 열 수 없습니다.`, detail: e.message }); continue; }
      const type = detect(wb, XLSX);
      detected.push({ name: input.name, slot: input.slot, type });
      if (input.slot === 'manager' && type !== 'summary') issues.push({ severity: 'error', code: '기준파일형식', message: `${input.name}은 출결업무 담당자 기준 파일 형식이 아닙니다.`, detail: '번호·성명과 결석·지각·조퇴·결과 집계 열이 있는 학급별 월결 출결 집계표를 올려 주세요.' });
      if (input.slot === 'monthly' && type !== 'monthly') issues.push({ severity: 'error', code: '반별현황형식', message: `${input.name}은 반별 월결 출결 현황 형식이 아닙니다.`, detail: '일자·번호·성명·출결구분 열이 있는 파일을 올려 주세요.' });
      if (type === 'monthly') {
        const result = parseMonthly(wb, XLSX); parsed.monthly.push(...result.records); issues.push(...result.warnings);
      } else if (type === 'summary') parsed.summary.push(...parseSummary(wb, XLSX));
      else if (type === 'absence') parsed.evidence.push(...parseAbsence(wb, XLSX).map(r => ({ ...r, file: input.name })));
      else if (type === 'movement') parsed.evidence.push(...parseMovement(wb, XLSX).map(r => ({ ...r, file: input.name })));
      else if (type === 'experience') parsed.evidence.push(...parseExperience(wb, XLSX).map(r => ({ ...r, file: input.name })));
      else issues.push({ severity: 'warning', code: '형식미지원', message: `${input.name}에서 비교 가능한 표를 찾지 못했습니다.`, detail: '보관은 가능하지만 자동 대조에서는 제외됩니다.' });
    }
    parsed.monthly = [...new Map(parsed.monthly.map(record => [rowKey(record), record])).values()];
    if (!parsed.summary.length) issues.push({ severity: 'error', code: '기준파일없음', message: '학급별 출결현황 파일이 없습니다.', detail: '번호·성명과 결석·지각·조퇴·결과 집계가 있는 PDF 또는 엑셀 파일을 올려 주세요.' });
    if (!parsed.monthly.length) issues.push({ severity: 'error', code: '반별현황없음', message: '대조할 반별 월결 출결 현황 파일이 없습니다.', detail: '일자·번호·성명·출결구분 열이 있는 반별 파일을 올려 주세요.' });

    const monthlyCounts = new Map();
    for (const r of parsed.monthly) {
      const key = r.no || r.name;
      if (!monthlyCounts.has(key)) monthlyCounts.set(key, { no: r.no, name: r.name, counts: Object.fromEntries(Object.keys(LABELS).map(k => [k, 0])) });
      monthlyCounts.get(key).counts[r.category]++;
    }
    const summaryByStudent = new Map(parsed.summary.map(r => [r.no || r.name, r]));
    if (parsed.summary.length) {
      for (const [key, classReport] of monthlyCounts) {
        const managerBase = summaryByStudent.get(key);
        if (!managerBase) { issues.push({ severity: 'error', code: '담당자명단누락', message: `${classReport.no}번 ${classReport.name}이 출결업무 담당자 기준 파일에 없습니다.`, detail: '' }); continue; }
        for (const cat of Object.keys(LABELS)) if (classReport.counts[cat] !== managerBase.counts[cat]) {
          issues.push({ severity: 'error', code: '집계불일치', message: `${classReport.no}번 ${classReport.name}의 ${LABELS[cat]} 횟수가 다릅니다.`, detail: `담당자 기준 ${managerBase.counts[cat]}회 · 반별 월결 현황 ${classReport.counts[cat]}회` });
        }
      }
      for (const [key, managerBase] of summaryByStudent) if (!monthlyCounts.has(key) && Object.values(managerBase.counts).some(Boolean)) issues.push({ severity: 'error', code: '기준명단누락', message: `${managerBase.no}번 ${managerBase.name}의 출결이 담당자 기준 파일에만 있습니다.`, detail: '' });
    }

    const monthlyMap = new Map(parsed.monthly.map(r => [rowKey(r), r]));
    const evidenceMap = new Map(parsed.evidence.map(r => [rowKey(r), r]));
    const baselineMonths = new Set(parsed.monthly.map(r => r.date.slice(0, 7)));
    if (parsed.evidence.length) {
      for (const [key, r] of monthlyMap) if (requiresParentNice(r) && !evidenceMap.has(key)) issues.push({ severity: 'error', code: '신고서누락', message: `${r.no}번 ${r.name} ${r.date} ${LABELS[r.category]} 신고가 없습니다.`, detail: '학부모 나이스 대조 대상이지만 올린 신고서에서 찾지 못했습니다.' });
      for (const [key, r] of evidenceMap) if (requiresParentNice(r) && !monthlyMap.has(key) && baselineMonths.has(r.date.slice(0, 7))) {
        const same = parsed.monthly.find(m => (m.no === r.no || m.name === r.name) && m.date === r.date);
        issues.push({ severity: 'error', code: same ? '구분불일치' : '기준기록누락', message: `${r.no}번 ${r.name} ${r.date} ${LABELS[r.category]} 기록이 기준과 맞지 않습니다.`, detail: same ? `월별 현황: ${LABELS[same.category]} · 신고서: ${LABELS[r.category]}` : `${r.file}에는 있으나 월별 현황에는 없습니다.` });
      }
    } else if (parsed.monthly.length) issues.push({ severity: 'warning', code: '신고서없음', message: '상세 신고서가 없어 날짜별 대조를 건너뛰었습니다.', detail: '' });

    const menstrual = new Map();
    for (const r of parsed.monthly.filter(r => /생리|생리통/.test(r.reason))) {
      const key = r.no || r.name;
      if (!menstrual.has(key)) menstrual.set(key, { no: r.no, name: r.name, dates: [] });
      menstrual.get(key).dates.push(r.date);
    }
    for (const m of menstrual.values()) if (m.dates.length >= 2) issues.push({ severity: 'error', code: '생리중복', message: `${m.no}번 ${m.name}의 생리/생리통 사유가 ${m.dates.length}회입니다.`, detail: `2회 이상 오류 · ${m.dates.join(', ')}` });

    const errors = issues.filter(i => i.severity === 'error').length, warnings = issues.filter(i => i.severity === 'warning').length;
    return { issues, detected, stats: { monthly: parsed.monthly.length, summaryStudents: parsed.summary.length, evidence: parsed.evidence.length, menstrualStudents: menstrual.size, errors, warnings } };
  }

  return { audit, detect, parseMonthly, parseSummary, parseSummaryPdfText, parseAbsence, parseMovement, parseExperience, parseOfficialOdt, parseOdtContent, requiresParentNice, LABELS };
});
