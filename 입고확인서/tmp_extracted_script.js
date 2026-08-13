












// ══════════════════════════════════════════════
// 탭 전환
// ══════════════════════════════════════════════
document.querySelectorAll('.tab').forEach(function(t) {
  t.addEventListener('click', function() { switchTab(t.dataset.tab); });
});
function switchTab(name) {
  document.querySelectorAll('.tab').forEach(function(e) { e.classList.remove('on'); });
  document.querySelectorAll('.pane').forEach(function(e) { e.classList.remove('on'); });
  var tabEl = document.querySelector('.tab[data-tab="' + name + '"]');
  var paneEl = document.getElementById('tab-' + name);
  if (tabEl)  tabEl.classList.add('on');
  if (paneEl) paneEl.classList.add('on');
}

// ══════════════════════════════════════════════
// 설정
// ══════════════════════════════════════════════
var $id = function(id) { return document.getElementById(id); };

$id('scriptUrl').value = localStorage.getItem('scriptUrl') || '';
$id('sheetId').value   = localStorage.getItem('sheetId')   || '1SLkS88rKFESYVcxMxLqz67YNAp6VTK0Icm9lTH5AYXU';

function saveSettings() {
  var url = $id('scriptUrl').value.trim();
  var sid = $id('sheetId').value.trim();
  if (!url) { alert('Apps Script URL을 입력해주세요'); return; }
  localStorage.setItem('scriptUrl', url);
  if (sid) localStorage.setItem('sheetId', sid);
  var ok = $id('settingOk');
  ok.style.display = 'inline';
  setTimeout(function() { ok.style.display = 'none'; }, 2000);
}

// ══════════════════════════════════════════════
// 파일 분류
// ══════════════════════════════════════════════
var files = new Map();

function classify(name) {
  var n = name.normalize('NFC');
  if (n.includes('수입스케줄') && n.endsWith('.xlsx'))       return {cls:'c-imp', tag:'수입스케줄'};
  if (n.includes('국내상품스케줄') && n.endsWith('.xlsx'))   return {cls:'c-dom', tag:'국내스케줄'};
  if (n.includes('품목확인용') && n.endsWith('.xlsx'))       return {cls:'c-inv', tag:'현재고'};
  if (n.includes('현재고') && (n.endsWith('.xlsx')||n.endsWith('.xls'))) return {cls:'c-inv', tag:'현재고'};
  if (n.includes('품절') && (n.endsWith('.xlsx')||n.endsWith('.xls'))) return {cls:'c-out', tag:'품절현황'};
  if (n.includes('입고수량'))                                   return {cls:'c-xls', tag:'입고수량'};
  return {cls:'c-unk', tag:'기타'};
}

function renderChips() {
  var list = $id('chipList');
  list.innerHTML = '';
  files.forEach(function(_, name) {
    var info = classify(name);
    var d = document.createElement('div');
    d.className = 'chip ' + info.cls;
    d.innerHTML = '<div class="cdot"></div><div class="cname">' + name + '</div><span class="ctag">' + info.tag + '</span><span class="crm" data-n="' + name + '">✕</span>';
    list.appendChild(d);
  });
  list.querySelectorAll('.crm').forEach(function(b) {
    b.onclick = function() { files.delete(b.dataset.n); renderChips(); resetAnalysis(); };
  });
  $id('anaBtn').disabled = files.size === 0;
  $id('fileCnt').textContent = files.size ? files.size + '개' : '';
}

var dz = $id('dropzone');
var fi = $id('fileInput');
dz.addEventListener('dragover', function(e) { e.preventDefault(); dz.classList.add('over'); });
dz.addEventListener('dragleave', function() { dz.classList.remove('over'); });
dz.addEventListener('drop', function(e) {
  e.preventDefault(); dz.classList.remove('over');
  Array.from(e.dataTransfer.files).forEach(function(f) { files.set(f.name, f); });
  renderChips(); resetAnalysis();
});
fi.addEventListener('change', function() {
  Array.from(fi.files).forEach(function(f) { files.set(f.name, f); });
  fi.value = ''; renderChips(); resetAnalysis();
});

function clearAll() {
  files.clear(); renderChips(); resetAnalysis();
  $id('progWrap').classList.remove('show');
  $id('resultBox').className = 'result';
  window._inventoryFileAttached = false;
  window._domesticFileAttached = false;
  window._importFileAttached = false;
  window._stockoutFileAttached = false;
}
function resetAnalysis() {
  $id('anaPanel').classList.remove('show');
  if ($id('runBtn')) $id('runBtn').disabled = false;
}

// ══════════════════════════════════════════════
// 파싱 유틸
// ══════════════════════════════════════════════

// ══════════════════════════════════════════════
// 파싱 유틸
// ══════════════════════════════════════════════

// ══════════════════════════════════════════════
// 파싱 유틸
// ══════════════════════════════════════════════
// 과거에는 CUTOFF(2026-04-01) 이전 행을 통째로 제외했으나,
// 벤더시트 헤더 셀의 수식(예: =수입스케줄!A1916)이 그 이전 행을 직접 가리키는 경우
// schedMap에서 빠져 매칭이 깨지고, 이미 완료된 옛 출하건이 "예정"으로 잘못 표시되는
// 버그가 있었음. 최근입고/다음입고 선택은 buildRows의 날짜 정렬에서 이미 처리되므로
// 여기서는 날짜로 미리 자르지 않는다 (null/유효하지 않은 행만 제외).
var TODAY  = new Date();

// 다이소 PB상품은 "다이소"라는 단어 하나만으로는 매칭 근거가 부족함(컨테이너 적재목록에
// 다이소가 워낙 자주 등장함). 로컬 품목명에 "다이소"가 들어있으면, 후보 텍스트에
// "다이소" + 그 외 의미있는 단어(예: "바다보양식","힘줄육포") 둘 다 있어야 인정한다.
// 예) "다이소 더주스 바다보양식 열빙어" → "다이소"+"바다보양식" 둘 다 후보에 있어야 함
//     "다이소 더주스 치킨 힘줄육포"   → "다이소"+"힘줄육포" 둘 다 후보에 있어야 함
function validateDaisoMatch(itemName, candItems) {
  if (String(itemName).indexOf('다이소') < 0) return true; // 다이소 상품이 아니면 이 검증 대상 아님
  if (!candItems) return false;
  var candStr = String(candItems).toLowerCase().replace(/\s+/g,'');
  if (candStr.indexOf('다이소') < 0) return false; // "다이소" 자체가 후보에 없으면 바로 탈락
  var candWords = splitWords(candItems);
  var nWords = splitWords(itemName);
  for (var i = 0; i < nWords.length; i++) {
    var w = nWords[i];
    if (w === '다이소' || w.length < 3) continue;
    if (/^[0-9]+[a-z]*$/.test(w)) continue; // 숫자/용량 표기(예: "70g","100ml") 제외
    if (GENERIC_WORDS.indexOf(w) >= 0) continue;
    if (candWords.indexOf(w) >= 0) return true; // "다이소" 외 다른 의미있는 단어도 겹침 → 인정
  }
  return false; // "다이소"만 겹치고 다른 근거가 없음 → 불인정
}

// 품목명을 "단어" 단위로 쪼갠다 (공백/괄호/기호 기준). 후보 키워드가 단일 단어일 때
// 이 단어 목록에 "완전히 똑같은 단어"로 있는지 확인하는 용도 — 부분포함(substring)
// 매칭은 "완두콩과양고기"처럼 단어 일부에 우연히 끼어있는 경우까지 다 잡아버려서 너무 느슨함.
function splitWords(str) {
  return String(str).toLowerCase().split(/[\s,\/()\[\]{}\-_~+&·,]+/).filter(function(w){ return w.length > 0; });
}

// 컨테이너 적재목록(품목 비고)에 흔히 같이 등장하지만 특정 제품을 식별하는 데는
// 쓸 수 없는 범용 재료/브랜드/포장 단어 목록. 글자수만으로 거르기 어려운
// 5자 이상 단어들도 여기서 함께 차단한다.
var GENERIC_WORDS = [
  '소고기','닭고기','양고기','닭가슴살','오리고기','연어','오리','치킨','딩고','개빼로',
  '트레이더스','몰리스','다이소','저온','수제','더펫','일킬로그램','터키츄','일킬로','마이독',
  '비스코티','황태','죽탕','시크릿크런치','짜먹는고양이간식','고구마말랭이','버블제로',
  '컨츄리치킨랩','덕텐돈','온리','생고기육포','슈퍼츄','다이소마이독'
];

function fd(d) {
  if (!d) return '';
  var dt = (d instanceof Date) ? d : new Date(d);
  if (isNaN(dt)) return '';
  return dt.getFullYear() + '-' + 
         String(dt.getMonth() + 1).padStart(2, '0') + '-' + 
         String(dt.getDate()).padStart(2, '0');
}
function fmtD(d) {
  if (!d) return '';
  var dt = (d instanceof Date) ? d : new Date(d);
  if (isNaN(dt)) return '';
  return String(dt.getFullYear()).slice(2) + '.' +
    String(dt.getMonth()+1).padStart(2,'0') + '.' +
    String(dt.getDate()).padStart(2,'0');
}
function nowStr() {
  var n = new Date();
  return fmtD(n) + ' ' + String(n.getHours()).padStart(2,'0') + ':' + String(n.getMinutes()).padStart(2,'0');
}
function readXlsx(file) {
  return new Promise(function(res, rej) {
    var r = new FileReader();
    r.onload = function(e) {
      try { res(XLSX.read(e.target.result, {type:'array', cellDates:false, cellNF:true})); }
      catch(err) { rej(err); }
    };
    r.onerror = rej;
    r.readAsArrayBuffer(file);
  });
}

// 엑셀 날짜 직렬번호 → Date (cellDates:false일 때)
function excelSerialToDate(v) {
  if (!v && v !== 0) return null;
  var n = parseFloat(v);
  if (isNaN(n) || n < 40000 || n > 60000) return null; // 2009~2064 범위
  var d = new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 86400000);
  return (d.getFullYear() > 2000 && d.getFullYear() < 2040) ? d : null;
}

// 날짜 직렬번호 OR 텍스트 모두 처리 (선적일 등 텍스트 허용)
// 숫자면 Date 반환, 텍스트면 문자열 그대로 반환, 빈값이면 null
function excelDateOrText(v) {
  if (v === null || v === undefined || v === '') return null;
  var s = String(v).trim();
  if (!s || s === '-') return null;
  var n = parseFloat(s);
  if (!isNaN(n) && n >= 40000 && n <= 60000) {
    var d = new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 86400000);
    if (d.getFullYear() > 2000 && d.getFullYear() < 2040) return d;
  }
  // 텍스트로 반환
  return s;
}

// 날짜 OR 텍스트 값을 문자열로 포맷
function fdOrText(v) {
  if (!v) return '';
  if (v instanceof Date) return fd(v);
  return String(v); // 텍스트 그대로
}

// ── 현재고 파싱 ──────────────────────────────
async function parseInventory(file) {
  var wb   = await readXlsx(file);
  var ws   = wb.Sheets[wb.SheetNames[0]];
  var data = XLSX.utils.sheet_to_json(ws, {header:1, defval:'', raw:true});
  var items = {}, disc = new Set(), dupes = [];
  if (data.length < 2) return {items, disc, dupes};

  var hdr = data[0];
  var CI = {};
  for (var i = 0; i < hdr.length; i++) {
    var h = String(hdr[i]||'').trim();
    if (h === '바코드' || h === '자사바코드' || h === '상품코드' || h === '관리코드') {
      if (CI.bc == null || h === '바코드' || h === '자사바코드') CI.bc = i;
    }
    if (h === '상품명' || h === '품명') {
      if (CI.name == null || h === '상품명') CI.name = i;
    }
    if (h === '현재고' || h === '가용재고' || h === '정상재고' || h === '재고수량' || h === '재고' || h === '수량') {
      if (CI.stock == null || h === '현재고' || h === '정상재고' || h === '가용재고') CI.stock = i;
    }
    if (h === '품절' || h === '단종' || h === '단종여부' || h === '상품상태' || h === '재고상태') {
      if (CI.sout == null || h === '품절' || h === '상품상태' || h === '단종여부') CI.sout = i;
    }
  }
  // 현재고_YYYYMMDD_HHMMSS 형식: E열(4)=바코드, M열(12)=현재고수량 (강제 고정)
  if (/현재고_\d/.test(file.name.normalize('NFC'))) {
    CI.bc    = 4;
    CI.stock = 12;
  } else {
    if (CI.bc    == null) CI.bc    = 7;  // 품목확인용: 자사바코드 H열
    if (CI.name  == null) CI.name  = 9;
    if (CI.stock == null) CI.stock = 12;
    if (CI.sout  == null) CI.sout  = 14;
  }

  for (var ri = 1; ri < data.length; ri++) {
    var row   = data[ri];
    var bcRaw = String(row[CI.bc]||'').trim();
    if (!bcRaw || bcRaw === 'undefined') continue;
    if (bcRaw.indexOf('.') >= 0) {
      try { bcRaw = String(parseInt(parseFloat(bcRaw))); } catch(e) {}
    }

    var name  = String(row[CI.name]  || '').trim();
    var stock = Number(row[CI.stock]) || 0;
    var sout  = String(row[CI.sout]  || '').trim();
    var isDisc = name.indexOf('(단종)') >= 0 || name.indexOf('[단종]') >= 0 || sout === '단종';
    if (isDisc) { disc.add(bcRaw); continue; }

    // 중복 바코드 감지
    if (items[bcRaw]) {
      // 중복이면 바코드 뒤에 _DUP 표시해서 둘 다 보존
      dupes.push(bcRaw);
      var dupKey = bcRaw + '_중복' + dupes.filter(function(d){return d===bcRaw;}).length;
      items[dupKey] = {name:'[중복] ' + name, stock:stock, isDupe:true, origBc:bcRaw};
    } else {
      items[bcRaw] = {name:name, stock:stock};
    }
  }
  return {items, disc, dupes};
}

// ── 품절 파싱 ────────────────────────────────
async function parseStockout(file) {
  var normName = file.name.normalize('NFC');
  // 파일명(확장자 제외)을 그대로 품절보고일로 사용: "0521품절.xlsx" → "0521품절"
  var rawName = normName.replace(/\.[^.]+$/, '');

  var wb   = await readXlsx(file);
  var ws   = wb.Sheets[wb.SheetNames[0]];
  var data = XLSX.utils.sheet_to_json(ws, {header:1, defval:'', raw:true});
  var out  = {}; // bc → 품절보고일
  if (data.length < 2) return out;

  // 헤더에서 바코드 열 찾기
  var hdr   = data[0];
  var bcIdx = 0;
  for (var i = 0; i < hdr.length; i++) {
    if (String(hdr[i]||'').trim() === '바코드') { bcIdx = i; break; }
  }

  for (var ri = 1; ri < data.length; ri++) {
    var row = data[ri];
    var bc  = String(row[bcIdx]||'').trim();
    if (!bc || bc === 'undefined') continue;
    try { if (bc.indexOf('.') >= 0) bc = String(parseInt(parseFloat(bc))); } catch(e) {}
    if (!/^\d/.test(bc)) continue;
    out[bc] = rawName;
  }
  return out;
}

// ── 날짜 파싱 유틸 ──────────────────────────
function parseHeaderDate(hv) {
  if (!hv) return null;
  var d = excelSerialToDate(hv);
  if (d) return d;
  var s = String(hv).trim();
  var m1 = s.match(/(\d{1,2})[\/\.-](\d{1,2})[\/\.-](\d{4})/);
  if (m1) return new Date(m1[3], m1[1]-1, m1[2]);
  var m2 = s.match(/(\d{4})[\/\.-](\d{1,2})[\/\.-](\d{1,2})/);
  if (m2) return new Date(m2[1], m2[2]-1, m2[3]);
  
  var m3 = s.match(/(\d{1,2})[\/\.-](\d{1,2})/);
  if (m3) {
    var cy = new Date().getFullYear();
    return new Date(cy, m3[1]-1, m3[2]);
  }
  
  var m4 = s.match(/(\d{1,2})월\s*(\d{1,2})일/);
  if (m4) {
    var cy = new Date().getFullYear();
    return new Date(cy, m4[1]-1, m4[2]);
  }
  
  return null;
}

// ── 국내상품스케줄 파싱 ──────────────────────
// 2행 날짜 열 헤더의 괄호/줄바꿈 유무에 따라 확정일과 대기(예정)일로 분류
async function parseDomestic(file) {
  var wb = await readXlsx(file);
  var domByBc = {};
  var sheetNames = ['펫원','원발란스','샌드원'];
  window._petoneMemoCount = 0;

  function cleanDomesticSheetName(name) {
    return String(name || '').replace(/\s+/g, '').trim();
  }
  function domesticWorkbookSheetName(target) {
    var cleanTarget = cleanDomesticSheetName(target);
    for (var i = 0; i < wb.SheetNames.length; i++) {
      if (cleanDomesticSheetName(wb.SheetNames[i]) === cleanTarget) return wb.SheetNames[i];
    }
    return '';
  }
  function isPetoneSheet(name) {
    return cleanDomesticSheetName(name) === '펫원';
  }

  function domesticCellDate(v) {
    if (v instanceof Date) return v;
    if (typeof v === 'number' && isFinite(v)) return excelSerialToDate(v);
    return parseHeaderDate(v);
  }
  function domesticCellText(v) {
    if (v instanceof Date) return fd(v);
    var d = domesticCellDate(v);
    if (d) return fd(d);
    return String(v || '').trim();
  }
  function buildPetoneNote(requestDate, shareDate, reason) {
    var parts = [];
    var req = domesticCellText(requestDate);
    var shared = domesticCellText(shareDate);
    var rsn = String(reason || '').trim();
    if (req) parts.push('입고요청일(리드타임기준): ' + req);
    if (shared) parts.push('펫원 공유 입고예정일: ' + shared);
    if (rsn) parts.push('지연사유 및 특이사항: ' + rsn);
    return parts.join(' / ');
  }
  function findHeaderCol(headerRow, names) {
    for (var i = 0; i < headerRow.length; i++) {
      var h = String(headerRow[i] || '').replace(/\s+/g, '');
      for (var n = 0; n < names.length; n++) {
        if (h === String(names[n]).replace(/\s+/g, '')) return i;
      }
    }
    return -1;
  }

  for (var si = 0; si < sheetNames.length; si++) {
    var sname = domesticWorkbookSheetName(sheetNames[si]);
    if (!sname) continue;
    var ws   = wb.Sheets[sname];
    var data = XLSX.utils.sheet_to_json(ws, {header:1, defval:'', raw:true});

    var hi = -1;
    for (var ri = 0; ri < data.length; ri++) {
      if (String(data[ri][0]).trim() === 'NO' && String(data[ri][1]).trim() === '바코드') {
        hi = ri; break;
      }
    }
    if (hi < 0) continue;
    var header = data[hi];
    var petoneRequestCol = -1, petoneShareCol = -1, petoneReasonCol = -1;
    if (isPetoneSheet(sname)) {
      petoneRequestCol = findHeaderCol(header, ['입고요청일(리드타임기준)','입고요청일']);
      petoneShareCol = findHeaderCol(header, ['펫원공유입고예정일','펫원 공유 입고예정일']);
      petoneReasonCol = findHeaderCol(header, ['지연사유및특이사항','지연사유 및 특이사항']);
      if (petoneRequestCol < 0) petoneRequestCol = 130; // EA fallback
      if (petoneShareCol < 0) petoneShareCol = 131;   // EB fallback
      if (petoneReasonCol < 0) petoneReasonCol = 133; // ED fallback
    }

    var confirmCols = {}; // ci → date
    var pendingCols = {}; // ci → date

    for (var ci = 2; ci < header.length; ci++) {
      var hv = header[ci];
      var d = parseHeaderDate(hv);
      if (!d) continue;
      var s = String(hv).trim();
      var isPending = s.indexOf('(') >= 0 || s.indexOf('\n') >= 0 || s.indexOf('\r') >= 0;
      if (isPending) {
        pendingCols[ci] = d;
      } else {
        confirmCols[ci] = d;
      }
    }

    for (var ri = hi+1; ri < data.length; ri++) {
      var row   = data[ri];
      var bcRaw = String(row[1]||'').trim();
      if (!bcRaw) continue;
      try { if (bcRaw.indexOf('.') >= 0) bcRaw = String(parseInt(parseFloat(bcRaw))); } catch(e) {}
      if (!/^\d/.test(bcRaw)) continue;

      var name = String(row[2]||'').trim();
        if (!domByBc[bcRaw]) domByBc[bcRaw] = {name:name, sheet:sname, done:[], pend:[]};
        else if (name.length > (domByBc[bcRaw].name||'').length) domByBc[bcRaw].name = name;

      if (isPetoneSheet(sname)) {
        var petoneRequestDate = row[petoneRequestCol]; // EA: 입고요청일(리드타임기준)
        var petoneShareDate = row[petoneShareCol];     // EB: 펫원 공유 입고예정일
        var petoneReason = row[petoneReasonCol];       // ED: 지연사유 및 특이사항
        var petoneNote = buildPetoneNote(petoneRequestDate, petoneShareDate, petoneReason);
        if (petoneNote) {
          domByBc[bcRaw].petoneNote = petoneNote;
          window._petoneMemoCount++;
        }
        var petoneShareD = domesticCellDate(petoneShareDate);
        if (petoneShareD) {
          domByBc[bcRaw].petoneShareDate = petoneShareD;
          domByBc[bcRaw].petoneReason = String(petoneReason || '').trim();
          domByBc[bcRaw].petoneRequestDate = domesticCellDate(petoneRequestDate);
        }
      }

      var confKeys = Object.keys(confirmCols);
      for (var ck = 0; ck < confKeys.length; ck++) {
        var ci2 = parseInt(confKeys[ck]);
        var rawVal = String(row[ci2]||'').trim();
        var qtyStr = rawVal.replace(/[^0-9\-]/g, '');
        var qty = parseInt(qtyStr);
        
        var dCellAddr = XLSX.utils.encode_cell({r: ri, c: ci2});
        var dCellObj  = ws[dCellAddr];
        var dMemoText = '';
        if (dCellObj && dCellObj.c && dCellObj.c.length > 0) {
          dMemoText = dCellObj.c.map(function(cm){ return cm.t || ''; }).join(' ').trim();
        }
        
        var hasContent = (rawVal !== '' && rawVal !== '-' && rawVal !== '0');
        var hasMemo = (dMemoText !== '');
        if (!hasContent && !hasMemo) continue;

        domByBc[bcRaw].done.push({date: confirmCols[ci2], qty: (qty > 0) ? qty : rawVal, memo: dMemoText});
      }

      var pendKeys = Object.keys(pendingCols);
      for (var pk = 0; pk < pendKeys.length; pk++) {
        var ci2 = parseInt(pendKeys[pk]);
        var rawVal = String(row[ci2]||'').trim();
        var qtyStr = rawVal.replace(/[^0-9\-]/g, '');
        var qty = parseInt(qtyStr);
        
        var cellAddr = XLSX.utils.encode_cell({r: ri, c: ci2});
        var cellObj  = ws[cellAddr];
        var memoDate = null;
        var memoText = '';
        if (cellObj && cellObj.c && cellObj.c.length > 0) {
          memoText = cellObj.c.map(function(cm){ return cm.t || ''; }).join(' ').trim();
          var mMatch = memoText.match(/(\d{1,2})\/(\d{1,2})/);
          if (mMatch) {
            var mMon = parseInt(mMatch[1]);
            var mDay = parseInt(mMatch[2]);
            var mYear = pendingCols[ci2].getFullYear();
            memoDate = new Date(mYear, mMon - 1, mDay);
          }
        }
        
        var hasContent = (rawVal !== '' && rawVal !== '-' && rawVal !== '0');
        var hasMemo = (memoText !== '');
        if (!hasContent && !hasMemo) continue;

        var headerText = String(header[ci2]).trim();
        domByBc[bcRaw].pend.push({
          date: memoDate || pendingCols[ci2],
          qty: (qty > 0) ? qty : rawVal,
          headerDate: pendingCols[ci2],
          headerText: headerText,
          memo: memoText
        });
      }
    }
  }
  return domByBc;
}

// ── 수입스케줄 파싱 ──────────────────────────
// 수식(f) + 값(v) 모두 지원
async function parseImport(file) {
  var wb = await readXlsx(file);
  if (wb.SheetNames.indexOf('수입스케줄') < 0) return {};

  var ws_s   = wb.Sheets['수입스케줄'];
  var sData  = XLSX.utils.sheet_to_json(ws_s, {header:1, defval:'', raw:true});

  // schedMap: row번호 → 컨테이너 정보
  // schedBySerial: 날짜직렬번호 → row번호 (값 기반 매칭용)
  var schedMap = {}, schedBySerial = {};

  for (var ri = 2; ri < sData.length; ri++) {
    var row    = sData[ri];
    var ship_d = excelDateOrText(row[0]); // 텍스트(5월선적예정 등) 허용
    var arr_d  = excelDateOrText(row[1]); // 텍스트 허용
    var inbd_d = excelDateOrText(row[2]); // 입고일: 날짜뿐 아니라 "입고가능" 같은 상태 텍스트도 허용
    var bl     = String(row[4]||'').trim();
    var vendor = String(row[5]||'').trim();
    var orderNo= String(row[8]||'').trim();
    var items  = String(row[9]||'').trim();
    var ecount = String(row[15]||'').trim(); // P열: 이카운트 정리 (O/X, 또는 "(O)통관전" 등)
    var memo   = String(row[16]||'').trim();
    // "(O)통관전" / "(O)송금전" 처럼 O 뒤에 괄호로 사유가 붙은 경우 그 사유만 추출
    var ecountNote = '';
    var ecMatch = ecount.match(/^\(O\)\s*(.+)$/i);
    if (ecMatch) ecountNote = ecMatch[1].trim();
    // best: 날짜 기준 정렬용 (텍스트면 ship_d로 대체)
    var arr_date  = (arr_d  instanceof Date) ? arr_d  : null;
    var ship_date = (ship_d instanceof Date) ? ship_d : null;
    var inbd_date = (inbd_d instanceof Date) ? inbd_d : null;
    var best   = inbd_date || arr_date || ship_date || (ship_d ? new Date('2099-01-01') : null);
    if (!best) continue;

    var done = !!(inbd_date && inbd_date <= TODAY); // 실제 날짜로 기재되고 오늘 이전일 때만 완료
    var est, estCalc = false;
    if (inbd_d) { est = inbd_d; estCalc = false; } // 날짜든 "입고가능" 같은 텍스트든 실제 기재값 → 추정 아님, 그대로 보존
    else if (arr_date) { est = new Date(arr_date.getTime() + 7*86400000); estCalc = true; } // 입항일+7일 추정
    else { est = null; }
    var rn   = ri + 1;
    schedMap[rn] = {ship:ship_d, arrive:arr_d, inbound:inbd_d,
                    bl:bl, vendor:vendor, orderNo:orderNo, items:items, done:done, est:est,
                    estCalc:estCalc, ecountNote:ecountNote, memo:memo};

    // 날짜 직렬번호로도 인덱싱 (값 기반 매칭용, 동일 날짜 다수 벤더 대비 배열 저장)
    function addSerial(val, rn) {
      if (val == null || val === '') return;
      var s = String(val).trim();
      if (!schedBySerial[s]) schedBySerial[s] = [];
      schedBySerial[s].push(rn);
    }
    addSerial(row[0], rn);
    addSerial(row[1], rn);
    addSerial(row[2], rn);
  }
  window._schedMap = schedMap;

  // 업체별 시트 파싱
  var impByBc = {};
  var SKIP = {수입스케줄:1, Sheet1:1, 현물검정:1, 국내상품스케줄:1};

  for (var si = 0; si < wb.SheetNames.length; si++) {
    var sname = wb.SheetNames[si];
    if (SKIP[sname]) continue;
    var wsRaw = wb.Sheets[sname];
    if (!wsRaw || !wsRaw['!ref']) continue;

    var range = XLSX.utils.decode_range(wsRaw['!ref']);
    if (range.e.r < 3) continue;

    function isVendorMatch(cand, sName) {
      if (!cand || !sName) return false;
      var v = String(cand.vendor||'').toLowerCase().replace(/\s+/g,'');
      var items = String(cand.items||'').toLowerCase().replace(/\s+/g,'');
      var s = String(sName).toLowerCase().replace(/\s+/g,'');
      if (v && (v.indexOf(s) >= 0 || s.indexOf(v) >= 0)) return true;
      if (items && items.indexOf(s) >= 0) return true;
      return false;
    }

    function isItemMatch(itemName, candItems) {
      if (!itemName || !candItems) return false;
      var n = String(itemName).toLowerCase().replace(/\s+/g,'');
      var nWords = splitWords(itemName);
      var parts = String(candItems).split(/[,/]/);
      for (var i=0; i<parts.length; i++) {
          var pRaw = parts[i].trim();
          var p = pRaw.toLowerCase().replace(/\s+/g,'');
          if (p.length < 5) continue; // 4글자 이하는 "양고기","소고기","닭고기","딩고","치킨","오리" 등
          // 컨테이너 적재목록에 흔히 같이 등장하는 범용 재료/브랜드명이라 매칭 근거로 너무 약함
          if (GENERIC_WORDS.indexOf(p) >= 0) continue; // 5글자 이상이어도 알려진 범용 단어는 제외
          if (pRaw.indexOf(' ') >= 0) {
            // 토큰이 여러 단어로 된 구(句) — 공백 제거 후 부분포함 허용 (구 전체가 그대로 들어있어야 하므로 충분히 엄격)
            if (n.indexOf(p) >= 0) return true;
          } else {
            // 단일 단어 — 품목명 안에 "완전히 같은 단어"로 있을 때만 인정 (부분포함 X)
            if (nWords.indexOf(p) >= 0) return true;
          }
      }
      return false;
    }

    // row2(인덱스1): 수식/값으로 cands 후보군 수집
    var colCandsData = {}; // ci -> { cands: [], fmInfo: obj, isDoneCol: boolean, h1: string }
    for (var ci = range.s.c+2; ci <= range.e.c; ci++) {
      var addr = XLSX.utils.encode_cell({r:1, c:ci});
      var cell = wsRaw[addr];
      if (!cell) continue;

      var cell1 = wsRaw[XLSX.utils.encode_cell({r:0, c:ci})];
      var h1Text = cell1 ? String(cell1.v||'').trim() : '';
      var h2Text = String(cell.v||'').trim();
      var isDoneCol = false;
      if ((h1Text.indexOf('입고') >= 0 || h1Text.indexOf('현물접수') >= 0 || h2Text.indexOf('입고') >= 0) && h1Text.indexOf('예정') < 0 && h2Text.indexOf('예정') < 0) {
        isDoneCol = true;
      }

      var fmInfo = null;
      var fm = null;
      if (cell.f) {
        fm = cell.f.match(/'?수입스케줄'?!(?:\$)?A(?:\$)?(\d+)/);
        if (fm) fmInfo = schedMap[parseInt(fm[1])] || null;
      }

      var cands = [];
      
      if (fmInfo && fm) {
        // 수식이 있으면 무조건 해당 행만 매칭 후보로 설정 (날짜로 재검색하지 않음)
        cands = [parseInt(fm[1])];
      } else {
        // 수식이 없을 때만 기존처럼 날짜 기반으로 후보군 검색
        if (cell.v != null) {
          var serial = String(cell.v).trim();
          cands = schedBySerial[serial] || [];
        }
        
        if (cands.length === 0) {
          var cellDate = parseHeaderDate(cell.v);
          if (!cellDate && cell.w) cellDate = parseHeaderDate(cell.w);
          
          if (cellDate) {
            var rKeys = Object.keys(schedMap);
            for (var k = 0; k < rKeys.length; k++) {
              var candCand = schedMap[rKeys[k]];
              if (candCand.ship && fd(candCand.ship) === fd(cellDate)) {
                cands.push(rKeys[k]);
              } else if (candCand.arrive && fd(candCand.arrive) === fd(cellDate)) {
                cands.push(rKeys[k]);
              }
            }
          }
        }
      }

      if (cands.length > 0 || fmInfo || isDoneCol) {
         colCandsData[ci] = { cands: cands, fmInfo: fmInfo, isDoneCol: isDoneCol, h1: h1Text };
      }
    }
    if (!Object.keys(colCandsData).length) continue;

    // 데이터 행 (row4~, 인덱스3~)
    for (var ri = 3; ri <= range.e.r; ri++) {
      var bcCell = wsRaw[XLSX.utils.encode_cell({r:ri, c:range.s.c})];
      if (!bcCell || bcCell.v == null) continue;
      var bc = String(bcCell.v).trim();
      if (bc.indexOf(' ') > 0) bc = bc.split(/\s+/)[0];
      if (!bc || !/^\d/.test(bc)) continue;
      if (bc.indexOf('.') >= 0) { try { bc = String(parseInt(parseFloat(bc))); } catch(e){} }

      var nmCell = wsRaw[XLSX.utils.encode_cell({r:ri, c:range.s.c+1})];
      var nm = nmCell ? String(nmCell.v||'').trim() : '';

      if (!impByBc[bc]) impByBc[bc] = {name:nm, vendor:'', done:[], pend:[]};
      else if (nm.length > (impByBc[bc].name||'').length) impByBc[bc].name = nm;

      var ciKeys = Object.keys(colCandsData);
      for (var ck = 0; ck < ciKeys.length; ck++) {
        var ci2   = parseInt(ciKeys[ck]);
        var qCell = wsRaw[XLSX.utils.encode_cell({r:ri, c:ci2})];
        if (!qCell || qCell.v == null) continue;
        var qty = parseInt(qCell.v);
        if (!qty || qty <= 0) continue;

        var candData = colCandsData[ci2];
        var cands = candData.cands;
        var fmInfo = candData.fmInfo;
        var info2 = null;

        // 1. 품목명(itemName)과 유사도 매칭 (가장 중요)
        if (nm) {
          for (var i = 0; i < cands.length; i++) {
             var candInfo = schedMap[cands[i]];
             if (candInfo && isItemMatch(nm, candInfo.items)) {
               info2 = candInfo; break;
             }
          }
        }
        
        // 2. 실패 시, 벤더/시트 이름으로 매칭
        // ⚠ 후보가 2개 이상이면(같은 날짜에 서로 다른 컨테이너/품목이 여러 건 있는 경우)
        //    벤더명만으로는 어느 게 진짜인지 구분이 안 되므로 추측하지 않는다.
        //    (이게 서로 무관한 여러 품목에 동일한 엉뚱한 품목텍스트가 붙는 버그의 원인이었음)
        if (!info2 && cands.length === 1) {
          var candInfo0 = schedMap[cands[0]];
          if (candInfo0 && isVendorMatch(candInfo0, sname)) {
            info2 = candInfo0;
          }
        }
        
        // 3. 실패 시, 수식(fmInfo) 강제 적용
        if (!info2 && fmInfo) {
           info2 = fmInfo;
        }

        // ★ 다이소 상품 추가 검증: "다이소" 단어만으로 매칭된 거면(다른 의미있는 단어가
        //   안 겹치면) 수식으로 강제 연결된 것이라도 신뢰하지 않고 매칭 실패로 되돌린다.
        if (info2 && !validateDaisoMatch(nm, info2.items)) {
           info2 = null;
        }

        var isMasterMatched = !!info2; // 1~3단계(품목명/벤더/수식)로 schedMap 매칭된 경우만 true
        if (!info2) {
           // 마지막 폴백: 헤더에 수식도 없고 schedMap 매칭도 안 되는 경우.
           // 벤더시트 헤더는 보통 '선적일'이므로 arrive를 같은 값으로 채우면
           // buildRows에서 (입항+7일) 규칙이 잘못 발동한다 → ship만 채우고
           // arrive는 비워서 (선적+14일) 규칙이 적용되게 한다.
           var dt = null;
           if (cands && cands.length > 0) { var fc = schedMap[cands[0]]; if(fc) dt = fc.ship || fc.arrive; } else { var hd = wsRaw[XLSX.utils.encode_cell({r:1, c:ci2})]; if(hd) dt = parseHeaderDate(hd.v) || parseHeaderDate(hd.w); }
           // 벤더는 빈 문자열로 두지 않는다: 이 데이터가 들어있는 시트 이름 자체가
           // 곧 업체명이므로(예: 'Wisetail' 시트 안의 행) sname을 그대로 사용한다.
           info2 = { bl:'', vendor:sname, ship: dt, arrive: null, inbound: null, est: null, qty: null, xlsQty: qty, items: nm, done: false };
        }
        // 헤더 라벨에 '입고'가 들어있다고 해서 무조건 완료(done) 처리하면 안 된다.
        // 예: "8/26입고" 컬럼이라도 그 날짜가 아직 오지 않았으면(TODAY 이전이 아니면) 미완료 상태다.
        var headerDate = parseHeaderDate(candData.h1);
        var inboundCandidate = info2.inbound || headerDate;
        var isDoneByDate = (inboundCandidate instanceof Date) && (inboundCandidate <= TODAY);
        var entryDone = info2.done || (candData.isDoneCol && isDoneByDate);
        // 마스터시트(수입스케줄)에 정식 등록되지 않은 "예정"(미확정) 데이터는 표시하지 않는다.
        // (이미 지나서 완료로 인정되는 건은 실제 발생한 입고이므로 마스터 미등록이어도 그대로 인정)
        if (!entryDone && !isMasterMatched) continue;
        var entryInbound = entryDone ? (inboundCandidate || new Date()) : info2.inbound;
        var entryEst = info2.est;
        if (!entryDone && candData.isDoneCol && inboundCandidate instanceof Date) {
          // 헤더는 "입고"라고 되어 있지만 날짜가 아직 안 됨 → 입고 "예정"일로 처리
          entryEst = inboundCandidate;
        }
        var entry = {
          bl:     info2.bl,      vendor:  info2.vendor,
          ship:   info2.ship,    arrive:  info2.arrive,
          inbound:entryInbound,  est:     entryEst,
          estCalc: info2.estCalc, ecountNote: info2.ecountNote,
          qty:    qty,           items:   info2.items,
          memo:   info2.memo,    matched: isMasterMatched,
          done:   entryDone,     sheet:   sname
        };
        if (entryDone)  impByBc[bc].done.push(entry);
        else            impByBc[bc].pend.push(entry);
      }
    }
  }
  return impByBc;
}

// ── 입고수량 XLS 파싱 ────────────────────────
async function parseInboundXls(file) {
  var normName = file.name.normalize('NFC');
  var mn = normName.match(/^(\d{8})_(.+?)_([\d-]+)__(.+?)_입고수량/);
  if (!mn) return [];
  var dateStr = mn[1], vendor = mn[2], seq = mn[3], bl = mn[4];
  var fileDate = new Date(dateStr.slice(0,4)+'-'+dateStr.slice(4,6)+'-'+dateStr.slice(6,8));

  var wb   = await readXlsx(file);
  var ws   = wb.Sheets[wb.SheetNames[0]];
  var data = XLSX.utils.sheet_to_json(ws, {defval:'', raw:true});
  var entries = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var bc  = String(row['바코드']||'').trim();
    if (!bc || bc === 'NaN') continue;
    if (bc.indexOf('.') >= 0) { try { bc = String(parseInt(parseFloat(bc))); } catch(e){} }
    if (!/^\d/.test(bc)) continue;
    var qty = parseInt(row['낱개수량']||0);
    if (!qty || qty <= 0) continue;
    entries.push({bc:bc, name:String(row['상품명']||'').trim(),
                  vendor:vendor, seq:seq, bl:bl, fileDate:fileDate, qty:qty, inboundDate:fileDate});
  }
  return entries;
}

// ── 입고수량 BL 매칭 적용 ────────────────────
function applyInboundXls(impByBc, xlsEntries) {
  for (var i = 0; i < xlsEntries.length; i++) {
    var e   = xlsEntries[i];
    var bc  = e.bc, bl = e.bl;
    var imp = impByBc[bc];

    if (imp) {
      if (!Array.isArray(imp.pend)) imp.pend = [];
      if (!Array.isArray(imp.done)) imp.done = [];
      var matched = false;
      // 1순위: BL 매칭
      for (var j = 0; j < imp.pend.length; j++) {
        if (imp.pend[j].bl === bl) {
          imp.pend[j].inbound = e.inboundDate;
          imp.pend[j].done    = true;
          imp.pend[j].xlsQty  = e.qty;
          imp.done.push(imp.pend.splice(j,1)[0]);
          matched = true; break;
        }
      }
      // 2순위: 수량 ±5% 매칭
      if (!matched) {
        for (var j = 0; j < imp.pend.length; j++) {
          if (Math.abs(imp.pend[j].qty - e.qty) <= e.qty * 0.05) {
            imp.pend[j].inbound = e.inboundDate;
            imp.pend[j].done    = true;
            imp.pend[j].xlsQty  = e.qty;
            imp.pend[j].bl      = imp.pend[j].bl || bl;
            imp.done.push(imp.pend.splice(j,1)[0]);
            matched = true; break;
          }
        }
      }
      // 3순위: 신규 완료 entry 추가
      if (!matched) {
        imp.done.push({bl:bl, vendor:e.vendor, ship:null, arrive:null,
                       inbound:e.inboundDate, est:null, qty:e.qty, xlsQty:e.qty,
                       items:e.name, done:true, sheet:''});
      }
    } else {
      impByBc[bc] = {name:e.name, vendor:e.vendor, done:[{
        bl:bl, vendor:e.vendor, ship:null, arrive:null,
        inbound:e.inboundDate, est:null, qty:e.qty, xlsQty:e.qty,
        items:e.name, done:true, sheet:''
      }], pend:[]};
    }
  }
}


// ── buildRows: 통합 행 생성 ──────────────────
function buildRows(impByBc, domByBc, items, disc, stockout, updatedAt) {
  function isItemMatchGlobal(itemName, candItems) {
    if (!itemName || !candItems) return false;
    var n = String(itemName).toLowerCase().replace(/\s+/g,'');
    var nWords = splitWords(itemName);
    var parts = String(candItems).split(/[,/]/);
    for (var i=0; i<parts.length; i++) {
        var pRaw = parts[i].trim();
        var p = pRaw.toLowerCase().replace(/\s+/g,'');
        if (p.length < 5) continue; // 4글자 이하는 "양고기","소고기","닭고기","딩고","치킨","오리" 등
          // 컨테이너 적재목록에 흔히 같이 등장하는 범용 재료/브랜드명이라 매칭 근거로 너무 약함
          if (GENERIC_WORDS.indexOf(p) >= 0) continue; // 5글자 이상이어도 알려진 범용 단어는 제외
        if (pRaw.indexOf(' ') >= 0) {
          // 토큰이 여러 단어로 된 구(句) — 공백 제거 후 부분포함 허용
          if (n.indexOf(p) >= 0) return true;
        } else {
          // 단일 단어 — 품목명 안에 "완전히 같은 단어"로 있을 때만 인정 (부분포함 X)
          if (nWords.indexOf(p) >= 0) return true;
        }
    }
    return false;
  }
  var allBc = {}, keys, i;
  keys = Object.keys(items);    for (i=0;i<keys.length;i++) allBc[keys[i]]=1;
  keys = Object.keys(impByBc);  for (i=0;i<keys.length;i++) allBc[keys[i]]=1;
  keys = Object.keys(domByBc);  for (i=0;i<keys.length;i++) allBc[keys[i]]=1;
  keys = Object.keys(stockout); for (i=0;i<keys.length;i++) allBc[keys[i]]=1;
  disc.forEach(function(bc){ allBc[bc]=1; });

  var sortedBc = Object.keys(allBc).sort();
  var rows = [];

  for (var bi = 0; bi < sortedBc.length; bi++) {
    var bc   = sortedBc[bi];
    var inv  = items[bc]    || {};
    var imp  = impByBc[bc]  || {done:[], pend:[]};
    var dom  = domByBc[bc]  || {name:'', sheet:'', done:[], pend:[]};

    // ── 바코드가 국내스케줄/수입스케줄 양쪽 파일 어디에도 없는 경우 ──
    // (둘 중 하나라도 바코드로 매칭됐으면 정상 케이스이므로 건너뜀)
    var bcMissingInDom = !domByBc[bc];
    var bcMissingInImp = !impByBc[bc];
    var orphanLabel = null; // null=정상(바코드 매칭됨), '바코드없음' 또는 '바코드없음(입고내용있음)'
    var ALLOW_KEYWORD_FOR = '스피루리나'; // ★ 키워드 매칭은 이 단어가 품목명에 들어간 경우에만 허용
    if (bcMissingInDom && bcMissingInImp && inv.name) {
      if (inv.name.indexOf(ALLOW_KEYWORD_FOR) >= 0) {
        var domKwHit = Object.keys(domByBc).some(function(k){
          return domByBc[k].name && isItemMatchGlobal(inv.name, domByBc[k].name);
        });
        var impKwHit = false;
        if (window._schedMap) {
          var smKeys2 = Object.keys(window._schedMap);
          for (var sk = 0; sk < smKeys2.length; sk++) {
            var sm2 = window._schedMap[smKeys2[sk]];
            if (sm2.items && isItemMatchGlobal(inv.name, sm2.items)) { impKwHit = true; break; }
          }
        }
        orphanLabel = (domKwHit || impKwHit) ? '바코드없음(입고내용있음)' : '바코드없음';
      } else {
        // 스피루리나가 아니면 키워드 검색 자체를 하지 않고 무조건 바코드없음
        orphanLabel = '바코드없음';
      }
    } else if (bcMissingInDom && bcMissingInImp) {
      // 상품명조차 없는 경우(인벤토리에도 없음)
      orphanLabel = '바코드없음';
    }
    var hasValidVendor = false;
    if (imp.done.filter(function(p){ return p.done; }).length > 0) hasValidVendor = true;
    if (imp.pend.filter(function(p){ 
       var d1 = typeof p.ship === 'string' ? p.ship.trim() : p.ship;
       var d2 = typeof p.arrive === 'string' ? p.arrive.trim() : p.arrive;
       var d3 = typeof p.est === 'string' ? p.est.trim() : p.est;
       var d4 = typeof p.date === 'string' ? p.date.trim() : p.date;
       var d5 = typeof p.inbound === 'string' ? p.inbound.trim() : p.inbound;
       return d1 || d2 || d3 || d4 || d5;
    }).length > 0) hasValidVendor = true;
    
    // 바코드 매칭이 아니더라도(orphanLabel이 있어도) 품목명 키워드가 마스터시트와 일치하면
    // 실데이터(선적/입항/입고)는 그대로 채운다. 단, J열 출처표시에서는 "바코드 매칭 아님"을 구분 표시한다.
    var kwFilled = false;
    var domConfirmed = (dom.done.length > 0 || dom.pend.length > 0); // 국내쪽에 실제 바코드매칭 데이터가 이미 있는지
    if (!hasValidVendor && !domConfirmed && window._schedMap && inv.name) {
       var smKeys = Object.keys(window._schedMap);
       for (var k = 0; k < smKeys.length; k++) {
         var sm = window._schedMap[smKeys[k]];
         if (sm.items && isItemMatchGlobal(inv.name, sm.items)) {
            var entry = {
              bl: sm.bl, orderNo: sm.orderNo, vendor: sm.vendor, ship: sm.ship, arrive: sm.arrive,
              inbound: sm.inbound, est: sm.est, estCalc: sm.estCalc, ecountNote: sm.ecountNote, qty: null, items: sm.items,
              memo: sm.memo, matched: true, done: !!sm.inbound, sheet: '수입스케줄_키워드매칭'
            };
            if (sm.inbound) imp.done.push(entry); else imp.pend.push(entry);
            kwFilled = true;
         }
       }
    }
    // 키워드 매칭으로 실데이터가 채워졌다면, '확인필요'/'바코드없음(입고내용있음)' 대신
    // "바코드 매칭은 아니지만 수입 데이터는 채워짐"을 의미하는 라벨로 격하한다.
    if (kwFilled && orphanLabel === '바코드없음(입고내용있음)') {
      orphanLabel = '수입(바코드미확인)';
    }
    
    var sout;
    if (window._stockoutFileAttached) {
      sout = stockout[bc] || '';
    } else {
      sout = stockout[bc] || '';
    }

    if (!Array.isArray(imp.done)) imp.done = [];
    if (!Array.isArray(imp.pend)) imp.pend = [];
    if (!Array.isArray(dom.done)) dom.done = [];
    if (!Array.isArray(dom.pend)) dom.pend = [];

    var name  = inv.name || '';
    if (!name || /^[\d\s]+$/.test(name.trim())) {
      name = imp.name || dom.name || name;
    }
    
    var stock = (inv.stock != null) ? inv.stock : '';
    var isDisc = disc.has(bc);

    // ── 국내 확정일(done): 오늘과 가장 가까운 최근 입고 ──
    var domDoneSorted = dom.done.slice()
                                .sort(function(a,b){ return b.date - a.date; });
    var domRecentD    = domDoneSorted[0] ? domDoneSorted[0].date : null;
    var domRecentQ    = domDoneSorted[0] ? domDoneSorted[0].qty  : '';
    var domRecentMemo = domDoneSorted[0] ? (domDoneSorted[0].memo || '') : '';

    // ── 국내 대기일(pend): 가장 빠른 예정일 → 다음 입고 예정 ──
    var domPendSorted = dom.pend.sort(function(a,b){ return a.date - b.date; });
    var domNextD = domPendSorted[0] ? domPendSorted[0].date : null;
    var domNextQ = domPendSorted[0] ? domPendSorted[0].qty  : '';
    var domNextMemo = domPendSorted[0] ? (domPendSorted[0].memo || '') : '';
    var domNextHdrText = domPendSorted[0] ? (domPendSorted[0].headerText || '') : '';
    var hDomNextD = dom.petoneShareDate || domNextD;
    var hDomNextMemo = dom.petoneShareDate ? (dom.petoneReason || '') : domNextMemo;
    var domSheet = dom.sheet || '';
    var domNoteDisplay = dom.petoneNote || '';

    // 국내다음계획일 표시 문자열 생성
    var domNextDisplay = '';
    if (domNextD) {
      if (domNextMemo) {
        // 메모 있으면: "05/22(5-3)건 메모 5/29 입고예정"
        domNextDisplay = domNextHdrText + '건 메모 ' + domNextMemo + ' 입고예정';
      } else {
        domNextDisplay = fd(domNextD);
      }
    }

    var impDone = imp.done.slice().sort(function(a,b){
      var ad = a.inbound||a.arrive||a.ship;
      var bd = b.inbound||b.arrive||b.ship;
      var at = (ad instanceof Date) ? ad.getTime() : 0;
      var bt = (bd instanceof Date) ? bd.getTime() : 0;
      return bt - at; // 가장 최근 입고 완료된 건 우선 (내림차순)
    });
    // 예정: 입고예상일 기준으로 가장 먼저 올 예정인 건 1건
    // ※ 벤더시트에는 수년 전(예: 2023년)의 미완료/미정리 "예정" 열이 그대로 남아있는 경우가 있어,
    //   날짜 비교만으로 가장 빠른 걸 고르면 이미 지나가버린 옛날 날짜가 "다음입고예정"으로 잘못 표시된다.
    //   → 날짜가 명시된 건은 입고예상일이 오늘(TODAY) 이후인 것만 후보로 삼는다.
    //   선적일이 지났어도 선적+14일 입고예상일이 미래면 아직 유효한 예정건이다.
    function bestPendDate(p) {
      // est가 날짜로 계산/기재돼 있으면 그 값으로 스테일 여부 판단
      if (p.est instanceof Date) return p.est;
      // est가 "입고가능" 같은 텍스트 상태값이면 날짜로 판단할 수 없으니
      // 스테일(과거) 필터 대상에서 빼고 그냥 표시되게 둔다.
      // (입항일이 이미 지났어도, 입고일이 아직 미확정이면 "지난 건"이 아니라 "아직 미입고"인 것임)
      if (p.est) return null;
      if (p.inbound instanceof Date) return p.inbound;
      if (p.arrive instanceof Date) return new Date(p.arrive.getTime() + 7*86400000);
      if (p.ship instanceof Date) return new Date(p.ship.getTime() + 14*86400000);
      return null;
    }
    var todayMid = new Date(TODAY.getFullYear(), TODAY.getMonth(), TODAY.getDate());
    var impPend = imp.pend.filter(function(p){ return p.ship || p.arrive || p.est || p.date || p.inbound; })
      .filter(function(p){
        var bd = bestPendDate(p);
        return !bd || bd >= todayMid; // 날짜가 있는 경우 과거(스테일)면 제외
      })
      .sort(function(a,b){
        var ad = bestPendDate(a);
        var bd = bestPendDate(b);
        var at = (ad instanceof Date) ? ad.getTime() : Infinity;
        var bt = (bd instanceof Date) ? bd.getTime() : Infinity;
        return at - bt; // 입고예상일이 가장 빠른 예정건 우선 (오름차순)
      });

    var il = impDone[0] || null;
    var iv = impPend[0] || null;

    var ilShip  = il && il.ship   ? fdOrText(il.ship)   : '';
    var ilArr   = il && il.arrive ? fdOrText(il.arrive) : '';
    var ilInbd  = il && il.inbound ? fd(il.inbound) : '';
    var ilQty   = il ? (il.xlsQty != null ? il.xlsQty : il.qty) || '' : '';
    var ilVdor  = il ? (il.vendor  || '') : '';
    var ilItems = il ? String(il.items||'') : '';

    var ivShipD = iv && iv.ship   ? iv.ship   : null;
    var ivArrD  = iv && iv.arrive ? iv.arrive : null;
    var ivShip  = ivShipD ? fdOrText(ivShipD) : '';
    var ivArr   = ivArrD  ? fdOrText(ivArrD)  : '';
    var ivEstFirst = '';
    var ivEstTag = ''; // 계산으로 채운 추정일 때만 '(예정)' 또는 '(예정/통관전)' 등
    function buildEstTag(note) { return '(예정' + (note ? '/' + note : '') + ')'; }
    if (iv) {
      if (iv.est) {
        ivEstFirst = fdOrText(iv.est);
        if (iv.estCalc) ivEstTag = buildEstTag(iv.ecountNote); // 입항일+7일로 계산된 추정치
        // iv.est가 날짜가 아니라 "입고가능" 같은 실제 기재 텍스트면, 추정이 아니므로 태그를 붙이지 않는다
      } else if (ivArrD instanceof Date) {
        ivEstFirst = fd(new Date(ivArrD.getTime() + 7*86400000));
        ivEstTag = buildEstTag(iv.ecountNote); // 입항일+7일 추정 (마스터시트 미매칭 케이스)
      } else if (typeof iv.arrive === 'string' && iv.arrive.trim()) {
        ivEstFirst = '확인필요(' + iv.arrive.trim() + ')';
      } else if (ivShipD instanceof Date) {
        ivEstFirst = fd(new Date(ivShipD.getTime() + 14*86400000));
        ivEstTag = buildEstTag(iv.ecountNote); // 선적일+14일 추정
      } else if (typeof iv.ship === 'string' && iv.ship.trim()) {
        ivEstFirst = '확인필요(' + iv.ship.trim() + ')';
      }

      if (ivEstFirst && ivEstTag) ivEstFirst += ivEstTag;

      if (iv && iv.sheet === '수입스케줄_키워드매칭' && iv.items) {
        ivEstFirst = (ivEstFirst || '') + ' (품목: ' + String(iv.items) + ')';
      }
      if (iv && iv.memo) {
        ivEstFirst = (ivEstFirst || '') + ' (메모: ' + iv.memo + ')';
      }
    }
    var ivQty   = iv ? (iv.qty   || '') : '';
    var ivBl    = iv ? (iv.bl    || '') : '';
    var ivVdor  = iv ? (iv.vendor || '') : '';
    var ivItems = iv ? String(iv.items||'') : '';

    var cands = [];
    if (hDomNextD) {
      var domNextStr = hDomNextMemo
        ? fmtD(hDomNextD) + '(메모: ' + hDomNextMemo + ')'
        : fmtD(hDomNextD);
      cands.push({type:'국내', date:hDomNextD, str:domNextStr});
    }
    if (iv) {
      var impEstD = null;
      if (iv.est) impEstD = iv.est;
      else if (ivArrD instanceof Date) impEstD = new Date(ivArrD.getTime() + 7*86400000);
      
      var isKwMatch = (iv.sheet === '수입스케줄_키워드매칭');
      var kwStr = '';
      if (isKwMatch) {
         var kwArr = [];
         for(var i=0; i<impPend.length; i++) {
             var pt = impPend[i];
             if (pt.sheet === '수입스케줄_키워드매칭') {
                 var sStr = (pt.ship instanceof Date) ? fd(pt.ship) : (pt.ship || '미정');
                 var aStr = (pt.arrive instanceof Date) ? fd(pt.arrive) : (pt.arrive || '미정');
                 kwArr.push('선적:' + sStr + ' 입항:' + aStr + ' / ' + (pt.vendor||'') + ' / ' + (pt.orderNo||'') + '/' + (pt.items||''));
             }
         }
         kwStr = kwArr.join('\n');
      }

      if (impEstD) {
        cands.push({type:'수입', date: impEstD, str: isKwMatch ? kwStr : fd(impEstD)});
      } else if (ivEstFirst || isKwMatch) {
        cands.push({type:'수입', date: new Date('2099-12-31'), str: isKwMatch ? kwStr : ivEstFirst});
      }
    }
    cands.sort(function(a,b){ return (a.date||0)-(b.date||0); });
    var nextInbd = cands.length ? '['+cands[0].type+'] '+cands[0].str : '';

    var recentCands = [];
    if (domRecentD) {
      recentCands.push({type:'국내', date:domRecentD, str:fd(domRecentD)});
    }
    if (il && il.inbound instanceof Date) {
      if (il.sheet === '수입스케줄_키워드매칭' && il.ship && il.arrive && il.inbound) {
        var ilKwStr = fd(il.inbound) + ' / ' + (il.vendor||'') + ' / ' + (il.orderNo||'');
        recentCands.push({type:'수입', date:il.inbound, str: ilKwStr});
      } else {
        recentCands.push({type:'수입', date:il.inbound, str:fd(il.inbound)});
      }
    } else if (ilInbd) {
      // fallback string
      if (il.sheet === '수입스케줄_키워드매칭') {
         var kwArr = [];
         for(var i=0; i<impDone.length; i++) {
             var pt = impDone[i];
             if (pt.sheet === '수입스케줄_키워드매칭') {
                 var sStr = (pt.ship instanceof Date) ? fd(pt.ship) : (pt.ship || '미정');
                 var aStr = (pt.arrive instanceof Date) ? fd(pt.arrive) : (pt.arrive || '미정');
                 var iStr = (pt.inbound instanceof Date) ? fd(pt.inbound) : (pt.inbound || '미정');
                 kwArr.push('선적:' + sStr + ' 입항:' + aStr + ' 입고:' + iStr + ' / ' + (pt.vendor||'') + ' / ' + (pt.orderNo||'') + '/' + (pt.items||''));
             }
         }
         var ilKwStr = kwArr.join('\n');
         recentCands.push({type:'수입', date:new Date('1900-01-01'), str: ilKwStr});
      } else {
        recentCands.push({type:'수입', date:new Date('1900-01-01'), str:ilInbd});
      }
    }
    recentCands.sort(function(a,b){ return (b.date||0)-(a.date||0); });
    var recentInbd = recentCands.length ? '['+recentCands[0].type+'] '+recentCands[0].str : '';
    if (!hasValidVendor) {
        recentInbd = recentInbd ? '수입파일바코드없음\n' + recentInbd : '수입파일바코드없음';
    }

    var noteVal = '';
    if (il && il.xlsQty != null && il.qty != null) {
      var diff = il.xlsQty - il.qty;
      if (diff !== 0) {
        noteVal = '오차: ' + (diff > 0 ? '+' + diff : diff);
      }
    }

    // 적용상태 결정
    var statusVal = '';
    if (sout) {
      statusVal = '품절';
    } else if (stock !== '' && Number(stock) >= 0 && Number(stock) < 10) {
      statusVal = '품절임박';
    }

    // J열: 국내/수입 출처 표시
    var srcParts = [];
    if (domSheet || domRecentD || domNextD) srcParts.push('국내');
    if (ilInbd || ivEstFirst) srcParts.push('수입');
    var srcLabel = srcParts.join('+');
    // 국내/수입 양쪽 모두 "진짜 바코드 매칭"(키워드 추정 아님)으로 데이터가 있는 경우는
    // 원래 있을 수 없는 케이스(품목은 국내/수입 둘 중 하나로 고정)이므로, 조용히
    // 국내+수입으로 합치지 않고 데이터 충돌을 바로 알 수 있게 표시한다.
    var impConfirmed = imp.done.concat(imp.pend).some(function(e){ return e.sheet !== '수입스케줄_키워드매칭'; });
    if (domConfirmed && impConfirmed) {
      srcLabel = '국내+수입(중복확인필요)';
    } else if (orphanLabel) {
      // 바코드가 국내/수입 양쪽 파일에 전부 없는 경우 → 확인필요 / 바코드없음(입고내용있음)으로 덮어씀
      // (키워드매칭으로 G/H에 추정값이 채워졌어도, 바코드 매칭이 아니므로 J는 항상 이 라벨을 우선)
      srcLabel = orphanLabel;
    }

    var upd = updatedAt;

    rows.push([
      '',                      // A  (1)  단종 — 수동관리, 건드리지 않음
      bc,                      // B  (2)  바코드 (키)
      '',                      // C  (3)  상품명 — 수동관리, 건드리지 않음
      '',                      // D  (4)  수동관리 — 건드리지 않음
      stock,                   // E  (5)  현재고수량
      sout,                    // F  (6)  품절일자
      recentInbd,              // G  (7)  최근입고 (자동계산)
      nextInbd,                // H  (8)  다음입고 (자동계산)
      upd,                     // I  (9)  업데이트일시
      srcLabel,                // J  (10) 국내/수입 출처 표시 (자동계산)
      domNoteDisplay,          // K  (11) 국내스케줄 EA/EB/ED 참고
      domSheet,                // L  (12) 국내업체
      fd(domRecentD),          // M  (13) 국내최근입고일
      domRecentQ ? String(domRecentQ) + (domRecentMemo ? ' (' + domRecentMemo + ')' : '') : '',  // N  (14) 최근수량
      domNextDisplay,          // O  (15) 국내다음계획일
      domNextQ ? String(domNextQ) : '',      // P  (16) 국내다음수량
      ilVdor,                  // Q  (17) 수입업체
      ilShip,                  // R  (18) 선적일
      ilArr,                   // S  (19) 입항일
      ilInbd,                  // T  (20) 입고일
      ilQty ? String(ilQty) : '',   // U  (21) 완료수량
      ilItems,                 // V  (22) 완료참고품목
      ivVdor,                  // W  (23) 수입예정업체
      ivShip,                  // X  (24) 예정선적일
      ivArr,                   // Y  (25) 예정입항일
      ivEstFirst,              // Z  (26) 입고예상일
      ivQty ? String(ivQty) : '',   // AA (27) 예정수량
      ivBl,                    // AB (28) BL
      ivItems,                 // AC (29) 예정참고품목
    ]);
  }
  return rows;
}

async function runAnalysis() {
  $id('anaBtn').disabled = true;
  $id('anaPanel').classList.remove('show');
  $id('resultBox').className = 'result';
  var t0 = Date.now();
  try {
    var impByBc = {}, domByBc = {}, allItems = {}, discSet = new Set(), stockout = {};
    var xlsFiles = [];
    // ★ 파일 첨부 플래그 초기화 — 첨부된 파일의 열만 업데이트
    window._inventoryFileAttached = false;
    window._domesticFileAttached = false;
    window._importFileAttached = false;
    window._stockoutFileAttached = false;

    var fileEntries = Array.from(files.entries());
    for (var _fi = 0; _fi < fileEntries.length; _fi++) {
      var name = fileEntries[_fi][0], file = fileEntries[_fi][1];
      var normName = name.normalize('NFC');
      if (normName.includes('수입스케줄') && normName.endsWith('.xlsx')) {
        impByBc = await parseImport(file);
        window._importFileAttached = true;
      }
      else if (normName.includes('국내상품스케줄') && normName.endsWith('.xlsx')) {
        domByBc = await parseDomestic(file);
        window._domesticFileAttached = true;
      }
      else if ((normName.includes('품목확인용') || normName.includes('현재고')) && (normName.endsWith('.xlsx')||normName.endsWith('.xls'))) {
        var invResult = await parseInventory(file);
        allItems = invResult.items; discSet = invResult.disc;
        if (invResult.dupes && invResult.dupes.length > 0) {
          console.warn('중복 바코드:', invResult.dupes);
          window._dupeBarcodes = invResult.dupes;
        }
        window._inventoryFileAttached = true;
      }
      else if (normName.includes('품절') && (normName.endsWith('.xlsx')||normName.endsWith('.xls'))) {
        stockout = await parseStockout(file);
        window._stockoutFileAttached = true;
      }
      else if (normName.includes('입고수량'))                                 xlsFiles.push(file);
    }

    // XLS 매칭 분석
    var xlsResults = [];
    for (var fi2 = 0; fi2 < xlsFiles.length; fi2++) { var file = xlsFiles[fi2];
      var entries = await parseInboundXls(file);
      if (!entries.length) continue;
      var matched = 0, unmatched = 0, unmatchedList = [];
      for (var ei = 0; ei < entries.length; ei++) { var e = entries[ei];
        var imp = impByBc[e.bc];
        if (imp) {
          var blMatch = imp.entries.find(function(x) { return x.bl === e.bl; });
          var qMatch  = !blMatch && imp.entries.find(function(x) { return !x.done && Math.abs(x.qty - e.qty) <= e.qty * 0.05; });
          if (blMatch || qMatch) matched++;
          else { unmatched++; unmatchedList.push({bc:e.bc, name:e.name, qty:e.qty}); }
        } else { unmatched++; unmatchedList.push({bc:e.bc, name:e.name, qty:e.qty}); }
      }
      var m = file.name.match(/^(\d{8})_(.+?)_([\d-]+)__(.+?)_입고수량/);
      xlsResults.push({
        filename: file.name,
        bl:     m ? m[4] : '?', vendor: m ? m[2] : '?',
        date:   m ? (m[1].slice(0,4)+'-'+m[1].slice(4,6)+'-'+m[1].slice(6,8)) : '?',
        total: entries.length, matched, unmatched, unmatchedList
      });
      applyInboundXls(impByBc, entries);
    }

    var updatedAt = nowStr();
    var rows = buildRows(impByBc, domByBc, allItems, discSet, stockout, updatedAt);

    // 적용상태 계산 (이전 반영 데이터와 비교)
    var prevDataRaw = localStorage.getItem('prevRowData');
    var prevData = prevDataRaw ? JSON.parse(prevDataRaw) : {};
    rows.forEach(function(r) {
      var bc = String(r[1]), prev = prevData[bc];
      if (!prev) {
        r[8] = '신규';
      } else if (
        String(prev[3])  !== String(r[3])  ||  // 현재고
        String(prev[4])  !== String(r[4])  ||  // 품절
        String(prev[5])  !== String(r[5])  ||  // 다음입고일
        String(prev[17]) !== String(r[17]) ||  // 수입완료 입고일
        String(prev[21]) !== String(r[21]) ||  // 수입예정 입항일
        String(prev[23]) !== String(r[23])     // 수입예정 수량
      ) {
        r[8] = '변경';
      } else {
        r[8] = '유지';
      }
    });

    var live   = rows.filter(function(r) { return !(r[0]===true||r[0]==='✓'); });
    var disc   = rows.filter(function(r) { return r[0]===true||r[0]==='✓'; });
    var sout   = live.filter(function(r) { return r[4]; });
    var minus  = live.filter(function(r) { return Number(r[3])<0; });
    var zero   = live.filter(function(r) { return Number(r[3])===0; });
    var urgent = live.filter(function(r) { return Number(r[3])<=0||r[4]; });
    var impPend= live.filter(function(r) { return r[20]&&r[23]; });
    var domPend= live.filter(function(r) { return r[11]; });
    var impVendors = new Set(Object.values(impByBc).map(function(v){return v.vendor;}));
    var impBls = new Set();
    Object.values(impByBc).forEach(function(v){ var all=(v.done||[]).concat(v.pend||[]); all.forEach(function(e){if(e&&e.bl)impBls.add(e.bl);}); });

    _analysisResult = { rows, updatedAt, stats: {
      total:rows.length, live:live.length, disc:disc.length,
      sout:sout.length, minus:minus.length, zero:zero.length, urgent:urgent.length,
      impBc:Object.keys(impByBc).length, impVendors:impVendors.size, impBls:impBls.size,
      domBc:Object.keys(domByBc).length, impPend:impPend.length, domPend:domPend.length,
      petoneMemoCount: window._petoneMemoCount || 0
    }, xlsResults };

    renderAnalysis(_analysisResult, Date.now()-t0);
  } catch(err) {
    showResult(false, '분석 오류: ' + err.message);
    console.error(err);
  } finally {
    $id('anaBtn').disabled = false;
  }
}

function renderAnalysis(result, elapsed) {
  var s = result.stats, xl = result.xlsResults;
  $id('anaTime').textContent = '(' + (elapsed/1000).toFixed(1) + '초)';
  var grid = $id('anaGrid');
  grid.innerHTML = '';

  var cards = [
    { label:'전체 품목', stats:[['전체',s.total,''],['활성',s.live,'grn'],['단종',s.disc,'']] },
    { label:'재고 현황', stats:[['긴급관리',s.urgent,s.urgent>0?'red':'grn'],['재고마이너스',s.minus,s.minus>0?'red':''],['재고0',s.zero,s.zero>0?'yel':''],['품절신고',s.sout,s.sout>0?'yel':'']] },
    { label:'수입 현황', stats:[['인식바코드',s.impBc,'blu'],['업체수',s.impVendors,''],['컨테이너',s.impBls,''],['입고대기',s.impPend,s.impPend>0?'pur':'']] },
    { label:'국내 현황', stats:[['인식바코드',s.domBc,'grn'],['입고대기',s.domPend,s.domPend>0?'grn':''],['펫원메모',s.petoneMemoCount || 0,(s.petoneMemoCount||0)>0?'yel':'']] },
  ];
  cards.forEach(function(card) {
    var el = document.createElement('div');
    el.className = 'ana-card';
    el.innerHTML = '<div class="ana-card-label">' + card.label + '</div>' +
      card.stats.map(function(st) {
        return '<div class="ana-stat"><span class="ana-stat-k">' + st[0] + '</span><span class="ana-stat-v ' + st[2] + '">' + st[1] + '</span></div>';
      }).join('');
    grid.appendChild(el);
  });

  var xlsSec = $id('xlsSection'), xlsList = $id('xlsList');
  if (xl.length > 0) {
    xlsSec.style.display = 'block';
    xlsList.innerHTML = xl.map(function(r) {
      return '<div class="xls-item">' +
        '<div class="xls-item-top"><span class="xls-bl">' + r.bl + '</span><span class="xls-vendor">' + r.vendor + '</span><span class="xls-date">' + r.date + '</span></div>' +
        '<div class="xls-stats"><div class="xls-stat ok">총 <span>' + r.total + '</span>품목</div><div class="xls-stat ok">매칭 <span>' + r.matched + '</span>건 ✓</div>' +
        (r.unmatched > 0 ? '<div class="xls-stat warn">미매칭 <span>' + r.unmatched + '</span>건</div>' : '') + '</div>' +
        (r.unmatchedList.length > 0 ? '<div class="unmatched">' + r.unmatchedList.slice(0,10).map(function(u){return '<div class="unmatched-item"><span class="unmatched-bc">'+u.bc+'</span><span>'+u.name.slice(0,30)+'</span></div>';}).join('') + '</div>' : '') +
        '</div>';
    }).join('');
  } else { xlsSec.style.display = 'none'; }

  var hint = $id('runHint');
  var hasUrl = !!localStorage.getItem('scriptUrl');
  if (!hasUrl) {
    hint.innerHTML = '<strong style="color:var(--yel)">⚠ Apps Script URL 미설정</strong> — 설정 탭에서 입력 후 저장';
    $id('runBtn').textContent = '⚙ 설정 먼저 입력하세요';
    $id('runBtn').style.opacity = '0.5';
  } else {
    hint.textContent = '총 ' + result.rows.length + '개 품목 → 입고현황_전체 업데이트';
    hint.style.color = 'var(--grn)';
    $id('runBtn').textContent = '구글 시트에 반영 (' + result.rows.length + '건)';
    $id('runBtn').style.opacity = '1';
  }
  $id('anaPanel').classList.add('show');
}

// ══════════════════════════════════════════════
// 구글 시트 반영
// ══════════════════════════════════════════════
var PROG_STEPS = ['데이터 준비 중...','구글 시트 전송 중...','반영 완료 확인 중...'];
var EXPECTED_GAS_VERSION = '2026-06-26-main-k-preserve-v6';

async function checkGasVersion(url) {
  var r = await fetch(url + '?test=1&ts=' + Date.now());
  var d = await r.json();
  if (!d.ok) return {ok:false, error:d.error || 'Apps Script 응답 오류'};
  if (d.version !== EXPECTED_GAS_VERSION) {
    return {
      ok:false,
      error:'Apps Script 코드가 구버전입니다. 가이드 탭의 코드를 다시 복사해서 Apps Script에 붙여넣고 새 배포를 해야 K열 보존이 반영됩니다. 현재 버전: ' + (d.version || '버전 없음')
    };
  }
  return {ok:true, data:d};
}

async function testConnection() {
  var url = document.getElementById('scriptUrl').value.trim();
  var res = document.getElementById('connResult');
  if (!url) { res.textContent = '❌ URL 없음'; res.style.color='var(--red)'; return; }
  res.textContent = '확인 중...'; res.style.color = 'var(--txt3)';
  try {
    var v = await checkGasVersion(url);
    if (v.ok) { res.textContent = '✅ 연결 성공 / 최신 코드'; res.style.color = 'var(--grn)'; }
    else      { res.textContent = '❌ ' + v.error; res.style.color = 'var(--red)'; }
  } catch(err) {
    res.textContent = '❌ 연결 실패: ' + err.message;
    res.style.color = 'var(--red)';
  }
}


async function runUpdate() {
  var scriptUrl = localStorage.getItem('scriptUrl');
  var sheetId   = localStorage.getItem('sheetId') || '1SLkS88rKFESYVcxMxLqz67YNAp6VTK0Icm9lTH5AYXU';
  if (!scriptUrl) { switchTab('settings'); alert('Apps Script URL을 먼저 입력해주세요'); return; }
  if (!scriptUrl.includes('script.google.com')) { alert('Apps Script URL 형식이 잘못됐습니다'); return; }
  if (!_analysisResult) { alert('먼저 분석하기를 실행해주세요'); return; }

  $id('runBtn').disabled = true;
  $id('progWrap').classList.add('show');
  $id('progList').innerHTML = '';
  $id('resultBox').className = 'result';

  PROG_STEPS.forEach(function(s, i) {
    var d = document.createElement('div');
    d.className = 'psi'; d.id = 'ps' + i;
    d.innerHTML = '<div class="psic">○</div><span>' + s + '</span>';
    $id('progList').appendChild(d);
  });

  try {
    setProg(0);
    var gasCheck = await checkGasVersion(scriptUrl);
    if (!gasCheck.ok) {
      switchTab('guide');
      throw new Error(gasCheck.error);
    }
    var rows = _analysisResult.rows, updatedAt = _analysisResult.updatedAt;
    var summary = {
      total: rows.length,
      urgent: rows.filter(function(r){return !(r[0]===true||r[0]==='✓')&&(Number(r[3])<=0||r[4]);}).length
    };
    setProg(1);
    var res, resData;
    try {
      res = await fetch(scriptUrl, {
        method:'POST',
        headers:{'Content-Type':'text/plain'},
        body: JSON.stringify({sheetId:sheetId, rows:rows, updatedAt:updatedAt, summary:summary, stockoutFileAttached: !!window._stockoutFileAttached, inventoryFileAttached: !!window._inventoryFileAttached, domesticFileAttached: !!window._domesticFileAttached, importFileAttached: !!window._importFileAttached})
      });
      resData = await res.json();
    } catch(fe) {
      // CORS 오류시 no-cors 재시도 (응답 확인 불가하지만 전송은 됨)
      try {
        await fetch(scriptUrl, {
          method:'POST', mode:'no-cors',
          headers:{'Content-Type':'text/plain'},
          body: JSON.stringify({sheetId:sheetId, rows:rows, updatedAt:updatedAt, summary:summary, stockoutFileAttached: !!window._stockoutFileAttached, inventoryFileAttached: !!window._inventoryFileAttached, domesticFileAttached: !!window._domesticFileAttached, importFileAttached: !!window._importFileAttached})
        });
        resData = {ok:true, fallback:true};
      } catch(fe2) { throw new Error('전송 실패: ' + fe2.message); }
    }
    setProg(2);
    await new Promise(function(r){setTimeout(r,800);});
    setProg(3);

    if (resData && resData.ok === false) {
      throw new Error('GAS 오류: ' + (resData.error || '알 수 없는 오류'));
    }

    var saveData = {};
    rows.forEach(function(r){if(r[1])saveData[r[1]]=r;});
    try { localStorage.setItem('prevRowData', JSON.stringify(saveData)); } catch(e){}
    var msg;
    if (resData && resData.fallback) {
      msg = '전송 완료 (응답확인 불가) - 구글 시트에서 직접 확인하세요';
    } else if (resData && resData.ok) {
      msg = '반영 완료! 업데이트 ' + (resData.count||rows.length) + '건'
          + (resData.added ? ' · 신규 ' + resData.added + '건' : '')
          + ' (' + updatedAt + ')';
    } else {
      msg = '전송 완료 (' + updatedAt + ')';
    }
    showResult(true, msg, sheetId);
  } catch(err) {
    showResult(false, '오류: ' + err.message);
  } finally {
    $id('runBtn').disabled = false;
  }
}

function setProg(idx) {
  $id('progFill').style.width = Math.round(idx/PROG_STEPS.length*100) + '%';
  PROG_STEPS.forEach(function(_, i) {
    var el = $id('ps'+i); if (!el) return;
    el.className = 'psi' + (i<idx?' ok':i===idx?' on':'');
    el.querySelector('.psic').innerHTML = i<idx ? '✓' : i===idx ? '<span class="spin">◌</span>' : '○';
  });
}

function showResult(ok, msg, sheetId) {
  var box = $id('resultBox');
  box.className = 'result show ' + (ok ? 'ok' : 'err');
  $id('resIco').textContent = ok ? '✅' : '❌';
  $id('resTxt').textContent = msg;
  var lnk = $id('sheetLink');
  if (ok && sheetId) { lnk.href = 'https://docs.google.com/spreadsheets/d/' + sheetId; lnk.style.display = 'inline-flex'; }
  else lnk.style.display = 'none';
  $id('progWrap').classList.remove('show');
}

function copyScript() {
  var code = document.getElementById('scriptCode').textContent;
  navigator.clipboard.writeText(code).then(function() {
    var b = document.querySelector('[onclick="copyScript()"]');
    b.textContent = '✓ 복사됨';
    setTimeout(function(){ b.textContent = '코드 복사'; }, 2000);
  });
}














