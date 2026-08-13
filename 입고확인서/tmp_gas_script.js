// ═══════════════════════════════════════════
// 입고현황 자동화 Apps Script v6
// 최종 수정: 2026-06-26 / 입고현황_전체 K열 보존
// ═══════════════════════════════════════════
// 시트: 입고현황_전체 / 단종
// 구조: 29열 (A~AC)
// E=현재고, F=품절일자, G=최근입고(자동), H=다음입고(자동)
// I=업데이트일시, J=국내/수입(자동), K=사용자 함수 보존
// A,C,D,K는 절대 덮어쓰지 않음 (상품정보/함수 보호)
// 신규 바코드 행 추가 없음 — 기존 바코드만 업데이트
// ═══════════════════════════════════════════

var SHEET_MAIN = '입고현황_전체';
var SHEET_DISC = '단종';
var APP_VERSION = '2026-06-26-main-k-preserve-v6';
var NCOLS = 29;

// 열 구성: A~AC (29열)
// A,C,D는 상품정보/수동관리 — 절대 덮어쓰지 않음
var HDR = [
  '단종','바코드','상품명','',          // A  B  C  D  ← 상품정보 (건드리지 않음)
  '현재고','품절일자',                   // E  F
  '최근입고','다음입고','업데이트일시','국내/수입','', // G  H  I  J  K(사용자 함수 보존)
  '국내업체','국내최근입고일','최근수량','국내다음계획일','국내다음수량', // L  M  N  O  P
  '수입업체','선적일','입항일','입고일','완료수량','완료참고품목',        // Q  R  S  T  U  V
  '수입예정업체','예정선적일','예정입항일','입고예상일','예정수량','BL','예정참고품목' // W  X  Y  Z  AA AB AC
];

var C = {
  DISC:1, BC:2, NAME:3,
  STOCK:5, SOUT:6,
  REC:7, NEXT:8,    // G=최근입고, H=다음입고 (자동계산)
  UPD:9,
  SRC:10,           // J=국내/수입 표시 (자동계산)
  NOTE:11,          // K=사용자 함수/메모 보존
  DOM_VDR:12, DOM_LD:13, DOM_LQ:14, DOM_ND:15, DOM_NQ:16,
  IMP_VDR:17, SHIP:18, ARR:19, INBD:20, IMP_QTY:21, IMP_REF:22,
  IV_VDR:23, IV_SHIP:24, IV_ARR:25, IV_EST:26, IV_QTY:27, IV_BL:28, IV_REF:29
};

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = data.sheetId ? SpreadsheetApp.openById(data.sheetId) : SpreadsheetApp.getActiveSpreadsheet();
    var rows = data.rows || [];
    var updatedAt = data.updatedAt || '';
    var stockoutFileAttached = data.stockoutFileAttached || false;
    var inventoryFileAttached = data.inventoryFileAttached || false;
    var domesticFileAttached = data.domesticFileAttached || false;
    var importFileAttached = data.importFileAttached || false;

    // 하위 호환성 필터
    if (data.inventoryFileAttached === undefined && data.domesticFileAttached === undefined && data.importFileAttached === undefined) {
      inventoryFileAttached = true;
      domesticFileAttached = true;
      importFileAttached = true;
    }

    if (rows.length === 0) {
      return respond({ok:false, error:'데이터 없음 - 파일 선택 후 분석하기를 먼저 실행하세요'});
    }

    var ws = getOrCreate(ss, SHEET_MAIN);
    ensureHeaders(ws);

    var result = updateMainSheet(ss, ws, rows, updatedAt, stockoutFileAttached, inventoryFileAttached, domesticFileAttached, importFileAttached);
    syncDiscSheet(ss, ws);
    setupTrigger(ss);

    return respond({ok:true, version:APP_VERSION, updated:result.updated, count:result.updated, added:result.added, updatedAt:updatedAt});
  } catch(err) {
    return respond({ok:false, error:err.toString()});
  }
}

function doGet(e) {
  return respond({ok:true, version:APP_VERSION, message:'연결 정상 - Apps Script 웹앱이 작동 중입니다'});
}

// ── 날짜 파싱 및 계산 유틸 ─────────────────
function parseDateFromString(str) {
  if (!str) return null;
  if (str instanceof Date) return str;
  str = String(str).trim();
  if (!str) return null;

  // YYYY-MM-DD
  var m1 = str.match(/(\d{4})[-.](\d{2})[-.](\d{2})/);
  if (m1) return new Date(m1[1], m1[2]-1, m1[3]);

  // YY.MM.DD
  var m2 = str.match(/(\d{2})[-.](\d{2})[-.](\d{2})/);
  if (m2) {
    var year = parseInt(m2[1]) + 2000;
    return new Date(year, m2[2]-1, m2[3]);
  }

  // MM/DD
  var m3 = str.match(/(\d{1,2})\/(\d{1,2})/);
  if (m3) {
    var year = new Date().getFullYear();
    return new Date(year, m3[1]-1, m3[2]);
  }

  return null;
}

function formatDateToYYYYMMDD(date) {
  if (!date) return '';
  var y = date.getFullYear();
  var m = String(date.getMonth() + 1).padStart(2, '0');
  var d = String(date.getDate()).padStart(2, '0');
  return y + '-' + m + '-' + d;
}

function formatNextDom(originalStr, date) {
  var yy = String(date.getFullYear()).slice(2);
  var mm = String(date.getMonth() + 1).padStart(2, '0');
  var dd = String(date.getDate()).padStart(2, '0');
  var base = yy + '.' + mm + '.' + dd;
  if (originalStr.indexOf('메모') >= 0) {
    var memoMatch = originalStr.match(/메모[:\s]*(.*?)\s*입고예정/);
    var memoText = memoMatch && memoMatch[1] ? memoMatch[1].trim() : '';
    return memoText ? base + '(메모: ' + memoText + ')' : base + '(메모)';
  }
  return base;
}

function petoneNextFromNote(row) {
  var note = String(row[C.NOTE-1] || '');
  if (!note || note.indexOf('펫원 공유 입고예정일') < 0) return null;
  var dateMatch = note.match(/펫원 공유\s*입고예정일:\s*([^/]+)/);
  var reasonMatch = note.match(/지연사유\s*및\s*특이사항:\s*(.*)$/);
  var d = dateMatch ? parseDateFromString(dateMatch[1].trim()) : null;
  if (!d) return null;
  var reason = reasonMatch && reasonMatch[1] ? reasonMatch[1].trim() : '';
  return {
    date: d,
    text: formatNextDom(reason ? ('메모: ' + reason + ' 입고예정') : '', d)
  };
}

// ── G열: 최근입고 (국내최근입고일 vs 수입입고일 비교) ──────
function computeRecent(row) {
  var domStr = row[C.DOM_LD-1];   // M(13) 국내최근입고일
  var impStr = row[C.INBD-1];     // T(20) 수입입고일
  var domD = parseDateFromString(domStr);
  var impD = parseDateFromString(impStr);
  if (domD && impD) {
    return domD.getTime() >= impD.getTime()
      ? '[국내] ' + formatDateToYYYYMMDD(domD)
      : '[수입] ' + formatDateToYYYYMMDD(impD);
  }
  if (domD) return '[국내] ' + formatDateToYYYYMMDD(domD);
  if (impD) return '[수입] ' + formatDateToYYYYMMDD(impD);
  return '';
}

// ── H열: 다음입고 (국내다음계획일 vs 수입입고예상일 비교) ──
function computeNext(row) {
  var domRaw = row[C.DOM_ND-1];   // O(15) 국내다음계획일
  var impRaw = row[C.IV_EST-1];   // Z(26) 수입입고예상일
  // GAS는 날짜 셀을 Date 객체로 반환 → indexOf 호출 전 문자열 변환
  var domStr = (domRaw instanceof Date) ? '' : String(domRaw || '');
  var impStr = (impRaw instanceof Date) ? '' : String(impRaw || '');
  var domD = parseDateFromString(domRaw);
  var impD = parseDateFromString(impRaw);
  var petoneNext = petoneNextFromNote(row);
  if (petoneNext) {
    domD = petoneNext.date;
    domStr = petoneNext.text;
  }
  // "(예정)" / "(예정/통관전)" 같이 입항일+7일 등으로 계산한 추정치 표시는
  // 날짜를 재포맷해도 그대로 보존
  var impSuffixMatch = impStr.match(/\(예정[^)]*\)/);
  var impSuffix = impSuffixMatch ? impSuffixMatch[0] : '';
  function formatImp() {
    return (impStr.indexOf('확인필요') >= 0) ? impStr : (formatDateToYYYYMMDD(impD) + impSuffix);
  }
  if (domD && impD) {
    return domD.getTime() <= impD.getTime()
      ? '[국내] ' + (petoneNext ? petoneNext.text : formatNextDom(domStr, domD))
      : '[수입] ' + formatImp();
  }
  if (domD) return '[국내] ' + (petoneNext ? petoneNext.text : formatNextDom(domStr, domD));
  if (impD) return '[수입] ' + formatImp();
  if (impStr && impStr.indexOf('확인필요') >= 0) return '[수입] ' + impStr;
  return '';
}

// ── J열: 국내/수입 출처 표시 ──────────────────────────────
function computeSrc(row, incomingSrc) {
  // 프런트에서 보낸 특수 상태 라벨은 항상 그대로 유지한다.
  // (이 라벨들은 hasDom/hasImp 컬럼이 채워진 상태에서도 의미가 있으므로,
  //  아래 일반 로직이 무조건 덮어쓰면 안 된다 — 예: 키워드매칭으로 데이터는
  //  채워졌지만 바코드 확인은 안 된 '수입(바코드미확인)' 같은 경우)
  if (incomingSrc === '바코드없음' ||
      incomingSrc === '바코드없음(입고내용있음)' ||
      incomingSrc === '수입(바코드미확인)' ||
      incomingSrc === '국내+수입(중복확인필요)') {
    return incomingSrc;
  }
  var hasDom = !!(row[C.DOM_LD-1] || row[C.DOM_ND-1]);
  var hasImp = !!(row[C.INBD-1]   || row[C.IV_EST-1]);
  if (hasDom && hasImp) return '국내+수입';
  if (hasDom) return '국내';
  if (hasImp) return '수입';
  return '';
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ── 매 실행 시 헤더 보장 ─────────────────────
// A,C,D는 상품정보/수동관리, K는 사용자 함수 — 건드리지 않음.
function ensureHeaders(ws) {
  if (ws.getLastRow() < 2) {
    initHeaders(ws);
  } else {
    ws.getRange(3,5,1,2).setValues([['현재고','품절일자']]);                             // E,F
    ws.getRange(3,7,1,4).setValues([['최근입고','다음입고','업데이트일시','국내/수입']]); // G,H,I,J
    ws.getRange(3,12,1,18).setValues([[
      '국내업체','국내최근입고일','최근수량','국내다음계획일','국내다음수량',
      '수입업체','선적일','입항일','입고일','완료수량','완료참고품목',
      '수입예정업체','예정선적일','예정입항일','입고예상일','예정수량','BL','예정참고품목'
    ]]);
  }
}

function getManualDiscMap(ss) {
  var map = {};
  try {
    var ws = ss.getSheetByName(SHEET_DISC);
    if (!ws || ws.getLastRow() < 3) return map;
    ws.getRange(3,1,ws.getLastRow()-2,1).getValues()
      .forEach(function(r){ if(r[0]) map[String(r[0])]=true; });
  } catch(ex) {}
  return map;
}

// 신규 바코드는 추가하지 않음 — 시트에 이미 있는 바코드만 업데이트
function updateMainSheet(ss, ws, rows, updatedAt, stockoutFileAttached, inventoryFileAttached, domesticFileAttached, importFileAttached) {
  var lastRow = Math.max(3, ws.getLastRow());
  var existingData = (lastRow >= 4) ? ws.getRange(4, 1, lastRow - 3, NCOLS).getValues() : [];
  var bcToRow = {};

  for (var i = 0; i < existingData.length; i++) {
    var bc = String(existingData[i][C.BC-1]||'').trim();
    if (bc) bcToRow[bc] = i + 4;
  }

  // 품절파일 첨부 시 전체 품절일자 열 초기화 (새 데이터로 덮어쓰기)
  if (stockoutFileAttached && existingData.length > 0) {
    for (var i = 0; i < existingData.length; i++) {
      existingData[i][C.SOUT-1] = '';
    }
  }

  var updatedCount = 0, newRowsBuffer = [];

  rows.forEach(function(r) {
    var bc = String(r[C.BC-1]||'').trim();
    if (!bc) return;

    // 시트에 없는 바코드는 건너뜀 (신규 행 추가 없음)
    if (!bcToRow[bc]) return;

    var ri = bcToRow[bc];
    var arrayIndex = ri - 4;
    var existingRow = existingData[arrayIndex];

    // E열: 현재고 (현재고파일 첨부 시만)
    if (inventoryFileAttached) {
      existingRow[C.STOCK-1] = (r[C.STOCK-1] !== null && r[C.STOCK-1] !== undefined) ? r[C.STOCK-1] : '';
    }

    // F열: 품절일자 (품절파일 첨부 시만)
    if (stockoutFileAttached) {
      existingRow[C.SOUT-1] = r[C.SOUT-1] || '';
    }

    // L~P열: 국내 (국내스케줄 첨부 시만)
    if (domesticFileAttached) {
      existingRow[C.DOM_VDR-1] = r[C.DOM_VDR-1];
      existingRow[C.DOM_LD-1]  = r[C.DOM_LD-1];
      existingRow[C.DOM_LQ-1]  = r[C.DOM_LQ-1];
      existingRow[C.DOM_ND-1]  = r[C.DOM_ND-1];
      existingRow[C.DOM_NQ-1]  = r[C.DOM_NQ-1];
    }

    // Q~AC열: 수입 (수입스케줄 첨부 시만)
    if (importFileAttached) {
      existingRow[C.IMP_VDR-1] = r[C.IMP_VDR-1];
      existingRow[C.SHIP-1]    = r[C.SHIP-1];
      existingRow[C.ARR-1]     = r[C.ARR-1];
      existingRow[C.INBD-1]    = r[C.INBD-1];
      existingRow[C.IMP_QTY-1] = r[C.IMP_QTY-1];
      existingRow[C.IMP_REF-1] = r[C.IMP_REF-1];
      existingRow[C.IV_VDR-1]  = r[C.IV_VDR-1];
      existingRow[C.IV_SHIP-1] = r[C.IV_SHIP-1];
      existingRow[C.IV_ARR-1]  = r[C.IV_ARR-1];
      existingRow[C.IV_EST-1]  = r[C.IV_EST-1];
      existingRow[C.IV_QTY-1]  = r[C.IV_QTY-1];
      existingRow[C.IV_BL-1]   = r[C.IV_BL-1];
      existingRow[C.IV_REF-1]  = r[C.IV_REF-1];
    }

    // G,H,J열: 기존 조건대로 재계산
    existingRow[C.REC-1]  = computeRecent(existingRow);
    existingRow[C.NEXT-1] = computeNext(existingRow);
    existingRow[C.SRC-1]  = computeSrc(existingRow, r[C.SRC-1]);

    // I열: 업데이트일시 (항상 갱신)
    existingRow[C.UPD-1] = r[C.UPD-1];
    // A(1),C(3),D(4)는 절대 수정하지 않음 — 읽어온 값 그대로 유지

    existingData[arrayIndex] = existingRow;
    newRowsBuffer.push({ri:ri, stock:Number(existingRow[C.STOCK-1])||0, sout:String(existingRow[C.SOUT-1]||'')});
    updatedCount++;
  });

  if (existingData.length > 0) {
    writeMainDataSkippingK(ws, existingData);
    ws.getRange(4, C.DISC, existingData.length, 1).insertCheckboxes();
  }

  applyStylesBatch(ws, newRowsBuffer, existingData.length);

  return {updated:updatedCount, added:0};
}

function writeMainDataSkippingK(ws, existingData) {
  if (!existingData.length) return;
  var left = existingData.map(function(r){ return r.slice(0, 10); });   // A:J
  var right = existingData.map(function(r){ return r.slice(11, NCOLS); }); // L:AC
  ws.getRange(4, 1, existingData.length, 10).setValues(left);
  ws.getRange(4, 12, existingData.length, NCOLS - 11).setValues(right);
}

function applyStylesBatch(ws, buffer, numRows) {
  if (numRows <= 0) return;

  var bgArray = [], fcArray = [], flArray = [], fsArray = [], fwArray = [];
  for (var i = 0; i < numRows; i++) {
    var bg = ((i + 4) % 2 === 0) ? '#F8F9FA' : '#FFFFFF';
    bgArray.push(new Array(NCOLS).fill(bg));
    fcArray.push(new Array(NCOLS).fill('#111827'));
    flArray.push(new Array(NCOLS).fill('none'));
    fsArray.push(new Array(NCOLS).fill(9));
    fwArray.push(new Array(NCOLS).fill('normal'));
  }

  buffer.forEach(function(item) {
    var idx = item.ri - 4;
    if (idx < 0 || idx >= numRows) return;

    var hasSout = !!item.sout;
    var stockLow = !hasSout && item.stock >= 0 && item.stock < 10;

    if (item.stock < 0) {
      bgArray[idx].fill('#FFEBEE');
      fcArray[idx][C.STOCK-1] = '#C62828';
      fwArray[idx][C.STOCK-1] = 'bold';
    } else if (hasSout) {
      bgArray[idx].fill('#FFEBEE');
      fcArray[idx][C.STOCK-1] = '#C62828';
      fwArray[idx][C.STOCK-1] = 'bold';
    } else if (item.stock === 0) {
      bgArray[idx].fill('#FFFDE7');
      fcArray[idx][C.STOCK-1] = '#F57F17';
      fwArray[idx][C.STOCK-1] = 'bold';
    } else if (stockLow) {
      bgArray[idx].fill('#FFF8E1');
      fcArray[idx][C.STOCK-1] = '#E65100';
      fwArray[idx][C.STOCK-1] = 'bold';
    }

    fcArray[idx][C.UPD-1] = '#9E9E9E';
    fsArray[idx][C.UPD-1] = 8;
  });

  applyRowStylesSkippingK(ws, bgArray, fcArray, flArray, fsArray, fwArray, numRows);
}

function applyRowStylesSkippingK(ws, bgArray, fcArray, flArray, fsArray, fwArray, numRows) {
  var groups = [
    {start:1, cols:10},
    {start:12, cols:NCOLS - 11}
  ];
  groups.forEach(function(g) {
    function pick(arr) { return arr.map(function(r){ return r.slice(g.start - 1, g.start - 1 + g.cols); }); }
    var r = ws.getRange(4, g.start, numRows, g.cols);
    r.setBackgrounds(pick(bgArray));
    r.setFontColors(pick(fcArray));
    r.setFontLines(pick(flArray));
    r.setFontSizes(pick(fsArray));
    r.setFontWeights(pick(fwArray));
  });
}

function initHeaders(ws) {
  ws.getRange(1,1).setValue('입고현황 관리표  |  자동반영: E,F,G,H,I,J,L~AC열  |  A,C,D,K는 수동/함수관리')
    .setBackground('#1a2235').setFontColor('#7EB3FF')
    .setFontSize(10).setFontWeight('bold').setVerticalAlignment('middle');
  ws.setRowHeight(1, 26);

  [[1,10,'기본 정보','#1B2A4A','#7EB3FF'],
   [12,16,'국내 입고 (파일 반영)','#1B3A28','#86EFAC'],
   [17,22,'수입 완료 (파일 반영)','#172B45','#93C5FD'],
   [23,29,'수입 예정 (파일 반영)','#0F2340','#60A5FA']
  ].forEach(function(s) {
    ws.getRange(2,s[0]).setValue(s[2]).setBackground(s[3]).setFontColor(s[4])
      .setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center');
  });
  ws.setRowHeight(2, 15);

  ws.getRange(3,1,1,10).setValues([HDR.slice(0,10)])
    .setBackground('#243450').setFontColor('#CBD5E1')
    .setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center');
  ws.getRange(3,12,1,NCOLS - 11).setValues([HDR.slice(11)])
    .setBackground('#243450').setFontColor('#CBD5E1')
    .setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center');
  ws.setRowHeight(3, 15);

  // 29열 너비: A~AC (K열은 사용자 함수 열이라 건드리지 않음)
  [5,13,34,10,10,10,13,13,12,8,null,  11,11,8,11,8,  11,10,10,10,8,22,  11,10,10,14,8,20,22]
    .forEach(function(w,i) { if (w) ws.setColumnWidth(i+1, w*7); });

  try { ws.setFrozenRows(3); } catch(e) {}
  try { ws.setFrozenColumns(3); } catch(e) {}
}

function syncDiscSheet(ss, mainWs) {
  var lastRow = mainWs.getLastRow();
  if (lastRow < 4) return;

  var allData = mainWs.getRange(4,1,lastRow-3,NCOLS).getValues();
  var discRows = allData.filter(function(r){ return r[C.DISC-1]===true; });

  var ws = getOrCreate(ss, SHEET_DISC);
  
  // 3행 이하 데이터 범위만 내용과 서식을 지웁니다. (1~2행 병합 유지)
  var lastR = ws.getLastRow();
  if (lastR >= 3) {
    ws.getRange(3, 1, lastR - 2, 5).clearContent().clearFormat();
  }

  // 1행과 2행의 값만 업데이트합니다. (서식은 건드리지 않음)
  ws.getRange(1,1).setValue('단종 품목  |  총 ' + discRows.length + '건');
  ws.getRange(2,1,1,5).setValues([['바코드','상품명','현재고','업데이트일시','']]);

  if (discRows.length > 0) {
    var data = discRows.map(function(r) {
      return [r[C.BC-1], r[C.NAME-1], r[C.STOCK-1], r[C.UPD-1], ''];
    });
    ws.getRange(3,1,data.length,5).setValues(data);
    data.forEach(function(_,i) {
      var ri = i + 3;
      ws.getRange(ri,1,1,5).setBackground(i%2===0?'#ECEFF1':'#FFFFFF').setFontSize(9);
      ws.getRange(ri,2).setFontLine('line-through').setFontColor('#757575');
      ws.setRowHeight(ri, 16);
    });
  }
  [14,30,8,12,10].forEach(function(w,i){ ws.setColumnWidth(i+1, w*7); });
  try { ws.setFrozenRows(2); } catch(e) {}
}

function onEdit(e) {
  try {
    var ws = e.range.getSheet();
    if (ws.getName() !== SHEET_MAIN) return;
    if (e.range.getColumn() !== 1) return;
    var ri = e.range.getRow();
    if (ri < 4) return;

    var checked = (e.range.getValue() === true);
    var ss = ws.getParent();
    var row = ws.getRange(ri,1,1,NCOLS).getValues()[0];
    var bc = String(row[C.BC-1]);
    var nm = String(row[C.NAME-1]);
    var st = row[C.STOCK-1];

    if (checked) {
      styleSingleRowSkippingK(ws, ri, '#EEEEEE', '#AAAAAA', 'line-through');
      ws.getRange(ri,C.DISC).setBackground('#E0E0E0');
      addToDisc(ss, bc, nm, st);
    } else {
      var bg = (ri%2===0) ? '#F8F9FA' : '#FFFFFF';
      styleSingleRowSkippingK(ws, ri, bg, '#111827', 'none');
      removeFromDisc(ss, bc);
    }
  } catch(err) {}
}

function styleSingleRowSkippingK(ws, row, bg, color, line) {
  [[1,10], [12, NCOLS - 11]].forEach(function(g) {
    ws.getRange(row, g[0], 1, g[1]).setBackground(bg).setFontColor(color).setFontLine(line);
  });
}

function addToDisc(ss, bc, nm, st) {
  var ws = ss.getSheetByName(SHEET_DISC);
  if (!ws) ws = ss.insertSheet(SHEET_DISC);
  var last = ws.getLastRow();
  if (last >= 3) {
    var ex = ws.getRange(3,1,last-2,1).getValues().map(function(r){return String(r[0]);});
    if (ex.indexOf(bc) >= 0) return;
  }
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yy.MM.dd HH:mm');
  var nr = Math.max(last,2) + 1;
  ws.getRange(nr,1,1,5).setValues([[bc,nm,st,now,'']]);
  ws.getRange(nr,1,1,5).setBackground('#ECEFF1').setFontSize(9);
  ws.getRange(nr,2).setFontLine('line-through').setFontColor('#757575');
}

function removeFromDisc(ss, bc) {
  var ws = ss.getSheetByName(SHEET_DISC);
  if (!ws || ws.getLastRow() < 3) return;
  var vals = ws.getRange(3,1,ws.getLastRow()-2,1).getValues();
  for (var i = vals.length-1; i >= 0; i--) {
    if (String(vals[i][0]) === bc) ws.deleteRow(i+3);
  }
}

function setupTrigger(ss) {
  try {
    var has = false;
    ScriptApp.getProjectTriggers().forEach(function(t){
      if (t.getHandlerFunction() === 'onEdit') has = true;
    });
    if (!has) ScriptApp.newTrigger('onEdit').forSpreadsheet(ss).onEdit().create();
  } catch(ex) {}
}

function getOrCreate(ss, name) {
  var ws = ss.getSheetByName(name);
  if (!ws) ws = ss.insertSheet(name);
  return ws;
}


