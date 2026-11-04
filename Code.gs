// 코인 배분 기록 서버 — Google Apps Script (구글 시트에 제출을 쌓습니다)
// 배포 순서는 07_코인배분앱_운영안내.md 2장을 따릅니다. 코드는 고칠 것이 없습니다.
const SHEET_NAME = '제출';
const ITEMS = ['A1','A2','B1','B2','B3','B4','B5','C1','C2','C3','D1','D2','D3','E1','E2','E3','E4','E5','F1','F2','F3','F4','F5','F6','G1','G2','G3'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['시각', '팀', '브랜드', '라운드', '합계', '이동코인'].concat(ITEMS).concat(['배분(JSON)']));
    sh.setFrozenRows(1);
  }
  return sh;
}

// 앱의 '제출' 버튼이 부릅니다. 한 줄씩 쌓이므로 팀별 이력이 전부 남습니다.
function doPost(e) {
  const p = JSON.parse(e.postData.contents || '{}');
  const alloc = p.alloc || {};
  const total = ITEMS.reduce((s, k) => s + (Number(alloc[k]) || 0), 0);
  sheet_().appendRow([new Date(), Number(p.team) || 0, String(p.brand || ''), Number(p.round) || 1, total, Number(p.moved) || 0]
    .concat(ITEMS.map(k => Number(alloc[k]) || 0)).concat([JSON.stringify(alloc)]));
  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}

// 진행자 보기가 부릅니다. ?round=1 또는 2. ?round=all 이면 전부.
function doGet(e) {
  const want = (e && e.parameter && e.parameter.round) || '1';
  const values = sheet_().getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (want !== 'all' && Number(r[3]) !== Number(want)) continue;
    const alloc = {};
    ITEMS.forEach((k, j) => alloc[k] = Number(r[6 + j]) || 0);
    rows.push({ ts: r[0] instanceof Date ? r[0].toISOString() : String(r[0]), team: Number(r[1]), brand: String(r[2] || ''),
                round: Number(r[3]), total: Number(r[4]) || 0, moved: Number(r[5]) || 0, alloc: alloc });
  }
  return ContentService.createTextOutput(JSON.stringify(rows)).setMimeType(ContentService.MimeType.JSON);
}

// 편집기에서 한 번 실행해 두면 시트 탭이 미리 만들어집니다(권한 승인도 이때 끝납니다).
function setup() { sheet_(); }
