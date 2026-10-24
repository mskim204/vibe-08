// ===== 설정 =====
// 이 스크립트는 구글 시트에 붙어 있어야 해요 (시트에서 확장 프로그램 → Apps Script)
const CHOICES = ['떡볶이', '치킨', '마라탕', '붕어빵'];
const CODE_SHEET = '코드';   // 1행: 코드 (제목), 2행부터 코드
const VOTE_SHEET = '투표';   // 1행: 시각 · 코드 · 고른 것 (제목)

// 주소 끝에 ?page=result 를 붙이면 결과 화면
function doGet(e) {
  const t = HtmlService.createTemplateFromFile('index');
  t.page = (e && e.parameter && e.parameter.page === 'result') ? 'result' : 'vote';
  return t.evaluate()
    .setTitle(t.page === 'result' ? '투표 결과' : '투표하기')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 화면에는 보기 목록만 보낸다 (코드 목록은 절대 보내지 않음)
function getChoices() {
  return CHOICES;
}

// 시트에 0123 이 숫자 123 으로 저장돼 있어도 같은 코드로 보이게 맞추기
function normalize(v) {
  const s = String(v).trim();
  return /^\d{1,4}$/.test(s) ? s.padStart(4, '0') : s;
}

// 2행부터 첫 번째 칸(또는 col번째 칸)의 값들을 가져오기
function columnValues(sheet, col) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, col, last - 1, 1).getValues().map(r => normalize(r[0]));
}

function vote(code, choice) {
  code = String(code || '').trim();
  if (!/^\d{4}$/.test(code)) return { ok: false, msg: '없는 코드예요' };
  if (CHOICES.indexOf(choice) < 0) return { ok: false, msg: '보기를 골라 주세요' };

  // 두 사람이 동시에 같은 코드로 눌러도 한 번만 저장되도록 잠그기
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const codes = columnValues(ss.getSheetByName(CODE_SHEET), 1);
    if (codes.indexOf(code) < 0) return { ok: false, msg: '없는 코드예요' };

    const voteSheet = ss.getSheetByName(VOTE_SHEET);
    const used = columnValues(voteSheet, 2);
    if (used.indexOf(code) >= 0) return { ok: false, msg: '이미 투표했어요' };

    // 코드는 글자로 저장 (앞의 0이 사라지지 않게 ' 를 붙임)
    voteSheet.appendRow([new Date(), "'" + code, choice]);
    return { ok: true, msg: '투표 완료! 고마워요 🎉' };
  } finally {
    lock.releaseLock();
  }
}

// 결과 화면: 보기별 표 수만 보낸다 (누가 무엇을 골랐는지는 보내지 않음)
function getResults() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VOTE_SHEET);
  const counts = {};
  CHOICES.forEach(c => counts[c] = 0);
  const last = sheet.getLastRow();
  if (last >= 2) {
    sheet.getRange(2, 3, last - 1, 1).getValues().forEach(r => {
      if (r[0] in counts) counts[r[0]]++;
    });
  }
  return CHOICES.map(c => ({ name: c, count: counts[c] }));
}
