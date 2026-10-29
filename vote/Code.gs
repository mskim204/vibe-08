// 시트 탭 이름
var CODE_SHEET = '코드';
var VOTE_SHEET = '투표';

// 투표 보기
var CHOICES = ['떡볶이', '치킨', '마라탕', '붕어빵'];

// 웹앱 주소로 들어오면 화면을 보여줘요.
// 주소 끝에 ?page=result 를 붙이면 결과 화면이 나와요.
function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) === 'result' ? 'result' : 'index';
  return HtmlService.createHtmlOutputFromFile(page)
    .setTitle(page === 'result' ? '투표 결과' : '간식 투표')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 코드를 비교하기 좋게 다듬어요. (앞뒤 공백 제거, 대문자로)
function normalize(value) {
  return String(value).trim().toUpperCase();
}

// 투표하기: 결과 메시지만 돌려줘요.
function vote(code, choice) {
  code = normalize(code || '');
  if (!code) return { ok: false, message: '코드를 입력해 주세요.' };
  if (CHOICES.indexOf(choice) === -1) return { ok: false, message: '보기를 골라 주세요.' };

  // 여러 명이 동시에 투표해도 꼬이지 않게 잠가요.
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1) 코드 탭에 있는 코드인지 확인 (A열)
    var codeSheet = ss.getSheetByName(CODE_SHEET);
    var codes = codeSheet.getRange(1, 1, Math.max(codeSheet.getLastRow(), 1), 1).getValues();
    var exists = codes.some(function (row) { return normalize(row[0]) === code; });
    if (!exists) return { ok: false, message: '없는 코드에요' };

    // 2) 투표 탭에 이미 쓴 코드인지 확인 (B열)
    var voteSheet = ss.getSheetByName(VOTE_SHEET);
    if (voteSheet.getLastRow() === 0) voteSheet.appendRow(['시간', '코드', '선택']);
    var used = voteSheet.getRange(1, 2, voteSheet.getLastRow(), 1).getValues();
    var already = used.some(function (row) { return normalize(row[0]) === code; });
    if (already) return { ok: false, message: '이미 투표했어요' };

    // 3) 처음 쓰는 맞는 코드일 때만 저장
    voteSheet.appendRow([new Date(), code, choice]);
    return { ok: true, message: '투표 완료! ' + choice + '에 투표했어요 🎉' };
  } finally {
    lock.releaseLock();
  }
}

// 결과: 보기별 표 수만 돌려줘요. (코드는 보내지 않아요)
function getResults() {
  var voteSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VOTE_SHEET);
  var counts = {};
  CHOICES.forEach(function (c) { counts[c] = 0; });
  if (voteSheet.getLastRow() > 0) {
    voteSheet.getRange(1, 3, voteSheet.getLastRow(), 1).getValues().forEach(function (row) {
      if (counts.hasOwnProperty(row[0])) counts[row[0]]++;
    });
  }
  return CHOICES.map(function (c) { return { name: c, count: counts[c] }; });
}
