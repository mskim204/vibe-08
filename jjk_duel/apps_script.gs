/**
 * 주술결전 카드 듀얼 — 온라인 대전 서버 (Google Apps Script)
 *
 * [설치 방법]
 * 1) Google Sheets에서 새 스프레드시트를 만듭니다.
 * 2) 메뉴: 확장 프로그램 → Apps Script 를 엽니다.
 * 3) Code.gs 내용을 모두 지우고 이 코드를 붙여넣은 뒤 저장(💾)합니다.
 * 4) 오른쪽 위 [배포] → [새 배포] → 톱니바퀴에서 [웹 앱] 선택
 *      - 실행 계정: 나
 *      - 액세스 권한: 모든 사용자
 *    → [배포] → 권한 승인 → 발급된 "웹 앱 URL"(…/exec) 복사
 * 5) (추천) 같은 프로젝트에 HTML 파일을 추가합니다: 왼쪽 [파일 +] → [HTML] → 이름 index
 *    → 내용을 모두 지우고 게임 index.html 전체를 붙여넣고 저장 → [배포 관리]에서 새 버전 배포.
 *    이제 웹 앱 URL을 열면 게임이 바로 뜹니다 (휴대폰 가능, 서버 설정 입력 불필요).
 * 6) 또는 index.html 파일을 직접 열어 [온라인 대전] → [서버 설정]에 URL을 붙여넣어도 됩니다.
 */
const SHEET_NAME = 'Rooms';
const TURN_SEC = 60;        // 턴당 제한시간(초)
const ROOM_TTL_MIN = 30;    // 이 시간(분) 동안 갱신이 없는 방은 자동 삭제
const HEAD = ['code', 'name', 'host', 'hostName', 'guest', 'guestName', 'guestDeck',
  'status', 'turn', 'deadline', 'version', 'updated', 'left', 'state'];
const CODE_CHARS = 'BCDFGHJKLMNPQRSTVWXZ'; // 숫자/모음 제외 → 시트가 숫자·불리언으로 변환하지 않음

function doGet(e) {
  const prm = (e && e.parameter) || {};
  // 주소를 그냥 열면(파라미터 없음) 게임 화면을 보여줌 → 휴대폰·PC 어디서나 URL 하나로 플레이
  if (!prm.q && !prm.action && !prm.callback) {
    try {
      return HtmlService.createHtmlOutputFromFile('index')
        .setTitle('주술결전 카드 듀얼')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } catch (err) {
      return out_({ ok: true, message: '서버는 정상입니다. 게임 화면을 띄우려면 Apps Script에 HTML 파일 "index"를 추가하고 게임 index.html 내용을 붙여넣은 뒤 다시 배포하세요.' });
    }
  }
  // 호환 방식(JSONP): ?q=<JSON>&callback=<함수명> 형태로 요청하면 JavaScript로 응답
  let p = prm;
  if (prm.q) { try { p = JSON.parse(prm.q); } catch (err) { p = {}; } }
  const out = route_(p);
  if (prm.callback && /^[A-Za-z_$][\w$]{0,60}$/.test(prm.callback)) {
    return ContentService.createTextOutput(prm.callback + '(' + out.getContent() + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return out;
}
// 게임 화면(HtmlService) 안에서 google.script.run 으로 호출하는 통신 함수
function apiCall(json) {
  let p = {};
  try { p = JSON.parse(json); } catch (err) { p = {}; }
  return route_(p).getContent();
}
function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) { p = {}; }
  return route_(p);
}
// 큰 데이터를 호환 방식으로 보낼 때는 여러 조각(chunk)으로 나눠 받은 뒤 합쳐서 처리
function route_(p) {
  if (p && p.action === 'chunk') {
    const cache = CacheService.getScriptCache();
    const key = 'ck_' + String(p.id || '').slice(0, 40);
    const n = Number(p.n), i = Number(p.i);
    if (!(n > 0 && n <= 60 && i >= 0 && i < n)) return out_({ ok: false, error: 'bad chunk' });
    cache.put(key + '_' + i, String(p.data || ''), 600);
    if (i < n - 1) return out_({ ok: true, chunk: i });
    let s = '';
    for (let k = 0; k < n; k++) {
      const part = cache.get(key + '_' + k);
      if (part === null) return out_({ ok: false, error: 'chunk missing' });
      s += part;
    }
    let q = {};
    try { q = JSON.parse(s); } catch (err) { return out_({ ok: false, error: 'chunk parse' }); }
    return handle_(q);
  }
  return handle_(p);
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]);
    sh.setFrozenRows(1);
  }
  return sh;
}
function clean_(s, n) {
  return String(s == null ? '' : s).replace(/^[=+\-@']+/, '').slice(0, n || 20);
}
function genCode_() {
  let c = '';
  for (let i = 0; i < 5; i++) c += CODE_CHARS.charAt(Math.floor(Math.random() * CODE_CHARS.length));
  return c;
}

function handle_(p) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return out_({ ok: false, error: '서버가 혼잡합니다. 잠시 후 다시 시도하세요.' });
  try {
    const now = Date.now();
    const sh = sheet_();
    const values = sh.getDataRange().getValues();
    const rows = [];
    for (let i = 1; i < values.length; i++) {
      const r = { _row: i + 1 };
      HEAD.forEach((h, j) => { r[h] = values[i][j]; });
      rows.push(r);
    }
    const find = c => rows.find(r => String(r.code) === String(c || '').toUpperCase().trim());
    const save = r => sh.getRange(r._row, 1, 1, HEAD.length)
      .setValues([HEAD.map(h => (r[h] === undefined || r[h] === null) ? '' : r[h])]);
    const pid = String(p.pid || '');
    const member = r => !!pid && (String(r.host) === pid || String(r.guest) === pid);
    const res = o => { o.now = now; return out_(o); };

    switch (p.action) {
      case 'ping':
        return res({ ok: true });

      case 'info': // 게임 공유용: 이 웹 앱의 주소
        return res({ ok: true, url: ScriptApp.getService().getUrl() });

      case 'list': { // 방 목록 (+ 오래된 방 정리)
        const alive = [];
        for (let i = rows.length - 1; i >= 0; i--) {
          const r = rows[i];
          if (now - Number(r.updated || 0) > ROOM_TTL_MIN * 60000) { sh.deleteRow(r._row); continue; }
          if (r.status !== 'finished') {
            alive.unshift({ code: r.code, name: r.name, hostName: r.hostName, guestName: r.guestName, status: r.status });
          }
        }
        return res({ ok: true, rooms: alive });
      }

      case 'create': {
        if (!pid) return res({ ok: false, error: '플레이어 ID가 없습니다.' });
        let code;
        do { code = genCode_(); } while (find(code));
        const r = {
          code: code, name: clean_(p.roomName, 24) || '주술 대전', host: pid, hostName: clean_(p.name, 16),
          guest: '', guestName: '', guestDeck: '', status: 'waiting', turn: '', deadline: 0,
          version: 0, updated: now, left: '', state: ''
        };
        sh.appendRow(HEAD.map(h => r[h]));
        return res({ ok: true, code: code, role: 'host' });
      }

      case 'join': {
        const r = find(p.code);
        if (!r) return res({ ok: false, error: '방을 찾을 수 없습니다.' });
        if (String(r.host) === pid) return res({ ok: true, role: 'host', code: r.code });
        if (String(r.guest) === pid) return res({ ok: true, role: 'guest', code: r.code });
        // 난입 / 관전 방어: 대기 중인 방에만, 빈 자리일 때만 입장 가능
        if (r.status !== 'waiting' || r.guest) {
          return res({ ok: false, error: '이미 대전 중인 방입니다. (난입·관전 불가)' });
        }
        r.guest = pid;
        r.guestName = clean_(p.name, 16);
        r.guestDeck = String(p.deck || '[]').slice(0, 5000);
        r.status = 'ready';
        r.updated = now;
        save(r);
        return res({ ok: true, role: 'guest', code: r.code });
      }

      case 'room': { // 방/게임 상태 조회 (참가자만)
        const r = find(p.code);
        if (!r) return res({ ok: false, error: 'gone' });
        if (!member(r)) return res({ ok: false, error: '대전 참가자만 접근할 수 있습니다. (관전 불가)' });
        const o = {
          ok: true, code: r.code, status: r.status, host: String(r.host), guest: String(r.guest),
          hostName: r.hostName, guestName: r.guestName, version: Number(r.version || 0),
          turn: String(r.turn || ''), deadline: Number(r.deadline || 0), left: String(r.left || '')
        };
        if (String(r.host) === pid) o.guestDeck = r.guestDeck;
        if (o.version > Number(p.since || 0)) o.state = r.state;
        return res(o);
      }

      case 'push': { // 게임 상태 업로드 (턴 잠금 + 버전 검사)
        const r = find(p.code);
        if (!r) return res({ ok: false, error: 'gone' });
        if (!member(r)) return res({ ok: false, error: '대전 참가자가 아닙니다.' });
        if (r.status === 'finished') return res({ ok: false, error: 'finished' });
        const v = Number(r.version || 0);
        if (Number(p.version) !== v) return res({ ok: false, error: 'conflict', version: v });
        if (r.turn && String(r.turn) !== pid) {
          // Turn Lock: 상대 턴에는 조작 불가. 단, 제한시간이 지나면 자동 턴 넘김 허용
          const timeoutOk = (p.timeout == 1 || p.timeout === true) && now > Number(r.deadline || 0);
          if (!timeoutOk) return res({ ok: false, error: 'turnlock' });
        }
        r.state = String(p.state || '');
        r.version = v + 1;
        r.updated = now;
        if (p.nextTurn && String(p.nextTurn) !== String(r.turn)) {
          r.turn = String(p.nextTurn);
          r.deadline = now + TURN_SEC * 1000;
        }
        if (p.status) r.status = String(p.status);
        save(r);
        return res({ ok: true, version: r.version, deadline: Number(r.deadline), turn: r.turn });
      }

      case 'leave': {
        const r = find(p.code);
        if (!r) return res({ ok: true });
        if (!member(r)) return res({ ok: false, error: '참가자가 아닙니다.' });
        if (r.status === 'waiting') { sh.deleteRow(r._row); return res({ ok: true }); }
        if (r.status !== 'finished') { r.left = pid; r.status = 'finished'; }
        r.updated = now;
        save(r);
        return res({ ok: true });
      }

      default:
        if (!p.action) {
          // 브라우저에서 URL을 직접 열었을 때: 서버가 정상 동작 중임을 알려줌
          return res({ ok: true, message: '주술결전 서버가 정상 동작 중입니다. 이 URL을 게임의 [온라인 대전 → 서버 설정]에 붙여넣으세요.' });
        }
        return res({ ok: false, error: 'unknown action: ' + p.action });
    }
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
