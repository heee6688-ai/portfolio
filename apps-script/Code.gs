/**
 * 포트폴리오 문의 폼 → 구글 시트 저장 + 알림 메일 + 자동회신
 *
 * 시트 헤더(1행) 순서: 이름 | 이메일 | 전화번호 | 회사,직함 | 강의 문의 내용
 * 배포: 확장 프로그램 > Apps Script > 배포 > 새 배포 > 유형 "웹 앱"
 *       실행 사용자: 나 / 액세스 권한: 모든 사용자
 */

var OWNER_EMAIL = 'heee6688@gmail.com';
var SHEET_NAME = '';              // 비워 두면 첫 번째 시트에 저장
var SENDER_NAME = '포트폴리오 문의';

var FIELDS = ['name', 'email', 'phone', 'company', 'message'];
var LIMITS = { name: 100, email: 200, phone: 50, company: 200, message: 5000 };

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var data = parseBody_(e);

    // 스팸 방지: 폼에 숨긴 입력칸(name="website")이 채워져 있으면 봇으로 보고 조용히 무시
    if (data.website) return json_({ ok: true });

    var error = validate_(data);
    if (error) return json_({ ok: false, error: error });

    var sheet = getSheet_();
    sheet.appendRow([data.name, data.email, data.phone, data.company, data.message]);

    sendOwnerNotice_(data);
    sendAutoReply_(data);

    return json_({ ok: true });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: '일시적인 오류가 발생했습니다. 잠시 후 다시 시도해 주세요.' });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

/** 브라우저 주소창으로 배포 주소를 열었을 때 동작 확인용 */
function doGet() {
  return json_({ ok: true, message: 'contact form endpoint' });
}

/* ---------- 입력 처리 ---------- */

function parseBody_(e) {
  var raw = {};
  var type = (e && e.postData && e.postData.type) || '';

  if (type.indexOf('application/json') === 0 || type.indexOf('text/plain') === 0) {
    try { raw = JSON.parse(e.postData.contents); } catch (ignore) { raw = {}; }
  } else if (e && e.parameter) {
    raw = e.parameter;
  }

  var out = { website: clean_(raw.website, 100) };
  FIELDS.forEach(function (key) {
    out[key] = clean_(raw[key], LIMITS[key]);
  });
  return out;
}

function clean_(value, max) {
  var s = value == null ? '' : String(value).trim();
  if (s.length > max) s = s.substring(0, max);
  // 구글 시트에서 수식으로 실행되지 않도록 앞의 = + - @ 를 막는다
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function validate_(d) {
  if (!d.name) return '이름을 입력해 주세요.';
  if (!/^[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]+$/.test(d.email)) return '이메일 형식을 확인해 주세요.';
  if (!d.message) return '문의 내용을 입력해 주세요.';
  return '';
}

/* ---------- 시트 ---------- */

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getSheets()[0];
  if (!sheet) throw new Error('시트를 찾을 수 없습니다: ' + SHEET_NAME);
  return sheet;
}

/* ---------- 메일 ---------- */

function sendOwnerNotice_(d) {
  var body =
    '새 문의가 도착했습니다.\n\n' +
    '이름: ' + d.name + '\n' +
    '이메일: ' + d.email + '\n' +
    '전화번호: ' + (d.phone || '-') + '\n' +
    '회사,직함: ' + (d.company || '-') + '\n\n' +
    '[강의 문의 내용]\n' + d.message + '\n';

  MailApp.sendEmail({
    to: OWNER_EMAIL,
    replyTo: d.email,
    name: SENDER_NAME,
    subject: '[새 문의 도착] ' + d.name + '님',
    body: body
  });
}

function sendAutoReply_(d) {
  var text =
    d.name + '님, 안녕하세요.\n\n' +
    '문의가 접수되었습니다. 남겨 주신 내용을 확인한 뒤 빠르게 회신드리겠습니다.\n\n' +
    '[접수 내용]\n' + d.message + '\n\n' +
    '감사합니다.';

  var html =
    '<div style="font-family:\'Malgun Gothic\',Arial,sans-serif;font-size:15px;line-height:1.7;color:#191f28">' +
    '<p>' + esc_(d.name) + '님, 안녕하세요.</p>' +
    '<p>문의가 접수되었습니다. 남겨 주신 내용을 확인한 뒤 빠르게 회신드리겠습니다.</p>' +
    '<p style="margin:20px 0 6px;color:#6b7684;font-size:13px">접수 내용</p>' +
    '<blockquote style="margin:0;padding:12px 16px;background:#f2f4f6;border-radius:8px;white-space:pre-wrap">' +
    esc_(d.message) + '</blockquote>' +
    '<p style="margin-top:24px">감사합니다.</p></div>';

  MailApp.sendEmail({
    to: d.email,
    replyTo: OWNER_EMAIL,
    name: SENDER_NAME,
    subject: '문의가 접수되었습니다',
    body: text,
    htmlBody: html
  });
}

function esc_(s) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/* ---------- 응답 ---------- */

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 에디터에서 직접 실행해 권한 승인과 동작을 테스트하는 함수 */
function testPost() {
  var res = doPost({
    postData: {
      type: 'application/json',
      contents: JSON.stringify({
        name: '테스트',
        email: OWNER_EMAIL,
        phone: '010-0000-0000',
        company: '테스트 회사 / 팀장',
        message: '강의 문의 테스트입니다.'
      })
    }
  });
  Logger.log(res.getContent());
}
