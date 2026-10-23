/**
 * PENTA SERVICE → Google 캘린더 자동 동기화 (Google Apps Script)
 *
 * 앱은 보낼 일정 목록을 Firestore(gcal/items)에 올려두기만 한다.
 * 이 스크립트가 Google 서버에서 10분마다 그 목록을 읽어 캘린더에 그대로 맞춘다.
 * 앱이 꺼져 있어도 돌고, 본인 Google 계정 권한으로 쓰기 때문에 팝업·1시간 만료가 없다.
 *
 * 처음 한 번만: 왼쪽 "서비스 +" 에서 Google Calendar API 를 추가하고(식별자 Calendar),
 * 위쪽 함수 선택에서 install 을 골라 실행 → 권한 허용.
 */
var PROJECT = 'pentaservice-3de4c';
var API_KEY = 'AIzaSyBK-VAiq9b5ugWFxjefQ4-t-zgkj6IJgfk';
var BASE = 'https://firestore.googleapis.com/v1/projects/' + PROJECT + '/databases/(default)/documents';

// Google 캘린더는 짧은 시간에 많이 쓰면 "Rate Limit Exceeded" 로 막는다(처음 수백 건 정리할 때 실제로 막힘).
var GAP_MS = 300;                    // 쓰기 한 번마다 쉬는 시간
var MAX_RUN_MS = 4.5 * 60 * 1000;    // 한 번 실행 한도(6분)보다 일찍 멈추고 나머지는 다음 실행에 넘긴다

/** 10분마다 syncNow 를 돌리는 트리거를 건다(여러 번 실행해도 하나만 남는다). */
function install() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncNow') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncNow').timeBased().everyMinutes(10).create();
  syncNow();
}

/** 자동 동기화를 끈다. */
function uninstall() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncNow') ScriptApp.deleteTrigger(t);
  });
}

function syncNow() {
  var started = new Date();
  var deadline = started.getTime() + MAX_RUN_MS;
  var pushed = 0, removed = 0, total = 0, remaining = 0, failed = 0, err = '';
  try {
    var doc = fetchDoc_('gcal/items');
    if (!doc) throw new Error('앱이 아직 일정 목록을 올리지 않았습니다 — 앱을 한 번 열어주세요');
    var calId = str_(doc.fields.calendarId);
    if (!calId) throw new Error('보낼 캘린더가 지정되지 않았습니다 — 앱 설정에서 캘린더를 고르세요');
    var desired = JSON.parse(str_(doc.fields.json) || '[]');
    total = desired.length;
    var existing = listOurs_(calId);
    // 안전장치: 목록이 통째로 비어 오면(앱이 덜 읽힌 상태 등) 캘린더를 지우지 않는다
    if (desired.length === 0 && Object.keys(existing).length > 20) {
      throw new Error('보낼 목록이 비어 있어 삭제를 멈췄습니다(안전장치)');
    }
    var want = {};
    desired.forEach(function (b) {
      want[b.id] = true;
      var ex = existing[b.id];
      if (ex && hashOf_(ex) === hashOf_(b)) return;   // 내용이 같으면 건너뛴다
      if (Date.now() > deadline) { remaining++; return; }
      try {
        withRetry_(function () { upsert_(calId, b, !!ex); });
        pushed++;
      } catch (e) {
        failed++;
        if (!err) err = String((e && e.message) || e);
      }
      Utilities.sleep(GAP_MS);
    });
    Object.keys(existing).forEach(function (id) {
      if (want[id]) return;
      if (Date.now() > deadline) { remaining++; return; }
      try {
        withRetry_(function () { Calendar.Events.remove(calId, id); });
        removed++;
      } catch (e) { /* 이미 지워진 것 등 — 다음 실행에서 다시 본다 */ }
      Utilities.sleep(GAP_MS);
    });
    if (failed) err = failed + '건 실패 — ' + err + ' (다음 실행에서 다시 시도합니다)';
  } catch (e) {
    err = String((e && e.message) || e);
  }
  writeAgent_({
    lastRun: started.toISOString(), lastError: err,
    pushed: pushed, removed: removed, total: total, remaining: remaining
  });
  var summary = '전체 ' + total + '건, 올림 ' + pushed + '건, 삭제 ' + removed + '건'
    + (remaining ? ', 남은 ' + remaining + '건은 다음 실행에서 이어서' : '');
  if (err) console.error(err + ' / ' + summary); else console.log(summary);
}

// 한도에 걸리거나 Google 쪽이 잠깐 바쁠 때는 기다렸다가 다시 시도한다(1·2·4·8·16초)
function withRetry_(fn) {
  var wait = 1000;
  for (var i = 0; ; i++) {
    try { return fn(); }
    catch (e) {
      var msg = String((e && e.message) || e);
      if (i < 5 && /rate ?limit|quota|exceeded|backend|try again|429|500|503/i.test(msg)) {
        Utilities.sleep(wait);
        wait *= 2;
        continue;
      }
      throw e;
    }
  }
}

// 우리 앱이 보낸 이벤트만(app=pentaservice 표시가 붙은 것) 가져온다 — 다른 일정은 건드리지 않는다
function listOurs_(calId) {
  var out = {}, pageToken = null;
  do {
    var opt = { privateExtendedProperty: 'app=pentaservice', maxResults: 2500, showDeleted: false };
    if (pageToken) opt.pageToken = pageToken;
    var res = withRetry_(function () { return Calendar.Events.list(calId, opt); });
    (res.items || []).forEach(function (ev) { out[ev.id] = ev; });
    pageToken = res.nextPageToken;
  } while (pageToken);
  return out;
}

function hashOf_(ev) {
  var p = ev.extendedProperties && ev.extendedProperties['private'];
  return (p && p.h) || '';
}

function upsert_(calId, b, exists) {
  if (exists) { Calendar.Events.update(b, calId, b.id); return; }
  try {
    Calendar.Events.insert(b, calId);
  } catch (e) {
    var msg = String((e && e.message) || e);
    if (/rate ?limit|quota|exceeded/i.test(msg)) throw e;   // 한도 문제는 위에서 기다렸다 다시 한다
    // 예전에 지웠던 같은 id 가 '취소됨' 으로 남아 있으면 insert 가 막힌다 — 되살린다
    var r = JSON.parse(JSON.stringify(b));
    r.status = 'confirmed';
    Calendar.Events.update(r, calId, b.id);
  }
}

function fetchDoc_(path) {
  var res = UrlFetchApp.fetch(BASE + '/' + path + '?key=' + API_KEY, { muteHttpExceptions: true });
  var code = res.getResponseCode();
  if (code === 404) return null;
  if (code !== 200) throw new Error('Firestore 읽기 실패 (' + code + ')');
  return JSON.parse(res.getContentText());
}

function str_(v) { return v && v.stringValue != null ? v.stringValue : ''; }

// 앱 설정 화면에 "마지막 실행 시각·오류"를 보여주려고 Firestore(gcal/agent)에 남긴다
function writeAgent_(o) {
  var fields = {};
  Object.keys(o).forEach(function (k) {
    fields[k] = (typeof o[k] === 'number') ? { integerValue: String(o[k]) } : { stringValue: String(o[k]) };
  });
  UrlFetchApp.fetch(BASE + '/gcal/agent?key=' + API_KEY, {
    method: 'patch', contentType: 'application/json',
    payload: JSON.stringify({ fields: fields }), muteHttpExceptions: true
  });
}
