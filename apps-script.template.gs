// 목원 출석부 — Apps Script (구글 시트 > 확장 프로그램 > Apps Script 에 붙여넣기)
//
// 1) 아래 ADMIN_PW 를 관리자 암호로 바꾸세요. (이 파일을 apps-script.gs 로 복사해서 고치면 git 에 안 올라가요)
// 2) 배포 > 배포 관리 > 연필(수정) > 버전: 새 버전 > 배포  (웹 앱 URL 은 그대로 유지돼요)
//    처음 배포라면: 배포 > 새 배포 > 웹 앱, 실행: 나, 액세스: 모든 사용자
//
// 권한
// - 관리자 암호로 들어오면 모든 기록을 보고 고칠 수 있고, 명단을 관리해요.
// - 개인 코드로 들어오면 본인 기록만 받고, 본인 기록만 고칠 수 있어요.
//   개인 코드는 명단을 저장할 때 새 목원에게 자동으로 만들어져요. ('목원' 시트 B열에서 직접 바꿔도 돼요)
// - 부부처럼 명단 한 줄에 같이 적은 사람들은 같은 코드로 들어와서 각자 기록을 따로 체크해요.

const ADMIN_PW = "여기에 관리자 암호";

const MEMBER_SHEET = "목원";   // 이름 | 코드
const RECORD_SHEET = "출석";   // 날짜 | 이름 | 예배 | 사유

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out({ ok: false, error: "bad" }); }

  const who = auth(body.pw);
  if (!who) {
    Utilities.sleep(1500);   // 코드 무작위 대입을 느리게
    return out({ ok: false, error: "pw" });
  }

  if (body.action === "load") return out(load(who));

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (body.action === "mark") {
      if (!who.admin && who.names.indexOf(body.name) < 0) return out({ ok: false, error: "forbidden" });
      mark(body.date, body.name, body.service, body.reason);
      return out({ ok: true });
    }
    if (body.action === "members") {
      if (!who.admin) return out({ ok: false, error: "forbidden" });
      PropertiesService.getScriptProperties().setProperty("title", body.title || "");
      const groups = body.groups || (body.members || []).map(n => [n]);
      return out({ ok: true, codes: saveMembers(groups) });
    }
  } finally {
    lock.releaseLock();
  }
  return out({ ok: false, error: "action" });
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 관리자면 { admin: true }, 목원이면 { admin: false, names: [같은 코드인 사람들] }, 아니면 null
function auth(pw) {
  pw = String(pw || "").trim();
  if (!pw) return null;
  if (pw === ADMIN_PW) return { admin: true };
  const names = memberRows().filter(r => r[1] && r[1] === pw).map(r => r[0]);
  return names.length ? { admin: false, names } : null;
}

function load(who) {
  const title = PropertiesService.getScriptProperties().getProperty("title") || "";
  const members = memberRows();
  const records = recordRows().filter(r => who.admin || who.names.indexOf(r[1]) >= 0);
  if (who.admin) {
    const codes = {};
    members.forEach(r => codes[r[0]] = r[1]);
    return { ok: true, role: "admin", title, members: members.map(r => r[0]), codes, records };
  }
  return { ok: true, role: "member", me: who.names.join(" · "), title, members: who.names, records };
}

function mark(date, name, service, reason) {
  const sh = sheet(RECORD_SHEET, ["날짜", "이름", "예배", "사유", "수정 시각"]);
  const rows = recordRows();
  const i = rows.findIndex(r => r[0] === date && r[1] === name);
  if (!service) {
    if (i >= 0) sh.deleteRow(i + 2);
    return;
  }
  const row = [[date, name, service, reason || "", new Date()]];
  const r = i >= 0 ? i + 2 : sh.getLastRow() + 1;
  if (i < 0) sh.getRange(r, 1, 1, 4).setNumberFormat("@");
  sh.getRange(r, 1, 1, 5).setValues(row);
  sh.getRange(r, 5).setNumberFormat("yyyy-mm-dd hh:mm");
}

// 명단을 다시 쓰면서 기존 코드는 유지하고, 새 가정에게 코드를 만들어 줌
// groups = [["홍길동", "김영희"], ["김철수"]] — 한 묶음(부부 등)은 같은 코드
function saveMembers(groups) {
  const sh = sheet(MEMBER_SHEET, ["이름", "코드"]);
  const old = {};
  memberRows().forEach(r => old[r[0]] = r[1]);
  const taken = new Set(Object.values(old).filter(Boolean));   // 새 코드는 예전 코드와도 안 겹치게
  const given = new Set();
  const codes = {}, rows = [];
  groups.forEach(g => {
    const names = g.map(n => String(n).trim()).filter(Boolean);
    if (!names.length) return;
    // 묶음 안 누군가의 예전 코드를 이어 씀 (다른 묶음이 이미 가져갔으면 새로 만듦)
    let c = names.map(n => old[n]).find(x => x && !given.has(x));
    if (!c) {
      do { c = String(1000 + Math.floor(Math.random() * 9000)); } while (taken.has(c) || given.has(c) || c === ADMIN_PW);
    }
    given.add(c);
    names.forEach(n => { codes[n] = c; rows.push([n, c]); });
  });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 2).setNumberFormat("@").setValues(rows);
  return codes;
}

// [[이름, 코드]]
function memberRows() {
  const sh = sheet(MEMBER_SHEET, ["이름", "코드"]);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
    .map(r => [String(r[0]).trim(), String(r[1]).trim()])
    .filter(r => r[0]);
}

// [[날짜, 이름, 예배, 사유]]
function recordRows() {
  const sh = sheet(RECORD_SHEET, ["날짜", "이름", "예배", "사유", "수정 시각"]);
  if (sh.getLastRow() < 2) return [];
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  return sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().map(r => [
    r[0] instanceof Date ? Utilities.formatDate(r[0], tz, "yyyy-MM-dd") : String(r[0]).trim(),
    String(r[1]).trim(), String(r[2]), String(r[3]),
  ]);
}

function sheet(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    sh.getRange("A:B").setNumberFormat("@");
  }
  return sh;
}
