/**
 * 구미은혜로교회 게시판 백엔드 (Google Apps Script)
 *
 * 사용 순서
 * 1) script.google.com 에서 새 프로젝트를 만들고 이 코드를 붙여 넣습니다.
 * 2) 상단 함수 선택에서 setup 을 골라 한 번 실행합니다 (권한 허용 필요). 이미 setup 을 했다면 다시 하지 않아도 됩니다.
 * 3) 프로젝트 설정(톱니바퀴) > 스크립트 속성 > 속성 추가
 *    - 속성: ADMIN_PASSWORD   값: 관리자 비밀번호 (글 삭제·수정을 관리자가 할 때 사용)
 * 4) 배포 > 새 배포 > 유형: 웹 앱 (실행 사용자: 나 / 액세스: 모든 사용자)
 *    이미 배포한 적이 있다면 배포 관리 > 연필 > 새 버전으로 다시 배포합니다. URL은 그대로입니다.
 *
 * 글 작성 시 정한 비밀번호로 본인 글을 수정·삭제할 수 있고, ADMIN_PASSWORD로는 모든 글을 수정·삭제할 수 있습니다.
 */

const SHEET_NAME = 'posts';
const HEADERS = ['id', 'board', 'title', 'author', 'link', 'body', 'createdAt', 'fileId', 'fileName', 'fileType', 'fileUrl', 'pwHash', 'updatedAt'];
const ALLOWED_BOARDS = ['service', 'bulletin', 'photo', 'music', 'articles', 'files'];
const MAX_FILE_BYTES = 10 * 1024 * 1024; // 10MB
const LIMITS = { title: 120, author: 40, link: 500, body: 5000 };

function setup() {
  const ss = SpreadsheetApp.create('구미은혜로교회 게시판 데이터');
  const sh = ss.getActiveSheet();
  sh.setName(SHEET_NAME);
  sh.appendRow(HEADERS);
  sh.setFrozenRows(1);

  const folder = DriveApp.createFolder('구미은혜로교회 첨부파일');
  PropertiesService.getScriptProperties().setProperties({
    SHEET_ID: ss.getId(),
    FOLDER_ID: folder.getId()
  });
  Logger.log('시트: ' + ss.getUrl());
  Logger.log('폴더: ' + folder.getUrl());
}

/* ---------- 조회 ---------- */
function doGet(e) {
  const board = (e.parameter.board || '').toString();
  if (!ALLOWED_BOARDS.includes(board)) return json_({ error: 'unknown board' });

  const values = sheet_().getDataRange().getValues();
  const head = values.shift();
  const rows = values
    .map(r => rowToObj_(head, r))
    .filter(r => r.board === board)
    .map(r => { delete r.pwHash; return r; })
    .sort((a, b) => Number(b.createdAt) - Number(a.createdAt));
  return json_(rows);
}

/* ---------- 작성 / 수정 / 삭제 ---------- */
function doPost(e) {
  let d;
  try {
    d = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad json' });
  }
  if (!ALLOWED_BOARDS.includes(d.board)) return json_({ ok: false, error: 'unknown board' });
  if (d.action === 'update') return updatePost_(d);
  if (d.action === 'delete') return deletePost_(d);
  return createPost_(d);
}

function createPost_(d) {
  if (!d.title) return json_({ ok: false, error: 'title required' });
  if (!d.password || String(d.password).length < 4) return json_({ ok: false, error: 'password too short' });

  const id = Utilities.getUuid();
  let fileId = '', fileName = '', fileType = '', fileUrl = '';

  if (d.fileData) {
    const bytes = Utilities.base64Decode(d.fileData);
    if (bytes.length > MAX_FILE_BYTES) return json_({ ok: false, error: 'file too large' });
    const folder = DriveApp.getFolderById(props_().FOLDER_ID);
    const blob = Utilities.newBlob(bytes, d.fileType || 'application/octet-stream', d.fileName || 'file');
    const f = folder.createFile(blob);
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    fileId = f.getId();
    fileName = d.fileName || '';
    fileType = d.fileType || '';
    fileUrl = fileType.startsWith('image/')
      ? 'https://drive.google.com/thumbnail?sz=w1200&id=' + fileId
      : 'https://drive.google.com/uc?export=download&id=' + fileId;
  }

  appendByHeader_(sheet_(), {
    id: id,
    board: d.board,
    title: clip_(d.title, LIMITS.title),
    author: clip_(d.author, LIMITS.author),
    link: clip_(d.link, LIMITS.link),
    body: clip_(d.body, LIMITS.body),
    createdAt: Date.now(),
    fileId: fileId,
    fileName: fileName,
    fileType: fileType,
    fileUrl: fileUrl,
    pwHash: hash_(id, String(d.password)),
    updatedAt: ''
  });
  return json_({ ok: true, id: id });
}

function updatePost_(d) {
  if (!d.title) return json_({ ok: false, error: 'title required' });
  const sh = sheet_();
  const found = findRow_(sh, d.id, d.board);
  if (!found) return json_({ ok: false, error: 'not found' });
  const obj = rowToObj_(found.head, found.row);
  if (!authorized_(obj, d.id, d.password)) return json_({ ok: false, error: 'wrong password' });

  const head = headers_(sh);
  ['title', 'author', 'link', 'body'].forEach(k => {
    if (d[k] === undefined) return;
    sh.getRange(found.rowNum, head.indexOf(k) + 1).setValue(clip_(d[k], LIMITS[k]));
  });
  sh.getRange(found.rowNum, head.indexOf('updatedAt') + 1).setValue(Date.now());
  return json_({ ok: true });
}

function deletePost_(d) {
  const sh = sheet_();
  const found = findRow_(sh, d.id, d.board);
  if (!found) return json_({ ok: false, error: 'not found' });
  const obj = rowToObj_(found.head, found.row);
  if (!authorized_(obj, d.id, d.password)) return json_({ ok: false, error: 'wrong password' });

  if (obj.fileId) {
    try { DriveApp.getFileById(obj.fileId).setTrashed(true); } catch (err) { Logger.log(err); }
  }
  sh.deleteRow(found.rowNum);
  return json_({ ok: true });
}

/* ---------- helpers ---------- */
function authorized_(obj, id, password) {
  const pw = String(password || '');
  if (!pw) return false;
  if (obj.pwHash && hash_(id, pw) === obj.pwHash) return true;
  const admin = props_().ADMIN_PASSWORD;
  return !!admin && pw === admin;
}

function findRow_(sh, id, board) {
  const values = sh.getDataRange().getValues();
  const head = values[0];
  const idI = head.indexOf('id');
  const bI = head.indexOf('board');
  for (let i = 1; i < values.length; i++) {
    if (values[i][idI] === id && values[i][bI] === board) {
      return { rowNum: i + 1, row: values[i], head: head };
    }
  }
  return null;
}

function headers_(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
}

// 예전에 만든 시트에 새 열(pwHash, updatedAt)이 없으면 자동으로 추가합니다.
function ensureHeaders_(sh) {
  const head = headers_(sh);
  HEADERS.forEach(h => {
    if (head.indexOf(h) < 0) {
      sh.getRange(1, head.length + 1).setValue(h);
      head.push(h);
    }
  });
  return head;
}

function appendByHeader_(sh, obj) {
  const head = ensureHeaders_(sh);
  sh.appendRow(head.map(h => (obj[h] === undefined ? '' : obj[h])));
}

function rowToObj_(head, row) {
  const o = {};
  head.forEach((h, i) => { o[h] = row[i]; });
  return o;
}

function hash_(id, password) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, id + '|' + password);
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function props_() { return PropertiesService.getScriptProperties().getProperties(); }

function sheet_() {
  const sh = SpreadsheetApp.openById(props_().SHEET_ID).getSheetByName(SHEET_NAME);
  ensureHeaders_(sh);
  return sh;
}

function clip_(s, n) { return String(s || '').slice(0, n); }

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
