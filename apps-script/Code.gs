/**
 * 구미은혜로교회 게시판 백엔드 (Google Apps Script)
 *
 * 사용 순서
 * 1) script.google.com 에서 새 프로젝트를 만들고 이 코드를 붙여 넣습니다.
 * 2) 상단 함수 선택에서 setup 을 골라 한 번 실행합니다 (권한 허용 필요).
 *    → 구글 시트 '구미은혜로교회 게시판 데이터'와 드라이브 폴더 '구미은혜로교회 첨부파일'이 만들어집니다.
 * 3) 배포 > 새 배포 > 유형: 웹 앱
 *    - 실행 사용자: 나
 *    - 액세스 권한: 모든 사용자
 *    배포 후 나오는 '웹 앱 URL'을 index.html 의 API_URL 에 넣습니다.
 *
 * 관리자 삭제: 구글 시트에서 해당 행을 지우고, 드라이브 폴더에서 같은 파일을 휴지통으로 보내면 됩니다.
 */

const SHEET_NAME = 'posts';
const HEADERS = ['id', 'board', 'title', 'author', 'link', 'body', 'createdAt', 'fileId', 'fileName', 'fileType', 'fileUrl'];
const ALLOWED_BOARDS = ['bulletin', 'photo', 'music', 'articles', 'files'];
const MAX_FILE_BYTES = 10 * 1024 * 1024; // 10MB (base64 기준 약 13MB)

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

function doGet(e) {
  const board = (e.parameter.board || '').toString();
  if (!ALLOWED_BOARDS.includes(board)) return json_({ error: 'unknown board' });

  const values = sheet_().getDataRange().getValues();
  const head = values.shift();
  const rows = values
    .map(r => Object.fromEntries(head.map((h, i) => [h, r[i]])))
    .filter(r => r.board === board)
    .sort((a, b) => Number(b.createdAt) - Number(a.createdAt));
  return json_(rows);
}

function doPost(e) {
  let d;
  try {
    d = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad json' });
  }
  if (!ALLOWED_BOARDS.includes(d.board)) return json_({ ok: false, error: 'unknown board' });
  if (!d.title) return json_({ ok: false, error: 'title required' });

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
    fileUrl = (fileType.startsWith('image/'))
      ? 'https://drive.google.com/thumbnail?sz=w1200&id=' + fileId
      : 'https://drive.google.com/uc?export=download&id=' + fileId;
  }

  sheet_().appendRow([
    id, d.board, clip_(d.title, 120), clip_(d.author, 40), clip_(d.link, 500),
    clip_(d.body, 5000), Date.now(), fileId, fileName, fileType, fileUrl
  ]);
  return json_({ ok: true, id: id });
}

/* ---------- helpers ---------- */
function props_() { return PropertiesService.getScriptProperties().getProperties(); }
function sheet_() { return SpreadsheetApp.openById(props_().SHEET_ID).getSheetByName(SHEET_NAME); }
function clip_(s, n) { return String(s || '').slice(0, n); }
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
