/**
 * はっぴょうメーカー — Google Apps Script ウェブアプリ用サーバーコード
 *
 * これを使うと、児童の作品が先生のGoogleドライブに集まります。
 * ・作品ファイル(JSON) … マイドライブ／はっぴょうメーカー_ていしゅつ／
 * ・一覧表(スプレッドシート) … 同じフォルダの「ていしゅつ一覧」
 *
 * デプロイ設定(重要):
 *   次のユーザーとして実行 = 自分(先生)
 *   アクセスできるユーザー = 学校のドメイン内の全員(または「全員」)
 *
 * 「全員(匿名)」で公開した場合は児童のメールが取れないため、
 * 先生モードは下の TEACHER_KEY をURLに付けて開きます。
 *   例) https://script.google.com/macros/s/xxxx/exec?key=あいことば
 */

// ★必ず自分だけのあいことばに書きかえてください(児童には教えない)
var TEACHER_KEY = 'sensei-himitsu-0000';

var FOLDER_NAME = 'はっぴょうメーカー_ていしゅつ';
var SHEET_NAME  = 'ていしゅつ一覧';
var TMP_NAME    = '_おくりちゅう';
var READ_CHUNK  = 200000;   // 先生が作品を読むときの分割サイズ

/* ========== アプリを表示する ========== */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('はっぴょうメーカー')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ========== だれが見ているか ========== */
function getWho(key) {
  var me = '';
  var you = '';
  try { me  = Session.getEffectiveUser().getEmail(); } catch (err) {}
  try { you = Session.getActiveUser().getEmail();    } catch (err) {}
  var isTeacher = (!!you && you === me) || (!!key && key === TEACHER_KEY);
  return { email: you, isTeacher: isTeacher };
}

function assertTeacher_(key) {
  if (!getWho(key).isTeacher) throw new Error('せんせい だけが つかえます');
}

/* ========== フォルダ・一覧表 ========== */
function getFolder_() {
  var it = DriveApp.getRootFolder().getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function getTmpFolder_() {
  var parent = getFolder_();
  var it = parent.getFoldersByName(TMP_NAME);
  return it.hasNext() ? it.next() : parent.createFolder(TMP_NAME);
}

function getSheet_() {
  var folder = getFolder_();
  var it = folder.getFilesByName(SHEET_NAME);
  var ss;
  if (it.hasNext()) {
    ss = SpreadsheetApp.open(it.next());
  } else {
    ss = SpreadsheetApp.create(SHEET_NAME);
    DriveApp.getFileById(ss.getId()).moveTo(folder);
    ss.getSheets()[0].appendRow(
      ['日時', 'なまえ', 'メール', 'まいすう', 'だいめい', 'かいた ことば', 'ファイルID']);
    ss.getSheets()[0].setFrozenRows(1);
  }
  return ss.getSheets()[0];
}

/* ========== 児童が作品を送る(3ステップ) ========== */
function startUpload(meta) {
  var id = Utilities.getUuid();
  var f = getTmpFolder_().createFolder(id);
  f.createFile('meta.json', JSON.stringify(meta || {}), MimeType.PLAIN_TEXT);
  return id;
}

function appendChunk(uploadId, index, text) {
  var f = findTmp_(uploadId);
  f.createFile(pad_(index) + '.part', text || '', MimeType.PLAIN_TEXT);
  return index;
}

function finishUpload(uploadId) {
  var f = findTmp_(uploadId);

  // meta を読む
  var meta = {};
  var mit = f.getFilesByName('meta.json');
  if (mit.hasNext()) {
    try { meta = JSON.parse(mit.next().getBlob().getDataAsString()); } catch (err) {}
  }

  // .part を番号順につなぐ
  var parts = [];
  var fit = f.getFiles();
  while (fit.hasNext()) {
    var file = fit.next();
    var nm = file.getName();
    if (nm.indexOf('.part') > 0) parts.push({ name: nm, file: file });
  }
  if (!parts.length) throw new Error('データが ありません');
  parts.sort(function (a, b) { return a.name < b.name ? -1 : 1; });
  var json = '';
  for (var i = 0; i < parts.length; i++) json += parts[i].file.getBlob().getDataAsString();

  // 中身が読めるか念のため確認
  var slides = meta.slides || 0;
  try {
    var d = JSON.parse(json);
    slides = (d && d.slides && d.slides.length) || slides;
  } catch (err) {
    throw new Error('データが こわれて います');
  }

  // 作品ファイルとして保存
  var who = getWho('');
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm');
  var name = String(meta.name || 'なまえなし').replace(/[\/\\:*?"<>|]/g, '_');
  var saved = getFolder_().createFile(stamp + '_' + name + '.json', json, MimeType.PLAIN_TEXT);

  // 一覧表に1行 追加
  getSheet_().appendRow([
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm'),
    meta.name || '',
    who.email || '',
    slides,
    meta.title || '',
    meta.text || '',
    saved.getId()
  ]);

  // 一時ファイルは片づける
  f.setTrashed(true);
  return { ok: true, file: saved.getName() };
}

function findTmp_(uploadId) {
  if (!uploadId) throw new Error('おくりさきが わかりません');
  var it = getTmpFolder_().getFoldersByName(String(uploadId));
  if (!it.hasNext()) throw new Error('おくる じゅんびが みつかりません');
  return it.next();
}

function pad_(n) {
  var s = '00000' + String(n);
  return s.slice(-5);
}

/* ========== 先生が見る ========== */
function listWorks(key) {
  assertTeacher_(key);
  var sh = getSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, 7).getValues();
  var out = [];
  for (var i = rows.length - 1; i >= 0; i--) {      // 新しいものから
    var r = rows[i];
    if (!r[6]) continue;
    out.push({
      date: String(r[0]),
      name: String(r[1]),
      email: String(r[2]),
      slides: r[3],
      title: String(r[4]),
      fileId: String(r[6])
    });
    if (out.length >= 300) break;
  }
  return out;
}

function getWorkChunk(fileId, index, key) {
  assertTeacher_(key);
  var text = DriveApp.getFileById(fileId).getBlob().getDataAsString();
  var start = index * READ_CHUNK;
  var slice = text.slice(start, start + READ_CHUNK);
  return { text: slice, done: (start + READ_CHUNK) >= text.length };
}

/* ========== 動作テスト用(スクリプトエディタから実行) ========== */
function テスト_フォルダを作る() {
  var f = getFolder_();
  getSheet_();
  Logger.log('できました: ' + f.getUrl());
}
