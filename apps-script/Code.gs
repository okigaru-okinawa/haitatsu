/**
 * 配達ルートアプリ — Google Apps Script（スプレッドシートに紐づけて使う）
 *
 * スクリプトプロパティに次の2つを設定する（プロジェクトの設定 → スクリプト プロパティ）
 *   API_KEY : Google Maps Platform の APIキー（Routes API / Geocoding API）
 *   TOKEN   : アプリとの合言葉（英数字20文字くらい）
 *
 * ウェブアプリとしてデプロイ：実行ユーザー＝自分、アクセス＝全員
 */

const MASTER = '薬局マスタ';
const LOG_SHEETS = ['便の記録', '配達記録', '変更の記録'];
// 薬局マスタの列（0始まり）
const C = { id:0, name:1, kana:2, addr:3, tel:4, lat:5, lng:6, plat:7, plng:8, tags:9, note:10, visits:11, avg:12, last:13, src:14 };

function props_() { return PropertiesService.getScriptProperties(); }
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function checkToken_(t) {
  const tok = props_().getProperty('TOKEN');
  if (!tok || t !== tok) throw new Error('token');
}

/** GET: 薬局マスタを返す */
function doGet(e) {
  try {
    checkToken_(e.parameter.t);
    if (e.parameter.a === 'master') return out_({ ok: true, master: readMaster_() });
    if (e.parameter.a === 'ping') return out_({ ok: true });
    return out_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  }
}

/** POST: ルート計算・記録・備考 */
function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents);
    checkToken_(b.t);
    switch (b.a) {
      case 'routes': return out_({ ok: true, data: computeRoutes_(b.body, b.mask) });
      case 'log':    return out_({ ok: true, n: appendLogs_(b.items) });
      case 'note':   return out_({ ok: true, row: updateNote_(b) });
      case 'visit':  return out_({ ok: true, row: updateVisit_(b) });
      case 'add':    return out_({ ok: true, row: addPharmacy_(b) });
      case 'delTrip': return out_({ ok: true, n: deleteRecords_(b.id, null) });
      case 'delStop': return out_({ ok: true, n: deleteRecords_(b.trip, b.id) });
      default:       return out_({ ok: false, error: 'unknown action' });
    }
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  }
}

function readMaster_() {
  const v = ss_().getSheetByName(MASTER).getDataRange().getValues();
  const res = [];
  for (let i = 1; i < v.length; i++) {
    const r = v[i];
    if (!r[C.id]) continue;
    res.push({
      id: String(r[C.id]), name: r[C.name], kana: r[C.kana], addr: r[C.addr], tel: String(r[C.tel] || ''),
      lat: num_(r[C.lat]), lng: num_(r[C.lng]), plat: num_(r[C.plat]), plng: num_(r[C.plng]),
      tags: String(r[C.tags] || ''), note: String(r[C.note] || ''),
      visits: Number(r[C.visits]) || 0, avg: num_(r[C.avg]),
      last: r[C.last] instanceof Date ? Utilities.formatDate(r[C.last], 'Asia/Tokyo', 'yyyy-MM-dd') : String(r[C.last] || '')
    });
  }
  return res;
}
function num_(x) { return (x === '' || x === null || x === undefined || isNaN(Number(x)) || Number(x) === 0) ? null : Number(x); }

function findRow_(sh, id) {
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  throw new Error('pharmacy not found: ' + id);
}

/** Routes API（computeRoutes）を代理で呼ぶ */
function computeRoutes_(body, reqMask) {
  const key = props_().getProperty('API_KEY');
  // 返してもらう項目。アプリ側から指定があればそれを使う（英字・点・カンマのみ許可）
  const mask = (typeof reqMask === 'string' && /^routes\.[A-Za-z.,]+$/.test(reqMask)) ? reqMask : [
    'routes.optimizedIntermediateWaypointIndex',
    'routes.legs.duration', 'routes.legs.distanceMeters',
    'routes.legs.polyline.encodedPolyline'
  ].join(',');
  const res = UrlFetchApp.fetch('https://routes.googleapis.com/directions/v2:computeRoutes', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'X-Goog-Api-Key': key, 'X-Goog-FieldMask': mask },
    payload: JSON.stringify(body)
  });
  const code = res.getResponseCode();
  const json = JSON.parse(res.getContentText() || '{}');
  if (code !== 200) throw new Error('Routes API ' + code + ': ' + (json.error && json.error.message));
  return json;
}

/** 記録シートに追記。items = [{sheet, row:[...]}, ...] */
function appendLogs_(items) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    let n = 0;
    (items || []).forEach(it => {
      if (LOG_SHEETS.indexOf(it.sheet) < 0) return;
      ss_().getSheetByName(it.sheet).appendRow(it.row); n++;
    });
    return n;
  } finally { lock.releaseLock(); }
}

/** タグ・備考・駐車位置を保存 */
function updateNote_(b) {
  const sh = ss_().getSheetByName(MASTER);
  const r = findRow_(sh, b.id);
  if (b.tags !== undefined) sh.getRange(r, C.tags + 1).setValue(b.tags);
  if (b.note !== undefined) sh.getRange(r, C.note + 1).setValue(b.note);
  if (b.plat !== undefined) { sh.getRange(r, C.plat + 1).setValue(b.plat); sh.getRange(r, C.plng + 1).setValue(b.plng); }
  return r;
}

/** 訪問回数・平均滞在・最終訪問日を更新 */
function updateVisit_(b) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = ss_().getSheetByName(MASTER);
    const r = findRow_(sh, b.id);
    const cur = sh.getRange(r, C.visits + 1, 1, 3).getValues()[0];
    const visits = (Number(cur[0]) || 0) + 1;
    let avg = num_(cur[1]);
    if (b.dwell !== null && b.dwell !== undefined && b.dwell > 0 && b.dwell < 90) {
      avg = avg === null ? b.dwell : Math.round(((avg * (visits - 1)) + b.dwell) / visits * 10) / 10;
    }
    sh.getRange(r, C.visits + 1, 1, 3).setValues([[visits, avg === null ? '' : avg, b.date || '']]);
    return r;
  } finally { lock.releaseLock(); }
}

/**
 * 記録を消す（テスト用）。tripId の便の記録を消す。pharmId を渡すと、その便のその薬局の配達記録だけ消す。
 * 消したあと、関係する薬局の「訪問回数・平均滞在分・最終訪問日」を、残っている配達記録から計算し直す
 */
function deleteRecords_(tripId, pharmId) {
  if (!tripId) throw new Error('trip id');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const touched = {}; let n = 0;
    (pharmId ? ['配達記録'] : LOG_SHEETS).forEach(name => {
      const sh = ss_().getSheetByName(name), v = sh.getDataRange().getValues();
      for (let i = v.length - 1; i >= 1; i--) {
        if (String(v[i][0]) !== String(tripId)) continue;
        if (pharmId && String(v[i][2]) !== String(pharmId)) continue;
        if (name === '配達記録') touched[String(v[i][2])] = true;
        sh.deleteRow(i + 1); n++;
      }
    });
    recomputeVisits_(Object.keys(touched));
    return n;
  } finally { lock.releaseLock(); }
}
function recomputeVisits_(ids) {
  if (!ids.length) return;
  const rows = ss_().getSheetByName('配達記録').getDataRange().getValues().slice(1);
  const sh = ss_().getSheetByName(MASTER);
  ids.forEach(id => {
    let visits = 0, sum = 0, cnt = 0, last = '';
    rows.forEach(r => {
      if (String(r[2]) !== id || String(r[9]).indexOf('完了') !== 0) return;
      visits++;
      const d = Number(r[8]); if (r[8] !== '' && d > 0 && d < 90) { sum += d; cnt++; }
      const day = r[1] instanceof Date ? Utilities.formatDate(r[1], 'Asia/Tokyo', 'yyyy-MM-dd') : String(r[1]);
      if (day > last) last = day;
    });
    let row; try { row = findRow_(sh, id); } catch (e) { return; }
    sh.getRange(row, C.visits + 1, 1, 3).setValues([[visits, cnt ? Math.round(sum / cnt * 10) / 10 : '', last]]);
  });
}

/** 新しい薬局を追加（住所から位置を求める） */
function addPharmacy_(b) {
  const g = geocode_(b.addr);
  const sh = ss_().getSheetByName(MASTER);
  const id = 'U' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyMMddHHmmss');
  const row = [id, b.name, '', b.addr, b.tel || '', g ? g.lat : '', g ? g.lng : '', '', '', '', '', 0, '', '', '手入力'];
  sh.appendRow(row);
  return { id: id, name: b.name, kana: '', addr: b.addr, tel: b.tel || '', lat: g ? g.lat : null, lng: g ? g.lng : null,
           plat: null, plng: null, tags: '', note: '', visits: 0, avg: null, last: '' };
}

function geocode_(addr) {
  const key = props_().getProperty('API_KEY');
  const url = 'https://maps.googleapis.com/maps/api/geocode/json?language=ja&region=jp&address=' +
              encodeURIComponent(addr) + '&key=' + key;
  const json = JSON.parse(UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText());
  if (json.status !== 'OK' || !json.results.length) return null;
  const loc = json.results[0].geometry.location;
  return { lat: loc.lat, lng: loc.lng };
}

/** 【最初に1回だけ手で実行】緯度経度が空の薬局を、住所から埋める */
function 位置をまとめて取得() {
  const sh = ss_().getSheetByName(MASTER);
  const v = sh.getDataRange().getValues();
  let done = 0, fail = [];
  for (let i = 1; i < v.length; i++) {
    const la = Number(v[i][C.lat]);
    if (v[i][C.lat] !== '' && v[i][C.lat] !== null && la > 20) continue; // 0 や空は取り直す
    if (!v[i][C.addr]) continue;
    const g = geocode_(String(v[i][C.addr]));
    if (g) { sh.getRange(i + 1, C.lat + 1, 1, 2).setValues([[g.lat, g.lng]]); done++; }
    else fail.push(v[i][C.name]);
    Utilities.sleep(60);
  }
  Logger.log('位置を入れた件数: ' + done + ' / 失敗: ' + fail.length + ' ' + fail.join('、'));
}

/** 【動作確認用】APIキーとトークンが入っているか確認 */
function 設定チェック() {
  Logger.log('API_KEY: ' + (props_().getProperty('API_KEY') ? 'OK' : '未設定'));
  Logger.log('TOKEN: ' + (props_().getProperty('TOKEN') ? 'OK' : '未設定'));
  Logger.log('薬局マスタ: ' + (ss_().getSheetByName(MASTER).getLastRow() - 1) + '件');
}

/** 【確認用】住所変換が動くか1件だけ試す */
function 住所変換テスト() {
  const key = props_().getProperty('API_KEY');
  const url = 'https://maps.googleapis.com/maps/api/geocode/json?language=ja&address=' +
              encodeURIComponent('沖縄県那覇市泉崎1-2-2') + '&key=' + key;
  Logger.log(UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText().slice(0, 600));
}
