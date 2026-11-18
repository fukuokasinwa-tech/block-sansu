/**
 * ブロック さんすう たんけん：きろくを Google スプレッドシートに ほぞんする プログラム
 *
 * つかいかた（くわしくは CLOUD-SETUP.md）
 *  1. この ファイルの なかみを ぜんぶ、スプレッドシートの「拡張機能 → Apps Script」に はりつける
 *  2. したの AIKOTOBA を、かぞくだけが しっている ことばに かえる（このリポジトリには かかない）
 *  3. 「デプロイ → 新しいデプロイ → ウェブアプリ」
 *     ・次のユーザーとして実行：自分
 *     ・アクセスできるユーザー：全員
 */
const AIKOTOBA = 'ここに あいことばを かく';
const SHEET_NAME = 'players';
const DELETED = '__deleted__';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['id', 'name', 'updated', 'data']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function ok_(key) {
  return AIKOTOBA && AIKOTOBA !== 'ここに あいことばを かく' && key === AIKOTOBA;
}

// よみこみ：?action=list&key=あいことば
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!ok_(p.key)) return out_({ ok: false, error: 'key' });
  if (p.action === 'ping') return out_({ ok: true });
  const rows = sheet_().getDataRange().getValues().slice(1);
  const players = rows.filter(r => r[0] !== '').map(r => ({
    id: String(r[0]), name: String(r[1]), updated: Number(r[2]) || 0, data: String(r[3] || '')
  }));
  return out_({ ok: true, players: players });
}

// かきこみ：{ key, players: [{ id, name, updated, data }] }
// あたらしい きろく（updated が おおきい）だけを うわがきする
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'json' }); }
  if (!ok_(body.key)) return out_({ ok: false, error: 'key' });
  const items = Array.isArray(body.players) ? body.players : [];
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const values = sh.getDataRange().getValues();
    const rowOf = {};
    for (let i = 1; i < values.length; i++) rowOf[String(values[i][0])] = i + 1;
    items.forEach(it => {
      if (!it || !it.id) return;
      const id = String(it.id).slice(0, 40);
      const updated = Number(it.updated) || Date.now();
      const data = it.data === DELETED ? DELETED : String(it.data || '').slice(0, 45000);
      const row = [id, String(it.name || '').slice(0, 20), updated, data];
      const r = rowOf[id];
      if (r) {
        const cur = Number(sh.getRange(r, 3).getValue()) || 0;
        if (updated >= cur) sh.getRange(r, 1, 1, 4).setValues([row]);
      } else {
        sh.appendRow(row);
        rowOf[id] = sh.getLastRow();
      }
    });
  } finally {
    lock.releaseLock();
  }
  return out_({ ok: true });
}
