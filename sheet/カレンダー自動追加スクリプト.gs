/**
 * 【試作版V3対応】カレンダー自動案件追加スクリプト
 * =====================================================================
 * Googleカレンダーの「【撮影】」予定を毎朝スキャンし、試作版のマスターに
 * 動画案件の行を自動追加する（2026-07-03 グリル合意の設計）。
 *
 * ■ V3対応版での変更（2026-07-03）
 *   ・設定は専用シート「自動追加設定」に分離（設定シートのV3ステータス表と衝突しないため）
 *   ・マスターのV3レイアウト（22列・ステータス手動）に対応。
 *     新規行のステータスは「タイトル待ち」で入れる
 *   ・クライアント名は _m/_d サフィックス付きのシート表記で入れる
 *   ・初版が設定シートに書いてしまった内容の修復処理入り
 *     （名簿に混ざったステータス名の除去・S1:X1ヘッダーの復元）
 *
 * ■ 使い方（初回セットアップ）
 *   1. 関数「setupCalendarAutoAdd」を選んで実行
 *      → 「自動追加設定」「自動追加ログ」シートを作成
 *      → 設定シートの汚れを修復
 *      → 毎朝6時のトリガーを設置
 *   2. 動作確認は「autoAddFromCalendar」を手動実行
 *
 * ■ 仕組み（合意した設計）
 *   ・対象: タイトルに【撮影】を含む予定。（仮）で始まる予定はスキップ
 *   ・クライアント特定: 「自動追加設定」の表と最長一致（別名も見る）。一覧外はスキップ＆ログ
 *   ・予定1件＝1撮影。本数はクライアント別
 *   ・投稿日: 投稿曜日で、シート内の最終投稿予定日の続きから順に埋める。
 *     既存がない/過去なら 撮影日＋リードタイム 以降の最初の枠から。
 *     投稿曜日が未設定のクライアントは投稿日を空欄で追加（あとで手動入力）
 *   ・行の初期値: 仮タイトル「7/6撮影 ①」＋ステータス「タイトル待ち」＋BO担当・案件タイプ既定値
 *   ・追加のみ。撮影日の変更・予定の削除は自動では書き換えず、ログ＋メモ欄に⚠
 *   ・重複防止: イベントIDを「自動追加ログ」に記録
 */

var AUTO_TARGET_ID = '1oXwpCyw_g4GT1SFh4V_DW6RL7Q8h5FhIx9nbEV_eLXo'; // 【試作版V2】案件進捗管理シート（中身はV3）
var AUTO_CAL_ID    = '';   // 空欄なら自分のメインカレンダー
var CFG_SHEET      = '自動追加設定';
var LOG_SHEET      = '自動追加ログ';

// V3マスターの列番号
var COL = { client:1, type:3, title:4, shoot:6, post:7, status:8, bo:18, memo:19 };

/* ==================== 初回セットアップ ==================== */
function setupCalendarAutoAdd(){
  var ss = SpreadsheetApp.openById(AUTO_TARGET_ID);
  repairSettingsSheet_(ss);   // 初版が設定シートに書いた分の修復
  buildAutoAddCfg_(ss);
  buildAutoAddLog_(ss);
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'autoAddFromCalendar') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('autoAddFromCalendar').timeBased().everyDays(1).atHour(6).create();
  Logger.log('セットアップ完了。毎朝6時に自動追加が動きます。\n「自動追加設定」シートでクライアント別の本数・投稿曜日を調整できます。');
}

/* ---- 初版が設定シートに書いてしまった内容の修復（何度実行してもOK） ---- */
function repairSettingsSheet_(ss){
  var st = ss.getSheetByName('設定');
  if (!st) return;
  // 1) クライアント名簿(A列)に混ざったステータス名の行を除去
  var statusNames = ['タイトル待ち','未着手','編集中','修正中','サムネ待ち','砂田確認中',
                     'CL確認中','クライアント共有待ち','アップロード待ち','リンク共有待ち','完了'];
  var roster = st.getRange('A2:A100').getValues();
  for (var i = 0; i < roster.length; i++){
    if (statusNames.indexOf(String(roster[i][0]).trim()) >= 0){
      st.getRange(i + 2, 1, 1, 4).clearContent(); // A:D（名前＋リードタイム）
    }
  }
  // 2) S1:X1 のヘッダー復元（初版が上書きした場合のみ）
  if (st.getRange('S1').getValue() === 'クライアント名'){
    st.getRange('S1:U1').setValues([['ステータス','工程','ボール（担当）']])
      .setFontWeight('bold').setBackground('#f1efe8');
    st.getRange('V1:X1').clearContent().setFontWeight('normal').setBackground(null);
  }
  // 3) Z列に書いた動作設定を撤去（新しい置き場は「自動追加設定」シート）
  if (st.getRange('Z1').getValue() === '自動追加の動作設定'){
    st.getRange('Z1:AA5').clearContent().setFontWeight('normal').setBackground(null);
  }
}

/* ---- 「自動追加設定」シート（専用シートなのでV3の表と衝突しない） ---- */
function buildAutoAddCfg_(ss){
  var sh = ss.getSheetByName(CFG_SHEET);
  if (!sh){
    sh = ss.insertSheet(CFG_SHEET);
    sh.getRange('A1:F1').setValues([['クライアント名（シート表記）','本数/撮影','投稿曜日','BO担当既定','案件タイプ既定','別名（・区切り）']])
      .setFontWeight('bold').setBackground('#e8f0e8');
    // _d（ダイレクト＝直接依頼）の案件は自動追加の対象外。撮影予定はすべて _m（MUSUBI経由）に入れる
    // BO担当既定は既存案件の実績から（2026-07-03時点・全クライアントで固定を確認済み）
    sh.getRange('A2:F16').setValues([
      ['DEP_m',            8, '水・日', '増田', '', 'DEP'],
      ['ハイテクノ_m',      3, '木',    '増田', '', 'ハイテクノ'],
      ['バイオテック_m',    4, '金',    '増田', '', 'バイオテック'],
      ['mug_m',            4, '金',    '岩渕', '', 'mug'],
      ['アーバンガレージ_m', 4, '金',    '岩渕', '', 'アーバンガレージ'],
      ['1sec._m',          4, '火',    '増田', '', '1sec.'],
      ['そうぞう_m',        4, '火',    '岩渕', '', 'そうぞう'],
      ['角川春樹',          4, '月',    '増田', '', ''],
      ['髭男会計士',        4, '',      '岩渕', '', '宮川さん'],
      ['四国物産_m',        1, '',      '岩渕', '', '四国物産'],
      ['井口鉱油_m',        2, '',      '岩渕', '', '井口鉱油'],
      ['EmpowerX',         1, '',      '岩渕', '', ''],
      ['千勝会_三浦',       4, '',      '岩渕', '', ''],
      ['Enny',             4, '',      '砂田', '', ''],
      ['（対象外）',        0, '',      '',    '', 'The Neutral・DEPサミット']
    ]);
    sh.getRange('H1').setValue('動作設定').setFontWeight('bold').setBackground('#e8f0e8');
    sh.getRange('H2:I7').setValues([
      ['初回投稿リードタイム（日）', 14],
      ['スキャン範囲：過去（日）', 7],
      ['スキャン範囲：未来（日）', 90],
      ['標準本数（表で空欄のとき）', 4],
      ['新規行のステータス', 'タイトル待ち'],
      ['Discord Webhook URL', '']
    ]);
    sh.setColumnWidth(1, 190); sh.setColumnWidth(3, 90); sh.setColumnWidth(6, 260); sh.setColumnWidth(8, 200);
    sh.getRange('A1').setNote(
      'カレンダー自動追加のクライアント表。\n' +
      '・カレンダーのタイトル（【撮影】のあと）と、ここの名前・別名を最長一致で照合します。\n' +
      '・クライアント名はマスターのシート表記（_m/_d付き）で書いてください。\n' +
      '・本数/撮影 … 予定1件で追加する行数（空欄なら標準本数）\n' +
      '・投稿曜日 … 「水・日」のように複数OK。空欄なら投稿日は空のまま追加\n' +
      '・（対象外）の行 … 別名に入れた名前はカウントせずスキップ\n' +
      '・別名が複数の名前に当てはまるときは、長い方が勝ちます');
  }
}

/* ---- ログシート ---- */
function buildAutoAddLog_(ss){
  var lg = ss.getSheetByName(LOG_SHEET);
  if (!lg){
    lg = ss.insertSheet(LOG_SHEET);
    lg.getRange('A1:G1').setValues([['処理日時','イベントID','予定タイトル','クライアント','撮影日','状態','詳細']])
      .setFontWeight('bold').setBackground('#f1efe8');
    lg.setFrozenRows(1);
    lg.setColumnWidth(1, 130); lg.setColumnWidth(2, 90); lg.setColumnWidth(3, 220);
    lg.setColumnWidth(6, 140); lg.setColumnWidth(7, 260);
  }
}

/* ==================== 毎朝の本体 ==================== */
function autoAddFromCalendar(){
  var ss  = SpreadsheetApp.openById(AUTO_TARGET_ID);
  var ms  = ss.getSheetByName('マスター');
  var lg  = ss.getSheetByName(LOG_SHEET);
  if (!lg){ buildAutoAddLog_(ss); lg = ss.getSheetByName(LOG_SHEET); }
  var cal = AUTO_CAL_ID ? CalendarApp.getCalendarById(AUTO_CAL_ID) : CalendarApp.getDefaultCalendar();

  var cfg     = readAutoConfig_(ss);
  var logRows = readLog_(lg);
  var notes   = []; // Discord通知用

  var now    = new Date();
  var from   = addDays_(dateOnly_(now), -cfg.pastDays);
  var to     = addDays_(dateOnly_(now), cfg.futureDays);
  var events = cal.getEvents(from, to);

  events.forEach(function(ev){
    var title = ev.getTitle() || '';
    if (title.indexOf('【撮影】') < 0) return;
    var id = ev.getId();
    var shoot = dateOnly_(ev.getStartTime());

    if (title.indexOf('（仮）') === 0 || title.indexOf('(仮)') === 0){
      logOnce_(lg, logRows, id, title, '', shoot, '（仮）のためスキップ', '');
      return;
    }
    if (hasLog_(logRows, id, '追加済')) return;

    var raw = title.substring(title.indexOf('【撮影】') + 4).split('@')[0].trim();
    var hit = matchClient_(raw, cfg.patterns);
    if (!hit){
      if (logOnce_(lg, logRows, id, title, raw, shoot, '一覧に無い名前', '「自動追加設定」シートの表に名前か別名を足すと次回から拾われます')){
        notes.push('❓ 「' + raw + '」（' + md_(shoot) + '撮影）は一覧に無い名前のためスキップしました。自動追加設定シートに別名を足すと次回から拾われます');
      }
      return;
    }
    if (hit.client === '（対象外）'){
      logOnce_(lg, logRows, id, title, raw, shoot, '対象外', '必要なら手動で行を追加してください');
      return;
    }

    var c = cfg.clients[hit.client];
    var count = c.count > 0 ? c.count : cfg.defaultCount;
    var res = addRows_(ms, hit.client, c, count, shoot, cfg);
    var postInfo = res.posts[0] instanceof Date
      ? '投稿 ' + md_(res.posts[0]) + '〜' + md_(res.posts[res.posts.length - 1])
      : '投稿日は手動入力待ち';
    if (res.added < count){
      appendLog_(lg, id, title, hit.client, shoot, 'エラー：空き行不足', res.added + '/' + count + '本のみ追加。マスターの空き行を増やしてください');
      notes.push('🚨 ' + hit.client + '：空き行不足で ' + res.added + '/' + count + '本しか追加できませんでした');
    } else {
      appendLog_(lg, id, title, hit.client, shoot, '追加済', count + '本追加（' + (c.days.length ? '投稿日自動' : '投稿日は手動入力待ち') + '）');
      notes.push('✅ ' + hit.client + '：' + md_(shoot) + '撮影 → ' + count + '本追加（' + postInfo + '）');
    }
    logRows.push({ id: id, client: hit.client, status: '追加済', shoot: fmt_(shoot), detail: '' });
  });

  detectChanges_(cal, ms, lg, logRows, notes);
  postDiscord_(cfg.webhook, notes, lg);

  // 進捗リマインダー（進捗確認リマインダー.gs）。失敗しても自動追加には影響させない
  try { sendProgressReminders_(ss, lg); }
  catch(e){ try { appendLog_(lg, '', '', '', '', '⚠進捗リマインダー失敗', String(e)); } catch(e2){} }
}

/* ==================== 設定の読み込み ==================== */
function readAutoConfig_(ss){
  var sh = ss.getSheetByName(CFG_SHEET);
  if (!sh) throw new Error('「' + CFG_SHEET + '」シートがありません。先に setupCalendarAutoAdd を実行してください。');
  var scal = {};
  sh.getRange('H2:I10').getValues().forEach(function(r){ if (r[0]) scal[String(r[0])] = r[1]; });
  var clients = {};
  var patterns = [];
  sh.getRange('A2:F60').getValues().forEach(function(r){
    var name = String(r[0]).trim();
    if (!name) return;
    clients[name] = {
      count: Number(r[1]) || 0,
      days:  parseDays_(String(r[2])),
      bo:    String(r[3]).trim(),
      type:  String(r[4]).trim()
    };
    patterns.push({ pat: name, client: name });
    String(r[5]).split(/[・、,\/]/).forEach(function(a){
      a = a.trim();
      if (a) patterns.push({ pat: a, client: name });
    });
  });
  patterns.sort(function(a, b){ return b.pat.length - a.pat.length; }); // 最長一致
  return {
    clients: clients,
    patterns: patterns,
    leadDays:     Number(scal['初回投稿リードタイム（日）']) || 14,
    pastDays:     Number(scal['スキャン範囲：過去（日）']) || 7,
    futureDays:   Number(scal['スキャン範囲：未来（日）']) || 90,
    defaultCount: Number(scal['標準本数（表で空欄のとき）']) || 4,
    newStatus:    String(scal['新規行のステータス'] || 'タイトル待ち'),
    webhook:      String(scal['Discord Webhook URL'] || '').trim(),
    lastPostCache: {}
  };
}

// 「水・日」→ [3,0]（区切りなしの「水日」もOK）
function parseDays_(s){
  var map = { '日':0, '月':1, '火':2, '水':3, '木':4, '金':5, '土':6 };
  var out = [];
  for (var i = 0; i < s.length; i++){
    if (s[i] in map && out.indexOf(map[s[i]]) < 0) out.push(map[s[i]]);
  }
  return out;
}

// タイトル先頭とクライアント名・別名の最長一致（「様」「さん」等の付記はそのまま通る）
function matchClient_(raw, patterns){
  for (var i = 0; i < patterns.length; i++){
    if (raw.indexOf(patterns[i].pat) === 0) return patterns[i];
  }
  return null;
}

// クライアント名の比較（大文字小文字ゆれ「empowerx」等を吸収）
function sameClient_(a, b){
  return String(a).trim().toLowerCase() === String(b).trim().toLowerCase();
}

/* ==================== 行の追加（V3レイアウト） ==================== */
function addRows_(ms, client, c, count, shoot, cfg){
  var rowsN = Math.min(1000, ms.getMaxRows() - 1);
  var vals = ms.getRange(2, 1, rowsN, 7).getValues(); // A〜G
  var empty = [];
  for (var i = 0; i < vals.length; i++){
    if (String(vals[i][0]).trim() === '' && empty.length < count) empty.push(i + 2);
  }
  var posts = calcPostDates_(vals, client, c.days, count, shoot, cfg);
  for (var k = 0; k < empty.length; k++){
    var r = empty[k];
    ms.getRange(r, COL.client).setValue(client);
    if (c.type) ms.getRange(r, COL.type).setValue(c.type);
    ms.getRange(r, COL.title).setValue(kariTitle_(shoot, k + 1));
    ms.getRange(r, COL.shoot).setValue(shoot);
    if (posts[k]) ms.getRange(r, COL.post).setValue(posts[k]);
    ms.getRange(r, COL.status).setValue(cfg.newStatus);
    if (c.bo) ms.getRange(r, COL.bo).setValue(c.bo);
  }
  return { added: empty.length, posts: posts };
}

// 仮タイトル「7/6撮影 ①」
function kariTitle_(shoot, n){
  var maru = n <= 20 ? String.fromCharCode(0x2460 + n - 1) : '#' + n;
  return (shoot.getMonth() + 1) + '/' + shoot.getDate() + '撮影 ' + maru;
}

/* ---- 投稿日: 最終投稿予定日の続き。なければ撮影日＋リードタイム以降の最初の枠 ---- */
function calcPostDates_(vals, client, days, count, shoot, cfg){
  var out = [];
  if (!days.length){
    for (var i = 0; i < count; i++) out.push('');
    return out;
  }
  var last = cfg.lastPostCache[client] || null;
  if (!last){
    vals.forEach(function(r){
      if (sameClient_(r[0], client) && r[6] instanceof Date){
        if (!last || r[6] > last) last = dateOnly_(r[6]);
      }
    });
  }
  var floor = addDays_(shoot, cfg.leadDays);
  var cur = last ? addDays_(last, 1) : floor;
  if (cur < floor) cur = floor;
  while (out.length < count){
    if (days.indexOf(cur.getDay()) >= 0) out.push(new Date(cur));
    cur = addDays_(cur, 1);
  }
  cfg.lastPostCache[client] = out[out.length - 1];
  return out;
}

/* ==================== 変更・削除の検知（通知のみ） ==================== */
function detectChanges_(cal, ms, lg, logRows, notes){
  logRows.forEach(function(row){
    if (row.status !== '追加済') return;
    var ev = null;
    try { ev = cal.getEventById(row.id); } catch(e){}
    if (!ev){
      if (hasLog_(logRows, row.id, '⚠予定削除')) return;
      appendLog_(lg, row.id, '', row.client || '', row.shoot, '⚠予定削除', 'カレンダーから撮影予定が消えています。行の要否を確認してください');
      logRows.push({ id: row.id, client: row.client, status: '⚠予定削除', shoot: row.shoot, detail: '' });
      markMemo_(ms, row.client, row.shoot, '⚠カレンダーから撮影予定が削除されています（要確認）');
      notes.push('⚠ ' + (row.client || '?') + '：' + row.shoot + ' の撮影予定がカレンダーから消えています。シートの行の要否を確認してください');
      return;
    }
    var d = fmt_(dateOnly_(ev.getStartTime()));
    if (d !== row.shoot){
      var detail = row.shoot + ' → ' + d;
      if (hasLogDetail_(logRows, row.id, '⚠撮影日変更', detail)) return;
      appendLog_(lg, row.id, ev.getTitle(), row.client || '', row.shoot, '⚠撮影日変更', detail + '。シートの撮影日・投稿日を手動で直してください');
      logRows.push({ id: row.id, client: row.client, status: '⚠撮影日変更', shoot: row.shoot, detail: detail });
      markMemo_(ms, row.client, row.shoot, '⚠カレンダーで撮影日が' + detail + 'に変更されています（要確認）');
      notes.push('⚠ ' + (row.client || '?') + '：撮影日が ' + detail + ' に変更されています。シートの撮影日・投稿日を直してください');
    }
  });
}

// 該当行（クライアント×撮影日）のメモ欄に⚠を追記（同じ内容は二重追記しない）
function markMemo_(ms, client, shootStr, note){
  if (!client) return;
  var rowsN = Math.min(1000, ms.getMaxRows() - 1);
  var vals = ms.getRange(2, 1, rowsN, COL.memo).getValues();
  for (var i = 0; i < vals.length; i++){
    var f = vals[i][COL.shoot - 1] instanceof Date ? fmt_(dateOnly_(vals[i][COL.shoot - 1])) : '';
    if (sameClient_(vals[i][0], client) && f === shootStr){
      var memo = String(vals[i][COL.memo - 1] || '');
      if (memo.indexOf(note) < 0){
        ms.getRange(i + 2, COL.memo).setValue(memo ? memo + ' / ' + note : note);
      }
    }
  }
}

/* ==================== ログの読み書き ==================== */
function readLog_(lg){
  var n = lg.getLastRow();
  if (n < 2) return [];
  return lg.getRange(2, 2, n - 1, 6).getValues().map(function(r){
    return {
      id: String(r[0]),
      client: String(r[2]),
      shoot: r[3] instanceof Date ? fmt_(dateOnly_(r[3])) : String(r[3]),
      status: String(r[4]),
      detail: String(r[5])
    };
  });
}

function hasLog_(rows, id, status){
  return rows.some(function(r){ return r.id === id && r.status === status; });
}
function hasLogDetail_(rows, id, status, detail){
  return rows.some(function(r){ return r.id === id && r.status === status && r.detail.indexOf(detail) === 0; });
}

function logOnce_(lg, rows, id, title, client, shoot, status, detail){
  if (hasLog_(rows, id, status)) return false;
  appendLog_(lg, id, title, client, shoot, status, detail);
  rows.push({ id: id, client: client, shoot: fmt_(shoot), status: status, detail: detail });
  return true;
}

function appendLog_(lg, id, title, client, shoot, status, detail){
  lg.appendRow([new Date(), id, title, client, shoot, status, detail]);
}

/* ==================== Discord通知 ==================== */
// Webhook設定の動作確認用（これを実行してDiscordにテスト通知が届けばOK）
function testDiscordNotify(){
  var ss = SpreadsheetApp.openById(AUTO_TARGET_ID);
  var cfg = readAutoConfig_(ss);
  if (!cfg.webhook){
    Logger.log('「自動追加設定」シートの「Discord Webhook URL」が空です。H列のラベルの隣（I列）にWebhookのURLを貼ってください。');
    return;
  }
  postDiscord_(cfg.webhook, ['✅ テスト通知です。この仕組みで毎朝の自動追加をお知らせします！'], ss.getSheetByName(LOG_SHEET));
  Logger.log('テスト通知を送りました。Discordのチャンネルを確認してください。');
}

// 追加・⚠・スキップをまとめて1通で投稿。URL未設定や通知ゼロの日は何もしない
// 冒頭に@everyone＋シートへのリンク付き
function postDiscord_(webhook, notes, lg){
  if (!webhook || !notes.length) return;
  var header = '@everyone\n' +
    '📋 **案件シート自動追加**（' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'M/d HH:mm') + '）\n' +
    '🔗 シートを開く → https://docs.google.com/spreadsheets/d/' + AUTO_TARGET_ID + '/edit\n';
  var chunks = [], cur = header;
  notes.forEach(function(line){
    if ((cur + line).length > 1900){ chunks.push(cur); cur = ''; } // Discordの2000字制限対策
    cur += line + '\n';
  });
  chunks.push(cur);
  chunks.forEach(function(msg){
    try {
      UrlFetchApp.fetch(webhook, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          content: msg,
          allowed_mentions: { parse: ['everyone'] } // @everyoneを確実に通知として効かせる
        }),
        muteHttpExceptions: true
      });
    } catch(e){
      try { appendLog_(lg, '', '', '', '', '⚠Discord通知失敗', String(e)); } catch(e2){}
    }
  });
}

/* ==================== 日付ユーティリティ ==================== */
function dateOnly_(d){ return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function addDays_(d, n){ var x = new Date(d); x.setDate(x.getDate() + n); return x; }
function fmt_(d){ return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd'); }
function md_(d){ return (d.getMonth() + 1) + '/' + d.getDate(); }
