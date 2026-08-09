/**
 * 【1回だけ実行】名前→Discord ID対応表の一括入力（2026-08-01）
 * =====================================================================
 * Discordログから抽出したメンバー表を「自動追加設定」シートのK:L列に流し込む。
 *
 * ■ 使い方
 *   1. このファイルをApps Scriptプロジェクトに追加（コード.gsに貼ってもOK）
 *   2. 関数「fillMentionTable」を実行（何度実行しても二重追加はされない）
 *   3. 実行後にやること：
 *      ・「岩渕」のIDが空欄 → Discordの「小渕」か「Yuika.I」のどちらが
 *        岩渕さんか確認して、該当のIDをコピーして埋める
 *      ・マスターの制作担当列で使っている表記（例:「千野」）と
 *        Discord表示名（例:「千野(ちの)_映像制作」）が違う人は、
 *        K列を担当列の表記に書き換える（IDはそのままでOK）
 *      ・担当列に出てこない人の行は消してもOK（残っていても害はない）
 *   4. 終わったらこのファイルは削除してOK
 *
 * ■ 注意
 *   DiscordのIDは18〜19桁で、数値として入ると末尾が丸められて壊れるため、
 *   この関数はL列を文字列書式にしてから入力する。手で追記するときも
 *   L列にそのまま貼れば文字列として入る。
 */

// [名前, DiscordユーザーID]。IDが「・」区切りの行は全員メンション（GS等）
var MENTION_ROWS = [
  // --- 担当列の表記が確定している人 ---
  ['増田',   '1415909781457076256'],  // 増田 育美_土日祝休み
  ['岩渕',   ''],                      // ⚠要確認: 「小渕」1375097471193190425 か「Yuika.I」1394647966454775829 か
  ['砂田',   '895528906302554182'],   // ITARU SUNADA
  ['GS',     '1131131638777118722・1426914585394417857'], // 田中ゆか里・舟瀬 渚 の2人
  // --- 以下はDiscord表示名のまま。担当列の表記に合わせてK列を直して使う ---
  ['田中ゆか里',        '1131131638777118722'],
  ['舟瀬 渚',           '1426914585394417857'],
  ['Yuika.I',           '1394647966454775829'],
  ['ASATO YUICHI',      '1233366519057219594'],
  ['小渕',              '1375097471193190425'],
  ['waryo',             '1320242584005640257'],
  ['千野(ちの)_映像制作', '1164871855782506527'],
  ['中畑 美波',          '1301405727939760161'],
  ['HIDEKI ZUNIGA',     '1287282022896762930'],
  ['松金智志',           '1292331753402404917'],
  ['榎本千帆',           '1334167859256889345'],
  ['金 祥鎬/KIN SHOKO',  '1335090607235792898'],
  ['イカラシ',           '1398850228328730625'],
  ['野村友作',           '763212900319756328'],
  ['森下蓮',             '211795495390019584'],
  ['上園千尋',           '1147413277094191205'],
  ['中込 健人',          '875686638573989969'],
  ['Yuki Hirano',       '825192029964599297'],
  ['岡 篤志',            '1337965596868087828'],
  ['沼田芳希',           '599582387310952449'],
  ['うえ',              '1112940044781629534'],
  ['古川聡史',           '1493477555985190973'],
  ['murakami noriko',   '1374534191219343422'],
  ['坂上空良',           '1374172187321176095'],
  ['2710_yuki',         '1391959851055321128'],
  ['ひーこ',             '1317113140844040192'],
  ['山地英明',           '546046176265240596']
];

/**
 * 【1回だけ実行】担当列の表記に合わせる最終リネーム（2026-08-01 確定分）
 * ---------------------------------------------------------------------
 * 読み仮名の照合で確定した5人分を、対応表のK列に反映する。
 *   ・新しい名前の行が既にあれば何もしない（IDだけ補完）
 *   ・古い名前の行があればK列を書き換え（IDはそのまま）
 *   ・どちらも無ければ行を追加
 * 何度実行しても安全。実行後は alignMentionTable で確認する。
 */
var FINAL_RENAMES = [
  // [今の名前, 担当列の表記, DiscordユーザーID]
  ['HIDEKI ZUNIGA', 'ヒデキ',   '1287282022896762930'],
  ['ASATO YUICHI',  '安里',     '1233366519057219594'],
  ['野村友作',       'ゆうさく', '763212900319756328'],
  ['waryo',         'かずあき', '1320242584005640257'],  // BO確認済み
  ['Yuki Hirano',   'ゆき',     '825192029964599297']    // BO確認済み
];

function applyFinalRenames(){
  var ss = SpreadsheetApp.openById(AUTO_TARGET_ID);
  var sh = ss.getSheetByName(CFG_SHEET);
  if (!sh) throw new Error('「' + CFG_SHEET + '」シートがありません');
  sh.getRange('L2:L100').setNumberFormat('@'); // IDの数値化け防止

  var names = sh.getRange('K2:K100').getValues().map(function(r){ return String(r[0]).trim(); });
  var done = [];
  FINAL_RENAMES.forEach(function(p){
    var idx = names.indexOf(p[1]); // 新しい名前が既にある？
    if (idx >= 0){
      done.push(p[1] + '（変更済みだったのでIDだけ確認）');
    } else {
      idx = names.indexOf(p[0]);   // 古い名前の行を探す
      if (idx >= 0){
        sh.getRange(idx + 2, 11).setValue(p[1]);
        names[idx] = p[1];
        done.push(p[0] + ' → ' + p[1]);
      } else {                     // どちらも無ければ追加
        idx = names.indexOf('');
        if (idx < 0) throw new Error('K列に空きがありません');
        sh.getRange(idx + 2, 11).setValue(p[1]);
        names[idx] = p[1];
        done.push(p[1] + '（行が無かったので追加）');
      }
    }
    sh.getRange(idx + 2, 12).setValue(p[2]); // IDを念のため上書き（正しい値で統一）
  });
  Logger.log('完了:\n' + done.join('\n') +
    '\n\n次は alignMentionTable を実行して、「対応できず」が増田・岩渕・砂田だけになっているか確認してください。');
}

/**
 * 【1回だけ実行】対応表のK列をマスターの担当列の表記に自動で合わせる
 * ---------------------------------------------------------------------
 * マスターの制作担当（K列）・BO担当（R列）に実際に出てくる表記を集めて、
 * 対応表の名前と部分一致で照合する。
 *   ・候補が1人に絞れた → 対応表のK列をマスターの表記に書き換え（IDはそのまま）
 *   ・候補が0人 or 複数 → 何もせずログに出す（手で対応する用）
 * 実行後は必ず実行ログを確認すること。
 */
function alignMentionTable(){
  var ss = SpreadsheetApp.openById(AUTO_TARGET_ID);
  var ms = ss.getSheetByName('マスター');
  var sh = ss.getSheetByName(CFG_SHEET);
  if (!ms || !sh) throw new Error('マスター／' + CFG_SHEET + ' シートが見つかりません');

  // 1) マスターに実際に出てくる担当表記を集める（制作担当K=11・BO担当R=18）
  var rowsN = Math.min(1000, ms.getMaxRows() - 1);
  var vals = ms.getRange(2, 1, rowsN, 18).getValues();
  var used = {};
  vals.forEach(function(r){
    if (!String(r[0]).trim()) return; // クライアント名が空の行はスキップ
    [String(r[10]).trim(), String(r[17]).trim()].forEach(function(n){
      if (n) used[n] = (used[n] || 0) + 1;
    });
  });

  // 2) 対応表の現在の名前
  var tableNames = sh.getRange('K2:K100').getValues().map(function(r){ return String(r[0]).trim(); });

  // 空白（半角・全角）を除いて小文字化した比較用の形
  function norm(s){ return String(s).replace(/[\s　]+/g, '').toLowerCase(); }

  var renamed = [], unmatched = [], already = [];
  Object.keys(used).forEach(function(masterName){
    if (tableNames.indexOf(masterName) >= 0){ already.push(masterName); return; } // 完全一致あり
    var nm = norm(masterName);
    var hits = [];
    tableNames.forEach(function(t, i){
      if (!t) return;
      var nt = norm(t);
      if (nt.indexOf(nm) >= 0 || nm.indexOf(nt) >= 0) hits.push(i);
    });
    if (hits.length === 1){
      var i = hits[0];
      renamed.push('「' + tableNames[i] + '」→「' + masterName + '」');
      sh.getRange(i + 2, 11).setValue(masterName);
      tableNames[i] = masterName;
    } else {
      unmatched.push(masterName + '（担当' + used[masterName] + '件・候補' + hits.length + '人）');
    }
  });

  Logger.log(
    '完全一致済み: ' + (already.join(', ') || 'なし') + '\n' +
    '自動リネーム: ' + (renamed.join(' / ') || 'なし') + '\n' +
    '対応できず（手で確認してください）: ' + (unmatched.join(', ') || 'なし') + '\n\n' +
    '※増田・岩渕・砂田は対応表から削除済みのため「対応できず」に出ますが、\n' +
    '  BOロールメンションで拾う運用にしたので、そのままで問題ありません。');
}

function fillMentionTable(){
  var ss = SpreadsheetApp.openById(AUTO_TARGET_ID);
  var sh = ss.getSheetByName(CFG_SHEET);
  if (!sh) throw new Error('「' + CFG_SHEET + '」シートがありません。先に setupProgressReminder を実行してください。');

  // IDの数値化け防止（18桁超は数値だと末尾が丸まる）
  sh.getRange('L2:L100').setNumberFormat('@');

  var names = sh.getRange('K2:K100').getValues().map(function(r){ return String(r[0]).trim(); });
  var updated = 0, added = 0;
  MENTION_ROWS.forEach(function(row){
    var idx = names.indexOf(row[0]);
    if (idx < 0){
      idx = names.indexOf('');
      if (idx < 0) throw new Error('K列に空きがありません（K2:K100）');
      names[idx] = row[0];
      sh.getRange(idx + 2, 11).setValue(row[0]);
      added++;
    } else {
      updated++;
    }
    if (row[1]) sh.getRange(idx + 2, 12).setValue(row[1]);
  });

  Logger.log('対応表を入力しました（既存行の更新 ' + updated + '件 / 追加 ' + added + '件）。\n' +
    'やること残り:\n' +
    '1. 「岩渕」のIDが空欄です。「小渕」か「Yuika.I」のどちらが岩渕さんか確認して埋めてください。\n' +
    '2. 担当列の表記とDiscord表示名が違う人は、K列を担当列の表記に書き換えてください。\n' +
    '3. 担当列に出てこない人の行は消してもOKです。');
}
