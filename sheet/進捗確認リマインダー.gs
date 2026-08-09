/**
 * 【試作版V3対応】制作進捗リマインダー（BOリクエスト・2026-07-30 グリル合意）
 * =====================================================================
 * 制作締切が近い／過ぎているのに制作が始まっていない案件を、毎朝Discordの
 * 専用チャンネルに通知して担当者を名指しで催促する。
 *
 * ■ 合意した設計（2026-07-30）
 *   ・対象: ステータスが「未着手」または「タイトル待ち」の行
 *       未着手     → 制作担当（K列）を@メンションして催促
 *       タイトル待ち → BO担当（R列）宛ての文面（タイトル確定を急いで）
 *   ・タイミング: 制作締切（L列）の前日から。当日・超過後もステータスが
 *     進むまで毎朝通知し続ける（⚠明日締切 → 🚨本日締切 → 🔥締切超過）
 *   ・通知先: 進捗確認専用チャンネルのWebhook（自動追加通知とは別）
 *   ・チーム担当（GS等）: 対応表のID欄に「・」区切りで複数ID → 全員メンション
 *   ・全通知の冒頭でBOロール（@BACKOFFICE）をメンション（ロールIDを設定した場合）
 *   ・トリガー: 既存の毎朝6時トリガーに相乗り
 *     （autoAddFromCalendar の最後から sendProgressReminders_ を呼ぶ）
 *   ・土日も含むカレンダー日ベース
 *
 * ■ 前提
 *   このファイルは「カレンダー自動追加スクリプト」と同じApps Script
 *   プロジェクトに入れること（AUTO_TARGET_ID・日付ユーティリティ・
 *   appendLog_ を共用しているため、単体では動きません）。
 *
 * ■ 使い方（初回セットアップ）
 *   1. このファイルをApps Scriptプロジェクトに追加
 *   2. カレンダー自動追加スクリプト側の autoAddFromCalendar 末尾に
 *      リマインダー呼び出しが入っていることを確認（今回の更新で追加済み）
 *   3. 関数「setupProgressReminder」を実行
 *      → 「自動追加設定」シートに設定行と「名前→Discord ID対応表」を追加
 *   4. Discordの専用チャンネルでWebhookを発行し、
 *      「進捗確認 Discord Webhook URL」の隣（I列）に貼る
 *   5. 対応表（K:L列）に編集者・BOの名前とDiscordユーザーIDを記入
 *      （IDが空の人はメンションなしで名前のテキスト表記になる）
 *   6. 動作確認は「testProgressReminder」を実行
 *      （送らずに内容だけ確認したいときは「dryRunProgressReminder」）
 */

/* ==================== 初回セットアップ ==================== */
function setupProgressReminder(){
  var ss = SpreadsheetApp.openById(AUTO_TARGET_ID);
  var sh = ss.getSheetByName(CFG_SHEET);
  if (!sh) throw new Error('「' + CFG_SHEET + '」シートがありません。先に setupCalendarAutoAdd を実行してください。');

  // 1) H列の動作設定に3行追加（既にあれば何もしない）
  addCfgRow_(sh, '進捗確認 Discord Webhook URL', '');
  addCfgRow_(sh, '超過通知の上限（日）', 30);
  addCfgRow_(sh, 'BOロールのID（全通知に@）', '');

  // 2) K:L列に「名前→Discord ID」対応表（既にあれば何もしない）
  if (String(sh.getRange('K1').getValue()).trim() === ''){
    sh.getRange('K1:L1').setValues([['メンバー名（担当列の表記）','DiscordユーザーID']])
      .setFontWeight('bold').setBackground('#e8f0e8');
    // BO担当は既定値から分かっているので名前だけ先に入れておく（IDはBOに記入してもらう）
    sh.getRange('K2:K4').setValues([['増田'],['岩渕'],['砂田']]);
    sh.setColumnWidth(11, 190); sh.setColumnWidth(12, 190);
  }
  // メモは毎回上書き（書き方の説明を最新にするため）
  sh.getRange('K1').setNote(
    '進捗リマインダーのメンション用対応表。\n' +
    '・メンバー名 … マスターの制作担当（K列）・BO担当（R列）と同じ表記で書く\n' +
    '・DiscordユーザーID … Discordで開発者モードをON→ユーザーを右クリック→「ユーザーIDをコピー」の数字\n' +
    '・チーム名の行（GS等）は、IDを「・」区切りで複数書くと全員メンションされます\n' +
    '・IDが空の人はメンションされず、名前のテキスト表記で通知されます');
  sh.getRange('H2:H30').getValues().forEach(function(r, i){
    if (String(r[0]).trim() === 'BOロールのID（全通知に@）'){
      sh.getRange(i + 2, 8).setNote(
        '毎回の通知の冒頭で@メンションするロール（@BACKOFFICE等）のID。\n' +
        'Discordで開発者モードをON→サーバー設定→ロール→対象ロールを右クリック→「IDをコピー」の数字。\n' +
        '空欄ならロールメンションなし。\n' +
        '※通知が鳴らない場合は、ロール設定の「全員にこのロールへの@mentionを許可する」をONにしてください。');
    }
  });

  // 3) トリガーは既存の毎朝6時（autoAddFromCalendar／エラー通知付きのsafe版）に相乗りするため新設しない
  var hasTrigger = ScriptApp.getProjectTriggers().some(function(t){
    var h = t.getHandlerFunction();
    return h === 'autoAddFromCalendar' || h === 'safeAutoAddFromCalendar';
  });
  Logger.log('セットアップ完了。\n' +
    '・「自動追加設定」シートのI列に進捗確認用WebhookのURLを貼ってください。\n' +
    '・K:L列の対応表に編集者の名前とDiscordユーザーIDを記入してください。\n' +
    (hasTrigger ? '・毎朝6時の既存トリガーに相乗りして動きます。'
                : '⚠ 毎朝6時のトリガーが見つかりません。先に setupCalendarAutoAdd を実行してください。'));
}

// H列の動作設定にラベル行を追加（既にあればスキップ）
function addCfgRow_(sh, label, defVal){
  var vals = sh.getRange('H2:H30').getValues();
  var firstEmpty = -1;
  for (var i = 0; i < vals.length; i++){
    var v = String(vals[i][0]).trim();
    if (v === label) return; // 追加済み
    if (v === '' && firstEmpty < 0) firstEmpty = i + 2;
  }
  if (firstEmpty < 0) throw new Error('自動追加設定シートのH列に空きがありません（H2:H30）');
  sh.getRange(firstEmpty, 8, 1, 2).setValues([[label, defVal]]);
}

/* ==================== 設定の読み込み ==================== */
function readReminderConfig_(ss){
  var sh = ss.getSheetByName(CFG_SHEET);
  if (!sh) return null;
  var scal = {};
  sh.getRange('H2:I30').getValues().forEach(function(r){ if (r[0]) scal[String(r[0])] = r[1]; });
  var mentions = {};
  sh.getRange('K2:L60').getValues().forEach(function(r){
    var name = String(r[0]).trim();
    var id   = String(r[1]).trim();
    if (name) mentions[name] = id;
  });
  return {
    webhook:        String(scal['進捗確認 Discord Webhook URL'] || '').trim(),
    overdueLimit:   Number(scal['超過通知の上限（日）']) || 30,
    boRole:         String(scal['BOロールのID（全通知に@）'] || '').replace(/\D/g, ''), // 数字だけ抽出
    mentions:       mentions
  };
}

// 名前 → Discordメンション文字列。
// ID欄に「・」区切りで複数書けば全員メンション（GS→田中・舟瀬の2人等）。
// 有効なIDが1つも無ければ名前＋さんのテキスト表記。
function mention_(name, mentions){
  name = String(name || '').trim();
  if (!name) return '担当未定';
  var ids = String(mentions[name] || '').split(/[・、,\/\s]+/).filter(function(s){
    return /^\d{5,}$/.test(s);
  });
  if (ids.length) return ids.map(function(id){ return '<@' + id + '>'; }).join(' ');
  return name + 'さん';
}

/* ==================== 毎朝の本体（autoAddFromCalendarから呼ばれる） ==================== */
function sendProgressReminders_(ss, lg){
  var cfg = readReminderConfig_(ss);
  if (!cfg || !cfg.webhook) return; // Webhook未設定なら何もしない（セットアップ前でも安全）
  var ms = ss.getSheetByName('マスター');
  if (!ms) return;

  var today = dateOnly_(new Date());
  var rowsN = Math.min(1000, ms.getMaxRows() - 1);
  // A クライアント(1) / D タイトル(4) / H ステータス(8) / K 制作担当(11) /
  // L 制作締切(12) / R BO担当(18)
  var vals = ms.getRange(2, 1, rowsN, 18).getValues();

  var items = [];
  for (var i = 0; i < vals.length; i++){
    var client = String(vals[i][0]).trim();
    if (!client) continue;
    var status = String(vals[i][7]).trim();
    if (status !== '未着手' && status !== 'タイトル待ち') continue;
    var due = vals[i][11];
    if (!(due instanceof Date)) continue;
    due = dateOnly_(due);
    var diff = Math.round((due - today) / 86400000); // 1=明日 0=本日 負=超過
    if (diff > 1) continue;                          // 締切前日より前はまだ通知しない
    if (diff < -cfg.overdueLimit) continue;          // 古すぎる超過は対象外（設定で変更可）
    items.push({
      row: i + 2, client: client,
      title: String(vals[i][3]).trim() || '（タイトル未定）',
      status: status, due: due, diff: diff,
      editor: String(vals[i][10]).trim(),  // K 制作担当
      bo:     String(vals[i][17]).trim()   // R BO担当
    });
  }
  if (!items.length) return; // 通知ゼロの日は何も送らない

  // 深刻な順（超過が大きい順 → 本日 → 明日）
  items.sort(function(a, b){ return a.diff - b.diff; });

  var lines = items.map(function(it){
    var icon = it.diff < 0 ? '🔥' : (it.diff === 0 ? '🚨' : '⚠');
    var when = it.diff < 0 ? md_(it.due) + '（' + (-it.diff) + '日超過）'
             : it.diff === 0 ? '本日 ' + md_(it.due)
             : '明日 ' + md_(it.due);
    if (it.status === '未着手'){
      return icon + ' ' + mention_(it.editor, cfg.mentions) +
        ' 【' + it.client + '】' + it.title + ' — 制作締切 ' + when + ' でまだ未着手です。進捗どうですか？';
    }
    // タイトル待ち → BO宛て
    return icon + ' ' + mention_(it.bo, cfg.mentions) +
      ' 【' + it.client + '】' + it.title + ' — 制作締切 ' + when + ' なのにタイトルが未確定です。確定を急いでください！';
  });

  postReminderDiscord_(cfg, lines, lg);
  try { appendLog_(lg, '', '', '', '', '進捗リマインダー送信', items.length + '件通知'); } catch(e){}
}

/* ==================== Discord投稿（担当者メンション版） ==================== */
function postReminderDiscord_(cfg, lines, lg){
  var webhook = cfg.webhook;
  if (!webhook || !lines.length) return;
  var header = (cfg.boRole ? '<@&' + cfg.boRole + '>\n' : '') + // BOロールを全通知にメンション
    '📢 **制作進捗リマインダー**（' +
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'M/d HH:mm') + '）\n' +
    '🔗 シートを開く → https://docs.google.com/spreadsheets/d/' + AUTO_TARGET_ID + '/edit\n';
  var chunks = [], cur = header;
  lines.forEach(function(line){
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
          allowed_mentions: { parse: ['users', 'roles'] } // 名指し＋BOロールのメンションを通知として効かせる
        }),
        muteHttpExceptions: true
      });
    } catch(e){
      try { appendLog_(lg, '', '', '', '', '⚠進捗リマインダー通知失敗', String(e)); } catch(e2){}
    }
  });
}

/* ==================== 動作確認用 ==================== */
// 実際にDiscordへ送って確認する（毎朝の処理と同じものが1回動く）
function testProgressReminder(){
  var ss = SpreadsheetApp.openById(AUTO_TARGET_ID);
  var cfg = readReminderConfig_(ss);
  if (!cfg || !cfg.webhook){
    Logger.log('「自動追加設定」シートの「進捗確認 Discord Webhook URL」が空です。' +
      'H列のラベルの隣（I列）に専用チャンネルのWebhook URLを貼ってください。');
    return;
  }
  sendProgressReminders_(ss, ss.getSheetByName(LOG_SHEET));
  Logger.log('実行しました。対象があればDiscordの専用チャンネルに届いています（対象ゼロなら何も送られません）。');
}

// 送らずに「今日送るとしたら誰に何が飛ぶか」をログで確認する
function dryRunProgressReminder(){
  var ss  = SpreadsheetApp.openById(AUTO_TARGET_ID);
  var cfg = readReminderConfig_(ss) || { webhook: '', overdueLimit: 30, mentions: {} };
  var ms  = ss.getSheetByName('マスター');
  var today = dateOnly_(new Date());
  var rowsN = Math.min(1000, ms.getMaxRows() - 1);
  var vals = ms.getRange(2, 1, rowsN, 18).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++){
    var client = String(vals[i][0]).trim();
    if (!client) continue;
    var status = String(vals[i][7]).trim();
    if (status !== '未着手' && status !== 'タイトル待ち') continue;
    var due = vals[i][11];
    if (!(due instanceof Date)) continue;
    var diff = Math.round((dateOnly_(due) - today) / 86400000);
    if (diff > 1 || diff < -cfg.overdueLimit) continue;
    out.push('行' + (i + 2) + ': 【' + client + '】' + (String(vals[i][3]).trim() || '（タイトル未定）') +
      ' / ' + status + ' / 締切' + fmt_(dateOnly_(due)) + '（diff=' + diff + '）' +
      ' → ' + (status === '未着手' ? '制作担当 ' + (String(vals[i][10]).trim() || '未定')
                                   : 'BO担当 ' + (String(vals[i][17]).trim() || '未定')));
  }
  Logger.log(out.length ? '今日送るとしたら以下の' + out.length + '件:\n' + out.join('\n')
                        : '今日の対象は0件です（何も送られません）。');
  Logger.log('Webhook設定: ' + (cfg.webhook ? '設定済み' : '未設定（このままだと毎朝の本番でも送られません）'));
}
