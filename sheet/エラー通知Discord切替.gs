/**
 * エラー通知Discord切替.gs — 朝の自動実行が失敗したらGmailではなくDiscordに知らせる
 *
 * ■ これは何？
 *   Apps Scriptの失敗通知メール（noreply-apps-scripts-notifications@google.com）は
 *   送り先をDiscordに変えられない仕様。なので発想を変えて、
 *   「毎朝の処理をtry/catchで包んだラッパー関数」経由で動かすようにする。
 *   エラーが起きたらDiscordに投稿 → Googleからは成功扱いになり、Gmailは来なくなる。
 *
 * ■ 対象（朝の自動実行3つ）
 *   ・朝5時台 rebuildSignalFormulas … 信号数式の一括再構築
 *   ・朝6時台 autoAddFromCalendar  … カレンダーから撮影日を自動追加
 *   ・朝7時台 dailyRefErrorScan    … #REF!破損チェック
 *
 * ■ 通知ルール
 *   ・エラーが起きたら45秒待って自動リトライ（最大3回実行）。
 *     Google側の一時的なタイムアウトはこれでほぼ吸収できる
 *   ・3回とも失敗したときだけDiscordに通知（通知なし＝正常）
 *   ・リトライで復旧したときは「🔁復旧しました」と軽くお知らせ
 *     （混雑が続いているかどうかの目安になるため）
 *   ・Webhook URLは「自動追加設定」シートの「Discord Webhook URL」欄を共用
 *   ・Discordへの投稿自体に失敗したときだけ、保険として従来どおりGmailに届く
 *
 * ■ 使い方（初回セットアップ）
 *   1. シートの 拡張機能 → Apps Script にこのファイルの中身を貼り付ける
 *      ※必ず「案件シート撮影日追加自動化」と同じプロジェクトに！
 *        （ラッパーが元の関数を呼び出すため。別プロジェクトだと動きません）
 *   2. 関数「setupErrorNotifyTriggers」を選んで実行
 *      → 既存の朝トリガー3つを、ラッパー経由のトリガーに置き換えます
 *   3. 動作確認は「testErrorNotifyDiscord」を実行
 *      → Discordにテスト通知が届けばOK
 *
 * ■ 知っておいてほしいこと
 *   ・スクリプトが6分の実行時間上限で強制終了した場合だけは、try/catchで
 *     捕まえられないので従来どおりGmailに届きます（めったに起きません）
 */

var ERRNOTIFY_SS_ID     = '1oXwpCyw_g4GT1SFh4V_DW6RL7Q8h5FhIx9nbEV_eLXo'; // iMuse案件進捗管理シート_V3
var ERRNOTIFY_CFG_SHEET = '自動追加設定'; // Discord Webhook URL をここから読む（既存と共用）

/* ==================== 初回セットアップ ==================== */
function setupErrorNotifyTriggers(){
  var jobs = [
    { original: 'rebuildSignalFormulas', wrapper: 'safeRebuildSignalFormulas', hour: 5 },
    { original: 'autoAddFromCalendar',   wrapper: 'safeAutoAddFromCalendar',   hour: 6 },
    { original: 'dailyRefErrorScan',     wrapper: 'safeDailyRefErrorScan',     hour: 7 }
  ];

  // 元の関数・ラッパー、どちらの既存トリガーも一旦ぜんぶ削除（二重実行防止）
  var targets = {};
  jobs.forEach(function(j){ targets[j.original] = true; targets[j.wrapper] = true; });
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (targets[t.getHandlerFunction()]) ScriptApp.deleteTrigger(t);
  });

  // ラッパー経由のトリガーを同じ時間帯に作り直す
  // （このプロジェクトに存在する関数だけが対象。無いものは自動でスキップ）
  var made = [], skipped = [];
  jobs.forEach(function(j){
    if (typeof globalThis[j.original] !== 'function'){
      skipped.push(j.original);
      return;
    }
    ScriptApp.newTrigger(j.wrapper).timeBased().everyDays(1).atHour(j.hour).create();
    made.push('朝' + j.hour + '時台：' + j.original);
  });

  Logger.log('セットアップ完了。Discordエラー通知つきに切り替えました。\n' +
    '・切り替えた処理：\n　' + (made.join('\n　') || 'なし') + '\n' +
    (skipped.length ? '・スキップ（このプロジェクトに関数が無いため）：' + skipped.join('、') + '\n' : '') +
    'エラーが起きたときだけDiscordに通知が飛びます（通知なし＝正常）。');
}

/* ==================== 毎朝動くラッパー（トリガーはこちらを呼ぶ） ==================== */
function safeRebuildSignalFormulas(){
  errNotifyRun_('rebuildSignalFormulas', '信号数式の一括再構築（朝5時台）');
}
function safeAutoAddFromCalendar(){
  errNotifyRun_('autoAddFromCalendar', 'カレンダーから撮影日を自動追加（朝6時台）');
}
function safeDailyRefErrorScan(){
  errNotifyRun_('dailyRefErrorScan', '#REF!破損チェック（朝7時台）');
}

/* ==================== 共通処理 ==================== */
var ERRNOTIFY_MAX_TRIES      = 3;               // 最初の実行＋リトライ2回
var ERRNOTIFY_RETRY_WAIT_SEC = 45;              // リトライ前に待つ秒数
var ERRNOTIFY_RETRY_DEADLINE = 4 * 60 * 1000;   // 開始から4分を過ぎたらリトライしない（6分の実行上限対策）

function errNotifyRun_(name, label){
  var start = Date.now();
  var attempt = 0;
  var lastErr = null;

  while (attempt < ERRNOTIFY_MAX_TRIES){
    attempt++;
    try {
      globalThis[name]();
      if (attempt > 1){
        errNotifyPostDiscord_('🔁 **リトライで復旧しました**\n' +
          '・処理：' + label + '\n' +
          '・' + attempt + '回目の実行で成功しました（Google側の一時的な混雑だったようです）\n' +
          '・日時：' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'M/d HH:mm'));
      }
      Logger.log('正常終了（' + attempt + '回目の実行で成功）：' + name);
      return;
    } catch (e) {
      lastErr = e;
      Logger.log('失敗（' + attempt + '回目）：' + name + ' → ' + (e && e.message ? e.message : String(e)));
      if (attempt >= ERRNOTIFY_MAX_TRIES) break;
      if (Date.now() - start > ERRNOTIFY_RETRY_DEADLINE){
        Logger.log('実行時間の上限が近いためリトライを打ち切ります：' + name);
        break;
      }
      Utilities.sleep(ERRNOTIFY_RETRY_WAIT_SEC * 1000);
    }
  }

  var msg = '⚠️ **朝の自動処理でエラーが起きました**\n' +
    '・処理：' + label + '\n' +
    '・関数：' + name + '\n' +
    '・エラー：' + (lastErr && lastErr.message ? lastErr.message : String(lastErr)) + '\n' +
    '・試行：' + attempt + '回実行してすべて失敗\n' +
    '・日時：' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'M/d HH:mm') + '\n' +
    '※リトライしてもダメだったケースです。連日続くようなら要確認です。';
  var sent = errNotifyPostDiscord_(msg);
  if (!sent) throw lastErr; // Discordに送れなかったときだけ、従来どおりGmailに知らせる
  Logger.log('エラーをDiscordに通知しました：' + name);
}

// 「自動追加設定」シートから Discord Webhook URL を読む（カレンダー自動追加と共用）
function errNotifyReadWebhook_(){
  var ss = SpreadsheetApp.openById(ERRNOTIFY_SS_ID);
  var sh = ss.getSheetByName(ERRNOTIFY_CFG_SHEET);
  if (!sh) return '';
  var vals = sh.getRange('H1:I20').getValues();
  for (var i = 0; i < vals.length; i++){
    if (String(vals[i][0]).trim() === 'Discord Webhook URL') return String(vals[i][1]).trim();
  }
  return '';
}

// Discordに投稿。成功したら true / 失敗したら false を返す
function errNotifyPostDiscord_(message){
  try {
    var webhook = errNotifyReadWebhook_();
    if (!webhook) return false;
    var res = UrlFetchApp.fetch(webhook, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ content: message }),
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    if (code < 200 || code >= 300){
      Logger.log('Discord送信エラー（HTTP ' + code + '）：' + res.getContentText().slice(0, 200));
      return false;
    }
    return true;
  } catch (e2) {
    return false;
  }
}

/* ==================== 動作確認用 ==================== */
function testErrorNotifyDiscord(){
  var ok = errNotifyPostDiscord_('✅ テスト通知です。朝の自動処理（数式再構築・カレンダー自動追加・#REF!チェック）が' +
    'エラーになったら、Gmailの代わりにこのチャンネルへお知らせします！');
  Logger.log(ok
    ? 'テスト通知を送りました。Discordのチャンネルを確認してください。'
    : '送信できませんでした。「自動追加設定」シートの「Discord Webhook URL」を確認してください。');
}
