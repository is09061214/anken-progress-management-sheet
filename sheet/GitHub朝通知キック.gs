/**
 * GitHub朝通知キック.gs — 朝の通知をGitHub任せにせず、Apps Scriptから時間ぴったりに起動する
 *
 * ■ これは何？
 *   「Discord 朝の進捗通知」はGitHub Actionsの定時実行（cron）で動かしていたが、
 *   2026年8月末からGitHub側のスケジューラ遅延がひどくなり、朝6:30起動のはずが
 *   昼過ぎに届く日も出てきた（8/28は8時間遅延）。
 *   そこで発想を変えて、時間に正確なApps Scriptの時間トリガー（毎朝6:30頃）から
 *   GitHubのAPIを叩いて「今すぐ実行して」と命令を飛ばす方式にする。
 *   命令を受けた実行（workflow_dispatch）は待ち行列に並ばず数秒で始まるので、
 *   7:00までにはほぼ確実に届くようになる。
 *
 * ■ 動き
 *   ・毎朝6:15〜6:45頃：kickMorningNotify が起動
 *       → 「Discord 朝の進捗通知」を即時実行
 *       → 月曜だけ「撮影予定 月曜朝通知」も一緒に実行
 *   ・失敗したら45秒待って自動リトライ（最大3回）。
 *     3回とも失敗したときだけDiscordに通知（通知なし＝正常）
 *   ・Webhook URLは「自動追加設定」シートの「Discord Webhook URL」欄を共用
 *
 * ■ 使い方（初回セットアップ・約10分）
 *   1. GitHubでアクセストークンを発行する
 *      github.com → 右上アイコン → Settings → Developer settings
 *      → Personal access tokens → Fine-grained tokens → Generate new token
 *        ・Token name: anken-morning-kick（何でもOK）
 *        ・Expiration: 1年など長めに（切れたらDiscordにエラー通知が来るので気づけます）
 *        ・Repository access: Only select repositories → anken-progress-management-sheet
 *        ・Permissions → Repository permissions → Actions を「Read and write」に
 *      → 発行された「github_pat_」で始まる文字列をコピー
 *   2. シートの 拡張機能 → Apps Script にこのファイルの中身を貼り付ける
 *      ※「案件シート撮影日追加自動化」と同じプロジェクトに！
 *   3. Apps Script画面の左メニュー ⚙️「プロジェクトの設定」→ 一番下の
 *      「スクリプト プロパティ」→「スクリプト プロパティを追加」
 *        ・プロパティ: GITHUB_TOKEN
 *        ・値: さっきコピーしたトークン
 *   4. 関数「testGithubToken」を実行 → ログに「接続OK」と出ればトークン設定は成功
 *   5. 関数「setupGithubKickTrigger」を実行 → 毎朝のトリガーが登録される
 *
 * ■ 知っておいてほしいこと
 *   ・GitHub側のcron（定時実行）はこの方式への切替と同時に外してあるので、
 *     二重に通知が飛ぶことはありません
 *   ・「今すぐ通知を飛ばしたい」ときは runMorningNotifyNow を実行すればOK
 *     （本物の通知が@everyone付きで飛ぶので、テスト連打には注意）
 */

var GHKICK_OWNER = 'is09061214';
var GHKICK_REPO  = 'anken-progress-management-sheet';
var GHKICK_BRANCH = 'main';
var GHKICK_MORNING_WORKFLOW = 'discord_morning.yml'; // Discord 朝の進捗通知
var GHKICK_MONDAY_WORKFLOW  = 'film_monday.yml';     // 撮影予定 月曜朝通知（月曜のみ）

var GHKICK_SS_ID     = '1oXwpCyw_g4GT1SFh4V_DW6RL7Q8h5FhIx9nbEV_eLXo'; // iMuse案件進捗管理シート_V3
var GHKICK_CFG_SHEET = '自動追加設定'; // Discord Webhook URL をここから読む（既存と共用）

var GHKICK_MAX_TRIES      = 3;  // 最初の実行＋リトライ2回
var GHKICK_RETRY_WAIT_SEC = 45; // リトライ前に待つ秒数

/* ==================== 初回セットアップ ==================== */
function setupGithubKickTrigger(){
  // 既存の同名トリガーを一旦削除（二重実行防止）
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'kickMorningNotify') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('kickMorningNotify')
    .timeBased().everyDays(1).atHour(6).nearMinute(30).create();
  Logger.log('セットアップ完了。毎朝6:15〜6:45頃にGitHubの朝通知を起動します。\n' +
    '月曜は撮影予定の通知も一緒に起動します。エラーのときだけDiscordに通知が飛びます。');
}

/* ==================== 毎朝動く本体（トリガーはこちらを呼ぶ） ==================== */
function kickMorningNotify(){
  var jobs = [{ file: GHKICK_MORNING_WORKFLOW, label: 'Discord 朝の進捗通知' }];
  if (Utilities.formatDate(new Date(), 'Asia/Tokyo', 'u') === '1'){ // 1 = 月曜
    jobs.push({ file: GHKICK_MONDAY_WORKFLOW, label: '撮影予定 月曜朝通知' });
  }
  jobs.forEach(function(j){ ghkickRunWithRetry_(j.file, j.label); });
}

/* ==================== 手動実行用 ==================== */
// 今すぐ朝の進捗通知を飛ばす（本物の@everyone通知が飛びます）
function runMorningNotifyNow(){
  ghkickRunWithRetry_(GHKICK_MORNING_WORKFLOW, 'Discord 朝の進捗通知（手動）');
}

// トークンの動作確認（通知は飛ばない・安全）
function testGithubToken(){
  var token = ghkickReadToken_();
  if (!token){
    Logger.log('スクリプト プロパティ「GITHUB_TOKEN」が未設定です。使い方の手順3を確認してください。');
    return;
  }
  var res = UrlFetchApp.fetch(
    'https://api.github.com/repos/' + GHKICK_OWNER + '/' + GHKICK_REPO + '/actions/workflows',
    { headers: ghkickHeaders_(token), muteHttpExceptions: true }
  );
  if (res.getResponseCode() === 200){
    var names = JSON.parse(res.getContentText()).workflows.map(function(w){ return w.name; });
    Logger.log('接続OK！ トークンは正しく設定されています。\n見えているワークフロー: ' + names.join(' / '));
  } else {
    Logger.log('接続できませんでした（HTTP ' + res.getResponseCode() + '）。\n' +
      'トークンの権限（Actions: Read and write）と対象リポジトリの設定を確認してください。\n' +
      res.getContentText().slice(0, 300));
  }
}

/* ==================== 共通処理 ==================== */
function ghkickRunWithRetry_(workflowFile, label){
  var lastErr = null;
  for (var attempt = 1; attempt <= GHKICK_MAX_TRIES; attempt++){
    try {
      ghkickDispatch_(workflowFile);
      if (attempt > 1){
        ghkickPostDiscord_('🔁 **リトライで起動できました**\n' +
          '・処理：' + label + '\n' +
          '・' + attempt + '回目で成功しました\n' +
          '・日時：' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'M/d HH:mm'));
      }
      Logger.log('起動成功（' + attempt + '回目）：' + label);
      return;
    } catch (e) {
      lastErr = e;
      Logger.log('起動失敗（' + attempt + '回目）：' + label + ' → ' + (e && e.message ? e.message : String(e)));
      if (attempt < GHKICK_MAX_TRIES) Utilities.sleep(GHKICK_RETRY_WAIT_SEC * 1000);
    }
  }
  var msg = '⚠️ **朝の通知をGitHubで起動できませんでした**\n' +
    '・処理：' + label + '\n' +
    '・エラー：' + (lastErr && lastErr.message ? lastErr.message : String(lastErr)) + '\n' +
    '・試行：' + GHKICK_MAX_TRIES + '回実行してすべて失敗\n' +
    '・日時：' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'M/d HH:mm') + '\n' +
    '※トークンの期限切れの可能性もあります。Apps Scriptの「testGithubToken」で確認できます。';
  var sent = ghkickPostDiscord_(msg);
  if (!sent) throw lastErr; // Discordに送れなかったときだけ、Gmailに知らせる
}

// GitHubに「このワークフローを今すぐ実行して」と命令する
function ghkickDispatch_(workflowFile){
  var token = ghkickReadToken_();
  if (!token) throw new Error('スクリプト プロパティ「GITHUB_TOKEN」が未設定です');
  var url = 'https://api.github.com/repos/' + GHKICK_OWNER + '/' + GHKICK_REPO +
    '/actions/workflows/' + workflowFile + '/dispatches';
  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: ghkickHeaders_(token),
    contentType: 'application/json',
    payload: JSON.stringify({ ref: GHKICK_BRANCH }),
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  if (code !== 204){
    throw new Error('GitHub API エラー（HTTP ' + code + '）: ' + res.getContentText().slice(0, 200));
  }
}

function ghkickHeaders_(token){
  return {
    'Authorization': 'Bearer ' + token,
    'Accept': 'application/vnd.github+json',
    'X-GitHub-Api-Version': '2022-11-28'
  };
}

function ghkickReadToken_(){
  return (PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN') || '').trim();
}

// 「自動追加設定」シートから Discord Webhook URL を読んで投稿（既存のエラー通知と同じ仕組み）
function ghkickPostDiscord_(message){
  try {
    var sh = SpreadsheetApp.openById(GHKICK_SS_ID).getSheetByName(GHKICK_CFG_SHEET);
    if (!sh) return false;
    var webhook = '';
    var vals = sh.getRange('H1:I20').getValues();
    for (var i = 0; i < vals.length; i++){
      if (String(vals[i][0]).trim() === 'Discord Webhook URL'){ webhook = String(vals[i][1]).trim(); break; }
    }
    if (!webhook) return false;
    var res = UrlFetchApp.fetch(webhook, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ content: message }),
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    return code >= 200 && code < 300;
  } catch (e) {
    return false;
  }
}
