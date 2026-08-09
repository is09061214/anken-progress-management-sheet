/**
 * 【毎朝の自動チェック】マスターの数式破損（#REF!）を検知してDiscordに通知
 * =====================================================================
 * 以前、マスター35行目の信号数式が #REF! に化けて「情報不足」扱いになる
 * 事故があったので（数式修復_マスター35行目.gs 参照）、同じ症状を毎朝
 * 自動でスキャンして、見つけたらDiscordに知らせる仕組みです。
 *
 * ■ チェック内容（マスターシート全体）
 *   1. 数式の中に #REF! が混ざっているセル（数式そのものが壊れているパターン）
 *   2. 計算結果が #REF! エラーになっているセル（参照先が消えたパターン）
 *
 * ■ 通知ルール
 *   ・異常が見つかったときだけDiscordに通知（通知が来ない＝正常）
 *   ・Webhook URLは「自動追加設定」シートの「Discord Webhook URL」欄を共用
 *
 * ■ 使い方（初回セットアップ）
 *   1. シートの 拡張機能 → Apps Script にこのファイルの中身を貼り付ける
 *      （カレンダー自動追加スクリプトと同じプロジェクトでOK。名前は衝突しません）
 *   2. 関数「setupRefErrorDailyScan」を選んで実行
 *      → 毎朝7時のトリガーを設置（6時のカレンダー自動追加が終わったあとに動く想定）
 *   3. 動作確認は「dailyRefErrorScan」を手動実行。ログに結果が出ます
 *      Discord通知のテストは「testRefScanNotify」を実行
 */

var REFSCAN_SS_ID     = '1oXwpCyw_g4GT1SFh4V_DW6RL7Q8h5FhIx9nbEV_eLXo'; // iMuse案件進捗管理シート_V3
var REFSCAN_SHEET     = 'マスター';
var REFSCAN_CFG_SHEET = '自動追加設定'; // Discord Webhook URL をここから読む
var REFSCAN_MAX_LIST  = 30;             // 通知に載せるセル数の上限（それ以上は「ほか◯件」）

/* ==================== 初回セットアップ ==================== */
function setupRefErrorDailyScan(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'dailyRefErrorScan') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyRefErrorScan').timeBased().everyDays(1).atHour(7).create();
  Logger.log('セットアップ完了。毎朝7時にマスターの数式チェックが動きます。\n' +
    '#REF! が見つかったときだけDiscordに通知します（通知なし＝正常）。');
}

/* ==================== 毎朝の本体 ==================== */
function dailyRefErrorScan(){
  var webhook = '';
  try {
    var ss = SpreadsheetApp.openById(REFSCAN_SS_ID);
    webhook = readRefScanWebhook_(ss);
    var ms = ss.getSheetByName(REFSCAN_SHEET);
    if (!ms) throw new Error('「' + REFSCAN_SHEET + '」シートが見つかりません');

    var range    = ms.getDataRange();
    var formulas = range.getFormulas();
    var values   = range.getDisplayValues();

    var brokenFormula = []; // 数式の中に #REF! が混入
    var brokenValue   = []; // 計算結果が #REF! エラー
    for (var r = 0; r < formulas.length; r++){
      for (var c = 0; c < formulas[r].length; c++){
        var addr = refScanColLetter_(c + 1) + (r + 1);
        if (formulas[r][c] && formulas[r][c].indexOf('#REF!') !== -1){
          brokenFormula.push(addr);
        } else if (String(values[r][c]).indexOf('#REF!') !== -1){
          brokenValue.push(addr);
        }
      }
    }

    if (!brokenFormula.length && !brokenValue.length){
      Logger.log('#REF!混入なし。マスターの数式は全部きれいです。');
      return;
    }

    var lines = ['🚨 案件進捗シートのマスターで数式破損（#REF!）を見つけました！'];
    if (brokenFormula.length){
      lines.push('🔧 数式そのものが壊れているセル（' + brokenFormula.length + '個）: ' +
        refScanJoin_(brokenFormula));
    }
    if (brokenValue.length){
      lines.push('⚠️ 計算結果がエラーになっているセル（' + brokenValue.length + '個）: ' +
        refScanJoin_(brokenValue));
    }
    lines.push('直し方: 正常な行の同じ列のセルをコピーして、壊れたセルに「特殊貼り付け → 数式のみ貼り付け」。' +
      'B列・V列なら fixRow35RefError と同じ要領です。');

    Logger.log(lines.join('\n'));
    postRefScanDiscord_(webhook, lines.join('\n'));

  } catch (e) {
    Logger.log('数式チェック自体が失敗: ' + e);
    postRefScanDiscord_(webhook,
      '🚨 毎朝の数式チェック（dailyRefErrorScan）の実行中にエラーが出ました: ' + e + '\n' +
      'Apps Scriptの実行ログを確認してください。');
  }
}

/* ==================== Discord通知 ==================== */
// Webhook設定の動作確認用（実行してDiscordにテスト通知が届けばOK）
function testRefScanNotify(){
  var ss = SpreadsheetApp.openById(REFSCAN_SS_ID);
  var webhook = readRefScanWebhook_(ss);
  if (!webhook){
    Logger.log('「' + REFSCAN_CFG_SHEET + '」シートの「Discord Webhook URL」が空です。' +
      'ラベルの隣のセルにWebhookのURLを貼ってください。');
    return;
  }
  postRefScanDiscord_(webhook, '✅ テスト通知です。数式破損（#REF!）を見つけたら、この仕組みでお知らせします！');
  Logger.log('テスト通知を送りました。Discordのチャンネルを確認してください。');
}

// 「自動追加設定」シートから Discord Webhook URL を読む（カレンダー自動追加と共用）
function readRefScanWebhook_(ss){
  var sh = ss.getSheetByName(REFSCAN_CFG_SHEET);
  if (!sh) return '';
  var vals = sh.getRange('H1:I20').getValues();
  for (var i = 0; i < vals.length; i++){
    if (String(vals[i][0]).trim() === 'Discord Webhook URL') return String(vals[i][1]).trim();
  }
  return '';
}

function postRefScanDiscord_(webhook, message){
  if (!webhook || !message) return;
  try {
    UrlFetchApp.fetch(webhook, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ content: message.substring(0, 1900) }), // Discordの2000字制限対策
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log('Discord通知に失敗: ' + e);
  }
}

/* ==================== 小物 ==================== */
// 30個を超えるセルは「ほか◯件」でまとめる（通知が長くなりすぎないように）
function refScanJoin_(cells){
  if (cells.length <= REFSCAN_MAX_LIST) return cells.join(', ');
  return cells.slice(0, REFSCAN_MAX_LIST).join(', ') + ' …ほか' + (cells.length - REFSCAN_MAX_LIST) + '件';
}

// 列番号 → 列文字（1→A, 27→AA）
function refScanColLetter_(n){
  var s = '';
  while (n > 0){
    var m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
