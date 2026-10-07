/**
 * ステータス「誤字脱字チェック待ち」を追加するスクリプト（1回実行すればOK）
 *
 * 【やること】
 *   1. 設定シートの「ステータス定義表」(S:U) の「編集中」のすぐ下に1行追加
 *        ステータス=誤字脱字チェック待ち / 工程=サムネ / ボール=BO担当
 *      （S:U の3列だけを下にずらすので、同じシートの他の表には影響しない）
 *      → 現在担当・現在締切・🚦信号の数式は定義表を見ているので自動で追従
 *   2. マスターH列のプルダウンに選択肢を追加（範囲参照なら追従を確認、リスト型なら追記）
 *   3. 条件付き書式（チップ色・行のボール色）をコピー
 *        BO担当の既存ステータス（アップロード待ち等）のルールを複製して、
 *        ステータス名だけ差し替える → 同じ緑で塗られる
 *
 * 【安全のための工夫】
 *   ・何度実行しても重複しない（追加済みの項目はスキップ）
 *   ・案件データ（マスターの各行の値）には一切触らない
 *   ・previewAddTypoCheckStatus で「何をするか」だけ先に確認できる（シートは変更しない）
 *
 * 【使い方】
 *   1. シートの 拡張機能 → Apps Script を開く
 *   2. 左の「ファイル」の＋ → スクリプト → 名前を付けて、このファイルの中身をまるごと貼り付けて保存
 *   3. 関数 previewAddTypoCheckStatus を実行 → ログで内容を確認（任意）
 *   4. 関数 addTypoCheckStatus を実行 → ログに「完了」が出たらOK
 *   5. 動作確認が済んだら、このファイルは削除してかまいません
 */

var TYPO_SS_ID = '1oXwpCyw_g4GT1SFh4V_DW6RL7Q8h5FhIx9nbEV_eLXo'; // iMuse案件進捗管理シート_V3
var TYPO_NEW_STATUS = '誤字脱字チェック待ち';
var TYPO_AFTER_STATUS = '編集中';          // この行のすぐ下に入れる
var TYPO_DEADLINE_FROM = 'サムネ待ち';     // 締切（工程）はこのステータスと同じにする＝サムネ締切
var TYPO_BALL = 'BO担当';
var TYPO_STATUS_COL = 8;                   // マスターH列

/** 確認だけ（シートは変更しない） */
function previewAddTypoCheckStatus() {
  runTypoCheckStatus_(true);
}

/** 本番：追加を実行 */
function addTypoCheckStatus() {
  runTypoCheckStatus_(false);
}

function runTypoCheckStatus_(dryRun) {
  var ss = SpreadsheetApp.openById(TYPO_SS_ID);
  var st = ss.getSheetByName('設定');
  var ms = ss.getSheetByName('マスター');
  if (!st || !ms) throw new Error('「設定」または「マスター」シートが見つかりません。');
  var log = [dryRun ? '【確認モード：シートは変更していません】' : '【実行結果】'];

  // ---------- 1) ステータス定義表 ----------
  var table = readStatusTable_(st);
  var newRow = null;
  if (table.rowOf[TYPO_NEW_STATUS]) {
    newRow = table.rowOf[TYPO_NEW_STATUS];
    log.push('1) 定義表：すでに ' + newRow + '行目にあります（スキップ）');
  } else {
    var afterRow = table.rowOf[TYPO_AFTER_STATUS];
    if (!afterRow) throw new Error('定義表に「' + TYPO_AFTER_STATUS + '」が見つかりません。中止しました。');
    var stage = table.stageOf[TYPO_DEADLINE_FROM];
    if (!stage) throw new Error('定義表の「' + TYPO_DEADLINE_FROM + '」の工程が空です。中止しました。');
    newRow = afterRow + 1;
    if (!dryRun) {
      st.getRange(newRow, 19, 1, 3).insertCells(SpreadsheetApp.Dimension.ROWS); // S:U だけ下にずらす
      st.getRange(newRow, 19, 1, 3).setValues([[TYPO_NEW_STATUS, stage, TYPO_BALL]]);
    }
    log.push('1) 定義表：' + newRow + '行目に [' + TYPO_NEW_STATUS + ' / ' + stage + ' / ' + TYPO_BALL + '] を追加' +
             '（「' + TYPO_AFTER_STATUS + '」の下。S:U列だけずらします）');
  }

  // ---------- 2) H列のプルダウン ----------
  log.push('2) プルダウン：' + updateDropdown_(st, ms, newRow, dryRun));

  // ---------- 3) 条件付き書式（チップ色・行の色） ----------
  log.push('3) 色：' + copyColorRules_(ms, table, dryRun));

  log.push('※プルダウン自体のチップ色（入力規則のパレットで付けた色）はスクリプトから設定できません。' +
           'H列のチップが灰色のままなら、データ→データの入力規則 で緑を選んでください（1分）。');
  if (!dryRun) SpreadsheetApp.flush();
  log.push(dryRun ? '問題なければ addTypoCheckStatus を実行してください。'
                  : '完了しました。テスト用の案件でステータスを切り替えて、担当・締切・色を確認してください。');
  Logger.log(log.join('\n'));
}

/** 設定シート S:U を読む → { rowOf: {ステータス:行}, stageOf: {ステータス:工程}, ballOf: {ステータス:ボール} } */
function readStatusTable_(st) {
  var last = st.getLastRow();
  var vals = st.getRange(1, 19, last, 3).getDisplayValues();
  var t = { rowOf: {}, stageOf: {}, ballOf: {} };
  for (var i = 1; i < vals.length; i++) { // 1行目は見出し
    var name = String(vals[i][0]).trim();
    if (!name) continue;
    t.rowOf[name] = i + 1;
    t.stageOf[name] = String(vals[i][1]).trim();
    t.ballOf[name] = String(vals[i][2]).trim();
  }
  return t;
}

/** マスターH列のプルダウンに新ステータスが出るようにする */
function updateDropdown_(st, ms, newRow, dryRun) {
  var n = ms.getMaxRows() - 1;
  var range = ms.getRange(2, TYPO_STATUS_COL, n, 1);
  var dvs = range.getDataValidations();
  var sample = null;
  for (var i = 0; i < dvs.length; i++) if (dvs[i][0]) { sample = dvs[i][0]; break; }
  if (!sample) return '⚠ H列に入力規則が見つかりません。手動で確認してください。';

  var type = sample.getCriteriaType();
  var args = sample.getCriteriaValues();

  if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
    var src = args[0];
    var inside = src.getSheet().getName() === '設定' && src.getColumn() <= 19 && src.getLastColumn() >= 19 &&
                 src.getRow() <= newRow && src.getLastRow() >= newRow;
    // 範囲の内側に行を挿入すると範囲は自動で1行伸びるので、確認モードでも同じ判定でよい
    if (inside) {
      return '設定シートの範囲 ' + src.getA1Notation() + ' を参照しているので自動で追従します（変更なし）';
    }
    if (src.getSheet().getName() !== '設定' || src.getColumn() > 19 || src.getLastColumn() < 19) {
      return '⚠ 想定外の範囲（' + src.getSheet().getName() + '!' + src.getA1Notation() + '）を参照しています。手動で確認してください。';
    }
    // 範囲が短くて新しい行を含まない場合は伸ばす
    var widened = src.getSheet().getRange(src.getRow(), src.getColumn(),
                                          Math.max(src.getLastRow(), newRow) - src.getRow() + 1, src.getNumColumns());
    if (!dryRun) rewriteValidations_(range, dvs, type, function(a) { return [widened, a[1]]; });
    return '参照範囲を ' + src.getA1Notation() + ' → ' + widened.getA1Notation() + ' に広げ' + (dryRun ? 'ます' : 'ました');
  }

  if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    var list = args[0].slice();
    if (list.indexOf(TYPO_NEW_STATUS) >= 0) return 'すでに選択肢にあります（スキップ）';
    var pos = list.indexOf(TYPO_AFTER_STATUS);
    list.splice(pos >= 0 ? pos + 1 : list.length, 0, TYPO_NEW_STATUS);
    if (!dryRun) rewriteValidations_(range, dvs, type, function(a) { return [list, a[1]]; });
    return '選択肢リストに追加' + (dryRun ? 'します' : 'しました') + '（「' + TYPO_AFTER_STATUS + '」の次）';
  }

  return '⚠ 想定外の入力規則（' + type + '）です。手動で確認してください。';
}

/** 同じ種類の入力規則が付いているセルだけ、条件を差し替えて書き戻す（他の設定はそのまま） */
function rewriteValidations_(range, dvs, type, mapArgs) {
  var out = dvs.map(function(row) {
    var dv = row[0];
    if (!dv || dv.getCriteriaType() !== type) return [dv];
    return [dv.copy().withCriteria(type, mapArgs(dv.getCriteriaValues())).build()];
  });
  range.setDataValidations(out);
}

/** BO担当の既存ステータスの色ルールを複製して、新ステータス用のルールを作る */
function copyColorRules_(ms, table, dryRun) {
  var rules = ms.getConditionalFormatRules();
  for (var i = 0; i < rules.length; i++) {
    if (ruleMentions_(rules[i], TYPO_NEW_STATUS)) return 'すでに新ステータス用の色ルールがあります（スキップ）';
  }

  // お手本：ボールがBO担当のステータス（完了以外）で、色ルールがあるもの
  var candidates = ['アップロード待ち', 'リンク共有待ち', 'クライアント共有待ち'];
  Object.keys(table.ballOf).forEach(function(name) {
    if (table.ballOf[name] === TYPO_BALL && candidates.indexOf(name) < 0) candidates.push(name);
  });
  var model = null;
  for (var c = 0; c < candidates.length && !model; c++) {
    var name = candidates[c];
    if (name === '完了' || name === TYPO_NEW_STATUS || table.ballOf[name] !== TYPO_BALL) continue;
    if (rules.some(function(r) { return ruleMentions_(r, name); })) model = name;
  }
  if (!model) return '⚠ お手本にできるBO担当ステータスの色ルールが見つかりません。色は手動で設定してください。';

  var out = [], copied = 0;
  rules.forEach(function(r) {
    if (ruleMentions_(r, model)) {
      var cond = r.getBooleanCondition();
      var vals = cond.getCriteriaValues().map(function(v) {
        return typeof v === 'string' ? v.split(model).join(TYPO_NEW_STATUS) : v;
      });
      // 複製はお手本の直前に置く（優先順位をお手本と同じ位置にするため）
      out.push(r.copy().withCriteria(cond.getCriteriaType(), vals).build());
      copied++;
    }
    out.push(r);
  });
  if (!dryRun) ms.setConditionalFormatRules(out);
  return '「' + model + '」の色ルール ' + copied + '件を複製' + (dryRun ? 'します' : 'しました') +
         '（チップ色・行の色が同じ緑になります）';
}

/** 条件付き書式ルールが、指定したステータス名を条件に含んでいるか */
function ruleMentions_(rule, name) {
  var cond = rule.getBooleanCondition();
  if (!cond) return false; // グラデーション等は対象外
  return cond.getCriteriaValues().some(function(v) {
    return typeof v === 'string' && v.indexOf(name) >= 0;
  });
}
