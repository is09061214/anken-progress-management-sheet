/* ========================================================================
 * 締切手入力サポートパッチ（V4スリム版）— 試作版V2（V3レイアウト）用
 * 2026-07-03
 * 2026-07-03 更新: BOの要望により「手入力セルを薄紫にする」機能を廃止。
 *              removeManualMarking を1回実行すると、シートから紫の
 *              条件付き書式が消え、ヘッダーのメモも紫の記述なしに更新される。
 *
 * 【前提】
 *   順算の日数変更（制作10/サムネ12/CL提出14）はITARUさんが設定シートJ2:J4を
 *   直接編集して対応済み。数式は従来どおり「撮影日＋J2/J3/J4」のまま変更しない。
 *
 * 【このパッチに今あるもの（2つ）】
 *   1. 紫色マーキングの削除機能（removeManualMarking）
 *        以前 applyManualMarking で追加した「手入力セルを薄紫にする」
 *        条件付き書式（L/N/O/Q列のISFORMULAルール）を削除する。
 *   2. 自動計算に戻す機能（restoreAutoDeadlines）
 *        設定シート J8 に行番号（カンマ区切り）か「all」を入力して
 *        restoreAutoDeadlines を実行 → その行の締切4列が元の数式に戻る。
 *        ※こちらは色と無関係に引き続き使える。
 *
 *   ※ マスターの数式・設定シートの日数には一切触りません。
 *   ※ 実行前にヘッダーを検証し、レイアウトが想定と違う場合は何も変更せず中断。
 *
 * 【使い方】
 *   Apps Scriptプロジェクト「案件管理シート_2026.7〜」のこのファイルを
 *   まるごと差し替え、removeManualMarking を実行（1回だけ）。
 *
 * 【対象シートのレイアウト（V3・22列）】
 *   A クライアント/B 🚦/C 案件タイプ/D タイトル/E 完成尺/F 撮影日/G 投稿予定日/
 *   H ステータス(手動)/I 修正対象/J 修正締切/K 制作担当/L 制作締切/M サムネ担当/
 *   N サムネ締切/O CL提出締切/P 提出稿/Q 公開設定締切/R BO担当/S メモ/
 *   T 現在担当(隠)/U 現在締切(隠)/V 信号(隠)
 * ======================================================================== */

var CHAIN_SS_ID = '1oXwpCyw_g4GT1SFh4V_DW6RL7Q8h5FhIx9nbEV_eLXo'; // 試作版V2

/* ---- 数式パーツ（既存protoV2.gsと名前が被らないよう c 接頭辞） ---- */
// 案件タイプの工程フラグ（n: 2=撮影 3=制作 4=サムネ 5=CL提出 6=公開設定）
function cTyp_(r, n){
  return 'IFERROR(VLOOKUP(IF($C'+r+'="","通常",$C'+r+'),\'設定\'!$L:$Q,'+n+',FALSE),"○")';
}
// 逆算リードタイム（設定A:D クライアント別、n: 2=制作 3=サムネ 4=CL提出）
function cLead_(r, n){
  return 'IFERROR(VLOOKUP($A'+r+',\'設定\'!$A:$D,'+n+',FALSE),VLOOKUP("（標準）",\'設定\'!$A:$D,'+n+',FALSE))';
}

/* ---- 締切の数式（現行の順算そのまま。restoreAutoDeadlines用） ---- */
// 工程締切 ＝ 撮影日 + 設定J列。撮影日空はJ6設定に従い逆算 or 空欄
function cDue_(r, typN, plusCell, leadN){
  return '=IF($A'+r+'="","",IF('+cTyp_(r,typN)+'="—","",' +
    'IF($F'+r+'<>"",$F'+r+'+\'設定\'!'+plusCell+',' +
    'IF(OR('+cTyp_(r,2)+'="—",\'設定\'!$J$6="逆算で仮計算"),' +
    'IF($G'+r+'="","",$G'+r+'-'+cLead_(r,leadN)+'),""))))';
}
function cFormulaL_(r){ return cDue_(r, 3, '$J$2', 2); } // 制作締切
function cFormulaN_(r){ return cDue_(r, 4, '$J$3', 3); } // サムネ締切
function cFormulaO_(r){ return cDue_(r, 5, '$J$4', 4); } // CL提出締切
// 公開設定締切 ＝ 投稿予定日 − J5
function cFormulaQ_(r){
  return '=IF($A'+r+'="","",IF('+cTyp_(r,6)+'="—","",' +
    'IF($G'+r+'="","",$G'+r+'-\'設定\'!$J$5)))';
}

/* ---- レイアウト検証（想定と違ったら何も変更せず中断） ---- */
function cVerifyLayout_(ms, st){
  var checks = [
    ['マスターA1', ms.getRange('A1').getValue(), 'クライアント'],
    ['マスターC1', ms.getRange('C1').getValue(), '案件タイプ'],
    ['マスターF1', ms.getRange('F1').getValue(), '撮影日'],
    ['マスターG1', ms.getRange('G1').getValue(), '投稿予定日'],
    ['マスターL1', ms.getRange('L1').getValue(), '制作締切'],
    ['マスターN1', ms.getRange('N1').getValue(), 'サムネ締切'],
    ['マスターO1', ms.getRange('O1').getValue(), 'CL提出締切'],
    ['マスターQ1', ms.getRange('Q1').getValue(), '公開設定'],
    ['設定L1',     st.getRange('L1').getValue(), '案件タイプ']
  ];
  var bad = [];
  checks.forEach(function(c){
    if (String(c[1]).indexOf(c[2]) === -1) bad.push(c[0]+'が「'+c[1]+'」（期待:「'+c[2]+'」を含む）');
  });
  var j6 = String(st.getRange('J6').getValue());
  if (j6 !== '逆算で仮計算' && j6 !== '情報不足(灰)'){
    bad.push('設定J6が「'+j6+'」（期待: 逆算で仮計算 / 情報不足(灰)）');
  }
  if (bad.length){
    throw new Error('レイアウトが想定と違うため中断しました（何も変更していません）: ' + bad.join(' / '));
  }
}

/* ---- フィルタの全条件を退避して解除（隠れ行にsetFormulaが効かない対策） ---- */
function cFilterOff_(ms){
  var fl = ms.getFilter();
  var saved = {};
  if (fl){
    var lastCol = ms.getLastColumn();
    for (var c = 1; c <= lastCol; c++){
      var cr = fl.getColumnFilterCriteria(c);
      if (cr){ saved[c] = cr.copy(); fl.removeColumnFilterCriteria(c); }
    }
    SpreadsheetApp.flush();
  }
  return saved;
}
function cFilterRestore_(ms, saved){
  var fl = ms.getFilter();
  if (!fl) return;
  for (var c in saved){
    try { fl.setColumnFilterCriteria(Number(c), saved[c].build()); }
    catch(e){ Logger.log('フィルタ復元スキップ col'+c+': '+e); }
  }
}

/* ==================== メイン：手入力セルの紫色マーキングを削除 ==================== */
// 2026-07-03 BOの要望で色分け機能を廃止。以前applyManualMarkingで追加した
// ISFORMULA条件付き書式（L/N/O/Q列の薄紫）を削除し、ヘッダーのメモも
// 紫の記述なしのものに置き換える。数式・設定日数・行の色分け（ボール所在）には触らない。
function removeManualMarking(){
  var ss = SpreadsheetApp.openById(CHAIN_SS_ID);
  var ms = ss.getSheetByName('マスター');
  var st = ss.getSheetByName('設定');
  cVerifyLayout_(ms, st);

  // 1) ISFORMULAを使った条件付き書式ルール（＝手入力セルの紫）だけを取り除く
  var before = ms.getConditionalFormatRules();
  var rules = before.filter(function(rule){
    var bc = rule.getBooleanCondition && rule.getBooleanCondition();
    if (!bc) return true;
    try {
      var vals = bc.getCriteriaValues();
      return !(vals && vals.length && String(vals[0]).indexOf('ISFORMULA') > -1);
    } catch(e){ return true; }
  });
  ms.setConditionalFormatRules(rules);

  // 2) マスターのヘッダーメモを「紫」の記述なしに更新（手入力OK＋自動に戻す方法は残す）
  var memo = '自動計算（撮影日から順算・日数は設定シートで変更可）。\n' +
             '個別調整したい場合は日付を直接入力してOK。\n' +
             '自動に戻すには設定シートJ8に行番号を入れて restoreAutoDeadlines を実行。';
  ['L1','N1','O1','Q1'].forEach(function(a){ ms.getRange(a).setNote(memo); });

  Logger.log('removeManualMarking 完了: 紫の条件付き書式を'+(before.length - rules.length)+'件削除。' +
    'ヘッダーメモを更新。数式・設定日数・行の色分けは無変更。');
}

/* ==================== 手入力を自動計算に戻す ==================== */
function restoreAutoDeadlines(){
  var ss = SpreadsheetApp.openById(CHAIN_SS_ID);
  var ms = ss.getSheetByName('マスター');
  var st = ss.getSheetByName('設定');
  cVerifyLayout_(ms, st);

  var last = ms.getMaxRows();
  var input = String(st.getRange('J8').getValue()).trim();
  if (!input){
    Logger.log('設定J8が空です。戻したい行番号（カンマ区切り）か all を入力してから実行してください。');
    return;
  }
  var rows = [];
  if (input.toLowerCase() === 'all'){
    for (var r = 2; r <= last; r++) rows.push(r);
  } else {
    input.split(/[,、\s]+/).forEach(function(t){
      var n = parseInt(t, 10);
      if (n >= 2 && n <= last) rows.push(n);
    });
  }
  if (!rows.length){
    Logger.log('J8「'+input+'」から有効な行番号が読み取れませんでした（2〜'+last+'）。');
    return;
  }

  var savedFilter = cFilterOff_(ms);
  rows.forEach(function(r){
    ms.getRange('L'+r).setFormula(cFormulaL_(r));
    ms.getRange('N'+r).setFormula(cFormulaN_(r));
    ms.getRange('O'+r).setFormula(cFormulaO_(r));
    ms.getRange('Q'+r).setFormula(cFormulaQ_(r));
  });
  cFilterRestore_(ms, savedFilter);
  st.getRange('J8').clearContent();

  Logger.log('restoreAutoDeadlines 完了: '+rows.length+'行（'+(rows.length>20?'先頭20行のみ表示: '+rows.slice(0,20).join(','):rows.join(','))+'）を自動計算に戻しました。');
}
