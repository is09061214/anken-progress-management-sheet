/**
 * 【試作版V2】案件進捗管理シート 改修プロトタイプ生成スクリプト
 * =====================================================================
 * BO2人の意見（2026-07 議論）を反映した改修を「本番の複製」に適用する。
 * 本番シート（1LLkk...）は読むだけで一切変更しない。
 *
 * ■ 使い方
 *   script.google.com の既存プロジェクトにこのファイルを追加（or 貼り付け）
 *   → 関数「buildPrototype」を選んで実行 → 実行ログに出るURLを開く
 *
 * ■ 改修内容（グリルで合意した設計）
 *   1. 行の色 ＝ ボールの所在（役割基準）
 *        紫=未着手 / 水色=制作・サムネ側 / 緑=BOの番(CL提出・公開設定) /
 *        ピンク=クライアント確認中 / オレンジ=修正中 / グレー=完了
 *   2. 緊急度は別レイヤー：B列「🚦」に絵文字ドット（🔴要対応/🟡もうすぐ/🔵順調/⚪情報不足）
 *      ＋ 締切セルの 赤(超過)・青(今ここ) は従来どおり
 *      ※ドットは「値」として出す（V25の教訓: グループ化でも確実に表示される唯一の方法）
 *      ※行の色塗りは「単一列ルールの束」で実装（V24/V27の教訓: 複数列範囲のCFは
 *        テーブルのグループ化で崩れる。単一列CF＋テーブル内A1参照なら効く）
 *   3. 「案件タイプ」列（通常/撮影なし/制作のみ/サムネなし）
 *        対象外の工程は締切が空欄になり、ステータス計算からスキップ。
 *        タイプの定義（○/—）は設定シートで編集可能。
 *   4. 「撮影日」列：工程締切（制作/サムネ/CL提出）は撮影日から順算（設定シートの日数）。
 *        公開設定・納品の締切だけは 投稿予定日−1日（前日アップのリミット）で逆算固定。
 *        順算のCL提出締切がリミットを超えたら信号=赤（スケジュール破綻）。
 *        撮影日が空のときの挙動は設定シートJ6で切替（逆算で仮計算 / 情報不足(灰)）。
 *   5. 「修正対象」列（動画/サムネ/両方）：「修正中（CL修正 / サムネ）」のように表示。
 *        修正対象=サムネ のときはボールがサムネ担当に戻る。
 *   6. 「提出稿」列（初稿/修正稿・手動）：「クライアント確認中（修正稿）」のように表示。
 *
 * ■ 試作版の新レイアウト（29列）
 *   A クライアント / B 🚦 / C 案件タイプ / D 動画タイトル / E 完成尺 / F 撮影日 /
 *   G 投稿予定日 / H 現在ステータス(自動) / I 差し戻し / J 修正対象 / K 修正締切 /
 *   L 制作担当 / M 制作着手✓ / N 制作締切(自動) / O 制作✓ /
 *   P サムネ担当 / Q サムネ締切(自動) / R サムネ✓ / S CL提出締切(自動) / T CL提出✓ /
 *   U 提出稿 / V CLチェック✓ / W 公開設定・納品締切(自動) / X 公開設定・納品✓ /
 *   Y BO担当 / Z メモ / AA 現在担当(隠) / AB 現在締切(隠) / AC 信号(隠)
 *
 * ■ 元レイアウト（本番・applyV31後）
 *   A クライアント / B タイトル / C 完成尺 / D 投稿予定日 / E ステータス / F 差し戻し /
 *   G 修正締切 / H 制作担当 / I 制作着手✓ / J 制作締切 / K 制作✓ / L サムネ担当 /
 *   M サムネ締切 / N サムネ✓ / O CL提出締切 / P CL提出✓ / Q CLチェック✓ /
 *   R 公開設定・納品締切 / S 公開設定・納品✓ / T BO担当 / U メモ / V-X (自動・隠)
 */

var PROTO_SRC = '1LLkkdljyokHqjLll9rcG-ClSYnXx43D-xZtmLkAp0Ds';
var PR = 200; // 入力行数

/* ---- 数式パーツ ---- */
// 案件タイプの工程フラグ（n: 2=撮影 3=制作 4=サムネ 5=CL提出 6=公開設定）。タイプ空欄は「通常」扱い
function pTyp_(r, n){
  return 'IFERROR(VLOOKUP(IF($C'+r+'="","通常",$C'+r+'),\'設定\'!$L:$Q,'+n+',FALSE),"○")';
}
// （超過）表示
function pOver_(r, col){
  return 'IF(AND($'+col+r+'<>"",$'+col+r+'<TODAY()),"（超過）","")';
}
// 逆算リードタイム（設定A:D クライアント別、n: 2=制作 3=サムネ 4=CL提出）
function pLead_(r, n){
  return 'IFERROR(VLOOKUP($A'+r+',\'設定\'!$A:$D,'+n+',FALSE),VLOOKUP("（標準）",\'設定\'!$A:$D,'+n+',FALSE))';
}

/* ==================== メイン ==================== */
function buildPrototype(){
  var ss = SpreadsheetApp.openById(PROTO_SRC).copy('【試作版V2】案件進捗管理シート');
  var ms = ss.getSheetByName('マスター');
  var L = PR + 1;

  // 0) 完了フィルタを一時解除（列挿入・書式コピーを通すため）
  var fl = ms.getFilter();
  if (fl){ try{ fl.removeColumnFilterCriteria(5); SpreadsheetApp.flush(); }catch(e){ Logger.log('filter off skip: '+e); } }

  // 1) 列挿入（元レイアウト基準・右→左）
  ms.insertColumnsAfter(16, 1); // 提出稿（CL提出✓の右）
  ms.insertColumnsAfter(6, 1);  // 修正対象（差し戻しの右）
  ms.insertColumnsAfter(3, 1);  // 撮影日（完成尺の右）
  ms.insertColumnsAfter(1, 1);  // 案件タイプ（クライアントの右）
  ms.insertColumnsAfter(1, 1);  // 🚦（クライアントの右＝タイプの左）※テーブル内に入れて Group by 崩れを防ぐ
  SpreadsheetApp.flush();

  // 2) 設定シート拡張（数式が参照するので先に作る）
  buildProtoSettings_(ss);

  // 3) ヘッダー
  [['B1','🚦'],['C1','案件タイプ'],['F1','撮影日'],['J1','修正対象'],['U1','提出稿']].forEach(function(h){
    ms.getRange('A1').copyTo(ms.getRange(h[0]), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    ms.getRange(h[0]).setValue(h[1]);
  });

  // 4) 新列の型・書式の整え（テーブルの列型継承への対策）
  // 🚦 B: クライアントの入力規則を継承しているのでクリア
  ms.getRange('B2:B'+L).clearDataValidations().clearContent();
  // 撮影日 F: 完成尺(テキスト)を継承 → 投稿予定日Gの書式をコピーして日付に
  ms.getRange('G2:G'+L).copyTo(ms.getRange('F2:F'+L), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  ms.getRange('F2:F'+L).clearDataValidations().clearContent();
  try{ ms.getRange('F2:F'+L).setNumberFormat('m/d'); }catch(e){ Logger.log('F format skip: '+e); }
  // 提出稿 U: CL提出✓(チェックボックス)を継承 → 解除して差し戻しIの書式に
  try{ ms.getRange('U2:U'+L).removeCheckboxes(); }catch(e){ Logger.log('U removeCheckboxes skip: '+e); }
  ms.getRange('U2:U'+L).clearDataValidations().clearContent();
  ms.getRange('I2:I'+L).copyTo(ms.getRange('U2:U'+L), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  // 5) 入力規則（ドロップダウン）
  var typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getRange('設定!L2:L20'), true).setAllowInvalid(true).build();
  ms.getRange('C2:C'+L).setDataValidation(typeRule);
  var taishoRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['動画','サムネ','両方'], true).setAllowInvalid(true).build();
  ms.getRange('J2:J'+L).setDataValidation(taishoRule);
  var kouRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['初稿','修正稿'], true).setAllowInvalid(true).build();
  ms.getRange('U2:U'+L).setDataValidation(kouRule);

  // 6) 数式（全行に明示セット）
  setProtoFormulas_(ms);

  // 7) 条件付き書式 全面刷新（旧ルールは全部破棄）
  setProtoCF_(ms);

  // 8) 列幅・固定・非表示
  ms.setColumnWidth(2, 30);   // 🚦
  ms.setColumnWidth(3, 92);   // 案件タイプ
  ms.setColumnWidth(6, 64);   // 撮影日
  ms.setColumnWidth(10, 78);  // 修正対象
  ms.setColumnWidth(21, 64);  // 提出稿
  try{ ms.hideColumns(27, 3); }catch(e){} // AA-AC
  ms.setFrozenColumns(4);     // クライアント/🚦/タイプ/タイトル

  // 9) ダッシュボード更新（信号=AC・新列参照・凡例）
  updateProtoDashboard_(ss);

  // 10) フィルタビュー（増田担当/岩渕担当）を作り直し（BO担当=Y列）
  try{ rebuildProtoFilterViews_(ss, ms); }catch(e){ Logger.log('filterViews skip: '+e); }

  // 11) 完了フィルタを再設定（現在ステータス=H列(8)）
  fl = ms.getFilter();
  if (fl){ try{ fl.setColumnFilterCriteria(8, SpreadsheetApp.newFilterCriteria().setHiddenValues(['完了']).build()); }catch(e){ Logger.log('filter on skip: '+e); } }

  Logger.log('試作版が完成しました。URL:\n' + ss.getUrl());
}

/* ==================== 設定シート拡張 ==================== */
function buildProtoSettings_(ss){
  var st = ss.getSheetByName('設定');
  // 順算設定
  st.getRange('I1').setValue('順算設定（撮影日から）').setFontWeight('bold').setBackground('#f1efe8');
  st.getRange('I2:J6').setValues([
    ['制作締切 ＝ 撮影日 ＋（日）', 7],
    ['サムネ締切 ＝ 撮影日 ＋（日）', 9],
    ['CL提出締切 ＝ 撮影日 ＋（日）', 12],
    ['公開リミット ＝ 投稿予定日 −（日）', 1],
    ['撮影日が空のときの締切', '逆算で仮計算']
  ]);
  st.getRange('J6').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['逆算で仮計算','情報不足(灰)'], true).setAllowInvalid(false).build());
  st.getRange('I1').setNote(
    '工程締切は「撮影日＋この日数」で順算します（日数は要調整・仮の値）。\n' +
    '公開設定・納品の締切だけは「投稿予定日−1日」（前日アップ必須のリミット）。\n' +
    '順算したCL提出締切がリミットを超えると、信号が赤（スケジュール破綻）になります。\n' +
    'J6「撮影日が空のときの締切」:\n' +
    '  逆算で仮計算 … 従来どおり投稿予定日から逆算で締切を出す（移行期間むけ・既定）\n' +
    '  情報不足(灰) … 撮影日が入るまで信号を灰にして入力を促す（本運用むけ・厳格）');
  st.setColumnWidth(9, 210);
  // 案件タイプ定義表
  st.getRange('L1:Q1').setValues([['案件タイプ','撮影','制作','サムネ','CL提出','公開設定']])
    .setFontWeight('bold').setBackground('#f1efe8');
  st.getRange('L2:Q5').setValues([
    ['通常',     '○','○','○','○','○'],
    ['撮影なし', '—','○','○','○','○'],
    ['制作のみ', '—','○','—','○','○'],
    ['サムネなし','○','○','—','○','○']
  ]);
  var maruRule = SpreadsheetApp.newDataValidation().requireValueInList(['○','—'], true).setAllowInvalid(true).build();
  st.getRange('M2:Q20').setDataValidation(maruRule);
  st.getRange('L1').setNote(
    '案件タイプごとに、どの工程があるかを ○/— で定義します。\n' +
    '—の工程は締切が空欄になり、ステータス計算からスキップされます。\n' +
    '行を足せば新しいタイプを作れます（マスターのドロップダウンに自動反映）。\n' +
    'マスターでタイプが空欄の案件は「通常」として扱います。');
}

/* ==================== マスター数式 ==================== */
function setProtoFormulas_(ms){
  var L = PR + 1;
  var bF=[], hF=[], nF=[], qF=[], sF=[], wF=[], aaF=[], abF=[], acF=[];
  for (var r = 2; r <= L; r++){
    var t2 = pTyp_(r,2), t3 = pTyp_(r,3), t4 = pTyp_(r,4), t5 = pTyp_(r,5), t6 = pTyp_(r,6);
    var ovN = pOver_(r,'N'), ovQ = pOver_(r,'Q'), ovS = pOver_(r,'S'), ovW = pOver_(r,'W'), ovK = pOver_(r,'K');

    // 工程締切（順算。撮影日空は J6 設定に従い逆算 or 空欄）
    function due(tp, plusCell, leadN){
      return '=IF($A'+r+'="","",IF('+tp+'="—","",IF($F'+r+'<>"",$F'+r+'+\'設定\'!'+plusCell+',' +
        'IF(OR('+pTyp_(r,2)+'="—",\'設定\'!$J$6="逆算で仮計算"),' +
        'IF($G'+r+'="","",$G'+r+'-'+pLead_(r,leadN)+'),""))))';
    }
    nF.push([ due(t3,'$J$2',2) ]); // 制作締切
    qF.push([ due(t4,'$J$3',3) ]); // サムネ締切
    sF.push([ due(t5,'$J$4',4) ]); // CL提出締切
    // 公開設定・納品締切 ＝ 投稿予定日 − 公開リミット（前日アップ）
    wF.push([ '=IF($A'+r+'="","",IF('+t6+'="—","",IF($G'+r+'="","",$G'+r+'-\'設定\'!$J$5)))' ]);

    // 🚦ドット B（値として出す＝グループ化でも確実。テーブル内の列だけ参照）
    var cdIFS = 'IFS($I'+r+'<>"",$K'+r+',' +
      'AND($M'+r+'<>TRUE,$O'+r+'<>TRUE,'+t3+'="○"),$N'+r+',' +
      'AND($O'+r+'<>TRUE,'+t3+'="○"),$N'+r+',' +
      'AND($R'+r+'<>TRUE,'+t4+'="○"),$Q'+r+',' +
      'AND($T'+r+'<>TRUE,'+t5+'="○"),$S'+r+',' +
      'AND($V'+r+'<>TRUE,'+t5+'="○"),"",' +
      'AND($X'+r+'<>TRUE,'+t6+'="○"),$W'+r+',' +
      'TRUE,"")';
    bF.push([ '=IF($A'+r+'="","",LET(cd,'+cdIFS+',' +
      'IF($G'+r+'="","⚪",' +
      'IF(AND('+t2+'="○",$F'+r+'="",\'設定\'!$J$6<>"逆算で仮計算"),"⚪",' +
      'IF($H'+r+'="完了","",' +
      'IF(OR($G'+r+'<=TODAY()+2,AND(cd<>"",cd<TODAY()),AND($S'+r+'<>"",$W'+r+'<>"",$S'+r+'>$W'+r+')),"🔴",' +
      'IF(OR($G'+r+'<=TODAY()+5,AND($O'+r+'<>TRUE,$N'+r+'<>"",$N'+r+'<=TODAY()+1)),"🟡","🔵")))))))' ]);

    // 現在ステータス H（修正中優先 → 工程を前から順に。—の工程はスキップ）
    hF.push([ '=IF($A'+r+'="","",' +
      'IF($I'+r+'<>"","修正中（"&$I'+r+'&IF($J'+r+'<>""," / "&$J'+r+',"")&"）"&'+ovK+',' +
      'IFS(' +
      'AND($M'+r+'<>TRUE,$O'+r+'<>TRUE,'+t3+'="○"),"未着手"&'+ovN+',' +
      'AND($O'+r+'<>TRUE,'+t3+'="○"),"制作待ち"&'+ovN+',' +
      'AND($R'+r+'<>TRUE,'+t4+'="○"),"サムネ待ち"&'+ovQ+',' +
      'AND($T'+r+'<>TRUE,'+t5+'="○"),"CL提出待ち"&'+ovS+',' +
      'AND($V'+r+'<>TRUE,'+t5+'="○"),"クライアント確認中"&IF($U'+r+'<>"","（"&$U'+r+'&"）",""),' +
      'AND($X'+r+'<>TRUE,'+t6+'="○"),"公開設定・納品待ち"&'+ovW+',' +
      'TRUE,"完了")))' ]);

    // 現在担当 AA（修正対象=サムネ → サムネ担当に戻る）
    aaF.push([ '=IF($A'+r+'="","",' +
      'IF($I'+r+'<>"",IF($J'+r+'="サムネ",$P'+r+',$L'+r+'),' +
      'IFS(' +
      'AND($M'+r+'<>TRUE,'+t3+'="○"),$L'+r+',' +
      'AND($O'+r+'<>TRUE,'+t3+'="○"),$L'+r+',' +
      'AND($R'+r+'<>TRUE,'+t4+'="○"),$P'+r+',' +
      'AND($T'+r+'<>TRUE,'+t5+'="○"),$Y'+r+',' +
      'AND($V'+r+'<>TRUE,'+t5+'="○"),"クライアント",' +
      'AND($X'+r+'<>TRUE,'+t6+'="○"),$Y'+r+',' +
      'TRUE,"—")))' ]);

    // 現在締切 AB（クライアント確認中は締切なし）
    abF.push([ '=IF($A'+r+'="","",' +
      'IF($I'+r+'<>"",$K'+r+',' +
      'IFS(' +
      'AND($M'+r+'<>TRUE,'+t3+'="○"),$N'+r+',' +
      'AND($O'+r+'<>TRUE,'+t3+'="○"),$N'+r+',' +
      'AND($R'+r+'<>TRUE,'+t4+'="○"),$Q'+r+',' +
      'AND($T'+r+'<>TRUE,'+t5+'="○"),$S'+r+',' +
      'AND($V'+r+'<>TRUE,'+t5+'="○"),"",' +
      'AND($X'+r+'<>TRUE,'+t6+'="○"),$W'+r+',' +
      'TRUE,"")))' ]);

    // 信号 AC：灰=投稿予定日空 or（撮影ありタイプ×撮影日空×厳格モード）
    //         赤=公開2日以内 or 工程超過 or スケジュール破綻（CL提出締切>公開リミット）
    //         黄=公開5日以内 or 制作締切1日前 ／ 青=それ以外・完了
    acF.push([ '=IF($A'+r+'="","",' +
      'IF($G'+r+'="","灰",' +
      'IF(AND('+t2+'="○",$F'+r+'="",\'設定\'!$J$6<>"逆算で仮計算"),"灰",' +
      'IF($H'+r+'="完了","青",' +
      'IF(OR($G'+r+'<=TODAY()+2,AND($AB'+r+'<>"",$AB'+r+'<TODAY()),AND($S'+r+'<>"",$W'+r+'<>"",$S'+r+'>$W'+r+')),"赤",' +
      'IF(OR($G'+r+'<=TODAY()+5,AND($O'+r+'<>TRUE,$N'+r+'<>"",$N'+r+'<=TODAY()+1)),"黄","青"))))))' ]);
  }
  ms.getRange('B2:B'+L).setFormulas(bF);
  ms.getRange('B2:B'+L).setHorizontalAlignment('center');
  ms.getRange('H2:H'+L).setFormulas(hF);
  ms.getRange('N2:N'+L).setFormulas(nF);
  ms.getRange('Q2:Q'+L).setFormulas(qF);
  ms.getRange('S2:S'+L).setFormulas(sF);
  ms.getRange('W2:W'+L).setFormulas(wF);
  ms.getRange('AA2:AA'+L).setFormulas(aaF);
  ms.getRange('AB2:AB'+L).setFormulas(abF);
  ms.getRange('AC2:AC'+L).setFormulas(acF);
}

/* ==================== 条件付き書式 ====================
 * V24/V27の教訓: 複数列範囲のCFはテーブルのグループ化で崩れる。
 * → すべて「単一列範囲＋テーブル内列のA1参照」のルールで構成する。 */
function setProtoCF_(ms){
  var L = PR + 1;
  var rules = [];
  function rule(a1, f, bg, fc){
    var b = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(f).setRanges([ms.getRange(a1)]);
    if (bg) b.setBackground(bg);
    if (fc) b.setFontColor(fc);
    rules.push(b.build());
  }
  // --- 締切セル：超過=赤（最優先）→ 今ここ=濃いめの青 ---
  rule('K2:K'+L, '=AND($I2<>"",$K2<>"",$K2<TODAY())', '#f6c7c7', '#a32d2d');                 // 修正締切
  rule('N2:N'+L, '=AND($O2<>TRUE,$N2<>"",$N2<TODAY())', '#f6c7c7', '#a32d2d');               // 制作
  rule('N2:N'+L, '=AND(OR(REGEXMATCH($H2&"","制作待ち"),REGEXMATCH($H2&"","未着手")),$N2<>"")', '#cfe4f7', '#14548f');
  rule('Q2:Q'+L, '=AND($R2<>TRUE,$Q2<>"",$Q2<TODAY())', '#f6c7c7', '#a32d2d');               // サムネ
  rule('Q2:Q'+L, '=AND(REGEXMATCH($H2&"","サムネ待ち"),$Q2<>"")', '#cfe4f7', '#14548f');
  rule('S2:S'+L, '=AND($T2<>TRUE,$S2<>"",$S2<TODAY())', '#f6c7c7', '#a32d2d');               // CL提出
  rule('S2:S'+L, '=AND(REGEXMATCH($H2&"","CL提出待ち"),$S2<>"")', '#cfe4f7', '#14548f');
  rule('W2:W'+L, '=AND($X2<>TRUE,$W2<>"",$W2<TODAY())', '#f6c7c7', '#a32d2d');               // 公開設定
  rule('W2:W'+L, '=AND(REGEXMATCH($H2&"","公開設定"),$W2<>"")', '#cfe4f7', '#14548f');
  // --- ステータスH：超過は赤で強調（ボール色より優先）---
  rule('H2:H'+L, '=REGEXMATCH($H2&"","超過")', '#f6c7c7', '#a32d2d');
  // --- 差し戻しI：選択中はオレンジ ---
  rule('I2:I'+L, '=$I2<>""', '#f9e3bf', '#854f0b');
  // --- 行の色 ＝ ボールの所在（単一列ルール × 対象列）---
  // 対象: 左の情報ブロック(A〜H) ＋ 担当列(L,P,Y) ＋ メモ(Z)。チェック・締切ゾーンは白のまま
  var BALL_COLS = ['A','B','C','D','E','F','G','H','L','P','Y','Z'];
  var BALL = [
    ['=REGEXMATCH($H2&"","修正中")', '#fdf1de', null],                                        // オレンジ=修正中
    ['=$H2="完了"', '#f1efe8', '#888780'],                                                    // グレー=完了
    ['=REGEXMATCH($H2&"","未着手")', '#f2ecfb', null],                                        // 紫=未着手
    ['=REGEXMATCH($H2&"","クライアント確認中")', '#fdecf3', null],                             // ピンク=CL確認中
    ['=OR(REGEXMATCH($H2&"","CL提出待ち"),REGEXMATCH($H2&"","公開設定"))', '#e5f2e8', null],   // 緑=BOの番
    ['=OR(REGEXMATCH($H2&"","制作待ち"),REGEXMATCH($H2&"","サムネ待ち"))', '#e9f2fc', null]    // 水色=制作側
  ];
  BALL.forEach(function(b){
    BALL_COLS.forEach(function(c){
      rule(c + '2:' + c + L, b[0], b[1], b[2]);
    });
  });
  ms.setConditionalFormatRules(rules);
}

/* ==================== ダッシュボード更新 ==================== */
function updateProtoDashboard_(ss){
  var dh = ss.getSheetByName('ダッシュボード');
  if (!dh) return;
  var M = 'マスター!';
  dh.getRange('A2').setFormula(
    '="本日 "&TEXT(TODAY(),"yyyy/mm/dd (ddd)")&"（Asia/Tokyo）　／　行の色=ボールの所在・🚦=緊急度（試作版V2）"');
  dh.getRange('A4').setNote(
    '信号の判定基準（試作版V2）\n' +
    '灰＝投稿予定日が未入力（厳格モード時は撮影日未入力も）\n' +
    '赤＝公開2日以内 / 工程の締切超過 / スケジュール破綻（CL提出締切が公開リミット超え）\n' +
    '黄＝公開5日以内 または 制作締切1日前\n' +
    '青＝それ以外・完了\n' +
    '行の色＝ボールの所在（紫=未着手・水色=制作サムネ側・緑=BOの番・ピンク=CL確認中・オレンジ=修正中・グレー=完了）');
  // サマリ（信号=AC）
  dh.getRange('A6').setFormula('=COUNTIF(' + M + 'AC2:AC,"赤")');
  dh.getRange('B6').setFormula('=COUNTIF(' + M + 'AC2:AC,"黄")');
  dh.getRange('C6').setFormula('=COUNTIF(' + M + 'AC2:AC,"青")');
  dh.getRange('D6').setFormula(
    '=SUMPRODUCT((' + M + 'A2:A201<>"")*(((' + M + 'D2:D201="")+(' + M + 'G2:G201="")+(' + M + 'L2:L201="")+(' + M + 'Y2:Y201=""))>0))');
  dh.getRange('E6').setFormula('=COUNTIF(' + M + 'A2:A,"<>")');
  // 判定基準の凡例（A7 結合セル）
  try{
    dh.getRange('A7').setValue(
      '判定基準：🔴要対応＝公開2日以内 or 締切超過 or 破綻(CL提出>公開リミット)／🟡もうすぐ＝公開5日以内 or 制作締切1日前／🔵順調／⚪情報不足　■行の色＝ボール：紫=未着手・水色=制作サムネ側・緑=BOの番・ピンク=CL確認中・オレンジ=修正中・グレー=完了');
  }catch(e){}
  // 赤黄リスト（クライアントA/タイトルD/投稿G/状況H/制作L/BO Y）
  dh.getRange('A10').setFormula(
    '=IFERROR(SORT(FILTER({' + M + 'AC2:AC,' + M + 'A2:A,' + M + 'D2:D,' + M + 'G2:G,' + M + 'G2:G-TODAY(),' + M + 'H2:H,' + M + 'L2:L,' + M + 'Y2:Y},' +
    '(' + M + 'AC2:AC="赤")+(' + M + 'AC2:AC="黄")),1,TRUE,4,TRUE),"いま確認が必要な案件はありません")');
  // 情報不足リスト（撮影日は厳格モード時のみ表示）
  dh.getRange('J8').setValue('情報不足の案件（タイトル/投稿予定日/担当/撮影日）').setFontWeight('bold').setFontSize(13).setFontColor('#5f5e5a');
  dh.getRange('J10').setFormula(
    '=IFERROR(FILTER({' + M + 'A2:A201,' + M + 'D2:D201,' + M + 'G2:G201,' +
    'IF(' + M + 'D2:D201="","タイトル ","")&IF(' + M + 'G2:G201="","投稿予定日 ","")&IF(' + M + 'L2:L201="","制作担当 ","")&IF(' + M + 'Y2:Y201="","BO担当 ","")&' +
    'IF(\'設定\'!$J$6="逆算で仮計算","",IF((IFERROR(VLOOKUP(IF(' + M + 'C2:C201="","通常",' + M + 'C2:C201),\'設定\'!$L:$Q,2,FALSE),"○")="○")*(' + M + 'F2:F201=""),"撮影日 ",""))},' +
    '(' + M + 'A2:A201<>"")*(((' + M + 'D2:D201="")+(' + M + 'G2:G201="")+(' + M + 'L2:L201="")+(' + M + 'Y2:Y201="")+' +
    'IF(\'設定\'!$J$6="逆算で仮計算",0,(IFERROR(VLOOKUP(IF(' + M + 'C2:C201="","通常",' + M + 'C2:C201),\'設定\'!$L:$Q,2,FALSE),"○")="○")*(' + M + 'F2:F201="")))>0)),"なし")');
  // 信号色（A10リスト）は既存CFが引き継がれるが、無ければ再作成
  var has = dh.getConditionalFormatRules().length > 0;
  if (!has){
    var rules=[];
    function sig(t,bg,fc){ rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(t).setBackground(bg).setFontColor(fc).setRanges([dh.getRange('A10:A'+PR)]).build()); }
    sig('赤','#fcebeb','#a32d2d'); sig('黄','#faeeda','#854f0b'); sig('青','#e6f1fb','#185fa5'); sig('灰','#f1efe8','#5f5e5a');
    dh.setConditionalFormatRules(rules);
  }
}

/* ==================== フィルタビュー再作成（増田担当/岩渕担当） ==================== */
function rebuildProtoFilterViews_(ss, ms){
  var id = ss.getId();
  var tok = ScriptApp.getOAuthToken();
  var meta = JSON.parse(UrlFetchApp.fetch(
    'https://sheets.googleapis.com/v4/spreadsheets/' + id + '?fields=sheets(properties(sheetId,title),filterViews(filterViewId,title))',
    { headers:{ Authorization:'Bearer ' + tok } }).getContentText());
  var reqs = [];
  (meta.sheets || []).forEach(function(s){
    if (!s.filterViews) return;
    s.filterViews.forEach(function(fv){
      if (fv.title === '増田担当' || fv.title === '岩渕担当'){
        reqs.push({ deleteFilterView:{ filterId: fv.filterViewId } });
      }
    });
  });
  var sid = ms.getSheetId();
  ['増田','岩渕'].forEach(function(n){
    reqs.push({ addFilterView:{ filter:{
      title: n + '担当',
      range: { sheetId: sid, startRowIndex: 0, endRowIndex: PR + 1, startColumnIndex: 0, endColumnIndex: 29 },
      filterSpecs: [{ columnIndex: 24, filterCriteria: { condition: { type:'TEXT_EQ', values:[{ userEnteredValue: n }] } } }] // Y=BO担当
    }}});
  });
  var resp = UrlFetchApp.fetch('https://sheets.googleapis.com/v4/spreadsheets/' + id + ':batchUpdate', {
    method:'post', contentType:'application/json',
    headers:{ Authorization:'Bearer ' + tok },
    payload: JSON.stringify({ requests: reqs }), muteHttpExceptions:true
  });
  Logger.log('filterViews: ' + resp.getResponseCode());
}

/* ========================================================================
 * ※デプロイ済みのApps Script（プロジェクト「案件管理シート_2026.7〜」protoV2.gs）は
 *   このローカル版からリファクタされている。主な差分:
 *   - buildPrototype の先頭に「probeAndRepair(); return;」ランチャー行あり
 *     （関数ドロップダウンが不安定なため。通常ビルド時はこの行を消す）
 *   - 行別数式は protoRowFormulas_(r) に共通化
 *   - PROTO_ID = '1oXwpCyw_g4GT1SFh4V_DW6RL7Q8h5FhIx9nbEV_eLXo'（試作版・正）
 *   - fixProtoRows(): PR をテーブル実行数(getMaxRows-1=1004)に上書きして
 *     数式/CF/入力規則/ダッシュボードを全行へ拡張するパッチ
 *   - probeAndRepair(): 旧数式（🔴🟡入り）残存行の特定＋修復。
 *     ★重要教訓: フィルタで非表示の行には setFormula/setFormulas が
 *     サイレントに効かない。全列 removeColumnFilterCriteria(1..29) で
 *     いったん全行を表示してから書き込むこと（列8だけの解除では不十分。
 *     本番から複製したフィルタは複数列に条件が残っていた）。
 *   - ダッシュボードのD6/J10は 201固定 → 全行(1005)に修正済み
 * ======================================================================== */

/* ========================================================================
 * 【V3・手動ステータス方式】2026-07-02 ITARU案で大転換（デプロイ版protoV2.gsが最新）
 *   - 現在ステータス(H)=手動ドロップダウン。旧シートの11ステータスをそのまま採用:
 *     タイトル待ち/未着手/編集中/修正中/サムネ待ち/砂田確認中/CL確認中/
 *     アップロード待ち/リンク共有待ち/クライアント共有待ち/完了
 *   - チェックボックス6列＋差し戻し＋公開設定✓を削除（計7列）
 *   - 設定シートに「ステータス定義表」(S:U): ステータス/締切工程/ボール(担当参照)
 *     → 現在締切=SWITCH(VLOOKUP(H,設定S:T))で動的。締切超過＝更新忘れ検知で🔴
 *   - 既存データは自動ステータス表示値から旧語彙へ自動変換して移行
 *     (制作待ち→編集中 / CL提出待ち→クライアント共有待ち / クライアント確認中→CL確認中 /
 *      公開設定・納品待ち→アップロード待ち)
 *   - V3レイアウト(22列): A クライアント/B 🚦/C 案件タイプ/D タイトル/E 完成尺/F 撮影日/
 *     G 投稿予定日/H ステータス(手動)/I 修正対象/J 修正締切/K 制作担当/L 制作締切/
 *     M サムネ担当/N サムネ締切/O CL提出締切/P 提出稿/Q 公開設定・納品締切/R BO担当/
 *     S メモ/T 現在担当(隠)/U 現在締切(隠)/V 信号(隠)
 *   - Hのチップ色CF(11種・旧シート風)＋行のボール色(単一列ルール×11列)
 *   - 注意: フィルタ条件を全列解除したため、以前マスターで非表示だった
 *     完了以外の行(旧フィルタの別列条件で隠れていた約34行)が表示されるようになった
 * ======================================================================== */
