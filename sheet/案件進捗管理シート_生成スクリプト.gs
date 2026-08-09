/**
 * 案件進捗管理シート（新）一括生成スクリプト
 * -------------------------------------------------
 * 実行すると、Google Drive に新しいスプレッドシートを1つ作成します。
 * 既存の本番シートには一切触れません。
 *
 * 作られるシートは3枚:
 *   1.「マスター」   … 1行=1動画。工程ごとの担当・締切・完了チェック。現在ステータスは自動。
 *   2.「担当者ビュー」… 名前を選ぶと、その人が今ボールを持つ案件だけ自動表示。
 *   3.「設定」       … クライアント別リードタイム / メンバー名簿。ここの数字を変えると締切が変わる。
 *
 * 工程フロー: 制作 → サムネ → チェック → CL提出 → 完了
 *   ＋ 差し戻し（チーム修正 / CL修正）が入ったら、現在ステータスが「修正中」に切替わり、
 *     担当は制作担当に戻る。直したらドロップダウンを空に戻せば元のフローに復帰。
 *
 * 使い方:
 *   script.google.com →「新しいプロジェクト」→ この内容を全部貼り付け →
 *   関数「buildSheet」を選んで実行 → 初回だけ権限を承認 →
 *   実行ログ（Ctrl+Enter で表示）に出るURLを開く。
 */

var NUM_ROWS = 200; // 入力できる行数（足りなくなったら下にコピーで増やせます）

function buildSheet() {
  var ss = SpreadsheetApp.create('案件進捗管理シート（新）');

  var setup  = ss.getActiveSheet().setName('設定');
  var master = ss.insertSheet('マスター');
  var view   = ss.insertSheet('担当者ビュー');

  buildSetup_(setup);
  buildMaster_(master);
  buildPersonView_(view);

  // 並び順: マスター → 担当者ビュー → 設定
  ss.setActiveSheet(master);
  ss.moveActiveSheet(1);
  ss.setActiveSheet(view);
  ss.moveActiveSheet(2);

  Logger.log('完成しました。次のURLを開いてください:\n' + ss.getUrl());
}

/* ========== 設定シート ========== */
function buildSetup_(sh) {
  sh.clear();

  // リードタイム表（投稿予定日の何日前を締切にするか）
  var lead = [
    ['クライアント', '制作(日前)', 'サムネ(日前)', 'チェック(日前)', 'CL提出(日前)'],
    ['（標準）', 10, 8, 6, 5],
    ['DEP',     10, 8, 6, 5],
    ['1sec.',   10, 8, 6, 5],
    ['そうぞう', 10, 8, 6, 5],
    ['モームリ', 10, 8, 6, 5],
    ['mug',     10, 8, 6, 5],
    ['Natuul',  10, 8, 6, 5]
  ];
  sh.getRange(1, 1, lead.length, 5).setValues(lead);

  // メンバー名簿（担当者ドロップダウンの選択肢）
  var members = [['メンバー'], ['ITARU'], ['増田'], ['岩渕'], ['GS'], ['']];
  sh.getRange(1, 7, members.length, 1).setValues(members);

  sh.getRange('A1:E1').setFontWeight('bold').setBackground('#f1efe8');
  sh.getRange('G1').setFontWeight('bold').setBackground('#f1efe8');
  sh.setColumnWidth(1, 110);
  sh.setFrozenRows(1);
  sh.getRange('A2').setNote('「（標準）」は、ここに載っていないクライアントに使う共通の日数です。\nクライアント名を追加すれば、個別の締切日数を設定できます。');
}

/* ========== マスターシート ==========
 * 列レイアウト:
 *  A クライアント   B 動画タイトル   C 完成尺   D 投稿予定日   E 現在ステータス(自動)
 *  F 差し戻し       G 制作担当   H 制作締切(自動)   I 制作✓
 *  J サムネ担当     K サムネ締切(自動)   L サムネ✓
 *  M チェック担当   N チェック締切(自動)   O チェック✓
 *  P CL提出締切(自動)   Q CL提出✓
 *  R BO担当   S メモ
 *  T 現在担当(自動・非表示)   U 現在締切(自動・非表示)
 */
function buildMaster_(sh) {
  sh.clear();

  var headers = [
    'クライアント','動画タイトル','完成尺','投稿予定日','現在ステータス',
    '差し戻し','制作担当','制作締切','制作✓',
    'サムネ担当','サムネ締切','サムネ✓',
    'チェック担当','チェック締切','チェック✓',
    'CL提出締切','CL提出✓',
    'BO担当','メモ',
    '現在担当','現在締切'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f1efe8')
    .setVerticalAlignment('middle').setWrap(true);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);

  var last = NUM_ROWS + 1;

  // 完了チェックボックス（制作I・サムネL・チェックO・CL提出Q）
  ['I','L','O','Q'].forEach(function(col){
    sh.getRange(col + '2:' + col + last).insertCheckboxes();
  });

  // チェック担当の初期値 = ITARU（M列）
  var itaru = [];
  for (var i = 0; i < NUM_ROWS; i++) itaru.push(['ITARU']);
  sh.getRange('M2:M' + last).setValues(itaru);

  // --- 行2に数式をセット ---
  // 各工程の締切（投稿予定日 D から逆算）
  sh.getRange('H2').setFormula(
    '=IF($D2="","",$D2-IFERROR(VLOOKUP($A2,\'設定\'!$A:$E,2,FALSE),VLOOKUP("（標準）",\'設定\'!$A:$E,2,FALSE)))');
  sh.getRange('K2').setFormula(
    '=IF($D2="","",$D2-IFERROR(VLOOKUP($A2,\'設定\'!$A:$E,3,FALSE),VLOOKUP("（標準）",\'設定\'!$A:$E,3,FALSE)))');
  sh.getRange('N2').setFormula(
    '=IF($D2="","",$D2-IFERROR(VLOOKUP($A2,\'設定\'!$A:$E,4,FALSE),VLOOKUP("（標準）",\'設定\'!$A:$E,4,FALSE)))');
  sh.getRange('P2').setFormula(
    '=IF($D2="","",$D2-IFERROR(VLOOKUP($A2,\'設定\'!$A:$E,5,FALSE),VLOOKUP("（標準）",\'設定\'!$A:$E,5,FALSE)))');

  // 現在ステータス E（差し戻しが最優先 → そのあとチェックの直列）
  sh.getRange('E2').setFormula(
    '=IF($D2="","",IF($F2<>"","修正中（"&$F2&"）",IFS(' +
    '$Q2=TRUE,"完了",' +
    '$O2=TRUE,"CL提出待ち"&IF($P2<TODAY(),"（超過）",""),' +
    '$L2=TRUE,"チェック待ち"&IF($N2<TODAY(),"（超過）",""),' +
    '$I2=TRUE,"サムネ待ち"&IF($K2<TODAY(),"（超過）",""),' +
    'TRUE,"制作待ち"&IF($H2<TODAY(),"（超過）",""))))');

  // 現在担当 T（修正中は制作担当に戻る）
  sh.getRange('T2').setFormula(
    '=IF($D2="","",IF($F2<>"",$G2,IFS($I2<>TRUE,$G2,$L2<>TRUE,$J2,$O2<>TRUE,$M2,$Q2<>TRUE,$R2,TRUE,"—")))');

  // 現在締切 U（修正中は空欄＝メモで管理）
  sh.getRange('U2').setFormula(
    '=IF($D2="","",IF($F2<>"","",IFS($I2<>TRUE,$H2,$L2<>TRUE,$K2,$O2<>TRUE,$N2,$Q2<>TRUE,$P2,TRUE,"")))');

  // 行2の数式を最終行までコピー
  ['E','H','K','N','P','T','U'].forEach(function(col){
    sh.getRange(col + '2').copyTo(sh.getRange(col + '3:' + col + last));
  });

  // 日付の表示形式
  ['D','H','K','N','P','U'].forEach(function(col){
    sh.getRange(col + '2:' + col + last).setNumberFormat('m/d');
  });

  // --- 入力規則（ドロップダウン）---
  var ssRef = sh.getParent();
  var clientRange = ssRef.getRange('設定!A2:A100');
  var memberRange = ssRef.getRange('設定!G2:G50');
  var clientRule = SpreadsheetApp.newDataValidation().requireValueInRange(clientRange, true).setAllowInvalid(true).build();
  var memberRule = SpreadsheetApp.newDataValidation().requireValueInRange(memberRange, true).setAllowInvalid(true).build();
  var sashiRule  = SpreadsheetApp.newDataValidation().requireValueInList(['チーム修正','CL修正'], true).setAllowInvalid(true).build();

  sh.getRange('A2:A' + last).setDataValidation(clientRule);
  ['G','J','M','R'].forEach(function(col){
    sh.getRange(col + '2:' + col + last).setDataValidation(memberRule);
  });
  sh.getRange('F2:F' + last).setDataValidation(sashiRule);

  // --- 条件付き書式 ---
  applyMasterFormats_(sh, last);

  // 裏方列を非表示
  sh.hideColumns(20, 2); // T,U

  // 列幅
  sh.setColumnWidth(1, 90);   // クライアント
  sh.setColumnWidth(2, 240);  // 動画タイトル
  sh.setColumnWidth(3, 55);   // 完成尺
  sh.setColumnWidth(4, 70);   // 投稿予定日
  sh.setColumnWidth(5, 120);  // 現在ステータス
  sh.setColumnWidth(6, 90);   // 差し戻し
  [7,10,13,18].forEach(function(c){ sh.setColumnWidth(c, 70); });  // 担当列
  [8,11,14,16].forEach(function(c){ sh.setColumnWidth(c, 60); });  // 締切列
  [9,12,15,17].forEach(function(c){ sh.setColumnWidth(c, 38); });  // ✓列
  sh.setColumnWidth(19, 180); // メモ
}

function applyMasterFormats_(sh, last) {
  var rules = [];
  function rule(rangeA1, formula, bg, font) {
    var b = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setRanges([sh.getRange(rangeA1)]);
    if (bg) b.setBackground(bg);
    if (font) b.setFontColor(font);
    rules.push(b.build());
  }

  // 現在ステータス E：修正中=オレンジ → 超過=赤 → 完了=グレー → 進行中=青
  rule('E2:E' + last, '=REGEXMATCH($E2&"","修正中")', '#faeeda', '#854f0b');
  rule('E2:E' + last, '=REGEXMATCH($E2&"","超過")',   '#fcebeb', '#a32d2d');
  rule('E2:E' + last, '=$E2="完了"',                  '#f1efe8', '#888780');
  rule('E2:E' + last, '=AND($D2<>"",$E2<>"完了")',     '#e6f1fb', '#185fa5');

  // 差し戻し F：選択中はオレンジで目立たせる
  rule('F2:F' + last, '=$F2<>""', '#faeeda', '#854f0b');

  // 締切セル：超過=赤（最優先）→ 今ここ=青（修正中でないとき）
  // 制作 H
  rule('H2:H' + last, '=AND($I2<>TRUE,$H2<>"",$H2<TODAY())', '#fcebeb', '#a32d2d');
  rule('H2:H' + last, '=AND($I2<>TRUE,$F2="",$H2<>"",$H2>=TODAY())', '#e6f1fb', '#185fa5');
  // サムネ K
  rule('K2:K' + last, '=AND($L2<>TRUE,$K2<>"",$K2<TODAY())', '#fcebeb', '#a32d2d');
  rule('K2:K' + last, '=AND($I2=TRUE,$L2<>TRUE,$F2="",$K2<>"",$K2>=TODAY())', '#e6f1fb', '#185fa5');
  // チェック N
  rule('N2:N' + last, '=AND($O2<>TRUE,$N2<>"",$N2<TODAY())', '#fcebeb', '#a32d2d');
  rule('N2:N' + last, '=AND($L2=TRUE,$O2<>TRUE,$F2="",$N2<>"",$N2>=TODAY())', '#e6f1fb', '#185fa5');
  // CL提出 P
  rule('P2:P' + last, '=AND($Q2<>TRUE,$P2<>"",$P2<TODAY())', '#fcebeb', '#a32d2d');
  rule('P2:P' + last, '=AND($O2=TRUE,$Q2<>TRUE,$F2="",$P2<>"",$P2>=TODAY())', '#e6f1fb', '#185fa5');

  // 完了行うっすらグレー（タイトル列）
  rule('B2:B' + last, '=$Q2=TRUE', '#f7f6f2', '#888780');

  sh.setConditionalFormatRules(rules);
}

/* ========== 担当者ビュー ========== */
function buildPersonView_(sh) {
  sh.clear();
  sh.getRange('A1').setValue('担当者を選択 →').setFontWeight('bold');

  var memberRange = sh.getParent().getRange('設定!G2:G50');
  var memberRule = SpreadsheetApp.newDataValidation().requireValueInRange(memberRange, true).setAllowInvalid(true).build();
  sh.getRange('B1').setDataValidation(memberRule).setValue('ITARU').setFontWeight('bold').setBackground('#e6f1fb');

  sh.getRange('A3:E3').setValues([['クライアント','動画タイトル','今やること','締切','投稿予定']])
    .setFontWeight('bold').setBackground('#f1efe8');

  // 選んだ人が今ボールを持つ案件だけを締切が近い順に表示
  sh.getRange('A4').setFormula(
    '=IFERROR(SORT(FILTER({マスター!A2:B,マスター!E2:E,マスター!U2:U,マスター!D2:D},' +
    'マスター!T2:T=$B$1,マスター!D2:D<>""),4,TRUE),"該当なし")');

  sh.getRange('D4:D' + (NUM_ROWS+3)).setNumberFormat('m/d');
  sh.getRange('E4:E' + (NUM_ROWS+3)).setNumberFormat('m/d');
  sh.setColumnWidth(1, 100);
  sh.setColumnWidth(2, 260);
  sh.setColumnWidth(3, 130);
  sh.setColumnWidth(4, 70);
  sh.setColumnWidth(5, 70);
  sh.setFrozenRows(3);
}

/* ========== 旧シート→新マスター 移行（投稿日が今日以降の未来分＋日付なしの制作中） ==========
 * 一度だけ実行する移行用。旧シートは読むだけ。再実行すると新マスターの移行データをクリアして入れ直す。 */
function migrate(){
  var OLD = '1K0sV4eFKy2DRXEDKqhd4XhDeBQaDe2jAlYlJDOB9CjY';
  var NEW = '1LLkkdljyokHqjLll9rcG-ClSYnXx43D-xZtmLkAp0Ds';
  var src = SpreadsheetApp.openById(OLD).getSheetByName('全体');
  var dst = SpreadsheetApp.openById(NEW).getSheetByName('マスター');
  var lastRow = src.getLastRow();
  if (lastRow < 3) { Logger.log('旧データなし'); return; }
  var today = new Date(); today.setHours(0,0,0,0);
  var INPROD = {'編集中':1,'サムネ待ち':1,'砂田確認中':1,'CL確認中':1,'修正中':1,'クライアント共有待ち':1,'タイトル待ち':1};
  var vals = src.getRange(3, 2, lastRow - 2, 8).getValues(); // B..I
  function pdate(x){
    if (Object.prototype.toString.call(x) === '[object Date]') return x;
    var p = String(x).split('/');
    if (p.length === 2){ var mo = parseInt(p[0],10), da = parseInt(p[1],10);
      if (mo>=1 && mo<=12 && da>=1 && da<=31) return new Date(2026, mo-1, da); }
    return '';
  }
  function mapS(s){
    if (s==='サムネ待ち')            return [true,false,false,false,''];
    if (s==='砂田確認中')            return [true,true,false,false,''];
    if (s==='クライアント共有待ち')  return [true,true,true,false,''];
    if (s==='CL確認中')              return [true,true,true,false,''];
    if (s==='修正中')                return [true,false,false,false,'チーム修正'];
    return [false,false,false,false,''];
  }
  var rows = [];
  for (var i=0;i<vals.length;i++){
    var client=vals[i][0], dRaw=vals[i][2], title=vals[i][3], stat=vals[i][5], editor=vals[i][6], bo=vals[i][7];
    if (!client) continue;
    if (stat==='完了' || stat==='' || stat==null) continue;
    var date = pdate(dRaw);
    var keep = (date && date.getTime() >= today.getTime()) || (!date && INPROD[stat]);
    if (!keep) continue;
    var mm = mapS(stat);
    rows.push({client:client,title:title,date:date,sashi:mm[4],s1:mm[0],s2:mm[1],s3:mm[2],s4:mm[3],editor:editor,bo:bo});
  }
  rows.sort(function(a,b){ var av=a.date?a.date.getTime():9e15, bv=b.date?b.date.getTime():9e15; return av-bv; });
  var n = rows.length;
  // 既存の移行データをクリア（数式列 E,H,K,N,P,T,U と M=ITARU は触らない）
  var clearCols = [1,2,3,4,6,7,9,10,12,15,17,18,19];
  for (var c=0;c<clearCols.length;c++){ dst.getRange(2, clearCols[c], NUM_ROWS, 1).clearContent(); }
  if (n===0){ Logger.log('対象0件'); return; }
  var A=[],B=[],D=[],F=[],G=[],I=[],L=[],O=[],Q=[],R=[];
  for (var j=0;j<n;j++){
    A.push([rows[j].client]); B.push([rows[j].title]); D.push([rows[j].date]);
    F.push([rows[j].sashi]); G.push([rows[j].editor]); R.push([rows[j].bo]);
    I.push([rows[j].s1]); L.push([rows[j].s2]); O.push([rows[j].s3]); Q.push([rows[j].s4]);
  }
  dst.getRange(2,1,n,1).setValues(A);
  dst.getRange(2,2,n,1).setValues(B);
  dst.getRange(2,4,n,1).setValues(D);
  dst.getRange(2,6,n,1).setValues(F);
  dst.getRange(2,7,n,1).setValues(G);
  dst.getRange(2,9,n,1).setValues(I);
  dst.getRange(2,12,n,1).setValues(L);
  dst.getRange(2,15,n,1).setValues(O);
  dst.getRange(2,17,n,1).setValues(Q);
  dst.getRange(2,18,n,1).setValues(R);
  Logger.log('移行完了: ' + n + '件');
}

/* ========== 担当者名簿に旧シートの編集者名を追加 ========== */
function addMembers(){
  var st = SpreadsheetApp.openById('1LLkkdljyokHqjLll9rcG-ClSYnXx43D-xZtmLkAp0Ds').getSheetByName('設定');
  var vals = st.getRange(2,7,49,1).getValues();
  var cur=[], lastIdx=-1;
  for (var i=0;i<vals.length;i++){ var x=String(vals[i][0]).trim(); if(x){ cur.push(x); lastIdx=i; } }
  var add = ['うえさん','かずあき','ゆき','イカラシ','ヒデキ','中畑','安里','森下','榎本'];
  var out=[]; add.forEach(function(n){ if(cur.indexOf(n)<0) out.push([n]); });
  if (out.length){ st.getRange(2+lastIdx+1,7,out.length,1).setValues(out); }
  Logger.log('追加: ' + out.length + '名');
}

/* ========== V2更新：クライアント別ビュー追加＋担当者ビューをBO別に ==========
 * 既存シートに対して一度だけ実行。マスターの列構成・入力データは一切変更しない。 */
function applyV2(){
  var ss = SpreadsheetApp.openById('1LLkkdljyokHqjLll9rcG-ClSYnXx43D-xZtmLkAp0Ds');

  // ビューを作り直し（マスターには触れない）
  var bo = ss.getSheetByName('担当者ビュー');
  if (bo) { bo.setName('BO別ビュー'); } else { bo = ss.getSheetByName('BO別ビュー') || ss.insertSheet('BO別ビュー'); }
  buildBOView_(bo);
  var cl = ss.getSheetByName('クライアント別ビュー') || ss.insertSheet('クライアント別ビュー');
  buildClientView_(cl);

  // 並び: マスター → クライアント別 → BO別 → 設定
  ss.setActiveSheet(cl); ss.moveActiveSheet(2);
  ss.setActiveSheet(bo); ss.moveActiveSheet(3);
  SpreadsheetApp.flush();
  Logger.log('V2適用完了');
}

/* BO別ビュー：増田/岩渕 を選ぶと、その人がBO担当の案件だけを締切が近い順に表示 */
function buildBOView_(sh){
  sh.clear();
  sh.getRange('A1').setValue('BO担当を選択 →').setFontWeight('bold');
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(['増田','岩渕'], true).setAllowInvalid(true).build();
  sh.getRange('B1').setDataValidation(rule).setValue('増田').setFontWeight('bold').setBackground('#e6f1fb');
  sh.getRange('A3:E3').setValues([['クライアント','動画タイトル','状況','締切','投稿予定']])
    .setFontWeight('bold').setBackground('#f1efe8');
  sh.getRange('A4').setFormula(
    '=IFERROR(SORT(FILTER({マスター!A2:B,マスター!E2:E,マスター!U2:U,マスター!D2:D},' +
    'マスター!R2:R=$B$1,マスター!A2:A<>""),4,TRUE),"該当なし")');
  sh.getRange('D4:D' + (NUM_ROWS+3)).setNumberFormat('m/d');
  sh.getRange('E4:E' + (NUM_ROWS+3)).setNumberFormat('m/d');
  sh.setColumnWidth(1,100); sh.setColumnWidth(2,260); sh.setColumnWidth(3,130); sh.setColumnWidth(4,70); sh.setColumnWidth(5,70);
  sh.setFrozenRows(3);
}

/* クライアント別ビュー：クライアントを選ぶと、その案件だけを投稿予定日順に表示 */
function buildClientView_(sh){
  sh.clear();
  sh.getRange('A1').setValue('クライアントを選択 →').setFontWeight('bold');
  var clientRange = sh.getParent().getRange('設定!A2:A100');
  var rule = SpreadsheetApp.newDataValidation().requireValueInRange(clientRange, true).setAllowInvalid(true).build();
  sh.getRange('B1').setDataValidation(rule).setFontWeight('bold').setBackground('#e6f1fb');
  sh.getRange('A3:E3').setValues([['動画タイトル','現在ステータス','BO担当','締切','投稿予定']])
    .setFontWeight('bold').setBackground('#f1efe8');
  sh.getRange('A4').setFormula(
    '=IFERROR(SORT(FILTER({マスター!B2:B,マスター!E2:E,マスター!R2:R,マスター!U2:U,マスター!D2:D},' +
    'マスター!A2:A=$B$1,マスター!A2:A<>""),5,TRUE),"該当なし")');
  sh.getRange('D4:D' + (NUM_ROWS+3)).setNumberFormat('m/d');
  sh.getRange('E4:E' + (NUM_ROWS+3)).setNumberFormat('m/d');
  sh.setColumnWidth(1,260); sh.setColumnWidth(2,130); sh.setColumnWidth(3,70); sh.setColumnWidth(4,70); sh.setColumnWidth(5,70);
  sh.setFrozenRows(3);
}

/* ========== クライアント名のクリーンアップ：末尾「_m」を除去＋全クライアントを設定名簿へ ========== */
function cleanupClients(){
  var ss = SpreadsheetApp.openById('1LLkkdljyokHqjLll9rcG-ClSYnXx43D-xZtmLkAp0Ds');
  var ms = ss.getSheetByName('マスター');
  var st = ss.getSheetByName('設定');
  // 1) マスターA列：末尾"_m"を除去
  var rng = ms.getRange(2, 1, NUM_ROWS, 1);
  var v = rng.getValues();
  for (var i=0;i<v.length;i++){
    var x = String(v[i][0]);
    if (x && x.length>2 && x.slice(-2)==='_m') v[i][0] = x.slice(0,-2);
  }
  rng.setValues(v);
  // 2) マスターに出てくる全クライアントを収集
  var set={}, list=[];
  for (var i=0;i<v.length;i++){ var x=v[i][0]; if(x && !set[x]){ set[x]=1; list.push(x); } }
  list.sort();
  // 3) 設定シートの既存クライアント
  var sv = st.getRange(2,1,98,1).getValues();
  var cur={}, lastRow=1;
  for (var i=0;i<sv.length;i++){ var c=String(sv[i][0]).trim(); if(c){ cur[c]=1; lastRow=2+i; } }
  // 4) 未登録クライアントを 標準リードタイム(10/8/6/5) で追加
  var add=[];
  list.forEach(function(c){ if(!cur[c]) add.push([c,10,8,6,5]); });
  if (add.length){ st.getRange(lastRow+1,1,add.length,5).setValues(add); }
  Logger.log('クライアント名クリーンアップ完了。設定に追加: ' + add.length + '社');
}

/* ========== 再調整：cleanupClientsの_m除去を取り消し、ITARUさんの「_m」命名に揃える ==========
 * 1) マスター：_mありのメイン名に戻す（設定の_mロスターを正とする）
 * 2) 設定：_m版と重複する素の名前の行を削除（リードタイムは保持、メンバー列Gは触らない）
 * 3) 設定：マスターに在るのに未登録のクライアントを追加 */
function reconcileClients(){
  var ss = SpreadsheetApp.openById('1LLkkdljyokHqjLll9rcG-ClSYnXx43D-xZtmLkAp0Ds');
  var ms = ss.getSheetByName('マスター');
  var st = ss.getSheetByName('設定');
  var sLast = st.getLastRow();
  var sd = st.getRange(2,1,sLast-1,5).getValues();
  var nameSet={}, baseHasM={};
  for (var i=0;i<sd.length;i++){ var n=String(sd[i][0]).trim(); if(!n) continue; nameSet[n]=1; if(n.length>2 && n.slice(-2)==='_m') baseHasM[n.slice(0,-2)]=1; }
  // 1) マスター：素の名前→_m付きに戻す
  var mr = ms.getRange(2,1,NUM_ROWS,1);
  var mv = mr.getValues();
  for (var i=0;i<mv.length;i++){ var x=String(mv[i][0]); if(x && !(x.length>2 && x.slice(-2)==='_m') && baseHasM[x]) mv[i][0]=x+'_m'; }
  mr.setValues(mv);
  // 2) 設定：_m版と重複する素の名前を落として再構成（リードタイム保持）
  var kept=[];
  for (var i=0;i<sd.length;i++){ var n=String(sd[i][0]).trim(); if(!n) continue;
    var dup = !(n.length>2 && n.slice(-2)==='_m') && nameSet[n+'_m'];
    if(!dup) kept.push([sd[i][0], sd[i][1]||10, sd[i][2]||8, sd[i][3]||6, sd[i][4]||5]); }
  // 3) マスターに在るのに未登録のクライアントを追加
  var keptSet={}; kept.forEach(function(r){ keptSet[String(r[0]).trim()]=1; });
  var seen={};
  for (var i=0;i<mv.length;i++){ var x=String(mv[i][0]).trim(); if(x && !seen[x]){ seen[x]=1; if(!keptSet[x]){ kept.push([x,10,8,6,5]); keptSet[x]=1; } } }
  // 書き戻し（A:E のみ。メンバー列Gは触らない）
  st.getRange(2,1,sd.length,5).clearContent();
  st.getRange(2,1,kept.length,5).setValues(kept);
  Logger.log('再調整完了。設定クライアント数: ' + kept.length);
}

/* ========== ビュー再構築（ヘッダー修復用） ========== */
function fixViews(){
  var ss = SpreadsheetApp.openById('1LLkkdljyokHqjLll9rcG-ClSYnXx43D-xZtmLkAp0Ds');
  buildBOView_(ss.getSheetByName('BO別ビュー'));
  buildClientView_(ss.getSheetByName('クライアント別ビュー'));
  Logger.log('views rebuilt');
}

/* ========== 進捗信号ダッシュボード追加 ==========
 * マスターに信号列(V)を足し、ダッシュボードタブを先頭に作る。一度だけ実行。 */
function addDashboard(){
  var ss = SpreadsheetApp.openById('1LLkkdljyokHqjLll9rcG-ClSYnXx43D-xZtmLkAp0Ds');
  var ms = ss.getSheetByName('マスター');
  var last = NUM_ROWS + 1;
  // 信号列 V(22)：灰=日付なし / 赤=公開3日以内 or 工程遅延(要対応) / 黄=公開7日以内かつ遅れてない(もうすぐ) / 青=それ以外
  ms.getRange('V1').setValue('信号');
  ms.getRange('V2').setFormula(
    '=IF($A2="","",IF($D2="","灰",IF($Q2=TRUE,"青",IF(OR($D2<=TODAY()+3,AND($U2<>"",$U2<TODAY())),"赤",IF($D2<=TODAY()+7,"黄","青")))))');
  ms.getRange('V2').copyTo(ms.getRange('V3:V' + last));
  ms.hideColumns(22);
  var dh = ss.getSheetByName('ダッシュボード') || ss.insertSheet('ダッシュボード');
  buildDashboard_(dh);
  ss.setActiveSheet(dh); ss.moveActiveSheet(1);
  Logger.log('ダッシュボード追加完了');
}

function buildDashboard_(sh){
  sh.clear();
  sh.setConditionalFormatRules([]);
  var M = 'マスター!';
  sh.getRange('A1').setValue('iMuseLLC 案件 進捗信号ダッシュボード').setFontSize(18).setFontWeight('bold');
  sh.getRange('A2').setFormula(
    '="本日 "&TEXT(TODAY(),"yyyy/mm/dd (ddd)")&"（Asia/Tokyo）　／　信号は 状況(E列) と 公開予定日(D列) から自動判定"');
  sh.getRange('A2').setFontColor('#888780');
  sh.getRange('A4').setValue('サマリ').setFontWeight('bold').setFontSize(13);
  sh.getRange('A4').setNote('信号の判定基準\n灰＝公開予定日(D列)が空欄＝情報不足\n赤＝公開予定まで3日以内、または今の工程の締切を過ぎている＋未完了＝要対応\n黄＝公開予定まで7日以内かつ遅れていない＋未完了＝もうすぐ\n青＝それ以外＝順調');
  sh.getRange('A5').setValue('🔴 要対応（赤）');
  sh.getRange('B5').setValue('🟡 もうすぐ（黄）');
  sh.getRange('C5').setValue('🔵 順調（青）');
  sh.getRange('D5').setValue('⚪ 情報不足（灰）');
  sh.getRange('E5').setValue('合計');
  sh.getRange('A6').setFormula('=COUNTIF(' + M + 'V2:V,"赤")');
  sh.getRange('B6').setFormula('=COUNTIF(' + M + 'V2:V,"黄")');
  sh.getRange('C6').setFormula('=COUNTIF(' + M + 'V2:V,"青")');
  sh.getRange('D6').setFormula('=COUNTIF(' + M + 'V2:V,"灰")');
  sh.getRange('E6').setFormula('=COUNTIF(' + M + 'A2:A,"<>")');
  sh.getRange('A5:E5').setHorizontalAlignment('center').setFontColor('#5f5e5a');
  sh.getRange('A6:E6').setFontSize(22).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('A6').setBackground('#fcebeb').setFontColor('#a32d2d');
  sh.getRange('B6').setBackground('#faeeda').setFontColor('#854f0b');
  sh.getRange('C6').setBackground('#e6f1fb').setFontColor('#185fa5');
  sh.getRange('D6').setBackground('#f1efe8').setFontColor('#5f5e5a');
  sh.getRange('A8').setValue('いますぐ確認が必要な案件（赤・黄）').setFontWeight('bold').setFontSize(13);
  sh.getRange('A9:H9').setValues([['信号','クライアント','タイトル','公開予定','残り(日)','状況','編集','BO']])
    .setFontWeight('bold').setBackground('#f1efe8');
  sh.getRange('A10').setFormula(
    '=IFERROR(SORT(FILTER({' + M + 'V2:V,' + M + 'A2:A,' + M + 'B2:B,' + M + 'D2:D,' + M + 'D2:D-TODAY(),' + M + 'E2:E,' + M + 'G2:G,' + M + 'R2:R},' +
    '(' + M + 'V2:V="赤")+(' + M + 'V2:V="黄")),4,TRUE),"いま確認が必要な案件はありません")');
  sh.getRange('D10:D' + NUM_ROWS).setNumberFormat('m/d');
  sh.setColumnWidth(1,60); sh.setColumnWidth(2,150); sh.setColumnWidth(3,300); sh.setColumnWidth(4,80);
  sh.setColumnWidth(5,70); sh.setColumnWidth(6,130); sh.setColumnWidth(7,90); sh.setColumnWidth(8,90);
  sh.setFrozenRows(9);
  var rules=[];
  function sig(t,bg,fc){ rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(t).setBackground(bg).setFontColor(fc).setRanges([sh.getRange('A10:A' + NUM_ROWS)]).build()); }
  sig('赤','#fcebeb','#a32d2d'); sig('黄','#faeeda','#854f0b'); sig('青','#e6f1fb','#185fa5'); sig('灰','#f1efe8','#5f5e5a');
  sh.setConditionalFormatRules(rules);
}

/* ========== 信号ルール（最終版）：赤=公開3日以内 or 工程遅延(要対応) / 黄=公開7日以内かつ遅れてない(もうすぐ) ========== */
function updateSignals(){
  var ss = SpreadsheetApp.openById('1LLkkdljyokHqjLll9rcG-ClSYnXx43D-xZtmLkAp0Ds');
  var ms = ss.getSheetByName('マスター');
  var last = NUM_ROWS + 1;
  ms.getRange('V2').setFormula(
    '=IF($A2="","",IF($D2="","灰",IF($Q2=TRUE,"青",IF(OR($D2<=TODAY()+3,AND($U2<>"",$U2<TODAY())),"赤",IF($D2<=TODAY()+7,"黄","青")))))');
  ms.getRange('V2').copyTo(ms.getRange('V3:V' + last));
  var dh = ss.getSheetByName('ダッシュボード');
  if (dh){
    dh.getRange('A5').setValue('🔴 要対応（赤）');
    dh.getRange('B5').setValue('🟡 もうすぐ（黄）');
    dh.getRange('A4').setNote('信号の判定基準\n灰＝公開予定日(D列)が空欄＝情報不足\n赤＝公開予定まで3日以内、または今の工程の締切を過ぎている＋未完了＝要対応\n黄＝公開予定まで7日以内かつ遅れていない＋未完了＝もうすぐ\n青＝それ以外＝順調');
  }
  Logger.log('信号ルール最終版適用完了');
}

/* ========== ダッシュボード視認性改善＋赤を上に＋マスターの完了非表示 ========== */
function upgradeDashboard(){
  var ss = SpreadsheetApp.openById('1LLkkdljyokHqjLll9rcG-ClSYnXx43D-xZtmLkAp0Ds');
  var ms = ss.getSheetByName('マスター');
  var dh = ss.getSheetByName('ダッシュボード');
  var M = 'マスター!';
  dh.clear();
  dh.setConditionalFormatRules([]);
  dh.getRange(1,1,9,9).breakApart();
  dh.getRange('A1').setValue('iMuseLLC 案件 進捗信号ダッシュボード').setFontSize(18).setFontWeight('bold');
  dh.getRange('A2').setFormula('="本日 "&TEXT(TODAY(),"yyyy/mm/dd (ddd)")&"（Asia/Tokyo）　／　信号は 状況(E列)と公開予定日(D列)から自動判定"');
  dh.getRange('A2').setFontColor('#888780');
  dh.getRange('A4').setValue('サマリ').setFontWeight('bold').setFontSize(13);
  dh.getRange('A4').setNote('信号の判定基準\n灰＝公開予定日が空欄＝情報不足\n赤＝公開3日以内 または 工程の締切超過＋未完了＝要対応\n黄＝公開7日以内かつ遅れてない＋未完了＝もうすぐ\n青＝それ以外＝順調');
  dh.getRange('A5:E5').setValues([['🔴 要対応','🟡 もうすぐ','🔵 順調','⚪ 情報不足','合計']])
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(12).setFontColor('#444441');
  dh.getRange('A6').setFormula('=COUNTIF(' + M + 'V2:V,"赤")');
  dh.getRange('B6').setFormula('=COUNTIF(' + M + 'V2:V,"黄")');
  dh.getRange('C6').setFormula('=COUNTIF(' + M + 'V2:V,"青")');
  dh.getRange('D6').setFormula('=COUNTIF(' + M + 'V2:V,"灰")');
  dh.getRange('E6').setFormula('=COUNTIF(' + M + 'A2:A,"<>")');
  dh.getRange('A6:E6').setFontSize(24).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  dh.getRange('A5:A6').setBackground('#fcebeb'); dh.getRange('A6').setFontColor('#a32d2d');
  dh.getRange('B5:B6').setBackground('#faeeda'); dh.getRange('B6').setFontColor('#854f0b');
  dh.getRange('C5:C6').setBackground('#e6f1fb'); dh.getRange('C6').setFontColor('#185fa5');
  dh.getRange('D5:D6').setBackground('#f1efe8'); dh.getRange('D6').setFontColor('#5f5e5a');
  dh.getRange('E5:E6').setBackground('#f7f6f2'); dh.getRange('E6').setFontColor('#2c2c2a');
  dh.getRange('A5:E6').setBorder(true,true,true,true,true,true,'#ffffff',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  dh.setRowHeight(5,26); dh.setRowHeight(6,52);
  dh.getRange('A8').setValue('いますぐ確認が必要な案件（赤＝要対応 → 黄＝もうすぐ の順）').setFontWeight('bold').setFontSize(13);
  dh.getRange('A9:H9').setValues([['信号','クライアント','タイトル','公開予定','残り(日)','状況','編集','BO']])
    .setFontWeight('bold').setBackground('#f1efe8').setVerticalAlignment('middle');
  dh.getRange('A10').setFormula(
    '=IFERROR(SORT(FILTER({' + M + 'V2:V,' + M + 'A2:A,' + M + 'B2:B,' + M + 'D2:D,' + M + 'D2:D-TODAY(),' + M + 'E2:E,' + M + 'G2:G,' + M + 'R2:R},' +
    '(' + M + 'V2:V="赤")+(' + M + 'V2:V="黄")),1,TRUE,4,TRUE),"いま確認が必要な案件はありません")');
  dh.getRange('D10:D' + NUM_ROWS).setNumberFormat('m/d');
  dh.setColumnWidth(1,70); dh.setColumnWidth(2,150); dh.setColumnWidth(3,300); dh.setColumnWidth(4,90);
  dh.setColumnWidth(5,75); dh.setColumnWidth(6,130); dh.setColumnWidth(7,90); dh.setColumnWidth(8,90);
  dh.setFrozenRows(9);
  var rules=[];
  function sig(t,bg,fc){ rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(t).setBackground(bg).setFontColor(fc).setRanges([dh.getRange('A10:A' + NUM_ROWS)]).build()); }
  sig('赤','#fcebeb','#a32d2d'); sig('黄','#faeeda','#854f0b'); sig('青','#e6f1fb','#185fa5'); sig('灰','#f1efe8','#5f5e5a');
  dh.setConditionalFormatRules(rules);
  // マスター：現在ステータス「完了」を隠すフィルタ（フィルタを開けば再表示できる）
  var ex = ms.getFilter(); if (ex) ex.remove();
  var f = ms.getRange(1,1,NUM_ROWS+1,19).createFilter();
  f.setColumnFilterCriteria(5, SpreadsheetApp.newFilterCriteria().setHiddenValues(['完了']).build());
  Logger.log('ダッシュボード改善＋完了非表示フィルタ 完了');
}

/* ================== V5 ==================
 * 1) CL提出の後に「公開設定・納品」工程を追加（完了＝ここまで済んだら）
 * 2) クライアント別/BO別ビューを色付き
 * 3) ダッシュボードに情報不足の詳細（確認リストとは別の場所＝右側）
 * 挿入後の列: ...P CL提出締切(16) Q CL提出✓(17) R 公開設定納品締切(18) S 公開設定納品✓(19)
 *            T BO担当(20) U メモ(21) V 現在担当(22) W 現在締切(23) X 信号(24) */
var SS_ID = '1LLkkdljyokHqjLll9rcG-ClSYnXx43D-xZtmLkAp0Ds';

function applyV5(){
  var ss = SpreadsheetApp.openById(SS_ID);
  addPublishStage_(ss);
  rebuildBOView_(ss.getSheetByName('BO別ビュー'));
  rebuildClientView_(ss.getSheetByName('クライアント別ビュー'));
  rebuildDashboardGray_(ss);
  Logger.log('V5適用完了');
}

function addPublishStage_(ss){
  var ms = ss.getSheetByName('マスター');
  var last = NUM_ROWS + 1;
  // CL提出✓(17)の後ろに2列挿入 → R=公開設定納品締切, S=公開設定納品✓
  if (ms.getRange('R1').getValue() !== '公開設定・納品締切') {
    ms.insertColumnsAfter(17, 2);
  }
  ms.getRange('R1').setValue('公開設定・納品締切');
  ms.getRange('S1').setValue('公開設定・納品✓');
  ms.getRange(1,1,1,24).setFontWeight('bold').setBackground('#f1efe8').setVerticalAlignment('middle').setWrap(true);
  ms.getRange('R2').setFormula('=IF($D2="","",$D2)'); // 締切=投稿予定日（公開日当日）
  ms.getRange('R2').copyTo(ms.getRange('R3:R' + last));
  ms.getRange('R2:R' + last).setNumberFormat('m/d');
  ms.getRange('S2:S' + last).insertCheckboxes();
  // 現在ステータス E：完了＝公開設定納品✓(S)
  ms.getRange('E2').setFormula('=IF($A2="","",IF($F2<>"","修正中（"&$F2&"）",IFS(' +
    '$S2=TRUE,"完了",' +
    '$Q2=TRUE,"公開設定・納品待ち"&IF($R2<TODAY(),"（超過）",""),' +
    '$O2=TRUE,"CL提出待ち"&IF($P2<TODAY(),"（超過）",""),' +
    '$L2=TRUE,"チェック待ち"&IF($N2<TODAY(),"（超過）",""),' +
    '$I2=TRUE,"サムネ待ち"&IF($K2<TODAY(),"（超過）",""),' +
    'TRUE,"制作待ち"&IF($H2<TODAY(),"（超過）",""))))');
  ms.getRange('E2').copyTo(ms.getRange('E3:E' + last));
  // 現在担当 V(22)
  ms.getRange('V2').setFormula('=IF($A2="","",IF($F2<>"",$G2,IFS($I2<>TRUE,$G2,$L2<>TRUE,$J2,$O2<>TRUE,$M2,$Q2<>TRUE,$T2,$S2<>TRUE,$T2,TRUE,"—")))');
  ms.getRange('V2').copyTo(ms.getRange('V3:V' + last));
  // 現在締切 W(23)
  ms.getRange('W2').setFormula('=IF($A2="","",IF($F2<>"","",IFS($I2<>TRUE,$H2,$L2<>TRUE,$K2,$O2<>TRUE,$N2,$Q2<>TRUE,$P2,$S2<>TRUE,$R2,TRUE,"")))');
  ms.getRange('W2').copyTo(ms.getRange('W3:W' + last));
  // 信号 X(24)
  ms.getRange('X2').setFormula('=IF($A2="","",IF($D2="","灰",IF($S2=TRUE,"青",IF(OR($D2<=TODAY()+3,AND($W2<>"",$W2<TODAY())),"赤",IF($D2<=TODAY()+7,"黄","青")))))');
  ms.getRange('X2').copyTo(ms.getRange('X3:X' + last));
  ms.setColumnWidth(18,60); ms.setColumnWidth(19,38); // R,S
  ms.showColumns(20,5); ms.hideColumns(22,3); // 念のため: T,U表示 / V,W,X非表示
  rebuildMasterCF_(ms, last);
}

function rebuildMasterCF_(sh, L){
  var rules=[];
  function rule(a1,f,bg,fc){ var b=SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(f).setRanges([sh.getRange(a1)]); if(bg)b.setBackground(bg); if(fc)b.setFontColor(fc); rules.push(b.build()); }
  rule('E2:E'+L,'=REGEXMATCH($E2&"","修正中")','#faeeda','#854f0b');
  rule('E2:E'+L,'=REGEXMATCH($E2&"","超過")','#fcebeb','#a32d2d');
  rule('E2:E'+L,'=$E2="完了"','#f1efe8','#888780');
  rule('E2:E'+L,'=AND($D2<>"",$E2<>"完了")','#e6f1fb','#185fa5');
  rule('F2:F'+L,'=$F2<>""','#faeeda','#854f0b');
  rule('H2:H'+L,'=AND($I2<>TRUE,$H2<>"",$H2<TODAY())','#fcebeb','#a32d2d');
  rule('H2:H'+L,'=AND($I2<>TRUE,$F2="",$H2<>"",$H2>=TODAY())','#e6f1fb','#185fa5');
  rule('K2:K'+L,'=AND($L2<>TRUE,$K2<>"",$K2<TODAY())','#fcebeb','#a32d2d');
  rule('K2:K'+L,'=AND($I2=TRUE,$L2<>TRUE,$F2="",$K2<>"",$K2>=TODAY())','#e6f1fb','#185fa5');
  rule('N2:N'+L,'=AND($O2<>TRUE,$N2<>"",$N2<TODAY())','#fcebeb','#a32d2d');
  rule('N2:N'+L,'=AND($L2=TRUE,$O2<>TRUE,$F2="",$N2<>"",$N2>=TODAY())','#e6f1fb','#185fa5');
  rule('P2:P'+L,'=AND($Q2<>TRUE,$P2<>"",$P2<TODAY())','#fcebeb','#a32d2d');
  rule('P2:P'+L,'=AND($O2=TRUE,$Q2<>TRUE,$F2="",$P2<>"",$P2>=TODAY())','#e6f1fb','#185fa5');
  rule('R2:R'+L,'=AND($S2<>TRUE,$R2<>"",$R2<TODAY())','#fcebeb','#a32d2d');
  rule('R2:R'+L,'=AND($Q2=TRUE,$S2<>TRUE,$F2="",$R2<>"",$R2>=TODAY())','#e6f1fb','#185fa5');
  rule('B2:B'+L,'=$S2=TRUE','#f7f6f2','#888780');
  sh.setConditionalFormatRules(rules);
}

function viewStatusCF_(sh, statusA1, dueA1){
  var rules=[];
  var rng=sh.getRange(statusA1);
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('超過').setBackground('#fcebeb').setFontColor('#a32d2d').setRanges([rng]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('修正中').setBackground('#faeeda').setFontColor('#854f0b').setRanges([rng]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('完了').setBackground('#f1efe8').setFontColor('#888780').setRanges([rng]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('待ち').setBackground('#e6f1fb').setFontColor('#185fa5').setRanges([rng]).build());
  var first=dueA1.split(':')[0];
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND('+first+'<>"",'+first+'<TODAY())').setFontColor('#a32d2d').setRanges([sh.getRange(dueA1)]).build());
  sh.setConditionalFormatRules(rules);
}

function rebuildBOView_(sh){
  sh.clear(); sh.setConditionalFormatRules([]);
  sh.getRange('A1').setValue('BO担当を選択 →').setFontWeight('bold');
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(['増田','岩渕'], true).setAllowInvalid(true).build();
  sh.getRange('B1').setDataValidation(rule).setValue('増田').setFontWeight('bold').setBackground('#e6f1fb');
  sh.getRange('A3:E3').setValues([['クライアント','動画タイトル','状況','締切','投稿予定']]).setFontWeight('bold').setBackground('#f1efe8');
  sh.getRange('A4').setFormula('=IFERROR(SORT(FILTER({マスター!A2:B,マスター!E2:E,マスター!W2:W,マスター!D2:D},マスター!T2:T=$B$1,マスター!A2:A<>""),4,TRUE),"該当なし")');
  sh.getRange('D4:D'+(NUM_ROWS+3)).setNumberFormat('m/d');
  sh.getRange('E4:E'+(NUM_ROWS+3)).setNumberFormat('m/d');
  sh.setColumnWidth(1,100); sh.setColumnWidth(2,260); sh.setColumnWidth(3,140); sh.setColumnWidth(4,70); sh.setColumnWidth(5,70);
  sh.setFrozenRows(3);
  viewStatusCF_(sh,'C4:C'+(NUM_ROWS+3),'D4:D'+(NUM_ROWS+3));
}

function rebuildClientView_(sh){
  sh.clear(); sh.setConditionalFormatRules([]);
  sh.getRange('A1').setValue('クライアントを選択 →').setFontWeight('bold');
  var clientRange = sh.getParent().getRange('設定!A2:A100');
  var rule = SpreadsheetApp.newDataValidation().requireValueInRange(clientRange, true).setAllowInvalid(true).build();
  sh.getRange('B1').setDataValidation(rule).setFontWeight('bold').setBackground('#e6f1fb');
  sh.getRange('A3:E3').setValues([['動画タイトル','現在ステータス','BO担当','締切','投稿予定']]).setFontWeight('bold').setBackground('#f1efe8');
  sh.getRange('A4').setFormula('=IFERROR(SORT(FILTER({マスター!B2:B,マスター!E2:E,マスター!T2:T,マスター!W2:W,マスター!D2:D},マスター!A2:A=$B$1,マスター!A2:A<>""),5,TRUE),"該当なし")');
  sh.getRange('D4:D'+(NUM_ROWS+3)).setNumberFormat('m/d');
  sh.getRange('E4:E'+(NUM_ROWS+3)).setNumberFormat('m/d');
  sh.setColumnWidth(1,260); sh.setColumnWidth(2,140); sh.setColumnWidth(3,80); sh.setColumnWidth(4,70); sh.setColumnWidth(5,70);
  sh.setFrozenRows(3);
  viewStatusCF_(sh,'B4:B'+(NUM_ROWS+3),'D4:D'+(NUM_ROWS+3));
}

function rebuildDashboardGray_(ss){
  var ms = ss.getSheetByName('マスター');
  var dh = ss.getSheetByName('ダッシュボード');
  var M = 'マスター!';
  dh.clear(); dh.setConditionalFormatRules([]); dh.getRange(1,1,9,13).breakApart();
  dh.getRange('A1').setValue('iMuseLLC 案件 進捗信号ダッシュボード').setFontSize(18).setFontWeight('bold');
  dh.getRange('A2').setFormula('="本日 "&TEXT(TODAY(),"yyyy/mm/dd (ddd)")&"（Asia/Tokyo）　／　信号は 状況(E列)と公開予定日(D列)から自動判定"').setFontColor('#888780');
  dh.getRange('A4').setValue('サマリ').setFontWeight('bold').setFontSize(13);
  dh.getRange('A4').setNote('信号の判定基準\n灰＝公開予定日が空欄＝情報不足\n赤＝公開3日以内 または 工程の締切超過＋未完了＝要対応\n黄＝公開7日以内かつ遅れてない＋未完了＝もうすぐ\n青＝それ以外＝順調\n※完了＝公開設定・納品まで済んだもの');
  dh.getRange('A5:E5').setValues([['🔴 要対応','🟡 もうすぐ','🔵 順調','⚪ 情報不足','合計']])
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(12).setFontColor('#444441');
  dh.getRange('A6').setFormula('=COUNTIF(' + M + 'X2:X,"赤")');
  dh.getRange('B6').setFormula('=COUNTIF(' + M + 'X2:X,"黄")');
  dh.getRange('C6').setFormula('=COUNTIF(' + M + 'X2:X,"青")');
  dh.getRange('D6').setFormula('=COUNTIF(' + M + 'X2:X,"灰")');
  dh.getRange('E6').setFormula('=COUNTIF(' + M + 'A2:A,"<>")');
  dh.getRange('A6:E6').setFontSize(24).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  dh.getRange('A5:A6').setBackground('#fcebeb'); dh.getRange('A6').setFontColor('#a32d2d');
  dh.getRange('B5:B6').setBackground('#faeeda'); dh.getRange('B6').setFontColor('#854f0b');
  dh.getRange('C5:C6').setBackground('#e6f1fb'); dh.getRange('C6').setFontColor('#185fa5');
  dh.getRange('D5:D6').setBackground('#f1efe8'); dh.getRange('D6').setFontColor('#5f5e5a');
  dh.getRange('E5:E6').setBackground('#f7f6f2'); dh.getRange('E6').setFontColor('#2c2c2a');
  dh.getRange('A5:E6').setBorder(true,true,true,true,true,true,'#ffffff',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  dh.setRowHeight(5,26); dh.setRowHeight(6,52);
  // 確認が必要な案件（赤→黄）A-H
  dh.getRange('A8').setValue('いますぐ確認が必要な案件（赤＝要対応 → 黄＝もうすぐ の順）').setFontWeight('bold').setFontSize(13);
  dh.getRange('A9:H9').setValues([['信号','クライアント','タイトル','公開予定','残り(日)','状況','編集','BO']]).setFontWeight('bold').setBackground('#f1efe8').setVerticalAlignment('middle');
  dh.getRange('A10').setFormula('=IFERROR(SORT(FILTER({' + M + 'X2:X,' + M + 'A2:A,' + M + 'B2:B,' + M + 'D2:D,' + M + 'D2:D-TODAY(),' + M + 'E2:E,' + M + 'G2:G,' + M + 'T2:T},(' + M + 'X2:X="赤")+(' + M + 'X2:X="黄")),1,TRUE,4,TRUE),"いま確認が必要な案件はありません")');
  dh.getRange('D10:D'+NUM_ROWS).setNumberFormat('m/d');
  // 情報不足（灰）J-M ＝ 確認リストとは別の場所
  dh.getRange('J8').setValue('情報不足の案件（公開予定日が未入力）').setFontWeight('bold').setFontSize(13).setFontColor('#5f5e5a');
  dh.getRange('J9:M9').setValues([['クライアント','タイトル','制作','BO']]).setFontWeight('bold').setBackground('#f1efe8').setVerticalAlignment('middle');
  dh.getRange('J10').setFormula('=IFERROR(FILTER({' + M + 'A2:A,' + M + 'B2:B,' + M + 'G2:G,' + M + 'T2:T},(' + M + 'A2:A<>"")*(' + M + 'D2:D="")),"なし")');
  dh.setColumnWidth(1,70); dh.setColumnWidth(2,150); dh.setColumnWidth(3,300); dh.setColumnWidth(4,90); dh.setColumnWidth(5,75); dh.setColumnWidth(6,130); dh.setColumnWidth(7,90); dh.setColumnWidth(8,90);
  dh.setColumnWidth(9,30); dh.setColumnWidth(10,120); dh.setColumnWidth(11,240); dh.setColumnWidth(12,80); dh.setColumnWidth(13,80);
  dh.setFrozenRows(9);
  var rules=[];
  function sig(t,bg,fc){ rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(t).setBackground(bg).setFontColor(fc).setRanges([dh.getRange('A10:A'+NUM_ROWS)]).build()); }
  sig('赤','#fcebeb','#a32d2d'); sig('黄','#faeeda','#854f0b'); sig('青','#e6f1fb','#185fa5'); sig('灰','#f1efe8','#5f5e5a');
  dh.setConditionalFormatRules(rules);
  // マスターの完了非表示フィルタ（列数が増えたので作り直し）
  var ex = ms.getFilter(); if (ex) ex.remove();
  var f = ms.getRange(1,1,NUM_ROWS+1,21).createFilter();
  f.setColumnFilterCriteria(5, SpreadsheetApp.newFilterCriteria().setHiddenValues(['完了']).build());
}

/* ================== V6 ==================
 * 1) 公開設定・納品締切＝投稿予定日の2日前
 * 2) BO別ビュー→「増田ビュー」「岩渕ビュー」固定タブ（制作担当・サムネ担当も表示）
 * クライアント別ビューは1つのまま（クライアント多数のため） */
function applyV6(){
  var ss = SpreadsheetApp.openById(SS_ID);
  var ms = ss.getSheetByName('マスター');
  var last = NUM_ROWS + 1;
  // 1) 公開設定・納品締切 = 投稿予定日の2日前
  ms.getRange('R2').setFormula('=IF($D2="","",$D2-2)');
  ms.getRange('R2').copyTo(ms.getRange('R3:R' + last));
  // 2) 個人別BOビュー
  var oldBO = ss.getSheetByName('BO別ビュー');
  var masu = ss.getSheetByName('増田ビュー') || ss.insertSheet('増田ビュー');
  buildPersonBOView_(masu, '増田');
  var iwa = ss.getSheetByName('岩渕ビュー') || ss.insertSheet('岩渕ビュー');
  buildPersonBOView_(iwa, '岩渕');
  if (oldBO) ss.deleteSheet(oldBO);
  // タブ順: ダッシュボード, マスター, クライアント別ビュー, 増田ビュー, 岩渕ビュー, 設定
  ss.setActiveSheet(ss.getSheetByName('クライアント別ビュー')); ss.moveActiveSheet(3);
  ss.setActiveSheet(masu); ss.moveActiveSheet(4);
  ss.setActiveSheet(iwa); ss.moveActiveSheet(5);
  Logger.log('V6適用完了');
}

function buildPersonBOView_(sh, name){
  sh.clear(); sh.setConditionalFormatRules([]);
  sh.getRange('A1').setValue(name + 'さんの担当案件（BO）').setFontWeight('bold').setFontSize(13).setFontColor('#185fa5');
  sh.getRange('A3:G3').setValues([['クライアント','動画タイトル','状況','締切','投稿予定','制作担当','サムネ担当']])
    .setFontWeight('bold').setBackground('#f1efe8');
  sh.getRange('A4').setFormula('=IFERROR(SORT(FILTER({マスター!A2:B,マスター!E2:E,マスター!W2:W,マスター!D2:D,マスター!G2:G,マスター!J2:J},マスター!T2:T="' + name + '",マスター!A2:A<>""),4,TRUE),"該当なし")');
  sh.getRange('D4:D' + (NUM_ROWS+3)).setNumberFormat('m/d');
  sh.getRange('E4:E' + (NUM_ROWS+3)).setNumberFormat('m/d');
  sh.setColumnWidth(1,100); sh.setColumnWidth(2,260); sh.setColumnWidth(3,140); sh.setColumnWidth(4,70); sh.setColumnWidth(5,70); sh.setColumnWidth(6,80); sh.setColumnWidth(7,80);
  sh.setFrozenRows(3);
  viewStatusCF_(sh, 'C4:C' + (NUM_ROWS+3), 'D4:D' + (NUM_ROWS+3));
}

/* ================== V7 ==================
 * 全ビューの列を: 動画タイトル / 投稿予定 / 現在ステータス / 締切 / 制作担当 / サムネ担当 に統一
 * クライアント別ビューはBO担当を外す。増田/岩渕ビューもクライアント列なし。 */
function applyV7(){
  var ss = SpreadsheetApp.openById(SS_ID);
  rebuildClientView2_(ss.getSheetByName('クライアント別ビュー'));
  buildPersonBOView2_(ss.getSheetByName('増田ビュー'), '増田');
  buildPersonBOView2_(ss.getSheetByName('岩渕ビュー'), '岩渕');
  Logger.log('V7適用完了');
}

function viewCols6_(sh){
  sh.getRange('A3:F3').setValues([['動画タイトル','投稿予定','現在ステータス','締切','制作担当','サムネ担当']])
    .setFontWeight('bold').setBackground('#f1efe8');
  sh.getRange('B4:B' + (NUM_ROWS+3)).setNumberFormat('m/d'); // 投稿予定
  sh.getRange('D4:D' + (NUM_ROWS+3)).setNumberFormat('m/d'); // 締切
  sh.setColumnWidth(1,280); sh.setColumnWidth(2,70); sh.setColumnWidth(3,140); sh.setColumnWidth(4,70); sh.setColumnWidth(5,80); sh.setColumnWidth(6,80);
  sh.setFrozenRows(3);
  viewStatusCF_(sh, 'C4:C' + (NUM_ROWS+3), 'D4:D' + (NUM_ROWS+3));
}

function rebuildClientView2_(sh){
  sh.clear(); sh.setConditionalFormatRules([]);
  sh.getRange('A1').setValue('クライアントを選択 →').setFontWeight('bold');
  var clientRange = sh.getParent().getRange('設定!A2:A100');
  var rule = SpreadsheetApp.newDataValidation().requireValueInRange(clientRange, true).setAllowInvalid(true).build();
  sh.getRange('B1').setDataValidation(rule).setFontWeight('bold').setBackground('#e6f1fb');
  // 列: B 動画タイトル, D 投稿予定, E 状況, W 締切, G 制作担当, J サムネ担当 / 投稿予定順
  sh.getRange('A4').setFormula('=IFERROR(SORT(FILTER({マスター!B2:B,マスター!D2:D,マスター!E2:E,マスター!W2:W,マスター!G2:G,マスター!J2:J},マスター!A2:A=$B$1,マスター!A2:A<>""),2,TRUE),"該当なし")');
  viewCols6_(sh);
}

function buildPersonBOView2_(sh, name){
  sh.clear(); sh.setConditionalFormatRules([]);
  sh.getRange('A1').setValue(name + 'さんの担当案件（BO）').setFontWeight('bold').setFontSize(13).setFontColor('#185fa5');
  // 締切が近い順
  sh.getRange('A4').setFormula('=IFERROR(SORT(FILTER({マスター!B2:B,マスター!D2:D,マスター!E2:E,マスター!W2:W,マスター!G2:G,マスター!J2:J},マスター!T2:T="' + name + '",マスター!A2:A<>""),4,TRUE),"該当なし")');
  viewCols6_(sh);
}

/* ================== V8 ==================
 * 制作担当・サムネ担当が自分の担当を見る「担当者ビュー」（名前を選ぶと制作担当orサムネ担当の案件） */
function applyV8(){
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName('担当者ビュー') || ss.insertSheet('担当者ビュー');
  buildStaffView_(sh);
  ss.setActiveSheet(sh); ss.moveActiveSheet(4); // ダッシュボード/マスター/クライアント別/担当者/増田/岩渕/設定
  Logger.log('V8適用完了');
}

function buildStaffView_(sh){
  sh.clear(); sh.setConditionalFormatRules([]);
  sh.getRange('A1').setValue('担当者を選択 →').setFontWeight('bold');
  var memberRange = sh.getParent().getRange('設定!G2:G50');
  var rule = SpreadsheetApp.newDataValidation().requireValueInRange(memberRange, true).setAllowInvalid(true).build();
  sh.getRange('B1').setDataValidation(rule).setFontWeight('bold').setBackground('#e6f1fb');
  sh.getRange('A2').setValue('（あなたが制作担当 または サムネ担当の案件が出ます）').setFontColor('#888780');
  sh.getRange('A3:H3').setValues([['クライアント','動画タイトル','投稿予定','現在ステータス','制作締切','サムネ締切','公開設定・納品','担当BO']])
    .setFontWeight('bold').setBackground('#f1efe8');
  // A client, B title, D 投稿予定, E 状況, H 制作締切, K サムネ締切, R 公開設定納品, T BO担当 / 制作担当(G)=B1 or サムネ担当(J)=B1 / 投稿予定順
  sh.getRange('A4').setFormula('=IFERROR(SORT(FILTER({マスター!A2:A,マスター!B2:B,マスター!D2:D,マスター!E2:E,マスター!H2:H,マスター!K2:K,マスター!R2:R,マスター!T2:T},((マスター!G2:G=$B$1)+(マスター!J2:J=$B$1))*(マスター!A2:A<>"")*(LEN($B$1)>0)),3,TRUE),"該当なし")');
  sh.getRange('C4:C' + (NUM_ROWS+3)).setNumberFormat('m/d'); // 投稿予定
  sh.getRange('E4:E' + (NUM_ROWS+3)).setNumberFormat('m/d'); // 制作締切
  sh.getRange('F4:F' + (NUM_ROWS+3)).setNumberFormat('m/d'); // サムネ締切
  sh.getRange('G4:G' + (NUM_ROWS+3)).setNumberFormat('m/d'); // 公開設定納品
  sh.setColumnWidth(1,100); sh.setColumnWidth(2,260); sh.setColumnWidth(3,70); sh.setColumnWidth(4,140); sh.setColumnWidth(5,70); sh.setColumnWidth(6,70); sh.setColumnWidth(7,90); sh.setColumnWidth(8,80);
  sh.setFrozenRows(3);
  viewStatusCF_(sh, 'D4:D' + (NUM_ROWS+3), 'C4:C' + (NUM_ROWS+3)); // 状況=D, 投稿予定=Cを超過赤
}

/* ================== V9 ==================
 * ダッシュボードの「情報不足」を拡張: タイトル/投稿予定日/制作担当/BO担当の未入力も表示（不足項目つき）
 * ※サムネ担当は全件未入力のため除外（含めると全件出る） */
function applyV9(){
  var ss = SpreadsheetApp.openById(SS_ID);
  var dh = ss.getSheetByName('ダッシュボード');
  var M = 'マスター!';
  dh.getRange('J8').setValue('情報不足の案件（タイトル/投稿予定日/担当が未入力）').setFontWeight('bold').setFontSize(13).setFontColor('#5f5e5a');
  dh.getRange('J9:M9').clearContent();
  dh.getRange('J9:L9').setValues([['クライアント','タイトル','不足項目']]).setFontWeight('bold').setBackground('#f1efe8').setVerticalAlignment('middle');
  dh.getRange('J10').setFormula('=IFERROR(FILTER({' + M + 'A2:A,' + M + 'B2:B,' +
    'IF(' + M + 'B2:B="","タイトル ","")&IF(' + M + 'D2:D="","投稿予定日 ","")&IF(' + M + 'G2:G="","制作担当 ","")&IF(' + M + 'T2:T="","BO担当 ","")},' +
    '(' + M + 'A2:A<>"")*((' + M + 'B2:B="")+(' + M + 'D2:D="")+(' + M + 'G2:G="")+(' + M + 'T2:T=""))),"なし")');
  dh.setColumnWidth(11,240); dh.setColumnWidth(12,220);
  Logger.log('V9適用完了');
}

/* ================== V10 ==================
 * 1) 情報不足リストに「投稿予定」列を追加（クライアント/タイトル/投稿予定/不足項目）
 * 2) マスターに「増田担当」「岩渕担当」のフィルタ表示を作成（BO担当=その人） */
function applyV10(){
  var ss = SpreadsheetApp.openById(SS_ID);
  var dh = ss.getSheetByName('ダッシュボード');
  var M = 'マスター!';
  dh.getRange('J9:M9').clearContent();
  dh.getRange('J9:M9').setValues([['クライアント','タイトル','投稿予定','不足項目']])
    .setFontWeight('bold').setBackground('#f1efe8').setVerticalAlignment('middle');
  dh.getRange('J10').setFormula('=IFERROR(FILTER({' + M + 'A2:A,' + M + 'B2:B,' + M + 'D2:D,' +
    'IF(' + M + 'B2:B="","タイトル ","")&IF(' + M + 'D2:D="","投稿予定日 ","")&IF(' + M + 'G2:G="","制作担当 ","")&IF(' + M + 'T2:T="","BO担当 ","")},' +
    '(' + M + 'A2:A<>"")*((' + M + 'B2:B="")+(' + M + 'D2:D="")+(' + M + 'G2:G="")+(' + M + 'T2:T=""))),"なし")');
  dh.getRange('L10:L' + NUM_ROWS).setNumberFormat('m/d'); // 投稿予定
  dh.setColumnWidth(10,120); dh.setColumnWidth(11,240); dh.setColumnWidth(12,70); dh.setColumnWidth(13,220);
  makeFilterViews_(ss);
  Logger.log('V10適用完了');
}

/* マスターに「増田担当」「岩渕担当」フィルタ表示を作成（Sheets API直叩き） */
function makeFilterViews_(ss){
  var ms = ss.getSheetByName('マスター');
  var sheetId = ms.getSheetId();
  var endRow = NUM_ROWS + 1;
  function req(title, name){
    return {addFilterView:{filter:{
      title: title,
      range:{sheetId:sheetId, startRowIndex:0, endRowIndex:endRow, startColumnIndex:0, endColumnIndex:21},
      filterSpecs:[{columnIndex:19, filterCriteria:{condition:{type:'TEXT_EQ', values:[{userEnteredValue:name}]}}}]
    }}};
  }
  var body = {requests:[req('増田担当','増田'), req('岩渕担当','岩渕')]};
  var resp = UrlFetchApp.fetch('https://sheets.googleapis.com/v4/spreadsheets/' + ss.getId() + ':batchUpdate', {
    method:'post', contentType:'application/json',
    headers:{Authorization:'Bearer ' + ScriptApp.getOAuthToken()},
    payload: JSON.stringify(body), muteHttpExceptions:true
  });
  Logger.log('filterViews: ' + resp.getResponseCode() + ' ' + resp.getContentText().slice(0,150));
}

/* ================== V11 ==================
 * 1) 公開設定・納品締切(R)のチェックボックス書式を解除して日付(投稿2日前)に戻す
 * 2) 「修正締切」列(Y)を追加。差し戻し中はそれが現在締切になり、過ぎたら超過(赤) */
function applyV11(){
  var ss = SpreadsheetApp.openById(SS_ID);
  var ms = ss.getSheetByName('マスター');
  var last = NUM_ROWS + 1;
  // 1) R: チェックボックスの継承を解除→日付式
  ms.getRange('R2:R' + last).setDataValidation(null);
  ms.getRange('R2').setFormula('=IF($D2="","",$D2-2)');
  ms.getRange('R2').copyTo(ms.getRange('R3:R' + last));
  ms.getRange('R2:R' + last).setNumberFormat('m/d');
  // 2) 修正締切 列 Y(25)
  ms.getRange('Y1').setValue('修正締切').setFontWeight('bold').setBackground('#f1efe8').setVerticalAlignment('middle').setWrap(true);
  ms.getRange('Y2:Y' + last).setNumberFormat('m/d');
  ms.setColumnWidth(25,70);
  // 3) 現在締切 W: 修正中は修正締切(Y)を使う
  ms.getRange('W2').setFormula('=IF($A2="","",IF($F2<>"",$Y2,IFS($I2<>TRUE,$H2,$L2<>TRUE,$K2,$O2<>TRUE,$N2,$Q2<>TRUE,$P2,$S2<>TRUE,$R2,TRUE,"")))');
  ms.getRange('W2').copyTo(ms.getRange('W3:W' + last));
  // 4) 現在ステータス E: 修正中で修正締切超過なら（超過）を付ける
  ms.getRange('E2').setFormula('=IF($A2="","",IF($F2<>"","修正中（"&$F2&"）"&IF(AND($Y2<>"",$Y2<TODAY()),"（超過）",""),IFS(' +
    '$S2=TRUE,"完了",' +
    '$Q2=TRUE,"公開設定・納品待ち"&IF($R2<TODAY(),"（超過）",""),' +
    '$O2=TRUE,"CL提出待ち"&IF($P2<TODAY(),"（超過）",""),' +
    '$L2=TRUE,"チェック待ち"&IF($N2<TODAY(),"（超過）",""),' +
    '$I2=TRUE,"サムネ待ち"&IF($K2<TODAY(),"（超過）",""),' +
    'TRUE,"制作待ち"&IF($H2<TODAY(),"（超過）",""))))');
  ms.getRange('E2').copyTo(ms.getRange('E3:E' + last));
  Logger.log('V11適用完了');
}

/* ========================================================================
 * ※注意: このローカル.gsは V11 までしか反映していない（V12〜V29はメモリ参照）。
 * 本番の現行レイアウトは applyV30 適用後のもの（下記）。
 *
 * 現行レイアウト（2026-06-27 applyV30 適用後・マスター/テーブルTable1）:
 *   A クライアント / B 動画タイトル / C 完成尺 / D 投稿予定日 / E 現在ステータス(自動) /
 *   F 差し戻し / G 修正締切 / H 制作担当 / I 制作締切 / J 制作✓ /
 *   K サムネ担当 / L サムネ締切 / M サムネ✓ / N CL提出締切 / O CL提出✓ /
 *   P CLチェック✓(新) / Q 公開設定・納品締切 / R 公開設定・納品✓ / S BO担当 / T メモ /
 *   U 現在担当(自動・非表示) / V 現在締切(自動・非表示) / W 信号(自動・非表示)
 * 工程フロー: 制作 → サムネ → CL提出 → ［クライアント確認中］ → 公開設定・納品 → 完了
 * ======================================================================== */

/**
 * applyV30 — 「CLチェック✓」列(P,16)を追加し、現在ステータスに「クライアント確認中」を新設。
 *  CL提出✓(O)ON → クライアント確認中 / CLチェック✓(P)ON → 公開設定・納品待ち。
 *  CLチェック✓ はBO(増田/岩渕)がクライアントOK確認後に付ける。締切なし（工程遅延の赤は出ない）。
 * 冪等（P1が「CLチェック✓」なら何もしない）。CFの色付け(38ルール)は列挿入でGoogleが
 * 参照を自動追従。クライアント確認中ステージ専用の色は未対応（要・後追い微調整）。
 */
function applyV30(){
  var ss=SpreadsheetApp.openById('1LLkkdljyokHqjLll9rcG-ClSYnXx43D-xZtmLkAp0Ds');
  var ms=ss.getSheetByName('マスター');
  var NR=200;
  if(ms.getRange('P1').getValue()==='CLチェック✓'){ Logger.log('already applied'); return; }
  var fl=ms.getFilter();
  if(fl){ try{ fl.removeColumnFilterCriteria(5); SpreadsheetApp.flush(); }catch(e){ Logger.log('filter off skip: '+e); } }
  ms.insertColumnsAfter(15,1);
  ms.getRange('O1').copyTo(ms.getRange('P1'), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  ms.getRange('P1').setValue('CLチェック✓');
  try{ ms.getRange('P2:P'+(NR+1)).insertCheckboxes(); }catch(e){ Logger.log('checkbox skip: '+e); }
  var eF=[],tF=[],uF=[],wF=[];
  for(var r=2;r<=NR+1;r++){
    var cd='IFS($F'+r+'<>"",$G'+r+',$J'+r+'<>TRUE,$I'+r+',$M'+r+'<>TRUE,$L'+r+',$O'+r+'<>TRUE,$N'+r+',$P'+r+'<>TRUE,"",$R'+r+'<>TRUE,$Q'+r+',TRUE,"")';
    eF.push(['=IF($A'+r+'="","",IF($D'+r+'="","",IF($R'+r+'=TRUE,"",IF(OR($D'+r+'<=TODAY()+2,AND('+cd+'<>"",'+cd+'<TODAY())),"🔴 ",IF(OR($D'+r+'<=TODAY()+5,AND($J'+r+'<>TRUE,$I'+r+'<>"",$I'+r+'<=TODAY()+1)),"🟡 ",""))))&(IF($F'+r+'<>"","修正中"&IF(AND($G'+r+'<>"",$G'+r+'<TODAY()),"（超過）",""),IFS($R'+r+'=TRUE,"完了",$P'+r+'=TRUE,"公開設定・納品待ち"&IF(AND($Q'+r+'<>"",$Q'+r+'<TODAY()),"（超過）",""),$O'+r+'=TRUE,"クライアント確認中",$M'+r+'=TRUE,"CL提出待ち"&IF(AND($N'+r+'<>"",$N'+r+'<TODAY()),"（超過）",""),$J'+r+'=TRUE,"サムネ待ち"&IF(AND($L'+r+'<>"",$L'+r+'<TODAY()),"（超過）",""),TRUE,"制作待ち"&IF(AND($I'+r+'<>"",$I'+r+'<TODAY()),"（超過）","")))))']);
    tF.push(['=IF($A'+r+'="","",IF($F'+r+'<>"",$H'+r+',IFS($J'+r+'<>TRUE,$H'+r+',$M'+r+'<>TRUE,$K'+r+',$O'+r+'<>TRUE,$S'+r+',$P'+r+'<>TRUE,$S'+r+',$R'+r+'<>TRUE,$S'+r+',TRUE,"—")))']);
    uF.push(['=IF($A'+r+'="","",IF($F'+r+'<>"",$G'+r+',IFS($J'+r+'<>TRUE,$I'+r+',$M'+r+'<>TRUE,$L'+r+',$O'+r+'<>TRUE,$N'+r+',$P'+r+'<>TRUE,"",$R'+r+'<>TRUE,$Q'+r+',TRUE,"")))']);
    wF.push(['=IF($A'+r+'="","",IF($D'+r+'="","灰",IF($R'+r+'=TRUE,"青",IF(OR($D'+r+'<=TODAY()+2,AND($V'+r+'<>"",$V'+r+'<TODAY())),"赤",IF(OR($D'+r+'<=TODAY()+5,AND($J'+r+'<>TRUE,$I'+r+'<>"",$I'+r+'<=TODAY()+1)),"黄","青")))))']);
  }
  ms.getRange('E2:E'+(NR+1)).setFormulas(eF);
  ms.getRange('U2:U'+(NR+1)).setFormulas(tF);
  ms.getRange('V2:V'+(NR+1)).setFormulas(uF);
  ms.getRange('W2:W'+(NR+1)).setFormulas(wF);
  if(fl){ try{ fl.setColumnFilterCriteria(5, SpreadsheetApp.newFilterCriteria().setHiddenValues(['完了']).build()); }catch(e){ Logger.log('filter on skip: '+e); } }
  Logger.log('applyV30 done');
}

/* ========================================================================
 * 現行レイアウト（2026-06-28 applyV31 適用後・最新）:
 *   A クライアント / B 動画タイトル / C 完成尺 / D 投稿予定日 / E 現在ステータス(自動) /
 *   F 差し戻し / G 修正締切 / H 制作担当 / I 制作着手✓(新) / J 制作締切 / K 制作✓ /
 *   L サムネ担当 / M サムネ締切 / N サムネ✓ / O CL提出締切 / P CL提出✓ /
 *   Q CLチェック✓ / R 公開設定・納品締切 / S 公開設定・納品✓ / T BO担当 / U メモ /
 *   V 現在担当(自動・非表示) / W 現在締切(自動・非表示) / X 信号(自動・非表示)
 * 工程フロー: 未着手 → 制作待ち → サムネ → CL提出 → ［クライアント確認中］ → 公開設定・納品 → 完了
 * ======================================================================== */

/**
 * applyV31 — 「制作着手✓」列(I,9)を追加し、現在ステータスに「未着手」を新設。
 *  制作着手✓(I)ON → 制作待ち（編集者の着手連絡を受けてBOが付ける）。未着手=デフォルト。
 *  締切は制作締切のまま＝制作締切超過で「未着手（超過）」＋赤。担当=制作担当。冪等。
 *  ※チェックボックス列はテキスト列(制作担当)の隣に挿入するため、clearDataValidations→
 *    insertCheckboxes で明示設定（CHECKBOX型で入ることを検証済み）。
 */
function applyV31(){
  var ss=SpreadsheetApp.openById('1LLkkdljyokHqjLll9rcG-ClSYnXx43D-xZtmLkAp0Ds');
  var ms=ss.getSheetByName('マスター');
  var NR=200;
  if(ms.getRange('I1').getValue()==='制作着手✓'){ Logger.log('already applied'); return; }
  var fl=ms.getFilter();
  if(fl){ try{ fl.removeColumnFilterCriteria(5); SpreadsheetApp.flush(); }catch(e){ Logger.log('filter off skip: '+e); } }
  ms.insertColumnsAfter(8,1); // 制作担当(8)の右 → I(9)=制作着手✓
  ms.getRange('H1').copyTo(ms.getRange('I1'), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  ms.getRange('I1').setValue('制作着手✓');
  try{ ms.getRange('I2:I'+(NR+1)).clearDataValidations(); ms.getRange('I2:I'+(NR+1)).clearContent(); }catch(e){ Logger.log('clear skip: '+e); }
  try{ ms.getRange('I2:I'+(NR+1)).insertCheckboxes(); }catch(e){ Logger.log('checkbox skip: '+e); }
  var eF=[],tF=[],uF=[],wF=[];
  for(var r=2;r<=NR+1;r++){
    var cd='IFS($F'+r+'<>"",$G'+r+',$K'+r+'<>TRUE,$J'+r+',$N'+r+'<>TRUE,$M'+r+',$P'+r+'<>TRUE,$O'+r+',$Q'+r+'<>TRUE,"",$S'+r+'<>TRUE,$R'+r+',TRUE,"")';
    eF.push(['=IF($A'+r+'="","",IF($D'+r+'="","",IF($S'+r+'=TRUE,"",IF(OR($D'+r+'<=TODAY()+2,AND('+cd+'<>"",'+cd+'<TODAY())),"🔴 ",IF(OR($D'+r+'<=TODAY()+5,AND($K'+r+'<>TRUE,$J'+r+'<>"",$J'+r+'<=TODAY()+1)),"🟡 ",""))))&(IF($F'+r+'<>"","修正中"&IF(AND($G'+r+'<>"",$G'+r+'<TODAY()),"（超過）",""),IFS($S'+r+'=TRUE,"完了",$Q'+r+'=TRUE,"公開設定・納品待ち"&IF(AND($R'+r+'<>"",$R'+r+'<TODAY()),"（超過）",""),$P'+r+'=TRUE,"クライアント確認中",$N'+r+'=TRUE,"CL提出待ち"&IF(AND($O'+r+'<>"",$O'+r+'<TODAY()),"（超過）",""),$K'+r+'=TRUE,"サムネ待ち"&IF(AND($M'+r+'<>"",$M'+r+'<TODAY()),"（超過）",""),$I'+r+'=TRUE,"制作待ち"&IF(AND($J'+r+'<>"",$J'+r+'<TODAY()),"（超過）",""),TRUE,"未着手"&IF(AND($J'+r+'<>"",$J'+r+'<TODAY()),"（超過）","")))))']);
    tF.push(['=IF($A'+r+'="","",IF($F'+r+'<>"",$H'+r+',IFS($K'+r+'<>TRUE,$H'+r+',$N'+r+'<>TRUE,$L'+r+',$P'+r+'<>TRUE,$T'+r+',$Q'+r+'<>TRUE,$T'+r+',$S'+r+'<>TRUE,$T'+r+',TRUE,"—")))']);
    uF.push(['=IF($A'+r+'="","",IF($F'+r+'<>"",$G'+r+',IFS($K'+r+'<>TRUE,$J'+r+',$N'+r+'<>TRUE,$M'+r+',$P'+r+'<>TRUE,$O'+r+',$Q'+r+'<>TRUE,"",$S'+r+'<>TRUE,$R'+r+',TRUE,"")))']);
    wF.push(['=IF($A'+r+'="","",IF($D'+r+'="","灰",IF($S'+r+'=TRUE,"青",IF(OR($D'+r+'<=TODAY()+2,AND($W'+r+'<>"",$W'+r+'<TODAY())),"赤",IF(OR($D'+r+'<=TODAY()+5,AND($K'+r+'<>TRUE,$J'+r+'<>"",$J'+r+'<=TODAY()+1)),"黄","青")))))']);
  }
  ms.getRange('E2:E'+(NR+1)).setFormulas(eF);
  ms.getRange('V2:V'+(NR+1)).setFormulas(tF); // 現在担当
  ms.getRange('W2:W'+(NR+1)).setFormulas(uF); // 現在締切
  ms.getRange('X2:X'+(NR+1)).setFormulas(wF); // 信号
  if(fl){ try{ fl.setColumnFilterCriteria(5, SpreadsheetApp.newFilterCriteria().setHiddenValues(['完了']).build()); }catch(e){ Logger.log('filter on skip: '+e); } }
  Logger.log('applyV31 done');
}
