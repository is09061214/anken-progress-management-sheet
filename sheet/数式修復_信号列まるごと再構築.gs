/**
 * マスターの信号系4列（B=🚦 / T=現在担当 / U=現在締切 / V=信号）を全行、正しい数式に張り替えるスクリプト
 *
 * 【背景】2026-07-30時点の調査で、行の移動・挿入の影響で12行の数式が壊れていた。
 *   ・#REF!エラー化：30行目(B/T/U/V)、36行目(B/V)
 *   ・参照が別の行にズレ：31〜35、37、38、129行目、44⇔154行目（B/V）
 *   壊れた行は強制「灰＝情報不足」になったり、隣の行の日付で信号が計算されたりする。
 *
 * 【やること】
 *   1. scanBrokenFormulas … 壊れている行を調べてログに出すだけ（シートは変更しない）
 *   2. rebuildSignalFormulas … 壊れているセルだけをテンプレートで書き直す（毎朝の自動実行用）
 *      （B/T/U/Vは完全に自動計算の列なので、書き直しても手入力データには一切触れない。
 *       締切列 L/N/O/Q は手動化パッチの対象なので触らない）
 *   3. rebuildSignalFormulasFull … 全行まるごと書き直す（緊急用・手動でだけ使う）
 *
 * 【2026-08-02の改修】
 *   ・毎朝の全行書き直し＋SpreadsheetApp.flush()が原因で
 *     「Service Spreadsheets timed out」が連日発生していた。
 *     （800セルの重い数式を一斉に書き直す→再計算が90秒制限を超える）
 *   ・rebuildSignalFormulas を「壊れたセルだけ直す」方式に変更。
 *     壊れゼロの日は何も書き込まない＝タイムアウトしようがない。
 *   ・flush()と確認用の読み取りを廃止（再計算の強制待ちがタイムアウトの正体。
 *     書き込みはスクリプト終了時に自動で反映されるので不要）
 *
 * 【2026-08-09の改修】信号の判定基準を変更（赤が多すぎる問題への対応）
 *   同日午後にさらにシンプル化。最終ルール：
 *   ・赤 ＝ 投稿予定日まで7日以内（それだけ）
 *   ・黄 ＝ 現在締切を超過 or スケジュール破綻(CL提出締切>公開リミット)
 *   ・青／灰／完了の扱いは従来どおり
 *   （締切超過だけでは赤にしない＝編集が済んでいれば当日〜1日で巻き返せる。
 *     締切前日の予告は毎朝のDiscordリマインダーが担当するのでシートからは削除）
 *   ※テンプレートを変えたら rebuildSignalFormulasFull を1回手動実行して
 *     全行に新ロジックを反映すること（rebuildSignalFormulas は壊れたセルしか直さない）
 *   ※同日の反映作業で「flush()なしの4列一括書き込みだとV列の一部行が
 *     定着しない」現象を確認。Fullは列ごとにflush()する方式に変更した。
 *
 * 【使い方】
 *   1. シートの 拡張機能 → Apps Script を開く
 *   2. このファイルの中身をまるごと貼り付けて保存
 *   3. まず scanBrokenFormulas を実行してログ確認（任意）
 *   4. rebuildSignalFormulas を実行 → ログに「修復完了」が出たらOK
 */

var REBUILD_SS_ID = '1oXwpCyw_g4GT1SFh4V_DW6RL7Q8h5FhIx9nbEV_eLXo'; // iMuse案件進捗管理シート_V3

// 正常な2行目から取った数式テンプレート（{R}を行番号に置換して使う）
var SIG_FORMULAS = {
  2: /* B列 🚦 */ '=IF($A{R}="","",IFERROR(LET(cd,SWITCH(IFERROR(VLOOKUP($H{R},\'設定\'!$S:$T,2,FALSE),"なし"),"制作",$L{R},"サムネ",$N{R},"CL提出",$O{R},"公開設定",$Q{R},"修正",$J{R},""),IF(OR($D{R}="",$G{R}="",$H{R}="",$K{R}="",$R{R}="",AND(IFERROR(VLOOKUP(IF($C{R}="","通常",$C{R}),\'設定\'!$L:$Q,2,FALSE),"○")="○",$F{R}="",\'設定\'!$J$6<>"逆算で仮計算")),"⚪",IF($H{R}="完了","",IF($G{R}<=TODAY()+7,"🔴",IF(OR(AND(cd<>"",cd<TODAY()),AND($O{R}<>"",$Q{R}<>"",$O{R}>$Q{R})),"🟡","🔵"))))),"⚪"))',
  20: /* T列 現在担当 */ '=IF($A{R}="","",IF($H{R}="修正中",IF($I{R}="サムネ",$M{R},$K{R}),LET(x,IFERROR(VLOOKUP($H{R},\'設定\'!$S:$U,3,FALSE),""),SWITCH(x,"制作担当",$K{R},"サムネ担当",$M{R},"BO担当",$R{R},x))))',
  21: /* U列 現在締切 */ '=IF($A{R}="","",SWITCH(IFERROR(VLOOKUP($H{R},\'設定\'!$S:$T,2,FALSE),"なし"),"制作",$L{R},"サムネ",$N{R},"CL提出",$O{R},"公開設定",$Q{R},"修正",$J{R},""))',
  22: /* V列 信号 */ '=IF($A{R}="","",IFERROR(IF(OR($D{R}="",$G{R}="",$H{R}="",$K{R}="",$R{R}="",AND(IFERROR(VLOOKUP(IF($C{R}="","通常",$C{R}),\'設定\'!$L:$Q,2,FALSE),"○")="○",$F{R}="",\'設定\'!$J$6<>"逆算で仮計算")),"灰",IF($H{R}="完了","青",IF($G{R}<=TODAY()+7,"赤",IF(OR(AND($U{R}<>"",$U{R}<TODAY()),AND($O{R}<>"",$Q{R}<>"",$O{R}>$Q{R})),"黄","青")))),"灰"))'
};

/** 1) 調査だけ：壊れた数式の行をログに出す（シートは変更しない） */
function scanBrokenFormulas() {
  var ms = SpreadsheetApp.openById(REBUILD_SS_ID).getSheetByName('マスター');
  var last = ms.getMaxRows();
  var hit = [];
  [2, 20, 21, 22].forEach(function(colNum) {
    var colLetter = ['B', 'T', 'U', 'V'][[2, 20, 21, 22].indexOf(colNum)];
    var fs = ms.getRange(2, colNum, last - 1, 1).getFormulas();
    for (var i = 0; i < fs.length; i++) {
      var f = fs[i][0], row = i + 2;
      if (!f) continue;
      if (f.indexOf('#REF!') !== -1) { hit.push(colLetter + row + ': #REF!'); continue; }
      // 自分以外の行を参照していないか（$J$6 などの絶対参照は対象外）
      var m = f.match(/\$[A-Z]{1,2}(\d+)/g) || [];
      var wrong = m.filter(function(r) { return parseInt(r.replace(/\$[A-Z]+/, ''), 10) !== row; });
      if (wrong.length) hit.push(colLetter + row + ': 行ズレ参照 ' + wrong.join(','));
    }
  });
  Logger.log(hit.length ? '壊れセル ' + hit.length + '件:\n' + hit.join('\n') : '壊れセルなし。全行きれいです。');
}

/** 2) 修復（毎朝の自動実行用）：壊れているセルだけをテンプレートで書き直す */
function rebuildSignalFormulas() {
  var ms = SpreadsheetApp.openById(REBUILD_SS_ID).getSheetByName('マスター');
  var last = ms.getMaxRows();
  var colLetters = { 2: 'B', 20: 'T', 21: 'U', 22: 'V' };
  var fixed = [];

  Object.keys(SIG_FORMULAS).forEach(function(colKey) {
    var colNum = parseInt(colKey, 10);
    var tmpl = SIG_FORMULAS[colKey];
    var fs = ms.getRange(2, colNum, last - 1, 1).getFormulas();

    // 壊れている行を集める（#REF! / 行ズレ参照 / 数式が消えている）
    var brokenRows = [];
    for (var i = 0; i < fs.length; i++) {
      var f = fs[i][0], row = i + 2;
      if (!f) { brokenRows.push(row); continue; } // 自動計算列なので数式なしも修復対象
      if (f.indexOf('#REF!') !== -1) { brokenRows.push(row); continue; }
      var m = f.match(/\$[A-Z]{1,2}(\d+)/g) || [];
      var wrong = m.filter(function(r) { return parseInt(r.replace(/\$[A-Z]+/, ''), 10) !== row; });
      if (wrong.length) brokenRows.push(row);
    }
    if (!brokenRows.length) return;

    if (brokenRows.length > 50) {
      // 大量に壊れているときだけ列まるごと書き直し（1回のAPI呼び出しで済ませる）
      var out = [];
      for (var r = 2; r <= last; r++) out.push([tmpl.replace(/\{R\}/g, r)]);
      ms.getRange(2, colNum, last - 1, 1).setFormulas(out);
      fixed.push(colLetters[colNum] + '列まるごと（' + brokenRows.length + '行壊れ）');
    } else {
      brokenRows.forEach(function(row) {
        ms.getRange(row, colNum).setFormula(tmpl.replace(/\{R\}/g, row));
        fixed.push(colLetters[colNum] + row);
      });
    }
  });

  // flush()や確認読み取りはしない：重い再計算を待つとタイムアウトするため。
  // 書き込みはスクリプト終了時に自動反映される。
  Logger.log(fixed.length
    ? '修復完了：' + fixed.length + 'セルを書き直しました → ' + fixed.join(', ')
    : '壊れセルなし。何も書き込みませんでした。');
}

/** 3) 緊急用（手動でだけ使う）：B/T/U/V列の数式を全行、テンプレートで書き直す */
function rebuildSignalFormulasFull() {
  var ms = SpreadsheetApp.openById(REBUILD_SS_ID).getSheetByName('マスター');
  var last = ms.getMaxRows();
  var n = last - 1;
  Object.keys(SIG_FORMULAS).forEach(function(colNum) {
    var tmpl = SIG_FORMULAS[colNum];
    var out = [];
    for (var r = 2; r <= last; r++) out.push([tmpl.replace(/\{R\}/g, r)]);
    ms.getRange(2, parseInt(colNum, 10), n, 1).setFormulas(out);
    // flush()なしの4列一括書き込みだと一部の行が定着しないことがある（2026-08-09に確認）
    SpreadsheetApp.flush();
  });
  Logger.log('全行再構築の書き込み完了：B/T/U/V列 2〜' + last + '行。\n検証 V98: ' + ms.getRange('V98').getFormula().slice(0, 160));
}
