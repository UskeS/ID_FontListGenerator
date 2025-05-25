/**
 * @fileoverview ID_FontListGenerator
 * @version  v0.9.0b
 * @author  Yusuke SAEGUSA
 * @description 
 * 動作保証バージョン InDesign 2024(v19.5.3)/2025(v20.2)
 */

// 参照用定数
var doc = app.documents[0];
var pStyleName = "例文";

// J、E以外のオプションを加えたらimportSettings()メソッド内も要修正
var constant = {
  weightChar: {
    J: "永",
    E: "Ag",
  }, 
  constantText: {
    J: "あいう アイウ 永辻葛 AWG xhigp 0123", 
    E: "AWG xhigp 0123 % $ ? ,", 
  },
  exSentence: {
    J: "和文サンプル",
    E: "欧文サンプル",
  },
  xrefFormName: "_段落テキスト全体",
  tableStyle: "フォントリスト",
};

// スクリプト実行前バリデーション
if (app.documents.length === 0) {
  alert("テンプレートドキュメントを開いてからスクリプトを実行してください。");
  exit();
}

if (!doc.textFrames.item(constant.exSentence.J).isValid 
  || !doc.textFrames.item(constant.exSentence.E).isValid 
  || !doc.tableStyles.item(constant.tableStyle).isValid
  || !doc.crossReferenceFormats.item(constant.xrefFormName).isValid
  ) {
  alert("テンプレートドキュメントの情報が正しく参照できません。\rテンプレートドキュメントを開いてから実行してください。もし改変してしまった場合は、改めてダウンロードしてから再度実行してください。");
  exit();
}

// メイン処理
app.doScript(main, ScriptLanguage.JAVASCRIPT, null, UndoModes.ENTIRE_SCRIPT);

/**
 * メイン関数
 *
 */
function main() {
  // 設定ファイルの読み込み
  var fontInfo = importSettings();
  var tgtTxf;
  // 読み込んだファイルそれぞれ(i)に対する処理
  for (var i = 0; i < fontInfo.length; i++) {
    var Fi = fontInfo[i];
    // ページ追加処理
    if (doc.pages.length > 1 || i !== 0) {
      tgtTxf = genCategoryPage(Fi.categoryName);
    } else {
      doc.sections[0].marker = Fi.categoryName;
      tgtTxf = doc.pages[0].textFrames[0];
    }
    // 読み込んだフォントリスト(j)に対する処理
    for (var j = 0, fl = Fi.fontList; j < fl.length ;j++) {
      var temp = Fi.fontList[j].split("\t");
      var contObj = {
        lang: temp[0],
        family: temp[1],
        weights: getWeightString(temp[1]),
        fixed: constant.constantText[temp[0]],
        exSentence: "" // 不使用（プロパティ参照用）
      };
      if (contObj.weights[0] === "–") {
        var tempFont = app.fonts.item(contObj.family);
      } else {
        var tempFont = app.fonts.item(contObj.family + "\t" + contObj.weights[0]);
      }
      if (tempFont.status === FontStatus.NOT_AVAILABLE || tempFont.status === FontStatus.FAUXED) {
        alert("「" + contObj.family + " " + contObj.weights[0] + "」が利用できません。");
        exit();
      }
      insertTable(tgtTxf.insertionPoints[-1], contObj, tempFont);
      if (tgtTxf.overflows) {
        var p = doc.pages.add(LocationOptions.AFTER, doc.pages[-1]);
        var t = genTextFrame(p);
        tgtTxf.nextTextFrame = t;
        tgtTxf = t;
      }
    }
  }
}

/**
 * 設定ファイルの読み込み
 *
 * @return {Object[]} 
 */
function importSettings() {
  var result = [], isError = false;
  var targetFiles = File.openDialog("フォントリスト設定ファイルを選択してください。", setFileType(), true);
  var re = /^[JE]\t/; // JとE以外のオプションを加えたらここも修正
  if (!targetFiles) { exit(); }
    for (var i = 0; i < targetFiles.length; i++) {
      try {
        var openFlag = targetFiles[i].open("r");
        if (!openFlag) {
          throw new Error(targetFiles[i].fsName + "が開けませんでした。");
        }
        var curInfo = {
          categoryName: targetFiles[i].displayName.replace(/\.txt/i, ""),
          fontList: targetFiles[i].read(99999).split("\n")
        };
        if (curInfo.fontList.length === 0) { throw new Error("フォントが記述されたファイルを指定してください。"); }
        else if (!hasEveryContent(curInfo.fontList, re)) { throw new Error("フォントリストの記述が正しくありません。\r和文サンプルであれば行頭にJ、欧文サンプルであれば行頭にEを入れ、タブで区切ったあとフォント名を記述してください。"); }
        result.push(curInfo);
      } catch(e) {
        alert(e);
        isError = true;
      } finally {
        targetFiles[i].close();
      }
      if (isError) { exit(); }
    }
    return result;
}

/**
 * File.openDialog用のコールバック関数
 *
 * @return {any} mac: Boolean, win: String
 */
function setFileType() {
  if (/^mac/i.test($.os)) {
    return function(f) {
      return (f instanceof Folder || /\.txt$/i.test(f.fsName));
    }
  } else if (/^win/i.test($.os)) {
    return "*.txt";
  }
}

/**
 * 配列の要素すべてを正規表現でチェックし、ひとつでもマッチしなければfalseを返す
 *
 * @param {String[]} ary フォントリスト
 * @param {RegExp} re リストをチェックする正規表現
 * @return {Boolean} ひとつでもfalseがあればfalse
 */
function hasEveryContent(ary, re) {
  for (var i = 0; i < ary.length; i++) {
    if (!re.test(ary[i])) { return false; }
  }
  return true;
}

/**
 * 表組を挿入する
 *
 * @param {InsertionPoint} ins 表組を挿入する挿入点
 * @param {Object} contObj 表組に挿入するコンテンツのオブジェクト
 * @param {Font} fontObj 表組のテキストに適用するFontオブジェクト
 */
function insertTable(ins, contObj, fontObj) {
  var cst = doc.cellStyles;
  var curTable = ins.tables.add(LocationOptions.AFTER, ins, {
    columnCount: 2,
    bodyRowCount: 2,
    headerRowCount: 1,
    topBorderStrokeWeight: 0,
    leftBorderStrokeWeight: 0,
    rightBorderStrokeWeight: 0,
    bottomBorderStrokeWeight: 0,
    appliedTableStyle: doc.tableStyles.item(constant.tableStyle),
  });
  curTable.cells[0].width = "94mm";
  curTable.cells[1].width = "96mm";
  curTable.cells[0].appliedCellStyle = cst.item("01_ヘッダ");
  curTable.cells[0].contents = contObj.family;
  curTable.cells[2].appliedCellStyle = cst.item("02_ウェイト一覧");
  curTable.cells[2].contents = genWeightChars(contObj.weights, contObj.lang);
  curTable.cells[3].appliedCellStyle = cst.item("04_例文");
  curTable.cells[3].contents = contObj.exSentence;
  curTable.cells[4].appliedCellStyle = cst.item("03_例文 固定テキスト");
  curTable.cells[4].contents = contObj.fixed;
  for (var i = 0, c = curTable.cells; i < c.length; i++) {
    c[i].clearCellStyleOverrides();
  }
  curTable.cells[5].merge(curTable.cells[3]);
  curTable.cells[1].merge(curTable.cells[0]);
  insertCrossReference(doc, constant.exSentence[contObj.lang], constant.xrefFormName, curTable.cells[2].insertionPoints[0]);
  for (var i = 0, c = curTable.cells; i < c.length; i++) {
    for (var j = 0, p = c[i].paragraphs; j < p.length; j++) {
      p[j].appliedFont = fontObj;
    }
  }
  setWeightAndRuby(contObj.weights, curTable.cells[1], contObj.lang);
  ins.parentStory.insertionPoints[-1].contents = "\r";
}

/**
 * ウェイト一覧のための文字を生成する
 *
 * @param {String[]} w ウェイトの配列
 * @param {String} lang J: 日本語用, E: 欧文用
 * @return {String} ウェイト表記用文字列 
 */
function genWeightChars(w, lang) {
  var result = [];
  for (var i = 0; i < w.length; i++) {
    result.push(constant.weightChar[lang]);
  }
  return result.join("\u2001");
}

/**
 * フォントファミリー名からウェイト一覧を取得する
 *
 * @param {String} fontFam フォントファミリー名
 * @return {String[]} 対応するウェイトの配列
 */
function getWeightString(fontFam) {
  var result = [];
  var ff = app.fonts.everyItem().name;
  var ft = app.fonts.everyItem().fontType;
  for (var i = 0; i < ff.length; i++) {
    var curFont = ff[i].split("\t");
    if (fontFam == curFont[0]) {
      if (ft[i] === FontTypes.ATC) {
        result.push("–");
      } else {
        result.push(curFont[1]);
      }
    }
  }
  if (result.length === 0) {
    alert("「"+ fontFam + "」が見つかりませんでした。フォント名を確認してください。");
    exit();
  }
  return result;
}

/**
 * 生成したウェイト一覧用文字列に、ウェイトを適用し、ウェイト名をルビとして挿入する
 *
 * @param {String[]} weightList ウェイト名の配列
 * @param {Text} tgt ウェイトとルビの適用対象テキスト
 * @param {String} lang 対応言語
 */
function setWeightAndRuby(weightList, tgt, lang) {
  app.findTextPreferences = app.changeTextPreferences = null;
  app.findTextPreferences.findWhat = constant.weightChar[lang];
  var found = tgt.findText();
  for (var i = 0; i < found.length; i++) {
    found[i].appliedFont = found[i].appliedFont.fontFamily + "\t" + weightList[i];
    found[i].rubyFlag = true;
    found[i].rubyString = weightList[i];
  }
}

/**
 * 相互参照を挿入する
 *
 * @param {Document} sourceDoc 対象Documentオブジェクト
 * @param {String} destTxfName 相互参照先対象テキストフレーム名
 * @param {String} formatName 相互参照フォーマット名
 * @param {InsertionPoint} insertionPoint 相互参照を挿入する挿入点
 * @return {Hyperlink} 挿入したHyperlinkオブジェクト
 */
function insertCrossReference(sourceDoc, destTxfName, formatName, insertionPoint) {
  var destPara = sourceDoc.textFrames.item(destTxfName).paragraphs[0];
  var refPara = sourceDoc.paragraphDestinations.add(destPara);
  var xRefForm = sourceDoc.crossReferenceFormats.item(formatName);
  var source = sourceDoc.crossReferenceSources.add(insertionPoint, xRefForm);
  var myLink = sourceDoc.hyperlinks.add(source, refPara);
  return myLink;
}

/**
 * カテゴリー用ページ生成
 *
 * @param {String} catName カテゴリー名
 * @return {TextFrame} 表組流し込み用のテキストフレーム
 */
function genCategoryPage(catName) {
  var p = doc.pages.add();
  doc.sections.add(p, { marker: catName });
  return genTextFrame(p);
}

/**
 * マージンいっぱいのテキストフレーム生成
 *
 * @param {Page} tgtPage 生成対象ページ
 * @return {TextFrame} 生成したテキストフレーム
 */
function genTextFrame(tgtPage) {
  var b = tgtPage.bounds;
  var m = tgtPage.marginPreferences;
  var t = tgtPage.textFrames.add({
    geometricBounds: [
      b[0] + m.top,
      b[1] + m.left,
      b[2] - m.right,
      b[3] - m.bottom,
    ]
  });
  t.insertionPoints[0].applyParagraphStyle(doc.paragraphStyles.itemByName(pStyleName));
  return t;
}
