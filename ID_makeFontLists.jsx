/**
 * @fileoverview FontListGenerator用のフォントリストファイルを生成する
 * @version  v0.9.0b
 * @author  Yusuke SAEGUSA
 * @description 
 * - 選択したテキストフレームに使われているフォントをリストにしてtxtファイルに書き出します
 * - 書き出し時、和文フォントか欧文フォントかを選択します
 * - 書き出し場所を指定してください
 * - 書き出されるテキストはUTF-8、改行コードLFになります
 */

/**
 * @type {Object} 環境変数
 * @prop {String} errorMessage エラーメッセージ
 * @prop {Document} currentDoc 処理対象ドキュメント
 * @prop {String} scriptVer スクリプトのバージョン表記
 * @prop {Object} fontInfo 取得したフォント情報の配列（フォントファミリー名）
 * @prop {String} fileName 書き出すファイルのファイル名（拡張子別）
 * @prop {String} lang 言語コード（和文の場合"J", 欧文の場合"E"）
*/
var _env = {
  errorMessage: "",
  currentDoc: {},
  scriptVer: "v0.9.0b",
  fontInfo: [],
  fileName: "",
  lang: "",
};

app.doScript(main, ScriptLanguage.JAVASCRIPT, null, UndoModes.ENTIRE_SCRIPT);

function main() {
  if (!validationBeforeInvoke()) { quitScriptWithMassage(); }
  if (!showExportSettingDialog()) { quitScriptWithMassage(); }
  if (!getFontsFromSelection()) { quitScriptWithMassage(); }
  if (!exportFontList()) { quitScriptWithMassage(); }
  alert("終了しました。");
}

/**
 * メッセージを表示してスクリプトを終了させる
 *
 */
function quitScriptWithMassage() {
  if (_env.errorMessage) {
    alert(_env.errorMessage);
    _env.errorMessage = "";
  }
  exit();
}

/**
 * スクリプト実行前のバリデーション
 *
 * @returns {Boolean} 問題なくバリデーションが終了したか
 */
function validationBeforeInvoke() {
  var docs = app.documents;
  if (docs.length === 0) {
    _env.errorMessage = "ドキュメントを開いてテキストやテキストフレーム等を選択してから実行してください。";
    return false;
  }
  if (docs[0].selection.length === 0) {
    _env.errorMessage = "テキストやテキストフレーム等を選択してから実行してください。";
    return false;
  }
  _env.currentDoc = app.activeDocument;
  return true;
}

/**
 * 選択したオブジェクトからフォント情報を取得する
 *
 */
function getFontsFromSelection() {
  var sel = _env.currentDoc.selection;
  var chars, fonts;
  var tempList = {};
  for (var i = 0; i < sel.length; i++) {
    if (!sel[i].hasOwnProperty("texts")) { continue; }
    chars = sel[i].texts[0].characters.everyItem();
    fonts = chars.appliedFont;
    for (var j = 0; j < fonts.length; j++) {
      if (!tempList[fonts[j].fontFamily]) {
        tempList[fonts[j].fontFamily] = true; // 重複を削除するためにkeyにフォントファミリーを設定する
      }
    }
  }
  for (var k in tempList) {
    _env.fontInfo.push(k);
  }
  if (_env.fontInfo.length === 0) {
    _env.errorMessage = "選択したオブジェクトからフォント情報が検出できませんでした。";
    return false;
  }
  return true;
}

/**
 * フォントリスト書き出し設定ダイアログの表示
 *
 * @return {Boolean} エラーなく終了したか
 */
function showExportSettingDialog() {
  var excludeChar = /[.$¥\/:*?\"<>|\\]/g; // ファイル名禁止文字（割と適当）
  var dlg = new Window("dialog", "フォントリスト書き出し設定");
  var gr_fileName = dlg.add("group {orientation: 'column', spacing: '6', alignChildren: ['center', 'fill']}");
  gr_fileName.add("staticText", undefined, "書き出しファイル名（拡張子不要）");
  var fileName = gr_fileName.add("editText {preferredSize: [200, -1]}");
  var gr_lang = dlg.add("group {orientation: 'column', spacing: '6', alignChildren: ['center', 'fill']}");
  gr_lang.add("staticText", undefined, "言語設定");
  var gr_radioButtons = gr_lang.add("group");
  var lang_j = gr_radioButtons.add("radioButton {value: true, text: '和文'}");
  var lang_e = gr_radioButtons.add("radioButton {value: false, text: '欧文'}");
  var gr_buttons = dlg.add("group");
  var btn_cancel = gr_buttons.add("button", undefined, "中断", { name: "cancel" });
  var btn_exec = gr_buttons.add("button", undefined, "実行", { name: "ok" });
  btn_cancel.onClick = function() {
    dlg.close(0);
  }
  btn_exec.onClick = function() {
    dlg.close(1);
  }
  var dlg_result = dlg.show();
  if (dlg_result === 0) { exit(); }
  if (fileName.text === "" || excludeChar.test(fileName.text)) {
    _env.errorMessage = "ファイル名には非推奨の文字列が含まれています。新ためてスクリプトを実行してください。\r" + fileName.text;
    return false;
  }
  _env.fileName = fileName.text;
  if (lang_j.value) { _env.lang = "J"; }
  else { _env.lang = "E"; }
  return true;
}

/**
 * ファイル書き出し
 *
 * @return {Boolean} ファイル書き出しが成功したかどうか 
 */
function exportFontList() {
  for (var i = 0; i < _env.fontInfo.length; i++) {
    _env.fontInfo[i] = _env.lang + "\t" + _env.fontInfo[i];
  }
  var contents = _env.fontInfo.join("\n");
  var tgtPath = Folder.selectDialog();
  var newFile, openFlag, isError;
  try {
    newFile = new File(tgtPath + "/" + _env.fileName + ".txt");
    if (newFile.exists) {
      if (!confirm("既にファイルが存在します。上書きしますか？", false)) {
        exit();
      }
    }
    newFile.encoding = "utf-8";
    newFile.lineFeed = "Unix";
    openFlag = newFile.open("w");
    newFile.write(contents);
  } catch (e) {
    alert(e);
    isError = true;
  } finally {
    newFile.close();
  }
  if (!openFlag || isError) {
    _env.errorMessage = "フォントリストファイルをうまく書き出せませんでした。選択したフォルダの権限などを確認してください。もしくは、別のフォルダを指定してください。";
    return false;
  }
  return true;
}
