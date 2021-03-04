/* --------------------------------------------------------------------------
 * save-png.jsx
 * 開いているすべてのファイルをPNG形式で保存して閉じる
 * -------------------------------------------------------------------------- */

 // PNG形式で保存
function savePng(doc, fileName) {
    var file = new File(fileName);
    var options = new PNGSaveOptions();

    options.compression = 9;
    options.interlaced = false;

    doc.saveAs(file, options, true, Extension.LOWERCASE);
}

// 以下の「Web用に保存」するコードを使うとファイルサイズが小さくなるが、
// ファイル名のスペースが自動でハイフンへ置換される
/*
function savePng(doc, fileName) {
    var file = new File(fileName);
    var options = new ExportOptionsSaveForWeb();

    options.format = SaveDocumentType.PNG;
    options.PNG8 = false;

    doc.exportDocument(file, ExportType.SAVEFORWEB, options);
}
*/

// メイン
function main() {
    // ファイルを閉じると番号が前に詰められるので大きい方から処理
    for (var i = app.documents.length - 1; 0 <= i; i--) {
        var doc = app.documents[i];

        // 保存するPNGファイル名
        var fileName = doc.fullName.toString().replace(/\.[^\.]+$/, '.png');

        // ファイルをアクティブにする
        app.activeDocument = doc;

        // RGB形式へ変更
        doc.changeMode(ChangeMode.RGB);

        // PNG形式で保存
        savePng(doc, fileName);

        // 元のファイルを保存せずに閉じる
        doc.close(SaveOptions.DONOTSAVECHANGES);
    }

    alert('完了');
}

main();
