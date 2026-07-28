// **************************************************
// 内容:提出物管理
// 作成日:2021/06/21
// 作成者:noboru ando
// **************************************************
function onOpen() {
  var sht = SpreadsheetApp.getActive().getSheetByName('読み込み');
  sht.activate();
  sht.clear();
  sht.getRange('A1').activate();
  SpreadsheetApp
    .getActiveSpreadsheet()
    .addMenu('バーコード', [
      {name: 'チェック', functionName: 'barcodelabelcheck'},
      {name: '入力初期化', functionName: 'syokika'},
      {name: 'バーコード作成', functionName: 'barcodelabel'},
    ]);
}

function syokika() {
  var sht = SpreadsheetApp.getActive().getSheetByName('読み込み');
  Browser.msgBox("半角英数入力モードにしてください。全角かなモードだと正常に読み込めません");
  sht.activate();
  sht.clear();
  sht.getRange('A1').activate();
}


function barcodelabel(){
  var sht = SpreadsheetApp.getActive().getSheetByName('バーコード作成');
  var lastRow = sht.getLastRow();
  for (var i=2; i<=lastRow; i++) {
      var barcodeId = String(sht.getRange(i, 1).getValue()) + String(sht.getRange(i, 2).getValue()) + String(sht.getRange(i, 3).getValue());
//      var barc1 = '=image(\"https://www.webarcode.com/barcode/image.php?code=\"&D' + i +'&\"&type=C128B&xres=1&height=75&width=150&font=3&output=png&style=196\")'
      var barc1 = '=image(\"https://www.webarcode.com/barcode/image.php?code=\"&D' + i +'&\"&type=C128B&xres=1&height=50&width=102&font=3&output=png&style=196\")' // パラメータを変更 (height=50, width=102, style=196)
//      var barc1 = '=image(\"https://www.webarcode.com/barcode/image.php?code=\"&D' + i +'&\"&type=C128B&xres=1&height=70&width=190&font=3&output=png&style=197\")'
      sht.getRange(i, 4).setValue(barcodeId);
      sht.getRange(i, 6).setValue(barc1);
  }
}

 function barcodelabelcheck() {
    var datetime = new Date();
    var today = Utilities.formatDate(datetime,'JST', 'yyyy/MM/dd');
    var sht1 = SpreadsheetApp.getActive().getSheetByName('読み込み');
    var sht2 = SpreadsheetApp.getActive().getSheetByName('バーコード作成');
    sht2.activate();
    var lastRow1 = sht1.getLastRow();
    var lastRow2 = sht2.getLastRow();
    var lastCol2 = sht2.getLastColumn();
    sht2.getRange(1, lastCol2 + 1).setValue(today);
    const values1 = sht1.getRange(1, 1, lastRow1, 1).getValues().flat();
    const values2 = sht2.getRange(2, 4, lastRow2 -1, 1).getValues().flat();
    Logger.log(values1);
    Logger.log(values2);
    for (var i=1; i<=lastRow1; i++) {
      var r = values2.indexOf(values1[i-1]) + 2;
      sht2.getRange(r, lastCol2 + 1).setValue(1);
    }
    //ここを配列にすれば処理が高速なるのは分かってるけど、初心者がわかりやすいようにあえてセル上で処理しています。
    //for (var i=1; i<=lastRow1; i++) {
    //  for (var j=2; j<=lastRow2; j++) {
    //   if (sht1.getRange(i, 1).getValue() == sht2.getRange(j, 4).getValue()) {
    //     sht2.getRange(j, lastCol+1).setValue(1);
    //     break;
    //    }
    //  }
    //}
  }
