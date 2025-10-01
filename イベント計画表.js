function onEdit(e) {
  var sheet = e.source.getActiveSheet(); // 編集があったシートを取得
  var sheetName = sheet.getName(); // シート名を取得
  //if (sheetName !== "イベント計画表") return; // 「イベント計画表」以外なら処理しない

  var range = e.range; // 編集されたセルの情報を取得
  var row = range.getRow(); // 入力された行
  var column = range.getColumn(); // 入力された列
  var value = range.getValue(); // 入力された値

  // 日付なら "MM/DD" 形式に統一
  var formattedValue = (value instanceof Date) 
    ? Utilities.formatDate(value, Session.getScriptTimeZone(), "MM/dd") 
    : value;

  if (column == 5 && value != "") {  // **E列（5列目）に入力があった場合 → 赤色に塗る←廃止する。（2/18）**
    var startColumn = 7; // **G列（7列目）からチェック**
    var dataRange = sheet.getRange(3, startColumn, 1, sheet.getLastColumn() - startColumn + 1);
    var data = dataRange.getValues()[0]; // 3行目のG列以降のデータを取得
    var matchingColumns = [];

    for (var i = 0; i < data.length; i++) {
      var cellValue = data[i];
      var formattedCellValue = (cellValue instanceof Date) 
        ? Utilities.formatDate(cellValue, Session.getScriptTimeZone(), "MM/dd") 
        : cellValue;

      if (formattedCellValue == formattedValue) {
        matchingColumns.push(i + startColumn);
      }
    }

    if (matchingColumns.length > 0) {
      matchingColumns.forEach(function(col) {
        var targetCell = sheet.getRange(row, col);
        //targetCell.setBackground("red"); // **セルを赤色に変更**
        targetCell.setValue("期"); // **「期」を入力**
        targetCell.setHorizontalAlignment("center"); // 文字を横方向に中央揃え
        targetCell.setVerticalAlignment("middle"); // 文字を縦方向に中央揃え
      });
    } else {
      Browser.msgBox("3行目のG列以降に '" + formattedValue + "' は見つかりませんでした。");
    }
  } 

  else if (column == 6 && value != "") {  // **F列（6列目）に入力があった場合 → 「済」を入力**
    var startColumn = 7; // **G列（7列目）からチェック（変更なし）**
    var dataRange = sheet.getRange(3, startColumn, 1, sheet.getLastColumn() - startColumn + 1);
    var data = dataRange.getValues()[0]; // 3行目のG列以降のデータを取得
    var matchingColumns = [];

    for (var i = 0; i < data.length; i++) {
      var cellValue = data[i];
      var formattedCellValue = (cellValue instanceof Date) 
        ? Utilities.formatDate(cellValue, Session.getScriptTimeZone(), "MM/dd") 
        : cellValue;

      if (formattedCellValue == formattedValue) {
        matchingColumns.push(i + startColumn);
      }
    }

    if (matchingColumns.length > 0) {
      matchingColumns.forEach(function(col) {
        var targetCell = sheet.getRange(row, col);
        targetCell.setValue("済"); // **「済」を入力**
        targetCell.setHorizontalAlignment("center"); // 文字を横方向に中央揃え
        targetCell.setVerticalAlignment("middle"); // 文字を縦方向に中央揃え
      });
    } else {
      Browser.msgBox("3行目のG列以降に '" + formattedValue + "' は見つかりませんでした。");
    }
  }
}