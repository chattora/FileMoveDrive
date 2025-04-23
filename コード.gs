
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Option')
    .addItem('FileMove', '_main')
    .addToUi();
}

//メイン
function _main()
{
  _moveFilesFromSheet();
}

//シートのファイルリストから指定ドライブに移動
function _moveFilesFromSheet() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const startRow = 2; 
  const colIndex = 1; 
  const folderId = data[0][1]; //移動先のフォルダIDを確保

  //ファイルだけをリスト化
  const fileList =  data
    .slice(startRow)                 // 下に向かっての行だけを対象に
    .map(row => row[colIndex])      // 指定列の値を取り出し
    .filter(value => value !== ""); // 空白（""）を除外

  //結果格納用
  var statusData = new Array();

  for (let i = 0; i < fileList.length; i++) {
    const url = fileList [i];
    const fileId = _extractIdFromUrl(url);

    if (!fileId || !folderId) {
      Logger.log(`スキップ: 行 ${i + 1} → 入力不備`);
      continue;
    }

    try {
      _moveSharedDriveFile(fileId, folderId);
      statusData.push("移動成功");
    } catch (e) {
      Logger.log(`エラー: 行 ${i + 1} → ${e.message}`);
      statusData.push("移動失敗");
    }
  }

  //結果をリスト化してシートに記入
  const statuslist =statusData.map(status => [status]);
  sheet.getRange(3, 3, statuslist.length, 1).setValues(statuslist); // B2セルから縦に書き出し
}

//URLからIDだけ抜き取る
function _extractIdFromUrl(url) {
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

//ファイルをドライブに移動させる
function _moveSharedDriveFile(fileId, targetFolderId) {
  const file = Drive.Files.get(fileId, { supportsAllDrives: true });
  const previousParents = file.parents.map(p => p.id).join(',');

  Drive.Files.update(
    { parents: [{ id: targetFolderId }] },
    fileId,
    null,
    {
      addParents: targetFolderId,
      removeParents: previousParents,
      supportsAllDrives: true
    }
  );

  Logger.log(` 移動成功: ${fileId} → ${targetFolderId}`);
}

