
//https://drive.google.com/drive/folders/1jGkhWFuz9WsmWlmJeSBlHJ2kDon08sgr
//https://docs.google.com/spreadsheets/d/1ekYwgxdgtC2p4mjJbS2SM9-AKgH7c9PN2pIp2a3vrGM/edit?gid=0#gid=0

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Option')
    .addItem('FileMove', '_main')
    .addToUi();
}

function _main()
{

  _moveFilesFromSheet();

}

function _moveFilesFromSheet() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  const startRow = 2; 
  const colIndex = 1; 

  const folderId = data[0][1]; //移動先のフォルダIDを確保
  const fileList =  data
    .slice(startRow)                 // 下に向かっての行だけを対象に
    .map(row => row[colIndex])      // 指定列の値を取り出し
    .filter(value => value !== ""); // 空白（""）を除外

    var statusData = new Array();

  for (let i = 0; i < fileList.length; i++) {
    const url = fileList [i];

    const fileId = extractIdFromUrl(url);
    if (!fileId || !folderId) {
      Logger.log(`スキップ: 行 ${i + 1} → 入力不備`);
      continue;
    }

    try {
      moveSharedDriveFile(fileId, folderId);
      statusData.push("移動成功");
    } catch (e) {
      Logger.log(`エラー: 行 ${i + 1} → ${e.message}`);
      statusData.push("移動失敗");
    }
  }

  const statuslist =statusData.map(status => [status]);

  
  sheet.getRange(3, 3, statuslist.length, 1).setValues(statuslist); // B2セルから縦に書き出し


}

function extractIdFromUrl(url) {
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

function moveSharedDriveFile(fileId, targetFolderId) {
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

  Logger.log(`✅ 移動成功: ${fileId} → ${targetFolderId}`);
}


















/*
function moveSharedDriveFile(fileId, targetFolderId) {
  const file = Drive.Files.get(fileId, { supportsAllDrives: true });
  const previousParents = file.parents.map(p => p.id).join(',');

  Drive.Files.update(
    { parents: [{ id: targetFolderId }] },
    fileId,
    null,  // ← mediaData は不要
    {
      addParents: targetFolderId,
      removeParents: previousParents,
      supportsAllDrives: true
    }
  );

  Logger.log(`ファイル ${fileId} を ${targetFolderId} に移動しました`);
}



function extractIdFromUrl(url) {
  // ファイル or フォルダのIDをURLから抜き出す
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}


function moveFileByUrls(fileUrl, folderUrl) {

  //const fileId = extractIdFromUrl(fileUrl);
  //const folderId = extractIdFromUrl(folderUrl);

//https://drive.google.com/file/d/13aJ3HyqeTSjiNkxv48cJCiXgc3V8YjEr/view?usp=drivesdk
  const fileId = "13aJ3HyqeTSjiNkxv48cJCiXgc3V8YjEr";
  const folderId = "1jGkhWFuz9WsmWlmJeSBlHJ2kDon08sgr";
  
  if (!fileId || !folderId) {
    Logger.log('URLからIDが取得できませんでした');
    return;
  }

  moveSharedDriveFile(fileId, folderId);
}
*/
