
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Option')
    .addItem('挨拶する', 'sayHello')
    .addToUi();
}


function moveFilesFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  // ヘッダーをスキップ
  for (let i = 1; i < data.length; i++) {
    const url = data[i][0];
    const targetFolderId = data[i][1];

    const fileId = extractIdFromUrl(url);
    if (!fileId || !targetFolderId) {
      Logger.log(`スキップ: 行 ${i + 1} → 入力不備`);
      continue;
    }

    try {
      moveSharedDriveFile(fileId, targetFolderId);
    } catch (e) {
      Logger.log(`エラー: 行 ${i + 1} → ${e.message}`);
    }
  }
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
