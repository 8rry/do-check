/**
 * ユーティリティ関数
 * パス処理、ファイル操作、共通処理を管理
 */

/**
 * ファイルパスを読み込み
 * @returns {string} ファイルパス
 */
function loadFolderPath() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.INFO_EXTRACTION);
    const folderPath = sheet.getRange(CONFIG.CELLS.FOLDER_PATH).getValue();
    
    console.log(`✅ ファイルパス読み込み完了: ${folderPath}`);
    return folderPath;
    
  } catch (error) {
    console.log(`❌ ファイルパス読み込みエラー: ${error.message}`);
    return null;
  }
}

/**
 * WindowsパスをGoogle Driveパスに変換（動作していたバージョン）
 * @param {string} windowsPath - Windowsパス
 * @returns {Object} フォルダキーとサブパスの情報
 */
function convertWindowsPathToDrivePath(windowsPath) {
  try {
    console.log(`🔍 パス解析開始: ${windowsPath}`);
    
    const pathParts = windowsPath.toString().split('\\');
    let driveIndex = -1;
    
    // 共有ドライブの位置を特定
    for (let i = 0; i < pathParts.length; i++) {
      if (pathParts[i] === '共有ドライブ') {
        driveIndex = i;
        break;
      }
    }
    
    if (driveIndex === -1 || driveIndex + 2 >= pathParts.length) {
      throw new Error('共有ドライブの位置が不正です');
    }
    
    // 自治体フォルダキー（共有ドライブの次のフォルダ）
    const folderKey = pathParts[driveIndex + 1];
    
    // サブパス（自治体フォルダ以降のパス）
    const subPath = pathParts.slice(driveIndex + 2).join('/');
    
    console.log(`📊 パス解析結果:`);
    console.log(`  - 自治体フォルダキー: ${folderKey}`);
    console.log(`  - サブパス: ${subPath}`);
    console.log(`  - パス構造: ${pathParts.join(' > ')}`);
    
    return {
      folderKey: folderKey,
      subPath: subPath,
      fullPath: windowsPath
    };
    
  } catch (error) {
    console.log(`❌ パス変換エラー: ${error.message}`);
    return null;
  }
}

/**
 * 自治体フォルダを検索
 * @param {string|Object} folderKey - フォルダキー（文字列またはオブジェクト）
 * @returns {string} フォルダID
 */
function findMunicipalityFolder(folderKey) {
  try {
    // オブジェクトが渡された場合、folderKeyプロパティを取得
    let actualFolderKey = folderKey;
    if (typeof folderKey === 'object' && folderKey !== null) {
      if (folderKey.folderKey) {
        actualFolderKey = folderKey.folderKey;
        console.log(`🔍 オブジェクトからフォルダキーを抽出: ${actualFolderKey}`);
      } else {
        throw new Error('フォルダキーオブジェクトにfolderKeyプロパティがありません');
      }
    }
    
    console.log(`🔍 自治体フォルダ検索開始: ${actualFolderKey}`);
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MUNICIPALITY_FOLDERS);
    
    if (!sheet) {
      throw new Error('自治体フォルダタブが見つかりません');
    }
    
    const lastRow = sheet.getLastRow();
    console.log(`📊 検索対象行数: ${lastRow - 1}行 (2行目から${lastRow}行目)`);
    
    for (let row = 2; row <= lastRow; row++) { // 2行目から開始
      const folderName = sheet.getRange(row, 2).getValue(); // B列
      const folderId = sheet.getRange(row, 3).getValue(); // C列
      
      console.log(`  - 行${row}: "${folderName}" (ID: ${folderId})`);
      
      if (folderName === actualFolderKey) {
        console.log(`✅ 自治体フォルダ検索完了: ${actualFolderKey} (ID: ${folderId})`);
        return folderId;
      }
    }
    
    throw new Error(`フォルダキー "${actualFolderKey}" が見つかりません`);
    
  } catch (error) {
    console.log(`❌ 自治体フォルダ検索エラー: ${error.message}`);
    return null;
  }
}

/**
 * ファイル名を抽出
 * @param {string} folderPath - ファイルパス
 * @returns {string} ファイル名
 */
function extractFileNameFromPath(folderPath) {
  const pathParts = folderPath.split('\\');
  const fileName = pathParts[pathParts.length - 1];
  console.log(`📁 対象ファイル: ${fileName}`);
  return fileName;
}

/**
 * フォルダ内から指定されたファイルを検索
 * @param {string} folderId - フォルダID
 * @param {string} fileName - ファイル名
 * @returns {string|null} ファイルID
 */
function findFileInFolder(folderId, fileName) {
  try {
    console.log(`🔍 フォルダ内でファイル検索開始: ${fileName}`);
    
    const folder = DriveApp.getFolderById(folderId);
    console.log(`📁 対象フォルダ: ${folder.getName()} (ID: ${folderId})`);
    
    const files = folder.getFiles();
    let fileCount = 0;
    let foundFiles = [];
    
    console.log(`📋 フォルダ内のファイル一覧:`);
    
    while (files.hasNext()) {
      const file = files.next();
      fileCount++;
      
      const fileInfo = {
        name: file.getName(),
        id: file.getId(),
        type: file.getMimeType(),
        size: file.getSize()
      };
      
      foundFiles.push(fileInfo);
      console.log(`  ${fileCount}. ${fileInfo.name} (ID: ${fileInfo.id}, タイプ: ${fileInfo.type}, サイズ: ${fileInfo.size} bytes)`);
      
          // ファイル名の完全一致をチェック
    if (file.getName() === fileName) {
      console.log(`✅ ファイル発見: ${fileName} (ID: ${file.getId()})`);
      return file.getId();
    }
    
    // 正規化したファイル名でもチェック（特殊文字を除去）
    const normalizedFileName = fileName.replace(/[^a-zA-Z0-9]/g, '');
    const normalizedFile = file.getName().replace(/[^a-zA-Z0-9]/g, '');
    
    if (normalizedFile === normalizedFileName) {
      console.log(`✅ 正規化ファイル名で発見: ${fileName} → ${file.getName()} (ID: ${file.getId()})`);
      return file.getId();
    }
    }
    
    console.log(`📊 フォルダ内ファイル数: ${fileCount}件`);
    
    // ファイルが見つからない場合、類似ファイル名を探す
    console.log(`🔍 類似ファイル名の検索:`);
    const similarFiles = foundFiles.filter(file => 
      file.name.toLowerCase().includes(fileName.toLowerCase().replace(/[^a-zA-Z0-9]/g, '')) ||
      fileName.toLowerCase().replace(/[^a-zA-Z0-9]/g, '').includes(file.name.toLowerCase())
    );
    
    if (similarFiles.length > 0) {
      console.log(`💡 類似ファイルが見つかりました:`);
      similarFiles.forEach(file => {
        console.log(`  - ${file.name} (ID: ${file.id})`);
      });
    }
    
    console.log(`❌ ファイルが見つかりません: ${fileName}`);
    
    // サブフォルダも検索してみる
    console.log(`🔍 サブフォルダ内も検索中...`);
    const subFolders = folder.getFolders();
    let subFolderCount = 0;
    
    while (subFolders.hasNext()) {
      const subFolder = subFolders.next();
      subFolderCount++;
      console.log(`📁 サブフォルダ${subFolderCount}: ${subFolder.getName()} (ID: ${subFolder.getId()})`);
      
      try {
        const subFolderFiles = subFolder.getFiles();
        while (subFolderFiles.hasNext()) {
          const file = subFolderFiles.next();
          console.log(`  📄 ${file.getName()} (ID: ${file.getId()}, タイプ: ${file.getMimeType()})`);
          
          if (file.getName() === fileName) {
            console.log(`✅ サブフォルダ内でファイル発見: ${fileName} (ID: ${file.getId()})`);
            return file.getId();
          }
        }
      } catch (subFolderError) {
        console.log(`⚠️ サブフォルダ ${subFolder.getName()} のアクセスエラー: ${subFolderError.message}`);
      }
    }
    
    console.log(`📊 サブフォルダ数: ${subFolderCount}件`);
    console.log(`❌ ファイルが見つかりません: ${fileName}`);
    return null;
    
  } catch (error) {
    console.log(`❌ ファイル検索エラー: ${error.message}`);
    return null;
  }
}

/**
 * サブパスを使用してファイルを検索（動作していたバージョン）
 * @param {string} rootFolderId - ルートフォルダID
 * @param {string} subPath - サブパス（例: "h_久山町/03_共通/99_資料/02_返礼品関係/k_株式会社亀井通産"）
 * @param {string} fileName - ファイル名
 * @returns {string|null} ファイルID
 */
function findFileInFolderWithSubPath(rootFolderId, subPath, fileName) {
  try {
    console.log(`🔍 サブパスを使用したファイル検索開始: ${subPath}/${fileName}`);
    
    const pathParts = subPath.split('/').filter(part => part.trim() !== '');
    const fileNameFromPath = pathParts.pop(); // 最後の要素をファイル名として取得
    const folderPath = pathParts; // 残りをフォルダパスとして使用
    
    console.log(`📁 ルートフォルダ: ${rootFolderId}`);
    console.log(`📂 フォルダパス: ${folderPath.join(' > ')}`);
    console.log(`📄 ファイル名: ${fileNameFromPath}`);
    
    let currentFolder = DriveApp.getFolderById(rootFolderId);
    
    // フォルダをたどる
    for (let i = 0; i < folderPath.length; i++) {
      const folderName = folderPath[i];
      console.log(`🔍 フォルダ検索中: ${folderName}`);
      
      const subFolders = currentFolder.getFoldersByName(folderName);
      if (!subFolders.hasNext()) {
        console.log(`❌ フォルダが見つかりません: ${folderName}`);
        return null;
      }
      
      currentFolder = subFolders.next();
      console.log(`✅ フォルダ発見: ${currentFolder.getName()} (ID: ${currentFolder.getId()})`);
    }
    
    // ファイル検索
    console.log(`🔍 最終フォルダ内でファイル検索: ${fileName}`);
    const files = currentFolder.getFilesByName(fileName);
    
    if (files.hasNext()) {
      const file = files.next();
      console.log(`✅ ファイル発見: ${fileName} (ID: ${file.getId()})`);
      return file.getId();
    } else {
      console.log(`❌ ファイルが見つかりません: ${fileName}`);
      return null;
    }
    
  } catch (error) {
    console.log(`❌ サブパス検索エラー: ${error.message}`);
    return null;
  }
}

/**
 * 列番号を列文字に変換
 * @param {number} columnNumber - 列番号（1から開始）
 * @returns {string} 列文字（A, B, C...）
 */
function getColumnLetter(columnNumber) {
  let result = '';
  while (columnNumber > 0) {
    columnNumber--;
    result = String.fromCharCode(65 + (columnNumber % 26)) + result;
    columnNumber = Math.floor(columnNumber / 26);
  }
  return result;
}

/**
 * 列文字を列番号に変換
 * @param {string} columnLetter - 列文字（A, B, C...）
 * @returns {number} 列番号（1から開始）
 */
function getColumnNumber(columnLetter) {
  let result = 0;
  for (let i = 0; i < columnLetter.length; i++) {
    result *= 26;
    result += columnLetter.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
  }
  return result;
}

/**
 * 結合セルの判定
 * @param {Sheet} sheet - スプレッドシート
 * @param {number} row - 行番号
 * @param {number} col - 列番号
 * @returns {boolean} 結合セルかどうか
 */
function isMergedCell(sheet, row, col) {
  try {
    const range = sheet.getRange(row, col);
    const mergedRanges = range.getMergedRanges();
    return mergedRanges.length > 0;
  } catch (error) {
    console.log(`結合セル判定エラー (${row}, ${col}): ${error.message}`);
    return false;
  }
}

/**
 * 結合セルの値を取得
 * @param {Sheet} sheet - スプレッドシート
 * @param {number} row - 行番号
 * @param {number} col - 列番号
 * @returns {*} セルの値
 */
function getMergedCellValue(sheet, row, col) {
  try {
    const range = sheet.getRange(row, col);
    const mergedRanges = range.getMergedRanges();
    
    if (mergedRanges.length > 0) {
      // 結合セルの場合、左上のセルの値を取得
      const mergedRange = mergedRanges[0];
      const value = sheet.getRange(mergedRange.getRow(), mergedRange.getColumn()).getValue();
      console.log(`結合セル (${row}, ${col}) → (${mergedRange.getRow()}, ${mergedRange.getColumn()}): "${value}"`);
      return value;
    } else {
      // 通常セルの場合
      const value = range.getValue();
      return value;
    }
  } catch (error) {
    console.log(`結合セル値取得エラー (${row}, ${col}): ${error.message}`);
    return '';
  }
}

/**
 * ファイルパスから必要な情報を取得してファイルIDを特定
 * @returns {Object} ファイル情報
 */
function resolveFilePathToFileId() {
  try {
    console.log('🔍 ファイルパス解決開始');
    
    // ファイルパスを読み込み
    const folderPath = loadFolderPath();
    if (!folderPath) {
      throw new Error('ファイルパスが設定されていません');
    }
    
    // ファイルパスから自治体フォルダキーとサブパスを抽出
    const pathInfo = convertWindowsPathToDrivePath(folderPath);
    if (!pathInfo) {
      throw new Error('パス情報の抽出に失敗しました');
    }
    
    console.log(`📋 抽出されたパス情報:`);
    console.log(`  - 自治体フォルダキー: ${pathInfo.folderKey}`);
    console.log(`  - サブパス: ${pathInfo.subPath}`);
    
    // 自治体フォルダタブからフォルダIDを取得
    const folderId = findMunicipalityFolder(pathInfo.folderKey);
    if (!folderId) {
      throw new Error(`自治体フォルダ "${pathInfo.folderKey}" が見つかりません`);
    }
    
    // ファイル名を抽出
    const fileName = extractFileNameFromPath(folderPath);
    
    // サブパスを使用してファイルを検索
    const fileId = findFileInFolderWithSubPath(folderId, pathInfo.subPath, fileName);
    if (!fileId) {
      throw new Error(`ファイル "${fileName}" がパス "${pathInfo.subPath}" 内に見つかりません`);
    }
    
    console.log(`✅ ファイルID取得完了: ${fileId}`);
    
    return {
      fileId: fileId,
      fileName: fileName,
      folderId: folderId,
      pathInfo: pathInfo
    };
    
  } catch (error) {
    console.log(`❌ ファイルパス解決エラー: ${error.message}`);
    throw error;
  }
}

/**
 * ユーティリティ関数のテスト実行
 */
function testUtils() {
  try {
    console.log('=== ユーティリティ関数テスト開始 ===');
    
    // 列変換テスト
    console.log('列番号→列文字変換テスト:');
    console.log('1 →', getColumnLetter(1));  // A
    console.log('26 →', getColumnLetter(26)); // Z
    console.log('27 →', getColumnLetter(27)); // AA
    
    console.log('列文字→列番号変換テスト:');
    console.log('A →', getColumnNumber('A'));   // 1
    console.log('Z →', getColumnNumber('Z'));   // 26
    console.log('AA →', getColumnNumber('AA')); // 27
    
    // パス処理テスト
    const testPath = 'G:\\共有ドライブ\\h_40348_福岡県_久山町_01\\test.xlsx';
    console.log('パス変換テスト:', testPath);
    const folderKey = convertWindowsPathToDrivePath(testPath);
    console.log('抽出されたフォルダキー:', folderKey);
    
    const fileName = extractFileNameFromPath(testPath);
    console.log('抽出されたファイル名:', fileName);
    
    // ファイル検索テスト（実際のフォルダIDが必要）
    console.log('ファイル検索テスト: 実際のフォルダIDが必要です');
    
    console.log('=== ユーティリティ関数テスト完了 ===');
    
  } catch (error) {
    console.error('ユーティリティ関数テストエラー:', error.message);
  }
}

/**
 * resolveFilePathToFileId関数のテスト実行
 */
function testResolveFilePathToFileId() {
  try {
    console.log('=== resolveFilePathToFileId テスト開始 ===');
    
    // 設定値の確認
    console.log('設定:', JSON.stringify(CONFIG, null, 2));
    
    // 関数の存在確認
    console.log('関数存在確認:');
    console.log('- loadFolderPath:', typeof loadFolderPath);
    console.log('- convertWindowsPathToDrivePath:', typeof convertWindowsPathToDrivePath);
    console.log('- findMunicipalityFolder:', typeof findMunicipalityFolder);
    console.log('- extractFileNameFromPath:', typeof extractFileNameFromPath);
    console.log('- findFileInFolderWithSubPath:', typeof findFileInFolderWithSubPath);
    console.log('- resolveFilePathToFileId:', typeof resolveFilePathToFileId);
    
    // 実際の処理を実行
    const result = resolveFilePathToFileId();
    console.log('結果:', result);
    
    console.log('=== resolveFilePathToFileId テスト完了 ===');
    
  } catch (error) {
    console.error('resolveFilePathToFileId テストエラー:', error.message);
    console.error('スタックトレース:', error.stack);
  }
}
