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
 * 結合セル内の情報を統合して取得（行・列方向の結合セル対応）
 * @param {Sheet} sheet - スプレッドシート
 * @param {number} row - 行番号
 * @param {number} col - 列番号
 * @returns {string} 統合された値
 */
function getMergedCellValueWithMergeInfo(sheet, row, col) {
  try {
    // 結合セルかチェック
    if (!isMergedCell(sheet, row, col)) {
      // 結合セルでない場合は通常の値を取得
      return sheet.getRange(row, col).getValue() || '';
    }
    
    // 結合セルの範囲を取得
    const mergedRanges = sheet.getRange(row, col).getMergedRanges();
    if (mergedRanges.length === 0) {
      return sheet.getRange(row, col).getValue() || '';
    }
    
    const mergedRange = mergedRanges[0];
    const startRow = mergedRange.getRow();
    const endRow = mergedRange.getLastRow();
    const startCol = mergedRange.getColumn();
    const endCol = mergedRange.getLastColumn();
    
    // 現在のセルが結合セルの開始位置でない場合は空値を返す
    // これにより、結合セル内の重複データを防ぐ
    if (row !== startRow || col !== startCol) {
      return '';
    }
    
    // 結合セル内のすべての値を収集
    const allValues = [];
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        const value = sheet.getRange(r, c).getValue();
        if (value && value !== '') {
          allValues.push(value.toString().trim());
        }
      }
    }
    
    // 値が1つの場合はそのまま返す
    if (allValues.length === 1) {
      return allValues[0];
    }
    
    // 複数の値がある場合は統合
    if (allValues.length > 1) {
      // 矢印（→）で区切られている場合は適切に処理
      const joinedValue = allValues.join(' ');
      
      // 矢印の前後の空白を調整
      const normalizedValue = joinedValue
        .replace(/\s*→\s*/g, ' → ')  // 矢印の前後の空白を統一
        .replace(/\s+/g, ' ')          // 連続する空白を1つに
        .trim();
      
      // 結合セルの範囲情報をログ出力
      const isRowMerged = (endRow - startRow) > 0;
      const isColMerged = (endCol - startCol) > 0;
      const mergeType = isRowMerged && isColMerged ? '行・列結合' : 
                       isRowMerged ? '行結合' : 
                       isColMerged ? '列結合' : '単一セル';
      
      console.log(`🔗 結合セル統合: 行${row}列${getColumnLetter(col)} (${mergeType}) → "${normalizedValue}"`);
      console.log(`  範囲: 行${startRow}-${endRow}, 列${getColumnLetter(startCol)}-${getColumnLetter(endCol)}`);
      
      return normalizedValue;
    }
    
    return '';
    
  } catch (error) {
    console.log(`❌ 結合セル値取得エラー: 行${row}列${getColumnLetter(col)} - ${error.message}`);
    return '';
  }
}

/**
 * 結合セルの情報をキャッシュして効率的に処理
 * @param {Sheet} sheet - スプレッドシート
 * @param {number} row - 行番号
 * @param {number} col - 列番号
 * @returns {string} 統合された値
 */
function getMergedCellValueWithMergeInfoOptimized(sheet, row, col) {
  try {
    // 結合セルかチェック
    if (!isMergedCell(sheet, row, col)) {
      // 結合セルでない場合は通常の値を取得
      return sheet.getRange(row, col).getValue() || '';
    }
    
    // 結合セルの範囲を取得
    const mergedRanges = sheet.getRange(row, col).getMergedRanges();
    if (mergedRanges.length === 0) {
      return sheet.getRange(row, col).getValue() || '';
    }
    
    const mergedRange = mergedRanges[0];
    const startRow = mergedRange.getRow();
    const endRow = mergedRange.getLastRow();
    const startCol = mergedRange.getColumn();
    const endCol = mergedRange.getLastColumn();
    
    // 現在のセルが結合セルの開始位置でない場合は空値を返す
    // これにより、結合セル内の重複データを防ぐ
    if (row !== startRow || col !== startCol) {
      return '';
    }
    
    // 結合セル内のすべての値を収集（効率化）
    const allValues = [];
    
    // 行方向の結合セルの場合
    if (endRow > startRow) {
      for (let r = startRow; r <= endRow; r++) {
        const value = sheet.getRange(r, startCol).getValue();
        if (value && value !== '') {
          allValues.push(value.toString().trim());
        }
      }
    }
    
    // 列方向の結合セルの場合
    if (endCol > startCol) {
      for (let c = startCol; c <= endCol; c++) {
        const value = sheet.getRange(startRow, c).getValue();
        if (value && value !== '') {
          allValues.push(value.toString().trim());
        }
      }
    }
    
    // 単一セルの場合
    if (startRow === endRow && startCol === endCol) {
      const value = sheet.getRange(startRow, startCol).getValue();
      return value || '';
    }
    
    // 値が1つの場合はそのまま返す
    if (allValues.length === 1) {
      return allValues[0];
    }
    
    // 複数の値がある場合は統合
    if (allValues.length > 1) {
      // 矢印（→）で区切られている場合は適切に処理
      const joinedValue = allValues.join(' ');
      
      // 矢印の前後の空白を調整
      const normalizedValue = joinedValue
        .replace(/\s*→\s*/g, ' → ')  // 矢印の前後の空白を統一
        .replace(/\s+/g, ' ')          // 連続する空白を1つに
        .trim();
      
      // 結合セルの範囲情報をログ出力
      const isRowMerged = (endRow - startRow) > 0;
      const isColMerged = (endCol - startCol) > 0;
      const mergeType = isRowMerged && isColMerged ? '行・列結合' : 
                       isRowMerged ? '行結合' : 
                       isColMerged ? '列結合' : '単一セル';
      
      console.log(`🔗 結合セル統合: 行${row}列${getColumnLetter(col)} (${mergeType}) → "${normalizedValue}"`);
      console.log(`  範囲: 行${startRow}-${endRow}, 列${getColumnLetter(startCol)}-${getColumnLetter(endCol)}`);
      
      return normalizedValue;
    }
    
    return '';
    
  } catch (error) {
    console.log(`❌ 結合セル値取得エラー: 行${row}列${getColumnLetter(col)} - ${error.message}`);
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
 * 実際にデータが入っている列の最終位置を取得
 * @param {Sheet} sheet - スプレッドシート
 * @param {number} startCol - 開始列番号
 * @param {number} maxCol - 最大列番号（getLastColumn()の結果）
 * @returns {number} 実際のデータが入っている最終列番号
 */
function getActualLastColumn(sheet, startCol, maxCol) {
  try {
    console.log(`🔍 有効列判定開始: ${startCol}列目から${maxCol}列目まで`);
    
    let actualLastCol = startCol - 1; // 開始列の前から開始
    
    // 各列のデータ有無をチェック
    for (let col = startCol; col <= maxCol; col++) {
      if (hasDataInColumn(sheet, col)) {
        actualLastCol = col;
      }
    }
    
    console.log(`✅ 有効列判定完了: 実際の最終列 = ${actualLastCol}列目 (${getColumnLetter(actualLastCol)})`);
    return actualLastCol;
    
  } catch (error) {
    console.log(`❌ 有効列判定エラー: ${error.message}`);
    return startCol; // エラーの場合は開始列を返す
  }
}

/**
 * 連続データの検出による最終列判定（高度版）
 * @param {Sheet} sheet - スプレッドシート
 * @param {number} startCol - 開始列番号
 * @param {number} maxCol - 最大列番号
 * @returns {number} 実際のデータが入っている最終列番号
 */
function getActualLastColumnAdvanced(sheet, startCol, maxCol) {
  try {
    console.log(`🔍 高度な有効列判定開始: ${startCol}列目から${maxCol}列目まで`);
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 4) {
      console.log(`⚠️ データ行が不足しています（最小4行必要）`);
      return startCol;
    }
    
    let actualLastCol = startCol - 1;
    let emptyColumnCount = 0;
    const maxEmptyColumns = 3; // 連続で空の列が3列以上続いたら終了
    
    // 各列のデータ有無をチェック
    for (let col = startCol; col <= maxCol; col++) {
      if (hasDataInColumn(sheet, col)) {
        actualLastCol = col;
        emptyColumnCount = 0; // データがある列が見つかったらリセット
        
        // データの詳細をログ出力
        const dataInfo = getColumnDataInfo(sheet, col);
        console.log(`  ✅ 列${getColumnLetter(col)}: データあり (${dataInfo.dataCount}行のデータ)`);
        
      } else {
        emptyColumnCount++;
        if (emptyColumnCount <= 2) {
          console.log(`  ⚪ 列${getColumnLetter(col)}: データなし (空列${emptyColumnCount}連続)`);
        }
        
        // 連続で空の列が一定数続いたら、それ以降は処理しない
        if (emptyColumnCount >= maxEmptyColumns) {
          console.log(`  🛑 列${getColumnLetter(col)}以降: 連続空列${maxEmptyColumns}列のため処理終了`);
          console.log(`  📍 最終有効列: 列${getColumnLetter(actualLastCol)} (${actualLastCol}列目)`);
          break;
        }
      }
    }
    
    console.log(`✅ 高度な有効列判定完了: 実際の最終列 = ${actualLastCol}列目 (${getColumnLetter(actualLastCol)})`);
    console.log(`📊 判定結果: ${startCol}列目〜${actualLastCol}列目 (${actualLastCol - startCol + 1}列)`);
    
    return actualLastCol;
    
  } catch (error) {
    console.log(`❌ 高度な有効列判定エラー: ${error.message}`);
    return startCol; // エラーの場合は開始列を返す
  }
}

/**
 * 指定列のデータ情報を取得
 * @param {Sheet} sheet - スプレッドシート
 * @param {number} col - 列番号
 * @returns {Object} 列のデータ情報
 */
function getColumnDataInfo(sheet, col) {
  try {
    const lastRow = sheet.getLastRow();
    const dataStartRow = 4;
    const dataEndRow = lastRow;
    
    if (dataEndRow < dataStartRow) {
      return { dataCount: 0, sampleData: [] };
    }
    
    // データ行のみの値を取得
    const columnRange = sheet.getRange(dataStartRow, col, dataEndRow - dataStartRow + 1, 1);
    const columnValues = columnRange.getValues();
    
    let dataCount = 0;
    const sampleData = [];
    
    // データ行で空でない値をカウント
    for (let row = 0; row < columnValues.length; row++) {
      const value = columnValues[row][0];
      if (value !== null && value !== undefined && value !== '') {
        if (typeof value === 'string' && value.trim() !== '') {
          dataCount++;
          if (sampleData.length < 3) { // 最初の3件をサンプルとして保存
            sampleData.push(value.toString().trim());
          }
        } else if (typeof value !== 'string') {
          dataCount++;
          if (sampleData.length < 3) {
            sampleData.push(value.toString());
          }
        }
      }
    }
    
    return {
      dataCount: dataCount,
      sampleData: sampleData
    };
    
  } catch (error) {
    console.log(`❌ 列データ情報取得エラー: 列${getColumnLetter(col)} - ${error.message}`);
    return { dataCount: 0, sampleData: [] };
  }
}

/**
 * 指定列にデータが入っているかを判定
 * @param {Sheet} sheet - スプレッドシート
 * @param {number} col - 列番号
 * @returns {boolean} データが入っているかどうか
 */
function hasDataInColumn(sheet, col) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow === 0) return false;
    
    // 列全体の値を取得
    const columnRange = sheet.getRange(1, col, lastRow, 1);
    const columnValues = columnRange.getValues();
    
    // 空でない値があるかをチェック
    for (let row = 0; row < columnValues.length; row++) {
      const value = columnValues[row][0];
      if (value !== null && value !== undefined && value !== '') {
        // 空白文字のみの場合は除外
        if (typeof value === 'string' && value.trim() !== '') {
          return true;
        } else if (typeof value !== 'string') {
          return true;
        }
      }
    }
    
    return false;
    
  } catch (error) {
    console.log(`❌ 列データ判定エラー: 列${getColumnLetter(col)} - ${error.message}`);
    return false;
  }
}

/**
 * 有効列の詳細情報をログ出力
 * @param {Sheet} sheet - スプレッドシート
 * @param {number} startCol - 開始列番号
 * @param {number} maxCol - 最大列番号
 */
function logColumnAnalysis(sheet, startCol, maxCol) {
  try {
    console.log(`📊 列分析結果:`);
    console.log(`  - 開始列: ${startCol}列目 (${getColumnLetter(startCol)})`);
    console.log(`  - 理論上の最終列: ${maxCol}列目 (${getColumnLetter(maxCol)})`);
    
    let dataColumnCount = 0;
    let emptyColumnCount = 0;
    
    for (let col = startCol; col <= maxCol; col++) {
      if (hasDataInColumn(sheet, col)) {
        dataColumnCount++;
        console.log(`  ✅ 列${getColumnLetter(col)}: データあり`);
      } else {
        emptyColumnCount++;
        if (emptyColumnCount <= 5) { // 最初の5列のみログ出力
          console.log(`  ⚪ 列${getColumnLetter(col)}: データなし`);
        } else if (emptyColumnCount === 6) {
          console.log(`  ... 他${maxCol - startCol - 4}列はデータなし`);
        }
      }
    }
    
    console.log(`📈 統計:`);
    console.log(`  - データあり: ${dataColumnCount}列`);
    console.log(`  - データなし: ${emptyColumnCount}列`);
    console.log(`  - 処理対象列: ${dataColumnCount}列`);
    
  } catch (error) {
    console.log(`❌ 列分析エラー: ${error.message}`);
  }
}
