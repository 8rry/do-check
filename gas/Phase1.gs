/**
 * Phase 1: 基本データ抽出
 * 
 * パフォーマンス改善:
 * - バッチ処理による一括操作
 * - キャッシュ機能による重複処理回避
 * - メモリ効率化
 */

/**
 * Phase 1実行
 * @param {string} fileId - ファイルID
 * @param {string} fileName - ファイル名
 * @returns {Object} 処理結果
 */
function executePhase1(fileId, fileName) {
  try {
    console.log('=== Phase 1: 基本データ抽出開始 ===');
    const startTime = new Date();
    
    // ExcelファイルをGoogle Sheetsに変換
    const extractedData = processExcelFileForPhase1(fileId, fileName);
    
    // 情報抽出タブに商品データを出力
    const result = outputProductDataToInfoExtractionTabOptimized(extractedData);
    
    const endTime = new Date();
    const processingTime = endTime - startTime;
    
    if (result.success) {
      console.log(`✅ 情報抽出タブへの出力完了: ${result.outputRows}行（${processingTime}ms）`);
    } else {
      console.log(`❌ 情報抽出タブへの出力エラー: ${result.error}`);
      throw new Error(result.error);
    }
    
    return {
      success: true,
      data: extractedData.data,
      sheet: extractedData.sheet,
      spreadsheetId: extractedData.spreadsheetId,
      processingTime: processingTime
    };
    
  } catch (error) {
    console.log(`❌ Phase 1エラー: ${error.message}`);
    throw error;
  }
}

/**
 * Excelファイル処理
 * @param {string} fileId - ファイルID
 * @param {string} fileName - ファイル名
 * @returns {Object} 抽出されたデータ
 */
function processExcelFileForPhase1(fileId, fileName) {
  let tempFileId = null;
  
  try {
    console.log(`📊 Excel処理開始: ${fileName}`);
    
    // Drive APIを使用した変換
    console.log('🔄 Drive APIを使用した変換を試行中...');
    
    const excelFile = DriveApp.getFileById(fileId);
    const blob = excelFile.getBlob();
    
    const tempFileName = 'temp_' + fileName.replace(/[^a-zA-Z0-9]/g, '_') + '_' + new Date().getTime();
    
    // バッチ処理による変換
    const resource = {
      title: tempFileName,
      mimeType: MimeType.GOOGLE_SHEETS
    };
    
    const convertedFile = Drive.Files.insert(resource, blob, {
      convert: true
    });
    
    tempFileId = convertedFile.id;
    console.log(`✅ Drive API変換完了: ${tempFileId}`);
    
    // 変換されたスプレッドシートを開く
    const ss = SpreadsheetApp.openById(tempFileId);
    const sheet = ss.getActiveSheet();
    
    // データ抽出
    const extractedData = extractProductDataFromSheet(sheet, tempFileId);
    
    return {
      ...extractedData,
      spreadsheetId: tempFileId
    };
    
  } catch (error) {
    console.log(`❌ Excel処理エラー: ${error.message}`);
    
    // Drive APIのエラーの場合、詳細情報を提供
    if (error.message.includes('Invalid mime type')) {
      console.log(`💡 解決方法: Drive APIの権限を確認してください`);
    }
    
    throw error;
  }
}

/**
 * スプレッドシートから商品名データを抽出
 * @param {Sheet} sheet - スプレッドシート
 * @param {string} tempFileId - 一時ファイルID
 * @returns {Object} 抽出されたデータとシート情報
 */
function extractProductDataFromSheet(sheet, tempFileId) {
  try {
    console.log(`🔍 商品名列の特定開始`);
    
    // キャッシュ機能による高速化
    const cacheKey = `product_data_${sheet.getSheetId()}_${sheet.getLastRow()}`;
    const cachedData = getCachedData(cacheKey);
    if (cachedData) {
      console.log('⚡ キャッシュヒット: 高速データ取得');
      return cachedData;
    }
    
    let productNameColumn = null;
    let productNameRow = null;
    
    // 検索処理
    const searchResults = findProductNameColumn(sheet);
    productNameColumn = searchResults.column;
    productNameRow = searchResults.row;
    
    if (!productNameColumn) {
      throw new Error('商品名を含む列が見つかりませんでした');
    }
    
    // 発見した列とその右隣列から4行目以降のデータを抽出
    const rightColumn = productNameColumn + 1;
    console.log(`📊 発見した列（${getColumnLetter(productNameColumn)}列）とその右隣列（${getColumnLetter(rightColumn)}列）からデータを抽出します`);
    
    const lastRow = sheet.getLastRow();
    const dataRange = sheet.getRange(4, productNameColumn, lastRow - 3, 2);
    const dataValues = dataRange.getValues();
    
    console.log(`📊 データ抽出範囲: ${getColumnLetter(productNameColumn)}4:${getColumnLetter(rightColumn)}${lastRow} (全${lastRow - 3}行)`);
    
    // 結合セル処理
    let mergedCellData = [];
    try {
      mergedCellData = processMergedCellsOptimized(sheet, productNameColumn, lastRow);
    } catch (error) {
      console.log(`⚠️ 結合セル処理でエラーが発生しました: ${error.message}`);
      console.log('ℹ️ 結合セル処理をスキップして通常のデータ処理を続行します。');
      mergedCellData = [];
    }
    
    // バッチ処理によるデータ構築
    const processedData = buildProductDataBatch(dataValues, mergedCellData, productNameColumn, rightColumn);
    
    // 結果をキャッシュに保存
    const result = {
      data: processedData,
      sheet: sheet,
      tempFileId: tempFileId
    };
    setCachedData(cacheKey, result);
    
    return result;
    
  } catch (error) {
    console.log(`❌ データ抽出エラー: ${error.message}`);
    throw error;
  }
}

/**
 * 商品名列を検索
 * @param {Sheet} sheet - スプレッドシート
 * @returns {Object} 検索結果
 */
function findProductNameColumn(sheet) {
  try {
    // 検索処理
    for (let col = 1; col <= 4; col++) {
      const result = searchColumnForProductName(sheet, col);
      if (result.found) {
        const colLetter = getColumnLetter(result.column);
        console.log(`✅ 商品名列発見: 列${colLetter}の${result.row}行目`);
        return result;
      }
    }
    
    throw new Error('最初の4列で商品名列が見つかりませんでした');
    
  } catch (error) {
    console.log(`❌ 商品名列検索エラー: ${error.message}`);
    throw error;
  }
}

/**
 * 単一列での商品名検索
 * @param {Sheet} sheet - スプレッドシート
 * @param {number} col - 列番号
 * @returns {Object} 検索結果
 */
function searchColumnForProductName(sheet, col) {
  try {
    const colLetter = getColumnLetter(col);
    
    // 最初の20行をチェック（高速化）
    for (let row = 1; row <= 20; row++) {
      const cellValue = sheet.getRange(row, col).getValue();
      
      if (cellValue && typeof cellValue === 'string' && cellValue.includes('商品名')) {
        return { found: true, column: col, row: row };
      }
    }
    
    return { found: false, column: col, row: null };
    
  } catch (error) {
    return { found: false, column: col, row: null };
  }
}

/**
 * 従来の順次検索（フォールバック）
 * @param {Sheet} sheet - スプレッドシート
 * @returns {Object} 検索結果
 */
function findProductNameColumnSequential(sheet) {
  try {
    for (let col = 1; col <= 4; col++) {
      const colLetter = getColumnLetter(col);
      console.log(`列${colLetter}をチェック中...`);
      
      for (let row = 1; row <= 20; row++) {
        const cellValue = sheet.getRange(row, col).getValue();
        
        if (cellValue && typeof cellValue === 'string' && cellValue.includes('商品名')) {
          console.log(`✅ 商品名列発見: 列${colLetter}の${row}行目`);
          return { found: true, column: col, row: row };
        }
      }
    }
    
    return { found: false, column: null, row: null };
    
  } catch (error) {
    console.log(`❌ 順次検索エラー: ${error.message}`);
    return { found: false, column: null, row: null };
  }
}

/**
 * 結合セル処理（超高速版）
 * @param {Sheet} sheet - スプレッドシート
 * @param {number} productNameColumn - 商品名列
 * @param {number} lastRow - 最終行
 * @returns {Object} 結合セルデータ
 */
function processMergedCellsOptimized(sheet, productNameColumn, lastRow) {
  try {
    console.log(`🔍 商品名列の結合セル状況をチェック中（超高速版）...`);
    
    // バッチ処理による結合セル情報取得
    let mergedRanges = [];
    try {
      // getMergedRanges()が利用可能かチェック
      if (typeof sheet.getMergedRanges === 'function') {
        mergedRanges = sheet.getMergedRanges();
      } else {
        console.log('⚠️ getMergedRanges()が利用できません。結合セル処理をスキップします。');
        return [];
      }
    } catch (error) {
      console.log(`⚠️ 結合セル情報取得エラー: ${error.message}。結合セル処理をスキップします。`);
      return [];
    }
    
    const mergedCellInfo = [];
    
    for (let i = 0; i < mergedRanges.length; i++) {
      const range = mergedRanges[i];
      if (range.getColumn() === productNameColumn) {
        const startRow = range.getRow();
        const endRow = startRow + range.getNumRows() - 1;
        
        if (startRow >= 4 && startRow <= lastRow) {
          mergedCellInfo.push({
            startRow: startRow,
            endRow: endRow,
            rowCount: range.getNumRows()
          });
        }
      }
    }
    
    console.log(`✅ 結合セル処理完了: ${mergedCellInfo.length}件の結合セルを検出`);
    return mergedCellInfo;
    
  } catch (error) {
    console.log(`❌ 結合セル処理エラー: ${error.message}`);
    return [];
  }
}

/**
 * バッチ処理による商品データ構築
 * @param {Array} dataValues - データ値
 * @param {Array} mergedCellInfo - 結合セル情報
 * @param {number} productNameColumn - 商品名列
 * @param {number} rightColumn - 右列
 * @returns {Array} 処理されたデータ
 */
function buildProductDataBatch(dataValues, mergedCellInfo, productNameColumn, rightColumn) {
  try {
    const processedData = [];
    let currentMergedIndex = 0;
    
    // 結合セル情報がない場合は通常の処理
    if (!mergedCellInfo || mergedCellInfo.length === 0) {
      console.log('ℹ️ 結合セル情報なし。通常のデータ処理を実行します。');
      for (let i = 0; i < dataValues.length; i++) {
        const row = i + 4; // 4行目から開始
        const productName = dataValues[i][0];
        const rightColumnValue = dataValues[i][1];
        
        processedData.push({
          productName: productName || '',
          rightColumn: rightColumnValue || '',
          row: row
        });
      }
      return processedData;
    }
    
    // 結合セル情報がある場合の処理
    for (let i = 0; i < dataValues.length; i++) {
      const row = i + 4; // 4行目から開始
      const productName = dataValues[i][0];
      const rightColumnValue = dataValues[i][1];
      
      // 結合セルの処理
      let actualProductName = productName;
      let shouldAddEmptyRows = false;
      
      if (currentMergedIndex < mergedCellInfo.length) {
        const mergedInfo = mergedCellInfo[currentMergedIndex];
        if (row === mergedInfo.startRow) {
          // 結合セルの開始行
          shouldAddEmptyRows = mergedInfo.rowCount > 1;
          currentMergedIndex++;
        } else if (row > mergedInfo.startRow && row <= mergedInfo.endRow) {
          // 結合セルの継続行
          actualProductName = '';
        }
      }
      
      // データを追加
      processedData.push({
        productName: actualProductName || '',
        rightColumn: rightColumnValue || '',
        row: row
      });
      
      // 空行を追加（必要に応じて）
      if (shouldAddEmptyRows) {
        for (let j = 1; j < mergedInfo.rowCount; j++) {
          processedData.push({
            productName: '',
            rightColumn: '',
            row: row + j
          });
        }
      }
    }
    
    console.log(`✅ バッチデータ構築完了: ${processedData.length}行のデータを処理`);
    return processedData;
    
  } catch (error) {
    console.log(`❌ バッチデータ構築エラー: ${error.message}`);
    // エラーが発生した場合は元のデータをそのまま返す
    console.log('🔄 フォールバック: 元のデータをそのまま返します。');
    return dataValues.map((rowData, index) => ({
      productName: rowData[0] || '',
      rightColumn: rowData[1] || '',
      row: index + 4
    }));
  }
}

/**
 * 情報抽出タブへの出力（超高速版）
 * @param {Object|Array} extractedData - 抽出されたデータ
 * @returns {Object} 処理結果
 */
function outputProductDataToInfoExtractionTabOptimized(extractedData) {
  try {
    console.log('🚀 情報抽出タブに商品データを出力開始（超高速版）');
    
    // 現在アクティブなスプレッドシートを取得
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      console.log('⚠️ アクティブなスプレッドシートが見つかりません');
      return null;
    }
    const infoSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.INFO_EXTRACTION);
    
    if (!infoSheet) {
      throw new Error('情報抽出タブが見つかりません');
    }
    
    // データの準備
    let outputData = [];
    if (Array.isArray(extractedData)) {
      outputData = extractedData.map(item => [item.productName, item.rightColumn]);
    } else if (extractedData.data && Array.isArray(extractedData.data)) {
      outputData = extractedData.data.map(item => [item.productName, item.rightColumn]);
    } else {
      throw new Error('無効なデータ形式です');
    }
    
    if (outputData.length === 0) {
      console.log('⚠️ 出力データがありません');
      return { success: false, outputRows: 0, error: '出力データがありません' };
    }
    
    // 既存データのクリア（バッチ処理）
    const lastRow = infoSheet.getLastRow();
    if (lastRow >= CONFIG.OUTPUT.START_ROW) {
      const clearRange = infoSheet.getRange(CONFIG.OUTPUT.START_ROW, CONFIG.OUTPUT.COL_B, lastRow - CONFIG.OUTPUT.START_ROW + 1, 2);
      clearRange.clear();
      console.log(`🗑️ 既存データクリア完了: ${CONFIG.OUTPUT.START_ROW}行目〜${lastRow}行目`);
    } else {
      console.log('ℹ️ クリア対象のデータなし');
    }
    
    // 超高速バッチ処理でデータを一括出力
    console.log('📤 データ出力開始（超高速バッチ処理）');
    if (outputData.length > 0) {
      const outputRange = infoSheet.getRange(
        CONFIG.OUTPUT.START_ROW, 
        CONFIG.OUTPUT.COL_B, 
        outputData.length, 
        2
      );
      outputRange.setValues(outputData);
      console.log(`✅ データ出力完了: ${CONFIG.OUTPUT.START_ROW}行目〜${CONFIG.OUTPUT.START_ROW + outputData.length - 1}行目`);
    }
    
    console.log('🎉 情報抽出タブへの商品データ出力完了（超高速版）');
    
    return {
      success: true,
      outputRows: outputData.length,
      message: `${outputData.length}行のデータを出力しました`
    };
    
  } catch (error) {
    console.log(`❌ 情報抽出タブへの商品データ出力エラー: ${error.message}`);
    console.log(`🔍 エラー詳細: ${error.toString()}`);
    return {
      success: false,
      outputRows: 0,
      error: error.message
    };
  }
}

// キャッシュ機能
const CACHE = {};

/**
 * キャッシュからデータを取得
 * @param {string} key - キャッシュキー
 * @returns {*} キャッシュされたデータ
 */
function getCachedData(key) {
  return CACHE[key] || null;
}

/**
 * データをキャッシュに保存
 * @param {string} key - キャッシュキー
 * @param {*} data - 保存するデータ
 */
function setCachedData(key, data) {
  CACHE[key] = data;
}
