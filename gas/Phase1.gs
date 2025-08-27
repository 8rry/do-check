/**
 * Phase 1: 基本データ抽出
 * 返礼品シートから商品名を含む列を自動特定し、データを抽出する
 */

/**
 * Phase 1のメイン処理
 * @param {string} fileId - ExcelファイルのID
 * @param {string} fileName - ファイル名
 * @returns {Object} 抽出されたデータ
 */
function executePhase1(fileId, fileName) {
  try {
    console.log('=== Phase 1: 基本データ抽出開始 ===');
    
    // ExcelファイルをGoogle Sheetsに変換
    const extractedData = processExcelFileForPhase1(fileId, fileName);
    
    // 情報抽出タブに出力（dataプロパティを渡す）
    const result = outputToInfoExtractionTab(extractedData);
    
    if (result.success) {
      console.log(`✅ 情報抽出タブへの出力完了: ${result.outputRows}行`);
    } else {
      console.log(`❌ 情報抽出タブへの出力失敗: ${result.error}`);
    }
    
    console.log('=== Phase 1: 基本データ抽出完了 ===');
    return extractedData;  // 元のデータ構造を返す
    
  } catch (error) {
    console.log(`❌ Phase 1エラー: ${error.message}`);
    throw error;
  }
}

/**
 * Excelファイルを処理（Phase 1用）
 * @param {string} fileId - ExcelファイルのID
 * @param {string} fileName - ファイル名
 * @returns {Array} 抽出されたデータ
 */
function processExcelFileForPhase1(fileId, fileName) {
  let tempFileId = null;
  try {
    console.log(`📊 Excel処理開始: ${fileName}`);
    
    // ExcelファイルをGoogle Sheetsに変換
    const excelFile = DriveApp.getFileById(fileId);
    const blob = excelFile.getBlob();
    const tempFileName = 'temp_' + fileName.replace(/[^a-zA-Z0-9]/g, '_') + '_' + new Date().getTime();
    
    let tempFileId = null;
    
    try {
             // 方法1: Drive APIを使用した変換（動作していたバージョン）
       console.log(`🔄 Drive APIを使用した変換を試行中...`);
       
       const resource = {
         title: tempFileName,
         mimeType: MimeType.GOOGLE_SHEETS
       };
       
       const convertedFile = Drive.Files.insert(resource, blob, { 
         convert: true
       });
       tempFileId = convertedFile.id;
       console.log(`✅ Drive API変換完了: ${tempFileId}`);
      
    } catch (driveApiError) {
      console.log(`⚠️ Drive API変換失敗: ${driveApiError.message}`);
      throw new Error(`Drive API変換に失敗しました: ${driveApiError.message}`);
    }
    
    // 変換されたスプレッドシートを開く
    const ss = SpreadsheetApp.openById(tempFileId);
    const sheet = ss.getActiveSheet();
    
    // 商品名データを抽出
    const extractedData = extractProductDataFromSheet(sheet, tempFileId);
    
    return extractedData;
    
  } catch (error) {
    console.log(`❌ Excel処理エラー: ${error.message}`);
    console.log(`🔍 エラー詳細: ${error.toString()}`);
    
    // Drive APIのエラーの場合、詳細情報を提供
    if (error.message.includes('Invalid mime type')) {
      console.log(`💡 解決方法: Drive APIの権限を確認してください`);
    }
    
    throw error;
  } finally {
    // 一時ファイルを削除
    if (tempFileId) {
      try {
        DriveApp.getFileById(tempFileId).setTrashed(true);
        console.log(`🗑️ 一時ファイルを削除: ${tempFileId}`);
      } catch (deleteError) {
        console.log(`⚠️ 一時ファイル削除エラー: ${deleteError.toString()}`);
      }
    }
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
    
    let productNameColumn = null;
    let productNameRow = null;
    
    // A列〜D列で「商品名」を含むセルを検索
    for (let col = 1; col <= 4; col++) {
      const colLetter = getColumnLetter(col);
      console.log(`列${colLetter}をチェック中...`);
      
      for (let row = 1; row <= 20; row++) { // 最初の20行をチェック
        const cellValue = sheet.getRange(row, col).getValue();
        
        if (cellValue && typeof cellValue === 'string' && cellValue.includes('商品名')) {
          productNameColumn = col;
          productNameRow = row;
          console.log(`✅ 商品名列発見: 列${colLetter}の${row}行目`);
          break;
        }
      }
      
      if (productNameColumn) break;
    }
    
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
    
    // 結合セルの状況を事前チェック（右列データ状況も考慮）
    console.log(`🔍 商品名列の結合セル状況をチェック中...`);
    let totalMergedRows = 0;
    let actualAddedRows = 0;
    
    for (let row = 4; row <= lastRow; row++) {
      if (isMergedCell(sheet, row, productNameColumn)) {
        const mergedRange = sheet.getRange(row, productNameColumn).getMergedRanges()[0];
        const mergedRowCount = mergedRange.getNumRows();
        
        // 右列のデータ状況を確認
        let hasRightColumnData = false;
        for (let j = 0; j < mergedRowCount; j++) {
          const currentRow = row + j;
          const rightColumnValue = sheet.getRange(currentRow, productNameColumn + 1).getValue();
          if (rightColumnValue && rightColumnValue !== '') {
            hasRightColumnData = true;
            break;
          }
        }
        
        if (hasRightColumnData) {
          console.log(`  - 行${row}: ${mergedRowCount}行結合 (右列データあり)`);
        } else {
          console.log(`  - 行${row}: ${mergedRowCount}行結合 (右列データなし)`);
          actualAddedRows += mergedRowCount - 1;
        }
        
        totalMergedRows += mergedRowCount - 1; // 理論上の追加行数
      }
    }
    
    console.log(`📏 結合セル状況:`);
    console.log(`  - 理論上の追加行数: ${totalMergedRows}行`);
    console.log(`  - 実際の追加行数: ${actualAddedRows}行 (右列データ状況を考慮)`);
    
    // 空行が5行連続したら抽出を終了
    let emptyRowCount = 0;
    let extractedData = [];
    
    for (let i = 0; i < dataValues.length; i++) {
      const row = dataValues[i];
      const productName = row[0];
      const rightColumnValue = row[1];
      
      // 空行チェック
      if (!productName && !rightColumnValue) {
        emptyRowCount++;
        if (emptyRowCount >= 5) {
          console.log(`✅ 空行が5行連続: 行${4 + i}で抽出終了`);
          break;
        }
      } else {
        emptyRowCount = 0;
      }
      
      // 結合セルの場合、右列のデータ状況を確認して処理を決定
      if (productName && productName !== '') {
        // 商品名列が結合セルかチェック
        const actualRow = 4 + i; // 実際の行番号
        const isMerged = isMergedCell(sheet, actualRow, productNameColumn);
        
        if (isMerged) {
          // 結合セルの行数を取得
          const mergedRange = sheet.getRange(actualRow, productNameColumn).getMergedRanges()[0];
          const mergedRowCount = mergedRange.getNumRows();
          
          console.log(`🔗 結合セル検出: 行${actualRow}列${getColumnLetter(productNameColumn)} (${mergedRowCount}行結合)`);
          
          // 右列のデータ状況を確認
          let hasRightColumnData = false;
          for (let j = 0; j < mergedRowCount; j++) {
            const currentRow = actualRow + j;
            const rightColumnValue = sheet.getRange(currentRow, productNameColumn + 1).getValue();
            if (rightColumnValue && rightColumnValue !== '') {
              hasRightColumnData = true;
              break;
            }
          }
          
          if (hasRightColumnData) {
            console.log(`  ✅ 右列にデータあり → 結合セル内の各行を個別処理`);
            // 右列にデータがある場合も、結合セル内の各行を処理
            for (let j = 0; j < mergedRowCount; j++) {
              const currentRow = actualRow + j;
              const currentProductName = (j === 0) ? productName : ''; // 最初の行のみ商品名を設定
              const currentRightColumn = sheet.getRange(currentRow, productNameColumn + 1).getValue() || '';
              
              extractedData.push({
                productName: currentProductName,
                rightColumn: currentRightColumn
              });
            }
            
            // 結合セルの分だけスキップ
            i += mergedRowCount - 1;
            continue;
          } else {
            console.log(`  ⚠️ 右列にデータなし → 結合セル分の行を追加`);
            // 結合されている分の行を空で追加
            for (let j = 0; j < mergedRowCount; j++) {
              const currentRow = actualRow + j;
              const currentProductName = (j === 0) ? productName : ''; // 最初の行のみ商品名を設定
              const currentRightColumn = sheet.getRange(currentRow, productNameColumn + 1).getValue() || '';
              
              extractedData.push({
                productName: currentProductName,
                rightColumn: currentRightColumn
              });
            }
            
            // 結合セルの分だけスキップ
            i += mergedRowCount - 1;
            continue;
          }
        }
      }
      
      // 通常の行または結合セルでない場合
      extractedData.push({
        productName: productName || '',
        rightColumn: rightColumnValue || ''
      });
    }
    
    console.log(`📝 抽出データ: ${extractedData.length}行のデータを取得 (結合セル対応済み)`);
    console.log(`📊 データ詳細:`);
    console.log(`  - 元データ行数: ${lastRow - 3}行`);
    console.log(`  - 理論上の追加行数: ${totalMergedRows}行`);
    console.log(`  - 実際の追加行数: ${actualAddedRows}行 (右列データ状況を考慮)`);
    console.log(`  - 最終抽出行数: ${extractedData.length}行`);
    
    return {
      data: extractedData,
      sheet: sheet, // スプレッドシートオブジェクトも含める
      spreadsheetId: tempFileId
    };
    
  } catch (error) {
    console.log(`❌ データ抽出エラー: ${error.message}`);
    throw error;
  }
}

/**
 * 情報抽出タブのB8、C8以降にデータを格納
 * @param {Object|Array} extractedData - 抽出されたデータ（配列またはオブジェクト）
 * @returns {Object} 処理結果
 */
function outputToInfoExtractionTab(extractedData) {
  try {
    console.log('🚀 情報抽出タブへの出力開始');
    
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const infoSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.INFO_EXTRACTION);
    
    if (!infoSheet) {
      console.log('❌ 情報抽出タブが見つかりません');
      return { success: false, error: '情報抽出タブが見つかりません' };
    }
    
    console.log(`📋 対象シート: ${infoSheet.getName()}`);
    
    // データ構造を確認して適切に処理
    let dataArray = extractedData;
    if (extractedData && typeof extractedData === 'object' && extractedData.data) {
      // { data, sheet, spreadsheetId } の形式の場合
      dataArray = extractedData.data;
      console.log(`📊 オブジェクト形式のデータを検出: ${dataArray.length}行`);
    } else if (Array.isArray(extractedData)) {
      // 配列形式の場合
      dataArray = extractedData;
      console.log(`📊 配列形式のデータを検出: ${dataArray.length}行`);
    } else {
      console.log('❌ 無効なデータ形式です');
      return { success: false, error: '無効なデータ形式です' };
    }
    
    // データの前処理
    console.log('🔄 データの前処理開始');
    const outputData = [];
    for (let i = 0; i < dataArray.length; i++) {
      const data = dataArray[i];
      outputData.push([
        data.productName || '',      // B列: 商品名
        data.rightColumn || ''       // C列: 右隣列の値
      ]);
    }
    console.log(`✅ データ前処理完了: ${outputData.length}行`);
    
    // 既存データのクリア
    console.log('🧹 既存データのクリア開始');
    const lastRow = infoSheet.getLastRow();
    if (lastRow >= CONFIG.OUTPUT.START_ROW) {
      const clearRange = infoSheet.getRange(
        CONFIG.OUTPUT.START_ROW, 
        CONFIG.OUTPUT.COL_B, 
        lastRow - CONFIG.OUTPUT.START_ROW + 1, 
        2
      );
      clearRange.clear();
      console.log(`🗑️ クリア完了: ${CONFIG.OUTPUT.START_ROW}行目〜${lastRow}行目`);
    } else {
      console.log('ℹ️ クリア対象のデータなし');
    }
    
    // バッチ処理でデータを一括出力（高速化）
    console.log('📤 データ出力開始（バッチ処理）');
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
    
    console.log('🎉 情報抽出タブへの出力完了');
    
    return {
      success: true,
      outputRows: outputData.length,
      outputRange: `${CONFIG.OUTPUT.START_ROW}行目〜${CONFIG.OUTPUT.START_ROW + outputData.length - 1}行目`
    };
    
  } catch (error) {
    console.log(`❌ 情報抽出タブへの出力エラー: ${error.message}`);
    console.log(`🔍 エラー詳細: ${error.toString()}`);
    return {
      success: false,
      error: error.message,
      stack: error.stack
    };
  }
}

// getColumnLetter関数はUtils.gsで管理
