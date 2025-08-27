/**
 * Phase 3: Do書き出し項目との紐付け
 * 
 * 概要: 情報抽出タブのB列・C列の値を部分検索でDoマスタの項目と紐付け
 * 
 * 機能:
 * 1. B列・C列の値からDo項目を自動判定
 * 2. 部分検索による柔軟なマッチング
 * 3. 優先順位を考慮した項目選択（旧・新の場合は新を優先）
 * 4. Do書き出し用タブへの出力
 */

/**
 * Phase 3のメイン実行関数
 * @param {Sheet} sheet - Phase 1で処理されたシート
 * @returns {Object} 実行結果
 */
function executePhase3(sheet) {
  try {
    console.log('=== Phase 3: Do書き出し項目との紐付け開始 ===');
    
    if (!sheet) {
      throw new Error('シートが指定されていません');
    }
    
    // 情報抽出タブからB列・C列のデータを取得
    const extractedData = extractInfoExtractionData();
    if (!extractedData || extractedData.length === 0) {
      console.log('⚠️ 情報抽出タブにデータがありません');
      return { success: false, error: 'データがありません' };
    }
    
    console.log(`📊 処理対象データ: ${extractedData.length}行`);
    
    // Do項目との紐付けを実行
    const mappingResults = performDoMapping(extractedData);
    
    // Do書き出し用タブに出力
    const outputResult = outputToDoExportTab(mappingResults);
    
    console.log('=== Phase 3: Do書き出し項目との紐付け完了 ===');
    
    return {
      success: true,
      processedRows: extractedData.length,
      mappedItems: mappingResults.length,
      outputResult: outputResult
    };
    
  } catch (error) {
    console.log(`❌ Phase 3エラー: ${error.message}`);
    console.log(`🔍 エラー詳細: ${error.toString()}`);
    return {
      success: false,
      error: error.message,
      stack: error.stack
    };
  }
}

/**
 * 情報抽出タブからB列・C列のデータを取得
 * @returns {Array} 抽出されたデータ配列
 */
function extractInfoExtractionData() {
  try {
    console.log('🔍 情報抽出タブからデータ取得開始');
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.INFO_EXTRACTION);
    
    if (!sheet) {
      throw new Error('情報抽出タブが見つかりません');
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < CONFIG.OUTPUT.START_ROW) {
      console.log('⚠️ データ行が不足しています');
      return [];
    }
    
    // B列・C列のデータを取得（8行目から）
    const dataRange = sheet.getRange(
      CONFIG.OUTPUT.START_ROW, 
      CONFIG.OUTPUT.COL_B, 
      lastRow - CONFIG.OUTPUT.START_ROW + 1, 
      2
    );
    const dataValues = dataRange.getValues();
    
    const extractedData = [];
    for (let i = 0; i < dataValues.length; i++) {
      const rowData = dataValues[i];
      const productName = rowData[0] || '';  // B列
      const rightColumn = rowData[1] || '';  // C列
      
      // 空行はスキップ
      if (!productName && !rightColumn) {
        continue;
      }
      
      extractedData.push({
        row: CONFIG.OUTPUT.START_ROW + i,
        productName: productName,
        rightColumn: rightColumn,
        combinedText: `${productName} ${rightColumn}`.trim()
      });
    }
    
    console.log(`✅ データ取得完了: ${extractedData.length}行`);
    return extractedData;
    
  } catch (error) {
    console.log(`❌ データ取得エラー: ${error.message}`);
    throw error;
  }
}

/**
 * Do項目との紐付けを実行
 * @param {Array} extractedData - 抽出されたデータ
 * @returns {Array} マッピング結果
 */
function performDoMapping(extractedData) {
  try {
    console.log('🔗 Do項目との紐付け開始');
    
    const mappingResults = [];
    let mappedCount = 0;
    let unmappedCount = 0;
    
    for (const data of extractedData) {
      // 商品名と右隣列の値を組み合わせて検索
      const searchText = data.combinedText;
      
      // 最適なDo項目を検索
      const doItem = findBestDoMapping(searchText);
      
      if (doItem) {
        mappingResults.push({
          ...data,
          doItem: doItem,
          mapped: true
        });
        mappedCount++;
        
        if (CONFIG.PERFORMANCE.LOG_DETAIL) {
          console.log(`✅ マッピング成功: 行${data.row} "${searchText}" → "${doItem}"`);
        }
      } else {
        mappingResults.push({
          ...data,
          doItem: null,
          mapped: false
        });
        unmappedCount++;
        
        if (CONFIG.PERFORMANCE.LOG_DETAIL) {
          console.log(`❌ マッピング失敗: 行${data.row} "${searchText}"`);
        }
      }
    }
    
    console.log(`📊 マッピング結果:`);
    console.log(`  - 成功: ${mappedCount}件`);
    console.log(`  - 失敗: ${unmappedCount}件`);
    console.log(`  - 成功率: ${((mappedCount / extractedData.length) * 100).toFixed(1)}%`);
    
    return mappingResults;
    
  } catch (error) {
    console.log(`❌ マッピング処理エラー: ${error.message}`);
    throw error;
  }
}

/**
 * Do書き出し用タブに出力
 * @param {Array} mappingResults - マッピング結果
 * @returns {Object} 出力結果
 */
function outputToDoExportTab(mappingResults) {
  try {
    console.log('📤 Do書き出し用タブへの出力開始');
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.DO_EXPORT);
    
    if (!sheet) {
      throw new Error('Do書き出し用タブが見つかりません');
    }
    
    // 既存データのクリア
    console.log('🧹 既存データのクリア開始');
    const lastRow = sheet.getLastRow();
    if (lastRow >= CONFIG.CELLS.MAPPING_START_ROW) {
      const clearRange = sheet.getRange(
        CONFIG.CELLS.MAPPING_START_ROW, 
        1, // A列から
        lastRow - CONFIG.CELLS.MAPPING_START_ROW + 1, 
        3  // A列〜C列
      );
      clearRange.clear();
      console.log(`🗑️ クリア完了: ${CONFIG.CELLS.MAPPING_START_ROW}行目〜${lastRow}行目`);
    }
    
    // ヘッダー行の設定
    console.log('📝 ヘッダー行の設定');
    const headerRange = sheet.getRange(CONFIG.CELLS.MAPPING_START_ROW, 1, 1, 3);
    headerRange.setValues([['Do項目', '商品名', '右隣列の値']]);
    
    // データの出力
    console.log('📤 データ出力開始');
    if (mappingResults.length > 0) {
      const outputData = mappingResults.map(result => [
        result.doItem || '',           // A列: Do項目
        result.productName || '',      // B列: 商品名
        result.rightColumn || ''       // C列: 右隣列の値
      ]);
      
      const outputRange = sheet.getRange(
        CONFIG.CELLS.MAPPING_START_ROW + 1, 
        1, 
        outputData.length, 
        3
      );
      outputRange.setValues(outputData);
      
      console.log(`✅ データ出力完了: ${CONFIG.CELLS.MAPPING_START_ROW + 1}行目〜${CONFIG.CELLS.MAPPING_START_ROW + outputData.length}行目`);
    }
    
    console.log('🎉 Do書き出し用タブへの出力完了');
    
    return {
      success: true,
      outputRows: mappingResults.length,
      outputRange: `${CONFIG.CELLS.MAPPING_START_ROW + 1}行目〜${CONFIG.CELLS.MAPPING_START_ROW + mappingResults.length}行目`
    };
    
  } catch (error) {
    console.log(`❌ Do書き出し用タブへの出力エラー: ${error.message}`);
    console.log(`🔍 エラー詳細: ${error.toString()}`);
    return {
      success: false,
      error: error.message,
      stack: error.stack
    };
  }
}

/**
 * マッピング結果の統計情報を取得
 * @param {Array} mappingResults - マッピング結果
 * @returns {Object} 統計情報
 */
function getMappingStatistics(mappingResults) {
  try {
    if (!mappingResults || mappingResults.length === 0) {
      return {
        total: 0,
        mapped: 0,
        unmapped: 0,
        successRate: 0,
        topDoItems: []
      };
    }
    
    const total = mappingResults.length;
    const mapped = mappingResults.filter(r => r.mapped).length;
    const unmapped = total - mapped;
    const successRate = (mapped / total) * 100;
    
    // 最も多くマッピングされたDo項目を集計
    const doItemCounts = {};
    mappingResults.forEach(result => {
      if (result.doItem) {
        doItemCounts[result.doItem] = (doItemCounts[result.doItem] || 0) + 1;
      }
    });
    
    const topDoItems = Object.entries(doItemCounts)
      .sort(([,a], [,b]) => b - a)
      .slice(0, 5)
      .map(([item, count]) => ({ item, count }));
    
    return {
      total,
      mapped,
      unmapped,
      successRate: Math.round(successRate * 10) / 10,
      topDoItems
    };
    
  } catch (error) {
    console.log(`❌ 統計情報取得エラー: ${error.message}`);
    return null;
  }
}
