/**
 * Phase 3: Do書き出し項目との紐付け（最適化版）
 * 
 * 概要: 情報抽出タブのB列・C列の値を部分検索でDoマスタの項目と紐付け
 * 
 * 機能:
 * 1. B列・C列の値からDo項目を自動判定
 * 2. 部分検索による柔軟なマッチング
 * 3. 優先順位を考慮した項目選択（旧・新の場合は新を優先）
 * 4. 情報抽出タブのA列へのバッチ出力
 * 5. ログ出力の最適化による高速化
 * 6. マッピング結果のキャッシュによる重複処理削減
 */

// Phase 3用のキャッシュ
let phase3MappingCache = {};
let phase3CacheTimestamp = null;

/**
 * Phase 3のメイン実行関数
 * @param {Sheet} sheet - Phase 1で処理されたシート
 * @returns {Object} 実行結果
 */
function executePhase3(sheet) {
  try {
    console.log('=== Phase 3: Do書き出し項目との紐付け開始 ===');
    const phase3StartTime = new Date();
    
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
    
    // 情報抽出タブのA列に出力
    const outputResult = outputToInfoExtractionTab(mappingResults);
    
    const phase3EndTime = new Date();
    const phase3ProcessingTime = phase3EndTime - phase3StartTime;
    console.log(`⚡ Phase 3処理時間: ${phase3ProcessingTime}ms`);
    console.log('=== Phase 3: Do書き出し項目との紐付け完了 ===');
    
    return {
      success: true,
      processedRows: extractedData.length,
      mappedItems: mappingResults.filter(r => r.mapped).length,
      outputResult: outputResult,
      processingTime: phase3ProcessingTime
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
    
    // 現在アクティブなスプレッドシートを取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      console.log('⚠️ アクティブなスプレッドシートが見つかりません');
      return [];
    }
    
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
      const actualRow = CONFIG.OUTPUT.START_ROW + i;  // 実際の行番号
      
      // 空行も含めて処理（行番号のずれを防ぐため）
      extractedData.push({
        row: actualRow,
        productName: productName,
        rightColumn: rightColumn,
        combinedText: `${productName} ${rightColumn}`.trim(),
        isEmpty: !productName && !rightColumn  // 空行フラグを追加
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
    
    // バッチ処理用の配列を準備
    const batchSize = CONFIG.PERFORMANCE.PHASE3_BATCH_SIZE || 50;
    let processedCount = 0;
    
    for (const data of extractedData) {
      // 空行の場合はスキップしてマッピング結果に追加
      if (data.isEmpty) {
        mappingResults.push({
          ...data,
          doItem: null,
          mapped: false
        });
        continue;
      }
      
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
      } else {
        mappingResults.push({
          ...data,
          doItem: null,
          mapped: false
        });
        unmappedCount++;
      }
      
      // バッチ処理の進捗をログ出力（詳細ログは無効化）
      processedCount++;
      if (processedCount % batchSize === 0) {
        console.log(`📊 マッピング進捗: ${processedCount}/${extractedData.length}件処理完了`);
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
 * 最適なDo項目を検索（キャッシュ付き）
 * @param {string} searchText - 検索テキスト
 * @returns {string|null} マッピング結果
 */
function findBestDoMapping(searchText) {
  try {
    if (!searchText || !CONFIG.DO_MAPPING) {
      return null;
    }
    
    // キャッシュをチェック
    const cacheKey = searchText.toLowerCase().trim();
    const now = new Date().getTime();
    const cacheTTL = CONFIG.PERFORMANCE.PHASE3_CACHE_TTL || 10 * 60 * 1000; // 10分
    
    if (phase3MappingCache[cacheKey] && phase3CacheTimestamp && 
        (now - phase3CacheTimestamp) < cacheTTL) {
      return phase3MappingCache[cacheKey];
    }
    
    let bestMatch = null;
    let bestScore = 0;
    let isOld = false;
    
    // 各マッピングルールをチェック
    for (const [doItem, mappingRule] of Object.entries(CONFIG.DO_MAPPING)) {
      // Utils.gsのマッチング関数を使用
      const matchResult = matchKeywordsWithOldNewCheck(searchText, mappingRule);
      
      if (matchResult.matched) {
        // スコア計算（簡略化）
        const totalScore = matchResult.isOld ? 0.5 : 1;
        
        if (totalScore > bestScore) {
          bestScore = totalScore;
          bestMatch = doItem;
          isOld = matchResult.isOld;
        }
      }
    }
    
    let result = null;
    if (bestMatch) {
      // 旧項目の場合はnullを返す（ラベルを付けない）
      if (!isOld) {
        result = bestMatch;
      }
    }
    
    // キャッシュに保存
    phase3MappingCache[cacheKey] = result;
    phase3CacheTimestamp = now;
    
    return result;
    
  } catch (error) {
    console.log(`❌ Do項目検索エラー: ${error.message}`);
    return null;
  }
}

/**
 * 情報抽出タブのA列にDo項目を出力
 * @param {Array} mappingResults - マッピング結果
 * @returns {Object} 出力結果
 */
function outputToInfoExtractionTab(mappingResults) {
  try {
    console.log('📤 情報抽出タブのA列にDo項目を出力開始');
    
    // 現在アクティブなスプレッドシートを取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      console.log('⚠️ アクティブなスプレッドシートが見つかりません');
      return {
        success: false,
        error: 'アクティブなスプレッドシートが見つかりません',
        outputRows: 0,
        outputRange: ''
      };
    }
    
    const sheet = ss.getSheetByName(CONFIG.SHEETS.INFO_EXTRACTION);
    
    if (!sheet) {
      throw new Error('情報抽出タブが見つかりません');
    }
    
    // 既存のA列データをクリア（8行目から）
    console.log('🧹 既存のA列データをクリア開始');
    const lastRow = sheet.getLastRow();
    if (lastRow >= CONFIG.OUTPUT.START_ROW) {
      const clearRange = sheet.getRange(
        CONFIG.OUTPUT.START_ROW, 
        1, // A列
        lastRow - CONFIG.OUTPUT.START_ROW + 1, 
        1  // A列のみ
      );
      clearRange.clear();
      console.log(`🗑️ A列クリア完了: ${CONFIG.OUTPUT.START_ROW}行目〜${lastRow}行目`);
    }
    
    // A列にDo項目を個別出力（正確な行番号で出力）
    console.log('📤 A列にDo項目を個別出力開始');
    if (mappingResults.length > 0) {
      let outputCount = 0;
      
      // 各行を正確な行番号に出力
      for (const result of mappingResults) {
        if (result.doItem) {
          sheet.getRange(result.row, 1).setValue(result.doItem);
          outputCount++;
        }
      }
      
      console.log(`✅ 個別出力完了: ${outputCount}件のDo項目を出力`);
    }
    
    console.log('🎉 情報抽出タブのA列への出力完了');
    
    return {
      success: true,
      outputRows: mappingResults.filter(r => r.doItem).length,
      outputRange: `A${CONFIG.OUTPUT.START_ROW}行目〜A${CONFIG.OUTPUT.START_ROW + mappingResults.length - 1}行目`
    };
    
  } catch (error) {
    console.log(`❌ 情報抽出タブのA列への出力エラー: ${error.message}`);
    console.log(`🔍 エラー詳細: ${error.toString()}`);
    return {
      success: false,
      error: error.message,
      stack: error.stack
    };
  }
}

/**
 * Phase 3のキャッシュをクリア
 */
function clearPhase3Cache() {
  phase3MappingCache = {};
  phase3CacheTimestamp = null;
  console.log('🗑️ Phase 3キャッシュをクリアしました');
}

/**
 * Phase 3のキャッシュ状態を取得
 * @returns {Object} キャッシュ状態
 */
function getPhase3CacheStatus() {
  const cacheSize = Object.keys(phase3MappingCache).length;
  const now = new Date().getTime();
  const cacheAge = phase3CacheTimestamp ? (now - phase3CacheTimestamp) / 1000 : 0;
  
  return {
    cacheSize: cacheSize,
    cacheAge: Math.round(cacheAge),
    isExpired: cacheAge > (CONFIG.PERFORMANCE.PHASE3_CACHE_TTL || 600000) / 1000
  };
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
