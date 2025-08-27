/**
 * Phase 3: Do書き出し項目との紐付け
 * 
 * 概要: 情報抽出タブのB列・C列の値を部分検索でDoマスタの項目と紐付け
 * 
 * 機能:
 * 1. B列・C列の値からDo項目を自動判定
 * 2. 部分検索による柔軟なマッチング
 * 3. 優先順位を考慮した項目選択（旧・新の場合は新を優先）
 * 4. 情報抽出タブのA列への出力
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
    
    // 情報抽出タブのA列に出力
    const outputResult = outputToInfoExtractionTab(mappingResults);
    
    console.log('=== Phase 3: Do書き出し項目との紐付け完了 ===');
    
    return {
      success: true,
      processedRows: extractedData.length,
      mappedItems: mappingResults.filter(r => r.mapped).length,
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
 * 最適なDo項目を検索
 * @param {string} searchText - 検索テキスト
 * @returns {string|null} マッピング結果
 */
function findBestDoMapping(searchText) {
  try {
    if (!searchText || !CONFIG.DO_MAPPING) {
      return null;
    }
    
    let bestMatch = null;
    let bestScore = 0;
    let isOld = false;
    
    // 各マッピングルールをチェック
    for (const [doItem, mappingRule] of Object.entries(CONFIG.DO_MAPPING)) {
      // メインキーワードでマッチング（AND検索）
      let hasMainKeywordMatch = false;
      let keywordScore = 0;
      
      if (mappingRule.keywords && mappingRule.keywords.length > 0) {
        // すべてのキーワードが含まれているかチェック（AND検索）
        const allKeywordsMatch = mappingRule.keywords.every(keyword => 
          searchText.toLowerCase().includes(keyword.toLowerCase())
        );
        
        if (allKeywordsMatch) {
          hasMainKeywordMatch = true;
          keywordScore = mappingRule.keywords.length;
        }
      }
      
      // fallbackKeywordsでマッチング（OR検索）
      let hasFallbackMatch = false;
      let fallbackScore = 0;
      
      if (mappingRule.fallbackKeywords) {
        for (const fallbackKeyword of mappingRule.fallbackKeywords) {
          if (searchText.toLowerCase().includes(fallbackKeyword.toLowerCase())) {
            hasFallbackMatch = true;
            fallbackScore += 1;
          }
        }
      }
      
      // マッチング判定
      if (hasMainKeywordMatch || hasFallbackMatch) {
        // 新・旧の判定
        let isNew = false;
        let isOldItem = false;
        
        if (mappingRule.fallbackKeywords) {
          for (const fallbackKeyword of mappingRule.fallbackKeywords) {
            if (searchText.toLowerCase().includes(fallbackKeyword.toLowerCase())) {
              if (fallbackKeyword.includes('(新)')) {
                isNew = true;
              } else if (fallbackKeyword.includes('(旧)')) {
                isOldItem = true;
              }
            }
          }
        }
        
        // 部分一致による新・旧判定も追加
        if (searchText.toLowerCase().includes('(新)') || 
            searchText.toLowerCase().includes('新') ||
            searchText.toLowerCase().includes('new')) {
          isNew = true;
        }
        if (searchText.toLowerCase().includes('(旧)') || 
            searchText.toLowerCase().includes('旧') ||
            searchText.toLowerCase().includes('old')) {
          isOldItem = true;
        }
        
        // 新・旧の判定結果を決定
        let finalIsOld = false;
        if (isOldItem && !isNew) {
          finalIsOld = true;
        } else if (isNew && !isOldItem) {
          finalIsOld = false;
        } else if (isNew && isOldItem) {
          // 新・旧両方の場合は設定された優先順位に従う
          finalIsOld = mappingRule.priority !== 'new';
        } else {
          // 新・旧の判定なしの場合は新項目として扱う
          finalIsOld = false;
        }
        
        // スコア計算（メインキーワードの方が高スコア）
        const totalScore = (keywordScore * 2) + fallbackScore;
        
        if (totalScore > bestScore) {
          bestScore = totalScore;
          bestMatch = doItem;
          isOld = finalIsOld;
        }
      }
    }
    
    if (bestMatch) {
      // 旧項目の場合はnullを返す（ラベルを付けない）
      if (isOld) {
        console.log(`🔍 Do項目マッチング: "${searchText}" → "${bestMatch}" (旧項目のためラベルを付けません)`);
        return null;
      }
      
      console.log(`🔍 Do項目マッチング: "${searchText}" → "${bestMatch}" (スコア: ${bestScore})`);
      return bestMatch;
    }
    
    return null;
    
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
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
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
    
    // A列にDo項目を出力
    console.log('📤 A列にDo項目を出力開始');
    if (mappingResults.length > 0) {
      let outputCount = 0;
      
      for (const result of mappingResults) {
        if (result.doItem) {
          const row = result.row;
          sheet.getRange(row, 1).setValue(result.doItem);
          outputCount++;
          
          if (CONFIG.PERFORMANCE.LOG_DETAIL) {
            console.log(`✅ A列にDo項目を出力: 行${row} "${result.doItem}"`);
          }
        }
      }
      
      console.log(`✅ 出力完了: ${outputCount}件のDo項目を出力`);
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
