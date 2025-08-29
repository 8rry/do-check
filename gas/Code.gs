/**
 * 返礼品情報整形GAS - メイン制御
 * 
 * 概要: 事業者から集めた返礼品情報を、マスタに登録する用に情報を整形する
 * 
 * Phase構成:
 * - Phase 1: 基本データ抽出
 * - Phase 2: 指定列データ抽出
 * - Phase 3: Do書き出し項目との紐付け
 * - Phase 4: Doへの書き出し
 * - Phase 5: データクリア
 */

/**
 * メイン処理
 * Phase 1、Phase 2、Phase 3を順序立てて実行
 */
function main() {
  try {
    console.log('=== 返礼品情報整形処理開始 ===');
    const startTime = new Date();
    
    // ファイルパスからファイルIDを解決
    const fileInfo = resolveFilePathToFileId();
    const { fileId, fileName } = fileInfo;
    
    console.log(`📁 処理開始: ${fileName}`);
    
    // Phase 1: 基本データ抽出
    const phase1Result = executePhase1(fileId, fileName);
    
    console.log(`✅ Phase 1完了: ${phase1Result.data.length}行抽出`);
    
    // Phase 2と3を順次実行
    const parallelResults = executePhasesInParallel(phase1Result.sheet);
    
    // 処理結果の表示
    if (parallelResults.phase2) {
      console.log(`✅ Phase 2完了: ${parallelResults.phase2.processedRows || 0}行処理`);
    }
    
    if (parallelResults.phase3) {
      console.log(`✅ Phase 3完了: ${parallelResults.phase3.processedRows || 0}行処理`);
    }
    
    // 一時ファイルの削除
    console.log('🗑️ 一時ファイル削除開始');
    try {
      const cleanupResult = cleanupTempFiles(phase1Result.spreadsheetId);
      if (cleanupResult.success) {
        console.log(`🗑️ 一時ファイル削除完了: ${cleanupResult.deletedFiles}件`);
      } else {
        console.log(`⚠️ 一時ファイル削除エラー: ${cleanupResult.error}`);
      }
    } catch (cleanupError) {
      console.log(`⚠️ 一時ファイル削除でエラーが発生しましたが、処理は継続します: ${cleanupError.message}`);
    }

    // スタイル処理
    SpreadsheetApp.getActiveSheet()
    .getDataRange()
    .setFontSize(9)
    .setFontFamily("Noto Sans JP")
    .setVerticalAlignment("middle");
    
    const endTime = new Date();
    const processingTime = endTime - startTime;
    console.log(`⚡ 処理完了: ${processingTime}ms`);
    console.log('=== 処理完了 ===');
    
  } catch (error) {
    console.log(`❌ メイン処理エラー: ${error.message}`);
    throw error;
  }
}

/**
 * Phase 2と3を順次実行
 * @param {Sheet} sheet - 処理対象シート
 * @returns {Object} 処理結果
 */
function executePhasesInParallel(sheet) {
  try {
    console.log('🔄 Phase 2と3の処理開始');
    
    // Phase 2を先に実行（Phase 3の処理に影響する可能性があるため）
    const phase2Result = executePhase2(sheet);
    const phase3Result = executePhase3(sheet);
    
    console.log('✅ Phase 2と3の処理完了');
    
    return {
      phase2: phase2Result,
      phase3: phase3Result
    };
    
  } catch (error) {
    console.log(`❌ Phase 2と3の処理エラー: ${error.message}`);
    // エラーが発生した場合は個別に実行
    let phase2Result = null;
    let phase3Result = null;
    
    try {
      phase2Result = executePhase2(sheet);
    } catch (phase2Error) {
      console.log(`⚠️ Phase 2個別実行エラー: ${phase2Error.message}`);
    }
    
    try {
      phase3Result = executePhase3(sheet);
    } catch (phase3Error) {
      console.log(`⚠️ Phase 3個別実行エラー: ${phase3Error.message}`);
    }
    
    return {
      phase2: phase2Result,
      phase3: phase3Result
    };
  }
}

/**
 * Phase 4: Doへの書き出し（独立実行用）
 * Phase 1-3とは独立して実行可能
 * 
 * 処理内容:
 * 1. チェックボックスを確認（チェック項目のみ処理する）
 * 2. 単一商品か、定期便かの見極め
 * 3. 項目名をキーにDo書き出し用に格納
 * 4. データをクレンジング
 */
function executePhase4Standalone() {
  try {
    console.log('=== Phase 4: Doへの書き出し開始 ===');
    const startTime = new Date();
    
    // Phase 4の実行
    const phase4Result = executePhase4();
    
    if (phase4Result) {
      const endTime = new Date();
      const processingTime = endTime - startTime;
      console.log(`✅ Phase 4完了: ${processingTime}ms`);
      console.log('=== Phase 4: Doへの書き出し完了 ===');
    } else {
      console.log('❌ Phase 4でエラーが発生しました');
    }
    
    return phase4Result;
    
  } catch (error) {
    console.log(`❌ Phase 4実行エラー: ${error.message}`);
    throw error;
  }
}

/**
 * Phase 4: Doへの書き出し（内部実行用）
 * @returns {boolean} 処理結果
 */
function executePhase4() {
  try {
    // 1. チェックボックスを確認（チェック項目のみ処理する）
    const checkedColumns = getCheckedColumns();
    if (checkedColumns.length === 0) {
      console.log('⚠️ チェックされている列がありません');
      return false;
    }
    console.log(`📋 チェック済み列数: ${checkedColumns.length}列`);
    
    // 2. 単一商品か、定期便かの見極め
    const productTypes = determineProductTypes(checkedColumns);
    const singleCount = Object.values(productTypes).filter(type => type === 'single').length;
    const subscriptionCount = Object.values(productTypes).filter(type => type === 'subscription').length;
    console.log(`🔍 商品種別判別完了: 単一商品${singleCount}件、定期便${subscriptionCount}件`);
    
    // 3. 項目名をキーにDo書き出し用に格納
    const extractedData = extractDataForDo(checkedColumns, productTypes);
    console.log(`📊 データ抽出完了: ${Object.keys(extractedData).length}列分のデータを抽出`);
    
    // 4. データをクレンジング
    const cleanedData = cleanData(extractedData);
    console.log(`🧹 データクレンジング完了: ${Object.keys(cleanedData).length}列分のデータをクレンジング`);
    
    // 5. Do書き出し用タブへの出力
    const outputResult = outputToDoTabs(cleanedData, productTypes);
    if (outputResult) {
      console.log(`📤 Do書き出し用タブへの出力完了`);
    }
    
    return true;
    
  } catch (error) {
    console.error('❌ Phase 4 エラー:', error);
    return false;
  }
}

/**
 * Phase 5: データクリア（独立実行用）
 * Phase 1-4とは独立して実行可能
 * 
 * 処理内容:
 * 1. 情報抽出タブの内容をクリア
 * 2. Do書き出し用タブの内容をクリア
 * 3. Do書き出し用(定期)タブの内容をクリア
 */
function executePhase5Standalone() {
  try {
    console.log('=== Phase 5: データクリア処理開始 ===');
    const startTime = new Date();
    
    // Phase 5の実行
    const phase5Result = executePhase5();
    
    if (phase5Result) {
      const endTime = new Date();
      const processingTime = endTime - startTime;
      console.log(`✅ Phase 5完了: ${processingTime}ms`);
      console.log('=== データクリア処理完了 ===');
    } else {
      console.log('❌ Phase 5でエラーが発生しました');
    }
    
    return phase5Result;
    
  } catch (error) {
    console.log(`❌ Phase 5実行エラー: ${error.message}`);
    throw error;
  }
}

/**
 * Phase 5: データクリア（内部実行用）
 * @returns {boolean} 処理結果
 */
function executePhase5() {
  try {
    // 1. 情報抽出タブのクリア
    const infoExtractionResult = clearInfoExtractionTab();
    if (infoExtractionResult) {
      console.log('✅ 情報抽出タブのクリア完了');
    } else {
      console.log('❌ 情報抽出タブのクリアでエラーが発生しました');
    }
    
    // 2. Do書き出し用タブのクリア
    const doOutputResult = clearDoOutputTab();
    if (doOutputResult) {
      console.log('✅ Do書き出し用タブのクリア完了');
    } else {
      console.log('❌ Do書き出し用タブのクリアでエラーが発生しました');
    }
    
    // 3. Do書き出し用(定期)タブのクリア
    const doOutputSubscriptionResult = clearDoOutputSubscriptionTab();
    if (doOutputSubscriptionResult) {
      console.log('✅ Do書き出し用(定期)タブのクリア完了');
    } else {
      console.log('❌ Do書き出し用(定期)タブのクリアでエラーが発生しました');
    }
    
    return infoExtractionResult && doOutputResult && doOutputSubscriptionResult;
    
  } catch (error) {
    console.error('❌ Phase 5 エラー:', error);
    return false;
  }
}


