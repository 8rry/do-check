/**
 * 返礼品情報整形GAS - メイン制御
 * 
 * 概要: 事業者から集めた返礼品情報を、マスタに登録する用に情報を整形する
 * 
 * Phase構成:
 * - Phase 1: 基本データ抽出
 * - Phase 2: 指定列データ抽出
 * - Phase 3: Do書き出し項目との紐付け
 * - Phase 4: データクレンジング
 * - Phase 5: Doへの書き出し
 * - Phase 6: データクリア
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


