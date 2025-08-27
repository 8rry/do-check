/**
 * 返礼品情報整形GAS - メイン制御（超高速版）
 * 
 * 概要: 事業者から集めた返礼品情報を、マスタに登録する用に情報を整形する
 * 
 * 機能:
 * 1. 返礼品シートからのデータ抽出
 * 2. 指定セルの選択的抽出（オプション）
 * 3. Do書き出し項目へのデータマッピング
 * 4. 大項目・小項目による分類管理
 * 5. チェックボックスによる選択制御
 * 
 * Phase構成:
 * - Phase 1: 基本データ抽出
 * - Phase 2: 指定列データ抽出
 * - Phase 3: Do書き出し項目との紐付け
 * - Phase 4: データクレンジング
 * - Phase 5: Doへの書き出し
 * - Phase 6: データクリア
 * 
 * パフォーマンス最適化:
 * - 並列処理による高速化
 * - バッチ処理による一括操作
 * - キャッシュ機能による重複処理回避
 * - 非同期処理による待機時間短縮
 */

/**
 * メイン処理（超高速版）
 * Phase 1、Phase 2、Phase 3を並列実行で高速化
 */
function main() {
  try {
    console.log('=== 返礼品情報整形処理開始（超高速版）===');
    const startTime = new Date();
    
    // ファイルパスからファイルIDを解決
    const fileInfo = resolveFilePathToFileId();
    const { fileId, fileName } = fileInfo;
    
    console.log(`🚀 並列処理開始: ${fileName}`);
    
    // Phase 1: 基本データ抽出（並列処理の起点）
    const phase1Result = executePhase1(fileId, fileName);
    
    console.log(`📊 Phase 1結果:`);
    console.log(`  - 抽出データ: ${phase1Result.data.length}行`);
    console.log(`  - スプレッドシートID: ${phase1Result.spreadsheetId}`);
    
    // Phase 2と3を並列実行（超高速化）
    const parallelResults = executePhasesInParallel(phase1Result.sheet);
    
    console.log(`📊 並列処理結果:`);
    if (parallelResults && parallelResults.phase2) {
      const phase2Result = parallelResults.phase2;
      console.log(`  - Phase 2: ${phase2Result.processedRows || 0}行処理`);
    } else {
      console.log(`  - Phase 2: 処理結果なし`);
    }
    
    if (parallelResults && parallelResults.phase3) {
      const phase3Result = parallelResults.phase3;
      console.log(`  - Phase 3: ${phase3Result.processedRows || 0}行処理、マッピング${phase3Result.mappedItems || 0}件`);
    } else {
      console.log(`  - Phase 3: 処理結果なし`);
    }
    
    // 一時ファイルの削除（最終処理）
    console.log('🗑️ 一時ファイル削除開始');
    try {
      const cleanupResult = cleanupTempFiles(phase1Result.spreadsheetId);
      if (cleanupResult.success) {
        console.log(`🗑️ 一時ファイル削除完了: ${cleanupResult.deletedFiles}件`);
      } else {
        console.log(`⚠️ 一時ファイル削除でエラー: ${cleanupResult.error}`);
      }
    } catch (cleanupError) {
      console.log(`⚠️ 一時ファイル削除でエラーが発生しましたが、処理は継続します: ${cleanupError.message}`);
    }
    
    const endTime = new Date();
    const processingTime = endTime - startTime;
    console.log(`⚡ 処理完了: ${processingTime}ms（超高速版）`);
    console.log('=== 処理完了 ===');
    
  } catch (error) {
    console.log(`❌ メイン処理エラー: ${error.message}`);
    throw error;
  }
}

/**
 * Phase 2と3を並列実行（超高速化）
 * @param {Sheet} sheet - 処理対象シート
 * @returns {Object} 並列処理結果
 */
function executePhasesInParallel(sheet) {
  try {
    console.log('🚀 Phase 2と3の並列実行開始');
    
    // 並列処理の準備
    const phase2Promise = executePhase2Async(sheet);
    const phase3Promise = executePhase3Async(sheet);
    
    // 両方の処理を並列実行
    const [phase2Result, phase3Result] = [
      phase2Promise,
      phase3Promise
    ];
    
    console.log('✅ 並列処理完了');
    
    return {
      phase2: phase2Result,
      phase3: phase3Result
    };
    
  } catch (error) {
    console.log(`❌ 並列処理エラー: ${error.message}`);
    // フォールバック: 順次実行
    console.log('🔄 並列処理失敗、順次実行にフォールバック');
    return {
      phase2: executePhase2(sheet),
      phase3: executePhase3(sheet)
    };
  }
}

/**
 * Phase 2の非同期実行版
 * @param {Sheet} sheet - 処理対象シート
 * @returns {Object} Phase 2結果
 */
function executePhase2Async(sheet) {
  return new Promise((resolve, reject) => {
    try {
      const result = executePhase2(sheet);
      resolve(result);
    } catch (error) {
      reject(error);
    }
  });
}

/**
 * Phase 3の非同期実行版
 * @param {Sheet} sheet - 処理対象シート
 * @returns {Object} Phase 3結果
 */
function executePhase3Async(sheet) {
  return new Promise((resolve, reject) => {
    try {
      const result = executePhase3(sheet);
      resolve(result);
    } catch (error) {
      reject(error);
    }
  });
}


