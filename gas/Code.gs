/**
 * 返礼品情報整形GAS - メイン制御
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
 */

/**
 * メイン処理
 * Phase 1、Phase 2、Phase 3を順次実行
 */
function main() {
  try {
    console.log('=== 返礼品情報整形処理開始 ===');
    
    // ファイルパスからファイルIDを解決
    const fileInfo = resolveFilePathToFileId();
    const { fileId, fileName } = fileInfo;
    
    // Phase 1: 基本データ抽出
    const phase1Result = executePhase1(fileId, fileName);
    
    console.log(`📊 Phase 1結果:`);
    console.log(`  - 抽出データ: ${phase1Result.data.length}行`);
    console.log(`  - スプレッドシートID: ${phase1Result.spreadsheetId}`);
    
    // Phase 2: 指定列データ抽出
    const phase2Result = executePhase2(phase1Result.sheet);
    
    // Phase 3: Do書き出し項目との紐付け
    const phase3Result = executePhase3(phase1Result.sheet);
    
    console.log(`📊 Phase 3結果:`);
    console.log(`  - 処理行数: ${phase3Result.processedRows}行`);
    console.log(`  - マッピング成功: ${phase3Result.mappedItems}件`);
    
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
    
    console.log('=== 処理完了 ===');
    
  } catch (error) {
    console.log(`❌ メイン処理エラー: ${error.message}`);
    throw error;
  }
}


