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
 * Phase 1とPhase 2を順次実行
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
    
    console.log('=== 処理完了 ===');
    
  } catch (error) {
    console.log(`❌ メイン処理エラー: ${error.message}`);
    throw error;
  }
}


