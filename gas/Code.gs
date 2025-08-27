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
    
    // ファイルパスを読み込み
    const folderPath = loadFolderPath();
    if (!folderPath) {
      throw new Error('ファイルパスが設定されていません');
    }
    
    // ファイルパスから自治体フォルダキーとサブパスを抽出
    const pathInfo = convertWindowsPathToDrivePath(folderPath);
    if (!pathInfo) {
      throw new Error('パス情報の抽出に失敗しました');
    }
    
    console.log(`📋 抽出されたパス情報:`);
    console.log(`  - 自治体フォルダキー: ${pathInfo.folderKey}`);
    console.log(`  - サブパス: ${pathInfo.subPath}`);
    
    // 自治体フォルダタブからフォルダIDを取得
    const folderId = findMunicipalityFolder(pathInfo.folderKey);
    if (!folderId) {
      throw new Error(`自治体フォルダ "${pathInfo.folderKey}" が見つかりません`);
    }
    
    // ファイル名を抽出
    const fileName = extractFileNameFromPath(folderPath);
    
    // サブパスを使用してファイルを検索
    const fileId = findFileInFolderWithSubPath(folderId, pathInfo.subPath, fileName);
    if (!fileId) {
      throw new Error(`ファイル "${fileName}" がパス "${pathInfo.subPath}" 内に見つかりません`);
    }
    
    console.log(`✅ ファイルID取得完了: ${fileId}`);
    
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

/**
 * 全体テスト実行
 */
function test() {
  try {
    console.log('=== 全体テスト開始 ===');
    
    // 設定値の確認
    console.log('設定:', JSON.stringify(CONFIG, null, 2));
    
    // スプレッドシートへの接続テスト
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    console.log('スプレッドシート名:', spreadsheet.getName());
    
    // シートの存在確認
    const sheets = spreadsheet.getSheets();
    console.log('利用可能なシート:', sheets.map(s => s.getName()));
    
    console.log('=== 全体テスト完了 ===');
    
  } catch (error) {
    console.error('全体テストエラー:', error.message);
  }
}
