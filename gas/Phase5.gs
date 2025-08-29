/**
 * 返礼品情報整形GAS - Phase 5: データクリア処理
 * 
 * 概要: 情報抽出タブ、Do書き出し用タブ、Do書き出し用(定期)タブの内容を一括でクリア
 * 
 * 処理内容:
 * 1. 情報抽出タブのクリア
 *    - D7:CQ7のチェック項目をfalseにする
 *    - B6を削除
 *    - A8:CQ200のデータを削除
 *    - D4:CQ6のデータを削除
 * 2. Do書き出し用タブのクリア
 *    - 1行目は残し、2行目以降を削除
 * 3. Do書き出し用(定期)タブのクリア
 *    - 1行目は残し、2行目以降を削除
 */

/**
 * 情報抽出タブの内容をクリア
 * @returns {boolean} 処理結果
 */
function clearInfoExtractionTab() {
  try {
    console.log('🗑️ 情報抽出タブのクリア開始');
    
    // 情報抽出タブを取得
    const sheet = getInfoExtractionSheet();
    if (!sheet) {
      console.log('❌ 情報抽出タブが見つかりません');
      return false;
    }
    
    // 1. D7:CQ7のチェック項目をfalseにする
    console.log('📋 チェックボックスをfalseに設定中...');
    const checkboxRange = sheet.getRange(7, 4, 1, 87); // D7:CQ7 (4列目から87列分)
    checkboxRange.setValue(false);
    console.log('✅ チェックボックスの設定完了');
    
    // 2. B6を削除
    console.log('🗑️ B6セルの内容を削除中...');
    sheet.getRange('B6').clearContent();
    console.log('✅ B6セルの削除完了');
    
    // 3. A8:CQ200のデータを削除
    console.log('🗑️ データ領域(A8:CQ200)を削除中...');
    const dataRange = sheet.getRange(8, 1, 193, 87); // A8:CQ200 (8行目から193行分、1列目から87列分)
    dataRange.clearContent();
    console.log('✅ データ領域の削除完了');
    
    // 4. D4:CQ6のデータを削除
    console.log('🗑️ ヘッダー領域(D4:CQ6)を削除中...');
    const headerRange = sheet.getRange(4, 4, 3, 87); // D4:CQ6 (4行目から3行分、4列目から87列分)
    headerRange.clearContent();
    console.log('✅ ヘッダー領域の削除完了');
    
    console.log('✅ 情報抽出タブのクリア完了');
    return true;
    
  } catch (error) {
    console.log(`❌ 情報抽出タブのクリア処理でエラー: ${error.message}`);
    return false;
  }
}

/**
 * Do書き出し用タブの内容をクリア
 * @returns {boolean} 処理結果
 */
function clearDoOutputTab() {
  try {
    console.log('🗑️ Do書き出し用タブのクリア開始');
    
    // Do書き出し用タブを取得
    const sheet = getDoOutputSheet();
    if (!sheet) {
      console.log('❌ Do書き出し用タブが見つかりません');
      return false;
    }
    
    // 1行目は残し、2行目以降を削除
    console.log('🗑️ 2行目以降のデータを削除中...');
    const lastRow = sheet.getLastRow();
    
    if (lastRow > 1) {
      const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
      dataRange.clearContent();
      console.log(`✅ 2行目から${lastRow}行目まで削除完了`);
    } else {
      console.log('ℹ️ 削除対象のデータがありません');
    }
    
    console.log('✅ Do書き出し用タブのクリア完了');
    return true;
    
  } catch (error) {
    console.log(`❌ Do書き出し用タブのクリア処理でエラー: ${error.message}`);
    return false;
  }
}

/**
 * Do書き出し用(定期)タブの内容をクリア
 * @returns {boolean} 処理結果
 */
function clearDoOutputSubscriptionTab() {
  try {
    console.log('🗑️ Do書き出し用(定期)タブのクリア開始');
    
    // Do書き出し用(定期)タブを取得
    const sheet = getDoOutputSubscriptionSheet();
    if (!sheet) {
      console.log('❌ Do書き出し用(定期)タブが見つかりません');
      return false;
    }
    
    // 1行目は残し、2行目以降を削除
    console.log('🗑️ 2行目以降のデータを削除中...');
    const lastRow = sheet.getLastRow();
    
    if (lastRow > 1) {
      const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
      dataRange.clearContent();
      console.log(`✅ 2行目から${lastRow}行目まで削除完了`);
    } else {
      console.log('ℹ️ 削除対象のデータがありません');
    }
    
    console.log('✅ Do書き出し用(定期)タブのクリア完了');
    return true;
    
  } catch (error) {
    console.log(`❌ Do書き出し用(定期)タブのクリア処理でエラー: ${error.message}`);
    return false;
  }
}

/**
 * 情報抽出タブを取得
 * @returns {Sheet|null} 情報抽出タブのシート、見つからない場合はnull
 */
function getInfoExtractionSheet() {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.INFO_EXTRACTION);
    return sheet;
  } catch (error) {
    console.log(`❌ 情報抽出タブの取得でエラー: ${error.message}`);
    return null;
  }
}

/**
 * Do書き出し用タブを取得
 * @returns {Sheet|null} Do書き出し用タブのシート、見つからない場合はnull
 */
function getDoOutputSheet() {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.DO_EXPORT);
    return sheet;
  } catch (error) {
    console.log(`❌ Do書き出し用タブの取得でエラー: ${error.message}`);
    return null;
  }
}

/**
 * Do書き出し用(定期)タブを取得
 * @returns {Sheet|null} Do書き出し用(定期)タブのシート、見つからない場合はnull
 */
function getDoOutputSubscriptionSheet() {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.DO_EXPORT_REGULAR);
    return sheet;
  } catch (error) {
    console.log(`❌ Do書き出し用(定期)タブの取得でエラー: ${error.message}`);
    return null;
  }
}
