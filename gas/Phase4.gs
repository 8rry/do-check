/**
 * Phase 4: Doへの書き出し
 * 
 * 概要: Phase 1-3で抽出・紐付けされたデータを、Do書き出し用タブに整形して出力
 * 
 * 主要機能:
 * - チェックボックス確認による選択的処理
 * - 単一商品・定期便の自動判別
 * - データクレンジング処理
 * - Do書き出し用タブへの出力
 * - 一時ファイルの自動クリーンアップ
 */

// 一時ファイル管理用グローバル変数
let tempFileIds = [];

/**
 * 一時ファイルIDを登録
 * @param {string} tempFileId - 一時ファイルID
 */
function registerTempFile(tempFileId) {
  if (tempFileId && !tempFileIds.includes(tempFileId)) {
    tempFileIds.push(tempFileId);
    console.log(`📝 一時ファイル登録: ${tempFileId}`);
  }
}

/**
 * 一時ファイルをクリーンアップ
 * @param {boolean} forceCleanup - 強制クリーンアップフラグ
 */
function cleanupPhase4TempFiles(forceCleanup = false) {
  try {
    if (tempFileIds.length === 0) {
      console.log('ℹ️ クリーンアップ対象の一時ファイルがありません');
      return;
    }
    
    console.log(`🗑️ Phase4一時ファイルクリーンアップ開始: ${tempFileIds.length}件`);
    
    let deletedCount = 0;
    let errorCount = 0;
    
    tempFileIds.forEach(tempFileId => {
      try {
        const tempFile = DriveApp.getFileById(tempFileId);
        if (tempFile) {
          tempFile.setTrashed(true);
          deletedCount++;
          console.log(`✅ 一時ファイル削除完了: ${tempFileId}`);
        } else {
          console.log(`⚠️ 一時ファイルが見つかりません: ${tempFileId}`);
        }
      } catch (error) {
        errorCount++;
        console.error(`❌ 一時ファイル削除エラー: ${tempFileId} - ${error.message}`);
      }
    });
    
    // クリーンアップ完了後は配列をクリア
    tempFileIds = [];
    
    if (errorCount === 0) {
      console.log(`🗑️ Phase4一時ファイルクリーンアップ完了: ${deletedCount}件削除`);
    } else {
      console.log(`⚠️ Phase4一時ファイルクリーンアップ完了（一部エラー）: ${deletedCount}件削除、${errorCount}件エラー`);
    }
    
  } catch (error) {
    console.error('❌ Phase4一時ファイルクリーンアップ処理エラー:', error);
  }
}

/**
 * 情報抽出タブを検索・取得
 * @returns {Sheet} 情報抽出タブのシートオブジェクト
 */
function getInfoExtractionSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // デバッグ用: 利用可能なタブ名を確認
  const sheets = ss.getSheets();
  console.log(`🔍 利用可能なタブ数: ${sheets.length}`);
  sheets.forEach((sheet, index) => {
    console.log(`  - タブ${index + 1}: "${sheet.getName()}"`);
  });
  
  // 情報抽出タブを探す（複数の候補名を試行）
  let infoSheet = null;
  const possibleNames = ['情報抽出タブ', '情報抽出', 'Info Extraction', 'info_extraction'];
  
  for (const name of possibleNames) {
    infoSheet = ss.getSheetByName(name);
    if (infoSheet) {
      console.log(`✅ タブ発見: "${name}"`);
      break;
    }
  }
  
  if (!infoSheet) {
    // タブが見つからない場合は、最初のタブを使用
    infoSheet = ss.getSheets()[0];
    console.log(`⚠️ 情報抽出タブが見つからないため、最初のタブ "${infoSheet.getName()}" を使用します`);
  }
  
  return infoSheet;
}

/**
 * チェックボックスを確認し、チェックされている列を取得
 * @returns {Array} チェックされている列のインデックス配列
 */
function getCheckedColumns() {
  try {
    const infoSheet = getInfoExtractionSheet();
    
    const checkedColumns = [];
    
    // D列以降の7行目をチェック
    for (let col = 4; col <= infoSheet.getLastColumn(); col++) {
      const cell = infoSheet.getRange(7, col);
      if (cell.getValue() === true) { // チェックボックスがチェックされている
        checkedColumns.push(col);
      }
    }
    
    console.log(`🔍 チェックボックス確認完了: ${checkedColumns.length}列がチェックされています`);
    return checkedColumns;
    
  } catch (error) {
    console.error('❌ チェックボックス確認エラー:', error);
    throw error;
  }
}

/**
 * 商品種別を判別（単一商品か定期便か）
 * @param {Array} checkedColumns - チェックされている列のインデックス配列
 * @returns {Object} 列インデックスをキーとした商品種別マップ
 */
function determineProductTypes(checkedColumns) {
  try {
    const infoSheet = getInfoExtractionSheet();
    
    const productTypes = {};
    
    checkedColumns.forEach(col => {
      // 各列の「商品名称」項目の値を確認
      let productName = null;
      
      // A列の項目名を順次チェックして「商品名称」を探す
      for (let row = 8; row <= 200; row++) {
        const itemName = infoSheet.getRange(row, 1).getValue(); // A列の項目名
        if (itemName === '商品名称') {
          // 商品名称項目が見つかったら、その列の値を取得
          productName = infoSheet.getRange(row, col).getValue();
          console.log(`🔍 列${col}の商品名称: "${productName}"`);
          break;
        }
      }
      
      // 「定期」文字が含まれているかチェック
      if (productName && productName.toString().includes('定期')) {
        productTypes[col] = 'subscription'; // 定期便
        console.log(`✅ 列${col}: 定期便として判定 "${productName}"`);
      } else {
        productTypes[col] = 'single'; // 単一商品
        console.log(`ℹ️ 列${col}: 単一商品として判定 "${productName}"`);
      }
    });
    
    console.log(`🔍 商品種別判別完了: ${Object.keys(productTypes).length}列の判別が完了`);
    return productTypes;
    
  } catch (error) {
    console.error('❌ 商品種別判別エラー:', error);
    throw error;
  }
}

/**
 * 項目名をキーとしてデータを抽出
 * @param {Array} checkedColumns - チェックされている列のインデックス配列
 * @param {Object} productTypes - 商品種別マップ
 * @returns {Object} 抽出されたデータ
 */
function extractDataForDo(checkedColumns, productTypes) {
  try {
    const infoSheet = getInfoExtractionSheet();
    
    // 実際のデータ範囲を取得（最適化）
    const lastDataRow = infoSheet.getLastRow();
    const actualEndRow = Math.min(lastDataRow, 200);
    const rowCount = actualEndRow - 7; // 8行目から開始
    
    console.log(`🔍 データ抽出最適化: 8行目から${actualEndRow}行目まで (${rowCount}行)`);
    
    // 一括でデータを取得（高速化）
    const dataRange = infoSheet.getRange(8, 1, rowCount, Math.max(...checkedColumns));
    const allData = dataRange.getValues();
    
    const extractedData = {};
    
    checkedColumns.forEach(col => {
      const productType = productTypes[col];
      const columnData = {};
      
      // 一括取得したデータから抽出
      for (let i = 0; i < allData.length; i++) {
        const itemName = allData[i][0]; // A列の項目名
        if (itemName && itemName.toString().trim() !== '') {
          const dataValue = allData[i][col - 1]; // 列インデックスは0ベース
          columnData[itemName] = dataValue;
        }
      }
      
      extractedData[col] = {
        type: productType,
        data: columnData
      };
    });
    
    console.log(`📊 データ抽出完了: ${Object.keys(extractedData).length}列分のデータを抽出 (最適化版)`);
    return extractedData;
    
  } catch (error) {
    console.error('❌ データ抽出エラー:', error);
    throw error;
  }
}

/**
 * データをクレンジング
 * @param {Object} extractedData - 抽出されたデータ
 * @returns {Object} クレンジングされたデータ
 */
function cleanData(extractedData) {
  try {
    const cleanedData = {};
    
    Object.keys(extractedData).forEach(col => {
      const columnData = extractedData[col];
      const cleanedColumnData = {};
      
      Object.keys(columnData.data).forEach(itemName => {
        let value = columnData.data[itemName];
        
        console.log(`🔍 データクレンジング処理中: 項目名"${itemName}", 値"${value}"`);
        
        // 数字処理
        if (['寄附金額1', '提供価格(税込)1', '固定送料1'].includes(itemName)) {
          value = extractNumericValue(value);
        }
        
        // 発送種別変換
        if (itemName === '発送種別') {
          value = convertShippingType(value);
        }
        
        // 日付フォーマット統一
        if (itemName.includes('期間') || itemName.includes('日付')) {
          value = normalizeDateFormat(value);
        }
        
        // 通年扱い処理
        if (itemName.includes('期間')) {
          value = processYearRoundHandling(value);
        }
        
        // 文字変換処理
        if (itemName === '配送伝票商品名称') {
          value = convertFullWidthParentheses(value);
        }
        
        // 商品コードを配達会社用商品コードにコピー
        if (itemName === '商品コード') {
          const productCode = value;
          // 配達会社用商品コードに同じ値を設定
          cleanedColumnData['配達会社用商品コード'] = productCode;
        }
        
        cleanedColumnData[itemName] = value;
      });
      
      cleanedData[col] = {
        type: columnData.type,
        data: cleanedColumnData
      };
    });
    
    console.log(`🧹 データクレンジング完了: ${Object.keys(cleanedData).length}列分のデータをクレンジング`);
    return cleanedData;
    
  } catch (error) {
    console.error('❌ データクレンジングエラー:', error);
    throw error;
  }
}

/**
 * 数字値を抽出
 * @param {string} text - テキスト
 * @returns {string} 抽出された数字値
 */
function extractNumericValue(text) {
  if (!text) return '';
  
  const numericMatch = text.toString().match(/[\d,]+/);
  if (numericMatch) {
    return numericMatch[0].replace(/,/g, '');
  }
  
  return '';
}

/**
 * 発送種別を変換
 * @param {string} shippingType - 発送種別
 * @returns {string} 変換後の発送種別
 */
function convertShippingType(shippingType) {
  if (!shippingType) return '';
  
  const type = shippingType.toString();
  if (type.includes('常温')) return '通常便';
  if (type.includes('冷蔵')) return '冷蔵便';
  if (type.includes('冷凍')) return '冷凍便';
  
  return shippingType;
}

/**
 * 日付フォーマットを統一
 * @param {string} dateText - 日付テキスト
 * @returns {string} 統一された日付フォーマット
 */
function normalizeDateFormat(dateText) {
  if (!dateText) return '';
  
  // JavaScriptのDateオブジェクトの場合
  if (dateText instanceof Date) {
    const year = dateText.getFullYear();
    const month = String(dateText.getMonth() + 1).padStart(2, '0');
    const day = String(dateText.getDate()).padStart(2, '0');
    console.log(`🔍 Dateオブジェクト変換: "${dateText}" → "${year}/${month}/${day}"`);
    return `${year}/${month}/${day}`;
  }
  
  const text = dateText.toString();
  
  // JavaScriptのDate文字列形式（Wed Mar 18 2026 16:00:00 GMT+0900 (日本標準時)など）
  if (text.includes('GMT') || text.includes('UTC') || text.includes('GMT+') || text.includes('GMT-')) {
    try {
      const date = new Date(text);
      if (!isNaN(date.getTime())) {
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        console.log(`🔍 Date文字列変換: "${text}" → "${year}/${month}/${day}"`);
        return `${year}/${month}/${day}`;
      }
    } catch (error) {
      console.log(`⚠️ Date文字列変換エラー: "${text}"`, error);
    }
  }
  
  // yyyy年mm月dd日 → yyyy/mm/dd
  const japaneseMatch = text.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
  if (japaneseMatch) {
    const year = japaneseMatch[1];
    const month = japaneseMatch[2].padStart(2, '0');
    const day = japaneseMatch[3].padStart(2, '0');
    return `${year}/${month}/${day}`;
  }
  
  // yyyy-mm-dd → yyyy/mm/dd
  const dashMatch = text.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (dashMatch) {
    const year = dashMatch[1];
    const month = dashMatch[2].padStart(2, '0');
    const day = dashMatch[3].padStart(2, '0');
    return `${year}/${month}/${day}`;
  }
  
  // yyyy/mm/dd → そのまま返す（既に正しい形式）
  const slashMatch = text.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
  if (slashMatch) {
    return text; // 既に正しい形式なのでそのまま返す
  }
  
  // 上旬/下旬 → 具体的日付
  const earlyMatch = text.match(/(\d{4})年(\d{1,2})月上旬/);
  if (earlyMatch) {
    const year = earlyMatch[1];
    const month = earlyMatch[2].padStart(2, '0');
    return `${year}/${month}/15`;
  }
  
  const lateMatch = text.match(/(\d{4})年(\d{1,2})月下旬/);
  if (lateMatch) {
    const year = lateMatch[1];
    const month = lateMatch[2].padStart(2, '0');
    const lastDay = getLastDayOfMonth(parseInt(year), parseInt(month));
    return `${year}/${month}/${lastDay}`;
  }
  
  // その他の形式の場合はログ出力して元の値を返す
  console.log(`⚠️ 未対応の日付形式: "${text}" (型: ${typeof dateText})`);
  return dateText;
}

/**
 * 通年扱い処理
 * @param {string} periodText - 期間テキスト
 * @returns {string} 処理後の期間テキスト
 */
function processYearRoundHandling(periodText) {
  if (!periodText) return '';
  
  const text = periodText.toString();
  
  // 通年扱いキーワードチェック
  if (text.includes('通年') || text.includes('順次') || text.includes('随時')) {
    return '通年扱い';
  }
  
  return periodText;
}

/**
 * 全角カッコを半角に変換
 * @param {string} text - テキスト
 * @returns {string} 変換後のテキスト
 */
function convertFullWidthParentheses(text) {
  if (!text) return '';
  
  return text.toString()
    .replace(/（/g, '(')
    .replace(/）/g, ')');
}

/**
 * 指定月の末日を取得
 * @param {number} year - 年
 * @param {number} month - 月
 * @returns {number} 末日
 */
function getLastDayOfMonth(year, month) {
  return new Date(year, month, 0).getDate();
}

/**
 * 外部シートから寄附金額(開始)1の値を取得
 * @param {string} keyValue - 情報抽出タブB1の値
 * @returns {string} 寄附金額(開始)1の値
 */
function getExternalPriceValue(keyValue) {
  try {
    if (!keyValue) return '';
    
    // 外部シートのID
    const externalSheetId = '1aRAvMW8-VEVmZQbAHiIas53Jcq6QVR8E0bE6tgTiL3s';
    const sheetName = '商品マスタ登録依頼表(CS) 2025/05/01';
    
    // 外部シートを開く
    const externalSheet = SpreadsheetApp.openById(externalSheetId);
    const targetSheet = externalSheet.getSheetByName(sheetName);
    
    if (!targetSheet) {
      console.log(`⚠️ 外部シート "${sheetName}" が見つかりません`);
      return '';
    }
    
    // H列（キー列）とD列（値列）を取得
    const lastRow = targetSheet.getLastRow();
    const keyColumn = targetSheet.getRange(1, 8, lastRow, 1).getValues(); // H列
    const valueColumn = targetSheet.getRange(1, 4, lastRow, 1).getValues(); // D列
    
    // 部分一致で検索
    for (let i = 0; i < keyColumn.length; i++) {
      const key = keyColumn[i][0];
      if (key && key.toString().includes(keyValue)) {
        const value = valueColumn[i][0];
        console.log(`🔍 外部シート参照: キー"${keyValue}" → 値"${value}"`);
        return value;
      }
    }
    
    console.log(`⚠️ 外部シートでキー"${keyValue}"に部分一致する行が見つかりません`);
    return '';
    
  } catch (error) {
    console.error('❌ 外部シート参照エラー:', error);
    return '';
  }
}

/**
 * 定期便特別処理
 * @param {Object} cleanedData - クレンジングされたデータ
 * @returns {Object} 定期便処理後のデータ
 */
function processSubscriptionProducts(cleanedData) {
  try {
    const subscriptionData = {};
    
    Object.keys(cleanedData).forEach(col => {
      if (cleanedData[col].type === 'subscription') {
        const data = cleanedData[col].data;
        
        // 定期便の種類と回数/月数を判定
        const subscriptionType = determineSubscriptionType(data['商品名称']);
        let subscriptionCount;
        
        if (subscriptionType.includes('ヶ月定期便')) {
          // ヶ月定期便の場合
          subscriptionCount = determineSubscriptionMonths(data['商品名称']);
          console.log(`🔍 ヶ月定期便判定: "${data['商品名称']}" → ${subscriptionCount}ヶ月`);
        } else {
          // 回定期便の場合
          subscriptionCount = determineSubscriptionCount(data['商品名称']);
          console.log(`🔍 回定期便判定: "${data['商品名称']}" → ${subscriptionCount}回`);
        }
        
        // 子マスタ生成（1回目→2回目→3回目の順序）
        for (let i = 1; i <= subscriptionCount; i++) {
          const childData = generateChildMaster(data, i);
          subscriptionData[`${col}_child_${i}`] = {
            type: 'subscription_child',
            data: childData
          };
        }
        
        // 親マスタ生成
        const parentData = generateParentMaster(data, subscriptionCount);
        subscriptionData[`${col}_parent`] = {
          type: 'subscription_parent',
          data: parentData
        };
      }
    });
    
    console.log(`📦 定期便特別処理完了: ${Object.keys(subscriptionData).length}件のマスタを生成`);
    return subscriptionData;
    
  } catch (error) {
    console.error('❌ 定期便特別処理エラー:', error);
    throw error;
  }
}

/**
 * 定期便回数を判定
 * @param {string} productName - 商品名称
 * @returns {number} 定期便回数
 */
function determineSubscriptionCount(productName) {
  if (!productName) return 1;
  
  const text = productName.toString();
  
  // 3回定期便、6回定期便、12回定期便などのパターンを検出
  // より柔軟なパターンマッチング（全角・半角対応）
  const patterns = [
    /(\d+)回定期便/,  // 3回定期便
    /(\d+)回/,        // 3回
    /(\d+)回目/       // 3回目
  ];
  
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) {
      const count = parseInt(match[1]);
      console.log(`🔍 回定期便判定: "${text}" → ${count}回`);
      return count;
    }
  }
  
  // デフォルトは1回
  console.log(`🔍 回定期便判定: "${text}" → デフォルト1回`);
  return 1;
}

/**
 * ヶ月定期便の月数を判定
 * @param {string} productName - 商品名称
 * @returns {number} ヶ月定期便の月数
 */
function determineSubscriptionMonths(productName) {
  if (!productName) return 1;
  
  const text = productName.toString();
  
  // 3ヶ月定期便、6ヶ月定期便、12ヶ月定期便などのパターンを検出
  // より柔軟なパターンマッチング（全角・半角対応）
  const patterns = [
    /(\d+)ヶ月定期便/,  // 3ヶ月定期便
    /(\d+)か月定期便/,  // 3か月定期便
    /(\d+)月定期便/,    // 3月定期便
    /(\d+)ヶ月/,        // 3ヶ月
    /(\d+)か月/,        // 3か月
    /(\d+)月/           // 3月
  ];
  
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) {
      const months = parseInt(match[1]);
      console.log(`🔍 ヶ月定期便判定: "${text}" → ${months}ヶ月`);
      return months;
    }
  }
  
  // デフォルトは1ヶ月
  console.log(`🔍 ヶ月定期便判定: "${text}" → デフォルト1ヶ月`);
  return 1;
}

/**
 * 子マスタを生成
 * @param {Object} data - 元データ
 * @param {number} count - 回数
 * @returns {Object} 子マスタデータ
 */
function generateChildMaster(data, count) {
  const childData = { ...data };
  
  // 商品コード変換: 元コード + "-" + 回数
  if (data['商品コード']) {
    childData['商品コード'] = `${data['商品コード']}-${count}`;
    // 配達会社用商品コードも同じ値に設定
    childData['配達会社用商品コード'] = `${data['商品コード']}-${count}`;
  }
  
  // 商品名称変換: 【3回定期便】の部分を【3回定期便1回目】に置換
  if (data['商品名称']) {
    const originalProductName = data['商品名称'].toString();
    console.log(`🔍 子マスタ${count}生成 - 元の商品名称: "${originalProductName}"`);
    
    // 定期の種類を判定
    const subscriptionType = determineSubscriptionType(originalProductName);
    let typeSuffix = '';
    
    if (subscriptionType.includes('回定期便')) {
      // 回定期便の場合: 【2回定期便1回目】形式
      const totalCount = determineSubscriptionCount(originalProductName);
      typeSuffix = `${totalCount}回定期便${count}回目`;
    } else if (subscriptionType.includes('ヶ月定期便')) {
      // ヶ月定期便の場合: 【3ヶ月定期便1ヶ月目】形式
      const totalMonths = determineSubscriptionMonths(originalProductName);
      typeSuffix = `${totalMonths}ヶ月定期便${count}ヶ月目`;
    } else {
      // デフォルト: 【2回定期便1回目】形式
      const totalCount = determineSubscriptionCount(originalProductName);
      typeSuffix = `${totalCount}回定期便${count}回目`;
    }
    
    console.log(`  - 生成するtypeSuffix: "${typeSuffix}"`);
    
    // 【3回定期便】や【3ヶ月定期便】の部分を【3回定期便1回目】や【3ヶ月定期便1ヶ月目】に置換
    let productName = originalProductName
      .replace(/【\d+回定期便】/, `【${typeSuffix}】`)
      .replace(/【\d+ヶ月定期便】/, `【${typeSuffix}】`)
      .replace(/【\d+か月定期便】/, `【${typeSuffix}】`)
      .replace(/【\d+月定期便】/, `【${typeSuffix}】`);
    console.log(`  - 置換後の商品名称: "${productName}"`);
    
    // 商品名称を設定
    childData['商品名称'] = productName;
    console.log(`  - 生成後の商品名称: "${childData['商品名称']}"`);
    
    // 伝票記載用商品名: ()形式で回数を記載、元の商品名も保持（例: (2回定期便1回目)商品名）
    let productNameWithoutBracket = originalProductName
      .replace(/【\d+回定期便】/, '')
      .replace(/【\d+ヶ月定期便】/, '')
      .replace(/【\d+か月定期便】/, '')
      .replace(/【\d+月定期便】/, '');
    childData['配送伝票商品名称'] = `(${typeSuffix})${productNameWithoutBracket}`;
    console.log(`  - 生成後の配送伝票商品名称: "${childData['配送伝票商品名称']}"`);
  }
  
  return childData;
}

/**
 * 親マスタを生成
 * @param {Object} data - 元データ
 * @param {number} count - 定期便回数
 * @returns {Object} 親マスタデータ
 */
function generateParentMaster(data, count) {
  const parentData = { ...data };
  
  // 親マスタ専用項目設定
  parentData['定期便フラグ'] = '有';
  parentData['定期便回数'] = count.toString();
  parentData['定期便種別'] = determineSubscriptionType(data['商品名称']);
  
  // 商品コードを配達会社用商品コードにコピー
  if (data['商品コード']) {
    parentData['配達会社用商品コード'] = data['商品コード'];
  }
  
  // 定期便/コラボ商品コードの設定
  const subscriptionType = determineSubscriptionType(data['商品名称']);
  if (subscriptionType.includes('回定期便')) {
    // 回定期便の場合：子マスタの商品コードを順番に反映
    for (let i = 1; i <= count; i++) {
      const childCode = `${data['商品コード']}-${i}`;
      parentData[`定期便/コラボ商品コード${i}`] = childCode;
      console.log(`  - 定期便/コラボ商品コード${i}: "${childCode}"`);
    }
  } else if (subscriptionType.includes('ヶ月定期便')) {
    // ヶ月定期便の場合：配送月が読み取れないため「要確認」
    for (let i = 1; i <= 12; i++) {
      parentData[`定期便/コラボ商品コード${i}`] = '要確認';
      console.log(`  - 定期便/コラボ商品コード${i}: "要確認"`);
    }
  }
  
  // 商品名称は元のまま保持（変更不要）
  // 配送伝票商品名称のみ設定
  if (data['商品名称']) {
    const originalProductName = data['商品名称'].toString();
    console.log(`🔍 親マスタ生成 - 元の商品名称: "${originalProductName}"`);
    
    // 定期の種類を判定
    const subscriptionType = determineSubscriptionType(originalProductName);
    let typeSuffix = '';
    
    if (subscriptionType.includes('回定期便')) {
      // 回定期便の場合
      typeSuffix = `${count}回定期便`;
    } else if (subscriptionType.includes('ヶ月定期便')) {
      // ヶ月定期便の場合
      const totalMonths = determineSubscriptionMonths(originalProductName);
      typeSuffix = `${totalMonths}ヶ月定期便`;
    } else {
      // デフォルト
      typeSuffix = `${count}回定期便`;
    }
    
    console.log(`  - 生成するtypeSuffix: "${typeSuffix}"`);
    
    // 伝票記載用商品名: ()形式で定期の種類のみ記載、元の商品名も保持
    // 【3回定期便】や【3ヶ月定期便】の部分を除去してから定期便種別を先頭に追加
    let productNameWithoutBracket = originalProductName
      .replace(/【\d+回定期便】/, '')
      .replace(/【\d+ヶ月定期便】/, '')
      .replace(/【\d+か月定期便】/, '')
      .replace(/【\d+月定期便】/, '');
    parentData['配送伝票商品名称'] = `(${typeSuffix})${productNameWithoutBracket}`;
    console.log(`  - 生成後の配送伝票商品名称: "${parentData['配送伝票商品名称']}"`);
  }
  
  return parentData;
}

/**
 * 定期便種別を判定
 * @param {string} productName - 商品名称
 * @returns {string} 定期便種別
 */
function determineSubscriptionType(productName) {
  if (!productName) return '';
  
  const text = productName.toString();
  
  // ヶ月定期便の判定（優先度を高く設定）
  const monthPatterns = [
    /(\d+)ヶ月定期便/,
    /(\d+)か月定期便/,
    /(\d+)月定期便/,
    /(\d+)ヶ月/,
    /(\d+)か月/,
    /(\d+)月/
  ];
  
  for (const pattern of monthPatterns) {
    const match = text.match(pattern);
    if (match) {
      const months = match[1];
      console.log(`🔍 定期便種別判定: "${text}" → ${months}ヶ月定期便`);
      return `月によらず商品配送順指定`;
    }
  }
  
  // 回定期便の判定
  const countPatterns = [
    /(\d+)回定期便/,
    /(\d+)回/,
    /(\d+)回目/
  ];
  
  for (const pattern of countPatterns) {
    const match = text.match(pattern);
    if (match) {
      const count = match[1];
      console.log(`🔍 定期便種別判定: "${text}" → ${count}回定期便`);
      return `月ごとに商品指定`;
    }
  }
  
  // デフォルト
  console.log(`🔍 定期便種別判定: "${text}" → デフォルト回定期便`);
  return '月ごとに商品指定';
}

/**
 * 期間処理（受付期間・発送期間）
 * @param {Object} cleanedData - クレンジングされたデータ
 * @returns {Object} 期間処理後のデータ
 */
function processPeriodData(cleanedData) {
  try {
    const processedData = {};
    
    Object.keys(cleanedData).forEach(col => {
      const columnData = cleanedData[col];
      const processedColumnData = { ...columnData.data };
      
      // 受付期間の処理
      const receptionStart = processedColumnData['受付期間(開始)'];
      const receptionEnd = processedColumnData['受付期間(終了)'];
      
      if (receptionStart || receptionEnd) {
        const receptionResult = processPeriod(
          receptionStart, 
          receptionEnd, 
          '受付期間(開始)', 
          '受付期間(終了)', 
          '受付期間種別'
        );
        
        // 処理結果を反映
        processedColumnData['受付期間(開始)'] = receptionResult.startDate;
        processedColumnData['受付期間(終了)'] = receptionResult.endDate;
        processedColumnData['受付期間種別'] = receptionResult.type;
        
        console.log(`🔍 受付期間処理完了: 開始="${receptionResult.startDate}", 終了="${receptionResult.endDate}", 種別="${receptionResult.type}"`);
      }
      
      // 発送期間の処理
      const shippingStart = processedColumnData['発送期間(開始)'];
      const shippingEnd = processedColumnData['発送期間(終了)'];
      
      if (shippingStart || shippingEnd) {
        const shippingResult = processPeriod(
          shippingStart, 
          shippingEnd, 
          '発送期間(開始)', 
          '発送期間(終了)', 
          '発送期間種別'
        );
        
        // 処理結果を反映
        processedColumnData['発送期間(開始)'] = shippingResult.startDate;
        processedColumnData['発送期間(終了)'] = shippingResult.endDate;
        processedColumnData['発送期間種別'] = shippingResult.type;
        
        console.log(`🔍 発送期間処理完了: 開始="${shippingResult.startDate}", 終了="${shippingResult.endDate}", 種別="${shippingResult.type}"`);
      }
      
      processedData[col] = {
        type: columnData.type,
        data: processedColumnData
      };
    });
    
    console.log(`📅 期間処理完了: ${Object.keys(processedData).length}列分の期間データを処理`);
    return processedData;
    
  } catch (error) {
    console.error('❌ 期間処理エラー:', error);
    return cleanedData; // エラーの場合は元のデータを返す
  }
}

/**
 * 個別期間の処理
 * @param {string} startDate - 開始日
 * @param {string} endDate - 終了日
 * @param {string} startField - 開始日フィールド名
 * @param {string} endField - 終了日フィールド名
 * @param {string} typeField - 種別フィールド名
 * @returns {Object} 処理結果
 */
function processPeriod(startDate, endDate, startField, endField, typeField) {
  const result = {
    startDate: '',
    endDate: '',
    type: ''
  };
  
  console.log(`🔍 期間処理開始: 開始="${startDate}", 終了="${endDate}"`);
  
  // パターン①: 両方に「通年」キーワードが含まれている場合
  if (isYearRound(startDate) && isYearRound(endDate)) {
    result.type = '通年扱い';
    result.startDate = '';
    result.endDate = '';
    console.log(`  - パターン①: 両方「通年」→「通年扱い」`);
  }
  // パターン②: 開始日付 + 終了「通年」
  else if (!isYearRound(startDate) && isYearRound(endDate)) {
    result.type = '季節限定扱い';
    result.startDate = normalizeDateWithEarlyLate(startDate, true); // 開始日
    result.endDate = '2099/12/31'; // 固定値
    console.log(`  - パターン②: 開始日付+終了「通年」→「季節限定扱い」`);
  }
  // パターン④: 開始「通年」 + 終了日付
  else if (isYearRound(startDate) && !isYearRound(endDate)) {
    result.type = '季節限定扱い';
    
    // 外部シートから日付を取得（情報抽出タブB1の値をキーとして使用）
    try {
      const b1Value = getInfoExtractionB1Value();
      if (b1Value) {
        result.startDate = getExternalPriceValueOptimized(b1Value);
        console.log(`  - パターン④: 外部シートから日付取得 "${result.startDate}" (キー: "${b1Value}")`);
      } else {
        result.startDate = getTodayDate();
        console.log(`  - パターン④: B1の値が空、今日の日付を使用 "${result.startDate}"`);
      }
    } catch (error) {
      result.startDate = getTodayDate();
      console.log(`  - パターン④: 外部シート取得失敗、今日の日付を使用 "${result.startDate}"`);
    }
    
    result.endDate = normalizeDateWithEarlyLate(endDate, false); // 終了日
  }
  // 両方に日付が入っている場合
  else if (startDate && endDate) {
    result.type = '季節限定扱い';
    result.startDate = normalizeDateWithEarlyLate(startDate, true); // 開始日
    result.endDate = normalizeDateWithEarlyLate(endDate, false); // 終了日
    console.log(`  - 両方日付: 「季節限定扱い」`);
  }
  // どちらか一方に日付が入っている場合
  else if (startDate || endDate) {
    result.type = '季節限定扱い';
    result.startDate = startDate ? normalizeDateWithEarlyLate(startDate, true) : ''; // 開始日
    result.endDate = endDate ? normalizeDateWithEarlyLate(endDate, false) : ''; // 終了日
    console.log(`  - 片方日付: 「季節限定扱い」`);
  }
  
  console.log(`🔍 期間処理完了: 開始="${result.startDate}", 終了="${result.endDate}", 種別="${result.type}"`);
  return result;
}

/**
 * 「通年」キーワードの判定
 * @param {string} text - 判定対象テキスト
 * @returns {boolean} 通年キーワードが含まれているか
 */
function isYearRound(text) {
  if (!text) return false;
  
  const yearRoundKeywords = ['通年', '順次', '随時', '常時', '準備でき次第', '受付でき次第'];
  const textStr = text.toString().toLowerCase();
  
  return yearRoundKeywords.some(keyword => textStr.includes(keyword));
}

/**
 * 今日の日付を取得（yyyy/mm/dd形式）
 * @returns {string} 今日の日付
 */
function getTodayDate() {
  const today = new Date();
  const year = today.getFullYear();
  const month = String(today.getMonth() + 1).padStart(2, '0');
  const day = String(today.getDate()).padStart(2, '0');
  return `${year}/${month}/${day}`;
}

/**
 * 情報抽出タブのB1の値を取得
 * @returns {string} B1の値
 */
function getInfoExtractionB1Value() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetNames = spreadsheet.getSheets().map(sheet => sheet.getName());
    console.log(`🔍 利用可能なシート名: ${JSON.stringify(sheetNames)}`);
    
    const sheet = spreadsheet.getSheetByName('情報抽出');
    if (sheet) {
      const b1Value = sheet.getRange('B1').getValue();
      console.log(`🔍 情報抽出タブB1の値: "${b1Value}"`);
      return b1Value ? b1Value.toString() : '';
    }
    console.log('⚠️ 情報抽出シートが見つかりません');
    return '';
  } catch (error) {
    console.error('❌ 情報抽出シートB1取得エラー:', error);
    return '';
  }
}

/**
 * 日付の正規化処理（上旬・下旬対応含む）
 * @param {string} dateText - 日付テキスト
 * @param {boolean} isStartDate - 開始日かどうか（デフォルト: true）
 * @returns {string} 正規化された日付
 */
function normalizeDateWithEarlyLate(dateText, isStartDate = true) {
  if (!dateText) return '';
  
  // 上旬・下旬が含まれている場合は変換
  if (dateText.includes('上旬') || dateText.includes('下旬')) {
    return convertEarlyLateToDate(dateText, null, isStartDate);
  }
  
  // 通常の日付正規化
  return normalizeDateFormat(dateText);
}

/**
 * 上旬・下旬を具体的な日付に変換
 * @param {string} text - 変換対象テキスト
 * @param {number} year - 年（デフォルト: 現在の年）
 * @param {boolean} isStartDate - 開始日かどうか（デフォルト: true）
 * @returns {string} 変換後の日付（yyyy/mm/dd形式）
 */
function convertEarlyLateToDate(text, year = null, isStartDate = true) {
  if (!text) return '';
  
  const textStr = text.toString();
  
  // 年が指定されていない場合は現在の年を使用
  const currentYear = year || new Date().getFullYear();
  
  // 年月の抽出パターン
  const monthPatterns = [
    /(\d{1,2})月/,
    /(\d{4})年(\d{1,2})月/
  ];
  
  console.log(`🔍 正規表現テスト: "${textStr}"`);
  console.log(`  - パターン1 (\\d{1,2})月: ${/(\d{1,2})月/g.test(textStr)}`);
  console.log(`  - パターン2 (\\d{4})年(\\d{1,2})月: ${/(\d{4})年(\d{1,2})月/g.test(textStr)}`);
  
  console.log(`🔍 上旬・下旬変換開始: "${text}"`);
  
  for (const pattern of monthPatterns) {
    const match = textStr.match(pattern);
    console.log(`  - パターン ${pattern.source} でマッチ: ${match ? 'あり' : 'なし'}`);
    if (match) {
      console.log(`  - マッチ結果: ${JSON.stringify(match)}`);
    }
    
    if (match) {
      let targetYear = currentYear;
      let targetMonth;
      
      if (pattern.source.includes('年')) {
        // 2025年6月形式
        const fullMatch = textStr.match(/(\d{4})年(\d{1,2})月/);
        if (fullMatch) {
          targetYear = parseInt(fullMatch[1]);
          targetMonth = parseInt(fullMatch[2]);
          console.log(`  - 年あり形式: ${targetYear}年${targetMonth}月`);
        }
      } else {
        // 6月形式（年が入っていない場合は現在の年を使用）
        targetMonth = parseInt(match[1]);
        console.log(`🔍 年なし上旬・下旬変換: "${text}" → ${currentYear}年${targetMonth}月 (match[1]: "${match[1]}")`);
      }
      
      if (targetMonth >= 1 && targetMonth <= 12) {
        if (textStr.includes('上旬')) {
          // 上旬の変換ルール
          let day;
          if (isStartDate) {
            day = '01'; // 開始日: 上旬→1日
          } else {
            day = '15'; // 終了日: 上旬→15日
          }
          const result = `${targetYear}/${String(targetMonth).padStart(2, '0')}/${day}`;
          console.log(`  - 上旬変換(${isStartDate ? '開始日' : '終了日'}): "${text}" → "${result}"`);
          return result;
        } else if (textStr.includes('下旬')) {
          // 下旬の変換ルール
          let day;
          if (isStartDate) {
            day = '16'; // 開始日: 下旬→16日
          } else {
            day = String(new Date(targetYear, targetMonth, 0).getDate()); // 終了日: 下旬→末日
          }
          const result = `${targetYear}/${String(targetMonth).padStart(2, '0')}/${String(day).padStart(2, '0')}`;
          console.log(`  - 下旬変換(${isStartDate ? '開始日' : '終了日'}): "${text}" → "${result}"`);
          return result;
        }
      }
    }
  }
  
  // 上旬・下旬が含まれていない場合は元のテキストを返す
  console.log(`🔍 上旬・下旬なし: "${text}" → そのまま返却`);
  return text;
}

/**
 * 配送会社名を変換
 * @param {string} shippingCompany - 元の配送会社名
 * @returns {string} 変換後の配送会社名
 */
function convertShippingCompany(shippingCompany) {
  if (!shippingCompany) return '';
  
  const company = shippingCompany.toString();
  
  // デフォルト値の場合は空文字を設定
  if (company.includes('配送方法をお選びください') || company.includes('選択してください')) {
    console.log(`🚚 配送会社変換: "${company}" → "" (デフォルト値)`);
    return '';
  }
  
  // ヤマト運輸
  if (company.includes('ヤマト')) {
    console.log(`🚚 配送会社変換: "${company}" → "ヤマト運輸"`);
    return 'ヤマト運輸';
  }
  
  // 佐川急便
  if (company.includes('佐川')) {
    console.log(`🚚 配送会社変換: "${company}" → "佐川急便"`);
    return '佐川急便';
  }
  
  // 日本郵便（OR検索）
  if (company.includes('パック') || company.includes('レター') || company.includes('郵便')) {
    console.log(`🚚 配送会社変換: "${company}" → "日本郵便"`);
    return '日本郵便';
  }
  
  // 変換対象外の場合は元の値を返す
  console.log(`ℹ️ 配送会社変換対象外: "${company}" (そのまま)`);
  return company;
}

/**
 * 税率種別を変換
 * @param {string} taxType - 元の税率種別
 * @returns {string} 変換後の税率種別
 */
function convertTaxType(taxType) {
  if (!taxType) return '';
  
  const tax = taxType.toString();
  
  // 標準税率
  if (tax.includes('標準') || tax.includes('10%') || tax.includes('10％')) {
    console.log(`💰 税率種別変換: "${tax}" → "標準税率"`);
    return '標準税率';
  }
  
  // 軽減税率
  if (tax.includes('軽減') || tax.includes('8%') || tax.includes('8％')) {
    console.log(`💰 税率種別変換: "${tax}" → "軽減税率"`);
    return '軽減税率';
  }
  
  // 非課税
  if (tax.includes('非課税') || tax.includes('0%') || tax.includes('0％') || tax.includes('免税')) {
    console.log(`💰 税率種別変換: "${tax}" → "非課税"`);
    return '非課税';
  }
  
  // 上記以外の場合は空文字を設定
  console.log(`💰 税率種別変換: "${tax}" → "" (未対応の値)`);
  return '';
}

/**
 * Do書き出し用タブにデータを出力
 * @param {Object} cleanedData - クレンジングされたデータ
 * @param {Object} productTypes - 商品種別マップ
 * @returns {boolean} 出力結果
 */
function outputToDoTabs(cleanedData, productTypes) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const singleTab = ss.getSheetByName('Do書き出し用');
    const subscriptionTab = ss.getSheetByName('Do書き出し用(定期)');
    
    if (!singleTab) {
      throw new Error('Do書き出し用タブが見つかりません');
    }
    
    if (!subscriptionTab) {
      throw new Error('Do書き出し用(定期)タブが見つかりません');
    }
    
    let singleCount = 0;
    let subscriptionCount = 0;
    
    // 期間処理（受付期間・発送期間）を先に実行
    const periodProcessedData = processPeriodData(cleanedData);
    
    // 定期便の特別処理を実行
    const processedData = processSubscriptionProducts(periodProcessedData);
    
    // 単一商品の出力
    Object.keys(periodProcessedData).forEach(col => {
      const data = periodProcessedData[col];
      if (data.type === 'single') {
        const result = outputToSingleTab(singleTab, data.data);
        if (result) singleCount++;
      }
    });
    
    // 定期便の出力（子マスタ・親マスタ）
    Object.keys(processedData).forEach(key => {
      const data = processedData[key];
      if (data.type === 'subscription_child' || data.type === 'subscription_parent') {
        const result = outputToSubscriptionTab(subscriptionTab, data.data);
        if (result) subscriptionCount++;
      }
    });
    
    console.log(`📤 Do書き出し用タブへの出力完了: 単一商品${singleCount}件、定期便${subscriptionCount}件`);
    return true;
    
  } catch (error) {
    console.error('❌ Do書き出し用タブへの出力エラー:', error);
    return false;
  }
}

/**
 * 単一商品用タブにデータを出力
 * @param {Sheet} tab - 出力先タブ
 * @param {Object} data - 出力データ
 * @returns {boolean} 出力結果
 */
function outputToSingleTab(tab, data) {
  try {
    // データの最終行に追加（上書き防止）
    const lastRow = tab.getLastRow();
    const targetRow = lastRow + 1;
    // 列インデックスキャッシュ（ヘッダー→列番号）を作成
    const columnIndexCache = createColumnIndexCache(tab);
    
    // 出力用のデータを作成（固定値と外部シート参照を含む）
    const outputData = { ...data };
    
    // 固定値設定
    outputData['寄附金額(終了)1'] = '2099/12/31';
    outputData['在庫数'] = '99999';
    outputData['アラート在庫数'] = '1';
    outputData['出荷可能日フラグ(月)'] = '有';
    outputData['出荷可能日フラグ(火)'] = '有';
    outputData['出荷可能日フラグ(水)'] = '有';
    outputData['出荷可能日フラグ(木)'] = '有';
    outputData['出荷可能日フラグ(金)'] = '有';
    outputData['出荷可能日フラグ(土)'] = '有';
    outputData['出荷可能日フラグ(日)'] = '有';
    outputData['出荷可能日フラグ(祝日)'] = '有';
    outputData['出品ステータス'] = '出品中';
    
    // 配送会社変換処理
    if (outputData['配送会社']) {
      outputData['配送会社'] = convertShippingCompany(outputData['配送会社']);
    }
    
    // 税率種別変換処理
    if (outputData['税率種別']) {
      outputData['税率種別'] = convertTaxType(outputData['税率種別']);
    }
    
    // 外部シート参照設定
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const infoSheet = ss.getSheetByName('情報抽出');
    if (infoSheet) {
      const keyValue = infoSheet.getRange('B1').getValue();
      const externalValue = getExternalPriceValueOptimized(keyValue);
      if (externalValue) {
        outputData['寄附金額(開始)1'] = externalValue;
        outputData['提供価格(開始)1'] = externalValue;
        console.log(`✅ 単品: 寄附金額(開始)1と提供価格(開始)1を外部シートの値に設定: "${externalValue}"`);
      }
    }
    
    // 変換を一括適用
    const convertedData = applyDataConversionsOptimized(outputData);
    // 項目名をキーとして適切な列にデータを配置（キャッシュ利用）
    Object.keys(convertedData).forEach(itemName => {
      const columnIndex = columnIndexCache[itemName] || 0;
      if (columnIndex > 0) {
        tab.getRange(targetRow, columnIndex).setValue(convertedData[itemName]);
      }
    });
    
    return true;
    
  } catch (error) {
    console.error('❌ 単一商品用タブへの出力エラー:', error);
    return false;
  }
}

/**
 * 定期便用タブにデータを出力
 * @param {Sheet} tab - 出力先タブ
 * @param {Object} data - 出力データ
 * @returns {boolean} 出力結果
 */
function outputToSubscriptionTab(tab, data) {
  try {
    // データの最終行に追加（上書き防止）
    const lastRow = tab.getLastRow();
    const targetRow = lastRow + 1;
    // 列インデックスキャッシュ（ヘッダー→列番号）を作成
    const columnIndexCache = createColumnIndexCache(tab);
    
    // 出力用のデータを作成（固定値と外部シート参照を含む）
    const outputData = { ...data };
    
    // 固定値設定
    outputData['寄附金額(終了)1'] = '2099/12/31';
    outputData['在庫数'] = '99999';
    outputData['アラート在庫数'] = '1';
    outputData['出荷可能日フラグ(月)'] = '有';
    outputData['出荷可能日フラグ(火)'] = '有';
    outputData['出荷可能日フラグ(水)'] = '有';
    outputData['出荷可能日フラグ(木)'] = '有';
    outputData['出荷可能日フラグ(金)'] = '有';
    outputData['出荷可能日フラグ(土)'] = '有';
    outputData['出荷可能日フラグ(日)'] = '有';
    outputData['出荷可能日フラグ(祝日)'] = '有';
    outputData['出品ステータス'] = '出品中';
    
    // 配送会社変換処理
    if (outputData['配送会社']) {
      outputData['配送会社'] = convertShippingCompany(outputData['配送会社']);
    }
    
    // 税率種別変換処理
    if (outputData['税率種別']) {
      outputData['税率種別'] = convertTaxType(outputData['税率種別']);
    }
    
    // 外部シート参照設定
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const infoSheet = ss.getSheetByName('情報抽出');
    if (infoSheet) {
      const keyValue = infoSheet.getRange('B1').getValue();
      const externalValue = getExternalPriceValueOptimized(keyValue);
      if (externalValue) {
        outputData['寄附金額(開始)1'] = externalValue;
        outputData['提供価格(開始)1'] = externalValue;
        console.log(`✅ 定期便: 寄附金額(開始)1と提供価格(開始)1を外部シートの値に設定: "${externalValue}"`);
      }
    }
    
    // 変換を一括適用
    const convertedData = applyDataConversionsOptimized(outputData);
    
    // バッチ処理による一括出力（高速化）
    const lastColumn = tab.getLastColumn();
    const outputValues = new Array(lastColumn).fill('');
    
    Object.keys(convertedData).forEach(itemName => {
      const columnIndex = columnIndexCache[itemName] || 0;
      if (columnIndex > 0) {
        outputValues[columnIndex - 1] = convertedData[itemName];
      }
    });
    
    // 一括で値を設定
    const outputRange = tab.getRange(targetRow, 1, 1, lastColumn);
    outputRange.setValues([outputValues]);
    
    return true;
    
  } catch (error) {
    console.error('❌ 定期便用タブへの出力エラー:', error);
    return false;
  }
}

/**
 * 項目名から列インデックスを検索
 * @param {Sheet} tab - 対象タブ
 * @param {string} itemName - 項目名
 * @returns {number} 列インデックス（見つからない場合は0）
 */
function findColumnIndexByItemName(tab, itemName) {
  try {
    // 1行目（ヘッダー行）から項目名を検索
    const headerRow = tab.getRange(1, 1, 1, tab.getLastColumn());
    const values = headerRow.getValues()[0];
    
    for (let i = 0; i < values.length; i++) {
      if (values[i] === itemName) {
        return i + 1; // 1ベースのインデックスに変換
      }
    }
    
    return 0; // 見つからない場合
    
  } catch (error) {
    console.error('❌ 項目名から列インデックス検索エラー:', error);
    return 0;
  }
}

/**
 * 高速化機能: 列インデックスキャッシュ
 * 項目名から列インデックスへの検索を高速化
 * @param {Sheet} tab - 対象タブ
 * @returns {Object} 列インデックスキャッシュ
 */
function createColumnIndexCache(tab) {
  try {
    const columnIndexCache = {};
    const headerRow = tab.getRange(1, 1, 1, tab.getLastColumn()).getValues()[0];
    
    headerRow.forEach((header, index) => {
      if (header) {
        columnIndexCache[header] = index + 1;
      }
    });
    
    console.log(`📋 列インデックスキャッシュ作成完了: ${Object.keys(columnIndexCache).length}項目`);
    return columnIndexCache;
    
  } catch (error) {
    console.error('❌ 列インデックスキャッシュ作成エラー:', error);
    return {};
  }
}

/**
 * 高速化機能: 外部シート参照キャッシュ
 * 同じキー値に対する外部シートアクセスをキャッシュ
 */
let globalExternalValueCache = {};

/**
 * 高速化機能: 外部シート参照最適化版
 * @param {string} keyValue - 情報抽出タブB1の値
 * @returns {string} 寄附金額(開始)1の値
 */
function getExternalPriceValueOptimized(keyValue) {
  try {
    if (!keyValue) return '';
    
    // キャッシュに存在しない場合のみ外部シートにアクセス
    if (!globalExternalValueCache[keyValue]) {
      console.log(`🔍 外部シートアクセス: "${keyValue}"`);
      
      const externalSheetId = '1aRAvMW8-VEVmZQbAHiIas53Jcq6QVR8E0bE6tgTiL3s';
      const sheetName = '商品マスタ登録依頼表(CS) 2025/05/01';
      
      const externalSheet = SpreadsheetApp.openById(externalSheetId);
      const targetSheet = externalSheet.getSheetByName(sheetName);
      
      if (targetSheet) {
        const searchValue = keyValue;
        const data = targetSheet.getDataRange().getValues();
        
        for (let i = 0; i < data.length; i++) {
          if (data[i][7] && data[i][7].toString().includes(searchValue)) {
            globalExternalValueCache[keyValue] = data[i][3] || '';
            console.log(`✅ 外部シート値取得: "${keyValue}" → "${globalExternalValueCache[keyValue]}"`);
            break;
          }
        }
      }
      
      // 見つからない場合は空文字をキャッシュ
      if (!globalExternalValueCache[keyValue]) {
        globalExternalValueCache[keyValue] = '';
        console.log(`⚠️ 外部シート値未発見: "${keyValue}"`);
      }
    } else {
      console.log(`📋 外部シートキャッシュ使用: "${keyValue}" → "${globalExternalValueCache[keyValue]}"`);
    }
    
    return globalExternalValueCache[keyValue];
    
  } catch (error) {
    console.error('❌ 外部シート参照最適化エラー:', error);
    return '';
  }
}

/**
 * 高速化機能: データ変換処理最適化
 * 変換ルールを事前定義して一括処理
 * @param {Object} outputData - 出力データ
 * @returns {Object} 変換後のデータ
 */
function applyDataConversionsOptimized(outputData) {
  try {
    // 変換ルールを事前定義
    const conversionRules = {
      '配送会社': convertShippingCompany,
      '税率種別': convertTaxType
    };
    
    // 一括変換処理
    Object.keys(conversionRules).forEach(field => {
      if (outputData[field]) {
        const originalValue = outputData[field];
        outputData[field] = conversionRules[field](outputData[field]);
        console.log(`🔄 データ変換: "${field}" "${originalValue}" → "${outputData[field]}"`);
      }
    });
    
    return outputData;
    
  } catch (error) {
    console.error('❌ データ変換処理最適化エラー:', error);
    return outputData;
  }
}

/**
 * 高速化機能: 単一商品用タブへの出力（最適化版）
 * @param {Sheet} tab - 出力先タブ
 * @param {Object} data - 出力データ
 * @returns {boolean} 出力結果
 */
function outputToSingleTabOptimized(tab, data) {
  try {
    const startTime = new Date().getTime();
    
    // 1. 列インデックスキャッシュを作成
    const columnIndexCache = createColumnIndexCache(tab);
    
    // 2. データの最終行に追加
    const lastRow = tab.getLastRow();
    const targetRow = lastRow + 1;
    
    // 3. 出力用のデータを作成
    const outputData = { ...data };
    
    // 4. 固定値設定
    outputData['寄附金額(終了)1'] = '2099/12/31';
    outputData['在庫数'] = '99999';
    outputData['アラート在庫数'] = '1';
    outputData['出荷可能日フラグ(月)'] = '有';
    outputData['出荷可能日フラグ(火)'] = '有';
    outputData['出荷可能日フラグ(水)'] = '有';
    outputData['出荷可能日フラグ(木)'] = '有';
    outputData['出荷可能日フラグ(金)'] = '有';
    outputData['出荷可能日フラグ(土)'] = '有';
    outputData['出荷可能日フラグ(日)'] = '有';
    outputData['出荷可能日フラグ(祝日)'] = '有';
    outputData['出品ステータス'] = '出品中';
    
    // 5. 外部シート参照（キャッシュ使用）
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const infoSheet = ss.getSheetByName('情報抽出');
    if (infoSheet) {
      const keyValue = infoSheet.getRange('B1').getValue();
      const externalValue = getExternalPriceValueOptimized(keyValue);
      if (externalValue) {
        outputData['寄附金額(開始)1'] = externalValue;
        outputData['提供価格(開始)1'] = externalValue;
        console.log(`✅ 単品: 寄附金額(開始)1と提供価格(開始)1を外部シートの値に設定: "${externalValue}"`);
      }
    }
    
    // 6. データ変換処理（一括処理）
    const convertedData = applyDataConversionsOptimized(outputData);
    
    // 7. キャッシュを使用して高速アクセス
    Object.keys(convertedData).forEach(itemName => {
      const columnIndex = columnIndexCache[itemName] || 0;
      if (columnIndex > 0) {
        tab.getRange(targetRow, columnIndex).setValue(convertedData[itemName]);
      }
    });
    
    const endTime = new Date().getTime();
    const executionTime = endTime - startTime;
    console.log(`⚡ 単一商品出力最適化版完了: ${executionTime}ms`);
    
    return true;
    
  } catch (error) {
    console.error('❌ 単一商品用タブへの出力最適化エラー:', error);
    return false;
  }
}

/**
 * 高速化機能: 定期便用タブへの出力（最適化版）
 * @param {Sheet} tab - 出力先タブ
 * @param {Object} data - 出力データ
 * @returns {boolean} 出力結果
 */
function outputToSubscriptionTabOptimized(tab, data) {
  try {
    const startTime = new Date().getTime();
    
    // 1. 列インデックスキャッシュを作成
    const columnIndexCache = createColumnIndexCache(tab);
    
    // 2. データの最終行に追加
    const lastRow = tab.getLastRow();
    const targetRow = lastRow + 1;
    
    // 3. 出力用のデータを作成
    const outputData = { ...data };
    
    // 4. 固定値設定
    outputData['寄附金額(終了)1'] = '2099/12/31';
    outputData['在庫数'] = '99999';
    outputData['アラート在庫数'] = '1';
    outputData['出荷可能日フラグ(月)'] = '有';
    outputData['出荷可能日フラグ(火)'] = '有';
    outputData['出荷可能日フラグ(水)'] = '有';
    outputData['出荷可能日フラグ(木)'] = '有';
    outputData['出荷可能日フラグ(金)'] = '有';
    outputData['出荷可能日フラグ(土)'] = '有';
    outputData['出荷可能日フラグ(日)'] = '有';
    outputData['出荷可能日フラグ(祝日)'] = '有';
    outputData['出品ステータス'] = '出品中';
    
    // 5. 外部シート参照（キャッシュ使用）
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const infoSheet = ss.getSheetByName('情報抽出');
    if (infoSheet) {
      const keyValue = infoSheet.getRange('B1').getValue();
      const externalValue = getExternalPriceValueOptimized(keyValue);
      if (externalValue) {
        outputData['寄附金額(開始)1'] = externalValue;
        outputData['提供価格(開始)1'] = externalValue;
        console.log(`✅ 定期便: 寄附金額(開始)1と提供価格(開始)1を外部シートの値に設定: "${externalValue}"`);
      }
    }
    
    // 6. データ変換処理（一括処理）
    const convertedData = applyDataConversionsOptimized(outputData);
    
    // 7. キャッシュを使用して高速アクセス
    Object.keys(convertedData).forEach(itemName => {
      const columnIndex = columnIndexCache[itemName] || 0;
      if (columnIndex > 0) {
        tab.getRange(targetRow, columnIndex).setValue(convertedData[itemName]);
      }
    });
    
    const endTime = new Date().getTime();
    const executionTime = endTime - startTime;
    console.log(`⚡ 定期便出力最適化版完了: ${executionTime}ms`);
    
    return true;
    
  } catch (error) {
    console.error('❌ 定期便用タブへの出力最適化エラー:', error);
    return false;
  }
}

/**
 * 高速化機能: Do書き出し用タブへの出力（最適化版）
 * @param {Object} cleanedData - クレンジングされたデータ
 * @param {Object} productTypes - 商品種別マップ
 * @returns {boolean} 出力結果
 */
function outputToDoTabsOptimized(cleanedData, productTypes) {
  try {
    const startTime = new Date().getTime();
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const singleTab = ss.getSheetByName('Do書き出し用');
    const subscriptionTab = ss.getSheetByName('Do書き出し用(定期)');
    
    if (!singleTab) {
      throw new Error('Do書き出し用タブが見つかりません');
    }
    
    if (!subscriptionTab) {
      throw new Error('Do書き出し用(定期)タブが見つかりません');
    }
    
    let singleCount = 0;
    let subscriptionCount = 0;
    
    // 期間処理（受付期間・発送期間）を先に実行
    const periodProcessedData = processPeriodData(cleanedData);
    
    // 定期便の特別処理を実行
    const processedData = processSubscriptionProducts(periodProcessedData);
    
    // 単一商品の出力（最適化版）
    Object.keys(periodProcessedData).forEach(col => {
      const data = periodProcessedData[col];
      if (data.type === 'single') {
        const result = outputToSingleTabOptimized(singleTab, data.data);
        if (result) singleCount++;
      }
    });
    
    // 定期便の出力（子マスタ・親マスタ）（最適化版）
    Object.keys(processedData).forEach(key => {
      const data = processedData[key];
      if (data.type === 'subscription_child' || data.type === 'subscription_parent') {
        const result = outputToSubscriptionTabOptimized(subscriptionTab, data.data);
        if (result) subscriptionCount++;
      }
    });
    
    const endTime = new Date().getTime();
    const executionTime = endTime - startTime;
    console.log(`⚡ Do書き出し用タブへの出力最適化版完了: 単一商品${singleCount}件、定期便${subscriptionCount}件 (${executionTime}ms)`);
    return true;
    
  } catch (error) {
    console.error('❌ Do書き出し用タブへの出力最適化エラー:', error);
    return false;
  }
}

/**
 * 高速化機能: 外部シート参照キャッシュのクリア
 * 必要に応じてキャッシュをリセット
 */
function clearExternalValueCache() {
  globalExternalValueCache = {};
  console.log('🗑️ 外部シート参照キャッシュをクリアしました');
}

/**
 * 高速化機能: 性能測定
 * 既存版と最適化版の性能を比較
 * @param {Object} cleanedData - クレンジングされたデータ
 * @param {Object} productTypes - 商品種別マップ
 */
function measurePerformance(cleanedData, productTypes) {
  try {
    console.log('📊 性能測定開始...');
    
    // 既存版の性能測定
    const startTime1 = new Date().getTime();
    const result1 = outputToDoTabs(cleanedData, productTypes);
    const endTime1 = new Date().getTime();
    const executionTime1 = endTime1 - startTime1;
    
    // キャッシュをクリア
    clearExternalValueCache();
    
    // 最適化版の性能測定
    const startTime2 = new Date().getTime();
    const result2 = outputToDoTabsOptimized(cleanedData, productTypes);
    const endTime2 = new Date().getTime();
    const executionTime2 = endTime2 - startTime2;
    
    // 結果表示
    console.log('📊 性能測定結果:');
    console.log(`  - 既存版: ${executionTime1}ms`);
    console.log(`  - 最適化版: ${executionTime2}ms`);
    console.log(`  - 高速化率: ${((executionTime1 - executionTime2) / executionTime1 * 100).toFixed(1)}%`);
    console.log(`  - 既存版結果: ${result1 ? '成功' : '失敗'}`);
    console.log(`  - 最適化版結果: ${result2 ? '成功' : '失敗'}`);
    
  } catch (error) {
    console.error('❌ 性能測定エラー:', error);
  }
}
