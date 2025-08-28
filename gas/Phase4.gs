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
 */

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
      // A列の「商品名称」行の値を確認（18行目）
      const productNameCell = infoSheet.getRange(18, col);
      const productName = productNameCell.getValue();
      
      // 「定期」文字が含まれているかチェック
      if (productName && productName.toString().includes('定期')) {
        productTypes[col] = 'subscription'; // 定期便
      } else {
        productTypes[col] = 'single'; // 単一商品
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
    
    const extractedData = {};
    
    checkedColumns.forEach(col => {
      const productType = productTypes[col];
      const columnData = {};
      
      // A8:A200の項目名をキーとしてデータを抽出
      for (let row = 8; row <= 200; row++) {
        const itemName = infoSheet.getRange(row, 1).getValue(); // A列の項目名
        if (itemName) {
          const dataValue = infoSheet.getRange(row, col).getValue();
          columnData[itemName] = dataValue;
        }
      }
      
      extractedData[col] = {
        type: productType,
        data: columnData
      };
    });
    
    console.log(`📊 データ抽出完了: ${Object.keys(extractedData).length}列分のデータを抽出`);
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
  
  const text = dateText.toString();
  
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
        
        // 定期便回数を判定
        const subscriptionCount = determineSubscriptionCount(data['商品名称']);
        
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
  const match = text.match(/(\d+)回定期便/);
  if (match) {
    return parseInt(match[1]);
  }
  
  // デフォルトは1回
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
  const match = text.match(/(\d+)ヶ月定期便/);
  if (match) {
    return parseInt(match[1]);
  }
  
  // デフォルトは1ヶ月
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
  
  // 商品名称変換: 【3回定期便】の部分を【1回目/3回定期便】に置換
  if (data['商品名称']) {
    const originalProductName = data['商品名称'].toString();
    console.log(`🔍 子マスタ${count}生成 - 元の商品名称: "${originalProductName}"`);
    
    // 定期の種類を判定
    const subscriptionType = determineSubscriptionType(originalProductName);
    let typeSuffix = '';
    
    if (subscriptionType.includes('回定期便')) {
      // 回定期便の場合
      const totalCount = determineSubscriptionCount(originalProductName);
      typeSuffix = `${count}回目/${totalCount}回定期便`;
    } else if (subscriptionType.includes('ヶ月定期便')) {
      // ヶ月定期便の場合
      const totalMonths = determineSubscriptionMonths(originalProductName);
      typeSuffix = `${count}ヶ月目/${totalMonths}ヶ月定期便`;
    } else {
      // デフォルト
      const totalCount = determineSubscriptionCount(originalProductName);
      typeSuffix = `${count}回目/${totalCount}回定期便`;
    }
    
    console.log(`  - 生成するtypeSuffix: "${typeSuffix}"`);
    
    // 【3回定期便】の部分を【1回目/3回定期便】に置換
    let productName = originalProductName.replace(/【\d+回定期便】/, `【${typeSuffix}】`);
    console.log(`  - 置換後の商品名称: "${productName}"`);
    
    // 商品名称を設定
    childData['商品名称'] = productName;
    console.log(`  - 生成後の商品名称: "${childData['商品名称']}"`);
    
    // 伝票記載用商品名: ()形式で回数を記載、元の商品名も保持
    childData['配送伝票商品名称'] = `(${typeSuffix})${originalProductName.replace(/【\d+回定期便】/, '')}`;
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
    // 【3回定期便】の部分を除去してから(3回定期便)を先頭に追加
    const productNameWithoutBracket = originalProductName.replace(/【\d+回定期便】/, '');
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
  
  if (text.includes('回定期便')) {
    // 回定期便の場合、数字を含める
    const match = text.match(/(\d+)回定期便/);
    if (match) {
      const count = match[1];
      return `${count}回定期便：月ごとに商品指定`;
    }
    return '回定期便：月ごとに商品指定';
  } else if (text.includes('ヶ月定期便')) {
    // ヶ月定期便の場合、数字を含める
    const match = text.match(/(\d+)ヶ月定期便/);
    if (match) {
      const months = match[1];
      return `${months}ヶ月定期便：月によらず商品配送順指定`;
    }
    return 'ヶ月定期便：月によらず商品配送順指定';
  }
  
  return '回定期便：月ごとに商品指定'; // デフォルト
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
    
    // 定期便の特別処理を先に実行
    const processedData = processSubscriptionProducts(cleanedData);
    
    // 単一商品の出力
    Object.keys(cleanedData).forEach(col => {
      const data = cleanedData[col];
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
    
    // 項目名をキーとして適切な列にデータを配置
    Object.keys(data).forEach(itemName => {
      const columnIndex = findColumnIndexByItemName(tab, itemName);
      if (columnIndex > 0) {
        tab.getRange(targetRow, columnIndex).setValue(data[itemName]);
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
    
    // 外部シート参照設定
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const infoSheet = ss.getSheetByName('情報抽出');
    if (infoSheet) {
      const keyValue = infoSheet.getRange('B1').getValue();
      const externalValue = getExternalPriceValue(keyValue);
      if (externalValue) {
        outputData['寄附金額(開始)1'] = externalValue;
        outputData['提供価格(開始)1'] = externalValue;
        console.log(`✅ 定期便: 寄附金額(開始)1と提供価格(開始)1を外部シートの値に設定: "${externalValue}"`);
      }
    }
    
    // 項目名をキーとして適切な列にデータを配置
    Object.keys(outputData).forEach(itemName => {
      const columnIndex = findColumnIndexByItemName(tab, itemName);
      if (columnIndex > 0) {
        tab.getRange(targetRow, columnIndex).setValue(outputData[itemName]);
      }
    });
    
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
