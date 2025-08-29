# 返礼品情報整形GAS 実装詳細

## 1. 実装概要

このドキュメントでは、Phase 1（基本データ抽出）、Phase 2（指定列データ抽出）、Phase 3（Do書き出し項目との紐付け）の実装詳細について説明します。

## 2. Phase 1: 基本データ抽出（実装完了）

### 2.1 主要関数

#### **`main()`関数**
- エントリーポイント
- Phase 1の処理を実行
- エラーハンドリングとログ出力

#### **`extractProductDataFromSheet(sheet)`関数**
- A列〜D列で「商品名」を含む列を特定
- 発見した列とその右隣列から4行目以降のデータを抽出
- 空行が5行連続したら抽出を終了

#### **`outputToInfoExtractionTab(extractedData)`関数**
- 情報抽出タブのB8、C8以降にデータを格納
- 発見した列と右隣列のデータを順次出力

### 2.2 処理フロー

```
1. ファイルパス読み込み（B1セル）
2. 自治体フォルダキー抽出
3. 自治体フォルダタブからフォルダID取得
4. Google Driveフォルダアクセス
5. Excelファイル検索・開封
6. Google Sheetsに変換
7. A列〜D列で「商品名」列を特定
8. 発見した列と右隣列の4行目以降を抽出
9. 空行5行連続で終了判定
10. 情報抽出タブのB8、C8以降に格納
```

### 2.3 出力形式

- **B8**: 発見した列の4行目の値
- **C8**: 右隣列の4行目の値
- **B9**: 発見した列の5行目の値
- **C9**: 右隣列の5行目の値
- 以下同様に続く（空行が5行連続するまで）

## 3. Phase 2: 指定列データ抽出（実装完了）

### 3.1 主要関数

#### **結合セル対応関数**
- `isMergedCell(sheet, row, col)` - 結合セルの判定
- `getMergedCellValue(sheet, row, col)` - 結合セルの値を取得

#### **列指定処理関数**
- `parseColumnSpec(columnSpec)` - カンマ区切りの列指定を解析
- `extractSpecifiedColumns(sheet, columnSpec)` - 指定された列のデータを抽出
- `extractFColumnData(sheet)` - F列起点の返礼品データを抽出（結合セル対応）

#### **出力関数**
- `outputColumnDataToInfoExtractionTab(extractedData)` - 列データを情報抽出タブに出力

### 3.2 処理フロー

#### **B2セルに指定がある場合**
```
1. B2セルの値をカンマ区切りでスプリット
2. 各列指定を数値に変換（F→6, H→8, J→10, L→12, N→14, P→16）
3. 指定された列のデータを抽出
4. 1行目-3行目: 情報抽出タブの4行目-6行目（D列から）
5. 4行目以降: 情報抽出タブの8行目以降（D列から）
```

#### **B2セルが空の場合**
```
1. F列（6列目）を起点として返礼品データを抽出
2. 結合セルの判定と処理
3. 1行目-3行目: 情報抽出タブの4行目-6行目（D列から）
4. 4行目以降: 情報抽出タブの8行目以降（D列から）
```

### 3.3 結合セル対応ロジック

#### **結合セルの判定方法**
```javascript
function isMergedCell(sheet, row, col) {
  try {
    var range = sheet.getRange(row, col);
    var mergedRanges = range.getMergedRanges();
    return mergedRanges.length > 0;
  } catch (error) {
    return false;
  }
}
```

#### **結合セルの値を取得**
```javascript
function getMergedCellValue(sheet, row, col) {
  try {
    var range = sheet.getRange(row, col);
    var mergedRanges = range.getMergedRanges();
    
    if (mergedRanges.length > 0) {
      // 結合セルの場合、左上のセルの値を取得
      var mergedRange = mergedRanges[0];
      return sheet.getRange(mergedRange.getRow(), mergedRange.getColumn()).getValue();
    } else {
      // 通常セルの場合
      return range.getValue();
    }
  } catch (error) {
    return '';
  }
}
```

### 3.4 出力形式

#### **4行目-6行目（1行目-3行目のデータ）**
- **D4**: 1行目の1列目データ
- **E4**: 1行目の2列目データ
- **F4**: 1行目の3列目データ
- **D5**: 2行目の1列目データ
- **E5**: 2行目の2列目データ
- **F5**: 2行目の3列目データ
- **D6**: 3行目の1列目データ
- **E6**: 3行目の2列目データ
- **F6**: 3行目の3列目データ

#### **8行目以降（4行目以降のデータ）**
- **D8**: 4行目の1列目データ
- **E8**: 5行目の1列目データ
- **F8**: 6行目の1列目データ
- **D9**: 7行目の1列目データ
- **E9**: 8行目の1列目データ
- **F9**: 9行目の1列目データ
- 以下同様に続く（**1列1返礼品**の形式）

## 4. Phase 3: Do書き出し項目との紐付け（実装完了）

### 4.1 概要
Phase 3では、情報抽出タブのB列・C列の値を部分検索でDoマスタの項目と紐付けを行います。

### 4.2 主要機能
- **自動マッチング**: キーワードベースの自動マッチング
- **検索タイプ対応**: AND検索・OR検索の両方に対応
- **スコア計算**: 最適なマッチング結果の選択
- **優先順位制御**: 新・旧項目の優先順位制御

### 4.3 実装詳細
- **キーワードマッチング**: 設定ファイルで定義されたキーワードによる検索
- **OR検索対応**: `searchType: 'or'`が指定された項目はOR検索で動作
- **AND検索**: デフォルトではAND検索で動作
- **スコアベース選択**: 複数のマッチング候補から最適なものを選択

### 4.4 設定例
```javascript
'提供価格(税込)1': {
  keywords: ['商品代金', '提供価格'],
  searchType: 'or'  // OR検索を指定
}
```

## 4. セル結合対応の詳細実装

### 4.1 イレギュラーパターン検知

```javascript
function detectIrregularPattern(sheet, columnIndex) {
  try {
    // 大項目の値を取得（1行目）
    const bigItemValue = sheet.getRange(1, columnIndex + 1).getValue();
    
    // 小項目の値を取得（2行目）
    const smallItemValue = sheet.getRange(2, columnIndex + 1).getValue();
    
    // イレギュラーパターン: 大項目のみ存在、小項目が空
    const isIrregular = bigItemValue && bigItemValue.trim() !== '' && 
                       (!smallItemValue || smallItemValue.trim() === '');
    
    // セル結合の行数を取得
    const mergedRowCount = getMergedRowCountForColumn(sheet, columnIndex);
    
    return {
      isIrregular: isIrregular,
      bigItemValue: bigItemValue,
      smallItemValue: smallItemValue,
      mergedRowCount: mergedRowCount
    };
    
  } catch (error) {
    console.log(`❌ イレギュラーパターン検知エラー: ${error.message}`);
    return {
      isIrregular: false,
      bigItemValue: '',
      smallItemValue: '',
      mergedRowCount: 3
    };
  }
}
```

### 4.2 セル結合行数の取得

```javascript
function getMergedRowCountForColumn(sheet, columnIndex) {
  try {
    const col = columnIndex + 1; // 0ベースから1ベースに変換
    
    // 1行目（大項目）のセル結合をチェック
    const range1 = sheet.getRange(1, col);
    const mergedRanges1 = range1.getMergedRanges();
    
    if (mergedRanges1.length > 0) {
      const mergedRange = mergedRanges1[0];
      const rowCount = mergedRange.getNumRows();
      console.log(`列${getColumnLetter(col)}: 1行目のセル結合行数=${rowCount}行`);
      return rowCount;
    }
    
    // 2行目（小項目）のセル結合をチェック
    const range2 = sheet.getRange(2, col);
    const mergedRanges2 = range2.getMergedRanges();
    
    if (mergedRanges2.length > 0) {
      const mergedRange = mergedRanges2[0];
      const rowCount = mergedRange.getNumRows();
      console.log(`列${getColumnLetter(col)}: 2行目のセル結合行数=${rowCount}行`);
      return rowCount;
    }
    
    // セル結合がない場合はデフォルト値
    return 3;
    
  } catch (error) {
    console.log(`❌ セル結合行数取得エラー: ${error.message}`);
    return 3;
  }
}
```

### 4.3 項目名の行数調整

```javascript
// 各列のイレギュラーパターンを事前分析
let columnOffsets = []; // 各列の累積オフセット
let cumulativeOffset = 0;

for (let j = 0; j < extractedData.headerData[0].length; j++) {
  const outputCol = 4 + j; // D列（4列目）から開始
  
  // イレギュラーパターンを検知（大項目のみで小項目が空の場合）
  const irregularInfo = detectIrregularPattern(sheet, j);
  
  columnOffsets[j] = cumulativeOffset;
  
  if (irregularInfo.isIrregular) {
    console.log(`📋 列${getColumnLetter(outputCol)}: 項目名にイレギュラーパターン検出`);
    console.log(`  - 大項目: "${irregularInfo.bigItemValue}"`);
    console.log(`  - この列の累積オフセット: ${cumulativeOffset}行`);
    
    // イレギュラーパターンの場合、次の列から1行追加
    cumulativeOffset += 1;
  }
}
```

### 4.4 返礼品データの出力位置調整

```javascript
// まず、全ての列の項目名におけるセル結合パターンを分析
let headerRowsNeeded = 3; // デフォルトは3行（大項目、小項目、最小項目）
let totalHeaderOffset = 0; // 項目名の累積オフセット

for (let j = 0; j < extractedData.bodyData[0].length; j++) {
  const outputCol = 4 + j; // D列（4列目）から開始
  
  // イレギュラーパターンを検知（大項目のみで小項目が空の場合）
  const irregularInfo = detectIrregularPattern(sheet, j);
  
  if (irregularInfo.isIrregular) {
    console.log(`⚠️ 列${getColumnLetter(outputCol)}: 項目名にイレギュラーパターン検出`);
    console.log(`  - 大項目: "${irregularInfo.bigItemValue}"`);
    console.log(`  - セル結合行数: ${irregularInfo.mergedRowCount}行`);
    
    // イレギュラーパターンの場合、項目名部分に1行追加
    totalHeaderOffset += 1;
    console.log(`  - 項目名累積オフセット: ${totalHeaderOffset}行`);
  }
}

// 項目名に必要な総行数を計算
const totalHeaderRows = headerRowsNeeded + totalHeaderOffset;
console.log(`📏 項目名総行数: ${totalHeaderRows}行 (基本${headerRowsNeeded}行 + オフセット${totalHeaderOffset}行)`);

// 返礼品データの開始行を計算（項目名分だけオフセット）
const dataStartRow = 8 + totalHeaderOffset;
console.log(`📍 返礼品データ開始行: ${dataStartRow}行`);
```

## 5. Phase 4: Doへの書き出し（実装予定）

### 5.1 概要
Phase 4では、Phase 1-3で抽出・紐付けされたデータを、Do書き出し用タブに整形して出力します。単一商品と定期便を自動判別し、それぞれ適切な形式でデータをクレンジングします。

### 5.2 主要機能
- **選択的処理**: チェックボックスで選択された項目のみを処理
- **商品種別自動判別**: 単一商品か定期便かの自動判定
- **データクレンジング**: 各種データの整形・変換処理
- **定期便対応**: 子マスタ・親マスタの自動生成

### 5.3 実装設計

#### **5.3.1 メイン制御フロー**
```javascript
function executePhase4() {
  try {
    // 1. チェックボックス確認
    const checkedColumns = getCheckedColumns();
    
    // 2. 商品種別判別
    const productTypes = determineProductTypes(checkedColumns);
    
    // 3. データ抽出・格納
    const extractedData = extractDataForDo(checkedColumns, productTypes);
    
    // 4. データクレンジング
    const cleanedData = cleanData(extractedData);
    
    // 5. Do書き出し用タブへの出力
    outputToDoTabs(cleanedData, productTypes);
    
    return true;
  } catch (error) {
    console.error('Phase 4 エラー:', error);
    return false;
  }
}
```

#### **5.3.2 チェックボックス確認処理**
```javascript
function getCheckedColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('情報抽出タブ');
  const checkedColumns = [];
  
  // D列以降の7行目をチェック
  for (let col = 4; col <= sheet.getLastColumn(); col++) {
    const cell = sheet.getRange(7, col);
    if (cell.getValue() === true) { // チェックボックスがチェックされている
      checkedColumns.push(col);
    }
  }
  
  return checkedColumns;
}
```

#### **5.3.3 商品種別判別処理**
```javascript
function determineProductTypes(checkedColumns) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('情報抽出タブ');
  const productTypes = {};
  
  checkedColumns.forEach(col => {
    // A列の「商品名称」行の値を確認
    const productNameCell = sheet.getRange(18, col); // 商品名称の行
    const productName = productNameCell.getValue();
    
    // 「定期」文字が含まれているかチェック
    if (productName && productName.toString().includes('定期')) {
      productTypes[col] = 'subscription'; // 定期便
    } else {
      productTypes[col] = 'single'; // 単一商品
    }
  });
  
  return productTypes;
}
```

#### **5.3.4 データ抽出・格納処理**
```javascript
function extractDataForDo(checkedColumns, productTypes) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('情報抽出タブ');
  const extractedData = {};
  
  checkedColumns.forEach(col => {
    const productType = productTypes[col];
    const columnData = {};
    
    // A8:A200の項目名をキーとしてデータを抽出
    for (let row = 8; row <= 200; row++) {
      const itemName = sheet.getRange(row, 1).getValue(); // A列の項目名
      if (itemName) {
        const dataValue = sheet.getRange(row, col).getValue();
        columnData[itemName] = dataValue;
      }
    }
    
    extractedData[col] = {
      type: productType,
      data: columnData
    };
  });
  
  return extractedData;
}
```

#### **5.3.5 データクレンジング処理**
```javascript
function cleanData(extractedData) {
  const cleanedData = {};
  
  Object.keys(extractedData).forEach(col => {
    const columnData = extractedData[col];
    const cleanedColumnData = {};
    
    Object.keys(columnData.data).forEach(itemName => {
      let value = columnData.data[itemName];
      
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
      
      cleanedColumnData[itemName] = value;
    });
    
    cleanedData[col] = {
      type: columnData.type,
      data: cleanedColumnData
    };
  });
  
  return cleanedData;
}
```

#### **5.3.6 定期便特別処理**
```javascript
function processSubscriptionProducts(cleanedData) {
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
  
  return subscriptionData;
}

function generateChildMaster(data, count) {
  const childData = { ...data };
  
  // 商品コード変換: 元コード + "-" + 回数
  if (data['商品コード']) {
    childData['商品コード'] = `${data['商品コード']}-${count}`;
  }
  
  // 商品名称変換: 定期便表記を除去
  if (data['商品名称']) {
    childData['商品名称'] = data['商品名称'].replace(/定期便?/g, '');
  }
  
  return childData;
}

function generateParentMaster(data, count) {
  const parentData = { ...data };
  
  // 親マスタ専用項目設定
  parentData['定期便フラグ'] = '有';
  parentData['定期便回数'] = count.toString();
  parentData['定期便種別'] = determineSubscriptionType(data['商品名称']);
  
  return parentData;
}
```

#### **5.3.7 Do書き出し用タブへの出力**
```javascript
function outputToDoTabs(cleanedData, productTypes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const singleTab = ss.getSheetByName('Do書き出し用');
  const subscriptionTab = ss.getSheetByName('Do書き出し用(定期)');
  
  Object.keys(cleanedData).forEach(col => {
    const data = cleanedData[col];
    
    if (data.type === 'single') {
      // 単一商品: Do書き出し用タブ
      outputToSingleTab(singleTab, data.data);
    } else if (data.type === 'subscription') {
      // 定期便: Do書き出し用(定期)タブ
      outputToSubscriptionTab(subscriptionTab, data.data);
    }
  });
}

function outputToSingleTab(tab, data) {
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
}
```

### 5.4 データクレンジング詳細

#### **5.4.1 数字抽出処理**
```javascript
function extractNumericValue(text) {
  if (!text) return '';
  
  const numericMatch = text.toString().match(/[\d,]+/);
  if (numericMatch) {
    return numericMatch[0].replace(/,/g, '');
  }
  
  return '';
}
```

#### **5.4.2 発送種別変換**
```javascript
function convertShippingType(shippingType) {
  if (!shippingType) return '';
  
  const type = shippingType.toString();
  if (type.includes('常温')) return '通常便';
  if (type.includes('冷蔵')) return '冷蔵便';
  if (type.includes('冷凍')) return '冷凍便';
  
  return shippingType;
}
```

#### **5.4.3 日付フォーマット統一**
```javascript
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
```

#### **5.4.4 通年扱い処理**
```javascript
function processYearRoundHandling(periodText) {
  if (!periodText) return '';
  
  const text = periodText.toString();
  
  // 通年扱いキーワードチェック
  if (text.includes('通年') || text.includes('順次') || text.includes('随時')) {
    return '通年扱い';
  }
  
  return periodText;
}
```

### 5.5 エラーハンドリング

#### **5.5.1 入力値検証**
- チェックボックスの存在確認
- 必須項目の値チェック
- データ型の妥当性確認

#### **5.5.2 処理エラー対応**
- ファイルアクセスエラー
- データ変換エラー
- 出力エラー

#### **5.5.3 ログ出力**
- 処理状況の詳細ログ
- エラー情報の記録
- 処理結果のサマリー

### 5.6 パフォーマンス最適化

#### **5.6.1 バッチ処理**
- 大量データの分割処理
- メモリ使用量の最適化

#### **5.6.2 キャッシュ活用**
- 頻繁にアクセスするデータのキャッシュ
- 計算結果の再利用

#### **5.6.3 非同期処理**
- 時間のかかる処理の非同期化
- ユーザー体験の向上

## 6. CONFIG定数

```javascript
var CONFIG = {
  SPREADSHEET_ID: '1W-Kmre4FTL5iU0VNSs5Z4vLVsXzFebLYlxSxnPWkxPQ',
  SHEETS: {
    INFO_EXTRACTION: '情報抽出',
    DO_EXPORT: 'Do書き出し用',
    DO_EXPORT_REGULAR: 'Do書き出し用(定期)',
    MUNICIPALITY_FOLDERS: '自治体フォルダ'
  },
  CELLS: {
    FOLDER_PATH: 'B1',
    COLUMN_SPEC: 'B2',        // Phase 2で追加
    MAPPING_START_ROW: 7
  },
  OUTPUT: {
    START_ROW: 8,
    COL_B: 'B',
    COL_C: 'C'
  }
};
```

## 7. テスト手順

### 7.1 結合セル対応のテスト
1. `isMergedCell()`関数を個別実行
2. `getMergedCellValue()`関数を個別実行
3. 結合セルを含むExcelファイルでテスト

### 7.2 列指定処理のテスト
1. `parseColumnSpec()`関数を個別実行
2. `extractSpecifiedColumns()`関数を個別実行
3. B2セルに`F,H,J,L,N,P`を設定してテスト

### 7.3 F列起点抽出のテスト
1. `extractFColumnData()`関数を個別実行
2. B2セルを空にしてテスト
3. 結合セルを含むデータでテスト

### 7.4 統合テスト
1. `main()`関数を実行
2. ログ出力を確認
3. 情報抽出タブの出力結果を確認

## 8. 実装完了後の動作確認

### 8.1 期待されるログ出力
```
=== 返礼品情報整形処理開始（Phase 1: 基本データ抽出） ===
...（Phase 1の処理）...
🔍 指定列データ抽出開始: F,H,J,L,N,P
列指定解析: F → 6列目
列指定解析: H → 8列目
列指定解析: J → 10列目
列指定解析: L → 12列目
列指定解析: N → 14列目
列指定解析: P → 16列目
📊 指定列データ抽出完了:
  - ヘッダーデータ: 3行
  - ボディデータ: 92行
📝 列データを情報抽出タブに出力開始
📝 ヘッダーデータ出力: 4行目以降 (D列から、セル結合対応版)
📝 ボディデータ出力: 8行目以降 (D列から、項目名セル結合対応版)
✅ 列データ出力完了
=== 処理完了 ===
```

### 8.2 出力結果の確認
- **4-6行目**: 1-3行目の指定列データ（D列から）
- **8行目以降**: 4行目以降の指定列データ（D列から）
- **1列1返礼品**の形式（縦方向に出力）
- 結合セルが正しく処理されている
- 指定された列のみが抽出されている

## 9. 重要な修正点

### 9.1 セル結合対応の修正
- **修正前**: 返礼品データの出力位置が列ごとにずれていた
- **修正後**: 項目名部分のみでセル結合対応し、返礼品データは固定位置から出力

### 9.2 出力ロジックの改善
- 項目名の行数調整は事前に計算
- 返礼品データの開始行は項目名の総オフセットから計算
- 各列の返礼品データは行をずらさずに出力

### 9.3 パフォーマンスの最適化
- 結合セルの判定は事前に実行
- 不要な関数呼び出しを削減
- ログ出力の最適化

## 10. Phase 5: データクリア処理（実装予定）

### 10.1 実装概要
情報抽出タブ、Do書き出し用タブ、Do書き出し用(定期)タブの内容を一括でクリアする機能を実装予定

### 10.2 実装設計

#### **10.2.1 ファイル構成**
- **Code.gs**: `executePhase5Standalone()`関数を追加
- **Phase5.gs**: 新規作成、データクリア処理のロジックを実装

#### **10.2.2 主要関数設計**
```javascript
// Code.gs
function executePhase5Standalone() {
  // Phase 5の独立実行
}

// Phase5.gs
function executePhase5() {
  // メイン処理
}

function clearInfoExtractionTab() {
  // 情報抽出タブのクリア
}

function clearDoOutputTab() {
  // Do書き出し用タブのクリア
}

function clearDoOutputSubscriptionTab() {
  // Do書き出し用(定期)タブのクリア
}
```

### 10.3 実装詳細

#### **10.3.1 情報抽出タブのクリア処理**
- D7:CQ7のチェックボックスをfalseに設定
- B6セルの内容を削除
- A8:CQ200のデータを削除
- D4:CQ6のデータを削除

#### **10.3.2 Do書き出し用タブのクリア処理**
- 1行目（項目名）を保持
- 2行目以降のデータを削除

#### **10.3.3 Do書き出し用(定期)タブのクリア処理**
- 1行目（項目名）を保持
- 2行目以降のデータを削除

### 10.4 技術要件

#### **10.4.1 エラーハンドリング**
- 各タブの存在確認
- データ削除処理のエラー処理
- 適切なログ出力

#### **10.4.2 パフォーマンス**
- 大量データの効率的な削除
- バッチ処理の活用

#### **10.4.3 安全性**
- 1行目（項目名）の保持
- 削除前の確認ログ

### 10.5 実装順序
1. Phase5.gsファイルの作成
2. 各クリア関数の実装
3. メイン処理関数の実装
4. Code.gsへの呼び出し関数追加
5. 動作確認とテスト

### 10.6 期待される動作
- 各タブの内容が適切にクリアされる
- 項目名は保持される
- 処理完了のログが出力される
- エラーが発生した場合は適切に処理される
