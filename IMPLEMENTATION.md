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

## 5. メイン処理の統合

### 5.1 `processExcelFile()`関数の修正

```javascript
function processExcelFile(fileId, fileName) {
  var tempFileId = null;
  try {
    console.log(`📊 Excel処理開始: ${fileName}`);
    
    // ExcelファイルをGoogle Sheetsに変換
    var excelFile = DriveApp.getFileById(fileId);
    var blob = excelFile.getBlob();
    var tempFileName = 'temp_' + fileName.replace(/[^a-zA-Z0-9]/g, '_') + '_' + new Date().getTime();
    
    var resource = {
      title: tempFileName,
      mimeType: MimeType.GOOGLE_SHEETS
    };
    
    var convertedFile = Drive.Files.insert(resource, blob, { convert: true });
    tempFileId = convertedFile.id;
    console.log(`✅ Google Sheetsに変換完了: ${tempFileId}`);
    
    // 変換されたスプレッドシートを開く
    var ss = SpreadsheetApp.openById(tempFileId);
    var sheet = ss.getActiveSheet();
    
    // Phase 1: 基本データ抽出
    var extractedData = extractProductDataFromSheet(sheet);
    outputToInfoExtractionTab(extractedData);
    
    // Phase 2: 指定列データ抽出
    var columnSpec = loadColumnSpec();
    var columnData = null;
    
    if (columnSpec && columnSpec.trim() !== '') {
      // B2セルに指定がある場合
      console.log(`🔍 指定列データ抽出開始: ${columnSpec}`);
      var columnNumbers = parseColumnSpec(columnSpec);
      if (columnNumbers && columnNumbers.length > 0) {
        columnData = extractSpecifiedColumns(sheet, columnNumbers);
      }
    } else {
      // B2セルが空の場合
      console.log(`🔍 F列起点データ抽出開始`);
      columnData = extractFColumnData(sheet);
    }
    
    if (columnData) {
      outputColumnDataToInfoExtractionTab(columnData);
    }
    
    return extractedData;
    
  } catch (error) {
    console.log(`❌ Excel処理エラー: ${error.message}`);
    throw error;
  } finally {
    // 一時ファイルを削除
    if (tempFileId) {
      try {
        DriveApp.getFileById(tempFileId).setTrashed(true);
        console.log(`🗑️ 一時ファイルを削除: ${tempFileId}`);
      } catch (deleteError) {
        console.log(`⚠️ 一時ファイル削除エラー: ${deleteError.toString()}`);
      }
    }
  }
}
```

### 5.2 `loadColumnSpec()`関数

```javascript
function loadColumnSpec() {
  try {
    console.log(`🔍 列指定読み込み開始: セル${CONFIG.CELLS.COLUMN_SPEC}`);
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.INFO_EXTRACTION);
    
    if (!sheet) {
      console.log(`❌ 情報抽出タブが見つかりません`);
      return '';
    }
    
    const range = sheet.getRange(CONFIG.CELLS.COLUMN_SPEC);
    const columnSpec = range.getValue();
    
    console.log(`✅ 列指定読み込み完了: "${columnSpec}" (型: ${typeof columnSpec})`);
    console.log(`🔍 セル${CONFIG.CELLS.COLUMN_SPEC}の詳細: 値="${columnSpec}", 空文字判定=${columnSpec === ''}, null判定=${columnSpec === null}`);
    
    return columnSpec;
    
  } catch (error) {
    console.log(`❌ 列指定読み込みエラー: ${error.message}`);
    return '';
  }
}
```

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
