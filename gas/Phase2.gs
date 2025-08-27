/**
 * Phase 2: 指定列データ抽出
 * B2セルで指定列の選択的抽出、またはF列起点の返礼品データ抽出を行う
 */

/**
 * Phase 2のメイン処理
 * @param {Sheet} sheet - 処理対象のスプレッドシート
 * @returns {Object} 抽出された列データ
 */
function executePhase2(sheet) {
  try {
    console.log('=== Phase 2: 指定列データ抽出開始 ===');
    
    // シートオブジェクトの存在確認
    if (!sheet) {
      throw new Error('シートオブジェクトが渡されていません');
    }
    
    console.log(`📊 処理対象シート: ${sheet.getName()}`);
    
    // 列指定を読み込み
    const columnSpec = loadColumnSpec();
    let columnData = null;
    
    if (columnSpec && columnSpec.trim() !== '') {
      // B2セルに指定がある場合
      console.log(`🔍 指定列データ抽出開始: ${columnSpec}`);
      const columnNumbers = parseColumnSpec(columnSpec);
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
    
    console.log('=== Phase 2: 指定列データ抽出完了 ===');
    return columnData;
    
  } catch (error) {
    console.log(`❌ Phase 2エラー: ${error.message}`);
    throw error;
  }
}

/**
 * 列指定を読み込み
 * @returns {string} 列指定文字列
 */
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
    return columnSpec;
    
  } catch (error) {
    console.log(`❌ 列指定読み込みエラー: ${error.message}`);
    return '';
  }
}

/**
 * カンマ区切りの列指定を解析
 * @param {string} columnSpec - 列指定文字列（例: "F,H,J,L,N,P"）
 * @returns {Array} 列番号の配列
 */
function parseColumnSpec(columnSpec) {
  try {
    if (!columnSpec || columnSpec.trim() === '') {
      return null;
    }
    
    const columns = columnSpec.split(',').map(function(col) {
      return col.trim().toUpperCase();
    });
    
    const columnNumbers = [];
    for (let i = 0; i < columns.length; i++) {
      const col = columns[i];
      const colNum = getColumnNumber(col);
      if (colNum > 0) {
        columnNumbers.push(colNum);
        console.log(`列指定解析: ${col} → ${colNum}列目`);
      } else {
        console.log(`⚠️ 無効な列指定: ${col}`);
      }
    }
    
    return columnNumbers;
  } catch (error) {
    console.log(`列指定解析エラー: ${error.message}`);
    return null;
  }
}

/**
 * 列文字を列番号に変換
 * @param {string} columnLetter - 列文字（A, B, C...）
 * @returns {number} 列番号（1から開始）
 */
function getColumnNumber(columnLetter) {
  let result = 0;
  for (let i = 0; i < columnLetter.length; i++) {
    result *= 26;
    result += columnLetter.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
  }
  return result;
}

/**
 * 指定された列のデータを抽出
 * @param {Sheet} sheet - スプレッドシート
 * @param {Array} columnNumbers - 列番号の配列
 * @returns {Object} 抽出されたデータ
 */
function extractSpecifiedColumns(sheet, columnNumbers) {
  try {
    if (!columnNumbers || columnNumbers.length === 0) {
      console.log(`⚠️ 有効な列指定がありません`);
      return null;
    }
    
    const result = {
      headerData: [], // 1-3行目のデータ
      bodyData: []    // 4行目以降のデータ
    };
    
    const lastRow = sheet.getLastRow();
    
    // 指定列の有効性をチェック
    const validColumnNumbers = columnNumbers.filter(col => {
      if (col < 1 || col > sheet.getLastColumn()) {
        console.log(`⚠️ 列${getColumnLetter(col)}: 範囲外のため除外`);
        return false;
      }
      return true;
    });
    
    if (validColumnNumbers.length === 0) {
      console.log(`❌ 有効な列指定がありません`);
      return null;
    }
    
    console.log(`📊 指定列データ抽出開始: ${validColumnNumbers.map(col => getColumnLetter(col)).join(', ')}列 (全${lastRow}行)`);
    
    // 1-3行目のデータを抽出（結合セル対応）
    for (let row = 1; row <= 3; row++) {
      const rowData = [];
      for (let i = 0; i < validColumnNumbers.length; i++) {
        const col = validColumnNumbers[i];
        const value = getMergedCellValueWithMergeInfo(sheet, row, col);
        rowData.push(value);
      }
      result.headerData.push(rowData);
      console.log(`行${row}: ${rowData.join(' | ')}`);
    }
    
    // 4行目以降のデータを抽出（結合セル対応）
    for (let row = 4; row <= lastRow; row++) {
      const rowData = [];
      for (let i = 0; i < validColumnNumbers.length; i++) {
        const col = validColumnNumbers[i];
        const value = getMergedCellValueWithMergeInfo(sheet, row, col);
        rowData.push(value);
      }
      result.bodyData.push(rowData);
      
      // 最初の10行と最後の10行の詳細ログ
      if (row <= 13 || row >= lastRow - 9) {
        console.log(`row${row}: ${rowData.join(' | ')}`);
      }
    }
    
    console.log(`📊 指定列データ抽出完了:`);
    console.log(`  - ヘッダーデータ: ${result.headerData.length}行`);
    console.log(`  - ボディデータ: ${result.bodyData.length}行`);
    console.log(`  - 処理対象列数: ${validColumnNumbers.length}列`);
    
    return result;
  } catch (error) {
    console.log(`❌ 指定列データ抽出エラー: ${error.message}`);
    throw error;
  }
}

/**
 * F列起点の返礼品データを抽出（結合セル対応）
 * @param {Sheet} sheet - スプレッドシート
 * @returns {Object} 抽出されたデータ
 */
function extractFColumnData(sheet) {
  try {
    // シートオブジェクトの存在確認
    if (!sheet) {
      throw new Error('シートオブジェクトが渡されていません');
    }
    
    console.log(`🔍 シート情報確認: ${sheet.getName()}, 最終行: ${sheet.getLastRow()}, 最終列: ${sheet.getLastColumn()}`);
    
    const lastRow = sheet.getLastRow();
    const maxCol = sheet.getLastColumn();
    const startCol = 6; // F列（6列目）
    
    // 実際にデータが入っている列の最終位置を取得（高度版）
    const actualLastCol = getActualLastColumnAdvanced(sheet, startCol, maxCol);
    
    // 列分析の詳細ログを出力
    logColumnAnalysis(sheet, startCol, maxCol);
    
    console.log(`📊 F列起点データ抽出開始: ${startCol}列目から${actualLastCol}列目まで (全${lastRow}行)`);
    
    const result = {
      headerData: [], // 1-3行目のデータ
      bodyData: []    // 4行目以降のデータ
    };
    
    // 1-3行目のデータを抽出（結合セル対応）
    for (let row = 1; row <= 3; row++) {
      const rowData = [];
      for (let col = startCol; col <= actualLastCol; col++) {
        const value = getMergedCellValueWithMergeInfo(sheet, row, col);
        rowData.push(value);
      }
      result.headerData.push(rowData);
      console.log(`行${row}: ${rowData.join(' | ')}`);
    }
    
    // 4行目以降のデータを抽出（結合セル対応）
    for (let row = 4; row <= lastRow; row++) {
      const rowData = [];
      for (let col = startCol; col <= actualLastCol; col++) {
        const value = getMergedCellValueWithMergeInfo(sheet, row, col);
        rowData.push(value);
      }
      result.bodyData.push(rowData);
      
      // 最初の10行と最後の10行の詳細ログ
      if (row <= 13 || row >= lastRow - 9) {
        console.log(`行${row}: ${rowData.join(' | ')}`);
      }
    }
    
    console.log(`📊 F列起点データ抽出完了:`);
    console.log(`  - ヘッダーデータ: ${result.headerData.length}行`);
    console.log(`  - ボディデータ: ${result.bodyData.length}行`);
    console.log(`  - 抽出列数: ${actualLastCol - startCol + 1}列`);
    
    return result;
  } catch (error) {
    console.log(`❌ F列起点データ抽出エラー: ${error.message}`);
    throw error;
  }
}

// 結合セル関連の関数はUtils.gsで管理

/**
 * 列データを情報抽出タブに出力
 * @param {Object} extractedData - 抽出されたデータ
 */
function outputColumnDataToInfoExtractionTab(extractedData) {
  try {
    if (!extractedData) {
      console.log(`⚠️ 出力データがありません`);
      return;
    }
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.INFO_EXTRACTION);
    
    console.log(`📝 列データを情報抽出タブに出力開始`);
    
    // ヘッダーデータ（1-3行目）を4-6行目に出力
    if (extractedData.headerData && extractedData.headerData.length > 0) {
      console.log(`📝 ヘッダーデータ出力: 4-6行目 (D列から)`);
      
      for (let i = 0; i < extractedData.headerData.length; i++) {
        const rowData = extractedData.headerData[i];
        const outputRow = 4 + i; // 4行目から開始
        
        // D列から開始してデータを出力
        for (let j = 0; j < rowData.length; j++) {
          const outputCol = 4 + j; // D列（4列目）から開始
          const value = rowData[j];
          sheet.getRange(outputRow, outputCol).setValue(value);
        }
        
        console.log(`行${outputRow}: ${rowData.join(' | ')}`);
      }
    }
    
    // ボディデータ（4行目以降）を8行目以降に出力
    if (extractedData.bodyData && extractedData.bodyData.length > 0) {
      console.log(`📝 ボディデータ出力: 8行目以降 (D列から)`);
      
      for (let i = 0; i < extractedData.bodyData.length; i++) {
        const rowData = extractedData.bodyData[i];
        const outputRow = 8 + i; // 8行目から開始
        
        // D列から開始してデータを出力
        for (let j = 0; j < rowData.length; j++) {
          const outputCol = 4 + j; // D列（4列目）から開始
          const value = rowData[j];
          sheet.getRange(outputRow, outputCol).setValue(value);
        }
        
        // 最初の10行と最後の10行の詳細ログ
        if (i < 10 || i >= extractedData.bodyData.length - 10) {
          console.log(`行${outputRow}: ${rowData.join(' | ')}`);
        }
      }
    }
    
    console.log(`✅ 列データ出力完了`);
    
  } catch (error) {
    console.log(`❌ 列データ出力エラー: ${error.message}`);
    throw error;
  }
}


