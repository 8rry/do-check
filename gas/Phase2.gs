/**
 * Phase 2: 指定列データ抽出（完全超高速版）
 * B2セルで指定列の選択的抽出、またはF列起点の返礼品データ抽出を行う
 * 
 * パフォーマンス最適化:
 * - バッチ処理による一括操作
 * - キャッシュ機能による重複処理回避
 * - 結合セル情報の一括取得
 * - メモリ効率化
 * - 並列処理のシミュレーション
 */

/**
 * Phase 2のメイン処理（完全超高速版）
 * @param {Sheet} sheet - 処理対象のスプレッドシート
 * @returns {Object} 抽出された列データ
 */
function executePhase2(sheet) {
  try {
    console.log('=== Phase 2: 指定列データ抽出開始（完全超高速版）===');
    const startTime = new Date();
    
    // シートオブジェクトの存在確認
    if (!sheet) {
      throw new Error('シートオブジェクトが渡されていません');
    }
    
    console.log(`📊 処理対象シート: ${sheet.getName()}`);
    
    // 超高速列指定読み込み
    const columnSpec = loadColumnSpecOptimized();
    let columnData = null;
    
    if (columnSpec && columnSpec.trim() !== '') {
      // B2セルに指定がある場合（完全超高速版）
      console.log(`🔍 指定列データ抽出開始（完全超高速版）: ${columnSpec}`);
      const columnNumbers = parseColumnSpecOptimized(columnSpec);
      if (columnNumbers && columnNumbers.length > 0) {
        columnData = extractSpecifiedColumnsFullyOptimized(sheet, columnNumbers);
      }
    } else {
      // B2セルが空の場合（完全超高速版）
      console.log(`🔍 F列起点データ抽出開始（完全超高速版）`);
      columnData = extractFColumnDataFullyOptimized(sheet);
    }
    
    if (columnData) {
      outputColumnDataToInfoExtractionTabFullyOptimized(columnData);
    }
    
    const endTime = new Date();
    const processingTime = endTime - startTime;
    console.log(`⚡ Phase 2完了: ${processingTime}ms（完全超高速版）`);
    console.log('=== Phase 2: 指定列データ抽出完了 ===');
    
    return {
      ...columnData,
      processingTime: processingTime
    };
    
  } catch (error) {
    console.log(`❌ Phase 2エラー: ${error.message}`);
    throw error;
  }
}

/**
 * 列指定を読み込み（超高速版）
 * @returns {string} 列指定文字列
 */
function loadColumnSpecOptimized() {
  try {
    // キャッシュ機能による高速化
    const cacheKey = `column_spec_${CONFIG.SPREADSHEET_ID}`;
    const cachedSpec = getPerformanceCache(cacheKey);
    if (cachedSpec && isPerformanceCacheValid(cacheKey)) {
      console.log('⚡ キャッシュヒット: 高速列指定取得');
      return cachedSpec.data;
    }
    
    console.log(`🔍 列指定読み込み開始: セル${CONFIG.CELLS.COLUMN_SPEC}`);
    
    // 現在アクティブなスプレッドシートを取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      console.log('⚠️ アクティブなスプレッドシートが見つかりません');
      return '';
    }
    
    const sheet = ss.getSheetByName(CONFIG.SHEETS.INFO_EXTRACTION);
    
    if (!sheet) {
      console.log(`❌ 情報抽出タブが見つかりません`);
      return '';
    }
    
    const range = sheet.getRange(CONFIG.CELLS.COLUMN_SPEC);
    const columnSpec = range.getValue();
    
    // 結果をキャッシュに保存
    setPerformanceCache(cacheKey, columnSpec, CONFIG.PERFORMANCE.CACHE_TTL);
    
    console.log(`✅ 列指定読み込み完了: "${columnSpec}" (型: ${typeof columnSpec})`);
    return columnSpec;
    
  } catch (error) {
    console.log(`❌ 列指定読み込みエラー: ${error.message}`);
    return '';
  }
}

/**
 * カンマ区切りの列指定を解析（超高速版）
 * @param {string} columnSpec - 列指定文字列（例: "F,H,J,L,N,P"）
 * @returns {Array} 列番号の配列
 */
function parseColumnSpecOptimized(columnSpec) {
  try {
    if (!columnSpec || columnSpec.trim() === '') {
      return null;
    }
    
    // バッチ処理による高速解析
    const columns = columnSpec.split(',').map(function(col) {
      return col.trim().toUpperCase();
    });
    
    const columnNumbers = [];
    for (let i = 0; i < columns.length; i++) {
      const col = columns[i];
      const colNum = getColumnNumber(col);
      if (colNum > 0) {
        columnNumbers.push(colNum);
      }
    }
    
    console.log(`✅ 列指定解析完了: ${columns.join(', ')} → ${columnNumbers.join(', ')}`);
    return columnNumbers;
    
  } catch (error) {
    console.log(`❌ 列指定解析エラー: ${error.message}`);
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
 * 指定された列のデータを抽出（完全超高速版・バッチ処理）
 * @param {Sheet} sheet - スプレッドシート
 * @param {Array} columnNumbers - 列番号の配列
 * @returns {Object} 抽出されたデータ
 */
function extractSpecifiedColumnsFullyOptimized(sheet, columnNumbers) {
  try {
    console.log(`🔍 指定列データ抽出開始（完全超高速版）: ${columnNumbers.join(', ')}列目`);
    
    const lastRow = sheet.getLastRow();
    const result = {
      headerData: [],
      bodyData: []
    };
    
    // バッチ処理による高速データ取得
    if (lastRow >= 1) {
      // ヘッダーデータ（1-3行目）を一括取得
      const minCol = Math.min(...columnNumbers);
      const maxCol = Math.max(...columnNumbers);
      const headerRange = sheet.getRange(1, minCol, 3, maxCol - minCol + 1);
      const headerValues = headerRange.getValues();
      
      // 結合セル情報を一括取得（エラーハンドリング付き）
      let mergedRanges = [];
      try {
        if (typeof sheet.getMergedRanges === 'function') {
          mergedRanges = sheet.getMergedRanges();
        } else {
          console.log('⚠️ getMergedRanges()が利用できません。結合セル処理をスキップします。');
        }
      } catch (error) {
        console.log(`⚠️ 結合セル情報取得エラー: ${error.message}。結合セル処理をスキップします。`);
      }
      
      const headerMergedInfo = getMergedCellInfoForRange(mergedRanges, 1, 3, minCol, maxCol);
      
      // ヘッダーデータの処理（バッチ処理）
      for (let row = 0; row < 3; row++) {
        const rowData = [];
        for (let i = 0; i < columnNumbers.length; i++) {
          const col = columnNumbers[i];
          const colIndex = col - minCol;
          const value = processMergedCellValueOptimized(headerValues[row][colIndex], row + 1, col, headerMergedInfo, sheet);
          rowData.push(value);
        }
        result.headerData.push(rowData);
        console.log(`行${row + 1}: ${rowData.join(' | ')}`);
      }
      
      // 4行目以降のデータを一括取得（バッチ処理）
      if (lastRow >= 4) {
        const bodyRange = sheet.getRange(4, minCol, lastRow - 3, maxCol - minCol + 1);
        const bodyValues = bodyRange.getValues();
        
        // 結合セル情報を一括取得（既に取得済みのmergedRangesを使用）
        const bodyMergedInfo = getMergedCellInfoForRange(mergedRanges, 4, lastRow, minCol, maxCol);
        
        // ボディデータの処理（チャンク処理によるメモリ効率化）
        const bodyData = processDataInChunks(bodyValues, (rowData, rowIndex) => {
          const actualRow = rowIndex + 4;
          const processedRow = [];
          
          for (let i = 0; i < columnNumbers.length; i++) {
            const col = columnNumbers[i];
            const colIndex = col - minCol;
            const value = processMergedCellValueOptimized(rowData[colIndex], actualRow, col, bodyMergedInfo, sheet);
            processedRow.push(value);
          }
          
          return processedRow;
        }, CONFIG.PERFORMANCE.CHUNK_SIZE || 50);
        
        result.bodyData = bodyData;
      }
    }
    
    console.log(`✅ 指定列データ抽出完了（完全超高速版）: ${result.headerData.length}行 + ${result.bodyData.length}行`);
    return result;
    
  } catch (error) {
    console.log(`❌ 指定列データ抽出エラー: ${error.message}`);
    throw error;
  }
}

/**
 * F列起点の返礼品データを抽出（完全超高速版・バッチ処理）
 * @param {Sheet} sheet - スプレッドシート
 * @returns {Object} 抽出されたデータ
 */
function extractFColumnDataFullyOptimized(sheet) {
  try {
    console.log(`🔍 F列起点データ抽出開始（完全超高速版）`);
    
    // シートオブジェクトの存在確認
    if (!sheet) {
      throw new Error('シートオブジェクトが渡されていません');
    }
    
    const lastRow = sheet.getLastRow();
    const maxCol = sheet.getLastColumn();
    const startCol = 6; // F列（6列目）
    
    // キャッシュ機能による高速化
    const cacheKey = `f_column_data_${sheet.getSheetId()}_${lastRow}_${maxCol}`;
    const cachedData = getPerformanceCache(cacheKey);
    if (cachedData && isPerformanceCacheValid(cacheKey)) {
      console.log('⚡ キャッシュヒット: 高速F列データ取得');
      return cachedData.data;
    }
    
    console.log(`🔍 シート情報確認: ${sheet.getName()}, 最終行: ${lastRow}, 最終列: ${maxCol}`);
    
    // 実際にデータが入っている列の最終位置を取得（高速版）
    const actualLastCol = getActualLastColumnOptimized(sheet, startCol, maxCol);
    
    console.log(`📊 F列起点データ抽出開始: ${startCol}列目から${actualLastCol}列目まで (全${lastRow}行)`);
    
    const result = {
      headerData: [], // 1-3行目のデータ
      bodyData: []    // 4行目以降のデータ
    };
    
    // バッチ処理による高速データ抽出
    if (lastRow >= 1) {
      // ヘッダーデータ（1-3行目）を一括取得
      const headerRange = sheet.getRange(1, startCol, 3, actualLastCol - startCol + 1);
      const headerValues = headerRange.getValues();
      
      // 結合セル情報を一括取得（エラーハンドリング付き）
      let mergedRanges = [];
      try {
        if (typeof sheet.getMergedRanges === 'function') {
          mergedRanges = sheet.getMergedRanges();
        } else {
          console.log('⚠️ getMergedRanges()が利用できません。結合セル処理をスキップします。');
        }
      } catch (error) {
        console.log(`⚠️ 結合セル情報取得エラー: ${error.message}。結合セル処理をスキップします。`);
      }
      
      const headerMergedInfo = getMergedCellInfoForRange(mergedRanges, 1, 3, startCol, actualLastCol);
      
      // ヘッダーデータの処理（バッチ処理）
      for (let row = 0; row < 3; row++) {
        const rowData = [];
        for (let col = 0; col < headerValues[row].length; col++) {
          const actualRow = row + 1;
          const actualCol = startCol + col;
          const value = processMergedCellValueOptimized(headerValues[row][col], actualRow, actualCol, headerMergedInfo, sheet);
          rowData.push(value);
        }
        result.headerData.push(rowData);
        console.log(`行${row + 1}: ${rowData.join(' | ')}`);
      }
      
      // 4行目以降のデータを一括取得（バッチ処理）
      if (lastRow >= 4) {
        const bodyRange = sheet.getRange(4, startCol, lastRow - 3, actualLastCol - startCol + 1);
        const bodyValues = bodyRange.getValues();
        
        // 結合セル情報を一括取得（既に取得済みのmergedRangesを使用）
        const bodyMergedInfo = getMergedCellInfoForRange(mergedRanges, 4, lastRow, startCol, actualLastCol);
        
        // ボディデータの処理（チャンク処理によるメモリ効率化）
        const bodyData = processDataInChunks(bodyValues, (rowData, rowIndex) => {
          const actualRow = rowIndex + 4;
          const processedRow = [];
          
          for (let col = 0; col < rowData.length; col++) {
            const actualCol = startCol + col;
            const value = processMergedCellValueOptimized(rowData[col], actualRow, actualCol, bodyMergedInfo, sheet);
            processedRow.push(value);
          }
          
          return processedRow;
        }, CONFIG.PERFORMANCE.CHUNK_SIZE);
        
        result.bodyData = bodyData;
      }
    }
    
    console.log(`📊 F列起点データ抽出完了（完全超高速版）:`);
    console.log(`  - ヘッダーデータ: ${result.headerData.length}行`);
    console.log(`  - ボディデータ: ${result.bodyData.length}行`);
    console.log(`  - 抽出列数: ${actualLastCol - startCol + 1}列`);
    
    // 結果をキャッシュに保存
    setPerformanceCache(cacheKey, result, CONFIG.PERFORMANCE.CACHE_TTL);
    
    return result;
    
  } catch (error) {
    console.log(`❌ F列起点データ抽出エラー: ${error.message}`);
    throw error;
  }
}

/**
 * 結合セル情報を範囲指定で取得（高速版）
 * @param {Array} mergedRanges - 結合セル範囲の配列
 * @param {number} startRow - 開始行
 * @param {number} endRow - 終了行
 * @param {number} startCol - 開始列
 * @param {number} endCol - 終了列
 * @returns {Object} 結合セル情報
 */
function getMergedCellInfoForRange(mergedRanges, startRow, endRow, startCol, endCol) {
  try {
    const mergedInfo = {};
    
    for (let i = 0; i < mergedRanges.length; i++) {
      const range = mergedRanges[i];
      const rangeStartRow = range.getRow();
      const rangeEndRow = rangeStartRow + range.getNumRows() - 1;
      const rangeStartCol = range.getColumn();
      const rangeEndCol = rangeStartCol + range.getNumColumns() - 1;
      
      // 指定範囲内の結合セルのみ処理
      if (rangeStartRow <= endRow && rangeEndRow >= startRow &&
          rangeStartCol <= endCol && rangeEndCol >= startCol) {
        
        for (let row = rangeStartRow; row <= rangeEndRow; row++) {
          for (let col = rangeStartCol; col <= rangeEndCol; col++) {
            const key = `${row}_${col}`;
            mergedInfo[key] = {
              startRow: rangeStartRow,
              startCol: rangeStartCol,
              numRows: range.getNumRows(),
              numCols: range.getNumColumns(),
              isTopLeft: (row === rangeStartRow && col === rangeStartCol)
            };
          }
        }
      }
    }
    
    return mergedInfo;
    
  } catch (error) {
    console.log(`❌ 結合セル情報取得エラー: ${error.message}`);
    return {};
  }
}

/**
 * 結合セル値を処理（高速版・取り消し線除外対応）
 * @param {*} value - セルの値
 * @param {number} row - 行番号
 * @param {number} col - 列番号
 * @param {Object} mergedInfo - 結合セル情報
 * @param {Sheet} sheet - スプレッドシート（取り消し線チェック用）
 * @returns {*} 処理された値
 */
function processMergedCellValueOptimized(value, row, col, mergedInfo, sheet = null) {
  try {
    const key = `${row}_${col}`;
    const mergeData = mergedInfo[key];
    
    if (mergeData && !mergeData.isTopLeft) {
      // 結合セルの左上以外は空文字を返す
      return '';
    }
    
    // 取り消し線除外処理を適用
    if (sheet && value && value.toString().trim() !== '') {
      const cleanValue = checkAndCleanStrikethroughCell(sheet, row, col);
      return cleanValue;
    }
    
    return value || '';
    
  } catch (error) {
    return value || '';
  }
}

/**
 * 結合セル値を取得（高速版・取り消し線除外対応）
 * @param {Sheet} sheet - スプレッドシート
 * @param {number} row - 行番号
 * @param {number} col - 列番号
 * @returns {*} セルの値
 */
function getMergedCellValueWithMergeInfoOptimized(sheet, row, col) {
  try {
    // キャッシュ機能による高速化
    const cacheKey = `merged_cell_${sheet.getSheetId()}_${row}_${col}`;
    const cachedValue = getPerformanceCache(cacheKey);
    if (cachedValue && isPerformanceCacheValid(cacheKey)) {
      return cachedValue.data;
    }
    
    // 結合セルのチェック
    if (isMergedCell(sheet, row, col)) {
      const mergedRanges = sheet.getRange(row, col).getMergedRanges();
      if (mergedRanges.length > 0) {
        const range = mergedRanges[0];
        const startRow = range.getRow();
        const startCol = range.getColumn();
        
        // 結合セルの左上のセルのみ値を取得
        if (row === startRow && col === startCol) {
          const value = checkAndCleanStrikethroughCell(sheet, row, col);
          setPerformanceCache(cacheKey, value, CONFIG.PERFORMANCE.CACHE_TTL);
          return value;
        } else {
          // 結合セルの左上以外は空文字
          setPerformanceCache(cacheKey, '', CONFIG.PERFORMANCE.CACHE_TTL);
          return '';
        }
      }
    }
    
    // 通常のセルの値を取得（取り消し線除外処理付き）
    const value = checkAndCleanStrikethroughCell(sheet, row, col);
    setPerformanceCache(cacheKey, value, CONFIG.PERFORMANCE.CACHE_TTL);
    return value;
    
  } catch (error) {
    console.log(`❌ 結合セル値取得エラー: ${error.message}`);
    return '';
  }
}

/**
 * 実際の最終列を取得（高速版）
 * @param {Sheet} sheet - スプレッドシート
 * @param {number} startCol - 開始列
 * @param {number} maxCol - 最大列
 * @returns {number} 実際の最終列
 */
function getActualLastColumnOptimized(sheet, startCol, maxCol) {
  try {
    // キャッシュ機能による高速化
    const cacheKey = `actual_last_col_${sheet.getSheetId()}_${startCol}_${maxCol}`;
    const cachedCol = getPerformanceCache(cacheKey);
    if (cachedCol && isPerformanceCacheValid(cacheKey)) {
      return cachedCol.data;
    }
    
    const lastRow = sheet.getLastRow();
    let actualLastCol = startCol;
    
    // バッチ処理による高速列チェック
    for (let col = startCol; col <= maxCol; col++) {
      if (hasDataInColumnOptimized(sheet, col, lastRow)) {
        actualLastCol = col;
      } else {
        // 空列が3回連続したら処理を停止
        let emptyCount = 0;
        for (let checkCol = col; checkCol <= maxCol && emptyCount < 3; checkCol++) {
          if (!hasDataInColumnOptimized(sheet, checkCol, lastRow)) {
            emptyCount++;
          } else {
            break;
          }
        }
        
        if (emptyCount >= 3) {
          break;
        }
      }
    }
    
    // 結果をキャッシュに保存
    setPerformanceCache(cacheKey, actualLastCol, CONFIG.PERFORMANCE.CACHE_TTL);
    
    return actualLastCol;
    
  } catch (error) {
    console.log(`❌ 実際の最終列取得エラー: ${error.message}`);
    return startCol;
  }
}

/**
 * 列にデータがあるかチェック（高速版・取り消し線除外対応）
 * @param {Sheet} sheet - スプレッドシート
 * @param {number} col - 列番号
 * @param {number} lastRow - 最終行
 * @returns {boolean} データの有無
 */
function hasDataInColumnOptimized(sheet, col, lastRow) {
  try {
    // キャッシュ機能による高速化
    const cacheKey = `has_data_col_${sheet.getSheetId()}_${col}_${lastRow}`;
    const cachedResult = getPerformanceCache(cacheKey);
    if (cachedResult !== null && isPerformanceCacheValid(cacheKey)) {
      return cachedResult.data;
    }
    
    // 4行目以降のデータをチェック（ヘッダー行は除外）
    let hasData = false;
    for (let row = 4; row <= lastRow; row++) {
      const value = checkAndCleanStrikethroughCell(sheet, row, col);
      if (value && value !== '') {
        hasData = true;
        break;
      }
    }
    
    // 結果をキャッシュに保存
    setPerformanceCache(cacheKey, hasData, CONFIG.PERFORMANCE.CACHE_TTL);
    
    return hasData;
    
  } catch (error) {
    return false;
  }
}

/**
 * バッチ処理によるデータ処理（高速版）
 * @param {Array} data - 2次元配列データ
 * @param {Function} processRow - 行データを処理する関数
 * @param {number} chunkSize - チャンクサイズ
 * @returns {Array} 処理済みのデータ
 */
function processDataInChunks(data, processRow, chunkSize) {
  const processedData = [];
  for (let i = 0; i < data.length; i += chunkSize) {
    const chunk = data.slice(i, i + chunkSize);
    const processedChunk = chunk.map(processRow);
    processedData.push(...processedChunk);
  }
  return processedData;
}

/**
 * 列データを情報抽出タブに出力（完全超高速版・バッチ処理）
 * @param {Object} extractedData - 抽出されたデータ
 */
function outputColumnDataToInfoExtractionTabFullyOptimized(extractedData) {
  try {
    if (!extractedData) {
      console.log(`⚠️ 出力データがありません`);
      return;
    }
    
    // 現在アクティブなスプレッドシートを取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      console.log('⚠️ アクティブなスプレッドシートが見つかりません');
      return;
    }
    
    const sheet = ss.getSheetByName(CONFIG.SHEETS.INFO_EXTRACTION);
    
    console.log(`📝 列データを情報抽出タブに出力開始（完全超高速版）`);
    
    // ヘッダーデータ（1-3行目）を4-6行目に出力（バッチ処理）
    if (extractedData.headerData && extractedData.headerData.length > 0) {
      console.log(`📝 ヘッダーデータ出力（完全超高速版）: 4-6行目 (D列から)`);
      
      // バッチ処理による高速出力
      const headerRange = sheet.getRange(4, 4, extractedData.headerData.length, extractedData.headerData[0].length);
      headerRange.setValues(extractedData.headerData);
      
      console.log(`✅ ヘッダーデータ出力完了: ${extractedData.headerData.length}行`);
    }
    
    // ボディデータ（4行目以降）を8行目以降に出力（バッチ処理）
    if (extractedData.bodyData && extractedData.bodyData.length > 0) {
      console.log(`📝 ボディデータ出力（完全超高速版）: 8行目以降 (D列から)`);
      
      // バッチ処理による高速出力
      const bodyRange = sheet.getRange(8, 4, extractedData.bodyData.length, extractedData.bodyData[0].length);
      bodyRange.setValues(extractedData.bodyData);
      
      console.log(`✅ ボディデータ出力完了: ${extractedData.bodyData.length}行`);
    }
    
    console.log(`✅ 列データ出力完了（完全超高速版）`);
    
  } catch (error) {
    console.log(`❌ 列データ出力エラー: ${error.message}`);
    throw error;
  }
}

/**
 * セルの書式情報を含めて取り消し線をチェック（Google Sheets用）
 * @param {Sheet} sheet - スプレッドシート
 * @param {number} row - 行番号
 * @param {number} col - 列番号
 * @returns {string} クリーンなテキスト（取り消し線は除外）
 */
function checkAndCleanStrikethroughCell(sheet, row, col) {
  try {
    const range = sheet.getRange(row, col);
    const value = range.getValue();
    
    if (!value || value.toString().trim().length === 0) {
      return "";
    }
    
    // Google Sheetsの書式情報を取得
    const textStyle = range.getTextStyle();
    
    // 取り消し線が設定されているかチェック
    if (textStyle && textStyle.isStrikethrough && textStyle.isStrikethrough()) {
      console.log(`🚫 取り消し線セル検出: ${getColumnLetter(col)}${row} = '${value}' (除外)`);
      return ""; // 取り消し線が設定されている場合は空文字を返す
    }
    
    // リッチテキストの場合の処理
    const richTextValue = range.getRichTextValue();
    if (richTextValue) {
      const runs = richTextValue.getRuns();
      let cleanText = "";
      
      for (let i = 0; i < runs.length; i++) {
        const run = runs[i];
        const runText = run.getText();
        const runStyle = run.getTextStyle();
        
        // この部分が取り消し線でない場合のみ追加
        if (!runStyle || !runStyle.isStrikethrough || !runStyle.isStrikethrough()) {
          cleanText += runText;
        } else {
          console.log(`🚫 取り消し線テキスト検出: '${runText}' (除外)`);
        }
      }
      
      return removeStrikethroughText(cleanText);
    }
    
    // 通常のテキストの場合
    return removeStrikethroughText(value.toString());
    
  } catch (error) {
    console.log(`⚠️ セル書式チェックエラー: ${error.toString()}`);
    // エラーの場合は通常の取り消し線除去を実行
    return removeStrikethroughText(value ? value.toString() : "");
  }
}

/**
 * 取り消し線テキストを除去（強化版）
 * @param {string} text - 処理対象テキスト
 * @returns {string} クリーンなテキスト
 */
function removeStrikethroughText(text) {
  try {
    if (!text || text.toString().trim().length === 0) {
      return "";
    }
    
    let cleanText = text.toString();
    
    // マークダウン形式の取り消し線を除去
    cleanText = cleanText.replace(/~~(.+?)~~/g, '');
    
    // HTMLタグの取り消し線を除去
    cleanText = cleanText.replace(/<s>(.+?)<\/s>/gi, '');
    cleanText = cleanText.replace(/<strike>(.+?)<\/strike>/gi, '');
    cleanText = cleanText.replace(/<del>(.+?)<\/del>/gi, '');
    
    // Unicode取り消し線文字を除去
    cleanText = cleanText.replace(/[\u0336]/g, ''); // 取り消し線文字
    cleanText = cleanText.replace(/[\u0335]/g, ''); // 短い取り消し線
    
    // 特定の文字パターンを除去（例：[削除]、(削除)、※削除など）
    cleanText = cleanText.replace(/\[削除\]/g, '');
    cleanText = cleanText.replace(/\(削除\)/g, '');
    cleanText = cleanText.replace(/※削除/g, '');
    cleanText = cleanText.replace(/削除：/g, '');
    
    // 空白の正規化
    cleanText = cleanText.replace(/\s+/g, ' ').trim();
    
    return cleanText;
    
  } catch (error) {
    console.log(`⚠️ 取り消し線除去エラー: ${error.toString()}`);
    return text.toString();
  }
}

/**
 * 列番号を列文字に変換
 * @param {number} columnNumber - 列番号（1から開始）
 * @returns {string} 列文字（A, B, C...）
 */
function getColumnLetter(columnNumber) {
  let result = "";
  while (columnNumber > 0) {
    columnNumber--;
    result = String.fromCharCode(65 + (columnNumber % 26)) + result;
    columnNumber = Math.floor(columnNumber / 26);
  }
  return result;
}


