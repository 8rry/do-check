/**
 * 設定定数
 * プロジェクト全体で使用する設定値を管理
 */

const CONFIG = {
  // スプレッドシートID
  SPREADSHEET_ID: '1W-Kmre4FTL5iU0VNSs5Z4vLVsXzFebLYlxSxnPWkxPQ',
  
  // シート名
  SHEETS: {
    INFO_EXTRACTION: '情報抽出',
    DO_EXPORT: 'Do書き出し用',
    DO_EXPORT_REGULAR: 'Do書き出し用(定期)',
    MUNICIPALITY_FOLDERS: '自治体フォルダ'
  },
  
  // セル位置
  CELLS: {
    FOLDER_PATH: 'B1',        // 返礼品シートのファイルパス
    COLUMN_SPEC: 'B2',        // 抽出する列の指定
    MAPPING_START_ROW: 7      // マッピング開始行
  },
  
  // 出力設定
  OUTPUT: {
    START_ROW: 8,             // 出力開始行（B8、C8から）
    COL_B: 2,                 // B列（2列目）
    COL_C: 3                  // C列（3列目）
  },
  
  // パフォーマンス設定
  PERFORMANCE: {
    LOG_DETAIL: true,         // 詳細ログの出力
    BATCH_SIZE: 100           // バッチ処理サイズ
  }
};
