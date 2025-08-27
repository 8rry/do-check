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
    BATCH_SIZE: 100,          // バッチ処理サイズ
    CHUNK_SIZE: 50,           // チャンク処理サイズ
    CACHE_TTL: 5 * 60 * 1000, // キャッシュ有効期限（5分）
    PARALLEL_PROCESSING: true, // 並列処理の有効化
    MEMORY_OPTIMIZATION: true, // メモリ最適化の有効化
    ASYNC_PROCESSING: true,   // 非同期処理の有効化
    CACHE_ENABLED: true,      // キャッシュ機能の有効化
    PERFORMANCE_MONITORING: true, // パフォーマンス監視の有効化
    
    // Phase 3専用の最適化設定
    PHASE3_BATCH_SIZE: 50,    // Phase 3のバッチ処理サイズ
    PHASE3_CACHE_TTL: 10 * 60 * 1000, // Phase 3のキャッシュ有効期限（10分）
    PHASE3_MAPPING_CACHE: true, // Phase 3のマッピング結果キャッシュ
    PHASE3_SEARCH_CACHE: true,  // Phase 3の検索結果キャッシュ
    PHASE3_OUTPUT_BATCH: true   // Phase 3の出力バッチ処理
  },
  
  // Phase 3: Do書き出し項目マッピングルール
  DO_MAPPING: {
    // 商品コード
    '商品コード': {
      keywords: ['返礼品コード'],
      fallbackKeywords: ['返礼品コード(旧)', '返礼品コード(新)'],
      priority: 'new'  // 新を優先
    },
    
    // 集荷先名
    '集荷先名': {
      keywords: ['発送元', '名称']
    },
    
    // 商品名称
    '商品名称': {
      keywords: ['商品名']
    },
    
    // 配送伝票商品名称
    '配送伝票商品名称': {
      keywords: ['商品名', '伝票記載用']
    },
    
    // 発送種別
    '発送種別': {
      keywords: ['発送', '発送温度帯']
    },
    
    // 受付期間(開始)
    '受付期間(開始)': {
      keywords: ['受付開始']
    },
    
    // 受付期間(終了)
    '受付期間(終了)': {
      keywords: ['受付終了']
    },
    
    // 発送期間(開始)
    '発送期間(開始)': {
      keywords: ['発送開始']
    },
    
    // 発送期間(終了)
    '発送期間(終了)': {
      keywords: ['発送終了']
    },
    
    // 提供価格(税込)1
    '提供価格(税込)1': {
      keywords: ['商品代金']
    },
    
    // 固定送料1
    '固定送料1': {
      keywords: ['送料']
    },
    
    // 税率種別
    '税率種別': {
      keywords: ['税区分']
    },
    
    // 出荷時サイズ
    '出荷時サイズ': {
      keywords: ['発送', '発送サイズ']
    },
    
    // 配送会社
    '配送会社': {
      keywords: ['発送', '発送方法']
    },
    
    // 寄附金額1
    '寄附金額1': {
      keywords: ['寄附額'],
      fallbackKeywords: ['寄附額(旧)', '寄附額(新)'],
      priority: 'new'  // 新を優先
    }
  }
};
