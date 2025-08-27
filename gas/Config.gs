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
  },
  
  // Phase 3: Do書き出し項目マッピングルール
  DO_MAPPING: {
    // 商品コード
    '商品コード': {
      keywords: ['返礼品コード'],
      newKeywords: ['新', '最新', '現在', '現行', '現在使用'],  // 新を示す部分一致キーワード
      oldKeywords: ['旧', '古い', '以前', '過去', '旧版', '旧式']   // 旧を示す部分一致キーワード
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
      keywords: ['発送温度帯']
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
      keywords: ['発送サイズ']
    },
    
    // 配送会社
    '配送会社': {
      keywords: ['発送', '発送方法']
    },
    
    // 寄附金額1
    '寄附金額1': {
      keywords: ['寄附額'],
      newKeywords: ['新', '最新', '現在', '現行', '現在使用'],  // 新を示す部分一致キーワード
      oldKeywords: ['旧', '古い', '以前', '過去', '旧版', '旧式']   // 旧を示す部分一致キーワード
    }
  }
};
