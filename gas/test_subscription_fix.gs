/**
 * 定期便処理のテスト関数
 * 修正内容の動作確認用
 */

/**
 * 定期便処理のテスト実行
 */
function testSubscriptionProcessing() {
  console.log('=== 定期便処理テスト開始 ===');
  
  // テストケース1: 2ヶ月定期便
  console.log('\n--- テストケース1: 2ヶ月定期便 ---');
  testSubscriptionCase('【2ヶ月定期便】いろはの自家製ぽん酢 360ｍl×各3本', 2);
  
  // テストケース2: 3回定期便
  console.log('\n--- テストケース2: 3回定期便 ---');
  testSubscriptionCase('【3回定期便】商品名', 3);
  
  // テストケース3: 6ヶ月定期便
  console.log('\n--- テストケース3: 6ヶ月定期便 ---');
  testSubscriptionCase('【6ヶ月定期便】商品名', 6);
  
  // テストケース4: 12回定期便
  console.log('\n--- テストケース4: 12回定期便 ---');
  testSubscriptionCase('【12回定期便】商品名', 12);
  
  console.log('\n=== 定期便処理テスト完了 ===');
}

/**
 * 個別テストケースの実行
 * @param {string} productName - 商品名称
 * @param {number} expectedCount - 期待される回数/月数
 */
function testSubscriptionCase(productName, expectedCount) {
  try {
    console.log(`入力: "${productName}"`);
    console.log(`期待: ${expectedCount}回/ヶ月`);
    
    // 定期便種別の判定テスト
    const subscriptionType = determineSubscriptionType(productName);
    console.log(`定期便種別: "${subscriptionType}"`);
    
    // 回数/月数の判定テスト
    let actualCount;
    if (subscriptionType === 'ヶ月定期便') {
      actualCount = determineSubscriptionMonths(productName);
      console.log(`月数判定: ${actualCount}ヶ月`);
    } else {
      actualCount = determineSubscriptionCount(productName);
      console.log(`回数判定: ${actualCount}回`);
    }
    
    const isSuccess = actualCount === expectedCount;
    console.log(`結果: ${actualCount}回/ヶ月`);
    console.log(`判定: ${isSuccess ? '✅ 成功' : '❌ 失敗'}`);
    
    if (!isSuccess) {
      console.log(`⚠️ 期待値と結果が一致しません`);
    }
    
  } catch (error) {
    console.log(`❌ テストエラー: ${error.message}`);
  }
}

/**
 * 子マスタ生成のテスト
 */
function testChildMasterGeneration() {
  console.log('=== 子マスタ生成テスト開始 ===');
  
  const testData = {
    '商品コード': '019-0339',
    '商品名称': '【2ヶ月定期便】いろはの自家製ぽん酢 360ｍl×各3本',
    '提供価格(税込)1': '3000',
    '寄附金額1': '5000'
  };
  
  console.log('元データ:', testData);
  
  // 子マスタ1の生成テスト
  console.log('\n--- 子マスタ1の生成 ---');
  const child1 = generateChildMaster(testData, 1);
  console.log('子マスタ1:', {
    '商品コード': child1['商品コード'],
    '商品名称': child1['商品名称'],
    '寄附金額1': child1['寄附金額1'],
    '提供価格(税込)1': child1['提供価格(税込)1']
  });
  
  // 子マスタ2の生成テスト
  console.log('\n--- 子マスタ2の生成 ---');
  const child2 = generateChildMaster(testData, 2);
  console.log('子マスタ2:', {
    '商品コード': child2['商品コード'],
    '商品名称': child2['商品名称'],
    '寄附金額1': child2['寄附金額1'],
    '提供価格(税込)1': child2['提供価格(税込)1']
  });
  
  console.log('\n=== 子マスタ生成テスト完了 ===');
}
