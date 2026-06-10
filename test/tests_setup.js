console.log('setupAll スモークテスト');
setupAll();
const names = Object.keys(SHEETS);
['設定_マスタ','設定_メンバー','kintone貼付','案件データ','案件入力','分配入力','集計','V_個人','V_ユニット','V_経営','監査ログ']
  .forEach(n => console.log((names.includes(n) ? '  ✓ ' : '  ✗ MISSING ') + n));
console.log('  集計 非表示:', SHEETS['集計'].hidden, '/ 監査ログ 非表示:', SHEETS['監査ログ'].hidden);
// マスタ初期値が読めるか
MASTER_CACHE_ = null;
const m = masters_();
console.log('  等級:', JSON.stringify(m.grades));
console.log('  マッピング件数:', m.mapping.length, '/ 必須:', m.mapping.filter(x=>x.required).length);
console.log('  ポジション:', Object.keys(m.positions).join(','));
console.log('  管理者:', JSON.stringify(m.admins));
// 再実行(冪等性)— 管理者=実行者なので通るはず
setupAll();
console.log('  ✓ 再実行OK(冪等)');
