
// ---- シート初期化(手動セットアップ:マスタ/メンバー/各シート) ----
function seed() {
  ['設定_マスタ','設定_メンバー','kintone貼付','案件データ','案件入力','分配入力','集計','監査ログ','V_個人','V_ユニット','V_経営']
    .forEach(n => SHEETS[n] = new MockSheet(n));
  const m = SHEETS['設定_マスタ'];
  const put = (row, col, vals) => vals.forEach((v, i) => {
    while (m.data.length < row) m.data.push([]);
    m.data[row - 1][col - 1 + i] = v;
  });
  // GRADE(A) TARGET(D) POSITION(H) MAPPING(M) SETTING(Q) ADMIN(T) CLOSED(V)
  put(2, 1, ['等級','順位']); put(3,1,['b',1]); put(4,1,['a1',2]); put(5,1,['a2',3]); put(6,1,['l1',4]);
  put(2, 4, ['等級','年次','半期目標額']);
  put(3,4,['b',1,500000]); put(4,4,['b',2,1500000]); put(5,4,['b',3,2500000]);
  put(6,4,['a1','',5400000]); put(7,4,['a2','',6300000]); put(8,4,['l1','',8100000]);
  put(2, 8, ['ポジション','最低等級','単価_1日','単価_半日']);
  put(3,8,['ディレクター','a2',50000,35000]); put(4,8,['アシスタントディレクター','a1',35000,25000]); put(5,8,['スタッフ','',25000,20000]);
  put(2,13,['論理名','kintoneヘッダー名','必須']);
  [['案件ID','案件ID','◎'],['案件名','案件名','◎'],['取引先名','取引先名','◎'],['納品予定日_終','納品予定日_終','◎'],
   ['フェーズ','フェーズ','◎'],['案件所有者','案件所有者','○'],['納品予定日_始','納品予定日_始','○'],
   ['日数','日数','○'],['親案件ID','親案件ID','○'],['親案件名','親案件名','○']]
    .forEach((r,i)=>put(3+i,13,r));
  put(2,17,['キー','値']); put(3,17,['期初月',7]); put(4,17,['ユニット長個人目標',5000000]);
  put(2,20,['管理者メール']); put(3,20,['admin@example.com']);
  put(2,22,['締め済み半期']);

  const mem = SHEETS['設定_メンバー'];
  mem.appendRow(MEMBER_COLS);
  mem.appendRow(['山田','member01@example.com','ユニット1','a2',5,true,'','']);
  mem.appendRow(['佐藤','member02@example.com','ユニット1','a1',3,true,'','']);
  mem.appendRow(['鈴木','member03@example.com','ユニット2','b',2,true,'','']);
  SHEETS['案件データ'].appendRow(PROJ_COLS);
  SHEETS['案件入力'].appendRow(PI_COLS);
  SHEETS['分配入力'].appendRow(AL_COLS);
  SHEETS['集計'].appendRow(AGG_COLS);
  SHEETS['監査ログ'].appendRow(AUDIT_COLS);
}
seed();

let pass = 0, failCnt = 0;
function check(label, actual, expected) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected);
  if (ok) { pass++; console.log('  ✓', label); }
  else { failCnt++; console.log('  ✗', label, '\n    expected:', JSON.stringify(expected), '\n    actual:  ', JSON.stringify(actual)); }
}
function expectError(label, fn, msgPart) {
  try { fn(); failCnt++; console.log('  ✗', label, '(エラーになりませんでした)'); }
  catch (e) {
    if (String(e.message).includes(msgPart)) { pass++; console.log('  ✓', label); }
    else { failCnt++; console.log('  ✗', label, '別のエラー:', e.message); }
  }
}
const piVal = (id, col) => {
  const sh = SHEETS['案件入力'];
  const r = sh.data.find(r => String(r[0]) === id);
  return r ? r[CI(PI_COLS, col)] : undefined;
};
const alRows = () => SHEETS['分配入力'].data.slice(1);

// ===== 1. kintone取込(列順をわざと実CSVと変えてヘッダーマッチングを確認) =====
console.log('1. kintone取込');
const paste = SHEETS['kintone貼付'];
paste.data = [
  ['フェーズ','余計な列','案件名','取引先名','案件所有者','納品予定日_終','納品予定日_始','日数','親案件ID','親案件名','案件ID'],
  ['受注','x','イベントA','クライアントX','member01@example.com','2026/06/01','2026/05/28',5,'','','1061'],
  ['受注','x','イベントB','クライアントY','member03@example.com\nmember04@example.com','2026/07/02','2026/03/31',2,'','','940'],
  ['失注','x','失注案件','クライアントZ','member01@example.com','2026/08/01','',1,'','','999'],
  ['受注','x','協賛A','主催Q','member01@example.com','2026/09/10','',1,'54','シナハロ2026','2001'],
  ['受注','x','協賛B','主催Q','member01@example.com','2026/10/15','',1,'54','シナハロ2026','2002'],
];
importKintone();
const proj = SHEETS['案件データ'].data;
check('upsert 5件', proj.length - 1, 5);
const p940 = proj.find(r => r[0] === '940');
check('請求月_自動 = 納品予定日_終の月', p940[CI(PROJ_COLS,'請求月_自動')], '2026/07');
check('所有者複数 → 1行目採用', p940[CI(PROJ_COLS,'案件所有者メール')], 'member03@example.com');
check('所有者複数 → 警告', p940[CI(PROJ_COLS,'所有者警告')].includes('複数'), true);
check('所有者照合 → 氏名', p940[CI(PROJ_COLS,'所有者メンバー名')], '鈴木');

// 再取込(更新)
paste.data[1][2] = 'イベントA改';
importKintone();
check('再取込で更新(行重複しない)', SHEETS['案件データ'].data.length - 1, 5);
check('案件名が更新される', SHEETS['案件データ'].data.find(r => r[0]==='1061')[CI(PROJ_COLS,'案件名')], 'イベントA改');

// 必須ヘッダー欠落
const saved = paste.data.map(r=>r.slice());
paste.data = paste.data.map(r => r.slice(0, 10)); // 案件ID列を落とす
expectError('必須ヘッダー欠落で中断', () => importKintone(), '案件ID');
paste.data = saved;

// ===== 2. 案件登録 =====
console.log('2. 案件登録(見積入力)');
const initP = fsGetInit('project');
check('失注は選択肢から除外', initP.projects.some(p => p.id === '999'), false);
check('owner初期値', initP.projects.find(p=>p.id==='1061').ownerName, '山田');
fsSubmitProject({ projId: '1061', mainName: '山田', ymOverride: '', sales: 1000000, cost: 200000, gp: 800000 });
check('入力ステータス', piVal('1061','入力ステータス'), ST_WAIT_FIX);
check('適用請求月', piVal('1061','適用請求月'), '2026/06');
check('半期(6月=後期)', piVal('1061','半期'), '2025後期');
check('メイン担当自動行が生成され100%', alRows().filter(r => r[CI(AL_COLS,'自動行')] === true && String(r[1])==='1061')[0][CI(AL_COLS,'制作%')], 100);
expectError('二重登録ブロック', () => fsSubmitProject({ projId: '1061', mainName: '山田', sales: 1, cost: 0, gp: 1 }), '登録済み');
expectError('失注は登録不可', () => fsSubmitProject({ projId: '999', mainName: '山田', sales: 1, cost: 0, gp: 1 }), '失注');

// ===== 3. 現場分配 =====
console.log('3. 現場分配');
fsSubmitField({ projId: '1061', memberName: '佐藤', position: 'ディレクター', days: 2.5 });
const fRow = alRows().find(r => r[CI(AL_COLS,'種別')] === KIND_FIELD);
check('金額 = 2日×50000 + 半日35000', fRow[CI(AL_COLS,'金額')], 135000);
check('稼働時等級スナップショット', fRow[CI(AL_COLS,'稼働時等級')], 'a1');
check('現場は予測=確定', [fRow[CI(AL_COLS,'予測額')], fRow[CI(AL_COLS,'確定額')]], [135000, 135000]);
check('PI 現場稼働合計', piVal('1061','現場稼働合計'), 135000);
check('PI 制作原資_予測 = 800000-135000', piVal('1061','制作原資_予測'), 665000);
check('等級条件違反 警告(a1がディレクター)', String(piVal('1061','警告')).includes('等級条件違反'), true);
expectError('日数0.5刻み以外はエラー', () => fsSubmitField({ projId: '1061', memberName: '佐藤', position: 'スタッフ', days: 1.3 }), '0.5刻み');

// ===== 4. 制作分配 =====
console.log('4. 制作分配');
fsSubmitProd({ projId: '1061', items: [{ member: '鈴木', pct: 30 }] });
const prodRow = alRows().find(r => r[CI(AL_COLS,'種別')] === KIND_PROD && r[CI(AL_COLS,'メンバー')] === '鈴木');
check('鈴木 予測額 = 665000×30%', prodRow[CI(AL_COLS,'予測額')], 199500);
check('鈴木 確定額は未確定なので空', prodRow[CI(AL_COLS,'確定額')], '');
const autoRow = alRows().find(r => r[CI(AL_COLS,'自動行')] === true && String(r[1]) === '1061');
check('自動行 残り70%', autoRow[CI(AL_COLS,'制作%')], 70);
check('自動行 予測額 = 665000×70%', autoRow[CI(AL_COLS,'予測額')], 465500);
check('PI 制作%合計 = 100', piVal('1061','制作%合計'), 100);
expectError('メイン担当の%指定はエラー', () => fsSubmitProd({ projId: '1061', items: [{ member: '山田', pct: 10 }] }), '自動計算');
expectError('100%超ハードブロック', () => fsSubmitProd({ projId: '1061', items: [{ member: '佐藤', pct: 80 }] }), '100%を超え');
expectError('同一メンバー重複ブロック', () => fsSubmitProd({ projId: '1061', items: [{ member: '鈴木', pct: 10 }] }), '設定済み');

// ===== 5. 確定値入力 =====
console.log('5. 確定値入力');
fsSubmitConfirm({ projId: '1061', sales: 900000, cost: 150000, gp: 750000 });
check('ステータス確定済', piVal('1061','ステータス'), '確定済');
check('制作原資_確定 = 750000-135000', piVal('1061','制作原資_確定'), 615000);
check('見積確定差額 = -50000', piVal('1061','見積確定差額'), -50000);
const prodRow2 = alRows().find(r => r[CI(AL_COLS,'種別')] === KIND_PROD && r[CI(AL_COLS,'メンバー')] === '鈴木');
check('鈴木 確定額 = 615000×30%', prodRow2[CI(AL_COLS,'確定額')], 184500);
expectError('確定値の再入力ブロック', () => fsSubmitConfirm({ projId: '1061', sales: 1, cost: 0, gp: 1 }), '管理者メニュー');

// ===== 6. まとめ分配 =====
console.log('6. まとめ分配');
fsSubmitProject({ projId: '2001', mainName: '山田', ymOverride: '', sales: 300000, cost: 0, gp: 300000 });
fsSubmitProject({ projId: '2002', mainName: '山田', ymOverride: '', sales: 500000, cost: 100000, gp: 400000 });
const initB = fsGetInit('bulk');
check('親案件リスト', initB.parents, [{ id: '54', name: 'シナハロ2026' }]);
fsSubmitBulk({ projIds: ['2001', '2002'], items: [{ member: '佐藤', pct: 50 }] });
const bulkRows = alRows().filter(r => r[CI(AL_COLS,'まとめ分配ID')] !== '' && r[CI(AL_COLS,'まとめ分配ID')] !== undefined && String(r[CI(AL_COLS,'まとめ分配ID')]).startsWith('B'));
check('まとめ分配 2行生成・同一ID', [bulkRows.length, bulkRows[0][CI(AL_COLS,'まとめ分配ID')] === bulkRows[1][CI(AL_COLS,'まとめ分配ID')]], [2, true]);
check('案件ごとの原資×% (2001: 300000×50%)', bulkRows.find(r=>String(r[1])==='2001')[CI(AL_COLS,'予測額')], 150000);
check('案件ごとの原資×% (2002: 400000×50%)', bulkRows.find(r=>String(r[1])==='2002')[CI(AL_COLS,'予測額')], 200000);
check('請求月は案件ごと', [bulkRows.find(r=>String(r[1])==='2001')[CI(AL_COLS,'請求月')], bulkRows.find(r=>String(r[1])==='2002')[CI(AL_COLS,'請求月')]], ['2026/09','2026/10']);
expectError('まとめ分配:1件でもNGなら全体中断', () => fsSubmitBulk({ projIds: ['2001', '940'], items: [{ member: '佐藤', pct: 10 }] }), '未登録');

// ===== 7. 取消・赤字・集計 =====
console.log('7. 取消/赤字/集計');
expectError('確定済み案件の取消は管理者のみ', () => fsSubmitCancel({ allocIds: [prodRow2[0]] }), '管理者のみ');
const bulkId1 = bulkRows.find(r=>String(r[1])==='2001')[0];
fsSubmitCancel({ allocIds: [bulkId1] });
const cancelled = alRows().find(r => r[0] === bulkId1);
check('論理削除(無効フラグ)', cancelled[CI(AL_COLS,'無効')], true);
check('取消後 自動行が100%に戻る', alRows().find(r => r[CI(AL_COLS,'自動行')] === true && String(r[1])==='2001')[CI(AL_COLS,'制作%')], 100);

// 赤字:現場分配が粗利を超えるケース(2001: 粗利300000 に現場 16日スタッフ=400000)
fsSubmitField({ projId: '2001', memberName: '鈴木', position: 'スタッフ', days: 16 });
check('赤字でもマイナスのまま原資計算', piVal('2001','制作原資_予測'), -100000);
check('赤字警告', String(piVal('2001','警告')).includes('赤字'), true);
check('マイナスのまま自動行100%按分', alRows().find(r => r[CI(AL_COLS,'自動行')] === true && String(r[1])==='2001')[CI(AL_COLS,'予測額')], -100000);

// 集計
const agg = SHEETS['集計'].data.slice(1);
const yamada06 = agg.find(r => r[0] === '山田' && r[1] === '2026/06');
check('集計:山田 2026/06 確定 = 615000×70%', yamada06[CI(AGG_COLS,'確定額')], 430500);
const sato06 = agg.find(r => r[0] === '佐藤' && r[1] === '2026/06');
check('集計:佐藤 2026/06 確定(現場)', sato06[CI(AGG_COLS,'確定額')], 135000);

// ===== 8. ビュー描画(例外なく走ること) =====
console.log('8. ビュー描画');
try {
  buildPersonView_('山田');
  buildUnitView_('ユニット1');
  buildMgmtView_();
  pass++; console.log('  ✓ 3ビューとも例外なく描画');
} catch (e) { failCnt++; console.log('  ✗ ビュー描画でエラー:', e.message, e.stack); }
const vp = SHEETS['V_個人'].data;
check('V_個人にサマリ見出し', vp.some(r => String(r[0]).includes('半期サマリ')), true);

// ===== 9. 半期締めガード =====
console.log('9. 半期締めガード');
SHEETS['設定_マスタ'].data[2][21] = '2026前期'; // 締め済みに2026前期を直接追加(V列=22)
MASTER_CACHE_ = null;
expectError('締め済み半期はフォーム拒否', () => fsSubmitField({ projId: '2001', memberName: '鈴木', position: 'スタッフ', days: 1 }), '締め済み');

console.log('\n結果: ' + pass + ' passed, ' + failCnt + ' failed');
process.exit(failCnt ? 1 : 0);
