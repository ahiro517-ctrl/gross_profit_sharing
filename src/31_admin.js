/**
 * 管理者機能:確定値修正 / 請求月上書き修正 / 半期締め
 * すべて監査ログに変更前→変更後を記録する
 */

function adminFixConfirmed() {
  assertAdmin_();
  const ui = SpreadsheetApp.getUi();
  const projId = promptValue_(ui, '確定値修正', '対象の案件IDを入力してください');
  if (projId === null) return;

  const piSh = sheet_(SHEET.PI);
  const c = function (n) { return CI(PI_COLS, n); };
  const idx = findPiRow_(piSh, projId);
  if (idx < 0) { ui.alert('案件が登録されていません: ' + projId); return; }
  const row = piSh.getRange(idx, 1, 1, PI_COLS.length).getValues()[0];
  const before = { 売上: row[c('確定_売上')], 原価: row[c('確定_原価')], 粗利: row[c('確定_粗利')] };

  const sales = promptValue_(ui, '確定値修正', '確定_売上(現在: ' + before.売上 + ')');
  if (sales === null) return;
  const cost = promptValue_(ui, '確定値修正', '確定_原価(現在: ' + before.原価 + ')');
  if (cost === null) return;
  const gp = promptValue_(ui, '確定値修正', '確定_粗利(現在: ' + before.粗利 + ')');
  if (gp === null) return;

  const after = { 売上: requireNum_(sales, '確定_売上'), 原価: requireNum_(cost, '確定_原価'), 粗利: requireNum_(gp, '確定_粗利') };
  piSh.getRange(idx, c('確定_売上') + 1, 1, 3).setValues([[after.売上, after.原価, after.粗利]]);
  piSh.getRange(idx, c('確定入力者') + 1, 1, 2).setValues([[userEmail_() + '(修正)', now_()]]);

  audit_('確定値修正(管理者)', projId, '確定', before, after);
  recalcAll_();
  toast_('確定値を修正しました: ' + projId);
}

function adminFixBillingMonth() {
  assertAdmin_();
  const ui = SpreadsheetApp.getUi();
  const projId = promptValue_(ui, '請求月上書き修正', '対象の案件IDを入力してください');
  if (projId === null) return;

  const piSh = sheet_(SHEET.PI);
  const c = function (n) { return CI(PI_COLS, n); };
  const idx = findPiRow_(piSh, projId);
  if (idx < 0) { ui.alert('案件が登録されていません: ' + projId); return; }

  const before = piSh.getRange(idx, c('請求月上書き') + 1).getValue();
  const v = promptValue_(ui, '請求月上書き修正', '請求月(yyyy/MM)。空にすると自動導出値に戻ります(現在: ' + (before || '(空)') + ')');
  if (v === null) return;
  const ym = v.trim() === '' ? '' : normYm_(v);
  if (v.trim() !== '' && !ym) { ui.alert('yyyy/MM 形式で入力してください'); return; }

  piSh.getRange(idx, c('請求月上書き') + 1).setValue(ym);
  audit_('請求月上書き修正(管理者)', projId, '請求月上書き', before, ym);
  recalcAll_();
  toast_('請求月上書きを変更しました: ' + projId + ' → ' + (ym || '(自動)'));
}

/**
 * 半期締め:当該半期(請求月ベース)の全データを編集不可化する。
 * 締め済み半期はフォーム経由の登録・確定・分配・取消をすべて拒否する。
 */
function adminCloseHalf() {
  assertAdmin_();
  const ui = SpreadsheetApp.getUi();
  const label = promptValue_(ui, '半期締め', '締める半期を入力してください(例: 2026前期)');
  if (label === null) return;
  if (!/^\d{4}(前期|後期)$/.test(label.trim())) { ui.alert('「2026前期」「2026後期」の形式で入力してください'); return; }
  const half = label.trim();
  if (isClosedHalf_(half)) { ui.alert(half + ' は既に締め済みです'); return; }

  const ans = ui.alert('半期締め',
    half + ' を締めます。以後この半期の案件・分配はフォームから変更できなくなります。よろしいですか?',
    ui.ButtonSet.OK_CANCEL);
  if (ans !== ui.Button.OK) return;

  // 締め直前に最新化してから記録
  recalcAll_();
  const block = M_BLOCK.CLOSED;
  const sh = sheet_(SHEET.MASTER);
  const existing = readBlock_(block).length;
  sh.getRange(3 + existing, block.col).setValue(half);
  MASTER_CACHE_ = null;

  audit_('半期締め(管理者)', '', half, '', '締め実行');
  toast_(half + ' を締めました');
}

/** プロンプト入力。キャンセル時は null */
function promptValue_(ui, title, msg) {
  const res = ui.prompt(title, msg, ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return null;
  return res.getResponseText();
}
