/**
 * カスタムメニューと simple triggers
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu(APP_NAME)
    .addItem('kintone取込', 'importKintone')
    .addSeparator()
    .addItem('案件登録(見積入力)', 'openFormProject')
    .addItem('確定値入力', 'openFormConfirm')
    .addSeparator()
    .addItem('現場分配を入力', 'openFormField')
    .addItem('制作分配を入力', 'openFormProd')
    .addItem('まとめ分配(複数案件へ一括)', 'openFormBulk')
    .addItem('分配取消', 'openFormCancel')
    .addSeparator()
    .addItem('ビュー更新', 'refreshAllViews')
    .addItem('再計算', 'menuRecalc')
    .addSubMenu(ui.createMenu('管理者')
      .addItem('初期セットアップ', 'setupAll')
      .addItem('保護を再設定', 'menuSetupProtections')
      .addItem('確定値修正', 'adminFixConfirmed')
      .addItem('請求月上書き修正', 'adminFixBillingMonth')
      .addItem('半期締め', 'adminCloseHalf'))
    .addToUi();

  // 個人ビューの初期選択(開いた人のメールから)
  try { presetPersonSelector_(); } catch (e) { /* 未セットアップ時は無視 */ }
}

function presetPersonSelector_() {
  const me = memberByEmail_(userEmail_());
  if (!me) return;
  const vp = ss_().getSheetByName(SHEET.V_PERSON);
  if (!vp) return;
  if (String(vp.getRange('B1').getValue()) === '') {
    vp.getRange('B1').setValue(me.name);
    buildPersonView_(me.name);
  }
}

/** セレクタ編集でビューを更新 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (e.range.getA1Notation() !== 'B1') return;
    if (sh.getName() === SHEET.V_PERSON) buildPersonView_(String(e.range.getValue()));
    if (sh.getName() === SHEET.V_UNIT) buildUnitView_(String(e.range.getValue()));
  } catch (err) {
    // onEditは静かに失敗させない(トーストのみ)
    try { toast_('ビュー更新エラー: ' + err.message); } catch (e2) { /* noop */ }
  }
}

function menuRecalc() {
  recalcAll_();
  toast_('再計算が完了しました');
}

function menuSetupProtections() {
  assertAdmin_();
  setupProtections();
  toast_('保護を再設定しました');
}

// ---- サイドバー起動 ----

function openSidebar_(file, title) {
  const html = HtmlService.createHtmlOutputFromFile(file).setTitle(title);
  SpreadsheetApp.getUi().showSidebar(html);
}

function openFormProject() { openSidebar_('f_project', '案件登録(見積入力)'); }
function openFormConfirm() { openSidebar_('f_confirm', '確定値入力'); }
function openFormField()   { openSidebar_('f_field', '現場分配'); }
function openFormProd()    { openSidebar_('f_prod', '制作分配'); }
function openFormBulk()    { openSidebar_('f_bulk', 'まとめ分配'); }
function openFormCancel()  { openSidebar_('f_cancel', '分配取消'); }
