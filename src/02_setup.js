/**
 * 初期セットアップ:全シート生成・マスタ初期値投入・保護・ビュー雛形
 * 管理者メニュー「初期セットアップ」から実行(再実行しても既存データは消さない)
 */

function setupAll() {
  const ss = ss_();
  const firstRun = !ss.getSheetByName(SHEET.MASTER);
  if (!firstRun) {
    // 2回目以降は管理者チェック(初回はマスタが無く管理者判定できないため実行者を管理者として登録)
    assertAdmin_();
  }

  setupMasterSheet_(firstRun);
  setupMemberSheet_();
  setupPasteSheet_();
  // 請求月・日付系の列はプレーンテキスト書式にする('2026/07' 等の文字列が
  // 日付値へ自動変換されるのを防ぎ、ヘッダー名マッチングや月比較を文字列で安定させる)
  setupHeaderSheet_(SHEET.PROJ, PROJ_COLS, false,
    ['請求月_自動', '納品予定日_始', '納品予定日_終', '最終取込日時']);
  setupHeaderSheet_(SHEET.PI, PI_COLS, false,
    ['請求月_自動', '請求月上書き', '適用請求月', '登録日時', '確定入力日時']);
  setupHeaderSheet_(SHEET.ALLOC, AL_COLS, false,
    ['請求月', '入力日時', '取消日時']);
  setupHeaderSheet_(SHEET.AGG, AGG_COLS, true, ['請求月']);
  setupHeaderSheet_(SHEET.AUDIT, AUDIT_COLS, true, ['日時']);
  setupViewSheets_();
  setupProtections();
  removeDefaultSheet_();

  toast_('初期セットアップが完了しました。設定_メンバーにメンバーを登録してください。');
  audit_('初期セットアップ', '', '', '', '', firstRun ? '初回' : '再実行');
}

function getOrCreateSheet_(name) {
  const ss = ss_();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function removeDefaultSheet_() {
  const ss = ss_();
  const def = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);
}

function setHeaderRow_(sh, col, headers) {
  sh.getRange(1, col, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#efefef');
}

/** ヘッダーのみのシート(任意で非表示・指定列をテキスト書式に) */
function setupHeaderSheet_(name, cols, hidden, textColNames) {
  const sh = getOrCreateSheet_(name);
  setHeaderRow_(sh, 1, cols); // ヘッダーは常に正に上書き(データ行は触らない)
  sh.setFrozenRows(1);
  (textColNames || []).forEach(function (n) {
    sh.getRange(2, CI(cols, n) + 1, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('@');
  });
  if (hidden) sh.hideSheet();
  return sh;
}

// ================= 設定_マスタ =================

function setupMasterSheet_(firstRun) {
  const sh = getOrCreateSheet_(SHEET.MASTER);

  writeBlock_(sh, M_BLOCK.GRADE, '■等級マスタ', firstRun ? [
    ['b', 1], ['a1', 2], ['a2', 3], ['l1', 4],
  ] : null);

  writeBlock_(sh, M_BLOCK.TARGET, '■半期粗利目標(等級×年次)', firstRun ? [
    ['b', 1, 500000],
    ['b', 2, 1500000],
    ['b', 3, 2500000],
    ['a1', '', 5400000],
    ['a2', '', 6300000],
    ['l1', '', 8100000],
  ] : null);

  writeBlock_(sh, M_BLOCK.POSITION, '■ポジション×単価表', firstRun ? [
    ['ディレクター', 'a2', 50000, 35000],
    ['アシスタントディレクター', 'a1', 35000, 25000],
    ['スタッフ', '', 25000, 20000],
  ] : null);

  writeBlock_(sh, M_BLOCK.MAPPING, '■kintone取込マッピング(ヘッダー名はここで変更可)', firstRun ? [
    [K.ID, '案件ID', '◎'],
    [K.NAME, '案件名', '◎'],
    [K.CLIENT, '取引先名', '◎'],
    [K.DUE_END, '納品予定日_終', '◎'],
    [K.PHASE, 'フェーズ', '◎'],
    [K.OWNER, '案件所有者', '○'],
    [K.DUE_START, '納品予定日_始', '○'],
    [K.DAYS, '日数', '○'],
    [K.PARENT_ID, '親案件ID', '○'],
    [K.PARENT_NAME, '親案件名', '○'],
  ] : null);

  writeBlock_(sh, M_BLOCK.SETTING, '■設定', firstRun ? [
    ['期初月', 7],
    ['ユニット長個人目標', 5000000],
  ] : null);

  writeBlock_(sh, M_BLOCK.ADMIN, '■管理者', firstRun ? [
    [userEmail_() || '(セットアップ実行者のメールを記入)'],
  ] : null);

  writeBlock_(sh, M_BLOCK.CLOSED, '■締め済み半期(半期締めで自動追記)', null);

  sh.setFrozenRows(2);
  MASTER_CACHE_ = null;
}

/** ブロック書込:1行目セクション名、2行目ヘッダー。dataRows が null なら見出しのみ整える */
function writeBlock_(sh, block, title, dataRows) {
  sh.getRange(1, block.col).setValue(title).setFontWeight('bold');
  sh.getRange(2, block.col, 1, block.headers.length).setValues([block.headers])
    .setFontWeight('bold').setBackground('#efefef');
  if (dataRows && dataRows.length) {
    sh.getRange(3, block.col, dataRows.length, block.headers.length).setValues(dataRows);
  }
}

// ================= 設定_メンバー =================

function setupMemberSheet_() {
  const sh = getOrCreateSheet_(SHEET.MEMBER);
  setHeaderRow_(sh, 1, MEMBER_COLS);
  sh.setFrozenRows(1);

  // 入力規則:等級はマスタ参照、評価対象/ユニット長はチェックボックス
  const gradeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['b', 'a1', 'a2', 'l1'], true).setAllowInvalid(true).build();
  sh.getRange(2, CI(MEMBER_COLS, '現等級') + 1, 200, 1).setDataValidation(gradeRule);
  sh.getRange(2, CI(MEMBER_COLS, '評価対象') + 1, 200, 1).insertCheckboxes();
  sh.getRange(2, CI(MEMBER_COLS, 'ユニット長') + 1, 200, 1).insertCheckboxes();
}

// ================= kintone貼付 =================

function setupPasteSheet_() {
  const sh = getOrCreateSheet_(SHEET.PASTE);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1).setValue('ここにkintone CSVの内容を1行目(ヘッダー)ごと貼り付けて、メニュー「' + APP_NAME + ' > kintone取込」を実行');
    sh.getRange(1, 1).setFontColor('#999999');
  }
}

// ================= ビュー雛形 =================

function setupViewSheets_() {
  // V_個人
  const vp = getOrCreateSheet_(SHEET.V_PERSON);
  vp.getRange('A1').setValue('メンバー選択▼').setFontWeight('bold');
  setMemberSelector_(vp.getRange('B1'));
  vp.getRange('D1').setValue('セレクタ変更で自動更新/メニュー「ビュー更新」でも更新できます').setFontColor('#999999');
  vp.setColumnWidth(2, 160);

  // V_ユニット
  const vu = getOrCreateSheet_(SHEET.V_UNIT);
  vu.getRange('A1').setValue('ユニット選択▼').setFontWeight('bold');
  setUnitSelector_(vu.getRange('B1'));
  vu.getRange('D1').setValue('セレクタ変更で自動更新').setFontColor('#999999');
  vu.setColumnWidth(2, 160);

  // V_経営
  const vm = getOrCreateSheet_(SHEET.V_MGMT);
  vm.getRange('A1').setValue('経営ビュー(全社+ユニット別)').setFontWeight('bold');
}

function setMemberSelector_(range) {
  let names = [];
  try { names = members_().map(function (m) { return m.name; }); } catch (e) { /* メンバー未登録 */ }
  if (names.length) {
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(names, true).setAllowInvalid(false).build();
    range.setDataValidation(rule);
  }
}

function setUnitSelector_(range) {
  let units = [];
  try {
    members_().forEach(function (m) { if (m.unit && units.indexOf(m.unit) < 0) units.push(m.unit); });
  } catch (e) { /* noop */ }
  if (units.length) {
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(units, true).setAllowInvalid(false).build();
    range.setDataValidation(rule);
  }
}

// ================= 保護 =================

/**
 * シート保護の方針(README参照):
 * - 設定_マスタ / 設定_メンバー:管理者のみ編集可(ハード保護)
 * - GASが書き込むシート(案件データ/案件入力/分配入力/集計/監査ログ/ビュー):警告付き保護
 *   ※コンテナバインドのGASは「実行した本人の権限」で動くため、メンバーのフォーム送信を
 *     通すにはハード保護にできない。誤編集は警告で防ぎ、改ざんは監査ログ+再計算で検知する。
 * - kintone貼付:保護なし(全員貼付可)
 */
function setupProtections() {
  const adminEmails = masters_().admins.filter(function (e) { return e.indexOf('@') > 0; });

  // 既存のシート保護をいったん整理して張り直す
  [SHEET.MASTER, SHEET.MEMBER, SHEET.PROJ, SHEET.PI, SHEET.ALLOC, SHEET.AGG, SHEET.AUDIT,
   SHEET.V_PERSON, SHEET.V_UNIT, SHEET.V_MGMT].forEach(function (name) {
    const sh = ss_().getSheetByName(name);
    if (!sh) return;
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) { p.remove(); });
  });

  // 管理者のみ(ハード保護)
  [SHEET.MASTER, SHEET.MEMBER].forEach(function (name) {
    const sh = ss_().getSheetByName(name);
    if (!sh) return;
    const p = sh.protect().setDescription('管理者のみ編集可');
    if (adminEmails.length) {
      p.removeEditors(p.getEditors().map(function (u) { return u.getEmail(); }).filter(function (e) {
        return adminEmails.indexOf(e.toLowerCase()) < 0;
      }));
      try { p.addEditors(adminEmails); } catch (e) { /* ドメイン外等 */ }
    }
  });

  // 警告付き保護(直接編集時に警告ダイアログ)
  [SHEET.PROJ, SHEET.PI, SHEET.ALLOC, SHEET.AGG, SHEET.AUDIT, SHEET.V_MGMT].forEach(function (name) {
    const sh = ss_().getSheetByName(name);
    if (!sh) return;
    sh.protect().setDescription('GASのみ書込(直接編集禁止)').setWarningOnly(true);
  });

  // ビュー(セレクタB1のみ警告なしで編集可)
  [SHEET.V_PERSON, SHEET.V_UNIT].forEach(function (name) {
    const sh = ss_().getSheetByName(name);
    if (!sh) return;
    const p = sh.protect().setDescription('閲覧用(セレクタのみ編集可)').setWarningOnly(true);
    p.setUnprotectedRanges([sh.getRange('B1')]);
  });
}
