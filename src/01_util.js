/**
 * 共通ユーティリティ:シートアクセス・マスタ読込・日付/半期計算・監査ログ
 */

function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('シートが見つかりません: ' + name + '(管理者メニュー>初期セットアップ を実行してください)');
  return sh;
}

function toast_(msg, title) {
  ss_().toast(msg, title || APP_NAME, 8);
}

function userEmail_() {
  return (Session.getActiveUser().getEmail() || '').toLowerCase();
}

function now_() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
}

/** ヘッダー行(1行目)+データを読む。戻り値: {headers, rows(2次元), sh} */
function readSheet_(name) {
  const sh = sheet_(name);
  const values = sh.getDataRange().getValues();
  return { sh: sh, headers: values[0] || [], rows: values.slice(1) };
}

/** Date/文字列/数値 を Date に変換(不可なら null) */
function asDate_(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === 'string' && v.trim()) {
    const m = v.trim().match(/^(\d{4})[\/\-](\d{1,2})(?:[\/\-](\d{1,2}))?/);
    if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3] || 1));
  }
  return null;
}

/** Date → 'yyyy/MM'(請求月表現) */
function ymOf_(v) {
  const d = asDate_(v);
  if (!d) return '';
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM');
}

/** 'yyyy/MM' 形式の正規化(セルにDateが入っていても/文字列でも) */
function normYm_(v) {
  if (v instanceof Date) return ymOf_(v);
  const s = String(v || '').trim();
  const m = s.match(/^(\d{4})[\/\-](\d{1,2})/);
  if (m) return m[1] + '/' + ('0' + m[2]).slice(-2);
  return '';
}

/**
 * 半期ラベル。期初=7月(設定_マスタの「期初月」)。
 * 2026/07〜2026/12 → '2026前期'、2027/01〜2027/06 → '2026後期'
 */
function halfOfYm_(ym, startMonth) {
  if (!ym) return '';
  const start = startMonth || getSettingNum_('期初月', 7);
  const parts = ym.split('/');
  const y = Number(parts[0]);
  const m = Number(parts[1]);
  if (m >= start) return y + '前期';
  return (y - 1) + '後期';
}

/** 半期ラベル → その半期に含まれる 'yyyy/MM' 6ヶ月分 */
function monthsOfHalf_(label) {
  const m = String(label).match(/^(\d{4})(前期|後期)$/);
  if (!m) return [];
  const y = Number(m[1]);
  const start = getSettingNum_('期初月', 7);
  const months = [];
  for (let i = 0; i < 6; i++) {
    let mm = (m[2] === '前期') ? start + i : start + 6 + i;
    let yy = y;
    while (mm > 12) { mm -= 12; yy += 1; }
    months.push(yy + '/' + ('0' + mm).slice(-2));
  }
  return months;
}

function currentHalf_() {
  return halfOfYm_(ymOf_(new Date()));
}

/** 半期ラベルの並び替えキー('2026前期'→202607) */
function halfSortKey_(label) {
  const m = String(label).match(/^(\d{4})(前期|後期)$/);
  if (!m) return 0;
  return Number(m[1]) * 100 + (m[2] === '前期' ? getSettingNum_('期初月', 7) : getSettingNum_('期初月', 7) + 6);
}

// ================= 設定_マスタ 読み出し =================

/** マスタブロックを {headers, rows} で読む(セクション名行はスキップ) */
function readBlock_(block) {
  const sh = sheet_(SHEET.MASTER);
  const last = Math.max(sh.getLastRow(), 3);
  const values = sh.getRange(2, block.col, last - 1, block.headers.length).getValues();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i].every(function (v) { return v === '' || v === null; })) continue;
    rows.push(values[i]);
  }
  return rows;
}

let MASTER_CACHE_ = null;

/** マスタ一括読込(1実行内キャッシュ) */
function masters_() {
  if (MASTER_CACHE_) return MASTER_CACHE_;
  const grades = {};   // 等級 → 順位
  readBlock_(M_BLOCK.GRADE).forEach(function (r) { if (r[0] !== '') grades[String(r[0])] = Number(r[1]); });

  const targets = readBlock_(M_BLOCK.TARGET).map(function (r) {
    return { grade: String(r[0]), nenji: r[1] === '' ? null : Number(r[1]), amount: Number(r[2]) };
  });

  const positions = {};  // ポジション → {minGrade, day, half}
  readBlock_(M_BLOCK.POSITION).forEach(function (r) {
    if (r[0] !== '') positions[String(r[0])] = { minGrade: String(r[1] || ''), day: Number(r[2]), half: Number(r[3]) };
  });

  const mapping = readBlock_(M_BLOCK.MAPPING).map(function (r) {
    return { logical: String(r[0]), header: String(r[1]), required: String(r[2]) === '◎' };
  });

  const settings = {};
  readBlock_(M_BLOCK.SETTING).forEach(function (r) { if (r[0] !== '') settings[String(r[0])] = r[1]; });

  const admins = readBlock_(M_BLOCK.ADMIN).map(function (r) { return String(r[0]).toLowerCase(); }).filter(String);
  const closedHalves = readBlock_(M_BLOCK.CLOSED).map(function (r) { return String(r[0]); }).filter(String);

  MASTER_CACHE_ = { grades: grades, targets: targets, positions: positions, mapping: mapping, settings: settings, admins: admins, closedHalves: closedHalves };
  return MASTER_CACHE_;
}

function getSettingNum_(key, defVal) {
  try {
    const v = masters_().settings[key];
    return (v === undefined || v === '') ? defVal : Number(v);
  } catch (e) {
    return defVal;
  }
}

function isAdmin_(email) {
  return masters_().admins.indexOf((email || userEmail_()).toLowerCase()) >= 0;
}

function assertAdmin_() {
  if (!isAdmin_()) throw new Error('この操作は管理者のみ実行できます(設定_マスタの管理者メール参照)');
}

function isClosedHalf_(halfLabel) {
  return masters_().closedHalves.indexOf(halfLabel) >= 0;
}

/** 締め済み半期チェック(書込系の共通ガード) */
function assertNotClosed_(ym) {
  const half = halfOfYm_(normYm_(ym));
  if (half && isClosedHalf_(half)) {
    throw new Error('半期「' + half + '」は締め済みのため変更できません(管理者にご相談ください)');
  }
}

// ================= 設定_メンバー =================

let MEMBER_CACHE_ = null;

/** メンバー一覧。{name, email, unit, grade, nenji, evalTarget, evalStart, isLead} */
function members_() {
  if (MEMBER_CACHE_) return MEMBER_CACHE_;
  const data = readSheet_(SHEET.MEMBER);
  MEMBER_CACHE_ = data.rows.filter(function (r) { return String(r[0]).trim() !== ''; }).map(function (r) {
    return {
      name: String(r[CI(MEMBER_COLS, '氏名')]).trim(),
      email: String(r[CI(MEMBER_COLS, 'メール')]).trim().toLowerCase(),
      unit: String(r[CI(MEMBER_COLS, 'ユニット')]).trim(),
      grade: String(r[CI(MEMBER_COLS, '現等級')]).trim(),
      nenji: Number(r[CI(MEMBER_COLS, '年次')]) || null,
      evalTarget: r[CI(MEMBER_COLS, '評価対象')] === true || String(r[CI(MEMBER_COLS, '評価対象')]).toUpperCase() === 'TRUE',
      evalStart: r[CI(MEMBER_COLS, '評価開始日')],
      isLead: r[CI(MEMBER_COLS, 'ユニット長')] === true || String(r[CI(MEMBER_COLS, 'ユニット長')]).toUpperCase() === 'TRUE',
    };
  });
  return MEMBER_CACHE_;
}

function memberByEmail_(email) {
  const e = (email || '').toLowerCase().trim();
  return members_().filter(function (m) { return m.email === e; })[0] || null;
}

function memberByName_(name) {
  const n = String(name || '').trim();
  return members_().filter(function (m) { return m.name === n; })[0] || null;
}

/** メンバーの半期目標額(円)。ユニット長は設定「ユニット長個人目標」を優先 */
function targetOf_(member) {
  if (!member) return 0;
  if (member.isLead) return getSettingNum_('ユニット長個人目標', 5000000);
  const ts = masters_().targets;
  // 年次条件あり(b等級)を優先して探す
  let hit = ts.filter(function (t) { return t.grade === member.grade && t.nenji !== null && t.nenji === member.nenji; })[0];
  if (!hit) hit = ts.filter(function (t) { return t.grade === member.grade && t.nenji === null; })[0];
  if (!hit) {
    // b等級で年次がマスタの上限超なら最大年次の値を使う
    const sameGrade = ts.filter(function (t) { return t.grade === member.grade && t.nenji !== null; });
    if (sameGrade.length) {
      sameGrade.sort(function (a, b) { return b.nenji - a.nenji; });
      hit = sameGrade[0];
    }
  }
  return hit ? hit.amount : 0;
}

// ================= 監査ログ =================

/** 監査ログ追記(非表示シート・行削除しない) */
function audit_(opType, projId, target, before, after, detail) {
  const sh = sheet_(SHEET.AUDIT);
  sh.appendRow([now_(), userEmail_(), opType, projId || '', target || '',
    typeof before === 'object' ? JSON.stringify(before) : (before === undefined ? '' : before),
    typeof after === 'object' ? JSON.stringify(after) : (after === undefined ? '' : after),
    detail || '']);
}

// ================= 汎用 =================

function newId_(prefix) {
  return prefix + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss') + '-' + Math.floor(Math.random() * 1000);
}

function roundYen_(v) {
  return Math.round(Number(v) || 0);
}

/** 0.5刻みチェック */
function isHalfStep_(v) {
  const n = Number(v);
  return isFinite(n) && n > 0 && Math.abs(n * 2 - Math.round(n * 2)) < 1e-9;
}
