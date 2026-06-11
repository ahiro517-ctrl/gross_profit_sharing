/**
 * kintone取込(仕様書§5.1)
 * - kintone貼付 シートの1行目ヘッダー名で列を特定(列位置に依存しない)
 * - マッピング表(設定_マスタ)の必要列のみ抽出し、案件IDをキーに案件データへ upsert
 * - 必須ヘッダー欠落時はどのヘッダーが無いかを明示して中断
 * - 取込ログをトースト+監査ログに記録
 */

function importKintone() {
  const paste = sheet_(SHEET.PASTE);
  const values = paste.getDataRange().getValues();
  if (values.length < 2) {
    throw uiError_('kintone貼付シートにデータがありません。CSVの内容をヘッダー行ごと貼り付けてください。');
  }

  // 1) ヘッダー名 → 列index の解決
  const headerRow = values[0].map(function (h) { return String(h).trim(); });
  const mapping = masters_().mapping; // [{logical, header, required}]
  const colOf = {};   // 論理名 → 列index
  const missing = [];
  mapping.forEach(function (m) {
    const idx = headerRow.indexOf(m.header);
    if (idx >= 0) {
      colOf[m.logical] = idx;
    } else if (m.required) {
      missing.push(m.header);
    }
  });
  if (missing.length) {
    throw uiError_('必須ヘッダーが見つからないため取込を中断しました:\n「' + missing.join('」「') + '」\n' +
      '貼り付けた1行目のヘッダー名、または 設定_マスタ のマッピング表を確認してください。');
  }

  // 2) 既存の案件データを読み込み(案件ID→行index)
  const projSh = sheet_(SHEET.PROJ);
  const projData = projSh.getDataRange().getValues();
  const idCol = CI(PROJ_COLS, '案件ID');
  const existing = {}; // 案件ID → 行index(0始まり、ヘッダー含むprojData上)
  for (let r = 1; r < projData.length; r++) {
    const id = String(projData[r][idCol]).trim();
    if (id) existing[id] = r;
  }

  // 3) 貼付データを案件データ行形式へ変換して upsert
  const ts = now_();
  let nNew = 0, nUpd = 0, nSkip = 0;
  const newRows = [];
  for (let r = 1; r < values.length; r++) {
    const src = values[r];
    const id = String(pick_(src, colOf, K.ID)).trim();
    if (!id) { nSkip++; continue; }

    const row = buildProjRow_(src, colOf, id, ts);
    if (existing[id] !== undefined) {
      projData[existing[id]] = row;
      nUpd++;
    } else {
      newRows.push(row);
      nNew++;
    }
  }

  // 4) 一括書込:請求月の新しい順に並べ直して書き戻す(請求月なしは末尾)
  //    下流はすべて案件IDで参照するため、並び順はいつ変わっても壊れない
  const ymCol = CI(PROJ_COLS, '請求月_自動');
  const all = projData.slice(1).concat(newRows).filter(function (r) {
    return String(r[idCol]).trim() !== '';
  });
  all.sort(function (x, y) {
    return String(y[ymCol]).localeCompare(String(x[ymCol]));
  });
  const lastRow = projSh.getLastRow();
  if (lastRow > 1) projSh.getRange(2, 1, lastRow - 1, PROJ_COLS.length).clearContent();
  if (all.length) projSh.getRange(2, 1, all.length, PROJ_COLS.length).setValues(all);

  // 5) 下流の派生値を更新
  recalcAll_();

  const msg = 'kintone取込 完了: 新規' + nNew + '件 / 更新' + nUpd + '件' + (nSkip ? ' / スキップ' + nSkip + '件(案件ID空)' : '');
  toast_(msg);
  audit_('kintone取込', '', '', '', '', msg);
}

function pick_(srcRow, colOf, logical) {
  const idx = colOf[logical];
  return (idx === undefined) ? '' : srcRow[idx];
}

/** 貼付1行 → 案件データ1行 */
function buildProjRow_(src, colOf, id, ts) {
  const row = new Array(PROJ_COLS.length).fill('');
  row[CI(PROJ_COLS, '案件ID')] = id;
  row[CI(PROJ_COLS, '案件名')] = String(pick_(src, colOf, K.NAME)).trim();
  row[CI(PROJ_COLS, '取引先名')] = String(pick_(src, colOf, K.CLIENT)).trim();
  row[CI(PROJ_COLS, 'フェーズ')] = String(pick_(src, colOf, K.PHASE)).trim();
  row[CI(PROJ_COLS, '納品予定日_始')] = formatDateCell_(pick_(src, colOf, K.DUE_START));
  row[CI(PROJ_COLS, '納品予定日_終')] = formatDateCell_(pick_(src, colOf, K.DUE_END));
  row[CI(PROJ_COLS, '日数')] = pick_(src, colOf, K.DAYS);
  row[CI(PROJ_COLS, '親案件ID')] = String(pick_(src, colOf, K.PARENT_ID)).trim();
  row[CI(PROJ_COLS, '親案件名')] = String(pick_(src, colOf, K.PARENT_NAME)).trim();
  row[CI(PROJ_COLS, '最終取込日時')] = ts;

  // 請求月 = 納品予定日_終 の月(自動導出。例外は案件入力の「請求月上書き」で対応)
  row[CI(PROJ_COLS, '請求月_自動')] = ymOf_(pick_(src, colOf, K.DUE_END));

  // 案件所有者:複数記載(改行区切り)は1行目を採用し警告。マスタ照合不可も警告
  const ownerRaw = String(pick_(src, colOf, K.OWNER) || '').trim();
  const warns = [];
  let ownerEmail = '';
  if (ownerRaw) {
    const lines = ownerRaw.split(/\r?\n/).map(function (s) { return s.trim(); }).filter(String);
    ownerEmail = (lines[0] || '').toLowerCase();
    if (lines.length > 1) warns.push('所有者複数記載(1行目を採用)');
  }
  row[CI(PROJ_COLS, '案件所有者メール')] = ownerEmail;
  const member = ownerEmail ? memberByEmail_(ownerEmail) : null;
  row[CI(PROJ_COLS, '所有者メンバー名')] = member ? member.name : '';
  if (ownerEmail && !member) warns.push('メンバーマスタと照合不可');
  row[CI(PROJ_COLS, '所有者警告')] = warns.join(' / ');
  return row;
}

/** 日付セルを 'yyyy/MM/dd' 文字列に正規化(値が日付でなければそのまま) */
function formatDateCell_(v) {
  const d = asDate_(v);
  return d ? Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd') : (v === null ? '' : v);
}

/** UIに出すエラー(メニュー実行時はalert、それ以外はthrow) */
function uiError_(msg) {
  try {
    SpreadsheetApp.getUi().alert(APP_NAME, msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) { /* UIなし実行 */ }
  return new Error(msg);
}
