/**
 * 再計算エンジン:
 * - 案件入力/分配入力 の自動列(請求月・半期・原資・警告・ステータス等)を更新
 * - メイン担当の制作%「自動行」(100 − 他メンバー合計)を分配入力にメンテナンス
 * - 集計シート(個人×請求月×予測/確定)を再構築
 * 取込・フォーム送信のたびに呼ばれる(30名×数百案件規模なので全件再計算で十分)
 */

// ---- 読込ヘルパー ----

/** 案件データ → Map(案件ID → info) */
function loadProjMap_() {
  const data = readSheet_(SHEET.PROJ);
  const map = {};
  data.rows.forEach(function (r) {
    const id = String(r[CI(PROJ_COLS, '案件ID')]).trim();
    if (!id) return;
    map[id] = {
      id: id,
      name: r[CI(PROJ_COLS, '案件名')],
      client: r[CI(PROJ_COLS, '取引先名')],
      phase: String(r[CI(PROJ_COLS, 'フェーズ')]),
      ymAuto: normYm_(r[CI(PROJ_COLS, '請求月_自動')]),
      ownerEmail: String(r[CI(PROJ_COLS, '案件所有者メール')]),
      ownerName: String(r[CI(PROJ_COLS, '所有者メンバー名')]),
      ownerWarn: String(r[CI(PROJ_COLS, '所有者警告')]),
      parentId: String(r[CI(PROJ_COLS, '親案件ID')]),
      parentName: String(r[CI(PROJ_COLS, '親案件名')]),
      dueStart: r[CI(PROJ_COLS, '納品予定日_始')],
      dueEnd: r[CI(PROJ_COLS, '納品予定日_終')],
    };
  });
  return map;
}

function numOrEmpty_(v) {
  if (v === '' || v === null || v === undefined) return '';
  const n = Number(v);
  return isFinite(n) ? n : '';
}

// ---- 本体 ----

function recalcAll_() {
  const proj = loadProjMap_();
  const piSh = sheet_(SHEET.PI);
  const alSh = sheet_(SHEET.ALLOC);
  const piData = piSh.getDataRange().getValues();
  const alData = alSh.getDataRange().getValues();
  const piRows = piData.slice(1);
  const alRows = alData.slice(1);

  const c = function (name) { return CI(PI_COLS, name); };
  const a = function (name) { return CI(AL_COLS, name); };
  const grades = masters_().grades;
  const positions = masters_().positions;

  // --- 案件入力:案件参照系の更新 + 案件ID→PI行 の索引 ---
  const piByProj = {};
  piRows.forEach(function (r) {
    const id = String(r[c('案件ID')]).trim();
    if (!id) return;
    piByProj[id] = r;
    const p = proj[id];
    if (p) {
      r[c('案件名')] = p.name;
      r[c('クライアント')] = p.client;
      r[c('請求月_自動')] = p.ymAuto;
      r[c('親案件名')] = p.parentName;
    }
    const ovr = normYm_(r[c('請求月上書き')]);
    const ym = ovr || normYm_(r[c('請求月_自動')]);
    r[c('適用請求月')] = ym;
    r[c('半期')] = halfOfYm_(ym);
  });

  // --- 分配入力:案件参照系の更新 + 案件別の積上げ ---
  const fieldTotal = {};   // 案件ID → 現場金額合計(有効行のみ)
  const prodOthers = {};   // 案件ID → メイン担当以外(非自動行)の%合計
  const gradeWarn = {};    // 案件ID → 等級条件違反あり
  alRows.forEach(function (r) {
    const id = String(r[a('案件ID')]).trim();
    if (!id) return;
    const p = proj[id];
    const pi = piByProj[id];
    if (p) r[a('案件名')] = p.name;
    const ym = pi ? pi[c('適用請求月')] : (p ? p.ymAuto : normYm_(r[a('請求月')]));
    r[a('請求月')] = ym;
    r[a('半期')] = halfOfYm_(ym);

    if (truthy_(r[a('無効')])) return; // 論理削除済みは積上げ対象外

    if (r[a('種別')] === KIND_FIELD) {
      const amt = Number(r[a('金額')]) || 0;
      fieldTotal[id] = (fieldTotal[id] || 0) + amt;
      // 等級条件チェック(警告のみ)
      const pos = positions[String(r[a('ポジション')])];
      const g = String(r[a('稼働時等級')]);
      if (pos && pos.minGrade && grades[g] !== undefined && grades[pos.minGrade] !== undefined
          && grades[g] < grades[pos.minGrade]) {
        gradeWarn[id] = true;
      }
      // 現場分は固定単価:予測額=確定額=金額(仕様§6.2)
      r[a('予測額')] = amt;
      r[a('確定額')] = amt;
    } else if (r[a('種別')] === KIND_PROD && !truthy_(r[a('自動行')])) {
      prodOthers[id] = (prodOthers[id] || 0) + (Number(r[a('制作%')]) || 0);
    }
  });

  // --- メイン担当の自動行を upsert(登録済み案件すべて) ---
  const autoRowByProj = {};
  alRows.forEach(function (r) {
    const id = String(r[a('案件ID')]).trim();
    if (id && truthy_(r[a('自動行')]) && !truthy_(r[a('無効')])) autoRowByProj[id] = r;
  });
  const appended = [];
  Object.keys(piByProj).forEach(function (id) {
    const pi = piByProj[id];
    const main = String(pi[c('メイン担当')]).trim();
    if (!main) return;
    const residual = Math.max(0, 100 - (prodOthers[id] || 0));
    let row = autoRowByProj[id];
    if (!row) {
      row = new Array(AL_COLS.length).fill('');
      row[a('分配ID')] = 'AUTO-' + id;
      row[a('案件ID')] = id;
      row[a('種別')] = KIND_PROD;
      row[a('自動行')] = true;
      row[a('入力者')] = '(自動)';
      row[a('入力日時')] = now_();
      const p = proj[id];
      if (p) row[a('案件名')] = p.name;
      row[a('請求月')] = pi[c('適用請求月')];
      row[a('半期')] = halfOfYm_(pi[c('適用請求月')]);
      appended.push(row);
      alRows.push(row);
    }
    row[a('メンバー')] = main;        // メイン担当変更に追従
    row[a('制作%')] = residual;       // 100 − 他メンバー合計
  });

  // --- 案件入力:集計列・警告・ステータス ---
  piRows.forEach(function (r) {
    const id = String(r[c('案件ID')]).trim();
    if (!id) return;
    const p = proj[id];
    const estGp = numOrEmpty_(r[c('見積_粗利')]);
    const fixGp = numOrEmpty_(r[c('確定_粗利')]);
    const ft = fieldTotal[id] || 0;

    r[c('現場稼働合計')] = ft;
    const baseEst = (estGp === '') ? '' : estGp - ft;
    const baseFix = (fixGp === '') ? '' : fixGp - ft;
    r[c('制作原資_予測')] = baseEst;
    r[c('制作原資_確定')] = baseFix;

    const others = prodOthers[id] || 0;
    const main = String(r[c('メイン担当')]).trim();
    const pctTotal = main ? Math.max(0, 100 - others) + others : others;
    r[c('制作%合計')] = pctTotal;

    // 入力ステータス / ステータス
    const inputSt = (estGp === '') ? ST_NO_ESTIMATE : (fixGp === '' ? ST_WAIT_FIX : ST_FIXED);
    r[c('入力ステータス')] = inputSt;
    r[c('ステータス')] = (fixGp === '') ? '進行中' : '確定済';
    r[c('見積確定差額')] = (estGp !== '' && fixGp !== '') ? fixGp - estGp : '';

    // 警告(止めない:表示のみ)
    const warns = [];
    if (!p) warns.push('案件データに存在しない案件ID');
    if (p && p.ownerWarn) warns.push(p.ownerWarn);
    if (baseEst !== '' && baseEst < 0) warns.push('赤字(予測原資マイナス)');
    if (baseFix !== '' && baseFix < 0) warns.push('赤字(確定原資マイナス)');
    if (pctTotal !== 100 && (others > 0 || main)) {
      if (pctTotal !== 100) warns.push('制作%合計≠100(' + pctTotal + '%)');
    }
    if (others > 100) warns.push('他メンバー%が100超');
    if (gradeWarn[id]) warns.push('等級条件違反の現場分配あり');
    if (!r[c('適用請求月')]) warns.push('請求月未確定(納品予定日_終が空)');
    r[c('警告')] = warns.join(' / ');
  });

  // --- 分配入力:制作行の予測額/確定額(原資×%。原資が動けば追従) ---
  alRows.forEach(function (r) {
    const id = String(r[a('案件ID')]).trim();
    if (!id || r[a('種別')] !== KIND_PROD) return;
    if (truthy_(r[a('無効')])) return;
    const pi = piByProj[id];
    const pct = Number(r[a('制作%')]) || 0;
    if (!pi) { r[a('予測額')] = ''; r[a('確定額')] = ''; return; }
    const baseEst = numOrEmpty_(pi[c('制作原資_予測')]);
    const baseFix = numOrEmpty_(pi[c('制作原資_確定')]);
    // 赤字でもマイナスのまま%按分(仕様§2)
    r[a('予測額')] = (baseEst === '') ? '' : roundYen_(baseEst * pct / 100);
    r[a('確定額')] = (baseFix === '') ? '' : roundYen_(baseFix * pct / 100);
  });

  // --- 書き戻し ---
  if (piRows.length) piSh.getRange(2, 1, piRows.length, PI_COLS.length).setValues(piRows);
  const existingAl = alRows.length - appended.length;
  if (existingAl > 0) alSh.getRange(2, 1, existingAl, AL_COLS.length).setValues(alRows.slice(0, existingAl));
  if (appended.length) alSh.getRange(existingAl + 2, 1, appended.length, AL_COLS.length).setValues(appended);

  // --- 集計シート再構築 ---
  buildAggregate_(alRows, piByProj);
}

function truthy_(v) {
  return v === true || String(v).toUpperCase() === 'TRUE';
}

/**
 * 集計(非表示):メンバー × 請求月 × 確定/予測/見込
 * 見込額 = 確定額があれば確定額、なければ予測額(=確定+進行中予測)
 */
function buildAggregate_(alRows, piByProj) {
  const a = function (name) { return CI(AL_COLS, name); };
  const agg = {}; // key: member|ym
  alRows.forEach(function (r) {
    if (truthy_(r[a('無効')])) return;
    const member = String(r[a('メンバー')]).trim();
    const ym = normYm_(r[a('請求月')]);
    const id = String(r[a('案件ID')]).trim();
    if (!member) return;
    const key = member + '|' + (ym || '(請求月未確定)');
    if (!agg[key]) agg[key] = { member: member, ym: ym || '(請求月未確定)', fixed: 0, est: 0, mikomi: 0, projects: {} };
    const fixed = numOrEmpty_(r[a('確定額')]);
    const est = numOrEmpty_(r[a('予測額')]);
    if (fixed !== '') agg[key].fixed += fixed;
    if (est !== '') agg[key].est += est;
    const mk = (fixed !== '') ? fixed : (est !== '' ? est : 0);
    agg[key].mikomi += mk;
    if (id) agg[key].projects[id] = true;
  });

  const out = Object.keys(agg).sort().map(function (k) {
    const v = agg[k];
    return [v.member, v.ym, halfOfYm_(normYm_(v.ym)), v.fixed, v.est, v.mikomi, Object.keys(v.projects).length];
  });

  const sh = sheet_(SHEET.AGG);
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, AGG_COLS.length).clearContent();
  if (out.length) sh.getRange(2, 1, out.length, AGG_COLS.length).setValues(out);
}
