/**
 * ビュー描画(V_個人 / V_ユニット / V_経営)
 * フィルター操作なしで見えるよう、GASが値を書き出す方式。
 * 各ビューは「半期サマリ」→「月別」→「案件明細」の3段構成(仕様§7)
 */

function refreshAllViews() {
  recalcAll_();
  const vp = ss_().getSheetByName(SHEET.V_PERSON);
  if (vp) {
    setMemberSelector_(vp.getRange('B1'));
    buildPersonView_(String(vp.getRange('B1').getValue()));
  }
  const vu = ss_().getSheetByName(SHEET.V_UNIT);
  if (vu) {
    setUnitSelector_(vu.getRange('B1'));
    buildUnitView_(String(vu.getRange('B1').getValue()));
  }
  buildMgmtView_();
  toast_('ビューを更新しました');
}

// ================= 共通データ =================

function viewData_() {
  const a = function (n) { return CI(AL_COLS, n); };
  const c = function (n) { return CI(PI_COLS, n); };
  const al = readSheet_(SHEET.ALLOC).rows.filter(function (r) {
    return String(r[a('分配ID')]).trim() && !truthy_(r[a('無効')]);
  });
  const pi = readSheet_(SHEET.PI).rows.filter(function (r) { return String(r[c('案件ID')]).trim(); });
  const piByProj = {};
  pi.forEach(function (r) { piByProj[String(r[c('案件ID')]).trim()] = r; });
  return { al: al, pi: pi, piByProj: piByProj, proj: loadProjMap_(), a: a, c: c };
}

/** メンバー×月 集計(有効分配行から):{member: {ym: {fixed, mikomi, projects:{}}}} */
function aggByMember_(d) {
  const out = {};
  d.al.forEach(function (r) {
    const member = String(r[d.a('メンバー')]).trim();
    if (!member) return;
    const ym = normYm_(r[d.a('請求月')]) || '(未確定)';
    const fixed = numOrEmpty_(r[d.a('確定額')]);
    const est = numOrEmpty_(r[d.a('予測額')]);
    if (!out[member]) out[member] = {};
    if (!out[member][ym]) out[member][ym] = { fixed: 0, mikomi: 0, projects: {} };
    if (fixed !== '') out[member][ym].fixed += fixed;
    out[member][ym].mikomi += (fixed !== '') ? fixed : (est !== '' ? est : 0);
    const id = String(r[d.a('案件ID')]).trim();
    if (id) out[member][ym].projects[id] = true;
  });
  return out;
}

function sumHalf_(memberAgg, halfMonths) {
  const res = { fixed: 0, mikomi: 0, projects: {} };
  if (!memberAgg) return res;
  halfMonths.forEach(function (ym) {
    const v = memberAgg[ym];
    if (!v) return;
    res.fixed += v.fixed;
    res.mikomi += v.mikomi;
    Object.keys(v.projects).forEach(function (id) { res.projects[id] = true; });
  });
  return res;
}

// ================= 描画ヘルパー =================

/**
 * blocks: [{title, headers, rows, highlights(行index配列)}] を縦に並べて書き出す
 */
function paintView_(sh, startRow, blocks, width) {
  const maxW = width || 10;
  const last = Math.max(sh.getLastRow(), startRow);
  sh.getRange(startRow, 1, last - startRow + 1, Math.max(sh.getLastColumn(), maxW)).clear({ contentsOnly: false, skipFilteredRows: false });

  const out = [];
  const titleRows = [];
  const headerRows = [];
  const hlRows = [];
  blocks.forEach(function (b) {
    out.push(padRow_(['■ ' + b.title], maxW));
    titleRows.push(startRow + out.length - 1);
    if (b.headers) {
      out.push(padRow_(b.headers, maxW));
      headerRows.push(startRow + out.length - 1);
    }
    (b.rows || []).forEach(function (r, i) {
      out.push(padRow_(r, maxW));
      if (b.highlights && b.highlights.indexOf(i) >= 0) hlRows.push(startRow + out.length - 1);
    });
    out.push(padRow_([], maxW)); // 区切りの空行
  });
  if (!out.length) return;
  sh.getRange(startRow, 1, out.length, maxW).setValues(out).setNumberFormat('@').setNumberFormats(numFormats_(out));
  titleRows.forEach(function (r) { sh.getRange(r, 1, 1, maxW).setFontWeight('bold').setBackground('#dce6f1'); });
  headerRows.forEach(function (r) { sh.getRange(r, 1, 1, maxW).setFontWeight('bold').setBackground('#efefef'); });
  hlRows.forEach(function (r) { sh.getRange(r, 1, 1, maxW).setBackground('#fce8e6'); });
}

function padRow_(r, w) {
  const row = r.slice(0, w);
  while (row.length < w) row.push('');
  return row;
}

/** 数値セルだけ #,##0 表示 */
function numFormats_(out) {
  return out.map(function (row) {
    return row.map(function (v) { return (typeof v === 'number') ? '#,##0' : '@'; });
  });
}

function pct_(num, den) {
  if (!den) return '—';
  return Math.round(num / den * 100) + '%';
}

// ================= V_個人 =================

function buildPersonView_(name) {
  const sh = sheet_(SHEET.V_PERSON);
  if (!name) { paintView_(sh, 3, [{ title: 'B1セルで自分の名前を選択してください', rows: [] }]); return; }
  const member = memberByName_(name);
  if (!member) { paintView_(sh, 3, [{ title: 'メンバーマスタに「' + name + '」がありません', rows: [] }]); return; }

  const d = viewData_();
  const agg = aggByMember_(d)[name] || {};
  const half = currentHalf_();
  const months = monthsOfHalf_(half);
  const sum = sumHalf_(agg, months);
  const target = targetOf_(member);
  const blocks = [];

  // --- あなたのやること(台帳の見える化:仕様§8) ---
  blocks.push({
    title: 'あなたのやること', headers: ['種別', '案件ID', '案件名', '内容'],
    rows: todoRowsFor_(d, member),
  });

  // --- 半期サマリ ---
  blocks.push({
    title: '半期サマリ(' + half + ')',
    headers: ['目標', '確定粗利', '予測粗利(確定+進行中)', '目標差額(見込−目標)', '達成率(確定)'],
    rows: [[target, sum.fixed, sum.mikomi, sum.mikomi - target, pct_(sum.fixed, target)]],
  });

  // --- 月別表(当半期6ヶ月+データのある先の月) ---
  const futureMonths = Object.keys(agg).filter(function (ym) {
    return /^\d{4}\/\d{2}$/.test(ym) && ym > months[months.length - 1];
  }).sort();
  const monthRows = months.concat(futureMonths).map(function (ym) {
    const v = agg[ym] || { fixed: 0, mikomi: 0, projects: {} };
    const mark = (ym > months[months.length - 1]) ? '(来期)' : '';
    return [ym + mark, v.fixed, v.mikomi, Object.keys(v.projects).length];
  });
  blocks.push({ title: '月別(請求月ベース)', headers: ['月', '確定', '見込', '関与案件数'], rows: monthRows });

  // --- 来期計上分(稼働したのに見えない、を防ぐ) ---
  const nextHalves = {};
  futureMonths.forEach(function (ym) {
    const h = halfOfYm_(ym);
    if (!nextHalves[h]) nextHalves[h] = 0;
    nextHalves[h] += (agg[ym] || { mikomi: 0 }).mikomi;
  });
  blocks.push({
    title: '来期以降の計上分', headers: ['半期', '見込額'],
    rows: Object.keys(nextHalves).sort(function (x, y) { return halfSortKey_(x) - halfSortKey_(y); })
      .map(function (h) { return [h, nextHalves[h]]; }),
  });

  // --- 自分の案件一覧(親案件グルーピング+小計:仕様§5.4) ---
  const list = personProjects_(d, member);
  blocks.push({
    title: '自分の案件一覧(新しい順・親案件でグループ)',
    headers: ['親案件名', '案件名', '案件ID', '役割', '請求月', '予測額', '確定額', 'ステータス', '警告'],
    rows: list.rows, highlights: list.highlights,
  });

  paintView_(sh, 3, blocks, 9);
}

/** やることリスト */
function todoRowsFor_(d, member) {
  const rows = [];
  const nowYm = ymOf_(new Date());

  // 1) 自分が案件所有者なのに未登録(kintone取込済・失注以外)
  Object.keys(d.proj).forEach(function (id) {
    const p = d.proj[id];
    if (p.phase === PHASE_LOST) return;
    if (p.ownerName === member.name && !d.piByProj[id]) {
      rows.push(['案件登録', id, p.name, '案件所有者ですが見積が未登録です']);
    }
  });

  d.pi.forEach(function (r) {
    const id = String(r[d.c('案件ID')]).trim();
    if (String(r[d.c('メイン担当')]).trim() !== member.name) return;
    // 2) 見積未入力(直接行が作られた場合など)
    if (r[d.c('入力ステータス')] === ST_NO_ESTIMATE) {
      rows.push(['見積入力', id, r[d.c('案件名')], '見積が未入力です']);
    }
    // 3) 請求月を過ぎたのに確定待ち
    const ym = normYm_(r[d.c('適用請求月')]);
    if (r[d.c('入力ステータス')] === ST_WAIT_FIX && ym && ym < nowYm) {
      rows.push(['確定値入力', id, r[d.c('案件名')], '請求月(' + ym + ')を過ぎています']);
    }
  });

  // 4) 自分の制作%が0のまま(メイン担当への催促材料)
  d.al.forEach(function (r) {
    if (String(r[d.a('メンバー')]).trim() !== member.name) return;
    if (r[d.a('種別')] === KIND_PROD && !truthy_(r[d.a('自動行')]) && (Number(r[d.a('制作%')]) || 0) === 0) {
      rows.push(['%未設定', String(r[d.a('案件ID')]), String(r[d.a('案件名')]), '制作%が未設定です(メイン担当に確認)']);
    }
  });

  return rows.length ? rows : [['—', '', '', '現在やることはありません']];
}

/** 個人の案件一覧(親案件グループ+小計行) */
function personProjects_(d, member) {
  const items = {}; // 案件ID → {roles:{}, est, fixed}
  d.al.forEach(function (r) {
    if (String(r[d.a('メンバー')]).trim() !== member.name) return;
    const id = String(r[d.a('案件ID')]).trim();
    if (!id) return;
    if (!items[id]) items[id] = { roles: {}, est: 0, fixed: 0, hasFixed: false };
    items[id].roles[r[d.a('種別')]] = true;
    const est = numOrEmpty_(r[d.a('予測額')]);
    const fixed = numOrEmpty_(r[d.a('確定額')]);
    if (est !== '') items[id].est += est;
    if (fixed !== '') { items[id].fixed += fixed; items[id].hasFixed = true; }
  });
  d.pi.forEach(function (r) {
    const id = String(r[d.c('案件ID')]).trim();
    if (String(r[d.c('メイン担当')]).trim() === member.name) {
      if (!items[id]) items[id] = { roles: {}, est: 0, fixed: 0, hasFixed: false };
      items[id].roles['メイン担当'] = true;
    }
  });

  const entries = Object.keys(items).map(function (id) {
    const p = d.proj[id] || {};
    const pi = d.piByProj[id];
    const roles = [];
    if (items[id].roles['メイン担当']) roles.push('メイン');
    if (items[id].roles[KIND_FIELD]) roles.push('現場');
    if (items[id].roles[KIND_PROD]) roles.push('制作');
    return {
      id: id,
      parent: p.parentName || '',
      name: p.name || (pi ? pi[d.c('案件名')] : ''),
      roles: roles.join('+'),
      ym: pi ? normYm_(pi[d.c('適用請求月')]) : (p.ymAuto || ''),
      est: items[id].est,
      fixed: items[id].hasFixed ? items[id].fixed : '',
      status: pi ? pi[d.c('入力ステータス')] : '未登録',
      warn: pi ? String(pi[d.c('警告')]) : '',
    };
  });

  // 親案件ごとにグループ化し、グループ内・グループ間とも請求月の新しい順
  const groups = {};
  entries.forEach(function (e) {
    const key = e.parent || '#single#' + e.id;
    if (!groups[key]) groups[key] = [];
    groups[key].push(e);
  });
  const groupList = Object.keys(groups).map(function (k) {
    const list = groups[k].sort(function (x, y) { return (y.ym || '').localeCompare(x.ym || ''); });
    return { key: k, list: list, maxYm: list[0].ym || '' };
  }).sort(function (x, y) { return y.maxYm.localeCompare(x.maxYm); });

  const rows = [];
  const highlights = [];
  groupList.forEach(function (g) {
    g.list.forEach(function (e) {
      rows.push([e.parent, e.name, e.id, e.roles, e.ym, e.est, e.fixed, e.status, e.warn]);
      if (e.warn) highlights.push(rows.length - 1);
    });
    if (g.list.length > 1) {
      const subEst = g.list.reduce(function (s, e) { return s + (Number(e.est) || 0); }, 0);
      const subFix = g.list.reduce(function (s, e) { return s + (Number(e.fixed) || 0); }, 0);
      rows.push(['└ 小計【' + g.list[0].parent + '】', '', '', '', '', subEst, subFix, '', '']);
    }
  });
  return { rows: rows, highlights: highlights };
}

// ================= V_ユニット =================

function buildUnitView_(unit) {
  const sh = sheet_(SHEET.V_UNIT);
  if (!unit) { paintView_(sh, 3, [{ title: 'B1セルでユニットを選択してください', rows: [] }]); return; }

  const d = viewData_();
  const agg = aggByMember_(d);
  const half = currentHalf_();
  const months = monthsOfHalf_(half);
  const unitMembers = members_().filter(function (m) { return m.unit === unit; });
  const blocks = [];

  const table = memberTable_(unitMembers, agg, months);
  blocks.push({
    title: 'ユニットサマリ(' + half + ') ' + unit,
    headers: ['ユニット目標(メンバー目標合計)', '確定合計', '見込合計', '差額(見込−目標)', '達成率(確定)'],
    rows: [[table.targetSum, table.fixedSum, table.mikomiSum, table.mikomiSum - table.targetSum, pct_(table.fixedSum, table.targetSum)]],
  });

  blocks.push({
    title: 'メンバー別(' + half + ')※赤=見込でも目標未達',
    headers: ['氏名', '等級', '目標', '確定', '見込', '差額(見込−目標)', '達成率(確定)', '関与案件数'],
    rows: table.rows, highlights: table.highlights,
  });

  // 未入力一覧(ユニット長が声をかけられるように)
  const todoRows = [];
  unitMembers.forEach(function (m) {
    todoRowsFor_(d, m).forEach(function (r) {
      if (r[0] !== '—') todoRows.push([m.name].concat(r));
    });
  });
  blocks.push({
    title: '未入力・要対応一覧', headers: ['メンバー', '種別', '案件ID', '案件名', '内容'],
    rows: todoRows.length ? todoRows : [['—', '', '', '', '対応が必要な項目はありません']],
  });

  paintView_(sh, 3, blocks, 8);
}

/** メンバー横並び表(共通) */
function memberTable_(memberList, agg, months) {
  let targetSum = 0, fixedSum = 0, mikomiSum = 0;
  const rows = [];
  const highlights = [];
  memberList.forEach(function (m) {
    const s = sumHalf_(agg[m.name], months);
    const target = targetOf_(m);
    targetSum += target;
    fixedSum += s.fixed;
    mikomiSum += s.mikomi;
    rows.push([m.name + (m.isLead ? '(ユニット長)' : ''), m.grade, target, s.fixed, s.mikomi,
      s.mikomi - target, pct_(s.fixed, target), Object.keys(s.projects).length]);
    if (s.mikomi < target) highlights.push(rows.length - 1); // 予測込みでも届かない
  });
  return { rows: rows, highlights: highlights, targetSum: targetSum, fixedSum: fixedSum, mikomiSum: mikomiSum };
}

// ================= V_経営 =================

function buildMgmtView_() {
  const sh = ss_().getSheetByName(SHEET.V_MGMT);
  if (!sh) return;

  const d = viewData_();
  const agg = aggByMember_(d);
  const half = currentHalf_();
  const months = monthsOfHalf_(half);
  const all = members_();
  const units = [];
  all.forEach(function (m) { if (m.unit && units.indexOf(m.unit) < 0) units.push(m.unit); });
  const blocks = [];

  // 全社+ユニット並列の半期サマリ
  const totalTable = memberTable_(all, agg, months);
  const summaryRows = [['全社', totalTable.targetSum, totalTable.fixedSum, totalTable.mikomiSum,
    totalTable.mikomiSum - totalTable.targetSum, pct_(totalTable.fixedSum, totalTable.targetSum)]];
  const unitTables = {};
  units.forEach(function (u) {
    const t = memberTable_(all.filter(function (m) { return m.unit === u; }), agg, months);
    unitTables[u] = t;
    summaryRows.push([u, t.targetSum, t.fixedSum, t.mikomiSum, t.mikomiSum - t.targetSum, pct_(t.fixedSum, t.targetSum)]);
  });
  blocks.push({
    title: '半期サマリ(' + half + ')全社+ユニット別',
    headers: ['区分', '目標', '確定', '見込', '差額', '達成率(確定)'],
    rows: summaryRows,
  });

  // 月別推移(全社+ユニット)
  const headers = ['月', '全社_確定', '全社_見込'];
  units.forEach(function (u) { headers.push(u + '_確定', u + '_見込'); });
  const monthRows = months.map(function (ym) {
    let row = [ym];
    let tf = 0, tm = 0;
    const perUnit = {};
    units.forEach(function (u) { perUnit[u] = { f: 0, m: 0 }; });
    all.forEach(function (m) {
      const v = (agg[m.name] || {})[ym];
      if (!v) return;
      tf += v.fixed; tm += v.mikomi;
      if (perUnit[m.unit]) { perUnit[m.unit].f += v.fixed; perUnit[m.unit].m += v.mikomi; }
    });
    row.push(tf, tm);
    units.forEach(function (u) { row.push(perUnit[u].f, perUnit[u].m); });
    return row;
  });
  blocks.push({ title: '月別推移(請求月ベース)', headers: headers, rows: monthRows });

  // ユニット別メンバー横並び表を縦に並べる(別ページに分散させない:仕様§7.3)
  units.forEach(function (u) {
    const t = unitTables[u];
    blocks.push({
      title: 'メンバー別 ' + u + '(' + half + ')',
      headers: ['氏名', '等級', '目標', '確定', '見込', '差額', '達成率(確定)', '関与案件数'],
      rows: t.rows, highlights: t.highlights,
    });
  });

  paintView_(sh, 3, blocks, Math.max(8, 3 + units.length * 2));
}
