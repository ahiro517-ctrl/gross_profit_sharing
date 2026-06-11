/**
 * サイドバーフォームのサーバ側処理
 * すべての書込はここを経由する(監査ログ+再計算をセットで実行)
 * ハードブロック:他メンバー%合計>100 / マスタ不整合 / 締め済み半期 / 確定値の再入力
 * 警告のみ(止めない):赤字 / 等級条件違反
 */

// ================= 初期データ =================

/** フォーム共通の初期データ取得(formKind: project|confirm|field|prod|bulk|cancel) */
function fsGetInit(formKind) {
  const proj = loadProjMap_();
  const piData = readSheet_(SHEET.PI);
  const c = function (n) { return CI(PI_COLS, n); };
  const piByProj = {};
  piData.rows.forEach(function (r) {
    const id = String(r[c('案件ID')]).trim();
    if (id) piByProj[id] = r;
  });

  const me = memberByEmail_(userEmail_());
  const base = {
    me: me ? { name: me.name, email: me.email } : null,
    members: members_().map(function (m) { return m.name; }),
    isAdmin: isAdmin_(),
  };

  if (formKind === 'project') {
    // 失注・運用開始月より前の案件は選択肢から除外。未登録の案件のみ
    base.projects = Object.keys(proj).filter(function (id) {
      return proj[id].phase !== PHASE_LOST && !piByProj[id] && !isBeforeOpStart_(proj[id].ymAuto);
    }).map(function (id) {
      const p = proj[id];
      return { id: id, label: id + ':' + p.name + '(' + p.client + ')', ownerName: p.ownerName, ym: p.ymAuto };
    });
  }

  if (formKind === 'confirm') {
    base.projects = Object.keys(piByProj).filter(function (id) {
      const r = piByProj[id];
      return numOrEmpty_(r[c('見積_粗利')]) !== '' && numOrEmpty_(r[c('確定_粗利')]) === '';
    }).map(function (id) {
      const r = piByProj[id];
      return {
        id: id, label: id + ':' + r[c('案件名')],
        estSales: numOrEmpty_(r[c('見積_売上')]), estCost: numOrEmpty_(r[c('見積_原価')]), estGp: numOrEmpty_(r[c('見積_粗利')]),
        main: String(r[c('メイン担当')]), ym: normYm_(r[c('適用請求月')]),
      };
    });
  }

  if (formKind === 'field') {
    // 現場分配は登録前の案件にも入力可(固定単価のため原資に依存しない)。
    // 失注と運用開始月より前の案件は除外
    base.projects = Object.keys(proj).filter(function (id) {
      return proj[id].phase !== PHASE_LOST && !isBeforeOpStart_(proj[id].ymAuto);
    }).map(function (id) {
      const p = proj[id];
      return { id: id, label: id + ':' + p.name + '(' + p.client + ')' };
    });
    const pos = masters_().positions;
    base.positions = Object.keys(pos).map(function (k) {
      return { name: k, minGrade: pos[k].minGrade, day: pos[k].day, half: pos[k].half };
    });
    base.myGrade = me ? me.grade : '';
    base.grades = masters_().grades;
    base.memberGrades = {};
    members_().forEach(function (m) { base.memberGrades[m.name] = m.grade; });
  }

  if (formKind === 'prod' || formKind === 'bulk') {
    // 既存の制作分配(有効・非自動)を案件ごとに要約(選択時に現状を表示するため)
    const alRows = readSheet_(SHEET.ALLOC).rows;
    const a = function (n) { return CI(AL_COLS, n); };
    const existingProd = {};
    alRows.forEach(function (r) {
      if (r[a('種別')] !== KIND_PROD || truthy_(r[a('無効')]) || truthy_(r[a('自動行')])) return;
      const id = String(r[a('案件ID')]).trim();
      if (!id) return;
      if (!existingProd[id]) existingProd[id] = [];
      existingProd[id].push(String(r[a('メンバー')]) + ' ' + (Number(r[a('制作%')]) || 0) + '%');
    });

    // 制作分配は登録済み(メイン担当あり)案件のみ。運用開始月より前は除外
    base.projects = Object.keys(piByProj).filter(function (id) {
      return String(piByProj[id][c('メイン担当')]).trim() !== ''
        && !isBeforeOpStart_(normYm_(piByProj[id][c('適用請求月')]));
    }).map(function (id) {
      const r = piByProj[id];
      const p = proj[id] || {};
      return {
        id: id, label: id + ':' + r[c('案件名')], main: String(r[c('メイン担当')]),
        parentId: p.parentId || '', parentName: p.parentName || '', ym: normYm_(r[c('適用請求月')]),
        existing: (existingProd[id] || []).join('、'),
      };
    });
    if (formKind === 'bulk') {
      const parents = {};
      base.projects.forEach(function (p) {
        if (p.parentId) parents[p.parentId] = p.parentName || p.parentId;
      });
      base.parents = Object.keys(parents).map(function (id) { return { id: id, name: parents[id] }; });
    }
  }

  if (formKind === 'cancel') {
    const al = readSheet_(SHEET.ALLOC);
    const a = function (n) { return CI(AL_COLS, n); };
    base.allocs = al.rows.filter(function (r) {
      return String(r[a('分配ID')]).trim() && !truthy_(r[a('無効')]) && !truthy_(r[a('自動行')]);
    }).map(function (r) {
      const memo = String(r[a('稼働メモ')] || '');
      const detail = r[a('種別')] === KIND_FIELD
        ? (r[a('ポジション')] + ' ' + r[a('日数')] + '日 ' + r[a('金額')] + '円' + (memo ? '【' + memo + '】' : ''))
        : (r[a('制作%')] + '%');
      return {
        id: r[a('分配ID')], projId: r[a('案件ID')], projName: r[a('案件名')],
        member: r[a('メンバー')], kind: r[a('種別')], detail: detail,
        by: r[a('入力者')], at: String(r[a('入力日時')]), bulkId: r[a('まとめ分配ID')],
      };
    });
  }

  return base;
}

// ================= 案件登録(見積入力) =================

function fsSubmitProject(p) {
  const proj = loadProjMap_()[String(p.projId).trim()];
  if (!proj) throw new Error('案件データに存在しない案件IDです: ' + p.projId);
  if (proj.phase === PHASE_LOST) throw new Error('失注案件は登録できません');

  const piSh = sheet_(SHEET.PI);
  const c = function (n) { return CI(PI_COLS, n); };
  if (findPiRow_(piSh, p.projId) > 0) {
    throw new Error('この案件は登録済みです。見積値の修正は管理者にご相談ください。');
  }
  const main = memberByName_(p.mainName);
  if (!main) throw new Error('メイン担当がメンバーマスタにありません: ' + p.mainName);

  const ymOverride = normYm_(p.ymOverride);
  if (p.ymOverride && !ymOverride) throw new Error('請求月上書きは yyyy/MM 形式で入力してください');
  const ym = ymOverride || proj.ymAuto;
  assertNotClosed_(ym);

  const sales = requireNum_(p.sales, '見積_売上');
  const cost = requireNum_(p.cost, '見積_原価');
  const gp = requireNum_(p.gp, '見積_粗利');

  const row = new Array(PI_COLS.length).fill('');
  row[c('案件ID')] = String(p.projId).trim();
  row[c('メイン担当')] = main.name;
  row[c('請求月上書き')] = ymOverride;
  row[c('見積_売上')] = sales;
  row[c('見積_原価')] = cost;
  row[c('見積_粗利')] = gp;
  row[c('登録者')] = userEmail_();
  row[c('登録日時')] = now_();
  piSh.appendRow(row);

  audit_('案件登録', p.projId, '見積', '', { 売上: sales, 原価: cost, 粗利: gp, メイン担当: main.name, 請求月上書き: ymOverride });
  recalcAll_();
  return '案件「' + proj.name + '」を登録しました(見積粗利 ' + gp.toLocaleString() + '円)';
}

// ================= 確定値入力 =================

function fsSubmitConfirm(p) {
  const piSh = sheet_(SHEET.PI);
  const c = function (n) { return CI(PI_COLS, n); };
  const rowIdx = findPiRow_(piSh, p.projId);
  if (rowIdx < 0) throw new Error('案件が登録されていません。先に「案件登録(見積入力)」を実行してください。');

  const row = piSh.getRange(rowIdx, 1, 1, PI_COLS.length).getValues()[0];
  // 改ざん防止:確定値は一度入力したらフォームから再編集不可(仕様§8)
  if (numOrEmpty_(row[c('確定_粗利')]) !== '') {
    throw new Error('確定値は入力済みです。修正は管理者メニュー「確定値修正」のみ可能です。');
  }
  assertNotClosed_(row[c('適用請求月')]);

  const sales = requireNum_(p.sales, '確定_売上');
  const cost = requireNum_(p.cost, '確定_原価');
  const gp = requireNum_(p.gp, '確定_粗利');

  piSh.getRange(rowIdx, c('確定_売上') + 1, 1, 3).setValues([[sales, cost, gp]]);
  piSh.getRange(rowIdx, c('確定入力者') + 1, 1, 2).setValues([[userEmail_(), now_()]]);

  audit_('確定値入力', p.projId, '確定', '', { 売上: sales, 原価: cost, 粗利: gp });
  recalcAll_();
  return '確定値を入力しました(確定粗利 ' + gp.toLocaleString() + '円)。以後の修正は管理者のみ可能です。';
}

// ================= 現場分配 =================

function fsSubmitField(p) {
  const proj = loadProjMap_()[String(p.projId).trim()];
  if (!proj) throw new Error('案件データに存在しない案件IDです: ' + p.projId);
  if (proj.phase === PHASE_LOST) throw new Error('失注案件には分配できません');

  const member = memberByName_(p.memberName);
  if (!member) throw new Error('メンバーマスタにありません: ' + p.memberName);
  const pos = masters_().positions[String(p.position)];
  if (!pos) throw new Error('ポジションがマスタにありません: ' + p.position);
  if (!isHalfStep_(p.days)) throw new Error('日数は0.5刻みの正の数で入力してください');

  const ym = piYmOrAuto_(p.projId, proj);
  assertNotClosed_(ym);

  const days = Number(p.days);
  const fullDays = Math.floor(days);
  const hasHalf = days - fullDays > 0;
  // 半日は係数計算ではなくマスタの半日単価を直接使用(仕様§4.3)
  const amount = fullDays * pos.day + (hasHalf ? pos.half : 0);

  // 稼働時点の等級・単価をスナップショットとして値で固定保存
  const a = function (n) { return CI(AL_COLS, n); };
  const row = new Array(AL_COLS.length).fill('');
  row[a('分配ID')] = newId_('F');
  row[a('案件ID')] = String(p.projId).trim();
  row[a('メンバー')] = member.name;
  row[a('種別')] = KIND_FIELD;
  row[a('ポジション')] = String(p.position);
  row[a('日数')] = days;
  row[a('稼働時等級')] = member.grade;
  row[a('単価_1日')] = pos.day;
  row[a('単価_半日')] = pos.half;
  row[a('金額')] = amount;
  row[a('稼働メモ')] = String(p.memo || '').trim(); // どの現場稼働か(例: 6/14-15 設営)
  row[a('入力者')] = userEmail_();
  row[a('入力日時')] = now_();
  sheet_(SHEET.ALLOC).appendRow(row);

  // 等級条件は警告のみ(入力は通す)
  let warn = '';
  const grades = masters_().grades;
  if (pos.minGrade && grades[member.grade] !== undefined && grades[pos.minGrade] !== undefined
      && grades[member.grade] < grades[pos.minGrade]) {
    warn = '\n⚠ 等級条件(' + p.position + 'は' + pos.minGrade + '以上)を満たしていません(警告のみ・入力は記録されました)';
  }

  audit_('現場分配', p.projId, member.name, '', { ポジション: p.position, 日数: days, 等級: member.grade, 金額: amount });
  recalcAll_();
  return member.name + ' に現場分配 ' + amount.toLocaleString() + '円(' + days + '日)を記録しました' + warn;
}

// ================= 制作分配 =================

function fsSubmitProd(p) {
  const res = prodValidateAndApply_(String(p.projId).trim(), p.items, '');
  recalcAll_();
  return res;
}

/**
 * まとめ分配(仕様§5.3):選択した各案件に同じメンバー%の制作分配行を生成。
 * 計上は案件単位(各案件の制作原資×%、各案件の請求月)のまま。
 */
function fsSubmitBulk(p) {
  const ids = (p.projIds || []).map(function (s) { return String(s).trim(); }).filter(String);
  if (!ids.length) throw new Error('対象案件が選択されていません');
  const bulkId = newId_('B');

  // 先に全案件をバリデーション(1件でもNGなら全体を中断して原因を表示)
  const errors = [];
  ids.forEach(function (id) {
    try { prodValidate_(id, p.items); } catch (e) { errors.push(id + ': ' + e.message); }
  });
  if (errors.length) throw new Error('以下の案件でエラーのため中断しました(1件も登録していません)\n' + errors.join('\n'));

  ids.forEach(function (id) { prodValidateAndApply_(id, p.items, bulkId); });
  audit_('まとめ分配', ids.join(','), 'まとめ分配ID=' + bulkId, '', p.items);
  recalcAll_();
  return ids.length + '案件に制作分配を適用しました(まとめ分配ID: ' + bulkId + ')';
}

/** 制作分配バリデーション(applyしない) */
function prodValidate_(projId, items) {
  const piSh = sheet_(SHEET.PI);
  const c = function (n) { return CI(PI_COLS, n); };
  const rowIdx = findPiRow_(piSh, projId);
  if (rowIdx < 0) throw new Error('案件が未登録です(先に案件登録が必要)');
  const piRow = piSh.getRange(rowIdx, 1, 1, PI_COLS.length).getValues()[0];
  const main = String(piRow[c('メイン担当')]).trim();
  if (!main) throw new Error('メイン担当が未設定です');
  assertNotClosed_(piRow[c('適用請求月')]);

  if (!items || !items.length) throw new Error('メンバーが指定されていません');
  const seen = {};
  let newPct = 0;
  items.forEach(function (it) {
    const m = memberByName_(it.member);
    if (!m) throw new Error('メンバーマスタにありません: ' + it.member);
    if (m.name === main) throw new Error('メイン担当(' + main + ')の%は自動計算のため指定できません');
    if (seen[m.name]) throw new Error('メンバーが重複しています: ' + m.name);
    seen[m.name] = true;
    const pct = Number(it.pct);
    if (!isFinite(pct) || pct <= 0) throw new Error(it.member + ' の%が不正です');
    newPct += pct;
  });

  // 既存の有効な制作行(非自動)と合算して100%超ならハードブロック(仕様§6.2)
  const al = readSheet_(SHEET.ALLOC);
  const a = function (n) { return CI(AL_COLS, n); };
  let existingPct = 0;
  al.rows.forEach(function (r) {
    if (String(r[a('案件ID')]).trim() !== projId) return;
    if (r[a('種別')] !== KIND_PROD || truthy_(r[a('無効')]) || truthy_(r[a('自動行')])) return;
    existingPct += Number(r[a('制作%')]) || 0;
    if (seen[String(r[a('メンバー')]).trim()]) {
      throw new Error(r[a('メンバー')] + ' には既に制作%が設定済みです(変更は分配取消→再入力)');
    }
  });
  if (existingPct + newPct > 100) {
    throw new Error('メイン担当以外の%合計が100%を超えます(既存' + existingPct + '% + 今回' + newPct + '%)');
  }
  return { main: main };
}

/** 制作分配の登録(バリデーション込み) */
function prodValidateAndApply_(projId, items, bulkId) {
  const v = prodValidate_(projId, items);
  const alSh = sheet_(SHEET.ALLOC);
  const a = function (n) { return CI(AL_COLS, n); };
  const rows = items.map(function (it) {
    const row = new Array(AL_COLS.length).fill('');
    row[a('分配ID')] = newId_('P');
    row[a('案件ID')] = projId;
    row[a('メンバー')] = String(it.member).trim();
    row[a('種別')] = KIND_PROD;
    row[a('制作%')] = Number(it.pct);
    row[a('入力者')] = userEmail_();
    row[a('入力日時')] = now_();
    row[a('まとめ分配ID')] = bulkId || '';
    return row;
  });
  alSh.getRange(alSh.getLastRow() + 1, 1, rows.length, AL_COLS.length).setValues(rows);
  if (!bulkId) audit_('制作分配', projId, items.map(function (i) { return i.member + ':' + i.pct + '%'; }).join(', '), '', items);
  return '制作分配を登録しました(メイン担当 ' + v.main + ' の%は自動で残りに更新されます)';
}

// ================= 分配取消(論理削除) =================

function fsSubmitCancel(p) {
  const ids = p.allocIds || [];
  if (!ids.length) throw new Error('取消対象が選択されていません');

  const alSh = sheet_(SHEET.ALLOC);
  const data = alSh.getDataRange().getValues();
  const a = function (n) { return CI(AL_COLS, n); };
  const piSh = sheet_(SHEET.PI);
  const c = function (n) { return CI(PI_COLS, n); };

  let done = 0;
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (ids.indexOf(String(row[a('分配ID')])) < 0) continue;
    if (truthy_(row[a('自動行')])) throw new Error('メイン担当の自動行は取消できません');
    if (truthy_(row[a('無効')])) continue;
    assertNotClosed_(row[a('請求月')]);

    // 確定済み案件の分配取消は管理者のみ(評価値が動くため)
    const piIdx = findPiRow_(piSh, row[a('案件ID')]);
    if (piIdx > 0) {
      const fixGp = piSh.getRange(piIdx, c('確定_粗利') + 1).getValue();
      if (numOrEmpty_(fixGp) !== '' && !isAdmin_()) {
        throw new Error('確定済み案件(' + row[a('案件ID')] + ')の分配取消は管理者のみ可能です');
      }
    }

    alSh.getRange(r + 1, a('無効') + 1).setValue(true);
    alSh.getRange(r + 1, a('取消者') + 1, 1, 2).setValues([[userEmail_(), now_()]]);
    audit_('分配取消', String(row[a('案件ID')]), String(row[a('分配ID')]),
      { メンバー: row[a('メンバー')], 種別: row[a('種別')], 金額: row[a('金額')], 制作pct: row[a('制作%')] }, '無効化');
    done++;
  }
  recalcAll_();
  return done + '件の分配を取消(論理削除)しました';
}

// ================= ヘルパー =================

/** 案件入力シートで案件IDの行番号(1始まり)。なければ -1 */
function findPiRow_(piSh, projId) {
  const ids = piSh.getRange(1, CI(PI_COLS, '案件ID') + 1, Math.max(piSh.getLastRow(), 1), 1).getValues();
  const target = String(projId).trim();
  for (let r = 1; r < ids.length; r++) {
    if (String(ids[r][0]).trim() === target) return r + 1;
  }
  return -1;
}

/** 案件の適用請求月(登録済みならPIの値、未登録なら自動導出値) */
function piYmOrAuto_(projId, proj) {
  const piSh = sheet_(SHEET.PI);
  const idx = findPiRow_(piSh, projId);
  if (idx > 0) return normYm_(piSh.getRange(idx, CI(PI_COLS, '適用請求月') + 1).getValue());
  return proj.ymAuto;
}

function requireNum_(v, label) {
  const n = Number(v);
  if (v === '' || v === null || v === undefined || !isFinite(n)) {
    throw new Error(label + ' は数値(税抜・円・整数)で入力してください');
  }
  return Math.round(n);
}
