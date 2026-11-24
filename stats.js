/* 精準輔導 Excel → 地圖統計資料。所有計算規則都集中在這裡。
   規則（與使用者確認過）：
   - 只讀「精準輔導」工作表；宣導場次不讀。個資欄位（聯絡人、執行人、工程聯絡人、缺失內容…）一律不讀。
   - 一列 = 一場次。總場次 = 該區列數。
   - 是否假日：以 Excel 欄位為準（只接受「假日／非假日」）；另與政府辦公日曆表比對，不一致只提醒、不修改。
   - 缺失項目：連續數字為一個項次，其他符號（空白 . 、 。 ， 等）只是分隔；「無」或「0」= 無違規。
     全形數字、全形標點先轉成半形；「1-5」「1~5」這類範圍寫法列為錯誤（避免被讀成第 1、5 項）。
   - 局限空間地點、類型、行業別：名稱中間的空白不算；不在原有分類中的名稱會提醒（可能是打錯字）。
     項次需在 1～16；同格重複的項次只算一次。空白或無法判讀的格子是錯誤。
     以數字格式儲存的格子：小數部分為「1」者（如 2.1）可能原為 2.10，列為錯誤；其餘照數字讀。
   - 違規分級（以不重複項次數計）：5 項以上／小於 5 項（1～4）／無違規。
   - 違規公司分布：假日有違規場次／非假日有違規場次／無違規場次。
   - 缺失五項以上公司：單一場次缺失 ≥5 項（含 5 項）才列入，不同場次不累計；
     每個符合的場次各列一列（同一家公司有多場就列多列），欄位：承攬廠商、缺失項次總數、缺失項目、日期。
   - 局限空間類型：各區頁面與全區總覽都統計。
   - 是否偏遠：115 年起才統計；115 年以前的年度不讀這一欄（就算 Excel 有這一欄也不讀）。 */
(function (root) {
  'use strict';

  const SHEET = '精準輔導';
  const REMOTE_FROM = 115;  // 是否偏遠從 115 年起才統計
  const COLS = [            // [代號, 表頭判斷, 顯示名稱]
    ['date', h => h === '日期', '日期'],
    ['company', h => h === '承攬廠商', '承攬廠商'],
    ['district', h => h === '行政區', '行政區'],
    ['remote', h => h === '是否偏遠', '是否偏遠'],
    ['place', h => h === '局限空間地點', '局限空間地點'],
    ['type', h => h === '局限空間類型', '局限空間類型'],
    ['industry', h => h.startsWith('行業別'), '行業別'],
    ['items', h => h.startsWith('缺失項目'), '缺失項目'],
    ['holiday', h => h === '是否假日', '是否假日'],
  ];
  const ORDER = {
    place: ['儲槽', '機械設備內部', '人孔', '管道與地下空間', '蓄水池或化糞池', '密閉空間'],
    industry: ['水、電、燃氣業', '製造業', '化學工業', '環境衛生服務業', '建築營造', '電信業', '食品加工業', '下水道工程業', '其他政府機關'],
    type: ['缺氧作業', '局限空間'],
  };
  const MAX_ITEM = 16;

  const norm = s => String(s ?? '').replace(/　/g, ' ').trim();
  const pad = n => String(n).padStart(2, '0');

  function cellText(c) {
    if (!c) return '';
    if (c.t === 's') return norm(c.v);
    if (c.t === 'n') return c.raw;
    if (c.t === 'd') return `${c.v.y}-${pad(c.v.m)}-${pad(c.v.d)}`;
    if (c.t === 'b') return c.v ? 'TRUE' : 'FALSE';
    return '#' + c.v;
  }

  function parseDate(c) {
    if (!c) return null;
    if (c.t === 'd') return c.v;
    if (c.t === 's') {
      const m = norm(c.v).match(/^(\d{2,4})\s*[\/\-.年]\s*(\d{1,2})\s*[\/\-.月]\s*(\d{1,2})\s*日?$/);
      if (!m) return null;
      let y = +m[1]; if (y < 1911) y += 1911;
      const d = { y, m: +m[2], d: +m[3] };
      const t = new Date(Date.UTC(d.y, d.m - 1, d.d));
      return t.getUTCMonth() === d.m - 1 && t.getUTCDate() === d.d ? d : null;
    }
    return null;
  }

  /* 缺失項目 → {items:[...], none:bool} 或 {error} */
  function parseItems(c) {
    if (!c || (c.t === 's' && norm(c.v) === '')) return { error: '缺失項目空白（無違規請填「無」）' };
    if (c.t === 'n') {
      const raw = String(c.raw).trim();
      if (/^0+(\.0+)?$/.test(raw)) return { items: [], none: true };      // 0 = 無缺失
      if (!/^\d+(\.\d+)?$/.test(raw)) return { error: `無法判讀的數字「${raw}」` };
      const [a, b] = raw.split('.');
      if (b !== undefined && b.length > 2) return { error: `無法判讀的數字「${raw}」，請把此欄改成文字格式後重新輸入` };
      // 數字格式會吃掉小數末尾的 0：只有「x.1」可能原為「x.10」，無法分辨
      if (b === '1') return { error: `此格以數字格式儲存為 ${raw}，可能原本是「${a}.10」，無法判斷。請把此欄改成文字格式後重新輸入` };
      const nums = [+a]; if (b !== undefined && !/^0+$/.test(b)) nums.push(+b);
      return checkNums(nums, raw, true);
    }
    if (c.t !== 's') return { error: '缺失項目格式無法判讀' };
    const s = norm(c.v).normalize('NFKC');          // 全形數字、全形標點 → 半形
    if (s === '無' || /^0+$/.test(s)) return { items: [], none: true };
    if (s.includes('無')) return { error: `同一格同時有「無」與其他內容「${s}」` };
    const nums = (s.match(/\d+/g) || []).map(Number);
    if (!nums.length) return { error: `沒有可判讀的項次「${s}」` };
    if (/\d\s*[-~－～—–]\s*\d/.test(s)) return { error: `看起來是範圍寫法「${s}」，請逐項列出（例如 1、2、3、4、5）` };
    if (/[^\d\s.,、。，．;；:：\/]/.test(s)) return { error: `含有非分隔符號的文字「${s}」` };
    return checkNums(nums, s, false);
  }
  function checkNums(nums, text, fromNumber) {
    const bad = nums.filter(n => !(n >= 1 && n <= MAX_ITEM));
    if (bad.length) return { error: `項次 ${bad.join('、')} 不在 1～${MAX_ITEM} 之間（原內容「${text}」）` };
    const set = [...new Set(nums)].sort((a, b) => a - b);
    return { items: set, none: false, dup: set.length !== nums.length, fromNumber };
  }

  // 只列出當年度資料中出現過的分類：先依原簡報順序，新分類接在後面，「其他」放最後
  function orderCats(key, seen) {
    const base = ORDER[key] || [];
    const out = [...base.filter(c => seen.has(c)), ...[...seen].filter(c => !base.includes(c) && c !== '其他')];
    if (seen.has('其他')) out.push('其他');
    return out;
  }

  function analyze(wb, roc, calendar, districtNames) {
    const errors = [], warnings = [], notes = [];
    const err = (row, col, msg, val) => errors.push({ row, col, msg, val });
    const sheet = wb.sheets.find(s => s.name.trim() === SHEET);
    if (!sheet) return { errors: [{ row: '-', col: '-', msg: `找不到名為「${SHEET}」的工作表（目前有：${wb.sheets.map(s => s.name).join('、')}）` }], warnings, notes };
    const header = sheet.rows.get(1) || {};
    const colOf = {};
    const useRemote = roc >= REMOTE_FROM;
    if (!useRemote) notes.push({ row: '-', col: '是否偏遠', msg: `${roc} 年不統計是否偏遠（${REMOTE_FROM} 年起才加入），這一欄不會讀取` });
    for (const [key, test, label] of COLS) {
      if (key === 'remote' && !useRemote) continue;
      const hits = Object.entries(header).filter(([, c]) => c && c.t === 's' && test(norm(c.v).replace(/\s/g, '')));
      if (!hits.length) err(1, label, `第 1 列找不到「${label}」欄`);
      else if (hits.length > 1) err(1, label, `「${label}」欄出現 ${hits.length} 次，無法判斷要用哪一欄`);
      else colOf[key] = +hits[0][0];
    }
    if (errors.length) return { errors, warnings, notes };

    const valid = new Set(districtNames);
    const rows = [];
    for (const [rn, cells] of [...sheet.rows.entries()].sort((a, b) => a[0] - b[0])) {
      if (rn === 1) continue;
      const get = k => cells[colOf[k]];
      if (COLS.every(([k]) => colOf[k] === undefined || cellText(get(k)) === '')) continue;      // 整列空白（只看需要的欄位）
      const r = { row: rn };
      const date = parseDate(get('date'));
      if (!date) err(rn, '日期', '日期無法判讀', cellText(get('date')));
      r.date = date;
      r.company = cellText(get('company'));
      if (!r.company) err(rn, '承攬廠商', '承攬廠商空白');
      r.district = cellText(get('district')).replace(/\s/g, '');
      if (!valid.has(r.district)) err(rn, '行政區', '不是高雄市 38 個行政區之一', cellText(get('district')));
      r.remote = useRemote ? cellText(get('remote')) : null;
      if (useRemote && r.remote !== '是' && r.remote !== '否') err(rn, '是否偏遠', '只能填「是」或「否」', r.remote);
      r.holiday = cellText(get('holiday'));
      if (r.holiday !== '假日' && r.holiday !== '非假日') err(rn, '是否假日', '只能填「假日」或「非假日」', r.holiday);
      for (const k of ['place', 'type', 'industry']) {
        r[k] = cellText(get(k)).replace(/\s/g, '');            // 分類名稱中間的空白不算
        if (!r[k]) err(rn, COLS.find(c => c[0] === k)[2], '空白');
      }
      const it = parseItems(get('items'));
      if (it.error) err(rn, '缺失項目', it.error, cellText(get('items')));
      else {
        r.items = it.items;
        if (it.dup) warnings.push({ row: rn, col: '缺失項目', msg: '同一格有重複的項次，已只算一次', val: cellText(get('items')) });
        if (it.fromNumber) notes.push({ row: rn, col: '缺失項目', msg: `此格是數字格式，已讀為第 ${it.items.join('、')} 項；建議改成文字格式`, val: cellText(get('items')) });
      }
      rows.push(r);
    }
    if (!rows.length) errors.push({ row: '-', col: '-', msg: '「精準輔導」工作表沒有資料' });

    // ---- 提醒（不影響統計） ----
    const westYear = roc + 1911;
    const offYear = rows.filter(r => r.date && r.date.y !== westYear);
    if (offYear.length) warnings.push({ row: offYear.map(r => r.row).join('、'), col: '日期', msg: `共 ${offYear.length} 筆日期不在 ${roc} 年（${westYear} 年）內，仍會計入 ${roc} 年統計，請確認` });
    const keyCount = {};
    for (const r of rows) if (r.date) { const k = `${r.date.y}-${r.date.m}-${r.date.d}|${r.company}|${r.district}`; (keyCount[k] = keyCount[k] || []).push(r.row); }
    for (const [k, list] of Object.entries(keyCount)) if (list.length > 1)
      warnings.push({ row: list.join('、'), col: '日期／承攬廠商／行政區', msg: '日期、承攬廠商、行政區完全相同，可能重複登打（仍各算一場）', val: k.split('|').slice(1).join('／') });
    const remoteBy = {};
    for (const r of rows) if (r.remote === '是' || r.remote === '否') (remoteBy[r.district] = remoteBy[r.district] || new Set()).add(r.remote);
    for (const [d, s] of Object.entries(remoteBy)) if (s.size > 1)
      warnings.push({ row: rows.filter(r => r.district === d).map(r => r.row).join('、'), col: '是否偏遠', msg: `${d}同時出現「是」與「否」，地圖上不會標示此區是否偏遠` });
    const missingCal = new Set();
    for (const r of rows) {
      if (!r.date || (r.holiday !== '假日' && r.holiday !== '非假日')) continue;
      const cal = calendar.years[r.date.y];
      if (!cal) { missingCal.add(r.date.y); continue; }
      const iso = `${r.date.y}-${pad(r.date.m)}-${pad(r.date.d)}`;
      const wd = new Date(Date.UTC(r.date.y, r.date.m - 1, r.date.d)).getUTCDay();
      const off = cal.holidays.includes(iso) || ((wd === 0 || wd === 6) && !cal.workdays.includes(iso));
      if ((r.holiday === '假日') !== off)
        warnings.push({ row: r.row, col: '是否假日', msg: `填「${r.holiday}」，但政府辦公日曆表此日為${off ? '放假日' : '上班日'}（星期${'日一二三四五六'[wd]}）`, val: iso });
    }
    for (const y of missingCal) notes.push({ row: '-', col: '是否假日', msg: `網站尚未收錄 ${y} 年（民國 ${y - 1911} 年）政府辦公日曆表，該年日期略過假日比對` });

    if (errors.length) return { errors, warnings, notes, rows: rows.length };

    // ---- 統計 ----
    const seen = { place: new Set(), industry: new Set(), type: new Set() };
    for (const r of rows) for (const k in seen) seen[k].add(r[k]);
    const colName = { place: '局限空間地點', industry: '行業別', type: '局限空間類型' };
    for (const k in seen) for (const v of seen[k]) if (v !== '其他' && !ORDER[k].includes(v))
      warnings.push({ row: rows.filter(r => r[k] === v).map(r => r.row).join('、'), col: colName[k], msg: `「${v}」不是原本的分類，會成為一個新類別；如果是打錯字，請修正後重新上傳`, val: v });
    const cats = { place: orderCats('place', seen.place), industry: orderCats('industry', seen.industry), type: orderCats('type', seen.type) };
    const zero = list => Object.fromEntries(list.map(k => [k, 0]));
    const blank = () => ({
      '場次分布': zero(['假日', '非假日']),
      '違規公司分布': zero(['假日', '非假日', '無違規']),
      '假日違規': zero(['五項以上', '小於五項', '無違規']),
      '非假日違規': zero(['五項以上', '小於五項', '無違規']),
      '局限空間地點': zero(cats.place),
      '行業別': zero(cats.industry),
      '局限空間類型': zero(cats.type),
    });
    const level = n => n >= 5 ? '五項以上' : n >= 1 ? '小於五項' : '無違規';
    function add(ch, r) {
      ch['場次分布'][r.holiday]++;
      ch['違規公司分布'][r.items.length ? r.holiday : '無違規']++;
      ch[r.holiday === '假日' ? '假日違規' : '非假日違規'][level(r.items.length)]++;
      ch['局限空間地點'][r.place]++;
      ch['行業別'][r.industry]++;
      ch['局限空間類型'][r.type]++;
    }
    const districts = {};
    for (const name of districtNames) districts[name] = { total: 0, remote: null, charts: blank() };
    const sum = blank(); if (useRemote) sum['是否偏遠'] = zero(['偏遠', '非偏遠']);
    for (const r of rows) {
      const d = districts[r.district];
      d.total++; add(d.charts, r);
      add(sum, r); if (useRemote) sum['是否偏遠'][r.remote === '是' ? '偏遠' : '非偏遠']++;
    }
    for (const [name, s] of Object.entries(remoteBy)) if (s.size === 1) districts[name].remote = [...s][0] === '是';
    for (const d of Object.values(districts)) if (!d.total) d.charts = {};

    const itemCount = zero(['無', ...Array.from({ length: MAX_ITEM }, (_, i) => `第${i + 1}項`)]);
    for (const r of rows) { if (!r.items.length) itemCount['無']++; for (const n of r.items) itemCount[`第${n}項`]++; }
    const companies = rows.filter(r => r.items.length >= 5)
      .map(r => [r.company, r.items.length, r.items.join('、'), `${r.date.y}-${pad(r.date.m)}-${pad(r.date.d)}`])
      .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], 'zh-Hant') || a[3].localeCompare(b[3]));

    const data = {
      year: roc, source: 'upload', generated: new Date().toISOString(), rows: rows.length,
      districts,
      summary: { total: rows.length, charts: sum, items: Object.entries(itemCount).map(([name, v]) => ({ name, v })), companies },
    };
    return { errors, warnings, notes, rows: rows.length, data };
  }

  const api = { analyze, parseItems };
  if (typeof module !== 'undefined') module.exports = api; else root.MapStats = api;
})(typeof window !== 'undefined' ? window : globalThis);
