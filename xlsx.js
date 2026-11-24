/* 精簡的 .xlsx 讀取器：只讀儲存格的值，不依賴任何外部套件。
   在瀏覽器與 Node.js（18 以上）都可執行。 */
(function (root) {
  'use strict';
  const td = new TextDecoder('utf-8');

  async function inflateRaw(bytes) {
    const ds = new DecompressionStream('deflate-raw');
    const stream = new Blob([bytes]).stream().pipeThrough(ds);
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  async function unzip(buf) {
    const u8 = new Uint8Array(buf), dv = new DataView(buf);
    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('這不是有效的 .xlsx 檔（找不到壓縮目錄）');
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const files = {};
    for (let n = 0; n < count; n++) {
      if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('壓縮目錄損毀');
      const method = dv.getUint16(p + 10, true);
      const csize = dv.getUint32(p + 20, true);
      const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
      const lho = dv.getUint32(p + 42, true);
      const name = td.decode(u8.subarray(p + 46, p + 46 + nlen));
      const lnlen = dv.getUint16(lho + 26, true), lxlen = dv.getUint16(lho + 28, true);
      const start = lho + 30 + lnlen + lxlen;
      const data = u8.subarray(start, start + csize);
      files[name] = { method, data };
      p += 46 + nlen + xlen + clen;
    }
    return {
      has: name => name in files,
      async text(name) {
        const f = files[name];
        if (!f) return null;
        const raw = f.method === 0 ? f.data : f.method === 8 ? await inflateRaw(f.data) : null;
        if (!raw) throw new Error('不支援的壓縮方式：' + name);
        return td.decode(raw);
      }
    };
  }

  /* 極簡 XML 解析：只取需要的元素與屬性 */
  function attrs(s) {
    const o = {}; const re = /([\w:]+)="([^"]*)"/g; let m;
    while ((m = re.exec(s))) o[m[1]] = m[2];
    return o;
  }
  function unesc(s) {
    return s.replace(/&(lt|gt|amp|quot|apos|#\d+|#x[0-9a-fA-F]+);/g, (_, e) =>
      e === 'lt' ? '<' : e === 'gt' ? '>' : e === 'amp' ? '&' : e === 'quot' ? '"' : e === 'apos' ? "'" :
      e[1] === 'x' ? String.fromCodePoint(parseInt(e.slice(2), 16)) : String.fromCodePoint(parseInt(e.slice(1), 10)));
  }
  function textOf(xml) {           // 串接 <t> 內容；略過 <rPh>（注音／假名標示）
    xml = xml.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
    let out = ''; const re = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>|<t(?:\s[^>]*)?\/>/g; let m;
    while ((m = re.exec(xml))) out += m[1] ? unesc(m[1]) : '';
    return out;
  }

  const BUILTIN_DATE = new Set([14, 15, 16, 17, 22, 27, 28, 29, 30, 31, 34, 35, 36, 50, 51, 52, 53, 54, 55, 56, 57, 58]);
  function isDateFormat(code) {
    const c = code.replace(/"[^"]*"|\[[^\]]*\]|\\./g, '');
    return /[ymd]/i.test(c) && !/^[#0.,%\s]*$/.test(c);
  }
  function colIndex(ref) {
    const letters = ref.match(/^[A-Z]+/)[0]; let n = 0;
    for (const ch of letters) n = n * 26 + ch.charCodeAt(0) - 64;
    return n - 1;
  }
  function serialToDate(v) {        // Excel 1900 日期系統序號 → {y,m,d}
    const ms = Math.round((v - 25569) * 86400000);
    const d = new Date(ms);
    return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate() };
  }

  async function readXlsx(buf) {
    const z = await unzip(buf);
    const wbXml = await z.text('xl/workbook.xml');
    if (!wbXml) throw new Error('這不是 Excel 活頁簿（缺少 workbook.xml）');
    if (/date1904="(1|true)"/.test(wbXml)) throw new Error('此檔案使用 1904 日期系統，請改用一般 Excel 日期格式');
    const rels = {};
    const relXml = await z.text('xl/_rels/workbook.xml.rels') || '';
    for (const m of relXml.matchAll(/<Relationship\b([^>]*)\/?>/g)) { const a = attrs(m[1]); rels[a.Id] = a.Target; }
    const sst = [];
    const sstXml = await z.text('xl/sharedStrings.xml');
    if (sstXml) for (const m of sstXml.matchAll(/<si>([\s\S]*?)<\/si>|<si\/>/g)) sst.push(m[1] ? textOf(m[1]) : '');
    const dateStyles = new Set();
    const stXml = await z.text('xl/styles.xml');
    if (stXml) {
      const custom = {};
      for (const m of stXml.matchAll(/<numFmt\b([^>]*)\/>/g)) { const a = attrs(m[1]); custom[a.numFmtId] = unesc(a.formatCode || ''); }
      const xfs = (stXml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/) || [])[1] || '';
      let i = 0;
      for (const m of xfs.matchAll(/<xf\b([^>]*?)(\/>|>)/g)) {
        const id = +(attrs(m[1]).numFmtId || 0);
        if (BUILTIN_DATE.has(id) || (custom[id] && isDateFormat(custom[id]))) dateStyles.add(i);
        i++;
      }
    }
    const sheets = [];
    for (const m of wbXml.matchAll(/<sheet\b([^>]*)\/?>/g)) {
      const a = attrs(m[1]);
      let target = rels[a['r:id']] || '';
      target = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
      const xml = await z.text(target);
      const rows = new Map();
      if (xml) {
        for (const rm of xml.matchAll(/<row\b([^>]*)>([\s\S]*?)<\/row>/g)) {
          const rn = +attrs(rm[1]).r; const cells = {};
          for (const cm of rm[2].matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
            const ca = attrs(cm[1]); const inner = cm[2] || '';
            const vm = inner.match(/<v>([\s\S]*?)<\/v>/);
            const raw = vm ? unesc(vm[1]) : null;
            let cell = null;
            if (ca.t === 's') cell = { t: 's', v: sst[+raw] ?? '' };
            else if (ca.t === 'inlineStr') cell = { t: 's', v: textOf(inner) };
            else if (ca.t === 'str') cell = { t: 's', v: raw ?? '' };
            else if (ca.t === 'b') cell = { t: 'b', v: raw === '1' };
            else if (ca.t === 'e') cell = { t: 'e', v: raw };
            else if (raw !== null) {
              const num = Number(raw);
              cell = dateStyles.has(+(ca.s || 0)) ? { t: 'd', v: serialToDate(num), raw } : { t: 'n', v: num, raw };
            }
            if (cell) cells[colIndex(ca.r)] = cell;
          }
          rows.set(rn, cells);
        }
      }
      sheets.push({ name: unesc(a.name), rows });
    }
    return { sheets };
  }

  const api = { readXlsx };
  if (typeof module !== 'undefined') module.exports = api; else root.XLSXLite = api;
})(typeof window !== 'undefined' ? window : globalThis);
