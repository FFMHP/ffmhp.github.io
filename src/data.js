/* Nạp bảng dữ liệu dùng chung từ thư mục data/ cạnh trang web (DMKH, DMSP, Code chuyển đổi, Form final).
   Anh Tài chỉ cần thay file Excel cùng tên trên GitHub là cả team có bảng mới.
   Không tải được (mở file HTML trên máy) thì dùng bảng gắn sẵn trong trang. */
(function (root) {
  const BASE = 'data/';
  const FILES = { dmkh: 'DMKH.xlsx', dmsp: 'DMSP.xlsx', map: 'Code chuyển đổi.xlsx', tpl: 'Form final.xlsx' };
  const PREFIX = { KKTC: 'KKTC', MOR: 'MORGT', BRG: 'BRGMART', AEON: 'AEON', SOIBIEN: 'SOIBIEN' };
  const LABEL = { dmkh: 'DMKH', dmsp: 'DMSP', map: 'Code chuyển đổi', tpl: 'Form final' };
  const CACHE = 'ffm-data-v1:';
  // Google Apps Script: thống kê + bảng "Code AEON" team tự thêm
  const API = root.FFM_LOG_URL || 'https://script.google.com/macros/s/AKfycbxLBieSz9Cp00--Eade0ZZOoDQxs_2QduXNVWVec236x7ny-YgnR-6curOzhY79ZNnovQ/exec';
  root.FFM_API = API;
  const status = {}; // key -> {src:'web'|'gắn sẵn', at}
  const url = n => BASE + encodeURIComponent(n);
  const norm = s => String(s == null ? '' : s).normalize('NFC').trim();
  const lget = k => { try { return JSON.parse(localStorage.getItem(CACHE + k) || 'null'); } catch (e) { return null; } };
  const lset = (k, v) => { try { localStorage.setItem(CACHE + k, JSON.stringify(v)); } catch (e) {} };

  async function stamp(name) {
    const r = await fetch(url(name), { method: 'HEAD', cache: 'no-cache' });
    if (!r.ok) return null;
    return { tag: r.headers.get('etag') || r.headers.get('last-modified') || r.headers.get('content-length') || '?', mod: r.headers.get('last-modified') };
  }
  async function getBook(name) {
    const r = await fetch(url(name), { cache: 'no-cache' });
    if (!r.ok) throw new Error(name + ': HTTP ' + r.status);
    const buf = await r.arrayBuffer();
    return { buf, wb: name.endsWith('.xlsx') && !/Form final/.test(name) ? XLSX.read(buf, { type: 'array', dense: true, cellHTML: false, cellText: false }) : null };
  }
  const rowsOf = ws => XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
  const headerIdx = (rows, test) => rows.findIndex(r => r && test(r.map(x => norm(x).toLowerCase())));

  // DMKH: cột "Mã KH" + "Tên nội bộ", lọc theo tiền tố mã của từng account
  function parseDMKH(wb) {
    const rows = rowsOf(wb.Sheets[wb.SheetNames[0]]);
    const h = headerIdx(rows, r => r.includes('mã kh') && r.includes('tên nội bộ'));
    if (h < 0) throw new Error('DMKH: không thấy cột "Mã KH" và "Tên nội bộ"');
    const ci = rows[h].findIndex(x => norm(x).toLowerCase() === 'mã kh'), ni = rows[h].findIndex(x => norm(x).toLowerCase() === 'tên nội bộ');
    const out = {}; Object.keys(PREFIX).forEach(k => out[k] = []);
    for (const r of rows.slice(h + 1)) {
      const code = norm(r[ci]), name = norm(r[ni]);
      if (!code || !name) continue;
      for (const k in PREFIX) if (code.toUpperCase().startsWith(PREFIX[k])) out[k].push({ code, name });
    }
    if (!Object.values(out).some(a => a.length)) throw new Error('DMKH: không có mã KH nào của KKTC/MORGT/BRGMART/AEON/SOIBIEN');
    return out;
  }
  // DMSP: Barcode -> [[product code, tên, active]]; bỏ vật tư POSM có mã chữ, bỏ barcode lạ (>14 số)
  function parseDMSP(wb) {
    const rows = rowsOf(wb.Sheets[wb.SheetNames[0]]);
    const h = headerIdx(rows, r => r.includes('barcode') && r.includes('product code'));
    if (h < 0) throw new Error('DMSP: không thấy cột "Barcode" và "PRODUCT CODE"');
    const H = rows[h].map(x => norm(x).toLowerCase());
    const bi = H.indexOf('barcode'), pi = H.indexOf('product code'), ni = H.indexOf('product name');
    const gi = H.findIndex(x => x.startsWith('ngành hàng')), ai = H.indexOf('active product');
    const pc = {};
    for (const r of rows.slice(h + 1)) {
      const bc = String(r[bi] == null ? '' : r[bi]).replace(/\D/g, ''), code = norm(r[pi]);
      if (bc.length < 8 || bc.length > 14 || !code) continue;
      if (gi >= 0 && norm(r[gi]).toLowerCase() === 'posm' && !/^\d/.test(code)) continue;
      (pc[bc.padStart(14, '0')] = pc[bc.padStart(14, '0')] || []).push([code, norm(r[ni]), ai < 0 || /active/i.test(norm(r[ai])) ? 1 : 0]);
    }
    if (Object.keys(pc).length < 50) throw new Error('DMSP: quá ít barcode, kiểm tra lại file');
    return pc;
  }
  // Code chuyển đổi: sheet có tiêu đề "Barcode Aeon"
  function parseMap(wb) {
    const ws = wb.SheetNames.map(n => wb.Sheets[n]).find(s => rowsOf(s).some(r => /barcode\s*aeon/i.test(norm(r[0])))) || wb.Sheets[wb.SheetNames[0]];
    const map = {};
    rowsOf(ws).forEach(r => { const k = norm(r[0]).replace(/\.0+$/, ''), v = norm(r[1]); if (/^\d{6,14}$/.test(k) && v) map[k] = v; });
    return map;
  }

  function apply(key, val) {
    if (key === 'dmkh') { for (const k in val) { const a = root.FFM.ACCOUNTS[k]; if (a && val[k].length) a.customers.splice(0, a.customers.length, ...val[k]); } applyCust(); }
    if (key === 'dmsp') {
      const extra = root.FFM_PC_EXTRA || {};
      for (const k in extra) if (!val[k]) val[k] = extra[k];
      root.FFM_PC = val; root.FFM_DMSP = {}; for (const k in val) root.FFM_DMSP[k] = val[k][0][1]; root.FFM_CATALOG = Object.keys(val);
    }
    if (key === 'map') root.FFM_AEON_MAP = val;
    if (key === 'tpl') root.FFM_TPL = Uint8Array.from(atob(val), c => c.charCodeAt(0));
  }
  const b64 = buf => { let s = ''; const u = new Uint8Array(buf); for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); };
  const PARSE = { dmkh: b => parseDMKH(b.wb), dmsp: b => parseDMSP(b.wb), map: b => parseMap(b.wb), tpl: b => b64(b.buf) };

  async function loadOne(key) {
    const name = FILES[key];
    const st = await stamp(name);
    if (!st) { status[key] = { src: 'gắn sẵn' }; return; }
    const c = lget(key + '@2');
    if (c && c.tag === st.tag) { apply(key, c.val); status[key] = { src: 'web', at: c.mod }; return; }
    const val = PARSE[key](await getBook(name));
    lset(key + '@2', { tag: st.tag, mod: st.mod, val });
    apply(key, val); status[key] = { src: 'web', at: st.mod, fresh: true };
  }

  // Mã KH team thêm: gắn thêm vào danh sách cửa hàng của từng account (đánh dấu team)
  function applyCust() {
    const A = root.FFM && root.FFM.ACCOUNTS; if (!A) return;
    for (const k in A) { const list = A[k].customers; for (let i = list.length - 1; i >= 0; i--) if (list[i].team) list.splice(i, 1); }
    (root.FFM_CUST_SHARED || []).forEach(x => {
      const a = A[x.acct]; if (!a) return;
      if (!a.customers.some(c => c.code === x.code && norm(c.name).toLowerCase() === norm(x.name).toLowerCase())) a.customers.push({ code: x.code, name: x.name, team: true });
    });
  }
  // bảng Code AEON team thêm: không chặn quá 15 giây; tải xong muộn vẫn áp dụng
  let sharedP = null;
  function fetchShared() {
    return sharedP || (sharedP = fetch(API + '?t=aeonmap', { cache: 'no-store' }).then(r => r.json()).then(j => {
      if (!j || !j.map) throw new Error('bad');
      root.FFM_AEON_MAP_SHARED = j.map; lset('aeonShared', j.map); status.shared = { n: Object.keys(j.map).length };
      root.FFM_CUST_SHARED = j.cust || []; lset('custShared', root.FFM_CUST_SHARED); applyCust();
      return true;
    }).catch(e => { sharedP = null; throw e; }));
  }
  async function loadShared() {
    if (!/^https:/.test(API)) return;
    const c = lget('aeonShared'); if (c) root.FFM_AEON_MAP_SHARED = c;
    const cc = lget('custShared'); if (cc) { root.FFM_CUST_SHARED = cc; applyCust(); }
    const n0 = Object.keys(root.FFM_AEON_MAP_SHARED || {}).length;
    const timeout = new Promise(ok => setTimeout(() => ok('timeout'), 15000));
    const r = await Promise.race([fetchShared().catch(() => 'err'), timeout]);
    if (r !== true) {
      status.shared = { err: true, n: n0 };
      fetchShared().catch(() => fetchShared()).then(() => root.FFM_DATA && root.FFM_DATA.onShared && root.FFM_DATA.onShared()).catch(() => {});
    }
  }

  function summary() {
    const fmt = s => { if (!s) return 'đang tải…'; if (s.err) return 'lỗi, dùng bảng gắn sẵn'; if (s.src !== 'web') return 'bảng gắn sẵn'; if (!s.at) return 'từ web'; const d = new Date(s.at), p = n => String(n).padStart(2, '0'); return `cập nhật ${p(d.getDate())}/${p(d.getMonth() + 1)} ${p(d.getHours())}:${p(d.getMinutes())}`; };
    const nc = (root.FFM_CUST_SHARED || []).length;
    const sh = status.shared ? ` · Code AEON team thêm: ${status.shared.n} mã · Mã KH team thêm: ${nc}${status.shared.err ? ' (chưa tải mới được)' : ''}` : '';
    return Object.keys(FILES).map(k => `${LABEL[k]}: ${fmt(status[k])}`).join(' · ') + sh;
  }

  root.FFM_DATA = {
    status, summary, parseDMKH, parseDMSP, parseMap,
    loadShared, applyCust,
    ready: (location.protocol === 'file:' ? Promise.resolve() : Promise.all(Object.keys(FILES).map(k =>
      loadOne(k).catch(e => { console.warn('[FFM data]', k, e); status[k] = { src: 'gắn sẵn', err: String(e.message || e) }; })
    ).concat([loadShared()]))).then(() => { if (location.protocol === 'file:') Object.keys(FILES).forEach(k => status[k] = { src: 'gắn sẵn' }); })
  };
})(window);
