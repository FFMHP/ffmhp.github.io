/* FFM converter core — works in browser (globals XLSX, JSZip) and Node (tests). */
(function (root) {
  const XLSXlib = root.XLSX || (typeof require !== 'undefined' ? require('xlsx') : null);
  const JSZipLib = root.JSZip || (typeof require !== 'undefined' ? require('jszip') : null);

  const norm = s => String(s == null ? '' : s).normalize('NFC').replace(/ /g, ' ').trim();
  const low = s => norm(s).toLowerCase();
  const pad = n => String(n).padStart(2, '0');
  const fmtDate = d => `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()}`;
  const toISO = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const addDays = (d, n) => { const x = new Date(d.getFullYear(), d.getMonth(), d.getDate()); x.setDate(x.getDate() + n); return x; };
  const excelSerial = iso => { const [y, m, d] = iso.split('-').map(Number); return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000); };

  function cellText(ws, r, c) {
    const cell = ws[XLSXlib.utils.encode_cell({ r, c })];
    if (!cell || cell.v == null) return '';
    if (cell.t === 'n' && Math.abs(cell.v) >= 1e6) return String(Math.round(cell.v)); // barcode stored as number
    return norm(cell.w != null ? cell.w : cell.v);
  }
  function cellNum(ws, r, c) {
    const cell = ws[XLSXlib.utils.encode_cell({ r, c })];
    if (!cell || cell.v == null || norm(cell.v) === '') return null;
    if (cell.t === 'n') return cell.v;
    const s = norm(cell.v).replace(/[.,\s]/g, '');
    return /^-?\d+$/.test(s) ? Number(s) : NaN;
  }

  // so khớp tên cửa hàng: bỏ dấu, bỏ tiền tố thương hiệu, bỏ ký tự đặc biệt
  const fold = s => norm(s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase();
  const KKTC_BRAND = /^(kktc_?|k[\s-]*mar(t|ket|tket)|kmarket|kmart)\b[\s_-]*/;
  const key = (s, brand) => { let k = fold(s).replace(/[^a-z0-9]+/g, ' ').trim(); if (!brand) return k; let p; do { p = k; k = k.replace(brand, '').trim(); } while (k !== p); return k; };
  function matchCustomer(list, name, brand) {
    const k = key(name, brand); if (!k) return { hit: null };
    const exact = list.filter(c => key(c.name, brand) === k);
    if (exact.length === 1) return { hit: exact[0], exact: true };
    // tên team đã gán tay trước đó được ưu tiên
    const teamHit = exact.filter(c => c.team);
    if (teamHit.length) return { hit: teamHit[teamHit.length - 1], exact: true };
    // tên trong DMKH nằm trọn trong tên trên file (vd "83 nguyen an ninh" chứa "nguyen an ninh") và chỉ có 1 cửa hàng như vậy
    const toks = new Set(k.split(' '));
    const sub = list.filter(c => { const ct = key(c.name, brand).split(' ').filter(Boolean); return ct.length >= 2 && ct.every(t => toks.has(t)); });
    // ngược lại: tên trên file nằm trọn trong tên DMKH (vd "mao khe" trong "mao khe quang ninh")
    const sup = toks.size >= 2 ? list.filter(c => { const ct = new Set(key(c.name, brand).split(' ')); return [...toks].every(t => ct.has(t)); }) : [];
    if (!sub.length && sup.length === 1) return { hit: sup[0], exact: false };
    if (sub.length === 1) { const ct = new Set(key(sub[0].name, brand).split(' ')); const extra = [...toks].filter(t => !ct.has(t)); return { hit: sub[0], exact: extra.every(t => /^\d{2,}$/.test(t)) }; }
    const compact = x => x.replace(/ /g, '');
    const kc = compact(k);
    const exact2 = list.filter(c => compact(key(c.name, brand)) === kc);
    if (exact2.length === 1) return { hit: exact2[0], exact: true };
    // độ giống theo cặp ký tự (chịu được viết dính, sai chính tả nhẹ: SUNSHINECITY ~ S3 Shunshine City)
    const grams = x => { const g = {}; for (let i = 0; i < x.length - 1; i++) { const t = x.slice(i, i + 2); g[t] = (g[t] || 0) + 1; } return g; };
    const dice = (a, b) => { const ga = grams(a), gb = grams(b); let inter = 0, na = 0, nb = 0; for (const t in ga) { na += ga[t]; if (gb[t]) inter += Math.min(ga[t], gb[t]); } for (const t in gb) nb += gb[t]; return na + nb ? 2 * inter / (na + nb) : 0; };
    const scored = list.map(c => ({ c, s: dice(kc, compact(key(c.name, brand))) })).sort((a, b) => b.s - a.s);
    if (scored[0] && scored[0].s >= 0.6 && (!scored[1] || scored[0].s - scored[1].s >= 0.1)) return { hit: scored[0].c, exact: false, score: scored[0].s };
    return { hit: null };
  }
  const yymmdd = iso => iso.slice(2, 4) + iso.slice(5, 7) + iso.slice(8, 10);
  /* Số đơn = prefix + viết tắt cửa hàng + YYMMDD ngày về đơn + -NN (thứ tự trong ngày, theo cửa hàng) */
  function assignPO(results, ngayVe, startSeq) {
    const count = {};
    results.forEach(r => {
      if (r.error || r.fixedPo) return;
      const d = r.poDate || ngayVe;
      if ((!r.abbr && !r.noAbbr) || !d) { r.po = ''; return; }
      const k = r.prefix + (r.abbr || '') + (r.poDate ? yymmdd(d) : '');
      count[k] = (count[k] || (startSeq && startSeq[k]) || 0) + 1;
      const base = r.abbrAfterDate ? `${r.prefix}${yymmdd(d)}${r.abbr || ''}` : `${r.prefix}${r.abbr || ''}${yymmdd(d)}`;
      // seq 'always': luôn có -01, -02 (Kmart) · 'dup': đơn đầu không đuôi, đơn thứ 2 trở đi -02, -03 (Morgant)
      r.po = r.seq === 'dup' ? (count[k] === 1 ? base : `${base}-${pad(count[k])}`) : `${base}-${pad(count[k])}`;
      r.poNote = '';
    });
  }

  // tiêu đề cột barcode: "Code", "Barcode", "Mã hàng barcode", "Mã vạch", "EAN"...
  const isBarcodeHeader = h => { const f = fold(h).trim(); return f === 'code' || /barcode|bar code|ma vach|\bean\b/.test(f); };
  /* ---------- Parser chung cho đơn dạng "1 cột số lượng theo cửa hàng" ----------
     Dòng tiêu đề: cột barcode (Code/Barcode) | tên SP | (giá) | cột SL có tiêu đề = tên cửa hàng | (mã KH) */
  function parseStoreColumn(acct, wb) {
    const ws = wb.Sheets[wb.SheetNames[0]];
    const range = XLSXlib.utils.decode_range(ws['!ref'] || 'A1:A1');
    const issues = [];
    let hr = -1, colCode = -1;
    for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 15) && hr < 0; r++)
      for (let c = range.s.c; c <= Math.min(range.e.c, 30); c++)
        if (isBarcodeHeader(cellText(ws, r, c))) { hr = r; colCode = c; break; }
    if (hr < 0) return { error: `Không tìm thấy cột "Code"/"Barcode" ở dòng tiêu đề. File này có đúng mẫu đơn ${acct.label} không?` };
    let colName = -1, colPrice = -1, colCust = -1, colStt = -1, store = '', customer = '';
    const skip = []; // cột giá / thành tiền / ĐVT... không phải cột số lượng
    for (let c = range.s.c; c <= Math.min(range.e.c, 30); c++) {
      const h = cellText(ws, hr, c), l = low(h);
      if (!h || c === colCode) continue;
      if (['stt', 'no', 'no.', '#'].includes(l)) colStt = c;
      else if (l.startsWith('tên') || l.includes('name')) colName = c;
      else if (/(^|\s|\()(đơn giá|giá|price|thành tiền|amount|vat|đvt|unit|quy cách|tổng)(\s|\)|$)/.test(l) || l.startsWith('giá') || l.startsWith('đơn giá') || l.startsWith('thành tiền')) { colPrice = colPrice < 0 ? c : colPrice; skip.push(c); }
      else if (/^[A-Z]{3,}\d{3,}$/.test(h.toUpperCase()) && colCust < 0) { colCust = c; customer = h.toUpperCase(); }
    }
    // cột số lượng = cột có tiêu đề là tên cửa hàng: ưu tiên cột dò được trong DMKH, rồi cột ngay trái cột mã KH, rồi cột còn lại đầu tiên
    const cands = [];
    for (let c = range.s.c; c <= Math.min(range.e.c, 30); c++) {
      if ([colCode, colName, colPrice, colCust, colStt].includes(c) || skip.includes(c)) continue;
      if (cellText(ws, hr, c)) cands.push(c);
    }
    let colQty = cands.find(c => matchCustomer(acct.customers, cellText(ws, hr, c), acct.brand).hit);
    if (colQty == null && colCust >= 0 && cands.includes(colCust - 1)) colQty = colCust - 1;
    if (colQty == null) colQty = cands.length ? cands[0] : -1;
    if (colQty >= 0) store = cellText(ws, hr, colQty);
    if (colQty < 0) return { error: 'Không tìm thấy cột số lượng (cột có tiêu đề là tên cửa hàng).' };
    const m = matchCustomer(acct.customers, store, acct.brand);
    const fileCust = customer;
    customer = m.hit ? m.hit.code : (fileCust || '');
    if (!m.hit && !fileCust) issues.push({ key: 'customer', level: 'err', msg: `Không dò được "${store}" trong DMKH (cột Tên nội bộ). Chọn cửa hàng ở ô Mã khách hàng.` });
    if (!m.hit && fileCust) issues.push({ level: 'warn', msg: `"${store}" không có trong DMKH; đang dùng mã KH ghi trong file (${fileCust}).` });
    if (m.hit && !m.exact) issues.push({ level: 'warn', msg: `Khớp gần đúng: "${store}" → ${m.hit.code} · ${m.hit.name}. Kiểm tra lại.` });
    if (m.hit && fileCust && fileCust !== m.hit.code) issues.push({ level: 'warn', msg: `File ghi mã KH ${fileCust}, DMKH ra ${m.hit.code}. Đang dùng mã theo DMKH.` });
    const ab = acct.makeAbbr(store, customer);
    if (!ab.abbr) issues.push({ key: 'abbr', level: 'err', msg: 'Không tạo được viết tắt cửa hàng. Nhập tay viết tắt.' });
    if (ab.note) issues.push({ key: 'abbrNote', level: 'warn', msg: ab.note });

    const lines = [];
    for (let r = hr + 1; r <= range.e.r; r++) {
      const q = cellNum(ws, r, colQty);
      if (q == null || q === 0) continue;
      let bc = cellText(ws, r, colCode);
      const name = colName >= 0 ? cellText(ws, r, colName) : '';
      const ln = { row: r + 1, barcode: bc, name, qty: q, flags: [] };
      if (/^\d{11}$/.test(bc)) { ln.barcode = bc = '0' + bc; ln.flags.push({ level: 'info', msg: 'Đã thêm số 0 đầu (file gốc mất số 0)' }); }
      if (Number.isNaN(q)) ln.flags.push({ level: 'err', msg: `Số lượng không phải số: "${cellText(ws, r, colQty)}"` });
      else if (q < 0 || !Number.isInteger(q)) ln.flags.push({ level: 'err', msg: 'Số lượng âm hoặc lẻ thập phân' });
      if (!bc) ln.flags.push({ level: 'err', msg: 'Thiếu barcode' });
      else if (!/^\d{8}$|^\d{12,14}$/.test(bc)) ln.flags.push({ level: 'warn', msg: 'Barcode lạ (không phải 8/12/13/14 số)' });
      else if (root.FFM_DMSP && !root.FFM_DMSP[bc.padStart(14, '0')]) ln.flags.push({ level: 'warn', msg: 'Barcode không có trong danh mục sản phẩm Hyperion (DMSP)' });
      lines.push(ln);
    }
    const seen = {};
    lines.forEach(l => { if (l.barcode) { if (seen[l.barcode]) l.flags.push({ level: 'warn', msg: `Trùng barcode với dòng ${seen[l.barcode]}` }); else seen[l.barcode] = l.row; } });
    if (!lines.length) issues.push({ level: 'err', msg: 'Không có dòng nào có số lượng đặt hàng.' });
    return { prefix: acct.prefix, seq: acct.seq, abbr: ab.abbr, customer, store, lines, issues };
  }

  /* ---------- Sói Biển: đơn Excel có sẵn Product code ----------
     Tiêu đề trên cùng: "ĐƠN HÀNG KHO TỔNG 2 ĐẶT NGÀY 12/11" · cột Barcode | Product code | Sản phẩm | ĐVT | TỔNG KHO ĐẶT | Giá
     Product code điền cột D, SL cột F, kho -> mã KH theo DMKH ("KHO 2" -> SOIBIEN074). */
  function parseSoiBien(acct, wb, fileName) {
    const ws = wb.Sheets[wb.SheetNames[0]];
    const range = XLSXlib.utils.decode_range(ws['!ref'] || 'A1:A1');
    const issues = [];
    let hr = -1, cBc = -1, cPc = -1, cName = -1, cQty = -1, title = '';
    for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 20) && hr < 0; r++) {
      const H = []; for (let c = range.s.c; c <= Math.min(range.e.c, 30); c++) H[c] = fold(cellText(ws, r, c)).trim();
      const pc = H.findIndex(h => h && /product\s*code|ma\s*(sp|san pham|hang hyperion)/.test(h));
      if (pc >= 0) {
        hr = r; cPc = pc;
        cBc = H.findIndex(h => h && /barcode|ma vach/.test(h));
        cName = H.findIndex(h => h && /san pham|ten/.test(h));
        cQty = H.findIndex((h, i) => h && i !== cPc && /dat|so luong|\bsl\b/.test(h) && !/gia|tien/.test(h));
      } else { for (let c = range.s.c; c <= Math.min(range.e.c, 30); c++) { const t = cellText(ws, r, c); if (t && t.length > title.length) title = t; } }
    }
    if (hr < 0) return { error: 'Không tìm thấy cột "Product code" ở dòng tiêu đề. File này có đúng mẫu đơn Sói Biển không?' };
    if (cQty < 0) return { error: 'Không tìm thấy cột số lượng đặt (vd "TỔNG KHO ĐẶT").' };
    const src = title + ' ' + fileName;
    // ngày đặt: "ĐẶT NGÀY 12/11" trong tiêu đề, không có thì lấy "ĐẶT 12.11" trong tên file
    let orderDate = '';
    const md = fold(title).match(/ngay\s*(\d{1,2})\s*[\/.\-]\s*(\d{1,2})(?:\s*[\/.\-]\s*(\d{2,4}))?/) || fold(fileName).match(/dat\s*(?:ngay\s*)?(\d{1,2})\s*[\/.\-]\s*(\d{1,2})(?:\s*[\/.\-]\s*(\d{2,4}))?/);
    if (md) {
      const now = new Date(); let y = md[3] ? +md[3] : now.getFullYear(); if (y < 100) y += 2000;
      if (!md[3] && +md[2] > now.getMonth() + 2) y--; // đơn tháng 12 mở vào tháng 1 năm sau
      orderDate = `${y}-${String(md[2]).padStart(2, '0')}-${String(md[1]).padStart(2, '0')}`;
    } else issues.push({ level: 'warn', msg: 'Không đọc được ngày đặt trên đơn; số đơn và ngày giao tính theo ô Ngày về đơn.' });
    // kho -> mã KH
    const mk = fold(src).match(/kho\s*(?:tong\s*)?(\d+)/);
    let hit = null;
    if (mk) { const re = new RegExp(`kho\\s*(?:tong\\s*)?0*${mk[1]}\\b`); const c = acct.customers.filter(x => re.test(fold(x.name))); if (c.length === 1) hit = c[0]; }
    if (!hit) { const m = matchCustomer(acct.customers, title, /^(don hang|soi bien)\b[\s_-]*/); if (m.hit && m.exact) hit = m.hit; }
    if (!hit) issues.push({ key: 'customer', level: 'err', msg: `Không dò được cửa hàng/kho từ "${title.trim() || fileName}". Chọn ở ô Mã khách hàng.` });
    const known = new Set(root.FFM_PC ? Object.values(root.FFM_PC).flat().map(x => x[0]) : []);
    const lines = [];
    for (let r = hr + 1; r <= range.e.r; r++) {
      const q = cellNum(ws, r, cQty);
      let pc = cellText(ws, r, cPc).trim().toUpperCase(), bc = cBc >= 0 ? cellText(ws, r, cBc).replace(/\D/g, '') : '';
      if (!pc && !bc) continue;
      if (q == null || q === 0) continue;
      const ln = { row: r + 1, barcode: bc, productCode: pc, qty: q, flags: [] };
      const nm = cName >= 0 ? cellText(ws, r, cName) : '';
      if (!pc && bc && root.FFM_PC && root.FFM_PC[bc.padStart(14, '0')]) {
        const c = root.FFM_PC[bc.padStart(14, '0')];
        if (c.length === 1) { ln.productCode = pc = c[0][0]; ln.flags.push({ level: 'warn', msg: `Đơn không ghi Product code; lấy theo DMSP: ${pc}` }); }
      }
      if (!pc) ln.flags.push({ level: 'err', msg: 'Thiếu Product code' });
      else if (known.size && !known.has(pc)) ln.flags.push({ level: 'warn', msg: `Product code ${pc} không có trong DMSP. Kiểm tra lại mã.` });
      if (Number.isNaN(q)) ln.flags.push({ level: 'err', msg: `Số lượng không phải số: "${cellText(ws, r, cQty)}"` });
      else if (q < 0 || !Number.isInteger(q)) ln.flags.push({ level: 'err', msg: 'Số lượng âm hoặc lẻ thập phân' });
      const dm = pc && root.FFM_PC ? (Object.values(root.FFM_PC).flat().find(x => x[0] === pc) || [])[1] : '';
      ln.name = `${pc || '?'} · ${dm || nm}`;
      lines.push(ln);
    }
    const seen = {};
    lines.forEach(l => { if (l.productCode) { if (seen[l.productCode]) l.flags.push({ level: 'warn', msg: `Trùng Product code với dòng ${seen[l.productCode]}` }); else seen[l.productCode] = l.row; } });
    if (!lines.length) issues.push({ level: 'err', msg: 'Không có dòng nào có số lượng đặt hàng.' });
    return { prefix: acct.prefix, seq: acct.seq, abbr: '', noAbbr: true, pcMode: true, poDate: orderDate, orderDate, customer: hit ? hit.code : '', store: title.trim(), lines, issues };
  }

  /* ---------- An Nhiên: 1 file nhiều sheet, mỗi sheet có số lượng = 1 đơn ----------
     Tên sheet có ngày đặt "ASM-12-11-26" · cột BARCODE + cột SL · mã KH cố định ANHNHCM0001 */
  function parseAnNhien(acct, wb, fileName) {
    const out = [], skipped = [];
    for (const sn of wb.SheetNames) {
      const ws = wb.Sheets[sn];
      if (!ws || !ws['!ref']) { skipped.push(`${sn} (trống)`); continue; }
      const range = XLSXlib.utils.decode_range(ws['!ref']);
      let hr = -1, cBc = -1, cQty = -1, cName = -1;
      for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 20) && hr < 0; r++) {
        const H = []; for (let c = range.s.c; c <= Math.min(range.e.c, 40); c++) H[c] = fold(cellText(ws, r, c)).replace(/_/g, ' ').trim();
        const b = H.findIndex(h => h && isBarcodeHeader(h));
        const q = H.findIndex(h => h && /^(sl|so luong|sl dat|so luong dat)$/.test(h));
        if (b >= 0 && q >= 0) { hr = r; cBc = b; cQty = q; cName = H.findIndex(h => h && /ten san pham|san pham|ten hang|vietnamese name/.test(h)); }
      }
      if (hr < 0) { skipped.push(`${sn} (không có cột SL)`); continue; }
      const lines = [];
      for (let r = hr + 1; r <= range.e.r; r++) {
        const q = cellNum(ws, r, cQty);
        if (q == null || q === 0) continue;
        let bc = cellText(ws, r, cBc).replace(/\D/g, '');
        const ln = { row: r + 1, barcode: bc, qty: q, name: cName >= 0 ? cellText(ws, r, cName) : '', flags: [] };
        if (/^\d{11}$/.test(bc)) { ln.barcode = bc = '0' + bc; ln.flags.push({ level: 'info', msg: 'Đã thêm số 0 đầu (file gốc mất số 0)' }); }
        if (Number.isNaN(q)) ln.flags.push({ level: 'err', msg: `Số lượng không phải số: "${cellText(ws, r, cQty)}"` });
        else if (q < 0 || !Number.isInteger(q)) ln.flags.push({ level: 'err', msg: 'Số lượng âm hoặc lẻ thập phân' });
        if (!bc) ln.flags.push({ level: 'err', msg: 'Thiếu barcode' });
        else if (!/^\d{8}$|^\d{12,14}$/.test(bc)) ln.flags.push({ level: 'warn', msg: 'Barcode lạ (không phải 8/12/13/14 số)' });
        else if (root.FFM_DMSP && !root.FFM_DMSP[bc.padStart(14, '0')]) ln.flags.push({ level: 'warn', msg: 'Barcode không có trong danh mục sản phẩm Hyperion (DMSP)' });
        lines.push(ln);
      }
      if (!lines.length) { skipped.push(`${sn} (không có dòng nào có SL)`); continue; }
      const seen = {};
      lines.forEach(l => { if (l.barcode) { if (seen[l.barcode]) l.flags.push({ level: 'warn', msg: `Trùng barcode với dòng ${seen[l.barcode]}` }); else seen[l.barcode] = l.row; } });
      const issues = [];
      let orderDate = '';
      const md = sn.match(/(\d{1,2})\s*[-.\/_]\s*(\d{1,2})\s*[-.\/_]\s*(\d{2,4})/);
      if (md) { let y = +md[3]; if (y < 100) y += 2000; orderDate = `${y}-${md[2].padStart(2, '0')}-${md[1].padStart(2, '0')}`; }
      else issues.push({ level: 'warn', msg: `Tên sheet "${sn}" không có ngày đặt (dạng 12-11-26); số đơn và ngày giao tính theo ô Ngày về đơn.` });
      out.push({ prefix: acct.prefix, seq: acct.seq, abbr: '', noAbbr: true, poDate: orderDate, orderDate, customer: acct.fixedCustomer, store: sn, sheet: sn, lines, issues });
    }
    if (!out.length) return { error: `Không sheet nào có số lượng đặt hàng${skipped.length ? ' (' + skipped.join(', ') + ')' : ''}.` };
    if (skipped.length) out[0].issues.push({ level: 'info', msg: `Bỏ qua sheet không phải đơn: ${skipped.join(', ')}.` });
    if (out.length > 1) out.forEach(o => o.issues.push({ level: 'info', msg: `File có ${out.length} sheet đặt hàng → ${out.length} đơn riêng.` }));
    return { multi: out };
  }

  /* ---------- Miomio: PO dạng PDF (PURCHASE ORDER) ----------
     Số đơn = "PO No." trên phiếu · ngày đặt = "Order time" · mã KH cố định MIOMIOW0004
     SL điền form = số cuối cùng trong cột Specification × cột Purchase Quantity
     (vd "1 thùng 24 hộp * 12 cái" × 30 = 360). Barcode điền cột C (anh Tài chốt 13/11). */
  function parseMiomio(acct, items, fileName) {
    const issues = [];
    const all = items.map(x => x.t).join(' ');
    const mPo = all.match(/PO\s*No\.?\s*[：:]?\s*(PO\d{6,})/i) || all.match(/\b(PO\d{10,})\b/) || fileName.match(/(PO\d{10,})/i);
    const po = mPo ? mPo[1].toUpperCase() : '';
    if (!po) issues.push({ key: 'po', level: 'err', msg: 'Không đọc được PO No. trên phiếu. Nhập tay Số đơn.' });
    const mD = all.match(/Order\s*time\s*[：:]?\s*(\d{4})-(\d{2})-(\d{2})/i);
    const orderDate = mD ? `${mD[1]}-${mD[2]}-${mD[3]}` : '';
    if (!orderDate) issues.push({ level: 'warn', msg: 'Không đọc được Order time; ngày giao tính theo ô Ngày về đơn.' });
    // cột theo tiêu đề
    const H = re => items.find(x => re.test(x.t));
    const hBc = H(/^piece\s*barcode$|^barcode$/i) || H(/barcode/i), hSpec = H(/^specification$/i), hDesc = H(/description/i);
    const hValid = H(/^valid/i), hPur = H(/^purchase$/i) || H(/purchase/i), hUnit = items.find(x => /^unit$/i.test(x.t) && hPur && x.x0 > hPur.x1);
    if (!hSpec || !hPur) return { error: 'Không tìm thấy cột Specification / Purchase Quantity trên PO. File này có đúng mẫu PO Miomio không?' };
    const specL = hDesc ? (hDesc.x1 + hSpec.x0) / 2 : hSpec.x0 - 20;
    const specR = hValid ? hValid.x0 - 2 : hSpec.x1 + 20;
    const purL = hValid ? hValid.x1 + 2 : hPur.x0 - 15, purR = hUnit ? hUnit.x0 - 2 : hPur.x1 + 15;
    const hdrY = hSpec.y1;
    const bcs = items.filter(x => /^\d{8,14}$/.test(x.t) && x.y0 > hdrY && (!hBc || x.x0 < (hDesc ? hDesc.x0 : hSpec.x0))).sort((a, b) => a.y0 - b.y0);
    const totalIt = items.find(x => /^total$/i.test(x.t) && x.y0 > hdrY);
    const lines = [];
    bcs.forEach((b, i) => {
      const top = b.y0 - 14, bot = i + 1 < bcs.length ? bcs[i + 1].y0 - 14 : (totalIt ? totalIt.y0 - 2 : b.y0 + 120);
      const band = items.filter(x => x.y0 >= top && x.y0 < bot);
      const spec = band.filter(x => x.x0 >= specL && x.x0 < specR).sort((a, c) => (a.y0 - c.y0) || (a.x0 - c.x0));
      let specTxt = '';
      spec.forEach(x => { specTxt += specTxt && !(/\d$/.test(specTxt) && /^\d/.test(x.t)) ? ' ' + x.t : x.t; });
      const desc = band.filter(x => hDesc && x.x0 >= hDesc.x0 - 40 && x.x0 < specL && !/^D\d{6,}$/.test(x.t)).sort((a, c) => (a.y0 - c.y0) || (a.x0 - c.x0)).map(x => x.t).join(' ').replace(/\s+/g, ' ');
      const pur = band.filter(x => x.x0 >= purL && x.x1 <= purR + 10 && /^\d+([.,]\d+)?$/.test(x.t));
      const pq = pur.length ? Number(pur[0].t.replace(',', '.')) : NaN;
      const nums = specTxt.match(/\d+/g) || [];
      const last = nums.length ? Number(nums[nums.length - 1]) : NaN;
      const ln = { row: i + 1, barcode: b.t, qty: (last && pq) ? last * pq : NaN, flags: [], spec: specTxt, pq, last };
      if (!specTxt || Number.isNaN(last)) ln.flags.push({ level: 'err', msg: `Không đọc được số trong cột Specification ("${specTxt}"). Nhập tay SL.` });
      if (Number.isNaN(pq)) ln.flags.push({ level: 'err', msg: 'Không đọc được Purchase Quantity. Nhập tay SL.' });
      // anh Tài chốt 13/11: Miomio điền BARCODE vào cột C (không điền product code); DMSP chỉ để đối chiếu tên
      const cands = root.FFM_PC ? (root.FFM_PC[b.t.padStart(14, '0')] || []) : [];
      if (!cands.length) ln.flags.push({ level: 'warn', msg: 'Barcode không có trong danh mục sản phẩm Hyperion (DMSP).' });
      if (root.FUJI && !root.FUJI.gtinOk(b.t)) ln.flags.push({ level: 'warn', msg: 'Barcode sai số kiểm tra' });
      ln.flags.push({ level: 'info', msg: `SL = ${Number.isNaN(last) ? '?' : last} (số cuối "${specTxt}") × ${Number.isNaN(pq) ? '?' : pq}` });
      const dm = cands.length ? cands[0][1] : '';
      ln.name = dm || desc;
      lines.push(ln);
    });
    if (!lines.length) return { error: 'Không thấy dòng hàng nào (barcode) trên PO.' };
    // đối chiếu tổng Purchase Quantity với dòng Total
    if (totalIt) {
      const tq = items.filter(x => Math.abs(x.y0 - totalIt.y0) < 6 && x.x0 >= purL - 10 && x.x1 <= purR + 15 && /^\d+$/.test(x.t));
      const sumPq = lines.reduce((s, l) => s + (l.pq || 0), 0);
      if (tq.length && Number(tq[0].t) !== sumPq) issues.push({ level: 'err', msg: `Tổng Purchase Quantity đọc được (${sumPq}) khác dòng Total trên PO (${tq[0].t}). Có thể thiếu hoặc đọc sai dòng.` });
      else if (tq.length) issues.push({ level: 'info', msg: `Khớp tổng Purchase Quantity với dòng Total: ${sumPq}.` });
    }
    return { fixedPo: true, editable: true, po, customer: acct.fixedCustomer, store: 'Miomio (Busy Ming)', orderDate, lines, issues, source: 'pdf' };
  }


  /* ================= 7 acc thêm 13/11: Aeon Fantasy, Fresh Garden, Hiền Lương, Hinode, Japaus, Sakuko, Siba Food ================= */
  // dòng hàng chuẩn: kiểm SL, barcode, DMSP; tên lấy theo DMSP (không có thì tên trên đơn)
  function stdLine(row, bcRaw, qty, srcName) {
    let bc = String(bcRaw == null ? '' : bcRaw).replace(/\D/g, '');
    const ln = { row, barcode: bc, qty, flags: [] };
    if (/^\d{11}$/.test(bc)) { ln.barcode = bc = '0' + bc; ln.flags.push({ level: 'info', msg: 'Đã thêm số 0 đầu (file gốc mất số 0)' }); }
    if (qty == null || Number.isNaN(qty)) ln.flags.push({ level: 'err', msg: 'Không đọc được số lượng. Nhập tay SL.' });
    else if (qty < 0 || !Number.isInteger(qty)) ln.flags.push({ level: 'err', msg: 'Số lượng âm hoặc lẻ thập phân' });
    if (!bc) ln.flags.push({ level: 'err', msg: 'Thiếu barcode' });
    else if (!/^\d{8}$|^\d{12,14}$/.test(bc)) ln.flags.push({ level: 'warn', msg: 'Barcode lạ (không phải 8/12/13/14 số)' });
    else if (root.FFM_DMSP && !root.FFM_DMSP[bc.padStart(14, '0')]) ln.flags.push({ level: 'warn', msg: 'Barcode không có trong danh mục sản phẩm Hyperion (DMSP)' });
    const dm = root.FFM_DMSP && bc ? root.FFM_DMSP[bc.padStart(14, '0')] : '';
    ln.name = dm || norm(srcName).replace(/\s+/g, ' ');
    return ln;
  }
  function dupCheck(lines, issues) {
    const seen = {};
    lines.forEach(l => { if (l.barcode) { if (seen[l.barcode]) l.flags.push({ level: 'warn', msg: `Trùng barcode với dòng ${seen[l.barcode]}` }); else seen[l.barcode] = l.row; } });
    if (!lines.length) issues.push({ level: 'err', msg: 'Không có dòng nào có số lượng đặt hàng.' });
  }
  const dmy = s => { const m = String(s || '').match(/(\d{1,2})\s*[\/.\-]\s*(\d{1,2})\s*[\/.\-]\s*(\d{4})/); return m && +m[3] > 1990 ? `${m[3]}-${pad(+m[2])}-${pad(+m[1])}` : ''; };
  // dò cửa hàng theo địa chỉ/tên trên đơn: tỉ lệ từ trong tên DMKH (bỏ thương hiệu) xuất hiện trong chữ trên đơn
  const ADDR_EXPAND = [[/\bcong nghiep\b/g, 'cn'], [/\bkhu cong nghiep\b/g, 'kcn'], [/\bthanh pho\b/g, 'tp'], [/\bho chi minh\b/g, 'hcm']];
  function matchAddr(list, text, brandWords) {
    let t = ' ' + key(text) + ' ';
    ADDR_EXPAND.forEach(([re, ab]) => { if (re.test(t)) t += ab + ' '; });
    const toks = new Set(t.trim().split(' '));
    const scored = list.map(c => {
      const ct = key(c.name).split(' ').filter(w => w && !brandWords.includes(w));
      if (!ct.length) return { c, s: 0, n: 0 };
      const hit = ct.filter(w => toks.has(w)).length;
      // số nhà phải khớp (34 Vũ Phạm Hàm ≠ 126 …)
      const numsOk = ct.filter(w => /^\d+$/.test(w)).every(w => toks.has(w) || toks.has(String(+w)));
      return { c, s: numsOk ? hit / ct.length : 0, n: hit };
    }).sort((a, b) => (b.s - a.s) || (b.n - a.n));
    const [a, b] = scored;
    if (a && a.s >= 0.75 && a.n >= 1 && (!b || b.s < a.s || b.n < a.n)) return { hit: a.c, exact: a.s === 1 };
    return { hit: null };
  }
  // chữ tắt cửa hàng: chữ cái đầu các từ trong tên DMKH (bỏ thương hiệu, bỏ số): "SAKUKO MART HÀNG BÔNG" -> HB
  const initials = (name, brandWords) => key(name).split(' ').filter(w => w && !brandWords.includes(w) && !/^\d+$/.test(w)).map(w => w[0]).join('').toUpperCase();
  function findHeader(ws, test, maxRows) {
    const range = XLSXlib.utils.decode_range(ws['!ref'] || 'A1:A1');
    for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + (maxRows || 25)); r++) {
      const H = []; for (let c = range.s.c; c <= Math.min(range.e.c, 40); c++) H[c] = fold(cellText(ws, r, c)).replace(/\s+/g, ' ').trim();
      const m = test(H); if (m) return Object.assign({ row: r, H, range }, m);
    }
    return null;
  }
  const rowText = (ws, r, range) => { const a = []; for (let c = range.s.c; c <= Math.min(range.e.c, 40); c++) { const t = cellText(ws, r, c); if (t) a.push(t); } return a.join(' '); };

  /* Excel 1 sheet, 1 cột barcode + 1 cột SL; phần đầu có địa chỉ giao, ngày đặt */
  function parseSimpleSheet(acct, wb, fileName, o) {
    const ws = wb.Sheets[wb.SheetNames[0]];
    const h = findHeader(ws, H => {
      const b = H.findIndex(x => x && o.bcHead.test(x)), q = H.findIndex(x => x && o.qtyHead.test(x));
      return b >= 0 && q >= 0 ? { cBc: b, cQty: q, cName: H.findIndex(x => x && /ten (san pham|sp|hang)/.test(x)) } : null;
    });
    if (!h) return { error: `Không thấy dòng tiêu đề (cột mã hàng/barcode và cột số lượng). File này có đúng mẫu đơn ${acct.label} không?` };
    const issues = [], top = [];
    for (let r = h.range.s.r; r < h.row; r++) top.push(rowText(ws, r, h.range));
    const head = top.join('\n');
    const lines = [];
    let sumFile = null;
    for (let r = h.row + 1; r <= h.range.e.r; r++) {
      const raw = cellText(ws, r, h.cBc), q = cellNum(ws, r, h.cQty);
      const label = fold(rowText(ws, r, h.range));
      if (!raw.replace(/\D/g, '')) { if (/^tong\b|\btong cong\b/.test(label) && q != null) sumFile = q; continue; }
      if (q == null || q === 0) continue;
      lines.push(stdLine(r + 1, raw, q, h.cName >= 0 ? cellText(ws, r, h.cName) : ''));
    }
    dupCheck(lines, issues);
    const sum = lines.reduce((s, l) => s + (l.qty || 0), 0);
    if (sumFile != null) issues.push(sumFile === sum ? { level: 'info', msg: `Khớp tổng SL với dòng TỔNG trên đơn: ${sum}.` } : { level: 'err', msg: `Tổng SL đọc được (${sum}) khác dòng TỔNG trên đơn (${sumFile}).` });
    return { head, lines, issues };
  }
  // mã KH: dò theo địa chỉ giao trên đơn; mã ghi tay trên file (vd FGARDEN0001) chỉ để đối chiếu
  function pickCustomer(acct, text, issues, codeRe) {
    const m = matchAddr(acct.customers, text, acct.brandWords || []);
    const written = (text.match(codeRe) || [])[0];
    let hit = m.hit;
    if (!hit && written) hit = acct.customers.find(c => c.code.toUpperCase() === written.toUpperCase()) || null;
    if (hit && written && hit.code.toUpperCase() !== written.toUpperCase()) issues.push({ key: 'customer', level: 'warn', msg: `Trên đơn có ghi mã ${written} nhưng địa chỉ giao khớp ${hit.code} (${hit.name}). Kiểm tra lại mã KH.` });
    if (!hit) issues.push({ key: 'customer', level: 'err', msg: `Không dò được cửa hàng từ địa chỉ "${text.replace(/\s+/g, ' ').slice(0, 90)}". Chọn ở ô Mã khách hàng (rồi bấm Ghi nhớ).` });
    else if (!m.exact && !written) issues.push({ key: 'customer', level: 'warn', msg: `Mã KH dò gần đúng: ${hit.code} (${hit.name}). Kiểm tra lại.` });
    return hit;
  }

  function parseFreshGarden(acct, wb, fileName) {
    const s = parseSimpleSheet(acct, wb, fileName, { bcHead: /^ma hang$|barcode|ma vach/, qtyHead: /^so luong/ });
    if (s.error) return s;
    const orderDate = dmy((fold(s.head).match(/ngay dat[^\n]*/) || [''])[0]);
    const addr = (s.head.split('\n').find(l => /dia chi/.test(fold(l))) || s.head);
    const hit = pickCustomer(acct, addr, s.issues, /FGARDEN\d{3,}/i);
    if (!orderDate) s.issues.push({ level: 'warn', msg: 'Không đọc được "Ngày đặt hàng" trên đơn.' });
    if (orderDate) s.issues.push({ level: 'info', msg: `Ngày đặt trên đơn: ${orderDate.split('-').reverse().join('/')}.` });
    return { prefix: acct.prefix, seq: acct.seq, abbr: '', noAbbr: true, customer: hit ? hit.code : '', store: addr.replace(/^.*?:\s*/, ''), lines: s.lines, issues: s.issues };
  }
  function parseSakuko(acct, wb, fileName) {
    const s = parseSimpleSheet(acct, wb, fileName, { bcHead: /^ma sp$|barcode|ma vach/, qtyHead: /^sl dat|^so luong/ });
    if (s.error) return s;
    const addr = (s.head.split('\n').find(l => /dia chi/.test(fold(l))) || s.head);
    const hit = pickCustomer(acct, addr, s.issues, /SAKUKO\d{3,}/i);
    const abbr = hit ? initials(hit.name, acct.brandWords) : '';
    if (hit && !abbr) s.issues.push({ key: 'po', level: 'err', msg: 'Không tạo được chữ tắt cửa hàng cho số đơn. Nhập tay Số đơn.' });
    const need = dmy((s.head.split('\n').find(l => /can giao|ngay giao/.test(fold(l))) || ''));
    if (need) s.issues.push({ level: 'info', msg: `Đơn ghi thời gian cần giao ${need.split('-').reverse().join('/')}; ngày giao trên form = ngày về đơn + ${acct.leadDays} ngày làm việc.` });
    return { prefix: acct.prefix, seq: acct.seq, abbr, customer: hit ? hit.code : '', store: addr.replace(/^.*?:\s*/, ''), lines: s.lines, issues: s.issues };
  }
  // Hiền Lương: file .xls cũ, font TCVN3 (chữ Việt hiển thị lỗi) -> dò theo vị trí: dòng có barcode 8-14 số, SL = số đầu tiên sau cột ĐVT
  function parseHienLuong(acct, wb, fileName) {
    const ws = wb.Sheets[wb.SheetNames[0]];
    const range = XLSXlib.utils.decode_range(ws['!ref'] || 'A1:A1');
    const issues = [];
    let cBc = -1, cQty = -1, orderDate = '', soPhieu = '';
    for (let r = range.s.r; r <= Math.min(range.e.r, 60); r++) for (let c = range.s.c; c <= Math.min(range.e.c, 120); c++) {
      const t = cellText(ws, r, c);
      if (cBc < 0 && /^barcode$/i.test(t.trim())) cBc = c;
      if (cQty < 0 && /^S\.L\s*\n?\s*§Æt/i.test(t)) cQty = c;
      if (!orderDate && /^\d{1,2}\/\d{1,2}\/\d{4}/.test(t) && !/1900/.test(t)) orderDate = dmy(t);
      if (!soPhieu && /^0{2,}\d{6,}$/.test(t.trim())) soPhieu = t.trim();
    }
    if (cBc < 0) return { error: 'Không thấy cột BARCODE. File này có đúng mẫu phiếu đặt hàng Hiền Lương (.xls) không?' };
    // ô gộp làm tiêu đề lệch cột so với dữ liệu -> mỗi dòng: tìm ô barcode (8-14 số), SL = ô số đầu tiên sau ô ĐVT
    const lines = [];
    for (let r = range.s.r; r <= range.e.r; r++) {
      let cb = -1;
      for (let c = range.s.c; c <= Math.min(range.e.c, 120) && cb < 0; c++) { const cell = ws[XLSXlib.utils.encode_cell({ r, c })], t = cellText(ws, r, c).trim(); if (cell && /^\d{8,14}$/.test(t) && !/^000/.test(t) && (cell.t === 's' || t.length >= 12)) cb = c; }
      if (cb < 0) continue;
      let q = null, unit = '';
      for (let c = cb + 1; c <= Math.min(range.e.c, cb + 40); c++) {
        const t = cellText(ws, r, c).trim(); if (!t) continue;
        const n = cellNum(ws, r, c);
        if (n == null || Number.isNaN(n)) { if (!unit) unit = t; continue; }
        q = n; break;
      }
      if (q === 0) continue;
      lines.push(stdLine(r + 1, cellText(ws, r, cb), q, ''));
    }
    dupCheck(lines, issues);
    if (!orderDate) issues.push({ level: 'info', msg: 'Không đọc được ngày đặt trên phiếu.' });
    if (soPhieu) issues.push({ level: 'info', msg: `Số phiếu Hiền Lương: ${soPhieu}.` });
    if (orderDate) issues.push({ level: 'info', msg: `Ngày đặt trên phiếu: ${orderDate.split('-').reverse().join('/')}.` });
    return { prefix: acct.prefix, seq: acct.seq, abbr: '', noAbbr: true, customer: acct.fixedCustomer, store: 'Siêu thị Hiền Lương', lines, issues };
  }

  /* ---------- PDF dạng bảng: gán từng chữ/số vào cột gần nhất theo tiêu đề; mỗi dòng hàng neo theo ô số lượng ---------- */
  function pdfTable(items, o) {
    // o.heads: {name: item tiêu đề}; o.qty, o.bc: tên cột; o.top: y dưới tiêu đề; o.stop: y dòng tổng (không có thì hết trang)
    const H = Object.entries(o.heads).filter(([, it]) => it).map(([k, it]) => ({ k, cx: (it.x0 + it.x1) / 2 }));
    const colOf = it => { const cx = (it.x0 + it.x1) / 2; let best = null, d = 1e9; H.forEach(h => { const dd = Math.abs(h.cx - cx); if (dd < d) { d = dd; best = h.k; } }); return best; };
    items = mergeRuns(items);
    const body = items.filter(it => it.y0 > o.top && (!o.stop || it.y0 < o.stop) && !(o.skip && o.skip(it)));
    const anchors = body.filter(it => colOf(it) === o.qty && /^\d{1,3}(?:[.,]\d{3})*(?:,\d+)?$|^\d+(?:[.,]\d+)?$/.test(it.t)).sort((a, b) => a.y0 - b.y0);
    return anchors.map((a, i) => {
      const lo = i ? (anchors[i - 1].y0 + a.y0) / 2 : o.top, hi = i + 1 < anchors.length ? (a.y0 + anchors[i + 1].y0) / 2 : (o.stop || a.y0 + 40);
      const band = body.filter(it => it.y0 >= lo && it.y0 < hi).sort((x, y) => (x.y0 - y.y0) || (x.x0 - y.x0));
      const cell = k => band.filter(it => colOf(it) === k);
      return { anchor: a, cell, bc: cell(o.bc).filter(it => /^\d+$/.test(it.t)).map(it => it.t).join(''), name: o.name ? cell(o.name).map(it => it.t).join(' ') : '' };
    });
  }
  // ghép chữ/số bị PDF cắt rời ("5" "," "00" -> "5,00")
  function mergeRuns(items) {
    const a = items.slice().sort((x, y) => (Math.round(x.y0) - Math.round(y.y0)) || (x.x0 - y.x0)), out = [];
    a.forEach(it => { const p = out[out.length - 1]; if (p && Math.abs(p.y0 - it.y0) < 1 && it.x0 - p.x1 < 0.8 && it.x0 >= p.x0) { p.t += it.t; p.x1 = it.x1; } else out.push(Object.assign({}, it)); });
    return out;
  }
  const pdfQty = t => { t = String(t).trim(); if (/^\d+,\d{1,2}$/.test(t)) return Number(t.replace(',', '.')); if (/^\d+\.\d{1,2}$/.test(t)) return Number(t); return Number(t.replace(/[.,]/g, '')); };
  const itemAfter = (items, re, maxDx) => { const l = items.find(x => re.test(x.t)); if (!l) return null; return items.filter(x => Math.abs(x.y0 - l.y0) < 3 && x.x0 >= l.x1 - 2 && x.x0 - l.x1 < (maxDx || 200)).sort((a, b) => a.x0 - b.x0)[0] || null; };
  const allText = items => items.slice().sort((a, b) => (a.y0 - b.y0) || (a.x0 - b.x0)).map(x => x.t).join(' ');
  function pdfLines(rows, issues) {
    const lines = rows.map((r, i) => {
      const ln = stdLine(i + 1, r.bc, pdfQty(r.anchor.t), r.name);
      if (root.FUJI && ln.barcode && /^\d{8}$|^\d{12,14}$/.test(ln.barcode) && !root.FUJI.gtinOk(ln.barcode)) ln.flags.push({ level: 'warn', msg: 'Barcode sai số kiểm tra (có thể đọc thiếu số)' });
      return ln;
    });
    dupCheck(lines, issues);
    return lines;
  }

  function parseAeonFantasy(acct, items, fileName) {
    const issues = [];
    const F = re => items.find(x => re.test(x.t));
    const hBc = F(/^barcode$/i) || F(/^mã hàng$/i), hQty = F(/^qty$/i) || F(/^số lượng$/i);
    if (!hBc || !hQty) return { error: 'Không thấy cột Barcode / Qty trên PO. File này có đúng mẫu đơn đặt hàng Aeon Fantasy không?' };
    const heads = { stt: F(/^no\.?$/i) || F(/^stt$/i), bc: hBc, name: F(/^item name$/i) || F(/^tên hàng$/i), unit: F(/^unit$/i), qty: hQty, price: F(/^unit price$/i) || F(/^đơn giá$/i) };
    const tot = F(/^tổng thanh toán|^total$/i);
    const rows = pdfTable(items, { heads, qty: 'qty', bc: 'bc', name: 'name', top: Math.max(hBc.y1, hQty.y1), stop: tot ? tot.y0 - 1 : 0 });
    if (!rows.length) return { error: 'Không thấy dòng hàng nào trên PO Aeon Fantasy.' };
    const lines = pdfLines(rows, issues);
    const no = (itemAfter(items, /^số\/no\.?:?$/i) || {}).t || '';
    const orderDate = dmy((itemAfter(items, /^ngày\/date:?$/i) || {}).t);
    const storeId = (itemAfter(items, /store id:?$/i) || {}).t || '';
    const storeName = (itemAfter(items, /store name:?$/i) || {}).t || '';
    const addr = items.filter(x => x.y0 > ((F(/address delivery/i) || {}).y0 || 1e9) && x.y0 < ((F(/LƯU Ý ĐỊA CHỈ/i) || {}).y0 || 0) && x.x0 > 300).map(x => x.t).join(' ');
    let hit = null; const m = matchAddr(acct.customers, `${storeId} ${storeName} ${addr}`, acct.brandWords);
    if (m.hit) hit = m.hit; else issues.push({ key: 'customer', level: 'err', msg: `Không dò được cửa hàng "${storeName || storeId}". Chọn ở ô Mã khách hàng (rồi bấm Ghi nhớ).` });
    // chữ tắt cửa hàng: Store ID "AEON HUE-PRIZE" -> HUE
    const abbr = key(storeId).replace(/^aeon\s*/, '').split(' ')[0].toUpperCase() || key(storeName).split(' ').pop().toUpperCase();
    if (tot) {
      const tq = items.filter(x => Math.abs(x.y0 - tot.y0) < 3 && Math.abs((x.x0 + x.x1) / 2 - (hQty.x0 + hQty.x1) / 2) < 30).map(x => x.t.replace(/\D/g, '')).filter(Boolean)[0];
      const sum = lines.reduce((s, l) => s + (l.qty || 0), 0);
      if (tq) issues.push(+tq === sum ? { level: 'info', msg: `Khớp tổng SL với dòng Tổng thanh toán: ${sum}.` } : { level: 'err', msg: `Tổng SL đọc được (${sum}) khác dòng Tổng trên PO (${tq}).` });
    }
        if (no) issues.push({ level: 'info', msg: `Số PO Aeon Fantasy: ${no}.` });
    return { prefix: acct.prefix, seq: acct.seq, abbr, docDate: orderDate, customer: hit ? hit.code : '', store: storeName || storeId, lines, issues, source: 'pdf' };
  }
  function parseHinode(acct, items, fileName) {
    const issues = [];
    const F = re => items.find(x => re.test(x.t));
    const hBc = F(/^barcode$/i), hQty = items.filter(x => /^số lượn/i.test(x.t)).sort((a, b) => a.x0 - b.x0)[0];
    if (!hBc || !hQty) return { error: 'Không thấy cột Barcode / Số lượng trên PO. File này có đúng mẫu đơn đặt hàng Hinode (IMF) không?' };
    const heads = { stt: F(/^stt$/i), item: F(/^mã hàng$/i), bc: hBc, name: F(/^tên hàng$/i), unit: F(/^đvt$/i), qty: hQty, price: F(/^đơn giá/i) };
    const tot = F(/^giá trị chưa vat/i);
    const rows = pdfTable(items, { heads, qty: 'qty', bc: 'bc', name: 'name', top: hBc.y1, stop: tot ? tot.y0 - 1 : 0, skip: it => /^quầy hàng/i.test(it.t) });
    if (!rows.length) return { error: 'Không thấy dòng hàng nào trên PO Hinode.' };
    const lines = pdfLines(rows, issues);
    const txt = allText(items);
    const mPo = txt.match(/SỐ\s*:\s*(POD\d{6,})/i) || txt.match(/\b(POD\d{8,})\b/) || fileName.match(/(POD\d{8,})/i);
    const po = mPo ? mPo[1].toUpperCase() : '';
    if (!po) issues.push({ key: 'po', level: 'err', msg: 'Không đọc được số PO (POD…). Nhập tay Số đơn.' });
    const md = txt.match(/NGÀY\s*:\s*(\d{1,2}\/\d{1,2}\/\d{4})/i);
    const orderDate = md ? dmy(md[1]) : '';
        const giao = dmy((itemAfter(items, /^thời gian giao hàng:?$/i) || {}).t);
    if (giao) issues.push({ level: 'info', msg: `Phiếu ghi thời gian giao ${giao.split('-').reverse().join('/')}; ngày giao trên form = ngày về đơn + ${acct.leadDays} ngày làm việc.` });
    return { fixedPo: true, editable: true, po, customer: acct.fixedCustomer, store: 'Hinode Hoài Đức', docDate: orderDate, lines, issues, source: 'pdf' };
  }
  function parseJapaus(acct, items, fileName) {
    const issues = [];
    const F = re => items.find(x => re.test(x.t));
    const hBc = F(/^mã vạch$/i), hQty = F(/^sl$/i);
    if (!hBc || !hQty) return { error: 'Không thấy cột Mã vạch / SL trên phiếu. File này có đúng mẫu phiếu đặt hàng Japaus không?' };
    const heads = { stt: F(/^stt$/i), sku: F(/^mã sku$/i), bc: hBc, name: F(/^tên hàng/i), unit: F(/^đvt$/i), qty: hQty, price: F(/^đơn giá$/i), amt: F(/^thành tiền$/i) };
    const tot = F(/^cộng tiền hàng/i);
    // bỏ tiêu đề lặp lại ở trang 2 và chữ "1/2"
    const skip = it => /^(stt|mã sku|mã vạch|tên hàng hoá|đvt|sl|đơn giá|thành tiền|%|ck|thuế suất|\d\/\d)$/i.test(it.t);
    const rows = pdfTable(items, { heads, qty: 'qty', bc: 'bc', name: 'name', top: hQty.y1, stop: tot ? tot.y0 - 1 : 0, skip });
    if (!rows.length) return { error: 'Không thấy dòng hàng nào trên phiếu Japaus.' };
    const lines = pdfLines(rows, issues);
    const txt = allText(items);
    const mPo = txt.match(/Số\s*:\s*(PDH)\s*(\d{3,})/i) || txt.match(/\b(PDH)(\d{3,})\b/i);
    const po = mPo ? (mPo[1] + mPo[2]).toUpperCase() : ((fileName.match(/pdh\d+/i) || [''])[0].toUpperCase());
    if (!po) issues.push({ key: 'po', level: 'err', msg: 'Không đọc được số phiếu (PDH…). Nhập tay Số đơn.' });
    const md = txt.match(/Ngày\s*(\d{1,2})\s*tháng\s*(\d{1,2})\s*năm\s*(\d{4})/i);
    const orderDate = md ? `${md[3]}-${pad(+md[2])}-${pad(+md[1])}` : '';
        return { fixedPo: true, editable: true, po, customer: acct.fixedCustomer, store: 'Japaus 193 Tôn Dật Tiên', docDate: orderDate, lines, issues, source: 'pdf' };
  }
  function parseSiba(acct, items, fileName) {
    const issues = [];
    const F = re => items.find(x => re.test(x.t));
    const hBc = F(/^mã vạch$/i), hQty = F(/^lượng$/i) || F(/^số lượng$/i);
    if (!hBc || !hQty) return { error: 'Không thấy cột Mã vạch / Số lượng trên PO. File này có đúng mẫu đơn đặt hàng Siba Food không?' };
    const heads = { stt: F(/^stt$/i), sku: F(/^mã sản$/i), bc: hBc, name: F(/^tên sản phẩm$/i), unit: F(/^đvt$/i), desc: F(/^mô tả$/i), qty: hQty, price: F(/^đơn$/i) || F(/^giá$/i), vat: F(/^vat$/i), ck: F(/^ckt$/i), amt: F(/^thành tiền$/i) };
    const tot = F(/^tổng giá trị trước thuế/i);
    const skip = it => /^(stt|mã sản|phẩm|mã vạch|tên sản phẩm|đvt|mô tả|mh|số|lượng|đơn|giá|vat|ckt|m|thành tiền|chưa vat|page \d+ of \d+)$/i.test(it.t);
    const rows = pdfTable(items, { heads, qty: 'qty', bc: 'bc', name: 'name', top: hQty.y1, stop: tot ? tot.y0 - 1 : 0, skip });
    if (!rows.length) return { error: 'Không thấy dòng hàng nào trên PO Siba Food.' };
    const lines = pdfLines(rows, issues);
    const po = ((itemAfter(items, /^số đơn hàng:?$/i) || {}).t || '').replace(/\D/g, '');
    const orderDate = dmy((itemAfter(items, /^ngày đặt hàng:?$/i) || {}).t);
    const giao = dmy((itemAfter(items, /^ngày giao hàng:?$/i) || {}).t);
    const nhan = F(/^nơi nhận hàng:?$/i);
    const addr = nhan ? items.filter(x => x !== nhan && x.x0 >= nhan.x0 - 2 && x.y0 >= nhan.y0 - 2 && x.y0 < nhan.y0 + 20 && !/:$/.test(x.t)).sort((a, b) => a.y0 - b.y0).map(x => x.t).join(' ') : '';
    const m = matchAddr(acct.customers, addr, acct.brandWords);
    if (!m.hit) issues.push({ key: 'customer', level: 'err', msg: `Cửa hàng "${addr || '?'}" chưa có mã KH trong DMKH. Chọn ở ô Mã khách hàng (rồi bấm Ghi nhớ).` });
    if (!po) issues.push({ key: 'po', level: 'err', msg: 'Không đọc được Số đơn hàng. Nhập tay Số đơn.' });
    if (giao) issues.push({ level: 'info', msg: `Phiếu ghi ngày giao ${giao.split('-').reverse().join('/')}; ngày giao trên form = ngày về đơn + ${acct.leadDays} ngày làm việc.` });
    return { fixedPo: true, editable: true, po, customer: m.hit ? m.hit.code : '', store: addr, docDate: orderDate, lines, issues, source: 'pdf' };
  }
  const pdfAcct = (label, fn, extra) => Object.assign({
    label, fixedPo: true, async: true, accept: '.pdf', leadDays: 3, workDays: true, fromNgayVe: true, customers: [],
    async parseFile(f, onProgress, libs) {
      if (!/\.pdf$/i.test(f.name)) return { error: 'Account này nhận file PDF.' };
      try { await libs.pdf(); } catch (e) { return { error: 'Không tải được bộ đọc PDF (cần internet).' }; }
      onProgress && onProgress('Đang đọc PDF…');
      return fn(this, await root.FUJI.pdfItems(f.buf), f.name);
    }
  }, extra);


  /* ================= đợt 3 (13/11 chiều): Anh Việt, Clever, Family Foods Market, Kaiser, Ohmee, Rfood, Top5 Mart, Wefresh ================= */
  // chữ tắt: 1 từ -> chữ đầu + phụ âm (Marina -> MRN) hoặc 4 chữ đầu (Times -> TIME); nhiều từ -> chữ cái đầu (Nguyễn Văn Lộc -> NVL)
  function shortName(name, brandWords, oneWord) {
    const w = key(name).split(' ').filter(x => x && !brandWords.includes(x) && !/^\d+[a-z]?$/.test(x));
    if (!w.length) return '';
    if (w.length === 1) return (oneWord === 'cons' ? (w[0][0] + w[0].slice(1).replace(/[aeiouy]/g, '')).slice(0, 3) : w[0].slice(0, 4)).toUpperCase();
    return w.map(x => x[0]).join('').toUpperCase();
  }
  // file Excel dạng "bảng hàng + cột cửa hàng": tiêu đề [Barcode | Tên | <tên cửa hàng> | <mã KH>], SL nằm dưới cột tên cửa hàng
  function parseStoreSheet(acct, ws, sheetName) {
    const range = XLSXlib.utils.decode_range(ws['!ref'] || 'A1:A1');
    for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 20); r++) {
      let cBc = -1, cCode = -1, code = '';
      for (let c = range.s.c; c <= Math.min(range.e.c, 40); c++) {
        const t = cellText(ws, r, c).trim(), f = fold(t);
        if (cBc < 0 && /^(barcode|ma hh|ma vach|ma hang hoa)$/.test(f)) cBc = c;
        if (cCode < 0 && acct.codeRe.test(t)) { cCode = c; code = t.toUpperCase(); }
      }
      if (cBc < 0 || cCode < 1) continue;
      const cQty = cCode - 1, store = cellText(ws, r, cQty).trim();
      const issues = [], lines = [];
      for (let rr = r + 1; rr <= range.e.r; rr++) {
        const raw = cellText(ws, rr, cBc), q = cellNum(ws, rr, cQty);
        if (!raw.replace(/\D/g, '') || q == null || q === 0) continue;
        lines.push(stdLine(rr + 1, raw, q, cellText(ws, rr, cBc + 1)));
      }
      dupCheck(lines, issues);
      const known = acct.customers.find(c => c.code.toUpperCase() === code);
      if (!known) issues.push({ key: 'customer', level: 'warn', msg: `Mã KH ${code} trên file không có trong DMKH. Kiểm tra lại.` });
      const abbr = acct.abbrOf ? acct.abbrOf(known ? known.name : store, sheetName) : '';
      return { lines, issues, store, code: known ? known.code : code, abbr };
    }
    return null;
  }
  function storeSheetAcct(acct, wb) {
    const out = [], skipped = [];
    for (const sn of wb.SheetNames) {
      const s = parseStoreSheet(acct, wb.Sheets[sn], sn);
      if (!s) { skipped.push(sn); continue; }
      if (!s.lines.length) { skipped.push(sn + ' (không có SL)'); continue; }
      if (acct.prefix && !acct.noAbbr && !s.abbr) s.issues.push({ key: 'po', level: 'err', msg: 'Không tạo được chữ tắt cửa hàng cho số đơn. Nhập tay Số đơn.' });
      out.push({ prefix: acct.prefix, seq: acct.seq, abbr: s.abbr, noAbbr: !!acct.noAbbr, abbrAfterDate: !!acct.abbrAfterDate, customer: s.code, store: s.store, sheet: wb.SheetNames.length > 1 ? sn : undefined, lines: s.lines, issues: s.issues });
    }
    if (!out.length) return { error: `Không thấy bảng hàng có cột cửa hàng + mã KH (${acct.codeHint}). File này có đúng mẫu đơn ${acct.label} không?` };
    if (skipped.length) out[0].issues.push({ level: 'info', msg: `Bỏ qua sheet: ${skipped.join(', ')}.` });
    if (out.length > 1) out.forEach(o => o.issues.push({ level: 'info', msg: `File có ${out.length} sheet đặt hàng → ${out.length} đơn riêng.` }));
    return out.length > 1 ? { multi: out } : out[0];
  }
  function parseAnhViet(acct, wb, fileName) {
    const s = parseSimpleSheet(acct, wb, fileName, { bcHead: /^barcode$/, qtyHead: /^sl$|^so luong/ });
    if (s.error) return s;
    return { prefix: acct.prefix, seq: acct.seq, abbr: '', noAbbr: true, customer: acct.fixedCustomer, store: 'Anh Việt Mart', lines: s.lines, issues: s.issues };
  }

  function parseFamilyFoods(acct, items, fileName) {
    const issues = [];
    const F = re => mergeRuns(items).find(x => re.test(x.t));
    const hBc = F(/^barcode$/i), hQty = F(/^qty$/i);
    if (!hBc || !hQty) return { error: 'Không thấy cột Barcode / Qty trên PO. File này có đúng mẫu Purchase Order của Family Foods Market không?' };
    const heads = { stt: F(/^stt$/i), ref: F(/^reference$/i), bc: hBc, name: F(/^full name$/i), qty: hQty, rec: F(/^received$/i), uom: F(/^uom$/i), price: F(/^price$/i), tax: F(/^taxes$/i), sub: F(/^subtotal$/i) };
    const skip = it => /^(stt|reference|barcode|full name|qty|received|uom|price|taxes|subtotal|internal|unit|page:.*|\d+ \/ \d+)$/i.test(it.t) || /familyfoodsmarket|tax id|^1900888659$/i.test(it.t);
    const tot = F(/^untaxed amount|^total/i);
    const rows = pdfTable(items, { heads, qty: 'qty', bc: 'bc', name: 'name', top: hQty.y1, stop: tot ? tot.y0 - 1 : 0, skip });
    if (!rows.length) return { error: 'Không thấy dòng hàng nào trên PO Family Foods Market.' };
    // SL in dạng "24.000" (= 24); bỏ số trang "1 / 4" lọt vào cột Qty
    for (let i = rows.length - 1; i >= 0; i--) { if (!/^\d{1,3}(,\d{3})*\.\d{3}$/.test(rows[i].anchor.t)) rows.splice(i, 1); else rows[i].anchor = Object.assign({}, rows[i].anchor, { t: String(Number(rows[i].anchor.t.replace(/,/g, ''))) }); }
    if (!rows.length) return { error: 'Không thấy dòng hàng nào trên PO Family Foods Market.' };
    const lines = pdfLines(rows, issues);
    const txt = allText(mergeRuns(items));
    const mPo = txt.match(/Purchase Order\s*#\s*(P\d+)/i) || fileName.match(/(P\d{5,})/);
    const po = mPo ? mPo[1].toUpperCase() : '';
    if (!po) issues.push({ key: 'po', level: 'err', msg: 'Không đọc được số Purchase Order (P…). Nhập tay Số đơn.' });
    const ship = txt.match(/Shipping address:\s*(.{0,240})/i);
    const m = matchAddr(acct.customers, ship ? ship[1] : txt.slice(0, 300), acct.brandWords);
    if (!m.hit) issues.push({ key: 'customer', level: 'err', msg: 'Không dò được cửa hàng theo Shipping address. Chọn ở ô Mã khách hàng (rồi bấm Ghi nhớ).' });
    return { fixedPo: true, editable: true, po, customer: m.hit ? m.hit.code : '', store: m.hit ? m.hit.name : (ship ? ship[1].replace(/\s+/g, ' ').trim().slice(0, 60) : 'Family Foods Market'), lines, issues, source: 'pdf' };
  }
  function parseOhmee(acct, items, fileName) {
    const issues = [];
    const M = mergeRuns(items), F = re => M.find(x => re.test(x.t));
    const hBc = F(/^mã vạch$/i), hQty = hBc && M.find(x => /^số$/i.test(x.t) && Math.abs(x.y0 - hBc.y0) < 8);
    if (!hBc || !hQty) return { error: 'Không thấy cột Mã vạch / Số lượng trên PO. File này có đúng mẫu đơn đặt hàng Ohmee không?' };
    const heads = { stt: F(/^stt$/i), item: F(/^mã hàng$/i), bc: hBc, name: F(/^tên hàng hóa$/i), unit: F(/^đvt$/i), quy: F(/^quy$/i), slt: M.find(x => /^sl$/i.test(x.t) && Math.abs(x.y0 - hBc.y0) < 8), qty: hQty, price: F(/^đơn giá$/i), giathung: F(/^giá thùng$/i), vat: F(/^vat%$/i) };
    const tot = F(/^số mặt hàng/i);
    const rows = pdfTable(items, { heads, qty: 'qty', bc: 'bc', name: 'name', top: hQty.y1 + 6, stop: tot ? tot.y0 - 1 : 0 });
    if (!rows.length) return { error: 'Không thấy dòng hàng nào trên PO Ohmee.' };
    const lines = pdfLines(rows, issues);
    const after = re => (M.find(x => re.test(x.t)) ? itemAfter(M, re) || {} : {}).t || '';
    const po = after(/^số po:?$|^po:$/i).replace(/\s/g, '') || (fileName.match(/PO\d{10,}/i) || [''])[0].toUpperCase();
    if (!po) issues.push({ key: 'po', level: 'err', msg: 'Không đọc được Số PO. Nhập tay Số đơn.' });
    const giao = dmy(after(/^ngày giao hàng:?$/i));
    const kho = after(/^n hàng:$|kho nhận hàng:?$/i) || (F(/\[WH\d+\]/) || {}).t || '';
    const addr = after(/^a chỉ:$|^địa chỉ:$/i);
    let hit = null;
    if (/central warehouse|\bWH001\b/i.test(kho)) hit = acct.customers.find(c => /kho dc/i.test(c.name)) || null;
    if (!hit) { const m = matchAddr(acct.customers, `${kho} ${addr}`, acct.brandWords); hit = m.hit; }
    if (!hit) issues.push({ key: 'customer', level: 'err', msg: `Không dò được cửa hàng từ "${kho}". Chọn ở ô Mã khách hàng (rồi bấm Ghi nhớ).` });
    const sum = lines.reduce((s, l) => s + (l.qty || 0), 0);
    const tq = after(/^tổng số lượng[：:]?$/i).replace(/\D/g, '');
    if (tq) issues.push(+tq === sum ? { level: 'info', msg: `Khớp Tổng số lượng trên PO: ${sum}.` } : { level: 'err', msg: `Tổng SL đọc được (${sum}) khác Tổng số lượng trên PO (${tq}).` });
    if (giao) issues.push({ level: 'info', msg: `Ngày giao lấy theo PO: ${giao.split('-').reverse().join('/')}.` });
    return { fixedPo: true, editable: true, po, customer: hit ? hit.code : '', store: kho || addr, dueFixed: giao, lines, issues, source: 'pdf' };
  }
  // Wefresh: phiếu in font TCVN3 (chữ Việt lỗi) -> chỉ dùng số
  function parseWefresh(acct, items, fileName) {
    const issues = [];
    const M = mergeRuns(items), F = re => M.find(x => re.test(x.t));
    const hBc = F(/^M· v¹ch$|^mã vạch$/i), hQty = F(/^Theo$/i);
    if (!hBc || !hQty) return { error: 'Không thấy cột Mã vạch / SL "Theo đơn" trên phiếu. File này có đúng mẫu phiếu đặt hàng Wefresh không?' };
    const heads = { stt: F(/^STT$/), item: F(/^M· hμng$|^M· hàng$|^mã hàng$/i), bc: hBc, unit: F(/^§VT$|^đvt$/i), qty: hQty, rec: F(/^Thùc$/), price: F(/^§¬n gi¸$/), amt: F(/^Thμnh tiÒn$/) };
    const skip = it => !/^[\d.,]+$/.test(it.t) || /^\d{18}$/.test(it.t);
    const tops = M.filter(x => x === hQty || (x.t === hQty.t)).map(x => x.y1);
    // phiếu nhiều trang: lấy từng trang dưới dòng tiêu đề
    const rows = [];
    const pages = [...new Set(M.map(x => x.page))];
    pages.forEach(pg => {
      const P = M.filter(x => x.page === pg), h = P.find(x => x.t === hQty.t);
      if (!h) return;
      const hd = {}; Object.entries(heads).forEach(([k, it]) => { if (it) hd[k] = P.find(x => x.t === it.t) || it; });
      const end = P.find(x => /^Tæng gi¸ trÞ|^- Hå s¬/.test(x.t));
      rows.push(...pdfTable(P, { heads: hd, qty: 'qty', bc: 'bc', top: h.y1 + 14, stop: end ? end.y0 - 1 : 0, skip }));
    });
    if (!rows.length) return { error: 'Không thấy dòng hàng nào trên phiếu Wefresh.' };
    const lines = pdfLines(rows, issues);
    const po = (M.find(x => /^\d{18}$/.test(x.t)) || {}).t || '';
    if (!po) issues.push({ key: 'po', level: 'err', msg: 'Không đọc được Số đơn (18 số). Nhập tay Số đơn.' });
    const nhan = F(/^N¬i nhËn:|^nơi nhận:/i);
    const addr = nhan ? M.filter(x => Math.abs(x.y0 - nhan.y0) < 3 && x.x0 > nhan.x1 && x.page === nhan.page).map(x => x.t).join(' ') : '';
    let hit = acct.customers.length === 1 ? acct.customers[0] : null;
    if (/174/.test(addr)) hit = acct.customers.find(c => /174/.test(c.name)) || hit;
    if (!hit) issues.push({ key: 'customer', level: 'err', msg: 'Chọn cửa hàng Wefresh ở ô Mã khách hàng (rồi bấm Ghi nhớ).' });
    return { fixedPo: true, editable: true, po, customer: hit ? hit.code : '', store: addr ? 'Wefresh ' + addr.replace(/^\d+\s*/, '') : 'Wefresh', lines, issues, source: 'pdf' };
  }

  /* ---------- Account rules ---------- */
  const ACCOUNTS = {
    KKTC: {
      label: 'Kmart (KKTC)', prefix: 'KKTC', seq: 'always', brand: KKTC_BRAND,
      accept: '.xlsx,.xls,.jpg,.jpeg,.png', async: true,
      leadDays: 4, // ngày giao = ngày về đơn + 4 (10/11 -> 14/11)
      // DMKH: mã KH + Tên nội bộ (lọc từ DMKH.xlsx, 45 cửa hàng KKTC, Active)
      customers: [{"code": "KKTC0064", "name": "K-Market 126 Nguyễn Cao - Bắc Ninh"}, {"code": "KKTC0068", "name": "K-MARKET SMART CITY"}, {"code": "KKTC0041", "name": "Kmart The Emerald"}, {"code": "KKTC0056", "name": "KKTC_ Mỹ Đình Pearl"}, {"code": "KKTC0067", "name": "K-MARKET 117 TRẦN DUY HƯNG"}, {"code": "KKTC0050", "name": "K-Market West Point"}, {"code": "KKTC0072", "name": "K MARTKET LONG BIÊN- Vinhome Symphony"}, {"code": "KKTC0036", "name": "Kmarket Ecopark"}, {"code": "KKTC0053", "name": "KKTC_S3 Shunshine City"}, {"code": "KKTC0043", "name": "K-Market Vinhomes Skylake S3"}, {"code": "KKTC0042", "name": "K-Market Thăng Long Number"}, {"code": "KKTC0052", "name": "KKTC_Dcapitale C6"}, {"code": "KKTC0062", "name": "K-Market Green Bay - Hạ Long - Quảng Ninh"}, {"code": "KKTC0063", "name": "K-Market Chelsea House - Văn Cao - Hải Phòng"}, {"code": "KKTC0057", "name": "KKTC_K-Market Vinhome Bắc Ninh"}, {"code": "KKTC0069", "name": "K-MARKET KTX SAMSUNG"}, {"code": "KKTC0013", "name": "Kmart Trung Hòa"}, {"code": "KKTC0014", "name": "Kmart Calidas"}, {"code": "KKTC0015", "name": "Kmart CipuTra"}, {"code": "KKTC0016", "name": "Kmart Long Biên"}, {"code": "KKTC0017", "name": "Kmart Royal City R1"}, {"code": "KKTC0023", "name": "Kmart Goldmart Ruby"}, {"code": "KKTC0025", "name": "Kmart Quang Minh"}, {"code": "KKTC0028", "name": "Kmart ciputra 2"}, {"code": "KKTC0051", "name": "KKTC_ Phạm văn Đồng"}, {"code": "KKTC0058", "name": "K-market 148 Xuân Diệu"}, {"code": "KKTC0066", "name": "K-MARKET 88 LÊ QUANG ĐẠO - ĐÀ NẴNG"}, {"code": "KKTC0059", "name": "K-MARKET VÕ NGUYÊN GIÁP- ĐÀ NẴNG"}, {"code": "KKTC0071", "name": "K-market CT9 Mỹ Đình"}, {"code": "KKTC0065", "name": "K-Market Dự án The Minato Residence - Hải Phòng"}, {"code": "KKTC0047", "name": "K-Market Việt Long Bắc Ninh"}, {"code": "KKTC0034", "name": "Kmarket D-Capitale"}, {"code": "KKTC0060", "name": "K-Market 342 Ngọc Hân Công Chúa - Bắc Ninh"}, {"code": "KKTC0002", "name": "Kmart Keangnam"}, {"code": "KKTC0003", "name": "Kmart Mỹ Đình"}, {"code": "KKTC0004", "name": "Kmart 17T3"}, {"code": "KKTC0005", "name": "Kmart Golden"}, {"code": "KKTC0007", "name": "Kmart Gardenia"}, {"code": "KKTC0008", "name": "Goldmak saphia"}, {"code": "KKTC0012", "name": "Kmarket Greenbay"}, {"code": "KKTC0039", "name": "K-market MONARCHY ĐÀ NẴNG"}, {"code": "KKTC0031", "name": "Kmart Vinhome Sky Lake"}, {"code": "KKTC0061", "name": "K-Market The Matrix One"}, {"code": "KKTC0049", "name": "K-Market Starlake Tây Hồ"}, {"code": "KKTC0055", "name": "KKTC_ TT4 - Mỹ Đình 1"}],
      // Số đơn: KKTC + 6 ký tự đầu tên cửa hàng (bỏ Kmart, bỏ "The", bỏ dấu) + YYMMDD + -01/-02
      makeAbbr(store) { return { abbr: key(store, KKTC_BRAND).split(' ').filter(w => w && w !== 'the').join('').toUpperCase().slice(0, 6) }; },
      parse(wb) { return parseStoreColumn(this, wb); },
      async parseFile(f, onProgress, libs) {
        if (!/\.(jpe?g|png)$/i.test(f.name)) return parseStoreColumn(this, XLSXlib.read(f.buf, { type: 'array', cellText: true }));
        // Ảnh chụp bảng đơn: không có tên cửa hàng -> lấy theo tên file, vẫn cho chọn tay
        try { await libs.ocr(); } catch (e) { return { error: 'Không tải được bộ nhận dạng chữ (cần internet lần đầu).' }; }
        let res;
        try { res = await root.FUJI.kktcImageParse(new Blob([f.buf]), onProgress, libs.tessOpts); }
        catch (e) { return { error: 'Không đọc được ảnh: ' + (e && e.message || e) }; }
        if (res.error) return { error: res.error };
        const issues = [];
        const fnStore = norm(f.name).replace(/\.[a-z]+$/i, '').replace(/^\s*20\d\d[.\-_]\d{1,2}[.\-_]\d{1,2}\s*-?\s*/, '').replace(/^((anh|ảnh)( dán| dan)?|image|screenshot|img)[\s_-]*\d*$/i, '').trim();
        // tên file hay có chữ thừa: "po kmart MONACHY ĐN" -> "kmart MONACHY Đà Nẵng"
        const clean = fnStore.replace(/^\s*((p\.?o|đơn hàng|don hang|đơn|don|dh|order)\b[\s_.-]*)+/i, '')
          .replace(/(^|[\s_-])(đn|dn)$/i, '$1Đà Nẵng').replace(/(^|[\s_-])(hn)$/i, '$1Hà Nội').replace(/(^|[\s_-])(hcm|sg)$/i, '$1Hồ Chí Minh').replace(/(^|[\s_-])(bn)$/i, '$1Bắc Ninh').trim();
        let m = clean ? matchCustomer(this.customers, clean, this.brand) : { hit: null };
        if (!m.hit && fnStore && clean !== fnStore) m = matchCustomer(this.customers, fnStore, this.brand);
        const customer = m.hit ? m.hit.code : '';
        // viết tắt lấy theo tên cửa hàng trong DMKH (tên file hay viết tắt / sai chính tả)
        const dm = m.hit ? (this.customers.find(c => c.code === m.hit.code && !c.team) || m.hit) : null;
        const abbr = dm ? this.makeAbbr(dm.name).abbr : '';
        if (!m.hit) issues.push({ key: 'customer', level: 'err', msg: fnStore ? `Không dò được cửa hàng "${fnStore}" (tên file) trong DMKH. Chọn cửa hàng ở ô Mã khách hàng.` : 'Ảnh không có tên cửa hàng. Chọn cửa hàng ở ô Mã khách hàng (hoặc đặt tên file ảnh theo tên cửa hàng).' });
        else if (!m.exact) issues.push({ level: 'warn', msg: `Khớp gần đúng theo tên file: "${fnStore}" → ${m.hit.code} · ${m.hit.name}. Kiểm tra lại.` });
        const lines = res.rows.map((x, i) => {
          const l = validateFuji({ row: i + 1, barcode: x.barcode, qty: x.slDv * x.slCai, preview: x.preview, bcState: x.bcSure ? 'sure' : x.bcGuessed ? 'guess' : 'weak' });
          l.flags = l.flags.filter(g => !/đơn giá\/thành tiền/.test(g.msg));
          if (x.slCai !== 1) l.flags.push({ level: 'info', msg: `SL = ${x.slDv} ĐV × ${x.slCai} cái` });
          if (Math.abs(x.giaEa * x.slCai - x.giaDv) > 2) l.flags.push({ level: 'warn', msg: `Giá ĐV (${x.giaDv.toLocaleString('vi-VN')}) ≠ Giá EA × SL cái: có thể đọc lệch SL cái hoặc số lượng. So với ảnh dòng.` });
          return l;
        });
        if (res.skipped) issues.push({ level: 'warn', msg: `Có ${res.skipped} dòng trong ảnh không đọc được đủ cột. So ảnh với bảng bên dưới xem có thiếu dòng không.` });
        issues.push({ level: 'info', msg: `Đọc từ ảnh: ${lines.length} dòng. SL = SL ĐV × SL cái. Ô đỏ là đọc nhầm: sửa trực tiếp rồi bấm Enter.` });
        return { prefix: this.prefix, seq: this.seq, abbr, customer, store: fnStore, lines, issues, editable: true, source: 'image' };
      }
    },

    MOR: {
      label: 'Morgant (MORGT)', prefix: 'MOR', seq: 'dup', brand: /^(morgt|morgant)\b\s*/,
      accept: '.xlsx,.xls',
      leadDays: 4,
      customers: [{ code: 'MORGT001', name: 'MORGT LÁNG HẠ ĐỐNG ĐA' }, { code: 'MORGT002', name: 'MORGT HOÀNG MINH GIÁM TRUNG HOÀ' }],
      // Số đơn: MOR + viết tắt cửa hàng + YYMMDD (không đuôi -01)
      abbr: { MORGT002: 'HMG' }, // giữ HMG như các đơn đã lên hệ thống
      makeAbbr(store, code) {
        if (this.abbr[code]) return { abbr: this.abbr[code] };
        const w = key(store, this.brand).split(' ').filter(Boolean);
        // chữ cái đầu mỗi chữ (đã bỏ MORGT, bỏ dấu): LÁNG HẠ ĐỐNG ĐA -> LHDD
        return { abbr: w.map(x => x[0]).join('').toUpperCase() };
      },
      parse(wb) { return parseStoreColumn(this, wb); }
    }
    ,
    BRG: {
      label: 'BRG / Fujimart (BRGMART)', fixedPo: true, async: true,
      brand: /^(brg\s*mart|brg\s*food|brg|fujimart|fuji\s*mart|fuji|haprofood|hapro\s*food|hapro|cua\s*hang|sieu\s*thi)\b[\s_-]*/,
      accept: '.pdf,.jpg,.jpeg,.png',
      leadDays: 4, // ngày giao = ngày đặt trên phiếu + 4
      customers: [{"code": "BRGMART201", "name": "FUJIMART NGỌC LÂM"}, {"code": "BRGMART199", "name": "Fujimart 328 Minh Khai"}, {"code": "BRGMART196", "name": "BRGMART HÀNG BÀI HÀ NỘI"}, {"code": "BRGMART202", "name": "FUJIMART LINH ĐÀM"}, {"code": "BRGMART097", "name": "FUJIMART_The Light"}, {"code": "BRGMART098", "name": "BRGMART_ UDIC Riverside"}, {"code": "BRGMART099", "name": "FUJIMART Lạc Long Quân"}, {"code": "BRGMART100", "name": "BRGMART C13 Thành Công"}, {"code": "BRGMART101", "name": "BRGMART Hải Dương"}, {"code": "BRGMART102", "name": "BRGMART K3 Việt Hưng"}, {"code": "BRGMART103", "name": "BRGMART 120 Hàng Trống"}, {"code": "BRGMART104", "name": "BRGMART 8 Phạm Ngọc Thạch"}, {"code": "BRGMART105", "name": "FUJIMART 15 -17 Ngọc Khánh"}, {"code": "BRGMART106", "name": "FUJIMART D2 Giảng Võ"}, {"code": "BRGMART107", "name": "BRGMART C12 Thanh Xuân"}, {"code": "BRGMART108", "name": "Fuji Mart Trần Phú"}, {"code": "BRGMART109", "name": "FUJIMART_ Nguyễn Cơ Thạch"}, {"code": "BRGMART110", "name": "BRGMART Moonlight"}, {"code": "BRGMART112", "name": "BRGFOOD_ HaproFood 9-11 Thổ Quan"}, {"code": "BRGMART113", "name": "BRGMART_N16 Sài Đồng"}, {"code": "BRGMART114", "name": "BRGMART_Intracom Reverside"}, {"code": "BRGMART115", "name": "BRGMART_Nguyễn Văn Cừ"}, {"code": "BRGMART116", "name": "BRGMART_Phố Nối - Hưng Yên"}, {"code": "BRGMART117", "name": "BRGMART_Minh Khai - Hải Phòng"}, {"code": "BRGMART118", "name": "BRGMART_63 Hàng Trống"}, {"code": "BRGMART119", "name": "BRGMART_Mạo Khê Quảng Ninh"}, {"code": "BRGMART120", "name": "BRGMART_Lý Nam Đế"}, {"code": "BRGMART121", "name": "BRGMART_Nguyễn Trãi"}, {"code": "BRGMART122", "name": "Fujimart Lê Duẩn"}, {"code": "BRGMART123", "name": "Fujimart Hoàng Cầu"}, {"code": "BRGMART124", "name": "Fuji mart Tây sơn"}, {"code": "BRGMART125", "name": "BRGMART_Giảng Võ"}, {"code": "BRGMART128", "name": "BRGFOOD_Lương Định Của"}, {"code": "BRGMART129", "name": "BRGFOOD_Vĩnh Phúc"}, {"code": "BRGMART134", "name": "BRGFOOD_Hàm Tử Quan"}, {"code": "BRGMART137", "name": "BRGFOOD_198 Lò Đúc"}, {"code": "BRGMART139", "name": "BRGFOOD_Trung Hòa Nhân Chính"}, {"code": "BRGMART140", "name": "BRGFOOD_Chợ bưởi"}, {"code": "BRGMART147", "name": "BRGFOOD_Ngõ Thái Thịnh 1"}, {"code": "BRGMART148", "name": "BRGFOOD_Nguyễn An Ninh"}, {"code": "BRGMART163", "name": "BRGFOOD_96 Tô Ngọc Vân"}, {"code": "BRGMART164", "name": "BRGFOOD_98 Tô Ngọc Vân"}, {"code": "BRGMART167", "name": "BRGFOOD_105 Lê Duẩn"}, {"code": "BRGMART169", "name": "BRGFOOD_Ngọc Lâm"}, {"code": "BRGMART171", "name": "BRGMART KHO DC"}, {"code": "BRGMART177", "name": "FUJIMART 51 Lê Đại Hành"}, {"code": "BRGMART178", "name": "BRGMART_Lộc Ninh"}, {"code": "BRGMART179", "name": "BRGFOOD_Tân Xuân_Ecohome 3"}, {"code": "BRGMART180", "name": "Fuji Mart Lê Văn Lương"}, {"code": "BRGMART181", "name": "Fuji Mart Trung Yên"}, {"code": "BRGMART182", "name": "Fuji Mart Thụy Khuê"}, {"code": "BRGMART183", "name": "BRGFOOD_24 Trần Nhật Duật"}, {"code": "BRGMART184", "name": "BRGFOOD_9 Lê Quý Đôn"}, {"code": "BRGMART187", "name": "FUJIMART 89 Bùi Ngọc Dương"}, {"code": "BRGMART188", "name": "FUJIMART Chính Kinh"}, {"code": "BRGMART189", "name": "Fuji Mart Huỳnh Thúc Kháng"}, {"code": "BRGMART191", "name": "Fuji Mart 89 Lạc Long Quân"}, {"code": "BRGMART197", "name": "Fujimart Ocean Park 1"}, {"code": "BRGMART193", "name": "Fujimart 109 Tân Mai"}, {"code": "BRGMART194", "name": "Fujimart T9 Times city"}, {"code": "BRGMART203", "name": "FUJIMART BẠCH MAI"}, {"code": "BRGMART205", "name": "FUJIMART NGỌC LÂM"}, {"code": "BRGMART206", "name": "FUJIMART OCEAN PARK HD09"}, {"code": "BRGMART207", "name": "FUJIMART SMART CITY 2"}, {"code": "BRGMART200", "name": "FUJIMART SMART CITY"}, {"code": "BRGMART198", "name": "Fujimart DC Complex"}, {"code": "BRGMART195", "name": "BRGMART Đồ Sơn Hải phòng"}, {"code": "BRGMART192", "name": "FUJIMART - 67 Trần Phú"}, {"code": "BRGMART204", "name": "FUJIMART L4 SÀI ĐỒNG"}],
      // Số đơn = "Số Đơn" in trên phiếu (18 số) · Mã KH = tra "Nơi nhận" trong DMKH · SL = cột "Theo đơn"
      async parseFile(f, onProgress, libs) {
        const F = root.FUJI, isPdf = /\.pdf$/i.test(f.name);
        let res;
        if (isPdf) {
          try { await libs.pdf(); } catch (e) { return { error: 'Không tải được bộ đọc PDF (cần internet).' }; }
          res = F.parseWords(await F.pdfWords(f.buf)); res.source = 'pdf';
        } else {
          try { await libs.ocr(); } catch (e) { return { error: 'Không tải được bộ nhận dạng chữ. Đọc ảnh chỉ chạy trong file FFM Chuyen Don.html mở trên máy có internet (link web không hỗ trợ). Nếu có file PDF gốc, dùng PDF.' }; }
          try { res = await F.imageParse(new Blob([f.buf]), onProgress, libs.tessOpts); }
          catch (e) { return { error: 'Không đọc được ảnh: ' + (e && e.message || e) + '. Nếu đang dùng link web, hãy mở file FFM Chuyen Don.html trên máy.' }; }
          res.source = 'image';
        }
        if (res.error) return { error: res.error };
        const issues = [];
        const fnPo = (norm(f.name).match(/\d{18}/) || [])[0] || '';
        const po = res.po || fnPo;
        if (!po) issues.push({ key: 'po', level: 'err', msg: 'Không đọc được Số Đơn (18 số) trên phiếu. Nhập tay.' });
        else if (res.po && fnPo && res.po !== fnPo) issues.push({ level: 'warn', msg: `Số đơn trên phiếu (${res.po}) khác số trong tên file (${fnPo}). Đang dùng số trên phiếu.` });
        const fnStore = norm(f.name).replace(/\.[a-z]+$/i, '').replace(/\d{18}/, '').replace(/^\s*20\d\d[.\-_]\d{1,2}[.\-_]\d{1,2}\s*-?/, '').replace(/[-_\s]+$/, '');
        let m = res.store ? matchCustomer(this.customers, res.store, this.brand) : { hit: null };
        let via = res.store;
        if (!m.hit && fnStore) { m = matchCustomer(this.customers, fnStore, this.brand); via = fnStore; }
        const customer = m.hit ? m.hit.code : '';
        if (!m.hit) issues.push({ key: 'customer', level: 'err', msg: `Không dò được cửa hàng "${res.store || fnStore}" trong DMKH. Chọn cửa hàng ở ô Mã khách hàng.` });
        else if (!m.exact) issues.push({ level: 'warn', msg: `Khớp gần đúng: "${via}" → ${m.hit.code} · ${m.hit.name}. Kiểm tra lại.` });
        if (!res.orderDate) issues.push({ level: 'warn', msg: 'Không đọc được Ngày đặt trên phiếu; ngày giao tính theo ô Ngày về đơn.' });
        const lines = res.rows.map((x, i) => validateFuji({ row: i + 1, barcode: x.barcode || '', qty: x.qty, price: x.price, amount: x.amount, preview: x.preview || '', bcState: res.source === 'pdf' || x.bcSure ? 'sure' : x.bcSnapped ? 'snap' : x.bcGuessed ? 'guess' : 'weak' }));
        const sumAmt = res.rows.reduce((a, x) => a + (x.amount || 0), 0);
        const totOk = res.total != null && Math.abs(sumAmt - res.total) <= 1;
        if (res.total != null && !totOk) issues.push({ level: 'warn', msg: `Tổng thành tiền các dòng (${sumAmt.toLocaleString('vi-VN')}) khác "Tổng giá trị đơn hàng" trên phiếu (${res.total.toLocaleString('vi-VN')}). Có thể thiếu dòng hoặc đọc sai số.` });
        issues.push({ level: 'info', msg: `Đọc từ ${res.source === 'pdf' ? 'PDF' : 'ảnh'}: ${lines.length} dòng${totOk ? ', tổng tiền khớp phiếu (' + res.total.toLocaleString('vi-VN') + ')' : ''}.${res.source === 'image' ? ' Ô đỏ là đọc nhầm: sửa trực tiếp trong bảng theo phiếu.' : ''}` });
        return { fixedPo: true, po, customer, store: res.store || fnStore, orderDate: res.orderDate || '', lines, issues, editable: true, source: res.source };
      }
    }
    ,
    AEON: {
      label: 'AEON (CSV hoặc PDF)', fixedPo: true, async: true, aeon: true,
      accept: '.csv,.pdf',
      leadDays: 4, // ngày giao = ngày đặt hàng trên CSV + 4
      customers: [{"code": "AEON0063", "name": "4003 - AEON MIDORI PARK"}, {"code": "AEON0048", "name": "1010 - AEON NGUYỄN VĂN LINH"}, {"code": "AEON0018", "name": "Aeon HCMC RDC"}, {"code": "AEON0001", "name": "1001 - Aeon - CELADON"}, {"code": "AEON0002", "name": "1002 - Aeon - CANARY"}, {"code": "AEON0004", "name": "1004 - Aeon - Bình Tân"}, {"code": "AEON0059", "name": "5488 - Hamper Center - Miền Nam"}, {"code": "AEON0105", "name": "4008 - AEON THẢO ĐIỀN"}, {"code": "AEONTONG", "name": "Công Ty TNHH Aeon Việt Nam"}, {"code": "AEON0039", "name": "1009 - AEON BINH DUONG NEW CITY"}, {"code": "AEONHP0008", "name": "1011 - Aeon Tạ Quang Bửu Hamper DC"}, {"code": "AEON0014", "name": "5402 - AEON CITIMART Miền Nam"}, {"code": "AEON0049", "name": "1011 - Aeon - Tạ Quang Bửu"}, {"code": "AEON0089", "name": "MVL DREAM HOME"}, {"code": "AEON0081", "name": "MAXVALU MIỀN NAM"}, {"code": "AEON0062", "name": "1013 - AEON TÂN AN"}, {"code": "AEONHP0010", "name": "1010 - Hamper_Aeon Nguyễn Văn Linh"}, {"code": "AEON0056", "name": "1010 - Aeon_Nguyễn Văn Linh Giao Cửa Hàng"}, {"code": "AEON0057", "name": "1011 - Aeon_Tạ Quang Bửu Giao Cửa Hàng"}, {"code": "AEON0053", "name": "1001 - Aeon_CELADON  Giao Cửa Hàng"}, {"code": "AEONHP0004", "name": "1004 - Hamper_Aeon - Bình Tân"}, {"code": "AEONHP0039", "name": "1009 - Aeon BINH DUONG NEW CITY Hamper DC"}, {"code": "AEONHP0001", "name": "1001 - Hamper_Aeon CELADON"}, {"code": "AEONHP0002", "name": "1002 - Hamper_Aeon CANARY"}, {"code": "AEON0078", "name": "MAXVALU - HAMPER DC"}, {"code": "AEONHP0013", "name": "1013 - Aeon TÂN AN-HAMPER DC"}, {"code": "AEONHP0014", "name": "4003 - Aeon MIDORI PARK - HAMPER DC"}, {"code": "AEONHP0015", "name": "1010 - Aeon Nguyễn Văn Linh - HAMPER DC"}, {"code": "AEONHP0016", "name": "1004 - Aeon - Bình Tân-Hamper DC"}, {"code": "AEONHP0017", "name": "1001 - Aeon CELADON -Hamper DC"}, {"code": "AEONHP0018", "name": "1002 - Aeon CANARY - Hamper DC"}, {"code": "AEON0054", "name": "1002 - Aeon_CANARY Giao Cửa Hàng"}, {"code": "AEON0032", "name": "AEON_MaxValu Westbay"}, {"code": "AEON0007", "name": "MAXVALU - ECOPARK"}, {"code": "AEON0079", "name": "1013 - AEON TÂN AN - Giao cửa hàng"}, {"code": "AEON0064", "name": "MAXVALU ECOPARK 3"}, {"code": "AEON0065", "name": "MAXVALU OASIS"}, {"code": "AEON0066", "name": "MAXVALU SAPPHIRE"}, {"code": "AEON0067", "name": "MAXVALU LANDMARK"}, {"code": "AEON0068", "name": "MAXVALU ROYAL"}, {"code": "AEON0069", "name": "MAXVALU MASTERI"}, {"code": "AEON0070", "name": "MAXVALU HAVENPARK"}, {"code": "AEON0071", "name": "MAXVALU HAWAII"}, {"code": "AEON0027", "name": "1006 - AEON Hải Phòng"}, {"code": "AEON0034", "name": "AEON_THE NINE"}, {"code": "AEON0041", "name": "MAXVALU_Horizon"}, {"code": "AEON0043", "name": "MAXVALU_Kosmo"}, {"code": "AEON0044", "name": "MAXVALU_The Five"}, {"code": "AEON0046", "name": "MAXVALU_Nam Trung Yên"}, {"code": "AEON0047", "name": "MAXVALU_Zenpark"}, {"code": "AEON0028", "name": "AEON_ Aeon Maxvalu Hà Trì"}, {"code": "AEON0051", "name": "Maxvalu Symphony"}, {"code": "AEON0108", "name": "AEON HN RDC HAMPER"}, {"code": "AEON0080", "name": "4005 - AEON Đà Nẵng Thanh Khê"}, {"code": "AEON0082", "name": "MAXVALU SMART CITY"}, {"code": "AEON0083", "name": "MAXALU CIPUTRA"}, {"code": "AEON0084", "name": "MAXVALU WESTPOINT"}, {"code": "AEON0085", "name": "MAXVALU SUNSHINE CITY"}, {"code": "AEON0033", "name": "8005 - AEON_MAXVALU"}, {"code": "AEON0103", "name": "1017 - AEON HẠ LONG"}, {"code": "AEON0010", "name": "MAXVALU RIVERSIDE"}, {"code": "AEONHP0011", "name": "1012 - Hamper_AEON Xuân Thủy"}, {"code": "AEON0061", "name": "1015 - AEON VĂN GIANG"}, {"code": "AEON0038", "name": "Aeon Maxvalu La Casta"}, {"code": "AEON0077", "name": "1015 - AEON VĂN GIANG Giao Cửa Hàng"}, {"code": "AEONHP0009", "name": "8005 - AEON_MAXVALU - HAMPER DC"}, {"code": "AEON0076", "name": "1006 - AEON Hải Phòng Giao Cửa Hàng"}, {"code": "AEONHP0019", "name": "1012 - AEON Xuân Thủy-Hamper DC"}, {"code": "AEONHP0020", "name": "1006 - AEON Hải Phòng-Hamper DC"}, {"code": "AEONHP0021", "name": "1015 - Aeon - Văn Giang-Hamper DC"}, {"code": "AEONHP0022", "name": "1003 - Aeon - Long Biên-Hamper DC"}, {"code": "AEONHP0023", "name": "1005 - Aeon - Hà Đông-Hamper DC"}, {"code": "AEONHP0024", "name": "1014 - Aeon Thừa Thiên Huế-HAMPER DC"}, {"code": "AEONHP0027", "name": "1006 - Hamper_AEON Hải Phòng"}, {"code": "AEON0058", "name": "MaxValu Five Star Garden"}, {"code": "AEON0088", "name": "AEON MAXVALU SOL FOREST"}, {"code": "AEON0035", "name": "AEON_Maxvalu Ocean Park"}, {"code": "AEON0036", "name": "AEON_Maxvalu Lĩnh Nam"}, {"code": "AEON0091", "name": "MAXVALU ZURICH"}, {"code": "AEON0095", "name": "MAXVALU PARK HILL"}, {"code": "AEON0097", "name": "MAXVALU SUNSHINE RIVER"}, {"code": "AEON0098", "name": "MAXVALU MIAMI"}, {"code": "AEON0099", "name": "MAXVALU SAKURA"}, {"code": "AEON0055", "name": "1004 - Aeon_Bình Tân Giao Cửa Hàng"}, {"code": "AEON_HN_0029", "name": "AEON_ Maxvalu Lotus"}, {"code": "AEONHP0012", "name": "1015 - Hamper_Aeon - Văn Giang"}, {"code": "AEON0072", "name": "1014 - Aeon - Thừa Thiên Huế Giao Cửa Hàng"}, {"code": "AEON0073", "name": "1003 - Aeon - Long Biên Giao Cửa Hàng"}, {"code": "AEON0074", "name": "1005 - Aeon - Hà Đông Giao Cửa Hàng"}, {"code": "AEON0075", "name": "1012 - Aeon - Xuân Thủy Giao Cửa Hàng"}, {"code": "AEON0106", "name": "MAXVALU TRẦN THỦ ĐỘ"}, {"code": "AEON0107", "name": "MAXVALU LAKEVIEW"}, {"code": "AEONHP0003", "name": "1003 - Hamper_Aeon - Long Biên"}, {"code": "AEONHP0006", "name": "1005 - Hamper_Aeon - Hà Đông"}, {"code": "AEON0050", "name": "1014 - Aeon - Thừa Thiên Huế"}, {"code": "AEON0037", "name": "AEON_DC Hà Nội"}, {"code": "AEON0029", "name": "AEON_ Maxvalu Lotus_Giao trực tiếp cửa hàng"}, {"code": "AEON0006", "name": "1005 - Aeon - Hà Đông"}, {"code": "AEON0086", "name": "1016 - AEON THANH HÓA"}, {"code": "AEON0087", "name": "1019 - AEON HẢI DƯƠNG"}, {"code": "AEONHP0007", "name": "8005 - Hamper Aeon Miền Bắc"}, {"code": "AEONHP0005", "name": "1014 - HAMPER _ Aeon Thừa Thiên Huế"}, {"code": "AEON0030", "name": "AEON_MaxValu Khuất Duy Tiến"}, {"code": "AEON0031", "name": "AEON_MaxValu Văn Tân"}, {"code": "AEON0003", "name": "1003 - Aeon - Long Biên"}, {"code": "AEON0060", "name": "1012 - Aeon - Xuân Thủy"}, {"code": "AEON0104", "name": "MAXVALU TROPIC"}, {"code": "AEON0100", "name": "MAXVALU HIM LAM 6"}, {"code": "AEON0101", "name": "MAXVALU CAO THẰNG"}, {"code": "AEON0052", "name": "1009 - Aeon_ BINH DUONG NEW CITY Giao Cửa Hàng"}, {"code": "AEON0102", "name": "AEON Đà Nẵng Thanh Khê - GIAO CỬA HÀNG"}],
      async parseFile(f, onProgress, libs) {
        let d;
        if (/\.pdf$/i.test(f.name)) {
          try { await libs.pdf(); await libs.ocr(); } catch (e) { return { error: 'Không tải được bộ đọc PDF / nhận dạng chữ. File PDF AEON chỉ đọc được trong file FFM Chuyen Don.html mở trên máy có internet. Nếu có file CSV của đơn, dùng CSV.' }; }
          let r;
          try { r = await root.FUJI.aeonPdfParse(f.buf, onProgress, libs.tessOpts); }
          catch (e) { return { error: 'Không đọc được PDF: ' + (e && e.message || e) + '. Nếu đang dùng link web, hãy mở file FFM Chuyen Don.html trên máy.' }; }
          d = { source: 'pdf', po: r.po, storeCode: r.storeCode || '', storeName: r.storeName || '', loc: r.loc || '', orderDate: r.orderDate || '', note: r.note || '', total: r.total,
            items: r.rows.map(x => ({ barcode: x.barcode, aeonName: x.aeonName, item: x.item, pack: x.pack, cartons: x.cartons, unit: x.unit, price: x.price, amount: x.amount, verified: x.verified })) };
          if (!d.items.length) return { error: 'Không tìm thấy dòng hàng nào trong PDF. File này có đúng là "STORE ORDER" của AEON không?' };
        } else {
          d = parseAeonCsv(f.buf);
          if (d.error) return d;
        }
        return aeonBuild(this, d);
      }
    }
    ,
    SOIBIEN: {
      label: 'Sói Biển (SOIBIEN)', prefix: 'SOIBIEN', seq: 'dup', noAbbr: true, pcMode: true,
      accept: '.xlsx,.xls',
      leadDays: 3, // Sói Biển: ngày giao = ngày đặt trên đơn + 3 (anh Tài chốt 12/11)
      // DMKH: mã KH bắt đầu SOIBIEN (Tên nội bộ vd "Sói Biển 074_DC_KHO 2 - Lô S3-8")
      customers: [{"code": "SOIBIEN062", "name": "Sói Biển 62_Số 8 Vũ Thạnh"}, {"code": "SOIBIEN069", "name": "Sói Biển 69_54 Giang Văn Minh"}, {"code": "SOIBIEN076", "name": "Sói Biển 076_16 Xuân La"}, {"code": "SOIBIEN075", "name": "Sói Biển 075_33 Trần Nguyên Đán"}, {"code": "SOIBIEN060", "name": "Sói Biển 60_Ngõ 34 Hoàng Cầu"}, {"code": "SOIBIEN071", "name": "Sói Biển 71 - Số 41 đường Vạn Phúc"}, {"code": "SOIBIEN081", "name": "Sói Biển 080_TT6-01, KĐT mới Văn Phú"}, {"code": "SOIBIEN064", "name": "Sói biển 64_Số 178 Lò Đúc"}, {"code": "SOIBIEN061", "name": "Sói Biển 61_Số 44 Nguyễn Tuân"}, {"code": "SOIBIEN070", "name": "Sói Biển 70 _ Khu đoàn ngoại giao"}, {"code": "SOIBIEN077", "name": "Sói Biển 077_KHO 1 - Số 524 Minh Khai"}, {"code": "SOIBIEN078", "name": "Sói Biển 078_ số 162 Nguyễn Đồng Chi"}, {"code": "SOIBIEN001", "name": "Sói Biển 1_182 Hoàng Văn Thái"}, {"code": "SOIBIEN002", "name": "Sói Biển 2_Số 13B Phan Huy Chú"}, {"code": "SOIBIEN004", "name": "Sói Biển 4_Số 65 Trần Nhân Tông"}, {"code": "SOIBIEN005", "name": "Sói Biển 5_Shophouse 9"}, {"code": "SOIBIEN008", "name": "Sói Biển 8_Số 20 Núi Trúc"}, {"code": "SOIBIEN009", "name": "Sói Biển 9_Số 48 Trần Hưng Đạo"}, {"code": "SOIBIEN010", "name": "Sói Biển 10_Số 16 N7A Nguyễn Thị Thập"}, {"code": "SOIBIEN012", "name": "Sói Biển 12_Shop 1 - Park 11"}, {"code": "SOIBIEN013", "name": "Sói Biển 13_Số 145 Nguyễn Tuân"}, {"code": "SOIBIEN015", "name": "Sói Biển 15_Shophouse 3"}, {"code": "SOIBIEN016", "name": "Sói Biển 16_Số 65 Nguyên Hồng"}, {"code": "SOIBIEN017", "name": "Sói Biển 17_Số 48 Nguyễn Khuyến"}, {"code": "SOIBIEN018", "name": "Sói Biển 18_Số 78 Láng Hạ"}, {"code": "SOIBIEN020", "name": "Sói Biển 20_Số 203 Trung Kính"}, {"code": "SOIBIEN021", "name": "Sói Biển 21_Số 25 Xuân La"}, {"code": "SOIBIEN035", "name": "Sói Biển 35_Số 86D Khu Nhà Vườn"}, {"code": "SOIBIEN040", "name": "Sói Biển 40_Số 50 Lò Đúc"}, {"code": "SOIBIEN041", "name": "Sói Biển 41_116E3 Thái Thịnh"}, {"code": "SOIBIEN042", "name": "Sói Biển 42_52 Nguyễn Huy Tưởng"}, {"code": "SOIBIEN043", "name": "Sói Biển 43_163 Phố Đặng Tiến Đông"}, {"code": "SOIBIEN044", "name": "Sói Biển 44_12C Láng Hạ"}, {"code": "SOIBIEN045", "name": "Sói Biển 45_117 Phan Văn Trường"}, {"code": "SOIBIEN046", "name": "Sói Biển 46_43 Đỗ Quang"}, {"code": "SOIBIEN047", "name": "Sói Biển 47_16 Vũ Trọng Phụng"}, {"code": "SOIBIEN049", "name": "Sói Biển 49_52 Chiến Thắng"}, {"code": "SOIBIEN050", "name": "Sói Biển 50_150 Lê Trọng Tấn"}, {"code": "SOIBIEN048", "name": "Sói Biển 48_423 Minh Khai"}, {"code": "SOIBIEN052", "name": "Sói Biển 52_55 Ngụy Như Kon Tum"}, {"code": "SOIBIEN051", "name": "Sói Biển 51_142 Trần Bình"}, {"code": "SOIBIEN053", "name": "Sói Biển 53_44 Nguyễn Sơn"}, {"code": "SOIBIEN054", "name": "Sói Biển 54_Gh 01S11 Tòa S2.07"}, {"code": "SOIBIEN055", "name": "Sói Biển 55_Gh 01S10 Tòa S2.16"}, {"code": "SOIBIEN056", "name": "Sói Biển 56_Times City"}, {"code": "SOIBIEN057", "name": "Sói Biển 57_106 Hoàng Quốc Việt"}, {"code": "SOIBIEN058", "name": "Sói Biển 58_16 Đoàn Thị Điểm"}, {"code": "SOIBIEN059", "name": "Sói Biển 59_Số 9, B7 Lưu Hữu Phước"}, {"code": "SOIBIEN065", "name": "Sói Biển 65_36 Nguyễn Hoàng Tôn"}, {"code": "SOIBIEN072", "name": "Sói Biển 072_179 Trần Đăng Ninh"}, {"code": "SOIBIEN073", "name": "Sói Biển 073_Vinhomes Smart City"}, {"code": "SOIBIEN080", "name": "Sói Biển 080_VT3, BT7 Xa La"}, {"code": "SOIBIEN063", "name": "Sói biển 63_Số 189 Hoàng Hoa Thám"}, {"code": "SOIBIEN067", "name": "Sói Biển 67_ Mỗ Lao"}, {"code": "SOIBIEN066", "name": "Sói Biển 66_ Số 56 Nguyễn Khuyến"}, {"code": "SOIBIEN068", "name": "Sói Biển 68_Linh Lang"}, {"code": "SOIBIEN074", "name": "Sói Biển 074_DC_KHO 2 - Lô S3-8"}, {"code": "SOIBIEN079", "name": "Sói Biển 079_158 Phố Lạc Trung"}],
      // Số đơn: SOIBIEN + YYMMDD ngày đặt (in trên đơn) · đơn thứ 2 cùng ngày thêm -02
      parse(wb, fileName) { return parseSoiBien(this, wb, fileName || ''); }
    }
    ,
    ANNHIEN: {
      label: 'An Nhiên (ANNHIEN)', prefix: 'ANNHIEN', seq: 'always', noAbbr: true,
      accept: '.xlsx,.xls',
      leadDays: 4,
      fixedCustomer: 'ANHNHCM0001', // anh Tài chốt: An Nhiên cố định mã KH này
      customers: [{ code: 'ANHNHCM0001', name: 'ANNHIEN_237-239-241 Nguyễn Công Trứ' }],
      // Số đơn: ANNHIEN + YYMMDD (ngày trong tên sheet) + -01, -02… (mỗi sheet 1 đơn)
      parse(wb, fileName) { return parseAnNhien(this, wb, fileName || ''); }
    }
    ,
    MIOMIO: {
      label: 'Miomio (PDF)', fixedPo: true, async: true, // anh Tài chốt 13/11: form FORM_2_HP, barcode điền cột C
      accept: '.pdf',
      leadDays: 4,
      fixedCustomer: 'MIOMIOW0004', // anh Tài chốt: Miomio cố định mã KH này
      customers: [{ code: 'MIOMIOW0004', name: 'Miomio' }],
      async parseFile(f, onProgress, libs) {
        try { await libs.pdf(); } catch (e) { return { error: 'Không tải được bộ đọc PDF (cần internet).' }; }
        onProgress && onProgress('Đang đọc PO Miomio…');
        const items = await root.FUJI.pdfItems(f.buf);
        return parseMiomio(this, items, f.name);
      }
    }
    ,
    FANTASY: pdfAcct('Aeon Fantasy (PDF)', parseAeonFantasy, {
      fixedPo: false, prefix: 'AEONFTS', seq: 'dup',
      // Số đơn: AEONFTS + chữ tắt Store ID (AEON HUE-PRIZE -> HUE) + YYMMDD ngày về đơn
      // các acc mới (mẫu anh Tài 13/11): ngày giao = ngày về đơn + 3 ngày làm việc (không tính Chủ nhật)
      brandWords: ['aeon', 'fantasy', 'fantasyaeon', 'fanasy', 'mall', 'dynalecx', 'kidzooona', 'prize', 'trung', 'tam', 'thuong', 'mai', 'tttm', 'center'],
      customers: [{"code": "FANTASY021", "name": "Aeon Fantasy_HA NOI CENTER"}, {"code": "FANTASY001", "name": "Aeon Fantasy Bình Dương Canary"}, {"code": "FANTASY014", "name": "FANTASYAEON Sense City Can Tho & Dynalecx Cần Thơ"}, {"code": "FANTASY006", "name": "Aeon Fantasy Mall Hà Đông"}, {"code": "FANTASY010", "name": "FANTASY Vincom Mega Mall Royal City"}, {"code": "FANTASY011", "name": "FANTASY Aeon Mall Long Biên"}, {"code": "FANTASY005", "name": "Aeon Fanasy ở Hải Phòng"}, {"code": "FANTASY009", "name": "Fantasy Go Nha Trang"}, {"code": "FANTASY018", "name": "Aeon Fantasy Aeon Huế"}, {"code": "FANTASY016", "name": "FANTASYAEON_ Đại lộ V.I.Lenin"}, {"code": "FANTASY008", "name": "Aeon Fantasy Buôn Ma Thuột"}, {"code": "FANTASY007", "name": "Aeon Fantasy Sora Garden Bình Dương"}, {"code": "FANTASY004", "name": "Khu vui chơi Molly Fantasy Tầng 2, Aeon Mall Long Biên"}, {"code": "FANTASY012", "name": "FANTASYAEON Fanpekka"}, {"code": "FANTASY013", "name": "FANTASYAEON Gigamall"}, {"code": "FANTASY022", "name": "FANTASY Vincom Mega Mall Royal Island"}, {"code": "FANTASY002", "name": "Aeon Fantasy Mall Bình Tân"}, {"code": "FANTASY003", "name": "Aeon Fantasy Mall Tân Phú Celadon"}, {"code": "FANTASY019", "name": "FANTASY AEON CITIMART PHƯỚC KIẾN"}, {"code": "FANTASY017", "name": "Aeon Fantasy Parc Mall"}, {"code": "FANTASY020", "name": "FANTASY AEON CITIMART OCEAN CITY"}, {"code": "FANTASY015", "name": "FANTASYAEON 469 Nguyễn Hữu Thọ"}, {"code": "FANTASY023", "name": "FANTASY AEON TÂN AN"}]
    }),
    FGARDEN: {
      label: 'Fresh Garden (FGARDEN)', prefix: 'FRESHGARDEN', seq: 'dup', noAbbr: true,
      accept: '.xlsx,.xls', leadDays: 3, workDays: true, fromNgayVe: true,
      // Số đơn: FRESHGARDEN + YYMMDD ngày về đơn · mã KH dò theo địa chỉ nhận hàng (Lai Xá -> FGARDEN0001)
      brandWords: ['freshgarden', 'fresh', 'garden'],
      customers: [{"code": "FGARDEN0002", "name": "FRESHGARDEN 179 Trâu Quỳ"}, {"code": "FGARDEN0003", "name": "FRESHGARDEN 113 Tô Hiệu, Hà Đông"}, {"code": "FGARDEN0004", "name": "FRESHGARDEN La Casta Văn Phú"}, {"code": "FGARDEN0005", "name": "FRESHGARDEN 93 Xuân La"}, {"code": "FGARDEN0006", "name": "FRESHGARDEN One Mount Times City"}, {"code": "FGARDEN0007", "name": "FRESHGARDEN 130 Lạc Trung"}, {"code": "FGARDEN0008", "name": "FRESHGARDEN 170 Ngô Xuân Quảng"}, {"code": "FGARDEN0009", "name": "FRESHGARDEN 147 Khương Trung"}, {"code": "FGARDEN0010", "name": "FRESHGARDEN trường ĐH Sư Phạm Hà Nội"}, {"code": "FGARDEN0011", "name": "FRESHGARDEN cánh Tây T2 Nội Bài"}, {"code": "FGARDEN0012", "name": "FRESHGARDEN Cánh đông T2 Nội bài"}, {"code": "FGARDEN0013", "name": "FRESHGARDEN Cánh B T1 Nội Bài"}, {"code": "FGARDEN0014", "name": "FRESHGARDEN VinsFRESHGARDENool Tiểu học Sài Đồng"}, {"code": "FGARDEN0015", "name": "FRESHGARDEN trường ĐH Vinuni Oceanpark"}, {"code": "FGARDEN0016", "name": "FRESHGARDEN VinsFRESHGARDENool Metropolis Liễu Giai"}, {"code": "FGARDEN0017", "name": "FRESHGARDEN VinsFRESHGARDENool T36 Time City"}, {"code": "FGARDEN0018", "name": "FRESHGARDEN VinsFRESHGARDENool T37 Times City"}, {"code": "FGARDEN0019", "name": "FRESHGARDEN 14 Ngô Thì Nhậm- Hà Đông"}, {"code": "FGARDEN0020", "name": "FRESHGARDEN 174 Hoàng Văn Thái"}, {"code": "FGARDEN0021", "name": "FRESHGARDEN Bệnh viện E bánh"}, {"code": "FGARDEN0022", "name": "FRESHGARDEN 135 Cổ Nhuế"}, {"code": "FGARDEN0023", "name": "FRESHGARDEN 300B Cầu Giấy"}, {"code": "FGARDEN0024", "name": "FRESHGARDEN 76 Ngụy Như Kom Tum"}, {"code": "FGARDEN0025", "name": "FRESHGARDEN 183 Tô Hiệu"}, {"code": "FGARDEN0026", "name": "FRESHGARDEN 28 Lê Văn Lương"}, {"code": "FGARDEN0027", "name": "FRESHGARDEN S2-08 Vinhomes Ocean Park"}, {"code": "FGARDEN0028", "name": "FRESHGARDEN 20 Láng Hạ"}, {"code": "FGARDEN0029", "name": "FRESHGARDEN 112 C2 Phạm Ngọc ThạFRESHGARDEN"}, {"code": "FGARDEN0030", "name": "FRESHGARDEN 18- BT11 Xa La"}, {"code": "FGARDEN0031", "name": "FRESHGARDEN 430 Vạn Phúc"}, {"code": "FGARDEN0032", "name": "FRESHGARDEN 16 Hàng Cót"}, {"code": "FGARDEN0033", "name": "FRESHGARDEN 42 Nguyễn Sơn"}, {"code": "FGARDEN0034", "name": "FRESHGARDEN 32 Trần Thái Tông"}, {"code": "FGARDEN0035", "name": "FRESHGARDEN 78 Kim Ngưu"}, {"code": "FGARDEN0036", "name": "FRESHGARDEN 65 Lạc Long Quân"}, {"code": "FGARDEN0037", "name": "FRESHGARDEN 36 Nguyễn Hoàng"}, {"code": "FGARDEN0038", "name": "FRESHGARDEN 440 BạFRESHGARDEN Mai"}, {"code": "FGARDEN0039", "name": "FRESHGARDEN CT12A Kim Văn Kim Lũ"}, {"code": "FGARDEN0040", "name": "FRESHGARDEN 41 Âu Cơ"}, {"code": "FGARDEN0041", "name": "FRESHGARDEN HH1B Linh Đàm"}, {"code": "FGARDEN0042", "name": "FRESHGARDEN 272 Thượng Đình"}, {"code": "FGARDEN0043", "name": "FRESHGARDEN 47 Lò Đúc"}, {"code": "FGARDEN0044", "name": "FRESHGARDEN 140 Lê Trọng Tấn"}, {"code": "FGARDEN0045", "name": "FRESHGARDEN 110 Thái Thịnh"}, {"code": "FGARDEN0046", "name": "FRESHGARDEN Vinmec Times City"}, {"code": "FGARDEN0047", "name": "FRESHGARDEN Green Bay"}, {"code": "FGARDEN0048", "name": "FRESHGARDEN 112 B4 Tôn Thất Tùng"}, {"code": "FGARDEN0049", "name": "FRESHGARDEN 19 dãy 16B4 Nguyễn Văn Lộc"}, {"code": "FGARDEN0050", "name": "FRESHGARDEN 46 An Dương"}, {"code": "FGARDEN0051", "name": "FRESHGARDEN T6 Times City"}, {"code": "FGARDEN0052", "name": "FRESHGARDEN 6 Tông Đản"}, {"code": "FGARDEN0053", "name": "FRESHGARDEN 102 FRESHGARDENiến Thắng"}, {"code": "FGARDEN0054", "name": "FRESHGARDEN 324 Xã Đàn"}, {"code": "FGARDEN0055", "name": "FRESHGARDEN 34 Hoàng Cầu"}, {"code": "FGARDEN0056", "name": "FRESHGARDEN Ngã tư Trôi"}, {"code": "FGARDEN0057", "name": "FRESHGARDEN 53 Kim Đồng"}, {"code": "FGARDEN0058", "name": "FRESHGARDEN 58 Nguyễn Khuyến"}, {"code": "FGARDEN0059", "name": "FRESHGARDEN S2 Vinhomes Symphony"}, {"code": "FGARDEN0060", "name": "FRESHGARDEN 287 Trần Đăng Ninh"}, {"code": "FGARDEN0061", "name": "FRESHGARDEN 71 Nguyễn An Ninh"}, {"code": "FGARDEN0062", "name": "FRESHGARDEN 65 Nguyễn FRESHGARDENí Thanh"}, {"code": "FGARDEN0063", "name": "FRESHGARDEN 44 Nguyễn Khánh Toàn"}, {"code": "FGARDEN0064", "name": "FRESHGARDEN 111 Núi Trúc"}, {"code": "FGARDEN0065", "name": "FRESHGARDEN 20C Trần Duy Hưng"}, {"code": "FGARDEN0066", "name": "FRESHGARDEN 41 Doãn Kế Thiện"}, {"code": "FGARDEN0067", "name": "FRESHGARDEN Bệnh viện E cafe"}, {"code": "FGARDEN0068", "name": "FRESHGARDEN TRÀNG THI"}, {"code": "FGARDEN0095", "name": "FRESHGARDEN 108 HỒ TÙNG MẬU"}, {"code": "FGARDEN0069", "name": "FRESHGARDEN 21B5 GREEN STAR"}, {"code": "FGARDEN0070", "name": "FRESHGARDEN 106 THÀNH CÔNG"}, {"code": "FGARDEN0071", "name": "FRESHGARDEN 72 NGUYỄN TUÂN"}, {"code": "FGARDEN0072", "name": "FRESHGARDEN S201 SMART CITY"}, {"code": "FGARDEN0073", "name": "FRESHGARDEN VINMEC SMART CITY"}, {"code": "FGARDEN0074", "name": "FRESHGARDEN TRUNG HỌC VINS SMART CITY"}, {"code": "FGARDEN0075", "name": "FRESHGARDEN TIỂU HỌC VINS SMART CITY"}, {"code": "FGARDEN0076", "name": "FRESHGARDEN 106 CHÙA LÁNG"}, {"code": "FGARDEN0077", "name": "FRESHGARDEN MASTERI TÂY MỖ"}, {"code": "FGARDEN0078", "name": "FRESHGARDEN G6 SMART CITY"}, {"code": "FGARDEN0079", "name": "FRESHGARDEN 66 TRẦN CUNG"}, {"code": "FGARDEN0080", "name": "FRESHGARDEN ADELAIDE"}, {"code": "FGARDEN0081", "name": "FRESHGARDEN LACASA VĂN KHÊ"}, {"code": "FGARDEN0082", "name": "FRESHGARDEN CT8 ĐẠI THANH"}, {"code": "FGARDEN0083", "name": "FRESHGARDEN 360 BẠCH MAI"}, {"code": "FGARDEN0084", "name": "FRESHGARDEN 11 TRẦN NGUYÊN ĐÁN"}, {"code": "FGARDEN0085", "name": "FRESHGARDEN VINSCHOOL OCEAN PART GIA LÂM"}, {"code": "FGARDEN0086", "name": "FRESHGARDEN TRUNG HỌC VINSCHOOL OCEAN PART GIA LÂM"}, {"code": "FGARDEN0087", "name": "FRESHGARDEN VINSCHOOL SÀI ĐỒNG"}, {"code": "FGARDEN0088", "name": "FRESHGARDEN TRƯỜNG PHỔ THÔNG VINSCHOOL OCEAN PART 2"}, {"code": "FGARDEN0089", "name": "FRESHGARDEN 08 HH3A KĐT LINH ĐÀM"}, {"code": "FGARDEN0090", "name": "FRESHGARDEN 90 SÀI ĐỒNG"}, {"code": "FGARDEN0091", "name": "FRESHGARDEN S1-09 VINHOMES OCEAN PART GIA LÂM"}, {"code": "FGARDEN0092", "name": "FRESHGARDEN R103 ZEN PARK"}, {"code": "FGARDEN0093", "name": "FRESHGARDEN 42W LÝ THƯỜNG KIỆT"}, {"code": "FGARDEN0094", "name": "FRESHGARDEN 46 NGUYỄN ĐỨC CẢNH"}, {"code": "FGARDEN0001", "name": "FRESHGARDEN Lô 8 cụm CN Lai Xá"}],
      parse(wb, fileName) { return parseFreshGarden(this, wb, fileName || ''); }
    },
    HIENLUONG: {
      label: 'Hiền Lương (.xls)', prefix: 'HIENLUONG', seq: 'dup', noAbbr: true,
      accept: '.xls,.xlsx', leadDays: 3, workDays: true, fromNgayVe: true,
      fixedCustomer: 'HIENLUONG0001',
      customers: [{ code: 'HIENLUONG0001', name: 'HIENLUONG Hòa Xá' }],
      parse(wb, fileName) { return parseHienLuong(this, wb, fileName || ''); }
    },
    HINODE: pdfAcct('Hinode (PDF)', parseHinode, {
      fixedCustomer: 'HINODE0001', customers: [{ code: 'HINODE0001', name: 'HINODE HOÀI ĐỨC' }]
    }),
    JAPAUS: pdfAcct('Japaus (PDF)', parseJapaus, {
      fixedCustomer: 'JAPAUS00001', customers: [{ code: 'JAPAUS00001', name: 'JAPAUS_193 Tôn Dật Tiên' }]
    }),
    SAKUKO: {
      label: 'Sakuko (SAKUKO)', prefix: 'SAKUKO', seq: 'always',
      accept: '.xlsx,.xls', leadDays: 3, workDays: true, fromNgayVe: true,
      // Số đơn: SAKUKO + chữ tắt cửa hàng (Hàng Bông -> HB) + YYMMDD ngày về đơn + -01
      brandWords: ['sakuko', 'daily', 'mart'],
      customers: [{"code": "SAKUKO0004", "name": "SAKUKO DAILY 165 XUÂN THỦY"}, {"code": "SAKUKO0005", "name": "SAKUKO DAILY NGUYỄN CÔNG HOAN"}, {"code": "SAKUKO0006", "name": "SAKUKO MART HÀNG BÔNG"}, {"code": "SAKUKO0001", "name": "SAKUKO DAILY 34 VŨ PHẠM HÀM"}, {"code": "SAKUKO0002", "name": "SAKUKO 126 LÊ TRỌNG TẤN"}, {"code": "SAKUKO0003", "name": "SAKUKO DAILY LING LANG"}],
      parse(wb, fileName) { return parseSakuko(this, wb, fileName || ''); }
    },
    SIBA: pdfAcct('Siba Food (PDF)', parseSiba, {
      // Số đơn = "Số đơn hàng" trên phiếu (chờ anh Tài chốt) · mã KH dò theo "Nơi nhận hàng"
      brandWords: ['siba', 'sibafood', 'food'],
      customers: [{"code": "SIBA010", "name": "SIBA_Nguyễn Thái Học"}, {"code": "SIBA004", "name": "SIBA Đặng Thái Thân"}, {"code": "SIBA005", "name": "SIBA An Dương Vương"}, {"code": "SIBA006", "name": "SIBA Hà Huy Tập"}, {"code": "SIBA007", "name": "SIBA Nguyễn Văn Cừ"}, {"code": "SIBA008", "name": "SIBA Nguyễn Trãi"}, {"code": "SIBA009", "name": "SIBA Nguyễn Trãi_Tổng kho hàng khô"}, {"code": "SIBAFOOD003", "name": "SIBA - SUNRISE"}, {"code": "SIBAFOOD004", "name": "SIBA - WESTGATE"}, {"code": "SIBAFOOD005", "name": "SIBA - MIZUKI"}, {"code": "SIBAFOOD006", "name": "SIBA - RESIDEN"}, {"code": "SIBAFOOD001", "name": "SIBA - Grand Park - GH201"}, {"code": "SIBAFOOD002", "name": "SIBA - Grand Park - BS1001"}, {"code": "SIBAFOOD011", "name": "SIBA - Lê Văn Khương"}, {"code": "SIBAFOOD012", "name": "SIBA - Chung cư Ehome 5"}, {"code": "SIBAFOOD008", "name": "SIBA - Green Town"}, {"code": "SIBAFOOD009", "name": "SIBA - Goldora Plaza"}, {"code": "SIBAFOOD007", "name": "SIBA - Grand Park - S06"}, {"code": "SIBAFOOD010", "name": "SIBA - Lê Đức Thọ"}]
    })

    ,
    ANHVIET: {
      label: 'Anh Việt Mart (ANHVIET)', prefix: 'ANHVIET', seq: 'dup', noAbbr: true,
      accept: '.xlsx,.xls', leadDays: 3, workDays: true, fromNgayVe: true,
      fixedCustomer: 'ANHVIET0001', customers: [{"code": "ANHVIET0001", "name": "Anh Việt Mart"}],
      parse(wb, fileName) { return parseAnhViet(this, wb, fileName || ''); }
    },
    CLEVER: {
      label: 'Clever Food (CLEVERFOOD)', prefix: 'CLEVER', seq: 'dup',
      accept: '.xlsx,.xls', leadDays: 3, workDays: true, fromNgayVe: true,
      // file: Barcode | Tên | <tên cửa hàng> (cột SL) | <mã CLEVERFOODxx> · số đơn CLEVER + chữ tắt + YYMMDD
      codeRe: /^CLEVERFOOD\d+$/i, codeHint: 'CLEVERFOODxx', brandWords: ['clever', 'food', 'cleverfood'],
      // chữ tắt = chữ cái đầu tên cửa hàng (bỏ số): LỮ ĐOÀN 21 LÊ ĐỨC THỌ -> LDLDT (anh Tài chốt 13/11)
      abbrOf(name) { return shortName(name, this.brandWords); },
      customers: [{"code": "CLEVERFOOD03", "name": "LỮ ĐOÀN 136 HỒ TÙNG MẬU"}, {"code": "CLEVERFOOD04", "name": "LỮ ĐOÀN NGHĨA ĐÔ"}, {"code": "CLEVERFOOD06", "name": "LỮ ĐOÀN 21 LÊ ĐỨC THỌ"}],
      parse(wb) { return storeSheetAcct(this, wb); }
    },
    FMKT: pdfAcct('Family Foods Market (PDF)', parseFamilyFoods, {
      brandWords: ['familyfoodmarket', 'family', 'foods', 'food', 'market'],
      customers: [{"code": "FMKTHNI0001", "name": "FAMILYFOODMARKET_06 Lê Văn Thiêm"}]
    }),
    KAISER: {
      label: 'Kaiser (CVS, nhiều sheet)', prefix: 'KAISER', seq: 'dup',
      accept: '.xlsx,.xls', leadDays: 3, workDays: true, fromNgayVe: true,
      // mỗi sheet = 1 cửa hàng = 1 đơn · cột SL = cột có tiêu đề tên cửa hàng, kế bên là mã CVS000xx · số đơn KAISER + chữ tắt + YYMMDD
      codeRe: /^CVS\d{3,}$/i, codeHint: 'CVS000xx', brandWords: ['kaiser', 'so', 'duong', 'pho', 'city'],
      abbrOf(name, sheet) { return shortName(String(name).replace(/^.*?_/, ''), this.brandWords) || shortName(sheet || '', this.brandWords); },
      customers: [{"code": "CVS00011", "name": "KAISER_Times City"}, {"code": "CVS00053", "name": "KAISER_Số 15 Nguyễn Văn Cừ"}, {"code": "CVS00008", "name": "KAISER_Số 37-39 Đường Phố Huế"}],
      parse(wb) { return storeSheetAcct(this, wb); }
    },
    OHMEE: pdfAcct('Ohmee (PDF)', parseOhmee, {
      brandWords: ['ohmee', 'kim', 'lam', 'enjoy', 'mart'],
      customers: [{"code": "OHMEE0003", "name": "OHMEE_VINHOMES SMART CITY"}, {"code": "OHMEE0004", "name": "OHMEE_B18, SỐ 53 PANDORA TRIỀU KHÚC"}, {"code": "OHMEE0008", "name": "OHMEE_ LINH ĐÀM"}, {"code": "OHMEE0009", "name": "OHMEE_ Vịnh Đảo Văn Giang Hưng Yên"}, {"code": "OHMEE0026", "name": "OHMEE_ S02 HẢI ĐĂNG HP"}, {"code": "OHMEE0027", "name": "OHMEE_ CT3 384 LÊ THÁNH TÔNG HP"}, {"code": "OHMEE0028", "name": "OHMEE_ 147 BẠCH ĐẰNG HẢI PHÒNG"}, {"code": "OHMEE0029", "name": "OHMEE_ KHU CN NGHĨA ĐÀN NGHỆ AN"}, {"code": "OHMEE0030", "name": "OHMEE_ 74 TRẦN THÁI TÔNG HN"}, {"code": "OHMEE0018", "name": "OHMEE_PHỐ TRÚC HƯNG YÊN"}, {"code": "OHMEE0019", "name": "OHMEE_ XUÂN PHƯƠNG HÀ NỘI"}, {"code": "OHMEE0020", "name": "OHMEE_ ROSE TOWN YÊN SỞ"}, {"code": "OHMEE0001", "name": "OHMEE_TRẦN KIM XUYẾN"}, {"code": "OHMEE0002", "name": "OHMEE_Kho DC"}, {"code": "OHMEE0021", "name": "OHMEE_ S2.06 OCEAN PARK HN"}, {"code": "OHMEE0022", "name": "OHMEE_ số 7 Tú Mỡ HN"}, {"code": "OHMEE0023", "name": "OHMEE_ I4 Imperia Smart City HN"}, {"code": "OHMEE0024", "name": "OHMEE_ HHB Tây Tân Đô HN"}, {"code": "OHMEE0025", "name": "OHMEE_ Tecco Daimond Thanh Trì HN"}, {"code": "OHMEE0005", "name": "OHMEE_ PHÁP VÂN TỨ HIỆP"}, {"code": "OHMEE0006", "name": "OHMEE_ BLUESTAR THÀNH TRUNG GIA LÂM"}, {"code": "OHMEE0007", "name": "OHMEE_R2 ECOPARK"}, {"code": "OHMEE0012", "name": "OHMEE_NGUYỄN BÌNH HƯNG YÊN"}, {"code": "OHMEE0013", "name": "OHMEE_ NGUYỄN TRÃI NGHỆ AN"}, {"code": "OHMEE0014", "name": "OHMEE_ NGỌC HỒI THANH TRÌ HN"}, {"code": "OHMEE0015", "name": "OHMEE_ NGUYỄN QUÝ ĐỨC HN"}, {"code": "OHMEE0016", "name": "OHMEE_ HỘI XÁ HN"}, {"code": "OHMEE0017", "name": "OHMEE_ NGÔ QUYỀN HẢI PHÒNG"}, {"code": "OHMEE0010", "name": "OHMEE_S08 Hồ Thiên Nga"}, {"code": "OHMEE0011", "name": "OHMEE_ S303 Vinhome Smart City"}]
    }),
    RFOOD: {
      label: 'Rfood (RFOODTP)', prefix: 'RFOOD', seq: 'dup', abbrAfterDate: true,
      accept: '.xlsx,.xls', leadDays: 3, workDays: true, fromNgayVe: true,
      // số đơn RFOOD + YYMMDD + chữ tắt (173 Marina -> MRN, 166 Nguyễn Văn Lộc -> NVL)
      codeRe: /^RFOODTP\d+$/i, codeHint: 'RFOODTPxxxx', brandWords: ['rfood'],
      abbrOf(name) { return shortName(name, this.brandWords, 'cons'); },
      customers: [{"code": "RFOODTP0002", "name": "RFOOD 166 Nguyễn Văn Lộc"}, {"code": "RFOODTP0005", "name": "RFOOD 40 Hoàng Như Tiếp"}, {"code": "RFOODTP0004", "name": "RFOOD 51 Trường Lâm"}, {"code": "RFOODTP0001", "name": "RFOOD 173 Marina"}],
      parse(wb) { return storeSheetAcct(this, wb); }
    },
    TOP5: {
      label: 'Top5 Mart (TOP5_MART)', prefix: 'TOP5MART', seq: 'dup', noAbbr: true,
      accept: '.xlsx,.xls', leadDays: 3, workDays: true, fromNgayVe: true,
      codeRe: /^TOP5_MART_\d+$/i, codeHint: 'TOP5_MART_xxx', brandWords: ['top5', 'mart'],
      customers: [{"code": "TOP5_MART_001", "name": "TOP5_MART Yên Phụ"}],
      parse(wb) { return storeSheetAcct(this, wb); }
    },
    WEFRESH: pdfAcct('Wefresh (PDF)', parseWefresh, {
      customers: [{"code": "WEFRESH0001", "name": "WEFRESH_174 Thái Hà"}]
    })

  };

  function parseAeonCsv(buf) {
    const text = new TextDecoder('utf-8').decode(buf);
    const rows = text.split(/\r?\n/).map(l => l.trim()).filter(l => l.startsWith('"') && l.endsWith('"')).map(l => l.slice(1, -1).split('","'));
    const hi = rows.findIndex(r => fold(r[0]) === 'loai nganh hang');
    if (hi < 0) return { error: 'Không thấy dòng tiêu đề của file CSV AEON ("Loai nganh hang"…). File này có đúng là PO AEON xuất CSV không?' };
    const H = rows[hi].map(h => fold(h).trim());
    const idx = (name, nth) => { let n = 0; for (let i = 0; i < H.length; i++) if (H[i] === name && n++ === (nth || 0)) return i; return -1; };
    const I = { po: idx('so don hang'), store: idx('ma cua hang'), storeName: idx('ten cua hang'), loc: idx('dia diem giao hang'), orderDate: idx('ngay dat hang'), note: idx('luu y'), name: idx('ten hang'), pack: idx('so luong', 0), unit: idx('don vi tinh'), item: idx('ma hang'), bc: idx('ma vach'), cartons: idx('so luong', 1) };
    const data = rows.slice(hi + 1);
    const sh = data.find(r => r[0] === 'SH'), head = sh || data.find(r => r[0] === 'D');
    if (!head) return { error: 'File CSV không có dòng đơn hàng nào.' };
    let items = data.filter(r => r[0] === 'SD'); if (!items.length) items = data.filter(r => r[0] === 'D');
    const num = v => { const x = Number(String(v || '').replace(/,/g, '')); return String(v || '').trim() !== '' && Number.isFinite(x) ? x : null; };
    const od = (head[I.orderDate] || '').trim();
    return { source: 'csv', po: (head[I.po] || '').trim(), storeCode: (head[I.store] || '').trim(), storeName: (head[I.storeName] || '').trim(), loc: (head[I.loc] || '').trim(),
      orderDate: /^\d{8}$/.test(od) ? `${od.slice(0, 4)}-${od.slice(4, 6)}-${od.slice(6, 8)}` : '', note: data.map(r => r[I.note] || '').join(' '),
      items: items.map(r => ({ barcode: (r[I.bc] || '').trim(), aeonName: (r[I.name] || '').trim(), item: (r[I.item] || '').trim(), pack: num(r[I.pack]), cartons: num(r[I.cartons]), unit: (r[I.unit] || '').trim() })) };
  }

  function aeonBuild(acct, d) {
    const issues = [];
    const { po, storeCode, storeName, loc, orderDate } = d;
    const hamper = /hamper/i.test(d.note) || d.items.some(r => /^HP-/i.test(r.aeonName || ''));
    const cands = acct.customers.filter(c => new RegExp('^' + storeCode + '\\s*-').test(c.name));
    let pick = null, alts = [];
    if (hamper) {
      const hp = cands.filter(c => /hamper/i.test(c.name));
      pick = hp.find(c => /hamper\s*_/i.test(c.name)) || hp[0] || null; alts = hp.filter(c => c !== pick);
      if (pick) issues.push({ level: 'warn', msg: `Đơn Hamper: chọn ${pick.code} · ${pick.name}.` + (alts.length ? ` Mã HP khác của cửa hàng ${storeCode}: ${alts.map(c => c.code + ' · ' + c.name).join('; ')}.` : '') + ' Kiểm tra lại.' });
    } else {
      const nm = cands.filter(c => !/hamper/i.test(c.name));
      const toStore = !/^dc\b/i.test(loc);
      pick = nm.find(c => /giao\s*c[uử]a\s*h[aà]ng/i.test(c.name) === toStore) || nm[0] || null; alts = nm.filter(c => c !== pick);
      if (pick && alts.length) issues.push({ level: 'warn', msg: `Chọn ${pick.code} · ${pick.name} (giao tại: ${loc || '?'}). Mã khác của cửa hàng ${storeCode}: ${alts.map(c => c.code + ' · ' + c.name).join('; ')}. Kiểm tra lại.` });
    }
    if (!pick && storeName) { const m = matchCustomer(acct.customers, storeName, /^aeon\b[\s_-]*/); if (m.hit) { pick = m.hit; issues.push({ level: 'warn', msg: `Dò theo tên cửa hàng: ${m.hit.code} · ${m.hit.name}. Kiểm tra lại.` }); } }
    if (!pick) issues.push({ key: 'customer', level: 'err', msg: `Không tìm thấy mã KH cho cửa hàng ${storeCode} ${storeName} trong DMKH. Chọn ở ô Mã khách hàng.` });
    if (!po) issues.push({ key: 'po', level: 'err', msg: 'Không đọc được Số đơn hàng. Nhập tay.' });
    if (!orderDate) issues.push({ level: 'warn', msg: 'Không đọc được Ngày đặt hàng; ngày giao tính theo ô Ngày về đơn.' });
    const lines = d.items.map((r, i) => {
      const qty = r.pack != null && r.cartons != null ? Math.round(r.pack * r.cartons) : (r.cartons ?? r.pack);
      const l = resolveAeon({ row: i + 1, barcode: r.barcode, aeonName: r.aeonName, item: r.item, pack: r.pack, cartons: r.cartons, qty, unit: r.unit, src: d.source, verified: r.verified });
      return l;
    });
    const sumQ = lines.reduce((a, l) => a + (l.qty || 0), 0);
    const sumAmt = d.items.reduce((a, r) => a + (r.amount || 0), 0);
    if (d.source === 'pdf' && d.total != null && Math.abs(sumAmt - d.total) > 1) issues.push({ level: 'warn', msg: `Tổng thành tiền các dòng (${sumAmt.toLocaleString('vi-VN')}) khác Tổng Tiền trên PDF (${d.total.toLocaleString('vi-VN')}). Có thể thiếu dòng hoặc đọc sai số.` });
    issues.push({ level: 'info', msg: `Đọc từ ${d.source === 'pdf' ? 'PDF' : 'CSV'}: cửa hàng ${storeCode} ${storeName}, giao tại ${loc || '?'}, ${lines.length} dòng, tổng ${sumQ.toLocaleString('vi-VN')} SL (= quy cách × số thùng)${d.source === 'pdf' && d.total != null && Math.abs(sumAmt - d.total) <= 1 ? ', tổng tiền khớp PDF' : ''}.` });
    return { fixedPo: true, po, customer: pick ? pick.code : '', store: `${storeCode} - ${storeName}${hamper ? ' (Hamper)' : ''}`, orderDate, lines, issues, aeon: true, source: d.source };
  }

  // AEON: barcode trên đơn -> product code điền form (bảng Code chuyển đổi trước, DMSP sau)
  function aeonMap() { return Object.assign({}, root.FFM_AEON_MAP_SHARED || {}, root.FFM_AEON_MAP || {}, root.FFM_AEON_MAP_USER || {}); }
  function resolveAeon(l) {
    const map = aeonMap(), PC = root.FFM_PC || {};
    const key = String(l.barcode).padStart(14, '0');
    const cands = (PC[key] || []).map(x => ({ code: x[0], name: x[1], active: !!x[2] }));
    l.cands = cands;
    if (map[l.barcode] || map[String(Number(l.barcode))]) { l.productCode = map[l.barcode] || map[String(Number(l.barcode))]; l.pcState = 'map'; }
    else if (cands.length === 1) { l.productCode = cands[0].code; l.pcState = 'dmsp'; }
    else if (cands.length > 1) { const base = cands.find(c => c.code.padStart(14, '0') === key); l.productCode = base ? base.code : ''; l.pcState = base ? 'multi' : 'pick'; }
    else { l.productCode = ''; l.pcState = 'none'; }
    return validateAeon(l);
  }
  function validateAeon(l) {
    l.flags = [];
    const c = (l.cands || []).find(x => x.code === l.productCode);
    l.name = c ? c.name : (l.productCode && root.FFM_PC ? (Object.values(root.FFM_PC).flat().find(x => x[0] === l.productCode) || [])[1] || '' : '');
    if (!l.productCode) l.flags.push({ level: 'err', msg: l.pcState === 'pick' ? `Barcode này có ${l.cands.length} mã sản phẩm trong DMSP. Chọn đúng mã.` : 'Barcode không có trong bảng Code chuyển đổi và DMSP. Nhập product code, rồi bổ sung vào bảng Code chuyển đổi.' });
    else if (l.pcState === 'multi') l.flags.push({ level: 'warn', msg: `Barcode này có ${l.cands.length} mã sản phẩm trong DMSP, đang chọn mã gốc. Kiểm tra lại; nếu hay gặp thì thêm vào bảng Code chuyển đổi.` });
    else if (l.pcState === 'manual' && !c) l.flags.push({ level: 'warn', msg: 'Product code nhập tay, không thấy trong DMSP.' });
    if (l.src === 'pdf') {
      const F = root.FUJI;
      if (F && l.barcode && !F.gtinOk(l.barcode)) l.flags.push({ level: 'warn', msg: 'Barcode đọc từ PDF sai số kiểm tra, có thể đọc nhầm. So với PDF.' });
      if (!l.verified) l.flags.push({ level: 'warn', msg: 'SL đặt × đơn giá ≠ thành tiền trên PDF. Kiểm tra số thùng.' });
    }
    if (l.qty == null || !(l.qty > 0)) l.flags.push({ level: 'err', msg: 'Không có số lượng' });
    else if (!Number.isInteger(l.qty)) l.flags.push({ level: 'err', msg: 'Số lượng lẻ thập phân' });
    return l;
  }

  // kiểm tra 1 dòng đọc từ PDF/ảnh
  function validateFuji(l) {
    l.flags = [];
    const F = root.FUJI || (typeof require !== 'undefined' ? require('./fuji.js') : null);
    if (!l.barcode) l.flags.push({ level: 'err', msg: 'Thiếu barcode' });
    else if (!F.gtinOk(l.barcode)) l.flags.push({ level: 'err', msg: 'Barcode sai số kiểm tra: còn ít nhất 1 chữ số chưa đúng. Sửa theo ảnh dòng rồi bấm Enter (viền ô chuyển xanh là đúng).' });
    else if (root.FFM_DMSP && !root.FFM_DMSP[l.barcode.padStart(14, '0')]) l.flags.push({ level: 'warn', msg: 'Barcode không có trong danh mục sản phẩm Hyperion (DMSP). So với ảnh dòng; nếu đúng là mã mới thì bỏ qua.' });
    else if (l.bcState === 'snap') l.flags.push({ level: 'info', msg: 'Ảnh mờ, barcode đã được khớp theo danh mục. Liếc ảnh dòng để chắc.' });
    else if (l.bcState === 'guess') l.flags.push({ level: 'warn', msg: 'Ảnh mờ, công cụ đã tự sửa 1 chữ số. So với ảnh dòng rồi sửa nếu khác.' });
    else if (l.bcState === 'weak') l.flags.push({ level: 'warn', msg: 'Ảnh mờ, barcode đọc chưa chắc. So với ảnh dòng rồi sửa nếu khác.' });
    if (l.qty == null || Number.isNaN(l.qty) || l.qty <= 0) l.flags.push({ level: 'err', msg: 'Không đọc được số lượng' });
    else if (!Number.isInteger(l.qty)) l.flags.push({ level: 'err', msg: 'Số lượng lẻ thập phân' });
    else if (l.price != null && l.amount != null && Math.abs(l.qty * l.price - l.amount) > 1) l.flags.push({ level: 'warn', msg: `SL × đơn giá ≠ thành tiền. Kiểm tra SL theo phiếu.` });
    else if (l.price == null || l.amount == null) l.flags.push({ level: 'warn', msg: 'Không đọc được đơn giá/thành tiền để kiểm tra SL.' });
    l.name = (root.FFM_DMSP && l.barcode && root.FFM_DMSP[l.barcode.padStart(14, '0')]) || '';
    return l;
  }

  /* ---------- Writer: fill Hyperion FORM_2_HP template ---------- */
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  /* FORM_5_HP (Miomio): giữ nguyên định dạng của form mẫu, chỉ thay giá trị ô.
     C9 số đơn · C10 mã KH · C13 ngày giao · dòng 20–39: D product code, E tên hàng, F số lượng. */
  async function buildFormProto(templateBytes, { po, customer, delivery, lines }) {
    if (lines.length > 20) throw new Error('Form FORM_5_HP chỉ có 20 dòng hàng (20–39); đơn này có ' + lines.length + ' dòng.');
    const zip = await JSZipLib.loadAsync(templateBytes);
    const sheetPath = 'xl/worksheets/sheet1.xml', sstPath = 'xl/sharedStrings.xml';
    let sheet = await zip.file(sheetPath).async('string');
    let sst = await zip.file(sstPath).async('string');
    let sis = (sst.match(/<si>/g) || []).length;
    const addS = text => { sst = sst.replace('</sst>', `<si><t xml:space="preserve">${esc(text)}</t></si></sst>`); return sis++; };
    // thay giá trị một ô, giữ kiểu định dạng (s="..") của ô trong form mẫu
    const put = (ref, val, kind) => {
      const re = new RegExp(`<c r="${ref}"((?: [^>]*?)?)(?:/>|>[\\s\\S]*?</c>)`);
      const m = sheet.match(re);
      const st = m ? ((m[1].match(/ s="\d+"/) || [''])[0]) : '';
      const xml = val == null || val === '' ? `<c r="${ref}"${st}/>`
        : kind === 'n' ? `<c r="${ref}"${st}><v>${val}</v></c>` : `<c r="${ref}"${st} t="s"><v>${addS(String(val))}</v></c>`;
      if (m) sheet = sheet.replace(re, xml);
      else { const rn = ref.replace(/^[A-Z]+/, ''); sheet = sheet.replace(new RegExp(`(<row r="${rn}"[^>]*>[\\s\\S]*?)(</row>)`), `$1${xml}$2`); }
    };
    put('C9', po); put('C10', customer); put('C13', excelSerial(delivery), 'n');
    for (let i = 0; i < 20; i++) {
      const n = 20 + i, l = lines[i];
      put(`C${n}`, '');
      put(`D${n}`, l ? (l.productCode || l.barcode) : '');
      put(`E${n}`, l ? (l.dmName || '') : '');
      put(`F${n}`, l ? l.qty : '', 'n');
    }
    sst = sst.replace(/count="\d+"/, `count="${sis}"`).replace(/uniqueCount="\d+"/, `uniqueCount="${sis}"`);
    zip.file(sheetPath, sheet); zip.file(sstPath, sst);
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  async function buildForm(templateBytes, { po, customer, delivery, lines }) {
    const zip = await JSZipLib.loadAsync(templateBytes);
    const sheetPath = 'xl/worksheets/sheet1.xml', sstPath = 'xl/sharedStrings.xml';
    let sheet = await zip.file(sheetPath).async('string');
    let sst = await zip.file(sstPath).async('string');
    let sis = (sst.match(/<si>/g) || []).length;
    const addS = text => { sst = sst.replace('</sst>', `<si><t xml:space="preserve">${esc(text)}</t></si></sst>`); return sis++; };
    const setCell = (ref, xml) => {
      const re = new RegExp(`<c r="${ref}"(?: [^>]*?)?(?:/>|>[\\s\\S]*?</c>)`);
      if (!re.test(sheet)) throw new Error('Template thiếu ô ' + ref);
      sheet = sheet.replace(re, xml);
    };
    setCell('C9', `<c r="C9" s="43" t="s"><v>${addS(po)}</v></c>`);
    setCell('C10', `<c r="C10" s="43" t="s"><v>${addS(customer)}</v></c>`);
    setCell('C13', `<c r="C13" s="38"><v>${excelSerial(delivery)}</v></c>`);

    // rebuild rows >= 20
    const rows = [...sheet.matchAll(/<row r="(\d+)"[\s\S]*?<\/row>/g)];
    const first20 = rows.find(m => +m[1] >= 20);
    const startIdx = first20.index;
    const endIdx = sheet.indexOf('</sheetData>');
    const tail27 = 'JKLMNOPQRSTUVWXYZ'.split('').map(c => `<c r="${c}20" s="27"/>`).join('') + ['AA', 'AB'].map(c => `<c r="${c}20" s="27"/>`).join('');
    const attrs = n => `<row r="${n}" spans="1:28" ht="56.25" customHeight="1" x14ac:dyDescent="0.4">`;
    let out = '';
    lines.forEach((l, i) => {
      const n = 20 + i;
      const cells = (i === 0 ? `<c r="A${n}" s="35"/><c r="B${n}" s="50" t="s"><v>${addS('1')}</v></c>` : '') +
        (l.productCode ? `<c r="C${n}" s="55"/><c r="D${n}" s="46" t="s"><v>${addS(l.productCode)}</v></c><c r="E${n}" s="47"/>` : `<c r="C${n}" s="55" t="s"><v>${addS(l.barcode)}</v></c><c r="D${n}" s="46"/><c r="E${n}" s="47"/>`) +
        `<c r="F${n}" s="56"><v>${l.qty}</v></c><c r="G${n}" s="48"/>` +
        (i === 0 ? `<c r="H${n}" s="51"/><c r="I${n}" s="52"/>` + tail27 : '');
      out += attrs(n) + cells + '</row>';
    });
    const last = Math.max(37, 19 + lines.length);
    for (let n = 20 + lines.length; n <= last; n++)
      out += attrs(n) + `<c r="C${n}" s="53"/><c r="D${n}" s="46"/><c r="E${n}" s="47"/><c r="F${n}" s="54"/></row>`;
    sheet = sheet.slice(0, startIdx) + out + sheet.slice(endIdx);
    sheet = sheet.replace(/<dimension ref="[^"]*"\/>/, `<dimension ref="A1:AB${last}"/>`);
    sst = sst.replace(/count="\d+"/, `count="${sis}"`).replace(/uniqueCount="\d+"/, `uniqueCount="${sis}"`);
    zip.file(sheetPath, sheet); zip.file(sstPath, sst);
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  const api = { validateAeon, resolveAeon, validateFuji, matchCustomer, ACCOUNTS, buildForm, buildFormProto, fmtDate, excelSerial, assignPO, toISO, addDays };
  if (typeof module !== 'undefined') module.exports = api; else root.FFM = api;
})(typeof window !== 'undefined' ? window : globalThis);
