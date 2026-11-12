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
      if (!r.abbr || !ngayVe) { r.po = ''; return; }
      const k = r.prefix + r.abbr;
      count[k] = (count[k] || (startSeq && startSeq[k]) || 0) + 1;
      const base = `${r.prefix}${r.abbr}${yymmdd(ngayVe)}`;
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
    for (let c = range.s.c; c <= Math.min(range.e.c, 30); c++) {
      const h = cellText(ws, hr, c), l = low(h);
      if (!h || c === colCode) continue;
      if (['stt', 'no', 'no.', '#'].includes(l)) colStt = c;
      else if (l.startsWith('tên') || l.includes('name')) colName = c;
      else if (l === 'giá' || l === 'price') colPrice = c;
      else if (/^[A-Z]{3,}\d{3,}$/.test(h.toUpperCase()) && colCust < 0) { colCust = c; customer = h.toUpperCase(); }
    }
    let colQty = -1;
    for (let c = range.s.c; c <= Math.min(range.e.c, 30); c++) {
      if ([colCode, colName, colPrice, colCust, colStt].includes(c)) continue;
      if (cellText(ws, hr, c)) { colQty = c; store = cellText(ws, hr, c); break; }
    }
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

  /* ---------- Account rules ---------- */
  const ACCOUNTS = {
    KKTC: {
      label: 'Kmart (KKTC)', prefix: 'KKTC', seq: 'always', brand: KKTC_BRAND,
      accept: '.xlsx,.xls',
      leadDays: 4, // ngày giao = ngày về đơn + 4 (10/11 -> 14/11)
      // DMKH: mã KH + Tên nội bộ (lọc từ DMKH.xlsx, 45 cửa hàng KKTC, Active)
      customers: [{"code": "KKTC0064", "name": "K-Market 126 Nguyễn Cao - Bắc Ninh"}, {"code": "KKTC0068", "name": "K-MARKET SMART CITY"}, {"code": "KKTC0041", "name": "Kmart The Emerald"}, {"code": "KKTC0056", "name": "KKTC_ Mỹ Đình Pearl"}, {"code": "KKTC0067", "name": "K-MARKET 117 TRẦN DUY HƯNG"}, {"code": "KKTC0050", "name": "K-Market West Point"}, {"code": "KKTC0072", "name": "K MARTKET LONG BIÊN- Vinhome Symphony"}, {"code": "KKTC0036", "name": "Kmarket Ecopark"}, {"code": "KKTC0053", "name": "KKTC_S3 Shunshine City"}, {"code": "KKTC0043", "name": "K-Market Vinhomes Skylake S3"}, {"code": "KKTC0042", "name": "K-Market Thăng Long Number"}, {"code": "KKTC0052", "name": "KKTC_Dcapitale C6"}, {"code": "KKTC0062", "name": "K-Market Green Bay - Hạ Long - Quảng Ninh"}, {"code": "KKTC0063", "name": "K-Market Chelsea House - Văn Cao - Hải Phòng"}, {"code": "KKTC0057", "name": "KKTC_K-Market Vinhome Bắc Ninh"}, {"code": "KKTC0069", "name": "K-MARKET KTX SAMSUNG"}, {"code": "KKTC0013", "name": "Kmart Trung Hòa"}, {"code": "KKTC0014", "name": "Kmart Calidas"}, {"code": "KKTC0015", "name": "Kmart CipuTra"}, {"code": "KKTC0016", "name": "Kmart Long Biên"}, {"code": "KKTC0017", "name": "Kmart Royal City R1"}, {"code": "KKTC0023", "name": "Kmart Goldmart Ruby"}, {"code": "KKTC0025", "name": "Kmart Quang Minh"}, {"code": "KKTC0028", "name": "Kmart ciputra 2"}, {"code": "KKTC0051", "name": "KKTC_ Phạm văn Đồng"}, {"code": "KKTC0058", "name": "K-market 148 Xuân Diệu"}, {"code": "KKTC0066", "name": "K-MARKET 88 LÊ QUANG ĐẠO - ĐÀ NẴNG"}, {"code": "KKTC0059", "name": "K-MARKET VÕ NGUYÊN GIÁP- ĐÀ NẴNG"}, {"code": "KKTC0071", "name": "K-market CT9 Mỹ Đình"}, {"code": "KKTC0065", "name": "K-Market Dự án The Minato Residence - Hải Phòng"}, {"code": "KKTC0047", "name": "K-Market Việt Long Bắc Ninh"}, {"code": "KKTC0034", "name": "Kmarket D-Capitale"}, {"code": "KKTC0060", "name": "K-Market 342 Ngọc Hân Công Chúa - Bắc Ninh"}, {"code": "KKTC0002", "name": "Kmart Keangnam"}, {"code": "KKTC0003", "name": "Kmart Mỹ Đình"}, {"code": "KKTC0004", "name": "Kmart 17T3"}, {"code": "KKTC0005", "name": "Kmart Golden"}, {"code": "KKTC0007", "name": "Kmart Gardenia"}, {"code": "KKTC0008", "name": "Goldmak saphia"}, {"code": "KKTC0012", "name": "Kmarket Greenbay"}, {"code": "KKTC0039", "name": "K-market MONARCHY ĐÀ NẴNG"}, {"code": "KKTC0031", "name": "Kmart Vinhome Sky Lake"}, {"code": "KKTC0061", "name": "K-Market The Matrix One"}, {"code": "KKTC0049", "name": "K-Market Starlake Tây Hồ"}, {"code": "KKTC0055", "name": "KKTC_ TT4 - Mỹ Đình 1"}],
      // Số đơn: KKTC + 6 ký tự đầu tên cửa hàng (bỏ Kmart, bỏ "The", bỏ dấu) + YYMMDD + -01/-02
      makeAbbr(store) { return { abbr: key(store, KKTC_BRAND).split(' ').filter(w => w && w !== 'the').join('').toUpperCase().slice(0, 6) }; },
      parse(wb) { return parseStoreColumn(this, wb); }
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
  function aeonMap() { return root.FFM_AEON_MAP_USER || root.FFM_AEON_MAP || {}; }
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

  const api = { validateAeon, resolveAeon, validateFuji, matchCustomer, ACCOUNTS, buildForm, fmtDate, excelSerial, assignPO, toISO, addDays };
  if (typeof module !== 'undefined') module.exports = api; else root.FFM = api;
})(typeof window !== 'undefined' ? window : globalThis);
