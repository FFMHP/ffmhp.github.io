/* BẢN THỬ (beta) · Team khác › Chuyển đổi file up DMS (team Đơn hàng GT).
   Đầu vào: "BẢNG KÊ CHI TIẾT HÓA ĐƠN (TẤT CẢ)" xuất từ phần mềm (vd 18112026143407PM.xlsx), 1 hoặc nhiều file.
   Quy tắc (suy từ "File final up DMS.xls", NPP Minh Tuấn Kiệt 06-17/11/2026):
   - Chỉ lấy Loại phiếu "Xuất bán KH" và "Xuất khuyến mãi" đã có số HĐ; bỏ đổi trả, hàng mẫu, xuất khác, hủy combo, "Chưa xuất HĐ".
   - Mã KH bỏ đuôi "HP" (TUANKIET001HP -> TUANKIET001). Mặc định chỉ xuất 7 NPP up DMS (DMS_KH / DMS_TEN bên dưới).
   - Dòng KM = "Xuất khuyến mãi" (hoặc số đơn "KM-…"). Số đơn KM bỏ tiền tố "KM-" để ghép với đơn bán.
   - Gộp KM vào hóa đơn bán cùng số đơn: cộng số lượng theo Mã vật tư + LOT + Ngày hết hạn.
     Số HĐ, ngày HĐ, ngày phiếu xuất, ngày đặt hàng lấy của hóa đơn bán.
   - Tiền tính lại theo giá bán: Tiền hàng = SL × Giá; CK và VAT theo đúng % của dòng bán cùng mã
     (thường CK 5%, VAT 8%, bia 10%). KM là mã không có trong đơn bán: lấy giá cùng mã ở đơn khác trong file, không có thì báo đỏ.
   - Bỏ trống Mã phiếu xuất, Mã phiếu đề xuất, Diễn giải, Giá vốn, Tiền vốn, Mã phiếu đề xuất HP; thêm 5 cột cuối để trống.
   - Mỗi Mã KH (NPP) một file .xls, sheet "Import". Kèm bảng đối chiếu kiểu UP WEB (PO bán / HĐ KM / SL UP / HĐ UP). */
(function () {
  const $ = id => document.getElementById(id);
  if (!$('dmDrop')) return;
  const fold = s => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase().replace(/\s+/g, ' ').trim();
  const num = v => typeof v === 'number' ? v : (Number(String(v == null ? '' : v).replace(/,/g, '').replace(/[^\d.-]/g, '')) || 0);
  const str = v => v == null ? '' : String(v).trim();
  const vnd = v => Math.round(v).toLocaleString('vi-VN');
  const el = (t, c, x) => { const e = document.createElement(t); if (c) e.className = c; if (x != null) e.textContent = x; return e; };
  const p2 = n => String(n).padStart(2, '0');
  const dstr = v => {
    if (v instanceof Date && !isNaN(v)) return `${p2(v.getDate())}/${p2(v.getMonth() + 1)}/${v.getFullYear()}`;
    if (typeof v === 'number' && v > 20000 && v < 80000) { const d = XLSX.SSF.parse_date_code(v); return `${p2(d.d)}/${p2(d.m)}/${d.y}`; }
    return str(v);
  };
  const r4 = x => Math.round(x * 10000) / 10000;
  const HEAD = ['Số TT', 'Loại phiếu', 'Ngày phiếu xuất', 'Mã phiếu xuất', 'Mã phiếu đề xuất', 'Ngày hóa đơn', 'Số hóa đơn', 'Mã KH / NCC\n(Customer Code / Vendor Code)', 'Tên KH 2 / NCC', 'Tên KH / NCC\n(Customer Name / Vendor Name)', 'Mã chủ hàng', 'Tên chủ hàng', 'Chi nhánh', 'Diễn giải\n(Description)', 'Mã vật tư\n(Goods Code)', 'BarCode', 'Tên vật tư\n(Goods Name)', 'Tên tiếng Anh\nEnglish Name)', 'Đơn vị\n(Unit)', 'Số lượng\n(Quantily)', 'Giá bán\n(Price)', 'Tiền hàng\n(Sub Total)', 'Tiền CK\n(Discount)', 'Tiền thuế\n(VAT)', 'Tồng tiền tt\n(Grand Total)', '', '', 'Mã kho\n(Warehouse Code)', 'Nhóm vật tư chính (Goods KeyGroup)', 'Nhóm vật tư phụ 1 - (Goods SubGroup1)', 'Nhóm vật tư phụ 2 - (Goods SubGroup2)', 'Nhóm vật tư phụ 3 - (Goods SubGroup3)', 'LOT', 'HSD', 'Ngày SX (Manufactur Date)', 'Ngày hết hạn (Expiry Date)', 'Số đơn đặt hàng', 'Ngày đặt hàng', 'Diễn giải', 'Mã Phiếu đề xuất của HP/ RW', 'Tổng số lượng trả lại', 'Giá trị trả lại', 'Giá trị giảm giá', 'Mã nhóm khách hàng', 'Scheme'];
  // 7 NPP up DMS (anh Tài chốt 18/11/2026): theo mã KH, hoặc tên KH có chữ tương ứng
  const DMS_KH = ['DD00008', 'TUANKIET001', 'DD00003', 'KINGBAC0001', 'NQKHUE0002', 'DD00010', 'DD00011'];
  const DMS_TEN = ['hoang huy phat', 'tuan kiet', 'an phat', 'kinh bac', 'quynh phat', 'ngoc an khang', 'thao huong tra'];
  // NPP team tự thêm: lưu chung sheet "Mã KH" (Apps Script, event "Thêm mã KH mới", acct = "DMS"), đọc lại qua ?t=aeonmap -> cust
  const normKh = c => String(c || '').trim().toUpperCase().replace(/HP$/, '');
  const added = () => (window.FFM_CUST_SHARED || []).filter(x => x.acct === 'DMS');
  const isDms = (kh, o) => DMS_KH.includes(String(kh).toUpperCase()) || added().some(x => normKh(x.code) === String(kh).toUpperCase()) || DMS_TEN.some(t => new RegExp('(^|[^a-z])' + t + '($|[^a-z])').test(fold(o.kh2 + ' | ' + o.tkh)));
  let files = [], R = null;

  const drop = $('dmDrop'), pick = $('dmFile');
  const setFiles = fl => { fl = [...(fl || [])].filter(f => /\.xlsx?$/i.test(f.name)); if (!fl.length) return; files = fl; $('dmName').textContent = fl.length === 1 ? fl[0].name : `${fl.length} file: ` + fl.map(f => f.name).join(', '); $('dmGo').disabled = false; $('dmMsg').textContent = ''; };
  drop.onclick = () => pick.click();
  drop.onkeydown = e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); pick.click(); } };
  pick.onchange = () => { setFiles(pick.files); pick.value = ''; };
  drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('over'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('over'));
  drop.addEventListener('drop', e => { e.preventDefault(); drop.classList.remove('over'); setFiles(e.dataTransfer.files); });

  // ---- đọc file: tìm dòng tiêu đề có "Số hóa đơn" + "Mã vật tư" + "Số lượng" ----
  function readFile(buf, name) {
    const wb = XLSX.read(buf, { type: 'array', cellDates: true, dense: true });
    for (const sn of wb.SheetNames) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: '' });
      const h = rows.slice(0, 30).findIndex(r => r.some(x => /^so hoa don/.test(fold(x))) && r.some(x => /^ma vat tu/.test(fold(x))) && r.some(x => /^so luong/.test(fold(x))));
      if (h < 0) continue;
      const H = rows[h].map(fold);
      const col = re => H.findIndex(x => re.test(x));
      const cols = re => H.map((x, i) => re.test(x) ? i : -1).filter(i => i >= 0);
      const C = {
        stt: col(/^so ?tt/), loai: col(/^loai phieu/), npx: col(/^ngay phieu xuat/), nhd: col(/^ngay hoa don/), hd: col(/^so hoa don/),
        kh: col(/^ma kh/), kh2: col(/^ten kh 2/), tkh: col(/^ten kh \/ ncc/), mch: col(/^ma chu hang/), tch: col(/^ten chu hang/), cn: col(/^chi nhanh/),
        mvt: col(/^ma vat tu/), bc: col(/^barcode/), ten: col(/^ten vat tu/), en: col(/^ten tieng anh/), dv: col(/^don vi/),
        sl: col(/^so luong/), gia: col(/^gia ban/), th: col(/^tien hang/), ck: col(/^tien ck/), vat: col(/^tien thue/), tt: col(/^tong tien tt|^tong tien/),
        kho: col(/^ma kho/), g0: col(/^nhom vat tu chinh/), g1: col(/^nhom vat tu phu 1/), g2: col(/^nhom vat tu phu 2/), g3: col(/^nhom vat tu phu 3/),
        lot: col(/^lot$/), hsd: col(/^hsd$/), nsx: col(/^ngay sx/), nhh: col(/^ngay het han/), ndh: col(/^ngay dat hang/),
        so: cols(/^so don dat hang/)
      };
      const miss = [['hd', 'Số hóa đơn'], ['mvt', 'Mã vật tư'], ['sl', 'Số lượng'], ['gia', 'Giá bán'], ['kh', 'Mã KH']].filter(([k]) => C[k] < 0).map(x => x[1]);
      if (!C.so.length) miss.push('Số đơn đặt hàng');
      if (miss.length) throw new Error(`${name}: thiếu cột ${miss.join(', ')}`);
      const g = (r, k) => C[k] >= 0 ? r[C[k]] : '';
      const out = [];
      out.skip = {}; out.noHd = 0;
      rows.slice(h + 1).forEach((r, i) => {
        const mvt = str(g(r, 'mvt')), hdv = g(r, 'hd'), hd = typeof hdv === 'number' ? String(hdv).padStart(8, '0') : str(hdv);
        if (!mvt || !hd) return;
        // chỉ lấy "Xuất bán KH" và "Xuất khuyến mãi"; bỏ đổi trả, hàng mẫu, xuất khác, hủy combo…
        const loai = str(g(r, 'loai')) || 'Xuất bán KH', fl = fold(loai);
        const isKmLoai = /khuyen mai/.test(fl);
        if (!isKmLoai && fl !== 'xuat ban kh') { out.skip[loai] = (out.skip[loai] || 0) + 1; return; }
        if (!/\d/.test(hd)) { out.noHd++; return; } // "Chưa xuất HĐ"
        const soRaw = C.so.map(c => str(r[c])).find(Boolean) || '';
        const isKmSo = /^KM[-_ ]?/i.test(soRaw);
        const so = soRaw.replace(/^KM[-_ ]?/i, '').trim();
        const gia = num(g(r, 'gia')), th = num(g(r, 'th'));
        out.push({
          file: name, line: h + 2 + i, loai: 'Xuất bán KH', npx: dstr(g(r, 'npx')), nhd: dstr(g(r, 'nhd')), hd,
          kh: str(g(r, 'kh')).replace(/HP$/i, ''), kh2: str(g(r, 'kh2')), tkh: str(g(r, 'tkh')), mch: str(g(r, 'mch')), tch: str(g(r, 'tch')), cn: str(g(r, 'cn')),
          mvt, bc: str(g(r, 'bc')), ten: str(g(r, 'ten')), en: str(g(r, 'en')), dv: str(g(r, 'dv')),
          sl: num(g(r, 'sl')), gia, th, ck: num(g(r, 'ck')), vat: num(g(r, 'vat')),
          kho: str(g(r, 'kho')), g0: str(g(r, 'g0')), g1: str(g(r, 'g1')), g2: str(g(r, 'g2')), g3: str(g(r, 'g3')),
          lot: str(g(r, 'lot')), hsd: str(g(r, 'hsd')), nsx: dstr(g(r, 'nsx')), nhh: dstr(g(r, 'nhh')), ndh: dstr(g(r, 'ndh')),
          so, soRaw, km: isKmLoai || isKmSo
        });
      });
      return out;
    }
    throw new Error(`${name}: không thấy dòng tiêu đề có "Số hóa đơn", "Mã vật tư", "Số lượng". File này có đúng là chi tiết hóa đơn xuất từ phần mềm không?`);
  }

  // ---- gộp KM vào hóa đơn bán ----
  function convert(lines) {
    const warn = [];
    let cur = '';
    const W = (lv, t) => warn.push({ lv, t, kh: cur });
    // giá bán theo mã (đơn bán) để dùng khi KM là mã không có trong đơn
    const priceBy = new Map();
    lines.forEach(o => { if (!o.km && o.gia > 0 && o.th > 0) { const k = o.mvt; if (!priceBy.has(k)) priceBy.set(k, { gia: o.gia, ckr: r4(o.ck / o.th), vr: o.th - o.ck ? r4(o.vat / (o.th - o.ck)) : 0, src: o }); } });
    // đơn: theo (Mã KH, số đơn) giữ thứ tự xuất hiện
    const orders = new Map();
    lines.forEach(o => {
      const k = o.kh + '|' + o.so;
      let d = orders.get(k); if (!d) orders.set(k, d = { kh: o.kh, so: o.so, sale: [], km: [] });
      (o.km ? d.km : d.sale).push(o);
    });
    const npps = new Map();
    orders.forEach(d => {
      cur = d.kh;
      const head = d.sale[0] || d.km[0];
      const saleHd = [...new Set(d.sale.map(o => o.hd))], kmHd = [...new Set(d.km.map(o => o.hd))];
      if (!d.sale.length) W('err', `Đơn ${d.so} (${d.kh}): chỉ có hàng KM (HĐ ${kmHd.join(', ')}), không có hóa đơn bán để gộp. Đã giữ nguyên dòng KM, giá 0.`);
      if (saleHd.length > 1) W('warn', `Đơn ${d.so}: có ${saleHd.length} hóa đơn bán (${saleHd.join(', ')}). Hàng KM gộp vào HĐ ${saleHd[0]}, kiểm tra lại.`);
      const main = d.sale.find(o => o.hd === saleHd[0]) || head;
      // gom dòng: theo HĐ (bán) + mã + LOT + hạn dùng
      const items = new Map();
      const add = (o, isKm) => {
        const hd = isKm ? main.hd : o.hd;
        const k = [hd, o.mvt, o.lot, o.nhh].join('|');
        let it = items.get(k);
        if (!it) items.set(k, it = { base: isKm ? Object.assign({}, o, { hd: main.hd, nhd: main.nhd, npx: main.npx, ndh: main.ndh, loai: main.loai }) : o, slBan: 0, slKm: 0, gia: 0, ckr: null, vr: null });
        if (isKm) it.slKm += o.sl; else { it.slBan += o.sl; if (!it.gia && o.gia) { it.gia = o.gia; it.ckr = o.th ? r4(o.ck / o.th) : 0; it.vr = o.th - o.ck ? r4(o.vat / (o.th - o.ck)) : 0; } }
      };
      d.sale.forEach(o => add(o, false));
      d.km.forEach(o => add(o, !!d.sale.length));
      const ckOrder = d.sale.length ? r4(d.sale[0].th ? d.sale[0].ck / d.sale[0].th : 0) : 0;
      let n = npps.get(d.kh); if (!n) npps.set(d.kh, n = { kh: d.kh, ten: head.kh2 || head.tkh, head, isNpp: isDms(d.kh, head), rows: [], orders: [], bad: 0 });
      items.forEach(it => {
        const o = it.base;
        if (!it.gia) {
          // dòng chỉ có KM: cùng mã trong đơn bán này (khác LOT) thì lấy giá đó, không thì lấy giá cùng mã ở đơn khác
          const same = d.sale.find(s => s.mvt === o.mvt && s.gia > 0 && s.th > 0);
          if (same) { it.gia = same.gia; it.ckr = r4(same.ck / same.th); it.vr = same.th - same.ck ? r4(same.vat / (same.th - same.ck)) : 0; }
        }
        if (!it.gia) {
          const p = priceBy.get(o.mvt);
          if (p && d.sale.length) { it.gia = p.gia; it.ckr = ckOrder || p.ckr; it.vr = p.vr; W('warn', `Đơn ${d.so}: KM mã ${o.mvt} (${o.ten}) không có trong hóa đơn bán, lấy giá ${vnd(p.gia)} của đơn ${p.src.so}.`); }
          else if (d.sale.length) { it.gia = 0; it.ckr = 0; it.vr = 0; it.bad = true; n.bad++; W('err', `Đơn ${d.so}: KM mã ${o.mvt} (${o.ten}) không tìm được giá bán trong file. Dòng này để giá 0, điền giá trước khi up.`); }
          else { it.ckr = 0; it.vr = 0; }
        }
        const sl = it.slBan + it.slKm, th = sl * it.gia, ck = th * it.ckr, vat = (th - ck) * it.vr;
        n.rows.push({ o, sl, gia: it.gia, th, ck, vat, tt: th - ck + vat, bad: !!it.bad, so: d.so });
      });
      n.orders.push({ npp: head.kh2 || head.tkh, so: d.so, kmSo: d.km.length && /^KM/i.test(d.km[0].soRaw) ? d.km[0].soRaw : 'KM-' + d.so, kmHd: kmHd.join(', '), slKm: d.km.reduce((s, o) => s + o.sl, 0), hd: saleHd.join(', '), slBan: d.sale.reduce((s, o) => s + o.sl, 0) });
    });
    return { npps: [...npps.values()], warn, lines: lines.length };
  }

  function toAoa(n) {
    const aoa = [HEAD];
    n.rows.forEach((x, i) => {
      const o = x.o;
      aoa.push([i + 1, o.loai, o.npx, '', '', o.nhd, o.hd, o.kh, o.kh2, o.tkh, o.mch, o.tch, o.cn, '', o.mvt, o.bc, o.ten, o.en, o.dv,
        x.sl, x.gia, x.th, x.ck, x.vat, x.tt, '', '', o.kho, o.g0, o.g1, o.g2, o.g3, o.lot, o.hsd, o.nsx, o.nhh, x.so, o.ndh, '', '', '', '', '', '', '']);
    });
    return aoa;
  }
  function fileName(n) {
    const t = new Date();
    return `File up DMS ${n.kh} ${p2(t.getDate())}.${p2(t.getMonth() + 1)}.${String(t.getFullYear()).slice(2)}.xls`;
  }
  function toXls(n) {
    const ws = XLSX.utils.aoa_to_sheet(toAoa(n));
    // mã/HĐ/số đơn luôn là chữ (giữ số 0 đầu)
    ws['!cols'] = HEAD.map((h, i) => ({ wch: [16, 14, 9, 36].includes(i) ? 18 : 12 }));
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Import');
    return new Blob([XLSX.write(wb, { bookType: 'biff8', type: 'array' })], { type: 'application/vnd.ms-excel' });
  }
  const save = (blob, name) => { const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.append(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000); };
  function logDm(d) {
    const API = window.FFM_API; if (!/^https:/.test(API || '')) return;
    try { fetch(API, { method: 'POST', mode: 'no-cors', keepalive: true, headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(Object.assign({ team: 'khac', user: (($('who') || {}).value || '').trim(), acct: 'UP DMS', files: files.length, ver: (($('ver') || {}).textContent || '').replace('Phiên bản ', '') + ' beta' }, d)) }).catch(() => {}); } catch (e) {}
  }

  // ---- hiển thị: cảnh báo + nút tải ở đầu, rồi từng NPP ----
  // khách đang chọn: mặc định chỉ khách tên "NPP …"
  const shown = () => { const v = $('dmKh').value; return R.npps.filter(n => v === '*' ? true : v === 'npp' ? n.isNpp : n.kh === v); };
  function fillKh() {
    R.npps.forEach(n => { n.isNpp = isDms(n.kh, n.head); });
    const sel = $('dmKh'), cur = sel.value; sel.innerHTML = '';
    const nNpp = R.npps.filter(n => n.isNpp).length;
    [['npp', `NPP up DMS (${nNpp} có trong file)`], ['*', `Tất cả khách (${R.npps.length})`]].concat(R.npps.map(n => [n.kh, `${n.kh} · ${n.ten}`])).forEach(([v, t]) => { const o = el('option', '', t); o.value = v; sel.append(o); });
    sel.value = [...sel.options].some(o => o.value === cur) ? cur : (nNpp ? 'npp' : '*');
  }
  function show() {
    const NS = shown(), KS = new Set(NS.map(n => n.kh));
    const wl = $('dmWarn'); wl.innerHTML = '';
    const ws = R.warn.filter(w => KS.has(w.kh));
    const errs = ws.filter(w => w.lv === 'err'), warns = ws.filter(w => w.lv !== 'err');
    errs.concat(warns).slice(0, 40).forEach(w => wl.append(el('li', w.lv, (w.lv === 'err' ? '❌ ' : '⚠️ ') + w.t)));
    if (ws.length > 40) wl.append(el('li', 'warn', `… và ${ws.length - 40} cảnh báo khác.`));
    wl.hidden = !ws.length;
    const kpi = $('dmKpi'); kpi.innerHTML = '';
    const k = (lab, v, sub) => { const d = el('div', 'tile'); d.append(el('small', '', lab), el('b', '', v)); if (sub) d.append(el('small', 'note', sub)); kpi.append(d); };
    const all = NS.flatMap(n => n.rows), ords = NS.flatMap(n => n.orders);
    k('Số khách', String(NS.length), NS.map(n => n.kh).slice(0, 8).join(', ') + (NS.length > 8 ? '…' : ''));
    k('Số đơn', String(ords.length), 'hóa đơn bán: ' + new Set(ords.map(o => o.hd)).size);
    k('Số dòng file up', all.length.toLocaleString('vi-VN'), R.info);
    k('Tổng SL up', all.reduce((s, x) => s + x.sl, 0).toLocaleString('vi-VN'), 'trong đó KM: ' + ords.reduce((s, o) => s + o.slKm, 0).toLocaleString('vi-VN'));
    k('Tổng tiền tt', vnd(all.reduce((s, x) => s + x.tt, 0)) + ' đ');
    const list = $('dmList'); list.innerHTML = '';
    NS.forEach((n, i) => {
      const d = el('details', 'gh-acc'); d.open = i === 0 || NS.length === 1;
      const s = el('summary');
      s.append(el('b', '', `${n.kh} · ${n.ten}`), el('span', 'note', ` · ${n.orders.length} đơn · ${n.rows.length} dòng · SL ${n.rows.reduce((a, x) => a + x.sl, 0).toLocaleString('vi-VN')} · ${vnd(n.rows.reduce((a, x) => a + x.tt, 0))} đ`));
      if (n.bad) s.append(el('span', 'gh-v', ` · ${n.bad} dòng thiếu giá`));
      d.append(s);
      const act = el('div', 'rp-act');
      const b1 = el('button', 'primary small', '⬇ Tải ' + fileName(n)); b1.type = 'button';
      b1.onclick = () => { save(toXls(n), fileName(n)); logDm({ event: 'Tải file up DMS', orders: n.orders.length, names: `${fileName(n)} | ${n.rows.length} dòng` }); };
      const b2 = el('button', 'small', '📋 Chép bảng UP WEB'); b2.type = 'button';
      b2.onclick = async () => { try { await navigator.clipboard.writeText(upweb(n)); b2.textContent = '✓ Đã chép, dán vào cột A của UP WEB 2026'; } catch (e) { b2.textContent = 'Không chép được, bôi đen bảng để chép'; } setTimeout(() => { b2.textContent = '📋 Chép bảng UP WEB'; }, 3000); };
      act.append(b1, b2); d.append(act);
      // bảng đối chiếu kiểu UP WEB
      const tb = el('table'), th = el('tr');
      ['NPP', 'Số PO KM', 'Số HĐ KM', 'SL KM', 'Số PO bán', 'Số HĐ bán', 'SL bán', 'SL UP', 'HĐ UP'].forEach(x => th.append(el('th', '', x)));
      tb.append(th);
      n.orders.forEach(o => { const tr = el('tr'); [o.npp, o.kmSo, o.kmHd || '0', o.slKm, o.so, o.hd, o.slBan, o.slBan + o.slKm, `${o.kmHd || '0'}, ${o.hd}`].forEach((v, j) => tr.append(el('td', [3, 6, 7].includes(j) ? 'n' : '', typeof v === 'number' ? v.toLocaleString('vi-VN') : v))); tb.append(tr); });
      const sc = el('div', 'tbl-scroll rp-tbl'); sc.append(el('h3', '', 'Đối chiếu từng đơn (kiểu UP WEB)'), tb); d.append(sc);
      // xem trước file up
      const tb2 = el('table'), th2 = el('tr');
      ['#', 'Số HĐ', 'Số đơn', 'Mã vật tư', 'Tên vật tư', 'LOT', 'Hết hạn', 'SL', 'Giá bán', 'Tiền CK', 'VAT', 'Tổng tiền tt'].forEach(x => th2.append(el('th', '', x)));
      tb2.append(th2);
      n.rows.forEach((x, j) => { const tr = el('tr', x.bad ? 'err' : ''); [j + 1, x.o.hd, x.so, x.o.mvt, x.o.ten, x.o.lot, x.o.nhh, x.sl, vnd(x.gia), vnd(x.ck), vnd(x.vat), vnd(x.tt)].forEach((v, c) => tr.append(el('td', c >= 7 ? 'n' : '', typeof v === 'number' ? v.toLocaleString('vi-VN') : v))); tb2.append(tr); });
      const sc2 = el('details', 'rp-more'); sc2.append(el('summary', '', `Xem trước ${n.rows.length} dòng file up`));
      const w2 = el('div', 'tbl-scroll rp-tbl'); w2.append(tb2); sc2.append(w2); d.append(sc2);
      list.append(d);
    });
    $('dmAll').hidden = NS.length < 2;
    $('dmAll').textContent = `⬇ Tải tất cả (${NS.length} file, .zip)`;
  }
  $('dmKh').addEventListener('change', () => R && show());
  // dán vào UP WEB 2026: NPP DD | SỐ PO KM | Số HĐ KM | SL KM | SỐ PO BÁN | SỐ HD BÁN | GIÁ TRỊ | DEALINE | GIỜ | SL BÁN | NOTE | dò số HĐ | | SL UP | HĐ UP
  const upweb = n => n.orders.map(o => [o.npp, o.kmSo, o.kmHd || '0', o.slKm, o.so, o.hd, '', '', '', o.slBan, '', o.hd, '', o.slBan + o.slKm, `${o.kmHd || '0'}, ${o.hd}`].join('\t')).join('\n');

  $('dmGo').onclick = async () => {
    if (!files.length) return;
    const btn = $('dmGo'); btn.disabled = true; $('dmMsg').textContent = 'Đang đọc file…';
    try {
      let lines = [], skip = {}, noHd = 0;
      for (const f of files) { const L = readFile(new Uint8Array(await f.arrayBuffer()), f.name); lines = lines.concat(L); noHd += L.noHd; for (const t in L.skip) skip[t] = (skip[t] || 0) + L.skip[t]; }
      if (!lines.length) throw new Error('Không có dòng "Xuất bán KH" / "Xuất khuyến mãi" nào đã xuất hóa đơn trong file.');
      R = convert(lines);
      const sk = Object.entries(skip).map(([t, n]) => `${t} ${n}`).concat(noHd ? [`chưa xuất HĐ ${noHd}`] : []);
      R.info = `${lines.length.toLocaleString('vi-VN')} dòng bán + KM` + (sk.length ? `; bỏ: ${sk.join(', ')}` : '');
      fillKh();
      $('dmOut').hidden = false; show();
      const NS = shown(), nErr = R.warn.filter(w => w.lv === 'err' && NS.some(n => n.kh === w.kh)).length;
      $('dmMsg').textContent = `Xong: ${NS.length} khách, ${NS.reduce((s, n) => s + n.rows.length, 0)} dòng file up.` + (nErr ? ` Có ${nErr} lỗi cần sửa (ô đỏ ở trên).` : '');
      logDm({ event: 'Chuyển đổi file up DMS', orders: NS.reduce((s, n) => s + n.orders.length, 0), ok: nErr ? 0 : 1, err: nErr, names: files.map(f => f.name).join(', ').slice(0, 200) });
    } catch (e) { $('dmMsg').textContent = '❌ ' + (e.message || e); $('dmOut').hidden = true; $('dmAll').hidden = true; R = null; logDm({ event: 'Chuyển đổi file up DMS', ok: 0, err: 1, names: String(e.message || e).slice(0, 150) }); }
    finally { btn.disabled = false; }
  };
  $('dmAll').onclick = async () => {
    if (!R) return;
    const zip = new JSZip(); shown().forEach(n => zip.file(fileName(n), toXls(n)));
    const t = new Date();
    save(await zip.generateAsync({ type: 'blob' }), `File up DMS ${p2(t.getDate())}.${p2(t.getMonth() + 1)}.zip`);
    logDm({ event: 'Tải file up DMS (zip)', orders: R.npps.length });
  };
  // ---- danh sách NPP up DMS + thêm tay ----
  const BASE = [['DD00008', 'NPP Hoàng Huy Phát'], ['TUANKIET001', 'NPP Minh Tuấn Kiệt'], ['DD00003', 'NPP An Phát'], ['KINGBAC0001', 'NPP Kinh Bắc'], ['', 'NPP Quỳnh Phát'], ['DD00010', 'NPP Ngọc An Khang'], ['DD00011', 'NPP Thảo Hương Trà']];
  function nppList() {
    const extra = added().filter(x => !DMS_KH.includes(normKh(x.code)));
    $('dmNppCount').textContent = `(${BASE.length + extra.length} NPP)`;
    $('dmNppList').textContent = 'Đang xuất: ' + BASE.map(([c, t]) => t + (c ? ` (${c})` : '')).concat(extra.map(x => `${x.name} (${normKh(x.code)}, team thêm)`)).join(', ') + '.';
  }
  try { if (!window.FFM_CUST_SHARED) { const c = JSON.parse(localStorage.getItem('ffm-data-v1:custShared') || 'null'); if (c) window.FFM_CUST_SHARED = c; } } catch (e) {}
  nppList();
  if (/^https:/.test(window.FFM_API || '')) fetch(window.FFM_API + '?t=aeonmap', { cache: 'no-store' }).then(r => r.json()).then(j => { if (j && j.cust) { window.FFM_CUST_SHARED = j.cust; nppList(); if (R) { fillKh(); show(); } } }).catch(() => {});
  $('dmAddBtn').onclick = async () => {
    const code = normKh($('dmAddCode').value), name = $('dmAddName').value.trim(), who = (($('who') || {}).value || '').trim(), note = $('dmAddNote');
    if (!who) { note.textContent = 'Nhập tên của bạn (góc trên bên phải) trước khi lưu.'; window.FFM_askName && window.FFM_askName(); return; }
    if (!/^[A-Z0-9_.-]{3,}$/.test(code)) { note.textContent = 'Nhập Mã KH như trong file bảng kê (vd C100061 hoặc C100061HP).'; return; }
    if (name.length < 3) { note.textContent = 'Nhập tên NPP.'; return; }
    if (DMS_KH.includes(code) || added().some(x => normKh(x.code) === code)) { note.textContent = `${code} đã có trong danh sách.`; return; }
    const btn = $('dmAddBtn'); btn.disabled = true; btn.textContent = 'Đang lưu…';
    window.FFM_CUST_SHARED = (window.FFM_CUST_SHARED || []).concat([{ acct: 'DMS', code, name }]);
    try { localStorage.setItem('ffm-data-v1:custShared', JSON.stringify(window.FFM_CUST_SHARED)); } catch (e) {}
    let ok = true;
    try { await fetch(window.FFM_API, { method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ team: 'khac', user: who, event: 'Thêm mã KH mới', acct: 'DMS', code, store: name, ver: (($('ver') || {}).textContent || '').replace('Phiên bản ', '') + ' beta' }) }); } catch (e) { ok = false; }
    btn.disabled = false; btn.textContent = 'Lưu vào bảng chung';
    $('dmAddCode').value = ''; $('dmAddName').value = '';
    note.textContent = ok ? `Đã thêm ${name} (${code}). Từ giờ file up DMS xuất cả NPP này.` : 'Chưa lưu lên bảng chung được (mất mạng?). Máy này vẫn nhớ, thử lại sau.';
    nppList(); if (R) { fillKh(); show(); }
  };
  window.FFM_dms = { readFile, convert, toAoa }; // để kiểm tra
})();
