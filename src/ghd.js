/* BẢN THỬ (beta) · Team khác › Mã giao thiếu theo ACC.
   Đầu vào: rawdata_giao_hang_du_<từ>_<đến>.xlsx (raw báo cáo giao hàng đủ).
   Dòng thiếu = "Chênh lệch" < 0 (SL xuất < SL trên đơn), bỏ đơn "PO hủy".
   SL thiếu = -Chênh lệch; Giá trị thiếu = -"Thành tiền chênh lệch"; Miền theo Kho xuất hàng (bảng Quy đổi như báo cáo List order).
   Gom theo ACC → mã (barcode) → lý do. File chỉ xử lý trong trình duyệt. */
(function () {
  const $ = id => document.getElementById(id);
  if (!$('ghDrop')) return;
  const fold = s => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase().replace(/\s+/g, ' ').trim();
  const num = v => typeof v === 'number' ? v : (Number(String(v || '').replace(/[^\d.-]/g, '')) || 0);
  const vnd = v => Math.round(v).toLocaleString('vi-VN');
  const fmtTr = v => { const tr = v / 1e6; return tr >= 1000 ? (Math.round(tr / 100) / 10).toString().replace('.', ',') + 'tỷ' : (tr >= 10 ? Math.round(tr) : Math.round(tr * 10) / 10).toString().replace('.', ',') + 'tr'; };
  const kvOf = k => { const KV = window.FFM_KV || {}, n = String(k || '').trim(); if (KV[n]) return KV[n]; const f = fold(n); for (const x in KV) if (fold(x) === f) return KV[x]; return 'Khác'; };
  const el = (t, c, x) => { const e = document.createElement(t); if (c) e.className = c; if (x != null) e.textContent = x; return e; };
  let file = null, G = null; // G: { rows, from, to }

  const drop = $('ghDrop'), pick = $('ghFile');
  const setFile = f => { if (!f) return; file = f; $('ghName').textContent = f.name; $('ghGo').disabled = false; $('ghMsg').textContent = ''; };
  drop.onclick = () => pick.click();
  drop.onkeydown = e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); pick.click(); } };
  pick.onchange = () => { setFile(pick.files[0]); pick.value = ''; };
  drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('over'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('over'));
  drop.addEventListener('drop', e => { e.preventDefault(); drop.classList.remove('over'); setFile(e.dataTransfer.files[0]); });

  function read(buf, name) {
    const wb = XLSX.read(buf, { type: 'array', dense: true });
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: '' });
    const h = rows.findIndex(r => r.some(x => fold(x) === 'chenh lech') && r.some(x => fold(x) === 'barcode'));
    if (h < 0) throw new Error('Không thấy cột "BarCode" và "Chênh lệch". File này có đúng là raw báo cáo giao hàng đủ không?');
    const H = rows[h].map(fold), col = re => H.findIndex(x => re.test(x));
    const C = { acc: col(/^account$/), so: col(/^ma phieu gt$/), po: col(/^so po$/), kh: col(/^ma kh$/), tkh: col(/^ten kh$/), brand: col(/^brand$/), mvt: col(/^ma vat tu$/), bc: col(/^barcode$/), ten: col(/^ten vat tu$/), dv: col(/^don vi$/), dat: col(/^so luong tren don hang$/), xuat: col(/^so luong xuat$/), cl: col(/^chenh lech$/), tien: col(/^thanh tien chenh lech$/), kho: col(/^kho xuat hang$/), why: col(/^ly do giao hang khong du$/), st: col(/^tinh trang don hang$/), ngay: col(/^ngay don/) };
    const miss = ['acc', 'bc', 'dat', 'xuat', 'cl'].filter(k => C[k] < 0);
    if (miss.length) throw new Error('File thiếu cột: ' + miss.join(', '));
    const out = [];
    let huy = 0;
    rows.slice(h + 1).forEach(r => {
      const cl = num(r[C.cl]); if (!(cl < 0)) return;
      if (C.st >= 0 && fold(r[C.st]) === 'po huy') { huy++; return; }
      out.push({ acc: String(r[C.acc] || '(trống)').trim(), reg: kvOf(r[C.kho]), kho: String(r[C.kho] || '').trim(), bc: String(r[C.bc] || '').trim(), mvt: C.mvt >= 0 ? String(r[C.mvt] || '').trim() : '', ten: C.ten >= 0 ? String(r[C.ten] || '').trim() : '', brand: C.brand >= 0 ? String(r[C.brand] || '').trim() : '', dv: C.dv >= 0 ? String(r[C.dv] || '') : '', dat: num(r[C.dat]), xuat: num(r[C.xuat]), thieu: -cl, tien: C.tien >= 0 ? -num(r[C.tien]) : 0, why: String(r[C.why] || '').trim() || '(chưa ghi lý do)', so: C.so >= 0 ? String(r[C.so] || '') : '', kh: C.kh >= 0 ? String(r[C.kh] || '') : '', ngay: C.ngay >= 0 ? String(r[C.ngay] || '') : '' });
    });
    const m = String(name || '').match(/(\d{4})-(\d{2})-(\d{2})_(\d{4})-(\d{2})-(\d{2})/);
    return { rows: out, huy, total: rows.length - h - 1, from: m ? `${m[3]}/${m[2]}/${m[1]}` : '', to: m ? `${m[6]}/${m[5]}/${m[4]}` : '' };
  }
  // gom: ACC -> mã (barcode + mã vật tư) -> lý do
  function group(rows) {
    const A = new Map();
    rows.forEach(o => {
      let a = A.get(o.acc); if (!a) A.set(o.acc, a = { acc: o.acc, thieu: 0, tien: 0, dong: 0, don: new Set(), items: new Map(), why: {} });
      a.thieu += o.thieu; a.tien += o.tien; a.dong++; a.don.add(o.so); a.why[o.why] = (a.why[o.why] || 0) + o.tien;
      const k = o.mvt || o.bc;
      let it = a.items.get(k); if (!it) a.items.set(k, it = { bc: o.bc, mvt: o.mvt, ten: o.ten, brand: o.brand, dv: o.dv, dat: 0, xuat: 0, thieu: 0, tien: 0, don: new Set(), regs: new Set(), why: new Map() });
      it.dat += o.dat; it.xuat += o.xuat; it.thieu += o.thieu; it.tien += o.tien; it.don.add(o.so); it.regs.add(o.reg);
      let w = it.why.get(o.why); if (!w) it.why.set(o.why, w = { why: o.why, dat: 0, xuat: 0, thieu: 0, tien: 0, don: new Set() });
      w.dat += o.dat; w.xuat += o.xuat; w.thieu += o.thieu; w.tien += o.tien; w.don.add(o.so);
    });
    return [...A.values()].sort((x, y) => y.tien - x.tien).map(a => Object.assign(a, { items: [...a.items.values()].sort((x, y) => y.tien - x.tien).map(it => Object.assign(it, { why: [...it.why.values()].sort((x, y) => y.tien - x.tien) })) }));
  }
  const topWhy = o => Object.entries(o).sort((x, y) => y[1] - x[1]).slice(0, 2).map(([k]) => k).join(', ');

  // ---- bộ lọc + danh sách ----
  function fillSelect(sel, values, all) { const cur = sel.value; sel.innerHTML = ''; [[all, '']].concat(values.map(v => [v, v])).forEach(([t, v]) => { const o = el('option', '', t); o.value = v; sel.append(o); }); if (values.includes(cur)) sel.value = cur; }
  function filtered() {
    const acc = $('ghAcc').value, reg = $('ghReg').value, why = $('ghWhy').value, q = fold($('ghQ').value);
    return G.rows.filter(o => (!acc || o.acc === acc) && (!reg || o.reg === reg) && (!why || o.why === why) && (!q || fold(o.bc + ' ' + o.mvt + ' ' + o.ten).includes(q)));
  }
  function render() {
    const rows = filtered(), groups = group(rows);
    const kpi = $('ghKpi'); kpi.innerHTML = '';
    const k = (lab, v, sub) => { const d = el('div', 'tile'); d.append(el('small', '', lab), el('b', '', v)); if (sub) d.append(el('small', 'note', sub)); kpi.append(d); };
    const tien = rows.reduce((s, o) => s + o.tien, 0), sl = rows.reduce((s, o) => s + o.thieu, 0);
    k('Giá trị giao thiếu', fmtTr(tien), vnd(tien) + ' đ');
    k('SL thiếu', sl.toLocaleString('vi-VN'), rows.length.toLocaleString('vi-VN') + ' dòng');
    k('Số ACC', String(groups.length), 'số mã: ' + new Set(rows.map(o => o.mvt || o.bc)).size);
    k('Số đơn bị thiếu', new Set(rows.map(o => o.so)).size.toLocaleString('vi-VN'));
    const list = $('ghList'); list.innerHTML = '';
    if (!groups.length) { list.append(el('p', 'note', 'Không có mã giao thiếu nào theo bộ lọc này.')); return; }
    groups.forEach((a, i) => {
      const d = el('details', 'gh-acc'); if (groups.length === 1 || i === 0) d.open = true;
      const s = el('summary');
      s.append(el('b', '', a.acc), el('span', 'note', ` · ${a.items.length} mã · ${a.don.size} đơn · thiếu ${a.thieu.toLocaleString('vi-VN')} · `), el('b', 'gh-v', fmtTr(a.tien)), el('span', 'note', ' · ' + topWhy(a.why)));
      d.append(s);
      const tb = el('table'), th = el('tr');
      ['Barcode', 'Tên sản phẩm', 'Số đơn', 'SL đặt', 'SL xuất', 'SL thiếu', 'Giá trị thiếu', 'Lý do'].forEach(x => th.append(el('th', '', x)));
      tb.append(th);
      a.items.forEach(it => {
        it.why.forEach((w, j) => {
          const tr = el('tr', j ? 'sub' : '');
          if (!j) { const td = el('td', 'bc', it.bc); if (it.mvt && it.mvt !== it.bc) td.append(el('small', 'note', ' ' + it.mvt.replace(it.bc, ''))); tr.append(td, el('td', '', it.ten)); }
          else tr.append(el('td'), el('td'));
          tr.append(el('td', 'n', String(w.don.size)), el('td', 'n', w.dat.toLocaleString('vi-VN')), el('td', 'n', w.xuat.toLocaleString('vi-VN')), el('td', 'n', w.thieu.toLocaleString('vi-VN')), el('td', 'n', vnd(w.tien)), el('td', 'why', w.why));
          tb.append(tr);
        });
      });
      const sc = el('div', 'tbl-scroll rp-tbl'); sc.append(tb); d.append(sc); list.append(d);
    });
  }
  ['ghAcc', 'ghReg', 'ghWhy'].forEach(id => $(id).addEventListener('change', render));
  let qt; $('ghQ').addEventListener('input', () => { clearTimeout(qt); qt = setTimeout(render, 200); });

  // ---- sheet Excel: theo ACC → mã → lý do, có dòng tổng ACC ----
  async function addSheet(wb, name) {
    if (!G) return;
    const ws = wb.addWorksheet(name || 'Giao thiếu theo ACC', { views: [{ state: 'frozen', ySplit: 3 }] });
    ws.getCell('A1').value = `MÃ GIAO THIẾU THEO ACC${G.from ? ` (từ ${G.from} đến ${G.to})` : ''}`; ws.getCell('A1').font = { bold: true, size: 13 };
    const head = ['ACC', 'Miền', 'Barcode', 'Mã vật tư', 'Tên sản phẩm', 'Brand', 'ĐVT', 'Số đơn', 'SL đặt', 'SL xuất', 'SL thiếu', 'Giá trị thiếu', 'Lý do giao không đủ'];
    ws.getRow(3).values = head;
    ws.getRow(3).eachCell(c => { c.font = { bold: true }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFCC99' } }; c.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }; c.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } }; });
    ws.getRow(3).height = 30;
    let r = 4;
    const groups = group(G.rows);
    groups.forEach(a => {
      const tr = ws.getRow(r++);
      tr.values = [a.acc, '', '', '', `Tổng ${a.acc}: ${a.items.length} mã`, '', '', a.don.size, '', '', a.thieu, a.tien, topWhy(a.why)];
      tr.eachCell(c => { c.font = { bold: true }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCFFCC' } }; });
      a.items.forEach(it => it.why.forEach(w => {
        ws.getRow(r++).values = [a.acc, [...it.regs].join(', '), it.bc, it.mvt, it.ten, it.brand, it.dv, w.don.size, w.dat, w.xuat, w.thieu, w.tien, w.why];
      }));
    });
    const tot = ws.getRow(r);
    tot.values = ['TỔNG CỘNG', '', '', '', '', '', '', new Set(G.rows.map(o => o.so)).size, '', '', G.rows.reduce((s, o) => s + o.thieu, 0), G.rows.reduce((s, o) => s + o.tien, 0), ''];
    tot.eachCell(c => { c.font = { bold: true }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF99' } }; });
    [12, 12, 16, 18, 46, 14, 8, 9, 10, 10, 10, 16, 36].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
    [8, 9, 10, 11, 12].forEach(c => { ws.getColumn(c).numFmt = '#,##0'; });
    ws.autoFilter = { from: { row: 3, column: 1 }, to: { row: r - 1, column: head.length } };
  }
  window.FFM_extraSheets = wb => addSheet(wb, 'Giao thiếu theo ACC');

  $('ghGo').onclick = async () => {
    if (!file) return;
    const btn = $('ghGo'); btn.disabled = true; $('ghMsg').textContent = 'Đang đọc file…';
    try {
      G = read(new Uint8Array(await file.arrayBuffer()), file.name);
      fillSelect($('ghAcc'), group(G.rows).map(a => a.acc), 'Tất cả ACC');
      fillSelect($('ghReg'), ['Miền Nam', 'Miền Bắc', 'Miền Trung', 'Khác'].filter(x => G.rows.some(o => o.reg === x)), 'Tất cả miền');
      const whys = {}; G.rows.forEach(o => { whys[o.why] = (whys[o.why] || 0) + o.tien; });
      fillSelect($('ghWhy'), Object.keys(whys).sort((x, y) => whys[y] - whys[x]), 'Tất cả lý do');
      $('ghOut').hidden = false; render();
      $('ghMsg').textContent = `${G.rows.length.toLocaleString('vi-VN')} dòng giao thiếu / ${G.total.toLocaleString('vi-VN')} dòng${G.from ? ` (từ ${G.from} đến ${G.to})` : ''}${G.huy ? `, bỏ ${G.huy} dòng PO hủy` : ''}. Báo cáo List order tải về sẽ có thêm sheet này.`;
      $('ghDl').hidden = false;
    } catch (e) { $('ghMsg').textContent = '❌ ' + (e.message || e); $('ghOut').hidden = true; $('ghDl').hidden = true; G = null; }
    finally { btn.disabled = false; }
  };
  $('ghDl').onclick = async () => {
    if (!G) return;
    $('ghDl').disabled = true;
    try {
      await window.FFM_loadExcelJS();
      const wb = new ExcelJS.Workbook(); await addSheet(wb, 'Giao thiếu theo ACC');
      const buf = await wb.xlsx.writeBuffer();
      const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([buf])); a.download = `Giao thieu theo ACC ${(G.to || '').replace(/\//g, '.')}.xlsx`.replace(' .xlsx', '.xlsx'); document.body.append(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
    } catch (e) { $('ghMsg').textContent = '❌ ' + (e.message || e); }
    finally { $('ghDl').disabled = false; }
  };
})();
