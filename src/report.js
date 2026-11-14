/* Team khác › Báo cáo List order hằng ngày.
   Đầu vào: file "Danh sách đơn đặt hàng (xuất bán)" xuất từ phần mềm (data dd.mm.xlsx).
   Đầu ra: "Total PO update dd.mm.yy.xlsx" gồm sheet Total (theo miền), National (đơn không hủy), Đơn hủy.
   Quy tắc (đối chiếu file Total PO update 13.11.26 của anh Tài):
   - Đã giao = có "Ngày bàn giao đơn hàng cho DVVC" và Smartlog không phải "Giao hàng thất bại"
   - Đã xuất HĐ = có "Số hóa đơn"; PO hủy = Tình trạng đơn hàng "PO hủy"
   - Khu vực theo Kho xuất (bảng Quy đổi); ACC + Kênh theo DMKH (bỏ đuôi HP ở mã khách hàng)
   File chỉ đọc trong trình duyệt, không gửi lên đâu. */
(function () {
  const $ = id => document.getElementById(id);
  if (!$('rpDrop')) return;
  const KV = {
    'Kho Bùi Thị Xuân': 'Miền Nam', 'Kho Bình Tân': 'Miền Nam', 'Kho Đà Nẵng': 'Miền Trung', 'Kho Hà Nội': 'Miền Bắc', 'Kho Online PVD': 'Miền Nam',
    'Kho Phạm Văn Đồng': 'Miền Nam', 'Kho Thuận An': 'Miền Nam', 'Kho Thuê ngoài - Đông Tân': 'Miền Nam', 'Kho Online Hà Nội': 'Miền Bắc', 'Kho Bình Hòa': 'Miền Nam',
    'Kho LogisValley - Long Biên': 'Miền Bắc', 'Kho Online Đà Nẵng': 'Miền Trung', 'Kho Online Nơ Trang Long': 'Miền Nam', 'Kho Upbase Nơ Trang Long': 'Miền Nam',
    'Kho Online Upbase Hà Nội': 'Miền Bắc', 'Kho Hateco Long Biên': 'Miền Bắc', 'Kho Gamuda Hà Nội': 'Miền Bắc', 'Kho thuê ngoài EIMSKIP Quận 9': 'Miền Nam',
    'Kho thuê dịch vụ Daso': 'Miền Nam', 'Kho Tam Bình': 'Miền Nam', 'Kho Bình Hòa - cũ': 'Miền Nam', 'Kho Woojin Bắc Ninh': 'Miền Bắc', 'Kho Lạnh ABA Tân Tạo': 'Miền Nam',
    'Kho SEC Hà Nội': 'Miền Bắc', 'Kho DC Đồng An': 'Miền Nam'
  };
  const REGIONS = ['Miền Nam', 'Miền Bắc', 'Miền Trung'];
  const fold = s => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase().replace(/\s+/g, ' ').trim();
  // khai báo của team (sheet "Khai báo báo cáo"): acc {MÃ: [ACC, Kênh]}, kho {tên kho: Miền}
  let SHARED = { acc: {}, kho: {} };
  try { SHARED = Object.assign(SHARED, JSON.parse(localStorage.getItem('ffm-rpmap') || '{}')); } catch (e) {}
  async function loadShared() {
    const API = window.FFM_API; if (!/^https:/.test(API || '')) return;
    try {
      const j = await Promise.race([fetch(API + '?t=rpmap', { cache: 'no-store' }).then(r => r.json()), new Promise((_, no) => setTimeout(() => no(new Error('timeout')), 8000))]);
      if (j && j.acc) { SHARED = { acc: j.acc || {}, kho: j.kho || {} }; try { localStorage.setItem('ffm-rpmap', JSON.stringify(SHARED)); } catch (e) {} }
    } catch (e) {}
  }
  const kvOf = k => { const n = String(k || '').trim(); if (KV[n]) return KV[n]; if (SHARED.kho[n]) return SHARED.kho[n]; const f = fold(n); for (const key in KV) if (fold(key) === f) return KV[key]; return ''; };
  const num = v => typeof v === 'number' ? v : (Number(String(v || '').replace(/[^\d.-]/g, '')) || 0);
  const fmtTr = v => { const tr = v / 1e6; return tr >= 1000 ? (Math.round(tr / 100) / 10).toString().replace('.', ',') + 'tỷ' : Math.round(tr) + 'tr'; };
  const vnd = v => Math.round(v).toLocaleString('vi-VN');
  let file = null, out = null;

  // ---- chọn file ----
  const drop = $('rpDrop'), pick = $('rpFile');
  const setFile = f => { if (!f) return; file = f; $('rpName').textContent = f.name; $('rpGo').disabled = false; $('rpMsg').textContent = ''; };
  drop.onclick = () => pick.click();
  drop.onkeydown = e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); pick.click(); } };
  pick.onchange = () => { setFile(pick.files[0]); pick.value = ''; };
  drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('over'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('over'));
  drop.addEventListener('drop', e => { e.preventDefault(); drop.classList.remove('over'); setFile(e.dataTransfer.files[0]); });

  // ---- đọc file data ----
  function readData(buf) {
    const wb = XLSX.read(buf, { type: 'array', cellDates: false, dense: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
    const h = rows.findIndex(r => r.some(x => /ma phieu xuat/.test(fold(x))) && r.some(x => /tinh trang don hang/.test(fold(x))));
    if (h < 0) throw new Error('Không thấy dòng tiêu đề (Mã phiếu xuất, Tình trạng đơn hàng…). File này có đúng là "Danh sách đơn đặt hàng (xuất bán)" không?');
    const H1 = rows[h], H2 = rows[h + 1] || [];
    const width = Math.max(H1.length, H2.length);
    const head = []; for (let c = 0; c < width; c++) head.push(String(H2[c] || H1[c] || '').replace(/\r/g, '').trim());
    const col = re => head.findIndex(x => re.test(fold(x)));
    const C = {
      so: col(/^ma phieu xuat/), cname: col(/^ten khach hang/), ddh: col(/^ngay ve don theo ddh/), code: col(/^ma khach hang/), val: col(/^tri gia don hang/), dvvc: col(/^ngay ban giao don hang cho dvvc/),
      inv: col(/^so hoa don/), amt: col(/^amount/), st: col(/^tinh trang don hang$/), sl: col(/smartlog/), kho: col(/^kho xuat/), stt: col(/^so tt/), de: col(/^ma phieu de xuat \(code\)/)
    };
    const miss = Object.entries(C).filter(([k, v]) => v < 0 && !['stt', 'de', 'ddh'].includes(k)).map(([k]) => k);
    if (miss.length) throw new Error('File thiếu cột: ' + miss.join(', '));
    const range = rows.slice(0, h).map(r => r.join(' ')).join(' ').match(/(\d{1,2}\/\d{1,2}\/\d{4})\s*đến ngày:?\s*(\d{1,2}\/\d{1,2}\/\d{4})/i);
    const data = rows.slice(h + 2).filter(r => String(r[C.de >= 0 ? C.de : C.so] || '').trim() || num(r[C.val]));
    return { head, H1: H1.map(x => String(x || '').replace(/\r/g, '')), H2: H2.map(x => String(x || '').replace(/\r/g, '')), C, data, from: range ? range[1] : '', to: range ? range[2] : '' };
  }

  function build(D, extra) {
    const A = window.FFM_ACCMAP || {};
    const S = SHARED.acc || {};
    const accOf = code => { const c = String(code || '').trim().toUpperCase(), b = c.replace(/HP$/, ''); return A[b] || A[c] || S[b] || S[c] || null; };
    const C = D.C, nat = [], huy = [], noAcc = new Map(), noKv = {};
    D.data.forEach(r => {
      const acc = accOf(r[C.code]), kv = kvOf(r[C.kho]);
      if (!acc) { const k = String(r[C.code]).trim().toUpperCase().replace(/HP$/, ''); const x = noAcc.get(k) || { name: C.cname >= 0 ? String(r[C.cname] || '').trim() : '', n: 0, v: 0 }; x.n++; x.v += num(r[C.val]); noAcc.set(k, x); }
      const daGiao = String(r[C.dvvc]).trim() && fold(r[C.sl]) !== 'giao hang that bai' ? 1 : 0;
      const daHD = String(r[C.inv]).trim() ? 1 : 0, isHuy = fold(r[C.st]) === 'po huy' ? 1 : 0;
      const o = { acc: acc ? acc[0] : '', kenh: acc ? acc[1] : '', kv, daGiao, daHD, isHuy, val: num(r[C.val]), amt: num(r[C.amt]), raw: r };
      if (!kv && !isHuy) { const k = String(r[C.kho] || '(trống)').trim(); noKv[k] = noKv[k] || { n: 0, v: 0 }; noKv[k].n++; noKv[k].v += o.val; }
      (isHuy ? huy : nat).push(o);
    });
    // bảng tổng theo miền
    const CATS = [
      { k: 'gx', label: 'Đơn đã giao hàng và đã xuất HD', f: o => o.daGiao && o.daHD },
      { k: 'cx', label: 'Đơn chưa giao hàng và đã xuất HD', f: o => !o.daGiao && o.daHD, note: 'đơn đang giao' },
      { k: 'gc', label: 'Đơn hàng đã giao nhưng chưa xuất HD', f: o => o.daGiao && !o.daHD },
      { k: 'cc', label: 'Đơn chưa giao và chưa xuất HĐ', f: o => !o.daGiao && !o.daHD, note: 'đơn đang nhặt' }
    ];
    const table = reg => {
      const list = nat.filter(o => reg ? o.kv === reg : o.kv);
      const rows = CATS.map(c => {
        const L = list.filter(c.f), v = L.reduce((s, o) => s + o.val, 0), a = L.reduce((s, o) => s + o.amt, 0);
        let note = '';
        if (c.note && reg) {
          const g = {}; L.forEach(o => { const k = o.acc || String(o.raw[C.code]).trim(); g[k] = (g[k] || 0) + o.val; });
          note = Object.entries(g).filter(([, x]) => x >= 50e6).sort((x, y) => y[1] - x[1]).slice(0, 6).map(([k, x]) => `${k} ${fmtTr(x)} - ${c.note}`).join('\n');
        }
        return { label: c.label, k: c.k, n: L.length, v, a, thieu: c.k === 'gx' || c.k === 'cx' ? v - a : 0, note };
      });
      const hl = huy.filter(o => reg ? o.kv === reg : true);
      const tot = rows.reduce((s, x) => ({ n: s.n + x.n, v: s.v + x.v, a: s.a + x.a, thieu: s.thieu + x.thieu }), { n: 0, v: 0, a: 0, thieu: 0 });
      return { rows, huy: { n: hl.length, v: hl.reduce((s, o) => s + o.val, 0) }, tot };
    };
    const T = { nw: table(''), reg: REGIONS.map(r => ({ r, t: table(r) })) };
    const m = (D.to || '').split('/'), mm = m.length === 3 ? `${m[1]}/${m[2]}` : '', prev = m.length === 3 ? (+m[1] === 1 ? `12/${+m[2] - 1}` : `${String(+m[1] - 1).padStart(2, '0')}/${m[2]}`) : '';
    const hdThang = T.nw.tot.a;
    return { nat, huy, T, noAcc: [...noAcc.entries()], noKv, mm, prev, hdThang, extra, D };
  }

  // ---- ghi file Excel: sheet Total lấy nguyên mẫu (định dạng + công thức) từ file Total PO update của anh Tài ----
  const EXCELJS = 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js';
  let exP = null;
  const loadExcelJS = () => window.ExcelJS ? Promise.resolve() : (exP = exP || new Promise((ok, no) => { const s = document.createElement('script'); s.src = (window.FFM_LIBS && window.FFM_LIBS.exceljs) || EXCELJS; s.onload = ok; s.onerror = () => { exP = null; no(new Error('Không tải được thư viện Excel (cần internet).')); }; document.head.append(s); }));
  const REG_ROW = { 'Miền Nam': 18, 'Miền Bắc': 28, 'Miền Trung': 38 };
  // tính lại kết quả các công thức đơn giản (+, -, SUM) để file mở bằng điện thoại/Zalo vẫn thấy số
  function evalFormulas(ws, known) {
    const memo = Object.assign({}, known), busy = {};
    const val = a => {
      if (a in memo) return memo[a];
      if (busy[a]) return 0; busy[a] = 1;
      const c = ws.getCell(a), f = c.formula;
      let v = 0;
      if (f) {
        const expr = f.replace(/SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)/g, (_, c1, r1, c2, r2) => { let s = 0; for (let i = +r1; i <= +r2; i++) s += val(c1 + i); return '(' + s + ')'; })
          .replace(/\b([A-Z]{1,2})(\d+)\b/g, (_, cc, rr) => '(' + val(cc + rr) + ')');
        v = /^[\d\s.()+\-e]+$/.test(expr) ? Function('return ' + expr)() : 0;
      } else v = typeof c.value === 'number' ? c.value : 0;
      memo[a] = v; return v;
    };
    return val;
  }
  async function toXlsx(R) {
    await loadExcelJS();
    const D = R.D, wb = new ExcelJS.Workbook();
    await wb.xlsx.load(Uint8Array.from(atob(TPL_TOTAL), c => c.charCodeAt(0)));
    const ws = wb.getWorksheet('Total') || wb.worksheets[0];
    // số liệu theo miền (đúng các ô COUNTIFS/SUMIFS của mẫu)
    const known = { X3: 0, H3: R.extra.prev || 0 };
    R.T.reg.forEach(({ r, t }) => {
      const b = REG_ROW[r];
      t.rows.forEach((x, i) => { known['G' + (b + 1 + i)] = x.n; known['H' + (b + 1 + i)] = x.v; known['I' + (b + 1 + i)] = x.a; });
      const h = r === 'Miền Trung' ? b + 5 : b + 7; known['G' + h] = 0; known['H' + h] = 0;
      const notes = { cx: b + 2, cc: b + 4 };
      t.rows.forEach(x => { if (notes[x.k]) ws.getCell('L' + notes[x.k]).value = x.note || null; });
    });
    const val = evalFormulas(ws, known);
    ws.eachRow({ includeEmpty: false }, row => row.eachCell({ includeEmpty: false }, cell => {
      // National không còn 10 dòng trống ở đầu: dữ liệu bắt đầu dòng 3 thay vì 13
      if (cell.formula) cell.value = { formula: cell.formula.replace(/(National!\$[A-Z]+\$)13:/g, (_, p) => p + '3:'), result: val(cell.address) };
    }));
    ws.getCell('H3').value = R.extra.prev || null;
    ws.getCell('H4').value = null;
    ws.getCell('G2').value = `Trị giá đơn hàng tháng ${R.mm} đã xuất hóa đơn`;
    ws.getCell('G3').value = `Trị giá đơn hàng tháng ${R.prev} đã xuất hóa đơn sang tháng ${R.mm}`;
    ws.getCell('G4').value = `Trị giá đơn hàng hủy - Đã xuất hóa đơn thay thế trong tháng ${R.mm}`;
    // National + ĐƠN HỦY: tiêu đề dòng 1-2, dữ liệu từ dòng 3 (cột giữ như file gốc: cột C khu vực, D đã giao, E đã xuất HĐ, F hủy, O trị giá, AA amount, AJ tình trạng)
    const keep = D.head.map((h, i) => i).filter(i => i !== D.C.ddh);
    const head1 = ['ACC', 'Kênh phân phối', 'Khu Vực', 'Đã giao', 'Đã xuất HĐ', 'PO HỦY'].concat(keep.map(i => D.H1[i] || ''));
    const head2 = ['', '', '', '', '', ''].concat(keep.map(i => D.H2[i] || ''));
    const sheet = (name, list) => {
      const s = wb.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 2, xSplit: 0 }] });
      s.getRow(1).values = head1; s.getRow(2).values = head2;
      [1, 2].forEach(r => { const row = s.getRow(r); row.height = 42; row.eachCell(c => { c.font = { bold: true, size: 10 }; c.alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } }; c.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } }; }); });
      list.forEach((o, i) => { s.getRow(3 + i).values = [o.acc, o.kenh, o.kv || '#N/A', o.daGiao, o.daHD, o.isHuy].concat(keep.map(k => o.raw[k] === undefined || o.raw[k] === null ? '' : o.raw[k])); });
      s.columns.forEach((c, i) => { c.width = i < 6 ? 11 : 16; });
      [15, 27].forEach(c => { s.getColumn(c).numFmt = '#,##0'; });
      s.autoFilter = { from: { row: 2, column: 1 }, to: { row: 2 + list.length, column: head1.length } };
      return s;
    };
    sheet('National', R.nat);
    sheet('ĐƠN HỦY ', R.huy);
    wb.calcProperties = Object.assign({}, wb.calcProperties, { fullCalcOnLoad: true });
    const buf = await wb.xlsx.writeBuffer();
    const [d, mo, y] = (R.D.to || '').split('/');
    const name = `Total PO update ${d ? `${d.padStart(2, '0')}.${mo.padStart(2, '0')}.${y.slice(2)}` : new Date().toISOString().slice(0, 10)}.xlsx`;
    return { blob: new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), name };
  }

  // ---- hiển thị ----
  function show(R) {
    const box = $('rpOut'); box.innerHTML = ''; box.hidden = false;
    const el = (t, c, x) => { const e = document.createElement(t); if (c) e.className = c; if (x != null) e.textContent = x; return e; };
    const sum = el('div', 'rp-kpi');
    const k = (lab, v, sub) => { const d = el('div', 'tile'); d.append(el('small', '', lab), el('b', '', v)); if (sub) d.append(el('small', 'note', sub)); sum.append(d); };
    k('Đơn (không tính hủy)', R.T.nw.tot.n.toLocaleString('vi-VN') + ' đơn', vnd(R.T.nw.tot.v) + ' đ');
    k('Đã xuất HĐ trong tháng', vnd(R.hdThang) + ' đ');
    k('Chưa giao', (R.T.nw.rows[1].n + R.T.nw.rows[3].n) + ' đơn', fmtTr(R.T.nw.rows[1].v + R.T.nw.rows[3].v));
    k('PO hủy', R.huy.length + ' đơn', fmtTr(R.huy.reduce((s, o) => s + o.val, 0)));
    box.append(sum);
    [['Nation Wide', R.T.nw]].concat(R.T.reg.map(x => [x.r, x.t])).forEach(([name, t]) => {
      const wrap = el('div', 'rp-tbl');
      wrap.append(el('h3', '', name));
      const tb = el('table'), th = el('tr');
      ['Tình trạng', 'Số PO', 'Trị giá đơn', 'Đã xuất HĐ', 'Thiếu'].concat(name === 'Nation Wide' ? [] : ['Ghi chú']).forEach(x => th.append(el('th', '', x)));
      tb.append(th);
      t.rows.concat([{ label: 'Tổng cộng', n: t.tot.n, v: t.tot.v, a: t.tot.a, thieu: t.tot.thieu, tot: 1 }]).forEach(x => {
        const tr = el('tr', x.tot ? 'tot' : '');
        tr.append(el('td', '', x.label), el('td', 'n', x.n.toLocaleString('vi-VN')), el('td', 'n', vnd(x.v)), el('td', 'n', vnd(x.a)), el('td', 'n', x.thieu ? vnd(x.thieu) : ''));
        if (name !== 'Nation Wide') tr.append(el('td', 'nt', x.note || ''));
        tb.append(tr);
      });
      const sc = el('div', 'tbl-scroll'); sc.append(tb); wrap.append(sc); box.append(wrap);
    });
    // mã KH thiếu ACC / kho thiếu miền: cho team khai báo + Ghi nhớ (dùng chung cả team)
    // kho chưa có miền: không hiện (anh Tài 14/11), vẫn không tính vào bảng theo miền như file gốc
    if (R.noAcc.length) {
      const w = el('div', 'rp-warn');
      const accList = [...new Set(Object.values(window.FFM_ACCMAP || {}).map(x => x[0]).concat(Object.values(SHARED.acc).map(x => x[0])).filter(Boolean))].sort();
      let dl = $('rpAccList'); if (!dl) { dl = el('datalist'); dl.id = 'rpAccList'; document.body.append(dl); }
      dl.innerHTML = ''; accList.forEach(a => { const o = el('option'); o.value = a; dl.append(o); });
      const KENH = ['MT', 'GT', 'B2B', 'HORECA', 'OTHERS'];
      if (R.noAcc.length) {
        w.append(el('p', 'rp-w-h', `⚠ ${R.noAcc.length} mã khách hàng chưa có trong DMKH (cột ACC đang trống). Nhập ACC + kênh rồi bấm Ghi nhớ:`));
        R.noAcc.forEach(([k, x]) => {
          const row = el('div', 'rp-fix');
          const inp = el('input'); inp.setAttribute('list', 'rpAccList'); inp.placeholder = 'ACC, vd COOP'; inp.setAttribute('aria-label', 'ACC của ' + k);
          const guess = accList.find(a => k.startsWith(a.replace(/[^A-Z0-9]/gi, '').toUpperCase())) || ''; inp.value = guess;
          const sel = el('select'); sel.setAttribute('aria-label', 'Kênh của ' + k);
          KENH.forEach(m => { const o = el('option', '', m); o.value = m; sel.append(o); });
          if (guess) { const any = Object.values(window.FFM_ACCMAP || {}).find(v => v[0] === guess); if (any && KENH.includes(any[1])) sel.value = any[1]; }
          const b = el('button', 'small primary', 'Ghi nhớ'); b.type = 'button';
          b.onclick = () => remember('acc', k, inp.value.trim().toUpperCase(), sel.value, b);
          row.append(el('span', 'rp-k', `${k}${x.name ? ' – ' + x.name : ''} (${x.n} đơn, ${fmtTr(x.v)})`), inp, sel, b); w.append(row);
        });
      }
      w.append(el('p', 'note', 'Ghi nhớ xong, báo cáo tự chạy lại. Lần sau cả team không phải khai báo lại.'));
      box.append(w);
    }
  }

  $('rpPrev').addEventListener('blur', e => { const n = Number(e.target.value.replace(/\D/g, '')); e.target.value = n ? n.toLocaleString('vi-VN') : ''; });
  // lịch sử: ghi vào sheet Log (cột Nhóm = khac) giống Team đơn hàng
  function logRp(d) {
    const API = window.FFM_API; if (!/^https:/.test(API || '')) return;
    try { fetch(API, { method: 'POST', mode: 'no-cors', keepalive: true, headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(Object.assign({ team: 'khac', user: ($('who').value || '').trim(), acct: 'LIST ORDER', files: 1, ver: ($('ver').textContent || '').replace('Phiên bản ', '') }, d)) }).catch(() => {}); } catch (e) {}
  }
  let lastD = null;
  async function rerun() {
    if (!lastD) return;
    const R = build(lastD, { prev: Number(String($('rpPrev').value).replace(/\D/g, '')) || 0 });
    out = await toXlsx(R); show(R);
    $('rpDl').hidden = false; $('rpDl').textContent = '⬇ Tải ' + out.name;
    return R;
  }
  async function remember(type, key, v1, v2, btn) {
    if (!($('who').value || '').trim()) { $('rpMsg').textContent = 'Nhập tên của bạn (góc trên bên phải) trước khi ghi nhớ.'; window.FFM_askName && window.FFM_askName(); return; }
    if (!v1) { $('rpMsg').textContent = type === 'kho' ? 'Chọn miền trước khi bấm Ghi nhớ.' : 'Nhập ACC trước khi bấm Ghi nhớ.'; return; }
    btn.disabled = true; btn.textContent = 'Đang lưu…';
    if (type === 'kho') SHARED.kho[key] = v1; else SHARED.acc[key.toUpperCase()] = [v1, v2];
    try { localStorage.setItem('ffm-rpmap', JSON.stringify(SHARED)); } catch (e) {}
    const API = window.FFM_API;
    if (/^https:/.test(API || '')) { try { await fetch(API, { method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ event: 'rpmap', type, key, v1, v2, user: ($('who').value || '').trim() }) }); } catch (e) { $('rpMsg').textContent = 'Chưa lưu lên bảng chung được (mất mạng?). Máy này vẫn nhớ, thử lại sau.'; } }
    logRp({ event: type === 'kho' ? 'Khai báo miền cho kho' : 'Khai báo ACC cho mã KH', names: `${key} → ${v1}${v2 ? ' / ' + v2 : ''}` });
    await rerun();
    $('rpMsg').textContent = `Đã ghi nhớ ${key} → ${v1}${v2 ? ' / ' + v2 : ''}. Báo cáo đã chạy lại.`;
  }
  $('rpGo').onclick = async () => {
    if (!file) return;
    if (!($('who').value || '').trim()) { $('rpMsg').textContent = 'Nhập tên của bạn (góc trên bên phải) trước khi chạy báo cáo.'; window.FFM_askName && window.FFM_askName(); return; }
    const btn = $('rpGo'); btn.disabled = true; $('rpMsg').textContent = 'Đang đọc file…';
    try {
      await Promise.all([window.FFM_DATA && window.FFM_DATA.ready, loadShared()]);
      if (!window.FFM_ACCMAP) $('rpMsg').textContent = 'Chưa tải được DMKH, cột ACC sẽ trống…';
      const D = readData(new Uint8Array(await file.arrayBuffer())); lastD = D;
      const R = build(D, { prev: Number(String($('rpPrev').value).replace(/\D/g, '')) || 0 });
      $('rpMsg').textContent = 'Đang tạo file Excel…';
      out = await toXlsx(R);
      show(R);
      $('rpMsg').textContent = `Đọc ${D.data.length.toLocaleString('vi-VN')} đơn${D.from ? ` (từ ${D.from} đến ${D.to})` : ''}. Bấm "Tải file kết quả".`;
      $('rpDl').hidden = false; $('rpDl').textContent = '⬇ Tải ' + out.name;
      logRp({ event: 'Báo cáo List order', orders: D.data.length, ok: 1, names: `${file.name} | ${D.from || '?'}–${D.to || '?'} | ${R.T.nw.tot.n} đơn, HĐ ${fmtTr(R.hdThang)}${R.extra.prev ? ', HĐ tháng trước ' + fmtTr(R.extra.prev) : ''}` });
    } catch (e) { $('rpMsg').textContent = '❌ ' + (e.message || e); $('rpOut').hidden = true; $('rpDl').hidden = true; logRp({ event: 'Báo cáo List order', ok: 0, err: 1, names: `${file.name} | lỗi: ${String(e.message || e).slice(0, 150)}` }); }
    finally { btn.disabled = false; }
  };
  $('rpDl').onclick = () => { if (!out) return; logRp({ event: 'Tải báo cáo List order', names: out.name }); const a = document.createElement('a'); a.href = URL.createObjectURL(out.blob); a.download = out.name; document.body.append(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000); };
})();
