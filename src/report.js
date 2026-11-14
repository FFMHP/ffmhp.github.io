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
  const kvOf = k => { const n = String(k || '').trim(); if (KV[n]) return KV[n]; const f = fold(n); for (const key in KV) if (fold(key) === f) return KV[key]; return ''; };
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
      so: col(/^ma phieu xuat/), code: col(/^ma khach hang/), val: col(/^tri gia don hang/), dvvc: col(/^ngay ban giao don hang cho dvvc/),
      inv: col(/^so hoa don/), amt: col(/^amount/), st: col(/^tinh trang don hang$/), sl: col(/smartlog/), kho: col(/^kho xuat/), stt: col(/^so tt/), de: col(/^ma phieu de xuat \(code\)/)
    };
    const miss = Object.entries(C).filter(([k, v]) => v < 0 && !['stt', 'de'].includes(k)).map(([k]) => k);
    if (miss.length) throw new Error('File thiếu cột: ' + miss.join(', '));
    const range = rows.slice(0, h).map(r => r.join(' ')).join(' ').match(/(\d{1,2}\/\d{1,2}\/\d{4})\s*đến ngày:?\s*(\d{1,2}\/\d{1,2}\/\d{4})/i);
    const data = rows.slice(h + 2).filter(r => String(r[C.de >= 0 ? C.de : C.so] || '').trim() || num(r[C.val]));
    return { head, C, data, from: range ? range[1] : '', to: range ? range[2] : '' };
  }

  function build(D, extra) {
    const A = window.FFM_ACCMAP || {};
    const accOf = code => { const c = String(code || '').trim().toUpperCase(); return A[c.replace(/HP$/, '')] || A[c] || null; };
    const C = D.C, nat = [], huy = [], noAcc = new Set(), noKv = {};
    D.data.forEach(r => {
      const acc = accOf(r[C.code]), kv = kvOf(r[C.kho]);
      if (!acc) noAcc.add(String(r[C.code]).trim());
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
    return { nat, huy, T, noAcc: [...noAcc], noKv, mm, prev, hdThang, extra, D };
  }

  // ---- ghi file Excel ----
  function toXlsx(R) {
    const D = R.D, head = ['ACC', 'Kênh phân phối', 'Khu Vực', 'Đã giao', 'Đã xuất HĐ', 'PO HỦY'].concat(D.head);
    const detail = list => [head].concat(list.map(o => [o.acc, o.kenh, o.kv || '#N/A', o.daGiao, o.daHD, o.isHuy].concat(D.head.map((_, i) => o.raw[i] === undefined ? '' : o.raw[i]))));
    const A = [[], ['', '', '', '', '', 1, `Trị giá đơn hàng tháng ${R.mm} đã xuất hóa đơn`, R.hdThang],
      ['', '', '', '', '', 2, `Trị giá đơn hàng tháng ${R.prev} đã xuất hóa đơn sang tháng ${R.mm}`, R.extra.prev || 0],
      ['', '', '', '', '', 3, `Trị giá đơn hàng hủy - Đã xuất hóa đơn thay thế trong tháng ${R.mm}`, R.extra.repl || 0],
      ['', '', '', '', '', 'Doanh số Net trừ hóa đơn hủy trong tháng (1+2)', '', R.hdThang + (R.extra.prev || 0)], []];
    const block = (title, region, t) => {
      A.push(['', '', '', '', region, title, 'Số lượng PO', 'Trị giá đơn hàng', 'Trị giá đã xuất HĐ trong tháng', 'Trị giá đơn hàng thiếu', region ? 'Note các đơn chưa giao' : '']);
      t.rows.forEach(x => {
        A.push(['', '', '', '', '', x.label, x.n, x.v, x.a, x.k === 'gx' || x.k === 'cx' ? x.thieu : '', x.note]);
        if (x.k === 'cc' && region !== 'Miền Trung') { A.push(['', '', '', '', '', '                     Đơn Win DC']); A.push(['', '', '', '', '', '                     CÁC ACC CÒN LẠI', x.n, x.v]); }
      });
      A.push(['', '', '', '', '', 'Đơn hàng hủy (không tính vào tổng)', t.huy.n, t.huy.v]);
      A.push(['', '', '', '', '', 'Tổng cộng', t.tot.n, t.tot.v, t.tot.a, t.tot.thieu]);
      A.push([]);
    };
    block('Tình trạng đơn hàng Nation Wide', '', R.T.nw);
    R.T.reg.forEach(x => block('Tình trạng đơn hàng ' + x.r, x.r, x.t));
    const wsT = XLSX.utils.aoa_to_sheet(A);
    wsT['!cols'] = [{ wch: 2 }, { wch: 2 }, { wch: 2 }, { wch: 2 }, { wch: 11 }, { wch: 44 }, { wch: 13 }, { wch: 18 }, { wch: 18 }, { wch: 18 }, { wch: 40 }];
    Object.keys(wsT).forEach(a => { const c = wsT[a]; if (a[0] !== '!' && c.t === 'n' && Math.abs(c.v) >= 1000) c.z = '#,##0'; });
    const wsN = XLSX.utils.aoa_to_sheet(detail(R.nat)), wsH = XLSX.utils.aoa_to_sheet(detail(R.huy));
    [wsN, wsH].forEach(ws => { ws['!cols'] = head.map((h, i) => ({ wch: i < 6 ? 10 : Math.min(28, Math.max(10, String(h).length * 0.6)) })); ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(1, ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']).e.r : 1), c: head.length - 1 } }) }; });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, wsT, 'Total');
    XLSX.utils.book_append_sheet(wb, wsN, 'National');
    XLSX.utils.book_append_sheet(wb, wsH, 'ĐƠN HỦY');
    const [d, mo, y] = (R.D.to || '').split('/');
    const name = `Total PO update ${d ? `${d.padStart(2, '0')}.${mo.padStart(2, '0')}.${y.slice(2)}` : new Date().toISOString().slice(0, 10)}.xlsx`;
    return { wb, name };
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
    const warn = [];
    const nk = Object.entries(R.noKv);
    if (nk.length) warn.push(`Không tính vào bảng theo miền (kho chưa có trong bảng Quy đổi): ${nk.map(([k, x]) => `${k} (${x.n} đơn, ${fmtTr(x.v)})`).join('; ')}.`);
    if (R.noAcc.length) warn.push(`${R.noAcc.length} mã khách hàng không có trong DMKH nên để trống ACC: ${R.noAcc.slice(0, 8).join(', ')}${R.noAcc.length > 8 ? '…' : ''}.`);
    if (warn.length) { const w = el('div', 'rp-warn'); warn.forEach(x => w.append(el('p', '', '⚠ ' + x))); box.append(w); }
  }

  $('rpGo').onclick = async () => {
    if (!file) return;
    const btn = $('rpGo'); btn.disabled = true; $('rpMsg').textContent = 'Đang đọc file…';
    try {
      await (window.FFM_DATA && window.FFM_DATA.ready);
      if (!window.FFM_ACCMAP) $('rpMsg').textContent = 'Chưa tải được DMKH, cột ACC sẽ trống…';
      const D = readData(new Uint8Array(await file.arrayBuffer()));
      const R = build(D, { prev: num($('rpPrev').value), repl: num($('rpRepl').value) });
      out = toXlsx(R);
      show(R);
      $('rpMsg').textContent = `Đọc ${D.data.length.toLocaleString('vi-VN')} đơn${D.from ? ` (từ ${D.from} đến ${D.to})` : ''}. Bấm "Tải file kết quả".`;
      $('rpDl').hidden = false; $('rpDl').textContent = '⬇ Tải ' + out.name;
      try { const API = window.FFM_API; if (/^https:/.test(API || '')) fetch(API, { method: 'POST', mode: 'no-cors', keepalive: true, headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ event: 'Báo cáo List order', team: 'khac', user: ($('who').value || '').trim(), acct: 'LISTORDER', files: 1, orders: D.data.length, ver: ($('ver').textContent || '').replace('Phiên bản ', '') }) }).catch(() => {}); } catch (e) {}
    } catch (e) { $('rpMsg').textContent = '❌ ' + (e.message || e); $('rpOut').hidden = true; $('rpDl').hidden = true; }
    finally { btn.disabled = false; }
  };
  $('rpDl').onclick = () => { if (out) XLSX.writeFile(out.wb, out.name, { compression: true }); };
})();
