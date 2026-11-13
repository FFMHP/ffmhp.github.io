/* Fujimart / BRG: đọc phiếu đặt hàng từ PDF (pdf.js, đọc chữ có sẵn) hoặc ảnh (Tesseract, nhận dạng chữ) — 0 token */
(function (root) {
  const fold = s => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase();
  const digits = s => String(s).replace(/\D/g, '');
  const med = a => { const b = a.slice().sort((x, y) => x - y); return b.length ? b[Math.floor(b.length / 2)] : 0; };

  function gtinOk(code) {
    if (!/^\d{8}$|^\d{12,14}$/.test(code)) return false;
    const d = code.split('').map(Number), check = d.pop();
    let sum = 0;
    for (let i = d.length - 1, w = 3; i >= 0; i--, w = w === 3 ? 1 : 3) sum += d[i] * w;
    return (10 - (sum % 10)) % 10 === check;
  }
  // "6.00" -> 6 · "51,200" / "51.200" -> 51200 · "2,567,352" -> 2567352
  function parseNum(t) {
    const s = String(t).replace(/[^\d.,]/g, '');
    if (!s) return null;
    if (/^\d{1,3}([.,]\d{3})+$/.test(s)) return Number(s.replace(/[.,]/g, ''));
    if (/^\d+[.,]\d{1,2}$/.test(s)) return Number(s.replace(',', '.'));
    if (/^\d+$/.test(s)) return Number(s);
    return null;
  }
  const isQtyTok = t => /^\d{1,4}[.,]\d{2}$/.test(String(t).replace(/[^\d.,]/g, ''));

  function groupLines(words) {
    const hs = med(words.map(w => w.y1 - w.y0)) || 10;
    const sorted = words.slice().sort((a, b) => (a.y0 + a.y1) - (b.y0 + b.y1));
    const lines = [];
    for (const w of sorted) {
      const yc = (w.y0 + w.y1) / 2;
      const ln = lines.find(l => Math.abs(l.yc - yc) < hs * 0.6);
      if (ln) { ln.words.push(w); ln.yc = (ln.yc * (ln.words.length - 1) + yc) / ln.words.length; }
      else lines.push({ yc, words: [w] });
    }
    lines.forEach(l => { l.words.sort((a, b) => a.x0 - b.x0); l.text = l.words.map(w => w.t).join(' '); });
    return { lines, hs };
  }

  /* words: [{t, x0, y0, x1, y1}] — cùng một hàm cho PDF và ảnh */
  function parseWords(words) {
    words = words.filter(w => String(w.t).trim());
    const { lines, hs } = groupLines(words);
    const out = { rows: [], issues: [] };

    // barcode: từ 12–14 chữ số (cho phép 1–2 ký tự nhiễu)
    const cand = words.filter(w => { const d = digits(w.t); return d.length >= 11 && d.length <= 14 && d.length >= String(w.t).replace(/\s/g, '').length - 2; });
    if (!cand.length) { out.error = 'Không tìm thấy cột Mã vạch (barcode) trên phiếu.'; return out; }
    const cx = med(cand.map(w => (w.x0 + w.x1) / 2)), cw = med(cand.map(w => w.x1 - w.x0));
    const bcs = cand.filter(w => Math.abs((w.x0 + w.x1) / 2 - cx) < cw * 0.6).sort((a, b) => a.y0 - b.y0);
    out.cols = { bc: { x0: cx - cw * 0.65, x1: cx + cw * 0.65 } };

    const qtyX = [], priceX = [], amtX = [];
    for (const b of bcs) {
      const yc = (b.y0 + b.y1) / 2;
      const right = words.filter(w => Math.abs((w.y0 + w.y1) / 2 - yc) < hs * 0.7 && w.x0 > b.x1).sort((a, c) => a.x0 - c.x0);
      const nums = right.filter(w => parseNum(w.t) != null);
      const qi = nums.findIndex(w => isQtyTok(w.t));
      const q = qi >= 0 ? nums[qi] : null, p = qi >= 0 ? nums[qi + 1] : null, a = qi >= 0 ? nums[qi + 2] : null;
      if (q) qtyX.push(q); if (p) priceX.push(p); if (a) amtX.push(a);
      out.rows.push({ y0: Math.min(b.y0, ...(q ? [q.y0] : [])), y1: Math.max(b.y1, ...(q ? [q.y1] : [])), barcode: digits(b.t), qty: q ? parseNum(q.t) : null, price: p ? parseNum(p.t) : null, amount: a ? parseNum(a.t) : null });
    }
    const colOf = arr => arr.length ? { x0: med(arr.map(w => w.x0)) - hs * 0.8, x1: med(arr.map(w => w.x1)) + hs * 0.8 } : null;
    out.cols.qty = colOf(qtyX); out.cols.price = colOf(priceX); out.cols.amt = colOf(amtX);
    out.rowH = hs;

    const lineHas = re => lines.find(l => re.test(fold(l.text)));
    // số đơn: 15–20 chữ số, ưu tiên dòng "Số Đơn"
    const poLine = lineHas(/so\s*d[oơ]n/);
    const poFrom = s => { const m = digits(s).match(/\d{18}/); return m ? m[0] : ''; };
    out.po = poLine ? (String(poLine.text).match(/\d[\d\s]{14,24}\d/g) || []).map(poFrom).find(Boolean) || '' : '';
    if (!out.po) { const w = words.find(w => /^\d{18}$/.test(digits(w.t)) && digits(w.t).length === String(w.t).replace(/\W/g, '').length); out.po = w ? digits(w.t) : ''; }
    const dl = lineHas(/ngay\s*dat/);
    const dm = dl && dl.text.match(/(\d{1,2})\s*\/\s*(\d{1,2})\s*\/\s*(20\d\d)/);
    if (dm) out.orderDate = `${dm[3]}-${String(dm[2]).padStart(2, '0')}-${String(dm[1]).padStart(2, '0')}`;
    const nl = lineHas(/noi\s*nhan/);
    if (nl) out.store = nl.text.replace(/^.*?n[ơo]i\s*nh[ậa]n\s*:?/i, '').replace(/^\s*\d+\s*/, '').trim();
    const tl = lineHas(/tai\s*tram/);
    if (tl) { const m = tl.text.match(/tr[ạa]m\s*:?\s*(\S+)/i); if (m) out.tram = m[1]; }
    const tot = lineHas(/tong\s*gia\s*tri\s*d\w{1,3}\s*hang/);
    if (tot) { const n = tot.words.map(w => parseNum(w.t)).filter(v => v != null); out.total = n.length ? n[n.length - 1] : null; }
    return out;
  }

  /* ---------- PDF: lấy chữ + toạ độ bằng pdf.js ---------- */
  // các đoạn chữ của PDF giữ nguyên từng ô (không tách theo dấu cách) — dùng cho PO Miomio
  async function pdfItems(buf) {
    const pdfjsLib = root.pdfjsLib;
    const doc = await pdfjsLib.getDocument({ data: new Uint8Array(buf).slice(), isEvalSupported: false }).promise;
    const out = [];
    for (let pn = 1; pn <= doc.numPages; pn++) {
      const page = await doc.getPage(pn);
      const vp = page.getViewport({ scale: 1 });
      const tc = await page.getTextContent();
      for (const it of tc.items) {
        const str = (it.str || '').trim(); if (!str) continue;
        const [, , , , e, f] = it.transform; const h = it.height || Math.abs(it.transform[3]) || 8;
        const y1 = vp.height - f + (pn - 1) * 10000;
        out.push({ t: str, x0: e, x1: e + (it.width || str.length * h * 0.5), y0: y1 - h, y1, page: pn });
      }
    }
    return out;
  }

  async function pdfWords(buf) {
    const pdfjsLib = root.pdfjsLib;
    const doc = await pdfjsLib.getDocument({ data: new Uint8Array(buf).slice(), isEvalSupported: false }).promise;
    const page = await doc.getPage(1);
    const vp = page.getViewport({ scale: 1 });
    const tc = await page.getTextContent();
    const words = [];
    for (const it of tc.items) {
      const str = it.str; if (!str || !str.trim()) continue;
      const [, , , , e, f] = it.transform; const h = it.height || Math.abs(it.transform[3]) || 8;
      const x0 = e, y1 = vp.height - f, y0 = y1 - h, wTot = it.width || str.length * h * 0.5;
      // tách theo khoảng trắng, chia bề rộng theo số ký tự
      let pos = 0;
      for (const part of str.split(/(\s+)/)) {
        const w = wTot * part.length / str.length;
        if (part.trim()) words.push({ t: part, x0: x0 + pos, x1: x0 + pos + w, y0, y1 });
        pos += w;
      }
    }
    return words;
  }

  /* ---------- Ảnh: Tesseract (vie+eng), đọc cả trang rồi đọc lại từng ô chỉ cho phép chữ số ---------- */
  let workerP = null;
  function getWorker(opts) {
    if (!workerP) workerP = root.Tesseract.createWorker(['vie', 'eng'], 1, Object.assign({ logger: () => {} }, opts || {}));
    return workerP;
  }
  async function loadImage(blob) {
    const url = URL.createObjectURL(blob);
    try { const img = new Image(); img.src = url; await img.decode(); return img; } finally { setTimeout(() => URL.revokeObjectURL(url), 5000); }
  }
  // mode: 'fixed' ngưỡng 170 · 'otsu' ngưỡng tự động · 'gray' chỉ tăng tương phản · 'none' giữ nguyên (ảnh xem trước)
  function cropCanvas(img, box, scale, mode) {
    mode = mode || 'fixed';
    const c = document.createElement('canvas');
    const w = Math.max(1, Math.round(box.x1 - box.x0)), h = Math.max(1, Math.round(box.y1 - box.y0));
    c.width = Math.round(w * scale); c.height = Math.round(h * scale);
    const g = c.getContext('2d', { willReadFrequently: true });
    g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height);
    g.imageSmoothingEnabled = true; g.imageSmoothingQuality = 'high';
    g.drawImage(img, box.x0, box.y0, w, h, 0, 0, c.width, c.height);
    if (mode === 'none') return c;
    const d = g.getImageData(0, 0, c.width, c.height), p = d.data, n = p.length / 4, L = new Uint8Array(n);
    let lo = 255, hi = 0;
    for (let i = 0; i < n; i++) { const v = p[i * 4] * 0.299 + p[i * 4 + 1] * 0.587 + p[i * 4 + 2] * 0.114; L[i] = v; if (v < lo) lo = v; if (v > hi) hi = v; }
    let t = 170;
    if (mode === 'otsu') {
      const hist = new Array(256).fill(0); for (let i = 0; i < n; i++) hist[L[i]]++;
      let sum = 0; for (let i = 0; i < 256; i++) sum += i * hist[i];
      let sB = 0, wB = 0, best = 0;
      for (let i = 0; i < 256; i++) { wB += hist[i]; if (!wB) continue; const wF = n - wB; if (!wF) break; sB += i * hist[i]; const mB = sB / wB, mF = (sum - sB) / wF, v = wB * wF * (mB - mF) * (mB - mF); if (v > best) { best = v; t = i; } }
    }
    for (let i = 0; i < n; i++) {
      const v = mode === 'gray' ? Math.max(0, Math.min(255, (L[i] - lo) * 255 / Math.max(1, hi - lo))) : (L[i] > t ? 255 : 0);
      p[i * 4] = p[i * 4 + 1] = p[i * 4 + 2] = v;
    }
    g.putImageData(d, 0, 0);
    return c;
  }
  const BC_VARIANTS = [{ s: 4, m: 'otsu' }, { s: 3, m: 'gray' }, { s: 5, m: 'fixed' }, { s: 6, m: 'otsu' }, { s: 4, m: 'gray' }, { s: 3, m: 'otsu' }, { s: 5, m: 'gray' }];
  const NUM_VARIANTS = [{ s: 4, m: 'otsu' }, { s: 3, m: 'gray' }, { s: 5, m: 'fixed' }];
  async function readVar(worker, img, box, whitelist, v) {
    await worker.setParameters({ tessedit_char_whitelist: whitelist, tessedit_pageseg_mode: '7' });
    const { data } = await worker.recognize(cropCanvas(img, box, v.s, v.m));
    return (data.text || '').replace(/\s+/g, '');
  }
  // các chữ số hay bị nhầm khi ảnh mờ
  const CONFUSE = { 0: '869', 1: '74', 2: '7', 3: '85', 4: '1', 5: '63', 6: '85', 7: '12', 8: '3609', 9: '80' };
  function pickBarcode(reads) {
    const ds = reads.map(digits).filter(x => x.length >= 8);
    if (!ds.length) return { code: '', bad: true };
    const cnt = {}; ds.forEach(x => cnt[x] = (cnt[x] || 0) + 1);
    const valid = Object.keys(cnt).filter(gtinOk).sort((a, b) => cnt[b] - cnt[a]);
    if (valid.length) return { code: valid[0], sure: cnt[valid[0]] >= 2 && (valid.length === 1 || cnt[valid[0]] > cnt[valid[1]]), votes: cnt[valid[0]] };
    const lc = {}; ds.forEach(x => lc[x.length] = (lc[x.length] || 0) + 1);
    const len = +Object.keys(lc).sort((a, b) => ([8, 12, 13, 14].includes(+b) - [8, 12, 13, 14].includes(+a)) || lc[b] - lc[a])[0];
    const same = ds.filter(x => x.length === len);
    const opts = [];
    for (let i = 0; i < len; i++) { const c = {}; same.forEach(x => c[x[i]] = (c[x[i]] || 0) + 1); opts.push(Object.keys(c).sort((a, b) => c[b] - c[a])); }
    const vote = opts.map(o => o[0]).join('');
    if (gtinOk(vote)) return { code: vote, guessed: true };
    // thử tổ hợp các chữ số mà các lần đọc không thống nhất
    let combos = [''];
    for (const o of opts) { combos = combos.flatMap(p => o.map(d => p + d)); if (combos.length > 4096) break; }
    const okC = combos.length <= 4096 ? [...new Set(combos.filter(gtinOk))] : [];
    if (okC.length === 1) return { code: okC[0], guessed: true };
    // thử sửa 1 chữ số hay bị nhầm
    const fix = new Set();
    for (let i = 0; i < vote.length; i++) for (const d of (CONFUSE[vote[i]] || '')) { const t = vote.slice(0, i) + d + vote.slice(i + 1); if (gtinOk(t)) fix.add(t); }
    if (fix.size === 1) return { code: [...fix][0], guessed: true };
    return { code: vote, bad: true };
  }
  // khớp với danh mục barcode đã biết: chấm điểm theo số chữ số trùng ở từng lần đọc
  function snapCatalog(reads, catalog) {
    const ds = reads.map(digits).filter(x => x.length >= 8);
    if (!ds.length || !catalog || !catalog.length) return null;
    const L = ds[0].length;
    catalog = catalog.map(c => c.length === 14 ? (/^0+$/.test(c.slice(0, 14 - L)) ? c.slice(14 - L) : null) : c).filter(Boolean);
    const scored = catalog.map(c => {
      let match = 0, n = 0;
      for (const r of ds) { if (r.length !== c.length) continue; n++; for (let i = 0; i < c.length; i++) if (r[i] === c[i]) match++; }
      return { c, miss: n ? (n * c.length - match) / n : 99 };
    }).sort((a, b) => a.miss - b.miss);
    const b = scored[0], s2 = scored[1];
    if (b && b.miss <= 3 && (!s2 || s2.miss - b.miss >= 1)) return b.c;
    return null;
  }
  // cột "Theo đơn" luôn có dạng 12.00 — đọc mất dấu chấm thành 1200 thì chia 100
  function parseQty(t) {
    const s = String(t).replace(/[^\d.,]/g, '');
    if (!s) return null;
    if (/[.,]/.test(s)) return parseNum(s.replace(/[.,](?=\d{3}\b)/g, '')) ?? parseNum(s);
    if (/^\d{3,}$/.test(s) && /00$/.test(s)) return Number(s) / 100;
    return Number(s);
  }
  async function readCell(worker, img, box, whitelist) {
    await worker.setParameters({ tessedit_char_whitelist: whitelist, tessedit_pageseg_mode: '7' });
    const { data } = await worker.recognize(cropCanvas(img, box, 3));
    return (data.text || '').replace(/\s+/g, '');
  }

  /* tìm lưới bảng (đường kẻ ngang/dọc) trên ảnh đã phóng to */
  function findGrid(canvas) {
    const W = canvas.width, H = canvas.height, g = canvas.getContext('2d', { willReadFrequently: true });
    const p = g.getImageData(0, 0, W, H).data, dark = new Uint8Array(W * H);
    for (let i = 0, j = 0; i < p.length; i += 4, j++) dark[j] = (p[i] * 0.299 + p[i + 1] * 0.587 + p[i + 2] * 0.114) < 150 ? 1 : 0;
    // dòng kẻ ngang: chuỗi điểm tối liên tục dài nhất trên mỗi hàng
    const runRow = y => { let best = 0, cur = 0, s0 = 0, bs = 0; for (let x = 0; x < W; x++) { if (dark[y * W + x]) { if (!cur) s0 = x; cur++; if (cur > best) { best = cur; bs = s0; } } else cur = 0; } return { len: best, x0: bs, x1: bs + best }; };
    const hl = [];
    for (let y = 0; y < H; y++) { const r = runRow(y); if (r.len > W * 0.5) hl.push({ y, x0: r.x0, x1: r.x1 }); }
    const group = (arr, key) => { const out = []; for (const v of arr) { const last = out[out.length - 1]; if (last && v[key] - last.end <= 3) { last.end = v[key]; last.items.push(v); } else out.push({ start: v[key], end: v[key], items: [v] }); } return out.map(o => Object.assign({ c: (o.start + o.end) / 2 }, o)); };
    const hLines = group(hl, 'y');
    if (hLines.length < 3) return null;
    // chọn cụm đường kẻ ngang cách đều nhau nhiều nhất (thân bảng)
    const tx0 = med(hLines.map(l => l.items[0].x0)), tx1 = med(hLines.map(l => l.items[0].x1));
    const top = hLines[0].c, bot = hLines[hLines.length - 1].c;
    const vl = [];
    for (let x = Math.max(0, Math.round(tx0) - 3); x <= Math.min(W - 1, Math.round(tx1) + 3); x++) {
      let n = 0; for (let y = Math.round(top); y <= Math.round(bot); y++) n += dark[y * W + x];
      if (n > (bot - top) * 0.6) vl.push({ x });
    }
    const vLines = group(vl, 'x');
    if (vLines.length < 4) return null;
    return { rows: hLines.map(l => l.c), cols: vLines.map(l => l.c) };
  }
  async function readBox(worker, src, box, whitelist, psm) {
    await worker.setParameters({ tessedit_char_whitelist: whitelist || '', tessedit_pageseg_mode: psm || '7' });
    const { data } = await worker.recognize(cropCanvas(src, box, 3));
    return (data.text || '').replace(/\s+/g, ' ').trim();
  }

  async function imageParse(blob, onProgress, tessOpts) {
    const img = await loadImage(blob);
    const k = Math.max(1, 1600 / img.naturalWidth), src = document.createElement('canvas');
    src.width = Math.round(img.naturalWidth * k); src.height = Math.round(img.naturalHeight * k);
    { const g = src.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, src.width, src.height); g.imageSmoothingQuality = 'high'; g.drawImage(img, 0, 0, src.width, src.height); }
    onProgress && onProgress('Đang tải bộ nhận dạng chữ (lần đầu cần internet)…');
    const worker = await getWorker(tessOpts);
    await worker.setParameters({ tessedit_char_whitelist: '', tessedit_pageseg_mode: '3' });
    onProgress && onProgress('Đang đọc thông tin đơn…');
    const { data } = await worker.recognize(src);
    const words = (data.words || []).map(w => ({ t: w.text, x0: w.bbox.x0, y0: w.bbox.y0, x1: w.bbox.x1, y1: w.bbox.y1 }));
    const res = parseWords(words);
    delete res.error; res.rows = [];
    // ảnh mờ: đọc lại góc trên bên phải (Số Đơn, Ngày đặt) ở độ phóng lớn hơn
    if (!res.po || !res.orderDate) {
      const box = { x0: img.naturalWidth * 0.55, x1: img.naturalWidth, y0: 0, y1: img.naturalHeight * 0.2 };
      for (const v of [{ s: 4, m: 'gray' }, { s: 3, m: 'otsu' }]) {
        await worker.setParameters({ tessedit_char_whitelist: '', tessedit_pageseg_mode: '6' });
        const t = ((await worker.recognize(cropCanvas(img, box, v.s, v.m))).data.text || '');
        if (!res.po) { const m = t.replace(/[ \t]/g, '').match(/\d{18}/); if (m) res.po = m[0]; }
        if (!res.orderDate) { const ln = t.split('\n').find(x => /ng[aà]y\s*[dđ][aặ]t/i.test(x)) || ''; const m = ln.match(/(\d{1,2})\s*\/\s*(\d{1,2})\s*\/\s*(20\d\d)/); if (m) res.orderDate = `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`; }
        if (res.po && res.orderDate) break;
      }
    }
    const grid = findGrid(src);
    if (!grid) { res.error = 'Không tìm thấy bảng hàng (đường kẻ bảng) trong ảnh. Ảnh có bị mờ, nghiêng hay cắt mất bảng không?'; return res; }
    const inset = 4, cell = (ri, ci) => ({ x0: grid.cols[ci] + inset, x1: grid.cols[ci + 1] - inset, y0: grid.rows[ri] + inset, y1: grid.rows[ri + 1] - inset });
    // dòng tiêu đề: đọc chữ để biết cột nào là Mã vạch / Theo đơn / Đơn giá / Thành tiền
    const nC = grid.cols.length - 1, heads = [];
    for (let ci = 0; ci < nC; ci++) heads.push(fold(await readBox(worker, src, cell(0, ci), '', '6')));
    const find = re => heads.findIndex(h => re.test(h));
    const col = { bc: find(/vach|barcode/), qty: find(/theo|so\s*luong/), price: find(/don\s*gia|gia/), amt: find(/thanh\s*tien|tien/) };
    if (col.bc < 0) col.bc = 3; if (col.qty < 0) col.qty = col.bc + 2; if (col.price < 0) col.price = nC - 2; if (col.amt < 0) col.amt = nC - 1;
    res.headers = heads;
    const toOrig = bx => ({ x0: bx.x0 / k, x1: bx.x1 / k, y0: bx.y0 / k, y1: bx.y1 / k });
    for (let ri = 1; ri < grid.rows.length - 1; ri++) {
      onProgress && onProgress(`Đang đọc dòng hàng ${ri}/${grid.rows.length - 2}…`);
      const bcBox = toOrig(cell(ri, col.bc));
      const first = await readVar(worker, img, bcBox, '0123456789', BC_VARIANTS[0]);
      if (digits(first).length < 8) continue; // dòng trống / dòng tổng
      const reads = [first];
      for (let vi = 1; vi < BC_VARIANTS.length; vi++) {
        const pk = pickBarcode(reads); if (pk.sure) break;
        reads.push(await readVar(worker, img, bcBox, '0123456789', BC_VARIANTS[vi]));
      }
      const pk = pickBarcode(reads);
      let snapped = false;
      const inCat = c => !!(root.FFM_DMSP && root.FFM_DMSP[String(c).padStart(14, '0')]);
      // đọc "chắc" nhưng không có trong danh mục thì vẫn coi là chưa chắc
      if (pk.sure && root.FFM_DMSP && !inCat(pk.code)) pk.sure = false;
      if (!pk.sure) { const sc = snapCatalog(reads, root.FFM_CATALOG); if (sc) { pk.code = sc; snapped = true; } }
      const nums = async (ci, parse) => { const out = []; for (const v of NUM_VARIANTS) { const x = parse(await readVar(worker, img, toOrig(cell(ri, ci)), '0123456789.,', v)); if (x != null && !Number.isNaN(x)) out.push(x); if (out.length >= 2 && out[0] === out[1]) break; } return out; };
      const qs = await nums(col.qty, parseQty), ps = await nums(col.price, parseNum), as = await nums(col.amt, parseNum);
      let q = qs[0] ?? null, p = ps[0] ?? null, a = as[0] ?? null, verified = false;
      outer: for (const qq of qs) for (const pp of ps) for (const aa of as) if (Math.abs(qq * pp - aa) <= 1) { q = qq; p = pp; a = aa; verified = true; break outer; }
      // xem trước: cả dòng từ cột Mã vạch đến Thành tiền
      const strip = toOrig({ x0: grid.cols[col.bc], x1: grid.cols[Math.min(nC, col.amt + 1)], y0: grid.rows[ri], y1: grid.rows[ri + 1] });
      let preview = '';
      try { preview = cropCanvas(img, strip, Math.min(3, Math.max(1, 28 / Math.max(1, strip.y1 - strip.y0))), 'none').toDataURL('image/jpeg', 0.85); } catch (e) {}
      res.rows.push({ reads, barcode: pk.code, bcSnapped: snapped, bcGuessed: !!pk.guessed && !snapped, bcSure: !!pk.sure, qty: q, price: p, amount: a, verified, preview });
    }
    if (!res.rows.length) res.error = 'Không đọc được dòng hàng nào trong bảng.';
    return res;
  }

  /* ---------- AEON: PDF "STORE ORDER" (font không có bảng mã chữ -> phải nhận dạng) ----------
     Mỗi đoạn chữ trong PDF có vị trí chính xác. Công cụ vẽ trang thành ảnh nét, nhận dạng từng đoạn,
     rồi học bảng "mã ký tự -> chữ" từ chính file (bỏ phiếu) để giải mã lại toàn bộ cho chắc. */

  /* ---------- KKTC: ảnh chụp màn hình bảng đơn Kmart ----------
     Cột: Tên SP | Mã SP | Quy cách | Barcode | ĐVT | SL cái | Giá EA | Giá ĐV | SL ĐV
     Đọc từng dòng từ phải sang trái: SL ĐV, Giá ĐV, Giá EA, SL cái, ĐVT, rồi barcode. SL điền form = SL ĐV × SL cái. */
  const isInt = t => /^\d{1,5}$/.test(t);
  const isMoney = t => /^\d{1,3}([.,]\d{3})+$|^\d{1,7}$/.test(t);
  const money = t => Number(String(t).replace(/[.,]/g, ''));
  function parseKktcTokens(toks) {
    const t = toks.map(x => x.t.replace(/[|\[\]{}]/g, '')).filter(Boolean), w = toks.filter(x => x.t.replace(/[|\[\]{}]/g, ''));
    let i = t.length - 1;
    if (i < 4) return null;
    const slDv = t[i], giaDv = t[i - 1], giaEa = t[i - 2], slCai = t[i - 3];
    if (!isInt(slDv) || !isMoney(giaDv) || !isMoney(giaEa) || !isInt(slCai)) return null;
    i -= 4;
    let dvt = '';
    if (i >= 0 && /^[A-Za-zĐđ]{1,6}$/.test(t[i])) { dvt = t[i]; i--; }
    // barcode: ghép các cụm số liền nhau bên trái ĐVT (OCR hay chèn dấu cách giữa barcode)
    let bc = '', bcWords = [], best = null;
    for (let j = i; j >= 0 && j >= i - 3; j--) {
      if (!/^\d+$/.test(t[j])) break;
      bc = t[j] + bc; bcWords.unshift(w[j]);
      if (bc.length > 14) break;
      if (bc.length >= 8 && (!best || gtinOk(bc))) { best = { code: bc, words: bcWords.slice(), j }; if (gtinOk(bc)) break; }
    }
    if (!best) return null;
    const rest = t.slice(0, best.j);
    const kmCode = (rest.find(x => /^\d{6,8}$/.test(x)) || '');
    const name = rest.slice(0, Math.max(0, rest.findIndex(x => /^\d{6,8}$/.test(x)))).join(' ');
    return { barcode: best.code, bcWords: best.words, slDv: +slDv, giaDv: money(giaDv), giaEa: money(giaEa), slCai: +slCai, dvt, kmCode, name };
  }

  async function kktcImageParse(blob, onProgress, tessOpts) {
    const img = await loadImage(blob);
    const W = img.naturalWidth, H = img.naturalHeight;
    const k = W < 1000 ? 3 : W < 1800 ? 2 : 1;
    onProgress && onProgress('Đang tải bộ nhận dạng chữ (lần đầu cần internet)…');
    const worker = await getWorker(tessOpts);
    await worker.setParameters({ tessedit_char_whitelist: '', tessedit_pageseg_mode: '6', preserve_interword_spaces: '0' });
    onProgress && onProgress('Đang đọc bảng hàng…');
    const big = cropCanvas(img, { x0: 0, y0: 0, x1: W, y1: H }, k, 'gray');
    const { data } = await worker.recognize(big);
    const words = (data.words || []).map(x => ({ t: x.text, x0: x.bbox.x0 / k, y0: x.bbox.y0 / k, x1: x.bbox.x1 / k, y1: x.bbox.y1 / k }));
    const { lines } = groupLines(words);
    const hdr = lines.findIndex(L => L.words.some(x => /barcode/i.test(x.t)));
    const rows = [];
    let skipped = 0;
    for (const L of lines.slice(hdr + 1)) {
      const toks = L.words.slice().sort((a, b) => a.x0 - b.x0);
      const p = parseKktcTokens(toks);
      if (!p) { if (toks.some(x => /\d{6,}/.test(x.t))) skipped++; continue; }
      const y0 = Math.max(0, Math.min(...toks.map(x => x.y0)) - 3), y1 = Math.min(H, Math.max(...toks.map(x => x.y1)) + 3);
      let bcSure = gtinOk(p.barcode), bcGuessed = false;
      if (!bcSure) {
        // đọc lại riêng ô barcode, chỉ cho phép chữ số
        const bx = { x0: Math.max(0, Math.min(...p.bcWords.map(x => x.x0)) - 4), x1: Math.min(W, Math.max(...p.bcWords.map(x => x.x1)) + 4), y0, y1 };
        const reads = [];
        for (const v of BC_VARIANTS.slice(0, 4)) reads.push(await readVar(worker, img, bx, '0123456789', v));
        const pick = pickBarcode(reads);
        if (pick.code && gtinOk(pick.code)) { p.barcode = pick.code; bcSure = !!pick.sure; bcGuessed = !pick.sure; }
        await worker.setParameters({ tessedit_char_whitelist: '', tessedit_pageseg_mode: '6' });
      }
      const preview = cropCanvas(img, { x0: 0, x1: W, y0, y1 }, 1, 'none').toDataURL('image/png');
      rows.push(Object.assign(p, { bcSure, bcGuessed, preview }));
    }
    if (!rows.length) return { error: 'Không đọc được dòng hàng nào. Ảnh có đủ các cột Barcode, ĐVT, SL cái, Giá EA, Giá ĐV, SL ĐV không?' };
    return { rows, skipped, header: hdr >= 0 };
  }

  async function aeonPdfParse(buf, onProgress, tessOpts) {
    const pdfjsLib = root.pdfjsLib;
    const doc = await pdfjsLib.getDocument({ data: new Uint8Array(buf).slice(), isEvalSupported: false }).promise;
    onProgress && onProgress('Đang tải bộ nhận dạng chữ…');
    const worker = await getWorker(tessOpts);
    const S = 3, toks = [];
    for (let pn = 1; pn <= doc.numPages; pn++) {
      const page = await doc.getPage(pn);
      const vp1 = page.getViewport({ scale: 1 }), vp = page.getViewport({ scale: S });
      const cv = document.createElement('canvas'); cv.width = Math.ceil(vp.width); cv.height = Math.ceil(vp.height);
      const g = cv.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, cv.width, cv.height);
      await page.render({ canvasContext: g, viewport: vp }).promise;
      const tc = await page.getTextContent();
      // ghép các đoạn sát nhau trên cùng dòng thành 1 cụm (font PDF này tách chữ số thành nhiều đoạn)
      const its = tc.items.filter(i => i.str && i.str.length && i.width > 0 && !(/^ +$/.test(i.str) && i.width > 1.5 * (i.height || 8) * i.str.length)).map(i => { const h = i.height || 8, yb = vp1.height - i.transform[5]; return { x0: i.transform[4], x1: i.transform[4] + i.width, y: yb, h }; }).sort((p, q) => (Math.round(p.y) - Math.round(q.y)) || (p.x0 - q.x0));
      const groups = [];
      for (const it of its) {
        const gq = groups[groups.length - 1];
        if (gq && Math.abs(gq.y - it.y) < 1.5 && it.x0 - gq.x1 < 0.8 && it.x0 >= gq.x1 - 1.5) { gq.x1 = Math.max(gq.x1, it.x1); gq.h = Math.max(gq.h, it.h); }
        else groups.push(Object.assign({}, it));
      }
      for (let k = 0; k < groups.length; k++) {
        const t = groups[k];
        // nới phải tối đa 1 ký tự (chữ số "6" của font này hay bị tách thành đoạn riêng), không lấn sang cụm kế
        const nxt = groups.slice(k + 1).find(q => Math.abs(q.y - t.y) < 1.5);
        t.x1 = Math.min(t.x1 + t.h * 0.75, nxt ? nxt.x0 - 0.6 : t.x1 + t.h * 0.75);
        if (k % 8 === 0) onProgress && onProgress(`Trang ${pn}/${doc.numPages}: đang đọc ${k + 1}/${groups.length} cụm chữ…`);
        const box = { x0: (t.x0 - 1.2) * S, x1: (t.x1 + 1.2) * S, y0: (t.y - t.h - 1.2) * S, y1: (t.y + t.h * 0.3 + 1.2) * S };
        await worker.setParameters({ tessedit_char_whitelist: '', tessedit_pageseg_mode: '7' });
        let text = ((await worker.recognize(cropCanvas(cv, box, 1, 'none'))).data.text || '').replace(/\s+/g, ' ').trim();
        // cụm trông như số: đọc lại chỉ cho phép chữ số
        if (/^[\d.,\s|:/-]{1,24}$/.test(text) && /\d/.test(text)) {
          await worker.setParameters({ tessedit_char_whitelist: '0123456789.,:/', tessedit_pageseg_mode: '7' });
          const t2 = ((await worker.recognize(cropCanvas(cv, box, 1, 'none'))).data.text || '').replace(/\s+/g, '').trim();
          if (t2 && digits(t2).length >= digits(text).length) text = t2;
        }
        text = text.replace(/^[\s|\[\]{}]+|[\s|\[\]{}]+$/g, '');
        toks.push({ page: pn, t: text, x0: t.x0, x1: t.x1, y: t.y, h: t.h });
      }
    }
    const items = toks;
    const p1 = items.filter(i => i.page === 1);
    const res = { rows: [] };
    const poIt = p1.filter(i => /^\d{10,16}$/.test(i.t) && i.y < 80).sort((a, b) => b.x0 - a.x0)[0];
    res.po = poIt ? poIt.t : '';
    const tb = p1.find(i => /tao\s*boi/.test(fold(i.t)));
    if (tb) { const m = tb.t.match(/:\s*(\d{3,5})\s*(.*)$/); if (m) { res.storeCode = m[1]; res.storeName = m[2].trim(); } }
    if (tb && !res.storeName) { const nx = p1.filter(i => Math.abs(i.y - tb.y) < 2 && i.x0 > tb.x1 && i.x0 < tb.x1 + 120).sort((a, b) => a.x0 - b.x0)[0]; if (nx) res.storeName = nx.t; }
    const nd = p1.find(i => /ngay\s*dat/.test(fold(i.t)));
    if (nd) {
      const below = p1.filter(i => i.y > nd.y + 2 && i.y < nd.y + 20 && i.x0 >= nd.x0 - 25 && i.x1 <= nd.x1 + 25 && /^\d{1,4}$/.test(i.t)).sort((a, b) => a.x0 - b.x0);
      if (below.length >= 3) { const [d, m, y] = below.slice(0, 3).map(i => i.t); if (/^20\d\d$/.test(y)) res.orderDate = `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`; }
    }
    const dc = p1.find(i => /^dc\s*\d/i.test(i.t)); res.loc = dc ? dc.t : '';
    res.note = items.map(i => i.t).join(' ');
    const isNum = t => /^[\d.,]+$/.test(t);
    for (const b of items.filter(i => /^\d{8,14}$/.test(i.t) && i !== poIt)) {
      const ln = items.filter(i => i.page === b.page && Math.abs(i.y - b.y) < 2).sort((p, q) => p.x0 - q.x0);
      const bi = ln.indexOf(b), left = ln.slice(0, bi), right = ln.slice(bi + 1);
      if (!right.length || !isNum(right[0].t) || !/[.,]/.test(right[0].t)) continue;
      const item = left.length && /^\d{5,10}$/.test(left[left.length - 1].t) ? left[left.length - 1].t : '';
      const li = left.length - (item ? 1 : 0);
      const unit = li > 0 && !isNum(left[li - 1].t) ? left[li - 1].t : '';
      const pi = li - (unit ? 1 : 0) - 1;
      const pack = pi >= 0 && isNum(left[pi].t) ? parseNum(left[pi].t) : null;
      const nameParts = left.slice(0, Math.max(0, pi)).filter((x, i) => !(i === 0 && /^\d{1,3}$/.test(x.t)));
      const nums = right.filter(i => isNum(i.t)).map(i => parseNum(i.t));
      const cartons = nums[0] ?? null, price = nums.length >= 3 ? nums[nums.length - 2] : null, amount = nums.length >= 2 ? nums[nums.length - 1] : null;
      res.rows.push({ barcode: b.t, item, unit, pack, cartons, price, amount, aeonName: nameParts.map(i => i.t).join(' ').replace(/^\d{1,3}\s+/, ''), verified: cartons != null && price != null && amount != null && Math.abs(cartons * price - amount) <= 1 });
    }
    const tot = p1.find(i => /^tong\s*tien$/.test(fold(i.t).replace(/[^a-z\s]/g, '').trim()));
    if (tot) { const v = p1.filter(i => Math.abs(i.y - tot.y) < 2 && i.x0 > tot.x1 && isNum(i.t))[0]; if (v) res.total = parseNum(v.t); }
    root.__lastAeonItems = items;
    return res;
  }

  root.__lastAeonItems = null;
  const api = { pdfItems, kktcImageParse, parseKktcTokens, aeonPdfParse, snapCatalog, pickBarcode, parseQty, gtinOk, parseNum, parseWords, pdfWords, imageParse, fold };
  if (typeof module !== 'undefined') module.exports = api; else root.FUJI = api;
})(typeof window !== 'undefined' ? window : globalThis);
