(function(){
const {ACCOUNTS, buildForm, fmtDate, assignPO, toISO, addDays, validateFuji} = window.FFM;
const TEMPLATE = Uint8Array.from(atob(TEMPLATE_B64), c => c.charCodeAt(0));
const COMING = [];
// thư viện tải khi cần (chỉ BRG)
const LIBS = Object.assign({ pdf: ['https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js', 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js'], ocr: ['https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js'], tessOpts: { langPath: 'https://cdn.jsdelivr.net/gh/tesseract-ocr/tessdata_fast@main', gzip: false, cachePath: 'ffm-tessdata-fast' } }, window.FFM_LIBS || {});
const loaded = {};
const loadScript = u => loaded[u] || (loaded[u] = new Promise((ok, bad) => { const s = document.createElement('script'); s.src = u; s.onload = ok; s.onerror = () => { delete loaded[u]; bad(new Error('Không tải được ' + u)); }; document.head.append(s); }));
const libs = {
  pdf: async () => { for (const u of LIBS.pdf) await loadScript(u); if (window.pdfjsLib) window.pdfjsLib.GlobalWorkerOptions.workerSrc = LIBS.pdf[1]; },
  ocr: async () => { for (const u of LIBS.ocr) await loadScript(u); },
  tessOpts: LIBS.tessOpts
};
const $ = id => document.getElementById(id);
// thống kê: gửi số liệu (không có nội dung đơn) về Google Sheet "FFM Chuyển Đơn - Thống kê"
const LOG_URL = window.FFM_API || '';
const state = { files: [], results: [], demo: true };
let dl = null;
if (window.claude && window.claude.use) window.claude.use('downloads').then(d => { dl = d; }).catch(() => {});

// account list
const sel = $('acct');
Object.entries(ACCOUNTS).forEach(([k,a]) => sel.add(new Option(a.label, k)));
COMING.forEach(n => { const o = new Option(n + ' · đang làm', ''); o.disabled = true; sel.add(o); });
try { const last = localStorage.getItem('ffm-acct'); if (last && ACCOUNTS[last]) sel.value = last; } catch(e) {}
const {resolveAeon, validateAeon} = window.FFM;
let mapInfo = 'gắn sẵn';
try { const saved = JSON.parse(localStorage.getItem('ffm-aeon-map') || 'null'); if (saved && saved.map) { window.FFM_AEON_MAP_USER = saved.map; mapInfo = `từ file ${saved.name}, nạp ${saved.at}`; } } catch (e) {}
const acctNote = () => {
  const a = ACCOUNTS[sel.value];
  $('acctNote').textContent = 'Nhận file ' + a.accept.replace(/,/g,', ') + ' · DMKH ' + a.customers.length + ' cửa hàng.';
  $('mapBox').hidden = !a.aeon;
  $('nvNote').textContent = a.fixedPo ? 'Account này lấy ngày đặt in trên đơn (ngày giao = ngày đặt + 4). Ô này chỉ dùng khi không đọc được ngày đặt.' : 'Dùng cho số đơn (YYMMDD) và ngày giao = ngày về đơn + 4.';
  if (a.aeon) $('mapNote').textContent = `Bảng Code chuyển đổi: ${Object.keys(window.FFM_AEON_MAP_USER || window.FFM_AEON_MAP || {}).length} mã (${mapInfo}).`;
};
$('mapBtn').addEventListener('click', () => $('mapIn').click());
$('mapIn').addEventListener('change', async () => {
  const f = $('mapIn').files[0]; $('mapIn').value = ''; if (!f) return;
  try {
    const wb = XLSX.read(await f.arrayBuffer(), { type:'array' });
    // sheet có tiêu đề "Barcode Aeon", không có thì sheet đầu tiên
    const ws = wb.SheetNames.map(n => wb.Sheets[n]).find(s => XLSX.utils.sheet_to_json(s, {header:1, raw:false}).some(r => /barcode\s*aeon/i.test(String(r[0] || '')))) || wb.Sheets[wb.SheetNames[0]];
    const map = {};
    XLSX.utils.sheet_to_json(ws, {header:1, raw:false}).forEach(r => { const k = String(r[0] || '').trim(), v = String(r[1] || '').trim(); if (/^\d{6,14}$/.test(k) && v) map[k] = v; });
    if (!Object.keys(map).length) throw new Error('không thấy cột Barcode Aeon / Product code');
    window.FFM_AEON_MAP_USER = map;
    const at = new Date().toLocaleString('vi-VN');
    mapInfo = `từ file ${f.name}, nạp ${at}`;
    try { localStorage.setItem('ffm-aeon-map', JSON.stringify({ map, name: f.name, at })); } catch (e) {}
    state.results.forEach(r => { if (r.aeon) r.lines.forEach(l => { if (!['chosen','manual'].includes(l.pcState)) resolveAeon(l); }); });
    acctNote(); if (!state.demo) render();
    toast(`Đã nạp ${Object.keys(map).length} mã từ ${f.name}`, 'ok');
  } catch (e) { toast('Không đọc được bảng Code chuyển đổi: ' + e.message, 'err'); }
});
$('ngayVe').value = toISO(addDays(new Date(), -1));
$('ngayVe').addEventListener('change', () => { if (!state.demo) { recompute(true); render(); } });
function recompute(resetDelivery){
  const acct = ACCOUNTS[sel.value], nv = $('ngayVe').value;
  assignPO(state.results, nv);
  state.results.forEach(r => { const base = r.orderDate || nv; if (!r.error && (resetDelivery || !r.delivery) && base) r.delivery = toISO(addDays(new Date(base+'T00:00'), acct.leadDays)); });
}
sel.addEventListener('change', () => { try{localStorage.setItem('ffm-acct', sel.value)}catch(e){}; acctNote(); });
acctNote();
const dataNote = () => { if (window.FFM_DATA) $('dataNote').textContent = 'Bảng dữ liệu · ' + window.FFM_DATA.summary(); };
dataNote();
if (window.FFM_DATA) window.FFM_DATA.ready.then(() => {
  if (window.FFM_DATA.status.map && window.FFM_DATA.status.map.src === 'web') { delete window.FFM_AEON_MAP_USER; mapInfo = 'bảng chung trên web'; $('mapBtn').hidden = true; }
  dataNote(); acctNote();
});

// files
const drop = $('drop'), fileIn = $('fileIn');
drop.addEventListener('click', () => fileIn.click());
drop.addEventListener('keydown', e => { if (e.key==='Enter'||e.key===' ') { e.preventDefault(); fileIn.click(); } });
['dragenter','dragover'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave','drop'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', e => addFiles(e.dataTransfer.files));
fileIn.addEventListener('change', () => { addFiles(fileIn.files); fileIn.value=''; });
async function addFiles(fl){
  const list = Array.from(fl);
  for (const f of list) {
    if (f.name.startsWith('~$')) continue; // Excel lock file
    state.files = state.files.filter(x => x.name !== f.name);
    state.files.push({ name: f.name, buf: await f.arrayBuffer() });
  }
  $('fileList').innerHTML = '';
  state.files.forEach(f => { const s = document.createElement('span'); s.className='chip'; s.textContent=f.name; s.title=f.name; $('fileList').appendChild(s); });
  $('btnGo').disabled = !state.files.length;
}

const who = $('who');
try { who.value = localStorage.getItem('ffm-user') || ''; } catch(e) {}
who.addEventListener('change', () => { who.value = who.value.trim(); try { localStorage.setItem('ffm-user', who.value); } catch(e) {} });
function logEvent(d){
  if (!/^https:/.test(LOG_URL)) return;
  try {
    fetch(LOG_URL, { method:'POST', mode:'no-cors', keepalive:true, headers:{'Content-Type':'text/plain;charset=utf-8'},
      body: JSON.stringify(Object.assign({ user: who.value.trim(), ver: ($('ver').textContent || '').replace('Phiên bản ', '') }, d)) }).catch(() => {});
  } catch(e) {}
}

$('btnGo').addEventListener('click', async () => {
  if (!who.value.trim()) { toast('Nhập tên người làm trước khi chuyển đổi (chỉ nhập 1 lần).', 'err'); who.focus(); return; }
  if (window.FFM_DATA) { const t = setTimeout(() => toast('Đang tải bảng DMKH/DMSP mới nhất…'), 300); await window.FFM_DATA.ready; clearTimeout(t); }
  const acct = ACCOUNTS[sel.value], key = sel.value;
  state.demo = false;
  const files = state.files.slice().sort((x,y) => x.name < y.name ? -1 : x.name > y.name ? 1 : 0);
  $('btnGo').disabled = true;
  const out = [];
  for (let i = 0; i < files.length; i++) {
    const f = files[i];
    try {
      let r;
      if (acct.async) r = await acct.parseFile(f, m => toast(`File ${i+1}/${files.length}: ${m}`), libs);
      else { const wb = XLSX.read(f.buf, { type: 'array', cellText: true }); r = acct.parse(wb, f.name); }
      if (!r.error) r.acctKey = key;
      out.push(Object.assign({ fileName: f.name }, r));
    } catch (e) { out.push({ fileName: f.name, error: 'Không đọc được file: ' + (e && e.message || e) }); }
  }
  state.results = out;
  $('btnGo').disabled = false;
  recompute(true);
  render();
  toast('');
  const okN = out.filter(r => !r.error && status(r) !== 'err').length;
  logEvent({ event:'Chuyển đổi', acct:key, files:files.length, orders:out.filter(r => !r.error).length, ok:okN, err:out.length - okN, names:files.map(f => f.name).join(' | ') });
});

const live = r => r.issues.filter(i => !(i.key && i.key !== 'abbrNote' && r[i.key])).concat(r.poNote ? [{level:'warn', msg:r.poNote}] : []);
// lý do chưa cho tải file (hiện cạnh nút tải)
function blockers(r){
  if (r.error) return [r.error];
  const out = [];
  if (!r.po) out.push('chưa có Số đơn (C9)');
  if (!r.customer) out.push('chưa chọn Mã khách hàng (C10)');
  if (!r.delivery) out.push('chưa có Ngày giao (C13)');
  if (!r.fixedPo && !$('ngayVe').value) out.push('chưa chọn Ngày về đơn');
  if (!r.lines.length) out.push('không có dòng hàng nào');
  live(r).filter(i => i.level === 'err' && !['po','customer'].includes(i.key)).forEach(i => out.push(i.msg));
  const bad = r.lines.filter(l => l.flags.some(x => x.level === 'err')).map(l => l.row);
  if (bad.length) out.push('dòng ' + bad.join(', ') + ' còn lỗi đỏ');
  return out;
}
function status(r){
  if (r.error) return 'err';
  const all = live(r).concat(...r.lines.map(l => l.flags));
  if (blockers(r).length) return 'err';
  return all.some(i => i.level==='warn') ? 'warn' : 'ok';
}
const PILL = { ok:'Sẵn sàng', warn:'Xem lại cảnh báo', err:'Cần sửa' };
const outName = r => (r.po || 'CHUA_CO_SO_DON') + '.xlsx';

function el(tag, attrs, ...kids){ const e=document.createElement(tag); for (const k in attrs||{}) { if(k==='class') e.className=attrs[k]; else if(k==='text') e.textContent=attrs[k]; else e.setAttribute(k,attrs[k]); } kids.flat().forEach(c => c!=null && e.append(c)); return e; }

function render(){
  const box = $('results'); box.innerHTML='';
  const rs = state.demo ? [DEMO] : state.results;
  rs.forEach((r,idx) => box.appendChild(card(r, idx)));
  if (!state.demo) {
    const n = rs.length, ok = rs.filter(r => status(r)!=='err').length;
    const lines = rs.reduce((s,r)=>s+(r.lines?r.lines.length:0),0), qty = rs.reduce((s,r)=>s+(r.lines?r.lines.reduce((a,l)=>a+(+l.qty||0),0):0),0);
    $('summary').textContent = `${n} đơn · ${lines} dòng hàng · tổng ${qty.toLocaleString('vi-VN')} SL · ${ok}/${n} đơn tải được`;
    $('btnAll').hidden = n < 2; $('btnAll').disabled = ok === 0;
  }
}

function editCell(r, l, key){
  const inp = el('input', {type:'text', class:'cell', value: l[key] == null ? '' : String(l[key]), 'aria-label': key === 'qty' ? 'Số lượng' : 'Barcode', inputmode:'numeric'});
  const commit = () => { const v = inp.value.replace(/\s/g, ''); l[key] = key === 'qty' ? (v === '' ? null : Number(v.replace(',', '.'))) : v.replace(/\D/g, ''); if (key === 'barcode') l.bcState = 'sure'; validateFuji(l); };
  inp.addEventListener('change', () => { commit(); render(); });
  inp.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); inp.blur(); } });
  // kiểm tra ngay khi gõ: viền xanh = barcode hợp lệ, đỏ = chưa đúng
  if (key === 'barcode') inp.addEventListener('input', () => { const d = inp.value.replace(/\D/g, ''); inp.classList.toggle('ok', window.FUJI.gtinOk(d)); inp.classList.toggle('bad', d.length >= 8 && !window.FUJI.gtinOk(d)); });
  return inp;
}

const flagEls = l => l.flags.map(x => el('span',{class:'flag'+(x.level==='info'?' info':''),text:(x.level==='info'?'ⓘ ':'⚠ ')+x.msg}));
// AEON: barcode trên đơn -> product code (cột D)
function saveAeonCode(l){
  if (!who.value.trim()) { toast('Nhập tên người làm trước.', 'err'); who.focus(); return; }
  const bc = l.barcode, pc = l.productCode;
  window.FFM_AEON_MAP_SHARED = Object.assign({}, window.FFM_AEON_MAP_SHARED || {}, { [bc]: pc });
  try { localStorage.setItem('ffm-data-v1:aeonShared', JSON.stringify(window.FFM_AEON_MAP_SHARED)); } catch(e) {}
  logEvent({ event:'Thêm code AEON', acct:'AEON', barcode:bc, code:pc, aeonName:l.aeonName || '' });
  // áp luôn cho các dòng cùng barcode đang mở
  state.results.forEach(r => { if (r.aeon) r.lines.forEach(x => { if (x.barcode === bc) { x.productCode = pc; x.pcState = 'map'; validateAeon(x); } }); });
  if (window.FFM_DATA && window.FFM_DATA.status.shared) window.FFM_DATA.status.shared.n = Object.keys(window.FFM_AEON_MAP_SHARED).length;
  dataNote(); render();
  toast(`Đã lưu ${bc} → ${pc} vào bảng chung. Cả team dùng được từ lần mở web tiếp theo.`, 'ok');
}
function aeonTable(r){
  const tb = el('tbody');
  r.lines.forEach((l, i) => {
    const lvl = l.flags.some(x=>x.level==='err') ? 'err' : l.flags.some(x=>x.level==='warn') ? 'warn' : '';
    let pcCtl;
    if (l.cands && l.cands.length > 1) {
      pcCtl = el('select', {class:'cell', 'aria-label':'Product code dòng ' + (i+1)});
      pcCtl.add(new Option('— chọn mã —', ''));
      l.cands.forEach(cd => pcCtl.add(new Option(`${cd.code} · ${cd.name}${cd.active ? '' : ' (ngừng)'}`, cd.code)));
      if (l.productCode && !l.cands.some(cd => cd.code === l.productCode)) pcCtl.add(new Option(l.productCode + ' · (bảng Code chuyển đổi)', l.productCode));
      pcCtl.value = l.productCode || '';
      pcCtl.addEventListener('change', () => { l.productCode = pcCtl.value; l.pcState = 'chosen'; validateAeon(l); render(); });
    } else {
      pcCtl = el('input', {type:'text', class:'cell', value:l.productCode || '', 'aria-label':'Product code dòng ' + (i+1)});
      pcCtl.addEventListener('change', () => { l.productCode = pcCtl.value.trim().toUpperCase(); l.pcState = 'manual'; validateAeon(l); render(); });
      pcCtl.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); pcCtl.blur(); } });
    }
    const shared = window.FFM_AEON_MAP_SHARED || {};
    const official = window.FFM_AEON_MAP || {};
    const canSave = ['chosen','manual'].includes(l.pcState) && l.productCode && l.barcode && !official[l.barcode] && shared[l.barcode] !== l.productCode;
    const inDmsp = l.productCode && window.FFM_PC && Object.values(window.FFM_PC).some(a => a.some(x => x[0] === l.productCode));
    let saveBtn = null;
    if (canSave) {
      saveBtn = el('button', {class:'small', type:'button', text: inDmsp ? '+ Lưu vào bảng Code chuyển đổi chung' : 'Mã này không có trong DMSP, chưa lưu được'});
      saveBtn.disabled = !inDmsp;
      saveBtn.title = 'Lần sau cả team gặp barcode ' + l.barcode + ' sẽ tự điền ' + l.productCode;
      saveBtn.addEventListener('click', () => saveAeonCode(l));
    }
    const src = { map: (shared[l.barcode] && !official[l.barcode] ? 'theo bảng Code chuyển đổi (team thêm)' : 'theo bảng Code chuyển đổi'), dmsp:'theo DMSP', multi:'mã gốc trong DMSP', chosen:'chọn tay', manual:'nhập tay', pick:'', none:'' }[l.pcState] || '';
    tb.append(el('tr', {class:lvl},
      el('td', {class:'num', text:String(i+1)}),
      el('td', {}, el('div', {class:'mono', text:l.barcode || '—'}), el('div', {class:'sub2', text:l.aeonName})),
      el('td', {}, pcCtl, el('div', {class:'pname', text:(l.name || '') + (src ? ' · ' + src : '')}), saveBtn, flagEls(l)),
      el('td', {class:'num', text: l.qty == null ? '?' : l.qty.toLocaleString('vi-VN')}),
      el('td', {class:'num', text: l.pack != null && l.cartons != null ? `${l.pack} × ${l.cartons}` : ''})));
  });
  const total = r.lines.reduce((s,l)=>s+(+l.qty||0),0);
  return el('div',{class:'tbl'}, el('table',{},
    el('thead',{}, el('tr',{}, el('th',{class:'num',text:'#'}), el('th',{text:'Barcode & tên trên đơn AEON'}), el('th',{text:'Product code → cột D'}), el('th',{class:'num',text:'SL → cột F'}), el('th',{class:'num',text:'Quy cách × thùng'}))),
    tb,
    el('tfoot',{}, el('tr',{}, el('td',{}), el('td',{text:r.lines.length+' dòng'}), el('td',{}), el('td',{class:'num',text:total.toLocaleString('vi-VN')}), el('td',{})))));
}

function card(r, idx){
  const st = r.demo ? 'demo' : status(r);
  const c = el('article', {class:'card'});
  // nút tải + lý do + cảnh báo đặt ngay đầu thẻ, khỏi phải cuộn xuống
  const btn = el('button', {class:'primary', text:'Tải file Excel'});
  btn.disabled = r.demo || st === 'err' || !!r.error;
  btn.addEventListener('click', async () => {
    btn.disabled = true;
    try { const bytes = await buildForm(window.FFM_TPL || TEMPLATE, r); if (await save(outName(r), new Blob([bytes]))) logEvent({ event:'Tải file', acct:r.acctKey || sel.value, orders:1, ok:1, names:outName(r) }); }
    catch(e) { toast(e.message, 'err'); }
    btn.disabled = false;
  });
  const dateTxt = r.delivery ? fmtDate(new Date(r.delivery+'T00:00')) : '—';
  const h = el('div', {class:'card-h'},
    el('div', {style:'min-width:0'}, el('h2', {text: r.po || '(chưa có số đơn)'}), el('div', {class:'src', text: 'Từ file: ' + r.fileName + (r.error ? '' : ` · → ${outName(r)} · giao ${dateTxt}`)})),
    el('div', {class:'card-act'}, el('span', {class:'pill '+st, text: r.demo ? 'Ví dụ' : PILL[st]}), r.error ? null : btn));
  c.append(h);
  if (r.error) { c.append(el('ul',{class:'issues',style:'margin-top:12px'}, el('li',{class:'err',text:r.error}))); return c; }
  if (!r.demo && st === 'err') c.append(el('div', {class:'why'}, el('b', {text:'Chưa tải được vì: '}), blockers(r).join(' · ')));
  {
    const ul0 = el('ul', {class:'issues top'});
    live(r).filter(i => i.msg).forEach(i => ul0.append(el('li', {class:i.level, text:i.msg})));
    const warnRows = r.lines.filter(l => l.flags.some(x => x.level === 'warn')).map(l => l.row);
    if (warnRows.length) ul0.append(el('li', {class:'warn', text:`Dòng ${warnRows.join(', ')} có cảnh báo vàng, xem trong bảng bên dưới.`}));
    if (ul0.children.length) c.append(ul0);
  }

  const f = el('div', {class:'fields'});
  const custSelect = () => {
    const acct = ACCOUNTS[r.acctKey || sel.value], id = `f-${idx}-customer`;
    const s = el('select', {id});
    s.add(new Option('— chọn cửa hàng —', ''));
    acct.customers.forEach(c => s.add(new Option(`${c.code} · ${c.name}`, c.code)));
    if (r.customer && !acct.customers.some(c => c.code === r.customer)) s.add(new Option(r.customer + ' · (ngoài DMKH)', r.customer));
    s.value = r.customer || '';
    if (r.demo) s.disabled = true;
    s.addEventListener('change', () => { r.customer = s.value; const c = acct.customers.find(x => x.code === s.value); if (c && acct.makeAbbr && (!r.abbr || (acct.abbr && acct.abbr[c.code]))) r.abbr = acct.makeAbbr(c.name, c.code).abbr; recompute(false); render(); });
    return el('div', {}, el('label', {class:'k', for:id, text:'Mã khách hàng (C10)'}), s);
  };
  const field = (label, key, type) => {
    const id = `f-${idx}-${key}`;
    const inp = el('input', {type, id, value: r[key] || ''});
    if (r.demo) inp.disabled = true;
    inp.addEventListener('change', () => { r[key] = type==='text' ? inp.value.trim().toUpperCase() : inp.value; if (key==='abbr') recompute(false); render(); });
    return el('div', {}, el('label', {class:'k', for:id, text:label}), inp);
  };
  f.append(el('div', {}, el('label',{class:'k',text: r.aeon ? 'Cửa hàng (trên đơn)' : r.fixedPo ? 'Nơi nhận (trên phiếu)' : 'Cửa hàng (ô tiêu đề)'}), el('div',{class:'ro',text:r.store || '—'})),
    r.fixedPo ? field('Số đơn (C9)', 'po', 'text') : field('Viết tắt cửa hàng', 'abbr', 'text'), custSelect(), field('Ngày giao yêu cầu (C13)', 'delivery', 'date'));
  c.append(f);


  if (r.aeon) { c.append(aeonTable(r)); }
  const tb = el('tbody');
  if (!r.aeon) r.lines.forEach((l,i) => {
    const lvl = l.flags.some(x=>x.level==='err') ? 'err' : l.flags.some(x=>x.level==='warn') ? 'warn' : '';
    tb.append(el('tr', {class:lvl},
      el('td', {class:'num', text:String(i+1)}),
      el('td', {class:'mono'}, r.editable ? editCell(r, l, 'barcode') : (l.barcode || '—')),
      el('td', {}, r.editable && l.name ? el('div', {class:'pname', text:l.name}) : null, l.preview ? el('img', {class:'rowimg', src:l.preview, alt:'Ảnh dòng ' + l.row + ' trên phiếu'}) : null, r.editable ? (l.price != null ? `${l.qty == null ? '?' : l.qty.toLocaleString('vi-VN')} × ${l.price.toLocaleString('vi-VN')} = ${l.amount != null ? l.amount.toLocaleString('vi-VN') : '?'}` : '') : (l.name || ''), l.flags.map(x => el('span',{class:'flag'+(x.level==='info'?' info':''),text:(x.level==='info'?'ⓘ ':'⚠ ')+x.msg}))),
      el('td', {class:'num'}, r.editable ? editCell(r, l, 'qty') : (Number.isNaN(l.qty) ? '?' : l.qty.toLocaleString('vi-VN'))),
      el('td', {class:'num', text:String(l.row)})));
  });
  const total = r.lines.reduce((a,l)=>a+(+l.qty||0),0);
  if (!r.aeon) c.append(el('div',{class:'tbl'}, el('table',{},
    el('thead',{}, el('tr',{}, el('th',{class:'num',text:'#'}), el('th',{text:'Barcode → cột C'}), el('th',{text: r.editable ? 'Tên SP (DMSP) · ảnh dòng · SL × đơn giá = thành tiền' : 'Tên sản phẩm (đối chiếu)'}), el('th',{class:'num',text:'SL → cột F'}), el('th',{class:'num',text: r.editable ? 'STT phiếu' : 'Dòng gốc'}))),
    tb,
    el('tfoot',{}, el('tr',{}, el('td',{}), el('td',{text:r.lines.length+' dòng'}), el('td',{}), el('td',{class:'num',text:total.toLocaleString('vi-VN')}), el('td',{}))))));

  return c;
}

$('btnAll').addEventListener('click', async () => {
  const ready = state.results.filter(r => status(r) !== 'err');
  const zip = new JSZip();
  for (const r of ready) zip.file(outName(r), await buildForm(window.FFM_TPL || TEMPLATE, r));
  const blob = await zip.generateAsync({type:'blob'});
  const d = new Date(), p = n => String(n).padStart(2,'0');
  if (await save(`${sel.value}_${d.getFullYear()}${p(d.getMonth()+1)}${p(d.getDate())}.zip`, blob)) logEvent({ event:'Tải tất cả (zip)', acct:sel.value, orders:ready.length, ok:ready.length, names:ready.map(outName).join(' | ') });
});

async function save(filename, blob){
  if (dl) {
    try { await dl.save({filename, data: blob}); toast('Đã tải ' + filename, 'ok'); return true; }
    catch(e) { toast(e && e.code === 'declined' ? 'Đã huỷ tải.' : 'Không tải được: ' + (e && e.message || e), 'err'); }
    return false;
  }
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = filename; document.body.append(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  toast('Đã tải ' + filename, 'ok');
  return true;
}
function toast(msg, kind){ const t=$('toast'); t.textContent=msg; t.className='toast '+(kind||''); }

const DEMO = { demo:true, fileName:'2026.11.11 - Emeral BK - KKTCEMERAL261110-01.xlsx', po:'KKTCEMERAL261110-01', abbr:'EMERAL', customer:'KKTC0041', store:'Kmart The Emerald', delivery:'2026-11-14', orderDate:'2026-11-11', issues:[],
  lines:[{row:54,barcode:'8001200045053',name:'Mì Ý Spaghettini Agnesi 500g',qty:6,flags:[]},{row:55,barcode:'8001200045091',name:'Mì Ý Linguine Agnesi 500g',qty:12,flags:[]},{row:58,barcode:'8001200023808',name:'Nui Xoắn Rau Củ 3 Màu Eliche Tricolori Agnesi 500g',qty:3,flags:[]}] };
render();
})();
