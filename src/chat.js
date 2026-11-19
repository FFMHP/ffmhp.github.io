/* Trang chủ (4 nhóm) + chat theo phòng.
   #/             -> trang chủ, phòng "tong" (chat chung)
   #/don-hang     -> Team đơn hàng MT (chuyển đơn FORM_2_HP), phòng "donhang"
   #/don-hang-gt  -> Team đơn hàng GT (Chuyển đổi file up DMS), phòng "khac"
   #/bao-cao      -> Team báo cáo (Báo cáo List order), phòng "khac"
   #/khac         -> Team khác (tính năng đang xây dựng), phòng "khac" (giữ lịch sử cũ; Apps Script chưa có phòng gt/baocao, phòng lạ bị đẩy về "tong")
   Tin nhắn lưu ở sheet "Chat" (cột Phòng). Hỏi tin mới 15 giây/lần khi mở khung, 60 giây/lần khi đóng. */
(function () {
  const $ = id => document.getElementById(id);
  const ROUTES = {
    '': { page: 'pageHome', room: 'tong', title: 'Chat chung FFM', doc: 'Công cụ team FFM' },
    'don-hang': { page: 'pageOrders', room: 'donhang', title: 'Chat team đơn hàng MT', doc: 'FFM Chuyển Đơn · Đơn hàng MT' },
    'don-hang-gt': { page: 'pageGT', room: 'khac', title: 'Chat team GT · Báo cáo · Khác', doc: 'FFM · Đơn hàng GT' },
    'bao-cao': { page: 'pageReport', room: 'khac', title: 'Chat team GT · Báo cáo · Khác', doc: 'FFM · Team báo cáo' },
    'khac': { page: 'pageOther', room: 'khac', title: 'Chat team GT · Báo cáo · Khác', doc: 'FFM · Team khác' },
    'sinh-hoat': { page: 'pageSocial', room: 'tong', title: 'Chat chung FFM', doc: 'FFM · Sinh hoạt chung' }
  };
  const API = window.FFM_API || '';
  const box = $('chatBox'), list = $('chatList'), inp = $('chatIn'), btn = $('chatBtn'), badge = $('chatBadge');
  const who = $('who'), chatWho = $('chatWho');
  let room = 'tong', msgs = [], last = 0, open = false, timer = null, busy = false, gen = 0;
  const seenKey = () => 'ffm-chat-seen-' + room;
  let seen = 0;
  const me = () => (who.value || '').trim();

  // ---- trang ----
  function route() {
    const key = (location.hash.replace(/^#\/?/, '').split(/[/?]/)[0] || '').toLowerCase();
    const R = ROUTES[key] || ROUTES[''];
    Object.values(ROUTES).forEach(x => { $(x.page).hidden = x.page !== R.page; });
    document.querySelectorAll('.nav a').forEach(a => { const on = a.dataset.route === (ROUTES[key] ? key : ''); a.classList.toggle('on', on); if (on) a.setAttribute('aria-current', 'page'); else a.removeAttribute('aria-current'); });
    if (R.page === 'pageHome') loadStats();
    if (window.FFM_social) { if (R.page === 'pageSocial') window.FFM_social.show(); else window.FFM_social.hide(); }
    document.title = R.doc;
    window.scrollTo(0, 0);
    setRoom(R.room, R.title);
  }
  window.addEventListener('hashchange', route);

  // ---- tên người dùng: nút trên thanh đầu trang (dùng chung cho thống kê + chat) ----
  const pop = $('userPop'), uIn = $('userIn');
  function setName(v) {
    who.value = (v || '').trim();
    try { localStorage.setItem('ffm-user', who.value); } catch (e) {}
    who.dispatchEvent(new Event('change'));
    syncWho(); draw();
  }
  function syncWho() {
    chatWho.value = me();
    $('chatWhoBox').hidden = !!me();
    $('userName').textContent = me() || 'Nhập tên';
    $('userAv').textContent = me() ? me().split(/\s+/).pop().charAt(0).toUpperCase() : '?';
    $('userBtn').classList.toggle('need', !me());
  }
  function openPop(v) { pop.hidden = !v; $('userBtn').setAttribute('aria-expanded', String(v)); if (v) { uIn.value = me(); uIn.focus(); uIn.select(); } }
  $('userBtn').addEventListener('click', () => openPop(pop.hidden));
  $('userSave').addEventListener('click', () => { if (uIn.value.trim()) { setName(uIn.value); openPop(false); } else uIn.focus(); });
  uIn.addEventListener('keydown', e => { if (e.key === 'Enter') $('userSave').click(); if (e.key === 'Escape') openPop(false); });
  document.addEventListener('click', e => { if (!pop.hidden && !e.target.closest('.user')) openPop(false); });
  window.FFM_askName = () => { setTimeout(() => openPop(true), 0); window.scrollTo(0, 0); };

  // ---- góp ý: mở khung chat ----
  $('fbBtn').addEventListener('click', () => { if (!open) btn.click(); if (!inp.value) inp.value = '[Góp ý] '; inp.focus(); });

  // ---- trang chủ: số liệu hôm nay từ sheet Log ----
  let statsAt = 0;
  async function loadStats() {
    if (!API || Date.now() - statsAt < 30000) return;
    statsAt = Date.now();
    try {
      const j = await (await fetch(API + '?t=stats', { cache: 'no-store' })).json();
      const tm = j.teams || {};
      $('stDon').textContent = (tm.donhang || 0).toLocaleString('vi-VN');
      $('stKhac').textContent = (tm.khac || 0).toLocaleString('vi-VN');
      $('statsNote').textContent = 'cập nhật ' + fmt(Date.now());
    } catch (e) { $('statsNote').textContent = 'chưa tải được số liệu'; }
  }
  chatWho.addEventListener('change', () => setName(chatWho.value));
  who.addEventListener('change', syncWho);


  // ---- file đính kèm (bật mặc định; tắt bằng window.FFM_CHAT_ATTACH=false) ----
  // tin nhắn lưu dạng "nội dung\n[[f:<id Drive>|<tên file>]]" (mỗi file 1 dòng)
  const ATTACH = window.FFM_CHAT_ATTACH !== false && /^https:/.test(API);
  const FILE_RE = /\[\[f:([\w-]{10,})\|([^\]\n]{0,120})\]\]/g;
  const isImg = n => /\.(jpe?g|png|gif|webp|heic)$/i.test(n);
  const icon = n => /\.pdf$/i.test(n) ? '📕' : /\.(xlsx?|xlsm|csv)$/i.test(n) ? '📗' : /\.docx?$/i.test(n) ? '📘' : isImg(n) ? '🖼️' : '📄';
  function renderMsg(row, text) {
    const files = []; let m; FILE_RE.lastIndex = 0;
    while ((m = FILE_RE.exec(text))) files.push({ id: m[1], name: m[2] || 'file' });
    const plain = text.replace(FILE_RE, '').replace(/\n{2,}/g, '\n').trim();
    if (plain || !files.length) { const b = document.createElement('div'); b.className = 'chat-text'; b.textContent = plain; row.append(b); }
    files.forEach(f => {
      const a = document.createElement('a'); a.className = 'chat-file' + (isImg(f.name) ? ' img' : '');
      a.href = 'https://drive.google.com/file/d/' + encodeURIComponent(f.id) + '/view'; a.target = '_blank'; a.rel = 'noopener';
      a.title = 'Mở / tải ' + f.name;
      if (isImg(f.name)) { const im = document.createElement('img'); im.loading = 'lazy'; im.alt = f.name; im.src = 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(f.id) + '&sz=w440'; a.append(im); }
      const t = document.createElement('span'); t.append(icon(f.name) + ' '); const bn = document.createElement('b'); bn.textContent = f.name; t.append(bn); a.append(t);
      const w = document.createElement('div'); w.append(a); row.append(w);
    });
  }
  let pend = [];
  const MAXB = 20 * 1024 * 1024;
  const MIME = { pdf: 'application/pdf', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', xlsm: 'application/vnd.ms-excel', xls: 'application/vnd.ms-excel', csv: 'text/csv', docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', doc: 'application/msword', txt: 'text/plain' };
  const extOf = n => ((n.match(/\.(\w+)$/) || [])[1] || '').toLowerCase();
  function drawPend() {
    const el = $('chatFiles'); el.innerHTML = ''; el.hidden = !pend.length;
    pend.forEach((f, i) => {
      const c = document.createElement('span'); c.className = 'chat-chip';
      const s = document.createElement('span'); s.textContent = icon(f.name) + ' ' + f.name; c.append(s);
      const x = document.createElement('button'); x.type = 'button'; x.textContent = '×'; x.title = 'Bỏ file này'; x.onclick = () => { pend.splice(i, 1); drawPend(); }; c.append(x);
      el.append(c);
    });
  }
  function addPend(list) {
    for (const f of Array.from(list || [])) {
      const e = extOf(f.name);
      if (!MIME[e] && !/^image\//.test(f.type)) { $('chatState').textContent = `Chưa nhận loại file "${f.name}". Gửi PDF, Excel, Word hoặc ảnh nha.`; continue; }
      if (f.size > MAXB) { $('chatState').textContent = `File "${f.name}" quá lớn (tối đa 20 MB).`; continue; }
      if (pend.length >= 5) { $('chatState').textContent = 'Mỗi lần gửi tối đa 5 file.'; break; }
      pend.push(f);
    }
    drawPend();
  }
  const toB64 = blob => new Promise((ok, no) => { const r = new FileReader(); r.onload = () => ok(String(r.result).split(',')[1]); r.onerror = no; r.readAsDataURL(blob); });
  async function uploadPend() {
    const out = [];
    for (let i = 0; i < pend.length; i++) {
      const f = pend[i];
      $('chatState').textContent = `Đang tải lên ${i + 1}/${pend.length}: ${f.name}…`;
      const mime = MIME[extOf(f.name)] || f.type || 'application/octet-stream';
      const r = await fetch(API, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ event: 'upload', kind: 'file', user: me(), name: f.name, mime, data: await toB64(f) }) });
      const j = await r.json(); if (j.error || !j.id) throw new Error(j.error || 'không nhận được file');
      out.push(`[[f:${j.id}|${f.name.replace(/[\]|\n]/g, ' ').slice(0, 100)}]]`);
    }
    $('chatState').textContent = '';
    return out;
  }
  if (ATTACH) {
    $('chatAtt').hidden = false;
    $('chatAtt').addEventListener('click', () => $('chatFile').click());
    $('chatFile').addEventListener('change', e => { addPend(e.target.files); e.target.value = ''; });
    box.addEventListener('dragover', e => { if (e.dataTransfer && [...e.dataTransfer.types].includes('Files')) { e.preventDefault(); box.classList.add('over'); } });
    box.addEventListener('dragleave', e => { if (!box.contains(e.relatedTarget)) box.classList.remove('over'); });
    box.addEventListener('drop', e => { e.preventDefault(); box.classList.remove('over'); addPend(e.dataTransfer.files); });
    inp.addEventListener('paste', e => { const fs = e.clipboardData && e.clipboardData.files; if (fs && fs.length) { e.preventDefault(); addPend(fs); } });
  }
  // ---- chat ----
  const fmt = t => { const d = new Date(t), p = n => String(n).padStart(2, '0'); const today = new Date().toDateString() === d.toDateString(); return (today ? '' : `${p(d.getDate())}/${p(d.getMonth() + 1)} `) + `${p(d.getHours())}:${p(d.getMinutes())}`; };
  function draw() {
    const atBottom = list.scrollHeight - list.scrollTop - list.clientHeight < 40;
    list.innerHTML = '';
    if (!msgs.length) { const e = document.createElement('p'); e.className = 'chat-empty'; e.textContent = room === 'tong' ? 'Chưa có tin nhắn chung. Nhắn gì cho cả phòng FFM cứ gõ ở đây.' : 'Chưa có tin nhắn trong nhóm này.'; list.append(e); }
    msgs.forEach(m => {
      const row = document.createElement('div'); row.className = 'chat-msg' + (m.u === me() ? ' mine' : '');
      const h = document.createElement('div'); h.className = 'chat-meta'; h.textContent = `${m.u || '?'} · ${fmt(m.t)}`;
      row.append(h); renderMsg(row, m.m); list.append(row);
    });
    if (atBottom || open) list.scrollTop = list.scrollHeight;
    const unread = msgs.filter(m => m.t > seen && m.u !== me()).length;
    badge.textContent = unread > 9 ? '9+' : String(unread); badge.hidden = open || !unread;
  }
  async function poll() {
    if (!API || busy || document.hidden) return;
    busy = true; const g = gen, r0 = room;
    try {
      const r = await fetch(`${API}?t=chat&room=${r0}&since=${last}`, { cache: 'no-store' });
      const j = await r.json();
      if (g !== gen) return; // đã đổi phòng trong lúc chờ
      if (j && j.msgs && j.msgs.length) {
        const keys = new Set(msgs.map(m => m.t + m.u + m.m));
        j.msgs.forEach(m => { if (!keys.has(m.t + m.u + m.m)) msgs.push(m); });
        msgs.sort((a, b) => a.t - b.t); msgs = msgs.slice(-100);
        last = msgs[msgs.length - 1].t;
      }
      draw();
      $('chatState').textContent = '';
    } catch (e) { if (g === gen) $('chatState').textContent = 'Mất kết nối, đang thử lại…'; }
    finally { busy = false; }
  }
  function schedule() { clearInterval(timer); timer = setInterval(poll, open ? 15000 : 60000); }
  function markSeen() { if (msgs.length) { seen = msgs[msgs.length - 1].t; try { localStorage.setItem(seenKey(), String(seen)); } catch (e) {} } draw(); }
  function setRoom(r, title) {
    if (r === room && msgs.length) return;
    room = r; gen++; msgs = []; last = 0; busy = false;
    try { seen = +localStorage.getItem(seenKey()) || 0; } catch (e) { seen = 0; }
    $('chatTitle').textContent = title;
    btn.firstChild.textContent = `💬 ${title} `;
    draw();
    poll().then(() => { if (open) markSeen(); });
  }
  btn.addEventListener('click', () => {
    open = !open; box.hidden = !open; btn.setAttribute('aria-expanded', String(open));
    if (open) { syncWho(); poll().then(markSeen); (me() ? inp : chatWho).focus(); }
    schedule();
  });
  $('chatClose').addEventListener('click', () => btn.click());
  async function send() {
    let text = inp.value.trim(); if (!text && !pend.length) return;
    if (!me()) { $('chatState').textContent = 'Nhập tên bạn trước khi nhắn.'; syncWho(); chatWho.focus(); return; }
    inp.value = ''; inp.disabled = true;
    if (pend.length) {
      $('chatSend').disabled = true;
      try { const refs = await uploadPend(); text = (text.slice(0, 480 - refs.join('\n').length) + '\n' + refs.join('\n')).trim(); pend = []; drawPend(); }
      catch (e) { $('chatState').textContent = 'Chưa tải file lên được: ' + (e.message || e) + '. Bấm Gửi lại.'; inp.value = text; inp.disabled = false; $('chatSend').disabled = false; return; }
      finally { $('chatSend').disabled = false; }
    }
    const local = { t: Date.now(), u: me(), m: text };
    msgs.push(local); draw(); list.scrollTop = list.scrollHeight;
    try {
      await fetch(API, { method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ event: 'chat', room, user: me(), msg: text }) });
      msgs = msgs.filter(m => m !== local);
      await new Promise(r => setTimeout(r, 800)); await poll(); markSeen();
    } catch (e) { $('chatState').textContent = 'Chưa gửi được, kiểm tra mạng rồi gửi lại.'; inp.value = text; msgs = msgs.filter(m => m !== local); draw(); }
    finally { inp.disabled = false; inp.focus(); }
  }
  $('chatSend').addEventListener('click', send);
  inp.addEventListener('keydown', e => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); send(); } });
  document.addEventListener('visibilitychange', () => { if (!document.hidden) poll(); });

  if (!/^https:/.test(API)) btn.hidden = true;
  syncWho();
  route();
  schedule();
})();
