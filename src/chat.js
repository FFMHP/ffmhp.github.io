/* Trang chủ (2 nhóm) + chat theo phòng.
   #/          -> trang chủ, phòng "tong" (chat chung)
   #/don-hang  -> công cụ chuyển đơn, phòng "donhang"
   #/khac      -> team khác (đang xây dựng), phòng "khac"
   Tin nhắn lưu ở sheet "Chat" (cột Phòng). Hỏi tin mới 15 giây/lần khi mở khung, 60 giây/lần khi đóng. */
(function () {
  const $ = id => document.getElementById(id);
  const ROUTES = {
    '': { page: 'pageHome', room: 'tong', title: 'Chat chung FFM', doc: 'Công cụ team FFM' },
    'don-hang': { page: 'pageOrders', room: 'donhang', title: 'Chat team đơn hàng', doc: 'FFM Chuyển Đơn' },
    'khac': { page: 'pageOther', room: 'khac', title: 'Chat team khác', doc: 'FFM · Team khác' }
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

  // ---- chat ----
  const fmt = t => { const d = new Date(t), p = n => String(n).padStart(2, '0'); const today = new Date().toDateString() === d.toDateString(); return (today ? '' : `${p(d.getDate())}/${p(d.getMonth() + 1)} `) + `${p(d.getHours())}:${p(d.getMinutes())}`; };
  function draw() {
    const atBottom = list.scrollHeight - list.scrollTop - list.clientHeight < 40;
    list.innerHTML = '';
    if (!msgs.length) { const e = document.createElement('p'); e.className = 'chat-empty'; e.textContent = room === 'tong' ? 'Chưa có tin nhắn chung. Nhắn gì cho cả phòng FFM cứ gõ ở đây.' : 'Chưa có tin nhắn trong nhóm này.'; list.append(e); }
    msgs.forEach(m => {
      const row = document.createElement('div'); row.className = 'chat-msg' + (m.u === me() ? ' mine' : '');
      const h = document.createElement('div'); h.className = 'chat-meta'; h.textContent = `${m.u || '?'} · ${fmt(m.t)}`;
      const b = document.createElement('div'); b.className = 'chat-text'; b.textContent = m.m;
      row.append(h, b); list.append(row);
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
    const text = inp.value.trim(); if (!text) return;
    if (!me()) { $('chatState').textContent = 'Nhập tên bạn trước khi nhắn.'; syncWho(); chatWho.focus(); return; }
    inp.value = ''; inp.disabled = true;
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
