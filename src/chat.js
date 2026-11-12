/* Khung chat team: tin nhắn lưu ở sheet "Chat" (Google Sheet thống kê). Hỏi tin mới 15 giây/lần khi mở khung, 60 giây/lần khi đóng. */
(function () {
  const $ = id => document.getElementById(id);
  const API = window.FFM_API || '';
  if (!/^https:/.test(API)) { const b = $('chatBtn'); if (b) b.hidden = true; return; }
  const box = $('chatBox'), list = $('chatList'), inp = $('chatIn'), btn = $('chatBtn'), badge = $('chatBadge');
  let msgs = [], last = 0, open = false, timer = null, busy = false;
  const LS = 'ffm-chat-seen';
  let seen = 0; try { seen = +localStorage.getItem(LS) || 0; } catch (e) {}
  const me = () => ($('who').value || '').trim();
  const fmt = t => { const d = new Date(t), p = n => String(n).padStart(2, '0'); const today = new Date().toDateString() === d.toDateString(); return (today ? '' : `${p(d.getDate())}/${p(d.getMonth() + 1)} `) + `${p(d.getHours())}:${p(d.getMinutes())}`; };
  function draw() {
    const atBottom = list.scrollHeight - list.scrollTop - list.clientHeight < 40;
    list.innerHTML = '';
    if (!msgs.length) { const e = document.createElement('p'); e.className = 'chat-empty'; e.textContent = 'Chưa có tin nhắn. Hỏi gì team cứ nhắn ở đây.'; list.append(e); }
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
    if (busy || document.hidden) return;
    busy = true;
    try {
      const r = await fetch(API + '?t=chat&since=' + last, { cache: 'no-store' });
      const j = await r.json();
      if (j && j.msgs && j.msgs.length) {
        const keys = new Set(msgs.map(m => m.t + m.u + m.m));
        j.msgs.forEach(m => { if (!keys.has(m.t + m.u + m.m)) msgs.push(m); });
        msgs.sort((a, b) => a.t - b.t); msgs = msgs.slice(-100);
        last = msgs[msgs.length - 1].t; draw();
      }
      $('chatState').textContent = '';
    } catch (e) { $('chatState').textContent = 'Mất kết nối, đang thử lại…'; }
    finally { busy = false; }
  }
  function schedule() { clearInterval(timer); timer = setInterval(poll, open ? 15000 : 60000); }
  function markSeen() { if (msgs.length) { seen = msgs[msgs.length - 1].t; try { localStorage.setItem(LS, String(seen)); } catch (e) {} } draw(); }
  btn.addEventListener('click', () => {
    open = !open; box.hidden = !open; btn.setAttribute('aria-expanded', String(open));
    if (open) { poll().then(markSeen); inp.focus(); } schedule();
  });
  $('chatClose').addEventListener('click', () => btn.click());
  async function send() {
    const text = inp.value.trim(); if (!text) return;
    if (!me()) { $('chatState').textContent = 'Nhập tên ở ô "Người làm" trước khi nhắn.'; $('who').focus(); return; }
    inp.value = ''; inp.disabled = true;
    const local = { t: Date.now(), u: me(), m: text, pending: true };
    msgs.push(local); draw(); list.scrollTop = list.scrollHeight;
    try {
      await fetch(API, { method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ event: 'chat', user: me(), msg: text }) });
      msgs = msgs.filter(m => m !== local); // lấy lại bản có giờ chuẩn từ máy chủ
      await new Promise(r => setTimeout(r, 800)); await poll(); markSeen();
    } catch (e) { $('chatState').textContent = 'Chưa gửi được, kiểm tra mạng rồi gửi lại.'; inp.value = text; msgs = msgs.filter(m => m !== local); draw(); }
    finally { inp.disabled = false; inp.focus(); }
  }
  $('chatSend').addEventListener('click', send);
  inp.addEventListener('keydown', e => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); send(); } });
  document.addEventListener('visibilitychange', () => { if (!document.hidden) poll(); });
  poll(); schedule();
})();
