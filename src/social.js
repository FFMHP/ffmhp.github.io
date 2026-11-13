/* Sinh hoạt chung: đăng status + ảnh/video, thích, bình luận (giống bảng tin Facebook).
   Bài lưu ở sheet "Bài đăng"; ảnh/video lưu trong thư mục Google Drive "FFM Sinh hoạt chung (ảnh, video)". */
(function () {
  const $ = id => document.getElementById(id);
  const API = window.FFM_API || '';
  const page = $('pageSocial'); if (!page) return;
  const who = $('who');
  const me = () => (who.value || '').trim();
  const MAX_VIDEO = 20 * 1024 * 1024, MAX_FILES = 10;
  let posts = [], files = [], loading = false, loadedAt = 0, timer = null;

  const fmt = t => {
    const d = new Date(t), s = (Date.now() - t) / 1000, p = n => String(n).padStart(2, '0');
    if (s < 60) return 'vừa xong'; if (s < 3600) return Math.floor(s / 60) + ' phút trước';
    if (s < 86400 && new Date().toDateString() === d.toDateString()) return Math.floor(s / 3600) + ' giờ trước';
    return `${p(d.getDate())}/${p(d.getMonth() + 1)}/${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`;
  };
  const el = (tag, cls, text) => { const e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; };
  const initial = n => (n || '?').trim().split(/\s+/).pop().charAt(0).toUpperCase();
  const img = (id, w) => `https://drive.google.com/thumbnail?id=${encodeURIComponent(id)}&sz=w${w}`;
  const state = t => { $('feedState').textContent = t || ''; };

  async function api(body) {
    const r = await fetch(API, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(body) });
    const j = await r.json();
    if (j.error) throw new Error(j.error);
    return j;
  }
  function needName() { if (me()) return false; window.FFM_askName && window.FFM_askName(); state('Nhập tên của bạn (góc trên bên phải) trước nhé.'); return true; }

  // ---- soạn bài ----
  const txt = $('postText'), pick = $('postFiles'), prev = $('postPrev'), send = $('postSend');
  function drawPrev() {
    prev.innerHTML = ''; prev.hidden = !files.length;
    files.forEach((f, i) => {
      const box = el('div', 'pv');
      const m = f.type.startsWith('video/') ? Object.assign(el('video'), { src: f.url, muted: true }) : Object.assign(el('img'), { src: f.url, alt: '' });
      const x = el('button', 'pv-x', '×'); x.type = 'button'; x.title = 'Bỏ file này'; x.setAttribute('aria-label', 'Bỏ ' + f.name);
      x.onclick = () => { URL.revokeObjectURL(f.url); files.splice(i, 1); drawPrev(); };
      box.append(m, x); prev.append(box);
    });
    send.disabled = !txt.value.trim() && !files.length;
  }
  function addFiles(list) {
    for (const f of list) {
      if (!/^(image|video)\//.test(f.type)) { state(`"${f.name}" không phải ảnh hay video.`); continue; }
      if (f.type.startsWith('video/') && f.size > MAX_VIDEO) { state(`Video "${f.name}" nặng ${(f.size / 1048576).toFixed(0)} MB, tối đa 20 MB. Hãy cắt ngắn hoặc nén lại.`); continue; }
      if (files.length >= MAX_FILES) { state('Mỗi bài tối đa 10 ảnh/video.'); break; }
      files.push({ file: f, type: f.type, name: f.name, url: URL.createObjectURL(f) });
    }
    drawPrev();
  }
  $('postAdd').onclick = () => pick.click();
  pick.onchange = () => { addFiles(pick.files); pick.value = ''; };
  txt.addEventListener('input', () => { send.disabled = !txt.value.trim() && !files.length; txt.style.height = 'auto'; txt.style.height = Math.min(txt.scrollHeight, 320) + 'px'; });
  txt.addEventListener('paste', e => { const fl = [...(e.clipboardData && e.clipboardData.files || [])]; if (fl.length) { e.preventDefault(); addFiles(fl); } });
  const comp = $('composer');
  comp.addEventListener('dragover', e => { e.preventDefault(); comp.classList.add('over'); });
  comp.addEventListener('dragleave', () => comp.classList.remove('over'));
  comp.addEventListener('drop', e => { e.preventDefault(); comp.classList.remove('over'); addFiles(e.dataTransfer.files); });

  // ảnh: thu nhỏ còn cạnh dài 1600px (JPEG) cho nhẹ; GIF giữ nguyên
  async function shrink(f) {
    if (!f.type.startsWith('image/') || f.type === 'image/gif') return f.file;
    try {
      const bmp = await createImageBitmap(f.file);
      const k = Math.min(1, 1600 / Math.max(bmp.width, bmp.height));
      if (k === 1 && f.file.size < 1.5e6) return f.file;
      const c = document.createElement('canvas'); c.width = Math.round(bmp.width * k); c.height = Math.round(bmp.height * k);
      c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height);
      return await new Promise(ok => c.toBlob(b => ok(b || f.file), 'image/jpeg', 0.85));
    } catch (e) { return f.file; }
  }
  const toB64 = blob => new Promise((ok, no) => { const r = new FileReader(); r.onload = () => ok(String(r.result).split(',')[1]); r.onerror = no; r.readAsDataURL(blob); });

  send.onclick = async () => {
    if (needName()) return;
    const text = txt.value.trim(); if (!text && !files.length) return;
    send.disabled = true; txt.disabled = true; $('postAdd').disabled = true;
    try {
      const media = [];
      for (let i = 0; i < files.length; i++) {
        const f = files[i];
        state(`Đang tải lên ${i + 1}/${files.length}: ${f.name}…`);
        const blob = await shrink(f);
        const mime = blob.type || f.type;
        const r = await api({ event: 'upload', user: me(), name: f.name.replace(/\.\w+$/, '') + (mime === 'image/jpeg' ? '.jpg' : (f.name.match(/\.\w+$/) || [''])[0]), mime, data: await toB64(blob) });
        media.push({ id: r.id, type: r.type });
      }
      state('Đang đăng…');
      await api({ event: 'post', user: me(), text, media });
      files.forEach(f => URL.revokeObjectURL(f.url)); files = []; txt.value = ''; txt.style.height = ''; drawPrev();
      state('Đã đăng.'); setTimeout(() => state(''), 2500);
      loadedAt = 0; await load();
    } catch (e) { state('Chưa đăng được: ' + (e.message || e) + '. Kiểm tra mạng rồi bấm Đăng lại.'); }
    finally { txt.disabled = false; $('postAdd').disabled = false; send.disabled = !txt.value.trim() && !files.length; }
  };

  // ---- bảng tin ----
  const feed = $('feed');
  function mediaGrid(media) {
    const g = el('div', 'media n' + Math.min(media.length, 4));
    media.forEach((m, i) => {
      if (i > 3 && media.length > 4) return;
      const cell = el('div', 'mcell');
      if (m.type === 'video') {
        const f = el('iframe'); f.src = `https://drive.google.com/file/d/${encodeURIComponent(m.id)}/preview`; f.allow = 'autoplay; fullscreen'; f.loading = 'lazy'; f.title = 'Video';
        cell.classList.add('vid'); cell.append(f);
      } else {
        const b = el('button', 'mimg'); b.type = 'button'; b.setAttribute('aria-label', 'Xem ảnh lớn');
        const im = el('img'); im.src = img(m.id, 900); im.alt = ''; im.loading = 'lazy';
        b.append(im); b.onclick = () => lightbox(media, i); cell.append(b);
      }
      if (i === 3 && media.length > 4) { const more = el('span', 'more', '+' + (media.length - 4)); cell.append(more); cell.onclick = () => lightbox(media, 3); }
      g.append(cell);
    });
    return g;
  }
  function draw() {
    feed.innerHTML = '';
    if (!posts.length) { feed.append(el('p', 'panel feed-empty', loading ? 'Đang tải bảng tin…' : 'Chưa có bài nào. Đăng bài đầu tiên cho cả phòng nhé!')); return; }
    posts.forEach(p => {
      const card = el('article', 'panel post');
      const hd = el('div', 'post-h');
      hd.append(el('span', 'avatar', initial(p.u)));
      const nm = el('div', 'post-who'); nm.append(el('b', '', p.u || '?'), el('small', 'note', fmt(p.t)));
      hd.append(nm);
      if (me() && p.u.toLowerCase() === me().toLowerCase()) {
        const del = el('button', 'small ghost', 'Xoá'); del.type = 'button'; del.title = 'Xoá bài này';
        del.onclick = () => askDelete(p, del); hd.append(del);
      }
      card.append(hd);
      if (p.text) card.append(el('div', 'post-text', p.text));
      if (p.media && p.media.length) card.append(mediaGrid(p.media));
      const liked = p.likes.some(u => u.toLowerCase() === me().toLowerCase());
      const cnt = el('div', 'post-cnt note');
      cnt.textContent = [p.likes.length ? `👍 ${p.likes.length}` : '', p.comments.length ? `${p.comments.length} bình luận` : ''].filter(Boolean).join(' · ');
      if (p.likes.length) cnt.title = p.likes.join(', ');
      if (cnt.textContent) card.append(cnt);
      const act = el('div', 'post-act');
      const lb = el('button', 'act' + (liked ? ' on' : ''), liked ? '👍 Đã thích' : '👍 Thích'); lb.type = 'button'; lb.setAttribute('aria-pressed', String(liked));
      lb.onclick = () => like(p);
      const cb = el('button', 'act', '💬 Bình luận'); cb.type = 'button';
      act.append(lb, cb); card.append(act);
      const cm = el('div', 'cmts');
      p.comments.forEach(c => {
        const row = el('div', 'cmt'); row.append(el('span', 'avatar sm', initial(c.u)));
        const bub = el('div', 'cmt-b'); bub.append(el('b', '', c.u), el('span', '', c.m)); row.append(bub);
        row.append(el('small', 'note', fmt(c.t))); cm.append(row);
      });
      const form = el('div', 'cmt-new'); const ci = el('input'); ci.type = 'text'; ci.placeholder = 'Viết bình luận…'; ci.maxLength = 1000; ci.setAttribute('aria-label', 'Bình luận');
      const cs = el('button', 'small primary', 'Gửi'); cs.type = 'button';
      const go = () => comment(p, ci, cs);
      cs.onclick = go; ci.onkeydown = e => { if (e.key === 'Enter') { e.preventDefault(); go(); } };
      form.append(ci, cs); cm.append(form); card.append(cm);
      cb.onclick = () => ci.focus();
      feed.append(card);
    });
  }
  async function like(p) {
    if (needName()) return;
    const u = me(), i = p.likes.findIndex(x => x.toLowerCase() === u.toLowerCase());
    if (i >= 0) p.likes.splice(i, 1); else p.likes.push(u); draw();
    try { const r = await api({ event: 'like', id: p.id, user: u }); p.likes = r.likes; draw(); }
    catch (e) { state('Chưa thích được: ' + e.message); loadedAt = 0; load(); }
  }
  async function comment(p, inp, b) {
    if (needName()) return;
    const m = inp.value.trim(); if (!m) return;
    inp.disabled = b.disabled = true;
    try { const r = await api({ event: 'comment', id: p.id, user: me(), msg: m }); p.comments = r.comments; draw(); }
    catch (e) { state('Chưa gửi được bình luận: ' + e.message); inp.disabled = b.disabled = false; }
  }
  function askDelete(p, btn) {
    if (btn.dataset.sure) {
      btn.disabled = true; btn.textContent = 'Đang xoá…';
      api({ event: 'delpost', id: p.id, user: me() }).then(() => { posts = posts.filter(x => x !== p); draw(); state('Đã xoá bài.'); })
        .catch(e => { state('Chưa xoá được: ' + e.message); btn.disabled = false; btn.textContent = 'Xoá'; delete btn.dataset.sure; });
      return;
    }
    btn.dataset.sure = '1'; btn.textContent = 'Bấm lần nữa để xoá'; btn.classList.add('danger');
    setTimeout(() => { if (btn.isConnected && !btn.disabled) { delete btn.dataset.sure; btn.textContent = 'Xoá'; btn.classList.remove('danger'); } }, 4000);
  }

  // ---- xem ảnh lớn ----
  const lb = $('lightbox'), lbImg = $('lbImg');
  let lbList = [], lbI = 0;
  function lightbox(media, i) { lbList = media.filter(m => m.type !== 'video'); lbI = Math.max(0, lbList.indexOf(media[i])); showLb(); lb.hidden = false; $('lbClose').focus(); }
  function showLb() { const m = lbList[lbI]; if (!m) return; lbImg.src = img(m.id, 2000); $('lbOpen').href = `https://drive.google.com/file/d/${m.id}/view`; $('lbPrev').hidden = $('lbNext').hidden = lbList.length < 2; }
  $('lbClose').onclick = () => { lb.hidden = true; lbImg.src = ''; };
  $('lbPrev').onclick = () => { lbI = (lbI - 1 + lbList.length) % lbList.length; showLb(); };
  $('lbNext').onclick = () => { lbI = (lbI + 1) % lbList.length; showLb(); };
  lb.addEventListener('click', e => { if (e.target === lb) $('lbClose').click(); });
  document.addEventListener('keydown', e => { if (lb.hidden) return; if (e.key === 'Escape') $('lbClose').click(); if (e.key === 'ArrowLeft') $('lbPrev').click(); if (e.key === 'ArrowRight') $('lbNext').click(); });

  // ---- tải bảng tin ----
  async function load() {
    if (!API || loading || page.hidden || Date.now() - loadedAt < 10000) return;
    loading = true; if (!posts.length) draw();
    try {
      const j = await (await fetch(API + '?t=feed', { cache: 'no-store' })).json();
      if (j && Array.isArray(j.posts)) {
        // giữ ô bình luận đang gõ dở
        const typing = document.activeElement && document.activeElement.closest('.cmt-new');
        posts = j.posts; loadedAt = Date.now();
        if (!typing) draw();
      }
    } catch (e) { state('Chưa tải được bảng tin, đang thử lại…'); }
    finally { loading = false; if (!posts.length) draw(); }
  }
  window.FFM_social = {
    show() { load(); clearInterval(timer); timer = setInterval(load, 45000); drawPrev(); },
    hide() { clearInterval(timer); }
  };
  who.addEventListener('change', () => { if (!page.hidden) draw(); });
  document.addEventListener('visibilitychange', () => { if (!document.hidden && !page.hidden) load(); });
})();
