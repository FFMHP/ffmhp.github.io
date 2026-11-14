cd /home/claude/tool
bash build.sh
python3 - <<'PY'
src=open('ffm-chuyen-don.html').read()
sec=open('ghd_section.html').read()
assert '<!--BETA_GHD-->' in src
b=src.replace('<!--BETA_GHD-->',sec,1)
b=b.replace('<title>','<meta name="robots" content="noindex"><script>window.FFM_DATA_BASE="../data/"</script><title>',1)
b=b.replace('</title>',' · BẢN THỬ</title>',1)
i=b.rindex('</body>') if '</body>' in b else len(b)
b=b[:i]+'<script>'+open('ghd.js').read()+'</script>'+b[i:]
open('beta.html','w').write(b)
t=open('t.html').read().replace('<!--BETA_GHD-->',sec,1).replace('<title>','<script>window.FFM_DATA_BASE="data/"</script><title>',1)
i=t.rindex('</body>') if '</body>' in t else len(t)
t=t[:i]+'<script>'+open('ghd.js').read()+'</script>'+t[i:]
open('tbeta.html','w').write(t)
PY
