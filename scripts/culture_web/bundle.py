import json, re, sys, os
root, out = sys.argv[1], sys.argv[2]
deck = json.load(open(os.path.join(root, "project/deck.json"), encoding="utf-8"))
secs = []
for sid in deck["order"]:
    h = open(os.path.join(root, f"project/slides/{sid}.html"), encoding="utf-8").read().strip()
    secs.append(h.replace("<section ", '<section class="slide" ', 1))
fonts = "".join(f'<link rel="stylesheet" href="{f["href"]}">\n' for f in deck["faces"].values() if "href" in f)
html = f'''<!doctype html>
<html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{deck["title"]}</title>
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
{fonts}<style>
*{{box-sizing:border-box}}
html,body{{margin:0;height:100%;background:#0B2F2E;overflow:hidden;font-family:'Jua','M PLUS Rounded 1c',Arial,sans-serif}}
#stage{{position:absolute;left:50%;top:50%;width:1920px;height:1080px;transform-origin:center center}}
.slide{{position:absolute;inset:0;width:1920px;height:1080px;overflow:hidden;display:none !important}}
.slide.on{{display:flex !important}}
.slide[style*="display:grid"].on{{display:grid !important}}
.slide h1,.slide h2,.slide h3,.slide p{{margin:0;word-break:keep-all}}
.slide h1{{line-height:1.1}} .slide h2{{line-height:1.15}} .slide h3{{line-height:1.2}} .slide p{{line-height:1.4}}
.slide aside{{display:none}}
[data-build-in]{{transition:opacity .45s ease}}
[data-build-in].hide{{opacity:0}}
#bar{{position:fixed;right:16px;bottom:12px;display:flex;gap:8px;align-items:center;font:16px Arial,sans-serif;color:#EAF6F5;opacity:.25;transition:opacity .2s}}
#bar:hover{{opacity:1}}
#bar button{{font:inherit;background:#0E7C77;color:#fff;border:0;border-radius:8px;padding:6px 12px;cursor:pointer}}
#hint{{position:fixed;left:16px;bottom:12px;font:14px Arial,sans-serif;color:#EAF6F5;opacity:.35}}
</style></head>
<body>
<div id="stage">
{chr(10).join(secs)}
</div>
<div id="bar"><button id="prev" aria-label="이전">◀</button><span id="num"></span><button id="next" aria-label="다음">▶</button><button id="fs" aria-label="전체 화면">⛶</button></div>
<div id="hint">→ / Enter / 클릭: 다음 · ←: 이전 · F: 전체 화면</div>
<script>
(function(){{
  var slides=[].slice.call(document.querySelectorAll('.slide')), cur=0, stage=document.getElementById('stage');
  function builds(s){{return [].slice.call(s.querySelectorAll('[data-build-in]')).sort(function(a,b){{
    var n=function(e){{var m=(e.getAttribute('data-build-in')||'').match(/\\d+/);return m?+m[0]:0}};return n(a)-n(b);}});}}
  function fit(){{var k=Math.min(innerWidth/1920,innerHeight/1080);stage.style.transform='translate(-50%,-50%) scale('+k+')';}}
  function show(i,fromEnd){{
    cur=Math.max(0,Math.min(slides.length-1,i));
    slides.forEach(function(s,j){{s.classList.toggle('on',j===cur)}});
    builds(slides[cur]).forEach(function(e){{e.classList.toggle('hide',!fromEnd)}});
    document.getElementById('num').textContent=(cur+1)+' / '+slides.length;
    history.replaceState(null,'','#slide-'+(cur+1));
  }}
  function next(){{var h=builds(slides[cur]).filter(function(e){{return e.classList.contains('hide')}});
    if(h.length){{h[0].classList.remove('hide');}} else if(cur<slides.length-1) show(cur+1);}}
  function prev(){{var v=builds(slides[cur]).filter(function(e){{return !e.classList.contains('hide')}});
    if(v.length){{v[v.length-1].classList.add('hide');}} else if(cur>0) show(cur-1,true);}}
  function fs(){{if(document.fullscreenElement)document.exitFullscreen();else document.documentElement.requestFullscreen&&document.documentElement.requestFullscreen();}}
  addEventListener('keydown',function(e){{
    if(['ArrowRight','ArrowDown','Enter',' ','PageDown'].indexOf(e.key)>=0){{e.preventDefault();next();}}
    else if(['ArrowLeft','ArrowUp','Backspace','PageUp'].indexOf(e.key)>=0){{e.preventDefault();prev();}}
    else if(e.key==='Home')show(0); else if(e.key==='End')show(slides.length-1,true);
    else if(e.key==='f'||e.key==='F')fs();
  }});
  stage.addEventListener('click',next);
  document.getElementById('next').onclick=next; document.getElementById('prev').onclick=prev; document.getElementById('fs').onclick=fs;
  var tx=null; addEventListener('touchstart',function(e){{tx=e.touches[0].clientX}},{{passive:true}});
  addEventListener('touchend',function(e){{if(tx===null)return;var d=e.changedTouches[0].clientX-tx;if(Math.abs(d)>50){{d<0?next():prev();}}tx=null;}});
  addEventListener('resize',fit); fit();
  var m=location.hash.match(/slide-(\\d+)/); show(m?+m[1]-1:0);
}})();
</script>
</body></html>
'''
open(out, "w", encoding="utf-8").write(html)
print("slides:", len(secs))
