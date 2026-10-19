"""Standalone, offline HTML for presenting (no artifact viewer needed).

Wraps the same page body in a full HTML document: charset, viewport, the artifact's base reset,
embedded fonts, a presenter bar (section links, light/dark, full screen) and keyboard steps:
  → / Space / Page Down = next stop     ← / Page Up = previous stop
  F = full screen                       T = light / dark                 Home = top
"""
import os

HERE = os.path.dirname(os.path.abspath(__file__))

RESET = """
:root{color-scheme:light;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}
*,*::before,*::after{box-sizing:border-box}
body{margin:0;-webkit-font-smoothing:antialiased}
img{max-width:100%}
[hidden]{display:none!important}
"""

PRESENT_CSS = """
.pbar{position:fixed;top:0;left:0;right:0;z-index:50;background:var(--surface);border-bottom:1px solid var(--line);
  padding:8px max(16px, env(safe-area-inset-right, 0px)) 8px max(16px, env(safe-area-inset-left, 0px));
  display:flex;gap:14px;align-items:center;font-family:var(--body);font-size:14px}
.pbar .brand{font-family:var(--display);font-weight:700;white-space:nowrap}
.pbar nav{display:flex;gap:4px;overflow-x:auto;flex:1;min-width:0;scrollbar-width:none}
.pbar nav a{padding:5px 10px;border-radius:999px;color:var(--muted);text-decoration:none;white-space:nowrap;font-weight:600}
.pbar nav a:hover{color:var(--ink);background:var(--paper)}
.pbar nav a.on{background:var(--teal);color:var(--paper)}
.pbar .ctl{display:flex;gap:6px;align-items:center;white-space:nowrap}
.pbar button{font:inherit;font-weight:600;border:1px solid var(--line);background:var(--paper);color:var(--ink);border-radius:6px;padding:5px 10px;cursor:pointer}
.pbar button:hover{border-color:var(--teal)}
.pbar button:focus-visible,.pbar a:focus-visible{outline:2px solid var(--teal);outline-offset:2px}
.pbar .count{color:var(--muted);font-variant-numeric:tabular-nums;min-width:52px;text-align:right}
body{padding-top:76px}
section,.part,header{scroll-margin-top:72px}
.stop-on{box-shadow:0 0 0 3px var(--teal-soft);border-radius:8px}
@media (max-width:720px){.pbar .brand{display:none}.pbar .hint{display:none}}
@media print{.pbar{display:none}body{padding-top:0}}
@media (prefers-reduced-motion: reduce){html{scroll-behavior:auto}}
"""

PRESENT_JS = """
(function(){
  var parts = document.querySelectorAll('.part');
  var names = ['part-a','part-b','part-c','appendix'];
  parts.forEach(function(p,i){ if(!p.id) p.id = names[i] || ('part-'+i); });
  var ids = ['mission','answer','part-a','a1','a2','a3','part-b','b1','b2','b3','part-c','c1','c2','c3','ask','appendix'];
  var stops = ids.map(function(id){ return document.getElementById(id); }).filter(Boolean);
  var links = document.querySelectorAll('.pbar nav a');
  var count = document.getElementById('pcount');
  var cur = 0, lockUntil = 0;
  var smooth = !window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  function mark(i){
    cur = Math.max(0, Math.min(stops.length-1, i));
    count.textContent = (cur+1) + ' / ' + stops.length;
    var id = stops[cur].id, group = id[0] === 'a' && id.length===2 ? 'part-a' : id[0] === 'b' && id.length===2 ? 'part-b' : id[0] === 'c' && id.length===2 ? 'part-c' : id;
    links.forEach(function(a){ a.classList.toggle('on', a.getAttribute('href') === '#'+group); });
  }
  function go(i){ mark(i); lockUntil = Date.now() + 900; stops[cur].scrollIntoView({behavior: smooth ? 'smooth' : 'auto', block: 'start'}); }
  var obs = new IntersectionObserver(function(es){
    if (Date.now() < lockUntil) return;
    es.forEach(function(e){ if(e.isIntersecting){ var i = stops.indexOf(e.target); if(i>=0) mark(i); } });
  }, {rootMargin: '-80px 0px -70% 0px'});
  stops.forEach(function(s){ obs.observe(s); });
  document.addEventListener('keydown', function(e){
    if (e.target.closest && e.target.closest('input,textarea,select')) return;
    if (e.metaKey || e.ctrlKey || e.altKey) return;
    var k = e.key;
    if (k === 'ArrowRight' || k === 'PageDown' || k === ' ') { e.preventDefault(); go(cur+1); }
    else if (k === 'ArrowLeft' || k === 'PageUp') { e.preventDefault(); go(cur-1); }
    else if (k === 'Home') { e.preventDefault(); go(0); }
    else if (k === 'End') { e.preventDefault(); go(stops.length-1); }
    else if (k === 'f' || k === 'F') { full(); }
    else if (k === 't' || k === 'T') { theme(); }
  });
  function full(){
    try { if (!document.fullscreenElement) document.documentElement.requestFullscreen(); else document.exitFullscreen(); } catch(err) {}
  }
  function theme(){
    var r = document.documentElement, dark = r.getAttribute('data-theme') === 'dark';
    r.setAttribute('data-theme', dark ? 'light' : 'dark');
    document.getElementById('ptheme').textContent = dark ? 'Dark' : 'Light';
  }
  document.getElementById('pfull').addEventListener('click', full);
  document.getElementById('ptheme').addEventListener('click', theme);
  document.getElementById('pprev').addEventListener('click', function(){ go(cur-1); });
  document.getElementById('pnext').addEventListener('click', function(){ go(cur+1); });
  mark(0);
})();
"""

BAR = """
<div class="pbar" role="navigation" aria-label="Presenter">
  <span class="brand">CoreItera</span>
  <nav>
    <a href="#mission">Mission</a><a href="#answer">Answer</a><a href="#part-a">A. Need</a><a href="#part-b">B. Product</a>
    <a href="#part-c">C. Plan and deal</a><a href="#ask">The ask</a><a href="#appendix">Appendix</a>
  </nav>
  <span class="ctl">
    <span class="hint" style="color:var(--muted)">← → to move</span>
    <button id="pprev" type="button" aria-label="Previous section">←</button>
    <button id="pnext" type="button" aria-label="Next section">→</button>
    <span class="count" id="pcount">1 / 16</span>
    <button id="ptheme" type="button">Dark</button>
    <button id="pfull" type="button">Full screen</button>
  </span>
</div>
"""


def standalone(title, css, body):
    fonts = open(os.path.join(HERE, "fonts_embedded.css")).read()
    return f"""<!doctype html>
<html lang="en" data-theme="light">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>{title}</title>
<style>
{fonts}
{RESET}
{css}
{PRESENT_CSS}
</style>
</head>
<body>
{BAR}
{body}
<script>{PRESENT_JS}</script>
</body>
</html>
"""
