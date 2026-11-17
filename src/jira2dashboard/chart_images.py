"""The PNG of the charts that a test exports to be looked at (AL, 2026-11-17): they stay after `build` / `vba-tests`
(`paths.CHARTS_DIR`, outside `out\\inter`) and each run replaces them: one sub-folder per scenario, emptied of its
old PNG before the export, and a note `_exported.txt` written AFTER it, so that a run that failed before the export
shows an old time (or none) instead of images that look current."""
import datetime as dt
import html
import subprocess
from pathlib import Path

NOTE_NAME = "_exported.txt"
INDEX_NAME = "index.html"


def prepare(folder: Path) -> None:
    """Make the folder and remove the images, the page and the note of the previous run (a chart renamed or removed leaves nothing)."""
    folder.mkdir(parents=True, exist_ok=True)
    for old in list(folder.glob("*.png")) + list(folder.glob("*.svg")) + [folder / NOTE_NAME, folder / INDEX_NAME]:
        old.unlink(missing_ok=True)


def commit_of(repo: Path) -> str:
    """The short hash of HEAD, or "unknown" (no git, not a repository)."""
    try:
        done = subprocess.run(["git", "rev-parse", "--short", "HEAD"], cwd=repo, capture_output=True, text=True,
                              timeout=10, check=True)
        return done.stdout.strip() or "unknown"
    except (OSError, subprocess.SubprocessError):
        return "unknown"


def note_text(scenario: str, names: list[str], commit: str, now: dt.datetime) -> str:
    return (f"charts of the scenario {scenario}\n"
            f"exported (UTC): {now:%Y-%m-%d %H:%M:%S}\n"
            f"commit: {commit}\n"
            f"images: {len(names)}\n" + "".join(f"  {n}\n" for n in sorted(names)))


def write_note(folder: Path, scenario: str, names: list[str], commit: str) -> Path:
    note = folder / NOTE_NAME
    note.write_text(note_text(scenario, names, commit, dt.datetime.now(dt.timezone.utc)), encoding="utf-8")
    return note


# The page: a dashboard of the SVG charts (Excel exports them as vectors), click = zoom. No script or font from outside:
# the page opens by double-click, offline.
_CSS = """
:root{--bg:#f4f5f7;--card:#fff;--ink:#1d2330;--muted:#5b6475;--line:#d9dde5;--accent:#1f5fbf}
@media (prefers-color-scheme:dark){:root{--bg:#14171d;--card:#1e222b;--ink:#e6e9ef;--muted:#9aa3b5;--line:#323846;--accent:#6ea8ff}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.45 system-ui,Segoe UI,sans-serif}
header{padding:20px 24px 8px}h1{margin:0;font-size:22px}.meta{color:var(--muted);font-size:13px;margin-top:4px}
h2{font-size:16px;margin:22px 24px 8px}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(min(100%,440px),1fr));gap:16px;padding:0 24px 24px}
.card{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:10px 12px}
.card h3{margin:0 0 8px;font-size:14px}.card a.pic{display:block;background:#fff;border-radius:4px;cursor:zoom-in}
.card img{display:block;width:100%;height:auto}.links{margin-top:6px;font-size:12px;color:var(--muted)}
.links a{color:var(--accent);margin-right:10px}
.verdict{margin:0 24px 8px;font-weight:600}.verdict.ok{color:#1a7f37}.verdict.diff{color:#c62828}
details.rep{margin:0 24px 10px;background:var(--card);border:1px solid var(--line);border-radius:8px;padding:6px 12px}
details.rep summary{cursor:pointer;font-weight:600}details.rep table{border-collapse:collapse;margin:8px 0;font-size:13px}
details.rep th,details.rep td{border-bottom:1px solid var(--line);padding:3px 12px 3px 0;text-align:left;font-variant-numeric:tabular-nums}
details.rep tr.diff td{color:#c62828;font-weight:600}
#zoom{position:fixed;inset:0;background:rgba(10,12,16,.92);display:flex;align-items:center;justify-content:center;overflow:hidden;touch-action:none;z-index:9}
#zoom[hidden]{display:none}#zoom img{max-width:92vw;max-height:84vh;background:#fff;transform-origin:center;cursor:grab;user-select:none}
#zoom .bar{position:fixed;top:0;left:0;right:0;display:flex;gap:8px;align-items:center;padding:10px 16px;color:#fff}
#zoom .cap{flex:1;font-size:14px}#zoom button{background:#2b3140;color:#fff;border:1px solid #475066;border-radius:6px;padding:4px 12px;font-size:15px;cursor:pointer}
"""

_JS = """
(function(){
  var box=document.getElementById('zoom'),img=box.querySelector('img'),cap=box.querySelector('.cap');
  var s=1,x=0,y=0,drag=null;
  function apply(){img.style.transform='translate('+x+'px,'+y+'px) scale('+s+')';}
  function reset(){s=1;x=0;y=0;apply();}
  function zoom(f,cx,cy){x=cx-(cx-x)*f;y=cy-(cy-y)*f;s*=f;apply();}
  function open(src,title){img.src=src;cap.textContent=title;reset();box.hidden=false;}
  function close(){box.hidden=true;img.removeAttribute('src');}
  Array.prototype.forEach.call(document.querySelectorAll('a[data-zoom]'),function(a){
    a.addEventListener('click',function(e){e.preventDefault();open(a.getAttribute('href'),a.getAttribute('data-title'));});});
  box.addEventListener('wheel',function(e){e.preventDefault();
    zoom(e.deltaY<0?1.15:1/1.15,e.clientX-innerWidth/2,e.clientY-innerHeight/2);},{passive:false});
  img.addEventListener('mousedown',function(e){e.preventDefault();drag={x:e.clientX-x,y:e.clientY-y};});
  window.addEventListener('mousemove',function(e){if(drag){x=e.clientX-drag.x;y=e.clientY-drag.y;apply();}});
  window.addEventListener('mouseup',function(){drag=null;});
  img.addEventListener('dblclick',reset);
  box.addEventListener('click',function(e){if(e.target===box)close();});
  document.getElementById('zin').onclick=function(){zoom(1.25,0,0);};
  document.getElementById('zout').onclick=function(){zoom(1/1.25,0,0);};
  document.getElementById('zreset').onclick=reset;
  document.getElementById('zclose').onclick=close;
  window.addEventListener('keydown',function(e){if(box.hidden)return;
    if(e.key==='Escape')close();else if(e.key==='+'||e.key==='=')zoom(1.25,0,0);
    else if(e.key==='-')zoom(1/1.25,0,0);else if(e.key==='0')reset();});
})();
"""


def index_html(scenario: str, groups: list[tuple[str, list[tuple[str, str]]]], commit: str, now: dt.datetime, report: str = "") -> str:
    """The page of one scenario. `groups` = [(sheet, [(title, chart name)])]: the image of a chart is `<name>.svg`, with `<name>.png` beside it.
    `report` = an HTML fragment (the report expected / obtained of tests/art_report.py) shown above the charts."""
    e = html.escape
    count = sum(len(charts) for _sheet, charts in groups)
    out = [f'<!doctype html><html lang="en"><head><meta charset="utf-8">'
           f'<meta name="viewport" content="width=device-width,initial-scale=1"><link rel="icon" href="data:,">'
           f'<title>Charts {e(scenario)}</title><style>{_CSS}</style></head><body>'
           f'<header><h1>Charts of the scenario {e(scenario)}</h1>'
           f'<div class="meta">{count} charts - exported (UTC) {now:%Y-%m-%d %H:%M:%S} - commit {e(commit)} - '
           f'click a chart to zoom (wheel = zoom, drag = move, double click = reset, Esc = close)</div></header>']
    if report:
        out.append(f'<h2>Expected (worked out by hand) and obtained (workbook)</h2>{report}')
    for sheet, charts in groups:
        out.append(f'<h2>{e(sheet)}</h2><div class="grid">')
        for title, name in charts:
            out.append(f'<div class="card"><h3>{e(title)}</h3>'
                       f'<a class="pic" data-zoom data-title="{e(title, quote=True)}" href="{e(name, quote=True)}.svg">'
                       f'<img loading="lazy" src="{e(name, quote=True)}.svg" alt="{e(title, quote=True)}"></a>'
                       f'<div class="links"><a href="{e(name, quote=True)}.svg">SVG</a><a href="{e(name, quote=True)}.png">PNG</a></div></div>')
        out.append('</div>')
    out.append('<div id="zoom" hidden><div class="bar"><span class="cap"></span>'
               '<button id="zout" title="-">-</button><button id="zin" title="+">+</button>'
               '<button id="zreset">1:1</button><button id="zclose">Close</button></div><img alt=""></div>'
               f'<script>{_JS}</script></body></html>')
    return "".join(out)


def write_index(folder: Path, scenario: str, groups: list[tuple[str, list[tuple[str, str]]]], commit: str, report: str = "") -> Path:
    page = folder / INDEX_NAME
    page.write_text(index_html(scenario, groups, commit, dt.datetime.now(dt.timezone.utc), report), encoding="utf-8")
    return page


def root_index_html(scenarios: list[tuple[str, str]]) -> str:
    """The page of the folder of all scenarios: `scenarios` = [(folder name, exported time text)]."""
    e = html.escape
    items = "".join(f'<li><a href="{e(name, quote=True)}/{INDEX_NAME}">{e(name)}</a> <span class="meta">{e(when)}</span></li>'
                    for name, when in scenarios)
    return (f'<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
            f'<title>Charts of the tests</title><style>{_CSS}li{{margin:8px 0}}ul{{padding:0 24px 24px 44px}}a{{color:var(--accent)}}</style></head>'
            f'<body><header><h1>Charts of the tests</h1><div class="meta">one page per scenario, replaced at each run</div></header>'
            f'<ul>{items}</ul></body></html>')


def write_root_index(charts_dir: Path) -> Path | None:
    """`charts_dir/index.html` links the scenario folders that have a page; None when there is none."""
    found = []
    for sub in sorted(p for p in charts_dir.iterdir() if p.is_dir() and (p / INDEX_NAME).exists()):
        note = sub / NOTE_NAME
        lines = note.read_text(encoding="utf-8").splitlines() if note.exists() else []
        found.append((sub.name, lines[1] if len(lines) > 1 else ""))
    if not found:
        return None
    page = charts_dir / INDEX_NAME
    page.write_text(root_index_html(found), encoding="utf-8")
    return page
