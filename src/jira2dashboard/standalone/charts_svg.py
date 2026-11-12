"""The charts of the Dashboard as SVG, drawn by Python from the results (no Excel): those of dashboard_spec.CHARTS.

One series per team (ART first, then the teams of the sheet Projects), or per work type for Flow Distribution (the
team of --team, or all teams weighted by their velocity, as the selector of the Dashboard). A point that is "not
computable" is a gap, never a 0, and the note under the chart counts the gaps as the workbook does.
"""
import datetime as dt
import math
from xml.sax.saxutils import escape

from jira2dashboard.flow import dashboard_spec as D
from jira2dashboard.flow.metrics import BUCKETS, PIPredictability, TeamIterationMetrics
from jira2dashboard.flow.model import NOT_COMPUTABLE
from jira2dashboard.standalone.xlsx_write import Picture

WIDTH, HEIGHT = 640, 380
LEFT, RIGHT, TOP = 58, 18, 44
FONT = 'font-family="Segoe UI, Arial, sans-serif"'
PALETTE = ("#2a6fb0", "#e07b00", "#3a9a4a", "#c0392b", "#7d5ba6", "#8c6d31", "#d6589f", "#5b6770", "#17a2b8",
           "#b5a300")
COLUMNS, STACKED, LINES, LINE_ONLY = "columns", "stacked", "lines", "line"
KINDS = {D.XL_COLUMN_CLUSTERED: COLUMNS, D.XL_COLUMN_STACKED_100: STACKED, D.XL_LINE_MARKERS: LINES,
         D.XL_LINE: LINE_ONLY}


def nice_top(highest: float) -> tuple[float, float]:
    """(top of the axis, step between two gridlines): 1, 2, 2.5 or 5 times a power of ten, five steps or so."""
    if highest <= 0:
        return 1.0, 0.2
    raw = highest / 5
    power = 10 ** math.floor(math.log10(raw))
    step = next(m * power for m in (1, 2, 2.5, 5, 10) if m * power >= raw)
    return step * math.ceil(highest / step - 1e-9), step


def _label(value: float, percent: bool) -> str:
    if percent:
        return f"{value * 100:.0f}%"
    return f"{value:.0f}" if float(value).is_integer() else f"{value:.1f}"


def _legend(names: list[str], top: float) -> tuple[list[str], float]:
    """Colored boxes and names under the plot, wrapped on as many lines as needed; returns the height used."""
    parts, x, y = [], LEFT, top
    for n, name in enumerate(names):
        width = 18 + 7 * len(name) + 14
        if x + width > WIDTH - RIGHT and x > LEFT:
            x, y = LEFT, y + 18
        parts.append(f'<rect x="{x}" y="{y - 9}" width="10" height="10" fill="{PALETTE[n % len(PALETTE)]}"/>'
                     f'<text x="{x + 14}" y="{y}" {FONT} font-size="11" fill="#333">{escape(name)}</text>')
        x += width
    return parts, y - top + 18


def draw(title: str, categories: list[str], series: dict[str, list[float | None]], kind: str, percent: bool = False,
         band: tuple[float, float] | None = None, note: str = "") -> str:
    """An SVG chart: `series` name -> one value per category (None = a gap)."""
    names = list(series)
    legend, legend_height = _legend(names, 0)
    bottom = 34 + legend_height + (16 if note else 0)
    plot_w, plot_h = WIDTH - LEFT - RIGHT, HEIGHT - TOP - bottom
    values = [v for column in series.values() for v in column if v is not None]
    highest = 1.0 if kind == STACKED else max(values + ([band[1]] if band else []) + [0])
    top, step = (1.0, 0.2) if kind == STACKED or (percent and highest <= 1) else nice_top(highest)
    if not percent and kind != STACKED and step < 1 and all(float(v).is_integer() for v in values):
        top, step = float(max(1, math.ceil(highest))), 1.0               # counts: no gridline between 1 and 2
    n = max(len(categories), 1)
    slot = plot_w / n

    def y(value: float) -> float:
        return TOP + plot_h - plot_h * value / top

    out = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{WIDTH}" height="{HEIGHT}" viewBox="0 0 {WIDTH} {HEIGHT}">',
           f'<rect width="{WIDTH}" height="{HEIGHT}" fill="#ffffff"/>',
           f'<text x="{WIDTH / 2}" y="24" text-anchor="middle" {FONT} font-size="14" fill="#222">{escape(title)}</text>']
    if band:
        out.append(f'<rect x="{LEFT}" y="{y(band[1]):.1f}" width="{plot_w}" height="{y(band[0]) - y(band[1]):.1f}" '
                   'fill="#e3f2e1"/>')
    tick = 0.0
    while tick <= top + 1e-9:
        out.append(f'<line x1="{LEFT}" y1="{y(tick):.1f}" x2="{WIDTH - RIGHT}" y2="{y(tick):.1f}" stroke="#e5e5e5"/>'
                   f'<text x="{LEFT - 6}" y="{y(tick) + 4:.1f}" text-anchor="end" {FONT} font-size="10" fill="#555">'
                   f'{_label(tick, percent or kind == STACKED)}</text>')
        tick += step
    out.append(f'<line x1="{LEFT}" y1="{TOP + plot_h}" x2="{WIDTH - RIGHT}" y2="{TOP + plot_h}" stroke="#888"/>')
    every = max(1, math.ceil(n / 12))
    for i, label in enumerate(categories):
        if i % every == 0:
            out.append(f'<text x="{LEFT + slot * (i + 0.5):.1f}" y="{TOP + plot_h + 16}" text-anchor="middle" {FONT} '
                       f'font-size="10" fill="#555">{escape(label)}</text>')
    if kind == COLUMNS:
        width = slot * 0.8 / max(len(names), 1)
        for s, name in enumerate(names):
            for i, value in enumerate(series[name]):
                if value is not None:
                    x = LEFT + slot * i + slot * 0.1 + width * s
                    out.append(f'<rect x="{x:.1f}" y="{y(value):.1f}" width="{width:.1f}" '
                               f'height="{TOP + plot_h - y(value):.1f}" fill="{PALETTE[s % len(PALETTE)]}"/>')
    elif kind == STACKED:
        for i in range(len(categories)):
            base = 0.0
            for s, name in enumerate(names):
                value = series[name][i]
                if value:
                    out.append(f'<rect x="{LEFT + slot * i + slot * 0.15:.1f}" y="{y(base + value):.1f}" '
                               f'width="{slot * 0.7:.1f}" height="{y(base) - y(base + value):.1f}" '
                               f'fill="{PALETTE[s % len(PALETTE)]}"/>')
                    base += value
    else:
        for s, name in enumerate(names):
            color, run = PALETTE[s % len(PALETTE)], []
            for i, value in enumerate(series[name] + [None]):
                if value is not None:
                    run.append((LEFT + slot * (i + 0.5), y(value)))
                    continue
                if len(run) > 1:
                    points = " ".join(f"{a:.1f},{b:.1f}" for a, b in run)
                    out.append(f'<polyline points="{points}" fill="none" stroke="{color}" stroke-width="2"/>')
                if kind == LINES:
                    out += [f'<circle cx="{a:.1f}" cy="{b:.1f}" r="3" fill="{color}"/>' for a, b in run]
                run = []
    if not values:
        out.append(f'<text x="{LEFT + plot_w / 2}" y="{TOP + plot_h / 2}" text-anchor="middle" {FONT} font-size="13" '
                   f'fill="#888">{NOT_COMPUTABLE}</text>')
    legend_top = TOP + plot_h + 34
    out.append(f'<g transform="translate(0,{legend_top})">' + "".join(legend) + "</g>")
    if note:
        out.append(f'<text x="{LEFT}" y="{HEIGHT - 8}" {FONT} font-size="10" fill="#666">{escape(note)}</text>')
    out.append("</svg>")
    return "\n".join(out)


def gaps_note(series: dict[str, list], reason: str) -> str:
    """The text under a chart, as the workbook writes it (dashboard_spec.note_formula); empty when there is no gap."""
    points = [v for column in series.values() for v in column]
    gaps = sum(1 for v in points if v is None)
    if not gaps:
        return ""
    return f"{gaps} of {len(points)} points are not computable (gaps)" + (f" - {reason}" if reason else "")


def _number(value) -> float | None:
    return None if value == NOT_COMPUTABLE or value is None else float(value)


def distribution(metrics: list[TeamIterationMetrics], iterations: list[str], team: str) -> dict[str, list[float | None]]:
    """The share of each work type among the items done in each iteration: one team, or all teams weighted by velocity."""
    out = {bucket: [] for bucket in BUCKETS}
    for iteration in iterations:
        rows = [m for m in metrics if m.iteration == iteration and (team == D.ALL_TEAMS or m.team == team)]
        done = sum(m.velocity for m in rows)
        for bucket in BUCKETS:
            if not done:
                out[bucket].append(None)
            else:
                out[bucket].append(sum(m.velocity * m.distribution_done[bucket] for m in rows if m.velocity) / done)
    return out


def chart_pictures(metrics: list[TeamIterationMetrics], teams: list[str], iterations: list[str],
                   loads: dict[str, dict[dt.datetime, int]], days: list[dt.datetime],
                   team: str = D.ALL_TEAMS, predictability: list[PIPredictability] = ()) -> list[Picture]:
    """The charts of dashboard_spec.CHARTS, in its order. `predictability`: Feature Predictability per PI (one row, a gap,
    when there is no PI, as in the workbook)."""
    by_key = {(m.iteration, m.team): m for m in metrics}
    attribute = {"Velocity": "velocity", "FlowTimeMedian": "flow_time_median", "FlowTimeP85": "flow_time_p85",
                 "FlowLoadAverage": "flow_load_average", "FlowEfficiency": "flow_efficiency",
                 "FlowPredictability": "flow_predictability"}
    pictures = []
    for spec in D.CHARTS:
        title = spec.title
        if spec.axis == D.DAY_AXIS:
            categories = [f"{d:%m-%d}" for d in days]
            series = {t: [loads[t].get(d) for d in days] for t in teams}
        elif spec.axis == D.PI_AXIS:
            categories = [p.pi for p in predictability] or [D.NO_PI]
            series = {name: [_number(p.ratio) for p in predictability] or [None] for name in D.series_names(spec)}
        elif spec.series == D.BUCKET_SERIES:
            categories = iterations
            series = distribution(metrics, iterations, team)
            title = title.replace("team of the selector", team)
        else:
            categories = iterations
            series = {t: [_number(getattr(by_key[(p, t)], attribute[spec.column])) if (p, t) in by_key else None
                          for p in iterations] for t in teams}
        svg = draw(title, categories, series, KINDS[spec.chart_type], spec.axis_format == "0%", spec.band,
                   gaps_note(series, spec.reason))
        pictures.append(Picture(spec.name, title, svg, WIDTH, HEIGHT))
    return pictures
