import html
import json
import math
import re
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parent
SHARED_DRIVE_DIR = Path(
    "/Users/taegyu/Library/CloudStorage/GoogleDrive-taegyu@tracecorp.co.kr/"
    "\u1100\u1169\u11bc\u110b\u1172 \u1103\u1173\u1105\u1161\u110b\u1175\u1107\u1173/"
    "(\u110c\u116e)\u1110\u1173\u1105\u1166\u110b\u1175\u1109\u1173"
)
OUTPUT_PATHS = [ROOT / "wbr_dashboard.html", ROOT / "index.html"]
WORKBOOK_PATTERN = re.compile(r"Trace_WBR_Master_W(\d+)\.xlsx$")

NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

BLUE = "#3f4cc6"
PINK = "#f7bbc7"
LEFT_AXIS = "#0284d8"
RIGHT_AXIS = "#f55a14"
TEXT = "#3f3f46"
GRID = "#d7dfeb"
BORDER = "#111111"

MONEY_METRICS = {
    "\uc21c\ub9e4\ucd9c",
    "\ub9e4\ucd9c\ucd1d\uc774\uc775",
    "\uacf5\ud5cc\uc774\uc775",
    "\uc601\uc5c5\uc774\uc775",
    "\uc601\uc5c5\ud604\uae08\ud750\ub984 (OCF)",
    "\ud604\uae08\uc794\uc561",
    "\ub9c8\ucf00\ud305\ube44: \ub0b4\ubd80 (\uad11\uace0)",
    "\ub9c8\ucf00\ud305\ube44: \uc678\ubd80 (\uccb4\ud5d8\ub2e8, \uc2ac\ub86f)",
    "\ubcc0\ub3d9\ube44",
    "\uace0\uc815\ube44",
    "\uac1d\ub2e8\uac00 (AOV)",
    "\uc8fc\ubb38\ub2f9 \ub9c8\ucf00\ud305\ube44",
}

PERCENT_METRICS = {
    "\uacf5\ud5cc\uc774\uc775\ub960",
    "\uc601\uc5c5\uc774\uc775\ub960",
    "\uc804\ud658\uc728 (CVR)",
    "TACOS (\ucd1d \ub9c8\ucf00\ud305\ube44%)",
    "ROAS (\ucfe0\ud321\uad11\uace0)",
    "\ub9c8\ucf00\ud305(\uc678\ubd80) ROI",
}

DAYS_METRICS = {"\u2514 \uc7ac\uace0 \ud655\ubcf4\uc77c\uc218 (DOS)"}

METRIC_ORDER = [
    "\uc21c\ub9e4\ucd9c",
    "\ub9e4\ucd9c\ucd1d\uc774\uc775",
    "\uacf5\ud5cc\uc774\uc775",
    "\uacf5\ud5cc\uc774\uc775\ub960",
    "\uc601\uc5c5\uc774\uc775",
    "\uc601\uc5c5\ud604\uae08\ud750\ub984 (OCF)",
    "\uc21c\ubc29\ubb38\uc790\uc218",
    "\uc21c\uc8fc\ubb38\uc218",
    "\uc7ac\uace0\uc218\ub7c9",
]

METRIC_META = {
    "\uc21c\ub9e4\ucd9c": {"title": "1. Net Revenue ($ K)", "legend": ("Net Revenue - CY", "Net Revenue - Trend")},
    "\ub9e4\ucd9c\ucd1d\uc774\uc775": {"title": "2. Gross Profit ($ K)", "legend": ("Gross Profit - CY", "Gross Profit - Trend")},
    "\uacf5\ud5cc\uc774\uc775": {"title": "3. Contribution Profit ($ K)", "legend": ("Contribution Profit - CY", "Contribution Profit - Trend")},
    "\uacf5\ud5cc\uc774\uc775\ub960": {"title": "4. Contribution Margin (%)", "legend": ("Contribution Margin - CY", "Contribution Margin - Trend")},
    "\uc601\uc5c5\uc774\uc775": {"title": "5. Operating Profit ($ K)", "legend": ("Operating Profit - CY", "Operating Profit - Trend")},
    "\uc601\uc5c5\ud604\uae08\ud750\ub984 (OCF)": {"title": "6. Operating Cash Flow ($ K)", "legend": ("Operating Cash Flow - CY", "Operating Cash Flow - Trend")},
    "\uc21c\ubc29\ubb38\uc790\uc218": {"title": "7. Net Visitors", "legend": ("Net Visitors - CY", "Net Visitors - Trend")},
    "\uc21c\uc8fc\ubb38\uc218": {"title": "8. Net Orders", "legend": ("Net Orders - CY", "Net Orders - Trend")},
    "\uc7ac\uace0\uc218\ub7c9": {"title": "9. Inventory Units", "legend": ("Inventory - CY", "Inventory - Trend")},
}


def find_latest_workbook() -> Path:
    candidates = []
    for path in SHARED_DRIVE_DIR.glob("Trace_WBR_Master_W*.xlsx"):
        match = WORKBOOK_PATTERN.fullmatch(path.name)
        if match:
            candidates.append((int(match.group(1)), path))
    if not candidates:
        raise FileNotFoundError(f"No Trace_WBR_Master_W*.xlsx file found in {SHARED_DRIVE_DIR}")
    return max(candidates, key=lambda item: item[0])[1]


def load_sheet_rows(path: Path, target_sheet_name: str) -> List[Dict[str, Optional[str]]]:
    with zipfile.ZipFile(path) as zf:
        shared_strings: List[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            shared_root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in shared_root.findall("a:si", NS):
                shared_strings.append("".join(t.text or "" for t in si.iterfind(".//a:t", NS)))

        workbook_root = ET.fromstring(zf.read("xl/workbook.xml"))
        rel_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rel_root}

        sheet_target = None
        for sheet in workbook_root.findall("a:sheets/a:sheet", NS):
            if sheet.attrib.get("name") != target_sheet_name:
                continue
            rel_id = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            sheet_target = rel_map[rel_id]
            break

        if sheet_target is None:
            raise KeyError(f"Sheet {target_sheet_name!r} not found in {path}")

        sheet_root = ET.fromstring(zf.read(f"xl/{sheet_target}"))
        rows: Dict[int, Dict[str, Optional[str]]] = {}
        for cell in sheet_root.iterfind(".//a:c", NS):
            ref = cell.attrib["r"]
            col = re.match(r"([A-Z]+)", ref).group(1)
            row_num = int(re.search(r"(\d+)$", ref).group(1))
            value_node = cell.find("a:v", NS)
            cell_type = cell.attrib.get("t")

            if cell_type == "s" and value_node is not None:
                value = shared_strings[int(value_node.text)]
            elif value_node is not None:
                value = value_node.text
            else:
                inline = cell.find("a:is", NS)
                value = "".join(t.text or "" for t in inline.iterfind(".//a:t", NS)) if inline is not None else None

            rows.setdefault(row_num, {})[col] = value

    return [rows[idx] for idx in sorted(rows)]


def parse_number(raw: Optional[str]) -> Optional[float]:
    if raw is None:
        return None
    text = str(raw).strip().replace(",", "")
    if text in {"", "-", "None", "#ERROR!", "#VALUE!"}:
        return None
    try:
        return float(text)
    except ValueError:
        if text.endswith("%"):
            try:
                return float(text[:-1]) / 100
            except ValueError:
                return None
        return None


def scale_value(metric: str, value: Optional[float]) -> Optional[float]:
    if value is None:
        return None
    if metric in MONEY_METRICS:
        return value / 1000
    return value


def format_scaled(metric: str, value: Optional[float], decimals: int = 0) -> str:
    if value is None:
        return "-"
    scaled = scale_value(metric, value)
    if scaled is None:
        return "-"
    if metric in PERCENT_METRICS:
        return f"{scaled * 100:.1f}%"
    if metric in DAYS_METRICS:
        return f"{scaled:.0f} Days"
    if metric in MONEY_METRICS:
        return f"{scaled:.{decimals}f}K"
    return f"{scaled:,.0f}"


def format_axis(metric: str, value: float) -> str:
    if metric in PERCENT_METRICS:
        return f"{value * 100:.0f}%"
    if metric in DAYS_METRICS:
        return f"{value:.0f}"
    if metric in MONEY_METRICS:
        return f"{value:.0f}K"
    if abs(value) >= 1000:
        return f"{value/1000:.1f}K"
    return f"{value:.0f}"


def format_point(metric: str, value: Optional[float]) -> str:
    if value is None:
        return "-"
    scaled = scale_value(metric, value)
    if metric in PERCENT_METRICS:
        return f"{scaled * 100:.1f}"
    if metric in DAYS_METRICS:
        return f"{scaled:.0f}"
    if metric in MONEY_METRICS:
        return f"{scaled:.0f}"
    return f"{scaled:,.0f}"


def format_change(metric: str, value: Optional[float]) -> str:
    if value is None:
        return "-"
    if metric in PERCENT_METRICS:
        return f"{value * 100:+.1f}%p"
    return f"{value * 100:+.1f}%"


def nice_bounds(values: List[float], include_zero: bool = False) -> tuple[float, float]:
    v_min = min(values)
    v_max = max(values)
    if include_zero:
        v_min = min(v_min, 0)
        v_max = max(v_max, 0)
    if math.isclose(v_min, v_max):
        bump = abs(v_min) * 0.1 or 1
        return v_min - bump, v_max + bump

    span = v_max - v_min
    padding = span * 0.15
    v_min -= padding
    v_max += padding
    if include_zero:
        v_min = min(v_min, 0)
        v_max = max(v_max, 0)
    return v_min, v_max


def rolling_trend(series: List[Optional[float]]) -> List[Optional[float]]:
    values: List[Optional[float]] = []
    for idx, value in enumerate(series):
        if value is None:
            values.append(None)
            continue
        window = [item for item in series[max(0, idx - 2) : idx + 1] if item is not None]
        values.append(sum(window) / len(window))
    return values


def path_from_points(points: List[tuple[float, float]]) -> str:
    if not points:
        return ""
    if len(points) == 1:
        x, y = points[0]
        return f"M{x:.1f},{y:.1f}"

    parts = [f"M{points[0][0]:.1f},{points[0][1]:.1f}"]
    for idx in range(len(points) - 1):
        p0 = points[idx - 1] if idx > 0 else points[idx]
        p1 = points[idx]
        p2 = points[idx + 1]
        p3 = points[idx + 2] if idx + 2 < len(points) else p2
        c1x = p1[0] + (p2[0] - p0[0]) / 6
        c1y = p1[1] + (p2[1] - p0[1]) / 6
        c2x = p2[0] - (p3[0] - p1[0]) / 6
        c2y = p2[1] - (p3[1] - p1[1]) / 6
        parts.append(f"C{c1x:.1f},{c1y:.1f} {c2x:.1f},{c2y:.1f} {p2[0]:.1f},{p2[1]:.1f}")
    return " ".join(parts)


def chart_svg(metric: str, labels: List[str], series: List[Optional[float]]) -> str:
    scaled_series = [scale_value(metric, value) for value in series]
    trend_series = [scale_value(metric, value) for value in rolling_trend(series)]

    actual_values = [value for value in scaled_series if value is not None]
    trend_values = [value for value in trend_series if value is not None]
    left_min, left_max = nice_bounds(actual_values, include_zero=any(v < 0 for v in actual_values))
    right_min, right_max = nice_bounds(trend_values, include_zero=any(v < 0 for v in trend_values))

    width = 1320
    height = 470
    left = 130
    right = 1160
    top = 58
    bottom = 350

    def x_pos(idx: int) -> float:
        if len(labels) == 1:
            return (left + right) / 2
        return left + (right - left) * idx / (len(labels) - 1)

    def y_left(value: float) -> float:
        return bottom - (value - left_min) * (bottom - top) / (left_max - left_min)

    def y_right(value: float) -> float:
        return bottom - (value - right_min) * (bottom - top) / (right_max - right_min)

    grid = []
    for idx in range(5):
        ratio = idx / 4
        y = top + (bottom - top) * ratio
        left_value = left_max - (left_max - left_min) * ratio
        right_value = right_max - (right_max - right_min) * ratio
        grid.append(f'<line x1="{left}" y1="{y:.1f}" x2="{right}" y2="{y:.1f}" class="grid-line" />')
        grid.append(
            f'<text x="{left - 18}" y="{y + 6:.1f}" text-anchor="end" class="axis-left">{html.escape(format_axis(metric, left_value))}</text>'
        )
        grid.append(
            f'<text x="{right + 18}" y="{y + 6:.1f}" text-anchor="start" class="axis-right">{html.escape(format_axis(metric, right_value))}</text>'
        )

    if left_min < 0 < left_max:
        zero_y = y_left(0)
        grid.append(f'<line x1="{left}" y1="{zero_y:.1f}" x2="{right}" y2="{zero_y:.1f}" class="zero-line" />')

    trend_points = [(x_pos(idx), y_right(value)) for idx, value in enumerate(trend_series) if value is not None]
    actual_points = [(x_pos(idx), y_left(value)) for idx, value in enumerate(scaled_series) if value is not None]

    x_labels = []
    for idx, label in enumerate(labels):
        x = x_pos(idx)
        x_labels.append(
            f'<text x="{x:.1f}" y="{bottom + 48}" transform="rotate(-32 {x:.1f} {bottom + 48})" text-anchor="end" class="x-axis">{html.escape(label.lower())}</text>'
        )

    point_labels = []
    for idx, value in enumerate(scaled_series):
        if value is None:
            continue
        x = x_pos(idx)
        y = y_left(value)
        point_labels.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="8" class="actual-dot" />')
        point_labels.append(
            f'<text x="{x:.1f}" y="{y - 18:.1f}" text-anchor="middle" class="point-label">{html.escape(format_point(metric, series[idx]))}</text>'
        )

    return f"""
    <svg viewBox="0 0 {width} {height}" class="chart" role="img" aria-label="{html.escape(metric)} chart">
      {''.join(grid)}
      <path d="{path_from_points(trend_points)}" class="trend-line" />
      <path d="{path_from_points(actual_points)}" class="actual-line" />
      {''.join(point_labels)}
      {''.join(x_labels)}
    </svg>
    """


def summary_row(metric: str, payload: Dict[str, Optional[float]]) -> str:
    items = [
        ("LastWk", format_scaled(metric, payload["series"][-1])),
        ("WOW", format_change(metric, payload["wow"])),
        ("6W Avg", format_scaled(metric, payload["avg6"])),
        ("QTD", format_scaled(metric, payload["qtd"])),
        ("YTD", format_scaled(metric, payload["ytd"])),
        ("MoM", format_change(metric, payload["mom"])),
    ]
    return "".join(
        f"""
        <div class="summary-item">
          <div class="summary-label">{label}</div>
          <div class="summary-value">{html.escape(value)}</div>
        </div>
        """
        for label, value in items
    )


def build_dashboard(workbook_path: Path) -> str:
    rows = load_sheet_rows(workbook_path, "WBR Dashboard")
    row_map = {idx + 1: row for idx, row in enumerate(rows)}

    labels = [row_map[6].get(col, "") or "" for col in ["B", "C", "D", "E", "F", "G"]]
    latest_week = row_map[4].get("B", "") or labels[-1]
    current_month = row_map[4].get("F", "") or "-"
    current_quarter = row_map[4].get("J", "") or "-"

    metrics: Dict[str, Dict[str, Optional[float]]] = {}
    for row_num in range(8, 33):
        row = row_map.get(row_num, {})
        metric = (row.get("A") or "").strip()
        if metric not in METRIC_META:
            continue
        metrics[metric] = {
            "series": [parse_number(row.get(col)) for col in ["B", "C", "D", "E", "F", "G"]],
            "wow": parse_number(row.get("H")),
            "avg6": parse_number(row.get("I")),
            "qtd": parse_number(row.get("J")),
            "ytd": parse_number(row.get("K")),
            "mom": parse_number(row.get("L")),
        }

    sections = []
    for metric in METRIC_ORDER:
        payload = metrics[metric]
        meta = METRIC_META[metric]
        legend_a, legend_b = meta["legend"]
        sections.append(
            f"""
            <section class="report-panel">
              <h2>{html.escape(meta['title'])}</h2>
              {chart_svg(metric, labels, payload['series'])}
              <div class="legend">
                <span class="legend-item"><span class="legend-line actual"></span>{html.escape(legend_a)}</span>
                <span class="legend-item"><span class="legend-line trend"></span>{html.escape(legend_b)}</span>
              </div>
              <div class="summary-grid">
                {summary_row(metric, payload)}
              </div>
            </section>
            """
        )

    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")
    payload = {
        "labels": labels,
        "latestWeek": latest_week,
        "month": current_month,
        "quarter": current_quarter,
        "metrics": metrics,
    }

    return f"""<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Trace WBR Dashboard</title>
  <style>
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      background: #f7f7f5;
      color: {TEXT};
      font-family: "Avenir Next", "Pretendard", "Apple SD Gothic Neo", sans-serif;
    }}
    .page {{
      max-width: 1520px;
      margin: 0 auto;
      padding: 20px 16px 44px;
    }}
    .masthead {{
      border: 2px solid {BORDER};
      background: #fff;
      padding: 20px 24px;
      margin-bottom: 18px;
    }}
    .masthead-top {{
      display: flex;
      justify-content: space-between;
      gap: 12px;
      align-items: baseline;
      flex-wrap: wrap;
    }}
    .masthead h1 {{
      margin: 0;
      font-size: 36px;
      color: #27272a;
    }}
    .masthead .meta {{
      font-size: 18px;
      color: #52525b;
    }}
    .masthead .sub {{
      margin-top: 8px;
      font-size: 14px;
      color: #71717a;
    }}
    .report-panel {{
      background: #fff;
      border: 2px solid {BORDER};
      padding: 22px 18px 18px;
      margin-bottom: 18px;
    }}
    .report-panel h2 {{
      margin: 0 0 4px;
      text-align: center;
      font-size: 30px;
      color: #4b5563;
      font-weight: 700;
    }}
    .chart {{
      width: 100%;
      height: auto;
      display: block;
    }}
    .grid-line {{
      stroke: {GRID};
      stroke-width: 2;
    }}
    .zero-line {{
      stroke: #9ca3af;
      stroke-width: 2;
      stroke-dasharray: 8 8;
    }}
    .axis-left {{
      fill: {LEFT_AXIS};
      font-size: 22px;
      font-weight: 500;
    }}
    .axis-right {{
      fill: {RIGHT_AXIS};
      font-size: 22px;
      font-weight: 500;
    }}
    .x-axis {{
      fill: #6b7280;
      font-size: 19px;
    }}
    .actual-line {{
      fill: none;
      stroke: {BLUE};
      stroke-width: 5;
      stroke-linecap: round;
      stroke-linejoin: round;
    }}
    .trend-line {{
      fill: none;
      stroke: {PINK};
      stroke-width: 4;
      stroke-linecap: round;
      stroke-linejoin: round;
    }}
    .actual-dot {{
      fill: {BLUE};
    }}
    .point-label {{
      fill: #303038;
      font-size: 19px;
      font-weight: 600;
    }}
    .legend {{
      display: flex;
      justify-content: center;
      gap: 28px;
      flex-wrap: wrap;
      margin: 2px 0 18px;
      font-size: 18px;
      color: #3f3f46;
    }}
    .legend-item {{
      display: inline-flex;
      align-items: center;
      gap: 10px;
    }}
    .legend-line {{
      width: 34px;
      height: 0;
      border-top-width: 4px;
      border-top-style: solid;
      position: relative;
    }}
    .legend-line::after {{
      content: "";
      width: 12px;
      height: 12px;
      border-radius: 50%;
      position: absolute;
      right: 9px;
      top: -8px;
      background: currentColor;
    }}
    .legend-line.actual {{
      border-top-color: {BLUE};
      color: {BLUE};
    }}
    .legend-line.trend {{
      border-top-color: {PINK};
      color: {PINK};
    }}
    .summary-grid {{
      display: grid;
      grid-template-columns: repeat(6, minmax(0, 1fr));
      gap: 10px;
      padding-top: 14px;
      border-top: 1px solid #d4d4d8;
    }}
    .summary-item {{
      text-align: center;
      min-height: 72px;
    }}
    .summary-label {{
      font-size: 17px;
      font-weight: 700;
      color: #27272a;
      margin-bottom: 6px;
    }}
    .summary-value {{
      font-size: 18px;
      color: #3f3f46;
    }}
    .footer {{
      text-align: center;
      font-size: 13px;
      color: #71717a;
      margin-top: 10px;
    }}
    @media (max-width: 980px) {{
      .masthead h1 {{ font-size: 30px; }}
      .report-panel h2 {{ font-size: 24px; }}
      .axis-left, .axis-right {{ font-size: 16px; }}
      .x-axis, .point-label {{ font-size: 14px; }}
      .summary-grid {{ grid-template-columns: repeat(3, minmax(0, 1fr)); }}
    }}
    @media (max-width: 640px) {{
      .page {{ padding: 12px 10px 28px; }}
      .masthead {{ padding: 16px; }}
      .masthead h1 {{ font-size: 24px; }}
      .masthead .meta {{ font-size: 15px; }}
      .report-panel {{ padding: 14px 10px; }}
      .report-panel h2 {{ font-size: 20px; }}
      .summary-grid {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
      .legend {{ font-size: 15px; gap: 16px; }}
    }}
  </style>
</head>
<body>
  <div class="page">
    <section class="masthead">
      <div class="masthead-top">
        <h1>Trace WBR Dashboard</h1>
        <div class="meta">{html.escape(str(latest_week))} | {html.escape(str(current_month))} | {html.escape(str(current_quarter))}</div>
      </div>
      <div class="sub">Source: {html.escape(workbook_path.name)} | Domain: https://wbr-dashboard.tracecorp.co.kr | Generated: {generated_at}</div>
    </section>

    {''.join(sections)}

    <div class="footer">Auto-generated from the latest WBR master workbook.</div>
  </div>
  <script type="application/json" id="wbr-data">{html.escape(json.dumps(payload, ensure_ascii=False))}</script>
</body>
</html>
"""


def main() -> None:
    workbook_path = find_latest_workbook()
    dashboard = build_dashboard(workbook_path)
    for output_path in OUTPUT_PATHS:
        output_path.write_text(dashboard, encoding="utf-8")
        print(output_path)


if __name__ == "__main__":
    main()
