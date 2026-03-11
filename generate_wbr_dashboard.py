import html
import json
import math
import re
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET


WORKBOOK_PATH = Path("/Users/taegyu/Downloads/2026 트레이스 WBR.xlsx")
OUTPUT_PATHS = [
    Path("/Users/taegyu/Documents/New project/wbr_dashboard.html"),
    Path("/Users/taegyu/Documents/New project/index.html"),
]

NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def load_sheet_rows(path: Path) -> List[Dict[str, Optional[str]]]:
    with zipfile.ZipFile(path) as zf:
        shared_strings: List[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            shared_root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in shared_root.findall("a:si", NS):
                shared_strings.append("".join(t.text or "" for t in si.iterfind(".//a:t", NS)))

        sheet_root = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
        rows: dict[int, dict[str, str | None]] = {}
        for cell in sheet_root.iterfind(".//a:c", NS):
            ref = cell.attrib["r"]
            col = re.match(r"([A-Z]+)", ref).group(1)
            row_num = int(re.search(r"(\d+)$", ref).group(1))
            cell_type = cell.attrib.get("t")
            value_node = cell.find("a:v", NS)

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
    if text in {"", "-", "#ERROR!", "#VALUE!"}:
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


def format_value(key: str, value: Optional[float]) -> str:
    if value is None:
        return "-"
    if key == "기여이익률":
        return f"{value * 100:.0f}%"
    if key in {"재고소진일"}:
        return f"{value:.0f}주"
    if key in {"일 평균 순판매량", "순판매량(체험단 제외)", "리뷰 체험단"}:
        return f"{value:.0f}"
    return f"{value:,.0f}"


def delta_text(key: str, current: Optional[float], previous: Optional[float]) -> Tuple[str, str]:
    if current is None or previous is None:
        return "-", "flat"
    delta = current - previous
    if abs(delta) < 1e-9:
        return "변화 없음", "flat"
    trend = "up" if delta > 0 else "down"
    sign = "+" if delta > 0 else ""
    if key == "기여이익률":
        return f"{sign}{delta * 100:.1f}%p", trend
    if key == "재고소진일":
        return f"{sign}{delta:.0f}주", trend
    return f"{sign}{delta:,.0f}", trend


def chart_svg(labels: List[str], series: List[Optional[float]], color: str, percent: bool = False, zero_line: bool = False) -> str:
    width = 420
    height = 220
    pad_l = 46
    pad_r = 20
    pad_t = 18
    pad_b = 38

    values = [v for v in series if v is not None]
    if not values:
        return f'<svg viewBox="0 0 {width} {height}" class="chart"></svg>'

    v_min = min(values)
    v_max = max(values)
    if zero_line:
        v_min = min(v_min, 0)
        v_max = max(v_max, 0)
    if math.isclose(v_min, v_max):
        v_min -= 1
        v_max += 1

    def x_pos(idx: int) -> float:
        usable = width - pad_l - pad_r
        return pad_l + usable * idx / max(len(labels) - 1, 1)

    def y_pos(value: float) -> float:
        usable = height - pad_t - pad_b
        return pad_t + (v_max - value) * usable / (v_max - v_min)

    grid = []
    for i in range(4):
        y_val = v_min + (v_max - v_min) * i / 3
        y = y_pos(y_val)
        label = f"{y_val * 100:.0f}%" if percent else f"{y_val:,.0f}"
        grid.append(f'<line x1="{pad_l}" y1="{y:.1f}" x2="{width - pad_r}" y2="{y:.1f}" class="grid" />')
        grid.append(f'<text x="{pad_l - 8}" y="{y + 4:.1f}" text-anchor="end" class="axis">{html.escape(label)}</text>')

    if zero_line and v_min < 0 < v_max:
        zero_y = y_pos(0)
        grid.append(f'<line x1="{pad_l}" y1="{zero_y:.1f}" x2="{width - pad_r}" y2="{zero_y:.1f}" class="zero" />')

    points = []
    path_cmds = []
    started = False
    for idx, value in enumerate(series):
        if value is None:
            started = False
            continue
        x = x_pos(idx)
        y = y_pos(value)
        path_cmds.append(("M" if not started else "L") + f"{x:.1f},{y:.1f}")
        points.append((x, y, value))
        started = True

    labels_svg = [
        f'<text x="{x_pos(i):.1f}" y="{height - 14}" text-anchor="middle" class="axis xlab">{html.escape(label)}</text>'
        for i, label in enumerate(labels)
    ]

    dots = []
    for x, y, value in points:
        text = f"{value * 100:.0f}%" if percent else f"{value:,.0f}"
        dots.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="4" fill="{color}" />')
        dots.append(f'<text x="{x:.1f}" y="{y - 10:.1f}" text-anchor="middle" class="point-label">{html.escape(text)}</text>')

    return f"""
    <svg viewBox="0 0 {width} {height}" class="chart" role="img">
      {''.join(grid)}
      <path d="{' '.join(path_cmds)}" fill="none" stroke="{color}" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round" />
      {''.join(dots)}
      {''.join(labels_svg)}
    </svg>
    """


def build_dashboard() -> str:
    rows = load_sheet_rows(WORKBOOK_PATH)
    labels = [rows[1].get(col, "") for col in ["D", "E", "F", "G", "H", "I"]]

    metric_rows = rows[2:12]
    data: Dict[str, List[Optional[float]]] = {}
    for row in metric_rows:
        key = row.get("C")
        if not key:
            continue
        data[key] = [parse_number(row.get(col)) for col in ["D", "E", "F", "G", "H", "I"]]

    cards = []
    for key in ["멜라체리 재고량", "순매출(체험단 포함)", "기여이익(체험단 제외)", "기여이익률", "잉여현금흐름"]:
        series = data[key]
        latest = next((v for v in reversed(series) if v is not None), None)
        prev = next((v for v in reversed(series[:-1]) if v is not None), None)
        delta, trend = delta_text(key, latest, prev)
        cards.append(
            f"""
            <div class="card stat-card">
              <div class="eyebrow">{html.escape(key)}</div>
              <div class="stat">{html.escape(format_value(key, latest))}</div>
              <div class="delta {trend}">직전 주 대비 {html.escape(delta)}</div>
            </div>
            """
        )

    charts = [
        ("재고 추이", "멜라체리 재고량", "#1d4ed8", False, False),
        ("판매량 추이", "순판매량(체험단 제외)", "#ea580c", False, False),
        ("일 평균 판매량", "일 평균 순판매량", "#7c3aed", False, False),
        ("매출 추이", "순매출(체험단 포함)", "#0f766e", False, False),
        ("기여이익과 이익률", "기여이익(체험단 제외)", "#dc2626", False, False),
        ("현금흐름", "잉여현금흐름", "#111827", False, True),
        ("체험단 운영", "리뷰 체험단", "#ca8a04", False, False),
        ("재고소진일", "재고소진일", "#2563eb", False, False),
    ]

    chart_blocks = []
    for title, key, color, percent, zero_line in charts:
        extra = ""
        if title == "기여이익과 이익률":
            extra = chart_svg(labels, data["기여이익률"], "#9333ea", True, False)
        chart_blocks.append(
            f"""
            <section class="card panel">
              <div class="panel-head">
                <h3>{html.escape(title)}</h3>
                <p>{html.escape(key)}</p>
              </div>
              {chart_svg(labels, data[key], color, percent, zero_line)}
              {extra}
            </section>
            """
        )

    latest_label = next((label for label, value in zip(reversed(labels), reversed(data["멜라체리 재고량"])) if value is not None), labels[-1])
    payload = {
        "labels": labels,
        "data": data,
    }

    return f"""<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>WBR KPI Dashboard</title>
  <style>
    :root {{
      --bg: #f4f1e8;
      --paper: #fffdf8;
      --ink: #18181b;
      --muted: #6b7280;
      --line: #d6d3d1;
      --accent: #0f766e;
      --up: #0f766e;
      --down: #b91c1c;
      --flat: #6b7280;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "Avenir Next", "Pretendard", "Apple SD Gothic Neo", sans-serif;
      color: var(--ink);
      background:
        radial-gradient(circle at top left, rgba(21, 128, 61, 0.12), transparent 30%),
        radial-gradient(circle at top right, rgba(37, 99, 235, 0.12), transparent 30%),
        var(--bg);
    }}
    .wrap {{
      max-width: 1400px;
      margin: 0 auto;
      padding: 32px 20px 56px;
      position: relative;
    }}
    .aux-link {{
      position: sticky;
      top: 16px;
      z-index: 10;
      display: flex;
      justify-content: flex-end;
      margin-bottom: 8px;
    }}
    .aux-link a {{
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 9px 12px;
      border-radius: 999px;
      text-decoration: none;
      color: #57534e;
      background: rgba(255, 253, 248, 0.82);
      border: 1px solid rgba(24, 24, 27, 0.08);
      box-shadow: 0 8px 20px rgba(24, 24, 27, 0.05);
      font-size: 12px;
      font-weight: 700;
      letter-spacing: 0.02em;
    }}
    .aux-link a::after {{
      content: "준비중";
      display: inline-flex;
      padding: 2px 7px;
      border-radius: 999px;
      background: #efe7da;
      font-size: 11px;
      font-weight: 600;
    }}
    .hero {{
      display: grid;
      grid-template-columns: 1.2fr 0.8fr;
      gap: 20px;
      margin-bottom: 20px;
    }}
    .card {{
      background: rgba(255, 253, 248, 0.88);
      border: 1px solid rgba(24, 24, 27, 0.08);
      border-radius: 22px;
      box-shadow: 0 12px 30px rgba(24, 24, 27, 0.06);
      backdrop-filter: blur(10px);
    }}
    .headline {{
      padding: 26px 28px;
    }}
    .headline h1 {{
      margin: 0 0 10px;
      font-size: clamp(30px, 5vw, 54px);
      line-height: 0.94;
      letter-spacing: -0.04em;
    }}
    .headline p {{
      margin: 0;
      color: var(--muted);
      font-size: 15px;
      line-height: 1.5;
    }}
    .meta {{
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin-top: 18px;
    }}
    .pill {{
      display: inline-flex;
      padding: 8px 12px;
      border-radius: 999px;
      background: #ece8dc;
      font-size: 13px;
    }}
    .snapshot {{
      padding: 26px 28px;
      display: grid;
      align-content: center;
    }}
    .snapshot .label {{
      font-size: 13px;
      color: var(--muted);
      text-transform: uppercase;
      letter-spacing: 0.1em;
    }}
    .snapshot .value {{
      margin-top: 10px;
      font-size: clamp(28px, 5vw, 44px);
      font-weight: 700;
      letter-spacing: -0.04em;
    }}
    .snapshot .sub {{
      margin-top: 6px;
      color: var(--muted);
      font-size: 15px;
    }}
    .stats {{
      display: grid;
      grid-template-columns: repeat(5, minmax(0, 1fr));
      gap: 14px;
      margin-bottom: 20px;
    }}
    .stat-card {{
      padding: 18px;
      min-height: 128px;
    }}
    .eyebrow {{
      color: var(--muted);
      font-size: 12px;
      letter-spacing: 0.04em;
    }}
    .stat {{
      margin-top: 14px;
      font-size: 30px;
      font-weight: 700;
      letter-spacing: -0.04em;
    }}
    .delta {{
      margin-top: 10px;
      font-size: 13px;
    }}
    .delta.up {{ color: var(--up); }}
    .delta.down {{ color: var(--down); }}
    .delta.flat {{ color: var(--flat); }}
    .grid-panels {{
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 16px;
    }}
    .panel {{
      padding: 20px 18px 10px;
    }}
    .panel-head {{
      display: flex;
      align-items: baseline;
      justify-content: space-between;
      gap: 12px;
      padding: 0 8px;
    }}
    .panel-head h3 {{
      margin: 0;
      font-size: 22px;
      letter-spacing: -0.03em;
    }}
    .panel-head p {{
      margin: 0;
      color: var(--muted);
      font-size: 13px;
    }}
    .chart {{
      width: 100%;
      height: auto;
      margin-top: 4px;
    }}
    .grid {{
      stroke: var(--line);
      stroke-width: 1;
    }}
    .zero {{
      stroke: #111827;
      stroke-dasharray: 4 4;
      stroke-width: 1.2;
    }}
    .axis {{
      fill: #78716c;
      font-size: 11px;
    }}
    .xlab {{
      font-size: 10px;
    }}
    .point-label {{
      fill: #292524;
      font-size: 11px;
      font-weight: 600;
    }}
    .footer {{
      margin-top: 18px;
      color: var(--muted);
      font-size: 12px;
    }}
    @media (max-width: 1100px) {{
      .stats {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
      .hero {{ grid-template-columns: 1fr; }}
    }}
    @media (max-width: 760px) {{
      .grid-panels {{ grid-template-columns: 1fr; }}
      .stats {{ grid-template-columns: 1fr; }}
    }}
  </style>
</head>
<body>
  <div class="wrap">
    <div class="aux-link">
      <a href="./claude.html">Claude Code</a>
    </div>
    <section class="hero">
      <div class="card headline">
        <h1>Trace WBR<br />KPI Dashboard</h1>
        <p>주차별 핵심 KPI를 한 화면에서 비교할 수 있도록 재고, 판매, 매출, 수익성, 현금흐름 중심으로 재구성했습니다.</p>
        <div class="meta">
          <span class="pill">기준 주차: {html.escape(latest_label)}</span>
          <span class="pill">기간: {html.escape(labels[0])} ~ {html.escape(labels[-1])}</span>
          <span class="pill">원본: 2026 트레이스 WBR.xlsx</span>
        </div>
      </div>
      <div class="card snapshot">
        <div class="label">현재 상황 요약</div>
        <div class="value">{html.escape(format_value("순매출(체험단 포함)", data["순매출(체험단 포함)"][-1]))}</div>
        <div class="sub">{html.escape(latest_label)} 순매출, 재고 {html.escape(format_value("멜라체리 재고량", data["멜라체리 재고량"][-1]))}, 기여이익률 {html.escape(format_value("기여이익률", data["기여이익률"][-1]))}</div>
      </div>
    </section>

    <section class="stats">
      {''.join(cards)}
    </section>

    <section class="grid-panels">
      {''.join(chart_blocks)}
    </section>

    <div class="footer">오류값이 있는 `3월 2주차` 열은 제외했고, 대시보드는 원본 XLSX 데이터를 직접 파싱해 생성했습니다. 마지막 검증 업데이트: 2026-03-11.</div>
  </div>
  <script type="application/json" id="wbr-data">{html.escape(json.dumps(payload, ensure_ascii=False))}</script>
</body>
</html>
"""


def main() -> None:
    dashboard = build_dashboard()
    for output_path in OUTPUT_PATHS:
        output_path.write_text(dashboard, encoding="utf-8")
        print(output_path)


if __name__ == "__main__":
    main()
