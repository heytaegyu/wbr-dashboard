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
CURRENT_WORKBOOK_PATTERN = re.compile(r"^2026\s+.*WBR\.xlsx$", re.IGNORECASE)

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
PAGE_BG = "#f7f7f5"

WEEKLY_COLS = ["B", "C", "D", "E", "F", "G"]
MONTHLY_COLS = ["M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X"]

MONEY_METRICS = {
    "순매출",
    "매출총이익",
    "공헌이익",
    "영업이익",
    "영업현금흐름 (OCF)",
    "영업현금흐름",
    "기타 변동비",
    "고정비",
    "현금잔액",
    "내부 마케팅비(광고)",
    "외부 마케팅비(체험단, 슬롯)",
    "실 객단가 (AOV)",
    "주문당 마케팅비",
}

PERCENT_METRICS = {
    "공헌이익률",
    "영업이익률",
    "실 전환율 (CVR)",
    "TACOS (총 마케팅비%)",
    "체험단 ROI",
}

RATIO_METRICS = set()

DAYS_METRICS = {"└ 재고 확보일수 (DOS)"}

COUNT_METRICS = {
    "순방문자수",
    "순주문수",
    "└ 체험단 주문수",
    "재고수량",
}

SECTION_HEADER_ROWS = range(7, 34)
SUMMARY_LABELS = ("LastWk", "WoW", "6W Avg", "MTD", "QTD", "YTD", "MoM")
TITLE_OVERRIDES = {
    "영업현금흐름 (OCF)": "영업현금흐름 (OCF)",
    "실 전환율 (CVR)": "실 전환율 (CVR)",
    "실 객단가 (AOV)": "실 객단가 (AOV)",
    "TACOS (총 마케팅비%)": "TACOS",
    "└ 체험단 주문수": "체험단 주문수",
    "└ 재고 확보일수 (DOS)": "재고 확보일수 (DOS)",
}

DASHBOARD_SECTIONS = (
    (
        "재무 결과",
        (
            ("순매출", ("순매출",)),
            ("매출총이익", ("매출총이익",)),
            ("공헌이익", ("공헌이익",)),
            ("공헌이익률", ("공헌이익률",)),
            ("영업이익", ("영업이익",)),
            ("영업이익률", ("영업이익률",)),
            ("영업현금흐름 (OCF)", ("영업현금흐름 (OCF)", "영업현금흐름")),
            ("기타 변동비", ("기타 변동비",)),
            ("고정비", ("고정비",)),
            ("현금잔액", ("현금잔액",)),
        ),
    ),
    (
        "마케팅. 유입 / 전환",
        (
            ("내부 마케팅비(광고)", ("내부 마케팅비(광고)",)),
            ("외부 마케팅비(체험단, 슬롯)", ("외부 마케팅비(체험단, 슬롯)",)),
            ("순방문자수", ("순방문자수",)),
            ("실 전환율 (CVR)", ("실 전환율 (CVR)",)),
            ("순주문수", ("순주문수",)),
            ("└ 체험단 주문수", ("└ 체험단 주문수", "체험단 주문수")),
            ("실 객단가 (AOV)", ("실 객단가 (AOV)",)),
            ("TACOS (총 마케팅비%)", ("TACOS (총 마케팅비%)",)),
            ("체험단 ROI", ("체험단 ROI",)),
            ("주문당 마케팅비", ("주문당 마케팅비",)),
        ),
    ),
    (
        "풀필먼트: 물류",
        (
            ("재고수량", ("재고수량",)),
            ("└ 재고 확보일수 (DOS)", ("└ 재고 확보일수 (DOS)", "재고 확보일수 (DOS)")),
        ),
    ),
)


def find_latest_workbook() -> Path:
    current_candidates = [
        path
        for path in SHARED_DRIVE_DIR.glob("*.xlsx")
        if not path.name.startswith("~$") and CURRENT_WORKBOOK_PATTERN.fullmatch(path.name)
    ]
    if current_candidates:
        return max(current_candidates, key=lambda path: path.stat().st_mtime)

    candidates = []
    for path in SHARED_DRIVE_DIR.glob("Trace_WBR_Master_W*.xlsx"):
        match = WORKBOOK_PATTERN.fullmatch(path.name)
        if match:
            candidates.append((int(match.group(1)), path))
    if not candidates:
        raise FileNotFoundError(f"No 2026 WBR or Trace_WBR_Master_W*.xlsx file found in {SHARED_DRIVE_DIR}")
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


def metric_kind(metric: str) -> str:
    if metric in MONEY_METRICS:
        return "money"
    if metric in PERCENT_METRICS:
        return "percent"
    if metric in RATIO_METRICS:
        return "ratio"
    if metric in DAYS_METRICS:
        return "days"
    return "count"


def money_decimals(value: Optional[float]) -> int:
    if value is None:
        return 0
    scaled = abs(value)
    if scaled >= 100:
        return 0
    if scaled >= 10:
        return 1
    return 2


def scale_value(metric: str, value: Optional[float]) -> Optional[float]:
    if value is None:
        return None
    if metric_kind(metric) == "money":
        return value / 1000
    return value


def format_scaled(metric: str, value: Optional[float], decimals: int = 0) -> str:
    if value is None:
        return "-"
    scaled = scale_value(metric, value)
    kind = metric_kind(metric)
    if kind == "percent":
        return f"{scaled * 100:.1f}%"
    if kind == "ratio":
        return f"{scaled:.2f}x"
    if kind == "days":
        return f"{scaled:.0f}일"
    if kind == "money":
        final_decimals = decimals if decimals else money_decimals(scaled)
        return f"{scaled:.{final_decimals}f}K"
    return f"{scaled:,.0f}"


def format_axis(metric: str, value: float) -> str:
    kind = metric_kind(metric)
    if kind == "percent":
        return f"{value * 100:.0f}%"
    if kind == "ratio":
        return f"{value:.1f}x"
    if kind == "days":
        return f"{value:.0f}"
    if kind == "money":
        return f"{value:.{money_decimals(value)}f}K"
    if abs(value) >= 1000:
        return f"{value / 1000:.1f}K"
    return f"{value:.0f}"


def format_point(metric: str, value: Optional[float]) -> str:
    if value is None:
        return "-"
    scaled = scale_value(metric, value)
    kind = metric_kind(metric)
    if kind == "percent":
        return f"{scaled * 100:.1f}"
    if kind == "ratio":
        return f"{scaled:.2f}"
    if kind == "days":
        return f"{scaled:.0f}"
    if kind == "money":
        return f"{scaled:.{money_decimals(scaled)}f}"
    return f"{scaled:,.0f}"


def format_change(value: Optional[float]) -> str:
    if value is None:
        return "-"
    return f"{value * 100:+.1f}%"


def nice_bounds(values: List[float], include_zero: bool = False) -> tuple[float, float]:
    if not values:
        return (0, 1)

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
    return v_min - padding, v_max + padding


def rolling_average(series: List[Optional[float]], window: int = 3) -> List[Optional[float]]:
    values: List[Optional[float]] = []
    for idx, value in enumerate(series):
        if value is None:
            values.append(None)
            continue
        window_values = [item for item in series[max(0, idx - window + 1) : idx + 1] if item is not None]
        values.append(sum(window_values) / len(window_values))
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


def slugify(text: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", text.lower())
    return slug.strip("-") or "section"


def title_text(metric: str, index: int) -> str:
    cleaned = TITLE_OVERRIDES.get(metric, metric.replace("└ ", "").strip())
    kind = metric_kind(metric)
    if kind == "money":
        unit = "(₩ K)"
    elif kind == "percent":
        unit = "(%)"
    elif kind == "ratio":
        unit = "(x)"
    elif kind == "days":
        unit = "(일)"
    else:
        unit = ""
    suffix = f" {unit}" if unit else ""
    return f"{index}. {cleaned}{suffix}"


def build_section_blocks(row_map: Dict[int, Dict[str, Optional[str]]], month_index: int) -> List[Dict[str, object]]:
    sections: List[Dict[str, object]] = []
    metric_index = 1

    rows_by_label = {
        (row.get("A") or "").strip(): row
        for row in row_map.values()
        if (row.get("A") or "").strip()
    }

    for section_index, (section_title, metrics) in enumerate(DASHBOARD_SECTIONS, start=1):
        current_section = {"title": section_title, "id": f"section-{section_index}", "metrics": []}
        for display_metric, aliases in metrics:
            row = next((rows_by_label[alias] for alias in aliases if alias in rows_by_label), {})
            payload = {
                "metric": display_metric,
                "metricId": f"metric-{metric_index}",
                "title": title_text(display_metric, metric_index),
                "weekly": [parse_number(row.get(col)) for col in WEEKLY_COLS],
                "monthly": [parse_number(row.get(col)) for col in MONTHLY_COLS],
                "wow": parse_number(row.get("H")),
                "avg6": parse_number(row.get("I")),
                "qtd": parse_number(row.get("J")),
                "ytd": parse_number(row.get("K")),
                "mom": parse_number(row.get("L")),
                "mtd": parse_number(row.get(MONTHLY_COLS[month_index])) if 0 <= month_index < len(MONTHLY_COLS) else None,
            }
            current_section["metrics"].append(payload)
            metric_index += 1
        sections.append(current_section)

    return [section for section in sections if section["metrics"]]


def chart_svg(metric: str, weekly_labels: List[str], month_labels: List[str], weekly: List[Optional[float]], monthly: List[Optional[float]]) -> str:
    scaled_weekly = [scale_value(metric, value) for value in weekly]
    scaled_monthly = [scale_value(metric, value) for value in monthly]
    trend_weekly = [scale_value(metric, value) for value in rolling_average(weekly)]
    trend_monthly = [scale_value(metric, value) for value in rolling_average(monthly)]

    weekly_values = [value for value in scaled_weekly + trend_weekly if value is not None]
    monthly_values = [value for value in scaled_monthly + trend_monthly if value is not None]

    include_zero_weekly = any((value or 0) < 0 for value in weekly_values)
    include_zero_monthly = any((value or 0) < 0 for value in monthly_values)
    left_min, left_max = nice_bounds(weekly_values, include_zero=include_zero_weekly)
    right_min, right_max = nice_bounds(monthly_values, include_zero=include_zero_monthly)

    width = 1320
    height = 500
    left = 130
    right = 1160
    top = 58
    bottom = 352
    weekly_end = 430
    monthly_start = 520

    def weekly_x(idx: int) -> float:
        if len(weekly_labels) == 1:
            return (left + weekly_end) / 2
        return left + (weekly_end - left) * idx / (len(weekly_labels) - 1)

    def monthly_x(idx: int) -> float:
        if len(month_labels) == 1:
            return (monthly_start + right) / 2
        return monthly_start + (right - monthly_start) * idx / (len(month_labels) - 1)

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
    elif right_min < 0 < right_max:
        zero_y = y_right(0)
        grid.append(f'<line x1="{left}" y1="{zero_y:.1f}" x2="{right}" y2="{zero_y:.1f}" class="zero-line" />')

    weekly_actual_points = [(weekly_x(idx), y_left(value)) for idx, value in enumerate(scaled_weekly) if value is not None]
    weekly_trend_points = [(weekly_x(idx), y_left(value)) for idx, value in enumerate(trend_weekly) if value is not None]
    monthly_actual_points = [(monthly_x(idx), y_right(value)) for idx, value in enumerate(scaled_monthly) if value is not None]
    monthly_trend_points = [(monthly_x(idx), y_right(value)) for idx, value in enumerate(trend_monthly) if value is not None]

    x_labels = []
    for idx, label in enumerate(weekly_labels):
        x = weekly_x(idx)
        x_labels.append(
            f'<text x="{x:.1f}" y="{bottom + 50}" transform="rotate(-32 {x:.1f} {bottom + 50})" text-anchor="end" class="x-axis">{html.escape(label)}</text>'
        )
    for idx, label in enumerate(month_labels):
        x = monthly_x(idx)
        x_labels.append(
            f'<text x="{x:.1f}" y="{bottom + 50}" transform="rotate(-32 {x:.1f} {bottom + 50})" text-anchor="end" class="x-axis">{html.escape(label)}</text>'
        )

    point_labels = []
    for idx, value in enumerate(weekly):
        if value is None:
            continue
        x = weekly_x(idx)
        y = y_left(scale_value(metric, value))
        point_labels.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="8" class="actual-dot" />')
        point_labels.append(
            f'<text x="{x:.1f}" y="{y - 18:.1f}" text-anchor="middle" class="point-label">{html.escape(format_point(metric, value))}</text>'
        )
    for idx, value in enumerate(monthly):
        if value is None:
            continue
        x = monthly_x(idx)
        y = y_right(scale_value(metric, value))
        point_labels.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="8" class="actual-dot" />')
        point_labels.append(
            f'<text x="{x:.1f}" y="{y - 18:.1f}" text-anchor="middle" class="point-label">{html.escape(format_point(metric, value))}</text>'
        )

    return f"""
    <svg viewBox="0 0 {width} {height}" class="chart" role="img" aria-label="{html.escape(metric)} chart">
      {''.join(grid)}
      <path d="{path_from_points(weekly_trend_points)}" class="trend-line" />
      <path d="{path_from_points(monthly_trend_points)}" class="trend-line" />
      <path d="{path_from_points(weekly_actual_points)}" class="actual-line" />
      <path d="{path_from_points(monthly_actual_points)}" class="actual-line" />
      <line x1="{weekly_end + 36}" y1="{top}" x2="{weekly_end + 36}" y2="{bottom + 12}" class="split-line" />
      {''.join(point_labels)}
      {''.join(x_labels)}
    </svg>
    """


def summary_row(metric: str, payload: Dict[str, Optional[float]]) -> str:
    items = (
        ("LastWk", format_scaled(metric, payload["weekly"][-1])),
        ("WoW", format_change(payload["wow"])),
        ("6W Avg", format_scaled(metric, payload["avg6"])),
        ("MTD", format_scaled(metric, payload["mtd"])),
        ("QTD", format_scaled(metric, payload["qtd"])),
        ("YTD", format_scaled(metric, payload["ytd"])),
        ("MoM", format_change(payload["mom"])),
    )
    return "".join(
        f"""
        <div class="summary-item">
          <div class="summary-label">{label}</div>
          <div class="summary-value">{html.escape(value)}</div>
        </div>
        """
        for label, value in items
    )


def find_metric_payload(sections: List[Dict[str, object]], metric: str) -> Optional[Dict[str, object]]:
    for section in sections:
        for metric_payload in section["metrics"]:
            if metric_payload["metric"] == metric:
                return metric_payload
    return None


def format_weekly_summary_value(metric: str, value: float) -> str:
    if metric_kind(metric) == "money" and value < 0:
        return f"({format_scaled(metric, abs(value), 1)})"
    return format_scaled(metric, value, 1)


def build_weekly_summary(sections: List[Dict[str, object]], inventory_asset: Optional[float]) -> str:
    parts = []
    for metric in ("순매출", "영업이익", "체험단 ROI"):
        metric_payload = find_metric_payload(sections, metric)
        if not metric_payload:
            continue
        latest_value = metric_payload["weekly"][-1]
        if latest_value is not None:
            parts.append(f"{metric} {format_weekly_summary_value(metric, latest_value)}")
    if inventory_asset is not None:
        parts.append(f"재고자산 {(inventory_asset / 1000000):.1f}M")
    return "이번 주: " + " · ".join(parts) if parts else ""


def render_panel(metric_payload: Dict[str, object], weekly_labels: List[str], month_labels: List[str]) -> str:
    metric = metric_payload["metric"]
    return f"""
      <article class="report-panel" id="{metric_payload['metricId']}">
        <h3>{html.escape(metric_payload['title'])}</h3>
{chart_svg(metric, weekly_labels, month_labels, metric_payload['weekly'], metric_payload['monthly'])}
        <div class="legend">
          <span class="legend-item"><span class="legend-line actual"></span>Actual</span>
          <span class="legend-item"><span class="legend-line trend"></span>Trend</span>
        </div>
        <div class="summary-grid">
          {summary_row(metric, metric_payload)}
        </div>
      </article>
    """


def build_dashboard(workbook_path: Path) -> str:
    rows = load_sheet_rows(workbook_path, "WBR Dashboard")
    row_map = {idx + 1: row for idx, row in enumerate(rows)}

    weekly_labels = [(row_map[6].get(col, "") or "").strip() for col in WEEKLY_COLS]
    month_labels = [(row_map[6].get(col, "") or "").strip() for col in MONTHLY_COLS]
    latest_week = (row_map[4].get("B", "") or weekly_labels[-1]).strip()
    current_month = (row_map[4].get("F", "") or "-").strip()
    current_quarter = (row_map[4].get("J", "") or "-").strip()
    runway = parse_number(row_map[4].get("O"))
    inventory_asset = parse_number(row_map[4].get("W"))

    month_match = re.search(r"(\d+)", current_month)
    month_index = max(0, min(int(month_match.group(1)) - 1, 11)) if month_match else 0

    sections = build_section_blocks(row_map, month_index)
    total_metrics = sum(len(section["metrics"]) for section in sections)
    weekly_summary = build_weekly_summary(sections, inventory_asset)

    section_nav = "".join(
        f'<a href="#{section["id"]}" class="section-chip">{html.escape(section["title"])}</a>'
        for section in sections
    )

    rendered_sections = []
    for section in sections:
        panels = "".join(render_panel(metric_payload, weekly_labels, month_labels) for metric_payload in section["metrics"])
        rendered_sections.append(
            f"""
            <section class="report-section" id="{section['id']}">
              <div class="section-heading">
                <span>{html.escape(section['title'])}</span>
                <small>{len(section['metrics'])} metrics</small>
              </div>
              <div class="report-grid">
                {panels}
              </div>
            </section>
            """
        )

    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")
    payload = {
        "weeklyLabels": weekly_labels,
        "monthLabels": month_labels,
        "latestWeek": latest_week,
        "month": current_month,
        "quarter": current_quarter,
        "runway": runway,
        "inventoryAsset": inventory_asset,
        "sections": sections,
    }

    runway_text = f"{runway:.1f}개월" if runway is not None else "-"
    inventory_asset_text = f"{(inventory_asset / 1000000):.1f}M" if inventory_asset is not None else "-"

    return f"""<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Trace WBR Dashboard</title>
  <link rel="icon" type="image/svg+xml" href="/favicon.svg" />
  <style>
    * {{ box-sizing: border-box; }}
    html {{ scroll-behavior: smooth; }}
    body {{
      margin: 0;
      background: {PAGE_BG};
      color: {TEXT};
      font-family: "Avenir Next", "Pretendard", "Apple SD Gothic Neo", sans-serif;
    }}
    .page {{
      max-width: 1560px;
      margin: 0 auto;
      padding: 22px 16px 52px;
    }}
    .masthead {{
      border: 2px solid {BORDER};
      background: #fff;
      padding: 24px 26px;
      margin-bottom: 18px;
    }}
    .masthead-top {{
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      gap: 18px;
      flex-wrap: wrap;
    }}
    .masthead-copy h1 {{
      margin: 0;
      font-size: 38px;
      color: #27272a;
    }}
    .masthead-copy p {{
      margin: 10px 0 0;
      font-size: 17px;
      color: #52525b;
    }}
    .week-summary {{
      margin-top: 8px;
      font-size: 15px;
      font-weight: 700;
      color: #27272a;
    }}
    .masthead-meta {{
      display: flex;
      gap: 28px;
      flex-wrap: wrap;
      align-items: flex-start;
      justify-content: flex-end;
    }}
    .meta-stat {{
      min-width: 112px;
      text-align: right;
    }}
    .meta-stat .label {{
      font-size: 12px;
      font-weight: 700;
      letter-spacing: 0.08em;
      text-transform: uppercase;
      color: #71717a;
    }}
    .meta-stat .value {{
      margin-top: 6px;
      font-size: 28px;
      font-weight: 700;
      color: #27272a;
    }}
    .meta-stat .subvalue {{
      margin-top: 4px;
      font-size: 14px;
      color: #71717a;
    }}
    .masthead-sub {{
      margin-top: 14px;
      font-size: 14px;
      color: #71717a;
    }}
    .section-nav {{
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
      margin-bottom: 18px;
    }}
    .section-chip {{
      display: inline-flex;
      align-items: center;
      min-height: 38px;
      padding: 0 14px;
      border: 1.5px solid {BORDER};
      border-radius: 8px;
      background: #fff;
      color: #27272a;
      text-decoration: none;
      font-size: 14px;
      font-weight: 600;
    }}
    .report-section {{
      margin-bottom: 22px;
    }}
    .section-heading {{
      display: flex;
      justify-content: space-between;
      align-items: baseline;
      gap: 12px;
      padding: 0 4px 10px;
      margin-bottom: 12px;
      border-bottom: 2px solid #d4d4d8;
    }}
    .section-heading span {{
      font-size: 28px;
      font-weight: 700;
      color: #27272a;
    }}
    .section-heading small {{
      font-size: 13px;
      color: #71717a;
      text-transform: uppercase;
      letter-spacing: 0.08em;
    }}
    .report-grid {{
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 16px;
    }}
    .report-panel {{
      background: #fff;
      border: 2px solid {BORDER};
      padding: 20px 16px 18px;
    }}
    .report-panel h3 {{
      margin: 0 0 4px;
      text-align: center;
      font-size: 30px;
      line-height: 1.2;
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
      stroke: #a1a1aa;
      stroke-width: 2;
      stroke-dasharray: 8 8;
    }}
    .split-line {{
      stroke: #cbd5e1;
      stroke-width: 1.5;
      stroke-dasharray: 5 6;
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
      grid-template-columns: repeat(7, minmax(0, 1fr));
      gap: 8px;
      padding-top: 14px;
      border-top: 1px solid #d4d4d8;
    }}
    .summary-item {{
      text-align: center;
      min-height: 72px;
    }}
    .summary-label {{
      font-size: 16px;
      font-weight: 700;
      color: #27272a;
      margin-bottom: 6px;
    }}
    .summary-value {{
      font-size: 18px;
      color: #3f3f46;
    }}
    .footer {{
      margin-top: 10px;
      text-align: center;
      font-size: 13px;
      color: #71717a;
    }}
    @media (max-width: 1260px) {{
      .report-grid {{ grid-template-columns: 1fr; }}
      .summary-grid {{ grid-template-columns: repeat(4, minmax(0, 1fr)); }}
    }}
    @media (max-width: 980px) {{
      .masthead-copy h1 {{ font-size: 30px; }}
      .meta-stat .value {{ font-size: 22px; }}
      .section-heading span {{ font-size: 24px; }}
      .report-panel h3 {{ font-size: 24px; }}
      .axis-left, .axis-right {{ font-size: 16px; }}
      .x-axis, .point-label {{ font-size: 14px; }}
    }}
    @media (max-width: 640px) {{
      .page {{ padding: 12px 10px 28px; }}
      .masthead {{ padding: 18px 16px; }}
      .masthead-copy h1 {{ font-size: 24px; }}
      .masthead-copy p {{ font-size: 15px; }}
      .masthead-meta {{
        width: 100%;
        gap: 16px;
        justify-content: flex-start;
      }}
      .meta-stat {{
        min-width: 0;
        text-align: left;
      }}
      .report-panel {{ padding: 14px 10px; }}
      .report-panel h3 {{ font-size: 20px; }}
      .legend {{ font-size: 15px; gap: 16px; }}
      .summary-grid {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
      .section-chip {{ font-size: 13px; }}
    }}
  </style>
</head>
<body>
  <div class="page">
    <section class="masthead">
      <div class="masthead-top">
        <div class="masthead-copy">
          <h1>Trace WBR Dashboard</h1>
          <p>{html.escape(latest_week)} 기준으로 {total_metrics}개 핵심 지표를 주간과 월간 흐름으로 한 번에 봅니다.</p>
          <div class="week-summary">{html.escape(weekly_summary)}</div>
        </div>
        <div class="masthead-meta">
          <div class="meta-stat">
            <div class="label">Month</div>
            <div class="value">{html.escape(current_month)}</div>
            <div class="subvalue">{html.escape(current_quarter)}</div>
          </div>
          <div class="meta-stat">
            <div class="label">Runway</div>
            <div class="value">{html.escape(runway_text)}</div>
            <div class="subvalue">운영비 포함</div>
          </div>
          <div class="meta-stat">
            <div class="label">Inventory Asset</div>
            <div class="value">{html.escape(inventory_asset_text)}</div>
            <div class="subvalue">KRW</div>
          </div>
        </div>
      </div>
      <div class="masthead-sub">Source: {html.escape(workbook_path.name)} | Domain: https://wbr-dashboard.tracecorp.co.kr | Generated: {generated_at}</div>
    </section>

    <nav class="section-nav" aria-label="Dashboard Sections">
      {section_nav}
    </nav>

    {''.join(rendered_sections)}

    <div class="footer">Auto-generated from the latest WBR dashboard workbook.</div>
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
