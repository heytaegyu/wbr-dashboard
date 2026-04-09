#!/usr/bin/env python3
"""Send a combined Rocket Growth daily report to Slack."""

from __future__ import annotations

import argparse
import hashlib
import hmac
import json
import math
import os
import sys
import time as time_mod
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta, timezone
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any


PROJECT_ROOT = Path("/Users/taegyu/Documents/New project")
DEFAULT_WEBHOOK_PATH = PROJECT_ROOT / ".codex-local" / "slack-webhook-url"
DEFAULT_CONFIG_PATH = PROJECT_ROOT / "config" / "coupang-rg-alert.json"
COUPANG_API_BASE = "https://api-gateway.coupang.com"
RG_INVENTORY_PATH = (
    "/v2/providers/rg_open_api/apis/api/v1/vendors/{vendor_id}/rg/inventory/summaries"
)
RG_PRODUCTS_PATH = "/v2/providers/seller_api/apis/api/v1/marketplace/seller-products"
RG_ORDERS_PATH = "/v2/providers/rg_open_api/apis/api/v1/vendors/{vendor_id}/rg/orders"
KST = timezone(timedelta(hours=9))
WON_QUANTIZE = Decimal("1")


@dataclass
class EconomicsSpec:
    reference_price: Decimal
    product_cost: Decimal
    inbound_shipping: Decimal
    ad_cost: Decimal
    selling_fee: Decimal
    advance_settlement: Decimal
    vat: Decimal

    @property
    def ad_rate(self) -> Decimal:
        return self.ad_cost / self.reference_price

    @property
    def selling_fee_rate(self) -> Decimal:
        return self.selling_fee / self.reference_price

    @property
    def advance_settlement_rate(self) -> Decimal:
        return self.advance_settlement / self.reference_price

    @property
    def vat_rate(self) -> Decimal:
        return self.vat / self.reference_price


@dataclass
class TargetSkuSpec:
    sku: str
    label: str
    units: int | None
    economics: EconomicsSpec | None
    order_index: int


@dataclass
class AlertItem:
    vendor_item_id: str
    external_sku_id: str | None
    seller_product_name: str
    item_name: str | None
    quantity: int
    sales_last_thirty_days: int
    estimated_days_left: float | None
    low_by_quantity: bool
    low_by_days: bool
    order_index: int
    label: str | None

    @property
    def display_name(self) -> str:
        if self.label:
            return self.label
        if self.item_name:
            return f"{self.seller_product_name} / {self.item_name}"
        return self.seller_product_name


@dataclass
class SalesTotals:
    quantity: int = 0
    gross_sales: Decimal = Decimal("0")


@dataclass
class CostBreakdown:
    gross_sales: Decimal
    product_cost: Decimal
    inbound_shipping: Decimal
    ad_cost: Decimal
    selling_fee: Decimal
    advance_settlement: Decimal
    vat: Decimal
    estimated_profit: Decimal

    @property
    def total_cost(self) -> Decimal:
        return (
            self.product_cost
            + self.inbound_shipping
            + self.ad_cost
            + self.selling_fee
            + self.advance_settlement
            + self.vat
        )


@dataclass
class SalesSummary:
    item: AlertItem
    yesterday: SalesTotals
    today_snapshot: SalesTotals
    economics: EconomicsSpec | None


@dataclass
class AggregateBreakdown:
    gross_sales: Decimal = Decimal("0")
    product_cost: Decimal = Decimal("0")
    inbound_shipping: Decimal = Decimal("0")
    ad_cost: Decimal = Decimal("0")
    selling_fee: Decimal = Decimal("0")
    advance_settlement: Decimal = Decimal("0")
    vat: Decimal = Decimal("0")
    estimated_profit: Decimal = Decimal("0")

    @property
    def total_cost(self) -> Decimal:
        return (
            self.product_cost
            + self.inbound_shipping
            + self.ad_cost
            + self.selling_fee
            + self.advance_settlement
            + self.vat
        )


@dataclass
class ReportContext:
    now_kst: datetime
    yesterday_date: date
    today_date: date
    snapshot_cutoff: datetime
    snapshot_label: str


@dataclass
class OrderLine:
    vendor_item_id: str
    quantity: int
    gross_sales: Decimal
    paid_at_kst: datetime


def parse_args() -> argparse.Namespace:
    config = load_config(DEFAULT_CONFIG_PATH)
    quantity_threshold_default = env_or_config("COUPANG_LOW_STOCK_QTY", config, "quantity_threshold", -1)
    days_threshold_default = env_or_config("COUPANG_LOW_STOCK_DAYS", config, "days_threshold", 28)
    max_items_default = env_or_config("COUPANG_ALERT_MAX_ITEMS", config, "max_items", 20)
    snapshot_hour_default = env_or_config("COUPANG_SNAPSHOT_HOUR_KST", config, "snapshot_hour_kst", 8)
    parser = argparse.ArgumentParser(
        description="Query Coupang Rocket Growth inventory/orders and post a daily report to Slack."
    )
    parser.add_argument(
        "--quantity-threshold",
        type=int,
        default=int(quantity_threshold_default),
        help=(
            "Alert when totalOrderableQuantity is at or below this value. "
            "Use -1 to disable quantity-based alerts."
        ),
    )
    parser.add_argument(
        "--days-threshold",
        type=float,
        default=float(days_threshold_default),
        help="Alert when estimated days left is at or below this value.",
    )
    parser.add_argument(
        "--max-items",
        type=int,
        default=int(max_items_default),
        help="Maximum number of low-stock items to include in Slack.",
    )
    parser.add_argument(
        "--snapshot-hour-kst",
        type=int,
        default=int(snapshot_hour_default),
        help="Today sales are reported up to this KST hour, or now if earlier.",
    )
    parser.add_argument(
        "--target-skus",
        default=str(env_or_config("COUPANG_TARGET_SKUS", {"target_skus_text": config_target_skus(config)}, "target_skus_text", "")),
        help=(
            "Comma-separated externalSkuId values to monitor. "
            "When omitted, all Rocket Growth items are checked."
        ),
    )
    parser.add_argument(
        "--webhook-path",
        default=str(DEFAULT_WEBHOOK_PATH),
        help=f"Path to a file containing a Slack webhook URL. Default: {DEFAULT_WEBHOOK_PATH}",
    )
    parser.add_argument(
        "--send-healthy-summary",
        action="store_true",
        help="Send a Slack message even when no low-stock items are found.",
    )
    parser.add_argument(
        "--skip-product-catalog",
        action="store_true",
        help="Skip the product-list lookup and use vendorItemId/externalSkuId only.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print the Slack payload instead of posting it.",
    )
    return parser.parse_args()


def require_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value


def load_config(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    with path.open("r", encoding="utf-8") as handle:
        data = json.load(handle)
    if not isinstance(data, dict):
        raise RuntimeError(f"Config file must contain a JSON object: {path}")
    return data


def env_or_config(env_name: str, config: dict[str, Any], key: str, fallback: Any) -> Any:
    env_value = os.getenv(env_name)
    if env_value not in (None, ""):
        return env_value
    value = config.get(key, fallback)
    if value is None:
        return fallback
    return value


def decimal_from(value: Any) -> Decimal:
    return Decimal(str(value))


def config_target_skus(config: dict[str, Any]) -> str:
    raw = config.get("target_skus", [])
    if isinstance(raw, list):
        values: list[str] = []
        for entry in raw:
            if isinstance(entry, dict):
                sku_value = str(entry.get("sku", "")).strip()
                if sku_value:
                    values.append(sku_value)
            else:
                sku_value = str(entry).strip()
                if sku_value:
                    values.append(sku_value)
        return ",".join(values)
    if isinstance(raw, str):
        return raw
    return ""


def parse_target_specs(config: dict[str, Any], target_skus_text: str) -> list[TargetSkuSpec]:
    config_entries = config.get("target_skus", [])
    specs: list[TargetSkuSpec] = []
    if isinstance(config_entries, list) and config_entries and isinstance(config_entries[0], dict):
        for index, entry in enumerate(config_entries):
            sku = str(entry.get("sku", "")).strip()
            if not sku:
                continue
            economics = None
            if all(key in entry for key in ("reference_price", "product_cost", "inbound_shipping", "ad_cost", "selling_fee", "advance_settlement", "vat")):
                economics = EconomicsSpec(
                    reference_price=decimal_from(entry["reference_price"]),
                    product_cost=decimal_from(entry["product_cost"]),
                    inbound_shipping=decimal_from(entry["inbound_shipping"]),
                    ad_cost=decimal_from(entry["ad_cost"]),
                    selling_fee=decimal_from(entry["selling_fee"]),
                    advance_settlement=decimal_from(entry["advance_settlement"]),
                    vat=decimal_from(entry["vat"]),
                )
            specs.append(
                TargetSkuSpec(
                    sku=sku,
                    label=str(entry.get("label") or sku),
                    units=int(entry["units"]) if entry.get("units") is not None else None,
                    economics=economics,
                    order_index=index,
                )
            )
        return specs

    text_values = [value.strip() for value in target_skus_text.split(",") if value.strip()]
    return [
        TargetSkuSpec(sku=value, label=value, units=None, economics=None, order_index=index)
        for index, value in enumerate(text_values)
    ]


def read_webhook_url(path_text: str) -> str:
    env_url = os.getenv("SLACK_WEBHOOK_URL", "").strip()
    if env_url:
        return env_url

    path = Path(path_text)
    if not path.exists():
        raise RuntimeError(
            "Slack webhook URL is missing. Set SLACK_WEBHOOK_URL or create "
            f"{path}."
        )
    return path.read_text(encoding="utf-8").strip()


def build_authorization(
    method: str,
    path: str,
    query: str,
    access_key: str,
    secret_key: str,
) -> str:
    signed_date = time_mod.strftime("%y%m%dT%H%M%SZ", time_mod.gmtime())
    message = f"{signed_date}{method}{path}{query}"
    signature = hmac.new(
        secret_key.encode("utf-8"),
        message.encode("utf-8"),
        hashlib.sha256,
    ).hexdigest()
    return (
        "CEA algorithm=HmacSHA256, "
        f"access-key={access_key}, "
        f"signed-date={signed_date}, "
        f"signature={signature}"
    )


def api_get(
    *,
    path: str,
    params: list[tuple[str, Any]],
    access_key: str,
    secret_key: str,
    vendor_id: str,
    keep_empty_keys: set[str] | None = None,
) -> dict[str, Any]:
    keep_empty_keys = keep_empty_keys or set()
    query_pairs: list[tuple[str, str]] = []
    for key, value in params:
        if value in (None, "") and key not in keep_empty_keys:
            continue
        query_pairs.append((key, "" if value is None else str(value)))

    query = urllib.parse.urlencode(query_pairs, doseq=True)
    url = f"{COUPANG_API_BASE}{path}"
    if query:
        url = f"{url}?{query}"

    request = urllib.request.Request(url, method="GET")
    request.add_header("Content-Type", "application/json;charset=UTF-8")
    request.add_header(
        "Authorization",
        build_authorization("GET", path, query, access_key, secret_key),
    )
    request.add_header("X-Requested-By", vendor_id)
    request.add_header("X-EXTENDED-TIMEOUT", "90000")

    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            body = response.read().decode("utf-8")
    except urllib.error.HTTPError as exc:
        error_body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(
            f"Coupang API request failed ({exc.code}) for {path}: {error_body}"
        ) from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Coupang API connection failed for {path}: {exc}") from exc

    data = json.loads(body)
    if isinstance(data, dict) and "code" in data and str(data.get("code")) not in {"200", "SUCCESS"}:
        raise RuntimeError(f"Coupang API returned an error for {path}: {body}")
    return data


def fetch_inventory_summaries(
    *,
    vendor_id: str,
    access_key: str,
    secret_key: str,
) -> list[dict[str, Any]]:
    path = RG_INVENTORY_PATH.format(vendor_id=vendor_id)
    results: list[dict[str, Any]] = []
    next_token: str | None = None

    while True:
        params: list[tuple[str, Any]] = []
        if next_token:
            params.append(("nextToken", next_token))
        response = api_get(
            path=path,
            params=params,
            access_key=access_key,
            secret_key=secret_key,
            vendor_id=vendor_id,
        )
        results.extend(response.get("data", []))
        next_token = response.get("nextToken")
        if not next_token:
            return results


def fetch_product_catalog(
    *,
    vendor_id: str,
    access_key: str,
    secret_key: str,
) -> dict[str, dict[str, str]]:
    results: dict[str, dict[str, str]] = {}
    next_token: str | None = None

    while True:
        params: list[tuple[str, Any]] = [
            ("vendorId", vendor_id),
            ("maxPerPage", 100),
            ("businessTypes", "rocketGrowth"),
        ]
        if next_token:
            params.append(("nextToken", next_token))

        response = api_get(
            path=RG_PRODUCTS_PATH,
            params=params,
            access_key=access_key,
            secret_key=secret_key,
            vendor_id=vendor_id,
        )

        for product in response.get("data", []):
            seller_product_name = product.get("sellerProductName") or "Unknown product"
            for item in product.get("items", []):
                rocket_growth_item = item.get("rocketGrowthItem") or item.get(
                    "rocketGrowthItemData"
                )
                if not rocket_growth_item:
                    continue
                vendor_item_id = rocket_growth_item.get("vendorItemId")
                if vendor_item_id is None:
                    continue
                results[str(vendor_item_id)] = {
                    "seller_product_name": seller_product_name,
                    "item_name": item.get("itemName") or "",
                }

        next_token = response.get("nextToken")
        if not next_token:
            return results


def fetch_rg_orders(
    *,
    vendor_id: str,
    access_key: str,
    secret_key: str,
    paid_date_from: date,
    paid_date_to: date,
) -> list[dict[str, Any]]:
    path = RG_ORDERS_PATH.format(vendor_id=vendor_id)
    results: list[dict[str, Any]] = []
    next_token: str | None = None

    while True:
        params: list[tuple[str, Any]] = [
            ("paidDateFrom", paid_date_from.strftime("%Y%m%d")),
            ("paidDateTo", paid_date_to.strftime("%Y%m%d")),
        ]
        if next_token:
            params.append(("nextToken", next_token))
        response = api_get(
            path=path,
            params=params,
            access_key=access_key,
            secret_key=secret_key,
            vendor_id=vendor_id,
        )
        results.extend(response.get("data", []))
        next_token = response.get("nextToken")
        if not next_token:
            return results


def estimate_days_left(quantity: int, sales_last_thirty_days: int) -> float | None:
    if sales_last_thirty_days <= 0:
        return None
    daily_sales = sales_last_thirty_days / 30
    if daily_sales <= 0:
        return None
    return quantity / daily_sales


def build_monitored_items(
    inventory_rows: list[dict[str, Any]],
    product_catalog: dict[str, dict[str, str]],
    quantity_threshold: int,
    days_threshold: float,
    specs_by_sku: dict[str, TargetSkuSpec],
) -> list[AlertItem]:
    items: list[AlertItem] = []

    for row in inventory_rows:
        vendor_item_id = str(row.get("vendorItemId", ""))
        if not vendor_item_id:
            continue

        external_sku_id = (
            str(row.get("externalSkuId")) if row.get("externalSkuId") is not None else None
        )
        if specs_by_sku and external_sku_id not in specs_by_sku:
            continue

        quantity = int(row.get("inventoryDetails", {}).get("totalOrderableQuantity") or 0)
        sales_last_thirty_days = int(
            row.get("salesCountMap", {}).get("SALES_COUNT_LAST_THIRTY_DAYS") or 0
        )
        estimated_days = estimate_days_left(quantity, sales_last_thirty_days)
        low_by_quantity = quantity_threshold >= 0 and quantity <= quantity_threshold
        low_by_days = estimated_days is not None and estimated_days <= days_threshold
        product_info = product_catalog.get(vendor_item_id, {})
        spec = specs_by_sku.get(external_sku_id or "")

        items.append(
            AlertItem(
                vendor_item_id=vendor_item_id,
                external_sku_id=external_sku_id,
                seller_product_name=product_info.get("seller_product_name")
                or f"vendorItemId {vendor_item_id}",
                item_name=product_info.get("item_name") or None,
                quantity=quantity,
                sales_last_thirty_days=sales_last_thirty_days,
                estimated_days_left=estimated_days,
                low_by_quantity=low_by_quantity,
                low_by_days=low_by_days,
                order_index=spec.order_index if spec else math.inf,
                label=spec.label if spec else None,
            )
        )

    return sorted(items, key=monitored_sort_key)


def monitored_sort_key(item: AlertItem) -> tuple[float, str]:
    return (item.order_index, item.display_name.lower())


def build_alert_items(monitored_items: list[AlertItem]) -> list[AlertItem]:
    return [item for item in monitored_items if item.low_by_quantity or item.low_by_days]


def reason_label(item: AlertItem) -> str:
    reasons: list[str] = []
    if item.low_by_quantity:
        reasons.append("수량 기준")
    if item.low_by_days:
        reasons.append("판매가능일 기준")
    return ", ".join(reasons)


def format_days_left(value: float | None) -> str:
    if value is None:
        return "-"
    if value > 70:
        return "70일+"
    return f"{value:.1f}일"


def build_threshold_summary(quantity_threshold: int, days_threshold: float) -> str:
    parts: list[str] = []
    if quantity_threshold >= 0:
        parts.append(f"재고 {quantity_threshold}개 이하")
    parts.append(f"판매가능 예상일수 {days_threshold:g}일 이하")
    return " 또는 ".join(parts)


def build_report_context(snapshot_hour_kst: int) -> ReportContext:
    if not 0 <= snapshot_hour_kst <= 23:
        raise RuntimeError("snapshot_hour_kst must be between 0 and 23")
    now_kst = datetime.now(timezone.utc).astimezone(KST)
    today_date = now_kst.date()
    yesterday_date = today_date - timedelta(days=1)
    scheduled_cutoff = datetime.combine(today_date, time(snapshot_hour_kst, 0), tzinfo=KST)
    snapshot_cutoff = min(now_kst, scheduled_cutoff)
    snapshot_label = f"오늘 {snapshot_cutoff.strftime('%H:%M')} 기준"
    return ReportContext(
        now_kst=now_kst,
        yesterday_date=yesterday_date,
        today_date=today_date,
        snapshot_cutoff=snapshot_cutoff,
        snapshot_label=snapshot_label,
    )


def money_to_won(amount: Decimal) -> int:
    return int(amount.quantize(WON_QUANTIZE, rounding=ROUND_HALF_UP))


def format_won(amount: Decimal) -> str:
    return f"{money_to_won(amount):,}원"


def estimate_cost_breakdown(totals: SalesTotals, economics: EconomicsSpec | None) -> CostBreakdown | None:
    if economics is None:
        return None

    gross_sales = totals.gross_sales
    quantity = Decimal(totals.quantity)
    product_cost = economics.product_cost * quantity
    inbound_shipping = economics.inbound_shipping * quantity
    ad_cost = (gross_sales * economics.ad_rate).quantize(WON_QUANTIZE, rounding=ROUND_HALF_UP)
    selling_fee = (gross_sales * economics.selling_fee_rate).quantize(WON_QUANTIZE, rounding=ROUND_HALF_UP)
    advance_settlement = (gross_sales * economics.advance_settlement_rate).quantize(
        WON_QUANTIZE, rounding=ROUND_HALF_UP
    )
    vat = (gross_sales * economics.vat_rate).quantize(WON_QUANTIZE, rounding=ROUND_HALF_UP)
    estimated_profit = gross_sales - product_cost - inbound_shipping - ad_cost - selling_fee - advance_settlement - vat
    return CostBreakdown(
        gross_sales=gross_sales,
        product_cost=product_cost,
        inbound_shipping=inbound_shipping,
        ad_cost=ad_cost,
        selling_fee=selling_fee,
        advance_settlement=advance_settlement,
        vat=vat,
        estimated_profit=estimated_profit,
    )


def build_sales_summaries(
    monitored_items: list[AlertItem],
    order_rows: list[dict[str, Any]],
    specs_by_sku: dict[str, TargetSkuSpec],
    report_context: ReportContext,
) -> list[SalesSummary]:
    summary_by_vendor_item: dict[str, SalesSummary] = {}
    for item in monitored_items:
        spec = specs_by_sku.get(item.external_sku_id or "")
        summary_by_vendor_item[item.vendor_item_id] = SalesSummary(
            item=item,
            yesterday=SalesTotals(),
            today_snapshot=SalesTotals(),
            economics=spec.economics if spec else None,
        )

    for row in order_rows:
        paid_at = datetime.fromtimestamp(row.get("paidAt", 0) / 1000, tz=timezone.utc).astimezone(KST)
        order_day = paid_at.date()
        for item in row.get("orderItems", []):
            vendor_item_id = str(item.get("vendorItemId", ""))
            summary = summary_by_vendor_item.get(vendor_item_id)
            if summary is None:
                continue
            quantity = int(item.get("salesQuantity") or 0)
            gross_sales = decimal_from(item.get("unitSalesPrice") or 0) * quantity
            if order_day == report_context.yesterday_date:
                summary.yesterday.quantity += quantity
                summary.yesterday.gross_sales += gross_sales
            if order_day == report_context.today_date and paid_at <= report_context.snapshot_cutoff:
                summary.today_snapshot.quantity += quantity
                summary.today_snapshot.gross_sales += gross_sales

    return sorted(summary_by_vendor_item.values(), key=lambda summary: summary.item.order_index)


def aggregate_costs(breakdowns: list[CostBreakdown | None]) -> AggregateBreakdown | None:
    valid = [entry for entry in breakdowns if entry is not None]
    if not valid:
        return None
    aggregate = AggregateBreakdown()
    for entry in valid:
        aggregate.gross_sales += entry.gross_sales
        aggregate.product_cost += entry.product_cost
        aggregate.inbound_shipping += entry.inbound_shipping
        aggregate.ad_cost += entry.ad_cost
        aggregate.selling_fee += entry.selling_fee
        aggregate.advance_settlement += entry.advance_settlement
        aggregate.vat += entry.vat
        aggregate.estimated_profit += entry.estimated_profit
    return aggregate


def profit_rate_text(profit: Decimal, gross_sales: Decimal) -> str:
    if gross_sales == 0:
        return "-"
    pct = (profit / gross_sales) * Decimal("100")
    return f"{pct.quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)}%"


def build_header_section(alert_items: list[AlertItem], quantity_threshold: int, days_threshold: float) -> str:
    icon = ":warning:" if alert_items else ":white_check_mark:"
    title = "품절임박 항목 있음" if alert_items else "품절임박 대상 없음"
    lines = [
        f"{icon} *쿠팡 로켓그로스 일일 점검*",
        f"기준: {build_threshold_summary(quantity_threshold, days_threshold)}",
        f"상태: {title}",
    ]
    if alert_items:
        lines.append("")
        lines.append("*품절임박 상세*")
        for item in alert_items:
            lines.append(
                "• "
                f"{item.display_name} | 재고: {item.quantity} | 판매가능 예상일수: {format_days_left(item.estimated_days_left)} | 사유: {reason_label(item)}"
            )
    return "\n".join(lines)


def build_stock_section(monitored_items: list[AlertItem]) -> str:
    lines = [":package: *재고 현황*"]
    for item in monitored_items:
        status = "품절임박" if (item.low_by_quantity or item.low_by_days) else "정상"
        lines.append(
            "• "
            f"{item.display_name} | 상태: {status} | 재고: {item.quantity} | 판매가능 예상일수: {format_days_left(item.estimated_days_left)} | 최근30일 판매: {item.sales_last_thirty_days}"
        )
    return "\n".join(lines)


def build_sales_section(sales_summaries: list[SalesSummary], report_context: ReportContext) -> str:
    yesterday_total_qty = sum(summary.yesterday.quantity for summary in sales_summaries)
    yesterday_total_sales = sum(summary.yesterday.gross_sales for summary in sales_summaries)
    today_total_qty = sum(summary.today_snapshot.quantity for summary in sales_summaries)
    today_total_sales = sum(summary.today_snapshot.gross_sales for summary in sales_summaries)

    lines = [":bar_chart: *판매 분석*", f"*어제 {report_context.yesterday_date.isoformat()}*"]
    lines.append(f"• 합계: *{yesterday_total_qty}개 / {format_won(yesterday_total_sales)}*")
    for summary in sales_summaries:
        lines.append(
            f"• {summary.item.display_name}: *{summary.yesterday.quantity}개* / *{format_won(summary.yesterday.gross_sales)}*"
        )
    lines.extend(["", f"_{report_context.snapshot_label} (참고)_"])
    lines.append(f"• 합계: {today_total_qty}개 / {format_won(today_total_sales)}")
    for summary in sales_summaries:
        lines.append(
            f"• {summary.item.display_name}: {summary.today_snapshot.quantity}개 / {format_won(summary.today_snapshot.gross_sales)}"
        )
    return "\n".join(lines)


def build_finance_period_lines(label: str, aggregate: AggregateBreakdown | None, *, secondary: bool = False) -> list[str]:
    if aggregate is None:
        return [f"*{label}*", "• 데이터 없음"]
    heading = f"*{label}*" if not secondary else f"*{label}* _(참고)_"
    return [
        heading,
        f"• 매출 *{format_won(aggregate.gross_sales)}* / 비용 *{format_won(aggregate.total_cost)}* / 이익 *{format_won(aggregate.estimated_profit)}* ({profit_rate_text(aggregate.estimated_profit, aggregate.gross_sales)})",
        (
            "• 비용구성 1: "
            f"제품원가 {format_won(aggregate.product_cost)} / "
            f"입출고+배송비 {format_won(aggregate.inbound_shipping)} / "
            f"광고비 {format_won(aggregate.ad_cost)}"
        ),
        (
            "• 비용구성 2: "
            f"판매수수료 {format_won(aggregate.selling_fee)} / "
            f"선정산 {format_won(aggregate.advance_settlement)} / "
            f"부가세 {format_won(aggregate.vat)}"
        ),
    ]


def build_finance_section(
    sales_summaries: list[SalesSummary],
    report_context: ReportContext,
) -> str:
    yesterday_breakdowns = [estimate_cost_breakdown(summary.yesterday, summary.economics) for summary in sales_summaries]
    today_breakdowns = [estimate_cost_breakdown(summary.today_snapshot, summary.economics) for summary in sales_summaries]
    yesterday_aggregate = aggregate_costs(yesterday_breakdowns)
    today_aggregate = aggregate_costs(today_breakdowns)

    lines = [":moneybag: *매출 / 비용 / 이익*", "원가표 기준 추정값입니다.", ""]
    lines.extend(build_finance_period_lines(f"어제 {report_context.yesterday_date.isoformat()}", yesterday_aggregate))
    lines.extend(["", *build_finance_period_lines(report_context.snapshot_label, today_aggregate, secondary=True)])
    return "\n".join(lines)


def build_payload(
    *,
    alert_items: list[AlertItem],
    monitored_items: list[AlertItem],
    sales_summaries: list[SalesSummary],
    quantity_threshold: int,
    days_threshold: float,
    send_healthy_summary: bool,
    report_context: ReportContext,
) -> dict[str, Any] | None:
    if not monitored_items and not send_healthy_summary:
        return None

    fallback_parts = [f"품절임박 {len(alert_items)}개"]
    total_yesterday = sum(summary.yesterday.quantity for summary in sales_summaries)
    fallback_parts.append(f"어제 판매 {total_yesterday}개")
    fallback_text = "[쿠팡 일일점검] " + " | ".join(fallback_parts)

    blocks = [
        {
            "type": "section",
            "text": {
                "type": "mrkdwn",
                "text": build_header_section(alert_items, quantity_threshold, days_threshold),
            },
        },
        {
            "type": "section",
            "text": {"type": "mrkdwn", "text": build_stock_section(monitored_items)},
        },
        {
            "type": "section",
            "text": {"type": "mrkdwn", "text": build_sales_section(sales_summaries, report_context)},
        },
        {
            "type": "section",
            "text": {"type": "mrkdwn", "text": build_finance_section(sales_summaries, report_context)},
        },
    ]
    return {"text": fallback_text, "blocks": blocks}


def send_slack_payload(webhook_url: str, payload: dict[str, Any]) -> None:
    body = json.dumps(payload).encode("utf-8")
    request = urllib.request.Request(
        webhook_url,
        data=body,
        headers={"Content-Type": "application/json"},
        method="POST",
    )

    try:
        with urllib.request.urlopen(request, timeout=15) as response:
            result = response.read().decode("utf-8").strip()
    except urllib.error.HTTPError as exc:
        error_body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(
            f"Slack webhook request failed ({exc.code}): {error_body}"
        ) from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Slack webhook connection failed: {exc}") from exc

    if response.status != 200 or result.lower() != "ok":
        raise RuntimeError(f"Slack webhook returned {response.status}: {result}")


def main() -> int:
    args = parse_args()

    try:
        config = load_config(DEFAULT_CONFIG_PATH)
        target_specs = parse_target_specs(config, args.target_skus)
        specs_by_sku = {spec.sku: spec for spec in target_specs}

        access_key = require_env("COUPANG_ACCESS_KEY")
        secret_key = require_env("COUPANG_SECRET_KEY")
        vendor_id = require_env("COUPANG_VENDOR_ID")
        webhook_url = read_webhook_url(args.webhook_path)
        report_context = build_report_context(args.snapshot_hour_kst)

        inventory_rows = fetch_inventory_summaries(
            vendor_id=vendor_id,
            access_key=access_key,
            secret_key=secret_key,
        )

        product_catalog: dict[str, dict[str, str]] = {}
        if not args.skip_product_catalog:
            try:
                product_catalog = fetch_product_catalog(
                    vendor_id=vendor_id,
                    access_key=access_key,
                    secret_key=secret_key,
                )
            except RuntimeError as exc:
                print(
                    "Warning: product catalog lookup failed, continuing with IDs only: "
                    f"{exc}",
                    file=sys.stderr,
                )

        monitored_items = build_monitored_items(
            inventory_rows=inventory_rows,
            product_catalog=product_catalog,
            quantity_threshold=args.quantity_threshold,
            days_threshold=args.days_threshold,
            specs_by_sku=specs_by_sku,
        )
        alert_items = build_alert_items(monitored_items)

        order_rows = fetch_rg_orders(
            vendor_id=vendor_id,
            access_key=access_key,
            secret_key=secret_key,
            paid_date_from=report_context.yesterday_date,
            paid_date_to=report_context.today_date,
        )
        sales_summaries = build_sales_summaries(
            monitored_items=monitored_items,
            order_rows=order_rows,
            specs_by_sku=specs_by_sku,
            report_context=report_context,
        )

        payload = build_payload(
            alert_items=alert_items,
            monitored_items=monitored_items,
            sales_summaries=sales_summaries,
            quantity_threshold=args.quantity_threshold,
            days_threshold=args.days_threshold,
            send_healthy_summary=args.send_healthy_summary,
            report_context=report_context,
        )

        if payload is None:
            print("No monitored items found. Slack message skipped.")
            return 0

        if args.dry_run:
            print(json.dumps(payload, ensure_ascii=False, indent=2))
            return 0

        send_slack_payload(webhook_url, payload)
        print(
            f"Sent Slack daily report with {len(alert_items)} low-stock item(s) and {len(sales_summaries)} monitored SKU summaries."
        )
        return 0
    except (RuntimeError, ValueError, json.JSONDecodeError) as exc:
        print(f"Failed to build Coupang daily report: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
