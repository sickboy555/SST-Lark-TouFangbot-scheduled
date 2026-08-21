from __future__ import annotations

import argparse
import base64
import datetime as dt
import hashlib
import hmac
import json
import mimetypes
import os
import re
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib.parse import quote, urlparse, parse_qs

from zoneinfo import ZoneInfo

try:
    from dotenv import load_dotenv
except ImportError:
    def load_dotenv() -> bool:
        return False


COUNTRY_COL_NAMES = ("国家", "country", "国家/地区", "市场")
PACKAGE_COL_NAMES = ("包名", "包", "package", "bundle", "产品包", "包体")
DATE_COL_NAMES = ("日期", "date", "day", "时间", "统计日期", "消耗日期")
VALUE_COL_NAMES = (
    "消耗",
    "消耗预算",
    "预算消耗",
    "预算",
    "总预算",
    "日消耗",
    "当日消耗",
    "cost",
    "spend",
    "amount",
    "value",
)
SUMMARY_KEYWORDS = ("汇总", "总计", "合计", "subtotal", "total")


@dataclass(frozen=True)
class Config:
    app_id: str
    app_secret: str
    chat_id: str
    sheet_url: str
    open_base_url: str
    sheet_range: str
    schedule_time: str
    timezone: str
    output_dir: Path
    send_mode: str
    webhook_url: str
    webhook_secret: str
    stability_retries: int
    stability_delay_seconds: float
    filter_zero_rows: bool

    @staticmethod
    def from_env() -> "Config":
        load_dotenv()
        send_mode = os.getenv("SEND_MODE", "app").strip().lower()
        if send_mode not in {"app", "webhook"}:
            raise RuntimeError("SEND_MODE 只能是 app 或 webhook。")

        required = {
            "LARK_APP_ID": os.getenv("LARK_APP_ID"),
            "LARK_APP_SECRET": os.getenv("LARK_APP_SECRET"),
            "LARK_SHEET_URL": os.getenv("LARK_SHEET_URL"),
        }
        if send_mode == "app":
            required["LARK_CHAT_ID"] = os.getenv("LARK_CHAT_ID")
        missing = [k for k, v in required.items() if not v or v.startswith("填入")]
        if missing:
            raise RuntimeError(f"缺少环境变量: {', '.join(missing)}。请先复制 .env.example 为 .env 并填写。")

        return Config(
            app_id=required["LARK_APP_ID"] or "",
            app_secret=required["LARK_APP_SECRET"] or "",
            chat_id=os.getenv("LARK_CHAT_ID", ""),
            sheet_url=required["LARK_SHEET_URL"] or "",
            open_base_url=os.getenv("LARK_OPEN_BASE_URL", "https://open.larksuite.com").rstrip("/"),
            sheet_range=os.getenv("SHEET_RANGE", "A1:AZ1000"),
            schedule_time=os.getenv("SCHEDULE_TIME", "10:00"),
            timezone=os.getenv("TIMEZONE", "Asia/Shanghai"),
            output_dir=Path(os.getenv("OUTPUT_DIR", "output")),
            send_mode=send_mode,
            webhook_url=os.getenv("LARK_WEBHOOK_URL", ""),
            webhook_secret=os.getenv("LARK_WEBHOOK_SECRET", ""),
            stability_retries=max(2, int(os.getenv("LARK_STABILITY_RETRIES", "4"))),
            stability_delay_seconds=max(0.0, float(os.getenv("LARK_STABILITY_DELAY_SECONDS", "2"))),
            filter_zero_rows=os.getenv("FILTER_ZERO_ROWS", "true").strip().lower() in {"1", "true", "yes", "on"},
        )


@dataclass(frozen=True)
class SheetRef:
    spreadsheet_token: str
    sheet_id: str


@dataclass
class BudgetRow:
    country: str
    package: str
    today: float
    yesterday: float

    @property
    def diff(self) -> float:
        return self.today - self.yesterday

    @property
    def is_summary(self) -> bool:
        return any(keyword.lower() in self.package.lower() for keyword in SUMMARY_KEYWORDS)


class LarkClient:
    def __init__(self, config: Config):
        import requests

        self.config = config
        self.session = requests.Session()
        self._tenant_access_token: str | None = None

    def tenant_access_token(self) -> str:
        if self._tenant_access_token:
            return self._tenant_access_token

        url = f"{self.config.open_base_url}/open-apis/auth/v3/tenant_access_token/internal"
        resp = self.session.post(
            url,
            json={"app_id": self.config.app_id, "app_secret": self.config.app_secret},
            timeout=30,
        )
        data = self._checked_json(resp, "获取 tenant_access_token 失败")
        token = data.get("tenant_access_token")
        if not token:
            raise RuntimeError(f"获取 tenant_access_token 失败: {data}")
        self._tenant_access_token = token
        return token

    def _headers(self) -> dict[str, str]:
        return {"Authorization": f"Bearer {self.tenant_access_token()}"}

    def get_sheet_values(self, sheet_ref: SheetRef, cell_range: str) -> list[list[Any]]:
        range_expr = f"{sheet_ref.sheet_id}!{cell_range}"
        encoded_range = quote(range_expr, safe="")
        url = (
            f"{self.config.open_base_url}/open-apis/sheets/v2/spreadsheets/"
            f"{sheet_ref.spreadsheet_token}/values/{encoded_range}"
        )
        # UnformattedValue 让公式单元格返回已计算数值；日期序列号由 parse_date_cell 转换。
        resp = self.session.get(
            url,
            headers=self._headers(),
            params={"valueRenderOption": "UnformattedValue"},
            timeout=30,
        )
        data = self._checked_json(resp, "读取表格失败")
        return data.get("data", {}).get("valueRange", {}).get("values", []) or []

    def upload_image(self, image_path: Path) -> str:
        url = f"{self.config.open_base_url}/open-apis/im/v1/images"
        mime = mimetypes.guess_type(str(image_path))[0] or "image/png"
        with image_path.open("rb") as fp:
            files = {"image": (image_path.name, fp, mime)}
            data = {"image_type": "message"}
            resp = self.session.post(url, headers=self._headers(), data=data, files=files, timeout=60)
        payload = self._checked_json(resp, "上传图片失败")
        image_key = payload.get("data", {}).get("image_key")
        if not image_key:
            raise RuntimeError(f"上传图片失败: {payload}")
        return image_key

    def send_text(self, text: str) -> None:
        self._send_message("text", {"text": text})

    def send_image(self, image_key: str) -> None:
        self._send_message("image", {"image_key": image_key})

    def _send_message(self, msg_type: str, content: dict[str, Any]) -> None:
        url = f"{self.config.open_base_url}/open-apis/im/v1/messages?receive_id_type=chat_id"
        resp = self.session.post(
            url,
            headers={**self._headers(), "Content-Type": "application/json"},
            json={
                "receive_id": self.config.chat_id,
                "msg_type": msg_type,
                "content": json.dumps(content, ensure_ascii=False),
            },
            timeout=30,
        )
        self._checked_json(resp, "发送群消息失败")

    @staticmethod
    def _checked_json(resp: Any, action: str) -> dict[str, Any]:
        try:
            data = resp.json()
        except ValueError as exc:
            raise RuntimeError(f"{action}: HTTP {resp.status_code}, {resp.text[:500]}") from exc
        if resp.status_code >= 400 or data.get("code", 0) != 0:
            raise RuntimeError(f"{action}: HTTP {resp.status_code}, {data}")
        return data


class LarkWebhookClient:
    def __init__(self, config: Config):
        import requests

        if not config.webhook_url or config.webhook_url.startswith("填入"):
            raise RuntimeError("SEND_MODE=webhook 时必须填写 LARK_WEBHOOK_URL。")
        self.config = config
        self.session = requests.Session()

    def send_text(self, text: str) -> None:
        self._post({"msg_type": "text", "content": {"text": text}})

    def send_image(self, image_key: str) -> None:
        self._post({"msg_type": "image", "content": {"image_key": image_key}})

    def _post(self, payload: dict[str, Any]) -> None:
        if self.config.webhook_secret:
            timestamp = str(int(dt.datetime.now().timestamp()))
            payload = {
                **payload,
                "timestamp": timestamp,
                "sign": self._sign(timestamp, self.config.webhook_secret),
            }

        resp = self.session.post(self.config.webhook_url, json=payload, timeout=30)
        try:
            data = resp.json()
        except ValueError as exc:
            raise RuntimeError(f"Webhook 发送失败: HTTP {resp.status_code}, {resp.text[:500]}") from exc

        error_code = data.get("code", data.get("StatusCode", 0))
        if resp.status_code >= 400 or error_code != 0:
            raise RuntimeError(f"Webhook 发送失败: HTTP {resp.status_code}, {data}")

    @staticmethod
    def _sign(timestamp: str, secret: str) -> str:
        string_to_sign = f"{timestamp}\n{secret}"
        digest = hmac.new(string_to_sign.encode("utf-8"), b"", hashlib.sha256).digest()
        return base64.b64encode(digest).decode("utf-8")


def parse_sheet_url(url: str) -> SheetRef:
    parsed = urlparse(url)
    token_match = re.search(r"/sheets/([^/?#]+)", parsed.path)
    if not token_match:
        raise ValueError(f"无法从表格 URL 解析 spreadsheet token: {url}")
    sheet_id = parse_qs(parsed.query).get("sheet", [""])[0]
    if not sheet_id:
        raise ValueError(f"无法从表格 URL 解析 sheet id: {url}")
    return SheetRef(spreadsheet_token=token_match.group(1), sheet_id=sheet_id)


def date_labels(now: dt.date) -> tuple[str, str]:
    yesterday = now - dt.timedelta(days=1)
    return format_cn_date(now), format_cn_date(yesterday)


def format_cn_date(value: dt.date) -> str:
    return f"{value.month}月{value.day}日"


def normalize_text(value: Any) -> str:
    return str(value).strip() if value is not None else ""


def parse_number(value: Any) -> float:
    if value is None or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace(",", "").strip()
    if not text:
        return 0.0
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    return float(match.group(0)) if match else 0.0


def same_date_header(cell: Any, target: dt.date) -> bool:
    return parse_date_cell(cell) == target


def parse_date_cell(value: Any) -> dt.date | None:
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value

    text = normalize_text(value)
    if not text:
        return None

    # UnformattedValue returns spreadsheet dates as Excel serial numbers.
    if isinstance(value, (int, float)) and 1 <= value <= 100000:
        return dt.date(1899, 12, 30) + dt.timedelta(days=int(value))

    compact = text.replace(" ", "")
    today_year = dt.date.today().year
    formats = (
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%Y.%m.%d",
        "%m-%d",
        "%m/%d",
        "%m.%d",
    )
    for fmt in formats:
        try:
            parsed = dt.datetime.strptime(compact, fmt)
        except ValueError:
            continue
        year = parsed.year if "%Y" in fmt else today_year
        return dt.date(year, parsed.month, parsed.day)

    match = re.search(r"(?:(\d{4})年)?(\d{1,2})月(\d{1,2})日", compact)
    if match:
        year = int(match.group(1) or today_year)
        return dt.date(year, int(match.group(2)), int(match.group(3)))

    return None


def find_col(normalized: list[str], candidates: tuple[str, ...]) -> int:
    candidate_set = {c.lower() for c in candidates}
    return next((i for i, cell in enumerate(normalized) if cell in candidate_set), -1)


def find_column_header(values: list[list[Any]]) -> tuple[int, int, int, list[tuple[dt.date, int]]] | None:
    """找到“日期在列”表头：包含国家、包名，且有至少两个可解析为日期的列。

    返回 (表头行号, 国家列, 包名列, [(日期, 列号), ...] 按日期从新到旧排序)。
    选取最近的两个日期列作为“今天 vs 昨天”，这样即使表格里最新一天还没到当天日历日，
    也能正确对比最近两天的数据。
    """
    for row_idx, row in enumerate(values[:20]):
        normalized = [normalize_text(cell).lower() for cell in row]
        country_col = find_col(normalized, COUNTRY_COL_NAMES)
        package_col = find_col(normalized, PACKAGE_COL_NAMES)
        if country_col < 0 or package_col < 0:
            continue
        date_cols = [(d, i) for i, cell in enumerate(row) if (d := parse_date_cell(cell)) is not None]
        if len(date_cols) >= 2:
            date_cols.sort(key=lambda pair: pair[0], reverse=True)
            return row_idx, country_col, package_col, date_cols
    return None


def find_row_header(values: list[list[Any]]) -> tuple[int, int, int, int, int] | None:
    for row_idx, row in enumerate(values[:20]):
        normalized = [normalize_text(cell).lower() for cell in row]
        country_col = find_col(normalized, COUNTRY_COL_NAMES)
        package_col = find_col(normalized, PACKAGE_COL_NAMES)
        date_col = find_col(normalized, DATE_COL_NAMES)
        value_col = find_col(normalized, VALUE_COL_NAMES)
        if country_col >= 0 and package_col >= 0 and date_col >= 0 and value_col >= 0:
            return row_idx, country_col, package_col, date_col, value_col
    return None


def parse_budget_rows(values: list[list[Any]], today: dt.date) -> tuple[list[BudgetRow], str, str]:
    """返回 (数据行, 今天列标签, 昨天列标签)。

    日期在列时使用表格里最近的两个日期列；日期在行时使用日历今天/昨天。
    """
    column_header = find_column_header(values)
    if column_header:
        return parse_budget_rows_by_date_columns(values, column_header)

    row_header = find_row_header(values)
    if row_header:
        return parse_budget_rows_by_date_rows(values, today, row_header)

    raise RuntimeError(
        "没有找到可用表头。日期在列时需要包含 国家、包名 以及至少两个日期列；"
        "日期在行时需要包含 国家、包名、日期、消耗。"
    )


def parse_budget_rows_by_date_columns(
    values: list[list[Any]],
    header: tuple[int, int, int, list[tuple[dt.date, int]]],
) -> tuple[list[BudgetRow], str, str]:
    header_idx, country_col, package_col, date_cols = header
    (today_date, today_col), (yesterday_date, yesterday_col) = date_cols[0], date_cols[1]
    rows: list[BudgetRow] = []
    current_country = ""

    for row in values[header_idx + 1 :]:
        max_idx = max(country_col, package_col, today_col, yesterday_col)
        padded = list(row) + [""] * (max_idx + 1 - len(row))
        country = normalize_text(padded[country_col])
        package = normalize_text(padded[package_col])

        if country:
            current_country = country
        if not package or not current_country:
            continue

        rows.append(
            BudgetRow(
                country=current_country,
                package=package,
                today=parse_number(padded[today_col]),
                yesterday=parse_number(padded[yesterday_col]),
            )
        )

    return require_rows(rows), format_cn_date(today_date), format_cn_date(yesterday_date)


def dedupe_duplicate_country_blocks(rows: list[BudgetRow]) -> list[BudgetRow]:
    """删除内容完全重复的国家区块，避免表尾残留复制区块重复计入。"""
    blocks: list[tuple[str, list[BudgetRow]]] = []
    for row in rows:
        if blocks and blocks[-1][0] == row.country:
            blocks[-1][1].append(row)
        else:
            blocks.append((row.country, [row]))

    seen: dict[tuple[str, frozenset[tuple[str, float, float]]], int] = {}
    result: list[tuple[str, list[BudgetRow]]] = []
    for country, block in blocks:
        signature = frozenset((normalize_text(row.package).casefold(), row.today, row.yesterday) for row in block)
        key = (country, signature)
        previous_index = seen.get(key)
        if previous_index is not None:
            result.pop(previous_index)
            seen = {k: (i if i < previous_index else i - 1) for k, i in seen.items() if i != previous_index}
        seen[key] = len(result)
        result.append((country, block))
    return [row for _, block in result for row in block]


def _declared_totals(values: list[list[Any]], header: tuple[int, int, int, list[tuple[dt.date, int]]]) -> tuple[float, float] | None:
    """读取表格顶部的公式总计；没有可解析总计时返回 None。"""
    if not values:
        return None
    _, _, _, date_cols = header
    if len(date_cols) < 2:
        return None
    today_col = date_cols[0][1]
    yesterday_col = date_cols[1][1]
    first = list(values[0])
    max_idx = max(today_col, yesterday_col)
    if len(first) <= max_idx:
        return None
    first_today = first[today_col]
    first_yesterday = first[yesterday_col]
    if not isinstance(first_today, (int, float)) or not isinstance(first_yesterday, (int, float)):
        return None
    return float(first_today), float(first_yesterday)


def _validated_values(values: list[list[Any]], today: dt.date) -> tuple[bool, tuple[tuple[str, str, float, float], ...]]:
    header = find_column_header(values)
    if header is None:
        return True, tuple()
    rows, _, _ = parse_budget_rows_by_date_columns(values, header)
    rows = dedupe_duplicate_country_blocks(rows)
    summary_rows = [row for row in rows if row.is_summary]
    if not summary_rows:
        return True, tuple((r.country, r.package, r.today, r.yesterday) for r in rows)
    expected = (sum(row.today for row in summary_rows), sum(row.yesterday for row in summary_rows))
    declared = _declared_totals(values, header)
    if declared is not None and declared != expected:
        return False, tuple((r.country, r.package, r.today, r.yesterday) for r in rows)
    signature = tuple((r.country, r.package, r.today, r.yesterday) for r in rows)
    return True, signature


def load_stable_budget_data(
    fetch_values: Any,
    today: dt.date,
    retries: int = 4,
    delay_seconds: float = 2.0,
    sleep_fn: Any = time.sleep,
) -> list[list[Any]]:
    """只在公式总计有效且连续两次读取完全一致时返回数据。"""
    previous_signature: tuple[tuple[str, str, float, float], ...] | None = None
    last_reason = ""
    for attempt in range(max(2, retries)):
        values = fetch_values()
        valid, signature = _validated_values(values, today)
        if not valid:
            last_reason = "总计与明细汇总不一致，等待 Lark 公式刷新"
        elif previous_signature is not None and signature == previous_signature:
            return values
        else:
            last_reason = "连续两次读取内容不一致，等待 Lark 公式刷新"
        if valid:
            previous_signature = signature
        if attempt + 1 < max(2, retries):
            sleep_fn(delay_seconds)
    raise RuntimeError(f"Lark 表格数据未稳定，已重试 {max(2, retries)} 次：{last_reason}。")


def parse_budget_rows_by_date_rows(
    values: list[list[Any]],
    today: dt.date,
    header: tuple[int, int, int, int, int],
) -> tuple[list[BudgetRow], str, str]:
    header_idx, country_col, package_col, date_col, value_col = header
    yesterday = today - dt.timedelta(days=1)
    by_key: dict[tuple[str, str], dict[str, float]] = {}
    order: list[tuple[str, str]] = []
    current_country = ""

    for row in values[header_idx + 1 :]:
        max_idx = max(country_col, package_col, date_col, value_col)
        padded = list(row) + [""] * (max_idx + 1 - len(row))
        country = normalize_text(padded[country_col])
        package = normalize_text(padded[package_col])
        row_date = parse_date_cell(padded[date_col])

        if country:
            current_country = country
        if not package or not current_country or row_date not in {today, yesterday}:
            continue

        key = (current_country, package)
        if key not in by_key:
            by_key[key] = {"today": 0.0, "yesterday": 0.0}
            order.append(key)

        bucket = "today" if row_date == today else "yesterday"
        by_key[key][bucket] += parse_number(padded[value_col])

    raw_rows = [BudgetRow(country, package, by_key[(country, package)]["today"], by_key[(country, package)]["yesterday"]) for country, package in order]
    rows = add_missing_summary_rows(raw_rows)
    today_label, yesterday_label = date_labels(today)
    return require_rows(rows), today_label, yesterday_label


def add_missing_summary_rows(rows: list[BudgetRow]) -> list[BudgetRow]:
    result: list[BudgetRow] = []
    idx = 0
    while idx < len(rows):
        country = rows[idx].country
        group: list[BudgetRow] = []
        while idx < len(rows) and rows[idx].country == country:
            group.append(rows[idx])
            idx += 1

        result.extend(group)
        if not any(row.is_summary for row in group):
            prefix = country_prefix(country)
            result.append(
                BudgetRow(
                    country=country,
                    package=f"{prefix}_汇总" if prefix else "汇总",
                    today=sum(row.today for row in group),
                    yesterday=sum(row.yesterday for row in group),
                )
            )
    return result


def country_prefix(country: str) -> str:
    mapping = {
        "墨西哥": "MX",
        "秘鲁": "PE",
        "哥伦比亚": "CO",
        "巴西": "BR",
        "智利": "CL",
        "尼日利亚": "NG",
        "菲律宾": "PH",
        "越南": "VN",
        "印尼": "ID",
        "印度尼西亚": "ID",
        "泰国": "TH",
    }
    return mapping.get(country, country[:2].upper() if country else "")


def require_rows(rows: list[BudgetRow]) -> list[BudgetRow]:
    if not rows:
        raise RuntimeError("表格里没有解析到今天/昨天的有效国家/包名数据。")
    return rows


def filter_zero_rows(rows: list[BudgetRow]) -> list[BudgetRow]:
    """过滤掉今天和昨天预算都为 0 的行；若某国汇总也为 0/0，整个国家区块删除。"""
    seen: set[str] = set()
    order: list[str] = []
    groups: dict[str, list[BudgetRow]] = {}
    for row in rows:
        if row.country not in seen:
            seen.add(row.country)
            order.append(row.country)
            groups[row.country] = []
        groups[row.country].append(row)

    result: list[BudgetRow] = []
    for country in order:
        country_rows = groups[country]
        summary = next((r for r in country_rows if r.is_summary), None)
        if summary is not None and summary.today == 0 and summary.yesterday == 0:
            continue
        for row in country_rows:
            if not row.is_summary and row.today == 0 and row.yesterday == 0:
                continue
            result.append(row)
    return result


def load_font(size: int, bold: bool = False) -> Any:
    from PIL import ImageFont

    candidates = [
        "C:/Windows/Fonts/msyhbd.ttc" if bold else "C:/Windows/Fonts/msyh.ttc",
        "C:/Windows/Fonts/simhei.ttf",
        "C:/Windows/Fonts/arial.ttf",
    ]
    for path in candidates:
        if path and Path(path).exists():
            return ImageFont.truetype(path, size=size)
    return ImageFont.load_default()


def fmt_amount(value: float) -> str:
    if abs(value - round(value)) < 0.0001:
        return f"{int(round(value)):,}"
    return f"{value:,.2f}".rstrip("0").rstrip(".")


def fmt_diff(value: float) -> str:
    if value > 0:
        return f"+{fmt_amount(value)}"
    return fmt_amount(value)


def fmt_diff_indicator(diff: float) -> str:
    if diff > 0.005:
        return f"↑ +{fmt_amount(diff)}"
    if diff < -0.005:
        return f"↓ {fmt_amount(diff)}"
    return "持平"


def make_report_image(
    rows: list[BudgetRow],
    today_label: str,
    yesterday_label: str,
    output_dir: Path,
    file_date: dt.date,
) -> Path:
    from PIL import Image, ImageDraw, ImageFont

    output_dir.mkdir(parents=True, exist_ok=True)

    row_h = 42
    title_h = 52
    header_h = 42
    footer_h = 18
    widths = [118, 220, 130, 130, 120]
    width = sum(widths) + 2
    height = title_h + header_h + row_h * len(rows) + footer_h

    img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(img)

    font_title = load_font(24, bold=True)
    font_header = load_font(18, bold=True)
    font_body = load_font(18)
    font_bold = load_font(18, bold=True)

    country_colors = [
        "#d8eef7",
        "#ddf3d8",
        "#f9dced",
        "#d8eef7",
        "#e8ddf6",
        "#f8f0cf",
        "#f7c3be",
        "#dedfe3",
        "#dbe5fa",
        "#f6dfd6",
        "#eef6c8",
        "#c9f1df",
    ]
    country_color_map: dict[str, str] = {}

    summary_rows = [r for r in rows if r.is_summary]
    total_source = summary_rows or [r for r in rows if not r.is_summary]
    total_today = sum(r.today for r in total_source)
    total_yesterday = sum(r.yesterday for r in total_source)

    def text_center(box: tuple[int, int, int, int], text: str, font: ImageFont.ImageFont, fill: str = "#222") -> None:
        bbox = draw.textbbox((0, 0), text, font=font)
        x = box[0] + (box[2] - box[0] - (bbox[2] - bbox[0])) / 2
        y = box[1] + (box[3] - box[1] - (bbox[3] - bbox[1])) / 2 - 1
        draw.text((x, y), text, font=font, fill=fill)

    def text_left(box: tuple[int, int, int, int], text: str, font: ImageFont.ImageFont, fill: str = "#222") -> None:
        bbox = draw.textbbox((0, 0), text, font=font)
        y = box[1] + (box[3] - box[1] - (bbox[3] - bbox[1])) / 2 - 1
        draw.text((box[0] + 12, y), text, font=font, fill=fill)

    x_positions = [0]
    for w in widths:
        x_positions.append(x_positions[-1] + w)

    y = 0
    draw.rectangle((0, y, width - 1, title_h), fill="#ffffff", outline="#2f2f2f")
    text_center((0, y, widths[0] + widths[1], title_h), "全市场总预算", font_title, "#e65764")
    text_center((x_positions[2], y, x_positions[3], title_h), fmt_amount(total_today), font_title, "#e65764")
    text_center((x_positions[3], y, x_positions[4], title_h), fmt_amount(total_yesterday), font_title, "#e65764")
    text_center((x_positions[4], y, x_positions[5], title_h), fmt_diff(total_today - total_yesterday), font_title, diff_color(total_today - total_yesterday))

    y += title_h
    draw.rectangle((0, y, width - 1, y + header_h), fill="#d7f0f8", outline="#2f2f2f")
    headers = ["国家", "包名", today_label, yesterday_label, "较昨日"]
    for i, title in enumerate(headers):
        text_center((x_positions[i], y, x_positions[i + 1], y + header_h), title, font_header)

    y += header_h
    for row in rows:
        if row.country not in country_color_map:
            country_color_map[row.country] = country_colors[len(country_color_map) % len(country_colors)]
        bg = country_color_map[row.country] if row.is_summary else "#ffffff"
        draw.rectangle((0, y, width - 1, y + row_h), fill=bg, outline="#2f2f2f")

        font = font_bold if row.is_summary else font_body
        text_left((x_positions[1], y, x_positions[2], y + row_h), row.package, font)
        text_center((x_positions[2], y, x_positions[3], y + row_h), fmt_amount(row.today), font, amount_color(row.today, row.yesterday))
        text_center((x_positions[3], y, x_positions[4], y + row_h), fmt_amount(row.yesterday), font)
        text_center((x_positions[4], y, x_positions[5], y + row_h), fmt_diff(row.diff), font, diff_color(row.diff))
        y += row_h

    # Draw country merged cells after package rows, so borders align like the source screenshot.
    start = title_h + header_h
    idx = 0
    while idx < len(rows):
        country = rows[idx].country
        end = idx
        while end + 1 < len(rows) and rows[end + 1].country == country:
            end += 1
        top = start + idx * row_h
        bottom = start + (end + 1) * row_h
        draw.rectangle((x_positions[0], top, x_positions[1], bottom), fill=country_color_map[country], outline="#2f2f2f")
        text_center((x_positions[0], top, x_positions[1], bottom), country, font_header)
        idx = end + 1

    for i, x in enumerate(x_positions):
        # 标题行里国家列与包名列合并，跳过两者之间的竖线
        y_start = title_h if i == 1 else 0
        draw.line((x, y_start, x, height - footer_h), fill="#2f2f2f", width=1)
    draw.line((width - 1, 0, width - 1, height - footer_h), fill="#2f2f2f", width=1)

    path = output_dir / f"budget_report_{file_date.isoformat()}.png"
    img.save(path)
    return path


def diff_color(value: float) -> str:
    if value > 0:
        return "#e65764"
    if value < 0:
        return "#38b653"
    return "#222222"


def amount_color(today: float, yesterday: float) -> str:
    return diff_color(today - yesterday)


def build_summary_text(rows: list[BudgetRow], today_label: str, yesterday_label: str) -> str:
    summary_rows = [r for r in rows if r.is_summary]
    total_source = summary_rows or rows
    total_today = sum(r.today for r in total_source)
    total_yesterday = sum(r.yesterday for r in total_source)
    total_diff = total_today - total_yesterday

    lines: list[str] = []
    lines.append(f"📊 预算日报 | {today_label}")
    lines.append(
        f"全市场总预算 {fmt_amount(total_today)}"
        f"（{yesterday_label} {fmt_amount(total_yesterday)}，{fmt_diff_indicator(total_diff)}）"
    )

    seen: set[str] = set()
    order: list[str] = []
    groups: dict[str, list[BudgetRow]] = {}
    for row in rows:
        if row.country not in seen:
            seen.add(row.country)
            order.append(row.country)
            groups[row.country] = []
        groups[row.country].append(row)

    for country in order:
        country_rows = groups[country]
        lines.append("")
        lines.append(f"【{country}】")

        summary_row: BudgetRow | None = None
        for row in country_rows:
            if row.is_summary:
                summary_row = row
            else:
                lines.append(
                    f"    • {row.package}：{fmt_amount(row.today)}（{fmt_diff_indicator(row.diff)}）"
                )

        if summary_row is not None:
            lines.append(
                f"  ▸ {summary_row.package}：{fmt_amount(summary_row.today)}"
                f"（{fmt_diff_indicator(summary_row.diff)}）"
            )

    return "\n".join(lines)


def run_once(config: Config, send: bool = True) -> Path:
    tz = ZoneInfo(config.timezone)
    today = dt.datetime.now(tz).date()
    sheet_ref = parse_sheet_url(config.sheet_url)
    client = LarkClient(config)

    values = load_stable_budget_data(
        lambda: client.get_sheet_values(sheet_ref, config.sheet_range),
        today,
        retries=config.stability_retries,
        delay_seconds=config.stability_delay_seconds,
    )
    rows, today_label, yesterday_label = parse_budget_rows(values, today)
    rows = dedupe_duplicate_country_blocks(rows)
    if config.filter_zero_rows:
        rows = filter_zero_rows(rows)
    image_path = make_report_image(rows, today_label, yesterday_label, config.output_dir, today)
    text = build_summary_text(rows, today_label, yesterday_label)
    print(f"对比日期列: {today_label} vs {yesterday_label}")

    if send:
        image_key = client.upload_image(image_path)
        sender = client if config.send_mode == "app" else LarkWebhookClient(config)
        sender.send_text(text)
        sender.send_image(image_key)

    print(f"生成图片: {image_path}")
    if send:
        print(f"已通过 {config.send_mode} 模式发送到 Lark 群。")
    else:
        print("已跳过发送，仅本地生成。")
    return image_path


def run_daemon(config: Config) -> None:
    from apscheduler.schedulers.blocking import BlockingScheduler
    from apscheduler.triggers.cron import CronTrigger

    hour, minute = parse_schedule_time(config.schedule_time)
    scheduler = BlockingScheduler(timezone=config.timezone)
    scheduler.add_job(
        lambda: run_with_log(config),
        CronTrigger(hour=hour, minute=minute, timezone=config.timezone),
        id="daily_budget_report",
        replace_existing=True,
    )
    print(f"机器人已启动，每天 {config.schedule_time} ({config.timezone}) 发送。按 Ctrl+C 退出。")
    scheduler.start()


def run_with_log(config: Config) -> None:
    try:
        run_once(config, send=True)
    except Exception as exc:
        print(f"[{dt.datetime.now().isoformat(timespec='seconds')}] 运行失败: {exc}", file=sys.stderr)


def parse_schedule_time(value: str) -> tuple[int, int]:
    match = re.fullmatch(r"([01]?\d|2[0-3]):([0-5]\d)", value.strip())
    if not match:
        raise ValueError("SCHEDULE_TIME 格式应为 HH:MM，例如 10:00")
    return int(match.group(1)), int(match.group(2))


def main() -> None:
    parser = argparse.ArgumentParser(description="Lark 表格预算消耗对比机器人")
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--once", action="store_true", help="立即读取表格并发送一次")
    mode.add_argument("--daemon", action="store_true", help="常驻进程，每天定时发送")
    parser.add_argument("--dry-run", action="store_true", help="只生成图片和摘要，不发送群消息")
    args = parser.parse_args()

    config = Config.from_env()
    if args.once:
        run_once(config, send=not args.dry_run)
    elif args.daemon:
        if args.dry_run:
            print("--daemon 模式不支持 --dry-run。")
            sys.exit(2)
        run_daemon(config)


if __name__ == "__main__":
    main()
