from __future__ import annotations

import argparse
import json
import os
import re
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Any
from urllib.parse import urlparse, urlunparse
from zoneinfo import ZoneInfo

import requests
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

GOOGLE_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
ALLOWED_UTM_RULES = {
    ("leadit", "cpa", "frml"): "Лидген КЦ",
    ("selfwalk", "organic", "frml"): "Самоход",
    ("avito", "cpc", "frml"): "Авито",
    ("recommendation", "call", "frml"): "Рекомендация",
}
LEAD_FALLBACK_UTM_FIELDS = {
    "source": ["UF_LEAD_FIRST_UTM_SOURCE"],
    "medium": ["UF_LEAD_FIRST_UTM_MEDIUM"],
    "campaign": ["UF_LEAD_FIRST_UTM_CAMPAIGN"],
}
MONTH_LABELS_RU = {
    1: "Январь",
    2: "Февраль",
    3: "Март",
    4: "Апрель",
    5: "Май",
    6: "Июнь",
    7: "Июль",
    8: "Август",
    9: "Сентябрь",
    10: "Октябрь",
    11: "Ноябрь",
    12: "Декабрь",
}
HEADER_ALIASES = {
    "number": {"№", "no", "number", "номер"},
    "utm_source": {"utm source", "utm_source"},
    "utm_medium": {"utm medium", "utm_medium"},
    "utm_campaign": {"utm campaign", "utm_campaign"},
    "date_created": {"дата создания", "date create", "date created", "дата"},
    "period": {"период", "period"},
    "source": {"источник", "source"},
    "total": {
        "суммарный объем",
        "итог по источникам",
        "объем",
        "объем лидов",
        "количество",
        "sum",
        "total",
    },
}
LEGACY_SUMMARY_BLOCK_LABELS = {
    "итог по источникам",
    "период",
    "источник",
    "суммарный объем",
}


class ConfigError(RuntimeError):
    """Raised when configuration is missing or invalid."""


@dataclass(frozen=True)
class ReportWindow:
    start: datetime
    end: datetime


@dataclass(frozen=True)
class UtmKey:
    utm_source: str
    utm_medium: str
    utm_campaign: str


@dataclass(frozen=True)
class SheetRange:
    start_col: str
    start_row: int
    end_col: str
    end_row: int

    @property
    def width(self) -> int:
        return a1_to_col_index(self.end_col) - a1_to_col_index(self.start_col) + 1

    def build_a1(self, sheet_title: str, row_end: int | None = None, row_start: int | None = None) -> str:
        escaped_title = sheet_title.replace("'", "''")
        quoted_title = sheet_title if re.fullmatch(r"[A-Za-z0-9_]+", sheet_title) else f"'{escaped_title}'"
        actual_row_start = row_start or self.start_row
        actual_row_end = row_end or self.end_row
        return f"{quoted_title}!{self.start_col}{actual_row_start}:{self.end_col}{actual_row_end}"


@dataclass
class SheetContext:
    spreadsheet_id: str
    spreadsheet_title: str
    sheet_id: int
    sheet_title: str
    allowed_range: SheetRange
    row_count: int
    row_groups: list[dict[str, Any]]


@dataclass
class Settings:
    bitrix_webhook_url: str
    bitrix_entity_type: str
    bitrix_date_field: str
    bitrix_utm_source_field: str
    bitrix_utm_medium_field: str
    bitrix_utm_campaign_field: str
    bitrix_request_timeout: int
    google_sheet_id: str
    google_sheet_name: str
    google_allowed_range: str
    google_source_summary_sheet_name: str
    google_service_account_file: str | None
    google_service_account_json: str | None
    report_timezone: str
    report_period_mode: str
    report_start_date: str
    report_require_all_utm: bool
    report_unknown_source: str


@dataclass
class ReportBuildResult:
    rows: list[list[Any]]
    month_groups: list[tuple[int, int]]
    day_groups: list[tuple[int, int]]
    summary_rows: list[int]
    month_count: int
    day_count: int
    detail_count: int
    record_count: int


@dataclass
class SheetWriteTarget:
    sheet_id: int
    sheet_title: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build and sync a Bitrix24 UTM report into Google Sheets with month/day grouping."
    )
    parser.add_argument(
        "--env-file",
        help="Path to the env file. If omitted, the script will try bitrix.env and .env.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Compute the report without writing to Google Sheets.",
    )
    return parser.parse_args()


def load_environment(env_file: str | None) -> None:
    if env_file:
        env_path = Path(env_file)
        if not env_path.exists():
            raise ConfigError(f"Env file not found: {env_path}")
        load_dotenv(env_path, override=True)
        return

    for candidate in ("bitrix.env", ".env"):
        candidate_path = Path(candidate)
        if candidate_path.exists():
            load_dotenv(candidate_path, override=False)


def require_env(name: str) -> str:
    value = os.getenv(name)
    if value is None or not value.strip():
        raise ConfigError(f"Missing required env var: {name}")
    return value.strip()


def read_bool(name: str, default: bool) -> bool:
    raw = os.getenv(name)
    if raw is None or not raw.strip():
        return default
    normalized = raw.strip().lower()
    if normalized in {"1", "true", "yes", "y"}:
        return True
    if normalized in {"0", "false", "no", "n"}:
        return False
    raise ConfigError(f"Invalid boolean value for {name}: {raw}")


def read_int(name: str, default: int) -> int:
    raw = os.getenv(name)
    if raw is None or not raw.strip():
        return default
    try:
        value = int(raw)
    except ValueError as exc:
        raise ConfigError(f"Invalid integer value for {name}: {raw}") from exc
    if value <= 0:
        raise ConfigError(f"{name} must be greater than zero.")
    return value


def load_settings() -> Settings:
    google_service_account_file = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "").strip() or None
    google_service_account_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip() or None

    if not google_service_account_file and google_service_account_json:
        looks_like_file_path = (
            "\n" not in google_service_account_json
            and "\r" not in google_service_account_json
            and "{" not in google_service_account_json
            and google_service_account_json.lower().endswith(".json")
        )
        if looks_like_file_path:
            candidate_path = Path(google_service_account_json)
            if not candidate_path.is_absolute():
                candidate_path = Path.cwd() / candidate_path
            if candidate_path.exists():
                google_service_account_file = str(candidate_path)
                google_service_account_json = None

    if not google_service_account_file:
        credentials_dir = Path.cwd() / "Credentials"
        discovered_files = sorted(credentials_dir.glob("*.json"))
        if len(discovered_files) == 1:
            google_service_account_file = str(discovered_files[0])

    if not google_service_account_file and not google_service_account_json:
        raise ConfigError(
            "Set GOOGLE_SERVICE_ACCOUNT_FILE for local runs or GOOGLE_SERVICE_ACCOUNT_JSON for GitHub."
        )

    return Settings(
        bitrix_webhook_url=require_env("BITRIX_WEBHOOK_URL"),
        bitrix_entity_type=os.getenv("BITRIX_ENTITY_TYPE", "lead").strip().lower() or "lead",
        bitrix_date_field=os.getenv("BITRIX_DATE_FIELD", "DATE_CREATE").strip() or "DATE_CREATE",
        bitrix_utm_source_field=os.getenv("BITRIX_UTM_SOURCE_FIELD", "UTM_SOURCE").strip() or "UTM_SOURCE",
        bitrix_utm_medium_field=os.getenv("BITRIX_UTM_MEDIUM_FIELD", "UTM_MEDIUM").strip() or "UTM_MEDIUM",
        bitrix_utm_campaign_field=os.getenv("BITRIX_UTM_CAMPAIGN_FIELD", "UTM_CAMPAIGN").strip() or "UTM_CAMPAIGN",
        bitrix_request_timeout=read_int("BITRIX_REQUEST_TIMEOUT", 120),
        google_sheet_id=require_env("GOOGLE_SHEET_ID"),
        google_sheet_name=require_env("GOOGLE_SHEET_NAME"),
        google_allowed_range=require_env("GOOGLE_ALLOWED_RANGE"),
        google_source_summary_sheet_name=(
            os.getenv("GOOGLE_SOURCE_SUMMARY_SHEET_NAME", "Итог по источникам").strip()
            or "Итог по источникам"
        ),
        google_service_account_file=google_service_account_file,
        google_service_account_json=google_service_account_json,
        report_timezone=os.getenv("REPORT_TIMEZONE", "Asia/Yekaterinburg").strip() or "Asia/Yekaterinburg",
        report_period_mode=os.getenv("REPORT_PERIOD_MODE", "from_start_date").strip().lower()
        or "from_start_date",
        report_start_date=os.getenv("REPORT_START_DATE", "2026-03-01").strip() or "2026-03-01",
        report_require_all_utm=read_bool("REPORT_REQUIRE_ALL_UTM", True),
        report_unknown_source=os.getenv("REPORT_UNKNOWN_SOURCE", "Не определено").strip()
        or "Не определено",
    )


def normalize_text(value: Any) -> str:
    text = str(value or "").strip().lower().replace("ё", "е")
    text = re.sub(r"[^0-9a-zа-я_ ]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def raw_cell_text(value: Any) -> str:
    return str(value or "").strip().lower().replace("ё", "е")


def normalize_key(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip().lower())


def cell_matches_alias(value: Any, aliases: set[str]) -> bool:
    return normalize_text(value) in aliases or raw_cell_text(value) in aliases


def a1_to_col_index(column_label: str) -> int:
    value = 0
    for char in column_label:
        value = value * 26 + (ord(char.upper()) - ord("A") + 1)
    return value - 1


def parse_a1_range(range_text: str) -> SheetRange:
    value = range_text.strip()
    if "!" in value:
        _, value = value.split("!", 1)

    match = re.fullmatch(r"([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)", value)
    if not match:
        raise ConfigError(
            "GOOGLE_ALLOWED_RANGE must be a fixed A1 range like A1:J5000 or Sheet1!A1:J5000."
        )

    start_col, start_row, end_col, end_row = match.groups()
    return SheetRange(
        start_col=start_col,
        start_row=int(start_row),
        end_col=end_col,
        end_row=int(end_row),
    )


def pad_row(values: list[Any], width: int) -> list[Any]:
    row = list(values[:width])
    row.extend([""] * (width - len(row)))
    return row


def pad_rows(values: list[list[Any]], width: int) -> list[list[Any]]:
    return [pad_row(row, width) for row in values]


def build_google_credentials(settings: Settings) -> Credentials:
    if settings.google_service_account_file:
        credential_path = Path(settings.google_service_account_file)
        if not credential_path.is_absolute():
            credential_path = Path.cwd() / credential_path
        if not credential_path.exists():
            raise ConfigError(f"Google service account file not found: {credential_path}")
        return Credentials.from_service_account_file(str(credential_path), scopes=GOOGLE_SCOPES)

    assert settings.google_service_account_json is not None
    try:
        info = json.loads(settings.google_service_account_json)
    except json.JSONDecodeError as exc:
        raise ConfigError(
            "GOOGLE_SERVICE_ACCOUNT_JSON must contain valid JSON. In GitHub Secrets paste the full JSON content as-is."
        ) from exc

    return Credentials.from_service_account_info(info, scopes=GOOGLE_SCOPES)


def build_sheets_service(settings: Settings):
    credentials = build_google_credentials(settings)
    return build("sheets", "v4", credentials=credentials, cache_discovery=False)


def resolve_sheet_context(service: Any, settings: Settings) -> SheetContext:
    metadata = (
        service.spreadsheets()
        .get(
            spreadsheetId=settings.google_sheet_id,
            fields="properties(title),sheets(properties(sheetId,title,gridProperties),rowGroups)",
        )
        .execute()
    )

    spreadsheet_title = metadata["properties"]["title"]
    sheets = metadata.get("sheets", [])
    if not sheets:
        raise ConfigError("The target spreadsheet has no sheets.")

    requested_title = settings.google_sheet_name.strip()
    selected_sheet = next(
        (sheet for sheet in sheets if sheet["properties"]["title"] == requested_title),
        None,
    )

    if selected_sheet is None and requested_title == spreadsheet_title and len(sheets) == 1:
        selected_sheet = sheets[0]
    elif selected_sheet is None and len(sheets) == 1:
        selected_sheet = sheets[0]

    if selected_sheet is None:
        available = ", ".join(sheet["properties"]["title"] for sheet in sheets)
        raise ConfigError(
            f"Sheet '{requested_title}' not found. Available sheets: {available}"
        )

    allowed_range = parse_a1_range(settings.google_allowed_range)
    props = selected_sheet["properties"]
    return SheetContext(
        spreadsheet_id=settings.google_sheet_id,
        spreadsheet_title=spreadsheet_title,
        sheet_id=props["sheetId"],
        sheet_title=props["title"],
        allowed_range=allowed_range,
        row_count=props.get("gridProperties", {}).get("rowCount", allowed_range.end_row),
        row_groups=selected_sheet.get("rowGroups", []),
    )


def ensure_sheet_exists(
    service: Any,
    spreadsheet_id: str,
    sheet_title: str,
    row_count: int = 2000,
    column_count: int = 10,
) -> SheetWriteTarget:
    metadata = (
        service.spreadsheets()
        .get(spreadsheetId=spreadsheet_id, fields="sheets(properties(sheetId,title))")
        .execute()
    )
    for sheet in metadata.get("sheets", []):
        properties = sheet["properties"]
        if properties["title"] == sheet_title:
            return SheetWriteTarget(sheet_id=properties["sheetId"], sheet_title=properties["title"])

    response = (
        service.spreadsheets()
        .batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={
                "requests": [
                    {
                        "addSheet": {
                            "properties": {
                                "title": sheet_title,
                                "gridProperties": {
                                    "rowCount": row_count,
                                    "columnCount": column_count,
                                },
                            }
                        }
                    }
                ]
            },
        )
        .execute()
    )
    properties = response["replies"][0]["addSheet"]["properties"]
    return SheetWriteTarget(sheet_id=properties["sheetId"], sheet_title=properties["title"])


def normalize_webhook_base(url: str) -> str:
    parsed = urlparse(url.strip())
    path = parsed.path or "/"
    if path.endswith(".json"):
        path = path.rsplit("/", 1)[0] + "/"
    elif not path.endswith("/"):
        path += "/"
    return urlunparse(parsed._replace(path=path, params="", query="", fragment=""))


def build_bitrix_method_url(base_url: str, method_name: str) -> str:
    normalized_base = normalize_webhook_base(base_url)
    return f"{normalized_base}{method_name}.json"


def resolve_report_window(settings: Settings) -> ReportWindow:
    tz = ZoneInfo(settings.report_timezone)
    now = datetime.now(tz)
    start_date = datetime.strptime(settings.report_start_date, "%Y-%m-%d").date()
    start = datetime.combine(start_date, time.min, tzinfo=tz)
    end = now

    if settings.report_period_mode not in {"from_start_date", "current_month", "previous_month", "all_time"}:
        raise ConfigError(
            "Unsupported REPORT_PERIOD_MODE. Use from_start_date, current_month, previous_month, or all_time."
        )

    if settings.report_period_mode == "current_month":
        start = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    elif settings.report_period_mode == "previous_month":
        current_month_start = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        end = current_month_start - timedelta(seconds=1)
        start = end.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    elif settings.report_period_mode == "all_time":
        start = datetime(2000, 1, 1, tzinfo=tz)

    if start > end:
        raise ConfigError("Resolved report window is invalid: start date is later than end date.")

    return ReportWindow(start=start, end=end)


def bitrix_method_name(entity_type: str) -> str:
    if entity_type == "lead":
        return "crm.lead.list"
    if entity_type == "deal":
        return "crm.deal.list"
    raise ConfigError("BITRIX_ENTITY_TYPE must be lead or deal.")


def get_utm_field_candidates(settings: Settings) -> dict[str, list[str]]:
    candidates = {
        "source": [settings.bitrix_utm_source_field],
        "medium": [settings.bitrix_utm_medium_field],
        "campaign": [settings.bitrix_utm_campaign_field],
    }

    if settings.bitrix_entity_type == "lead":
        for key, fallback_fields in LEAD_FALLBACK_UTM_FIELDS.items():
            for field_name in fallback_fields:
                if field_name not in candidates[key]:
                    candidates[key].append(field_name)

    return candidates


def resolve_record_value(record: dict[str, Any], field_names: list[str]) -> str:
    for field_name in field_names:
        value = normalize_key(record.get(field_name))
        if value:
            return value
    return ""


def resolve_record_utm_key(record: dict[str, Any], field_candidates: dict[str, list[str]]) -> UtmKey:
    return UtmKey(
        utm_source=resolve_record_value(record, field_candidates["source"]),
        utm_medium=resolve_record_value(record, field_candidates["medium"]),
        utm_campaign=resolve_record_value(record, field_candidates["campaign"]),
    )


def build_bitrix_session() -> requests.Session:
    session = requests.Session()
    retries = Retry(
        total=3,
        connect=3,
        read=3,
        backoff_factor=1.5,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset({"GET"}),
    )
    session.mount("https://", HTTPAdapter(max_retries=retries))
    session.mount("http://", HTTPAdapter(max_retries=retries))
    return session


def execute_bitrix_list_request(
    session: requests.Session,
    settings: Settings,
    filters: dict[str, str],
    select_fields: list[str],
    start: int = 0,
) -> dict[str, Any]:
    method = bitrix_method_name(settings.bitrix_entity_type)
    method_url = build_bitrix_method_url(settings.bitrix_webhook_url, method)

    params: list[tuple[str, Any]] = [("start", start)]
    for field_name, field_value in filters.items():
        params.append((f"filter[{field_name}]", field_value))
    for field_name in select_fields:
        params.append(("select[]", field_name))

    response = session.get(method_url, params=params, timeout=settings.bitrix_request_timeout)
    response.raise_for_status()
    payload = response.json()

    if "error" in payload:
        raise RuntimeError(f"Bitrix API error: {payload['error']} - {payload.get('error_description', '')}")

    result = payload.get("result", [])
    if not isinstance(result, list):
        raise RuntimeError("Unexpected Bitrix API response: result is not a list.")

    return payload


def iterate_dates(start_date: date, end_date: date) -> list[date]:
    days: list[date] = []
    current = start_date
    while current <= end_date:
        days.append(current)
        current += timedelta(days=1)
    return days


def build_day_filters(settings: Settings, day_start: datetime, day_end: datetime) -> dict[str, str]:
    return {
        f">={settings.bitrix_date_field}": day_start.isoformat(timespec="seconds"),
        f"<={settings.bitrix_date_field}": day_end.isoformat(timespec="seconds"),
    }


def fetch_day_records(
    session: requests.Session,
    settings: Settings,
    day_start: datetime,
    day_end: datetime,
    select_fields: list[str],
) -> list[dict[str, Any]]:
    filters = build_day_filters(settings, day_start, day_end)
    records: list[dict[str, Any]] = []
    next_page: int | None = 0

    while next_page is not None:
        payload = execute_bitrix_list_request(
            session=session,
            settings=settings,
            filters=filters,
            select_fields=select_fields,
            start=next_page,
        )
        records.extend(payload.get("result", []))
        raw_next = payload.get("next")
        next_page = int(raw_next) if raw_next is not None else None

    return records


def build_daily_counts(settings: Settings, window: ReportWindow) -> tuple[dict[date, Counter[UtmKey]], int]:
    session = build_bitrix_session()
    day_counters: dict[date, Counter[UtmKey]] = {}
    counted_records = 0
    field_candidates = get_utm_field_candidates(settings)
    select_fields = ["ID", settings.bitrix_date_field]
    for field_name in (
        field_candidates["source"] + field_candidates["medium"] + field_candidates["campaign"]
    ):
        if field_name not in select_fields:
            select_fields.append(field_name)

    for current_date in iterate_dates(window.start.date(), window.end.date()):
        day_start = datetime.combine(current_date, time.min, tzinfo=window.start.tzinfo)
        day_end = datetime.combine(current_date, time.max, tzinfo=window.start.tzinfo)
        if current_date == window.end.date():
            day_end = window.end

        records = fetch_day_records(session, settings, day_start, day_end, select_fields)
        counter: Counter[UtmKey] = Counter()

        for record in records:
            key = resolve_record_utm_key(record, field_candidates)
            if not (key.utm_source and key.utm_medium and key.utm_campaign):
                if settings.report_require_all_utm:
                    continue

            if resolve_allowed_source_label(key) is None:
                continue

            counter[key] += 1
            counted_records += 1

        if counter:
            day_counters[current_date] = counter

    return day_counters, counted_records


def fetch_header_rows(service: Any, context: SheetContext) -> list[list[Any]]:
    header_end_row = min(context.allowed_range.start_row + 4, context.allowed_range.end_row)
    header_range = context.allowed_range.build_a1(context.sheet_title, row_end=header_end_row)
    values = (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=context.spreadsheet_id, range=header_range, majorDimension="ROWS")
        .execute()
        .get("values", [])
    )
    return pad_rows(values, context.allowed_range.width)


def find_header_columns(rows: list[list[Any]]) -> tuple[int, dict[str, int]]:
    required_columns = ("utm_source", "utm_medium", "utm_campaign", "date_created", "total")
    column_map: dict[str, int] = {}
    matched_rows: list[int] = []

    for canonical_name in required_columns:
        match = find_first_matching_alias(rows, HEADER_ALIASES[canonical_name])
        if match is None:
            raise ConfigError(
                "Could not find required columns: utm source, utm medium, utm campaign, Дата создания, Объем."
            )
        row_index, column_index = match
        column_map[canonical_name] = column_index
        matched_rows.append(row_index)

    number_match = find_first_matching_alias(rows, HEADER_ALIASES["number"])
    if number_match is not None:
        _, number_column = number_match
        column_map["number"] = number_column

    return max(matched_rows), column_map


def find_first_matching_alias(rows: list[list[Any]], aliases: set[str]) -> tuple[int, int] | None:
    for row_index, row in enumerate(rows):
        for column_index, cell in enumerate(row):
            if cell_matches_alias(cell, aliases):
                return row_index, column_index
    return None


def find_first_matching_column(rows: list[list[Any]], aliases: set[str]) -> int | None:
    for row in rows:
        for column_index, cell in enumerate(row):
            if cell_matches_alias(cell, aliases):
                return column_index
    return None


def sanitize_header_rows(rows: list[list[Any]]) -> list[list[Any]]:
    sanitized_rows = [list(row) for row in rows]
    matched_cells: list[tuple[int, int]] = []

    for row_index, row in enumerate(sanitized_rows):
        for column_index, cell in enumerate(row):
            if normalize_text(cell) in LEGACY_SUMMARY_BLOCK_LABELS:
                matched_cells.append((row_index, column_index))

    if not matched_cells:
        return sanitized_rows

    min_row = min(row_index for row_index, _ in matched_cells)
    max_row = max(row_index for row_index, _ in matched_cells)
    min_col = min(column_index for _, column_index in matched_cells)
    max_col = max(column_index for _, column_index in matched_cells)

    for row_index in range(min_row, max_row + 1):
        for column_index in range(min_col, max_col + 1):
            sanitized_rows[row_index][column_index] = ""

    return sanitized_rows


def month_summary_label(year: int, month: int) -> str:
    return f"Итого за {MONTH_LABELS_RU[month]} {year}"


def day_summary_label(current_date: date) -> str:
    return f"Итого за {current_date.isoformat()}"


def month_period_label(year: int, month: int) -> str:
    return f"{MONTH_LABELS_RU[month]} {year}"


def resolve_source_label(raw_source: str, unknown_source: str) -> str:
    normalized_source = normalize_key(raw_source)
    if not normalized_source:
        return unknown_source
    return normalized_source


def resolve_allowed_source_label(key: UtmKey) -> str | None:
    return ALLOWED_UTM_RULES.get((key.utm_source, key.utm_medium, key.utm_campaign))


def format_sheet_date(current_date: date) -> str:
    return current_date.strftime("%d.%m.%Y")


def build_row(width: int) -> list[Any]:
    return [""] * width


def build_report_rows(
    daily_counts: dict[date, Counter[UtmKey]],
    width: int,
    header_rows: list[list[Any]],
    header_row_index: int,
    column_map: dict[str, int],
) -> ReportBuildResult:
    output_rows = [list(row) for row in header_rows[: header_row_index + 1]]
    detail_number = 1
    month_groups: list[tuple[int, int]] = []
    day_groups: list[tuple[int, int]] = []
    summary_rows: list[int] = []
    month_count = 0
    day_count = 0
    detail_count = 0
    record_count = sum(sum(counter.values()) for counter in daily_counts.values())

    date_column = column_map["date_created"]
    total_column = column_map["total"]
    number_column = column_map.get("number")

    grouped_by_month: dict[tuple[int, int], dict[date, Counter[UtmKey]]] = defaultdict(dict)
    for current_date, counter in sorted(daily_counts.items()):
        grouped_by_month[(current_date.year, current_date.month)][current_date] = counter

    for month_key in sorted(grouped_by_month):
        year, month = month_key
        month_days = grouped_by_month[month_key]
        month_total = sum(sum(counter.values()) for counter in month_days.values())
        month_row = build_row(width)
        month_row[date_column] = month_summary_label(year, month)
        month_row[total_column] = month_total
        output_rows.append(month_row)
        summary_rows.append(len(output_rows))
        month_count += 1

        month_group_start = len(output_rows) + 1
        month_child_started = False

        for current_date in sorted(month_days):
            day_counter = month_days[current_date]
            day_total = sum(day_counter.values())

            day_row = build_row(width)
            day_row[date_column] = day_summary_label(current_date)
            day_row[total_column] = day_total
            output_rows.append(day_row)
            summary_rows.append(len(output_rows))
            day_count += 1
            month_child_started = True

            day_group_start = len(output_rows) + 1
            day_detail_started = False

            for key in sorted(
                day_counter,
                key=lambda item: (
                    item.utm_source,
                    item.utm_medium,
                    item.utm_campaign,
                ),
            ):
                detail_row = build_row(width)
                if number_column is not None:
                    detail_row[number_column] = detail_number
                detail_row[column_map["utm_source"]] = key.utm_source
                detail_row[column_map["utm_medium"]] = key.utm_medium
                detail_row[column_map["utm_campaign"]] = key.utm_campaign
                detail_row[date_column] = format_sheet_date(current_date)
                detail_row[total_column] = day_counter[key]
                output_rows.append(detail_row)
                detail_number += 1
                detail_count += 1
                day_detail_started = True

            if day_detail_started:
                day_groups.append((day_group_start, len(output_rows)))

        if month_child_started:
            month_groups.append((month_group_start, len(output_rows)))

    return ReportBuildResult(
        rows=output_rows,
        month_groups=month_groups,
        day_groups=day_groups,
        summary_rows=summary_rows,
        month_count=month_count,
        day_count=day_count,
        detail_count=detail_count,
        record_count=record_count,
    )


def build_source_summary_rows(
    daily_counts: dict[date, Counter[UtmKey]],
    unknown_source: str,
) -> list[list[Any]]:
    monthly_source_totals: dict[tuple[int, int], Counter[str]] = defaultdict(Counter)
    overall_source_totals: Counter[str] = Counter()

    for current_date, counter in daily_counts.items():
        month_key = (current_date.year, current_date.month)
        for utm_key, amount in counter.items():
            source_label = resolve_allowed_source_label(utm_key) or unknown_source
            monthly_source_totals[month_key][source_label] += amount
            overall_source_totals[source_label] += amount

    rows: list[list[Any]] = [["Период", "Источник", "Суммарный объем"]]
    for month_key in sorted(monthly_source_totals):
        year, month = month_key
        period_label = month_period_label(year, month)
        for source_label in sorted(monthly_source_totals[month_key]):
            rows.append([period_label, source_label, monthly_source_totals[month_key][source_label]])

    if rows == [["Период", "Источник", "Суммарный объем"]]:
        rows.append(["Все время", unknown_source, 0])
        return rows

    rows.append([])
    rows.append(["Все время", "", ""])
    for source_label in sorted(overall_source_totals):
        rows.append(["Все время", source_label, overall_source_totals[source_label]])

    return rows


def clear_report_values(service: Any, context: SheetContext, header_rows: int) -> None:
    clear_start_row = context.allowed_range.start_row + header_rows
    if clear_start_row > context.allowed_range.end_row:
        return
    clear_range = context.allowed_range.build_a1(
        context.sheet_title,
        row_start=clear_start_row,
        row_end=context.allowed_range.end_row,
    )
    service.spreadsheets().values().clear(
        spreadsheetId=context.spreadsheet_id,
        range=clear_range,
        body={},
    ).execute()


def clear_sheet_values(service: Any, spreadsheet_id: str, sheet_title: str) -> None:
    escaped_title = sheet_title.replace("'", "''")
    quoted_title = sheet_title if re.fullmatch(r"[A-Za-z0-9_]+", sheet_title) else f"'{escaped_title}'"
    service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range=f"{quoted_title}!A:Z",
        body={},
    ).execute()


def write_report_rows(
    service: Any,
    context: SheetContext,
    rows: list[list[Any]],
) -> None:
    if not rows:
        return

    row_end = context.allowed_range.start_row + len(rows) - 1
    write_range = context.allowed_range.build_a1(context.sheet_title, row_end=row_end)
    service.spreadsheets().values().update(
        spreadsheetId=context.spreadsheet_id,
        range=write_range,
        valueInputOption="USER_ENTERED",
        body={"majorDimension": "ROWS", "values": rows},
    ).execute()


def write_sheet_rows(
    service: Any,
    spreadsheet_id: str,
    sheet_title: str,
    rows: list[list[Any]],
) -> None:
    if not rows:
        return
    escaped_title = sheet_title.replace("'", "''")
    quoted_title = sheet_title if re.fullmatch(r"[A-Za-z0-9_]+", sheet_title) else f"'{escaped_title}'"
    end_row = len(rows)
    end_col = chr(ord("A") + max(len(row) for row in rows) - 1)
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"{quoted_title}!A1:{end_col}{end_row}",
        valueInputOption="USER_ENTERED",
        body={"majorDimension": "ROWS", "values": rows},
    ).execute()


def delete_existing_row_groups(service: Any, context: SheetContext) -> None:
    if not context.row_groups:
        return

    sorted_groups = sorted(
        context.row_groups,
        key=lambda group: (
            group.get("depth", 0),
            group.get("range", {}).get("startIndex", 0),
            group.get("range", {}).get("endIndex", 0),
        ),
        reverse=True,
    )
    requests = [
        {
            "deleteDimensionGroup": {
                "range": {
                    "sheetId": context.sheet_id,
                    "dimension": "ROWS",
                    "startIndex": group["range"]["startIndex"],
                    "endIndex": group["range"]["endIndex"],
                }
            }
        }
        for group in sorted_groups
    ]

    service.spreadsheets().batchUpdate(
        spreadsheetId=context.spreadsheet_id,
        body={"requests": requests},
    ).execute()


def build_group_requests(context: SheetContext, group_ranges: list[tuple[int, int]]) -> list[dict[str, Any]]:
    requests: list[dict[str, Any]] = []
    for start_row, end_row in group_ranges:
        if end_row < start_row:
            continue
        requests.append(
            {
                "addDimensionGroup": {
                    "range": {
                        "sheetId": context.sheet_id,
                        "dimension": "ROWS",
                        "startIndex": start_row - 1,
                        "endIndex": end_row,
                    }
                }
            }
        )
    return requests


def apply_row_groups(service: Any, context: SheetContext, result: ReportBuildResult) -> None:
    delete_existing_row_groups(service, context)
    requests = []
    requests.extend(build_group_requests(context, result.month_groups))
    requests.extend(build_group_requests(context, result.day_groups))
    if not requests:
        return
    service.spreadsheets().batchUpdate(
        spreadsheetId=context.spreadsheet_id,
        body={"requests": requests},
    ).execute()


def apply_summary_row_formatting(service: Any, context: SheetContext, result: ReportBuildResult) -> None:
    if not result.summary_rows:
        return

    requests = [
        {
            "repeatCell": {
                "range": {
                    "sheetId": context.sheet_id,
                    "startRowIndex": row_number - 1,
                    "endRowIndex": row_number,
                    "startColumnIndex": a1_to_col_index(context.allowed_range.start_col),
                    "endColumnIndex": a1_to_col_index(context.allowed_range.end_col) + 1,
                },
                "cell": {
                    "userEnteredFormat": {
                        "horizontalAlignment": "CENTER",
                        "textFormat": {"bold": True},
                    }
                },
                "fields": "userEnteredFormat(horizontalAlignment,textFormat.bold)",
            }
        }
        for row_number in result.summary_rows
    ]

    service.spreadsheets().batchUpdate(
        spreadsheetId=context.spreadsheet_id,
        body={"requests": requests},
    ).execute()


def print_summary(context: SheetContext, result: ReportBuildResult, dry_run: bool) -> None:
    mode_label = "Dry run" if dry_run else "Sync completed"
    print(f"{mode_label}: {context.sheet_title}")
    print(f"Counted records: {result.record_count}")
    print(f"Month groups: {result.month_count}")
    print(f"Day groups: {result.day_count}")
    print(f"Detail rows: {result.detail_count}")
    print(f"Written rows: {len(result.rows)}")


def main() -> int:
    try:
        args = parse_args()
        load_environment(args.env_file)
        settings = load_settings()
        window = resolve_report_window(settings)
        sheets_service = build_sheets_service(settings)
        context = resolve_sheet_context(sheets_service, settings)
        header_rows = sanitize_header_rows(fetch_header_rows(sheets_service, context))
        header_row_index, column_map = find_header_columns(header_rows)

        day_counters, _ = build_daily_counts(settings, window)
        result = build_report_rows(
            daily_counts=day_counters,
            width=context.allowed_range.width,
            header_rows=header_rows,
            header_row_index=header_row_index,
            column_map=column_map,
        )
        source_summary_rows = build_source_summary_rows(
            daily_counts=day_counters,
            unknown_source=settings.report_unknown_source,
        )

        if not args.dry_run:
            clear_report_values(sheets_service, context, header_row_index + 1)
            write_report_rows(sheets_service, context, result.rows)
            apply_row_groups(sheets_service, context, result)
            apply_summary_row_formatting(sheets_service, context, result)
            summary_target = ensure_sheet_exists(
                sheets_service,
                context.spreadsheet_id,
                settings.google_source_summary_sheet_name,
            )
            clear_sheet_values(
                sheets_service,
                context.spreadsheet_id,
                summary_target.sheet_title,
            )
            write_sheet_rows(
                sheets_service,
                context.spreadsheet_id,
                summary_target.sheet_title,
                source_summary_rows,
            )

        print_summary(context, result, args.dry_run)
        return 0
    except ConfigError as exc:
        print(f"Configuration error: {exc}", file=sys.stderr)
        return 2
    except requests.Timeout as exc:
        print(f"Timeout error: {exc}", file=sys.stderr)
        return 4
    except requests.HTTPError as exc:
        print(f"HTTP error: {exc}", file=sys.stderr)
        return 3
    except Exception as exc:  # pragma: no cover
        print(f"Unhandled error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
