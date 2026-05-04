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
SOURCE_MAP = {
    "leadit": "Лидген КЦ",
    "selfwalk": "Самоход",
    "avito": "Авито",
    "recommendation": "Рекомендация",
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
    month_count: int
    day_count: int
    detail_count: int
    record_count: int


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


def normalize_key(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip().lower())


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
) -> list[dict[str, Any]]:
    filters = build_day_filters(settings, day_start, day_end)
    select_fields = [
        "ID",
        settings.bitrix_date_field,
        settings.bitrix_utm_source_field,
        settings.bitrix_utm_medium_field,
        settings.bitrix_utm_campaign_field,
    ]
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


def record_has_all_utm(record: dict[str, Any], settings: Settings) -> bool:
    return all(
        normalize_key(record.get(field))
        for field in (
            settings.bitrix_utm_source_field,
            settings.bitrix_utm_medium_field,
            settings.bitrix_utm_campaign_field,
        )
    )


def build_daily_counts(settings: Settings, window: ReportWindow) -> tuple[dict[date, Counter[UtmKey]], int]:
    session = build_bitrix_session()
    day_counters: dict[date, Counter[UtmKey]] = {}
    counted_records = 0

    for current_date in iterate_dates(window.start.date(), window.end.date()):
        day_start = datetime.combine(current_date, time.min, tzinfo=window.start.tzinfo)
        day_end = datetime.combine(current_date, time.max, tzinfo=window.start.tzinfo)
        if current_date == window.end.date():
            day_end = window.end

        records = fetch_day_records(session, settings, day_start, day_end)
        counter: Counter[UtmKey] = Counter()

        for record in records:
            if settings.report_require_all_utm and not record_has_all_utm(record, settings):
                continue

            key = UtmKey(
                utm_source=normalize_key(record.get(settings.bitrix_utm_source_field)),
                utm_medium=normalize_key(record.get(settings.bitrix_utm_medium_field)),
                utm_campaign=normalize_key(record.get(settings.bitrix_utm_campaign_field)),
            )
            if not (key.utm_source and key.utm_medium and key.utm_campaign):
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
    for row_index, row in enumerate(rows):
        normalized_cells = [normalize_text(cell) for cell in row]
        column_map: dict[str, int] = {}
        for canonical_name, aliases in HEADER_ALIASES.items():
            for column_index, cell in enumerate(normalized_cells):
                if cell in aliases:
                    column_map[canonical_name] = column_index
                    break

        required = {"utm_source", "utm_medium", "utm_campaign", "period", "source", "total"}
        if required.issubset(column_map):
            return row_index, column_map

    raise ConfigError(
        "Could not find a header row with columns: utm source, utm medium, utm campaign, Период, Источник, Суммарный объем."
    )


def find_first_matching_column(rows: list[list[Any]], aliases: set[str]) -> int | None:
    for row in rows:
        for column_index, cell in enumerate(row):
            if normalize_text(cell) in aliases:
                return column_index
    return None


def month_summary_label(year: int, month: int) -> str:
    return f"Итого за {MONTH_LABELS_RU[month]} {year}"


def day_summary_label(current_date: date) -> str:
    return f"Итого за {current_date.isoformat()}"


def build_row(width: int) -> list[Any]:
    return [""] * width


def build_report_rows(
    daily_counts: dict[date, Counter[UtmKey]],
    width: int,
    header_rows: list[list[Any]],
    header_row_index: int,
    column_map: dict[str, int],
    unknown_source: str,
) -> ReportBuildResult:
    output_rows = [list(row) for row in header_rows[: header_row_index + 1]]
    detail_number = 1
    month_groups: list[tuple[int, int]] = []
    day_groups: list[tuple[int, int]] = []
    month_count = 0
    day_count = 0
    detail_count = 0
    record_count = sum(sum(counter.values()) for counter in daily_counts.values())

    date_column = column_map.get("date_created")
    number_column = column_map.get("number")

    grouped_by_month: dict[tuple[int, int], dict[date, Counter[UtmKey]]] = defaultdict(dict)
    for current_date, counter in sorted(daily_counts.items()):
        grouped_by_month[(current_date.year, current_date.month)][current_date] = counter

    for month_key in sorted(grouped_by_month):
        year, month = month_key
        month_days = grouped_by_month[month_key]
        month_total = sum(sum(counter.values()) for counter in month_days.values())
        month_row = build_row(width)
        month_row[column_map["period"]] = month_summary_label(year, month)
        month_row[column_map["total"]] = month_total
        output_rows.append(month_row)
        month_count += 1

        month_group_start = len(output_rows) + 1
        month_child_started = False

        for current_date in sorted(month_days):
            day_counter = month_days[current_date]
            day_total = sum(day_counter.values())

            day_row = build_row(width)
            day_row[column_map["period"]] = day_summary_label(current_date)
            day_row[column_map["total"]] = day_total
            output_rows.append(day_row)
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
                if date_column is not None:
                    detail_row[date_column] = current_date.isoformat()
                detail_row[column_map["source"]] = SOURCE_MAP.get(key.utm_source, unknown_source)
                detail_row[column_map["total"]] = day_counter[key]
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
        month_count=month_count,
        day_count=day_count,
        detail_count=detail_count,
        record_count=record_count,
    )


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
        header_rows = fetch_header_rows(sheets_service, context)
        header_row_index, column_map = find_header_columns(header_rows)

        if "number" not in column_map:
            number_column = find_first_matching_column(header_rows, HEADER_ALIASES["number"])
            if number_column is not None:
                column_map["number"] = number_column
        if "date_created" not in column_map:
            date_column = find_first_matching_column(header_rows, HEADER_ALIASES["date_created"])
            if date_column is not None:
                column_map["date_created"] = date_column

        day_counters, _ = build_daily_counts(settings, window)
        result = build_report_rows(
            daily_counts=day_counters,
            width=context.allowed_range.width,
            header_rows=header_rows,
            header_row_index=header_row_index,
            column_map=column_map,
            unknown_source=settings.report_unknown_source,
        )

        if not args.dry_run:
            clear_report_values(sheets_service, context, header_row_index + 1)
            write_report_rows(sheets_service, context, result.rows)
            apply_row_groups(sheets_service, context, result)

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
