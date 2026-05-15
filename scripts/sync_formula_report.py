from __future__ import annotations

import argparse
import json
import os
import re
import sys
from collections import defaultdict
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
KC_DASHBOARD_URL = "https://formula-agency.github.io/otchety/"
DEFAULT_MEETING_LOG_SHEET_ID = "1CNT1xTe5uBHo4W4ZLUh3qZLmgWy7wxe7nSsCtDXwwIo"
DEFAULT_MEETING_LOG_SHEET_NAME = "Meetings"
LEAD_FALLBACK_UTM_FIELDS = {
    "source": ["UF_LEAD_FIRST_UTM_SOURCE"],
    "medium": ["UF_LEAD_FIRST_UTM_MEDIUM"],
    "campaign": ["UF_LEAD_FIRST_UTM_CAMPAIGN"],
}
DEAL_FALLBACK_UTM_FIELDS = {
    "source": ["UF_DEAL_FIRST_UTM_SOURCE"],
    "medium": ["UF_DEAL_FIRST_UTM_MEDIUM"],
    "campaign": ["UF_DEAL_FIRST_UTM_CAMPAIGN"],
}
DEFAULT_DEAL_APPROVED_MORTGAGE_FIELD = "UF_DEAL_MORTGAGE_APPROVED"
DEFAULT_DEAL_MEETING_SHOW_FIELD = "UF_DEAL_SHOW"
DEFAULT_DEAL_RESERVATION_FIELD = "UF_DEAL_WHERE_PUT_RESERVATION"
METRIC_COLUMN_TITLES = {
    "approved_mortgage": "Одобрена ипотека",
    "meeting_show": "Проведена встреча/показ",
    "reservation": "Зафиксирована бронь",
    "closed_deals": "Закрыто сделок",
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
    "approved_mortgage": {"одобрена ипотека", "одобренная ипотека", "approved mortgage"},
    "meeting_show": {
        "проведена встреча/показ",
        "проведена встреча",
        "проведен показ",
        "встреча/показ",
        "встречи и показы",
    },
    "reservation": {"зафиксирована бронь", "бронь", "бронирование", "забронирована бронь"},
    "closed_deals": {"закрыто сделок", "закрыто", "закрытые сделки", "closed deals"},
}
MEETING_LOG_HEADER_ALIASES = {
    "status": {"статус", "status", "результат", "result"},
    "meeting_start": {"начало встречи", "start", "meeting start"},
    "deal_id": {"id сделки", "deal id", "id deal"},
    "deal_link": {"ссылка на сделку", "deal link", "link"},
}
SUCCESSFUL_MEETING_STATUSES = {"прошла успешно"}
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


@dataclass
class UtmMetrics:
    records: int = 0
    deals: int = 0
    approved_mortgage: int = 0
    meeting_show: int = 0
    reservation: int = 0
    closed: int = 0

    def add(
        self,
        approved_mortgage: bool,
        meeting_show: bool,
        reservation: bool,
        closed: bool,
    ) -> None:
        self.records += 1
        self.approved_mortgage += int(bool(approved_mortgage))
        self.meeting_show += int(bool(meeting_show))
        self.reservation += int(bool(reservation))
        self.closed += int(bool(closed))

    def merge(self, other: "UtmMetrics") -> "UtmMetrics":
        return UtmMetrics(
            records=self.records + other.records,
            deals=self.deals + other.deals,
            approved_mortgage=self.approved_mortgage + other.approved_mortgage,
            meeting_show=self.meeting_show + other.meeting_show,
            reservation=self.reservation + other.reservation,
            closed=self.closed + other.closed,
        )


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
    bitrix_stage_field: str
    bitrix_approved_mortgage_field: str | None
    bitrix_meeting_show_field: str | None
    bitrix_reservation_field: str | None
    bitrix_success_stage_ids: tuple[str, ...]
    bitrix_request_timeout: int
    google_sheet_id: str
    google_sheet_name: str
    google_allowed_range: str
    google_source_summary_sheet_name: str
    google_meeting_log_sheet_id: str
    google_meeting_log_sheet_name: str
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
    month_summary_rows: list[int]
    day_summary_rows: list[int]
    summary_label_column: int
    summary_total_column: int
    month_count: int
    day_count: int
    detail_count: int
    record_count: int


@dataclass
class SheetWriteTarget:
    sheet_id: int
    sheet_title: str


@dataclass(frozen=True)
class MeetingLogEntry:
    meeting_date: date
    deal_id: str


@dataclass(frozen=True)
class AttributionCaches:
    deal_cache: dict[str, dict[str, Any]]
    lead_cache: dict[str, dict[str, Any]]
    contact_cache: dict[str, dict[str, Any]]
    phone_utm_cache: dict[str, UtmKey | None]


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

    bitrix_entity_type = os.getenv("BITRIX_ENTITY_TYPE", "lead").strip().lower() or "lead"
    default_stage_field = "STAGE_ID" if bitrix_entity_type == "deal" else "STATUS_ID"
    raw_success_stages = os.getenv("BITRIX_SUCCESS_STAGE_IDS", "WON,CLOSED").strip()
    if raw_success_stages:
        bitrix_success_stage_ids = tuple(
            normalize_key(stage) for stage in raw_success_stages.split(",") if stage.strip()
        )
    else:
        bitrix_success_stage_ids = ("won", "closed", "contract", "success", "complete")

    return Settings(
        bitrix_webhook_url=require_env("BITRIX_WEBHOOK_URL"),
        bitrix_entity_type=bitrix_entity_type,
        bitrix_date_field=os.getenv("BITRIX_DATE_FIELD", "DATE_CREATE").strip() or "DATE_CREATE",
        bitrix_utm_source_field=os.getenv("BITRIX_UTM_SOURCE_FIELD", "UTM_SOURCE").strip() or "UTM_SOURCE",
        bitrix_utm_medium_field=os.getenv("BITRIX_UTM_MEDIUM_FIELD", "UTM_MEDIUM").strip() or "UTM_MEDIUM",
        bitrix_utm_campaign_field=os.getenv("BITRIX_UTM_CAMPAIGN_FIELD", "UTM_CAMPAIGN").strip() or "UTM_CAMPAIGN",
        bitrix_stage_field=os.getenv("BITRIX_STAGE_FIELD", default_stage_field).strip() or default_stage_field,
        bitrix_approved_mortgage_field=(
            os.getenv("BITRIX_APPROVED_MORTGAGE_FIELD", "").strip() or DEFAULT_DEAL_APPROVED_MORTGAGE_FIELD
        ),
        bitrix_meeting_show_field=(
            os.getenv("BITRIX_MEETING_SHOW_FIELD", "").strip() or DEFAULT_DEAL_MEETING_SHOW_FIELD
        ),
        bitrix_reservation_field=(
            os.getenv("BITRIX_RESERVATION_FIELD", "").strip() or DEFAULT_DEAL_RESERVATION_FIELD
        ),
        bitrix_success_stage_ids=bitrix_success_stage_ids,
        bitrix_request_timeout=read_int("BITRIX_REQUEST_TIMEOUT", 120),
        google_sheet_id=require_env("GOOGLE_SHEET_ID"),
        google_sheet_name=require_env("GOOGLE_SHEET_NAME"),
        google_allowed_range=require_env("GOOGLE_ALLOWED_RANGE"),
        google_source_summary_sheet_name=(
            os.getenv("GOOGLE_SOURCE_SUMMARY_SHEET_NAME", "Итог по источникам").strip()
            or "Итог по источникам"
        ),
        google_meeting_log_sheet_id=(
            os.getenv("GOOGLE_MEETING_LOG_SHEET_ID", DEFAULT_MEETING_LOG_SHEET_ID).strip()
            or DEFAULT_MEETING_LOG_SHEET_ID
        ),
        google_meeting_log_sheet_name=(
            os.getenv("GOOGLE_MEETING_LOG_SHEET_NAME", DEFAULT_MEETING_LOG_SHEET_NAME).strip()
            or DEFAULT_MEETING_LOG_SHEET_NAME
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


def resolve_sheet_title(
    service: Any,
    spreadsheet_id: str,
    requested_title: str,
) -> str:
    metadata = (
        service.spreadsheets()
        .get(spreadsheetId=spreadsheet_id, fields="properties(title),sheets(properties(title))")
        .execute()
    )

    sheets = metadata.get("sheets", [])
    if not sheets:
        raise ConfigError("The target spreadsheet has no sheets.")

    spreadsheet_title = metadata["properties"]["title"]
    normalized_title = requested_title.strip()
    selected_sheet = next(
        (sheet for sheet in sheets if sheet["properties"]["title"] == normalized_title),
        None,
    )

    if selected_sheet is None and normalized_title == spreadsheet_title and len(sheets) == 1:
        selected_sheet = sheets[0]
    elif selected_sheet is None and len(sheets) == 1:
        selected_sheet = sheets[0]

    if selected_sheet is None:
        available = ", ".join(sheet["properties"]["title"] for sheet in sheets)
        raise ConfigError(
            f"Sheet '{normalized_title}' not found in spreadsheet {spreadsheet_id}. Available sheets: {available}"
        )

    return selected_sheet["properties"]["title"]


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


def get_deal_utm_field_candidates(settings: Settings) -> dict[str, list[str]]:
    candidates = {
        "source": [settings.bitrix_utm_source_field],
        "medium": [settings.bitrix_utm_medium_field],
        "campaign": [settings.bitrix_utm_campaign_field],
    }

    for key, fallback_fields in DEAL_FALLBACK_UTM_FIELDS.items():
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


def resolve_allowed_utm_key(key: UtmKey) -> UtmKey | None:
    return key if resolve_allowed_source_label(key) is not None else None


def resolve_boolean_field(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    if isinstance(value, (int, float)):
        return value != 0
    text = str(value).strip().lower()
    return text in {
        "1",
        "true",
        "yes",
        "y",
        "да",
        "on",
        "ok",
        "checked",
        "t",
    }


def resolve_non_empty_field(value: Any) -> bool:
    if value is None:
        return False
    if isinstance(value, (list, tuple, set, dict)):
        return len(value) > 0
    if isinstance(value, str):
        return bool(value.strip())
    return bool(value)


def resolve_deal_closed(record: dict[str, Any], settings: Settings) -> bool:
    semantic_value = normalize_key(record.get("STAGE_SEMANTIC_ID"))
    if semantic_value == "s":
        return True
    stage_value = record.get(settings.bitrix_stage_field)
    if stage_value is None:
        return False
    return normalize_key(stage_value) in settings.bitrix_success_stage_ids


def resolve_deal_record_metrics(
    record: dict[str, Any],
    settings: Settings,
) -> UtmMetrics:
    approved = False
    reservation = False
    closed = False

    if settings.bitrix_approved_mortgage_field:
        approved = resolve_boolean_field(record.get(settings.bitrix_approved_mortgage_field))

    if settings.bitrix_reservation_field:
        reservation = resolve_non_empty_field(record.get(settings.bitrix_reservation_field))

    closed = resolve_deal_closed(record, settings)

    metrics = UtmMetrics()
    metrics.deals = 1
    metrics.approved_mortgage = int(approved)
    metrics.reservation = int(reservation)
    metrics.closed = int(closed)
    return metrics


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
    entity_type: str,
    filters: dict[str, str],
    select_fields: list[str],
    start: int = 0,
) -> dict[str, Any]:
    method = bitrix_method_name(entity_type)
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


def execute_bitrix_get_request(
    session: requests.Session,
    settings: Settings,
    method_name: str,
    params: dict[str, Any],
) -> dict[str, Any]:
    method_url = build_bitrix_method_url(settings.bitrix_webhook_url, method_name)
    response = session.get(method_url, params=params, timeout=settings.bitrix_request_timeout)
    response.raise_for_status()
    payload = response.json()

    if "error" in payload:
        raise RuntimeError(f"Bitrix API error: {payload['error']} - {payload.get('error_description', '')}")

    result = payload.get("result")
    if not isinstance(result, dict):
        raise RuntimeError("Unexpected Bitrix API response: result is not an object.")

    return result


def iterate_dates(start_date: date, end_date: date) -> list[date]:
    days: list[date] = []
    current = start_date
    while current <= end_date:
        days.append(current)
        current += timedelta(days=1)
    return days


def build_day_filters(date_field: str, day_start: datetime, day_end: datetime) -> dict[str, str]:
    return {
        f">={date_field}": day_start.isoformat(timespec="seconds"),
        f"<={date_field}": day_end.isoformat(timespec="seconds"),
    }


def fetch_day_records_for_entity(
    session: requests.Session,
    settings: Settings,
    entity_type: str,
    date_field: str,
    day_start: datetime,
    day_end: datetime,
    select_fields: list[str],
) -> list[dict[str, Any]]:
    filters = build_day_filters(date_field, day_start, day_end)
    records: list[dict[str, Any]] = []
    next_page: int | None = 0

    while next_page is not None:
        payload = execute_bitrix_list_request(
            session=session,
            settings=settings,
            entity_type=entity_type,
            filters=filters,
            select_fields=select_fields,
            start=next_page,
        )
        records.extend(payload.get("result", []))
        raw_next = payload.get("next")
        next_page = int(raw_next) if raw_next is not None else None

    return records


def parse_sheet_datetime(value: Any, timezone_name: str) -> datetime | None:
    text = str(value or "").strip()
    if not text:
        return None

    tz = ZoneInfo(timezone_name)
    normalized = text.replace("T", " ").replace("/", "-")

    for parser in (datetime.fromisoformat,):
        try:
            parsed = parser(normalized)
            break
        except ValueError:
            parsed = None
    else:
        parsed = None

    if parsed is None:
        for pattern in (
            "%Y-%m-%d %H:%M:%S",
            "%Y-%m-%d %H:%M",
            "%d.%m.%Y %H:%M:%S",
            "%d.%m.%Y %H:%M",
            "%Y-%m-%d",
            "%d.%m.%Y",
        ):
            try:
                parsed = datetime.strptime(normalized, pattern)
                break
            except ValueError:
                continue

    if parsed is None:
        return None
    if parsed.tzinfo is None:
        return parsed.replace(tzinfo=tz)
    return parsed.astimezone(tz)


def parse_bitrix_datetime(value: Any, timezone_name: str) -> datetime | None:
    text = str(value or "").strip()
    if not text:
        return None

    tz = ZoneInfo(timezone_name)
    normalized = text.replace("Z", "+00:00")
    try:
        parsed = datetime.fromisoformat(normalized)
    except ValueError:
        parsed = parse_sheet_datetime(text, timezone_name)

    if parsed is None:
        return None
    if parsed.tzinfo is None:
        return parsed.replace(tzinfo=tz)
    return parsed.astimezone(tz)


def normalize_phone(value: Any) -> str:
    digits = "".join(char for char in str(value or "") if char.isdigit())
    if len(digits) == 11 and digits.startswith("8"):
        digits = "7" + digits[1:]
    return digits


def extract_entity_phone_numbers(record: dict[str, Any]) -> list[str]:
    raw_values = record.get("PHONE") or []
    phones: list[str] = []

    if isinstance(raw_values, dict):
        raw_values = [raw_values]
    elif not isinstance(raw_values, list):
        raw_values = [raw_values]

    for item in raw_values:
        if isinstance(item, dict):
            value = item.get("VALUE")
        else:
            value = item
        phone = normalize_phone(value)
        if phone:
            phones.append(phone)

    return phones


def build_successful_meeting_entries(service: Any, settings: Settings) -> list[MeetingLogEntry]:
    sheet_title = resolve_sheet_title(
        service,
        settings.google_meeting_log_sheet_id,
        settings.google_meeting_log_sheet_name,
    )
    values = (
        service.spreadsheets()
        .values()
        .get(
            spreadsheetId=settings.google_meeting_log_sheet_id,
            range=f"'{sheet_title}'!A:Z",
            majorDimension="ROWS",
        )
        .execute()
        .get("values", [])
    )

    if not values:
        return []

    header_row_index, column_map = find_meeting_log_columns(values)
    entries: list[MeetingLogEntry] = []

    for row in values[header_row_index + 1 :]:
        status_index = column_map["status"]
        status_value = row[status_index] if status_index < len(row) else ""
        if normalize_text(status_value) not in SUCCESSFUL_MEETING_STATUSES:
            continue

        meeting_start_index = column_map["meeting_start"]
        meeting_start_value = row[meeting_start_index] if meeting_start_index < len(row) else ""
        meeting_datetime = parse_sheet_datetime(meeting_start_value, settings.report_timezone)
        if meeting_datetime is None:
            continue

        deal_id = ""
        deal_id_index = column_map.get("deal_id")
        if deal_id_index is not None and deal_id_index < len(row):
            deal_id = parse_numeric_id(row[deal_id_index])

        if not deal_id:
            deal_link_index = column_map.get("deal_link")
            if deal_link_index is not None and deal_link_index < len(row):
                deal_id = extract_deal_id_from_link(row[deal_link_index])

        if not deal_id:
            continue

        entries.append(MeetingLogEntry(meeting_date=meeting_datetime.date(), deal_id=deal_id))

    return entries


def build_primary_daily_counts(settings: Settings, window: ReportWindow) -> tuple[dict[date, dict[UtmKey, UtmMetrics]], int]:
    session = build_bitrix_session()
    day_counters: dict[date, dict[UtmKey, UtmMetrics]] = {}
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

        records = fetch_day_records_for_entity(
            session=session,
            settings=settings,
            entity_type=settings.bitrix_entity_type,
            date_field=settings.bitrix_date_field,
            day_start=day_start,
            day_end=day_end,
            select_fields=select_fields,
        )
        counter: dict[UtmKey, UtmMetrics] = defaultdict(UtmMetrics)

        for record in records:
            key = resolve_record_utm_key(record, field_candidates)
            if not (key.utm_source and key.utm_medium and key.utm_campaign):
                if settings.report_require_all_utm:
                    continue

            if resolve_allowed_source_label(key) is None:
                continue

            counter[key].records += 1
            counted_records += 1

        if counter:
            day_counters[current_date] = counter

    return day_counters, counted_records


def build_daily_deal_metrics(
    settings: Settings,
    window: ReportWindow,
) -> dict[date, dict[UtmKey, UtmMetrics]]:
    session = build_bitrix_session()
    day_counters: dict[date, dict[UtmKey, UtmMetrics]] = {}
    field_candidates = get_deal_utm_field_candidates(settings)
    caches = AttributionCaches(
        deal_cache={},
        lead_cache={},
        contact_cache={},
        phone_utm_cache={},
    )

    select_fields = [
        "ID",
        "DATE_CREATE",
        "STAGE_ID",
        "STAGE_SEMANTIC_ID",
        "LEAD_ID",
        "CONTACT_ID",
    ]
    for field_name in (
        field_candidates["source"]
        + field_candidates["medium"]
        + field_candidates["campaign"]
        + [
            settings.bitrix_approved_mortgage_field,
            settings.bitrix_reservation_field,
        ]
    ):
        if field_name and field_name not in select_fields:
            select_fields.append(field_name)

    for current_date in iterate_dates(window.start.date(), window.end.date()):
        day_start = datetime.combine(current_date, time.min, tzinfo=window.start.tzinfo)
        day_end = datetime.combine(current_date, time.max, tzinfo=window.start.tzinfo)
        if current_date == window.end.date():
            day_end = window.end

        records = fetch_day_records_for_entity(
            session=session,
            settings=settings,
            entity_type="deal",
            date_field="DATE_CREATE",
            day_start=day_start,
            day_end=day_end,
            select_fields=select_fields,
        )
        counter: dict[UtmKey, UtmMetrics] = defaultdict(UtmMetrics)

        for record in records:
            deal_id = str(record.get("ID") or "").strip()
            key = resolve_allowed_utm_key_for_deal(
                session=session,
                settings=settings,
                deal_id=deal_id,
                caches=caches,
                deal_record=record,
            )
            if key is None:
                continue

            metrics = resolve_deal_record_metrics(record, settings)
            target = counter[key]
            target.deals += metrics.deals
            target.approved_mortgage += metrics.approved_mortgage
            target.reservation += metrics.reservation
            target.closed += metrics.closed

        if counter:
            day_counters[current_date] = counter

    return day_counters


def resolve_allowed_utm_key_for_deal(
    session: requests.Session,
    settings: Settings,
    deal_id: str,
    caches: AttributionCaches,
    deal_record: dict[str, Any] | None = None,
) -> UtmKey | None:
    if not deal_id and not deal_record:
        return None

    deal = deal_record
    if deal is None and deal_id:
        deal = caches.deal_cache.get(deal_id)
        if deal is None:
            deal = execute_bitrix_get_request(session, settings, "crm.deal.get", {"id": deal_id})
            caches.deal_cache[deal_id] = deal

    if deal is None:
        return None

    resolved_deal_id = deal_id or str(deal.get("ID") or "").strip()
    if resolved_deal_id and resolved_deal_id not in caches.deal_cache:
        caches.deal_cache[resolved_deal_id] = deal

    deal_key = resolve_allowed_utm_key(
        resolve_record_utm_key(deal, get_deal_utm_field_candidates(settings))
    )
    if deal_key:
        return deal_key

    lead_id = str(deal.get("LEAD_ID") or "").strip()
    if lead_id:
        lead = caches.lead_cache.get(lead_id)
        if lead is None:
            lead = execute_bitrix_get_request(session, settings, "crm.lead.get", {"id": lead_id})
            caches.lead_cache[lead_id] = lead

        lead_key = resolve_allowed_utm_key(
            resolve_record_utm_key(lead, get_utm_field_candidates(settings))
        )
        if lead_key:
            return lead_key

        for phone in extract_entity_phone_numbers(lead):
            if phone not in caches.phone_utm_cache:
                caches.phone_utm_cache[phone] = find_first_allowed_utm_key_by_phone(session, settings, phone)
            if caches.phone_utm_cache[phone]:
                return caches.phone_utm_cache[phone]

    contact_id = str(deal.get("CONTACT_ID") or "").strip()
    if contact_id:
        contact = caches.contact_cache.get(contact_id)
        if contact is None:
            contact = execute_bitrix_get_request(session, settings, "crm.contact.get", {"id": contact_id})
            caches.contact_cache[contact_id] = contact
        for phone in extract_entity_phone_numbers(contact):
            if phone not in caches.phone_utm_cache:
                caches.phone_utm_cache[phone] = find_first_allowed_utm_key_by_phone(session, settings, phone)
            if caches.phone_utm_cache[phone]:
                return caches.phone_utm_cache[phone]

    return None


def phone_search_variants(phone: str) -> list[str]:
    variants = [phone]
    if len(phone) == 11 and phone.startswith("7"):
        variants.append("8" + phone[1:])
        variants.append("+" + phone)
        variants.append("+7 (" + phone[1:4] + ") " + phone[4:7] + "-" + phone[7:9] + "-" + phone[9:11])
    elif len(phone) == 10:
        variants.append("7" + phone)
        variants.append("8" + phone)
    return list(dict.fromkeys(filter(None, variants)))


def find_first_allowed_utm_key_by_phone(
    session: requests.Session,
    settings: Settings,
    phone: str,
) -> UtmKey | None:
    lead_fields = get_utm_field_candidates(settings)
    select_fields = ["ID", "DATE_CREATE", "PHONE"] + lead_fields["source"] + lead_fields["medium"] + lead_fields["campaign"]
    collected_records: dict[str, dict[str, Any]] = {}

    for variant in phone_search_variants(phone):
        start = 0
        while True:
            try:
                payload = execute_bitrix_list_request(
                    session=session,
                    settings=settings,
                    entity_type="lead",
                    filters={"PHONE": variant},
                    select_fields=select_fields,
                    start=start,
                )
            except RuntimeError:
                break

            for record in payload.get("result", []):
                record_id = str(record.get("ID") or "").strip()
                if record_id:
                    collected_records[record_id] = record

            raw_next = payload.get("next")
            if raw_next is None:
                break
            start = int(raw_next)

    sorted_records = sorted(
        collected_records.values(),
        key=lambda record: (
            parse_bitrix_datetime(record.get("DATE_CREATE"), settings.report_timezone) or datetime.max.replace(tzinfo=ZoneInfo(settings.report_timezone)),
            int(str(record.get("ID") or "0")),
        ),
    )

    for record in sorted_records:
        key = resolve_allowed_utm_key(resolve_record_utm_key(record, lead_fields))
        if key:
            return key

    return None


def build_daily_meeting_metrics(
    service: Any,
    settings: Settings,
    window: ReportWindow,
) -> dict[date, dict[UtmKey, UtmMetrics]]:
    session = build_bitrix_session()
    meeting_entries = build_successful_meeting_entries(service, settings)
    daily_counts: dict[date, dict[UtmKey, UtmMetrics]] = defaultdict(lambda: defaultdict(UtmMetrics))
    caches = AttributionCaches(
        deal_cache={},
        lead_cache={},
        contact_cache={},
        phone_utm_cache={},
    )

    for entry in meeting_entries:
        if entry.meeting_date < window.start.date() or entry.meeting_date > window.end.date():
            continue

        utm_key = resolve_allowed_utm_key_for_deal(
            session=session,
            settings=settings,
            deal_id=entry.deal_id,
            caches=caches,
        )
        if utm_key is None:
            continue
        daily_counts[entry.meeting_date][utm_key].meeting_show += 1

    return {current_date: dict(counter) for current_date, counter in daily_counts.items()}


def overlay_deal_metrics(
    primary_counts: dict[date, dict[UtmKey, UtmMetrics]],
    deal_metrics: dict[date, dict[UtmKey, UtmMetrics]],
) -> dict[date, dict[UtmKey, UtmMetrics]]:
    merged_counts: dict[date, dict[UtmKey, UtmMetrics]] = {
        current_date: {
            key: UtmMetrics(
                records=metrics.records,
                deals=metrics.deals,
                approved_mortgage=metrics.approved_mortgage,
                meeting_show=metrics.meeting_show,
                reservation=metrics.reservation,
                closed=metrics.closed,
            )
            for key, metrics in day_counter.items()
        }
        for current_date, day_counter in primary_counts.items()
    }

    for current_date, day_counter in deal_metrics.items():
        target_counter = merged_counts.setdefault(current_date, {})
        for key, metrics in day_counter.items():
            target = target_counter.setdefault(key, UtmMetrics())
            target.deals += metrics.deals
            target.approved_mortgage += metrics.approved_mortgage
            target.meeting_show += metrics.meeting_show
            target.reservation += metrics.reservation
            target.closed += metrics.closed

    return merged_counts


def overlay_meeting_metrics(
    base_counts: dict[date, dict[UtmKey, UtmMetrics]],
    meeting_counts: dict[date, dict[UtmKey, UtmMetrics]],
) -> dict[date, dict[UtmKey, UtmMetrics]]:
    merged_counts: dict[date, dict[UtmKey, UtmMetrics]] = {
        current_date: {
            key: UtmMetrics(
                records=metrics.records,
                deals=metrics.deals,
                approved_mortgage=metrics.approved_mortgage,
                meeting_show=metrics.meeting_show,
                reservation=metrics.reservation,
                closed=metrics.closed,
            )
            for key, metrics in day_counter.items()
        }
        for current_date, day_counter in base_counts.items()
    }

    for current_date, day_counter in meeting_counts.items():
        target_counter = merged_counts.setdefault(current_date, {})
        for key, metrics in day_counter.items():
            target = target_counter.setdefault(key, UtmMetrics())
            target.meeting_show += metrics.meeting_show

    return merged_counts


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
    optional_columns = (
        "approved_mortgage",
        "meeting_show",
        "reservation",
        "closed_deals",
        "number",
    )
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

    for canonical_name in optional_columns:
        match = find_first_matching_alias(rows, HEADER_ALIASES[canonical_name])
        if match is not None:
            _, column_index = match
            column_map[canonical_name] = column_index

    return max(matched_rows), column_map


def ensure_metric_header_columns(
    rows: list[list[Any]],
    column_map: dict[str, int],
) -> tuple[list[list[Any]], dict[str, int]]:
    updated_rows = [list(row) for row in rows]
    next_column = column_map["total"] + 1

    for canonical_name in ("approved_mortgage", "meeting_show", "reservation", "closed_deals"):
        if canonical_name in column_map:
            continue
        if next_column >= len(updated_rows[0]):
            raise ConfigError(
                "GOOGLE_ALLOWED_RANGE is too narrow for the required report columns."
            )
        updated_rows[0][next_column] = METRIC_COLUMN_TITLES[canonical_name]
        if len(updated_rows) > 1:
            updated_rows[1][next_column] = ""
        column_map[canonical_name] = next_column
        next_column += 1

    return updated_rows, column_map


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


def parse_numeric_id(value: Any) -> str:
    match = re.search(r"\d+", str(value or ""))
    return match.group(0) if match else ""


def extract_deal_id_from_link(value: Any) -> str:
    match = re.search(r"/crm/deal/details/(\d+)/?", str(value or ""), flags=re.IGNORECASE)
    return match.group(1) if match else ""


def find_meeting_log_columns(rows: list[list[Any]]) -> tuple[int, dict[str, int]]:
    required_columns = ("status", "meeting_start")
    optional_columns = ("deal_id", "deal_link")
    column_map: dict[str, int] = {}
    matched_rows: list[int] = []

    for canonical_name in required_columns:
        match = find_first_matching_alias(rows, MEETING_LOG_HEADER_ALIASES[canonical_name])
        if match is None:
            raise ConfigError(
                "Could not find required meeting log columns. Expected 'Результат/Статус' and 'Начало встречи'."
            )
        row_index, column_index = match
        column_map[canonical_name] = column_index
        matched_rows.append(row_index)

    for canonical_name in optional_columns:
        match = find_first_matching_alias(rows, MEETING_LOG_HEADER_ALIASES[canonical_name])
        if match is not None:
            row_index, column_index = match
            column_map[canonical_name] = column_index
            matched_rows.append(row_index)

    if "deal_id" not in column_map and "deal_link" not in column_map:
        raise ConfigError(
            "Could not find 'ID сделки' or 'Ссылка на сделку' columns in the meeting log sheet."
        )

    return max(matched_rows), column_map


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
    return f"Итого за {format_sheet_date(current_date)}"


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
    daily_counts: dict[date, dict[UtmKey, UtmMetrics]],
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
    month_summary_rows: list[int] = []
    day_summary_rows: list[int] = []
    month_count = 0
    day_count = 0
    detail_count = 0
    record_count = sum(
        sum(metrics.records for metrics in counter.values())
        for counter in daily_counts.values()
    )

    date_column = column_map["date_created"]
    total_column = column_map["total"]
    number_column = column_map.get("number")
    label_column = column_map.get("utm_source", date_column)
    approved_column = column_map.get("approved_mortgage")
    meeting_column = column_map.get("meeting_show")
    reservation_column = column_map.get("reservation")
    closed_column = column_map.get("closed_deals")

    grouped_by_month: dict[tuple[int, int], dict[date, dict[UtmKey, UtmMetrics]]] = defaultdict(dict)
    for current_date, counter in sorted(daily_counts.items()):
        grouped_by_month[(current_date.year, current_date.month)][current_date] = counter

    for month_key in sorted(grouped_by_month):
        year, month = month_key
        month_days = grouped_by_month[month_key]
        month_total = sum(
            metrics.records for counter in month_days.values() for metrics in counter.values()
        )
        month_approved = sum(
            metrics.approved_mortgage for counter in month_days.values() for metrics in counter.values()
        )
        month_meetings = sum(
            metrics.meeting_show for counter in month_days.values() for metrics in counter.values()
        )
        month_reservations = sum(
            metrics.reservation for counter in month_days.values() for metrics in counter.values()
        )
        month_closed = sum(
            metrics.closed for counter in month_days.values() for metrics in counter.values()
        )

        month_row = build_row(width)
        month_row[label_column] = month_summary_label(year, month)
        month_row[total_column] = month_total
        if approved_column is not None:
            month_row[approved_column] = month_approved
        if meeting_column is not None:
            month_row[meeting_column] = month_meetings
        if reservation_column is not None:
            month_row[reservation_column] = month_reservations
        if closed_column is not None:
            month_row[closed_column] = month_closed
        output_rows.append(month_row)
        summary_rows.append(len(output_rows))
        month_summary_rows.append(len(output_rows))
        month_count += 1

        month_group_start = len(output_rows) + 1
        month_child_started = False

        for current_date in sorted(month_days):
            day_counter = month_days[current_date]
            day_total = sum(metrics.records for metrics in day_counter.values())
            day_approved = sum(metrics.approved_mortgage for metrics in day_counter.values())
            day_meetings = sum(metrics.meeting_show for metrics in day_counter.values())
            day_reservations = sum(metrics.reservation for metrics in day_counter.values())
            day_closed = sum(metrics.closed for metrics in day_counter.values())

            day_row = build_row(width)
            day_row[label_column] = day_summary_label(current_date)
            day_row[total_column] = day_total
            if approved_column is not None:
                day_row[approved_column] = day_approved
            if meeting_column is not None:
                day_row[meeting_column] = day_meetings
            if reservation_column is not None:
                day_row[reservation_column] = day_reservations
            if closed_column is not None:
                day_row[closed_column] = day_closed
            output_rows.append(day_row)
            summary_rows.append(len(output_rows))
            day_summary_rows.append(len(output_rows))
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
                metrics = day_counter[key]
                detail_row[total_column] = metrics.records
                if approved_column is not None:
                    detail_row[approved_column] = metrics.approved_mortgage
                if meeting_column is not None:
                    detail_row[meeting_column] = metrics.meeting_show
                if reservation_column is not None:
                    detail_row[reservation_column] = metrics.reservation
                if closed_column is not None:
                    detail_row[closed_column] = metrics.closed
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
        month_summary_rows=month_summary_rows,
        day_summary_rows=day_summary_rows,
        summary_label_column=label_column,
        summary_total_column=total_column,
        month_count=month_count,
        day_count=day_count,
        detail_count=detail_count,
        record_count=record_count,
    )


def build_source_summary_rows(
    daily_counts: dict[date, dict[UtmKey, UtmMetrics]],
    unknown_source: str,
) -> list[list[Any]]:
    monthly_source_totals: dict[tuple[int, int], dict[str, UtmMetrics]] = defaultdict(dict)
    overall_source_totals: dict[str, UtmMetrics] = {}

    for current_date, counter in daily_counts.items():
        month_key = (current_date.year, current_date.month)
        for utm_key, metrics in counter.items():
            source_label = resolve_allowed_source_label(utm_key) or unknown_source
            month_metrics = monthly_source_totals[month_key].setdefault(source_label, UtmMetrics())
            month_metrics.records += metrics.records
            month_metrics.approved_mortgage += metrics.approved_mortgage
            month_metrics.meeting_show += metrics.meeting_show
            month_metrics.reservation += metrics.reservation
            month_metrics.closed += metrics.closed

            overall_metrics = overall_source_totals.setdefault(source_label, UtmMetrics())
            overall_metrics.records += metrics.records
            overall_metrics.approved_mortgage += metrics.approved_mortgage
            overall_metrics.meeting_show += metrics.meeting_show
            overall_metrics.reservation += metrics.reservation
            overall_metrics.closed += metrics.closed

    rows: list[list[Any]] = [[
        "Период",
        "Источник",
        "Суммарный объем",
        "Одобрена ипотека",
        "Проведена встреча/показ",
        "Зафиксирована бронь",
        "Закрыто сделок",
    ]]
    for month_key in sorted(monthly_source_totals):
        year, month = month_key
        period_label = month_period_label(year, month)
        for source_label in sorted(monthly_source_totals[month_key]):
            metrics = monthly_source_totals[month_key][source_label]
            rows.append([
                period_label,
                source_label,
                metrics.records,
                metrics.approved_mortgage,
                metrics.meeting_show,
                metrics.reservation,
                metrics.closed,
            ])

    if rows == [[
        "Период",
        "Источник",
        "Суммарный объем",
        "Одобрена ипотека",
        "Проведена встреча/показ",
        "Зафиксирована бронь",
        "Закрыто сделок",
    ]]:
        rows.append(["Все время", unknown_source, 0, 0, 0, 0, 0])
        return rows

    rows.append([])
    rows.append(["Все время", "", "", "", "", "", ""])
    for source_label in sorted(overall_source_totals):
        metrics = overall_source_totals[source_label]
        rows.append([
            "Все время",
            source_label,
            metrics.records,
            metrics.approved_mortgage,
            metrics.meeting_show,
            metrics.reservation,
            metrics.closed,
        ])

    return rows


def build_dashboard_payload(
    daily_counts: dict[date, dict[UtmKey, UtmMetrics]],
    settings: Settings,
    window: ReportWindow,
) -> dict[str, Any]:
    dashboard_rows: list[dict[str, Any]] = []
    source_labels: set[str] = set()
    utm_sources: set[str] = set()
    utm_mediums: set[str] = set()
    utm_campaigns: set[str] = set()
    utm_combinations: set[tuple[str, str, str]] = set()

    totals = {
        "records": 0,
        "deals": 0,
        "approvedMortgage": 0,
        "meetingShow": 0,
        "reservation": 0,
        "sales": 0,
        "closedDeals": 0,
    }

    for current_date in sorted(daily_counts):
        current_date_iso = current_date.isoformat()
        current_month = current_date.strftime("%Y-%m")
        for key in sorted(
            daily_counts[current_date],
            key=lambda item: (item.utm_source, item.utm_medium, item.utm_campaign),
        ):
            metrics = daily_counts[current_date][key]
            source_label = resolve_allowed_source_label(key) or settings.report_unknown_source

            source_labels.add(source_label)
            utm_sources.add(key.utm_source)
            utm_mediums.add(key.utm_medium)
            utm_campaigns.add(key.utm_campaign)
            utm_combinations.add((key.utm_source, key.utm_medium, key.utm_campaign))

            totals["records"] += metrics.records
            totals["deals"] += metrics.deals
            totals["approvedMortgage"] += metrics.approved_mortgage
            totals["meetingShow"] += metrics.meeting_show
            totals["reservation"] += metrics.reservation
            totals["sales"] += metrics.closed
            totals["closedDeals"] += metrics.closed

            dashboard_rows.append(
                {
                    "month": current_month,
                    "uploadDate": current_date_iso,
                    "sourceLabel": source_label,
                    "utmSource": key.utm_source,
                    "utmMedium": key.utm_medium,
                    "utmCampaign": key.utm_campaign,
                    "records": metrics.records,
                    "deals": metrics.deals,
                    "approvedMortgage": metrics.approved_mortgage,
                    "meetingShow": metrics.meeting_show,
                    "reservation": metrics.reservation,
                    "sales": metrics.closed,
                    "closedDeals": metrics.closed,
                }
            )

    generated_at = datetime.now(ZoneInfo(settings.report_timezone)).isoformat()
    table_url = f"https://docs.google.com/spreadsheets/d/{settings.google_sheet_id}/edit?gid=0#gid=0"

    return {
        "report": {
            "name": "Отчет ОП",
            "from": window.start.date().isoformat(),
            "to": window.end.date().isoformat(),
            "timezone": settings.report_timezone,
            "tableUrl": table_url,
            "kcDashboardUrl": KC_DASHBOARD_URL,
            "opDashboardUrl": "./",
        },
        "generatedAt": generated_at,
        "filters": {
            "sourceLabels": sorted(source_labels),
            "utmSources": sorted(utm_sources),
            "utmMediums": sorted(utm_mediums),
            "utmCampaigns": sorted(utm_campaigns),
            "minDate": dashboard_rows[0]["uploadDate"] if dashboard_rows else "",
            "maxDate": dashboard_rows[-1]["uploadDate"] if dashboard_rows else "",
        },
        "totals": totals,
        "overview": {
            "sourceCount": len(source_labels),
            "utmCombinationCount": len(utm_combinations),
        },
        "baseRows": dashboard_rows,
    }


def write_dashboard_files(payload: dict[str, Any]) -> None:
    dashboard_data_dir = Path("dashboard") / "data"
    dashboard_data_dir.mkdir(parents=True, exist_ok=True)

    json_path = dashboard_data_dir / "report-data.json"
    js_path = dashboard_data_dir / "report-data.js"

    json_text = json.dumps(payload, ensure_ascii=False, indent=2)
    json_path.write_text(f"{json_text}\n", encoding="utf-8")
    js_path.write_text(
        f"window.REPORT_DASHBOARD_DATA = {json_text};\n",
        encoding="utf-8",
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


def clear_report_body_merges(service: Any, context: SheetContext, header_rows: int) -> None:
    merge_start_row = context.allowed_range.start_row + header_rows
    if merge_start_row > context.allowed_range.end_row:
        return
    service.spreadsheets().batchUpdate(
        spreadsheetId=context.spreadsheet_id,
        body={
            "requests": [
                {
                    "unmergeCells": {
                        "range": {
                            "sheetId": context.sheet_id,
                            "startRowIndex": merge_start_row - 1,
                            "endRowIndex": context.allowed_range.end_row,
                            "startColumnIndex": a1_to_col_index(context.allowed_range.start_col),
                            "endColumnIndex": a1_to_col_index(context.allowed_range.end_col) + 1,
                        }
                    }
                }
            ]
        },
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

    base_col = a1_to_col_index(context.allowed_range.start_col)
    end_col = a1_to_col_index(context.allowed_range.end_col) + 1
    requests: list[dict[str, Any]] = []

    for row_number in result.summary_rows:
        font_size = 12 if row_number in result.month_summary_rows else 11
        requests.append(
            {
                "repeatCell": {
                    "range": {
                        "sheetId": context.sheet_id,
                        "startRowIndex": row_number - 1,
                        "endRowIndex": row_number,
                        "startColumnIndex": base_col,
                        "endColumnIndex": end_col,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "horizontalAlignment": "CENTER",
                            "textFormat": {
                                "bold": True,
                                "fontSize": font_size,
                            },
                        }
                    },
                    "fields": "userEnteredFormat(horizontalAlignment,textFormat.bold,textFormat.fontSize)",
                }
            }
        )

        if result.summary_label_column < result.summary_total_column:
            requests.append(
                {
                    "mergeCells": {
                        "range": {
                            "sheetId": context.sheet_id,
                            "startRowIndex": row_number - 1,
                            "endRowIndex": row_number,
                            "startColumnIndex": base_col + result.summary_label_column,
                            "endColumnIndex": base_col + result.summary_total_column,
                        },
                        "mergeType": "MERGE_ALL",
                    }
                }
            )

    service.spreadsheets().batchUpdate(
        spreadsheetId=context.spreadsheet_id,
        body={"requests": requests},
    ).execute()


def apply_metric_header_formatting(service: Any, context: SheetContext, column_map: dict[str, int]) -> None:
    base_col = a1_to_col_index(context.allowed_range.start_col)
    metric_columns = [
        column_map.get(canonical_name)
        for canonical_name in ("approved_mortgage", "meeting_show", "reservation", "closed_deals")
        if column_map.get(canonical_name) is not None
    ]
    if not metric_columns:
        return

    requests: list[dict[str, Any]] = []
    requests.append(
        {
            "unmergeCells": {
                "range": {
                    "sheetId": context.sheet_id,
                    "startRowIndex": context.allowed_range.start_row - 1,
                    "endRowIndex": context.allowed_range.start_row + 1,
                    "startColumnIndex": base_col + min(metric_columns),
                    "endColumnIndex": base_col + max(metric_columns) + 1,
                }
            }
        }
    )

    for canonical_name in ("approved_mortgage", "meeting_show", "reservation", "closed_deals"):
        column_index = column_map.get(canonical_name)
        if column_index is None:
            continue
        requests.append(
            {
                "repeatCell": {
                    "range": {
                        "sheetId": context.sheet_id,
                        "startRowIndex": context.allowed_range.start_row - 1,
                        "endRowIndex": context.allowed_range.start_row + 1,
                        "startColumnIndex": base_col + column_index,
                        "endColumnIndex": base_col + column_index + 1,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {
                                "red": 0.7882353,
                                "green": 0.85490197,
                                "blue": 0.972549,
                            },
                            "horizontalAlignment": "CENTER",
                            "textFormat": {
                                "fontFamily": "Arial",
                                "fontSize": 12,
                                "bold": False,
                            },
                            "borders": {
                                "top": {"style": "SOLID_MEDIUM"},
                                "bottom": {"style": "SOLID_MEDIUM"},
                                "left": {"style": "SOLID_MEDIUM"},
                                "right": {"style": "SOLID_MEDIUM"},
                            },
                        }
                    },
                    "fields": (
                        "userEnteredFormat(backgroundColor,horizontalAlignment,"
                        "textFormat.fontFamily,textFormat.fontSize,textFormat.bold,borders)"
                    ),
                }
            }
        )
        requests.append(
            {
                "mergeCells": {
                    "range": {
                        "sheetId": context.sheet_id,
                        "startRowIndex": context.allowed_range.start_row - 1,
                        "endRowIndex": context.allowed_range.start_row + 1,
                        "startColumnIndex": base_col + column_index,
                        "endColumnIndex": base_col + column_index + 1,
                    },
                    "mergeType": "MERGE_ALL",
                }
            }
        )

    service.spreadsheets().batchUpdate(
        spreadsheetId=context.spreadsheet_id,
        body={"requests": requests},
    ).execute()


def apply_report_body_alignment(
    service: Any,
    context: SheetContext,
    header_rows: int,
    written_rows: int,
) -> None:
    body_start_row = context.allowed_range.start_row + header_rows
    body_end_row = context.allowed_range.start_row + written_rows - 1
    if body_start_row > body_end_row:
        return
    service.spreadsheets().batchUpdate(
        spreadsheetId=context.spreadsheet_id,
        body={
            "requests": [
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": context.sheet_id,
                            "startRowIndex": body_start_row - 1,
                            "endRowIndex": body_end_row,
                            "startColumnIndex": a1_to_col_index(context.allowed_range.start_col),
                            "endColumnIndex": a1_to_col_index(context.allowed_range.end_col) + 1,
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "horizontalAlignment": "LEFT",
                            }
                        },
                        "fields": "userEnteredFormat.horizontalAlignment",
                    }
                }
            ]
        },
    ).execute()


def apply_detail_date_format(
    service: Any,
    context: SheetContext,
    result: ReportBuildResult,
    header_rows: int,
    date_column: int,
) -> None:
    first_body_row = context.allowed_range.start_row + header_rows
    last_body_row = context.allowed_range.start_row + len(result.rows) - 1
    summary_row_set = set(result.summary_rows)
    requests: list[dict[str, Any]] = []

    for row_number in range(first_body_row, last_body_row + 1):
        if row_number in summary_row_set:
            continue
        requests.append(
            {
                "repeatCell": {
                    "range": {
                        "sheetId": context.sheet_id,
                        "startRowIndex": row_number - 1,
                        "endRowIndex": row_number,
                        "startColumnIndex": a1_to_col_index(context.allowed_range.start_col) + date_column,
                        "endColumnIndex": a1_to_col_index(context.allowed_range.start_col) + date_column + 1,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {
                                "type": "DATE",
                                "pattern": "dd.mm.yyyy",
                            }
                        }
                    },
                    "fields": "userEnteredFormat.numberFormat",
                }
            }
        )

    if requests:
        service.spreadsheets().batchUpdate(
            spreadsheetId=context.spreadsheet_id,
            body={"requests": requests},
        ).execute()


def rewrite_date_column_as_text(
    service: Any,
    context: SheetContext,
    result: ReportBuildResult,
    header_rows: int,
    date_column: int,
) -> None:
    body_start_row = context.allowed_range.start_row + header_rows
    if body_start_row > context.allowed_range.end_row:
        return

    date_col_label = chr(ord("A") + a1_to_col_index(context.allowed_range.start_col) + date_column)
    body_values = [[row[date_column] if date_column < len(row) else ""] for row in result.rows[header_rows:]]
    escaped_title = context.sheet_title.replace("'", "''")
    quoted_title = context.sheet_title if re.fullmatch(r"[A-Za-z0-9_]+", context.sheet_title) else f"'{escaped_title}'"
    end_row = body_start_row + len(body_values) - 1
    service.spreadsheets().values().update(
        spreadsheetId=context.spreadsheet_id,
        range=f"{quoted_title}!{date_col_label}{body_start_row}:{date_col_label}{end_row}",
        valueInputOption="RAW",
        body={"majorDimension": "ROWS", "values": body_values},
    ).execute()


def apply_sheet_alignment(
    service: Any,
    spreadsheet_id: str,
    sheet_id: int,
    row_count: int,
    column_count: int,
) -> None:
    if row_count <= 0 or column_count <= 0:
        return

    requests: list[dict[str, Any]] = [
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 0,
                    "endRowIndex": 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": column_count,
                },
                "cell": {
                    "userEnteredFormat": {
                        "horizontalAlignment": "CENTER",
                    }
                },
                "fields": "userEnteredFormat.horizontalAlignment",
            }
        }
    ]

    if row_count > 1:
        requests.append(
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 1,
                        "endRowIndex": row_count,
                        "startColumnIndex": 0,
                        "endColumnIndex": column_count,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "horizontalAlignment": "LEFT",
                        }
                    },
                    "fields": "userEnteredFormat.horizontalAlignment",
                }
            }
        )

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
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
        header_rows, column_map = ensure_metric_header_columns(header_rows, column_map)

        primary_day_counters, _ = build_primary_daily_counts(settings, window)
        deal_day_metrics = build_daily_deal_metrics(settings, window)
        meeting_day_metrics = build_daily_meeting_metrics(sheets_service, settings, window)
        day_counters = overlay_deal_metrics(primary_day_counters, deal_day_metrics)
        day_counters = overlay_meeting_metrics(day_counters, meeting_day_metrics)
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
        dashboard_payload = build_dashboard_payload(
            daily_counts=day_counters,
            settings=settings,
            window=window,
        )
        write_dashboard_files(dashboard_payload)

        if not args.dry_run:
            clear_report_values(sheets_service, context, header_row_index + 1)
            clear_report_body_merges(sheets_service, context, header_row_index + 1)
            write_report_rows(sheets_service, context, result.rows)
            rewrite_date_column_as_text(
                sheets_service,
                context,
                result,
                header_row_index + 1,
                column_map["date_created"],
            )
            apply_metric_header_formatting(sheets_service, context, column_map)
            apply_row_groups(sheets_service, context, result)
            apply_summary_row_formatting(sheets_service, context, result)
            apply_report_body_alignment(
                sheets_service,
                context,
                header_row_index + 1,
                len(result.rows),
            )
            apply_detail_date_format(
                sheets_service,
                context,
                result,
                header_row_index + 1,
                column_map["date_created"],
            )
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
            apply_sheet_alignment(
                sheets_service,
                context.spreadsheet_id,
                summary_target.sheet_id,
                len(source_summary_rows),
                max(len(row) for row in source_summary_rows),
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
