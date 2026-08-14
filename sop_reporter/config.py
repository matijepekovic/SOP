from __future__ import annotations

import fnmatch
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Mapping, Sequence

import yaml

from sop_reporter.exceptions import ConfigurationError


SUPPORTED_TYPES = {
    "text",
    "integer",
    "number",
    "currency",
    "percent",
    "date",
    "datetime",
    "boolean",
}
SUPPORTED_FILTERS = {
    "equals",
    "not_equals",
    "contains",
    "not_contains",
    "in",
    "not_in",
    "range",
    "gt",
    "gte",
    "lt",
    "lte",
    "is_blank",
    "not_blank",
}
SUPPORTED_AGGREGATIONS = {"sum", "count", "count_rows", "avg", "min", "max"}
WEEKDAYS = {
    "monday",
    "tuesday",
    "wednesday",
    "thursday",
    "friday",
    "saturday",
    "sunday",
}


def _load_yaml(path: Path) -> dict[str, Any]:
    try:
        with path.open("r", encoding="utf-8") as handle:
            loaded = yaml.safe_load(handle)
    except FileNotFoundError as exc:
        raise ConfigurationError(f"Configuration file not found: {path}") from exc
    except yaml.YAMLError as exc:
        raise ConfigurationError(f"Invalid YAML in {path}: {exc}") from exc
    except OSError as exc:
        raise ConfigurationError(f"Cannot read configuration file {path}: {exc}") from exc
    if not isinstance(loaded, dict):
        raise ConfigurationError(f"Configuration root must be a mapping: {path}")
    return loaded


def _mapping(value: Any, location: str) -> Mapping[str, Any]:
    if not isinstance(value, Mapping):
        raise ConfigurationError(f"{location} must be a mapping")
    return value


def _sequence(value: Any, location: str) -> Sequence[Any]:
    if isinstance(value, (str, bytes)) or not isinstance(value, Sequence):
        raise ConfigurationError(f"{location} must be a list")
    return value


def _positive_int(value: Any, location: str, *, allow_zero: bool = False) -> int:
    if isinstance(value, bool):
        raise ConfigurationError(f"{location} must be an integer")
    try:
        parsed = int(value)
    except (TypeError, ValueError) as exc:
        raise ConfigurationError(f"{location} must be an integer") from exc
    minimum = 0 if allow_zero else 1
    if parsed < minimum:
        raise ConfigurationError(f"{location} must be at least {minimum}")
    return parsed


def _optional_float(value: Any, location: str) -> float | None:
    if value is None:
        return None
    try:
        return float(value)
    except (TypeError, ValueError) as exc:
        raise ConfigurationError(f"{location} must be a number") from exc


def _validate_color(value: Any, location: str, default: str) -> str:
    text = str(value if value is not None else default).strip().lstrip("#").upper()
    if len(text) not in {6, 8} or any(character not in "0123456789ABCDEF" for character in text):
        raise ConfigurationError(f"{location} must be a 6- or 8-character hex color")
    return text


@dataclass(frozen=True)
class EmailSearchConfig:
    sender: str = ""
    subject_contains: str = ""
    since_days: int = 14
    unread_only: bool = False


@dataclass(frozen=True)
class AttachmentConfig:
    filename_patterns: tuple[str, ...] = ("*.xlsx", "*.xlsm")


@dataclass(frozen=True)
class EmailConfig:
    account: str
    imap_host: str
    imap_port: int
    mailbox: str
    search: EmailSearchConfig
    attachments: AttachmentConfig


@dataclass(frozen=True)
class ScheduleConfig:
    enabled: bool
    days: tuple[str, ...]
    time: str
    polling_seconds: int


@dataclass(frozen=True)
class OutputConfig:
    downloads_directory: str
    reports_directory: str
    report_filename: str


@dataclass(frozen=True)
class PrinterConfig:
    enabled: bool
    name: str
    paper_size: str
    orientation: str
    copies: int
    fit_to_pages_wide: int
    fit_to_pages_tall: int


@dataclass(frozen=True)
class LoggingConfig:
    level: str
    max_bytes: int
    backup_count: int


@dataclass(frozen=True)
class TrayConfig:
    tooltip: str


@dataclass(frozen=True)
class UpdateConfig:
    enabled: bool
    repository: str
    check_on_startup: bool
    include_prereleases: bool


@dataclass(frozen=True)
class AppConfig:
    version: int
    email: EmailConfig
    schedule: ScheduleConfig
    output: OutputConfig
    printer: PrinterConfig
    logging: LoggingConfig
    tray: TrayConfig
    update: UpdateConfig


@dataclass(frozen=True)
class InputConfig:
    sheet_name: str | None
    # 0 means "locate the header row automatically" (header_row: auto).
    # Exports that begin with title or filter rows shift the header down, and
    # the offset is not stable between runs.
    header_row: int
    data_start_row: int

    @property
    def auto_header(self) -> bool:
        return self.header_row == 0
    stop_at_first_blank_row: bool
    skip_blank_rows: bool


@dataclass(frozen=True)
class ColumnRule:
    source: str
    target: str
    value_type: str = "text"
    required: bool = True
    default: Any = None
    number_format: str | None = None
    width: float | None = None
    include_in_report: bool = True
    fill_down: bool = False


@dataclass(frozen=True)
class FilterRule:
    column: str
    operator: str
    value: Any = None
    values: tuple[Any, ...] = ()
    minimum: Any = None
    maximum: Any = None
    case_sensitive: bool = False


@dataclass(frozen=True)
class FiltersConfig:
    mode: str = "all"
    rules: tuple[FilterRule, ...] = ()


@dataclass(frozen=True)
class AggregationRule:
    source: str | None
    target: str
    operation: str
    number_format: str | None = None
    width: float | None = None


@dataclass(frozen=True)
class GroupingConfig:
    enabled: bool = False
    by: tuple[str, ...] = ()
    aggregations: tuple[AggregationRule, ...] = ()


@dataclass(frozen=True)
class SplitConfig:
    enabled: bool = False
    by: str | None = None
    title_template: str = "{base_title} — {value}"
    filename_suffix: str = "_{value}"
    include_blank: bool = True
    blank_label: str = "No Sub Status"


@dataclass(frozen=True)
class SortRule:
    column: str
    direction: str = "asc"


@dataclass(frozen=True)
class TextStyle:
    font_name: str = "Aptos"
    font_size: float = 10.0
    bold: bool = False
    font_color: str = "111827"
    fill_color: str = "FFFFFF"
    alignment: str = "left"
    row_height: float | None = None
    alternate_fill_color: str | None = None
    border_color: str | None = None


@dataclass(frozen=True)
class PageConfig:
    paper_size: str = "tabloid"
    orientation: str = "landscape"
    fit_to_pages_wide: int = 1
    fit_to_pages_tall: int = 0
    margin_inches: float = 0.25


@dataclass(frozen=True)
class ReportConfig:
    worksheet_name: str
    title: str
    table_start_row: int
    show_generated_at: bool
    generated_at_label: str
    freeze_panes: str
    auto_filter: bool
    page: PageConfig
    title_style: TextStyle
    header_style: TextStyle
    body_style: TextStyle


@dataclass(frozen=True)
class ExtractionConfig:
    version: int
    input: InputConfig
    columns: tuple[ColumnRule, ...]
    filters: FiltersConfig = field(default_factory=FiltersConfig)
    split: SplitConfig = field(default_factory=SplitConfig)
    grouping: GroupingConfig = field(default_factory=GroupingConfig)
    sort: tuple[SortRule, ...] = ()
    report: ReportConfig | None = None


def load_app_config(path: Path) -> AppConfig:
    root = _load_yaml(path)
    version = _positive_int(root.get("version", 1), "version")
    if version != 1:
        raise ConfigurationError(f"Unsupported app config version: {version}")

    email_data = _mapping(root.get("email", {}), "email")
    search_data = _mapping(email_data.get("search", {}), "email.search")
    attachment_data = _mapping(email_data.get("attachments", {}), "email.attachments")
    patterns = tuple(
        str(item).strip()
        for item in _sequence(
            attachment_data.get("filename_patterns", ["*.xlsx", "*.xlsm"]),
            "email.attachments.filename_patterns",
        )
        if str(item).strip()
    )
    if not patterns:
        raise ConfigurationError("email.attachments.filename_patterns cannot be empty")
    email = EmailConfig(
        account=str(email_data.get("account", "")).strip(),
        imap_host=str(email_data.get("imap_host", "imap.gmail.com")).strip(),
        imap_port=_positive_int(email_data.get("imap_port", 993), "email.imap_port"),
        mailbox=str(email_data.get("mailbox", "INBOX")).strip() or "INBOX",
        search=EmailSearchConfig(
            sender=str(search_data.get("from", "")).strip(),
            subject_contains=str(search_data.get("subject_contains", "")).strip(),
            since_days=_positive_int(
                search_data.get("since_days", 14), "email.search.since_days"
            ),
            unread_only=bool(search_data.get("unread_only", False)),
        ),
        attachments=AttachmentConfig(filename_patterns=patterns),
    )
    if not email.imap_host:
        raise ConfigurationError("email.imap_host cannot be empty")

    schedule_data = _mapping(root.get("schedule", {}), "schedule")
    days = tuple(
        str(item).strip().lower()
        for item in _sequence(
            schedule_data.get(
                "days", ["monday", "tuesday", "wednesday", "thursday", "friday"]
            ),
            "schedule.days",
        )
    )
    invalid_days = sorted(set(days) - WEEKDAYS)
    if invalid_days:
        raise ConfigurationError(f"Invalid schedule day(s): {', '.join(invalid_days)}")
    schedule_time = str(schedule_data.get("time", "07:00")).strip()
    try:
        datetime.strptime(schedule_time, "%H:%M")
    except ValueError as exc:
        raise ConfigurationError("schedule.time must use 24-hour HH:MM format") from exc
    schedule = ScheduleConfig(
        enabled=bool(schedule_data.get("enabled", True)),
        days=days,
        time=schedule_time,
        polling_seconds=_positive_int(
            schedule_data.get("polling_seconds", 20), "schedule.polling_seconds"
        ),
    )

    output_data = _mapping(root.get("output", {}), "output")
    output = OutputConfig(
        downloads_directory=str(
            output_data.get("downloads_directory", "downloads")
        ).strip(),
        reports_directory=str(output_data.get("reports_directory", "reports")).strip(),
        report_filename=str(
            output_data.get("report_filename", "SOP_Report_%Y-%m-%d_%H%M%S.xlsx")
        ).strip(),
    )
    if not output.downloads_directory or not output.reports_directory:
        raise ConfigurationError("output directory values cannot be empty")
    if not output.report_filename.lower().endswith(".xlsx"):
        raise ConfigurationError("output.report_filename must end in .xlsx")

    printer_data = _mapping(root.get("printer", {}), "printer")
    paper_size = str(printer_data.get("paper_size", "tabloid")).strip().lower()
    orientation = str(printer_data.get("orientation", "landscape")).strip().lower()
    if paper_size != "tabloid":
        raise ConfigurationError("printer.paper_size currently supports only 'tabloid'")
    if orientation not in {"landscape", "portrait"}:
        raise ConfigurationError("printer.orientation must be landscape or portrait")
    printer = PrinterConfig(
        enabled=bool(printer_data.get("enabled", True)),
        name=str(printer_data.get("name", "")).strip(),
        paper_size=paper_size,
        orientation=orientation,
        copies=_positive_int(printer_data.get("copies", 1), "printer.copies"),
        fit_to_pages_wide=_positive_int(
            printer_data.get("fit_to_pages_wide", 1),
            "printer.fit_to_pages_wide",
        ),
        fit_to_pages_tall=_positive_int(
            printer_data.get("fit_to_pages_tall", 0),
            "printer.fit_to_pages_tall",
            allow_zero=True,
        ),
    )

    logging_data = _mapping(root.get("logging", {}), "logging")
    logging_config = LoggingConfig(
        level=str(logging_data.get("level", "INFO")).strip().upper(),
        max_bytes=_positive_int(
            logging_data.get("max_bytes", 2_097_152), "logging.max_bytes"
        ),
        backup_count=_positive_int(
            logging_data.get("backup_count", 5),
            "logging.backup_count",
            allow_zero=True,
        ),
    )
    if logging_config.level not in {"DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"}:
        raise ConfigurationError("logging.level is invalid")

    tray_data = _mapping(root.get("tray", {}), "tray")
    tray = TrayConfig(tooltip=str(tray_data.get("tooltip", "SOP Reporter")).strip())

    update_data = _mapping(root.get("update", {}), "update")
    repository = str(update_data.get("repository", "matijepekovic/SOP")).strip()
    if repository.count("/") != 1 or repository.startswith("/") or repository.endswith("/"):
        raise ConfigurationError(
            "update.repository must look like owner/name, for example matijepekovic/SOP"
        )
    update = UpdateConfig(
        enabled=bool(update_data.get("enabled", True)),
        repository=repository,
        check_on_startup=bool(update_data.get("check_on_startup", True)),
        include_prereleases=bool(update_data.get("include_prereleases", False)),
    )

    return AppConfig(
        version=version,
        email=email,
        schedule=schedule,
        output=output,
        printer=printer,
        logging=logging_config,
        tray=tray,
        update=update,
    )


def _parse_text_style(
    data: Mapping[str, Any], location: str, defaults: TextStyle
) -> TextStyle:
    alignment = str(data.get("alignment", defaults.alignment)).strip().lower()
    if alignment not in {"left", "center", "right"}:
        raise ConfigurationError(f"{location}.alignment must be left, center, or right")
    alternate = data.get("alternate_fill_color", defaults.alternate_fill_color)
    border = data.get("border_color", defaults.border_color)
    return TextStyle(
        font_name=str(data.get("font_name", defaults.font_name)).strip(),
        font_size=float(data.get("font_size", defaults.font_size)),
        bold=bool(data.get("bold", defaults.bold)),
        font_color=_validate_color(
            data.get("font_color"), f"{location}.font_color", defaults.font_color
        ),
        fill_color=_validate_color(
            data.get("fill_color"), f"{location}.fill_color", defaults.fill_color
        ),
        alignment=alignment,
        row_height=_optional_float(
            data.get("row_height", defaults.row_height), f"{location}.row_height"
        ),
        alternate_fill_color=(
            _validate_color(alternate, f"{location}.alternate_fill_color", "FFFFFF")
            if alternate is not None
            else None
        ),
        border_color=(
            _validate_color(border, f"{location}.border_color", "B8C4CE")
            if border is not None
            else None
        ),
    )


def load_extraction_config(path: Path) -> ExtractionConfig:
    return extraction_config_from_mapping(_load_yaml(path))


def extraction_config_from_mapping(root: Mapping[str, Any]) -> ExtractionConfig:
    version = _positive_int(root.get("version", 1), "version")
    if version != 1:
        raise ConfigurationError(f"Unsupported extraction config version: {version}")

    input_data = _mapping(root.get("input", {}), "input")
    raw_header_row = input_data.get("header_row", 1)
    if isinstance(raw_header_row, str) and raw_header_row.strip().casefold() == "auto":
        header_row = 0
    else:
        header_row = _positive_int(raw_header_row, "input.header_row")

    raw_data_start = input_data.get("data_start_row")
    if header_row == 0:
        # Data begins on the row after whichever row is detected, so an
        # explicit data_start_row would contradict the detection.
        if raw_data_start is not None:
            raise ConfigurationError(
                "input.data_start_row cannot be set when input.header_row is auto; "
                "data starts on the row after the detected header"
            )
        data_start_row = 0
    else:
        data_start_row = _positive_int(
            raw_data_start if raw_data_start is not None else header_row + 1,
            "input.data_start_row",
        )
        if data_start_row <= header_row:
            raise ConfigurationError(
                "input.data_start_row must be after input.header_row"
            )
    raw_sheet_name = input_data.get("sheet_name")
    input_config = InputConfig(
        sheet_name=str(raw_sheet_name).strip() if raw_sheet_name not in {None, ""} else None,
        header_row=header_row,
        data_start_row=data_start_row,
        stop_at_first_blank_row=bool(input_data.get("stop_at_first_blank_row", False)),
        skip_blank_rows=bool(input_data.get("skip_blank_rows", True)),
    )

    column_items = _sequence(root.get("columns", []), "columns")
    columns: list[ColumnRule] = []
    for index, raw_item in enumerate(column_items):
        item = _mapping(raw_item, f"columns[{index}]")
        source = str(item.get("source", "")).strip()
        target = str(item.get("target", source)).strip()
        value_type = str(item.get("type", "text")).strip().lower()
        if not source or not target:
            raise ConfigurationError(f"columns[{index}] needs source and target")
        if value_type not in SUPPORTED_TYPES:
            raise ConfigurationError(
                f"columns[{index}].type must be one of {sorted(SUPPORTED_TYPES)}"
            )
        columns.append(
            ColumnRule(
                source=source,
                target=target,
                value_type=value_type,
                required=bool(item.get("required", True)),
                default=item.get("default"),
                number_format=(
                    str(item["number_format"])
                    if item.get("number_format") is not None
                    else None
                ),
                width=_optional_float(item.get("width"), f"columns[{index}].width"),
                include_in_report=bool(item.get("include_in_report", True)),
                fill_down=bool(item.get("fill_down", False)),
            )
        )
    if not columns:
        raise ConfigurationError("At least one columns entry is required")
    targets = [column.target.casefold() for column in columns]
    if len(targets) != len(set(targets)):
        raise ConfigurationError("Column target names must be unique")

    filters_data = _mapping(root.get("filters", {}), "filters")
    filter_mode = str(filters_data.get("mode", "all")).strip().lower()
    if filter_mode not in {"all", "any"}:
        raise ConfigurationError("filters.mode must be all or any")
    filter_rules: list[FilterRule] = []
    for index, raw_item in enumerate(
        _sequence(filters_data.get("rules", []), "filters.rules")
    ):
        item = _mapping(raw_item, f"filters.rules[{index}]")
        column = str(item.get("column", "")).strip()
        operator = str(item.get("operator", item.get("op", ""))).strip().lower()
        if not column or operator not in SUPPORTED_FILTERS:
            raise ConfigurationError(
                f"filters.rules[{index}] needs a column and a supported operator"
            )
        values_raw = item.get("values", item.get("value", []))
        if operator in {"in", "not_in"}:
            if isinstance(values_raw, (str, bytes)) or not isinstance(values_raw, Sequence):
                raise ConfigurationError(f"filters.rules[{index}].values must be a list")
            values = tuple(values_raw)
        else:
            values = ()
        filter_rules.append(
            FilterRule(
                column=column,
                operator=operator,
                value=item.get("value"),
                values=values,
                minimum=item.get("min", item.get("minimum")),
                maximum=item.get("max", item.get("maximum")),
                case_sensitive=bool(item.get("case_sensitive", False)),
            )
        )
    filters = FiltersConfig(mode=filter_mode, rules=tuple(filter_rules))

    column_targets = {column.target for column in columns}
    split_data = _mapping(root.get("split", {}), "split")
    split_enabled = bool(split_data.get("enabled", False))
    raw_split_by = split_data.get("by")
    split_by = (
        str(raw_split_by).strip() if raw_split_by not in {None, ""} else None
    )
    title_template = str(
        split_data.get("title_template", "{base_title} — {value}")
    ).strip()
    filename_suffix = str(
        split_data.get("filename_suffix", "_{value}")
    ).strip()
    blank_label = str(split_data.get("blank_label", "No Sub Status")).strip()
    if split_enabled:
        if split_by is None:
            raise ConfigurationError("Enabled split requires split.by")
        if split_by not in column_targets:
            raise ConfigurationError(
                f"split.by is not a mapped target column: {split_by}"
            )
        if not title_template or not filename_suffix:
            raise ConfigurationError(
                "Enabled split requires title_template and filename_suffix"
            )
        if not blank_label:
            raise ConfigurationError("split.blank_label cannot be empty")
        try:
            title_template.format(base_title="SOP REPORT", value="Status")
            filename_suffix.format(base_title="SOP_REPORT", value="Status")
        except (IndexError, KeyError, ValueError) as exc:
            raise ConfigurationError(
                "split templates may use only {base_title} and {value}"
            ) from exc
    split = SplitConfig(
        enabled=split_enabled,
        by=split_by,
        title_template=title_template,
        filename_suffix=filename_suffix,
        include_blank=bool(split_data.get("include_blank", True)),
        blank_label=blank_label,
    )

    grouping_data = _mapping(root.get("grouping", {}), "grouping")
    group_by = tuple(
        str(item).strip()
        for item in _sequence(grouping_data.get("by", []), "grouping.by")
        if str(item).strip()
    )
    aggregations: list[AggregationRule] = []
    for index, raw_item in enumerate(
        _sequence(grouping_data.get("aggregations", []), "grouping.aggregations")
    ):
        item = _mapping(raw_item, f"grouping.aggregations[{index}]")
        operation = str(item.get("operation", "")).strip().lower()
        source_value = item.get("source", item.get("column"))
        source = str(source_value).strip() if source_value not in {None, ""} else None
        target = str(item.get("target", "")).strip()
        if operation not in SUPPORTED_AGGREGATIONS or not target:
            raise ConfigurationError(
                f"grouping.aggregations[{index}] needs a target and supported operation"
            )
        if operation != "count_rows" and source is None:
            raise ConfigurationError(
                f"grouping.aggregations[{index}].source is required for {operation}"
            )
        aggregations.append(
            AggregationRule(
                source=source,
                target=target,
                operation=operation,
                number_format=(
                    str(item["number_format"])
                    if item.get("number_format") is not None
                    else None
                ),
                width=_optional_float(
                    item.get("width"), f"grouping.aggregations[{index}].width"
                ),
            )
        )
    grouping = GroupingConfig(
        enabled=bool(grouping_data.get("enabled", False)),
        by=group_by,
        aggregations=tuple(aggregations),
    )
    if grouping.enabled:
        missing_group_columns = [name for name in grouping.by if name not in column_targets]
        if missing_group_columns:
            raise ConfigurationError(
                "grouping.by contains unknown target column(s): "
                + ", ".join(missing_group_columns)
            )
        if not grouping.by or not grouping.aggregations:
            raise ConfigurationError(
                "Enabled grouping requires at least one by column and aggregation"
            )
        for aggregation in grouping.aggregations:
            if aggregation.source is not None and aggregation.source not in column_targets:
                raise ConfigurationError(
                    f"Aggregation source is not a mapped target column: {aggregation.source}"
                )

    sort_rules: list[SortRule] = []
    for index, raw_item in enumerate(_sequence(root.get("sort", []), "sort")):
        item = _mapping(raw_item, f"sort[{index}]")
        column = str(item.get("column", "")).strip()
        direction = str(item.get("direction", "asc")).strip().lower()
        if not column or direction not in {"asc", "desc"}:
            raise ConfigurationError(
                f"sort[{index}] needs a column and asc/desc direction"
            )
        sort_rules.append(SortRule(column=column, direction=direction))

    report_data = _mapping(root.get("report", {}), "report")
    page_data = _mapping(report_data.get("page", {}), "report.page")
    paper_size = str(page_data.get("paper_size", "tabloid")).strip().lower()
    orientation = str(page_data.get("orientation", "landscape")).strip().lower()
    if paper_size != "tabloid":
        raise ConfigurationError("report.page.paper_size currently supports only tabloid")
    if orientation not in {"landscape", "portrait"}:
        raise ConfigurationError("report.page.orientation must be landscape or portrait")
    page = PageConfig(
        paper_size=paper_size,
        orientation=orientation,
        fit_to_pages_wide=_positive_int(
            page_data.get("fit_to_pages_wide", 1),
            "report.page.fit_to_pages_wide",
        ),
        fit_to_pages_tall=_positive_int(
            page_data.get("fit_to_pages_tall", 0),
            "report.page.fit_to_pages_tall",
            allow_zero=True,
        ),
        margin_inches=float(page_data.get("margin_inches", 0.25)),
    )
    if page.margin_inches < 0:
        raise ConfigurationError("report.page.margin_inches cannot be negative")

    styles_data = _mapping(report_data.get("styles", {}), "report.styles")
    title_style = _parse_text_style(
        _mapping(styles_data.get("title", {}), "report.styles.title"),
        "report.styles.title",
        TextStyle(
            font_name="Aptos Display",
            font_size=20,
            bold=True,
            font_color="FFFFFF",
            fill_color="172033",
        ),
    )
    header_style = _parse_text_style(
        _mapping(styles_data.get("header", {}), "report.styles.header"),
        "report.styles.header",
        TextStyle(
            font_name="Aptos",
            font_size=11,
            bold=True,
            font_color="FFFFFF",
            fill_color="1F4E78",
            alignment="center",
            row_height=24,
        ),
    )
    body_style = _parse_text_style(
        _mapping(styles_data.get("body", {}), "report.styles.body"),
        "report.styles.body",
        TextStyle(
            alternate_fill_color="EAF2F8",
            border_color="B8C4CE",
            row_height=20,
        ),
    )
    table_start_row = _positive_int(
        report_data.get("table_start_row", 4), "report.table_start_row"
    )
    report = ReportConfig(
        worksheet_name=str(report_data.get("worksheet_name", "Report")).strip()[:31]
        or "Report",
        title=str(report_data.get("title", "SOP REPORT")).strip(),
        table_start_row=table_start_row,
        show_generated_at=bool(report_data.get("show_generated_at", True)),
        generated_at_label=str(
            report_data.get("generated_at_label", "Generated")
        ).strip(),
        freeze_panes=str(report_data.get("freeze_panes", "auto")).strip(),
        auto_filter=bool(report_data.get("auto_filter", True)),
        page=page,
        title_style=title_style,
        header_style=header_style,
        body_style=body_style,
    )

    output_headers = (
        list(grouping.by) + [item.target for item in grouping.aggregations]
        if grouping.enabled
        else [column.target for column in columns if column.include_in_report]
    )
    if not output_headers:
        raise ConfigurationError("At least one column must be included in the report")
    for sort_rule in sort_rules:
        if sort_rule.column not in output_headers:
            raise ConfigurationError(f"Sort column is not in report output: {sort_rule.column}")

    return ExtractionConfig(
        version=version,
        input=input_config,
        columns=tuple(columns),
        filters=filters,
        split=split,
        grouping=grouping,
        sort=tuple(sort_rules),
        report=report,
    )


RULES_DIRECTORY_NAME = "rules"
LEGACY_RULES_FILENAME = "extraction_rules.yaml"
DEFAULT_ATTACHMENT_PATTERNS = ("*.xlsx", "*.xlsm")


@dataclass(frozen=True)
class MatchConfig:
    """Which downloaded attachments a report definition claims."""

    filename_patterns: tuple[str, ...] = DEFAULT_ATTACHMENT_PATTERNS
    enabled: bool = True

    def matches(self, filename: str) -> bool:
        if not self.enabled:
            return False
        name = Path(filename).name.casefold()
        return any(
            fnmatch.fnmatch(name, pattern.casefold())
            for pattern in self.filename_patterns
        )


@dataclass(frozen=True)
class ReportDefinition:
    """One emailed report: how to recognize it and how to turn it into output."""

    name: str
    match: MatchConfig
    extraction: ExtractionConfig
    report_filename: str | None = None
    source: Path | None = None


def _report_definition_from_mapping(
    root: Mapping[str, Any], *, name: str, source: Path | None
) -> ReportDefinition:
    match_data = _mapping(root.get("match", {}), "match")
    patterns = tuple(
        str(item).strip()
        for item in _sequence(
            match_data.get("filename_patterns", DEFAULT_ATTACHMENT_PATTERNS),
            "match.filename_patterns",
        )
        if str(item).strip()
    )
    if not patterns:
        raise ConfigurationError("match.filename_patterns cannot be empty")

    output_data = _mapping(root.get("output", {}), "output")
    report_filename = str(output_data.get("report_filename", "")).strip() or None

    return ReportDefinition(
        name=str(root.get("name", "")).strip() or name,
        match=MatchConfig(
            filename_patterns=patterns,
            enabled=bool(match_data.get("enabled", True)),
        ),
        extraction=extraction_config_from_mapping(root),
        report_filename=report_filename,
        source=source,
    )


def load_report_definitions(config_dir: Path) -> tuple[ReportDefinition, ...]:
    """Load every report definition from ``config_dir``.

    Each ``rules/*.yaml`` file is one report. When that directory is absent or
    empty the single legacy ``extraction_rules.yaml`` is used instead, so
    installations created before per-report rules existed keep working.
    """
    config_dir = Path(config_dir)
    rules_dir = config_dir / RULES_DIRECTORY_NAME
    definitions: list[ReportDefinition] = []

    if rules_dir.is_dir():
        # Sorted so routing order — and therefore which definition wins an
        # overlapping pattern — is deterministic rather than filesystem order.
        for path in sorted(rules_dir.glob("*.yaml")):
            try:
                definitions.append(
                    _report_definition_from_mapping(
                        _load_yaml(path), name=path.stem, source=path
                    )
                )
            except ConfigurationError as exc:
                raise ConfigurationError(f"{path.name}: {exc}") from exc

    if definitions:
        _reject_duplicate_names(definitions)
        return tuple(definitions)

    legacy = config_dir / LEGACY_RULES_FILENAME
    if not legacy.is_file():
        raise ConfigurationError(
            f"No report rules found. Expected YAML files in {rules_dir} "
            f"or a {LEGACY_RULES_FILENAME} in {config_dir}."
        )
    return (
        ReportDefinition(
            name=legacy.stem,
            match=MatchConfig(),
            extraction=load_extraction_config(legacy),
            source=legacy,
        ),
    )


def _reject_duplicate_names(definitions: Sequence[ReportDefinition]) -> None:
    seen: set[str] = set()
    for definition in definitions:
        folded = definition.name.casefold()
        if folded in seen:
            raise ConfigurationError(
                f"Two report rule files share the name {definition.name!r}; "
                "names identify reports in logs and must be unique"
            )
        seen.add(folded)


def update_email_account(path: Path, account: str) -> None:
    root = _load_yaml(path)
    email_data = root.setdefault("email", {})
    if not isinstance(email_data, dict):
        raise ConfigurationError("email must be a mapping")
    email_data["account"] = account.strip()
    temporary = path.with_suffix(path.suffix + ".tmp")
    try:
        with temporary.open("w", encoding="utf-8", newline="\n") as handle:
            yaml.safe_dump(root, handle, sort_keys=False, allow_unicode=True)
        temporary.replace(path)
    except OSError as exc:
        raise ConfigurationError(f"Cannot update Gmail account in {path}: {exc}") from exc
