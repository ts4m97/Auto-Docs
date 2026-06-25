from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
import re
import unicodedata


@dataclass
class SqlConnectionConfig:
    server: str
    database: str
    username: str
    password: str


@dataclass
class SqlImportResult:
    imported: int
    skipped: int
    cccd_column: str
    course_join: str


CCCD_COLUMNS = [
    "cccd",
    "so_cccd",
    "socccd",
    "cmnd",
    "so_cmnd",
    "socmnd",
    "cmt",
    "so_cmt",
    "socmt",
    "so_giay_to",
    "sogiayto",
    "so_giayto",
]

COURSE_KEY_COLUMNS = [
    "ma_khoa_hoc",
    "makhoahoc",
    "khoa_hoc_id",
    "khoahocid",
    "id_khoa_hoc",
    "idkhoahoc",
    "ma_kh",
    "makh",
]


def connect_sql_server(config: SqlConnectionConfig):
    try:
        import pyodbc
    except ImportError as exc:
        raise RuntimeError("Chua cai thu vien pyodbc. Hay chay: python -m pip install pyodbc") from exc

    drivers = pyodbc.drivers()
    preferred = [
        "ODBC Driver 18 for SQL Server",
        "ODBC Driver 17 for SQL Server",
        "SQL Server Native Client 11.0",
        "SQL Server",
    ]
    driver = next((item for item in preferred if item in drivers), None)
    if not driver:
        raise RuntimeError("Khong tim thay ODBC Driver cho SQL Server tren may nay.")

    connection_string = (
        f"DRIVER={{{driver}}};"
        f"SERVER={config.server};"
        f"DATABASE={config.database};"
        f"UID={config.username};"
        f"PWD={config.password};"
        "Encrypt=no;"
        "TrustServerCertificate=yes;"
    )
    return pyodbc.connect(connection_string, timeout=8)


def fetch_table_columns(connection, table_name: str) -> list[str]:
    schema, name = table_name.split(".", 1)
    rows = connection.cursor().execute(
        """
        SELECT COLUMN_NAME
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = ? AND TABLE_NAME = ?
        ORDER BY ORDINAL_POSITION
        """,
        schema,
        name,
    ).fetchall()
    return [row[0] for row in rows]


def fetch_dvhc_lookup(connection) -> dict[tuple[str, str], str]:
    if not fetch_table_columns(connection, "dbo.DM_DVHC"):
        return {}
    rows = connection.cursor().execute(
        """
        SELECT MaDvhc, MaDVQL, TenDayDu, TenDvhc, TenNganGon
        FROM dbo.DM_DVHC
        """
    ).fetchall()
    lookup: dict[tuple[str, str], str] = {}
    by_code: dict[str, list[str]] = {}
    for ma_dvhc, ma_dvql, ten_day_du, ten_dvhc, ten_ngan_gon in rows:
        code = format_sql_value(ma_dvhc)
        parent_code = format_sql_value(ma_dvql)
        name = first_non_empty(ten_day_du, ten_dvhc, ten_ngan_gon)
        if not code or not name:
            continue
        lookup[(code, parent_code)] = name
        by_code.setdefault(code, []).append(name)

    for code, names in by_code.items():
        unique_names = sorted(set(names))
        if len(unique_names) == 1:
            lookup[(code, "")] = unique_names[0]
    return lookup


def import_drivers_from_sql(config: SqlConnectionConfig, store) -> SqlImportResult:
    with connect_sql_server(config) as connection:
        driver_columns = fetch_table_columns(connection, "dbo.NguoiLX")
        profile_columns = fetch_table_columns(connection, "dbo.NguoiLX_HoSo")
        course_columns = fetch_table_columns(connection, "dbo.KhoaHoc")
        dvhc_lookup = fetch_dvhc_lookup(connection)
        if not driver_columns:
            raise RuntimeError("Khong tim thay bang dbo.NguoiLX hoac bang khong co cot.")

        cccd_column = choose_column(driver_columns, CCCD_COLUMNS)
        if not cccd_column:
            raise RuntimeError("Khong tim thay cot CCCD/CMND trong dbo.NguoiLX.")

        has_profile_join = can_join_profile(driver_columns, profile_columns)
        has_course_join = can_join_course(profile_columns, course_columns)
        query = build_import_query(driver_columns, profile_columns, course_columns, has_profile_join, has_course_join)
        cursor = connection.cursor()
        cursor.execute(query)
        columns = [column[0] for column in cursor.description]
        rows = cursor.fetchall()

    imported = 0
    skipped = 0
    cccd_alias = f"nguoilx__{cccd_column}"
    for row in rows:
        values = row_to_fields(columns, row)
        cccd = values.get(cccd_alias, "").strip()
        if not cccd:
            skipped += 1
            continue

        fields = normalize_imported_fields(values)
        enrich_administrative_names(fields, dvhc_lookup)
        fields["cccd"] = cccd
        store.upsert_customer(cccd, fields)
        imported += 1

    join_label = (
        "NguoiLX.MaDK = NguoiLX_HoSo.MaDK; NguoiLX_HoSo.MaKhoaHoc = KhoaHoc.MaKH"
        if has_profile_join and has_course_join
        else "khong ghep KhoaHoc"
    )
    return SqlImportResult(
        imported=imported,
        skipped=skipped,
        cccd_column=cccd_column,
        course_join=join_label,
    )


def choose_column(columns: list[str], candidates: list[str]) -> str | None:
    normalized = {normalize_name(column): column for column in columns}
    for candidate in candidates:
        value = normalized.get(normalize_name(candidate))
        if value:
            return value
    for column in columns:
        key = normalize_name(column)
        if any(part in key for part in ("cccd", "cmnd", "cmt")):
            return column
    return None


def choose_course_join(driver_columns: list[str], course_columns: list[str]) -> tuple[str | None, str | None]:
    if not course_columns:
        return None, None

    normalized_driver = {normalize_name(column): column for column in driver_columns}
    normalized_course = {normalize_name(column): column for column in course_columns}
    for candidate in COURSE_KEY_COLUMNS:
        key = normalize_name(candidate)
        if key in normalized_driver and key in normalized_course:
            return normalized_driver[key], normalized_course[key]

    for key, driver_column in normalized_driver.items():
        course_column = normalized_course.get(key)
        if course_column and "khoa" in key and ("hoc" in key or key.endswith("kh")):
            return driver_column, course_column
    return None, None


def can_join_profile(driver_columns: list[str], profile_columns: list[str]) -> bool:
    return has_column(driver_columns, "MaDK") and has_column(profile_columns, "MaDK")


def can_join_course(profile_columns: list[str], course_columns: list[str]) -> bool:
    return has_column(profile_columns, "MaKhoaHoc") and has_column(course_columns, "MaKH")


def has_column(columns: list[str], column_name: str) -> bool:
    wanted = normalize_name(column_name)
    return any(normalize_name(column) == wanted for column in columns)


def build_import_query(
    driver_columns: list[str],
    profile_columns: list[str],
    course_columns: list[str],
    has_profile_join: bool,
    has_course_join: bool,
) -> str:
    driver_select = [f"n.{quote_sql(column)} AS {quote_sql('nguoilx__' + column)}" for column in driver_columns]
    profile_select = []
    if has_profile_join:
        profile_select = [f"h.{quote_sql(column)} AS {quote_sql('hoso__' + column)}" for column in profile_columns]
    course_select = []
    if has_profile_join and has_course_join:
        course_select = [f"k.{quote_sql(column)} AS {quote_sql('khoahoc__' + column)}" for column in course_columns]
    select_parts = driver_select + profile_select + course_select
    query = f"SELECT {', '.join(select_parts)} FROM dbo.NguoiLX n"
    if has_profile_join:
        query += " LEFT JOIN dbo.NguoiLX_HoSo h ON n.[MaDK] = h.[MaDK]"
    if has_profile_join and has_course_join:
        query += " LEFT JOIN dbo.KhoaHoc k ON h.[MaKhoaHoc] = k.[MaKH]"
    return query


def row_to_fields(columns: list[str], row) -> dict[str, str]:
    return {column: format_sql_value(value, column) for column, value in zip(columns, row)}


def normalize_imported_fields(values: dict[str, str]) -> dict[str, str]:
    fields: dict[str, str] = {}
    for key, value in values.items():
        if value == "":
            continue
        source, _, raw_name = key.partition("__")
        prefix = source_prefix(source)
        normalized = normalize_output_key(raw_name)
        fields[normalized] = value
        fields[f"{prefix}_{normalized}"] = value
    apply_common_aliases(fields)
    return fields


def source_prefix(source: str) -> str:
    if source == "khoahoc":
        return "khoa_hoc"
    if source == "hoso":
        return "ho_so"
    return "nguoi_lx"


def apply_common_aliases(fields: dict[str, str]) -> None:
    aliases = {
        "ten_khach_hang": ["ho_ten", "hoten", "ten", "ho_va_ten", "hovaten"],
        "so_dien_thoai": ["dien_thoai", "dienthoai", "sdt", "so_dt", "sodt"],
        "dia_chi": ["dia_chi", "diachi", "dia_chi_thuong_tru", "diachithuongtru"],
        "ngay_sinh": ["ngay_sinh", "ngaysinh"],
        "ma_khoa_hoc": ["ma_khoa_hoc", "makhoahoc", "makh", "khoa_hoc_id", "khoahocid"],
        "ten_khoa_hoc": ["ten_khoa_hoc", "tenkhoahoc", "ten_khoa", "tenkhoa", "tenkh"],
        "hang_dao_tao": ["hang_dao_tao", "hangdaotao", "hang_dt", "hangdt", "hang_gplx", "hanggplx"],
    }
    by_normalized = {normalize_name(key): value for key, value in fields.items()}
    for canonical, candidates in aliases.items():
        if fields.get(canonical):
            continue
        for candidate in candidates:
            value = by_normalized.get(normalize_name(candidate))
            if value:
                fields[canonical] = value
                break


def enrich_administrative_names(fields: dict[str, str], dvhc_lookup: dict[tuple[str, str], str]) -> None:
    residence = resolve_dvhc_name(fields.get("noitt_madvhc", ""), fields.get("noitt_madvql", ""), dvhc_lookup)
    if residence:
        set_dvhc_display_fields(fields, "noitt", residence)
        fields["noi_thuong_tru"] = residence
        fields["noi_dang_ky_thuong_tru"] = residence
        fields["dia_chi_thuong_tru"] = residence

    current_place = resolve_dvhc_name(fields.get("noict_madvhc", ""), fields.get("noict_madvql", ""), dvhc_lookup)
    if current_place:
        set_dvhc_display_fields(fields, "noict", current_place)
        fields["noi_cu_tru"] = current_place
        fields["noi_o_hien_tai"] = current_place


def set_dvhc_display_fields(fields: dict[str, str], prefix: str, display_name: str) -> None:
    code_key = f"{prefix}_madvhc"
    source_code_key = f"nguoi_lx_{prefix}_madvhc"
    if code_key in fields:
        fields[f"{prefix}_ma_dvhc_goc"] = fields[code_key]
        fields[code_key] = display_name
    if source_code_key in fields:
        fields[f"nguoi_lx_{prefix}_ma_dvhc_goc"] = fields[source_code_key]
        fields[source_code_key] = display_name
    fields[prefix] = display_name
    fields[f"{prefix}_ten_dvhc"] = display_name


def resolve_dvhc_name(code: str, parent_code: str, dvhc_lookup: dict[tuple[str, str], str]) -> str:
    code = str(code or "").strip()
    parent_code = str(parent_code or "").strip()
    if not code:
        return ""
    return dvhc_lookup.get((code, parent_code)) or dvhc_lookup.get((code, "")) or code


def format_sql_value(value, column_name: str = "") -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y") if is_date_only_column(column_name) else value.strftime("%d/%m/%Y %H:%M:%S")
    if isinstance(value, date):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, Decimal):
        value = format(value, "f").rstrip("0").rstrip(".")

    text = str(value).strip()
    parsed_date = parse_temporal_text(text)
    if parsed_date and is_date_like_column(column_name):
        if isinstance(parsed_date, datetime) and not is_date_only_column(column_name):
            return parsed_date.strftime("%d/%m/%Y %H:%M:%S")
        return parsed_date.strftime("%d/%m/%Y")
    return text


def first_non_empty(*values) -> str:
    for value in values:
        text = format_sql_value(value)
        if text:
            return text
    return ""


def is_date_like_column(column_name: str) -> bool:
    key = normalize_name(column_name)
    return any(part in key for part in ("ngay", "date", "sinh"))


def is_date_only_column(column_name: str) -> bool:
    key = normalize_name(column_name)
    return any(part in key for part in ("sinh", "ngaycap", "ngaythi", "ngayhethan"))


def parse_date_text(value: str) -> date | None:
    parsed = parse_temporal_text(value)
    if isinstance(parsed, datetime):
        return parsed.date()
    return parsed


def parse_temporal_text(value: str) -> date | datetime | None:
    if not value:
        return None

    compact = re.fullmatch(r"(\d{4})(\d{2})(\d{2})", value)
    if compact:
        return safe_date(int(compact.group(1)), int(compact.group(2)), int(compact.group(3)))

    normalized = value.replace("T", " ")
    for pattern in ("%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            parsed = datetime.strptime(normalized[:19], pattern)
            return parsed if "H" in pattern else parsed.date()
        except ValueError:
            continue
    return None


def safe_date(year: int, month: int, day: int) -> date | None:
    try:
        return date(year, month, day)
    except ValueError:
        return None


def normalize_output_key(value: str) -> str:
    value = remove_accents(value)
    key = []
    last_was_sep = False
    for char in value.strip():
        if char.isalnum():
            key.append(char.lower())
            last_was_sep = False
        elif not last_was_sep:
            key.append("_")
            last_was_sep = True
    return "".join(key).strip("_")


def normalize_name(value: str) -> str:
    value = remove_accents(value)
    return "".join(char.casefold() for char in value if char.isalnum())


def remove_accents(value: str) -> str:
    text = unicodedata.normalize("NFKD", value)
    return "".join(char for char in text if not unicodedata.combining(char))


def quote_sql(identifier: str) -> str:
    return "[" + identifier.replace("]", "]]") + "]"
