from __future__ import annotations

import argparse
import math
from pathlib import Path
import re
import time
import unicodedata
import warnings

import pandas as pd

from check_atom import build_atom_check
from field_profile import apply_field_profile, load_field_profile
from non_pickup_validation import (
    non_pickup_atoms_in_record_scope,
    non_pickup_candidate_validation_reason,
)


PROJECT_ROOT = Path(__file__).resolve().parent
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "data" / "output"
DEFAULT_MODEL_COMBO_PATH = PROJECT_ROOT / "database" / "model_combo.tsv"
BACKSIZE_SOURCE_COLUMN = "最终尺码"
REQUIRED_COLUMNS = [
    "品牌",
    "前台车型",
    "结构",
    "版本",
    "年份区间",
    BACKSIZE_SOURCE_COLUMN,
]
NON_PICKUP_REQUIRED_COLUMNS = REQUIRED_COLUMNS
PICKUP_REQUIRED_COLUMNS = [
    "品牌",
    "前台车型",
    "版本",
    "年份区间",
    BACKSIZE_SOURCE_COLUMN,
    "驾驶室类型",
    "货斗长度_ft",
]
NON_PICKUP_FINAL_COLUMNS = [
    "主车型",
    "BRAND",
    "MODEL",
    "YEAR",
    "VERSION",
    "CONST",
    "BACKSIZE",
]
PICKUP_FINAL_COLUMNS = [
    "主车型",
    "BRAND",
    "MODEL",
    "YEAR",
    "VERSION",
    "CAB",
    "BED_FT",
    "BACKSIZE",
]
NON_PICKUP_EXPORT_COLUMNS = [
    "CAR",
    "MAKE",
    "MODEL",
    "YEAR",
    "VERSION",
    "CONST",
    "BACKSIZE",
]
PICKUP_EXPORT_COLUMNS = [
    "CAR",
    "MAKE",
    "MODEL",
    "YEAR",
    "VERSION",
    "CAB",
    "BED",
    "BACKSIZE",
]
ATOM_FINAL_COLUMNS = [
    "压缩类型",
    "BRAND",
    "MODEL",
    "YEAR",
    "VERSION",
    "CONST",
    "CAB",
    "BED_FT",
    "BACKSIZE",
]
PROCESS_FINAL_COLUMNS = [
    "压缩类型",
    "BRAND",
    "MODEL",
    "原始行数",
    "原子事实数",
    "无损年份行数",
    "无损年份结构行数",
    "特定性行数",
    "相邻合并尝试次数",
    "相邻合并成功次数",
    "相邻合并Fallback次数",
    "皮卡年份闭合尝试组数",
    "皮卡年份闭合成功组数",
    "皮卡年份闭合Fallback组数",
    "皮卡Bed合并尝试组数",
    "皮卡Bed合并成功组数",
    "皮卡Bed合并Fallback组数",
    "风险提示",
]
LOG_FINAL_COLUMNS = [
    "压缩类型",
    "阶段",
    "BRAND",
    "MODEL",
    "合并MODEL",
    "CAB",
    "BED_FT",
    "BACKSIZE",
    "候选YEAR",
    "候选CONST",
    "候选VERSION",
    "合并CAB",
    "合并BED_FT",
    "合并YEAR",
    "合并CONST",
    "合并VERSION",
    "结果",
    "原因",
]
LEGACY_FRONT_MODEL_COLUMNS = ["车型名", "车姓名"]


class ProgressReporter:
    def __init__(self, interval_seconds: float = 10.0, enabled: bool = True) -> None:
        self.interval_seconds = interval_seconds
        self.enabled = enabled
        self.phase = ""
        self.total_models = 0
        self.completed_models = 0
        self.unit_label = "MODEL"
        self.current_make = ""
        self.current_model = ""
        self.merge_count = 0
        self.attempt_count = 0
        self.started_at = time.monotonic()
        self.last_emit_at = 0.0

    def start(self, phase: str, total_models: int, unit_label: str = "MODEL") -> None:
        self.phase = phase
        self.total_models = total_models
        self.completed_models = 0
        self.unit_label = unit_label
        self.current_make = ""
        self.current_model = ""
        self.merge_count = 0
        self.attempt_count = 0
        self.started_at = time.monotonic()
        self.last_emit_at = 0.0
        self.emit(force=True)

    def update(
        self,
        current_make: object = "",
        current_model: object = "",
        completed_models: int | None = None,
        merge_count: int | None = None,
        attempt_count: int | None = None,
        force: bool = False,
    ) -> None:
        if current_make != "":
            self.current_make = normalize_text(current_make)
        if current_model != "":
            self.current_model = normalize_text(current_model)
        if completed_models is not None:
            self.completed_models = completed_models
        if merge_count is not None:
            self.merge_count = merge_count
        if attempt_count is not None:
            self.attempt_count = attempt_count
        self.emit(force=force)

    def finish(self) -> None:
        self.completed_models = self.total_models
        self.emit(force=True)

    def emit(self, force: bool = False) -> None:
        if not self.enabled:
            return
        now = time.monotonic()
        if not force and now - self.last_emit_at < self.interval_seconds:
            return
        self.last_emit_at = now
        elapsed = int(now - self.started_at)
        total_text = str(self.total_models) if self.total_models else "?"
        print(
            "[进度] "
            f"{self.phase} | "
            f"MAKE={self.current_make or '-'} | "
            f"MODEL={self.current_model or '-'} | "
            f"已完成{self.unit_label}={self.completed_models}/{total_text} | "
            f"合并项数量={self.merge_count} | "
            f"尝试次数={self.attempt_count} | "
            f"用时={elapsed}s",
            flush=True,
        )


def read_tsv(path: Path, encoding: str = "utf-8-sig") -> pd.DataFrame:
    """Read a TSV file while preserving text-like values as strings."""
    return pd.read_csv(path, sep="\t", dtype=str, encoding=encoding, keep_default_na=False)


EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xls", ".xlsb"}


_MODEL_COMBO_CACHE: dict[Path, dict[tuple[str, str], str]] = {}


def load_model_combo_map(path: Path = DEFAULT_MODEL_COMBO_PATH) -> dict[tuple[str, str], str]:
    path = path.resolve()
    if path in _MODEL_COMBO_CACHE:
        return _MODEL_COMBO_CACHE[path]
    if not path.exists():
        _MODEL_COMBO_CACHE[path] = {}
        return {}

    combo_df = read_tsv(path)
    combo_df.columns = [normalize_text(column).upper() for column in combo_df.columns]
    if not {"MAKE", "MODELS"}.issubset(combo_df.columns):
        _MODEL_COMBO_CACHE[path] = {}
        return {}

    mapping: dict[tuple[str, str], str] = {}
    for _, row in combo_df.iterrows():
        make = normalize_text(row.get("MAKE", ""))
        models = [normalize_text(item) for item in normalize_text(row.get("MODELS", "")).split("|")]
        models = [model for model in models if model]
        if not make or len(models) < 2:
            continue
        group_key = "|".join(models)
        for model in models:
            mapping[(make.casefold(), model.casefold())] = group_key
    _MODEL_COMBO_CACHE[path] = mapping
    return mapping


def model_combo_group(brand: object, model: object) -> str:
    brand_text = normalize_text(brand)
    model_text = normalize_text(model)
    return load_model_combo_map().get((brand_text.casefold(), model_text.casefold()), model_text)


def normalize_input_schema(df: pd.DataFrame, field_profile: dict[str, object] | None = None) -> pd.DataFrame:
    """Normalize historical column names to the current allcars schema."""
    return apply_field_profile(df, field_profile)


def normalize_text(value: object) -> str:
    if value is None or value is pd.NA or value is pd.NaT:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""

    text = str(value)
    text = text.replace("\u00a0", " ")
    text = text.replace("\u200b", "")
    text = text.replace("\ufeff", "")
    text = "".join(ch for ch in text if unicodedata.category(ch)[0] != "C")
    return text.strip()


def profile_input_config(field_profile: dict[str, object] | None) -> dict[str, object]:
    profile = field_profile or {}
    configured = profile.get("input", {})
    if configured is None:
        return {}
    if not isinstance(configured, dict):
        raise ValueError("Field profile 'input' must be a mapping.")
    return configured


def profile_input_path(field_profile: dict[str, object] | None) -> str:
    config = profile_input_config(field_profile)
    value = config.get("path", config.get("file", ""))
    return normalize_text(value)


def profile_input_sheets(field_profile: dict[str, object] | None) -> list[str]:
    config = profile_input_config(field_profile)
    value = config.get("sheets", config.get("sheet_names", config.get("tables", [])))
    if value is None:
        return []
    if isinstance(value, str):
        return [normalize_text(value)] if normalize_text(value) else []
    if isinstance(value, (list, tuple)):
        sheets: list[str] = []
        for item in value:
            if isinstance(item, dict):
                item = item.get("sheet", item.get("name", ""))
            sheet = normalize_text(item)
            if sheet:
                sheets.append(sheet)
        return sheets
    raise ValueError("Field profile 'input.sheets' must be a string or list.")


def resolve_input_path(
    cli_input: Path | None,
    field_profile: dict[str, object],
    field_profile_path: Path | None,
) -> Path:
    if cli_input is not None:
        return cli_input.resolve()

    configured_path = profile_input_path(field_profile)
    if not configured_path:
        raise ValueError("Input file is required. Pass it as an argument or set input.path in --field-profile YAML.")

    path = Path(configured_path)
    if path.is_absolute():
        return path.resolve()

    base_dir = field_profile_path.parent if field_profile_path is not None else PROJECT_ROOT
    return (base_dir / path).resolve()


def resolve_excel_sheet_name(requested_sheet: str, available_sheets: list[str]) -> str:
    normalized_available = {normalize_text(sheet): sheet for sheet in available_sheets}
    normalized_requested = normalize_text(requested_sheet)
    if normalized_requested in normalized_available:
        return normalized_available[normalized_requested]

    if normalized_requested.endswith("表"):
        without_table_suffix = normalized_requested[:-1]
        if without_table_suffix in normalized_available:
            return normalized_available[without_table_suffix]

    with_table_suffix = f"{normalized_requested}表"
    if with_table_suffix in normalized_available:
        return normalized_available[with_table_suffix]

    available_text = "、".join(available_sheets)
    raise ValueError(f"Excel sheet not found: {requested_sheet}. Available sheets: {available_text}")


def drop_blank_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    blank_mask = df.apply(lambda row: all(normalize_text(value) == "" for value in row), axis=1)
    return df.loc[~blank_mask].reset_index(drop=True)


def drop_duplicate_columns(df: pd.DataFrame) -> pd.DataFrame:
    if not df.columns.has_duplicates:
        return df
    return df.loc[:, ~df.columns.duplicated()].copy()


def read_excel_input(path: Path, sheet_names: list[str]) -> pd.DataFrame:
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message=".*extension is not supported.*", category=UserWarning)
        excel_file_context = pd.ExcelFile(path)
    with excel_file_context as excel_file:
        if not sheet_names:
            if len(excel_file.sheet_names) != 1:
                available_text = "、".join(excel_file.sheet_names)
                raise ValueError(
                    f"Excel input has multiple sheets. Set input.sheets in YAML. Available sheets: {available_text}"
                )
            sheet_names = [excel_file.sheet_names[0]]

        frames: list[pd.DataFrame] = []
        for requested_sheet in sheet_names:
            sheet_name = resolve_excel_sheet_name(requested_sheet, excel_file.sheet_names)
            with warnings.catch_warnings():
                warnings.filterwarnings("ignore", message=".*extension is not supported.*", category=UserWarning)
                frame = pd.read_excel(excel_file, sheet_name=sheet_name, dtype=str, keep_default_na=False)
            frame = drop_blank_rows(frame)
            if not frame.empty:
                frames.append(frame)
            print(f"读取工作表: {sheet_name} ({len(frame)} rows)")

    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True, sort=False)


def safe_path_part(value: object) -> str:
    text = normalize_text(value)
    text = re.sub(r'[<>:"/\\|?*]+', "_", text)
    text = re.sub(r"\s+", " ", text).strip(" .")
    return text or "sheet"


def read_excel_input_tables(path: Path, sheet_names: list[str]) -> list[tuple[str, Path, pd.DataFrame]]:
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message=".*extension is not supported.*", category=UserWarning)
        excel_file_context = pd.ExcelFile(path)
    with excel_file_context as excel_file:
        if not sheet_names:
            if len(excel_file.sheet_names) != 1:
                available_text = "、".join(excel_file.sheet_names)
                raise ValueError(
                    f"Excel input has multiple sheets. Set input.sheets in YAML. Available sheets: {available_text}"
                )
            sheet_names = [excel_file.sheet_names[0]]

        tables: list[tuple[str, Path, pd.DataFrame]] = []
        for requested_sheet in sheet_names:
            sheet_name = resolve_excel_sheet_name(requested_sheet, excel_file.sheet_names)
            with warnings.catch_warnings():
                warnings.filterwarnings("ignore", message=".*extension is not supported.*", category=UserWarning)
                frame = pd.read_excel(excel_file, sheet_name=sheet_name, dtype=str, keep_default_na=False)
            frame = drop_blank_rows(frame)
            table_stem = f"{path.stem}_{safe_path_part(sheet_name)}"
            table_path = path.with_name(f"{table_stem}{path.suffix}")
            tables.append((sheet_name, table_path, frame))
            print(f"读取工作表: {sheet_name} ({len(frame)} rows)")
    return tables


def read_input_data(path: Path, encoding: str, field_profile: dict[str, object]) -> pd.DataFrame:
    if path.suffix.lower() in EXCEL_SUFFIXES:
        return read_excel_input(path, profile_input_sheets(field_profile))
    return read_tsv(path, encoding=encoding)


def read_input_tables(path: Path, encoding: str, field_profile: dict[str, object]) -> list[tuple[str, Path, pd.DataFrame]]:
    if path.suffix.lower() in EXCEL_SUFFIXES:
        return read_excel_input_tables(path, profile_input_sheets(field_profile))
    return [(path.stem, path, read_tsv(path, encoding=encoding))]


def parse_year_list(value: object) -> list[int]:
    text = normalize_text(value).replace("，", "/").replace(",", "/")
    years: list[int] = []

    for part in text.split("/"):
        part = part.strip()
        if not part:
            continue

        pieces = [item.strip() for item in part.split("-")]
        start = to_int(pieces[0]) if pieces else None
        end = to_int(pieces[1]) if len(pieces) >= 2 else start

        if start is None or end is None:
            continue

        lo, hi = sorted((start, end))
        years.extend(range(lo, hi + 1))

    return sorted(set(years))


def to_int(value: object) -> int | None:
    text = normalize_text(value)
    if not text:
        return None

    try:
        return int(float(text))
    except ValueError:
        match = re.search(r"\d{4}", text)
        return int(match.group(0)) if match else None


def combine_year_text(values: pd.Series) -> str:
    years: set[int] = set()
    for value in values:
        for part in normalize_text(value).split("/"):
            year = to_int(part)
            if year is not None:
                years.add(year)
    return "/".join(str(year) for year in sorted(years))


def combine_unique_text(values: pd.Series) -> str:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        for part in normalize_text(value).split("/"):
            text = normalize_text(part)
            if text and text not in seen:
                seen.add(text)
                result.append(text)
    return "/".join(result)


def split_const_atoms(value: object) -> list[str]:
    text = normalize_text(value)
    if not text:
        return [""]

    atoms: list[str] = []
    seen: set[str] = set()
    for part in text.split("/"):
        const = normalize_text(part)
        if const and const not in seen:
            seen.add(const)
            atoms.append(const)
    return atoms or [""]


def join_unique(values: list[str]) -> str:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        for part in normalize_text(value).split("/"):
            text = normalize_text(part)
            if text and text not in seen:
                seen.add(text)
                result.append(text)
    return "/".join(result)


def text_sort_key(value: object) -> tuple[tuple[int, object], ...]:
    text = normalize_text(value)
    parts: list[tuple[int, object]] = []
    for part in re.split(r"(\d+)", text.casefold()):
        if not part:
            continue
        if part.isdigit():
            parts.append((0, int(part)))
        else:
            parts.append((1, part))
    return tuple(parts)


def join_unique_sorted(values: list[str]) -> str:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        for part in normalize_text(value).split("/"):
            text = normalize_text(part)
            if text and text not in seen:
                seen.add(text)
                result.append(text)
    return "/".join(sorted(result, key=text_sort_key))


def split_joined_text(value: object) -> set[str]:
    text = normalize_text(value)
    if not text:
        return set()
    return set(part for part in (normalize_text(item) for item in text.split("/")) if part)


def split_joined_atoms(value: object) -> list[str]:
    text = normalize_text(value)
    if not text:
        return [""]
    result: list[str] = []
    seen: set[str] = set()
    for part in text.split("/"):
        atom = normalize_text(part)
        if atom and atom not in seen:
            seen.add(atom)
            result.append(atom)
    return result or [""]


def const_sets_overlap(left: set[str], right: set[str]) -> bool:
    return not left or not right or bool(left & right)


def split_version_tokens(value: object) -> list[str]:
    text = normalize_text(value)
    if not text:
        return []

    text = re.sub(r"\b(?:INCL|Incl|incl|EXCL|Excl|excl|EXP|Exp|exp)\s*:", "", text)
    text = re.sub(r"\b([A-Za-z])\s*/\s*([A-Za-z])\b", r"\1__VERSION_SLASH__\2", text)
    result: list[str] = []
    seen: set[str] = set()
    for part in text.split("/"):
        version = normalize_text(part).replace("__VERSION_SLASH__", "/")
        if version and version not in seen:
            seen.add(version)
            result.append(version)
    return result


def version_door_key(value: object) -> str:
    doors: set[str] = set()
    for token in split_version_tokens(value):
        text = normalize_text(token).lower()
        for match in re.finditer(r"\b([24])\s*(?:-| )?\s*(?:door|dr)\b", text):
            doors.add(match.group(1))
        for match in re.finditer(r"\b([24])dr\b", text):
            doors.add(match.group(1))
    return "/".join(sorted(doors))


def normalize_door_key(value: object) -> str:
    text = normalize_text(value).lower()
    if not text:
        return ""

    doors: set[str] = set()
    for match in re.finditer(r"\b([24])\s*(?:-| )?\s*(?:door|doors|dr|门)?\b", text):
        doors.add(match.group(1))
    return "/".join(sorted(doors))


def resolve_door_key(door_value: object, version_value: object = "") -> str:
    return normalize_door_key(door_value) or version_door_key(version_value)


def combine_version_tokens(values: pd.Series, add_incl: bool = False) -> str:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        for version in split_version_tokens(value):
            if version not in seen:
                seen.add(version)
                result.append(version)

    if not result:
        return ""
    text = "/".join(result)
    return f"Incl: {text}" if add_incl else text


def combine_versions(values: pd.Series) -> str:
    versions = [normalize_text(value) for value in values]
    has_base = "" in versions
    non_blank = combine_version_tokens(pd.Series([value for value in versions if value]))

    if not non_blank:
        return ""
    if has_base:
        return f"Incl: {non_blank}"
    return non_blank


def split_dedupe_versions(values: pd.Series, add_incl: bool = False) -> str:
    return combine_version_tokens(values, add_incl=add_incl)


def combine_pickup_versions(values: pd.Series) -> str:
    versions = [normalize_text(value) for value in values]
    return split_dedupe_versions(pd.Series(versions), add_incl="" in versions)


def append_excl(version: object, excl_list: object) -> str:
    version_text = normalize_text(version)
    excl_text = normalize_text(excl_list)
    if not excl_text:
        return version_text
    return f"Excl: {excl_text}" if not version_text else f"{version_text} Excl: {excl_text}"


def build_version_text(incl_versions: list[str], excl_versions: list[str]) -> str:
    parts: list[str] = []
    incl = join_unique_sorted(incl_versions)
    excl = join_unique_sorted(excl_versions)
    if incl:
        parts.append(f"Incl: {incl}")
    if excl:
        parts.append(f"Excl: {excl}")
    return " ".join(parts)


def format_year_range(years: list[int]) -> str:
    if not years:
        return ""
    minimum = min(years)
    maximum = max(years)
    return str(minimum) if minimum == maximum else f"{minimum}-{maximum}"


def format_year_segments(years: list[int]) -> str:
    sorted_years = sorted(set(years))
    if not sorted_years:
        return ""

    segments: list[str] = []
    start = sorted_years[0]
    previous = sorted_years[0]
    for year in sorted_years[1:]:
        if year != previous + 1:
            segments.append(format_year_range([start, previous]))
            start = year
        previous = year

    segments.append(format_year_range([start, previous]))
    return ";".join(segments)


def format_year_segments_from_text(value: object) -> str:
    return format_year_segments(parse_year_list(value))


def continuous_year_segments(years: list[int]) -> list[tuple[int, int]]:
    sorted_years = sorted(set(years))
    if not sorted_years:
        return []

    segments: list[tuple[int, int]] = []
    start = sorted_years[0]
    previous = sorted_years[0]
    for year in sorted_years[1:]:
        if year != previous + 1:
            segments.append((start, previous))
            start = year
        previous = year

    segments.append((start, previous))
    return segments


def split_rows_by_year_boundaries(
    source: pd.DataFrame,
    boundary_rows: pd.DataFrame,
    model_keys: list[str],
) -> pd.DataFrame:
    if source.empty:
        return source

    pieces: list[pd.Series] = []
    for model_values, model_boundary_rows in boundary_rows.groupby(model_keys, dropna=False, sort=False):
        if not isinstance(model_values, tuple):
            model_values = (model_values,)

        boundary_points: set[int] = set()
        for value in model_boundary_rows["原始年份列表"]:
            for start, end in continuous_year_segments(parse_year_list(value)):
                boundary_points.add(start)
                boundary_points.add(end + 1)

        model_source = source.copy()
        for key, value in zip(model_keys, model_values):
            model_source = model_source[model_source[key] == value]

        for _, row in model_source.iterrows():
            for start, end in continuous_year_segments(parse_year_list(row["原始年份列表"])):
                cut_points = sorted(point for point in boundary_points if start < point <= end)
                segment_start = start
                for cut_point in [*cut_points, end + 1]:
                    segment_end = cut_point - 1
                    piece = row.copy()
                    piece["year_start"] = segment_start
                    piece["year_end"] = segment_end
                    piece["原始年份列表"] = "/".join(str(year) for year in range(segment_start, segment_end + 1))
                    pieces.append(piece)
                    segment_start = cut_point

    if not pieces:
        return source.iloc[0:0].copy()
    return pd.DataFrame(pieces).reset_index(drop=True)


def combine_excl_versions(rows: pd.DataFrame) -> str:
    if rows.empty:
        return ""
    special_versions = rows[rows["VERSION_RAW"].map(normalize_text) != ""]["VERSION_RAW"]
    return combine_version_tokens(special_versions)


def bridge_status(year_text: object) -> str:
    years = parse_year_list(year_text)
    if not years:
        return "连续年份"

    segment_count = 1
    for previous, current in zip(years, years[1:]):
        if current != previous + 1:
            segment_count += 1

    return "已桥接断档" if segment_count > 1 else "连续年份"


def build_excl_list(group: pd.DataFrame) -> str:
    special_rows = group[
        (group["BackSize"].map(normalize_text) == "无可用尺码")
        & (group["VERSION"].map(normalize_text).str.lower().str.startswith("incl:"))
    ]
    versions = []
    for value in special_rows["VERSION"]:
        version = normalize_text(value)
        versions.extend(split_version_tokens(version))
    return "/".join(sorted(set(value for value in versions if value)))


def build_pickup_excl_list(group: pd.DataFrame) -> str:
    special_rows = group[
        (group["BackSize"].map(normalize_text) == "无可用尺码")
        & (group["VERSION"].map(normalize_text).str.lower().str.startswith("incl:"))
    ]
    return split_dedupe_versions(special_rows["VERSION"], add_incl=False)


def combine_bed_ft(values: pd.Series) -> str:
    bed_numbers: list[float] = []
    text_parts: list[str] = []
    seen_text: set[str] = set()
    for value in values:
        for part in normalize_text(value).split("/"):
            bed = normalize_text(part)
            if not bed:
                continue
            if "-" in bed:
                left, right = (normalize_text(item) for item in bed.split("-", 1))
                left_number = parse_bed_number(left)
                right_number = parse_bed_number(right)
                if left_number is not None and right_number is not None:
                    bed_numbers.extend([left_number, right_number])
                    continue
            bed_number = parse_bed_number(bed)
            if bed_number is not None:
                bed_numbers.append(bed_number)
            elif bed not in seen_text:
                seen_text.add(bed)
                text_parts.append(bed)

    if bed_numbers and not text_parts:
        minimum = min(bed_numbers)
        maximum = max(bed_numbers)
        if minimum == maximum:
            return format_number(minimum)
        return f"{format_number(minimum)}-{format_number(maximum)}"

    if not bed_numbers and not text_parts:
        return ""

    if bed_numbers:
        minimum = min(bed_numbers)
        maximum = max(bed_numbers)
        text_parts.insert(0, format_number(minimum) if minimum == maximum else f"{format_number(minimum)}-{format_number(maximum)}")

    return "/".join(text_parts)


def parse_bed_number(value: object) -> float | None:
    text = normalize_text(value)
    if not text:
        return None

    try:
        return float(text)
    except ValueError:
        return None


def bed_matches_expression(bed_value: object, expression: object) -> bool:
    bed_text = normalize_text(bed_value)
    expression_text = normalize_text(expression)
    if not bed_text or not expression_text:
        return bed_text == expression_text
    bed_number = parse_bed_number(bed_text)
    expression_number = parse_bed_number(expression_text)
    if bed_number is not None and expression_number is not None:
        return bed_number == expression_number

    if "/" in expression_text:
        return any(bed_matches_expression(bed_text, part) for part in expression_text.split("/"))

    if "-" in expression_text:
        left, right = (normalize_text(part) for part in expression_text.split("-", 1))
        bed_number = parse_bed_number(bed_text)
        left_number = parse_bed_number(left)
        right_number = parse_bed_number(right)
        if bed_number is not None and left_number is not None and right_number is not None:
            lo, hi = sorted((left_number, right_number))
            return lo <= bed_number <= hi

    return bed_text == expression_text


def format_number(value: float) -> str:
    return str(int(value)) if value.is_integer() else str(value)


def split_model_raw_list(value: object) -> list[str]:
    text = normalize_text(value)
    text = text.replace("；", ";").replace("\n", ";").replace("\r", ";").replace("｜", "|")
    if not text:
        return []

    result: list[str] = []
    seen: set[str] = set()
    for part in text.split(";"):
        model_raw = normalize_text(part)
        if model_raw and model_raw not in seen:
            seen.add(model_raw)
            result.append(model_raw)
    return result


def parse_year_range(value: object) -> list[int]:
    text = normalize_text(value)
    if not text:
        return []

    years: list[int] = []
    for segment in text.replace("；", ";").split(";"):
        parts = [part.strip() for part in segment.split("-")]
        start = to_int(parts[0]) if parts else None
        end = to_int(parts[1]) if len(parts) > 1 else start

        if start is None or end is None:
            continue

        lo, hi = sorted((start, end))
        years.extend(range(lo, hi + 1))

    return sorted(set(years))


def start_year_from_range(value: object) -> int | None:
    years = parse_year_list(value)
    return min(years) if years else None


def resolve_start_year(start_value: object, year_range_value: object) -> int | None:
    start = to_int(start_value)
    if start is not None:
        return start
    return start_year_from_range(year_range_value)


def split_front_model_atoms(value: object) -> list[str]:
    text = normalize_text(value)
    if not text:
        return [""]
    result: list[str] = []
    seen: set[str] = set()
    for part in text.split("|"):
        model = normalize_text(part)
        if model and model not in seen:
            seen.add(model)
            result.append(model)
    return result or [""]


def split_model_expression(value: object) -> list[str]:
    text = normalize_text(value)
    if not text:
        return []
    result: list[str] = []
    seen: set[str] = set()
    for part in text.split("/"):
        model = normalize_text(part)
        if model and model not in seen:
            seen.add(model)
            result.append(model)
    return result


def model_expression_matches(record_model: object, atom_model: object) -> bool:
    if normalize_text(record_model) == normalize_text(atom_model):
        return True
    models = split_model_expression(record_model)
    atom = normalize_text(atom_model)
    return atom in models if models else atom == ""


def combine_model_expression(values: list[object]) -> str:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        for model in split_model_expression(value):
            if model and model not in seen:
                seen.add(model)
                result.append(model)
    return "/".join(result)


def combine_model_expression_for_brand(brand: object, values: list[object]) -> str:
    combined = split_model_expression(combine_model_expression(values))
    if len(combined) <= 1:
        return "/".join(combined)

    brand_text = normalize_text(brand)
    combo_map = load_model_combo_map()
    combo_group = ""
    for model in combined:
        group = combo_map.get((brand_text.casefold(), model.casefold()), "")
        if group:
            combo_group = group
            break
    if not combo_group:
        return "/".join(combined)

    order = {model.casefold(): index for index, model in enumerate(combo_group.split("|"))}
    return "/".join(sorted(combined, key=lambda model: order.get(model.casefold(), len(order))))


def split_output_tokens(value: object) -> set[str]:
    return split_joined_text(value)


def record_version_matches(record_version: object, atom_version: object) -> bool:
    record_text = normalize_text(record_version)
    atom_text = normalize_text(atom_version)
    if not record_text:
        return atom_text == ""
    atom_tokens = set(split_version_tokens(atom_text))
    lower = record_text.lower()
    if lower.startswith("incl:"):
        tokens = set(split_version_tokens(record_text))
        return atom_text == "" or atom_text in tokens or bool(atom_tokens & tokens)
    tokens = set(split_version_tokens(record_text))
    return atom_text in tokens or bool(atom_tokens & tokens)


def combine_lossy_versions(values: list[object]) -> str:
    normalized = [normalize_text(value) for value in values]
    has_base = any(value == "" or value.lower().startswith("incl:") for value in normalized)
    tokens: list[str] = []
    seen: set[str] = set()
    for value in normalized:
        if not value:
            continue
        for token in split_version_tokens(value):
            if token not in seen:
                seen.add(token)
                tokens.append(token)
    tokens = sorted(tokens, key=text_sort_key)
    if has_base and tokens:
        return f"Incl: {'/'.join(tokens)}"
    return "/".join(tokens)


def require_columns(df: pd.DataFrame, required_columns: list[str] | None = None) -> None:
    columns = required_columns or REQUIRED_COLUMNS
    missing = [column for column in columns if column not in df.columns]
    if missing:
        raise ValueError("Input TSV is missing required columns: " + ", ".join(missing))


def has_columns(df: pd.DataFrame, columns: list[str]) -> bool:
    return all(column in df.columns for column in columns)


def empty_non_pickup_result() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    summary = {
        "model_groups": 0,
        "yearcross_true_groups": 0,
        "yearcross_false_groups": 0,
        "special_atomic_rows": 0,
        "special_combined_incl": 0,
        "special_excluded": 0,
        "special_unmatched": 0,
        "base_atomic_rows": 0,
        "output_rows": 0,
        "lossless_rows": 0,
        "high_rows": 0,
        "higher_rows": 0,
        "process_rows": [],
        "compression_log": empty_log_table(),
    }
    result = pd.DataFrame(columns=NON_PICKUP_FINAL_COLUMNS)
    high = result.copy()
    higher = result.copy()
    for item in [result, high, higher]:
        item.attrs["summary"] = summary
        item.attrs["atom_table"] = empty_atom_table()
    return result, high, higher


def empty_pickup_result() -> tuple[pd.DataFrame, pd.DataFrame]:
    summary = {
        "pickup_atomic_rows": 0,
        "pickup_bed_merged_year_groups": 0,
        "pickup_gap_closed_groups": 0,
        "pickup_cab_bed_closed_groups": 0,
        "process_rows": [],
        "compression_log": empty_log_table(),
    }
    lossless = pd.DataFrame(columns=PICKUP_FINAL_COLUMNS)
    specificity = lossless.copy()
    for item in [lossless, specificity]:
        item.attrs["summary"] = summary
        item.attrs["atom_table"] = empty_atom_table()
    return lossless, specificity


def empty_atom_table() -> pd.DataFrame:
    return pd.DataFrame(columns=ATOM_FINAL_COLUMNS)


def empty_log_table() -> pd.DataFrame:
    return pd.DataFrame(columns=LOG_FINAL_COLUMNS)


def pair_text(left: object, right: object) -> str:
    return f"{normalize_text(left)}/{normalize_text(right)}"


def candidate_log_fields(
    left: pd.Series,
    right: pd.Series,
    merged: dict[str, object],
    include_const: bool = True,
) -> dict[str, object]:
    return {
        "CAB": pair_text(left.get("CAB", ""), right.get("CAB", "")),
        "BED_FT": pair_text(left.get("BED_FT", ""), right.get("BED_FT", "")),
        "候选YEAR": pair_text(left.get("YEAR", ""), right.get("YEAR", "")),
        "候选CONST": pair_text(left.get("CONST", ""), right.get("CONST", "")) if include_const else "",
        "候选VERSION": pair_text(left.get("VERSION", ""), right.get("VERSION", "")),
        "合并CAB": normalize_text(merged.get("CAB", "")),
        "合并BED_FT": normalize_text(merged.get("BED_FT", "")),
        "合并YEAR": normalize_text(merged.get("YEAR", "")),
        "合并CONST": normalize_text(merged.get("CONST", "")) if include_const else "",
        "合并VERSION": normalize_text(merged.get("VERSION", "")),
    }


def build_non_pickup_atom_table(atoms: pd.DataFrame) -> pd.DataFrame:
    if atoms.empty:
        return empty_atom_table()

    result = pd.DataFrame(
        {
            "压缩类型": "非皮卡",
            "BRAND": atoms["BRAND"].map(normalize_text),
            "MODEL": atoms["MODEL"].map(normalize_text),
            "YEAR": atoms["YEAR_SINGLE"].astype(int),
            "VERSION": atoms["VERSION_RAW"].map(normalize_text),
            "CONST": atoms["Const"].map(normalize_text),
            "CAB": "",
            "BED_FT": "",
            "BACKSIZE": atoms["BackSize"].map(normalize_text),
        }
    )
    return result[ATOM_FINAL_COLUMNS].sort_values(["压缩类型", "BRAND", "MODEL", "YEAR", "BACKSIZE", "CONST", "VERSION"], kind="mergesort").reset_index(drop=True)


def build_pickup_atom_table(atoms: pd.DataFrame) -> pd.DataFrame:
    if atoms.empty:
        return empty_atom_table()

    result = pd.DataFrame(
        {
            "压缩类型": "皮卡",
            "BRAND": atoms["BRAND"].map(normalize_text),
            "MODEL": atoms["MODEL"].map(normalize_text),
            "YEAR": atoms["YEAR_SINGLE"].astype(int),
            "VERSION": atoms["VERSION_RAW"].map(normalize_text),
            "CONST": "",
            "CAB": atoms["CAB"].map(normalize_text),
            "BED_FT": atoms["BED_FT"].map(normalize_text),
            "BACKSIZE": atoms["BackSize"].map(normalize_text),
        }
    )
    return result[ATOM_FINAL_COLUMNS].sort_values(["压缩类型", "BRAND", "MODEL", "YEAR", "BACKSIZE", "CAB", "BED_FT", "VERSION"], kind="mergesort").reset_index(drop=True)


def make_non_pickup_output_record(
    brand: str,
    model: str,
    start: int,
    end: int,
    backsize: str,
    const: str,
    version: str,
) -> dict[str, object]:
    return {
        "主车型": " ".join(part for part in [brand, model] if part),
        "BRAND": brand,
        "MODEL": model,
        "YEAR": format_year_range([start, end]),
        "VERSION": version,
        "RAW-CONST": const,
        "CONST": const,
        "BACKSIZE": backsize,
        "year_start": start,
        "year_end": end,
        "START_YEAR": start,
    }


def build_non_pickup_lossless_new(atoms: pd.DataFrame) -> pd.DataFrame:
    if atoms.empty:
        return pd.DataFrame(columns=[*NON_PICKUP_FINAL_COLUMNS, "year_start", "year_end", "START_YEAR"])

    records: list[dict[str, object]] = []
    group_keys = ["BRAND", "MODEL", "BackSize", "Const", "VERSION_RAW"]
    atoms = atoms.sort_values(["BRAND", "MODEL", "START_YEAR", "YEAR_SINGLE"], kind="mergesort")
    for key_values, group in atoms.groupby(group_keys, dropna=False, sort=False):
        brand, model, backsize, const, version = (
            normalize_text(value) for value in (key_values if isinstance(key_values, tuple) else (key_values,))
        )
        years = sorted(set(group["YEAR_SINGLE"].astype(int)))
        for start, end in continuous_year_segments(years):
            records.append(make_non_pickup_output_record(brand, model, start, end, backsize, const, version))

    if not records:
        return pd.DataFrame(columns=[*NON_PICKUP_FINAL_COLUMNS, "year_start", "year_end", "START_YEAR"])
    return pd.DataFrame(records).sort_values(["BRAND", "MODEL", "START_YEAR", "year_start", "BACKSIZE", "CONST", "VERSION"], kind="mergesort").reset_index(drop=True)


def merge_non_pickup_high_records(left: pd.Series, right: pd.Series) -> dict[str, object]:
    brand = normalize_text(left["BRAND"])
    model = combine_model_expression_for_brand(brand, [left["MODEL"], right["MODEL"]])
    start = min(int(left["year_start"]), int(right["year_start"]))
    end = max(int(left["year_end"]), int(right["year_end"]))
    return make_non_pickup_output_record(
        brand,
        model,
        start,
        end,
        normalize_text(left["BACKSIZE"]),
        join_unique_sorted([normalize_text(left.get("CONST", "")), normalize_text(right.get("CONST", ""))]),
        combine_lossy_versions([left.get("VERSION", ""), right.get("VERSION", "")]),
    )


def build_non_pickup_high_new(
    lossless: pd.DataFrame,
    atoms: pd.DataFrame,
    progress: ProgressReporter | None = None,
) -> tuple[pd.DataFrame, dict[tuple[str, str], dict[str, int]], pd.DataFrame]:
    stats: dict[tuple[str, str], dict[str, int]] = {}
    log_rows: list[dict[str, object]] = []
    if lossless.empty:
        return lossless.copy(), stats, empty_log_table()

    output_groups: list[pd.DataFrame] = []
    lossless_work = lossless.copy()
    atoms_work = atoms.copy()
    lossless_work["__MODEL_GROUP"] = lossless_work.apply(lambda row: model_combo_group(row["BRAND"], row["MODEL"]), axis=1)
    atoms_work["__MODEL_GROUP"] = atoms_work.apply(lambda row: model_combo_group(row["BRAND"], row["MODEL"]), axis=1)
    total_groups = int(lossless_work.groupby(["BRAND", "__MODEL_GROUP"], dropna=False, sort=False).ngroups)
    completed_groups = 0
    merge_count = 0
    attempt_count = 0
    if progress:
        progress.start("非皮卡高度压缩", total_groups)
    for group_values, group in lossless_work.groupby(["BRAND", "__MODEL_GROUP"], dropna=False, sort=False):
        brand, model_group = group_values
        key = (normalize_text(brand), normalize_text(model_group))
        key_stats = stats.setdefault(key, {"attempts": 0, "successes": 0, "fallbacks": 0})
        if progress:
            progress.update(current_make=key[0], current_model=key[1], completed_models=completed_groups)
        atoms_group = atoms_work[
            (atoms_work["BRAND"].map(normalize_text) == key[0])
            & (atoms_work["__MODEL_GROUP"].map(normalize_text) == key[1])
        ].drop(columns=["__MODEL_GROUP"])
        group = group.drop(columns=["__MODEL_GROUP"])
        working = group.sort_values(["START_YEAR", "year_start", "year_end", "MODEL", "BACKSIZE", "CONST", "VERSION"], kind="mergesort").reset_index(drop=True)

        changed = True
        while changed and len(working) > 1:
            changed = False
            for left_index in range(len(working)):
                if changed:
                    break
                for right_index in range(left_index + 1, len(working)):
                    left = working.iloc[left_index]
                    right = working.iloc[right_index]
                    if normalize_text(left.get("BACKSIZE", "")) != normalize_text(right.get("BACKSIZE", "")):
                        continue
                    key_stats["attempts"] += 1
                    attempt_count += 1
                    merged = merge_non_pickup_high_records(left, right)
                    candidate_rows = [
                        *working.iloc[:left_index].to_dict("records"),
                        *working.iloc[left_index + 1 : right_index].to_dict("records"),
                        *working.iloc[right_index + 1 :].to_dict("records"),
                        merged,
                    ]
                    candidate = pd.DataFrame(candidate_rows, columns=working.columns)
                    scoped_atoms = non_pickup_atoms_in_record_scope(atoms_group, merged)
                    reason = (
                        "候选合并范围内没有可验证原子事实"
                        if scoped_atoms.empty
                        else non_pickup_candidate_validation_reason(candidate, scoped_atoms)
                    )
                    log_base = {
                        "压缩类型": "非皮卡",
                        "阶段": "高度压缩两两组合",
                        "BRAND": key[0],
                        "MODEL": pair_text(left.get("MODEL", ""), right.get("MODEL", "")),
                        "合并MODEL": normalize_text(merged["MODEL"]),
                        "BACKSIZE": normalize_text(merged["BACKSIZE"]),
                        **candidate_log_fields(left, right, merged, include_const=True),
                    }
                    if not reason:
                        key_stats["successes"] += 1
                        merge_count += 1
                        log_rows.append({**log_base, "结果": "success", "原因": ""})
                        if progress:
                            progress.update(merge_count=merge_count, attempt_count=attempt_count)
                        working = candidate.sort_values(
                            ["START_YEAR", "year_start", "year_end", "BACKSIZE", "CONST", "VERSION"], kind="mergesort"
                        ).reset_index(drop=True)
                        changed = True
                        break
                    key_stats["fallbacks"] += 1
                    log_rows.append({**log_base, "结果": "fallback", "原因": reason})
                    if progress:
                        progress.update(merge_count=merge_count, attempt_count=attempt_count)

        output_groups.append(working)
        completed_groups += 1
        if progress:
            progress.update(completed_models=completed_groups, merge_count=merge_count, attempt_count=attempt_count)

    if not output_groups:
        return lossless.iloc[0:0].copy(), stats, pd.DataFrame(log_rows, columns=LOG_FINAL_COLUMNS)
    if progress:
        progress.finish()
    result = pd.concat(output_groups, ignore_index=True).sort_values(
        ["BRAND", "MODEL", "START_YEAR", "year_start", "BACKSIZE", "CONST", "VERSION"], kind="mergesort"
    ).reset_index(drop=True)
    return result, stats, pd.DataFrame(log_rows, columns=LOG_FINAL_COLUMNS)


def row_const_atoms(row: pd.Series) -> set[str]:
    consts = split_joined_text(row.get("Const", ""))
    if consts:
        return consts
    raw_consts = split_joined_text(row.get("RawConst", ""))
    return raw_consts or {"*"}


def row_included_version_tokens(row: pd.Series) -> list[str]:
    version = normalize_text(row.get("VERSION", ""))
    if not version or "Excl:" in version:
        return []
    return split_version_tokens(version)


def non_pickup_version_parts(value: object) -> tuple[bool, set[str], set[str]]:
    text = normalize_text(value)
    if not text:
        return True, set(), set()

    lower = text.lower()
    has_label = "incl:" in lower or "excl:" in lower
    has_base = has_label
    incl_tokens: set[str] = set()
    excl_tokens: set[str] = set()

    incl_match = re.search(r"\bincl\s*:\s*(.*?)(?=\bexcl\s*:|$)", text, flags=re.IGNORECASE)
    if incl_match:
        incl_tokens.update(split_version_tokens(incl_match.group(1)))

    excl_match = re.search(r"\bexcl\s*:\s*(.*)$", text, flags=re.IGNORECASE)
    if excl_match:
        excl_tokens.update(split_version_tokens(excl_match.group(1)))

    if not has_label:
        incl_tokens.update(split_version_tokens(text))

    return has_base, incl_tokens, excl_tokens


def non_pickup_record_covers_atom(record: pd.Series, atom: pd.Series) -> bool:
    if normalize_text(record.get("BRAND", "")) != normalize_text(atom.get("BRAND", "")):
        return False
    if not model_expression_matches(record.get("MODEL", ""), atom.get("MODEL", "")):
        return False
    if normalize_text(record.get("BACKSIZE", "")) != normalize_text(atom.get("BackSize", "")):
        return False
    if int(atom["YEAR_SINGLE"]) not in parse_year_list(record.get("YEAR", "")):
        return False

    record_consts = split_joined_text(record.get("CONST", "")) or split_joined_text(record.get("RAW-CONST", ""))
    atom_consts = row_const_atoms(atom)
    if record_consts and atom_consts and not (record_consts & atom_consts):
        return False

    record_has_base, record_incl, record_excl = non_pickup_version_parts(record.get("VERSION", ""))
    atom_has_base, atom_incl, _atom_excl = non_pickup_version_parts(atom.get("VERSION", ""))
    if atom_has_base:
        return record_has_base

    atom_tokens = atom_incl
    if not atom_tokens:
        return record_has_base
    if record_has_base and atom_tokens & record_excl:
        return False
    return bool(atom_tokens & record_incl)


def non_pickup_records_have_unique_atom_matches(rows: pd.DataFrame, atoms: pd.DataFrame) -> bool:
    for _, atom in atoms.iterrows():
        matches = 0
        for _, row in rows.iterrows():
            if non_pickup_record_covers_atom(row, atom):
                matches += 1
                if matches > 1:
                    return False
        if matches == 0:
            return False
    return True


def combine_non_pickup_versions(values: list[object]) -> str:
    has_base = False
    incl_tokens: list[str] = []
    excl_tokens: list[str] = []
    seen_incl: set[str] = set()
    seen_excl: set[str] = set()

    for value in values:
        value_has_base, value_incl, value_excl = non_pickup_version_parts(value)
        has_base = has_base or value_has_base
        for token in split_version_tokens("/".join(sorted(value_incl))):
            if token not in seen_incl:
                seen_incl.add(token)
                incl_tokens.append(token)
        for token in split_version_tokens("/".join(sorted(value_excl))):
            if token not in seen_excl:
                seen_excl.add(token)
                excl_tokens.append(token)

    if has_base:
        return build_version_text(incl_tokens, excl_tokens)
    return "/".join(sorted(incl_tokens, key=text_sort_key))


def merge_non_pickup_specificity_records(left: pd.Series, right: pd.Series) -> dict[str, object]:
    start = min(int(left["year_start"]), int(right["year_start"]))
    end = max(int(left["year_end"]), int(right["year_end"]))
    return {
        "主车型": left["主车型"],
        "BRAND": left["BRAND"],
        "MODEL": left["MODEL"],
        "YEAR": format_year_range([start, end]),
        "CONST": join_unique_sorted([normalize_text(left.get("CONST", "")), normalize_text(right.get("CONST", ""))]),
        "RAW-CONST": join_unique_sorted([normalize_text(left.get("RAW-CONST", "")), normalize_text(right.get("RAW-CONST", ""))]),
        "VERSION": combine_non_pickup_versions([left.get("VERSION", ""), right.get("VERSION", "")]),
        "BACKSIZE": left["BACKSIZE"],
        "DOOR_KEY": "",
        "year_start": start,
        "year_end": end,
    }


def merge_adjacent_non_pickup_specificity_rows(rows: pd.DataFrame, atoms: pd.DataFrame) -> pd.DataFrame:
    stats = {"adjacent_merge_attempts": 0, "adjacent_merge_successes": 0, "adjacent_merge_fallbacks": 0}
    if len(rows) <= 1:
        rows = rows.copy()
        rows.attrs["merge_stats"] = stats
        return rows

    working = rows.sort_values(["year_start", "year_end", "CONST", "VERSION"], kind="mergesort").reset_index(drop=True)
    i = 0
    while i < len(working) - 1:
        merged = merge_non_pickup_specificity_records(working.iloc[i], working.iloc[i + 1])
        stats["adjacent_merge_attempts"] += 1
        candidate_records = [
            *working.iloc[:i].to_dict("records"),
            merged,
            *working.iloc[i + 2 :].to_dict("records"),
        ]
        candidate = pd.DataFrame(candidate_records, columns=working.columns)
        if non_pickup_records_have_unique_atom_matches(candidate, atoms):
            stats["adjacent_merge_successes"] += 1
            working = candidate.sort_values(["year_start", "year_end", "CONST", "VERSION"], kind="mergesort").reset_index(
                drop=True
            )
            i = max(i - 1, 0)
        else:
            stats["adjacent_merge_fallbacks"] += 1
            i += 1

    working.attrs["merge_stats"] = stats
    return working


def split_non_pickup_structure_rows(rows: pd.DataFrame) -> pd.DataFrame:
    if rows.empty:
        return rows

    split_rows: list[dict[str, object]] = []
    for _, row in rows.iterrows():
        consts = split_joined_text(row.get("Const", ""))
        raw_consts = split_joined_text(row.get("RawConst", ""))
        if consts:
            for const in sorted(consts):
                item = row.to_dict()
                item["Const"] = const
                item["RawConst"] = const
                split_rows.append(item)
        elif raw_consts:
            for raw_const in sorted(raw_consts):
                item = row.to_dict()
                item["Const"] = ""
                item["RawConst"] = raw_const
                split_rows.append(item)
        else:
            split_rows.append(row.to_dict())

    return pd.DataFrame(split_rows)


def close_same_fact_year_gaps(rows: pd.DataFrame, atom_result: pd.DataFrame) -> pd.DataFrame:
    if rows.empty or "DOOR_KEY" not in rows.columns:
        return rows

    model_keys = ["主车型", "BRAND", "MODEL"]
    facts: dict[tuple[object, ...], set[tuple[int, str, str]]] = {}
    for _, atom in atom_result.iterrows():
        scope = tuple(atom[key] for key in [*model_keys, "DOOR_KEY"])
        scope_facts = facts.setdefault(scope, set())
        for const in row_const_atoms(atom):
            scope_facts.add((int(atom["YEAR_SINGLE"]), const, normalize_text(atom["BackSize"])))

    gap_keys = [*model_keys, "VERSION", "RAW-CONST", "CONST", "BACKSIZE", "DOOR_KEY"]
    closed_rows: list[dict[str, object]] = []
    for key_values, group in rows.groupby(gap_keys, dropna=False, sort=False):
        values = dict(zip(gap_keys, key_values if isinstance(key_values, tuple) else (key_values,)))
        years: set[int] = set()
        for _, row in group.iterrows():
            years.update(parse_year_list(row["YEAR"]))

        if not years:
            closed_rows.extend(group.to_dict("records"))
            continue

        consts = split_joined_text(values["CONST"]) or split_joined_text(values["RAW-CONST"]) or {"*"}
        scope = tuple(values[key] for key in [*model_keys, "DOOR_KEY"])
        scoped_facts = facts.get(scope, set())
        allowed_years: list[int] = []
        for year in range(min(years), max(years) + 1):
            if year in years:
                allowed_years.append(year)
                continue

            has_existing_fact = any(
                fact_year == year and fact_const in consts
                for fact_year, fact_const, _fact_size in scoped_facts
            )
            if not has_existing_fact:
                allowed_years.append(year)

        first = group.iloc[0].to_dict()
        for start, end in continuous_year_segments(allowed_years):
            if not any(year in years for year in range(start, end + 1)):
                continue
            row = first.copy()
            row["YEAR"] = format_year_range([start, end])
            row["year_start"] = start
            row["year_end"] = end
            closed_rows.append(row)

    return pd.DataFrame(closed_rows)


def build_non_pickup_table_from_atoms(
    atom_result: pd.DataFrame, mode: str, keep_internal_columns: bool = False
) -> pd.DataFrame:
    model_keys = ["主车型", "BRAND", "MODEL"]
    output_columns = NON_PICKUP_FINAL_COLUMNS
    internal_columns = [*output_columns, "DOOR_KEY"] if keep_internal_columns else output_columns
    if atom_result.empty:
        return pd.DataFrame(columns=internal_columns)

    atom_result = atom_result.copy()
    if "DOOR_KEY" not in atom_result.columns:
        atom_result["DOOR_KEY"] = atom_result["VERSION"].map(version_door_key)
    atom_result["DOOR_KEY"] = atom_result["DOOR_KEY"].map(normalize_text)

    if mode == "lossless":
        atom_result = split_non_pickup_structure_rows(atom_result)

    global_facts: dict[tuple[object, ...], set[tuple[int, str]]] = {}
    for _, row in atom_result.iterrows():
        model_tuple = tuple(row[key] for key in [*model_keys, "DOOR_KEY"])
        facts = global_facts.setdefault(model_tuple, set())
        for const in row_const_atoms(row):
            facts.add((int(row["YEAR_SINGLE"]), const))

    def make_record(group: pd.DataFrame, year_start: int, year_end: int) -> dict[str, object]:
        segment = group[(group["YEAR_SINGLE"] >= year_start) & (group["YEAR_SINGLE"] <= year_end)]
        first = group.iloc[0]
        return {
            **{key: first[key] for key in model_keys},
            "YEAR": format_year_range([year_start, year_end]),
            "CONST": join_unique(segment["Const"].map(normalize_text).tolist()),
            "RAW-CONST": join_unique(segment["RawConst"].map(normalize_text).tolist()),
            "VERSION": first["VERSION"],
            "BACKSIZE": first["BackSize"],
            "DOOR_KEY": normalize_text(first.get("DOOR_KEY", "")),
            "year_start": year_start,
            "year_end": year_end,
        }

    records: list[dict[str, object]] = []

    if mode == "lossless":
        group_keys = [*model_keys, "Const", "RawConst", "VERSION", "BackSize", "DOOR_KEY"]
        for _, group in atom_result.groupby(group_keys, dropna=False, sort=False):
            years = sorted(set(group["YEAR_SINGLE"].astype(int)))
            for start, end in continuous_year_segments(years):
                records.append(make_record(group, start, end))

    elif mode == "high":
        group_keys = [*model_keys, "VERSION", "BackSize", "Const", "RawConst", "DOOR_KEY"]
        yearset_rows: list[dict[str, object]] = []
        for _, group in atom_result.groupby(group_keys, dropna=False, sort=False):
            first = group.iloc[0]
            yearset_rows.append(
                {
                    **{key: first[key] for key in model_keys},
                    "VERSION": first["VERSION"],
                    "BackSize": first["BackSize"],
                    "Const": first["Const"],
                    "RawConst": first["RawConst"],
                    "DOOR_KEY": first["DOOR_KEY"],
                    "YEARSET": tuple(sorted(set(group["YEAR_SINGLE"].astype(int)))),
                }
            )

        yearset_df = pd.DataFrame(yearset_rows)
        for _, group in yearset_df.groupby(
            [*model_keys, "VERSION", "BackSize", "DOOR_KEY", "YEARSET"], dropna=False, sort=False
        ):
            years = list(group.iloc[0]["YEARSET"])
            if not years:
                continue
            first = group.iloc[0]
            for start, end in continuous_year_segments(years):
                records.append(
                    {
                        **{key: first[key] for key in model_keys},
                        "YEAR": format_year_range([start, end]),
                        "CONST": join_unique(group["Const"].map(normalize_text).tolist()),
                        "RAW-CONST": join_unique(group["RawConst"].map(normalize_text).tolist()),
                        "VERSION": first["VERSION"],
                        "BACKSIZE": first["BackSize"],
                        "DOOR_KEY": first["DOOR_KEY"],
                        "year_start": start,
                        "year_end": end,
                    }
                )

    elif mode == "higher":
        high_fallback = build_non_pickup_table_from_atoms(atom_result, "high", keep_internal_columns=True)
        high_fallback = close_same_fact_year_gaps(high_fallback, atom_result)
        fallback_group_keys = [*model_keys, "BACKSIZE"]
        atom_group_keys = [*model_keys, "BackSize"]
        specificity_process_stats: dict[object, dict[str, int]] = {}
        atom_lookup: dict[tuple[object, ...], pd.DataFrame] = {}
        for key_values, group in atom_result.groupby(atom_group_keys, dropna=False, sort=False):
            key_tuple = key_values if isinstance(key_values, tuple) else (key_values,)
            atom_lookup[key_tuple] = group

        for key_values, group in high_fallback.groupby(fallback_group_keys, dropna=False, sort=False):
            key_tuple = key_values if isinstance(key_values, tuple) else (key_values,)
            atoms = atom_lookup.get(key_tuple)
            if atoms is None or atoms.empty:
                records.extend(group.to_dict("records"))
                continue
            merged_group = merge_adjacent_non_pickup_specificity_rows(group, atoms)
            model_name = key_tuple[0]
            model_stats = specificity_process_stats.setdefault(
                model_name,
                {"adjacent_merge_attempts": 0, "adjacent_merge_successes": 0, "adjacent_merge_fallbacks": 0},
            )
            for stat_key, stat_value in merged_group.attrs.get("merge_stats", {}).items():
                model_stats[stat_key] = model_stats.get(stat_key, 0) + int(stat_value)
            records.extend(merged_group.to_dict("records"))

    else:
        raise ValueError(f"Unknown non-pickup compression mode: {mode}")

    if not records:
        return pd.DataFrame(columns=output_columns)

    result = pd.DataFrame(records)
    for _, model_group in result.groupby(model_keys, dropna=False, sort=False):
        raw_consts: set[str] = set()
        for value in model_group["RAW-CONST"]:
            raw_consts.update(split_joined_text(value))
        if len(raw_consts) == 1:
            result.loc[model_group.index, "CONST"] = ""

    result = result.sort_values(
        ["主车型", "BRAND", "MODEL", "year_start", "year_end", "BACKSIZE", "DOOR_KEY", "CONST", "VERSION"],
        kind="mergesort",
    )[internal_columns].reset_index(drop=True)
    if mode == "higher":
        result.attrs["specificity_process_stats"] = specificity_process_stats
    return result


def build_non_pickup_compressed(renamed: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    model_keys = ["主车型", "BRAND", "MODEL"]

    output_columns = NON_PICKUP_FINAL_COLUMNS
    if renamed.empty:
        result = pd.DataFrame(columns=output_columns)
        result.attrs["summary"] = {
            "model_groups": 0,
            "yearcross_true_groups": 0,
            "yearcross_false_groups": 0,
            "special_atomic_rows": 0,
            "special_combined_incl": 0,
            "special_excluded": 0,
            "special_unmatched": 0,
            "base_atomic_rows": 0,
            "output_rows": 0,
            "lossless_rows": 0,
            "high_rows": 0,
            "higher_rows": 0,
            "process_rows": [],
        }
        return result, result.copy(), result.copy()

    work = renamed.copy()
    if "DOOR_KEY" not in work.columns:
        work["DOOR_KEY"] = work["VERSION_RAW"].map(version_door_key)
    for column in [*model_keys, "Const", "VERSION_RAW", "BackSize"]:
        work[column] = work[column].map(normalize_text)
    work["DOOR_KEY"] = work["DOOR_KEY"].map(normalize_text)
    work["YEAR_SINGLE"] = work["YEAR_SINGLE"].astype(int)

    records: list[dict[str, object]] = []
    summary = {
        "model_groups": 0,
        "yearcross_true_groups": 0,
        "yearcross_false_groups": 0,
        "special_atomic_rows": int((work["VERSION_RAW"] != "").sum()),
        "special_combined_incl": 0,
        "special_excluded": 0,
        "special_unmatched": 0,
        "base_atomic_rows": int((work["VERSION_RAW"] == "").sum()),
        "output_rows": 0,
        "lossless_rows": 0,
        "high_rows": 0,
        "higher_rows": 0,
        "process_rows": [],
    }
    for _, group in work.groupby(model_keys, dropna=False, sort=False):
        summary["model_groups"] += 1
        base = group[group["VERSION_RAW"] == ""].copy()
        special = group[group["VERSION_RAW"] != ""].copy()

        base_year_rows: list[dict[str, object]] = []
        if not base.empty:
            year_fact_counts = (
                base.groupby("YEAR_SINGLE", dropna=False)
                .apply(lambda item: item[["Const", "BackSize", "DOOR_KEY"]].drop_duplicates().shape[0])
            )
            year_cross = bool((year_fact_counts > 1).any())
            if year_cross:
                summary["yearcross_true_groups"] += 1
            else:
                summary["yearcross_false_groups"] += 1

            if year_cross:
                for (year, size, door_key), year_size_rows in base.groupby(
                    ["YEAR_SINGLE", "BackSize", "DOOR_KEY"], dropna=False, sort=False
                ):
                    consts = join_unique(year_size_rows["Const"].map(normalize_text).tolist())
                    base_year_rows.append(
                        {
                            **{key: year_size_rows.iloc[0][key] for key in model_keys},
                            "YEAR_SINGLE": int(year),
                            "Const": consts,
                            "RawConst": consts,
                            "BackSize": normalize_text(size),
                            "DOOR_KEY": normalize_text(door_key),
                            "OVERLAP_CONSTS": split_joined_text(consts),
                        }
                    )
            else:
                for (year, size, door_key), year_size_rows in base.groupby(
                    ["YEAR_SINGLE", "BackSize", "DOOR_KEY"], dropna=False, sort=False
                ):
                    raw_consts = join_unique(year_size_rows["Const"].map(normalize_text).tolist())
                    base_year_rows.append(
                        {
                            **{key: year_size_rows.iloc[0][key] for key in model_keys},
                            "YEAR_SINGLE": int(year),
                            "Const": "",
                            "RawConst": raw_consts,
                            "BackSize": normalize_text(size),
                            "DOOR_KEY": normalize_text(door_key),
                            "OVERLAP_CONSTS": set(year_size_rows["Const"].map(normalize_text)),
                        }
                    )

        used_special_indexes: set[int] = set()
        for base_row in base_year_rows:
            same_year_special = special[
                (special["YEAR_SINGLE"] == base_row["YEAR_SINGLE"])
                & (special["DOOR_KEY"].map(normalize_text) == base_row["DOOR_KEY"])
            ]
            same_size_incl_by_door: dict[str, list[tuple[int, str]]] = {}
            different_size_excl_by_door: dict[str, list[str]] = {}
            for special_index, special_row in same_year_special.iterrows():
                special_consts = split_joined_text(special_row["Const"])
                if not const_sets_overlap(base_row["OVERLAP_CONSTS"], special_consts):
                    continue

                special_version = normalize_text(special_row["VERSION_RAW"])
                door_key = normalize_text(special_row["DOOR_KEY"])
                if normalize_text(special_row["BackSize"]) == base_row["BackSize"]:
                    same_size_incl_by_door.setdefault(door_key, []).append((int(special_index), special_version))
                else:
                    different_size_excl_by_door.setdefault(door_key, []).append(special_version)

            door_keys = {
                key
                for key in [*same_size_incl_by_door.keys(), *different_size_excl_by_door.keys()]
                if key
            }
            can_merge_special_versions = len(door_keys) <= 1
            same_size_incl: list[str] = []
            different_size_excl: list[str] = []
            if can_merge_special_versions:
                for items in same_size_incl_by_door.values():
                    for special_index, special_version in items:
                        same_size_incl.append(special_version)
                        used_special_indexes.add(special_index)
                        summary["special_combined_incl"] += 1
                for versions in different_size_excl_by_door.values():
                    different_size_excl.extend(versions)
                    summary["special_excluded"] += len(versions)

            records.append(
                {
                    **{key: base_row[key] for key in model_keys},
                    "YEAR_SINGLE": base_row["YEAR_SINGLE"],
                    "Const": base_row["Const"],
                    "RawConst": base_row.get("RawConst", ""),
                    "VERSION": build_version_text(same_size_incl, different_size_excl),
                    "BackSize": base_row["BackSize"],
                    "DOOR_KEY": base_row["DOOR_KEY"],
                }
            )

        special = special.copy()
        for (year, size, _door_key), special_rows in special.groupby(["YEAR_SINGLE", "BackSize", "DOOR_KEY"], dropna=False, sort=False):
            remaining_rows = special_rows[~special_rows.index.isin(used_special_indexes)]
            if remaining_rows.empty:
                continue

            consts = join_unique(remaining_rows["Const"].map(normalize_text).tolist())
            summary["special_unmatched"] += len(remaining_rows)
            records.append(
                {
                    **{key: remaining_rows.iloc[0][key] for key in model_keys},
                    "YEAR_SINGLE": int(year),
                    "Const": consts,
                    "RawConst": consts,
                    "VERSION": combine_version_tokens(remaining_rows["VERSION_RAW"]),
                    "BackSize": normalize_text(size),
                    "DOOR_KEY": normalize_text(_door_key),
                }
            )

    if not records:
        result = pd.DataFrame(columns=output_columns)
        result.attrs["summary"] = summary
        return result, result.copy(), result.copy()

    atom_result = pd.DataFrame(records)
    lossless = build_non_pickup_table_from_atoms(atom_result, "lossless")
    high = build_non_pickup_table_from_atoms(atom_result, "high")
    higher = build_non_pickup_table_from_atoms(atom_result, "higher")
    summary["lossless_rows"] = len(lossless)
    summary["high_rows"] = len(high)
    summary["higher_rows"] = len(higher)
    summary["output_rows"] = len(lossless)
    raw_counts = renamed.attrs.get("raw_counts", {})
    merge_stats = higher.attrs.get("specificity_process_stats", {})
    process_rows: list[dict[str, object]] = []
    for model_name in sorted(set(work["主车型"].map(normalize_text))):
        model_atoms = work[work["主车型"].map(normalize_text) == model_name]
        model_lossless = lossless[lossless["主车型"].map(normalize_text) == model_name]
        model_high = high[high["主车型"].map(normalize_text) == model_name]
        model_higher = higher[higher["主车型"].map(normalize_text) == model_name]
        model_stats = merge_stats.get(model_name, {})
        fallback_count = int(model_stats.get("adjacent_merge_fallbacks", 0))
        risks = []
        if fallback_count:
            risks.append("存在相邻合并验证失败，已退回保守记录")
        process_rows.append(
            {
                "压缩类型": "非皮卡",
                "主车型": model_name,
                "原始行数": int(raw_counts.get(model_name, 0)),
                "原子事实数": int(len(model_atoms)),
                "无损年份行数": int(len(model_lossless)),
                "无损年份结构行数": int(len(model_high)),
                "特定性行数": int(len(model_higher)),
                "相邻合并尝试次数": int(model_stats.get("adjacent_merge_attempts", 0)),
                "相邻合并成功次数": int(model_stats.get("adjacent_merge_successes", 0)),
                "相邻合并Fallback次数": fallback_count,
                "皮卡年份闭合尝试组数": 0,
                "皮卡年份闭合成功组数": 0,
                "皮卡年份闭合Fallback组数": 0,
                "皮卡Bed合并尝试组数": 0,
                "皮卡Bed合并成功组数": 0,
                "皮卡Bed合并Fallback组数": 0,
                "风险提示": "；".join(risks),
            }
        )
    summary["process_rows"] = process_rows
    lossless.attrs["summary"] = summary
    high.attrs["summary"] = summary
    higher.attrs["summary"] = summary

    return lossless, high, higher


def transform_non_pickup(
    df: pd.DataFrame,
    progress: ProgressReporter | None = None,
    remove_null_size: bool = False,
    field_profile: dict[str, object] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Convert non-pickup rows into lossless and high-compression tables."""
    df = normalize_input_schema(df, field_profile=field_profile)
    if not has_columns(df, NON_PICKUP_REQUIRED_COLUMNS):
        return empty_non_pickup_result()
    require_columns(df, NON_PICKUP_REQUIRED_COLUMNS)

    text_columns = [
        "主车型",
        "品牌",
        "前台车型",
        "开始年",
        "分类",
        "结构",
        "版本",
        "年份区间",
        BACKSIZE_SOURCE_COLUMN,
        "驾驶室类型",
        "货斗长度_ft",
    ]
    work = df.copy()
    for column in text_columns:
        if column not in work.columns:
            work[column] = ""
        work[column] = work[column].map(normalize_text)

    if work["分类"].str.contains("皮卡", na=False).any():
        non_pickup_mask = ~work["分类"].str.contains("皮卡", na=False)
    else:
        non_pickup_mask = (work["驾驶室类型"] == "") & (work["货斗长度_ft"] == "")

    size_mask = work[BACKSIZE_SOURCE_COLUMN] != ""
    if remove_null_size:
        size_mask &= work[BACKSIZE_SOURCE_COLUMN] != "无可用尺码"
    work = work[non_pickup_mask & size_mask].copy()
    if (work["年份区间"] == "").any():
        bad_count = int((work["年份区间"] == "").sum())
        raise ValueError(f"非皮卡存在 {bad_count} 行年份区间为空，请补齐年份区间后再压缩。")

    if work.empty:
        return empty_non_pickup_result()

    work["MODEL_LIST"] = work["前台车型"].map(split_front_model_atoms)
    work = work.explode("MODEL_LIST")
    work["前台车型"] = work["MODEL_LIST"].map(normalize_text)
    work = work.drop(columns=["MODEL_LIST"])
    work["START_YEAR"] = work.apply(lambda row: resolve_start_year(row.get("开始年", ""), row.get("年份区间", "")), axis=1)
    if work["START_YEAR"].isna().any():
        bad_count = int(work["START_YEAR"].isna().sum())
        raise ValueError(f"非皮卡存在 {bad_count} 行无法解析开始年，请检查开始年或年份区间。")
    work["START_YEAR"] = work["START_YEAR"].astype(int)

    raw_counts = work.groupby(["品牌", "前台车型"], dropna=False).size().to_dict()
    work["YEAR_LIST"] = work["年份区间"].map(parse_year_list)
    work = work.explode("YEAR_LIST")
    work = work[work["YEAR_LIST"].notna()].copy()
    work["YEAR_SINGLE"] = work["YEAR_LIST"].astype(int)

    work["CONST_LIST"] = work["结构"].map(split_const_atoms)
    work = work.explode("CONST_LIST")
    work["结构"] = work["CONST_LIST"].map(normalize_text)
    work = work.drop(columns=["CONST_LIST"])

    renamed = work.rename(
        columns={
            "品牌": "BRAND",
            "前台车型": "MODEL",
            "结构": "Const",
            "版本": "VERSION_RAW",
            BACKSIZE_SOURCE_COLUMN: "BackSize",
        }
    )
    renamed = drop_duplicate_columns(renamed)

    renamed = renamed[
        ["BRAND", "MODEL", "START_YEAR", "YEAR_SINGLE", "Const", "VERSION_RAW", "BackSize"]
    ].drop_duplicates()
    renamed.attrs["raw_counts"] = {
        (normalize_text(key[0]), normalize_text(key[1])): int(value)
        for key, value in raw_counts.items()
    }

    for column in ["BRAND", "MODEL", "Const", "VERSION_RAW", "BackSize"]:
        renamed[column] = renamed[column].map(normalize_text)
    renamed = renamed.sort_values(["BRAND", "MODEL", "START_YEAR", "YEAR_SINGLE", "BackSize", "Const", "VERSION_RAW"], kind="mergesort").reset_index(drop=True)

    lossless_internal = build_non_pickup_lossless_new(renamed)
    high_internal = lossless_internal.copy()
    higher_internal, merge_stats, compression_log = build_non_pickup_high_new(lossless_internal, renamed, progress=progress)
    atom_table = build_non_pickup_atom_table(renamed)
    process_rows: list[dict[str, object]] = []
    for (brand, model), atoms_group in renamed.groupby(["BRAND", "MODEL"], dropna=False, sort=False):
        key = (normalize_text(brand), normalize_text(model))
        lossless_group = lossless_internal[
            (lossless_internal["BRAND"].map(normalize_text) == key[0])
            & (lossless_internal["MODEL"].map(normalize_text) == key[1])
        ]
        higher_group = higher_internal[
            (higher_internal["BRAND"].map(normalize_text) == key[0])
            & (higher_internal["MODEL"].map(normalize_text) == key[1])
        ]
        stats = merge_stats.get(key, {})
        fallback_count = int(stats.get("fallbacks", 0))
        risks = ["存在组合验证失败，已fallback"] if fallback_count else []
        process_rows.append(
            {
                "压缩类型": "非皮卡",
                "BRAND": key[0],
                "MODEL": key[1],
                "原始行数": int(renamed.attrs.get("raw_counts", {}).get(key, 0)),
                "原子事实数": int(len(atoms_group)),
                "无损年份行数": int(len(lossless_group)),
                "无损年份结构行数": "",
                "特定性行数": int(len(higher_group)),
                "相邻合并尝试次数": int(stats.get("attempts", 0)),
                "相邻合并成功次数": int(stats.get("successes", 0)),
                "相邻合并Fallback次数": fallback_count,
                "皮卡年份闭合尝试组数": 0,
                "皮卡年份闭合成功组数": 0,
                "皮卡年份闭合Fallback组数": 0,
                "皮卡Bed合并尝试组数": 0,
                "皮卡Bed合并成功组数": 0,
                "皮卡Bed合并Fallback组数": 0,
                "风险提示": "；".join(risks),
            }
        )

    summary = {
        "model_groups": int(renamed[["BRAND", "MODEL"]].drop_duplicates().shape[0]),
        "yearcross_true_groups": 0,
        "yearcross_false_groups": 0,
        "special_atomic_rows": int((renamed["VERSION_RAW"] != "").sum()),
        "special_combined_incl": 0,
        "special_excluded": 0,
        "special_unmatched": 0,
        "base_atomic_rows": int((renamed["VERSION_RAW"] == "").sum()),
        "output_rows": len(lossless_internal),
        "lossless_rows": len(lossless_internal),
        "high_rows": len(lossless_internal),
        "higher_rows": len(higher_internal),
        "process_rows": process_rows,
    }

    lossless = lossless_internal[NON_PICKUP_FINAL_COLUMNS].reset_index(drop=True)
    high = high_internal[NON_PICKUP_FINAL_COLUMNS].reset_index(drop=True)
    higher = higher_internal[NON_PICKUP_FINAL_COLUMNS].reset_index(drop=True)
    for table in [lossless, high, higher]:
        table.attrs["summary"] = summary
        table.attrs["atom_table"] = atom_table
        table.attrs["compression_log"] = compression_log
    return lossless, high, higher


def build_pickup_lossless_compressed(renamed: pd.DataFrame) -> pd.DataFrame:
    output_columns = PICKUP_FINAL_COLUMNS
    if renamed.empty:
        return pd.DataFrame(columns=output_columns)

    rows: list[dict[str, object]] = []
    group_keys = ["主车型", "BRAND", "MODEL", "VERSION_RAW", "CAB", "BED_FT", "BackSize"]
    for key_values, group in renamed.groupby(group_keys, dropna=False, sort=False):
        values = dict(zip(group_keys, key_values if isinstance(key_values, tuple) else (key_values,)))
        years = sorted(set(group["YEAR_SINGLE"].astype(int)))
        for start, end in continuous_year_segments(years):
            rows.append(
                {
                    "主车型": values["主车型"],
                    "BRAND": values["BRAND"],
                    "MODEL": values["MODEL"],
                    "YEAR": format_year_range([start, end]),
                    "VERSION": values["VERSION_RAW"],
                    "CAB": values["CAB"],
                    "BED_FT": values["BED_FT"],
                    "BACKSIZE": values["BackSize"],
                    "_year_start": start,
                    "_year_end": end,
                }
            )

    result = pd.DataFrame(rows)
    if result.empty:
        return pd.DataFrame(columns=output_columns)

    result = result.sort_values(
        ["主车型", "BRAND", "MODEL", "_year_start", "_year_end", "VERSION", "CAB", "BED_FT", "BACKSIZE"],
        kind="mergesort",
    )
    result = result[output_columns].reset_index(drop=True)
    return result


def pickup_record_matches_atom_scope(record: pd.Series, atom: pd.Series) -> bool:
    if normalize_text(record.get("BRAND", "")) != normalize_text(atom.get("BRAND", "")):
        return False
    if not model_expression_matches(record.get("MODEL", ""), atom.get("MODEL", "")):
        return False
    if not record_version_matches(record.get("VERSION", ""), atom.get("VERSION_RAW", "")):
        return False
    cab_text = normalize_text(record.get("CAB", ""))
    atom_cab = normalize_text(atom.get("CAB", ""))
    if cab_text and atom_cab not in split_joined_text(cab_text):
        return False
    if not cab_text and atom_cab:
        return False
    if int(atom["YEAR_SINGLE"]) not in parse_year_list(record.get("YEAR", "")):
        return False
    return bed_matches_expression(atom.get("BED_FT", ""), record.get("BED_FT", ""))


def pickup_record_covers_atom(record: pd.Series, atom: pd.Series) -> bool:
    return pickup_record_matches_atom_scope(record, atom) and normalize_text(record.get("BACKSIZE", "")) == normalize_text(
        atom.get("BackSize", "")
    )


def pickup_candidate_validation_reason(rows: pd.DataFrame, atoms: pd.DataFrame) -> str:
    for _, atom in atoms.iterrows():
        matched_sizes: list[str] = []
        for _, row in rows.iterrows():
            if pickup_record_matches_atom_scope(row, atom):
                matched_sizes.append(normalize_text(row.get("BACKSIZE", "")))
        if len(matched_sizes) > 1:
            return "原子事实对应多条候选记录"
        if not matched_sizes:
            return "原子事实未被候选记录覆盖"
        if matched_sizes[0] != normalize_text(atom.get("BackSize", "")):
            return "原子事实命中不同尺码候选记录"
    return ""


def pickup_internal_from_lossless(lossless: pd.DataFrame) -> pd.DataFrame:
    if lossless.empty:
        return pd.DataFrame(columns=[*PICKUP_FINAL_COLUMNS, "VERSION_RAW", "BackSize", "_year_start", "_year_end"])

    rows = []
    for _, row in lossless.iterrows():
        years = parse_year_list(row.get("YEAR", ""))
        if not years:
            continue
        item = row.to_dict()
        item["VERSION_RAW"] = normalize_text(row.get("VERSION", ""))
        item["BackSize"] = normalize_text(row.get("BACKSIZE", ""))
        item["_year_start"] = min(years)
        item["_year_end"] = max(years)
        rows.append(item)
    return pd.DataFrame(rows)


def merge_pickup_high_records(left: pd.Series, right: pd.Series) -> dict[str, object]:
    start = min(int(left["_year_start"]), int(right["_year_start"]))
    end = max(int(left["_year_end"]), int(right["_year_end"]))
    bed_ft = combine_bed_ft(pd.Series([left.get("BED_FT", ""), right.get("BED_FT", "")]))
    model = combine_model_expression_for_brand(left["BRAND"], [left["MODEL"], right["MODEL"]])
    return {
        "主车型": " ".join(part for part in [left["BRAND"], model] if part),
        "BRAND": left["BRAND"],
        "MODEL": model,
        "YEAR": format_year_range([start, end]),
        "VERSION": combine_lossy_versions([left.get("VERSION", ""), right.get("VERSION", "")]),
        "VERSION_RAW": combine_lossy_versions([left.get("VERSION", ""), right.get("VERSION", "")]),
        "CAB": join_unique([left.get("CAB", ""), right.get("CAB", "")]),
        "BED_FT": bed_ft,
        "BACKSIZE": left["BACKSIZE"],
        "BackSize": left["BACKSIZE"],
        "_year_start": start,
        "_year_end": end,
    }


def build_pickup_high_from_lossless(
    lossless: pd.DataFrame,
    atoms: pd.DataFrame,
    progress: ProgressReporter | None = None,
) -> pd.DataFrame:
    output_columns = PICKUP_FINAL_COLUMNS
    process_stats: dict[object, dict[str, int]] = {}
    log_rows: list[dict[str, object]] = []
    if lossless.empty:
        result = pd.DataFrame(columns=output_columns)
        result.attrs["summary"] = {
            "pickup_bed_merged_year_groups": 0,
            "pickup_gap_closed_groups": 0,
            "pickup_cab_bed_closed_groups": 0,
            "process_stats": process_stats,
            "compression_log": empty_log_table(),
        }
        return result

    output_groups: list[pd.DataFrame] = []
    work_source = pickup_internal_from_lossless(lossless)
    work_source["__MODEL_GROUP"] = work_source.apply(lambda row: model_combo_group(row["BRAND"], row["MODEL"]), axis=1)
    atoms_work = atoms.copy()
    atoms_work["__MODEL_GROUP"] = atoms_work.apply(lambda row: model_combo_group(row["BRAND"], row["MODEL"]), axis=1)
    total_groups = int(work_source.groupby(["BRAND", "__MODEL_GROUP"], dropna=False, sort=False).ngroups)
    completed_groups = 0
    merge_count = 0
    attempt_count = 0
    if progress:
        progress.start("皮卡高度压缩", total_groups)
    for group_values, group in work_source.groupby(["BRAND", "__MODEL_GROUP"], dropna=False, sort=False):
        brand, model_group = group_values
        group = group.drop(columns=["__MODEL_GROUP"])
        model_stats = process_stats.setdefault(
            f"{normalize_text(brand)} {normalize_text(model_group)}",
            {"year_gap_attempts": 0, "year_gap_successes": 0, "year_gap_fallbacks": 0, "bed_merge_attempts": 0, "bed_merge_successes": 0, "bed_merge_fallbacks": 0},
        )
        if progress:
            progress.update(current_make=brand, current_model=model_group, completed_models=completed_groups)
        atoms_group = atoms_work[
            (atoms_work["BRAND"].map(normalize_text) == normalize_text(brand))
            & (atoms_work["__MODEL_GROUP"].map(normalize_text) == normalize_text(model_group))
        ].drop(columns=["__MODEL_GROUP"])
        working = group.sort_values(["_year_start", "_year_end", "VERSION", "CAB", "BED_FT", "BACKSIZE"], kind="mergesort").reset_index(drop=True)

        changed = True
        while changed and len(working) > 1:
            changed = False
            for left_index in range(len(working)):
                if changed:
                    break
                for right_index in range(left_index + 1, len(working)):
                    left = working.iloc[left_index]
                    right = working.iloc[right_index]
                    if normalize_text(left.get("BACKSIZE", "")) != normalize_text(right.get("BACKSIZE", "")):
                        continue
                    model_stats["bed_merge_attempts"] += 1
                    attempt_count += 1
                    merged = merge_pickup_high_records(left, right)
                    log_base = {
                        "压缩类型": "皮卡",
                        "阶段": "高度压缩两两组合",
                        "BRAND": normalize_text(brand),
                        "MODEL": pair_text(left.get("MODEL", ""), right.get("MODEL", "")),
                        "合并MODEL": normalize_text(merged["MODEL"]),
                        "BACKSIZE": normalize_text(merged["BACKSIZE"]),
                        **candidate_log_fields(left, right, merged, include_const=False),
                    }

                    candidate_rows = [
                        *working.iloc[:left_index].to_dict("records"),
                        *working.iloc[left_index + 1 : right_index].to_dict("records"),
                        *working.iloc[right_index + 1 :].to_dict("records"),
                        merged,
                    ]
                    candidate = pd.DataFrame(candidate_rows, columns=working.columns)
                    reason = pickup_candidate_validation_reason(candidate, atoms_group)
                    if not reason:
                        model_stats["bed_merge_successes"] += 1
                        merge_count += 1
                        log_rows.append({**log_base, "结果": "success", "原因": ""})
                        if progress:
                            progress.update(merge_count=merge_count, attempt_count=attempt_count)
                        working = candidate.sort_values(
                            ["_year_start", "_year_end", "VERSION", "CAB", "BED_FT", "BACKSIZE"], kind="mergesort"
                        ).reset_index(drop=True)
                        changed = True
                        break
                    model_stats["bed_merge_fallbacks"] += 1
                    log_rows.append({**log_base, "结果": "fallback", "原因": reason})
                    if progress:
                        progress.update(merge_count=merge_count, attempt_count=attempt_count)

        output_groups.append(working)
        completed_groups += 1
        if progress:
            progress.update(completed_models=completed_groups, merge_count=merge_count, attempt_count=attempt_count)

    if not output_groups:
        result = pd.DataFrame(columns=output_columns)
    else:
        result = pd.concat(output_groups, ignore_index=True).sort_values(
            ["主车型", "BRAND", "MODEL", "_year_start", "_year_end", "VERSION", "CAB", "BED_FT", "BACKSIZE"],
            kind="mergesort",
        )
        result = result[output_columns].reset_index(drop=True)
    if progress:
        progress.finish()

    result.attrs["summary"] = {
        "pickup_bed_merged_year_groups": sum(stats["bed_merge_successes"] for stats in process_stats.values()),
        "pickup_gap_closed_groups": 0,
        "pickup_cab_bed_closed_groups": sum(stats["bed_merge_successes"] for stats in process_stats.values()),
        "process_stats": process_stats,
        "compression_log": pd.DataFrame(log_rows, columns=LOG_FINAL_COLUMNS),
    }
    return result


def pickup_candidate_has_size_conflict(
    facts: pd.DataFrame,
    values: dict[str, object],
    start: int,
    end: int,
    bed_expression: object,
) -> bool:
    scoped = facts[
        (facts["主车型"] == values["主车型"])
        & (facts["BRAND"] == values["BRAND"])
        & (facts["MODEL"] == values["MODEL"])
        & (facts["VERSION_RAW"] == values["VERSION_RAW"])
        & (facts["CAB"] == values["CAB"])
        & (facts["YEAR_SINGLE"].astype(int).between(start, end))
    ]
    for _, fact in scoped.iterrows():
        if bed_matches_expression(fact["BED_FT"], bed_expression) and fact["BackSize"] != values["BackSize"]:
            return True
    return False


def build_pickup_output_row(
    values: dict[str, object],
    start: int,
    end: int,
    bed_ft: object,
) -> dict[str, object]:
    return {
        "主车型": values["主车型"],
        "BRAND": values["BRAND"],
        "MODEL": values["MODEL"],
        "YEAR": format_year_range([start, end]),
        "VERSION": values["VERSION_RAW"],
        "VERSION_RAW": values["VERSION_RAW"],
        "CAB": values["CAB"],
        "BED_FT": normalize_text(bed_ft),
        "BACKSIZE": values["BackSize"],
        "BackSize": values["BackSize"],
        "_year_start": start,
        "_year_end": end,
    }


def build_pickup_specificity_compressed(renamed: pd.DataFrame) -> pd.DataFrame:
    output_columns = PICKUP_FINAL_COLUMNS
    if renamed.empty:
        result = pd.DataFrame(columns=output_columns)
        result.attrs["summary"] = {
            "pickup_bed_merged_year_groups": 0,
            "pickup_gap_closed_groups": 0,
            "pickup_cab_bed_closed_groups": 0,
            "process_stats": {},
        }
        return result

    model_keys = ["主车型", "BRAND", "MODEL", "VERSION_RAW"]
    bed_rows: list[dict[str, object]] = []
    bed_keys = [*model_keys, "CAB", "BED_FT", "BackSize"]
    gap_closed_groups = 0
    process_stats: dict[object, dict[str, int]] = {}
    log_rows: list[dict[str, object]] = []
    for key_values, group in renamed.groupby(bed_keys, dropna=False, sort=False):
        values = dict(zip(bed_keys, key_values if isinstance(key_values, tuple) else (key_values,)))
        model_stats = process_stats.setdefault(
            values["主车型"],
            {"year_gap_attempts": 0, "year_gap_successes": 0, "year_gap_fallbacks": 0, "bed_merge_attempts": 0, "bed_merge_successes": 0, "bed_merge_fallbacks": 0},
        )
        years = sorted(set(group["YEAR_SINGLE"].astype(int)))
        if not years:
            continue

        start = min(years)
        end = max(years)
        if len(continuous_year_segments(years)) > 1:
            model_stats["year_gap_attempts"] += 1
            has_conflict = pickup_candidate_has_size_conflict(renamed, values, start, end, values["BED_FT"])
            log_base = {
                "压缩类型": "皮卡",
                "阶段": "年份空洞闭合",
                "BRAND": values["BRAND"],
                "MODEL": values["MODEL"],
                "合并MODEL": values["MODEL"],
                "CAB": values["CAB"],
                "BED_FT": values["BED_FT"],
                "BACKSIZE": values["BackSize"],
                "候选YEAR": format_year_range([start, end]),
                "候选CONST": "",
                "候选VERSION": values["VERSION_RAW"],
                "合并CAB": values["CAB"],
                "合并BED_FT": values["BED_FT"],
                "合并YEAR": format_year_range([start, end]),
                "合并CONST": "",
                "合并VERSION": values["VERSION_RAW"],
            }
            if not has_conflict:
                bed_rows.append(build_pickup_output_row(values, start, end, values["BED_FT"]))
                gap_closed_groups += 1
                model_stats["year_gap_successes"] += 1
                log_rows.append({**log_base, "结果": "success", "原因": ""})
                continue
            model_stats["year_gap_fallbacks"] += 1
            log_rows.append({**log_base, "结果": "fallback", "原因": "候选年份区间内存在同BED不同尺码事实"})

        for segment_start, segment_end in continuous_year_segments(years):
            bed_rows.append(build_pickup_output_row(values, segment_start, segment_end, values["BED_FT"]))

    if not bed_rows:
        result = pd.DataFrame(columns=output_columns)
        result.attrs["summary"] = {
            "pickup_bed_merged_year_groups": 0,
            "pickup_gap_closed_groups": gap_closed_groups,
            "pickup_cab_bed_closed_groups": 0,
            "process_stats": process_stats,
        }
        return result

    bed_df = pd.DataFrame(bed_rows)
    rows: list[dict[str, object]] = []
    cab_keys = [*model_keys, "CAB", "BackSize"]
    bed_merged_year_groups = 0
    cab_bed_closed_groups = 0
    for key_values, group in bed_df.groupby(cab_keys, dropna=False, sort=False):
        values = dict(zip(cab_keys, key_values if isinstance(key_values, tuple) else (key_values,)))
        start = int(group["_year_start"].min())
        end = int(group["_year_end"].max())
        bed_ft = combine_bed_ft(group["BED_FT"])
        bed_variants = len(set(normalize_text(value) for value in group["BED_FT"] if normalize_text(value)))

        if bed_variants > 1:
            process_stats.setdefault(
                values["主车型"],
                {"year_gap_attempts": 0, "year_gap_successes": 0, "year_gap_fallbacks": 0, "bed_merge_attempts": 0, "bed_merge_successes": 0, "bed_merge_fallbacks": 0},
            )["bed_merge_attempts"] += 1
            has_conflict = pickup_candidate_has_size_conflict(renamed, values, start, end, bed_ft)
            log_base = {
                "压缩类型": "皮卡",
                "阶段": "BED合并",
                "BRAND": values["BRAND"],
                "MODEL": values["MODEL"],
                "合并MODEL": values["MODEL"],
                "CAB": values["CAB"],
                "BED_FT": bed_ft,
                "BACKSIZE": values["BackSize"],
                "候选YEAR": format_year_range([start, end]),
                "候选CONST": "",
                "候选VERSION": values["VERSION_RAW"],
                "合并CAB": values["CAB"],
                "合并BED_FT": bed_ft,
                "合并YEAR": format_year_range([start, end]),
                "合并CONST": "",
                "合并VERSION": values["VERSION_RAW"],
            }
        if bed_variants > 1 and not has_conflict:
            rows.append(build_pickup_output_row(values, start, end, bed_ft))
            bed_merged_year_groups += int(
                group.groupby(["_year_start", "_year_end"], dropna=False)["BED_FT"].nunique().gt(1).sum()
            )
            cab_bed_closed_groups += 1
            process_stats[values["主车型"]]["bed_merge_successes"] += 1
            log_rows.append({**log_base, "结果": "success", "原因": ""})
            continue
        if bed_variants > 1:
            model_stats = process_stats.setdefault(
                values["主车型"],
                {"year_gap_attempts": 0, "year_gap_successes": 0, "year_gap_fallbacks": 0, "bed_merge_attempts": 0, "bed_merge_successes": 0, "bed_merge_fallbacks": 0},
            )
            model_stats["bed_merge_fallbacks"] += 1
            log_rows.append({**log_base, "结果": "fallback", "原因": "候选BED范围内存在不同尺码事实"})

        compress_keys = [*cab_keys, "BED_FT"]
        for fallback_key_values, fallback_group in group.groupby(compress_keys, dropna=False, sort=False):
            fallback_values = dict(
                zip(compress_keys, fallback_key_values if isinstance(fallback_key_values, tuple) else (fallback_key_values,))
            )
            years: set[int] = set()
            for _, row in fallback_group.iterrows():
                years.update(range(int(row["_year_start"]), int(row["_year_end"]) + 1))
            for segment_start, segment_end in continuous_year_segments(sorted(years)):
                rows.append(build_pickup_output_row(fallback_values, segment_start, segment_end, fallback_values["BED_FT"]))

    result = pd.DataFrame(rows)
    if result.empty:
        result = pd.DataFrame(columns=output_columns)
    else:
        result = result.sort_values(
            ["主车型", "BRAND", "MODEL", "_year_start", "_year_end", "VERSION", "CAB", "BED_FT", "BACKSIZE"],
            kind="mergesort",
        )
        result = result[output_columns].reset_index(drop=True)

    result.attrs["summary"] = {
        "pickup_bed_merged_year_groups": bed_merged_year_groups,
        "pickup_gap_closed_groups": gap_closed_groups,
        "pickup_cab_bed_closed_groups": cab_bed_closed_groups,
        "process_stats": process_stats,
        "compression_log": pd.DataFrame(log_rows, columns=LOG_FINAL_COLUMNS),
    }
    return result


def transform_pickup(
    df: pd.DataFrame,
    progress: ProgressReporter | None = None,
    remove_null_size: bool = False,
    field_profile: dict[str, object] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Convert pickup rows into lossless and specificity-preserving compressed tables."""
    df = normalize_input_schema(df, field_profile=field_profile)
    if not has_columns(df, PICKUP_REQUIRED_COLUMNS):
        return empty_pickup_result()
    require_columns(df, PICKUP_REQUIRED_COLUMNS)

    text_columns = [
        "主车型",
        "分类",
        "品牌",
        "前台车型",
        "版本",
        "年份区间",
        "驾驶室类型",
        "货斗长度_ft",
        BACKSIZE_SOURCE_COLUMN,
    ]
    work = df.copy()
    for column in text_columns:
        if column not in work.columns:
            work[column] = ""
        work[column] = work[column].map(normalize_text)

    if work["分类"].str.contains("皮卡", na=False).any():
        pickup_mask = work["分类"].str.contains("皮卡", na=False)
    else:
        pickup_mask = (work["驾驶室类型"] != "") | (work["货斗长度_ft"] != "")

    size_mask = work[BACKSIZE_SOURCE_COLUMN] != ""
    if remove_null_size:
        size_mask &= work[BACKSIZE_SOURCE_COLUMN] != "无可用尺码"
    work = work[pickup_mask & size_mask].copy()
    if (work["年份区间"] == "").any():
        bad_count = int((work["年份区间"] == "").sum())
        raise ValueError(f"皮卡存在 {bad_count} 行年份区间为空，请补齐年份区间后再压缩。")

    if work.empty:
        return empty_pickup_result()

    raw_counts = work.groupby("主车型", dropna=False).size().to_dict()
    work["YEAR_LIST"] = work["年份区间"].map(parse_year_list)
    work = work.explode("YEAR_LIST")
    work = work[work["YEAR_LIST"].notna()].copy()
    work["YEAR_SINGLE"] = work["YEAR_LIST"].astype(int)

    renamed = work.rename(
        columns={
            "品牌": "BRAND",
            "前台车型": "MODEL",
            "版本": "VERSION_RAW",
            "驾驶室类型": "CAB",
            "货斗长度_ft": "BED_FT",
            BACKSIZE_SOURCE_COLUMN: "BackSize",
        }
    )
    renamed = drop_duplicate_columns(renamed)

    for column in [
        "主车型",
        "BRAND",
        "MODEL",
        "VERSION_RAW",
        "CAB",
        "BED_FT",
        "BackSize",
    ]:
        renamed[column] = renamed[column].map(normalize_text)

    renamed["CAB"] = renamed["CAB"].map(split_joined_atoms)
    renamed = renamed.explode("CAB")

    renamed = renamed[
        ["主车型", "BRAND", "MODEL", "YEAR_SINGLE", "VERSION_RAW", "CAB", "BED_FT", "BackSize"]
    ].drop_duplicates()
    atom_table = build_pickup_atom_table(renamed)

    pickup_lossless = build_pickup_lossless_compressed(renamed)
    pickup_specificity = build_pickup_high_from_lossless(pickup_lossless, renamed, progress=progress)
    summary = {
        "pickup_atomic_rows": len(renamed),
        "pickup_bed_merged_year_groups": pickup_specificity.attrs.get("summary", {}).get("pickup_bed_merged_year_groups", 0),
        "pickup_gap_closed_groups": pickup_specificity.attrs.get("summary", {}).get("pickup_gap_closed_groups", 0),
        "pickup_cab_bed_closed_groups": pickup_specificity.attrs.get("summary", {}).get("pickup_cab_bed_closed_groups", 0),
        "compression_log": pickup_specificity.attrs.get("summary", {}).get("compression_log", empty_log_table()),
    }
    process_stats = pickup_specificity.attrs.get("summary", {}).get("process_stats", {})
    process_rows = []
    for model_name in sorted(set(renamed["主车型"].map(normalize_text))):
        model_lossless = pickup_lossless[pickup_lossless["主车型"].map(normalize_text) == model_name]
        model_specificity = pickup_specificity[pickup_specificity["主车型"].map(normalize_text) == model_name]
        stats = process_stats.get(model_name, {})
        fallback_count = int(stats.get("year_gap_fallbacks", 0)) + int(stats.get("bed_merge_fallbacks", 0))
        risks = []
        if fallback_count:
            risks.append("存在皮卡闭合或Bed合并验证失败，已退回保守记录")
        process_rows.append(
            {
                "压缩类型": "皮卡",
                "BRAND": normalize_text(model_lossless.iloc[0]["BRAND"]) if not model_lossless.empty else "",
                "MODEL": normalize_text(model_lossless.iloc[0]["MODEL"]) if not model_lossless.empty else "",
                "原始行数": int(raw_counts.get(model_name, 0)),
                "原子事实数": int(len(renamed[renamed["主车型"].map(normalize_text) == model_name])),
                "无损年份行数": int(len(model_lossless)),
                "无损年份结构行数": "",
                "特定性行数": int(len(model_specificity)),
                "相邻合并尝试次数": 0,
                "相邻合并成功次数": 0,
                "相邻合并Fallback次数": 0,
                "皮卡年份闭合尝试组数": int(stats.get("year_gap_attempts", 0)),
                "皮卡年份闭合成功组数": int(stats.get("year_gap_successes", 0)),
                "皮卡年份闭合Fallback组数": int(stats.get("year_gap_fallbacks", 0)),
                "皮卡Bed合并尝试组数": int(stats.get("bed_merge_attempts", 0)),
                "皮卡Bed合并成功组数": int(stats.get("bed_merge_successes", 0)),
                "皮卡Bed合并Fallback组数": int(stats.get("bed_merge_fallbacks", 0)),
                "风险提示": "；".join(risks),
            }
        )
    summary["process_rows"] = process_rows
    pickup_lossless.attrs["summary"] = summary
    pickup_specificity.attrs["summary"] = summary
    pickup_lossless.attrs["atom_table"] = atom_table
    pickup_specificity.attrs["atom_table"] = atom_table
    pickup_lossless.attrs["compression_log"] = summary["compression_log"]
    pickup_specificity.attrs["compression_log"] = summary["compression_log"]
    return pickup_lossless, pickup_specificity


def transform(df: pd.DataFrame) -> pd.DataFrame:
    """Backward-compatible alias for the non-pickup table."""
    return transform_non_pickup(df)[0]


def transform_all(
    df: pd.DataFrame,
    progress: ProgressReporter | None = None,
    remove_null_size: bool = False,
    field_profile: dict[str, object] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    non_pickup_lossless, non_pickup_high, non_pickup_higher = transform_non_pickup(
        df,
        progress=progress,
        remove_null_size=remove_null_size,
        field_profile=field_profile,
    )
    pickup_lossless, pickup_specificity = transform_pickup(
        df,
        progress=progress,
        remove_null_size=remove_null_size,
        field_profile=field_profile,
    )
    return non_pickup_lossless, non_pickup_high, non_pickup_higher, pickup_lossless, pickup_specificity


def build_process_table(non_pickup_lossless: pd.DataFrame, pickup_lossless: pd.DataFrame) -> pd.DataFrame:
    rows = [
        *non_pickup_lossless.attrs.get("summary", {}).get("process_rows", []),
        *pickup_lossless.attrs.get("summary", {}).get("process_rows", []),
    ]
    if not rows:
        return pd.DataFrame(columns=PROCESS_FINAL_COLUMNS)
    return pd.DataFrame(rows)[PROCESS_FINAL_COLUMNS].sort_values(["压缩类型", "BRAND", "MODEL"], kind="mergesort").reset_index(drop=True)


def build_atom_table(non_pickup_lossless: pd.DataFrame, pickup_lossless: pd.DataFrame) -> pd.DataFrame:
    tables = [
        non_pickup_lossless.attrs.get("atom_table", empty_atom_table()),
        pickup_lossless.attrs.get("atom_table", empty_atom_table()),
    ]
    non_empty = [table for table in tables if not table.empty]
    if not non_empty:
        return empty_atom_table()
    return pd.concat(non_empty, ignore_index=True)[ATOM_FINAL_COLUMNS].sort_values(
        ["压缩类型", "BRAND", "MODEL", "YEAR", "BACKSIZE"], kind="mergesort"
    ).reset_index(drop=True)


def build_compression_log(non_pickup_higher: pd.DataFrame, pickup_specificity: pd.DataFrame) -> pd.DataFrame:
    tables = [
        non_pickup_higher.attrs.get("compression_log", empty_log_table()),
        pickup_specificity.attrs.get("compression_log", empty_log_table()),
    ]
    non_empty = [table for table in tables if not table.empty]
    if not non_empty:
        return empty_log_table()
    return pd.concat(non_empty, ignore_index=True)[LOG_FINAL_COLUMNS].reset_index(drop=True)


def transform_all_outputs(
    df: pd.DataFrame,
    remove_null_size: bool = False,
    progress: ProgressReporter | None = None,
    field_profile: dict[str, object] | None = None,
) -> tuple[
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame,
]:
    non_pickup_lossless, non_pickup_high, non_pickup_higher, pickup_lossless, pickup_specificity = transform_all(
        df,
        progress=progress,
        remove_null_size=remove_null_size,
        field_profile=field_profile,
    )
    log_df = build_compression_log(non_pickup_higher, pickup_specificity)
    atom_df = build_atom_table(non_pickup_lossless, pickup_lossless)
    return (
        non_pickup_lossless,
        non_pickup_high,
        non_pickup_higher,
        pickup_lossless,
        pickup_specificity,
        log_df,
        atom_df,
    )


def build_summary_markdown(
    input_path: Path,
    non_pickup_lossless_df: pd.DataFrame,
    non_pickup_high_df: pd.DataFrame,
    non_pickup_higher_df: pd.DataFrame,
    pickup_lossless_df: pd.DataFrame,
    pickup_specificity_df: pd.DataFrame,
    log_df: pd.DataFrame,
    atom_df: pd.DataFrame,
) -> str:
    summary = non_pickup_lossless_df.attrs.get("summary", {})
    pickup_summary = pickup_lossless_df.attrs.get("summary", {})
    lines = [
        "# 尺码压缩处理总结",
        "",
        f"- 输入文件: {input_path.name}",
        f"- 非皮卡无损压缩表行数: {len(non_pickup_lossless_df)}",
        f"- 非皮卡高度压缩行数: {len(non_pickup_higher_df)}",
        f"- 皮卡无损压缩行数: {len(pickup_lossless_df)}",
        f"- 皮卡特定性压缩行数: {len(pickup_specificity_df)}",
        f"- 压缩log行数: {len(log_df)}",
        f"- 原子事实表行数: {len(atom_df)}",
        "",
        "## 非皮卡压缩动作",
        "",
        f"- BRAND + MODEL 分组数: {summary.get('model_groups', 0)}",
        f"- 基础版本空原子记录数: {summary.get('base_atomic_rows', 0)}",
        f"- 特殊版本非空原子记录数: {summary.get('special_atomic_rows', 0)}",
        f"- 特殊版本结合进基础记录为 Incl 的次数: {summary.get('special_combined_incl', 0)}",
        f"- 特殊版本导致基础记录增加 Excl 的次数: {summary.get('special_excluded', 0)}",
        f"- 特殊版本未匹配基础记录、单独输出的次数: {summary.get('special_unmatched', 0)}",
        f"- 高度压缩减少行数: {len(non_pickup_lossless_df) - len(non_pickup_higher_df)}",
        "",
        "## 皮卡压缩动作",
        "",
        f"- 皮卡原子记录数: {pickup_summary.get('pickup_atomic_rows', 0)}",
        f"- 皮卡无损压缩减少行数: {pickup_summary.get('pickup_atomic_rows', 0) - len(pickup_lossless_df)}",
        f"- 皮卡特定性压缩减少行数: {len(pickup_lossless_df) - len(pickup_specificity_df)}",
        f"- 同年同 Cab 同版本同尺码合并 Bed 的组数: {pickup_summary.get('pickup_bed_merged_year_groups', 0)}",
        f"- 同 Cab 同 Bed 同版本同尺码闭合年份空洞的组数: {pickup_summary.get('pickup_gap_closed_groups', 0)}",
        f"- 同 Cab 同版本同尺码合并 Bed 并闭合年份的组数: {pickup_summary.get('pickup_cab_bed_closed_groups', 0)}",
        "",
        "## 输出规则",
        "",
        "- 压缩表不输出桥接状态和原始年份列表。",
        "- 中间压缩表不再生成。",
        "- 特殊版本结合进基础记录时使用 Incl: 前缀。",
        "- 特殊版本未匹配基础记录而单独输出时不加 Incl: 前缀。",
        "- 非皮卡原子事实会按 MODEL 中的 | 拆成多条记录。",
        "- 非皮卡压缩前按 BRAND、MODEL、开始年升序排序；开始年为空时从年份区间左端计算。",
        "- 非皮卡无损压缩只合并同 BRAND、MODEL 下年份连续且尺码、结构、版本都相同的记录。",
        "- 非皮卡高度压缩从无损压缩表出发，对同 BRAND、MODEL 组内每两条记录尝试组合。",
        "- 非皮卡高度压缩合并结构和版本时会去重；候选结果临时展开后，原子事实不能对应多个尺码，也不能对应不同尺码。",
        "- 皮卡无损压缩只合并同车型、同版本、同 Cab、同 Bed、同尺码的连续年份。",
        "- 皮卡特定性压缩先尝试同车型、同版本、同 Cab、同 Bed、同尺码闭合年份空洞。",
        "- 皮卡特定性压缩再尝试同车型、同版本、同 Cab、同尺码合并 Bed 并闭合年份。",
        "- 皮卡特定性压缩补出的年份和 Bed 组合不能和已有尺码事实冲突。",
        "- 皮卡特定性压缩同样保证原表原子事实不会对应多条压缩表展开事实。",
    ]
    return "\n".join(lines) + "\n"


def export_table(df: pd.DataFrame) -> pd.DataFrame:
    return df.rename(columns={"BRAND": "MAKE"})


def export_non_pickup_table(df: pd.DataFrame) -> pd.DataFrame:
    return df.rename(columns={"主车型": "CAR", "BRAND": "MAKE"})[NON_PICKUP_EXPORT_COLUMNS]


def export_pickup_table(df: pd.DataFrame) -> pd.DataFrame:
    work = df.rename(columns={"主车型": "CAR", "BRAND": "MAKE", "BED_FT": "BED"})
    return work[PICKUP_EXPORT_COLUMNS]


def write_outputs(
    non_pickup_lossless_df: pd.DataFrame,
    non_pickup_high_df: pd.DataFrame,
    non_pickup_higher_df: pd.DataFrame,
    pickup_lossless_df: pd.DataFrame,
    pickup_specificity_df: pd.DataFrame,
    log_df: pd.DataFrame,
    atom_df: pd.DataFrame,
    input_path: Path,
    output_dir: Path,
    check_atom: bool = False,
    progress: ProgressReporter | None = None,
) -> dict[str, Path]:
    stem = input_path.stem
    project_output_dir = output_dir / stem
    project_output_dir.mkdir(parents=True, exist_ok=True)
    compress_dir = project_output_dir / "compress"
    check_dir = project_output_dir / "check"
    compress_dir.mkdir(parents=True, exist_ok=True)
    if check_atom:
        check_dir.mkdir(parents=True, exist_ok=True)

    for stale_path in [*compress_dir.glob(f"{stem}_*"), *check_dir.glob(f"{stem}_*"), *project_output_dir.glob(f"{stem}_*.xlsx")]:
        if stale_path.is_file():
            stale_path.unlink()

    non_pickup_lossless_tsv_path = compress_dir / f"{stem}_非皮卡无损压缩表.tsv"
    non_pickup_higher_tsv_path = compress_dir / f"{stem}_非皮卡高度压缩表.tsv"
    pickup_lossless_tsv_path = compress_dir / f"{stem}_皮卡无损压缩.tsv"
    pickup_specificity_tsv_path = compress_dir / f"{stem}_皮卡高度压缩表.tsv"
    log_tsv_path = compress_dir / f"{stem}_压缩log.tsv"
    atom_tsv_path = compress_dir / f"{stem}_原子事实表.tsv"
    non_pickup_check_tsv_path = check_dir / f"{stem}_非皮卡原子检查.tsv"
    pickup_check_tsv_path = check_dir / f"{stem}_皮卡原子检查.tsv"
    xlsx_path = project_output_dir / f"{stem}_output.xlsx"

    if not non_pickup_lossless_df.empty:
        export_non_pickup_table(non_pickup_lossless_df).to_csv(
            non_pickup_lossless_tsv_path, sep="\t", index=False, encoding="utf-8-sig"
        )
        export_non_pickup_table(non_pickup_higher_df).to_csv(
            non_pickup_higher_tsv_path, sep="\t", index=False, encoding="utf-8-sig"
        )
    if not pickup_lossless_df.empty:
        export_pickup_table(pickup_lossless_df).to_csv(pickup_lossless_tsv_path, sep="\t", index=False, encoding="utf-8-sig")
        export_pickup_table(pickup_specificity_df).to_csv(
            pickup_specificity_tsv_path, sep="\t", index=False, encoding="utf-8-sig"
        )
    export_table(log_df).to_csv(log_tsv_path, sep="\t", index=False, encoding="utf-8-sig")
    export_table(atom_df).to_csv(atom_tsv_path, sep="\t", index=False, encoding="utf-8-sig")
    non_pickup_check_df = pd.DataFrame()
    pickup_check_df = pd.DataFrame()
    if check_atom:
        atom_export = export_table(atom_df)
        if not non_pickup_higher_df.empty:
            non_pickup_atoms = atom_export[atom_export["压缩类型"].map(normalize_text) == "非皮卡"].copy()
            non_pickup_check_df = build_atom_check(
                non_pickup_atoms,
                export_non_pickup_table(non_pickup_higher_df),
                progress=progress,
                progress_phase="非皮卡原子检查",
            )
            non_pickup_check_df.to_csv(non_pickup_check_tsv_path, sep="\t", index=False, encoding="utf-8-sig")
        if not pickup_specificity_df.empty:
            pickup_atoms = atom_export[atom_export["压缩类型"].map(normalize_text) == "皮卡"].copy()
            pickup_check_df = build_atom_check(
                pickup_atoms,
                export_pickup_table(pickup_specificity_df),
                progress=progress,
                progress_phase="皮卡原子检查",
            )
            pickup_check_df.to_csv(pickup_check_tsv_path, sep="\t", index=False, encoding="utf-8-sig")

    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        if not non_pickup_lossless_df.empty:
            export_non_pickup_table(non_pickup_lossless_df).to_excel(writer, sheet_name="非皮卡无损", index=False)
            export_non_pickup_table(non_pickup_higher_df).to_excel(writer, sheet_name="非皮卡高度", index=False)
        if not pickup_lossless_df.empty:
            export_pickup_table(pickup_lossless_df).to_excel(writer, sheet_name="皮卡无损", index=False)
            export_pickup_table(pickup_specificity_df).to_excel(writer, sheet_name="皮卡高度", index=False)
        export_table(log_df).to_excel(writer, sheet_name="压缩log", index=False)
        export_table(atom_df).to_excel(writer, sheet_name="原子事实表", index=False)
        if check_atom and not non_pickup_check_df.empty:
            non_pickup_check_df.to_excel(writer, sheet_name="非皮卡原子检查", index=False)
        if check_atom and not pickup_check_df.empty:
            pickup_check_df.to_excel(writer, sheet_name="皮卡原子检查", index=False)

    paths = {
        "output_dir": project_output_dir,
        "compress_dir": compress_dir,
        "log_tsv": log_tsv_path,
        "atom_tsv": atom_tsv_path,
        "xlsx": xlsx_path,
    }
    if check_atom:
        paths["check_dir"] = check_dir
        if not non_pickup_check_df.empty:
            paths["non_pickup_check_tsv"] = non_pickup_check_tsv_path
        if not pickup_check_df.empty:
            paths["pickup_check_tsv"] = pickup_check_tsv_path
    if not non_pickup_lossless_df.empty:
        paths["non_pickup_lossless_tsv"] = non_pickup_lossless_tsv_path
        paths["non_pickup_higher_tsv"] = non_pickup_higher_tsv_path
    if not pickup_lossless_df.empty:
        paths["pickup_lossless_tsv"] = pickup_lossless_tsv_path
        paths["pickup_higher_tsv"] = pickup_specificity_tsv_path
    return paths


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Compress fitment size data from TSV or configured Excel sheets.")
    parser.add_argument(
        "input",
        type=Path,
        nargs="?",
        help="Path to the input TSV/Excel file. Can be omitted when --field-profile YAML sets input.path.",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        type=Path,
        default=DEFAULT_OUTPUT_DIR,
        help="Directory for generated files. Defaults to data/output.",
    )
    parser.add_argument(
        "--encoding",
        default="utf-8-sig",
        help="Input file encoding. Defaults to utf-8-sig.",
    )
    parser.add_argument(
        "--field-profile",
        type=Path,
        default=None,
        help="YAML file that maps input column names and can optionally set input.path/input.sheets.",
    )
    parser.add_argument(
        "--remove-null-size",
        action="store_true",
        help="Filter out rows whose size is 无可用尺码. By default these rows are kept.",
    )
    parser.add_argument(
        "--check-atom",
        action="store_true",
        help="Check atom facts against high-compression tables and write check TSVs.",
    )
    parser.add_argument(
        "--progress-interval",
        type=float,
        default=10.0,
        help="Seconds between progress updates. Defaults to 10.",
    )
    parser.add_argument(
        "--no-progress",
        action="store_true",
        help="Disable periodic progress output.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    field_profile_path = args.field_profile.resolve() if args.field_profile else None
    field_profile = load_field_profile(field_profile_path)
    input_path = resolve_input_path(args.input, field_profile, field_profile_path)

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    input_tables = read_input_tables(input_path, encoding=args.encoding, field_profile=field_profile)
    for table_index, (table_name, table_path, df) in enumerate(input_tables, start=1):
        if len(input_tables) > 1:
            print(f"\n处理工作表 {table_index}/{len(input_tables)}: {table_name}")
        progress = ProgressReporter(interval_seconds=args.progress_interval, enabled=not args.no_progress)
        (
            non_pickup_lossless_df,
            non_pickup_high_df,
            non_pickup_higher_df,
            pickup_lossless_df,
            pickup_specificity_df,
            log_df,
            atom_df,
        ) = transform_all_outputs(
            df,
            remove_null_size=args.remove_null_size,
            progress=progress,
            field_profile=field_profile,
        )
        output_paths = write_outputs(
            non_pickup_lossless_df,
            non_pickup_high_df,
            non_pickup_higher_df,
            pickup_lossless_df,
            pickup_specificity_df,
            log_df,
            atom_df,
            table_path,
            args.output_dir,
            check_atom=args.check_atom,
            progress=progress,
        )

        print("写入完成")
        print(f"Input table: {table_name}")
        print(f"Output dir: {output_paths['output_dir']}")
        if "non_pickup_lossless_tsv" in output_paths:
            print(f"Non-pickup lossless TSV: {output_paths['non_pickup_lossless_tsv']}")
            print(f"Non-pickup high TSV: {output_paths['non_pickup_higher_tsv']}")
        if "pickup_lossless_tsv" in output_paths:
            print(f"Pickup lossless TSV: {output_paths['pickup_lossless_tsv']}")
            print(f"Pickup high TSV: {output_paths['pickup_higher_tsv']}")
        print(f"Log TSV: {output_paths['log_tsv']}")
        print(f"Atom TSV: {output_paths['atom_tsv']}")
        if args.check_atom:
            if "non_pickup_check_tsv" in output_paths:
                print(f"Non-pickup atom check TSV: {output_paths['non_pickup_check_tsv']}")
            if "pickup_check_tsv" in output_paths:
                print(f"Pickup atom check TSV: {output_paths['pickup_check_tsv']}")
        print(f"Excel: {output_paths['xlsx']}")


if __name__ == "__main__":
    main()
