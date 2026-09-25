from __future__ import annotations

import argparse
from functools import lru_cache
from pathlib import Path
import re
from typing import Any
import unicodedata

import pandas as pd

from non_pickup_validation import (
    atom_profile,
    profile_matches,
    record_profile,
    split_model_expression,
)


PROJECT_ROOT = Path(__file__).resolve().parent
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "data" / "output"
CHECK_FINAL_COLUMNS = [
    "原子行号",
    "压缩命中数",
    "检查结果",
    "压缩行号",
    "命中尺码",
    "压缩类型",
    "MAKE",
    "MODEL",
    "YEAR",
    "VERSION",
    "CONST",
    "CAB",
    "BED_FT",
    "BACKSIZE",
    "命中压缩行",
]
ATOM_MATCH_KEY_COLUMNS = [
    "压缩类型",
    "MAKE",
    "MODEL",
    "YEAR",
    "VERSION",
    "CONST",
    "CAB",
    "BED_FT",
]
MATCH_KEY_COLUMNS = [
    "压缩类型",
    "MAKE",
    "MODEL",
    "YEAR",
    "VERSION",
    "CONST",
    "CAB",
    "BED_FT",
    "BACKSIZE",
]


def read_tsv(path: Path, encoding: str = "utf-8-sig") -> pd.DataFrame:
    return pd.read_csv(path, sep="\t", dtype=str, encoding=encoding, keep_default_na=False)


def normalize_text(value: object) -> str:
    if isinstance(value, str):
        return _normalize_str(value)
    if pd.isna(value):
        return ""
    return _normalize_str(str(value))


@lru_cache(maxsize=1 << 20)
def _normalize_str(text: str) -> str:
    text = text.replace("\u00a0", " ")
    text = text.replace("\u200b", "")
    text = text.replace("\ufeff", "")
    text = "".join(ch for ch in text if unicodedata.category(ch)[0] != "C")
    return text.strip()


def to_int(value: object) -> int | None:
    text = normalize_text(value)
    if not text:
        return None
    try:
        return int(float(text))
    except ValueError:
        match = re.search(r"\d{4}", text)
        return int(match.group(0)) if match else None


def parse_year_list(value: object) -> list[int]:
    text = normalize_text(value).replace("，", "/").replace(",", "/").replace(";", "/")
    years: list[int] = []
    for part in text.split("/"):
        part = normalize_text(part)
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


def split_tokens(value: object, separator: str = "/") -> set[str]:
    text = normalize_text(value)
    if not text:
        return set()
    return {normalize_text(part) for part in text.split(separator) if normalize_text(part)}


def protect_version_slashes(text: str) -> str:
    return re.sub(r"\b([A-Za-z])\s*/\s*([A-Za-z])\b", r"\1__VERSION_SLASH__\2", text)


def split_version_tokens(value: object) -> set[str]:
    text = normalize_text(value)
    if not text:
        return set()
    text = re.sub(r"\b(?:INCL|Incl|incl|EXCL|Excl|excl|EXP|Exp|exp)\s*:", "", text)
    text = protect_version_slashes(text)
    return {token.replace("__VERSION_SLASH__", "/") for token in split_tokens(text)}


def record_version_matches(record_version: object, atom_version: object) -> bool:
    record_text = normalize_text(record_version)
    atom_text = normalize_text(atom_version)
    if not record_text:
        return atom_text == ""
    record_tokens = split_version_tokens(record_text)
    atom_tokens = split_version_tokens(atom_text)
    if record_text.lower().startswith("incl:"):
        return atom_text == "" or atom_text in record_tokens or bool(record_tokens & atom_tokens)
    return atom_text in record_tokens or bool(record_tokens & atom_tokens)


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


def normalize_schema(df: pd.DataFrame) -> pd.DataFrame:
    work = df.copy()
    work.columns = ["MAKE" if normalize_text(column).upper() == "BRAND" else normalize_text(column) for column in work.columns]
    if "BED_FT" not in work.columns and "BED" in work.columns:
        work["BED_FT"] = work["BED"]
    for column in CHECK_FINAL_COLUMNS:
        if column not in work.columns and column not in {"原子行号", "压缩命中数", "检查结果", "压缩行号", "命中压缩行"}:
            work[column] = ""
    return work


def model_matches(record_model: object, atom_model: object) -> bool:
    if normalize_text(record_model) == normalize_text(atom_model):
        return True
    models = split_tokens(record_model)
    atom = normalize_text(atom_model)
    return atom in models if models else atom == ""


def const_matches(record: pd.Series, atom: pd.Series) -> bool:
    atom_consts = split_tokens(atom.get("CONST", "")) or {"*"}
    record_consts = split_tokens(record.get("CONST", ""))
    return not record_consts or "*" in atom_consts or bool(record_consts & atom_consts)


def row_matches(record: pd.Series, atom: pd.Series) -> bool:
    if normalize_text(record.get("MAKE", "")) != normalize_text(atom.get("MAKE", "")):
        return False
    if not model_matches(record.get("MODEL", ""), atom.get("MODEL", "")):
        return False
    atom_year = to_int(atom.get("YEAR", ""))
    if atom_year is None or atom_year not in parse_year_list(record.get("YEAR", "")):
        return False
    if not record_version_matches(record.get("VERSION", ""), atom.get("VERSION", "")):
        return False
    if normalize_text(atom.get("CAB", "")):
        record_cabs = split_tokens(record.get("CAB", ""))
        if normalize_text(atom.get("CAB", "")) not in record_cabs:
            return False
    if normalize_text(atom.get("BED_FT", "")) and not bed_matches_expression(atom.get("BED_FT", ""), record.get("BED_FT", "")):
        return False
    return const_matches(record, atom)


def atom_key(row: pd.Series) -> tuple[str, ...]:
    return tuple(normalize_text(row.get(column, "")) for column in ATOM_MATCH_KEY_COLUMNS)


def brand_model_key(row: pd.Series) -> tuple[str, str]:
    return (normalize_text(row.get("MAKE", "")), normalize_text(row.get("MODEL", "")))


def is_non_pickup_atom(row: pd.Series) -> bool:
    return not normalize_text(row.get("CAB", "")) and not normalize_text(row.get("BED_FT", ""))


def is_non_pickup_record(row: pd.Series) -> bool:
    return not normalize_text(row.get("CAB", "")) and not normalize_text(row.get("BED_FT", ""))


def compress_summary(line_no: int, record: pd.Series) -> str:
    return (
        f"{line_no}: {normalize_text(record.get('MAKE', ''))} {normalize_text(record.get('MODEL', ''))} "
        f"{normalize_text(record.get('YEAR', ''))} {normalize_text(record.get('VERSION', ''))} "
        f"{normalize_text(record.get('CONST', ''))} {normalize_text(record.get('CAB', ''))} "
        f"{normalize_text(record.get('BED_FT', ''))} {normalize_text(record.get('BACKSIZE', ''))}"
    )


def expand_compress_atoms(
    atom_df: pd.DataFrame,
    compress_df: pd.DataFrame,
    progress: Any = None,
    completed_offset: int = 0,
) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    atoms_by_make: dict[str, list[dict[str, object]]] = {}
    for atom in atom_df.to_dict("records"):
        atoms_by_make.setdefault(normalize_text(atom.get("MAKE", "")), []).append(atom)
    for completed_count, (record_index, record) in enumerate(zip(compress_df.index, compress_df.to_dict("records")), start=1):
        if progress:
            progress.update(
                current_make=record.get("MAKE", ""),
                current_model=record.get("MODEL", ""),
                completed_models=completed_offset + completed_count,
            )
        line_no = int(record_index) + 2
        summary = compress_summary(line_no, record)
        for atom in atoms_by_make.get(normalize_text(record.get("MAKE", "")), ()):
            if not row_matches(record, atom):
                continue
            item = {column: normalize_text(atom.get(column, "")) for column in MATCH_KEY_COLUMNS}
            item["BACKSIZE"] = normalize_text(record.get("BACKSIZE", ""))
            item["压缩行号"] = line_no
            item["压缩行摘要"] = summary
            rows.append(item)
    if not rows:
        return pd.DataFrame(columns=[*MATCH_KEY_COLUMNS, "压缩行号", "压缩行摘要"])
    return pd.DataFrame(rows).drop_duplicates().reset_index(drop=True)


def build_non_pickup_record_lookup(compress_df: pd.DataFrame) -> dict[tuple[str, str], list[dict[str, object]]]:
    grouped_records: dict[tuple[str, str], list[dict[str, object]]] = {}
    for record_index, record in zip(compress_df.index, compress_df.to_dict("records")):
        if not is_non_pickup_record(record):
            continue
        make = normalize_text(record.get("MAKE", ""))
        models = split_model_expression(record.get("MODEL", "")) or [normalize_text(record.get("MODEL", ""))]
        for model in models:
            key = (make, model)
            item = dict(record)
            item["_compress_line_no"] = int(record_index) + 2
            item["_compress_summary"] = compress_summary(int(record_index) + 2, record)
            item["_match_profile"] = record_profile(item)
            grouped_records.setdefault(key, []).append(item)
    return grouped_records


def build_non_pickup_atom_check_row(atom_index: object, atom: pd.Series, records: list[dict[str, object]]) -> dict[str, object]:
    profile = atom_profile(atom)
    matches = [record for record in records if profile_matches(record["_match_profile"], profile)]
    matched_line_numbers: list[str] = []
    matched_sizes: list[str] = []
    match_lines: list[str] = []
    for record in matches:
        matched_line_numbers.append(str(record.get("_compress_line_no", "")))
        matched_sizes.append(normalize_text(record["_match_profile"][-1]))
        match_lines.append(normalize_text(record.get("_compress_summary", "")))

    match_count = len(matches)
    unique_sizes = sorted(set(matched_sizes))
    atom_size = normalize_text(atom.get("BACKSIZE", ""))
    if match_count == 0:
        result = "MISS"
    elif match_count == 1 and unique_sizes == [atom_size]:
        result = "OK"
    elif atom_size not in unique_sizes:
        result = "SIZE_MISMATCH"
    else:
        result = "MULTI"

    return {
        "原子行号": int(atom_index) + 2,
        "压缩命中数": match_count,
        "检查结果": result,
        "压缩行号": "/".join(matched_line_numbers),
        "命中尺码": "/".join(unique_sizes),
        "压缩类型": normalize_text(atom.get("压缩类型", "")),
        "MAKE": normalize_text(atom.get("MAKE", "")),
        "MODEL": normalize_text(atom.get("MODEL", "")),
        "YEAR": normalize_text(atom.get("YEAR", "")),
        "VERSION": normalize_text(atom.get("VERSION", "")),
        "CONST": normalize_text(atom.get("CONST", "")),
        "CAB": normalize_text(atom.get("CAB", "")),
        "BED_FT": normalize_text(atom.get("BED_FT", "")),
        "BACKSIZE": atom_size,
        "命中压缩行": " || ".join(match_lines),
    }


def build_atom_check(
    atom_df: pd.DataFrame,
    compress_df: pd.DataFrame,
    progress: Any = None,
    progress_phase: str = "原子检查",
) -> pd.DataFrame:
    atom_df = normalize_schema(atom_df)
    compress_df = normalize_schema(compress_df)
    if progress:
        progress.start(progress_phase, len(compress_df) + len(atom_df), unit_label="行")
    non_pickup_lookup = build_non_pickup_record_lookup(compress_df)
    atom_records = atom_df.to_dict("records")
    pickup_atoms = atom_df[pd.Series([not is_non_pickup_atom(atom) for atom in atom_records], index=atom_df.index, dtype=bool)]
    pickup_compress = compress_df[
        pd.Series([not is_non_pickup_record(record) for record in compress_df.to_dict("records")], index=compress_df.index, dtype=bool)
    ]
    expanded_compress_df = expand_compress_atoms(pickup_atoms, pickup_compress, progress=progress)
    expanded_positions: dict[tuple[str, ...], list[int]] = {}
    if not expanded_compress_df.empty:
        expanded_keys = zip(*(expanded_compress_df[column].map(normalize_text) for column in ATOM_MATCH_KEY_COLUMNS))
        for position, row_key in enumerate(expanded_keys):
            expanded_positions.setdefault(row_key, []).append(position)
    rows: list[dict[str, object]] = []
    completed_offset = len(compress_df)
    for completed_count, (atom_index, atom) in enumerate(zip(atom_df.index, atom_records), start=1):
        if progress:
            progress.update(
                current_make=atom.get("MAKE", ""),
                current_model=atom.get("MODEL", ""),
                completed_models=completed_offset + completed_count,
            )
        if is_non_pickup_atom(atom):
            records = non_pickup_lookup.get(brand_model_key(atom), [])
            rows.append(build_non_pickup_atom_check_row(atom_index, atom, records))
            continue

        key = atom_key(atom)
        if expanded_compress_df.empty:
            matched_rows = expanded_compress_df
        else:
            matched_rows = expanded_compress_df.iloc[expanded_positions.get(key, [])]
        matches = [str(value) for value in matched_rows["压缩行号"].tolist()]
        match_lines = [normalize_text(value) for value in matched_rows["压缩行摘要"].tolist()]
        match_count = len(matches)
        matched_sizes = sorted(set(matched_rows["BACKSIZE"].map(normalize_text).tolist())) if not matched_rows.empty else []
        atom_size = normalize_text(atom.get("BACKSIZE", ""))
        if match_count == 0:
            result = "MISS"
        elif match_count == 1 and matched_sizes == [atom_size]:
            result = "OK"
        elif atom_size not in matched_sizes:
            result = "SIZE_MISMATCH"
        else:
            result = "MULTI"
        rows.append(
            {
                "原子行号": int(atom_index) + 2,
                "压缩命中数": match_count,
                "检查结果": result,
                "压缩行号": "/".join(matches),
                "命中尺码": "/".join(matched_sizes),
                "压缩类型": normalize_text(atom.get("压缩类型", "")),
                "MAKE": normalize_text(atom.get("MAKE", "")),
                "MODEL": normalize_text(atom.get("MODEL", "")),
                "YEAR": normalize_text(atom.get("YEAR", "")),
                "VERSION": normalize_text(atom.get("VERSION", "")),
                "CONST": normalize_text(atom.get("CONST", "")),
                "CAB": normalize_text(atom.get("CAB", "")),
                "BED_FT": normalize_text(atom.get("BED_FT", "")),
                "BACKSIZE": normalize_text(atom.get("BACKSIZE", "")),
                "命中压缩行": " || ".join(match_lines),
            }
        )
    if progress:
        progress.finish()
    return pd.DataFrame(rows, columns=CHECK_FINAL_COLUMNS)


def check_atom_files(atom_path: Path, compress_path: Path, output_dir: Path, encoding: str = "utf-8-sig") -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    result = build_atom_check(read_tsv(atom_path, encoding=encoding), read_tsv(compress_path, encoding=encoding))
    output_path = output_dir / f"{atom_path.stem}_check.tsv"
    result.to_csv(output_path, sep="\t", index=False, encoding="utf-8-sig")
    return output_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Check atom facts against a compressed table.")
    parser.add_argument("--atom", type=Path, required=True, help="Atom fact TSV.")
    parser.add_argument("--compress", type=Path, required=True, help="Compressed TSV.")
    parser.add_argument("-o", "--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR, help="Output directory.")
    parser.add_argument("--encoding", default="utf-8-sig", help="Input file encoding. Defaults to utf-8-sig.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    output_path = check_atom_files(args.atom.resolve(), args.compress.resolve(), args.output_dir.resolve(), args.encoding)
    print("写入完成")
    print(f"Atom check TSV: {output_path}")


if __name__ == "__main__":
    main()
