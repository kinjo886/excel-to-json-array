"""
从 Excel（.xlsx）读取数据，导出为 JSON 二维字符串数组。

- 若每行占多列：按列导出。
- 若每行只有 1 个单元格且内容是「英文逗号分隔」：自动按 CSV 规则拆成多列（常见：整行粘到 A 列）。

依赖：pip install openpyxl

用法：
  python txt_to_2d_json_array.py
  python txt_to_2d_json_array.py 二维数组.xlsx -o 别的名字.json
  python txt_to_2d_json_array.py 二维数组.xlsx --sheet Sheet1 --skip-rows 1
  python txt_to_2d_json_array.py 二维数组.xlsx --no-split-csv-cell
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import io
import json
import sys
from pathlib import Path
from typing import Any


DEFAULT_BATCH_SIZE = 10


def _cell_to_str(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return str(value)
    if isinstance(value, dt.datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, dt.date):
        return value.isoformat()
    return str(value).strip()


def _maybe_split_csv_row(cells: list[str], *, split_single_cell_csv: bool) -> list[str]:
    """若整行被粘在一个单元格里（逗号分隔），拆成多列。"""
    if not split_single_cell_csv or len(cells) != 1:
        return cells
    s = cells[0]
    if "," not in s:
        return cells
    reader = csv.reader(io.StringIO(s), delimiter=",", skipinitialspace=True)
    for row in reader:
        return [c.strip() for c in row]
    return cells


def read_rows_from_xlsx(
    path: Path,
    *,
    sheet_name: str | None,
    skip_rows: int,
    split_single_cell_csv: bool,
) -> list[list[str]]:
    try:
        from openpyxl import load_workbook  # type: ignore
    except ImportError as exc:
        raise SystemExit("请先安装：pip install openpyxl") from exc

    wb = load_workbook(path, read_only=True, data_only=True)
    if sheet_name and sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        if sheet_name:
            print(f"警告：未找到工作表 {sheet_name!r}，使用第一个表：{wb.sheetnames[0]!r}", file=sys.stderr)
        ws = wb[wb.sheetnames[0]]

    rows: list[list[str]] = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i < skip_rows:
            continue
        if row is None:
            continue
        cells = [_cell_to_str(c) for c in row]
        cells = _maybe_split_csv_row(cells, split_single_cell_csv=split_single_cell_csv)
        if not any(cells):
            continue
        rows.append(cells)

    return rows


def chunk_rows(rows: list[list[str]], batch_size: int) -> list[list[list[str]]]:
    """将二维数组按固定批次拆分。

    Args:
        rows: 原始二维数组数据。
        batch_size: 每批最大行数，必须大于 0。

    Returns:
        按批次切分后的三维数组，最外层每个元素代表一个批次。

    Raises:
        ValueError: 当 batch_size 小于等于 0 时抛出。
    """
    if batch_size <= 0:
        raise ValueError("batch_size 必须大于 0。")
    return [rows[i : i + batch_size] for i in range(0, len(rows), batch_size)]


def build_chunk_output_path(base_output: Path, chunk_index: int) -> Path:
    """生成分批导出的目标文件路径。

    Args:
        base_output: 基础输出路径（例如 data.json）。
        chunk_index: 分批序号（从 1 开始）。

    Returns:
        带批次后缀的路径（例如 data_001.json）。
    """
    return base_output.with_name(f"{base_output.stem}_{chunk_index:03d}{base_output.suffix}")


def main() -> int:
    p = argparse.ArgumentParser(description="Excel → JSON 二维字符串数组")
    p.add_argument(
        "xlsx",
        nargs="?",
        default="二维数组.xlsx",
        help="输入 .xlsx 路径（默认：当前目录 二维数组.xlsx）",
    )
    p.add_argument("-o", "--output", help="输出 .json 路径（默认同名 .json）")
    p.add_argument("--sheet", default=None, help="工作表名称（默认第一个 sheet）")
    p.add_argument(
        "--skip-rows",
        type=int,
        default=0,
        help="跳过前 N 行（可用于跳过表头），默认 0",
    )
    p.add_argument(
        "--no-split-csv-cell",
        action="store_true",
        help="禁止把「单列且含逗号」的单元格按 CSV 拆成多列（默认会拆分）",
    )
    p.add_argument(
        "--batch-size",
        type=int,
        default=DEFAULT_BATCH_SIZE,
        help="每个 JSON 文件最多包含多少行，默认 10",
    )
    args = p.parse_args()

    src = Path(args.xlsx)
    if not src.is_file():
        print(f"错误：文件不存在: {src.resolve()}", file=sys.stderr)
        return 1

    rows = read_rows_from_xlsx(
        src,
        sheet_name=args.sheet,
        skip_rows=args.skip_rows,
        split_single_cell_csv=not args.no_split_csv_cell,
    )
    if not rows:
        print(f"错误：未读到任何数据行: {src.resolve()}", file=sys.stderr)
        return 2

    if args.batch_size <= 0:
        print("错误：--batch-size 必须大于 0", file=sys.stderr)
        return 3

    # 先统一列数，保证每个批次结构一致
    max_cols = max(len(r) for r in rows)
    normalized = [r + [""] * (max_cols - len(r)) for r in rows]

    out = Path(args.output) if args.output else src.with_suffix(".json")
    row_chunks = chunk_rows(normalized, args.batch_size)
    for idx, one_chunk in enumerate(row_chunks, start=1):
        payload = json.dumps(one_chunk, ensure_ascii=False, separators=(",", ":"))
        out_path = build_chunk_output_path(out, idx)
        out_path.write_text(payload, encoding="utf-8")
        print(f"批次: {idx}，行数: {len(one_chunk)}，列数: {max_cols}，已写入: {out_path.resolve()}")

    print(f"总行数: {len(normalized)}，总批次: {len(row_chunks)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())