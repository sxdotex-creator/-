"""命令行接口。"""

import argparse
import sys
from typing import List, Optional

from .core import load_json_config, parse_map_args, generate_documents, DEFAULT_PLACEHOLDER_STYLES


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="根据 Excel 数据填充 Word 模板并生成文档。"
    )
    parser.add_argument("--excel", required=True, help="Excel 文件路径 (.xlsx)")
    parser.add_argument("--template", required=True, help="Word 模板文件路径 (.docx)")
    parser.add_argument("--output", required=True, help="输出目录")
    parser.add_argument("--sheet", default=None, help="工作表名称，默认第一个")
    parser.add_argument("--start-row", type=int, default=2, help="数据起始行，默认 2")
    parser.add_argument(
        "--map",
        action="append",
        default=[],
        help="列映射: Excel表头=模板占位符，例如 --map 盒号=BOX_NO",
    )
    parser.add_argument(
        "--group-by",
        default=None,
        help="按 Excel 表头分组输出一个 Word 文档，例如 --group-by 盒号",
    )
    parser.add_argument(
        "--output-name",
        default=None,
        help="输出文件名模板，例如 {BOX_NO}_{DOC_NO}.docx",
    )
    parser.add_argument(
        "--config",
        default=None,
        help="JSON 配置文件路径，优先于 --map",
    )
    parser.add_argument(
        "--repeat-marker",
        default="#repeat",
        help="表格行复制标记占位符名称，默认 #repeat（例如 {{#repeat}}）",
    )
    return parser


def parse_config_or_args(parser_values) -> tuple[dict, Optional[str], Optional[str], Optional[str], Optional[str], List[tuple[str, str]]]:
    mappings = {}
    output_name = None
    group_by = None
    repeat_marker = "#repeat"
    placeholder_styles = DEFAULT_PLACEHOLDER_STYLES

    if parser_values.config:
        config = load_json_config(parser_values.config)
        mappings = config.get("mappings", {})
        output_name = config.get("output_name")
        group_by = config.get("group_by")
        repeat_marker = config.get("repeat_marker", repeat_marker)
        placeholder_styles = config.get("placeholder_styles", placeholder_styles)
    else:
        mappings = parse_map_args(parser_values.map)

    if parser_values.output_name:
        output_name = parser_values.output_name
    if parser_values.group_by:
        group_by = parser_values.group_by
    if parser_values.repeat_marker:
        repeat_marker = parser_values.repeat_marker

    return mappings, output_name, group_by, repeat_marker, placeholder_styles


def main(argv: Optional[List[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        mappings, output_name, group_by, repeat_marker, placeholder_styles = parse_config_or_args(args)
        if not mappings:
            raise ValueError("必须通过 --map 或 --config 提供列映射关系")

        files = generate_documents(
            excel_path=args.excel,
            template_path=args.template,
            output_dir=args.output,
            mapping=mappings,
            output_name_pattern=output_name,
            sheet_name=args.sheet,
            start_row=args.start_row,
            group_by=group_by,
            repeat_marker=repeat_marker,
            placeholder_styles=placeholder_styles,
        )

        print(f"成功生成 {len(files)} 个文档：")
        for path in files:
            print(f"  - {path}")
        return 0
    except Exception as exc:
        print(f"错误: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
