"""CLI 入口 — 支持占位符替换、节级内容填充和模板扫描三种模式。"""
from argparse import ArgumentParser
from json import load, loads
from pathlib import Path
from shutil import copy2
from sys import exit, stderr

from .docx_section_filler import scan_docx, fill_docx_sections
from .filler import fill_template
from .md_parser import parse_sections_md


def main() -> None:
    parser = ArgumentParser(description='在模板文件中替换占位符或填充节内容')
    parser.add_argument('--template', '-t', required=True, help='模板文件路径')
    parser.add_argument('--output', '-o', help='输出文件路径（--scan 模式下不需要）')
    # 占位符模式
    parser.add_argument(
        '--set', '-s', action='append', default=[],
        help='设置占位符值: KEY=VALUE（可重复）'
    )
    parser.add_argument('--data', '-d', help='JSON 格式的占位符数据')
    parser.add_argument('--data-file', help='JSON 文件格式的占位符数据')
    parser.add_argument(
        '--pattern', default=r'\{\{\s*(\w+)\s*\}\}',
        help='占位符正则模式（默认: {{ name }}）'
    )
    # 节级填充模式
    parser.add_argument(
        '--section-data-file',
        help='节级填充的 JSON/Markdown 文件（标题 -> 内容项列表）',
    )
    parser.add_argument(
        '--section-mode', choices=('replace', 'append'), default='replace',
        help='节级填充模式：replace（清空旧内容）/ append（追加），默认 replace',
    )
    parser.add_argument(
        '--heading-style',
        help='手动指定 docx 中的标题样式名（如 a4），用于自动检测失败时覆盖',
    )
    # 扫描模式
    parser.add_argument(
        '--scan', action='store_true',
        help='扫描模板标题结构和样式，不做任何修改',
    )
    parser.add_argument(
        '--dry-run', action='store_true',
        help='仅检测标题定位和边界，不修改文件',
    )
    parser.add_argument('--force', '-f', action='store_true', help='覆盖已存在的输出文件')

    args = parser.parse_args()
    template = Path(args.template)
    
    if not template.exists():
        print(f'错误: 模板不存在: {template}', file=stderr)
        exit(1)

    # --- 扫描模式 ---
    if args.scan:
        scan_docx(template)
        return

    # 以下模式都需要 --output
    if not args.output:
        print('错误: 非扫描模式下 --output 是必需的。', file=stderr)
        exit(1)

    output = Path(args.output)
    same_file = template.resolve() == output.resolve()

    # --- 节级填充模式 ---
    if args.section_data_file:
        section_path = Path(args.section_data_file)

        # 按后缀自动判断格式：.md -> Markdown，.json -> JSON
        if section_path.suffix.lower() == '.md':
            sections = parse_sections_md(section_path.read_text(encoding='utf-8'))
        else:
            with open(section_path, 'r', encoding='utf-8') as f:
                sections = load(f)

        # 同文件 -> 原地修改；不同文件 -> 先复制
        # dry-run 模式不创建输出文件（只验证定位），直接对模板读取
        if not same_file and not args.dry_run:
            if output.exists() and not args.force:
                print(f'错误: {output} 已存在。使用 --force 覆盖。', file=stderr)
                exit(1)
            if output.exists():
                output.unlink()
            output.parent.mkdir(parents=True, exist_ok=True)
            copy2(template, output)

        # dry-run 时对模板操作（不修改任何文件）；否则对输出文件操作
        target = template if args.dry_run else output
        count = fill_docx_sections(
            target, sections, mode=args.section_mode,
            heading_style=args.heading_style,
            dry_run=args.dry_run,
        )
        if args.dry_run:
            print(f'Dry-run on template: {template}')
        else:
            print(f'已填充 {count} 个节: {output}')
        return

    # --- 占位符模式 ---
    if same_file:
        print('错误: --output 不能与 --template 相同。模板填写始终输出到新文件。', file=stderr)
        exit(1)

    if output.exists() and not args.force:
        print(f'错误: {output} 已存在。使用 --force 覆盖。', file=stderr)
        exit(1)

    if output.exists():
        output.unlink()

    content_map: dict[str, str] = {}
    if args.data:
        content_map.update(loads(args.data))
    if args.data_file:
        with open(args.data_file, 'r', encoding='utf-8') as f:
            content_map.update(load(f))

    for kv in args.set:
        if '=' not in kv:
            print(f'无效的 --set 格式: {kv}（应为 KEY=VALUE）', file=stderr)
            exit(1)
        key, value = kv.split('=', 1)
        content_map[key] = value

    if not content_map:
        print('未提供内容。请使用 --set、--data、--data-file 或 --section-data-file。', file=stderr)
        exit(1)

    result = fill_template(args.template, args.output, content_map, args.pattern)
    print(f'已写入: {result}')


if __name__ == '__main__':
    main()
