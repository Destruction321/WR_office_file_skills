"""
CLI 入口 — 支持命令行直接调用。
"""
from argparse import ArgumentParser
from json import load, loads
from pathlib import Path
from sys import exit, stderr

from .filler import fill_template


def main():
    parser = ArgumentParser(description='在模板文件中替换占位符')
    parser.add_argument('--template', '-t', required=True, help='模板文件路径')
    parser.add_argument('--output', '-o', required=True, help='输出文件路径')
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
    parser.add_argument('--force', '-f', action='store_true', help='覆盖已存在的输出文件')

    args = parser.parse_args()

    # 从所有来源合并占位符数据
    content_map = {}

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
        print('未提供内容。请使用 --set、--data 或 --data-file。', file=stderr)
        exit(1)

    try:
        result = fill_template(args.template, args.output, content_map, args.pattern)
        print(f'已写入: {result}')
    except FileExistsError:
        if args.force:
            existing = Path(args.output)
            if existing.exists():
                existing.unlink()
            result = fill_template(args.template, args.output, content_map, args.pattern)
            print(f'已写入（覆盖）: {result}')
        else:
            print(f'错误: {args.output} 已存在。使用 --force 覆盖。', file=stderr)
            exit(1)
    except Exception as e:
        print(f'错误: {e}', file=stderr)
        exit(1)


if __name__ == '__main__':
    main()
