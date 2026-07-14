"""CLI 入口 — 参数解析、流程编排、输出。"""
from argparse import ArgumentParser
from pathlib import Path
from sys import exit, stderr

from . import discovery
from .extractors import extract_file
from .section import filter_section


def main() -> None:
    """CLI 入口：解析参数、收集文件、提取内容、写入输出。"""
    parser = ArgumentParser(
        prog='extract_files',
        description='从文档文件中提取文本和资源（docx, pptx, pdf, xlsx, ...）',
    )
    # --- 搜索模式 ---
    parser.add_argument('--root', help='ASCII 安全的祖先目录（搜索模式）')
    parser.add_argument('--glob', help='不区分大小写的子串匹配（搜索模式）')
    parser.add_argument('--max-depth', type=int, default=6, help='最大目录深度（默认 6）')
    parser.add_argument('--list-only', action='store_true', help='仅列出文件，不做提取')

    # --- 直连模式 ---
    parser.add_argument('--paths-file', help='UTF-8 文件，每行一个绝对路径')

    # --- 通用选项 ---
    parser.add_argument('--section', help='只提取匹配此关键字的小节')
    parser.add_argument('--assets-dir', help='将嵌入资源提取到此目录')

    args = parser.parse_args()

    # --- 标准化所有路径参数（MSYS -> Windows） ---
    if args.root:
        args.root = discovery.normalize_path(args.root)
    if args.paths_file:
        args.paths_file = discovery.normalize_path(args.paths_file)
    if args.assets_dir:
        args.assets_dir = discovery.normalize_path(args.assets_dir)

    # --- 收集文件列表 ---
    file_list: list[str] = []
    if args.paths_file:
        if not Path(args.paths_file).exists():
            print(f'[Error] --paths-file 文件不存在: {args.paths_file}', file=stderr)
            exit(1)
        file_list = discovery.read_paths_file(args.paths_file)
    
    elif args.root:
        if not Path(args.root).is_dir():
            print(f'[Error] --root 不是目录: {args.root}', file=stderr)
            exit(1)
        file_list = discovery.find_files(args.root, pattern=args.glob, max_depth=args.max_depth)
    
    else:
        parser.error('必须指定 --root 或 --paths-file 之一。')

    if not file_list:
        print('[未找到支持的文档文件]', file=stderr)
        exit(0)

    # --- 仅列表模式 ---
    if args.list_only:
        for f in file_list:
            print(f)
        exit(0)

    # --- 确定并创建输出目录 ---
    if args.paths_file:
        output_dir = Path(args.paths_file).parent
    else:
        output_dir = Path(file_list[0]).parent / 'temp'
    output_dir.mkdir(parents=True, exist_ok=True)

    # --- 逐个提取文件（各格式提取器内部自行处理缺失依赖） ---
    all_lines: list[str] = []
    for i, filepath in enumerate(file_list, 1):
        filepath = discovery.normalize_path(filepath)
        all_lines.append(f'{"=" * 60}')
        all_lines.append(f'[{i}/{len(file_list)}] {Path(filepath).name}')
        all_lines.append(f'路径: {filepath}')
        all_lines.append(f'{"=" * 60}')
        all_lines.append('')
        try:
            lines = extract_file(filepath, assets_dir=args.assets_dir)
            # 逐文件过滤小节，避免跨文件标记互相干扰
            if args.section:
                lines = filter_section(lines, args.section)
            all_lines.extend(lines)
        
        except Exception as e:
            all_lines.append(f'[提取 {filepath} 时出错: {e}]')
        
        all_lines.append('')

    # --- 写入输出（output_dir 已在上方创建） ---
    output_path = str(output_dir / 'tmp_output.txt')
    Path(output_path).write_text('\n'.join(all_lines), encoding='utf-8')

    # SEARCH 模式需要告知输出路径（DIRECT 模式路径是约定的）
    if not args.paths_file:
        print(f'OUTPUT_PATH:{output_path}', file=stderr)
