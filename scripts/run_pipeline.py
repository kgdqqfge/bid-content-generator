"""
标书务虚方案一键生成编排脚本。

整合 extract_bid_structure.py → generate_content.py → fill_docx.py 的完整流程。
WorkBuddy Agent调用此脚本即可完成从招标文件到完整务虚方案的全流程。

用法:
  python run_pipeline.py <招标文件.docx> <输出文件.docx> \
    --api-key sk-xxx \
    --model glm-4.5-air \
    --api-base https://open.bigmodel.cn/api/paas/v4 \
    [--project-context "项目描述"] \
    [--industry "行业领域"] \
    [--placeholder "此处填写"] \
    [--batch-size 10] \
    [--chars-per-node 700] \
    [--body-font "宋体"] \
    [--body-size 10.5]

示例:
  # 最简用法（API Key从环境变量读取）
  python run_pipeline.py input.docx output.docx

  # 完整用法
  python run_pipeline.py input.docx output.docx \
    --api-key sk-xxx \
    --model glm-4.5-air \
    --project-context "三亚崖州湾科技城车联网项目" \
    --industry "智慧交通/车联网"
"""

import sys, os, json, argparse, tempfile, subprocess
sys.stdout.reconfigure(encoding='utf-8')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def run_step(script_name, args_list, description):
    """执行一个步骤"""
    script_path = os.path.join(SCRIPT_DIR, script_name)
    if not os.path.exists(script_path):
        print(f"[错误] 脚本不存在: {script_path}")
        return False

    cmd = [sys.executable, '-u', script_path] + args_list
    print(f"\n{'='*60}")
    print(f"步骤: {description}")
    print(f"命令: {' '.join(a[:50] for a in cmd)}")
    print(f"{'='*60}")

    result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
    # 合并stdout和stderr
    output = (result.stdout or '') + (result.stderr or '')
    if output.strip():
        # 按行输出，避免超长行
        for line in output.split('\n'):
            print(line)
    return result.returncode == 0


def main():
    parser = argparse.ArgumentParser(
        description='标书务虚方案一键生成',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
完整流程:
  1. 解析招标文件框架（提取标题结构和占位符位置）
  2. 批量调用LLM生成务虚内容
  3. 将内容回填到docx并打包输出

支持的API: 任何OpenAI兼容API（智谱AI、DeepSeek、通义千问、本地Ollama等）
        """
    )
    parser.add_argument('input_docx', help='招标文件docx（含框架目录）')
    parser.add_argument('output_docx', help='输出的完整方案docx')
    parser.add_argument('--api-key', default=None, help='LLM API Key')
    parser.add_argument('--model', default='glm-4.5-air', help='LLM模型名称')
    parser.add_argument('--api-base', default='https://open.bigmodel.cn/api/paas/v4',
                        help='API Base URL')
    parser.add_argument('--project-context', default=None, help='项目背景描述')
    parser.add_argument('--industry', default=None, help='行业领域')
    parser.add_argument('--placeholder', default='此处填写', help='占位符关键词')
    parser.add_argument('--batch-size', type=int, default=10, help='每批生成数量')
    parser.add_argument('--delay', type=float, default=3, help='API请求间隔秒数')
    parser.add_argument('--max-tokens', type=int, default=4000, help='最大生成token')
    parser.add_argument('--chars-per-node', type=int, default=700, help='每节点目标字数')
    parser.add_argument('--body-font', default='宋体', help='正文字体')
    parser.add_argument('--body-size', type=float, default=10.5, help='正文字号pt')
    parser.add_argument('--line-spacing', type=float, default=1.5, help='行距倍数')
    parser.add_argument('--first-line-indent', type=int, default=2, help='首行缩进字符数')
    parser.add_argument('--work-dir', default=None, help='工作目录（存放中间文件）')
    parser.add_argument('--skip-extract', action='store_true', help='跳过步骤1（已有结构JSON）')
    parser.add_argument('--skip-generate', action='store_true', help='跳过步骤2（已有内容JSON）')
    parser.add_argument('--validate', action='store_true', help='验证打包结果（中文路径建议关闭）')
    args = parser.parse_args()

    # 工作目录
    work_dir = args.work_dir or tempfile.mkdtemp(prefix='bid_pipeline_')
    os.makedirs(work_dir, exist_ok=True)

    structure_json = os.path.join(work_dir, 'structure.json')
    content_json = os.path.join(work_dir, 'generated_content.json')
    unpacked_dir = os.path.join(work_dir, 'unpacked')

    print(f"工作目录: {work_dir}")
    print(f"输入文件: {args.input_docx}")
    print(f"输出文件: {args.output_docx}")
    print(f"模型: {args.model}")

    # 检查输入文件
    if not os.path.exists(args.input_docx):
        print(f"[错误] 输入文件不存在: {args.input_docx}")
        sys.exit(1)

    # ========== 步骤1: 解析结构 ==========
    if not args.skip_extract:
        extract_args = [
            args.input_docx,
            '--placeholder', args.placeholder,
            '--unpacked-dir', unpacked_dir,
            '--output', structure_json,
        ]
        if not run_step('extract_bid_structure.py', extract_args, '解析招标文件框架'):
            print("[错误] 步骤1失败")
            sys.exit(1)
    else:
        print(f"\n跳过步骤1，使用已有结构: {structure_json}")

    # 检查结构JSON
    if not os.path.exists(structure_json):
        print(f"[错误] 结构文件不存在: {structure_json}")
        sys.exit(1)

    with open(structure_json, 'r', encoding='utf-8') as f:
        structure = json.load(f)
    stats = structure.get('stats', {})
    print(f"\n结构分析完成: {stats.get('placeholders_found', 0)}个待填充章节")

    if stats.get('placeholders_found', 0) == 0:
        print("[警告] 未找到占位符章节，检查关键词是否正确")

    # ========== 步骤2: 生成内容 ==========
    if not args.skip_generate:
        gen_args = [
            structure_json,
            '--output', content_json,
            '--model', args.model,
            '--api-base', args.api_base,
            '--batch-size', str(args.batch_size),
            '--delay', str(args.delay),
            '--max-tokens', str(args.max_tokens),
            '--chars-per-node', str(args.chars_per_node),
        ]
        if args.api_key:
            gen_args.extend(['--api-key', args.api_key])
        if args.project_context:
            gen_args.extend(['--project-context', args.project_context])
        if args.industry:
            gen_args.extend(['--industry', args.industry])

        if not run_step('generate_content.py', gen_args, '批量调用LLM生成务虚内容'):
            print("[错误] 步骤2失败")
            sys.exit(1)
    else:
        print(f"\n跳过步骤2，使用已有内容: {content_json}")

    # 检查内容JSON
    if not os.path.exists(content_json):
        print(f"[错误] 内容文件不存在: {content_json}")
        sys.exit(1)

    with open(content_json, 'r', encoding='utf-8') as f:
        content_data = json.load(f)
    total_chars = sum(len(v) for v in content_data.values())
    print(f"\n内容生成完成: {len(content_data)}条, 总字数: {total_chars}")

    # ========== 步骤3: 回填到docx ==========
    fill_args = [
        args.input_docx,
        content_json,
        args.output_docx,
        '--placeholder', args.placeholder,
        '--unpacked-dir', unpacked_dir,
        '--body-font', args.body_font,
        '--body-size', str(args.body_size),
        '--line-spacing', str(args.line_spacing),
        '--first-line-indent', str(args.first_line_indent),
    ]
    if args.validate:
        fill_args.append('--validate')

    if not run_step('fill_docx.py', fill_args, '将内容回填到docx并打包'):
        print("[错误] 步骤3失败")
        sys.exit(1)

    # ========== 完成 ==========
    print(f"\n{'='*60}")
    print(f"✅ 全部完成!")
    print(f"输出文件: {args.output_docx}")
    print(f"总字数: {total_chars}")
    print(f"工作目录: {work_dir}")
    print(f"{'='*60}")


if __name__ == '__main__':
    main()
