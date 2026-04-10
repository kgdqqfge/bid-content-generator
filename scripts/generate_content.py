"""
调用LLM API批量生成标书务虚内容。

从extract_bid_structure.py输出的结构JSON中读取需要填充的章节，
为每个叶子节点生成写作Prompt，调用兼容OpenAI格式的LLM API生成正文内容。

支持任意OpenAI兼容API（智谱AI、DeepSeek、通义千问、本地Ollama等）。

用法:
  python generate_content.py structure.json \
    --api-key sk-xxx \
    --model glm-4.5-air \
    --api-base https://open.bigmodel.cn/api/paas/v4 \
    --output generated_content.json \
    [--batch-size 10] \
    [--delay 3] \
    [--max-tokens 4000] \
    [--chars-per-node 700] \
    [--batch 1]

参数说明:
  structure.json    - 上一步extract_bid_structure.py输出的结构JSON
  --api-key         - LLM API Key（也可通过环境变量LLM_API_KEY或文件~/.llm_api_key）
  --model           - 模型名称（默认: glm-4.5-air）
  --api-base        - API Base URL（默认: https://open.bigmodel.cn/api/paas/v4）
  --output          - 输出JSON文件（默认: generated_content.json）
  --batch-size      - 每批生成数量（默认: 10）
  --delay           - 请求间隔秒数（默认: 3）
  --max-tokens      - 每次生成的最大token（默认: 4000）
  --chars-per-node  - 每个节点的目标字数（默认: 700）
  --batch           - 指定运行某一批次（不传则运行全部）
  --retries         - 重试次数（默认: 3）
  --project-context - 项目背景描述（如"三亚崖州湾科技城车联网项目"）
  --industry        - 行业领域（默认: 从文件名或上下文推断）
"""

import sys, os, json, time, re, argparse, urllib.request, urllib.error
sys.stdout.reconfigure(encoding='utf-8')

# 务虚风格系统Prompt
DEFAULT_SYSTEM_PROMPT = """你是一位资深的IT项目实施与咨询服务专家，擅长撰写投标文件中的实施方案、项目管理方案、售后服务方案、培训方案、质量保障方案等章节。

写作要求：
1. 务实但不虚浮：内容要有深度和专业性，按照项目管理/实施/售后的通用知识体系来写，不要硬往某个具体项目上靠。
2. 提到公司时用"我司"代替，不要出现具体公司名称。
3. 每个章节按照指定字数撰写高质量正文。
4. 不要在正文开头加数字编号或序号（如"一、""1.""（1）"等），直接写正文内容。
5. 内容要连贯、专业，逻辑清晰，可以引用行业标准和最佳实践。
6. 涉及表格时用文字描述即可，不要使用Markdown表格语法。
7. 语言风格正式但不生硬，体现专业水准和丰富的行业经验。"""


def read_api_key(api_key_arg):
    """读取API Key: 优先使用参数，其次环境变量，最后配置文件"""
    if api_key_arg:
        return api_key_arg
    key = os.environ.get('LLM_API_KEY', '')
    if key:
        return key
    key = os.environ.get('ZHIPUAI_API_KEY', '')
    if key:
        return key
    # 尝试常见key文件
    for path in ['~/.llm_api_key', '~/.zhipuai_key']:
        full_path = os.path.expanduser(path)
        if os.path.exists(full_path):
            with open(full_path) as f:
                return f.read().strip()
    return None


def call_llm(prompt, system_prompt, api_key, api_base, model, max_tokens, retries=3, delay=3):
    """调用LLM API（OpenAI兼容格式）"""
    url = f"{api_base.rstrip('/')}/chat/completions"
    headers = {
        'Authorization': f'Bearer {api_key}',
        'Content-Type': 'application/json'
    }
    data = json.dumps({
        'model': model,
        'messages': [
            {'role': 'system', 'content': system_prompt},
            {'role': 'user', 'content': prompt}
        ],
        'max_tokens': max_tokens,
        'temperature': 0.7,
    }).encode('utf-8')

    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, data=data, headers=headers)
            resp = urllib.request.urlopen(req, timeout=120)
            result = json.loads(resp.read())
            content = result['choices'][0]['message']['content']
            return content.strip()
        except urllib.error.HTTPError as e:
            if e.code == 429:
                wait = (attempt + 1) * 10
                print(f'    限速429, 等待{wait}秒...')
                time.sleep(wait)
            else:
                try:
                    body = e.read().decode('utf-8', errors='replace')
                    print(f'    HTTP {e.code}: {body[:200]}')
                except:
                    print(f'    HTTP {e.code}')
                time.sleep(5)
        except Exception as e:
            err_str = str(e)[:150]
            print(f'    错误: {err_str}')
            time.sleep(5)
    return None


def get_leaf_nodes(structure):
    """从结构JSON中提取所有需要填充内容的叶子节点"""
    flat = structure.get('flat', [])
    return [n for n in flat if n.get('needs_content') and n.get('has_placeholder')]


def build_parent_chain(flat, node_idx):
    """构建从根到当前节点的标题链"""
    chain = []
    node = flat[node_idx]
    for i in range(node_idx - 1, -1, -1):
        prev = flat[i]
        if prev['level'] < node['level']:
            chain.insert(0, prev['text'])
            node = prev
            if node['level'] == 1:
                break
    return chain


def build_number_path(flat, node_idx):
    """构建编号路径，如 5.1.2.1"""
    node = flat[node_idx]
    chain = build_parent_chain(flat, node_idx)
    chain.append(node['text'])

    # 尝试从文本中提取编号
    numbers = []
    for title in chain:
        m = re.match(r'^(\d+(?:\.\d+)*)', title)
        if m:
            numbers = m.group(1).split('.')
        else:
            numbers.append('1')
    return '.'.join(numbers[:len(chain)])


def make_unique_key(node, flat, node_idx):
    """为节点生成唯一key，处理重复标题"""
    parent_chain = build_parent_chain(flat, node_idx)
    title = node['text']

    # 如果有编号前缀，用编号+标题
    m = re.match(r'^(\d+(?:\.\d+)*)\s+(.+)', title)
    if m:
        return f"{m.group(1)} {m.group(2)}"

    # 否则用父级链+标题
    if parent_chain:
        return f"{parent_chain[-1]}>{title}"
    return title


def make_prompt(node, flat, node_idx, chars_per_node, project_context, industry):
    """为每个叶子节点生成写作Prompt"""
    parent_chain = build_parent_chain(flat, node_idx)
    parents_str = ' > '.join(parent_chain) if parent_chain else '(顶级章节)'

    context_section = ""
    if project_context:
        context_section = f"\n项目背景：{project_context}"
    if industry:
        context_section += f"\n行业领域：{industry}"

    return f"""请为以下投标文件章节撰写正文内容。

章节标题：{node['text']}
所属大纲：{parents_str}
目标字数：{chars_per_node}字左右
{context_section}

请按照该标题的语义范围，撰写{chars_per_node}字左右的正文内容。注意：
- 这是投标文件中的正文段落，直接写内容，不要加任何编号前缀
- 按照该领域的通用专业知识和最佳实践来写
- 如果上级标题是宏观性的（如"质量保证措施"），请聚焦到当前标题的具体内容
- 保持与上下级标题的逻辑连贯性
- 内容体现专业水准和丰富的行业经验"""


def main():
    parser = argparse.ArgumentParser(description='调用LLM批量生成标书务虚内容')
    parser.add_argument('structure_json', help='结构JSON文件路径')
    parser.add_argument('--api-key', default=None, help='LLM API Key')
    parser.add_argument('--model', default='glm-4.5-air', help='模型名称')
    parser.add_argument('--api-base', default='https://open.bigmodel.cn/api/paas/v4', help='API Base URL')
    parser.add_argument('--output', default='generated_content.json', help='输出JSON文件')
    parser.add_argument('--batch-size', type=int, default=10, help='每批数量')
    parser.add_argument('--delay', type=float, default=3, help='请求间隔秒数')
    parser.add_argument('--max-tokens', type=int, default=4000, help='最大token')
    parser.add_argument('--chars-per-node', type=int, default=700, help='每节点目标字数')
    parser.add_argument('--batch', type=int, default=None, help='指定批次号')
    parser.add_argument('--retries', type=int, default=3, help='重试次数')
    parser.add_argument('--project-context', default=None, help='项目背景')
    parser.add_argument('--industry', default=None, help='行业领域')
    args = parser.parse_args()

    # 读取结构
    with open(args.structure_json, 'r', encoding='utf-8') as f:
        structure = json.load(f)

    # 获取叶子节点
    leaves = get_leaf_nodes(structure)
    if not leaves:
        print("[警告] 未找到需要填充内容的叶子节点")
        print("请检查: 1) 占位符关键词是否正确 2) 文档是否包含标题+占位符的结构")
        # 也尝试所有有占位符的节点
        flat = structure.get('flat', [])
        leaves = [n for n in flat if n.get('has_placeholder')]
        if not leaves:
            print("[错误] 没有任何可填充的节点")
            sys.exit(1)
        print(f"回退到使用所有有占位符的节点: {len(leaves)}个")

    print(f"需要生成内容的节点数: {len(leaves)}")

    # 读取API Key
    api_key = read_api_key(args.api_key)
    if not api_key:
        print("[错误] 未找到API Key，请通过 --api-key 或环境变量 LLM_API_KEY 提供")
        sys.exit(1)

    print(f"模型: {args.model}")
    print(f"API Base: {args.api_base}")
    print(f"每节点目标: {args.chars_per_node}字")

    # 加载已有内容（断点续传）
    existing = {}
    if os.path.exists(args.output):
        with open(args.output, 'r', encoding='utf-8') as f:
            existing = json.load(f)
        print(f"已有内容: {len(existing)}条")

    flat = structure.get('flat', [])

    # 生成key映射
    node_items = []
    for node in leaves:
        idx = node['idx']
        # 在flat中找到对应位置
        flat_idx = None
        for fi, fn in enumerate(flat):
            if fn['idx'] == idx:
                flat_idx = fi
                break
        if flat_idx is None:
            continue

        key = make_unique_key(node, flat, flat_idx)
        node_items.append({
            'key': key,
            'node': node,
            'flat_idx': flat_idx,
            'doc_title': node['text'],  # 文档中的原始标题文本
        })

    # 批次处理
    total = len(node_items)
    total_batches = (total + args.batch_size - 1) // args.batch_size
    batch_range = [args.batch - 1] if args.batch else range(total_batches)

    for bi in batch_range:
        start_idx = bi * args.batch_size
        end_idx = min(start_idx + args.batch_size, total)
        batch = node_items[start_idx:end_idx]

        print(f"\n--- 批次 {bi+1}/{total_batches} (第{start_idx+1}-{end_idx}个) ---")

        for li, item in enumerate(batch):
            key = item['key']

            # 跳过已生成的（字数>=200才认为有效）
            if key in existing and len(existing[key]) >= 200:
                print(f'  [{start_idx + li + 1}/{total}] SKIP (已有{len(existing[key])}字) {item["doc_title"][:40]}')
                continue

            prompt = make_prompt(
                item['node'], flat, item['flat_idx'],
                args.chars_per_node, args.project_context, args.industry
            )
            print(f'  [{start_idx + li + 1}/{total}] 生成: {item["doc_title"][:40]}...', end='', flush=True)

            content = call_llm(
                prompt, DEFAULT_SYSTEM_PROMPT,
                api_key, args.api_base, args.model,
                args.max_tokens, args.retries, args.delay
            )

            if content and len(content) >= 100:
                # 清理内容：去掉可能的开头编号
                content = re.sub(r'^[\s]*[（(]?\s*[一二三四五六七八九十]+\s*[）).、]', '', content).strip()
                content = re.sub(r'^[\s]*\d+[.、]\s*', '', content).strip()

                existing[key] = content
                # 实时保存（断点续传）
                with open(args.output, 'w', encoding='utf-8') as f:
                    json.dump(existing, f, ensure_ascii=False, indent=2)
                print(f' OK ({len(content)}字)')
            else:
                content_len = len(content) if content else 0
                print(f' FAIL ({content_len}字)')

            time.sleep(args.delay)

    # 统计
    total_chars = sum(len(v) for v in existing.values())
    total_done = sum(1 for item in node_items if item['key'] in existing and len(existing[item['key']]) >= 200)

    print(f"\n=== 生成统计 ===")
    print(f"已完成: {total_done}/{total} 个节点")
    print(f"总字数: {total_chars}")
    print(f"输出文件: {args.output}")


if __name__ == '__main__':
    main()
