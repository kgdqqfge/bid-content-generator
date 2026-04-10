"""
从docx文档中提取目录框架结构，识别需要填充内容的章节。

输入: 招标文件docx（含目录和分级标题的框架文档）
输出: JSON格式的章节结构树，标记哪些是叶子节点（需要填充内容）

用法:
  python extract_bid_structure.py <input.docx> [--placeholder "此处填写"] [--unpacked-dir /tmp/unpacked] [--output structure.json]

参数说明:
  input.docx       - 招标文件docx路径
  --placeholder    - 占位符关键词，用于识别需要填充的章节（默认: "此处填写"）
  --unpacked-dir   - 解包临时目录（默认: 自动创建）
  --output         - 输出JSON文件路径（默认: structure.json）
  --min-level      - 最低标题级别（数字越大层级越深，默认: 0 表示自动检测）
"""

import sys, os, json, argparse, tempfile, subprocess, re
sys.stdout.reconfigure(encoding='utf-8')

def unpack_docx(docx_path, unpacked_dir):
    """使用docx skill的unpack脚本解包docx"""
    pack_script = os.path.join(
        os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
        'plugins', 'marketplaces', 'codebuddy-plugins-official', 'plugins', 'docx',
        'scripts', 'office', 'unpack.py'
    )
    # 也尝试WorkBuddy插件路径
    if not os.path.exists(pack_script):
        pack_script = os.path.expanduser(
            '~/.workbuddy/plugins/marketplaces/codebuddy-plugins-official/'
            'plugins/docx/scripts/office/unpack.py'
        )
    if not os.path.exists(pack_script):
        # 尝试常见安装路径
        import glob
        candidates = glob.glob(
            r'C:\Users\*\.workbuddy\plugins\marketplaces\codebuddy-plugins-official'
            r'\plugins\docx\scripts\office\unpack.py'
        ) + glob.glob(
            r'D:\ProgramData\WorkBuddy\resources\app\extensions\genie\out\extension'
            r'\builtin\docx\scripts\office\unpack.py'
        )
        if candidates:
            pack_script = candidates[0]

    if not os.path.exists(pack_script):
        print(f"[错误] 找不到unpack.py脚本，请确保docx skill已安装")
        print(f"尝试的路径: {pack_script}")
        return False

    os.makedirs(unpacked_dir, exist_ok=True)
    result = subprocess.run(
        [sys.executable, '-u', pack_script, os.path.abspath(docx_path), os.path.abspath(unpacked_dir)],
        capture_output=True
    )
    return result.returncode == 0


def analyze_structure(unpacked_dir, placeholder_keyword):
    """分析解包后的XML，提取标题结构"""
    from lxml import etree

    NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    xml_path = os.path.join(unpacked_dir, 'word', 'document.xml')

    if not os.path.exists(xml_path):
        print(f"[错误] 找不到document.xml: {xml_path}")
        return None

    tree = etree.parse(xml_path)
    root = tree.getroot()
    body = root.find('.//w:body', NS)
    all_elements = list(body)

    # 1. 扫描所有标题样式，建立样式ID -> 级别的映射
    # heading1的pStyle val可能是'1', '2', 'heading1', 'a1'等
    style_level_map = {}
    for elem in all_elements:
        if elem.tag != f'{{{NS["w"]}}}p':
            continue
        ppr = elem.find('w:pPr', NS)
        if ppr is None:
            continue
        style_elem = ppr.find('w:pStyle', NS)
        if style_elem is None:
            continue
        style_val = style_elem.get(f'{{{NS["w"]}}}val', '')
        text = ''.join(t.text or '' for t in elem.findall('.//w:t', NS)).strip()
        if not text:
            continue
        # 通过文本内容推断级别（编号模式: "1 ", "1.1 ", "1.1.1 " 等）
        m = re.match(r'^(\d+(?:\.\d+)*)\s+', text)
        if m:
            level = len(m.group(1).split('.'))
            if style_val not in style_level_map:
                style_level_map[style_val] = level

    # 2. 扫描文档中的标题样式名称（从styles.xml）
    styles_xml = os.path.join(unpacked_dir, 'word', 'styles.xml')
    if os.path.exists(styles_xml):
        styles_tree = etree.parse(styles_xml)
        styles_root = styles_tree.getroot()
        for style in styles_root.findall('.//w:style', NS):
            style_id = style.get(f'{{{NS["w"]}}}styleId', '')
            name_elem = style.find('w:name', NS)
            style_name = name_elem.get(f'{{{NS["w"]}}}val', '') if name_elem is not None else ''
            if 'heading' in style_name.lower() or '标题' in style_name:
                for m in re.finditer(r'(\d+)', style_name):
                    level = int(m.group(1))
                    style_level_map[style_id] = level
                    break

    # 3. 也检查numbering.xml来识别标题编号
    # (简化处理: 通过样式名和编号模式双重匹配)

    # 如果没找到映射，尝试常见默认值
    if not style_level_map:
        # Word中默认: heading1 -> style val可能是各种值
        # 检测文档中所有非正文段落，按出现频率推断
        para_style_count = {}
        for elem in all_elements:
            if elem.tag != f'{{{NS["w"]}}}p':
                continue
            ppr = elem.find('w:pPr', NS)
            if ppr is None:
                continue
            style_elem = ppr.find('w:pStyle', NS)
            if style_elem is None:
                continue
            style_val = style_elem.get(f'{{{NS["w"]}}}val', '')
            text = ''.join(t.text or '' for t in elem.findall('.//w:t', NS)).strip()
            m = re.match(r'^(\d+(?:\.\d+)*)\s+', text)
            if m:
                level = len(m.group(1).split('.'))
                key = (style_val, level)
                para_style_count[key] = para_style_count.get(key, 0) + 1
        for (sv, lv), count in sorted(para_style_count.items(), key=lambda x: -x[1]):
            if sv not in style_level_map:
                style_level_map[sv] = lv

    if not style_level_map:
        print("[警告] 无法自动识别标题样式层级，将尝试按编号模式推断")

    # 4. 构建文档结构树
    doc_structure = []
    for i, elem in enumerate(all_elements):
        if elem.tag != f'{{{NS["w"]}}}p':
            continue

        ppr = elem.find('w:pPr', NS)
        style_elem = ppr.find('w:pStyle', NS) if ppr is not None else None
        style_val = style_elem.get(f'{{{NS["w"]}}}val', '') if style_elem is not None else ''

        text = ''.join(t.text or '' for t in elem.findall('.//w:t', NS)).strip()
        if not text:
            continue

        # 判断是否是标题
        level = None
        if style_val in style_level_map:
            level = style_level_map[style_val]

        if level is None:
            # 通过编号模式推断
            m = re.match(r'^(\d+(?:\.\d+)*)\s+(.+)', text)
            if m:
                level = len(m.group(1).split('.'))

        if level is None or level < 1:
            continue

        # 检查后面是否有占位符段落
        has_placeholder = False
        placeholder_text = ''
        if i + 1 < len(all_elements):
            next_elem = all_elements[i + 1]
            if next_elem.tag == f'{{{NS["w"]}}}p':
                next_text = ''.join(t.text or '' for t in next_elem.findall('.//w:t', NS)).strip()
                if placeholder_keyword in next_text:
                    has_placeholder = True
                    placeholder_text = next_text

        doc_structure.append({
            'idx': i,
            'level': level,
            'text': text,
            'style_val': style_val,
            'has_placeholder': has_placeholder,
            'placeholder': placeholder_text,
            'needs_content': has_placeholder,
        })

    # 5. 识别叶子节点（没有更低级别标题跟在后面的节点）
    for i, node in enumerate(doc_structure):
        if not node['has_placeholder']:
            continue
        # 检查下一个标题是否是更深层级
        is_leaf = True
        if i + 1 < len(doc_structure):
            next_node = doc_structure[i + 1]
            if next_node['level'] > node['level']:
                is_leaf = False
        node['is_leaf'] = is_leaf

    # 6. 构建树形结构
    def build_tree(nodes):
        """将扁平列表转为嵌套树"""
        root_children = []
        stack = [(0, root_children)]  # (level, children_list)

        for node in nodes:
            children = []
            entry = {**node, 'children': children}

            # 找到父级
            while stack and stack[-1][0] >= node['level']:
                stack.pop()

            if stack:
                stack[-1][1].append(entry)
            else:
                root_children.append(entry)

            if node['has_placeholder']:
                stack.append((node['level'], children))

        return root_children

    tree = build_tree(doc_structure)

    return {
        'flat': doc_structure,
        'tree': tree,
        'style_level_map': style_level_map,
        'stats': {
            'total_headings': len(doc_structure),
            'total_levels': len(set(n['level'] for n in doc_structure)),
            'placeholders_found': sum(1 for n in doc_structure if n['has_placeholder']),
            'leaf_nodes': sum(1 for n in doc_structure if n.get('is_leaf')),
        }
    }


def main():
    parser = argparse.ArgumentParser(description='从docx中提取标书框架结构')
    parser.add_argument('input_docx', help='招标文件docx路径')
    parser.add_argument('--placeholder', default='此处填写', help='占位符关键词')
    parser.add_argument('--unpacked-dir', default=None, help='解包目录')
    parser.add_argument('--output', default='structure.json', help='输出JSON路径')
    parser.add_argument('--min-level', type=int, default=0, help='最低标题级别')
    args = parser.parse_args()

    docx_path = args.input_docx
    if not os.path.exists(docx_path):
        print(f"[错误] 文件不存在: {docx_path}")
        sys.exit(1)

    # 解包
    unpacked_dir = args.unpacked_dir or tempfile.mkdtemp(prefix='bid_unpack_')
    print(f"解包目录: {unpacked_dir}")

    if not unpack_docx(docx_path, unpacked_dir):
        print("[错误] 解包失败")
        sys.exit(1)

    print("解包成功，分析结构...")

    # 分析
    result = analyze_structure(unpacked_dir, args.placeholder)
    if result is None:
        print("[错误] 结构分析失败")
        sys.exit(1)

    # 过滤最低级别
    if args.min_level > 0:
        result['flat'] = [n for n in result['flat'] if n['level'] <= args.min_level]

    # 输出
    output_path = args.output
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    # 打印摘要
    stats = result['stats']
    print(f"\n=== 结构分析结果 ===")
    print(f"总标题数: {stats['total_headings']}")
    print(f"层级数: {stats['total_levels']}")
    print(f"占位符章节: {stats['placeholders_found']}")
    print(f"叶子节点: {stats['leaf_nodes']}")
    print(f"样式映射: {result['style_level_map']}")
    print(f"\n结构已保存到: {output_path}")

    # 打印树形概览
    print(f"\n=== 目录结构 ===")
    for node in result['flat']:
        indent = '  ' * (node['level'] - 1)
        marker = '📝' if node['has_placeholder'] else '📌'
        if node.get('is_leaf'):
            marker = '🟢'
        print(f"{indent}{marker} [{node['level']}] {node['text'][:60]}")


if __name__ == '__main__':
    main()
