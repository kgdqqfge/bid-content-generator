"""
将生成的务虚内容正确回填到docx文档中。

关键设计原则（基于实际经验总结）：
1. 替换策略：找到标题后面紧跟的占位段落，替换其文本内容
2. 正文格式：保持模板中的格式属性，只替换文本
3. 多段落支持：生成内容包含换行时，拆分为多个段落并添加到占位段落后面
4. 样式安全：不创建新的样式引用，使用与占位段落相同的格式

用法:
  python fill_docx.py <template.docx> <generated_content.json> <output.docx> \
    [--placeholder "此处填写"] \
    [--unpacked-dir /tmp/unpacked] \
    [--body-font "宋体"] \
    [--body-size 10.5] \
    [--line-spacing 1.5] \
    [--first-line-indent 2]

参数说明:
  template.docx       - 模板docx路径（含占位符的原始文档）
  generated_content   - generate_content.py输出的内容JSON
  output.docx         - 输出docx路径
  --placeholder       - 占位符关键词（默认: "此处填写"）
  --unpacked-dir      - 解包临时目录
  --body-font         - 正文字体（默认: 宋体）
  --body-size         - 正文字号pt（默认: 10.5）
  --line-spacing      - 行距倍数（默认: 1.5）
  --first-line-indent - 首行缩进字符数（默认: 2）
  --validate          - 是否验证打包结果（默认: false，中文路径建议关闭）
"""

import sys, os, json, copy, re, argparse, tempfile, subprocess
sys.stdout.reconfigure(encoding='utf-8')
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
XML_NS = 'http://www.w3.org/XML/1998/namespace'


def wtag(name):
    return f'{{{W}}}{name}'


def qn(ns, name):
    return f'{{{ns}}}{name}'


def find_pack_script():
    """查找pack.py脚本路径"""
    candidates = [
        os.path.expanduser(
            '~/.workbuddy/plugins/marketplaces/codebuddy-plugins-official/'
            'plugins/docx/scripts/office/pack.py'
        ),
    ]
    import glob
    candidates += glob.glob(
        r'C:\Users\*\.workbuddy\plugins\marketplaces\codebuddy-plugins-official'
        r'\plugins\docx\scripts\office\pack.py'
    ) + glob.glob(
        r'D:\ProgramData\WorkBuddy\resources\app\extensions\genie\out\extension'
        r'\builtin\docx\scripts\office\pack.py'
    )
    for p in candidates:
        if os.path.exists(p):
            return p
    return None


def find_unpack_script():
    """查找unpack.py脚本路径"""
    candidates = [
        os.path.expanduser(
            '~/.workbuddy/plugins/marketplaces/codebuddy-plugins-official/'
            'plugins/docx/scripts/office/unpack.py'
        ),
    ]
    import glob
    candidates += glob.glob(
        r'C:\Users\*\.workbuddy\plugins\marketplaces\codebuddy-plugins-official'
        r'\plugins\docx\scripts\office\unpack.py'
    ) + glob.glob(
        r'D:\ProgramData\WorkBuddy\resources\app\extensions\genie\out\extension'
        r'\builtin\docx\scripts\office\unpack.py'
    )
    for p in candidates:
        if os.path.exists(p):
            return p
    return None


def unpack_docx(docx_path, unpacked_dir):
    """解包docx"""
    script = find_unpack_script()
    if not script:
        print("[错误] 找不到unpack.py，请确保docx skill已安装")
        return False
    os.makedirs(unpacked_dir, exist_ok=True)
    result = subprocess.run(
        [sys.executable, '-u', script, os.path.abspath(docx_path), os.path.abspath(unpacked_dir)],
        capture_output=True
    )
    return result.returncode == 0


def pack_docx(unpacked_dir, output_path, original_path, validate=False):
    """打包为docx"""
    script = find_pack_script()
    if not script:
        print("[错误] 找不到pack.py，请确保docx skill已安装")
        return False
    cmd = [sys.executable, '-u', script, os.path.abspath(unpacked_dir), os.path.abspath(output_path),
           '--original', os.path.abspath(original_path)]
    if not validate:
        cmd.extend(['--validate', 'false'])
    result = subprocess.run(cmd, capture_output=True)
    out = result.stdout.decode('utf-8', errors='replace') if result.stdout else ''
    return result.returncode == 0 or 'Successfully' in out


def make_rpr(font_name='宋体', font_size_pt=10.5):
    """创建run格式属性"""
    half_pt = int(font_size_pt * 2)  # docx用半磅
    rpr = etree.Element(wtag('rPr'))
    rfonts = etree.SubElement(rpr, wtag('rFonts'))
    rfonts.set(qn(W, 'ascii'), 'Times New Roman')
    rfonts.set(qn(W, 'hAnsi'), 'Times New Roman')
    rfonts.set(qn(W, 'eastAsia'), font_name)
    rfonts.set(qn(W, 'cs'), 'Times New Roman')
    sz = etree.SubElement(rpr, wtag('sz'))
    sz.set(qn(W, 'val'), str(half_pt))
    szCs = etree.SubElement(rpr, wtag('szCs'))
    szCs.set(qn(W, 'val'), str(half_pt))
    return rpr


def make_content_ppr(line_spacing=1.5, first_line_indent=2, font_size_pt=10.5):
    """创建段落格式属性"""
    ppr = etree.Element(wtag('pPr'))
    # 行距: 1.5倍 = 360 * 1.5 / 1 = 240 (实际Word中用twips)
    # 1倍行距 = 240, 1.5倍 = 360
    spacing = etree.SubElement(ppr, wtag('spacing'))
    line_val = str(int(240 * line_spacing))
    spacing.set(qn(W, 'line'), line_val)
    spacing.set(qn(W, 'lineRule'), 'auto')
    # 首行缩进: 每字符约21磅 = 420 twips for 宋体10.5pt
    # 简化: 2字符 ≈ 420 twips
    twips_per_char = int(font_size_pt * 20 * first_line_indent)  # 近似值
    ind = etree.SubElement(ppr, wtag('ind'))
    ind.set(qn(W, 'firstLine'), str(twips_per_char))
    jc = etree.SubElement(ppr, wtag('jc'))
    jc.set(qn(W, 'val'), 'left')
    return ppr


def make_content_para(text, font_name='宋体', font_size_pt=10.5, line_spacing=1.5, first_line_indent=2):
    """创建一个正文段落元素"""
    para = etree.Element(wtag('p'))
    ppr = make_content_ppr(line_spacing, first_line_indent, font_size_pt)
    para.insert(0, ppr)
    run = etree.SubElement(para, wtag('r'))
    rpr = make_rpr(font_name, font_size_pt)
    run.insert(0, rpr)
    t = etree.SubElement(run, wtag('t'))
    t.set(qn(XML_NS, 'space'), 'preserve')
    t.text = text
    return para


def main():
    parser = argparse.ArgumentParser(description='将生成内容回填到docx文档')
    parser.add_argument('template_docx', help='模板docx路径')
    parser.add_argument('content_json', help='生成的内容JSON路径')
    parser.add_argument('output_docx', help='输出docx路径')
    parser.add_argument('--placeholder', default='此处填写', help='占位符关键词')
    parser.add_argument('--unpacked-dir', default=None, help='解包目录')
    parser.add_argument('--body-font', default='宋体', help='正文字体')
    parser.add_argument('--body-size', type=float, default=10.5, help='正文字号pt')
    parser.add_argument('--line-spacing', type=float, default=1.5, help='行距倍数')
    parser.add_argument('--first-line-indent', type=int, default=2, help='首行缩进字符数')
    parser.add_argument('--validate', action='store_true', help='验证打包结果')
    args = parser.parse_args()

    template_path = args.template_docx
    content_path = args.content_json
    output_path = args.output_docx

    if not os.path.exists(template_path):
        print(f"[错误] 模板文件不存在: {template_path}")
        sys.exit(1)
    if not os.path.exists(content_path):
        print(f"[错误] 内容文件不存在: {content_path}")
        sys.exit(1)

    # 读取生成的内容
    with open(content_path, 'r', encoding='utf-8') as f:
        content_data = json.load(f)
    print(f"加载内容: {len(content_data)}条")

    # 构建标题到内容的映射（支持多种匹配方式）
    title_to_content = {}
    for key, text in content_data.items():
        # key格式可能是 "5.1.2.1 标题" 或 "父级>标题"
        if '>' in key:
            # "父级标题>子标题" -> 用子标题匹配
            title_text = key.split('>', 1)[1].strip()
            title_to_content[title_text] = text
        else:
            title_text = key.strip()
            title_to_content[title_text] = text
        # 也用完整key匹配
        title_to_content[key] = text

    # 解包
    unpacked_dir = args.unpacked_dir or tempfile.mkdtemp(prefix='bid_fill_')
    print(f"解包目录: {unpacked_dir}")

    if not unpack_docx(template_path, unpacked_dir):
        print("[错误] 解包失败")
        sys.exit(1)

    print("解包成功，开始填充内容...")

    # 解析XML
    doc_tree = etree.parse(os.path.join(unpacked_dir, 'word', 'document.xml'))
    root = doc_tree.getroot()
    body = root.find(f'.//{wtag("body")}')
    all_elements = list(body)

    # 遍历所有标题段落，找到后面有占位符的
    insertions = []  # (placeholder_elem, content, title_text)
    not_matched = []

    for i, elem in enumerate(all_elements):
        if elem.tag != wtag('p'):
            continue

        text = ''.join(t.text or '' for t in elem.findall(f'.//{wtag("t")}')).strip()
        if not text:
            continue

        # 检查后面是否有占位符段落
        if i + 1 < len(all_elements):
            next_elem = all_elements[i + 1]
            if next_elem.tag == wtag('p'):
                next_text = ''.join(t.text or '' for t in next_elem.findall(f'.//{wtag("t")}')).strip()
                if args.placeholder in next_text:
                    # 这个标题后面有占位符，尝试匹配内容
                    content = None
                    # 精确匹配标题文本
                    if text in title_to_content:
                        content = title_to_content[text]
                    # 模糊匹配：key以标题文本结尾
                    if content is None:
                        for key in title_to_content:
                            if key.endswith(text) and text:
                                content = title_to_content[key]
                                break

                    if content is not None:
                        insertions.append((next_elem, content, text))
                    else:
                        not_matched.append(text)

    print(f"匹配: {len(insertions)}, 未匹配: {len(not_matched)}")
    if not_matched:
        for nm in not_matched[:10]:
            print(f"  未匹配: {nm[:60]}")
        if len(not_matched) > 10:
            print(f"  ... 还有{len(not_matched)-10}个")

    if not insertions:
        print("[警告] 没有找到匹配的标题-占位符对，请检查：")
        print("  1. 占位符关键词是否正确")
        print("  2. 内容JSON中的key是否与文档标题一致")
        sys.exit(1)

    # 执行替换
    replace_count = 0
    total_extra_paras = 0

    for placeholder, content, title in insertions:
        # 将内容按换行拆分为多个段落
        paragraphs = re.split(r'\n+', content)
        paragraphs = [p.strip() for p in paragraphs if p.strip()]

        if not paragraphs:
            continue

        # 清除占位段落的原有run元素
        for r in placeholder.findall(wtag('r')):
            placeholder.remove(r)

        # 添加正文格式属性
        existing_ppr = placeholder.find(wtag('pPr'))
        if existing_ppr is not None:
            placeholder.remove(existing_ppr)
        ppr = make_content_ppr(args.line_spacing, args.first_line_indent, args.body_size)
        placeholder.insert(0, ppr)

        # 写入第一段
        run = etree.SubElement(placeholder, wtag('r'))
        rpr = make_rpr(args.body_font, args.body_size)
        run.insert(0, rpr)
        t = etree.SubElement(run, wtag('t'))
        t.set(qn(XML_NS, 'space'), 'preserve')
        t.text = paragraphs[0]

        # 追加额外段落
        current = placeholder
        for extra_text in paragraphs[1:]:
            extra_para = make_content_para(
                extra_text, args.body_font, args.body_size,
                args.line_spacing, args.first_line_indent
            )
            current.addnext(extra_para)
            current = extra_para
            total_extra_paras += 1

        replace_count += 1

    print(f"替换占位符: {replace_count}")
    print(f"追加段落: {total_extra_paras}")

    # 保存XML
    output_xml = os.path.join(unpacked_dir, 'word', 'document.xml')
    doc_tree.write(output_xml, xml_declaration=True, encoding='UTF-8', standalone=True)
    print("XML已保存")

    # 打包
    print("正在打包docx...")
    if pack_docx(unpacked_dir, output_path, template_path, validate=args.validate):
        size_mb = os.path.getsize(output_path) / 1024 / 1024
        print(f"打包成功: {output_path}")
        print(f"文件大小: {size_mb:.2f} MB")
    else:
        print("[错误] 打包失败")
        sys.exit(1)

    # 验证：统计输出文档中的占位符残留
    verify_unpack_dir = tempfile.mkdtemp(prefix='bid_verify_')
    if unpack_docx(output_path, verify_unpack_dir):
        verify_tree = etree.parse(os.path.join(verify_unpack_dir, 'word', 'document.xml'))
        verify_root = verify_tree.getroot()
        remaining = 0
        total_chars = 0
        for p in verify_root.iter(wtag('p')):
            p_text = ''.join(t.text or '' for t in p.findall(f'.//{wtag("t")}')).strip()
            if args.placeholder in p_text:
                remaining += 1
            elif p_text and not p_text.startswith(('第', '1', '2', '3', '4', '5', '6', '7', '8', '9')):
                # 粗略统计正文内容字数
                total_chars += len(p_text)

        print(f"\n=== 验证结果 ===")
        print(f"剩余占位符: {remaining}")
        print(f"估算正文总字数: {total_chars}")


if __name__ == '__main__':
    main()
