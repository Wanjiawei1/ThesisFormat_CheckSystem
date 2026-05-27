"""
浙江工业大学研究生学位论文 - 参考文献格式检测模块
严格按照 GB/T 7714-2015 顺序编码制 + 浙工大模板格式要求

检测项:
1. 标题"参考文献"格式：黑体三号(16pt)、加粗、居中、1.5倍行距、段前段后0行
2. 条目编号：[1][2]...左顶格、悬挂缩进2字符、连续性
3. 条目内容：宋体小四(12pt)、英文Times New Roman小四(12pt)、1.5倍行距
4. 标点：半角括号/逗号/冒号、后空一格、每条以"."结束
5. 作者：3名以上列前3名+", 等"/", et al"
6. 各文献类型结构验证（M/J/D/C/P/EB等）
"""

from docx import Document
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re


class GraduateReferenceChecker:
    """参考文献格式检测器"""

    # 段落格式期望值
    TITLE_FONT_SIZE = 16.0       # 三号 = 16pt
    TITLE_FONT_CN = '黑体'
    TITLE_FONT_EN = 'Times New Roman'  # 英文部分
    BODY_FONT_SIZE = 12.0        # 小四 = 12pt
    BODY_FONT_CN = '宋体'
    BODY_FONT_EN = 'Times New Roman'
    LINE_SPACING = 1.5           # 1.5倍行距
    HANGING_INDENT_CM = 2.0      # 悬挂缩进2字符 ≈ 0.74cm (12pt宋体)

    def __init__(self):
        self.type_labels = {
            'M': '专著(图书)', 'J': '期刊', 'D': '学位论文', 'C': '论文集',
            'R': '报告', 'S': '标准', 'P': '专利', 'N': '报纸',
            'EB/OL': '电子文献(在线)', 'DB/OL': '数据库(在线)',
            'J/OL': '期刊(在线)', 'M/CD': '光盘专著',
        }

    # ======== 字体属性提取 ========

    def _get_run_fonts(self, run, paragraph=None):
        """获取 run 的 ascii 字体和 eastAsia 字体（分别用于英文和中文）"""
        info = {
            'ascii': None,       # 英文字体
            'eastAsia': None,    # 中文字体（显式设置的）
            'eastAsia_set': False,  # eastAsia 是否在 XML 中显式设置
            'size': None,
            'bold': run.font.bold,
        }

        try:
            if hasattr(run._element, 'rPr') and run._element.rPr is not None:
                rpr = run._element.rPr
                rfonts = rpr.find(qn('w:rFonts'))
                if rfonts is not None:
                    info['ascii'] = rfonts.get(qn('w:ascii'))
                    ea = rfonts.get(qn('w:eastAsia'))
                    if ea:
                        info['eastAsia'] = ea
                        info['eastAsia_set'] = True
                    if not info['ascii']:
                        info['ascii'] = rfonts.get(qn('w:hAnsi'))
        except:
            pass

        # 回退
        if not info['ascii'] and run.font.name:
            info['ascii'] = run.font.name

        # 字号
        if run.font.size:
            info['size'] = run.font.size.pt
        elif paragraph:
            try:
                if paragraph.style and paragraph.style.font and paragraph.style.font.size:
                    info['size'] = paragraph.style.font.size.pt
            except:
                pass

        return info

    def _has_chinese(self, text):
        return any('\u4e00' <= c <= '\u9fff' for c in text)

    def _get_alignment_name(self, para):
        a = para.paragraph_format.alignment
        if a == WD_ALIGN_PARAGRAPH.CENTER:
            return 'CENTER'
        if a == WD_ALIGN_PARAGRAPH.LEFT:
            return 'LEFT'
        if a == WD_ALIGN_PARAGRAPH.RIGHT:
            return 'RIGHT'
        if a == WD_ALIGN_PARAGRAPH.JUSTIFY:
            return 'JUSTIFY'
        return str(a)

    def _get_line_spacing(self, para):
        try:
            pPr = para._element.find(qn('w:pPr'))
            if pPr is not None:
                spacing = pPr.find(qn('w:spacing'))
                if spacing is not None:
                    line_val = spacing.get(qn('w:line'))
                    if line_val is not None:
                        line_val = int(line_val)
                        if line_val > 1000:
                            return line_val / 12700.0
                        elif line_val > 100:
                            return line_val / 240.0
                        else:
                            return line_val / 20.0
        except:
            pass
        return None

    def _get_spacing_before_after(self, para):
        """获取段前段后间距(磅)"""
        before = para.paragraph_format.space_before
        after = para.paragraph_format.space_after
        return (
            before.pt if before else 0.0,
            after.pt if after else 0.0,
        )

    def _get_hanging_indent(self, para):
        """获取悬挂缩进(磅)"""
        try:
            # 段落格式的 leftIndent
            left = para.paragraph_format.left_indent
            first = para.paragraph_format.first_line_indent
            if left and first and first.pt < 0:
                return abs(first.pt)
        except:
            pass
        # 从 XML 检查
        try:
            pPr = para._element.find(qn('w:pPr'))
            if pPr is not None:
                ind = pPr.find(qn('w:ind'))
                if ind is not None:
                    hanging = ind.get(qn('w:hanging'))
                    if hanging:
                        return int(hanging) / 20  # twips → pt
        except:
            pass
        return 0

    # ======== 提取参考文献区域 ========

    def _find_ref_section(self, doc):
        """定位参考文献标题段落索引"""
        for i, para in enumerate(doc.paragraphs):
            text = re.sub(r'\s+', '', para.text.strip())
            if text == '参考文献':
                return i
        return -1

    def _is_section_end(self, text):
        normalized = re.sub(r'\s+', '', text)
        for t in ['致谢', '附录', '作者简介', '学位论文数据集']:
            if normalized.startswith(t):
                return True
        if re.match(r'^第[一二三四五六七八九十\d]+\s*章', text):
            return True
        return False

    def extract_references(self, docx_path):
        """提取参考文献条目及其段落对象"""
        doc = Document(docx_path)
        ref_start = self._find_ref_section(doc)
        refs = []

        if ref_start < 0:
            return refs, doc, -1

        # 检测是否使用自动编号（Word 列表编号），而非文本 [1] [2] 形式
        uses_auto_numbering = False
        for i in range(ref_start + 1, min(ref_start + 5, len(doc.paragraphs))):
            para = doc.paragraphs[i]
            if not para.text.strip():
                continue
            pPr = para._element.find(qn('w:pPr'))
            if pPr is not None and pPr.find(qn('w:numPr')) is not None:
                uses_auto_numbering = True
                break

        seq_num = 0
        for i in range(ref_start + 1, len(doc.paragraphs)):
            para = doc.paragraphs[i]
            text = para.text.strip()
            if not text:
                continue
            if refs and self._is_section_end(text):
                break

            if uses_auto_numbering:
                pPr = para._element.find(qn('w:pPr'))
                has_num = pPr is not None and pPr.find(qn('w:numPr')) is not None
                if has_num:
                    seq_num += 1
                    refs.append({
                        'number': seq_num,
                        'content': text,
                        'paragraph_index': i,
                        'paragraph': para,
                    })
                elif refs:
                    refs[-1]['content'] += ' ' + text
            else:
                m = re.match(r'^\[(\d+)\]\s*(.*)', text)
                if m:
                    refs.append({
                        'number': int(m.group(1)),
                        'content': m.group(2).strip(),
                        'paragraph_index': i,
                        'paragraph': para,
                    })
                elif refs:
                    refs[-1]['content'] += ' ' + text

        return refs, doc, ref_start

    # ======== 1. 标题格式检测 ========

    def check_title_format(self, doc, title_index):
        """检查"参考文献"标题的格式"""
        issues = []
        if title_index < 0:
            issues.append("[警告] 未找到'参考文献'标题")
            return issues

        para = doc.paragraphs[title_index]

        # 行距 → 1.5倍
        ls = self._get_line_spacing(para)
        if ls is not None and abs(ls - 1.5) > 0.1:
            issues.append(f"[错误] 标题行距应为1.5倍，当前为: {ls}")

        # 段前段后 → 0
        sp_before, sp_after = self._get_spacing_before_after(para)
        if sp_before and sp_before > 1:
            issues.append(f"[错误] 标题段前间距应为0行，当前为: {sp_before}pt")
        if sp_after and sp_after > 1:
            issues.append(f"[错误] 标题段后间距应为0行，当前为: {sp_after}pt")

        # 字体和加粗
        for run in para.runs:
            if not run.text.strip():
                continue
            fonts = self._get_run_fonts(run, para)
            has_cn = self._has_chinese(run.text)

            # 字号 → 三号=16pt
            if fonts['size'] and abs(fonts['size'] - self.TITLE_FONT_SIZE) > 1.0:
                issues.append(f"[错误] 标题字号应为三号(16pt)，当前为: {fonts['size']}pt")
                break

            # 加粗
            if fonts['bold'] is False:
                issues.append(f"[错误] 标题应为加粗，当前未加粗")
                break

            # 字体（中文字体应为黑体，仅在 eastAsia 显式设置时检查）
            if has_cn and fonts['eastAsia_set'] and fonts['eastAsia'] and '黑体' not in fonts['eastAsia']:
                issues.append(f"[错误] 标题中文字体应为黑体，当前为: {fonts['eastAsia']}")
                break

        if not issues:
            issues.append("[通过] 标题格式符合要求")

        return issues

    # ======== 2. 条目编号检测 ========

    def check_numbering(self, refs):
        """检查编号连续性和格式"""
        issues = []
        if not refs:
            return issues

        # 连续性
        for i, ref in enumerate(refs):
            expected = i + 1
            if ref['number'] != expected:
                issues.append(f"[警告] 编号不连续: [{ref['number']}] 位置应为 [{expected}]")

        # 段落格式：悬挂缩进
        for ref in refs:
            para = ref['paragraph']
            hanging = self._get_hanging_indent(para)
            if hanging < 15:  # 2字符 ≈ 24pt at 12pt font, 宽松到15pt
                issues.append(f"[警告] [{ref['number']}] 悬挂缩进不足(当前{hanging:.0f}pt，应≥约24pt)")

        if not issues:
            issues.append("[通过] 编号连续性正常")
        return issues

    # ======== 3. 条目段落格式检测 ========

    def check_item_format(self, refs):
        """检查每条文献的段落格式（字体、字号、行距）"""
        issues = []

        for ref in refs:
            para = ref['paragraph']
            num = ref['number']
            item_issues = []

            # 行距 → 1.5
            ls = self._get_line_spacing(para)
            if ls is not None and abs(ls - 1.5) > 0.15:
                item_issues.append(f"行距应为1.5倍，当前为: {ls}")

            # 段前段后 → 0
            sp_before, sp_after = self._get_spacing_before_after(para)
            if sp_before and sp_before > 2:
                item_issues.append(f"段前间距应为0，当前: {sp_before}pt")
            if sp_after and sp_after > 2:
                item_issues.append(f"段后间距应为0，当前: {sp_after}pt")

            # 检查 run 的字体和字号
            for run in para.runs:
                if not run.text.strip():
                    continue
                fonts = self._get_run_fonts(run, para)
                has_cn = self._has_chinese(run.text)
                has_en = bool(re.search(r'[a-zA-Z]', run.text))

                # 字号 → 小四=12pt
                if fonts['size'] and abs(fonts['size'] - self.BODY_FONT_SIZE) > 1.0:
                    item_issues.append(f"字号应为小四(12pt)，当前为: {fonts['size']}pt")

                # 中文字体 → 宋体 (仅在 eastAsia 显式设置时检查)
                if has_cn and fonts['eastAsia_set'] and fonts['eastAsia'] and '宋体' not in fonts['eastAsia']:
                    item_issues.append(f"中文字体应为宋体，当前为: {fonts['eastAsia']}")

                # 英文字体 → Times New Roman (查 ascii)
                if has_en and fonts['ascii'] and 'Times' not in fonts['ascii'] and 'times' not in fonts['ascii'].lower():
                    item_issues.append(f"英文字体应为Times New Roman，当前为: {fonts['ascii']}")

            if item_issues:
                issues.append(f"[{num}] " + "; ".join(item_issues))

        return issues

    # ======== 4. 标点和格式检测 ========

    def check_punctuation(self, refs):
        """检查标点符号和格式"""
        issues = []

        for ref in refs:
            num = ref['number']
            content = ref['content']
            if not content:
                issues.append(f"[{num}] 内容为空")
                continue

            item_issues = []

            # 每条应以 "." 结束
            if not content.rstrip().endswith('.'):
                item_issues.append("应以句号 '.' 结束")

            # 全角标点检查（括号、逗号、冒号应为半角）
            fullwidth_punct = re.findall(r'[，：；（）]', content)
            if fullwidth_punct:
                # 但年份中的中文括号如（2003）可能被允许，这里只报标点
                item_issues.append(f"存在全角标点（应使用半角）: {''.join(set(fullwidth_punct))}")

            # 检查文献类型标识
            type_match = re.search(r'\[([A-Z]+(?:/[A-Z]+)?)\]', content)
            if not type_match:
                item_issues.append("缺少文献类型标识，如 [M]、[J]、[D] 等")
            else:
                ref_type = type_match.group(1)
                if ref_type not in self.type_labels:
                    item_issues.append(f"未知文献类型标识: [{ref_type}]")

            # 类型标识前应有 ". " 分隔（仅当有作者时才检查）
            # 无作者的参考文献（如网络公告、标准等）不检查此项
            # 判断方式：
            # 1. 类型标识前没有内容（以[开头）→ 无作者
            # 2. 类型标识前只有标点、空格、数字 → 无作者
            # 3. 类型标识前内容以书名号《》开头 → 无作者（以题名开头）
            # 4. 类型标识前内容包含题名特征词（关于、论文、研究等）→ 无作者
            type_pos = content.find('[')
            if type_pos > 0:
                before = content[:type_pos].strip()
                # 如果类型标识前没有实质内容（只有标点、空格等）→ 无作者
                is_no_author = not before or bool(re.fullmatch(r'[\s.．,，（）()\[\]【】\d]*', before))
                # 如果以书名号开头 → 无作者（以题名开头）
                if before.startswith('《'):
                    is_no_author = True
                # 如果包含题名特征词 → 无作者（以题名开头，而非作者名）
                title_keywords = ['关于', '论文', '研究', '规范', '要求', '标准', '指南', '报告', '分析', '设计', '实现', '方法', '技术', '系统', '应用', '综述', '进展', '现状', '趋势']
                if any(kw in before for kw in title_keywords):
                    is_no_author = True
                # 如果内容包含". "，说明可能是"作者. 题名"格式，是有作者的
                # 但如果只是"题名. 期刊名"格式（无作者），则不检查
                if not is_no_author and not re.search(r'[.．]\s', before):
                    item_issues.append("作者与题名/类型标识之间应以 '. ' 分隔")

            # 类型标识后应有 "."
            if type_match and not re.search(r'\[[A-Z]+(?:/[A-Z]+)?\]\s*\.', content):
                item_issues.append("类型标识后应有句号，如 [M].")

            # 标点后缺少空格检查
            missing_space_checks = [
                (r'\.(?=[A-Z])', '句点后缺少空格'),
                (r',(?=[a-zA-Z0-9])', '逗号后缺少空格'),
                (r'(?<!http)(?<!https)(?<!ftp):(?=[A-Z])', '冒号后缺少空格'),
                (r';(?=[a-zA-Z])', '分号后缺少空格'),
                (r'\](?=[A-Z])', '右括号后缺少空格'),
            ]
            for pat, desc in missing_space_checks:
                for m in re.finditer(pat, content):
                    ctx = content[max(0, m.start()-10):m.end()+10]
                    if re.search(r'\.\d+', m.group()):
                        continue
                    if desc == '句点后缺少空格' and re.search(r'\.(?=\d)', content[m.start():m.start()+3]):
                        continue
                    item_issues.append(f"{desc} (\"...{ctx}...\")")
                    break

            if item_issues:
                issues.append(f"[{num}] " + "; ".join(item_issues))

        return issues

    # ======== 5. 作者格式检测 ========

    def check_authors(self, refs):
        """检查作者格式（3人以上、姓前名后等）"""
        issues = []

        for ref in refs:
            num = ref['number']
            content = ref['content']

            # 提取作者部分（类型标识前的内容）
            type_pos = content.find('[')
            author_part = content[:type_pos].strip() if type_pos > 0 else content

            # 去掉末尾句点
            author_part = author_part.rstrip('.').strip()

            # 统计作者数量（按逗号分隔）
            if not author_part:
                continue

            # 中文作者
            has_cn = self._has_chinese(author_part)
            if has_cn:
                # 中文用逗号分隔
                authors = [a.strip() for a in re.split(r'[,，]', author_part) if a.strip()]
                if len(authors) > 3:
                    # 应该是 "前3名, 等"
                    if '等' not in author_part:
                        issues.append(f"[{num}] 作者超过3人应缩写为'前3名, 等'，当前{len(authors)}人")
            else:
                # 英文作者
                authors = [a.strip() for a in re.split(r',', author_part) if a.strip()]
                if len(authors) > 3:
                    if 'et al' not in author_part.lower():
                        issues.append(f"[{num}] 作者超过3人应缩写为'前3名, et al'，当前{len(authors)}人")

        return issues

    # ======== 6. 按类型检查结构 ========

    def check_type_structure(self, refs):
        """按文献类型检查基本结构"""
        issues = []

        for ref in refs:
            num = ref['number']
            content = ref['content']
            type_match = re.search(r'\[([A-Z]+(?:/[A-Z]+)?)\]', content)
            if not type_match:
                continue

            ref_type = type_match.group(1)
            item_issues = []

            if ref_type == 'M':  # 专著
                if not re.search(r'[;:：]', content.split(']')[-1]):
                    item_issues.append("专著缺少出版信息（应含出版地:出版社）")
                if not re.search(r'\d{4}', content):
                    item_issues.append("专著缺少出版年份")

            elif ref_type == 'J':  # 期刊
                if not re.search(r'[,，]\s*\d{4}', content):
                    # 年份前标点不是逗号，也可能是格式问题（如用了句号）
                    if re.search(r'[.．]\s*\d{4}', content):
                        item_issues.append("期刊年份前应使用逗号 ',' 而非句号 '.'")
                    else:
                        item_issues.append("期刊缺少年份")
                # 检查卷(期)号格式
                # 支持格式：卷(期)如34(04)、只有期如(06)、卷:页码如34:251
                has_volume_issue = False
                if re.search(r'\d+\(\d+\)', content):
                    # 有 卷(期) 格式，正确
                    pass
                elif re.search(r'\(\d+\)', content):
                    # 只有 (期) 格式，也允许
                    pass
                elif re.search(r'\d+\s*:', content):
                    # 有 卷: 页码 格式，正确
                    pass
                else:
                    item_issues.append("期刊缺少卷(期)号，应为 卷(期) 或 (期) 格式")
                # 检查页码：支持起止页码(123-125)、单页码(: 97)、文章编号
                has_page_info = False
                # 起止页码
                if re.search(r'\d+[-–—]\d+', content):
                    has_page_info = True
                # 单页码：卷(期): 页码 或 卷(期):页码
                elif re.search(r'\(\d+\)\s*:\s*\d+', content):
                    has_page_info = True
                # 文章编号格式：卷(期) 数字（无冒号）
                elif re.search(r'\d+\(\d+\)\s+\d{3,}', content):
                    has_page_info = True
                if not has_page_info:
                    item_issues.append("期刊缺少页码信息（起止页码、单页码或文章编号）")

            elif ref_type == 'D':  # 学位论文
                has_unit = bool(re.search(
                    r'[\u4e00-\u9fa5]+大学|[\u4e00-\u9fa5]+学院|Univ\.|University',
                    content
                ))
                if not has_unit:
                    item_issues.append("学位论文缺少授予/保存单位")
                if not re.search(r'\d{4}', content):
                    item_issues.append("学位论文缺少年份")

            elif ref_type == 'C':  # 会议论文集
                if not re.search(r'[.]\s*//', content) and '//' not in content:
                    item_issues.append("会议论文集应包含 '//' 编者/文集名分隔符")
                if not re.search(r'\d+[-–—]\d+', content):
                    # 允许文章编号替代起止页码
                    if not re.search(r'[:\s]\d{3,}', content):
                        item_issues.append("会议论文集缺少页码（或文章编号）")

            elif ref_type == 'P':  # 专利
                if not re.search(r'[\d\.]+.*\[P\]|CN\s*[\d]+', content):
                    item_issues.append("专利缺少专利号")

            elif ref_type in ('EB/OL', 'DB/OL', 'J/OL'):  # 电子文献
                if not re.search(r'https?[:\s/]+|www\.|[a-zA-Z0-9-]+\.[a-z]{2,}', content):
                    item_issues.append("电子文献缺少URL链接")

            if item_issues:
                issues.append(f"[{num}] " + "; ".join(item_issues))

        return issues

    # ======== 引用匹配检测 ========

    def _expand_citation_range(self, num_str: str) -> set:
        """展开引用范围字符串，如 '4-7' -> {4,5,6,7}, '2,3,5-8' -> {2,3,5,6,7,8}"""
        result = set()
        # 先按逗号分隔（支持半角和中文逗号）
        parts = re.split(r'[,，]', num_str)
        for part in parts:
            part = part.strip()
            if not part:
                continue
            # 范围：支持 - – — 三种破折号
            range_match = re.match(r'^(\d+)\s*[-–—]\s*(\d+)$', part)
            if range_match:
                start, end = int(range_match.group(1)), int(range_match.group(2))
                if start <= end:
                    result.update(range(start, end + 1))
            else:
                # 单个编号
                single = re.match(r'^(\d+)$', part)
                if single:
                    result.add(int(single.group(1)))
        return result

    def check_citation_matching(self, doc: Document, refs: list) -> dict:
        """验证正文引用与参考文献列表的匹配关系"""
        # 定位正文区域：从第1章到参考文献
        body_start = -1
        body_end = len(doc.paragraphs)
        for i, para in enumerate(doc.paragraphs):
            text = re.sub(r'\s+', '', para.text.strip())
            if body_start < 0 and re.match(r'^第[1一]章', text):
                body_start = i
            if re.match(r'^参考文献$', text):
                body_end = i
                break

        if body_start < 0:
            return {
                "total_citations_found": 0,
                "issues": ["[警告] 未能定位正文区域，无法进行引用匹配检测"]
            }

        # 提取正文中所有引用标记
        cited_numbers = set()
        for i in range(body_start, body_end):
            text = doc.paragraphs[i].text
            # 匹配 [数字] 或 [数字,数字] 或 [数字-数字] 格式
            matches = re.findall(r'\[(\d+(?:\s*[,，\-–—]\s*\d+)*)\]', text)
            for m in matches:
                # 过滤形如 [1.2]、[2.3.4] 的非引用标记
                if re.match(r'^\d+\.\d+', m):
                    continue
                cited_numbers.update(self._expand_citation_range(m))

        # 构建参考文献列表编号集合
        ref_list_numbers = {ref['number'] for ref in refs}

        refs_not_cited = sorted(ref_list_numbers - cited_numbers)
        citations_not_in_refs = sorted(cited_numbers - ref_list_numbers)

        issues = []
        for n in refs_not_cited:
            # 找到对应的参考文献内容预览
            ref_preview = ""
            for ref in refs:
                if ref['number'] == n:
                    ref_preview = ref['content'][:60]
                    break
            issues.append(f"[警告] 参考文献[{n}]在正文中未被引用 (内容: \"{ref_preview}...\")")

        for n in citations_not_in_refs:
            issues.append(f"[警告] 正文引用了[{n}]但参考文献列表中不存在")

        return {
            "total_citations_found": len(cited_numbers),
            "total_refs_in_list": len(ref_list_numbers),
            "refs_not_cited": refs_not_cited,
            "citations_not_in_refs": citations_not_in_refs,
            "issues": issues,
        }

    # ======== 主入口 ========

    def check_references(self, docx_path):
        """执行全部检测，返回结构化结果"""
        refs, doc, title_idx = self.extract_references(docx_path)

        citation_result = self.check_citation_matching(doc, refs)

        results = {
            'total': len(refs),
            'has_title': title_idx >= 0,
            'refs': refs,
            'title_issues': self.check_title_format(doc, title_idx),
            'numbering_issues': self.check_numbering(refs),
            'format_issues': self.check_item_format(refs),
            'punct_issues': self.check_punctuation(refs),
            'author_issues': self.check_authors(refs),
            'type_issues': self.check_type_structure(refs),
            'citation_issues': citation_result.get('issues', []),
        }

        # 统计文献类型
        type_stats = {}
        for ref in refs:
            m = re.search(r'\[([A-Z]+(?:/[A-Z]+)?)\]', ref['content'])
            t = m.group(1) if m else '未识别'
            type_stats[t] = type_stats.get(t, 0) + 1
        results['type_stats'] = type_stats

        return results

    # ======== 报告生成 ========

    def generate_report(self, docx_path, report_path=None):
        import io, sys
        buf = io.StringIO()
        _orig = sys.stdout
        sys.stdout = buf

        r = self.check_references(docx_path)
        lines = []

        lines.append("# 参考文献格式检测报告\n")
        lines.append("## 基本信息")
        lines.append(f"- 参考文献总数: {r['total']}")
        lines.append("")

        # 1. 标题
        lines.append("## 1. 标题格式检测")
        for it in r['title_issues']:
            lines.append(f"- {it}")
        lines.append("")

        # 2. 编号
        lines.append("## 2. 编号连续性")
        for it in r['numbering_issues']:
            lines.append(f"- {it}")
        lines.append("")

        # 3. 段落格式
        lines.append("## 3. 段落格式检测（字体/字号/行距）")
        if not r['format_issues']:
            lines.append("[通过] 所有条目段落格式符合要求")
        else:
            lines.append(f"[警告] {len(r['format_issues'])} 条存在问题:")
            for it in r['format_issues']:
                lines.append(f"- [警告] {it}")
        lines.append("")

        # 4. 标点
        lines.append("## 4. 标点与格式检测")
        if not r['punct_issues']:
            lines.append("[通过] 标点格式符合要求")
        else:
            lines.append(f"[警告] {len(r['punct_issues'])} 条存在问题:")
            for it in r['punct_issues']:
                lines.append(f"- [警告] {it}")
        lines.append("")

        # 5. 作者
        lines.append("## 5. 作者格式检测")
        if not r['author_issues']:
            lines.append("[通过] 作者格式符合要求")
        else:
            for it in r['author_issues']:
                lines.append(f"- [警告] {it}")
        lines.append("")

        # 6. 类型结构
        lines.append("## 6. 文献类型结构检测")
        if not r['type_issues']:
            lines.append("[通过] 文献类型结构符合要求")
        else:
            for it in r['type_issues']:
                lines.append(f"- [警告] {it}")
        lines.append("")

        # 7. 引用匹配
        lines.append("## 7. 引用匹配检测")
        ci = r.get('citation_issues', [])
        if not ci:
            lines.append("[通过] 引用匹配检测通过")
        else:
            lines.append(f"[警告] **发现 {len(ci)} 个引用匹配问题**")
            for it in ci:
                lines.append(f"- {it}")
        lines.append("")

        # 类型统计
        if r['type_stats']:
            lines.append("## 文献类型统计")
            for t, c in sorted(r['type_stats'].items(), key=lambda x: -x[1]):
                label = self.type_labels.get(t, '未知')
                lines.append(f"- [{t}] {label}: {c} 篇")
            lines.append("")

        # 结论
        all_issues = (r['numbering_issues'] + r['format_issues'] +
                      r['punct_issues'] + r['author_issues'] + r['type_issues'] +
                      r.get('citation_issues', []))
        # 去掉通过类消息计数（只保留警告和错误）
        real_issues = [i for i in all_issues if i.startswith('[警告]') or i.startswith('[错误]')]
        lines.append("## 结论")
        if not r['has_title']:
            lines.append("[警告] 未找到'参考文献'标题，无法检测")
        elif r['total'] == 0:
            lines.append("[信息] 暂未检测到参考文献条目")
        elif not real_issues:
            lines.append("[通过] 参考文献格式检测通过")
        else:
            lines.append(f"**共发现 {len(real_issues)} 个问题**")

        report = "\n".join(lines)

        sys.stdout = _orig

        if report_path:
            with open(report_path, 'w', encoding='utf-8') as f:
                f.write(report)

        print(report)
        return report


def main():
    import os, sys
    if len(sys.argv) < 2:
        print("用法: python reference_checker.py <docx文件路径>")
        sys.exit(1)
    path = sys.argv[1]
    if not os.path.exists(path):
        print(f"[错误] 文件不存在: {path}")
        sys.exit(1)
    checker = GraduateReferenceChecker()
    d = os.path.join(os.path.dirname(path), "reports")
    os.makedirs(d, exist_ok=True)
    checker.generate_report(path, os.path.join(d, "ref_report.md"))


if __name__ == "__main__":
    main()
