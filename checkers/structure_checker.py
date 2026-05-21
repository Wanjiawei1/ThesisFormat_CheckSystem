from docx import Document
from docx.oxml.ns import qn
import os
import re


def roman_to_int(s: str) -> int | None:
    """将罗马数字（I, II, IV...）转成 int；不合法返回 None。"""
    if not s:
        return None
    s = s.strip().upper()
    if not re.fullmatch(r"[IVXLCDM]+", s):
        return None
    values = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100, "D": 500, "M": 1000}
    total = 0
    prev = 0
    for ch in reversed(s):
        v = values[ch]
        if v < prev:
            total -= v
        else:
            total += v
            prev = v
    return total


class ThesisStructureChecker:
    """
    论文整体结构检测：
    1) 检查章节顺序（基于目录项/标题识别）
    2) 验证页码连续性（基于目录页码：罗马/阿拉伯分别检测递增与连续）
    """

    def __init__(self):
        # 允许缺失（不同学院模板不一定都有），但若出现则检查相对顺序
        # 用“同义组”避免：同一章节在目录里写成“摘  要/摘要”导致自相矛盾
        self.expected_groups = [
            ["摘要", "摘  要", "摘 要"],
            ["ABSTRACT"],
            ["目录", "目  录", "目 录"],
            ["第1章"],  # 前缀匹配
            ["参考文献"],
            ["致谢"],
            ["附录"],  # 前缀匹配
        ]

    def _normalize_title(self, t: str) -> str:
        t = re.sub(r"\s+", "", t.strip())
        return t

    def _parse_toc_line(self, text: str):
        """
        解析目录行：返回 (title, page_raw, page_type, page_value) 或 None
        支持：
        - 标题 \t 页码
        - 标题 ... 页码
        """
        raw = text.strip()
        if not raw:
            return None

        # 常见：标题\t页码
        if "\t" in raw:
            left, right = raw.rsplit("\t", 1)
            title = left.strip()
            page = right.strip()
        else:
            # 尝试：标题 .... 12 / I
            m = re.match(r"^(.*?)[\.\·…\s]{2,}([IVXLCDM]+|\d+)\s*$", raw, re.IGNORECASE)
            if not m:
                return None
            title = m.group(1).strip()
            page = m.group(2).strip()

        if not title or not page:
            return None

        if re.fullmatch(r"\d+", page):
            return title, page, "arabic", int(page)

        r = roman_to_int(page)
        if r is not None:
            return title, page, "roman", r

        return None

    def extract_toc_entries(self, doc: Document):
        """从文档段落中提取目录项（尽量只取目录区域）。"""
        paras = doc.paragraphs
        # 1) 找目录起点
        toc_start = None
        for i, p in enumerate(paras):
            t = self._normalize_title(p.text)
            if t in ("目录", "目 录", "目录目录"):
                toc_start = i
                break
        if toc_start is None:
            # 没有明确目录标题：退化为全局扫描“标题\t页码”模式
            toc_start = 0

        entries = []
        in_toc = toc_start == 0
        non_parsable_streak = 0

        for i in range(toc_start, len(paras)):
            raw = paras[i].text
            tnorm = self._normalize_title(raw)

            if not in_toc:
                # 跳过“目录”标题行本身
                if tnorm in ("目录", "目录", "目录目录", "目 录"):
                    in_toc = True
                continue

            parsed = self._parse_toc_line(raw)
            if parsed:
                entries.append(parsed)
                non_parsable_streak = 0
                continue

            # 没解析出来：可能是空行/目录说明/真正正文
            if not raw.strip():
                # 空行不累计
                continue

            non_parsable_streak += 1

            # 经验规则：如果已经解析到一些目录项，并且连续多行都不像目录项，
            # 同时出现“第X章/参考文献/致谢/附录”这种正文标题形态（且没有页码），认为目录结束。
            if entries and non_parsable_streak >= 3:
                if re.match(r"^(第\s*[一二三四五六七八九十\d]+\s*章|参\s*考\s*文\s*献|致\s*谢|附\s*录)", raw.strip()):
                    break

        return entries

    def check_order(self, toc_entries):
        """基于目录项检查关键章节相对顺序。"""
        titles = [self._normalize_title(t) for (t, _p, _pt, _pv) in toc_entries]

        # 将 expected_groups 映射到出现位置（只对出现的组检查顺序）
        positions = []
        for group in self.expected_groups:
            idx = None
            picked = None
            for exp in group:
                exp_norm = self._normalize_title(exp)
                if exp_norm == "第1章":
                    idx = next((i for i, tt in enumerate(titles) if tt.startswith("第1章")), None)
                elif exp_norm == "附录":
                    idx = next((i for i, tt in enumerate(titles) if tt.startswith("附录")), None)
                else:
                    idx = next((i for i, tt in enumerate(titles) if tt == exp_norm), None)
                if idx is not None:
                    picked = exp
                    break
            if idx is not None:
                positions.append((picked, idx))

        issues = []
        for (a, ia), (b, ib) in zip(positions, positions[1:]):
            if ia >= ib:
                issues.append(f"章节顺序异常：`{a}` 应在 `{b}` 之前（目录位置 {ia} vs {ib}）")

        return {
            "found": [x[0] for x in positions],
            "issues": issues,
        }

    def check_page_continuity(self, toc_entries):
        """
        按目录页码检查连续性：
        - 目录页码不要求“覆盖每一页”，因此不做“缺页”判断（否则会大量误报）
        - 只检查：页码是否递增（不允许倒退）/ 阿拉伯页码是否从 1 开始 / 罗马页码是否从 I(=1) 开始（若存在）
        """
        roman_pages = []
        arabic_pages = []

        for title, page_raw, page_type, page_value in toc_entries:
            if page_type == "roman":
                roman_pages.append((title, page_raw, page_value))
            elif page_type == "arabic":
                arabic_pages.append((title, page_raw, page_value))

        issues = []

        def _check_sequence(items, label, expect_start=None):
            if not items:
                return
            values = [v for (_t, _pr, v) in items]
            min_v = min(values)
            # 目录里应该递增
            for i in range(1, len(values)):
                if values[i] < values[i - 1]:
                    issues.append(f"{label}页码非递增：`{items[i-1][0]}`({items[i-1][1]}) -> `{items[i][0]}`({items[i][1]})")

            if expect_start is not None and min_v != expect_start:
                issues.append(f"{label}页码起始异常：应从 {expect_start} 开始，实际最小为 {min_v}")

        _check_sequence(roman_pages, "罗马")
        _check_sequence(arabic_pages, "阿拉伯", expect_start=1 if arabic_pages else None)

        return {
            "roman_count": len(roman_pages),
            "arabic_count": len(arabic_pages),
            "issues": issues,
        }

    def _chinese_num_to_int(self, s: str) -> int | None:
        """将中文数字（一、二、...、十、十一、...）转为整数"""
        mapping = {'一':1,'二':2,'三':3,'四':4,'五':5,'六':6,'七':7,'八':8,'九':9,'十':10}
        s = s.strip()
        if s in mapping:
            return mapping[s]
        if s.startswith('十'):
            suffix = s[1:]
            if not suffix:
                return 10
            return 10 + mapping.get(suffix, 0)
        if s.endswith('十') and len(s) == 2:
            return mapping.get(s[0], 0) * 10
        return None

    def _find_chapter_headings(self, doc: Document) -> list:
        """查找所有一级章标题段落（排除目录区域），返回 [(段落索引, 章号, 标题文本), ...]"""
        chapters = []
        # 先定位目录标题，找到目录区域的结束位置
        body_start = 0
        toc_heading_idx = -1
        for i, para in enumerate(doc.paragraphs):
            text_norm = re.sub(r'\s+', '', para.text.strip())
            if text_norm in ('目录', '目录目录'):
                toc_heading_idx = i
                break

        if toc_heading_idx >= 0:
            # 从目录标题之后找到最后一个目录条目
            last_toc = toc_heading_idx
            for i in range(toc_heading_idx + 1, len(doc.paragraphs)):
                text = doc.paragraphs[i].text.strip()
                style_name = doc.paragraphs[i].style.name.lower() if doc.paragraphs[i].style else ''
                if not text:
                    continue
                # 判断是否为目录条目：toc样式、或包含tab+页码、或包含引导点+页码
                is_toc = False
                if 'toc' in style_name:
                    is_toc = True
                elif '\t' in text and re.search(r'\t\s*[IVXLCDM\d]+$', text):
                    is_toc = True
                elif re.search(r'[\.\·…]{3,}[IVXLCDM\d]+$', text):
                    is_toc = True
                if is_toc:
                    last_toc = i
            body_start = last_toc + 1

        # 从正文区域扫描章标题
        for i, para in enumerate(doc.paragraphs):
            if i < body_start:
                continue
            text = para.text.strip()
            # 安全过滤：跳过目录条目格式（防止TOC区域未被正确排除）
            if '\t' in text and re.search(r'\t\s*[IVXLCDM\d]+$', text):
                continue
            if re.search(r'[\.\·…]{3,}[IVXLCDM\d]+$', text):
                continue
            m = re.match(r'^第([一二三四五六七八九十\d]+)章\s', text)
            if m:
                num_str = m.group(1)
                if num_str.isdigit():
                    num = int(num_str)
                else:
                    num = self._chinese_num_to_int(num_str)
                    if num is None:
                        continue
                chapters.append((i, num, text))
        return chapters

    def _find_body_start(self, doc: Document) -> int:
        """定位正文起始位置（跳过目录区域），返回段落索引"""
        # 复用 _find_chapter_headings 中的目录边界定位逻辑
        toc_heading_idx = -1
        for i, para in enumerate(doc.paragraphs):
            text_norm = re.sub(r'\s+', '', para.text.strip())
            if text_norm in ('目录', '目录目录'):
                toc_heading_idx = i
                break
        if toc_heading_idx < 0:
            return 0
        last_toc = toc_heading_idx
        for i in range(toc_heading_idx + 1, len(doc.paragraphs)):
            text = doc.paragraphs[i].text.strip()
            style_name = doc.paragraphs[i].style.name.lower() if doc.paragraphs[i].style else ''
            if not text:
                continue
            is_toc = False
            if 'toc' in style_name:
                is_toc = True
            elif '\t' in text and re.search(r'\t\s*[IVXLCDM\d]+$', text):
                is_toc = True
            elif re.search(r'[\.\·…]{3,}[IVXLCDM\d]+$', text):
                is_toc = True
            if is_toc:
                last_toc = i
        return last_toc + 1

    def check_chapter_summary(self, doc: Document) -> dict:
        """检查每章是否包含"本章小结"段落"""
        chapters = self._find_chapter_headings(doc)
        if not chapters:
            return {"total_chapters": 0, "chapters_missing_summary": [], "issues": ["未找到章节标题"]}

        # 找正文结束位置（参考文献、致谢或附录），从正文区域开始扫描，避免目录条目干扰
        body_start = self._find_body_start(doc)
        body_end = len(doc.paragraphs)
        for i in range(body_start, len(doc.paragraphs)):
            para = doc.paragraphs[i]
            text = re.sub(r'\s+', '', para.text.strip())
            if re.match(r'^(参考文献|致谢|附录)', text):
                # 二次验证：排除正文中提到这些词的段落（如"致谢与附录检测。系统..."）
                # 真正的章节标题很短（≤20字符）或使用了标题样式
                raw_text = para.text.strip()
                style_name = para.style.name.lower() if para.style else ''
                is_heading_style = 'heading' in style_name or 'toc' in style_name
                is_short = len(raw_text) <= 20
                if is_heading_style or is_short:
                    body_end = i
                    break

        missing = []
        for idx in range(len(chapters)):
            p_idx, ch_num, ch_title = chapters[idx]
            # 确定该章段落范围：从本章标题之后到下一章标题之前（或正文结束）
            start = p_idx + 1
            end = chapters[idx + 1][0] if idx + 1 < len(chapters) else body_end

            found = False
            for j in range(start, min(end, len(doc.paragraphs))):
                p_text = re.sub(r'\s+', '', doc.paragraphs[j].text)
                if '本章小结' in p_text:
                    found = True
                    break

            if not found:
                missing.append({
                    "chapter_number": ch_num,
                    "chapter_title": ch_title,
                    "paragraph_index": p_idx,
                })

        issues = []
        for m in missing:
            ch_title_short = re.sub(r'^第[一二三四五六七八九十\d]+章\s*', '', m['chapter_title'])
            issues.append(f"第{m['chapter_number']}章  {ch_title_short}")

        return {
            "total_chapters": len(chapters),
            "chapters_missing_summary": missing,
            "issues": issues,
        }

    def _has_page_break_before(self, doc: Document, para_idx: int) -> dict:
        """检测第 para_idx 段是否从新页开始，返回详细信息"""
        result = {"has_explicit_break": False, "has_auto_indicator": False, "has_break_before": False}
        para = doc.paragraphs[para_idx]

        # A: 当前段落 runs 中的显式分页符
        for run in para.runs:
            for br in run._element.iter(qn('w:br')):
                if br.get(qn('w:type')) == 'page':
                    result["has_explicit_break"] = True
                    result["has_break_before"] = True
                    return result

        # B: 前一段落 runs 中的显式分页符
        if para_idx > 0:
            prev = doc.paragraphs[para_idx - 1]
            for run in prev.runs:
                for br in run._element.iter(qn('w:br')):
                    if br.get(qn('w:type')) == 'page':
                        result["has_explicit_break"] = True
                        result["has_break_before"] = True
                        return result

        # C: 段前分页属性
        pPr = para._element.find(qn('w:pPr'))
        if pPr is not None:
            pbp = pPr.find(qn('w:pageBreakBefore'))
            if pbp is not None:
                result["has_explicit_break"] = True
                result["has_break_before"] = True
                return result

        # D: 前一段落的自动分页标记
        if para_idx > 0:
            prev = doc.paragraphs[para_idx - 1]
            for run in prev.runs:
                for _ in run._element.iter(qn('w:lastRenderedPageBreak')):
                    result["has_auto_indicator"] = True
                    result["has_break_before"] = True
                    return result

        # E: 向前查找最近的分节符（限制在前5段以内）
        for offset in range(1, min(6, para_idx + 1)):
            prev = doc.paragraphs[para_idx - offset]
            pPr_prev = prev._element.find(qn('w:pPr'))
            if pPr_prev is not None:
                sectPr = pPr_prev.find(qn('w:sectPr'))
                if sectPr is not None:
                    sect_type = sectPr.find(qn('w:type'))
                    if sect_type is not None:
                        val = sect_type.get(qn('w:val'))
                        if val in ('nextPage', 'evenPage', 'oddPage'):
                            result["has_explicit_break"] = True
                            result["has_break_before"] = True
                            return result
                        # continuous 等不分页类型：最近的分节符决定了本节不从新页开始
                        return result
                    else:
                        # 无 type 属性，默认为 nextPage
                        result["has_explicit_break"] = True
                        result["has_break_before"] = True
                        return result

        return result

    def check_chapter_page_breaks(self, doc: Document) -> dict:
        """检查每章标题是否从新页开始"""
        chapters = self._find_chapter_headings(doc)
        if not chapters:
            return {"total_chapters": 0, "chapters_with_issues": [], "issues": ["未找到章节标题"]}

        problems = []
        for p_idx, ch_num, ch_title in chapters:
            pb = self._has_page_break_before(doc, p_idx)
            if not pb["has_break_before"]:
                problems.append({
                    "chapter_number": ch_num,
                    "chapter_title": ch_title,
                    "paragraph_index": p_idx,
                    "has_explicit_break": pb["has_explicit_break"],
                    "has_auto_indicator": pb["has_auto_indicator"],
                    "has_break_before": pb["has_break_before"],
                })

        issues = []
        for p in problems:
            ch_title_short = re.sub(r'^第[一二三四五六七八九十\d]+章\s*', '', p['chapter_title'])
            issues.append(f"第{p['chapter_number']}章  {ch_title_short}")

        return {
            "total_chapters": len(chapters),
            "chapters_with_issues": problems,
            "issues": issues,
        }

    def generate_report(self, docx_path, report_path="structure_report.md"):
        doc = Document(docx_path)
        toc_entries = self.extract_toc_entries(doc)

        order_result = self.check_order(toc_entries)
        page_result = self.check_page_continuity(toc_entries)
        summary_result = self.check_chapter_summary(doc)
        page_break_result = self.check_chapter_page_breaks(doc)

        lines = []
        lines.append("# 论文整体结构检测报告\n")

        # ---- 检测概览 ----
        lines.append("## 检测概览\n")
        lines.append("| 检测项 | 结果 |")
        lines.append("|--------|------|")

        def _status_cell(issues_list):
            if not issues_list:
                return "[通过]"
            return f"[警告] {len(issues_list)}项"

        order_ok = not order_result["issues"] if toc_entries else None
        page_ok = not page_result["issues"] if toc_entries else None
        summary_ok = not summary_result["issues"] if summary_result["total_chapters"] > 0 else None
        pb_ok = not page_break_result["issues"] if page_break_result["total_chapters"] > 0 else None

        if order_ok is None:
            lines.append(f"| 章节顺序 | [跳过] 无目录数据 |")
        elif order_ok:
            lines.append(f"| 章节顺序 | [通过] |")
        else:
            lines.append(f"| 章节顺序 | [警告] {len(order_result['issues'])}项 |")

        if page_ok is None:
            lines.append(f"| 页码连续性 | [跳过] 无目录数据 |")
        elif page_ok:
            lines.append(f"| 页码连续性 | [通过] |")
        else:
            lines.append(f"| 页码连续性 | [警告] {len(page_result['issues'])}项 |")

        if summary_ok is None:
            lines.append(f"| 本章小结 | [跳过] 无章节数据 |")
        elif summary_ok:
            lines.append(f"| 本章小结 | [通过] |")
        else:
            lines.append(f"| 本章小结 | [警告] **{len(summary_result['issues'])}个章节**缺少 |")

        if pb_ok is None:
            lines.append(f"| 章节分页 | [跳过] 无章节数据 |")
        elif pb_ok:
            lines.append(f"| 章节分页 | [通过] |")
        else:
            lines.append(f"| 章节分页 | [警告] **{len(page_break_result['issues'])}个章节**未另起一页 |")

        lines.append("")
        lines.append(f"> 共检测 **{summary_result['total_chapters']}** 个章节，目录可解析条目 **{len(toc_entries)}** 项")
        lines.append("")

        # ---- 章节顺序检测 ----
        lines.append("## 章节顺序检测")
        if not toc_entries:
            lines.append("[警告] 未能解析目录项，请确认文档有自动目录且包含页码")
        elif not order_result["issues"]:
            lines.append("[通过] 关键章节顺序正常")
        else:
            lines.append(f"[警告] 发现 {len(order_result['issues'])} 个顺序问题")
            for it in order_result["issues"]:
                lines.append(f"- {it}")
        lines.append("")

        # ---- 页码连续性 ----
        lines.append("## 页码连续性检测")
        if not toc_entries:
            lines.append("[警告] 未能解析目录页码")
        elif not page_result["issues"]:
            lines.append("[通过] 目录页码递增正常")
        else:
            lines.append(f"[警告] 发现 {len(page_result['issues'])} 个问题")
            for it in page_result["issues"]:
                lines.append(f"- {it}")
        lines.append("")

        # ---- 本章小结 ----
        lines.append("## 本章小结检查")
        if summary_result["total_chapters"] == 0:
            lines.append("[警告] 未检测到章节标题")
        elif not summary_result["issues"]:
            lines.append(f"[通过] 全部 {summary_result['total_chapters']} 个章节均包含\"本章小结\"")
        else:
            lines.append(f"[警告] **发现 {len(summary_result['issues'])} 个章节缺少\"本章小结\"**：\n")
            for it in summary_result["issues"]:
                lines.append(f"- {it}")
        lines.append("")

        # ---- 章节分页 ----
        lines.append("## 章节分页检测")
        if page_break_result["total_chapters"] == 0:
            lines.append("[警告] 未检测到章节标题")
        elif not page_break_result["issues"]:
            lines.append(f"[通过] 全部 {page_break_result['total_chapters']} 个章节均从新页开始")
        else:
            lines.append(f"[警告] **发现 {len(page_break_result['issues'])} 个章节未另起一页**：\n")
            for it in page_break_result["issues"]:
                lines.append(f"- {it}")
        lines.append("")

        # ---- 目录项 ----
        lines.append("## 目录项（完整）")
        if not toc_entries:
            lines.append("（未解析到目录项）")
        else:
            lines.append('<details><summary>展开查看全部 {0} 条目录项</summary>\n'.format(len(toc_entries)))
            for (t, p_raw, p_type, p_val) in toc_entries:
                lines.append(f"- {t}\t{p_raw}  ({p_type}:{p_val})")
            lines.append('</details>')

        report = "\n".join(lines)

        with open(report_path, "w", encoding="utf-8") as f:
            f.write(report)

        print(report)
        print(f"\n报告已保存至: {os.path.abspath(report_path)}")
        return report


def main():
    """主函数"""
    import os
    import sys
    current_dir = os.path.dirname(os.path.abspath(__file__))
    base_dir = os.path.dirname(current_dir)
    docx_path = os.path.join(base_dir, "测试论文.docx")
    
    if not os.path.exists(docx_path):
        print(f"[错误] 文件不存在：{docx_path}")
        print(f"   当前目录：{current_dir}")
        sys.exit(1)
    
    try:
        checker = ThesisStructureChecker()
        reports_dir = os.path.join(base_dir, "reports")
        os.makedirs(reports_dir, exist_ok=True)
        checker.generate_report(docx_path, os.path.join(reports_dir, "structure_report.md"))
    except Exception as e:
        print(f"[错误] 检测过程中出错：{str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()


