from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re


class AckAppendixChecker:
    """
    致谢 & 附录 格式检测器（第一部分）

    当前这一段代码完成的功能：
    1. 在文档中定位「致谢」和「附录」标题位置
    2. 读取这些标题及其后正文段落的字体、字号、对齐方式
    3. 按预设规则进行简单合规性判断并打印报告

    后续如果需要，可以再追加：
    - 更细的行距、段前段后、首行缩进检测
    - 更复杂的多附录（附录A/B/C...）结构判断
    """

    def __init__(self):
        # 标准正文格式（参考 BodyTextChecker）
        self.body_format = {
            "chinese_font": ["宋体", "SimSun"],
            "english_font": ["Times New Roman"],
            "font_size": 12.0,  # 小四号 = 12pt
            "line_spacing": 1.5,  # 1.5倍行距
            "first_line_indent": 24.0,  # 首行缩进2字符 = 12*2 = 24pt
            "alignment": "JUSTIFY",  # 两端对齐
            "space_before": 0,  # 段前0行
            "space_after": 0,  # 段后0行
        }

        # 致谢标题格式（按你确认：三号=16pt）
        self.ack_title_format = {
            "fonts": ["黑体", "SimHei"],
            "size": 16.0,  # 三号
            "bold": True,
            "alignment": WD_ALIGN_PARAGRAPH.CENTER,
        }

        # 附录标题格式（按你确认：三号=16pt）
        self.appendix_title_format = {
            "fonts": ["黑体", "SimHei"],
            "size": 16.0,
            "bold": True,
            "alignment": WD_ALIGN_PARAGRAPH.CENTER,
        }

    # ========= 通用工具函数 =========
    def get_font_name(self, run, paragraph=None):
        """获取运行块的字体名称（优先中文）"""
        has_chinese = any("\u4e00" <= ch <= "\u9fff" for ch in run.text)

        try:
            if hasattr(run._element, "rPr") and run._element.rPr is not None:
                rpr = run._element.rPr
                rfonts_elem = rpr.find(qn("w:rFonts"))
                if rfonts_elem is not None:
                    if has_chinese:
                        eastasia = rfonts_elem.get(qn("w:eastAsia"))
                        if eastasia:
                            return eastasia
                    ascii_font = rfonts_elem.get(qn("w:ascii"))
                    if ascii_font:
                        return ascii_font
        except Exception:
            pass

        if run.font.name:
            return run.font.name

        if paragraph and paragraph.style and paragraph.style.font and paragraph.style.font.name:
            return paragraph.style.font.name

        return None

    def get_font_size(self, run, paragraph=None):
        """获取字号（pt）"""
        if run.font.size:
            return run.font.size.pt
        if run.style and run.style.font and run.style.font.size:
            return run.style.font.size.pt
        if paragraph and paragraph.style and paragraph.style.font and paragraph.style.font.size:
            return paragraph.style.font.size.pt
        return None

    def is_run_effectively_bold(self, run, paragraph=None):
        """
        获取 run 的“有效加粗”状态：
        - True/False 表示能确定
        - None 表示无法确定（可能继承样式且 python-docx 未暴露）
        """
        # 1) 直接属性
        if run.bold is True:
            return True
        if run.bold is False:
            return False

        # 2) font 层（有时会在这里）
        try:
            if run.font and run.font.bold is True:
                return True
            if run.font and run.font.bold is False:
                return False
        except Exception:
            pass

        # 3) run 的样式
        try:
            if run.style and run.style.font:
                if run.style.font.bold is True:
                    return True
                if run.style.font.bold is False:
                    return False
        except Exception:
            pass

        # 4) 段落样式
        try:
            if paragraph and paragraph.style and paragraph.style.font:
                if paragraph.style.font.bold is True:
                    return True
                if paragraph.style.font.bold is False:
                    return False
        except Exception:
            pass

        # 5) XML（w:b / w:bCs）
        try:
            rPr = run._element.rPr
            if rPr is not None:
                b = rPr.find(qn("w:b"))
                if b is not None:
                    val = b.get(qn("w:val"))
                    # w:val 可能缺失（等同 true），或为 "0"/"false"
                    if val is None:
                        return True
                    if str(val).lower() in ("0", "false", "off"):
                        return False
                    return True

                bcs = rPr.find(qn("w:bCs"))
                if bcs is not None:
                    val = bcs.get(qn("w:val"))
                    if val is None:
                        return True
                    if str(val).lower() in ("0", "false", "off"):
                        return False
                    return True
        except Exception:
            pass

        return None

    def is_paragraph_effectively_bold(self, paragraph):
        """
        判断段落是否“有效加粗”：
        - 如果存在任何明确 False 的 run，则判定 False
        - 如果存在任何明确 True 的 run，则判定 True
        - 否则回退段落样式（bold=True 判定 True），最后返回 None
        """
        any_true = False
        for run in paragraph.runs:
            if not run.text.strip():
                continue
            eff = self.is_run_effectively_bold(run, paragraph)
            if eff is False:
                return False
            if eff is True:
                any_true = True

        if any_true:
            return True

        try:
            if paragraph.style and paragraph.style.font:
                if paragraph.style.font.bold is True:
                    return True
                if paragraph.style.font.bold is False:
                    return False
        except Exception:
            pass

        return None

    def get_paragraph_main_font_size(self, paragraph):
        """统计段落内出现次数最多的字号，作为该段字号"""
        size_count = {}
        for run in paragraph.runs:
            if not run.text.strip():
                continue
            sz = self.get_font_size(run, paragraph)
            if sz:
                size_count[sz] = size_count.get(sz, 0) + 1
        if not size_count:
            return None
        # 出现次数最多的字号
        return max(size_count.items(), key=lambda kv: kv[1])[0]
    
    def get_line_spacing(self, paragraph):
        """获取行距（倍数）"""
        try:
            pPr = paragraph._element.find(qn('w:pPr'))
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
    
    def get_first_line_indent(self, paragraph):
        """获取首行缩进（单位：磅）"""
        # 方法1：检查标准首行缩进属性
        try:
            indent = paragraph.paragraph_format.first_line_indent
            if indent and indent.pt != 0:
                return indent.pt
        except:
            pass
        
        # 方法2：从段落样式获取首行缩进
        try:
            if paragraph.style:
                style_element = paragraph.style.element
                pPr = style_element.find(qn('w:pPr'))
                if pPr is not None:
                    ind = pPr.find(qn('w:ind'))
                    if ind is not None:
                        firstLine = ind.get(qn('w:firstLine'))
                        if firstLine:
                            # firstLine的单位是twips（1/20 pt）
                            return int(firstLine) / 20.0
        except:
            pass
        
        # 方法3：检查段落开头的空格或制表符（手动缩进）
        text = paragraph.text
        if text:
            # 检查开头的空格数
            leading_spaces = 0
            for char in text:
                if char == ' ':  # 半角空格
                    leading_spaces += 0.5
                elif char == '　':  # 全角空格
                    leading_spaces += 1
                elif char == '\t':  # 制表符
                    return 24.0  # 制表符通常表示2字符缩进
                else:
                    break
            
            # 如果有至少1个字符的空格缩进，按比例计算
            if leading_spaces >= 1:
                # 按每个字符12pt计算（因为字号是12pt）
                return leading_spaces * 12.0
        
        return 0
    
    def get_alignment(self, paragraph):
        """
        获取“有效”的对齐方式。
        注意：python-docx 的 paragraph.alignment 可能为 None（表示继承样式），但在 Word 中实际显示可能是居中/两端对齐等。
        """
        alignment_map = {
            0: "LEFT",
            1: "CENTER",
            2: "RIGHT",
            3: "JUSTIFY",
            4: "DISTRIBUTE",
        }

        # 1) 段落直接设置（最可靠）
        if paragraph.alignment is not None:
            return alignment_map.get(paragraph.alignment, "UNKNOWN")

        # 2) 段落格式（有时会在这里）
        try:
            pf_align = paragraph.paragraph_format.alignment
            if pf_align is not None:
                return alignment_map.get(pf_align, "UNKNOWN")
        except Exception:
            pass

        # 3) 段落样式（常见：标题居中是通过样式实现的）
        try:
            if paragraph.style and paragraph.style.paragraph_format:
                st_align = paragraph.style.paragraph_format.alignment
                if st_align is not None:
                    return alignment_map.get(st_align, "UNKNOWN")
        except Exception:
            pass

        # 4) 直接读段落 XML（w:jc），以及样式 XML（w:pPr/w:jc）
        def _map_jc_val(val: str | None):
            if not val:
                return None
            val = val.lower()
            if val in ("center",):
                return "CENTER"
            if val in ("left", "start"):
                return "LEFT"
            if val in ("right", "end"):
                return "RIGHT"
            if val in ("both", "justify"):
                return "JUSTIFY"
            if val in ("distribute",):
                return "DISTRIBUTE"
            return None

        try:
            pPr = paragraph._element.pPr
            if pPr is not None:
                jc = pPr.find(qn("w:jc"))
                if jc is not None:
                    mapped = _map_jc_val(jc.get(qn("w:val")))
                    if mapped:
                        return mapped
        except Exception:
            pass

        try:
            if paragraph.style:
                style_element = paragraph.style.element
                pPr = style_element.find(qn("w:pPr"))
                if pPr is not None:
                    jc = pPr.find(qn("w:jc"))
                    if jc is not None:
                        mapped = _map_jc_val(jc.get(qn("w:val")))
                        if mapped:
                            return mapped
        except Exception:
            pass

        # 5) 最后兜底：None=继承/默认，无法确定
        return "DEFAULT"

    # ========= 标题识别 =========
    def is_ack_title(self, paragraph):
        text = paragraph.text.strip()
        return bool(re.match(r"^致\s*谢$", text))

    def is_appendix_title(self, paragraph):
        text = paragraph.text.strip()
        # 「附录」「附录A」「附录A 标题」等
        return bool(re.match(r"^附\s*录(\s*[A-ZＡ-Ｚ]?)", text))

    # ========= 核心检测逻辑 =========
    def check_ack_and_appendix(self, docx_path):
        """
        检测文档中的致谢和附录标题及其正文的格式。
        返回检测结果字典，用于生成报告。
        """
        doc = Document(docx_path)

        print("========== 致谢 & 附录 格式检测 ==========")
        print(f"文档：{docx_path}\n")

        ack_index = None
        appendix_indices = []

        paragraphs = doc.paragraphs

        # 识别目录部分并收集附录标题
        in_toc = False

        for i, para in enumerate(paragraphs):
            text = para.text.strip()
            if not text:
                continue

            # 检查是否是目录标题
            if re.match(r"^[目目]\s*录$", text) or re.match(r"^Contents$", text, re.IGNORECASE):
                in_toc = True
                continue

            # 如果在目录中，跳过所有内容（包括附录条目）
            if in_toc:
                # 检查是否是目录结束的标志（明显不是目录条目的内容）
                if not re.search(r"\d+$|^\d+\.", text):
                    in_toc = False
                continue

            # 不在目录中，正常检测
            if ack_index is None and self.is_ack_title(para):
                ack_index = i

            if self.is_appendix_title(para):
                # 检查是否是目录中的附录条目（通常包含页码或数字）
                appendix_text = para.text.strip()
                if re.search(r"\d+", appendix_text):
                    # 跳过目录中的附录条目
                    continue
                appendix_indices.append(i)

        results = {
            'docx_path': docx_path,
            'acknowledgment': None,
            'appendices': []
        }

        # ---- 致谢部分检测 ----
        if ack_index is None:
            print("[错误] 未找到「致谢」标题（精确匹配\"致谢\"）")
        else:
            print(f"[通过] 找到「致谢」标题，段落索引：{ack_index}")
            ack_result = self._check_ack_section(paragraphs, ack_index, return_result=True)
            results['acknowledgment'] = ack_result
            self._check_ack_section(paragraphs, ack_index, return_result=False)

        print("\n----------------------------------------\n")

        # ---- 附录部分检测 ----
        if not appendix_indices:
            print("[警告] 未检测到任何「附录」标题（匹配\"附录\"\"附录A/B...\"）")
        else:
            print(f"[通过] 共检测到 {len(appendix_indices)} 个附录标题：{appendix_indices}")
            for idx in appendix_indices:
                appendix_result = self._check_appendix_section(paragraphs, idx, return_result=True)
                results['appendices'].append(appendix_result)
                self._check_appendix_section(paragraphs, idx, return_result=False)
        
        return results

    # ========= 致谢检测 =========
    def _check_ack_section(self, paragraphs, title_index, return_result=False):
        title_para = paragraphs[title_index]
        fmt = self.ack_title_format

        # 标题字体/字号/对齐检测
        title_main_size = self.get_paragraph_main_font_size(title_para)
        title_fonts = {self.get_font_name(run, title_para) for run in title_para.runs if run.text.strip()}
        title_fonts.discard(None)

        title_problems = []

        # 字体
        if not title_fonts:
            title_problems.append("无法检测到标题字体")
        else:
            if not any(f in fmt["fonts"] for f in title_fonts):
                title_problems.append(f"标题字体应为：{fmt['fonts']}，当前：{list(title_fonts)}")

        # 字号
        if title_main_size is None:
            title_problems.append("无法检测到标题字号")
        elif abs(title_main_size - fmt["size"]) > 0.5:
            title_problems.append(f"标题字号应为：{fmt['size']}pt，当前：{title_main_size}pt")

        # 对齐（用“有效对齐”，避免样式继承导致 paragraph.alignment=None 误判）
        if self.get_alignment(title_para) != "CENTER":
            title_problems.append("标题应为居中对齐")

        # 加粗（只要有一部分不加粗就提示）
        eff_bold = self.is_paragraph_effectively_bold(title_para)
        if eff_bold is False:
            title_problems.append("标题建议整体加粗显示")
        elif eff_bold is None:
            # 无法确定时不误报（通常是继承样式导致 bold 为 None）
            pass

        if not return_result:
            print("\n【致谢 标题检测】")
            print(f"  标题文本：{title_para.text.strip()}")
            print(f"  检测到字号：{title_main_size} pt，字体集合：{title_fonts or {'(未检测到)'}}")
            if title_problems:
                print("  [错误] 标题格式存在问题：")
                for p in title_problems:
                    print("   -", p)
            else:
                print("  [通过] 标题格式符合预期")

        # 致谢正文详细检测：连续正文段，直到遇到下一个大节（参考文献/附录等）
        print("\n【致谢 正文检测】")
        body_start = title_index + 1
        body_end = len(paragraphs)

        for i in range(body_start, len(paragraphs)):
            text = paragraphs[i].text.strip()
            if not text:
                continue
            # 下一个模块开始：参考文献 / 附录 / Abstract 等
            if re.match(r"^参\s*考\s*文\s*献", text) or re.match(r"^附\s*录", text) or re.match(
                r"^Abstract", text, re.IGNORECASE
            ):
                body_end = i
                break

        if body_start >= body_end:
            print("  [警告] 未检测到致谢正文内容（标题后没有有效段落）")
            return

        problem_count = 0
        total_paragraphs = 0
        body_issues = []
        
        for i in range(body_start, body_end):
            para = paragraphs[i]
            text = para.text.strip()
            if not text:
                continue
            
            total_paragraphs += 1
            local_problems = []
            
            # 1. 字号检测
            main_size = self.get_paragraph_main_font_size(para)
            if main_size is not None:
                if abs(main_size - self.body_format["font_size"]) > 0.5:
                    local_problems.append(f"字号应为 {self.body_format['font_size']}pt，当前 {main_size}pt")
            
            # 2. 字体检测
            fonts = {self.get_font_name(run, para) for run in para.runs if run.text.strip()}
            fonts.discard(None)
            if fonts:
                valid_fonts = self.body_format["chinese_font"] + self.body_format["english_font"]
                if not any(f in valid_fonts for f in fonts):
                    local_problems.append(f"字体应为宋体/Times New Roman，当前 {list(fonts)}")
            
            # 3. 行距检测
            line_spacing = self.get_line_spacing(para)
            if line_spacing is not None:
                expected_spacing = self.body_format["line_spacing"]
                if abs(line_spacing - expected_spacing) > 0.1:
                    local_problems.append(f"行距应为 {expected_spacing}倍，当前 {line_spacing}倍")
            
            # 4. 首行缩进检测
            first_indent = self.get_first_line_indent(para)
            expected_indent = self.body_format["first_line_indent"]
            if abs(first_indent - expected_indent) > 2.0:  # 允许2pt误差
                local_problems.append(f"首行缩进应为 {expected_indent}pt（2字符），当前 {first_indent:.1f}pt")
            
            # 5. 对齐方式检测
            alignment = self.get_alignment(para)
            expected_alignment = self.body_format["alignment"]
            if alignment != expected_alignment:
                local_problems.append(f"对齐方式应为 {expected_alignment}，当前 {alignment}")

            if local_problems:
                problem_count += 1
                preview = text[:50] + ("..." if len(text) > 50 else "")
                body_issues.append({
                    'paragraph_index': i,
                    'text_preview': preview,
                    'problems': local_problems,
                })
                print(f"\n  段落 {i}：{preview}")
                for p in local_problems:
                    print(f"    [错误] {p}")

        if not return_result:
            print(f"\n  检测统计：共 {total_paragraphs} 个段落")
            if problem_count == 0:
                print("  [通过] 致谢正文格式全部符合要求")
            else:
                print(f"  [警告] 发现 {problem_count} 个段落存在格式问题")
        
        if return_result:
            return {
                'title_index': title_index,
                'title_text': title_para.text.strip(),
                'title_problems': title_problems,
                'title_font_size': title_main_size,
                'title_fonts': list(title_fonts) if title_fonts else [],
                'body_paragraphs': total_paragraphs,
                'body_problems': problem_count,
                'body_issues': body_issues,
            }

    # ========= 附录检测（详细） =========
    def _check_appendix_section(self, paragraphs, title_index, return_result=False):
        title_para = paragraphs[title_index]
        fmt = self.appendix_title_format

        title_main_size = self.get_paragraph_main_font_size(title_para)
        title_fonts = {self.get_font_name(run, title_para) for run in title_para.runs if run.text.strip()}
        title_fonts.discard(None)

        title_problems = []

        if not title_fonts:
            title_problems.append("无法检测到标题字体")
        else:
            if not any(f in fmt["fonts"] for f in title_fonts):
                title_problems.append(f"标题字体应为：{fmt['fonts']}，当前：{list(title_fonts)}")

        if title_main_size is None:
            title_problems.append("无法检测到标题字号")
        elif abs(title_main_size - fmt["size"]) > 0.5:
            title_problems.append(f"标题字号应为：{fmt['size']}pt，当前：{title_main_size}pt")

        if self.get_alignment(title_para) != "CENTER":
            title_problems.append("标题应为居中对齐")

        eff_bold = self.is_paragraph_effectively_bold(title_para)
        if eff_bold is False:
            title_problems.append("标题建议整体加粗显示")
        elif eff_bold is None:
            pass

        if not return_result:
            print("\n【附录 标题检测】")
            print(f"  标题文本：{title_para.text.strip()}")
            print(f"  段落索引：{title_index}")
            print(f"  检测到字号：{title_main_size} pt，字体集合：{title_fonts or {'(未检测到)'}}")
            if title_problems:
                print("  [错误] 标题格式存在问题：")
                for p in title_problems:
                    print("   -", p)
            else:
                print("  [通过] 标题格式符合预期")
        
        # 附录正文详细检测
        if not return_result:
            print("\n【附录 正文检测】")
        body_start = title_index + 1
        body_end = len(paragraphs)

        # 找到下一个附录标题或参考文献等
        for i in range(body_start, len(paragraphs)):
            text = paragraphs[i].text.strip()
            if not text:
                continue
            # 下一个模块开始：参考文献 / 另一个附录 / Abstract 等
            if (re.match(r"^参\s*考\s*文\s*献", text) or 
                re.match(r"^附\s*录", text) or 
                re.match(r"^Abstract", text, re.IGNORECASE) or
                re.match(r"^致\s*谢", text)):
                body_end = i
                break

        if body_start >= body_end:
            if not return_result:
                print("  [警告] 未检测到附录正文内容（标题后没有有效段落）")
            if return_result:
                return {
                    'title_index': title_index,
                    'title_text': title_para.text.strip(),
                    'title_problems': title_problems,
                    'title_font_size': title_main_size,
                    'title_fonts': list(title_fonts) if title_fonts else [],
                    'body_paragraphs': 0,
                    'body_problems': 0,
                    'body_issues': []
                }
            return

        problem_count = 0
        total_paragraphs = 0
        body_issues = []
        
        for i in range(body_start, body_end):
            para = paragraphs[i]
            text = para.text.strip()
            if not text:
                continue
            
            total_paragraphs += 1
            local_problems = []
            
            # 1. 字号检测
            main_size = self.get_paragraph_main_font_size(para)
            if main_size is not None:
                if abs(main_size - self.body_format["font_size"]) > 0.5:
                    local_problems.append(f"字号应为 {self.body_format['font_size']}pt，当前 {main_size}pt")
            
            # 2. 字体检测
            fonts = {self.get_font_name(run, para) for run in para.runs if run.text.strip()}
            fonts.discard(None)
            if fonts:
                valid_fonts = self.body_format["chinese_font"] + self.body_format["english_font"]
                if not any(f in valid_fonts for f in fonts):
                    local_problems.append(f"字体应为宋体/Times New Roman，当前 {list(fonts)}")
            
            # 3. 行距检测
            line_spacing = self.get_line_spacing(para)
            if line_spacing is not None:
                expected_spacing = self.body_format["line_spacing"]
                if abs(line_spacing - expected_spacing) > 0.1:
                    local_problems.append(f"行距应为 {expected_spacing}倍，当前 {line_spacing}倍")
            
            # 4. 首行缩进检测
            first_indent = self.get_first_line_indent(para)
            expected_indent = self.body_format["first_line_indent"]
            if abs(first_indent - expected_indent) > 2.0:  # 允许2pt误差
                local_problems.append(f"首行缩进应为 {expected_indent}pt（2字符），当前 {first_indent:.1f}pt")
            
            # 5. 对齐方式检测
            alignment = self.get_alignment(para)
            expected_alignment = self.body_format["alignment"]
            if alignment != expected_alignment:
                local_problems.append(f"对齐方式应为 {expected_alignment}，当前 {alignment}")

            if local_problems:
                problem_count += 1
                preview = text[:50] + ("..." if len(text) > 50 else "")
                body_issues.append({
                    'paragraph_index': i,
                    'text_preview': preview,
                    'problems': local_problems,
                })
                print(f"\n  段落 {i}：{preview}")
                for p in local_problems:
                    print(f"    [错误] {p}")

        if not return_result:
            print(f"\n  检测统计：共 {total_paragraphs} 个段落")
            if problem_count == 0:
                print("  [通过] 附录正文格式全部符合要求")
            else:
                print(f"  [警告] 发现 {problem_count} 个段落存在格式问题")
        
        if return_result:
            return {
                'title_index': title_index,
                'title_text': title_para.text.strip(),
                'title_problems': title_problems,
                'title_font_size': title_main_size,
                'title_fonts': list(title_fonts) if title_fonts else [],
                'body_paragraphs': total_paragraphs,
                'body_problems': problem_count,
                'body_issues': body_issues,
            }
    
    def highlight_paragraph(self, paragraph):
        """在段落中添加高亮背景"""
        for run in paragraph.runs:
            run.font.highlight_color = 7  # 7 表示黄色

    def generate_report(self, docx_path, report_path=None, highlight_path=None):
        """生成致谢和附录格式检测报告"""
        results = self.check_ack_and_appendix(docx_path)

        # 如果提供了高亮路径，对有问题的段落进行高亮
        if highlight_path:
            try:
                hl_doc = Document(highlight_path)

                # 高亮致谢标题和正文问题
                ack = results.get('acknowledgment')
                if ack:
                    if ack.get('title_problems') and ack.get('title_index') is not None:
                        ti = ack['title_index']
                        if 0 <= ti < len(hl_doc.paragraphs):
                            self.highlight_paragraph(hl_doc.paragraphs[ti])
                    for iss in ack.get('body_issues', []):
                        pi = iss['paragraph_index']
                        if 0 <= pi < len(hl_doc.paragraphs):
                            self.highlight_paragraph(hl_doc.paragraphs[pi])

                # 高亮附录标题和正文问题
                for appendix in results.get('appendices', []):
                    if appendix.get('title_problems') and appendix.get('title_index') is not None:
                        ti = appendix['title_index']
                        if 0 <= ti < len(hl_doc.paragraphs):
                            self.highlight_paragraph(hl_doc.paragraphs[ti])
                    for iss in appendix.get('body_issues', []):
                        pi = iss['paragraph_index']
                        if 0 <= pi < len(hl_doc.paragraphs):
                            self.highlight_paragraph(hl_doc.paragraphs[pi])

                hl_doc.save(highlight_path)
                print(f"致谢/附录问题已高亮至: {highlight_path}")
            except Exception as e:
                print(f"高亮致谢/附录问题失败: {e}")

        report_lines = []
        report_lines.append("# 致谢和附录格式检测报告\n")
        
        # 基本信息
        report_lines.append("## 基本信息")
        report_lines.append(f"- 文档路径: {results['docx_path']}")
        report_lines.append("")
        
        # 格式标准
        report_lines.append("## 格式标准\n")
        report_lines.append("### 致谢/附录标题格式")
        report_lines.append("- **字体**: 黑体")
        report_lines.append("- **字号**: 三号（16pt）")
        report_lines.append("- **对齐**: 居中")
        report_lines.append("- **加粗**: 是")
        report_lines.append("")
        report_lines.append("### 致谢/附录正文格式")
        report_lines.append("- **中文字体**: 宋体")
        report_lines.append("- **英文字体**: Times New Roman")
        report_lines.append("- **字号**: 小四号（12pt）")
        report_lines.append("- **行距**: 1.5倍")
        report_lines.append("- **首行缩进**: 2字符（24pt）")
        report_lines.append("- **对齐方式**: 两端对齐")
        report_lines.append("")
        
        # 致谢检测结果
        report_lines.append("## 致谢检测结果")
        if results['acknowledgment'] is None:
            report_lines.append("[错误] **未找到致谢部分**")
        else:
            ack = results['acknowledgment']
            if len(ack['title_problems']) == 0 and ack['body_problems'] == 0:
                report_lines.append("[通过] **致谢格式全部符合要求**")
            else:
                report_lines.append("[警告] **致谢格式存在问题**\n")
                
                if len(ack['title_problems']) > 0:
                    report_lines.append("### 标题问题")
                    for problem in ack['title_problems']:
                        report_lines.append(f"- {problem}")
                    report_lines.append("")
                
                if ack['body_problems'] > 0:
                    report_lines.append(f"### 正文问题")
                    report_lines.append(f"- 共发现 {ack['body_problems']} 个段落存在格式问题")
                    report_lines.append(f"- 总段落数: {ack['body_paragraphs']}")
                    report_lines.append("")
                    if ack.get('body_issues'):
                        report_lines.append("#### 问题明细")
                        for issue in ack['body_issues']:
                            report_lines.append(f"\n**段落 {issue['paragraph_index']}**: {issue['text_preview']}")
                            for p in issue['problems']:
                                report_lines.append(f"  - [错误] {p}")
                        report_lines.append("")
        report_lines.append("")
        
        # 附录检测结果
        report_lines.append("## 附录检测结果")
        if len(results['appendices']) == 0:
            report_lines.append("[警告] **未检测到附录部分**")
        else:
            report_lines.append(f"共检测到 {len(results['appendices'])} 个附录\n")
            
            all_appendix_ok = True
            for idx, appendix in enumerate(results['appendices'], 1):
                report_lines.append(f"### 附录 {idx}: {appendix['title_text']}")
                
                if len(appendix['title_problems']) == 0 and appendix['body_problems'] == 0:
                    report_lines.append("[通过] **格式全部符合要求**")
                else:
                    all_appendix_ok = False
                    if len(appendix['title_problems']) > 0:
                        report_lines.append("#### 标题问题")
                        for problem in appendix['title_problems']:
                            report_lines.append(f"- {problem}")
                        report_lines.append("")
                    
                    if appendix['body_problems'] > 0:
                        report_lines.append(f"#### 正文问题")
                        report_lines.append(f"- 共发现 {appendix['body_problems']} 个段落存在格式问题")
                        report_lines.append(f"- 总段落数: {appendix['body_paragraphs']}")
                        if appendix.get('body_issues'):
                            report_lines.append("")
                            report_lines.append("#### 问题明细")
                            for issue in appendix['body_issues']:
                                report_lines.append(f"\n**段落 {issue['paragraph_index']}**: {issue['text_preview']}")
                                for p in issue['problems']:
                                    report_lines.append(f"  - [错误] {p}")
                            report_lines.append("")
                report_lines.append("")
            
            if all_appendix_ok:
                report_lines.insert(-len(results['appendices']), "[通过] **所有附录格式全部符合要求**\n")
        report_lines.append("")
        
        # 总结
        report_lines.append("## 总结")
        total_title_problems = 0
        total_body_problems = 0
        
        if results['acknowledgment']:
            total_title_problems += len(results['acknowledgment']['title_problems'])
            total_body_problems += results['acknowledgment']['body_problems']
        
        for appendix in results['appendices']:
            total_title_problems += len(appendix['title_problems'])
            total_body_problems += appendix['body_problems']
        
        total_problems = total_title_problems + total_body_problems
        
        if total_problems == 0:
            report_lines.append("[通过] **格式检测通过，所有内容格式规范**")
        else:
            report_lines.append(f"[警告] **共发现 {total_problems} 个格式问题**")
            report_lines.append(f"- 标题问题: {total_title_problems}个")
            report_lines.append(f"- 正文问题: {total_body_problems}个")
            report_lines.append("\n**建议**: 请检查字体、字号、对齐方式、首行缩进、行距等格式设置")
        
        report_text = '\n'.join(report_lines)
        
        # 输出到控制台
        print("\n" + "="*60)
        print("检测报告")
        print("="*60)
        print(report_text)
        
        # 保存报告
        if report_path:
            with open(report_path, 'w', encoding='utf-8') as f:
                f.write(report_text)
            print(f"\n报告已保存至: {report_path}")
        
        return report_text


def main():
    """主函数"""
    import os
    import sys
    current_dir = os.path.dirname(os.path.abspath(__file__))
    base_dir = os.path.dirname(current_dir)
    # 使用样本文件夹中的文件进行测试
    sample_dir = os.path.join(base_dir, "样本")
    docx_path = os.path.join(sample_dir, "毕业论文-董炜博.docx")

    if not os.path.exists(docx_path):
        print(f"[错误] 文件不存在：{docx_path}")
        print(f"   当前目录：{current_dir}")
        sys.exit(1)

    try:
        checker = AckAppendixChecker()
        reports_dir = os.path.join(base_dir, "reports")
        os.makedirs(reports_dir, exist_ok=True)
        report_path = os.path.join(reports_dir, "ack_appendix_report.md")
        checker.generate_report(docx_path, report_path)
    except Exception as e:
        print(f"[错误] 检测过程中出错：{str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()


