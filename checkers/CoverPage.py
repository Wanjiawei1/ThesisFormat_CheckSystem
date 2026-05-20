from docx import Document
from docx.oxml.ns import qn
import re

font_size_map = {
    "小二": 36,
    "三号": 16,
    "小三号": 15
}

def get_effective_font_name(run, paragraph):
    """
    获取有效的字体名称，优先从 XML 获取中文字体（eastAsia）
    """
    has_chinese = any('\u4e00' <= char <= '\u9fff' for char in run.text)
    
    try:
        if hasattr(run._element, 'rPr') and run._element.rPr is not None:
            rpr = run._element.rPr
            rfonts_elem = rpr.find(qn('w:rFonts'))
            if rfonts_elem is not None:
                if has_chinese:
                    eastasia = rfonts_elem.get(qn('w:eastAsia'))
                    if eastasia:
                        return eastasia
                    hAnsi = rfonts_elem.get(qn('w:hAnsi'))
                    if hAnsi:
                        return hAnsi
                else:
                    ascii_font = rfonts_elem.get(qn('w:ascii'))
                    if ascii_font:
                        return ascii_font
                    hAnsi = rfonts_elem.get(qn('w:hAnsi'))
                    if hAnsi:
                        return hAnsi
    except:
        pass
    
    if run.font.name:
        return run.font.name
    
    try:
        if paragraph.style.font.name:
            return paragraph.style.font.name
    except:
        pass
    
    return "None"

def get_font_size(run, paragraph):
    """获取字号（pt）"""
    try:
        if hasattr(run._element, 'rPr') and run._element.rPr is not None:
            rpr = run._element.rPr
            # 先尝试读取 sz 标签（西文字号）
            sz_elem = rpr.find(qn('w:sz'))
            if sz_elem is not None and sz_elem.get(qn('w:val')):
                try:
                    return int(sz_elem.get(qn('w:val'))) / 2
                except:
                    pass
            # 再尝试读取 szCs 标签（中文字号）
            szcs_elem = rpr.find(qn('w:szCs'))
            if szcs_elem is not None and szcs_elem.get(qn('w:val')):
                try:
                    return int(szcs_elem.get(qn('w:val'))) / 2
                except:
                    pass
    except:
        pass
    
    if run.font.size:
        return run.font.size.pt
    if run.style and run.style.font and run.style.font.size:
        return run.style.font.size.pt
    if paragraph and paragraph.style and paragraph.style.font and paragraph.style.font.size:
        return paragraph.style.font.size.pt
    return None

def is_run_bold(run, paragraph):
    """
    检测运行块是否加粗
    """
    if run.bold is True:
        return True
    if run.bold is False:
        return False
    
    try:
        if hasattr(run._element, 'rPr') and run._element.rPr is not None:
            rpr = run._element.rPr
            b_elem = rpr.find(qn('w:b'))
            if b_elem is not None:
                return True
    except:
        pass
    
    try:
        if paragraph and paragraph.style and paragraph.style.font:
            if paragraph.style.font.bold is True:
                return True
    except:
        pass
    
    return False

def check_font_and_bold(paragraph, expected_fonts, expected_bold, expected_size=None, field_name=""):
    """
    检查段落的字体、加粗和字号状态
    返回: (字体是否正确, 加粗是否正确, 字号是否正确, 错误原因列表)
    """
    reasons = []
    font_ok = True
    bold_ok = True
    size_ok = True
    
    if not paragraph or not paragraph.text.strip():
        return True, True, True, []
    
    content_runs = [run for run in paragraph.runs if run.text.strip()]
    if not content_runs:
        return True, True, True, []
    
    font_names = set()
    all_bold = True
    has_bold_info = False
    font_sizes = []
    
    for run in content_runs:
        run_text = run.text.strip()
        
        # 忽略各种字段的标签文字
        skip_run = False
        
        # 论文题目标签 - 灵活匹配
        if field_name == "论文题目" and any(run_text.startswith(label) for label in ["题目", "题目：", "题目:"]):
            skip_run = True
        # 学院标签 - 灵活匹配
        elif field_name == "学院" and any(run_text.startswith(label) for label in ["学院", "学院：", "学院:", "学"]):
            skip_run = True
        # 班级标签 - 灵活匹配
        elif field_name == "班级" and any(run_text.startswith(label) for label in ["班级", "班级：", "班级:", "班"]):
            skip_run = True
        # 学号标签 - 灵活匹配
        elif field_name == "学号" and any(run_text.startswith(label) for label in ["学号", "学号：", "学号:", "学"]):
            skip_run = True
        # 学生姓名标签 - 灵活匹配
        elif field_name == "学生姓名" and any(run_text.startswith(label) for label in ["学生姓名", "学生姓名：", "学生姓名:", "姓名", "姓名：", "姓名:"]):
            skip_run = True
        # 指导老师标签 - 灵活匹配
        elif field_name == "指导老师" and any(run_text.startswith(label) for label in ["指导老师", "指导老师：", "指导老师:", "指导教师", "指导教师：", "指导教师:"]):
            skip_run = True
        # 专业标签 - 灵活匹配
        elif field_name == "专业" and any(run_text.startswith(label) for label in ["专业", "专业：", "专业:", "专"]):
            skip_run = True
        # 校外指导教师标签 - 灵活匹配
        elif field_name == "校外指导教师" and any(run_text.startswith(label) for label in ["校外指导教师", "校外指导教师：", "校外指导教师:"]):
            skip_run = True
        # 提交日期标签 - 灵活匹配
        elif field_name == "提交日期" and any(run_text.startswith(label) for label in ["提交日期", "提交日期：", "提交日期:"]):
            skip_run = True
        
        # 对于学院字段，特殊处理
        if field_name == "学院":
            if run_text == "学    院：":
                skip_run = True
            elif run_text == "信息工程学院":
                skip_run = False
        
        if skip_run:
            continue
        
        font_name = get_effective_font_name(run, paragraph)
        if font_name and font_name != "None":
            font_names.add(font_name)
        
        is_bold = is_run_bold(run, paragraph)
        if is_bold is not None:
            has_bold_info = True
            if not is_bold:
                all_bold = False
        
        size = get_font_size(run, paragraph)
        if size is not None:
            font_sizes.append(size)
    
    if font_names:
        font_matched = False
        for fn in font_names:
            if any(ef in fn or fn in ef for ef in expected_fonts):
                font_matched = True
                break
        if not font_matched:
            font_ok = False
            reasons.append(f"字体应为{'或'.join(expected_fonts)}，当前为：{', '.join(font_names)}")
    
    if has_bold_info and expected_bold and not all_bold:
        bold_ok = False
        reasons.append("应加粗但未加粗")
    
    if expected_size is not None and font_sizes:
        # 检查每个运行块的字号
        all_sizes_ok = True
        for size in font_sizes:
            if abs(size - expected_size) > 0.7:
                all_sizes_ok = False
                break
        if not all_sizes_ok:
            size_ok = False
            # 显示所有不同的字号
            unique_sizes = sorted(list(set(font_sizes)))
            sizes_str = ", ".join([f"{s:.1f}pt" for s in unique_sizes])
            reasons.append(f"字号应为{expected_size}pt，当前为：{sizes_str}")
    
    return font_ok, bold_ok, size_ok, reasons

def check_cover_page_final(docx_path, school_name="浙江工业大学"):
    doc = Document(docx_path)
    paragraphs = [p for p in doc.paragraphs if p.text.strip()]
    para_texts = [p.text.strip() for p in paragraphs]

    # 同时收集表格单元格中的段落（封面常用表格布局）
    table_paragraphs = []
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if p.text.strip():
                        table_paragraphs.append(p)
                        para_texts.append(p.text.strip())

    # 合并顶层段落和表格段落，取前20个非空段落作为封面区域
    all_cover_paras = paragraphs + table_paragraphs
    cover_paragraphs = all_cover_paras[:30]  # 多取一些以覆盖表格内容
    full_text = "\n".join(para_texts)

    result = {}

    cover_fields_config = {
        "学校名称": {
            "keywords": [school_name] if school_name else [],
            "fonts": ["黑体", "SimHei"],
            "bold": False,
            "size": 36
        },
        "论文题目": {
            "keywords": [],
            "fonts": ["黑体", "SimHei"],
            "bold": False,
            "size": 16
        },
        "学院": {
            "keywords": ["学院"],
            "fonts": ["宋体", "SimSun", "黑体", "SimHei"],
            "bold": False,
            "size": 15
        },
        "班级": {
            "keywords": ["班级"],
            "fonts": ["宋体", "SimSun", "黑体", "SimHei"],
            "bold": False,
            "size": 15
        },
        "学号": {
            "keywords": ["学号"],
            "fonts": ["宋体", "SimSun", "黑体", "SimHei"],
            "bold": False,
            "size": 15
        },
        "学生姓名": {
            "keywords": ["学生姓名", "姓名"],
            "fonts": ["宋体", "SimSun", "黑体", "SimHei"],
            "bold": False,
            "size": 15
        },
        "指导老师": {
            "keywords": ["指导老师", "指导教师"],
            "fonts": ["宋体", "SimSun", "黑体", "SimHei"],
            "bold": False,
            "size": 15
        },
        "专业": {
            "keywords": [],
            "fonts": ["宋体", "SimSun", "黑体", "SimHei"],
            "bold": False,
            "size": 15
        },
        "校外指导教师": {
            "keywords": ["校外指导教师"],
            "fonts": ["宋体", "SimSun", "黑体", "SimHei"],
            "bold": False,
            "size": 15
        },
        "提交日期": {
            "keywords": [],
            "fonts": ["宋体", "SimSun", "黑体", "SimHei"],
            "bold": False,
            "size": 15
        }
    }

    def find_paragraph_by_keywords(keywords):
        if not keywords:
            return None
        for p in cover_paragraphs:
            text = p.text.strip()
            for kw in keywords:
                pattern = re.sub(r"(.)", lambda m: re.escape(m.group(1)) + r"\s*", kw)
                if re.search(pattern, text):
                    return p
        return None

    for field, config in cover_fields_config.items():
        keywords = config["keywords"]
        expected_fonts = config["fonts"]
        expected_bold = config["bold"]
        
        content_ok = False
        style_ok = True
        reasons = []
        matched_para = None
        
        if field == "学校名称":
            # 优先查找字号较大的段落（通常是学校名称）
            candidate_paras = []
            for p in cover_paragraphs:
                text = p.text.strip()
                if text:
                    # 计算每个段落的平均字号
                    sizes = []
                    for run in p.runs:
                        size = get_font_size(run, p)
                        if size:
                            sizes.append(size)
                    if sizes:
                        avg_size = sum(sizes) / len(sizes)
                        candidate_paras.append((avg_size, p))
            # 按字号降序排序
            if candidate_paras:
                candidate_paras.sort(key=lambda x: x[0], reverse=True)
                # 选择字号最大的段落作为学校名称
                matched_para = candidate_paras[0][1]
                content_ok = True
                # 同时检查是否包含学校名称文本
                if school_name and school_name in matched_para.text:
                    # 如果包含学校名称，更好
                    pass
                else:
                    # 即使不包含学校名称文本，也选择字号最大的段落
                    # 因为有些封面的学校名称可能是图片或者特殊格式
                    pass
        elif field == "论文题目":
            # 先找到学校名称的段落（字号最大的段落）
            school_para = None
            max_size = 0
            for p in cover_paragraphs:
                text = p.text.strip()
                if text:
                    sizes = []
                    for run in p.runs:
                        size = get_font_size(run, p)
                        if size:
                            sizes.append(size)
                    if sizes:
                        avg_size = sum(sizes) / len(sizes)
                        if avg_size > max_size:
                            max_size = avg_size
                            school_para = p
            
            # 如果找到了学校名称段落，在其后找论文题目
            if school_para:
                school_idx = cover_paragraphs.index(school_para)
                if school_idx + 1 < len(cover_paragraphs):
                    # 从学校名称之后开始查找
                    for i in range(school_idx + 1, min(school_idx + 5, len(cover_paragraphs))):
                        p = cover_paragraphs[i]
                        text = p.text.strip()
                        if text and len(text) > 5:  # 论文题目通常较长
                            matched_para = p
                            content_ok = True
                            break
            # 如果还没找到，尝试找学校名称段落之后的几个段落
            if not matched_para and school_para:
                school_idx = cover_paragraphs.index(school_para)
                # 尝试合并学校名称之后的几个段落作为论文题目
                if school_idx + 1 < len(cover_paragraphs):
                    # 取学校名称之后的2-3个段落作为论文题目
                    for i in range(school_idx + 1, min(school_idx + 4, len(cover_paragraphs))):
                        p = cover_paragraphs[i]
                        text = p.text.strip()
                        if text and len(text) > 5:
                            matched_para = p
                            content_ok = True
                            break
            # 最后的 fallback
            if not matched_para and cover_paragraphs:
                # 尝试找第二个非空段落
                non_empty_paras = [p for p in cover_paragraphs if p.text.strip()]
                if len(non_empty_paras) > 1:
                    matched_para = non_empty_paras[1]
                    content_ok = len(matched_para.text.strip()) > 5
                elif non_empty_paras:
                    matched_para = non_empty_paras[0]
                    content_ok = len(matched_para.text.strip()) > 5
        elif field == "专业":
            # 优先通过关键词"专业"查找
            for p in cover_paragraphs:
                text = p.text.strip()
                # 检查是否包含"专业"或"专"和"业"
                if "专业" in text or ("专" in text and "业" in text):
                    matched_para = p
                    content_ok = True
                    break
            # 如果没有找到，使用原来的逻辑作为fallback
            if not matched_para:
                used_texts = set()
                for p in cover_paragraphs:
                    text = p.text.strip()
                    if (
                        text not in used_texts
                        and 3 <= len(text) <= 20
                        and not any(k in text for k in ["浙江工业大学", "论文", "指导", "学号", "班级", "学院", "姓名", "题目", "提交", "学生"])
                    ):
                        matched_para = p
                        break
                content_ok = matched_para is not None
        elif field == "校外指导教师":
            matched_para = find_paragraph_by_keywords(keywords)
            content_ok = matched_para is not None
            if not content_ok:
                result[field] = {
                    "内容": False,
                    "样式": None,
                    "原因": ""
                }
                continue
        elif field == "提交日期":
            date_match = re.search(r"\d{4}年\d{1,2}月", full_text)
            content_ok = bool(date_match)
            if content_ok:
                for p in cover_paragraphs:
                    if re.search(r"\d{4}年\d{1,2}月", p.text):
                        matched_para = p
                        break
        else:
            matched_para = find_paragraph_by_keywords(keywords)
            content_ok = matched_para is not None
        
        if content_ok and matched_para:
            font_ok, bold_ok, size_ok, font_reasons = check_font_and_bold(matched_para, expected_fonts, expected_bold, config.get("size"), field)
            if not font_ok or not bold_ok or not size_ok:
                style_ok = False
                reasons = font_reasons
        
        result[field] = {
            "内容": content_ok,
            "样式": style_ok,
            "原因": "，".join(reasons) if reasons else ""
        }

    print("\n=== 封面格式检测结果 ===")
    for field, status in result.items():
        if field == "校外指导教师" and not status["内容"]:
            continue
        c = "[OK]" if status["内容"] else "[FAIL]"
        s = "[OK]" if status["样式"] else "[FAIL]" if status["样式"] is not None else "（未检测）"
        reason = f" - {status['原因']}" if status['原因'] else ""
        print(f"- {field}：内容 {c} / 样式 {s}{reason}")

    return result


def check_english_names(docx_path, school_name="浙江工业大学"):
    """检测英文封面页的姓名格式是否符合名在前、姓在后的规范"""
    doc = Document(docx_path)
    found = {"Student": None, "Advisor": None}
    issues = []

    # 查找英文姓名行
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        m_student = re.match(r'^Student\s*:\s*(.+)', text, re.IGNORECASE)
        if m_student:
            found["Student"] = (i, m_student.group(1).strip())
        m_advisor = re.match(r'^(Advisor|Supervisor|Adviser)\s*:\s*(.+)', text, re.IGNORECASE)
        if m_advisor:
            found["Advisor"] = (i, m_advisor.group(2).strip())

    for label, info in found.items():
        if info is None:
            issues.append(f"[WARN] 未找到英文{label}姓名行")
            continue

        p_idx, name_str = info
        # 去除头衔前缀
        name_str = re.sub(r'^(Prof\.|Dr\.|Mr\.|Ms\.|Mrs\.|Miss\.)\s*', '', name_str, flags=re.IGNORECASE).strip()

        # 提取英文姓名单词（取前两个大写开头的单词作为姓和名）
        words = name_str.split()
        if len(words) < 2:
            continue

        name_words = []
        for w in words:
            if re.match(r'^[A-Z][a-z]+$', w):
                name_words.append(w)
            if len(name_words) >= 2:
                break

        if len(name_words) < 2:
            continue

        first, second = name_words[0], name_words[1]

        # 常见中文姓氏拼音
        common_surnames = {
            'Wang', 'Zhang', 'Li', 'Liu', 'Chen', 'Yang', 'Huang', 'Zhao', 'Wu',
            'Zhou', 'Xu', 'Sun', 'Ma', 'Zhu', 'Hu', 'Guo', 'He', 'Gao', 'Lin',
            'Luo', 'Zheng', 'Liang', 'Xie', 'Song', 'Tang', 'Han', 'Cao', 'Deng',
            'Xiao', 'Feng', 'Zeng', 'Cheng', 'Cai', 'Pan', 'Yuan', 'Yu', 'Dong',
            'Su', 'Ye', 'Lu', 'Wei', 'Jiang', 'Tian', 'Du', 'Ding', 'Shen', 'Ren',
            'Yao', 'Fang', 'Jin', 'Xia', 'Qiu', 'Tan', 'Jia', 'Zou', 'Shi', 'Xiong',
            'Meng', 'Qin', 'Yan', 'Xue', 'Hou', 'Lei', 'Bai', 'Long', 'Wan',
            'Duan', 'Qian', 'Yin', 'Mao', 'Chang', 'Wen', 'Lai', 'Gong', 'Hong',
            'Bao', 'Shu', 'Gu', 'Ji', 'Nie',
        }

        first_is_surname = first in common_surnames
        second_is_surname = second in common_surnames
        first_shorter = len(first) <= len(second)

        if first_is_surname and first_shorter and not second_is_surname:
            issues.append(
                f"[FAIL] 英文{label}姓名格式错误：\"{first} {second}\" 应为名在前、姓在后，"
                f"正确格式应为 \"{second} {first}\""
            )

    if not issues:
        issues.append("[OK] 英文姓名格式符合规范（名在前、姓在后）")

    return {
        "found": found,
        "issues": issues,
    }


if __name__ == '__main__':
    import os
    current_dir = os.path.dirname(os.path.abspath(__file__))
    base_dir = os.path.dirname(current_dir)
    sample_dir = os.path.join(base_dir, "样本")
    docx_path = os.path.join(sample_dir, "毕业论文-万嘉玮.docx")
    if os.path.exists(docx_path):
        check_cover_page_final(docx_path)
    else:
        print(f"文件不存在: {docx_path}")
