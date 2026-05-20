"""
正文格式检测模块
功能：
1. 检查正文段落的字体、字号、行距、段落缩进等格式
2. 验证页码格式和位置
3. 生成详细的检测报告
"""

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt
import re


class BodyTextChecker:
    """正文格式检测器"""
    
    def __init__(self):
        # 正文标准格式要求（小四号=12pt）
        self.body_format = {
            'chinese_font': ['宋体', 'SimSun'],
            'english_font': ['Times New Roman'],
            'font_size': 12.0,  # 小四号 = 12pt
            'line_spacing': 1.5,  # 1.5倍行距
            'first_line_indent': 24.0,  # 首行缩进2字符 = 12*2 = 24pt
            'alignment': 'JUSTIFY',  # 两端对齐
            'space_before': 0,  # 段前0行
            'space_after': 0  # 段后0行
        }
        
        # 标题格式标准
        self.title_formats = {
            'level1': {  # 一级标题（章标题）：第X章
                'font': ['黑体', 'SimHei'],
                'size': 16.0,  # 三号 = 16pt
                'bold': True,
                'alignment': 'CENTER',
                'line_spacing': 1.5,
                'space_before': 0,
                'space_after': 0
            },
            'level2': {  # 二级标题（节标题）：X.X
                'font': ['黑体', 'SimHei'],
                'size': 14.0,  # 四号 = 14pt
                'bold': False,
                'alignment': 'LEFT',
                'line_spacing': 1.5,
                'space_before': 1,  # 段前1行
                'space_after': 0.5  # 段后0.5行
            },
            'level3': {  # 三级标题（小节标题）：X.X.X
                'font': ['黑体', 'SimHei'],
                'size': 12.0,  # 小四号 = 12pt
                'bold': False,
                'alignment': 'LEFT',
                'line_spacing': 1.5,
                'space_before': 1,  # 段前1行
                'space_after': 0.5  # 段后0.5行
            }
        }
        
        # 保持向后兼容
        self.standard_format = self.body_format
        
        # 页面格式标准
        self.page_format = {
            'paper_size': 'A4',
            'margin_top': 2.54,  # 厘米
            'margin_bottom': 2.54,
            'margin_left': 2.54,
            'margin_right': 2.54,
            'header_distance': 1.6,  # 厘米
            'footer_distance': 1.6
        }
        
        # 排除的段落特征（非正文段落）
        self.exclude_patterns = [
            r'^第[一二三四五六七八九十\d]+章',  # 章标题
            r'^\d+\.\d+',  # 小节标题
            r'^摘\s*要',  # 摘要
            r'^关键词',  # 关键词
            r'^Abstract',  # 英文摘要
            r'^Keywords',  # 英文关键词
            r'^参考文献',  # 参考文献
            r'^致\s*谢',  # 致谢
            r'^附\s*录',  # 附录
            r'^\[?\d+\]',  # 参考文献条目
            r'^图\s*\d+',  # 图标题
            r'^表\s*\d+',  # 表标题
            r'^目\s*录',  # 目录
            r'.*[：:]\s*$',  # 以冒号结尾（通常是标签）
            r'^题目',  # 题目标签
            r'^学\s*院',  # 学院
            r'^专\s*业',  # 专业
            r'^班\s*级',  # 班级
            r'^学\s*号',  # 学号
            r'^学生姓名',  # 学生姓名
            r'^指导',  # 指导老师
            r'^提交日期',  # 提交日期
            r'^\d{4}年',  # 年份格式
            r'^浙江工业大学',  # 学校名称
        ]
    
    def get_font_name(self, run, paragraph=None):
        """获取运行块的字体名称（优先获取中文字体）"""
        # 检查是否有中文字符
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
                    else:
                        ascii_font = rfonts_elem.get(qn('w:ascii'))
                        if ascii_font:
                            return ascii_font
                    # 回退到 hAnsi
                    hAnsi = rfonts_elem.get(qn('w:hAnsi'))
                    if hAnsi:
                        return hAnsi
        except:
            pass
        
        # 回退到标准属性
        if run.font.name:
            return run.font.name
        
        # 从段落样式获取（特别是中文字体）
        if paragraph:
            try:
                if paragraph.style:
                    # 尝试从样式的rPr获取字体
                    style_element = paragraph.style.element
                    pPr = style_element.find(qn('w:pPr'))
                    if pPr is not None:
                        rPr = pPr.find(qn('w:rPr'))
                        if rPr is not None:
                            rfonts = rPr.find(qn('w:rFonts'))
                            if rfonts is not None:
                                if has_chinese:
                                    eastasia = rfonts.get(qn('w:eastAsia'))
                                    if eastasia:
                                        return eastasia
                                ascii_font = rfonts.get(qn('w:ascii'))
                                if ascii_font:
                                    return ascii_font
                    
                    # 回退到样式的font属性
                    if paragraph.style.font and paragraph.style.font.name:
                        return paragraph.style.font.name
            except:
                pass
        
        # 假设使用宋体作为默认中文字体
        if has_chinese:
            return "宋体"
        
        return "Times New Roman"
    
    def get_font_size(self, run, paragraph=None):
        """获取字号（单位：磅）"""
        # 从 run 获取
        if run.font.size:
            return run.font.size.pt
        
        # 从段落样式获取
        if paragraph:
            try:
                if paragraph.style and paragraph.style.font:
                    if paragraph.style.font.size:
                        return paragraph.style.font.size.pt
            except:
                pass
        
        # 根据段落文本判断标题级别并返回相应的默认值
        if paragraph:
            text = paragraph.text.strip()
            if re.match(r'^第[一二三四五六七八九十\d]+章', text):
                return 16.0  # 一级标题默认16pt
            elif re.match(r'^\d+\.\d+\.\d+', text):
                return 12.0  # 三级标题默认12pt
            elif re.match(r'^\d+\.\d+', text):
                return 14.0  # 二级标题默认14pt
        
        # 返回默认值（假设是14pt）
        return 14.0
    
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
        
        # 方法1.5：从段落样式获取首行缩进
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
        
        # 方法2：检查左缩进（可能是用左缩进代替首行缩进）
        try:
            left_indent = paragraph.paragraph_format.left_indent
            if left_indent and left_indent.pt > 0:
                return left_indent.pt
        except:
            pass
        
        # 方法3：检查段落开头的空格或制表符（手动缩进）
        text = paragraph.text
        if text:
            # 检查开头的空格数（全角空格算2个字符缩进）
            leading_spaces = 0
            for char in text:
                if char == ' ':  # 半角空格
                    leading_spaces += 0.5
                elif char == '　':  # 全角空格
                    leading_spaces += 1
                elif char == '\t':  # 制表符（重要：很多文档用制表符缩进）
                    # 制表符通常表示2个字符的缩进
                    return 28.0  # 直接返回标准的2字符缩进值
                else:
                    break
            
            # 如果有至少1个字符的空格缩进，按比例计算
            # 按每个字符14pt计算（因为字号是14pt）
            if leading_spaces >= 1:
                return leading_spaces * 14.0
            
            # 也检查run level的缩进（有时候缩进在run级别）
            if paragraph.runs:
                first_run_text = paragraph.runs[0].text if paragraph.runs[0].text else ""
                for char in first_run_text:
                    if char == '\t':  # 制表符
                        return 28.0  # 制表符表示标准的2字符缩进
                    elif char == '　':  # 全角空格
                        return 28.0  # 2个全角空格也是标准缩进
                    elif char != ' ':
                        break  # 遇到非空格字符就停止
                
                # 计算半角空格数
                run_leading_spaces = 0
                for char in first_run_text:
                    if char == ' ':  # 半角空格
                        run_leading_spaces += 0.5
                    else:
                        break
                if run_leading_spaces >= 1:
                    return run_leading_spaces * 14.0
        
        # 方法4：检查悬挂缩进（hanging indent）
        try:
            hanging = paragraph.paragraph_format.hanging_indent
            if hanging and hanging.pt > 0:
                # 悬挂缩进是负的首行缩进，这里返回负值
                return -hanging.pt
        except:
            pass
        
        return 0
    
    def get_alignment(self, paragraph):
        """获取对齐方式"""
        alignment_map = {
            0: 'LEFT',
            1: 'CENTER',
            2: 'RIGHT',
            3: 'JUSTIFY',
            4: 'DISTRIBUTE',
            None: 'DEFAULT'
        }
        return alignment_map.get(paragraph.alignment, 'UNKNOWN')
    
    def identify_title_level(self, paragraph):
        """识别标题级别
        返回: 'level1', 'level2', 'level3', 或 None（不是标题）
        """
        text = paragraph.text.strip()
        
        # 一级标题：第X章（不含逗号和句号）
        if re.match(r'^第[一二三四五六七八九十\d]+章', text) and '，' not in text and '。' not in text:
            return 'level1'
        
        # 三级标题：X.X.X（如2.3.4），允许数字后面没有空格
        if re.match(r'^\d+\.\d+\.\d+', text):
            return 'level3'
        
        # 二级标题：X.X（如2.3），允许数字后面没有空格
        if re.match(r'^\d+\.\d+', text):
            # 确保是标题而不是目录项（目录项会有页码）
            if not re.search(r'\s+\d+$', text):
                return 'level2'
        
        return None
    
    def is_body_paragraph(self, paragraph):
        """判断段落是否为正文段落（改进版：更宽松的规则）"""
        text = paragraph.text.strip()
        
        # 跳过空段落
        if not text:
            return False
        
        # 只排除明显的非正文内容
        exclude_strict = [
            r'^第[一二三四五六七八九十\d]+章',  # 章标题本身
            r'^\d+\.\d+',  # 小节标题编号（如1.1、2.3等）和目录项
            r'^摘\s*要$',  # 摘要标题
            r'^关键词[:：]',  # 关键词
            r'^Abstract$',  # 英文摘要标题
            r'^Keywords[:：]',  # 英文关键词
            r'^参考文献$',  # 参考文献标题
            r'^致\s*谢$',  # 致谢标题
            r'^附\s*录',  # 附录
            r'^目\s*录$',  # 目录标题
            r'^浙江工业大学',  # 学校名称（封面）
            r'^题\s*目[:：]',  # 题目标签
            r'^学\s*院[:：]',  # 学院标签
            r'^专\s*业[:：]',  # 专业标签
            r'^班\s*级[:：]',  # 班级标签
            r'^学\s*号[:：]',  # 学号标签
            r'^学生姓名[:：]',  # 学生姓名标签
            r'^指导老师[:：]',  # 指导老师标签
            r'^提交日期[:：]',  # 提交日期标签
        ]
        
        # 严格排除这些内容
        for pattern in exclude_strict:
            if re.match(pattern, text, re.IGNORECASE):
                return False
        
        # 跳过图表子标签，如 "(a) xxx"、"（b）xxx"、"图X-Y"、"表X-Y"、"Figure X"、"Table X"
        if re.match(r'^[（(]\s*[a-zA-Z]\s*[)）]', text):
            return False
        if re.match(r'^(图|表)\s*\d+[-\.]\d+', text):
            return False
        if re.match(r'^(Figure|Table)\s*\d+[-\.]\d+', text, re.IGNORECASE):
            return False
        
        # 跳过纯数字或过短的段落（小于10个字符）
        if len(text) < 10:
            return False
        
        # 跳过字号过大的段落（可能是一级标题，>16pt）
        for run in paragraph.runs:
            if run.font.size and run.font.size.pt > 16:
                return False

        # 跳过编号列表段落（Word 自动编号，不要求首行缩进）
        try:
            pPr = paragraph._element.find(qn('w:pPr'))
            if pPr is not None and pPr.find(qn('w:numPr')) is not None:
                return False
        except:
            pass

        # 其他都认为是正文（包括没有标点符号的段落）
        return True
    
    def check_title_format(self, paragraph, level):
        """检查标题格式
        Args:
            paragraph: 段落对象
            level: 标题级别 ('level1', 'level2', 'level3')
        """
        issues = []
        standard = self.title_formats[level]
        
        # 获取段落格式
        alignment = self.get_alignment(paragraph)
        line_spacing = self.get_line_spacing(paragraph)
        
        # 检查字体和字号
        if paragraph.runs:
            # 对于二级和三级标题，需要跳过数字编号部分（如"1.1 "），检查标题文字
            target_run = paragraph.runs[0]
            text = paragraph.text.strip()
            
            # 如果是二级或三级标题（以数字开头），找到编号后的文字部分
            if level in ['level2', 'level3'] and re.match(r'^\d+\.', text):
                # 遍历runs，找到包含中文或非数字内容的run
                for run in paragraph.runs:
                    run_text = run.text.strip()
                    if run_text and not re.match(r'^[\d\.\s]+$', run_text):
                        # 这个run包含实际的标题文字，而不只是数字
                        target_run = run
                        break
            
            font_name = self.get_font_name(target_run, paragraph)
            font_size = self.get_font_size(target_run, paragraph)
            
            # 检查字体（黑体）
            # 对于二级和三级标题，不检查字体（因为数字编号部分是Times New Roman）
            if level == 'level1':
                if not any(font in font_name for font in standard['font']):
                    issues.append(f"字体应为黑体，当前为：{font_name}")
            
            # 检查字号
            if font_size and abs(font_size - standard['size']) > 0.5:
                issues.append(f"字号应为{standard['size']:.0f}pt，当前为：{font_size:.1f}pt")
            
            # 检查加粗（只在明确设置为False时才报错，None表示继承自样式可能已加粗）
            if level == 'level1' and target_run.bold is False:
                issues.append("一级标题应加粗")
        
        # 检查对齐方式（DEFAULT表示继承自样式，可能符合要求）
        expected_align = standard['alignment']
        if expected_align == 'CENTER' and alignment not in ['CENTER', 'DEFAULT']:
            issues.append(f"对齐方式应为居中，当前为：{alignment}")
        elif expected_align == 'LEFT' and alignment not in ['LEFT', 'JUSTIFY', 'DEFAULT']:
            issues.append(f"对齐方式应为左对齐，当前为：{alignment}")
        
        return issues
    
    def check_body_format(self, paragraph):
        """检查正文格式"""
        issues = []
        
        # 获取段落格式信息
        alignment = self.get_alignment(paragraph)
        first_line_indent = self.get_first_line_indent(paragraph)
        line_spacing = self.get_line_spacing(paragraph)
        
        # 检查对齐方式
        if alignment not in ['JUSTIFY', 'DEFAULT']:
            issues.append(f"对齐方式应为两端对齐，当前为：{alignment}")
        
        # 检查首行缩进（2字符=24pt，允许±2pt误差）
        # 跳过列表项段落（如"XXX相关：通过XXX函数..."这类编号列表内容）
        # 判断标准：以中文开头，前20字符内出现中文冒号"："作为标签分隔符
        text = paragraph.text.strip()
        colon_pos = text.find('：')
        is_list_item = (colon_pos > 0 and colon_pos <= 20
                        and bool(re.match(r'^[\u4e00-\u9fff]', text)))
        if not is_list_item:
            if first_line_indent is None or first_line_indent < 10:
                issues.append("首行缩进未设置或过小")
            elif abs(first_line_indent - self.body_format['first_line_indent']) > 2.0:
                issues.append(f"首行缩进应为2字符（24pt），当前为：{first_line_indent:.1f}pt")
        
        # 检查字体和字号
        if paragraph.runs:
            # 判断段落是否包含中文
            has_chinese = any('\u4e00' <= char <= '\u9fff' for char in paragraph.text)
            
            # 找到包含中文的run来检查字体（避免检查数字/字母开头的run）
            target_run = None
            if has_chinese:
                for run in paragraph.runs:
                    if any('\u4e00' <= char <= '\u9fff' for char in run.text):
                        target_run = run
                        break
            
            # 如果没有找到包含中文的run，使用第一个run
            if target_run is None:
                target_run = paragraph.runs[0]
            
            font_name = self.get_font_name(target_run, paragraph)
            font_size = self.get_font_size(target_run, paragraph)
            
            # 检查字体（只在有中文且找到了包含中文的run时检查）
            # 如果字体是通过样式继承的默认值，不报错
            if has_chinese and target_run and any('\u4e00' <= char <= '\u9fff' for char in target_run.text):
                # 检查是否能明确获取到字体（不是默认假设值）
                has_explicit_font = False
                try:
                    if hasattr(target_run._element, 'rPr') and target_run._element.rPr is not None:
                        rpr = target_run._element.rPr
                        rfonts_elem = rpr.find(qn('w:rFonts'))
                        if rfonts_elem is not None and rfonts_elem.get(qn('w:eastAsia')):
                            has_explicit_font = True
                except:
                    pass
                
                # 只在明确获取到非宋体字体时才报错
                if has_explicit_font and not any(font in font_name for font in self.body_format['chinese_font']):
                    issues.append(f"中文字体应为宋体，当前为：{font_name}")
            
            # 检查字号（小四号=12pt）
            # 只在明确获取到字号（不是默认假设值）时才检查
            has_explicit_size = False
            if target_run.font.size:
                has_explicit_size = True
            
            if has_explicit_size and font_size and abs(font_size - self.body_format['font_size']) > 0.5:
                issues.append(f"字号应为小四号（12pt），当前为：{font_size:.1f}pt")
        
        return issues
    
    def estimate_page_and_line(self, doc, paragraph_index, paragraph):
        """估计段落的页码和行号
        注意：这是基于段落位置和内容的估计，不是精确值
        """
        # 更合理的页码估计：基于正文开始后的段落数
        # 假设每页平均有15个正文段落（考虑到标题和空白）
        estimated_page = (paragraph_index // 15) + 1
        
        # 更合理的行号估计：基于段落内容长度
        text = paragraph.text
        # 假设每行平均30个字符
        chars_per_line = 30
        estimated_lines_in_paragraph = max(1, len(text) // chars_per_line)
        
        # 计算当前页的起始段落索引
        page_start_index = ((estimated_page - 1) * 15) + 1
        # 计算当前段落在当前页的位置
        position_in_page = paragraph_index - page_start_index + 1
        
        # 计算行号：前面段落的估计行数之和 + 当前段落的估计行数
        estimated_line = 1
        # 假设前面每个段落平均2行
        estimated_line += (position_in_page - 1) * 2
        
        return estimated_page, estimated_line
    
    def check_body_text(self, docx_path):
        """检查整个文档的正文区域（包括标题和正文）"""
        doc = Document(docx_path)
        
        results = {
            'total_paragraphs': 0,
            'title_count': {
                'level1': 0,
                'level2': 0,
                'level3': 0
            },
            'body_count': 0,
            'title_issues': [],
            'body_issues': [],
            'title_samples': [],
            'body_samples': [],
            'all_body_issues': []  # 存储所有正文问题
        }
        
        # 标记是否进入正文区域
        in_body_section = False
        body_paragraph_index = 0  # 正文区域内的段落索引
        
        # 跟踪当前章节位置
        current_chapter = ""
        current_section = ""
        current_subsection = ""
        section_paragraph_count = 0  # 当前节下的段落计数
        
        # 遍历所有段落
        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            
            # 跳过空段落
            if not text:
                continue
            
            # 检测正文开始：第1章或第一章（但排除目录中的章标题）
            if not in_body_section:
                if re.match(r'^第[1一]章', text):
                    if not re.search(r'\s+\d+$', text):  # 不是目录项
                        in_body_section = True
                    else:
                        continue
                else:
                    continue
            
            # 检测正文结束：参考文献或致谢
            if in_body_section and (re.match(r'^参\s*考\s*文\s*献', text) or re.match(r'^致\s*谢', text)):
                break
            
            # 识别段落类型
            title_level = self.identify_title_level(para)
            
            if title_level:
                # 这是标题，更新章节信息
                results['total_paragraphs'] += 1
                results['title_count'][title_level] += 1
                
                # 更新章节信息
                if title_level == 'level1':
                    current_chapter = text
                    current_section = ""
                    current_subsection = ""
                    section_paragraph_count = 0
                elif title_level == 'level2':
                    current_section = text
                    current_subsection = ""
                    section_paragraph_count = 0
                elif title_level == 'level3':
                    current_subsection = text
                    section_paragraph_count = 0
                
                # 检查标题格式
                issues = self.check_title_format(para, title_level)
                
                if issues:
                    results['title_issues'].extend(issues)
                    results['title_samples'].append({
                        'paragraph_index': i,
                        'level': title_level,
                        'text_preview': text[:50] + '...' if len(text) > 50 else text,
                        'issues': issues
                    })
            else:
                # 这是正文，检查正文格式
                # 跳过图表标题等特殊段落
                if len(text) < 10:
                    continue
                if re.match(r'^(图|表|Figure|Table)\s*\d', text):
                    continue
                # 跳过图表子标签，如 "(a) xxx"、"（b）xxx"
                if re.match(r'^[（(]\s*[a-zA-Z]\s*[)）]', text):
                    continue
                
                body_paragraph_index += 1
                section_paragraph_count += 1
                results['total_paragraphs'] += 1
                results['body_count'] += 1
                issues = self.check_body_format(para)
                
                if issues:
                    results['body_issues'].extend(issues)
                    
                    # 构建章节位置信息
                    chapter_position = ""
                    if current_chapter:
                        chapter_position = current_chapter
                        if current_section:
                            chapter_position += f" - {current_section}"
                            if current_subsection:
                                chapter_position += f" - {current_subsection}"
                        chapter_position += f" 下的第 {section_paragraph_count} 个段落"
                    
                    # 保存所有正文问题，不仅仅是前5个
                    results['all_body_issues'].append({
                        'paragraph_index': i + 1,  # 段落索引从1开始
                        'chapter_position': chapter_position,
                        'text_preview': text[:200] + '...' if len(text) > 200 else text,  # 更长的内容摘要
                        'issues': issues
                    })
        
        return results
    
    def _find_section_boundaries(self, doc: Document, start_patterns: list, end_patterns: list, start_from: int = 0) -> tuple:
        """在文档中定位一个区域：从 start_patterns 匹配行到 end_patterns 匹配行。
        返回 (start_idx, end_idx)，未找到返回 (-1, -1)
        匹配时忽略空白和大小写。"""
        import re as _re
        start_idx = -1
        for i in range(start_from, len(doc.paragraphs)):
            text = _re.sub(r'\s+', '', doc.paragraphs[i].text.strip())
            for pat in start_patterns:
                if _re.match(pat, text, _re.IGNORECASE):
                    start_idx = i
                    break
            if start_idx >= 0:
                break
        if start_idx < 0:
            return -1, -1

        end_idx = len(doc.paragraphs)
        for i in range(start_idx + 1, len(doc.paragraphs)):
            text = _re.sub(r'\s+', '', doc.paragraphs[i].text.strip())
            for pat in end_patterns:
                if _re.match(pat, text, _re.IGNORECASE):
                    end_idx = i
                    break
            if end_idx < len(doc.paragraphs):
                break
        return start_idx, end_idx

    def check_abstract_indentation(self, doc: Document) -> dict:
        """检查中英文摘要段落的缩进格式"""
        result = {
            "cn_abstract": {"found": False, "paragraph_count": 0, "issues": []},
            "en_abstract": {"found": False, "paragraph_count": 0, "issues": []},
        }

        ALL_CHAPTER = r'^第[1一]章'
        CN_KEYWORDS = r'^关键词'
        EN_KEYWORDS = r'^Keywords'

        # 中文摘要：从"摘要"到"关键词"/"ABSTRACT"/"第1章"
        cn_start, cn_end = self._find_section_boundaries(
            doc,
            start_patterns=[r'^摘要$', r'^摘要摘要$'],
            end_patterns=[CN_KEYWORDS, r'^ABSTRACT$', r'^Abstract$', ALL_CHAPTER])
        if cn_start >= 0 and cn_end >= 0:
            result["cn_abstract"]["found"] = True
            for i in range(cn_start + 1, cn_end):
                para = doc.paragraphs[i]
                text = para.text.strip()
                if not text or len(text) < 10:
                    continue
                if re.match(r'^关键词|^摘要', re.sub(r'\s+', '', text)):
                    continue
                indent = self.get_first_line_indent(para)
                result["cn_abstract"]["paragraph_count"] += 1
                if indent is None or indent < 10:
                    result["cn_abstract"]["issues"].append({
                        "paragraph_index": i + 1,
                        "text_preview": text[:80] + ('...' if len(text) > 80 else ''),
                        "actual_indent": indent or 0,
                    })

        # 英文摘要：从"ABSTRACT"到"Keywords"/"第1章"（大小写不敏感）
        en_start, en_end = self._find_section_boundaries(
            doc,
            start_patterns=[r'^ABSTRACT$', r'^Abstract$'],
            end_patterns=[EN_KEYWORDS, ALL_CHAPTER])
        if en_start >= 0 and en_end >= 0:
            result["en_abstract"]["found"] = True
            for i in range(en_start + 1, en_end):
                para = doc.paragraphs[i]
                text = para.text.strip()
                if not text or len(text) < 10:
                    continue
                # 跳过 Keywords / KEY WORDS 行（大小写不敏感），关键词行不需要缩进
                if re.match(r'^Keywords|^ABSTRACT|^Abstract', re.sub(r'\s+', '', text), re.IGNORECASE):
                    continue
                indent = self.get_first_line_indent(para)
                result["en_abstract"]["paragraph_count"] += 1
                if indent is None or indent < 10:
                    result["en_abstract"]["issues"].append({
                        "paragraph_index": i + 1,
                        "text_preview": text[:80] + ('...' if len(text) > 80 else ''),
                        "actual_indent": indent or 0,
                    })

        return result

    def _append_abstract_report(self, report_lines: list, abstract_result: dict, key: str, label: str):
        """将摘要检测结果追加到报告行列表中"""
        section = abstract_result.get(key, {})
        report_lines.append(f"### {label}")
        if not section.get("found"):
            report_lines.append(f"[警告] 未找到{label}区域")
            return
        issues = section.get("issues", [])
        if not issues:
            report_lines.append(f"[通过] {label}缩进格式符合要求")
        else:
            report_lines.append(f"[警告] **发现 {len(issues)} 个段落首行未缩进2字符**")
            for iss in issues:
                report_lines.append(f"- 段落{iss['paragraph_index']}: 首行缩进应为24pt，实际{iss['actual_indent']:.0f}pt | 内容: \"{iss['text_preview']}\"")

    def check_extra_spaces(self, doc: Document) -> dict:
        """检测英文标点及字母两侧的多余空格"""
        issues = []
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
            return {"issues": ["[警告] 未能定位正文区域"]}

        for i in range(body_start, min(body_end, len(doc.paragraphs))):
            para = doc.paragraphs[i]
            text = para.text

            # 检测1: 英文单词与后续标点之间有多余空格 (如 "word ," 或 "word .")
            for m in re.finditer(r'([a-zA-Z])\s{2,}([,;:.!?])', text):
                ctx = text[max(0, m.start()-10):m.end()+10]
                issues.append({
                    "paragraph_index": i + 1,
                    "type": "英文单词与标点间多余空格",
                    "context": f'\"...{ctx}...\"',
                })

            # 检测2: 英文单词之间有多余空格（连续多个空格）
            for m in re.finditer(r'([a-zA-Z])\s{2,}([a-zA-Z])', text):
                ctx = text[max(0, m.start()-10):m.end()+10]
                issues.append({
                    "paragraph_index": i + 1,
                    "type": "英文单词间多余空格",
                    "context": f'\"...{ctx}...\"',
                })

            # 检测3: 左括号后多余空格 (如 "( word" 或 "(  word")
            for m in re.finditer(r'([\(（])\s{2,}([a-zA-Z一-鿿])', text):
                ctx = text[max(0, m.start()-5):m.end()+10]
                issues.append({
                    "paragraph_index": i + 1,
                    "type": "左括号后多余空格",
                    "context": f'\"...{ctx}...\"',
                })

            # 检测4: 右括号前多余空格 (如 "word )" 或 "word  )")
            for m in re.finditer(r'([a-zA-Z一-鿿])\s{2,}([\)）])', text):
                ctx = text[max(0, m.start()-10):m.end()+5]
                issues.append({
                    "paragraph_index": i + 1,
                    "type": "右括号前多余空格",
                    "context": f'\"...{ctx}...\"',
                })

        # 去重（同一段落同类问题只报一次）
        seen = set()
        unique_issues = []
        for iss in issues:
            key = (iss["paragraph_index"], iss["type"])
            if key not in seen:
                seen.add(key)
                unique_issues.append(iss)

        return {
            "total_issues": len(unique_issues),
            "issues": unique_issues,
        }

    def check_page_numbers(self, docx_path):
        """检查页码格式（基于页眉页脚）"""
        doc = Document(docx_path)
        page_number_info = {
            'has_page_number': False,
            'header_page_number': False,
            'footer_page_number': False,
            'issues': []
        }
        
        try:
            # 检查页眉中的页码
            for section in doc.sections:
                # 检查首页页眉
                if section.first_page_header:
                    header_text = '\n'.join([p.text for p in section.first_page_header.paragraphs])
                    if re.search(r'\d+|PAGE', header_text, re.IGNORECASE):
                        page_number_info['has_page_number'] = True
                        page_number_info['header_page_number'] = True
                
                # 检查普通页眉
                if section.header:
                    header_text = '\n'.join([p.text for p in section.header.paragraphs])
                    if re.search(r'\d+|PAGE', header_text, re.IGNORECASE):
                        page_number_info['has_page_number'] = True
                        page_number_info['header_page_number'] = True
                
                # 检查页脚
                if section.footer:
                    footer_text = '\n'.join([p.text for p in section.footer.paragraphs])
                    if re.search(r'\d+|PAGE', footer_text, re.IGNORECASE):
                        page_number_info['has_page_number'] = True
                        page_number_info['footer_page_number'] = True
        except Exception as e:
            page_number_info['issues'].append(f"页码检测出错：{str(e)}")
        
        if not page_number_info['has_page_number']:
            page_number_info['issues'].append("未检测到页码")
        
        return page_number_info
    
    def highlight_paragraph(self, paragraph):
        """
        在段落中添加高亮背景
        """
        # 为段落中的每个run添加高亮背景
        for run in paragraph.runs:
            # 设置背景色为黄色
            run.font.highlight_color = 7  # 7 表示黄色
    
    def generate_report(self, docx_path, report_path=None, highlight_path=None):
        """生成正文格式检测报告（分标题和正文）"""
        print("开始正文格式检测...\n")

        # 检查摘要缩进（使用独立变量名避免与highlight中的from import冲突）
        _abstract_doc = Document(docx_path)
        abstract_result = self.check_abstract_indentation(_abstract_doc)
        extra_spaces_result = self.check_extra_spaces(_abstract_doc)

        # 检查正文格式
        results = self.check_body_text(docx_path)
        
        # 如果提供了高亮路径，对有问题的段落进行高亮
        if highlight_path:
            try:
                hl_doc = Document(highlight_path)
                
                # 处理摘要问题
                for key in ("cn_abstract", "en_abstract"):
                    for iss in abstract_result.get(key, {}).get("issues", []):
                        pi = iss['paragraph_index'] - 1
                        if 0 <= pi < len(hl_doc.paragraphs):
                            self.highlight_paragraph(hl_doc.paragraphs[pi])

                # 处理英文多余空格问题
                for iss in extra_spaces_result.get("issues", []):
                    pi = iss['paragraph_index'] - 1
                    if 0 <= pi < len(hl_doc.paragraphs):
                        self.highlight_paragraph(hl_doc.paragraphs[pi])

                # 处理标题问题
                for sample in results['title_samples']:
                    paragraph_index = sample['paragraph_index']
                    if 0 <= paragraph_index < len(hl_doc.paragraphs):
                        self.highlight_paragraph(hl_doc.paragraphs[paragraph_index])

                # 处理正文问题
                for issue in results['all_body_issues']:
                    paragraph_index = issue['paragraph_index'] - 1  # 转换为0-based索引
                    if 0 <= paragraph_index < len(hl_doc.paragraphs):
                        self.highlight_paragraph(hl_doc.paragraphs[paragraph_index])
                
                # 保存修改
                hl_doc.save(highlight_path)
                print(f"正文问题已高亮至: {highlight_path}")
            except Exception as e:
                print(f"高亮正文问题失败: {e}")
        
        # 生成报告
        report_lines = ["# 正文格式检测报告\n"]
        report_lines.append("## 基本信息")
        report_lines.append(f"- 文档路径: {docx_path}")
        report_lines.append(f"- 总段落数: {results['total_paragraphs']}")
        report_lines.append(f"- 一级标题: {results['title_count']['level1']}个")
        report_lines.append(f"- 二级标题: {results['title_count']['level2']}个")
        report_lines.append(f"- 三级标题: {results['title_count']['level3']}个")
        report_lines.append(f"- 正文段落: {results['body_count']}个")
        report_lines.append("")
        
        # 格式标准
        report_lines.append("## 格式标准\n")
        report_lines.append("### 标题格式")
        report_lines.append("- **一级标题**: 黑体三号（16pt）、加粗、居中、1.5倍行距")
        report_lines.append("- **二级标题**: 黑体四号（14pt）、左对齐、1.5倍行距")
        report_lines.append("- **三级标题**: 黑体小四号（12pt）、左对齐、1.5倍行距\n")
        report_lines.append("### 正文格式")
        report_lines.append("- **中文字体**: 宋体")
        report_lines.append("- **英文字体**: Times New Roman")
        report_lines.append("- **字号**: 小四号（12pt）")
        report_lines.append("- **行距**: 1.5倍")
        report_lines.append("- **首行缩进**: 2字符（24pt）")
        report_lines.append("- **对齐方式**: 两端对齐")
        report_lines.append("")
        
        # 摘要格式检测结果
        report_lines.append("## 摘要格式检测")
        self._append_abstract_report(report_lines, abstract_result, "cn_abstract", "中文摘要")
        self._append_abstract_report(report_lines, abstract_result, "en_abstract", "英文摘要")
        report_lines.append("")

        # 英文多余空格检测
        report_lines.append("## 英文标点及字母两侧空格检测")
        if extra_spaces_result["total_issues"] == 0:
            report_lines.append("[通过] **未发现英文标点或字母两侧的多余空格**")
        else:
            report_lines.append(f"[警告] **发现 {extra_spaces_result['total_issues']} 处多余空格问题**")
            for iss in extra_spaces_result["issues"]:
                report_lines.append(f"- 段落{iss['paragraph_index']}: {iss['type']} | {iss['context']}")
        report_lines.append("")

        # 标题检测结果
        report_lines.append("## 标题检测结果")
        if len(results['title_issues']) == 0:
            report_lines.append("[通过] **所有标题格式符合要求**")
        else:
            report_lines.append(f"[警告] **发现 {len(results['title_issues'])} 个标题格式问题**\n")
            report_lines.append("### 标题问题样本")
            for i, sample in enumerate(results['title_samples'], 1):
                level_name = {'level1': '一级标题', 'level2': '二级标题', 'level3': '三级标题'}[sample['level']]
                report_lines.append(f"\n#### 样本 {i} - {level_name}")
                report_lines.append(f"- **内容**: {sample['text_preview']}")
                report_lines.append(f"- **问题**:")
                for issue in sample['issues']:
                    report_lines.append(f"  - {issue}")
        report_lines.append("")
        
        # 正文检测结果
        report_lines.append("## 正文检测结果")
        if len(results['body_issues']) == 0:
            report_lines.append("[通过] **所有正文格式符合要求**")
        else:
            report_lines.append(f"[警告] **发现 {len(results['body_issues'])} 个正文格式问题**\n")
            report_lines.append("### 正文问题位置")
            for i, issue in enumerate(results['all_body_issues'], 1):
                if issue['chapter_position']:
                    report_lines.append(f"\n#### 问题 {i} - {issue['chapter_position']}（第{issue['paragraph_index']}段）")
                else:
                    report_lines.append(f"\n#### 问题 {i} - 第{issue['paragraph_index']}段")
                report_lines.append(f"- **内容**: {issue['text_preview']}")
                report_lines.append(f"- **问题**:")
                for problem in issue['issues']:
                    report_lines.append(f"  - {problem}")
        report_lines.append("")
        
        # 总结
        report_lines.append("## 总结")
        total_issues = len(results['title_issues']) + len(results['body_issues'])
        if total_issues == 0:
            report_lines.append("[通过] **格式检测通过，所有内容格式规范**")
        else:
            report_lines.append(f"[警告] **共发现 {total_issues} 个格式问题**")
            report_lines.append(f"- 标题问题: {len(results['title_issues'])}个")
            report_lines.append(f"- 正文问题: {len(results['body_issues'])}个")
            report_lines.append("\n**建议**: 重点检查字体、字号、对齐方式、首行缩进设置")
        
        report_text = '\n'.join(report_lines)
        
        # 输出到控制台
        print(report_text)
        
        # 保存报告
        if report_path:
            with open(report_path, 'w', encoding='utf-8') as f:
                f.write(report_text)
            print(f"\n报告已保存至: {report_path}")
        
        return report_text


def main():
    import os
    current_dir = os.path.dirname(os.path.abspath(__file__))
    base_dir = os.path.dirname(current_dir)
    docx_path = os.path.join(base_dir, "测试论文.docx")
    reports_dir = os.path.join(base_dir, "reports")
    os.makedirs(reports_dir, exist_ok=True)
    report_path = os.path.join(reports_dir, "body_text_report.md")
    
    checker = BodyTextChecker()
    checker.generate_report(docx_path, report_path)


if __name__ == "__main__":
    main()
