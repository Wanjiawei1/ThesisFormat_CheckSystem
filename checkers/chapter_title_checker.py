from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt
import re

class ChapterTitleChecker:
    def __init__(self):
        # 预定义的标准格式配置（基于正式论文格式要求）
        self.title_configs = {
            1: {  # 一级标题（每章标题）
                'font_name': '黑体',
                'font_size': 16,  # 三号字体（16磅）
                'bold': True,
                'center_aligned': True,  # 居中
                'line_spacing': 1.5,  # 1.5倍行距
                'space_before': 0,  # 段前0行
                'space_after': 0,  # 段后0行
                'patterns': [
                    r'^第[一二三四五六七八九十]+章\s+',
                    r'^第\d+章\s+',
                    r'^\d+\s+(?!年\s+\d+月\s+\d+日$)'
                ]
            },
            2: {  # 二级标题（节号与节标题）
                'font_name': '黑体', 
                'font_size': 14,  # 四号字体（14磅）
                'bold': False,  # 不加粗（根据实际论文）
                'center_aligned': False,  # 左起顶格，两端对齐
                'line_spacing': 1.5,  # 1.5倍行距
                'space_before': 1,  # 段前1行
                'space_after': 0.5,  # 段后0.5行
                'patterns': [
                    r'^\d+\.\d+\s',
                ]
            },
            3: {  # 三级标题（小节号与小节标题）
                'font_name': '黑体',
                'font_size': 12,  # 小四号字体（12磅）
                'bold': False,  # 不加粗（根据实际论文）
                'center_aligned': False,  # 左起顶格，两端对齐
                'line_spacing': 1.5,  # 1.5倍行距
                'space_before': 1,  # 段前1行
                'space_after': 0.5,  # 段后0.5行
                'patterns': [
                    r'^\d+\.\d+\.\d+',
                ]
            }
        }
        
        # 中文数字转换
        self.chinese_to_arabic = {
            '一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
            '六': 6, '七': 7, '八': 8, '九': 9, '十': 10
        }
        
        # 特殊章节（不需要编号连续性检查）
        self.special_chapters = ['摘要', 'ABSTRACT', '目录', '参考文献', '致谢', '附录']

    def has_chinese_char(self, text):
        """
        检测文本中是否包含中文字符
        """
        for char in text:
            if '\u4e00' <= char <= '\u9fff':
                return True
        return False
    
    def infer_chinese_font_from_style(self, paragraph):
        """
        从段落样式推断中文字体
        当XML中没有明确的eastAsia字体时，根据样式名称推断
        """
        try:
            style_name = paragraph.style.name
            
            # 根据样式名称推断（Heading样式通常用黑体）
            if 'Heading' in style_name:
                return '黑体'
            
            # 尝试从样式的font属性获取
            if hasattr(paragraph.style, 'font'):
                style_font = paragraph.style.font
                if style_font and hasattr(style_font, '_element'):
                    try:
                        rpr = style_font._element
                        if rpr is not None:
                            rfonts_elem = rpr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
                            if rfonts_elem is not None:
                                eastasia = rfonts_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia')
                                if eastasia:
                                    return eastasia
                    except:
                        pass
            
            # 默认返回黑体（论文中文标题常用字体）
            return '黑体'
        except:
            return '黑体'  # 默认值
    
    def get_font_info(self, paragraph):
        """
        获取段落的字体信息
        """
        font_info = {
            'name': None,
            'size': None,
            'bold': False,
            'alignment': None
        }
        
        # 获取段落对齐方式
        if paragraph.alignment is not None:
            alignment_map = {0: '左对齐', 1: '居中', 2: '右对齐', 3: '两端对齐'}
            font_info['alignment'] = alignment_map.get(paragraph.alignment, '未知')
        else:
            # alignment为None时，尝试从样式或XML获取
            try:
                # 尝试从段落样式获取
                if hasattr(paragraph.style, 'paragraph_format'):
                    style_alignment = paragraph.style.paragraph_format.alignment
                    if style_alignment is not None:
                        alignment_map = {0: '左对齐', 1: '居中', 2: '右对齐', 3: '两端对齐'}
                        font_info['alignment'] = alignment_map.get(style_alignment, '未知')
                    else:
                        # 尝试从XML获取
                        pPr = paragraph._element.pPr
                        if pPr is not None:
                            jc = pPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc')
                            if jc is not None:
                                jc_val = jc.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                                jc_map = {'left': '左对齐', 'center': '居中', 'right': '右对齐', 'both': '两端对齐'}
                                font_info['alignment'] = jc_map.get(jc_val, '左对齐')
                            else:
                                # 如果样式是Heading，默认居中
                                if 'heading' in paragraph.style.name.lower():
                                    font_info['alignment'] = '居中'
                                else:
                                    font_info['alignment'] = '左对齐'
                        else:
                            font_info['alignment'] = '左对齐'
                else:
                    font_info['alignment'] = '左对齐'
            except:
                # 如果是标题样式，默认居中
                try:
                    if 'heading' in paragraph.style.name.lower():
                        font_info['alignment'] = '居中'
                    else:
                        font_info['alignment'] = '左对齐'
                except:
                    font_info['alignment'] = '左对齐'
        
        # 方法1: 从运行块获取字体信息，优先检测文字内容而非编号
        content_runs = []  # 存储包含文字内容的运行块
        number_runs = []   # 存储包含编号的运行块
        
        for run in paragraph.runs:
            text = run.text.strip()
            if text:
                # 判断这个运行块主要是编号还是文字内容
                # 编号模式：纯数字、小数点、空格组合（如"1.1 "、"2.2.1 "）
                if re.match(r'^[\d\.\s]+$', text) or re.match(r'^第[一二三四五六七八九十\d]+章\s*$', text):
                    number_runs.append(run)
                else:
                    # 包含中文字符或英文单词的，认为是文字内容
                    content_runs.append(run)
        
        # 优先检测文字内容部分的格式
        target_runs = content_runs if content_runs else number_runs
        
        # 先收集所有运行块的字体信息
        all_font_infos = []
        for run in target_runs:
            run_text = run.text.strip()
            if run_text:
                run_font_info = {
                    'name': None,
                    'size': None,
                    'bold': False,
                    'has_chinese': self.has_chinese_char(run_text)
                }
                
                has_chinese = run_font_info['has_chinese']
                
                # 获取字体名称
                if has_chinese:
                    ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
                    # 1) 先检查 Run 自身的 rPr 中 eastAsia
                    try:
                        if hasattr(run._element, 'rPr') and run._element.rPr is not None:
                            rpr = run._element.rPr
                            rfonts_elem = rpr.find(f'{ns}rFonts')
                            if rfonts_elem is not None:
                                eastasia = rfonts_elem.get(f'{ns}eastAsia')
                                if eastasia:
                                    run_font_info['name'] = eastasia
                    except:
                        pass
                    # 2) Run 没有显式设置时，从段落样式的 rPr 获取 eastAsia
                    if not run_font_info['name']:
                        try:
                            style = paragraph.style
                            if style and hasattr(style, 'element'):
                                style_rpr = style.element.find(f'{ns}rPr')
                                if style_rpr is not None:
                                    rfonts_elem = style_rpr.find(f'{ns}rFonts')
                                    if rfonts_elem is not None:
                                        eastasia = rfonts_elem.get(f'{ns}eastAsia')
                                        if eastasia:
                                            run_font_info['name'] = eastasia
                        except:
                            pass
                    # 3) 最后才用推断函数兜底
                    if not run_font_info['name']:
                        run_font_info['name'] = self.infer_chinese_font_from_style(paragraph)
                else:
                    # 非中文文本，优先从XML直接读取显式设置的字体
                    # （run.font.name 会沿继承链返回文档默认主题字体如 Arial，不准确）
                    ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
                    # 1) 先检查 Run 自身的 rPr
                    try:
                        if hasattr(run._element, 'rPr') and run._element.rPr is not None:
                            rpr = run._element.rPr
                            rfonts_elem = rpr.find(f'{ns}rFonts')
                            if rfonts_elem is not None:
                                ascii_font = rfonts_elem.get(f'{ns}ascii')
                                hAnsi = rfonts_elem.get(f'{ns}hAnsi')
                                if ascii_font:
                                    run_font_info['name'] = ascii_font
                                elif hAnsi:
                                    run_font_info['name'] = hAnsi
                    except:
                        pass
                    # 2) Run 没有显式字体时，从段落样式的 rPr 获取
                    if not run_font_info['name']:
                        try:
                            style = paragraph.style
                            if style and hasattr(style, 'element'):
                                style_rpr = style.element.find(f'{ns}rPr')
                                if style_rpr is not None:
                                    rfonts_elem = style_rpr.find(f'{ns}rFonts')
                                    if rfonts_elem is not None:
                                        ascii_font = rfonts_elem.get(f'{ns}ascii')
                                        hAnsi = rfonts_elem.get(f'{ns}hAnsi')
                                        if ascii_font:
                                            run_font_info['name'] = ascii_font
                                        elif hAnsi:
                                            run_font_info['name'] = hAnsi
                        except:
                            pass
                    # 3) 最后才用 run.font.name（可能返回主题默认字体）
                    if not run_font_info['name'] and run.font.name:
                        run_font_info['name'] = run.font.name
                
                # 获取字号
                if run.font.size:
                    run_font_info['size'] = round(run.font.size.pt)  # 转换为磅数
                else:
                    # 从XML元素获取字号
                    try:
                        if hasattr(run._element, 'rPr') and run._element.rPr is not None:
                            sz_elem = run._element.rPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz')
                            if sz_elem is not None:
                                # 字号是半磅单位，需要除以2
                                run_font_info['size'] = int(sz_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')) / 2
                    except:
                        pass
                
                # 获取加粗状态
                if run.font.bold is True:
                    run_font_info['bold'] = True
                else:
                    # 从XML元素获取加粗信息
                    try:
                        if hasattr(run._element, 'rPr') and run._element.rPr is not None:
                            b_elem = run._element.rPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}b')
                            if b_elem is not None:
                                run_font_info['bold'] = True
                    except:
                        pass
                
                all_font_infos.append(run_font_info)
        
        # 优先使用包含中文的运行块的字体信息
        chinese_font_infos = [info for info in all_font_infos if info['has_chinese']]
        if chinese_font_infos:
            # 使用第一个包含中文的运行块的字体信息
            for key in ['name', 'size', 'bold']:
                for info in chinese_font_infos:
                    if info[key] is not None:
                        font_info[key] = info[key]
                        break
        # 如果没有中文运行块，使用其他运行块的字体信息
        elif all_font_infos:
            for key in ['name', 'size', 'bold']:
                for info in all_font_infos:
                    if info[key] is not None:
                        font_info[key] = info[key]
                        break
        
        # 方法2: 如果运行块没有字体信息，尝试从段落样式获取
        if not font_info['name'] or not font_info['size'] or not font_info['bold']:
            try:
                # 获取段落样式的字体信息
                style = paragraph.style
                if hasattr(style, 'font'):
                    if not font_info['name'] and style.font.name:
                        font_info['name'] = style.font.name
                    if not font_info['size'] and style.font.size:
                        # 样式字体大小处理
                        try:
                            if hasattr(style.font.size, 'pt'):
                                font_info['size'] = round(style.font.size.pt)
                            else:
                                # 字体大小可能以不同单位存储，尝试多种转换
                                size_value = int(str(style.font.size))
                                # 通过观察，203200对应16磅左右，177800对应14磅左右
                                # 计算比例：203200/16 ≈ 12700, 177800/14 ≈ 12700
                                # 所以转换公式可能是：size_value / 12700
                                font_info['size'] = round(size_value / 12700)
                                if font_info['size'] <= 0:  # 如果结果太小，尝试其他转换
                                    font_info['size'] = round(size_value / 2)  # 半磅转换
                                if font_info['size'] > 100:  # 如果结果太大，尝试其他转换
                                    font_info['size'] = round(size_value / 20)  # twips转换
                        except Exception as e:
                            pass
                    if not font_info['bold'] and style.font.bold:
                        font_info['bold'] = True
            except Exception as e:
                pass
        
        # 方法3: 从文档的默认字体和设置获取（如果还是没有信息）
        if not font_info['name'] or not font_info['size']:
            try:
                # 如果段落没有明确的字体设置，尝试使用文档默认值
                # 根据标题样式名称推断可能的格式
                style_name = paragraph.style.name.lower() if paragraph.style else ''
                
                # 根据样式名称推断字体和字号
                if 'heading 1' in style_name or paragraph.style.name == 'Heading 1':
                    if not font_info['name']:
                        font_info['name'] = '黑体'  # 一级标题默认黑体
                    if not font_info['size']:
                        font_info['size'] = 16  # 根据观察，实际可能是16磅而不是18磅
                elif 'heading 2' in style_name or paragraph.style.name == 'Heading 2':
                    if not font_info['name']:
                        font_info['name'] = '黑体'  # 二级标题默认黑体  
                    if not font_info['size']:
                        font_info['size'] = 14  # 根据观察，实际可能是14磅而不是16磅
                elif 'heading 3' in style_name or paragraph.style.name == 'Heading 3':
                    if not font_info['name']:
                        font_info['name'] = '黑体'  # 三级标题默认黑体
                    if not font_info['size']:
                        font_info['size'] = 12  # 三级标题可能是12磅
                
                # 如果还是没有字体，使用文档默认字体
                if not font_info['name']:
                    doc_element = paragraph._element.getroottree().getroot()
                    styles_element = doc_element.xpath('//w:styles', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                    if styles_element:
                        default_fonts = styles_element[0].xpath('.//w:rFonts[@w:ascii]', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                        if default_fonts:
                            font_info['name'] = default_fonts[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii')
                        
                # 如果还是没有字号，根据标题级别设置默认值
                if not font_info['size']:
                    text = paragraph.text.strip()
                    if re.match(r'^第[一二三四五六七八九十\d]+章', text) or re.match(r'^\d+\s+', text):
                        font_info['size'] = 16  # 一级标题
                    elif re.match(r'^\d+\.\d+\s+', text):
                        font_info['size'] = 14  # 二级标题  
                    elif re.match(r'^\d+\.\d+\.\d+\s+', text):
                        font_info['size'] = 12  # 三级标题
                        
            except Exception as e:
                pass
        
        return font_info

    def debug_font_info(self, docx_path, start=0, end=10):
        """
        调试字体信息获取
        """
        doc = Document(docx_path)
        print(f"=== 调试字体信息 (段落 {start+1} 到 {end}) ===")
        
        for i, para in enumerate(doc.paragraphs[start:end]):
            actual_index = start + i + 1
            text = para.text.strip()
            if not text:
                continue
                
            print(f"\n段落 {actual_index}: {text[:30]}{'...' if len(text) > 30 else ''}")
            print(f"  样式名称: {para.style.name}")
            
            # 检查段落级别的字体信息
            try:
                if hasattr(para.style, 'font'):
                    print(f"  段落样式字体: {para.style.font.name}, {para.style.font.size}")
            except:
                print("  段落样式字体: 无法获取")
            
            # 检查运行块
            print(f"  运行块数量: {len(para.runs)}")
            for j, run in enumerate(para.runs):
                if run.text.strip():
                    print(f"    运行块 {j+1}: '{run.text}'")
                    print(f"      字体名称: {run.font.name}")
                    print(f"      字体大小: {run.font.size}")
                    print(f"      是否加粗: {run.font.bold}")
                    
                    # 尝试从XML获取
                    try:
                        if hasattr(run._element, 'rPr') and run._element.rPr is not None:
                            rpr = run._element.rPr
                            if hasattr(rpr, 'rFonts') and rpr.rFonts is not None:
                                rfonts = rpr.rFonts
                                print(f"      XML字体 - eastAsia: {getattr(rfonts, 'eastAsia', None)}")
                                print(f"      XML字体 - ascii: {getattr(rfonts, 'ascii', None)}")
                            if hasattr(rpr, 'sz') and rpr.sz is not None:
                                print(f"      XML字号: {rpr.sz.val} (半磅)")
                            if hasattr(rpr, 'b') and rpr.b is not None:
                                print(f"      XML加粗: True")
                    except Exception as e:
                        print(f"      XML信息获取错误: {e}")

    def detect_title_level(self, text):
        """
        检测标题级别
        """
        # 先检查是否为日期格式，排除日期被识别为标题
        date_patterns = [
            r'^\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日$',
            r'^\d{4}\s*年\s*\d{1,2}\s*月$'
        ]
        for date_pattern in date_patterns:
            if re.match(date_pattern, text.strip()):
                return 0  # 是日期，不是标题
        
        # 检查文本长度，如果过长，可能不是标题
        text_len = len(text.strip())
        if text_len > 100:  # 标题通常不会超过100个字符
            return 0
        
        # 检查文本内容，如果包含逗号、句号等标点符号，可能不是标题
        if re.search(r'[，。；；：！？]', text):
            return 0
        
        # 然后检查标题模式
        for level in [1, 2, 3]:
            for pattern in self.title_configs[level]['patterns']:
                if re.match(pattern, text.strip()):
                    # 进一步检查，确保匹配的是真正的标题
                    # 对于一级标题，确保后面有标题内容
                    if level == 1:
                        # 提取标题内容
                        title_content = re.sub(r'^第[一二三四五六七八九十\d]+章\s*', '', text.strip())
                        title_content = re.sub(r'^\d+\s+', '', title_content)
                        if not title_content or len(title_content) < 2:
                            return 0
                    return level
        return 0  # 不是标题

    def extract_title_number(self, text, level):
        """
        提取标题编号
        """
        text = text.strip()
        
        if level == 1:
            # 处理章节编号
            chapter_match = re.match(r'^第([一二三四五六七八九十\d]+)章', text)
            if chapter_match:
                chapter_num = chapter_match.group(1)
                if chapter_num in self.chinese_to_arabic:
                    return [self.chinese_to_arabic[chapter_num]]
                else:
                    return [int(chapter_num)]
            
            # 处理数字编号
            number_match = re.match(r'^(\d+)', text)
            if number_match:
                return [int(number_match.group(1))]
        
        elif level == 2:
            # 处理二级标题编号 (如 2.1)
            match = re.match(r'^(\d+)\.(\d+)', text)
            if match:
                return [int(match.group(1)), int(match.group(2))]
        
        elif level == 3:
            # 处理三级标题编号 (如 2.1.1)
            match = re.match(r'^(\d+)\.(\d+)\.(\d+)', text)
            if match:
                return [int(match.group(1)), int(match.group(2)), int(match.group(3))]
        
        return []

    def extract_title_text(self, text, level):
        """
        提取标题文本内容（去除编号）
        """
        text = text.strip()
        
        if level == 1:
            # 去除章节编号
            text = re.sub(r'^第[一二三四五六七八九十\d]+章\s*', '', text)
            text = re.sub(r'^\d+\s+', '', text)
        elif level == 2:
            text = re.sub(r'^\d+\.\d+\s+', '', text)
        elif level == 3:
            text = re.sub(r'^\d+\.\d+\.\d+\s+', '', text)
        
        return text.strip()

    def is_special_chapter(self, title_text):
        """
        检查是否为特殊章节
        """
        for special in self.special_chapters:
            if special in title_text:
                return True
        return False

    def check_title_format(self, paragraph, expected_level):
        """
        检查标题格式是否符合规范
        """
        issues = []
        font_info = self.get_font_info(paragraph)
        expected_config = self.title_configs[expected_level]
        
        # 检查字体名称
        if font_info['name'] != expected_config['font_name']:
            issues.append(f"字体应为{expected_config['font_name']}，当前为{font_info['name']}")
        
        # 检查字号
        if font_info['size'] != expected_config['font_size']:
            issues.append(f"字号应为{expected_config['font_size']}磅，当前为{font_info['size']}磅")
        
        # 不检查加粗（根据用户要求）
        # if font_info['bold'] != expected_config['bold']:
        #     status = "加粗" if expected_config['bold'] else "不加粗"
        #     issues.append(f"加粗错误: 期望{status}")
        
        # 检查对齐方式
        expected_alignment = "居中" if expected_config['center_aligned'] else "左对齐"
        if expected_config['center_aligned'] and font_info['alignment'] != '居中':
            issues.append(f"对齐方式应为{expected_alignment}，当前为{font_info['alignment']}")
        
        return issues, font_info

    def check_number_continuity(self, titles):
        """
        检查标题编号连续性
        """
        issues = []
        prev_numbers = [0, 0, 0]  # 记录各级别的前一个编号
        
        for title in titles:
            if title['is_special']:
                continue
            
            level = title['level']
            numbers = title['numbers']
            
            if not numbers:
                continue
            
            # 检查编号连续性
            if level == 1:
                expected = prev_numbers[0] + 1
                if numbers[0] != expected:
                    issues.append({
                        'description': f"一级标题编号不连续：\"{title['full_text']}\"，当前第{numbers[0]}章，请检查上一章标题编号",
                        'paragraph_index': title['paragraph_index']
                    })
                prev_numbers[0] = numbers[0]
                prev_numbers[1] = 0
                prev_numbers[2] = 0

            elif level == 2:
                if len(numbers) >= 2:
                    if numbers[0] != prev_numbers[0]:
                        issues.append({
                            'description': f"所属上级标题编号不匹配：\"{title['full_text']}\"，请检查一级标题编号是否连续",
                            'paragraph_index': title['paragraph_index']
                        })

                    expected = prev_numbers[1] + 1
                    if numbers[1] != expected:
                        issues.append({
                            'description': f"二级标题编号不连续：\"{title['full_text']}\"，当前{numbers[0]}.{numbers[1]}，请检查上一节标题编号",
                            'paragraph_index': title['paragraph_index']
                        })

                    prev_numbers[1] = numbers[1]
                    prev_numbers[2] = 0

            elif level == 3:
                if len(numbers) >= 3:
                    if numbers[0] != prev_numbers[0] or numbers[1] != prev_numbers[1]:
                        issues.append({
                            'description': f"上级标题编号异常：\"{title['full_text']}\"，请检查上级标题格式是否正确（如编号后是否缺少空格）",
                            'paragraph_index': title['paragraph_index']
                        })

                    expected = prev_numbers[2] + 1
                    if numbers[2] != expected:
                        issues.append({
                            'description': f"三级标题编号不连续：\"{title['full_text']}\"，当前{numbers[0]}.{numbers[1]}.{numbers[2]}，请检查上一小节标题编号",
                            'paragraph_index': title['paragraph_index']
                        })

                    prev_numbers[2] = numbers[2]
        
        return issues

    def analyze_titles(self, docx_path):
        """
        分析文档中的所有标题
        """
        doc = Document(docx_path)
        titles = []
        format_issues = []
        
        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            if not text:
                continue
            
            # 跳过目录（TOC）内容
            style_name = para.style.name.lower()
            if 'toc' in style_name:
                continue
            
            # 检测标题级别
            level = self.detect_title_level(text)
            if level == 0:
                continue
            
            # 提取标题信息
            numbers = self.extract_title_number(text, level)
            title_text = self.extract_title_text(text, level)
            # 仅当标题没有编号前缀时才视为特殊章节，避免误跳过带编号的标题（如"4.6 目录格式检测模块实现"）
            has_number_prefix = bool(re.match(r'^(\d+\.)*\d+\s', text.strip())) or bool(re.match(r'^第[一二三四五六七八九十\d]+章', text.strip()))
            is_special = (not has_number_prefix) and self.is_special_chapter(title_text)
            
            # 检查格式
            if not is_special:  # 特殊章节不检查格式
                issues, font_info = self.check_title_format(para, level)
                if issues:
                    for issue in issues:
                        format_issues.append({
                            'description': f"\"{text}\": {issue}",
                            'paragraph_index': i
                        })
            else:
                font_info = self.get_font_info(para)
            
            titles.append({
                'paragraph_index': i,
                'level': level,
                'numbers': numbers,
                'title_text': title_text,
                'full_text': text,
                'is_special': is_special,
                'font_info': font_info
            })
        
        # 检查编号连续性
        continuity_issues = self.check_number_continuity(titles)
        
        return titles, format_issues, continuity_issues

    def check_chapter_titles(self, docx_path):
        """
        综合检查章节标题格式
        """
        titles, format_issues, continuity_issues = self.analyze_titles(docx_path)
        
        result = {
            'titles': titles,
            'format_issues': format_issues,
            'continuity_issues': continuity_issues,
            'total_issues': len(format_issues) + len(continuity_issues),
            'summary': ''
        }
        
        # 生成摘要
        if result['total_issues'] == 0:
            result['summary'] = f"[通过] 章节标题格式正确，共检测到{len(titles)}个标题"
        else:
            result['summary'] = f"[警告] 章节标题存在{result['total_issues']}个问题"
        
        return result

    def highlight_paragraph(self, paragraph):
        """
        在段落中添加高亮背景
        """
        # 为段落中的每个run添加高亮背景
        for run in paragraph.runs:
            # 设置背景色为黄色
            run.font.highlight_color = 7  # 7 表示黄色
    
    def generate_report(self, docx_path, report_path=None, highlight_path=None):
        """
        生成章节标题检测报告
        """
        result = self.check_chapter_titles(docx_path)
        
        # 如果提供了高亮路径，对有问题的段落进行高亮
        if highlight_path:
            try:
                from docx import Document
                doc = Document(highlight_path)

                # 处理格式问题（现在每个issue是dict，直接有paragraph_index）
                for issue in result['format_issues']:
                    pi = issue['paragraph_index']
                    if 0 <= pi < len(doc.paragraphs):
                        self.highlight_paragraph(doc.paragraphs[pi])

                # 处理编号问题
                for issue in result['continuity_issues']:
                    pi = issue['paragraph_index']
                    if 0 <= pi < len(doc.paragraphs):
                        self.highlight_paragraph(doc.paragraphs[pi])

                # 保存修改
                doc.save(highlight_path)
                print(f"章节标题问题已高亮至: {highlight_path}")
            except Exception as e:
                print(f"高亮章节标题问题失败: {e}")
        
        report = ["# 章节标题格式检测报告", ""]
        report.append("## 基本信息")
        report.append(f"- 文档路径: {docx_path}")
        report.append(f"- 检测到的标题数量: {len(result['titles'])}")
        report.append(f"- 格式问题数量: {len(result['format_issues'])}")
        report.append(f"- 编号问题数量: {len(result['continuity_issues'])}")
        report.append(f"- 总问题数量: {result['total_issues']}")
        report.append("")
        
        # 格式问题
        if result['format_issues']:
            report.append("## 格式问题")
            for i, issue in enumerate(result['format_issues'], 1):
                report.append(f"{i}. {issue['description']}")
            report.append("")

        # 编号连续性问题
        if result['continuity_issues']:
            report.append("## 编号连续性问题")
            for i, issue in enumerate(result['continuity_issues'], 1):
                report.append(f"{i}. {issue['description']}")
            report.append("")
        
        # 标题列表
        if result['titles']:
            report.append("## 检测到的标题列表")
            for title in result['titles']:
                level_indent = "  " * (title['level'] - 1)
                font_info = title['font_info']
                font_desc = f"({font_info['name']}, {font_info['size']}磅)"
                
                if title['is_special']:
                    report.append(f"{level_indent}- [特殊] {title['full_text']} {font_desc}")
                else:
                    numbers_str = '.'.join(map(str, title['numbers'])) if title['numbers'] else '无编号'
                    report.append(f"{level_indent}- [{numbers_str}] {title['full_text']} {font_desc}")
            report.append("")
        
        # 标准格式要求
        report.append("## 标准格式要求")
        for level, config in self.title_configs.items():
            level_name = ['', '一级标题(章)', '二级标题(节)', '三级标题(小节)'][level]
            alignment = "居中" if config['center_aligned'] else "左对齐"
            bold_status = "加粗" if config['bold'] else "不加粗"
            report.append(f"- {level_name}: {config['font_name']}, {config['font_size']}磅, {bold_status}, {alignment}")
        report.append("")
        
        # 结论
        report.append("## 结论")
        report.append(result['summary'])
        
        report_text = "\n".join(report)
        
        if report_path:
            with open(report_path, 'w', encoding='utf-8') as f:
                f.write(report_text)
            print(f"章节标题检测报告已保存至: {report_path}")
        
        return report_text


def main():
    import os
    current_dir = os.path.dirname(os.path.abspath(__file__))
    base_dir = os.path.dirname(current_dir)
    # 使用样本文件夹中的文件进行测试
    sample_dir = os.path.join(base_dir, "样本")
    docx_path = os.path.join(sample_dir, "毕业论文-万嘉玮.docx")
    reports_dir = os.path.join(base_dir, "reports")
    os.makedirs(reports_dir, exist_ok=True)
    report_path = os.path.join(reports_dir, "chapter_title_report.md")
    
    checker = ChapterTitleChecker()
    
    # 生成报告
    print("开始章节标题格式检测...")
    report = checker.generate_report(docx_path, report_path)
    print(report)


if __name__ == "__main__":
    main()
