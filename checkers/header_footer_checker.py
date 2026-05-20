"""
页眉页脚格式检测模块

功能：
- 检查页眉字体（宋体）和字号（9pt / 10.5pt）
- 检查页脚字体（Times New Roman 或 宋体）和字号（9pt / 10.5pt）
- 检查页脚是否包含页码
- 检查各节的页眉页脚设置
- 生成 Markdown 检测报告
"""

from docx import Document
from docx.oxml.ns import qn
import zipfile
from lxml import etree
import re
import os


class HeaderFooterChecker:
    """页眉页脚格式检测器"""

    EXPECTED_HEADER_FONT = '宋体'
    EXPECTED_FOOTER_FONTS = ['Times New Roman', '宋体']
    EXPECTED_SIZES = [9.0, 10.5]

    def get_default_fonts(self, docx_path):
        """获取文档默认字体"""
        default_fonts = {}
        try:
            with zipfile.ZipFile(docx_path) as docx_zip:
                with docx_zip.open('word/styles.xml') as styles_file:
                    styles_xml = etree.parse(styles_file)
                    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
                    rpr_default = styles_xml.find('.//w:docDefaults/w:rPrDefault/w:rPr', ns)
                    if rpr_default is not None:
                        rFonts = rpr_default.find('w:rFonts', ns)
                        if rFonts is not None:
                            default_fonts['eastAsia'] = rFonts.get(qn('w:eastAsia'))
                            default_fonts['ascii'] = rFonts.get(qn('w:ascii'))
        except Exception:
            pass
        return default_fonts

    def get_font_name(self, run, default_fonts, para=None):
        """获取运行块的字体名称"""
        has_chinese = any('\u4e00' <= char <= '\u9fff' for char in run.text)

        try:
            rPr = run._element.rPr
            if rPr is not None and rPr.rFonts is not None:
                rFonts = rPr.rFonts
                font_east = rFonts.get(qn('w:eastAsia'))
                font_ascii = rFonts.get(qn('w:ascii'))
                font_hAnsi = rFonts.get(qn('w:hAnsi'))

                if has_chinese:
                    if font_east:
                        return font_east
                    elif font_hAnsi:
                        return font_hAnsi
                    elif font_ascii:
                        return font_ascii
                else:
                    if font_ascii or font_hAnsi or font_east:
                        return font_ascii or font_hAnsi or font_east

            if run.font.name:
                return run.font.name
            if para and para.style and para.style.font and para.style.font.name:
                return para.style.font.name
        except Exception:
            pass

        if has_chinese:
            return default_fonts.get('eastAsia') or '宋体'
        return default_fonts.get('ascii') or "Times New Roman"

    def get_font_size(self, run, para):
        """获取运行块的字号"""
        if run.font.size:
            return run.font.size.pt
        if run.style and run.style.font and run.style.font.size:
            return run.style.font.size.pt
        if para.style and para.style.font and para.style.font.size:
            return para.style.font.size.pt
        return None

    def get_header_footer_info(self, header_footer, default_fonts):
        """获取页眉或页脚的信息"""
        if not header_footer:
            return None

        info = {
            'text': '',
            'paragraphs': [],
            'fonts': [],
            'has_content': False,
            'has_page_field': False
        }

        for para in header_footer.paragraphs:
            para_text = para.text.strip()

            # 检查是否包含页码字段
            try:
                para_xml = para._element.xml
                if 'PAGE' in para_xml or 'w:fldChar' in para_xml:
                    info['has_page_field'] = True
                    info['has_content'] = True
            except Exception:
                pass

            if para_text:
                info['has_content'] = True
                info['text'] += para_text + ' '
                para_info = {
                    'text': para_text,
                    'alignment': para.alignment,
                    'fonts': []
                }

                for run in para.runs:
                    if run.text.strip():
                        font_name = self.get_font_name(run, default_fonts, para)
                        font_size = self.get_font_size(run, para)
                        para_info['fonts'].append((run.text, font_name, font_size))
                        info['fonts'].append((run.text, font_name, font_size))
                    else:
                        try:
                            if run._element.xml and ('PAGE' in run._element.xml or 'w:fldChar' in run._element.xml):
                                info['has_page_field'] = True
                                info['has_content'] = True
                                font_name = self.get_font_name(run, default_fonts, para)
                                font_size = self.get_font_size(run, para)
                                if font_name or font_size:
                                    para_info['fonts'].append(('[页码字段]', font_name, font_size))
                                    info['fonts'].append(('[页码字段]', font_name, font_size))
                        except Exception:
                            pass

                info['paragraphs'].append(para_info)

        info['text'] = info['text'].strip()

        if info['has_page_field'] and not info['text']:
            info['text'] = '[页码字段]'

        return info

    def _is_linked_to_previous(self, header_footer):
        """检查页眉/页脚是否链接到前一节（即没有独立内容）"""
        try:
            return header_footer.is_linked_to_previous
        except Exception:
            return False

    def check(self, docx_path):
        """
        执行页眉页脚检测，返回结构化结果。

        Returns:
            dict: {
                'sections': [每个节的检测结果],
                'issues': [问题列表],
                'summary': str
            }
        """
        default_fonts = self.get_default_fonts(docx_path)
        doc = Document(docx_path)

        sections = []
        issues = []
        total_issues = 0

        for section_idx, section in enumerate(doc.sections):
            sec_data = {
                'index': section_idx + 1,
                'items': [],
                'has_even_odd': False,
            }

            # 定义要检查的区域
            areas = [
                ('首页页眉', section.first_page_header, True),
                ('普通页眉', section.header, True),
                ('偶数页页眉', section.even_page_header, True),
                ('首页页脚', section.first_page_footer, False),
                ('普通页脚', section.footer, False),
                ('偶数页页脚', section.even_page_footer, False),
            ]

            for label, hf_obj, is_header in areas:
                info = self.get_header_footer_info(hf_obj, default_fonts)
                if info and info['has_content']:
                    item = self._check_single_hf(label, info, is_header, section_idx + 1)
                    sec_data['items'].append(item)
                    if item['errors']:
                        total_issues += len(item['errors'])
                        for err in item['errors']:
                            issues.append(f"节{section_idx + 1} - {label}：{err}")

            # 检查奇偶页设置
            if section.even_page_header or section.even_page_footer:
                sec_data['has_even_odd'] = True

            # 检查普通页眉/页脚是否缺失
            # 只有当该节的普通页眉/页脚既没有独立内容、也没有链接到前一节时才报问题
            has_normal_header = any(it['label'] == '普通页眉' for it in sec_data['items'])
            has_normal_footer = any(it['label'] == '普通页脚' for it in sec_data['items'])

            if not has_normal_header:
                # 检查是否链接到前一节（有继承内容也算合格）
                if not self._is_linked_to_previous(section.header):
                    # 首页不同模式下，如果首页页眉有内容，普通页眉为空可以接受
                    has_first_page = any(it['label'] == '首页页眉' for it in sec_data['items'])
                    if not has_first_page:
                        # 既没有普通页眉也没有首页页眉，才算问题
                        issues.append(f"节{section_idx + 1}：未检测到普通页眉内容")
                        total_issues += 1

            if not has_normal_footer:
                if not self._is_linked_to_previous(section.footer):
                    has_first_page = any(it['label'] == '首页页脚' for it in sec_data['items'])
                    if not has_first_page:
                        issues.append(f"节{section_idx + 1}：未检测到普通页脚内容")
                        total_issues += 1

            sections.append(sec_data)

        # 如果所有节都没有任何页眉页脚内容
        all_empty = all(not sec['items'] for sec in sections)
        if all_empty:
            issues.append("文档中未检测到任何页眉页脚内容")
            total_issues += 1

        summary = f"共检测 {len(doc.sections)} 个节，发现 {total_issues} 个问题" if total_issues else f"共检测 {len(doc.sections)} 个节，页眉页脚格式均符合要求"

        return {
            'sections': sections,
            'issues': issues,
            'summary': summary,
            'total_issues': total_issues,
        }

    def _check_single_hf(self, label, info, is_header, section_num):
        """检查单个页眉/页脚的格式"""
        item = {
            'label': label,
            'text': info['text'][:80],
            'fonts': info['fonts'],
            'has_page_field': info.get('has_page_field', False),
            'errors': [],
            'font_ok': True,
            'size_ok': True,
        }

        if not info['fonts']:
            return item

        for txt, font, size in info['fonts']:
            if not txt.strip():
                continue

            # 字体检查
            if is_header:
                if font and font != self.EXPECTED_HEADER_FONT:
                    item['errors'].append(f"字体「{txt.strip()[:20]}」为 {font}，应为 {self.EXPECTED_HEADER_FONT}")
                    item['font_ok'] = False
            else:
                if font and font not in self.EXPECTED_FOOTER_FONTS:
                    item['errors'].append(f"字体「{txt.strip()[:20]}」为 {font}，应为 {' 或 '.join(self.EXPECTED_FOOTER_FONTS)}")
                    item['font_ok'] = False

            # 字号检查
            if size is not None and size not in self.EXPECTED_SIZES:
                item['errors'].append(f"字号「{txt.strip()[:20]}」为 {size}pt，应为 {' 或 '.join(str(s) + 'pt' for s in self.EXPECTED_SIZES)}")
                item['size_ok'] = False

        # 页脚检查页码
        if '页脚' in label:
            has_page_number = bool(re.search(r'\d+', info['text']))
            has_page_field = info.get('has_page_field', False) or 'PAGE' in info['text'].upper() or bool(re.search(r'第\s*\d+\s*页', info['text']))
            if not has_page_number and not has_page_field:
                item['errors'].append("未检测到页码")
                item['has_page_field'] = False
            else:
                item['has_page_field'] = True

        return item

    def generate_report(self, docx_path, report_path=None):
        """
        生成页眉页脚检测报告。

        Args:
            docx_path: Word 文档路径
            report_path: 报告输出路径（可选）

        Returns:
            str: Markdown 格式的报告文本
        """
        result = self.check(docx_path)

        lines = []
        lines.append("# 页眉页脚格式检测报告")
        lines.append("")
        lines.append("## 检测摘要")
        lines.append("")
        lines.append(f"- 文档：`{os.path.basename(docx_path)}`")
        lines.append(f"- 节数：{len(result['sections'])}")
        lines.append(f"- 问题数：{result['total_issues']}")
        lines.append("")
        lines.append(f"**结论：{result['summary']}**")
        lines.append("")

        # 各节详情
        for sec in result['sections']:
            lines.append(f"## 第 {sec['index']} 节")
            lines.append("")

            if not sec['items']:
                lines.append("[警告] 该节未检测到任何页眉页脚内容")
                lines.append("")
                continue

            for item in sec['items']:
                status = "[通过] 合格" if not item['errors'] else "[错误] 有问题"
                lines.append(f"### {item['label']}：{status}")
                lines.append(f"- 内容：{item['text']}")
                if item['fonts']:
                    font_details = []
                    for txt, font, size in item['fonts']:
                        if txt.strip():
                            font_details.append(f"  `{txt.strip()[:15]}` → 字体: {font}, 字号: {size}pt" if size else f"  `{txt.strip()[:15]}` → 字体: {font}")
                    if font_details:
                        lines.append("- 字体详情：")
                        lines.extend(font_details)
                if '页脚' in item['label']:
                    lines.append(f"- 页码：{'[OK] 已检测到' if item['has_page_field'] else '[错误] 未检测到'}")

                if item['errors']:
                    lines.append("- 问题：")
                    for err in item['errors']:
                        lines.append(f"  - [错误] {err}")
                lines.append("")

            if sec['has_even_odd']:
                lines.append(f"> 💡 该节已设置奇偶页不同")
                lines.append("")

        # 问题汇总
        if result['issues']:
            lines.append("## 问题汇总")
            lines.append("")
            for i, issue in enumerate(result['issues'], 1):
                lines.append(f"{i}. {issue}")
            lines.append("")

        report_text = "\n".join(lines)

        if report_path:
            with open(report_path, 'w', encoding='utf-8') as f:
                f.write(report_text)
            print(f"页眉页脚检测报告已保存至: {report_path}")

        return report_text


def check_header_footer_format(docx_path, report_path=None):
    """
    兼容旧接口：执行页眉页脚检测并打印/写入报告。

    Args:
        docx_path: Word 文档路径
        report_path: 报告输出路径（可选）
    """
    import sys
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')

    checker = HeaderFooterChecker()
    report = checker.generate_report(docx_path, report_path)
    print(report)
    return report


def main():
    """独立运行入口"""
    import sys
    current_dir = os.path.dirname(os.path.abspath(__file__))
    docx_path = os.path.join(current_dir, "测试论文.docx")

    if not os.path.exists(docx_path):
        print(f"[错误] 文件不存在：{docx_path}")
        print(f"   当前目录：{current_dir}")
        sys.exit(1)

    reports_dir = os.path.join(current_dir, "reports")
    os.makedirs(reports_dir, exist_ok=True)
    report_path = os.path.join(reports_dir, "header_footer_report.md")

    print(f"正在检测文档：{docx_path}\n")
    try:
        check_header_footer_format(docx_path, report_path)
    except Exception as e:
        print(f"[错误] 检测过程中出错：{str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
