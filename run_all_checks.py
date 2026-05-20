"""
论文格式检测主程序
统一调用所有检测模块，生成综合报告
"""

import os
import sys
import io

# 确保 Windows 控制台能正确处理 UTF-8 输出
if sys.platform == 'win32':
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
    except (AttributeError, OSError):
        pass

from datetime import datetime
from config_loader import load_config, get_cfg

# 处理 PyInstaller 打包后的路径
def get_base_dir():
    """获取基础目录，兼容 PyInstaller 打包"""
    if getattr(sys, 'frozen', False):
        # PyInstaller 打包后的情况
        # 数据文件在临时目录，但工作目录应该在 exe 所在目录
        return os.path.dirname(sys.executable)
    else:
        # 正常 Python 脚本运行
        return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = get_base_dir()
CHECKERS_DIR = os.path.join(BASE_DIR, "checkers")

# 如果是打包后的情况，需要从临时目录加载数据文件
if getattr(sys, 'frozen', False):
    # PyInstaller 会将数据文件提取到 sys._MEIPASS
    if hasattr(sys, '_MEIPASS'):
        meipass = sys._MEIPASS
        # 将 checkers 目录添加到路径
        checkers_in_temp = os.path.join(meipass, "checkers")
        if os.path.exists(checkers_in_temp):
            sys.path.insert(0, meipass)
            CHECKERS_DIR = checkers_in_temp

# 确保能导入 checkers 包
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

try:
    from checkers.CoverPage import check_cover_page_final, check_english_names
    from checkers.commitment import check_commitment_info_by_zip
    from checkers.toc_checker import TocChecker
    from checkers.chapter_title_checker import ChapterTitleChecker
    from checkers.body_text_checker import BodyTextChecker
    from checkers.figure_table_formula_term_checker import FigureTableFormulaTermChecker
    from checkers.header_footer_checker import check_header_footer_format
    from checkers.ack_appendix_checker import AckAppendixChecker
    from checkers.structure_checker import ThesisStructureChecker
except ImportError as e:
    print(f"导入检测模块失败: {e}")
    print(f"提示：请确认检测脚本在 `{CHECKERS_DIR}` 且包含 `__init__.py`")
    sys.exit(1)


class ThesisChecker:
    """论文格式检测主类"""
    
    def __init__(self, docx_path, reports_dir, config=None):
        self.docx_path = docx_path
        self.reports_dir = reports_dir
        self.config = config or {}
        self.current_dir = BASE_DIR
        self.reports = {}
        self.errors = {}
        self.highlight_path = self.config.get('highlight_path')
        
    def check_cover_page(self):
        """检测封面格式"""
        print("\n" + "="*60)
        print("【1/9】封面格式检测")
        print("="*60)
        try:
            cover_cfg = get_cfg(self.config, "cover", {}) or {}
            school_name = cover_cfg.get("school_name") or get_cfg(self.config, "school_name", "浙江工业大学")
            result = check_cover_page_final(self.docx_path, school_name=school_name)
            english_name_result = check_english_names(self.docx_path, school_name=school_name)
            result['english_names'] = english_name_result
            self.reports['cover_page'] = {
                'name': '封面格式检测',
                'status': 'completed',
                'result': result
            }
            print("[OK] 封面检测完成")
            if english_name_result.get('issues'):
                for iss in english_name_result['issues']:
                    print(f"  {iss}")
        except Exception as e:
            self.errors['cover_page'] = str(e)
            print(f"[FAIL] 封面检测出错: {e}")
            
    def check_commitment(self):
        """检测承诺书"""
        print("\n" + "="*60)
        print("【2/9】承诺书检测")
        print("="*60)
        try:
            comm_cfg = get_cfg(self.config, "commitment", {}) or {}
            keywords = comm_cfg.get("keywords")
            school_kw = comm_cfg.get("school_keyword") or get_cfg(self.config, "school_name")
            result = check_commitment_info_by_zip(self.docx_path, keywords=keywords, school_keyword=school_kw)
            self.reports['commitment'] = {
                'name': '承诺书检测',
                'status': 'completed',
                'result': result
            }
            print("[OK] 承诺书检测完成")
        except Exception as e:
            self.errors['commitment'] = str(e)
            print(f"[FAIL] 承诺书检测出错: {e}")
            
    def check_toc(self):
        """检测目录格式"""
        print("\n" + "="*60)
        print("【3/9】目录格式检测")
        print("="*60)
        try:
            checker = TocChecker()
            report_path = os.path.join(self.reports_dir, "toc_report.md")
            checker.generate_report(self.docx_path, report_path)
            self.reports['toc'] = {
                'name': '目录格式检测',
                'status': 'completed',
                'report_path': report_path
            }
            print("[OK] 目录检测完成")
        except Exception as e:
            self.errors['toc'] = str(e)
            print(f"[FAIL] 目录检测出错: {e}")
            
    def check_chapter_titles(self):
        """检测章节标题格式"""
        print("\n" + "="*60)
        print("【4/9】章节标题格式检测")
        print("="*60)
        try:
            checker = ChapterTitleChecker()
            report_path = os.path.join(self.reports_dir, "chapter_title_report.md")
            checker.generate_report(self.docx_path, report_path, self.highlight_path)
            self.reports['chapter_titles'] = {
                'name': '章节标题格式检测',
                'status': 'completed',
                'report_path': report_path
            }
            print("[OK] 章节标题检测完成")
        except Exception as e:
            self.errors['chapter_titles'] = str(e)
            print(f"[FAIL] 章节标题检测出错: {e}")
            
    def check_body_text(self):
        """检测正文格式"""
        print("\n" + "="*60)
        print("【5/9】正文格式检测")
        print("="*60)
        try:
            checker = BodyTextChecker()
            report_path = os.path.join(self.reports_dir, "body_text_report.md")
            checker.generate_report(self.docx_path, report_path, self.highlight_path)
            self.reports['body_text'] = {
                'name': '正文格式检测',
                'status': 'completed',
                'report_path': report_path
            }
            print("[OK] 正文检测完成")
        except Exception as e:
            self.errors['body_text'] = str(e)
            print(f"[FAIL] 正文检测出错: {e}")
            
    def check_figures_tables_formulas(self):
        """检测图表公式术语"""
        print("\n" + "="*60)
        print("【6/9】图表公式术语检测")
        print("="*60)
        try:
            checker = FigureTableFormulaTermChecker()
            report_path = os.path.join(self.reports_dir, "综合检测报告.md")
            checker.generate_report(self.docx_path, report_path, self.highlight_path)
            self.reports['figures_tables'] = {
                'name': '图表公式术语检测',
                'status': 'completed',
                'report_path': report_path
            }
            print("[OK] 图表公式术语检测完成")
        except Exception as e:
            self.errors['figures_tables'] = str(e)
            print(f"[FAIL] 图表公式术语检测出错: {e}")
            
    def check_header_footer(self):
        """检测页眉页脚格式"""
        print("\n" + "="*60)
        print("【7/9】页眉页脚格式检测")
        print("="*60)
        try:
            # header_footer_checker 是函数，不是类
            report_path = os.path.join(self.reports_dir, "header_footer_report.md")
            check_header_footer_format(self.docx_path, report_path=report_path)
            self.reports['header_footer'] = {
                'name': '页眉页脚格式检测',
                'status': 'completed',
                'report_path': report_path
            }
            print("[OK] 页眉页脚检测完成")
        except Exception as e:
            self.errors['header_footer'] = str(e)
            print(f"[FAIL] 页眉页脚检测出错: {e}")
            
    def check_ack_appendix(self):
        """检测致谢和附录格式"""
        print("\n" + "="*60)
        print("【8/9】致谢和附录格式检测")
        print("="*60)
        try:
            checker = AckAppendixChecker()
            report_path = os.path.join(self.reports_dir, "ack_appendix_report.md")
            checker.generate_report(self.docx_path, report_path, self.highlight_path)
            self.reports['ack_appendix'] = {
                'name': '致谢和附录格式检测',
                'status': 'completed',
                'report_path': report_path
            }
            print("[OK] 致谢和附录检测完成")
        except Exception as e:
            self.errors['ack_appendix'] = str(e)
            print(f"[FAIL] 致谢和附录检测出错: {e}")
            
    def check_structure(self):
        """检测论文整体结构"""
        print("\n" + "="*60)
        print("【9/9】论文整体结构检测")
        print("="*60)
        try:
            checker = ThesisStructureChecker()
            report_path = os.path.join(self.reports_dir, "structure_report.md")
            checker.generate_report(self.docx_path, report_path)
            self.reports['structure'] = {
                'name': '论文整体结构检测',
                'status': 'completed',
                'report_path': report_path
            }
            print("[OK] 整体结构检测完成")
        except Exception as e:
            self.errors['structure'] = str(e)
            print(f"[FAIL] 整体结构检测出错: {e}")
            
    def run_all_checks(self):
        """运行所有检测"""
        print("\n" + "="*60)
        print("开始论文格式全面检测")
        print("="*60)
        print(f"文档路径: {self.docx_path}")
        print(f"检测时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # 依次运行所有检测
        self.check_cover_page()
        self.check_commitment()
        self.check_toc()
        self.check_chapter_titles()
        self.check_body_text()
        self.check_figures_tables_formulas()
        self.check_header_footer()
        self.check_ack_appendix()
        self.check_structure()
        
        # 生成汇总报告
        self.generate_summary_report()
        
    def generate_summary_report(self):
        """生成汇总报告"""
        report_path = os.path.join(self.reports_dir, "检测汇总报告.md")
        
        report_lines = []
        report_lines.append("# 论文格式检测汇总报告\n")
        report_lines.append(f"**检测时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        report_lines.append(f"**文档路径**: {self.docx_path}\n")
        report_lines.append("")
        
        report_lines.append("## 检测项目汇总\n")
        
        # 检测项目列表
        check_items = [
            ('cover_page', '封面格式检测'),
            ('commitment', '承诺书检测'),
            ('toc', '目录格式检测'),
            ('chapter_titles', '章节标题格式检测'),
            ('body_text', '正文格式检测'),
            ('figures_tables', '图表公式术语检测'),
            ('header_footer', '页眉页脚格式检测'),
            ('ack_appendix', '致谢和附录格式检测'),
            ('structure', '论文整体结构检测'),
        ]
        
        completed_count = 0
        error_count = 0
        
        for key, name in check_items:
            if key in self.reports:
                status = "[OK] 已完成"
                completed_count += 1
                report_lines.append(f"- **{name}**: {status}")
                if 'report_path' in self.reports[key]:
                    report_lines.append(f"  - 报告文件: `{os.path.basename(self.reports[key]['report_path'])}`")
            elif key in self.errors:
                status = f"[FAIL] 出错: {self.errors[key]}"
                error_count += 1
                report_lines.append(f"- **{name}**: {status}")
            else:
                status = "[跳过] 未执行"
                report_lines.append(f"- **{name}**: {status}")
        
        report_lines.append("")
        report_lines.append("## 检测统计\n")
        report_lines.append(f"- 已完成检测: {completed_count}/9")
        report_lines.append(f"- 检测出错: {error_count}/9")
        report_lines.append("")
        
        report_lines.append("## 详细报告\n")
        report_lines.append("各检测项目的详细报告请查看对应的 `.md` 文件：\n")
        
        report_files = [
            ("toc_report.md", "目录格式检测报告"),
            ("chapter_title_report.md", "章节标题格式检测报告"),
            ("body_text_report.md", "正文格式检测报告"),
            ("综合检测报告.md", "图表公式术语综合检测报告"),
            ("ack_appendix_report.md", "致谢和附录格式检测报告"),
            ("structure_report.md", "论文整体结构检测报告"),
        ]
        
        for filename, description in report_files:
            filepath = os.path.join(self.reports_dir, filename)
            if os.path.exists(filepath):
                report_lines.append(f"- **{description}**: `{filename}`")
        
        report_lines.append("")
        report_lines.append("---")
        report_lines.append("\n*本报告由论文格式检测系统自动生成*")
        
        report_text = '\n'.join(report_lines)
        
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(report_text)
        
        print("\n" + "="*60)
        print("检测完成！")
        print("="*60)
        print(f"汇总报告已保存至: {report_path}")
        print(f"已完成检测: {completed_count}/9")
        if error_count > 0:
            print(f"检测出错: {error_count}/9")
        print("="*60)


def main():
    """主函数"""
    import argparse
    
    parser = argparse.ArgumentParser(description='论文格式全面检测')
    parser.add_argument('--config', default=os.path.join(BASE_DIR, "config.json"), help='配置文件路径（默认 config.json）')
    parser.add_argument('docx_path', nargs='?', default=None, help='待检测的Word文档路径（覆盖配置）')
    args = parser.parse_args()

    # 加载配置
    config_path = args.config
    if not os.path.isabs(config_path):
        config_path = os.path.join(BASE_DIR, config_path)
    if not os.path.exists(config_path):
        print(f"[FAIL] 配置文件不存在: {config_path}")
        print("   你可以复制 `config.example.json` 为 `config.json` 后修改。")
        sys.exit(1)
    cfg = load_config(config_path)

    # docx_path 优先级：命令行 > 配置文件
    docx_path = args.docx_path or get_cfg(cfg, "docx_path") or os.path.join(BASE_DIR, "测试论文.docx")
    if not os.path.isabs(docx_path):
        docx_path = os.path.join(BASE_DIR, docx_path)

    reports_dir_base = get_cfg(cfg, "reports_dir") or os.path.join(BASE_DIR, "reports")
    if not os.path.isabs(reports_dir_base):
        reports_dir_base = os.path.join(BASE_DIR, reports_dir_base)

    # 为每次检测创建独立子目录，避免不同文档的报告混淆
    doc_name = os.path.splitext(os.path.basename(docx_path))[0]
    task_subdir = f"{doc_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    reports_dir = os.path.join(reports_dir_base, task_subdir)
    os.makedirs(reports_dir, exist_ok=True)

    if not os.path.exists(docx_path):
        print(f"[FAIL] 文件不存在: {docx_path}")
        print(f"   配置文件: {config_path}")
        sys.exit(1)
    
    # 运行检测
    checker = ThesisChecker(docx_path, reports_dir=reports_dir, config=cfg)
    try:
        checker.run_all_checks()
    except KeyboardInterrupt:
        print("\n\n检测被用户中断")
        sys.exit(1)
    except Exception as e:
        print(f"\n\n检测过程中发生严重错误: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()

