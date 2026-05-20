from checkers.chapter_title_checker import ChapterTitleChecker
import os

# 打开万嘉玮的论文
current_dir = os.path.dirname(os.path.abspath(__file__))
sample_dir = os.path.join(current_dir, "样本")
docx_path = os.path.join(sample_dir, "毕业论文-万嘉玮.docx")

print(f"\n=== 生成章节标题报告 ===")
print(f"文件: {docx_path}")

# 创建检查器实例
checker = ChapterTitleChecker()

# 运行检测并生成报告
result = checker.check_chapter_titles(docx_path)
report_text = checker.generate_report(docx_path)

# 显示报告内容
print(f"\n=== 章节标题报告 ===")
print(report_text)

print("\n=== 生成完成 ===")