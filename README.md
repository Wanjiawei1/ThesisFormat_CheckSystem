# 论文格式检测系统

基于 Python 的本科/研究生学位论文格式自动检测系统，支持 10 大检测模块，提供 Web 界面和批注文档生成功能。

## 功能模块

| 模块 | 说明 |
|------|------|
| 封面格式检测 | 检查封面字段完整性、字体、字号、对齐方式 |
| 承诺书检测 | 检查诚信承诺书内容、签名图片、日期填写 |
| 目录格式检测 | 检查目录层级格式、自动生成 vs 手工目录 |
| 章节标题检测 | 检查标题编号连续性、字体、字号、对齐 |
| 正文格式检测 | 检查正文字体（中文宋体/英文Times New Roman）、行距、首行缩进 |
| 图表公式术语检测 | 检查图题/表题/公式编号连续性与格式、术语缩略词 |
| 页眉页脚检测 | 检查页眉内容、页脚页码格式 |
| 致谢与附录检测 | 检查致谢/附录标题和正文格式 |
| 参考文献检测 | 严格按 GB/T 7714-2015 检查条目格式、作者、标点、引用匹配 |
| 论文整体结构检测 | 检查章节顺序、页码连续性、本章小结完整性、章节分页 |

## 技术栈

- **Python 3.10+**
- **Flask** — Web 服务框架
- **python-docx** — Word 文档解析与高亮标注
- **Markdown** — 检测报告生成
- **lxml** — XML 底层操作（批注、分节符检测）

## 快速开始

### 安装

```bash
pip install -r requirements.txt
```

### 命令行使用

```bash
# 检测指定文档（需先在 config.json 中配置或命令行传入）
python run_all_checks.py 论文.docx

# 指定配置文件
python run_all_checks.py --config config.json 论文.docx
```

检测报告将生成在 `reports/` 目录下。

### Web 界面

```bash
python start_web_ui.py
```

浏览器访问 `http://localhost:5000`，上传 .docx 文件即可在线检测并下载高亮/批注版文档。

## 项目结构

```
├── run_all_checks.py          # CLI 入口，调度所有检测模块
├── api_server.py              # API 服务端（检测逻辑核心）
├── start_web_ui.py            # Web UI 入口
├── config.json                # 配置文件（学校名称、关键字段等）
├── config_loader.py           # 配置加载器
├── requirements.txt           # Python 依赖
├── checkers/                  # 检测模块
│   ├── CoverPage.py           # 封面格式检测
│   ├── commitment.py          # 承诺书检测
│   ├── toc_checker.py         # 目录格式检测
│   ├── chapter_title_checker.py  # 章节标题检测
│   ├── body_text_checker.py   # 正文格式检测
│   ├── figure_table_formula_term_checker.py  # 图表公式术语检测
│   ├── header_footer_checker.py  # 页眉页脚检测
│   ├── ack_appendix_checker.py   # 致谢与附录检测
│   ├── reference_checker.py   # 参考文献检测
│   ├── structure_checker.py   # 论文整体结构检测
│   ├── comment_utils.py       # 批注生成工具（Word批注/高亮）
│   └── highlight_errors.py    # 高亮标注工具
└── utils/
    └── comment_adder.py       # Word 批注底层 XML 操作
```

## 配置说明

编辑 `config.json` 配置检测参数：

```json
{
  "docx_path": "测试论文.docx",
  "reports_dir": "reports",
  "school_name": "浙江工业大学",
  "cover": {
    "required_fields": ["学校名称", "论文题目", "学院", ...]
  }
}
```

## API 接口

详见 [API.md](API.md)
