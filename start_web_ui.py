"""
启动本地网页版论文检测的简单入口脚本

功能：
- 启动 Flask 后端（集成在此文件中）
- 在浏览器中自动打开网页界面

用法：
- 直接在资源管理器中双击本文件，或在终端运行：
  python start_web_ui.py
"""

import os
import re
import sys
import threading
import time
import webbrowser
import uuid
from datetime import datetime
import shutil
from concurrent.futures import ThreadPoolExecutor

import markdown
from flask import Flask, request, render_template_string, send_file

# 导入检测模块
try:
    from config_loader import load_config, get_cfg
    from run_all_checks import ThesisChecker, BASE_DIR
except ImportError as e:
    print(f"[ERROR] 导入模块失败：{e}")
    print("请确保所有依赖文件都在正确的位置")
    sys.exit(1)

# 参考文献检测器
try:
    from checkers.reference_checker import GraduateReferenceChecker
except ImportError:
    GraduateReferenceChecker = None

# 批注功能
COMMENT_ENABLED = False
try:
    from checkers.comment_utils import run_all_checks_with_comments
    COMMENT_ENABLED = True
    print("[OK] 批注功能已加载")
except ImportError as e:
    print(f"[WARN] 批注功能未加载: {e}")
    run_all_checks_with_comments = None

# 页眉页脚检测器
try:
    from checkers.header_footer_checker import HeaderFooterChecker
except ImportError:
    HeaderFooterChecker = None


def get_base_dir():
    """获取基础目录，兼容 PyInstaller 打包"""
    if getattr(sys, 'frozen', False):
        # PyInstaller 打包后的情况
        return os.path.dirname(sys.executable)
    else:
        # 正常 Python 脚本运行
        return os.path.dirname(os.path.abspath(__file__))


def get_resource_path(relative_path):
    """获取资源文件路径，兼容 PyInstaller 打包"""
    if getattr(sys, 'frozen', False):
        # PyInstaller 打包后的情况，数据文件在临时目录
        if hasattr(sys, '_MEIPASS'):
            return os.path.join(sys._MEIPASS, relative_path)
    # 正常情况
    return os.path.join(get_base_dir(), relative_path)


# 初始化 Flask 应用
app = Flask(__name__)

# 添加basename过滤器
@app.template_filter('basename')
def basename_filter(path):
    return os.path.basename(path)

BASE_DIR_RUNTIME = get_base_dir()
UPLOAD_DIR = os.path.join(BASE_DIR_RUNTIME, "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# 配置文件路径：优先从临时目录（打包后），否则从工作目录
DEFAULT_CONFIG_PATH = get_resource_path("config.json")
if not os.path.exists(DEFAULT_CONFIG_PATH):
    DEFAULT_CONFIG_PATH = os.path.join(BASE_DIR_RUNTIME, "config.json")


def load_base_config():
    """加载基础配置（用于学校名称等通用规则）"""
    if os.path.exists(DEFAULT_CONFIG_PATH):
        return load_config(DEFAULT_CONFIG_PATH)
    return {}


INDEX_TEMPLATE = r"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>论文格式检测系统</title>
<style>
  :root {
    --primary: #1a56db;
    --primary-light: #e8effc;
    --primary-deep: #1e3a8a;
    --bg: #f5f7fa;
    --card: #ffffff;
    --text: #1f2937;
    --text-light: #6b7280;
    --border: #e5e7eb;
    --green: #059669;
    --red: #dc2626;
    --orange: #d97706;
    --shadow: 0 1px 3px rgba(0,0,0,.1), 0 1px 2px rgba(0,0,0,.06);
  }
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
    background: var(--bg); color: var(--text); line-height: 1.6;
  }

  /* ---- 顶栏 ---- */
  .header {
    background: linear-gradient(135deg, #1e3a8a 0%, #1a56db 100%);
    color: #fff; padding: 28px 0; text-align: center;
  }
  .header h1 { font-size: 24px; font-weight: 600; }
  .header p  { font-size: 14px; opacity: .8; margin-top: 6px; }

  .container { max-width: 960px; margin: 0 auto; padding: 24px 16px; }

  /* ---- 上传区 ---- */
  .upload-card {
    background: var(--card); border-radius: 12px; box-shadow: var(--shadow);
    padding: 32px; text-align: center; margin-bottom: 24px;
  }
  .drop-zone {
    border: 2px dashed var(--border); border-radius: 10px;
    padding: 48px 24px; cursor: pointer; transition: .2s;
  }
  .drop-zone:hover, .drop-zone.drag-over {
    border-color: var(--primary); background: var(--primary-light);
  }
  .drop-zone .icon { font-size: 48px; margin-bottom: 12px; }
  .drop-zone h3 { font-size: 18px; margin-bottom: 6px; }
  .drop-zone p  { color: var(--text-light); font-size: 14px; }
  .file-input { display: none; }

  .file-info {
    display: none; margin-top: 16px; padding: 12px 16px;
    background: var(--primary-light); border-radius: 8px;
    font-size: 14px; text-align: left;
  }
  .file-info .name { font-weight: 600; color: var(--primary); }
  .file-info .size { color: var(--text-light); margin-left: 8px; }

  .btn-row { margin-top: 20px; display: flex; gap: 12px; justify-content: center; flex-wrap: wrap; }
  .btn {
    display: inline-flex; align-items: center; gap: 6px;
    padding: 10px 28px; border-radius: 8px; font-size: 15px;
    border: none; cursor: pointer; transition: .2s; font-weight: 500;
    text-decoration: none;
  }
  .btn-primary { background: var(--primary); color: #fff; }
  .btn-primary:hover { background: #1648b8; }
  .btn-primary:disabled { opacity: .5; cursor: not-allowed; }
  .btn-outline { background: #fff; color: var(--primary); border: 1.5px solid var(--primary); }
  .btn-outline:hover { background: var(--primary-light); }
  .btn-sm { padding: 6px 16px; font-size: 13px; }

  /* ---- 进度 ---- */
  .progress-bar { display: none; margin-top: 16px; }
  .progress-bar .track {
    height: 6px; background: var(--border); border-radius: 3px; overflow: hidden;
  }
  .progress-bar .fill {
    height: 100%; width: 0; background: var(--primary); border-radius: 3px;
    transition: width .3s;
  }
  .progress-bar .label { font-size: 13px; color: var(--text-light); margin-top: 4px; }

  /* ---- 报告 ---- */
  .report-card {
    background: var(--card); border-radius: 12px; box-shadow: var(--shadow);
    padding: 32px; display: none; margin-bottom: 24px;
  }
  .report-card h2 { font-size: 20px; margin-bottom: 16px; color: var(--primary); }
  .report-card .report-header {
    display: flex; align-items: center; justify-content: space-between;
    margin-bottom: 16px; flex-wrap: wrap; gap: 8px;
  }
  .report-card .report-header h2 { margin-bottom: 0; }
  .report-card .report-header .actions { display: flex; gap: 8px; flex-wrap: wrap; }
  .report-body {
    font-size: 14px; line-height: 1.8; word-break: break-word;
  }
  .report-body h1 { font-size: 20px; margin: 24px 0 12px; padding-bottom: 8px; border-bottom: 2px solid var(--border); }
  .report-body h2 { font-size: 17px; margin: 20px 0 10px; }
  .report-body h3 { font-size: 15px; margin: 16px 0 8px; }
  .report-body h4 { font-size: 14px; margin: 12px 0 6px; }
  .report-body ul, .report-body ol { padding-left: 24px; margin: 6px 0; }
  .report-body li { margin: 3px 0; }
  .report-body hr { border: none; border-top: 1px solid var(--border); margin: 20px 0; }
  .report-body code {
    background: #f3f4f6; padding: 2px 6px; border-radius: 4px; font-size: 13px;
  }
  .report-body strong { font-weight: 600; }
  .report-body .pass { color: var(--green); }
  .report-body .warn { color: var(--orange); }
  .report-body .fail { color: var(--red); }

  /* ---- 详情报告表格 ---- */
  .report-body table {
    width: 100%; border-collapse: collapse; margin: 12px 0; font-size: 13px;
  }
  .report-body table th,
  .report-body table td {
    padding: 8px 12px; border: 1px solid var(--border); text-align: left;
  }
  .report-body table th {
    background: #f9fafb; font-weight: 600;
  }

  /* ---- 摘要卡片 ---- */
  .summary-cards {
    display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 12px; margin-bottom: 20px;
  }
  .summary-item {
    padding: 16px; border-radius: 10px; text-align: center;
    cursor: pointer; transition: transform .18s, box-shadow .18s;
  }
  .summary-item:hover {
    transform: translateY(-3px);
    box-shadow: 0 4px 12px rgba(0,0,0,.12);
  }
  .summary-item.pass  { background: #ecfdf5; }
  .summary-item.pass:hover  { background: #d1fae5; }
  .summary-item.warn  { background: #fffbeb; }
  .summary-item.warn:hover  { background: #fef3c7; }
  .summary-item.none  { background: #f3f4f6; }
  .summary-item.none:hover  { background: #e5e7eb; }
  .summary-item .label { font-size: 13px; color: var(--text-light); }
  .summary-item .value { font-size: 28px; font-weight: 700; margin-top: 4px; }
  .summary-item .hint { font-size: 11px; color: var(--text-light); opacity: 0; transition: opacity .2s; margin-top: 4px; }
  .summary-item:hover .hint { opacity: 1; }
  .summary-item.pass .value { color: var(--green); }
  .summary-item.warn .value { color: var(--orange); }
  .summary-item.none .value { color: var(--text-light); }

  .highlight-card {
    background: linear-gradient(135deg, #ecfdf5 0%, #d1fae5 100%);
    border: 1.5px solid #6ee7b7; border-radius: 12px; box-shadow: var(--shadow);
    padding: 20px 24px; display: none; text-align: center; margin-bottom: 24px;
  }
  .highlight-card p { color: #065f46; font-size: 15px; font-weight: 500; margin-bottom: 14px; }

  /* ---- 错误 ---- */
  .error-card {
    background: #fef2f2; border: 1.5px solid #fca5a5; border-radius: 12px;
    padding: 16px 24px; text-align: center; margin-bottom: 24px;
  }
  .error-card p { color: var(--red); font-size: 15px; font-weight: 500; }

  /* ---- Footer ---- */
  .footer { text-align: center; padding: 24px; color: var(--text-light); font-size: 13px; }

  @media (max-width: 600px) {
    .container { padding: 12px 8px; }
    .upload-card, .report-card { padding: 16px; }
  }
</style>
</head>
<body>

<div class="header">
  <h1>论文格式检测系统</h1>
  <p>自动检测目录 / 正文 / 参考文献等格式 · 支持高亮版 & 批注版导出</p>
</div>

<div class="container">
  <!-- 错误提示 -->
  {% if error %}
  <div class="error-card">
    <p>{{ error }}</p>
  </div>
  {% endif %}

  <!-- 上传 -->
  <div class="upload-card">
    <form method="post" enctype="multipart/form-data" id="uploadForm">
      <div class="drop-zone" id="dropZone">
        <div class="icon">[+]</div>
        <h3>拖拽 Word 文档到此处</h3>
        <p>或点击选择文件 · 支持 .docx 格式</p>
        <input type="file" class="file-input" id="fileInput" name="file" accept=".docx" required>
      </div>
      <div class="file-info" id="fileInfo">
        <span class="name" id="fileName"></span>
        <span class="size" id="fileSize"></span>
      </div>
      <div class="btn-row">
        <button class="btn btn-primary" type="submit" id="btnCheck" disabled>开始检测</button>
        <button class="btn btn-outline" type="button" id="btnClear" style="display:none">重新选择</button>
      </div>
    </form>
    <div class="progress-bar" id="progressBar">
      <div class="track"><div class="fill" id="progressFill"></div></div>
      <div class="label" id="progressLabel">准备中...</div>
    </div>
  </div>

  {% if summary %}
  <!-- 高亮下载 -->
  {% if highlight_filename %}
  <div class="highlight-card" id="highlightCard" style="display:block;">
    <p>已生成高亮标注文档，问题段落以黄色背景标记</p>
    <a class="btn btn-primary" href="/download/{{ highlight_filename }}">下载高亮版文档</a>
  </div>
  {% endif %}

  <!-- 批注版下载 -->
  {% if comment_filename %}
  <div class="highlight-card" id="commentCard" style="display:block; background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%); border-color: #93c5fd;">
    <p style="color:#1e40af;">已生成批注版文档，错误信息以 Word 批注形式标记在对应位置</p>
    <a class="btn btn-primary" href="/download/{{ comment_filename }}" style="background:#2563eb;">下载批注版文档</a>
  </div>
  {% endif %}

  <!-- 摘要 -->
  <div class="report-card" id="summaryCard" style="display:block;">
    <div class="report-header">
      <h2>检测摘要</h2>
      <div class="actions">
        <button type="button" class="btn btn-outline btn-sm" onclick="copySummary()">复制结果</button>
      </div>
    </div>
    <p style="color:var(--text-light);font-size:13px;margin-bottom:16px;">
      检测时间：{{ time_str }} · {{ filename }}
    </p>
    {{ summary_cards_html|safe }}
  </div>

  <!-- 详细报告 -->
  <div class="report-card" id="reportCard" style="display:block;">
    <h2>详细报告</h2>
    <div class="report-body" id="reportBody">
      {% if report_html %}
        {{ report_html|safe }}
      {% elif summary_html %}
        {{ summary_html|safe }}
      {% endif %}
    </div>
  </div>
  {% endif %}
</div>

<div class="footer">论文格式检测系统 · 本地运行 · 保护论文隐私</div>

<script>
/* ---- 拖拽 & 选择 ---- */
const dropZone  = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const fileInfo  = document.getElementById('fileInfo');
const fileName  = document.getElementById('fileName');
const fileSize  = document.getElementById('fileSize');
const btnCheck  = document.getElementById('btnCheck');
const btnClear  = document.getElementById('btnClear');
const progressBar   = document.getElementById('progressBar');
const progressFill  = document.getElementById('progressFill');
const progressLabel = document.getElementById('progressLabel');

dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault(); dropZone.classList.remove('drag-over');
  if (e.dataTransfer.files.length) { fileInput.files = e.dataTransfer.files; selectFile(e.dataTransfer.files[0]); }
});
fileInput.addEventListener('change', () => { if (fileInput.files.length) selectFile(fileInput.files[0]); });

function selectFile(f) {
  if (!f.name.endsWith('.docx')) { alert('请选择 .docx 格式的 Word 文档'); return; }
  fileName.textContent = f.name;
  fileSize.textContent = `(${(f.size / 1024 / 1024).toFixed(2)} MB)`;
  fileInfo.style.display = 'block';
  btnCheck.disabled = false;
  btnClear.style.display = 'inline-flex';
}
btnClear.addEventListener('click', () => {
  fileInput.value = '';
  fileInfo.style.display = 'none';
  btnCheck.disabled = true;
  btnClear.style.display = 'none';
});

// 表单提交时显示进度
document.getElementById('uploadForm').addEventListener('submit', function() {
  if (!fileInput.files.length) return;
  btnCheck.disabled = true;
  btnCheck.textContent = '检测中...';
  progressBar.style.display = 'block';
  progressFill.style.width = '30%';
  progressLabel.textContent = '正在上传并分析文档...';
  // 模拟进度动画
  let w = 30;
  const iv = setInterval(() => {
    w = Math.min(w + Math.random() * 5, 90);
    progressFill.style.width = w + '%';
    if (w > 60) progressLabel.textContent = '正在生成报告...';
  }, 500);
});

/* ---- 复制结果 ---- */
function copySummary() {
  const raw = document.getElementById('summaryRaw');
  const text = raw ? raw.value : document.body.innerText;
  navigator.clipboard.writeText(text).then(() => {
    const btn = document.querySelector('button[onclick="copySummary()"]');
    const original = btn.textContent;
    btn.textContent = '已复制!';
    setTimeout(() => { btn.textContent = original; }, 2000);
  }).catch(() => alert('复制失败，请手动选择文本'));
}

/* ---- 摘要卡片点击跳转 ---- */
document.querySelectorAll('.summary-item[data-target]').forEach(card => {
  card.addEventListener('click', function() {
    const targetId = this.getAttribute('data-target');
    if (!targetId) return;
    const target = document.getElementById(targetId);
    if (target) {
      target.scrollIntoView({ behavior: 'smooth', block: 'start' });
      // 短暂高亮目标区域
      target.style.transition = 'background .3s';
      target.style.background = 'rgba(26,86,219,.08)';
      target.style.borderRadius = '8px';
      target.style.padding = '4px 8px';
      target.style.margin = '-4px -8px';
      setTimeout(() => {
        target.style.background = '';
        target.style.padding = '';
        target.style.margin = '';
      }, 1500);
    }
  });
});

</script>
</body>
</html>
"""


def _build_summary_cards_html(summary, detailed_reports=None):
    """从汇总报告和详细报告中提取各类检测状态，生成摘要卡片 HTML"""
    if not summary:
        return ""

    def _count_issues(text):
        """统计文本中的问题数量（从报告文本中提取实际数字）"""
        if not text:
            return 0
        # 从 [警告] 行提取 "发现 X 个"（覆盖结构检测等非标准表述）
        nums = re.findall(r'\[警告\].*?发现\s*(\d+)\s*个', text)
        if nums:
            return max(int(n) for n in nums)
        # 从"发现 X 个...问题"中提取
        nums = re.findall(r'发现\s*(\d+)\s*个.*?问题', text)
        if nums:
            return max(int(n) for n in nums)
        # 从"共发现 X 个...问题"中提取
        nums = re.findall(r'共发现\s*(\d+)\s*个.*?问题', text)
        if nums:
            return max(int(n) for n in nums)
        # 从"共 X 个段落存在格式问题"中提取
        nums = re.findall(r'共\s*(\d+)\s*个段落存在格式问题', text)
        if nums:
            return max(int(n) for n in nums)
        # 从"总问题数量: X"中提取
        nums = re.findall(r'总问题数量[：:]\s*(\d+)', text)
        if nums:
            return max(int(n) for n in nums)
        # 从"格式问题数量: X"中提取
        nums = re.findall(r'格式问题数量[：:]\s*(\d+)', text)
        if nums:
            return max(int(n) for n in nums)
        return 0

    def _extract_status(label, keywords, report_keys):
        """优先从详细报告提取状态，汇总报告作为兜底"""
        # 主策略：从详细报告中统计问题
        if detailed_reports:
            for key in report_keys:
                if key in detailed_reports:
                    content = detailed_reports[key].get('content', '')
                    issues = _count_issues(content)
                    if issues > 0:
                        return ('warn', f'{issues}个问题')
                    else:
                        return ('pass', '通过')
                # 也检查 checker.reports（没有单独报告文件的检测项）
                # 如果 detailed_reports 里有按关键词匹配的条目
                for dk, dv in detailed_reports.items():
                    name = dv.get('name', '')
                    if any(kw in name for kw in keywords):
                        content = dv.get('content', '')
                        issues = _count_issues(content)
                        if issues > 0:
                            return ('warn', f'{issues}个问题')
                        else:
                            return ('pass', '通过')

        # 兜底：从汇总报告文本解析
        lines = summary.split('\n')
        for i, line in enumerate(lines):
            if any(kw in line for kw in keywords):
                # 看这一行及上下3行
                start = max(0, i - 3)
                end = min(len(lines), i + 4)
                context = '\n'.join(lines[start:end])
                issues = _count_issues(context)
                if issues > 0:
                    nums = re.findall(r'\d+', context)
                    count = nums[0] if nums else str(issues)
                    return ('warn', f'{count}个问题')

        # 没找到任何相关检测结果 → 未检测
        return ('none', '未检测')

    checks = [
        ('封面格式', ['封面'], ['cover_page']),
        ('目录格式', ['目录'], ['toc']),
        ('正文格式', ['正文'], ['body_text']),
        ('章节标题', ['章节标题', '标题'], ['chapter_titles']),
        ('参考文献', ['参考文献'], ['ref']),
        ('图表公式', ['图表', '图片', '表格', '公式', '术语'], ['figures_tables']),
        ('页眉页脚', ['页眉', '页脚', '页码'], ['header_footer']),
        ('致谢附录', ['致谢', '附录'], ['ack_appendix']),
        ('论文结构', ['结构'], ['structure']),
        ('承诺书', ['承诺书'], ['commitment']),
    ]

    items = [(label, *_extract_status(label, kws, rkeys)) for label, kws, rkeys in checks]

    # 映射检测项 → 报告区域 anchor id（与 _build_report_html 中一致）
    anchor_map = {
        '封面格式': 'sec-cover_page',
        '目录格式': 'sec-toc',
        '正文格式': 'sec-body_text',
        '章节标题': 'sec-chapter_titles',
        '参考文献': 'sec-ref',
        '图表公式': 'sec-figures_tables',
        '页眉页脚': 'sec-header_footer',
        '致谢附录': 'sec-ack_appendix',
        '论文结构': 'sec-structure',
        '承诺书': 'sec-commitment',
    }

    html_parts = ['<div class="summary-cards">']
    for label, status, value in items:
        if status == 'pass':
            css = 'pass'
            display = value
        elif status == 'warn':
            css = 'warn'
            display = value
        else:
            css = 'none'
            display = value
        anchor = anchor_map.get(label, '')
        data_attr = f' data-target="{anchor}"' if anchor else ''
        html_parts.append(
            f'<div class="summary-item {css}"{data_attr}>'
            f'<div class="label">{label}</div>'
            f'<div class="value">{display}</div>'
            f'<div class="hint">点击查看详情</div>'
            f'</div>'
        )
    html_parts.append('</div>')
    return ''.join(html_parts)


def _build_report_html(detailed_reports):
    """将所有详细报告合并为一份 HTML"""
    if not detailed_reports:
        return ""

    order = ['cover_page', 'commitment', 'toc', 'chapter_titles', 'body_text',
             'ref', 'figures_tables', 'header_footer', 'ack_appendix', 'structure']

    parts = []
    for key in order:
        if key not in detailed_reports:
            continue
        name = detailed_reports[key]['name']
        html_content = detailed_reports[key]['html']
        parts.append(f'<hr><h2 id="sec-{key}">{name}</h2>\n{html_content}')

    return '\n'.join(parts)


@app.route("/favicon.ico")
def favicon():
    """提供网站图标"""
    from flask import send_from_directory
    return send_from_directory(os.path.dirname(__file__), "zjut.ico", mimetype="image/vnd.microsoft.icon")


@app.route("/", methods=["GET", "POST"])
def index():
    error = None
    summary = None
    summary_html = None
    summary_cards_html = None
    report_html = None
    filename = None
    time_str = None
    reports_subdir = None
    highlight_path = None
    highlight_filename = None
    comment_filename = None

    if request.method == "POST":
        file = request.files.get("file")
        if not file or file.filename == "":
            error = "请选择一个 .docx 文件。"
        elif not file.filename.lower().endswith(".docx"):
            error = "目前仅支持 .docx 格式的 Word 文档。"
        else:
            # 保存上传文件
            os.makedirs(UPLOAD_DIR, exist_ok=True)
            uid = datetime.now().strftime("%Y%m%d_%H%M%S_") + uuid.uuid4().hex[:8]
            safe_name = uid + "_" + os.path.basename(file.filename)
            save_path = os.path.join(UPLOAD_DIR, safe_name)
            file.save(save_path)

            # 复制一份文件用于高亮
            highlight_filename = f"{uid}_{os.path.splitext(os.path.basename(file.filename))[0]}_高亮版.docx"
            highlight_path = os.path.join(UPLOAD_DIR, highlight_filename)
            shutil.copyfile(save_path, highlight_path)

            # 为本次检测创建专属 reports 子目录
            reports_subdir = os.path.join(BASE_DIR_RUNTIME, "reports", uid)
            os.makedirs(reports_subdir, exist_ok=True)

            # 加载基础配置并为本次运行构造一个副本
            base_cfg = load_base_config()
            run_cfg = dict(base_cfg)
            run_cfg["docx_path"] = save_path
            run_cfg["reports_dir"] = reports_subdir
            run_cfg["highlight_path"] = highlight_path

            # 批注版文件名（提前计算，与核心检测并行执行）
            comment_filename = None
            comment_path = None
            if COMMENT_ENABLED:
                comment_filename = f"{uid}_{os.path.splitext(os.path.basename(file.filename))[0]}_批注版.docx"
                comment_path = os.path.join(UPLOAD_DIR, comment_filename)

            # 并行执行：核心检测、批注生成、参考文献检测、页眉页脚检测
            checker = ThesisChecker(save_path, reports_dir=reports_subdir, config=run_cfg)

            def _run_comment_gen():
                if comment_path is not None:
                    try:
                        print(f"[INFO] 开始生成批注版: {comment_path}")
                        run_all_checks_with_comments(save_path, comment_path)
                        if os.path.exists(comment_path):
                            print(f"[OK] 批注版生成成功: {comment_filename}")
                            return True
                        else:
                            print(f"[WARN] 批注版文件未生成")
                    except Exception as e:
                        print(f"[ERROR] 批注版生成失败: {e}")
                        import traceback
                        traceback.print_exc()
                    return False
                return None

            def _run_ref_checker():
                if GraduateReferenceChecker is None:
                    return ""
                import io as _io
                from contextlib import redirect_stdout as _redirect
                try:
                    ref_buf = _io.StringIO()
                    with _redirect(ref_buf):
                        GraduateReferenceChecker().generate_report(str(save_path))
                    return ref_buf.getvalue()
                except Exception as e:
                    return f"参考文献检测出错: {e}"

            def _run_hf_checker():
                if HeaderFooterChecker is not None:
                    try:
                        hf_report_path = os.path.join(reports_subdir, "header_footer_report.md")
                        HeaderFooterChecker().generate_report(save_path, hf_report_path)
                    except Exception as e:
                        print(f"页眉页脚检测出错: {e}")

            try:
                with ThreadPoolExecutor(max_workers=4) as executor:
                    comment_future = executor.submit(_run_comment_gen)
                    core_future = executor.submit(checker.run_all_checks)
                    ref_future = executor.submit(_run_ref_checker)
                    hf_future = executor.submit(_run_hf_checker)

                    core_future.result()
                    comment_ok = comment_future.result()
                    if comment_ok is False:
                        comment_filename = None
                    ref_report_text = ref_future.result()
                    hf_future.result()


                summary_path = os.path.join(reports_subdir, "检测汇总报告.md")
                
                # 读取汇总报告
                if os.path.exists(summary_path):
                    with open(summary_path, "r", encoding="utf-8") as f:
                        summary_raw = f.read()
                    summary = summary_raw
                else:
                    summary = None
                    error = "检测已完成，但未找到汇总报告。"
                
                # 读取所有详细报告文件，组合成分块展示
                detailed_reports = {}
                report_files = {
                    'cover_page': ('封面格式检测', None),
                    'commitment': ('承诺书检测', None),
                    'toc': ('目录格式检测', 'toc_report.md'),
                    'chapter_titles': ('章节标题格式检测', 'chapter_title_report.md'),
                    'body_text': ('正文格式检测', 'body_text_report.md'),
                    'ref': ('参考文献格式检测', None),  # 独立检测结果注入
                    'figures_tables': ('图表公式术语检测', '综合检测报告.md'),
                    'header_footer': ('页眉页脚格式检测', 'header_footer_report.md'),
                    'ack_appendix': ('致谢和附录格式检测', 'ack_appendix_report.md'),
                    'structure': ('论文整体结构检测', 'structure_report.md'),
                }
                
                # 读取各个报告文件
                for key, (name, fname) in report_files.items():
                    if fname and os.path.exists(os.path.join(reports_subdir, fname)):
                        try:
                            with open(os.path.join(reports_subdir, fname), "r", encoding="utf-8") as f:
                                content = f.read()
                                detailed_reports[key] = {
                                    'name': name,
                                    'content': content,
                                    'html': markdown.markdown(content, extensions=["extra"])
                                }
                        except Exception as e:
                            detailed_reports[key] = {
                                'name': name,
                                'content': f"读取报告文件失败: {e}",
                                'html': f"<p>读取报告文件失败: {e}</p>"
                            }
                    elif key in checker.reports:
                        result = checker.reports[key].get('result', {})
                        if result:
                            content_lines = [f"# {name}\n"]
                            if isinstance(result, dict):
                                for k, v in result.items():
                                    # 校外指导教师为空时跳过（不是所有人都有校外导师）
                                    if k == "校外指导教师" and isinstance(v, dict) and not v.get('内容'):
                                        continue
                                    if isinstance(v, dict) and '内容' in v and '样式' in v:
                                        content_status = "正常" if v['内容'] else "缺失"
                                        style_status = "正常" if v['样式'] else "异常" if v['样式'] is not None else "（未检测）"
                                        content_lines.append(f"- **{k}**:")
                                        content_lines.append(f"  - 内容: {content_status}")
                                        content_lines.append(f"  - 样式: {style_status}")
                                        if v['原因']:
                                            content_lines.append(f"  - 原因: {v['原因']}")
                                    elif isinstance(v, tuple) and len(v) == 3:
                                        has_commitment, has_jpeg, has_date = v
                                        content_lines.append(f"- **{k}**:")
                                        content_lines.append(f"  - 承诺书内容: {'正常' if has_commitment else '缺失'}")
                                        content_lines.append(f"  - JPEG图片（签名）: {'正常' if has_jpeg else '未检测到'}")
                                        content_lines.append(f"  - 日期填写: {'正常' if has_date else '未检测到'}")
                                    else:
                                        content_lines.append(f"- **{k}**: {v}")
                            elif isinstance(result, tuple) and len(result) == 3:
                                has_commitment, has_jpeg, has_date = result
                                content_lines.append("- **承诺书检测**:")
                                content_lines.append(f"  - 承诺书内容: {'正常' if has_commitment else '缺失'}")
                                content_lines.append(f"  - JPEG图片（签名）: {'正常' if has_jpeg else '未检测到'}")
                                content_lines.append(f"  - 日期填写: {'正常' if has_date else '未检测到'}")
                            else:
                                content_lines.append(str(result))
                            content = "\n".join(content_lines)
                            detailed_reports[key] = {
                                'name': name,
                                'content': content,
                                'html': markdown.markdown(content, extensions=["extra"])
                            }
                
                # 注入参考文献检测结果
                if ref_report_text:
                    detailed_reports['ref'] = {
                        'name': '参考文献格式检测',
                        'content': ref_report_text,
                        'html': markdown.markdown(ref_report_text, extensions=["extra"])
                    }

                # 构建摘要卡片
                summary_cards_html = _build_summary_cards_html(summary, detailed_reports)

                # 构建报告 HTML（合并为一份）
                report_html = _build_report_html(detailed_reports)

                # 兜底：如果没有详细报告，用汇总报告 markdown 渲染
                if not report_html and summary:
                    try:
                        summary_html = markdown.markdown(summary_raw, extensions=["extra"])
                    except Exception:
                        esc = (
                            summary_raw.replace("&", "&amp;")
                            .replace("<", "&lt;")
                            .replace(">", "&gt;")
                        )
                        summary_html = f"<pre>{esc}</pre>"
                elif not report_html and not summary:
                    summary_html = "<p>暂无检测结果</p>"
                    
            except Exception as e:
                error = f"检测过程中出错：{e}"
                summary = None
                summary_html = None
                summary_cards_html = None
                report_html = None

            filename = file.filename
            time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            print(f"高亮版文档路径: {highlight_path}")
            print(f"批注版文件名: {comment_filename}")
            print(f"COMMENT_ENABLED: {COMMENT_ENABLED}")

    return render_template_string(
        INDEX_TEMPLATE,
        error=error,
        summary=summary,
        summary_html=summary_html,
        summary_cards_html=summary_cards_html,
        report_html=report_html,
        filename=filename,
        time_str=time_str,
        reports_subdir=reports_subdir,
        highlight_path=highlight_path,
        highlight_filename=highlight_filename,
        comment_filename=comment_filename
    )


@app.route("/download/<path:filename>")
def download_file(filename):
    """下载文件"""
    file_path = os.path.join(UPLOAD_DIR, filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        return "文件不存在", 404

def run_flask():
    """在后台线程中运行 Flask 应用"""
    app.run(host="127.0.0.1", port=5000, debug=False, use_reloader=False)


def main():
    """主函数"""
    print("=" * 50)
    print("论文格式检测系统 - 网页版")
    print("=" * 50)
    print()
    
    # 在后台线程启动 Flask
    flask_thread = threading.Thread(target=run_flask, daemon=True)
    flask_thread.start()
    
    # 等待 Flask 启动
    print("正在启动服务器...")
    time.sleep(2)
    
    # 打开浏览器
    url = "http://127.0.0.1:5000/"
    print(f"[OK] 服务器已启动：{url}")
    print("正在打开浏览器...")
    webbrowser.open(url)
    
    print()
    print("提示：")
    print("  - 在浏览器中上传 .docx 文件进行检测")
    print("  - 检测报告会保存在 exe 所在目录的 reports 文件夹中")
    print("  - 按 Ctrl+C 可停止服务器")
    print()
    
    try:
        # 保持主线程运行
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n\n正在关闭服务器...")
        print("再见！")


if __name__ == "__main__":
    main()
