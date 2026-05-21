"""
论文格式检测系统 - REST API 接口层

接口概览：
  POST   /api/v1/check                          上传论文并启动检测
  GET    /api/v1/status/<task_id>                查询任务状态
  GET    /api/v1/report/<task_id>                获取检测报告（JSON / Markdown 文本）
  GET    /api/v1/report/<task_id>/download       下载汇总报告文件（.md）
  GET    /api/v1/download/<task_id>              下载高亮版文档（.docx）
  GET    /api/v1/download/<task_id>/comments     下载批注版文档（.docx）
  GET    /api/v1/health                          健康检查
"""

import os
import sys
import uuid
import shutil
import threading
import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from functools import wraps, lru_cache
from io import BytesIO

from flask import Flask, request, jsonify, send_file
from flask.json.provider import DefaultJSONProvider
from docx import Document as _RawDocument

# 文档加载缓存，避免同一文件被重复解析
@lru_cache(maxsize=4)
def _cached_document(path: str):
    return _RawDocument(path)

try:
    from config_loader import load_config, get_cfg
    from run_all_checks import ThesisChecker, BASE_DIR
except ImportError as e:
    print(f"[ERROR] 导入核心模块失败：{e}")
    sys.exit(1)

try:
    from checkers.comment_utils import run_all_checks_with_comments
except ImportError:
    run_all_checks_with_comments = None

try:
    from checkers.reference_checker import GraduateReferenceChecker
except ImportError:
    GraduateReferenceChecker = None

try:
    from checkers.header_footer_checker import HeaderFooterChecker
except ImportError:
    HeaderFooterChecker = None

app = Flask(__name__)


class SafeJSONProvider(DefaultJSONProvider):
    def default(self, o):
        if isinstance(o, bytes):
            return f"<{len(o)} bytes>"
        return super().default(o)

app.json_provider_class = SafeJSONProvider
app.json = SafeJSONProvider(app)

API_KEY = None
UPLOAD_DIR = "uploads"
REPORTS_DIR = "reports"
MAX_FILE_SIZE = 50 * 1024 * 1024

_tasks = {}


def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_resource_path(relative_path):
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(get_base_dir(), relative_path)


def load_base_config():
    config_path = get_resource_path("config.json")
    if not os.path.exists(config_path):
        config_path = os.path.join(get_base_dir(), "config.json")
    if os.path.exists(config_path):
        return load_config(config_path)
    return {}


def require_api_key(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if API_KEY:
            auth = request.headers.get("Authorization", "")
            token = ""
            if auth.startswith("Bearer "):
                token = auth[7:]
            elif auth.startswith("Token "):
                token = auth[6:]
            if not token:
                token = request.args.get("api_key", "")
            if token != API_KEY:
                return jsonify({"error": "认证失败", "detail": "无效的 API Key"}), 401
        return f(*args, **kwargs)
    return decorated


def _parse_check_status(content: str, check_key: str) -> tuple:
    import re as _re

    if not content:
        return "skip", []

    issues = []

    # 从 [警告] 行提取 "发现 X 个"（覆盖结构检测等非标准表述）
    nums = _re.findall(r'\[警告\].*?发现\s*(\d+)\s*个', content)
    if not nums:
        nums = _re.findall(r'发现\s*(\d+)\s*个.*?问题', content)
    if not nums:
        nums = _re.findall(r'共发现\s*(\d+)\s*个.*?问题', content)
    if not nums:
        nums = _re.findall(r'总问题数量[：:]\s*(\d+)', content)
    if not nums:
        nums = _re.findall(r'格式问题数量[：:]\s*(\d+)', content)

    if nums:
        total_issues = max(int(n) for n in nums)
        issues = [{"message": f"共发现 {total_issues} 个问题", "severity": "error", "count": total_issues}]
        return "fail", issues

    # 回退：从报告中解析问题条目（以 "- " 或数字序号开头的行）
    for line in content.split("\n"):
        line_stripped = line.strip()
        if _re.match(r'^\d+[\.\)]\s', line_stripped) or _re.match(r'^-\s', line_stripped):
            issues.append({"message": line_stripped, "severity": "warning"})

    if issues:
        return "warn", issues
    return "pass", []


def _build_consolidated_report(original_filename: str, summary: dict, checks: dict, reports_dir: str) -> str:
    from datetime import datetime as _dt
    lines = []

    lines.append("# 论文格式检测报告")
    lines.append("")
    lines.append(f"- **文件名**: {original_filename}")
    lines.append(f"- **检测时间**: {_dt.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"- **检测模块数**: {summary.get('total', 0)}")
    lines.append("")

    pass_n = summary.get("pass", 0)
    warn_n = summary.get("warn", 0)
    fail_n = summary.get("fail", 0)
    skip_n = summary.get("skip", 0)

    lines.append("## 总体统计")
    lines.append("")
    lines.append("| 指标 | 数量 |")
    lines.append("|------|------|")
    lines.append(f"| [通过] | {pass_n} |")
    lines.append(f"| [警告] | {warn_n} |")
    lines.append(f"| [失败] | {fail_n} |")
    lines.append(f"| [跳过] | {skip_n} |")
    lines.append(f"| **合计** | **{summary.get('total', 0)}** |")
    lines.append("")

    if fail_n > 0:
        lines.append(f"> **结论**: [失败] 检测未通过，存在 {fail_n} 项失败，请根据下方详细报告修正后重新检测。")
    elif warn_n > 0:
        lines.append(f"> **结论**: [警告] 检测基本通过，但存在 {warn_n} 项警告，建议修正。")
    else:
        lines.append("> **结论**: [通过] 所有检测项目均通过。")
    lines.append("")

    lines.append("## 检测概览")
    lines.append("")
    lines.append("| 序号 | 检测模块 | 状态 | 问题数 |")
    lines.append("|------|----------|------|--------|")

    status_icon = {"pass": "[通过]", "warn": "[警告]", "fail": "[失败]", "skip": "[跳过]"}

    ordered_keys = [
        "cover_page", "commitment", "toc", "chapter_titles",
        "body_text", "ref", "figures_tables", "header_footer",
        "ack_appendix", "structure",
    ]

    for idx, key in enumerate(ordered_keys, 1):
        if key not in checks:
            continue
        ck = checks[key]
        name = ck.get("name", key)
        st = ck.get("status", "skip")
        issue_count = ck.get("issue_count", 0)
        issue_str = str(issue_count) if issue_count > 0 else "-"
        lines.append(f"| {idx} | {name} | {status_icon.get(st, st)} | {issue_str} |")

    lines.append("")
    lines.append("---")
    lines.append("")

    lines.append("## 详细报告")
    lines.append("")

    for key in ordered_keys:
        if key not in checks:
            continue
        ck = checks[key]
        name = ck.get("name", key)
        st = ck.get("status", "skip")
        detail = ck.get("detail", "")
        issues = ck.get("issues", [])

        icon = {"pass": "[通过]", "warn": "[警告]", "fail": "[失败]", "skip": "[跳过]"}.get(st, "")
        lines.append(f"### {icon} {name}")
        lines.append("")

        if issues:
            actual_count = ck.get("issue_count", len(issues))
            lines.append(f"**发现问题 {actual_count} 个：**")
            lines.append("")
            for iss in issues:
                sev = iss.get("severity", "info")
                sev_label = "[错误]" if sev == "error" else "[警告]" if sev == "warning" else "[信息]"
                lines.append(f"- {sev_label} {iss.get('message', '')}")
            lines.append("")

        if detail:
            detail_lines = detail.strip().split("\n")
            cleaned = []
            skip_header = True
            for dl in detail_lines:
                if skip_header and dl.strip().startswith("#"):
                    continue
                if skip_header and dl.strip() == "":
                    continue
                skip_header = False
                cleaned.append(dl)
            if cleaned:
                lines.append("\n".join(cleaned))
                lines.append("")
        else:
            lines.append("*该模块未生成详细报告。*")
            lines.append("")

        lines.append("---")
        lines.append("")

    lines.append("## 附录：原始报告文件")
    lines.append("")
    # reports_dir 可能已被清理，检查是否存在
    if reports_dir and os.path.isdir(reports_dir):
        for fname in sorted(os.listdir(reports_dir)):
            if fname.endswith(".md") and fname != "consolidated_report.md":
                lines.append(f"- `{fname}`")
    lines.append("")
    lines.append("---")
    lines.append("")
    lines.append("*本报告由论文格式检测系统自动生成*")

    return "\n".join(lines)


def _run_check(file_path: str, task_id: str, uploaded_file_path: str = None) -> dict:
    reports_subdir = os.path.join(get_base_dir(), REPORTS_DIR, task_id)
    os.makedirs(reports_subdir, exist_ok=True)

    highlight_filename = f"{task_id}_highlighted.docx"
    highlight_path = os.path.join(UPLOAD_DIR, highlight_filename)
    shutil.copyfile(file_path, highlight_path)

    base_cfg = load_base_config()
    run_cfg = dict(base_cfg)
    run_cfg["docx_path"] = file_path
    run_cfg["reports_dir"] = reports_subdir
    run_cfg["highlight_path"] = highlight_path

    result = {
        "task_id": task_id,
        "status": "completed",
        "filename": os.path.basename(file_path),
        "summary": {"pass": 0, "warn": 0, "fail": 0, "skip": 0, "total": 0},
        "checks": {},
        "report_markdown": "",
        "highlight_available": False,
    }

    try:
        checker = ThesisChecker(file_path, reports_dir=reports_subdir, config=run_cfg)
        checker.run_all_checks()

        # 并行运行 3 个独立后续任务：参考文献检测、页眉页脚检测、批注生成
        def _run_ref_checker():
            if GraduateReferenceChecker is None:
                return ""
            import io as _io
            from contextlib import redirect_stdout as _redirect
            try:
                ref_buf = _io.StringIO()
                with _redirect(ref_buf):
                    GraduateReferenceChecker().generate_report(str(file_path))
                return ref_buf.getvalue()
            except Exception as e:
                return f"参考文献检测出错: {e}"

        def _run_hf_checker():
            if HeaderFooterChecker is not None:
                try:
                    hf_report_path = os.path.join(reports_subdir, "header_footer_report.md")
                    HeaderFooterChecker().generate_report(file_path, hf_report_path)
                except Exception as e:
                    print(f"页眉页脚检测出错: {e}")

        def _run_comments():
            comment_path = os.path.join(UPLOAD_DIR, f"{task_id}_comments.docx")
            if run_all_checks_with_comments is not None:
                try:
                    run_all_checks_with_comments(file_path, output_path=comment_path)
                    if os.path.exists(comment_path):
                        with open(comment_path, "rb") as f:
                            return f.read()
                except Exception as e:
                    print(f"批注版生成出错: {e}")
            return None

        with ThreadPoolExecutor(max_workers=3) as post_executor:
            ref_future = post_executor.submit(_run_ref_checker)
            hf_future = post_executor.submit(_run_hf_checker)
            comment_future = post_executor.submit(_run_comments)
            ref_report_text = ref_future.result()
            hf_future.result()
            comment_bytes = comment_future.result()

        report_files = {
            "cover_page":    ("封面格式检测",   None),
            "commitment":    ("承诺书检测",     None),
            "toc":           ("目录格式检测",   "toc_report.md"),
            "chapter_titles":("章节标题格式检测","chapter_title_report.md"),
            "body_text":     ("正文格式检测",   "body_text_report.md"),
            "ref":           ("参考文献格式检测", None),
            "figures_tables":("图表公式术语检测","综合检测报告.md"),
            "header_footer": ("页眉页脚格式检测","header_footer_report.md"),
            "ack_appendix":  ("致谢和附录格式检测","ack_appendix_report.md"),
            "structure":     ("论文整体结构检测","structure_report.md"),
        }

        all_markdown_parts = []

        for key, (display_name, fname) in report_files.items():
            content = ""
            if fname:
                fpath = os.path.join(reports_subdir, fname)
                if os.path.exists(fpath):
                    with open(fpath, "r", encoding="utf-8") as f:
                        content = f.read()
            if key == "ref" and ref_report_text:
                content = ref_report_text
            if not content and key in checker.reports:
                result_data = checker.reports[key].get("result", {})
                if result_data:
                    content = str(result_data)

            check_status, issues = _parse_check_status(content, key)
            issue_count = issues[0].get("count", len(issues)) if issues else 0
            result["checks"][key] = {
                "name": display_name,
                "status": check_status,
                "issues": issues,
                "issue_count": issue_count,
                "detail": content,
            }
            result["summary"]["total"] += 1
            result["summary"][check_status] = result["summary"].get(check_status, 0) + 1

            if content:
                all_markdown_parts.append(f"## {display_name}\\n\\n{content}")

        consolidated_md = _build_consolidated_report(
            original_filename=os.path.basename(file_path),
            summary=result["summary"],
            checks=result["checks"],
            reports_dir=reports_subdir,
        )
        result["report_markdown"] = consolidated_md

        if os.path.exists(highlight_path):
            with open(highlight_path, "rb") as f:
                result["highlight_bytes"] = f.read()
            result["highlight_available"] = True

        # 从并行任务结果中设置批注文档
        result["comment_available"] = False
        if comment_bytes is not None:
            result["comment_bytes"] = comment_bytes
            result["comment_available"] = True

    except Exception as e:
        result["status"] = "error"
        result["error"] = str(e)

    # 清理磁盘缓存文件
    comment_file = os.path.join(UPLOAD_DIR, f"{task_id}_comments.docx")
    _cleanup_disk(reports_subdir, highlight_path, uploaded_file_path)
    _cleanup_disk(None, comment_file, None)

    return result


def _cleanup_disk(reports_dir: str, highlight_path: str, uploaded_file_path: str = None):
    try:
        if os.path.isdir(reports_dir):
            shutil.rmtree(reports_dir)
    except Exception:
        pass
    try:
        if highlight_path and os.path.exists(highlight_path):
            os.remove(highlight_path)
    except Exception:
        pass
    try:
        if uploaded_file_path and os.path.exists(uploaded_file_path):
            os.remove(uploaded_file_path)
    except Exception:
        pass


# ---- API 路由 ----

@app.route("/api/v1/health", methods=["GET"])
def health():
    return jsonify({
        "status": "ok",
        "service": "thesis-format-checker",
        "version": "1.0.0",
        "time": datetime.now().isoformat(),
    })


@app.route("/api/v1/check", methods=["POST"])
@require_api_key
def check_thesis():
    task_id = datetime.now().strftime("%Y%m%d_%H%M%S_") + uuid.uuid4().hex[:8]
    save_path = None
    original_filename = None
    is_upload = False
    content_type = request.content_type or ""

    if "application/json" in content_type or request.is_json:
        data = request.get_json(silent=True) or {}
        file_path = data.get("file_path", "").strip()
        mode = data.get("mode", "sync").lower()

        if not file_path:
            return jsonify({"error": "请提供 file_path 参数", "code": "MISSING_PATH"}), 400

        file_path = os.path.abspath(file_path)
        if not os.path.exists(file_path):
            return jsonify({"error": f"文件不存在: {file_path}", "code": "FILE_NOT_FOUND"}), 404
        if not file_path.lower().endswith(".docx"):
            return jsonify({"error": "仅支持 .docx 格式", "code": "INVALID_FORMAT"}), 400
        if os.path.getsize(file_path) > MAX_FILE_SIZE:
            return jsonify({"error": f"文件过大，最大允许 {MAX_FILE_SIZE // 1024 // 1024}MB"}), 400

        save_path = file_path
        original_filename = os.path.basename(file_path)

    else:
        file = request.files.get("file")
        mode = request.form.get("mode", "sync").lower()

        if not file or file.filename == "":
            return jsonify({"error": "请上传 .docx 文件或提供 file_path", "code": "MISSING_FILE"}), 400
        if not file.filename.lower().endswith(".docx"):
            return jsonify({"error": "仅支持 .docx 格式", "code": "INVALID_FORMAT"}), 400

        os.makedirs(UPLOAD_DIR, exist_ok=True)
        safe_name = f"{task_id}_{os.path.basename(file.filename)}"
        save_path = os.path.join(UPLOAD_DIR, safe_name)
        file.save(save_path)
        is_upload = True

        if os.path.getsize(save_path) > MAX_FILE_SIZE:
            os.remove(save_path)
            return jsonify({"error": f"文件过大，最大允许 {MAX_FILE_SIZE // 1024 // 1024}MB"}), 400

        original_filename = file.filename

    if mode == "async":
        _tasks[task_id] = {
            "status": "pending",
            "created_at": datetime.now().isoformat(),
            "filename": original_filename,
            "file_path": save_path,
        }
        uploaded_path = save_path if is_upload else None
        thread = threading.Thread(target=_async_run, args=(task_id, save_path, uploaded_path), daemon=True)
        thread.start()
        return jsonify({
            "task_id": task_id,
            "status": "pending",
            "message": "检测任务已提交，请通过 /api/v1/status/{task_id} 查询进度",
        }), 202
    else:
        uploaded_path = save_path if is_upload else None
        result = _run_check(save_path, task_id, uploaded_file_path=uploaded_path)

        # bytes 不能 JSON 序列化，单独存储
        highlight_bytes = result.pop("highlight_bytes", None)
        comment_bytes = result.pop("comment_bytes", None)

        _tasks[task_id] = {
            "status": result["status"],
            "result": result,
            "created_at": datetime.now().isoformat(),
            "filename": original_filename,
        }
        if highlight_bytes:
            _tasks[task_id]["result"]["highlight_bytes"] = highlight_bytes
        if comment_bytes:
            _tasks[task_id]["result"]["comment_bytes"] = comment_bytes

        return jsonify(result), 200


def _async_run(task_id: str, file_path: str, uploaded_file_path: str = None):
    _tasks[task_id]["status"] = "running"
    try:
        result = _run_check(file_path, task_id, uploaded_file_path=uploaded_file_path)
        _tasks[task_id]["status"] = result["status"]
        _tasks[task_id]["result"] = result
    except Exception as e:
        _tasks[task_id]["status"] = "error"
        _tasks[task_id]["error"] = str(e)


@app.route("/api/v1/status/<task_id>", methods=["GET"])
@require_api_key
def task_status(task_id):
    task = _tasks.get(task_id)
    if not task:
        return jsonify({"error": "任务不存在", "code": "NOT_FOUND"}), 404

    resp = {
        "task_id": task_id,
        "status": task["status"],
        "filename": task.get("filename"),
        "created_at": task.get("created_at"),
    }
    if task["status"] in ("completed", "error") and "result" in task:
        resp["result"] = task["result"]
    if task["status"] == "error" and "error" in task:
        resp["error"] = task["error"]

    return jsonify(resp), 200


@app.route("/api/v1/report/<task_id>", methods=["GET"])
@require_api_key
def get_report(task_id):
    task = _tasks.get(task_id)
    if not task:
        return jsonify({"error": "任务不存在"}), 404
    if task["status"] != "completed":
        return jsonify({"error": f"任务状态为 {task['status']}，尚未完成"}), 409

    fmt = request.args.get("format", "json").lower()
    result = task.get("result", {})

    if fmt == "markdown":
        md = result.get("report_markdown", "暂无报告")
        return md, 200, {"Content-Type": "text/plain; charset=utf-8"}
    else:
        return jsonify(result), 200


@app.route("/api/v1/report/<task_id>/download", methods=["GET"])
@require_api_key
def download_report_md(task_id):
    task = _tasks.get(task_id)
    if not task:
        return jsonify({"error": "任务不存在", "code": "NOT_FOUND"}), 404
    if task["status"] != "completed":
        return jsonify({"error": f"任务状态为 {task['status']}，尚未完成", "code": "NOT_READY"}), 409

    result = task.get("result", {})
    md_content = result.get("report_markdown", "")
    if not md_content:
        return jsonify({"error": "报告内容为空", "code": "EMPTY_REPORT"}), 404

    original_name = task.get("filename", "thesis")
    base_name = os.path.splitext(original_name)[0]
    buf = BytesIO(md_content.encode("utf-8"))
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name=f"{base_name}_检测报告.md",
        mimetype="text/markdown; charset=utf-8",
    )


@app.route("/api/v1/download/<task_id>", methods=["GET"])
@require_api_key
def download_highlight(task_id):
    task = _tasks.get(task_id)
    if not task:
        return jsonify({"error": "任务不存在", "code": "NOT_FOUND"}), 404
    if task["status"] != "completed":
        return jsonify({"error": "任务尚未完成", "code": "NOT_READY"}), 409

    result = task.get("result", {})
    highlight_bytes = result.get("highlight_bytes")
    if not highlight_bytes:
        return jsonify({"error": "高亮文档不可用", "code": "NO_HIGHLIGHT"}), 404

    original_name = task.get("filename", "thesis.docx")
    base_name = os.path.splitext(original_name)[0]
    buf = BytesIO(highlight_bytes)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name=f"{base_name}_高亮版.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.route("/api/v1/download/<task_id>/comments", methods=["GET"])
@require_api_key
def download_comments(task_id):
    task = _tasks.get(task_id)
    if not task:
        return jsonify({"error": "任务不存在", "code": "NOT_FOUND"}), 404
    if task["status"] != "completed":
        return jsonify({"error": "任务尚未完成", "code": "NOT_READY"}), 409

    result = task.get("result", {})
    comment_bytes = result.get("comment_bytes")
    if not comment_bytes:
        return jsonify({"error": "批注版文档不可用", "code": "NO_COMMENTS"}), 404

    original_name = task.get("filename", "thesis.docx")
    base_name = os.path.splitext(original_name)[0]
    buf = BytesIO(comment_bytes)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name=f"{base_name}_批注版.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.route("/api/v1/checkers", methods=["GET"])
@require_api_key
def list_checkers():
    checkers = [
        {"key": "cover_page",     "name": "封面格式检测",   "description": "检测论文封面的格式、字体、字号等"},
        {"key": "commitment",     "name": "承诺书检测",     "description": "检测诚信承诺书内容、签名图片、日期"},
        {"key": "toc",            "name": "目录格式检测",   "description": "检测目录的格式、页码对应等"},
        {"key": "chapter_titles", "name": "章节标题格式检测","description": "检测各章节标题的编号和格式"},
        {"key": "body_text",      "name": "正文格式检测",   "description": "检测正文字体、字号、行距、段落格式"},
        {"key": "ref",            "name": "参考文献格式检测","description": "检测参考文献的格式和引用规范"},
        {"key": "figures_tables", "name": "图表公式术语检测","description": "检测图、表、公式编号及术语使用"},
        {"key": "header_footer",  "name": "页眉页脚格式检测","description": "检测页眉页脚和页码设置"},
        {"key": "ack_appendix",   "name": "致谢和附录格式检测","description": "检测致谢和附录的格式"},
        {"key": "structure",      "name": "论文整体结构检测","description": "检测论文章节顺序和整体结构"},
    ]
    return jsonify({"checkers": checkers}), 200


def main():
    parser = argparse.ArgumentParser(description="论文格式检测系统 - API 服务")
    parser.add_argument("--host", default="0.0.0.0", help="监听地址（默认 0.0.0.0）")
    parser.add_argument("--port", type=int, default=5080, help="监听端口（默认 5080）")
    parser.add_argument("--api-key", default=None, help="API Key 鉴权")
    parser.add_argument("--debug", action="store_true", help="调试模式")
    args = parser.parse_args()

    global API_KEY
    API_KEY = args.api_key

    os.makedirs(UPLOAD_DIR, exist_ok=True)
    os.makedirs(REPORTS_DIR, exist_ok=True)

    print("=" * 55)
    print("  论文格式检测系统 - REST API 服务")
    print("=" * 55)
    print(f"  地址: http://{args.host}:{args.port}")
    print(f"  鉴权: {'API Key 已启用' if API_KEY else '无鉴权（开放访问）'}")
    print()
    print("  接口列表:")
    print("    GET  /api/v1/health                      健康检查")
    print("    POST /api/v1/check                       上传论文并检测")
    print("    GET  /api/v1/status/<task_id>             查询任务状态")
    print("    GET  /api/v1/report/<task_id>             获取检测报告 (JSON/MD文本)")
    print("    GET  /api/v1/report/<task_id>/download    下载汇总报告 (.md 文件)")
    print("    GET  /api/v1/download/<task_id>           下载高亮文档 (.docx 文件)")
    print("    GET  /api/v1/download/<task_id>/comments  下载批注文档 (.docx 文件)")
    print("    GET  /api/v1/checkers                    列出检测项")
    print("=" * 55)

    app.run(host=args.host, port=args.port, debug=args.debug, threaded=True)


if __name__ == "__main__":
    main()
