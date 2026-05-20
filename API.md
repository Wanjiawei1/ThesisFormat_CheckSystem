# 论文格式检测系统 - API 接口文档

> 版本：v1.1 | 基础路径：`http://<host>:<port>/api/v1`

---

## 1. 快速开始

### 启动服务

```bash
# 无鉴权（开发/内网环境）
python api_server.py --port 5080

# 调试模式
python api_server.py --port 5080 --debug
```

### 一行调用示例

```bash
# 方式一：传文件路径（云平台推荐）
curl -X POST http://localhost:5080/api/v1/check \
     -H "Content-Type: application/json" \
     -d '{"file_path": "/data/papers/论文.docx", "mode": "sync"}'

# 方式二：上传文件
curl -X POST http://localhost:5080/api/v1/check \
     -F "file=@毕业论文.docx"

# 带 API Key
curl -X POST http://localhost:5080/api/v1/check \
     -H "Authorization: Bearer YOUR_SECRET_KEY" \
     -H "Content-Type: application/json" \
     -d '{"file_path": "/data/papers/论文.docx"}'
```

---

## 2. 接口详情

### 2.1 健康检查

```
GET /api/v1/health
```

**无需鉴权**

**响应示例：**
```json
{
  "status": "ok",
  "service": "thesis-format-checker",
  "version": "1.0.0",
  "time": "2026-04-28T10:30:00"
}
```

---

### 2.2 论文格式检测

```
POST /api/v1/check
```

支持两种调用方式（二选一）：

#### 传文件路径（推荐，云平台调用场景）

```
POST /api/v1/check
Content-Type: application/json

{
  "file_path": "/data/papers/2026001_张三_毕业论文.docx",
  "mode": "sync"
}
```

**请求参数：**

| 参数          | 类型     | 必填  | 说明                                    |
| ----------- | ------ | --- | ------------------------------------- |
| `file_path` | String | 必填  | 论文 .docx 文件的服务器本地绝对路径                 |
| `mode`      | String | ❌   | `sync`（默认）同步返回 / `async` 异步返回 task_id |

> `file_path` 和 `file` 二选一，优先识别 `file_path`（JSON 请求）。

**同步模式响应（200）：**

```json
{
  "task_id": "20260428_103000_a1b2c3d4",
  "status": "completed",
  "filename": "毕业论文.docx",
  "summary": {
    "pass": 6,
    "warn": 2,
    "fail": 1,
    "skip": 1,
    "total": 10
  },
  "checks": {
    "cover_page": {
      "name": "封面格式检测",
      "status": "pass",
      "issues": [],
      "detail": "..."
    },
    "toc": {
      "name": "目录格式检测",
      "status": "warn",
      "issues": [
        {"message": "目录第3项页码与正文不一致", "severity": "warning"}
      ],
      "detail": "..."
    },
    "body_text": {
      "name": "正文格式检测",
      "status": "fail",
      "issues": [
        {"message": "第2章正文字体应为宋体，实际为黑体", "severity": "error"},
        {"message": "第3段行距不符合要求", "severity": "warning"}
      ],
      "detail": "..."
    }
  },
  "report_markdown": "# 检测汇总报告\n\n...",
  "highlight_available": true,
  "comment_available": true
}
```

**异步模式响应（202）：**

```json
{
  "task_id": "20260428_103000_a1b2c3d4",
  "status": "pending",
  "message": "检测任务已提交，请通过 /api/v1/status/{task_id} 查询进度"
}
```

---

### 2.3 查询任务状态（异步模式）

```
GET /api/v1/status/<task_id>
```

**响应示例：**

```json
{
  "task_id": "20260428_103000_a1b2c3d4",
  "status": "completed",
  "filename": "毕业论文.docx",
  "created_at": "2026-04-28T10:30:00",
  "result": { ... }
}
```

**status 取值：**
- `pending` — 排队中
- `running` — 检测进行中
- `completed` — 已完成
- `error` — 出错

---

### 2.4 获取检测报告

```
GET /api/v1/report/<task_id>?format=json
```

**参数：**

| 参数 | 说明 |
|------|------|
| `format` | `json`（默认）返回完整 JSON / `markdown` 返回 Markdown 文本 |

**Markdown 格式响应（Content-Type: text/plain）：**

```markdown
# 检测汇总报告

## 封面格式检测
✅ 封面格式通过

## 正文格式检测
❌ 第2章正文字体应为宋体，实际为黑体
⚠️ 第3段行距不符合要求
...
```

---

### 2.5 下载汇总报告文件（MD）

```
GET /api/v1/report/<task_id>/download
```

**说明：** 下载包含所有模块检测结果的汇总 Markdown 文件（`.md`），供云管理平台存储或展示。

**响应：** 直接返回 `.md` 文件下载。

- Content-Type: `text/markdown; charset=utf-8`
- 文件名格式: `{原文件名}_检测报告.md`

**报告结构：**
```markdown
# 论文格式检测报告

- **文件名**: 毕业论文.docx
- **检测时间**: 2026-04-28 10:30:00

## 总体统计
| 指标 | 数量 |
|------|------|
| ✅ 通过 | 6 |
| ⚠️ 警告 | 2 |
| ❌ 失败 | 1 |
| ➖ 跳过 | 1 |

> **结论**: ❌ 检测未通过，存在 1 项失败

## 检测概览
| 序号 | 检测模块 | 状态 | 问题数 |
|------|----------|------|--------|
| 1 | 封面格式检测 | ✅ 通过 | - |
| 2 | 正文格式检测 | ❌ 失败 | 3 |
| ... | ... | ... | ... |

## 详细报告
### ❌ 正文格式检测
**发现问题 3 个：**
- ❌ 第2章正文字体应为宋体，实际为黑体
- ⚠️ 第3段行距不符合要求
...

---
*本报告由论文格式检测系统自动生成*
```

**调用示例：**
```bash
# 直接下载保存
curl -o 检测报告.md http://localhost:5080/api/v1/report/20260428_103000_a1b2c3d4/download

# 带 API Key
curl -H "Authorization: Bearer YOUR_KEY" \
     -o 检测报告.md \
     http://localhost:5080/api/v1/report/20260428_103000_a1b2c3d4/download
```

---

### 2.6 下载高亮版文档

```
GET /api/v1/download/<task_id>
```

**说明：** 下载带有错误高亮标注的 Word 文档（`.docx`），问题段落以黄色背景标记。

**响应：** 直接返回 `.docx` 文件下载。

- Content-Type: `application/vnd.openxmlformats-officedocument.wordprocessingml.document`
- 文件名格式: `{原文件名}_高亮版.docx`

**调用示例：**
```bash
# 直接下载保存
curl -o 高亮版.docx http://localhost:5080/api/v1/download/20260428_103000_a1b2c3d4

# 带 API Key
curl -H "Authorization: Bearer YOUR_KEY" \
     -o 高亮版.docx \
     http://localhost:5080/api/v1/download/20260428_103000_a1b2c3d4
```

---

### 2.7 下载批注版文档

```
GET /api/v1/download/<task_id>/comments
```

**说明：** 下载带有格式问题批注的 Word 文档（`.docx`）。系统自动检测各模块的格式问题，并以 Word 批注（Comment）的形式标注在对应位置，方便在 Word 中逐条查看和修改。

**响应：** 直接返回 `.docx` 文件下载。

- Content-Type: `application/vnd.openxmlformats-officedocument.wordprocessingml.document`
- 文件名格式: `{原文件名}_批注版.docx`

**调用示例：**
```bash
# 直接下载保存
curl -o 批注版.docx http://localhost:5080/api/v1/download/20260428_103000_a1b2c3d4/comments

# 带 API Key
curl -H "Authorization: Bearer YOUR_KEY" \
     -o 批注版.docx \
     http://localhost:5080/api/v1/download/20260428_103000_a1b2c3d4/comments
```

---

### 2.8 列出所有检测项

```
GET /api/v1/checkers
```

**响应：**

```json
{
  "checkers": [
    {"key": "cover_page",     "name": "封面格式检测",    "description": "检测论文封面的格式、字体、字号等"},
    {"key": "commitment",     "name": "承诺书检测",      "description": "检测诚信承诺书内容、签名图片、日期"},
    {"key": "toc",            "name": "目录格式检测",    "description": "检测目录的格式、页码对应等"},
    {"key": "chapter_titles", "name": "章节标题格式检测", "description": "检测各章节标题的编号和格式"},
    {"key": "body_text",      "name": "正文格式检测",    "description": "检测正文字体、字号、行距、段落格式"},
    {"key": "ref",            "name": "参考文献格式检测", "description": "检测参考文献的格式和引用规范"},
    {"key": "figures_tables", "name": "图表公式术语检测", "description": "检测图、表、公式编号及术语使用"},
    {"key": "header_footer",  "name": "页眉页脚格式检测", "description": "检测页眉页脚和页码设置"},
    {"key": "ack_appendix",   "name": "致谢和附录格式检测","description": "检测致谢和附录的格式"},
    {"key": "structure",      "name": "论文整体结构检测", "description": "检测论文章节顺序和整体结构"}
  ]
}
```

---

## 3. 鉴权方式

如果启动时设置了 `--api-key`，所有接口（除 `/health`）需要携带 API Key：

```bash
# 方式一：Header（推荐）
curl -H "Authorization: Bearer YOUR_KEY" http://...

# 方式二：Query 参数
curl "http://...?api_key=YOUR_KEY"
```

---

## 4. 错误码

| HTTP 状态码 | 场景 |
|-------------|------|
| 400 | 缺少文件 / 格式不支持 |
| 401 | API Key 无效 |
| 404 | task_id 不存在 |
| 409 | 任务未完成就尝试获取报告 |
| 500 | 服务端内部错误 |

**错误响应格式：**
```json
{
  "error": "错误描述",
  "code": "ERROR_CODE"
}
```

---

## 5. 调用流程

### 同步模式（推荐用于快速集成）

```
云管理平台                       API 服务
  │                               │
  │── POST /check (JSON) ────────>│
  │   { "file_path": "/xxx.docx"} │
  │                               │── 根据路径读取文件
  │                               │── 执行全部检测
  │                               │── 生成报告
  │<── 200 完整结果 ──────────────│
  │                               │
  │── GET /report/<id>/download ─>│  ← 下载 MD 汇总报告
  │<── xxx_检测报告.md ───────────│
  │                               │
  │── GET /download/<task_id> ───>│  ← 下载高亮版文档
  │<── xxx_高亮版.docx ──────────│
  │                               │
  │── GET /download/<id>/comments>│  ← 下载批注版文档
  │<── xxx_批注版.docx ──────────│
```

### 异步模式（推荐大文件 / 不阻塞调用方）

```
云管理平台                       API 服务
  │                               │
  │── POST /check (async) ───────>│
  │<── 202 { task_id } ──────────│
  │                               │── 后台执行检测...
  │                               │
  │── GET /status/<task_id> ─────>│  (轮询)
  │<── { status: "completed" } ──│
  │                               │
  │── GET /report/<id>/download ─>│  ← 下载 MD 汇总报告
  │<── xxx_检测报告.md ───────────│
  │                               │
  │── GET /download/<task_id> ───>│  ← 下载高亮版文档
  │<── xxx_高亮版.docx ──────────│
  │                               │
  │── GET /download/<id>/comments>│  ← 下载批注版文档
  │<── xxx_批注版.docx ──────────│
```

---

## 6. 集成示例（Python）

```python
import requests

API_BASE = "http://your-server:5080/api/v1"
HEADERS = {"Authorization": "Bearer YOUR_KEY"}

# ---- 方式一：传路径（云平台调用场景） ----
resp = requests.post(
    f"{API_BASE}/check",
    headers=HEADERS,
    json={
        "file_path": "/data/papers/2026001_张三_毕业论文.docx",
        "mode": "sync",
    },
)

# ---- 方式二：上传文件 ----
# with open("论文.docx", "rb") as f:
#     resp = requests.post(
#         f"{API_BASE}/check",
#         headers=HEADERS,
#         files={"file": ("论文.docx", f, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
#         data={"mode": "sync"},
#     )

result = resp.json()
print(f"检测状态: {result['status']}")
print(f"通过: {result['summary']['pass']}, 警告: {result['summary']['warn']}, 失败: {result['summary']['fail']}")

# 查看各项详情
for key, check in result["checks"].items():
    icon = {"pass": "✅", "warn": "⚠️", "fail": "❌", "skip": "➖"}.get(check["status"], "?")
    print(f"  {icon} {check['name']}: {check['status']}")
    for issue in check["issues"]:
        print(f"      - [{issue['severity']}] {issue['message']}")

# 下载汇总报告（MD 文件）
report_dl = requests.get(f"{API_BASE}/report/{result['task_id']}/download", headers=HEADERS)
base = result["filename"].rsplit(".", 1)[0]
with open(f"{base}_检测报告.md", "wb") as f:
    f.write(report_dl.content)
print(f"已下载: {base}_检测报告.md")

# 下载高亮文档
if result["highlight_available"]:
    dl = requests.get(f"{API_BASE}/download/{result['task_id']}", headers=HEADERS)
    with open(f"{base}_高亮版.docx", "wb") as f:
        f.write(dl.content)
    print(f"已下载: {base}_高亮版.docx")

# 下载批注版文档
if result.get("comment_available"):
    dl = requests.get(f"{API_BASE}/download/{result['task_id']}/comments", headers=HEADERS)
    with open(f"{base}_批注版.docx", "wb") as f:
        f.write(dl.content)
    print(f"已下载: {base}_批注版.docx")
```

---

## 7. 集成示例（Java / Spring Boot）

```java
import okhttp3.*;
import java.io.File;

// ---- 方式一：传路径（推荐） ----
OkHttpClient client = new OkHttpClient();

String json = "{\"file_path\": \"/data/papers/论文.docx\", \"mode\": \"sync\"}";
RequestBody body = RequestBody.create(json, MediaType.parse("application/json"));

Request request = new Request.Builder()
    .url("http://your-server:5080/api/v1/check")
    .header("Authorization", "Bearer YOUR_KEY")
    .post(body)
    .build();

Response response = client.newCall(request).execute();
String result = response.body().string();
// 解析 JSON...
```

---

## 8. 注意事项

1. **文件大小限制**：单文件最大 50MB
2. **文件格式**：仅支持 `.docx`（不支持 `.doc`）
3. **路径要求**：传 `file_path` 时必须是 API 服务所在机器的本地绝对路径，确保服务进程有读取权限
4. **文件共享**：如果云平台和 API 服务不在同一台机器，需要挂载共享存储（NFS/SMB）或使用文件同步，确保路径可达
5. **并发**：服务支持多线程并发请求
6. **临时文件**：检测完成后服务端自动清理缓存文件，不会占用磁盘空间
7. **内存占用**：检测结果（含高亮文档）保存在内存中，单个任务约占几 MB 内存
7. **隐私**：论文仅在服务端本地处理，不会上传到外部服务
