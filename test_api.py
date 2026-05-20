import requests

API = "http://127.0.0.1:5080/api/v1"
FILE_PATH = "/root/wanjiawei/毕业论文-沈丞琦.docx"

# 1. 提交检测
resp = requests.post(f"{API}/check", json={"file_path": FILE_PATH})
result = resp.json()

task_id = result["task_id"]
filename = result["filename"]
base = filename.rsplit(".", 1)[0]

print(f"任务ID: {task_id}")
print(f"文件名: {filename}")
print(f"状态: {result['status']}")
print(f"通过: {result['summary']['pass']}, 警告: {result['summary']['warn']}, 失败: {result['summary']['fail']}")

# 2. 下载 MD 汇总报告
md = requests.get(f"{API}/report/{task_id}/download")
md_name = f"{base}_检测报告.md"
with open(md_name, "wb") as f:
    f.write(md.content)
print(f"已下载: {md_name}")

# 3. 下载高亮版文档
hl = requests.get(f"{API}/download/{task_id}")
hl_name = f"{base}_高亮版.docx"
with open(hl_name, "wb") as f:
    f.write(hl.content)
print(f"已下载: {hl_name}")

# 4. 下载批注版文档
cm = requests.get(f"{API}/download/{task_id}/comments")
if cm.status_code == 200:
    cm_name = f"{base}_批注版.docx"
    with open(cm_name, "wb") as f:
        f.write(cm.content)
    print(f"已下载: {cm_name}")
else:
    print(f"批注版不可用: {cm.json()}")
