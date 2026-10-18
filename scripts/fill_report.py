#!/usr/bin/env python3
"""在原版 Word 报告的副本上原位插入结果截图与实验小结（WorkBuddy editor_sdk / edsdk）。

核心保证：原版模板只读（自动复制副本再编辑），插入采用「段落结构树定位」——
即使查找标签只是段落前缀（如“五、实验小结”是完整标题的前缀），也不会插进标题中间。

用法：
  1. 编写计划文件 plan.json（结构见 --print-plan-template）：
     {
       "template_doc": "/abs/报告模板.doc",
       "out_docx": "/abs/完成版.docx",
       "img_w": 548, "img_h": 331,
       "sections": [
         {"find": "结果截图", "images": ["/abs/截图/1-1.png", "/abs/截图/1-2.png"]}
       ],
       "summary": {"find": "五、实验小结", "markdown": "/abs/实验小结.md"}
     }
  2. python3 fill_report.py plan.json

说明：
- sections 按文档顺序列出；每个 find 标签的“整段匹配”按文档顺序依次消耗
  （段落中除标签外还有其他文字的匹配会被跳过，如“格式要求：…结果截图居中…”）。
- 同一节多张图会在该节插入点按顺序各占一段（14.5cm 居中由 img_w/img_h 控制）。
- 工作副本 = template_doc 的复制（同目录，文件名加“-完成版”）；源为 .doc 时
  保存自动转 .docx（内核只写 OOXML），原版 .doc 始终不变。
"""
import argparse
import json
import os
import re
import shutil
import subprocess
import sys

EDSDK_CANDIDATES = (
    "/Applications/AI/WorkBuddy.app/Contents/Resources/app.asar.unpacked/resources/plugins/"
    "workbuddy-builtin/skills/workbuddy-local-office-edit/edsdk.py",
    "/Applications/AI/WorkBuddy.app/Contents/Resources/app.asar.unpacked/resources/plugins/"
    "workbuddy-builtin/skills/tencent-local-office-edit/edsdk.py",
)


def resolve_edsdk(explicit):
    """WorkBuddy 更新可能改变内置技能目录，按候选路径自动探测。"""
    if explicit:
        if not os.path.exists(explicit):
            sys.exit(f"!! 找不到 edsdk.py: {explicit}")
        return explicit
    for p in EDSDK_CANDIDATES:
        if os.path.exists(p):
            return p
    sys.exit("!! 找不到 edsdk.py。请定位后用 --edsdk 指定：\n"
             "   find /Applications/AI/WorkBuddy.app -name edsdk.py 2>/dev/null")

def wait_ready(edsdk, file_id, timeout=60):
    """.doc 的 open_file 是异步流式加载：轮询轻量查询直到文档真正可用。"""
    import time
    for i in range(timeout):
        r = subprocess.run(["python3", edsdk, "call", "doc_get_last_operable_pos",
                            "--json", json.dumps({"file_id": file_id})],
                           capture_output=True, text=True)
        if r.returncode == 0 and "document is not open" not in r.stdout and "error" not in r.stdout:
            return True
        time.sleep(1)
    return False

PLAN_TEMPLATE = """{
  "template_doc": "/abs/路径/报告模板.doc",
  "out_docx": "/abs/路径/完成版.docx",
  "img_w": 548,
  "img_h": 331,
  "sections": [
    {"find": "结果截图", "images": ["/abs/截图/题1-截图1.png", "/abs/截图/题1-截图2.png"]}
  ],
  "summary": {"find": "五、实验小结", "markdown": "/abs/实验小结.md"}
}"""


def call(edsdk, tool, args, check=True):
    r = subprocess.run(["python3", edsdk, "call", tool, "--json",
                        json.dumps(args, ensure_ascii=False)],
                       capture_output=True, text=True)
    out = r.stdout
    if check and ('"ok": false' in out or "JSON-RPC 错误" in out):
        print(f"!! {tool} 调用失败:\n{out[:600]}", file=sys.stderr)
        sys.exit(1)
    depth = 0; start = -1; objs = []
    for i, ch in enumerate(out):
        if ch == '{':
            if depth == 0:
                start = i
            depth += 1
        elif ch == '}':
            depth -= 1
            if depth == 0 and start >= 0:
                objs.append(out[start:i + 1])
    return json.loads(objs[-1]) if objs else None


def resolve_paragraph(edsdk, file_id, idx):
    """返回包含 idx 的段落节点（含 start_index/end_index）。"""
    m = call(edsdk, "doc_resolve_document_structure",
             {"file_id": file_id, "idx": idx, "mode": "full",
              "text_preview_length": 60})
    nodes = m.get("nodes", [])
    return nodes[0] if nodes else None


def insert_images(edsdk, file_id, insert_idx, image_paths, w, h):
    """在插入点放入 N 张图，各占一段：先插最后一张，再依次「插段落符+插前一张」。"""
    if not image_paths:
        return
    call(edsdk, "doc_insert_image",
         {"file_id": file_id, "idx": insert_idx,
          "image_path": image_paths[-1], "w": w, "h": h})
    print(f"      + {os.path.basename(image_paths[-1])}")
    for k in range(len(image_paths) - 2, -1, -1):
        call(edsdk, "doc_insert_paragraph", {"file_id": file_id, "idx": insert_idx})
        call(edsdk, "doc_insert_image",
             {"file_id": file_id, "idx": insert_idx,
              "image_path": image_paths[k], "w": w, "h": h})
        print(f"      + {os.path.basename(image_paths[k])}")


def next_label_paragraph(edsdk, file_id, label, consumed):
    """查找标签，返回第一个「整段恰好为标签」且未被消耗的匹配。"""
    m = call(edsdk, "doc_find", {"file_id": file_id, "text": label})
    locs = sorted(m.get("locations", []), key=lambda x: x["begin"])
    for loc in locs:
        if loc["begin"] in consumed:
            continue
        para = resolve_paragraph(edsdk, file_id, loc["begin"])
        if not para:
            continue
        # 整段恰好等于标签（纯标签段落），且段落边界与匹配一致
        if para["start_index"] == loc["begin"] and para["end_index"] == loc["end"]:
            consumed.add(loc["begin"])
            return para
    return None


def main():
    ap = argparse.ArgumentParser(description="原版 Word 报告原位填充（截图+小结）")
    ap.add_argument("plan", help="计划文件 plan.json")
    ap.add_argument("--print-plan-template", action="store_true",
                    help="打印计划文件模板后退出")
    ap.add_argument("--edsdk", default=None, help="edsdk.py 路径（缺省自动探测）")
    args = ap.parse_args()

    if args.print_plan_template:
        print(PLAN_TEMPLATE)
        return
    edsdk = resolve_edsdk(args.edsdk)

    plan = json.load(open(args.plan, encoding="utf-8"))
    template = plan["template_doc"]
    out_docx = plan["out_docx"]
    w, h = plan.get("img_w", 548), plan.get("img_h", 331)

    # 工作副本（原版模板绝不直接编辑）
    work = plan.get("work_doc") or os.path.splitext(template)[0] + "-完成版.doc"
    if not os.path.exists(work):
        shutil.copyfile(template, work)
        print(f"已复制原版模板为工作副本: {work}")
    elif os.path.getmtime(work) < os.path.getmtime(template):
        print(f"(提示) 工作副本比模板旧: {work}")

    file_id = os.path.abspath(work)
    # 打开工作副本（若已在用户预览中打开则复用其实例），并从响应解析真实 file_id
    r = subprocess.run(["python3", edsdk, "call", "open_file", "--json",
                        json.dumps({"file_path": file_id, "open_with_existing": True})],
                       capture_output=True, text=True)
    fm = re.search(r"file_id=(.+)", r.stdout or "")
    if fm:
        file_id = fm.group(1).strip()
        print(f"file_id: {file_id}")
    if not wait_ready(edsdk, file_id):
        sys.exit("!! 文档打开超时（重试或重启 editor_sdk）")

    # 各节截图
    consumed = set()
    for si, sec in enumerate(plan.get("sections", []), 1):
        label = sec["find"]
        images = sec.get("images", [])
        print(f"[{si}] 查找标签「{label}」")
        para = next_label_paragraph(edsdk, file_id, label, consumed)
        if not para:
            print(f"   !! 未找到未消耗的整段标签，跳过")
            continue
        insert_idx = para["end_index"] + 1  # +1 跳过段落分隔符 → 下一空段起点
        print(f"   插入点: {insert_idx}（共 {len(images)} 张图）")
        insert_images(edsdk, file_id, insert_idx, images, w, h)

    # 实验小结（锚点可以是长段落的前缀，结构树定位保证插在其段落之后）
    summary = plan.get("summary")
    if summary:
        anchor = summary["find"]
        print(f"[小结] 查找锚点「{anchor}」")
        m = call(edsdk, "doc_find", {"file_id": file_id, "text": anchor})
        locs = sorted(m.get("locations", []), key=lambda x: x["begin"])
        if not locs:
            print("   !! 未找到锚点，跳过")
        else:
            para = resolve_paragraph(edsdk, file_id, locs[0]["begin"])
            insert_idx = para["end_index"] + 1
            md = summary["markdown"]
            call(edsdk, "doc_insert_markdown",
                 {"file_id": file_id, "idx": insert_idx,
                  "markdown": "file://" + md if not md.startswith("file://") else md})
            print(f"   小结已插入（锚点段落 end={para['end_index']}）")

    # 保存
    save_args = {"file_id": file_id, "file_path": out_docx}
    if work.lower().endswith(".doc"):
        save_args["allow_format_change"] = True
    call(edsdk, "save_file", save_args)
    print(f"✓ 已保存: {out_docx}")

    subprocess.run(["python3", edsdk, "call", "close_file", "--json",
                    json.dumps({"file_id": file_id})], capture_output=True, text=True)
    print("完成。建议校验：unzip -l 完成版.docx | grep media 核对图片数；"
          "pandoc 转 html 目检版式。")


if __name__ == "__main__":
    main()
