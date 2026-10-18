# Designed by TiantianYe
#!/usr/bin/env python3
"""通过 WorkBuddy editor_sdk (edsdk) 替换/删除 Word 文档中的图片。

定位方式：doc_get_images 返回的「drawing 类嵌入图片」按文档顺序排列，
用 1-based 序号指定要操作的图片（.doc 老式 pict 图片不会出现在列表中，
但保留在文档里——以 unzip -l 文件.docx | grep media 计数核对总量）。

依赖：WorkBuddy 内置 editor_sdk（默认自动探测 edsdk.py 路径）。

用法：
  # 只列出文档中的图片及序号（不修改）
  python3 update_docx_images.py /abs/报告.docx --dry-run

  # 替换第5、6张为修复后的新截图；删除第17张
  python3 update_docx_images.py /abs/报告.docx \
      --replace 5=/abs/截图/新图1.png --replace 6=/abs/截图/新图2.png --delete 17

注意：
- --replace/--delete 中的序号都按【操作前列表】计；替换不改变位置，
  多个删除按从大到小自动排序执行。
- .doc 源保存时必须转 .docx（内核只写 OOXML）；.docx 源原格式保存。
"""
import argparse
import json
import subprocess
import os
import re
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


def call(edsdk, tool, args):
    r = subprocess.run(["python3", edsdk, "call", tool, "--json",
                        json.dumps(args, ensure_ascii=False)],
                       capture_output=True, text=True)
    out = r.stdout
    if '"ok": false' in out or "JSON-RPC 错误" in out:
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


def list_images(edsdk, file_id):
    m = call(edsdk, "doc_get_images", {"file_id": file_id})
    return m.get("images", [])


def main():
    ap = argparse.ArgumentParser(description="替换/删除 Word 文档中的图片（edsdk）")
    ap.add_argument("docx", help="Word 文件路径（.docx 或 .doc）")
    ap.add_argument("--replace", action="append", default=[], metavar="N=图片路径",
                    help="把第N张(1-based,文档顺序)替换为指定图片，可多次")
    ap.add_argument("--delete", action="append", default=[], type=int, metavar="N",
                    help="删除第N张所在的整个段落，可多次")
    ap.add_argument("--dry-run", action="store_true", help="只列出图片，不做修改")
    ap.add_argument("--out", help="另存路径；缺省覆盖原文件（.doc 会自动转 .docx 并需确认）")
    ap.add_argument("--edsdk", default=None, help="edsdk.py 路径（缺省自动探测）")
    args = ap.parse_args()

    edsdk = resolve_edsdk(args.edsdk)

    file_id = os.path.abspath(args.docx)
    # 后台打开（若已在用户预览中打开则复用其实例），并从响应解析真实 file_id
    r = subprocess.run(["python3", edsdk, "call", "open_file", "--json",
                        json.dumps({"file_path": file_id, "open_with_existing": True})],
                       capture_output=True, text=True)
    fm = re.search(r"file_id=(.+)", r.stdout or "")
    if fm:
        file_id = fm.group(1).strip()
        print(f"file_id: {file_id}")
    if not wait_ready(edsdk, file_id):
        sys.exit("!! 文档打开超时（重试或重启 editor_sdk）")

    images = list_images(edsdk, file_id)
    print(f"文档中共 {len(images)} 张 drawing 图片（按文档顺序）：")
    for i, img in enumerate(images):
        print(f"  {i+1:2d}. idx={img['index']} url={img['image_url'][:60]}...")
    if args.dry_run:
        return
    if not args.replace and not args.delete:
        print("(未指定 --replace/--delete，仅列出。加 --dry-run 可避免此提示)")

    # 替换（不改变图片位置，按用户给定顺序执行）
    for spec in args.replace:
        pos, _, path = spec.partition("=")
        pos = int(pos)
        if not (1 <= pos <= len(images)):
            sys.exit(f"!! 序号越界: {pos}（共 {len(images)} 张）")
        if not os.path.isabs(path):
            sys.exit(f"!! 图片路径必须为绝对路径: {path}")
        img = images[pos - 1]
        call(edsdk, "doc_replace_image", {
            "file_id": file_id, "idx": img["index"],
            "old_image_url": img["image_url"], "new_image_path": path,
        })
        print(f"✓ 替换第 {pos} 张 → {os.path.basename(path)}")

    # 删除（位置按操作前列表计；从大到小执行，避免偏移）
    if args.delete:
        for pos in sorted(args.delete, reverse=True):
            if not (1 <= pos <= len(images)):
                sys.exit(f"!! 序号越界: {pos}")
            # 删除会改变坐标：每次删除前重新获取最新位置
            cur = list_images(edsdk, file_id)
            img = cur[pos - 1]
            call(edsdk, "doc_delete_paragraph", {"file_id": file_id, "idx": img["index"]})
            print(f"✓ 删除第 {pos} 张所在段落 (idx={img['index']})")

    # 保存
    save_args = {"file_id": file_id}
    if args.out:
        save_args["file_path"] = args.out
    if args.docx.lower().endswith(".doc"):
        save_args["allow_format_change"] = True
        save_args.setdefault("file_path",
                             os.path.splitext(args.docx)[0] + ".docx")
        print("(源为 .doc，内核只写 OOXML：将另存为 .docx)")
    call(edsdk, "save_file", save_args)
    print(f"✓ 已保存: {save_args.get('file_path', file_id)}")


if __name__ == "__main__":
    main()
