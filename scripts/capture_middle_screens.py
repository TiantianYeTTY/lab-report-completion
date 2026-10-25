# Designed by TiantianYe
#!/usr/bin/env python3
"""长笔记本补截中间屏：capture_screenshots.py 的补充。

问题背景：capture_screenshots.py 对超一屏的笔记本只截首尾两屏；
当内容高度 > 2×视口时（如 5 个 code cell 的折线图笔记本，内容 3270px vs 视口 647px），
首尾两屏之间会跳过中间 cell 的输出（函数定义输出、中间图表）。

本脚本打开同一笔记本（Jupyter + Chrome 服务须仍在运行），滚动到内容中点再截一张，
存为 `<stem>_3.png`。plan.json 中按内容顺序排列：`_1, _3, _2`（顶部→中间→底部）。

前置条件（与 capture_screenshots.py 相同）：
1. Jupyter Notebook 服务（--notebooks-dir 所在目录）：--port=8899 --token=''
2. Chrome 带远程调试：--remote-debugging-port=9222 --remote-allow-origins='*'
3. 宿主终端已授予"屏幕录制"权限。
依赖：pip install websocket-client pyobjc-framework-Quartz

用法：
  python3 capture_middle_screens.py --notebooks-dir /abs/笔记本目录 --out-dir /abs/截图目录 \
      [--notebooks a.ipynb b.ipynb] [--base-url http://localhost:8899/notebooks/]

行为：--notebooks 缺省时自动探测"内容高度 > 2×视口"的笔记本（即两屏装不下的）；
手动指定时无条件补截。输出文件名 `<stem>_3.png`，已存在则跳过（可重复运行）。
"""
import argparse
import json
import os
import subprocess
import sys
import time
import urllib.parse
import urllib.request

try:
    from websocket import create_connection
except ImportError:
    sys.exit("缺少依赖：请先 pip install websocket-client")
try:
    import Quartz
except ImportError:
    sys.exit("缺少依赖：请先 pip install pyobjc-framework-Quartz")


# ---------------- CDP（与 capture_screenshots.py 同款逻辑） ----------------

def make_opener():
    """绕过环境代理（HTTP_PROXY 会劫持 localhost 请求）。"""
    return urllib.request.build_opener(urllib.request.ProxyHandler({}))


def probe_cdp():
    for host in ("http://[::1]:9222", "http://127.0.0.1:9222"):
        try:
            make_opener().open(host + "/json/version", timeout=3).read()
            return host
        except Exception:
            continue
    sys.exit("!! Chrome DevTools 端口不可达。请先按 §4.1 启动 Chrome（remote-debugging-port=9222）")


class Tab:
    def __init__(self, cdp, url, timeout=60):
        req = urllib.request.Request(
            cdp + "/json/new?" + urllib.parse.quote(url, safe=":/?=&.-"), method="PUT")
        self.cdp_base = cdp
        self.info = json.loads(make_opener().open(req, timeout=10).read())
        self.ws = create_connection(self.info["webSocketDebuggerUrl"], timeout=timeout)
        self._id = 0

    def send(self, method, params=None):
        self._id += 1
        self.ws.send(json.dumps({"id": self._id, "method": method, "params": params or {}}))
        while True:
            msg = json.loads(self.ws.recv())
            if msg.get("method") == "Page.javascriptDialogOpening":
                try:
                    self.ws.send(json.dumps({
                        "id": 999900 + self._id,
                        "method": "Page.handleJavaScriptDialog",
                        "params": {"accept": True},
                    }))
                except Exception:
                    pass
                continue
            if msg.get("id") == self._id:
                return msg.get("result", {})

    def evaluate(self, expr, quiet=False):
        try:
            r = self.send("Runtime.evaluate", {"expression": expr, "returnByValue": True})
            return r.get("result", {}).get("value")
        except Exception as e:
            if not quiet:
                print(f"   [evaluate异常] {type(e).__name__}: {e}")
            return None

    def close_tab(self):
        try:
            self.send("Page.navigate", {"url": "about:blank"})
            time.sleep(1)
        except Exception:
            pass
        try:
            make_opener().open(self.cdp_base + "/json/close/" + self.info["id"], timeout=5).read()
            self.ws.close()
        except Exception:
            pass


SCROLLER_JS = """(function(){
  function pick(){
    var b = document.scrollingElement;
    if (b && b.scrollHeight > b.clientHeight + 30) return b;
    var best=null, bestOver=0;
    var all = document.querySelectorAll('body,div,section,main');
    for (var i=0;i<all.length;i++){
      var e = all[i]; var ov = getComputedStyle(e).overflowY;
      if ((ov==='auto'||ov==='scroll') && e.scrollHeight > e.clientHeight){
        var over = e.scrollHeight - e.clientHeight;
        if (over > bestOver){ bestOver=over; best=e; }
      }
    }
    return best || b;
  }
  window.__scroller = pick();
  return window.__scroller.scrollHeight;
})()"""


def wait_stable(tab):
    """等渲染稳定：高度连续3次不变且 cell 已出现。"""
    prev = -1
    stable = 0
    for _ in range(60):
        h = tab.evaluate("(window.__scroller ? window.__scroller.scrollHeight "
                         ": document.scrollingElement.scrollHeight)", quiet=True)
        n = tab.evaluate("document.querySelectorAll('.cell').length", quiet=True)
        if h and h == prev and n and n > 0:
            stable += 1
            if stable >= 3:
                return h
        else:
            stable = 0
        prev = h
        time.sleep(0.5)
    return prev


def chrome_window_bounds(title_keyword):
    """按【进程PID + 窗口标题】找自动化 Chrome 的窗口（防误截用户窗口）。"""
    our = set(subprocess.run(["pgrep", "-f", "remote-debugging-port=9222"],
                             capture_output=True, text=True).stdout.split())
    wl = Quartz.CGWindowListCopyWindowInfo(
        Quartz.kCGWindowListOptionOnScreenOnly, Quartz.kCGNullWindowID)
    best = None
    for w in wl:
        if (w.get("kCGWindowOwnerName") == "Google Chrome"
                and str(w.get("kCGWindowOwnerPID")) in our):
            name = w.get("kCGWindowName") or ""
            if title_keyword in name and "Jupyter" in name:
                b = w.get("kCGWindowBounds")
                if best is None or b["Width"] * b["Height"] > best[2] * best[3]:
                    best = (b["X"], b["Y"], b["Width"], b["Height"])
    return best


def main():
    ap = argparse.ArgumentParser(description="长笔记本补截中间屏（_3.png）")
    ap.add_argument("--notebooks-dir", required=True)
    ap.add_argument("--out-dir", required=True)
    ap.add_argument("--base-url", default="http://localhost:8899/notebooks/")
    ap.add_argument("--notebooks", nargs="*", help="指定笔记本文件名；缺省=自动探测内容高>2×视口的")
    args = ap.parse_args()

    os.makedirs(args.out_dir, exist_ok=True)
    cdp = probe_cdp()
    print("CDP端点:", cdp)

    names = args.notebooks or sorted(
        f for f in os.listdir(args.notebooks_dir) if f.endswith(".ipynb")
        and not f.startswith("."))
    if not names:
        sys.exit(f"!! {args.notebooks_dir} 下没有 .ipynb 文件")

    done = 0
    for nb in names:
        stem = nb.replace(".ipynb", "")
        out_path = os.path.join(args.out_dir, f"{stem}_3.png")
        if os.path.exists(out_path):
            print(f"-> {stem}: _3.png 已存在，跳过")
            continue

        tab = Tab(cdp, args.base_url + nb)
        try:
            tab.send("Page.enable")
        except Exception:
            pass
        ready = False
        for _ in range(60):
            title = tab.evaluate("document.title", quiet=True)
            if title and stem in title:
                ready = True
                break
            time.sleep(1)
        if not ready:
            print(f"-> {stem}: !! 页面标题未就绪，跳过")
            tab.close_tab()
            continue

        sh = wait_stable(tab)
        ch = tab.evaluate("(window.__scroller||document.scrollingElement).clientHeight")
        if sh is None or ch is None or sh <= ch * 2:
            print(f"-> {stem}: 内容高 {sh} / 视口 {ch}，两屏可覆盖，无需补截")
            tab.close_tab()
            continue

        mid = (sh - ch) // 2
        tab.evaluate(f"(window.__scroller||document.scrollingElement).scrollTop = {mid}")
        time.sleep(0.8)
        try:
            tab.send("Page.bringToFront")
        except Exception:
            pass
        time.sleep(1.0)
        bounds = chrome_window_bounds(stem)
        if not bounds:
            print(f"-> {stem}: !! 找不到自动化 Chrome 窗口，跳过")
            tab.close_tab()
            continue
        subprocess.run(["screencapture", "-x", "-R",
                        f"{bounds[0]},{bounds[1]},{bounds[2]},{bounds[3]}", out_path], check=True)
        print(f"-> {stem}: sh={sh} ch={ch} mid={mid} | 截图 {stem}_3.png")
        done += 1
        tab.close_tab()
        time.sleep(1)

    print(f"补截完成：{done} 张（plan.json 中按 _1, _3, _2 内容顺序排列）")


if __name__ == "__main__":
    main()
