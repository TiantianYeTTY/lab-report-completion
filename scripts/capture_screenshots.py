# Designed by TiantianYe
#!/usr/bin/env python3
"""真机截图管线：用真实 Chrome 窗口打开已执行过的 Jupyter 笔记本，系统级 screencapture 截屏。

前置条件（由调用方先启动，脚本会自动探测并给出启动命令提示）：
1. Jupyter Notebook 服务（在笔记本所在目录启动）：
   jupyter-notebook --no-browser --port=8899 --NotebookApp.token='' --NotebookApp.disable_check_xsrf=True
2. Chrome（带远程调试，专用 profile）：
   "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" --no-sandbox --test-type \
     --remote-debugging-port=9222 --remote-allow-origins='*' \
     --user-data-dir=/abs/path/chrome_profile --no-first-run --no-default-browser-check \
     --hide-crash-restore-bubble --window-size=1400,845 --window-position=30,25 about:blank
3. 宿主终端已授予"屏幕录制"权限。

依赖：pip install websocket-client pyobjc-framework-Quartz

用法示例：
  python3 capture_screenshots.py --notebooks-dir /abs/笔记本目录 --out-dir /abs/截图目录 \
      [--base-url http://localhost:8899/notebooks/] [--notebooks a.ipynb b.ipynb] \
      [--img-w 548 --img-h 331]

行为：单 tab 持久导航（避免多 tab 开关导致 DevTools 通道超时）；自动关闭 Notebook7
迁移横幅；内容超一屏时自动输出 _1/_2 两张连续截图。
"""
import argparse
import json
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


# ---------------- CDP 连接 ----------------

def make_opener():
    """绕过环境代理（HTTP_PROXY 会劫持 localhost 请求）。"""
    return urllib.request.build_opener(urllib.request.ProxyHandler({}))


def probe_cdp():
    """CDP 端点双栈探测：IPv4 可能被环境占用，优先试 IPv6。"""
    for host in ("http://[::1]:9222", "http://127.0.0.1:9222"):
        try:
            make_opener().open(host + "/json/version", timeout=3).read()
            return host
        except Exception:
            continue
    print("!! Chrome DevTools 端口不可达。请先启动 Chrome，例如：", file=sys.stderr)
    print('   "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" --no-sandbox --test-type '
          '--remote-debugging-port=9222 --remote-allow-origins=\'*\' --user-data-dir=/tmp/chrome_profile '
          '--no-first-run --no-default-browser-check --hide-crash-restore-bubble '
          '--window-size=1400,845 --window-position=30,25 about:blank', file=sys.stderr)
    sys.exit(1)


class Tab:
    """一个 CDP 页面标签，带 JS 对话框自动应答（Jupyter 关 tab 的 beforeunload 会卡死操作）。"""

    def __init__(self, cdp, url, timeout=60):
        self.cdp = cdp
        self.opener = make_opener()
        req = urllib.request.Request(
            cdp + "/json/new?" + urllib.parse.quote(url, safe=":/?=&.-"), method="PUT")
        self.info = json.loads(self.opener.open(req, timeout=10).read())
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

    def evaluate(self, expr, retries=2, quiet=False):
        for attempt in range(retries + 1):
            try:
                r = self.send("Runtime.evaluate", {"expression": expr, "returnByValue": True})
                return r.get("result", {}).get("value")
            except Exception as e:
                if not quiet:
                    print(f"   [evaluate异常#{attempt}] {type(e).__name__}: {e}", flush=True)
                if attempt == retries:
                    return None
                time.sleep(1)
                try:
                    self.ws = create_connection(self.info["webSocketDebuggerUrl"], timeout=60)
                except Exception:
                    pass
        return None

    def close_tab(self):
        try:
            self.send("Page.navigate", {"url": "about:blank"})
            time.sleep(1)
        except Exception:
            pass
        try:
            self.opener.open(self.cdp + "/json/close/" + self.info["id"], timeout=5).read()
            self.ws.close()
        except Exception:
            pass


# ---------------- 窗口定位与截屏 ----------------

def get_our_pids():
    out = subprocess.run(["pgrep", "-f", "remote-debugging-port=9222"],
                         capture_output=True, text=True).stdout.split()
    return set(out)


def chrome_window_bounds(title_keyword):
    """按【进程PID + 窗口标题】找自动化 Chrome 的窗口（防止误截用户自己的浏览器窗口）。"""
    our = get_our_pids()
    wl = Quartz.CGWindowListCopyWindowInfo(
        Quartz.kCGWindowListOptionOnScreenOnly, Quartz.kCGNullWindowID)
    best = None
    for w in wl:
        if (w.get("kCGWindowOwnerName") == "Google Chrome"
                and str(w.get("kCGWindowOwnerPID")) in our):
            name = w.get("kCGWindowName") or ""
            if title_keyword in name and "Jupyter" in name:
                b = w.get("kCGWindowBounds")
                area = b["Width"] * b["Height"]
                if best is None or area > best[2] * best[3]:
                    best = (b["X"], b["Y"], b["Width"], b["Height"])
    return best


def capture(path, bounds):
    x, y, w, h = bounds
    subprocess.run(["screencapture", "-x", "-R", f"{x},{y},{w},{h}", path], check=True)


# ---------------- 渲染等待 ----------------

SCROLLER_JS = """(function(){
  function pick(){
    var b = document.scrollingElement;
    if (b && b.scrollHeight > b.clientHeight + 30) return {el:b, sh:b.scrollHeight, ch:b.clientHeight};
    var best=null, bestOver=0;
    var all = document.querySelectorAll('body,div,section,main');
    for (var i=0;i<all.length;i++){
      var e = all[i];
      var ov = getComputedStyle(e).overflowY;
      if ((ov==='auto'||ov==='scroll') && e.scrollHeight > e.clientHeight){
        var over = e.scrollHeight - e.clientHeight;
        if (over > bestOver){ bestOver=over; best=e; }
      }
    }
    if (best) return {el:best, sh:best.scrollHeight, ch:best.clientHeight};
    return {el:b, sh:b?b.scrollHeight:0, ch:b?b.clientHeight:0};
  }
  var p = pick();
  window.__scroller = p.el;
  return {sh: p.sh, ch: p.ch, id: p.el.id || p.el.tagName};
})()"""

DISMISS_BANNER_JS = """(function(){
  var btns = document.querySelectorAll('button');
  for (var i=0;i<btns.length;i++){
    if (/don.t show anymore/i.test(btns[i].textContent)){ btns[i].click(); return 'clicked'; }
  }
  return 'not-found';
})()"""


def wait_stable(tab):
    """等渲染稳定：高度连续3次不变且 cell 已出现。"""
    prev = -1
    stable = 0
    for _ in range(60):
        h = tab.evaluate("(window.__scroller ? window.__scroller.scrollHeight "
                         ": document.scrollingElement.scrollHeight)")
        n_cells = tab.evaluate("document.querySelectorAll('.cell').length")
        if h and h == prev and n_cells and n_cells > 0:
            stable += 1
            if stable >= 3:
                return True
        else:
            stable = 0
        prev = h
        time.sleep(0.5)
    return False


def bring_front(tab):
    try:
        tab.send("Page.bringToFront")
    except Exception:
        pass


# ---------------- 主流程 ----------------

def main():
    ap = argparse.ArgumentParser(description="Jupyter 笔记本真机截图管线")
    ap.add_argument("--notebooks-dir", required=True, help="笔记本所在目录（绝对路径）")
    ap.add_argument("--out-dir", required=True, help="截图输出目录")
    ap.add_argument("--base-url", default="http://localhost:8899/notebooks/")
    ap.add_argument("--notebooks", nargs="*", help="要截的笔记本文件名；缺省=目录下全部 .ipynb")
    ap.add_argument("--img-w", type=int, default=548, help="插入用显示宽（96DPI像素，548≈14.5cm）")
    ap.add_argument("--img-h", type=int, default=331, help="插入用显示高")
    args = ap.parse_args()

    import os
    os.makedirs(args.out_dir, exist_ok=True)

    names = args.notebooks or sorted(
        f for f in os.listdir(args.notebooks_dir) if f.endswith(".ipynb")
        and not f.startswith("."))
    if not names:
        sys.exit(f"!! {args.notebooks_dir} 下没有 .ipynb 文件")

    cdp = probe_cdp()
    print("CDP端点:", cdp)
    tab = Tab(cdp, args.base_url + names[0])
    try:
        tab.send("Page.enable")
    except Exception:
        pass

    for nb in names:
        url = args.base_url + nb
        stem = nb.replace(".ipynb", "")
        print(f"-> 打开 {nb}")
        try:
            tab.send("Page.navigate", {"url": url})
        except Exception as e:
            print(f"   导航异常，重建连接: {e}")
            tab.ws = create_connection(tab.info["webSocketDebuggerUrl"], timeout=60)
            tab.send("Page.navigate", {"url": url})

        # 轮询页面标题确认加载到位（最多60s）
        ready = False
        for _ in range(60):
            title = tab.evaluate("document.title", quiet=True)
            if title and stem in title:
                ready = True
                break
            time.sleep(1)
        if not ready:
            print("   !! 页面标题未就绪（检查 Jupyter 服务是否假死：重启 jupyter-notebook）")
            continue
        wait_stable(tab)
        time.sleep(1)

        # 关闭 Notebook7 迁移横幅（幂等）
        if tab.evaluate(DISMISS_BANNER_JS) == "clicked":
            print("   迁移提示条已关闭")
            time.sleep(1)
            wait_stable(tab)

        info = tab.evaluate(SCROLLER_JS) or {"sh": 0, "ch": 0, "id": "?"}
        max_scroll = info["sh"] - info["ch"]
        print(f"   滚动容器: {info['id']} | 内容高度 {info['sh']} | 视口 {info['ch']} | 需滚动 {max_scroll}")

        bring_front(tab)
        time.sleep(1.0)
        bounds = chrome_window_bounds(stem)
        if not bounds:
            print("   !! 找不到我们的 Chrome 窗口（检查是否被用户窗口遮挡/最小化）")
            continue

        tab.evaluate("(window.__scroller||document.scrollingElement).scrollTop = 0")
        time.sleep(0.6)
        bring_front(tab)
        time.sleep(0.5)
        capture(os.path.join(args.out_dir, f"{stem}_1.png"), bounds)
        print(f"   截图 {stem}_1.png")
        if max_scroll > 60:
            tab.evaluate(
                f"(window.__scroller||document.scrollingElement).scrollTop = {int(max_scroll)}")
            time.sleep(0.6)
            bring_front(tab)
            time.sleep(0.5)
            capture(os.path.join(args.out_dir, f"{stem}_2.png"), bounds)
            print(f"   截图 {stem}_2.png (内容超出一屏)")

    tab.close_tab()
    print("全部截图完成（务必用 Read/看图工具逐张目检后再使用）")


if __name__ == "__main__":
    main()
