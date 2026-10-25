---
name: lab-report-completion
description: 完成《数据科学与分析》等课程的 Python 实验报告全流程技能：解析实验指导书 PDF 与报告模板 .doc（含从 doc 二进制中提取代码截图并逐字转录）、只读检查数据文件、构建版本兼容的 pandas/Jupyter 运行环境并用 Chrome+CDP+系统级 screencapture 生成真机运行截图、整理代码源文件交付包、最后用本地 Office 编辑引擎（editor_sdk/edsdk）在原版 Word 的副本上原位插入结果截图与实验小结。当用户给出实验指导书、数据文件（xlsx/csv）与报告模板，并要求"完成实验报告、贴真实截图（不是渲染/画的）、代码逐字按截图、源码文件保存供亲自运行、原文件禁止更改"时使用本技能。
agent_created: true
---

# 实验报告完成全流程（lab-report-completion）

## 0. 用户核心诉求（不可违背的铁律）

1. **真实截图，严禁编造**：报告中的"结果截图"必须是真机上真实运行代码的屏幕截图（系统级 screencapture 截取真实浏览器窗口），绝不允许用 HTML 渲染、画图、PIL 合成等方式伪造输出。
2. **代码逐字一致**：报告/模板中以截图形式给出的代码，必须逐字转录执行——包括分号、空格、怪异变量名（如 `'Zhio'`）、非标准写法（如 `str.split('.',1,False)` 位置参数）。不得"修正"代码。
3. **原文件与数据禁止更改**：指导书、报告模板原始文件、数据文件（xlsx/csv）只读。所有操作基于副本；交付前用 md5 校验数据副本与原件一致。
4. **代码源文件按题保存在文件夹**：交付 `.py`（供终端运行）与 `.ipynb`（与报告截图逐字一致，供 Jupyter 复现），数据文件副本一并放入，README 说明运行方式。
5. **Word 必须在原版副本上原位填充**：复制原版报告 → 在副本上插入截图与小结 → 保留原有全部标题、字体、版式、原代码截图。**禁止用 python-docx 从头重建报告**（会丢失原模板的样式结构，已被用户否决过一次）。
6. **先盘点既有环境，再考虑装包**（实验3教训）：做过后继实验时，前序实验几乎必然已建好课程专用环境。重复安装既浪费时间，还会把工具库误装进错误的 Python（系统 vs venv），污染两边的依赖状态。

## 0.5 阶段〇：环境盘点（先于一切 pip install）

**触发条件**：只要课程文件夹里存在任何前序实验（实验1/实验2…），必须先完成本阶段，禁止直接 `pip install`。

### 0.5.1 三步找既有环境

```bash
# ① 找启动脚本（学生双击用的 .command 通常写死了 venv 解释器路径）
cat "/课程根目录/启动Jupyter.command" 2>/dev/null
# ② 找既有 venv（注意深度，runenv 可能藏在 工作目录/ 下一层）
find "/课程根目录" -maxdepth 4 -name "pyvenv.cfg" 2>/dev/null
# ③ 核对该环境的关键包是否齐全
"<venv>/bin/python" --version
"<venv>/bin/pip" list | grep -iE "pandas|numpy|notebook|ipykernel|matplotlib|openpyxl"
```

实例（实验3）：`启动Jupyter.command` 指向 `实验2：大数据处理-学生用/工作目录/runenv/bin/jupyter-notebook`，即实验2 已建好 Python 3.9.6 + pandas 1.4.4 + notebook 6.5.7 的专用环境——实验3 直接复用，仅补装缺失的 matplotlib，零重复建设。

### 0.5.2 环境分层原则

| 层 | 用途 | 装在哪 |
|---|---|---|
| **课程运行环境**（venv） | 跑课程代码、执行笔记本、Jupyter 服务 | pandas/numpy/notebook/ipykernel/matplotlib/openpyxl——版本必须满足课程代码的旧 API |
| **工具解析环境**（系统 python3） | 解析 PDF（pymupdf）、雕刻 .doc 图片（olefile）、截图管线（websocket-client、pyobjc）、pandoc 校验 | 与课程环境无关的一次性工具，装系统 Python 即可，**不要**装进课程 venv |

判断口诀：**跑报告代码的 → venv；帮你读文件/截屏的 → 系统**。检查依赖时先想清楚当前是哪一层，用对应解释器 `python3 -c "import X"` 探测，不要拿系统 python3 去探测课程包（必然误判为缺失）。

### 0.5.3 确实需要新建环境时

才按 §2 的环境配方走后台安装。若新实验与最近前序实验的课程代码同源（同一本教材），优先复制/复用前序 venv，只补差量包。

## 1. 阶段一：解析输入材料

### 1.1 实验指导书 PDF

`Read` 工具可能报 "Cannot display content of binary file"。用 pymupdf 提取：

```python
import pymupdf  # pip install pymupdf
doc = pymupdf.open('实验指导书.pdf')
text = ''.join(f'\n===== 第{i+1}页 =====\n' + p.get_text() for i, p in enumerate(doc))
# 同时渲染每页为 PNG，用于查看指导书中的结果截图（核对预期输出样式）
for i, page in enumerate(doc):
    page.get_pixmap(dpi=150).save(f'指导书页面/page_{i+1:02d}.png')
```

指导书中的输出截图是核对真实运行结果是否正确的基准（如去重后剩几行、切片结果样式）。

### 1.2 报告模板 .doc（关键：拿到代码截图）

`.doc` 是 OLE 二进制格式，图片无法用 python-docx 读取。两步走：

```bash
# ① 文本结构（textutil 保留字体/字号/加粗/对齐等格式信息，但会丢弃全部图片）
textutil -convert docx 原模板.doc -output 模板.docx
textutil -convert txt  原模板.doc -output 模板.txt
unzip -o 模板.docx -d 模板解包   # 解析 word/document.xml 可得到段落顺序与格式细节
```

```python
# ② 从 .doc 二进制按签名雕刻提取全部图片（代码截图就在其中）
import olefile, re
ole = olefile.OleFileIO('原模板.doc')   # pip install olefile
data = ole.openstream(['Data']).read()  # 图片通常在 Data 流
n = 0
for m in re.finditer(rb'\x89PNG\r\n\x1a\n', data):
    end = data.find(b'IEND', m.start()) + 8
    n += 1
    open(f'img_{n:02d}.png','wb').write(data[m.start():end])
# JPEG 同理: \xff\xd8\xff ... \xff\xd9
```

提取出的 PNG 按 Data 流偏移顺序 = 文档中出现顺序。逐张查看（Read 工具），**逐字转录每张代码截图**，并按文本结构的空段数量与顺序把图片映射到"（x.x）输入以下命令"和"结果截图"标签之间。易错点：textutil 转出的 txt/docx 中图片位置表现为空段落，空段数量 ≈ 图片数，但需结合语义校对。

模板尾部若出现"实 验 报 告/姓名/学号"等封面文字，属正常——它实际存放在**页眉**（header2.xml）中，textutil 将其线性化到文末。填充时不要动它。

模板尾部若出现"实 验 报 告/姓名/学号"等封面文字，属正常——它实际存放在**页眉**（header2.xml）中，textutil 将其线性化到文末。填充时不要动它。

### 1.3 数据文件（只读）

用 pandas 读取全部数据文件，记录：形状、列名、每列 dtype（含 NaN 的数值列会变 float64）、NaN 位置、重复行位置、sheet 列表。这些决定后续输出是否与指导书一致。

### 1.4 模板类型 B：教材案例 + 扩展题（实验3 形态）

模板可能不止"指导书案例"一种题源。实验3 的模板每类图表有**两道题**：① 教材上机案例（代码 = 模板中的代码截图，数据 = 附带的 csv/xlsx）；② **扩展题**（题目文字 + 数据表格直接印在模板里，代码截图也是模板给定的、带 `rcParams` 中文字体设置的完整代码，数据内嵌在代码中）。

识别特征：文本结构里每类图表出现两段"输入以下命令，并显示结果"+"结果截图"，中间夹着数据表格文字。此时**代码截图总数会大于案例代码段数**，雕刻出的 PNG 必须逐张核对内容归属，映射以模板语义结构为准（Data 流偏移顺序可能有局部乱序，如 img_07 排在 img_08 后面，别盲信）。对应到 Word 填充：多个同名"结果截图"标签，fill_report.py 的 sections 按文档顺序依次消耗即可。扩展题表格数据转录后必须与代码截图中的数据逐值核对（如 `scores = [82,91,78,86,80]` ↔ 表格"语文82/数学91/…"），防止转录错位。

## 2. 阶段二：干跑验证（先于一切截图）

**把报告中全部代码逐字拼成一个脚本干跑**，任何报错/警告都在此阶段消灭：

```python
import matplotlib          # ① 先 import 所有三方库（matplotlib/pandas/numpy...）
matplotlib.use('Agg')
import warnings            # ② 库导入完成后才开严格模式
warnings.simplefilter('error', Warning)
warnings.filterwarnings('ignore', message='FigureCanvasAgg is non-interactive',
                        category=UserWarning)  # ③ Agg 无头后端对 plt.show() 的假警报，仅此一条放行
# ...逐字执行报告代码...
```

**时序铁律**：`simplefilter('error')` 若放在 `import matplotlib` 之前，三方库内部的弃用警告（如 pyparsing 的 `PyparsingDeprecationWarning: 'oneOf' deprecated`）会被升级为错误，干跑直接崩——这不是课程代码的问题，别误判。同理，**同一题的多个 cell 要共享同一个 exec 命名空间**（cell2 的 `gb` 依赖 cell1 的 `df`），不同题之间命名空间隔离（模拟独立笔记本）。

坑与对策（均实际踩过）：

| 现象 | 根因 | 正确处理 |
|---|---|---|
| `fillna(method='pad')` 报 TypeError | pandas 3.x 已移除 method 参数 | 用 pandas 1.x/2.x 环境运行（代码不许改） |
| `str.split('.',1,False)` FutureWarning（Jupyter 中显示为粉色框） | pandas 1.5 起位置参数将变 keyword-only | 用 pandas ≤1.4.4 |
| numpy ≥1.25 时 fillna 内部 DeprecationWarning | pandas 调用 np.find_common_type | ipykernel 默认过滤器不显示（非 `__main__` 触发），实测 nbconvert 输出干净；若仍出现则降 numpy 到 1.24.x 或在 kernel 配置过滤 |
| 综合题(3)(4) 报 `KeyError: 'Revenue'` | 这些题依赖题(1)创建的 Revenue 列 | 笔记本必须包含前置单元格（数据 cell + Revenue cell），按顺序执行 |
| `read_csv(..., names=['band','area','num'])` 数据却是4列 | 文件首列序号被 pandas 自动当作索引 | 这是预期行为（与指导书输出一致），不要改代码 |
| import matplotlib 即报 `PyparsingDeprecationWarning` 被升为错误 | matplotlib 3.9 + pyparsing 3.2 在 import 期发库内部弃用警告 | `pip install "pyparsing==3.1.4"` 降级根治；**不降级会污染 Jupyter 截图**（首次 import 的 cell 出现粉色警告框） |
| 干跑 plt.show() 报 `FigureCanvasAgg is non-interactive` | Agg 无头后端没有 GUI | 假警报，filterwarnings 放行该条即可；Jupyter inline 后端无此问题 |
| 终端跑 .py 卡住直到超时被 SIGTERM | `MPLBACKEND=Agg &&` 少了 export，环境变量没生效，plt.show() 打开 GUI 窗口阻塞 | 必须 `export MPLBACKEND=Agg`；或每条命令前缀 `MPLBACKEND=Agg python3 ...` |
| `timeout 60 python ...` 报 exit 127 | macOS 默认无 GNU coreutils 的 timeout | 去掉 timeout 直接跑（Agg 已 export 时不会阻塞），或 `brew install coreutils` 用 gtimeout |

**环境配方（已验证）**：Python 3.9 + venv + `numpy==1.26.4` + `pandas==1.4.4` + `matplotlib==3.9.4` + `pyparsing==3.1.4` + `openpyxl` + `notebook==6.5.7` + `ipykernel`。安装耗时较长，用后台任务跑，给足超时（pip 每次装包 1-3 分钟）。安装完 matplotlib 后务必先干跑再截图，把 import 期/运行期警告全部清零。

### 2.1 中文字体（macOS 必做，否则截图中文全是方框）

课程代码通常写死 `SimHei`（Windows 黑体）。macOS 无自带，matplotlib 找不到字体会：① 图内中文渲染为方框（不可交付）；② 每次 `rcParams` 触发 `findfont` 警告污染输出。

```bash
# ① 检查是否已装（此前实验可能已装过——先查再下，同 §0.5 原则）
find ~/Library/Fonts /Library/Fonts -iname "simhei*" 2>/dev/null
# ② 未装则下载到用户字体目录（公开镜像，9.7MB）
curl -sL --max-time 60 -o ~/Library/Fonts/SimHei.ttf \
  "https://github.com/StellarCN/scp_zh/raw/master/fonts/SimHei.ttf"
file ~/Library/Fonts/SimHei.ttf   # 应输出 TrueType Font data
```

```python
# ③ 清 matplotlib 字体缓存并强制重建（venv 的 python 执行）
import glob, os, matplotlib, matplotlib.font_manager as fm
for f in glob.glob(os.path.join(matplotlib.get_cachedir(), 'fontlist*.json')):
    os.remove(f)
import importlib; importlib.reload(fm)   # 重新扫描
print(fm.findfont('SimHei'))             # 应输出 /Users/xxx/Library/Fonts/SimHei.ttf
```

**缓存时序坑**：在同一个 Python 进程里"清缓存 + `_load_fontmanager(try_read_cache=False)`"可能仍报 SimHei 不存在（构建时序问题）；换一个**全新的 Python 进程**重新探测即正常。验证以 `findfont('SimHei')` 返回正确路径为准。

## 3. 阶段三：构建笔记本并真实执行

1. 每道题一个 `.ipynb`，**每个代码截图 = 一个 code cell**，单元格内容逐字一致（裸表达式 `df`、`newDF` 保持裸写，Jupyter 会显示 Out[n]）。综合题笔记本需含前置 cell。
   - 实践验证：每道题开头加一个 **markdown 标题 cell**（`# 饼图` 等），截图里能看出题号，报告更易读。
   - 实践验证：由"干跑.py 的 SECTIONS 字典"统一派生 .ipynb（build_notebooks.py）与 .py（build_py.py，裸表达式改 print()），保证两份交付与干跑验证三者同源一致，不手抄第二遍。
2. 执行（较慢，每个 10-25 秒，务必给长超时或逐个跑，否则会被工具默认 120s 超时 SIGTERM）：

```bash
jupyter nbconvert --to notebook --execute --inplace q01_xxx.ipynb
```

3. 校验执行结果：读 JSON 检查 `outputs` 中无 `output_type=="error"`、每 cell 有 `execution_count`。（nbformat 的 `MissingIDFieldWarning` 是构建元数据警告，无害可忽略。）
4. **信任笔记本**：`jupyter trust q*.ipynb`——否则 Jupyter 界面显示 "Not Trusted" 且可能抑制 HTML 输出。

## 4. 阶段四：真机截图管线（核心）

原理：真实执行过的笔记本 → 本地 Jupyter 服务 → 真实 Chrome 窗口加载 → CDP 控制导航/滚动/置前 → Quartz 定位窗口 → 系统级 `screencapture -R` 截取窗口区域。

### 4.1 启动服务

```bash
# Jupyter（空 token 方便直达 URL）
jupyter-notebook --no-browser --port=8899 --NotebookApp.token='' --NotebookApp.disable_check_xsrf=True
# Chrome（专用 profile；参数缺一不可，原因见下）
"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" --no-sandbox --test-type \
  --remote-debugging-port=9222 --remote-allow-origins='*' \
  --user-data-dir=/abs/path/chrome配置 --no-first-run --no-default-browser-check \
  --hide-crash-restore-bubble --window-size=1400,845 --window-position=30,25 about:blank
# 后台运行，日志重定向到文件（不要管道接 head，会 SIGPIPE 杀掉 Chrome）
```

- `--no-sandbox`：Bash 沙箱会阻止 Chrome 自带沙箱初始化（"sandbox initialization failed"）导致崩溃。若可用 `dangerouslyDisableSandbox` 启动则可去掉此参数（无提示条更美观），失败再回退。
- `--test-type`：隐藏"不受支持的命令行旗标 --no-sandbox"提示条。
- `--remote-allow-origins=*`：没有它 CDP WebSocket 握手返回 403 Forbidden。
- 窗口高度 ≤845：避免底部被 Dock 遮挡。
- 屏幕为 Retina 时 `screencapture` 输出 2x 像素图（逻辑 1470×956 → 2940×1912 像素）。

### 4.2 CDP 连接的四个坑

```python
import urllib.request, json, urllib.parse
from websocket import create_connection  # pip install websocket-client
opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))  # 坑①绕过代理
# 坑②: IPv4 127.0.0.1:9222 可能被环境占用返回404 → 用 http://[::1]:9222
# 坑③: 每个笔记本 PUT /json/new?<urlencode的url> 新建 tab，比复用旧 tab 稳
url = "http://localhost:8899/notebooks/q01_xxx.ipynb"
req = urllib.request.Request("http://[::1]:9222/json/new?" + urllib.parse.quote(url, safe=":/?=&.-"), method="PUT")
tab = json.loads(opener.open(req, timeout=10).read())
ws = create_connection(tab['webSocketDebuggerUrl'], timeout=60)  # 坑④需 --remote-allow-origins=*
```

- 坑①：环境注入 `HTTP_PROXY=127.0.0.1:<port>`，所有 localhost 请求会被代理劫持（表现为 404/"upstream connect failed"），必须 `ProxyHandler({})` / `NO_PROXY='*'` / `curl --noproxy '*'`。
- send 循环里必须自动应答 JS 对话框（Jupyter 关 tab 会触发 beforeunload 原生确认框，卡死后续操作）：收到 `Page.javascriptDialogOpening` 立即回 `Page.handleJavaScriptDialog {accept:true}`；关 tab 前先 `Page.navigate` 到 about:blank，再 `/json/close/<id>`。

### 4.3 渲染等待与滚动（大坑）

- **Jupyter classic 的滚动容器是 `#site`，不是 documentElement**——`window.scrollTo` 完全无效，且 `document.body.scrollHeight` 恒等于视口高度（表现为"所有页面都正好一屏"）。必须探测：

```js
(function(){
  var b = document.scrollingElement;
  if (b && b.scrollHeight > b.clientHeight + 30) return {el:b, sh:b.scrollHeight, ch:b.clientHeight};
  var best=null, bestOver=0;
  document.querySelectorAll('body,div,section,main').forEach(function(e){
    var ov = getComputedStyle(e).overflowY;
    if ((ov==='auto'||ov==='scroll') && e.scrollHeight > e.clientHeight){
      if (e.scrollHeight - e.clientHeight > bestOver){ bestOver = e.scrollHeight - e.clientHeight; best=e; }
    }
  });
  if (best) return {el:best, sh:best.scrollHeight, ch:best.clientHeight};
  return {el:b, sh:b?b.scrollHeight:0, ch:b?b.clientHeight:0};
})()
```

把返回的 el 存到 `window.__scroller`，之后用 `__scroller.scrollTop = N` 滚动。

- 等待渲染稳定：轮询（每 0.5s）直到 `scrollHeight` 连续 3 次不变且 `document.querySelectorAll('.cell').length > 0`，再追加 1.5s。首次打开点掉 Notebook7 迁移横幅（找文本含 "Don't show anymore" 的 button 并 click，持久化到 localStorage）。
- 内容超高（max_scroll = sh - ch > 60px）时截两屏：顶部一张 + 滚动到底再一张；报告中两图上下排列。
- **长笔记本两屏会漏中间内容**（实验3 实测）：内容高度 > 2×视口时（如折线图 5 cell、高 3270px vs 视口 647px），首屏+尾屏之间会跳过中间 cell 的输出（函数定义输出、中间图表）。补截方法：打开同一笔记本，滚动到 `mid = (sh - ch) // 2` 处再截一张，存为 `_3.png`。用 `scripts/capture_middle_screens.py`。**plan.json 中的图片顺序按内容排列**：capture 先截顶(_1)后截底(_2)、补截为中(_3)，故 3 屏题在 plan 里应写 `_1, _3, _2`（顶部→中间→底部）。

### 4.4 窗口定位与截屏（最容易截错的一步）

```python
# 坑：用户自己的 Chrome 可能正在运行！按应用名找窗口会截到用户的窗口。
# 必须按【进程PID + 窗口标题】双重过滤。
import subprocess, Quartz  # pip install pyobjc-framework-Quartz
our_pids = set(subprocess.run(['pgrep','-f','remote-debugging-port=9222'],
               capture_output=True,text=True).stdout.split())
wl = Quartz.CGWindowListCopyWindowInfo(Quartz.kCGWindowListOptionOnScreenOnly, Quartz.kCGNullWindowID)
for w in wl:
    if (w.get('kCGWindowOwnerName')=='Google Chrome'
        and str(w.get('kCGWindowOwnerPID')) in our_pids
        and notebook_stem in (w.get('kCGWindowName') or '')
        and 'Jupyter' in (w.get('kCGWindowName') or '')):
        b = w['kCGWindowBounds']; bounds = (b['X'],b['Y'],b['Width'],b['Height'])
# 截屏前把 tab 置前（同时抬高窗口；不要用 open -a 激活——那会把用户的 Chrome 调到前台）
send("Page.bringToFront"); time.sleep(1)
subprocess.run(['screencapture','-x','-R',f'{x},{y},{w},{h}', out_png], check=True)
```

**每张截图必须用 Read 工具目检**：确认是我们的窗口、内容渲染完整、无对话框遮挡、无用户窗口混入。

### 4.5 Chrome 对话框三连坑（截屏杀手）

| 现象 | 根因 | 处理 |
|---|---|---|
| 「Google Chrome」未預期的結束 对话框盖住页面 | SIGTERM/SIGKILL 杀 Chrome 被记为崩溃，macOS UserNotificationCenter 弹窗 | `pkill -x UserNotificationCenter` 关掉弹窗；启动后用 Quartz 预检无此进程窗口 |
| "恢复标签页？"气泡 | 同上，Chrome 认为上次异常退出 | 彻底方案：CDP `Browser.close` 干净退出 → 等 pgrep 确认进程消失 → 改 `chrome配置/Default/Preferences` 的 `profile.exit_type="Normal"`、`exited_cleanly=true` → 删除 `Default/Sessions/` 目录 → 重启 |
| 启动即崩溃退出 | Bash 沙箱阻止 Chrome 沙箱初始化 / 输出管道 head 触发 SIGPIPE | `--no-sandbox`；日志重定向文件；必要时 dangerouslyDisableSandbox |

启动后先做**对话框预检**：Quartz 扫描 `UserNotificationCenter` 窗口和本 PID 的 <600×400、layer=8 的小窗，有则先处理再截图。

### 4.6 收尾

截完用 `Browser.close` 退出（可能因挂起对话框失败，届时 SIGKILL → 清理 Sessions/Preferences → 重启再验证）；Jupyter 服务 pkill 停止；临时 Chrome profile 删除。

## 5. 阶段五：代码源文件交付包

```
实验N代码/
├── README.md               # 环境要求/运行方式/题目-文件对应表/版本注意事项
├── py源代码/               # N个 .py：代码逐字一致，仅把裸表达式改为 print() 以便终端查看；数据副本同目录
├── jupyter笔记本/          # N个 .ipynb（与报告截图逐字一致，含已执行输出）+ 数据文件副本
├── 运行结果截图/           # 截图按"题号-中文题名-截图N.png"重命名
└── (数据文件副本)          # md5 必须与原件一致
```

每个 `.py` 运行一遍确认成功；README 中写明 pandas 版本注意事项（如 `fillna(method=)` 在 3.x 报错，替代写法 `ffill()/bfill()`；`str.split` 位置参数需关键字化）。

**放入课程文件夹的惯例**（实验2/3 均验证）：交付包整体复制到 `实验N：xxx-学生用/实验N代码/`，数据文件副本同时放包根部一份（`实验2代码/rz2.xlsx` 同款结构）；完成版 Word 也放课程文件夹根部（`028125435-xxx-实验报告N-完成版.docx`）。

**复制防嵌套坑**：`cp -R 源 目标` 若目标路径已存在（哪怕只是之前误操作残留），会把源目录整个复制成目标的**子目录**，产生 `实验N代码/交付/…` 嵌套副本。复制前先确认目标不存在，复制后 `ls` 检查一层结构。交付前顺手清理 `.ipynb_checkpoints/`、`.DS_Store`。

## 6. 阶段六：Word 原位填充（editor_sdk / edsdk）

### 6.1 强制流程

1. 复制原版：`cp 原模板.doc 完成版.doc`（原版永远不动）。
2. 调 `Skill(tencent-local-office-edit)`，**先读其 doc.md**，然后对本次要用的每个工具先 `python3 edsdk.py schema <工具名>` 再调用（违反会参数错误）。edsdk 位置：WorkBuddy 内置插件目录 `skills/tencent-local-office-edit/edsdk.py`；editor_sdk 服务默认探测端口 39099 起。
3. `present_files` 呈现副本 → `get_pool_status` 拿 file_id；若未注册则 `open_file file_path=... open_with_existing=true`（返回的 file_id = 路径字符串本身）。
4. 编辑 → `save_file` → 三重验证（见 6.5）。

### 6.2 定位插入点（最大的坑）

- `doc_find("结果截图")` 返回**逆序**（文档尾部在前）的全部匹配，且匹配范围只是匹配文本本身，每个匹配带 begin/end 与 paragraph_id。
- **大坑**：匹配文本若是段落前缀（如"五、实验小结"只是完整标题"五、实验小结（可以包括但不限于…）"的前6字），`match.end + 1` 仍在标题段内部——markdown 会插进标题中间，把标题撕成两半（前半段+插入内容+后半段落在文档尾部）。
- **正确做法**：用 `doc_resolve_document_structure(idx=匹配位置, mode=full, text_preview_length=50)` 拿到所在段落的 `start_index/end_index`（即文本末尾，段落分隔符占1个坐标单位），插入点 = `end_index + 1`（跳过分隔符，落到下一空段起点）。
- 空段落表现为零宽节点（start==end），每个段落分隔符占 1 个 UTF-16 坐标单位。

### 6.3 插入截图与文本

```bash
# 图片：idx=DOC坐标, image_path=绝对路径, w/h 单位是 96DPI 像素（14.5cm ≈ 548px；2800×1690 的 2x 截图 → 548×331）
python3 edsdk.py call doc_insert_image --json '{"file_id":"...","idx":1347,"image_path":"/abs/截图.png","w":548,"h":331}'
```

- **从文档后部往前处理**：前面的编辑只影响更大偏移，预计算的坐标保持有效；同一段内多次插入都写同一个 idx 即可。
- 同一位置放两张图（上下排列）：在**同一个 idx** 按「img2 → doc_insert_paragraph → img1」逆序堆叠，结果为 [img1¶][img2¶]，各占一段。
- 长文本（实验小结）用 `doc_insert_markdown`，markdown 参数支持 `file:///abs/小结.md` 直接读文件；每个段落自动带段落分隔符，插入后各自成段（继承所在处正文样式）。
- 撤销失误：`doc_list_recent_ai_edits`（按 tool_name 过滤）找到 version → `doc_undo_ai_edit(target_version=version)`（禁止跳过 list 直接 undo；OT 回滚不影响后续无关编辑）。

### 6.4 保存与格式限制

- **编辑内核只能写 OOXML**：对 97-2003 `.doc` 保存会被拒绝，必须 `save_file` 加 `allow_format_change=true` 另存为 `.docx`。内容与排版 100% 保留，仅容器格式变化——向用户说明这一点（原版 .doc 仍未动，副本可删）。
- `doc_get_images` 只列出 drawing 类图片：老 .doc 的 pict 图片（如模板原有代码截图）**不会出现在列表里但实际保留**，以 `unzip -l 完成版.docx | grep media` 计数为准。

### 6.5 交付前验证（三重）

1. **XML 校验**：解包 docx，正则提取段落序列（文本+图片数），确认每个"结果截图"段后紧跟 1-2 个图片段、标题段完整无拼接、小结段落独立、封面文字在 header*.xml 中；
2. **数量校验**：media 中图片数 = 原代码截图数 + 新插入截图数（+原模板杂项图）；
3. **视觉校验**：`pandoc 完成版.docx -o 验证.html --extract-media=媒体目录 --standalone`（pandoc 是少数能保留 docx 图片的转换器；textutil 会丢图），在 Chrome 中按 4.4 的方法截 2-3 屏目检。
4. 数据文件 md5 与原件比对。

## 7. 易错点速查表

| # | 错误现象 | 根因 | 正确处理 |
|---|---|---|---|
| 1 | Read 读 PDF 报二进制错误 | PDF 非文本 | pymupdf 提文本+渲染页面图 |
| 2 | textutil 后图片全丢 | textutil 不保留 doc 图片 | olefile 按签名从 Data 流雕刻图片 |
| 3 | pandas 3.x 跑课程代码报 TypeError | fillna(method=)/str.split 位置参数已移除 | 换 pandas 1.x/2.x 环境，代码不改 |
| 4 | Jupyter 输出出现粉色 FutureWarning | pandas 1.5+ 对 str.split 位置参数警告 | 用 pandas 1.4.4 |
| 5 | 笔记本执行中途被 SIGTERM | nbconvert 单个 10-25s，循环超默认 120s 超时 | 后台执行/显式长超时/逐个跑 |
| 6 | 笔记本显示 Not Trusted | 未签名 | `jupyter trust *.ipynb` |
| 7 | curl/urllib 访问 localhost 返回 404 或 upstream failed | 环境注入 HTTP_PROXY 劫持 localhost | ProxyHandler({}) / NO_PROXY / --noproxy '*' |
| 8 | CDP WebSocket 403 Forbidden | 未允许 origins | --remote-allow-origins='*' |
| 9 | 9222 端口 IPv4 连不上/404 | IPv4 被环境其他进程占用 | 用 http://[::1]:9222 |
| 10 | Chrome 启动即退出 | 沙箱初始化失败 / 输出管道 head 触发 SIGPIPE | --no-sandbox；日志重定向文件；必要时 dangerouslyDisableSandbox |
| 11 | 截图内容是用户自己的浏览器窗口 | Quartz 按应用名取窗口 | 按 PID+标题过滤（pgrep remote-debugging-port） |
| 12 | 页面"高度恒等于视口"且滚动无效 | Jupyter classic 滚动容器是 #site | 探测 overflow 滚动容器并对它 scrollTop |
| 13 | 截图有崩溃对话框/恢复气泡 | 此前 SIGTERM/SIGKILL 杀 Chrome | pkill -x UserNotificationCenter；清 Sessions/ 改 Preferences；尽量 Browser.close 干净退出 |
| 14 | 关不掉 Jupyter 标签页 | beforeunload 原生确认框阻塞 | CDP send 循环自动应答 javascriptDialogOpening；先导航 about:blank 再 close |
| 15 | markdown 把标题撕成两半 | doc_find 匹配只是段落前缀，end+1 仍在段内 | 用结构树段落边界定位；失误用 list_recent_ai_edits+undo_ai_edit 撤销 |
| 16 | .doc 保存被拒 | 内核只写 OOXML | save_file allow_format_change=true 另存 .docx |
| 17 | 以为封面丢了 | 封面信息块在页眉 header2.xml | 检查各 XML 部件，勿慌 |
| 18 | pip/长命令中途被杀 | 工具默认 120s 超时 | 后台执行 + 显式长超时 |
| 19 | 数据文件被误改 | 直接在原目录跑代码 | 一律副本操作 + md5 终检 |
| 20 | 报告被重建成"格式不对" | 用 python-docx 从头写 | **必须复制原版 .doc 原位填充** |
| 21 | 给前序实验装过的包又装一遍（pymupdf/olefile 等） | 未盘点既有环境，拿系统 python3 探测课程包 | 走 §0.5 环境盘点：找 .command/venv → 分层探测 → 只补差量 |
| 22 | 干跑 import 阶段就崩 PyparsingDeprecationWarning | simplefilter('error') 放在 import matplotlib 之前 + pyparsing 3.2 与 matplotlib 3.9 不兼容 | 警告开关放在三方 import 之后；`pip install "pyparsing==3.1.4"` 根治（否则污染 Jupyter 截图） |
| 23 | 截图里中文全是方框 | macOS 无 SimHei，matplotlib findfont 失败 | §2.1：下载 SimHei.ttf 到 ~/Library/Fonts + 清 fontlist 缓存重建 + 新进程验证 |
| 24 | 终端跑 .py 卡死到超时 | `MPLBACKEND=Agg &&` 未 export，plt.show() 开 GUI 阻塞 | `export MPLBACKEND=Agg` 或命令前缀写法 |
| 25 | `timeout 60 …` 报 exit 127 | macOS 无 GNU timeout | 去掉（Agg 下不阻塞）或 brew coreutils 用 gtimeout |
| 26 | 长笔记本截图只有首尾，中间 cell 输出缺失 | 内容高度 > 2×视口，两屏方案跳过中段 | capture_middle_screens.py 补中间屏；plan.json 按 `_1,_3,_2` 内容顺序排 |
| 27 | 交付包出现 `实验N代码/交付/…` 嵌套目录 | cp -R 时目标已存在，源整个变成子目录 | 复制前确认目标不存在；复制后 ls 检查；清 .ipynb_checkpoints |

## 8. 可复用脚本（scripts/）

四个脚本来自真实实验全流程实测并已参数化，可直接使用（也可按前文原理重写）：

| 脚本 | 作用 | 核心参数 |
|---|---|---|
| `capture_screenshots.py` | 真机截图管线（单tab持久导航版，自动关横幅、超一屏自动两连拍） | `--notebooks-dir` `--out-dir` `--notebooks` `--base-url` `--img-w/--img-h` |
| `capture_middle_screens.py` | 长笔记本补截中间屏（滚动到中点，存 `_3.png`；依赖主截图管线已跑过） | `--notebooks-dir` `--out-dir` `--notebooks`（缺省自动选内容高>2×视口的） `--base-url` |
| `update_docx_images.py` | 按 doc_get_images 文档顺序替换/删除 Word 图片 | `--replace N=图片路径` `--delete N` `--dry-run` `--out` |
| `fill_report.py` | 原版 Word 副本原位填充截图与小结（JSON 计划驱动） | `plan.json`（`--print-plan-template` 查看结构） |

### capture_screenshots.py

前置：jupyter-notebook 与 Chrome 已按 §4.1 启动；依赖 `pip install websocket-client pyobjc-framework-Quartz`。

```bash
python3 scripts/capture_screenshots.py \
    --notebooks-dir /abs/实验N代码/jupyter笔记本 \
    --out-dir /abs/工作目录/截图
```

### capture_middle_screens.py

在 capture_screenshots.py 跑完后使用（同一 Jupyter + Chrome 服务仍在）。自动探测内容高 > 2×视口的笔记本，滚动到中点补截 `_3.png`；也可 `--notebooks` 手动指定：

```bash
python3 scripts/capture_middle_screens.py \
    --notebooks-dir /abs/实验N代码/jupyter笔记本 \
    --out-dir /abs/工作目录/截图
```

注意：补截图在 plan.json 中按内容顺序排列为 `_1, _3, _2`（顶部→中间→底部）。

### update_docx_images.py

先 `--dry-run` 列出图片序号，确认后再替换/删除（序号按操作前列表计，删除自动从大到小执行）：

```bash
python3 scripts/update_docx_images.py /abs/完成版.docx --dry-run
python3 scripts/update_docx_images.py /abs/完成版.docx \
    --replace 5=/abs/截图/新图1.png --replace 6=/abs/截图/新图2.png --delete 17
```

### fill_report.py

```bash
python3 scripts/fill_report.py --print-plan-template   # 查看计划文件结构
python3 scripts/fill_report.py plan.json
```

plan.json 字段：`template_doc`（原版模板，只读）/ `work_doc`（工作副本，可选，缺省为模板旁“-完成版.doc”）/ `out_docx` / `img_w,img_h`（548×331≈14.5cm）/ `sections[]`（按文档顺序；`find` 标签须为独立整段，混排段落自动跳过）/ `summary{find,markdown}`。脚本自动处理：模板复制、段落结构树定位（防前缀标签撕标题）、.doc→.docx 保存、文档异步打开就绪轮询、edsdk 路径变更探测。

## 9. 交付前检查清单

- [ ] 报告中每个"（x.x）输入以下命令"的代码与原模板截图逐字一致
- [ ] 每个题的运行截图为真机截图且经目检（无对话框、无用户窗口、无 Not Trusted）
- [ ] 结果数值与指导书输出核对一致
- [ ] 原版 .doc、指导书 PDF、全部数据文件 md5 未变
- [ ] `实验N代码/` 中每个 .py 亲自运行成功；.ipynb 在 Jupyter 打开可复现
- [ ] README 含运行说明与 pandas 版本注意事项
- [ ] 完成版 Word 原版结构完好 + 截图位置正确 + 小结已填 + 封面页眉未动
