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

### 1.3 数据文件（只读）

用 pandas 读取全部数据文件，记录：形状、列名、每列 dtype（含 NaN 的数值列会变 float64）、NaN 位置、重复行位置、sheet 列表。这些决定后续输出是否与指导书一致。

## 2. 阶段二：干跑验证（先于一切截图）

**把报告中全部代码逐字拼成一个脚本干跑**，任何报错/警告都在此阶段消灭：

```python
import warnings
warnings.simplefilter('error', Warning)  # 全部警告升级为错误，最严格
# ...逐字执行报告代码...
```

坑与对策（均实际踩过）：

| 现象 | 根因 | 正确处理 |
|---|---|---|
| `fillna(method='pad')` 报 TypeError | pandas 3.x 已移除 method 参数 | 用 pandas 1.x/2.x 环境运行（代码不许改） |
| `str.split('.',1,False)` FutureWarning（Jupyter 中显示为粉色框） | pandas 1.5 起位置参数将变 keyword-only | 用 pandas ≤1.4.4 |
| numpy ≥1.25 时 fillna 内部 DeprecationWarning | pandas 调用 np.find_common_type | ipykernel 默认过滤器不显示（非 `__main__` 触发），实测 nbconvert 输出干净；若仍出现则降 numpy 到 1.24.x 或在 kernel 配置过滤 |
| 综合题(3)(4) 报 `KeyError: 'Revenue'` | 这些题依赖题(1)创建的 Revenue 列 | 笔记本必须包含前置单元格（数据 cell + Revenue cell），按顺序执行 |
| `read_csv(..., names=['band','area','num'])` 数据却是4列 | 文件首列序号被 pandas 自动当作索引 | 这是预期行为（与指导书输出一致），不要改代码 |

**环境配方（已验证）**：系统 Python 3.9 + venv + `numpy==1.26.4` + `pandas==1.4.4` + `openpyxl` + `notebook==6.5.7` + `ipykernel`。安装耗时较长，用后台任务跑，给足超时（pip 每次装包 1-3 分钟）。

## 3. 阶段三：构建笔记本并真实执行

1. 每道题一个 `.ipynb`，**每个代码截图 = 一个 code cell**，单元格内容逐字一致（裸表达式 `df`、`newDF` 保持裸写，Jupyter 会显示 Out[n]）。综合题笔记本需含前置 cell。
2. 执行（较慢，每个 10-25 秒，务必给长超时或逐个跑，否则会被工具默认 120s 超时 SIGTERM）：

```bash
jupyter nbconvert --to notebook --execute --inplace q01_xxx.ipynb
```

3. 校验执行结果：读 JSON 检查 `outputs` 中无 `output_type=="error"`、每 cell 有 `execution_count`。
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

## 8. 交付前检查清单

- [ ] 报告中每个"（x.x）输入以下命令"的代码与原模板截图逐字一致
- [ ] 每个题的运行截图为真机截图且经目检（无对话框、无用户窗口、无 Not Trusted）
- [ ] 结果数值与指导书输出核对一致
- [ ] 原版 .doc、指导书 PDF、全部数据文件 md5 未变
- [ ] `实验N代码/` 中每个 .py 亲自运行成功；.ipynb 在 Jupyter 打开可复现
- [ ] README 含运行说明与 pandas 版本注意事项
- [ ] 完成版 Word 原版结构完好 + 截图位置正确 + 小结已填 + 封面页眉未动
