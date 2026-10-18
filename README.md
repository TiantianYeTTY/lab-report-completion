# lab-report-completion

> 面向《数据科学与分析》等数据类课程的 **Python 实验报告自动化完成技能**（WorkBuddy Agent Skill）。
>
> 从解析实验指导书与报告模板，到搭建运行环境、生成**真机运行截图**、整理代码源文件，再到在**原版 Word 副本上原位填充**——一条龙完成，且全程不改动原始文件。
>
> 作者：**TiantianYe**

---

## ✨ 这是什么

`lab-report-completion` 是一个 Agent Skill（智能体技能包），把一次完整的实验报告制作流程沉淀为可复用的工作流与踩坑经验。它解决的问题非常具体：

- 课程实验报告的模板是 `.doc` 格式，里面的代码以**截图**形式给出，学生需要逐字抄写运行、再手工截图粘贴回 Word——繁琐且容易出错；
- 不同版本的 pandas/Jupyter 对同一份课程代码的表现不同，稍有差池截图里就会出现报错或警告；
- 截图需要是**真机运行结果**（不是渲染或伪造的图片），并且要精确插入到 Word 的对应位置。

本技能将整个流程标准化：**代码逐字转录执行 → 真实浏览器截图 → 原版 Word 原位填充**，并对过程中的 20+ 个易错点给出了根因分析与标准解法。

## 🎯 使用场景

当你的任务符合以下画像时，即可使用本技能：

| 要素 | 说明 |
|---|---|
| 课程 | 《数据科学与分析》等以 pandas 为核心的数据类课程 |
| 输入材料 | ① 实验指导书 PDF；② 实验报告模板（.doc，代码为截图）；③ 数据文件（.xlsx / .csv） |
| 典型内容 | 数据清洗、数据抽取、排名索引、数据合并、groupby / pivot_table 等 pandas 操作 |
| 用户诉求 | "完成实验报告、贴真实截图（不是渲染/画的）、代码逐字按截图、源码保存供亲自运行、原文件禁止更改" |

典型触发指令示例：

> 根据实验指导书和数据文件（原文件和数据禁止更改），按照实验报告要求，完成实验报告内容。务必完全按照实验报告中的代码截图输入代码、并在实验报告中贴上符合要求的真实截图（不是渲染的、画的，真机截图）、切勿编造结果。每道题的代码源文件保存在文件夹中，以便我亲自运行检验。

## 📦 安装

将本仓库克隆（或复制 `lab-report-completion/` 文件夹）到 Agent 技能目录：

```bash
git clone https://github.com/TiantianYeTTY/lab-report-completion.git ~/.workbuddy/skills/lab-report-completion
```

重启会话后技能即被识别。之后在做实验报告时，用上面示例中的方式描述任务（附上指导书、数据文件与报告模板），智能体会自动加载本技能；也可以显式调用 `/lab-report-completion`。

## 🚀 使用方法（工作流概览）

技能按以下 7 个阶段自动执行，完整细节与全部命令见 [SKILL.md](SKILL.md)：

| 阶段 | 内容 | 关键点 |
|---|---|---|
| 1. 解析输入 | pymupdf 提取指导书文本与页面图；textutil + olefile 从 .doc 二进制中提取全部代码截图并**逐字转录** | 图片按 Data 流签名雕刻提取 |
| 2. 数据检查 | 只读检查数据文件的结构、dtype、NaN 与重复行 | 原件零改动 |
| 3. 干跑验证 | 全部代码以"警告升级为错误"的最严格模式预跑 | 提前消灭版本兼容问题 |
| 4. 环境构建 | Python 3.9 + `pandas==1.4.4` + `numpy==1.26.4` + `notebook==6.5.7` 专用 venv | 课程代码在 pandas 3.x 会报错，详见 SKILL.md 版本矩阵 |
| 5. 真机截图 | Jupyter 真实执行笔记本 → Chrome + CDP 控制导航/滚动 → Quartz 按 PID 定位窗口 → 系统 `screencapture` 截取 | 每张截图自动目检，杜绝对话框/串窗 |
| 6. 交付包 | `实验N代码/`：py 源码 + ipynb + 截图 + 数据副本 + README | 每个 .py 均运行验证 |
| 7. Word 原位填充 | 复制原版 .doc → 通过 editor_sdk 在副本上原位插入截图与实验小结 → 另存 .docx | 保留原模板全部标题、字体、版式与原代码截图 |

### 输入 / 产出

```
输入                              产出
─────────────────────────        ─────────────────────────────
实验指导书.pdf            ───▶    完成版报告 .docx（原版副本原位填充）
报告模板.doc（含代码截图）  ───▶    实验N代码/（py + ipynb + 截图 + 数据副本）
数据文件 xlsx/csv          ───▶    运行结果截图/（真机截图）
```

## 🔧 环境依赖

- **macOS**（截屏管线依赖系统 `screencapture` 与 Quartz）
- **Google Chrome**（真实浏览器渲染 + CDP 远程调试）
- **Python 3.9** 及上表所列包（建议专用 venv，避免污染全局环境）
- **WorkBuddy 内置 editor_sdk**（Word 原位编辑通道；在 WorkBuddy 环境外可参考 SKILL.md 中的原理，用 LibreOffice/Word COM 等价替代）
- 屏幕录制权限（终端/宿主应用需在系统设置中授权）

## ⚠️ 重要说明

1. **截图真实性是硬约束**：所有结果截图来自真实运行的浏览器窗口截屏，绝不合成或伪造。
2. **原文件零改动**：指导书、报告模板原件与数据文件全程只读，交付前以 md5 校验副本一致性。
3. **版本敏感**：课程代码使用了 `fillna(method='pad')`、`str.split('.',1,False)` 等旧式写法，pandas 3.x 下会报错——这不是代码错误，请按 SKILL.md 的环境配方运行。
4. **.doc → .docx**：编辑内核只支持写出 OOXML，填充后的报告以 `.docx` 保存（内容与排版与原版一致），原版 `.doc` 保持原样。

## 📁 目录结构

```
lab-report-completion/
├── README.md               ← 本文件
├── SKILL.md                ← 技能正文：完整工作流 + 20 条易错点速查表 + 检查清单
└── scripts/
    ├── capture_screenshots.py   ← 真机截图管线（Chrome+CDP+screencapture，单tab持久导航）
    ├── update_docx_images.py    ← Word 图片按文档序号替换/删除（edsdk）
    └── fill_report.py           ← 原版 Word 副本原位填充截图与小结（JSON 计划驱动）
```

## 🐛 已知踩坑点（节选）

SKILL.md 中记录了完整的 20 条「错误现象 → 根因 → 正确处理」，包括：

- 环境代理劫持 localhost 导致 CDP 连接 404；
- Jupyter classic 的滚动容器是 `#site` 而非 documentElement；
- 多标签页反复开关导致 DevTools WebSocket 全面超时（应使用单 tab 持久导航）；
- `doc_find` 匹配只是段落前缀，直接 `end+1` 插入会把标题撕成两半；
- 强杀 Chrome 后的崩溃对话框与"恢复标签页"气泡的清理方法；
- ……

## 🛠️ 可复用脚本

`scripts/` 内三个脚本均为真实实验全流程实测产物，参数化后可直接运行（详见 SKILL.md §8）：

```bash
# ① 真机截图（前置：按 SKILL.md §4.1 启动 Jupyter 与 Chrome）
python3 scripts/capture_screenshots.py --notebooks-dir /abs/笔记本目录 --out-dir /abs/截图目录

# ② Word 图片替换/删除（先 --dry-run 查看序号）
python3 scripts/update_docx_images.py /abs/报告.docx --dry-run

# ③ 原版 Word 副本原位填充（JSON 计划驱动，--print-plan-template 查看结构）
python3 scripts/fill_report.py plan.json
```

## 👤 作者

**TiantianYe**

欢迎通过 Issue 反馈问题，或提交 PR 完善工作流（例如补充 Windows 下的截图管线、适配其他课程模板）。

## 📄 许可

仅供学习交流使用。
