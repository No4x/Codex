# 踩坑总表（43 条 · 按主题归类）

> 来源：《电气控制技术实训》项目一 60 页课件重构全过程实测记录。
> 标注 ⚠️ 的是**每次都会踩**的高频坑。

---

## A · ppt-master SVG 契约

1. **多行同段文字必须合成一个 `<text>` + `<tspan dy>`**：同 x 相邻、行距约 1.5× 的多个
   `<text>` 会被质检判为「段落被拆成兄弟节点」。列表项（行距 ≈2×）不受影响。
2. **根级 `<text>` / `<line>` 要么放进带 `data-pptx-bounds` 的 `<g>`，要么加
   `data-pptx-role="header"|"footer"|"decoration"`**，否则报 `ungrouped top-level element`。
3. **`<g>` 的 bounds 必须覆盖组内所有文字实际宽度**（含左侧起点，差 4px 也报溢出）。
   标定速率：`body24≈25.6px/字`｜`title43≈45.5`｜`subtitle27≈28.6`｜`annotation19≈20.1`｜
   `lead30≈32.3`｜`cover_title72≈76.9`。
4. **CSS 继承类属性（`text-anchor`）提到父 `<g>` 后转换器能正确继承** —— 紧凑写法可用，
   可清零「非紧凑写法」警告。但 ⚠️ **不要**跑 `compact_svg_styles.py --inplace` 全册：
   会重排全册格式、diff 噪声大、按行号写的批量脚本全失效。只对目标页做 group 提升。
5. **按行号批改 SVG 前必须核对文件是否含 `<defs>`**：有 `<defs>` 的页整体后移若干行；
   卡片组内「rect 在前、text 在后」也会错位一行。
6. **`svg_to_pptx` 会把 `data-pptx-bounds` 分组转成 PPT 的 GROUP 对象**，组内文字用
   `python-pptx` 的 `slide.shapes[].text_frame` **读不到**，必须递归进 `shape_type == 6`
   的子形状。顶层遍历判断「某页有没有某段文字」会得到**假阴性**。

## B · 质量门与导出

7. ⚠️ **导出前必须刷新 final 质量报告**：`svg_quality_checker.py --stage final --json`。
   漏 `--json` 只打印不落文件 → 导出报
   `requires a passing final SVG quality report for the current svg_output/; found stale`。
8. **`design_spec.md §VIII` 的图片状态列只接受合法枚举**。写成 `Ready` 会报 N 项 blocking
   `illustration_planned_image_invalid_status_error`。
9. **未被引用的图片必须同时从 §VIII 表格与 `spec_lock.md` images 段移出**，只删一边触发
   反向规则 `illustration_planned_image_missing_lock_error`。替代写法：改成表下说明行。
10. **验证脚本退出码为 1 但断言全过**：拼图阶段（PIL contact sheet）失败也会让退出码非 0。
    **判定成败要看日志里的断言，不能只看退出码。**

## C · 本机环境与工具

11. ⚠️ **本机无 pandoc** → `.ppt / .doc` 必须先经 Office COM 转换。
12. **Office COM 首次 `Presentations.Open` 会挂起** → 先杀 `POWERPNT.EXE` / `WINWORD.EXE`，重试即成功。
13. ⚠️ **PowerShell 工具在本机不回显 stdout**，且 **含中文路径的 `.ps1` 会被 PS 5.1 按 ANSI
    解码而静默失败** → 命令**内联**传给 PowerShell 工具，输出**写日志文件**再用 Read 读。
14. **Python 环境**：`~/.workbuddy/binaries/python/envs/default/Scripts/python.exe`
    （`--system-site-packages` 建的，因此同见 `playwright`）。ppt-master 的 `source_to_md.py`
    必须用这个 venv 跑。
15. **HTML 渲染成图**：Playwright（Chromium 已缓存）+ 临时 `http.server`。页面里引用本地图片
    **必须用相对路径**（`file:///D:/...` 会被 Chromium 拦掉）。flex 容器里的固定宽度 SVG 要加
    `flex:0 0 auto`，否则被压缩变形。
16. ⚠️ **无头浏览器截图必须传绝对路径的 URI**：`Path(相对路径).as_uri()` 抛
    `ValueError: relative path can't be expressed as a file URI` → 改 `.resolve().as_uri()`。
    命令：
    `msedge.exe --headless --disable-gpu --no-sandbox --hide-scrollbars --force-device-scale-factor=N --window-size=1280,720 --screenshot=<绝对路径> <file-URI>`
17. **PIL 画中文变方块**：`ImageDraw.text` 默认字体不含中文 →
    `ImageFont.truetype(r"C:\Windows\Fonts\msyh.ttc", 15)`，并做 msyh → simhei → simsun 逐级回退。

## D · 版面修订

18. ⚠️ **批量改头部/局部元素前必须先扫「区域占位」**：不能假设某区域空着。用 `data-pptx-bounds`
    过滤目标区域，一次性揪出全部冲突页。否则会「改完 → 回退」。
19. **徽标 bounds 与内容并非等边距**：bounds 右缘常比内容右缘多 4–8px →
    **按内容右缘对齐**更可靠，并顺手把 bounds 右缘收进安全边距。
20. **预览面板可能是旧版本**：`get_pool_status.last_saved_ms` 就是编辑器载入该文件的时间。
    若早于最近几次导出，说明面板里看的是**旧版**。
    → 选区类需求**先核对编辑器载入版本**；**绝不在陈旧实例里保存**（`save_file` 会覆盖回旧内容）。
    只读几何（`slide_get_page_info`）是安全的。
21. **批量改动的差异核验口径**：全册「恰好 N 页有差异、每页恰好 K 行」，其余页字节未变 ——
    用这个来证明没有误伤。

## E · 图片（含 AI 生成）

22. ⚠️ **图片受 §VIII ↔ spec_lock 双向闭合校验**，且 `Acquire Via` 与 `Status` 有配对约束：
    `ai→generated|needs-manual`｜`user→existing|needs-manual`｜`web→sourced|needs-manual`｜
    `formula→rendered|needs-manual`｜`placeholder→placeholder`。AI 图写 `Existing` 报
    `planned_image_status_mismatch`。
23. **采样倍率合法区间**：`render_scale = 显示尺寸 / 源图尺寸`。`>2.0` 报「会发虚」；
    `<0.25` 且源图 ≥1 MiB 报「体积过大」。安全区间 **0.5×–4×**。
24. **圆角用底色填充而非 alpha**：落位底色是 `#F1F5F9` 就把四角填成该色，避免转换器不吃 alpha
    出现黑角，也省 PNG 体积。
25. ⚠️ **AI 生成图右下角带「AI生成」水印** → 裁切必须**从底部切**（不是居中裁），
    按目标宽高比算裁切高度。
26. ⚠️ **配图落点不能只看「像有空白」**：必须**渲染 + 读源码**双确认。本次 P39/P06/P56 渲染看似
    可插，读源码发现正文+页脚已占满。
27. **章首页与场景页结构不同**：章首页留白在 `chapter-divider` **之后**；场景页没有该结构，
    只能插在**场景条内部**，且先量出正文真实右缘（如 P48 正文止于 `x≈896`），否则压字。

## F · 嵌入对象（OLE）

28. **ppt-master 导出契约禁止 `file://` 本地链接**（只允许 `https://` 与 `#slide-N`）→
    「打开本地 HTML」**写不进 SVG 源**，只能做成导出后工序。
29. ⚠️ **`Shapes.AddOLEObject` 只接受关键字参数**（位置参数报
    `could not convert string to float: ''`）；命名参数里的 `Left/Top` 会被忽略，
    **必须创建后再赋 `sh.Left/Top/Width/Height`**；`Width` 还会被 PowerPoint 按图标原始比例回收
    （设 247.5 pt 实际存 38.3 pt）。
30. ⚠️ **PowerPoint 的 `OLEFormat` 没有 `IconLabel` / `IconIndex` / `DisplayAsIcon` / `FileName`**
    （typelib 里就没有，Word 才有）。图标只显示系统文件类型名（如「Microsoft Edge HTML Document」）
    → **必须在图标左侧另加一个文本框**写中文短标题。
31. **嵌入对象保真校验**：`ppt/embeddings/oleObjectN.bin` 是 OLE 复合文档，**源 HTML 字节原样在里面**，
    直接 `源字节 in bin字节` 子串比对即可（本项目 17/17 逐字节命中）。
    ⚠️ 但**勘误**：原始文件名**不在 OLE 目录项里**（目录项只有 `\x01Ole10Native` 之类），
    而在 `\x01Ole10Native` **流内部**，且是 **ANSI = mbcs / GBK 编码**。
    校验 `.html` 扩展名必须解析该流并 `bytes.decode("mbcs")` 再 `endswith(".html")`；
    用 UTF-8 正则、或只看 `olefile.listdir()` 的目录项名，都会得 **0 命中（假失败）**。
    可直接用 `scripts/verify_embeddings.py`（已含 `parse_ole10native`）。
32. ⚠️ **重导 PPT 后必须重跑嵌入脚本**，并把脚本里 `DEFAULT_SRC` 换成**最新时间戳**那份导出件；
    否则 N 个嵌入对象全丢。另：**目标 PPTX 被预览/编辑器占用时 `os.replace` 报
    `WinError 5 拒绝访问`** → 不硬覆盖，**升版本号**。

---

## G · 代码级坑（生成 HTML 时）

33. **`"%%IMG:%s:%s%%" % (...)` 里的 `%%` 会被 `%`-格式化当成转义** → 实际替换串只剩一个 `%`，
    结果是每张图 `src` 前多一个 `%`，**图片全部变成 file:// 相对路径而加载失败**，
    且事后用正则查残留占位符查不出来（**假通过**）。
    改法：用字符串拼接；并加两条断言 ——「输出含 `%%` 即报错」「内嵌图数量 == 占位符数」。

---

## H · 执行环境坑（2026-10-01 补）

34. ⚠️ **PowerShell 沙箱会拦截 COM 启动**：`New-Object -ComObject PowerPoint.Application` 在沙箱内执行会让
    **整条命令 `exit 1`**，而且**连 COM 之前就执行的 `Out-File` 日志也不落地**。
    诊断特征就是「日志文件根本不存在」→ 说明是**命令级拒绝**，不是脚本逻辑错。
    解法：该条命令加 `dangerouslyDisableSandbox`（需用户批准）。
    附带确认：`New-Item` / `Out-File` 这类纯文件操作在沙箱内**正常**，可用来做对照排查。

35. **`sed` / `grep -o` 打印中文会乱码**（Git Bash 按本地编码解析 UTF-8 文本），
    用它预览 md 会得到不可读输出 → 读文件内容一律走 **Read 工具**；
    只做「计数 / 是否存在」这类不含中文输出的检查时才用 grep。

36. ⚠️ **429 频率限制可能在工具已执行完之后才返回**：本项目实测 `embed_html_objects.py` 的 429 报错
    出现在**嵌入已完成之后**（`..._embedded.pptx` 已生成、全部对象已挂上）。
    → 遇到 429 **先核对产物是否已落地**（看导出目录时间戳/体积、跑一遍保真校验），再决定是否重跑；
    盲目重试会**重复嵌入**全部对象（对象数与页脚图标翻倍）。

37. **md 台账与页面会「各改一半」**：改完页面（页脚句式等）后，同一份 `项目一_课件脚本.md` 的
    **表格正文里往往还残留旧文本**（本轮 P03/P04/P05 三行的「依据：」就这么漏了）。
    → 收尾时对改动点在全库做一次 `grep` 全文检索，把**表格/清单里的同步列**一并改掉，别只改「未核实项」段。

---

## I · 页脚口径与合规标注（2026-10-06 补）

38. ⚠️ **删改指令涉及「合规标注」时，先问边界再动手**：本轮按「页脚注文全删」执行，
    把 8 页 AI 免责标注一并删了 → 用户要求保留 → 返工一轮写回。
    教训：「全删」不等于连**免责声明 / 版权标注**一起删，这类通常在保护范围内。
    丢给用户确认只花一句话，返工要重跑一整轮导出。
39. ⚠️ **AI 图 ↔ 页脚标注要做「一一对应」审计**：脚本扫全册
    `href="../images/scene_*_ai.png"`（或 `design_spec §VIII` 的 ai 项）与
    「配图为 AI 生成示意图」逐页比对。本轮据此查出 **P53 有 AI 配图却从未标注**（v3 遗留漏标）。
    人工记忆查不出，必须做成**固定核对手段**（8 图 ↔ 8 页标注）。
40. **`footer-note` 批量删除的稳妥写法**：正则
    `[ \t]*<text id="footer-note"[^>]*>.*?</text>[ \t]*\r?\n?`，`re.subn` 后
    **断言每页恰好 1 处命中**（防误删/漏删）；删完 `grep` 关键词复核全册 0 残留。
    写回时保持原 `id`、`data-pptx-role="footer"` 与坐标（`x=64 y=688 font-size=15 fill=#64748B`）。
41. ⚠️ **页码写死「P xx / 60」→ 加页必须连带同步全册页码**：本套每页页脚都写着总页数，
    加 1 页要 ① 全册总页数 `/ 60` → `/ 61`；② 插入位**之后**各页 `P xx` 顺延 +1；
    ③ 目录 / 导航 / 备注 / 嵌入挂接表里的页号引用同步；④ 全册 `grep` 复核页码连续无重复。
    **策略：优先加页** —— 内容值得独立成页就加，不要为省页码硬挤排版（挤排版会牺牲可读性）；
    但**页码同步是不可省的一步**，用脚本按「页号 ≥ 插入位」批量 +1（**文件名 + 页内文本一起改**），别手改。
    —— 并页法（压排版）只在「内容很少、不值得占一页」时用：本轮 P60 曾把五档作业条
    58px 压成 30px 行，那是**当时的选择**，不是默认策略。
42. ⚠️ **嵌入对象数变化必须三处同步**：① 嵌入脚本映射表（1 基页号 → 文件 → 标签）
    ② `knowledge/INDEX.md` 的数量 ③ 对应 PPT 页教师备注的提及。
    缺任何一处即「账实不符」（本轮 16 → 17，新增「互动10 三只元件待判」挂 P60）。
43. ⚠️ **别把「本次的条数」写成规则**：skill 里凡出现「7 动画 / 10 互动 / 17 对象」
    这类数字，必须**同时标明它是当次实况**，不能读成「每套课件都要这么多」。
    互动与动画的条数**由内容决定、无配额**（判据＝哪些知识点需要学生**动手判一次**
    或**看一次动作**；该多就多、该少就少，不为凑数硬塞，也不照搬上一套）。
    → 本轮自查发现 3 处被写死并已修：`SKILL.md` 的「①～⑩」与「动画 1–7 与互动 1–10」、
    两个脚本 docstring 的旧数（「共 16 个」「9 个教学互动」）。
    → 通用做法：**规范段只用变量 N/M；实例数字一律加「本项目 / 本次」限定**。
