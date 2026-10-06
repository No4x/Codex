---
name: pptx-embed-html-interactions
description: 把单文件 HTML 互动资源（交互动画 / 课堂互动练习）作为「嵌入对象」挂进 PPTX，使 PPT 自包含、放映时双击图标用浏览器打开；并覆盖用「无头浏览器 + Playwright」真机验证这批 HTML 的完整流程。当用户说「把 HTML 插进 PPT」「把互动/动画嵌到课件里」「让 PPT 自带这些网页」「讲义要能带走所有互动」时使用。也适用于任何需要给 PPTX 追加「不可丢的附属文件」的场景。
agent_created: true
---

# 把单文件 HTML 互动资源嵌入 PPTX

## 何时用

- 已有（或刚做好）一批**单文件零依赖 HTML**（交互动画、课堂互动练习、可点选的小游戏），要随 PPT 一起交付；
- 用户要求「插进 PPT」「嵌到课件里」「换电脑也能打开」；
- 需要把任意附属文件（PDF、Excel、音频、网页）作为 Package 对象封进 PPTX。

## 对象数量：不设配额

**嵌入几个对象，由内容需要决定，没有固定值。** 本技能只负责「把已经做好的 HTML 挂上去」，
不规定做几个 —— 动画与互动练习的**条数按知识点需要算**（哪些点需要学生**动手判一次**
或**看一次动作**），该多就多、该少就少，不为凑数硬塞。

> 本项目落在 **7 动画 + 10 互动 = 17 个对象**，只是**该次实况，不构成标准**。
> 换一套课件必须**重新数**，不要沿用旧数。
> 脚本映射表、知识库索引、备注提及三处**始终用同一个 N**（见下「对象数量变化必须三处同步」）。

## 铁律：为什么必须做成「导出后工序」

**ppt-master 的 SVG→PPTX 导出契约禁止 `file://` 本地链接**，只允许 `https://` 与 `#slide-N`
（见 `ppt-master/references/native-hyperlinks.md`：relative paths / filesystem paths / `file:` 一律 fail closed；
「OLE or file actions」被列为 unsupported action settings）。

所以「让 PPT 打开本地 HTML」**写不进 SVG 源**。正确分层是：

```
SVG 源（版面唯一真源） → svg_to_pptx 导出 → [本技能] 导出后嵌入工序 → 交付 PPTX
```

推论：
- **SVG 版面一个字都不改**，嵌入是纯附加层；
- **每次重导 PPT 后必须重跑嵌入脚本**，否则嵌入对象全丢；
- 目标文件升版本号（v2/v3），不要把「有嵌入」和「无嵌入」混在同一个文件名里。

## 工作流

### 0. 先探一次（新环境必做）
不要假设 COM 行为，先用一个 HTML 跑通最小用例：
创建对象、设位置、另存、解包看 `ppt/embeddings/`。确认后再批量。

### 1. 准备 HTML（若还没有）
单文件零依赖 = 图片必须 base64 内嵌。见 `scripts/build_interaction_html.py`：
模板里写占位符 `%%IMG:文件名:宽度%%`，脚本等比缩放到目标宽度、白底转 JPEG q82、内嵌。

**构造替换串时不要用 `%`-格式化**。`"%%IMG:%s:%s%%" % (a, b)` 里 `%%` 会被当成转义，
替换串只剩一个 `%`，结果是 `src` 变成 `%"data:image/...` —— 图片静默全挂，且事后用正则查残留占位符**查不出来**。
用字符串拼接，并加断言：输出里含 `%%` 即报错；内嵌图数量必须等于占位符数量。

### 2. 真机验证 HTML（别靠肉眼）
`scripts/test_interaction_html.py`：Playwright + 本机 Edge（`channel="msedge"`，无需另装浏览器）。
对每个文件：捕获 `pageerror` / `console.error`，点主按钮断言答案区显隐，再点遍所有交互元素，最后 full_page 截图。
**注意**：`file://` URL 含中文必须用 `pathlib.Path.as_uri()` 做百分号编码；直接拼字符串会 `ERR_FILE_NOT_FOUND`。

### 3. 嵌入（核心）
`scripts/embed_html_objects.py`：pywin32 驱动 PowerPoint COM。

```python
sh = sl.Shapes.AddOLEObject(FileName=str(html_path), DisplayAsIcon=True)   # 只用关键字参数！
sh.Left, sh.Top, sh.Width, sh.Height = L, T, W, H                          # 创建后再赋坐标
tb = sl.Shapes.AddTextbox(1, lx, ty, tw, th)                               # 图标旁边补中文标签
```

必须知道的 5 个坑：

1. **`AddOLEObject` 只吃关键字参数**。传位置参数（ClassType 优先的那个签名）报
   `could not convert string to float: ''`。
2. **命名参数里的 `Left/Top` 会被忽略**，必须在创建之后赋 `sh.Left/Top/Width/Height`。
   且 `Width` 会被 PowerPoint 按图标原始比例回收（设 247.5 pt，实际存 38.3 pt）→ 别指望精确控宽。
3. **PowerPoint 的 `OLEFormat` 里没有 `IconLabel` / `IconIndex` / `DisplayAsIcon` / `FileName`**
   （typelib 就没有；Word 的才有）。图标只能显示**系统文件类型名**（例如「Microsoft Edge HTML Document」），
   **认不出是哪一个文件** → 必须用 `AddTextbox` 在图标旁边写中文短标题，否则课件里就是一堆无法区分的图标。
   文本框参数：`TextFrame.VerticalAnchor = 3`（居中）、`ParagraphFormat.Alignment = 3`（右对齐）、
   `Font.NameFarEast = "Microsoft YaHei"`、字号 9pt。
4. **目标 PPTX 被预览器/编辑器打开时 `os.replace` 会 `WinError 5 拒绝访问`**。
   不要硬覆盖：先探是否有锁，有锁就升版本号输出新文件（v2）。
5. **嵌入后要校验保真**（`scripts/inspect_ole.py`）：
   `ppt/embeddings/oleObjectN.bin` 是 OLE 复合文档，卸载后用 `olefile` 读 `\x01Ole10Native` 流，
   头部依次是 `DWORD 总长 | WORD 0x0002 | [可选版本头] | cstr 标签 | cstr 原始路径 | ... | DWORD 数据长 | 数据`。
   注意**头部字段长度各版本略有差异**，别按固定偏移硬算 —— 直接 `blob.index(b"<!DOCTYPE html>")` 定位数据段起点最稳。
   校验两件事：
   - 数据段 sha256 == 源文件 sha256（证明内容逐字节未损）；
   - 头部保留的**原始文件名带 `.html` 扩展名** → 双击后才会释放成 `.html` 并用浏览器打开（丢掉扩展名就只会得到一个打不开的临时文件）。
   再用 `ppt/slides/_rels/slideN.xml.rels` 里的 `../embeddings/oleObjectN.bin` 核对「页码 ↔ 文件」挂接表。

### 4. 落点选择（避免压字）
在 PPT 里给图标找**固定的空白窗**，不要逐页微调。做法：
1. 导出目标页 PNG（`pres.Slides(n).Export(path, "PNG", 1920, 1080)`）；
2. 用 numpy 在目标横带内找「暗像素簇」，量出既有元素（页脚注、页码）的实际起止；
3. 选一个所有目标页都不冲突的窗，再加 ≥40px 余量；
4. 页码 → 磅换算：`pt = px × 0.75`（16:9 幻灯片 960×540 pt 对应 SVG 1280×720 px，即 1 px = 0.75 pt）。

经验落点（1280×720 版式，页脚带）：标签盒 x 516–692、图标 x 700–749、y 666–710。

**落点区的前置条件**：早期版本的页脚在横线与页码之外还有一行左侧注文（`footer-note`，
如「素材来源：…」），量坐标时必须把它算进去；该注文现已**全册删除**（页脚只留横线 + 页码，
AI 配图页保留免责标注），落点区不再有文字冲突。**换一套课件时仍要重新量，不要照搬坐标。**

⚠️ **对象数量变化必须三处同步**：① 本脚本映射表 ② 知识库索引里的数量
③ 对应 PPT 页教师备注的提及 —— 任一不同步即「账实不符」。
（本项目曾在 16 → 17 时按此走；**数字随内容变**，见上文「对象数量：不设配额」。）

### 5. 交付前核验清单
- [ ] `slides` / `notesSlides` 数量与源一致（嵌入不该丢备注）
- [ ] `ppt/embeddings/` 条目数 == 计划嵌入数
- [ ] 每个 `.bin` 都能子串命中对应源文件
- [ ] 页码 ↔ oleObject 挂接表与计划一致
- [ ] 嵌入对象数与脚本映射表、知识库索引三处一致
- [ ] PowerPoint 实测导出若干页 PNG，目视确认图标落点无重叠、标签可读
- [ ] 交付文件 sha256 == 导出产物 sha256（证明确实落盘没串版本）

## 环境

- Python：`C:/Users/xzz-n/.workbuddy/binaries/python/envs/default/Scripts/python.exe`（已装 pywin32、comtypes、PIL、playwright）
- 浏览器：本机 Edge（`C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe`），Playwright 用 `channel="msedge"`
- COM 前必须 `pythoncom.CoInitialize()`，结束 `CoUninitialize()`，并 `pres.Close()` / `app.Quit()`

## 明确不适用的替代方案

- **SVG 里写 `<a href="file:///...">`**：被导出契约拒（fail closed）。
- **把 HTML 截成图片插页**：能看不能点，等于放弃交互；只在用户明确要「静态预览」时用。
- **发布成在线站点再放 https 链接**：完全兼容契约、版本永远最新，但需要联网与用户授权发布；
  离线课堂场景不要选它。
