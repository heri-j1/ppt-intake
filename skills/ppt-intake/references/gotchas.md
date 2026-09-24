# 坑集（Gotchas）

排错按症状查。修的是**计划或用法**，不是手改输出文件。

## 症状：文件打不开 / PowerPoint 提示修复

- 克隆页的 rels 没修干净：新页里 `r:embed` 指向不存在的 rId。
  → validate_output.py 的 rels 检查会提前抓到；apply 走的是包 API 注册，别绕开它手拼 XML。
- 手改过 `[Content_Types].xml`。→ 不要手改；用包 API 添加 part 后由保存流程维护。
- 图片 part 字节为空（从材料包取 blob 失败）。→ import_image 的 from_material
  路径写错，或材料页该形状不是图片。

## 症状：替换没生效 / 替换错框

- `\x0b` 软换行（Shift+Enter）藏在文本里，肉眼看不见。
  → replace 的匹配在脚本里做了 `_norm()`（去 `\x0b\r` 去首尾空白），
  但计划里的 old_text 若抄自"你看到的渲染文本"仍可能差一个空格。
  → 抄 inspect 清单里的 `text` 字段，那是 repr 级准确。
- 短词碰撞：`01`、`18`、`2024` 这类 token 多页重复。
  → replace_text 必须带 `slide`（+必要时 `shape_id`）作用域；replacements 只作用于克隆出的那一页，天然隔离。
- 同一框里 old_text 只匹配到段落级。→ replace 支持整框替换与逐段替换两种
  （plan schema 里 `match: full|para`），长文用 full。

## 症状：克隆页格式不对 / 字体突兀

- 克隆源选错角色（拿数据页当章节页）。→ 对照清单 `role` 与 shape 数量重选 `clone_from`。
- 替换文本把首 run 格式带丢：检查是否误用了 `paragraph.text =`。
  merge_apply 不会犯这个错——如果你绕开脚本手写了代码，改回脚本。
- 新文本里带英文但框的 `<a:ea>` 是中文字体 → 正常（CJK 字体渲染拉丁字符），不用处理。

## 症状：文本溢出

- 违反 R4：new_text 超了 budget。→ validate 会列出超预算框；删字优先，
  实在不行 `widen_to_fit` 思路（见官方 pptx skill）是下策。
- 克隆页塞了材料页的**全部**要点。→ 按 mapping-rules 拆两页，不要硬塞。

## 症状：删页后内容错位

- 索引漂移：先删低索引页导致后续页号全变。→ plan 里的 delete 全部写基准原始页号，
  脚本内部排序后高索引先删。**你永远按原始页号写计划。**

## 症状：插入页位置不对

- insert_slide.after 的页号在多次插入后漂移。→ 同上：全部写基准原始页号，
  脚本维护"原始页号 → 当前位置"映射。

## 环境坑

- Windows 路径含空格/中文：给脚本传参时整体加引号。
- python-pptx 版本 <1.0 的旧 API 差异（如 `slides._sldIdLst`）→ requirements 钉 `python-pptx>=1.0`。
- 同名输出文件被 PowerPoint/WPS 占用锁定 → 保存失败，关掉预览窗口重跑。
