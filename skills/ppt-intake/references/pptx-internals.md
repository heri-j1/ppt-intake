# PPTX 内部结构速查（pptx-internals）

写给需要在 XML 级操作 pptx 的场合（克隆页、修 rels、读主题色）。
python-pptx 是这些结构的薄封装，它的对象模型都能映射回来。

## 包结构

```
foo.pptx (ZIP)
├── [Content_Types].xml        # 每个 part 的类型声明；新增 part 必须 Default/Override
├── _rels/.rels                # 根关系：指向 presentation.xml
├── ppt/
│   ├── presentation.xml       # <p:sldIdLst> 页序（r:id → slides/slideN.xml）
│   ├── _rels/presentation.xml.rels
│   ├── slides/slideN.xml      # 每页
│   ├── slides/_rels/slideN.xml.rels   # 页级关系：layout / 图片 / 备注
│   ├── slideLayouts/…         # 版式（继承 master，占位符几何在这层）
│   ├── slideMasters/…         # 母版（<p:clrMap> 把 scheme 映射到明暗槽位）
│   ├── theme/theme1.xml       # 主题：<a:clrScheme> 十二色 + <a:fontScheme>
│   └── media/imageN.*         # 图片字节
```

关键链条：`slide →(rels)→ layout →(rels)→ master →(rels)→ theme`。
页面上没写的格式属性沿这条链继承 —— 这就是"克隆基准页必然样式一致"的原因，
也是"直接搬材料页必然污染"的原因（链条终点是材料的 theme）。

## 颜色间接寻址

slide XML 里几乎不写死十六进制色，而是 `<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>`。
真实色值在 `theme1.xml` 的 `<a:clrScheme>`，经 master 的 `<p:clrMap>` 重映射
（如 `bg2 → lt2`）。inspect 脚本已代读出最终 hex，人工排查才需要碰这层。

字体同理：run 上常只有 `<a:latin typeface="+mn-lt"/>`（主题次要字体）；
CJK 字体看 `<a:ea>`，这才是中文实际渲染用的 face。

## XML 级克隆一页（merge_apply.py 内部做法）

python-pptx 无官方 API，标准做法：

1. `copy.deepcopy(src_slide._element)` 得到新 `<p:sld>` 树
2. 让包新建一个 slide part 挂这棵树（拿到新 part）
3. 遍历源页 `.rels`：layout 关系指向**基准自己的 layout**（本来就在同一包内，rId 重连即可）；
   图片等媒体关系用新 part 的 `relate_to(target_part, reltype)` 重新注册，
   并把树里旧的 `r:embed`/`r:id` 改成新 rId
4. 在 `presentation.xml` 的 `<p:sldIdLst>` 目标位置插入 `<p:sldId>`
   （新 sldId 数字取现有最大 +1）
5. `[Content_Types].xml` 由 python-pptx 保存时自动维护（part 走包 API 添加时）

判定克隆成功的信号：保存后 `Presentation(path)` 能重开且新页图片正常显示
（rels 断了图片会变空框 —— validate_output.py 专门查这个）。

## 文本格式保存原理（为什么首 run 替换不丢格式）

格式存在 run 级（`<a:rPr>` 挂在 `<a:r>` 上）。段落里多个 run 各带各的格式，
保留第一个 run 只改它的 `text`、删掉其余 run，视觉格式 = 原 run0 格式。
`paragraph.text = "..."` 会重建 run（格式回退到继承值）——所以禁止这么写。

## 页序与删页

页序 = `sldIdLst` 顺序，与文件名 slideN.xml 的 N 无关。删页要同时：
`drop_rel(rId)` + 从 `sldIdLst` remove。索引漂移：删第 i 页后所有 >i 的页索引 -1，
所以批量删按**高索引先**。
