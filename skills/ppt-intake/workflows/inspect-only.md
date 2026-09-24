# 只读工作流：结构检查 / 差异报告（inspect-only）

用户只想"看看两份 PPT 结构/内容差异"或"我的 PPT 里都有什么"时使用。
**不产出任何 pptx 文件。**

## 步骤

1. 对涉及的每份文件跑盘点：

```bash
python scripts/inspect_pptx.py <a.pptx> -o work/a.json
python scripts/inspect_pptx.py <b.pptx> -o work/b.json
```

2. 读 JSON，输出两份人读报告（直接答给用户，不写文件除非用户要）：

**结构清单**（单文件时）：
- 页数、画布尺寸、主题色、主要字体
- 每页一行：`页号 [role] 标题 —— N 个文本框 / 图 / 表`

**差异报告**（双文件对比时）：
- 主题对比：两边的主色/字体/画布尺寸是否一致（不一致 = 整合时需走 ppt-intake 主流程）
- 内容重叠：哪些页讲同一主题、各自的数据口径
- 独有内容：a 有 b 没有的章节、b 有 a 没有的章节
- 体量对比：页数、图表数、数据密度

## 边界

- 本工作流**不**给合并建议的执行细节；用户说"那帮我合并一下" → 切到
  `workflows/integrate-materials.md` 从 Step 2 开始（Step 1 产物可复用）。
