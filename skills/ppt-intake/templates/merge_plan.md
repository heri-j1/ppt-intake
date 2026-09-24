# Merge Plan JSON Schema（v1）

合并计划的唯一格式。字段名区分大小量。所有页号（`slide`/`after`/`clone_from`/
`delete`）均为**基准原始页号（0-based）**，脚本内部处理执行漂移。

```jsonc
{
  // 必填。版本号，当前 "1"
  "version": "1",
  // 必填。输出文件名（相对于 merge_apply 的 -o 参数所在目录；仅作记录）
  "output": "base_merged.pptx",

  // 必填（可为空数组）。按数组顺序执行。
  "operations": [

    // ── 1. 替换基准现有页的文本（update 类动作）──
    {
      "action": "replace_text",
      "slide": 3,                    // 基准原始页号 0-based
      "shape_id": 5,                 // 可选；不给则按 old_text 全页匹配（易误伤，慎用）
      "match": "full",               // full=整框替换 | para=逐段匹配替换，默认 full
      "old_text": "市场规模 1.2 亿元",   // 抄 inspect 清单 text 字段（含换行）
      "new_text": "市场规模 1.35 亿元",
      "source": "material_a.pptx#1"  // 可选但强烈建议：内容出处（溯源用，
                                     // 无可靠来源的内容不该进计划）
    },

    // ── 2. 克隆基准页并替换文本后插入（append 类动作）──
    {
      "action": "insert_slide",
      "after": 5,                    // 插到基准原始页号 5 之后
      "clone_from": 4,               // 克隆源 = 基准原始页号 4（同角色页）
      "clone_id": "p_chan",          // 可选；后续操作用 "slide": "p_chan" 引用这个新页
      "replacements": [              // 作用于克隆出的新页
        {"old_text": "市场概览", "new_text": "竞品格局", "match": "full"},
        {"old_text": "三行要点\n第一行\n第二行\n第三行", "new_text": "要点A\n要点B"}
      ]
    },

    // ── 3. 删除页 ──
    { "action": "delete_slide", "slide": 12 },

    // ── 4. 从材料导入图片到基准现有形状（从严使用）──
    {
      "action": "import_image",
      "from_material": "material_a.pptx",  // 材料文件名（merge_apply 的 -m 搜索路径中）
      "material_slide": 4,            // 材料页号 0-based
      "material_shape_id": 7,         // 材料里的图片 shape_id
      "slide": 6,                     // 目标：基准原始页号
      "shape_id": 3                   // 目标图片形状（保留其位置尺寸，仅换图内容）
    }
  ],

  // 必填（可为空数组）。drop 的记录，仅进变更说明
  "skipped": [
    {"source": "material_a.pptx#0", "reason": "封面与成品重复"}
  ],

  // 必填（可为空数组）。conflict 记录，不产生操作
  "conflicts": [
    {"desc": "2024营收：基准1.2亿 vs material_a 1.35亿", "resolution": "待用户确认，本次未采用任何一方"}
  ]
}
```

## 规则

1. `old_text` 一律从 inspect 清单的 `text` 字段复制（repr 级准确，含 `\n`）。
2. `new_text` 长度 ≤ 目标框 `budget`（label 框另见 baseline-rules R4）。
3. `insert_slide` 可连续多条；`after` 都写基准原始页号。
4. `import_image` 换图不换框：目标形状的位置/尺寸/裁剪保持基准原样。
5. 计划要可重放：merge_apply 每次都从基准原件冷启动执行全量 operations。

## 处置回执（receipt）

merge_apply 执行成功后，会在输出文件旁写 `<out>.receipt.json`：逐页去向
（buckets: untouched / updated / cloned / deleted + 逐页 disposition 明细）。
`validate_output.py` 会自动探测该文件并与计划推导的期望分桶**精确对账**——
"未涉及的页保持原样"由此从承诺变为可验证的事实。回执是工具产物，不要手改。
