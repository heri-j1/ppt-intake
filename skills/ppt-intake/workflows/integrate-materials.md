# 主工作流：整合材料（integrate-materials）

输入：`base.pptx`（基准，格式赢家）+ 1..n 份材料 PPT。
输出：`<stem>_merged.pptx` + `变更说明.md`。全程遵守 `rules/baseline-rules.md`。

## Step 1 · 盘点（脚本，确定性）

对基准和每份材料各跑一次，得到 JSON 清单：

```bash
python scripts/inspect_pptx.py base.pptx -o work/base.json
python scripts/inspect_pptx.py material_a.pptx -o work/material_a.json
```

清单包含：幻灯片尺寸、每页 role 推断、逐 shape 位置/尺寸/文本（exact，
含换行）/字号/占位符类型/图片/表格、每个文本框的 `budget`、主题色与字体。

**读清单，不要猜**。后续所有 `old_text` 都从这里复制。

## Step 2 · 对齐分析（你，判断性）

按 `rules/mapping-rules.md` 的判定流程，把每份材料的每一页归入
`update / append / drop / conflict`，先写成人读的映射表（材料页 → 动作 → 目标 → 理由）。
材料总量大时按**章节**粗分组再逐页判定，避免遗漏；每页都要有去向（R5）。

## Step 2.5 · 合并方案确认（交互式任务时）

把 Step 2 的映射表（update/append/drop/conflict 四类 + 理由）先给用户过目，
**conflict 必须让用户裁决**，drop 有争议时也给机会翻案。用户确认后再写
merge_plan.json。非交互（批处理/已授权全自动）时跳过本步，但变更说明里
仍要完整披露。这一步对应"设计稿先确认再量产"——比生成完再返工便宜得多。

## Step 3 · 写合并计划（你 → JSON）

按 `templates/merge_plan.md` 的 schema 写 `work/merge_plan.json`：
- update → `replace_text` 操作（`slide` + `shape_id` 定位 + `new_text`）
- append → `insert_slide` 操作（`clone_from` 基准页 + `replacements` 逐框映射）
- drop → 不产生操作，记入 `skipped`
- conflict → 不产生操作，记入 `conflicts`
- 材料图片确需保留 → `import_image`（从严使用，R6）

每条 op 建议带 `source` 字段标注内容出处（如 `"material_a.pptx#1"`）——
无可靠来源的内容不该进计划（事实溯源）。

写完自查三件事：
1. 每个 `new_text` 长度 ≤ 对应框 `budget`（清单字段）
2. 每个 `replacements` 的 `old_text` 与克隆源页的 exact text 一致（含换行）
3. `insert_slide.after` 的页号基于**基准当前页序**（insert 按计划顺序执行，页号会随执行漂移——计划里的页号全部按"基准原始页号"写，脚本内部处理漂移）

## Step 4 · 执行（脚本）

```bash
python "${SKILL_DIR}/scripts/merge_apply.py" base.pptx work/merge_plan.json -o base_merged.pptx -m <材料目录>
```

脚本在副本上执行：克隆（XML 级 + 修 rels）→ 替换（首 run 保格式）→
插页 → 删页（高索引先）→ 图片导入；成功后写 `base_merged.receipt.json`
逐页处置回执（untouched/updated/cloned/deleted）。任何一步失败会报出
操作序号和原因，修计划后重跑（脚本是幂等重放：每次都从基准原件重新开始）。

## Step 5 · 验证与交付（脚本 + 你）

```bash
python "${SKILL_DIR}/scripts/validate_output.py" base_merged.pptx --base base.pptx --plan work/merge_plan.json
```

全绿的标志：receipt 与计划分桶对账通过、每条替换落地、untouched 页逐字未动。
之后抽查 3 页（1 克隆页 + 1 替换页 + 1 原样页）对照基准确认格式一致，
最后写 `变更说明.md`：

```markdown
# 变更说明（base + material_a → base_merged）
## 更新（update）
- 第3页「市场规模」：市占率 15% → 18%（来源：material_a 第2页，数据期更新）
## 新增（append）
- 第6页「竞品对比」：克隆自基准第4页样式，内容取 material_a 第4页
## 丢弃（drop）
- material_a 第1页（封面）：与成品封面重复
## 冲突（conflict，未处理，待你确认）
- 2024 营收：基准 1.2 亿 vs 材料 1.35 亿 —— 两边都未注明口径
```

## 失败回路

- apply/validate 报错 → 读 `references/gotchas.md`，按**修复层级原则**在拥有
  故障的最浅层修（单页问题改那条 op；结构取舍改计划整体；调用问题改命令），
  不要手改输出文件（下次 apply 会覆盖）。
- 克隆页观感不对（如字号突兀）→ 大概率替换进了错误角色的框，
  对照清单的 shape_id 重写 replacements。
