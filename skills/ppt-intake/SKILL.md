---
name: ppt-intake
description: "Use when the user wants to merge/integrate material from other people's presentations into their own existing PPT while keeping THEIR deck's format, style and template intact. Trigger words: 合并PPT, 整合PPT材料, 把别人的PPT合并到我的PPT, PPT材料整合, 汇报材料整合, merge presentations, integrate slides from another deck, combine PPTs keeping my template, import content from material deck. Also when formats/styles of the two decks conflict and the user's own format must win. NOT for creating a deck from scratch or topic-to-deck generation."
license: MIT
---

# PPT Intake —— 材料整合

把别人给的 PPT 材料（格式、模板、字体与你的不一致）**摄入**你自己的基准 PPT，
产出一版**格式与基准完全一致**的成品。核心思想：**基准 PPT 的每一页都是模板**——
新页 = 克隆基准同角色页 + 替换文本；绝不新建页面，绝不直接搬运材料页。

## Always Read（每次任务开始前必读）

1. `rules/baseline-rules.md` —— 铁律（违反任何一条 = 返工）：基准格式优先、
   输出新文件不动原件、文本预算 `orig_len × 1.1`、绝不用 python-pptx 从零建页
2. `templates/merge_plan.md` —— 合并计划的 JSON schema（第 3 步要写这个）

## 任务路由

| 任务 | 去读 | 产出 |
|---|---|---|
| **整合材料**（主流程） | `workflows/integrate-materials.md` | `<stem>_merged.pptx` + `变更说明.md` |
| 只想看差异/结构，不改文件 | `workflows/inspect-only.md` | 结构清单 + 差异报告 |
| 需要克隆页 / 修 rels / 主题色 | `references/pptx-internals.md` | — |
| 脚本报错 / 结果损坏 / 排错 | `references/gotchas.md` | — |

## 脚本（确定性工作交给脚本，判断性工作由你做）

```bash
# 1. 盘点：PPT → 结构 JSON 清单（含每页角色、文本预算、主题色/字体）
python scripts/inspect_pptx.py <file.pptx> -o <out.json>

# 2. 执行合并计划（在基准副本上操作）
python scripts/merge_apply.py <base.pptx> <merge_plan.json> -o <out.pptx>

# 3. 验证输出（必须全绿才算完成）
python scripts/validate_output.py <out.pptx> --base <base.pptx> --plan <merge_plan.json>
```

依赖：仅 `python-pptx`（`pip install python-pptx`）。

## Known Gotchas（摘要，详见 references/gotchas.md）

- python-pptx **没有**官方克隆 slide API —— 必须按 internals 文档在 XML 级克隆并修 rels
- 材料页直接复制会**带入外来母版/主题/字体** —— 这是本 skill 最要防的事故
- 文本替换只改**首 run**（保格式）；`\x0b` 软换行先 `_norm()` 再匹配
- 删页**高索引先删**；克隆页的媒体 rId 必须全部重注册
- `Presentation.save()` 会覆写目标路径 —— 输出路径永远用新文件名，绝不指向原件
- 新替换文案长度 ≤ `orig_len × 1.1`（清单里有 budget 字段，写计划时对着抄）

## 完成标准

`validate_output.py` 全部通过 + 抽查 3 页（1 克隆页 + 1 替换页 + 1 原样页）确认
格式与基准一致 + 向用户报告：改了什么、丢了什么（附理由）、有什么冲突。
