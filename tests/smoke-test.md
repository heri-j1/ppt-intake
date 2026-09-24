# Smoke Test（全链路冒烟）

目的：验证 skill 的脚本链路 + agent 判定环节端到端可用。
预计耗时 < 2 分钟。依赖：`pip install python-pptx`。

## 前置

```bash
SKILL_DIR="$(pwd)/skills/ppt-intake"    # 路径先于命令：后续一律用 ${SKILL_DIR}
python samples/make_samples.py          # 生成 base.pptx + material_a.pptx（格式刻意冲突）
```

## T1 · 盘点脚本

```bash
python "${SKILL_DIR}/scripts/inspect_pptx.py" samples/base.pptx -o work/t1.json
```

预期：`OK ... 8 slides`；JSON 里 8 页、theme.colors 有 accent1、
每页有 role 与逐 shape 的 text/budget 字段。
核对：`#0` role=cover；`#4` role=stats（"18%" 框 budget=4）。

## T2 · 执行合并（含克隆/替换/图片导入）

```bash
python "${SKILL_DIR}/scripts/merge_apply.py" samples/base.pptx work/merge_plan.json -o work/t2.pptx -m samples
```

预期：`OK 6/6 ops`，无 WARN；11 slides, clones: 3；随后打印
`receipt: untouched=6 updated=2 cloned=3 deleted=0 -> work/t2.receipt.json`（untouched=6：页0/1/2/5/6/7；页6虽是克隆源但自身未被改）。
若报 WARN「未找到匹配」→ old_text 没抄 inspect 清单，或克隆源被先改过
（脚本已用预克隆规避后者，只剩抄写问题）。

## T3 · 验证

```bash
python "${SKILL_DIR}/scripts/validate_output.py" work/t2.pptx --base samples/base.pptx --plan work/merge_plan.json -m samples
```

预期：10 项 PASS（含 receipt 对账 / pages / landed），`RESULT: ALL GREEN`，退出码 0。

## T4 · agent 抽查（人工/AI 检查）

1. 用 inspect 盘点 t2.pptx，确认页序：
   `cover, toc, section, content(1.35亿/18%), stats(20%), section, 产品与竞品, 渠道结构, 竞品概览, 政策风险, closing`
2. 全部 font_name == 微软雅黑（材料宋体零泄漏）
3. #7 渠道结构 有 PICTURE 且 ext=png
4. #6 竞品 A 售价仍为 99 元（conflict 未擅改）

## T5 · delete 路径

```bash
echo '{"version":"1","operations":[{"action":"delete_slide","slide":1}],"skipped":[],"conflicts":[]}' > work/t5.json
python "${SKILL_DIR}/scripts/merge_apply.py" samples/base.pptx work/t5.json -o work/t5.pptx
python "${SKILL_DIR}/scripts/validate_output.py" work/t5.pptx --base samples/base.pptx --plan work/t5.json
```

预期：7 slides，`receipt: ... deleted=1`，ALL GREEN（含 pages.deleted 检查）。

## T6 · 输入保护

T2/T5 跑完后核对 `samples/base.pptx` 与 `samples/material_a.pptx` 的修改时间
未变（脚本从不写输入路径；validate 第 7 项亦检查存在性）。

## 最近一次结果

- 日期：2026-09-25（v2：回执对账 / landed / source 溯源 / SKILL_DIR 升级后重跑）
- T1–T6 全部通过；产物 `samples/base_merged.pptx` + `samples/变更说明.md` + 回执
- 冒烟中发现并已修复：克隆源被先前 replace_text 污染（→ 预克隆）、
  validate 相对路径误报（→ 按 plan/out 目录解析）
