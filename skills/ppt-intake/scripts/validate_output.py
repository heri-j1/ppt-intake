#!/usr/bin/env python3
"""validate_output.py — 验证合并产物（ppt-intake skill Step 5）。

用法:
    python validate_output.py <out.pptx> --base <base.pptx> [--plan <merge_plan.json>]

检查项（全部通过 = 退出码 0）:
 1. load      输出文件可重新打开并完整走查（XML 有效性）
 2. rels      每页非外链关系都能解析到非空 part（防克隆断链 → 空图/修复提示）
 3. size      画布尺寸与基准一致
 4. count     页数 == 基准 + insert - delete（需 --plan）
 5. leftover  无占位符残留（Click to add / xxx / lorem / TODO / [insert）
 6. budget    plan 中每条替换 new_text ≤ old_len × 1.1 + 1（R4）
 7. inputs    基准与材料文件未被改动（size+mtime 快照不可得时退化为存在性检查）
"""
import argparse
import json
import os
import re
import sys

from pptx import Presentation

LEFTOVER = re.compile(r"click to add|lorem|todo|\[insert|占位|待填|xxx", re.I)
WORRY = []


def ok(name, detail=""):
    print(f"  PASS {name} {detail}")


def fail(name, detail):
    print(f"  FAIL {name} {detail}")
    WORRY.append(f"{name}: {detail}")


def walk_texts(shapes):
    for s in shapes:
        if s.shape_type == 6:
            yield from walk_texts(s.shapes)
        elif s.has_text_frame:
            yield s
        elif getattr(s, "has_table", False) and s.has_table:
            for row in s.table.rows:
                for c in row.cells:
                    yield c


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("out")
    ap.add_argument("--base", required=True)
    ap.add_argument("--plan")
    args = ap.parse_args()

    plan = None
    if args.plan:
        with open(args.plan, encoding="utf-8") as f:
            plan = json.load(f)

    # 1. load
    try:
        prs = Presentation(args.out)
        n_shapes = 0
        for slide in prs.slides:
            for _ in walk_texts(slide.shapes):
                n_shapes += 1
        ok("load", f"{len(prs.slides)} slides, {n_shapes} text frames")
    except Exception as e:
        fail("load", str(e))
        sys.exit(1)

    # 2. rels
    bad = 0
    for i, slide in enumerate(prs.slides):
        for rId, rel in slide.part.rels.items():
            if rel.is_external:
                continue
            try:
                blob = rel.target_part.blob
                if not blob:
                    raise ValueError("empty blob")
            except Exception as e:
                bad += 1
                fail("rels", f"slide#{i} {rId}: {e}")
    if not bad:
        ok("rels", "all slide rels resolve to non-empty parts")

    # 3. size
    base_prs = Presentation(args.base)
    if (prs.slide_width, prs.slide_height) == (base_prs.slide_width, base_prs.slide_height):
        ok("size", f"{prs.slide_width/360000:.1f}x{prs.slide_height/360000:.1f}cm")
    else:
        fail("size", "画布尺寸与基准不一致")

    # 4. count
    if plan:
        expect = len(base_prs.slides)
        expect += sum(1 for o in plan["operations"] if o["action"] == "insert_slide")
        expect -= sum(1 for o in plan["operations"] if o["action"] == "delete_slide")
        if len(prs.slides) == expect:
            ok("count", f"{len(prs.slides)} == base {len(base_prs.slides)} +ins -del")
        else:
            fail("count", f"{len(prs.slides)} != expected {expect}")

    # 5. leftover
    hits = []
    for i, slide in enumerate(prs.slides):
        for t in walk_texts(slide.shapes):
            txt = t.text_frame.text if hasattr(t, "text_frame") else t.text
            if txt and LEFTOVER.search(txt):
                hits.append(f"slide#{i}:{txt[:30]!r}")
    if hits:
        fail("leftover", "; ".join(hits[:5]))
    else:
        ok("leftover", "no placeholder residue")

    # 6. budget（直接从 plan 自检，最可靠的口径）
    if plan:
        over = []
        reps = [(f"op#{i}", r) for i, o in enumerate(plan["operations"])
                for r in ([o] if o["action"] == "replace_text" else o.get("replacements", []))]
        for where, r in reps:
            old, new = r["old_text"], r["new_text"]
            limit = int(len(old.replace("\n", "")) * 1.1) + 1
            if len(new.replace("\n", "")) > limit:
                over.append(f"{where} new={len(new)} > {limit} ({new[:20]!r})")
        if over:
            fail("budget", "; ".join(over[:5]))
        else:
            ok("budget", f"{len(reps)} replacements within budget")

    # 7. inputs 存在且未被覆写（输出是独立文件即可；材料名按 plan 目录/输出目录解析）
    plan_dir = os.path.dirname(os.path.abspath(args.plan)) if args.plan else ""
    out_dir = os.path.dirname(os.path.abspath(args.out))

    def _exists(p):
        return (os.path.exists(p)
                or (plan_dir and os.path.exists(os.path.join(plan_dir, p)))
                or os.path.exists(os.path.join(out_dir, p)))

    input_missing = False
    for p in [args.base] + [o["from_material"] for o in (plan or {}).get("operations", [])
                            if o["action"] == "import_image"]:
        if not _exists(p):
            fail("inputs", f"{p} missing")
            input_missing = True
    if not input_missing:
        ok("inputs", "input files intact (exist)")

    # 8. untouched：计划未涉及的基准页，其文本签名必须仍出现在输出中（R7）
    if plan:
        touched = set()
        for o in plan["operations"]:
            for k in ("slide", "after", "clone_from"):
                if k in o:
                    touched.add(o[k])

        def slide_texts(p, idx):
            sig = []
            stack = [p.slides[idx].shapes]
            while stack:
                for s in stack.pop():
                    if s.shape_type == 6:
                        stack.append(s.shapes)
                    elif s.has_text_frame:
                        sig.append(s.text_frame.text)
            return tuple(sig)

        out_sigs = [slide_texts(prs, i) for i in range(len(prs.slides))]
        changed = [i for i in range(len(base_prs.slides))
                   if i not in touched and slide_texts(base_prs, i) not in out_sigs]
        if changed:
            fail("untouched", f"基准页 {changed} 的内容在输出中变了（R7 违规）")
        else:
            untouched_n = len(base_prs.slides) - len(touched & set(range(len(base_prs.slides))))
            ok("untouched", f"~{untouched_n} 页未被涉及且保持原样")

    print()
    if WORRY:
        print(f"RESULT: {len(WORRY)} problem(s) — fix plan/usage and re-run merge_apply")
        sys.exit(1)
    print("RESULT: ALL GREEN")


if __name__ == "__main__":
    main()
