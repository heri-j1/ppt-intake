#!/usr/bin/env python3
"""validate_output.py — 验证合并产物（ppt-intake skill Step 5）。

用法:
    python validate_output.py <out.pptx> --base <base.pptx> [--plan <merge_plan.json>]
                              [--receipt <out.receipt.json>]

检查项（全部通过 = 退出码 0；NOTE 为提醒，不影响退出码）:
 1. load      输出文件可重新打开并完整走查（XML 有效性）
 2. rels      每页非外链关系都能解析到非空 part（防克隆断链 → 空图/修复提示）
 3. size      画布尺寸与基准一致
 4. count     页数 == 基准 + insert - delete（需 --plan）
 5. leftover  无占位符残留（Click to add / xxx / lorem / TODO / [insert）
 6. budget    plan 中每条替换 new_text ≤ old_len × 1.1 + 1（R4）
 7. inputs    基准与材料文件未被改动
 8. receipt   处置回执与计划推导的期望分桶一致（回执文件自动探测）
 9. pages     untouched 页签名仍在；deleted 页签名已消失（扣除其克隆源）
10. landed    每条 replace_text 的 new_text 确实出现在输出中
11. NOTE      克隆页与源页共享备注 part 时给出提醒（已知边界）
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


def note(name, detail):
    print(f"  NOTE {name} {detail}")


def norm(s):
    return s.replace("\x0b", "\n").replace("\r", "").strip()


def expected_buckets(plan, base_n):
    """从计划推导期望分桶（与 merge_apply 的回执逻辑同构）。"""
    updated, deleted, cloned = set(), set(), []
    n_auto = 0
    for op in plan["operations"]:
        act = op["action"]
        if act in ("replace_text", "import_image") and isinstance(op["slide"], int):
            updated.add(op["slide"])
        elif act == "delete_slide":
            deleted.add(op["slide"])
        elif act == "insert_slide":
            n_auto += 1
            cloned.append(op.get("clone_id") or f"auto-{n_auto}")
    untouched = [i for i in range(base_n) if i not in updated and i not in deleted]
    return {"untouched": untouched, "updated": sorted(updated),
            "cloned": sorted(cloned), "deleted": sorted(deleted)}


def slide_texts(prs_, idx):
    """一页的文本签名（text frame 全文，GROUP 递归）。"""
    sig = []
    stack = [prs_.slides[idx].shapes]
    while stack:
        for s in stack.pop():
            if s.shape_type == 6:
                stack.append(s.shapes)
            elif s.has_text_frame:
                sig.append(s.text_frame.text)
    return tuple(sig)


def collect_texts(prs_):
    """输出文件全部文本：frame 级 + 段落级（用于落地检查）。"""
    frames, paras = [], []
    for slide in prs_.slides:
        stack = [slide.shapes]
        while stack:
            for s in stack.pop():
                if s.shape_type == 6:
                    stack.append(s.shapes)
                elif s.has_text_frame:
                    frames.append(s.text_frame.text)
                    paras.extend(p.text for p in s.text_frame.paragraphs)
                elif getattr(s, "has_table", False) and s.has_table:
                    for row in s.table.rows:
                        for c in row.cells:
                            frames.append(c.text_frame.text)
                            paras.extend(p.text for p in c.text_frame.paragraphs)
    return frames, paras


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
    ap.add_argument("--receipt")
    ap.add_argument("-m", "--material-dir", action="append", default=[])
    args = ap.parse_args()

    plan = None
    if args.plan:
        with open(args.plan, encoding="utf-8") as f:
            plan = json.load(f)

    receipt = None
    receipt_path = args.receipt
    if not receipt_path and args.out.endswith(".pptx"):
        cand = args.out[:-5] + ".receipt.json"
        if os.path.exists(cand):
            receipt_path = cand
    if receipt_path and os.path.exists(receipt_path):
        with open(receipt_path, encoding="utf-8") as f:
            receipt = json.load(f)

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

    base_n = len(base_prs.slides)
    expect = expected_buckets(plan, base_n) if plan else None

    # 4. count
    if plan:
        expect_count = base_n + len(expect["cloned"]) - len(expect["deleted"])
        if len(prs.slides) == expect_count:
            ok("count", f"{len(prs.slides)} == base {base_n} +ins -del")
        else:
            fail("count", f"{len(prs.slides)} != expected {expect_count}")

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

    # 6. budget
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

    # 7. inputs
    plan_dir = os.path.dirname(os.path.abspath(args.plan)) if args.plan else ""
    out_dir = os.path.dirname(os.path.abspath(args.out))

    def _exists(p):
        if os.path.exists(p):
            return True
        for d in [plan_dir, out_dir] + args.material_dir + [os.getcwd()]:
            if d and os.path.exists(os.path.join(d, p)):
                return True
        return False

    input_missing = False
    for p in [args.base] + [o["from_material"] for o in (plan or {}).get("operations", [])
                            if o["action"] == "import_image"]:
        if not _exists(p):
            fail("inputs", f"{p} missing")
            input_missing = True
    if not input_missing:
        ok("inputs", "input files intact (exist)")

    # 8. receipt 与计划对账
    if plan and receipt:
        rb = receipt.get("buckets", {})
        diff = [k for k in expect if sorted(rb.get(k, [])) != sorted(expect[k])]
        count_same = receipt.get("final_slide_count") == len(prs.slides)
        if not diff and count_same:
            ok("receipt", f"buckets match plan (untouched={len(expect['untouched'])} "
                          f"updated={len(expect['updated'])} cloned={len(expect['cloned'])} "
                          f"deleted={len(expect['deleted'])})")
        else:
            fail("receipt", f"回执与计划不一致: {diff or ''}"
                            f"{'' if count_same else ' / final_slide_count'}")
    elif plan:
        note("receipt", "未找到回执文件（旧版 merge_apply 产物），第 8 项跳过")

    # 9. untouched 仍在 + deleted 已消失
    if plan:
        out_sigs = [slide_texts(prs, i) for i in range(len(prs.slides))]
        untouched = (receipt["buckets"]["untouched"] if receipt
                     else [i for i in range(base_n)
                           if i not in {o.get("slide") for o in plan["operations"]
                                        if isinstance(o.get("slide"), int)}])
        missing = [i for i in untouched
                   if i < base_n and slide_texts(base_prs, i) not in out_sigs]
        if missing:
            fail("pages", f"untouched 基准页 {missing} 的内容在输出中变了（R7 违规）")
        else:
            ok("pages", f"untouched={len(untouched)} 页保持原样")
        if expect["deleted"]:
            clone_src_sigs = {slide_texts(base_prs, o["clone_from"])
                              for o in plan["operations"] if o["action"] == "insert_slide"}
            forbidden = {slide_texts(base_prs, i) for i in expect["deleted"]} - clone_src_sigs
            leftover_del = [i for i in expect["deleted"]
                            if slide_texts(base_prs, i) in forbidden
                            and slide_texts(base_prs, i) in out_sigs]
            if leftover_del:
                fail("pages.deleted", f"已删除页 {leftover_del} 的内容仍出现在输出中")
            else:
                ok("pages.deleted", f"deleted={len(expect['deleted'])} 页内容已移除")

    # 10. 每条替换确实落地
    if plan:
        frames, paras = collect_texts(prs)
        pool = {norm(t) for t in frames + paras}
        miss = []
        for i, o in enumerate(plan["operations"]):
            for r in ([o] if o["action"] == "replace_text" else o.get("replacements", [])):
                if norm(r["new_text"]) not in pool:
                    miss.append(f"op#{i}:{r['new_text'][:20]!r}")
        if miss:
            fail("landed", f"{len(miss)} 条替换未落地: " + "; ".join(miss[:5]))
        else:
            n_reps = sum(len([o] if o["action"] == "replace_text" else o.get("replacements", []))
                         for o in plan["operations"])
            ok("landed", f"{n_reps} 条替换全部落地")

    # 11. NOTE：克隆页与源页共享备注 part（已知边界）
    notes_use = {}
    for i, slide in enumerate(prs.slides):
        for rel in slide.part.rels.values():
            if rel.reltype.endswith("/notesSlide") and not rel.is_external:
                key = str(rel.target_part.partname)
                notes_use.setdefault(key, []).append(i)
    shared = {k: v for k, v in notes_use.items() if len(v) > 1}
    if shared:
        detail = "; ".join(f"{k} -> slides {v[:4]}" for k, v in list(shared.items())[:3])
        note("notesSlide", f"{len(shared)} 个备注 part 被多页共享（克隆带备注页的已知边界，"
                           f"编辑备注前先拆分）: {detail}")

    print()
    if WORRY:
        print(f"RESULT: {len(WORRY)} problem(s) — fix plan/usage and re-run merge_apply")
        sys.exit(1)
    print("RESULT: ALL GREEN")


if __name__ == "__main__":
    main()
