#!/usr/bin/env python3
"""merge_apply.py — 在基准副本上执行 merge_plan.json（ppt-intake skill Step 4）。

用法:
    python merge_apply.py <base.pptx> <merge_plan.json> -o <out.pptx> [-m <材料目录>]

规则（见 templates/merge_plan.md）:
- 所有页号均按基准原始页号（0-based）书写；脚本用"页身份表"处理执行漂移。
- 每次运行都从基准原件冷启动全量重放（幂等）。
- 绝不写入任何输入文件路径。
"""
import argparse
import copy
import json
import os
import shutil
import sys
from io import BytesIO

from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.text.text import _Paragraph

R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def _norm(s):
    return s.replace("\x0b", "").replace("\r", "").strip()


def replace_in_paragraph(p, new_text):
    """首 run 替换保格式；new_text 不含 \\n（full 模式在外层拆行）。"""
    runs = p.runs
    if not runs:
        if new_text:
            p.add_run().text = new_text
        return
    runs[0].text = new_text
    for r in runs[1:]:
        r._r.getparent().remove(r._r)


def replace_full(tf, new_text):
    """整框替换：按行分配到现有段落；行数多于段落时按末段格式追加。"""
    lines = new_text.split("\n")
    paras = list(tf.paragraphs)
    template = copy.deepcopy(paras[-1]._p) if len(lines) > len(paras) else None
    for i, p in enumerate(paras):
        replace_in_paragraph(p, lines[i] if i < len(lines) else "")
    for line in lines[len(paras):]:
        el = copy.deepcopy(template)
        tf._txBody.append(el)
        replace_in_paragraph(_Paragraph(el, tf), line)


def replace_para(tf, old_text, new_text):
    """逐段匹配替换（new_text 视为单行）。"""
    m = _norm(old_text)
    for p in tf.paragraphs:
        if _norm(p.text) == m:
            replace_in_paragraph(p, new_text)
            return True
    return False


def apply_replacement(slide, rep, where):
    """对一页执行单条替换。rep 可带 shape_id 精确定位。"""
    match = rep.get("match", "full")
    old, new = rep["old_text"], rep["new_text"]
    sid = rep.get("shape_id")
    frames = []

    def walk(shapes):
        for s in shapes:
            if s.shape_type == 6:
                walk(s.shapes)
            elif s.has_text_frame:
                frames.append(s)

    walk(slide.shapes)
    if sid is not None:
        frames = [s for s in frames if s.shape_id == sid]
        if not frames:
            print(f"  WARN [{where}] shape_id={sid} 不在这页上")
            return False
    for shape in frames:
        tf = shape.text_frame
        if match == "full" and _norm("\n".join(p.text for p in tf.paragraphs)) == _norm(old):
            replace_full(tf, new)
            return True
        if match == "para" and replace_para(tf, old, new):
            return True
    print(f"  WARN [{where}] 未找到匹配: old_text={old[:40]!r} shape_id={sid}")
    return False


def clone_slide(prs, src_slide):
    """XML 级克隆：deepcopy 形状树 + 背景，重注册 rels 并重映射 rId。"""
    new_slide = prs.slides.add_slide(src_slide.slide_layout)
    for shp in list(new_slide.shapes):  # 清掉 add_slide 自动生成的占位符
        shp._element.getparent().remove(shp._element)
    src_cSld = src_slide._element.find(qn("p:cSld"))
    bg = src_cSld.find(qn("p:bg"))
    if bg is not None:
        new_cSld = new_slide._element.find(qn("p:cSld"))
        new_cSld.insert(0, copy.deepcopy(bg))
    remap = {}
    for rId, rel in src_slide.part.rels.items():
        if rel.reltype.endswith("/slideLayout"):
            continue  # add_slide 已建立（同一 layout 对象）
        if rel.is_external:
            new_rid = new_slide.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        else:
            new_rid = new_slide.part.relate_to(rel.target_part, rel.reltype)
        remap[rId] = new_rid
    for shp in src_slide.shapes:
        new_slide.shapes._spTree.append(copy.deepcopy(shp._element))
    if remap:
        for el in new_slide.shapes._spTree.iter():
            for attr in (qn("r:embed"), qn("r:id"), qn("r:link")):
                v = el.get(attr)
                if v in remap:
                    el.set(attr, remap[v])  # 每 attr 只读一次再写，交换安全
    return new_slide


class Deck:
    """页身份表：identity -> slide 对象 + 期望页序。
    identity ∈ 基准原始页号 | 克隆序号(n, n+1, …)。
    物理页序只在最终 sync() 时对齐一次，中途所有操作与位置解耦。
    同一锚点(after)多次插入时，按计划顺序依次排列（链尾追踪）。
    """

    def __init__(self, prs):
        self.prs = prs
        self.order = list(range(len(prs.slides)))
        self.map = {i: prs.slides[i] for i in self.order}
        self.next_clone = len(prs.slides)
        self.clones = {}
        self.anchor_tail = {}

    def slide(self, identity):
        return self.map[identity]

    def insert_after(self, cid, after, slide_obj):
        self.map[cid] = slide_obj
        self.clones[cid] = after
        tail = self.anchor_tail.get(after)
        anchor = tail if tail is not None else after
        self.order.insert(self.order.index(anchor) + 1, cid)
        self.anchor_tail[after] = cid

    def delete(self, identity):
        slide_obj = self.map.pop(identity)
        self.order.remove(identity)
        lst = self.prs.slides._sldIdLst
        for sld in list(lst):
            rId = sld.get(qn("r:id"))
            if self.prs.part.rels[rId].target_part is slide_obj.part:
                self.prs.part.drop_rel(rId)
                lst.remove(sld)
                return
        raise RuntimeError(f"delete: 找不到页 {identity} 的 sldId 条目")

    def sync(self):
        """把物理页序对齐到 self.order。"""
        lst = self.prs.slides._sldIdLst
        el_of_part = {}
        for sld in list(lst):
            el_of_part[self.prs.part.rels[sld.get(qn("r:id"))].target_part] = sld
        for identity in self.order:
            lst.append(el_of_part[self.map[identity].part])  # append 移动到末尾，循环后即期望序


def import_image(deck, op, material_paths, where):
    from_material = op["from_material"]
    path = material_paths.get(from_material)
    if not path:
        raise FileNotFoundError(f"{where}: 找不到材料文件 {from_material}（-m 目录里没有）")
    mat = Presentation(path)
    mslide = mat.slides[op["material_slide"]]
    mshape = next((s for s in mslide.shapes if s.shape_id == op["material_shape_id"]), None)
    if mshape is None or mshape.shape_type != 13:
        raise ValueError(f"{where}: 材料页 {op['material_slide']} 没有 shape_id="
                         f"{op['material_shape_id']} 的图片")
    blob, ext = mshape.image.blob, mshape.image.ext
    target = deck.slide(op["slide"])
    tshape = next((s for s in target.shapes if s.shape_id == op["shape_id"]), None)
    if tshape is None or tshape.shape_type != 13:
        raise ValueError(f"{where}: 目标页 {op['slide']} 没有 shape_id={op['shape_id']} 的图片形状")
    img_part, rId = target.part.get_or_add_image_part(BytesIO(blob))
    tshape._element.blipFill.blip.set(qn("r:embed"), rId)
    return True


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("base")
    ap.add_argument("plan")
    ap.add_argument("-o", "--out", required=True)
    ap.add_argument("-m", "--material-dir", action="append", default=[])
    args = ap.parse_args()

    with open(args.plan, encoding="utf-8") as f:
        plan = json.load(f)

    material_paths = {}
    for name in {op.get("from_material") for op in plan["operations"]
                 if op.get("action") == "import_image"}:
        for d in [os.path.dirname(os.path.abspath(args.plan))] + args.material_dir + ["."]:
            p = os.path.join(d, name) if d else name
            if os.path.exists(p):
                material_paths[name] = os.path.abspath(p)
                break

    inputs = {os.path.abspath(args.base)} | set(material_paths.values())
    out_abs = os.path.abspath(args.out)
    if out_abs in inputs:
        sys.exit(f"REFUSE: 输出路径 {args.out} 与输入文件相同（baseline-rules R3）")

    shutil.copyfile(args.base, args.out)  # 冷启动：从基准原件开始
    prs = Presentation(args.out)
    deck = Deck(prs)
    done = 0

    # 预克隆：克隆源一律取自未修改的基准原件（计划的 clone_from 语义 = 原始页号，
    # 与 ops 顺序无关——即使该页先被 replace_text/delete，克隆的仍是原始内容）
    pre = {}
    cids = []
    n_auto = 0
    for op in plan["operations"]:
        if op["action"] == "insert_slide":
            n_auto += 1
            cid = op.get("clone_id") or f"auto-{n_auto}"
            if cid in pre:
                raise ValueError(f"clone_id 重复: {cid}")
            pre[cid] = clone_slide(prs, deck.slide(op["clone_from"]))
            cids.append(cid)
    cid_iter = iter(cids)

    for i, op in enumerate(plan["operations"]):
        where = f"op#{i} {op.get('action')}"
        act = op["action"]
        if act == "replace_text":
            slide = deck.slide(op["slide"])
            if apply_replacement(slide, op, where):
                done += 1
        elif act == "insert_slide":
            cid = op.get("clone_id") or next(cid_iter)
            clone = pre[cid]
            for rep in op.get("replacements", []):
                apply_replacement(clone, rep, f"{where} rep")
            deck.insert_after(cid, op["after"], clone)
            done += 1
        elif act == "delete_slide":
            deck.delete(op["slide"])
            done += 1
        elif act == "import_image":
            import_image(deck, op, material_paths, where)
            done += 1
        else:
            raise ValueError(f"{where}: 未知 action")
    deck.sync()
    prs.save(args.out)

    # 处置回执（receipt）：逐页去向，validate_output 用它与计划精确对账。
    # buckets: untouched=未涉及 | updated=被 replace/import 改过 |
    #          cloned=克隆新增 | deleted=已删除
    base_n = deck.next_clone
    updated, deleted = set(), set()
    for op in plan["operations"]:
        act = op["action"]
        if act in ("replace_text", "import_image") and isinstance(op["slide"], int):
            updated.add(op["slide"])
        elif act == "delete_slide":
            deleted.add(op["slide"])
    untouched = [i for i in range(base_n) if i not in updated and i not in deleted]
    pages = []
    for pos, ident in enumerate(deck.order):
        if ident in deck.clones:
            pages.append({"position": pos, "disposition": "cloned",
                          "clone_id": ident, "clone_from": deck.clones[ident]})
        elif ident in updated:
            pages.append({"position": pos, "disposition": "updated", "base_slide": ident})
        else:
            pages.append({"position": pos, "disposition": "untouched", "base_slide": ident})
    receipt = {"receipt_version": "1", "output": os.path.basename(args.out),
               "base_slide_count": base_n, "final_slide_count": len(deck.order),
               "buckets": {"untouched": untouched, "updated": sorted(updated),
                           "cloned": sorted(deck.clones), "deleted": sorted(deleted)},
               "pages": pages}
    stem = args.out[:-5] if args.out.endswith(".pptx") else args.out
    receipt_path = stem + ".receipt.json"
    with open(receipt_path, "w", encoding="utf-8") as f:
        json.dump(receipt, f, ensure_ascii=False, indent=1)
    print(f"OK {done}/{len(plan['operations'])} ops -> {args.out} "
          f"(slides: {len(deck.order)}, clones: {len(deck.clones)}, "
          f"skipped: {len(plan.get('skipped', []))}, conflicts: {len(plan.get('conflicts', []))})")
    print(f"receipt: untouched={len(untouched)} updated={len(updated)} "
          f"cloned={len(deck.clones)} deleted={len(deleted)} -> {receipt_path}")


if __name__ == "__main__":
    main()
