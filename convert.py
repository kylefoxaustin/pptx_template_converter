#!/usr/bin/env python3
"""Apply a corporate PowerPoint template's theme/master/layouts to an existing deck.

Strategy A — "theme swap": use the template as the base, port source slides onto
the template's Blank layout. Source's absolutely-positioned shapes are preserved
exactly; the template's theme (colors, fonts) and layout branding take over.

Usage:
    python convert.py --input INPUT.pptx --template TEMPLATE.pptx --output OUTPUT.pptx
"""
from __future__ import annotations

import argparse
import re
import sys
from copy import deepcopy
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.parts.image import Image, ImagePart

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
SRGB_RE = re.compile(r'srgbClr\s+val="([0-9A-Fa-f]{6})"')

# spTree children that must stay put when we wipe a slide's default shapes.
GROUP_FRAME_TAGS = {qn("p:nvGrpSpPr"), qn("p:grpSpPr")}


def convert(src_path: Path, tpl_path: Path, out_path: Path) -> None:
    src = Presentation(src_path)
    dst = Presentation(tpl_path)

    if (src.slide_width, src.slide_height) != (dst.slide_width, dst.slide_height):
        print(
            f"[warn] canvas size differs: source {src.slide_width}x{src.slide_height}, "
            f"template {dst.slide_width}x{dst.slide_height}. Keeping source dimensions; "
            f"template branding designed for the other ratio may look stretched.",
            file=sys.stderr,
        )
        dst.slide_width = src.slide_width
        dst.slide_height = src.slide_height

    blank_layout = _find_blank_layout(dst)
    print(f"[info] using template layout: {blank_layout.name!r}")

    _delete_all_slides(dst)

    per_slide_report: list[tuple[int, set[str]]] = []
    for idx, src_slide in enumerate(src.slides):
        new_slide = dst.slides.add_slide(blank_layout)
        _strip_default_shapes(new_slide)
        rid_map = _clone_slide_rels(src_slide, new_slide, idx)
        _copy_shapes(src_slide, new_slide, rid_map)
        per_slide_report.append((idx, _hardcoded_colors(src_slide)))

    out_path.parent.mkdir(parents=True, exist_ok=True)
    dst.save(out_path)
    print(f"[done] wrote {out_path}")
    _print_color_report(per_slide_report, _theme_colors(dst))


def _find_blank_layout(pres):
    for layout in pres.slide_layouts:
        if layout.name.lower() == "blank":
            return layout
    return pres.slide_layouts[0]


def _delete_all_slides(pres):
    """Remove every slide from `pres` — drop the sldId entries AND their relationships."""
    sld_id_lst = pres.slides._sldIdLst
    r_id_attr = f"{{{R_NS}}}id"
    for sld_id in list(sld_id_lst):
        rid = sld_id.get(r_id_attr)
        pres.part.drop_rel(rid)
        sld_id_lst.remove(sld_id)


def _strip_default_shapes(slide):
    sp_tree = slide.shapes._spTree
    for child in list(sp_tree):
        if child.tag not in GROUP_FRAME_TAGS:
            sp_tree.remove(child)


def _clone_slide_rels(src_slide, dst_slide, slide_idx):
    """Port images and hyperlinks from src_slide onto dst_slide's part.

    Returns a map of src-side rId -> dst-side rId so we can remap r:embed / r:link / r:id
    references on the copied shape XML.
    """
    rid_map: dict[str, str] = {}
    for rel in src_slide.part.rels.values():
        # dst_slide already has the correct layout relationship; don't overwrite.
        if rel.reltype.endswith("/slideLayout"):
            continue

        if rel.is_external:
            new_rid = dst_slide.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
            rid_map[rel.rId] = new_rid
            continue

        src_part = rel.target_part
        if isinstance(src_part, ImagePart):
            img = Image.from_blob(src_part.blob, src_part.desc)
            new_part = ImagePart.new(dst_slide.part.package, img)
            new_rid = dst_slide.part.relate_to(new_part, rel.reltype)
            rid_map[rel.rId] = new_rid
        else:
            print(
                f"[warn] slide {idx_for_log(slide_idx)}: skipping related part "
                f"{src_part.partname} (reltype={rel.reltype}). Visuals tied to it may be broken.",
                file=sys.stderr,
            )
    return rid_map


def idx_for_log(i: int) -> str:
    return f"{i + 1}"


def _copy_shapes(src_slide, dst_slide, rid_map):
    dst_tree = dst_slide.shapes._spTree
    r_attrs = (f"{{{R_NS}}}embed", f"{{{R_NS}}}link", f"{{{R_NS}}}id")
    for src_el in list(src_slide.shapes._spTree):
        if src_el.tag in GROUP_FRAME_TAGS:
            continue
        new_el = deepcopy(src_el)
        for el in new_el.iter():
            for attr in r_attrs:
                old = el.attrib.get(attr)
                if old is not None and old in rid_map:
                    el.attrib[attr] = rid_map[old]
        dst_tree.append(new_el)


def _hardcoded_colors(slide) -> set[str]:
    xml = etree.tostring(slide._element, encoding="unicode")
    return {m.group(1).upper() for m in SRGB_RE.finditer(xml)}


def _theme_colors(pres) -> dict[str, str]:
    master = pres.slide_masters[0]
    theme_part = next(
        (r.target_part for r in master.part.rels.values() if r.reltype.endswith("/theme")),
        None,
    )
    if theme_part is None:
        return {}
    root = etree.fromstring(theme_part.blob)
    out: dict[str, str] = {}
    for name in ("dk1", "lt1", "dk2", "lt2",
                 "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"):
        el = root.find(f".//{{{A_NS}}}{name}")
        if el is None:
            continue
        srgb = el.find(f"{{{A_NS}}}srgbClr")
        if srgb is not None:
            out[name] = srgb.get("val").upper()
    return out


def _print_color_report(per_slide_report, theme_colors):
    print()
    print("=== Color report ===")
    if theme_colors:
        print("Template theme palette (colors at these values ARE re-themed):")
        for name, hx in theme_colors.items():
            print(f"  {name:8s} #{hx}")
    theme_set = set(theme_colors.values())
    print()
    print("Hardcoded colors in source slides (these were NOT re-themed by the swap):")
    any_reported = False
    for idx, colors in per_slide_report:
        off_theme = sorted(colors - theme_set)
        if off_theme:
            any_reported = True
            print(f"  slide {idx + 1}: " + ", ".join("#" + c for c in off_theme))
    if not any_reported:
        print("  (none — source only used theme-equivalent colors)")


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--input", required=True, type=Path, help="Source .pptx")
    ap.add_argument("--template", required=True, type=Path, help="Corporate template .pptx")
    ap.add_argument("--output", required=True, type=Path, help="Output merged .pptx")
    args = ap.parse_args()
    for label, p in (("input", args.input), ("template", args.template)):
        if not p.exists():
            ap.error(f"{label} not found: {p}")
    convert(args.input, args.template, args.output)


if __name__ == "__main__":
    main()
