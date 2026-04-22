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
import json
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

THEME_COLOR_NAMES = {
    "dk1", "lt1", "dk2", "lt2",
    "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
}
HEX_RE = re.compile(r"^#?[0-9A-Fa-f]{6}$")


def convert(src_path: Path, tpl_path: Path, out_path: Path,
            color_map_path: Path | None = None) -> None:
    src = Presentation(src_path)
    dst = Presentation(tpl_path)

    if (src.slide_width, src.slide_height) != (dst.slide_width, dst.slide_height):
        # Ignore sub-inch rounding differences that show up between decks authored in
        # different tools — they don't warrant a warning.
        width_delta = abs(src.slide_width - dst.slide_width)
        height_delta = abs(src.slide_height - dst.slide_height)
        if width_delta > 9144 or height_delta > 9144:  # > 0.01"
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

    theme_colors = _theme_colors(dst)
    color_map = _load_color_map(color_map_path, theme_colors) if color_map_path else {}
    if color_map_path:
        print(f"[info] loaded {len(color_map)} color mappings from {color_map_path}")

    _delete_all_slides(dst)

    per_slide_report: list[tuple[int, dict[str, set[str]]]] = []
    for idx, src_slide in enumerate(src.slides):
        new_slide = dst.slides.add_slide(blank_layout)
        _strip_default_shapes(new_slide)
        rid_map = _clone_slide_rels(src_slide, new_slide, idx)
        _copy_shapes(src_slide, new_slide, rid_map, color_map)
        per_slide_report.append((idx, _categorize_colors(src_slide, color_map)))

    out_path.parent.mkdir(parents=True, exist_ok=True)
    dst.save(out_path)
    print(f"[done] wrote {out_path}")
    _print_color_report(per_slide_report, theme_colors, bool(color_map))


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


def _copy_shapes(src_slide, dst_slide, rid_map, color_map):
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
        if color_map:
            _remap_colors(new_el, color_map)
        dst_tree.append(new_el)


def _remap_colors(root, color_map):
    """Rewrite matching <a:srgbClr val="X"/> elements per color_map.

    A mapping target can be either a template theme-slot name (rendered as
    <a:schemeClr val="accent1"/>) or a hex literal (stays as <a:srgbClr>).
    Preserves any shade/tint/alpha child modifiers on the original color element.
    """
    srgb_tag = f"{{{A_NS}}}srgbClr"
    scheme_tag = f"{{{A_NS}}}schemeClr"
    for old_el in list(root.iter(srgb_tag)):
        val = (old_el.get("val") or "").upper()
        target = color_map.get(val)
        if target is None:
            continue
        is_theme = not target.startswith("#")
        new_el = etree.Element(scheme_tag if is_theme else srgb_tag)
        new_el.set("val", target.lstrip("#") if not is_theme else target)
        for child in list(old_el):
            new_el.append(child)
        old_el.addnext(new_el)
        old_el.getparent().remove(old_el)


def _load_color_map(path: Path, theme_colors: dict[str, str]) -> dict[str, str]:
    """Load a color-map JSON file and validate entries.

    The file is { "description": "...", "map": { "<srcHex>": "<target>", ... } }.
    Targets are either a theme-slot name (dk1/lt1/dk2/lt2/accent1..accent6/hlink/folHlink)
    or a '#RRGGBB' literal. Unknown theme names cause a fatal error so typos don't
    silently no-op.
    """
    data = json.loads(path.read_text())
    raw = data.get("map", {}) if isinstance(data, dict) else {}
    normalized: dict[str, str] = {}
    for src, target in raw.items():
        if not isinstance(src, str) or not HEX_RE.match(src):
            raise SystemExit(f"color-map key not a 6-hex string: {src!r}")
        src_key = src.lstrip("#").upper()
        if not isinstance(target, str):
            raise SystemExit(f"color-map value for {src!r} is not a string: {target!r}")
        if target.startswith("#"):
            if not HEX_RE.match(target):
                raise SystemExit(f"color-map hex target for {src!r} invalid: {target!r}")
            normalized[src_key] = "#" + target.lstrip("#").upper()
        elif target in THEME_COLOR_NAMES:
            if target in ("accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
                          "dk1", "lt1", "dk2", "lt2") and target not in theme_colors:
                print(f"[warn] color-map: template theme has no {target!r}; mapping of "
                      f"{src!r} will still emit a schemeClr reference (PowerPoint will resolve).",
                      file=sys.stderr)
            normalized[src_key] = target
        else:
            raise SystemExit(
                f"color-map target for {src!r} not recognized: {target!r}. "
                f"Expected a theme name ({', '.join(sorted(THEME_COLOR_NAMES))}) or '#RRGGBB'."
            )
    return normalized


def _categorize_colors(slide, color_map) -> dict[str, set[str]]:
    """Return per-slide color buckets: 'remapped' (source hexes that were hit by color_map)
    and 'unmapped' (source hexes with no mapping — still hardcoded in output)."""
    xml = etree.tostring(slide._element, encoding="unicode")
    all_colors = {m.group(1).upper() for m in SRGB_RE.finditer(xml)}
    return {
        "remapped": {c for c in all_colors if c in color_map},
        "unmapped": {c for c in all_colors if c not in color_map},
    }


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


def _print_color_report(per_slide_report, theme_colors, used_color_map: bool):
    print()
    print("=== Color report ===")
    if theme_colors:
        print("Template theme palette:")
        for name, hx in theme_colors.items():
            print(f"  {name:8s} #{hx}")
    print()
    # Aggregate across slides so the user sees the full set of unmapped colors to address.
    all_remapped: set[str] = set()
    all_unmapped: set[str] = set()
    for _, buckets in per_slide_report:
        all_remapped |= buckets["remapped"]
        all_unmapped |= buckets["unmapped"]

    if used_color_map:
        print(f"Source colors remapped ({len(all_remapped)} unique):")
        if all_remapped:
            print("  " + ", ".join("#" + c for c in sorted(all_remapped)))
        else:
            print("  (none — the mapping had no hits on this source)")
        print()
        print(f"Source colors still hardcoded — add these to your color map to re-theme them "
              f"({len(all_unmapped)} unique):")
        if all_unmapped:
            print("  " + ", ".join("#" + c for c in sorted(all_unmapped)))
        else:
            print("  (none — mapping covered everything)")
    else:
        print("Source colors hardcoded in output (no --color-map supplied):")
        print("  " + ", ".join("#" + c for c in sorted(all_unmapped)))


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--input", required=True, type=Path, help="Source .pptx")
    ap.add_argument("--template", required=True, type=Path, help="Corporate template .pptx")
    ap.add_argument("--output", required=True, type=Path, help="Output merged .pptx")
    ap.add_argument("--color-map", type=Path, default=None,
                    help="Optional JSON file mapping source hex colors to template theme slots "
                         "or hex literals. See mappings/keyhole_to_corporate.json for an example.")
    args = ap.parse_args()
    for label, p in (("input", args.input), ("template", args.template)):
        if not p.exists():
            ap.error(f"{label} not found: {p}")
    if args.color_map is not None and not args.color_map.exists():
        ap.error(f"color-map not found: {args.color_map}")
    convert(args.input, args.template, args.output, color_map_path=args.color_map)


if __name__ == "__main__":
    main()
