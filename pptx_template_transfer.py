#!/usr/bin/env python3
"""PPTX Template Transfer — single-file, zero-config.

Transfer the design/theme/layouts from one PPTX to another, preserving content.

Usage:
    python3 pptx_template_transfer.py template.pptx content.pptx output.pptx
    python3 pptx_template_transfer.py template.pptx content.pptx output.pptx --layout-map mapping.json
"""

import argparse
import json
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from defusedxml.minidom import parse as parse_xml

# ============================================================================
# SHARED HELPERS
# ============================================================================

def _find_all(parent, tag_local):
    """Find all descendant elements matching a local tag name (ignoring ns prefix)."""
    if parent is None:
        return []
    return [c for c in parent.getElementsByTagName("*")
            if (c.localName or c.tagName.split(":")[-1]) == tag_local]


def _find_first(parent, tag_local):
    m = _find_all(parent, tag_local)
    return m[0] if m else None


def _attr(node, attr, default=""):
    if node is None:
        return default
    return node.getAttribute(attr) if node.hasAttribute(attr) else default


def _write_xml(doc, path):
    xml_str = doc.toxml().replace(
        '<?xml version="1.0" ?>',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    Path(path).write_text(xml_str, encoding="utf-8")


def _extract_number(filename):
    m = re.search(r"(\d+)", filename)
    return int(m.group(1)) if m else 0


def _parse_color_element(elem):
    if elem is None:
        return None
    for child in elem.childNodes:
        if getattr(child, "nodeType", None) != 1:
            continue
        local = child.localName or child.tagName.split(":")[-1]
        if local == "srgbClr":
            return "#" + _attr(child, "val", "000000")
        if local == "sysClr":
            lc = _attr(child, "lastClr", "")
            return f"sys:{_attr(child, 'val', '')}(#{lc})" if lc else f"sys:{_attr(child, 'val', '')}"
    return None


def _hex_to_rgb(h):
    h = h.lstrip("#")
    if len(h) != 6:
        return (128, 128, 128)
    try:
        return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except ValueError:
        return (128, 128, 128)


def _luminance(r, g, b):
    return (0.299 * r + 0.587 * g + 0.114 * b) / 255.0


def _is_dark(hex_color):
    r, g, b = _hex_to_rgb(hex_color)
    return _luminance(r, g, b) < 0.5


# ============================================================================
# INSPECT
# ============================================================================

def inspect_rels(rels_path):
    if not rels_path.exists():
        return []
    try:
        doc = parse_xml(str(rels_path))
    except Exception:
        return []
    return [{"id": _attr(r, "Id"), "type": _attr(r, "Type").split("/")[-1],
             "target": _attr(r, "Target")} for r in _find_all(doc, "Relationship")]


def inspect_theme(unpacked_dir):
    theme_dir = unpacked_dir / "ppt" / "theme"
    themes = []
    if not theme_dir.exists():
        return themes
    for tf in sorted(theme_dir.glob("theme*.xml")):
        try:
            doc = parse_xml(str(tf))
        except Exception as e:
            themes.append({"file": tf.name, "error": str(e)}); continue
        info = {"file": tf.name, "colors": {}, "fonts": {}, "name": ""}
        te = _find_first(doc, "theme")
        if te:
            info["name"] = _attr(te, "name", "")
        cs = _find_first(doc, "clrScheme")
        if cs:
            info["color_scheme_name"] = _attr(cs, "name", "")
            for cn in ["dk1","lt1","dk2","lt2","accent1","accent2","accent3",
                        "accent4","accent5","accent6","hlink","folHlink"]:
                info["colors"][cn] = _parse_color_element(_find_first(cs, cn))
        fs = _find_first(doc, "fontScheme")
        if fs:
            info["font_scheme_name"] = _attr(fs, "name", "")
            for key, tag in [("major","majorFont"),("minor","minorFont")]:
                f = _find_first(fs, tag)
                if f:
                    l = _find_first(f, "latin")
                    info["fonts"][key] = _attr(l, "typeface", "") if l else ""
        themes.append(info)
    return themes


def inspect_masters_and_layouts(unpacked_dir):
    masters_dir = unpacked_dir / "ppt" / "slideMasters"
    layouts_dir = unpacked_dir / "ppt" / "slideLayouts"
    masters = []
    if not masters_dir.exists():
        return masters
    for mf in sorted(masters_dir.glob("slideMaster*.xml")):
        mi = {"file": mf.name, "layouts": [], "background": None, "media": []}
        rels = inspect_rels(masters_dir / "_rels" / f"{mf.name}.rels")
        layout_files = [r["target"].split("/")[-1] for r in rels if r["type"] == "slideLayout"]
        for r in rels:
            if r["type"] in ("image", "oleObject"):
                mi["media"].append(r["target"])
        try:
            doc = parse_xml(str(mf))
            if _find_first(doc, "bg"):
                mi["background"] = "defined"
        except Exception:
            pass
        for lf in sorted(layout_files):
            li = {"file": lf, "name": "", "placeholders": [], "used_by_slides": []}
            lp = layouts_dir / lf
            if lp.exists():
                try:
                    ld = parse_xml(str(lp))
                    c = _find_first(ld, "cSld")
                    if c:
                        li["name"] = _attr(c, "name", "")
                    for ph in _find_all(ld, "ph"):
                        li["placeholders"].append({"type": _attr(ph, "type", "body"), "idx": _attr(ph, "idx", "")})
                except Exception:
                    pass
            mi["layouts"].append(li)
        masters.append(mi)
    return masters


def inspect_slides(unpacked_dir):
    sd = unpacked_dir / "ppt" / "slides"
    slides = []
    if not sd.exists():
        return slides
    for sf in sorted(sd.glob("slide*.xml")):
        if sf.name.startswith("slide") and sf.suffix == ".xml":
            info = {"file": sf.name, "layout": None}
            for r in inspect_rels(sd / "_rels" / f"{sf.name}.rels"):
                if r["type"] == "slideLayout":
                    info["layout"] = r["target"].split("/")[-1]; break
            slides.append(info)
    return slides


def do_inspect(unpacked_dir):
    results = {"unpacked_dir": str(unpacked_dir), "themes": inspect_theme(unpacked_dir),
               "masters": inspect_masters_and_layouts(unpacked_dir), "slides": inspect_slides(unpacked_dir)}
    sl_map = {}
    for s in results["slides"]:
        if s["layout"]:
            sl_map.setdefault(s["layout"], []).append(s["file"])
    for m in results["masters"]:
        for l in m["layouts"]:
            l["used_by_slides"] = sl_map.get(l["file"], [])
    return results


def print_report(r):
    print("=" * 70)
    print(f"PPTX Template Inspection: {r['unpacked_dir']}")
    print("=" * 70)
    for t in r["themes"]:
        print(f"\nTheme: {t.get('name','N/A')} ({t['file']})")
        if "error" in t:
            print(f"  ERROR: {t['error']}"); continue
        print(f"  Color Scheme: {t.get('color_scheme_name','N/A')}")
        for cn, cv in t.get("colors", {}).items():
            print(f"    {cn:12s} = {cv}")
        print(f"  Font Scheme: {t.get('font_scheme_name','N/A')}")
        for ft, fv in t.get("fonts", {}).items():
            print(f"    {ft:12s} = {fv}")
    print(f"\nSlide Masters: {len(r['masters'])}")
    for m in r["masters"]:
        print(f"\n  {m['file']}")
        print(f"    Layouts ({len(m['layouts'])}):")
        for l in m["layouts"]:
            used = l.get("used_by_slides", [])
            ph = [p["type"] for p in l["placeholders"]]
            print(f'      {l["file"]:30s} name="{l["name"]}"')
            if ph:
                print(f"        Placeholders: {', '.join(ph)}")
            if used:
                print(f"        Used by: {', '.join(used)}")
    print(f"\nSlides: {len(r['slides'])}")
    for s in r["slides"]:
        print(f"  {s['file']:20s} -> {s['layout']}")
    print()


# ============================================================================
# EXTRACT THEME
# ============================================================================

def extract_theme(source_dir, output_dir):
    output_dir = Path(output_dir); output_dir.mkdir(parents=True, exist_ok=True)
    theme_src = Path(source_dir) / "ppt" / "theme"
    manifest = {"themes": [], "media_files": []}
    themes_out = output_dir / "theme"; themes_out.mkdir(exist_ok=True)
    for tf in sorted(theme_src.glob("theme*.xml")):
        shutil.copy2(tf, themes_out / tf.name); print(f"  Copied {tf.name}")
        rs = theme_src / "_rels" / f"{tf.name}.rels"
        if rs.exists():
            ro = themes_out / "_rels"; ro.mkdir(exist_ok=True)
            shutil.copy2(rs, ro / rs.name)
            for r in inspect_rels(rs):
                if r["type"] in ("image","oleObject","audio","video"):
                    mn = Path(r["target"]).name
                    mp = Path(source_dir) / "ppt" / "media" / mn
                    if mp.exists():
                        mo = output_dir / "media"; mo.mkdir(exist_ok=True)
                        shutil.copy2(mp, mo / mn); manifest["media_files"].append(mn)
    for ti in inspect_theme(Path(source_dir)):
        manifest["themes"].append({k: ti.get(k, "") for k in
            ["file","name","colors","fonts","color_scheme_name","font_scheme_name"]})
    (output_dir / "manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    print("  Wrote manifest.json")
    return manifest


# ============================================================================
# APPLY THEME
# ============================================================================

def apply_theme(bundle_dir, target_dir):
    bd, td = Path(bundle_dir), Path(target_dir)
    manifest = json.loads((bd / "manifest.json").read_text(encoding="utf-8"))
    ttd = td / "ppt" / "theme"; ttd.mkdir(parents=True, exist_ok=True)
    btd = bd / "theme"
    if btd.exists():
        for tf in sorted(btd.glob("theme*.xml")):
            shutil.copy2(tf, ttd / tf.name); print(f"  Replaced {tf.name}")
        brd = btd / "_rels"
        if brd.exists():
            trd = ttd / "_rels"; trd.mkdir(exist_ok=True)
            for rf in brd.glob("*.rels"):
                shutil.copy2(rf, trd / rf.name)
    bmd = bd / "media"; new_media = []
    if bmd.exists():
        tmd = td / "ppt" / "media"; tmd.mkdir(parents=True, exist_ok=True)
        for mf in bmd.iterdir():
            shutil.copy2(mf, tmd / mf.name); new_media.append(mf.name)
    if new_media:
        _update_content_types_for_media(td, new_media)
    print("  Theme applied successfully.")


def _update_content_types_for_media(target_dir, new_media_files):
    ct_path = Path(target_dir) / "[Content_Types].xml"
    if not ct_path.exists():
        return
    doc = parse_xml(str(ct_path))
    existing = {_attr(e, "Extension", "").lower() for e in _find_all(doc, "Default")}
    ext_ct = {"png":"image/png","jpg":"image/jpeg","jpeg":"image/jpeg","gif":"image/gif",
              "bmp":"image/bmp","emf":"image/x-emf","wmf":"image/x-wmf","svg":"image/svg+xml"}
    for mf in new_media_files:
        ext = Path(mf).suffix.lstrip(".").lower()
        if ext and ext not in existing and ext in ext_ct:
            d = doc.createElement("Default")
            d.setAttribute("Extension", ext); d.setAttribute("ContentType", ext_ct[ext])
            doc.documentElement.appendChild(d); existing.add(ext)
    _write_xml(doc, ct_path)


# ============================================================================
# TRANSFER MASTERS AND LAYOUTS
# ============================================================================

def transfer_masters_and_layouts(source_dir, target_dir, mode="replace"):
    sd, td = Path(source_dir), Path(target_dir)
    smd, sld = sd/"ppt"/"slideMasters", sd/"ppt"/"slideLayouts"
    tmd, tld = td/"ppt"/"slideMasters", td/"ppt"/"slideLayouts"

    layout_offset = 0; master_offset = 0
    if mode == "replace":
        print("  Mode: REPLACE - removing existing masters and layouts")
        for d, p in [(tmd,"slideMaster*.xml"),(tld,"slideLayout*.xml")]:
            if d.exists():
                for f in d.glob(p): f.unlink()
                rd = d / "_rels"
                if rd.exists():
                    for f in rd.glob(p + ".rels"): f.unlink()
    else:
        print("  Mode: MERGE")
        master_offset = max([_extract_number(f.name) for f in tmd.glob("slideMaster*.xml")] or [0])
        layout_offset = max([_extract_number(f.name) for f in tld.glob("slideLayout*.xml")] or [0])

    for d in [tmd, tld, tmd/"_rels", tld/"_rels"]:
        d.mkdir(parents=True, exist_ok=True)

    # Build maps
    src_master_files = sorted(smd.glob("slideMaster*.xml"))
    all_src_layouts = set()
    for mf in src_master_files:
        rp = smd / "_rels" / f"{mf.name}.rels"
        if rp.exists():
            for r in _find_all(parse_xml(str(rp)), "Relationship"):
                t = _attr(r, "Target", "")
                if "slideLayout" in t:
                    all_src_layouts.add(Path(t).name)
    if sld.exists():
        for lf in sld.glob("slideLayout*.xml"):
            all_src_layouts.add(lf.name)

    n = layout_offset + 1
    layout_map = {}
    for ln in sorted(all_src_layouts, key=_extract_number):
        layout_map[ln] = f"slideLayout{n}.xml"; n += 1

    n = master_offset + 1
    master_map = {}
    for mf in src_master_files:
        master_map[mf.name] = f"slideMaster{n}.xml"; n += 1

    print(f"  Masters to copy: {len(master_map)}")
    print(f"  Layouts to copy: {len(layout_map)}")

    # Copy layouts
    all_media = set()
    for old, new in layout_map.items():
        sp = sld / old
        if not sp.exists():
            continue
        shutil.copy2(sp, tld / new)
        sr = sld / "_rels" / f"{old}.rels"
        if sr.exists():
            rd = parse_xml(str(sr))
            for rel in _find_all(rd, "Relationship"):
                t = _attr(rel, "Target", "")
                rt = _attr(rel, "Type", "").split("/")[-1]
                for om, nm in master_map.items():
                    if om in t:
                        rel.setAttribute("Target", t.replace(om, nm)); break
                if rt in ("image","oleObject","audio","video"):
                    all_media.add(Path(t).name)
            _write_xml(rd, tld / "_rels" / f"{new}.rels")
        print(f"  Copied layout {old} -> {new}")

    # Copy masters
    for old, new in master_map.items():
        sp = smd / old
        shutil.copy2(sp, tmd / new)
        sr = smd / "_rels" / f"{old}.rels"
        if sr.exists():
            rd = parse_xml(str(sr))
            for rel in _find_all(rd, "Relationship"):
                t = _attr(rel, "Target", "")
                rt = _attr(rel, "Type", "").split("/")[-1]
                for ol, nl in layout_map.items():
                    if ol in t:
                        rel.setAttribute("Target", t.replace(ol, nl)); break
                if rt in ("image","oleObject","audio","video"):
                    all_media.add(Path(t).name)
            _write_xml(rd, tmd / "_rels" / f"{new}.rels")
        print(f"  Copied master {old} -> {new}")

    # Copy media
    sm = sd / "ppt" / "media"; tm = td / "ppt" / "media"
    if all_media and sm.exists():
        tm.mkdir(parents=True, exist_ok=True)
        for mn in all_media:
            mp = sm / mn
            if mp.exists():
                shutil.copy2(mp, tm / mn)

    # Update presentation.xml
    pp = td / "ppt" / "presentation.xml"
    if pp.exists():
        pdoc = parse_xml(str(pp))
        mil = _find_all(pdoc, "sldMasterIdLst")
        mil = mil[0] if mil else None
        if mil is None:
            mil = pdoc.createElement("p:sldMasterIdLst")
            sil = _find_all(pdoc, "sldIdLst")
            if sil:
                pdoc.documentElement.insertBefore(mil, sil[0])
            else:
                pdoc.documentElement.appendChild(mil)
        if mode == "replace":
            while mil.firstChild:
                mil.removeChild(mil.firstChild)

        existing_ids = {int(_attr(m, "id", "0")) for m in _find_all(pdoc, "sldMasterId") if _attr(m, "id")}
        next_id = (max(existing_ids) + 1) if existing_ids else 2147483648

        prp = td / "ppt" / "_rels" / "presentation.xml.rels"
        prd = parse_xml(str(prp)) if prp.exists() else parse_xml(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')

        existing_rids = {int(m.group(1)) for r in _find_all(prd, "Relationship")
                         if (m := re.search(r"(\d+)", _attr(r, "Id", "")))}
        next_rid = (max(existing_rids) + 1) if existing_rids else 1

        if mode == "replace":
            for rel in list(_find_all(prd, "Relationship")):
                if "slideMaster" in _attr(rel, "Type", ""):
                    rel.parentNode.removeChild(rel)

        for old, new in sorted(master_map.items(), key=lambda x: _extract_number(x[1])):
            rid = f"rId{next_rid}"; next_rid += 1
            me = pdoc.createElement("p:sldMasterId")
            me.setAttribute("id", str(next_id)); me.setAttribute("r:id", rid)
            mil.appendChild(me); next_id += 1
            re2 = prd.createElement("Relationship")
            re2.setAttribute("Id", rid)
            re2.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster")
            re2.setAttribute("Target", f"slideMasters/{new}")
            prd.documentElement.appendChild(re2)
            print(f"  Added master reference: {rid} -> {new}")

        _write_xml(pdoc, pp); _write_xml(prd, prp)

    # Update [Content_Types].xml
    ctp = td / "[Content_Types].xml"
    if ctp.exists():
        ctd = parse_xml(str(ctp))
        existing_ov = {_attr(o, "PartName", "") for o in _find_all(ctd, "Override")}
        mct = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"
        lct = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
        for new in master_map.values():
            pn = f"/ppt/slideMasters/{new}"
            if pn not in existing_ov:
                o = ctd.createElement("Override"); o.setAttribute("PartName", pn); o.setAttribute("ContentType", mct)
                ctd.documentElement.appendChild(o)
        for new in layout_map.values():
            pn = f"/ppt/slideLayouts/{new}"
            if pn not in existing_ov:
                o = ctd.createElement("Override"); o.setAttribute("PartName", pn); o.setAttribute("ContentType", lct)
                ctd.documentElement.appendChild(o)
        if mode == "replace":
            for ov in list(_find_all(ctd, "Override")):
                pn = _attr(ov, "PartName", "")
                if ("/slideMasters/" in pn or "/slideLayouts/" in pn):
                    fn = Path(pn).name
                    if fn not in master_map.values() and fn not in layout_map.values():
                        if not (td / pn.lstrip("/")).exists():
                            ov.parentNode.removeChild(ov)
        _write_xml(ctd, ctp)
        print("  Updated [Content_Types].xml")

    print(f"\n  Transfer complete: {len(master_map)} masters, {len(layout_map)} layouts")
    return master_map, layout_map


# ============================================================================
# MAP LAYOUTS (with content-aware analysis)
# ============================================================================

SYNONYMS = {
    "two content": {"two column","two_content","2 content","2 column"},
    "two column": {"two content","two_content","2 content","2 column"},
    "title only": {"title_only","titleonly"},
    "title slide": {"title_slide","titleslide"},
    "section header": {"section_header","sectionheader","section title"},
    "comparison": {"compare"},
    "blank": {"empty"},
}
GENERIC_NAMES = {"default","blank","empty","custom","custom layout","layout",
                 "1_default","2_default","default design"}
LAYOUT_KEYWORDS = {
    "title": ["title slide","cover","title page","branded_title","branded title"],
    "content": ["content","body","text","title and content","branded_content","branded content"],
    "section": ["section","divider","break","header"],
    "blank": ["blank","empty"],
}
TITLE_EXCLUDES = ["title and content","title, content","title and body","title only","title and text"]


def _normalize(name):
    return name.strip().lower().replace("_"," ").replace("-"," ")


def _get_layout_info(layouts_dir):
    layouts = {}
    if not layouts_dir.exists():
        return layouts
    for lf in sorted(layouts_dir.glob("slideLayout*.xml")):
        info = {"file": lf.name, "name": "", "placeholder_types": set()}
        try:
            doc = parse_xml(str(lf))
            c = _find_all(doc, "cSld")
            if c: info["name"] = _attr(c[0], "name", "")
            for ph in _find_all(doc, "ph"):
                info["placeholder_types"].add(_attr(ph, "type", "body"))
        except Exception:
            pass
        layouts[lf.name] = info
    return layouts


def _get_slides_with_layouts(slides_dir):
    slides = []
    if not slides_dir.exists():
        return slides
    for sf in sorted(slides_dir.glob("slide*.xml")):
        if sf.name.startswith("slide") and sf.suffix == ".xml":
            info = {"file": sf.name, "layout": None, "layout_name": ""}
            rp = slides_dir / "_rels" / f"{sf.name}.rels"
            if rp.exists():
                try:
                    rd = parse_xml(str(rp))
                    for r in _find_all(rd, "Relationship"):
                        if "slideLayout" in _attr(r, "Type", ""):
                            info["layout"] = Path(_attr(r, "Target", "")).name; break
                except Exception:
                    pass
            slides.append(info)
    return slides


def _analyze_slide_content(slide_path):
    result = {"text_boxes": 0, "max_font_size": 0, "has_title_like": False,
              "has_body_text": False, "has_images": False,
              "is_first_slide": slide_path.name == "slide1.xml", "inferred_type": "unknown"}
    try:
        doc = parse_xml(str(slide_path))
    except Exception:
        return result

    shapes = []
    for sp in _find_all(doc, "sp"):
        si = {"has_text": False, "text": "", "max_font_size": 0, "top": 0, "is_centered_h": False}
        for off in _find_all(sp, "off"):
            try: si["top"] = int(_attr(off, "y", "0"))
            except: pass
        for ext in _find_all(sp, "ext"):
            try:
                if int(_attr(ext, "cx", "0")) > 7_000_000: si["is_centered_h"] = True
            except: pass
        for txBody in _find_all(sp, "txBody"):
            txt = ""
            for p in _find_all(txBody, "p"):
                for r in _find_all(p, "r"):
                    for c in r.childNodes:
                        if (c.localName or getattr(c, 'tagName', '').split(":")[-1] if hasattr(c, 'tagName') else '') == "t":
                            txt += c.firstChild.nodeValue if c.firstChild else ""
                    for rPr in _find_all(r, "rPr"):
                        sz = _attr(rPr, "sz", "")
                        if sz:
                            try:
                                fs = int(sz)
                                si["max_font_size"] = max(si["max_font_size"], fs)
                                result["max_font_size"] = max(result["max_font_size"], fs)
                            except: pass
            si["text"] = txt.strip()
            if txt.strip(): si["has_text"] = True
        if si["has_text"]: shapes.append(si)

    if _find_all(doc, "pic"): result["has_images"] = True
    result["text_boxes"] = len(shapes)

    title_shapes = []
    for s in shapes:
        if s["max_font_size"] >= 2400 and s["top"] < 2_500_000:
            title_shapes.append(s); result["has_title_like"] = True
        elif s["has_text"] and s["max_font_size"] < 2400:
            result["has_body_text"] = True

    n = result["text_boxes"]
    if n == 0:
        result["inferred_type"] = "picture" if result["has_images"] else "blank"
    elif result["is_first_slide"] and n <= 2 and result["has_title_like"]:
        result["inferred_type"] = "title"
    elif result["has_title_like"] and result["has_body_text"]:
        result["inferred_type"] = "content"
    elif n <= 2 and result["has_title_like"] and any(s["is_centered_h"] for s in title_shapes):
        result["inferred_type"] = "section"
    elif result["has_title_like"]:
        result["inferred_type"] = "content"
    else:
        result["inferred_type"] = "content"
    return result


def _find_layout_for_type(inferred_type, source_layouts, template_dir=None):
    if inferred_type == "unknown":
        return None, "none", "none"
    keywords = LAYOUT_KEYWORDS.get(inferred_type, [])
    excludes = TITLE_EXCLUDES if inferred_type == "title" else []
    for fn, info in source_layouts.items():
        ln = _normalize(info["name"])
        if not ln: continue
        if any(ex in ln for ex in excludes): continue
        if any(kw in ln for kw in keywords):
            return fn, "content_analysis", "high"
    # Fallback: section/picture → content keywords
    if inferred_type in ("section", "picture"):
        for fn, info in source_layouts.items():
            ln = _normalize(info["name"])
            if not ln: continue
            if any(kw in ln for kw in LAYOUT_KEYWORDS.get("content", [])):
                return fn, "content_analysis_fallback", "medium"
    # Positional fallback from template usage
    if template_dir:
        ts = _get_slides_with_layouts(Path(template_dir) / "ppt" / "slides")
        target = "slide1.xml" if inferred_type == "title" else None
        for s in ts:
            if target and s["file"] == target and s["layout"]:
                return s["layout"], "positional_fallback", "medium"
            if not target and s["file"] != "slide1.xml" and s["layout"]:
                return s["layout"], "positional_fallback", "medium"
    return None, "none", "none"


def map_layouts(source_dir, target_dir):
    sd, td = Path(source_dir), Path(target_dir)
    src_layouts = _get_layout_info(sd / "ppt" / "slideLayouts")
    tgt_layouts = _get_layout_info(td / "ppt" / "slideLayouts")
    tgt_slides = _get_slides_with_layouts(td / "ppt" / "slides")

    for s in tgt_slides:
        if s["layout"] and s["layout"] in tgt_layouts:
            s["layout_name"] = tgt_layouts[s["layout"]]["name"]

    fallback = None
    for fn, info in src_layouts.items():
        if fallback is None: fallback = fn
        if _normalize(info["name"]) in ("blank", "empty"):
            fallback = fn; break

    mappings = []
    for slide in tgt_slides:
        cl = slide["layout"] or ""; cn = slide.get("layout_name", "")
        cph = tgt_layouts.get(cl, {}).get("placeholder_types", set())

        best_match, best_score, best_type, best_conf = None, -1, "none", "none"

        # Name-based matching
        for sfn, sinfo in src_layouts.items():
            sn = _normalize(sinfo["name"]); tn = _normalize(cn)
            if sn == tn:
                sc, mt, co = 100, "exact_name", "high"
            elif tn in SYNONYMS.get(sn, set()) or sn in SYNONYMS.get(tn, set()):
                sc, mt, co = 90, "synonym_name", "high"
            elif sn and tn and (sn in tn or tn in sn):
                sc, mt, co = 70, "partial_name", "medium"
            else:
                sph = sinfo.get("placeholder_types", set())
                if sph and cph:
                    j = len(sph & cph) / len(sph | cph) if (sph | cph) else 0
                    if j > 0.7: sc, mt, co = int(60*j), "placeholder_match", "medium"
                    elif j > 0.3: sc, mt, co = int(40*j), "placeholder_partial", "low"
                    else: sc, mt, co = 0, "none", "none"
                else:
                    sc, mt, co = 0, "none", "none"
            if sc > best_score:
                best_match, best_score, best_type, best_conf = sfn, sc, mt, co

        # Content-aware analysis if generic match
        use_ca = best_score <= 0 or _normalize(cn) in GENERIC_NAMES or (not cn)
        if not use_ca and best_type in ("exact_name","synonym_name"):
            mn = _normalize(src_layouts.get(best_match, {}).get("name", ""))
            if mn in GENERIC_NAMES: use_ca = True

        if use_ca:
            sp = td / "ppt" / "slides" / slide["file"]
            a = _analyze_slide_content(sp)
            if a["inferred_type"] != "unknown":
                cm, ct, cc = _find_layout_for_type(a["inferred_type"], src_layouts, sd)
                if cm:
                    # When triggered by generic layout, always prefer CA over generic match
                    best_match, best_score, best_type, best_conf = cm, 80, ct, cc

        if best_score <= 0 or best_match is None:
            if slide["file"] == "slide1.xml":
                fm, ft, fc = _find_layout_for_type("title", src_layouts, sd)
            else:
                fm, ft, fc = _find_layout_for_type("content", src_layouts, sd)
            best_match = fm or fallback
            best_type, best_conf = "fallback", "low"

        mappings.append({
            "target_slide": slide["file"], "current_layout": cl,
            "current_layout_name": cn, "suggested_layout": best_match,
            "suggested_layout_name": src_layouts.get(best_match, {}).get("name", ""),
            "match_type": best_type, "confidence": best_conf})

    return {"mappings": mappings, "unmapped_slides": [],
            "available_source_layouts": [{"file": fn, "name": info["name"]}
                                          for fn, info in sorted(src_layouts.items())]}


# ============================================================================
# REMAP SLIDES
# ============================================================================

def remap_slides(target_dir, mapping_path):
    td = Path(target_dir)
    data = json.loads(Path(mapping_path).read_text(encoding="utf-8"))
    remapped = 0
    for entry in data.get("mappings", []):
        sf, nl = entry["target_slide"], entry.get("suggested_layout")
        if not nl: continue
        rp = td / "ppt" / "slides" / "_rels" / f"{sf}.rels"
        if not rp.exists(): continue
        try:
            rd = parse_xml(str(rp)); updated = False
            for rel in _find_all(rd, "Relationship"):
                if "slideLayout" in _attr(rel, "Type", ""):
                    old = _attr(rel, "Target", ""); new = f"../slideLayouts/{nl}"
                    if old != new:
                        rel.setAttribute("Target", new); updated = True
                        print(f"  {sf}: {old} -> {new}")
                    break
            if updated:
                _write_xml(rd, rp); remapped += 1
            else:
                print(f"  {sf}: already using correct layout")
        except Exception as e:
            print(f"  Error processing {sf}: {e}", file=sys.stderr)
    print(f"\nRemapped {remapped}/{len(data.get('mappings', []))} slides")


# ============================================================================
# ADAPT TEXT COLORS
# ============================================================================

def _get_solid_fill_color(parent):
    for sf in _find_all(parent, "solidFill"):
        for c in sf.childNodes:
            if getattr(c, "nodeType", None) != 1: continue
            if (c.localName or c.tagName.split(":")[-1]) == "srgbClr":
                return _attr(c, "val", "")
    return None


def _get_bg_color_from_xml(doc):
    for bg in _find_all(doc, "bg"):
        for bgPr in _find_all(bg, "bgPr"):
            c = _get_solid_fill_color(bgPr)
            if c: return c
        for bgRef in _find_all(bg, "bgRef"):
            c = _get_solid_fill_color(bgRef)
            if c: return c
    return None


def get_effective_bg(slide_path, unpacked_dir):
    ud = Path(unpacked_dir)
    try:
        c = _get_bg_color_from_xml(parse_xml(str(slide_path)))
        if c: return c
    except: pass
    # Check layout
    lf = None
    rp = ud/"ppt"/"slides"/"_rels"/f"{slide_path.name}.rels"
    if rp.exists():
        try:
            for r in _find_all(parse_xml(str(rp)), "Relationship"):
                if "slideLayout" in _attr(r, "Type", ""):
                    lf = Path(_attr(r, "Target", "")).name; break
        except: pass
    if lf:
        lp = ud/"ppt"/"slideLayouts"/lf
        if lp.exists():
            try:
                c = _get_bg_color_from_xml(parse_xml(str(lp)))
                if c: return c
            except: pass
            # Check master
            mf = None
            lr = ud/"ppt"/"slideLayouts"/"_rels"/f"{lf}.rels"
            if lr.exists():
                try:
                    for r in _find_all(parse_xml(str(lr)), "Relationship"):
                        if "slideMaster" in _attr(r, "Type", ""):
                            mf = Path(_attr(r, "Target", "")).name; break
                except: pass
            if mf:
                mp = ud/"ppt"/"slideMasters"/mf
                if mp.exists():
                    try:
                        c = _get_bg_color_from_xml(parse_xml(str(mp)))
                        if c: return c
                    except: pass
    return "FFFFFF"


def adapt_text_colors(unpacked_dir):
    ud = Path(unpacked_dir); sd = ud / "ppt" / "slides"
    if not sd.exists(): return
    total = 0
    for sp in sorted(sd.glob("slide*.xml")):
        if not sp.name.startswith("slide") or sp.suffix != ".xml": continue
        bg = get_effective_bg(sp, ud); bg_dark = _is_dark(bg)
        print(f"  {sp.name}: bg=#{bg} ({'dark' if bg_dark else 'light'})")
        try:
            doc = parse_xml(str(sp)); modified = False
            for rPr in _find_all(doc, "rPr"):
                for sf in _find_all(rPr, "solidFill"):
                    for c in sf.childNodes:
                        if getattr(c, "nodeType", None) != 1: continue
                        if (c.localName or c.tagName.split(":")[-1]) != "srgbClr": continue
                        old = _attr(c, "val", "")
                        if not old or len(old) != 6: continue
                        td2 = _is_dark(old)
                        new = None
                        if bg_dark and td2: new = "FFFFFF"
                        elif not bg_dark and not td2: new = "333333"
                        if new and new.upper() != old.upper():
                            c.setAttribute("val", new); modified = True; total += 1
                            print(f"    Flipped #{old} -> #{new}")
            if modified: _write_xml(doc, sp)
        except Exception as e:
            print(f"  Warning: {sp.name}: {e}")
    if total == 0: print("  No color adaptations needed.")
    else: print(f"\n  Adapted {total} text color(s)")


# ============================================================================
# RECONCILE RIDS
# ============================================================================

def reconcile_rids(unpacked_dir):
    ud = Path(unpacked_dir); total = 0
    rels_files = list(ud.rglob("*.rels"))
    print(f"Scanning {len(rels_files)} .rels files...")
    for rp in sorted(rels_files):
        try:
            rd = parse_xml(str(rp))
        except: continue
        rels = _find_all(rd, "Relationship")
        if not rels: continue
        seen, dups = {}, []
        for r in rels:
            rid = _attr(r, "Id", "")
            if rid in seen: dups.append(r)
            else: seen[rid] = r
        if not dups: continue
        max_n = max([int(m.group(1)) for r in rels if (m := re.search(r"(\d+)", _attr(r, "Id", "")))] or [0])
        nn = max_n + 1; renames = {}
        for r in dups:
            old = _attr(r, "Id", ""); new = f"rId{nn}"; nn += 1
            r.setAttribute("Id", new); renames[old] = new; total += 1
        _write_xml(rd, rp)
        # Update corresponding XML
        xn = rp.name[:-5] if rp.name.endswith(".rels") else None
        if xn:
            xp = rp.parent.parent / xn
            if xp.exists() and renames:
                try:
                    xd = parse_xml(str(xp)); changed = False
                    for e in xd.getElementsByTagName("*"):
                        for i in range(e.attributes.length):
                            a = e.attributes.item(i)
                            if a.value in renames:
                                e.setAttribute(a.name, renames[a.value]); changed = True
                    if changed: _write_xml(xd, xp)
                except: pass
    if total == 0: print("No relationship ID collisions found.")
    else: print(f"Fixed {total} collisions.")


# ============================================================================
# CLEAN ORPHANS
# ============================================================================

def clean_orphans(unpacked_dir):
    ud = Path(unpacked_dir); referenced = set()
    for rf in ud.rglob("*.rels"):
        try:
            doc = parse_xml(str(rf))
            for r in _find_all(doc, "Relationship"):
                t = _attr(r, "Target", "")
                if t: referenced.add((rf.parent.parent / t).resolve())
        except: pass
    ct = ud / "[Content_Types].xml"
    if ct.exists():
        try:
            doc = parse_xml(str(ct))
            for o in _find_all(doc, "Override"):
                p = _attr(o, "PartName", "")
                if p: referenced.add((ud / p.lstrip("/")).resolve())
        except: pass
    md = ud / "ppt" / "media"; removed = 0
    if md.exists():
        for mf in md.iterdir():
            if mf.resolve() not in referenced:
                found = False
                for xf in ud.rglob("*.xml"):
                    try:
                        if mf.name in xf.read_text(encoding="utf-8", errors="ignore"):
                            found = True; break
                    except: pass
                if not found:
                    for rf in ud.rglob("*.rels"):
                        try:
                            if mf.name in rf.read_text(encoding="utf-8", errors="ignore"):
                                found = True; break
                        except: pass
                if not found:
                    mf.unlink(); removed += 1
    if removed: print(f"  Cleaned {removed} orphaned media files")
    else: print("  No orphaned files found")


# ============================================================================
# ORCHESTRATOR
# ============================================================================

def apply_template(template_pptx, content_pptx, output_pptx, layout_map_path=None):
    tp, cp, op = Path(template_pptx), Path(content_pptx), Path(output_pptx)
    wd = Path(tempfile.mkdtemp(prefix="pptx_tt_"))
    td, cd = wd / "template", wd / "content"
    tbd = wd / "theme_bundle"; mp = Path(layout_map_path) if layout_map_path else wd / "mapping.json"
    ok = False
    try:
        print("\n[1/11] Unpacking...")
        for src, dst in [(tp, td), (cp, cd)]:
            dst.mkdir(parents=True, exist_ok=True)
            with zipfile.ZipFile(str(src), "r") as zf: zf.extractall(str(dst))
            print(f"  {src.name} -> {dst}")

        print("\n[2/11] Inspecting...")
        print_report(do_inspect(td)); print_report(do_inspect(cd))

        print("\n[3/11] Extracting theme...")
        extract_theme(td, tbd)

        print("\n[4/11] Applying theme...")
        apply_theme(tbd, cd)

        print("\n[5/11] Mapping layouts...")
        if layout_map_path and Path(layout_map_path).exists():
            print(f"  Using provided: {layout_map_path}")
            mr = json.loads(Path(layout_map_path).read_text(encoding="utf-8"))
            mp = Path(layout_map_path)
        else:
            mr = map_layouts(td, cd)
            mp.write_text(json.dumps(mr, indent=2), encoding="utf-8")
            h = sum(1 for m in mr["mappings"] if m["confidence"] == "high")
            md2 = sum(1 for m in mr["mappings"] if m["confidence"] == "medium")
            lo = sum(1 for m in mr["mappings"] if m["confidence"] == "low")
            print(f"  {len(mr['mappings'])} slides: {h} high, {md2} medium, {lo} low confidence")

        print("\n[6/11] Transferring masters & layouts...")
        _mm, lrm = transfer_masters_and_layouts(td, cd, mode="replace")

        print("\n[7/11] Translating layout mapping...")
        for e in mr.get("mappings", []):
            old = e.get("suggested_layout", "")
            if old in lrm:
                e["suggested_layout"] = lrm[old]
                if old != lrm[old]: print(f"  {e['target_slide']}: {old} -> {lrm[old]}")
        mp2 = wd / "mapping_final.json"
        mp2.write_text(json.dumps(mr, indent=2), encoding="utf-8")

        print("\n[8/11] Remapping slides...")
        remap_slides(cd, mp2)

        print("\n[9/11] Adapting text colors...")
        adapt_text_colors(cd)

        print("\n[10/11] Reconciling rIds...")
        reconcile_rids(cd)

        print("\n[11/11] Cleaning & packing...")
        clean_orphans(cd)
        with zipfile.ZipFile(str(op), "w", zipfile.ZIP_DEFLATED) as zf:
            for fp in sorted(cd.rglob("*")):
                if fp.is_file(): zf.write(str(fp), str(fp.relative_to(cd)))
        print(f"  Packed -> {op}")
        ok = True
        print(f"\nDone! Output: {op}")
    except Exception as e:
        import traceback; traceback.print_exc()
        print(f"\nFailed. Work dir preserved: {wd}", file=sys.stderr)
        sys.exit(1)
    finally:
        if ok:
            try: shutil.rmtree(wd)
            except: pass


# ============================================================================
# CLI
# ============================================================================

if __name__ == "__main__":
    p = argparse.ArgumentParser(description="PPTX Template Transfer")
    p.add_argument("template_pptx", type=Path, help="Template PPTX (design source)")
    p.add_argument("content_pptx", type=Path, help="Content PPTX (text/images)")
    p.add_argument("output_pptx", type=Path, help="Output PPTX path")
    p.add_argument("--layout-map", type=Path, default=None, help="Manual layout mapping JSON")
    a = p.parse_args()
    for f, n in [(a.template_pptx, "Template"), (a.content_pptx, "Content")]:
        if not f.exists():
            print(f"Error: {n} not found: {f}", file=sys.stderr); sys.exit(1)
    apply_template(a.template_pptx, a.content_pptx, a.output_pptx, a.layout_map)
