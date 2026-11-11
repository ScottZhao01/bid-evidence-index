#!/usr/bin/env python3
"""Read-only, standard-library DOCX inventory and baseline comparison.

Usage: audit_docx.py source.docx [result.docx] --out report.json
This checks package/XML evidence; it does not render, repair, or update fields.
Exit 0 means the report was written, not that the document is correct.
"""

import argparse
from collections import Counter, defaultdict
from decimal import Decimal, InvalidOperation
import hashlib
import json
import os
from pathlib import Path
import posixpath
import re
import sys
import tempfile
from urllib.parse import unquote, urlsplit
import xml.etree.ElementTree as ET
import zipfile


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "asvg": "http://schemas.microsoft.com/office/drawing/2016/SVG/main",
}
STRICT = {
    "http://purl.oclc.org/ooxml/wordprocessingml/main": NS["w"],
    "http://purl.oclc.org/ooxml/officeDocument/relationships": NS["r"],
    "http://purl.oclc.org/ooxml/drawingml/main": NS["a"],
    "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing": NS["wp"],
    "http://purl.oclc.org/ooxml/drawingml/picture": NS["pic"],
}
STORY_ROOTS = {"document", "hdr", "ftr", "footnotes", "endnotes", "comments", "glossaryDocument"}
LIMITATIONS = [
    "基础审计，仅检查ZIP部件、XML结构和关系；不是投标内容、真实性或评分结论。",
    "未渲染页面，不能证明正文逐字完整、列边界正确、图片可读或未被版面遮挡/裁切。无检测项不等于通过完整性验收。",
    "图片按XML引用出现次数审计，残留但未引用的media不能证明显示保留。AlternateContent的Choice/Fallback及SVG替代引用均列出，不能据此认定它们同时可见。",
    "同一字节哈希被多次引用时按出现序号比较；相同图片插入、删除或调换可能使具体实例的尺寸配对有歧义。次数变化单独报告。",
    "图片顺序在每个XML部件内比较，不等于视觉阅读顺序；浮动定位、文本框、页眉页脚及重命名的story部件需要人工复核。",
    "书签和域按字面解析；不计算域、不更新目录、不验证页码缓存是否最新，不验证外部网址或外部文档的目标。",
    "TOC内容控件通过目录域、docPartGallery或明确目录标签识别；非标准手工目录可能只能作为普通正文块出现。",
    "表格行列、跨度及宽度为存储值，不证明合并单元格或列边界渲染正确。嵌套或修订标记的显示取决于Word设置。",
    "文本摘要包括已存储的w:t（可能含修订、隐藏文字或缓存），不是Word当前可见正文的逐字提取。嵌入的其他文档/对象未递归审计。",
]


def q(prefix, name):
    return "{%s}%s" % (NS[prefix], name)


def local(name):
    return name.rsplit("}", 1)[-1]


def digest(data):
    return hashlib.sha256(data).hexdigest()


def canonical_ns(name):
    if name.startswith("{"):
        uri, tail = name[1:].split("}", 1)
        return "{%s}%s" % (STRICT.get(uri, uri), tail)
    return name


def parse_xml(data):
    root = ET.fromstring(data)
    for e in root.iter():
        e.tag = canonical_ns(e.tag)
        e.attrib = {canonical_ns(k): v for k, v in e.attrib.items()}
    return root


def text_of(element):
    return "".join(e.text or "" for e in element.iter(q("w", "t")))


def summary(text, limit=240):
    return re.sub(r"\s+", " ", text).strip()[:limit]


def number(value):
    if value is None:
        return None
    try:
        return format(Decimal(value).normalize(), "f")
    except InvalidOperation:
        return value


def attributes(element):
    return {} if element is None else {local(k): v for k, v in element.attrib.items()}


def css_length_emu(value):
    if not value:
        return None
    m = re.fullmatch(r"\s*([+-]?[\d.]+)\s*(in|pt|pc|cm|mm|px)?\s*", value, re.I)
    if not m:
        return None
    factors = {"in": 914400, "pt": 12700, "pc": 152400, "cm": 360000, "mm": 36000, "px": 9525}
    # Unitless VML dimensions are not assumed to be points.
    if not m.group(2):
        return None
    try:
        return number(str(Decimal(m.group(1)) * factors[m.group(2).lower()]))
    except InvalidOperation:
        return None


def relationship_owner(name):
    if name == "_rels/.rels":
        return ""
    folder, filename = posixpath.split(name)
    if posixpath.basename(folder) == "_rels" and filename.endswith(".rels"):
        return posixpath.join(posixpath.dirname(folder), filename[:-5])
    return None


def resolve_target(owner, target):
    parsed = urlsplit(target)
    path = unquote(parsed.path)
    if not path:
        resolved = owner
    elif path.startswith("/"):
        resolved = posixpath.normpath(path.lstrip("/"))
    else:
        resolved = posixpath.normpath(posixpath.join(posixpath.dirname(owner), path))
    return resolved, unquote(parsed.fragment)


def parse_instruction(instruction):
    tokens = re.findall(r'"(?:[^"]|"")*"|[^\s]+', instruction.strip())
    tokens = [t[1:-1].replace('""', '"') if t.startswith('"') and t.endswith('"') else t for t in tokens]
    command = tokens[0].upper() if tokens else ""
    result = {"command": command, "target": None, "target_scope": None}
    if command in {"REF", "PAGEREF", "NOTEREF"}:
        # In a valid field the bookmark is the first positional argument.
        i = 1
        while i < len(tokens):
            if tokens[i].startswith("\\"):
                i += 2 if tokens[i] in {"\\*", "\\#", "\\@"} else 1
            else:
                result.update(target=tokens[i], target_scope="document")
                break
    elif command == "HYPERLINK":
        url = tokens[1] if len(tokens) > 1 and not tokens[1].startswith("\\") else ""
        result["external_target"] = url or None
        for i, token in enumerate(tokens[:-1]):
            if token.lower() == "\\l":
                result.update(target=tokens[i + 1], target_scope="external_document" if url else "document")
                break
        if not result["target"] and url.startswith("#"):
            result.update(target=unquote(url[1:]), target_scope="document", external_target=None)
    return result


class Package:
    def __init__(self, filename):
        self.path = Path(filename)
        self.xml_errors = []
        self.parts = {}
        self.roots = {}
        self.relationships = defaultdict(dict)
        self.duplicate_zip_names = []
        self.duplicate_relationship_ids = []
        with zipfile.ZipFile(self.path) as z:
            counts = Counter(i.filename for i in z.infolist() if not i.is_dir())
            self.duplicate_zip_names = [n for n, c in counts.items() if c > 1]
            for info in z.infolist():
                if info.is_dir():
                    continue
                data = z.read(info)
                self.parts[info.filename] = {"size": len(data), "sha256": digest(data)}
                if info.filename.endswith((".xml", ".rels")):
                    try:
                        self.roots[info.filename] = parse_xml(data)
                    except ET.ParseError as exc:
                        self.xml_errors.append({"part": info.filename, "error": str(exc)})
        for name, root in self.roots.items():
            owner = relationship_owner(name)
            if owner is None:
                continue
            for e in root:
                if local(e.tag) != "Relationship":
                    continue
                rid = e.get("Id", "")
                rel = dict(e.attrib)
                rel["external"] = e.get("TargetMode", "").lower() == "external"
                if not rel["external"]:
                    rel["resolved_part"], rel["fragment"] = resolve_target(owner, e.get("Target", ""))
                if rid in self.relationships[owner]:
                    self.duplicate_relationship_ids.append({"owner": owner, "id": rid})
                self.relationships[owner][rid] = rel
        self.main_part = next((r["resolved_part"] for r in self.relationships[""].values()
                               if r.get("Type", "").endswith("/officeDocument") and not r["external"]), "word/document.xml")
        if self.main_part not in self.roots:
            raise ValueError("Main document XML is missing or cannot be parsed: " + self.main_part)

    def image_reference(self, part, rid):
        rel = self.relationships[part].get(rid)
        if rel is None:
            return {"relationship_id": rid, "status": "missing_relationship", "sha256": None}
        if rel["external"]:
            return {"relationship_id": rid, "status": "external_not_checked", "target": rel.get("Target"), "sha256": None}
        target = rel["resolved_part"]
        return {"relationship_id": rid, "status": "resolved" if target in self.parts else "missing_part",
                "target": rel.get("Target"), "resolved_part": target,
                "sha256": self.parts.get(target, {}).get("sha256")}


def geometry(element, ancestors):
    lineage = ancestors + [element]
    drawing = next((e for e in reversed(lineage) if e.tag == q("w", "drawing")), None)
    shape = next((e for e in reversed(lineage) if e.tag in {q("v", "shape"), q("v", "rect"), q("v", "image")}), None)
    result = {"container": None, "frame_emu": None, "transform_emu": None,
              "rotation_degrees": "0", "flip_h": False, "flip_v": False, "crop": {}}
    if drawing is not None:
        frame = next((e for e in drawing.iter() if e.tag in {q("wp", "inline"), q("wp", "anchor")}), drawing)
        result["container"] = local(frame.tag)
        ext = frame.find(q("wp", "extent"))
        if ext is not None:
            result["frame_emu"] = {k: number(ext.get(k)) for k in ("cx", "cy")}
        picture = next((e for e in reversed(lineage) if e.tag == q("pic", "pic")), drawing)
        xfrm = picture.find(".//" + q("a", "xfrm"))
        if xfrm is not None:
            ext = xfrm.find(q("a", "ext"))
            if ext is not None:
                result["transform_emu"] = {k: number(ext.get(k)) for k in ("cx", "cy")}
            try:
                result["rotation_degrees"] = number(str(Decimal(xfrm.get("rot", "0")) / 60000))
            except InvalidOperation:
                result["rotation_degrees"] = xfrm.get("rot")
            result["flip_h"] = xfrm.get("flipH", "0").lower() in {"1", "true", "on"}
            result["flip_v"] = xfrm.get("flipV", "0").lower() in {"1", "true", "on"}
        crop = picture.find(".//" + q("a", "srcRect"))
        if crop is not None:
            result["crop"] = {k: number(v) for k, v in crop.attrib.items() if number(v) != "0"}
    elif shape is not None:
        result["container"] = "vml"
        styles = dict((k.strip().lower(), v.strip()) for k, sep, v in
                      (s.partition(":") for s in shape.get("style", "").split(";")) if sep)
        result["vml_style_dimensions"] = {k: styles.get(k) for k in ("width", "height")}
        result["frame_emu"] = {"cx": css_length_emu(styles.get("width")), "cy": css_length_emu(styles.get("height"))}
        result["rotation_degrees"] = number(styles.get("rotation", "0"))
        result["flip_h"] = "x" in styles.get("flip", "")
        result["flip_v"] = "y" in styles.get("flip", "")
        result["crop"] = {k: v for k, v in element.attrib.items() if local(k).startswith("crop") and v not in {"0", "0f"}}
    return result


def table_info(table):
    props = table.find(q("w", "tblPr"))
    grid = table.find(q("w", "tblGrid"))
    rows = []
    for tr in table.findall(q("w", "tr")):
        cells = []
        for tc in tr.findall(q("w", "tc")):
            pr = tc.find(q("w", "tcPr"))
            span = None if pr is None else pr.find(q("w", "gridSpan"))
            merge = None if pr is None else pr.find(q("w", "vMerge"))
            width = None if pr is None else pr.find(q("w", "tcW"))
            cells.append({"width": attributes(width), "grid_span": span.get(q("w", "val"), "1") if span is not None else "1",
                          "vertical_merge": merge.get(q("w", "val"), "continue") if merge is not None else None})
        rows.append({"physical_cells": len(cells), "cells": cells,
                     "row_properties": [dict(tag=local(e.tag), **attributes(e)) for e in tr.find(q("w", "trPr"))]
                     if tr.find(q("w", "trPr")) is not None else []})
    return {"rows": len(rows), "row_cell_counts": [r["physical_cells"] for r in rows],
            "grid_columns": [] if grid is None else [attributes(e) for e in grid],
            "table_width": attributes(None if props is None else props.find(q("w", "tblW"))),
            "layout": attributes(None if props is None else props.find(q("w", "tblLayout"))), "row_details": rows}


def inventory(filename):
    pkg = Package(filename)
    result = {"file": str(Path(filename).resolve()), "main_part": pkg.main_part,
              "zip_parts": pkg.parts, "xml_parse_errors": pkg.xml_errors,
              "duplicate_zip_names": pkg.duplicate_zip_names,
              "duplicate_relationship_ids": pkg.duplicate_relationship_ids,
              "body_blocks": [], "tables": [], "images": [], "bookmarks": [],
              "bookmark_ends": [], "fields": [], "hyperlinks": [], "content_controls": [], "story_parts": []}
    main = pkg.roots[pkg.main_part]
    body = main.find(q("w", "body"))
    body_index = {id(e): i for i, e in enumerate(body)} if body is not None else {}
    for i, block in enumerate(body if body is not None else []):
        text = text_of(block)
        result["body_blocks"].append({"index": i, "tag": local(block.tag), "summary": summary(text),
            "text_characters": len(text), "stored_text_sha256": digest(text.encode()),
            "normalized_xml_sha256": digest(ET.tostring(block)),
            "paragraph_count": sum(1 for _ in block.iter(q("w", "p"))),
            "table_count": sum(1 for _ in block.iter(q("w", "tbl"))),
            "image_ids": [], "field_ids": [], "hyperlink_ids": [], "bookmark_names": []})

    for part in sorted(pkg.roots, key=lambda s: (s != pkg.main_part, s)):
        root = pkg.roots[part]
        if local(root.tag) not in STORY_ROOTS or not root.tag.startswith("{" + NS["w"] + "}"):
            continue
        result["story_parts"].append(part)
        stack = []
        controls = []

        def location(path, ancestors):
            block = next((body_index[id(e)] for e in ancestors if part == pkg.main_part and id(e) in body_index), None)
            return {"part": part, "path": path, "body_block": block}

        def append_record(key, rec):
            rec["id"] = key.rstrip("s") + ":" + str(len(result[key]))
            result[key].append(rec)
            b = rec.get("body_block")
            if b is not None:
                dest = {"images": "image_ids", "fields": "field_ids", "hyperlinks": "hyperlink_ids"}.get(key)
                if dest:
                    result["body_blocks"][b][dest].append(rec["id"])
            return rec

        def field_finish(rec):
            rec["instruction"] = "".join(rec.pop("_instruction", []))
            cache = "".join(rec.pop("_cached", []))
            rec.update(parse_instruction(rec["instruction"]))
            rec.update(cached_text_present=bool(cache.strip()), cached_text_characters=len(cache), cached_summary=summary(cache))
            if rec["command"] in {"REF", "PAGEREF", "NOTEREF", "HYPERLINK"} and rec.get("target"):
                append_record("hyperlinks", {k: rec[k] for k in ("part", "path", "body_block", "target", "target_scope")} |
                              {"source": "field", "field_id": rec["id"], "command": rec["command"],
                               "external_target": rec.get("external_target")})

        def visit(e, path, ancestors):
            loc = location(path, ancestors + [e])
            if e.tag == q("w", "tbl"):
                result["tables"].append(dict(loc, **table_info(e)))
            if e.tag == q("w", "sdt"):
                pr = e.find(q("w", "sdtPr"))
                vals = {local(x.tag): x.get(q("w", "val")) for x in pr.iter()} if pr is not None else {}
                stored = text_of(e)
                control = dict(loc, properties=vals, is_toc=False, toc_evidence=[],
                               stored_text_present=bool(stored.strip()), stored_text_characters=len(stored),
                               stored_text_summary=summary(stored),
                               stored_hyperlink_count=sum(1 for _ in e.iter(q("w", "hyperlink"))))
                controls.append(control)
                result["content_controls"].append(control)
            if e.tag == q("w", "bookmarkStart"):
                rec = dict(loc, name=e.get(q("w", "name")), bookmark_id=e.get(q("w", "id")))
                result["bookmarks"].append(rec)
                if loc["body_block"] is not None:
                    result["body_blocks"][loc["body_block"]]["bookmark_names"].append(rec["name"])
            elif e.tag == q("w", "bookmarkEnd"):
                result["bookmark_ends"].append(dict(loc, bookmark_id=e.get(q("w", "id"))))
            if e.tag == q("w", "hyperlink"):
                anchor = e.get(q("w", "anchor"))
                rid = e.get(q("r", "id"))
                rel = pkg.relationships[part].get(rid, {})
                target_scope = "document" if anchor else None
                external_target = rel.get("Target") if rel.get("external") else None
                if rel and not rel.get("external") and rel.get("fragment"):
                    anchor = anchor or rel["fragment"]
                    target_scope = "document" if rel.get("resolved_part") in {part, pkg.main_part} else "other_part"
                if external_target:
                    if external_target.startswith("#"):
                        anchor, target_scope, external_target = unquote(external_target[1:]), "document", None
                    elif anchor:
                        target_scope = "external_document"
                append_record("hyperlinks", dict(loc, source="w:hyperlink", target=anchor,
                              target_scope=target_scope, relationship_id=rid, external_target=external_target,
                              text_summary=summary(text_of(e)), relationship_missing=bool(rid and not rel)))
            if e.tag == q("w", "fldSimple"):
                rec = append_record("fields", dict(loc, kind="simple", closed=True, result_region_present=len(e) > 0,
                    _instruction=[e.get(q("w", "instr"), "")], _cached=[text_of(e)]))
                field_finish(rec)
            elif e.tag == q("w", "fldChar"):
                kind = e.get(q("w", "fldCharType"))
                if kind == "begin":
                    stack.append(append_record("fields", dict(loc, kind="complex", closed=False,
                                 result_region_present=False, _instruction=[], _cached=[])))
                elif kind == "separate" and stack:
                    stack[-1]["result_region_present"] = True
                elif kind == "end" and stack:
                    rec = stack.pop()
                    rec["closed"] = True
                    field_finish(rec)
            elif e.tag == q("w", "instrText") and stack and not stack[-1]["result_region_present"]:
                stack[-1]["_instruction"].append(e.text or "")
            elif e.tag == q("w", "t"):
                for rec in stack:
                    if rec["result_region_present"]:
                        rec["_cached"].append(e.text or "")
            image_tags = {q("a", "blip"), q("asvg", "svgBlip"), q("v", "imagedata"), q("v", "fill"), q("v", "stroke"), q("v", "image")}
            if e.tag in image_tags:
                refs = []
                for attr in (q("r", "embed"), q("r", "link"), q("r", "id"), q("o", "relid")):
                    if e.get(attr):
                        refs.append(dict(pkg.image_reference(part, e.get(attr)), attribute=local(attr)))
                direct = e.get("src")
                if direct and not refs:
                    refs.append({"status": "direct_source_not_resolved", "target": direct, "sha256": None})
                # A VML fill without an image reference is a color/gradient, not an image occurrence.
                if refs or e.tag in {q("a", "blip"), q("asvg", "svgBlip"), q("v", "imagedata")}:
                    primary = next((r for r in refs if r.get("sha256")), refs[0] if refs else {})
                    branch = [{"kind": local(a.tag), "requires": a.get("Requires")} for a in ancestors
                              if a.tag in {q("mc", "Choice"), q("mc", "Fallback")}]
                    append_record("images", dict(loc, kind=local(e.tag), references=refs,
                        sha256=primary.get("sha256"), external_target=primary.get("target") if not primary.get("sha256") else None,
                        geometry=geometry(e, ancestors), alternate_content_branches=branch))
            counts = Counter()
            for child in e:
                counts[child.tag] += 1
                visit(child, path + "/" + local(child.tag) + "[" + str(counts[child.tag]) + "]", ancestors + [e])

        visit(root, "/" + local(root.tag) + "[1]", [])
        for rec in stack:
            field_finish(rec)
        for control in controls:
            vals = control["properties"]
            if (vals.get("docPartGallery") or "").strip().casefold() in {"table of contents", "tables of contents", "目录"}:
                control["toc_evidence"].append("docPartGallery")
            if (vals.get("alias") or "").strip().casefold() in {"table of contents", "toc", "目录"}:
                control["toc_evidence"].append("alias")
            if any(f.get("command") == "TOC" and f["part"] == part and f["path"].startswith(control["path"] + "/") for f in result["fields"]):
                control["toc_evidence"].append("TOC_field")
            control["is_toc"] = bool(control["toc_evidence"])

    referenced = Counter(r["resolved_part"] for i in result["images"] for r in i["references"] if r.get("resolved_part"))
    image_targets = set(referenced)
    defaults, overrides = {}, {}
    ct = pkg.roots.get("[Content_Types].xml")
    if ct is not None:
        for e in ct:
            if local(e.tag) == "Default":
                defaults[e.get("Extension", "").lower()] = e.get("ContentType", "")
            elif local(e.tag) == "Override":
                overrides[e.get("PartName", "").lstrip("/")] = e.get("ContentType", "")
    media = []
    for name, info in pkg.parts.items():
        content_type = overrides.get(name, defaults.get(name.rsplit(".", 1)[-1].lower(), ""))
        if "/media/" in "/" + name or name in image_targets or content_type.startswith("image/"):
            media.append(dict(part=name, **info, content_type=content_type,
                              xml_reference_count=referenced[name], referenced_in_xml=bool(referenced[name])))
    result["media"] = media
    names = defaultdict(list)
    for b in result["bookmarks"]:
        if b["name"]:
            names[b["name"]].append(b)
    bad_links = [h for h in result["hyperlinks"] if h.get("target_scope") == "document" and h.get("target") not in names]
    starts = Counter((b["part"], b["bookmark_id"]) for b in result["bookmarks"])
    ends = Counter((b["part"], b["bookmark_id"]) for b in result["bookmark_ends"])
    result["integrity"] = {
        "broken_local_links": bad_links,
        "duplicate_bookmark_names": [{"name": name, "occurrences": rows} for name, rows in names.items() if len(rows) > 1],
        "bookmark_boundary_count_mismatches": [{"part": p, "bookmark_id": bid, "starts": starts[(p, bid)], "ends": ends[(p, bid)]}
                                               for p, bid in sorted(starts.keys() | ends.keys(), key=str) if starts[(p, bid)] != ends[(p, bid)]],
        "duplicate_bookmark_ids_per_part": [{"part": p, "bookmark_id": bid, "count": c} for (p, bid), c in starts.items() if c > 1],
        "unclosed_complex_fields": [f["id"] for f in result["fields"] if not f["closed"]],
        "unresolved_image_references": [{"image_id": i["id"], **r} for i in result["images"] for r in i["references"] if r["status"] in {"missing_relationship", "missing_part", "direct_source_not_resolved"}],
        "image_nodes_without_references": [i["id"] for i in result["images"] if not i["references"]],
        "nonpositive_image_extents": [i["id"] for i in result["images"] if any(v is not None and re.fullmatch(r"-?\d+(?:\.\d+)?", v) and Decimal(v) <= 0 for v in (i["geometry"].get("frame_emu") or {}).values())],
        "hyperlinks_missing_relationships": [h["id"] for h in result["hyperlinks"] if h.get("relationship_missing")],
    }
    toc_fields = [f for f in result["fields"] if f.get("command") == "TOC"]
    toc_controls = [s for s in result["content_controls"] if s["is_toc"]]
    result["toc"] = {"field_count": len(toc_fields), "field_ids": [f["id"] for f in toc_fields],
                     "fields_with_cached_text": sum(f["cached_text_present"] for f in toc_fields),
                     "content_control_count": len(toc_controls), "content_controls": toc_controls,
                     "content_controls_with_cached_or_static_text": sum(s["stored_text_present"] for s in toc_controls),
                     "cache_freshness": "not_evaluated"}
    result["counts"] = {k: len(result[k]) for k in ("body_blocks", "tables", "images", "bookmarks", "fields", "hyperlinks", "media")}
    return result


def image_identity(image):
    if image.get("sha256"):
        return "sha256:" + image["sha256"]
    if image.get("external_target"):
        return "unverified-target:" + image["external_target"]
    return "unresolved-image-node"


def comparable_geometry(g):
    g = dict(g)
    if g.get("frame_emu") and all(v is not None for v in g["frame_emu"].values()):
        g.pop("vml_style_dimensions", None)
    return g


def compare(source, target):
    old, new = source["zip_parts"], target["zip_parts"]
    modified = [n for n in old.keys() & new.keys() if old[n]["sha256"] != new[n]["sha256"]]
    old_hashes = defaultdict(list)
    for m in source["media"]:
        old_hashes[m["sha256"]].append(m["part"])
    # Byte presence is deliberately separate from being an image or being referenced.
    # A renamed part with an unusual extension still retains the bytes, but cannot
    # satisfy the per-occurrence reference checks just by remaining in the ZIP.
    new_hashes = {info["sha256"] for info in target["zip_parts"].values()}
    old_names = Counter(b["name"] for b in source["bookmarks"] if b["name"])
    new_names = Counter(b["name"] for b in target["bookmarks"] if b["name"])
    per_part = []
    for part in sorted(set(source["story_parts"]) | set(target["story_parts"])):
        before = [i for i in source["images"] if i["part"] == part]
        after = [i for i in target["images"] if i["part"] == part]
        if not before and not after:
            continue
        a, b = [image_identity(i) for i in before], [image_identity(i) for i in after]
        ca, cb = Counter(a), Counter(b)
        missing = [{"identity": key, "source_count": ca[key], "result_count": cb[key], "missing_count": ca[key] - cb[key]}
                   for key in ca if ca[key] > cb[key]]
        added = [{"identity": key, "source_count": ca[key], "result_count": cb[key], "added_count": cb[key] - ca[key]}
                 for key in cb if cb[key] > ca[key]]
        # Relative order is assessed using the common multiset. Duplicate identities
        # have no intrinsic instance ID; the limitation is deliberately explicit.
        common = ca & cb
        def projection(items):
            count = Counter()
            out = []
            for key in items:
                count[key] += 1
                if count[key] <= common[key]:
                    out.append((key, count[key]))
            return out
        geometry_changes = []
        groups = defaultdict(list)
        for i in after:
            groups[image_identity(i)].append(i)
        seen = Counter()
        for i in before:
            key = image_identity(i)
            ordinal = seen[key]
            seen[key] += 1
            if ordinal < len(groups[key]):
                j = groups[key][ordinal]
                if comparable_geometry(i["geometry"]) != comparable_geometry(j["geometry"]):
                    geometry_changes.append({"identity": key, "occurrence_ordinal": ordinal + 1,
                        "source_id": i["id"], "result_id": j["id"], "source_geometry": i["geometry"], "result_geometry": j["geometry"],
                        "duplicate_pairing_ambiguous": ca[key] > 1 or cb[key] > 1})
        source_subsequence = iter(b)
        subsequence = all(any(candidate == key for candidate in source_subsequence) for key in a)
        per_part.append({"part": part, "source_count": len(before), "result_count": len(after),
            "missing_occurrences": missing, "added_occurrences": added,
            "exact_reference_identity_sequence_equal": a == b,
            "source_sequence_is_subsequence_of_result": subsequence,
            "relative_order_of_common_occurrences_changed": projection(a) != projection(b),
            "geometry_changes": geometry_changes,
            "source_identity_sequence": a, "result_identity_sequence": b})
    toc = {k: {"source": source["toc"][k], "result": target["toc"][k],
               "decreased": target["toc"][k] < source["toc"][k]}
           for k in ("field_count", "fields_with_cached_text", "content_control_count", "content_controls_with_cached_or_static_text")}
    return {"zip_parts": {"added": sorted(new.keys() - old.keys()), "removed": sorted(old.keys() - new.keys()),
                          "modified": sorted(modified), "unchanged_count": len(old.keys() & new.keys()) - len(modified)},
            "original_media_hashes_missing_from_result_package": [{"sha256": h, "source_parts": parts} for h, parts in old_hashes.items() if h not in new_hashes],
            "original_bookmark_names_missing": sorted(old_names.keys() - new_names.keys()),
            "bookmark_occurrence_count_decreases": [{"name": n, "source_count": old_names[n], "result_count": new_names[n]} for n in old_names if new_names[n] < old_names[n]],
            "images_by_part": per_part, "toc": toc,
            "body_block_counts": {"source": len(source["body_blocks"]), "result": len(target["body_blocks"]), "content_completeness_proven": False},
            "result_integrity": target["integrity"]}


def audit(source_path, result_path=None):
    source = inventory(source_path)
    report = {"schema_version": 1, "audit_scope": "basic_read_only_docx_package_and_xml_audit",
              "document_completeness_proven": False, "limitations": LIMITATIONS, "source": source}
    if result_path is not None:
        result = inventory(result_path)
        report["result"] = result
        report["comparison"] = compare(source, result)
    return report


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("result", type=Path, nargs="?")
    parser.add_argument("--out", type=Path, required=True, help="JSON report; must not overwrite either input")
    args = parser.parse_args(argv)
    inputs = [p for p in (args.source, args.result) if p is not None]
    for p in inputs:
        if args.out.resolve() == p.resolve() or (args.out.exists() and p.exists() and os.path.samefile(args.out, p)):
            parser.error("--out must be different from all input files (including links)")
    temp_name = None
    try:
        report = audit(args.source, args.result)
        args.out.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile("w", encoding="utf-8", dir=args.out.parent, prefix=".docx-audit-", suffix=".json", delete=False) as f:
            temp_name = f.name
            json.dump(report, f, ensure_ascii=False, indent=2)
            f.write("\n")
        os.replace(temp_name, args.out)
        print("Wrote basic audit report: " + str(args.out))
        return 0
    except (OSError, ValueError, zipfile.BadZipFile, RuntimeError) as exc:
        print("Audit failed: " + str(exc), file=sys.stderr)
        return 1
    finally:
        if temp_name and os.path.exists(temp_name):
            os.unlink(temp_name)


if __name__ == "__main__":
    raise SystemExit(main())
