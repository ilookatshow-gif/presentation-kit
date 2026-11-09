#!/usr/bin/env python3
"""Original, offline JSON-to-editable-PPTX writer with a bounded fixed template."""
import argparse
import hashlib
import io
import json
import math
import os
from pathlib import Path, PurePosixPath
import posixpath
import re
import sys
import tempfile
import unicodedata
import xml.etree.ElementTree as ET
import zipfile

VERSION = "0.1.0"
MAX_INPUT = 512 * 1024
MAX_PPTX = 8 * 1024 * 1024
MAX_EXPANDED = 16 * 1024 * 1024
MAX_PART = 1024 * 1024
EMU = 12700
WIDTH, HEIGHT = 960, 540
LAYOUTS = ("cover", "bullets", "two_columns", "table")
NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
      "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
      "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
      "dc": "http://purl.org/dc/elements/1.1/",
      "ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"}
for _prefix in ("p", "a", "r", "cp", "dc"):
    ET.register_namespace(_prefix, NS[_prefix])
OFFICE_REL = NS["r"] + "/"
CORE_REL = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
MIME = {"presentation": "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
        "slideMaster": "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
        "slideLayout": "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
        "slide": "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
        "presProps": "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml",
        "theme": "application/vnd.openxmlformats-officedocument.theme+xml"}


class KitError(Exception):
    def __init__(self, code, message):
        self.code, self.message = code, message
        super().__init__(code)


def fail(code, message):
    raise KitError(code, message)


def digest(data):
    return hashlib.sha256(data).hexdigest()


def safe_path(value):
    path = Path(os.path.abspath(str(value)))
    if any(p.is_symlink() for p in (path,) + tuple(path.parents)):
        fail("UNSAFE_PATH", "Symlink paths are not accepted.")
    return path


def read_bytes(path, maximum):
    path = safe_path(path)
    if not path.is_file():
        fail("INPUT_INVALID", "Input must be a regular file.")
    with path.open("rb") as handle:
        data = handle.read(maximum + 1)
    if len(data) > maximum:
        fail("INPUT_TOO_LARGE", "Input exceeds the documented size limit.")
    return data


def unique_keys(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError("duplicate key")
        result[key] = value
    return result


def load_input(path):
    raw = read_bytes(path, MAX_INPUT)
    try:
        value = json.loads(raw.decode("utf-8"), object_pairs_hook=unique_keys)
    except (ValueError, UnicodeError):
        fail("INPUT_INVALID", "Input must be UTF-8 JSON with unique object keys.")
    validate_deck(value)
    return value, raw


def object_keys(value, required, optional=()):
    if (not isinstance(value, dict) or not set(required) <= set(value)
            or set(value) - set(required) - set(optional)):
        fail("INPUT_INVALID", "An object has missing or unsupported fields.")


def text(value, maximum):
    if (not isinstance(value, str) or not value.strip() or len(value) > maximum
            or any(unicodedata.category(c).startswith("C") for c in value)):
        fail("INPUT_INVALID", "Text must be nonempty, bounded and contain no control characters.")
    return value


def weight(character):
    if character.isspace():
        return 0.36
    if unicodedata.east_asian_width(character) in ("W", "F"):
        return 1.10
    if character in "MW@%ШЩЖЮФЫ":
        return 1.05
    if character in "ilI.,:;'!|":
        return 0.34
    return 0.72


def line_count(value, size, width):
    """Conservative heuristic, not a font renderer or an Office fit guarantee."""
    available = width - 12
    lines, used = 1, 0.0
    for token in re.findall(r"\S+|\s+", value):
        length = sum(weight(c) for c in token) * size
        if length > available and not token.isspace():
            fail("TEXT_OVERFLOW", "A word is too wide for the selected fixed layout.")
        if used and used + length > available:
            lines += 1
            used = 0.0
        used += length
    return lines


def fit(value, size, width, height):
    if line_count(value, size, width) * size * 1.25 > height - 8:
        fail("TEXT_OVERFLOW", "Text exceeds the conservative fixed-layout budget; shorten it.")


def bullet_values(values, maximum_count, maximum_length, size, width, height):
    if not isinstance(values, list) or not 1 <= len(values) <= maximum_count:
        fail("INPUT_INVALID", "A bullet list has an invalid length.")
    for value in values:
        text(value, maximum_length)
    if sum(line_count(v, size, width - 20) * size * 1.25 + 12 for v in values) > height - 8:
        fail("TEXT_OVERFLOW", "Bullets exceed the conservative fixed-layout budget; shorten them.")


def table_heights(slide):
    width = 864 / len(slide["columns"])
    heights = []
    for row in [slide["columns"]] + slide["rows"]:
        lines = max(line_count(cell, 18, width - 16) for cell in row)
        if lines > 3:
            fail("TEXT_OVERFLOW", "A table cell exceeds three estimated lines.")
        heights.append(lines * 23 + 18)
    if sum(heights) > 330:
        fail("TEXT_OVERFLOW", "The table exceeds the fixed-slide height; shorten or split it.")
    return heights


def validate_deck(value):
    object_keys(value, ("title", "slides"), ("author",))
    text(value["title"], 100)
    if "author" in value:
        text(value["author"], 100)
        fit(value["author"], 12, 740, 24)
    slides = value["slides"]
    if not isinstance(slides, list) or not 1 <= len(slides) <= 30:
        fail("INPUT_INVALID", "Provide between 1 and 30 slides.")
    for slide in slides:
        if not isinstance(slide, dict) or slide.get("layout") not in LAYOUTS:
            fail("INPUT_INVALID", "Choose a supported slide layout.")
        layout = slide["layout"]
        fields = {"cover": ("layout", "title"), "bullets": ("layout", "title", "bullets"),
                  "two_columns": ("layout", "title", "left", "right"),
                  "table": ("layout", "title", "columns", "rows")}[layout]
        object_keys(slide, fields, ("subtitle",) if layout == "cover" else ())
        text(slide["title"], 70)
        fit(slide["title"], 44 if layout == "cover" else 34, 864, 200 if layout == "cover" else 100)
        if layout == "cover":
            if "subtitle" in slide:
                text(slide["subtitle"], 180)
                fit(slide["subtitle"], 22, 864, 100)
        elif layout == "bullets":
            bullet_values(slide["bullets"], 6, 180, 24, 844, 320)
        elif layout == "two_columns":
            for side in ("left", "right"):
                column = slide[side]
                object_keys(column, ("heading", "bullets"))
                text(column["heading"], 45)
                fit(column["heading"], 24, 414, 70)
                bullet_values(column["bullets"], 5, 140, 22, 394, 245)
        else:
            columns, rows = slide["columns"], slide["rows"]
            if not isinstance(columns, list) or not 2 <= len(columns) <= 4:
                fail("INPUT_INVALID", "A table needs between 2 and 4 columns.")
            if not isinstance(rows, list) or not 1 <= len(rows) <= 6:
                fail("INPUT_INVALID", "A table needs between 1 and 6 body rows.")
            for cell in columns:
                text(cell, 35)
            for row in rows:
                if not isinstance(row, list) or len(row) != len(columns):
                    fail("INPUT_INVALID", "Every row must have the same number of cells as columns.")
                for cell in row:
                    text(cell, 100)
            table_heights(slide)
    return {"ok": True, "command": "validate", "slides": len(slides),
            "layouts": [s["layout"] for s in slides], "estimated_text_fit": "pass", "network_used": False}


def q(name):
    prefix, local = name.split(":")
    return "{" + NS[prefix] + "}" + local


def element(name, attributes=None, parent=None, content=None):
    node = ET.Element(q(name), attributes or {}) if parent is None else ET.SubElement(parent, q(name), attributes or {})
    if content is not None:
        node.text = content
    return node


def xml(node):
    namespace = node.tag.split("}", 1)[0].lstrip("{")
    if namespace in (NS["ct"], NS["pr"], NS["ep"]):
        # Conventional OPC serialization avoids importer incompatibility with
        # otherwise namespace-equivalent ns0-prefixed package metadata.
        ET.register_namespace("", namespace)
    return ET.tostring(node, encoding="utf-8", xml_declaration=True)


def relations(items):
    node = element("pr:Relationships")
    for ident, kind, target in items:
        element("pr:Relationship", {"Id": ident, "Type": kind, "Target": target}, node)
    return xml(node)


def transform(parent, x, y, width, height, prefix="a"):
    node = element(prefix + ":xfrm", parent=parent)
    element("a:off", {"x": str(round(x * EMU)), "y": str(round(y * EMU))}, node)
    element("a:ext", {"cx": str(round(width * EMU)), "cy": str(round(height * EMU))}, node)
    return node


def shape_tree(parent):
    tree = element("p:spTree", parent=parent)
    nv = element("p:nvGrpSpPr", parent=tree)
    element("p:cNvPr", {"id": "1", "name": ""}, nv)
    element("p:cNvGrpSpPr", parent=nv)
    element("p:nvPr", parent=nv)
    group = element("p:grpSpPr", parent=tree)
    trans = transform(group, 0, 0, 0, 0)
    element("a:chOff", {"x": "0", "y": "0"}, trans)
    element("a:chExt", {"cx": "0", "cy": "0"}, trans)
    return tree


def text_body(parent, values, size, color="142B3A", bold=False, bullet=False, cell=False):
    body = element("a:txBody" if cell else "p:txBody", parent=parent)
    body_pr = element("a:bodyPr", {"wrap": "square", "lIns": "0", "rIns": "0", "tIns": "0", "bIns": "0", "anchor": "t"}, body)
    element("a:noAutofit", parent=body_pr)
    element("a:lstStyle", parent=body)
    for value in values:
        para = element("a:p", parent=body)
        attributes = {"marL": str(20 * EMU), "indent": str(-16 * EMU)} if bullet else {}
        prop = element("a:pPr", attributes, para)
        spacing = element("a:lnSpc", parent=prop)
        element("a:spcPct", {"val": "120000"}, spacing)
        after = element("a:spcAft", parent=prop)
        element("a:spcPts", {"val": "900" if bullet else "0"}, after)
        element("a:buChar", {"char": "•"}, prop) if bullet else element("a:buNone", parent=prop)
        run = element("a:r", parent=para)
        rp = element("a:rPr", {"lang": "en-US", "sz": str(size * 100), "b": "1" if bold else "0"}, run)
        fill = element("a:solidFill", parent=rp)
        element("a:srgbClr", {"val": color}, fill)
        for kind in ("latin", "ea", "cs"):
            element("a:" + kind, {"typeface": "Arial"}, rp)
        element("a:t", {"{http://www.w3.org/XML/1998/namespace}space": "preserve"}, run, value)
        element("a:endParaRPr", {"lang": "en-US", "sz": str(size * 100)}, para)
    return body


def textbox(tree, ident, name, values, x, y, width, height, size, color="142B3A", bold=False, bullet=False):
    shape = element("p:sp", parent=tree)
    nv = element("p:nvSpPr", parent=shape)
    element("p:cNvPr", {"id": str(ident), "name": name}, nv)
    element("p:cNvSpPr", {"txBox": "1"}, nv)
    element("p:nvPr", parent=nv)
    props = element("p:spPr", parent=shape)
    transform(props, x, y, width, height)
    geom = element("a:prstGeom", {"prst": "rect"}, props)
    element("a:avLst", parent=geom)
    element("a:noFill", parent=props)
    line = element("a:ln", parent=props)
    element("a:noFill", parent=line)
    text_body(shape, values, size, color, bold, bullet)


def native_table(tree, slide):
    frame = element("p:graphicFrame", parent=tree)
    nv = element("p:nvGraphicFramePr", parent=frame)
    element("p:cNvPr", {"id": "3", "name": "Editable table"}, nv)
    element("p:cNvGraphicFramePr", parent=nv)
    element("p:nvPr", parent=nv)
    heights = table_heights(slide)
    transform(frame, 48, 150, 864, sum(heights), "p")
    graphic = element("a:graphic", parent=frame)
    data = element("a:graphicData", {"uri": "http://schemas.openxmlformats.org/drawingml/2006/table"}, graphic)
    table = element("a:tbl", parent=data)
    element("a:tblPr", {"firstRow": "1", "bandRow": "0"}, table)
    grid = element("a:tblGrid", parent=table)
    columns = len(slide["columns"])
    col_widths = [round(864 * EMU / columns)] * columns
    col_widths[-1] = 864 * EMU - sum(col_widths[:-1])
    for width in col_widths:
        element("a:gridCol", {"w": str(width)}, grid)
    for index, values in enumerate([slide["columns"]] + slide["rows"]):
        row = element("a:tr", {"h": str(heights[index] * EMU)}, table)
        for value in values:
            cell = element("a:tc", parent=row)
            text_body(cell, [value], 18, "FFFFFF" if index == 0 else "142B3A", index == 0, cell=True)
            prop = element("a:tcPr", {"marL": str(8 * EMU), "marR": str(8 * EMU), "marT": str(8 * EMU), "marB": str(8 * EMU)}, cell)
            for side in ("L", "R", "T", "B"):
                line = element("a:ln" + side, {"w": "6350"}, prop)
                solid = element("a:solidFill", parent=line)
                element("a:srgbClr", {"val": "C7D5DA"}, solid)
            fill = element("a:solidFill", parent=prop)
            element("a:srgbClr", {"val": "142B3A" if index == 0 else ("EDF4F3" if index % 2 else "FFFFFF")}, fill)


def slide_xml(slide, index, count, author):
    root = element("p:sld")
    common = element("p:cSld", {"name": "presentation-kit:" + slide["layout"]}, root)
    background = element("p:bg", parent=common)
    prop = element("p:bgPr", parent=background)
    fill = element("a:solidFill", parent=prop)
    element("a:srgbClr", {"val": "FFFFFF"}, fill)
    element("a:effectLst", parent=prop)
    tree = shape_tree(common)
    if slide["layout"] == "cover":
        textbox(tree, 2, "Title", [slide["title"]], 48, 112, 864, 200, 44, bold=True)
        if "subtitle" in slide:
            textbox(tree, 3, "Subtitle", [slide["subtitle"]], 48, 338, 864, 100, 22, "375C68")
    else:
        textbox(tree, 2, "Title", [slide["title"]], 48, 32, 864, 100, 34, bold=True)
        if slide["layout"] == "bullets":
            textbox(tree, 3, "Bullets", slide["bullets"], 58, 150, 844, 320, 24, bullet=True)
        elif slide["layout"] == "two_columns":
            for ident, side, x in ((3, "left", 48), (5, "right", 498)):
                column = slide[side]
                textbox(tree, ident, side + " heading", [column["heading"]], x, 148, 414, 70, 24, "00796B", True)
                textbox(tree, ident + 1, side + " bullets", column["bullets"], x + 10, 225, 394, 245, 22, bullet=True)
        else:
            native_table(tree, slide)
    if author:
        textbox(tree, 20, "Author supplied by user", [author], 48, 505, 740, 24, 12, "536F7A")
    textbox(tree, 21, "Page number", [str(index) + " / " + str(count)], 840, 505, 72, 24, 12, "536F7A")
    override = element("p:clrMapOvr", parent=root)
    element("a:masterClrMapping", parent=override)
    return xml(root)


def theme_xml():
    root = element("a:theme", {"name": "Presentation Kit original text theme"})
    group = element("a:themeElements", parent=root)
    colors = element("a:clrScheme", {"name": "Original"}, group)
    values = {"dk1": "142B3A", "lt1": "FFFFFF", "dk2": "375C68", "lt2": "EDF4F3", "accent1": "00796B",
              "accent2": "385A91", "accent3": "AB6C32", "accent4": "7960A1", "accent5": "3E8279", "accent6": "94515B", "hlink": "215B9B", "folHlink": "744D89"}
    for key, value in values.items():
        node = element("a:" + key, parent=colors)
        element("a:srgbClr", {"val": value}, node)
    fonts = element("a:fontScheme", {"name": "Arial references; no fonts embedded"}, group)
    for family in ("majorFont", "minorFont"):
        node = element("a:" + family, parent=fonts)
        for kind in ("latin", "ea", "cs"):
            element("a:" + kind, {"typeface": "Arial"}, node)
    fmt = element("a:fmtScheme", {"name": "Original"}, group)
    fills = element("a:fillStyleLst", parent=fmt)
    for unused in range(3):
        fill = element("a:solidFill", parent=fills)
        element("a:schemeClr", {"val": "phClr"}, fill)
    lines = element("a:lnStyleLst", parent=fmt)
    for width in (6350, 12700, 19050):
        line = element("a:ln", {"w": str(width), "cap": "flat", "cmpd": "sng", "algn": "ctr"}, lines)
        fill = element("a:solidFill", parent=line)
        element("a:schemeClr", {"val": "phClr"}, fill)
        element("a:prstDash", {"val": "solid"}, line)
        element("a:miter", {"lim": "800000"}, line)
    effects = element("a:effectStyleLst", parent=fmt)
    for unused in range(3):
        effect = element("a:effectStyle", parent=effects)
        element("a:effectLst", parent=effect)
    backgrounds = element("a:bgFillStyleLst", parent=fmt)
    for unused in range(3):
        fill = element("a:solidFill", parent=backgrounds)
        element("a:schemeClr", {"val": "phClr"}, fill)
    return xml(root)


def package_parts(deck):
    count = len(deck["slides"])
    presentation = element("p:presentation")
    masters = element("p:sldMasterIdLst", parent=presentation)
    element("p:sldMasterId", {"id": "2147483648", q("r:id"): "rId1"}, masters)
    slides = element("p:sldIdLst", parent=presentation)
    for index in range(1, count + 1):
        element("p:sldId", {"id": str(255 + index), q("r:id"): "rId" + str(index + 2)}, slides)
    element("p:sldSz", {"cx": str(WIDTH * EMU), "cy": str(HEIGHT * EMU), "type": "screen16x9"}, presentation)
    element("p:notesSz", {"cx": "6858000", "cy": "9144000"}, presentation)
    master = element("p:sldMaster")
    shape_tree(element("p:cSld", parent=master))
    element("p:clrMap", {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2", "accent1": "accent1", "accent2": "accent2", "accent3": "accent3", "accent4": "accent4", "accent5": "accent5", "accent6": "accent6", "hlink": "hlink", "folHlink": "folHlink"}, master)
    layouts = element("p:sldLayoutIdLst", parent=master)
    element("p:sldLayoutId", {"id": "2147483649", q("r:id"): "rId1"}, layouts)
    styles = element("p:txStyles", parent=master)
    for name in ("titleStyle", "bodyStyle", "otherStyle"):
        element("p:" + name, parent=styles)
    layout = element("p:sldLayout", {"type": "blank", "preserve": "1"})
    shape_tree(element("p:cSld", {"name": "Blank"}, layout))
    element("a:masterClrMapping", parent=element("p:clrMapOvr", parent=layout))
    core = element("cp:coreProperties")
    element("dc:title", parent=core, content=deck["title"])
    element("dc:creator", parent=core, content=deck.get("author", ""))
    app = element("ep:Properties")
    for name, content in (("Application", "Lookatshow Presentation Kit"), ("AppVersion", VERSION), ("Slides", str(count)), ("PresentationFormat", "On-screen Show (16:9)")):
        element("ep:" + name, parent=app, content=content)
    parts = {"_rels/.rels": relations([("rId1", OFFICE_REL + "officeDocument", "ppt/presentation.xml"), ("rId2", CORE_REL, "docProps/core.xml"), ("rId3", OFFICE_REL + "extended-properties", "docProps/app.xml")]),
             "docProps/core.xml": xml(core), "docProps/app.xml": xml(app),
             "ppt/presentation.xml": xml(presentation), "ppt/presProps.xml": xml(element("p:presentationPr")),
             "ppt/_rels/presentation.xml.rels": relations([("rId1", OFFICE_REL + "slideMaster", "slideMasters/slideMaster1.xml"), ("rId2", OFFICE_REL + "presProps", "presProps.xml")] + [("rId" + str(i + 2), OFFICE_REL + "slide", "slides/slide" + str(i) + ".xml") for i in range(1, count + 1)]),
             "ppt/slideMasters/slideMaster1.xml": xml(master),
             "ppt/slideMasters/_rels/slideMaster1.xml.rels": relations([("rId1", OFFICE_REL + "slideLayout", "../slideLayouts/slideLayout1.xml"), ("rId2", OFFICE_REL + "theme", "../theme/theme1.xml")]),
             "ppt/slideLayouts/slideLayout1.xml": xml(layout),
             "ppt/slideLayouts/_rels/slideLayout1.xml.rels": relations([("rId1", OFFICE_REL + "slideMaster", "../slideMasters/slideMaster1.xml")]),
             "ppt/theme/theme1.xml": theme_xml()}
    content_types = element("ct:Types")
    element("ct:Default", {"Extension": "rels", "ContentType": "application/vnd.openxmlformats-package.relationships+xml"}, content_types)
    element("ct:Default", {"Extension": "xml", "ContentType": "application/xml"}, content_types)
    overrides = {"ppt/presentation.xml": MIME["presentation"], "ppt/presProps.xml": MIME["presProps"], "ppt/slideMasters/slideMaster1.xml": MIME["slideMaster"], "ppt/slideLayouts/slideLayout1.xml": MIME["slideLayout"], "ppt/theme/theme1.xml": MIME["theme"], "docProps/core.xml": "application/vnd.openxmlformats-package.core-properties+xml", "docProps/app.xml": "application/vnd.openxmlformats-officedocument.extended-properties+xml"}
    for index, slide in enumerate(deck["slides"], 1):
        path = "ppt/slides/slide" + str(index) + ".xml"
        parts[path] = slide_xml(slide, index, count, deck.get("author", ""))
        parts["ppt/slides/_rels/slide" + str(index) + ".xml.rels"] = relations([("rId1", OFFICE_REL + "slideLayout", "../slideLayouts/slideLayout1.xml")])
        overrides[path] = MIME["slide"]
    for path, kind in overrides.items():
        element("ct:Override", {"PartName": "/" + path, "ContentType": kind}, content_types)
    parts["[Content_Types].xml"] = xml(content_types)
    return parts


def make_pptx(deck):
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w", zipfile.ZIP_STORED) as archive:
        for name, data in sorted(package_parts(deck).items()):
            info = zipfile.ZipInfo(name, (2026, 1, 1, 0, 0, 0))
            info.create_system = 3
            info.compress_type = zipfile.ZIP_STORED
            info.external_attr = 0o100644 << 16
            archive.writestr(info, data)
    return stream.getvalue()


def write_output(path, data, source, replace):
    output = safe_path(path)
    source = safe_path(source)
    if output.suffix.lower() != ".pptx":
        fail("INPUT_INVALID", "Output must have the .pptx extension.")
    if output == source or (output.exists() and os.path.samefile(output, source)):
        fail("OUTPUT_COLLISION", "Output must not replace the input document or its hardlink.")
    if not output.parent.is_dir():
        fail("UNSAFE_PATH", "Create the output parent directory explicitly first.")
    if output.exists() and (not replace or not output.is_file()):
        fail("OUTPUT_EXISTS", "Output exists; use --replace only for an intentional rebuild.")
    fd, temporary = tempfile.mkstemp(prefix=".presentation-kit-", dir=str(output.parent))
    try:
        with os.fdopen(fd, "wb") as handle:
            handle.write(data)
            handle.flush()
            os.fsync(handle.fileno())
        safe_path(output)
        if replace:
            os.replace(temporary, str(output))
        else:
            os.link(temporary, str(output))
    except FileExistsError:
        fail("OUTPUT_EXISTS", "Output exists; no overwrite was performed.")
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)
    return output


def build(input_path, output, replace=False):
    deck, source = load_input(input_path)
    data = make_pptx(deck)
    inspect_bytes(data)
    written = write_output(output, data, input_path, replace)
    return {"ok": True, "command": "build", "slides": len(deck["slides"]), "output": str(written),
            "bytes": len(data), "pptx_sha256": digest(data), "source_sha256": digest(source),
            "editable_text": True, "editable_tables": True, "text_fit": "conservative_estimate", "network_used": False}


def inspect_bytes(raw):
    if len(raw) > MAX_PPTX:
        fail("INVALID_PPTX", "PPTX exceeds the inspection size limit.")
    try:
        with zipfile.ZipFile(io.BytesIO(raw)) as archive:
            infos = archive.infolist()
            names = [item.filename for item in infos]
            if (len(infos) > 72 or len(set(names)) != len(names)
                    or sum(item.file_size for item in infos) > MAX_EXPANDED):
                fail("INVALID_PPTX", "PPTX has duplicate, excessive or oversized parts.")
            parts = {}
            for item in infos:
                if (item.file_size > MAX_PART or item.flag_bits & 1 or item.is_dir()
                        or item.external_attr >> 16 & 0o170000 == 0o120000
                        or "\\" in item.filename or item.filename.startswith("/")
                        or any(p in ("", ".", "..") for p in item.filename.split("/"))):
                    fail("INVALID_PPTX", "PPTX contains an unsafe or unsupported part.")
                data = archive.read(item)
                decoded = data.decode("utf-8")
                declared = re.search(r"<\?xml[^>]*encoding\s*=\s*['\"]([^'\"]+)", decoded, re.I)
                if "\x00" in decoded or (declared and declared.group(1).lower() not in ("utf-8", "utf8")):
                    fail("INVALID_PPTX", "Only UTF-8 XML parts without NUL bytes are supported.")
                if "<!DOCTYPE" in decoded.upper() or "<!ENTITY" in decoded.upper():
                    fail("INVALID_PPTX", "DTD and entity declarations are not supported.")
                parts[item.filename] = ET.fromstring(decoded)
        presentation = parts["ppt/presentation.xml"]
        ids = presentation.findall("p:sldIdLst/p:sldId", NS)
        count = len(ids)
        if (not 1 <= count <= 30 or len({x.get(q("r:id")) for x in ids}) != count
                or {x.get("id") for x in ids} != {str(i) for i in range(256, 256 + count)}):
            fail("INVALID_PPTX", "Slide references are missing or duplicated.")
        expected_parts = package_parts({"title": "check", "slides": [{"layout": "cover", "title": "check"}] * count})
        if set(parts) != set(expected_parts):
            fail("INVALID_PPTX", "Inspection supports only the presentation-kit fixed package format.")
        for name, root in parts.items():
            if root.tag != ET.fromstring(expected_parts[name]).tag:
                fail("INVALID_PPTX", "A package part has an unsupported root element.")
        def content_types(root):
            records = {}
            for child in root:
                if child.tag == q("ct:Default") and set(child.attrib) == {"Extension", "ContentType"}:
                    key = ("default", child.get("Extension"))
                elif child.tag == q("ct:Override") and set(child.attrib) == {"PartName", "ContentType"}:
                    key = ("override", child.get("PartName"))
                else:
                    fail("INVALID_PPTX", "Unsupported content-type declaration.")
                if key in records or len(child):
                    fail("INVALID_PPTX", "Duplicate or nested content-type declaration.")
                records[key] = child.attrib
            return records
        if content_types(parts["[Content_Types].xml"]) != content_types(ET.fromstring(expected_parts["[Content_Types].xml"])):
            fail("INVALID_PPTX", "Content types do not match the fixed non-macro package format.")
        app = parts["docProps/app.xml"]
        if app.findtext("ep:Application", namespaces=NS) != "Lookatshow Presentation Kit" or app.findtext("ep:AppVersion", namespaces=NS) != VERSION:
            fail("INVALID_PPTX", "The package is not the supported presentation-kit version.")
        size = presentation.find("p:sldSz", NS)
        if size is None or size.get("cx") != str(WIDTH * EMU) or size.get("cy") != str(HEIGHT * EMU):
            fail("INVALID_PPTX", "The fixed package requires 16:9 dimensions.")
        for name, root in parts.items():
            if not name.endswith(".rels"):
                continue
            if root.tag != q("pr:Relationships"):
                fail("INVALID_PPTX", "Invalid relationships part.")
            seen = set()
            folder = "" if name == "_rels/.rels" else posixpath.dirname(posixpath.dirname(name))
            for rel in root:
                if (rel.tag != q("pr:Relationship") or not rel.get("Id") or rel.get("Id") in seen
                        or rel.get("TargetMode", "Internal") != "Internal"
                        or rel.get("Type") not in {CORE_REL} | {OFFICE_REL + s for s in ("officeDocument", "extended-properties", "slide", "slideMaster", "slideLayout", "theme", "presProps")}):
                    fail("INVALID_PPTX", "Unsupported, duplicate or external relationship.")
                target = rel.get("Target", "")
                if not target or ":" in target or "\\" in target or target.startswith("/"):
                    fail("INVALID_PPTX", "Unsafe relationship target.")
                resolved = posixpath.normpath(posixpath.join(folder, target))
                if resolved not in parts:
                    fail("INVALID_PPTX", "Relationship target is missing.")
                seen.add(rel.get("Id"))
            expected_relationships = {child.get("Id"): child.attrib for child in ET.fromstring(expected_parts[name])}
            if {child.get("Id"): child.attrib for child in root} != expected_relationships:
                fail("INVALID_PPTX", "The fixed package requires its complete internal relationship graph.")
        rels = {x.get("Id"): x for x in parts["ppt/_rels/presentation.xml.rels"]}
        outline = []
        for number, ident in enumerate(ids, 1):
            rel = rels.get(ident.get(q("r:id")))
            if rel is None or rel.get("Type") != OFFICE_REL + "slide":
                fail("INVALID_PPTX", "A slide references a non-slide relationship.")
            path = posixpath.normpath(posixpath.join("ppt", rel.get("Target")))
            slide = parts[path]
            common = slide.find("p:cSld", NS)
            layout = common.get("name", "").removeprefix("presentation-kit:") if common is not None else ""
            if layout not in LAYOUTS or slide.tag != q("p:sld"):
                fail("INVALID_PPTX", "Unsupported slide layout marker.")
            shapes = slide.findall(".//p:sp", NS)
            cells = slide.findall(".//a:tc", NS)
            for graphic in slide.findall(".//a:graphicData", NS):
                if graphic.get("uri") != "http://schemas.openxmlformats.org/drawingml/2006/table":
                    fail("INVALID_PPTX", "Unsupported native graphic type.")
            for trans in slide.findall(".//a:xfrm", NS) + slide.findall(".//p:graphicFrame/p:xfrm", NS):
                off, ext = trans.find("a:off", NS), trans.find("a:ext", NS)
                if off is None or ext is None:
                    fail("INVALID_PPTX", "Missing geometry.")
                values = (off.get("x"), off.get("y"), ext.get("cx"), ext.get("cy"))
                if any(not isinstance(v, str) or re.fullmatch(r"[0-9]{1,10}", v) is None for v in values):
                    fail("INVALID_PPTX", "Invalid or excessive geometry value.")
                x, y, width, height = [int(v) for v in values]
                if min(x, y, width, height) < 0 or x + width > WIDTH * EMU or y + height > HEIGHT * EMU:
                    fail("INVALID_PPTX", "A native object exceeds the slide canvas.")
            outline.append({"slide": number, "layout": layout, "text_boxes": len(shapes), "table_cells": len(cells),
                            "text": [x.text or "" for x in slide.findall(".//a:t", NS)]})
        return {"ok": True, "command": "inspect", "format": "presentation-kit/1", "slides": count,
                "outline": outline, "external_relationships": 0, "network_used": False}
    except (zipfile.BadZipFile, ET.ParseError, KeyError, ValueError, TypeError, OSError, RuntimeError):
        fail("INVALID_PPTX", "PPTX is malformed or outside the supported fixed package format.")


def inspect(path):
    return inspect_bytes(read_bytes(path, MAX_PPTX))


def demo(output, replace=False):
    result = build(Path(__file__).resolve().parent / "examples/synthetic-deck.json", output, replace)
    result.update(command="demo", synthetic=True)
    return result


def onboard():
    marker = Path(__file__).resolve().parent / ".onboarded"
    try:
        with marker.open("x", encoding="ascii") as handle:
            handle.write(VERSION + "\n")
    except FileExistsError:
        return {"ok": True, "command": "onboard", "already_shown": True}
    return {"ok": True, "command": "onboard", "already_shown": False,
            "author": "Дмитрий Лукашов", "channel": "https://t.me/lookatsoul",
            "next_step": "https://lookatshow.ru/executive?utm_source=github&utm_medium=open_module&utm_campaign=presentation-kit&utm_content=onboard"}


def schema():
    return {"ok": True, "command": "schema", "kind": "descriptive_contract_not_json_schema",
            "top": {"required": ["title", "slides"], "optional": ["author"]}, "slides": [1, 30],
            "layouts": {"cover": ["layout", "title", "subtitle?"], "bullets": ["layout", "title", "bullets"],
                        "two_columns": ["layout", "title", "left:heading+bullets", "right:heading+bullets"],
                        "table": ["layout", "title", "columns", "rows"]},
            "limits": {"input_bytes": MAX_INPUT, "deck_title_characters": 100, "author_characters": 100,
                       "slide_title_characters": 70, "subtitle_characters": 180, "bullets_count": 6,
                       "bullet_characters": 180, "column_heading_characters": 45, "column_bullets_count": 5,
                       "column_bullet_characters": 140, "table_columns": [2, 4], "table_body_rows": [1, 6],
                       "table_header_characters": 35, "table_cell_characters": 100},
            "additional_keys": False, "text_fit": "conservative_estimate_can_refuse_below_character_caps",
            "network_used": False}


class Parser(argparse.ArgumentParser):
    def error(self, message):
        fail("CLI_USAGE", "Invalid arguments; use --help for the command contract.")


def main(argv=None):
    try:
        parser = Parser(description="Build an editable fixed-layout PPTX from explicit local JSON; no model or network.")
        parser.add_argument("--json", action="store_true", help="All command results are already stable JSON.")
        parser.add_argument("--version", action="version", version=VERSION)
        sub = parser.add_subparsers(dest="command", required=True)
        for name in ("doctor", "schema", "onboard"):
            sub.add_parser(name)
        validate = sub.add_parser("validate", help="Validate JSON and conservative text-fit limits without writing.")
        validate.add_argument("input")
        build_parser = sub.add_parser("build", help="Create an editable PPTX from local JSON.")
        build_parser.add_argument("input")
        for item in (build_parser, sub.add_parser("demo", help="Write the invented five-slide example.")):
            item.add_argument("--output", required=True)
            item.add_argument("--replace", action="store_true")
        read = sub.add_parser("inspect", help="Read bounded native text/table structure from this tool's own format.")
        read.add_argument("input")
        args = parser.parse_args(argv)
        if args.command == "doctor":
            fixtures = (Path(__file__).resolve().parent / "examples/synthetic-deck.json").is_file()
            result = {"ok": sys.version_info >= (3, 9) and fixtures, "command": "doctor", "version": VERSION,
                      "python_supported": sys.version_info >= (3, 9), "synthetic_fixtures_found": fixtures,
                      "auth": "not_required", "mode": "offline", "network_used": False}
        elif args.command == "schema":
            result = schema()
        elif args.command == "validate":
            result = validate_deck(load_input(args.input)[0])
        elif args.command == "build":
            result = build(args.input, args.output, args.replace)
        elif args.command == "demo":
            result = demo(args.output, args.replace)
        elif args.command == "inspect":
            result = inspect(args.input)
        else:
            result = onboard()
        print(json.dumps(result, ensure_ascii=True, sort_keys=True))
        return 0 if result["ok"] else 2
    except KitError as error:
        print(json.dumps({"ok": False, "error": {"code": error.code, "message": error.message}}, sort_keys=True))
        return 2
    except (OSError, ValueError, TypeError, UnicodeError, RuntimeError):
        print(json.dumps({"ok": False, "error": {"code": "LOCAL_IO_ERROR", "message": "Local input or output could not be processed; no private path or content is exposed."}}, sort_keys=True))
        return 2


if __name__ == "__main__":
    sys.exit(main())
