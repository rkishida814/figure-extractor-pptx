"""pptx の解析と、描画用 pptx の組み立て。

- list_shapes(): 各スライドの最上位オブジェクト（図形・画像・グループ・表・グラフなど）の位置を返す。
- build_preview_pptx(): 非表示スライドも含めて全スライドを PDF にできるようにした複製を作る。
- build_export_pptx(): 選んだオブジェクト1つだけを載せたスライドを、選んだ数だけ先頭に並べた複製を作る。
  背景は白または黒の単色にし、マスターの図形は消すので、PDF にしたときそのオブジェクトだけが描かれる。
"""
from __future__ import annotations

import copy
import posixpath
import zipfile
from dataclasses import dataclass

from lxml import etree

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}
P = "{%s}" % NS["p"]
A = "{%s}" % NS["a"]
R = "{%s}" % NS["r"]
MC = "{%s}" % NS["mc"]

SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
SLIDE_CT = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
SHAPE_TAGS = {"sp", "grpSp", "graphicFrame", "cxnSp", "pic", "contentPart"}
SECTION_EXT_URI = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"

KIND_LABEL = {
    "sp": "図形",
    "grpSp": "グループ",
    "graphicFrame": "表・グラフ",
    "cxnSp": "線",
    "pic": "画像",
    "contentPart": "インク",
}


def _local(tag) -> str:
    return etree.QName(tag).localname if isinstance(tag, str) else ""


def _read_zip(path: str) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}


def _write_zip(parts: dict[str, bytes], path: str) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        # [Content_Types].xml を先頭に置く
        order = sorted(parts, key=lambda n: (n != "[Content_Types].xml", n))
        for n in order:
            media = n.startswith("ppt/media/") or n.startswith("ppt/embeddings/")
            z.writestr(n, parts[n], compress_type=zipfile.ZIP_STORED if media else zipfile.ZIP_DEFLATED)


def _xml(b: bytes):
    return etree.fromstring(b, parser=etree.XMLParser(remove_blank_text=False, huge_tree=True))


def _dump(el) -> bytes:
    return etree.tostring(el, xml_declaration=True, encoding="UTF-8", standalone=True)


def _rels_path(part: str) -> str:
    d, f = posixpath.split(part)
    return posixpath.join(d, "_rels", f + ".rels")


@dataclass
class Package:
    parts: dict[str, bytes]
    pres: etree._Element
    pres_rels: etree._Element
    slide_parts: list[str]  # 表示順

    @classmethod
    def load(cls, path: str) -> "Package":
        parts = _read_zip(path)
        pres = _xml(parts["ppt/presentation.xml"])
        pres_rels = _xml(parts["ppt/_rels/presentation.xml.rels"])
        targets = {
            r.get("Id"): posixpath.normpath(posixpath.join("ppt", r.get("Target")))
            for r in pres_rels.findall("rel:Relationship", NS)
        }
        slide_parts = []
        lst = pres.find("p:sldIdLst", NS)
        if lst is not None:
            for s in lst.findall("p:sldId", NS):
                slide_parts.append(targets[s.get(R + "id")])
        return cls(parts, pres, pres_rels, slide_parts)

    def slide_size(self) -> tuple[int, int]:
        sz = self.pres.find("p:sldSz", NS)
        return int(sz.get("cx")), int(sz.get("cy"))


# ---------------------------------------------------------------- 図形の一覧

def _inner(el):
    """mc:AlternateContent なら中身の図形要素を返す。"""
    if _local(el.tag) != "AlternateContent":
        return el
    for branch in (el.find("mc:Choice", NS), el.find("mc:Fallback", NS)):
        if branch is not None:
            for c in branch:
                if _local(c.tag) in SHAPE_TAGS:
                    return c
    return None


def _top_shapes(sptree):
    for el in sptree:
        inner = _inner(el)
        if inner is not None and _local(inner.tag) in SHAPE_TAGS:
            yield el, inner


def _xfrm_of(inner):
    for e in inner.iter():
        if _local(e.tag) == "xfrm":
            off = e.find("a:off", NS)
            ext = e.find("a:ext", NS)
            if off is None or ext is None:
                return None
            return (int(off.get("x")), int(off.get("y")), int(ext.get("cx")), int(ext.get("cy")),
                    int(e.get("rot", "0")) / 60000.0)
    return None


def list_shapes(path: str) -> dict:
    """スライドごとの最上位オブジェクトの位置（EMU）を返す。

    オブジェクトは spTree の中での並び順（"id"）で見分ける。cNvPr の id は、別のファイルから
    貼り付けたオブジェクトで重複していることがあり、見分けに使えない。"""
    from pptx import Presentation
    from pptx.shapes.shapetree import SlideShapeFactory  # プレースホルダーの位置の継承に使う

    pkg = Package.load(path)
    prs = Presentation(path)
    sw, sh = pkg.slide_size()
    slides = []
    for idx, slide in enumerate(prs.slides):
        hidden_slide = slide._element.get("show") == "0"
        shapes = []
        sptree = slide.shapes._spTree
        for order, (el, inner) in enumerate(_top_shapes(sptree)):
            cnv = inner.find(".//p:cNvPr", NS)
            if cnv is None or cnv.get("hidden") in ("1", "true"):
                continue
            geom = None
            try:
                shp = SlideShapeFactory(inner, slide.shapes)
                if shp.width is not None and shp.left is not None:
                    rot = float(getattr(shp, "rotation", 0.0) or 0.0)
                    geom = (int(shp.left), int(shp.top), int(shp.width), int(shp.height), rot)
            except Exception:
                geom = None
            if geom is None:
                geom = _xfrm_of(inner)
            if geom is None:
                continue
            x, y, w, h, rot = geom
            shapes.append({
                "id": order,
                "name": cnv.get("name", ""),
                "kind": KIND_LABEL.get(_local(inner.tag), "図形"),
                "x": x, "y": y, "w": w, "h": h, "rot": rot,
            })
        slides.append({"index": idx, "hidden": hidden_slide, "shapes": shapes})
    return {"slideW": sw, "slideH": sh, "slides": slides}


# ---------------------------------------------------------------- 描画用 pptx

def build_preview_pptx(src: str, dst: str) -> int:
    """非表示スライドも PDF に出るようにした複製を作る。スライド数を返す。"""
    pkg = Package.load(src)
    for part in pkg.slide_parts:
        root = _xml(pkg.parts[part])
        if root.get("show") is not None:
            del root.attrib["show"]
            pkg.parts[part] = _dump(root)
    _write_zip(pkg.parts, dst)
    return len(pkg.slide_parts)


def _isolate_slide(slide_xml: bytes, keep: set[int], bg_rgb: str) -> bytes:
    """keep に入った並び順のオブジェクトだけを残す（重なり順はそのまま）。"""
    root = _xml(slide_xml)
    if root.get("show") is not None:
        del root.attrib["show"]
    root.set("showMasterSp", "0")
    # 画面切り替え・アニメーション（mc:AlternateContent で包まれていることもある）を消す
    for child in list(root):
        if _local(child.tag) in ("timing", "transition", "AlternateContent"):
            root.remove(child)
    csld = root.find("p:cSld", NS)
    old_bg = csld.find("p:bg", NS)
    if old_bg is not None:
        csld.remove(old_bg)
    bg = etree.Element(P + "bg")
    bgpr = etree.SubElement(bg, P + "bgPr")
    solid = etree.SubElement(bgpr, A + "solidFill")
    etree.SubElement(solid, A + "srgbClr").set("val", bg_rgb)
    etree.SubElement(bgpr, A + "effectLst")
    csld.insert(0, bg)

    sptree = csld.find("p:spTree", NS)
    found = set()
    for order, (el, _inner_el) in enumerate(list(_top_shapes(sptree))):
        if order in keep:
            found.add(order)
            continue
        sptree.remove(el)
    missing = set(keep) - found
    if missing:
        raise ValueError(f"{sorted(m + 1 for m in missing)} 番目のオブジェクトが見つかりません")
    return _dump(root)


def _strip_slide_rels(rels_xml: bytes) -> bytes:
    root = _xml(rels_xml)
    for r in list(root):
        t = r.get("Type", "")
        if t.endswith("/notesSlide") or "comments" in t.lower():
            root.remove(r)
    return _dump(root)


def build_export_pptx(src: str, dst: str, jobs: list[tuple[int, list[int]]]) -> None:
    """jobs = [(スライド番号0始まり, [スライド内の並び順0始まり, ...]), ...]。1つの job が1枚の図になる。
    出力の 1..n 枚目が各オブジェクトだけを白背景に、n+1..2n 枚目が黒背景に載せたスライドになる
    （2枚の差から透明度と範囲を求める）。元のスライドは非表示にして後ろに残す
    （元のスライドを消すと、ハイパーリンクなどの参照が切れて PowerPoint が修復を求めることがあるため）。"""
    pkg = Package.load(src)
    parts = pkg.parts
    ct = _xml(parts["[Content_Types].xml"])
    pres, pres_rels = pkg.pres, pkg.pres_rels

    sld_lst = pres.find("p:sldIdLst", NS)
    old_ids = list(sld_lst)
    for s in old_ids:
        sld_lst.remove(s)
    used_ids = {int(s.get("id")) for s in old_ids}
    next_id = max(used_ids | {255}) + 1

    doubled = [(j, "FFFFFF") for j in jobs] + [(j, "000000") for j in jobs]
    for i, ((slide_idx, orders), bg_rgb) in enumerate(doubled, start=1):
        src_part = pkg.slide_parts[slide_idx]
        new_part = f"ppt/slides/fxslide{i}.xml"
        parts[new_part] = _isolate_slide(parts[src_part], set(orders), bg_rgb)
        src_rels = _rels_path(src_part)
        if src_rels in parts:
            parts[_rels_path(new_part)] = _strip_slide_rels(parts[src_rels])
        ov = etree.SubElement(ct, "{%s}Override" % NS["ct"])
        ov.set("PartName", "/" + new_part)
        ov.set("ContentType", SLIDE_CT)
        rid = f"rIdFx{i}"
        rel = etree.SubElement(pres_rels, "{%s}Relationship" % NS["rel"])
        rel.set("Id", rid)
        rel.set("Type", SLIDE_REL_TYPE)
        rel.set("Target", f"slides/fxslide{i}.xml")
        sid = etree.SubElement(sld_lst, P + "sldId")
        sid.set("id", str(next_id))
        sid.set(R + "id", rid)
        next_id += 1

    # 元のスライドは非表示にして後ろへ
    for s, part in zip(old_ids, pkg.slide_parts):
        sld_lst.append(s)
        root = _xml(parts[part])
        root.set("show", "0")
        parts[part] = _dump(root)

    # セクション情報は新しいスライドを含まないので消す
    ext_lst = pres.find("p:extLst", NS)
    if ext_lst is not None:
        for ext in list(ext_lst):
            if ext.get("uri") == SECTION_EXT_URI:
                ext_lst.remove(ext)

    parts["ppt/presentation.xml"] = _dump(pres)
    parts["ppt/_rels/presentation.xml.rels"] = _dump(pres_rels)
    parts["[Content_Types].xml"] = _dump(ct)
    _write_zip(parts, dst)
