import argparse
import json
import os
import re
import shutil
import sys
import tempfile
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from copy import deepcopy
from io import BytesIO
from pathlib import Path

from docx.shared import Pt, Twips

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

PAGE_WIDTH = 11906
PAGE_HEIGHT = 16838
MARGIN_TOP = 2098
MARGIN_BOTTOM = 2041
MARGIN_LEFT = 1588
MARGIN_RIGHT = 1588
HEADER_DISTANCE = 851
FOOTER_DISTANCE = 1644

FONT_HEADER = "方正小标宋简体"
FONT_BODY = "仿宋_GB2312"
FONT_LEVEL1 = "黑体"
FONT_ASCII = "Times New Roman"
FONT_DISCLOSURE_KEY = "黑体"
FONT_DISCLOSURE_VALUE = "仿宋_GB2312"

SIZE_HEADER = Pt(46)
SIZE_TITLE = Pt(22)
SIZE_BODY = Pt(16)
SIZE_FOOTER = Pt(14)
SIZE_DISCLOSURE = Pt(14)

LINE_SPACING = Pt(27)
TWO_CHAR_INDENT = Twips(640)
RIGHT_INDENT_SIGNER = Twips(1247)
RIGHT_INDENT_DATE = Twips(10)

COLOR_RED = "FF0000"
COLOR_BLACK = "000000"

DEFAULT_DATA = {
    "title": "关于开展年度工作总结的通知",
    "recipients": ["各相关单位"],
    "body": [
        "为全面总结年度工作成果，梳理经验做法，现就有关事项通知如下。",
        "一、总体要求",
        "（一）突出重点。各单位要围绕中心任务，突出亮点工作。",
        "1. 做到数据准确、材料完整。",
        "（1）按时报送，总结材料不超过三页。",
        "请于2月15日前报送电子版材料。",
    ],
    "attachments": ["年度工作总结模板"],
}

MARKDOWN_EXTENSIONS = {".md", ".markdown"}

def _register_docx_namespaces():
    for prefix, uri in {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
        "v": "urn:schemas-microsoft-com:vml",
        "o": "urn:schemas-microsoft-com:office:office",
    }.items():
        ET.register_namespace(prefix, uri)


def _extract_header_elements(template_path):
    if not template_path or not os.path.exists(template_path):
        return None, [], {}, None

    _register_docx_namespaces()
    temp_dir = tempfile.mkdtemp(prefix="gongwen_template_")

    try:
        with zipfile.ZipFile(template_path, "r") as zf:
            zf.extractall(temp_dir)

        doc_xml_path = os.path.join(temp_dir, "word", "document.xml")
        if not os.path.exists(doc_xml_path):
            return None, [], {}, None

        tree = ET.parse(doc_xml_path)
        root = tree.getroot()
        body = root.find(f"{W_NS}body")
        if body is None:
            return None, [], {}, None

        paragraphs = body.findall(f"{W_NS}p")
        if len(paragraphs) < 2:
            return None, [], {}, None

        header_para = deepcopy(paragraphs[0])
        line_para = deepcopy(paragraphs[1])

        media_dir = os.path.join(temp_dir, "word", "media")
        media_files = {}
        if os.path.exists(media_dir):
            for f in os.listdir(media_dir):
                src_path = os.path.join(media_dir, f)
                with open(src_path, "rb") as rf:
                    media_files[f] = rf.read()

        rels_path = os.path.join(temp_dir, "word", "_rels", "document.xml.rels")
        rels_content = None
        if os.path.exists(rels_path):
            with open(rels_path, "rb") as rf:
                rels_content = rf.read()

        return (header_para, line_para), media_files, rels_content, temp_dir

    except Exception as e:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)
        print(f"提取红头失败: {e}")
        return None, [], {}, None


def _normalize_quotes(text: str) -> str:
    if not text:
        return text
    left_quote = "\u201c"
    right_quote = "\u201d"
    return re.sub(r'"([^"\n]+)"', left_quote + r"\1" + right_quote, text)


_DATE_YMD_RE = re.compile(r"^\s*(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日\s*$")


def _to_ascii_digits(text: str) -> str:
    if not text:
        return text
    # Convert full-width digits to ASCII digits (e.g., '０' -> '0').
    fullwidth_zero = ord("０")
    return text.translate({fullwidth_zero + i: ord("0") + i for i in range(10)})


def _estimate_text_width_hanzi(text: str) -> float:
    if not text:
        return 0.0
    width = 0.0
    for ch in text:
        if ch.isspace():
            continue
        # Treat East Asian wide/fullwidth characters as 1 "hanzi" unit,
        # and everything else (ASCII, halfwidth punctuation) as 0.5.
        width += 1.0 if unicodedata.east_asian_width(ch) in ("W", "F") else 0.5
    return width


def _estimate_date_width_hanzi(date_text: str):
    """
    Estimate date width using the user's empirical mapping (hanzi-equivalent units):
    - XXXX年X月X日 ≈ 7.5
    - XXXX年XX月X日 or XXXX年X月XX日 ≈ 7.8
    - XXXX年XX月XX日 ≈ 8.2
    """
    if not date_text:
        return None
    normalized = _to_ascii_digits(str(date_text).strip())
    m = _DATE_YMD_RE.match(normalized)
    if not m:
        return None
    month = m.group(2)
    day = m.group(3)
    month_2 = len(month) == 2
    day_2 = len(day) == 2
    if not month_2 and not day_2:
        return 7.5
    if month_2 and day_2:
        return 8.2
    return 7.8


def _compute_date_right_indent_twips(
    *, signer: str, date: str, signer_right_indent_twips: int, debug: bool = False
) -> int:
    """
    Compute date paragraph right indent so that the date is *center-aligned* under the signer,
    using simple width estimation.
    """
    signer = (signer or "").strip()
    date = (date or "").strip()
    if not signer or not date:
        return int(getattr(RIGHT_INDENT_DATE, "twips", RIGHT_INDENT_DATE))

    signer_width = _estimate_text_width_hanzi(signer)
    date_width = _estimate_date_width_hanzi(date)
    date_width_source = "mapping"
    if date_width is None:
        date_width = _estimate_text_width_hanzi(date)
        date_width_source = "fallback"

    # Calibrate "1 hanzi" in twips from the existing signer indent (≈ right 空四字).
    twips_per_hanzi = float(signer_right_indent_twips) / 4.0 if signer_right_indent_twips else 0.0

    # With right-aligned paragraphs, "right_indent" moves the text block left from the page right edge.
    # To keep the *centers* aligned:
    #   R - indent_s - w_s/2 == R - indent_d - w_d/2  =>  indent_d = indent_s + (w_s - w_d)/2
    indent = float(signer_right_indent_twips) + (float(signer_width) - float(date_width)) * 0.5 * twips_per_hanzi
    indent_twips = max(0, int(round(indent)))

    if debug:
        indent_hanzi_equiv = (indent_twips / twips_per_hanzi) if twips_per_hanzi else None
        print("[luokuan-indent] mode=center")
        print(f"[luokuan-indent] signer='{signer}' date='{date}'")
        print(
            f"[luokuan-indent] widths(hanzi-equiv): signer={signer_width:.2f} date={date_width:.2f} source={date_width_source}"
        )
        print(
            f"[luokuan-indent] signer_right_indent_twips={signer_right_indent_twips} twips_per_hanzi={twips_per_hanzi:.3f}"
        )
        if indent_hanzi_equiv is not None:
            print(
                f"[luokuan-indent] date_right_indent_twips={indent_twips} (~{indent_hanzi_equiv:.2f} hanzi)"
            )
        else:
            print(f"[luokuan-indent] date_right_indent_twips={indent_twips}")

    return indent_twips





def _choose_font_for_paragraph(text):
    stripped = text.strip()
    if stripped.startswith(tuple("一二三四五六七八九十")) and "、" in stripped[:3]:
        return FONT_LEVEL1
    if stripped.startswith("（") and "）" in stripped[:4]:
        return FONT_BODY
    return FONT_BODY


def _normalize_recipients(value):
    if value is None:
        return ""
    if isinstance(value, list):
        return "、".join([str(item).strip() for item in value if str(item).strip()])
    return str(value).strip()


def _normalize_list(value):
    if value is None:
        return []
    if isinstance(value, list):
        return [str(item).strip() for item in value if str(item).strip()]
    text = str(value).strip()
    return [text] if text else []


def _parse_body_text(text):
    if not text:
        return []
    lines = []
    for line in text.splitlines():
        cleaned = line.strip()
        if cleaned:
            lines.append(cleaned)
    return lines


def _load_json_input(path):
    if path == "-":
        content = sys.stdin.read()
    else:
        content = Path(path).read_text(encoding="utf-8")
    return json.loads(content)


def _read_text_input(path: str) -> str:
    if path == "-":
        return sys.stdin.read()
    return Path(path).read_text(encoding="utf-8")


def _parse_simple_front_matter(text: str):
    stripped = text.lstrip("\ufeff")
    if not stripped.startswith("---\n") and stripped != "---":
        return {}, text

    lines = stripped.splitlines()
    if not lines or lines[0].strip() != "---":
        return {}, text

    meta = {}
    i = 1
    current_list_key = None
    while i < len(lines):
        line = lines[i]
        if line.strip() == "---":
            i += 1
            break
        if not line.strip() or line.lstrip().startswith("#"):
            i += 1
            continue

        if line.startswith("  - ") or line.startswith("- "):
            if current_list_key:
                item = line.split("-", 1)[1].strip()
                if item:
                    meta.setdefault(current_list_key, []).append(item)
            i += 1
            continue

        current_list_key = None
        if ":" in line:
            key, value = line.split(":", 1)
            key = key.strip()
            value = value.strip()
            if not key:
                i += 1
                continue
            if value == "":
                meta[key] = []
                current_list_key = key
                i += 1
                continue
            if value.startswith("[") and value.endswith("]"):
                inner = value[1:-1].strip()
                if inner:
                    meta[key] = [
                        part.strip().strip("\"'")
                        for part in inner.split(",")
                        if part.strip()
                    ]
                else:
                    meta[key] = []
                i += 1
                continue
            meta[key] = value.strip("\"'")
        i += 1

    rest_text = "\n".join(lines[i:])
    return meta, rest_text


def _parse_controlled_markdown(text: str):
    meta, body_text = _parse_simple_front_matter(text)

    blocks = []
    for raw_line in body_text.splitlines():
        line = raw_line.rstrip()
        if not line.strip():
            continue
        if line.startswith("#"):
            level = 0
            for ch in line:
                if ch == "#":
                    level += 1
                else:
                    break
            if 1 <= level <= 5:
                content = line[level:].lstrip()
                if content:
                    blocks.append({"type": f"h{level}", "text": content})
                    continue
        blocks.append({"type": "p", "text": line.strip()})

    title = ""
    for b in blocks:
        if b["type"] == "h1":
            title = b["text"].strip()
            break

    data = {
        "title": title,
        "doc_number": meta.get("doc_number"),
        "recipients": meta.get("recipients"),
        "attachments": meta.get("attachments"),
        "signer": meta.get("signer"),
        "date": meta.get("date"),
        "contact": meta.get("contact"),
        "blocks": blocks,
    }
    return data


def _load_markdown_input(path: str):
    content = _read_text_input(path)
    return _parse_controlled_markdown(content)


def _build_with_template(template_path, output_path, data):
    _register_docx_namespaces()
    temp_dir = tempfile.mkdtemp(prefix="gongwen_build_")

    try:
        with zipfile.ZipFile(template_path, "r") as zf:
            zf.extractall(temp_dir)

        doc_xml_path = os.path.join(temp_dir, "word", "document.xml")
        tree = ET.parse(doc_xml_path)
        root = tree.getroot()
        body = root.find(f"{W_NS}body")

        paragraphs = body.findall(f"{W_NS}p")
        if len(paragraphs) < 2:
            shutil.rmtree(temp_dir, ignore_errors=True)
            raise ValueError("模板缺少红头或红线，无法生成文档")

        sect_pr = body.find(f"{W_NS}sectPr")

        for child in list(body):
            if child != sect_pr:
                body.remove(child)

        if sect_pr is not None:
            body.remove(sect_pr)

        body.append(deepcopy(paragraphs[0]))
        body.append(deepcopy(paragraphs[1]))

        title = str(data.get("title", "")).strip()
        doc_number = str(data.get("doc_number", "")).strip()
        recipients = _normalize_recipients(data.get("recipients"))
        body_list = data.get("body", [])
        attachments = _normalize_list(data.get("attachments"))
        signer = str(data.get("signer", "")).strip()
        date = str(data.get("date", "")).strip()
        contact = str(data.get("contact", "")).strip()
        blocks = data.get("blocks")

        if isinstance(body_list, str):
            body_list = _parse_body_text(body_list)
        elif isinstance(body_list, list):
            body_list = [str(item).strip() for item in body_list if str(item).strip()]
        else:
            body_list = []

        content_paras = []

        if doc_number:
            p = ET.Element(f"{W_NS}p")
            pPr = ET.SubElement(p, f"{W_NS}pPr")
            jc = ET.SubElement(pPr, f"{W_NS}jc")
            jc.set(f"{W_NS}val", "right")
            spacing = ET.SubElement(pPr, f"{W_NS}spacing")
            spacing.set(f"{W_NS}line", "540")
            spacing.set(f"{W_NS}lineRule", "exact")
            r = ET.SubElement(p, f"{W_NS}r")
            rPr = ET.SubElement(r, f"{W_NS}rPr")
            rFonts = ET.SubElement(rPr, f"{W_NS}rFonts")
            rFonts.set(f"{W_NS}ascii", FONT_ASCII)
            rFonts.set(f"{W_NS}hAnsi", FONT_ASCII)
            rFonts.set(f"{W_NS}eastAsia", FONT_BODY)
            sz = ET.SubElement(rPr, f"{W_NS}sz")
            sz.set(f"{W_NS}val", "32")
            szCs = ET.SubElement(rPr, f"{W_NS}szCs")
            szCs.set(f"{W_NS}val", "32")
            t = ET.SubElement(r, f"{W_NS}t")
            t.text = doc_number
            content_paras.append(p)

            p_blank = ET.Element(f"{W_NS}p")
            pPr_blank = ET.SubElement(p_blank, f"{W_NS}pPr")
            spacing_blank = ET.SubElement(pPr_blank, f"{W_NS}spacing")
            spacing_blank.set(f"{W_NS}line", "540")
            spacing_blank.set(f"{W_NS}lineRule", "exact")
            content_paras.append(p_blank)

        def add_para(
            text,
            font_cn,
            font_en,
            size,
            align="both",
            first_indent=None,
            right_indent=None,
        ):
            p = ET.Element(f"{W_NS}p")
            pPr = ET.SubElement(p, f"{W_NS}pPr")
            if align:
                jc = ET.SubElement(pPr, f"{W_NS}jc")
                jc.set(f"{W_NS}val", align)
            spacing = ET.SubElement(pPr, f"{W_NS}spacing")
            spacing.set(f"{W_NS}line", "540")
            spacing.set(f"{W_NS}lineRule", "exact")
            if first_indent is not None or right_indent is not None:
                ind = ET.SubElement(pPr, f"{W_NS}ind")
                if first_indent is not None:
                    ind.set(f"{W_NS}firstLine", str(first_indent))
                    ind.set(f"{W_NS}firstLineChars", "200")
                if right_indent is not None:
                    ind.set(f"{W_NS}right", str(right_indent))
            r = ET.SubElement(p, f"{W_NS}r")
            rPr = ET.SubElement(r, f"{W_NS}rPr")
            rFonts = ET.SubElement(rPr, f"{W_NS}rFonts")
            rFonts.set(f"{W_NS}ascii", font_en)
            rFonts.set(f"{W_NS}hAnsi", font_en)
            rFonts.set(f"{W_NS}eastAsia", font_cn)
            sz = ET.SubElement(rPr, f"{W_NS}sz")
            sz.set(f"{W_NS}val", str(size))
            szCs = ET.SubElement(rPr, f"{W_NS}szCs")
            szCs.set(f"{W_NS}val", str(size))
            t = ET.SubElement(r, f"{W_NS}t")
            t.text = _normalize_quotes(text)
            content_paras.append(p)

        if isinstance(blocks, list) and blocks:
            title_done = False
            for block in blocks:
                if block.get("type") == "h1" and not title_done:
                    text = str(block.get("text", "")).strip()
                    if text:
                        add_para(text, FONT_HEADER, FONT_ASCII, 44, "center")
                        p_blank = ET.Element(f"{W_NS}p")
                        pPr_blank = ET.SubElement(p_blank, f"{W_NS}pPr")
                        spacing_blank = ET.SubElement(pPr_blank, f"{W_NS}spacing")
                        spacing_blank.set(f"{W_NS}line", "540")
                        spacing_blank.set(f"{W_NS}lineRule", "exact")
                        content_paras.append(p_blank)
                    title_done = True
                    break

            if not title_done and title:
                add_para(title, FONT_HEADER, FONT_ASCII, 44, "center")
                p_blank = ET.Element(f"{W_NS}p")
                pPr_blank = ET.SubElement(p_blank, f"{W_NS}pPr")
                spacing_blank = ET.SubElement(pPr_blank, f"{W_NS}spacing")
                spacing_blank.set(f"{W_NS}line", "540")
                spacing_blank.set(f"{W_NS}lineRule", "exact")
                content_paras.append(p_blank)

            if recipients:
                add_para(f"{recipients}：", FONT_BODY, FONT_ASCII, 32, "left")

            skipped_first_h1 = False
            for block in blocks:
                btype = block.get("type")
                text = str(block.get("text", "")).strip()
                if not text:
                    continue

                if btype == "h1" and not skipped_first_h1:
                    skipped_first_h1 = True
                    continue

                if btype and btype.startswith("h"):
                    try:
                        level = int(btype[1:])
                    except ValueError:
                        level = 0

                    if level == 2:
                        add_para(text, FONT_LEVEL1, FONT_ASCII, 32, "both", 640)
                    elif level >= 3:
                        add_para(text, FONT_BODY, FONT_ASCII, 32, "both", 640)
                    else:
                        add_para(text, FONT_BODY, FONT_ASCII, 32, "both", 640)
                else:
                    add_para(text, FONT_BODY, FONT_ASCII, 32, "both", 640)
        else:
            if title:
                add_para(title, FONT_HEADER, FONT_ASCII, 44, "center")
                p_blank = ET.Element(f"{W_NS}p")
                pPr_blank = ET.SubElement(p_blank, f"{W_NS}pPr")
                spacing_blank = ET.SubElement(pPr_blank, f"{W_NS}spacing")
                spacing_blank.set(f"{W_NS}line", "540")
                spacing_blank.set(f"{W_NS}lineRule", "exact")
                content_paras.append(p_blank)

            if recipients:
                add_para(f"{recipients}：", FONT_BODY, FONT_ASCII, 32, "left")

            for paragraph_text in body_list:
                font_name = _choose_font_for_paragraph(paragraph_text)
                align = "both" if font_name == FONT_BODY else "left"
                add_para(paragraph_text, font_name, FONT_ASCII, 32, align, 640)

        if attachments:
            p_blank = ET.Element(f"{W_NS}p")
            pPr_blank = ET.SubElement(p_blank, f"{W_NS}pPr")
            spacing_blank = ET.SubElement(pPr_blank, f"{W_NS}spacing")
            spacing_blank.set(f"{W_NS}line", "540")
            spacing_blank.set(f"{W_NS}lineRule", "exact")
            content_paras.append(p_blank)

            if len(attachments) == 1:
                add_para(
                    f"附件：{attachments[0]}", FONT_BODY, FONT_ASCII, 32, "both", 640
                )
            else:
                for idx, name in enumerate(attachments, 1):
                    prefix = "附件：" if idx == 1 else ""
                    text = f"{prefix}{idx}. {name}" if prefix else f"{idx}. {name}"
                    add_para(text, FONT_BODY, FONT_ASCII, 32, "both", 640)

        if signer or date or contact:
            p_blank = ET.Element(f"{W_NS}p")
            pPr_blank = ET.SubElement(p_blank, f"{W_NS}pPr")
            spacing_blank = ET.SubElement(pPr_blank, f"{W_NS}spacing")
            spacing_blank.set(f"{W_NS}line", "540")
            spacing_blank.set(f"{W_NS}lineRule", "exact")
            content_paras.append(p_blank)

            if signer:
                add_para(signer, FONT_BODY, FONT_ASCII, 32, "right", None, 1247)
            if date:
                date_right = _compute_date_right_indent_twips(
                    signer=signer,
                    date=date,
                    signer_right_indent_twips=int(getattr(RIGHT_INDENT_SIGNER, "twips", RIGHT_INDENT_SIGNER)),
                    debug=bool(data.get("_debug_indent")),
                )
                add_para(date, FONT_BODY, FONT_ASCII, 32, "right", None, date_right)

            if contact:
                p_blank2 = ET.Element(f"{W_NS}p")
                pPr_blank2 = ET.SubElement(p_blank2, f"{W_NS}pPr")
                spacing_blank2 = ET.SubElement(pPr_blank2, f"{W_NS}spacing")
                spacing_blank2.set(f"{W_NS}line", "540")
                spacing_blank2.set(f"{W_NS}lineRule", "exact")
                content_paras.append(p_blank2)
                add_para(
                    f"（联系人：{contact}）", FONT_BODY, FONT_ASCII, 32, "both", 640
                )

        for p in content_paras:
            body.append(p)

        if sect_pr is not None:
            body.append(sect_pr)

        tree.write(doc_xml_path, encoding="UTF-8", xml_declaration=True)

        with zipfile.ZipFile(str(output_path), "w", zipfile.ZIP_DEFLATED) as zf:
            for root_dir, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    arc_name = os.path.relpath(file_path, temp_dir)
                    zf.write(file_path, arc_name)

        print(f"已生成（含红头）：{output_path}")

        inject_banji_section(str(output_path))

    finally:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)


def estimate_page_count(doc_xml):
    """Estimate page count from paragraph count in document body."""
    body_match = re.search(r"<w:body[^>]*>([\s\S]*)</w:body>", doc_xml)
    if not body_match:
        return 1
    body_inner = body_match.group(1)
    para_count = len(re.findall(r"<w:p[\s>]", body_inner))
    LINES_PER_PAGE = 23
    return max(1, (para_count + LINES_PER_PAGE - 1) // LINES_PER_PAGE)


def inject_banji_section(output_path):
    """Post-process docx: wrap existing sectPr in a section break, add section 2 with 版记 footer."""
    with open(output_path, "rb") as f:
        zip_data = f.read()

    with zipfile.ZipFile(BytesIO(zip_data), "r") as zf:
        files = {name: zf.read(name) for name in zf.namelist()}

    rels_xml = files["word/_rels/document.xml.rels"].decode("utf-8")
    max_id = 0
    for m in re.finditer(r'Id="rId(\d+)"', rels_xml):
        max_id = max(max_id, int(m.group(1)))
    new_rid = f"rId{max_id + 1}"

    banji_para = (
        '<w:p><w:pPr><w:jc w:val="left"/></w:pPr>'
        '<w:r><w:rPr><w:rFonts w:eastAsia="黑体"/><w:sz w:val="28"/></w:rPr><w:t>公开方式：</w:t></w:r>'
        '<w:r><w:rPr><w:rFonts w:eastAsia="仿宋_GB2312"/><w:sz w:val="28"/></w:rPr><w:t xml:space="preserve">依申请公开</w:t></w:r>'
        "</w:p>"
    )
    # 从模板默认页脚提取原始 <w:ftr> 命名空间声明和页码内容
    template_footer = files.get("word/footer2.xml")
    if template_footer:
        tf = template_footer.decode("utf-8")
        root_match = re.match(r"(<\?xml[^?]*\?>\s*)?(<w:ftr[^>]*>)", tf)
        if root_match:
            ftr_open = root_match.group(2)
            content_start = root_match.end()
            content_end = tf.rfind("</w:ftr>")
            inner_content = tf[content_start:content_end].replace(
                'w:jc w:val="right"', 'w:jc w:val="left"'
            )
            footer_xml = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>{ftr_open}{banji_para}{inner_content}</w:ftr>'
        else:
            footer_xml = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{banji_para}</w:ftr>'
    else:
        footer_xml = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{banji_para}</w:ftr>'
    files["word/footer99.xml"] = footer_xml.encode("utf-8")

    ct_xml = files["[Content_Types].xml"].decode("utf-8")
    ct_override = (
        '<Override PartName="/word/footer99.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
    )
    if ct_override not in ct_xml:
        ct_xml = ct_xml.replace(
            "</Types>", ct_override + "</Types>", 1
        )
    files["[Content_Types].xml"] = ct_xml.encode("utf-8")

    footer_rel = (
        f'<Relationship Id="{new_rid}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" '
        'Target="footer99.xml"/>'
    )
    rels_updated = rels_xml.replace("</Relationships>", footer_rel + "</Relationships>")
    files["word/_rels/document.xml.rels"] = rels_updated.encode("utf-8")

    doc_xml = files["word/document.xml"].decode("utf-8")
    sect_pr_match = re.search(r"<w:sectPr\b.*?</w:sectPr>", doc_xml, re.DOTALL)
    if sect_pr_match:
        orig_sect_pr = sect_pr_match.group(0)
        # 估计页数：偶数页用连续分节符（不新增空白页），奇数页用下一页分节符
        page_count = estimate_page_count(doc_xml)
        mod_sect_pr = orig_sect_pr
        if page_count % 2 == 0:
            mod_sect_pr = re.sub(r"^(<w:sectPr)", r'\1<w:type w:val="continuous"/>', orig_sect_pr)
        section_break = f"<w:p><w:pPr>{mod_sect_pr}</w:pPr></w:p>"
        empty_para = "<w:p><w:pPr><w:rPr><w:sz w:val=\"28\"/></w:rPr></w:pPr></w:p>"
        sec2_sect_pr = (
            f"<w:sectPr>"
            f'<w:footerReference w:type="even" r:id="{new_rid}"/>'
            f'<w:footerReference w:type="default" r:id="{new_rid}"/>'
            f'<w:footerReference w:type="first" r:id="{new_rid}"/>'
            f'<w:pgSz w:w="11906" w:h="16838"/>'
            f'<w:pgMar w:top="2098" w:right="1588" w:bottom="2041" w:left="1588" w:header="851" w:footer="1644" w:gutter="0"/>'
            f"</w:sectPr>"
        )
        doc_xml = doc_xml.replace(orig_sect_pr, section_break + empty_para + sec2_sect_pr, 1)
        files["word/document.xml"] = doc_xml.encode("utf-8")

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in files.items():
            zf.writestr(name, data)


def build_document(data, output_path: Path, template_path: str = None):
    if template_path and os.path.exists(template_path):
        _build_with_template(template_path, output_path, data)
        return

    raise FileNotFoundError(f"模板文件未找到: {template_path}")


def main():
    script_dir = Path(__file__).parent.parent
    default_template = script_dir / "templates" / "岭南学院公文模板.docx"

    parser = argparse.ArgumentParser(description="生成符合岭南学院公文格式的Word文档")
    parser.add_argument("--input", help="输入JSON文件路径，或 '-' 从STDIN读取")
    parser.add_argument(
        "--md", help="输入Markdown文件路径（受控格式），或 '-' 从STDIN读取"
    )
    parser.add_argument("--title", help="标题")
    parser.add_argument("--doc-number", help="文号（如：岭院函〔2026〕1号）")
    parser.add_argument("--recipients", help="主送机关，多个用顿号分隔")
    parser.add_argument("--body", help="正文，多段用换行分隔")
    parser.add_argument("--body-file", help="正文文本文件，每行一段")
    parser.add_argument("--attachment", action="append", help="附件名称，可多次传入")
    parser.add_argument("--signer", help="发文单位")
    parser.add_argument("--date", help="发文日期")
    parser.add_argument("--contact", help="联系人信息")
    parser.add_argument(
        "--debug-indent",
        action="store_true",
        help="打印落款中发文单位、日期缩进计算的中间过程",
    )
    parser.add_argument(
        "-o",
        "--output",
        default="示例公文.docx",
        help="输出文件路径（默认：示例公文.docx）",
    )
    args = parser.parse_args()

    data = {}
    has_custom = any(
        [
            args.title,
            args.doc_number,
            args.recipients,
            args.body,
            args.body_file,
            args.attachment,
            args.signer,
            args.date,
            args.contact,
        ]
    )

    if args.md:
        data = _load_markdown_input(args.md)
    elif args.input:
        suffix = Path(args.input).suffix.lower() if args.input != "-" else ""
        if suffix in MARKDOWN_EXTENSIONS:
            data = _load_markdown_input(args.input)
        else:
            data = _load_json_input(args.input)
    elif has_custom:
        data = {}
    else:
        data = DEFAULT_DATA.copy()

    if args.title:
        data["title"] = args.title
    if args.doc_number:
        data["doc_number"] = args.doc_number
    if args.recipients:
        data["recipients"] = args.recipients

    if args.body_file:
        body_text = Path(args.body_file).read_text(encoding="utf-8")
        data["body"] = _parse_body_text(body_text)
    elif args.body:
        data["body"] = _parse_body_text(args.body)

    if args.attachment:
        data["attachments"] = args.attachment
    if args.signer:
        data["signer"] = args.signer
    if args.date:
        data["date"] = args.date
    if args.contact:
        data["contact"] = args.contact
    if args.debug_indent:
        data["_debug_indent"] = True

    output_path = Path(args.output).resolve()

    build_document(data, output_path, str(default_template))


if __name__ == "__main__":
    main()
