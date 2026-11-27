"""Document rendering and batch execution; importing this module has no UI side effects."""
import os
import io
import re
import uuid
import shutil
import tempfile
import zipfile
import subprocess
import json
from datetime import datetime
from zoneinfo import ZoneInfo
import fitz
from pathlib import Path
from lxml import etree
import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate, InlineImage, RichText
from docx.shared import Mm
import barcode
from barcode.writer import ImageWriter
from PIL import Image, ImageOps
from rules import (CDU_COLS, MERGED_SCHEMA, NAME_COL, derive_outer_size, get_orms, int_str, normalize_spaces, preflight, resolve_bl, resolved_name, safe_str, split_orms_bold, validate_code128, validate_itf)
from rules import (build_plan, input_fingerprint)

def get_resource_path(relative_path):
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)


def crop_vertical_margin(image_path):
    with Image.open(image_path) as img:
        gray = img.convert("L")
        inverted = ImageOps.invert(gray)
        bbox = inverted.getbbox()
        if bbox:
            img.crop((0, bbox[1], img.width, bbox[3])).save(image_path)


def _find_soffice():
    """定位 LibreOffice 可执行文件，兼容云端(libreoffice)与本地 Mac/Win。"""
    for cand in ("libreoffice", "soffice"):
        p = shutil.which(cand)
        if p:
            return p
    for p in ("/Applications/LibreOffice.app/Contents/MacOS/soffice",
              "/opt/homebrew/bin/soffice", "/usr/local/bin/soffice",
              r"C:\Program Files\LibreOffice\program\soffice.exe"):
        if os.path.exists(p):
            return p
    return None


def convert_docx_to_pdf(docx_path, output_dir, profile_dir):
    soffice = _find_soffice()
    if not soffice:
        raise RuntimeError("未找到 LibreOffice。本地测试请先安装：Mac 执行 "
                           "`brew install --cask libreoffice`（或到 libreoffice.org 下载）。")
    result = subprocess.run(
        [soffice, f"-env:UserInstallation={Path(profile_dir).resolve().as_uri()}",
         "--headless", "--convert-to", "pdf", "--outdir", output_dir, docx_path],
        capture_output=True, text=True, timeout=120,
    )
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice 转换失败: {result.stderr.strip()}")





def calc_fit_font_pt(text, max_mm, font_path=None, max_pt=8.0, min_pt=4.0):
    if not text:
        return max_pt
    try:
        from PIL import ImageFont
        RENDER_DPI, MARGIN = 288.0, 0.80
        max_px = max_mm * MARGIN * RENDER_DPI / 25.4
        pt = max_pt
        while pt >= min_pt:
            size_px = max(1, round(pt * RENDER_DPI / 72.0))
            if font_path and os.path.exists(font_path):
                fo = ImageFont.truetype(font_path, size=size_px)
                bb = fo.getbbox(text)
                if (bb[2] - bb[0]) <= max_px:
                    return pt
            else:
                break
            pt -= 0.5
        return min_pt
    except Exception:
        return max(min_pt, min(max_pt, max_mm / max(len(text), 1) / 0.31))


@st.cache_resource
def _install_fonts():
    """把 fonts/ 里的字体装到 LibreOffice 能识别的位置，保证云端与本地渲染字体一致。
    macOS → ~/Library/Fonts；Linux → ~/.fonts + fc-cache。"""
    fonts_src = get_resource_path("fonts")
    if not os.path.exists(fonts_src):
        return
    import sys
    targets = [os.path.expanduser("~/.fonts")]
    if sys.platform == "darwin":
        targets.append(os.path.expanduser("~/Library/Fonts"))
    for dst in targets:
        os.makedirs(dst, exist_ok=True)
        for f in os.listdir(fonts_src):
            if f.lower().endswith((".ttf", ".otf")):
                try:
                    shutil.copy2(os.path.join(fonts_src, f), os.path.join(dst, f))
                except Exception:
                    pass
    if shutil.which("fc-cache"):
        subprocess.run(["fc-cache", "-f"], capture_output=True, timeout=60)


_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _cell_avail_mm(root, p_el, fallback_mm):
    """按品名段落实际所在表格单元格的真实宽度，算出可用宽度（mm）：
    单元格宽 - 单元格左右边距(tcMar，缺省取表格 tblCellMar，再缺省 108twips*2) - 首行缩进(w:ind firstLine)。
    取不到几何信息时退回 fallback_mm，保证兼容旧模板。"""
    W = _W
    try:
        # 找最近的 w:tc 祖先
        tc = p_el.getparent()
        while tc is not None and tc.tag != f"{W}tc":
            tc = tc.getparent()
        if tc is None:
            return fallback_mm
        tcPr = tc.find(f"{W}tcPr")
        tcW_el = tcPr.find(f"{W}tcW") if tcPr is not None else None
        if tcW_el is None or tcW_el.get(f"{W}w") is None:
            return fallback_mm
        cell_dxa = float(tcW_el.get(f"{W}w"))

        # 单元格左右边距：优先 tcMar，否则表格级 tblCellMar，否则默认 108+108
        mar_l = mar_r = 108.0
        tcMar = tcPr.find(f"{W}tcMar") if tcPr is not None else None
        if tcMar is not None:
            le = tcMar.find(f"{W}left"); re_ = tcMar.find(f"{W}right")
            if le is not None and le.get(f"{W}w") is not None:
                mar_l = float(le.get(f"{W}w"))
            if re_ is not None and re_.get(f"{W}w") is not None:
                mar_r = float(re_.get(f"{W}w"))
        else:
            tbl = tc.getparent()
            while tbl is not None and tbl.tag != f"{W}tbl":
                tbl = tbl.getparent()
            if tbl is not None:
                tblPr = tbl.find(f"{W}tblPr")
                cm = tblPr.find(f"{W}tblCellMar") if tblPr is not None else None
                if cm is not None:
                    le = cm.find(f"{W}left"); re_ = cm.find(f"{W}right")
                    if le is not None and le.get(f"{W}w") is not None:
                        mar_l = float(le.get(f"{W}w"))
                    if re_ is not None and re_.get(f"{W}w") is not None:
                        mar_r = float(re_.get(f"{W}w"))

        # 首行缩进
        indent_dxa = 0.0
        pPr = p_el.find(f"{W}pPr")
        ind = pPr.find(f"{W}ind") if pPr is not None else None
        if ind is not None and ind.get(f"{W}firstLine") is not None:
            indent_dxa = float(ind.get(f"{W}firstLine"))

        avail_dxa = cell_dxa - mar_l - mar_r - indent_dxa
        if avail_dxa <= 0:
            return fallback_mm
        return avail_dxa / 1440.0 * 25.4
    except Exception:
        return fallback_mm


def _set_paragraph_font_pt(p_el, pt):
    """把某个 w:p 内所有含文字的 run 字号设为 pt（保留其余格式）。返回是否改动。"""
    W = _W
    half = str(int(round(pt * 2)))
    changed = False
    for r in p_el.iter(f"{W}r"):
        if not r.findall(f"{W}t"):
            continue
        rPr = r.find(f"{W}rPr")
        if rPr is None:
            rPr = etree.SubElement(r, f"{W}rPr"); r.insert(0, rPr)
        for tag in ("sz", "szCs"):
            e = rPr.find(f"{W}{tag}")
            if e is None:
                e = etree.SubElement(rPr, f"{W}{tag}")
            e.set(f"{W}val", half)
        changed = True
    return changed


def shrink_name_in_docx(docx_path, name_text, max_mm, max_pt=4.0, min_pt=2.0, forced_pt=None):
    """外箱/内盒品名在文本框/单元格里、docxtpl RichText 不生效，改为渲染后处理：
    按品名所在单元格的真实可用宽度自适应缩小字号，保证单行不换行（保留原有白色/字体/加粗）。
    max_mm 仅作为取不到真实几何信息时的兜底宽度；多处出现时取其中最窄单元格为准，保证各份一致不换行。
    forced_pt：跳过自动测算，直接强制使用该字号（配合 ensure_name_no_wrap 的换行校验回退使用）。
    返回实际使用的字号（pt）。"""
    W = _W
    target = normalize_spaces(name_text)
    if not target:
        return None
    zin = zipfile.ZipFile(docx_path)
    parts = {it.filename: zin.read(it.filename) for it in zin.infolist()}
    infos = zin.infolist(); zin.close()
    root = etree.fromstring(parts["word/document.xml"])

    matched_paras = [p for p in root.iter(f"{W}p")
                      if normalize_spaces("".join(t.text or "" for t in p.iter(f"{W}t"))) == target]
    if not matched_paras:
        return None

    if forced_pt is not None:
        fit = forced_pt
    else:
        avail_list = [_cell_avail_mm(root, p, max_mm) for p in matched_paras]
        real_max_mm = min(avail_list) if avail_list else max_mm
        real_max_mm = min(real_max_mm, max_mm) if max_mm else real_max_mm
        fit = calc_fit_font_pt(target, real_max_mm, get_resource_path("fonts/微软雅黑.ttf"),
                                max_pt=max_pt, min_pt=min_pt)

    changed = 0
    for p in matched_paras:
        if _set_paragraph_font_pt(p, fit):
            changed += 1
    if changed:
        parts["word/document.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        tmp = docx_path + ".t"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
            for it in infos:
                zo.writestr(it, parts[it.filename])
        os.replace(tmp, docx_path)
    return fit


def _pdf_word_boxes(pdf_path):
    words = []
    with fitz.open(pdf_path) as doc:
        for page_i, page in enumerate(doc):
            for x0, y0, x1, y1, word, *_ in page.get_text("words"):
                words.append({"page": page_i, "text": word, "xMin": x0, "yMin": y0})
    return words


def name_wrapped_in_pdf(pdf_path, name_text, y_tol=1.0):
    """校验品名在生成的 PDF 里是否被换行拆成了两行。
    做法：按 y 坐标把词聚成"行"，若某一行的文本恰好是品名的真前缀/后缀（而非完整品名），
    说明品名被拆到了相邻行 —— 判定为换行。"""
    target = normalize_spaces(name_text)
    if not target:
        return False
    words = _pdf_word_boxes(pdf_path)
    if not words:
        return False
    lines = []  # list of (page, y_key, [words])
    for w in words:
        found = None
        for entry in lines:
            if entry[0] == w["page"] and abs(entry[1] - w["yMin"]) <= y_tol:
                found = entry
                break
        if found is None:
            found = (w["page"], w["yMin"], [])
            lines.append(found)
        found[2].append(w)
    for _page, _y, ws in lines:
        line_text = normalize_spaces(" ".join(x["text"] for x in sorted(ws, key=lambda x: x["xMin"])))
        if line_text and line_text != target and (target.startswith(line_text) or target.endswith(line_text)):
            return True
    return False


def ensure_name_no_wrap(docx_path, name_text, out_dir, profile, max_mm=34.0, max_pt=6.0, min_pt=2.0, step=0.5):
    """渲染并转 PDF 后用真实版面校验品名有没有被换行；如果换行了就调小字号重转，
    直到不换行或到达最小字号为止（保证"绝不允许换行"这条硬性要求，不完全依赖字体宽度估算）。"""
    target = normalize_spaces(name_text)
    if not target:
        convert_docx_to_pdf(docx_path, out_dir, profile)
        return
    fit = shrink_name_in_docx(docx_path, name_text, max_mm, max_pt=max_pt, min_pt=min_pt)
    convert_docx_to_pdf(docx_path, out_dir, profile)
    if fit is None:
        return
    base = os.path.splitext(os.path.basename(docx_path))[0]
    pdf_path = os.path.join(out_dir, f"{base}.pdf")
    tries = 0
    while os.path.exists(pdf_path) and name_wrapped_in_pdf(pdf_path, name_text) and fit > min_pt and tries < 8:
        fit = max(min_pt, fit - step)
        shrink_name_in_docx(docx_path, name_text, max_mm, forced_pt=fit)
        convert_docx_to_pdf(docx_path, out_dir, profile)
        tries += 1
    if name_wrapped_in_pdf(pdf_path, name_text):
        os.remove(pdf_path)
        raise ValueError("品名在最小字号下仍换行，请缩短品名或调整模板")


@st.cache_data
def build_blank_template():
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
    wb = Workbook()
    ws = wb.active
    ws.title = "数据源"
    schema = [c for c in MERGED_SCHEMA if c['kind'] != 'formula']
    for i, spec in enumerate(schema, 1):
        col = get_column_letter(i)
        cell = ws.cell(1, i, spec['n'])
        cell.font = Font(bold=True)
        cell.fill = PatternFill('solid', fgColor='C6E0B4' if spec['kind'] == 'data' else 'FFF2CC')
        ws.column_dimensions[col].width = 22
        for row in range(2, 502):
            ws.cell(row, i).number_format = '@' if spec.get('text') else ('0' if spec.get('intfmt') else 'General')
        if spec['kind'] == 'select':
            values = 'Y,N' if spec['dd'] == 'YN' else ','.join(spec['dd'])
            dv = DataValidation(type='list', formula1='"' + values + '"', allow_blank=True)
            dv.showErrorMessage = True
            ws.add_data_validation(dv)
            dv.add(f'{col}2:{col}501')
    ws.freeze_panes = 'A2'
    notes = wb.create_sheet('填写说明')
    for i, text in enumerate([
        '只填原始数据，无需 Excel 公式；网站统一重算条码、尺寸、国家和文件名。',
        '绿色为数据；黄色为选择项。Y=生成，N=不生成，留空=按规则自动判断。',
        '数量必须为正整数；有内盒时，外箱装量必须能被内盒装量整除。',
        '无内盒请留空内盒装量；内外装量相等也视为无内盒。',
        '号码列已按文本格式设置至第 501 行；不要转成数字以免丢失前导 0。',
        '旧表中的公式结果与命名列会被忽略。体检通过后才能生成。',
    ], 1): notes.cell(i, 1, text)
    notes.column_dimensions['A'].width = 110
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


CDU_CSV = get_resource_path("cdu_list.csv")


@st.cache_data
def load_cdu_df_bundled():
    if os.path.exists(CDU_CSV):
        try:
            df = pd.read_csv(CDU_CSV, dtype=str).fillna("")
            for c in CDU_COLS:
                if c not in df.columns:
                    df[c] = ""
            return df[CDU_COLS]
        except Exception:
            pass
    return pd.DataFrame(columns=CDU_COLS)


def _confirm_pdf(output_dir, work_dir, name_base):
    pdf_path = os.path.join(output_dir, f"{name_base}.pdf")
    if not os.path.exists(pdf_path):
        raise RuntimeError(f"未找到输出 PDF：{name_base}.pdf")
    with fitz.open(pdf_path) as doc:
        if len(doc) != 1:
            raise ValueError(f"标签应为单页，实际为 {len(doc)} 页：{name_base}")
        if not doc[0].get_text().strip():
            raise ValueError(f"标签缺少可读文字：{name_base}")
    return True





def gen_outer(row, work_dir, out_dir, profile):
    po_str = safe_str(row.get("PO号"))
    if not po_str:
        raise ValueError("PO号 为空，无法生成外箱标")
    is_uk = po_str.startswith("380") or po_str.startswith("520")
    is_ce = po_str.startswith("999")
    if not (is_uk or is_ce):
        raise ValueError(f"PO号 '{po_str}' 前缀非 380/520/999，无法判断 UK/CE")

    tpl = DocxTemplate(get_resource_path("模板_OuterCarton.docx"))
    ctx = {}
    if is_uk:
        ctx["TPNB"], ctx["TPND"] = safe_str(row.get("TPNB")), safe_str(row.get("TPND"))
        ctx["CON不加粗"], ctx["CON加粗"], ctx["SKU"] = "", "", ""
    else:  # CE
        ctx["TPNB"], ctx["TPND"] = "", ""
        ctx["SKU"] = safe_str(row.get("SKU"))
        con_n, con_b = safe_str(row.get("CON不加粗")), safe_str(row.get("CON加粗"))
        if not con_n and not con_b:
            con_n, con_b = split_orms_bold(get_orms(row))
        ctx["CON不加粗"], ctx["CON加粗"] = con_n, con_b

    ctx["产品英文名"] = safe_str(row.get("产品英文名"))
    ctx["外箱BD号"] = safe_str(row.get("外箱 BD 号", row.get("外箱BD号")))
    ctx["有CDU加印文字"] = safe_str(row.get("有CDU-加印文字", row.get("有CDU加印文字")))
    ctx.update({
        "PO号": po_str, "category": safe_str(row.get("category")),
        "EAN": safe_str(row.get("EAN")), "毛重": safe_str(row.get("毛重")),
        "净重": safe_str(row.get("净重")), "外箱尺寸": derive_outer_size(row),
        "其他备注": safe_str(row.get("其他备注")), "订单国家": safe_str(row.get("订单国家")),
        "总外箱数": int_str(row.get("总外箱数")),
    })
    bl = resolve_bl(row)
    bl_path = get_resource_path(f"2. Brand Logo/{bl}.png")
    ctx["BL"] = InlineImage(tpl, bl_path, height=Mm(3)) if bl and os.path.exists(bl_path) else ""
    se = safe_str(row.get("节日logo"))
    se_path = get_resource_path(f"3. Seasonal Logo/{se}.png")
    ctx["节日logo"] = InlineImage(tpl, se_path, height=Mm(6)) if se and os.path.exists(se_path) else ""
    hv_path = get_resource_path("4. 双人抬 Logo/双人抬.png")
    ctx["双人抬"] = InlineImage(tpl, hv_path, height=Mm(6)) if safe_str(row.get("双人抬标识")) == "双人抬" and os.path.exists(hv_path) else ""

    name = resolved_name("outer", row)
    docx = os.path.join(work_dir, f"{name}.docx")
    tpl.render(ctx, autoescape=True)
    tpl.save(docx)
    ensure_name_no_wrap(docx, safe_str(row.get("产品英文名")), out_dir, profile, max_mm=34.0, max_pt=6.5, min_pt=2.0)
    _confirm_pdf(out_dir, work_dir, name)
    return [{"命名": name, "状态": "成功", "说明": "UK" if is_uk else "CE"}]


def gen_inner(row, work_dir, out_dir, profile):
    tpl = DocxTemplate(get_resource_path("模板_InnerCarton.docx"))
    ctx = {
        # 品名在文本框内（docxtpl RichText 在文本框里不生效），仍用纯字符串；
        # 换行问题改在模板里把品名字号调小解决。
        "产品英文名": safe_str(row.get("产品英文名")),
        "category": safe_str(row.get("category")),
        "EAN": safe_str(row.get("EAN")),
        "有CDU加印文字内盒版": safe_str(row.get("有CDU加印文字内盒版")),
    }
    bl = resolve_bl(row)
    bl_path = get_resource_path(f"2. Brand Logo/{bl}.png")
    # 内盒黑条较窄，Logo 用 2.2mm 以免顶部溢出被交界处遮住
    ctx["BL"] = InlineImage(tpl, bl_path, height=Mm(2.2)) if bl and os.path.exists(bl_path) else ""
    name = resolved_name("inner", row)
    docx = os.path.join(work_dir, f"{name}.docx")
    tpl.render(ctx, autoescape=True)
    tpl.save(docx)
    ensure_name_no_wrap(docx, safe_str(row.get("产品英文名")), out_dir, profile, max_mm=34.0, max_pt=6.0, min_pt=2.0)
    _confirm_pdf(out_dir, work_dir, name)
    return [{"命名": name, "状态": "成功", "说明": ""}]


def gen_itf(row, work_dir, out_dir, profile, want_uk, want_ce):
    itf_val = safe_str(row.get("ITF"))
    if not itf_val:
        raise ValueError("ITF 为空")
    itf_spaced = safe_str(row.get("ITF加空格", itf_val))
    validate_itf(itf_val, itf_spaced)

    eng_raw = safe_str(row.get("产品英文名"))
    inner_raw = safe_str(row.get("内盒装量"))
    # 内盒装量无数字 → 该产品无内盒，ITF 上的 Case Size 直接用外箱装量
    if not any(ch.isdigit() for ch in inner_raw):
        inner_qty = int_str(row.get("外箱装量"))
    else:
        inner_qty = int_str(inner_raw)
    # 品名自适应字号：用与渲染相同的字体(微软雅黑)测量，保证宽度一致、长名绝不换行
    font_path = get_resource_path("fonts/微软雅黑.ttf")
    fit_pt = calc_fit_font_pt(normalize_spaces(eng_raw), 64.0, font_path, max_pt=7.5, min_pt=3.0)

    ITF_Class = barcode.get_barcode_class("itf")
    img_tmp = os.path.join(work_dir, f"barcode_itf_{uuid.uuid4().hex}")
    saved = ITF_Class(itf_val, writer=ImageWriter()).save(img_tmp, options={"write_text": False, "module_height": 20.0})
    crop_vertical_margin(saved)
    tpl_path = get_resource_path("模板_ITF.docx")
    results = []
    try:
        def one(kind, want, con_n, con_b, tpnb, tpnd):
            if not want:
                return
            name = resolved_name(kind, row)
            tpl = DocxTemplate(tpl_path)
            rt = RichText(); rt.add(eng_raw, bold=True, size=int(round(fit_pt * 2)), font="微软雅黑")
            ctx = {"TPNB": tpnb, "TPND": tpnd, "CON不加粗": con_n, "CON加粗": con_b,
                   "产品英文名": rt, "内盒装量": inner_qty, "ITF加空格": itf_spaced,
                   "Barcode_Img": InlineImage(tpl, saved, width=Mm(95), height=Mm(28))}
            docx = os.path.join(work_dir, f"{name}.docx")
            tpl.render(ctx, autoescape=True); tpl.save(docx)
            convert_docx_to_pdf(docx, out_dir, profile)
            _confirm_pdf(out_dir, work_dir, name)
            results.append({"命名": name, "状态": "成功", "说明": kind.upper()})

        one("itf_uk", want_uk, "", "", safe_str(row.get("TPNB")), safe_str(row.get("TPND")))
        cn, cb = safe_str(row.get("CON不加粗")), safe_str(row.get("CON加粗"))
        if not cn and not cb:
            cn, cb = split_orms_bold(get_orms(row))
        one("itf_ce", want_ce, cn, cb, "", "")
    finally:
        if os.path.exists(saved):
            os.remove(saved)
    return results


def gen_bd(row, work_dir, out_dir, profile):
    bd_val = safe_str(row.get("BD条码号"))
    name = resolved_name("bd", row)
    if not bd_val:
        # 无内盒(装量相等)时 BD条码号 公式为空 → 该产品无 BD，跳过（非错误）
        return [{"命名": name, "状态": "跳过", "说明": "无内盒或BD条码号为空，无需BD"}]
    bd_spaced = safe_str(row.get("BD条码号加空格", bd_val))
    validate_code128(bd_val, bd_spaced)

    # GS1-128 自动在起始符后添加 FNC1；号码和目视文本保持原样。
    GS1_128 = barcode.get_barcode_class("gs1_128")
    img_tmp = os.path.join(work_dir, f"barcode_bd_{uuid.uuid4().hex}")
    saved = GS1_128(bd_val, writer=ImageWriter()).save(img_tmp, options={"write_text": False, "module_height": 22.0})
    crop_vertical_margin(saved)
    tpl = DocxTemplate(get_resource_path("模板_BD.docx"))
    ctx = {"BD条码号加空格": bd_spaced, "Barcode_Img": InlineImage(tpl, saved, width=Mm(140), height=Mm(70))}
    try:
        tpl.render(ctx, autoescape=True)
        docx = os.path.join(work_dir, f"{name}.docx")
        tpl.save(docx)
        convert_docx_to_pdf(docx, out_dir, profile)
        _confirm_pdf(out_dir, work_dir, name)
        return [{"命名": name, "状态": "成功", "说明": ""}]
    finally:
        if os.path.exists(saved):
            os.remove(saved)


def run_all(df, selected, cdu_map=None, on_progress=None):
    issues = preflight(df, selected, cdu_map)
    if issues:
        raise ValueError(f"数据体检未通过（{len(issues)} 项），未生成文件")
    plan = build_plan(df, selected, cdu_map)
    if not plan:
        raise ValueError("没有待生成的标签，请检查类型和开关")
    _install_fonts()
    completed, results, pdfs = set(), [], {}
    with tempfile.TemporaryDirectory() as work:
        profile = os.path.join(work, "lo")
        os.makedirs(profile)
        for i, item in enumerate(plan):
            record = {"行": item['source_row'], "类型": item['type'], "命名": item['name']}
            if item['key'] in completed:
                record.update(状态="跳过", 说明="相同产品与标签内容已成功生成")
            else:
                # Each job gets its own output directory: partial failed PDFs never enter the ZIP.
                job = os.path.join(work, str(i))
                output = os.path.join(job, "out")
                os.makedirs(output)
                row = dict(item['row'])
                row[NAME_COL[item['kind']]] = item['name']
                try:
                    if item['kind'] == 'outer': gen_outer(row, job, output, profile)
                    elif item['kind'] == 'inner': gen_inner(row, job, output, profile)
                    elif item['kind'] == 'bd': gen_bd(row, job, output, profile)
                    else: gen_itf(row, job, output, profile, item['kind'] == 'itf_uk', item['kind'] == 'itf_ce')
                    _confirm_pdf(output, job, item['name'])
                    path = f"{item['type']}/{item['name']}.pdf"
                    pdfs[path] = Path(output, item['name'] + '.pdf').read_bytes()
                    completed.add(item['key'])
                    record.update(状态="成功", 说明=item['kind'])
                except Exception as exc:
                    record.update(状态="失败", 说明=str(exc))
            results.append(record)
            if on_progress: on_progress((i + 1) / len(plan))
    created_at = datetime.now(ZoneInfo("Asia/Shanghai")).isoformat(timespec="seconds")
    fingerprint = input_fingerprint(df, selected, cdu_map or {})
    manifest = {"generated_at": created_at, "input_fingerprint": fingerprint,
                "types": selected, "input_rows": len(df), "pdf_count": len(pdfs),
                "failed": sum(r['状态'] == '失败' for r in results),
                "skipped": sum(r['状态'] == '跳过' for r in results)}
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, data in pdfs.items(): archive.writestr(name, data)
        archive.writestr('生成明细.csv', pd.DataFrame(results).to_csv(index=False).encode('utf-8-sig'))
        archive.writestr('批次信息.json', json.dumps(manifest, ensure_ascii=False, indent=2))
    return dict(zip_data=buf.getvalue(), pdfs=pdfs, results=results, manifest=manifest,
                fingerprint=fingerprint, name=f"Tesco标签_{len(pdfs)}个.zip")



@st.cache_data(max_entries=20, show_spinner=False)
def preview_pdf(data):
    with fitz.open(stream=data, filetype="pdf") as doc:
        return doc[0].get_pixmap(matrix=fitz.Matrix(1.5, 1.5)).tobytes("png")
