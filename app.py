import os
import io
import re
import uuid
import shutil
import tempfile
import zipfile
import subprocess
from xml.sax.saxutils import escape as xml_escape
from lxml import etree

import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate, InlineImage, RichText
from docx.shared import Mm
import barcode
from barcode.writer import ImageWriter
from PIL import Image, ImageOps

# ==========================================
# 页面基础设置
# ==========================================
st.set_page_config(page_title="Tesco 标贴与纸箱生成系统", page_icon="📦", layout="wide")

for k, v in {"download_data": None, "download_name": "", "auth_ok": False}.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ==========================================
# 访问控制（可选口令门；未在 Secrets 配置则不拦截）
# ==========================================
def check_password():
    try:
        expected = st.secrets.get("app_password", None)
    except Exception:
        expected = None
    if not expected or st.session_state.get("auth_ok"):
        return
    st.title("🔒 Tesco 标贴系统 · 内部访问")
    pwd = st.text_input("请输入访问口令", type="password")
    if pwd:
        if pwd == expected:
            st.session_state.auth_ok = True
            st.rerun()
        else:
            st.error("口令错误，请重试。")
    st.stop()


check_password()


# ==========================================
# 基础辅助
# ==========================================
def safe_str(val):
    if pd.isna(val):
        return ""
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    if isinstance(val, int):
        return str(val)
    return str(val).strip()


def esc(text):
    return xml_escape(text, {'"': "&quot;", "'": "&apos;"})


def normalize_spaces(text):
    """把不间断空格等归一化为普通空格并压缩连续空格（仅用于文件名，不动标签正文）。"""
    return re.sub(r"\s+", " ", text.replace("\xa0", " ")).strip()


def is_yes(val):
    return safe_str(val).upper() in {"Y", "YES", "是", "1", "TRUE", "T", "✓", "√", "要", "需要"}


def is_neg(val):
    return safe_str(val).upper() in {"N", "NO", "否", "0", "FALSE", "F", "×", "X", "不"}


def wants(row, col, default):
    """每行开关：显式 Y→做、显式 N→不做、留空→按 default（自然规则）。
    这样粘贴原始数据不填开关也能按规则自动生成。"""
    v = safe_str(row.get(col))
    if not v:
        return default
    if is_yes(v):
        return True
    if is_neg(v):
        return False
    return default


def qty_equal(row):
    """内盒装量 == 外箱装量 → 该产品无内盒（返回 True 时不生成内盒标）。"""
    a, b = safe_str(row.get("外箱装量")), safe_str(row.get("内盒装量"))
    if not a or not b:
        return False
    try:
        return float(a) == float(b)
    except ValueError:
        return a == b


def no_inner(row):
    """判定该产品无内盒（返回 True 则不生成内盒标）。两种情况都算无内盒：
    ① 内盒装量为空或不含数字 → 产品直接入外箱（无内盒）；
    ② 内盒装量 == 外箱装量 → 无内盒。
    无内盒时即使界面勾了'是否要做内盒=Y'也跳过内盒设计稿。"""
    inner = safe_str(row.get("内盒装量"))
    if not inner or not any(ch.isdigit() for ch in inner):
        return True
    return qty_equal(row)


def sanitize_name(name):
    """净化文件名：路径分隔符拍平为下划线，防穿越，纯点名回退。"""
    cleaned = normalize_spaces(name).replace("/", "_").replace("\\", "_").strip()
    if not cleaned or set(cleaned) <= {"."}:
        return "unnamed"
    return cleaned


def get_resource_path(relative_path):
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)


def get_orms(row):
    exact = safe_str(row.get("CEORMSNO"))
    if exact:
        return exact
    for k in row.keys():
        ku = str(k).upper()
        if "CE" in ku and "ORMS" in ku:
            v = safe_str(row[k])
            if v:
                return v
    return ""


def split_orms_bold(orms_str):
    """后5位加粗，纯字符串切片，保留前导/内部 0。"""
    if len(orms_str) >= 5:
        return orms_str[:-5], orms_str[-5:]
    return orms_str, ""


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
        [soffice, f"-env:UserInstallation=file://{profile_dir}",
         "--headless", "--convert-to", "pdf", "--outdir", output_dir, docx_path],
        capture_output=True, text=True, timeout=120,
    )
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice 转换失败: {result.stderr.strip()}")


def create_zip(source_dir):
    zip_path = os.path.join(tempfile.gettempdir(), f"Tesco_Labels_{uuid.uuid4().hex}.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for root, _d, files in os.walk(source_dir):
            for file in files:
                if file.endswith(".pdf"):
                    fp = os.path.join(root, file)
                    zipf.write(fp, os.path.relpath(fp, source_dir))
    return zip_path


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


# ---- 条码一致性校验 ----
def _strip_spaces(s):
    return "".join(ch for ch in s if not ch.isspace() and ch != "\xa0")


def validate_itf(itf_val, visual):
    v = itf_val.strip()
    if not v.isdigit():
        raise ValueError(f"ITF 条码值含非数字字符：'{itf_val}'")
    if len(v) % 2 != 0:
        raise ValueError(f"ITF 条码值为奇数位（{len(v)} 位）：'{v}'，会被静默补 0 导致条码与号码不符（Tesco ITF-14 应为 14 位）")
    if visual and _strip_spaces(visual) != v:
        raise ValueError(f"ITF 目视号码 '{visual}' 去空格后与条码值 '{v}' 不一致")


def validate_code128(bd_val, visual):
    v = bd_val.strip()
    if not v:
        raise ValueError("BD 条码值为空")
    if visual and _strip_spaces(visual) != v:
        raise ValueError(f"BD 目视号码 '{visual}' 去空格后与条码值 '{v}' 不一致")


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
    subprocess.run(["fc-cache", "-f"], capture_output=True)


_install_fonts()


# ==========================================
# 自动命名 & 自动推导
# ==========================================
def build_name(kind, row):
    pin = normalize_spaces(safe_str(row.get("产品品名")))
    eng = normalize_spaces(safe_str(row.get("产品英文名")))
    po = safe_str(row.get("PO号"))
    qty = safe_str(row.get("总外箱数"))
    if kind == "outer":
        return f"{po}{pin}{qty}个外箱"
    if kind == "inner":
        return f"内盒-{pin}"
    if kind == "itf_uk":
        return f"ITF-UK-{pin}-{eng}"
    if kind == "itf_ce":
        return f"ITF-CE-{pin}-{eng}"
    if kind == "bd":
        return f"BD-{pin}"
    return ""


def resolved_name(kind, row):
    """命名优先读数据源里的命名列（Excel 公式算好的值）；为空时回退到内置规则生成。"""
    col_val = safe_str(row.get(NAME_COL.get(kind, "")))
    return sanitize_name(col_val) if col_val else sanitize_name(build_name(kind, row))


def derive_outer_size(row):
    existing = safe_str(row.get("外箱尺寸"))
    if existing:
        return existing
    l, w, h = safe_str(row.get("外箱长")), safe_str(row.get("外箱宽")), safe_str(row.get("外箱高"))
    return f"{l}*{w}*{h}cm" if (l and w and h) else ""


def int_str(val):
    """整数字段（总外箱数/外箱装量/内盒装量）：去掉小数点，只显示整数。"""
    s = safe_str(val)
    if not s:
        return ""
    try:
        return str(int(round(float(s))))
    except ValueError:
        return s


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
    """用 pdftotext -bbox 拿每个词的坐标，用于校验品名是否被换行拆开。"""
    out = subprocess.run(["pdftotext", "-bbox", pdf_path, "-"], capture_output=True, text=True, timeout=60)
    if out.returncode != 0 or not out.stdout:
        return []
    root = etree.fromstring(out.stdout.encode("utf-8"))
    ns = {"h": "http://www.w3.org/1999/xhtml"}
    words = []
    for page_i, page in enumerate(root.findall(".//h:page", ns)):
        for w in page.findall(".//h:word", ns):
            try:
                words.append({
                    "page": page_i, "text": w.text or "",
                    "xMin": float(w.get("xMin")), "yMin": float(w.get("yMin")),
                })
            except (TypeError, ValueError):
                continue
    return words


def name_wrapped_in_pdf(pdf_path, name_text, y_tol=1.0):
    """校验品名在生成的 PDF 里是否被换行拆成了两行。
    做法：按 y 坐标把词聚成"行"，若某一行的文本恰好是品名的真前缀/后缀（而非完整品名），
    说明品名被拆到了相邻行 —— 判定为换行。"""
    target = normalize_spaces(name_text)
    if not target:
        return False
    try:
        words = _pdf_word_boxes(pdf_path)
    except Exception:
        return False
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


def resolve_bl(row):
    """品牌 Logo 判定：优先读 BL 列（Excel 公式值）；为空则按归一化英文名判断，
    兼容数据里的不间断空格（原公式用普通空格 SEARCH 会漏判，导致 Logo 缺失）。"""
    bl = safe_str(row.get("BL"))
    if bl:
        return bl
    low = normalize_spaces(safe_str(row.get("产品英文名"))).lower()
    if "go cook" in low or "tesco" in low:
        return "Tesco"
    if "f&f home" in low:
        return "FF"
    return ""


TYPE_LABELS = {"外箱": "外箱", "内盒": "内盒", "ITF": "ITF 标", "BD": "BD 标"}
ALL_TYPES = ["外箱", "内盒", "ITF", "BD"]

# 命名列 <-> 生成类型
NAME_COL = {"outer": "外箱文件命名列", "inner": "内盒文件命名列",
            "itf_uk": "ITF_UK命名", "itf_ce": "ITF_CE命名", "bd": "BD命名"}

# 节日logo → 胶带颜色 查找表（原 Sheet2）
TAPE_LOOKUP = [
    ("Halloween", "橙色"), ("Christmas", "请与客户确认"), ("Valentines", "红色"),
    ("Mother’s Day", "红色"), ("Father’s Day", "红色"), ("Easter", "红色"), ("Gardening", "绿色"),
]

# ---- 合并数据源三色结构 ----
# kind: data=绿色手填 / formula=橙蓝公式（Excel自动算） / select=黄色人工选择
# f: 公式模板（{r}=行号），忠实还原原表并修好 #REF；命名用用户确认的新规则
_CDU = "Slit tape at base, lift off lid display full tray on shelf"
MERGED_SCHEMA = [
    # 绿色：手填数据（顺序固定 A–S，与原表一致）
    {"n": "产品品名", "kind": "data", "text": True},
    {"n": "产品英文名", "kind": "data", "text": True},
    {"n": "PO号", "kind": "data", "text": True},
    {"n": "总外箱数", "kind": "data", "intfmt": True},
    {"n": "TPNB", "kind": "data", "text": True},
    {"n": "TPND", "kind": "data", "text": True},
    {"n": "ITF", "kind": "data", "text": True},
    {"n": "category", "kind": "data", "text": True},
    {"n": "EAN", "kind": "data", "text": True},
    {"n": "外箱装量", "kind": "data", "intfmt": True},
    {"n": "内盒装量", "kind": "data", "intfmt": True},
    {"n": "CEORMSNO", "kind": "data", "text": True},
    {"n": "SKU", "kind": "data", "text": True},
    {"n": "VSN", "kind": "data", "text": True},
    {"n": "毛重", "kind": "data"},
    {"n": "净重", "kind": "data"},
    {"n": "外箱长", "kind": "data"},
    {"n": "外箱宽", "kind": "data"},
    {"n": "外箱高", "kind": "data"},
    # 橙/蓝：公式列（列引用见上方注释的字母布局）
    {"n": "订单国家", "kind": "formula",
     "f": '=IF(LEFT(C{r},3)="380","UK",IF(LEFT(C{r},3)="999","CE",IF(LEFT(C{r},3)="520","Ireland","")))'},
    {"n": "CON不加粗", "kind": "formula", "f": "=LEFT(L{r},8)"},
    {"n": "CON加粗", "kind": "formula", "f": "=RIGHT(L{r},5)"},
    {"n": "CON", "kind": "formula", "f": "=L{r}"},
    {"n": "外箱BD号", "kind": "formula", "f": '=IFERROR(IF(J{r}/K{r}=1," ","B/D "&J{r}/K{r}),"")'},
    {"n": "外箱尺寸", "kind": "formula", "f": '=Q{r}&"*"&R{r}&"*"&S{r}&"cm"'},
    {"n": "双人抬标识", "kind": "formula",
     "f": '=IF(OR(AND(LEFT(C{r},3)="380",O{r}>23),AND(LEFT(C{r},3)="999",O{r}>14)),"双人抬","")'},
    {"n": "BL", "kind": "formula",
     "f": '=IFERROR(IF(OR(ISNUMBER(SEARCH("Go Cook",SUBSTITUTE(B{r},UNICHAR(160)," "))),ISNUMBER(SEARCH("Tesco",SUBSTITUTE(B{r},UNICHAR(160)," ")))),"Tesco",IF(ISNUMBER(SEARCH("F&F Home",SUBSTITUTE(B{r},UNICHAR(160)," "))),"FF","")),"")'},
    {"n": "ITF加空格", "kind": "formula", "f": '=LEFT(G{r},3)&" "&MID(G{r},4,5)&" "&RIGHT(G{r},6)'},
    {"n": "BD条码号", "kind": "formula", "f": '=IF(J{r}=K{r},"",IFERROR("02"&G{r}&"37"&TEXT(J{r}/K{r},"00"),""))'},
    {"n": "BD条码号加空格", "kind": "formula",
     "f": '=IF(AC{r}="","",LEFT(AC{r},2)&" "&MID(AC{r},3,14)&" "&MID(AC{r},17,2)&" "&RIGHT(AC{r},2))'},
    # 黄色人工 + 其后公式
    {"n": "节日logo", "kind": "select", "dd": [x[0] for x in TAPE_LOOKUP]},
    {"n": "胶带颜色", "kind": "formula",
     "f": '=IF(AE{r}=""," ",_xlfn.XLOOKUP(AE{r},选项字典!$A$2:$A$8,选项字典!$B$2:$B$8,"")&"胶带用于上口封箱")'},
    {"n": "是否有CDU", "kind": "select", "dd": "YN"},
    {"n": "有CDU加印文字", "kind": "formula", "f": f'=IF(AG{{r}}="Y","{_CDU}","")'},
    {"n": "内盒是否直接接触CDU", "kind": "select", "dd": "YN"},
    {"n": "有CDU加印文字内盒版", "kind": "formula", "f": f'=IF(AI{{r}}="Y","{_CDU}","")'},
    {"n": "是否要做外箱", "kind": "select", "dd": "YN"},
    {"n": "外箱文件命名列", "kind": "formula", "f": '=IF(AK{r}="Y",C{r}&A{r}&D{r}&"个外箱","")'},
    {"n": "是否要做内盒", "kind": "select", "dd": "YN"},
    {"n": "内盒文件命名列", "kind": "formula", "f": '=IF(AND(AM{r}="Y",K{r}<>J{r}),"内盒-"&A{r},"")'},
    {"n": "是否需要ITF-UK", "kind": "select", "dd": "YN"},
    {"n": "ITF_UK命名", "kind": "formula", "f": '=IF(AO{r}="Y","ITF-UK-"&A{r}&"-"&B{r},"")'},
    {"n": "是否需要ITF-CE", "kind": "select", "dd": "YN"},
    {"n": "ITF_CE命名", "kind": "formula", "f": '=IF(AQ{r}="Y","ITF-CE-"&A{r}&"-"&B{r},"")'},
    {"n": "是否要做BD", "kind": "select", "dd": "YN"},
    {"n": "BD命名", "kind": "formula", "f": '=IF(AND(AS{r}="Y",AC{r}<>""),"BD-"&A{r},"")'},
]


@st.cache_data
def build_blank_template():
    from openpyxl import Workbook
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.styles import Font, PatternFill, Alignment

    wb = Workbook()
    ws = wb.active
    ws.title = "数据源"

    fills = {
        "data": PatternFill("solid", fgColor="C6E0B4"),      # 绿：手填
        "formula": PatternFill("solid", fgColor="FCE4D6"),   # 橙：公式（自动）
        "select": PatternFill("solid", fgColor="FFFF00"),    # 黄：人工选择
    }
    # 示例两行的“手填/选择”值（公式列留空，交给 Excel 计算）
    ex1 = {"产品品名": "相思木黑陶瓷碗套装", "产品英文名": "F&F Home Acacia Wooden Tray with 3 bowls black",
           "PO号": "380-68804", "总外箱数": 845, "TPNB": "099065895", "TPND": "000384156",
           "ITF": "05063638227053", "category": "Cook & Dining", "EAN": "5063638226940",
           "外箱装量": 6, "内盒装量": 6, "VSN": "ST1238-TESCO", "毛重": 41, "净重": 27,
           "外箱长": 41, "外箱宽": 27, "外箱高": 25, "是否要做外箱": "Y", "是否要做内盒": "N",
           "是否需要ITF-UK": "Y", "是否需要ITF-CE": "N", "是否要做BD": "N"}
    ex2 = {"产品品名": "不锈钢刮皮刀", "产品英文名": "Go Cook Peeler", "PO号": "999-88153",
           "总外箱数": 16, "ITF": "05063446647500", "category": "Cook & Dining",
           "CEORMSNO": "2005101001984", "外箱装量": 96, "内盒装量": 6, "VSN": "ST1864-TESCO",
           "毛重": 6.7, "净重": 6.1, "外箱长": 45, "外箱宽": 17.5, "外箱高": 38,
           "是否要做外箱": "Y", "是否要做内盒": "N", "是否需要ITF-UK": "N",
           "是否需要ITF-CE": "Y", "是否要做BD": "Y"}

    for c, spec in enumerate(MERGED_SCHEMA, start=1):
        letter = get_column_letter(c)
        h = ws.cell(row=1, column=c, value=spec["n"])
        h.fill = fills[spec["kind"]]
        h.font = Font(bold=True, color="000000")
        h.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[letter].width = max(9, min(22, len(spec["n"]) + 4))
        for ri, ex in enumerate((ex1, ex2), start=2):
            cell = ws.cell(row=ri, column=c)
            if spec["kind"] == "formula":
                cell.value = spec["f"].format(r=ri)
            else:
                cell.value = ex.get(spec["n"], "")
                if spec.get("text"):
                    cell.number_format = "@"   # 代码列设文本，保护前导0
                elif spec.get("intfmt"):
                    cell.number_format = "0"   # 总外箱数/装量：整数显示，无小数点

    # 数据验证（黄色列下拉）
    for c, spec in enumerate(MERGED_SCHEMA, start=1):
        if spec["kind"] != "select":
            continue
        letter = get_column_letter(c)
        if spec["dd"] == "YN":
            dv = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
        else:
            dv = DataValidation(type="list", formula1="选项字典!$A$2:$A$8", allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"{letter}2:{letter}500")

    # 选项字典表（节日logo→胶带颜色，供 XLOOKUP）
    ws2 = wb.create_sheet("选项字典")
    ws2["A1"] = "节日logo"; ws2["B1"] = "胶带颜色"
    ws2["A1"].font = ws2["B1"].font = Font(bold=True)
    for ri, (a, b) in enumerate(TAPE_LOOKUP, start=2):
        ws2.cell(row=ri, column=1, value=a)
        ws2.cell(row=ri, column=2, value=b)

    # 说明表
    ws3 = wb.create_sheet("填写说明")
    notes = [
        "颜色含义：绿色=你手填的数据；橙色=公式列（Excel 自动算，请勿手改）；黄色=人工选择（下拉 Y/N 或节日）。",
        "务必用 Excel 打开填写并保存，让公式先算好，再上传网站（网站读取的是公式算出的值）。",
        "带前导 0 的号码（如 ITF、CEORMSNO）已设为文本格式，不会丢 0。",
        "每行开关：是否要做外箱/内盒、是否需要ITF-UK/ITF-CE、是否要做BD，填 Y 才生成对应标签。",
        "命名自动生成：外箱=PO+品名+总箱数+个外箱；内盒=内盒-品名；ITF=ITF-UK/CE-品名-英文名；BD=BD-品名。",
        "BD条码号与外箱BD号由装量自动计算（外箱装量=内盒装量时不产生BD）。",
        "内盒装量=外箱装量、或内盒装量为空/无数字时，视为无内盒（产品直接入外箱），自动不生成内盒标（即使勾了是否要做内盒=Y）。",
        "内盒装量无数字（无内盒）时，ITF 标上的 Case Size 直接取外箱装量。",
    ]
    for ri, t in enumerate(notes, start=1):
        ws3.cell(row=ri, column=1, value=("• " + t))
    ws3.column_dimensions["A"].width = 100

    ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ==========================================
# 后台公式引擎（C）：从原始数据算出所有派生列，不再依赖 Excel 公式
# ==========================================
def derive_country(po):
    return {"380": "UK", "999": "CE", "520": "Ireland"}.get(safe_str(po)[:3], "")


def derive_heavy(po, gross_weight):
    p = safe_str(po)[:3]
    try:
        gw = float(safe_str(gross_weight))
    except ValueError:
        return ""
    if (p == "380" and gw > 23) or (p == "999" and gw > 14):
        return "双人抬"
    return ""


def space_itf(itf):
    v = _strip_spaces(safe_str(itf))
    return f"{v[:3]} {v[3:8]} {v[8:]}" if len(v) >= 9 else v


def _ratio(outer_qty, inner_qty):
    try:
        oq, iq = float(safe_str(outer_qty)), float(safe_str(inner_qty))
    except ValueError:
        return None
    if iq == 0:
        return None
    return oq / iq


def derive_bd_num(outer_qty, inner_qty):
    r = _ratio(outer_qty, inner_qty)
    if r is None or r == 1:
        return " "
    return f"B/D {int(r) if r == int(r) else r}"


def derive_bd_barcode(itf, outer_qty, inner_qty):
    v = _strip_spaces(safe_str(itf))
    r = _ratio(outer_qty, inner_qty)
    if not v or r is None or r == 1:
        return ""
    return f"02{v}37{int(round(r)):02d}"


def space_bd(bd):
    v = _strip_spaces(safe_str(bd))
    return f"{v[:2]} {v[2:16]} {v[16:18]} {v[18:]}" if len(v) >= 20 else v


def derive_tape(festival):
    f = safe_str(festival)
    if not f:
        return " "
    for a, b in TAPE_LOOKUP:
        if a == f:
            return f"{b}胶带用于上口封箱"
    return " "


def enrich_row(raw, cdu_map=None):
    """把一行原始数据补全为含所有派生列的完整行（仅填空缺项，已有值不覆盖）。
    这样'网页粘贴原始数据'与'上传带公式的Excel'两条路都能生成。"""
    row = dict(raw)

    def cur(k):
        return safe_str(row.get(k))

    itf = cur("ITF")
    po = cur("PO号")
    if not cur("订单国家"):
        row["订单国家"] = derive_country(po)
    if not cur("CON不加粗") and not cur("CON加粗"):
        row["CON不加粗"], row["CON加粗"] = split_orms_bold(get_orms(row))
    if not cur("外箱尺寸"):
        row["外箱尺寸"] = derive_outer_size(row)
    if not cur("双人抬标识"):
        row["双人抬标识"] = derive_heavy(po, cur("毛重"))
    if not cur("BL"):
        row["BL"] = resolve_bl(row)
    if not cur("ITF加空格"):
        row["ITF加空格"] = space_itf(itf)
    if not cur("外箱BD号"):
        row["外箱BD号"] = derive_bd_num(cur("外箱装量"), cur("内盒装量"))
    if not cur("BD条码号"):
        row["BD条码号"] = derive_bd_barcode(itf, cur("外箱装量"), cur("内盒装量"))
    if not cur("BD条码号加空格"):
        row["BD条码号加空格"] = space_bd(cur("BD条码号"))
    if not cur("胶带颜色"):
        row["胶带颜色"] = derive_tape(cur("节日logo"))
    # CDU（D）：按 ITF 从清单自动匹配外箱/内盒是否接触 CDU
    if cdu_map:
        key = _strip_spaces(itf)
        if key in cdu_map:
            oflag, iflag = cdu_map[key]
            if not cur("是否有CDU") and oflag:
                row["是否有CDU"] = oflag
            if not cur("内盒是否直接接触CDU") and iflag:
                row["内盒是否直接接触CDU"] = iflag
    if not cur("有CDU加印文字") and is_yes(row.get("是否有CDU")):
        row["有CDU加印文字"] = _CDU
    if not cur("有CDU加印文字内盒版") and is_yes(row.get("内盒是否直接接触CDU")):
        row["有CDU加印文字内盒版"] = _CDU
    return row


# ==========================================
# CDU 清单（D）：仓库内置 CSV，可在前端查看/编辑/导入/导出，按 ITF 匹配
# ==========================================
CDU_CSV = get_resource_path("cdu_list.csv")
CDU_COLS = ["ITF", "外箱接触CDU", "内盒接触CDU"]


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


def cdu_df_to_map(df):
    m = {}
    if df is None or len(df) == 0:
        return m
    for _, r in df.iterrows():
        key = _strip_spaces(safe_str(r.get("ITF")))
        if key:
            m[key] = (safe_str(r.get("外箱接触CDU")), safe_str(r.get("内盒接触CDU")))
    return m


# ==========================================
# 生成逻辑（每行开关 gate + 自动命名 + 自动推导）
# ==========================================
def _confirm_pdf(output_dir, work_dir, name_base):
    pdf_path = os.path.join(output_dir, f"{name_base}.pdf")
    docx_path = os.path.join(work_dir, f"{name_base}.docx")
    if os.path.exists(pdf_path):
        if os.path.exists(docx_path):
            os.remove(docx_path)
        return True
    raise RuntimeError(f"LibreOffice 返回成功但未找到输出 PDF：{name_base}.pdf")


def _register(used, name):
    """登记输出名。返回 True=首次（应生成）；False=已存在（同款产品已做过，跳过）。"""
    if name in used:
        return False
    used.add(name)
    return True


def gen_outer(row, work_dir, out_dir, used, profile):
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

    ctx["产品英文名"] = esc(safe_str(row.get("产品英文名")))
    ctx["外箱BD号"] = safe_str(row.get("外箱 BD 号", row.get("外箱BD号")))
    ctx["有CDU加印文字"] = safe_str(row.get("有CDU-加印文字", row.get("有CDU加印文字")))
    ctx.update({
        "PO号": po_str, "category": esc(safe_str(row.get("category"))),
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
    if not _register(used, name):
        return [{"命名": name, "状态": "跳过", "说明": "重复，已生成一次"}]
    docx = os.path.join(work_dir, f"{name}.docx")
    tpl.render(ctx)
    tpl.save(docx)
    ensure_name_no_wrap(docx, safe_str(row.get("产品英文名")), out_dir, profile, max_mm=34.0, max_pt=6.5, min_pt=2.0)
    _confirm_pdf(out_dir, work_dir, name)
    return [{"命名": name, "状态": "成功", "说明": "UK" if is_uk else "CE"}]


def gen_inner(row, work_dir, out_dir, used, profile):
    tpl = DocxTemplate(get_resource_path("模板_InnerCarton.docx"))
    ctx = {
        # 品名在文本框内（docxtpl RichText 在文本框里不生效），仍用纯字符串；
        # 换行问题改在模板里把品名字号调小解决。
        "产品英文名": esc(safe_str(row.get("产品英文名"))),
        "category": esc(safe_str(row.get("category"))),
        "EAN": safe_str(row.get("EAN")),
        "有CDU加印文字内盒版": safe_str(row.get("有CDU加印文字内盒版")),
    }
    bl = resolve_bl(row)
    bl_path = get_resource_path(f"2. Brand Logo/{bl}.png")
    # 内盒黑条较窄，Logo 用 2.2mm 以免顶部溢出被交界处遮住
    ctx["BL"] = InlineImage(tpl, bl_path, height=Mm(2.2)) if bl and os.path.exists(bl_path) else ""
    name = resolved_name("inner", row)
    if not _register(used, name):
        return [{"命名": name, "状态": "跳过", "说明": "同款产品已生成一次"}]
    docx = os.path.join(work_dir, f"{name}.docx")
    tpl.render(ctx)
    tpl.save(docx)
    ensure_name_no_wrap(docx, safe_str(row.get("产品英文名")), out_dir, profile, max_mm=34.0, max_pt=6.0, min_pt=2.0)
    _confirm_pdf(out_dir, work_dir, name)
    return [{"命名": name, "状态": "成功", "说明": ""}]


def gen_itf(row, work_dir, out_dir, used, profile, want_uk, want_ce):
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
            if not _register(used, name):
                results.append({"命名": name, "状态": "跳过", "说明": kind.upper() + " 同款已生成一次"})
                return
            tpl = DocxTemplate(tpl_path)
            rt = RichText(); rt.add(eng_raw, bold=True, size=int(round(fit_pt * 2)), font="微软雅黑")
            ctx = {"TPNB": tpnb, "TPND": tpnd, "CON不加粗": con_n, "CON加粗": con_b,
                   "产品英文名": rt, "内盒装量": inner_qty, "ITF加空格": itf_spaced,
                   "Barcode_Img": InlineImage(tpl, saved, width=Mm(95), height=Mm(28))}
            docx = os.path.join(work_dir, f"{name}.docx")
            tpl.render(ctx); tpl.save(docx)
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


def gen_bd(row, work_dir, out_dir, used, profile):
    bd_val = safe_str(row.get("BD条码号"))
    name = resolved_name("bd", row)
    if not bd_val:
        # 无内盒(装量相等)时 BD条码号 公式为空 → 该产品无 BD，跳过（非错误）
        return [{"命名": name, "状态": "跳过", "说明": "无内盒或BD条码号为空，无需BD"}]
    bd_spaced = safe_str(row.get("BD条码号加空格", bd_val))
    validate_code128(bd_val, bd_spaced)
    if not _register(used, name):
        return [{"命名": name, "状态": "跳过", "说明": "同款产品已生成一次"}]

    Code128 = barcode.get_barcode_class("code128")
    img_tmp = os.path.join(work_dir, f"barcode_bd_{uuid.uuid4().hex}")
    saved = Code128(bd_val, writer=ImageWriter()).save(img_tmp, options={"write_text": False, "module_height": 22.0})
    crop_vertical_margin(saved)
    tpl = DocxTemplate(get_resource_path("模板_BD.docx"))
    ctx = {"BD条码号加空格": bd_spaced, "Barcode_Img": InlineImage(tpl, saved, width=Mm(140), height=Mm(70))}
    try:
        tpl.render(ctx)
        docx = os.path.join(work_dir, f"{name}.docx")
        tpl.save(docx)
        convert_docx_to_pdf(docx, out_dir, profile)
        _confirm_pdf(out_dir, work_dir, name)
        return [{"命名": name, "状态": "成功", "说明": ""}]
    finally:
        if os.path.exists(saved):
            os.remove(saved)


# ==========================================
# 生成前“体检”：不出 PDF，只做校验
# ==========================================
# 完整性必填规则（按订单前缀）：
_DATA_FIELDS = ["产品品名", "产品英文名", "PO号", "总外箱数", "TPNB", "TPND", "ITF",
                "category", "EAN", "外箱装量", "内盒装量", "CEORMSNO", "SKU", "VSN",
                "毛重", "净重", "外箱长", "外箱宽", "外箱高"]
_REQ_UK = [c for c in _DATA_FIELDS if c not in ("CEORMSNO", "SKU")]   # 380/520：除 CEORMSNO/SKU 外都必填
_REQ_CE = [c for c in _DATA_FIELDS if c not in ("TPNB", "TPND")]      # 999：除 TPNB/TPND 外都必填


def _missing_or_zero(v):
    s = safe_str(v)
    if not s:
        return True
    try:
        return float(s) == 0
    except ValueError:
        return False


# 固定位数校验（数字位数）：TPNB/TPND 9、ITF 14、EAN 13、CEORMSNO 13、SKU 9
_DIGIT_LEN = {"TPNB": 9, "TPND": 9, "ITF": 14, "EAN": 13, "CEORMSNO": 13, "SKU": 9}


def _iss(idx, po, col, msg):
    return {"idx": idx, "行": idx + 2, "PO号": po or "（PO号为空）", "列": col, "问题": msg}


def preflight(df, selected, cdu_map=None):
    """返回结构化问题列表：[{idx, 行, PO号, 列, 问题}]。"""
    issues = []
    for idx, raw in df.iterrows():
        row = enrich_row(raw, cdu_map)
        country = safe_str(row.get("订单国家"))
        has_inner = not no_inner(row)
        po = safe_str(row.get("PO号"))
        prefix = po[:3]

        # ① 完整性 + TPNB/TPND 的 "/" 规则（与所选类型无关，始终校验）
        if not po:
            issues.append(_iss(idx, "", "PO号", "PO号为空"))
        elif prefix in ("380", "520"):
            for c in _REQ_UK:
                if _missing_or_zero(row.get(c)):
                    issues.append(_iss(idx, po, c, "不能为空或0（380/520 必填）"))
            for c in ("TPNB", "TPND"):
                if "/" in safe_str(row.get(c)):
                    issues.append(_iss(idx, po, c, "含 '/'（380/520 不允许）"))
        elif prefix == "999":
            for c in _REQ_CE:
                if _missing_or_zero(row.get(c)):
                    issues.append(_iss(idx, po, c, "不能为空或0（999 必填）"))
            # 999：TPNB/TPND 允许留空或含 '/'
        else:
            issues.append(_iss(idx, po, "PO号", f"前缀 '{prefix}' 非 380/520/999"))

        # ② 固定位数校验（有值且非 '/' 才校验；空缺由必填规则负责）
        for c, n in _DIGIT_LEN.items():
            # 380/520（UK）订单不使用 CEORMSNO/SKU，跳过其位数校验
            if prefix in ("380", "520") and c in ("CEORMSNO", "SKU"):
                continue
            v = safe_str(row.get(c))
            if not v or (c in ("TPNB", "TPND") and "/" in v):
                continue
            if not (v.isdigit() and len(v) == n):
                issues.append(_iss(idx, po, c, f"必须是 {n} 位数字（当前 {len(v)} 位）"))

        # ③ 条码正确性
        if "ITF" in selected and (wants(row, "是否需要ITF-UK", country in ("UK", "Ireland")) or wants(row, "是否需要ITF-CE", country == "CE")):
            try:
                validate_itf(safe_str(row.get("ITF")), safe_str(row.get("ITF加空格", safe_str(row.get("ITF")))))
            except ValueError as e:
                issues.append(_iss(idx, po, "ITF", str(e)))
        if "BD" in selected and has_inner and wants(row, "是否要做BD", True):
            bdv = safe_str(row.get("BD条码号"))
            if bdv:
                try:
                    validate_code128(bdv, safe_str(row.get("BD条码号加空格", bdv)))
                except ValueError as e:
                    issues.append(_iss(idx, po, "BD条码号", str(e)))
    return issues


def run_all(df, selected, cdu_map=None):
    with tempfile.TemporaryDirectory() as work:
        base = os.path.join(work, "out")
        profile = os.path.join(work, "lo")
        os.makedirs(profile, exist_ok=True)
        subdirs = {t: os.path.join(base, t) for t in selected}
        for d in subdirs.values():
            os.makedirs(d, exist_ok=True)
        used = {t: set() for t in selected}
        results = []
        prog = st.progress(0)
        status = st.empty()
        total = len(df)

        for idx, raw in df.iterrows():
            rn = idx + 2
            row = enrich_row(raw, cdu_map)  # 后台算全所有派生列
            country = safe_str(row.get("订单国家"))
            has_inner = not no_inner(row)
            w_uk = wants(row, "是否需要ITF-UK", country in ("UK", "Ireland"))
            w_ce = wants(row, "是否需要ITF-CE", country == "CE")
            jobs = []
            if "外箱" in selected and wants(row, "是否要做外箱", True):
                jobs.append(("外箱", lambda r: gen_outer(r, work, subdirs["外箱"], used["外箱"], profile)))
            if "内盒" in selected and has_inner and wants(row, "是否要做内盒", True):
                jobs.append(("内盒", lambda r: gen_inner(r, work, subdirs["内盒"], used["内盒"], profile)))
            if "ITF" in selected and (w_uk or w_ce):
                jobs.append(("ITF", lambda r, u=w_uk, c=w_ce: gen_itf(r, work, subdirs["ITF"], used["ITF"], profile, u, c)))
            if "BD" in selected and has_inner and wants(row, "是否要做BD", True):
                jobs.append(("BD", lambda r: gen_bd(r, work, subdirs["BD"], used["BD"], profile)))
            for tname, fn in jobs:
                try:
                    for res in fn(row):
                        res.update({"行": rn, "类型": tname})
                        results.append(res)
                except Exception as e:
                    results.append({"行": rn, "类型": tname, "命名": "", "状态": "失败", "说明": str(e)})
            prog.progress((idx + 1) / total if total else 1.0)
            status.text(f"正在生成中... ({idx + 1}/{total})")
        prog.empty(); status.empty()

        pdfs = [os.path.join(r, f) for r, _d, fs in os.walk(base) for f in fs if f.endswith(".pdf")]
        actual = len(pdfs)
        preview_png = None
        if pdfs:
            try:
                subprocess.run(["pdftoppm", "-png", "-r", "80", "-singlefile", sorted(pdfs)[0],
                                os.path.join(work, "preview")], capture_output=True, timeout=30)
                pv = os.path.join(work, "preview.png")
                if os.path.exists(pv):
                    with open(pv, "rb") as fp:
                        preview_png = fp.read()
            except Exception:
                preview_png = None
            zp = create_zip(base)
            with open(zp, "rb") as fp:
                st.session_state.download_data = fp.read()
            st.session_state.download_name = f"Tesco标签_{actual}个.zip"
            os.remove(zp)
        return results, actual, preview_png


# ==========================================
# UI（v2：网页粘贴原始数据 + 公式后台化 + CDU 按 ITF 自动匹配）
# ==========================================
INPUT_DATA_COLS = ["产品品名", "产品英文名", "PO号", "总外箱数", "TPNB", "TPND", "ITF",
                   "category", "EAN", "外箱装量", "内盒装量", "CEORMSNO", "SKU", "VSN",
                   "毛重", "净重", "外箱长", "外箱宽", "外箱高"]
INPUT_SELECT_COLS = ["节日logo", "是否要做外箱", "是否要做内盒", "是否需要ITF-UK", "是否需要ITF-CE", "是否要做BD"]
INPUT_COLUMNS = INPUT_DATA_COLS + INPUT_SELECT_COLS
FESTIVALS = [x[0] for x in TAPE_LOOKUP]


def _empty_input_df(n=15):
    return pd.DataFrame("", index=range(n), columns=INPUT_COLUMNS)


def _clean_input_df(df):
    df = df.fillna("").astype(str)
    df.columns = [str(c).strip() for c in df.columns]
    key = [c for c in ["产品品名", "产品英文名", "PO号", "ITF"] if c in df.columns]
    if key:
        mask = df[key].apply(lambda r: any(str(v).strip() for v in r), axis=1)
        df = df[mask]
    return df.reset_index(drop=True)


st.title("📦 Tesco 标贴与纸箱生成系统")
st.caption("Made By Sherry ｜ 仅供 Suncha 内部使用 ｜ v2 · 网页录入 + 后台公式")

# ---- CDU 清单（D）：按 ITF 自动匹配 ----
if "cdu_df" not in st.session_state:
    st.session_state.cdu_df = load_cdu_df_bundled()
with st.sidebar:
    st.header("🏷️ CDU 清单")
    st.caption("按 ITF 自动匹配：哪些产品外箱/内盒直接接触 CDU。填 Y 的会在标签上加印 CDU 文案。")
    cdu_up = st.file_uploader("导入 CDU 清单 CSV", type=["csv"], key="cdu_up")
    if cdu_up is not None:
        try:
            st.session_state.cdu_df = pd.read_csv(cdu_up, dtype=str).fillna("")
            st.success("CDU 清单已导入。")
        except Exception as e:
            st.error(f"导入失败：{e}")
    cdu_edited = st.data_editor(
        st.session_state.cdu_df, num_rows="dynamic", use_container_width=True, key="cdu_editor",
        column_config={
            "ITF": st.column_config.TextColumn("ITF", help="产品 ITF 条码值"),
            "外箱接触CDU": st.column_config.SelectboxColumn("外箱接触CDU", options=["", "Y", "N"]),
            "内盒接触CDU": st.column_config.SelectboxColumn("内盒接触CDU", options=["", "Y", "N"]),
        })
    st.download_button("⬇️ 导出当前 CDU 清单", cdu_edited.to_csv(index=False).encode("utf-8-sig"),
                       "cdu_list.csv", "text/csv", use_container_width=True)
    st.caption("云端重启会还原为仓库内置版；新增/修改后请导出 CSV 重新提交仓库（或下次导入）。")
    st.divider()
    st.download_button("⬇️ 下载空白 Excel 模板（备用）", data=build_blank_template(),
                       file_name="Tesco数据源模板.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)
cdu_map = cdu_df_to_map(cdu_edited)

# ---- 第一步：录入数据 ----
st.markdown("**第一步：录入数据**（只填原始数据，公式/命名/条码等由系统自动算）")
mode = st.radio("录入方式", ["① 粘贴 / 编辑数据（推荐）", "② 上传 Excel"], horizontal=True, label_visibility="collapsed")

df = None
if mode.startswith("①"):
    st.caption("可直接从 Excel 复制一整片数据粘贴进来，或逐格填写。右侧几列是人工选择——**留空即按规则自动**"
               "（外箱默认做、有内盒才做内盒/BD、ITF 按订单国家）；只有明确不做才填 N。")
    if "input_df" not in st.session_state:
        st.session_state.input_df = _empty_input_df()
    yn = ["", "Y", "N"]
    colcfg = {c: st.column_config.SelectboxColumn(c, options=yn, width="small")
              for c in ["是否要做外箱", "是否要做内盒", "是否需要ITF-UK", "是否需要ITF-CE", "是否要做BD"]}
    colcfg["节日logo"] = st.column_config.SelectboxColumn("节日logo", options=[""] + FESTIVALS)
    edited = st.data_editor(st.session_state.input_df, num_rows="dynamic",
                            use_container_width=True, key="input_editor", column_config=colcfg)
    df = _clean_input_df(edited)
    if len(df):
        st.success(f"当前有效数据 {len(df)} 行。")
else:
    uploaded = st.file_uploader("上传数据源 Excel（原始数据即可，无需公式）", type=["xlsx", "xls"])
    if uploaded is not None:
        df = _clean_input_df(pd.read_excel(uploaded, dtype=str))
        st.success(f"读取成功，共 {len(df)} 行。")

# ---- 第二步：自动体检 + 选类型 + 生成 ----
if df is not None and len(df):
    # 录入后自动体检（每次编辑实时刷新）；按 PO 合并显示 + 问题格标红
    issues = preflight(df, ALL_TYPES, cdu_map)
    if issues:
        from collections import OrderedDict
        grouped = OrderedDict()
        for it in issues:
            grouped.setdefault(it["PO号"], []).append(it)
        st.error(f"⚠️ 数据体检发现 {len(issues)} 处问题，涉及 {len(grouped)} 个 PO（按 PO 对照上方表格修改）：")
        gp_rows = [{"PO号": po, "问题数": len(its),
                    "问题": "；".join((f"{i['列']}：{i['问题']}" if i["列"] else i["问题"]) for i in its)}
                   for po, its in grouped.items()]
        st.dataframe(pd.DataFrame(gp_rows), use_container_width=True, hide_index=True)
    else:
        st.success("✅ 数据体检通过，未发现问题。")

    st.markdown("**第二步：勾选本次要生成的类型**")
    cols = st.columns(4)
    selected = [t for i, t in enumerate(ALL_TYPES)
                if cols[i].checkbox(TYPE_LABELS[t], value=True, key=f"chk_{t}")]

    if st.button("🚀 开始生成", type="primary", use_container_width=True):
        st.session_state.download_data = None
        if not selected:
            st.warning("请至少勾选一种类型。")
        else:
            results, actual, preview_png = run_all(df, selected, cdu_map)
            if actual > 0:
                st.success(f"🎉 完成！实际生成 {actual} 个 PDF（按类型分文件夹打包）。")
            else:
                st.warning("没有生成任何文件，请检查数据与开关列。")
            if results:
                rdf = pd.DataFrame(results)[["行", "类型", "命名", "状态", "说明"]]
                fails = rdf[rdf["状态"] == "失败"]
                if not fails.empty:
                    st.error(f"其中 {len(fails)} 项失败：")
                    st.dataframe(fails, use_container_width=True, hide_index=True)
                with st.expander(f"查看全部 {len(rdf)} 条生成明细", expanded=False):
                    st.dataframe(rdf, use_container_width=True, hide_index=True)
            if preview_png:
                st.markdown("**首张预览：**")
                st.image(preview_png, width=360)

if st.session_state.download_data is not None:
    st.divider()
    st.download_button("📥 下载生成结果 (.zip)", data=st.session_state.download_data,
                       file_name=st.session_state.download_name, mime="application/zip",
                       type="primary", use_container_width=True)
