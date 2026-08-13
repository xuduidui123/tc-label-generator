import os
import io
import re
import uuid
import shutil
import tempfile
import zipfile
import subprocess
from xml.sax.saxutils import escape as xml_escape

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


def qty_equal(row):
    """内盒装量 == 外箱装量 → 该产品无内盒（返回 True 时不生成内盒标）。"""
    a, b = safe_str(row.get("外箱装量")), safe_str(row.get("内盒装量"))
    if not a or not b:
        return False
    try:
        return float(a) == float(b)
    except ValueError:
        return a == b


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


TYPE_LABELS = {"外箱": "外箱标 (Outer)", "内盒": "内盒标 (Inner)", "ITF": "ITF 标", "BD": "BD 标"}
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
    {"n": "总外箱数", "kind": "data"},
    {"n": "TPNB", "kind": "data", "text": True},
    {"n": "TPND", "kind": "data", "text": True},
    {"n": "ITF", "kind": "data", "text": True},
    {"n": "category", "kind": "data", "text": True},
    {"n": "EAN", "kind": "data", "text": True},
    {"n": "外箱装量", "kind": "data"},
    {"n": "内盒装量", "kind": "data"},
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
                    cell.number_format = "@"   # 仅对手填代码列设文本，保护前导0（不影响公式引用的数值列）

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
        "内盒装量=外箱装量时视为无内盒，自动不生成内盒标（即使勾了是否要做内盒=Y）。",
    ]
    for ri, t in enumerate(notes, start=1):
        ws3.cell(row=ri, column=1, value=("• " + t))
    ws3.column_dimensions["A"].width = 100

    ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


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
        "总外箱数": safe_str(row.get("总外箱数")),
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
    convert_docx_to_pdf(docx, out_dir, profile)
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
    convert_docx_to_pdf(docx, out_dir, profile)
    _confirm_pdf(out_dir, work_dir, name)
    return [{"命名": name, "状态": "成功", "说明": ""}]


def gen_itf(row, work_dir, out_dir, used, profile, want_uk, want_ce):
    itf_val = safe_str(row.get("ITF"))
    if not itf_val:
        raise ValueError("ITF 为空")
    itf_spaced = safe_str(row.get("ITF加空格", itf_val))
    validate_itf(itf_val, itf_spaced)

    eng_raw = safe_str(row.get("产品英文名"))
    inner_qty = safe_str(row.get("内盒装量"))
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
def preflight(df, selected):
    issues = []
    if "PO号" not in df.columns and "外箱" in selected:
        issues.append("缺少列：PO号（外箱标必需）")
    if "ITF" not in df.columns and "ITF" in selected:
        issues.append("缺少列：ITF（ITF 标必需）")
    if "BD条码号" not in df.columns and "BD" in selected:
        issues.append("缺少列：BD条码号（BD 标必需）")

    # 公式列整片为空 → 多半是没在 Excel 打开保存、公式未计算
    check_cols = [c for c in ["外箱文件命名列", "BD条码号", "双人抬标识", "ITF加空格", "CON加粗"] if c in df.columns]
    if check_cols and all(df[c].astype(str).str.strip().replace("nan", "").eq("").all() for c in check_cols):
        issues.append("⚠️ 公式列全为空：数据源似乎未在 Excel 中打开并保存，公式尚未计算。"
                      "请先用 Excel 打开、保存一次再上传，否则命名/条码等会缺失。")

    # 只报“真正的数据错误”。同款去重、无内盒无BD 属正常，生成时自动跳过，不在此报错。
    for idx, row in df.iterrows():
        rn = idx + 2
        if "外箱" in selected and is_yes(row.get("是否要做外箱")):
            po = safe_str(row.get("PO号"))
            if not po:
                issues.append(f"第{rn}行 外箱：PO号为空")
            elif not (po.startswith("380") or po.startswith("520") or po.startswith("999")):
                issues.append(f"第{rn}行 外箱：PO号'{po}'前缀非380/520/999")
        if "ITF" in selected and (is_yes(row.get("是否需要ITF-UK")) or is_yes(row.get("是否需要ITF-CE"))):
            try:
                validate_itf(safe_str(row.get("ITF")), safe_str(row.get("ITF加空格", safe_str(row.get("ITF")))))
            except ValueError as e:
                issues.append(f"第{rn}行 ITF：{e}")
        # BD 仅在有内盒(装量不等)且有条码值时校验；无内盒→无BD，属正常不报
        if "BD" in selected and is_yes(row.get("是否要做BD")) and not qty_equal(row):
            bdv = safe_str(row.get("BD条码号"))
            if bdv:
                try:
                    validate_code128(bdv, safe_str(row.get("BD条码号加空格", bdv)))
                except ValueError as e:
                    issues.append(f"第{rn}行 BD：{e}")
    return issues


def run_all(df, selected):
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

        for idx, row in df.iterrows():
            rn = idx + 2
            jobs = []
            if "外箱" in selected and is_yes(row.get("是否要做外箱")):
                jobs.append(("外箱", lambda r: gen_outer(r, work, subdirs["外箱"], used["外箱"], profile)))
            if "内盒" in selected and is_yes(row.get("是否要做内盒")) and not qty_equal(row):
                jobs.append(("内盒", lambda r: gen_inner(r, work, subdirs["内盒"], used["内盒"], profile)))
            if "ITF" in selected and (is_yes(row.get("是否需要ITF-UK")) or is_yes(row.get("是否需要ITF-CE"))):
                jobs.append(("ITF", lambda r: gen_itf(r, work, subdirs["ITF"], used["ITF"], profile,
                                                       is_yes(r.get("是否需要ITF-UK")), is_yes(r.get("是否需要ITF-CE")))))
            if "BD" in selected and is_yes(row.get("是否要做BD")) and not qty_equal(row):
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
# UI
# ==========================================
st.title("📦 Tesco 标贴与纸箱生成系统")
st.caption("Made By Sherry ｜ 仅供 Suncha 内部使用 ｜ 合并数据源版")

with st.sidebar:
    st.header("📄 数据源模板")
    st.write("第一次使用？下载空白模板，**用 Excel 打开填写并保存**，让公式先算好再上传。")
    st.download_button("⬇️ 下载空白数据源模板", data=build_blank_template(),
                       file_name="Tesco数据源模板.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)
    st.divider()
    st.markdown("**表头三色**：\n\n"
                "- 🟩 绿色：你手填的数据\n"
                "- 🟧 橙色：公式列，Excel 自动算（勿手改）\n"
                "- 🟨 黄色：人工选择（下拉 Y/N 或节日）\n\n"
                "命名、CON、外箱尺寸、双人抬、BL、ITF加空格、BD条码号等均由公式自动生成。")

uploaded = st.file_uploader("📥 第一步：上传数据源 Excel", type=["xlsx", "xls"])

if uploaded is not None:
    df = pd.read_excel(uploaded, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    st.success(f"读取成功，共 {len(df)} 行。")

    with st.expander("👀 预览数据（前 20 行）", expanded=False):
        preview = df.head(20).copy()
        preview.insert(0, "Excel行号", range(2, 2 + len(preview)))  # 表头是第1行，数据从第2行起
        st.caption("行号与 Excel 一致：表头为第 1 行，数据从第 2 行开始（下方报错的“第N行”即此行号）。")
        st.dataframe(preview, use_container_width=True, hide_index=True)

    st.markdown("**第二步：勾选本次要生成的类型**（每行还会看它自己的开关列）")
    cols = st.columns(4)
    selected = []
    for i, t in enumerate(ALL_TYPES):
        if cols[i].checkbox(TYPE_LABELS[t], value=True, key=f"chk_{t}"):
            selected.append(t)

    c1, c2 = st.columns([1, 1])
    if c1.button("🩺 生成前体检", use_container_width=True):
        if not selected:
            st.warning("请至少勾选一种类型。")
        else:
            issues = preflight(df, selected)
            if issues:
                st.error(f"发现 {len(issues)} 处隐患，建议修好再生成：")
                st.dataframe(pd.DataFrame({"问题": issues}), use_container_width=True, hide_index=True)
            else:
                st.success("体检通过，未发现明显问题 ✅")

    if c2.button("🚀 开始生成", type="primary", use_container_width=True):
        st.session_state.download_data = None
        if not selected:
            st.warning("请至少勾选一种类型。")
        else:
            results, actual, preview_png = run_all(df, selected)
            if actual > 0:
                st.success(f"🎉 完成！实际生成 {actual} 个 PDF（按类型分文件夹打包）。")
            else:
                st.warning("没有生成任何文件，请检查开关列与数据。")
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
