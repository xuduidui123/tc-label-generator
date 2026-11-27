"""Pure data rules shared by validation, planning and rendering."""
import re
import json
import hashlib
from decimal import Decimal, InvalidOperation
from xml.sax.saxutils import escape as xml_escape
import pandas as pd

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
    try:
        return positive_integer(row.get("外箱装量"), "外箱装量") == positive_integer(row.get("内盒装量"), "内盒装量")
    except ValueError:
        return False


def no_inner(row):
    return not safe_str(row.get("内盒装量")) or qty_equal(row)


def sanitize_name(name):
    """净化文件名：路径分隔符拍平为下划线，防穿越，纯点名回退。"""
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', normalize_spaces(name)).strip().rstrip('.')
    if not cleaned or set(cleaned) <= {"."}:
        return "unnamed"
    return cleaned


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


def _strip_spaces(s):
    return "".join(ch for ch in s if not ch.isspace() and ch != "\xa0")


def validate_itf(itf_val, visual):
    v = itf_val.strip()
    if not re.fullmatch(r"[0-9]{14}", v):
        raise ValueError("ITF 必须是 14 位数字")
    if visual and _strip_spaces(visual) != v:
        raise ValueError("ITF 目视号码与条码值不一致")


def validate_code128(bd_val, visual):
    v = bd_val.strip()
    if not re.fullmatch(r"02[0-9]{14}37[0-9]{1,8}", v):
        raise ValueError("BD 必须由 02 + 14 位 ITF + 37 + 1 至 8 位正整数数量组成")
    if int(v[18:]) <= 0:
        raise ValueError("BD 数量必须大于 0")
    if visual and _strip_spaces(visual) != v:
        raise ValueError("BD 目视号码与条码值不一致")


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
    return str(positive_integer(val, "数量")) if safe_str(val) else ""


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


NAME_COL = {"outer": "外箱文件命名列", "inner": "内盒文件命名列",
            "itf_uk": "ITF_UK命名", "itf_ce": "ITF_CE命名", "bd": "BD命名"}


TAPE_LOOKUP = [
    ("Halloween", "橙色"), ("Christmas", "请与客户确认"), ("Valentines", "红色"),
    ("Mother’s Day", "红色"), ("Father’s Day", "红色"), ("Easter", "红色"), ("Gardening", "绿色"),
]


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
    if not safe_str(inner_qty):
        return None
    oq = positive_integer(outer_qty, "外箱装量")
    iq = positive_integer(inner_qty, "内盒装量")
    if oq % iq:
        raise ValueError("外箱装量必须能被内盒装量整除")
    return oq // iq


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
    return f"02{v}37{r:02d}"


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
    """统一重算派生字段，旧 Excel 公式缓存不参与标签生成。"""
    row = dict(raw)
    for key in DERIVED_COLUMNS:
        row.pop(key, None)

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
        try:
            row["外箱BD号"] = derive_bd_num(cur("外箱装量"), cur("内盒装量"))
        except ValueError:
            row["外箱BD号"] = ""
    if not cur("BD条码号"):
        try:
            row["BD条码号"] = derive_bd_barcode(itf, cur("外箱装量"), cur("内盒装量"))
        except ValueError:
            row["BD条码号"] = ""
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


CDU_COLS = ["ITF", "外箱接触CDU", "内盒接触CDU"]


def cdu_df_to_map(df):
    m = {}
    if df is None or len(df) == 0: return m
    if not set(CDU_COLS).issubset(df.columns): raise ValueError("缺少 CDU 清单列")
    for _, row in df.iterrows():
        key = _strip_spaces(safe_str(row.get("ITF")))
        if not key: continue
        validate_itf(key, key)
        flags = tuple(safe_str(row.get(c)).upper() for c in CDU_COLS[1:])
        if any(flag not in ("", "Y", "N") for flag in flags):
            raise ValueError(f"{key} 的 CDU 标志必须为 Y、N 或留空")
        if key in m and m[key] != flags: raise ValueError(f"ITF {key} 存在相互冲突的 CDU 设置")
        m[key] = flags
    return m


_DATA_FIELDS = ["产品品名", "产品英文名", "PO号", "总外箱数", "TPNB", "TPND", "ITF",
                "category", "EAN", "外箱装量", "内盒装量", "CEORMSNO", "SKU", "VSN",
                "毛重", "净重", "外箱长", "外箱宽", "外箱高"]


_REQ_UK = [c for c in _DATA_FIELDS if c not in ("CEORMSNO", "SKU")]   # 380/520：除 CEORMSNO/SKU 外都必填


_REQ_CE = [c for c in _DATA_FIELDS if c not in ("TPNB", "TPND")]      # 999：除 TPNB/TPND 外都必填





_DIGIT_LEN = {"TPNB": 9, "TPND": 9, "ITF": 14, "EAN": 13, "CEORMSNO": 13, "SKU": 9}


def _iss(idx, po, col, msg):
    return {"idx": idx, "行": idx + 2, "PO号": po or "（PO号为空）", "列": col, "问题": msg}








def preflight(df, selected, cdu_map=None):
    issues = []
    for idx, raw in df.iterrows():
        row = enrich_row(raw, cdu_map)
        po = safe_str(row.get("PO号"))
        kinds = requested_kinds(row, selected)
        def error(col, msg):
            issues.append(_iss(idx, po, col, msg))
        for col in SWITCH_COLUMNS:
            value = safe_str(raw.get(col))
            if value and not (is_yes(value) or is_neg(value)):
                error(col, "只允许 Y、N 或留空")
        if not kinds:
            # Explicitly excluded rows are allowed; ambiguous automatic ITF selection is not.
            if "ITF" in selected and not derive_country(po) and not all(is_neg(row.get(c)) for c in ("是否需要ITF-UK", "是否需要ITF-CE")):
                error("PO号", "无法判断 ITF UK/CE，请填有效 PO 前缀或手动选择版本")
            continue
        required = {"产品品名"}
        numeric = set()
        if "outer" in kinds:
            required.update(["产品英文名", "PO号", "总外箱数", "category", "EAN", "毛重", "净重", "外箱长", "外箱宽", "外箱高", "外箱装量"])
            required.update(["TPNB", "TPND"] if po[:3] in ("380", "520") else ["CEORMSNO", "SKU"])
            numeric.update(["总外箱数", "外箱装量"])
            if not derive_country(po): error("PO号", "前缀必须为 380、520 或 999")
            for col in ("毛重", "净重", "外箱长", "外箱宽", "外箱高"):
                try: positive_number(row.get(col), col)
                except ValueError as exc: error(col, str(exc))
            try:
                if positive_number(row.get("净重"), "净重") > positive_number(row.get("毛重"), "毛重"):
                    error("净重", "净重不能大于毛重")
            except ValueError: pass
        if "inner" in kinds: required.update(["产品英文名", "category", "EAN"])
        if any(k.startswith("itf_") for k in kinds):
            required.update(["产品英文名", "ITF"])
            numeric.add("内盒装量" if safe_str(row.get("内盒装量")) else "外箱装量")
        if "itf_uk" in kinds: required.update(["TPNB", "TPND"])
        if "itf_ce" in kinds: required.add("CEORMSNO")
        if "bd" in kinds:
            required.add("ITF")
            numeric.update(["外箱装量", "内盒装量"])
        if safe_str(row.get("内盒装量")) and any(k in kinds for k in ("outer", "inner", "bd")):
            numeric.update(["外箱装量", "内盒装量"])
            try: _ratio(row.get("外箱装量"), row.get("内盒装量"))
            except ValueError as exc: error("外箱装量/内盒装量", str(exc))
        for col in sorted(required):
            value = safe_str(row.get(col))
            if not value: error(col, "本次标签必填")
            elif col in _DIGIT_LEN and not re.fullmatch(r"[0-9]{%d}" % _DIGIT_LEN[col], value):
                error(col, f"必须是 {_DIGIT_LEN[col]} 位数字")
        for col in sorted(numeric):
            try: positive_integer(row.get(col), col)
            except ValueError as exc: error(col, str(exc))
        if "bd" in kinds and safe_str(row.get("BD条码号")):
            try: validate_code128(row["BD条码号"], row["BD条码号加空格"])
            except ValueError as exc: error("BD条码号", str(exc))
    return issues


# Legacy spreadsheet formulas are never authoritative.
DERIVED_COLUMNS = {c['n'] for c in MERGED_SCHEMA if c['kind'] == 'formula'} | {
    '外箱 BD 号', '有CDU-加印文字'
}
SWITCH_COLUMNS = [c['n'] for c in MERGED_SCHEMA if c['kind'] == 'select' and c.get('dd') == 'YN']
KIND_TYPE = {'outer': '外箱', 'inner': '内盒', 'itf_uk': 'ITF', 'itf_ce': 'ITF', 'bd': 'BD'}


def positive_number(value, label):
    try:
        number = Decimal(safe_str(value))
        if not number.is_finite() or number <= 0:
            raise InvalidOperation
        return number
    except (InvalidOperation, ValueError):
        raise ValueError(f'{label}必须是大于 0 的有效数值') from None


def positive_integer(value, label):
    number = positive_number(value, label)
    if number != number.to_integral_value():
        raise ValueError(f'{label}必须是正整数，不能四舍五入')
    return int(number)


def requested_kinds(row, selected):
    country = derive_country(row.get('PO号'))
    kinds = []
    if '外箱' in selected and wants(row, '是否要做外箱', True): kinds.append('outer')
    if '内盒' in selected and not no_inner(row) and wants(row, '是否要做内盒', True): kinds.append('inner')
    if 'ITF' in selected:
        if wants(row, '是否需要ITF-UK', country in ('UK', 'Ireland')): kinds.append('itf_uk')
        if wants(row, '是否需要ITF-CE', country == 'CE'): kinds.append('itf_ce')
    if 'BD' in selected and not no_inner(row) and wants(row, '是否要做BD', True): kinds.append('bd')
    return kinds


# Identity includes product identifiers AND every field affecting the printed label.
CONTENT_FIELDS = {
    'outer': ['PO号','总外箱数','TPNB','TPND','SKU','CON不加粗','CON加粗','产品英文名','外箱BD号','有CDU加印文字','category','EAN','毛重','净重','外箱尺寸','其他备注','订单国家','BL','节日logo','双人抬标识'],
    'inner': ['产品英文名','category','EAN','BL','有CDU加印文字内盒版'],
    'itf_uk': ['产品英文名','TPNB','TPND','内盒装量','外箱装量','ITF加空格'],
    'itf_ce': ['产品英文名','CON不加粗','CON加粗','内盒装量','外箱装量','ITF加空格'],
    'bd': ['BD条码号','BD条码号加空格'],
}


def content_key(kind, row):
    fields = set(CONTENT_FIELDS[kind]) | {'ITF', 'EAN', '产品品名'}
    payload = [kind, {k: safe_str(row.get(k)) for k in sorted(fields)}]
    return hashlib.sha256(json.dumps(payload, ensure_ascii=False, sort_keys=True).encode()).hexdigest()


def build_plan(df, selected, cdu_map=None):
    plan, identities, filenames = [], {}, {}
    for idx, raw in df.iterrows():
        row = enrich_row(raw, cdu_map)
        for kind in requested_kinds(row, selected):
            key = content_key(kind, row)
            typ = KIND_TYPE[kind]
            duplicate = identities.get(key)
            if duplicate:
                name = duplicate['name']
            else:
                base = resolved_name(kind, row).encode("utf-8")[:180].decode("utf-8", errors="ignore")
                name, counter = base, 1
                while (typ, name.casefold()) in filenames:
                    counter += 1
                    name = f'{base}-{counter}'
                filenames[(typ, name.casefold())] = key
            item = dict(row=row, kind=kind, type=typ, key=key, name=name,
                        source_row=idx + 2, duplicate_of=duplicate['source_row'] if duplicate else None)
            identities.setdefault(key, item)
            plan.append(item)
    return plan


def input_fingerprint(df, selected, cdu_map):
    payload = {'rows': df.fillna('').astype(str).to_dict('records') if df is not None else [],
               'types': sorted(selected), 'cdu': cdu_map}
    return hashlib.sha256(json.dumps(payload, sort_keys=True, ensure_ascii=False).encode()).hexdigest()
