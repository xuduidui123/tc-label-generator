import os
import sys
import shutil
import tempfile
import zipfile
import subprocess
import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate, InlineImage, RichText
from docx import Document as DocxDocument
from docx.shared import Mm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn as oxqn
import barcode
from barcode.writer import ImageWriter
from PIL import Image, ImageOps

# ==========================================
# 页面基础设置
# ==========================================
st.set_page_config(page_title="Tesco 标贴与纸箱生成系统", page_icon="📦", layout="centered")

# 初始化 session_state 用于暂存生成的压缩包，防止页面刷新丢失
if 'download_data' not in st.session_state:
    st.session_state.download_data = None
if 'download_name' not in st.session_state:
    st.session_state.download_name = ""

# ==========================================
# 辅助函数区
# ==========================================
def safe_str(val):
    if pd.isna(val): return ""
    if isinstance(val, (float, int)) and val == int(val): return str(int(val))
    return str(val).strip()

def crop_vertical_margin(image_path):
    with Image.open(image_path) as img:
        gray = img.convert('L')
        inverted = ImageOps.invert(gray)
        bbox = inverted.getbbox()
        if bbox:
            left, upper, right, lower = 0, bbox[1], img.width, bbox[3]
            img.crop((left, upper, right, lower)).save(image_path)

def get_resource_path(relative_path):
    # 资源文件与 app.py 放在同一目录下
    base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

def convert_docx_to_pdf(docx_path, output_dir):
    """
    使用 LibreOffice 将 docx 转换为 pdf。
    替代原来依赖 Mac 本地 Word 的 docx2pdf。
    服务器需提前安装：sudo apt install libreoffice -y
    """
    result = subprocess.run(
        [
            'libreoffice',
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', output_dir,
            docx_path
        ],
        capture_output=True,
        text=True,
        timeout=60  # 单个文件转换超时 60 秒
    )
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice 转换失败: {result.stderr}")

def create_zip(source_dir):
    """将生成的 PDF 文件夹打包为 ZIP"""
    zip_path = os.path.join(tempfile.gettempdir(), "Tesco_Labels.zip")
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(source_dir):
            for file in files:
                if file.endswith('.pdf'):
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, source_dir)
                    zipf.write(file_path, arcname)
    return zip_path

def calc_fit_font_pt(text, max_mm, font_path=None, max_pt=8.0, min_pt=4.0):
    """用 PIL 精确测量，找到能让 text 在 max_mm 宽度内单行显示的最大字号（pt）。
    使用 288 DPI 避免低分辨率四舍五入误差，并留 20% 余量补偿 LibreOffice 渲染比 PIL 宽约 16% 的问题。
    """
    if not text:
        return max_pt
    try:
        from PIL import ImageFont
        RENDER_DPI = 288.0  # 高分辨率避免像素取整误差
        LIBREOFFICE_MARGIN = 0.80  # PIL 测量值留 20% 余量，补偿 LibreOffice 渲染偏宽
        max_px = max_mm * LIBREOFFICE_MARGIN * RENDER_DPI / 25.4
        pt = max_pt
        while pt >= min_pt:
            size_px = max(1, round(pt * RENDER_DPI / 72.0))
            if font_path and os.path.exists(font_path):
                font_obj = ImageFont.truetype(font_path, size=size_px)
                bbox = font_obj.getbbox(text)
                if (bbox[2] - bbox[0]) <= max_px:
                    return pt
            else:
                break
            pt -= 0.5
        return min_pt
    except Exception:
        return max(min_pt, min(max_pt, max_mm / max(len(text), 1) / 0.31))

def _set_table_borders(table):
    """给表格所有边（含内部）加单线黑色边框。"""
    tbl = table._tbl
    tblPr = tbl.find(oxqn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    tblBorders = OxmlElement('w:tblBorders')
    for side in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{side}')
        border.set(oxqn('w:val'), 'single')
        border.set(oxqn('w:sz'), '6')
        border.set(oxqn('w:color'), '000000')
        tblBorders.append(border)
    tblPr.append(tblBorders)

def _set_bearer_bars(cell, sz=72):
    """给单元格四边加粗黑框，模拟条码 Bearer Bar（上下左右贯穿粗线）。sz 单位为 1/8pt，72=9pt≈3.2mm。"""
    tc = cell._tc
    tcPr = tc.find(oxqn('w:tcPr'))
    if tcPr is None:
        tcPr = OxmlElement('w:tcPr')
        tc.insert(0, tcPr)
    existing = tcPr.find(oxqn('w:tcBorders'))
    if existing is not None:
        tcPr.remove(existing)
    tcBorders = OxmlElement('w:tcBorders')
    for side in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{side}')
        border.set(oxqn('w:val'), 'single')
        border.set(oxqn('w:sz'), str(sz))
        border.set(oxqn('w:color'), '000000')
        tcBorders.append(border)
    tcPr.append(tcBorders)

@st.cache_resource
def _install_fonts():
    # 将 fonts/ 目录中的字体复制到用户字体目录，确保 LibreOffice 云端转换时字体一致
    fonts_src = get_resource_path("fonts")
    if not os.path.exists(fonts_src):
        return
    home_fonts = os.path.expanduser("~/.fonts")
    os.makedirs(home_fonts, exist_ok=True)
    for f in os.listdir(fonts_src):
        if f.lower().endswith('.ttf'):
            shutil.copy2(os.path.join(fonts_src, f), os.path.join(home_fonts, f))
    subprocess.run(["fc-cache", "-f"], capture_output=True)

_install_fonts()

# ==========================================
# 核心业务逻辑包装器
# ==========================================
def run_generation_task(task_name, df):
    # 使用系统临时目录，不再依赖 Mac Office 沙盒路径
    with tempfile.TemporaryDirectory() as work_dir:
        # work_dir：存放中间的 docx 和 barcode 图片
        # final_pdf_dir：只存放最终 PDF，用于打包
        final_pdf_dir = os.path.join(work_dir, "output_pdfs")
        os.makedirs(final_pdf_dir, exist_ok=True)

        success_count = 0
        errors = []
        total_items = len(df)

        progress_bar = st.progress(0)
        status_text = st.empty()

        for index, row in df.iterrows():
            try:
                if task_name == "外箱":
                    success = process_outer(row, work_dir, final_pdf_dir, index)
                elif task_name == "内盒":
                    success = process_inner(row, work_dir, final_pdf_dir, index)
                elif task_name == "ITF":
                    success = process_itf(row, work_dir, final_pdf_dir, index)
                elif task_name == "BD":
                    success = process_bd(row, work_dir, final_pdf_dir, index)

                if success:
                    success_count += success
            except Exception as e:
                errors.append(f"第 {index + 2} 行 发生错误: {str(e)}")

            current_progress = (index + 1) / total_items
            progress_bar.progress(current_progress)
            status_text.text(f"正在生成中... ({index + 1}/{total_items})")

        progress_bar.empty()
        status_text.empty()

        if success_count > 0:
            zip_file_path = create_zip(final_pdf_dir)
            with open(zip_file_path, "rb") as fp:
                st.session_state.download_data = fp.read()
            st.session_state.download_name = f"{task_name}生成结果_{success_count}个.zip"
            os.remove(zip_file_path)
            st.success(f"🎉 {task_name} 任务执行完毕！成功生成 {success_count} 个文件。")
        else:
            st.warning(f"⚠️ {task_name} 任务结束，但没有生成任何文件。请检查数据源。")

        if errors:
            with st.expander("点击查看失败详情"):
                for err in errors:
                    st.write(err)

# ==========================================
# 具体生成逻辑
# 参数变化：新增 work_dir 用于存放中间文件
# ==========================================
def process_outer(row, work_dir, output_dir, index):
    file_name_base = safe_str(row.get('外箱文件命名列'))
    if not file_name_base:
        return 0

    po_str = safe_str(row.get('PO号'))
    if not po_str:
        return 0

    template_path = get_resource_path("模板_OuterCarton.docx")
    tpl = DocxTemplate(template_path)
    context = {}

    is_uk_or_ireland = po_str.startswith('380') or po_str.startswith('520')
    is_ce = po_str.startswith('999')

    orms_str = ""
    for k in row.keys():
        if 'CE' in str(k).upper() and 'ORMS' in str(k).upper():
            orms_str = safe_str(row[k])
            break

    if is_uk_or_ireland:
        context['TPNB'], context['TPND'] = safe_str(row.get('TPNB')), safe_str(row.get('TPND'))
        context['CON不加粗'], context['CON加粗'], context['SKU'] = "", "", ""
    elif is_ce:
        context['TPNB'], context['TPND'] = "", ""
        context['SKU'] = safe_str(row.get('SKU'))
        con_normal, con_bold = safe_str(row.get('CON不加粗')), safe_str(row.get('CON加粗'))
        if not con_normal and not con_bold and orms_str:
            if len(orms_str) >= 5:
                con_normal, con_bold = orms_str[:-5], orms_str[-5:]
            else:
                con_normal = orms_str
        context['CON不加粗'], context['CON加粗'] = con_normal, con_bold
    else:
        return 0

    context['产品英文名'] = safe_str(row.get('产品英文名')).replace('&', '&amp;')
    context['外箱BD号'] = safe_str(row.get('外箱 BD 号', row.get('外箱BD号')))
    context['有CDU加印文字'] = safe_str(row.get('有CDU-加印文字', row.get('有CDU加印文字')))
    context.update({
        'PO号': po_str,
        'category': safe_str(row.get('category')).replace('&', '&amp;'),
        'EAN': safe_str(row.get('EAN')),
        '毛重': safe_str(row.get('毛重')),
        '净重': safe_str(row.get('净重')),
        '外箱尺寸': safe_str(row.get('外箱尺寸')),
        '其他备注': safe_str(row.get('其他备注')),
        '订单国家': safe_str(row.get('订单国家')),
        '总外箱数': safe_str(row.get('总外箱数'))
    })

    brand_name = safe_str(row.get('BL'))
    brand_img_path = get_resource_path(f"2. Brand Logo/{brand_name}.png")
    context['BL'] = InlineImage(tpl, brand_img_path, height=Mm(3)) if brand_name and os.path.exists(brand_img_path) else ""

    season_name = safe_str(row.get('节日logo'))
    season_img_path = get_resource_path(f"3. Seasonal Logo/{season_name}.png")
    context['节日logo'] = InlineImage(tpl, season_img_path, height=Mm(6)) if season_name and os.path.exists(season_img_path) else ""

    heavy_flag = safe_str(row.get('双人抬标识'))
    heavy_img_path = get_resource_path("4. 双人抬 Logo/双人抬.png")
    context['双人抬'] = InlineImage(tpl, heavy_img_path, height=Mm(6)) if heavy_flag == "双人抬" and os.path.exists(heavy_img_path) else ""

    file_name_base = file_name_base.replace("/", "_")

    temp_docx = os.path.join(work_dir, f"{file_name_base}.docx")
    final_pdf = os.path.join(output_dir, f"{file_name_base}.pdf")

    tpl.render(context)
    tpl.save(temp_docx)
    convert_docx_to_pdf(temp_docx, output_dir)

    # LibreOffice 输出文件名是将 .docx 改为 .pdf
    libreoffice_pdf = os.path.join(output_dir, f"{file_name_base}.pdf")
    if os.path.exists(libreoffice_pdf):
        if os.path.exists(temp_docx):
            os.remove(temp_docx)
        return 1
    return 0

def process_inner(row, work_dir, output_dir, index):
    file_name_base = safe_str(row.get('内盒文件命名列'))
    if not file_name_base:
        return 0

    file_name_base = file_name_base.replace("/", "_")

    template_path = get_resource_path("模板_InnerCarton.docx")
    tpl = DocxTemplate(template_path)
    context = {
        '产品英文名': safe_str(row.get('产品英文名')).replace('&', '&amp;'),
        'category': safe_str(row.get('category')).replace('&', '&amp;'),
        'EAN': safe_str(row.get('EAN')),
        '有CDU加印文字内盒版': safe_str(row.get('有CDU加印文字内盒版'))
    }

    brand_name = safe_str(row.get('BL'))
    brand_img_path = get_resource_path(f"2. Brand Logo/{brand_name}.png")
    context['BL'] = InlineImage(tpl, brand_img_path, height=Mm(3)) if brand_name and os.path.exists(brand_img_path) else ""

    temp_docx = os.path.join(work_dir, f"{file_name_base}.docx")

    tpl.render(context)
    tpl.save(temp_docx)
    convert_docx_to_pdf(temp_docx, output_dir)

    if os.path.exists(os.path.join(output_dir, f"{file_name_base}.pdf")):
        if os.path.exists(temp_docx):
            os.remove(temp_docx)
        return 1
    return 0

def process_itf(row, work_dir, output_dir, index):
    itf_val = safe_str(row.get('ITF'))
    uk_name = safe_str(row.get('ITF_UK命名'))
    ce_name = safe_str(row.get('ITF_CE命名'))

    if (not uk_name and not ce_name) or not itf_val:
        return 0

    eng_name_raw = safe_str(row.get('产品英文名'))
    inner_qty = safe_str(row.get('内盒装量'))
    itf_spaced = safe_str(row.get('ITF加空格', itf_val))

    # 用 PIL 计算自适应字号，确保品名在单行内显示（可用宽约 73mm）
    font_path = get_resource_path("fonts/Verdana Bold.ttf")
    fit_pt = calc_fit_font_pt(eng_name_raw, 73.0, font_path, max_pt=8.0, min_pt=4.0)
    rt_eng = RichText()
    rt_eng.add(eng_name_raw, bold=True, size=int(fit_pt * 2))

    ITF_Class = barcode.get_barcode_class('itf')
    barcode_obj = ITF_Class(itf_val, writer=ImageWriter())
    img_temp_path = os.path.join(work_dir, f"barcode_{index}")
    saved_img_path = barcode_obj.save(img_temp_path, options={"write_text": False, "module_height": 20.0})
    crop_vertical_margin(saved_img_path)

    template_path = get_resource_path("模板_ITF.docx")
    count = 0

    if uk_name:
        uk_name = uk_name.replace("/", "_")
        tpl_uk = DocxTemplate(template_path)
        context_uk = {
            'TPNB': safe_str(row.get('TPNB')),
            'TPND': safe_str(row.get('TPND')),
            'CON不加粗': "",
            'CON加粗': "",
            '产品英文名': rt_eng,
            '内盒装量': inner_qty,
            'ITF加空格': itf_spaced,
            'Barcode_Img': InlineImage(tpl_uk, saved_img_path, width=Mm(95), height=Mm(28))
        }
        temp_docx = os.path.join(work_dir, f"{uk_name}.docx")

        tpl_uk.render(context_uk)
        tpl_uk.save(temp_docx)
        convert_docx_to_pdf(temp_docx, output_dir)

        if os.path.exists(os.path.join(output_dir, f"{uk_name}.pdf")):
            count += 1
        if os.path.exists(temp_docx):
            os.remove(temp_docx)

    if ce_name:
        ce_name = ce_name.replace("/", "_")
        tpl_ce = DocxTemplate(template_path)
        orms_str = safe_str(row.get('CEORMSNO'))
        if len(orms_str) >= 5:
            con_normal, con_bold = orms_str[:-5], orms_str[-5:]
        else:
            con_normal, con_bold = orms_str, ""

        context_ce = {
            'TPNB': "",
            'TPND': "",
            'CON不加粗': safe_str(row.get('CON不加粗', con_normal)),
            'CON加粗': safe_str(row.get('CON加粗', con_bold)),
            '产品英文名': rt_eng,
            '内盒装量': inner_qty,
            'ITF加空格': itf_spaced,
            'Barcode_Img': InlineImage(tpl_ce, saved_img_path, width=Mm(95), height=Mm(28))
        }
        temp_docx = os.path.join(work_dir, f"{ce_name}.docx")

        tpl_ce.render(context_ce)
        tpl_ce.save(temp_docx)
        convert_docx_to_pdf(temp_docx, output_dir)

        if os.path.exists(os.path.join(output_dir, f"{ce_name}.pdf")):
            count += 1
        if os.path.exists(temp_docx):
            os.remove(temp_docx)

    if os.path.exists(saved_img_path):
        os.remove(saved_img_path)
    return count

def process_bd(row, work_dir, output_dir, index):
    bd_name = safe_str(row.get('BD命名'))
    bd_val = safe_str(row.get('BD条码号'))
    if not bd_name or not bd_val:
        return 0

    bd_name = bd_name.replace("/", "_")
    bd_spaced = safe_str(row.get('BD条码号加空格', bd_val))

    Code128_Class = barcode.get_barcode_class('code128')
    barcode_obj = Code128_Class(bd_val, writer=ImageWriter())
    img_temp_path = os.path.join(work_dir, f"barcode_{index}")
    saved_img_path = barcode_obj.save(img_temp_path, options={"write_text": False, "module_height": 22.0})
    crop_vertical_margin(saved_img_path)

    # 直接使用用户设计好的模板，填入条码图片和号码
    template_path = get_resource_path("模板_BD.docx")
    tpl = DocxTemplate(template_path)
    context = {
        'BD条码号加空格': bd_spaced,
        'Barcode_Img': InlineImage(tpl, saved_img_path, width=Mm(140), height=Mm(70))
    }
    tpl.render(context)
    temp_docx = os.path.join(work_dir, f"{bd_name}.docx")
    tpl.save(temp_docx)
    convert_docx_to_pdf(temp_docx, output_dir)

    if os.path.exists(os.path.join(output_dir, f"{bd_name}.pdf")):
        if os.path.exists(temp_docx):
            os.remove(temp_docx)
        if os.path.exists(saved_img_path):
            os.remove(saved_img_path)
        return 1
    return 0


# ==========================================
# UI 界面渲染
# ==========================================
st.title("📦 Tesco 标贴与纸箱生成系统")
st.markdown("**Made By Sherry | 仅供Suncha内部使用**")
st.divider()

# 1. 上传文件区
uploaded_file = st.file_uploader("📥 第一步：请上传您的 Excel 数据源文件", type=['xlsx', 'xls'])

if uploaded_file is not None:
    df = pd.read_excel(
        uploaded_file,
        dtype={
            'BD条码号': str,
            'BD条码号加空格': str,
            'TPNB': str,
            'TPND': str,
            'EAN': str,
            '外箱 BD 号': str,
            '外箱BD号': str,
            'ITF': str,
            'ITF加空格': str
        }
    )
    st.success(f"成功读取数据源！共识别到 {len(df)} 行数据。")

    st.markdown("🛠️ **第二步：请选择要执行的生成任务**")

    col1, col2, col3, col4 = st.columns(4)

    if col1.button("生成外箱标 (Outer)", use_container_width=True):
        st.session_state.download_data = None
        run_generation_task("外箱", df)

    if col2.button("生成内盒标 (Inner)", use_container_width=True):
        st.session_state.download_data = None
        run_generation_task("内盒", df)

    if col3.button("生成 ITF 标", use_container_width=True):
        st.session_state.download_data = None
        run_generation_task("ITF", df)

    if col4.button("生成 BD 标", use_container_width=True):
        st.session_state.download_data = None
        run_generation_task("BD", df)

st.divider()

# 3. 下载结果区
if st.session_state.download_data is not None:
    st.markdown("### ⬇️ 您的文件已打包完毕，请点击下载：")
    st.download_button(
        label="📥 点击下载生成结果 (.zip)",
        data=st.session_state.download_data,
        file_name=st.session_state.download_name,
        mime="application/zip",
        type="primary"
    )