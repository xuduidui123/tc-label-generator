import hashlib
import pandas as pd
import streamlit as st

# ==========================================
# 页面基础设置
# ==========================================
st.set_page_config(page_title="Tesco 标贴与纸箱生成系统", page_icon="📦", layout="wide")

for k, v in {"batch": None, "auth_ok": False}.items():
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


from engine import build_blank_template, load_cdu_df_bundled, run_all, preview_pdf
from rules import ALL_TYPES, TYPE_LABELS, TAPE_LOOKUP, cdu_df_to_map, preflight, build_plan, input_fingerprint

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
    if df.columns.duplicated().any():
        raise ValueError("存在重复表头，请按模板整理列名")
    key = [c for c in ["产品品名", "产品英文名", "PO号", "ITF"] if c in df.columns]
    if key:
        mask = df[key].apply(lambda r: any(str(v).strip() for v in r), axis=1)
        df = df[mask]
    return df.reset_index(drop=True)


st.title("📦 Tesco 标贴与纸箱生成系统")
st.caption("Made By Sherry ｜ 仅供 Suncha 内部使用 ｜ v3 · 体检拦截 + 批次核对")

# ---- CDU 清单（D）：按 ITF 自动匹配 ----
if "cdu_df" not in st.session_state:
    st.session_state.cdu_df = load_cdu_df_bundled()
with st.sidebar:
    st.header("🏷️ CDU 清单")
    st.caption("按 ITF 自动匹配：哪些产品外箱/内盒直接接触 CDU。填 Y 的会在标签上加印 CDU 文案。")
    cdu_up = st.file_uploader("导入 CDU 清单 CSV", type=["csv"], key="cdu_up")
    if cdu_up is not None:
        try:
            upload_key = hashlib.sha256(cdu_up.getvalue()).hexdigest()
            if st.session_state.get("cdu_upload_key") != upload_key:
                imported = pd.read_csv(cdu_up, dtype=str).fillna("")
                missing = {"ITF", "外箱接触CDU", "内盒接触CDU"} - set(imported.columns)
                if missing: raise ValueError("缺少列：" + "、".join(sorted(missing)))
                st.session_state.cdu_df = imported
                st.session_state.cdu_upload_key = upload_key
                st.session_state.pop("cdu_editor", None)
            st.success("CDU 清单已导入。")
        except Exception as e:
            st.error(f"导入失败：{e}")
    cdu_edited = st.data_editor(
        st.session_state.cdu_df, num_rows="dynamic", width="stretch", key="cdu_editor",
        column_config={
            "ITF": st.column_config.TextColumn("ITF", help="产品 ITF 条码值"),
            "外箱接触CDU": st.column_config.SelectboxColumn("外箱接触CDU", options=["", "Y", "N"]),
            "内盒接触CDU": st.column_config.SelectboxColumn("内盒接触CDU", options=["", "Y", "N"]),
        })
    st.download_button("⬇️ 导出当前 CDU 清单", cdu_edited.to_csv(index=False).encode("utf-8-sig"),
                       "cdu_list.csv", "text/csv", width="stretch")
    st.caption("云端重启会还原为仓库内置版；新增/修改后请导出 CSV 重新提交仓库（或下次导入）。")
    st.divider()
    st.download_button("⬇️ 下载空白 Excel 模板（备用）", data=build_blank_template(),
                       file_name="Tesco数据源模板.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       width="stretch")
try:
    cdu_map = cdu_df_to_map(cdu_edited)
except ValueError as exc:
    st.error(f"CDU 清单有误：{exc}")
    st.stop()

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
                            width="stretch", key="input_editor", column_config=colcfg)
    df = _clean_input_df(edited)
    if len(df):
        st.success(f"当前有效数据 {len(df)} 行。")
else:
    uploaded = st.file_uploader("上传数据源 Excel（原始数据即可，无需公式）", type=["xlsx"])
    if uploaded is not None:
        try:
            df = _clean_input_df(pd.read_excel(uploaded, dtype=str))
            st.success(f"读取成功，共 {len(df)} 行。派生字段已按原始数据重新计算。")
        except Exception as exc:
            st.error(f"读取失败，请使用下载的 .xlsx 模板：{exc}")
            df = None

# ---- 第二步：自动体检 + 选类型 + 生成 ----
selected = []
if df is not None and len(df):
    # 录入后自动体检（每次编辑实时刷新）；按 PO 合并显示 + 问题格标红
    st.markdown("**第二步：勾选本次要生成的类型**")
    cols = st.columns(4)
    selected = [t for i, t in enumerate(ALL_TYPES)
                if cols[i].checkbox(TYPE_LABELS[t], value=True, key=f"chk_{t}")]

    # 按所选类型体检（只勾选 ITF 时，只校验 ITF 所需字段，其余数据不必填/不报错）
    issues = preflight(df, selected, cdu_map) if selected else []
    if not selected:
        st.info("请在上方勾选要生成的类型。")
    elif issues:
        from collections import OrderedDict
        grouped = OrderedDict()
        for it in issues:
            grouped.setdefault(it["PO号"], []).append(it)
        st.error(f"⚠️ 数据体检发现 {len(issues)} 处问题，涉及 {len(grouped)} 个 PO（按 PO 对照上方表格修改）：")
        gp_rows = [{"PO号": po, "问题数": len(its),
                    "问题": "；".join((f"{i['列']}：{i['问题']}" if i["列"] else i["问题"]) for i in its)}
                   for po, its in grouped.items()]
        st.dataframe(pd.DataFrame(gp_rows), width="stretch", hide_index=True)
    else:
        st.success("✅ 数据体检通过，未发现问题。")

    plan = build_plan(df, selected, cdu_map) if selected else []
    unique_plan = [p for p in plan if p["duplicate_of"] is None]
    counts = {t: sum(p["type"] == t for p in unique_plan) for t in selected}
    st.caption("预计输出：" + " · ".join(f"{t} {n} 份" for t, n in counts.items()) + f"；相同内容 {len(plan) - len(unique_plan)} 项将在成功后合并。")
    with st.expander("生成前核对任务清单", expanded=False):
        if plan:
            st.dataframe(pd.DataFrame([{
                "行": p['source_row'], "类型": p['kind'], "文件名": p['name'],
                "处理": f"与第 {p['duplicate_of']} 行相同，成功后合并" if p['duplicate_of'] else "生成",
            } for p in plan]), hide_index=True, width="stretch")
        else:
            st.info("当前没有待生成项，请检查开关。")
    st.caption("修改数据、生成类型或 CDU 清单后，旧批次将失效；请重新生成。")
    if st.button("🚀 开始生成", type="primary", width="stretch",
                 disabled=bool(issues) or not plan):
        st.session_state.batch = None
        progress = st.progress(0.0)
        try:
            st.session_state.batch = run_all(df, selected, cdu_map, on_progress=progress.progress)
        except Exception as exc:
            st.error(f"生成未完成：{exc}")
        finally:
            progress.empty()

fingerprint = input_fingerprint(df, selected, cdu_map)
batch = st.session_state.batch
if batch and batch['fingerprint'] != fingerprint:
    st.session_state.batch = None
    batch = None
    st.info("数据或选项已变化，旧结果已失效，请重新生成。")
if batch:
    manifest = batch['manifest']
    st.divider()
    summary = f"生成 {manifest['pdf_count']} 个 PDF · 失败 {manifest['failed']} 项 · 合并 {manifest['skipped']} 项"
    if manifest['failed']: st.warning("本批次部分失败：" + summary)
    elif manifest['pdf_count']: st.success(summary)
    else: st.warning("没有生成任何 PDF。")
    st.caption(f"生成时间：{manifest['generated_at']} · 数据 {manifest['input_rows']} 行 · 批次 {batch['fingerprint'][:10]}")
    st.dataframe(pd.DataFrame(batch['results']), hide_index=True, width="stretch")
    if batch['pdfs']:
        choice = st.selectbox("预览标签（按类型 / 文件名）", sorted(batch['pdfs']))
        st.image(preview_pdf(batch['pdfs'][choice]), width=600)
        st.download_button("📥 下载本批次结果 (.zip)", data=batch['zip_data'],
                           file_name=batch['name'], mime="application/zip", type="primary")
        st.caption("ZIP 包含成功的 PDF、生成明细和批次信息。失败项不会包含在 PDF 中。")
