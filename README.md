# Tesco 标贴与纸箱生成系统

供 Suncha 内部使用。从网页表格或 Excel 原始数据批量生成外箱、内盒、ITF-UK/CE 和 BD 标签，导出 PDF ZIP。

## 使用

1. 在网页录入数据，或下载空白 `.xlsx` 模板填写后上传。无需 Excel 公式，不支持旧版 `.xls`。
2. 选择需要的标签类型。行级开关 `Y` 表示生成、`N` 表示不生成、空白按业务规则自动判断。
3. 修复体检提示，并核对预计数量和任务清单。存在错误时不能生成，后台执行入口也会再次校验。
4. 生成后按类型/文件名选择预览，再下载 ZIP。包内附生成明细 CSV 和批次信息 JSON；部分失败会明确提示，只打包成功 PDF。
5. 修改数据、类型或 CDU 清单后，旧下载结果失效，必须重新生成。

## 数据规则

- ITF、EAN、TPNB、TPND、CEORMSNO、SKU 按文本填写，保留前导 0。已经在 Excel 中丢掉的 0 无法自动恢复。
- 数量必须为正整数，不会四舍五入。毛重、净重、尺寸必须大于 0，净重不能超过毛重。
- 无内盒请留空内盒装量；内外装量相等也视为无内盒，自动跳过内盒/BD。填写内盒装量时需为正整数；涉及内盒/BD/外箱的装量比例必须整除。
- 校验按实际生成类型执行。仅生成 ITF 时不要求毛重、外箱尺寸等无关字段。UK/CE 可按 PO 前缀自动判断，也可用行级开关指定。
- 国家、尺寸、双人抬、条码、CDU 文案、文件名等派生字段统一由后台重算，上传旧表中的对应公式结果会被忽略。不支持隐式手工覆盖派生字段。
- 去重依据标签类型、产品标识和实际印刷字段。不同内容即使同名也会生成独立文件；同内容仅在首次成功后合并，失败后遇到相同记录仍可重试。
- CDU 清单可导入/导出 CSV；相同 ITF 的冲突设置会阻止生成。清单保存在当前会话，不是多人共享数据库，重新打开会话后需重新导入或配置仓库内置 `cdu_list.csv`。
- ZIP 中每份标签必须为单页且包含可读文字。长品名在最小字号下仍换行时会报告失败，不能把生成成功等同于实物印刷验收。

## 本地运行

建议 Python 3.12。Python 直接依赖固定在 `requirements.txt`，系统 LibreOffice 版本由部署环境提供。

```bash
pip install -r requirements.txt
streamlit run app.py
```

系统还需 LibreOffice；Linux 安装 `libreoffice fonts-noto-cjk fontconfig`，macOS 可安装 LibreOffice 应用。项目启动不会立即安装字体；首次生成时会将随附字体复制到用户字体目录，Linux 刷新 fontconfig。PDF 文本检查和预览通过 PyMuPDF 完成。

Streamlit Community Cloud 的系统依赖由 `packages.txt` 安装。配置应用访问权限以限定内部用户；可选口令门在 `.streamlit/secrets.toml` 或云端 Secrets 设置：

```toml
app_password = "替换为内部口令"
```

未配置口令时沿用开放访问行为，适合本地使用；部署访问权限由管理员配置。

## 维护与验证

- `app.py`：页面、访问口令、输入及批次状态。
- `rules.py`：无页面副作用的数据规则、体检、任务计划、输入指纹。
- `engine.py`：Word 模板渲染、PDF 验证、批量执行和打包。
- `tests/test_regressions.py`：错标/漏标边界、失败重试、数据变化、页面冒烟测试。
- `tests/render_smoke.py`：真实模板经 LibreOffice 生成 UK、CE、Ireland、长品名及同名不同码样例，检查单页、模板占位符、预期数量，并从最终 PDF 解码核对 ITF 和 GS1-128 BD。

```bash
pip install -r requirements-test.txt
python -m unittest discover -s tests -v
python tests/render_smoke.py
```

CI 配置模板位于 `ci/github-actions.yml`；将它复制到 `.github/workflows/tests.yml` 后启用。工作流会运行两类测试并保留 PDF/PNG 检查产物。`output/`、生成 ZIP、上传 Excel 和本地 Secrets 不纳入版本控制；仓库原有模板文件仍被保留。

原有 `修复说明.md` 为历史版本记录，以本说明和当前代码为准。尚未包含实物打印后的扫码与尺寸验收。
