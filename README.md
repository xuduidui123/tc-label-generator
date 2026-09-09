# Tesco 标贴与纸箱生成系统

从 Excel 数据源批量生成 Tesco 外箱标 / 内盒标 / ITF 条码标 / BD 条码标（导出为打包 PDF）。基于 Streamlit + LibreOffice。

> 仅供 Suncha 内部使用。
>
> 

## 功能（v2）
- **网页直接录入**：在网页表格里粘贴/编辑原始数据即可，**不再依赖 Excel 公式**；也可上传 Excel（原始数据即可）。
- **公式全后台化**：命名、CON、外箱尺寸、UK/CE、双人抬（按毛重）、BL（按英文名）、ITF加空格（3-5-6）、BD条码号、胶带颜色等全部由程序自动算。
- **CDU 按 ITF 自动匹配**：维护一份 CDU 清单（ITF→外箱/内盒是否接触 CDU），系统按 ITF 自动给标签加印 CDU 文案；清单可在网页查看/编辑/导入/导出。
- **每行开关留空即自动**：外箱默认做、有内盒才做内盒/BD、ITF 按订单国家；只有明确不做才填 N。
- **一款产品只做一次**：ITF/BD/内盒 同款自动去重；外箱按每个 PO 一张。
- 勾选要生成的类型 →「一键生成」→ 按类型分子文件夹打包 ZIP；生成前可体检，生成后给首张预览。
- 数据格式：总外箱数/装量为整数（ITF 的 UNIT SIZE 只显示整数）；TPNB/TPND/ITF/EAN/CEORMSNO/SKU 文本护零；毛重/净重/尺寸可小数。

## 本地运行
```bash
# 1) 系统依赖（Debian/Ubuntu）
sudo apt update && sudo apt install -y libreoffice fonts-noto-cjk fontconfig

# 2) Python 依赖
pip install -r requirements.txt

# 3) 启动
streamlit run app.py
```

## 部署到网页（Streamlit Community Cloud）

本项目已按 Streamlit Community Cloud 规范配置，`packages.txt` 会在云端自动安装 LibreOffice 与中日韩字体。

1. 把本文件夹作为一个 GitHub 仓库推送（见下方“推送到 GitHub”）。
2. 打开 https://share.streamlit.io ，用 GitHub 账号登录。
3. **New app** → 选择你的仓库、分支 `main`、主文件 `app.py` → **Deploy**。
4. 首次部署会安装 LibreOffice，冷启动较慢（数分钟），属正常现象。

### 限制访问（内部使用）
两层，任选或叠加：

- **应用私有化（推荐）**：在应用的 **Settings → Sharing** 中关闭公开，改为 “Only specific people”，按邮箱邀请内部同事。只有被邀请的 Google/GitHub 邮箱能打开。
- **口令门（补充）**：在 **Settings → Secrets** 里填入
  ```toml
  app_password = "你的内部口令"
  ```
  保存后，应用会先要求输入口令。未设置该密钥时不拦截（方便本地调试）。

## 推送到 GitHub
在本文件夹内执行：
```bash
git init
git add .
git commit -m "Tesco 标贴系统（修复版）"
git branch -M main
git remote add origin https://github.com/<你的用户名>/<仓库名>.git
git push -u origin main
```
> `.gitignore` 已排除数据源 `*.xlsx`、生成的 `*.zip` 和机密 `secrets.toml`，不会误传敏感数据。

## Excel 数据源（合并版·三色结构）
在网站左侧点「下载空白数据源模板」，或用仓库里的 `Tesco数据源模板.xlsx`。表头分三色，忠实沿用原「外箱内盒数据源」的公式设计并合并了标贴源：

- 🟩 **绿色 = 手填数据**（顺序固定 A–S）：产品品名、产品英文名、PO号、总外箱数、TPNB、TPND、ITF、category、EAN、外箱装量、内盒装量、CEORMSNO、SKU、VSN、毛重、净重、外箱长/宽/高。代码类列已设文本格式，**保护前导 0**。
- 🟧 **橙色 = 公式列**（Excel 自动算，勿手改）：订单国家、CON不加粗/加粗/CON、外箱BD号、外箱尺寸、双人抬标识、BL、ITF加空格、BD条码号、BD条码号加空格、胶带颜色、有CDU加印文字(内盒版)、五个文件命名列。
- 🟨 **黄色 = 人工选择**（下拉）：节日logo、是否有CDU、内盒是否直接接触CDU、是否要做外箱/内盒、是否需要ITF-UK/ITF-CE、是否要做BD（填 `Y`/`N`）。

**重要：必须用 Excel 打开填写并保存**，让公式计算出结果后再上传——网站读取的是公式算出的值（若某公式列为空，程序会回退用内置规则兜底命名/尺寸）。

公式要点（已修好原表中的 `#REF!`）：
- 命名：外箱 `{PO}{品名}{总箱数}个外箱`；内盒 `内盒-{品名}`；ITF-UK/CE `ITF-UK/CE-{品名}-{英文名}`；BD `BD-{品名}`。
- 双人抬：380 单件毛重>23 或 999 毛重>14 自动标「双人抬」。
- BL：英文名含 Go Cook/Tesco→Tesco，含 F&F Home→FF。
- BD条码号：`02`+ITF+`37`+（外箱装量/内盒装量），装量相等时不产生 BD。
- ITF 位数须偶数（ITF-14 为 14 位）；奇数位报错而**不会**被静默补 0。条码目视号去空格后须与条码值一致。

## 本次修复与改造要点
详见 `修复说明.md`。
