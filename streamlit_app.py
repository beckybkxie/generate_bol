import os
import io
import zipfile
import tempfile
from datetime import date

import streamlit as st
import generate_bol as core

st.set_page_config(page_title="BOL HERE", page_icon="📦", layout="centered")

# ---------- 简单密码保护 ----------
# 密码在 Streamlit Cloud 的 Settings -> Secrets 里配置, 本地跑就用环境变量或默认值
PASSWORD = st.secrets.get("password", os.environ.get("BOL_APP_PASSWORD", "changeme"))

if "authed" not in st.session_state:
    st.session_state.authed = False

if not st.session_state.authed:
    st.title("📦 BOL HERE")
    pw = st.text_input("请输入访问密码", type="password")
    if st.button("登录"):
        if pw == PASSWORD:
            st.session_state.authed = True
            st.rerun()
        else:
            st.error("密码错误")
    st.stop()

# ---------- 主界面 ----------
st.title("BOL GENERATOR")
st.caption("上传腾讯文档导出的 CSV / Excel, 按日期批量生成BOL")

uploaded = st.file_uploader("上传出入库记录表 (CSV 或 xlsx)", type=["csv", "xlsx"])

col1, col2 = st.columns(2)
month = col1.number_input("月", min_value=1, max_value=12, value=date.today().month)
day = col2.number_input("日", min_value=1, max_value=31, value=date.today().day)

if uploaded is not None:
    if st.button("生成 BOL", type="primary"):
        with tempfile.TemporaryDirectory() as tmpdir:
            src_path = os.path.join(tmpdir, uploaded.name)
            with open(src_path, "wb") as f:
                f.write(uploaded.getbuffer())

            with st.spinner("读取数据中..."):
                rows = core.load_source_rows(src_path)

            st.write("识别到的字段：", ", ".join(rows[0].keys()) if rows else "(无数据)")

            target_rows = [r for r in rows if core.match_date(r.get("派送时间"), month, day)]
            st.write(f"筛选到 {month}/{day} 的记录: **{len(target_rows)}** 条")

            groups = core.group_by_shipto(target_rows)

            if not groups:
                st.warning("没有找到匹配的数据, 请检查日期或表格字段")
            else:
                out_dir = os.path.join(tmpdir, "output")
                os.makedirs(out_dir, exist_ok=True)
                core.OUT_DIR = out_dir

                today_str = date.today().strftime("%m/%d/%Y")
                results = []
                progress = st.progress(0.0)
                for i, (ship_to, items) in enumerate(groups, 1):
                    bol_number = f"EBOL-{month:02d}{day:02d}-{i:03d}"
                    out_xlsx = os.path.join(out_dir, f"{bol_number}.xlsx")
                    core.fill_bol(items, ship_to, bol_number, today_str, out_xlsx)
                    core.convert_to_pdf(out_xlsx)
                    pdf_path = os.path.join(out_dir, f"{bol_number}.pdf")
                    results.append((bol_number, len(items), ship_to, pdf_path, out_xlsx))
                    progress.progress(i / len(groups))

                st.success(f"生成了 {len(results)} 张 BOL")

                # 打包全部PDF成zip
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, "w") as zf:
                    for bol_number, n, ship_to, pdf_path, out_xlsx in results:
                        if os.path.exists(pdf_path):
                            zf.write(pdf_path, f"{bol_number}.pdf")

                        if os.path.exists(out_xlsx):
                            zf.write(out_xlsx, f"{bol_number}.xlsx")
                st.download_button(
                    "⬇️ 下载全部 PDF (zip)",
                    zip_buf.getvalue(),
                    file_name=f"BOL_{month:02d}{day:02d}.zip",
                    mime="application/zip",
                )

                st.divider()
                for bol_number, n, ship_to, pdf_path, out_xlsx in results:
                    with st.container(border=True):
                        st.write(f"**{bol_number}** — 明细{n}条 — {ship_to[:60]}")
                        if os.path.exists(pdf_path):
                            with open(pdf_path, "rb") as f:
                                st.download_button(
                                    f"下载 {bol_number}.pdf",
                                    f.read(),
                                    file_name=f"{bol_number}.pdf",
                                    mime="application/pdf",
                                    key=bol_number,
                                )

                        if os.path.exists(out_xlsx):
                            with open(out_xlsx, "rb") as f:
                                st.download_button(
                                    f"下载 {bol_number}.xlsx",
                                    f.read(),
                                    file_name=f"{bol_number}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key=bol_number + "_xlsx",
                                )
                        else:
                            st.info("PDF转换失败(可能是LibreOffice不可用), 但xlsx已生成, 请检查部署环境")
