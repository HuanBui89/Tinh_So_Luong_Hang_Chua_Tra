# App.py generated for Hang Chua Tra
import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Hàng chưa trả", page_icon="📦", layout="wide")
st.title("📦 THỐNG KÊ HÀNG CHƯA TRẢ")

def get_sale_name(df_raw):
    for i in range(min(20, len(df_raw))):
        row_text = " ".join(df_raw.iloc[i].fillna("").astype(str).tolist())
        if "Tên khách hàng:" in row_text:
            parts = row_text.split("Tên khách hàng:")
            if len(parts) > 1:
                return parts[1].strip()
    return "Không xác định"

def process_file(uploaded_file):
    excel = pd.ExcelFile(uploaded_file)
    first_sheet = excel.sheet_names[0]
    df_raw = pd.read_excel(uploaded_file, sheet_name=first_sheet, header=None)
    sale_name = get_sale_name(df_raw)

    df = pd.read_excel(uploaded_file, sheet_name=first_sheet, header=3)
    df.columns = [str(c).strip() for c in df.columns]

    required = ["Số chứng từ", "Diễn giải", "Số lượng"]
    for c in required:
        if c not in df.columns:
            st.error(f"Không tìm thấy cột: {c}")
            return None, None, None

    df = df[required].copy()
    df = df.dropna(subset=["Diễn giải"])
    df["Số chứng từ"] = df["Số chứng từ"].astype(str).str.upper().str.strip()
    df["Số lượng"] = pd.to_numeric(df["Số lượng"], errors="coerce").fillna(0)

    df = df[df["Số chứng từ"].str.startswith(("XK", "NK"))]

    df["Xuất kho"] = df.apply(lambda x: x["Số lượng"] if x["Số chứng từ"].startswith("XK") else 0, axis=1)
    df["Nhập kho"] = df.apply(lambda x: x["Số lượng"] if x["Số chứng từ"].startswith("NK") else 0, axis=1)

    summary = df.groupby("Diễn giải", as_index=False).agg({"Xuất kho":"sum","Nhập kho":"sum"})
    summary.rename(columns={"Diễn giải":"Mã hàng"}, inplace=True)
    summary["Chưa trả"] = summary["Xuất kho"] - summary["Nhập kho"]

    outstanding = summary[summary["Chưa trả"] > 0].copy()
    outstanding = outstanding.sort_values("Chưa trả", ascending=False)

    return sale_name, summary, outstanding

def export_excel(uploaded_file, outstanding):
    output = BytesIO()

    total_row = pd.DataFrame({
        "Mã hàng":["TỔNG CỘNG"],
        "Xuất kho":[outstanding["Xuất kho"].sum()],
        "Nhập kho":[outstanding["Nhập kho"].sum()],
        "Chưa trả":[outstanding["Chưa trả"].sum()]
    })

    export_df = pd.concat([outstanding, total_row], ignore_index=True)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        excel = pd.ExcelFile(uploaded_file)
        first_sheet = excel.sheet_names[0]

        original_df = pd.read_excel(uploaded_file, sheet_name=first_sheet, header=None)
        original_df.to_excel(writer, sheet_name=first_sheet[:31], index=False, header=False)

        export_df = export_df[["Mã hàng","Xuất kho","Nhập kho","Chưa trả"]]
        export_df.columns = ["Mã hàng","Số lượng xuất","Số lượng nhập","Số lượng chưa trả"]
        export_df.to_excel(writer, sheet_name="Hàng chưa trả", index=False)

    output.seek(0)
    return output

uploaded_file = st.file_uploader("Upload file Excel", type=["xlsx","xls"])

if uploaded_file:
    sale_name, summary, outstanding = process_file(uploaded_file)

    if summary is not None:
        st.success(f"Sale: {sale_name}")

        c1, c2, c3 = st.columns(3)
        c1.metric("Tổng mã hàng", len(summary))
        c2.metric("Tổng xuất", int(summary["Xuất kho"].sum()))
        c3.metric("Tổng chưa trả", int(outstanding["Chưa trả"].sum()))

        st.dataframe(outstanding, use_container_width=True, hide_index=True)

        excel_file = export_excel(uploaded_file, outstanding)

        st.download_button(
            "📥 Download Excel",
            data=excel_file,
            file_name="Hang_Chua_Tra.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
