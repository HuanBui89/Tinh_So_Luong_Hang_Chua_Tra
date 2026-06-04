import streamlit as st
import pandas as pd
from io import BytesIO

# =====================================
# CẤU HÌNH
# =====================================

st.set_page_config(
    page_title="Quản lý hàng mượn Sales",
    page_icon="📦",
    layout="wide"
)

st.title("📦 QUẢN LÝ HÀNG MƯỢN SALES")
st.write("Upload file công nợ từ hệ thống để thống kê hàng mượn.")

# =====================================
# HÀM XỬ LÝ FILE
# =====================================

def process_file(uploaded_file):

    try:

        excel = pd.ExcelFile(uploaded_file)

        # Luôn lấy sheet đầu tiên
        first_sheet = excel.sheet_names[0]

        sale_name = first_sheet

        df = pd.read_excel(
            uploaded_file,
            sheet_name=first_sheet,
            header=4
        )

        df.columns = [
            str(col).strip()
            for col in df.columns
        ]

        required_cols = [
            "Số chứng từ",
            "Diễn giải",
            "Số lượng"
        ]

        missing = [
            c for c in required_cols
            if c not in df.columns
        ]

        if missing:
            st.error(
                f"Thiếu cột: {', '.join(missing)}"
            )
            return None, None, None

        df = df[
            [
                "Số chứng từ",
                "Diễn giải",
                "Số lượng"
            ]
        ].copy()

        # Xóa dòng rỗng
        df = df.dropna(
            subset=["Diễn giải"]
        )

        # Chuẩn hóa
        df["Số lượng"] = pd.to_numeric(
            df["Số lượng"],
            errors="coerce"
        ).fillna(0)

        df["Số chứng từ"] = (
            df["Số chứng từ"]
            .astype(str)
            .str.upper()
            .str.strip()
        )

        df["Mã hàng"] = (
            df["Diễn giải"]
            .astype(str)
            .str.strip()
        )

        df["Tên sale"] = sale_name

        # XK
        df["Xuất kho"] = df.apply(
            lambda x:
            x["Số lượng"]
            if x["Số chứng từ"].startswith("XK")
            else 0,
            axis=1
        )

        # NK
        df["Nhập kho"] = df.apply(
            lambda x:
            x["Số lượng"]
            if x["Số chứng từ"].startswith("NK")
            else 0,
            axis=1
        )

        summary = (
            df.groupby(
                ["Tên sale", "Mã hàng"],
                as_index=False
            )
            .agg(
                {
                    "Xuất kho": "sum",
                    "Nhập kho": "sum"
                }
            )
        )

        summary["Chưa trả"] = (
            summary["Xuất kho"]
            - summary["Nhập kho"]
        )

        summary = summary.sort_values(
            by="Chưa trả",
            ascending=False
        )

        outstanding = summary[
            summary["Chưa trả"] > 0
        ].copy()

        sale_summary = (
            summary.groupby(
                "Tên sale",
                as_index=False
            )
            .agg(
                {
                    "Xuất kho": "sum",
                    "Nhập kho": "sum",
                    "Chưa trả": "sum"
                }
            )
        )

        return (
            summary,
            outstanding,
            sale_summary
        )

    except Exception as e:

        st.error(f"Lỗi xử lý file: {e}")

        return None, None, None


# =====================================
# EXPORT EXCEL
# =====================================

def export_excel(
    summary,
    outstanding,
    sale_summary
):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        summary.to_excel(
            writer,
            sheet_name="TongHop",
            index=False
        )

        outstanding.to_excel(
            writer,
            sheet_name="HangChuaTra",
            index=False
        )

        sale_summary.to_excel(
            writer,
            sheet_name="TongHopSale",
            index=False
        )

    output.seek(0)

    return output


# =====================================
# UPLOAD
# =====================================

uploaded_file = st.file_uploader(
    "Chọn file Excel",
    type=["xlsx", "xls"]
)

# =====================================
# XỬ LÝ
# =====================================

if uploaded_file:

    with st.spinner("Đang xử lý dữ liệu..."):

        summary, outstanding, sale_summary = process_file(
            uploaded_file
        )

    if summary is not None:

        # KPI

        col1, col2, col3 = st.columns(3)

        col1.metric(
            "Tổng mã hàng",
            len(summary)
        )

        col2.metric(
            "Tổng xuất kho",
            int(summary["Xuất kho"].sum())
        )

        col3.metric(
            "Tổng chưa trả",
            int(summary["Chưa trả"].sum())
        )

        st.divider()

        # Filter

        sale_list = [
            "Tất cả"
        ] + sorted(
            summary["Tên sale"]
            .unique()
            .tolist()
        )

        selected_sale = st.selectbox(
            "Chọn Sale",
            sale_list
        )

        view_df = summary.copy()

        if selected_sale != "Tất cả":

            view_df = view_df[
                view_df["Tên sale"]
                == selected_sale
            ]

        only_outstanding = st.checkbox(
            "Chỉ hiện hàng chưa trả"
        )

        if only_outstanding:

            view_df = view_df[
                view_df["Chưa trả"] > 0
            ]

        # TABS

        tab1, tab2, tab3 = st.tabs(
            [
                "📦 Tổng hợp",
                "⚠️ Hàng chưa trả",
                "👨‍💼 Theo Sale"
            ]
        )

        with tab1:

            st.subheader(
                "Tổng hợp hàng mượn"
            )

            st.dataframe(
                view_df,
                use_container_width=True,
                hide_index=True
            )

        with tab2:

            st.subheader(
                "Danh sách hàng chưa trả"
            )

            st.dataframe(
                outstanding,
                use_container_width=True,
                hide_index=True
            )

        with tab3:

            st.subheader(
                "Tổng hợp theo Sale"
            )

            st.dataframe(
                sale_summary,
                use_container_width=True,
                hide_index=True
            )

        st.divider()

        excel_file = export_excel(
            summary,
            outstanding,
            sale_summary
        )

        st.download_button(
            label="📥 Download Excel",
            data=excel_file,
            file_name="BaoCaoHangMuon.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
