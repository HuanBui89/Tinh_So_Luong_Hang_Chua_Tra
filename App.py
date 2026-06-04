import streamlit as st
import pandas as pd
from io import BytesIO

# ==========================
# CẤU HÌNH TRANG
# ==========================

st.set_page_config(
    page_title="Quản lý hàng mượn Sales",
    page_icon="📦",
    layout="wide"
)

st.title("📦 QUẢN LÝ HÀNG MƯỢN SALES")
st.markdown("Upload file công nợ từ hệ thống để thống kê hàng mượn.")

# ==========================
# HÀM TÌM DÒNG HEADER
# ==========================

def find_header_row(df_raw):

    for idx in range(min(20, len(df_raw))):

        row_text = " ".join(
            df_raw.iloc[idx].astype(str).tolist()
        ).lower()

        if (
            "số chứng từ" in row_text
            and "diễn giải" in row_text
            and "số lượng" in row_text
        ):
            return idx

    return None


# ==========================
# XỬ LÝ FILE
# ==========================

def process_file(uploaded_file):

    excel = pd.ExcelFile(uploaded_file)

    all_data = []

    skip_sheet_keywords = [
        "nk",
        "xk",
        "chưa trả",
        "chua tra",
        "tong hop",
        "tổng hợp"
    ]

    for sheet_name in excel.sheet_names:

        sheet_lower = sheet_name.lower()

        if any(x in sheet_lower for x in skip_sheet_keywords):
            continue

        try:

            df_raw = pd.read_excel(
                uploaded_file,
                sheet_name=sheet_name,
                header=None
            )

            header_row = find_header_row(df_raw)

            if header_row is None:
                continue

            df = pd.read_excel(
                uploaded_file,
                sheet_name=sheet_name,
                header=header_row
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

            if not all(col in df.columns for col in required_cols):
                continue

            temp = df[
                [
                    "Số chứng từ",
                    "Diễn giải",
                    "Số lượng"
                ]
            ].copy()

            temp["Tên sale"] = sheet_name

            all_data.append(temp)

        except Exception:
            continue

    if len(all_data) == 0:
        return None, None, None

    data = pd.concat(
        all_data,
        ignore_index=True
    )

    # Làm sạch dữ liệu

    data["Số lượng"] = pd.to_numeric(
        data["Số lượng"],
        errors="coerce"
    ).fillna(0)

    data["Số chứng từ"] = (
        data["Số chứng từ"]
        .astype(str)
        .str.upper()
        .str.strip()
    )

    data["Mã hàng"] = (
        data["Diễn giải"]
        .astype(str)
        .str.strip()
    )

    # Phân loại XK / NK

    data["Xuất kho"] = data.apply(
        lambda x:
        x["Số lượng"]
        if x["Số chứng từ"].startswith("XK")
        else 0,
        axis=1
    )

    data["Nhập kho"] = data.apply(
        lambda x:
        x["Số lượng"]
        if x["Số chứng từ"].startswith("NK")
        else 0,
        axis=1
    )

    # Tổng hợp

    summary = (
        data.groupby(
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
        ["Tên sale", "Chưa trả"],
        ascending=[True, False]
    )

    # Hàng chưa trả

    outstanding = summary[
        summary["Chưa trả"] > 0
    ].copy()

    # Tổng hợp sale

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

    sale_summary = sale_summary.sort_values(
        "Chưa trả",
        ascending=False
    )

    return (
        summary,
        outstanding,
        sale_summary
    )


# ==========================
# XUẤT EXCEL
# ==========================

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


# ==========================
# UPLOAD FILE
# ==========================

uploaded_file = st.file_uploader(
    "Chọn file Excel",
    type=["xlsx", "xls"]
)

# ==========================
# XỬ LÝ
# ==========================

if uploaded_file:

    with st.spinner("Đang xử lý dữ liệu..."):

        summary, outstanding, sale_summary = process_file(
            uploaded_file
        )

    if summary is None:

        st.error(
            "Không tìm thấy dữ liệu hợp lệ."
        )

    else:

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

        # Bộ lọc sale

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

        # TAB

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
