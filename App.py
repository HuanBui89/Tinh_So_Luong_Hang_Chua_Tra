import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="Quản lý hàng mượn",
    page_icon="📦",
    layout="wide"
)

st.title("📦 QUẢN LÝ HÀNG MƯỢN SALES")


# =====================================================
# TÌM TÊN SALE
# =====================================================

def get_sale_name(df_raw):

    for i in range(min(30, len(df_raw))):

        row_text = " ".join(
            df_raw.iloc[i]
            .fillna("")
            .astype(str)
            .tolist()
        )

        if "Tên khách hàng:" in row_text:

            try:
                return row_text.split(
                    "Tên khách hàng:"
                )[1].strip()
            except:
                pass

    return "Không xác định"


# =====================================================
# TÌM HEADER
# =====================================================

def find_header_row(df_raw):

    for i in range(min(30, len(df_raw))):

        row_text = " ".join(
            df_raw.iloc[i]
            .fillna("")
            .astype(str)
            .tolist()
        ).lower()

        if (
            "số chứng từ" in row_text
            and "diễn giải" in row_text
            and "số lượng" in row_text
        ):
            return i

    return None


# =====================================================
# XỬ LÝ FILE
# =====================================================

def process_file(uploaded_file):

    excel = pd.ExcelFile(uploaded_file)

    first_sheet = excel.sheet_names[0]

    # đọc thô
    df_raw = pd.read_excel(
        uploaded_file,
        sheet_name=first_sheet,
        header=None
    )

    # tên sale
    sale_name = get_sale_name(df_raw)

    # dòng header
    header_row = find_header_row(df_raw)

    if header_row is None:
        st.error(
            "Không tìm thấy dòng tiêu đề."
        )
        return None, None, None

    # đọc lại đúng header
    df = pd.read_excel(
        uploaded_file,
        sheet_name=first_sheet,
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

    # bỏ dòng trống
    df = df.dropna(
        subset=["Diễn giải"]
    )

    df["Tên sale"] = sale_name

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
        "Chưa trả",
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


# =====================================================
# XUẤT EXCEL
# =====================================================

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


# =====================================================
# GIAO DIỆN
# =====================================================

uploaded_file = st.file_uploader(
    "Upload file Excel",
    type=["xlsx", "xls"]
)

if uploaded_file:

    with st.spinner(
        "Đang xử lý dữ liệu..."
    ):

        summary, outstanding, sale_summary = process_file(
            uploaded_file
        )

    if summary is not None:

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

        only_outstanding = st.checkbox(
            "Chỉ hiện hàng chưa trả"
        )

        view_df = summary.copy()

        if only_outstanding:

            view_df = view_df[
                view_df["Chưa trả"] > 0
            ]

        tab1, tab2, tab3 = st.tabs(
            [
                "📦 Tổng hợp",
                "⚠️ Hàng chưa trả",
                "👨‍💼 Theo Sale"
            ]
        )

        with tab1:

            st.dataframe(
                view_df,
                use_container_width=True,
                hide_index=True
            )

        with tab2:

            st.dataframe(
                outstanding,
                use_container_width=True,
                hide_index=True
            )

        with tab3:

            st.dataframe(
                sale_summary,
                use_container_width=True,
                hide_index=True
            )

        excel_file = export_excel(
            summary,
            outstanding,
            sale_summary
        )

        st.download_button(
            "📥 Download Excel",
            excel_file,
            file_name="BaoCaoHangMuon.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
