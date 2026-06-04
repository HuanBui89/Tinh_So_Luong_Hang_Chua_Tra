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

    # Đọc thô để lấy tên sale
    df_raw = pd.read_excel(
        uploaded_file,
        sheet_name=first_sheet,
        header=None
    )

    sale_name = "Không xác định"

    for i in range(min(20, len(df_raw))):

        row_text = " ".join(
            df_raw.iloc[i]
            .fillna("")
            .astype(str)
            .tolist()
        )

        if "Tên khách hàng:" in row_text:

            sale_name = (
                row_text
                .split("Tên khách hàng:")[1]
                .strip()
            )

            break

    # Header thực tế ở dòng 4
    df = pd.read_excel(
        uploaded_file,
        sheet_name=first_sheet,
        header=3
    )

    df.columns = [
        str(col).strip()
        for col in df.columns
    ]

    # Chỉ lấy cột cần thiết
    df = df[
        [
            "Số chứng từ",
            "Diễn giải",
            "Số lượng"
        ]
    ].copy()

    # Bỏ dòng không có mã hàng
    df = df.dropna(
        subset=["Diễn giải"]
    )

    # Chỉ giữ XK/NK
    df["Số chứng từ"] = (
        df["Số chứng từ"]
        .astype(str)
        .str.upper()
        .str.strip()
    )

    df = df[
        df["Số chứng từ"]
        .str.startswith(("XK", "NK"))
    ]

    df["Số lượng"] = pd.to_numeric(
        df["Số lượng"],
        errors="coerce"
    ).fillna(0)

    df["Tên sale"] = sale_name

    df["Mã hàng"] = (
        df["Diễn giải"]
        .astype(str)
        .str.strip()
    )

    df["Xuất kho"] = df.apply(
        lambda x: x["Số lượng"]
        if x["Số chứng từ"].startswith("XK")
        else 0,
        axis=1
    )

    df["Nhập kho"] = df.apply(
        lambda x: x["Số lượng"]
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
def export_excel(uploaded_file, outstanding):

  output = BytesIO()

# Tạo dòng tổng cộng
total_row = pd.DataFrame({
    "Mã hàng": ["TỔNG CỘNG"],
    "Xuất kho": [outstanding["Xuất kho"].sum()],
    "Nhập kho": [outstanding["Nhập kho"].sum()],
    "Chưa trả": [outstanding["Chưa trả"].sum()]
})

export_df = pd.concat(
    [outstanding, total_row],
    ignore_index=True
)

with pd.ExcelWriter(
    output,
    engine="openpyxl"
) as writer:

    # ==========================
    # SHEET 1 - GIỮ NGUYÊN FILE GỐC
    # ==========================

    original_excel = pd.ExcelFile(uploaded_file)

    first_sheet = original_excel.sheet_names[0]

    original_df = pd.read_excel(
        uploaded_file,
        sheet_name=first_sheet,
        header=None
    )

    original_df.to_excel(
        writer,
        sheet_name=first_sheet[:31],
        index=False,
        header=False
    )

    # ==========================
    # SHEET 2 - HÀNG CHƯA TRẢ
    # ==========================

    export_df = export_df[
        [
            "Mã hàng",
            "Xuất kho",
            "Nhập kho",
            "Chưa trả"
        ]
    ]

    export_df.columns = [
        "Mã hàng",
        "Số lượng xuất",
        "Số lượng nhập",
        "Số lượng chưa trả"
    ]

    export_df.to_excel(
        writer,
        sheet_name="Hàng chưa trả",
        index=False
    )

    # Format Excel
    ws = writer.book["Hàng chưa trả"]

    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter

        for cell in column:

            try:
                max_length = max(
                    max_length,
                    len(str(cell.value))
                )
            except:
                pass

        ws.column_dimensions[
            column_letter
        ].width = max_length + 5

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
         uploaded_file,
            outstanding
        )

st.download_button(
    label="📥 Download Excel",
    data=excel_file,
    file_name="Hang_Chua_Tra.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

