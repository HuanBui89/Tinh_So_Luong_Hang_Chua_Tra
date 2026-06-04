import streamlit as st
import pandas as pd
from io import BytesIO

# ==================================================

# CẤU HÌNH

# ==================================================

st.set_page_config(
page_title="Hàng chưa trả",
page_icon="📦",
layout="wide"
)

st.title("📦 THỐNG KÊ HÀNG CHƯA TRẢ")

# ==================================================

# LẤY TÊN SALE

# ==================================================

def get_sale_name(df_raw):

for i in range(min(20, len(df_raw))):

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
            return "Không xác định"

return "Không xác định"


# ==================================================

# XỬ LÝ FILE

# ==================================================

def process_file(uploaded_file):


excel = pd.ExcelFile(uploaded_file)

first_sheet = excel.sheet_names[0]

# Đọc thô để lấy tên sale
df_raw = pd.read_excel(
    uploaded_file,
    sheet_name=first_sheet,
    header=None
)

sale_name = get_sale_name(df_raw)

# Header thực tế nằm ở dòng 4
df = pd.read_excel(
    uploaded_file,
    sheet_name=first_sheet,
    header=3
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

for col in required_cols:

    if col not in df.columns:

        st.error(f"Không tìm thấy cột: {col}")

        return None, None, None

df = df[
    [
        "Số chứng từ",
        "Diễn giải",
        "Số lượng"
    ]
].copy()

# Bỏ dòng trống
df = df.dropna(subset=["Diễn giải"])

# Chuẩn hóa
df["Số chứng từ"] = (
    df["Số chứng từ"]
    .astype(str)
    .str.upper()
    .str.strip()
)

df["Số lượng"] = pd.to_numeric(
    df["Số lượng"],
    errors="coerce"
).fillna(0)

# Chỉ lấy XK và NK
df = df[
    df["Số chứng từ"]
    .str.startswith(("XK", "NK"))
]

# Tính XK
df["Xuất kho"] = df.apply(
    lambda x:
    x["Số lượng"]
    if x["Số chứng từ"].startswith("XK")
    else 0,
    axis=1
)

# Tính NK
df["Nhập kho"] = df.apply(
    lambda x:
    x["Số lượng"]
    if x["Số chứng từ"].startswith("NK")
    else 0,
    axis=1
)

summary = (
    df.groupby(
        "Diễn giải",
        as_index=False
    )
    .agg(
        {
            "Xuất kho": "sum",
            "Nhập kho": "sum"
        }
    )
)

summary.rename(
    columns={
        "Diễn giải": "Mã hàng"
    },
    inplace=True
)

summary["Chưa trả"] = (
    summary["Xuất kho"]
    - summary["Nhập kho"]
)

outstanding = summary[
    summary["Chưa trả"] > 0
].copy()

return sale_name, summary, outstanding
```

# ==================================================

# XUẤT EXCEL

# ==================================================

def export_excel(uploaded_file, outstanding):

output = BytesIO()

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

    # Sheet gốc
    excel = pd.ExcelFile(uploaded_file)

    first_sheet = excel.sheet_names[0]

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

    # Sheet Hàng chưa trả
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

output.seek(0)

return output


# ==================================================

# GIAO DIỆN

# ==================================================

uploaded_file = st.file_uploader(
"Upload file Excel",
type=["xlsx", "xls"]
)

if uploaded_file:


sale_name, summary, outstanding = process_file(
    uploaded_file
)

if summary is not None:

    st.success(
        f"Sale: {sale_name}"
    )

    col1, col2, col3 = st.columns(3)

    col1.metric(
        "Tổng mã hàng",
        len(summary)
    )

    col2.metric(
        "Tổng xuất",
        int(summary["Xuất kho"].sum())
    )

    col3.metric(
        "Tổng chưa trả",
        int(outstanding["Chưa trả"].sum())
    )

    st.subheader("Danh sách hàng chưa trả")

    st.dataframe(
        outstanding,
        use_container_width=True,
        hide_index=True
    )

    excel_file = export_excel(
        uploaded_file,
        outstanding
    )

    st.download_button(
        "📥 Download Excel",
        data=excel_file,
        file_name="Hang_Chua_Tra.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

