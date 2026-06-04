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
