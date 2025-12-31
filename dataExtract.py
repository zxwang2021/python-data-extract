import glob
import csv
import pandas as pd


def read_csv_loose(path: str, encoding: str = "utf-8-sig") -> pd.DataFrame:
    """Read a CSV that may have ragged rows (inconsistent field counts)."""

    rows: list[list[str]] = []
    with open(path, "r", encoding=encoding, newline="", errors="replace") as f:
        reader = csv.reader(f)
        for r in reader:
            rows.append([cell.strip() for cell in r])

    if not rows:
        return pd.DataFrame()

    max_cols = max(len(r) for r in rows)
    padded = [r + [""] * (max_cols - len(r)) for r in rows]

    df = pd.DataFrame(padded)
    return df.replace({"": pd.NA})

for file in glob.glob("*.csv"):
    print(f"Processing {file}...")
    outFile = file.replace(".csv", ".xlsx")
    
    df = read_csv_loose(file, encoding="utf-8-sig")

    segments = []
    current_company = None
    current_header = None
    current_rows = []

    def flush_segment():
        if current_company and current_header and current_rows:
            segments.append({
                "company": current_company,
                "header": current_header,
                "rows": current_rows.copy()
            })

    global_headers = set()

    for i, row in df.iterrows():
        values = row.tolist()

        # 行是否“只有公司名”
        if pd.notna(values[0]) and row.count() == 1:
            flush_segment()
            current_company = values[0]
            current_header = None
            current_rows = []
            continue

        # 识别表头（本段第一个非空、含较多文字的行）
        if current_company and current_header is None and row.count() > 0:
            current_header = [str(x).strip() if pd.notna(x) else "" for x in values]
            global_headers.update([h for h in current_header if h])
            continue

        # 普通数据行
        if current_company and current_header and row.count() > 0:
            current_rows.append(values)

    # 最后一段
    flush_segment()

    # —— 构建“全局统一表头” ——
    global_headers = list(global_headers)
    global_headers.append("公司名称")

    records = []

    # —— 映射每段数据到统一表头 ——
    for seg in segments:
        header = seg["header"]
        company = seg["company"]

        for r in seg["rows"]:
            rec = {h: None for h in global_headers}

            for idx, value in enumerate(r):
                if idx < len(header) and header[idx]:
                    rec[header[idx]] = value

            rec["公司名称"] = company
            records.append(rec)

    result = pd.DataFrame(records, columns=global_headers)
    result.to_excel(outFile, index=False)

    print(f"完成，已生成 {outFile}")