import re
from pathlib import Path

import pandas as pd
import tabula


PDF_AREA = [250, 15, 800, 830]      # [top, left, bottom, right]
PDF_COLUMNS = [85, 380, 470, 830]   # split into: tanggal | keterangan | mutasi | saldo

GARBAGE_PATTERNS = [
    r"^TANGGAL$",
    r"^KETERANGAN$",
    r"^MUTASI$",
    r"^SALDO$",
    r"^CATATAN:?$",
    r"^\d+\s*/\s*\d+$",   # matches page markers like 1 / 4
    r"^BERSAMBUNG KE HALAMAN BERIKUT$",
]

SUMMARY_REGEX = r"^SALDO AWAL\s*:|^MUTASI CR\s*:|^MUTASI DB\s*:|^SALDO AKHIR\s*:"


def clean_text(value):
    """Normalize text: remove line breaks, trim spaces, collapse multiple spaces."""
    if pd.isna(value):
        return ""
    text = str(value).replace("\r", " ").replace("\n", " ").strip()
    return re.sub(r"\s+", " ", text)


def parse_number(value):
    """
    Convert numeric text like '6,903,500.00' or '36,000.00 DB' into int.
    Returns pd.NA if conversion fails.
    """
    text = str(value).strip().upper()

    if text in ("", "NAN"):
        return pd.NA

    text = text.replace("DB", "").replace("CR", "").replace(",", "").strip()

    try:
        return int(float(text))
    except ValueError:
        return pd.NA


def parse_mutasi(value):
    """
    Split Mutasi into DB and CR.
    Example:
    - '36,000.00 DB' -> DB=36000, CR=<NA>
    - '6,903,500.00' -> DB=<NA>, CR=6903500
    """
    text = str(value).strip().upper()

    if text in ("", "NAN"):
        return pd.Series([pd.NA, pd.NA], index=["DB", "CR"])

    number = parse_number(text)
    if pd.isna(number):
        return pd.Series([pd.NA, pd.NA], index=["DB", "CR"])

    if "DB" in text:
        return pd.Series([number, pd.NA], index=["DB", "CR"])

    return pd.Series([pd.NA, number], index=["DB", "CR"])


def is_garbage_row(row, garbage_patterns):
    """Detect repeated headers, page markers, and empty rows."""
    values = [str(v).strip() for v in row.tolist() if pd.notna(v) and str(v).strip()]
    joined = " ".join(values).strip()

    if not joined:
        return True

    return any(re.fullmatch(pattern, joined, flags=re.IGNORECASE) for pattern in garbage_patterns)


def append_text(base, extra):
    """Append extra text to base with a single space."""
    base = (base or "").strip()
    extra = (extra or "").strip()

    if not extra:
        return base
    if not base:
        return extra
    return f"{base} {extra}"


def read_pdf_table(pdf_file):
    """Read all pages from PDF using Tabula and combine them into one dataframe."""
    dfs = tabula.read_pdf(
        str(pdf_file),
        pages="all",
        stream=True,
        guess=False,
        area=PDF_AREA,
        columns=PDF_COLUMNS,
        pandas_options={"header": None},
        multiple_tables=False,
    )

    if not dfs:
        raise ValueError("No table found. Try adjusting area/columns slightly.")

    print(f"Proses membaca PDF selesai, ditemukan {len(dfs)} potongan tabel. Menggabungkan data...")

    df = pd.concat(dfs, ignore_index=True)
    df = df.iloc[:, :4].copy()
    df.columns = ["Tanggal", "Keterangan", "Mutasi", "Saldo"]

    for col in df.columns:
        df[col] = df[col].apply(clean_text)

    return df


def remove_garbage_rows(df):
    """Remove repeated headers, page markers, and other obvious non-transaction rows."""
    mask = df.apply(lambda row: is_garbage_row(row, GARBAGE_PATTERNS), axis=1)
    return df[~mask].reset_index(drop=True)


def split_summary_rows(df):
    """Separate transaction rows from summary rows at the bottom."""
    summary_mask = (
        df["Tanggal"].str.strip().eq("") &
        df["Keterangan"].str.upper().str.contains(SUMMARY_REGEX, regex=True, na=False)
    )

    print(f"Terdeteksi {summary_mask.sum()} baris summary. Memisahkan transaksi dan summary...")

    df_summary = df[summary_mask].reset_index(drop=True)
    df_trans = df[~summary_mask].reset_index(drop=True)

    return df_trans, df_summary


def merge_continuation_rows(df_trans):
    """
    Merge multiline transaction rows.
    If Tanggal is empty, the row is treated as continuation of previous row.
    """
    merged_rows = []
    current = None

    for _, row in df_trans.iterrows():
        tanggal = row["Tanggal"].strip()
        keterangan = row["Keterangan"].strip()
        mutasi = row["Mutasi"].strip()
        saldo = row["Saldo"].strip()

        if tanggal:
            if current is not None:
                merged_rows.append(current)

            current = {
                "Tanggal": tanggal,
                "Keterangan": keterangan,
                "Mutasi": mutasi,
                "Saldo": saldo,
            }
        else:
            if current is None:
                continue

            if keterangan:
                current["Keterangan"] = append_text(current["Keterangan"], keterangan)

            if mutasi:
                if current["Mutasi"] == "":
                    current["Mutasi"] = mutasi
                else:
                    current["Keterangan"] = append_text(current["Keterangan"], mutasi)

            if saldo:
                if current["Saldo"] == "":
                    current["Saldo"] = saldo
                else:
                    current["Keterangan"] = append_text(current["Keterangan"], saldo)

    if current is not None:
        merged_rows.append(current)

    return pd.DataFrame(merged_rows, columns=["Tanggal", "Keterangan", "Mutasi", "Saldo"])


def build_summary_dataframe(df_summary):
    """Convert summary rows into structured summary dataframe."""
    summary_clean = []

    for _, row in df_summary.iterrows():
        text = " ".join([row["Keterangan"], row["Mutasi"], row["Saldo"]]).strip()
        text = re.sub(r"\s+", " ", text)
        summary_clean.append(text)

    parsed_rows = []

    for row in summary_clean:
        parts = row.split(":", 1)

        if len(parts) == 2:
            label = parts[0].strip()
            values = parts[1].strip().split()

            amount = values[0] if len(values) > 0 else ""
            freq = values[1] if len(values) > 1 else ""

            parsed_rows.append({
                "Keterangan": label,
                "Amount": parse_number(amount),
                "Frekuensi": parse_number(freq),
            })

    return pd.DataFrame(parsed_rows, columns=["Keterangan", "Amount", "Frekuensi"])


def finalize_transactions(clean_df):
    """Create final transaction dataframe with DB, CR, and cleaned saldo."""
    clean_df[["DB", "CR"]] = clean_df["Mutasi"].apply(parse_mutasi)
    clean_df["Saldo"] = clean_df["Saldo"].apply(parse_number)

    clean_df = clean_df.drop(columns=["Mutasi"])
    clean_df = clean_df[["Tanggal", "Keterangan", "DB", "CR", "Saldo"]]

    return clean_df


def auto_fit_columns(sheet):
    """Auto-fit Excel column widths based on content length."""
    for col_cells in sheet.columns:
        col_letter = col_cells[0].column_letter
        max_length = max(
            len(str(cell.value)) if cell.value is not None else 0
            for cell in col_cells
        )
        sheet.column_dimensions[col_letter].width = max_length + 2


def export_to_excel(clean_df, df_summary_final, output_file):
    """Write transaction and summary dataframes to Excel."""
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        clean_df.to_excel(writer, sheet_name="Transaksi", index=False)
        df_summary_final.to_excel(writer, sheet_name="Summary", index=False)

        sheet1 = writer.sheets["Transaksi"]
        sheet2 = writer.sheets["Summary"]

        auto_fit_columns(sheet1)
        auto_fit_columns(sheet2)


def main():
    file_name = input("Enter the filename without extension: ").strip()

    pdf_dir = Path("pdf_file")
    excel_dir = Path("excel_file")
    excel_dir.mkdir(parents=True, exist_ok=True)

    pdf_file = pdf_dir / f"{file_name}.pdf"
    output_file = excel_dir / f"{file_name}.xlsx"

    if not pdf_file.exists():
        raise FileNotFoundError(f"PDF file not found: {pdf_file}")

    df = read_pdf_table(pdf_file)
    df = remove_garbage_rows(df)

    df_trans, df_summary = split_summary_rows(df)
    clean_df = merge_continuation_rows(df_trans)
    clean_df = finalize_transactions(clean_df)

    df_summary_final = build_summary_dataframe(df_summary)

    print("Proses pembersihan data selesai. Menulis ke Excel...")
    export_to_excel(clean_df, df_summary_final, output_file)

    print(f"Data successfully written to {output_file}")


if __name__ == "__main__":
    main()