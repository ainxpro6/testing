import pdfplumber
import pandas as pd
import re
import os
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side


def extract_and_process_pdf(pdf_path):

    print("Memulai metode 'Grid' untuk ekstraksi data mentah...")

    KOLOM_BOUNDARIES = [
        (0, 350), (350, 470), (470, 540), (540, 595)
    ]

    all_rows_structured = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            print(f"Memproses Halaman {page_num + 1}...")
            
            if page.width != 595:
                KOLOM_BOUNDARIES[3] = (540, page.width)

            h_lines = sorted(list(set([edge['top'] for edge in page.horizontal_edges] + [0, page.height])))

            for i in range(len(h_lines) - 1):
                top = h_lines[i]
                bottom = h_lines[i+1]
                
                row_data = []
                for x0, x1 in KOLOM_BOUNDARIES:
                    cell_crop = page.crop((x0, top, x1, bottom))
                    text = cell_crop.extract_text(x_tolerance=2, y_tolerance=2)
                    row_data.append(text.strip() if text else '')
                
                if any(row_data):
                    all_rows_structured.append(row_data)

    if not all_rows_structured:
        raise Exception("Tidak ada data yang bisa diekstrak.")

    df_raw = pd.DataFrame(all_rows_structured, columns=['Nama Produk', 'SKU', 'Slot', 'Qty'])
    df_raw = df_raw[~df_raw['Nama Produk'].str.contains("Nama Produk", na=False)].reset_index(drop=True)
    
    return df_raw


def clean_data(df_raw):

    print("Menerapkan aturan pembersihan pada data...")
    processed_data = []
    for index, row in df_raw.iterrows():
        nama_produk_raw = str(row.get('Nama Produk', ''))
        sku_raw = str(row.get('SKU', ''))
        qty = str(row.get('Qty', ''))

        if qty and qty.isdigit():
            
            if 'Buyer Notes:' in nama_produk_raw:
                nama_produk_raw = nama_produk_raw.split('Buyer Notes:')[0]

            sku_joined = sku_raw.replace('\n', '')
            sku_cleaned = re.sub('defa', '', sku_joined, flags=re.IGNORECASE).strip()
            sku_final = re.sub(r'^.\s', '', sku_cleaned)
            
            nama_produk_clean = ' '.join(nama_produk_raw.replace('\n', ' ').split())
            varian = ''
            match = re.search(r'(variant:|riant:)(.*)', nama_produk_clean, re.IGNORECASE)
            if match:
                nama_produk_clean = nama_produk_clean.split(match.group(0))[0].strip()
                varian = match.group(2).strip()
            
            processed_data.append({
                'Nama Produk': nama_produk_clean,
                'Varian': varian,
                'SKU': sku_final,
                'Qty': int(qty)
            })
            
    return processed_data


def save_to_excel(data, output_file):

    wb = Workbook()
    ws = wb.active

    headers = ["Nama Produk", "Varian", "SKU", "Qty"]
    ws.append(headers)

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for item in data:
        row = [
            item.get("Nama Produk", ""),
            item.get("Varian", ""),
            item.get("SKU", ""),
            item.get("Qty", ""),
        ]
        ws.append(row)

    ws.column_dimensions['A'].width = 52
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 4

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=3):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, horizontal='left', vertical='center')
            
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=4):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')

    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = thin_border

    wb.save(output_file)


def main(file_path):
    
    file_name = os.path.splitext(os.path.basename(file_path))[0]
    output_file = os.path.join(os.path.dirname(file_path), f"{file_name}.xlsx")

    raw_data_df = extract_and_process_pdf(file_path)
    cleaned_data = clean_data(raw_data_df)
    save_to_excel(cleaned_data, output_file)

    print(f"\nProses Selesai!")
    print(f"Data telah disimpan ke {output_file}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Penggunaan: python ubah.py <file_pdf>")
        sys.exit(1)

    pdf_file_path = sys.argv[1]
    main(pdf_file_path)
    
