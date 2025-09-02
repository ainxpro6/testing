import pdfplumber
import re
import os
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side


def extract_text_from_pdf(pdf_path):
    """
    Ekstrak teks per halaman dari PDF dan gabungkan jadi satu list baris.
    """
    lines = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                lines.extend(page_text.split('\n'))
    return lines


def process_data(lines):
    """
    Proses teks menjadi list data terstruktur: Nama Produk, SKU, Variant, Qty.
    """
    data = []

    for i in range(len(lines)):
        line = lines[i].strip()

        # Deteksi baris produk + SKU
        if re.search(r'[A-Z0-9\-]{4,}$', line) and 'Default Slot' not in line:
            parts = line.rsplit(' ', 1)
            if len(parts) == 2:
                nama_produk, sku = parts[0].strip(), parts[1].strip()

                # Ambil qty dari baris setelahnya
                if i + 1 < len(lines) and "Default Slot" in lines[i + 1]:
                    qty_match = re.search(r'Default Slot (\d+)', lines[i + 1])
                    if qty_match:
                        qty = qty_match.group(1)

                        # Cari varian di 1–3 baris setelahnya
                        varian = ""
                        for j in range(1, 4):
                            if i + j + 1 < len(lines):
                                next_line = lines[i + j + 1]
                                if "Variant:" in next_line:
                                    varian = next_line.split("Variant:")[-1].strip()
                                    break

                        data.append({
                            'Nama Produk': nama_produk,
                            'SKU': sku,
                            'Varian': varian,
                            'Qty': int(qty)
                        })
    return data


def clean_data(data):
    """
    Bersihkan atau modifikasi data jika diperlukan.
    Untuk sementara ini return data langsung.
    """
    return data


def save_to_excel(data, output_file):
    """
    Simpan data ke dalam file Excel dengan format rapi.
    """
    wb = Workbook()
    ws = wb.active

    headers = ["Nama Produk", "Variant", "SKU", "Qty"]
    ws.append(headers)

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for item in data:
        row = [
            str(item.get("Nama Produk", "")),
            str(item.get("Varian", "")),
            str(item.get("SKU", "")),
            str(item.get("Qty", "")),
        ]
        ws.append(row)

    # Format kolom
    ws.column_dimensions['A'].width = 50
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 4

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, horizontal='left', vertical='center')

    # Hapus baris yang kosong
    for row in range(ws.max_row, 1, -1):
        non_empty_cells = [cell for cell in ws[row] if cell.value and str(cell.value).strip()]
        if len(non_empty_cells) <= 1:
            ws.delete_rows(row)

    # Tambahkan border
    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = thin_border

    wb.save(output_file)


def main(file_path):
    """
    Fungsi utama untuk memproses file PDF menjadi file Excel.
    """
    file_name = os.path.splitext(os.path.basename(file_path))[0]
    output_file = os.path.join(os.path.dirname(file_path), f"{file_name}.xlsx")

    text_lines = extract_text_from_pdf(file_path)
    data = process_data(text_lines)
    cleaned_data = clean_data(data)
    save_to_excel(cleaned_data, output_file)

    print(f"Data telah disimpan ke {output_file}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Penggunaan: ubah.py <file_pdf>")
        sys.exit(1)

    pdf_file_path = sys.argv[1]
    main(pdf_file_path)
