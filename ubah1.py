import pdfplumber
import re
import os
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

def load_master_skus(file_path='daftar_sku.txt'):
    """
    Memuat daftar SKU utuh dari file teks.
    Mengembalikan sebuah set untuk pencarian yang lebih cepat.
    """
    if not os.path.exists(file_path):
        print(f"Peringatan: File '{file_path}' tidak ditemukan. Pencocokan SKU tidak akan dilakukan.")
        return None
    with open(file_path, 'r') as f:
        # Menghapus spasi putih dan baris kosong
        return {line.strip() for line in f if line.strip()}

def find_matching_sku(partial_sku, master_skus):
    """
    Mencari SKU utuh di master_skus yang cocok dengan SKU terpotong.
    """
    # Menghapus karakter non-alfanumerik aneh seperti 'Β' (Beta Yunani)
    normalized_sku = partial_sku.replace('Β', 'B')

    for full_sku in master_skus:
        if full_sku.startswith(normalized_sku):
            return full_sku
    return partial_sku # Jika tidak ada yang cocok, kembalikan SKU asli

def extract_text_from_pdf(pdf_path):
    """
    Ekstrak teks per halaman dari PDF.
    """
    lines = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                lines.extend(page_text.split('\n'))
    return lines

def process_data(lines, master_skus):
    """
    Proses teks menjadi data terstruktur dengan pencocokan SKU.
    """
    data = []
    processed_indices = set()

    for i in range(len(lines)):
        if i in processed_indices:
            continue

        line = lines[i].strip()
        
        # Heuristik untuk mendeteksi baris nama produk dan bagian pertama SKU
        match = re.search(r'^(.*[a-zA-Z].*)(\s+)([A-Z0-9\-]{4,})$', line)
        
        if match and 'Default Slot' not in line and 'Variant:' not in line:
            nama_produk, sku_part1 = match.group(1).strip(), match.group(3).strip()
            qty = None
            
            # --- Logika Baru untuk SKU terpotong di 2 baris ---
            # Cek apakah baris berikutnya adalah kelanjutan SKU
            if i + 2 < len(lines) and "Default Slot" in lines[i+2]:
                next_line_parts = lines[i+1].strip().split()
                # Jika baris berikutnya hanya berisi 1 bagian (calon bagian kedua SKU)
                if len(next_line_parts) == 1:
                    sku_part2 = next_line_parts[0]
                    combined_sku = sku_part1 + sku_part2
                    
                    # Cocokkan dengan daftar SKU utuh
                    final_sku = find_matching_sku(combined_sku, master_skus) if master_skus else combined_sku
                    
                    qty_match = re.search(r'Default Slot (\d+)', lines[i+2])
                    if qty_match:
                        qty = qty_match.group(1)
                        processed_indices.update([i, i + 1, i + 2])

            # --- Logika Lama untuk SKU di 1 baris ---
            if qty is None and i + 1 < len(lines) and "Default Slot" in lines[i+1]:
                final_sku = find_matching_sku(sku_part1, master_skus) if master_skus else sku_part1
                qty_match = re.search(r'Default Slot (\d+)', lines[i+1])
                if qty_match:
                    qty = qty_match.group(1)
                    processed_indices.update([i, i + 1])

            if qty:
                varian = ""
                # Cari Varian di baris berikutnya
                if i + 1 < len(lines) and "Variant:" in lines[i+1]:
                    varian = lines[i+1].split("Variant:")[-1].strip()
                # Jika varian ada di baris yang sama dengan produk (jarang terjadi)
                elif "Variant:" in nama_produk:
                     nama_produk, varian = nama_produk.split("Variant:")
                     nama_produk = nama_produk.strip()
                     varian = varian.strip()


                data.append({
                    'Nama Produk': nama_produk,
                    'SKU': final_sku,
                    'Varian': varian,
                    'Qty': int(qty)
                })

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
            item.get("Nama Produk", ""),
            item.get("Varian", ""),
            item.get("SKU", ""),
            item.get("Qty", ""),
        ]
        ws.append(row)

    ws.column_dimensions['A'].width = 60
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 25
    ws.column_dimensions['D'].width = 5

    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    for row in ws.iter_rows(min_row=1):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, horizontal='left', vertical='center')
            cell.border = thin_border

    wb.save(output_file)

def main(file_path):
    """
    Fungsi utama untuk memproses file PDF menjadi file Excel.
    """
    file_name = os.path.splitext(os.path.basename(file_path))[0]
    output_file = os.path.join(os.path.dirname(file_path), f"{file_name}.xlsx")

    # Muat daftar SKU utuh
    master_skus = load_master_skus()
    
    text_lines = extract_text_from_pdf(file_path)
    processed_data = process_data(text_lines, master_skus)
    save_to_excel(processed_data, output_file)

    print(f"Data telah disimpan ke {output_file}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Penggunaan: python ubah1.py <file_pdf>")
        sys.exit(1)

    pdf_file_path = sys.argv[1]
    main(pdf_file_path)