import pdfplumber
import re
import os
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

def load_master_skus(file_path='daftar_sku.txt'):
    """Memuat daftar SKU utuh dari file teks."""
    if not os.path.exists(file_path):
        print(f"Peringatan: File '{file_path}' tidak ditemukan.")
        return set()
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return {line.strip() for line in f if line.strip()}
    except Exception as e:
        print(f"Error saat membaca {file_path}: {e}")
        return set()

def find_matching_sku(partial_sku, master_skus):
    """Mencocokkan dan membersihkan SKU."""
    # Normalisasi karakter yang sering salah baca
    normalized_sku = partial_sku.replace('Β', 'B').replace('Ο', 'O').replace('Υ', 'Y')
    if master_skus:
        for full_sku in master_skus:
            # Menggunakan startswith untuk menangani SKU terpotong
            if full_sku.startswith(normalized_sku):
                return full_sku
    return normalized_sku

def extract_text_from_pdf(pdf_path):
    """Mengekstrak teks mentah dari PDF."""
    lines = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text(x_tolerance=2, y_tolerance=2)
                if page_text:
                    lines.extend(page_text.split('\n'))
    except Exception as e:
        print(f"Gagal membaca PDF: {e}")
    return lines

def process_data(lines, master_skus):
    """
    Logika baru: Mengidentifikasi blok data dan membedahnya.
    Ini jauh lebih andal daripada metode per baris.
    """
    data = []
    i = 0
    while i < len(lines):
        line = lines[i].strip()

        # Lewati baris yang tidak relevan di awal
        if not line or line.startswith("desty") or line.startswith("Jumlah") or line.startswith("Tanggal") or line.startswith("Dicetak") or line.startswith("Picking List") or line.startswith("Halaman") or line.startswith("Nama Produk"):
            i += 1
            continue

        # Awal dari sebuah blok data produk
        block_lines = [line]
        end_of_block_index = i

        # Cari akhir dari blok (baris dengan "Default Slot")
        for j in range(i + 1, min(i + 5, len(lines))):
            next_line = lines[j].strip()
            block_lines.append(next_line)
            if "Default Slot" in next_line:
                end_of_block_index = j
                break
        
        # Jika akhir blok ditemukan, proses blok tersebut
        if end_of_block_index > i:
            full_block_text = " ".join(block_lines)
            
            # Ekstrak Qty
            qty_match = re.search(r'Default Slot (\d+)$', full_block_text)
            qty = qty_match.group(1) if qty_match else '0'

            # Ekstrak Varian (jika ada)
            varian_match = re.search(r'Variant: (.*?)(?=\s[A-Z0-9\-]{4,}|$)', full_block_text)
            varian = varian_match.group(1).strip() if varian_match else ''

            # Ekstrak Nama Produk dan SKU
            # Menghapus varian dan qty dari teks blok untuk menyisakan Nama Produk & SKU
            text_for_sku = full_block_text.replace(f"Default Slot {qty}", "")
            if varian:
                text_for_sku = text_for_sku.replace(f"Variant: {varian}", "")
            
            # Pola untuk menemukan SKU di akhir teks
            sku_match = re.search(r'(\s+)([A-Z0-9\-ΒΟΥ]+(?:\s[A-Z0-9\-ΒΟΥ]+)?)$', text_for_sku.strip())
            
            if sku_match:
                # SKU adalah bagian terakhir, sisanya adalah Nama Produk
                raw_sku = sku_match.group(2).replace(" ", "")
                nama_produk = text_for_sku[:sku_match.start()].strip()
                final_sku = find_matching_sku(raw_sku, master_skus)

                # Tambahkan data yang berhasil diekstrak
                data.append({
                    'Nama Produk': nama_produk,
                    'SKU': final_sku,
                    'Varian': varian,
                    'Qty': int(qty)
                })

            # Lanjutkan loop dari setelah blok ini
            i = end_of_block_index + 1
        else:
            # Jika tidak ditemukan akhir blok, lanjutkan ke baris berikutnya
            i += 1
            
    return data

def save_to_excel(data, output_file):
    """Menyimpan data ke file Excel."""
    wb = Workbook()
    ws = wb.active
    headers = ["Nama Produk", "Variant", "SKU", "Qty"]
    ws.append(headers)

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for item in data:
        row = [ item.get(h, "") for h in ["Nama Produk", "Varian", "SKU", "Qty"] ]
        ws.append(row)

    ws.column_dimensions['A'].width = 60
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 30
    ws.column_dimensions['D'].width = 8

    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    for row in ws.iter_rows(min_row=1):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, horizontal='left', vertical='center')
            cell.border = thin_border

    wb.save(output_file)

def main(file_path):
    """Fungsi utama."""
    file_name = os.path.splitext(os.path.basename(file_path))[0]
    output_file = os.path.join(os.path.dirname(file_path), f"{file_name}.xlsx")

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