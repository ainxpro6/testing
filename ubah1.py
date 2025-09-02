import pdfplumber
import re
import os
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

def load_master_skus(file_path='daftar_sku.txt'):
    """Memuat daftar SKU utuh dari file teks."""
    if not os.path.exists(file_path):
        return set()
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return {line.strip() for line in f if line.strip()}
    except Exception as e:
        print(f"Error saat membaca {file_path}: {e}")
        return set()

def find_matching_sku(partial_sku, master_skus):
    """Mencocokkan dan membersihkan SKU."""
    normalized_sku = partial_sku.replace('Β', 'B').replace('Ο', 'O').replace('Υ', 'Y')
    if master_skus:
        for full_sku in master_skus:
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
    return [line for line in lines if line.strip()] # Hapus baris kosong

def process_data(lines, master_skus):
    """
    Logika baru: Bekerja mundur dari baris 'Default Slot' untuk memastikan akurasi.
    """
    data = []
    processed_indices = set()

    # Iterasi mundur agar modifikasi indeks tidak mengganggu loop
    for i in range(len(lines) - 1, -1, -1):
        if i in processed_indices:
            continue

        line = lines[i].strip()
        
        # 1. TEMUKAN JANGKARNYA
        if "Default Slot" in line:
            qty_match = re.search(r'Default Slot (\d+)$', line)
            if not qty_match:
                continue
            
            qty = int(qty_match.group(1))
            
            # 2. KUMPULKAN BLOK TEKS DI ATAS JANGKAR
            # Kumpulkan 3 baris di atasnya, yang kemungkinan besar berisi semua info
            start_index = max(0, i - 3)
            block_lines = []
            for j in range(start_index, i):
                # Jangan sertakan header atau data yang sudah diproses
                if j not in processed_indices and "Nama Produk" not in lines[j] and "desty" not in lines[j]:
                    block_lines.append(lines[j].strip())
            
            if not block_lines:
                continue

            full_block_text = " ".join(block_lines)
            
            # 3. EKSTRAKSI INFORMASI DARI BLOK
            
            # Ekstrak Varian (jika ada)
            varian = ""
            varian_match = re.search(r'Variant: (.*)', full_block_text)
            if varian_match:
                varian = varian_match.group(1).strip()
                # Hapus varian dari teks untuk mempermudah ekstraksi selanjutnya
                full_block_text = full_block_text.replace(varian_match.group(0), "").strip()
            
            # Ekstrak SKU (kata terakhir yang terlihat seperti SKU)
            # Pola ini mencari kata terakhir yang terdiri dari huruf besar, angka, dan strip.
            # Ia juga bisa menangani SKU yang terpotong menjadi dua kata.
            sku_match = re.search(r'([A-Z0-9\-]{4,})\s*([A-Z0-9ΒΟΥ]+)?$', full_block_text)
            raw_sku = ""
            nama_produk = full_block_text # Default nama produk adalah semua teks
            
            if sku_match:
                part1 = sku_match.group(1) or ""
                part2 = sku_match.group(2) or ""
                raw_sku = part1 + part2
                # Nama produk adalah semua teks sebelum SKU
                nama_produk = full_block_text[:sku_match.start()].strip()
                
            final_sku = find_matching_sku(raw_sku, master_skus)
            
            # 4. SIMPAN DATA & TANDAI BARIS YANG SUDAH DIPROSES
            if nama_produk: # Hanya simpan jika ada nama produk
                data.append({
                    'Nama Produk': nama_produk,
                    'SKU': final_sku,
                    'Varian': varian,
                    'Qty': qty
                })
                # Tandai semua baris dalam blok ini sebagai sudah diproses
                for j in range(start_index, i + 1):
                    processed_indices.add(j)

    # Kembalikan urutan seperti semula (karena kita memprosesnya terbalik)
    data.reverse()
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
    print(f"Data telah disimpan ke {output_file} dengan {len(processed_data)} baris.")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Penggunaan: python ubah1.py <file_pdf>")
        sys.exit(1)

    pdf_file_path = sys.argv[1]
    main(pdf_file_path)