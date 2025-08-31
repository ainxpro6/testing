from flask import Flask, render_template, request, send_file, flash
import os
import time
from ubah1 import main as process_pdf
from ubah2 import main as process_excel

app = Flask(__name__)
app.secret_key = "supersecretkey"

UPLOAD_FOLDER = "/tmp/uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

ALLOWED_EXTENSIONS = {"pdf", "xlsx", "xls"}

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        if "file" not in request.files:
            flash("❌ Tidak ada file yang diupload!", "danger")
            return render_template("index.html")

        file = request.files["file"]

        if file.filename == "":
            flash("⚠️ Pilih file terlebih dahulu!", "warning")
            return render_template("index.html")

        if not allowed_file(file.filename):
            flash("🚫 Format file tidak diizinkan! Hanya PDF dan Excel yang diperbolehkan.", "danger")
            return render_template("index.html")

        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)

        # Ambil nama file tanpa ekstensi
        file_name = os.path.splitext(file.filename)[0]
        
        # Tentukan nama file output sesuai dengan file input (ubah ekstensi menjadi .xlsx)
        output_file = os.path.join(UPLOAD_FOLDER, f"{file_name}.xlsx")

        file_extension = file.filename.rsplit(".", 1)[1].lower()

        try:
            if file_extension == "pdf":
                process_pdf(file_path)  # Memproses file PDF
                output_file = output_file  # Nama file output
            elif file_extension in ["xlsx", "xls"]:
                with open(file_path, "rb") as f:
                    output_file = process_excel(f)  # Memproses file Excel

            # Menambahkan download_name dan mimetype pada send_file
            response = send_file(output_file,
                                 as_attachment=True,
                                 download_name=f"{file_name}.xlsx",  # Nama file output yang sesuai
                                 mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            # Tunggu sebentar sebelum menghapus file
            time.sleep(1)

            # Pastikan hanya menghapus file yang ada di disk, bukan objek BytesIO
            if isinstance(output_file, str):
                if os.path.exists(file_path):
                    os.remove(file_path)  # Menghapus file input (PDF atau Excel asli)
                if os.path.exists(output_file):
                    os.remove(output_file)  # Menghapus file output Excel

            return response
        except Exception as e:
            print(f"Error during file processing: {e}")
            flash(f"Terjadi kesalahan: {e}", "danger")
            return render_template("index.html")
        finally:
            # Jika menggunakan file dalam memori seperti BytesIO, tidak perlu dihapus
            # Pastikan tidak ada file yang tertinggal terkunci
            print("Proses selesai. File output tidak terkunci.")
            
    return render_template("index.html")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080, debug=True)
