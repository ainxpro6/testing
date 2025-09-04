from flask import Flask, render_template, request, send_file, jsonify
import os
import time
from werkzeug.utils import secure_filename
from ubah1 import main as process_pdf
from ubah2 import main as process_excel

app = Flask(__name__)

ALLOWED_EXTENSIONS = {"pdf", "xlsx", "xls"}

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        if "file" not in request.files:
            return jsonify({"error": "Tidak ada file yang diupload!"}), 400

        file = request.files["file"]

        if file.filename == "":
            return jsonify({"error": "Pilih file terlebih dahulu!"}), 400

        if not allowed_file(file.filename):
            return jsonify({"error": "Format file tidak diizinkan! Hanya PDF dan Excel yang diperbolehkan."}), 400

        try:
            filename = secure_filename(file.filename)
            file_name_without_ext = os.path.splitext(filename)[0]
            output_filename = f"{file_name_without_ext}.xlsx"
            
            file_extension = filename.rsplit(".", 1)[1].lower()

            output_to_send = None
            if file_extension == "pdf":
                output_to_send = process_pdf(file) 
            elif file_extension in ["xlsx", "xls"]:
                output_to_send = process_excel(file)

            if output_to_send is None:
                raise ValueError("Gagal memproses file.")

            response = send_file(
                output_to_send,
                as_attachment=True,
                download_name=output_filename,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            response.headers["Access-Control-Expose-Headers"] = "Content-Disposition"
            
            return response

        except Exception as e:
            print(f"Error during file processing: {repr(e)}")
            return jsonify({"error": f"Terjadi kesalahan: {str(e)}"}), 500
            
    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)