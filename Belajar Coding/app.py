from flask import Flask, render_template, request, send_file, jsonify
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
            return jsonify({"error": "Tidak ada file yang diupload!"}), 400

        file = request.files["file"]

        if file.filename == "":
            return jsonify({"error": "Pilih file terlebih dahulu!"}), 400

        if not allowed_file(file.filename):
            return jsonify({"error": "Format file tidak diizinkan! Hanya PDF dan Excel yang diperbolehkan."}), 400

        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)

        file_name = os.path.splitext(file.filename)[0]
        output_file_path = os.path.join(UPLOAD_FOLDER, f"{file_name}.xlsx")
        output_name = f"{file_name}.xlsx"
        
        file_extension = file.filename.rsplit(".", 1)[1].lower()

        try:
            if file_extension == "pdf":
                process_pdf(file_path)
                output_file = output_file_path
            elif file_extension in ["xlsx", "xls"]:
                with open(file_path, "rb") as f:
                    output_file = process_excel(f)

            response = send_file(output_file,
                                 as_attachment=True,
                                 download_name=output_name,
                                 mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            time.sleep(1)
            if isinstance(output_file, str) and os.path.exists(output_file):
                os.remove(output_file)

            return response
        except Exception as e:
            print(f"Error during file processing: {e}")
            return jsonify({"error": f"Terjadi kesalahan: {e}"}), 500
        finally:
            if os.path.exists(file_path):
                os.remove(file_path)
            print("Proses selesai.")
            
    return render_template("index.html")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080, debug=True)
