import os
from flask import Flask, render_template, request, url_for, send_from_directory
from werkzeug.utils import secure_filename

from statmaster_logic import analyze_pdf, analyze_two_pdfs_comparison


# Configurazione base Flask
app = Flask(__name__)

# Cartelle per PDF caricati e report
BASE_DIR = os.path.dirname(__file__)
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
REPORT_FOLDER = os.path.join(BASE_DIR, "reports")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(REPORT_FOLDER, exist_ok=True)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["REPORT_FOLDER"] = REPORT_FOLDER

ALLOWED_EXTENSIONS = {"pdf"}


def allowed_file(filename: str) -> bool:
    """Controlla se il file ha un'estensione ammessa (solo .pdf)."""
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


# -------------------------
#   ROUTE: SINGOLO REPORT
# -------------------------
@app.route("/", methods=["GET", "POST"])
def index():
    """Pagina principale: analisi di un solo dipendente."""
    if request.method == "POST":
        employee_name = request.form.get("employee_name", "").strip()
        weekly_hours_str = request.form.get("weekly_hours", "").strip()

        if not employee_name:
            return render_template("index.html", error="Inserisci il nome della dipendente.")

        if not weekly_hours_str:
            return render_template("index.html", error="Inserisci le ore settimanali di contratto.")

        try:
            weekly_hours = float(weekly_hours_str.replace(",", "."))
        except ValueError:
            return render_template(
                "index.html",
                error="Le ore settimanali devono essere un numero (es. 20 o 15.5)."
            )

        if "pdf_file" not in request.files:
            return render_template("index.html", error="Nessun file ricevuto.")

        file = request.files["pdf_file"]

        if file.filename == "":
            return render_template("index.html", error="Nessun file selezionato.")

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            save_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            file.save(save_path)

            try:
                report_filename, summary = analyze_pdf(
                    pdf_path=save_path,
                    employee_name=employee_name,
                    weekly_hours=weekly_hours,
                    report_folder=app.config["REPORT_FOLDER"],
                )
            except Exception as e:
                return render_template(
                    "index.html",
                    error=f"Errore durante l'analisi del PDF: {e}"
                )

            download_url = url_for("download_report", filename=report_filename)

            return render_template(
                "success.html",
                employee_name=employee_name,
                weekly_hours=weekly_hours,
                summary=summary,
                download_url=download_url,
                report_filename=report_filename,
            )

        return render_template("index.html", error="Per favore carica un file PDF (.pdf) valido.")

    # GET
    return render_template("index.html")


# -------------------------
#   ROUTE: CONFRONTO 2 DIPENDENTI
# -------------------------
@app.route("/compare", methods=["GET", "POST"])
def compare():
    """Pagina per confrontare due dipendenti con un unico report PDF di confronto."""
    if request.method == "POST":
        # Dipendente 1
        name1 = request.form.get("employee1_name", "").strip()
        h1_str = request.form.get("weekly_hours1", "").strip()

        # Dipendente 2
        name2 = request.form.get("employee2_name", "").strip()
        h2_str = request.form.get("weekly_hours2", "").strip()

        if not name1 or not name2:
            return render_template("compare.html", error="Inserisci il nome di entrambi i dipendenti.")

        if not h1_str or not h2_str:
            return render_template(
                "compare.html",
                error="Inserisci le ore settimanali di contratto per entrambi i dipendenti."
            )

        try:
            weekly_hours1 = float(h1_str.replace(",", "."))
            weekly_hours2 = float(h2_str.replace(",", "."))
        except ValueError:
            return render_template(
                "compare.html",
                error="Le ore settimanali devono essere numeri (es. 20 o 15.5)."
            )

        files = request.files
        if "pdf_file1" not in files or "pdf_file2" not in files:
            return render_template("compare.html", error="Carica i PDF per entrambi i dipendenti.")

        f1 = files["pdf_file1"]
        f2 = files["pdf_file2"]

        if f1.filename == "" or f2.filename == "":
            return render_template("compare.html", error="Seleziona entrambi i file PDF.")

        if not (allowed_file(f1.filename) and allowed_file(f2.filename)):
            return render_template("compare.html", error="Carica solo file PDF (.pdf).")

        filename1 = "1_" + secure_filename(f1.filename)
        filename2 = "2_" + secure_filename(f2.filename)
        path1 = os.path.join(app.config["UPLOAD_FOLDER"], filename1)
        path2 = os.path.join(app.config["UPLOAD_FOLDER"], filename2)
        f1.save(path1)
        f2.save(path2)

        try:
            comparison_filename, summary1, summary2 = analyze_two_pdfs_comparison(
                pdf1_path=path1,
                employee1_name=name1,
                weekly_hours1=weekly_hours1,
                pdf2_path=path2,
                employee2_name=name2,
                weekly_hours2=weekly_hours2,
                report_folder=app.config["REPORT_FOLDER"],
            )
        except Exception as e:
            return render_template(
                "compare.html",
                error=f"Errore durante l'analisi dei PDF: {e}"
            )

        comparison_download_url = url_for("download_report", filename=comparison_filename)

        return render_template(
            "compare_result.html",
            employee1_name=name1,
            employee2_name=name2,
            weekly_hours1=weekly_hours1,
            weekly_hours2=weekly_hours2,
            summary1=summary1,
            summary2=summary2,
            comparison_download_url=comparison_download_url,
            comparison_filename=comparison_filename,
        )

    # GET
    return render_template("compare.html")



# -------------------------
#   DOWNLOAD REPORT
# -------------------------
@app.route("/download/<path:filename>")
def download_report(filename):
    """Permette di scaricare il report PDF generato da StatMaster."""
    return send_from_directory(app.config["REPORT_FOLDER"], filename, as_attachment=True)


if __name__ == "__main__":
    # Avvia il server in locale su http://127.0.0.1:5000
    app.run(debug=False)
