import os
from flask import Flask, render_template, request, url_for, send_from_directory
from werkzeug.utils import secure_filename

# Usiamo SOLO le funzioni di alto livello del modulo di logica
from statmaster_logic import (
    analyze_pdf,
    analyze_two_pdfs_comparison,
)

# -------------------------------------------------
# CONFIGURAZIONE BASE FLASK
# -------------------------------------------------
app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
REPORT_FOLDER = os.path.join(BASE_DIR, "reports")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(REPORT_FOLDER, exist_ok=True)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["REPORT_FOLDER"] = REPORT_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # max 16 MB per file


# -------------------------------------------------
# ROUTE: HOME / SINGOLO DIPENDENTE
# -------------------------------------------------
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        employee_name = request.form.get("employee_name", "").strip()
        weekly_hours_str = request.form.get("weekly_hours", "").strip()
        file = request.files.get("pdf_file")

        # Controlli base sui campi del form
        if not employee_name or not weekly_hours_str or not file:
            return render_template(
                "index.html",
                error="Compila tutti i campi e carica un PDF.",
                employee_name=employee_name,
                weekly_hours=weekly_hours_str,
            )

        try:
            weekly_hours = float(weekly_hours_str.replace(",", "."))
        except ValueError:
            return render_template(
                "index.html",
                error="Le ore settimanali devono essere un numero (es. 20).",
                employee_name=employee_name,
                weekly_hours=weekly_hours_str,
            )

        filename = secure_filename(file.filename or "")
        if not filename.lower().endswith(".pdf"):
            return render_template(
                "index.html",
                error="Carica solo file in formato PDF (.pdf).",
                employee_name=employee_name,
                weekly_hours=weekly_hours_str,
            )

        # Salvo il PDF caricato
        upload_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        file.save(upload_path)

        try:
            # Funzione di alto livello: fa parsing, calcoli e genera il report PDF
            report_filename, summary = analyze_pdf(
                pdf_path=upload_path,
                employee_name=employee_name,
                weekly_hours=weekly_hours,
                report_folder=app.config["REPORT_FOLDER"],
            )




        except ValueError as e:
            # Errori “previsti” (PDF non leggibile, nessuna timbratura, ecc.)
            return render_template(
                "index.html",
                error=f"Errore durante l'analisi del PDF: {e}",
                employee_name=employee_name,
                weekly_hours=weekly_hours_str,
            )
        except Exception as e:
            # Qualsiasi altro errore: messaggio generico ma leggibile
            return render_template(
                "index.html",
                error=f"Errore imprevisto durante l'elaborazione del PDF: {e}",
                employee_name=employee_name,
                weekly_hours=weekly_hours_str,
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

    # GET: mostro la pagina vuota
    return render_template("index.html")


# -------------------------------------------------
# ROUTE: CONFRONTO 2 DIPENDENTI
# -------------------------------------------------
@app.route("/compare", methods=["GET", "POST"])
def compare():
    """
    Pagina per confrontare due dipendenti con un unico report PDF di confronto.
    """
    if request.method == "POST":
        # Campi dipendente 1
        employee1_name = request.form.get("employee1_name", "").strip()
        weekly_hours1_str = request.form.get("weekly_hours1", "").strip()
        file1 = request.files.get("pdf_file1")

        # Campi dipendente 2
        employee2_name = request.form.get("employee2_name", "").strip()
        weekly_hours2_str = request.form.get("weekly_hours2", "").strip()
        file2 = request.files.get("pdf_file2")

        # Controlli base
        if (
            not employee1_name
            or not weekly_hours1_str
            or not file1
            or not employee2_name
            or not weekly_hours2_str
            or not file2
        ):
            return render_template(
                "compare.html",
                error="Compila tutti i campi e carica i due PDF.",
            )

        try:
            weekly_hours1 = float(weekly_hours1_str.replace(",", "."))
            weekly_hours2 = float(weekly_hours2_str.replace(",", "."))
        except ValueError:
            return render_template(
                "compare.html",
                error="Le ore settimanali devono essere numeriche (es. 20 e 15).",
            )

        filename1 = secure_filename(file1.filename or "")
        filename2 = secure_filename(file2.filename or "")

        if not filename1.lower().endswith(".pdf") or not filename2.lower().endswith(".pdf"):
            return render_template(
                "compare.html",
                error="Carica solo file in formato PDF (.pdf) per entrambi i dipendenti.",
            )

        # Salvo i due PDF
        path1 = os.path.join(app.config["UPLOAD_FOLDER"], "1_" + filename1)
        path2 = os.path.join(app.config["UPLOAD_FOLDER"], "2_" + filename2)
        file1.save(path1)
        file2.save(path2)

        try:
            # Funzione di alto livello che genera il report di confronto
            comparison_filename, summary1, summary2 = analyze_two_pdfs_comparison(
                pdf1_path=path1,
                employee1_name=employee1_name,
                weekly_hours1=weekly_hours1,
                pdf2_path=path2,
                employee2_name=employee2_name,
                weekly_hours2=weekly_hours2,
                report_folder=app.config["REPORT_FOLDER"],
            )

        except ValueError as e:
            return render_template(
                "compare.html",
                error=f"Errore durante l'analisi dei PDF: {e}",
            )
        except Exception as e:
            return render_template(
                "compare.html",
                error=f"Errore imprevisto durante l'elaborazione dei PDF: {e}",
            )

        download_url = url_for("download_report", filename=comparison_filename)

        return render_template(
            "compare_result.html",
            employee1_name=employee1_name,
            employee2_name=employee2_name,
            weekly_hours1=weekly_hours1,
            weekly_hours2=weekly_hours2,
            summary1=summary1,
            summary2=summary2,
            download_url=download_url,
            report_filename=comparison_filename,
        )

    # GET
    return render_template("compare.html")


# -------------------------------------------------
# ROUTE: DOWNLOAD REPORT
# -------------------------------------------------
@app.route("/download/<path:filename>")
def download_report(filename):
    """Permette di scaricare il report PDF generato da StatMaster."""
    return send_from_directory(
        app.config["REPORT_FOLDER"],
        filename,
        as_attachment=True,
    )


# -------------------------------------------------
# AVVIO LOCALE (sviluppo)
# -------------------------------------------------
if __name__ == "__main__":
    # Avvia il server in locale su http://127.0.0.1:5000
    app.run(debug=False)
