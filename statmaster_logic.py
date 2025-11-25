
import os
import re
import calendar
from datetime import datetime, time, timedelta
from typing import Tuple, Dict

import pdfplumber
import pandas as pd

import matplotlib
matplotlib.use("Agg")          # <--- AGGIUNGI QUESTE DUE RIGHE

import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages



def extract_text_from_pdf(pdf_path: str) -> str:
    """Estrae il testo da tutte le pagine del PDF."""
    texts = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ''
            texts.append(page_text)
    return '\n'.join(texts)


def parse_pdf_to_dataframe(pdf_path: str) -> pd.DataFrame:
    """
    Parsing specifico per il PDF 'Calendario Periodico Lavori'.

    - La data è su una riga a parte nel formato dd/mm/yyyy (es. 02/01/2025)
    - Le righe dei turni hanno orari nel formato: 5, 30 6, 30 ...
      (virgola al posto dei due punti).

    Per ogni riga con orari:
        start = primo orario (es. 5:30)
        end   = secondo orario (es. 6:30)
    Le ore vengono ricalcolate da noi, ignorando la colonna "Totale".
    """
    text = extract_text_from_pdf(pdf_path)

    lines = text.splitlines()
    current_date = None
    records = []

    # data tipo 02/01/2025
    date_re = re.compile(r"(\d{1,2}/\d{1,2}/\d{4})")

    # orari tipo "5, 30 6, 30" oppure "4, 45 8, 30"
    time_re = re.compile(r"\b(\d{1,2}),\s*(\d{1,2})\s+(\d{1,2}),\s*(\d{1,2})\b")

    for line in lines:
        # Se troviamo una data, la impostiamo come data corrente
        mdate = date_re.search(line)
        if mdate:
            try:
                current_date = datetime.strptime(mdate.group(1), "%d/%m/%Y").date()
            except ValueError:
                current_date = None
            continue

        # Senza data non ha senso leggere orari
        if current_date is None:
            continue

        # Cerchiamo una coppia di orari nella riga
        mt = time_re.search(line)
        if not mt:
            continue

        sh, sm, eh, em = [int(x) for x in mt.groups()]

        # Controllino veloce su valori impossibili
        if not (0 <= sh < 24 and 0 <= eh < 24 and 0 <= sm < 60 and 0 <= em < 60):
            continue

        start_dt = datetime.combine(current_date, time(sh, sm))
        end_dt = datetime.combine(current_date, time(eh, em))

        # Se l'uscita è "prima" dell'entrata, vuol dire che è passata la mezzanotte
        if end_dt < start_dt:
            end_dt += timedelta(days=1)

        hours = (end_dt - start_dt).total_seconds() / 3600.0

        records.append({
            "date": current_date,
            "start": start_dt,
            "end": end_dt,
            "hours": hours,
        })

    if not records:
        raise ValueError(
            "Nessuna timbratura riconosciuta nel PDF. "
            "Il parser è stato cucito sul formato 'Calendario Periodico Lavori', "
            "ma non ha trovato righe con data + orari."
        )

    df = pd.DataFrame(records)
    return df




def compute_monthly_stats(df: pd.DataFrame, weekly_hours: float) -> Tuple[Dict, pd.DataFrame]:
    """
    Calcola gli indicatori mensili e un riepilogo complessivo stile StatMaster.

    Ritorna:
        summary: dict con periodo, totali e medie
        monthly_df: DataFrame con una riga per mese
    """
    # Aggregazione per giorno
    daily = df.groupby("date")["hours"].sum().reset_index()

    # ⚠️ IMPORTANTE: convertiamo la colonna 'date' nel formato datetime di pandas
    daily["date"] = pd.to_datetime(daily["date"])

    # Anno e mese per ogni giorno
    daily["year"] = daily["date"].dt.year
    daily["month"] = daily["date"].dt.month

    # Aggregazione per mese
    monthly_rows = []
    for (year, month), group in daily.groupby(["year", "month"]):
        hours_worked = group["hours"].sum()
        days_worked = group["date"].nunique()
        days_in_month = calendar.monthrange(year, month)[1]
        theoretical_hours = weekly_hours * (days_in_month / 7.0)
        overtime = hours_worked - theoretical_hours
        avg_hours_per_day = hours_worked / days_worked if days_worked > 0 else 0.0

        monthly_rows.append({
            "year": year,
            "month": month,
            "month_label": f"{month:02d}/{year}",
            "days_worked": days_worked,
            "hours_worked": hours_worked,
            "theoretical_hours": theoretical_hours,
            "overtime": overtime,
            "avg_hours_per_day": avg_hours_per_day,
            "days_in_month": days_in_month,
        })

    monthly_df = pd.DataFrame(monthly_rows).sort_values(["year", "month"]).reset_index(drop=True)

    total_hours_worked = monthly_df["hours_worked"].sum()
    total_theoretical_hours = monthly_df["theoretical_hours"].sum()
    total_overtime = total_hours_worked - total_theoretical_hours
    avg_monthly_hours = total_hours_worked / len(monthly_df) if len(monthly_df) > 0 else 0.0

    period_start = daily["date"].min()
    period_end = daily["date"].max()
    period_label = f"{period_start.strftime('%d/%m/%Y')} - {period_end.strftime('%d/%m/%Y')}"

    summary = {
        "period_start": period_start,
        "period_end": period_end,
        "period_label": period_label,
        "total_hours_worked": total_hours_worked,
        "total_theoretical_hours": total_theoretical_hours,
        "total_overtime": total_overtime,
        "avg_monthly_hours": avg_monthly_hours,
        "monthly_count": len(monthly_df),
    }

    return summary, monthly_df



def _page1_overview(pdf: PdfPages, employee_name: str, weekly_hours: float, summary: Dict):
    """Pagina 1: riepilogo sintetico."""
    fig = plt.figure(figsize=(8.27, 11.69))  # A4
    plt.axis('off')

    lines = []
    lines.append(f'Report StatMaster - {employee_name}')
    lines.append('')
    lines.append(f'Periodo analizzato: {summary["period_label"]}')
    lines.append(f'Ore settimanali da contratto: {weekly_hours:.2f} h')
    lines.append('')
    lines.append(f'Totale ore lavorate: {summary["total_hours_worked"]:.2f} h')
    lines.append(f'Totale ore teoriche: {summary["total_theoretical_hours"]:.2f} h')
    lines.append(f'Straordinario netto complessivo: {summary["total_overtime"]:.2f} h')
    lines.append('')
    lines.append(f'Media ore lavorate al mese: {summary["avg_monthly_hours"]:.2f} h')
    lines.append('')
    lines.append('Nota: le ore teoriche sono calcolate come ore_settimanali * (giorni_del_mese / 7).')

    y = 0.9
    for line in lines:
        plt.text(0.1, y, line, fontsize=12, transform=plt.gcf().transFigure)
        y -= 0.04

    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)


def _page2_table(pdf: PdfPages, monthly_df):
    """Pagina 2: tabella mensile con indicatori principali."""
    fig = plt.figure(figsize=(11.69, 8.27))  # A4 orizzontale
    ax = fig.add_subplot(111)
    ax.axis('off')

    # Costruiamo i dati per la tabella
    columns = [
        'Mese',
        'Giorni lavorati',
        'Ore lavorate',
        'Ore teoriche',
        'Straordinario netto',
        'Media ore/giorno lavorato',
    ]

    table_data = []
    for _, row in monthly_df.iterrows():
        mese = f"{row['month']:02d}/{row['year']}"
        table_data.append([
            mese,
            int(row['days_worked']),
            f"{row['hours_worked']:.2f}",
            f"{row['theoretical_hours']:.2f}",
            f"{row['overtime']:.2f}",
            f"{row['avg_hours_per_day']:.2f}",
        ])

    table = ax.table(
        cellText=table_data,
        colLabels=columns,
        loc='center',
    )
    table.auto_set_font_size(False)
    table.set_fontsize(8)
    table.scale(1, 1.5)

    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)


def _page3_bars_hours(pdf: PdfPages, monthly_df):
    """Pagina 3: grafico barre ore lavorate vs teoriche."""
    fig = plt.figure(figsize=(11.69, 8.27))
    ax = fig.add_subplot(111)

    labels = [f"{int(m)}/{int(y)}" for y, m in zip(monthly_df['year'], monthly_df['month'])]
    x = range(len(labels))

    ax.bar(x, monthly_df['hours_worked'], width=0.4, label='Ore lavorate')
    ax.bar([i + 0.4 for i in x], monthly_df['theoretical_hours'], width=0.4, label='Ore teoriche')

    ax.set_xticks([i + 0.2 for i in x])
    ax.set_xticklabels(labels, rotation=45)
    ax.set_ylabel('Ore')
    ax.set_title('Ore lavorate vs ore teoriche per mese')
    ax.legend()

    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)


def _page4_bars_overtime(pdf: PdfPages, monthly_df):
    """Pagina 4: grafico barre straordinario netto per mese."""
    fig = plt.figure(figsize=(11.69, 8.27))
    ax = fig.add_subplot(111)

    labels = [f"{int(m)}/{int(y)}" for y, m in zip(monthly_df['year'], monthly_df['month'])]
    x = range(len(labels))

    ax.bar(x, monthly_df['overtime'])

    ax.set_xticks(list(x))
    ax.set_xticklabels(labels, rotation=45)
    ax.set_ylabel('Ore')
    ax.set_title('Straordinario netto per mese')

    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)


def generate_pdf_report(output_path: str, employee_name: str, weekly_hours: float, summary: Dict, monthly_df):
    """Genera il PDF multipagina stile StatMaster usando matplotlib + PdfPages."""
    with PdfPages(output_path) as pdf:
        _page1_overview(pdf, employee_name, weekly_hours, summary)
        _page2_table(pdf, monthly_df)
        _page3_bars_hours(pdf, monthly_df)
        _page4_bars_overtime(pdf, monthly_df)


def analyze_pdf(pdf_path: str, employee_name: str, weekly_hours: float, report_folder: str):
    """
    Funzione principale chiamata da Flask.

    1. Legge il PDF delle timbrature.
    2. Calcola le statistiche mensili e il riepilogo.
    3. Genera il report PDF multipagina.
    4. Restituisce il nome del file di report e il riepilogo.
    """
    df = parse_pdf_to_dataframe(pdf_path)
    summary, monthly_df = compute_monthly_stats(df, weekly_hours)

    # Costruiamo un nome file pulito per il report
    safe_name = re.sub(r'[^a-zA-Z0-9_-]+', '_', employee_name.strip()) or 'dipendente'
    period_tag = summary['period_label'].replace(' ', '').replace('/', '-').replace(':', '-')
    report_filename = f'Report_{safe_name}_{period_tag}.pdf'
    output_path = os.path.join(report_folder, report_filename)

    generate_pdf_report(output_path, employee_name, weekly_hours, summary, monthly_df)

    return report_filename, summary

# ================================================================
#  FUNZIONI DI SUPPORTO PER IL REPORT DI CONFRONTO
# ================================================================

def _format_hours_hm(hours: float) -> str:
    """Converte ore decimali in stringa 'X h YY min', con eventuale segno."""
    total_minutes = int(round(hours * 60))
    sign = "-" if total_minutes < 0 else ""
    total_minutes = abs(total_minutes)
    h = total_minutes // 60
    m = total_minutes % 60
    return f"{sign}{h} h {m:02d} min"


def _estimate_weekly_average(summary: Dict) -> float:
    """Stima delle ore settimanali medie sul periodo analizzato."""
    period_start = summary["period_start"]
    period_end = summary["period_end"]
    delta_days = (period_end - period_start).days + 1
    weeks = delta_days / 7.0 if delta_days > 0 else 1.0
    return summary["total_hours_worked"] / weeks


def _slugify_name(name: str) -> str:
    cleaned = "".join(
        c if c.isalnum() or c in ("_", "-") else "_"
        for c in name.strip()
    )
    return cleaned or "dipendente"


def _build_comparison_filename(name1: str, name2: str) -> str:
    return f"Confronto_{_slugify_name(name1)}_vs_{_slugify_name(name2)}.pdf"


def _merge_monthly_data(df1: pd.DataFrame, df2: pd.DataFrame):
    """
    Unisce i dati mensili dei due dipendenti in base a 'month_label'
    e restituisce:
        - lista di mesi
        - dizionari ore lavorate/straordinari per ciascun dipendente
    """
    labels1 = df1["month_label"].tolist() if not df1.empty else []
    labels2 = df2["month_label"].tolist() if not df2.empty else []
    labels = sorted(set(labels1) | set(labels2))

    h1 = {row["month_label"]: row["hours_worked"] for _, row in df1.iterrows()}
    o1 = {row["month_label"]: row["overtime"] for _, row in df1.iterrows()}
    h2 = {row["month_label"]: row["hours_worked"] for _, row in df2.iterrows()}
    o2 = {row["month_label"]: row["overtime"] for _, row in df2.iterrows()}

    return labels, h1, o1, h2, o2


def _comparison_page1_overview(
    pdf,
    employee1_name: str,
    weekly_hours1: float,
    summary1: Dict,
    employee2_name: str,
    weekly_hours2: float,
    summary2: Dict,
):
    """Pagina 1: riepilogo sintetico confronto."""
    fig = plt.figure(figsize=(8.27, 11.69))  # A4
    plt.axis("off")

    # Titolo
    fig.text(0.5, 0.92, "Confronto ore lavorate", ha="center",
             fontsize=18, fontweight="bold")
    fig.text(0.5, 0.88, f"{employee1_name} vs {employee2_name}", ha="center",
             fontsize=14)

    period_label = summary1.get("period_label", "")
    fig.text(0.1, 0.82, f"Periodo analizzato: {period_label}", fontsize=11)

    avg_weekly1 = _estimate_weekly_average(summary1)
    avg_weekly2 = _estimate_weekly_average(summary2)

    y = 0.76
    fig.text(0.1, y, f"{employee1_name} (contratto {weekly_hours1:.0f} h/sett):",
             fontsize=11, fontweight="bold")
    y -= 0.03
    fig.text(0.12, y, f"Totale ore lavorate: {_format_hours_hm(summary1['total_hours_worked'])}",
             fontsize=10)
    y -= 0.025
    fig.text(0.12, y, f"Totale ore teoriche: {_format_hours_hm(summary1['total_theoretical_hours'])}",
             fontsize=10)
    y -= 0.025
    fig.text(0.12, y, f"Straordinario netto complessivo: {_format_hours_hm(summary1['total_overtime'])}",
             fontsize=10)
    y -= 0.025
    fig.text(0.12, y, f"Media ore settimanali (stima): {_format_hours_hm(avg_weekly1)}",
             fontsize=10)

    y -= 0.05
    fig.text(0.1, y, f"{employee2_name} (contratto {weekly_hours2:.0f} h/sett):",
             fontsize=11, fontweight="bold")
    y -= 0.03
    fig.text(0.12, y, f"Totale ore lavorate: {_format_hours_hm(summary2['total_hours_worked'])}",
             fontsize=10)
    y -= 0.025
    fig.text(0.12, y, f"Totale ore teoriche: {_format_hours_hm(summary2['total_theoretical_hours'])}",
             fontsize=10)
    y -= 0.025
    fig.text(0.12, y, f"Straordinario netto complessivo: {_format_hours_hm(summary2['total_overtime'])}",
             fontsize=10)
    y -= 0.025
    fig.text(0.12, y, f"Media ore settimanali (stima): {_format_hours_hm(avg_weekly2)}",
             fontsize=10)

    # Differenze principali
    y -= 0.05
    diff_hours = summary1["total_hours_worked"] - summary2["total_hours_worked"]
    diff_overtime = summary1["total_overtime"] - summary2["total_overtime"]

    fig.text(0.1, y, "Differenze principali:",
             fontsize=11, fontweight="bold")
    y -= 0.03
    fig.text(
        0.12,
        y,
        f"Differenza ore totali lavorate ({employee1_name} - {employee2_name}): "
        f"{_format_hours_hm(diff_hours)}",
        fontsize=10,
    )
    y -= 0.025
    fig.text(
        0.12,
        y,
        f"Differenza straordinari netti complessivi: {_format_hours_hm(diff_overtime)}",
        fontsize=10,
    )

    pdf.savefig(fig)
    plt.close(fig)


def _comparison_page2_monthly_table(
    pdf,
    employee1_name: str,
    monthly_df1: pd.DataFrame,
    employee2_name: str,
    monthly_df2: pd.DataFrame,
):
    """Pagina 2: tabella mensile ore / straordinari per entrambi."""
    labels, h1, o1, h2, o2 = _merge_monthly_data(monthly_df1, monthly_df2)

    fig, ax = plt.subplots(figsize=(8.27, 11.69))
    ax.axis("off")
    ax.set_title("Confronto mensile ore lavorate / straordinari",
                 fontsize=14, pad=20)

    header = [
        "Mese",
        f"Ore lavorate {employee1_name}",
        f"Straord. {employee1_name}",
        f"Ore lavorate {employee2_name}",
        f"Straord. {employee2_name}",
    ]
    table_data = [header]
    for m in labels:
        table_data.append([
            m,
            f"{h1.get(m, 0.0):.2f}",
            f"{o1.get(m, 0.0):.2f}",
            f"{h2.get(m, 0.0):.2f}",
            f"{o2.get(m, 0.0):.2f}",
        ])

    table = ax.table(cellText=table_data, loc="center", cellLoc="center")
    table.auto_set_font_size(False)
    table.set_fontsize(9)
    table.scale(1, 1.4)

    pdf.savefig(fig)
    plt.close(fig)


def _comparison_page3_monthly_hours(
    pdf,
    employee1_name: str,
    monthly_df1: pd.DataFrame,
    employee2_name: str,
    monthly_df2: pd.DataFrame,
):
    """Pagina 3: grafico ore lavorate per mese (linee), in A4 orizzontale."""
    labels, h1, _, h2, _ = _merge_monthly_data(monthly_df1, monthly_df2)
    y1 = [h1.get(m, 0.0) for m in labels]
    y2 = [h2.get(m, 0.0) for m in labels]

    # A4 orizzontale
    fig, ax = plt.subplots(figsize=(11.69, 8.27))
    ax.plot(labels, y1, marker="o", label=employee1_name)
    ax.plot(labels, y2, marker="o", label=employee2_name)
    ax.set_xlabel("Mese")
    ax.set_ylabel("Ore")
    ax.set_title("Ore lavorate per mese - Confronto")
    ax.grid(True, linestyle="--", alpha=0.4)
    ax.legend()
    plt.xticks(rotation=45)

    fig.tight_layout(pad=2.0)
    pdf.savefig(fig)
    plt.close(fig)



def _comparison_page4_monthly_overtime(
    pdf,
    employee1_name: str,
    monthly_df1: pd.DataFrame,
    employee2_name: str,
    monthly_df2: pd.DataFrame,
):
    """Pagina 4: grafico straordinari netti per mese (linee), in A4 orizzontale."""
    labels, _, o1, _, o2 = _merge_monthly_data(monthly_df1, monthly_df2)
    y1 = [o1.get(m, 0.0) for m in labels]
    y2 = [o2.get(m, 0.0) for m in labels]

    fig, ax = plt.subplots(figsize=(11.69, 8.27))  # A4 orizzontale
    ax.plot(labels, y1, marker="o", label=employee1_name)
    ax.plot(labels, y2, marker="o", label=employee2_name)
    ax.set_xlabel("Mese")
    ax.set_ylabel("Ore straordinarie")
    ax.set_title("Straordinario netto per mese - Confronto")
    ax.grid(True, linestyle="--", alpha=0.4)
    ax.legend()
    plt.xticks(rotation=45)

    fig.tight_layout(pad=2.0)
    pdf.savefig(fig)
    plt.close(fig)



def _comparison_page5_totals(
    pdf,
    employee1_name: str,
    summary1: Dict,
    employee2_name: str,
    summary2: Dict,
):
    """Pagina 5: barre complessive ore totali / straordinari totali, in A4 orizzontale."""
    labels = [employee1_name, employee2_name]
    ore_tot = [summary1["total_hours_worked"], summary2["total_hours_worked"]]
    straord_tot = [summary1["total_overtime"], summary2["total_overtime"]]

    x = range(len(labels))
    width = 0.35

    fig, ax = plt.subplots(figsize=(11.69, 8.27))  # A4 orizzontale
    ax.bar([p - width / 2 for p in x], ore_tot, width, label="Ore totali")
    ax.bar([p + width / 2 for p in x], straord_tot, width, label="Straordinari totali")

    ax.set_xticks(list(x))
    ax.set_xticklabels(labels)
    ax.set_ylabel("Ore")
    ax.set_title("Confronto complessivo ore e straordinari")
    ax.legend()
    ax.grid(axis="y", linestyle="--", alpha=0.4)

    fig.tight_layout(pad=2.0)
    pdf.savefig(fig)
    plt.close(fig)



def analyze_two_pdfs_comparison(
    pdf1_path: str,
    employee1_name: str,
    weekly_hours1: float,
    pdf2_path: str,
    employee2_name: str,
    weekly_hours2: float,
    report_folder: str,
) -> Tuple[str, Dict, Dict]:
    """
    Analizza due PDF e genera un unico report di confronto stile StatMaster.

    Ritorna:
        - nome file PDF di confronto
        - summary del dipendente 1
        - summary del dipendente 2
    """
    df1 = parse_pdf_to_dataframe(pdf1_path)
    summary1, monthly_df1 = compute_monthly_stats(df1, weekly_hours1)

    df2 = parse_pdf_to_dataframe(pdf2_path)
    summary2, monthly_df2 = compute_monthly_stats(df2, weekly_hours2)

    filename = _build_comparison_filename(employee1_name, employee2_name)
    report_path = os.path.join(report_folder, filename)

    with PdfPages(report_path) as pdf:
        _comparison_page1_overview(
            pdf,
            employee1_name,
            weekly_hours1,
            summary1,
            employee2_name,
            weekly_hours2,
            summary2,
        )
        _comparison_page2_monthly_table(
            pdf,
            employee1_name,
            monthly_df1,
            employee2_name,
            monthly_df2,
        )
        _comparison_page3_monthly_hours(
            pdf,
            employee1_name,
            monthly_df1,
            employee2_name,
            monthly_df2,
        )
        _comparison_page4_monthly_overtime(
            pdf,
            employee1_name,
            monthly_df1,
            employee2_name,
            monthly_df2,
        )
        _comparison_page5_totals(
            pdf,
            employee1_name,
            summary1,
            employee2_name,
            summary2,
        )

    return filename, summary1, summary2
