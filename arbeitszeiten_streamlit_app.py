
import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment

st.set_page_config(page_title="Arbeitszeiten-Extraktion", layout="wide")
st.title("🕒 Arbeitszeit-Extraktion aus MyTMA-PDF")

st.markdown("""
## ℹ️ Anleitung zur Nutzung

1. **Zeiten aus MyTMA exportieren:**
   - Menüpunkt *Auskunft → Selbstauskunft*
   - Dann **Monat und Jahr wählen** und unten die beiden Haken bei **„Bemerkungen“** und **„Kalenderwochen“** deaktivieren
   - Auf **„Drucken“ klicken** und das PDF irgendwo abspeichern

2. **PDF-Datei hochladen**, die aus dem MyTMA-System exportiert wurde.

4. Es berechnet **Von_gesamt** (erste Zeit) und **Bis_gesamt** (letzte Zeit). Achtung, die Pausen werden nicht rausgerechnet.

5. Du kannst die berechneten **Stunden und Minuten als Excel-Datei herunterladen**.

💡 6. Markiere die 4 Spalten mit den von und bis Stunden/Minuten und kopiere diese (mit Werte einfügen) in die Zeiterfassungstabelle.
    Die für das Projekt gearbeiteten Minuten kannst Du dann von Hand in der Spalte N ergänzen.

7. (optional) Bitte die Verwaltung, in Zukunft auf solche Prozesse zu verzichten, geeignete Workflows
   (copy-paste statt Zahlen vom einen Verwaltungssystem in ein anderes zu übertragen) zur Verfügung zu stellen
   oder solche Arbeiten selbst auszuführen ;).

Fragen, Anregungen zum Tool: faberm@rki.de
""")

uploaded_file = st.file_uploader("PDF-Datei hochladen", type="pdf")

def extract_times_from_pdf(pdf_bytes):
    pages = []
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                pages.append(pd.DataFrame(table[1:], columns=table[0]))

    extrahiert = []
    for page in pages:
        for i, row in page.iterrows():
            row_text = " ".join(str(cell) for cell in row if cell)
            zeiten = re.findall(r"\d{2}:\d{2}", row_text)
            if re.match(r"^\d{2}\.\d{2}\.", row_text):
                datum = row[0]
                wochentag = row[1]
                von1 = zeiten[0] if len(zeiten) > 0 else ""
                bis1 = zeiten[1] if len(zeiten) > 1 else ""
                von2 = zeiten[2] if len(zeiten) > 2 else ""
                bis2 = zeiten[3] if len(zeiten) > 3 else ""
                extrahiert.append([datum, wochentag, von1, bis1, von2, bis2])

    df = pd.DataFrame(extrahiert, columns=["Datum", "Wochentag", "Von1", "Bis1", "Von2", "Bis2"])
    if df.iloc[0]["Datum"].startswith(("28.", "29.", "30.", "31.")):
        df = df.iloc[1:]

    def parse_time(text):
        match = re.match(r"(\d{1,2})[:\.]?(\d{2})", str(text))
        if match:
            return int(match.group(1)), int(match.group(2))
        return pd.NA, pd.NA

    for col in ["Von1", "Bis1", "Von2", "Bis2"]:
        df[f"{col}_Stunde"], df[f"{col}_Minute"] = zip(*df[col].apply(parse_time))

    df["Von_gesamt"] = df["Von1"]
    df["Bis_gesamt"] = df["Bis2"].where(df["Bis2"].fillna("").astype(str).str.strip().astype(bool), df["Bis1"])

    df["Von_gesamt_Stunde"], df["Von_gesamt_Minute"] = zip(*df["Von_gesamt"].apply(parse_time))
    df["Bis_gesamt_Stunde"], df["Bis_gesamt_Minute"] = zip(*df["Bis_gesamt"].apply(parse_time))

    def ist_gueltige_zeit(st, mi):
        return pd.notna(st) and 0 <= st <= 23 and 0 <= mi <= 59

    for col in ["Von_gesamt", "Bis_gesamt"]:
        stc, mic = f"{col}_Stunde", f"{col}_Minute"
        mask = ~df[[stc, mic]].apply(lambda x: ist_gueltige_zeit(x[0], x[1]), axis=1)
        df.loc[mask, [col, stc, mic]] = pd.NA

    df = df.astype({col: "Int64" for col in df.columns if col.endswith("_Stunde") or col.endswith("_Minute")})
    return df

def create_formatted_excel(df):
    blue_border_thin = Side(style="thin", color="0000FF")
    blue_border_thick = Side(style="medium", color="0000FF")
    wb = Workbook()
    ws = wb.active
    ws["L1"] = "ℹ️ Anleitung zur Nutzung:"
    ws["L2"] = "1. Zeiten aus MyTMA exportieren:"
    ws["L3"] = "   - Auskunft → Selbstauskunft"
    ws["L4"] = "   - Monat und Jahr wählen, Haken bei 'Bemerkungen' und 'Kalenderwochen' deaktivieren"
    ws["L5"] = "   - Auf 'Drucken' klicken und PDF abspeichern"
    ws["L6"] = "2. PDF-Datei hochladen, die aus dem MyTMA-System exportiert wurde."
    ws["L7"] = "3. Von_gesamt = erste Zeit, Bis_gesamt = letzte Zeit. Pausen werden nicht abgezogen."
    ws["L8"] = "4. Excel herunterladen, Zeitspalten kopieren, in Zeiterfassungstabelle einfügen."
    ws["L9"] = "5. Minuten in Spalte N manuell ergänzen."
    ws["L10"] = "6. (optional) Verwaltung um geeignete Workflows bitten."
    ws["L11"] = "7. Fragen, Anregungen zum Tool: faberm@rki.de"
    for col, width in zip("ABCDEFG", [5, 5, 6, 6, 5, 6, 5]):
        ws.column_dimensions[col].width = width


        ws.column_dimensions[col].width = width

    for i, row in df.iterrows():
        r = i + 2
        ws.cell(row=r, column=1).value = row["Datum"]
        ws.cell(row=r, column=2).value = row["Datum"]
        ws.cell(row=r, column=3).value = row["Wochentag"]
        ws.cell(row=r, column=4).value = int(row["Von_gesamt_Stunde"]) if pd.notna(row["Von_gesamt_Stunde"]) else None
        ws.cell(row=r, column=5).value = int(row["Von_gesamt_Minute"]) if pd.notna(row["Von_gesamt_Minute"]) else None
        ws.cell(row=r, column=6).value = int(row["Bis_gesamt_Stunde"]) if pd.notna(row["Bis_gesamt_Stunde"]) else None
        ws.cell(row=r, column=7).value = int(row["Bis_gesamt_Minute"]) if pd.notna(row["Bis_gesamt_Minute"]) else None

    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=1, max_col=7):
        wtag = row[2].value
        is_weekend = str(wtag).strip() in {"Sa", "So"}

        for cell in row:
            cell.border = Border(top=blue_border_thin, bottom=blue_border_thin,
                                 left=blue_border_thin, right=blue_border_thin)
            if is_weekend:
                cell.fill = PatternFill("solid", fgColor="FFFF99")

        for i in [3, 4, 5, 6]:
            row[i].border = Border(top=blue_border_thick, bottom=blue_border_thick,
                                   left=blue_border_thick, right=blue_border_thick)

        buffer = BytesIO()
    
    wb.save(buffer)
    return buffer.getvalue()

if uploaded_file:
    pdf_bytes = uploaded_file.read()
    with st.spinner("Verarbeite PDF..."):
        df_result = extract_times_from_pdf(pdf_bytes)
        st.success("Extraktion abgeschlossen!")
        st.dataframe(df_result)

        excel_bytes = create_formatted_excel(df_result)
        st.download_button("📥 Excel herunterladen", excel_bytes,
                           file_name="Arbeitszeiten_Export_formatiert.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
