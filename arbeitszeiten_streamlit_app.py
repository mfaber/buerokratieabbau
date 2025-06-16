
import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Arbeitszeiten-Extraktion", layout="wide")
st.title("🕒 Arbeitszeit-Extraktion aus MyTMA-PDF")

st.markdown("""
### ℹ️ Anleitung zur Nutzung

1. **Zeiten aus MyTMA exportieren:**  
   - Menüpunkt *Auskunft → Selbstauskunft*  
   - Dann **Monat und Jahr wählen** und unten die beiden Haken bei **„Bemerkungen“** und **„Kalenderwochen“** deaktivieren  
   - Auf **„Drucken“ klicken** und das PDF irgendwo abspeichern  

2. **PDF-Datei hochladen**, die aus dem MyTMA-System exportiert wurde.

3. Das Tool liest die Tabelle automatisch aus und extrahiert bis zu **vier Zeitangaben** (Von1, Bis1, Von2, Bis2).

4. Es berechnet **Von_gesamt** (erste Zeit) und **Bis_gesamt** (letzte Zeit mit Fallback).

5. Du kannst die berechneten **Stunden und Minuten als Excel-Datei herunterladen** – komplett formatiert wie deine Vorlage.
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
    df = df.iloc[1:]  # erste Zeile entfernen

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

def export_formatiert_excel(df):
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Arbeitszeiten"

    # Farben/Rahmen
    yellow_fill = PatternFill("solid", fgColor="FFFF99")
    blue_border_thin = Side(style="thin", color="0000FF")
    blue_border_thick = Side(style="medium", color="0000FF")

    # Kopfzeile
    ws["A1"] = "Wo."
    ws["B1"] = "tag"
    ws["C1"] = "Tag"
    ws.merge_cells("A1:B1")
    ws.merge_cells("C1:C2")
    ws["D1"] = "Beginn"
    ws["F1"] = "Ende"
    ws.merge_cells("D1:E1")
    ws.merge_cells("F1:G1")
    ws["D2"] = "Std"
    ws["E2"] = "Min"
    ws["F2"] = "Std"
    ws["G2"] = "Min"

    # Spaltenbreiten
    for col, width in zip("ABCDEFG", [5, 5, 6, 6, 5, 6, 5]):
        ws.column_dimensions[col].width = width

    for i, row in df.iterrows():
        r = i + 3
        ws.cell(row=r, column=1).value = row["Datum"]
        ws.cell(row=r, column=2).value = row["Datum"]
        ws.cell(row=r, column=3).value = row["Wochentag"]
        ws.cell(row=r, column=4).value = row["Von_gesamt_Stunde"]
        ws.cell(row=r, column=5).value = row["Von_gesamt_Minute"]
        ws.cell(row=r, column=6).value = row["Bis_gesamt_Stunde"]
        ws.cell(row=r, column=7).value = row["Bis_gesamt_Minute"]

    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=1, max_col=7):
        wtag = row[2].value
        is_weekend = str(wtag).strip() in {"Sa", "So"}

        for cell in row:
            cell.border = Border(top=blue_border_thin, bottom=blue_border_thin,
                                 left=blue_border_thin, right=blue_border_thin)
            if is_weekend:
                cell.fill = yellow_fill

        for i in [3, 4, 5, 6]:  # CDEF → Beginn/Ende
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

        excel_bytes = export_formatiert_excel(df_result)
        st.download_button("📥 Excel herunterladen", excel_bytes,
                           file_name="Arbeitszeiten_Export_formatiert.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
