
import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO

st.set_page_config(page_title="Arbeitszeiten-Extraktion", layout="wide")
st.title("🕒 Arbeitszeit-Extraktion aus MyTMA-PDF")

st.markdown("""
### ℹ️ Anleitung zur Nutzung

1. **Zeiten aus MyTMA exportieren:**  
   - Menüpunkt *Auskunft → Selbstauskunft*  
   - Dann **Monat und Jahr wählen** und unten die beiden Haken bei **„Bemerkungen“** und **„Kalenderwochen“** deaktivieren  
   - Auf **„Drucken“ klicken** und das PDF irgendwo abspeichern  

2. **PDF-Datei hochladen**, die aus dem MyTMA-System exportiert wurde.

3. Das Tool liest die Tabelle automatisch aus und extrahiert die Zeitangaben** (Von1, Bis1, Von2, Bis2).

4. Es berechnet **Von_gesamt** (erste Zeit) und **Bis_gesamt** (letzte Zeit). Achtung, die Pausen werden nicht rausgerechnet.

5. Du kannst die berechneten **Stunden und Minuten als Excel-Datei herunterladen**.

💡 6. Markiere die 4 Spalten mit den von und bis Stunden/Minuten und kopiere diese (mit Werte einfügen) in die Zeiterfassungstabelle. Die für das Projekt gearbeiteten Minuten kannst Du dann von Hand in der Spalte N ergänzen.
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

    # ❗ Filter: Nur Zeilen mit Datum 01.–31. zulassen
    df = df[df["Datum"].str.match(r"^(0[1-9]|[12][0-9]|3[01])\.")]

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

    export_df = df[[
        "Datum", "Wochentag",
        "Von_gesamt_Stunde", "Von_gesamt_Minute",
        "Bis_gesamt_Stunde", "Bis_gesamt_Minute"
    ]]

    return export_df

if uploaded_file:
    pdf_bytes = uploaded_file.read()
    with st.spinner("Verarbeite PDF..."):
        df_result = extract_times_from_pdf(pdf_bytes)
        st.success("Extraktion abgeschlossen!")
        st.dataframe(df_result)

        buffer = BytesIO()
        df_result.to_excel(buffer, index=False, engine="openpyxl")
        st.download_button("📥 Excel herunterladen", buffer.getvalue(),
                           file_name="Arbeitszeiten_Export.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
