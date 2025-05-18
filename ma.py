import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

st.title("Tourenauswertung – beide Seiten, gefiltert per Fahrernamen")

uploaded_files = st.file_uploader("Excel-Dateien hochladen", type=["xlsx"], accept_multiple_files=True)
fahrersuche = st.text_input("Fahrername eingeben (z. B. 'demuth' oder 'harry')").strip().lower()

# Wochentage Deutsch
wochentage_deutsch = {
    "Monday": "Montag", "Tuesday": "Dienstag", "Wednesday": "Mittwoch",
    "Thursday": "Donnerstag", "Friday": "Freitag",
    "Saturday": "Samstag", "Sunday": "Sonntag"
}

# KW mit Sonntag als Start
def get_kw_and_year_sunday_start(datum):
    try:
        dt = pd.to_datetime(datum)
        kw = int(dt.strftime("%U")) + 1
        jahr = dt.year
        if dt.month == 1 and kw >= 52:
            jahr -= 1
        return kw, jahr
    except:
        return None, None

def extract_entries_both_sides(row):
    eintraege = []

    # Basisdaten
    datum = pd.to_datetime(row[14], errors="coerce")
    if pd.isna(datum):
        return []

    kw, jahr = get_kw_and_year_sunday_start(datum)
    wochentag = datum.day_name()
    wochentag_de = wochentage_deutsch.get(wochentag, wochentag)
    datum_formatiert = datum.strftime('%d.%m.%Y')
    datum_komplett = f"{wochentag_de}, {datum_formatiert}"

    # Daten für links (3+4)
    if pd.notna(row[3]) and pd.notna(row[4]):
        name = f"{str(row[3]).strip()} {str(row[4]).strip()}"
        eintraege.append({
            "KW": kw,
            "Jahr": jahr,
            "Datum": datum_komplett,
            "Name": name,
            "Tour": row[15],
            "Uhrzeit": row[8],
            "LKW": row[11]
        })

    # Daten für rechts (6+7)
    if pd.notna(row[6]) and pd.notna(row[7]):
        name = f"{str(row[6]).strip()} {str(row[7]).strip()}"
        eintraege.append({
            "KW": kw,
            "Jahr": jahr,
            "Datum": datum_komplett,
            "Name": name,
            "Tour": row[15],
            "Uhrzeit": row[8],
            "LKW": row[11]
        })

    return eintraege

# Hauptlogik
if uploaded_files:
    eintraege_gesamt = []

    for file in uploaded_files:
        try:
            df = pd.read_excel(file, sheet_name="Touren", header=None)
            df = df.iloc[5:].reset_index(drop=True)

            # Für jede Zeile ggf. zwei Einträge erzeugen
            for _, row in df.iterrows():
                eintraege = extract_entries_both_sides(row)
                eintraege_gesamt.extend(eintraege)
        except Exception as e:
            st.error(f"Fehler beim Verarbeiten von {file.name}: {e}")

    if eintraege_gesamt:
        df_final = pd.DataFrame(eintraege_gesamt)

        # Filter nach Fahrernamen (Teilwort, case-insensitive)
        if fahrersuche:
            df_final = df_final[df_final["Name"].str.lower().str.contains(fahrersuche)]

        if df_final.empty:
            st.warning("Keine Einträge für diesen Fahrer gefunden.")
        else:
            df_final.sort_values(by=["Jahr", "KW", "Datum"], inplace=True)

            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                sheet = "Touren"
                ws = writer.book.create_sheet(title=sheet)
                writer.sheets[sheet] = ws

                start_row = 1
                ws.append(["KW", "Jahr", "Datum", "Name", "Tour", "Uhrzeit", "LKW"])

                # Format Kopfzeile
                for col_num in range(1, 8):
                    cell = ws.cell(row=start_row, column=col_num)
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                    cell.alignment = Alignment(horizontal="left")
                start_row += 1

                # Daten einfügen
                for row in df_final.itertuples(index=False):
                    ws.append([row.KW, row.Jahr, row.Datum, row.Name, row.Tour, row.Uhrzeit, row.LKW])

                # Autobreite
                for col in ws.columns:
                    max_length = 0
                    col_letter = get_column_letter(col[0].column)
                    for cell in col:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    ws.column_dimensions[col_letter].width = int(max_length * 1.5)

            output.seek(0)
            st.success("Auswertung abgeschlossen.")
            st.download_button("Excel-Datei herunterladen",
                               output,
                               file_name=f"touren_auswertung.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
