import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

st.title("Tourenauswertung – nach Jahr, KW (Sonntag), Name & Tour")

uploaded_files = st.file_uploader("Excel-Dateien hochladen", type=["xlsx"], accept_multiple_files=True)
name_query = st.text_input("Gesuchten Fahrer eingeben (Teil von Vor- oder Nachname):")

# Deutsche Wochentage
wochentage_deutsch = {
    "Monday": "Montag", "Tuesday": "Dienstag", "Wednesday": "Mittwoch",
    "Thursday": "Donnerstag", "Friday": "Freitag",
    "Saturday": "Samstag", "Sunday": "Sonntag"
}

def extract_name(row):
    if pd.notna(row[3]) and pd.notna(row[4]):
        return f"{row[3]} {row[4]}"
    elif pd.notna(row[6]) and pd.notna(row[7]):
        return f"{row[6]} {row[7]}"
    return None

# Sonntag als Start der Kalenderwoche
def get_kw_and_year_sunday_start(datum):
    try:
        dt = pd.to_datetime(datum)
        verschoben = dt - pd.DateOffset(days=(dt.weekday() + 1) % 7)
        iso = verschoben.isocalendar()
        return iso.week, iso.year
    except:
        return None, None

if uploaded_files and name_query:
    all_data = []

    for file in uploaded_files:
        try:
            df = pd.read_excel(file, sheet_name="Touren", header=None)
            df = df.iloc[5:]  # Ab Zeile 6
            df = df.reset_index(drop=True)

            df["Name"] = df.apply(extract_name, axis=1)
            df["Datum"] = pd.to_datetime(df[14], errors='coerce')
            df[["KW", "Jahr"]] = df["Datum"].apply(lambda x: pd.Series(get_kw_and_year_sunday_start(x)))
            df["Tour"] = df[15]
            df["Uhrzeit"] = df[8]
            df["LKW"] = df[11]

            df = df[df["Name"].str.contains(name_query, case=False, na=False)]
            df = df[["KW", "Jahr", "Datum", "Name", "Tour", "Uhrzeit", "LKW"]]

            all_data.append(df)
        except Exception as e:
            st.error(f"Fehler beim Verarbeiten von {file.name}: {e}")

    if all_data:
        result_df = pd.concat(all_data)
        result_df.sort_values(by=["Jahr", "KW", "Datum"], inplace=True)

        # Deutsches Datum mit Wochentag
        result_df["Wochentag"] = result_df["Datum"].dt.day_name().map(wochentage_deutsch)
        result_df["Datum_formatiert"] = result_df["Datum"].dt.strftime('%d.%m.%Y')
        result_df["Datum_komplett"] = result_df["Wochentag"] + ", " + result_df["Datum_formatiert"]

        # Finaler Export-DataFrame
        df_final = result_df[["KW", "Jahr", "Datum_komplett", "Name", "Tour", "Uhrzeit", "LKW"]]

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            sheet = "Alle_KWs"
            ws = writer.book.create_sheet(title=sheet)
            writer.sheets[sheet] = ws
            start_row = 1

            for (jahr, kw), group in df_final.groupby(["Jahr", "KW"]):
                group = group.reset_index(drop=True)

                # KW-Jahr-Blocküberschrift
                ws.cell(row=start_row, column=1, value=f"KW {kw} ({jahr})")
                ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=7)
                cell = ws.cell(row=start_row, column=1)
                cell.font = Font(bold=True, size=14)
                cell.alignment = Alignment(horizontal="left")
                cell.fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
                start_row += 1

                # Spaltenüberschriften
                header = ["KW", "Jahr", "Datum", "Name", "Tour", "Uhrzeit", "LKW"]
                for col_num, column_title in enumerate(header, 1):
                    cell = ws.cell(row=start_row, column=col_num, value=column_title)
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                start_row += 1

                # Daten
                for row in group.itertuples(index=False):
                    values = [row.KW, row.Jahr, row.Datum_komplett, row.Name, row.Tour, row.Uhrzeit, row.LKW]
                    for col_num, value in enumerate(values, 1):
                        cell = ws.cell(row=start_row, column=col_num, value=value)
                        cell.alignment = Alignment(horizontal="left", vertical="center")
                    start_row += 1

                # Leere Zeile zwischen KW-Blöcken
                start_row += 1

            # Autobreite auf 150 % Inhalt
            for col in ws.columns:
                max_length = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                ws.column_dimensions[col_letter].width = int(max_length * 1.5)

        output.seek(0)
        st.success("Auswertung abgeschlossen.")
        st.download_button("Ausgewertete Excel-Datei herunterladen",
                           output,
                           file_name="tourenauswertung.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.warning("Kein passender Name gefunden.")
