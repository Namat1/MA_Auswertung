import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

st.title("Tourenauswertung – nach Jahr, KW (Sonntag), Name & Tour")

uploaded_files = st.file_uploader("Excel-Dateien hochladen", type=["xlsx"], accept_multiple_files=True)

# Deutsche Wochentage
wochentage_deutsch = {
    "Monday": "Montag", "Tuesday": "Dienstag", "Wednesday": "Mittwoch",
    "Thursday": "Donnerstag", "Friday": "Freitag",
    "Saturday": "Samstag", "Sunday": "Sonntag"
}

# ✅ Name aus 3+4 oder alternativ 6+7, robust gegen Nicht-Strings
def extract_name(row):
    if pd.notna(row[3]) and pd.notna(row[4]):
        return f"{str(row[3]).strip()} {str(row[4]).strip()}"
    elif pd.notna(row[6]) and pd.notna(row[7]):
        return f"{str(row[6]).strip()} {str(row[7]).strip()}"
    return None

# ✅ KW mit Sonntag als Wochenstart + korrektes Jahr
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

# ----------------- Streamlit Hauptteil -----------------
if uploaded_files:
    all_data = []
    alle_namen = set()

    for file in uploaded_files:
        try:
            df = pd.read_excel(file, sheet_name="Touren", header=None)
            df = df.iloc[5:].reset_index(drop=True)
            df["Name"] = df.apply(extract_name, axis=1)
            alle_namen.update(df["Name"].dropna().unique())

            df["Datum"] = pd.to_datetime(df[14], errors='coerce')
            df[["KW", "Jahr"]] = df["Datum"].apply(lambda x: pd.Series(get_kw_and_year_sunday_start(x)))
            df["Tour"] = df[15]
            df["Uhrzeit"] = df[8]
            df["LKW"] = df[11]

            df = df[["KW", "Jahr", "Datum", "Name", "Tour", "Uhrzeit", "LKW"]]
            all_data.append(df)
        except Exception as e:
            st.error(f"Fehler beim Verarbeiten von {file.name}: {e}")

    if alle_namen:
        # Sortiere alphabetisch nach Nachname
        def sort_nachname(name):
            return name.split()[0].lower() if isinstance(name, str) else ""

        sorted_names = sorted(list(alle_namen), key=sort_nachname)
        selected_name = st.selectbox("Fahrer auswählen", sorted_names)

        if selected_name:
            result_df = pd.concat(all_data)
            result_df = result_df[result_df["Name"] == selected_name]
            result_df.sort_values(by=["Jahr", "KW", "Datum"], inplace=True)

            result_df["Wochentag"] = result_df["Datum"].dt.day_name().map(wochentage_deutsch)
            result_df["Datum_formatiert"] = result_df["Datum"].dt.strftime('%d.%m.%Y')
            result_df["Datum_komplett"] = result_df["Wochentag"] + ", " + result_df["Datum_formatiert"]

            df_final = result_df[["KW", "Jahr", "Datum_komplett", "Name", "Tour", "Uhrzeit", "LKW"]]

            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                sheet = "Alle_KWs"
                ws = writer.book.create_sheet(title=sheet)
                writer.sheets[sheet] = ws
                start_row = 1

                for (jahr, kw), group in df_final.groupby(["Jahr", "KW"]):
                    group = group.reset_index(drop=True)

                    ws.cell(row=start_row, column=1, value=f"KW {int(kw)} ({int(jahr)})")
                    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=7)
                    cell = ws.cell(row=start_row, column=1)
                    cell.font = Font(bold=True, size=14)
                    cell.alignment = Alignment(horizontal="left")
                    cell.fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
                    start_row += 1

                    header = ["KW", "Jahr", "Datum", "Name", "Tour", "Uhrzeit", "LKW"]
                    for col_num, column_title in enumerate(header, 1):
                        cell = ws.cell(row=start_row, column=col_num, value=column_title)
                        cell.font = Font(bold=True)
                        cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                        cell.alignment = Alignment(horizontal="left", vertical="center")
                    start_row += 1

                    for row in group.itertuples(index=False):
                        values = [row.KW, row.Jahr, row.Datum_komplett, row.Name, row.Tour, row.Uhrzeit, row.LKW]
                        for col_num, value in enumerate(values, 1):
                            cell = ws.cell(row=start_row, column=col_num, value=value)
                            cell.alignment = Alignment(horizontal="left", vertical="center")
                        start_row += 1

                    start_row += 1

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
                               file_name=f"{selected_name}_touren.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.warning("Keine Fahrernamen erkannt.")
