import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

st.title("Tourenauswertung – nach KW, Fahrername & Tour")

uploaded_files = st.file_uploader("Excel-Dateien hochladen", type=["xlsx"], accept_multiple_files=True)
name_query = st.text_input("Gesuchten Fahrer eingeben (Teil von Vor- oder Nachname):")

# Deutsche Wochentage manuell
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

def get_kw(datum):
    try:
        return pd.to_datetime(datum).isocalendar().week
    except:
        return None

if uploaded_files and name_query:
    all_data = []

    for file in uploaded_files:
        try:
            df = pd.read_excel(file, sheet_name="Touren", header=None)
            df = df.iloc[5:]  # ab Zeile 6
            df = df.reset_index(drop=True)

            df["Name"] = df.apply(extract_name, axis=1)
            df["Datum"] = pd.to_datetime(df[14], errors='coerce')
            df["KW"] = df["Datum"].apply(get_kw)
            df["Tour"] = df[15]  # Tournummer ist in Spalte 15
            df["Uhrzeit"] = df[8]
            df["LKW"] = df[11]

            df = df[df["Name"].str.contains(name_query, case=False, na=False)]
            df = df[["KW", "Datum", "Name", "Tour", "Uhrzeit", "LKW"]]

            all_data.append(df)
        except Exception as e:
            st.error(f"Fehler beim Verarbeiten von {file.name}: {e}")

    if all_data:
        result_df = pd.concat(all_data)
        result_df.sort_values(by=["KW", "Datum", "Name"], inplace=True)

        # Wochentag + deutsches Datumsformat
        result_df["Wochentag"] = result_df["Datum"].dt.day_name().map(wochentage_deutsch)
        result_df["Datum_formatiert"] = result_df["Datum"].dt.strftime('%d.%m.%Y')
        result_df["Datum_komplett"] = result_df["Wochentag"] + ", " + result_df["Datum_formatiert"]
        result_df.drop(columns=["Datum", "Wochentag", "Datum_formatiert"], inplace=True)
        result_df = result_df[["KW", "Datum_komplett", "Name", "Tour", "Uhrzeit", "LKW"]]

        # Excel-Ausgabe
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            sheet_name = "Alle_KWs"
            start_row = 1
            wb = writer.book

            for kw, group in result_df.groupby("KW"):
                group = group.sort_values(by="Datum_komplett")
                ws = writer.book.create_sheet(title=sheet_name) if writer.sheets == {} else writer.sheets[sheet_name]

                # KW-Titelzeile
                ws.cell(row=start_row, column=1, value=f"KW {kw}")
                ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=6)
                title_cell = ws.cell(row=start_row, column=1)
                title_cell.font = Font(bold=True, size=14)
                title_cell.alignment = Alignment(horizontal="center")
                title_cell.fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
                start_row += 1

                # Spaltenüberschriften
                for col_num, column_title in enumerate(group.columns, 1):
                    cell = ws.cell(row=start_row, column=col_num, value=column_title)
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                start_row += 1

                # Datenzeilen
                for row in group.itertuples(index=False):
                    for col_num, value in enumerate(row, 1):
                        cell = ws.cell(row=start_row, column=col_num, value=value)
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    start_row += 1

                # Leerzeile zwischen KW-Blöcken
                start_row += 1

            # Spaltenbreiten auf 150 % Inhalt
            for col in ws.columns:
                max_length = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                ws.column_dimensions[col_letter].width = int(max_length * 1.5)

        output.seek(0)
        st.success("Auswertung abgeschlossen.")
        st.download_button("Ausgewertete Excel-Datei herunterladen",
                           output,
                           file_name="tourenauswertung.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.warning("Kein passender Name gefunden.")
