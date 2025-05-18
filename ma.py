import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

st.title("Tourenauswertung – links & rechts getrennt")

uploaded_files = st.file_uploader("Excel-Dateien hochladen", type=["xlsx"], accept_multiple_files=True)

wochentage_deutsch = {
    "Monday": "Montag", "Tuesday": "Dienstag", "Wednesday": "Mittwoch",
    "Thursday": "Donnerstag", "Friday": "Freitag",
    "Saturday": "Samstag", "Sunday": "Sonntag"
}

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

def extract_both_names(row):
    eintraege = []
    for side, (n_col, v_col) in zip(["links", "rechts"], [(3, 4), (6, 7)]):
        if pd.notna(row[n_col]) and pd.notna(row[v_col]):
            name = f"{str(row[n_col]).strip()} {str(row[v_col]).strip()}"
            eintraege.append({
                "Name": name,
                "Datum": pd.to_datetime(row[14], errors='coerce'),
                "Tour": row[15],
                "Uhrzeit": row[8],
                "LKW": row[11]
            })
    return eintraege

if uploaded_files:
    result_rows = []

    for file in uploaded_files:
        try:
            df = pd.read_excel(file, sheet_name="Touren", header=None)
            df = df.iloc[5:].reset_index(drop=True)

            for _, row in df.iterrows():
                eintraege = extract_both_names(row)
                for eintrag in eintraege:
                    eintrag["KW"], eintrag["Jahr"] = get_kw_and_year_sunday_start(eintrag["Datum"])
                    result_rows.append(eintrag)

        except Exception as e:
            st.error(f"Fehler in Datei {file.name}: {e}")

    if result_rows:
        df_final = pd.DataFrame(result_rows)
        df_final["Wochentag"] = df_final["Datum"].dt.day_name().map(wochentage_deutsch)
        df_final["Datum_komplett"] = df_final["Wochentag"] + ", " + df_final["Datum"].dt.strftime('%d.%m.%Y')

        df_export = df_final[["KW", "Jahr", "Datum_komplett", "Name", "Tour", "Uhrzeit", "LKW"]]
        df_export.sort_values(by=["Jahr", "KW", "Datum_komplett"], inplace=True)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            ws = writer.book.create_sheet(title="Alle_KWs")
            writer.sheets["Alle_KWs"] = ws
            start_row = 1

            for (jahr, kw), group in df_export.groupby(["Jahr", "KW"]):
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
                           file_name="touren_alle_seiten.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
