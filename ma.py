import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

st.title("Tourenauswertung – KW, Name, Tour in einem Blatt")

uploaded_files = st.file_uploader("Excel-Dateien hochladen", type=["xlsx"], accept_multiple_files=True)

name_query = st.text_input("Gesuchten Fahrer eingeben (Teil von Vor- oder Nachname):")

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
            df["Tour"] = df[0]
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

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            sheet_name = "Alle_KWs"
            result_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)

            wb = writer.book
            ws = writer.sheets[sheet_name]

            # Kopfzeile einfärben und fett
            header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
            for col_num, column_title in enumerate(result_df.columns, 1):
                cell = ws.cell(row=2, column=col_num)
                cell.font = Font(bold=True)
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")

            # Alle Inhalte zentrieren
            for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                for cell in row:
                    cell.alignment = Alignment(horizontal="center", vertical="center")

            # Spaltenbreite automatisch (150 %)
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
