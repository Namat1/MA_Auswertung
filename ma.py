import streamlit as st
import pandas as pd
import datetime
from io import BytesIO

st.title("Touren-Auswertung nach Kalenderwoche, Name & Tour")

uploaded_files = st.file_uploader("Excel-Dateien hochladen", type=["xlsx"], accept_multiple_files=True)

name_query = st.text_input("Bitte gesuchten Fahrer (Teil von Vor- oder Nachname) eingeben:")

def extract_name(row):
    """Extrahiert Namen aus Spalte 3+4 oder 6+7"""
    if pd.notna(row[3]) and pd.notna(row[4]):
        return f"{row[3]} {row[4]}"
    elif pd.notna(row[6]) and pd.notna(row[7]):
        return f"{row[6]} {row[7]}"
    return None

def get_kw(datum):
    try:
        return pd.to_datetime(datum).isocalendar().week
    except Exception:
        return None

if uploaded_files and name_query:
    all_data = []

    for file in uploaded_files:
        try:
            df = pd.read_excel(file, sheet_name="Touren", header=None)
            df = df.iloc[5:]  # Ab Zeile 6 beginnen
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
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for kw, group in result_df.groupby("KW"):
                sheet_name = f"KW{kw}" if kw is not None else "Unbekannt"
                group.to_excel(writer, sheet_name=sheet_name, index=False)

        output.seek(0)
        st.success("Auswertung abgeschlossen.")
        st.download_button("Ausgewertete Excel-Datei herunterladen",
                           output,
                           file_name="tourenauswertung.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.warning("Kein passender Name in den Dateien gefunden.")
