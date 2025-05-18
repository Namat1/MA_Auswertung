import streamlit as st
import pandas as pd
import datetime
from io import BytesIO

# Streamlit UI
st.title("Excel-Auswertung nach Name, KW und Tour")

uploaded_files = st.file_uploader("Excel-Dateien hochladen", type=["xlsx"], accept_multiple_files=True)

name_query = st.text_input("Bitte Nachnamen oder Vornamen eingeben (Teil reicht):")

def extract_name(row):
    # Priorität auf Spalte 4+5, sonst 7+8
    if pd.notna(row[4]) and pd.notna(row[5]):
        return f"{row[4]} {row[5]}"
    elif pd.notna(row[7]) and pd.notna(row[8]):
        return f"{row[7]} {row[8]}"
    return None

def get_kw(datum):
    try:
        return pd.to_datetime(datum).isocalendar().week
    except Exception:
        return None

if uploaded_files and name_query:
    all_data = []

    for file in uploaded_files:
        df = pd.read_excel(file, header=None)
        df = df.iloc[4:]  # ab Zeile 5
        df = df.reset_index(drop=True)

        df["Name"] = df.apply(extract_name, axis=1)
        df["Datum"] = pd.to_datetime(df[14], errors='coerce')
        df["KW"] = df["Datum"].apply(get_kw)
        df["Tour"] = df[0]
        df["Uhrzeit"] = df[8]
        df["LKW"] = df[11]

        df = df[df["Name"].str.contains(name_query, case=False, na=False)]
        df = df[["KW", "Name", "Datum", "Tour", "Uhrzeit", "LKW"]]

        all_data.append(df)

    if all_data:
        result_df = pd.concat(all_data)
        result_df.sort_values(by=["KW", "Name", "Datum"], inplace=True)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for kw, group in result_df.groupby("KW"):
                group.to_excel(writer, sheet_name=f"KW{kw}", index=False)
        
        output.seek(0)
        st.download_button("Ausgewertete Excel-Datei herunterladen", output, file_name="auswertung.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.warning("Kein passender Name in den Dateien gefunden.")
