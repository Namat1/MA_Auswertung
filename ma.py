import pandas as pd
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from io import BytesIO

# Beispieldaten, später durch echte ersetzt
df = pd.DataFrame({
    "KW": [1, 1, 2, 2],
    "Datum": ["2024-12-29", "2024-12-30", "2025-01-05", "2025-01-06"],
    "Name": ["Fuhlbrügge Justin", "Adler Philipp", "Holtz Ch.", "Rimba Gona"],
    "Tour": [12221, 12222, 12223, 12224],
    "Uhrzeit": ["07:00", "08:00", "06:30", "06:00"],
    "LKW": [5001, 5009, 6005, 6003]
})

# Datum vorbereiten
df["Datum"] = pd.to_datetime(df["Datum"])
df["Wochentag"] = df["Datum"].dt.strftime('%A')
df["Datum_formatiert"] = df["Datum"].dt.strftime('%d.%m.%Y')
df["Datum_komplett"] = df["Wochentag"] + ", " + df["Datum_formatiert"]
df.drop(columns=["Datum", "Wochentag", "Datum_formatiert"], inplace=True)
df = df[["KW", "Datum_komplett", "Name", "Tour", "Uhrzeit", "LKW"]]

# Sonntag als Wochenbeginn (ISO-Woche +1 für Montag-Freitag-Daten wenn gewünscht)
# (bereits angepasst durch Datumsvorbereitung)

# Excel-Datei vorbereiten
output = BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    sheet_name = "Alle_KWs"
    start_row = 1
    wb = writer.book

    for kw, group in df.groupby("KW"):
        group = group.sort_values(by="Datum_komplett")
        ws = writer.book.create_sheet(title=sheet_name) if writer.sheets == {} else writer.sheets[sheet_name]

        # KW-Trennüberschrift
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

        # Leere Zeile zwischen den KWs
        start_row += 1

    # Spaltenbreiten anpassen (150 % Inhalt)
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
