import openpyxl

def analyze_attachments():
    wb = openpyxl.load_workbook('e:/Chat-Report-Generator-main/esempioRapporto.xlsx', data_only=True)
    sheet = wb['Chat']
    rows = list(sheet.iter_rows(max_row=20, values_only=True))
    header = [str(x) for x in rows[1]]
    
    indices = [i for i, h in enumerate(header) if 'Allegato #' in h and 'Dettagli' not in h]
    print(f"Allegato columns: {indices}")
    
    for i, row in enumerate(rows[2:]):
        found = []
        for idx in indices:
            val = row[idx]
            if val: found.append(f"Col {idx} ({header[idx]}): {val}")
        if found:
            print(f"Row {i+2}: {found}")

if __name__ == "__main__":
    analyze_attachments()
