import openpyxl

def analyze_sheet(name):
    wb = openpyxl.load_workbook('e:/Chat-Report-Generator-main/esempioRapporto.xlsx', data_only=True)
    sheet = wb[name]
    print(f"--- {name} ---")
    for i, row in enumerate(sheet.iter_rows(max_row=10, values_only=True)):
        r = [str(x) for x in row if x]
        if i == 0: print(f"Headers: {r[:20]}")
        elif i < 5: print(f"Row {i}: {r[:10]}")

analyze_sheet('Chat')
analyze_sheet('Riepilogo')
