import openpyxl

def find_header():
    wb = openpyxl.load_workbook('e:/Chat-Report-Generator-main/esempioRapporto.xlsx', data_only=True)
    sheet = wb.active
    for i, row in enumerate(sheet.iter_rows(max_row=20, values_only=True)):
        r = [str(x) for x in row if x]
        print(f"Row {i}: {r[:10]}")
        if 'Da' in r or 'A' in r or 'Corpo' in r:
            print(f"!!! HEADER FOUND AT ROW {i} !!!")

if __name__ == "__main__":
    find_header()
