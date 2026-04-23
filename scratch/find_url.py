import openpyxl

def find_url():
    wb = openpyxl.load_workbook('e:/Chat-Report-Generator-main/esempioRapporto.xlsx', data_only=True)
    sheet = wb['Chat']
    for i, row in enumerate(sheet.iter_rows(max_row=100, values_only=True)):
        for idx, val in enumerate(row):
            if val and 'https' in str(val):
                print(f"URL found in Row {i}, Col {idx}: {str(val)[:50]}...")

if __name__ == "__main__":
    find_url()
