import openpyxl
import re

def find_media():
    wb = openpyxl.load_workbook('e:/Chat-Report-Generator-main/esempioRapporto.xlsx', data_only=True)
    sheet = wb['Chat']
    exts = r"\.(jpg|jpeg|png|gif|mp4|webm|mov|avi|wav|mp3|opus|amr)"
    for i, row in enumerate(sheet.iter_rows(max_row=500, values_only=True)):
        for idx, val in enumerate(row):
            if val and re.search(exts, str(val), re.I):
                print(f"Media found! Row {i}, Col {idx}: {str(val)[:100]}")

if __name__ == "__main__":
    find_media()
