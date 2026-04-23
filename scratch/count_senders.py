import openpyxl

def analyze():
    wb = openpyxl.load_workbook('e:/Chat-Report-Generator-main/esempioRapporto.xlsx', data_only=True)
    sheet = wb.active
    rows = list(sheet.iter_rows(max_row=100, values_only=True))
    header = [str(x) for x in rows[0]]
    
    da_idx = -1
    for i, h in enumerate(header):
        if h == 'Da': da_idx = i; break
        
    if da_idx == -1: print("Da not found"); return

    senders = {}
    for row in rows[1:]:
        s = str(row[da_idx])
        senders[s] = senders.get(s, 0) + 1
        
    print(f"Total rows: {len(rows)-1}")
    print("Senders found in 'Da' column:")
    for s, count in senders.items():
        print(f"  {s}: {count} messages")

if __name__ == "__main__":
    analyze()
