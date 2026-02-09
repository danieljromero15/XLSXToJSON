import openpyxl
import sys
import json

if __name__ == '__main__':
    sheet_number = 0
    indent = 2
    if len(sys.argv) not in [2, 3, 4]:
        print("Usage: python XLSXToJSON.py <XLSXFile.xlsx> [sheet_number = 0] [indent = 2]")
        exit(1)

    if len(sys.argv) >= 3:
        try:
            sheet_number = int(sys.argv[2])
        except ValueError:
            print("Invalid sheet number")
            exit(1)

    if len(sys.argv) >= 4:
        try:
            indent = int(sys.argv[3])
        except ValueError:
            print("Invalid indent")
            exit(1)

    wb = openpyxl.load_workbook(sys.argv[1])
    ws = wb.worksheets[sheet_number]
    print(ws)

    max_col = 0
    headers = []
    json_list = []
    for row_num, row in enumerate(ws.iter_rows()):
        row_dict = {}
        for cell in row:
            if cell.value is not None:
                if row_num == 0:
                    headers.append(cell.value)
                    if max_col < cell.column:
                        max_col = cell.column
                else:
                    row_dict[headers[cell.column - 1]] = cell.value
        if row_dict:
            json_list.append(row_dict)

    # print(json_list)
    with open('output.json', 'w') as json_file:
        json.dump(json_list, json_file, indent=indent)

    print("output.json created")
