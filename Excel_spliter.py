import openpyxl

def split_excel_file(file_path):
    # Load the Excel file
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    # Extract data from each row and column
    data = []
    for row in sheet.iter_rows(values_only=True):
        row_data = []
        for cell in row:
            row_data.append(cell)
        data.append(row_data)

    return data

# Provide the path to the Excel file
file_path = 'WORK SHOP CERTIFICATE.xlsx'

# Split the Excel file into rows and columns
result = split_excel_file(file_path)
print(result[1][1])
print(result[1][3])