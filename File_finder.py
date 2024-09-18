import os
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
def list_files_by_time(directory):
    # Get a list of all files in the directory
    files = os.listdir(directory)

    # Create a list of file paths with their corresponding modification times
    file_times = [(os.path.join(directory, file), os.path.getmtime(os.path.join(directory, file))) for file in files]

    # Sort the list of file paths by modification time (oldest files first)
    sorted_files = sorted(file_times, key=lambda x: x[1])

    # Create an array to store the file paths
    file_paths = []

    # Append the file paths to the array
    for file_path, modification_time in sorted_files:
        file_paths.append(file_path)

    return file_paths
# Provide the path to the Excel file
file_path = 'check.xlsx'
result = split_excel_file(file_path)
folder_path = r'C:\Users\Admin\PycharmProjects\pythonProject\Bluck_email\Final_Certificate'
sorted_file_paths = list_files_by_time(folder_path)
for i in range(1,3):
    path= sorted_file_paths[i-1]
    name= result[i][1]
    email=result[i][3]
    Send=f'name={name} ad email={email} and path ={path}'
    print(Send)

