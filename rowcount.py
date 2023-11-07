import openpyxl

def count_rows_and_write_to_sheet(file_path):
    workbook = openpyxl.load_workbook(file_path)

    for sheet_name in workbook.sheetnames:
        worksheet = workbook[sheet_name]
        count_old = 0
        count_new = 0
        flag = None

        for row in worksheet.iter_rows():
            for cell in row:
                if cell.value == "QA02 - Old":
                    flag = "QA02 - Old"
                elif cell.value == "QA03 - New":
                    flag = "QA03 - New"
                elif flag == "QA02 - Old":
                    count_old += 1
                elif flag == "QA03 - New":
                    count_new += 1

        last_row = worksheet.max_row
        worksheet.cell(row=last_row + 1, column=1, value="Count of QA02 - Old Rows")
        worksheet.cell(row=last_row + 1, column=2, value=count_old)
        worksheet.cell(row=last_row + 2, column=1, value="Count of QA03 - New Rows")
        worksheet.cell(row=last_row + 2, column=2, value=count_new)

    workbook.save(file_path)

# Replace 'your_file.xlsx' with the path to your Excel file
count_rows_and_write_to_sheet('your_file.xlsx')
