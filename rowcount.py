import openpyxl

def count_data_rows_and_write_to_workbook(file_path):
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

        worksheet.append(["Count of QA02 - Old Rows", count_old])
        worksheet.append(["Count of QA03 - New Rows", count_new])

    workbook.save(file_path)

# Replace 'your_file.xlsx' with the path to your Excel file
count_data_rows_and_write_to_workbook('your_file.xlsx')
