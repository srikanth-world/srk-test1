import openpyxl

# Load your Excel file
workbook = openpyxl.load_workbook('your_excel_file.xlsx')

# Iterate through each sheet in the workbook
for sheet_name in workbook.sheetnames:
    sheet = workbook[sheet_name]

    # Initialize variables to keep track of row counts
    old_data_row_count = 0
    new_data_row_count = 0

    # Flag to indicate when to start counting new data
    counting_new_data = False

    # Iterate through rows in the sheet
    for row in sheet.iter_rows(min_row=4):
        for cell in row:
            if cell.value == "QA02 - Old":
                old_data_row_count = 0
                counting_new_data = False
            if cell.value == "QA03 - New":
                new_data_row_count = 0
                counting_new_data = True
                continue  # Skip to the next iteration

            if counting_new_data:
                new_data_row_count += 1
            else:
                old_data_row_count += 1

    # Write the row counts at the end of the sheet
    sheet.cell(row=sheet.max_row + 1, column=1, value="Old Data Row Count:")
    sheet.cell(row=sheet.max_row, column=2, value=old_data_row_count)
    sheet.cell(row=sheet.max_row + 2, column=1, value="New Data Row Count:")
    sheet.cell(row=sheet.max_row + 2, column=2, value=new_data_row_count)

# Save the modified workbook
workbook.save('your_updated_excel_file.xlsx')
