import openpyxl

def compare_excel_files(file1, file2, output_file, primary_key):
    wb1 = openpyxl.load_workbook(file1)
    wb2 = openpyxl.load_workbook(file2)
    diff_wb = openpyxl.Workbook()
    diff_ws = diff_wb.active

    # Create dictionaries to store rows based on the primary key for both files
    rows_dict1 = {}
    rows_dict2 = {}

    for sheet1, sheet2 in zip(wb1.sheetnames, wb2.sheetnames):
        ws1 = wb1[sheet1]
        ws2 = wb2[sheet2]

        # Populate rows_dict1 with rows from the first file
        for row in range(2, ws1.max_row + 1):
            cell = ws1.cell(row=row, column=primary_key)
            rows_dict1[cell.value] = [cell.value] + [cell2.value for cell2 in ws1[row]]

        # Populate rows_dict2 with rows from the second file
        for row in range(2, ws2.max_row + 1):
            cell = ws2.cell(row=row, column=primary_key)
            rows_dict2[cell.value] = [cell.value] + [cell2.value for cell2 in ws2[row]]

    # Copy the data from the first workbook to the difference workbook with column names
    diff_ws.append([cell.value for cell in ws2[1]])

    for primary_key_value, row2 in rows_dict2.items():
        row1 = rows_dict1.get(primary_key_value)

        if row1:
            diff_ws.append(row2 if row2[1:] != row1[1:] else row1)

    # Compare the cells in both workbooks and highlight the differences
    for row in range(2, diff_ws.max_row + 1):
        for col in range(1, diff_ws.max_column + 1):
            cell1 = rows_dict1.get(diff_ws.cell(row=row, column=primary_key).value)
            cell2 = diff_ws.cell(row=row, column=col)

            if cell1 and cell1[col - 1] != cell2.value:
                cell2.font = openpyxl.styles.Font(color="FF0000")
                cell2.comment = openpyxl.comments.Comment(f"Original Value: {cell1[col - 1]}", "Author")

    diff_wb.save(output_file)
    print(f"Differences saved in {output_file}")

# Usage
file1 = "file1.xlsx"
file2 = "file2.xlsx"
output_file = "difference.xlsx"
primary_key = 1  # Assuming primary key column is 1st column (A)

compare_excel_files(file1, file2, output_file, primary_key)

