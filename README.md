import openpyxl

def compare_excel_files(file1, file2, output_file, primary_key):
    wb1 = openpyxl.load_workbook(file1)
    wb2 = openpyxl.load_workbook(file2)
    diff_wb = openpyxl.Workbook()
    diff_ws = diff_wb.active

    # Create dictionary to store rows based on the primary key for the first file
    rows_dict1 = {}

    for sheet1 in wb1.sheetnames:
        ws1 = wb1[sheet1]

        # Populate rows_dict1 with rows from the first file
        for row in range(2, ws1.max_row + 1):
            cell = ws1.cell(row=row, column=primary_key)
            rows_dict1[cell.value] = [cell.value] + [cell2.value for cell2 in ws1[row]]

    for sheet2 in wb2.sheetnames:
        ws2 = wb2[sheet2]

        # Copy the data from the second workbook to the difference workbook with column names
        diff_ws.append([cell.value for cell in ws2[1]])

        # Compare the cells in both workbooks and highlight the differences
        for row in range(2, ws2.max_row + 1):
            primary_key_value = ws2.cell(row=row, column=primary_key).value
            row2 = [cell.value for cell in ws2[row]]
            row1 = rows_dict1.get(primary_key_value)

            if row1:
                if row1[1:] != row2[1:]:
                    diff_row = [None] * len(row2)
                    for col in range(1, len(row2)):
                        if row1[col] != row2[col]:
                            diff_row[col] = row2[col]
                            cell = diff_ws.cell(row=row, column=col+1)
                            cell.font = openpyxl.styles.Font(color="FF0000")
                            cell.comment = openpyxl.comments.Comment(f"Original Value: {row1[col]}", "Author")
                    diff_ws.append(diff_row)
            else:
                diff_ws.append(row2)

    diff_wb.save(output_file)
    print(f"Differences saved in {output_file}")

# Usage
file1 = "file1.xlsx"
file2 = "file2.xlsx"
output_file = "difference.xlsx"
primary_key = 1  # Assuming primary key column is 1st column (A)

compare_excel_files(file1, file2, output_file, primary_key)
