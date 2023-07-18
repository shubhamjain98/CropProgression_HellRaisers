import openpyxl

def compare_excel_files(file1, file2, output_file, primary_keys):
    wb1 = openpyxl.load_workbook(file1)
    wb2 = openpyxl.load_workbook(file2)
    diff_wb = openpyxl.Workbook()
    diff_ws = diff_wb.active

    # Create dictionaries to store rows based on the primary keys for both files
    rows_dict1 = {}
    rows_dict2 = {}

    for sheet1, sheet2 in zip(wb1.sheetnames, wb2.sheetnames):
        ws1 = wb1[sheet1]
        ws2 = wb2[sheet2]

        # Get the column names
        column_names = [cell.value for cell in ws1[1]]
        diff_ws.append(column_names)

        # Populate rows_dict1 with rows from the first file
        for row in range(2, ws1.max_row + 1):
            primary_key_values = tuple(ws1.cell(row=row, column=col).value for col in primary_keys)
            rows_dict1[primary_key_values] = [cell.value for cell in ws1[row]]

        # Populate rows_dict2 with rows from the second file
        for row in range(2, ws2.max_row + 1):
            primary_key_values = tuple(ws2.cell(row=row, column=col).value for col in primary_keys)
            rows_dict2[primary_key_values] = [cell.value for cell in ws2[row]]

    # Compare the cells in both workbooks and highlight the differences
    for primary_key_values, row1 in rows_dict1.items():
        row2 = rows_dict2.get(primary_key_values)

        if row2:
            diff_row = []
            for cell1, cell2 in zip(row1, row2):
                if cell1 != cell2:
                    diff_row.append(cell2)
                    cell = diff_ws.cell(row=len(diff_ws["A"]) + 1, column=len(diff_row))
                    cell.font = openpyxl.styles.Font(color="FF0000")
                    cell.comment = openpyxl.comments.Comment(f"Original Value: {cell1}", "Author")
                else:
                    diff_row.append(None)
        else:
            diff_row = row1 + [None] * (len(column_names) - len(row1))

        diff_ws.append(diff_row)

    diff_wb.save(output_file)
    print(f"Differences saved in {output_file}")

# Usage
file1 = "file1.xlsx"
file2 = "file2.xlsx"
output_file = "difference.xlsx"
primary_keys = [1, 2]  # Assuming primary key columns are 1st and 2nd columns (A and B)

compare_excel_files(file1, file2, output_file, primary_keys)
