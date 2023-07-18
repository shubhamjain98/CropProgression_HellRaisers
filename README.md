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

        # Populate rows_dict1 with rows from the first file
        for row in range(2, ws1.max_row + 1):
            primary_key_values = tuple(ws1.cell(row=row, column=col).value for col in primary_keys)
            rows_dict1[primary_key_values] = [cell.value for cell in ws1[row]]

        # Populate rows_dict2 with rows from the second file
        for row in range(2, ws2.max_row + 1):
            primary_key_values = tuple(ws2.cell(row=row, column=col).value for col in primary_keys)
            rows_dict2[primary_key_values] = [cell.value for cell in ws2[row]]

    # Copy the data from the first workbook to the difference workbook with column names
    for col, cell in enumerate(ws1[1], start=1):
        diff_ws.cell(row=1, column=col, value=cell.value)

    # Compare the cells in both workbooks and highlight the differences
    for row, (primary_key_values, row1) in enumerate(rows_dict1.items(), start=2):
        row2 = rows_dict2.get(primary_key_values)

        if row2:
            diff_row = [primary_key_values[0]] + [None] * (len(row1) - 1)  # Initialize diff_row with primary key value
            for col, (cell1, cell2) in enumerate(zip(row1[1:], row2[1:]), start=1):
                if cell1 != cell2:
                    diff_row[col] = cell2
                    diff_ws.cell(row=row, column=col+1).font = openpyxl.styles.Font(color="FF0000")
                    diff_ws.cell(row=row, column=col+1).comment = openpyxl.comments.Comment(f"Original Value: {cell1}", "Author")
            diff_ws.append(diff_row)

    diff_wb.save(output_file)
    print(f"Differences saved in {output_file}")

# Usage
file1 = "file1.xlsx"
file2 = "file2.xlsx"
output_file = "difference.xlsx"
primary_keys = [1, 2]  # Assuming primary key columns are 1st and 2nd columns (A and B)

compare_excel_files(file1, file2, output_file, primary_keys)
