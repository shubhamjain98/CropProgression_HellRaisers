import openpyxl

def compare_excel_files(file1, file2, output_file, primary_key):
    wb1 = openpyxl.load_workbook(file1)
    wb2 = openpyxl.load_workbook(file2)
    diff_wb = openpyxl.Workbook()
    
    for sheet1, sheet2 in zip(wb1.sheetnames, wb2.sheetnames):
        ws1 = wb1[sheet1]
        ws2 = wb2[sheet2]
        diff_ws = diff_wb.create_sheet(title=sheet2)
        
        for row in range(1, ws2.max_row + 1):
            for col in range(1, ws2.max_column + 1):
                cell1 = ws1.cell(row=row, column=col)
                cell2 = ws2.cell(row=row, column=col)
                diff_cell = diff_ws.cell(row=row, column=col)
                
                if col == primary_key:
                    diff_cell.value = cell2.value
                    continue  # Skip comparison for primary key column
                
                if cell1.value != cell2.value:
                    diff_cell.value = cell2.value
                    diff_cell.font = openpyxl.styles.Font(color="FF0000")
                    diff_cell.comment = openpyxl.comments.Comment(f"Original Value: {cell1.value}", "Author")
                else:
                    diff_cell.value = cell2.value
    
    diff_wb.save(output_file)
    print(f"Differences saved in {output_file}")

# Usage
file1 = "file1.xlsx"
file2 = "file2.xlsx"
output_file = "difference.xlsx"
primary_key = 1  # Assuming primary key column is 1st column (A)

compare_excel_files(file1, file2, output_file, primary_key)
