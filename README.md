import openpyxl

def compare_excel_files(file1, file2, output_file):
    wb1 = openpyxl.load_workbook(file1)
    wb2 = openpyxl.load_workbook(file2)
    diff_wb = openpyxl.Workbook()
    diff_ws = diff_wb.active
    
    for sheet1, sheet2 in zip(wb1.sheetnames, wb2.sheetnames):
        ws1 = wb1[sheet1]
        ws2 = wb2[sheet2]
        diff_ws = diff_wb.create_sheet(title=sheet2)
        
        # Copy data from the second workbook to the difference workbook
        for row in ws2.iter_rows(min_row=1, max_row=ws2.max_row, min_col=1, max_col=ws2.max_column):
            diff_row = [cell.value for cell in row]
            diff_ws.append(diff_row)
        
        # Compare the cells in both workbooks and highlight the differences
        for row in range(1, ws1.max_row + 1):
            for col in range(1, ws1.max_column + 1):
                cell1 = ws1.cell(row=row, column=col)
                cell2 = ws2.cell(row=row, column=col)
                diff_cell = diff_ws.cell(row=row, column=col)
                
                if cell1.value != cell2.value:
                    diff_cell.font = openpyxl.styles.Font(color="FF0000")
                    diff_cell.comment = openpyxl.comments.Comment(f"Original Value: {cell1.value}", "Author")
    
    diff_wb.save(output_file)
    print(f"Differences saved in {output_file}")

# Usage
file1 = "file1.xlsx"
file2 = "file2.xlsx"
output_file = "difference.xlsx"

compare_excel_files(file1, file2, output_file)

