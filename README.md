# CropProgression_HellRaisers
import openpyxl

def compare_excel_files(file1, file2, output_file, primary_key):
    wb1 = openpyxl.load_workbook(file1)
    wb2 = openpyxl.load_workbook(file2)
    diff_wb = openpyxl.Workbook()
    diff_sheet = diff_wb.active
    
    for sheet1, sheet2, diff_sheet in zip(wb1.sheetnames, wb2.sheetnames, diff_wb.sheetnames):
        ws1 = wb1[sheet1]
        ws2 = wb2[sheet2]
        diff_ws = diff_wb[diff_sheet]
        
        # Create a dictionary to store rows based on the primary key
        rows_dict = {}
        for row in range(2, ws2.max_row + 1):
            cell = ws2.cell(row=row, column=primary_key)
            rows_dict[cell.value] = ws2[row]
        
        # Compare rows in the first file with matching rows in the second file based on the primary key
        for row in range(2, ws1.max_row + 1):
            cell1 = ws1.cell(row=row, column=primary_key)
            matching_row = rows_dict.get(cell1.value)
            
            if matching_row:
                for col in range(1, ws1.max_column + 1):
                    diff_cell = diff_ws.cell(row=row, column=col)
                    cell1 = ws1.cell(row=row, column=col)
                    cell2 = matching_row[col - 1]
                    
                    if cell1.value != cell2.value:
                        diff_cell.value = cell1.value
                        diff_cell.font = openpyxl.styles.Font(color="FF0000")
                        diff_cell.comment = openpyxl.comments.Comment(f"Actual Value: {cell1.value}", "Author")
    
    diff_wb.save(output_file)
    print(f"Differences saved in {output_file}")

# Usage
file1 = "file1.xlsx"
file2 = "file2.xlsx"
output_file = "difference.xlsx"
primary_key = 1  # Assuming primary key column is 1st column (A)

compare_excel_files(file1, file2, output_file, primary_key)
