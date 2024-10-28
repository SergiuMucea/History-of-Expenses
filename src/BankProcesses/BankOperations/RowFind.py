# Find the starting row in the processed table, excluding headers
def find_start_row(ws, d):
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == d:
                return cell.row + 1
                

# last_row will always be the first empty row from the table (in accordance to Python range handling as well)
# rows_no represents the number of rows to be excluded. Some files have additional redundant values in the last rows.
def find_last_row(ws, rows_no=0):    
    last_row = 1
    for row in ws:
        # Checking all cells in a row and not only singular cells
        if not all([cell.value == None for cell in row]):  
            last_row += 1
    return last_row - rows_no

