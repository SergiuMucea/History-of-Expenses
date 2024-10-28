"""This module gathers all downloaded attachments from gmail
and consolidates the information from their pdf format to excel.
It is only needed for the Third bank, where we have statements only in pdf format.

Consists of actions such as:
- Converting tables from pdf to excel
- Moving converted files from downloads folder to archive
- Checking if there are duplicates and if yes, moving them to the duplicates folder
- Consolidating and cleaning up information for further processing in the ThirdBank module 
"""

import aspose.pdf as ap
from tkinter import filedialog as fd
from pathlib import Path
import os

input_folder = Path("path/to/the/Project/InputFiles/Statements/")
archive_folder = Path("path/to/the/Project/InputFiles/Statements/Archive/")
data_folder = Path("path/to/the/Project/OutputFiles/")
bank_statement = data_folder / "Statements.xlsx"


# After downloading the files from Gmail, they will be selected for processing using this function
def file_selection():
    input_statements = []

    for file in os.listdir(input_folder):
        if file.endswith(".pdf"):
            input_statements.append(file)
            
    return input_statements

# Convert the tables from the pdf file to excel and save the row_index for multiple file processing
# Uses updated_row_index in case of multiple file processing
def process_tables(ws, document, updated_row_index):
    # Initialize row index
    row_index = 1 + updated_row_index
    
    # Iterate through all pages of the document
    for page in document.pages:
        absorber = ap.text.TableAbsorber()
        absorber.visit(page)
                
        # Extract tables from the current page
        for table in absorber.table_list:
            for row in table.row_list:
                col_index = 1
                for cell in row.cell_list:
                    text_fragment_collection = cell.text_fragments
                    cell_text = ""
                    for fragment in text_fragment_collection:
                        cell_text += fragment.text + '\n'  # Add a new line between text fragments
                    ws.cell(row=row_index, column=col_index, value=cell_text.strip())  # Remove trailing space
                    col_index += 1
                row_index += 1
    return row_index
                

# Deletes redundant rows that are not related to transactions
def delete_redundant_rows(ws, keyword):
    rows_to_delete = []
        
    for row in ws.iter_rows():
        checked_cell = row[0]
        cell_value = str(checked_cell.value) if checked_cell.value else ""
        
        if checked_cell.value == keyword:
            continue
        if not cell_value or not str(checked_cell.value)[0].isdigit():
            rows_to_delete.append(checked_cell.row)

    for row_num in reversed(rows_to_delete):
        ws.delete_rows(row_num)

# In case of multiple pages, the header of the table will duplicate itself, so this function corrects this
def delete_duplicate_header(ws, keyword):
    rows_to_delete = []
    
    for row in ws.iter_rows(min_row=2):
        checked_cell = row[0]
        if checked_cell.value == keyword:
            rows_to_delete.append(checked_cell.row)

    for row_num in reversed(rows_to_delete):
        ws.delete_rows(row_num)
        
# Replaces symbols from the converted file with whitespace, 
# in order for the converted file to interract properly with the script
def fix_transaction_details(ws):
    symbols = ['<', '>', ',', ';', '/', '\\', '-']
    for symbol in symbols:
        for row in ws.iter_rows(min_row=2):
            checked_cell = row[1]
            if checked_cell.value != None:
                checked_cell.value = str(checked_cell.value.replace(symbol, ' '))
    
def fix_amounts_format(ws):
    for row in ws.iter_rows(min_row=2):
        checked_cell = row[2]
        if checked_cell.value != None:
            checked_cell.value = float(checked_cell.value.replace('.', '').replace(',', '.'))
            
# Settles proper arrangment of amounts in table. Expenses are debit and earnings are credit.
def fix_credit_debit(ws):
    for row in ws.iter_rows(min_row=2):
        checked_cell = row[2]
        if checked_cell.value != None:
            if checked_cell.value > 0:
                checked_cell.value = -abs(checked_cell.value)
            elif checked_cell.value < 0:
                checked_cell.value = abs(checked_cell.value)

# Small amounts which are normally comissions from the bank will be removed,
# so that those statements are ignored from processing
def remove_small_amounts(ws):
    rows_to_delete = []
    for row in ws.iter_rows(min_row=2):
        checked_cell = row[2]
        if checked_cell.value != None:
            if checked_cell.value > -5 and checked_cell.value < 5:
                rows_to_delete.append(checked_cell.row)
                
    for row_num in reversed(rows_to_delete):
        ws.delete_rows(row_num)


# Main function which prepares the statement file for processing in the main ThirdBank module
def pdf_conversion(wb, ws, keyword):

    input_statements = file_selection()
    
    # Necessary to define before the for loop, in order to avoid resetting its value too early within the loop
    updated_row_index = 0
    
    for file in input_statements:
        input_file = str(input_folder / file)
        archived_file = str(archive_folder / file)
        
        # Open PDF document
        document = ap.Document(input_file, password="********")
        
        # Call the process_tables function and use the returned row_index for saving its value for next processed statement
        # updated_row_index resets itself after each use, for consistent and correct use of the row_index value
        row_index = process_tables(ws, document, updated_row_index)
        updated_row_index = 0
        updated_row_index += row_index
        
        # Moving processed file to archive
        os.renames(input_file, archived_file)
            
    # Preparing the statement file for correct and clear processing              
    ws.delete_cols(3)
    delete_redundant_rows(ws, keyword)
    delete_duplicate_header(ws, keyword)
    fix_transaction_details(ws)
    fix_amounts_format(ws)
    fix_credit_debit(ws)
    remove_small_amounts(ws)
    
    wb.save(bank_statement)