"""This module is the backbone of this project.
All transactions are defined here as categories and based on the category-list, certain labels will be assigned to each transaction.
It needs manual maintanance when new transactions arise, so whenever user considers that some transactions will repeat themselves,
they can update the lists accordingly.

This  module also contains certain functions that just work with the excel file based on the openpyxl module,
They will sometimes be imported and reused in the main modules, just for ease of interraction with the code.

For future developments, machine learning could be implemented here, but we're far from this type of upgrade - 
first, let's learn the basics
"""

int_trans = "Internal transaction"
ext_trans = "External transaction"
man_check = "***Manual check"
dummy_iban = "IBAN"  # Used here only for privacy reasons. Normally replaced by actual IBAN number.

def set_value(dst_ws, dst_row, dst_col, value):
    dst_ws.cell(row=dst_row, column=dst_col).value = value
    
def set_bank(dst_ws, dst_row, bank: str):
    set_value(dst_ws, dst_row, 8, bank)

# Amount of inserted rows will exclude header + last row in the range, which is always empty/redunant.
# The data range will be the number in between and will be inserted from row 2 (excluding header in final file).
def insert_rows(ws, s_last_row, s_start_row):
    ws.insert_rows(2, amount = s_last_row - s_start_row)
    
# The cell with the transaction details
def check_cell(src_ws, src_row, col=2):
    return src_ws.cell(row=src_row, column=col)

# Ignore empty cells
def skip_trans(dst_ws, dst_row):
    if dst_ws.cell(row=dst_row, column=5).value != None:
        return True


# Assings transactions that are related to savings 
# See also function convert_trans for the handling of savings
savings = [
    dummy_iban,
    dummy_iban,
]

def set_savings_transactions(src_ws, dst_ws, src_row, dst_row):
    checked_cell = check_cell(src_ws, src_row)
    
    for item in savings:
        if item.lower() in checked_cell.value.lower().split():
            set_value(dst_ws, dst_row, 5, "Savings")
            set_value(dst_ws, dst_row, 6, "Savings")    

# Turn some of the incoming amounts into credit amounts and move them under expenses, so that balance is done
# Applies mostly to Savings, Second Bank and other types of refunds from suppliers. 
# Only employer benefits and salary kept as "Income"
# Better to have a balance than have the amounts separated. Separation is enough for only actual incomings, such as salaries / bonuses.
unconverted_trans = ["Income"]

def convert_trans(dst_ws, dst_row):
    checked_cell = dst_ws.cell(row=dst_row, column=4)
    dest_cell = dst_ws.cell(row=dst_row, column=3)
    trans_check = dst_ws.cell(row=dst_row, column=5)
    
    for item in unconverted_trans:
        if checked_cell.value == None or item.lower() in trans_check.value.lower().split():
            continue
        dest_cell.value = -abs(checked_cell.value)
        checked_cell.value = None   


# Assigns monthly expenses such as bills that are paid each month.
monthly_expenses = {
    "Supplier1": "Category_supplier_1",
    "Supplier2": "Category_supplier_2",
    "Supplier3": "Category_supplier_3",
    "Provider1": "Category_supplier_1",
    "Provider2": "Category_supplier_2",
    "Provider3": "Category_provider_3",
}

def set_monthly_expenses(src_ws, dst_ws, src_row, dst_row):
    checked_cell = check_cell(src_ws, src_row)
        
    for key, value in monthly_expenses.items():
        if skip_trans(dst_ws, dst_row):
            continue
        
        if key.lower() in checked_cell.value.lower().split():
            set_value(dst_ws, dst_row, 5, "Monthly expenses")
            set_value(dst_ws, dst_row, 6, value)


# This section encompasses all external transactions and assings categories for each transaction
# Manual maintenance on adding the information after new statement processing so that they are assigned automatically next time. 
# This represents one representative word of the transaction that make reference to the paid supplier.
supermarket = [
    "image", "Penny","REALFOODS", "MOBILPAY*DRIEDFRUITS", "AUC", "CRISTIM", "Amma", "PROFI", 
    "DRM", "SUPERMARKET", "Fair", "LIDL", "Carrefour", "CORA", "MEGAIMAGE", "DM", "Selgros",
    "Kaufland", "Auchan",
]

health = [
    "Synlab", "zostalab", "delta", "catena", "naturaliabio", "plafar(bio", "FARMACIA", "MEDIDENT",
    "PLATION*FARMACIATEI", "TEI", "FARMAROM", "QUANTUM", "OFTAPRO", "TEH",
]

car = [
    "ready2wash", "ROVIGNETA", "Directia", "OMV", "MOL", "ROMPETROL", "SOCAR", "LUKOIL", "Tires",
]

sport = [
    "WWW.WORLDCLASS.RO", "INTERSPORT", "SPORTISIMO.RO", "GURU", "Decathlon", "ZUMONT", "ROUMASPORT",
]

clothes = [
    "Spencer", "sephora", "Douglas", "H&M", "TAILOR", "C&M", "ccc.eu", "OLIMPIA", "ZARA", "PENTI", 
    "WAIKIKI", "M&S",
]

electronics = [
    "PAYU*ALTEX.RO", "flanco.ro","GALAXY", "Flanco",
]

house = [
    "PAYU*HORNBACH.RO", "mpy*mobexpert", "LEROY", "HOME", "DEDEMAN", "IKEA", "JYSK",
]

fun = [
    "multiplex", "EP*bilet.ro", "PAYU*IABILET.RO", "STARBUCKS", "MANUFAKTURA", "BIKALDI", "KAFE",
]

restaurants = [
    "domino", "tazz", "food", "Glovo", "Turkiseria", "MESOPOTAMIA", "PIZZERIA",
]

travels = [
    "BGN", "USD", "EUR", "beach", "ISK", "WIZZ", "TOLL",
]

public_transport = [
    "lime", "uber", "Bolt", "STB", "Metrorex",
]

credit_account1 = [dummy_iban]
credit_account2 = [dummy_iban]
presents = []

main_transactions = {
    "Supermarket": supermarket, 
    "Health": health, 
    "Car Expenses": car,  
    "Sport": sport, 
    "Clothes": clothes,
    "Electronics": electronics,
    "House Expenses": house,
    "Fun": fun,
    "Restaurants": restaurants,
    "Travels": travels,
    "Transport Bucharest": public_transport,
    "Gifts": presents,
    "Credit Account1": credit_account1,
    "Credit Account2": credit_account2,
}


# Function to assign a separate exception category based on external transaction 'car' and amount smaller than 25 RON
# Think of the situation where a cup of coffee is bought in a gas station, for example.
def set_exception_car_transactions(src_ws, dst_ws, src_row, dst_row):
    checked_cell = check_cell(src_ws, src_row)
    ex_cell = check_cell(src_ws, src_row, col=3)
    
    for item in car:
        if skip_trans(dst_ws, dst_row):
            continue
        
        if item.lower() in checked_cell.value.lower().split() and ex_cell.value < 25:
            set_value(dst_ws, dst_row, 5, "Expenses")
            set_value(dst_ws, dst_row, 6, "Other Expenses")
            
            
# Main function for the external transactions
# Makes a difference between Travels and Expenses
def set_main_transactions(src_ws, dst_ws, src_row, dst_row):
    checked_cell = check_cell(src_ws, src_row)
    
    for key, value_list in main_transactions.items():
        if skip_trans(dst_ws, dst_row):
            continue
        
        for item in value_list:
            if item.lower() in checked_cell.value.lower().split() and item in travels:
                set_value(dst_ws, dst_row, 5, "Travels")
                set_value(dst_ws, dst_row, 6, key)
                break             
            if item.lower() in checked_cell.value.lower().split() and item not in travels:
                set_value(dst_ws, dst_row, 5, "Expenses")
                set_value(dst_ws, dst_row, 6, key)
                break
            
                
# Assings values for the incoming payments
favorable_trans = {
    dummy_iban: "Salary",
    "referral": "Second Bank benefits",
}

def set_earnings(src_ws, dst_ws, src_row, dst_row):
    checked_cell = check_cell(src_ws, src_row)

    for key, value in favorable_trans.items():
        if skip_trans(dst_ws, dst_row):
            continue
            
        if key.lower() in checked_cell.value.lower().split():
            set_value(dst_ws, dst_row, 5, "Income")
            set_value(dst_ws, dst_row, 6, value)


# Checks cash addition and withdrawal. 
# Amounts above 250 need to be checked manually, because they might have to do with cash monthly payments, or whatever else
cash_trans = {
    "Addition": "Cash Addition",
    "Withdrawal": "Cash Withdrawal",  # Cash withdrawal
}

cash = "Cash"

def set_cash_transactions(src_ws, dst_ws, src_row, dst_row):
    checked_cell = check_cell(src_ws, src_row)
    amount = check_cell(src_ws, src_row, col=3)
    
    for key, value in cash_trans.items():
        if skip_trans(dst_ws, dst_row):
            continue
        
        if key.lower() in checked_cell.value.lower().split() and key == "Addition":
            set_value(dst_ws, dst_row, 5, "Expenses")    
            set_value(dst_ws, dst_row, 6, cash)
            break
        
        if key.lower() in checked_cell.value.lower().split() and key == "Withdrawal":
            if amount.value != None and amount.value < 250:
                set_value(dst_ws, dst_row, 5, "Expenses") 
                set_value(dst_ws, dst_row, 6, cash)
                break
            else:
                set_value(dst_ws, dst_row, 5, man_check)    
                set_value(dst_ws, dst_row, 6, value)
                
                
# Assigns the rest values based on the main transactions dictionary. 
# This function is last to be called, so that already assigned values are not overriden.
# It looks at values that are not present in any of the main transactions dictionaries and the result is simple:
# Amounts below 100 RON will be assigned "other expenses", and above 100 RON will be checked manually.
def set_rest_transactions(src_ws, dst_ws, src_row, dst_row):
    checked_cell = check_cell(src_ws, src_row)
    ex_cell = check_cell(src_ws, src_row, col=3)
    
    for key, value_list in main_transactions.items():
        if skip_trans(dst_ws, dst_row):
            continue
        
        for item in value_list:
            if item.lower() not in checked_cell.value.lower().split() and ex_cell.value != None and ex_cell.value < 100:                    
                set_value(dst_ws, dst_row, 5, "Expenses")
                set_value(dst_ws, dst_row, 6, "Other Expenses")
                break   
            else:
                set_value(dst_ws, dst_row, 5, man_check)
                set_value(dst_ws, dst_row, 6, man_check)


def set_transactions(src_ws, dst_ws, src_row, dst_row):
    set_savings_transactions(src_ws, dst_ws, src_row, dst_row)
    set_monthly_expenses(src_ws, dst_ws, src_row, dst_row)
    set_exception_car_transactions(src_ws, dst_ws, src_row, dst_row)
    set_main_transactions(src_ws, dst_ws, src_row, dst_row)
    set_earnings(src_ws, dst_ws, src_row, dst_row)
    set_cash_transactions(src_ws, dst_ws, src_row, dst_row)
    set_rest_transactions(src_ws, dst_ws, src_row, dst_row)



# Deletes internal transactions between personal accounts (redundant transactions). 
# They will be accounted for in the overall transactions checks
excluded_trans = [
    dummy_iban,
    dummy_iban,
    dummy_iban,
    "Reference",
    "Reference2",
]

def remove_redundant(src_ws, dst_ws, src_row, dst_row):
    checked_cell = check_cell(src_ws, src_row)
    removed_rows = []

    for item in excluded_trans:
        if checked_cell.value != None and item.lower() in checked_cell.value.lower().split():
            removed_rows.append(dst_row)
            dst_ws.delete_rows(dst_row)
    
    return removed_rows
            

def last_procedures(src_ws, dst_ws, src_row, dst_row):
    convert_trans(dst_ws, dst_row)
    remove_redundant(src_ws, dst_ws, src_row, dst_row)

