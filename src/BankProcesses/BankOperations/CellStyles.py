# Small module for formatting ans styling cells in the final file
def format_date(dst_ws):
    for cell in dst_ws["A"]:
        if cell.value == "Data":
            continue
        if cell.number_format != 'DD/MM/YYYY':
            cell.number_format = 'DD/MM/YYYY'
        cell.alignment = cell.alignment.copy(horizontal='right', vertical='center')

def set_accounting_format(dst_ws):
    acc = '#,###0.00'
    for cell in dst_ws['C']:
        if cell.number_format != acc:
            cell.number_format = acc
    for cell in dst_ws['D']:
        if cell.number_format != acc:
            cell.number_format = acc

def wrap(dst_ws):
    for cell in dst_ws['B']:
        if cell.alignment.copy(wrap_text=False):
            cell.alignment = cell.alignment.copy(wrap_text=True)