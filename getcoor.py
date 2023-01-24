# getting coordinate functions.

def get_pno_col_ltr(ws):
  for rows in ws[6:7]:
    for cell in rows:
      if str(cell.value).upper() == 'PART-NO':
        pno_col_letter = cell.column_letter
        return pno_col_letter

def get_pname_col_ltr(ws):
  for rows in ws[6:7]:
    for cell in rows:
      if str(cell.value).upper() == 'PART NAME':
        pname_col_letter = cell.column_letter
        return pname_col_letter

def get_upg_col_ltr(ws):
  for rows in ws[6:7]:
    for cell in rows:
      if str(cell.value).upper() == 'UPG':
        upg_col_letter = cell.column_letter
        return upg_col_letter

#charge needs to parameter reference sheet.
def get_charge_col_ltr(ws):
  for cell in ws[1:1]:
#    for cell in rows:
    if str(cell.value).upper() == 'CHARGE':
      charge_col_letter = cell.column_letter
      return charge_col_letter

def get_eono_col_ltr(ws):
  for rows in ws[6:7]:
    for cell in rows:
      if str(cell.value).upper() == 'EO-NO':
        eono_col_letter = cell.column_letter
        return eono_col_letter

#def get_eodate_col_ltr(ws):
#  for rows in ws[6:7]:
#    for cell in rows:
#      if str(cell.value).() == 'eo date':
#        eodate_col_letter = cell.column_letter
#        return eodate_col_letter
 
def get_sup_col_ltr(ws):
  for rows in ws[6:7]:
    for cell in rows:
      if cell.value == '업체명-1':
        sup_col = cell.column
        return sup_col

def get_sup_name(ws, sup_col, get_row):
  if ws.cell(row=get_row, column=sup_col).value:
    return ws.cell(row=get_row, column=sup_col).value
  elif ws.cell(row=get_row, column=(sup_col+1)).value:
    return ws.cell(row=get_row, column=(sup_col+1)).value
  elif ws.cell(row=get_row, column=(sup_col+2)).value:
    return ws. cell(row=get_row, column=(sup_col+2)).value
  else:
    return None
