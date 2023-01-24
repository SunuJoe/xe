# check_unique function

def check_unique(ws, p_row, code):
  if ws['E'+str(p_row)].value[-5:-3] == code:
    return True
  else:
    return False
