import g_var

#def put_sel_ref_path(path):
#  g_var.select_path = path

def get_sel_ref_path():
  return g_var.select_ref_path[0]

#def put_sel_raw_path(path):
#  g_var.select_path = path

def get_sel_raw_path():
  return g_var.select_raw_path[0]

def put_save_path(path):
  g_var.save_path = path

def get_save_path():
  return g_var.save_path[0]

#def put_pj_code(code):
#  g_var.pj_code = code

def get_pj_code():
  return g_var.pj_code[0] 
