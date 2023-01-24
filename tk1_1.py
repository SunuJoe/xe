#!/usr/bin/env python3
# making file finder . reference file
import sys
import os
import g_var
import tkinter
from tkinter import * 
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo, showerror
from get_var import *


lay = [] #layout 
root = Tk()
root.title("Test1")
root.resizable(False, False)
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
rt_width = 480
rt_height = 180
x_coor = int((screen_width/2) - (rt_width/2))
y_coor = int((screen_height/2) - (rt_height/2))
root.geometry('{}x{}+{}+{}'.format(rt_width, rt_height, x_coor, y_coor)) 

arg_sel1 = []
arg_sel2 = []
#select_path = ''
#save_path = ''
#pj_code = ''

e1 = Entry(root, width = 3)
e1.place(x=100, y=10) #project code entry
e2text = StringVar() #ref path text var
e3text = StringVar() #raw path text var
e4text = StringVar() #save path text var

e2 = Entry(root, width = 32, state='disabled', textvariable=e2text) # readonly using textvaraiable and set using insert 
e3 = Entry(root, width = 32, state='disabled', textvariable=e3text) 
e4 = Entry(root, width = 32, state='disabled', textvariable=e4text) # readonly using textvaraiable and set using insert 

e2.place(x=100, y=40)
e3.place(x=100, y=73)
e4.place(x=100, y=105)

#label
label_0 = Label(root, text="Project Code").place(x=10, y=10)
label_1 = Label(root, text="Select ref. file").place(x=10, y=42)
label_2 = Label(root, text="Select raw file").place(x=10, y=73)
label_3 = Label(root, text="Save path").place(x=10, y=105)

def select_file(sel_path, text):
#  global select_path
#  global e2text
  filetypes = (
        ('text files', '*.txt'),
        ('excel file', '*.xlsx')
  )
  file_path = fd.askopenfilename(title='Open a file', initialdir='/', filetypes=filetypes)
  if file_path:
    select_path = os.path.abspath(file_path)
    #e2.insert(0, select_path)  #show entry 2 : filepath 
#  print(select_path)
  text.set(select_path)
  sel_path.append(select_path)

def save_directory():
  global save_path
  global e4text
  filetypes = (
        ('excel file', '*.xlsx'),
        ('text file', '*.txt'),
        ('py', '*.py')
  )
  file_path = fd.askdirectory()
  if file_path:
    save_path = os.path.abspath(file_path)
    #e3.insert(0, filepath)
  e4text.set(save_path) #in the entry. show save_path contents
  g_var.save_path.append(save_path) 

top_e1_text = StringVar()
def msg_get_name():
  toplevel = Toplevel(root)
  lay.append(toplevel)
  toplevel.title("File name")
  t_w = 250
  t_h = 100
  x_c = int((screen_width/2) - (t_w/2))
  y_c = int((screen_height/2) - (t_h/2))
  toplevel.geometry('{}x{}+{}+{}'.format(t_w, t_h, x_c, y_c)) 

  top_label = Label(toplevel, text="File name").place(x=10, y=10)
  top_e1 = Entry(toplevel, width = 12)
  top_e1.place(x=100, y=10)
#  top_e1_text = StringVar()
  top_excute_bt = Button(toplevel, text='excute', width=3, command=lambda: run(top_e1.get())).place(x=90, y=60)
#save_button = Button(root, text='3. Save', width=3, command=save_directory).place(x=405, y=104)
  
#def excu():
#  g_var.file_name.append(top_e1.get())
#  print(g_var.file_name) # confirm store
#  try:
#    import x11_1
#  except:
#    showerror(title="Error", message="Exception Error!", parent=root)

def run(text):
  #no select file
#  try:
    g_var.pj_code.append(e1.get())
    g_var.file_name.append(text) 
    print(g_var.pj_code)
    print(g_var.file_name)
    top = lay[0]
    import x11_1
#    top.quit()
    top.destroy()
#    msg_get_name()
#    import x11_1
#  except:
#    showerror(title="Error", message="Exception Error!", parent=root)  
    
#  try:
#    import x11_1
#  except:
#    showerror(title="Error", message="Exception Error!", parent=root)  
#  global pj_code
#  pj_code = e1.get()
#  g_var.pj_code.append(e1.get())  
#  print(get_sel_ref_path())
#  print(get_sel_raw_path())
#  print(get_save_path())
#  print(get_pj_code())
#  print(type(get_pj_code()))
  #file name input popup.
#  os.system('python3 "/Users/jo/learn/3_pylang/xls/x11_1.py"')
  #run x11.py(making excel) 


#e2 = Entry(root, width = 5).place(x=100, y=30)

# button

open1_button = Button(root, text='1. Open', width=3, command=lambda: select_file(g_var.select_ref_path, e2text)).place(x=405, y=37)
open2_button = Button(root, text='2. Open', width=3, command=lambda: select_file(g_var.select_raw_path, e3text)).place(x=405, y=70)
save_button = Button(root, text='3. Save', width=3, command=save_directory).place(x=405, y=104)
run_button = Button(root, text='Run', width=2, command=msg_get_name, fg='blue', activebackground='blue').place(x=410, y=140)


root.mainloop()

