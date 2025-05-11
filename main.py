from tkinter import *
from tkinter import messagebox
from tkinter import ttk,filedialog
import os
import pandas

root=Tk()
root.title("Excel Datasheet")
root.geometry("1100x400+200+200")

def Open():
    
    filename = filedialog.askopenfilename(title="Open a file", filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
    
    if filename:
        try:
            filename=r"{}".format(filename)
            df=pandas.read_excel(filename)
        except:
            messagebox.showerror("Error","You can't access this file")

    tree.delete(*tree.get_children())

    #data sheet heading
    tree["column"]=list(df.columns)
    tree["show"]="headings"
    #heading title
    for col in tree ["column"]:
        tree.heading(col,text=col)

    #entry
    df_rows=df.to_numpy().tolist()
    for row in df_rows:
        tree.insert("","end",values=row)     


#icon
image_icon=PhotoImage(file="logo.png")
root.iconphoto(False,image_icon)

#frame
frame=Frame(root)
frame.pack(pady=25)

#treeview
tree=ttk.Treeview(frame)
tree.pack()

#button
button=Button(root,text="OPEN",width=60,height=2,font=30,fg="white",bg="#0078d7",command=Open)
button.pack(padx=10,pady=20)



root.mainloop()