import openpyxl
from tkinter import *
import tkinter.ttk as ttk
import tkinter.messagebox as mb
import tkinter.simpledialog as sd
import os

# Ensure folder exists
folder = r"C:\python"
if not os.path.exists(folder):
    os.makedirs(folder)

# Excel file setup (always in C:\python)
file = os.path.join(folder, "library.xlsx")

# If file does not exist, create it with headers
if not os.path.exists(file):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Library"
    ws.append(["Book Name", "Book ID", "Author", "Status", "Issuer Card ID"])
    wb.save(file)
    mb.showinfo("File Created", f"{file} was created automatically!")

# Load workbook
wb = openpyxl.load_workbook(file)
ws = wb["Library"]

# Functions
def issuer_card():
    Cid = sd.askstring('Issuer Card ID', 'What is the Issuer\'s Card ID?\t\t\t')
    if not Cid:
        mb.showerror('Error', 'Issuer ID cannot be empty!')
    else:
        return Cid

def display_records():
    tree.delete(*tree.get_children())
    for row in ws.iter_rows(min_row=2, values_only=True):
        tree.insert('', END, values=row)

def clear_fields():
    bk_status.set('Available')
    bk_name.set('')
    bk_id.set('')
    author_name.set('')
    card_id.set('')
    try:
        tree.selection_remove(tree.selection()[0])
    except:
        pass

def clear_and_display():
    clear_fields()
    display_records()

def add_record():
    if bk_status.get() == 'Issued':
        card_id.set(issuer_card())
    else:
        card_id.set('N/A')

    # Check duplicate Book ID
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[1] == bk_id.get():
            mb.showerror('Error', 'Book ID already exists!')
            return

    ws.append([bk_name.get(), bk_id.get(), author_name.get(), bk_status.get(), card_id.get()])
    wb.save(file)
    clear_and_display()
    mb.showinfo('Success', 'Record added to Excel!')

def view_record():
    if not tree.focus():
        mb.showerror('Error', 'Please select a record to view')
        return
    current_item = tree.focus()
    values = tree.item(current_item)["values"]

    bk_name.set(values[0])
    bk_id.set(values[1])
    author_name.set(values[2])
    bk_status.set(values[3])
    card_id.set(values[4])

def update_record():
    def update():
        if bk_status.get() == 'Issued':
            card_id.set(issuer_card())
        else:
            card_id.set('N/A')

        # Find and update row in Excel
        for row in ws.iter_rows(min_row=2):
            if row[1].value == bk_id.get():
                row[0].value = bk_name.get()
                row[2].value = author_name.get()
                row[3].value = bk_status.get()
                row[4].value = card_id.get()
                break
        wb.save(file)
        clear_and_display()
        edit.destroy()
        bk_id_entry.config(state='normal')
        clear.config(state='normal')

    view_record()
    bk_id_entry.config(state='disable')
    clear.config(state='disable')

    edit = Button(left_frame, text='Update Record', font=btn_font, bg=btn_hlb_bg, width=20, command=update)
    edit.place(x=50, y=375)

def remove_record():
    if not tree.selection():
        mb.showerror('Error!', 'Please select a record to delete')
        return

    current_item = tree.focus()
    values = tree.item(current_item)["values"]

    # Delete from Excel
    for row in ws.iter_rows(min_row=2):
        if row[1].value == values[1]:
            ws.delete_rows(row[0].row)
            break
    wb.save(file)
    clear_and_display()
    mb.showinfo('Done', 'Record deleted successfully!')

def delete_inventory():
    if mb.askyesno('Confirm', 'Are you sure you want to delete the entire inventory?'):
        ws.delete_rows(2, ws.max_row)
        wb.save(file)
        clear_and_display()

def change_availability():
    if not tree.selection():
        mb.showerror('Error!', 'Please select a record')
        return

    current_item = tree.focus()
    values = tree.item(current_item)["values"]
    book_id = values[1]
    status = values[3]

    for row in ws.iter_rows(min_row=2):
        if row[1].value == book_id:
            if status == "Issued":
                surety = mb.askyesno('Confirm Return', 'Has the book been returned?')
                if surety:
                    row[3].value = "Available"
                    row[4].value = "N/A"
                else:
                    mb.showinfo('Info', 'Book must be returned first.')
            else:
                row[3].value = "Issued"
                row[4].value = issuer_card()
            break

    wb.save(file)
    clear_and_display()

# GUI Setup
lf_bg = 'LightSkyBlue'
rtf_bg = 'DeepSkyBlue'
rbf_bg = 'DodgerBlue'
btn_hlb_bg = 'SteelBlue'

lbl_font = ('Georgia', 13)
entry_font = ('Times New Roman', 12)
btn_font = ('Gill Sans MT', 13)

root = Tk()
root.title('Library Management System (Excel)')
root.geometry('1010x530')
root.resizable(0, 0)

Label(root, text='LIBRARY MANAGEMENT SYSTEM', font=("Noto Sans CJK TC", 15, 'bold'),
      bg=btn_hlb_bg, fg='White').pack(side=TOP, fill=X)

# StringVars
bk_status = StringVar()
bk_name = StringVar()
bk_id = StringVar()
author_name = StringVar()
card_id = StringVar()

# Frames
left_frame = Frame(root, bg=lf_bg)
left_frame.place(x=0, y=30, relwidth=0.3, relheight=0.96)

RT_frame = Frame(root, bg=rtf_bg)
RT_frame.place(relx=0.3, y=30, relheight=0.2, relwidth=0.7)

RB_frame = Frame(root)
RB_frame.place(relx=0.3, rely=0.24, relheight=0.785, relwidth=0.7)

# Left Frame
Label(left_frame, text='Book Name', bg=lf_bg, font=lbl_font).place(x=98, y=25)
Entry(left_frame, width=25, font=entry_font, textvariable=bk_name).place(x=45, y=55)

Label(left_frame, text='Book ID', bg=lf_bg, font=lbl_font).place(x=110, y=105)
bk_id_entry = Entry(left_frame, width=25, font=entry_font, textvariable=bk_id)
bk_id_entry.place(x=45, y=135)

Label(left_frame, text='Author Name', bg=lf_bg, font=lbl_font).place(x=90, y=185)
Entry(left_frame, width=25, font=entry_font, textvariable=author_name).place(x=45, y=215)

Label(left_frame, text='Status of the Book', bg=lf_bg, font=lbl_font).place(x=75, y=265)
dd = OptionMenu(left_frame, bk_status, *['Available', 'Issued'])
dd.configure(font=entry_font, width=12)
dd.place(x=75, y=300)

submit = Button(left_frame, text='Add new record', font=btn_font, bg=btn_hlb_bg, width=20, command=add_record)
submit.place(x=50, y=375)

clear = Button(left_frame, text='Clear fields', font=btn_font, bg=btn_hlb_bg, width=20, command=clear_fields)
clear.place(x=50, y=435)

# Right Top Frame
Button(RT_frame, text='Delete book record', font=btn_font, bg=btn_hlb_bg, width=17, command=remove_record).place(x=8, y=30)
Button(RT_frame, text='Delete full inventory', font=btn_font, bg=btn_hlb_bg, width=17, command=delete_inventory).place(x=178, y=30)
Button(RT_frame, text='Update book details', font=btn_font, bg=btn_hlb_bg, width=17, command=update_record).place(x=348, y=30)
Button(RT_frame, text='Change Book Availability', font=btn_font, bg=btn_hlb_bg, width=19, command=change_availability).place(x=518, y=30)

# Right Bottom Frame
Label(RB_frame, text='BOOK INVENTORY', bg=rbf_bg, font=("Noto Sans CJK TC", 15, 'bold')).pack(side=TOP, fill=X)

tree = ttk.Treeview(RB_frame, selectmode=BROWSE, columns=('Book Name', 'Book ID', 'Author', 'Status', 'Issuer Card ID'))
XScrollbar = Scrollbar(tree, orient=HORIZONTAL, command=tree.xview)
YScrollbar = Scrollbar(tree, orient=VERTICAL, command=tree.yview)
XScrollbar.pack(side=BOTTOM, fill=X)
YScrollbar.pack(side=RIGHT, fill=Y)
tree.config(xscrollcommand=XScrollbar.set, yscrollcommand=YScrollbar.set)

tree.heading('Book Name', text='Book Name', anchor=CENTER)
tree.heading('Book ID', text='Book ID', anchor=CENTER)
tree.heading('Author', text='Author', anchor=CENTER)
tree.heading('Status', text='Status of the Book', anchor=CENTER)
tree.heading('Issuer Card ID', text='Card ID of the Issuer', anchor=CENTER)

tree.column('#0', width=0, stretch=NO)
tree.column('#1', width=225, stretch=NO)
tree.column('#2', width=70, stretch=NO)
tree.column('#3', width=150, stretch=NO)
tree.column('#4', width=105, stretch=NO)
tree.column('#5', width=132, stretch=NO)

tree.place(y=30, x=0, relheight=0.9, relwidth=1)

clear_and_display()

# Finalizing
root.update()
root.mainloop()
