import tkinter as tk
from tkinter import ttk, Label, Entry, Button, messagebox
import xlsxwriter

def save_to_excel():
    
    name = entry_1.get()
    email = entry_3.get()
    contact = entry_5.get()
    address = entry_6.get()

    
    try:
        wb = xlsxwriter.Workbook('data.xlsx')
        sheet = wb.add_worksheet()
    except Exception as e:
        messagebox.showerror("Error", f"Failed to create Excel workbook: {e}")
        return

    headers = ["Name", "Email", "Contact No", "Address"]
    for col, header in enumerate(headers):
        sheet.write(0, col, header)

    row = 1
    data = [name, email, contact, address]
    for col, value in enumerate(data):
        sheet.write(row, col, value)

    try:
        wb.close()
        messagebox.showinfo("Success", "Data saved to Excel!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save Excel workbook: {e}")

root = tk.Tk()
root.geometry("500x500")
root.title('Registration form')

label_0 = Label(root, text="Registration form", width=20, font=("bold", 20))
label_0.place(x=90, y=60)

label_1 = Label(root, text="Name", width=20, font=("bold", 10))
label_1.place(x=80, y=130)
entry_1 = Entry(root)
entry_1.place(x=240, y=130)

label_3 = Label(root, text="Email", width=20, font=("bold", 10))
label_3.place(x=68, y=180)
entry_3 = Entry(root)
entry_3.place(x=240, y=180)

label_5 = Label(root, text="Contact No", width=20, font=("bold", 10))
label_5.place(x=67, y=240)
entry_5 = Entry(root)
entry_5.place(x=245, y=240)

label_6 = Label(root, text="Address", width=20, font=("bold", 10))
label_6.place(x=69, y=280)
entry_6 = Entry(root)
entry_6.place(x=245, y=280)

Button(root, text='Submit', width=20, bg="blue", fg='white', command=save_to_excel).place(x=180, y=380)

root.mainloop()