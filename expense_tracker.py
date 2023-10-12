import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook

expenses_list = []

def add_expenses():
    name = name_entry.get()
    date = date_entry.get()
    amount = amount_entry.get()

    if name and date and amount:
        expenses_list.append((name, date, amount))
        update_expenses()
        clear_entries()
    else:
        messagebox.showerror("Error", "Please fill in all fields.")

def update_expenses():
    expenses_box.delete(0, tk.END)
    for i, expense in enumerate(expenses_list, start=1):
        expenses_box.insert(tk.END, f'{i}. Name: {expense[0]}, Date: {expense[1]}, Amount: {expense[2]}')

def clear_entries():
    name_entry.delete(0, tk.END)
    date_entry.delete(0, tk.END)
    amount_entry.delete(0, tk.END)

def clear_all_entries():
    expenses_list.clear()
    update_expenses()
    clear_entries()

def search_expenses():
    search_text = search_entry.get().lower()
    results = []

    for expense in expenses_list:
        if search_text in expense[0].lower() or search_text in expense[1].lower() or search_text in expense[2].lower():
            results.append(expense)

    expenses_box.delete(0, tk.END)
    for i, expense in enumerate(results, start=1):
        expenses_box.insert(tk.END, f'{i}. Name: {expense[0]}, Date: {expense[1]}, Amount: {expense[2]}')

def export_expenses_excel():
    wb = Workbook()
    ws = wb.active
    ws.append(["Name", "Date", "Amount"])

    for expense in expenses_list:
        ws.append(expense)

    wb.save('Expenses.xlsx')
    messagebox.showinfo('Export', "Expenses Exported to expenses.xlsx")

root = tk.Tk()
root.title("Expenses Tracker")

name_label = tk.Label(root, text='Expenses Name:')
date_label = tk.Label(root, text='Date:')
amount_label = tk.Label(root, text='Amount:')

name_entry = tk.Entry(root)
date_entry = tk.Entry(root)
amount_entry = tk.Entry(root)

add_button = tk.Button(root, text='Add Expense', command=add_expenses)
search_button = tk.Button(root, text='Search', command=search_expenses)
export_button = tk.Button(root, text='Export to Excel', command=export_expenses_excel)
clear_button = tk.Button(root, text='Clear All', command=clear_all_entries)

search_label = tk.Label(root, text="Search:")
search_entry = tk.Entry(root)

expenses_box = tk.Listbox(root, height=10, width=60)

name_label.grid(row=0, column=0)
date_label.grid(row=0, column=1)
amount_label.grid(row=0, column=2)

name_entry.grid(row=1, column=0)
date_entry.grid(row=1, column=1)
amount_entry.grid(row=1, column=2)

add_button.grid(row=2, column=0, columnspan=3)
search_label.grid(row=3, column=0)
search_entry.grid(row=3, column=1)
search_button.grid(row=3, column=2)

expenses_box.grid(row=4, column=0, columnspan=3)

export_button.grid(row=5, column=0, padx=10, pady=10)
clear_button.grid(row=5, column=1, padx=10, pady=10)

root.mainloop()
