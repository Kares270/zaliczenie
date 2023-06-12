from tkinter import *
from tkinter import ttk
import tkinter as tk
import mysql.connector
import sqlite3
def CreateBase():
    conn = sqlite3.connect('Workers.db')
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='workers'")
    result = cursor.fetchone()

    if result is not None:
        print("Tabela istnieje.")
    else:
        print("Tabela nie istnieje.")
        cursor.execute('''
    
    CREATE TABLE workers (
        ID INTEGER PRIMARY KEY,
        F_Name TEXT NOT NULL,
        L_Name TEXT NOT NULL,
        Position TEXT NOT NULL,
        Hired_Date TEXT NOT NULL,
        Salary TEXT NOT NULL
    )
    ''')


mydb = mysql.connector.connect(
    host="192.168.0.101",
    user="PythonEntrance",
    password="Nzoz2003",
    database="stosowana"
)


root = Tk()
geometry_width = 700
geometry_height = 1000
geometry_size = f"{geometry_width}x{geometry_height}"
root.geometry(geometry_size)






FirstName = Entry(root,highlightcolor="red")
FirstName.grid(row=0,column=1)
Button_CreateBase = Button(root, text="Stwórz Baze", command=CreateBase)
Button_CreateBase.grid(row=1,column=1)

if mydb.is_connected():
    print("Połączenie z bazą danych zostało ustanowione.")

    mydb_cursor = mydb.cursor()
    mydb_cursor.execute("SELECT ID, IMIE, NAZWISKO FROM pacjent")
    rows = mydb_cursor.fetchall()
    x = 0
    scrollbar = ttk.Scrollbar(root)
    scrollbar.grid(row=2, column=2, sticky=tk.N+tk.S)
    tree = ttk.Treeview(root, yscrollcommand=scrollbar.set)
    tree["columns"] = ("column1", "column2", "column3")
    tree.heading("#0", text="ID")
    tree.heading("column1", text="Imie")
    tree.heading("column2", text="Nazwisko")
    tree.heading("column3", text="Pesel")
    tree.column("#0", width=50)
    tree.column("column1", width=100)
    tree.column("column2", width=100)
    tree.column("column3", width=100)
    scrollbar.config(command=tree.yview)
    for row in rows:
        x += 1
        tree.insert("", "end", text=f"{row[0]}", values=(f"{row[1]}",f"{row[2]}"))
        #text = f"{row[0]}--{row[1]}--{row[2]}"
        #Box.insert(x, text)
else:
    print("Połączenie z bazą danych nie zostało ustanowione.")
tree.grid(row=2,column=1)

root.mainloop()

