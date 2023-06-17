import tkinter as tk
import sqlite3
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
from openpyxl import Workbook
import tkcalendar as Cal
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog

MessageOn = False

def BaseCheck(): # Funcja sprawdzająca czy baza istnieje, jeśli nie tworzy podstawową wersje.
    isConnected = False
    conn = sqlite3.connect('Workers.db')
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='workers'")
    result = cursor.fetchone()
    if result is not None:
        print("Tabela istnieje.")
        isConnected = True
    else:
        messagebox.showwarning("Brak Bazy", "Nie wykryto bazy danych, utworzono nową")
        print("Tabela nie istnieje.")
        cursor.execute('''
    CREATE TABLE workers (
        ID INTEGER PRIMARY KEY AUTOINCREMENT,
        F_Name TEXT NOT NULL,
        L_Name TEXT NOT NULL,
        Position TEXT NOT NULL,
        Hired_Date TEXT NOT NULL,
        Salary TEXT NOT NULL
    )
    ''')
        conn.close()
    
    return isConnected

def ExcelExport():
    # Wybierz miejsce zapisu pliku
    filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    
    if filepath:
        # Tworzenie nowego arkusza
        workbook = Workbook()
        sheet = workbook.active
        
        # Pobranie danych z bazy danych SQLite
        conn = sqlite3.connect("Workers.db")
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM workers")
        results = cursor.fetchall()
        
        # Zapis danych do arkusza Excel
        sheet.append(["ID", "First Name", "Last Name", "Position", "Hired Date", "Salary"])
        for row in results:
            sheet.append(row)
        
        # Zapis pliku Excel
        workbook.save(filepath)
        
        # Zamykanie połączenia z bazą danych
        conn.close()

def BaseAdd(val, wind): #Funkcja dodająca informacje do bazy
    print(f'base Add: {val}')
    isConnected = BaseCheck()
    if isConnected:
        conn = sqlite3.connect('Workers.db')
        cursor = conn.cursor()
        sql = "INSERT INTO workers (F_Name, L_Name, Position, Hired_Date, Salary) VALUES (?, ?, ?, ?, ?)"
        cursor.execute(sql, val)
        conn.commit()
        conn.close()
    
def BaseUpdate(val, ID ,wind):
    isConnected = BaseCheck()
    print(val)
    sql = f"UPDATE workers SET F_Name = '{val[0]}', L_Name = '{val[1]}', Position = '{val[2]}', Hired_Date = '{val[3]}', Salary = '{val[4]}' WHERE ID = {ID};"
    print(sql)
    if isConnected:
        conn = sqlite3.connect('Workers.db')
        cursor = conn.cursor()
        cursor.execute(sql)
        conn.commit()
        conn.close()
    

def get_person(tree):
    selected_items = tree.selection()  # Pobranie zaznaczonych elementów
    if selected_items:
        for item in selected_items:
            values = tree.item(item)['values']  # Pobranie wartości zaznaczonego elementu
            person_id = values[0]  # Pobranie wartości ID (pierwsza kolumna)
            print("ID:", person_id)
            print(f"Drzewo{tree}")
            return person_id
    else:
        print("Nie zaznaczono żadnej osoby.")
        return None
    
# Pobiera dane z widgetów Entry i sprawdza je przed wysłaniem ich dalej do BaseAdd().
def GetData(*val, wind, choice = None, ID = None):
    isFilled = True
    global MessageOn
    print(f'get data: {val}')
    Val_list = []
    for values in val:
        if values.get() == "" or len(values.get()) == 0:
            isFilled = False
            break
        else:
            Val_list.append(values.get())
    if not isFilled:
        messagebox.showwarning("Ostrzeżenie", "Nie wszystkie pola zostały wypełnione")
        MessageOn = True
    elif choice == None:
        for values in val:
            values.delete(0,'end')    
        BaseAdd(Val_list, wind)
    elif choice == 1:
        for values in val:
            values.delete(0,'end')
            MessageOn = False
        BaseUpdate(Val_list, ID, wind)

def DrawEditing(window, ID):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    window_width = 200
    window_height = 300
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2

    window.geometry(f"{window_width}x{window_height}+{x}+{y}")

    isConnected = BaseCheck()
    if isConnected:
        conn = sqlite3.connect('Workers.db')
        cursor = conn.cursor()
        sql = f"SELECT * FROM workers WHERE ID={ID}"
        cursor.execute(sql)
        result = cursor.fetchone()
        print(f"Result: {result} || ID: {ID}")

    FirstNameLabel = ttk.Label(text=f"Imię [{result[0]}]")
    FirstNameLabel.pack(anchor="center")
    FirstName = ttk.Entry(window, width=23)
    FirstName.pack(anchor="center")
    FirstName.insert(tk.END, result[1])

    LastNameLabel = ttk.Label(text="Nazwisko")
    LastNameLabel.pack(anchor="center")
    LastName = ttk.Entry(window, width=23)
    LastName.pack(anchor="center")
    LastName.insert(tk.END, result[2])

    PositionEntryLabel = ttk.Label(text="Stanowisko")
    PositionEntryLabel.pack(anchor="center")
    ComboValues = ["Właściciel","Sprzatacz","Helpdesk"]
    PositionEntry = ttk.Combobox(window, values=ComboValues,width=20)
    PositionEntry.pack(anchor="center")
    PositionEntry.set( result[3])

    calendarLabel = ttk.Label(text="Data")
    calendarLabel.pack(anchor="center")
    calendar = Cal.DateEntry(window, width=20, locale="pl_PL", background="green", foreground="white")
    calendar.pack(anchor="center")
    calendar.set_date(result[4])

    SalaryLabel = ttk.Label(text="Wynagrodzenie")
    SalaryLabel.pack(anchor="center")
    Salary = ttk.Entry(window, width=23)
    Salary.pack(anchor="center")
    Salary.insert(tk.END, result[5])

    EditButton = ttk.Button(window, text="Zaktualizuj",width=19, command=lambda: [GetData(FirstName, LastName, PositionEntry, calendar, Salary,wind=window, choice=1, ID=ID), DeleteMenu(FirstNameLabel, FirstName,LastNameLabel,LastName,PositionEntryLabel,PositionEntry,calendarLabel, calendar,SalaryLabel,Salary,EditButton,ReturnButton, choice=4, window=window)])
    EditButton.pack(anchor="center", pady=10)

    ReturnButton = ttk.Button(window, text="Powrot",width=19, command=lambda: DeleteMenu(FirstNameLabel, FirstName,LastNameLabel,LastName,PositionEntryLabel,PositionEntry,calendarLabel, calendar,SalaryLabel,Salary,EditButton,ReturnButton, choice=2, window=window))
    ReturnButton.pack(anchor="center", pady=5)
    conn.close()
#Funkcja rysująca menu dodawania
def DrawAdding(window):
    
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    window_width = 200
    window_height = 300
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2

    window.geometry(f"{window_width}x{window_height}+{x}+{y}")
    FirstNameLabel = ttk.Label(text="Imię")
    FirstNameLabel.pack(anchor="center")
    FirstName = ttk.Entry(window, width=23)
    FirstName.pack(anchor="center")

    LastNameLabel = ttk.Label(text="Nazwisko")
    LastNameLabel.pack(anchor="center")
    LastName = ttk.Entry(window, width=23)
    LastName.pack(anchor="center")

    PositionEntryLabel = ttk.Label(text="Stanowisko")
    PositionEntryLabel.pack(anchor="center")
    ComboValues = ["Właściciel","Sprzatacz","Helpdesk"]
    PositionEntry = ttk.Combobox(window, values=ComboValues,width=20)
    PositionEntry.pack(anchor="center")

    calendarLabel = ttk.Label(text="Data")
    calendarLabel.pack(anchor="center")
    calendar = Cal.DateEntry(window, width=20, locale="pl_PL", background="green", foreground="white")
    calendar.pack(anchor="center")

    SalaryLabel = ttk.Label(text="Wynagrodzenie")
    SalaryLabel.pack(anchor="center")
    Salary = ttk.Entry(window, width=23)
    Salary.pack(anchor="center")

    AddButton = ttk.Button(window, text="Zapisz",width=19, command=lambda: GetData(FirstName, LastName, PositionEntry, calendar, Salary,wind=window))
    AddButton.pack(anchor="center", pady=10)

    ReturnButton = ttk.Button(window, text="Powrot",width=19, command=lambda: DeleteMenu(FirstNameLabel, FirstName,LastNameLabel,LastName,PositionEntryLabel,PositionEntry,calendarLabel, calendar,SalaryLabel,Salary,AddButton,ReturnButton, choice=2, window=window))
    ReturnButton.pack(anchor="center", pady=5)
#Funkcja usuwająca widgety i rysująca inne
def DeleteMenu(*widgets, choice, window, ID = None):
    global MessageOn
    if choice == 1:
        for widget in widgets:
            widget.destroy()
        DrawAdding(window) 
    if choice == 2:
        for widget in widgets:
            widget.destroy()
        DrawList(window, BaseCheck)   
    if choice == 3 and ID != None:
      for widget in widgets:
            widget.destroy()
      DrawEditing(window, ID)
    elif choice == 3 and ID == None:
      messagebox.showwarning("Ostrzeżenie", "Nie wybrano pracownika z listy")
    if choice == 4 and MessageOn == False:
        for widget in widgets:
            widget.destroy()
        DrawList(window, BaseCheck)

def DeleteWorker(ID):
    sql = f"DELETE FROM Workers WHERE ID={ID};"
    result = messagebox.askyesno("Informacja", "Czy na pewno chcesz usunąć pracownika? Wszystkie informacje o nim zostaną usunięte.")
    print(f"INFO RESULT:{result}")
    isConnected = BaseCheck()
    if isConnected and result == True:
        conn = sqlite3.connect('Workers.db')
        cursor = conn.cursor()
        cursor.execute(sql)
        conn.commit()
        conn.close()
    
#Funkcja rysująca pierwotne menu       
def DrawList(window, isConnected):

        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        window_width = 950
        window_height = 500
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2

        window.geometry(f"{window_width}x{window_height}+{x}+{y}")

        tree = ttk.Treeview(window, columns=('ID', 'First Name', 'Last Name', 'Position', 'Hired Date', 'Salary'), show='headings')
        tree.heading('ID', text='ID')
        tree.heading('First Name', text='Imię')
        tree.heading('Last Name', text='Nazwisko')
        tree.heading('Position', text='Stanowisko')
        tree.heading('Hired Date', text='Data Zatrudnienia')
        tree.heading('Salary', text='Wynagrodzenie')
        tree.grid(column=1,row=1,sticky="E")

        tree.column('ID', width=50)
        tree.column('First Name', width=150)
        tree.column('Last Name', width=150)
        tree.column('Position', width=150)
        tree.column('Hired Date', width=150)
        tree.column('Salary', width=150)

        scrollbar = ttk.Scrollbar(window, orient=tk.VERTICAL, command=tree.yview)
        scrollbar.grid(column=2, row=1, sticky="NS")
        tree.configure(yscrollcommand=scrollbar.set)

        AddButton = ttk.Button(window, text="Dodaj", width=10, command=lambda: DeleteMenu(tree,AddButton,EditButton,DeleteButton, choice=1, window=window))
        AddButton.grid(column=0,row=1,sticky="NW",pady=10)
        EditButton = ttk.Button(window, text="Edytuj",width=10, command=lambda: DeleteMenu(tree,AddButton,EditButton,DeleteButton, choice=3, window=window, ID=get_person(tree)))
        EditButton.grid(column=0,row=1,sticky="NW",pady = 40)

        ExportButton = ttk.Button(window, text="Eksportuj",width=10, command=lambda: ExcelExport())
        ExportButton.grid(column=0,row=1,sticky="NW",pady = 70)
        DeleteButton = ttk.Button(window, text="Usuń", width=10, command=lambda: [DeleteWorker(ID=get_person(tree)),DeleteMenu(tree,AddButton,EditButton,DeleteButton,ExportButton, choice=2,window=window)])
        DeleteButton.grid(column=0, row=1,sticky="NW",pady=100)
        if isConnected:
        # Połączenie z bazą danych SQLite
            conn = sqlite3.connect("Workers.db")
            cursor = conn.cursor()

        # Pobieranie danych z tabeli workers
            cursor.execute("SELECT * FROM workers")
            results = cursor.fetchall()

        # Dodawanie danych do Listbox
            for row in results:
                tree.insert('', 'end', values=row)

        # Zamykanie połączenia z bazą danych
            conn.close()
        else:
            print("Brak Połączenia z bazą")

