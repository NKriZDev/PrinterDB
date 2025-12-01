**<h1>PrinterDB (PySide6)</h1>**

A small Windows desktop app for tracking printer inventory, repairs, and user assignments using a shared Excel workbook as the database.

**Features**

- Uses one .xlsx file with 3 sheets:

-> Data: inventory counts per model

-> Logs: issue/repair/return history with Jalali dates

-> Lists: printer names, users, and user→printer mappings


- Three GUI tabs:

-> Data (inventory + main actions)

-> Logs

-> Manage (users & printer dropdowns)


- Primary workflow: Replace & Send to Repair

-> Select user → select a printer the user has → pick Jalali date

-> Decrements Storage, increments Repair, logs actions

- Repair return workflow: move items from Repair back to Storage

- Hover tooltips show users who own a model, quantities, and acquire dates

- Atomic Excel saving to prevent corruption (network-share safe)



**Requirements**

- Python 3.10+

- Packages:
```pip install PySide6 pandas openpyxl```

**Run**
```python PrinterDB.py```

If the workbook doesn’t exist, the app creates it on first run.


**Excel Structure**

Sheet 1: ``Data``
Columns (exact):
``Row, Model, New in storage, Storage, Repair, Total, User``

Sheet 2: ``Logs``
Columns:
``Date (Jalali), Event, User, Model, Quantity, Notes``

Sheet 3: ``Lists``
Columns:
``PrinterNames, Usernames, UserHasPrinters``

``UserHasPrinters`` rows are:
``username|model``
Example: ``ali|HP 401`` (one row per unit)


**Network Workbook Location**

Set near the top of the script:

``NETWORK_DIR = r"\\192.168.20.15\DataCenter Office\IT"``
``EXCEL_PATH = os.path.join(NETWORK_DIR, "printers.xlsx")``

**Build a Windows ``.exe``**
``pip install pyinstaller
pyinstaller --onefile --windowed PrinterDB.py``


Output:
``dist/PrinterDB.exe``

If imports are missed:

```pyinstaller --onefile --windowed --hidden-import=openpyxl --hidden-import=pandas PrinterDB.py```

**Troubleshooting**

- Excel says file is corrupt: close Excel, delete the workbook, rerun the app. Atomic saves prevent future corruption.

- PermissionError on save (network): someone has the file open in Excel; close it and retry.

- AttributeError / missing methods: ensure methods are inside the correct class and properly indented.

**License**

MIT
