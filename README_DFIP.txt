DFIP CASE SORTING SYSTEM - MICROSOFT ACCESS EDITION
==================================================

WHAT THIS VERSION DOES
----------------------
The programme reads DFIP case records from a Microsoft Access database file
(.accdb or .mdb). It expects a table with the following six fields:

    CaseID
    Case_Title
    Case_Type
    Date_Reported
    Severity
    Status

It displays these fields in a table and allows sorting by any of them.

IMPORTANT DATABASE NOTE
-----------------------
This version is designed for Microsoft Access databases (.accdb / .mdb).
It uses the Microsoft Access ODBC driver through the Python package pyodbc.
It does not currently connect directly to a Microsoft SQL Server database.

REQUIREMENTS
------------
1. Python 3.
2. The pyodbc package. Install it from a terminal with:

       python -m pip install pyodbc

3. The Microsoft Access Database Engine / Access ODBC driver installed on
   Windows. The driver architecture (32-bit or 64-bit) must match the Python
   installation. The driver is usually called:

       Microsoft Access Driver (*.mdb, *.accdb)

HOW TO RUN
----------
1. Keep dfip_case_sorting_system.py and this README in the same folder.
2. Open the folder in Visual Studio Code, PyCharm or IDLE.
3. Run dfip_case_sorting_system.py.
4. Click "Connect / Load Database".
5. Select your .accdb or .mdb file.
6. Select the table containing the DFIP case records if more than one table
   is available.
7. The programme loads and displays the six specified fields.

HOW TO USE THE SORTING FEATURES
-------------------------------
1. Choose a field from "Sort by".
2. Choose Ascending or Descending.
3. Click "Sort Cases".
4. Click "Display All" to restore the order returned by the database query.
5. Click "Refresh Data" to load the latest records from the selected table.
6. Click "Disconnect" to clear the current database and displayed records.

SEVERITY SORTING
----------------
Known severity values are ranked as follows:
Informational / Info, Low / Minor, Medium / Moderate, High / Major,
Critical / Severe.

Use Descending to show the most serious cases first. Unrecognised severity
labels are sorted alphabetically after the recognised values. Blank values
are kept at the bottom in either sort direction.

DATABASE SAFETY
---------------
The programme only runs a SELECT query. It does not add, edit or delete rows
in the Access database. "Remove Selected From View" only removes a record
from the current programme display; it does not delete it from the database.

The six field names must match exactly. The database table can have any name.
The Date_Reported field should use an Access Date/Time data type where possible.

HOW TO RUN THE TESTS
--------------------
Open a terminal in the folder containing both Python files and run:

    python -m unittest -v test_dfip_case_sorting_system.py

The tests check sorting by ID, title, status, severity and date, along with
empty lists, one-row lists, duplicate values and missing dates.
