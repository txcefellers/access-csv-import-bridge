# Access CSV Import Bridge

Imports rows from a CSV file into an existing table in a Microsoft Access database (`.mdb`/`.accdb`)
over ODBC, matching columns by name and skipping auto-increment columns.

## Requirements

- Windows with the "Microsoft Access Driver (*.mdb, *.accdb)" ODBC driver installed
  (comes with Microsoft Access, or install the free Microsoft Access Database Engine Redistributable).
- `pip install -r requirements.txt`

## Run

```bash
python access_csv_import.py --db-path C:\path\to\database.accdb --table Employees --csv-path data.csv
```

Optional flags:
- `--delimiter` — CSV delimiter (default: `,`)
- `--encoding` — CSV encoding (default: `utf-8-sig`)
- `--driver` — override the ODBC driver name
- `--truncate` — delete existing rows in the target table before importing

Only CSV columns that match existing table column names (case-sensitive) are imported; the target
table must already exist with a compatible schema.
