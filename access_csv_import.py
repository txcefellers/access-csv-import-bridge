import argparse
import csv
from pathlib import Path
from typing import Iterable

import pyodbc


DEFAULT_DRIVER = "{Microsoft Access Driver (*.mdb, *.accdb)}"


def build_connection_string(db_path: Path, driver: str) -> str:
    return f"DRIVER={driver};DBQ={db_path};"


def get_insertable_columns(cursor: pyodbc.Cursor, table: str) -> list[str]:
    columns = []
    for col in cursor.columns(table=table):
        type_name = (col.type_name or "").upper()
        if "COUNTER" in type_name or "AUTOINCREMENT" in type_name:
            continue
        columns.append(col.column_name)
    if not columns:
        raise ValueError(f"No insertable columns found for table '{table}'.")
    return columns


def normalize_row(row: dict[str, str], columns: Iterable[str]) -> list[object]:
    values: list[object] = []
    for col in columns:
        raw = row.get(col)
        if raw is None:
            values.append(None)
            continue
        value = raw.strip()
        values.append(None if value == "" else value)
    return values


def import_csv_to_access(
    db_path: Path,
    table: str,
    csv_path: Path,
    delimiter: str,
    encoding: str,
    driver: str,
    truncate: bool,
) -> int:
    if not db_path.exists():
        raise FileNotFoundError(f"Access database not found: {db_path}")
    if not csv_path.exists():
        raise FileNotFoundError(f"CSV file not found: {csv_path}")

    connection_string = build_connection_string(db_path, driver)
    conn = pyodbc.connect(connection_string)
    try:
        cursor = conn.cursor()
        table_columns = get_insertable_columns(cursor, table)

        with csv_path.open("r", newline="", encoding=encoding) as csv_file:
            reader = csv.DictReader(csv_file, delimiter=delimiter)
            if reader.fieldnames is None:
                raise ValueError("CSV is empty or missing a header row.")

            csv_columns = [name.strip() for name in reader.fieldnames]
            common_columns = [c for c in table_columns if c in csv_columns]
            if not common_columns:
                raise ValueError(
                    f"No matching columns between CSV and table '{table}'. "
                    f"Table columns: {table_columns}"
                )

            if truncate:
                cursor.execute(f"DELETE FROM [{table}]")

            placeholders = ", ".join(["?"] * len(common_columns))
            col_sql = ", ".join(f"[{c}]" for c in common_columns)
            sql = f"INSERT INTO [{table}] ({col_sql}) VALUES ({placeholders})"

            inserted = 0
            for row in reader:
                values = normalize_row(row, common_columns)
                cursor.execute(sql, values)
                inserted += 1

            conn.commit()
            return inserted
    finally:
        conn.close()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Import rows from CSV into an existing Microsoft Access table."
    )
    parser.add_argument("--db-path", required=True, type=Path, help="Path to .mdb/.accdb file")
    parser.add_argument("--table", required=True, help="Target table name in Access")
    parser.add_argument("--csv-path", required=True, type=Path, help="Path to source CSV")
    parser.add_argument("--delimiter", default=",", help="CSV delimiter (default: ,)")
    parser.add_argument("--encoding", default="utf-8-sig", help="CSV encoding (default: utf-8-sig)")
    parser.add_argument(
        "--driver",
        default=DEFAULT_DRIVER,
        help=f"ODBC driver name (default: {DEFAULT_DRIVER})",
    )
    parser.add_argument(
        "--truncate",
        action="store_true",
        help="Delete existing rows from the table before import",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    inserted = import_csv_to_access(
        db_path=args.db_path,
        table=args.table,
        csv_path=args.csv_path,
        delimiter=args.delimiter,
        encoding=args.encoding,
        driver=args.driver,
        truncate=args.truncate,
    )
    print(f"Imported {inserted} row(s) into '{args.table}'.")


if __name__ == "__main__":
    main()
