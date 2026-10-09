import sqlite3
from tabulate import tabulate

conn = sqlite3.connect("database.db")
cursor = conn.cursor()

cursor.execute("SELECT id, name, email FROM users")

rows = cursor.fetchall()

print(tabulate(rows, headers=["ID", "Name", "Email"], tablefmt="grid"))

conn.close()
