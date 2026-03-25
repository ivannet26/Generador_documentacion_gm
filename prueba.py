import pandas as pd
import sqlite3

URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSBK_5xmZ9uRO7p7AVWCRuis41Q0kvlZ7uFnmni4WC5jgBeGw2AZXVXU8jV5GYgqjnqEeCFoF-unTxu/pub?gid=1680576094&single=true&output=csv"

df = pd.read_csv(URL)

conn = sqlite3.connect("certificados.db")
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS solicitudes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    documento TEXT UNIQUE,
    nombres TEXT,
    apellidos TEXT,
    estado TEXT
)
""")

nuevos = 0

for _, row in df.iterrows():
    documento = str(row["N° DOCUMENTO"]).strip()
    nombres = row["NOMBRES"]
    apellidos = row["APELLIDOS"]
    estado = row["ESTADO"]

    try:
        cursor.execute("""
        INSERT INTO solicitudes (documento, nombres, apellidos, estado)
        VALUES (?, ?, ?, ?)
        """, (documento, nombres, apellidos, estado))
        nuevos += 1
    except sqlite3.IntegrityError:
        pass

conn.commit()
conn.close()

print("Nuevos registros insertados:", nuevos)



