import pandas as pd

URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSBK_5xmZ9uRO7p7AVWCRuis41Q0kvlZ7uFnmni4WC5jgBeGw2AZXVXU8jV5GYgqjnqEeCFoF-unTxu/pub?gid=1680576094&single=true&output=csv"

df = pd.read_csv(URL)
print("FILAS:", len(df))
print("COLUMNAS:", list(df.columns))
print(df.head(3))
