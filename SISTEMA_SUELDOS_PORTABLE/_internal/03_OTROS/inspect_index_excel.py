import pandas as pd

excel_path = r"C:\Users\rjiso\OneDrive\Escritorio\carjorExcelJS\API\public\ACOMODAR_PDF\index.xlsx"
try:
    df = pd.read_excel(excel_path)
    print(df.head(10))
    print(df.columns)
except Exception as e:
    print(f"Error: {e}")
