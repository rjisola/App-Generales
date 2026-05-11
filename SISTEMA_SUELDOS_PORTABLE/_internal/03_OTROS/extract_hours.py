import pandas as pd
import os

# Data extracted from the image
data = {
    'Nombre': [
        'ERPEN', 
        'DOUGLAS B', 
        'CACERES', 
        'SILVERO', 
        'PALACIOS', 
        'TAPIA +15%', 
        'AYALA +15%'
    ],
    '18': [12, 12, 11.5, 10.5, 10.5, 11, 11],
    '19': [0, 0, 0, 0, 0, 0, 0], # Paro CGT
    '20': [11, 11, 11, 11, 11, 11, 11],
    '21': [7.5, 7.5, 0, 0, 0, 0, 0],
    '23': [11.5, 11.5, 11.5, 11.5, 11.5, 12, 12],
    '24': [12, 12, 10.5, 10.5, 10.5, 11, 11],
    '25': [11, 11, 10.5, 10.5, 10.5, 11.5, 11.5],
    '26': [10.5, 10.5, 10.5, 10.5, 10.5, 11, 11],
    '27': [11, 11, 10.5, 10.5, 10.5, 11, 11]
}

df = pd.DataFrame(data)

# Sumar por fila
numeric_cols = [col for col in df.columns if col != 'Nombre']
df['Total Horas'] = df[numeric_cols].sum(axis=1)

# Path to save the Excel file
output_path = r'C:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\horas_febrero.xlsx'

# Save to Excel
try:
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Horas Febrero')
    print(f"Archivo Excel generado exitosamente en: {output_path}")
    print("\nResumen de datos:")
    print(df[['Nombre', 'Total Horas']])
except Exception as e:
    print(f"Error al generar el Excel: {e}")
