import pandas as pd
import os

original_path = "03_OTROS/PROGRAMA DEPOSITO.xlsm"
modified_path = "03_OTROS/PROGRAMA DEPOSITO_MODIFICADO.xlsm"
sheet_name = "RECUENTO TOTAL"

def compare_recuento():
    print(f"\n--- Comparando RECUENTO TOTAL (Columna K) ---")
    try:
        df_orig = pd.read_excel(original_path, sheet_name=sheet_name, header=None, engine='openpyxl')
        df_mod = pd.read_excel(modified_path, sheet_name=sheet_name, header=None, engine='openpyxl')
        
        diffs = []
        # Buscamos nombres en la Columna D (índice 3)
        for r in range(len(df_orig)):
            nombre = df_orig.iloc[r, 3]
            if pd.notna(nombre) and str(nombre).strip() != "":
                val_orig = df_orig.iloc[r, 10] # Col K (índice 10)
                val_mod = df_mod.iloc[r, 10]
                
                # Intentar convertir a float para comparar
                try:
                    f_orig = float(val_orig) if pd.notna(val_orig) else 0.0
                    f_mod = float(val_mod) if pd.notna(val_mod) else 0.0
                    
                    if abs(f_orig - f_mod) > 1.0: # Diferencia de más de 1 peso
                        diffs.append({
                            'Nombre': nombre,
                            'Original': f_orig,
                            'Modificado': f_mod,
                            'Diferencia': f_mod - f_orig
                        })
                except:
                    pass
        
        if not diffs:
            print("✓ No se detectaron diferencias significativas en RECUENTO TOTAL.")
        else:
            print(f"Se encontraron {len(diffs)} diferencias:")
            print(f"{'Empleado':<30} | {'Original':<12} | {'Modificado':<12} | {'Diferencia':<12}")
            print("-" * 75)
            for d in diffs:
                print(f"{str(d['Nombre']):<30} | {d['Original']:>12,.2f} | {d['Modificado']:>12,.2f} | {d['Diferencia']:>12,.2f}")
                
    except Exception as e:
        print(f"Error al comparar: {e}")

compare_recuento()
