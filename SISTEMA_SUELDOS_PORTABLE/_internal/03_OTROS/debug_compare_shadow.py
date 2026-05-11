import pandas as pd
import os
import shutil
import tempfile
import uuid

def safe_read(file_path, sheet_name):
    temp_dir = tempfile.gettempdir()
    temp_path = os.path.join(temp_dir, f"shadow_copy_{uuid.uuid4().hex}.xlsm")
    shutil.copy2(file_path, temp_path)
    try:
        return pd.read_excel(temp_path, sheet_name=sheet_name, header=None, engine='openpyxl')
    finally:
        try: os.remove(temp_path)
        except: pass

def compare_recuento():
    original_path = "03_OTROS/PROGRAMA DEPOSITO.xlsm"
    modified_path = "03_OTROS/PROGRAMA DEPOSITO_MODIFICADO.xlsm"
    sheet_name = "RECUENTO TOTAL"
    
    print(f"\n--- Comparando RECUENTO TOTAL (Columna K) ---")
    try:
        df_orig = safe_read(original_path, sheet_name)
        df_mod = safe_read(modified_path, sheet_name)
        
        diffs = []
        for r in range(len(df_orig)):
            nombre = df_orig.iloc[r, 3] # Col D
            if pd.notna(nombre) and str(nombre).strip() != "":
                val_orig = df_orig.iloc[r, 10] # Col K
                val_mod = df_mod.iloc[r, 10]
                
                try:
                    f_orig = float(val_orig) if pd.notna(val_orig) else 0.0
                    f_mod = float(val_mod) if pd.notna(val_mod) else 0.0
                    
                    if abs(f_orig - f_mod) > 1.0:
                        diffs.append({'Nombre': nombre, 'Orig': f_orig, 'Mod': f_mod, 'Diff': f_mod - f_orig})
                except: pass
        
        if not diffs:
            print("✓ No se detectaron diferencias significativas en RECUENTO TOTAL.")
        else:
            print(f"Se encontraron {len(diffs)} diferencias:")
            print(f"{'Empleado':<30} | {'Original':<12} | {'Modificado':<12} | {'Diferencia':<12}")
            print("-" * 75)
            # Ordenar por mayor diferencia
            diffs.sort(key=lambda x: abs(x['Diff']), reverse=True)
            for d in diffs[:20]: # Mostrar top 20
                print(f"{str(d['Nombre']):<30} | {d['Orig']:>12,.0f} | {d['Mod']:>12,.0f} | {d['Diff']:>12,.0f}")
                
    except Exception as e:
        print(f"Error: {e}")

compare_recuento()
