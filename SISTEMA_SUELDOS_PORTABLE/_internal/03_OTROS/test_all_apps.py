import os
import glob
import subprocess
import time

def test_apps():
    apps = [f for f in glob.glob("*.pyw") if not f.startswith("debug_") and not f.startswith("borrar\\")]
    results = {}

    print(f"Buscando {len(apps)} aplicaciones...")
    
    for app in apps:
        print(f"Probando {app}...")
        try:
            # Ejecutar con python normal para capturar salida (usando .venv)
            python_exe = os.path.join(".venv", "Scripts", "python.exe")
            if not os.path.exists(python_exe):
                python_exe = "python" # fallback
                
            p = subprocess.Popen([python_exe, app], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            
            # Esperar 3 segundos. Si no crashea, asumimos que la GUI abrio bien.
            time.sleep(3)
            
            # Chequear si el proceso terminó (crash)
            if p.poll() is not None:
                stdout, stderr = p.communicate()
                results[app] = {"status": "ERROR", "stderr": stderr}
                print(f"  -> ERROR: {stderr.strip()[:100]}...")
            else:
                p.terminate() # Cerrar app
                results[app] = {"status": "OK", "stderr": ""}
                print(f"  -> OK")
                
        except Exception as e:
            results[app] = {"status": "EXCEPTION", "stderr": str(e)}
            print(f"  -> EXCEPTION: {e}")

    # Guardar resultados
    with open("test_results_all.txt", "w", encoding="utf-8") as f:
        f.write("=== RESULTADOS DE PRUEBA DE ARRANQUE ===\n\n")
        f.write(f"Total apps probadas: {len(apps)}\n")
        
        exitos = [k for k, v in results.items() if v['status'] == 'OK']
        f.write(f"Exitosas ({len(exitos)}): {', '.join(exitos)}\n\n")
        
        errores = [k for k, v in results.items() if v['status'] != 'OK']
        if errores:
            f.write(f"Con Errores ({len(errores)}):\n")
            for e in errores:
                f.write(f"--- {e} ---\n")
                f.write(f"Status: {results[e]['status']}\n")
                f.write(f"Detalle: {results[e]['stderr']}\n\n")

if __name__ == "__main__":
    test_apps()
    print("Prueba finalizada. Resultados en test_results_all.txt")
