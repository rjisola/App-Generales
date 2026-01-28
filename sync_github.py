import os
import subprocess
import datetime
import sys
import webbrowser

# Asegurar que el directorio del script esté en sys.path
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)

def run_command(command):
    try:
        # shell=True para Windows
        result = subprocess.run(command, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, shell=True)
        return True, result.stdout
    except subprocess.CalledProcessError as e:
        return False, e.stderr

def main():
    print("="*60)
    print("   SINCRONIZACIÓN AUTOMÁTICA CON GITHUB")
    print("="*60)
    
    # 1. Verificar instalación de Git
    ok, version = run_command("git --version")
    if not ok:
        print("❌ ERROR: Git no está instalado o no se encuentra en el PATH.")
        print("   Por favor descarga e instala Git desde: https://git-scm.com/download/win")
        input("\nPresiona Enter para salir...")
        return
    print(f"✓ Git detectado: {version.strip()}")

    # 2. Configuración de Usuario (si falta)
    ok, user_name = run_command("git config user.name")
    if not ok or not user_name.strip():
        print("\n⚠️ Configuración de usuario Git no detectada.")
        name = input("   Ingresa tu nombre para Git: ").strip()
        email = input("   Ingresa tu email para Git: ").strip()
        if name and email:
            run_command(f'git config --global user.name "{name}"')
            run_command(f'git config --global user.email "{email}"')
            print("✓ Usuario configurado.")

    # 3. Inicializar Repositorio
    if not os.path.exists(".git"):
        print("\nInitializing Git repository...")
        run_command("git init")
        run_command("git branch -M main")
        print("✓ Repositorio inicializado.")
    
    # 4. Configurar Remoto
    ok, remotes = run_command("git remote -v")
    if not ok or "origin" not in remotes:
        print("\n⚠️ No hay repositorio remoto configurado.")
        print("   --- CONFIGURACIÓN AUTOMÁTICA PARA 'App-Generales' ---")
        
        github_user = input("   Ingresa tu nombre de usuario de GitHub: ").strip()
        
        if github_user:
            url = f"https://github.com/{github_user}/App-Generales.git"
            print(f"   URL Destino: {url}")
            
            print("   ¿Ya existe el repositorio 'App-Generales' en tu GitHub?")
            resp = input("   (s = sí / n = no, crearlo ahora): ").lower().strip()
            
            if resp == 'n':
                print("   Abriendo navegador para crear el repositorio...")
                try:
                    webbrowser.open("https://github.com/new?name=App-Generales")
                except:
                    pass
                input("   >>> Presiona Enter una vez hayas creado el repositorio en el navegador...")
            
            run_command(f"git remote add origin {url}")
            print("✓ Remoto 'origin' configurado.")
        else:
            print("⚠️ No se ingresó usuario. Modo manual.")
            url = input("   URL del repositorio (ej: https://github.com/usuario/repo.git): ").strip()
            if url:
                run_command(f"git remote add origin {url}")
                print("✓ Remoto 'origin' configurado.")
            else:
                print("⚠️ Sin URL remota. Solo se guardará localmente.")

    # 5. Agregar y Commitear
    print("\n📦 Preparando archivos...")
    run_command("git add .")
    
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    commit_msg = f"Backup {timestamp}"
    
    # Verificar estado
    ok, status = run_command("git status --porcelain")
    if not status.strip():
        print("✓ No hay cambios nuevos para guardar.")
    else:
        print(f"💾 Creando commit: '{commit_msg}'")
        ok, out = run_command(f'git commit -m "{commit_msg}"')
        if ok:
            print("✓ Cambios guardados localmente.")

    # 6. Subir a GitHub
    ok, remotes = run_command("git remote -v")
    if ok and "origin" in remotes:
        print("\n☁️  Subiendo a GitHub...")
        ok, out = run_command("git push -u origin main")
        if ok:
            print("✅ ¡ÉXITO! Sincronización completada.")
        else:
            print("❌ Error al subir a GitHub:")
            print(out)
            print("\nConsejo: Si es la primera vez, asegúrate de tener permisos o haber iniciado sesión.")

    input("\nPresiona Enter para cerrar...")

if __name__ == "__main__":
    main()