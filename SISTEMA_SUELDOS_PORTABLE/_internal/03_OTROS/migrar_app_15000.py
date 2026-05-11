import os
import shutil
import codecs

# --- CONFIGURACIÓN DE RUTAS ---
# Ajustamos las rutas basándonos en tu solicitud
BASE_DIR = r"C:\Users\rjiso\OneDrive\Escritorio\asistente-sueldos-js"
SOURCE_DIR = os.path.join(BASE_DIR, "app_12000")
DEST_DIR = os.path.join(BASE_DIR, "app_15000")

OLD_PORT = "12000"
NEW_PORT = "15000"

# --- CONTENIDO DE LA NUEVA GUI (Caratula Principal.html) ---
NEW_HTML_CONTENT = r"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Suite Workspace - Glassmorphism</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600&display=swap" rel="stylesheet">
    <script src="https://unpkg.com/@phosphor-icons/web"></script>

    <style>
        /* --- 1. CONFIGURACIÓN BASE --- */
        :root {
            --glass-bg: rgba(255, 255, 255, 0.05);
            --glass-border: rgba(255, 255, 255, 0.1);
            --glass-highlight: rgba(255, 255, 255, 0.2);
            --text-main: #ffffff;
            --text-muted: #a1a1aa;
            
            /* Los 4 Colores Principales (Neón suave) */
            --color-1: #6366f1; /* Indigo */
            --color-2: #ec4899; /* Pink */
            --color-3: #06b6d4; /* Cyan */
            --color-4: #8b5cf6; /* Violet */
        }

        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Inter', sans-serif;
            background-color: #0f172a; /* Fondo oscuro base */
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
            overflow-x: hidden;
            position: relative;
        }

        /* --- 2. FONDO ANIMADO (Para resaltar el efecto vidrio) --- */
        .background-blobs {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            z-index: -1;
            overflow: hidden;
        }

        .blob {
            position: absolute;
            border-radius: 50%;
            filter: blur(80px);
            opacity: 0.6;
            animation: float 10s infinite ease-in-out;
        }

        .blob-1 { width: 400px; height: 400px; background: var(--color-1); top: -10%; left: -10%; }
        .blob-2 { width: 300px; height: 300px; background: var(--color-2); bottom: 10%; right: -5%; animation-delay: 2s; }
        .blob-3 { width: 350px; height: 350px; background: var(--color-3); bottom: -10%; left: 20%; animation-delay: 4s; }
        .blob-4 { width: 250px; height: 250px; background: var(--color-4); top: 20%; right: 30%; animation-delay: 1s; }

        @keyframes float {
            0% { transform: translate(0, 0); }
            50% { transform: translate(20px, 40px); }
            100% { transform: translate(0, 0); }
        }

        /* --- 3. CONTENEDOR PRINCIPAL --- */
        .main-container {
            width: 90%;
            max-width: 1200px;
            padding: 40px;
            z-index: 10;
        }

        .header-title {
            color: var(--text-main);
            font-size: 2.5rem;
            margin-bottom: 10px;
            font-weight: 600;
            text-shadow: 0 4px 10px rgba(0,0,0,0.3);
        }

        .header-subtitle {
            color: var(--text-muted);
            margin-bottom: 40px;
            font-size: 1.1rem;
        }

        /* --- 4. GRID DE TARJETAS --- */
        .cards-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(320px, 1fr));
            gap: 24px;
        }

        /* --- 5. ESTILOS GLASSMORPHISM CARD --- */
        .glass-card {
            /* La magia del vidrio */
            background: var(--glass-bg);
            backdrop-filter: blur(16px);
            -webkit-backdrop-filter: blur(16px); /* Soporte Safari */
            border: 1px solid var(--glass-border);
            border-top: 1px solid var(--glass-highlight); /* Borde superior más brillante para efecto luz */
            border-left: 1px solid var(--glass-highlight);
            
            border-radius: 20px;
            padding: 30px;
            transition: all 0.3s ease;
            box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.3);
            
            display: flex;
            flex-direction: column;
            align-items: flex-start;
            position: relative;
            overflow: hidden;
            cursor: pointer;
        }

        .glass-card:hover {
            transform: translateY(-5px);
            background: rgba(255, 255, 255, 0.1);
            border-color: rgba(255, 255, 255, 0.3);
            box-shadow: 0 15px 40px 0 rgba(0, 0, 0, 0.5);
        }

        /* Icono */
        .card-icon {
            width: 50px;
            height: 50px;
            border-radius: 12px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 28px;
            margin-bottom: 20px;
            background: rgba(255,255,255,0.05);
            border: 1px solid rgba(255,255,255,0.1);
        }

        /* Variaciones de color para iconos */
        .icon-blue { color: var(--color-3); box-shadow: 0 0 15px rgba(6, 182, 212, 0.2); }
        .icon-purple { color: var(--color-4); box-shadow: 0 0 15px rgba(139, 92, 246, 0.2); }
        .icon-pink { color: var(--color-2); box-shadow: 0 0 15px rgba(236, 72, 153, 0.2); }
        .icon-indigo { color: var(--color-1); box-shadow: 0 0 15px rgba(99, 102, 241, 0.2); }

        /* Textos */
        .card-title {
            color: var(--text-main);
            font-size: 1.25rem;
            margin-bottom: 8px;
            font-weight: 600;
        }

        .card-desc {
            color: var(--text-muted);
            font-size: 0.9rem;
            line-height: 1.5;
            margin-bottom: 25px;
            flex-grow: 1;
        }

        /* Botón de acción */
        .launch-btn {
            padding: 10px 20px;
            border-radius: 30px;
            border: none;
            background: rgba(255, 255, 255, 0.1);
            color: white;
            font-weight: 500;
            font-size: 0.85rem;
            display: flex;
            align-items: center;
            gap: 8px;
            transition: background 0.2s;
        }

        .glass-card:hover .launch-btn {
            background: rgba(255, 255, 255, 0.25);
        }

    </style>
</head>
<body>

    <div class="background-blobs">
        <div class="blob blob-1"></div>
        <div class="blob blob-2"></div>
        <div class="blob blob-3"></div>
        <div class="blob blob-4"></div>
    </div>

    <main class="main-container">
        <h1 class="header-title">Workspace Suite</h1>
        <p class="header-subtitle">Selecciona una aplicación para comenzar a trabajar.</p>

        <div class="cards-grid">
            
            <div class="glass-card">
                <div class="card-icon icon-blue">
                    <i class="ph ph-chart-bar"></i>
                </div>
                <h3 class="card-title">Analytics Pro</h3>
                <p class="card-desc">Visualiza métricas en tiempo real, reportes de ventas y crecimiento con gráficos interactivos.</p>
                <button class="launch-btn">Abrir App <i class="ph ph-arrow-right"></i></button>
            </div>

            <div class="glass-card">
                <div class="card-icon icon-purple">
                    <i class="ph ph-users-three"></i>
                </div>
                <h3 class="card-title">Team Connect</h3>
                <p class="card-desc">Gestión de recursos humanos, chat interno y organización de equipos ágiles.</p>
                <button class="launch-btn">Abrir App <i class="ph ph-arrow-right"></i></button>
            </div>

            <div class="glass-card">
                <div class="card-icon icon-pink">
                    <i class="ph ph-kanban"></i>
                </div>
                <h3 class="card-title">Task Master</h3>
                <p class="card-desc">Tableros Kanban, seguimiento de proyectos y control de fechas límite.</p>
                <button class="launch-btn">Abrir App <i class="ph ph-arrow-right"></i></button>
            </div>

            <div class="glass-card">
                <div class="card-icon icon-indigo">
                    <i class="ph ph-cloud-arrow-up"></i>
                </div>
                <h3 class="card-title">Cloud Drive</h3>
                <p class="card-desc">Almacenamiento seguro, gestión de archivos y copias de seguridad automáticas.</p>
                <button class="launch-btn">Abrir App <i class="ph ph-arrow-right"></i></button>
            </div>

             <div class="glass-card">
                <div class="card-icon icon-blue">
                    <i class="ph ph-envelope-simple"></i>
                </div>
                <h3 class="card-title">Mail Hub</h3>
                <p class="card-desc">Cliente de correo centralizado con filtros inteligentes y respuestas automáticas.</p>
                <button class="launch-btn">Abrir App <i class="ph ph-arrow-right"></i></button>
            </div>

             <div class="glass-card">
                <div class="card-icon icon-purple">
                    <i class="ph ph-gear"></i>
                </div>
                <h3 class="card-title">Settings</h3>
                <p class="card-desc">Configuración global de la suite, permisos de usuario y personalización.</p>
                <button class="launch-btn">Abrir App <i class="ph ph-arrow-right"></i></button>
            </div>

        </div>
    </main>

</body>
</html>
"""

def main():
    print("=== Iniciando Migración de App 12000 a 15000 ===")
    
    # 1. Verificar origen
    if not os.path.exists(SOURCE_DIR):
        print(f"ERROR: No se encuentra el directorio origen: {SOURCE_DIR}")
        return

    # 2. Copiar carpeta (Limpiar destino si existe)
    if os.path.exists(DEST_DIR):
        print(f"El destino {DEST_DIR} ya existe. Eliminando versión anterior...")
        try:
            shutil.rmtree(DEST_DIR)
        except Exception as e:
            print(f"Error al eliminar destino: {e}")
            return
    
    print(f"Copiando archivos de {SOURCE_DIR} a {DEST_DIR}...")
    try:
        shutil.copytree(SOURCE_DIR, DEST_DIR)
    except Exception as e:
        print(f"Error al copiar archivos: {e}")
        return

    # 3. Procesar archivos (Reemplazo de puerto y GUI)
    print("Procesando archivos copiados...")
    
    files_modified = 0
    gui_updated = False
    
    # Archivos que probablemente sean la entrada principal
    possible_index_files = ['index.html', 'default.html', 'main.html', 'home.html', 'AbrirLauncher.html']

    for root, dirs, files in os.walk(DEST_DIR):
        for filename in files:
            file_path = os.path.join(root, filename)
            
            # A. Reemplazar puerto 12000 -> 15000 en archivos de texto
            try:
                # Intentar leer con utf-8
                with codecs.open(file_path, 'r', 'utf-8') as f:
                    content = f.read()
                
                if OLD_PORT in content:
                    new_content = content.replace(OLD_PORT, NEW_PORT)
                    with codecs.open(file_path, 'w', 'utf-8') as f:
                        f.write(new_content)
                    print(f"  [PORT] Actualizado en: {filename}")
                    files_modified += 1
            except UnicodeDecodeError:
                # Si falla utf-8, intentar latin-1 o saltar si es binario
                pass
            except Exception as e:
                print(f"  [WARN] No se pudo leer {filename}: {e}")

            # B. Detectar y reemplazar GUI principal
            # Si el archivo coincide con nombres comunes de index, lo reemplazamos
            if filename in possible_index_files and not gui_updated:
                print(f"  [GUI] Encontrado posible entry point: {filename}. Reemplazando con nueva carátula...")
                
                # Backup del original
                shutil.move(file_path, file_path + ".bak")
                
                # Escribir nueva GUI
                with codecs.open(file_path, 'w', 'utf-8') as f:
                    f.write(NEW_HTML_CONTENT)
                
                gui_updated = True

    # Si no se encontró ningún archivo index para reemplazar, crear uno nuevo en la raíz
    if not gui_updated:
        new_index_path = os.path.join(DEST_DIR, "index.html")
        print(f"  [GUI] No se encontró archivo index existente. Creando {new_index_path}...")
        with codecs.open(new_index_path, 'w', 'utf-8') as f:
            f.write(NEW_HTML_CONTENT)

    print("\n=== Migración Completada ===")
    print(f"Directorio generado: {DEST_DIR}")
    print(f"Archivos con puerto modificado: {files_modified}")
    print("La nueva interfaz gráfica ha sido aplicada.")

if __name__ == "__main__":
    main()
import os
import shutil
import codecs
import ctypes

# --- CONFIGURACIÓN DE RUTAS ---
# Ajustamos las rutas basándonos en tu solicitud
BASE_DIR = r"C:\Users\rjiso\OneDrive\Escritorio\asistente-sueldos-js"
SOURCE_DIR = os.path.join(BASE_DIR, "app_12000")
DEST_DIR = os.path.join(BASE_DIR, "app_15000")

OLD_PORT = "12000"
NEW_PORT = "15000"

# --- CONTENIDO DE LA NUEVA GUI (Caratula Principal.html) ---
NEW_HTML_CONTENT = r"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Suite Workspace - Glassmorphism</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600&display=swap" rel="stylesheet">
    <script src="https://unpkg.com/@phosphor-icons/web"></script>

    <style>
        /* --- 1. CONFIGURACIÓN BASE --- */
        :root {
            --glass-bg: rgba(255, 255, 255, 0.05);
            --glass-border: rgba(255, 255, 255, 0.1);
            --glass-highlight: rgba(255, 255, 255, 0.2);
            --text-main: #ffffff;
            --text-muted: #a1a1aa;
            
            /* Los 4 Colores Principales (Neón suave) */
            --color-1: #6366f1; /* Indigo */
            --color-2: #ec4899; /* Pink */
            --color-3: #06b6d4; /* Cyan */
            --color-4: #8b5cf6; /* Violet */
        }

        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Inter', sans-serif;
            background-color: #0f172a; /* Fondo oscuro base */
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
            overflow-x: hidden;
            position: relative;
        }

        /* --- 2. FONDO ANIMADO (Para resaltar el efecto vidrio) --- */
        .background-blobs {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            z-index: -1;
            overflow: hidden;
        }

        .blob {
            position: absolute;
            border-radius: 50%;
            filter: blur(80px);
            opacity: 0.6;
            animation: float 10s infinite ease-in-out;
        }

        .blob-1 { width: 400px; height: 400px; background: var(--color-1); top: -10%; left: -10%; }
        .blob-2 { width: 300px; height: 300px; background: var(--color-2); bottom: 10%; right: -5%; animation-delay: 2s; }
        .blob-3 { width: 350px; height: 350px; background: var(--color-3); bottom: -10%; left: 20%; animation-delay: 4s; }
        .blob-4 { width: 250px; height: 250px; background: var(--color-4); top: 20%; right: 30%; animation-delay: 1s; }

        @keyframes float {
            0% { transform: translate(0, 0); }
            50% { transform: translate(20px, 40px); }
            100% { transform: translate(0, 0); }
        }

        /* --- 3. CONTENEDOR PRINCIPAL --- */
        .main-container {
            width: 90%;
            max-width: 1200px;
            padding: 40px;
            z-index: 10;
        }

        .header-title {
            color: var(--text-main);
            font-size: 2.5rem;
            margin-bottom: 10px;
            font-weight: 600;
            text-shadow: 0 4px 10px rgba(0,0,0,0.3);
        }

        .header-subtitle {
            color: var(--text-muted);
            margin-bottom: 40px;
            font-size: 1.1rem;
        }

        /* --- 4. GRID DE TARJETAS --- */
        .cards-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(320px, 1fr));
            gap: 24px;
        }

        /* --- 5. ESTILOS GLASSMORPHISM CARD --- */
        .glass-card {
            /* La magia del vidrio */
            background: var(--glass-bg);
            backdrop-filter: blur(16px);
            -webkit-backdrop-filter: blur(16px); /* Soporte Safari */
            border: 1px solid var(--glass-border);
            border-top: 1px solid var(--glass-highlight); /* Borde superior más brillante para efecto luz */
            border-left: 1px solid var(--glass-highlight);
            
            border-radius: 20px;
            padding: 30px;
            transition: all 0.3s ease;
            box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.3);
            
            display: flex;
            flex-direction: column;
            align-items: flex-start;
            position: relative;
            overflow: hidden;
            cursor: pointer;
        }

        .glass-card:hover {
            transform: translateY(-5px);
            background: rgba(255, 255, 255, 0.1);
            border-color: rgba(255, 255, 255, 0.3);
            box-shadow: 0 15px 40px 0 rgba(0, 0, 0, 0.5);
        }

        /* Icono */
        .card-icon {
            width: 50px;
            height: 50px;
            border-radius: 12px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 28px;
            margin-bottom: 20px;
            background: rgba(255,255,255,0.05);
            border: 1px solid rgba(255,255,255,0.1);
        }

        /* Variaciones de color para iconos */
        .icon-blue { color: var(--color-3); box-shadow: 0 0 15px rgba(6, 182, 212, 0.2); }
        .icon-purple { color: var(--color-4); box-shadow: 0 0 15px rgba(139, 92, 246, 0.2); }
        .icon-pink { color: var(--color-2); box-shadow: 0 0 15px rgba(236, 72, 153, 0.2); }
        .icon-indigo { color: var(--color-1); box-shadow: 0 0 15px rgba(99, 102, 241, 0.2); }

        /* Textos */
        .card-title {
            color: var(--text-main);
            font-size: 1.25rem;
            margin-bottom: 8px;
            font-weight: 600;
        }

        .card-desc {
            color: var(--text-muted);
            font-size: 0.9rem;
            line-height: 1.5;
            margin-bottom: 25px;
            flex-grow: 1;
        }

        /* Botón de acción */
        .launch-btn {
            padding: 10px 20px;
            border-radius: 30px;
            border: none;
            background: rgba(255, 255, 255, 0.1);
            color: white;
            font-weight: 500;
            font-size: 0.85rem;
            display: flex;
            align-items: center;
            gap: 8px;
            transition: background 0.2s;
        }

        .glass-card:hover .launch-btn {
            background: rgba(255, 255, 255, 0.25);
        }

    </style>
</head>
<body>

    <div class="background-blobs">
        <div class="blob blob-1"></div>
        <div class="blob blob-2"></div>
        <div class="blob blob-3"></div>
        <div class="blob blob-4"></div>
    </div>

    <main class="main-container">
        <h1 class="header-title">Workspace Suite</h1>
        <p class="header-subtitle">Selecciona una aplicación para comenzar a trabajar.</p>

        <div class="cards-grid">
            
            <div class="glass-card">
                <div class="card-icon icon-blue">
                    <i class="ph ph-chart-bar"></i>
                </div>
                <h3 class="card-title">Analytics Pro</h3>
                <p class="card-desc">Visualiza métricas en tiempo real, reportes de ventas y crecimiento con gráficos interactivos.</p>
                <button class="launch-btn">Abrir App <i class="ph ph-arrow-right"></i></button>
            </div>

            <div class="glass-card">
                <div class="card-icon icon-purple">
                    <i class="ph ph-users-three"></i>
                </div>
                <h3 class="card-title">Team Connect</h3>
                <p class="card-desc">Gestión de recursos humanos, chat interno y organización de equipos ágiles.</p>
                <button class="launch-btn">Abrir App <i class="ph ph-arrow-right"></i></button>
            </div>

            <div class="glass-card">
                <div class="card-icon icon-pink">
                    <i class="ph ph-kanban"></i>
                </div>
                <h3 class="card-title">Task Master</h3>
                <p class="card-desc">Tableros Kanban, seguimiento de proyectos y control de fechas límite.</p>
                <button class="launch-btn">Abrir App <i class="ph ph-arrow-right"></i></button>
            </div>

            <div class="glass-card">
                <div class="card-icon icon-indigo">
                    <i class="ph ph-cloud-arrow-up"></i>
                </div>
                <h3 class="card-title">Cloud Drive</h3>
                <p class="card-desc">Almacenamiento seguro, gestión de archivos y copias de seguridad automáticas.</p>
                <button class="launch-btn">Abrir App <i class="ph ph-arrow-right"></i></button>
            </div>

             <div class="glass-card">
                <div class="card-icon icon-blue">
                    <i class="ph ph-envelope-simple"></i>
                </div>
                <h3 class="card-title">Mail Hub</h3>
                <p class="card-desc">Cliente de correo centralizado con filtros inteligentes y respuestas automáticas.</p>
                <button class="launch-btn">Abrir App <i class="ph ph-arrow-right"></i></button>
            </div>

             <div class="glass-card">
                <div class="card-icon icon-purple">
                    <i class="ph ph-gear"></i>
                </div>
                <h3 class="card-title">Settings</h3>
                <p class="card-desc">Configuración global de la suite, permisos de usuario y personalización.</p>
                <button class="launch-btn">Abrir App <i class="ph ph-arrow-right"></i></button>
            </div>

        </div>
    </main>

</body>
</html>
"""

def main():
    print("=== Iniciando Migración de App 12000 a 15000 ===")
    
    # 1. Verificar origen
    if not os.path.exists(SOURCE_DIR):
        print(f"ERROR: No se encuentra el directorio origen: {SOURCE_DIR}")
        return

    # 2. Copiar carpeta (Limpiar destino si existe)
    if os.path.exists(DEST_DIR):
        print(f"El destino {DEST_DIR} ya existe. Eliminando versión anterior...")
        try:
            shutil.rmtree(DEST_DIR)
        except Exception as e:
            print(f"Error al eliminar destino: {e}")
            return
    
    print(f"Copiando archivos de {SOURCE_DIR} a {DEST_DIR}...")
    try:
        shutil.copytree(SOURCE_DIR, DEST_DIR)
    except Exception as e:
        print(f"Error al copiar archivos: {e}")
        return

    # 3. Procesar archivos (Reemplazo de puerto y GUI)
    print("Procesando archivos copiados...")
    
    files_modified = 0
    gui_updated = False
    
    # Archivos que probablemente sean la entrada principal
    possible_index_files = ['index.html', 'default.html', 'main.html', 'home.html', 'AbrirLauncher.html']

    for root, dirs, files in os.walk(DEST_DIR):
        for filename in files:
            file_path = os.path.join(root, filename)
            
            # A. Reemplazar puerto 12000 -> 15000 en archivos de texto
            try:
                # Intentar leer con utf-8
                with codecs.open(file_path, 'r', 'utf-8') as f:
                    content = f.read()
                
                if OLD_PORT in content:
                    new_content = content.replace(OLD_PORT, NEW_PORT)
                    with codecs.open(file_path, 'w', 'utf-8') as f:
                        f.write(new_content)
                    print(f"  [PORT] Actualizado en: {filename}")
                    files_modified += 1
            except UnicodeDecodeError:
                # Si falla utf-8, intentar latin-1 o saltar si es binario
                pass
            except Exception as e:
                print(f"  [WARN] No se pudo leer {filename}: {e}")

            # B. Detectar y reemplazar GUI principal
            # Si el archivo coincide con nombres comunes de index, lo reemplazamos
            if filename in possible_index_files and not gui_updated:
                print(f"  [GUI] Encontrado posible entry point: {filename}. Reemplazando con nueva carátula...")
                
                # Backup del original
                shutil.move(file_path, file_path + ".bak")
                
                # Escribir nueva GUI
                with codecs.open(file_path, 'w', 'utf-8') as f:
                    f.write(NEW_HTML_CONTENT)
                
                gui_updated = True

    # Si no se encontró ningún archivo index para reemplazar, crear uno nuevo en la raíz
    if not gui_updated:
        new_index_path = os.path.join(DEST_DIR, "index.html")
        print(f"  [GUI] No se encontró archivo index existente. Creando {new_index_path}...")
        with codecs.open(new_index_path, 'w', 'utf-8') as f:
            f.write(NEW_HTML_CONTENT)

    print("\n=== Migración Completada ===")
    print(f"Directorio generado: {DEST_DIR}")
    print(f"Archivos con puerto modificado: {files_modified}")
    print("La nueva interfaz gráfica ha sido aplicada.")

    # Mostrar ventana emergente de aviso
    ctypes.windll.user32.MessageBoxW(0, f"La migración a app_15000 ha finalizado correctamente.\n\nArchivos modificados: {files_modified}", "Proceso Terminado", 0x40)

if __name__ == "__main__":
    main()