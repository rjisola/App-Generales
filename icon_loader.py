# -*- coding: utf-8 -*-
"""
Icon Loader Helper
Módulo para cargar iconos PNG en aplicaciones Tkinter
"""

from PIL import Image, ImageTk
import os

class IconLoader:
    """Clase para cargar y gestionar iconos de aplicaciones."""
    
    def __init__(self, icons_dir="app_icons"):
        """
        Inicializa el cargador de iconos.
        
        Args:
            icons_dir: Nombre del directorio de iconos (relativo al script)
        """
        # Obtener directorio base del script que llama
        import sys
        if hasattr(sys.modules['__main__'], '__file__'):
            base_dir = os.path.dirname(os.path.abspath(sys.modules['__main__'].__file__))
        else:
            base_dir = os.getcwd()
        
        self.icons_dir = os.path.join(base_dir, icons_dir)
        self.icons = {}
        self._loaded = False
    
    def load_icon(self, icon_name, size=(64, 64)):
        """
        Carga un icono específico.
        
        Args:
            icon_name: Nombre del icono (sin extensión .png)
            size: Tupla (ancho, alto) para redimensionar
            
        Returns:
            PhotoImage del icono o None si no se pudo cargar
        """
        # Crear clave única para este icono y tamaño
        cache_key = f"{icon_name}_{size[0]}x{size[1]}"
        
        # Verificar si ya está en caché
        if cache_key in self.icons:
            return self.icons[cache_key]
        
        # Construir ruta del archivo
        icon_file = f"{icon_name}.png"
        icon_path = os.path.join(self.icons_dir, icon_file)
        
        if not os.path.exists(icon_path):
            print(f"⚠ Icono no encontrado: {icon_path}")
            return None
        
        try:
            # Cargar y redimensionar imagen
            img = Image.open(icon_path)
            img = img.resize(size, Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            
            # Guardar en caché
            self.icons[cache_key] = photo
            return photo
            
        except Exception as e:
            print(f"⚠ Error al cargar icono {icon_file}: {e}")
            return None
    
    def load_multiple(self, icon_specs):
        """
        Carga múltiples iconos de una vez.
        
        Args:
            icon_specs: Lista de tuplas (nombre, tamaño) o diccionarios
                       Ejemplos:
                       [('calculator', (64, 64)), ('folder', (32, 32))]
                       [{'name': 'calculator', 'size': (64, 64)}]
        
        Returns:
            Diccionario {nombre: PhotoImage}
        """
        loaded = {}
        
        for spec in icon_specs:
            if isinstance(spec, tuple):
                name, size = spec
            elif isinstance(spec, dict):
                name = spec['name']
                size = spec.get('size', (64, 64))
            else:
                continue
            
            icon = self.load_icon(name, size)
            if icon:
                loaded[name] = icon
        
        return loaded
    
    def set_window_icon(self, window, icon_name):
        """
        Establece el icono de una ventana Tkinter.
        
        Args:
            window: Ventana Tk o Toplevel
            icon_name: Nombre del icono a usar
        """
        icon_path = os.path.join(self.icons_dir, f"{icon_name}.png")
        
        if os.path.exists(icon_path):
            try:
                # Cargar icono para la ventana (tamaño original)
                img = Image.open(icon_path)
                photo = ImageTk.PhotoImage(img)
                
                # Establecer icono de ventana
                window.iconphoto(True, photo)
                
                # Guardar referencia para evitar garbage collection
                if not hasattr(window, '_icon_ref'):
                    window._icon_ref = []
                window._icon_ref.append(photo)
                
            except Exception as e:
                print(f"⚠ Error al establecer icono de ventana: {e}")

# Instancia global para uso conveniente
_global_loader = None

def get_icon_loader(icons_dir="app_icons"):
    """
    Obtiene la instancia global del cargador de iconos.
    
    Args:
        icons_dir: Directorio de iconos
        
    Returns:
        IconLoader instance
    """
    global _global_loader
    if _global_loader is None:
        _global_loader = IconLoader(icons_dir)
    return _global_loader

def load_icon(icon_name, size=(64, 64)):
    """
    Función de conveniencia para cargar un icono.
    
    Args:
        icon_name: Nombre del icono
        size: Tamaño (ancho, alto)
        
    Returns:
        PhotoImage del icono
    """
    loader = get_icon_loader()
    return loader.load_icon(icon_name, size)

def set_window_icon(window, icon_name):
    """
    Función de conveniencia para establecer icono de ventana.
    
    Args:
        window: Ventana Tk/Toplevel
        icon_name: Nombre del icono
    """
    loader = get_icon_loader()
    loader.set_window_icon(window, icon_name)

# Ejemplo de uso:
"""
from icon_loader import load_icon, set_window_icon

# En tu aplicación:
class MiApp:
    def __init__(self, root):
        self.root = root
        
        # Establecer icono de ventana
        set_window_icon(self.root, 'calculator')
        
        # Cargar icono para un botón
        folder_icon = load_icon('folder', size=(32, 32))
        btn = tk.Button(root, image=folder_icon, ...)
        btn.image = folder_icon  # Mantener referencia
"""
