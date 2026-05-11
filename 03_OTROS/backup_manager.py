"""
Módulo de backup automático para archivos Excel.
Crea copias de seguridad antes de procesar archivos.
"""
import shutil
from pathlib import Path
from datetime import datetime
from typing import Optional

class BackupManager:
    """
    Gestor de backups automáticos para archivos.
    """
    
    def __init__(self, backup_dir: str = "backups"):
        """
        Inicializa el gestor de backups en una ruta centralizada del proyecto.
        """
        # Calcular ruta centralizada relativa a este módulo (03_OTROS/)
        motor_dir = Path(__file__).parent.resolve()
        project_root = motor_dir.parent
        self.backup_dir = project_root / "02_CARPETAS" / backup_dir
        self.backup_dir.mkdir(parents=True, exist_ok=True)
    
    def create_backup(self, file_path: str, prefix: str = "backup") -> Optional[str]:
        """
        Crea un backup del archivo especificado.
        
        Args:
            file_path: Ruta al archivo a respaldar
            prefix: Prefijo para el nombre del backup
        
        Returns:
            str: Ruta al archivo de backup creado, o None si falla
        """
        try:
            source_path = Path(file_path)
            
            if not source_path.exists():
                print(f"Error: Archivo no encontrado: {file_path}")
                return None
            
            # Generar nombre de backup con timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_name = f"{prefix}_{source_path.stem}_{timestamp}{source_path.suffix}"
            backup_path = self.backup_dir / backup_name
            
            # Copiar archivo
            print(f"Creando backup: {backup_path}")
            shutil.copy2(source_path, backup_path)
            print(f"✓ Backup creado exitosamente")
            
            return str(backup_path)
            
        except Exception as e:
            print(f"Error al crear backup: {e}")
            return None
    
    def restore_backup(self, backup_path: str, destination_path: str) -> bool:
        """
        Restaura un archivo desde un backup.
        
        Args:
            backup_path: Ruta al archivo de backup
            destination_path: Ruta donde restaurar el archivo
        
        Returns:
            bool: True si la restauración fue exitosa
        """
        try:
            source = Path(backup_path)
            dest = Path(destination_path)
            
            if not source.exists():
                print(f"Error: Backup no encontrado: {backup_path}")
                return False
            
            # Confirmar con el usuario
            if dest.exists():
                response = input(f"¿Sobrescribir {destination_path}? (s/n): ")
                if response.lower() != 's':
                    print("Restauración cancelada")
                    return False
            
            print(f"Restaurando desde: {backup_path}")
            shutil.copy2(source, dest)
            print(f"✓ Archivo restaurado exitosamente a: {destination_path}")
            
            return True
            
        except Exception as e:
            print(f"Error al restaurar backup: {e}")
            return False
    
    def list_backups(self, pattern: str = "*") -> list:
        """
        Lista todos los backups disponibles.
        
        Args:
            pattern: Patrón de búsqueda (glob)
        
        Returns:
            list: Lista de rutas de backups
        """
        backups = sorted(self.backup_dir.glob(pattern), reverse=True)
        return [str(b) for b in backups]
    
    def clean_old_backups(self, keep_last: int = 5, pattern: str = "*"):
        """
        Elimina backups antiguos, manteniendo solo los más recientes.
        
        Args:
            keep_last: Número de backups a mantener
            pattern: Patrón de archivos a limpiar
        """
        backups = sorted(self.backup_dir.glob(pattern), reverse=True)
        
        if len(backups) <= keep_last:
            print(f"Solo hay {len(backups)} backups, no se eliminará ninguno")
            return
        
        to_delete = backups[keep_last:]
        print(f"Eliminando {len(to_delete)} backups antiguos...")
        
        for backup in to_delete:
            try:
                backup.unlink()
                print(f"  ✓ Eliminado: {backup.name}")
            except Exception as e:
                print(f"  ✗ Error eliminando {backup.name}: {e}")
        
        print(f"✓ Limpieza completada. Mantenidos {keep_last} backups más recientes")
    
    def get_backup_info(self, backup_path: str) -> dict:
        """
        Obtiene información sobre un backup.
        
        Args:
            backup_path: Ruta al backup
        
        Returns:
            dict: Información del backup
        """
        path = Path(backup_path)
        
        if not path.exists():
            return {"error": "Backup no encontrado"}
        
        stat = path.stat()
        
        return {
            "name": path.name,
            "size_mb": stat.st_size / (1024 * 1024),
            "created": datetime.fromtimestamp(stat.st_ctime).strftime("%Y-%m-%d %H:%M:%S"),
            "modified": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
        }

# Función de conveniencia
def create_auto_backup(file_path: str, backup_dir: str = "backups") -> Optional[str]:
    """
    Crea un backup automático de un archivo.
    
    Args:
        file_path: Ruta al archivo
        backup_dir: Directorio de backups
    
    Returns:
        str: Ruta al backup creado
    """
    manager = BackupManager(backup_dir)
    return manager.create_backup(file_path, prefix="auto")

# Ejemplo de uso
if __name__ == "__main__":
    # Crear gestor de backups
    manager = BackupManager("backups")
    
    # Crear backup
    file_to_backup = "PROGRAMA DEPOSITO 1ERA DICIEMBRE2025.xlsm"
    backup_path = manager.create_backup(file_to_backup)
    
    if backup_path:
        # Mostrar información del backup
        info = manager.get_backup_info(backup_path)
        print(f"\nInformación del backup:")
        print(f"  Nombre: {info['name']}")
        print(f"  Tamaño: {info['size_mb']:.2f} MB")
        print(f"  Creado: {info['created']}")
        
        # Listar todos los backups
        print(f"\nBackups disponibles:")
        for i, backup in enumerate(manager.list_backups(), 1):
            print(f"  {i}. {Path(backup).name}")
        
        # Limpiar backups antiguos (mantener últimos 5)
        manager.clean_old_backups(keep_last=5)
