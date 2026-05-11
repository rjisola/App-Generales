import sqlite3
import os

class Database:
    def __init__(self, db_path="ordenes.db"):
        self.db_path = db_path
        self._init_db()

    def _get_connection(self):
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def _init_db(self):
        schema_path = os.path.join(os.path.dirname(__file__), "schema.sql")
        if os.path.exists(schema_path):
            with open(schema_path, "r") as f:
                schema = f.read()
            
            with self._get_connection() as conn:
                conn.executescript(schema)
                
                # Migración simple: Verificar si faltan columnas nuevas
                cursor = conn.cursor()
                cursor.execute("PRAGMA table_info(ordenes)")
                columns = [info['name'] for info in cursor.fetchall()]
                
                new_cols = {
                    "fecha_entrega": "TEXT",
                    "retira": "TEXT",
                    "destino": "TEXT"
                }
                
                for col, col_type in new_cols.items():
                    if col not in columns:
                        try:
                            conn.execute(f"ALTER TABLE ordenes ADD COLUMN {col} {col_type}")
                        except Exception as e:
                            print(f"Error agregando columna {col}: {e}")
                
                conn.commit()

    # --- PROVEEDORES ---
    def get_proveedores(self):
        with self._get_connection() as conn:
            return conn.execute("SELECT * FROM proveedores ORDER BY nombre ASC").fetchall()

    def get_proveedor_by_nombre(self, nombre):
        with self._get_connection() as conn:
            return conn.execute("SELECT * FROM proveedores WHERE nombre = ?", (nombre,)).fetchone()

    def delete_proveedor(self, nombre):
        with self._get_connection() as conn:
            conn.execute("DELETE FROM proveedores WHERE nombre = ?", (nombre.upper(),))
            conn.commit()

    def upsert_proveedor(self, nombre, domicilio="", domicilio2="", categoria_iva="", cuit=""):
        with self._get_connection() as conn:
            conn.execute("""
                INSERT INTO proveedores (nombre, domicilio, domicilio2, categoria_iva, cuit)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(nombre) DO UPDATE SET
                    domicilio = excluded.domicilio,
                    domicilio2 = excluded.domicilio2,
                    categoria_iva = excluded.categoria_iva,
                    cuit = excluded.cuit
            """, (nombre.upper(), domicilio.upper(), domicilio2.upper(), categoria_iva.upper(), cuit))
            conn.commit()

    def save_proveedor(self, p_info):
        """
        p_info: dict con nombre, domicilio, domicilio2, domicilio3, categoria_iva, cuit
        """
        # Unificamos domicilios para el esquema actual
        dom1 = p_info.get('domicilio', '')
        dom2 = (p_info.get('domicilio2', '') + " " + p_info.get('domicilio3', '')).strip()
        
        self.upsert_proveedor(
            p_info['nombre'], 
            dom1, 
            dom2, 
            p_info['categoria_iva'], 
            p_info['cuit']
        )

    # --- ORDENES ---
    def save_orden(self, orden_data, items):
        """
        orden_data: dict con numero_orden, fecha, proveedor_nombre, obra, autorizado, forma_pago, subtotal, iibb, ley23966, ley27430, iva, total
        items: list de dicts con descripcion, cantidad, precio_unitario, total_item
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()
            try:
                # Obtener la próxima orden esperada antes de insertar
                res = cursor.execute("SELECT valor FROM configuracion WHERE clave = 'proxima_orden'").fetchone()
                esperada = int(res['valor']) if res else 1

                cursor.execute("""
                    INSERT INTO ordenes (numero_orden, fecha, proveedor_nombre, obra, autorizado, forma_pago, fecha_entrega, retira, destino, subtotal, iibb, ley23966, ley27430, iva, total)
                    VALUES (:numero_orden, :fecha, :proveedor_nombre, :obra, :autorizado, :forma_pago, :fecha_entrega, :retira, :destino, :subtotal, :iibb, :ley23966, :ley27430, :iva, :total)
                """, orden_data)
                orden_id = cursor.lastrowid

                for item in items:
                    item['orden_id'] = orden_id
                    cursor.execute("""
                        INSERT INTO items_orden (orden_id, descripcion, cantidad, precio_unitario, total_item)
                        VALUES (:orden_id, :descripcion, :cantidad, :precio_unitario, :total_item)
                    """, item)
                
                # Solo incrementar la secuencia si el usuario usó el número que correspondía
                if int(orden_data['numero_orden']) == esperada:
                    proxima = esperada + 1
                    cursor.execute("UPDATE configuracion SET valor = ? WHERE clave = 'proxima_orden'", (str(proxima),))
                
                conn.commit()
                return orden_id
            except Exception as e:
                conn.rollback()
                raise e

    def get_ultima_orden_num(self):
        with self._get_connection() as conn:
            res = conn.execute("SELECT valor FROM configuracion WHERE clave = 'proxima_orden'").fetchone()
            return int(res['valor']) if res else 1

    def set_config(self, clave, valor):
        with self._get_connection() as conn:
            conn.execute("INSERT OR REPLACE INTO configuracion (clave, valor) VALUES (?, ?)", (clave, str(valor)))
            conn.commit()

    def get_config(self, clave, default=None):
        with self._get_connection() as conn:
            res = conn.execute("SELECT valor FROM configuracion WHERE clave = ?", (clave,)).fetchone()
            return res['valor'] if res else default
