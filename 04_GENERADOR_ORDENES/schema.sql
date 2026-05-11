CREATE TABLE IF NOT EXISTS proveedores (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT UNIQUE NOT NULL,
    domicilio TEXT,
    domicilio2 TEXT,
    categoria_iva TEXT,
    cuit TEXT
);

CREATE TABLE IF NOT EXISTS ordenes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    numero_orden INTEGER UNIQUE NOT NULL,
    fecha DATE NOT NULL,
    proveedor_nombre TEXT, -- Guardamos el nombre por si el proveedor cambia en el futuro
    obra TEXT,
    autorizado TEXT,
    forma_pago TEXT,
    fecha_entrega TEXT,
    retira TEXT,
    destino TEXT,
    subtotal REAL,
    iibb REAL,
    ley23966 REAL,
    ley27430 REAL,
    iva REAL,
    total REAL
);

CREATE TABLE IF NOT EXISTS items_orden (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    orden_id INTEGER,
    descripcion TEXT,
    cantidad REAL,
    precio_unitario REAL,
    total_item REAL,
    FOREIGN KEY (orden_id) REFERENCES ordenes(id)
);

CREATE TABLE IF NOT EXISTS configuracion (
    clave TEXT PRIMARY KEY,
    valor TEXT
);

INSERT OR IGNORE INTO configuracion (clave, valor) VALUES ('proxima_orden', '1');
INSERT OR IGNORE INTO configuracion (clave, valor) VALUES ('porcentaje_iva', '21');
INSERT OR IGNORE INTO configuracion (clave, valor) VALUES ('porcentaje_iibb', '0');
INSERT OR IGNORE INTO configuracion (clave, valor) VALUES ('porcentaje_ley23966', '0');
INSERT OR IGNORE INTO configuracion (clave, valor) VALUES ('porcentaje_ley27430', '0');
