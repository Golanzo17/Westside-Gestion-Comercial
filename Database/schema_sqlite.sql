-- =======================================================================
-- BASE DE DATOS SQLITE: gestion_comercial.db
-- Sistema de Gestión Comercial para Local de Ropa / Indumentaria
-- =======================================================================

-- 1. TABLA: configuracion
CREATE TABLE IF NOT EXISTS configuracion (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre_comercio TEXT NOT NULL DEFAULT 'Mi Local de Ropa',
    cuit TEXT DEFAULT '',
    direccion TEXT DEFAULT '',
    telefono TEXT DEFAULT '',
    email TEXT DEFAULT '',
    condicion_iva TEXT DEFAULT 'Responsable Inscripto',
    mensaje_ticket TEXT DEFAULT '¡Gracias por su compra! Cambios dentro de los 30 días con ticket.',
    moneda_simbolo TEXT DEFAULT '$',
    updated_at TEXT DEFAULT (datetime('now', 'localtime'))
);

-- 2. TABLA: usuarios
CREATE TABLE IF NOT EXISTS usuarios (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL UNIQUE,
    password_hash TEXT NOT NULL,
    nombre_completo TEXT NOT NULL,
    rol TEXT NOT NULL DEFAULT 'Vendedor',
    activo INTEGER NOT NULL DEFAULT 1,
    ultimo_login TEXT NULL,
    created_at TEXT DEFAULT (datetime('now', 'localtime'))
);

-- 3. TABLA: categorias
CREATE TABLE IF NOT EXISTS categorias (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL UNIQUE,
    descripcion TEXT NULL,
    activo INTEGER NOT NULL DEFAULT 1,
    created_at TEXT DEFAULT (datetime('now', 'localtime'))
);

-- 4. TABLA: talles
CREATE TABLE IF NOT EXISTS talles (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL UNIQUE,
    orden INTEGER NOT NULL DEFAULT 0,
    created_at TEXT DEFAULT (datetime('now', 'localtime'))
);

-- 5. TABLA: productos
CREATE TABLE IF NOT EXISTS productos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    codigo_barra TEXT UNIQUE NOT NULL,
    nombre TEXT NOT NULL,
    descripcion TEXT NULL,
    categoria_id INTEGER NOT NULL,
    precio_costo REAL NOT NULL DEFAULT 0.0,
    precio_venta REAL NOT NULL DEFAULT 0.0,
    porcentaje_ganancia REAL DEFAULT 50.0,
    imagen_ruta TEXT NULL,
    activo INTEGER NOT NULL DEFAULT 1,
    created_at TEXT DEFAULT (datetime('now', 'localtime')),
    updated_at TEXT DEFAULT (datetime('now', 'localtime')),
    FOREIGN KEY (categoria_id) REFERENCES categorias(id) ON UPDATE CASCADE
);

-- 6. TABLA: producto_talles (Matriz de stock por Talle y Color)
CREATE TABLE IF NOT EXISTS producto_talles (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    producto_id INTEGER NOT NULL,
    talle_id INTEGER NOT NULL,
    color TEXT NOT NULL DEFAULT 'Único',
    stock_actual INTEGER NOT NULL DEFAULT 0,
    stock_minimo INTEGER NOT NULL DEFAULT 2,
    sku_especifico TEXT NULL,
    created_at TEXT DEFAULT (datetime('now', 'localtime')),
    updated_at TEXT DEFAULT (datetime('now', 'localtime')),
    UNIQUE(producto_id, talle_id, color),
    FOREIGN KEY (producto_id) REFERENCES productos(id) ON DELETE CASCADE,
    FOREIGN KEY (talle_id) REFERENCES talles(id) ON UPDATE CASCADE
);

-- 7. TABLA: clientes
CREATE TABLE IF NOT EXISTS clientes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    dni_cuit TEXT NOT NULL UNIQUE,
    nombre TEXT NOT NULL,
    apellido TEXT NOT NULL,
    telefono TEXT NULL,
    email TEXT NULL,
    direccion TEXT NULL,
    ciudad TEXT NULL,
    notas TEXT NULL,
    activo INTEGER NOT NULL DEFAULT 1,
    created_at TEXT DEFAULT (datetime('now', 'localtime'))
);

-- 8. TABLA: cajas
CREATE TABLE IF NOT EXISTS cajas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    usuario_id INTEGER NOT NULL,
    fecha_apertura TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    fecha_cierre TEXT NULL,
    monto_inicial REAL NOT NULL DEFAULT 0.0,
    total_ventas_efectivo REAL NOT NULL DEFAULT 0.0,
    total_ventas_digital REAL NOT NULL DEFAULT 0.0,
    total_ingresos REAL NOT NULL DEFAULT 0.0,
    total_egresos REAL NOT NULL DEFAULT 0.0,
    monto_esperado REAL NOT NULL DEFAULT 0.0,
    monto_real REAL NULL,
    diferencia REAL NULL,
    estado TEXT NOT NULL DEFAULT 'Abierta',
    observaciones TEXT NULL,
    FOREIGN KEY (usuario_id) REFERENCES usuarios(id)
);

-- 9. TABLA: movimientos_caja
CREATE TABLE IF NOT EXISTS movimientos_caja (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    caja_id INTEGER NOT NULL,
    usuario_id INTEGER NOT NULL,
    fecha TEXT DEFAULT (datetime('now', 'localtime')),
    tipo TEXT NOT NULL,
    concepto TEXT NOT NULL,
    monto REAL NOT NULL,
    referencia TEXT NULL,
    FOREIGN KEY (caja_id) REFERENCES cajas(id) ON DELETE CASCADE,
    FOREIGN KEY (usuario_id) REFERENCES usuarios(id)
);

-- 10. TABLA: ventas
CREATE TABLE IF NOT EXISTS ventas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    numero_ticket TEXT NOT NULL UNIQUE,
    fecha TEXT DEFAULT (datetime('now', 'localtime')),
    usuario_id INTEGER NOT NULL,
    cliente_id INTEGER NOT NULL,
    caja_id INTEGER NOT NULL,
    tipo_comprobante TEXT NOT NULL DEFAULT 'Ticket',
    metodo_pago TEXT NOT NULL DEFAULT 'Efectivo',
    subtotal REAL NOT NULL DEFAULT 0.0,
    descuento_porcentaje REAL NOT NULL DEFAULT 0.0,
    descuento_monto REAL NOT NULL DEFAULT 0.0,
    recargo_monto REAL NOT NULL DEFAULT 0.0,
    total REAL NOT NULL DEFAULT 0.0,
    monto_abonado REAL NOT NULL DEFAULT 0.0,
    vuelto REAL NOT NULL DEFAULT 0.0,
    estado TEXT NOT NULL DEFAULT 'Completada',
    observaciones TEXT NULL,
    FOREIGN KEY (usuario_id) REFERENCES usuarios(id),
    FOREIGN KEY (cliente_id) REFERENCES clientes(id),
    FOREIGN KEY (caja_id) REFERENCES cajas(id)
);

-- 11. TABLA: detalle_ventas
CREATE TABLE IF NOT EXISTS detalle_ventas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    venta_id INTEGER NOT NULL,
    producto_id INTEGER NOT NULL,
    talle_id INTEGER NOT NULL,
    color TEXT NOT NULL DEFAULT 'Único',
    codigo_barra TEXT NOT NULL,
    descripcion_articulo TEXT NOT NULL,
    precio_unitario REAL NOT NULL,
    costo_unitario REAL NOT NULL DEFAULT 0.0,
    cantidad INTEGER NOT NULL,
    subtotal REAL NOT NULL,
    FOREIGN KEY (venta_id) REFERENCES ventas(id) ON DELETE CASCADE,
    FOREIGN KEY (producto_id) REFERENCES productos(id),
    FOREIGN KEY (talle_id) REFERENCES talles(id)
);

-- 12. TABLA: movimientos_stock
CREATE TABLE IF NOT EXISTS movimientos_stock (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    producto_id INTEGER NOT NULL,
    talle_id INTEGER NOT NULL,
    color TEXT NOT NULL DEFAULT 'Único',
    tipo_movimiento TEXT NOT NULL,
    cantidad INTEGER NOT NULL,
    stock_anterior INTEGER NOT NULL,
    stock_posterior INTEGER NOT NULL,
    motivo TEXT NULL,
    usuario_id INTEGER NOT NULL,
    fecha TEXT DEFAULT (datetime('now', 'localtime')),
    FOREIGN KEY (producto_id) REFERENCES productos(id),
    FOREIGN KEY (talle_id) REFERENCES talles(id),
    FOREIGN KEY (usuario_id) REFERENCES usuarios(id)
);

-- Índices
CREATE INDEX IF NOT EXISTS idx_prod_codigo ON productos (codigo_barra);
CREATE INDEX IF NOT EXISTS idx_prod_nombre ON productos (nombre);
CREATE INDEX IF NOT EXISTS idx_ventas_fecha ON ventas (fecha);
CREATE INDEX IF NOT EXISTS idx_ventas_ticket ON ventas (numero_ticket);
