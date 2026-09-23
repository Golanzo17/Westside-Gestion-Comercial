-- =======================================================================
-- DATOS INICIALES PARA SQLITE: gestion_comercial.db
-- =======================================================================

-- 1. USUARIOS (admin / admin123, vendedor / 1234 y gerente / gerente123)
INSERT OR REPLACE INTO usuarios (id, username, dni, password_hash, nombre, apellido, nombre_completo, rol, activo)
VALUES 
(1, 'admin', '10000001', 'PBKDF2$SHA256$100000$KzlmMZm4PZ0Hh1YNHZqQDw==$LM2PQmr9ClthJwIKhhlPCN/XNyfw7+gXNM1O801EqiU=', 'Administrador', 'General', 'Administrador General', 'Administrador', 1),
(2, 'vendedor', '10000002', 'PBKDF2$SHA256$100000$5J0bYHMjG/MrsBvPlSRiZQ==$AximwF6BMDNCSTFo/Ji1a1Wc+15Nw5XgC+nmlZL192Y=', 'Vendedor', 'De Turno', 'Vendedor de Turno', 'Vendedor', 1),
(3, 'gerente', '10000003', 'PBKDF2$SHA256$100000$a606/LnGvERCjMeOWvEduQ==$KbsU0/h6RLoTTtqBps8fg16jvCN2SoxpP+OrX51GS6E=', 'Gerente', 'General', 'Gerente General', 'Gerente', 1);

-- 3. TALLES
INSERT OR REPLACE INTO talles (id, nombre, orden) VALUES
(1, 'XS', 1), (2, 'S', 2), (3, 'M', 3), (4, 'L', 4), (5, 'XL', 5), (6, 'XXL', 6),
(7, '36', 7), (8, '38', 8), (9, '40', 9), (10, '42', 10), (11, '44', 11), (12, 'Único', 12);

-- 4. CATEGORÍAS
INSERT OR REPLACE INTO categorias (id, nombre, descripcion, activo) VALUES
(1, 'Remeras y Tops', 'Remeras estampadas, lisas, musculosas y tops', 1),
(2, 'Pantalones y Jeans', 'Jeans mom, skinny, wide leg, joggers y pantalones de vestir', 1),
(3, 'Buzos y Camperas', 'Buzos hoodie, camperas de abrigo, bombers y cardigans', 1),
(4, 'Vestidos y Polleras', 'Vestidos de fiesta, casuales y polleras', 1),
(5, 'Accesorios y Calzado', 'Cinturones, gorras, medias y calzado urbano', 1);

-- 5. CLIENTE INICIAL (Consumidor Final)
INSERT OR REPLACE INTO clientes (id, dni_cuit, nombre, apellido, telefono, email, direccion, ciudad, notas, activo)
VALUES 
(1, '00000000', 'Consumidor', 'Final', '000-0000', 'consumidor@final.com', 'Mostrador', 'Local', 'Cliente genérico', 1),
(2, '35123456', 'Gonzalo', 'Fernández', '11-3456-7890', 'gonzalo@email.com', 'Calle Falsa 123', 'Buenos Aires', 'Cliente habitual', 1);

-- 6. PRENDAS DE DEMOSTRACIÓN
INSERT OR REPLACE INTO productos (id, codigo_barra, nombre, descripcion, categoria_id, precio_costo, precio_venta, porcentaje_ganancia, activo) VALUES
(1, '779001001', 'Remera Oversize Vintage Acid Wash', 'Remera de algodón peinado 100% corte oversize', 1, 9500.0, 18500.0, 94.74, 1),
(2, '779001002', 'Jean Mom Fit Celeste Nevado', 'Jean rígido tiro alto con calce relajado', 2, 18000.0, 36000.0, 100.0, 1),
(3, '779001003', 'Buzo Hoodie Canguro Frisa Pesada', 'Buzo con capucha y bolsillo delantero unisex', 3, 22000.0, 42500.0, 93.18, 1),
(4, '779001004', 'Campera Puffer Corta Negra', 'Campera térmica impermeable con cierre reforzado', 3, 35000.0, 68000.0, 94.29, 1),
(5, '779001005', 'Gorra Trucker Urbana Bordada', 'Gorra con visera curva y red respirable', 5, 4500.0, 9800.0, 117.78, 1);

-- 7. STOCK POR TALLE Y COLOR
INSERT OR REPLACE INTO producto_talles (producto_id, talle_id, color, stock_actual, stock_minimo, sku_especifico) VALUES
(1, 2, 'Negro', 10, 3, 'REM-OV-BLK-S'),
(1, 3, 'Negro', 15, 3, 'REM-OV-BLK-M'),
(1, 4, 'Negro', 8, 3, 'REM-OV-BLK-L'),
(1, 5, 'Negro', 5, 2, 'REM-OV-BLK-XL'),
(1, 2, 'Blanco', 8, 3, 'REM-OV-WHT-S'),
(1, 3, 'Blanco', 12, 3, 'REM-OV-WHT-M'),
(1, 4, 'Blanco', 6, 2, 'REM-OV-WHT-L'),
(2, 7, 'Celeste', 4, 2, 'JN-MOM-CEL-36'),
(2, 8, 'Celeste', 10, 3, 'JN-MOM-CEL-38'),
(2, 9, 'Celeste', 12, 3, 'JN-MOM-CEL-40'),
(2, 10, 'Celeste', 7, 2, 'JN-MOM-CEL-42'),
(2, 11, 'Celeste', 3, 2, 'JN-MOM-CEL-44'),
(3, 3, 'Gris Melange', 8, 2, 'BUZ-HD-GRS-M'),
(3, 4, 'Gris Melange', 10, 3, 'BUZ-HD-GRS-L'),
(3, 5, 'Gris Melange', 5, 2, 'BUZ-HD-GRS-XL'),
(3, 3, 'Negro', 9, 3, 'BUZ-HD-BLK-M'),
(3, 4, 'Negro', 11, 3, 'BUZ-HD-BLK-L'),
(4, 2, 'Negro', 4, 2, 'CMP-PUF-BLK-S'),
(4, 3, 'Negro', 6, 2, 'CMP-PUF-BLK-M'),
(4, 4, 'Negro', 5, 2, 'CMP-PUF-BLK-L'),
(5, 12, 'Negro', 20, 5, 'ACC-GOR-BLK-U');
