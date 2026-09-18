-- =======================================================================
-- DATOS INICIALES Y CATÁLOGO DE EJEMPLO PARA LOCAL DE ROPA
-- Base de Datos: gestion_comercial_db
-- =======================================================================

USE `gestion_comercial_db`;

-- 1. CONFIGURACIÓN DEL COMERCIO
INSERT INTO `configuracion` (`id`, `nombre_comercio`, `cuit`, `direccion`, `telefono`, `email`, `condicion_iva`, `mensaje_ticket`, `moneda_simbolo`)
VALUES (1, 'Boutique Urbana - Local de Ropa', '20-38491234-9', 'Av. Santa Fe 1540, CABA', '11-4567-8901', 'ventas@boutiqueurbana.com', 'Responsable Inscripto', '¡Gracias por elegirnos! Cambios dentro de los 30 días presentando este ticket.', '$')
ON DUPLICATE KEY UPDATE `nombre_comercio` = VALUES(`nombre_comercio`);

-- 2. USUARIOS INICIALES (Contraseñas con hash SHA256)
-- admin / admin123  -> 240be518fabd2724ddb6f04eeb1da5967448d7e831c08c8fa822809f74c720a9
-- vendedor / 1234   -> 03ac674216f3e15c761ee1a5e255f067953623c8b388b4459e13f978d7c846f4
-- gerente / gerente123 -> ecfba551324356e5bd27b548adf36b728783f60d9b573d142caac7baad62be49
INSERT INTO `usuarios` (`id`, `username`, `password_hash`, `nombre_completo`, `rol`, `activo`)
VALUES 
(1, 'admin', '240be518fabd2724ddb6f04eeb1da5967448d7e831c08c8fa822809f74c720a9', 'Administrador General', 'Administrador', 1),
(2, 'vendedor', '03ac674216f3e15c761ee1a5e255f067953623c8b388b4459e13f978d7c846f4', 'Vendedor de Turno', 'Vendedor', 1),
(3, 'gerente', 'ecfba551324356e5bd27b548adf36b728783f60d9b573d142caac7baad62be49', 'Gerente General', 'Gerente', 1)
ON DUPLICATE KEY UPDATE `nombre_completo` = VALUES(`nombre_completo`);

-- 3. TALLES DE ROPA
INSERT INTO `talles` (`id`, `nombre`, `orden`) VALUES
(1, 'XS', 1),
(2, 'S', 2),
(3, 'M', 3),
(4, 'L', 4),
(5, 'XL', 5),
(6, 'XXL', 6),
(7, '36', 7),
(8, '38', 8),
(9, '40', 9),
(10, '42', 10),
(11, '44', 11),
(12, 'Único', 12)
ON DUPLICATE KEY UPDATE `orden` = VALUES(`orden`);

-- 4. CATEGORÍAS DE INDUMENTARIA
INSERT INTO `categorias` (`id`, `nombre`, `descripcion`, `activo`) VALUES
(1, 'Remeras y Tops', 'Remeras estampadas, lisas, musculosas y tops', 1),
(2, 'Pantalones y Jeans', 'Jeans mom, skinny, wide leg, joggers y pantalones de vestir', 1),
(3, 'Buzos y Camperas', 'Buzos hoodie, camperas de abrigo, bombers y cardigans', 1),
(4, 'Vestidos y Polleras', 'Vestidos de fiesta, casuales y polleras', 1),
(5, 'Accesorios y Calzado', 'Cinturones, gorras, medias y calzado urbano', 1)
ON DUPLICATE KEY UPDATE `nombre` = VALUES(`nombre`);

-- 5. CLIENTE PREDETERMINADO (Consumidor Final)
INSERT INTO `clientes` (`id`, `dni_cuit`, `nombre`, `apellido`, `telefono`, `email`, `direccion`, `ciudad`, `notas`, `activo`)
VALUES 
(1, '00000000', 'Consumidor', 'Final', '000-0000', 'consumidor@final.com', 'Mostrador', 'Local', 'Cliente genérico de mostrador', 1),
(2, '35123456', 'Gonzalo', 'Fernández', '11-3456-7890', 'gonzalo@email.com', 'Calle Falsa 123', 'Buenos Aires', 'Cliente frecuente', 1)
ON DUPLICATE KEY UPDATE `nombre` = VALUES(`nombre`);

-- 6. PRODUCTOS DE DEMOSTRACIÓN
INSERT INTO `productos` (`id`, `codigo_barra`, `nombre`, `descripcion`, `categoria_id`, `precio_costo`, `precio_venta`, `porcentaje_ganancia`, `activo`) VALUES
(1, '779001001', 'Remera Oversize Vintage Acid Wash', 'Remera de algodón peinado 100% corte oversize', 1, 9500.00, 18500.00, 94.74, 1),
(2, '779001002', 'Jean Mom Fit Celeste Nevado', 'Jean rígido tiro alto con calce relajado', 2, 18000.00, 36000.00, 100.00, 1),
(3, '779001003', 'Buzo Hoodie Canguro Frisa Pesada', 'Buzo con capucha y bolsillo delantero unisex', 3, 22000.00, 42500.00, 93.18, 1),
(4, '779001004', 'Campera Puffer Corta Negra', 'Campera térmica impermeable con cierre reforzado', 3, 35000.00, 68000.00, 94.29, 1),
(5, '779001005', 'Gorra Trucker Urbana Bordada', 'Gorra con visera curva y red respirable', 5, 4500.00, 9800.00, 117.78, 1)
ON DUPLICATE KEY UPDATE `nombre` = VALUES(`nombre`);

-- 7. STOCK POR TALLE Y COLOR (MATRIZ DE INDUMENTARIA)
INSERT INTO `producto_talles` (`producto_id`, `talle_id`, `color`, `stock_actual`, `stock_minimo`, `sku_especifico`) VALUES
-- Remera Oversize (S, M, L, XL en Negro y Blanco)
(1, 2, 'Negro', 10, 3, 'REM-OV-BLK-S'),
(1, 3, 'Negro', 15, 3, 'REM-OV-BLK-M'),
(1, 4, 'Negro', 8, 3, 'REM-OV-BLK-L'),
(1, 5, 'Negro', 5, 2, 'REM-OV-BLK-XL'),
(1, 2, 'Blanco', 8, 3, 'REM-OV-WHT-S'),
(1, 3, 'Blanco', 12, 3, 'REM-OV-WHT-M'),
(1, 4, 'Blanco', 6, 2, 'REM-OV-WHT-L'),

-- Jean Mom Fit (Talles 36, 38, 40, 42, 44 en Celeste)
(2, 7, 'Celeste', 4, 2, 'JN-MOM-CEL-36'),
(2, 8, 'Celeste', 10, 3, 'JN-MOM-CEL-38'),
(2, 9, 'Celeste', 12, 3, 'JN-MOM-CEL-40'),
(2, 10, 'Celeste', 7, 2, 'JN-MOM-CEL-42'),
(2, 11, 'Celeste', 3, 2, 'JN-MOM-CEL-44'),

-- Buzo Hoodie (M, L, XL en Gris Melange y Negro)
(3, 3, 'Gris Melange', 8, 2, 'BUZ-HD-GRS-M'),
(3, 4, 'Gris Melange', 10, 3, 'BUZ-HD-GRS-L'),
(3, 5, 'Gris Melange', 5, 2, 'BUZ-HD-GRS-XL'),
(3, 3, 'Negro', 9, 3, 'BUZ-HD-BLK-M'),
(3, 4, 'Negro', 11, 3, 'BUZ-HD-BLK-L'),

-- Campera Puffer (S, M, L en Negro)
(4, 2, 'Negro', 4, 2, 'CMP-PUF-BLK-S'),
(4, 3, 'Negro', 6, 2, 'CMP-PUF-BLK-M'),
(4, 4, 'Negro', 5, 2, 'CMP-PUF-BLK-L'),

-- Gorra Trucker (Talle Único en Negro)
(5, 12, 'Negro', 20, 5, 'ACC-GOR-BLK-U')
ON DUPLICATE KEY UPDATE `stock_actual` = VALUES(`stock_actual`);
