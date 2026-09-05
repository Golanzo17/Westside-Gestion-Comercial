-- =======================================================================
-- SCRIPT DE MIGRACIÓN: DESDE E-COMMERCE (grupo5) A GESTIÓN COMERCIAL
-- =======================================================================
-- Este script permite importar automáticamente las categorías, talles,
-- productos y stock que ya existen en la base de datos de tu tienda web
-- hacia la nueva base de datos del Software de Gestión Comercial.
-- =======================================================================

USE `gestion_comercial_db`;

-- 1. IMPORTAR CATEGORÍAS
-- Reemplaza 'grupo5' por el nombre exacto de la BD de tu ecommerce si es diferente
INSERT INTO `gestion_comercial_db`.`categorias` (`id`, `nombre`, `descripcion`, `activo`)
SELECT 
    c.id, 
    c.nombre, 
    CONCAT('Importado desde e-commerce (slug: ', c.slug, ')'),
    1
FROM `grupo5`.`categorias` c
ON DUPLICATE KEY UPDATE `nombre` = VALUES(`nombre`);

-- 2. IMPORTAR TALLES
INSERT INTO `gestion_comercial_db`.`talles` (`id`, `nombre`, `orden`)
SELECT 
    t.id, 
    t.nombre, 
    t.id
FROM `grupo5`.`talles` t
ON DUPLICATE KEY UPDATE `nombre` = VALUES(`nombre`);

-- 3. IMPORTAR PRODUCTOS
-- Mapea los productos web asignando un código de barra si no lo tienen y calculando costo estimado
INSERT INTO `gestion_comercial_db`.`productos` 
(`id`, `codigo_barra`, `nombre`, `descripcion`, `categoria_id`, `precio_costo`, `precio_venta`, `porcentaje_ganancia`, `imagen_ruta`, `activo`, `created_at`)
SELECT 
    p.id,
    LPAD(p.id, 8, '77900000'), -- Genera código de barra con prefijo 779
    p.nombre,
    p.descripcion,
    p.categoria_id,
    ROUND(p.precio * 0.50, 2), -- Costo estimado al 50% del precio web (ajustable)
    p.precio,
    100.00,
    p.imagen_ruta,
    IFNULL(p.activo, 1),
    IFNULL(p.created_at, NOW())
FROM `grupo5`.`productos` p
ON DUPLICATE KEY UPDATE 
    `nombre` = VALUES(`nombre`),
    `precio_venta` = VALUES(`precio_venta`),
    `categoria_id` = VALUES(`categoria_id`);

-- 4. IMPORTAR RELACIÓN PRODUCTO-TALLE Y STOCK
INSERT INTO `gestion_comercial_db`.`producto_talles`
(`producto_id`, `talle_id`, `color`, `stock_actual`, `stock_minimo`, `sku_especifico`)
SELECT 
    pt.producto_id,
    pt.talle_id,
    'Único' AS color,
    pt.stock,
    2 AS stock_minimo,
    CONCAT('SKU-', pt.producto_id, '-', pt.talle_id)
FROM `grupo5`.`producto_talle` pt
ON DUPLICATE KEY UPDATE 
    `stock_actual` = VALUES(`stock_actual`);

-- Fin del script de migración
