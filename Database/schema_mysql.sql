-- =======================================================================
-- BASE DE DATOS: gestion_comercial_db
-- Sistema de Gestión Comercial para Local de Ropa / Indumentaria
-- =======================================================================

CREATE DATABASE IF NOT EXISTS `gestion_comercial_db` 
CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;

USE `gestion_comercial_db`;

-- Desactivar chequeo de claves foráneas temporalmente para reconstrucción limpia
SET FOREIGN_KEY_CHECKS = 0;

DROP TABLE IF EXISTS `detalle_ventas`;
DROP TABLE IF EXISTS `ventas`;
DROP TABLE IF EXISTS `movimientos_stock`;
DROP TABLE IF EXISTS `movimientos_caja`;
DROP TABLE IF EXISTS `cajas`;
DROP TABLE IF EXISTS `producto_talles`;
DROP TABLE IF EXISTS `productos`;
DROP TABLE IF EXISTS `talles`;
DROP TABLE IF EXISTS `categorias`;
DROP TABLE IF EXISTS `clientes`;
DROP TABLE IF EXISTS `usuarios`;
DROP TABLE IF EXISTS `configuracion`;

SET FOREIGN_KEY_CHECKS = 1;

-- -----------------------------------------------------------------------
-- 1. TABLA: configuracion (Datos del comercio / local)
-- -----------------------------------------------------------------------
CREATE TABLE `configuracion` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `nombre_comercio` VARCHAR(150) NOT NULL DEFAULT 'Mi Local de Ropa',
    `cuit` VARCHAR(20) DEFAULT '',
    `direccion` VARCHAR(255) DEFAULT '',
    `telefono` VARCHAR(50) DEFAULT '',
    `email` VARCHAR(100) DEFAULT '',
    `condicion_iva` VARCHAR(50) DEFAULT 'Responsable Inscripto',
    `mensaje_ticket` VARCHAR(255) DEFAULT '¡Gracias por su compra! Cambios dentro de los 30 días con ticket.',
    `moneda_simbolo` VARCHAR(10) DEFAULT '$',
    `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- -----------------------------------------------------------------------
-- 2. TABLA: usuarios (Personal con roles)
-- -----------------------------------------------------------------------
CREATE TABLE `usuarios` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `username` VARCHAR(50) NOT NULL UNIQUE,
    `dni` VARCHAR(20) UNIQUE NULL,
    `password_hash` VARCHAR(255) NOT NULL,
    `nombre_completo` VARCHAR(100) NOT NULL,
    `rol` ENUM('Administrador', 'Vendedor', 'Cajero') NOT NULL DEFAULT 'Vendedor',
    `activo` TINYINT(1) NOT NULL DEFAULT 1,
    `ultimo_login` DATETIME NULL,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- -----------------------------------------------------------------------
-- 3. TABLA: categorias (Categorías de indumentaria)
-- -----------------------------------------------------------------------
CREATE TABLE `categorias` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `nombre` VARCHAR(100) NOT NULL UNIQUE,
    `descripcion` VARCHAR(255) NULL,
    `activo` TINYINT(1) NOT NULL DEFAULT 1,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- -----------------------------------------------------------------------
-- 4. TABLA: talles (Talles estándar y numéricos)
-- -----------------------------------------------------------------------
CREATE TABLE `talles` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `nombre` VARCHAR(20) NOT NULL UNIQUE,
    `orden` INT NOT NULL DEFAULT 0,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- -----------------------------------------------------------------------
-- 5. TABLA: productos (Prendas y artículos principales)
-- -----------------------------------------------------------------------
CREATE TABLE `productos` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `codigo_barra` VARCHAR(50) UNIQUE NOT NULL,
    `nombre` VARCHAR(150) NOT NULL,
    `descripcion` TEXT NULL,
    `categoria_id` INT NOT NULL,
    `precio_costo` DECIMAL(12,2) NOT NULL DEFAULT 0.00,
    `precio_venta` DECIMAL(12,2) NOT NULL DEFAULT 0.00,
    `porcentaje_ganancia` DECIMAL(5,2) DEFAULT 50.00,
    `imagen_ruta` VARCHAR(255) NULL,
    `activo` TINYINT(1) NOT NULL DEFAULT 1,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    CONSTRAINT `fk_prod_categoria` FOREIGN KEY (`categoria_id`) REFERENCES `categorias`(`id`) ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- -----------------------------------------------------------------------
-- 6. TABLA: producto_talles (Matriz de stock por Talle y Color)
-- -----------------------------------------------------------------------
CREATE TABLE `producto_talles` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `producto_id` INT NOT NULL,
    `talle_id` INT NOT NULL,
    `color` VARCHAR(50) NOT NULL DEFAULT 'Único',
    `stock_actual` INT NOT NULL DEFAULT 0,
    `stock_minimo` INT NOT NULL DEFAULT 2,
    `sku_especifico` VARCHAR(60) NULL,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    UNIQUE KEY `uk_prod_talle_color` (`producto_id`, `talle_id`, `color`),
    CONSTRAINT `fk_pt_producto` FOREIGN KEY (`producto_id`) REFERENCES `productos`(`id`) ON DELETE CASCADE,
    CONSTRAINT `fk_pt_talle` FOREIGN KEY (`talle_id`) REFERENCES `talles`(`id`) ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- -----------------------------------------------------------------------
-- 7. TABLA: clientes (Directorio de compradores)
-- -----------------------------------------------------------------------
CREATE TABLE `clientes` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `dni_cuit` VARCHAR(20) NOT NULL UNIQUE,
    `nombre` VARCHAR(80) NOT NULL,
    `apellido` VARCHAR(80) NOT NULL,
    `telefono` VARCHAR(50) NULL,
    `email` VARCHAR(100) NULL,
    `direccion` VARCHAR(200) NULL,
    `ciudad` VARCHAR(100) NULL,
    `notas` TEXT NULL,
    `activo` TINYINT(1) NOT NULL DEFAULT 1,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- -----------------------------------------------------------------------
-- 8. TABLA: cajas (Apertura, movimientos y arqueo diario)
-- -----------------------------------------------------------------------
CREATE TABLE `cajas` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `usuario_id` INT NOT NULL,
    `fecha_apertura` DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    `fecha_cierre` DATETIME NULL,
    `monto_inicial` DECIMAL(12,2) NOT NULL DEFAULT 0.00,
    `total_ventas_efectivo` DECIMAL(12,2) NOT NULL DEFAULT 0.00,
    `total_ventas_digital` DECIMAL(12,2) NOT NULL DEFAULT 0.00,
    `total_ingresos` DECIMAL(12,2) NOT NULL DEFAULT 0.00,
    `total_egresos` DECIMAL(12,2) NOT NULL DEFAULT 0.00,
    `monto_esperado` DECIMAL(12,2) NOT NULL DEFAULT 0.00,
    `monto_real` DECIMAL(12,2) NULL,
    `diferencia` DECIMAL(12,2) NULL,
    `estado` ENUM('Abierta', 'Cerrada') NOT NULL DEFAULT 'Abierta',
    `observaciones` TEXT NULL,
    CONSTRAINT `fk_caja_usuario` FOREIGN KEY (`usuario_id`) REFERENCES `usuarios`(`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- -----------------------------------------------------------------------
-- 9. TABLA: movimientos_caja (Ingresos y egresos manuales)
-- -----------------------------------------------------------------------
CREATE TABLE `movimientos_caja` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `caja_id` INT NOT NULL,
    `usuario_id` INT NOT NULL,
    `fecha` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `tipo` ENUM('Ingreso', 'Egreso') NOT NULL,
    `concepto` VARCHAR(200) NOT NULL,
    `monto` DECIMAL(12,2) NOT NULL,
    `referencia` VARCHAR(100) NULL,
    CONSTRAINT `fk_mc_caja` FOREIGN KEY (`caja_id`) REFERENCES `cajas`(`id`) ON DELETE CASCADE,
    CONSTRAINT `fk_mc_usuario` FOREIGN KEY (`usuario_id`) REFERENCES `usuarios`(`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- -----------------------------------------------------------------------
-- 10. TABLA: ventas (Encabezado de comprobantes)
-- -----------------------------------------------------------------------
CREATE TABLE `ventas` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `numero_ticket` VARCHAR(30) NOT NULL UNIQUE,
    `fecha` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `usuario_id` INT NOT NULL,
    `cliente_id` INT NOT NULL,
    `caja_id` INT NOT NULL,
    `tipo_comprobante` ENUM('Ticket', 'Factura B', 'Factura A', 'Presupuesto') NOT NULL DEFAULT 'Ticket',
    `metodo_pago` ENUM('Efectivo', 'Tarjeta Débito', 'Tarjeta Crédito', 'Transferencia', 'Múltiple') NOT NULL DEFAULT 'Efectivo',
    `subtotal` DECIMAL(12,2) NOT NULL DEFAULT 0.00,
    `descuento_porcentaje` DECIMAL(5,2) NOT NULL DEFAULT 0.00,
    `descuento_monto` DECIMAL(12,2) NOT NULL DEFAULT 0.00,
    `recargo_monto` DECIMAL(12,2) NOT NULL DEFAULT 0.00,
    `total` DECIMAL(12,2) NOT NULL DEFAULT 0.00,
    `monto_abonado` DECIMAL(12,2) NOT NULL DEFAULT 0.00,
    `vuelto` DECIMAL(12,2) NOT NULL DEFAULT 0.00,
    `estado` ENUM('Completada', 'Anulada') NOT NULL DEFAULT 'Completada',
    `observaciones` VARCHAR(255) NULL,
    CONSTRAINT `fk_venta_usuario` FOREIGN KEY (`usuario_id`) REFERENCES `usuarios`(`id`),
    CONSTRAINT `fk_venta_cliente` FOREIGN KEY (`cliente_id`) REFERENCES `clientes`(`id`),
    CONSTRAINT `fk_venta_caja` FOREIGN KEY (`caja_id`) REFERENCES `cajas`(`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- -----------------------------------------------------------------------
-- 11. TABLA: detalle_ventas (Renglones de prendas vendidas)
-- -----------------------------------------------------------------------
CREATE TABLE `detalle_ventas` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `venta_id` INT NOT NULL,
    `producto_id` INT NOT NULL,
    `talle_id` INT NOT NULL,
    `color` VARCHAR(50) NOT NULL DEFAULT 'Único',
    `codigo_barra` VARCHAR(50) NOT NULL,
    `descripcion_articulo` VARCHAR(200) NOT NULL,
    `precio_unitario` DECIMAL(12,2) NOT NULL,
    `costo_unitario` DECIMAL(12,2) NOT NULL DEFAULT 0.00,
    `cantidad` INT NOT NULL,
    `subtotal` DECIMAL(12,2) NOT NULL,
    CONSTRAINT `fk_dv_venta` FOREIGN KEY (`venta_id`) REFERENCES `ventas`(`id`) ON DELETE CASCADE,
    CONSTRAINT `fk_dv_producto` FOREIGN KEY (`producto_id`) REFERENCES `productos`(`id`),
    CONSTRAINT `fk_dv_talle` FOREIGN KEY (`talle_id`) REFERENCES `talles`(`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- -----------------------------------------------------------------------
-- 12. TABLA: movimientos_stock (Auditoría de inventario)
-- -----------------------------------------------------------------------
CREATE TABLE `movimientos_stock` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `producto_id` INT NOT NULL,
    `talle_id` INT NOT NULL,
    `color` VARCHAR(50) NOT NULL DEFAULT 'Único',
    `tipo_movimiento` ENUM('Venta', 'Ingreso_Compra', 'Ajuste_Manual', 'Devolucion', 'Anulacion_Venta') NOT NULL,
    `cantidad` INT NOT NULL,
    `stock_anterior` INT NOT NULL,
    `stock_posterior` INT NOT NULL,
    `motivo` VARCHAR(255) NULL,
    `usuario_id` INT NOT NULL,
    `fecha` DATETIME DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT `fk_ms_producto` FOREIGN KEY (`producto_id`) REFERENCES `productos`(`id`),
    CONSTRAINT `fk_ms_talle` FOREIGN KEY (`talle_id`) REFERENCES `talles`(`id`),
    CONSTRAINT `fk_ms_usuario` FOREIGN KEY (`usuario_id`) REFERENCES `usuarios`(`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- Índices adicionales para búsquedas instantáneas
CREATE INDEX `idx_prod_codigo` ON `productos` (`codigo_barra`);
CREATE INDEX `idx_prod_nombre` ON `productos` (`nombre`);
CREATE INDEX `idx_ventas_fecha` ON `ventas` (`fecha`);
CREATE INDEX `idx_ventas_ticket` ON `ventas` (`numero_ticket`);
