-- ============================================================
-- MIGRACIÓN: Ampliar datos personales de usuarios y clientes
-- Opción A: nombre_completo se migra a la columna nombre
-- Ejecutar UNA sola vez sobre la base de datos existente
-- ============================================================

-- ---- TABLA usuarios ----------------------------------------
ALTER TABLE usuarios ADD COLUMN nombre    TEXT NOT NULL DEFAULT '';
ALTER TABLE usuarios ADD COLUMN apellido  TEXT NOT NULL DEFAULT '';
ALTER TABLE usuarios ADD COLUMN telefono  TEXT NULL;
ALTER TABLE usuarios ADD COLUMN email     TEXT NULL;
ALTER TABLE usuarios ADD COLUMN direccion TEXT NULL;
ALTER TABLE usuarios ADD COLUMN ciudad    TEXT NULL;
ALTER TABLE usuarios ADD COLUMN notas     TEXT NULL;
ALTER TABLE usuarios ADD COLUMN fecha_nacimiento TEXT NULL;

-- Opción A: migrar el valor de nombre_completo a la nueva columna nombre
UPDATE usuarios SET nombre = nombre_completo WHERE nombre = '' OR nombre IS NULL;

-- ---- TABLA clientes ----------------------------------------
ALTER TABLE clientes ADD COLUMN fecha_nacimiento TEXT NULL;
