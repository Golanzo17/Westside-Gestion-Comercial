# 👕 Explicación de CatalogService.vb (Catálogo, Talles y Stock)

* **Ubicación**: [GestionComercial/Services/CatalogService.vb]
* **Propósito**: Gestionar todo el catálogo de indumentaria: dar de alta prendas, categorías, talles, manejar la matriz de variantes (talle/color) y realizar ajustes manuales de stock por inventario o compras.

---

## 🔍 Explicación Sección por Sección

### 1. Sección Categorías
* **`GetCategorias(soloActivas) As List(Of Categoria)`**:
  * Devuelve la lista de categorías ordenadas alfabéticamente. Si `soloActivas = True`, excluye las inactivas.
* **`GuardarCategoria(cat, ByRef errorMessage) As Boolean`**:
  * Valida que el nombre no esté vacío.
  * Si `cat.Id = 0`, ejecuta un `INSERT` para crear una nueva categoría.
  * Si `cat.Id > 0`, ejecuta un `UPDATE` para modificar la existente.

---

### 2. Sección Talles
* **`GetTalles() As List(Of Talle)`**:
  * Devuelve los talles ordenados por la columna `orden` (para que "S" aparezca antes de "M" y "L", en vez de ordenarse alfabéticamente).
* **`GuardarTalle(talle, ByRef errorMessage) As Boolean`**:
  * Guarda o actualiza un talle con su número de orden.

---

### 3. Sección Productos y Búsqueda

* **`GetProductos(filtro, categoriaId, soloActivos) As List(Of Producto)`**:
  * Es el buscador del catálogo. Construye dinámicamente la consulta SQL con los filtros que el usuario elija en pantalla (por texto de nombre/código de barra, por categoría o solo activos).
  * Hace un `LEFT JOIN producto_talles` para calcular en tiempo real el stock total sumando las unidades de todos los talles.
* **`GetProductoPorCodigo(codigo) As Producto`**:
  * Usado por el **Lector de Código de Barras** en el Punto de Venta.
  * Busca si el código coincide con el código de barras escaneado o con el ID numérico del producto. Si lo encuentra, carga el producto con toda su matriz de talles y stock.

---

### 4. `GuardarProducto(prod, ByRef errorMessage) As Boolean`
Es la operación central del catálogo:
1. Valida con cláusula de guarda que el producto tenga nombre y categoría.
2. Abre una transacción de base de datos.
3. Inserta o actualiza el encabezado del producto (precios, código, margen de ganancia).
4. Si es un producto nuevo (`prod.Id = 0`), recupera el `ID` generado automáticamente.
5. **`GuardarMatrizTalles(...)`**:
   * Recorre la lista de talles asociados a la prenda.
   * Utiliza la sintaxis SQL adecuada (`ON CONFLICT` en SQLite o `ON DUPLICATE KEY UPDATE` en MySQL) para insertar las combinaciones nuevas o actualizar las existentes en una sola pasada eficiente.

---

### 5. `AjustarStockManual(...) As Boolean`
Permite a los encargados corregir el stock cuando hacen recuentos físicos o cuando ingresa un pedido del proveedor:
* Parámetros: `productoId`, `talleId`, `color`, `cantidadDelta`, `motivo`, `usuarioId`.
* Si `cantidadDelta` es positivo (ej: `+10`), es un ingreso por compra; si es negativo (ej: `-2`), es una merma o rotura.
* Si la fila talle/color no existía en la base de datos, la crea automáticamente con stock en 0 antes de sumar.
* Registra la auditoría en `movimientos_stock` con `tipo = 'Ingreso_Compra'` o `'Ajuste_Manual'`.

---

### 6. `EliminarProducto(id, ByRef errorMessage) As Boolean`
* **¿Qué es un Soft Delete (Borrado Lógico)?**:
  En un sistema comercial **nunca se debe hacer `DELETE` físico** sobre un producto que ya tuvo ventas, porque rompería el historial contable del local.
  En su lugar, este método hace `UPDATE productos SET activo = 0 WHERE id = @id`. La prenda ya no aparecerá para vender ni en el catálogo, pero los tickets pasados conservan intacta su información histórica.
