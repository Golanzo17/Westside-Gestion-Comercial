# 🌐 Explicación de EcommerceSyncService.vb (Sincronización Web)

* **Ubicación**: [GestionComercial/Services/EcommerceSyncService.vb]
* **Propósito**: Conectar el software de gestión local con la base de datos MySQL de una tienda web (E-commerce) para importar el catálogo de prendas, categorías, talles y stock existente.

---

## 🔍 Explicación Función por Función

### 1. `TestEcommerceConnection(...) As Boolean`
* Valida parámetros (`host`, `port`, `dbName`) con cláusulas de guarda.
* Intenta abrir una conexión a la base de datos remota con un timeout de 5 segundos.
* Si conecta, devuelve `True`; si el servidor está apagado o las credenciales son incorrectas, devuelve `False` y llena `errorMessage`.

---

### 2. `GetEcommerceStats(...) As Dictionary(Of String, Integer)`
* Se ejecuta al presionar "Analizar Web".
* Realiza consultas `COUNT(*)` en la web para contar:
  * Cuántas categorías hay en la tienda online.
  * Cuántos talles existen.
  * Cuántos productos están publicados.
  * La suma total del stock en línea.
* Devuelve un `Dictionary` que la pantalla utiliza para mostrar las estadísticas previas a la migración.

---

### 3. `ImportCatalogFromEcommerce(...) As Boolean`
Es el proceso de migración masiva en dos fases:

#### Fase 1: Lectura desde la Web (MySQL)
Abre la conexión a la base web y llena 4 tablas en memoria (`DataTable`):
```vb
FillDataTable(srcConn, "SELECT * FROM `categorias`;", dtCategorias)
FillDataTable(srcConn, "SELECT * FROM `talles`;", dtTalles)
FillDataTable(srcConn, "SELECT * FROM `productos`;", dtProductos)
TryFillDataTable(srcConn, "SELECT * FROM `producto_talle`;", dtStock)
```

#### Fase 2: Inserción Atómica en el Software Local
Abre una transacción atómica local y delega en métodos especializados:
* **`ImportCategorias`**: Inserta las categorías web manteniendo sus mismos IDs originales.
* **`ImportTalles`**: Inserta los talles web respetando el orden.
* **`ImportProductos`**:
  * Toma el precio de la web (`precio_venta`).
  * Asigna un costo estimado (`precioWeb * 0.5D`, ganancia del 100%).
  * Genera un código de barras estándar nacional tipo EAN (`"779" & prodId.ToString("D6")`).
  * Inserta o actualiza el producto en la base local.
* **`ImportStock`**: Inserta la relación entre cada producto y su talle con el stock real que tenía en la web.
* Si todo sale bien, ejecuta `trans.Commit()` y genera un reporte detallado con la cantidad exacta de registros importados.
