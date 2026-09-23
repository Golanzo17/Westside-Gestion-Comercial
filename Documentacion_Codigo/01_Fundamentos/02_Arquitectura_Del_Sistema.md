# 🏛️ Arquitectura del Sistema Westside Gestión Comercial

Tu proyecto no está hecho en un solo archivo gigante, sino organizado en **capas profesionales**. Cada carpeta tiene una responsabilidad clara y específica.

---

## El Flujo de la Información (Diagrama de Capas)

Cuando un usuario interactúa con la aplicación, la información viaja en este orden:

```text
[ 1. Pantalla (Forms / UI) ]
            │
            ▼  (Pide procesar una venta o buscar un producto)
[ 2. Lógica del Negocio (Services) ]
            │
            ▼  (Usa los objetos de datos: Entities / Models)
[ 3. Capa de Acceso a Datos (DatabaseHelper) ]
            │
            ▼  (Ejecuta SQL con parámetros seguros)
[ 4. Motor de Base de Datos (SQLite) ]
```

---

## Descripción de las Carpetas del Proyecto

### 1. `Forms/` (Formularios e Interfaz de Usuario)
* **¿Qué hay aquí?** Las ventanas que el usuario ve y toca en Windows:
  * `FrmLogin.vb`: Pantalla de inicio de sesión.
  * `FrmMain.vb`: Menú principal y navegación.
  * `FrmVentasPOS.vb`: Punto de venta (cobro rápido, lector de códigos).
  * `FrmProductos.vb`: Catálogo y alta de artículos.
  * `FrmStock.vb`: Control de inventario y ajustes manuales.
  * `FrmCaja.vb`: Apertura, arqueo y cierre de caja.
  * `FrmUsuarios.vb`: Gestión de usuarios, vendedores y permisos.
* **Regla de oro**: Los formularios NO deben escribir consultas SQL directamente. Solo capturan los datos que el usuario escribe, llaman al `Service` correspondiente y muestran los resultados o alertas.

---

### 2. `Services/` (Lógica de Negocio)
* **¿Qué hay aquí?** El "cerebro" del sistema. Aquí están las reglas del negocio:
  * `VentaService.vb`: ¿Hay stock suficiente para vender? ¿Cómo se calcula el ticket? ¿Cómo se anula una venta restituyendo el stock?
  * `CatalogService.vb`: ¿Cómo se guarda un producto y su matriz de talles y colores?
  * `AuthService.vb`: ¿Es correcta la contraseña? ¿Está bloqueado por 5 intentos fallidos?
  * `CajaService.vb`: ¿La caja ya está abierta? ¿Cuánto dinero debería haber en el cajón?
  * `UsuarioService.vb` y `ClienteService.vb`: Validación preventiva de unicidad de DNI y gestión de altas y bajas.
  * `ReporteService.vb`: Métricas de ventas, rankings comerciales e inteligencia de negocio.
  * `ConfiguracionService.vb`: Administración de datos comerciales y leyenda de tickets.

---

### 3. `Models/` (Entidades y Objetos)
* **¿Qué hay aquí?** En `Entities.vb` se definen las "plantillas" de datos:
  * `Producto`, `Venta`, `DetalleVenta`, `Cliente`, `Caja`, `Usuario`.
  * Son clases simples con propiedades (POCOs / DTOs) que permiten transportar la información entre las capas sin mezclar SQL con la interfaz.

---

### 4. `Data/` (Acceso a Bases de Datos)
* **¿Qué hay aquí?** En `DatabaseHelper.vb` está todo el código que abre conexiones, ejecuta `SELECT`, `INSERT`, `UPDATE`, y gestiona transacciones.
* **Gran ventaja de tu sistema**: Utiliza **SQLite** como motor relacional embebido (archivo `gestion_comercial.db`). Es completamente autónomo, no requiere servicios ni configuración previa, es portable y garantiza consistencia ACID total.

---

### 5. `Config/` (Configuración de la Aplicación)
* **¿Qué hay aquí?** `AppConfig.vb` lee el archivo `appsettings.json` al arrancar el programa, administrando la ruta de la base de datos y los parámetros de inicialización.

---

### 6. `UI/` (Diseño y Estilos)
* **¿Qué hay aquí?** `UITheme.vb` define la paleta de colores moderna (tonos oscuros, azules y grises elegantes), tipografías y métodos para dar estilo uniforme a botones, cajas de texto y tablas.
