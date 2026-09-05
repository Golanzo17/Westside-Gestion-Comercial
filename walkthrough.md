# Walkthrough: Software de Gestión Comercial de Indumentaria en Visual Basic y MySQL

Se ha completado el desarrollo del **Software de Gestión Comercial y Punto de Venta (POS)** para el local de ropa, construido en **Visual Basic (.NET WinForms)** con motor de base de datos **MySQL**, preparado con una arquitectura limpia por capas y un módulo de migración directa desde el e-commerce previo.

---

## 🏛 Arquitectura y Componentes Desarrollados

### 1. Base de Datos Relacional Robusta (MySQL)
Ubicada en [c:/Users/gonza/Desktop/Proyecto/Database/](file:///c:/Users/gonza/Desktop/Proyecto/Database/):
- [schema_mysql.sql](file:///c:/Users/gonza/Desktop/Proyecto/Database/schema_mysql.sql): Define 12 tablas relacionales con claves foráneas, integridad referencial y soporte `utf8mb4`:
  - `configuracion`, `usuarios`, `categorias`, `talles`, `productos`, `producto_talles` (matriz de stock), `clientes`, `cajas`, `movimientos_caja`, `ventas`, `detalle_ventas`, `movimientos_stock`.
- [seed_data.sql](file:///c:/Users/gonza/Desktop/Proyecto/Database/seed_data.sql): Carga categorías de indumentaria, talles estándar (XS a XXL y 36 a 44), cliente genérico (Consumidor Final), prendas de muestra y usuarios con contraseñas cifradas en SHA256.
- [migration_from_ecommerce.sql](file:///c:/Users/gonza/Desktop/Proyecto/Database/migration_from_ecommerce.sql): Script SQL para migrar directamente las categorías, talles, productos y stock desde la base de datos del e-commerce previo (`grupo5`).

### 2. Capa de Datos y Servicios en Visual Basic (.NET)
- [DatabaseHelper.vb](file:///c:/Users/gonza/Desktop/Proyecto/GestionComercial/Data/DatabaseHelper.vb): Manejo de conexiones con `MySqlConnector`, ejecución de consultas parametrizadas, transacciones seguras e **inicialización automática de tablas**.
- [AppConfig.vb](file:///c:/Users/gonza/Desktop/Proyecto/GestionComercial/Config/AppConfig.vb): Carga y guardado de parámetros en `appsettings.json`.
- [Entities.vb](file:///c:/Users/gonza/Desktop/Proyecto/GestionComercial/Models/Entities.vb): Modelos POCO desacoplados.
- **Servicios de Negocio**:
  - [AuthService.vb](file:///c:/Users/gonza/Desktop/Proyecto/GestionComercial/Services/AuthService.vb): Control de sesión, roles (Administrador, Vendedor, Cajero) y verificación de contraseñas.
  - [CatalogService.vb](file:///c:/Users/gonza/Desktop/Proyecto/GestionComercial/Services/CatalogService.vb): CRUD de productos, categorías, talles, matriz de stock y ajustes.
  - [VentaService.vb](file:///c:/Users/gonza/Desktop/Proyecto/GestionComercial/Services/VentaService.vb): Procesamiento transaccional de ventas, deducción y restitución de stock por talle, actualización de caja e historial.
  - [CajaService.vb](file:///c:/Users/gonza/Desktop/Proyecto/GestionComercial/Services/CajaService.vb): Apertura, cierre de turno, arqueo de billetes y movimientos varios.
  - [ClienteService.vb](file:///c:/Users/gonza/Desktop/Proyecto/GestionComercial/Services/ClienteService.vb): Directorio de clientes y búsqueda por DNI.
  - [ReporteService.vb](file:///c:/Users/gonza/Desktop/Proyecto/GestionComercial/Services/ReporteService.vb): Indicadores de ventas, recaudación y ranking de prendas.
  - [EcommerceSyncService.vb](file:///c:/Users/gonza/Desktop/Proyecto/GestionComercial/Services/EcommerceSyncService.vb): Diagnóstico e importación automática desde la base de datos de la web.

### 3. Interfaz de Usuario Moderna (WinForms)
- [UITheme.vb](file:///c:/Users/gonza/Desktop/Proyecto/GestionComercial/UI/UITheme.vb): Paleta de colores contemporánea (Slate 900 `#0F172A`, Indigo `#4F46E5`, Emerald `#10B981`), estilo uniforme para botones, DataGridViews y tarjetas KPI.
- [FrmLogin.vb](file:///c:/Users/gonza/Desktop/Proyecto/GestionComercial/Forms/FrmLogin.vb): Autenticación y acceso con indicador de estado de conexión.
- [FrmMain.vb](file:///c:/Users/gonza/Desktop/Proyecto/GestionComercial/Forms/FrmMain.vb): Dashboard principal con menú lateral, reloj en tiempo real, status de caja y KPIs de ventas y stock crítico.
- [FrmVentasPOS.vb](file:///c:/Users/gonza/Desktop/Proyecto/GestionComercial/Forms/FrmVentasPOS.vb): Punto de venta ágil con selector dinámico de talles con stock en vivo, calculadora de vuelto, multiformas de pago y emisión de ticket térmico.
- [FrmProductos.vb](file:///c:/Users/gonza/Desktop/Proyecto/GestionComercial/Forms/FrmProductos.vb): Catálogo de prendas y editor modal con matriz de stock por talle.
- [FrmStock.vb](file:///c:/Users/gonza/Desktop/Proyecto/GestionComercial/Forms/FrmStock.vb): Semáforo de reposición de inventario y auditoría de movimientos.
- [FrmClientes.vb](file:///c:/Users/gonza/Desktop/Proyecto/GestionComercial/Forms/FrmClientes.vb): Gestión de clientes y ventas asociadas.
- [FrmCaja.vb](file:///c:/Users/gonza/Desktop/Proyecto/GestionComercial/Forms/FrmCaja.vb): Apertura, registro de gastos/ingresos y arqueo de caja con cálculo de diferencias.
- [FrmReportes.vb](file:///c:/Users/gonza/Desktop/Proyecto/GestionComercial/Forms/FrmReportes.vb): Métricas por fecha, ranking de modelos y anulación de tickets.
- [FrmSincronizacionEcommerce.vb](file:///c:/Users/gonza/Desktop/Proyecto/GestionComercial/Forms/FrmSincronizacionEcommerce.vb): Asistente para importar el catálogo desde la tienda web en un clic.
- [FrmConfiguracion.vb](file:///c:/Users/gonza/Desktop/Proyecto/GestionComercial/Forms/FrmConfiguracion.vb): Configuración de conexión MySQL y datos fiscales para los tickets.

---

## 🧪 Verificación y Pruebas Realizadas

### Compilación de la Solución
Ejecución de compilación mediante `dotnet build`:
```
GestionComercial -> C:\Users\gonza\Desktop\Proyecto\GestionComercial\bin\Debug\net10.0-windows\GestionComercial.dll
Compilación correcta.
    0 Advertencia(s)
    0 Errores
```

Tanto la solución estándar `GestionComercial.sln` como el archivo de proyecto compilan con cero advertencias y cero errores.

---

## 🔑 Credenciales Predeterminadas

| Rol | Usuario | Contraseña |
|---|---|---|
| **Administrador** | `admin` | `admin123` |
| **Vendedor / Cajero** | `vendedor` | `1234` |
