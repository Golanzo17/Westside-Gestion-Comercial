# 👗 Software de Gestión Comercial para Local de Ropa

Sistema integral de gestión de ventas, stock, caja y clientes desarrollado en **Visual Basic (.NET WinForms)** con motor de base de datos **MySQL**, diseñado específicamente para locales de indumentaria y preparado para importar/sincronizar los datos del e-commerce web previo.

---

## 🚀 Inicio Rápido

### 1. Requisitos Previos
- **.NET SDK** (.NET 8 LTS o .NET 10).
- **Servidor MySQL** (MySQL Community Server, MariaDB, XAMPP, Laragon o Herd).
- **Visual Studio 2022/2026** o **Visual Studio Code**.

### 2. Base de Datos MySQL
1. En tu gestor MySQL preferido (phpMyAdmin, MySQL Workbench, DBeaver o consola), puedes ejecutar el script:
   - `Database/schema_mysql.sql`: Crea la base de datos `gestion_comercial_db` y todas sus tablas con índices y relaciones.
   - `Database/seed_data.sql`: Carga datos iniciales (talles estándar XS a XXL y 36 a 44, categorías de ropa, usuarios y catálogo de muestra).
2. **Inicialización Automática**: Si prefieres, el propio software cuenta con una herramienta en el menú de **Configuración** (`⚙ Conexión MySQL -> Crear / Inicializar Tablas`) que crea la base de datos y la estructura automáticamente sin necesidad de hacerlo a mano.

### 3. Ejecutar la Aplicación
- Desde **Visual Studio**: Abre `GestionComercial.sln` y presiona **F5**.
- Desde **Terminal / PowerShell**:
  ```powershell
  cd c:\Users\gonza\Desktop\Proyecto\GestionComercial
  dotnet run
  ```

### 4. Credenciales de Acceso por Defecto
| Rol | Usuario | Contraseña |
|---|---|---|
| **Administrador** | `admin` | `admin123` |
| **Vendedor / Cajero** | `vendedor` | `1234` |

---

## 🌟 Módulos y Funcionalidades Principales

### 🛒 1. Punto de Venta (POS) y Facturación
- **Lectura Ágil**: Búsqueda por código de barras / SKU o nombre de la prenda.
- **Matriz de Talles en Vivo**: Visualización interactiva del stock disponible por talle en botones dinámicos (`[S (8)]`, `[M (12)]`, `[L (0 - Sin Stock)]`).
- **Carrito de Venta**: Cálculo automático de subtotales, descuentos porcentuales y recargos.
- **Múltiples Medios de Pago**: Efectivo (con calculadora de vuelto en tiempo real), Tarjeta Débito, Tarjeta Crédito y Transferencia / QR.
- **Emisión de Ticket**: Vista previa e impresión de ticket térmico con datos fiscales del comercio y política de cambios.
- **Transaccionalidad Estricta**: Cada venta descuenta automáticamente las unidades del talle y color específico, registra la auditoría de stock e impacta en la caja del turno.

### 👕 2. Catálogo de Ropa y Matriz de Stock
- ABM completo de prendas, categorías y talles.
- Cálculo automático de márgenes de ganancia (`Precio Venta = Costo + Margen %`).
- **Matriz de Talles x Color**: Permite definir stock individual y niveles de stock mínimo por cada talle (XS, S, M, L, XL, 36, 38, etc.).

### 📦 3. Control de Stock y Reposición
- **Semáforo de Alertas**: Detección automática de artículos agotados o en nivel crítico por debajo del stock mínimo.
- **Ingreso de Mercadería**: Registro de compras a proveedores con actualización directa del inventario.
- **Auditoría de Movimientos**: Historial detallado de cada entrada, salida, venta o ajuste manual con fecha, usuario y motivo.

### 👥 4. Directorio de Clientes
- Registro de clientes con DNI/CUIT, teléfono, WhatsApp, email y notas de preferencias.
- Cliente predeterminado *"Consumidor Final"* para ventas rápidas de mostrador.

### 💵 5. Caja Diaria y Arqueo
- Apertura de caja con monto inicial para cambio.
- Registro de ingresos y egresos varios (gastos de flete, pagos a proveedores, retiros).
- Cierre de turno y **Arqueo de Billetes**: compara el monto contado físicamente contra el esperado por el sistema y calcula diferencias (sobrante / faltante).

### 📈 6. Reportes Comerciales y Estadísticas
- Ventas por rango de fechas (diario, semanal, mensual).
- Discriminación de recaudación en efectivo vs medios electrónicos.
- **Ranking de Prendas Más Vendidas** (unidades y facturación).
- Módulo de **Anulación de Ventas**: reintegra el stock al talle vendido y reversa el monto en caja.

### 🌐 7. Migración y Sincronización con el E-commerce Web
- Asistente integrado para conectar con la base de datos de tu tienda web (ej. `grupo5`).
- Diagnóstico automático de prendas, categorías, talles y stock creados en la web.
- **Botón de Migración en 1 Clic**: importa todo el catálogo existente al nuevo software de gestión sin pérdida de datos.
- Script SQL alternativo disponible en `Database/migration_from_ecommerce.sql`.

---

## 📁 Estructura del Proyecto

```
c:\Users\gonza\Desktop\Proyecto\
├── Database/
│   ├── schema_mysql.sql              # Estructura DDL completa para MySQL
│   ├── seed_data.sql                 # Datos maestros iniciales y catálogo demo
│   └── migration_from_ecommerce.sql  # Script de migración desde e-commerce
├── GestionComercial.sln              # Solución estándar Visual Studio
├── GestionComercial.slnx             # Solución formato .NET moderno
└── GestionComercial/
    ├── GestionComercial.vbproj       # Proyecto WinForms VB.NET (.NET 10 / 8)
    ├── Program.vb                    # Punto de entrada de la aplicación
    ├── appsettings.json              # Configuración local de conexión MySQL
    ├── Config/
    │   └── AppConfig.vb              # Manejador de configuración JSON y connection string
    ├── Data/
    │   └── DatabaseHelper.vb         # Conexión ADO.NET, MySqlConnector, SHA256 y scripts
    ├── Models/
    │   └── Entities.vb               # Clases POCO (Producto, Talle, Venta, Cliente, Caja, etc.)
    ├── Services/
    │   ├── AuthService.vb            # Autenticación y roles de usuario
    │   ├── CatalogService.vb         # Lógica de productos, talles y matriz de stock
    │   ├── ClienteService.vb         # Directorio de clientes
    │   ├── CajaService.vb            # Apertura, movimientos y arqueo de caja
    │   ├── VentaService.vb           # Transacciones POS, deducción de stock y tickets
    │   ├── ReporteService.vb         # Métricas de ventas y ranking
    │   ├── ConfiguracionService.vb   # Datos comerciales y del ticket
    │   └── EcommerceSyncService.vb   # Asistente de importación desde la web
    ├── Forms/
    │   ├── FrmMain.vb                # Dashboard principal con panel lateral
    │   ├── FrmLogin.vb               # Pantalla de acceso al sistema
    │   ├── FrmVentasPOS.vb           # Punto de venta ágil para indumentaria
    │   ├── FrmProductos.vb           # Catálogo y editor con matriz de talles
    │   ├── FrmStock.vb               # Reposición, alertas y auditoría
    │   ├── FrmClientes.vb            # ABM de clientes
    │   ├── FrmCaja.vb                # Control de caja y arqueo
    │   ├── FrmReportes.vb            # Estadísticas y anulación de ventas
    │   ├── FrmSincronizacionEcommerce.vb # Asistente de migración web
    │   └── FrmConfiguracion.vb       # Configuración MySQL y datos del local
    └── UI/
        └── UITheme.vb                # Sistema de diseño, paleta Slate/Indigo y estilos
```
