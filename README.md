# Software de Gestión Comercial para Local de Ropa

Sistema integral de gestión de ventas, stock, caja y clientes desarrollado en **Visual Basic (.NET WinForms)** con motor de base de datos embebido **SQLite**, diseñado específicamente para locales de indumentaria.

---

## 🚀 Inicio Rápido

### 1. Requisitos Previos
- **.NET SDK** (.NET 8 LTS o .NET 10).
- **Visual Studio 2022/2026** o **Visual Studio Code**.

### 2. Base de Datos
- **SQLite Embebido (Zero-Config)**: No requiere ningún servidor ni instalación externa de servicios. El sistema gestiona automáticamente el archivo `Database/gestion_comercial.db`, inicializando tablas (`schema_sqlite.sql`) y datos semilla (`seed_sqlite.sql`) de forma transparente al iniciar por primera vez.
- **Inicialización Asistida**: El propio software cuenta con una herramienta en el menú de **Configuración** (`⚙ Configuración -> Inicializar tablas y datos`) para verificar y regenerar la estructura cuando sea necesario.

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
| **Gerente** | `gerente` | `gerente123` |
| **Vendedor** | `vendedor` | `1234` |

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
- Registro de clientes con DNI/CUIT (con validación de duplicados en tiempo real), teléfono, WhatsApp, email y notas de preferencias.
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

---

## 📁 Estructura del Proyecto

```
c:\Users\gonza\Desktop\Proyecto\
├── Database/
│   ├── gestion_comercial.db          # Base de datos SQLite local activa
│   ├── schema_sqlite.sql             # Estructura DDL completa para SQLite
│   └── seed_sqlite.sql               # Datos iniciales para SQLite (usuarios, talles, categorías)
├── GestionComercial.sln              # Solución estándar Visual Studio
├── GestionComercial.slnx             # Solución formato .NET moderno
└── GestionComercial/
    ├── GestionComercial.vbproj       # Proyecto WinForms VB.NET (.NET 10 / 8)
    ├── Program.vb                    # Punto de entrada de la aplicación
    ├── appsettings.json              # Configuración de base de datos
    ├── Config/
    │   └── AppConfig.vb              # Manejador de configuración JSON y connection string
    ├── Data/
    │   └── DatabaseHelper.vb         # Capa de persistencia SQLite y seguridad criptográfica
    ├── Models/
    │   └── Entities.vb               # Clases POCO (Producto, Talle, Venta, Cliente, Caja, etc.)
    ├── Services/
    │   ├── AuthService.vb            # Autenticación y roles de usuario
    │   ├── CatalogService.vb         # Lógica de productos, talles y matriz de stock
    │   ├── ClienteService.vb         # Directorio y validación de clientes
    │   ├── CajaService.vb            # Apertura, movimientos y arqueo de caja
    │   ├── VentaService.vb           # Transacciones POS, deducción de stock y tickets
    │   ├── ReporteService.vb         # Métricas de ventas y ranking
    │   ├── UsuarioService.vb         # Gestión y auditoría de usuarios y DNI
    │   └── ConfiguracionService.vb   # Datos comerciales y del ticket
    ├── Forms/
    │   ├── FrmMain.vb                # Dashboard principal con panel lateral
    │   ├── FrmLogin.vb               # Pantalla de acceso al sistema
    │   ├── FrmVentasPOS.vb           # Punto de venta ágil para indumentaria
    │   ├── FrmProductos.vb           # Catálogo y editor con matriz de talles
    │   ├── FrmStock.vb               # Reposición, alertas y auditoría
    │   ├── FrmClientes.vb            # ABM de clientes
    │   ├── FrmCaja.vb                # Control de caja y arqueo
    │   ├── FrmUsuarios.vb            # Gestión de usuarios del sistema
    │   ├── FrmReportes.vb            # Estadísticas y anulación de ventas
    │   └── FrmConfiguracion.vb       # Configuración de conexión y datos del local
    └── UI/
        └── UITheme.vb                # Sistema de diseño, paleta Slate/Indigo y estilos
```
