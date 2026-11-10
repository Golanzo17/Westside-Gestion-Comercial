# Westside - Sistema de Gestión Comercial

Sistema de escritorio para la gestión integral de un local de indumentaria: ventas (POS), catálogo con matriz de talles y colores, stock, clientes, caja con arqueo, reportes, usuarios con roles y copias de seguridad. Desarrollado en **Visual Basic .NET (WinForms, .NET 10)** con base de datos embebida **SQLite**.

---

## Tabla de contenidos

1. [Inicio rápido](#inicio-rápido)
2. [Roles y permisos](#roles-y-permisos)
3. [Módulos y funcionalidades](#módulos-y-funcionalidades)
4. [Seguridad y control](#seguridad-y-control)
5. [Arquitectura](#arquitectura)
6. [Estructura del proyecto](#estructura-del-proyecto)
7. [Pruebas](#pruebas)
8. [Documentación adicional](#documentación-adicional)

---

## Inicio rápido

### Requisitos previos
- **.NET SDK 10** (o compatible con el `TargetFramework` del proyecto).
- **Visual Studio 2022 o superior**, o Visual Studio Code con la extensión de C# Dev Kit / .NET.
- Windows (aplicación WinForms).

### Base de datos
- **SQLite embebido, sin configuración**: no requiere servidor ni instalación de servicios. El sistema utiliza el archivo `Database/gestion_comercial.db`.
- Si la base no existe o está vacía, se crean las tablas (`schema_sqlite.sql`) y se cargan los datos iniciales (`seed_sqlite.sql`) al primer arranque.
- `seed_catalogo_demo.sql` carga un catálogo de demostración (productos, talles, stock) y clientes ficticios de Corrientes Capital, útiles para pruebas y presentaciones.
- Desde el módulo **Configuración** se puede verificar y regenerar la estructura de la base cuando sea necesario.

### Ejecución
- **Visual Studio**: abrir `GestionComercial.sln` y presionar **F5**.
- **Terminal**:
  ```powershell
  cd GestionComercial
  dotnet run
  ```

### Credenciales por defecto
| Rol | Usuario | Contraseña |
|---|---|---|
| Administrador | `admin` | `admin123` |
| Gerente | `gerente` | `gerente123` |
| Vendedor | `vendedor` | `1234` |

> Se recomienda cambiar estas contraseñas antes de usar el sistema en un entorno real.

---

## Roles y permisos

La segregación de funciones se aplica en el menú lateral y se vuelve a validar al abrir cada módulo.

| Módulo |Vendedor | Gerente | Administrador |
|---|---|---|---|
| Punto de venta (POS) | Sí | No | No |
| Mis ventas del turno | Sí | No | No |
| Caja diaria y arqueo | Sí (arqueo ciego) | No | No |
| Catálogo y talles | Solo consulta (sin costos ni márgenes) | ABM completo | Sin acceso desde el menú |
| Stock y reposición | No | Sí | No |
| Clientes | Sí | Sí | Sí |
| Cierres y arqueos (auditoría de cajas) | No | Sí | No |
| Reportes y anulación de ventas | No | Sí | No |
| Usuarios y equipo | No | No | Sí |
| Copias de seguridad | No | No | Sí, con autorización de un Gerente |
| Configuración | No | No | Sí |

Operaciones que piden autorización de supervisor:

| Operación | Quién autoriza |
|---|---|
| Descuento superior al 15% en el POS (vendedor) | Administrador o Gerente |
| Egresos o retiros de efectivo de caja (vendedor) | Administrador o Gerente |
| Alta o edición de prendas y precios, ingresos y ajustes de stock (vendedor) | Administrador o Gerente |
| Crear, exportar o restaurar un backup | Gerente (siempre, incluso si la sesión es de Administrador) |
| Anular una venta | Gerente |
| Acceder a Configuración | Administrador |

---

## Módulos y funcionalidades

### 1. Punto de venta (POS)
- Búsqueda por código de barras / SKU o por nombre de la prenda.
- Matriz de talles en vivo: botones dinámicos con el stock disponible por talle y color.
- Carrito con subtotales, descuentos porcentuales y recargos.
- Medios de pago: Efectivo (con calculadora de vuelto), Tarjeta de Débito, Tarjeta de Crédito y Transferencia / QR.
- **Confirmación previa a cobrar**: un cuadro de diálogo resume cliente, cantidad de artículos, medio de pago, total y vuelto antes de registrar la venta.
- Ticket con vista previa e impresión, con datos del comercio y política de cambios.
- Transaccionalidad estricta: cada venta descuenta stock del talle/color vendido, registra el movimiento de stock e impacta en la caja del turno.
- Consulta de **Mis ventas del turno** para el vendedor en sesión.

### 2. Catálogo de productos
- ABM de prendas, categorías y talles.
- Cálculo automático del precio de venta a partir del costo y el margen porcentual.
- Matriz talle x color con stock individual y stock mínimo por variante.

### 3. Control de stock y reposición
- Semáforo de alertas para artículos agotados o por debajo del mínimo.
- Ingreso de mercadería con actualización directa del inventario.
- Auditoría de movimientos (entradas, salidas, ventas, ajustes) con fecha, usuario y motivo.

### 4. Clientes
- Alta, edición y baja con DNI/CUIT validado contra duplicados, teléfono, WhatsApp, email y notas.
- Cliente predeterminado "Consumidor Final" para ventas de mostrador.
- Datos de demostración con direcciones y teléfonos (característica 3794) de Corrientes Capital.

### 5. Caja y arqueo
- Apertura de caja con fondo inicial de cambio.
- Ingresos y egresos manuales (fletes, pagos a proveedores, retiros).
- KPIs del turno con desglose por medio de pago: efectivo, tarjeta y Transferencia / QR.
- Cierre de turno con **arqueo**: compara el efectivo contado contra el esperado y calcula el sobrante o faltante.

### 6. Control de cajas (Gerente)
- Listado de turnos de todos los vendedores.
- Detalle de auditoría de cada caja con la fórmula del efectivo esperado: fondo inicial, ventas en efectivo, ingresos, egresos y total esperado.

### 7. Reportes
- Ventas por rango de fechas y filtros por vendedor, medio de pago y estado.
- Discriminación entre efectivo y medios electrónicos.
- Ranking de prendas más vendidas (unidades y facturación).
- Anulación de ventas: reintegra el stock y revierte el monto en caja, con motivo obligatorio y autorización de Gerente.

### 8. Usuarios y configuración
- ABM de usuarios con asignación de rol y contraseñas almacenadas con hash.
- Datos comerciales del local y del ticket.
- Parámetros de conexión e inicialización de la base de datos.

### 9. Copias de seguridad
- Acceso desde el menú solo para el Administrador.
- Creación de backups de la base SQLite en `Database/Backups/`, exportación a una ruta elegida y restauración desde un archivo de respaldo.
- Las tres operaciones requieren la contraseña de un **Gerente** (doble control).

---

## Seguridad y control

- **Autenticación** con contraseñas hasheadas y control de acceso por rol en el menú y en cada formulario.
- **Doble control**: las acciones sensibles abren un diálogo de autorización que solicita la contraseña de un usuario activo con el rol requerido (por ejemplo Gerente para backups). Cada autorización queda registrada en el log de auditoría con supervisor, operador y motivo.
- **Protección anti fuerza bruta**: bloqueo temporal tras varios intentos fallidos, tanto en el login como en el diálogo de autorización.
- **Registro de eventos** (`AppLogger`) para errores y acciones relevantes.
- **Consultas parametrizadas** en toda la capa de datos para prevenir inyección SQL.
- **Integridad transaccional**: ventas, anulaciones y movimientos de stock se ejecutan dentro de transacciones.

---

## Arquitectura

Aplicación de escritorio por capas:

```
Forms (UI)  ->  Services (lógica de negocio)  ->  Data (DatabaseHelper / SQLite)
                         |
                      Models (entidades)
```

- **Interfaz 100% por código** con un sistema de diseño propio (`UITheme`): paleta Slate/Indigo, tarjetas KPI, grillas estilizadas y diseño responsivo (Dock/Anchor, `SplitContainer`, `TableLayoutPanel`) para resoluciones desde 1024 px hasta pantallas grandes.
- **Servicios** independientes por dominio (autenticación, catálogo, ventas, caja, clientes, reportes, usuarios, configuración, backup).
- **Persistencia local** con SQLite, sin dependencias de red ni servicios en la nube.

---

## Estructura del proyecto

```
Proyecto/
├── Database/
│   ├── gestion_comercial.db          # Base de datos SQLite activa
│   ├── schema_sqlite.sql             # Estructura (DDL)
│   ├── seed_sqlite.sql               # Datos iniciales (usuarios, talles, categorías, clientes)
│   ├── seed_catalogo_demo.sql        # Catálogo y clientes de demostración
│   └── Backups/                      # Copias de seguridad generadas
├── Documentacion_Codigo/             # Documentación técnica y guía de defensa
├── GestionComercial.sln
├── GestionComercial/
│   ├── GestionComercial.vbproj
│   ├── Program.vb                    # Punto de entrada
│   ├── appsettings.json              # Configuración de la aplicación
│   ├── Config/
│   │   └── AppConfig.vb
│   ├── Data/
│   │   └── DatabaseHelper.vb         # Acceso a SQLite y hashing
│   ├── Models/
│   │   └── Entities.vb               # Entidades del dominio
│   ├── Services/
│   │   ├── AuthService.vb            # Autenticación, roles y autorización de gerente
│   │   ├── CatalogService.vb         # Productos, talles y matriz de stock
│   │   ├── ClienteService.vb
│   │   ├── CajaService.vb            # Apertura, movimientos, arqueo y desglose por medio
│   │   ├── VentaService.vb           # Ventas, descuento de stock y anulaciones
│   │   ├── ReporteService.vb
│   │   ├── UsuarioService.vb
│   │   ├── ConfiguracionService.vb
│   │   ├── BackupService.vb          # Respaldo y restauración
│   │   └── AppLogger.vb              # Registro de eventos
│   ├── Forms/
│   │   ├── FrmLogin.vb
│   │   ├── FrmMain.vb                # Panel principal con navegación lateral
│   │   ├── FrmVentasPOS.vb
│   │   ├── FrmMisVentasTurno.vb
│   │   ├── FrmProductos.vb
│   │   ├── FrmStock.vb
│   │   ├── FrmClientes.vb
│   │   ├── FrmCaja.vb
│   │   ├── FrmControlCajas.vb
│   │   ├── FrmDetalleArqueoCaja.vb
│   │   ├── FrmReportes.vb
│   │   ├── FrmUsuarios.vb
│   │   ├── FrmConfiguracion.vb
│   │   ├── FrmBackup.vb
│   │   └── FrmAutorizacionSupervisor.vb
│   └── UI/
│       └── UITheme.vb                # Sistema de diseño y estilos
└── GestionComercial.Tests/
    ├── GestionComercial.Tests.vbproj
    └── SystemIntegrationTests.vb     # Pruebas de integración
```

---

## Pruebas

```powershell
dotnet build GestionComercial.sln
dotnet test GestionComercial.Tests/GestionComercial.Tests.vbproj
```

---

## Documentación adicional

- `Documentacion_Codigo/`: explicación técnica por capas, flujo completo del programa y guía para la defensa oral.
