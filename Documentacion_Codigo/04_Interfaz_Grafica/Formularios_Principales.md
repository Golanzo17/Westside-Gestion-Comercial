# 🖥️ Explicación de los Formularios y Ventanas (Forms)

* **Ubicación**: [GestionComercial/Forms/]
* **Propósito**: Contiene todas las pantallas gráficas con las que interactúa el usuario final. Cada formulario habla con uno o más servicios (`Services`), nunca directamente con la base de datos.

---

## 💡 Conceptos Clave de Windows Forms en VB.NET

### 1. `Inherits Form`
Indica que nuestra clase es una ventana de Windows. Hereda automáticamente todas las capacidades visuales del sistema operativo (minimizar, maximizar, cerrar, redimensionar).

### 2. Controles Dinámicos por Código vs Diseñador Visual
En este proyecto, los formularios se crean **100% por código limpio** en un método llamado `InitializeUI()`.
Esto evita que se corrompan los archivos `.Designer.vb` que a menudo causan dolores de cabeza en Visual Studio.

### 3. Eventos (`AddHandler ... AddressOf ...`)
Para que un botón haga algo cuando el usuario le hace clic:
```vb
AddHandler btnCobrar.Click, AddressOf BtnCobrar_Click
```
* `btnCobrar.Click`: El evento que dispara Windows cuando se presiona el botón.
* `AddressOf BtnCobrar_Click`: La dirección de la subrutina (`Sub`) que contiene el código a ejecutar.

### 4. Diseño Responsivo
Los formularios usan `Dock`, `Anchor`, `TableLayoutPanel`, `FlowLayoutPanel` y `SplitContainer` en lugar de posiciones fijas, para verse bien desde 1024 px hasta pantallas grandes. Todos los estilos salen de `UITheme` (ver `UITheme.md`).

---

## 🔐 Segregación de Funciones por Rol

El sistema aplica el **principio de mínimo privilegio**: cada rol ve solo lo que necesita. `FrmMain` oculta los botones del menú según el rol y, además, cada acción de navegación (`Nav_...`) vuelve a validar el rol antes de abrir la pantalla.

* **Vendedor** (atención al público): opera el POS, su caja y sus ventas del turno, consulta el catálogo y gestiona clientes.
* **Gerente** (supervisión): administra catálogo y stock, audita las cajas de todos los vendedores, consulta reportes y autoriza operaciones sensibles. **No tiene caja ni POS propios.**
* **Administrador** (TI / propietario): gestiona usuarios, configuración y copias de seguridad. No opera dinero ni ventas.

| Módulo / Función | Administrador | Gerente | Vendedor |
| :--- | :---: | :---: | :---: |
| **Panel Principal (KPIs)** | Visible | Visible | Visible |
| **Punto de Venta (POS)** | No | No | Operativo |
| **Mis ventas del turno** | No | No | Operativo |
| **Caja diaria y arqueo** | No | No | Operativo (arqueo ciego) |
| **Catálogo y Talles** | No (sin acceso desde el menú) | Gestión total | Solo consulta, sin costos ni márgenes |
| **Stock y Reposición** | No | Operativo | No |
| **Clientes** | Visible | Visible | Visible |
| **Cierres y arqueos (auditoría)** | No | Operativo | No |
| **Reportes y Ventas** | No | Completo | No |
| **Usuarios y Equipo** | Exclusivo | No | No |
| **Copias de Seguridad** | Exclusivo, con autorización de un Gerente | No | No |
| **Configuración General** | Exclusivo | No | No |

**Operaciones con autorización de supervisor** (diálogo `FrmAutorizacionSupervisor`):
* Descuentos mayores al 15% y egresos de caja de un vendedor: Administrador o Gerente.
* Alta o modificación de prendas, ingreso de mercadería y ajustes de stock hechos por un vendedor: Administrador o Gerente.
* Anulación de ventas: Gerente.
* Crear, exportar o restaurar un backup: Gerente (siempre).

---

## 📑 Resumen de los Formularios

### 1. [FrmLogin.vb] (Inicio de Sesión)
* **Controles**: cajas de texto para Usuario y Contraseña (con `PasswordChar = "*"c`), botón "Ingresar".
* **Atajo de teclado**: captura la tecla `Enter` en el campo contraseña para que no sea necesario tocar el ratón.
* **Indicador de conexión**: al cargar, comprueba que la base responda y muestra un indicador verde o rojo (`CheckInitialConnection`).
* **Lógica**: llama a `AuthService.Login(...)`. Si es correcto, oculta el login (`Me.Hide()`) y abre `FrmMain`. Vincula el evento `FormClosed` de `FrmMain` para que, al salir, se cierre el proceso por completo.

---

### 2. [FrmMain.vb] (Menú Principal y Navegación)
* **Barra lateral izquierda (`pnlSidebar`)**: botones agrupados en secciones (Operaciones, Gestión, Sistema), visibles según el rol.
* **Barra superior (`pnlTopBar`)**: sesión activa con distintivo del rol, **estado de caja en tiempo real** (solo para el Vendedor), botón de venta rápida (F1, solo Vendedor) y reloj en vivo.
* **Panel de contenido central (`pnlContentHost`)**: donde se cargan los formularios hijos, sin ventanas flotantes.
* **Panel principal (dashboard)**: tarjetas KPI y accesos rápidos que cambian con el rol:
  * Vendedor: "Mis ventas del turno" y "Mis tickets del turno", con la recaudación de su propia caja.
  * Gerente: ventas y tickets globales del día, stock crítico y **"Cajas en curso"** (cuántos turnos de vendedores siguen abiertos).
  * Administrador: ventas globales, stock crítico y cantidad de usuarios registrados.
* **Salir**: el cierre de sesión vuelve al login y el cierre de la ventana termina el proceso con `Application.Exit()`.

---

### 3. [FrmVentasPOS.vb] (Punto de Venta)
Es la pantalla más utilizada en el mostrador del local:
* **Lector de Código de Barras**: caja de texto con foco permanente (`txtBusqueda`). Al escanear o escribir y dar Enter, busca con `CatalogService.GetProductoPorCodigo`.
* **Matriz de Talles y Colores**: botones dinámicos con el stock disponible de cada talle y color (`GetStockPorTalles`).
* **Carrito (`dgvCarrito`)**: artículos, cantidades, precios unitarios y subtotales.
* **Descuentos, recargos y vuelto**: recalcula el total en tiempo real. Un descuento **superior al 15%** exige la autorización de un Administrador o Gerente.
* **Medios de pago**: Efectivo, Tarjeta Débito, Tarjeta Crédito, Transferencia / QR y Múltiple.
* **Confirmación previa al cobro**: un cuadro de diálogo (`MessageBoxButtons.YesNo`) resume cliente, cantidad de artículos, medio de pago, total y vuelto. La venta se registra solo si el cajero confirma.
* **Cobro**: llama a `VentaService.ProcesarVenta(...)`. Si es exitoso, emite el ticket, descuenta el stock y vacía el carrito.

---

### 4. [FrmProductos.vb] y [FrmStock.vb]
* **`FrmProductos`** (Gerente y Vendedor):
  * Listado con filtros de búsqueda rápida y matriz de talles.
  * El Gerente crea y edita prendas, define costo, margen y precio, y activa o desactiva prendas (`ToggleActivo`).
  * El **Vendedor solo consulta**: las columnas de costo y margen están ocultas y, si intentara modificar algo, se pide autorización.
* **`FrmStock`** (Gerente):
  * Inventario con **semáforo**: resalta en rojo/naranja las prendas con stock crítico o agotado.
  * "Ingreso de mercadería" y ajustes manuales con motivo obligatorio, auditados en `movimientos_stock`.

---

### 5. [FrmCaja.vb] (Caja del Vendedor)
* **Tarjetas KPI**: estado del turno y totales cobrados en **Efectivo**, **Tarjetas** y **Transferencia / QR**, calculados con `GetDesgloseMediosTurno`.
* **Abrir Caja**: solicita el fondo inicial de cambio.
* **Ingreso / retiro de efectivo**: un egreso pide la autorización de un Administrador o Gerente y queda registrado con quién lo aprobó.
* **Cerrar Caja (Arqueo ciego)**: el vendedor cuenta el efectivo **sin ver el monto esperado**; el sistema calcula el sobrante o faltante y guarda el cierre.

---

### 6. [FrmMisVentasTurno.vb] (Mis Ventas del Turno)
Exclusiva del **Vendedor**. Muestra únicamente los tickets que él emitió en su caja abierta (`GetVentasPorTurnoVendedor`), sin costos, márgenes ni ventas de otros vendedores.
* **Ver detalle** del ticket.
* **Reimprimir en PDF** el comprobante.
* **Cambio de prenda** por otro talle o color: abre un modal y llama a `VentaService.ProcesarCambioPrenda`, que reingresa la prenda devuelta, descuenta la nueva y deja el rastro en el stock.

---

### 7. [FrmClientes.vb] (Directorio de Clientes)
* ABM de clientes con DNI/CUIT, teléfono, email, dirección, ciudad, notas (talles preferidos) y fecha de nacimiento.
* Validación de DNI duplicado antes de guardar.
* **Borrado lógico**: el botón activar/desactivar usa `ClienteService.ToggleActivo`. El cliente `ID = 1` ("Consumidor Final") no se puede desactivar.
* Filtro para mostrar activos, inactivos o todos; las filas inactivas se distinguen por color (`CellFormatting`).

---

### 8. [FrmControlCajas.vb] (Cierres y Arqueos, Gerente)
Tablero de auditoría de los turnos de **todos** los vendedores:
* **Filtros**: rango de fechas (con atajos rápidos), vendedor, estado (abierta/cerrada) y discrepancia (faltante, sobrante, exacto).
* **Tarjetas KPI**: cantidad de turnos, facturación, esperado vs real y balance neto de los arqueos.
* **Grilla**: marca en rojo los faltantes, en azul los sobrantes y en verde los arqueos exactos.
* **Acciones**: auditar el turno (abre `FrmDetalleArqueoCaja`), ver el acta, ver los movimientos manuales y ver los tickets. Doble clic sobre un turno también abre el detalle.
* **Exportación** de la auditoría a CSV para conciliación externa.
* Usa `CajaService.GetHistorialCajasAvanzado`.

---

### 9. [FrmDetalleArqueoCaja.vb] (Detalle de un Arqueo)
Ventana de auditoría de un turno, organizada en pestañas:
* **Conciliación**: una "tarjeta de fórmula" muestra el efectivo esperado paso a paso (fondo inicial, ventas en efectivo, ingresos, egresos, total esperado) contra lo contado y la diferencia, más el desglose de cobros por medio de pago.
* **Ventas**: todos los tickets del turno (`GetVentasPorCaja`).
* **Movimientos**: ingresos y egresos manuales con el supervisor que los autorizó.
* **Acta**: genera el **acta oficial de arqueo y cierre** para firma del vendedor y del Gerente, lista para imprimir.

---

### 10. [FrmReportes.vb] (Reportes, Gerente)
Inteligencia comercial en pestañas:
* **Ventas**: buscador de tickets por fechas, vendedor, cliente, medio de pago y estado, con KPIs financieros (`GetMetricasFinancierasPeriodo`) y descarga de ticket en PDF.
* **Anulaciones**: historial de ventas anuladas.
* **Clientes**: ventas agrupadas por cliente, con tickets de cada uno.
* **Vendedores**: desempeño por vendedor y sus tickets.
* **Ranking**: prendas más vendidas, con modo **Mayor Salida** o **Mayor Margen**, y reporte de **stock inmovilizado**.
* **Anular venta**: exige la autorización de un Gerente y un motivo, y revierte el stock y la caja de forma atómica (`VentaService.AnularVenta`).

---

### 11. [FrmUsuarios.vb] (Usuarios y Equipo, Administrador)
* Alta, modificación y desactivación de empleados con su rol.
* Campos: Nombre, Apellido, DNI (con validación de duplicados), Teléfono, Email, Ciudad, Dirección y Notas.
* Contraseñas con hashing PBKDF2. El sistema impide desactivar al único Administrador activo.

---

### 12. [FrmBackup.vb] (Copias de Seguridad, Administrador)
* **Tarjetas KPI**: base de datos activa (tablas y tickets), cantidad de copias, último respaldo e integridad.
* **Grilla** con el historial de copias (fecha y tamaño).
* **Crear copia**, **exportar a una ruta** elegida, **restaurar** una copia y abrir la carpeta de respaldos (`BackupService`).
* **Doble control**: crear, exportar y restaurar piden la contraseña de un **Gerente** (`SolicitarAutorizacionGerente`), aunque quien opere sea Administrador. La restauración además pide una confirmación estricta, porque reemplaza toda la base.

---

### 13. [FrmConfiguracion.vb] (Configuración, Administrador)
* **Base de datos**: ruta del archivo SQLite, botón para **probar la conexión** e **inicializar tablas y datos** (esquema + semillas).
* **Datos del comercio**: nombre, CUIT, condición de IVA, dirección, teléfono, email y mensaje al pie del ticket, que se guardan con `ConfiguracionService`.
* Atajos para ir a Usuarios y a Copias de Seguridad.
* Acceder desde el menú pide la autorización de un Administrador (`SolicitarAutorizacionAdmin`).

---

### 14. [FrmAutorizacionSupervisor.vb] (Diálogo de Autorización)
* Ventana modal reutilizable que usan todas las acciones con doble control.
* Recibe un título, el rol requerido, el motivo y los roles permitidos, y pide **solo la contraseña** del supervisor.
* Valida con `AuthService.AutenticarSupervisorPorRoles`, que aplica el bloqueo anti fuerza bruta y registra en auditoría quién autorizó, a quién y por qué.
* Si se confirma, devuelve `DialogResult.OK` y el supervisor queda en `AuthService.LastAuthorizingSupervisor`.
