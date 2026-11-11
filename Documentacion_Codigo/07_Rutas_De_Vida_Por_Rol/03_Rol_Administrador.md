# 🛡️ Ruta de Vida por Formulario — Rol: Administrador

Este documento detalla **todas las pantallas y funciones disponibles para el rol `Administrador`**, mostrando la **ruta de vida completa** de cada botón o acción: dónde se dispara en la interfaz (`Forms/`), qué servicio lo procesa (`Services/`), qué consultas/acciones ejecuta en SQLite (`Data/DatabaseHelper.vb`) y qué actualiza en pantalla.

---

## 🔐 Resumen de Permisos del Rol (`AuthService.vb`)

En [AuthService.vb](../../GestionComercial/Services/AuthService.vb):
* `AuthService.IsAdmin` → **`True`**
* `AuthService.IsAdminOrManager` → **`True`**
* `AuthService.IsManager` → **`False`**
* `AuthService.IsVendor` → **`False`**

Por aplicación estricta del **Principio de Menor Privilegio** en [FrmMain.vb](../../GestionComercial/Forms/FrmMain.vb#L136-L147):
* **Ve y opera:** Panel principal administrativo (`FrmMain`), Usuarios y Equipo (`FrmUsuarios` + `FrmUsuarioEditor` + `FrmCambiarPassword`), Copias de Seguridad (`FrmBackup`), Configuración del Sistema (`FrmConfiguracion`) y Clientes (`FrmClientes`).
* **Tiene oculto / bloqueado:** Terminal POS (`FrmVentasPOS`), Mis Ventas del Turno (`FrmMisVentasTurno`), Caja (`FrmCaja`), Catálogo (`FrmProductos` bloqueado en `Nav_Productos` por `If AuthService.IsAdmin Then Return`), Reposición de Stock (`FrmStock`), Auditoría de Cierres (`FrmControlCajas`) y Reportes (`FrmReportes`).
* **Política de Control Dual (Cuatro Ojos) en Backups:** Aunque el Administrador opera `FrmBackup`, generar o restaurar respaldos exige adicionalmente la contraseña presencial de un **Gerente** (`AuthService.SolicitarAutorizacionGerente`).

---

## 📋 Mapa de Formularios del Administrador

1. [FrmLogin.vb](#1-frmloginvb--inicio-de-sesión) — Ingreso al sistema
2. [FrmMain.vb](#2-frmmainvb--dashboard-administrativo-y-navegación) — Indicadores generales y acceso a configuración
3. [FrmUsuarios.vb](#3-frmusuariosvb--gestión-de-usuarios-empleados-y-claves) — Alta, edición, cambio de contraseña y baja lógica
4. [FrmBackup.vb](#4-frmbackupvb--copias-de-seguridad-y-restauración) — Respaldos SQLite y control dual con Gerente
5. [FrmConfiguracion.vb](#5-frmconfiguracionvb--configuración-de-bd-y-comercio) — Prueba/inicialización de BD y datos fiscales del ticket
6. [FrmClientes.vb](#6-frmclientesvb--directorio-y-abm-de-clientes) — Consulta, alta, edición y baja lógica
7. [FrmAutorizacionSupervisor.vb](#7-frmautorizacionsupervisorvb--confirmación-y-control-dual) — Validación de clave de Admin o Gerente

---

## 1. `FrmLogin.vb` — Inicio de Sesión

Archivo: [GestionComercial/Forms/FrmLogin.vb](../../GestionComercial/Forms/FrmLogin.vb)

* **Al iniciar:** `New()` → [`CheckInitialConnection()`](../../GestionComercial/Forms/FrmLogin.vb#L158) → [`DatabaseHelper.TestConnection()`](../../GestionComercial/Data/DatabaseHelper.vb#L59) valida el archivo SQLite `gestion_comercial.db` y asegura tablas/migraciones.
* **Botón `INGRESAR AL SISTEMA`:** [`BtnLogin_Click`](../../GestionComercial/Forms/FrmLogin.vb#L171) → [`AuthService.Login()`](../../GestionComercial/Services/AuthService.vb#L104) verifica intentos fallidos, busca usuario activo en `usuarios`, compara el hash con [`DatabaseHelper.VerifyPassword`](../../GestionComercial/Data/DatabaseHelper.vb#L207), actualiza `ultimo_login`, registra `AppLogger.LogAudit` y abre `FrmMain`.

---

## 2. `FrmMain.vb` — Dashboard Administrativo y Navegación

Archivo: [GestionComercial/Forms/FrmMain.vb](../../GestionComercial/Forms/FrmMain.vb)

### 🔹 Carga de KPIs Administrativos (`LoadDashboardData`)
* **Código:** [`BuildDashboardHome()`](../../GestionComercial/Forms/FrmMain.vb#L339) + [`LoadDashboardData()`](../../GestionComercial/Forms/FrmMain.vb#L631-L684).
* **Qué ejecuta:**
  * [`ReporteService.GetResumenHoy()`](../../GestionComercial/Services/ReporteService.vb#L122): puebla **VENTAS DE HOY**, **TICKETS EMITIDOS** y **PRENDAS STOCK CRÍTICO**.
  * [`UsuarioService.GetUsuarios()`](../../GestionComercial/Services/UsuarioService.vb#L25): cuenta las cuentas totales y las muestra en la tarjeta 4 (**USUARIOS REGISTRADOS**).
  * Oculta `lblCajaStatus` y oculta el botón `+ NUEVA VENTA (F1)` en la barra superior.

### 🔹 Botones de Navegación del Administrador
* `Panel principal` → [`Nav_Home`](../../GestionComercial/Forms/FrmMain.vb#L707)
* `Usuarios y equipo` → [`Nav_Usuarios`](../../GestionComercial/Forms/FrmMain.vb#L780) → verifica `AuthService.IsAdmin` e incrusta `FrmUsuarios` vía `OpenChildForm()`.
* `Copias de seguridad` → [`Nav_Backup`](../../GestionComercial/Forms/FrmMain.vb#L790) → verifica `AuthService.IsAdmin` e incrusta `FrmBackup`.
* `Configuración` → [`Nav_Config`](../../GestionComercial/Forms/FrmMain.vb#L801) → llama primero a [`AuthService.SolicitarAutorizacionAdmin()`](../../GestionComercial/Services/AuthService.vb#L334) y abre `FrmConfiguracion` como modal; al cerrarse con `OK`, refresca `LoadDashboardData()`.
* `Clientes` → [`Nav_Clientes`](../../GestionComercial/Forms/FrmMain.vb#L740) → incrusta `FrmClientes`.
* `Cerrar sesión` → [`BtnLogout_Click`](../../GestionComercial/Forms/FrmMain.vb#L812) → [`AuthService.Logout()`](../../GestionComercial/Services/AuthService.vb#L240) y regresa a `FrmLogin`.

---

## 3. `FrmUsuarios.vb` — Gestión de Usuarios, Empleados y Claves

Archivo: [GestionComercial/Forms/FrmUsuarios.vb](../../GestionComercial/Forms/FrmUsuarios.vb)

### 🔹 Control de entrada y carga de nómina
* [`OnLoad()`](../../GestionComercial/Forms/FrmUsuarios.vb#L41): doble barrera de seguridad que cierra automáticamente la ventana si `AuthService.IsAdmin` es `False`.
* [`LoadUsuarios()`](../../GestionComercial/Forms/FrmUsuarios.vb#L230): llama a [`UsuarioService.GetUsuarios()`](../../GestionComercial/Services/UsuarioService.vb#L25) (`SELECT * FROM usuarios ORDER BY apellido, nombre`), llena `dgvUsuarios` y actualiza el contador de activos/inactivos en `lblTotalUsuarios`.
* [`DgvUsuarios_CellFormatting()`](../../GestionComercial/Forms/FrmUsuarios.vb#L252): colorea estado `Activo`/`Inactivo` y resalta roles `Administrador` y `Gerente`.

### 🔹 Botón `+ Nuevo Usuario` → Modal `FrmUsuarioEditor(0)`
* **Apertura:** [`BtnNuevo_Click`](../../GestionComercial/Forms/FrmUsuarios.vb#L284) abre `FrmUsuarioEditor(0)`.
* **Botón `💾 Guardar Empleado` ([`BtnGuardar_Click`](../../GestionComercial/Forms/FrmUsuarios.vb#L514-L571)):**
  1. Valida que no falte ninguno de los campos obligatorios (Nombre, Apellido, DNI, Usuario, Teléfono, Email, Dirección, Ciudad), rol seleccionado y contraseña inicial mínima de 4 caracteres.
  2. Llama a [`UsuarioService.CrearUsuario(...)`](../../GestionComercial/Services/UsuarioService.vb#L83):
     * Valida rol contra `UsuarioService.RolesDisponibles` (`Vendedor`, `Gerente`, `Administrador`).
     * Comprueba en `usuarios` que el `username` no esté repetido y que el `dni` sea único (`ExisteDni()`).
     * Encripta la clave llamando a [`DatabaseHelper.HashPasswordSecure(password)`](../../GestionComercial/Data/DatabaseHelper.vb#L194) (PBKDF2-SHA256, 100.000 iteraciones + Salt aleatorio).
     * Ejecuta `INSERT INTO usuarios (...)` y registra `AppLogger.LogAudit`.
  3. Cierra el modal con `OK` y recarga `LoadUsuarios()`.

### 🔹 Botón `✏ Editar Datos` (o doble clic) → Modal `FrmUsuarioEditor(id)`
* **Apertura:** [`BtnEditar_Click`](../../GestionComercial/Forms/FrmUsuarios.vb#L292) / [`DgvUsuarios_CellDoubleClick`](../../GestionComercial/Forms/FrmUsuarios.vb#L277) → [`EditarSeleccionado()`](../../GestionComercial/Forms/FrmUsuarios.vb#L297).
* **Carga de datos:** [`LoadUserData()`](../../GestionComercial/Forms/FrmUsuarios.vb#L484) llama a [`UsuarioService.GetUsuarioById(_usuarioId)`](../../GestionComercial/Services/UsuarioService.vb#L52), rellena los campos y deja `txtUsername` en solo lectura.
* **Botón `💾 Guardar Empleado` ([`BtnGuardar_Click`](../../GestionComercial/Forms/FrmUsuarios.vb#L514)):**
  * Llama a [`UsuarioService.ActualizarUsuario(...)`](../../GestionComercial/Services/UsuarioService.vb#L195) que valida DNI único (`ExisteDni(dni, usuarioId)`), rol válido y ejecuta `UPDATE usuarios SET ... WHERE id = @id` + `AppLogger.LogAudit`.

### 🔹 Botón `🔑 Cambiar Clave` → Modal `FrmCambiarPassword`
* **Apertura:** [`BtnCambiarPass_Click`](../../GestionComercial/Forms/FrmUsuarios.vb#L310) abre `FrmCambiarPassword(uId, username)`.
* **Botón `Actualizar Clave` ([`BtnGuardar_Click`](../../GestionComercial/Forms/FrmUsuarios.vb#L638)):**
  * Verifica que `txtNueva.Text = txtConfirmar.Text`.
  * Llama a [`UsuarioService.CambiarPassword(uId, nuevaPassword, errMsg)`](../../GestionComercial/Services/UsuarioService.vb#L295) → valida longitud mínima de 4 caracteres, genera nuevo hash con `DatabaseHelper.HashPasswordSecure()`, ejecuta `UPDATE usuarios SET password_hash = @hash` y audita en `AppLogger`.

### 🔹 Botón `🚫 Activar / Desactivar` (Borrado lógico)
* **Código:** [`BtnToggleActivo_Click`](../../GestionComercial/Forms/FrmUsuarios.vb#L327-L348).
* **Servicio:** Llama a [`UsuarioService.ToggleActivo(uId, errMsg)`](../../GestionComercial/Services/UsuarioService.vb#L323), que **protege al usuario `admin` principal** impidiendo que el sistema quede sin administradores, alterna `activo` con `UPDATE usuarios SET activo = @activo` y registra `AppLogger.LogAudit`.

---

## 4. `FrmBackup.vb` — Copias de Seguridad y Restauración

Archivo: [GestionComercial/Forms/FrmBackup.vb](../../GestionComercial/Forms/FrmBackup.vb)

### 🔹 Carga del estado de la BD y lista de respaldos
* [`OnLoad()`](../../GestionComercial/Forms/FrmBackup.vb#L43) verifica `AuthService.IsAdmin`.
* [`LoadDashboardBackup()`](../../GestionComercial/Forms/FrmBackup.vb#L253):
  * Llama a [`BackupService.ObtenerInfoBaseDatosActual()`](../../GestionComercial/Services/BackupService.vb#L191) para mostrar ruta y tamaño del archivo `gestion_comercial.db`.
  * Llama a [`BackupService.GetDirectorioBackups()`](../../GestionComercial/Services/BackupService.vb#L65) y [`BackupService.GetBackups()`](../../GestionComercial/Services/BackupService.vb#L77) para listar los archivos `.db` de `Database/Backups/`.

### 🔹 Botón `➕ Crear Backup Ahora`
* **Código:** [`BtnCrearBackup_Click`](../../GestionComercial/Forms/FrmBackup.vb#L279-L294).
* **Control Dual:** Llama primero a [`AuthService.SolicitarAutorizacionGerente()`](../../GestionComercial/Services/AuthService.vb#L309) exigiendo la clave presencial de un Gerente en `FrmAutorizacionSupervisor`.
* **Servicio:** Llama a [`BackupService.CrearBackup(errMsg)`](../../GestionComercial/Services/BackupService.vb#L103) → `CrearBackup(rutaDestino, errMsg)` (L111), que usa `SqliteConnection.BackupDatabase()` para copiar de forma segura y consistente la base activa sin bloquearla, y registra `AppLogger.LogAudit("BACKUP_CREADO", ...)`.

### 🔹 Botón `💾 Guardar Copia en Otra Carpeta...`
* **Código:** [`BtnCrearEnRuta_Click`](../../GestionComercial/Forms/FrmBackup.vb#L297-L319).
* **Flujo:** Solicita clave de Gerente (`AuthService.SolicitarAutorizacionGerente`), abre `SaveFileDialog` para elegir destino (ej. pendrive o disco externo) y ejecuta [`BackupService.CrearBackup(sfd.FileName, errMsg)`](../../GestionComercial/Services/BackupService.vb#L111).

### 🔹 Botón `♻️ Restaurar Backup Seleccionado` (o doble clic)
* **Código:** [`BtnRestaurar_Click`](../../GestionComercial/Forms/FrmBackup.vb#L322-L352).
* **Flujo de máxima seguridad:**
  1. Muestra advertencia crítica pidiendo confirmación `Sí/No`.
  2. Exige contraseña presencial de un Gerente (`AuthService.SolicitarAutorizacionGerente`).
  3. Llama a [`BackupService.RestaurarBackup(rutaArchivo, errMsg)`](../../GestionComercial/Services/BackupService.vb#L158):
     * Crea automáticamente un **respaldo preventivo de emergencia** de la base actual antes de sobrescribir.
     * Vacía el pool de conexiones (`SqliteConnection.ClearAllPools()`) para liberar locks del archivo.
     * Copia el backup seleccionado sobre `gestion_comercial.db`.
     * Llama a `DatabaseHelper.TestConnection()` para verificar integridad y migrar columnas si hiciera falta.
     * Registra `AppLogger.LogAudit("BACKUP_RESTAURADO", ...)`.

### 🔹 Botón `📂 Abrir Carpeta de Backups`
* **Código:** [`BtnAbrirCarpeta_Click`](../../GestionComercial/Forms/FrmBackup.vb#L355) obtiene la ruta con `BackupService.GetDirectorioBackups()`, la crea si no existe y abre `explorer.exe`.

---

## 5. `FrmConfiguracion.vb` — Configuración de BD y Comercio

Archivo: [GestionComercial/Forms/FrmConfiguracion.vb](../../GestionComercial/Forms/FrmConfiguracion.vb)

### 🔹 Carga inicial de rutas y datos institucionales
* **Código:** [`LoadConfigData()`](../../GestionComercial/Forms/FrmConfiguracion.vb#L220-L249) lee la cadena SQLite con `AppConfig.BuildConnectionString()` y consulta [`ConfiguracionService.GetConfiguracion()`](../../GestionComercial/Services/ConfiguracionService.vb#L18) (`SELECT * FROM configuracion_comercio LIMIT 1`).

### 🔹 Botón `Probar Conexión`
* **Código:** [`BtnProbar_Click`](../../GestionComercial/Forms/FrmConfiguracion.vb#L253) → [`DatabaseHelper.TestConnection(errMsg)`](../../GestionComercial/Data/DatabaseHelper.vb#L59) → informa si SQLite responde correctamente.

### 🔹 Botón `Inicializar BD y Tablas`
* **Código:** [`BtnInicializarDb_Click`](../../GestionComercial/Forms/FrmConfiguracion.vb#L272) → [`DatabaseHelper.InitializeDatabaseAndTables(outMsg)`](../../GestionComercial/Data/DatabaseHelper.vb#L268):
  * Resuelve las rutas de `Database/schema_sqlite.sql` y `Database/seed_sqlite.sql` (`ResolveDatabaseScriptPath`).
  * Crea tablas si no existen, siembra datos base (`seed_sqlite.sql`), ejecuta `ApplyDemoCatalogSeed` y `EnsureSchemaUpdates()`.

### 🔹 Botones de acceso directo: `Copias de Seguridad (Backup)...` y `Usuarios y empleados`
* `btnIrBackup.Click` (L118) abre `FrmBackup` en diálogo modal.
* `btnIrUsuarios.Click` (L192) verifica `AuthService.IsAdmin` y abre `FrmUsuarios` en diálogo modal.

### 🔹 Botón `Guardar configuración`
* **Código:** [`BtnGuardarTodo_Click`](../../GestionComercial/Forms/FrmConfiguracion.vb#L288-L309) toma los campos de comercio (Nombre, CUIT, Condición IVA, Dirección, Teléfono, Email, Mensaje del ticket) y llama a [`ConfiguracionService.GuardarConfiguracion(cfg, errMsg)`](../../GestionComercial/Services/ConfiguracionService.vb#L43) que ejecuta `UPDATE` (o `INSERT` si no existía fila) en `configuracion_comercio`.

---

## 6. `FrmClientes.vb` — Directorio y ABM de Clientes

Archivo: [GestionComercial/Forms/FrmClientes.vb](../../GestionComercial/Forms/FrmClientes.vb)

* Visible en el panel lateral para todos los roles (`FrmMain.vb` L174).
* [`LoadClientes()`](../../GestionComercial/Forms/FrmClientes.vb#L177) → `ClienteService.GetClientes()`.
* [`BtnNuevo_Click`](../../GestionComercial/Forms/FrmClientes.vb#L214) / [`BtnEditar_Click`](../../GestionComercial/Forms/FrmClientes.vb#L222) → `FrmClienteEditor` → [`ClienteService.GuardarCliente()`](../../GestionComercial/Services/ClienteService.vb#L112).
* [`BtnToggleActivo_Click`](../../GestionComercial/Forms/FrmClientes.vb#L241) → `ClienteService.ToggleActivo()` (protegiendo `Id = 1` Consumidor Final).

---

## 7. `FrmAutorizacionSupervisor.vb` — Confirmación y Control Dual

Archivo: [GestionComercial/Forms/FrmAutorizacionSupervisor.vb](../../GestionComercial/Forms/FrmAutorizacionSupervisor.vb)

* **Cuándo interviene en el flujo del Administrador:**
  * Si el Administrador confirma acciones que soliciten `SolicitarAutorizacionAdmin` con `forzarConfirmacionClave:=True`.
  * En `FrmBackup`, aunque el Administrador inició la acción, el sistema abre este modal pidiendo específicamente la contraseña de un **Gerente** (`SolicitarAutorizacionGerente`).
* [`ProcesarConfirmacion()`](../../GestionComercial/Forms/FrmAutorizacionSupervisor.vb#L103) llama a [`AuthService.AutenticarSupervisorPorRoles()`](../../GestionComercial/Services/AuthService.vb#L269) y deja asentada la auditoría en `AppLogger.LogAudit`.

