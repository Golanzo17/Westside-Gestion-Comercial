# 🛠️ Explicación de Otros Servicios Clave

En esta guía se explican los servicios complementarios del sistema:
1. `ClienteService.vb`
2. `UsuarioService.vb`
3. `ReporteService.vb`
4. `ConfiguracionService.vb`
5. `BackupService.vb`
6. `AppLogger.vb`

Los servicios de mayor peso tienen su propia guía: `AuthService`, `CajaService`, `CatalogService` y `VentaService`.

---

## 1. `ClienteService.vb` (Gestión de Clientes)
* **Ubicación**: [GestionComercial/Services/ClienteService.vb]

### Métodos:
* **`GetClientes(busqueda, soloActivos)`**:
  * Permite buscar clientes por DNI/CUIT, nombre, apellido o teléfono.
* **`GetClienteById(id)`**:
  * Devuelve los datos completos de un cliente puntual.
* **`GuardarCliente(cli, ByRef errorMessage)`**:
  * Da de alta o actualiza un cliente.
  * Valida que el nombre, apellido y DNI/CUIT no estén vacíos.
  * **Control preventivo de unicidad**: comprueba si el DNI ya está registrado por otro cliente (activo o inactivo) y devuelve un mensaje claro identificando al titular previo antes del intento de inserción.
* **`ExisteDni(dni, excludeId)`** y **`ExisteDni(dni, excludeId, ByRef clienteExistente)`**:
  * Dos versiones (sobrecarga): una solo responde si el DNI/CUIT existe; la otra además devuelve el cliente que lo tiene. Permiten excluir un ID para no chocar consigo mismo al editar.
* **`ToggleActivo(id, ByRef errorMessage)`**:
  * Aplica **borrado lógico**: alterna `activo` entre 1 y 0, así el historial de ventas no se rompe.
  * Impide desactivar al cliente con `ID = 1` ("Consumidor Final"), que usa el POS para ventas de mostrador.

---

## 2. `UsuarioService.vb` (Gestión de Personal y Cuentas)
* **Ubicación**: [GestionComercial/Services/UsuarioService.vb]
* **Roles disponibles** (`RolesDisponibles`): `Vendedor`, `Gerente` y `Administrador`.

### Métodos:
* **`GetUsuarios() As List(Of Usuario)`**:
  * Devuelve la lista de usuarios del sistema con sus roles, estados y DNI.
* **`GetUsuarioById(id)`**:
  * Devuelve un usuario puntual (lo usan `ToggleActivo` y la edición).
* **`CrearUsuario(username, password, nombre, apellido, rol, ByRef errorMessage, dni, ...)`**:
  * Crea un usuario encriptando la clave con **PBKDF2**.
  * Valida que el rol sea uno de los permitidos y la unicidad estricta tanto del **nombre de usuario** como del **DNI**.
* **`ActualizarUsuario(usuarioId, nombre, apellido, rol, ByRef errorMessage, dni, ...)`**:
  * Actualiza los datos validando que el DNI no pertenezca a otro empleado.
  * Si se intenta quitar el rol de Administrador al único administrador activo, lo rechaza.
* **`ExisteDni(dni, excludeId)`**:
  * Verifica preventivamente si un DNI ya existe en la tabla de usuarios.
* **`CambiarPassword(usuarioId, nuevaPassword, ByRef errorMessage)`**:
  * Permite resetear la clave de acceso de forma segura (se guarda el nuevo hash PBKDF2).
* **`ToggleActivo(usuarioId, ByRef errorMessage)`**:
  * Activa o desactiva al usuario (`activo = 1/0`) sin borrarlo, para conservar la trazabilidad de sus ventas y cajas.
  * Protege al sistema: no permite desactivar al **único Administrador activo**.
* Todas las altas, cambios y bajas quedan registradas en el log de auditoría (`AppLogger.LogAudit`).

---

## 3. `ReporteService.vb` (Estadísticas y Alertas)
* **Ubicación**: [GestionComercial/Services/ReporteService.vb]
* Para ordenar los rankings usa el enum `CriterioOrdenCliente`: `MayorFacturacion`, `MayorSalida` y `MayorMargen`.

### Métodos:
* **`GetResumenHoy() As ResumenVentasHoy`**:
  * Alimenta el **Dashboard Principal** en tiempo real: total recaudado hoy, cantidad de tickets, efectivo vs medios digitales y cantidad de prendas bajo el stock mínimo.
* **`GetTopProductosVendidos(limite, fechaDesde, fechaHasta, criterio)`**:
  * Ranking de prendas en un período, con dos modos: **Mayor Salida** (unidades vendidas) o **Mayor Margen de Ganancia** (utilidad neta en pesos y porcentaje de rentabilidad).
* **`GetReporteVentasPorCliente(fechaDesde, fechaHasta, clienteId, criterio)`**:
  * Compras por cliente en un período: tickets, prendas acumuladas, facturación total y margen aportado. Permite filtrar por un cliente o listar todos.
* **`GetReporteVentasPorVendedor(fechaDesde, fechaHasta, usuarioId, criterio)`**:
  * Desempeño de cada vendedor: cantidad de ventas, unidades, facturación y margen. Es la base del reporte de **ventas por vendedor** que consulta el Gerente.
* **`GetMetricasFinancierasPeriodo(fechaDesde, fechaHasta, usuarioId, clienteId, metodoPago) As MetricasFinancierasPeriodo`**:
  * Calcula los indicadores financieros del período sobre ventas **Completadas** (las anuladas no cuentan): total facturado, costo de mercadería, ganancia neta, porcentaje de margen, ticket promedio, cantidad de tickets y total en efectivo vs digital. Acepta filtros por vendedor, cliente y medio de pago. Alimenta las tarjetas KPI de `FrmReportes`.
* **`GetAlertasStockBajo() As List(Of AlertaStock)`**:
  * Devuelve las prendas que necesitan reposición (`stock_actual <= stock_minimo`).
* **`GetStockInmovilizado(fechaDesde, fechaHasta, limite) As List(Of PrendaStockInmovilizado)`**:
  * Detecta prendas activas con stock que **no tuvieron ventas en el período**, es decir, capital parado. Informa el costo inmovilizado (`stock × costo`), el valor de venta potencial y la fecha de la última venta histórica. Sirve para decidir liquidaciones o promociones.

---

## 4. `ConfiguracionService.vb` (Datos de la Empresa)
* **Ubicación**: [GestionComercial/Services/ConfiguracionService.vb]

### Métodos:
* **`GetConfiguracion() As ConfiguracionComercio`**:
  * Lee la tabla `configuracion` (nombre del local, CUIT, dirección, teléfono, email, condición de IVA, mensaje del ticket y símbolo de moneda).
* **`GuardarConfiguracion(cfg, ByRef errorMessage)`**:
  * Guarda los datos con un **UPSERT** (`INSERT ... ON CONFLICT(id) DO UPDATE`): crea la fila si no existe y la actualiza si existe. Esos datos salen impresos en los tickets.

---

## 5. `BackupService.vb` (Copias de Seguridad y Restauración)
* **Ubicación**: [GestionComercial/Services/BackupService.vb]
* **Propósito**: Gestión atómica de respaldos de la base SQLite mediante el *Online Backup API*.
* **Control dual**: el formulario `FrmBackup` exige la contraseña de un **Gerente** (`SolicitarAutorizacionGerente`) antes de llamar a crear, exportar o restaurar.

### Métodos:
* **`GetDirectorioBackups() As String`**:
  * Devuelve la carpeta de respaldos (`Database/Backups/`) y la crea si no existe.
* **`CrearBackup(ByRef mensajeError) As String`** y **`CrearBackup(rutaDestinoPersonalizada, ByRef mensajeError) As String`**:
  * Sobrecarga: la primera guarda en la carpeta por defecto con estampa temporal (`backup_gestion_comercial_yyyyMMdd_HHmmss.db`); la segunda permite **exportar** a una ruta elegida por el usuario.
  * Hace una copia consistente con `sourceConn.BackupDatabase(destConn)`, incluso con la app en uso.
* **`RestaurarBackup(rutaBackup, ByRef mensajeError) As Boolean`**:
  * Libera las conexiones activas con `SqliteConnection.ClearAllPools()` y reemplaza la base activa por el archivo elegido. **Sobrescribe todo**: no fusiona datos.
* **`GetBackups(Optional customDir) As List(Of BackupItem)`**:
  * Lista las copias disponibles (de la carpeta por defecto o de una indicada), de la más nueva a la más vieja, con su tamaño formateado.
* **`ObtenerInfoBaseDatosActual() As DatabaseInfo`**:
  * Metadatos de la base activa: tamaño en disco, cantidad de tablas y de tickets registrados.

---

## 6. `AppLogger.vb` (Registro de Eventos y Auditoría)
* **Ubicación**: [GestionComercial/Services/AppLogger.vb]
* **Propósito**: Registrar errores técnicos y acciones sensibles para poder investigar qué pasó. Es un `Module` (clase estática) usado por todos los demás servicios.
* **Dónde escribe**: un archivo por día en la carpeta `logs/` junto al ejecutable (`gestion_comercial_yyyyMMdd.log`) y en la salida de depuración. Usa `SyncLock` para que varios hilos no corrompan el archivo. Si el disco es de solo lectura, el error se ignora para no frenar la aplicación.

### Niveles:
* **`LogInfo(message)`**: eventos informativos.
* **`LogAudit(action, details, [usuarioId])`**: **auditoría** de acciones sensibles (login, logout, caja abierta, movimientos de caja, autorizaciones de supervisor, alta/baja de usuarios).
* **`LogWarning(message, [ex])`**: situaciones anómalas, por ejemplo un bloqueo por fuerza bruta.
* **`LogError(context, ex)`**: excepciones capturadas, con tipo y mensaje del error.

Formato de cada línea: `[2026-11-10 14:05:33.120] [AUDIT] AUTORIZACION_SUPERVISOR | Supervisor=gerente (Gerente) | Operador=admin | Motivo=...`
