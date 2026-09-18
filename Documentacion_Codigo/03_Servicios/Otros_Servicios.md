# 🛠️ Explicación de Otros Servicios Clave

En esta guía se explican los 4 servicios complementarios del sistema:
1. `ClienteService.vb`
2. `UsuarioService.vb`
3. `ReporteService.vb`
4. `ConfiguracionService.vb`

---

## 1. `ClienteService.vb` (Gestión de Clientes y Cuentas)
* **Ubicación**: [GestionComercial/Services/ClienteService.vb]

### Métodos:
* **`GetClientes(busqueda, soloActivos)`**:
  * Permite buscar clientes por DNI/CUIT, nombre, apellido o teléfono.
* **`GetClienteById(id)`**:
  * Devuelve los datos completos de un cliente puntual.
* **`GuardarCliente(cli, ByRef errorMessage)`**:
  * Da de alta o actualiza un cliente.
  * Valida que el nombre, apellido y DNI/CUIT no estén vacíos.
  * **Control preventivo de unicidad**: Comprueba si el DNI ya está registrado por otro cliente (activo o inactivo) y devuelve un mensaje claro identificando al titular previo antes del intento de inserción.
* **`ExisteDni(dni, excludeId, ByRef clienteExistente)`**:
  * Consulta si un DNI/CUIT ya existe en el sistema excluyendo un ID opcional.
* **`EliminarCliente(id, ByRef errorMessage)`**:
  * Aplica borrado lógico (`activo = 0`). Impide eliminar al cliente con `ID = 1` (que es el "Consumidor Final" por defecto).

---

## 2. `UsuarioService.vb` (Gestión de Personal y Cuentas)
* **Ubicación**: [GestionComercial/Services/UsuarioService.vb]

### Métodos:
* **`GetUsuarios() As List(Of Usuario)`**:
  * Devuelve la lista de usuarios del sistema con sus roles, estados y DNI.
* **`CrearUsuario(username, password, nombre, apellido, rol, ByRef errorMessage, dni, ...)`**:
  * Crea un usuario encriptando la clave con **PBKDF2**.
  * Valida unicidad estricta tanto del **nombre de usuario (`username`)** como del **DNI**, impidiendo duplicados.
* **`ActualizarUsuario(usuarioId, nombre, apellido, rol, ByRef errorMessage, dni, ...)`**:
  * Actualiza los datos del usuario validando que el DNI no esté utilizado por otro empleado y preservando al menos un administrador activo.
* **`ExisteDni(dni, excludeId)`**:
  * Verifica preventivamente la existencia de un DNI en la tabla de usuarios.
* **`CambiarPassword(usuarioId, nuevaPassword, ByRef errorMessage)`**:
  * Permite al usuario o a un administrador resetear la clave de acceso de forma segura.
* **`EliminarUsuario(id, ByRef errorMessage)`**:
  * Desactiva al usuario (`activo = 0`). Protege al usuario con `ID = 1` ("admin") para evitar que el sistema se quede sin acceso administrativo.

---

## 3. `ReporteService.vb` (Estadísticas y Alertas)
* **Ubicación**: [GestionComercial/Services/ReporteService.vb]

### Métodos:
* **`GetResumenHoy() As ResumenVentasHoy`**:
  * Alimenta el **Dashboard Principal** en tiempo real:
    * Total recaudado hoy.
    * Cantidad de tickets emitidos hoy.
    * Total cobrado en Efectivo vs Total cobrado con Medios Digitales (Tarjeta/Transferencia).
    * Cantidad de prendas que están por debajo del stock mínimo.
* **`GetTopProductosVendidos(top, fechaDesde, fechaHasta)`**:
  * Ranking de las prendas más vendidas del local (las que más rotación y recaudación generan) mediante consultas `SUM(dv.cantidad)` agrupadas por producto y talle.
* **`GetAlertasStockBajo() As List(Of AlertaStock)`**:
  * Devuelve la lista de prendas que necesitan reposición urgente (`stock_actual <= stock_minimo`).
* **`GetReportePeriodo(desde, hasta)`**:
  * Muestra el resumen global de facturación entre dos fechas para balances mensuales o anuales.

---

## 4. `ConfiguracionService.vb` (Datos de la Empresa)
* **Ubicación**: [GestionComercial/Services/ConfiguracionService.vb]

### Métodos:
* **`GetConfiguracion() As ConfiguracionComercio`**:
  * Lee la tabla `configuracion_comercio` (nombre del local, CUIT, dirección, teléfono, pie de ticket).
* **`GuardarConfiguracion(cfg, ByRef errorMessage)`**:
  * Actualiza los datos de la empresa para que salgan impresos en los comprobantes de venta.
