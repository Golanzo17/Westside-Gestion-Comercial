# 🧾 Ruta de Vida por Formulario — Rol: Vendedor

Este documento explica **todas las pantallas y funciones disponibles para el rol `Vendedor`**, siguiendo la **ruta de vida completa** de cada botón o acción: desde que el usuario hace clic en la interfaz (`Forms/`), pasando por la lógica de negocio (`Services/`), hasta llegar a SQLite (`Data/DatabaseHelper.vb`) y regresar con el resultado a la pantalla.

---

## 🔐 Resumen de Permisos del Rol (`AuthService.vb`)

En [AuthService.vb](../../GestionComercial/Services/AuthService.vb) las propiedades de rol para este perfil devuelven:
* `AuthService.IsVendor` → **`True`**
* `AuthService.IsManager` → **`False`**
* `AuthService.IsAdmin` → **`False`**
* `AuthService.IsAdminOrManager` → **`False`**

Por regla de **Segregación de Funciones** en [FrmMain.vb](../../GestionComercial/Forms/FrmMain.vb#L136-L147):
* **Ve y opera:** Panel principal (`FrmMain`), Punto de Venta (`FrmVentasPOS`), Mis ventas del turno (`FrmMisVentasTurno`), Caja diaria (`FrmCaja`), Clientes (`FrmClientes`) y consulta del Catálogo (`FrmProductos`).
* **Tiene oculto / bloqueado:** Reposición de Stock (`FrmStock`), Auditoría de Cierres (`FrmControlCajas`), Reportes (`FrmReportes`), Usuarios (`FrmUsuarios`), Backups (`FrmBackup`) y Configuración (`FrmConfiguracion`).

---

## 📋 Mapa de Formularios del Vendedor

1. [FrmLogin.vb](#1-frmloginvb--inicio-de-sesión) — Ingreso al sistema
2. [FrmMain.vb](#2-frmmainvb--menú-principal-y-dashboard-del-vendedor) — Tablero principal y navegación
3. [FrmVentasPOS.vb](#3-frmventasposvb--punto-de-venta-pos) — Cobro de ventas, carrito, descuentos y ticket PDF
4. [FrmMisVentasTurno.vb](#4-frmmisventasturnovb--mis-ventas-del-turno-y-cambios-de-prenda) — Tickets del turno y cambios de talle/color
5. [FrmCaja.vb](#5-frmcajavb--caja-diaria-y-arqueo-ciego) — Apertura, ingresos/egresos y cierre ciego
6. [FrmClientes.vb](#6-frmclientesvb--directorio-y-abm-de-clientes) — Consulta, alta y edición de clientes
7. [FrmProductos.vb](#7-frmproductosvb--consulta-de-catálogo-y-acciones-supervisadas) — Consulta de prendas (edición requiere supervisor)
8. [FrmAutorizacionSupervisor.vb](#8-frmautorizacionsupervisorvb--modal-de-autorización-presencial) — Validación de clave de Gerente o Admin

---

## 1. `FrmLogin.vb` — Inicio de Sesión

Archivo: [GestionComercial/Forms/FrmLogin.vb](../../GestionComercial/Forms/FrmLogin.vb)

### 🔹 Al abrir la ventana: Verificación de Base de Datos
* **Dónde empieza:** Constructor [`New()`](../../GestionComercial/Forms/FrmLogin.vb#L28-L37) → [`CheckInitialConnection()`](../../GestionComercial/Forms/FrmLogin.vb#L158-L168).
* **Qué código ejecuta:** Llama a [`DatabaseHelper.TestConnection()`](../../GestionComercial/Data/DatabaseHelper.vb#L59) que abre conexión SQLite, crea carpeta/archivo si falta y ejecuta esquema/migraciones.
* **Resultado en pantalla:** Actualiza `lblStatus` en verde (`● Base de Datos SQLite Lista`) o rojo si falla.

### 🔹 Botón `INGRESAR AL SISTEMA` (o tecla `Enter`)
1. **Captura en UI ([`BtnLogin_Click`](../../GestionComercial/Forms/FrmLogin.vb#L171-L193)):**
   * Toma `txtUser.Text.Trim()` y `txtPass.Text` y valida que no estén vacíos.
2. **Regla de seguridad y consulta SQL ([`AuthService.Login`](../../GestionComercial/Services/AuthService.vb#L104-L156)):**
   * Llama a `EstaBloqueadoPorFuerzaBruta()` (bloquea 60 segundos si acumuló 5 fallos).
   * Ejecuta `DatabaseHelper.ExecuteQuery` buscando en la tabla `usuarios` donde `username = @username AND activo = 1`.
   * Llama a [`DatabaseHelper.VerifyPassword`](../../GestionComercial/Data/DatabaseHelper.vb#L207) (PBKDF2-SHA256 con 100.000 iteraciones y comparación `FixedTimeEquals`).
   * Si falla: `RegistrarFalloLogin()` resta intentos restantes y retorna `False`.
   * Si coincide: `LimpiarIntentosFallidos()`, guarda el usuario en `AuthService.CurrentUser`, ejecuta `UPDATE usuarios SET ultimo_login = NOW()` y registra auditoría en `AppLogger.LogAudit`.
3. **Resultado en pantalla:** Oculta `FrmLogin` (`Me.Hide()`) e instancia y muestra `FrmMain`.

---

## 2. `FrmMain.vb` — Menú Principal y Dashboard del Vendedor

Archivo: [GestionComercial/Forms/FrmMain.vb](../../GestionComercial/Forms/FrmMain.vb)

### 🔹 Carga de tarjetas KPI y estado de caja en barra superior
* **Dónde empieza:** [`New()`](../../GestionComercial/Forms/FrmMain.vb#L51-L61) → [`InitializeUI()`](../../GestionComercial/Forms/FrmMain.vb#L64-L293) + [`BuildDashboardHome()`](../../GestionComercial/Forms/FrmMain.vb#L339-L618) + [`LoadDashboardData()`](../../GestionComercial/Forms/FrmMain.vb#L631-L684).
* **Dónde está el código que trae los datos:**
  * [`ReporteService.GetResumenHoy()`](../../GestionComercial/Services/ReporteService.vb#L122): consulta stock bajo (`producto_talles`) para la tarjeta **PRENDAS STOCK CRÍTICO**.
  * [`CajaService.GetCajaAbierta(userId)`](../../GestionComercial/Services/CajaService.vb#L31): busca en `cajas` si el vendedor tiene una caja en estado `'Abierta'`.
  * [`CajaService.GetDesgloseMediosTurno(caja.Id)`](../../GestionComercial/Services/CajaService.vb#L365): consulta `ventas` de esa `caja_id` para mostrar **únicamente** lo recaudado y cantidad de tickets de **su propio turno** (nunca la facturación global del local).
  * La tarjeta 4 muestra `"Arqueo Ciego"` mientras la caja esté abierta (para no revelar el monto esperado) o `"$ 0,00"` si está cerrada.
  * `StartClock()` (L622) refresca `lblClock` cada 1 segundo.

### 🔹 Botones de Navegación (Sidebar y Acciones Rápidas) + Atajo `F1`
* **Dónde está el código:** En `FrmMain.vb` cada botón llama a [`OpenChildForm()`](../../GestionComercial/Forms/FrmMain.vb#L688-L704), que cierra la vista hija anterior e incrusta el nuevo formulario dentro de `pnlContentHost`:
  * `Panel principal` → [`Nav_Home`](../../GestionComercial/Forms/FrmMain.vb#L707) (vuelve al dashboard y refresca `LoadDashboardData()`).
  * `Punto de venta (POS)` / `+ NUEVA VENTA (F1)` / tecla `F1` → [`Nav_POS`](../../GestionComercial/Forms/FrmMain.vb#L719) → abre `FrmVentasPOS`.
  * `Mis ventas del turno` / `Mis Ventas del Turno / Cambios` → [`Nav_MisVentas`](../../GestionComercial/Forms/FrmMain.vb#L753) → abre `FrmMisVentasTurno`.
  * `Caja diaria y arqueo` / `Arqueo / Cierre de Caja` → [`Nav_Caja`](../../GestionComercial/Forms/FrmMain.vb#L746) → abre `FrmCaja`.
  * `Clientes` / `Directorio de Clientes` → [`Nav_Clientes`](../../GestionComercial/Forms/FrmMain.vb#L740) → abre `FrmClientes`.
  * `Catálogo y talles` → [`Nav_Productos`](../../GestionComercial/Forms/FrmMain.vb#L726) → abre `FrmProductos`.
  * `Cerrar sesión` → [`BtnLogout_Click`](../../GestionComercial/Forms/FrmMain.vb#L812) → llama a [`AuthService.Logout()`](../../GestionComercial/Services/AuthService.vb#L240), registra log de salida y vuelve a abrir `FrmLogin`.

---

## 3. `FrmVentasPOS.vb` — Punto de Venta (POS)

Archivo: [GestionComercial/Forms/FrmVentasPOS.vb](../../GestionComercial/Forms/FrmVentasPOS.vb)

### 🔹 Al abrir el POS: Carga de clientes activos
* **Ruta de vida:** `New()` → [`LoadInitialData()`](../../GestionComercial/Forms/FrmVentasPOS.vb#L451) → [`CargarClientes()`](../../GestionComercial/Forms/FrmVentasPOS.vb#L456) → [`ClienteService.GetClientes("", True)`](../../GestionComercial/Services/ClienteService.vb#L20) → `SELECT * FROM clientes WHERE activo = 1`.
* **Resultado:** Puebla `listaClientes` en memoria y llena `cboClientes` dejándolo sin selección por defecto para obligar a elegir cliente.

### 🔹 Buscador de Cliente (`txtBuscarCliente`, botón `🔍` y botón `+`)
* **Autocompletado al escribir:** [`TxtBuscarCliente_TextChanged`](../../GestionComercial/Forms/FrmVentasPOS.vb#L473) llama a `FiltrarClientesLocal()` (L516) filtrando en memoria por nombre, apellido o DNI/CUIT y despliega `lstSugerenciasCliente`.
* **Al confirmar con Enter, flechas o botón `🔍`:** `TxtBuscarCliente_KeyDown` / `BtnBuscarCliente_Click` → [`EjecutarBusquedaCliente()`](../../GestionComercial/Forms/FrmVentasPOS.vb#L573) → [`SeleccionarClienteEnCombo()`](../../GestionComercial/Forms/FrmVentasPOS.vb#L603).
* **Botón `+` (Nuevo Cliente rápido):** [`BtnNuevoCliente_Click`](../../GestionComercial/Forms/FrmVentasPOS.vb#L952) abre `FrmClientes` como diálogo modal y al cerrarse vuelve a ejecutar `CargarClientes()`.

### 🔹 Buscador de Prenda (Lector de Código de Barras / Texto) y Selección de Talle
1. **Mientras escribe:** [`TxtBuscarArticulo_TextChanged`](../../GestionComercial/Forms/FrmVentasPOS.vb#L624) → [`CatalogService.GetProductos(term, 0, True)`](../../GestionComercial/Services/CatalogService.vb#L82) → muestra `lstSugerenciasPOS`.
2. **Al presionar `Enter` (pistola lectora) o botón `Buscar`:** [`BtnBuscarArticulo_Click`](../../GestionComercial/Forms/FrmVentasPOS.vb#L710) / [`BuscarArticulo()`](../../GestionComercial/Forms/FrmVentasPOS.vb#L716) → llama a [`CatalogService.GetProductoPorCodigo(term)`](../../GestionComercial/Services/CatalogService.vb#L118) (consulta `productos` + `producto_talles` + `talles`). Si no hay código exacto, toma la primera coincidencia de `GetProductos()`.
3. **Dibujo de talles disponibles:** [`MostrarPrendaSeleccionada(prod)`](../../GestionComercial/Forms/FrmVentasPOS.vb#L739) limpia `pnlTallesContainer`, crea un `Button` por cada variante en `prod.TallesStock` (deshabilitando las que tienen `StockActual = 0`) y auto-selecciona el primero disponible.
4. **Clic en un botón de talle:** Dispara [`SeleccionarTalle(pt)`](../../GestionComercial/Forms/FrmVentasPOS.vb#L797), resalta el botón en azul, selecciona su color en `cboColor` y fija `numCantidad.Maximum = pt.StockActual`.

### 🔹 Botón `+ AGREGAR AL CARRITO`
* **Dónde está el código:** [`BtnAgregarAlCarrito_Click`](../../GestionComercial/Forms/FrmVentasPOS.vb#L823-L865).
* **Qué hace:**
  1. Verifica que exista `productoActual` y `talleSeleccionado`.
  2. Busca si esa combinación `(ProductoId, TalleId, Color)` ya está en `ventaActual.Detalles`: si ya existe, valida no superar `talleSeleccionado.StockActual` y suma la cantidad; si no existe, crea un nuevo `DetalleVenta` con `PrecioUnitario`, `CostoUnitario` y `Subtotal`.
  3. Llama a `RefrescarGrillaCarrito()` (L867) y `RecalcularTotales()` (L898).
  4. Limpia `txtBuscarArticulo` y le devuelve el foco para escanear la próxima prenda.

### 🔹 Botones `✖ Quitar Ítem Seleccionado` y `Vaciar Carrito`
* [`BtnQuitarItem_Click`](../../GestionComercial/Forms/FrmVentasPOS.vb#L875): quita el índice seleccionado de `ventaActual.Detalles`, refresca grilla y recalcula totales.
* [`BtnVaciarCarrito_Click`](../../GestionComercial/Forms/FrmVentasPOS.vb#L887): pide confirmación y limpia `ventaActual.Detalles`.

### 🔹 Control de Descuento (`numDescuentoPorc`) y Cálculo de Vuelto (`txtMontoAbonado`)
* **Totales:** `cboMetodoPago.SelectedIndexChanged` y `numDescuentoPorc.ValueChanged` llaman a [`RecalcularTotales()`](../../GestionComercial/Forms/FrmVentasPOS.vb#L898) (`Subtotal`, `DescuentoMonto`, `Total`) y `CalcularVuelto()`.
* **Regla de tope 15% para Vendedor:** Al salir de `numDescuentoPorc` (`.Leave`) o al cobrar, [`ValidarDescuentoVendedor()`](../../GestionComercial/Forms/FrmVentasPOS.vb#L935-L950) comprueba si `numDescuentoPorc.Value > 15`. Si supera 15% y no fue autorizado, abre `FrmAutorizacionSupervisor` vía [`AuthService.SolicitarAutorizacionAdminOManager()`](../../GestionComercial/Services/AuthService.vb#L319). Si el supervisor cancela o falla la clave, fuerza el descuento de vuelta a `15%`.
* **Vuelto:** [`TxtMontoAbonado_TextChanged`](../../GestionComercial/Forms/FrmVentasPOS.vb#L916) → [`CalcularVuelto()`](../../GestionComercial/Forms/FrmVentasPOS.vb#L921) muestra `abonado - Total` en `lblVuelto`.

### 🔹 Botón `COBRAR VENTA (F5)` (o tecla `F5`) — Ruta de Vida Completa
1. **Entrada en UI ([`FinalizarVenta()`](../../GestionComercial/Forms/FrmVentasPOS.vb#L972-L1076)):**
   * Valida que `ventaActual.Detalles.Count > 0`.
   * Ejecuta `ValidarDescuentoVendedor()`.
   * Consulta `CajaService.GetCajaAbierta(userId)`. Si no hay caja abierta, pregunta si desea abrirla y muestra `FrmCaja` modal.
   * Valida que haya un cliente elegido en `cboClientes`.
   * Llama a [`VentaService.GenerarNumeroTicket()`](../../GestionComercial/Services/VentaService.vb#L23) para previsualizar el código (`T-0001-XXXXXXXX`).
   * Calcula `MontoEfectivo` según el medio de pago (`Efectivo`, `Múltiple` o digital `0`) y `Vuelto`.
   * Muestra un `MessageBox` de resumen para que el vendedor confirme el cobro.
2. **Transacción Atómica en Servicio ([`VentaService.ProcesarVenta()`](../../GestionComercial/Services/VentaService.vb#L68-L142)):**
   * Abre conexión SQLite y `BeginTransaction()`.
   * `GarantizarNumeroTicket` (L343): genera/asegura número único dentro de la transacción.
   * `ValidarStockDisponible` (L360): consulta `SELECT stock_actual FROM producto_talles` con parámetros para evitar vender sin stock.
   * `InsertarVenta` (L424): `INSERT INTO ventas (...)` y obtiene el ID insertado con `DatabaseHelper.GetLastInsertedId`.
   * `InsertarDetalles` (L454): `INSERT INTO detalle_ventas (...)` por cada prenda del carrito.
   * `DescontarStockYAuditar` (L385): descuenta `UPDATE producto_talles SET stock_actual = stock_actual - @cant` e inserta traza `'Salida'` en `movimientos_stock` (`stock_anterior`, `stock_posterior`, motivo `"Venta Ticket #..."`).
   * `ActualizarCajaVenta` (L476): `UPDATE cajas SET total_ventas_efectivo = ... , total_ventas_digital = ...`.
   * `trans.Commit()` (si algo falla en cualquier paso, `trans.Rollback()` revierte todo).
   * Registra evento en `AppLogger.LogAudit("VENTA_PROCESADA", ...)`.
3. **Emisión de Ticket y Reseteo en Pantalla:**
   * Vuelve a `FinalizarVenta()`, llama a [`MostrarTicketImpresion(ventaActual)`](../../GestionComercial/Forms/FrmVentasPOS.vb#L1078) que dibuja la vista previa térmica con botón `📄 Guardar ticket en PDF` → [`FrmReportes.GuardarTicketEnPdf()`](../../GestionComercial/Forms/FrmReportes.vb#L1769) + [`ConstruirLineasTicket()`](../../GestionComercial/Forms/FrmReportes.vb#L1849).
   * Reinicia carrito, descuento, cliente y devuelve el foco a `txtBuscarArticulo`.

---

## 4. `FrmMisVentasTurno.vb` — Mis Ventas del Turno y Cambios de Prenda

Archivo: [GestionComercial/Forms/FrmMisVentasTurno.vb](../../GestionComercial/Forms/FrmMisVentasTurno.vb)

### 🔹 Carga y búsqueda de tickets del turno
* **Dónde empieza:** `New()` / `txtBuscarTicket.TextChanged` / `btnRefrescar.Click` → [`CargarVentasDelTurno()`](../../GestionComercial/Forms/FrmMisVentasTurno.vb#L200-L238).
* **Lógica en Servicios:**
  * Llama a [`CajaService.GetCajaAbierta(userId)`](../../GestionComercial/Services/CajaService.vb#L31). Si no hay turno abierto, limpia la grilla e informa que no hay caja activa.
  * Con `cajaAbierta.Id`, llama a [`VentaService.GetVentasPorTurnoVendedor(cajaId, userId, busqueda)`](../../GestionComercial/Services/VentaService.vb#L612) (`SELECT` filtrado por `caja_id`, `usuario_id` y búsqueda opcional por número de ticket o cliente).
* **Resultado:** Puebla `dgvMisVentas` y actualiza `lblResumenTurno` (cantidad de tickets y total vendido en el turno).

### 🔹 Botón `🔍 Ver Detalle de Ticket` (o doble clic en una fila)
* **Código:** [`BtnVerDetalleTicket_Click`](../../GestionComercial/Forms/FrmMisVentasTurno.vb#L250) → [`MostrarDetalleTicketVendedor()`](../../GestionComercial/Forms/FrmMisVentasTurno.vb#L276-L409).
* **Servicio:** Llama a [`VentaService.GetDetalleVenta(ventaId)`](../../GestionComercial/Services/VentaService.vb#L200) para listar todas las prendas de `detalle_ventas`.
* **Acciones dentro del modal:**
  * `📄 Guardar Ticket PDF` → `FrmReportes.GuardarTicketEnPdf(frmDet, ventaSel)`.
  * `🔄 Cambio de Prenda / Talle` → cierra el modal y llama a `MostrarModalCambioPrenda(ventaSel)`.

### 🔹 Botón `📄 Reimprimir / PDF Ticket`
* **Código:** [`BtnReimprimirPdfTicket_Click`](../../GestionComercial/Forms/FrmMisVentasTurno.vb#L257).
* **Qué hace:** Verifica con `ObtenerVentaSeleccionada()` (L240), carga los ítems desde `VentaService.GetDetalleVenta(ventaSel.Id)` y llama a [`FrmReportes.GuardarTicketEnPdf()`](../../GestionComercial/Forms/FrmReportes.vb#L1769) para generar el `.pdf` con *Microsoft Print to PDF*.

### 🔹 Botón `🔄 Cambio de Prenda / Talle` — Ruta de Vida Completa
1. **Apertura del modal ([`BtnCambioPrenda_Click`](../../GestionComercial/Forms/FrmMisVentasTurno.vb#L265) → [`MostrarModalCambioPrenda()`](../../GestionComercial/Forms/FrmMisVentasTurno.vb#L411-L560)):**
   * Valida que la venta seleccionada tenga `Estado = "Completada"`.
   * Carga los renglones con `VentaService.GetDetalleVenta(ventaSel.Id)` en `cboItemTicket`.
   * Cada vez que el vendedor elige una prenda (`cboItemTicket.SelectedIndexChanged`), consulta [`CatalogService.GetStockPorTalles(detSel.ProductoId)`](../../GestionComercial/Services/CatalogService.vb#L151) para ofrecer únicamente otros talles/colores con `StockActual >= detSel.Cantidad`.
2. **Al presionar `Confirmar Cambio de Prenda` (L532):**
   * Llama a [`VentaService.ProcesarCambioPrenda()`](../../GestionComercial/Services/VentaService.vb#L651-L834) dentro de una transacción SQLite:
     1. Verifica que la venta siga en estado `'Completada'` y lee la fila original de `detalle_ventas`.
     2. Consulta `producto_talles` de la nueva variante y valida stock suficiente.
     3. **Devolución de la prenda vieja:** suma la cantidad en `producto_talles` y registra movimiento `'Entrada'` en `movimientos_stock`.
     4. **Entrega de la prenda nueva:** resta la cantidad en `producto_talles` y registra movimiento `'Salida'` en `movimientos_stock`.
     5. Actualiza `detalle_ventas` (`talle_id = @nuevoTalleId`, `color = @nuevoColor`) y anexa nota de trazabilidad en `ventas.motivo_anulacion`.
     6. Confirma `trans.Commit()` y registra `AppLogger.LogAudit("CAMBIO_PRENDA", ...)`.
3. **Resultado en pantalla:** Cierra el modal con `DialogResult.OK` y recarga `CargarVentasDelTurno()`.

---

## 5. `FrmCaja.vb` — Caja Diaria y Arqueo Ciego

Archivo: [GestionComercial/Forms/FrmCaja.vb](../../GestionComercial/Forms/FrmCaja.vb)

### 🔹 Carga de estado y movimientos (`LoadCajaData`)
* **Código:** [`LoadCajaData()`](../../GestionComercial/Forms/FrmCaja.vb#L209-L251).
* **Servicios que consulta:**
  * [`CajaService.GetCajaAbierta(userId)`](../../GestionComercial/Services/CajaService.vb#L31).
  * Si está abierta: [`CajaService.GetDesgloseMediosTurno(cajaActual.Id)`](../../GestionComercial/Services/CajaService.vb#L365) y [`CajaService.GetMovimientosCaja(cajaActual.Id)`](../../GestionComercial/Services/CajaService.vb#L187).
  * **Seguridad de Arqueo Ciego:** Como `AuthService.IsAdminOrManager` es `False`, el formulario oculta el monto esperado de efectivo y muestra `"ARQUEO CIEGO"` en la tarjeta de saldo, habilitando `btnNuevoMovimiento` y `btnCerrarCaja`.

### 🔹 Botón `Abrir Caja del Día`
* **Código:** [`BtnAbrirCaja_Click`](../../GestionComercial/Forms/FrmCaja.vb#L254-L287).
* **Qué hace:** Abre un diálogo pidiendo el fondo inicial de cambio, y llama a [`CajaService.AbrirCaja(userId, montoInicial, errMsg)`](../../GestionComercial/Services/CajaService.vb#L58), que verifica que el vendedor no tenga otra caja abierta y ejecuta `INSERT INTO cajas (...)`. Luego refresca `LoadCajaData()`.

### 🔹 Botón `Registrar Ingreso / Egreso`
* **Código:** [`BtnNuevoMovimiento_Click`](../../GestionComercial/Forms/FrmCaja.vb#L290-L348).
* **Control de seguridad:** Permite registrar `"Ingreso"` o `"Egreso"`. Si el vendedor elige **`Egreso`** (retiro de efectivo / pago de gasto), **exige autorización presencial** vía [`AuthService.SolicitarAutorizacionAdminOManager()`](../../GestionComercial/Services/AuthService.vb#L319) (abre `FrmAutorizacionSupervisor`).
* **Capa de Servicio:** Llama a [`CajaService.RegistrarMovimiento(...)`](../../GestionComercial/Services/CajaService.vb#L89) dentro de una transacción: inserta en `movimientos_caja` (incluyendo `autorizado_por_id`) y actualiza `total_ingresos` o `total_egresos` en `cajas`.

### 🔹 Botón `Cerrar Caja / Arqueo`
* **Código:** [`BtnCerrarCaja_Click`](../../GestionComercial/Forms/FrmCaja.vb#L351-L448).
* **Arqueo Ciego:** Como el usuario es Vendedor (`esAdmin = False`), oculta `lblEsp` y `lblDif`: el vendedor **solo cuenta sus billetes físicos**, escribe el monto real y las observaciones.
* **Capa de Servicio:** Al confirmar, llama a [`CajaService.CerrarCaja(cajaActual.Id, numReal.Value, txtObs, errMsg)`](../../GestionComercial/Services/CajaService.vb#L143), que recalcula `monto_esperado = monto_inicial + total_ventas_efectivo + total_ingresos - total_egresos`, calcula `diferencia = monto_real - monto_esperado`, actualiza `estado = 'Cerrada'`, `fecha_cierre = NOW()` en `cajas` y graba auditoría en `AppLogger`.

---

## 6. `FrmClientes.vb` — Directorio y ABM de Clientes

Archivo: [GestionComercial/Forms/FrmClientes.vb](../../GestionComercial/Forms/FrmClientes.vb)

### 🔹 Búsqueda y carga de grilla
* **Código:** `txtBuscar.KeyDown` / `btnBuscar.Click` / `chkMostrarInactivos.CheckedChanged` → [`LoadClientes()`](../../GestionComercial/Forms/FrmClientes.vb#L177-L196).
* **Servicio:** [`ClienteService.GetClientes(filtro, soloActivos)`](../../GestionComercial/Services/ClienteService.vb#L20) → `SELECT * FROM clientes` filtrado por nombre, apellido, DNI o teléfono. `DgvClientes_CellFormatting` colorea `'Activo'` en verde e `'Inactivo'` en rojo.

### 🔹 Botones `+ Nuevo Cliente` y `✏ Editar` (o doble clic)
* **Código:** [`BtnNuevo_Click`](../../GestionComercial/Forms/FrmClientes.vb#L214) abre `New FrmClienteEditor(0)`; [`BtnEditar_Click`](../../GestionComercial/Forms/FrmClientes.vb#L222) / [`EditarSeleccionado()`](../../GestionComercial/Forms/FrmClientes.vb#L227) abre `New FrmClienteEditor(id)`.
* **Dentro de `FrmClienteEditor`:**
  * `LoadData()` (L357): si es edición, llama a [`ClienteService.GetClienteById(_clienteId)`](../../GestionComercial/Services/ClienteService.vb#L46).
  * `BtnGuardar_Click` (L382): valida campos obligatorios (DNI, Nombre, Apellido, Teléfono, Email, Dirección, Ciudad) y llama a [`ClienteService.GuardarCliente(clienteActual, errMsg)`](../../GestionComercial/Services/ClienteService.vb#L112), que verifica unicidad del DNI con `ExisteDni()` y ejecuta `INSERT` o `UPDATE` en la tabla `clientes`.

### 🔹 Botón `🚫 Activar / Desactivar`
* **Código:** [`BtnToggleActivo_Click`](../../GestionComercial/Forms/FrmClientes.vb#L241-L263).
* **Qué hace:** Impide desactivar el cliente ID `1` (`Consumidor Final`). Para cualquier otro, pide confirmación y llama a [`ClienteService.ToggleActivo(id, errMsg)`](../../GestionComercial/Services/ClienteService.vb#L204) (`UPDATE clientes SET activo = ...`).

---

## 7. `FrmProductos.vb` — Consulta de Catálogo y Acciones Supervisadas

Archivo: [GestionComercial/Forms/FrmProductos.vb](../../GestionComercial/Forms/FrmProductos.vb)

* **Consulta libre para el Vendedor:**
  * [`LoadCategorias()`](../../GestionComercial/Forms/FrmProductos.vb#L232) → `CatalogService.GetCategorias()`.
  * [`LoadProductos()`](../../GestionComercial/Forms/FrmProductos.vb#L243) → [`CatalogService.GetProductos(filtro, catId, soloActivos)`](../../GestionComercial/Services/CatalogService.vb#L82) muestra código, nombre, categoría, precio y stock total.
  * El doble clic en grilla (`DgvProductos_CellDoubleClick` L300) está bloqueado si no es `IsAdminOrManager`.
* **Botones protegidos (`+ Nueva Prenda`, `✏ Editar`, `🚫 Activar / Desactivar`):**
  * `BtnNuevo_Click` (L286), `EditarSeleccionado` (L307) y `BtnToggleActivo_Click` (L321) llaman primero a `AuthService.SolicitarAutorizacionAdminOManager(...)`, abriendo `FrmAutorizacionSupervisor`. Solo si un Gerente o Administrador ingresa su clave en mostrador se abre `FrmProductoEditor` o se cambia el estado en `CatalogService`.

---

## 8. `FrmAutorizacionSupervisor.vb` — Modal de Autorización Presencial

Archivo: [GestionComercial/Forms/FrmAutorizacionSupervisor.vb](../../GestionComercial/Forms/FrmAutorizacionSupervisor.vb)

* **Cuándo aparece al Vendedor:** Descuento > 15% en `FrmVentasPOS`, egreso de efectivo en `FrmCaja`, o intento de alta/edición/baja de prendas en `FrmProductos`.
* **Botón `Autorizar` (`Enter`):** [`ProcesarConfirmacion()`](../../GestionComercial/Forms/FrmAutorizacionSupervisor.vb#L103-L122) → llama a [`AuthService.AutenticarSupervisorPorRoles()`](../../GestionComercial/Services/AuthService.vb#L269) → valida contraseña contra la tabla `usuarios` filtrando por roles permitidos (con bloqueo de 60s tras 5 fallos) y registra al supervisor en `AuthService.LastAuthorizingSupervisor` + `AppLogger.LogAudit`.

