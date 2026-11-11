# 🧑‍💼 Ruta de Vida por Formulario — Rol: Gerente

Este documento detalla **todas las pantallas y botones disponibles para el rol `Gerente`**, mostrando el recorrido completo del código: desde el evento del botón en el formulario (`Forms/`), pasando por las reglas de negocio en los servicios (`Services/`), hasta las consultas y transacciones en SQLite (`Data/DatabaseHelper.vb`).

---

## 🔐 Resumen de Permisos del Rol (`AuthService.vb`)

En [AuthService.vb](../../GestionComercial/Services/AuthService.vb):
* `AuthService.IsManager` → **`True`**
* `AuthService.IsAdminOrManager` → **`True`**
* `AuthService.IsVendor` → **`False`**
* `AuthService.IsAdmin` → **`False`**

Por diseño de control interno en [FrmMain.vb](../../GestionComercial/Forms/FrmMain.vb#L136-L147):
* **Ve y opera:** Panel principal gerencial (`FrmMain`), Catálogo y Talles (`FrmProductos` + `FrmProductoEditor`), Stock y Reposición (`FrmStock` + `FrmAjusteStock`), Cierres y Arqueos (`FrmControlCajas` + `FrmDetalleArqueoCaja`), Reportes y Anulaciones (`FrmReportes`) y Clientes (`FrmClientes`).
* **Tiene oculto / restringido:** Terminal POS (`FrmVentasPOS`), Mis Ventas del Turno (`FrmMisVentasTurno`), Caja propia (`FrmCaja`), ABM de Usuarios (`FrmUsuarios`), Backups (`FrmBackup`) y Configuración (`FrmConfiguracion`).
* **Rol de supervisor:** Su contraseña habilita autorizaciones presenciales (`FrmAutorizacionSupervisor`) cuando un Vendedor necesita descuento >15% o retiro de caja, o cuando el Administrador ejecuta/restaura un Backup (política de control dual / 4 ojos).

---

## 📋 Mapa de Formularios del Gerente

1. [FrmLogin.vb](#1-frmloginvb--inicio-de-sesión) — Ingreso al sistema
2. [FrmMain.vb](#2-frmmainvb--dashboard-gerencial-y-navegación) — Métricas globales del local y menú
3. [FrmProductos.vb](#3-frmproductosvb--catálogo-precios-y-matriz-de-talles) — Alta, edición, precios y matriz de stock
4. [FrmStock.vb](#4-frmstockvb--control-de-stock-y-reposición) — Alertas críticas, historial y ajustes manuales
5. [FrmControlCajas.vb](#5-frmcontrolcajasvb--auditoría-de-cierres-y-arqueos-de-caja) — Supervisión de turnos y exportación CSV
6. [FrmDetalleArqueoCaja.vb](#6-frmdetallearqueocajavb--auditoría-profunda-de-un-turno) — Fórmula matemática, medios de pago y acta
7. [FrmReportes.vb](#7-frmreportesvb--inteligencia-comercial-pdf-y-anulaciones) — Métricas financieras, rankings y anulación de ventas
8. [FrmClientes.vb](#8-frmclientesvb--directorio-y-abm-de-clientes) — Consulta, alta, edición y baja lógica
9. [FrmAutorizacionSupervisor.vb](#9-frmautorizacionsupervisorvb--autorizaciones-presenciales) — Control dual sobre Vendedor y Administrador

---

## 1. `FrmLogin.vb` — Inicio de Sesión

Archivo: [GestionComercial/Forms/FrmLogin.vb](../../GestionComercial/Forms/FrmLogin.vb)

* **Verificación inicial:** `New()` → [`CheckInitialConnection()`](../../GestionComercial/Forms/FrmLogin.vb#L158) → [`DatabaseHelper.TestConnection()`](../../GestionComercial/Data/DatabaseHelper.vb#L59).
* **Botón `INGRESAR AL SISTEMA`:** [`BtnLogin_Click`](../../GestionComercial/Forms/FrmLogin.vb#L171) → [`AuthService.Login()`](../../GestionComercial/Services/AuthService.vb#L104) verifica bloqueo anti fuerza bruta, consulta `usuarios`, valida PBKDF2 con [`DatabaseHelper.VerifyPassword`](../../GestionComercial/Data/DatabaseHelper.vb#L207), actualiza `ultimo_login`, registra `AppLogger.LogAudit` y abre `FrmMain`.

---

## 2. `FrmMain.vb` — Dashboard Gerencial y Navegación

Archivo: [GestionComercial/Forms/FrmMain.vb](../../GestionComercial/Forms/FrmMain.vb)

### 🔹 Carga de KPIs Gerenciales (`LoadDashboardData`)
* **Código:** [`BuildDashboardHome()`](../../GestionComercial/Forms/FrmMain.vb#L339) + [`LoadDashboardData()`](../../GestionComercial/Forms/FrmMain.vb#L631-L684).
* **Qué ejecuta:**
  * [`ReporteService.GetResumenHoy()`](../../GestionComercial/Services/ReporteService.vb#L122): llena **VENTAS DE HOY** (facturación global del día), **TICKETS EMITIDOS** y **PRENDAS STOCK CRÍTICO**.
  * [`CajaService.GetHistorialCajasAvanzado(..., "Abierta", Nothing)`](../../GestionComercial/Services/CajaService.vb#L249): cuenta cuántos turnos de vendedores están abiertos en ese momento y lo muestra en la tarjeta **CAJAS EN CURSO** (el Gerente no tiene caja propia).
  * `StartClock()` (L622) mantiene actualizado el reloj de la barra superior.

### 🔹 Accesos del Sidebar y Acciones Rápidas (`OpenChildForm`)
* `Panel principal` → [`Nav_Home`](../../GestionComercial/Forms/FrmMain.vb#L707)
* `Catálogo y talles` → [`Nav_Productos`](../../GestionComercial/Forms/FrmMain.vb#L726) → `FrmProductos`
* `Stock y reposición` → [`Nav_Stock`](../../GestionComercial/Forms/FrmMain.vb#L733) → `FrmStock`
* `Cierres y arqueos` → [`Nav_ControlCajas`](../../GestionComercial/Forms/FrmMain.vb#L760) → `FrmControlCajas`
* `Reportes y ventas` → [`Nav_Reportes`](../../GestionComercial/Forms/FrmMain.vb#L770) → `FrmReportes`
* `Clientes` → [`Nav_Clientes`](../../GestionComercial/Forms/FrmMain.vb#L740) → `FrmClientes`
* `Cerrar sesión` → [`BtnLogout_Click`](../../GestionComercial/Forms/FrmMain.vb#L812) → `AuthService.Logout()` y regresa a `FrmLogin`.

---

## 3. `FrmProductos.vb` — Catálogo, Precios y Matriz de Talles

Archivo: [GestionComercial/Forms/FrmProductos.vb](../../GestionComercial/Forms/FrmProductos.vb)

### 🔹 Carga y filtrado del catálogo
* [`LoadCategorias()`](../../GestionComercial/Forms/FrmProductos.vb#L232) → `CatalogService.GetCategorias()` (tabla `categorias`).
* [`LoadProductos()`](../../GestionComercial/Forms/FrmProductos.vb#L243) → [`CatalogService.GetProductos(filtro, catId, soloActivos)`](../../GestionComercial/Services/CatalogService.vb#L82) consulta `productos` + `categorias` + suma de stock en `producto_talles`.
* [`DgvProductos_CellFormatting()`](../../GestionComercial/Forms/FrmProductos.vb#L270) resalta stock en `0` (rojo) o `<= 5` (amarillo) y estado Activo/Inactivo.

### 🔹 Botones `+ Nueva Prenda` y `✏ Editar Prenda` (o doble clic)
* **Inicio:** [`BtnNuevo_Click`](../../GestionComercial/Forms/FrmProductos.vb#L286), [`BtnEditar_Click`](../../GestionComercial/Forms/FrmProductos.vb#L295), [`DgvProductos_CellDoubleClick`](../../GestionComercial/Forms/FrmProductos.vb#L300) → [`EditarSeleccionado()`](../../GestionComercial/Forms/FrmProductos.vb#L307).
* **Autorización:** Llama a `AuthService.SolicitarAutorizacionAdminOManager()`, que retorna `True` inmediatamente porque el usuario conectado ya es Gerente. Abre `FrmProductoEditor(prodId)`.

### 🔹 Dentro del modal `FrmProductoEditor`
* **Carga de datos:** [`LoadData()`](../../GestionComercial/Forms/FrmProductos.vb#L492) trae `CatalogService.GetCategorias()`, `CatalogService.GetTalles()` y, si es edición, `CatalogService.GetProductoPorCodigo()` para poblar la matriz editable `dgvTalles` (Talle, Color, Stock, Stock Mínimo, SKU).
* **Cálculo automático de precio:** Al modificar `numCosto` o `numGanancia`, [`CalcularPrecios()`](../../GestionComercial/Forms/FrmProductos.vb#L544) calcula `PrecioVenta = Costo * (1 + Ganancia / 100)`.
* **Botón `💾 Guardar Prenda` ([`BtnGuardar_Click`](../../GestionComercial/Forms/FrmProductos.vb#L551-L611)):**
  1. Valida código de barra, nombre, categoría y arma la lista `productoActual.TallesStock` desde `dgvTalles`.
  2. Llama a [`CatalogService.GuardarProducto(productoActual, errMsg)`](../../GestionComercial/Services/CatalogService.vb#L176): dentro de una transacción SQLite hace `INSERT` o `UPDATE` en `productos` y luego `GuardarMatrizTalles` (L322) inserta o actualiza cada variante en `producto_talles`.
  3. Registra auditoría en `AppLogger` y al cerrar recarga `LoadProductos()`.

### 🔹 Botón `🚫 Activar / Desactivar`
* **Código:** [`BtnToggleActivo_Click`](../../GestionComercial/Forms/FrmProductos.vb#L321-L340) → pide confirmación → [`CatalogService.ToggleActivo(prodId, errMsg)`](../../GestionComercial/Services/CatalogService.vb#L297) (`UPDATE productos SET activo = ...`).

---

## 4. `FrmStock.vb` — Control de Stock y Reposición

Archivo: [GestionComercial/Forms/FrmStock.vb](../../GestionComercial/Forms/FrmStock.vb)

### 🔹 Carga de alertas críticas e historial de auditoría (`LoadStockData`)
* **Dónde está el código:** [`LoadStockData()`](../../GestionComercial/Forms/FrmStock.vb#L251-L287) (disparado al iniciar y con `🔄 Actualizar`).
* **Qué ejecuta:**
  * Grilla superior (Alertas): llama a [`ReporteService.GetAlertasStockBajo()`](../../GestionComercial/Services/ReporteService.vb#L415) (`SELECT` en `producto_talles` JOIN `productos`/`talles` donde `stock_actual <= stock_minimo`). Colorea `AGOTADO` en rojo y `REPONER URGENTE` en amarillo.
  * Grilla inferior (Historial): llama directamente a `DatabaseHelper.ExecuteQuery` trayendo los últimos 100 registros de `movimientos_stock` ordenados por fecha descendente.

### 🔹 Botón `+ Ingreso de Mercadería / Ajuste` → Modal `FrmAjusteStock`
* **Apertura:** [`BtnIngresoMercaderia_Click`](../../GestionComercial/Forms/FrmStock.vb#L292) verifica `AuthService.SolicitarAutorizacionAdminOManager` (pasa directo por ser Gerente) y abre `FrmAjusteStock`.
* **Buscador predictivo de prendas:**
  * [`CargarCatalogo()`](../../GestionComercial/Forms/FrmStock.vb#L446) llama a `CatalogService.GetProductos("", 0, True)`.
  * [`TxtBuscarPrenda_TextChanged`](../../GestionComercial/Forms/FrmStock.vb#L468) filtra en memoria y muestra `lstSugerencias`; [`ConfirmarSeleccionSugerencia()`](../../GestionComercial/Forms/FrmStock.vb#L527) → [`SeleccionarPrenda()`](../../GestionComercial/Forms/FrmStock.vb#L538) habilita los controles de talle/color.
* **Selector de color y alta rápida (`+`):**
  * [`CargarColores()`](../../GestionComercial/Forms/FrmStock.vb#L454) llama a `CatalogService.GetColoresDisponibles()`.
  * Botón `+` o ítem `[+ Crear nuevo color...]` → [`CrearNuevoColor()`](../../GestionComercial/Forms/FrmStock.vb#L561) abre diálogo para agregar un color nuevo al combo.
* **Botón `💾 Registrar Movimiento` ([`BtnGuardar_Click`](../../GestionComercial/Forms/FrmStock.vb#L649-L686)):**
  1. Valida prenda, talle y cantidad distinta de `0` (admite positivos para ingresos y negativos para mermas/ajustes).
  2. Llama a [`CatalogService.AjustarStockManual(productoId, talleId, color, cant, motivo, userId, errMsg)`](../../GestionComercial/Services/CatalogService.vb#L254):
     * Abre transacción SQLite (`BeginTransaction`).
     * `ObtenerOInsertarStockFila` (L346): obtiene el stock actual en `producto_talles` (o crea la fila si es un talle/color nuevo).
     * Verifica que el nuevo stock no quede negativo (`nuevoStock < 0`).
     * `ActualizarStockFila` (L372): `UPDATE producto_talles SET stock_actual = @nuevo`.
     * `RegistrarAuditoriaStock` (L386): `INSERT INTO movimientos_stock` registrando tipo (`Entrada` o `Salida`), `stock_anterior`, `stock_posterior`, `motivo` y `usuario_id`.
     * `trans.Commit()` + `AppLogger.LogAudit`.
  3. Cierra el modal y refresca `LoadStockData()`.

---

## 5. `FrmControlCajas.vb` — Auditoría de Cierres y Arqueos de Caja

Archivo: [GestionComercial/Forms/FrmControlCajas.vb](../../GestionComercial/Forms/FrmControlCajas.vb)

### 🔹 Filtros, KPIs y Grilla de Auditoría
* **Carga inicial de filtro de vendedores:** [`CargarFiltroVendedores()`](../../GestionComercial/Forms/FrmControlCajas.vb#L458) → `UsuarioService.GetUsuarios(soloActivos:=False)`.
* **Filtros rápidos de período:** [`CboRapidoFechas_SelectedIndexChanged`](../../GestionComercial/Forms/FrmControlCajas.vb#L471) actualiza `dtpDesde`/`dtpHasta` y llama a `CargarDatosAuditoria()`. Botón `Restablecer` ([`BtnLimpiar_Click`](../../GestionComercial/Forms/FrmControlCajas.vb#L504)) restaura últimos 90 días.
* **Botón `🔍 Filtrar` / `🔄 Actualizar` ([`CargarDatosAuditoria()`](../../GestionComercial/Forms/FrmControlCajas.vb#L517-L626)):**
  * Arma filtros de rango, vendedor, estado (`Cerrada`/`Abierta`) y tipo de discrepancia (`CON_DESCUADRE`, `FALTANTE`, `SOBRANTE`, `EXACTO`).
  * Llama a [`CajaService.GetHistorialCajasAvanzado(...)`](../../GestionComercial/Services/CajaService.vb#L249) (`SELECT` en `cajas` JOIN `usuarios`).
  * Puebla `dgvCajas` y calcula las 4 tarjetas KPI superiores (**TURNOS AUDITADOS**, **RECAUDACIÓN TOTAL**, **EFECTIVO ESPERADO** vs **Declarado**, **BALANCE DE ARQUEO**).
  * [`DgvCajas_CellFormatting()`](../../GestionComercial/Forms/FrmControlCajas.vb#L628) colorea diferencias: `0` en verde, `> 0` sobrante en azul, `< 0` faltante en rojo.

### 🔹 Botones de Inspección del Turno Seleccionado
* **`🔍 Auditar Turno Seleccionado` (o doble clic):** [`BtnAuditar_Click`](../../GestionComercial/Forms/FrmControlCajas.vb#L683) / [`DgvCajas_CellDoubleClick`](../../GestionComercial/Forms/FrmControlCajas.vb#L676) → [`AbrirAuditoriaSeleccionada()`](../../GestionComercial/Forms/FrmControlCajas.vb#L688) abre modal `FrmDetalleArqueoCaja(id)`.
* **`📄 Ver Acta Oficial`:** [`BtnVerActa_Click`](../../GestionComercial/Forms/FrmControlCajas.vb#L701) abre `FrmDetalleArqueoCaja(id, seleccionarTabActa:=True)`.
* **`💵 Gastos y Retiros`:** [`BtnVerMovimientos_Click`](../../GestionComercial/Forms/FrmControlCajas.vb#L714) llama a [`CajaService.GetMovimientosCaja(id)`](../../GestionComercial/Services/CajaService.vb#L187) y abre una ventana rápida con retiros e ingresos del turno.
* **`🧾 Tickets de Venta`:** [`BtnVerTickets_Click`](../../GestionComercial/Forms/FrmControlCajas.vb#L776) llama a [`CajaService.GetVentasPorCaja(id)`](../../GestionComercial/Services/CajaService.vb#L303) y lista los comprobantes de ese turno.
* **`📊 Exportar a Excel / CSV`:** [`BtnExportarCSV_Click`](../../GestionComercial/Forms/FrmControlCajas.vb#L835) recorre `listaCajasActual`, abre `SaveFileDialog` y escribe un archivo `.csv` UTF-8 delimitado por `;` con `File.WriteAllText`.

---

## 6. `FrmDetalleArqueoCaja.vb` — Auditoría Profunda de un Turno

Archivo: [GestionComercial/Forms/FrmDetalleArqueoCaja.vb](../../GestionComercial/Forms/FrmDetalleArqueoCaja.vb)

### 🔹 Carga de las 4 pestañas de auditoría
* **Ruta de vida:** `New(cajaId)` → [`InitializeComponent()`](../../GestionComercial/Forms/FrmDetalleArqueoCaja.vb#L62) → [`CargarDatosTurno()`](../../GestionComercial/Forms/FrmDetalleArqueoCaja.vb#L536-L636):
  1. [`CajaService.GetCajaById(_cajaId)`](../../GestionComercial/Services/CajaService.vb#L230) → puebla el desglose matemático en **Pestaña 1 (Conciliación y Arqueo)**: Fondo inicial + Ventas efectivo + Ingresos - Egresos = Esperado vs Real declarado y Diferencia.
  2. [`CajaService.GetResumenMediosPagoCaja(_cajaId)`](../../GestionComercial/Services/CajaService.vb#L346) → llena grilla derecha con totales por medio de pago (`Efectivo`, `Débito`, `Crédito`, `Transferencia / QR`).
  3. [`CajaService.GetVentasPorCaja(_cajaId)`](../../GestionComercial/Services/CajaService.vb#L303) → llena **Pestaña 2 (Tickets de Venta)**.
  4. [`CajaService.GetMovimientosCaja(_cajaId)`](../../GestionComercial/Services/CajaService.vb#L187) → llena **Pestaña 3 (Retiros e Ingresos)** con autorizante.
  5. [`GenerarActaOficial()`](../../GestionComercial/Forms/FrmDetalleArqueoCaja.vb#L638) → consulta [`ConfiguracionService.GetConfiguracion()`](../../GestionComercial/Services/ConfiguracionService.vb#L18) y redacta el acta formal en **Pestaña 4 (Acta de Cierre)** lista para firma.

### 🔹 Botón `🖨 Imprimir / Guardar Acta`
* **Código:** [`BtnImprimirActa_Click`](../../GestionComercial/Forms/FrmDetalleArqueoCaja.vb#L701-L718) abre `SaveFileDialog` y guarda el texto del acta en un archivo `.txt` UTF-8.

---

## 7. `FrmReportes.vb` — Inteligencia Comercial, PDF y Anulaciones

Archivo: [GestionComercial/Forms/FrmReportes.vb](../../GestionComercial/Forms/FrmReportes.vb)

### 🔹 Filtros globales y Botón `Generar Reportes` (`LoadReportes`)
* Carga inicial de selectores: [`CargarComboClientes()`](../../GestionComercial/Forms/FrmReportes.vb#L760) (`ClienteService.GetClientes`) y [`CargarComboVendedores()`](../../GestionComercial/Forms/FrmReportes.vb#L783) (`UsuarioService.GetUsuarios`).
* [`CboRapidoFechas_SelectedIndexChanged`](../../GestionComercial/Forms/FrmReportes.vb#L1191) y `btnFiltrar.Click` ejecutan [`LoadReportes()`](../../GestionComercial/Forms/FrmReportes.vb#L1226), que refresca las 5 pestañas:
  1. **Pestaña `Listado General de Ventas` ([`LoadVentasGenerales`](../../GestionComercial/Forms/FrmReportes.vb#L1235)):**
     * Llama a [`VentaService.GetVentas(...)`](../../GestionComercial/Services/VentaService.vb#L144) aplicando filtros de fechas, vendedor, cliente, medio de pago y estado.
     * Llama a [`ReporteService.GetMetricasFinancierasPeriodo(...)`](../../GestionComercial/Services/ReporteService.vb#L447) para actualizar las 4 tarjetas KPI: **FACTURACIÓN TOTAL**, **COSTO MERCADERÍA**, **GANANCIA NETA (MARGEN)** y **TICKET PROMEDIO**.
  2. **Pestaña `Anulaciones y Devoluciones` ([`LoadAnulaciones`](../../GestionComercial/Forms/FrmReportes.vb#L1287)):**
     * Llama a `VentaService.GetVentas(..., "Anulada")` mostrando quién autorizó la anulación, monto revertido y motivo.
  3. **Pestaña `Ventas por Cliente y Fecha` ([`LoadVentasPorCliente`](../../GestionComercial/Forms/FrmReportes.vb#L1320)):**
     * Llama a [`ReporteService.GetReporteVentasPorCliente(...)`](../../GestionComercial/Services/ReporteService.vb#L221) ordenando por Mayor Facturación, Mayor Salida (unidades) o Mayor Margen.
     * Al seleccionar un cliente en la grilla superior, [`DgvClientesReporte_SelectionChanged`](../../GestionComercial/Forms/FrmReportes.vb#L1440) → [`LoadTicketsCliente()`](../../GestionComercial/Forms/FrmReportes.vb#L1452) consulta `VentaService.GetVentas` para la grilla inferior.
  4. **Pestaña `Ventas por Vendedor y Fecha` ([`LoadVentasPorVendedor`](../../GestionComercial/Forms/FrmReportes.vb#L1365)):**
     * Llama a [`ReporteService.GetReporteVentasPorVendedor(...)`](../../GestionComercial/Services/ReporteService.vb#L317); al seleccionar un vendedor, [`DgvVendedoresReporte_SelectionChanged`](../../GestionComercial/Forms/FrmReportes.vb#L1410) → [`LoadTicketsVendedor()`](../../GestionComercial/Forms/FrmReportes.vb#L1422) muestra sus tickets debajo.
  5. **Pestaña `Ranking de Prendas y Rentabilidad` ([`LoadRanking`](../../GestionComercial/Forms/FrmReportes.vb#L1470)):**
     * Si elige *Mayor Salida* o *Mayor Margen*: reconfigura columnas y llama a [`ReporteService.GetTopProductosVendidos(...)`](../../GestionComercial/Services/ReporteService.vb#L158).
     * Si elige *Stock Inmovilizado*: llama a [`ReporteService.GetStockInmovilizado(...)`](../../GestionComercial/Services/ReporteService.vb#L506) mostrando prendas sin ventas en el período y el capital retenido al costo.

### 🔹 Ver Detalle de Prendas (Doble clic o botones `Ver Detalle de Prendas del Ticket`)
* [`MostrarModalDetalleVenta()`](../../GestionComercial/Forms/FrmReportes.vb#L1661-L1767) consulta [`VentaService.GetVentaPorId(ventaId)`](../../GestionComercial/Services/VentaService.vb#L221) y [`VentaService.GetDetalleVenta(ventaId)`](../../GestionComercial/Services/VentaService.vb#L200) y abre modal con botones de reimpresión PDF.

### 🔹 Botones `📄 Descargar Ticket PDF`
* [`BtnDescargarPdf_Click`](../../GestionComercial/Forms/FrmReportes.vb#L1608), [`BtnDescargarPdfTicketCliente_Click`](../../GestionComercial/Forms/FrmReportes.vb#L1633), [`BtnDescargarPdfTicketVendedor_Click`](../../GestionComercial/Forms/FrmReportes.vb#L1598) → [`DescargarTicketPdfPorId()`](../../GestionComercial/Forms/FrmReportes.vb#L1643) → [`GuardarTicketEnPdf()`](../../GestionComercial/Forms/FrmReportes.vb#L1769).
* **Cómo funciona:** Usa `ConstruirLineasTicket()` (L1849) con datos de `ConfiguracionService.GetConfiguracion()`, verifica la impresora `"Microsoft Print to PDF"`, abre `SaveFileDialog` e imprime vía `System.Drawing.Printing.PrintDocument`.

### 🔹 Botón `Anular Venta Seleccionada` — Ruta de Vida Completa
1. **Validación en UI ([`BtnAnularVenta_Click`](../../GestionComercial/Forms/FrmReportes.vb#L1932-L1988)):**
   * Verifica que sea Gerente (`AuthService.SolicitarAutorizacionManager` si fuera necesario).
   * Comprueba en `dgvVentas` que el estado actual no sea `"Anulada"`.
   * Abre diálogo modal solicitando el **motivo de anulación**.
2. **Transacción Atómica en Servicio ([`VentaService.AnularVenta()`](../../GestionComercial/Services/VentaService.vb#L248-L323)):**
   * `ObtenerVentaRow` (L492): lee la venta en `ventas` y verifica que su `estado` sea `'Completada'`.
   * `ObtenerDetallesVentaEnTransaccion` (L508): lee todas las prendas en `detalle_ventas`.
   * `RestituirStockYAuditar` (L540): devuelve las unidades a `producto_talles` (`stock_actual + cantidad`) e inserta movimiento `'Entrada'` en `movimientos_stock`.
   * `MarcarVentaAnulada` (L583): `UPDATE ventas SET estado = 'Anulada', motivo_anulacion = ..., autorizado_por_id = ...`.
   * `ActualizarCajaAnulacion` (L594): resta los importes de `total_ventas_efectivo` y `total_ventas_digital` en `cajas`.
   * `trans.Commit()` y `AppLogger.LogAudit("VENTA_ANULADA", ...)`.
3. **Resultado en pantalla:** Muestra confirmación y llama a `LoadReportes()` para actualizar ventas, anulaciones y métricas financieras.

---

## 8. `FrmClientes.vb` — Directorio y ABM de Clientes

Archivo: [GestionComercial/Forms/FrmClientes.vb](../../GestionComercial/Forms/FrmClientes.vb)

* **Listado y filtros:** [`LoadClientes()`](../../GestionComercial/Forms/FrmClientes.vb#L177) → `ClienteService.GetClientes(filtro, soloActivos)`.
* **Alta / Edición (`+ Nuevo Cliente`, `✏ Editar`):** Abre `FrmClienteEditor` → [`BtnGuardar_Click`](../../GestionComercial/Forms/FrmClientes.vb#L382) → [`ClienteService.GuardarCliente()`](../../GestionComercial/Services/ClienteService.vb#L112) (`INSERT` o `UPDATE` en `clientes` con validación de DNI único).
* **Baja lógica / Reactivación (`🚫 Activar / Desactivar`):** [`BtnToggleActivo_Click`](../../GestionComercial/Forms/FrmClientes.vb#L241) protege al cliente `Consumidor Final` (`Id = 1`) y llama a `ClienteService.ToggleActivo()`.

---

## 9. `FrmAutorizacionSupervisor.vb` — Autorizaciones Presenciales

Archivo: [GestionComercial/Forms/FrmAutorizacionSupervisor.vb](../../GestionComercial/Forms/FrmAutorizacionSupervisor.vb)

* El Gerente ingresa su contraseña en este diálogo cuando se requiere su aval presencial:
  * **En POS/Caja/Catálogo:** Descuentos mayores al 15% o egresos manuales de caja iniciados por un Vendedor (`AuthService.SolicitarAutorizacionAdminOManager`).
  * **En Backups:** Creación, exportación o restauración de copias de seguridad iniciadas por el Administrador (`AuthService.SolicitarAutorizacionGerente`).
* [`ProcesarConfirmacion()`](../../GestionComercial/Forms/FrmAutorizacionSupervisor.vb#L103) invoca [`AuthService.AutenticarSupervisorPorRoles()`](../../GestionComercial/Services/AuthService.vb#L269), registra `LastAuthorizingSupervisor` y audita el evento en `AppLogger.LogAudit`.

