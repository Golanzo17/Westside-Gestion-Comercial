# 💰 Explicación de CajaService.vb (Apertura, Movimientos y Arqueo)

* **Ubicación**: [GestionComercial/Services/CajaService.vb]
* **Propósito**: Controlar todo el flujo de dinero en efectivo y digital del local: apertura del turno con fondo inicial, registro de gastos e ingresos varios, cálculo del dinero esperado, arqueo al cierre y consultas de auditoría para el Gerente.

---

## 🔍 Operación de la caja (Vendedor)

### 1. `GetCajaAbierta(Optional usuarioId = 0) As Caja`
* Con `usuarioId > 0` busca **únicamente la caja abierta de ese usuario** (`estado = 'Abierta'`). Así un vendedor nunca ve ni opera la caja de otro.
* Sin parámetro (`0`) devuelve la última caja abierta del sistema (uso administrativo).
* Si no hay ninguna, devuelve `Nothing`.

### 2. `AbrirCaja(usuarioId, montoInicial, ByRef errorMessage) As Caja`
* **Regla de negocio**: cada vendedor puede tener **una sola caja abierta a la vez**. Si ya tiene una, rechaza la operación.
* Si no, inserta una fila en `cajas` con `fecha_apertura = NOW()`, `monto_inicial` (el cambio que se deja en la gaveta), `monto_esperado = montoInicial` y `estado = 'Abierta'`. Registra el evento `CAJA_ABIERTA` en auditoría.

### 3. `RegistrarMovimiento(cajaId, usuarioId, tipo, concepto, monto, referencia, ByRef errorMessage, Optional autorizadoPorId) As Boolean`
Controla las entradas o salidas de dinero que **no son ventas**:
* `tipo = "Ingreso"`: aporte de cambio extra o cobro de una deuda.
* `tipo = "Egreso"`: pago de flete, compra de insumos, retiro para depósito.
* **Transacción de dos pasos**:
  1. Inserta el movimiento en `movimientos_caja`, guardando también `autorizado_por_id` (el supervisor que aprobó un retiro).
  2. Actualiza `total_ingresos` o `total_egresos` en `cajas` y recalcula el `monto_esperado`.

### 4. `CerrarCaja(cajaId, montoReal, observaciones, ByRef errorMessage) As Boolean`
Es el momento del **arqueo de caja** al final del turno:
$$\text{Monto Esperado} = \text{Inicial} + \text{Ventas Efectivo} + \text{Ingresos} - \text{Egresos}$$
$$\text{Diferencia} = \text{Monto Real} - \text{Monto Esperado}$$
* `0`: caja perfecta. Positivo: sobrante. Negativo: faltante.
* Actualiza la fila con `estado = 'Cerrada'`, fecha de cierre, monto real, diferencia y observaciones.
* El vendedor hace un **arqueo ciego**: cuenta el efectivo sin ver el monto esperado.

### 5. `GetMovimientosCaja(cajaId) As List(Of MovimientoCaja)`
* Lista los ingresos y egresos manuales de un turno, con el nombre de quien los hizo.

---

## 🕵️ Consultas de auditoría (Gerente)

Las usa la pantalla **Cierres y arqueos** (`FrmControlCajas`) y el detalle del arqueo (`FrmDetalleArqueoCaja`).

### 6. `GetCajaById(id) As Caja`
* Devuelve un turno puntual (abierto o cerrado) con el nombre del vendedor.

### 7. `GetHistorialCajasAvanzado(fechaDesde, fechaHasta, usuarioId, estado, filtroDiferencia) As List(Of Caja)`
* Construye dinámicamente el SQL con los filtros elegidos:
  * Rango de fechas (por fecha de apertura).
  * Vendedor (`usuarioId`, opcional).
  * Estado: `"Abierta"` o `"Cerrada"` (o ninguno para todas).
  * Discrepancia: `CON_DESCUADRE`, `FALTANTE`, `SOBRANTE` o `EXACTO`.

### 8. `GetVentasPorCaja(cajaId) As List(Of Venta)`
* Todos los tickets emitidos en un turno, de más reciente a más antiguo. Se usa en "Ver tickets" del control de cajas.

### 9. `GetResumenMediosPagoCaja(cajaId) As DataTable`
* Agrupa las ventas no anuladas del turno por medio de pago: cantidad de tickets y total cobrado.

### 10. `GetDesgloseMediosTurno(cajaId) As DesgloseTurnoCaja`
* Calcula el desglose real del turno: **Efectivo**, **Tarjetas** y **Transferencia / QR**, más cantidad de tickets y total vendido.
* Usa `LIKE 'Tarjeta%'` y `LIKE 'Transferencia%'` porque el POS guarda `"Transferencia / QR"`, `"Tarjeta Débito"` y `"Tarjeta Crédito"`: una comparación exacta daría 0.
* En pagos **Múltiple** reparte el total: lo que entró en efectivo suma a Efectivo y el resto a Transferencia.
* Alimenta los KPIs de `FrmCaja` y de `FrmDetalleArqueoCaja`.
