# 💰 Explicación de CajaService.vb (Apertura, Movimientos y Arqueo)

* **Ubicación**: [GestionComercial/Services/CajaService.vb]
* **Propósito**: Controlar todo el flujo de dinero en efectivo y digital del local: apertura del turno con fondo inicial, registro de gastos e ingresos varios, cálculo del dinero esperado y arqueo al cierre.

---

## 🔍 Explicación de Métodos Principales

### 1. `GetCajaAbierta(usuarioId) As Caja`
* Busca en la base de datos si actualmente existe una fila con `estado = 'Abierta'`.
* Si existe, devuelve el objeto `Caja` con los acumulados en tiempo real. Si no hay ninguna abierta, devuelve `Nothing`.

---

### 2. `AbrirCaja(usuarioId, montoInicial, ByRef errorMessage) As Caja`
* **Regla de negocio**: No se pueden abrir dos cajas en simultáneo. Si ya hay una abierta, rechaza la operación.
* Si no hay ninguna abierta, inserta una nueva fila en `cajas`:
  * `fecha_apertura = NOW()`
  * `monto_inicial = montoInicial` (el dinero en cambio que se deja en la gaveta).
  * `monto_esperado = montoInicial`
  * `estado = 'Abierta'`

---

### 3. `RegistrarMovimiento(cajaId, usuarioId, tipo, concepto, monto, referencia, ByRef errorMessage) As Boolean`
Controla las entradas o salidas de dinero que **no son ventas**:
* Ejemplos:
  * `tipo = "Ingreso"`: Aporte de cambio extra o cobro de una deuda.
  * `tipo = "Egreso"`: Pago a un repartidor, compra de artículos de limpieza, adelanto de sueldo.
* **Transacción de dos pasos**:
  1. Inserta el movimiento en `movimientos_caja`.
  2. Actualiza los acumulados en `cajas`:
     * Si fue Ingreso: suma a `total_ingresos`.
     * Si fue Egreso: suma a `total_egresos`.
     * Recalcula automáticamente el `monto_esperado`.

---

### 4. `CerrarCaja(cajaId, montoReal, observaciones, ByRef errorMessage) As Boolean`
Es el momento del **arqueo de caja** al final del día o cambio de turno:
1. Obtiene los valores registrados de la caja activa:
   $$\text{Monto Esperado} = \text{Inicial} + \text{Ventas Efectivo} + \text{Ingresos} - \text{Egresos}$$
2. El cajero ingresa el dinero físico que contó en la mano (`montoReal`).
3. El sistema calcula la **diferencia**:
   $$\text{Diferencia} = \text{Monto Real} - \text{Monto Esperado}$$
   * Si es `0`: Caja perfecta (sin sobrantes ni faltantes).
   * Si es positivo (`+500`): Hay dinero sobrante.
   * Si es negativo (`-1200`): Hay dinero faltante.
4. Actualiza la fila con `estado = 'Cerrada'`, fecha de cierre y las observaciones escritas por el usuario.

---

### 5. `GetMovimientosCaja` y `GetHistorialCajas`
* `GetMovimientosCaja`: Lista todos los movimientos manuales del turno actual.
* `GetHistorialCajas`: Permite a los administradores auditar cierres de caja de días anteriores filtrando por un rango de fechas.
