# 💳 Explicación de VentaService.vb (Ventas, Stock y Facturación)

* **Ubicación**: [GestionComercial/Services/VentaService.vb]
* **Propósito**: Es el servicio más importante del sistema comercial. Controla el cobro de tickets, garantiza que no se venda mercadería sin stock, registra la auditoría de cada movimiento y gestiona la anulación de tickets.

---

## 🌟 Concepto Clave: ¿Qué es una Transacción Atómica (`BeginTransaction`)?
Cuando se cobra una venta deben ocurrir muchas cosas a la vez:
1. Descontar el stock de cada prenda.
2. Crear los registros de auditoría en `movimientos_stock`.
3. Guardar el ticket en `ventas`.
4. Guardar cada renglón en `detalle_ventas`.
5. Sumar el dinero en la `caja`.

Si la luz se corta o el disco falla en el paso 4, **no podemos dejar la venta a medias**.  
Por eso usamos `trans = conn.BeginTransaction()`. Si todo sale bien, se ejecuta `trans.Commit()` (se guardan todos los cambios juntos). Si algo falla en cualquier paso, se ejecuta `trans.Rollback()` (la base de datos retrocede en el tiempo como si nada hubiera pasado).

---

## 📝 Explicación Función por Función

### 1. `GenerarNumeroTicket() As String`
Genera el número correlativo diario para el ticket:
* Obtiene la fecha actual en formato `yyyyMMdd` (ej: `20260914`).
* Busca en la base de datos el último ticket emitido hoy con el prefijo `TICK-20260914-`.
* Si encuentra el `0003`, le suma 1 para generar `TICK-20260914-0004`.
* Si es la primera venta del día, empieza en `0001`.

---

### 2. `ProcesarVenta(venta As Venta, ByRef errorMessage As String) As Boolean`
Es la función principal de cobro. Su flujo sigue las buenas prácticas de la skill (flujo explícito y plano):

```vb
' Paso 0: Valida que el carrito no esté vacío
If venta.Detalles Is Nothing OrElse venta.Detalles.Count = 0 Then
    errorMessage = "No hay artículos en el carrito de venta."
    Return False
End If
```

Luego, dentro de la transacción:
1. **`GarantizarNumeroTicket`**: Asegura un correlativo único.
2. **`ValidarStockDisponible`**: Revisa prenda por prenda que haya stock suficiente para cada combinación de talle y color. Si falta una sola prenda, cancela la venta antes de descontar nada.
3. **`DescontarStockYAuditar`**: Resta el stock en `producto_talles` e inserta una fila en `movimientos_stock` con `tipo_movimiento = 'Venta'`.
4. **`InsertarVenta`**: Inserta el encabezado (total, vuelto, cliente, forma de pago) y obtiene el nuevo `ID` de venta.
5. **`InsertarDetalles`**: Guarda cada renglón vendido con su precio, costo y cantidad.
6. **`ActualizarCajaVenta`**: Si la venta fue en Efectivo, suma a `total_ventas_efectivo` y al `monto_esperado`; si fue tarjeta/transferencia, suma a `total_ventas_digital`.
7. `trans.Commit()`: Confirma todos los cambios en la base de datos.

---

### 3. `AnularVenta(ventaId, usuarioId, motivo, ByRef errorMessage) As Boolean`
Permite dar de baja un ticket emitido por error:
1. Comprueba que la venta exista y que no haya sido anulada previamente.
2. **`RestituirStockYAuditar`**: Lee los artículos que se habían vendido y le suma nuevamente las cantidades al stock (`stock_actual = stock_actual + cantidad`). Crea un registro de auditoría con tipo `'Anulacion_Venta'`.
3. **`MarcarVentaAnulada`**: Cambia el estado de la venta a `'Anulada'` y agrega el motivo a las observaciones.
4. **`ActualizarCajaAnulacion`**: Resta el importe de la caja correspondiente para que el arqueo de dinero al final del día sea exacto.

---

## 🛡️ Prevención de Concurrencia y Ventas Simultáneas (Race Conditions)

### Escenario Crítico: Dos cajeros vendiendo la última unidad (Stock = 1) al mismo tiempo
¿Qué sucede si el Puesto 1 (POS 1) y el Puesto 2 (POS 2) tienen la misma prenda cargada en su pantalla (con Stock = 1) y ambos cajeros presionan **"Confirmar Cobro" al mismo milisegundo**?

El sistema previene ventas duplicadas o stock negativo mediante 3 barreras de seguridad en el código:

#### 1. Cerrojo de Escritura en Base de Datos (SQLite Locking)
Aunque las dos peticiones lleguen casi al unísono, el motor relacional SQLite opera con bloqueo transaccional. La primera conexión en iniciar la escritura adquiere un cerrojo exclusivo (`RESERVED/EXCLUSIVE lock`). La segunda conexión espera su turno de forma segura.

#### 2. Re-validación Obligatoria Dentro de la Transacción (`ValidarStockDisponible`)
No se confía en la cantidad que el cajero cargó visualmente en su pantalla minutos antes. Al momento de presionar el cobro, la función `ProcesarVenta` consulta **en vivo** dentro de la transacción la cantidad remanente en la tabla `producto_talles`:

```vb
' ARCHIVO: VentaService.vb -> ValidarStockDisponible()
Dim queryStock = "SELECT stock_actual FROM `producto_talles` WHERE `producto_id` = @prodId AND `talle_id` = @talleId AND `color` = @color;"
Dim res = cmdStock.ExecuteScalar()
Dim stockActual = Convert.ToInt32(res)

If stockActual < item.Cantidad Then
    errorMessage = $"Stock insuficiente para '{item.DescripcionArticulo}' (Talle: {item.TalleNombre}). Stock disponible: {stockActual}, solicitado: {item.Cantidad}."
    Return False
End If
```

#### 3. Secuencia Cronológica Real entre POS 1 y POS 2:

* **POS 1 (Llega 1 milisegundo antes)**:
  1. Abre transacción: `conn.BeginTransaction()`.
  2. Consulta `ValidarStockDisponible()`: Lee `stock_actual = 1`. Como `1 >= 1`, valida con éxito.
  3. Ejecuta `DescontarStockYAuditar()`: Actualiza `stock_actual = 1 - 1 = 0`.
  4. Inserta el ticket y confirma con `trans.Commit()`.
  5. **Resultado**: Venta exitosa. El stock en la base de datos ahora es `0`.

* **POS 2 (Llega 1 milisegundo después o tras liberarse el cerrojo)**:
  1. Abre su propia transacción: `conn.BeginTransaction()`.
  2. Consulta `ValidarStockDisponible()`: Lee el stock actualizado en tiempo real, que ahora es **`0`**.
  3. Evalúa la condición `If stockActual < item.Cantidad` (`0 < 1`):
     * Asigna el mensaje de error: `"Stock insuficiente para [Prenda] (Talle: M). Stock disponible: 0, solicitado: 1."`
     * Retorna `False`.
  4. Se ejecuta de inmediato el Rollback:
     ```vb
     If Not ValidarStockDisponible(conn, trans, venta.Detalles, errorMessage) Then
         trans.Rollback()
         Return False
     End If
     ```
  5. **Resultado**: La venta se cancela limpiamente. No se descuenta stock, no se altera la caja y el vendedor recibe una notificación en pantalla impidiendo el cobro.

#### 4. Protección contra Choque Simultáneo (`SQLITE_BUSY`)
Si ambas conexiones intentaran forzar la escritura en el mismo instante exacto sin esperar, SQLite arroja una excepción de base de datos ocupada (`database is locked`). El bloque `Try...Catch` de `ProcesarVenta` captura el error y aplica `trans.Rollback()`:

```vb
Catch exTrans As Exception
    trans.Rollback()
    errorMessage = "Error al procesar la venta: " & exTrans.Message
    Return False
End Try
```

> **Garantía del Sistema**: La consistencia es total. Bajo ninguna circunstancia el stock quedará en números negativos ni existirán dos tickets cobrando la misma unidad física de prenda.
