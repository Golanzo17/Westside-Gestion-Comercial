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
