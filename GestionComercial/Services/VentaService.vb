' ARCHIVO: VentaService.vb
' PROPÓSITO: Lógica de negocio de cobros, transacciones atómicas y control de inventario.
' Este es el servicio central del circuito comercial:
' 1. Transacciones ACID (DbTransaction): Todo el cobro se ejecuta como una sola unidad atómica.
'    Si falla el descuento de stock de una sola prenda, se hace Rollback y no se cobra nada.
' 2. Formato de Comprobantes Único: Los tickets se generan con numeración cronológica
'    "TICK-YYYYMMDD-0001", garantizando trazabilidad contable diaria.
' 3. Auditoría de Stock en Vivo: No solo restamos del inventario, sino que insertamos un
'    registro en movimientos_stock asociando el ticket y el vendedor que hizo la operación.
' 4. Impacto Inmediato en Caja: Actualiza en tiempo real el efectivo o dinero digital
'    esperado en el turno abierto de caja (incluyendo desglose en pagos mixtos).

Imports System.Data
Imports System.Data.Common
Imports GestionComercial.Data
Imports GestionComercial.Models

Namespace Services
    Public Class VentaService

        ' Genera el número secuencial de ticket para el día de hoy (ej: TICK-20260922-0001).
        ' Busca el último emitido hoy, extrae el número y le suma 1 con formato de 4 dígitos.
        Public Function GenerarNumeroTicket(Optional conn As DbConnection = Nothing, Optional trans As DbTransaction = Nothing) As String
            Dim hoyStr = DateTime.Now.ToString("yyyyMMdd")
            Dim prefijo = "TICK-" & hoyStr & "-"
            Dim query = "SELECT `numero_ticket` FROM `ventas` WHERE `numero_ticket` LIKE @prefijo ORDER BY `id` DESC LIMIT 1;"

            Dim lastTicket As String = Nothing
            If conn IsNot Nothing Then
                Using cmd = DatabaseHelper.CreateCommand(conn, query, trans)
                    DatabaseHelper.AddParam(cmd, "@prefijo", prefijo & "%")
                    Dim obj = cmd.ExecuteScalar()
                    If obj IsNot Nothing AndAlso Not Convert.IsDBNull(obj) Then
                        lastTicket = obj.ToString()
                    End If
                End Using
            Else
                Dim params = New Dictionary(Of String, Object) From {{"@prefijo", prefijo & "%"}}
                Dim dt = DatabaseHelper.ExecuteQuery(query, params)
                If dt.Rows.Count > 0 Then
                    lastTicket = dt.Rows(0)("numero_ticket").ToString()
                End If
            End If

            Dim nextSeq = 1
            If Not String.IsNullOrEmpty(lastTicket) Then
                Dim parts = lastTicket.Split("-"c)
                If parts.Length >= 3 Then
                    Dim currentSeq As Integer
                    If Integer.TryParse(parts(2), currentSeq) Then
                        nextSeq = currentSeq + 1
                    End If
                End If
            End If

            Return prefijo & nextSeq.ToString("D4")
        End Function

        ' Procesamiento de Venta con Transacción ACID.
        ' Flujo seguro de 5 pasos:
        '   Paso 0: Garantizamos que el número de ticket no colisione con otro usuario.
        '   Paso 1: Validamos que haya stock físico suficiente de cada talle y color.
        '   Paso 2: Descontamos las unidades y registramos el movimiento en la auditoría.
        '   Paso 3: Insertamos la cabecera del ticket en la tabla ventas.
        '   Paso 4: Insertamos cada artículo vendido en detalle_ventas.
        '   Paso 5: Sumamos el importe a la caja activa del turno (Efectivo y/o Digital).
        ' Si ocurre cualquier error en el camino, se ejecuta trans.Rollback() y no se guarda nada.
        Public Function ProcesarVenta(venta As Venta, ByRef errorMessage As String) As Boolean
            If venta Is Nothing Then
                errorMessage = "La información de la venta no es válida."
                Return False
            End If

            If venta.Detalles Is Nothing OrElse venta.Detalles.Count = 0 Then
                errorMessage = "No hay artículos en el carrito de venta."
                Return False
            End If

            Try
                Using conn = DatabaseHelper.GetConnection()
                    conn.Open()
                    Using trans = conn.BeginTransaction()
                        Try
                            ' Normalizamos redondeo contable a 2 decimales
                            venta.Subtotal = Math.Round(venta.Subtotal, 2)
                            venta.DescuentoMonto = Math.Round(venta.DescuentoMonto, 2)
                            venta.RecargoMonto = Math.Round(venta.RecargoMonto, 2)
                            venta.Total = Math.Round(venta.Total, 2)
                            venta.MontoAbonado = Math.Round(venta.MontoAbonado, 2)
                            venta.Vuelto = Math.Round(venta.Vuelto, 2)

                            Dim montoEfectivo = CalcularMontoEfectivoVenta(venta)
                            venta.MontoEfectivo = montoEfectivo

                            ' Paso 0: Verificamos o generamos el correlativo único del comprobante
                            GarantizarNumeroTicket(conn, trans, venta)

                            ' Paso 1: Verificamos stock físico disponible para cada variante elegida
                            If Not ValidarStockDisponible(conn, trans, venta.Detalles, errorMessage) Then
                                trans.Rollback()
                                Return False
                            End If

                            ' Paso 2: Descontamos stock e insertamos los registros en movimientos_stock (Kardex)
                            DescontarStockYAuditar(conn, trans, venta.Detalles, venta.NumeroTicket, venta.UsuarioId)

                            ' Paso 3: Guardamos la cabecera del ticket en la tabla ventas
                            Dim ventaId = InsertarVenta(conn, trans, venta)
                            venta.Id = ventaId

                            ' Paso 4: Guardamos los renglones comprados en detalle_ventas
                            InsertarDetalles(conn, trans, ventaId, venta.Detalles)

                            ' Paso 5: Impactamos el dinero cobrado en la caja activa del turno
                            If venta.CajaId > 0 Then
                                Dim montoDigital = Math.Round(venta.Total - montoEfectivo, 2)
                                ActualizarCajaVenta(conn, trans, venta.CajaId, montoEfectivo, montoDigital)
                            End If

                            ' Si todo salió perfecto, confirmamos la transacción de forma atómica
                            trans.Commit()
                            AppLogger.LogAudit("VENTA_PROCESADA", $"Ticket={venta.NumeroTicket} | Total={venta.Total:F2} | UsuarioId={venta.UsuarioId}")
                            errorMessage = String.Empty
                            Return True
                        Catch exTrans As Exception
                            trans.Rollback()
                            AppLogger.LogError("Error en transacción ProcesarVenta", exTrans)
                            errorMessage = "Error al procesar la venta: " & exTrans.Message
                            Return False
                        End Try
                    End Using
                End Using
            Catch ex As Exception
                AppLogger.LogError("Error de conexión en ProcesarVenta", ex)
                errorMessage = "Error de conexión con la base de datos: " & ex.Message
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Busca y filtra el historial de ventas por rango de fechas, texto (ticket/cliente/DNI), cliente, vendedor, medio de pago y estado.
        ''' Es el método principal que alimenta la grilla del módulo de Reportes.
        ''' </summary>
        Public Function GetVentas(fechaDesde As DateTime, fechaHasta As DateTime, Optional busqueda As String = "", Optional clienteId As Integer = 0, Optional usuarioId As Integer = 0, Optional metodoPago As String = "", Optional estado As String = "") As List(Of Venta)
            Dim list = New List(Of Venta)()
            Dim query = "SELECT v.*, u.nombre_completo AS usuario_nombre, IFNULL(ua.nombre_completo, '') AS autorizado_por_nombre, (c.apellido || ', ' || c.nombre) AS cliente_nombre " &
                        "FROM `ventas` v " &
                        "INNER JOIN `usuarios` u ON v.usuario_id = u.id " &
                        "LEFT JOIN `usuarios` ua ON v.autorizado_por_id = ua.id " &
                        "INNER JOIN `clientes` c ON v.cliente_id = c.id " &
                        "WHERE DATE(v.fecha) >= @desde AND DATE(v.fecha) <= @hasta "

            Dim params = New Dictionary(Of String, Object) From {
                {"@desde", fechaDesde.ToString("yyyy-MM-dd")},
                {"@hasta", fechaHasta.ToString("yyyy-MM-dd")}
            }

            If clienteId > 0 Then
                query &= "AND v.cliente_id = @cliId "
                params.Add("@cliId", clienteId)
            End If

            If usuarioId > 0 Then
                query &= "AND v.usuario_id = @usrId "
                params.Add("@usrId", usuarioId)
            End If

            If Not String.IsNullOrWhiteSpace(metodoPago) Then
                If metodoPago.Trim().StartsWith("Transferencia", StringComparison.OrdinalIgnoreCase) Then
                    query &= "AND v.metodo_pago LIKE 'Transferencia%' "
                Else
                    query &= "AND v.metodo_pago = @metodo "
                    params.Add("@metodo", metodoPago.Trim())
                End If
            End If

            If Not String.IsNullOrWhiteSpace(estado) Then
                query &= "AND v.estado = @estado "
                params.Add("@estado", estado.Trim())
            End If

            Dim hasBusqueda = Not String.IsNullOrWhiteSpace(busqueda)
            If hasBusqueda Then
                query &= "AND (v.numero_ticket LIKE @b OR c.nombre LIKE @b OR c.apellido LIKE @b OR c.dni_cuit LIKE @b) "
                params.Add("@b", "%" & busqueda.Trim() & "%")
            End If

            query &= "ORDER BY v.fecha DESC;"

            Dim dt = DatabaseHelper.ExecuteQuery(query, params)
            For Each row As DataRow In dt.Rows
                list.Add(MapVenta(row))
            Next
            Return list
        End Function

        ''' <summary>
        ''' Devuelve la lista de renglones/prendas vendidas que pertenecen a un ticket específico.
        ''' </summary>
        Public Function GetDetalleVenta(ventaId As Integer) As List(Of DetalleVenta)
            Dim list = New List(Of DetalleVenta)()
            If ventaId <= 0 Then
                Return list
            End If

            Dim query = "SELECT dv.*, t.nombre AS talle_nombre " &
                        "FROM `detalle_ventas` dv " &
                        "INNER JOIN `talles` t ON dv.talle_id = t.id " &
                        "WHERE dv.venta_id = @vId;"

            Dim dt = DatabaseHelper.ExecuteQuery(query, New Dictionary(Of String, Object) From {{"@vId", ventaId}})
            For Each row As DataRow In dt.Rows
                list.Add(MapDetalleVenta(row))
            Next
            Return list
        End Function

        ''' <summary>
        ''' Obtiene la cabecera completa de una venta por su ID (incluyendo nombres de vendedor, cliente y supervisor autorizante).
        ''' </summary>
        Public Function GetVentaPorId(ventaId As Integer) As Venta
            If ventaId <= 0 Then
                Return Nothing
            End If

            Dim query = "SELECT v.*, u.nombre_completo AS usuario_nombre, IFNULL(ua.nombre_completo, '') AS autorizado_por_nombre, (c.apellido || ', ' || c.nombre) AS cliente_nombre " &
                        "FROM `ventas` v " &
                        "INNER JOIN `usuarios` u ON v.usuario_id = u.id " &
                        "LEFT JOIN `usuarios` ua ON v.autorizado_por_id = ua.id " &
                        "INNER JOIN `clientes` c ON v.cliente_id = c.id " &
                        "WHERE v.id = @vId LIMIT 1;"
            Dim params = New Dictionary(Of String, Object) From {{"@vId", ventaId}}
            Dim dt = DatabaseHelper.ExecuteQuery(query, params)
            If dt.Rows.Count = 0 Then
                Return Nothing
            End If

            Return MapVenta(dt.Rows(0))
        End Function

        ' Anulación Segura de Venta con Reversa Transaccional.
        ' Si un cliente devuelve prendas o hubo un error de cobro:
        ' 1. Verificamos que la venta exista y no haya sido anulada previamente.
        ' 2. Restituimos el stock físico sumando las unidades a producto_talles.
        ' 3. Auditamos en movimientos_stock con stock_anterior y stock_posterior reales.
        ' 4. Marcamos el estado del ticket como "Anulada".
        ' 5. Reversamos el monto en la caja para que el arqueo final coincida con el dinero real.
        Public Function AnularVenta(ventaId As Integer, usuarioId As Integer, motivo As String, ByRef errorMessage As String) As Boolean
            If ventaId <= 0 Then
                errorMessage = "Identificador de venta inválido."
                Return False
            End If

            Try
                Using conn = DatabaseHelper.GetConnection()
                    conn.Open()
                    Using trans = conn.BeginTransaction()
                        Try
                            Dim ventaRow = ObtenerVentaRow(conn, trans, ventaId)
                            If ventaRow Is Nothing Then
                                trans.Rollback()
                                errorMessage = "Venta no encontrada."
                                Return False
                            End If

                            Dim yaAnulada = ventaRow("estado").ToString().Equals("Anulada", StringComparison.OrdinalIgnoreCase)
                            If yaAnulada Then
                                trans.Rollback()
                                errorMessage = "Esta venta ya se encuentra anulada."
                                Return False
                            End If

                            ' Paso 1: Restituir stock físico de cada artículo dentro de la misma transacción y auditar en movimientos_stock
                            Dim detalles = ObtenerDetallesVentaEnTransaccion(conn, trans, ventaId)
                            Dim ticket = ventaRow("numero_ticket").ToString()
                            RestituirStockYAuditar(conn, trans, detalles, ticket, motivo, usuarioId)

                            ' Paso 2: Marcar comprobante como Anulado
                            MarcarVentaAnulada(conn, trans, ventaId, motivo, usuarioId)

                            ' Paso 3: Restar el monto de la caja activa para mantener el balance contable
                            Dim cajaId = Convert.ToInt32(ventaRow("caja_id"))
                            Dim total = Math.Round(Convert.ToDecimal(ventaRow("total")), 2)
                            Dim metodo = ventaRow("metodo_pago").ToString()
                            Dim montoEfecGuardado As Nullable(Of Decimal) = Nothing
                            If ventaRow.Table.Columns.Contains("monto_efectivo") AndAlso Not Convert.IsDBNull(ventaRow("monto_efectivo")) Then
                                montoEfecGuardado = Math.Round(Convert.ToDecimal(ventaRow("monto_efectivo")), 2)
                            End If

                            If cajaId > 0 Then
                                Dim efecReversa As Decimal
                                If montoEfecGuardado.HasValue Then
                                    efecReversa = montoEfecGuardado.Value
                                ElseIf metodo.Equals("Efectivo", StringComparison.OrdinalIgnoreCase) Then
                                    efecReversa = total
                                Else
                                    efecReversa = 0D
                                End If
                                Dim digReversa = Math.Round(total - efecReversa, 2)
                                ActualizarCajaAnulacion(conn, trans, cajaId, efecReversa, digReversa)
                            End If

                            trans.Commit()
                            AppLogger.LogAudit("VENTA_ANULADA", $"Ticket={ticket} | VentaId={ventaId} | AutorizadoPor={usuarioId} | Motivo={motivo}")
                            errorMessage = String.Empty
                            Return True
                        Catch exTrans As Exception
                            trans.Rollback()
                            AppLogger.LogError("Error en transacción AnularVenta", exTrans)
                            errorMessage = "Error al anular la venta: " & exTrans.Message
                            Return False
                        End Try
                    End Using
                End Using
            Catch ex As Exception
                AppLogger.LogError("Error de conexión en AnularVenta", ex)
                errorMessage = "Error de conexión: " & ex.Message
                Return False
            End Try
        End Function

        ' --- Helpers privados de soporte ---

        ' Determina cuánto dinero de la venta ingresó efectivamente en billetes (total si es Efectivo, o la parte en efectivo si es pago Múltiple)
        Private Shared Function CalcularMontoEfectivoVenta(venta As Venta) As Decimal
            If venta.MontoEfectivo.HasValue Then
                Return Math.Clamp(Math.Round(venta.MontoEfectivo.Value, 2), 0D, venta.Total)
            End If

            Dim metodo = If(venta.MetodoPago, "Efectivo").Trim()
            If metodo.Equals("Efectivo", StringComparison.OrdinalIgnoreCase) Then
                Return venta.Total
            End If

            If metodo.Equals("Múltiple", StringComparison.OrdinalIgnoreCase) Then
                Return Math.Clamp(Math.Round(venta.MontoAbonado, 2), 0D, venta.Total)
            End If

            Return 0D
        End Function

        ' Asigna un número de ticket nuevo si no tiene uno o si el correlativo ya fue usado por otra caja simultánea
        Private Sub GarantizarNumeroTicket(conn As DbConnection, trans As DbTransaction, venta As Venta)
            If String.IsNullOrWhiteSpace(venta.NumeroTicket) Then
                venta.NumeroTicket = GenerarNumeroTicket(conn, trans)
                Return
            End If

            Dim queryCheckTicket = "SELECT COUNT(*) FROM `ventas` WHERE `numero_ticket` = @num;"
            Using cmdCheck = DatabaseHelper.CreateCommand(conn, queryCheckTicket, trans)
                DatabaseHelper.AddParam(cmdCheck, "@num", venta.NumeroTicket)
                Dim countTicket = Convert.ToInt64(cmdCheck.ExecuteScalar())
                If countTicket > 0 Then
                    venta.NumeroTicket = GenerarNumeroTicket(conn, trans)
                End If
            End Using
        End Sub

        ' Verifica dentro de la transacción que cada prenda (talle + color) tenga stock suficiente antes de cobrar
        Private Function ValidarStockDisponible(conn As DbConnection, trans As DbTransaction, detalles As List(Of DetalleVenta), ByRef errorMessage As String) As Boolean
            Dim queryStock = "SELECT stock_actual FROM `producto_talles` WHERE `producto_id` = @prodId AND `talle_id` = @talleId AND `color` = @color;"
            For Each item In detalles
                Using cmdStock = DatabaseHelper.CreateCommand(conn, queryStock, trans)
                    DatabaseHelper.AddParam(cmdStock, "@prodId", item.ProductoId)
                    DatabaseHelper.AddParam(cmdStock, "@talleId", item.TalleId)
                    DatabaseHelper.AddParam(cmdStock, "@color", item.Color)
                    Dim res = cmdStock.ExecuteScalar()

                    If res Is Nothing OrElse Convert.IsDBNull(res) Then
                        errorMessage = $"El artículo '{item.DescripcionArticulo}' (Talle: {item.TalleNombre}, Color: {item.Color}) no tiene stock asignado."
                        Return False
                    End If

                    Dim stockActual = Convert.ToInt32(res)
                    If stockActual < item.Cantidad Then
                        errorMessage = $"Stock insuficiente para '{item.DescripcionArticulo}' (Talle: {item.TalleNombre}). Stock disponible: {stockActual}, solicitado: {item.Cantidad}."
                        Return False
                    End If
                End Using
            Next
            Return True
        End Function

        ' Resta las unidades vendidas en producto_talles y guarda el movimiento de salida en movimientos_stock
        Private Sub DescontarStockYAuditar(conn As DbConnection, trans As DbTransaction, detalles As List(Of DetalleVenta), numeroTicket As String, usuarioId As Integer)
            Dim querySelect = "SELECT stock_actual FROM `producto_talles` WHERE `producto_id` = @prodId AND `talle_id` = @talleId AND `color` = @color;"
            Dim queryDescontar = "UPDATE `producto_talles` SET `stock_actual` = @nuevoStock WHERE `producto_id` = @prodId AND `talle_id` = @talleId AND `color` = @color;"
            Dim queryMovStock = "INSERT INTO `movimientos_stock` (`producto_id`, `talle_id`, `color`, `tipo_movimiento`, `cantidad`, `stock_anterior`, `stock_posterior`, `motivo`, `usuario_id`) " &
                                "VALUES (@prodId, @talleId, @color, 'Venta', @cant, @ant, @post, @motivo, @userId);"

            For Each item In detalles
                Dim stockActual = 0
                Using cmdSel = DatabaseHelper.CreateCommand(conn, querySelect, trans)
                    DatabaseHelper.AddParam(cmdSel, "@prodId", item.ProductoId)
                    DatabaseHelper.AddParam(cmdSel, "@talleId", item.TalleId)
                    DatabaseHelper.AddParam(cmdSel, "@color", item.Color)
                    stockActual = Convert.ToInt32(cmdSel.ExecuteScalar())
                End Using

                Dim nuevoStock = stockActual - item.Cantidad
                Using cmdDesc = DatabaseHelper.CreateCommand(conn, queryDescontar, trans)
                    DatabaseHelper.AddParam(cmdDesc, "@nuevoStock", nuevoStock)
                    DatabaseHelper.AddParam(cmdDesc, "@prodId", item.ProductoId)
                    DatabaseHelper.AddParam(cmdDesc, "@talleId", item.TalleId)
                    DatabaseHelper.AddParam(cmdDesc, "@color", item.Color)
                    cmdDesc.ExecuteNonQuery()
                End Using

                Using cmdMov = DatabaseHelper.CreateCommand(conn, queryMovStock, trans)
                    DatabaseHelper.AddParam(cmdMov, "@prodId", item.ProductoId)
                    DatabaseHelper.AddParam(cmdMov, "@talleId", item.TalleId)
                    DatabaseHelper.AddParam(cmdMov, "@color", item.Color)
                    DatabaseHelper.AddParam(cmdMov, "@cant", item.Cantidad)
                    DatabaseHelper.AddParam(cmdMov, "@ant", stockActual)
                    DatabaseHelper.AddParam(cmdMov, "@post", nuevoStock)
                    DatabaseHelper.AddParam(cmdMov, "@motivo", "Venta Ticket #" & numeroTicket)
                    DatabaseHelper.AddParam(cmdMov, "@userId", usuarioId)
                    cmdMov.ExecuteNonQuery()
                End Using
            Next
        End Sub

        ' Inserta la cabecera de la venta en la tabla 'ventas' y retorna el ID recién generado
        Private Function InsertarVenta(conn As DbConnection, trans As DbTransaction, venta As Venta) As Integer
            Dim queryVenta = "INSERT INTO `ventas` (`numero_ticket`, `fecha`, `usuario_id`, `autorizado_por_id`, `cliente_id`, `caja_id`, " &
                             "`tipo_comprobante`, `metodo_pago`, `subtotal`, `descuento_porcentaje`, `descuento_monto`, " &
                             "`recargo_monto`, `total`, `monto_efectivo`, `monto_abonado`, `vuelto`, `estado`, `observaciones`) " &
                             "VALUES (@num, NOW(), @userId, @autId, @cliId, @cajaId, @tipoComp, @metodo, @sub, @descPorc, @descMonto, @recargo, @total, @montoEfec, @abonado, @vuelto, 'Completada', @obs);"

            Using cmdVenta = DatabaseHelper.CreateCommand(conn, queryVenta, trans)
                DatabaseHelper.AddParam(cmdVenta, "@num", venta.NumeroTicket)
                DatabaseHelper.AddParam(cmdVenta, "@userId", venta.UsuarioId)
                DatabaseHelper.AddParam(cmdVenta, "@autId", If(venta.AutorizadoPorId.HasValue, CObj(venta.AutorizadoPorId.Value), DBNull.Value))
                DatabaseHelper.AddParam(cmdVenta, "@cliId", venta.ClienteId)
                DatabaseHelper.AddParam(cmdVenta, "@cajaId", venta.CajaId)
                DatabaseHelper.AddParam(cmdVenta, "@tipoComp", venta.TipoComprobante)
                DatabaseHelper.AddParam(cmdVenta, "@metodo", venta.MetodoPago)
                DatabaseHelper.AddParam(cmdVenta, "@sub", venta.Subtotal)
                DatabaseHelper.AddParam(cmdVenta, "@descPorc", venta.DescuentoPorcentaje)
                DatabaseHelper.AddParam(cmdVenta, "@descMonto", venta.DescuentoMonto)
                DatabaseHelper.AddParam(cmdVenta, "@recargo", venta.RecargoMonto)
                DatabaseHelper.AddParam(cmdVenta, "@total", venta.Total)
                DatabaseHelper.AddParam(cmdVenta, "@montoEfec", If(venta.MontoEfectivo.HasValue, CObj(venta.MontoEfectivo.Value), DBNull.Value))
                DatabaseHelper.AddParam(cmdVenta, "@abonado", venta.MontoAbonado)
                DatabaseHelper.AddParam(cmdVenta, "@vuelto", venta.Vuelto)
                DatabaseHelper.AddParam(cmdVenta, "@obs", venta.Observaciones)
                cmdVenta.ExecuteNonQuery()
            End Using

            Return Convert.ToInt32(DatabaseHelper.GetLastInsertedId(conn, trans))
        End Function

        ' Guarda cada ítem/prenda del carrito en la tabla 'detalle_ventas' asociado a la venta
        Private Sub InsertarDetalles(conn As DbConnection, trans As DbTransaction, ventaId As Integer, detalles As List(Of DetalleVenta))
            Dim queryDetalle = "INSERT INTO `detalle_ventas` (`venta_id`, `producto_id`, `talle_id`, `color`, `codigo_barra`, `descripcion_articulo`, `precio_unitario`, `costo_unitario`, `cantidad`, `subtotal`) " &
                               "VALUES (@vId, @pId, @tId, @color, @cod, @desc, @pu, @cu, @cant, @sub);"

            For Each d In detalles
                Using cmdDet = DatabaseHelper.CreateCommand(conn, queryDetalle, trans)
                    DatabaseHelper.AddParam(cmdDet, "@vId", ventaId)
                    DatabaseHelper.AddParam(cmdDet, "@pId", d.ProductoId)
                    DatabaseHelper.AddParam(cmdDet, "@tId", d.TalleId)
                    DatabaseHelper.AddParam(cmdDet, "@color", d.Color)
                    DatabaseHelper.AddParam(cmdDet, "@cod", d.CodigoBarra)
                    DatabaseHelper.AddParam(cmdDet, "@desc", d.DescripcionArticulo)
                    DatabaseHelper.AddParam(cmdDet, "@pu", Math.Round(d.PrecioUnitario, 2))
                    DatabaseHelper.AddParam(cmdDet, "@cu", Math.Round(d.CostoUnitario, 2))
                    DatabaseHelper.AddParam(cmdDet, "@cant", d.Cantidad)
                    DatabaseHelper.AddParam(cmdDet, "@sub", Math.Round(d.Subtotal, 2))
                    cmdDet.ExecuteNonQuery()
                End Using
            Next
        End Sub

        ' Suma a la caja del turno los importes cobrados en efectivo y en medios digitales, y actualiza el esperado en cajón
        Private Sub ActualizarCajaVenta(conn As DbConnection, trans As DbTransaction, cajaId As Integer, montoEfectivo As Decimal, montoDigital As Decimal)
            Dim updateCajaQuery = "UPDATE `cajas` SET " &
                                  "`total_ventas_efectivo` = ROUND(`total_ventas_efectivo` + @efec, 2), " &
                                  "`total_ventas_digital` = ROUND(`total_ventas_digital` + @dig, 2), " &
                                  "`monto_esperado` = ROUND(`monto_esperado` + @efec, 2) " &
                                  "WHERE `id` = @cajaId;"

            Using cmdUpdCaja = DatabaseHelper.CreateCommand(conn, updateCajaQuery, trans)
                DatabaseHelper.AddParam(cmdUpdCaja, "@efec", montoEfectivo)
                DatabaseHelper.AddParam(cmdUpdCaja, "@dig", montoDigital)
                DatabaseHelper.AddParam(cmdUpdCaja, "@cajaId", cajaId)
                cmdUpdCaja.ExecuteNonQuery()
            End Using
        End Sub

        ' Lee los datos crudos de la fila de una venta dentro de la transacción activa (para anulación)
        Private Function ObtenerVentaRow(conn As DbConnection, trans As DbTransaction, ventaId As Integer) As DataRow
            Dim queryVenta = "SELECT * FROM `ventas` WHERE `id` = @vId;"
            Using cmdVenta = DatabaseHelper.CreateCommand(conn, queryVenta, trans)
                DatabaseHelper.AddParam(cmdVenta, "@vId", ventaId)
                Using reader = cmdVenta.ExecuteReader()
                    Dim dt = New DataTable()
                    dt.Load(reader)
                    If dt.Rows.Count = 0 Then
                        Return Nothing
                    End If
                    Return dt.Rows(0)
                End Using
            End Using
        End Function

        ' Recupera los ítems vendidos dentro de la misma transacción de anulación para saber qué cantidades devolver al stock
        Private Function ObtenerDetallesVentaEnTransaccion(conn As DbConnection, trans As DbTransaction, ventaId As Integer) As List(Of DetalleVenta)
            Dim list = New List(Of DetalleVenta)()
            Dim query = "SELECT dv.*, t.nombre AS talle_nombre " &
                        "FROM `detalle_ventas` dv " &
                        "INNER JOIN `talles` t ON dv.talle_id = t.id " &
                        "WHERE dv.venta_id = @vId;"

            Using cmd = DatabaseHelper.CreateCommand(conn, query, trans)
                DatabaseHelper.AddParam(cmd, "@vId", ventaId)
                Using reader = cmd.ExecuteReader()
                    Dim dt = New DataTable()
                    For i As Integer = 0 To reader.FieldCount - 1
                        Dim colType = reader.GetFieldType(i)
                        dt.Columns.Add(reader.GetName(i), If(colType, GetType(Object)))
                    Next
                    While reader.Read()
                        Dim row = dt.NewRow()
                        For i As Integer = 0 To reader.FieldCount - 1
                            row(i) = reader.GetValue(i)
                        Next
                        dt.Rows.Add(row)
                    End While
                    For Each r As DataRow In dt.Rows
                        list.Add(MapDetalleVenta(r))
                    Next
                End Using
            End Using

            Return list
        End Function

        ' Devuelve al inventario las prendas de una venta anulada y registra el ingreso en movimientos_stock
        Private Sub RestituirStockYAuditar(conn As DbConnection, trans As DbTransaction, detalles As List(Of DetalleVenta), numeroTicket As String, motivo As String, usuarioId As Integer)
            Dim querySelect = "SELECT stock_actual FROM `producto_talles` WHERE `producto_id` = @pId AND `talle_id` = @tId AND `color` = @color;"
            Dim queryStock = "UPDATE `producto_talles` SET `stock_actual` = @nuevoStock WHERE `producto_id` = @pId AND `talle_id` = @tId AND `color` = @color;"
            Dim queryAud = "INSERT INTO `movimientos_stock` (`producto_id`, `talle_id`, `color`, `tipo_movimiento`, `cantidad`, `stock_anterior`, `stock_posterior`, `motivo`, `usuario_id`) " &
                           "VALUES (@pId, @tId, @color, 'Anulacion_Venta', @cant, @ant, @post, @motivo, @uId);"

            For Each item In detalles
                Dim stockActual = 0
                Using cmdSel = DatabaseHelper.CreateCommand(conn, querySelect, trans)
                    DatabaseHelper.AddParam(cmdSel, "@pId", item.ProductoId)
                    DatabaseHelper.AddParam(cmdSel, "@tId", item.TalleId)
                    DatabaseHelper.AddParam(cmdSel, "@color", item.Color)
                    Dim val = cmdSel.ExecuteScalar()
                    If val IsNot Nothing AndAlso Not Convert.IsDBNull(val) Then
                        stockActual = Convert.ToInt32(val)
                    End If
                End Using

                Dim nuevoStock = stockActual + item.Cantidad

                Using cmdStock = DatabaseHelper.CreateCommand(conn, queryStock, trans)
                    DatabaseHelper.AddParam(cmdStock, "@nuevoStock", nuevoStock)
                    DatabaseHelper.AddParam(cmdStock, "@pId", item.ProductoId)
                    DatabaseHelper.AddParam(cmdStock, "@tId", item.TalleId)
                    DatabaseHelper.AddParam(cmdStock, "@color", item.Color)
                    cmdStock.ExecuteNonQuery()
                End Using

                Using cmdAud = DatabaseHelper.CreateCommand(conn, queryAud, trans)
                    DatabaseHelper.AddParam(cmdAud, "@pId", item.ProductoId)
                    DatabaseHelper.AddParam(cmdAud, "@tId", item.TalleId)
                    DatabaseHelper.AddParam(cmdAud, "@color", item.Color)
                    DatabaseHelper.AddParam(cmdAud, "@cant", item.Cantidad)
                    DatabaseHelper.AddParam(cmdAud, "@ant", stockActual)
                    DatabaseHelper.AddParam(cmdAud, "@post", nuevoStock)
                    DatabaseHelper.AddParam(cmdAud, "@motivo", "Anulación venta #" & numeroTicket & ": " & motivo)
                    DatabaseHelper.AddParam(cmdAud, "@uId", usuarioId)
                    cmdAud.ExecuteNonQuery()
                End Using
            Next
        End Sub

        ' Cambia el estado de la venta a 'Anulada', guarda quién autorizó y anexa el motivo en observaciones
        Private Sub MarcarVentaAnulada(conn As DbConnection, trans As DbTransaction, ventaId As Integer, motivo As String, usuarioId As Integer)
            Dim queryUpdateVenta = "UPDATE `ventas` SET `estado` = 'Anulada', `autorizado_por_id` = @autId, `observaciones` = (IFNULL(`observaciones`,'') || ' | ANULADA: ' || @motivo) WHERE `id` = @vId;"
            Using cmdUpdVenta = DatabaseHelper.CreateCommand(conn, queryUpdateVenta, trans)
                DatabaseHelper.AddParam(cmdUpdVenta, "@motivo", motivo)
                DatabaseHelper.AddParam(cmdUpdVenta, "@autId", usuarioId)
                DatabaseHelper.AddParam(cmdUpdVenta, "@vId", ventaId)
                cmdUpdVenta.ExecuteNonQuery()
            End Using
        End Sub

        ' Descuenta de la caja del turno los montos de una venta anulada para mantener cuadradas las cuentas del arqueo
        Private Sub ActualizarCajaAnulacion(conn As DbConnection, trans As DbTransaction, cajaId As Integer, montoEfectivo As Decimal, montoDigital As Decimal)
            Dim updateCajaQuery = "UPDATE `cajas` SET " &
                                  "`total_ventas_efectivo` = ROUND(`total_ventas_efectivo` - @efec, 2), " &
                                  "`total_ventas_digital` = ROUND(`total_ventas_digital` - @dig, 2), " &
                                  "`monto_esperado` = ROUND(`monto_esperado` - @efec, 2) " &
                                  "WHERE `id` = @cajaId;"

            Using cmdCaja = DatabaseHelper.CreateCommand(conn, updateCajaQuery, trans)
                DatabaseHelper.AddParam(cmdCaja, "@efec", montoEfectivo)
                DatabaseHelper.AddParam(cmdCaja, "@dig", montoDigital)
                DatabaseHelper.AddParam(cmdCaja, "@cajaId", cajaId)
                cmdCaja.ExecuteNonQuery()
            End Using
        End Sub

        ''' <summary>
        ''' Obtiene únicamente los tickets emitidos por un vendedor puntual en su caja abierta actual (pantalla Mis Ventas del Turno).
        ''' </summary>
        Public Function GetVentasPorTurnoVendedor(cajaId As Integer, usuarioId As Integer, Optional busqueda As String = "") As List(Of Venta)
            Dim list = New List(Of Venta)()
            Try
                Dim query = "SELECT v.*, u.nombre_completo AS usuario_nombre, " &
                            "IFNULL(ua.nombre_completo, '') AS autorizado_por_nombre, " &
                            "(c.apellido || ', ' || c.nombre) AS cliente_nombre, " &
                            "c.dni_cuit AS cliente_dni " &
                            "FROM `ventas` v " &
                            "INNER JOIN `usuarios` u ON v.usuario_id = u.id " &
                            "LEFT JOIN `usuarios` ua ON v.autorizado_por_id = ua.id " &
                            "INNER JOIN `clientes` c ON v.cliente_id = c.id " &
                            "WHERE v.caja_id = @cId AND v.usuario_id = @uId "

                Dim params = New Dictionary(Of String, Object) From {
                    {"@cId", cajaId},
                    {"@uId", usuarioId}
                }

                If Not String.IsNullOrWhiteSpace(busqueda) Then
                    query &= "AND (v.numero_ticket LIKE @b OR c.nombre LIKE @b OR c.apellido LIKE @b OR c.dni_cuit LIKE @b) "
                    params.Add("@b", "%" & busqueda.Trim() & "%")
                End If

                query &= "ORDER BY v.fecha DESC;"

                Dim dt = DatabaseHelper.ExecuteQuery(query, params)
                For Each row As DataRow In dt.Rows
                    list.Add(MapVenta(row))
                Next
            Catch ex As Exception
                AppLogger.LogError("VentaService.GetVentasPorTurnoVendedor", ex)
            End Try
            Return list
        End Function

        ''' <summary>
        ''' Procesa el cambio de talle o color de una prenda ya vendida sin anular el ticket, dentro de una transacción ACID.
        ''' Devuelve la prenda original al stock, descuenta la nueva variante, audita ambos movimientos y actualiza el renglón.
        ''' </summary>
        Public Function ProcesarCambioPrenda(ventaId As Integer, detalleId As Integer, nuevoTalleId As Integer, nuevoColor As String, usuarioId As Integer, motivo As String, ByRef errorMessage As String) As Boolean
            Try
                Using conn = DatabaseHelper.GetConnection()
                    conn.Open()
                    Using trans = conn.BeginTransaction()
                        Try
                            ' 1. Obtener datos actuales del renglón y de la venta
                            Dim qDet = "SELECT dv.*, v.numero_ticket, v.estado, p.codigo_barra, p.nombre AS prod_nombre " &
                                       "FROM `detalle_ventas` dv " &
                                       "INNER JOIN `ventas` v ON dv.venta_id = v.id " &
                                       "INNER JOIN `productos` p ON dv.producto_id = p.id " &
                                       "WHERE dv.id = @dId AND dv.venta_id = @vId;"
                            Dim prodId As Integer = 0
                            Dim oldTalleId As Integer = 0
                            Dim oldColor As String = ""
                            Dim cantidad As Integer = 0
                            Dim numeroTicket As String = ""
                            Dim estadoVenta As String = ""
                            Dim prodNombre As String = ""
                            Dim codBarra As String = ""

                            Using cmdDet = DatabaseHelper.CreateCommand(conn, qDet, trans)
                                DatabaseHelper.AddParam(cmdDet, "@dId", detalleId)
                                DatabaseHelper.AddParam(cmdDet, "@vId", ventaId)
                                Using r = cmdDet.ExecuteReader()
                                    If Not r.Read() Then
                                        errorMessage = "No se encontró el ítem de la venta seleccionado."
                                        trans.Rollback()
                                        Return False
                                    End If
                                    prodId = Convert.ToInt32(r("producto_id"))
                                    oldTalleId = Convert.ToInt32(r("talle_id"))
                                    oldColor = r("color").ToString()
                                    cantidad = Convert.ToInt32(r("cantidad"))
                                    numeroTicket = r("numero_ticket").ToString()
                                    estadoVenta = r("estado").ToString()
                                    prodNombre = r("prod_nombre").ToString()
                                    codBarra = r("codigo_barra").ToString()
                                End Using
                            End Using

                            If estadoVenta.Equals("Anulada", StringComparison.OrdinalIgnoreCase) Then
                                errorMessage = "No se puede realizar un cambio sobre una venta anulada."
                                trans.Rollback()
                                Return False
                            End If

                            If oldTalleId = nuevoTalleId AndAlso String.Equals(oldColor, nuevoColor, StringComparison.OrdinalIgnoreCase) Then
                                errorMessage = "Debe seleccionar un talle o color distinto al original para realizar el cambio."
                                trans.Rollback()
                                Return False
                            End If

                            ' 2. Verificar stock disponible de la nueva variante (talle/color)
                            Dim qNuevoStock = "SELECT pt.stock_actual, t.nombre AS talle_nombre " &
                                              "FROM `producto_talles` pt " &
                                              "INNER JOIN `talles` t ON pt.talle_id = t.id " &
                                              "WHERE pt.producto_id = @pId AND pt.talle_id = @tId AND pt.color = @col;"
                            Dim stockNuevoActual As Integer = -1
                            Dim nuevoTalleNombre As String = ""
                            Using cmdNS = DatabaseHelper.CreateCommand(conn, qNuevoStock, trans)
                                DatabaseHelper.AddParam(cmdNS, "@pId", prodId)
                                DatabaseHelper.AddParam(cmdNS, "@tId", nuevoTalleId)
                                DatabaseHelper.AddParam(cmdNS, "@col", nuevoColor)
                                Using r = cmdNS.ExecuteReader()
                                    If r.Read() Then
                                        stockNuevoActual = Convert.ToInt32(r("stock_actual"))
                                        nuevoTalleNombre = r("talle_nombre").ToString()
                                    End If
                                End Using
                            End Using

                            If stockNuevoActual < cantidad Then
                                errorMessage = $"Stock insuficiente en el nuevo talle/color seleccionado (Disponible: {Math.Max(0, stockNuevoActual)}, Requerido: {cantidad})."
                                trans.Rollback()
                                Return False
                            End If

                            ' 3. Obtener stock actual de la variante devuelta
                            Dim qOldStock = "SELECT pt.stock_actual, t.nombre AS talle_nombre " &
                                            "FROM `producto_talles` pt " &
                                            "INNER JOIN `talles` t ON pt.talle_id = t.id " &
                                            "WHERE pt.producto_id = @pId AND pt.talle_id = @tId AND pt.color = @col;"
                            Dim stockOldActual As Integer = 0
                            Dim oldTalleNombre As String = ""
                            Using cmdOS = DatabaseHelper.CreateCommand(conn, qOldStock, trans)
                                DatabaseHelper.AddParam(cmdOS, "@pId", prodId)
                                DatabaseHelper.AddParam(cmdOS, "@tId", oldTalleId)
                                DatabaseHelper.AddParam(cmdOS, "@col", oldColor)
                                Using r = cmdOS.ExecuteReader()
                                    If r.Read() Then
                                        stockOldActual = Convert.ToInt32(r("stock_actual"))
                                        oldTalleNombre = r("talle_nombre").ToString()
                                    End If
                                End Using
                            End Using

                            Dim queryUpdStock = "UPDATE `producto_talles` SET `stock_actual` = @nuevoStock WHERE `producto_id` = @pId AND `talle_id` = @tId AND `color` = @col;"
                            Dim queryAud = "INSERT INTO `movimientos_stock` (`producto_id`, `talle_id`, `color`, `tipo_movimiento`, `cantidad`, `stock_anterior`, `stock_posterior`, `motivo`, `usuario_id`) " &
                                           "VALUES (@pId, @tId, @col, @tipo, @cant, @ant, @post, @motivo, @uId);"

                            ' 3a. Reingresar al stock la prenda devuelta
                            Dim postOld As Integer = stockOldActual + cantidad
                            Using cmdUpdOld = DatabaseHelper.CreateCommand(conn, queryUpdStock, trans)
                                DatabaseHelper.AddParam(cmdUpdOld, "@nuevoStock", postOld)
                                DatabaseHelper.AddParam(cmdUpdOld, "@pId", prodId)
                                DatabaseHelper.AddParam(cmdUpdOld, "@tId", oldTalleId)
                                DatabaseHelper.AddParam(cmdUpdOld, "@col", oldColor)
                                cmdUpdOld.ExecuteNonQuery()
                            End Using

                            Dim motivoTexto As String = If(String.IsNullOrWhiteSpace(motivo), "Cambio de talle/color", motivo.Trim())
                            Using cmdAudIn = DatabaseHelper.CreateCommand(conn, queryAud, trans)
                                DatabaseHelper.AddParam(cmdAudIn, "@pId", prodId)
                                DatabaseHelper.AddParam(cmdAudIn, "@tId", oldTalleId)
                                DatabaseHelper.AddParam(cmdAudIn, "@col", oldColor)
                                DatabaseHelper.AddParam(cmdAudIn, "@tipo", "Devolucion")
                                DatabaseHelper.AddParam(cmdAudIn, "@cant", cantidad)
                                DatabaseHelper.AddParam(cmdAudIn, "@ant", stockOldActual)
                                DatabaseHelper.AddParam(cmdAudIn, "@post", postOld)
                                DatabaseHelper.AddParam(cmdAudIn, "@motivo", $"Cambio en Ticket #{numeroTicket} (Ingresa T.{oldTalleNombre} {oldColor}): {motivoTexto}")
                                DatabaseHelper.AddParam(cmdAudIn, "@uId", usuarioId)
                                cmdAudIn.ExecuteNonQuery()
                            End Using

                            ' 3b. Descontar del stock la nueva prenda entregada
                            Dim postNew As Integer = stockNuevoActual - cantidad
                            Using cmdUpdNew = DatabaseHelper.CreateCommand(conn, queryUpdStock, trans)
                                DatabaseHelper.AddParam(cmdUpdNew, "@nuevoStock", postNew)
                                DatabaseHelper.AddParam(cmdUpdNew, "@pId", prodId)
                                DatabaseHelper.AddParam(cmdUpdNew, "@tId", nuevoTalleId)
                                DatabaseHelper.AddParam(cmdUpdNew, "@col", nuevoColor)
                                cmdUpdNew.ExecuteNonQuery()
                            End Using

                            Using cmdAudOut = DatabaseHelper.CreateCommand(conn, queryAud, trans)
                                DatabaseHelper.AddParam(cmdAudOut, "@pId", prodId)
                                DatabaseHelper.AddParam(cmdAudOut, "@tId", nuevoTalleId)
                                DatabaseHelper.AddParam(cmdAudOut, "@col", nuevoColor)
                                DatabaseHelper.AddParam(cmdAudOut, "@tipo", "Cambio_Prenda")
                                DatabaseHelper.AddParam(cmdAudOut, "@cant", cantidad)
                                DatabaseHelper.AddParam(cmdAudOut, "@ant", stockNuevoActual)
                                DatabaseHelper.AddParam(cmdAudOut, "@post", postNew)
                                DatabaseHelper.AddParam(cmdAudOut, "@motivo", $"Cambio en Ticket #{numeroTicket} (Egresa T.{nuevoTalleNombre} {nuevoColor}): {motivoTexto}")
                                DatabaseHelper.AddParam(cmdAudOut, "@uId", usuarioId)
                                cmdAudOut.ExecuteNonQuery()
                            End Using

                            ' 4. Actualizar el renglón del ticket y dejar constancia en observaciones de la venta
                            Dim nuevaDesc As String = $"{prodNombre} (T.{nuevoTalleNombre} - {nuevoColor})"
                            Dim qUpdDet = "UPDATE `detalle_ventas` SET `talle_id` = @tId, `color` = @col, `descripcion_articulo` = @desc WHERE `id` = @dId;"
                            Using cmdUpdDet = DatabaseHelper.CreateCommand(conn, qUpdDet, trans)
                                DatabaseHelper.AddParam(cmdUpdDet, "@tId", nuevoTalleId)
                                DatabaseHelper.AddParam(cmdUpdDet, "@col", nuevoColor)
                                DatabaseHelper.AddParam(cmdUpdDet, "@desc", nuevaDesc)
                                DatabaseHelper.AddParam(cmdUpdDet, "@dId", detalleId)
                                cmdUpdDet.ExecuteNonQuery()
                            End Using

                            Dim notaCambio As String = $"CAMBIO ({DateTime.Now:dd/MM HH:mm}): {prodNombre} de [T.{oldTalleNombre} {oldColor}] a [T.{nuevoTalleNombre} {nuevoColor}]"
                            Dim qUpdVenta = "UPDATE `ventas` SET `observaciones` = CASE WHEN IFNULL(`observaciones`, '') = '' THEN @nota ELSE (`observaciones` || ' | ' || @nota) END WHERE `id` = @vId;"
                            Using cmdUpdV = DatabaseHelper.CreateCommand(conn, qUpdVenta, trans)
                                DatabaseHelper.AddParam(cmdUpdV, "@nota", notaCambio)
                                DatabaseHelper.AddParam(cmdUpdV, "@vId", ventaId)
                                cmdUpdV.ExecuteNonQuery()
                            End Using

                            trans.Commit()
                            Return True
                        Catch exTrans As Exception
                            trans.Rollback()
                            errorMessage = exTrans.Message
                            AppLogger.LogError("VentaService.ProcesarCambioPrenda [Transaction]", exTrans)
                            Return False
                        End Try
                    End Using
                End Using
            Catch ex As Exception
                errorMessage = ex.Message
                AppLogger.LogError("VentaService.ProcesarCambioPrenda", ex)
                Return False
            End Try
        End Function

        ' Convierte una fila DataRow de la tabla 'ventas' en un objeto Venta en memoria
        Private Shared Function MapVenta(row As DataRow) As Venta
            Dim autId As Nullable(Of Integer) = Nothing
            If row.Table.Columns.Contains("autorizado_por_id") AndAlso Not Convert.IsDBNull(row("autorizado_por_id")) Then
                autId = Convert.ToInt32(row("autorizado_por_id"))
            End If

            Dim autNombre As String = ""
            If row.Table.Columns.Contains("autorizado_por_nombre") AndAlso Not Convert.IsDBNull(row("autorizado_por_nombre")) Then
                autNombre = row("autorizado_por_nombre").ToString()
            End If
            If String.IsNullOrWhiteSpace(autNombre) Then
                autNombre = "-"
            End If

            Dim montoEfec As Nullable(Of Decimal) = Nothing
            If row.Table.Columns.Contains("monto_efectivo") AndAlso Not Convert.IsDBNull(row("monto_efectivo")) Then
                montoEfec = Math.Round(Convert.ToDecimal(row("monto_efectivo")), 2)
            End If

            Return New Venta() With {
                .Id = Convert.ToInt32(row("id")),
                .NumeroTicket = row("numero_ticket").ToString(),
                .Fecha = Convert.ToDateTime(row("fecha")),
                .UsuarioId = Convert.ToInt32(row("usuario_id")),
                .UsuarioNombre = row("usuario_nombre").ToString(),
                .AutorizadoPorId = autId,
                .AutorizadoPorNombre = autNombre,
                .ClienteId = Convert.ToInt32(row("cliente_id")),
                .ClienteNombre = row("cliente_nombre").ToString(),
                .CajaId = Convert.ToInt32(row("caja_id")),
                .TipoComprobante = row("tipo_comprobante").ToString(),
                .MetodoPago = row("metodo_pago").ToString(),
                .Subtotal = Math.Round(Convert.ToDecimal(row("subtotal")), 2),
                .DescuentoPorcentaje = Math.Round(Convert.ToDecimal(row("descuento_porcentaje")), 2),
                .DescuentoMonto = Math.Round(Convert.ToDecimal(row("descuento_monto")), 2),
                .RecargoMonto = Math.Round(Convert.ToDecimal(row("recargo_monto")), 2),
                .Total = Math.Round(Convert.ToDecimal(row("total")), 2),
                .MontoEfectivo = montoEfec,
                .MontoAbonado = Math.Round(Convert.ToDecimal(row("monto_abonado")), 2),
                .Vuelto = Math.Round(Convert.ToDecimal(row("vuelto")), 2),
                .Estado = row("estado").ToString(),
                .Observaciones = If(Convert.IsDBNull(row("observaciones")), "", row("observaciones").ToString())
            }
        End Function

        ' Convierte una fila DataRow de la tabla 'detalle_ventas' en un objeto DetalleVenta en memoria
        Private Shared Function MapDetalleVenta(row As DataRow) As DetalleVenta
            Return New DetalleVenta() With {
                .Id = Convert.ToInt32(row("id")),
                .VentaId = Convert.ToInt32(row("venta_id")),
                .ProductoId = Convert.ToInt32(row("producto_id")),
                .TalleId = Convert.ToInt32(row("talle_id")),
                .TalleNombre = row("talle_nombre").ToString(),
                .Color = row("color").ToString(),
                .CodigoBarra = row("codigo_barra").ToString(),
                .DescripcionArticulo = row("descripcion_articulo").ToString(),
                .PrecioUnitario = Math.Round(Convert.ToDecimal(row("precio_unitario")), 2),
                .CostoUnitario = Math.Round(Convert.ToDecimal(row("costo_unitario")), 2),
                .Cantidad = Convert.ToInt32(row("cantidad")),
                .Subtotal = Math.Round(Convert.ToDecimal(row("subtotal")), 2)
            }
        End Function

    End Class
End Namespace
