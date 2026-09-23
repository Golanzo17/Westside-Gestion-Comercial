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
'    esperado en el turno abierto de caja.

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
        '   Paso 5: Sumamos el importe a la caja activa del vendedor (Efectivo o Digital).
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
                                ActualizarCajaVenta(conn, trans, venta.CajaId, venta.MetodoPago, venta.Total)
                            End If

                            ' Si todo salió perfecto, confirmamos la transacción de forma atómica
                            trans.Commit()
                            errorMessage = String.Empty
                            Return True
                        Catch exTrans As Exception
                            trans.Rollback()
                            errorMessage = "Error al procesar la venta: " & exTrans.Message
                            Return False
                        End Try
                    End Using
                End Using
            Catch ex As Exception
                errorMessage = "Error de conexión con la base de datos: " & ex.Message
                Return False
            End Try
        End Function

        Public Function GetVentas(fechaDesde As DateTime, fechaHasta As DateTime, Optional busqueda As String = "") As List(Of Venta)
            Dim list = New List(Of Venta)()
            Dim query = "SELECT v.*, u.nombre_completo AS usuario_nombre, (c.apellido || ', ' || c.nombre) AS cliente_nombre " &
                        "FROM `ventas` v " &
                        "INNER JOIN `usuarios` u ON v.usuario_id = u.id " &
                        "INNER JOIN `clientes` c ON v.cliente_id = c.id " &
                        "WHERE DATE(v.fecha) >= @desde AND DATE(v.fecha) <= @hasta "

            Dim params = New Dictionary(Of String, Object) From {
                {"@desde", fechaDesde.ToString("yyyy-MM-dd")},
                {"@hasta", fechaHasta.ToString("yyyy-MM-dd")}
            }

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

        ' Anulación Segura de Venta con Reversa Transaccional.
        ' Si un cliente devuelve prendas o hubo un error de cobro:
        ' 1. Verificamos que la venta exista y no haya sido anulada previamente.
        ' 2. Restituimos el stock físico sumando las unidades a producto_talles.
        ' 3. Auditamos en movimientos_stock con motivo "Anulacion_Venta" y el usuario supervisor.
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

                            ' Paso 1: Restituir stock físico de cada artículo y auditar en movimientos_stock
                            Dim detalles = GetDetalleVenta(ventaId)
                            Dim ticket = ventaRow("numero_ticket").ToString()
                            RestituirStockYAuditar(conn, trans, detalles, ticket, motivo, usuarioId)

                            ' Paso 2: Marcar comprobante como Anulado
                            MarcarVentaAnulada(conn, trans, ventaId, motivo)

                            ' Paso 3: Restar el monto de la caja activa para mantener el balance contable
                            Dim cajaId = Convert.ToInt32(ventaRow("caja_id"))
                            Dim total = Convert.ToDecimal(ventaRow("total"))
                            Dim metodo = ventaRow("metodo_pago").ToString()
                            If cajaId > 0 Then
                                ActualizarCajaAnulacion(conn, trans, cajaId, metodo, total)
                            End If

                            trans.Commit()
                            errorMessage = String.Empty
                            Return True
                        Catch exTrans As Exception
                            trans.Rollback()
                            errorMessage = "Error al anular la venta: " & exTrans.Message
                            Return False
                        End Try
                    End Using
                End Using
            Catch ex As Exception
                errorMessage = "Error de conexión: " & ex.Message
                Return False
            End Try
        End Function

        ' --- Helpers privados de soporte ---

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

        Private Function InsertarVenta(conn As DbConnection, trans As DbTransaction, venta As Venta) As Integer
            Dim queryVenta = "INSERT INTO `ventas` (`numero_ticket`, `fecha`, `usuario_id`, `cliente_id`, `caja_id`, " &
                             "`tipo_comprobante`, `metodo_pago`, `subtotal`, `descuento_porcentaje`, `descuento_monto`, " &
                             "`recargo_monto`, `total`, `monto_abonado`, `vuelto`, `estado`, `observaciones`) " &
                             "VALUES (@num, NOW(), @userId, @cliId, @cajaId, @tipoComp, @metodo, @sub, @descPorc, @descMonto, @recargo, @total, @abonado, @vuelto, 'Completada', @obs);"

            Using cmdVenta = DatabaseHelper.CreateCommand(conn, queryVenta, trans)
                DatabaseHelper.AddParam(cmdVenta, "@num", venta.NumeroTicket)
                DatabaseHelper.AddParam(cmdVenta, "@userId", venta.UsuarioId)
                DatabaseHelper.AddParam(cmdVenta, "@cliId", venta.ClienteId)
                DatabaseHelper.AddParam(cmdVenta, "@cajaId", venta.CajaId)
                DatabaseHelper.AddParam(cmdVenta, "@tipoComp", venta.TipoComprobante)
                DatabaseHelper.AddParam(cmdVenta, "@metodo", venta.MetodoPago)
                DatabaseHelper.AddParam(cmdVenta, "@sub", venta.Subtotal)
                DatabaseHelper.AddParam(cmdVenta, "@descPorc", venta.DescuentoPorcentaje)
                DatabaseHelper.AddParam(cmdVenta, "@descMonto", venta.DescuentoMonto)
                DatabaseHelper.AddParam(cmdVenta, "@recargo", venta.RecargoMonto)
                DatabaseHelper.AddParam(cmdVenta, "@total", venta.Total)
                DatabaseHelper.AddParam(cmdVenta, "@abonado", venta.MontoAbonado)
                DatabaseHelper.AddParam(cmdVenta, "@vuelto", venta.Vuelto)
                DatabaseHelper.AddParam(cmdVenta, "@obs", venta.Observaciones)
                cmdVenta.ExecuteNonQuery()
            End Using

            Return Convert.ToInt32(DatabaseHelper.GetLastInsertedId(conn, trans))
        End Function

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
                    DatabaseHelper.AddParam(cmdDet, "@pu", d.PrecioUnitario)
                    DatabaseHelper.AddParam(cmdDet, "@cu", d.CostoUnitario)
                    DatabaseHelper.AddParam(cmdDet, "@cant", d.Cantidad)
                    DatabaseHelper.AddParam(cmdDet, "@sub", d.Subtotal)
                    cmdDet.ExecuteNonQuery()
                End Using
            Next
        End Sub

        Private Sub ActualizarCajaVenta(conn As DbConnection, trans As DbTransaction, cajaId As Integer, metodoPago As String, total As Decimal)
            Dim esEfectivo = metodoPago.Equals("Efectivo", StringComparison.OrdinalIgnoreCase)
            Dim updateCajaQuery = If(esEfectivo,
                "UPDATE `cajas` SET `total_ventas_efectivo` = `total_ventas_efectivo` + @tot, `monto_esperado` = `monto_esperado` + @tot WHERE `id` = @cajaId;",
                "UPDATE `cajas` SET `total_ventas_digital` = `total_ventas_digital` + @tot WHERE `id` = @cajaId;")

            Using cmdUpdCaja = DatabaseHelper.CreateCommand(conn, updateCajaQuery, trans)
                DatabaseHelper.AddParam(cmdUpdCaja, "@tot", total)
                DatabaseHelper.AddParam(cmdUpdCaja, "@cajaId", cajaId)
                cmdUpdCaja.ExecuteNonQuery()
            End Using
        End Sub

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

        Private Sub RestituirStockYAuditar(conn As DbConnection, trans As DbTransaction, detalles As List(Of DetalleVenta), numeroTicket As String, motivo As String, usuarioId As Integer)
            Dim queryStock = "UPDATE `producto_talles` SET `stock_actual` = `stock_actual` + @cant WHERE `producto_id` = @pId AND `talle_id` = @tId AND `color` = @color;"
            Dim queryAud = "INSERT INTO `movimientos_stock` (`producto_id`, `talle_id`, `color`, `tipo_movimiento`, `cantidad`, `stock_anterior`, `stock_posterior`, `motivo`, `usuario_id`) " &
                           "VALUES (@pId, @tId, @color, 'Anulacion_Venta', @cant, 0, 0, @motivo, @uId);"

            For Each item In detalles
                Using cmdStock = DatabaseHelper.CreateCommand(conn, queryStock, trans)
                    DatabaseHelper.AddParam(cmdStock, "@cant", item.Cantidad)
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
                    DatabaseHelper.AddParam(cmdAud, "@motivo", "Anulación venta #" & numeroTicket & ": " & motivo)
                    DatabaseHelper.AddParam(cmdAud, "@uId", usuarioId)
                    cmdAud.ExecuteNonQuery()
                End Using
            Next
        End Sub

        Private Sub MarcarVentaAnulada(conn As DbConnection, trans As DbTransaction, ventaId As Integer, motivo As String)
            Dim queryUpdateVenta = "UPDATE `ventas` SET `estado` = 'Anulada', `observaciones` = (IFNULL(`observaciones`,'') || ' | ANULADA: ' || @motivo) WHERE `id` = @vId;"
            Using cmdUpdVenta = DatabaseHelper.CreateCommand(conn, queryUpdateVenta, trans)
                DatabaseHelper.AddParam(cmdUpdVenta, "@motivo", motivo)
                DatabaseHelper.AddParam(cmdUpdVenta, "@vId", ventaId)
                cmdUpdVenta.ExecuteNonQuery()
            End Using
        End Sub

        Private Sub ActualizarCajaAnulacion(conn As DbConnection, trans As DbTransaction, cajaId As Integer, metodo As String, total As Decimal)
            Dim esEfectivo = metodo.Equals("Efectivo", StringComparison.OrdinalIgnoreCase)
            Dim updateCajaQuery = If(esEfectivo,
                "UPDATE `cajas` SET `total_ventas_efectivo` = `total_ventas_efectivo` - @tot, `monto_esperado` = `monto_esperado` - @tot WHERE `id` = @cajaId;",
                "UPDATE `cajas` SET `total_ventas_digital` = `total_ventas_digital` - @tot WHERE `id` = @cajaId;")

            Using cmdCaja = DatabaseHelper.CreateCommand(conn, updateCajaQuery, trans)
                DatabaseHelper.AddParam(cmdCaja, "@tot", total)
                DatabaseHelper.AddParam(cmdCaja, "@cajaId", cajaId)
                cmdCaja.ExecuteNonQuery()
            End Using
        End Sub

        Private Shared Function MapVenta(row As DataRow) As Venta
            Return New Venta() With {
                .Id = Convert.ToInt32(row("id")),
                .NumeroTicket = row("numero_ticket").ToString(),
                .Fecha = Convert.ToDateTime(row("fecha")),
                .UsuarioId = Convert.ToInt32(row("usuario_id")),
                .UsuarioNombre = row("usuario_nombre").ToString(),
                .ClienteId = Convert.ToInt32(row("cliente_id")),
                .ClienteNombre = row("cliente_nombre").ToString(),
                .CajaId = Convert.ToInt32(row("caja_id")),
                .TipoComprobante = row("tipo_comprobante").ToString(),
                .MetodoPago = row("metodo_pago").ToString(),
                .Subtotal = Convert.ToDecimal(row("subtotal")),
                .DescuentoPorcentaje = Convert.ToDecimal(row("descuento_porcentaje")),
                .DescuentoMonto = Convert.ToDecimal(row("descuento_monto")),
                .RecargoMonto = Convert.ToDecimal(row("recargo_monto")),
                .Total = Convert.ToDecimal(row("total")),
                .MontoAbonado = Convert.ToDecimal(row("monto_abonado")),
                .Vuelto = Convert.ToDecimal(row("vuelto")),
                .Estado = row("estado").ToString(),
                .Observaciones = If(Convert.IsDBNull(row("observaciones")), "", row("observaciones").ToString())
            }
        End Function

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
                .PrecioUnitario = Convert.ToDecimal(row("precio_unitario")),
                .CostoUnitario = Convert.ToDecimal(row("costo_unitario")),
                .Cantidad = Convert.ToInt32(row("cantidad")),
                .Subtotal = Convert.ToDecimal(row("subtotal"))
            }
        End Function

    End Class
End Namespace
