Imports System.Data
Imports MySqlConnector
Imports GestionComercial.Data
Imports GestionComercial.Models

Namespace Services
    Public Class VentaService

        Public Function GenerarNumeroTicket() As String
            Dim prefijo As String = "TICK-" & DateTime.Now.ToString("yyyyMMdd") & "-"
            Dim query As String = "SELECT COUNT(*) FROM `ventas` WHERE DATE(`fecha`) = CURDATE();"
            Dim count As Long = Convert.ToInt64(DatabaseHelper.ExecuteScalar(query))
            Return prefijo & (count + 1).ToString("D4")
        End Function

        Public Function ProcesarVenta(venta As Venta, ByRef errorMessage As String) As Boolean
            If venta.Detalles Is Nothing OrElse venta.Detalles.Count = 0 Then
                errorMessage = "No hay artículos en el carrito de venta."
                Return False
            End If

            Try
                Using conn As Common.DbConnection = DatabaseHelper.GetConnection()
                    conn.Open()
                    Using trans As Common.DbTransaction = conn.BeginTransaction()
                        Try
                            ' 1. Validar y descontar stock de cada prenda por su Talle y Color
                            For Each item In venta.Detalles
                                Dim queryStock As String = "SELECT stock_actual FROM `producto_talles` WHERE `producto_id` = @prodId AND `talle_id` = @talleId AND `color` = @color FOR UPDATE;"
                                Dim stockActual As Integer = 0
                                Using cmdStock = DatabaseHelper.CreateCommand(conn, queryStock, trans)
                                    DatabaseHelper.AddParam(cmdStock, "@prodId", item.ProductoId)
                                    DatabaseHelper.AddParam(cmdStock, "@talleId", item.TalleId)
                                    DatabaseHelper.AddParam(cmdStock, "@color", item.Color)
                                    Dim res = cmdStock.ExecuteScalar()
                                    If res Is Nothing OrElse IsDBNull(res) Then
                                        Throw New Exception($"El artículo '{item.DescripcionArticulo}' (Talle: {item.TalleNombre}, Color: {item.Color}) no tiene stock asignado.")
                                    End If
                                    stockActual = Convert.ToInt32(res)
                                End Using

                                If stockActual < item.Cantidad Then
                                    Throw New Exception($"Stock insuficiente para '{item.DescripcionArticulo}' (Talle: {item.TalleNombre}). Stock disponible: {stockActual}, solicitado: {item.Cantidad}.")
                                End If

                                ' Descontar stock
                                Dim nuevoStock As Integer = stockActual - item.Cantidad
                                Dim queryDescontar As String = "UPDATE `producto_talles` SET `stock_actual` = @nuevoStock WHERE `producto_id` = @prodId AND `talle_id` = @talleId AND `color` = @color;"
                                Using cmdDesc = DatabaseHelper.CreateCommand(conn, queryDescontar, trans)
                                    DatabaseHelper.AddParam(cmdDesc, "@nuevoStock", nuevoStock)
                                    DatabaseHelper.AddParam(cmdDesc, "@prodId", item.ProductoId)
                                    DatabaseHelper.AddParam(cmdDesc, "@talleId", item.TalleId)
                                    DatabaseHelper.AddParam(cmdDesc, "@color", item.Color)
                                    cmdDesc.ExecuteNonQuery()
                                End Using

                                ' Registrar en movimientos_stock
                                Dim queryMovStock As String = "INSERT INTO `movimientos_stock` (`producto_id`, `talle_id`, `color`, `tipo_movimiento`, `cantidad`, `stock_anterior`, `stock_posterior`, `motivo`, `usuario_id`) " &
                                                             "VALUES (@prodId, @talleId, @color, 'Venta', @cant, @ant, @post, @motivo, @userId);"
                                Using cmdMov = DatabaseHelper.CreateCommand(conn, queryMovStock, trans)
                                    DatabaseHelper.AddParam(cmdMov, "@prodId", item.ProductoId)
                                    DatabaseHelper.AddParam(cmdMov, "@talleId", item.TalleId)
                                    DatabaseHelper.AddParam(cmdMov, "@color", item.Color)
                                    DatabaseHelper.AddParam(cmdMov, "@cant", item.Cantidad)
                                    DatabaseHelper.AddParam(cmdMov, "@ant", stockActual)
                                    DatabaseHelper.AddParam(cmdMov, "@post", nuevoStock)
                                    DatabaseHelper.AddParam(cmdMov, "@motivo", "Venta Ticket #" & venta.NumeroTicket)
                                    DatabaseHelper.AddParam(cmdMov, "@userId", venta.UsuarioId)
                                    cmdMov.ExecuteNonQuery()
                                End Using
                            Next

                            ' 2. Insertar Encabezado de Venta
                            Dim queryVenta As String = "INSERT INTO `ventas` (`numero_ticket`, `fecha`, `usuario_id`, `cliente_id`, `caja_id`, " &
                                                      "`tipo_comprobante`, `metodo_pago`, `subtotal`, `descuento_porcentaje`, `descuento_monto`, " &
                                                      "`recargo_monto`, `total`, `monto_abonado`, `vuelto`, `estado`, `observaciones`) " &
                                                      "VALUES (@num, NOW(), @userId, @cliId, @cajaId, @tipoComp, @metodo, @sub, @descPorc, @descMonto, @recargo, @total, @abonado, @vuelto, 'Completada', @obs);"

                            Dim ventaId As Integer = 0
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
                                ventaId = Convert.ToInt32(DatabaseHelper.GetLastInsertedId(conn, trans))
                                venta.Id = ventaId
                            End Using

                            ' 3. Insertar Detalles de la Venta
                            For Each d In venta.Detalles
                                Dim queryDetalle As String = "INSERT INTO `detalle_ventas` (`venta_id`, `producto_id`, `talle_id`, `color`, `codigo_barra`, `descripcion_articulo`, `precio_unitario`, `costo_unitario`, `cantidad`, `subtotal`) " &
                                                            "VALUES (@vId, @pId, @tId, @color, @cod, @desc, @pu, @cu, @cant, @sub);"
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

                            ' 4. Actualizar totales de la Caja activa si corresponde
                            If venta.CajaId > 0 Then
                                Dim updateCajaQuery As String
                                If venta.MetodoPago.Equals("Efectivo", StringComparison.OrdinalIgnoreCase) Then
                                    updateCajaQuery = "UPDATE `cajas` SET `total_ventas_efectivo` = `total_ventas_efectivo` + @tot, `monto_esperado` = `monto_esperado` + @tot WHERE `id` = @cajaId;"
                                Else
                                    updateCajaQuery = "UPDATE `cajas` SET `total_ventas_digital` = `total_ventas_digital` + @tot WHERE `id` = @cajaId;"
                                End If

                                Using cmdUpdCaja = DatabaseHelper.CreateCommand(conn, updateCajaQuery, trans)
                                    DatabaseHelper.AddParam(cmdUpdCaja, "@tot", venta.Total)
                                    DatabaseHelper.AddParam(cmdUpdCaja, "@cajaId", venta.CajaId)
                                    cmdUpdCaja.ExecuteNonQuery()
                                End Using
                            End If

                            trans.Commit()
                            errorMessage = String.Empty
                            Return True
                        Catch exTrans As Exception
                            trans.Rollback()
                            errorMessage = exTrans.Message
                            Return False
                        End Try
                    End Using
                End Using
            Catch ex As Exception
                errorMessage = ex.Message
                Return False
            End Try
        End Function

        Public Function GetVentas(fechaDesde As DateTime, fechaHasta As DateTime, Optional busqueda As String = "") As List(Of Venta)
            Dim list As New List(Of Venta)()
            Dim query As String = "SELECT v.*, u.nombre_completo AS usuario_nombre, (c.apellido || ', ' || c.nombre) AS cliente_nombre " &
                                 "FROM `ventas` v " &
                                 "INNER JOIN `usuarios` u ON v.usuario_id = u.id " &
                                 "INNER JOIN `clientes` c ON v.cliente_id = c.id " &
                                 "WHERE DATE(v.fecha) >= @desde AND DATE(v.fecha) <= @hasta "

            Dim params As New Dictionary(Of String, Object) From {
                {"@desde", fechaDesde.ToString("yyyy-MM-dd")},
                {"@hasta", fechaHasta.ToString("yyyy-MM-dd")}
            }

            If Not String.IsNullOrWhiteSpace(busqueda) Then
                query &= "AND (v.numero_ticket LIKE @b OR c.nombre LIKE @b OR c.apellido LIKE @b OR c.dni_cuit LIKE @b) "
                params.Add("@b", "%" & busqueda.Trim() & "%")
            End If

            query &= "ORDER BY v.fecha DESC;"

            Dim dt As DataTable = DatabaseHelper.ExecuteQuery(query, params)
            For Each row As DataRow In dt.Rows
                list.Add(New Venta() With {
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
                    .Observaciones = If(IsDBNull(row("observaciones")), "", row("observaciones").ToString())
                })
            Next
            Return list
        End Function

        Public Function GetDetalleVenta(ventaId As Integer) As List(Of DetalleVenta)
            Dim list As New List(Of DetalleVenta)()
            Dim query As String = "SELECT dv.*, t.nombre AS talle_nombre " &
                                 "FROM `detalle_ventas` dv " &
                                 "INNER JOIN `talles` t ON dv.talle_id = t.id " &
                                 "WHERE dv.venta_id = @vId;"

            Dim dt As DataTable = DatabaseHelper.ExecuteQuery(query, New Dictionary(Of String, Object) From {{"@vId", ventaId}})
            For Each row As DataRow In dt.Rows
                list.Add(New DetalleVenta() With {
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
                })
            Next
            Return list
        End Function

        Public Function AnularVenta(ventaId As Integer, usuarioId As Integer, motivo As String, ByRef errorMessage As String) As Boolean
            Try
                Using conn As Common.DbConnection = DatabaseHelper.GetConnection()
                    conn.Open()
                    Using trans As Common.DbTransaction = conn.BeginTransaction()
                        Try
                            ' Obtener la venta
                            Dim queryVenta As String = "SELECT * FROM `ventas` WHERE `id` = @vId FOR UPDATE;"
                            Dim ventaRow As DataRow = Nothing
                            Using cmdVenta = DatabaseHelper.CreateCommand(conn, queryVenta, trans)
                                DatabaseHelper.AddParam(cmdVenta, "@vId", ventaId)
                                Using reader = cmdVenta.ExecuteReader()
                                    Dim dt As New DataTable()
                                    dt.Load(reader)
                                    If dt.Rows.Count = 0 Then Throw New Exception("Venta no encontrada.")
                                    ventaRow = dt.Rows(0)
                                End Using
                            End Using

                            If ventaRow("estado").ToString().Equals("Anulada", StringComparison.OrdinalIgnoreCase) Then
                                Throw New Exception("Esta venta ya se encuentra anulada.")
                            End If

                            ' Restituir el stock de cada artículo vendido
                            Dim detalles = GetDetalleVenta(ventaId)
                            For Each item In detalles
                                Dim queryStock As String = "UPDATE `producto_talles` SET `stock_actual` = `stock_actual` + @cant WHERE `producto_id` = @pId AND `talle_id` = @tId AND `color` = @color;"
                                Using cmdStock = DatabaseHelper.CreateCommand(conn, queryStock, trans)
                                    DatabaseHelper.AddParam(cmdStock, "@cant", item.Cantidad)
                                    DatabaseHelper.AddParam(cmdStock, "@pId", item.ProductoId)
                                    DatabaseHelper.AddParam(cmdStock, "@tId", item.TalleId)
                                    DatabaseHelper.AddParam(cmdStock, "@color", item.Color)
                                    cmdStock.ExecuteNonQuery()
                                End Using

                                ' Registrar auditoría de restitución
                                Dim queryAud As String = "INSERT INTO `movimientos_stock` (`producto_id`, `talle_id`, `color`, `tipo_movimiento`, `cantidad`, `stock_anterior`, `stock_posterior`, `motivo`, `usuario_id`) " &
                                                        "VALUES (@pId, @tId, @color, 'Anulacion_Venta', @cant, 0, 0, @motivo, @uId);"
                                Using cmdAud = DatabaseHelper.CreateCommand(conn, queryAud, trans)
                                    DatabaseHelper.AddParam(cmdAud, "@pId", item.ProductoId)
                                    DatabaseHelper.AddParam(cmdAud, "@tId", item.TalleId)
                                    DatabaseHelper.AddParam(cmdAud, "@color", item.Color)
                                    DatabaseHelper.AddParam(cmdAud, "@cant", item.Cantidad)
                                    DatabaseHelper.AddParam(cmdAud, "@motivo", "Anulación venta #" & ventaRow("numero_ticket").ToString() & ": " & motivo)
                                    DatabaseHelper.AddParam(cmdAud, "@uId", usuarioId)
                                    cmdAud.ExecuteNonQuery()
                                End Using
                            Next

                            ' Marcar venta como Anulada
                            Dim queryUpdateVenta As String = "UPDATE `ventas` SET `estado` = 'Anulada', `observaciones` = (IFNULL(`observaciones`,'') || ' | ANULADA: ' || @motivo) WHERE `id` = @vId;"
                            Using cmdUpdVenta = DatabaseHelper.CreateCommand(conn, queryUpdateVenta, trans)
                                DatabaseHelper.AddParam(cmdUpdVenta, "@motivo", motivo)
                                DatabaseHelper.AddParam(cmdUpdVenta, "@vId", ventaId)
                                cmdUpdVenta.ExecuteNonQuery()
                            End Using

                            ' Reversar impacto en caja
                            Dim cajaId As Integer = Convert.ToInt32(ventaRow("caja_id"))
                            Dim total As Decimal = Convert.ToDecimal(ventaRow("total"))
                            Dim metodo As String = ventaRow("metodo_pago").ToString()

                            If cajaId > 0 Then
                                Dim updateCajaQuery As String
                                If metodo.Equals("Efectivo", StringComparison.OrdinalIgnoreCase) Then
                                    updateCajaQuery = "UPDATE `cajas` SET `total_ventas_efectivo` = `total_ventas_efectivo` - @tot, `monto_esperado` = `monto_esperado` - @tot WHERE `id` = @cajaId;"
                                Else
                                    updateCajaQuery = "UPDATE `cajas` SET `total_ventas_digital` = `total_ventas_digital` - @tot WHERE `id` = @cajaId;"
                                End If

                                Using cmdCaja = DatabaseHelper.CreateCommand(conn, updateCajaQuery, trans)
                                    DatabaseHelper.AddParam(cmdCaja, "@tot", total)
                                    DatabaseHelper.AddParam(cmdCaja, "@cajaId", cajaId)
                                    cmdCaja.ExecuteNonQuery()
                                End Using
                            End If

                            trans.Commit()
                            errorMessage = String.Empty
                            Return True
                        Catch exTrans As Exception
                            trans.Rollback()
                            errorMessage = exTrans.Message
                            Return False
                        End Try
                    End Using
                End Using
            Catch ex As Exception
                errorMessage = ex.Message
                Return False
            End Try
        End Function

    End Class
End Namespace
