' ARCHIVO: CajaService.vb
' PROPÓSITO: Lógica de negocio de turnos de caja, arqueo de efectivo y control de gastos.
' En este servicio resolvemos el control financiero diario del salón de ventas:
' 1. Apertura de Turno: Exige registrar el fondo inicial para cambio antes de vender.
' 2. Movimientos Manuales en Transacción: Registra retiros (egresos) o ingresos extraordinarios
'    actualizando de inmediato el dinero que debería haber en el cajón.
' 3. Arqueo y Cierre Ciego: Cuando el cajero cuenta el dinero físico real (montoReal),
'    el sistema calcula automáticamente la diferencia (Monto Real - Monto Esperado),
'    dejando registro transparente de si hubo sobrante o faltante de dinero en el turno.

Imports System.Data
Imports GestionComercial.Data
Imports GestionComercial.Models

Namespace Services
    Public Class CajaService

        ' Busca si existe una caja con estado "Abierta".

        Public Function GetCajaAbierta(usuarioId As Integer) As Caja
            Dim query As String = "SELECT c.*, u.nombre_completo AS usuario_nombre " &
                                 "FROM `cajas` c " &
                                 "INNER JOIN `usuarios` u ON c.usuario_id = u.id " &
                                 "WHERE c.estado = 'Abierta' " &
                                 "ORDER BY c.id DESC LIMIT 1;"

            Dim dt As DataTable = DatabaseHelper.ExecuteQuery(query)
            If dt.Rows.Count = 0 Then Return Nothing

            Dim row As DataRow = dt.Rows(0)
            Return MapCaja(row)
        End Function

        ' Abre un nuevo turno de caja con el monto de cambio inicial.

        Public Function AbrirCaja(usuarioId As Integer, montoInicial As Decimal, ByRef errorMessage As String) As Caja
            Try
                Dim cajaExistente As Caja = GetCajaAbierta(usuarioId)
                If cajaExistente IsNot Nothing Then
                    errorMessage = "Ya existe una caja abierta en este momento (Caja #" & cajaExistente.Id & ")."
                    Return Nothing
                End If

                Dim query As String = "INSERT INTO `cajas` (`usuario_id`, `fecha_apertura`, `monto_inicial`, `monto_esperado`, `estado`) " &
                                     "VALUES (@userId, NOW(), @montoIni, @montoIni, 'Abierta');"

                Dim params As New Dictionary(Of String, Object) From {
                    {"@userId", usuarioId},
                    {"@montoIni", montoInicial}
                }

                DatabaseHelper.ExecuteNonQuery(query, params)
                errorMessage = String.Empty
                Return GetCajaAbierta(usuarioId)
            Catch ex As Exception
                errorMessage = ex.Message
                Return Nothing
            End Try
        End Function

        ' Movimientos Manuales con Transacción Atómica.
        ' Si el vendedor paga un flete o saca cambio, guardamos el comprobante en
        ' movimientos_caja y recalculamos monto_esperado en cajas dentro de la misma transacción.
        Public Function RegistrarMovimiento(cajaId As Integer, usuarioId As Integer, tipo As String, concepto As String, monto As Decimal, referencia As String, ByRef errorMessage As String) As Boolean
            Try
                Using conn As Common.DbConnection = DatabaseHelper.GetConnection()
                    conn.Open()
                    Using trans As Common.DbTransaction = conn.BeginTransaction()
                        Try
                            Dim queryMov As String = "INSERT INTO `movimientos_caja` (`caja_id`, `usuario_id`, `fecha`, `tipo`, `concepto`, `monto`, `referencia`) " &
                                                    "VALUES (@cajaId, @userId, NOW(), @tipo, @concepto, @monto, @ref);"
                            Using cmdMov = DatabaseHelper.CreateCommand(conn, queryMov, trans)
                                DatabaseHelper.AddParam(cmdMov, "@cajaId", cajaId)
                                DatabaseHelper.AddParam(cmdMov, "@userId", usuarioId)
                                DatabaseHelper.AddParam(cmdMov, "@tipo", tipo)
                                DatabaseHelper.AddParam(cmdMov, "@concepto", concepto)
                                DatabaseHelper.AddParam(cmdMov, "@monto", monto)
                                DatabaseHelper.AddParam(cmdMov, "@ref", referencia)
                                cmdMov.ExecuteNonQuery()
                            End Using

                            ' Actualizamos los acumuladores y el dinero esperado en la caja
                            Dim updateCol As String = If(tipo = "Ingreso", "`total_ingresos` = `total_ingresos` + @monto", "`total_egresos` = `total_egresos` + @monto")
                            Dim updateCaja As String = $"UPDATE `cajas` SET {updateCol}, `monto_esperado` = `monto_inicial` + `total_ventas_efectivo` + `total_ingresos` - `total_egresos` WHERE `id` = @cajaId;"
                            Using cmdCaja = DatabaseHelper.CreateCommand(conn, updateCaja, trans)
                                DatabaseHelper.AddParam(cmdCaja, "@monto", monto)
                                DatabaseHelper.AddParam(cmdCaja, "@cajaId", cajaId)
                                cmdCaja.ExecuteNonQuery()
                            End Using

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

        ' Arqueo y Cierre de Turno.
        ' El cajero ingresa el dinero físico que contó en mano (montoReal).
        ' El sistema calcula: Diferencia = Monto Real - Monto Esperado
        ' Si es 0: Arqueo exacto. Si es positivo: Sobrante. Si es negativo: Faltante.
        Public Function CerrarCaja(cajaId As Integer, montoReal As Decimal, observaciones As String, ByRef errorMessage As String) As Boolean
            Try
                Dim queryGet As String = "SELECT * FROM `cajas` WHERE `id` = @cajaId LIMIT 1;"
                Dim dt As DataTable = DatabaseHelper.ExecuteQuery(queryGet, New Dictionary(Of String, Object) From {{"@cajaId", cajaId}})
                If dt.Rows.Count = 0 Then
                    errorMessage = "Caja no encontrada."
                    Return False
                End If

                Dim row As DataRow = dt.Rows(0)
                Dim inicial As Decimal = Convert.ToDecimal(row("monto_inicial"))
                Dim ventasEfec As Decimal = Convert.ToDecimal(row("total_ventas_efectivo"))
                Dim ingresos As Decimal = Convert.ToDecimal(row("total_ingresos"))
                Dim egresos As Decimal = Convert.ToDecimal(row("total_egresos"))
                Dim esperado As Decimal = inicial + ventasEfec + ingresos - egresos
                Dim diferencia As Decimal = montoReal - esperado

                Dim queryClose As String = "UPDATE `cajas` SET `fecha_cierre` = NOW(), `monto_esperado` = @esp, `monto_real` = @real, " &
                                          "`diferencia` = @dif, `estado` = 'Cerrada', `observaciones` = @obs WHERE `id` = @cajaId;"

                Dim params As New Dictionary(Of String, Object) From {
                    {"@esp", esperado},
                    {"@real", montoReal},
                    {"@dif", diferencia},
                    {"@obs", observaciones},
                    {"@cajaId", cajaId}
                }

                DatabaseHelper.ExecuteNonQuery(queryClose, params)
                errorMessage = String.Empty
                Return True
            Catch ex As Exception
                errorMessage = ex.Message
                Return False
            End Try
        End Function

        Public Function GetMovimientosCaja(cajaId As Integer) As List(Of MovimientoCaja)
            Dim list As New List(Of MovimientoCaja)()
            Dim query As String = "SELECT mc.*, u.nombre_completo AS usuario_nombre " &
                                 "FROM `movimientos_caja` mc " &
                                 "INNER JOIN `usuarios` u ON mc.usuario_id = u.id " &
                                 "WHERE mc.caja_id = @cajaId ORDER BY mc.fecha DESC;"

            Dim dt As DataTable = DatabaseHelper.ExecuteQuery(query, New Dictionary(Of String, Object) From {{"@cajaId", cajaId}})
            For Each row As DataRow In dt.Rows
                list.Add(New MovimientoCaja() With {
                    .Id = Convert.ToInt32(row("id")),
                    .CajaId = Convert.ToInt32(row("caja_id")),
                    .UsuarioId = Convert.ToInt32(row("usuario_id")),
                    .UsuarioNombre = row("usuario_nombre").ToString(),
                    .Fecha = Convert.ToDateTime(row("fecha")),
                    .Tipo = row("tipo").ToString(),
                    .Concepto = row("concepto").ToString(),
                    .Monto = Convert.ToDecimal(row("monto")),
                    .Referencia = If(IsDBNull(row("referencia")), "", row("referencia").ToString())
                })
            Next
            Return list
        End Function

        Public Function GetCajaById(id As Integer) As Caja
            Try
                Dim query As String = "SELECT c.*, u.nombre_completo AS usuario_nombre " &
                                     "FROM `cajas` c " &
                                     "INNER JOIN `usuarios` u ON c.usuario_id = u.id " &
                                     "WHERE c.id = @id LIMIT 1;"
                Dim dt As DataTable = DatabaseHelper.ExecuteQuery(query, New Dictionary(Of String, Object) From {{"@id", id}})
                If dt.Rows.Count = 0 Then Return Nothing
                Return MapCaja(dt.Rows(0))
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        Public Function GetHistorialCajas(fechaDesde As DateTime, fechaHasta As DateTime) As List(Of Caja)
            Return GetHistorialCajasAvanzado(fechaDesde, fechaHasta, Nothing, Nothing, Nothing)
        End Function

        ''' <summary>
        ''' Búsqueda y filtrado exhaustivo de turnos de caja para la auditoría gerencial.
        ''' Permite filtrar por rango de fechas, vendedor individual, estado de apertura y tipo de diferencia de arqueo.
        ''' </summary>
        Public Function GetHistorialCajasAvanzado(fechaDesde As DateTime, fechaHasta As DateTime, usuarioId As Nullable(Of Integer), estado As String, filtroDiferencia As String) As List(Of Caja)
            Dim list As New List(Of Caja)()
            Try
                Dim sql As New System.Text.StringBuilder()
                sql.Append("SELECT c.*, u.nombre_completo AS usuario_nombre ")
                sql.Append("FROM `cajas` c ")
                sql.Append("INNER JOIN `usuarios` u ON c.usuario_id = u.id ")
                sql.Append("WHERE DATE(c.fecha_apertura) >= @desde AND DATE(c.fecha_apertura) <= @hasta ")

                Dim params As New Dictionary(Of String, Object) From {
                    {"@desde", fechaDesde.ToString("yyyy-MM-dd")},
                    {"@hasta", fechaHasta.ToString("yyyy-MM-dd")}
                }

                If usuarioId.HasValue AndAlso usuarioId.Value > 0 Then
                    sql.Append("AND c.usuario_id = @usuarioId ")
                    params.Add("@usuarioId", usuarioId.Value)
                End If

                If Not String.IsNullOrWhiteSpace(estado) AndAlso Not estado.Equals("Todos", StringComparison.OrdinalIgnoreCase) Then
                    sql.Append("AND c.estado = @estado ")
                    params.Add("@estado", estado)
                End If

                If Not String.IsNullOrWhiteSpace(filtroDiferencia) Then
                    Select Case filtroDiferencia.ToUpperInvariant()
                        Case "CON_DESCUADRE", "DESCUADRE"
                            sql.Append("AND c.diferencia IS NOT NULL AND c.diferencia <> 0 ")
                        Case "FALTANTE", "CON_FALTANTE"
                            sql.Append("AND c.diferencia IS NOT NULL AND c.diferencia < 0 ")
                        Case "SOBRANTE", "CON_SOBRANTE"
                            sql.Append("AND c.diferencia IS NOT NULL AND c.diferencia > 0 ")
                        Case "EXACTO", "EXACTAS"
                            sql.Append("AND c.diferencia IS NOT NULL AND c.diferencia = 0 ")
                        Case "ABIERTAS"
                            sql.Append("AND c.estado = 'Abierta' ")
                    End Select
                End If

                sql.Append("ORDER BY c.id DESC;")

                Dim dt As DataTable = DatabaseHelper.ExecuteQuery(sql.ToString(), params)
                For Each row As DataRow In dt.Rows
                    list.Add(MapCaja(row))
                Next
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("Error en GetHistorialCajasAvanzado: " & ex.Message)
            End Try
            Return list
        End Function

        ''' <summary>
        ''' Recupera todas las ventas y tickets facturados en un turno de caja específico.
        ''' </summary>
        Public Function GetVentasPorCaja(cajaId As Integer) As List(Of Venta)
            Dim list As New List(Of Venta)()
            Try
                Dim query As String = "SELECT v.*, u.nombre_completo AS usuario_nombre, c.nombre_completo AS cliente_nombre " &
                                     "FROM `ventas` v " &
                                     "INNER JOIN `usuarios` u ON v.usuario_id = u.id " &
                                     "INNER JOIN `clientes` c ON v.cliente_id = c.id " &
                                     "WHERE v.caja_id = @cajaId " &
                                     "ORDER BY v.fecha DESC;"

                Dim dt As DataTable = DatabaseHelper.ExecuteQuery(query, New Dictionary(Of String, Object) From {{"@cajaId", cajaId}})
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
                        .Observaciones = If(Convert.IsDBNull(row("observaciones")), "", row("observaciones").ToString())
                    })
                Next
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("Error en GetVentasPorCaja: " & ex.Message)
            End Try
            Return list
        End Function

        ''' <summary>
        ''' Genera una conciliación de totales cobrados por cada medio de pago durante el turno auditado.
        ''' </summary>
        Public Function GetResumenMediosPagoCaja(cajaId As Integer) As DataTable
            Try
                Dim query As String = "SELECT `metodo_pago` AS MetodoPago, COUNT(*) AS CantidadTickets, " &
                                     "SUM(`total`) AS TotalCobrado " &
                                     "FROM `ventas` " &
                                     "WHERE `caja_id` = @cajaId AND `estado` <> 'Anulada' " &
                                     "GROUP BY `metodo_pago` " &
                                     "ORDER BY TotalCobrado DESC;"

                Return DatabaseHelper.ExecuteQuery(query, New Dictionary(Of String, Object) From {{"@cajaId", cajaId}})
            Catch ex As Exception
                Return New DataTable()
            End Try
        End Function

        Private Function MapCaja(row As DataRow) As Caja
            Return New Caja() With {
                .Id = Convert.ToInt32(row("id")),
                .UsuarioId = Convert.ToInt32(row("usuario_id")),
                .UsuarioNombre = row("usuario_nombre").ToString(),
                .FechaApertura = Convert.ToDateTime(row("fecha_apertura")),
                .FechaCierre = If(IsDBNull(row("fecha_cierre")), Nothing, Convert.ToDateTime(row("fecha_cierre"))),
                .MontoInicial = Convert.ToDecimal(row("monto_inicial")),
                .TotalVentasEfectivo = Convert.ToDecimal(row("total_ventas_efectivo")),
                .TotalVentasDigital = Convert.ToDecimal(row("total_ventas_digital")),
                .TotalIngresos = Convert.ToDecimal(row("total_ingresos")),
                .TotalEgresos = Convert.ToDecimal(row("total_egresos")),
                .MontoEsperado = Convert.ToDecimal(row("monto_esperado")),
                .MontoReal = If(IsDBNull(row("monto_real")), Nothing, Convert.ToDecimal(row("monto_real"))),
                .Diferencia = If(IsDBNull(row("diferencia")), Nothing, Convert.ToDecimal(row("diferencia"))),
                .Estado = row("estado").ToString(),
                .Observaciones = If(IsDBNull(row("observaciones")), "", row("observaciones").ToString())
            }
        End Function

    End Class
End Namespace
