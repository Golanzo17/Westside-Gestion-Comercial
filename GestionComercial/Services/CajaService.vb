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
Imports System.Data.Common
Imports GestionComercial.Data
Imports GestionComercial.Models

Namespace Services
    Public Class CajaService

        ' Busca si existe una caja con estado "Abierta".
        ' Si se indica usuarioId (> 0), prioriza la caja abierta por ese usuario y, si no tiene una propia,
        ' devuelve la caja abierta activa del salón comercial.
        Public Function GetCajaAbierta(Optional usuarioId As Integer = 0) As Caja
            If usuarioId > 0 Then
                Dim queryUsuario As String = "SELECT c.*, u.nombre_completo AS usuario_nombre " &
                                             "FROM `cajas` c " &
                                             "INNER JOIN `usuarios` u ON c.usuario_id = u.id " &
                                             "WHERE c.estado = 'Abierta' AND c.usuario_id = @uid " &
                                             "ORDER BY c.id DESC LIMIT 1;"
                Dim dtUser = DatabaseHelper.ExecuteQuery(queryUsuario, New Dictionary(Of String, Object) From {{"@uid", usuarioId}})
                If dtUser.Rows.Count > 0 Then
                    Return MapCaja(dtUser.Rows(0))
                End If
            End If

            Dim query As String = "SELECT c.*, u.nombre_completo AS usuario_nombre " &
                                  "FROM `cajas` c " &
                                  "INNER JOIN `usuarios` u ON c.usuario_id = u.id " &
                                  "WHERE c.estado = 'Abierta' " &
                                  "ORDER BY c.id DESC LIMIT 1;"

            Dim dt As DataTable = DatabaseHelper.ExecuteQuery(query)
            If dt.Rows.Count = 0 Then Return Nothing

            Return MapCaja(dt.Rows(0))
        End Function

        ' Abre un nuevo turno de caja con el monto de cambio inicial.
        Public Function AbrirCaja(usuarioId As Integer, montoInicial As Decimal, ByRef errorMessage As String) As Caja
            Try
                Dim cajaExistente As Caja = GetCajaAbierta(usuarioId)
                If cajaExistente IsNot Nothing Then
                    errorMessage = "Ya existe una caja abierta en este momento (Caja #" & cajaExistente.Id & ")."
                    Return Nothing
                End If

                Dim montoRedondeado = Math.Round(montoInicial, 2)
                Dim query As String = "INSERT INTO `cajas` (`usuario_id`, `fecha_apertura`, `monto_inicial`, `monto_esperado`, `estado`) " &
                                     "VALUES (@userId, NOW(), @montoIni, @montoIni, 'Abierta');"

                Dim params As New Dictionary(Of String, Object) From {
                    {"@userId", usuarioId},
                    {"@montoIni", montoRedondeado}
                }

                DatabaseHelper.ExecuteNonQuery(query, params)
                AppLogger.LogAudit("CAJA_ABIERTA", $"UsuarioId={usuarioId} | MontoInicial={montoRedondeado:F2}")
                errorMessage = String.Empty
                Return GetCajaAbierta(usuarioId)
            Catch ex As Exception
                AppLogger.LogError("Error en AbrirCaja", ex)
                errorMessage = ex.Message
                Return Nothing
            End Try
        End Function

        ' Movimientos Manuales con Transacción Atómica.
        ' Si el vendedor paga un flete o saca cambio, guardamos el comprobante en
        ' movimientos_caja y recalculamos monto_esperado en cajas dentro de la misma transacción.
        Public Function RegistrarMovimiento(cajaId As Integer, usuarioId As Integer, tipo As String, concepto As String, monto As Decimal, referencia As String, ByRef errorMessage As String, Optional autorizadoPorId As Nullable(Of Integer) = Nothing) As Boolean
            Try
                Dim montoRedondeado = Math.Round(monto, 2)
                Using conn As DbConnection = DatabaseHelper.GetConnection()
                    conn.Open()
                    Using trans As DbTransaction = conn.BeginTransaction()
                        Try
                            Dim queryMov As String = "INSERT INTO `movimientos_caja` (`caja_id`, `usuario_id`, `autorizado_por_id`, `fecha`, `tipo`, `concepto`, `monto`, `referencia`) " &
                                                    "VALUES (@cajaId, @userId, @autId, NOW(), @tipo, @concepto, @monto, @ref);"
                            Using cmdMov = DatabaseHelper.CreateCommand(conn, queryMov, trans)
                                DatabaseHelper.AddParam(cmdMov, "@cajaId", cajaId)
                                DatabaseHelper.AddParam(cmdMov, "@userId", usuarioId)
                                DatabaseHelper.AddParam(cmdMov, "@autId", If(autorizadoPorId.HasValue, CObj(autorizadoPorId.Value), DBNull.Value))
                                DatabaseHelper.AddParam(cmdMov, "@tipo", tipo)
                                DatabaseHelper.AddParam(cmdMov, "@concepto", concepto)
                                DatabaseHelper.AddParam(cmdMov, "@monto", montoRedondeado)
                                DatabaseHelper.AddParam(cmdMov, "@ref", referencia)
                                cmdMov.ExecuteNonQuery()
                            End Using

                            ' Actualizamos los acumuladores y el dinero esperado en la caja
                            Dim updateCol As String = If(tipo.Equals("Ingreso", StringComparison.OrdinalIgnoreCase),
                                                         "`total_ingresos` = ROUND(`total_ingresos` + @monto, 2)",
                                                         "`total_egresos` = ROUND(`total_egresos` + @monto, 2)")
                            Dim updateCaja As String = $"UPDATE `cajas` SET {updateCol}, `monto_esperado` = ROUND(`monto_inicial` + `total_ventas_efectivo` + `total_ingresos` - `total_egresos`, 2) WHERE `id` = @cajaId;"
                            Using cmdCaja = DatabaseHelper.CreateCommand(conn, updateCaja, trans)
                                DatabaseHelper.AddParam(cmdCaja, "@monto", montoRedondeado)
                                DatabaseHelper.AddParam(cmdCaja, "@cajaId", cajaId)
                                cmdCaja.ExecuteNonQuery()
                            End Using

                            trans.Commit()
                            AppLogger.LogAudit("MOVIMIENTO_CAJA", $"CajaId={cajaId} | Tipo={tipo} | Monto={montoRedondeado:F2} | UsuarioId={usuarioId} | AutorizadoPor={If(autorizadoPorId.HasValue, autorizadoPorId.Value.ToString(), "-")}")
                            errorMessage = String.Empty
                            Return True
                        Catch exTrans As Exception
                            trans.Rollback()
                            AppLogger.LogError("Error en transacción RegistrarMovimiento", exTrans)
                            errorMessage = exTrans.Message
                            Return False
                        End Try
                    End Using
                End Using
            Catch ex As Exception
                AppLogger.LogError("Error de conexión en RegistrarMovimiento", ex)
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
                Dim inicial As Decimal = Math.Round(Convert.ToDecimal(row("monto_inicial")), 2)
                Dim ventasEfec As Decimal = Math.Round(Convert.ToDecimal(row("total_ventas_efectivo")), 2)
                Dim ingresos As Decimal = Math.Round(Convert.ToDecimal(row("total_ingresos")), 2)
                Dim egresos As Decimal = Math.Round(Convert.ToDecimal(row("total_egresos")), 2)
                Dim esperado As Decimal = Math.Round(inicial + ventasEfec + ingresos - egresos, 2)
                Dim realRedondeado As Decimal = Math.Round(montoReal, 2)
                Dim diferencia As Decimal = Math.Round(realRedondeado - esperado, 2)

                Dim queryClose As String = "UPDATE `cajas` SET `fecha_cierre` = NOW(), `monto_esperado` = @esp, `monto_real` = @real, " &
                                          "`diferencia` = @dif, `estado` = 'Cerrada', `observaciones` = @obs WHERE `id` = @cajaId;"

                Dim params As New Dictionary(Of String, Object) From {
                    {"@esp", esperado},
                    {"@real", realRedondeado},
                    {"@dif", diferencia},
                    {"@obs", observaciones},
                    {"@cajaId", cajaId}
                }

                DatabaseHelper.ExecuteNonQuery(queryClose, params)
                AppLogger.LogAudit("CAJA_CERRADA", $"CajaId={cajaId} | Esperado={esperado:F2} | Real={realRedondeado:F2} | Diferencia={diferencia:F2}")
                errorMessage = String.Empty
                Return True
            Catch ex As Exception
                AppLogger.LogError("Error en CerrarCaja", ex)
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
                Dim autId As Nullable(Of Integer) = Nothing
                If row.Table.Columns.Contains("autorizado_por_id") AndAlso Not Convert.IsDBNull(row("autorizado_por_id")) Then
                    autId = Convert.ToInt32(row("autorizado_por_id"))
                End If

                list.Add(New MovimientoCaja() With {
                    .Id = Convert.ToInt32(row("id")),
                    .CajaId = Convert.ToInt32(row("caja_id")),
                    .UsuarioId = Convert.ToInt32(row("usuario_id")),
                    .UsuarioNombre = row("usuario_nombre").ToString(),
                    .AutorizadoPorId = autId,
                    .Fecha = Convert.ToDateTime(row("fecha")),
                    .Tipo = row("tipo").ToString(),
                    .Concepto = row("concepto").ToString(),
                    .Monto = Math.Round(Convert.ToDecimal(row("monto")), 2),
                    .Referencia = If(Convert.IsDBNull(row("referencia")), "", row("referencia").ToString())
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
                AppLogger.LogError("Error en GetCajaById", ex)
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
                AppLogger.LogError("Error en GetHistorialCajasAvanzado", ex)
            End Try
            Return list
        End Function

        ''' <summary>
        ''' Recupera todas las ventas y tickets facturados en un turno de caja específico.
        ''' </summary>
        Public Function GetVentasPorCaja(cajaId As Integer) As List(Of Venta)
            Dim list As New List(Of Venta)()
            Try
                Dim query As String = "SELECT v.*, u.nombre_completo AS usuario_nombre, (c.apellido || ', ' || c.nombre) AS cliente_nombre " &
                                     "FROM `ventas` v " &
                                     "INNER JOIN `usuarios` u ON v.usuario_id = u.id " &
                                     "INNER JOIN `clientes` c ON v.cliente_id = c.id " &
                                     "WHERE v.caja_id = @cajaId " &
                                     "ORDER BY v.fecha DESC;"

                Dim dt As DataTable = DatabaseHelper.ExecuteQuery(query, New Dictionary(Of String, Object) From {{"@cajaId", cajaId}})
                For Each row As DataRow In dt.Rows
                    list.Add( New Venta() With {
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
                        .Subtotal = Math.Round(Convert.ToDecimal(row("subtotal")), 2),
                        .DescuentoPorcentaje = Math.Round(Convert.ToDecimal(row("descuento_porcentaje")), 2),
                        .DescuentoMonto = Math.Round(Convert.ToDecimal(row("descuento_monto")), 2),
                        .RecargoMonto = Math.Round(Convert.ToDecimal(row("recargo_monto")), 2),
                        .Total = Math.Round(Convert.ToDecimal(row("total")), 2),
                        .MontoAbonado = Math.Round(Convert.ToDecimal(row("monto_abonado")), 2),
                        .Vuelto = Math.Round(Convert.ToDecimal(row("vuelto")), 2),
                        .Estado = row("estado").ToString(),
                        .Observaciones = If(Convert.IsDBNull(row("observaciones")), "", row("observaciones").ToString())
                    })
                Next
            Catch ex As Exception
                AppLogger.LogError("Error en GetVentasPorCaja", ex)
            End Try
            Return list
        End Function

        ''' <summary>
        ''' Genera una conciliación de totales cobrados por cada medio de pago durante el turno auditado.
        ''' </summary>
        Public Function GetResumenMediosPagoCaja(cajaId As Integer) As DataTable
            Try
                Dim query As String = "SELECT `metodo_pago` AS MetodoPago, COUNT(*) AS CantidadTickets, " &
                                     "ROUND(SUM(`total`), 2) AS TotalCobrado " &
                                     "FROM `ventas` " &
                                     "WHERE `caja_id` = @cajaId AND `estado` <> 'Anulada' " &
                                     "GROUP BY `metodo_pago` " &
                                     "ORDER BY TotalCobrado DESC;"

                Return DatabaseHelper.ExecuteQuery(query, New Dictionary(Of String, Object) From {{"@cajaId", cajaId}})
            Catch ex As Exception
                AppLogger.LogError("Error en GetResumenMediosPagoCaja", ex)
                Return New DataTable()
            End Try
        End Function

        Private Function MapCaja(row As DataRow) As Caja
            Dim fechaCierreVal As Nullable(Of DateTime) = Nothing
            If Not Convert.IsDBNull(row("fecha_cierre")) Then
                fechaCierreVal = Convert.ToDateTime(row("fecha_cierre"))
            End If

            Dim montoRealVal As Nullable(Of Decimal) = Nothing
            If Not Convert.IsDBNull(row("monto_real")) Then
                montoRealVal = Math.Round(Convert.ToDecimal(row("monto_real")), 2)
            End If

            Dim difVal As Nullable(Of Decimal) = Nothing
            If Not Convert.IsDBNull(row("diferencia")) Then
                difVal = Math.Round(Convert.ToDecimal(row("diferencia")), 2)
            End If

            Return New Caja() With {
                .Id = Convert.ToInt32(row("id")),
                .UsuarioId = Convert.ToInt32(row("usuario_id")),
                .UsuarioNombre = row("usuario_nombre").ToString(),
                .FechaApertura = Convert.ToDateTime(row("fecha_apertura")),
                .FechaCierre = fechaCierreVal,
                .MontoInicial = Math.Round(Convert.ToDecimal(row("monto_inicial")), 2),
                .TotalVentasEfectivo = Math.Round(Convert.ToDecimal(row("total_ventas_efectivo")), 2),
                .TotalVentasDigital = Math.Round(Convert.ToDecimal(row("total_ventas_digital")), 2),
                .TotalIngresos = Math.Round(Convert.ToDecimal(row("total_ingresos")), 2),
                .TotalEgresos = Math.Round(Convert.ToDecimal(row("total_egresos")), 2),
                .MontoEsperado = Math.Round(Convert.ToDecimal(row("monto_esperado")), 2),
                .MontoReal = montoRealVal,
                .Diferencia = difVal,
                .Estado = row("estado").ToString(),
                .Observaciones = If(Convert.IsDBNull(row("observaciones")), "", row("observaciones").ToString())
            }
        End Function

    End Class
End Namespace
