Imports System.Data
Imports MySqlConnector
Imports GestionComercial.Data
Imports GestionComercial.Models

Namespace Services
    Public Class CajaService

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

                            ' Actualizar totales en la caja
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

        Public Function CerrarCaja(cajaId As Integer, montoReal As Decimal, observaciones As String, ByRef errorMessage As String) As Boolean
            Try
                ' Calcular monto esperado final
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
