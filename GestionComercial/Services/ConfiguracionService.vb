Imports System.Data
Imports GestionComercial.Data
Imports GestionComercial.Models

Namespace Services
    Public Class ConfiguracionService

        Public Function GetConfiguracion() As ConfiguracionComercio
            Dim config As New ConfiguracionComercio()
            Try
                Dim query As String = "SELECT * FROM `configuracion` LIMIT 1;"
                Dim dt As DataTable = DatabaseHelper.ExecuteQuery(query)
                If dt.Rows.Count > 0 Then
                    Dim row As DataRow = dt.Rows(0)
                    config.Id = Convert.ToInt32(row("id"))
                    config.NombreComercio = row("nombre_comercio").ToString()
                    config.Cuit = row("cuit").ToString()
                    config.Direccion = row("direccion").ToString()
                    config.Telefono = row("telefono").ToString()
                    config.Email = row("email").ToString()
                    config.CondicionIva = row("condicion_iva").ToString()
                    config.MensajeTicket = row("mensaje_ticket").ToString()
                    config.MonedaSimbolo = row("moneda_simbolo").ToString()
                End If
            Catch ex As Exception
                ' Retorna defaults
            End Try
            Return config
        End Function

        Public Function GuardarConfiguracion(cfg As ConfiguracionComercio, ByRef errorMessage As String) As Boolean
            Try
                Dim query As String
                If DatabaseHelper.IsSQLite Then
                    query = "INSERT INTO `configuracion` (`id`, `nombre_comercio`, `cuit`, `direccion`, `telefono`, `email`, `condicion_iva`, `mensaje_ticket`, `moneda_simbolo`) " &
                            "VALUES (1, @nombre, @cuit, @dir, @tel, @email, @iva, @mensaje, @moneda) " &
                            "ON CONFLICT(`id`) DO UPDATE SET " &
                            "`nombre_comercio` = @nombre, `cuit` = @cuit, `direccion` = @dir, `telefono` = @tel, `email` = @email, `condicion_iva` = @iva, `mensaje_ticket` = @mensaje, `moneda_simbolo` = @moneda;"
                Else
                    query = "INSERT INTO `configuracion` (`id`, `nombre_comercio`, `cuit`, `direccion`, `telefono`, `email`, `condicion_iva`, `mensaje_ticket`, `moneda_simbolo`) " &
                            "VALUES (1, @nombre, @cuit, @dir, @tel, @email, @iva, @mensaje, @moneda) " &
                            "ON DUPLICATE KEY UPDATE " &
                            "`nombre_comercio` = @nombre, `cuit` = @cuit, `direccion` = @dir, `telefono` = @tel, `email` = @email, `condicion_iva` = @iva, `mensaje_ticket` = @mensaje, `moneda_simbolo` = @moneda;"
                End If

                Dim params As New Dictionary(Of String, Object) From {
                    {"@nombre", cfg.NombreComercio},
                    {"@cuit", cfg.Cuit},
                    {"@dir", cfg.Direccion},
                    {"@tel", cfg.Telefono},
                    {"@email", cfg.Email},
                    {"@iva", cfg.CondicionIva},
                    {"@mensaje", cfg.MensajeTicket},
                    {"@moneda", cfg.MonedaSimbolo}
                }

                DatabaseHelper.ExecuteNonQuery(query, params)
                errorMessage = String.Empty
                Return True
            Catch ex As Exception
                errorMessage = ex.Message
                Return False
            End Try
        End Function

    End Class
End Namespace
