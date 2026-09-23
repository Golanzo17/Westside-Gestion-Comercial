' ARCHIVO: ConfiguracionService.vb
' PROPÓSITO: Persistencia de datos institucionales de la empresa y pie de ticket.
' Administra los datos del comercio que luego se imprimen en los comprobantes (Nombre,
' CUIT, Dirección, Teléfono, Condición de IVA y mensaje de cambio de prendas).
' Para persistir estos datos usamos la técnica UPSERT (Insert or Update):
' Sintaxis nativa SQLite "ON CONFLICT(id) DO UPDATE SET...".
' De este modo garantizamos que siempre exista una única fila (id = 1) sin duplicados.


Imports System.Data
Imports GestionComercial.Data
Imports GestionComercial.Models

Namespace Services
    Public Class ConfiguracionService

        ' Recupera la configuración del local; si no existe aún, devuelve valores predeterminados
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

        ' Guarda los datos comerciales asegurando una única fila mediante UPSERT.

        Public Function GuardarConfiguracion(cfg As ConfiguracionComercio, ByRef errorMessage As String) As Boolean
            Try
                Dim query As String = "INSERT INTO `configuracion` (`id`, `nombre_comercio`, `cuit`, `direccion`, `telefono`, `email`, `condicion_iva`, `mensaje_ticket`, `moneda_simbolo`) " &
                                      "VALUES (1, @nombre, @cuit, @dir, @tel, @email, @iva, @mensaje, @moneda) " &
                                      "ON CONFLICT(`id`) DO UPDATE SET " &
                                      "`nombre_comercio` = @nombre, `cuit` = @cuit, `direccion` = @dir, `telefono` = @tel, `email` = @email, `condicion_iva` = @iva, `mensaje_ticket` = @mensaje, `moneda_simbolo` = @moneda;"

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
