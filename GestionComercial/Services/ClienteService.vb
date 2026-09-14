Imports System.Data
Imports GestionComercial.Data
Imports GestionComercial.Models

Namespace Services
    Public Class ClienteService

        Public Function GetClientes(Optional busqueda As String = "", Optional soloActivos As Boolean = True) As List(Of Cliente)
            Dim list As New List(Of Cliente)()
            Dim query As String = "SELECT * FROM `clientes` WHERE 1=1 "
            Dim params As New Dictionary(Of String, Object)()

            If soloActivos Then
                query &= "AND `activo` = 1 "
            End If

            If Not String.IsNullOrWhiteSpace(busqueda) Then
                query &= "AND (`dni_cuit` LIKE @b OR `nombre` LIKE @b OR `apellido` LIKE @b OR `telefono` LIKE @b) "
                params.Add("@b", "%" & busqueda.Trim() & "%")
            End If

            query &= "ORDER BY `apellido` ASC, `nombre` ASC;"

            Dim dt As DataTable = DatabaseHelper.ExecuteQuery(query, params)
            For Each row As DataRow In dt.Rows
                list.Add(MapCliente(row))
            Next
            Return list
        End Function

        Public Function GetClienteById(id As Integer) As Cliente
            Dim query As String = "SELECT * FROM `clientes` WHERE `id` = @id LIMIT 1;"
            Dim dt As DataTable = DatabaseHelper.ExecuteQuery(query, New Dictionary(Of String, Object) From {{"@id", id}})
            If dt.Rows.Count = 0 Then Return Nothing
            Return MapCliente(dt.Rows(0))
        End Function

        Private Function MapCliente(row As DataRow) As Cliente
            Dim fnac As Nullable(Of DateTime) = Nothing
            If Not IsDBNull(row("fecha_nacimiento")) Then
                Dim parsed As DateTime
                If DateTime.TryParse(row("fecha_nacimiento").ToString(), parsed) Then
                    fnac = parsed
                End If
            End If
            Return New Cliente() With {
                .Id = Convert.ToInt32(row("id")),
                .DniCuit = row("dni_cuit").ToString(),
                .Nombre = row("nombre").ToString(),
                .Apellido = row("apellido").ToString(),
                .Telefono = If(IsDBNull(row("telefono")), "", row("telefono").ToString()),
                .Email = If(IsDBNull(row("email")), "", row("email").ToString()),
                .Direccion = If(IsDBNull(row("direccion")), "", row("direccion").ToString()),
                .Ciudad = If(IsDBNull(row("ciudad")), "", row("ciudad").ToString()),
                .Notas = If(IsDBNull(row("notas")), "", row("notas").ToString()),
                .FechaNacimiento = fnac,
                .Activo = Convert.ToBoolean(row("activo"))
            }
        End Function

        Public Function GuardarCliente(cli As Cliente, ByRef errorMessage As String) As Boolean
            Try
                Dim query As String
                Dim params As New Dictionary(Of String, Object) From {
                    {"@dni",   cli.DniCuit.Trim()},
                    {"@nom",   cli.Nombre.Trim()},
                    {"@ape",   cli.Apellido.Trim()},
                    {"@tel",   cli.Telefono.Trim()},
                    {"@email", cli.Email.Trim()},
                    {"@dir",   cli.Direccion.Trim()},
                    {"@ciu",   cli.Ciudad.Trim()},
                    {"@notas", cli.Notas.Trim()},
                    {"@fnac",  If(cli.FechaNacimiento.HasValue, CObj(cli.FechaNacimiento.Value.ToString("yyyy-MM-dd")), DBNull.Value)},
                    {"@activo", If(cli.Activo, 1, 0)}
                }

                If cli.Id = 0 Then
                    query = "INSERT INTO `clientes` (`dni_cuit`, `nombre`, `apellido`, `telefono`, `email`, `direccion`, `ciudad`, `notas`, `fecha_nacimiento`, `activo`) " &
                            "VALUES (@dni, @nom, @ape, @tel, @email, @dir, @ciu, @notas, @fnac, @activo);"
                Else
                    query = "UPDATE `clientes` SET `dni_cuit` = @dni, `nombre` = @nom, `apellido` = @ape, `telefono` = @tel, " &
                            "`email` = @email, `direccion` = @dir, `ciudad` = @ciu, `notas` = @notas, " &
                            "`fecha_nacimiento` = @fnac, `activo` = @activo WHERE `id` = @id;"
                    params.Add("@id", cli.Id)
                End If

                DatabaseHelper.ExecuteNonQuery(query, params)
                errorMessage = String.Empty
                Return True
            Catch ex As Exception
                errorMessage = ex.Message
                Return False
            End Try
        End Function

        Public Function EliminarCliente(id As Integer, ByRef errorMessage As String) As Boolean
            Try
                Dim query As String = "UPDATE `clientes` SET `activo` = 0 WHERE `id` = @id;"
                DatabaseHelper.ExecuteNonQuery(query, New Dictionary(Of String, Object) From {{"@id", id}})
                errorMessage = String.Empty
                Return True
            Catch ex As Exception
                errorMessage = ex.Message
                Return False
            End Try
        End Function

    End Class
End Namespace
