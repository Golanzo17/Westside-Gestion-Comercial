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

        ''' <summary>
        ''' Verifica si ya existe otro cliente con el mismo DNI/CUIT.
        ''' </summary>
        Public Function ExisteDni(dni As String, Optional excludeId As Integer = 0) As Boolean
            Dim dummy As Cliente = Nothing
            Return ExisteDni(dni, excludeId, dummy)
        End Function

        ''' <summary>
        ''' Verifica si ya existe otro cliente con el mismo DNI/CUIT y obtiene el cliente existente si coincide.
        ''' </summary>
        Public Function ExisteDni(dni As String, excludeId As Integer, ByRef clienteExistente As Cliente) As Boolean
            clienteExistente = Nothing
            Dim dniTrimmed = If(dni, "").Trim()
            If String.IsNullOrWhiteSpace(dniTrimmed) Then Return False

            Dim sqlCheck As String = "SELECT * FROM `clientes` WHERE LOWER(TRIM(`dni_cuit`)) = LOWER(@dni) AND `id` <> @id LIMIT 1;"
            Dim dtCheck = DatabaseHelper.ExecuteQuery(sqlCheck, New Dictionary(Of String, Object) From {
                {"@dni", dniTrimmed},
                {"@id", excludeId}
            })

            If dtCheck.Rows.Count > 0 Then
                clienteExistente = MapCliente(dtCheck.Rows(0))
                Return True
            End If
            Return False
        End Function

        Public Function GuardarCliente(cli As Cliente, ByRef errorMessage As String) As Boolean
            Try
                Dim dniTrimmed = If(cli.DniCuit, "").Trim()
                If String.IsNullOrWhiteSpace(dniTrimmed) Then
                    errorMessage = "El DNI/CUIT es obligatorio."
                    Return False
                End If

                If String.IsNullOrWhiteSpace(cli.Nombre) Then
                    errorMessage = "El nombre del cliente es obligatorio."
                    Return False
                End If

                ' Control preventivo de unicidad de DNI
                Dim clienteDuplicado As Cliente = Nothing
                If ExisteDni(dniTrimmed, cli.Id, clienteDuplicado) Then
                    Dim estadoDesc = If(clienteDuplicado.Activo, "activo", "dado de baja")
                    If clienteDuplicado.Activo Then
                        errorMessage = $"Ya existe un cliente registrado con el DNI/CUIT '{dniTrimmed}': {clienteDuplicado.NombreCompleto}."
                    Else
                        errorMessage = $"El DNI/CUIT '{dniTrimmed}' pertenece al cliente {clienteDuplicado.NombreCompleto}, actualmente inactivo. Puede reactivarlo desde el listado."
                    End If
                    Return False
                End If

                Dim query As String
                Dim params As New Dictionary(Of String, Object) From {
                    {"@dni",   dniTrimmed},
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
