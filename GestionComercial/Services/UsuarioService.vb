Imports System.Data
Imports GestionComercial.Data
Imports GestionComercial.Models

Namespace Services
    Public Class UsuarioService

        Public Function GetUsuarios(Optional soloActivos As Boolean = False) As List(Of Usuario)
            Dim list As New List(Of Usuario)()
            Try
                Dim sql As String = "SELECT * FROM `usuarios` "
                If soloActivos Then
                    sql &= "WHERE `activo` = 1 "
                End If

                Dim dt As DataTable
                Try
                    dt = DatabaseHelper.ExecuteQuery(sql & "ORDER BY `apellido` ASC, `nombre` ASC;")
                Catch
                    dt = DatabaseHelper.ExecuteQuery(sql & "ORDER BY `id` ASC;")
                End Try

                For Each row As DataRow In dt.Rows
                    list.Add(MapUsuario(row))
                Next
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("Error al cargar usuarios: " & ex.Message)
            End Try
            Return list
        End Function

        Public Function GetUsuarioById(id As Integer) As Usuario
            Try
                Dim sql As String = "SELECT * FROM `usuarios` WHERE `id` = @id LIMIT 1;"
                Dim dt = DatabaseHelper.ExecuteQuery(sql, New Dictionary(Of String, Object) From {{"@id", id}})
                If dt.Rows.Count > 0 Then
                    Return MapUsuario(dt.Rows(0))
                End If
            Catch ex As Exception
                Return Nothing
            End Try
            Return Nothing
        End Function

        ''' <summary>
        ''' Verifica si un DNI ya pertenece a otro usuario en el sistema.
        ''' </summary>
        Public Function ExisteDni(dni As String, Optional excludeId As Integer = 0) As Boolean
            Dim dniTrimmed = If(dni, "").Trim()
            If String.IsNullOrWhiteSpace(dniTrimmed) Then Return False
            Dim sqlCheck As String = "SELECT COUNT(*) FROM `usuarios` WHERE LOWER(TRIM(`dni`)) = LOWER(@d) AND `id` <> @id;"
            Dim count = Convert.ToInt32(DatabaseHelper.ExecuteScalar(sqlCheck, New Dictionary(Of String, Object) From {
                {"@d", dniTrimmed},
                {"@id", excludeId}
            }))
            Return count > 0
        End Function

        Public Function CrearUsuario(username As String, password As String,
                                      nombre As String, apellido As String,
                                      rol As String, ByRef errorMessage As String,
                                      Optional dni As String = "",
                                      Optional telefono As String = "",
                                      Optional email As String = "",
                                      Optional direccion As String = "",
                                      Optional ciudad As String = "",
                                      Optional notas As String = "",
                                      Optional fechaNacimiento As Nullable(Of DateTime) = Nothing) As Boolean
            Try
                If String.IsNullOrWhiteSpace(username) Then
                    errorMessage = "El nombre de usuario es obligatorio."
                    Return False
                End If
                If String.IsNullOrWhiteSpace(password) OrElse password.Length < 4 Then
                    errorMessage = "La contraseña debe tener al menos 4 caracteres."
                    Return False
                End If
                If String.IsNullOrWhiteSpace(nombre) Then
                    errorMessage = "El nombre del empleado es obligatorio."
                    Return False
                End If
                If String.IsNullOrWhiteSpace(rol) Then
                    rol = "Vendedor"
                End If
                If Not EsRolValido(rol) Then
                    errorMessage = "El rol seleccionado no es válido."
                    Return False
                End If

                Dim dniTrimmed = If(dni, "").Trim()
                If String.IsNullOrWhiteSpace(dniTrimmed) Then
                    errorMessage = "El DNI del empleado es obligatorio."
                    Return False
                End If

                ' Verificar que el username no exista ya
                Dim sqlCheck As String = "SELECT COUNT(*) FROM `usuarios` WHERE LOWER(`username`) = LOWER(@u);"
                Dim count = Convert.ToInt32(DatabaseHelper.ExecuteScalar(sqlCheck, New Dictionary(Of String, Object) From {{"@u", username.Trim()}}))
                If count > 0 Then
                    errorMessage = $"El nombre de usuario '{username}' ya está en uso. Elija otro."
                    Return False
                End If

                ' Verificar que el DNI no exista ya
                If ExisteDni(dniTrimmed) Then
                    errorMessage = $"Ya existe otro usuario registrado con el DNI '{dniTrimmed}'."
                    Return False
                End If

                Dim hash = DatabaseHelper.HashPasswordSecure(password)
                Dim nombreCompleto As String = $"{apellido}, {nombre}".Trim(" "c, ","c)
                Dim sqlInsert As String =
                    "INSERT INTO `usuarios` (`username`, `dni`, `password_hash`, `nombre`, `apellido`, `nombre_completo`, `rol`, " &
                    "`telefono`, `email`, `direccion`, `ciudad`, `notas`, `fecha_nacimiento`, `activo`, `created_at`) " &
                    "VALUES (@u, @dni, @p, @nom, @ape, @nomcomp, @rol, @tel, @email, @dir, @ciu, @notas, @fnac, 1, NOW());"
                Dim prms As New Dictionary(Of String, Object) From {
                    {"@u",       username.Trim().ToLowerInvariant()},
                    {"@dni",     dniTrimmed},
                    {"@p",       hash},
                    {"@nom",     nombre.Trim()},
                    {"@ape",     apellido.Trim()},
                    {"@nomcomp", nombreCompleto},
                    {"@rol",     rol.Trim()},
                    {"@tel",     telefono.Trim()},
                    {"@email",   email.Trim()},
                    {"@dir",     direccion.Trim()},
                    {"@ciu",     ciudad.Trim()},
                    {"@notas",   notas.Trim()},
                    {"@fnac",    If(fechaNacimiento.HasValue, CObj(fechaNacimiento.Value.ToString("yyyy-MM-dd")), DBNull.Value)}
                }

                DatabaseHelper.ExecuteNonQuery(sqlInsert, prms)
                errorMessage = String.Empty
                Return True
            Catch ex As Exception
                errorMessage = "Error al crear usuario: " & ex.Message
                Return False
            End Try
        End Function

        Public Function ActualizarUsuario(usuarioId As Integer,
                                           nombre As String, apellido As String,
                                           rol As String, ByRef errorMessage As String,
                                           Optional dni As String = "",
                                           Optional telefono As String = "",
                                           Optional email As String = "",
                                           Optional direccion As String = "",
                                           Optional ciudad As String = "",
                                           Optional notas As String = "",
                                           Optional fechaNacimiento As Nullable(Of DateTime) = Nothing) As Boolean
            Try
                If String.IsNullOrWhiteSpace(nombre) Then
                    errorMessage = "El nombre es obligatorio."
                    Return False
                End If
                If Not EsRolValido(rol) Then
                    errorMessage = "El rol seleccionado no es válido."
                    Return False
                End If

                Dim dniTrimmed = If(dni, "").Trim()
                If String.IsNullOrWhiteSpace(dniTrimmed) Then
                    errorMessage = "El DNI del empleado es obligatorio."
                    Return False
                End If

                ' Verificar que el DNI no pertenezca a otro usuario
                If ExisteDni(dniTrimmed, usuarioId) Then
                    errorMessage = $"Ya existe otro usuario registrado con el DNI '{dniTrimmed}'."
                    Return False
                End If

                ' Si se intenta cambiar de rol al único admin, verificar que quede al menos uno activo
                If rol <> "Administrador" Then
                    Dim sqlAdmins As String = "SELECT COUNT(*) FROM `usuarios` WHERE `rol` = 'Administrador' AND `activo` = 1 AND `id` <> @id;"
                    Dim otrosAdmins = Convert.ToInt32(DatabaseHelper.ExecuteScalar(sqlAdmins, New Dictionary(Of String, Object) From {{"@id", usuarioId}}))
                    If otrosAdmins = 0 Then
                        errorMessage = "No se puede quitar el rol de Administrador al único Administrador activo del sistema."
                        Return False
                    End If
                End If

                Dim nombreCompleto As String = $"{apellido}, {nombre}".Trim(" "c, ","c)
                Dim sqlUpdate As String =
                    "UPDATE `usuarios` SET `dni` = @dni, `nombre` = @nom, `apellido` = @ape, `nombre_completo` = @nomcomp, `rol` = @rol, " &
                    "`telefono` = @tel, `email` = @email, `direccion` = @dir, `ciudad` = @ciu, `notas` = @notas, " &
                    "`fecha_nacimiento` = @fnac WHERE `id` = @id;"
                Dim prms As New Dictionary(Of String, Object) From {
                    {"@dni",     dniTrimmed},
                    {"@nom",     nombre.Trim()},
                    {"@ape",     apellido.Trim()},
                    {"@nomcomp", nombreCompleto},
                    {"@rol",     rol.Trim()},
                    {"@tel",     telefono.Trim()},
                    {"@email",   email.Trim()},
                    {"@dir",     direccion.Trim()},
                    {"@ciu",     ciudad.Trim()},
                    {"@notas",   notas.Trim()},
                    {"@fnac",    If(fechaNacimiento.HasValue, CObj(fechaNacimiento.Value.ToString("yyyy-MM-dd")), DBNull.Value)},
                    {"@id",      usuarioId}
                }

                DatabaseHelper.ExecuteNonQuery(sqlUpdate, prms)
                errorMessage = String.Empty
                Return True
            Catch ex As Exception
                errorMessage = "Error al actualizar usuario: " & ex.Message
                Return False
            End Try
        End Function

        Public Function CambiarPassword(usuarioId As Integer, nuevaPassword As String, ByRef errorMessage As String) As Boolean
            Try
                If String.IsNullOrWhiteSpace(nuevaPassword) OrElse nuevaPassword.Length < 4 Then
                    errorMessage = "La nueva contraseña debe tener al menos 4 caracteres."
                    Return False
                End If

                Dim hash = DatabaseHelper.HashPasswordSecure(nuevaPassword)
                Dim sqlUpdate As String = "UPDATE `usuarios` SET `password_hash` = @p WHERE `id` = @id;"
                DatabaseHelper.ExecuteNonQuery(sqlUpdate, New Dictionary(Of String, Object) From {
                    {"@p", hash},
                    {"@id", usuarioId}
                })

                errorMessage = String.Empty
                Return True
            Catch ex As Exception
                errorMessage = "Error al cambiar la contraseña: " & ex.Message
                Return False
            End Try
        End Function

        Public Function ToggleActivo(usuarioId As Integer, ByRef errorMessage As String) As Boolean
            Try
                Dim user = GetUsuarioById(usuarioId)
                If user Is Nothing Then
                    errorMessage = "Usuario no encontrado."
                    Return False
                End If

                Dim nuevoEstado As Boolean = Not user.Activo

                ' Si se intenta desactivar a un administrador, asegurar que quede al menos otro activo
                If Not nuevoEstado AndAlso user.Rol.Equals("Administrador", StringComparison.OrdinalIgnoreCase) Then
                    Dim sqlAdmins As String = "SELECT COUNT(*) FROM `usuarios` WHERE `rol` = 'Administrador' AND `activo` = 1 AND `id` <> @id;"
                    Dim otrosAdmins = Convert.ToInt32(DatabaseHelper.ExecuteScalar(sqlAdmins, New Dictionary(Of String, Object) From {{"@id", usuarioId}}))
                    If otrosAdmins = 0 Then
                        errorMessage = "No es posible desactivar al único Administrador activo del sistema."
                        Return False
                    End If
                End If

                Dim sqlUpdate As String = "UPDATE `usuarios` SET `activo` = @act WHERE `id` = @id;"
                DatabaseHelper.ExecuteNonQuery(sqlUpdate, New Dictionary(Of String, Object) From {
                    {"@act", If(nuevoEstado, 1, 0)},
                    {"@id", usuarioId}
                })

                errorMessage = String.Empty
                Return True
            Catch ex As Exception
                errorMessage = "Error al modificar estado: " & ex.Message
                Return False
            End Try
        End Function

        Private Function EsRolValido(rol As String) As Boolean
            Return rol.Equals("Administrador", StringComparison.OrdinalIgnoreCase) OrElse
                   rol.Equals("Gerente", StringComparison.OrdinalIgnoreCase) OrElse
                   rol.Equals("Vendedor", StringComparison.OrdinalIgnoreCase) OrElse
                   rol.Equals("Cajero", StringComparison.OrdinalIgnoreCase)
        End Function

        Private Function MapUsuario(row As DataRow) As Usuario
            Dim fnac As Nullable(Of DateTime) = Nothing
            If row.Table.Columns.Contains("fecha_nacimiento") AndAlso Not IsDBNull(row("fecha_nacimiento")) Then
                Dim parsed As DateTime
                If DateTime.TryParse(row("fecha_nacimiento").ToString(), parsed) Then
                    fnac = parsed
                End If
            End If

            Dim dniVal As String = ""
            If row.Table.Columns.Contains("dni") AndAlso Not IsDBNull(row("dni")) Then
                dniVal = row("dni").ToString()
            End If

            Dim nomVal As String = ""
            If row.Table.Columns.Contains("nombre") AndAlso Not IsDBNull(row("nombre")) Then
                nomVal = row("nombre").ToString()
            ElseIf row.Table.Columns.Contains("nombre_completo") AndAlso Not IsDBNull(row("nombre_completo")) Then
                nomVal = row("nombre_completo").ToString()
            End If

            Dim apeVal As String = ""
            If row.Table.Columns.Contains("apellido") AndAlso Not IsDBNull(row("apellido")) Then
                apeVal = row("apellido").ToString()
            End If

            Dim telVal As String = ""
            If row.Table.Columns.Contains("telefono") AndAlso Not IsDBNull(row("telefono")) Then
                telVal = row("telefono").ToString()
            End If

            Dim emailVal As String = ""
            If row.Table.Columns.Contains("email") AndAlso Not IsDBNull(row("email")) Then
                emailVal = row("email").ToString()
            End If

            Dim dirVal As String = ""
            If row.Table.Columns.Contains("direccion") AndAlso Not IsDBNull(row("direccion")) Then
                dirVal = row("direccion").ToString()
            End If

            Dim ciuVal As String = ""
            If row.Table.Columns.Contains("ciudad") AndAlso Not IsDBNull(row("ciudad")) Then
                ciuVal = row("ciudad").ToString()
            End If

            Dim notasVal As String = ""
            If row.Table.Columns.Contains("notas") AndAlso Not IsDBNull(row("notas")) Then
                notasVal = row("notas").ToString()
            End If

            Dim ultLogin As Nullable(Of DateTime) = Nothing
            If row.Table.Columns.Contains("ultimo_login") AndAlso Not IsDBNull(row("ultimo_login")) Then
                Dim parsedLogin As DateTime
                If DateTime.TryParse(row("ultimo_login").ToString(), parsedLogin) Then
                    ultLogin = parsedLogin
                End If
            End If

            Dim createdAtVal As DateTime = DateTime.MinValue
            If row.Table.Columns.Contains("created_at") AndAlso Not IsDBNull(row("created_at")) Then
                Dim parsedCreated As DateTime
                If DateTime.TryParse(row("created_at").ToString(), parsedCreated) Then
                    createdAtVal = parsedCreated
                End If
            End If

            Return New Usuario() With {
                .Id = Convert.ToInt32(row("id")),
                .Username = row("username").ToString(),
                .Dni = dniVal,
                .Nombre = nomVal,
                .Apellido = apeVal,
                .Rol = If(row.Table.Columns.Contains("rol"), row("rol").ToString(), "Vendedor"),
                .Telefono = telVal,
                .Email = emailVal,
                .Direccion = dirVal,
                .Ciudad = ciuVal,
                .Notas = notasVal,
                .FechaNacimiento = fnac,
                .Activo = If(row.Table.Columns.Contains("activo"), Convert.ToBoolean(row("activo")), True),
                .UltimoLogin = ultLogin,
                .CreatedAt = createdAtVal
            }
        End Function

    End Class
End Namespace
