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
                sql &= "ORDER BY `apellido` ASC, `nombre` ASC;"

                Dim dt = DatabaseHelper.ExecuteQuery(sql)
                For Each row As DataRow In dt.Rows
                    list.Add(MapUsuario(row))
                Next
            Catch ex As Exception
                ' Error silencioso en carga
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

        Public Function CrearUsuario(username As String, password As String,
                                      nombre As String, apellido As String,
                                      rol As String, ByRef errorMessage As String,
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

                ' Verificar que el username no exista ya
                Dim sqlCheck As String = "SELECT COUNT(*) FROM `usuarios` WHERE LOWER(`username`) = LOWER(@u);"
                Dim count = Convert.ToInt32(DatabaseHelper.ExecuteScalar(sqlCheck, New Dictionary(Of String, Object) From {{"@u", username.Trim()}}))
                If count > 0 Then
                    errorMessage = $"El nombre de usuario '{username}' ya está en uso. Elija otro."
                    Return False
                End If

                Dim hash = DatabaseHelper.HashPasswordSecure(password)
                Dim nombreCompleto As String = $"{apellido}, {nombre}".Trim(" "c, ","c)
                Dim sqlInsert As String =
                    "INSERT INTO `usuarios` (`username`, `password_hash`, `nombre`, `apellido`, `nombre_completo`, `rol`, " &
                    "`telefono`, `email`, `direccion`, `ciudad`, `notas`, `fecha_nacimiento`, `activo`, `created_at`) " &
                    "VALUES (@u, @p, @nom, @ape, @nomcomp, @rol, @tel, @email, @dir, @ciu, @notas, @fnac, 1, NOW());"
                Dim prms As New Dictionary(Of String, Object) From {
                    {"@u",       username.Trim().ToLowerInvariant()},
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
                    "UPDATE `usuarios` SET `nombre` = @nom, `apellido` = @ape, `nombre_completo` = @nomcomp, `rol` = @rol, " &
                    "`telefono` = @tel, `email` = @email, `direccion` = @dir, `ciudad` = @ciu, `notas` = @notas, " &
                    "`fecha_nacimiento` = @fnac WHERE `id` = @id;"
                Dim prms As New Dictionary(Of String, Object) From {
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

        Private Function MapUsuario(row As DataRow) As Usuario
            Dim fnac As Nullable(Of DateTime) = Nothing
            If Not IsDBNull(row("fecha_nacimiento")) Then
                Dim parsed As DateTime
                If DateTime.TryParse(row("fecha_nacimiento").ToString(), parsed) Then
                    fnac = parsed
                End If
            End If
            Return New Usuario() With {
                .Id = Convert.ToInt32(row("id")),
                .Username = row("username").ToString(),
                .Nombre = If(IsDBNull(row("nombre")), "", row("nombre").ToString()),
                .Apellido = If(IsDBNull(row("apellido")), "", row("apellido").ToString()),
                .Rol = row("rol").ToString(),
                .Telefono = If(IsDBNull(row("telefono")), "", row("telefono").ToString()),
                .Email = If(IsDBNull(row("email")), "", row("email").ToString()),
                .Direccion = If(IsDBNull(row("direccion")), "", row("direccion").ToString()),
                .Ciudad = If(IsDBNull(row("ciudad")), "", row("ciudad").ToString()),
                .Notas = If(IsDBNull(row("notas")), "", row("notas").ToString()),
                .FechaNacimiento = fnac,
                .Activo = Convert.ToBoolean(row("activo")),
                .UltimoLogin = If(IsDBNull(row("ultimo_login")), Nothing, Convert.ToDateTime(row("ultimo_login"))),
                .CreatedAt = If(IsDBNull(row("created_at")), DateTime.MinValue, Convert.ToDateTime(row("created_at")))
            }
        End Function

    End Class
End Namespace
