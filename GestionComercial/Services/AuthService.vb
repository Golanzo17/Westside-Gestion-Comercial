Imports System.Data
Imports GestionComercial.Data
Imports GestionComercial.Models

Namespace Services
    Public Module AuthService
        Private _currentUser As Usuario = Nothing

        Private Structure LoginAttemptInfo
            Public FailedCount As Integer
            Public LockoutUntil As DateTime
        End Structure

        Private ReadOnly _attempts As New Dictionary(Of String, LoginAttemptInfo)(StringComparer.OrdinalIgnoreCase)
        Private ReadOnly _lockObj As New Object()

        Public Property CurrentUser As Usuario
            Get
                Return _currentUser
            End Get
            Private Set(value As Usuario)
                _currentUser = value
            End Set
        End Property

        Public ReadOnly Property IsAuthenticated As Boolean
            Get
                Return _currentUser IsNot Nothing
            End Get
        End Property

        Public ReadOnly Property IsAdmin As Boolean
            Get
                Return _currentUser IsNot Nothing AndAlso _currentUser.Rol.Equals("Administrador", StringComparison.OrdinalIgnoreCase)
            End Get
        End Property

        Public Function Login(username As String, password As String, ByRef errorMessage As String) As Boolean
            Try
                If String.IsNullOrWhiteSpace(username) OrElse String.IsNullOrWhiteSpace(password) Then
                    errorMessage = "Por favor ingrese usuario y contraseña."
                    Return False
                End If

                Dim userKey As String = username.Trim().ToLowerInvariant()

                ' 1. Verificación de bloqueo por intentos fallidos (Anti Fuerza Bruta)
                SyncLock _lockObj
                    If _attempts.ContainsKey(userKey) Then
                        Dim attempt = _attempts(userKey)
                        If attempt.LockoutUntil > DateTime.Now Then
                            Dim remainingSec As Integer = Math.Max(1, CInt((attempt.LockoutUntil - DateTime.Now).TotalSeconds))
                            errorMessage = $"Demasiados intentos fallidos. Por seguridad, espere {remainingSec} segundos antes de volver a intentar."
                            Return False
                        ElseIf attempt.LockoutUntil > DateTime.MinValue Then
                            ' El bloqueo ya expiró: resetear contador
                            _attempts.Remove(userKey)
                        End If
                    End If
                End SyncLock

                Dim query As String = "SELECT * FROM `usuarios` WHERE `username` = @username AND `activo` = 1 LIMIT 1;"
                Dim params As New Dictionary(Of String, Object) From {{"@username", username.Trim()}}
                Dim dt As DataTable = DatabaseHelper.ExecuteQuery(query, params)

                If dt.Rows.Count = 0 Then
                    RegistrarFalloLogin(userKey, errorMessage)
                    Return False
                End If

                Dim row As DataRow = dt.Rows(0)
                Dim storedHash As String = row("password_hash").ToString()
                Dim needsRehash As Boolean = False

                ' 2. Validación de contraseña con PBKDF2 / SHA-256 legado y protección contra timing attacks
                Dim passwordValid As Boolean = DatabaseHelper.VerifyPassword(password, storedHash, needsRehash)

                If Not passwordValid Then
                    RegistrarFalloLogin(userKey, errorMessage)
                    Return False
                End If

                ' 3. Auto-migración transparente: si la clave era SHA-256 legacy, actualizarla a PBKDF2 salado
                Dim userId As Integer = Convert.ToInt32(row("id"))
                If needsRehash Then
                    Try
                        Dim newSecureHash = DatabaseHelper.HashPasswordSecure(password)
                        Dim updateHashSql = "UPDATE `usuarios` SET `password_hash` = @newHash WHERE `id` = @id;"
                        DatabaseHelper.ExecuteNonQuery(updateHashSql, New Dictionary(Of String, Object) From {
                            {"@newHash", newSecureHash},
                            {"@id", userId}
                        })
                    Catch
                        ' Si falla la actualización en BD no se interrumpe la sesión del usuario
                    End Try
                End If

                ' 4. Login exitoso: limpiar historial de intentos fallidos
                SyncLock _lockObj
                    If _attempts.ContainsKey(userKey) Then
                        _attempts.Remove(userKey)
                    End If
                End SyncLock

                CurrentUser = New Usuario() With {
                    .Id = userId,
                    .Username = row("username").ToString(),
                    .NombreCompleto = row("nombre_completo").ToString(),
                    .Rol = row("rol").ToString(),
                    .Activo = Convert.ToBoolean(row("activo")),
                    .UltimoLogin = DateTime.Now
                }

                ' Actualizar fecha de último login
                Dim updateLoginQuery As String = "UPDATE `usuarios` SET `ultimo_login` = NOW() WHERE `id` = @id;"
                DatabaseHelper.ExecuteNonQuery(updateLoginQuery, New Dictionary(Of String, Object) From {{"@id", CurrentUser.Id}})

                errorMessage = String.Empty
                Return True
            Catch ex As Exception
                errorMessage = "Error al autenticar: " & ex.Message
                Return False
            End Try
        End Function

        Private Sub RegistrarFalloLogin(userKey As String, ByRef errorMessage As String)
            SyncLock _lockObj
                Dim info As LoginAttemptInfo = If(_attempts.ContainsKey(userKey), _attempts(userKey), New LoginAttemptInfo())
                info.FailedCount += 1

                Const MaxIntentos As Integer = 5
                Const SegundosBloqueo As Integer = 60

                If info.FailedCount >= MaxIntentos Then
                    info.LockoutUntil = DateTime.Now.AddSeconds(SegundosBloqueo)
                    errorMessage = $"Demasiados intentos fallidos. Su cuenta ha sido bloqueada temporalmente por {SegundosBloqueo} segundos."
                Else
                    Dim restantes = MaxIntentos - info.FailedCount
                    errorMessage = $"Usuario o contraseña incorrectos. (Le quedan {restantes} intento{If(restantes > 1, "s", "")})"
                End If

                _attempts(userKey) = info
            End SyncLock
        End Sub

        Public Sub Logout()
            CurrentUser = Nothing
        End Sub
    End Module
End Namespace
