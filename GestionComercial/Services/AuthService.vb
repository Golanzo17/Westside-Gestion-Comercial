Imports System.Data
Imports System.Drawing
Imports System.Windows.Forms
Imports GestionComercial.Data
Imports GestionComercial.Models
Imports GestionComercial.UI

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

        Public ReadOnly Property IsManager As Boolean
            Get
                Return _currentUser IsNot Nothing AndAlso _currentUser.Rol.Equals("Gerente", StringComparison.OrdinalIgnoreCase)
            End Get
        End Property

        Public ReadOnly Property IsVendor As Boolean
            Get
                Return _currentUser IsNot Nothing AndAlso _currentUser.Rol.Equals("Vendedor", StringComparison.OrdinalIgnoreCase)
            End Get
        End Property

        Public ReadOnly Property IsAdminOrManager As Boolean
            Get
                Return IsAdmin OrElse IsManager
            End Get
        End Property

        Public Function SolicitarAutorizacionAdminOManager(owner As Form, motivo As String) As Boolean
            If IsAdminOrManager Then Return True
            Return SolicitarAutorizacionAdmin(owner, motivo)
        End Function

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

        ''' <summary>
        ''' Valida una contraseña contra los administradores activos del sistema.
        ''' Si se especifica un username, valida contra ese administrador específico.
        ''' </summary>
        Public Function ValidarCredencialesAdmin(password As String, Optional username As String = "") As Boolean
            Try
                If String.IsNullOrWhiteSpace(password) Then Return False

                Dim sql As String
                Dim prms As New Dictionary(Of String, Object)()
                If Not String.IsNullOrWhiteSpace(username) Then
                    sql = "SELECT id, password_hash FROM `usuarios` WHERE `username` = @user AND `rol` = 'Administrador' AND `activo` = 1 LIMIT 1;"
                    prms.Add("@user", username.Trim())
                Else
                    sql = "SELECT id, password_hash FROM `usuarios` WHERE `rol` = 'Administrador' AND `activo` = 1;"
                End If

                Dim dt = DatabaseHelper.ExecuteQuery(sql, prms)
                For Each row As DataRow In dt.Rows
                    Dim storedHash = row("password_hash").ToString()
                    Dim dummyRehash As Boolean = False
                    If DatabaseHelper.VerifyPassword(password, storedHash, dummyRehash) Then
                        Return True
                    End If
                Next
                Return False
            Catch ex As Exception
                Return False
            End Try
        End Function

        Public Function ValidarCredencialesManager(password As String) As Boolean
            Return ValidarCredencialesPorRol(password, "Gerente")
        End Function

        Private Function ValidarCredencialesPorRol(password As String, rol As String) As Boolean
            Try
                If String.IsNullOrWhiteSpace(password) Then Return False

                Dim sql As String = "SELECT password_hash FROM `usuarios` WHERE `rol` = @rol AND `activo` = 1;"
                Dim dt = DatabaseHelper.ExecuteQuery(sql, New Dictionary(Of String, Object) From {{"@rol", rol}})
                For Each row As DataRow In dt.Rows
                    Dim dummyRehash As Boolean = False
                    If DatabaseHelper.VerifyPassword(password, row("password_hash").ToString(), dummyRehash) Then Return True
                Next
                Return False
            Catch
                Return False
            End Try
        End Function

        Public Function SolicitarAutorizacionManager(owner As Form, motivo As String) As Boolean
            If IsManager Then Return True

            Using dlg As New Form()
                dlg.Text = "Autorización de Gerente"
                dlg.Size = New Size(430, 270)
                dlg.StartPosition = FormStartPosition.CenterParent
                dlg.FormBorderStyle = FormBorderStyle.FixedDialog
                dlg.MaximizeBox = False
                dlg.MinimizeBox = False
                dlg.BackColor = Color.White

                Dim lblMotivo As New Label() With {
                    .Text = motivo,
                    .Font = UITheme.FontRegular,
                    .ForeColor = UITheme.ColorDanger,
                    .Location = New Point(20, 25),
                    .Size = New Size(375, 55)
                }
                Dim lblPass As New Label() With {
                    .Text = "Contraseña de Gerente:",
                    .Font = UITheme.FontBold,
                    .Location = New Point(20, 95),
                    .AutoSize = True
                }
                Dim txtPass As New TextBox() With {
                    .Location = New Point(20, 120),
                    .Size = New Size(375, 26),
                    .UseSystemPasswordChar = True
                }
                UITheme.StyleTextBox(txtPass)
                Dim btnConfirmar As New Button() With {.Text = "Autorizar", .Location = New Point(185, 170), .Size = New Size(100, 34)}
                UITheme.StyleButton(btnConfirmar, "Primary")
                Dim btnCancelar As New Button() With {.Text = "Cancelar", .Location = New Point(295, 170), .Size = New Size(100, 34)}
                UITheme.StyleButton(btnCancelar, "Secondary")
                Dim autorizado As Boolean = False

                AddHandler btnConfirmar.Click, Sub()
                                                    If ValidarCredencialesManager(txtPass.Text) Then
                                                        autorizado = True
                                                        dlg.Close()
                                                    Else
                                                        MessageBox.Show("Contraseña de Gerente incorrecta.", "Acceso Denegado", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                        txtPass.Clear()
                                                        txtPass.Focus()
                                                    End If
                                                End Sub
                AddHandler btnCancelar.Click, Sub() dlg.Close()
                dlg.Controls.AddRange({lblMotivo, lblPass, txtPass, btnConfirmar, btnCancelar})
                dlg.AcceptButton = btnConfirmar
                dlg.CancelButton = btnCancelar
                dlg.ShowDialog(owner)
                Return autorizado
            End Using
        End Function

        ''' <summary>
        ''' Solicita autorización de un Administrador mediante un diálogo modal si el usuario actual no es Administrador.
        ''' Si el usuario actual ya es Administrador, retorna True directamente.
        ''' </summary>
        Public Function SolicitarAutorizacionAdmin(owner As Form, motivo As String) As Boolean
            If IsAdmin Then Return True

            Using dlg As New Form()
                dlg.Text = "Autorización Requerida"
                dlg.Size = New Size(430, 270)
                dlg.StartPosition = FormStartPosition.CenterParent
                dlg.FormBorderStyle = FormBorderStyle.FixedDialog
                dlg.MaximizeBox = False
                dlg.MinimizeBox = False
                dlg.BackColor = Color.White

                Dim pnlTop As New Panel() With {
                    .Dock = DockStyle.Top,
                    .Height = 55,
                    .BackColor = UITheme.ColorSecondary,
                    .Padding = New Padding(15, 12, 15, 10)
                }
                Dim lblTitle As New Label() With {
                    .Text = "🔒 AUTORIZACIÓN DE ADMINISTRADOR",
                    .Font = UITheme.FontSubheading,
                    .ForeColor = Color.White,
                    .AutoSize = True,
                    .Location = New Point(12, 15)
                }
                pnlTop.Controls.Add(lblTitle)

                Dim lblMotivo As New Label() With {
                    .Text = motivo,
                    .Font = UITheme.FontRegular,
                    .ForeColor = UITheme.ColorDanger,
                    .Location = New Point(20, 68),
                    .Size = New Size(375, 40)
                }

                Dim lblPass As New Label() With {
                    .Text = "Contraseña de Administrador:",
                    .Font = UITheme.FontBold,
                    .Location = New Point(20, 115),
                    .AutoSize = True
                }

                Dim txtPass As New TextBox() With {
                    .Location = New Point(20, 138),
                    .Size = New Size(375, 26),
                    .UseSystemPasswordChar = True
                }
                UITheme.StyleTextBox(txtPass)

                Dim btnConfirmar As New Button() With {
                    .Text = "Autorizar",
                    .Location = New Point(185, 180),
                    .Size = New Size(100, 34)
                }
                UITheme.StyleButton(btnConfirmar, "Primary")

                Dim btnCancelar As New Button() With {
                    .Text = "Cancelar",
                    .Location = New Point(295, 180),
                    .Size = New Size(100, 34)
                }
                UITheme.StyleButton(btnCancelar, "Secondary")

                Dim autorizado As Boolean = False

                AddHandler btnConfirmar.Click, Sub()
                    If String.IsNullOrWhiteSpace(txtPass.Text) Then
                        MessageBox.Show("Ingrese la contraseña de Administrador.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return
                    End If

                    If ValidarCredencialesAdmin(txtPass.Text) Then
                        autorizado = True
                        dlg.DialogResult = DialogResult.OK
                        dlg.Close()
                    Else
                        MessageBox.Show("Contraseña de Administrador incorrecta.", "Acceso Denegado", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        txtPass.Clear()
                        txtPass.Focus()
                    End If
                End Sub

                AddHandler btnCancelar.Click, Sub()
                    dlg.DialogResult = DialogResult.Cancel
                    dlg.Close()
                End Sub

                AddHandler txtPass.KeyDown, Sub(s, e)
                    If e.KeyCode = Keys.Enter Then
                        btnConfirmar.PerformClick()
                    End If
                End Sub

                dlg.Controls.AddRange({pnlTop, lblMotivo, lblPass, txtPass, btnConfirmar, btnCancelar})
                dlg.AcceptButton = btnConfirmar
                dlg.CancelButton = btnCancelar

                dlg.ShowDialog(owner)
                Return autorizado
            End Using
        End Function
    End Module
End Namespace
