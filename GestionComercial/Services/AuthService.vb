Imports System.Data
Imports GestionComercial.Data
Imports GestionComercial.Models

Namespace Services
    Public Module AuthService
        Private _currentUser As Usuario = Nothing

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

                Dim query As String = "SELECT * FROM `usuarios` WHERE `username` = @username AND `activo` = 1 LIMIT 1;"
                Dim params As New Dictionary(Of String, Object) From {{"@username", username.Trim()}}
                Dim dt As DataTable = DatabaseHelper.ExecuteQuery(query, params)

                If dt.Rows.Count = 0 Then
                    errorMessage = "Usuario o contraseña incorrectos."
                    Return False
                End If

                Dim row As DataRow = dt.Rows(0)
                Dim storedHash As String = row("password_hash").ToString()
                Dim inputHash As String = DatabaseHelper.HashPassword(password)

                If Not String.Equals(storedHash, inputHash, StringComparison.OrdinalIgnoreCase) Then
                    errorMessage = "Usuario o contraseña incorrectos."
                    Return False
                End If

                CurrentUser = New Usuario() With {
                    .Id = Convert.ToInt32(row("id")),
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

        Public Sub Logout()
            CurrentUser = Nothing
        End Sub
    End Module
End Namespace
