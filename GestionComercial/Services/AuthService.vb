' ARCHIVO: AuthService.vb
' PROPÓSITO: Autenticación, control de sesiones, protección anti fuerza bruta y permisos por rol.
' En este servicio resolvemos la seguridad de acceso al sistema:
' 1. Matriz de Roles (RBAC):
'    Diferenciamos 3 perfiles: Vendedor, Gerente y Administrador.
'    Para aplicar el Principio de Menor Privilegio (Least Privilege), el Administrador
'    no tiene acceso a cobrar en el POS ni a la Caja física para no mezclar funciones de auditoría con la manipulación material del dinero.
' 2. Seguridad en Contraseñas:
'    Verifica credenciales mediante PBKDF2 con Salt usando comparación en tiempo constante (FixedTimeEquals)
'    para evitar vulnerabilidades de análisis de tiempo de respuesta (Timing Attacks).
' 3. Protección contra Ataques de Fuerza Bruta:
'    Si una cuenta acumula 5 intentos fallidos consecutivos, queda bloqueada temporalmente por 60 segundos.
' 4. Autorizaciones en Punto de Venta:
'    Permite que un Gerente o Administrador autorice descuentos mayores al 15% o egresos de caja
'    ingresando sus credenciales en un diálogo modal sin tener que cerrar la sesión del vendedor,
'    registrando qué supervisor autorizó la operación.

Imports System.Data
Imports System.Windows.Forms
Imports GestionComercial.Data
Imports GestionComercial.Forms
Imports GestionComercial.Models

Namespace Services
    Public Module AuthService
        Private _currentUser As Usuario = Nothing
        Private _lastAuthorizingSupervisor As Usuario = Nothing

        ' Estructura para registrar los intentos fallidos de cada cuenta
        Private Structure LoginAttemptInfo
            Public FailedCount As Integer
            Public LockoutUntil As DateTime
        End Structure

        Private ReadOnly _attempts As New Dictionary(Of String, LoginAttemptInfo)(StringComparer.OrdinalIgnoreCase)
        Private ReadOnly _lockObj As New Object()
        Private Const SupervisorAttemptKey As String = "__SUPERVISOR_AUTH__"

        ' Usuario actualmente autenticado en la sesión
        Public Property CurrentUser As Usuario
            Get
                Return _currentUser
            End Get
            Private Set(value As Usuario)
                _currentUser = value
            End Set
        End Property

        ' Último supervisor que autorizó una operación excepcional en el turno actual
        Public Property LastAuthorizingSupervisor As Usuario
            Get
                Return _lastAuthorizingSupervisor
            End Get
            Private Set(value As Usuario)
                _lastAuthorizingSupervisor = value
            End Set
        End Property

        ' Indica si hay algún usuario con sesión iniciada en este momento
        Public ReadOnly Property IsAuthenticated As Boolean
            Get
                Return _currentUser IsNot Nothing
            End Get
        End Property

        ' Consultas de Roles para Segregación de Funciones:
        ' Estas propiedades las usamos en los formularios para mostrar u ocultar botones
        ' según los privilegios del perfil activo.

        ' Devuelve True únicamente si el usuario conectado tiene el rol "Administrador"
        Public ReadOnly Property IsAdmin As Boolean
            Get
                Return _currentUser IsNot Nothing AndAlso _currentUser.Rol.Equals("Administrador", StringComparison.OrdinalIgnoreCase)
            End Get
        End Property

        ' Devuelve True únicamente si el usuario conectado tiene el rol "Gerente"
        Public ReadOnly Property IsManager As Boolean
            Get
                Return _currentUser IsNot Nothing AndAlso _currentUser.Rol.Equals("Gerente", StringComparison.OrdinalIgnoreCase)
            End Get
        End Property

        ' Devuelve True únicamente si el usuario conectado tiene el rol "Vendedor"
        Public ReadOnly Property IsVendor As Boolean
            Get
                Return _currentUser IsNot Nothing AndAlso _currentUser.Rol.Equals("Vendedor", StringComparison.OrdinalIgnoreCase)
            End Get
        End Property

        ' Devuelve True si el usuario conectado es "Administrador" o "Gerente" (perfil de supervisión)
        Public ReadOnly Property IsAdminOrManager As Boolean
            Get
                Return IsAdmin OrElse IsManager
            End Get
        End Property

        ' MÉTODO PRINCIPAL DE INICIO DE SESIÓN
        ' Flujo de ejecución:
        ' 1. Revisa si el usuario está bloqueado por demasiados intentos fallidos.
        ' 2. Busca al usuario en la BD de forma segura con parámetros.
        ' 3. Valida la contraseña con PBKDF2 en tiempo constante.
        ' 4. Guarda la sesión en _currentUser y actualiza la fecha de último login.
        Public Function Login(username As String, password As String, ByRef errorMessage As String) As Boolean
            Try
                If String.IsNullOrWhiteSpace(username) OrElse String.IsNullOrWhiteSpace(password) Then
                    errorMessage = "Por favor ingrese usuario y contraseña."
                    Return False
                End If

                Dim userKey As String = username.Trim().ToLowerInvariant()

                ' 1. Verificación de bloqueo por intentos fallidos (Anti Fuerza Bruta)
                If EstaBloqueadoPorFuerzaBruta(userKey, errorMessage) Then
                    Return False
                End If

                Dim query As String = "SELECT * FROM `usuarios` WHERE `username` = @username AND `activo` = 1 LIMIT 1;"
                Dim params As New Dictionary(Of String, Object) From {{"@username", username.Trim()}}
                Dim dt As DataTable = DatabaseHelper.ExecuteQuery(query, params)

                If dt.Rows.Count = 0 Then
                    RegistrarFalloLogin(userKey, errorMessage)
                    Return False
                End If

                Dim row As DataRow = dt.Rows(0)
                Dim userId As Integer = Convert.ToInt32(row("id"))
                Dim storedHash As String = row("password_hash").ToString()

                ' Validación de credenciales con PBKDF2
                Dim passwordValid As Boolean = DatabaseHelper.VerifyPassword(password, storedHash)

                If Not passwordValid Then
                    RegistrarFalloLogin(userKey, errorMessage)
                    Return False
                End If

                ' Login exitoso: reiniciar contador de intentos fallidos
                LimpiarIntentosFallidos(userKey)

                CurrentUser = ConstruirUsuarioDesdeFila(row)

                ' Actualizar fecha de último login
                Dim updateLoginQuery As String = "UPDATE `usuarios` SET `ultimo_login` = NOW() WHERE `id` = @id;"
                DatabaseHelper.ExecuteNonQuery(updateLoginQuery, New Dictionary(Of String, Object) From {{"@id", userId}})

                AppLogger.LogAudit("LOGIN_EXITOSO", $"Usuario={CurrentUser.Username} | Rol={CurrentUser.Rol}")
                errorMessage = String.Empty
                Return True
            Catch ex As Exception
                AppLogger.LogError("Error en Login", ex)
                errorMessage = "Error al autenticar: " & ex.Message
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Verifica si la cuenta o clave indicada se encuentra temporalmente bloqueada por haber superado los 5 intentos fallidos.
        ''' Si el tiempo de bloqueo ya venció, limpia el registro y permite volver a intentar.
        ''' </summary>
        Private Function EstaBloqueadoPorFuerzaBruta(userKey As String, ByRef errorMessage As String) As Boolean
            SyncLock _lockObj
                If Not _attempts.ContainsKey(userKey) Then
                    Return False
                End If

                Dim attempt = _attempts(userKey)
                If attempt.LockoutUntil > DateTime.Now Then
                    Dim remainingSec As Integer = Math.Max(1, CInt((attempt.LockoutUntil - DateTime.Now).TotalSeconds))
                    errorMessage = $"Demasiados intentos fallidos. Por seguridad, espere {remainingSec} segundos antes de volver a intentar."
                    Return True
                End If

                If attempt.LockoutUntil > DateTime.MinValue Then
                    _attempts.Remove(userKey)
                End If
                Return False
            End SyncLock
        End Function

        ''' <summary>
        ''' Borra el historial de intentos fallidos de un usuario luego de iniciar sesión o autorizar correctamente.
        ''' </summary>
        Private Sub LimpiarIntentosFallidos(userKey As String)
            SyncLock _lockObj
                If _attempts.ContainsKey(userKey) Then
                    _attempts.Remove(userKey)
                End If
            End SyncLock
        End Sub

        ' Protección Anti Fuerza Bruta (Brute-Force Lockout):
        ' Llevamos la cuenta de intentos fallidos en un diccionario en memoria protegido
        ' con SyncLock para seguridad multihilo (Thread-Safe). Si el usuario acumula
        ' 5 fallos, congelamos la cuenta por 60 segundos y le avisamos en pantalla.
        Private Sub RegistrarFalloLogin(userKey As String, ByRef errorMessage As String)
            SyncLock _lockObj
                Dim info As LoginAttemptInfo = If(_attempts.ContainsKey(userKey), _attempts(userKey), New LoginAttemptInfo())
                info.FailedCount += 1

                Const MaxIntentos As Integer = 5
                Const SegundosBloqueo As Integer = 60

                If info.FailedCount >= MaxIntentos Then
                    info.LockoutUntil = DateTime.Now.AddSeconds(SegundosBloqueo)
                    errorMessage = $"Demasiados intentos fallidos. Su cuenta ha sido bloqueada temporalmente por {SegundosBloqueo} segundos."
                    AppLogger.LogWarning($"Bloqueo temporal por fuerza bruta activado para clave '{userKey}'.")
                Else
                    Dim restantes = MaxIntentos - info.FailedCount
                    errorMessage = $"Usuario o contraseña incorrectos. (Le quedan {restantes} intento{If(restantes > 1, "s", "")})"
                End If

                _attempts(userKey) = info
            End SyncLock
        End Sub

        ''' <summary>
        ''' Convierte una fila (DataRow) de la tabla 'usuarios' en un objeto Usuario en memoria.
        ''' </summary>
        Private Function ConstruirUsuarioDesdeFila(row As DataRow) As Usuario
            Dim dniVal As String = ""
            If row.Table.Columns.Contains("dni") AndAlso Not Convert.IsDBNull(row("dni")) Then
                dniVal = row("dni").ToString()
            End If

            Return New Usuario() With {
                .Id = Convert.ToInt32(row("id")),
                .Username = row("username").ToString(),
                .Dni = dniVal,
                .Nombre = If(Convert.IsDBNull(row("nombre")), "", row("nombre").ToString()),
                .Apellido = If(Convert.IsDBNull(row("apellido")), "", row("apellido").ToString()),
                .Rol = row("rol").ToString(),
                .Activo = Convert.ToBoolean(row("activo")),
                .UltimoLogin = DateTime.Now
            }
        End Function

        ' Cierra la sesión activa liberando el objeto de usuario en memoria
        Public Sub Logout()
            If CurrentUser IsNot Nothing Then
                AppLogger.LogAudit("LOGOUT", $"Usuario={CurrentUser.Username}")
            End If
            CurrentUser = Nothing
            LastAuthorizingSupervisor = Nothing
        End Sub

        ' Helper interno que consulta contraseñas filtrando por rol específico
        Private Function ValidarCredencialesPorRol(password As String, rol As String) As Boolean
            Try
                If String.IsNullOrWhiteSpace(password) Then Return False

                Dim sql As String = "SELECT * FROM `usuarios` WHERE `rol` = @rol AND `activo` = 1;"
                Dim dt = DatabaseHelper.ExecuteQuery(sql, New Dictionary(Of String, Object) From {{"@rol", rol}})
                For Each row As DataRow In dt.Rows
                    If DatabaseHelper.VerifyPassword(password, row("password_hash").ToString()) Then
                        LastAuthorizingSupervisor = ConstruirUsuarioDesdeFila(row)
                        Return True
                    End If
                Next
                Return False
            Catch ex As Exception
                AppLogger.LogError("Error en ValidarCredencialesPorRol", ex)
                Return False
            End Try
        End Function

        ' Autentica a un supervisor dentro de los roles permitidos aplicando bloqueo anti fuerza bruta y auditoría
        Public Function AutenticarSupervisorPorRoles(password As String, rolesPermitidos As String(), motivo As String, ByRef supervisorEncontrado As Usuario, ByRef errorMessage As String) As Boolean
            supervisorEncontrado = Nothing
            If EstaBloqueadoPorFuerzaBruta(SupervisorAttemptKey, errorMessage) Then
                Return False
            End If

            For Each rol In rolesPermitidos
                If ValidarCredencialesPorRol(password, rol) Then
                    LimpiarIntentosFallidos(SupervisorAttemptKey)
                    supervisorEncontrado = LastAuthorizingSupervisor
                    Dim operador = If(CurrentUser IsNot Nothing, CurrentUser.Username, "Sistema")
                    AppLogger.LogAudit("AUTORIZACION_SUPERVISOR", $"Supervisor={supervisorEncontrado.Username} ({supervisorEncontrado.Rol}) | Operador={operador} | Motivo={motivo}")
                    errorMessage = String.Empty
                    Return True
                End If
            Next

            RegistrarFalloLogin(SupervisorAttemptKey, errorMessage)
            Return False
        End Function

        ''' <summary>
        ''' Solicita autorización de un Gerente si el usuario actual no tiene ese rol (por ejemplo, para anular una venta).
        ''' Si quien opera ya es Gerente, retorna True sin abrir el diálogo.
        ''' </summary>
        Public Function SolicitarAutorizacionManager(owner As Form, motivo As String) As Boolean
            If IsManager Then
                LastAuthorizingSupervisor = CurrentUser
                Return True
            End If

            Using dlg As New FrmAutorizacionSupervisor("Autorización de Gerente", "Gerente", motivo, {"Gerente"})
                Return dlg.ShowDialog(owner) = DialogResult.OK
            End Using
        End Function

        ''' <summary>
        ''' Solicita autorización de un Gerente mediante un diálogo modal (incluso si el usuario actual es Administrador),
        ''' garantizando el principio de control dual (cuatro ojos) para operaciones críticas de respaldo y restauración.
        ''' </summary>
        Public Function SolicitarAutorizacionGerente(owner As Form, motivo As String) As Boolean
            Using dlg As New FrmAutorizacionSupervisor("Autorización de Gerente Requerida", "Gerente", motivo, {"Gerente"})
                Return dlg.ShowDialog(owner) = DialogResult.OK
            End Using
        End Function

        ''' <summary>
        ''' Solicita autorización de un Administrador o Gerente mediante diálogo modal si el usuario actual es Vendedor.
        ''' Si el usuario actual ya es Administrador o Gerente, retorna True directamente.
        ''' </summary>
        Public Function SolicitarAutorizacionAdminOManager(owner As Form, motivo As String) As Boolean
            If IsAdminOrManager Then
                LastAuthorizingSupervisor = CurrentUser
                Return True
            End If

            Using dlg As New FrmAutorizacionSupervisor("Autorización de Supervisor", "Administrador o Gerente", motivo, {"Administrador", "Gerente"})
                Return dlg.ShowDialog(owner) = DialogResult.OK
            End Using
        End Function

        ''' <summary>
        ''' Solicita autorización de un Administrador mediante un diálogo modal si el usuario actual no es Administrador,
        ''' o siempre que forzarConfirmacionClave sea True (para operaciones críticas como crear/restaurar copias de seguridad).
        ''' </summary>
        Public Function SolicitarAutorizacionAdmin(owner As Form, motivo As String, Optional forzarConfirmacionClave As Boolean = False) As Boolean
            If IsAdmin AndAlso Not forzarConfirmacionClave Then
                LastAuthorizingSupervisor = CurrentUser
                Return True
            End If

            Using dlg As New FrmAutorizacionSupervisor("Autorización de Administrador", "Administrador", motivo, {"Administrador"})
                Return dlg.ShowDialog(owner) = DialogResult.OK
            End Using
        End Function
    End Module
End Namespace
