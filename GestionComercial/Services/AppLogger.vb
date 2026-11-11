' ARCHIVO: AppLogger.vb
' PROPÓSITO: Servicio centralizado de registro de eventos, errores técnicos y auditoría de seguridad.
' Escribe de forma segura para múltiples hilos (Thread-Safe) en la carpeta local "logs/" y en la salida de depuración.

Imports System.Diagnostics
Imports System.IO
Imports System.Text

Namespace Services
    Public Module AppLogger
        ' Objeto de bloqueo para garantizar que varios hilos no escriban al mismo tiempo en el archivo .log
        Private ReadOnly _syncRoot As New Object()

        ''' <summary>
        ''' Construye y devuelve la ruta del archivo de log del día actual (ej: logs/gestion_comercial_20261110.log).
        ''' Si la carpeta "logs" no existe junto al ejecutable, la crea automáticamente.
        ''' </summary>
        Private Function GetLogFilePath() As String
            Dim baseDir = AppDomain.CurrentDomain.BaseDirectory
            Dim logDir = Path.Combine(baseDir, "logs")
            If Not Directory.Exists(logDir) Then
                Directory.CreateDirectory(logDir)
            End If
            Dim dateStamp = DateTime.Now.ToString("yyyyMMdd")
            Return Path.Combine(logDir, $"gestion_comercial_{dateStamp}.log")
        End Function

        ''' <summary>
        ''' Registra un mensaje informativo general sobre el funcionamiento del sistema (nivel INFO).
        ''' </summary>
        Public Sub LogInfo(message As String)
            WriteEntry("INFO", message, Nothing)
        End Sub

        ''' <summary>
        ''' Registra un evento de auditoría de seguridad o negocio (ej: login, apertura de caja, anulación, autorización).
        ''' Incluye la acción realizada, el detalle y opcionalmente el ID del usuario que la ejecutó (nivel AUDIT).
        ''' </summary>
        Public Sub LogAudit(action As String, details As String, Optional usuarioId As Nullable(Of Integer) = Nothing)
            Dim userTag = If(usuarioId.HasValue, $" [User:#{usuarioId.Value}]", String.Empty)
            WriteEntry("AUDIT", $"{action}{userTag} | {details}", Nothing)
        End Sub

        ''' <summary>
        ''' Registra una advertencia o situación anómala que no detiene el sistema (ej: bloqueo por fuerza bruta) (nivel WARN).
        ''' </summary>
        Public Sub LogWarning(message As String, Optional ex As Exception = Nothing)
            WriteEntry("WARN", message, ex)
        End Sub

        ''' <summary>
        ''' Registra un error técnico o excepción capturada en un bloque Try/Catch indicando en qué método ocurrió (nivel ERROR).
        ''' </summary>
        Public Sub LogError(context As String, ex As Exception)
            WriteEntry("ERROR", context, ex)
        End Sub

        ''' <summary>
        ''' Método interno que arma la línea de texto con fecha, hora, nivel y mensaje, y la guarda en el archivo de log de forma segura (SyncLock).
        ''' </summary>
        Private Sub WriteEntry(level As String, message As String, ex As Exception)
            Dim timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")
            Dim errorDetail = If(ex IsNot Nothing, $" | Ex: {ex.GetType().Name}: {ex.Message}", String.Empty)
            Dim line = $"[{timestamp}] [{level}] {message}{errorDetail}"

            Debug.WriteLine(line)

            Try
                SyncLock _syncRoot
                    Dim filePath = GetLogFilePath()
                    File.AppendAllText(filePath, line & Environment.NewLine, Encoding.UTF8)
                End SyncLock
            Catch
                ' No interrumpimos la aplicación si el disco es de solo lectura
            End Try
        End Sub
    End Module
End Namespace
