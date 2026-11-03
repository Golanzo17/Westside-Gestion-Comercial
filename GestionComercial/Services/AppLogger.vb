' ARCHIVO: AppLogger.vb
' PROPÓSITO: Servicio centralizado de registro de eventos, errores técnicos y auditoría de seguridad.
' Escribe de forma segura para múltiples hilos (Thread-Safe) en la carpeta local "logs/" y en la salida de depuración.

Imports System.Diagnostics
Imports System.IO
Imports System.Text

Namespace Services
    Public Module AppLogger
        Private ReadOnly _syncRoot As New Object()

        Private Function GetLogFilePath() As String
            Dim baseDir = AppDomain.CurrentDomain.BaseDirectory
            Dim logDir = Path.Combine(baseDir, "logs")
            If Not Directory.Exists(logDir) Then
                Directory.CreateDirectory(logDir)
            End If
            Dim dateStamp = DateTime.Now.ToString("yyyyMMdd")
            Return Path.Combine(logDir, $"gestion_comercial_{dateStamp}.log")
        End Function

        Public Sub LogInfo(message As String)
            WriteEntry("INFO", message, Nothing)
        End Sub

        Public Sub LogAudit(action As String, details As String, Optional usuarioId As Nullable(Of Integer) = Nothing)
            Dim userTag = If(usuarioId.HasValue, $" [User:#{usuarioId.Value}]", String.Empty)
            WriteEntry("AUDIT", $"{action}{userTag} | {details}", Nothing)
        End Sub

        Public Sub LogWarning(message As String, Optional ex As Exception = Nothing)
            WriteEntry("WARN", message, ex)
        End Sub

        Public Sub LogError(context As String, ex As Exception)
            WriteEntry("ERROR", context, ex)
        End Sub

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
