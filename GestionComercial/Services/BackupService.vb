Imports System.IO
Imports Microsoft.Data.Sqlite
Imports GestionComercial.Config
Imports GestionComercial.Data

' ARCHIVO: BackupService.vb
' PROPÓSITO: Copias de seguridad atómicas (Online Backup API de SQLite) y restauración
' - Consistencia Transaccional Total (Online Backup API):
'   Utiliza la API oficial de respaldo en vivo de SQLite (SqliteConnection.BackupDatabase).
'   Permite realizar una copia fiel byte a byte sin bloquear lectores ni corromper transacciones en curso.
' - Seguridad de Integridad:
'   Antes de restaurar una copia, libera los pools de conexiones activas (ClearAllPools).
' - Directorio de Respaldo Centralizado:
'   Organiza los respaldos en 'Database/Backups/' con estampas cronológicas precisas (yyyyMMdd_HHmmss).

Namespace Services
    Public Class BackupItem
        Public Property NombreArchivo As String = String.Empty
        Public Property RutaCompleta As String = String.Empty
        Public Property FechaCreacion As DateTime = DateTime.MinValue
        Public Property TamanoBytes As Long = 0

        Public ReadOnly Property TamanoFormateado As String
            Get
                If TamanoBytes >= 1024 * 1024 Then
                    Return $"{TamanoBytes / (1024.0 * 1024.0):F2} MB"
                ElseIf TamanoBytes >= 1024 Then
                    Return $"{TamanoBytes / 1024.0:F1} KB"
                Else
                    Return $"{TamanoBytes} B"
                End If
            End Get
        End Property
    End Class

    Public Class DatabaseInfo
        Public Property RutaDb As String = String.Empty
        Public Property TamanoBytes As Long = 0
        Public Property UltimaModificacion As DateTime = DateTime.MinValue
        Public Property CantidadTablas As Integer = 0
        Public Property CantidadVentas As Integer = 0

        Public ReadOnly Property TamanoFormateado As String
            Get
                If TamanoBytes >= 1024 * 1024 Then
                    Return $"{TamanoBytes / (1024.0 * 1024.0):F2} MB"
                ElseIf TamanoBytes >= 1024 Then
                    Return $"{TamanoBytes / 1024.0:F1} KB"
                Else
                    Return $"{TamanoBytes} B"
                End If
            End Get
        End Property
    End Class

    Public Class BackupService

        Public Function GetDirectorioBackups() As String
            Dim dbPath = AppConfig.Settings.GetSqlitePath()
            Dim baseDir = Path.GetDirectoryName(dbPath)
            If String.IsNullOrEmpty(baseDir) Then
                baseDir = AppDomain.CurrentDomain.BaseDirectory
            End If
            Return Path.Combine(baseDir, "Backups")
        End Function

        Public Function GetBackups(Optional customDir As String = "") As List(Of BackupItem)
            Dim list As New List(Of BackupItem)()
            Dim dir = If(String.IsNullOrWhiteSpace(customDir), GetDirectorioBackups(), customDir)

            If Not Directory.Exists(dir) Then
                Return list
            End If

            Dim di As New DirectoryInfo(dir)
            Dim files = di.GetFiles("*.db").OrderByDescending(Function(f) f.CreationTime).ToList()

            For Each f In files
                list.Add(New BackupItem() With {
                    .NombreArchivo = f.Name,
                    .RutaCompleta = f.FullName,
                    .FechaCreacion = f.CreationTime,
                    .TamanoBytes = f.Length
                })
            Next

            Return list
        End Function

        Public Function CrearBackup(ByRef mensajeError As String) As String
            Return CrearBackup(String.Empty, mensajeError)
        End Function

        Public Function CrearBackup(rutaDestinoPersonalizada As String, ByRef mensajeError As String) As String
            mensajeError = String.Empty
            Try
                Dim origenPath = AppConfig.Settings.GetSqlitePath()
                If Not File.Exists(origenPath) Then
                    mensajeError = "No se encontró el archivo de base de datos origen en: " & origenPath
                    Return String.Empty
                End If

                Dim rutaFinal = rutaDestinoPersonalizada
                If String.IsNullOrWhiteSpace(rutaFinal) Then
                    Dim dirDefault = GetDirectorioBackups()
                    If Not Directory.Exists(dirDefault) Then
                        Directory.CreateDirectory(dirDefault)
                    End If
                    Dim timeStamp = DateTime.Now.ToString("yyyyMMdd_HHmmss")
                    rutaFinal = Path.Combine(dirDefault, $"backup_gestion_comercial_{timeStamp}.db")
                Else
                    Dim dirDestino = Path.GetDirectoryName(rutaFinal)
                    If Not String.IsNullOrEmpty(dirDestino) AndAlso Not Directory.Exists(dirDestino) Then
                        Directory.CreateDirectory(dirDestino)
                    End If
                End If

                ' Usamos el SQLite Online Backup API oficial de Microsoft.Data.Sqlite
                Using sourceConn As New SqliteConnection(AppConfig.ConnectionString)
                    sourceConn.Open()
                    Dim destConnStr = $"Data Source={rutaFinal}"
                    Using destConn As New SqliteConnection(destConnStr)
                        destConn.Open()
                        sourceConn.BackupDatabase(destConn)
                    End Using
                End Using

                Return rutaFinal
            Catch ex As Exception
                mensajeError = ex.Message
                Return String.Empty
            End Try
        End Function

        Public Function RestaurarBackup(rutaBackup As String, ByRef mensajeError As String) As Boolean
            mensajeError = String.Empty
            Try
                If Not File.Exists(rutaBackup) Then
                    mensajeError = "El archivo de respaldo seleccionado no existe."
                    Return False
                End If

                Dim origenPath = AppConfig.Settings.GetSqlitePath()

                ' Limpiamos todos los pools de conexiones activas en la aplicación
                SqliteConnection.ClearAllPools()

                ' Realizamos la restauración usando el SQLite Backup API (backup -> base activa)
                Dim backupConnStr = $"Data Source={rutaBackup}"
                Using sourceConn As New SqliteConnection(backupConnStr)
                    sourceConn.Open()
                    Using destConn As New SqliteConnection(AppConfig.ConnectionString)
                        destConn.Open()
                        sourceConn.BackupDatabase(destConn)
                    End Using
                End Using

                Return True
            Catch ex As Exception
                mensajeError = ex.Message
                Return False
            End Try
        End Function

        Public Function ObtenerInfoBaseDatosActual() As DatabaseInfo
            Dim info As New DatabaseInfo()
            Try
                Dim dbPath = AppConfig.Settings.GetSqlitePath()
                info.RutaDb = dbPath
                If File.Exists(dbPath) Then
                    Dim fi As New FileInfo(dbPath)
                    info.TamanoBytes = fi.Length
                    info.UltimaModificacion = fi.LastWriteTime
                End If

                Dim queryTablas = "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%';"
                info.CantidadTablas = Convert.ToInt32(DatabaseHelper.ExecuteScalar(queryTablas))

                Dim queryVentas = "SELECT COUNT(*) FROM ventas;"
                info.CantidadVentas = Convert.ToInt32(DatabaseHelper.ExecuteScalar(queryVentas))
            Catch ex As Exception
                ' Valores por defecto
            End Try
            Return info
        End Function

    End Class
End Namespace
