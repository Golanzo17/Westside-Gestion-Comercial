' ARCHIVO: AppConfig.vb
' PROPÓSITO: Administrador centralizado de la configuración de base de datos del sistema.
' En este archivo manejamos la configuración de persistencia del sistema. Decidimos
' guardar los parámetros en formato JSON (appsettings.json) usando System.Text.Json.
' El sistema utiliza como motor principal SQLite embebido: un archivo .db local que
' ofrece total portabilidad, cero dependencias de servicios externos y cumplimiento ACID.

Imports System.IO
Imports System.Text.Json
Imports System.Text.Json.Serialization
Imports GestionComercial.Services

Namespace Config

    ' Configuración de persistencia para conexión a la base de datos SQLite.
    Public Class DatabaseSettings
        ' Motor seleccionado: "SQLite"
        Public Property Provider As String = "SQLite"
        Public Property SqliteFileName As String = "gestion_comercial.db"

        ' Permite redirigir la conexión a un archivo temporal aislado durante pruebas automatizadas.
        <JsonIgnore>
        Public Property CustomDatabasePath As String = String.Empty

        ' Ubica el archivo SQLite (.db) buscando en las carpetas relativas del proyecto o junto al ejecutable.
        Public Function GetSqlitePath() As String
            If Not String.IsNullOrWhiteSpace(CustomDatabasePath) Then
                Return Path.GetFullPath(CustomDatabasePath)
            End If

            Dim baseDir As String = AppDomain.CurrentDomain.BaseDirectory
            Dim candidates As String() = {
                Path.Combine(baseDir, "..", "..", "..", "..", "Database", SqliteFileName),
                Path.Combine(baseDir, "..", "..", "..", "Database", SqliteFileName),
                Path.Combine(baseDir, "Database", SqliteFileName),
                Path.Combine(baseDir, SqliteFileName)
            }

            ' Recorremos las rutas candidatas y si el archivo ya existe físicamente, usamos esa
            For Each c In candidates
                Try
                    Dim full = Path.GetFullPath(c)
                    If File.Exists(full) Then
                        Return full
                    End If
                Catch ex As Exception
                    AppLogger.LogWarning("Ruta candidata SQLite inválida: " & c, ex)
                End Try
            Next

            ' Si el archivo aún no existe (primera ejecución), buscamos si la carpeta Database existe
            For Each c In candidates
                Try
                    Dim dir = Path.GetDirectoryName(Path.GetFullPath(c))
                    If Not String.IsNullOrEmpty(dir) AndAlso Directory.Exists(dir) Then
                        Return Path.Combine(dir, SqliteFileName)
                    End If
                Catch ex As Exception
                    AppLogger.LogWarning("Directorio candidato SQLite inválido: " & c, ex)
                End Try
            Next

            ' Si no encontramos nada, devolvemos la ruta junto al ejecutable como respaldo (fallback)
            Return Path.Combine(baseDir, SqliteFileName)
        End Function

        ' Construye la cadena de conexión completa para SQLite con integridad referencial activa (Foreign Keys=True).
        Public Function GetConnectionString() As String
            Return $"Data Source={GetSqlitePath()};Foreign Keys=True;"
        End Function
    End Class

    ' Módulo global para lectura y guardado de appsettings.json.
    Public Module AppConfig
        Private ReadOnly ConfigFilePath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json")
        Private _settings As DatabaseSettings

        ' Propiedad estática para acceder a la configuración desde cualquier parte del sistema
        Public Property Settings As DatabaseSettings
            Get
                If _settings Is Nothing Then
                    LoadSettings()
                End If
                Return _settings
            End Get
            Set(value As DatabaseSettings)
                _settings = value
            End Set
        End Property

        ' Carga la configuración desde appsettings.json. Si no existe, genera la configuración por defecto.
        Public Sub LoadSettings()
            Try
                If File.Exists(ConfigFilePath) Then
                    Dim json As String = File.ReadAllText(ConfigFilePath)
                    _settings = JsonSerializer.Deserialize(Of DatabaseSettings)(json)
                End If
            Catch ex As Exception
                ' Si hay error de lectura, registramos la advertencia y usamos valores por defecto
                AppLogger.LogWarning("No se pudo leer appsettings.json, usando configuración por defecto.", ex)
            End Try

            If _settings Is Nothing Then
                _settings = New DatabaseSettings()
                SaveSettings()
                Return
            End If

            If String.IsNullOrWhiteSpace(_settings.Provider) Then
                _settings.Provider = "SQLite"
                SaveSettings()
            End If
        End Sub

        ' Guarda los ajustes actuales en appsettings.json con formato legible.
        Public Sub SaveSettings()
            Try
                Dim options As New JsonSerializerOptions With {.WriteIndented = True}
                Dim json As String = JsonSerializer.Serialize(_settings, options)
                File.WriteAllText(ConfigFilePath, json)
            Catch ex As Exception
                ' Registramos el aviso sin interrumpir el flujo si el archivo está bloqueado
                AppLogger.LogWarning("No se pudo guardar appsettings.json.", ex)
            End Try
        End Sub

        ' Acceso rápido a la cadena de conexión activa
        Public ReadOnly Property ConnectionString As String
            Get
                Return Settings.GetConnectionString()
            End Get
        End Property
    End Module
End Namespace
