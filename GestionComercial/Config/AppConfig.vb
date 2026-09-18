Imports System.IO
Imports System.Text.Json

Namespace Config
    Public Class DatabaseSettings
        Public Property Provider As String = "SQLite" ' "SQLite" (recomendado, sin servidor) o "MySQL"
        Public Property SqliteFileName As String = "gestion_comercial.db"
        Public Property Host As String = "localhost"
        Public Property Port As Integer = 3306
        Public Property Database As String = "gestion_comercial_db"
        Public Property Username As String = "root"
        Public Property Password As String = ""

        Public Function GetSqlitePath() As String
            ' Buscar el archivo gestion_comercial.db priorizando la carpeta raíz del proyecto
            Dim baseDir As String = AppDomain.CurrentDomain.BaseDirectory
            Dim candidates As String() = {
                Path.Combine(baseDir, "..", "..", "..", "..", "Database", SqliteFileName),
                Path.Combine(baseDir, "..", "..", "..", "Database", SqliteFileName),
                Path.Combine(baseDir, "Database", SqliteFileName),
                Path.Combine(baseDir, SqliteFileName)
            }

            For Each c In candidates
                Try
                    Dim full = Path.GetFullPath(c)
                    If File.Exists(full) Then
                        Return full
                    End If
                Catch
                End Try
            Next

            ' Si no existe el archivo aún, devolver la ruta en la carpeta Database del proyecto si la carpeta existe
            For Each c In candidates
                Try
                    Dim dir = Path.GetDirectoryName(Path.GetFullPath(c))
                    If Directory.Exists(dir) Then
                        Return Path.Combine(dir, SqliteFileName)
                    End If
                Catch
                End Try
            Next

            ' Fallback al directorio ejecutable
            Return Path.Combine(baseDir, SqliteFileName)
        End Function

        Public Function GetConnectionString() As String
            If Provider.Equals("SQLite", StringComparison.OrdinalIgnoreCase) Then
                Return $"Data Source={GetSqlitePath()};"
            Else
                Dim builder As New MySqlConnector.MySqlConnectionStringBuilder()
                builder.Server = Host
                builder.Port = CUInt(Port)
                builder.Database = Database
                builder.UserID = Username
                builder.Password = Password
                builder.CharacterSet = "utf8mb4"
                builder.ConnectionTimeout = 5
                builder.DefaultCommandTimeout = 30
                builder.AllowUserVariables = True
                Return builder.ConnectionString
            End If
        End Function

        Public Function GetServerOnlyConnectionString() As String
            If Provider.Equals("SQLite", StringComparison.OrdinalIgnoreCase) Then
                Return GetConnectionString()
            Else
                Dim builder As New MySqlConnector.MySqlConnectionStringBuilder()
                builder.Server = Host
                builder.Port = CUInt(Port)
                builder.UserID = Username
                builder.Password = Password
                builder.CharacterSet = "utf8mb4"
                builder.ConnectionTimeout = 5
                Return builder.ConnectionString
            End If
        End Function
    End Class

    Public Module AppConfig
        Private ReadOnly ConfigFilePath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json")
        Private _settings As DatabaseSettings

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

        Public Sub LoadSettings()
            Try
                If File.Exists(ConfigFilePath) Then
                    Dim json As String = File.ReadAllText(ConfigFilePath)
                    _settings = JsonSerializer.Deserialize(Of DatabaseSettings)(json)
                End If
            Catch ex As Exception
                ' Fallback
            End Try

            If _settings Is Nothing Then
                _settings = New DatabaseSettings()
                SaveSettings()
            Else
                If String.IsNullOrWhiteSpace(_settings.Provider) Then
                    _settings.Provider = "SQLite"
                    SaveSettings()
                End If
            End If
        End Sub

        Public Sub SaveSettings()
            Try
                Dim options As New JsonSerializerOptions With {.WriteIndented = True}
                Dim json As String = JsonSerializer.Serialize(_settings, options)
                File.WriteAllText(ConfigFilePath, json)
            Catch ex As Exception
                ' Silently continue
            End Try
        End Sub

        Public ReadOnly Property ConnectionString As String
            Get
                Return Settings.GetConnectionString()
            End Get
        End Property
    End Module
End Namespace
