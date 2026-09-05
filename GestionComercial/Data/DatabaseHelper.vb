Imports System.Data
Imports System.Data.Common
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports Microsoft.Data.Sqlite
Imports MySqlConnector
Imports GestionComercial.Config

Namespace Data
    Public Module DatabaseHelper

        Public ReadOnly Property IsSQLite As Boolean
            Get
                Return AppConfig.Settings.Provider.Equals("SQLite", StringComparison.OrdinalIgnoreCase)
            End Get
        End Property

        Public Function GetConnection() As DbConnection
            If IsSQLite Then
                ' Asegurar que el directorio de la BD SQLite exista
                Dim dbPath = AppConfig.Settings.GetSqlitePath()
                Dim dir = Path.GetDirectoryName(dbPath)
                If Not String.IsNullOrEmpty(dir) AndAlso Not Directory.Exists(dir) Then
                    Directory.CreateDirectory(dir)
                End If
                Return New SqliteConnection(AppConfig.ConnectionString)
            Else
                Return New MySqlConnection(AppConfig.ConnectionString)
            End If
        End Function

        Public Function PrepareQuery(query As String) As String
            If IsSQLite Then
                ' Adaptaciones automáticas de sintaxis para SQLite
                Dim q = query
                q = q.Replace("NOW()", "datetime('now', 'localtime')", StringComparison.OrdinalIgnoreCase)
                q = q.Replace("CURDATE()", "date('now', 'localtime')", StringComparison.OrdinalIgnoreCase)
                q = q.Replace("FOR UPDATE", "", StringComparison.OrdinalIgnoreCase)
                Return q
            End If
            Return query
        End Function

        Public Function TestConnection(ByRef errorMessage As String) As Boolean
            Try
                If IsSQLite Then
                    ' Para SQLite, verificar que exista y esté inicializada
                    Dim dbPath = AppConfig.Settings.GetSqlitePath()
                    If Not File.Exists(dbPath) Then
                        Dim initMsg As String = ""
                        InitializeDatabaseAndTables(initMsg)
                    End If
                End If

                Using conn As DbConnection = GetConnection()
                    conn.Open()

                    ' Si es SQLite, verificar que las tablas principales existan; si no, inicializar
                    If IsSQLite Then
                        Using cmdCheck = conn.CreateCommand()
                            cmdCheck.CommandText = "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='usuarios';"
                            Dim tblCount As Long = Convert.ToInt64(cmdCheck.ExecuteScalar())
                            If tblCount = 0 Then
                                Dim initMsg As String = ""
                                InitializeDatabaseAndTables(initMsg)
                            End If
                        End Using
                    End If

                    errorMessage = String.Empty
                    Return True
                End Using
            Catch ex As Exception
                errorMessage = ex.Message
                Return False
            End Try
        End Function

        Public Function CreateCommand(conn As DbConnection, query As String, Optional trans As DbTransaction = Nothing) As DbCommand
            Dim cmd As DbCommand = conn.CreateCommand()
            cmd.CommandText = PrepareQuery(query)
            If trans IsNot Nothing Then cmd.Transaction = trans
            Return cmd
        End Function

        Public Sub AddParam(cmd As DbCommand, name As String, value As Object)
            Dim p = cmd.CreateParameter()
            p.ParameterName = name
            p.Value = If(value, DBNull.Value)
            cmd.Parameters.Add(p)
        End Sub

        Public Function GetLastInsertedId(conn As DbConnection, trans As DbTransaction) As Long
            Using cmd As DbCommand = conn.CreateCommand()
                cmd.Transaction = trans
                If IsSQLite Then
                    cmd.CommandText = "SELECT last_insert_rowid();"
                Else
                    cmd.CommandText = "SELECT LAST_INSERT_ID();"
                End If
                Return Convert.ToInt64(cmd.ExecuteScalar())
            End Using
        End Function

        Public Function ExecuteQuery(query As String, Optional parameters As Dictionary(Of String, Object) = Nothing) As DataTable
            Dim dt As New DataTable()
            Using conn As DbConnection = GetConnection()
                conn.Open()
                Using cmd As DbCommand = CreateCommand(conn, query)
                    If parameters IsNot Nothing Then
                        For Each kvp In parameters
                            AddParam(cmd, kvp.Key, kvp.Value)
                        Next
                    End If
                    Using reader As DbDataReader = cmd.ExecuteReader()
                        dt.Load(reader)
                    End Using
                End Using
            End Using
            Return dt
        End Function

        Public Function ExecuteNonQuery(query As String, Optional parameters As Dictionary(Of String, Object) = Nothing) As Integer
            Using conn As DbConnection = GetConnection()
                conn.Open()
                Using cmd As DbCommand = CreateCommand(conn, query)
                    If parameters IsNot Nothing Then
                        For Each kvp In parameters
                            AddParam(cmd, kvp.Key, kvp.Value)
                        Next
                    End If
                    Return cmd.ExecuteNonQuery()
                End Using
            End Using
        End Function

        Public Function ExecuteScalar(query As String, Optional parameters As Dictionary(Of String, Object) = Nothing) As Object
            Using conn As DbConnection = GetConnection()
                conn.Open()
                Using cmd As DbCommand = CreateCommand(conn, query)
                    If parameters IsNot Nothing Then
                        For Each kvp In parameters
                            AddParam(cmd, kvp.Key, kvp.Value)
                        Next
                    End If
                    Return cmd.ExecuteScalar()
                End Using
            End Using
        End Function

        ''' <summary>
        ''' Genera hash SHA256 seguro para contraseñas de usuarios
        ''' </summary>
        Public Function HashPassword(password As String) As String
            If String.IsNullOrEmpty(password) Then Return String.Empty
            Using sha As SHA256 = SHA256.Create()
                Dim bytes As Byte() = sha.ComputeHash(Encoding.UTF8.GetBytes(password))
                Dim sb As New StringBuilder()
                For Each b In bytes
                    sb.Append(b.ToString("x2"))
                Next
                Return sb.ToString()
            End Using
        End Function

        ''' <summary>
        ''' Inicializa automáticamente la base de datos (SQLite o MySQL) y crea las tablas si aún no existen
        ''' </summary>
        Public Function InitializeDatabaseAndTables(ByRef outMessage As String) As Boolean
            Try
                If IsSQLite Then
                    ' Inicialización nativa SQLite
                    Using conn As DbConnection = GetConnection()
                        conn.Open()

                        Dim baseDir As String = AppDomain.CurrentDomain.BaseDirectory
                        Dim searchPaths As String() = {
                            Path.Combine(baseDir, "Database", "schema_sqlite.sql"),
                            Path.Combine(baseDir, "..", "..", "..", "..", "Database", "schema_sqlite.sql"),
                            Path.Combine(baseDir, "..", "..", "..", "Database", "schema_sqlite.sql"),
                            "c:\Users\gonza\Desktop\Proyecto\Database\schema_sqlite.sql"
                        }

                        Dim schemaFile As String = ""
                        For Each p In searchPaths
                            Dim fullPath = Path.GetFullPath(p)
                            If File.Exists(fullPath) Then
                                schemaFile = fullPath
                                Exit For
                            End If
                        Next

                        If Not String.IsNullOrEmpty(schemaFile) Then
                            Dim sql = File.ReadAllText(schemaFile)
                            Using cmd = conn.CreateCommand()
                                cmd.CommandText = sql
                                cmd.ExecuteNonQuery()
                            End Using

                            ' Sembrar si usuarios está vacío
                            Using cmdCount = conn.CreateCommand()
                                cmdCount.CommandText = "SELECT COUNT(*) FROM usuarios;"
                                Dim count As Long = Convert.ToInt64(cmdCount.ExecuteScalar())
                                If count = 0 Then
                                    Dim seedFile = Path.Combine(Path.GetDirectoryName(schemaFile), "seed_sqlite.sql")
                                    If File.Exists(seedFile) Then
                                        Dim seedSql = File.ReadAllText(seedFile)
                                        Using cmdSeed = conn.CreateCommand()
                                            cmdSeed.CommandText = seedSql
                                            cmdSeed.ExecuteNonQuery()
                                        End Using
                                    End If
                                End If
                            End Using
                        End If
                    End Using

                    outMessage = "Base de datos SQLite local inicializada exitosamente."
                    Return True
                Else
                    ' Inicialización MySQL
                    Using serverConn As New MySqlConnection(AppConfig.Settings.GetServerOnlyConnectionString())
                        serverConn.Open()
                        Dim createDbQuery As String = $"CREATE DATABASE IF NOT EXISTS `{AppConfig.Settings.Database}` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;"
                        Using cmd As New MySqlCommand(createDbQuery, serverConn)
                            cmd.ExecuteNonQuery()
                        End Using
                    End Using

                    Using dbConn As MySqlConnection = New MySqlConnection(AppConfig.ConnectionString)
                        dbConn.Open()

                        Dim baseDir As String = AppDomain.CurrentDomain.BaseDirectory
                        Dim searchPaths As String() = {
                            Path.Combine(baseDir, "Database", "schema_mysql.sql"),
                            Path.Combine(baseDir, "..", "..", "..", "..", "Database", "schema_mysql.sql"),
                            Path.Combine(baseDir, "..", "..", "..", "Database", "schema_mysql.sql"),
                            "c:\Users\gonza\Desktop\Proyecto\Database\schema_mysql.sql"
                        }

                        Dim schemaFile As String = ""
                        For Each p In searchPaths
                            Dim fullPath = Path.GetFullPath(p)
                            If File.Exists(fullPath) Then
                                schemaFile = fullPath
                                Exit For
                            End If
                        Next

                        If Not String.IsNullOrEmpty(schemaFile) Then
                            Dim schemaSql As String = File.ReadAllText(schemaFile)
                            Using cmd As New MySqlCommand(schemaSql, dbConn)
                                cmd.CommandTimeout = 120
                                cmd.ExecuteNonQuery()
                            End Using

                            Dim countCmd As New MySqlCommand("SELECT COUNT(*) FROM `usuarios`;", dbConn)
                            Dim userCount As Long = Convert.ToInt64(countCmd.ExecuteScalar())
                            If userCount = 0 Then
                                Dim seedFile = Path.Combine(Path.GetDirectoryName(schemaFile), "seed_data.sql")
                                If File.Exists(seedFile) Then
                                    Dim seedSql As String = File.ReadAllText(seedFile)
                                    Using cmdSeed As New MySqlCommand(seedSql, dbConn)
                                        cmdSeed.CommandTimeout = 120
                                        cmdSeed.ExecuteNonQuery()
                                    End Using
                                End If
                            End If
                        End If
                    End Using

                    outMessage = "Base de datos MySQL verificada e inicializada exitosamente."
                    Return True
                End If
            Catch ex As Exception
                outMessage = "Error al inicializar la base de datos: " & ex.Message
                Return False
            End Try
        End Function

    End Module
End Namespace
