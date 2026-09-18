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

                    ' Asegurar actualizaciones automáticas de esquema para columnas incorporadas
                    EnsureSchemaUpdates(conn)

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
                        ' Carga manual sin inferir restricciones en memoria de DataTable (previene ConstraintException en columnas UNIQUE con NULLs)
                        For i As Integer = 0 To reader.FieldCount - 1
                            Dim colType = reader.GetFieldType(i)
                            dt.Columns.Add(reader.GetName(i), If(colType, GetType(Object)))
                        Next
                        dt.BeginLoadData()
                        While reader.Read()
                            Dim row As DataRow = dt.NewRow()
                            For i As Integer = 0 To reader.FieldCount - 1
                                row(i) = reader.GetValue(i)
                            Next
                            dt.Rows.Add(row)
                        End While
                        dt.EndLoadData()
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
        ''' Genera hash seguro PBKDF2 con Salt aleatorio y 100.000 iteraciones (Estándar OWASP)
        ''' Formato: PBKDF2$SHA256${iterations}${saltBase64}${hashBase64}
        ''' </summary>
        Public Function HashPasswordSecure(password As String) As String
            If String.IsNullOrEmpty(password) Then Return String.Empty
            Const iterations As Integer = 100000
            Const saltSize As Integer = 16
            Const keySize As Integer = 32

            Dim salt As Byte() = RandomNumberGenerator.GetBytes(saltSize)
            Dim subKey As Byte() = Rfc2898DeriveBytes.Pbkdf2(password, salt, iterations, HashAlgorithmName.SHA256, keySize)

            Return $"PBKDF2$SHA256${iterations}${Convert.ToBase64String(salt)}${Convert.ToBase64String(subKey)}"
        End Function

        ''' <summary>
        ''' Hash legacy SHA-256 (mantenido para verificar y migrar contraseñas existentes)
        ''' </summary>
        Public Function LegacyHashPassword(password As String) As String
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
        ''' Valida una contraseña contra un hash almacenado (soporta PBKDF2 y legado SHA-256).
        ''' Si el hash es legado y la clave es correcta, needsRehash se pone en True para auto-migrar.
        ''' Utiliza CryptographicOperations.FixedTimeEquals para prevenir ataques de temporización (timing attacks).
        ''' </summary>
        Public Function VerifyPassword(password As String, storedHash As String, ByRef needsRehash As Boolean) As Boolean
            needsRehash = False
            If String.IsNullOrEmpty(password) OrElse String.IsNullOrEmpty(storedHash) Then Return False

            Try
                ' 1. Caso formato moderno: PBKDF2$SHA256${iterations}${salt}${hash}
                If storedHash.StartsWith("PBKDF2$SHA256$", StringComparison.OrdinalIgnoreCase) Then
                    Dim parts As String() = storedHash.Split("$"c)
                    If parts.Length >= 5 Then
                        Dim iterations As Integer = 100000
                        Integer.TryParse(parts(2), iterations)
                        Dim salt As Byte() = Convert.FromBase64String(parts(3))
                        Dim expectedKey As Byte() = Convert.FromBase64String(parts(4))
                        Dim actualKey As Byte() = Rfc2898DeriveBytes.Pbkdf2(password, salt, iterations, HashAlgorithmName.SHA256, expectedKey.Length)

                        Return CryptographicOperations.FixedTimeEquals(actualKey, expectedKey)
                    End If
                End If

                ' 2. Caso formato legado: SHA-256 hexadecimal (64 caracteres)
                Dim legacyExpected = LegacyHashPassword(password)
                Dim match = CryptographicOperations.FixedTimeEquals(
                    Encoding.UTF8.GetBytes(legacyExpected.ToLowerInvariant()),
                    Encoding.UTF8.GetBytes(storedHash.Trim().ToLowerInvariant())
                )
                If match Then
                    needsRehash = True
                    Return True
                End If
            Catch
                Return False
            End Try

            Return False
        End Function

        ''' <summary>
        ''' Mapeo predeterminado de hash para nuevos usuarios o cambios de contraseña
        ''' </summary>
        Public Function HashPassword(password As String) As String
            Return HashPasswordSecure(password)
        End Function

        ''' <summary>
        ''' Resuelve dinámicamente la ubicación de scripts SQL sin rutas hardcodeadas
        ''' </summary>
        Public Function ResolveDatabaseScriptPath(fileName As String) As String
            ' 1. Directorio base del ejecutable / bin
            Dim baseDir As String = AppDomain.CurrentDomain.BaseDirectory
            Dim p1 = Path.Combine(baseDir, "Database", fileName)
            If File.Exists(p1) Then Return Path.GetFullPath(p1)

            Dim p2 = Path.Combine(baseDir, fileName)
            If File.Exists(p2) Then Return Path.GetFullPath(p2)

            ' 2. Búsqueda recursiva ascendente hacia la raíz del proyecto (hasta 6 niveles)
            Dim currentDir As DirectoryInfo = New DirectoryInfo(baseDir)
            For i As Integer = 0 To 5
                If currentDir Is Nothing Then Exit For
                Dim candDir = Path.Combine(currentDir.FullName, "Database", fileName)
                If File.Exists(candDir) Then Return Path.GetFullPath(candDir)

                Dim candRoot = Path.Combine(currentDir.FullName, fileName)
                If File.Exists(candRoot) Then Return Path.GetFullPath(candRoot)

                currentDir = currentDir.Parent
            Next

            ' 3. Directorio de trabajo actual
            Dim cwd = Environment.CurrentDirectory
            Dim pCwd = Path.Combine(cwd, "Database", fileName)
            If File.Exists(pCwd) Then Return Path.GetFullPath(pCwd)

            Return String.Empty
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

                        Dim schemaFile As String = ResolveDatabaseScriptPath("schema_sqlite.sql")

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
                                    Dim seedFile = ResolveDatabaseScriptPath("seed_sqlite.sql")
                                    If Not String.IsNullOrEmpty(seedFile) AndAlso File.Exists(seedFile) Then
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

                        Dim schemaFile As String = ResolveDatabaseScriptPath("schema_mysql.sql")

                        If Not String.IsNullOrEmpty(schemaFile) Then
                            Dim schemaSql As String = File.ReadAllText(schemaFile)
                            Using cmd As New MySqlCommand(schemaSql, dbConn)
                                cmd.CommandTimeout = 120
                                cmd.ExecuteNonQuery()
                            End Using

                            Dim roleMigrationSql As String = "ALTER TABLE `usuarios` MODIFY COLUMN `rol` ENUM('Administrador', 'Gerente', 'Vendedor', 'Cajero') NOT NULL DEFAULT 'Vendedor';"
                            Using roleMigrationCmd As New MySqlCommand(roleMigrationSql, dbConn)
                                roleMigrationCmd.ExecuteNonQuery()
                            End Using

                            Dim countCmd As New MySqlCommand("SELECT COUNT(*) FROM `usuarios`;", dbConn)
                            Dim userCount As Long = Convert.ToInt64(countCmd.ExecuteScalar())
                            If userCount = 0 Then
                                Dim seedFile = ResolveDatabaseScriptPath("seed_data.sql")
                                If Not String.IsNullOrEmpty(seedFile) AndAlso File.Exists(seedFile) Then
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

        ''' <summary>
        ''' Aplica migraciones y ajustes automáticos de esquema para garantizar retrocompatibilidad con bases de datos existentes.
        ''' </summary>
        Private Sub EnsureSchemaUpdates(conn As DbConnection)
            Try
                If IsSQLite Then
                    Using cmd = conn.CreateCommand()
                        cmd.CommandText = "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='usuarios';"
                        Dim hasTable = Convert.ToInt64(cmd.ExecuteScalar()) > 0
                        If Not hasTable Then Return

                        ' Columnas a asegurar en usuarios
                        Dim requiredCols As String() = {
                            "nombre TEXT NOT NULL DEFAULT ''",
                            "apellido TEXT NOT NULL DEFAULT ''",
                            "telefono TEXT NULL",
                            "email TEXT NULL",
                            "direccion TEXT NULL",
                            "ciudad TEXT NULL",
                            "notas TEXT NULL",
                            "fecha_nacimiento TEXT NULL",
                            "dni TEXT NULL"
                        }

                        For Each colDef In requiredCols
                            Dim colName = colDef.Split(" "c)(0)
                            cmd.CommandText = "PRAGMA table_info(usuarios);"
                            Dim exists As Boolean = False
                            Using reader = cmd.ExecuteReader()
                                While reader.Read()
                                    If reader("name").ToString().Equals(colName, StringComparison.OrdinalIgnoreCase) Then
                                        exists = True
                                        Exit While
                                    End If
                                End While
                            End Using

                            If Not exists Then
                                Using cmdAlter = conn.CreateCommand()
                                    cmdAlter.CommandText = $"ALTER TABLE usuarios ADD COLUMN {colDef};"
                                    cmdAlter.ExecuteNonQuery()
                                End Using
                            End If
                        Next

                        ' Migración de datos existentes
                        cmd.CommandText = "UPDATE usuarios SET nombre = nombre_completo WHERE (nombre = '' OR nombre IS NULL) AND nombre_completo <> '';"
                        cmd.ExecuteNonQuery()

                        cmd.CommandText = "UPDATE usuarios SET dni = '1000000' || id WHERE dni IS NULL OR dni = '';"
                        cmd.ExecuteNonQuery()

                        cmd.CommandText = "CREATE UNIQUE INDEX IF NOT EXISTS uk_usuarios_dni ON usuarios(dni);"
                        cmd.ExecuteNonQuery()

                        ' Columna en clientes
                        cmd.CommandText = "PRAGMA table_info(clientes);"
                        Dim hasFnacCli As Boolean = False
                        Using readerCli = cmd.ExecuteReader()
                            While readerCli.Read()
                                If readerCli("name").ToString().Equals("fecha_nacimiento", StringComparison.OrdinalIgnoreCase) Then
                                    hasFnacCli = True
                                    Exit While
                                End If
                            End While
                        End Using

                        If Not hasFnacCli Then
                            Using cmdAlter = conn.CreateCommand()
                                cmdAlter.CommandText = "ALTER TABLE clientes ADD COLUMN fecha_nacimiento TEXT NULL;"
                                cmdAlter.ExecuteNonQuery()
                            End Using
                        End If

                        ' Agregar el gerente inicial a bases creadas con una semilla anterior.
                        cmd.CommandText = "INSERT INTO usuarios (username, dni, password_hash, nombre, apellido, nombre_completo, rol, activo) " &
                                           "SELECT 'gerente', NULL, 'ecfba551324356e5bd27b548adf36b728783f60d9b573d142caac7baad62be49', " &
                                           "'Gerente', 'General', 'Gerente General', 'Gerente', 1 " &
                                           "WHERE NOT EXISTS (SELECT 1 FROM usuarios WHERE username = 'gerente');"
                        cmd.ExecuteNonQuery()
                    End Using
                Else
                    Using cmd = conn.CreateCommand()
                        cmd.CommandText = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'usuarios' AND COLUMN_NAME = 'dni';"
                        Dim count = Convert.ToInt64(cmd.ExecuteScalar())
                        If count = 0 Then
                            Using cmdAlter = conn.CreateCommand()
                                cmdAlter.CommandText = "ALTER TABLE `usuarios` ADD COLUMN `dni` VARCHAR(20) UNIQUE NULL AFTER `username`;"
                                cmdAlter.ExecuteNonQuery()
                            End Using
                        End If
                    End Using
                End If
            Catch ex As Exception
                ' Error silencioso en auto-migración para no impedir inicio si ya fue aplicado
            End Try
        End Sub

    End Module
End Namespace
