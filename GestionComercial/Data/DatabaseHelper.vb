' ARCHIVO: DatabaseHelper.vb
' PROPÓSITO: Capa de Acceso a Datos (DAL), utilidades de base de datos y seguridad criptográfica.
' Aquí centralizamos la persistencia y la seguridad del sistema:
' 1. Adaptador de Persistencia (SQLite):
'    El sistema gestiona la conectividad a la base de datos relacional SQLite usando SqliteConnection,
'    DbCommand y DbDataReader de Microsoft.Data.Sqlite y System.Data.Common para abstracción de comandos.
' 2. Seguridad contra Inyección SQL:
'    Todas las consultas utilizan parámetros tipados (AddParam) para prevenir ataques de SQL Injection.
' 3. Seguridad Criptográfica de Contraseñas (PBKDF2):
'    Hashea contraseñas con Rfc2898DeriveBytes, 100.000 iteraciones, HMAC-SHA-256 y Salt aleatorio de 16 bytes.
'    La verificación se realiza en tiempo constante (FixedTimeEquals) para mitigar ataques de temporización (Timing Attacks).
' 4. Inicialización Inteligente de Tablas y Datos Semilla:
'    Si el archivo de base de datos no existe o está vacío, ejecuta automáticamente los scripts de esquema
'    y datos semilla para dejar el sistema listo para operar de inmediato sin configuración manual.
' 5. Auto-migraciones de Esquema:
'    Verifica dinámicamente si faltan columnas (como datos personales o fecha de nacimiento)
'    y las incorpora con ALTER TABLE de forma no destructiva sin perder registros previos.

Imports System.Data
Imports System.Data.Common
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports Microsoft.Data.Sqlite
Imports GestionComercial.Config

Namespace Data
    Public Module DatabaseHelper

        ' Bandera booleana rápida que confirma que trabajamos con SQLite como motor de persistencia.
        Public ReadOnly Property IsSQLite As Boolean
            Get
                Return True
            End Get
        End Property

        ' Fábrica de conexiones: Devuelve una instancia de SqliteConnection conectada
        ' al archivo de base de datos local según appsettings.json.
        ' Nos aseguramos de que la carpeta de destino exista antes de abrir.
        Public Function GetConnection() As DbConnection
            ' Aseguramos que la carpeta Database donde se aloja el archivo .db exista
            Dim dbPath = AppConfig.Settings.GetSqlitePath()
            Dim dir = Path.GetDirectoryName(dbPath)
            If Not String.IsNullOrEmpty(dir) AndAlso Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If
            Return New SqliteConnection(AppConfig.ConnectionString)
        End Function

        ' Adaptador de sintaxis SQL: Convierte funciones de fecha a sintaxis nativa de SQLite
        ' (datetime(...) y date(...)) y remueve cláusulas incompatibles como "FOR UPDATE".
        Public Function PrepareQuery(query As String) As String
            Dim q = query
            q = q.Replace("NOW()", "datetime('now', 'localtime')", StringComparison.OrdinalIgnoreCase)
            q = q.Replace("CURDATE()", "date('now', 'localtime')", StringComparison.OrdinalIgnoreCase)
            q = q.Replace("FOR UPDATE", "", StringComparison.OrdinalIgnoreCase)
            Return q
        End Function

        ' Método de verificación de conectividad al iniciar el sistema.
        ' Si el archivo físico o las tablas aún no existen, ejecuta la inicialización
        ' automática de tablas y corre las migraciones de esquema pendientes.
        Public Function TestConnection(ByRef errorMessage As String) As Boolean
            Try
                ' Si no existe el archivo físico, corremos la inicialización inicial
                Dim dbPath = AppConfig.Settings.GetSqlitePath()
                If Not File.Exists(dbPath) Then
                    Dim initMsg As String = ""
                    InitializeDatabaseAndTables(initMsg)
                End If

                Using conn As DbConnection = GetConnection()
                    conn.Open()

                    ' Si aún no tiene la tabla usuarios, sembramos las tablas
                    Using cmdCheck = conn.CreateCommand()
                        cmdCheck.CommandText = "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='usuarios';"
                        Dim tblCount As Long = Convert.ToInt64(cmdCheck.ExecuteScalar())
                        If tblCount = 0 Then
                            Dim initMsg As String = ""
                            InitializeDatabaseAndTables(initMsg)
                        End If
                    End Using

                    ' Verificamos que todas las columnas necesarias existan (migración de esquema)
                    EnsureSchemaUpdates(conn)

                    errorMessage = String.Empty
                    Return True
                End Using
            Catch ex As Exception
                errorMessage = ex.Message
                Return False
            End Try
        End Function

        ' Crea un comando SQL limpio asociando la conexión activa y la transacción opcional.
        ' Pasa la consulta por PrepareQuery para compatibilidad de funciones de fecha.
        Public Function CreateCommand(conn As DbConnection, query As String, Optional trans As DbTransaction = Nothing) As DbCommand
            Dim cmd As DbCommand = conn.CreateCommand()
            cmd.CommandText = PrepareQuery(query)
            If trans IsNot Nothing Then cmd.Transaction = trans
            Return cmd
        End Function

        ' Protección contra Inyección SQL (SQL Injection):
        ' Todo valor ingresado por el usuario se asigna como parámetro formal. Si el valor
        ' es Nothing, le asignamos DBNull.Value para no romper columnas de base de datos.
        Public Sub AddParam(cmd As DbCommand, name As String, value As Object)
            Dim p = cmd.CreateParameter()
            p.ParameterName = name
            p.Value = If(value, DBNull.Value)
            cmd.Parameters.Add(p)
        End Sub

        ' Recupera el ID autoincremental generado tras un INSERT dentro de una transacción.
        ' En SQLite se consulta last_insert_rowid().
        Public Function GetLastInsertedId(conn As DbConnection, trans As DbTransaction) As Long
            Using cmd As DbCommand = conn.CreateCommand()
                cmd.Transaction = trans
                cmd.CommandText = "SELECT last_insert_rowid();"
                Return Convert.ToInt64(cmd.ExecuteScalar())
            End Using
        End Function

        ' Ejecuta una consulta SELECT y devuelve los resultados en un DataTable en memoria.
        ' Llenamos los datos columna por columna manualmente para evitar ConstraintExceptions
        ' si una columna única contiene valores nulos en SQLite.
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

        ' Ejecuta sentencias INSERT, UPDATE o DELETE y retorna la cantidad de filas afectadas.
        ' El bloque Using garantiza que la conexión física se cierre siempre al terminar.
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

        ' Ejecuta consultas que devuelven un único dato escalar (ej: COUNT(*), SUM(...), etc.).
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

        ' Genera el hash de contraseña utilizando PBKDF2 con sal aleatoria y 100.000 iteraciones
        Public Function HashPasswordSecure(password As String) As String
            If String.IsNullOrEmpty(password) Then Return String.Empty
            Const iterations As Integer = 100000
            Const saltSize As Integer = 16
            Const keySize As Integer = 32

            Dim salt As Byte() = RandomNumberGenerator.GetBytes(saltSize)
            Dim subKey As Byte() = Rfc2898DeriveBytes.Pbkdf2(password, salt, iterations, HashAlgorithmName.SHA256, keySize)

            Return $"PBKDF2$SHA256${iterations}${Convert.ToBase64String(salt)}${Convert.ToBase64String(subKey)}"
        End Function

        ' Valida la contraseña comparando el hash PBKDF2 con FixedTimeEquals para prevenir timing attacks
        Public Function VerifyPassword(password As String, storedHash As String) As Boolean
            If String.IsNullOrEmpty(password) OrElse String.IsNullOrEmpty(storedHash) Then Return False

            Try
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
            Catch
                Return False
            End Try

            Return False
        End Function

        Public Function HashPassword(password As String) As String
            Return HashPasswordSecure(password)
        End Function

        ' Resolución dinámica de scripts SQL (schema_sqlite, seed_data, etc.).
        ' Busca el archivo SQL en el directorio del binario y subiendo hasta 6 niveles
        ' hacia la carpeta raíz del proyecto, evitando rutas absolutas fijas.
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

        ' Inicialización de Base de Datos y Tablas.
        ' Si el software se ejecuta por primera vez en una máquina limpia:
        ' Crea las tablas ejecutando schema_sqlite.sql y si la tabla usuarios
        ' está vacía, ejecuta seed_sqlite.sql para tener usuarios y catálogo listos para operar.
        Public Function InitializeDatabaseAndTables(ByRef outMessage As String) As Boolean
            Try
                Using conn As DbConnection = GetConnection()
                    conn.Open()

                    Dim schemaFile As String = ResolveDatabaseScriptPath("schema_sqlite.sql")

                    If Not String.IsNullOrEmpty(schemaFile) Then
                        Dim sql = File.ReadAllText(schemaFile)
                        Using cmd = conn.CreateCommand()
                            cmd.CommandText = sql
                            cmd.ExecuteNonQuery()
                        End Using

                        ' Verificamos si usuarios está vacío para cargar las semillas de prueba
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
            Catch ex As Exception
                outMessage = "Error al inicializar la base de datos: " & ex.Message
                Return False
            End Try
        End Function

        ' Migraciones automáticas de esquema (Auto-Migrations):
        ' Detecta si faltan columnas de datos personales de usuarios o fecha de nacimiento
        ' de clientes y las agrega con ALTER TABLE sin perder los registros existentes.
        Private Sub EnsureSchemaUpdates(conn As DbConnection)
            Try
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
                                       "SELECT 'gerente', '10000003', 'ecfba551324356e5bd27b548adf36b728783f60d9b573d142caac7baad62be49', " &
                                       "'Gerente', 'General', 'Gerente General', 'Gerente', 1 " &
                                       "WHERE NOT EXISTS (SELECT 1 FROM usuarios WHERE username = 'gerente');"
                    cmd.ExecuteNonQuery()
                End Using
            Catch ex As Exception
                ' Error silencioso en auto-migración para no impedir inicio si ya fue aplicado
            End Try
        End Sub

    End Module
End Namespace
