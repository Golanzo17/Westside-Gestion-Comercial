Imports System.Data
Imports System.Data.Common
Imports MySqlConnector
Imports GestionComercial.Data

Namespace Services
    Public Class EcommerceSyncService

        Public Function TestEcommerceConnection(host As String, port As Integer, user As String, pass As String, dbName As String, ByRef errorMessage As String) As Boolean
            Dim hasInvalidParams = String.IsNullOrWhiteSpace(host) OrElse port <= 0 OrElse port > 65535 OrElse String.IsNullOrWhiteSpace(dbName)
            If hasInvalidParams Then
                errorMessage = "Los parámetros de conexión son inválidos. Verifique host, puerto y base de datos."
                Return False
            End If

            Try
                Dim builder = CreateConnectionStringBuilder(host, port, user, pass, dbName, 5)
                Using conn = New MySqlConnection(builder.ConnectionString)
                    conn.Open()
                    errorMessage = String.Empty
                    Return True
                End Using
            Catch ex As Exception
                errorMessage = ex.Message
                Return False
            End Try
        End Function

        Public Function GetEcommerceStats(host As String, port As Integer, user As String, pass As String, dbName As String) As Dictionary(Of String, Integer)
            Dim stats = New Dictionary(Of String, Integer) From {
                {"categorias", 0},
                {"talles", 0},
                {"productos", 0},
                {"stock", 0}
            }

            Dim hasInvalidParams = String.IsNullOrWhiteSpace(host) OrElse port <= 0 OrElse String.IsNullOrWhiteSpace(dbName)
            If hasInvalidParams Then
                Return stats
            End If

            Try
                Dim builder = CreateConnectionStringBuilder(host, port, user, pass, dbName)
                Using conn = New MySqlConnection(builder.ConnectionString)
                    conn.Open()
                    stats("categorias") = ExecuteScalarInt(conn, "SELECT COUNT(*) FROM `categorias`;")
                    stats("talles") = ExecuteScalarInt(conn, "SELECT COUNT(*) FROM `talles`;")
                    stats("productos") = ExecuteScalarInt(conn, "SELECT COUNT(*) FROM `productos`;")
                    stats("stock") = ExecuteScalarInt(conn, "SELECT IFNULL(SUM(stock), 0) FROM `producto_talle`;")
                End Using
            Catch
                ' Devuelve los conteos que se hayan podido obtener
            End Try

            Return stats
        End Function

        Public Function ImportCatalogFromEcommerce(host As String, port As Integer, user As String, pass As String, dbName As String, ByRef outReport As String) As Boolean
            Dim hasInvalidParams = String.IsNullOrWhiteSpace(host) OrElse port <= 0 OrElse String.IsNullOrWhiteSpace(dbName)
            If hasInvalidParams Then
                outReport = "Parámetros de conexión inválidos. Verifique servidor, puerto y base de datos."
                Return False
            End If

            Try
                Dim srcBuilder = CreateConnectionStringBuilder(host, port, user, pass, dbName)

                ' 1. Leer datos del origen E-Commerce
                Dim dtCategorias = New DataTable()
                Dim dtTalles = New DataTable()
                Dim dtProductos = New DataTable()
                Dim dtStock = New DataTable()

                Using srcConn = New MySqlConnection(srcBuilder.ConnectionString)
                    srcConn.Open()
                    FillDataTable(srcConn, "SELECT * FROM `categorias`;", dtCategorias)
                    FillDataTable(srcConn, "SELECT * FROM `talles`;", dtTalles)
                    FillDataTable(srcConn, "SELECT * FROM `productos`;", dtProductos)
                    TryFillDataTable(srcConn, "SELECT * FROM `producto_talle`;", dtStock)
                End Using

                ' 2. Persistir en la base local bajo transacción
                Using dstConn = DatabaseHelper.GetConnection()
                    dstConn.Open()
                    Using trans = dstConn.BeginTransaction()
                        Try
                            Dim isSqlite = DatabaseHelper.IsSQLite
                            Dim catCount = ImportCategorias(dstConn, trans, dtCategorias, isSqlite)
                            Dim talleCount = ImportTalles(dstConn, trans, dtTalles, isSqlite)
                            Dim prodCount = ImportProductos(dstConn, trans, dtProductos, isSqlite)
                            Dim stockCount = ImportStock(dstConn, trans, dtStock, isSqlite)

                            trans.Commit()
                            outReport = $"Migración exitosa: {catCount} categorías, {talleCount} talles, {prodCount} productos y {stockCount} registros de stock importados correctamente."
                            Return True
                        Catch exTrans As Exception
                            trans.Rollback()
                            outReport = "Error durante la migración en destino: " & exTrans.Message
                            Return False
                        End Try
                    End Using
                End Using

            Catch ex As Exception
                outReport = "Error de conexión o lectura: " & ex.Message
                Return False
            End Try
        End Function

        ' --- Helpers de infraestructura y consultas ---

        Private Shared Function CreateConnectionStringBuilder(host As String, port As Integer, user As String, pass As String, dbName As String, Optional timeoutSec As Integer = 15) As MySqlConnectionStringBuilder
            Return New MySqlConnectionStringBuilder() With {
                .Server = host,
                .Port = CUInt(port),
                .Database = dbName,
                .UserID = user,
                .Password = pass,
                .CharacterSet = "utf8mb4",
                .ConnectionTimeout = CUInt(timeoutSec)
            }
        End Function

        Private Shared Function ExecuteScalarInt(conn As MySqlConnection, sql As String) As Integer
            Using cmd = New MySqlCommand(sql, conn)
                Dim result = cmd.ExecuteScalar()
                If result Is Nothing OrElse Convert.IsDBNull(result) Then
                    Return 0
                End If
                Return Convert.ToInt32(result)
            End Using
        End Function

        Private Shared Sub FillDataTable(conn As MySqlConnection, query As String, table As DataTable)
            Using cmd = New MySqlCommand(query, conn)
                Using da = New MySqlDataAdapter(cmd)
                    da.Fill(table)
                End Using
            End Using
        End Sub

        Private Shared Sub TryFillDataTable(conn As MySqlConnection, query As String, table As DataTable)
            Try
                FillDataTable(conn, query, table)
            Catch
                ' Tabla opcional en esquemas anteriores
            End Try
        End Sub

        Private Shared Function ImportCategorias(conn As DbConnection, trans As DbTransaction, dt As DataTable, isSqlite As Boolean) As Integer
            Dim sql = If(isSqlite,
                "INSERT INTO `categorias` (`id`, `nombre`, `descripcion`, `activo`) VALUES (@id, @nom, @desc, 1) ON CONFLICT(`id`) DO UPDATE SET `nombre` = excluded.`nombre`;",
                "INSERT INTO `categorias` (`id`, `nombre`, `descripcion`, `activo`) VALUES (@id, @nom, @desc, 1) ON DUPLICATE KEY UPDATE `nombre` = @nom;")

            Dim count = 0
            For Each row As DataRow In dt.Rows
                Using cmd = DatabaseHelper.CreateCommand(conn, sql, trans)
                    DatabaseHelper.AddParam(cmd, "@id", row("id"))
                    DatabaseHelper.AddParam(cmd, "@nom", row("nombre").ToString())
                    DatabaseHelper.AddParam(cmd, "@desc", "Importado desde E-commerce")
                    cmd.ExecuteNonQuery()
                    count += 1
                End Using
            Next
            Return count
        End Function

        Private Shared Function ImportTalles(conn As DbConnection, trans As DbTransaction, dt As DataTable, isSqlite As Boolean) As Integer
            Dim sql = If(isSqlite,
                "INSERT INTO `talles` (`id`, `nombre`, `orden`) VALUES (@id, @nom, @id) ON CONFLICT(`id`) DO UPDATE SET `nombre` = excluded.`nombre`;",
                "INSERT INTO `talles` (`id`, `nombre`, `orden`) VALUES (@id, @nom, @id) ON DUPLICATE KEY UPDATE `nombre` = @nom;")

            Dim count = 0
            For Each row As DataRow In dt.Rows
                Using cmd = DatabaseHelper.CreateCommand(conn, sql, trans)
                    DatabaseHelper.AddParam(cmd, "@id", row("id"))
                    DatabaseHelper.AddParam(cmd, "@nom", row("nombre").ToString())
                    cmd.ExecuteNonQuery()
                    count += 1
                End Using
            Next
            Return count
        End Function

        Private Shared Function ImportProductos(conn As DbConnection, trans As DbTransaction, dt As DataTable, isSqlite As Boolean) As Integer
            Dim sql = If(isSqlite,
                "INSERT INTO `productos` (`id`, `codigo_barra`, `nombre`, `descripcion`, `categoria_id`, `precio_costo`, `precio_venta`, `porcentaje_ganancia`, `activo`) VALUES (@id, @cod, @nom, @desc, @catId, @costo, @venta, 100, 1) ON CONFLICT(`id`) DO UPDATE SET `nombre` = excluded.`nombre`, `precio_venta` = excluded.`precio_venta`, `categoria_id` = excluded.`categoria_id`;",
                "INSERT INTO `productos` (`id`, `codigo_barra`, `nombre`, `descripcion`, `categoria_id`, `precio_costo`, `precio_venta`, `porcentaje_ganancia`, `activo`) VALUES (@id, @cod, @nom, @desc, @catId, @costo, @venta, 100, 1) ON DUPLICATE KEY UPDATE `nombre` = @nom, `precio_venta` = @venta, `categoria_id` = @catId;")

            Dim count = 0
            For Each row As DataRow In dt.Rows
                Dim prodId = Convert.ToInt32(row("id"))
                Dim nombre = row("nombre").ToString()
                Dim precioWeb = Convert.ToDecimal(row("precio"))
                Dim costoEstimado = Math.Round(precioWeb * 0.5D, 2)
                Dim codigoBarra = "779" & prodId.ToString("D6")
                Dim catId = Convert.ToInt32(row("categoria_id"))
                Dim descripcion = If(row.IsNull("descripcion"), "", row("descripcion").ToString())

                Using cmd = DatabaseHelper.CreateCommand(conn, sql, trans)
                    DatabaseHelper.AddParam(cmd, "@id", prodId)
                    DatabaseHelper.AddParam(cmd, "@cod", codigoBarra)
                    DatabaseHelper.AddParam(cmd, "@nom", nombre)
                    DatabaseHelper.AddParam(cmd, "@desc", descripcion)
                    DatabaseHelper.AddParam(cmd, "@catId", catId)
                    DatabaseHelper.AddParam(cmd, "@costo", costoEstimado)
                    DatabaseHelper.AddParam(cmd, "@venta", precioWeb)
                    cmd.ExecuteNonQuery()
                    count += 1
                End Using
            Next
            Return count
        End Function

        Private Shared Function ImportStock(conn As DbConnection, trans As DbTransaction, dt As DataTable, isSqlite As Boolean) As Integer
            Dim sql = If(isSqlite,
                "INSERT INTO `producto_talles` (`producto_id`, `talle_id`, `color`, `stock_actual`, `stock_minimo`, `sku_especifico`) VALUES (@pId, @tId, 'Único', @stock, 2, @sku) ON CONFLICT(`producto_id`, `talle_id`, `color`) DO UPDATE SET `stock_actual` = excluded.`stock_actual`;",
                "INSERT INTO `producto_talles` (`producto_id`, `talle_id`, `color`, `stock_actual`, `stock_minimo`, `sku_especifico`) VALUES (@pId, @tId, 'Único', @stock, 2, @sku) ON DUPLICATE KEY UPDATE `stock_actual` = @stock;")

            Dim count = 0
            For Each row As DataRow In dt.Rows
                Dim pId = Convert.ToInt32(row("producto_id"))
                Dim tId = Convert.ToInt32(row("talle_id"))
                Dim stock = Convert.ToInt32(row("stock"))
                Dim sku = $"WEB-{pId}-{tId}"

                Using cmd = DatabaseHelper.CreateCommand(conn, sql, trans)
                    DatabaseHelper.AddParam(cmd, "@pId", pId)
                    DatabaseHelper.AddParam(cmd, "@tId", tId)
                    DatabaseHelper.AddParam(cmd, "@stock", stock)
                    DatabaseHelper.AddParam(cmd, "@sku", sku)
                    cmd.ExecuteNonQuery()
                    count += 1
                End Using
            Next
            Return count
        End Function

    End Class
End Namespace
