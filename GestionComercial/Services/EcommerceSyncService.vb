Imports System.Data
Imports System.Data.Common
Imports MySqlConnector
Imports GestionComercial.Data

Namespace Services
    Public Class EcommerceSyncService

        Public Function TestEcommerceConnection(host As String, port As Integer, user As String, pass As String, dbName As String, ByRef errorMessage As String) As Boolean
            Try
                Dim builder As New MySqlConnectionStringBuilder() With {
                    .Server = host,
                    .Port = CUInt(port),
                    .Database = dbName,
                    .UserID = user,
                    .Password = pass,
                    .CharacterSet = "utf8mb4",
                    .ConnectionTimeout = 5
                }

                Using conn As New MySqlConnection(builder.ConnectionString)
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
            Dim stats As New Dictionary(Of String, Integer) From {
                {"categorias", 0},
                {"talles", 0},
                {"productos", 0},
                {"stock", 0}
            }

            Try
                Dim builder As New MySqlConnectionStringBuilder() With {
                    .Server = host,
                    .Port = CUInt(port),
                    .Database = dbName,
                    .UserID = user,
                    .Password = pass,
                    .CharacterSet = "utf8mb4"
                }

                Using conn As New MySqlConnection(builder.ConnectionString)
                    conn.Open()

                    ' Contar categorías
                    Using cmd As New MySqlCommand("SELECT COUNT(*) FROM `categorias`;", conn)
                        stats("categorias") = Convert.ToInt32(cmd.ExecuteScalar())
                    End Using

                    ' Contar talles
                    Using cmd As New MySqlCommand("SELECT COUNT(*) FROM `talles`;", conn)
                        stats("talles") = Convert.ToInt32(cmd.ExecuteScalar())
                    End Using

                    ' Contar productos
                    Using cmd As New MySqlCommand("SELECT COUNT(*) FROM `productos`;", conn)
                        stats("productos") = Convert.ToInt32(cmd.ExecuteScalar())
                    End Using

                    ' Sumar stock si existe producto_talle
                    Using cmd As New MySqlCommand("SELECT IFNULL(SUM(stock), 0) FROM `producto_talle`;", conn)
                        stats("stock") = Convert.ToInt32(cmd.ExecuteScalar())
                    End Using
                End Using
            Catch ex As Exception
                ' Devuelve lo que pudo obtener
            End Try

            Return stats
        End Function

        Public Function ImportCatalogFromEcommerce(host As String, port As Integer, user As String, pass As String, dbName As String, ByRef outReport As String) As Boolean
            Try
                Dim srcBuilder As New MySqlConnectionStringBuilder() With {
                    .Server = host,
                    .Port = CUInt(port),
                    .Database = dbName,
                    .UserID = user,
                    .Password = pass,
                    .CharacterSet = "utf8mb4"
                }

                Dim catCount As Integer = 0
                Dim talleCount As Integer = 0
                Dim prodCount As Integer = 0
                Dim stockRelCount As Integer = 0

                ' 1. Leer datos desde la BD del E-commerce
                Dim dtCategorias As New DataTable()
                Dim dtTalles As New DataTable()
                Dim dtProductos As New DataTable()
                Dim dtStock As New DataTable()

                Using srcConn As New MySqlConnection(srcBuilder.ConnectionString)
                    srcConn.Open()

                    Using cmd As New MySqlCommand("SELECT * FROM `categorias`;", srcConn)
                        Using da As New MySqlDataAdapter(cmd)
                            da.Fill(dtCategorias)
                        End Using
                    End Using

                    Using cmd As New MySqlCommand("SELECT * FROM `talles`;", srcConn)
                        Using da As New MySqlDataAdapter(cmd)
                            da.Fill(dtTalles)
                        End Using
                    End Using

                    Using cmd As New MySqlCommand("SELECT * FROM `productos`;", srcConn)
                        Using da As New MySqlDataAdapter(cmd)
                            da.Fill(dtProductos)
                        End Using
                    End Using

                    Try
                        Using cmd As New MySqlCommand("SELECT * FROM `producto_talle`;", srcConn)
                            Using da As New MySqlDataAdapter(cmd)
                                da.Fill(dtStock)
                            End Using
                        End Using
                    Catch
                        ' Si no existe producto_talle, continuamos
                    End Try
                End Using

                ' 2. Insertar/Actualizar en la base de datos de Gestión Comercial
                Using dstConn As DbConnection = DatabaseHelper.GetConnection()
                    dstConn.Open()
                    Using trans As DbTransaction = dstConn.BeginTransaction()
                        Try
                            Dim isSqlite = DatabaseHelper.IsSQLite

                            ' Importar categorías
                            For Each row As DataRow In dtCategorias.Rows
                                Dim q As String
                                If isSqlite Then
                                    q = "INSERT INTO `categorias` (`id`, `nombre`, `descripcion`, `activo`) " &
                                        "VALUES (@id, @nom, @desc, 1) " &
                                        "ON CONFLICT(`id`) DO UPDATE SET `nombre` = excluded.`nombre`;"
                                Else
                                    q = "INSERT INTO `categorias` (`id`, `nombre`, `descripcion`, `activo`) " &
                                        "VALUES (@id, @nom, @desc, 1) " &
                                        "ON DUPLICATE KEY UPDATE `nombre` = @nom;"
                                End If

                                Using cmd = DatabaseHelper.CreateCommand(dstConn, q, trans)
                                    DatabaseHelper.AddParam(cmd, "@id", row("id"))
                                    DatabaseHelper.AddParam(cmd, "@nom", row("nombre").ToString())
                                    DatabaseHelper.AddParam(cmd, "@desc", "Importado desde E-commerce")
                                    cmd.ExecuteNonQuery()
                                    catCount += 1
                                End Using
                            Next

                            ' Importar talles
                            For Each row As DataRow In dtTalles.Rows
                                Dim q As String
                                If isSqlite Then
                                    q = "INSERT INTO `talles` (`id`, `nombre`, `orden`) " &
                                        "VALUES (@id, @nom, @id) " &
                                        "ON CONFLICT(`id`) DO UPDATE SET `nombre` = excluded.`nombre`;"
                                Else
                                    q = "INSERT INTO `talles` (`id`, `nombre`, `orden`) " &
                                        "VALUES (@id, @nom, @id) " &
                                        "ON DUPLICATE KEY UPDATE `nombre` = @nom;"
                                End If

                                Using cmd = DatabaseHelper.CreateCommand(dstConn, q, trans)
                                    DatabaseHelper.AddParam(cmd, "@id", row("id"))
                                    DatabaseHelper.AddParam(cmd, "@nom", row("nombre").ToString())
                                    cmd.ExecuteNonQuery()
                                    talleCount += 1
                                End Using
                            Next

                            ' Importar productos
                            For Each row As DataRow In dtProductos.Rows
                                Dim prodId As Integer = Convert.ToInt32(row("id"))
                                Dim nombre As String = row("nombre").ToString()
                                Dim precioWeb As Decimal = Convert.ToDecimal(row("precio"))
                                Dim costoEstimado As Decimal = Math.Round(precioWeb * 0.5D, 2)
                                Dim codigoBarra As String = "779" & prodId.ToString("D6")
                                Dim catId As Integer = Convert.ToInt32(row("categoria_id"))

                                Dim q As String
                                If isSqlite Then
                                    q = "INSERT INTO `productos` (`id`, `codigo_barra`, `nombre`, `descripcion`, `categoria_id`, `precio_costo`, `precio_venta`, `porcentaje_ganancia`, `activo`) " &
                                        "VALUES (@id, @cod, @nom, @desc, @catId, @costo, @venta, 100, 1) " &
                                        "ON CONFLICT(`id`) DO UPDATE SET `nombre` = excluded.`nombre`, `precio_venta` = excluded.`precio_venta`, `categoria_id` = excluded.`categoria_id`;"
                                Else
                                    q = "INSERT INTO `productos` (`id`, `codigo_barra`, `nombre`, `descripcion`, `categoria_id`, `precio_costo`, `precio_venta`, `porcentaje_ganancia`, `activo`) " &
                                        "VALUES (@id, @cod, @nom, @desc, @catId, @costo, @venta, 100, 1) " &
                                        "ON DUPLICATE KEY UPDATE `nombre` = @nom, `precio_venta` = @venta, `categoria_id` = @catId;"
                                End If

                                Using cmd = DatabaseHelper.CreateCommand(dstConn, q, trans)
                                    DatabaseHelper.AddParam(cmd, "@id", prodId)
                                    DatabaseHelper.AddParam(cmd, "@cod", codigoBarra)
                                    DatabaseHelper.AddParam(cmd, "@nom", nombre)
                                    DatabaseHelper.AddParam(cmd, "@desc", If(IsDBNull(row("descripcion")), "", row("descripcion").ToString()))
                                    DatabaseHelper.AddParam(cmd, "@catId", catId)
                                    DatabaseHelper.AddParam(cmd, "@costo", costoEstimado)
                                    DatabaseHelper.AddParam(cmd, "@venta", precioWeb)
                                    cmd.ExecuteNonQuery()
                                    prodCount += 1
                                End Using
                            Next

                            ' Importar stock por talle
                            For Each row As DataRow In dtStock.Rows
                                Dim pId As Integer = Convert.ToInt32(row("producto_id"))
                                Dim tId As Integer = Convert.ToInt32(row("talle_id"))
                                Dim stock As Integer = Convert.ToInt32(row("stock"))

                                Dim q As String
                                If isSqlite Then
                                    q = "INSERT INTO `producto_talles` (`producto_id`, `talle_id`, `color`, `stock_actual`, `stock_minimo`, `sku_especifico`) " &
                                        "VALUES (@pId, @tId, 'Único', @stock, 2, @sku) " &
                                        "ON CONFLICT(`producto_id`, `talle_id`, `color`) DO UPDATE SET `stock_actual` = excluded.`stock_actual`;"
                                Else
                                    q = "INSERT INTO `producto_talles` (`producto_id`, `talle_id`, `color`, `stock_actual`, `stock_minimo`, `sku_especifico`) " &
                                        "VALUES (@pId, @tId, 'Único', @stock, 2, @sku) " &
                                        "ON DUPLICATE KEY UPDATE `stock_actual` = @stock;"
                                End If

                                Using cmd = DatabaseHelper.CreateCommand(dstConn, q, trans)
                                    DatabaseHelper.AddParam(cmd, "@pId", pId)
                                    DatabaseHelper.AddParam(cmd, "@tId", tId)
                                    DatabaseHelper.AddParam(cmd, "@stock", stock)
                                    DatabaseHelper.AddParam(cmd, "@sku", $"WEB-{pId}-{tId}")
                                    cmd.ExecuteNonQuery()
                                    stockRelCount += 1
                                End Using
                            Next

                            trans.Commit()
                            outReport = $"Migración exitosa: {catCount} categorías, {talleCount} talles, {prodCount} productos y {stockRelCount} registros de stock importados correctamente."
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

    End Class
End Namespace
