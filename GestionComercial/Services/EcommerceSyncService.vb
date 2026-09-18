Imports System.Data.Common
Imports MySqlConnector
Imports GestionComercial.Data

Namespace Services
    Public Class EcommerceSyncService
        Public Function TestEcommerceConnection(host As String, port As Integer, username As String, password As String, databaseName As String, ByRef errorMessage As String) As Boolean
            Try
                Using connection = CreateEcommerceConnection(host, port, username, password, databaseName)
                    connection.Open()
                    Using command = connection.CreateCommand()
                        command.CommandText = "SELECT 1;"
                        command.ExecuteScalar()
                    End Using
                End Using

                errorMessage = String.Empty
                Return True
            Catch ex As Exception
                errorMessage = ex.Message
                Return False
            End Try
        End Function

        Public Function GetEcommerceStats(host As String, port As Integer, username As String, password As String, databaseName As String) As Dictionary(Of String, Integer)
            Dim stats = New Dictionary(Of String, Integer) From {
                {"categorias", 0},
                {"talles", 0},
                {"productos", 0},
                {"stock", 0}
            }

            Try
                Using connection = CreateEcommerceConnection(host, port, username, password, databaseName)
                    connection.Open()
                    stats("categorias") = ReadCount(connection, "SELECT COUNT(*) FROM `categorias`;")
                    stats("talles") = ReadCount(connection, "SELECT COUNT(*) FROM `talles`;")
                    stats("productos") = ReadCount(connection, "SELECT COUNT(*) FROM `productos`;")
                    stats("stock") = ReadCount(connection, "SELECT COALESCE(SUM(`stock`), 0) FROM `producto_talle`;")
                End Using
            Catch
                Return stats
            End Try

            Return stats
        End Function

        Public Function ImportCatalogFromEcommerce(host As String, port As Integer, username As String, password As String, databaseName As String, ByRef report As String) As Boolean
            Try
                Using sourceConnection = CreateEcommerceConnection(host, port, username, password, databaseName)
                    sourceConnection.Open()

                    Using targetConnection As DbConnection = DatabaseHelper.GetConnection()
                        targetConnection.Open()
                        Using transaction = targetConnection.BeginTransaction()
                            Try
                                Dim categoryCount = ImportCategories(sourceConnection, targetConnection, transaction)
                                Dim sizeCount = ImportSizes(sourceConnection, targetConnection, transaction)
                                Dim productCount = ImportProducts(sourceConnection, targetConnection, transaction)
                                Dim stockCount = ImportStock(sourceConnection, targetConnection, transaction)
                                transaction.Commit()

                                report = $"Importación completada: {categoryCount} categorías, {sizeCount} talles, {productCount} productos y {stockCount} relaciones de stock."
                                Return True
                            Catch
                                transaction.Rollback()
                                Throw
                            End Try
                        End Using
                    End Using
                End Using
            Catch ex As Exception
                report = "No se pudo completar la importación: " & ex.Message
                Return False
            End Try
        End Function

        Private Function CreateEcommerceConnection(host As String, port As Integer, username As String, password As String, databaseName As String) As MySqlConnection
            Dim builder = New MySqlConnectionStringBuilder With {
                .Server = If(String.IsNullOrWhiteSpace(host), "localhost", host.Trim()),
                .Port = CUInt(Math.Max(1, port)),
                .UserID = username,
                .Password = password,
                .Database = databaseName,
                .CharacterSet = "utf8mb4",
                .ConnectionTimeout = 10,
                .AllowUserVariables = True
            }
            Return New MySqlConnection(builder.ConnectionString)
        End Function

        Private Function ReadCount(connection As MySqlConnection, query As String) As Integer
            Using command = connection.CreateCommand()
                command.CommandText = query
                Return Convert.ToInt32(command.ExecuteScalar())
            End Using
        End Function

        Private Function ImportCategories(sourceConnection As MySqlConnection, targetConnection As DbConnection, transaction As DbTransaction) As Integer
            Dim count = 0
            Using sourceCommand = sourceConnection.CreateCommand()
                sourceCommand.CommandText = "SELECT `id`, `nombre`, `slug` FROM `categorias` ORDER BY `id`;"
                Using reader = sourceCommand.ExecuteReader()
                    While reader.Read()
                        Dim categoryId = Convert.ToInt32(reader("id"))
                        Dim name = ReadText(reader, "nombre")
                        Dim slug = ReadText(reader, "slug")
                        Dim description = If(String.IsNullOrWhiteSpace(slug), "Importado desde e-commerce", "Importado desde e-commerce (slug: " & slug & ")")
                        UpsertCategory(targetConnection, transaction, categoryId, name, description)
                        count += 1
                    End While
                End Using
            End Using
            Return count
        End Function

        Private Function ImportSizes(sourceConnection As MySqlConnection, targetConnection As DbConnection, transaction As DbTransaction) As Integer
            Dim count = 0
            Using sourceCommand = sourceConnection.CreateCommand()
                sourceCommand.CommandText = "SELECT `id`, `nombre` FROM `talles` ORDER BY `id`;"
                Using reader = sourceCommand.ExecuteReader()
                    While reader.Read()
                        Dim sizeId = Convert.ToInt32(reader("id"))
                        UpsertSize(targetConnection, transaction, sizeId, ReadText(reader, "nombre"), sizeId)
                        count += 1
                    End While
                End Using
            End Using
            Return count
        End Function

        Private Function ImportProducts(sourceConnection As MySqlConnection, targetConnection As DbConnection, transaction As DbTransaction) As Integer
            Dim count = 0
            Using sourceCommand = sourceConnection.CreateCommand()
                sourceCommand.CommandText = "SELECT `id`, `nombre`, `descripcion`, `categoria_id`, `precio`, `imagen_ruta`, `activo` FROM `productos` ORDER BY `id`;"
                Using reader = sourceCommand.ExecuteReader()
                    While reader.Read()
                        Dim productId = Convert.ToInt32(reader("id"))
                        Dim price = ReadDecimal(reader, "precio")
                        Dim barcode = "779" & productId.ToString("00000000")
                        UpsertProduct(targetConnection, transaction, productId, barcode, ReadText(reader, "nombre"), ReadText(reader, "descripcion"), Convert.ToInt32(reader("categoria_id")), price * 0.5D, price, ReadText(reader, "imagen_ruta"), ReadBoolean(reader, "activo"))
                        count += 1
                    End While
                End Using
            End Using
            Return count
        End Function

        Private Function ImportStock(sourceConnection As MySqlConnection, targetConnection As DbConnection, transaction As DbTransaction) As Integer
            Dim count = 0
            Using sourceCommand = sourceConnection.CreateCommand()
                sourceCommand.CommandText = "SELECT `producto_id`, `talle_id`, `stock` FROM `producto_talle` ORDER BY `producto_id`, `talle_id`;"
                Using reader = sourceCommand.ExecuteReader()
                    While reader.Read()
                        UpsertStock(targetConnection, transaction, Convert.ToInt32(reader("producto_id")), Convert.ToInt32(reader("talle_id")), Convert.ToInt32(reader("stock")))
                        count += 1
                    End While
                End Using
            End Using
            Return count
        End Function

        Private Sub UpsertCategory(connection As DbConnection, transaction As DbTransaction, id As Integer, name As String, description As String)
            If Exists(connection, transaction, "categorias", id) Then
                ExecuteTarget(connection, transaction, "UPDATE `categorias` SET `nombre` = @name, `descripcion` = @description, `activo` = 1 WHERE `id` = @id;", {"@id", id}, {"@name", name}, {"@description", description})
                Return
            End If
            ExecuteTarget(connection, transaction, "INSERT INTO `categorias` (`id`, `nombre`, `descripcion`, `activo`) VALUES (@id, @name, @description, 1);", {"@id", id}, {"@name", name}, {"@description", description})
        End Sub

        Private Sub UpsertSize(connection As DbConnection, transaction As DbTransaction, id As Integer, name As String, orderValue As Integer)
            If Exists(connection, transaction, "talles", id) Then
                ExecuteTarget(connection, transaction, "UPDATE `talles` SET `nombre` = @name, `orden` = @orderValue WHERE `id` = @id;", {"@id", id}, {"@name", name}, {"@orderValue", orderValue})
                Return
            End If
            ExecuteTarget(connection, transaction, "INSERT INTO `talles` (`id`, `nombre`, `orden`) VALUES (@id, @name, @orderValue);", {"@id", id}, {"@name", name}, {"@orderValue", orderValue})
        End Sub

        Private Sub UpsertProduct(connection As DbConnection, transaction As DbTransaction, id As Integer, barcode As String, name As String, description As String, categoryId As Integer, cost As Decimal, price As Decimal, imagePath As String, active As Boolean)
            Dim parameters = {"@id", id, "@barcode", barcode, "@name", name, "@description", description, "@categoryId", categoryId, "@cost", cost, "@price", price, "@imagePath", If(String.IsNullOrWhiteSpace(imagePath), DBNull.Value, imagePath), "@active", If(active, 1, 0)}
            If Exists(connection, transaction, "productos", id) Then
                ExecuteTarget(connection, transaction, "UPDATE `productos` SET `codigo_barra` = @barcode, `nombre` = @name, `descripcion` = @description, `categoria_id` = @categoryId, `precio_costo` = @cost, `precio_venta` = @price, `porcentaje_ganancia` = 100, `imagen_ruta` = @imagePath, `activo` = @active WHERE `id` = @id;", parameters)
                Return
            End If
            ExecuteTarget(connection, transaction, "INSERT INTO `productos` (`id`, `codigo_barra`, `nombre`, `descripcion`, `categoria_id`, `precio_costo`, `precio_venta`, `porcentaje_ganancia`, `imagen_ruta`, `activo`) VALUES (@id, @barcode, @name, @description, @categoryId, @cost, @price, 100, @imagePath, @active);", parameters)
        End Sub

        Private Sub UpsertStock(connection As DbConnection, transaction As DbTransaction, productId As Integer, sizeId As Integer, stock As Integer)
            Dim existsQuery = "SELECT COUNT(*) FROM `producto_talles` WHERE `producto_id` = @productId AND `talle_id` = @sizeId AND `color` = 'Único';"
            Dim existsParameters = {"@productId", productId, "@sizeId", sizeId}
            If Convert.ToInt32(ExecuteTargetScalar(connection, transaction, existsQuery, existsParameters)) > 0 Then
                ExecuteTarget(connection, transaction, "UPDATE `producto_talles` SET `stock_actual` = @stock WHERE `producto_id` = @productId AND `talle_id` = @sizeId AND `color` = 'Único';", {"@productId", productId}, {"@sizeId", sizeId}, {"@stock", stock})
                Return
            End If
            ExecuteTarget(connection, transaction, "INSERT INTO `producto_talles` (`producto_id`, `talle_id`, `color`, `stock_actual`, `stock_minimo`, `sku_especifico`) VALUES (@productId, @sizeId, 'Único', @stock, 2, @sku);", {"@productId", productId}, {"@sizeId", sizeId}, {"@stock", stock}, {"@sku", $"SKU-{productId}-{sizeId}"})
        End Sub

        Private Function Exists(connection As DbConnection, transaction As DbTransaction, tableName As String, id As Integer) As Boolean
            Dim query = "SELECT COUNT(*) FROM `" & tableName & "` WHERE `id` = @id;"
            Return Convert.ToInt32(ExecuteTargetScalar(connection, transaction, query, {"@id", id})) > 0
        End Function

        Private Function ExecuteTargetScalar(connection As DbConnection, transaction As DbTransaction, query As String, ParamArray parameters() As Object) As Object
            Using command = DatabaseHelper.CreateCommand(connection, query, transaction)
                AddParameters(command, parameters)
                Return command.ExecuteScalar()
            End Using
        End Function

        Private Sub ExecuteTarget(connection As DbConnection, transaction As DbTransaction, query As String, ParamArray parameters() As Object)
            Using command = DatabaseHelper.CreateCommand(connection, query, transaction)
                AddParameters(command, parameters)
                command.ExecuteNonQuery()
            End Using
        End Sub

        Private Sub AddParameters(command As DbCommand, parameters() As Object)
            For index As Integer = 0 To parameters.Length - 1 Step 2
                DatabaseHelper.AddParam(command, CStr(parameters(index)), parameters(index + 1))
            Next
        End Sub

        Private Function ReadText(reader As DbDataReader, columnName As String) As String
            If reader.IsDBNull(reader.GetOrdinal(columnName)) Then Return String.Empty
            Return Convert.ToString(reader(columnName)).Trim()
        End Function

        Private Function ReadDecimal(reader As DbDataReader, columnName As String) As Decimal
            If reader.IsDBNull(reader.GetOrdinal(columnName)) Then Return 0D
            Return Convert.ToDecimal(reader(columnName))
        End Function

        Private Function ReadBoolean(reader As DbDataReader, columnName As String) As Boolean
            If reader.IsDBNull(reader.GetOrdinal(columnName)) Then Return True
            Return Convert.ToBoolean(reader(columnName))
        End Function
    End Class
End Namespace
