Imports System.Data
Imports MySqlConnector
Imports GestionComercial.Data
Imports GestionComercial.Models

Namespace Services
    Public Class CatalogService

#Region "Categorías"
        Public Function GetCategorias(Optional soloActivas As Boolean = True) As List(Of Categoria)
            Dim list As New List(Of Categoria)()
            Dim query As String = "SELECT * FROM `categorias` " & If(soloActivas, "WHERE `activo` = 1 ", "") & "ORDER BY `nombre` ASC;"
            Dim dt As DataTable = DatabaseHelper.ExecuteQuery(query)

            For Each row As DataRow In dt.Rows
                list.Add(New Categoria() With {
                    .Id = Convert.ToInt32(row("id")),
                    .Nombre = row("nombre").ToString(),
                    .Descripcion = If(IsDBNull(row("descripcion")), "", row("descripcion").ToString()),
                    .Activo = Convert.ToBoolean(row("activo"))
                })
            Next
            Return list
        End Function

        Public Function GuardarCategoria(cat As Categoria, ByRef errorMessage As String) As Boolean
            Try
                Dim query As String
                Dim params As New Dictionary(Of String, Object) From {
                    {"@nombre", cat.Nombre.Trim()},
                    {"@desc", cat.Descripcion.Trim()},
                    {"@activo", If(cat.Activo, 1, 0)}
                }

                If cat.Id = 0 Then
                    query = "INSERT INTO `categorias` (`nombre`, `descripcion`, `activo`) VALUES (@nombre, @desc, @activo);"
                Else
                    query = "UPDATE `categorias` SET `nombre` = @nombre, `descripcion` = @desc, `activo` = @activo WHERE `id` = @id;"
                    params.Add("@id", cat.Id)
                End If

                DatabaseHelper.ExecuteNonQuery(query, params)
                errorMessage = String.Empty
                Return True
            Catch ex As Exception
                errorMessage = ex.Message
                Return False
            End Try
        End Function
#End Region

#Region "Talles"
        Public Function GetTalles() As List(Of Talle)
            Dim list As New List(Of Talle)()
            Dim query As String = "SELECT * FROM `talles` ORDER BY `orden` ASC, `id` ASC;"
            Dim dt As DataTable = DatabaseHelper.ExecuteQuery(query)

            For Each row As DataRow In dt.Rows
                list.Add(New Talle() With {
                    .Id = Convert.ToInt32(row("id")),
                    .Nombre = row("nombre").ToString(),
                    .Orden = Convert.ToInt32(row("orden"))
                })
            Next
            Return list
        End Function

        Public Function GuardarTalle(talle As Talle, ByRef errorMessage As String) As Boolean
            Try
                Dim query As String
                Dim params As New Dictionary(Of String, Object) From {
                    {"@nombre", talle.Nombre.Trim()},
                    {"@orden", talle.Orden}
                }

                If talle.Id = 0 Then
                    query = "INSERT INTO `talles` (`nombre`, `orden`) VALUES (@nombre, @orden);"
                Else
                    query = "UPDATE `talles` SET `nombre` = @nombre, `orden` = @orden WHERE `id` = @id;"
                    params.Add("@id", talle.Id)
                End If

                DatabaseHelper.ExecuteNonQuery(query, params)
                errorMessage = String.Empty
                Return True
            Catch ex As Exception
                errorMessage = ex.Message
                Return False
            End Try
        End Function
#End Region

#Region "Productos y Matriz de Stock"
        Public Function GetProductos(Optional filtro As String = "", Optional categoriaId As Integer = 0, Optional soloActivos As Boolean = True) As List(Of Producto)
            Dim list As New List(Of Producto)()
            Dim query As String = "SELECT p.*, c.nombre AS categoria_nombre, IFNULL(SUM(pt.stock_actual), 0) AS total_stock " &
                                 "FROM `productos` p " &
                                 "INNER JOIN `categorias` c ON p.categoria_id = c.id " &
                                 "LEFT JOIN `producto_talles` pt ON p.id = pt.producto_id " &
                                 "WHERE 1=1 "

            Dim params As New Dictionary(Of String, Object)()

            If soloActivos Then
                query &= "AND p.activo = 1 "
            End If

            If categoriaId > 0 Then
                query &= "AND p.categoria_id = @categoriaId "
                params.Add("@categoriaId", categoriaId)
            End If

            If Not String.IsNullOrWhiteSpace(filtro) Then
                query &= "AND (p.nombre LIKE @filtro OR p.codigo_barra LIKE @filtro) "
                params.Add("@filtro", "%" & filtro.Trim() & "%")
            End If

            query &= "GROUP BY p.id ORDER BY p.nombre ASC;"

            Dim dt As DataTable = DatabaseHelper.ExecuteQuery(query, params)

            For Each row As DataRow In dt.Rows
                Dim prod As New Producto() With {
                    .Id = Convert.ToInt32(row("id")),
                    .CodigoBarra = row("codigo_barra").ToString(),
                    .Nombre = row("nombre").ToString(),
                    .Descripcion = If(IsDBNull(row("descripcion")), "", row("descripcion").ToString()),
                    .CategoriaId = Convert.ToInt32(row("categoria_id")),
                    .CategoriaNombre = row("categoria_nombre").ToString(),
                    .PrecioCosto = Convert.ToDecimal(row("precio_costo")),
                    .PrecioVenta = Convert.ToDecimal(row("precio_venta")),
                    .PorcentajeGanancia = Convert.ToDecimal(row("porcentaje_ganancia")),
                    .ImagenRuta = If(IsDBNull(row("imagen_ruta")), "", row("imagen_ruta").ToString()),
                    .Activo = Convert.ToBoolean(row("activo")),
                    .TotalStock = Convert.ToInt32(row("total_stock"))
                }
                list.Add(prod)
            Next
            Return list
        End Function

        Public Function GetProductoPorCodigo(codigo As String) As Producto
            Dim query As String = "SELECT p.*, c.nombre AS categoria_nombre, IFNULL(SUM(pt.stock_actual), 0) AS total_stock " &
                                 "FROM `productos` p " &
                                 "INNER JOIN `categorias` c ON p.categoria_id = c.id " &
                                 "LEFT JOIN `producto_talles` pt ON p.id = pt.producto_id " &
                                 "WHERE (p.codigo_barra = @cod OR p.id = @codInt) AND p.activo = 1 " &
                                 "GROUP BY p.id LIMIT 1;"

            Dim codInt As Integer = 0
            Integer.TryParse(codigo, codInt)

            Dim params As New Dictionary(Of String, Object) From {
                {"@cod", codigo.Trim()},
                {"@codInt", codInt}
            }

            Dim dt As DataTable = DatabaseHelper.ExecuteQuery(query, params)
            If dt.Rows.Count = 0 Then Return Nothing

            Dim row As DataRow = dt.Rows(0)
            Dim prod As New Producto() With {
                .Id = Convert.ToInt32(row("id")),
                .CodigoBarra = row("codigo_barra").ToString(),
                .Nombre = row("nombre").ToString(),
                .Descripcion = If(IsDBNull(row("descripcion")), "", row("descripcion").ToString()),
                .CategoriaId = Convert.ToInt32(row("categoria_id")),
                .CategoriaNombre = row("categoria_nombre").ToString(),
                .PrecioCosto = Convert.ToDecimal(row("precio_costo")),
                .PrecioVenta = Convert.ToDecimal(row("precio_venta")),
                .PorcentajeGanancia = Convert.ToDecimal(row("porcentaje_ganancia")),
                .ImagenRuta = If(IsDBNull(row("imagen_ruta")), "", row("imagen_ruta").ToString()),
                .Activo = Convert.ToBoolean(row("activo")),
                .TotalStock = Convert.ToInt32(row("total_stock")),
                .TallesStock = GetStockPorTalles(Convert.ToInt32(row("id")))
            }
            Return prod
        End Function

        Public Function GetStockPorTalles(productoId As Integer) As List(Of ProductoTalle)
            Dim list As New List(Of ProductoTalle)()
            Dim query As String = "SELECT pt.*, t.nombre AS talle_nombre " &
                                 "FROM `producto_talles` pt " &
                                 "INNER JOIN `talles` t ON pt.talle_id = t.id " &
                                 "WHERE pt.producto_id = @prodId " &
                                 "ORDER BY t.orden ASC, t.id ASC;"

            Dim params As New Dictionary(Of String, Object) From {{"@prodId", productoId}}
            Dim dt As DataTable = DatabaseHelper.ExecuteQuery(query, params)

            For Each row As DataRow In dt.Rows
                list.Add(New ProductoTalle() With {
                    .Id = Convert.ToInt32(row("id")),
                    .ProductoId = Convert.ToInt32(row("producto_id")),
                    .TalleId = Convert.ToInt32(row("talle_id")),
                    .TalleNombre = row("talle_nombre").ToString(),
                    .Color = row("color").ToString(),
                    .StockActual = Convert.ToInt32(row("stock_actual")),
                    .StockMinimo = Convert.ToInt32(row("stock_minimo")),
                    .SkuEspecifico = If(IsDBNull(row("sku_especifico")), "", row("sku_especifico").ToString())
                })
            Next
            Return list
        End Function

        Public Function GuardarProducto(prod As Producto, ByRef errorMessage As String) As Boolean
            Try
                Using conn As Common.DbConnection = DatabaseHelper.GetConnection()
                    conn.Open()
                    Using trans As Common.DbTransaction = conn.BeginTransaction()
                        Try
                            Dim queryProd As String
                            If prod.Id = 0 Then
                                queryProd = "INSERT INTO `productos` (`codigo_barra`, `nombre`, `descripcion`, `categoria_id`, `precio_costo`, `precio_venta`, `porcentaje_ganancia`, `imagen_ruta`, `activo`) " &
                                            "VALUES (@cod, @nom, @desc, @cat, @costo, @venta, @gan, @img, @activo);"
                            Else
                                queryProd = "UPDATE `productos` SET `codigo_barra` = @cod, `nombre` = @nom, `descripcion` = @desc, " &
                                            "`categoria_id` = @cat, `precio_costo` = @costo, `precio_venta` = @venta, `porcentaje_ganancia` = @gan, " &
                                            "`imagen_ruta` = @img, `activo` = @activo WHERE `id` = @id;"
                            End If

                            Using cmdProd = DatabaseHelper.CreateCommand(conn, queryProd, trans)
                                DatabaseHelper.AddParam(cmdProd, "@cod", prod.CodigoBarra.Trim())
                                DatabaseHelper.AddParam(cmdProd, "@nom", prod.Nombre.Trim())
                                DatabaseHelper.AddParam(cmdProd, "@desc", prod.Descripcion.Trim())
                                DatabaseHelper.AddParam(cmdProd, "@cat", prod.CategoriaId)
                                DatabaseHelper.AddParam(cmdProd, "@costo", prod.PrecioCosto)
                                DatabaseHelper.AddParam(cmdProd, "@venta", prod.PrecioVenta)
                                DatabaseHelper.AddParam(cmdProd, "@gan", prod.PorcentajeGanancia)
                                DatabaseHelper.AddParam(cmdProd, "@img", If(String.IsNullOrEmpty(prod.ImagenRuta), DBNull.Value, prod.ImagenRuta))
                                DatabaseHelper.AddParam(cmdProd, "@activo", If(prod.Activo, 1, 0))
                                If prod.Id > 0 Then
                                    DatabaseHelper.AddParam(cmdProd, "@id", prod.Id)
                                End If
                                cmdProd.ExecuteNonQuery()

                                If prod.Id = 0 Then
                                    prod.Id = Convert.ToInt32(DatabaseHelper.GetLastInsertedId(conn, trans))
                                End If
                            End Using

                            ' Guardar / Actualizar la matriz de talles y stock
                            If prod.TallesStock IsNot Nothing Then
                                For Each pt In prod.TallesStock
                                    Dim queryPt As String
                                    If DatabaseHelper.IsSQLite Then
                                        queryPt = "INSERT INTO `producto_talles` (`producto_id`, `talle_id`, `color`, `stock_actual`, `stock_minimo`, `sku_especifico`) " &
                                                  "VALUES (@prodId, @talleId, @color, @stock, @stockMin, @sku) " &
                                                  "ON CONFLICT(`producto_id`, `talle_id`, `color`) DO UPDATE SET `stock_actual` = @stock, `stock_minimo` = @stockMin, `sku_especifico` = @sku;"
                                    Else
                                        queryPt = "INSERT INTO `producto_talles` (`producto_id`, `talle_id`, `color`, `stock_actual`, `stock_minimo`, `sku_especifico`) " &
                                                  "VALUES (@prodId, @talleId, @color, @stock, @stockMin, @sku) " &
                                                  "ON DUPLICATE KEY UPDATE `stock_actual` = @stock, `stock_minimo` = @stockMin, `sku_especifico` = @sku;"
                                    End If

                                    Using cmdPt = DatabaseHelper.CreateCommand(conn, queryPt, trans)
                                        DatabaseHelper.AddParam(cmdPt, "@prodId", prod.Id)
                                        DatabaseHelper.AddParam(cmdPt, "@talleId", pt.TalleId)
                                        DatabaseHelper.AddParam(cmdPt, "@color", If(String.IsNullOrWhiteSpace(pt.Color), "Único", pt.Color.Trim()))
                                        DatabaseHelper.AddParam(cmdPt, "@stock", pt.StockActual)
                                        DatabaseHelper.AddParam(cmdPt, "@stockMin", pt.StockMinimo)
                                        DatabaseHelper.AddParam(cmdPt, "@sku", If(String.IsNullOrEmpty(pt.SkuEspecifico), $"{prod.CodigoBarra}-{pt.TalleId}", pt.SkuEspecifico))
                                        cmdPt.ExecuteNonQuery()
                                    End Using
                                Next
                            End If

                            trans.Commit()
                            errorMessage = String.Empty
                            Return True
                        Catch exTrans As Exception
                            trans.Rollback()
                            errorMessage = "Error en transacción de producto: " & exTrans.Message
                            Return False
                        End Try
                    End Using
                End Using
            Catch ex As Exception
                errorMessage = ex.Message
                Return False
            End Try
        End Function

        Public Function AjustarStockManual(productoId As Integer, talleId As Integer, color As String, cantidadDelta As Integer, motivo As String, usuarioId As Integer, ByRef errorMessage As String) As Boolean
            Try
                Using conn As Common.DbConnection = DatabaseHelper.GetConnection()
                    conn.Open()
                    Using trans As Common.DbTransaction = conn.BeginTransaction()
                        Try
                            ' Obtener stock actual
                            Dim querySelect As String = "SELECT stock_actual FROM `producto_talles` WHERE `producto_id` = @prodId AND `talle_id` = @talleId AND `color` = @color FOR UPDATE;"
                            Dim stockActual As Integer = 0
                            Using cmdSelect = DatabaseHelper.CreateCommand(conn, querySelect, trans)
                                DatabaseHelper.AddParam(cmdSelect, "@prodId", productoId)
                                DatabaseHelper.AddParam(cmdSelect, "@talleId", talleId)
                                DatabaseHelper.AddParam(cmdSelect, "@color", color)
                                Dim res = cmdSelect.ExecuteScalar()
                                If res IsNot Nothing AndAlso Not IsDBNull(res) Then
                                    stockActual = Convert.ToInt32(res)
                                Else
                                    ' Crear fila con 0 si no existía
                                    Dim queryInsertPt As String = "INSERT INTO `producto_talles` (`producto_id`, `talle_id`, `color`, `stock_actual`, `stock_minimo`) VALUES (@prodId, @talleId, @color, 0, 2);"
                                    Using cmdIns = DatabaseHelper.CreateCommand(conn, queryInsertPt, trans)
                                        DatabaseHelper.AddParam(cmdIns, "@prodId", productoId)
                                        DatabaseHelper.AddParam(cmdIns, "@talleId", talleId)
                                        DatabaseHelper.AddParam(cmdIns, "@color", color)
                                        cmdIns.ExecuteNonQuery()
                                    End Using
                                    stockActual = 0
                                End If
                            End Using

                            Dim nuevoStock As Integer = stockActual + cantidadDelta
                            If nuevoStock < 0 Then nuevoStock = 0

                            ' Actualizar stock
                            Dim queryUpdate As String = "UPDATE `producto_talles` SET `stock_actual` = @nuevoStock WHERE `producto_id` = @prodId AND `talle_id` = @talleId AND `color` = @color;"
                            Using cmdUpd = DatabaseHelper.CreateCommand(conn, queryUpdate, trans)
                                DatabaseHelper.AddParam(cmdUpd, "@nuevoStock", nuevoStock)
                                DatabaseHelper.AddParam(cmdUpd, "@prodId", productoId)
                                DatabaseHelper.AddParam(cmdUpd, "@talleId", talleId)
                                DatabaseHelper.AddParam(cmdUpd, "@color", color)
                                cmdUpd.ExecuteNonQuery()
                            End Using

                            ' Registrar en auditoría de stock
                            Dim queryMov As String = "INSERT INTO `movimientos_stock` (`producto_id`, `talle_id`, `color`, `tipo_movimiento`, `cantidad`, `stock_anterior`, `stock_posterior`, `motivo`, `usuario_id`) " &
                                                    "VALUES (@prodId, @talleId, @color, @tipo, @cant, @ant, @post, @motivo, @userId);"
                            Using cmdMov = DatabaseHelper.CreateCommand(conn, queryMov, trans)
                                DatabaseHelper.AddParam(cmdMov, "@prodId", productoId)
                                DatabaseHelper.AddParam(cmdMov, "@talleId", talleId)
                                DatabaseHelper.AddParam(cmdMov, "@color", color)
                                DatabaseHelper.AddParam(cmdMov, "@tipo", If(cantidadDelta >= 0, "Ingreso_Compra", "Ajuste_Manual"))
                                DatabaseHelper.AddParam(cmdMov, "@cant", Math.Abs(cantidadDelta))
                                DatabaseHelper.AddParam(cmdMov, "@ant", stockActual)
                                DatabaseHelper.AddParam(cmdMov, "@post", nuevoStock)
                                DatabaseHelper.AddParam(cmdMov, "@motivo", motivo)
                                DatabaseHelper.AddParam(cmdMov, "@userId", usuarioId)
                                cmdMov.ExecuteNonQuery()
                            End Using

                            trans.Commit()
                            errorMessage = String.Empty
                            Return True
                        Catch exTrans As Exception
                            trans.Rollback()
                            errorMessage = exTrans.Message
                            Return False
                        End Try
                    End Using
                End Using
            Catch ex As Exception
                errorMessage = ex.Message
                Return False
            End Try
        End Function

        Public Function EliminarProducto(id As Integer, ByRef errorMessage As String) As Boolean
            Try
                ' Soft delete para mantener integridad de ventas pasadas
                Dim query As String = "UPDATE `productos` SET `activo` = 0 WHERE `id` = @id;"
                DatabaseHelper.ExecuteNonQuery(query, New Dictionary(Of String, Object) From {{"@id", id}})
                errorMessage = String.Empty
                Return True
            Catch ex As Exception
                errorMessage = ex.Message
                Return False
            End Try
        End Function
#End Region

    End Class
End Namespace
