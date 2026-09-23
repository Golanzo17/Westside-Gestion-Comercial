' ARCHIVO: CatalogService.vb
' PROPÓSITO: Lógica de catálogo de indumentaria, talles, categorías y matriz de stock.
' Este servicio resuelve la estructura comercial de las prendas de ropa:
' 1. Matriz de Variantes: Una prenda no tiene una única cantidad fija; tiene múltiples
'    talles (S, M, L, XL, etc.) y colores registrados en la tabla producto_talles.
' 2. Búsqueda por Código de Barras o ID: Permite que el vendedor escanee la etiqueta
'    con una pistola láser en el mostrador para cargar la prenda al instante en el POS.
' 3. Ajustes Manuales de Inventario: Cuando llega mercadería de un proveedor o se detecta
'    una prenda rota, este servicio actualiza el stock y genera el movimiento de auditoría
'    asociando el motivo y el usuario responsable.


Imports System.Data
Imports System.Data.Common
Imports GestionComercial.Data
Imports GestionComercial.Models

Namespace Services
    Public Class CatalogService

#Region "Categorías"
        ' Retorna el listado de categorías de ropa (ej: Remeras, Jeans, Abrigos, etc.)

        Public Function GetCategorias(Optional soloActivas As Boolean = True) As List(Of Categoria)
            Dim list = New List(Of Categoria)()
            Dim query = "SELECT * FROM `categorias` " & If(soloActivas, "WHERE `activo` = 1 ", "") & "ORDER BY `nombre` ASC;"
            Dim dt = DatabaseHelper.ExecuteQuery(query)

            For Each row As DataRow In dt.Rows
                list.Add(MapCategoria(row))
            Next
            Return list
        End Function

        ' Guarda o modifica una categoría en la base de datos
        Public Function GuardarCategoria(cat As Categoria, ByRef errorMessage As String) As Boolean
            If cat Is Nothing OrElse String.IsNullOrWhiteSpace(cat.Nombre) Then
                errorMessage = "El nombre de la categoría es obligatorio."
                Return False
            End If

            Try
                Dim query As String
                Dim params = New Dictionary(Of String, Object) From {
                    {"@nombre", cat.Nombre.Trim()},
                    {"@desc", If(cat.Descripcion, "").Trim()},
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
        ' Retorna los talles ordenados por su campo "orden" (XS -> S -> M -> L -> XL) para visualización ordenada.

        Public Function GetTalles() As List(Of Talle)
            Dim list = New List(Of Talle)()
            Dim query = "SELECT * FROM `talles` ORDER BY `orden` ASC, `id` ASC;"
            Dim dt = DatabaseHelper.ExecuteQuery(query)

            For Each row As DataRow In dt.Rows
                list.Add(MapTalle(row))
            Next
            Return list
        End Function

        Public Function GuardarTalle(talle As Talle, ByRef errorMessage As String) As Boolean
            If talle Is Nothing OrElse String.IsNullOrWhiteSpace(talle.Nombre) Then
                errorMessage = "El nombre del talle es obligatorio."
                Return False
            End If

            Try
                Dim query As String
                Dim params = New Dictionary(Of String, Object) From {
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
        ' Consulta el catálogo de prendas calculando la suma total de stock a partir de producto_talles.

        Public Function GetProductos(Optional filtro As String = "", Optional categoriaId As Integer = 0, Optional soloActivos As Boolean = True) As List(Of Producto)
            Dim list = New List(Of Producto)()
            Dim query = "SELECT p.*, c.nombre AS categoria_nombre, IFNULL(SUM(pt.stock_actual), 0) AS total_stock " &
                        "FROM `productos` p " &
                        "INNER JOIN `categorias` c ON p.categoria_id = c.id " &
                        "LEFT JOIN `producto_talles` pt ON p.id = pt.producto_id " &
                        "WHERE 1=1 "

            Dim params = New Dictionary(Of String, Object)()

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

            Dim dt = DatabaseHelper.ExecuteQuery(query, params)
            For Each row As DataRow In dt.Rows
                list.Add(MapProducto(row, False))
            Next
            Return list
        End Function

        ' Búsqueda rápida para el Punto de Venta (Lector de Código de Barras o ID).
        ' Carga la prenda y sus variantes de talle y color correspondientes.

        Public Function GetProductoPorCodigo(codigo As String) As Producto
            If String.IsNullOrWhiteSpace(codigo) Then
                Return Nothing
            End If

            Dim codLimpio = codigo.Trim()
            Dim codInt = 0
            Integer.TryParse(codLimpio, codInt)

            Dim query = "SELECT p.*, c.nombre AS categoria_nombre, IFNULL(SUM(pt.stock_actual), 0) AS total_stock " &
                        "FROM `productos` p " &
                        "INNER JOIN `categorias` c ON p.categoria_id = c.id " &
                        "LEFT JOIN `producto_talles` pt ON p.id = pt.producto_id " &
                        "WHERE (p.codigo_barra = @cod OR p.id = @codInt) AND p.activo = 1 " &
                        "GROUP BY p.id LIMIT 1;"

            Dim params = New Dictionary(Of String, Object) From {
                {"@cod", codLimpio},
                {"@codInt", codInt}
            }

            Dim dt = DatabaseHelper.ExecuteQuery(query, params)
            If dt.Rows.Count = 0 Then
                Return Nothing
            End If

            Dim prod = MapProducto(dt.Rows(0), True)
            Return prod
        End Function

        Public Function GetStockPorTalles(productoId As Integer) As List(Of ProductoTalle)
            Dim list = New List(Of ProductoTalle)()
            If productoId <= 0 Then
                Return list
            End If

            Dim query = "SELECT pt.*, t.nombre AS talle_nombre " &
                        "FROM `producto_talles` pt " &
                        "INNER JOIN `talles` t ON pt.talle_id = t.id " &
                        "WHERE pt.producto_id = @prodId " &
                        "ORDER BY t.orden ASC, t.id ASC;"

            Dim params = New Dictionary(Of String, Object) From {{"@prodId", productoId}}
            Dim dt = DatabaseHelper.ExecuteQuery(query, params)

            For Each row As DataRow In dt.Rows
                list.Add(MapProductoTalle(row))
            Next
            Return list
        End Function

        Public Function GuardarProducto(prod As Producto, ByRef errorMessage As String) As Boolean
            If prod Is Nothing Then
                errorMessage = "Los datos del producto son inválidos."
                Return False
            End If

            Dim tieneDatosValidos = Not String.IsNullOrWhiteSpace(prod.Nombre) AndAlso prod.CategoriaId > 0
            If Not tieneDatosValidos Then
                errorMessage = "El producto debe tener un nombre y una categoría asignada."
                Return False
            End If

            Try
                Using conn = DatabaseHelper.GetConnection()
                    conn.Open()
                    Using trans = conn.BeginTransaction()
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
                                DatabaseHelper.AddParam(cmdProd, "@cod", If(prod.CodigoBarra, "").Trim())
                                DatabaseHelper.AddParam(cmdProd, "@nom", prod.Nombre.Trim())
                                DatabaseHelper.AddParam(cmdProd, "@desc", If(prod.Descripcion, "").Trim())
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

                            ' Guardar matriz de talles y stock
                            If prod.TallesStock IsNot Nothing AndAlso prod.TallesStock.Count > 0 Then
                                GuardarMatrizTalles(conn, trans, prod.Id, prod.CodigoBarra, prod.TallesStock)
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
            Dim esAjusteValido = productoId > 0 AndAlso talleId > 0 AndAlso cantidadDelta <> 0
            If Not esAjusteValido Then
                errorMessage = "Parámetros de ajuste inválidos (se requiere producto, talle y una cantidad distinta de cero)."
                Return False
            End If

            Dim colorLimpio = If(String.IsNullOrWhiteSpace(color), "Único", color.Trim())

            Try
                Using conn = DatabaseHelper.GetConnection()
                    conn.Open()
                    Using trans = conn.BeginTransaction()
                        Try
                            Dim stockActual = ObtenerOInsertarStockFila(conn, trans, productoId, talleId, colorLimpio)
                            Dim nuevoStock = Math.Max(0, stockActual + cantidadDelta)

                            ActualizarStockFila(conn, trans, productoId, talleId, colorLimpio, nuevoStock)
                            RegistrarAuditoriaStock(conn, trans, productoId, talleId, colorLimpio, cantidadDelta, stockActual, nuevoStock, motivo, usuarioId)

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
            If id <= 0 Then
                errorMessage = "Identificador de producto inválido."
                Return False
            End If

            Try
                ' Soft delete para mantener integridad de ventas pasadas
                Dim query = "UPDATE `productos` SET `activo` = 0 WHERE `id` = @id;"
                DatabaseHelper.ExecuteNonQuery(query, New Dictionary(Of String, Object) From {{"@id", id}})
                errorMessage = String.Empty
                Return True
            Catch ex As Exception
                errorMessage = ex.Message
                Return False
            End Try
        End Function
#End Region

#Region "Helpers Privados de Soporte"

        Private Sub GuardarMatrizTalles(conn As DbConnection, trans As DbTransaction, prodId As Integer, codBarra As String, talles As List(Of ProductoTalle))
            Dim queryPt = "INSERT INTO `producto_talles` (`producto_id`, `talle_id`, `color`, `stock_actual`, `stock_minimo`, `sku_especifico`) " &
                          "VALUES (@prodId, @talleId, @color, @stock, @stockMin, @sku) " &
                          "ON CONFLICT(`producto_id`, `talle_id`, `color`) DO UPDATE SET `stock_actual` = @stock, `stock_minimo` = @stockMin, `sku_especifico` = @sku;"

            For Each pt In talles
                Dim colorFinal = If(String.IsNullOrWhiteSpace(pt.Color), "Único", pt.Color.Trim())
                Dim skuFinal = If(String.IsNullOrEmpty(pt.SkuEspecifico), $"{codBarra}-{pt.TalleId}", pt.SkuEspecifico)

                Using cmdPt = DatabaseHelper.CreateCommand(conn, queryPt, trans)
                    DatabaseHelper.AddParam(cmdPt, "@prodId", prodId)
                    DatabaseHelper.AddParam(cmdPt, "@talleId", pt.TalleId)
                    DatabaseHelper.AddParam(cmdPt, "@color", colorFinal)
                    DatabaseHelper.AddParam(cmdPt, "@stock", pt.StockActual)
                    DatabaseHelper.AddParam(cmdPt, "@stockMin", pt.StockMinimo)
                    DatabaseHelper.AddParam(cmdPt, "@sku", skuFinal)
                    cmdPt.ExecuteNonQuery()
                End Using
            Next
        End Sub

        Private Function ObtenerOInsertarStockFila(conn As DbConnection, trans As DbTransaction, productoId As Integer, talleId As Integer, color As String) As Integer
            Dim querySelect = "SELECT stock_actual FROM `producto_talles` WHERE `producto_id` = @prodId AND `talle_id` = @talleId AND `color` = @color;"
            Using cmdSelect = DatabaseHelper.CreateCommand(conn, querySelect, trans)
                DatabaseHelper.AddParam(cmdSelect, "@prodId", productoId)
                DatabaseHelper.AddParam(cmdSelect, "@talleId", talleId)
                DatabaseHelper.AddParam(cmdSelect, "@color", color)
                Dim res = cmdSelect.ExecuteScalar()
                If res IsNot Nothing AndAlso Not Convert.IsDBNull(res) Then
                    Return Convert.ToInt32(res)
                End If
            End Using

            Dim queryInsertPt = "INSERT INTO `producto_talles` (`producto_id`, `talle_id`, `color`, `stock_actual`, `stock_minimo`) VALUES (@prodId, @talleId, @color, 0, 2);"
            Using cmdIns = DatabaseHelper.CreateCommand(conn, queryInsertPt, trans)
                DatabaseHelper.AddParam(cmdIns, "@prodId", productoId)
                DatabaseHelper.AddParam(cmdIns, "@talleId", talleId)
                DatabaseHelper.AddParam(cmdIns, "@color", color)
                cmdIns.ExecuteNonQuery()
            End Using

            Return 0
        End Function

        Private Sub ActualizarStockFila(conn As DbConnection, trans As DbTransaction, productoId As Integer, talleId As Integer, color As String, nuevoStock As Integer)
            Dim queryUpdate = "UPDATE `producto_talles` SET `stock_actual` = @nuevoStock WHERE `producto_id` = @prodId AND `talle_id` = @talleId AND `color` = @color;"
            Using cmdUpd = DatabaseHelper.CreateCommand(conn, queryUpdate, trans)
                DatabaseHelper.AddParam(cmdUpd, "@nuevoStock", nuevoStock)
                DatabaseHelper.AddParam(cmdUpd, "@prodId", productoId)
                DatabaseHelper.AddParam(cmdUpd, "@talleId", talleId)
                DatabaseHelper.AddParam(cmdUpd, "@color", color)
                cmdUpd.ExecuteNonQuery()
            End Using
        End Sub

        Private Sub RegistrarAuditoriaStock(conn As DbConnection, trans As DbTransaction, productoId As Integer, talleId As Integer, color As String, cantidadDelta As Integer, stockActual As Integer, nuevoStock As Integer, motivo As String, usuarioId As Integer)
            Dim queryMov = "INSERT INTO `movimientos_stock` (`producto_id`, `talle_id`, `color`, `tipo_movimiento`, `cantidad`, `stock_anterior`, `stock_posterior`, `motivo`, `usuario_id`) " &
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
        End Sub

        Private Shared Function MapCategoria(row As DataRow) As Categoria
            Return New Categoria() With {
                .Id = Convert.ToInt32(row("id")),
                .Nombre = row("nombre").ToString(),
                .Descripcion = If(Convert.IsDBNull(row("descripcion")), "", row("descripcion").ToString()),
                .Activo = Convert.ToBoolean(row("activo"))
            }
        End Function

        Private Shared Function MapTalle(row As DataRow) As Talle
            Return New Talle() With {
                .Id = Convert.ToInt32(row("id")),
                .Nombre = row("nombre").ToString(),
                .Orden = Convert.ToInt32(row("orden"))
            }
        End Function

        Private Function MapProducto(row As DataRow, Optional incluirTalles As Boolean = False) As Producto
            Dim prodId = Convert.ToInt32(row("id"))
            Dim prod = New Producto() With {
                .Id = prodId,
                .CodigoBarra = row("codigo_barra").ToString(),
                .Nombre = row("nombre").ToString(),
                .Descripcion = If(Convert.IsDBNull(row("descripcion")), "", row("descripcion").ToString()),
                .CategoriaId = Convert.ToInt32(row("categoria_id")),
                .CategoriaNombre = row("categoria_nombre").ToString(),
                .PrecioCosto = Convert.ToDecimal(row("precio_costo")),
                .PrecioVenta = Convert.ToDecimal(row("precio_venta")),
                .PorcentajeGanancia = Convert.ToDecimal(row("porcentaje_ganancia")),
                .ImagenRuta = If(Convert.IsDBNull(row("imagen_ruta")), "", row("imagen_ruta").ToString()),
                .Activo = Convert.ToBoolean(row("activo")),
                .TotalStock = Convert.ToInt32(row("total_stock"))
            }

            If incluirTalles Then
                prod.TallesStock = GetStockPorTalles(prodId)
            End If

            Return prod
        End Function

        Private Shared Function MapProductoTalle(row As DataRow) As ProductoTalle
            Return New ProductoTalle() With {
                .Id = Convert.ToInt32(row("id")),
                .ProductoId = Convert.ToInt32(row("producto_id")),
                .TalleId = Convert.ToInt32(row("talle_id")),
                .TalleNombre = row("talle_nombre").ToString(),
                .Color = row("color").ToString(),
                .StockActual = Convert.ToInt32(row("stock_actual")),
                .StockMinimo = Convert.ToInt32(row("stock_minimo")),
                .SkuEspecifico = If(Convert.IsDBNull(row("sku_especifico")), "", row("sku_especifico").ToString())
            }
        End Function

#End Region

    End Class
End Namespace
