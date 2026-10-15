' ARCHIVO: ReporteService.vb
' PROPÓSITO: Cálculo de métricas comerciales, KPIs de ventas y alertas de stock crítico.
' Este servicio alimenta el Dashboard del formulario principal (FrmMain) y la pantalla de reportes analíticos:
' 1. KPIs en Tiempo Real: Suma ventas del día de hoy, cuenta tickets emitidos y discrimina
'    cuánto dinero entró en efectivo y cuánto por canales digitales (tarjetas / transferencias).
' 2. Ranking de Prendas más Vendidas: Agrupa por producto, talle y monto recaudado
'    para que el encargado sepa qué artículos son los más rentables del local.
' 3. Alertas de Reposición: Identifica automáticamente las prendas cuyo stock actual
'    es menor o igual a su stock mínimo para emitir pedidos a proveedores a tiempo.


Imports System.Data
Imports GestionComercial.Data

Namespace Services
    ' DTO para las tarjetas superiores (KPI Cards) del Panel Principal
    Public Class ResumenVentasHoy
        Public Property TotalVendido As Decimal = 0D
        Public Property CantidadTickets As Integer = 0
        Public Property TotalEfectivo As Decimal = 0D
        Public Property TotalDigital As Decimal = 0D
        Public Property ArticulosStockBajo As Integer = 0
    End Class

    ' Criterio para el ranking de productos
    Public Enum CriterioRankingProducto
        MayorSalida
        MayorMargen
    End Enum

    ' DTO para el ranking de artículos estrella y rentabilidad
    Public Class TopProductoVendido
        Public Property ProductoId As Integer
        Public Property Nombre As String = String.Empty
        Public Property Talle As String = String.Empty
        Public Property CantidadVendida As Integer = 0
        Public Property TotalRecaudado As Decimal = 0D
        Public Property CostoTotal As Decimal = 0D
        Public Property MargenGanancia As Decimal = 0D
        Public Property PorcentajeMargen As Decimal = 0D
    End Class

    ' Criterio de ordenación para el reporte de clientes
    Public Enum CriterioOrdenCliente
        MayorFacturacion
        MayorSalida
        MayorMargen
    End Enum

    ' DTO para el reporte analítico de ventas por cliente
    Public Class ReporteVentaCliente
        Public Property ClienteId As Integer
        Public Property ClienteNombre As String = String.Empty
        Public Property DniCuit As String = String.Empty
        Public Property Telefono As String = String.Empty
        Public Property CantidadCompras As Integer = 0
        Public Property TotalUnidades As Integer = 0
        Public Property TotalFacturado As Decimal = 0D
        Public Property CostoTotal As Decimal = 0D
        Public Property MargenGanancia As Decimal = 0D
        Public Property PorcentajeMargen As Decimal = 0D
        Public Property UltimaCompra As Nullable(Of DateTime) = Nothing
    End Class

    ' DTO para alertar prendas por agotarse
    Public Class AlertaStock
        Public Property ProductoId As Integer
        Public Property CodigoBarra As String = String.Empty
        Public Property Nombre As String = String.Empty
        Public Property Talle As String = String.Empty
        Public Property Color As String = "Único"
        Public Property StockActual As Integer = 0
        Public Property StockMinimo As Integer = 0
    End Class

    Public Class ReporteService

        ' Consulta agregada de KPIs para el Dashboard del día de hoy (ventas, tickets y medios de pago).

        Public Function GetResumenHoy() As ResumenVentasHoy
            Dim resumen As New ResumenVentasHoy()
            Try
                Dim query As String = "SELECT " &
                                     "IFNULL(SUM(total), 0) AS total_vendido, " &
                                     "COUNT(*) AS cantidad_tickets, " &
                                     "IFNULL(SUM(CASE WHEN metodo_pago = 'Efectivo' THEN total ELSE 0 END), 0) AS total_efectivo, " &
                                     "IFNULL(SUM(CASE WHEN metodo_pago != 'Efectivo' THEN total ELSE 0 END), 0) AS total_digital " &
                                     "FROM `ventas` WHERE DATE(`fecha`) = CURDATE() AND `estado` = 'Completada';"

                Dim dt As DataTable = DatabaseHelper.ExecuteQuery(query)
                If dt.Rows.Count > 0 Then
                    Dim row = dt.Rows(0)
                    resumen.TotalVendido = Convert.ToDecimal(row("total_vendido"))
                    resumen.CantidadTickets = Convert.ToInt32(row("cantidad_tickets"))
                    resumen.TotalEfectivo = Convert.ToDecimal(row("total_efectivo"))
                    resumen.TotalDigital = Convert.ToDecimal(row("total_digital"))
                End If

                ' Cantidad de artículos con stock bajo
                Dim queryStock As String = "SELECT COUNT(*) FROM `producto_talles` pt " &
                                          "INNER JOIN `productos` p ON pt.producto_id = p.id " &
                                          "WHERE p.activo = 1 AND pt.stock_actual <= pt.stock_minimo;"
                resumen.ArticulosStockBajo = Convert.ToInt32(DatabaseHelper.ExecuteScalar(queryStock))

            Catch ex As Exception
                ' Retorna ceros
            End Try
            Return resumen
        End Function

        Public Function GetTopProductosVendidos(limite As Integer, fechaDesde As DateTime, fechaHasta As DateTime, Optional criterio As CriterioRankingProducto = CriterioRankingProducto.MayorSalida) As List(Of TopProductoVendido)
            Dim list As New List(Of TopProductoVendido)()
            Try
                Dim orderByClause As String = If(criterio = CriterioRankingProducto.MayorMargen,
                                                 "ORDER BY total_margen DESC, total_cantidad DESC ",
                                                 "ORDER BY total_cantidad DESC, total_monto DESC ")

                Dim limitClause As String = If(limite > 0, "LIMIT @limite;", ";")

                Dim query As String = "SELECT dv.producto_id, dv.descripcion_articulo, t.nombre AS talle_nombre, " &
                                     "SUM(dv.cantidad) AS total_cantidad, SUM(dv.subtotal) AS total_monto, " &
                                     "SUM(COALESCE(NULLIF(dv.costo_unitario, 0), p.precio_costo, 0) * dv.cantidad) AS total_costo, " &
                                     "SUM(dv.subtotal - (COALESCE(NULLIF(dv.costo_unitario, 0), p.precio_costo, 0) * dv.cantidad)) AS total_margen " &
                                     "FROM `detalle_ventas` dv " &
                                     "INNER JOIN `ventas` v ON dv.venta_id = v.id " &
                                     "INNER JOIN `productos` p ON dv.producto_id = p.id " &
                                     "LEFT JOIN `talles` t ON dv.talle_id = t.id " &
                                     "WHERE DATE(v.fecha) >= @desde AND DATE(v.fecha) <= @hasta AND v.estado = 'Completada' " &
                                     "GROUP BY dv.producto_id, dv.descripcion_articulo, dv.talle_id, t.nombre " &
                                     orderByClause & limitClause

                Dim params As New Dictionary(Of String, Object) From {
                    {"@desde", fechaDesde.ToString("yyyy-MM-dd")},
                    {"@hasta", fechaHasta.ToString("yyyy-MM-dd")}
                }
                If limite > 0 Then
                    params.Add("@limite", limite)
                End If

                Dim dt As DataTable = DatabaseHelper.ExecuteQuery(query, params)
                For Each row As DataRow In dt.Rows
                    Dim recaudado As Decimal = Convert.ToDecimal(row("total_monto"))
                    Dim costo As Decimal = Convert.ToDecimal(row("total_costo"))
                    Dim margen As Decimal = Convert.ToDecimal(row("total_margen"))
                    Dim pctMargen As Decimal = If(recaudado > 0D, Math.Round((margen / recaudado) * 100D, 1), 0D)

                    list.Add(New TopProductoVendido() With {
                        .ProductoId = Convert.ToInt32(row("producto_id")),
                        .Nombre = row("descripcion_articulo").ToString(),
                        .Talle = If(IsDBNull(row("talle_nombre")), "-", row("talle_nombre").ToString()),
                        .CantidadVendida = Convert.ToInt32(row("total_cantidad")),
                        .TotalRecaudado = recaudado,
                        .CostoTotal = costo,
                        .MargenGanancia = margen,
                        .PorcentajeMargen = pctMargen
                    })
                Next
            Catch ex As Exception
                ' Silencioso
            End Try
            Return list
        End Function

        Public Function GetReporteVentasPorCliente(fechaDesde As DateTime, fechaHasta As DateTime, Optional clienteId As Integer = 0, Optional criterio As CriterioOrdenCliente = CriterioOrdenCliente.MayorFacturacion) As List(Of ReporteVentaCliente)
            Dim list As New List(Of ReporteVentaCliente)()
            Try
                Dim query As String = "SELECT " &
                                     "c.id AS cliente_id, " &
                                     "(c.apellido || ', ' || c.nombre) AS cliente_nombre, " &
                                     "c.dni_cuit, " &
                                     "IFNULL(c.telefono, '') AS telefono, " &
                                     "COUNT(sub.venta_id) AS cantidad_compras, " &
                                     "IFNULL(SUM(sub.unidades_venta), 0) AS total_unidades, " &
                                     "IFNULL(SUM(sub.total), 0) AS total_facturado, " &
                                     "IFNULL(SUM(sub.costo_venta), 0) AS total_costo, " &
                                     "(IFNULL(SUM(sub.total), 0) - IFNULL(SUM(sub.costo_venta), 0)) AS margen_ganancia, " &
                                     "MAX(sub.fecha) AS ultima_compra " &
                                     "FROM `clientes` c " &
                                     "INNER JOIN ( " &
                                     "    SELECT " &
                                     "        v.cliente_id, " &
                                     "        v.id AS venta_id, " &
                                     "        v.fecha, " &
                                     "        v.total, " &
                                     "        IFNULL(dv.unidades_venta, 0) AS unidades_venta, " &
                                     "        IFNULL(dv.costo_venta, 0) AS costo_venta " &
                                     "    FROM `ventas` v " &
                                     "    LEFT JOIN ( " &
                                     "        SELECT " &
                                     "            dv_inner.venta_id, " &
                                     "            SUM(dv_inner.cantidad) AS unidades_venta, " &
                                     "            SUM(COALESCE(NULLIF(dv_inner.costo_unitario, 0), p.precio_costo, 0) * dv_inner.cantidad) AS costo_venta " &
                                     "        FROM `detalle_ventas` dv_inner " &
                                     "        INNER JOIN `productos` p ON dv_inner.producto_id = p.id " &
                                     "        GROUP BY dv_inner.venta_id " &
                                     "    ) dv ON v.id = dv.venta_id " &
                                     "    WHERE DATE(v.fecha) >= @desde AND DATE(v.fecha) <= @hasta AND v.estado = 'Completada' " &
                                     ") sub ON c.id = sub.cliente_id " &
                                     "WHERE 1=1 "

                Dim params As New Dictionary(Of String, Object) From {
                    {"@desde", fechaDesde.ToString("yyyy-MM-dd")},
                    {"@hasta", fechaHasta.ToString("yyyy-MM-dd")}
                }

                If clienteId > 0 Then
                    query &= "AND c.id = @cliId "
                    params.Add("@cliId", clienteId)
                End If

                query &= "GROUP BY c.id, c.apellido, c.nombre, c.dni_cuit, c.telefono "

                Select Case criterio
                    Case CriterioOrdenCliente.MayorSalida
                        query &= "ORDER BY total_unidades DESC, total_facturado DESC;"
                    Case CriterioOrdenCliente.MayorMargen
                        query &= "ORDER BY margen_ganancia DESC, total_facturado DESC;"
                    Case Else
                        query &= "ORDER BY total_facturado DESC, cantidad_compras DESC;"
                End Select

                Dim dt As DataTable = DatabaseHelper.ExecuteQuery(query, params)
                For Each row As DataRow In dt.Rows
                    Dim facturado As Decimal = Convert.ToDecimal(row("total_facturado"))
                    Dim costo As Decimal = Convert.ToDecimal(row("total_costo"))
                    Dim margen As Decimal = Convert.ToDecimal(row("margen_ganancia"))
                    Dim pctMargen As Decimal = If(facturado > 0D, Math.Round((margen / facturado) * 100D, 1), 0D)

                    Dim ultimaFecha As Nullable(Of DateTime) = Nothing
                    If Not IsDBNull(row("ultima_compra")) Then
                        Dim parsed As DateTime
                        If DateTime.TryParse(row("ultima_compra").ToString(), parsed) Then
                            ultimaFecha = parsed
                        End If
                    End If

                    list.Add(New ReporteVentaCliente() With {
                        .ClienteId = Convert.ToInt32(row("cliente_id")),
                        .ClienteNombre = row("cliente_nombre").ToString(),
                        .DniCuit = row("dni_cuit").ToString(),
                        .Telefono = row("telefono").ToString(),
                        .CantidadCompras = Convert.ToInt32(row("cantidad_compras")),
                        .TotalUnidades = Convert.ToInt32(row("total_unidades")),
                        .TotalFacturado = facturado,
                        .CostoTotal = costo,
                        .MargenGanancia = margen,
                        .PorcentajeMargen = pctMargen,
                        .UltimaCompra = ultimaFecha
                    })
                Next
            Catch ex As Exception
                ' Silencioso
            End Try
            Return list
        End Function

        Public Function GetAlertasStockBajo() As List(Of AlertaStock)
            Dim list As New List(Of AlertaStock)()
            Try
                Dim query As String = "SELECT pt.producto_id, p.codigo_barra, p.nombre, t.nombre AS talle_nombre, pt.color, pt.stock_actual, pt.stock_minimo " &
                                     "FROM `producto_talles` pt " &
                                     "INNER JOIN `productos` p ON pt.producto_id = p.id " &
                                     "INNER JOIN `talles` t ON pt.talle_id = t.id " &
                                     "WHERE p.activo = 1 AND pt.stock_actual <= pt.stock_minimo " &
                                     "ORDER BY pt.stock_actual ASC, p.nombre ASC;"

                Dim dt As DataTable = DatabaseHelper.ExecuteQuery(query)
                For Each row As DataRow In dt.Rows
                    list.Add(New AlertaStock() With {
                        .ProductoId = Convert.ToInt32(row("producto_id")),
                        .CodigoBarra = row("codigo_barra").ToString(),
                        .Nombre = row("nombre").ToString(),
                        .Talle = row("talle_nombre").ToString(),
                        .Color = row("color").ToString(),
                        .StockActual = Convert.ToInt32(row("stock_actual")),
                        .StockMinimo = Convert.ToInt32(row("stock_minimo"))
                    })
                Next
            Catch ex As Exception
                ' Silencioso
            End Try
            Return list
        End Function

    End Class
End Namespace
