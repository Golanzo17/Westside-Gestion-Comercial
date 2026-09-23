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

    ' DTO para el ranking de artículos estrella
    Public Class TopProductoVendido
        Public Property ProductoId As Integer
        Public Property Nombre As String = String.Empty
        Public Property Talle As String = String.Empty
        Public Property CantidadVendida As Integer = 0
        Public Property TotalRecaudado As Decimal = 0D
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

        Public Function GetTopProductosVendidos(limite As Integer, fechaDesde As DateTime, fechaHasta As DateTime) As List(Of TopProductoVendido)
            Dim list As New List(Of TopProductoVendido)()
            Try
                Dim query As String = "SELECT dv.producto_id, dv.descripcion_articulo, t.nombre AS talle_nombre, " &
                                     "SUM(dv.cantidad) AS total_cantidad, SUM(dv.subtotal) AS total_monto " &
                                     "FROM `detalle_ventas` dv " &
                                     "INNER JOIN `ventas` v ON dv.venta_id = v.id " &
                                     "LEFT JOIN `talles` t ON dv.talle_id = t.id " &
                                     "WHERE DATE(v.fecha) >= @desde AND DATE(v.fecha) <= @hasta AND v.estado = 'Completada' " &
                                     "GROUP BY dv.producto_id, dv.descripcion_articulo, dv.talle_id, t.nombre " &
                                     "ORDER BY total_cantidad DESC LIMIT @limite;"

                Dim params As New Dictionary(Of String, Object) From {
                    {"@desde", fechaDesde.ToString("yyyy-MM-dd")},
                    {"@hasta", fechaHasta.ToString("yyyy-MM-dd")},
                    {"@limite", limite}
                }

                Dim dt As DataTable = DatabaseHelper.ExecuteQuery(query, params)
                For Each row As DataRow In dt.Rows
                    list.Add(New TopProductoVendido() With {
                        .ProductoId = Convert.ToInt32(row("producto_id")),
                        .Nombre = row("descripcion_articulo").ToString(),
                        .Talle = If(IsDBNull(row("talle_nombre")), "-", row("talle_nombre").ToString()),
                        .CantidadVendida = Convert.ToInt32(row("total_cantidad")),
                        .TotalRecaudado = Convert.ToDecimal(row("total_monto"))
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
