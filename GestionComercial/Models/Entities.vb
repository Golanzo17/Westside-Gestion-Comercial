' ARCHIVO: Entities.vb
' PROPÓSITO: Modelos de dominio y DTOs (Data Transfer Objects) del sistema.
' En este archivo definimos las clases que representan la información de nuestro local
' en memoria. Cada clase se corresponde con una tabla de la base de datos:
' - Usuario: Empleados, roles y permisos de acceso.
' - Producto y ProductoTalle: Catálogo y la matriz de stock por talle y color.
' - Venta y DetalleVenta: Encabezado y renglones de los tickets de venta.
' - Caja y MovimientoCaja: Control de turnos, efectivo inicial, ingresos y arqueo.
' - MovimientoStock: Registro histórico (Kardex) para auditoría de inventario.
' - Cliente: Directorio comercial de compradores.
' - ConfiguracionComercio: Datos institucionales impresos en el ticket.

Namespace Models

    ' Modela los empleados del local con su rol asignado (Administrador, Gerente o Vendedor) y datos de contacto.
    Public Class Usuario
        Public Property Id As Integer
        Public Property Username As String = String.Empty
        Public Property Dni As String = String.Empty
        Public Property PasswordHash As String = String.Empty
        Public Property Nombre As String = String.Empty
        Public Property Apellido As String = String.Empty
        Public Property Rol As String = "Vendedor" ' Administrador, Gerente, Vendedor
        Public Property Telefono As String = String.Empty
        Public Property Email As String = String.Empty
        Public Property Direccion As String = String.Empty
        Public Property Ciudad As String = String.Empty
        Public Property Notas As String = String.Empty
        Public Property FechaNacimiento As Nullable(Of DateTime)
        Public Property Activo As Boolean = True
        Public Property UltimoLogin As Nullable(Of DateTime)
        Public Property CreatedAt As DateTime = DateTime.Now

        ' Propiedad calculada para mostrar "Apellido, Nombre" en grillas y formularios
        Public ReadOnly Property NombreCompleto As String
            Get
                Dim s = $"{Apellido}, {Nombre}".Trim(" "c, ","c)
                Return If(String.IsNullOrWhiteSpace(s), Username, s)
            End Get
        End Property
    End Class

    ' Clasifica las prendas del local (Remeras, Pantalones, Camperas, etc.).
    Public Class Categoria
        Public Property Id As Integer
        Public Property Nombre As String = String.Empty
        Public Property Descripcion As String = String.Empty
        Public Property Activo As Boolean = True

        Public Overrides Function ToString() As String
            Return Nombre
        End Function
    End Class

    ' Talles universales de indumentaria (XS, S, M, L, XL, XXL, 38, 40, etc.).
    Public Class Talle
        Public Property Id As Integer
        Public Property Nombre As String = String.Empty
        Public Property Orden As Integer = 0

        Public Overrides Function ToString() As String
            Return Nombre
        End Function
    End Class

    ' Representa el artículo genérico del catálogo con sus precios y categoría.
    ' Un producto contiene una lista de ProductoTalle para discriminar cantidades por cada variante de talle y color.
    Public Class Producto
        Public Property Id As Integer
        Public Property CodigoBarra As String = String.Empty
        Public Property Nombre As String = String.Empty
        Public Property Descripcion As String = String.Empty
        Public Property CategoriaId As Integer
        Public Property CategoriaNombre As String = String.Empty
        Public Property PrecioCosto As Decimal = 0D
        Public Property PrecioVenta As Decimal = 0D
        Public Property PorcentajeGanancia As Decimal = 50D
        Public Property ImagenRuta As String = String.Empty
        Public Property Activo As Boolean = True
        Public Property TotalStock As Integer = 0
        Public Property TallesStock As List(Of ProductoTalle) = New List(Of ProductoTalle)()
    End Class

    ' Matriz de stock por prenda: define las unidades disponibles por talle y color, con su umbral mínimo de alerta.
    Public Class ProductoTalle
        Public Property Id As Integer
        Public Property ProductoId As Integer
        Public Property TalleId As Integer
        Public Property TalleNombre As String = String.Empty
        Public Property Color As String = "Único"
        Public Property StockActual As Integer = 0
        Public Property StockMinimo As Integer = 2
        Public Property SkuEspecifico As String = String.Empty

        Public Overrides Function ToString() As String
            Return $"{TalleNombre} - {Color} (Stock: {StockActual})"
        End Function
    End Class

    ' Clientes registrados en el sistema para asignación de tickets y cuenta corriente.
    Public Class Cliente
        Public Property Id As Integer
        Public Property DniCuit As String = String.Empty
        Public Property Nombre As String = String.Empty
        Public Property Apellido As String = String.Empty
        Public Property Telefono As String = String.Empty
        Public Property Email As String = String.Empty
        Public Property Direccion As String = String.Empty
        Public Property Ciudad As String = String.Empty
        Public Property Notas As String = String.Empty
        Public Property FechaNacimiento As Nullable(Of DateTime)
        Public Property Activo As Boolean = True

        Public ReadOnly Property NombreCompleto As String
            Get
                Return $"{Apellido}, {Nombre}".Trim(" "c, ","c)
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return $"{NombreCompleto} ({DniCuit})"
        End Function
    End Class

    ' Control de turnos y arqueo diario de caja:
    ' Registra monto inicial, cobros por medio de pago, movimientos manuales y cálculo de diferencia de cierre.
    Public Class Caja
        Public Property Id As Integer
        Public Property UsuarioId As Integer
        Public Property UsuarioNombre As String = String.Empty
        Public Property FechaApertura As DateTime = DateTime.Now
        Public Property FechaCierre As Nullable(Of DateTime)
        Public Property MontoInicial As Decimal = 0D
        Public Property TotalVentasEfectivo As Decimal = 0D
        Public Property TotalVentasDigital As Decimal = 0D
        Public Property TotalIngresos As Decimal = 0D
        Public Property TotalEgresos As Decimal = 0D
        Public Property MontoEsperado As Decimal = 0D
        Public Property MontoReal As Nullable(Of Decimal)
        Public Property Diferencia As Nullable(Of Decimal)
        Public Property Estado As String = "Abierta" ' Abierta, Cerrada
        Public Property Observaciones As String = String.Empty
    End Class

    ' Movimiento manual de caja (ingresos o egresos extraordinarios como retiros o gastos operativos).
    Public Class MovimientoCaja
        Public Property Id As Integer
        Public Property CajaId As Integer
        Public Property UsuarioId As Integer
        Public Property UsuarioNombre As String = String.Empty
        Public Property Fecha As DateTime = DateTime.Now
        Public Property Tipo As String = "Ingreso" ' Ingreso, Egreso
        Public Property Concepto As String = String.Empty
        Public Property Monto As Decimal = 0D
        Public Property Referencia As String = String.Empty
    End Class

    ' Encabezado de la venta (ticket o comprobante emitido en el POS).
    Public Class Venta
        Public Property Id As Integer
        Public Property NumeroTicket As String = String.Empty
        Public Property Fecha As DateTime = DateTime.Now
        Public Property UsuarioId As Integer
        Public Property UsuarioNombre As String = String.Empty
        Public Property ClienteId As Integer = 1
        Public Property ClienteNombre As String = "Consumidor Final"
        Public Property CajaId As Integer
        Public Property TipoComprobante As String = "Ticket" ' Ticket, Factura B, Factura A, Presupuesto
        Public Property MetodoPago As String = "Efectivo" ' Efectivo, Tarjeta Débito, Tarjeta Crédito, Transferencia, Múltiple
        Public Property Subtotal As Decimal = 0D
        Public Property DescuentoPorcentaje As Decimal = 0D
        Public Property DescuentoMonto As Decimal = 0D
        Public Property RecargoMonto As Decimal = 0D
        Public Property Total As Decimal = 0D
        Public Property MontoAbonado As Decimal = 0D
        Public Property Vuelto As Decimal = 0D
        Public Property Estado As String = "Completada" ' Completada, Anulada
        Public Property Observaciones As String = String.Empty
        Public Property Detalles As List(Of DetalleVenta) = New List(Of DetalleVenta)()
    End Class

    ' Renglón de detalle de una venta con los datos puntuales de la prenda al momento de la transacción.
    Public Class DetalleVenta
        Public Property Id As Integer
        Public Property VentaId As Integer
        Public Property ProductoId As Integer
        Public Property TalleId As Integer
        Public Property TalleNombre As String = String.Empty
        Public Property Color As String = "Único"
        Public Property CodigoBarra As String = String.Empty
        Public Property DescripcionArticulo As String = String.Empty
        Public Property PrecioUnitario As Decimal = 0D
        Public Property CostoUnitario As Decimal = 0D
        Public Property Cantidad As Integer = 1
        Public Property Subtotal As Decimal = 0D
    End Class

    ' Registro de auditoría de inventario (Kardex) para trazabilidad de variaciones de stock.
    Public Class MovimientoStock
        Public Property Id As Integer
        Public Property ProductoId As Integer
        Public Property ProductoNombre As String = String.Empty
        Public Property TalleId As Integer
        Public Property TalleNombre As String = String.Empty
        Public Property Color As String = "Único"
        Public Property TipoMovimiento As String = "Venta" ' Venta, Ingreso_Compra, Ajuste_Manual, Devolucion, Anulacion_Venta
        Public Property Cantidad As Integer = 0
        Public Property StockAnterior As Integer = 0
        Public Property StockPosterior As Integer = 0
        Public Property Motivo As String = String.Empty
        Public Property UsuarioId As Integer
        Public Property UsuarioNombre As String = String.Empty
        Public Property Fecha As DateTime = DateTime.Now
    End Class

    ' Parámetros de identidad del local comercial para la cabecera y pie de tickets.
    Public Class ConfiguracionComercio
        Public Property Id As Integer = 1
        Public Property NombreComercio As String = "Mi Local de Ropa"
        Public Property Cuit As String = ""
        Public Property Direccion As String = ""
        Public Property Telefono As String = ""
        Public Property Email As String = ""
        Public Property CondicionIva As String = "Responsable Inscripto"
        Public Property MensajeTicket As String = "¡Gracias por su compra! Cambios dentro de los 30 días con ticket."
        Public Property MonedaSimbolo As String = "$"
    End Class

End Namespace
