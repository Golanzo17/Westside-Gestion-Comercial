Imports System.Data
Imports System.IO
Imports Microsoft.Data.Sqlite
Imports Xunit
Imports GestionComercial.Config
Imports GestionComercial.Data
Imports GestionComercial.Models
Imports GestionComercial.Services

Namespace GestionComercial.Tests

    Public Class SystemIntegrationTests
        Implements IDisposable

        Private ReadOnly _tempDbPath As String

        Public Sub New()
            _tempDbPath = Path.Combine(Path.GetTempPath(), $"gestion_comercial_xunit_{Guid.NewGuid():N}.db")
            AppConfig.Settings.CustomDatabasePath = _tempDbPath
            Dim msg As String = ""
            Dim ok = DatabaseHelper.InitializeDatabaseAndTables(msg)
            Assert.True(ok, "Falló la inicialización de la base de datos de prueba: " & msg)
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            AppConfig.Settings.CustomDatabasePath = String.Empty
            SqliteConnection.ClearAllPools()
            Try
                If File.Exists(_tempDbPath) Then File.Delete(_tempDbPath)
                If File.Exists(_tempDbPath & "-wal") Then File.Delete(_tempDbPath & "-wal")
                If File.Exists(_tempDbPath & "-shm") Then File.Delete(_tempDbPath & "-shm")
            Catch
            End Try
        End Sub

        <Fact>
        Public Sub PasswordHashing_PBKDF2_ShouldAuthenticateAdminAndManager()
            Dim errMsg As String = ""
            Dim okAdmin = AuthService.Login("admin", "admin123", errMsg)
            Assert.True(okAdmin, errMsg)
            Assert.NotNull(AuthService.CurrentUser)
            Assert.Equal("Administrador", AuthService.CurrentUser.Rol)

            Dim supMgr As Usuario = Nothing
            Dim okMgr = AuthService.AutenticarSupervisorPorRoles("gerente123", {"Gerente"}, "Test xUnit", supMgr, errMsg)
            Assert.True(okMgr, errMsg)
            Assert.NotNull(supMgr)
            Assert.Equal("Gerente", supMgr.Rol)
        End Sub

        <Fact>
        Public Sub BruteForceProtection_ShouldLockAccountAfterFiveFailedAttempts()
            Dim errMsg As String = ""
            Dim locked As Boolean = False
            For i As Integer = 1 To 6
                AuthService.Login("usuario_inexistente_test", "clave_erronea", errMsg)
                If errMsg.Contains("bloqueada") OrElse errMsg.Contains("Demasiados intentos") Then
                    locked = True
                    Exit For
                End If
            Next
            Assert.True(locked, "Se esperaba bloqueo por fuerza bruta tras 5 intentos fallidos.")
        End Sub

        <Fact>
        Public Sub VentaAndAnulacion_ShouldUpdateStockKardexAndCajaCorrectly()
            Dim errMsg As String = ""
            Assert.True(AuthService.Login("vendedor", "1234", errMsg))
            Dim vendedorId = AuthService.CurrentUser.Id

            Dim cajaSvc As New CajaService()
            Dim caja = cajaSvc.AbrirCaja(vendedorId, 10000D, errMsg)
            Assert.NotNull(caja)

            ' Fijamos stock inicial = 10 para Producto 1, Talle 2, Color Negro
            DatabaseHelper.ExecuteNonQuery("UPDATE `producto_talles` SET `stock_actual` = 10 WHERE `producto_id` = 1 AND `talle_id` = 2 AND `color` = 'Negro';")

            Dim ventaSvc As New VentaService()
            Dim venta As New Venta() With {
                .NumeroTicket = ventaSvc.GenerarNumeroTicket(),
                .UsuarioId = vendedorId,
                .ClienteId = 1,
                .CajaId = caja.Id,
                .TipoComprobante = "Ticket",
                .MetodoPago = "Múltiple",
                .Subtotal = 18500D,
                .Total = 18500D,
                .MontoAbonado = 10000D,
                .MontoEfectivo = 10000D,
                .Vuelto = 0D,
                .Estado = "Completada"
            }
            venta.Detalles.Add(New DetalleVenta() With {
                .ProductoId = 1,
                .TalleId = 2,
                .Color = "Negro",
                .CodigoBarra = "779001001",
                .DescripcionArticulo = "Remera Oversize Vintage Acid Wash",
                .PrecioUnitario = 18500D,
                .CostoUnitario = 9500D,
                .Cantidad = 2,
                .Subtotal = 37000D
            })
            venta.Subtotal = 37000D
            venta.Total = 37000D
            venta.MontoAbonado = 15000D
            venta.MontoEfectivo = 15000D

            Assert.True(ventaSvc.ProcesarVenta(venta, errMsg), errMsg)

            ' Verificar que el stock bajó de 10 a 8
            Dim stockPostVenta = Convert.ToInt32(DatabaseHelper.ExecuteScalar("SELECT stock_actual FROM producto_talles WHERE producto_id = 1 AND talle_id = 2 AND color = 'Negro';"))
            Assert.Equal(8, stockPostVenta)

            ' Verificar que Caja registró $15.000 en efectivo y $22.000 en digital (pago mixto)
            Dim cajaAct = cajaSvc.GetCajaAbierta(vendedorId)
            Assert.NotNull(cajaAct)
            Assert.Equal(15000D, cajaAct.TotalVentasEfectivo)
            Assert.Equal(22000D, cajaAct.TotalVentasDigital)

            ' Verificar que GetVentasPorCaja y GetVentaPorId funcionan sin errores de columna SQL
            Dim ventasCaja = cajaSvc.GetVentasPorCaja(caja.Id)
            Assert.Single(ventasCaja)
            Assert.Equal(venta.NumeroTicket, ventasCaja(0).NumeroTicket)

            Dim ventaRecuperada = ventaSvc.GetVentaPorId(venta.Id)
            Assert.NotNull(ventaRecuperada)
            Assert.Equal(venta.NumeroTicket, ventaRecuperada.NumeroTicket)
            ventaRecuperada.Detalles = ventaSvc.GetDetalleVenta(venta.Id)
            Dim lineasPdf = Forms.FrmReportes.ConstruirLineasTicket(ventaRecuperada, New ConfiguracionService().GetConfiguracion())
            Assert.NotEmpty(lineasPdf)

            ' Anular la venta y verificar reversión de stock, Kardex exacto y reversión de caja
            Assert.True(ventaSvc.AnularVenta(venta.Id, 1, "Devolución por cambio de talle", errMsg), errMsg)

            Dim stockPostAnulacion = Convert.ToInt32(DatabaseHelper.ExecuteScalar("SELECT stock_actual FROM producto_talles WHERE producto_id = 1 AND talle_id = 2 AND color = 'Negro';"))
            Assert.Equal(10, stockPostAnulacion)

            Dim dtKardex As DataTable = DatabaseHelper.ExecuteQuery("SELECT stock_anterior, stock_posterior, cantidad FROM movimientos_stock WHERE tipo_movimiento = 'Anulacion_Venta' ORDER BY id DESC LIMIT 1;")
            Assert.Equal(1, dtKardex.Rows.Count)
            Assert.Equal(8, Convert.ToInt32(dtKardex.Rows(0)("stock_anterior")))
            Assert.Equal(10, Convert.ToInt32(dtKardex.Rows(0)("stock_posterior")))
            Assert.Equal(2, Convert.ToInt32(dtKardex.Rows(0)("cantidad")))

            Dim cajaFinal = cajaSvc.GetCajaAbierta(vendedorId)
            Assert.Equal(0D, cajaFinal.TotalVentasEfectivo)
            Assert.Equal(0D, cajaFinal.TotalVentasDigital)
        End Sub

    End Class

End Namespace

