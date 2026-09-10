Friend Module Program

    <STAThread()>
    Friend Sub Main(args As String())
        If args IsNot Nothing AndAlso args.Length > 0 AndAlso args(0) = "--init-db" Then
            Dim outMsg As String = ""
            Dim ok = Data.DatabaseHelper.InitializeDatabaseAndTables(outMsg)
            Console.WriteLine("INIT_RESULT:" & ok.ToString() & ":" & outMsg)
            Return
        End If

        If args IsNot Nothing AndAlso args.Length > 0 AndAlso args(0) = "--test" Then
            Dim outMsg As String = ""
            Dim okConn = Data.DatabaseHelper.TestConnection(outMsg)
            Console.WriteLine("TEST_CONN:" & okConn.ToString() & ":" & outMsg)

            Dim authErr As String = ""
            Dim okLogin = Services.AuthService.Login("admin", "admin123", authErr)
            Console.WriteLine("TEST_AUTH:" & okLogin.ToString() & ":" & authErr)

            Dim catSvc As New Services.CatalogService()
            Dim prods = catSvc.GetProductos()
            Console.WriteLine("TEST_PRODUCTS:" & prods.Count.ToString())

            Dim talles = catSvc.GetTalles()
            Console.WriteLine("TEST_TALLES:" & talles.Count.ToString())

            Dim cajaSvc As New Services.CajaService()
            Dim cajaErr As String = ""
            Dim caja = cajaSvc.GetCajaAbierta(Services.AuthService.CurrentUser.Id)
            If caja Is Nothing Then
                caja = cajaSvc.AbrirCaja(Services.AuthService.CurrentUser.Id, 15000D, cajaErr)
                Console.WriteLine("TEST_CAJA_ABRIR:" & (caja IsNot Nothing).ToString() & ":" & cajaErr)
            End If
            Console.WriteLine("TEST_CAJA_ID:" & If(caja IsNot Nothing, caja.Id.ToString(), "0"))

            Dim ventaSvc As New Services.VentaService()
            Dim venta As New Models.Venta() With {
                .NumeroTicket = ventaSvc.GenerarNumeroTicket(),
                .UsuarioId = Services.AuthService.CurrentUser.Id,
                .ClienteId = 1,
                .CajaId = If(caja IsNot Nothing, caja.Id, 1),
                .TipoComprobante = "Ticket",
                .MetodoPago = "Efectivo",
                .Subtotal = 18500D,
                .Total = 18500D,
                .MontoAbonado = 20000D,
                .Vuelto = 1500D,
                .Estado = "Completada"
            }
            venta.Detalles.Add(New Models.DetalleVenta() With {
                .ProductoId = 1,
                .TalleId = 2,
                .Color = "Negro",
                .CodigoBarra = "779001001",
                .DescripcionArticulo = "Remera Oversize Vintage Acid Wash (Talle: S, Color: Negro)",
                .PrecioUnitario = 18500D,
                .CostoUnitario = 9500D,
                .Cantidad = 1,
                .Subtotal = 18500D
            })

            Dim ventaErr As String = ""
            Dim okVenta = ventaSvc.ProcesarVenta(venta, ventaErr)
            Console.WriteLine("TEST_VENTA:" & okVenta.ToString() & ":Ticket=" & venta.NumeroTicket & ":Err=" & ventaErr)
            Return
        End If

        Application.SetHighDpiMode(HighDpiMode.PerMonitorV2)
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Application.Run(New Forms.FrmLogin())
    End Sub

End Module
