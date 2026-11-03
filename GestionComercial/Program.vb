' ARCHIVO: Program.vb
' PROPÓSITO: Punto de entrada principal (Entry Point) de nuestro sistema.
' Aquí arranca la aplicación. Diseñamos este módulo para que cumpla dos funciones:
' 1. Si se ejecuta normalmente, configura la compatibilidad visual y abre el Login.
' 2. Si recibe argumentos de consola (--init-db o --test), ejecuta pruebas automatizadas
'    en una base de datos aislada temporal sin levantar la interfaz gráfica ni alterar la base productiva.

Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Data.Sqlite

Friend Module Program

    <STAThread()>
    Friend Sub Main(args As String())
        ' Inicialización de base de datos y datos semilla desde consola
        If args IsNot Nothing AndAlso args.Length > 0 AndAlso args(0) = "--init-db" Then
            Dim outMsg As String = ""
            Dim ok = Data.DatabaseHelper.InitializeDatabaseAndTables(outMsg)
            Console.WriteLine("INIT_RESULT:" & ok.ToString() & ":" & outMsg)
            Return
        End If

        ' Pruebas automatizadas de diagnóstico en base de datos temporal aislada:
        ' Conexión a BD, login con PBKDF2, mitigación de fuerza bruta,
        ' lectura de catálogo, apertura de caja, procesamiento/anulación de venta y auditoría Kardex.
        If args IsNot Nothing AndAlso args.Length > 0 AndAlso args(0) = "--test" Then
            Dim tempDbPath As String = Path.Combine(Path.GetTempPath(), $"gestion_comercial_test_{Guid.NewGuid():N}.db")
            Config.AppConfig.Settings.CustomDatabasePath = tempDbPath
            Try
                Dim initMsg As String = ""
                Dim okInit = Data.DatabaseHelper.InitializeDatabaseAndTables(initMsg)
                Console.WriteLine("TEST_INIT_ISOLATED_DB:" & okInit.ToString() & ":" & initMsg)

                ' 1. Probamos conectividad básica y PRAGMAs
                Dim outMsg As String = ""
                Dim okConn = Data.DatabaseHelper.TestConnection(outMsg)
                Console.WriteLine("TEST_CONN:" & okConn.ToString() & ":" & outMsg)

                ' 2. Probamos autenticación con el usuario administrador y gerente
                Dim authErr As String = ""
                Dim okLogin = Services.AuthService.Login("admin", "admin123", authErr)
                Console.WriteLine("TEST_AUTH:" & okLogin.ToString() & ":" & authErr)

                Dim mgrErr As String = ""
                Dim supMgr As Models.Usuario = Nothing
                Dim okMgr = Services.AuthService.AutenticarSupervisorPorRoles("gerente123", {"Gerente"}, "Test diagnóstico", supMgr, mgrErr)
                Console.WriteLine("TEST_AUTH_GERENTE:" & okMgr.ToString() & ":" & mgrErr)

                ' 3. Verificamos que la contraseña almacenada utilice hash PBKDF2
                Dim userRow = Data.DatabaseHelper.ExecuteQuery("SELECT password_hash FROM usuarios WHERE username = 'admin' LIMIT 1;")
                Dim currentHash As String = If(userRow.Rows.Count > 0, userRow.Rows(0)("password_hash").ToString(), "")
                Dim isPbkdf2 As Boolean = currentHash.StartsWith("PBKDF2$SHA256$")
                Console.WriteLine("TEST_PBKDF2_AUTH:" & isPbkdf2.ToString() & ":Prefix=" & If(isPbkdf2, "PBKDF2_OK", currentHash.Substring(0, Math.Min(10, currentHash.Length))))

                ' 4. Probamos el bloqueo temporal por fuerza bruta tras intentos fallidos repetidos
                Dim fakeErr As String = ""
                Dim lockoutTriggered As Boolean = False
                For i As Integer = 1 To 6
                    Services.AuthService.Login("test_brute_user", "wrong_password", fakeErr)
                    If fakeErr.Contains("bloqueada") OrElse fakeErr.Contains("bloqueo") OrElse fakeErr.Contains("Demasiados intentos") Then
                        lockoutTriggered = True
                        Exit For
                    End If
                Next
                Console.WriteLine("TEST_BRUTE_FORCE_LOCKOUT:" & lockoutTriggered.ToString() & ":" & fakeErr)

                ' 5. Verificamos que el sistema encuentre los scripts SQL sin rutas fijas
                Dim resolvedScript = Data.DatabaseHelper.ResolveDatabaseScriptPath("schema_sqlite.sql")
                Dim scriptFound = Not String.IsNullOrEmpty(resolvedScript) AndAlso File.Exists(resolvedScript)
                Console.WriteLine("TEST_SCRIPT_PATH_RESOLVED:" & scriptFound.ToString() & ":" & Path.GetFileName(resolvedScript))

                ' 6. Probamos recuperar datos de la empresa
                Dim cfgSvc As New Services.ConfiguracionService()
                Dim cfg = cfgSvc.GetConfiguracion()
                Dim okCfg = (cfg IsNot Nothing AndAlso Not String.IsNullOrEmpty(cfg.NombreComercio))
                Console.WriteLine("TEST_GUARDAR_CONFIG:" & okCfg.ToString() & ":Err=")

                ' 7. Verificamos catálogo de prendas y talles cargados
                Dim catSvc As New Services.CatalogService()
                Dim prods = catSvc.GetProductos()
                Console.WriteLine("TEST_PRODUCTS:" & prods.Count.ToString())

                Dim talles = catSvc.GetTalles()
                Console.WriteLine("TEST_TALLES:" & talles.Count.ToString())

                ' 8. Verificamos apertura o recuperación de caja del día
                Dim cajaSvc As New Services.CajaService()
                Dim cajaErr As String = ""
                Dim currentUserId As Integer = If(Services.AuthService.CurrentUser IsNot Nothing, Services.AuthService.CurrentUser.Id, 1)
                Dim caja = cajaSvc.GetCajaAbierta(currentUserId)
                If caja Is Nothing Then
                    caja = cajaSvc.AbrirCaja(currentUserId, 15000D, cajaErr)
                    Console.WriteLine("TEST_CAJA_ABRIR:" & (caja IsNot Nothing).ToString() & ":" & cajaErr)
                End If
                Console.WriteLine("TEST_CAJA_ID:" & If(caja IsNot Nothing, caja.Id.ToString(), "0"))

                ' 9. Probamos el circuito completo de cobro de un ticket de venta
                Dim ventaSvc As New Services.VentaService()
                Dim venta As New Models.Venta() With {
                    .NumeroTicket = ventaSvc.GenerarNumeroTicket(),
                    .UsuarioId = currentUserId,
                    .ClienteId = 1,
                    .CajaId = If(caja IsNot Nothing, caja.Id, 1),
                    .TipoComprobante = "Ticket",
                    .MetodoPago = "Efectivo",
                    .Subtotal = 18500D,
                    .Total = 18500D,
                    .MontoAbonado = 20000D,
                    .MontoEfectivo = 18500D,
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

                Data.DatabaseHelper.ExecuteNonQuery("UPDATE `producto_talles` SET `stock_actual` = 10 WHERE `producto_id` = 1 AND `talle_id` = 2 AND `color` = 'Negro';")
                Dim ventaErr As String = ""
                Dim okVenta = ventaSvc.ProcesarVenta(venta, ventaErr)
                Console.WriteLine("TEST_VENTA:" & okVenta.ToString() & ":Ticket=" & venta.NumeroTicket & ":Err=" & ventaErr)

                ' 10. Probamos las funciones de auditoría gerencial de cajas (incluyendo GetVentasPorCaja corregido)
                Dim listCajas = cajaSvc.GetHistorialCajasAvanzado(DateTime.Today.AddDays(-60), DateTime.Today, Nothing, Nothing, Nothing)
                Console.WriteLine("TEST_HISTORIAL_CAJAS:" & listCajas.Count.ToString())
                If caja IsNot Nothing Then
                    Dim vtasCaja = cajaSvc.GetVentasPorCaja(caja.Id)
                    Dim dtMedios = cajaSvc.GetResumenMediosPagoCaja(caja.Id)
                    Console.WriteLine("TEST_VENTAS_CAJA:" & vtasCaja.Count.ToString() & ":MEDIOS=" & dtMedios.Rows.Count.ToString())
                End If

                ' 11. Probamos AnularVenta y verificamos integridad de Kardex (stock_anterior = 9, stock_posterior = 10)
                Dim anularErr As String = ""
                Dim okAnular = ventaSvc.AnularVenta(venta.Id, currentUserId, "Prueba automatizada de reversión", anularErr)
                Dim dtKardex = Data.DatabaseHelper.ExecuteQuery("SELECT stock_anterior, stock_posterior FROM movimientos_stock WHERE tipo_movimiento = 'Anulacion_Venta' ORDER BY id DESC LIMIT 1;")
                Dim antStock As Integer = If(dtKardex.Rows.Count > 0, Convert.ToInt32(dtKardex.Rows(0)("stock_anterior")), -1)
                Dim postStock As Integer = If(dtKardex.Rows.Count > 0, Convert.ToInt32(dtKardex.Rows(0)("stock_posterior")), -1)
                Console.WriteLine($"TEST_ANULAR_KARDEX:{okAnular}:Ant={antStock}:Post={postStock}:Err={anularErr}")
            Finally
                Config.AppConfig.Settings.CustomDatabasePath = String.Empty
                SqliteConnection.ClearAllPools()
                Try
                    If File.Exists(tempDbPath) Then File.Delete(tempDbPath)
                    If File.Exists(tempDbPath & "-wal") Then File.Delete(tempDbPath & "-wal")
                    If File.Exists(tempDbPath & "-shm") Then File.Delete(tempDbPath & "-shm")
                Catch
                End Try
            End Try
            Return
        End If

        ' Arranque normal de la aplicación gráfica (Windows Forms)
        Application.SetHighDpiMode(HighDpiMode.PerMonitorV2)
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Application.Run(New Forms.FrmLogin())
    End Sub

End Module
