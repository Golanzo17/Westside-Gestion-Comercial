' ARCHIVO: FrmMain.vb
' PROPÓSITO: Ventana Principal, Panel de Navegación (Sidebar) y Dashboard de KPIs.
' Este formulario es el centro de control del sistema:
' 1. Patrón Contenedor (MDI Moderno): Para evitar ventanas flotantes desordenadas,
'    utilizamos un panel central (pnlContentHost) donde incrustamos los formularios
'    hijos (FrmVentasPOS, FrmCaja, etc.) configurando TopLevel = False y Dock = Fill.
' 2. Segregación de Roles (Principle of Least Privilege):
'    Configuramos la visibilidad de los accesos según el rol activo:
'    - Administrador: Gestiona Configuración y Equipo de Usuarios.
'    - Gerente: Ve Catálogo, Stock, Caja y Reportes, con potestad de autorizar excepciones comerciales.
'    - Vendedor: Opera POS, Caja, Clientes y Catálogo; restringido en Configuración y Usuarios.
' 3. Dashboard con KPIs en Vivo: Muestra métricas comerciales del día (total facturado,
'    tickets emitidos, prendas en stock crítico y estado de caja abierta).


Imports System.Drawing
Imports System.Windows.Forms
Imports GestionComercial.Models
Imports GestionComercial.Services
Imports GestionComercial.UI

Namespace Forms
    Public Class FrmMain
        Inherits Form

        Private pnlSidebar As Panel
        Private pnlTopBar As Panel
        Private pnlContentHost As Panel
        Private pnlDashboardHome As Panel

        ' Indicadores TopBar
        Private lblUserSession As Label
        Private lblCajaStatus As Label
        Private lblClock As Label
        Private timerClock As Timer

        ' KPI Cards
        Private lblKpiVentasHoy As Label
        Private lblKpiTicketsHoy As Label
        Private lblKpiStockCritico As Label
        Private lblKpiCajaTurno As Label

        Private reporteService As New ReporteService()
        Private cajaService As New CajaService()
        Private usuarioService As New UsuarioService()
        Private activeChildForm As Form = Nothing
        Private activeNavBtn As Button = Nothing  ' Botón de nav activo (para indicador visual)

        Public Sub New()
            InitializeUI()
            StartClock()
            LoadDashboardData()
        End Sub

        Private Sub InitializeUI()
            Me.Text = "Software de Gestión Comercial - Local de Ropa & Indumentaria"
            Me.Size = New Size(1280, 800)
            Me.MinimumSize = New Size(1024, 700)
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.BackColor = UITheme.ColorBackground
            Me.Font = UITheme.FontRegular
            Me.AutoScaleMode = AutoScaleMode.Dpi

            ' ─── Panel lateral de navegación (Sidebar) ───
            pnlSidebar = New Panel() With {
                .Dock = DockStyle.Left,
                .Width = 260,
                .BackColor = UITheme.ColorSidebar,
                .Padding = New Padding(0, 0, 0, 10)
            }

            ' Logo y Marca en Sidebar
            Dim pnlBrand As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 110,
                .BackColor = UITheme.ColorSecondary,
                .Padding = New Padding(15, 18, 15, 10)
            }

            Dim lblBrandIcon As New Label() With {
                .Text = "WESTSIDE",
                .Font = New Font("Segoe UI", 18.0F, FontStyle.Bold),
                .ForeColor = Color.White,
                .Dock = DockStyle.Top,
                .Height = 38,
                .TextAlign = ContentAlignment.MiddleCenter
            }

            Dim lblBrandTitle As New Label() With {
                .Text = "GESTIÓN COMERCIAL",
                .Font = New Font("Segoe UI", 8.0F, FontStyle.Bold),
                .ForeColor = UITheme.ColorPrimaryLight,
                .Dock = DockStyle.Top,
                .Height = 20,
                .TextAlign = ContentAlignment.MiddleCenter
            }

            Dim lblBrandSubtitle As New Label() With {
                .Text = "Indumentaria & Calzado",
                .Font = UITheme.FontSmall,
                .ForeColor = Color.FromArgb(180, 200, 220),
                .Dock = DockStyle.Top,
                .Height = 18,
                .TextAlign = ContentAlignment.MiddleCenter
            }

            pnlBrand.Controls.AddRange({lblBrandIcon, lblBrandTitle, lblBrandSubtitle})

            ' ─── Botones de Navegación agrupados por sección ───
            Dim pnlNavMenu As New Panel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(0, 5, 0, 5),
                .AutoScroll = True
            }

            Dim btnHome = CreateNavButton("Panel principal", AddressOf Nav_Home)
            Dim btnPos = CreateNavButton("Punto de venta (POS)", AddressOf Nav_POS)
            Dim btnProductos = CreateNavButton("Catálogo y talles", AddressOf Nav_Productos)
            Dim btnStock = CreateNavButton("Stock y reposición", AddressOf Nav_Stock)
            Dim btnClientes = CreateNavButton("Clientes", AddressOf Nav_Clientes)
            Dim btnCaja = CreateNavButton("Caja diaria y arqueo", AddressOf Nav_Caja)
            Dim btnReportes = CreateNavButton("Reportes y ventas", AddressOf Nav_Reportes)
            Dim btnUsuarios = CreateNavButton("Usuarios y equipo", AddressOf Nav_Usuarios)
            Dim btnBackup = CreateNavButton("Copias de seguridad", AddressOf Nav_Backup)
            Dim btnConfig = CreateNavButton("Configuración", AddressOf Nav_Config)

            ' Configuración de accesos del menú lateral según el rol activo
            btnPos.Visible = AuthService.IsVendor
            btnStock.Visible = AuthService.IsManager
            btnProductos.Visible = Not AuthService.IsAdmin
            btnCaja.Visible = AuthService.IsVendor
            btnReportes.Visible = AuthService.IsManager OrElse AuthService.IsVendor
            btnUsuarios.Visible = AuthService.IsAdmin
            btnBackup.Visible = AuthService.IsAdmin
            btnConfig.Visible = AuthService.IsAdmin

            ' Separadores de sección visuales para agrupar la navegación
            Dim secSistema = UITheme.CreateSidebarSection("Sistema")
            secSistema.Visible = AuthService.IsAdmin
            Dim divSistema = UITheme.CreateDivider(1, 6, 2)
            divSistema.Visible = AuthService.IsAdmin

            Dim secGestion = UITheme.CreateSidebarSection("Gestión")
            secGestion.Visible = AuthService.IsManager OrElse AuthService.IsVendor

            Dim secOperaciones = UITheme.CreateSidebarSection("Operaciones")
            secOperaciones.Visible = AuthService.IsVendor

            ' Orden de docking: se agregan en orden inverso (Bottom-up)
            ' Resultado visual: Home > [Operaciones: POS, Caja, Clientes] > [Gestión: Catálogo, Stock, Reportes] > [Sistema: Usuarios, Backup, Config]
            pnlNavMenu.Controls.Add(btnConfig)
            pnlNavMenu.Controls.Add(btnBackup)
            pnlNavMenu.Controls.Add(btnUsuarios)
            pnlNavMenu.Controls.Add(secSistema)
            pnlNavMenu.Controls.Add(divSistema)
            pnlNavMenu.Controls.Add(btnReportes)
            pnlNavMenu.Controls.Add(btnStock)
            pnlNavMenu.Controls.Add(btnProductos)
            pnlNavMenu.Controls.Add(secGestion)
            pnlNavMenu.Controls.Add(btnCaja)
            pnlNavMenu.Controls.Add(btnClientes)
            pnlNavMenu.Controls.Add(btnPos)
            pnlNavMenu.Controls.Add(secOperaciones)
            pnlNavMenu.Controls.Add(btnHome)

            ' Botón Cerrar Sesión abajo
            Dim btnLogout = New Button() With {
                .Text = "Cerrar sesión",
                .Dock = DockStyle.Bottom,
                .Height = 45
            }
            UITheme.StyleButton(btnLogout, "SIDEBAR")
            btnLogout.ForeColor = UITheme.ColorDanger
            AddHandler btnLogout.Click, AddressOf BtnLogout_Click

            ' Orden exacto de Docking en Sidebar:
            pnlSidebar.Controls.Add(pnlNavMenu)
            pnlSidebar.Controls.Add(btnLogout)
            pnlSidebar.Controls.Add(pnlBrand)

            ' ─── Barra superior de estado (TopBar) ───
            pnlTopBar = New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 60,
                .BackColor = UITheme.ColorSurface
            }

            Dim pnlBorderBottom As New Panel() With {
                .Dock = DockStyle.Bottom,
                .Height = 1,
                .BackColor = UITheme.ColorBorder
            }
            pnlTopBar.Controls.Add(pnlBorderBottom)

            Dim tlpTopBar As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 4,
                .RowCount = 1,
                .Padding = New Padding(12, 0, 12, 0)
            }
            tlpTopBar.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 30.0F))  ' usuario
            tlpTopBar.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 30.0F))  ' caja
            tlpTopBar.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))         ' botón
            tlpTopBar.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 40.0F))  ' reloj
            tlpTopBar.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            ' Usuario Conectado con Distintivo de Rol
            Dim rolBadge As String = If(AuthService.IsAdmin, "ADMINISTRADOR", If(AuthService.IsManager, "GERENTE", "VENDEDOR"))
            Dim nombreUser As String = If(AuthService.CurrentUser IsNot Nothing, AuthService.CurrentUser.NombreCompleto, "Invitado")
            lblUserSession = New Label() With {
                .Text = $"{rolBadge}  ·  {nombreUser}",
                .Font = UITheme.FontBold,
                .ForeColor = If(AuthService.IsAdminOrManager, UITheme.ColorPrimaryDark, UITheme.ColorTextPrimary),
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleLeft,
                .AutoEllipsis = True
            }

            ' Estado de Caja en TopBar (Solo visible para Gerente y Vendedor)
            lblCajaStatus = New Label() With {
                .Text = "Verificando caja...",
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorWarning,
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleLeft,
                .AutoEllipsis = True,
                .Visible = Not AuthService.IsAdmin
            }

            ' Botón Venta Rápida
            Dim btnQuickPOS As New Button() With {
                .Text = "+ NUEVA VENTA (F1)",
                .Size = New Size(195, 36),
                .Anchor = AnchorStyles.None,
                .Visible = AuthService.IsVendor
            }
            UITheme.StyleButton(btnQuickPOS, "Primary")
            AddHandler btnQuickPOS.Click, AddressOf Nav_POS

            ' Reloj en vivo
            lblClock = New Label() With {
                .Text = DateTime.Now.ToString("dddd, dd MMMM yyyy - HH:mm:ss"),
                .Font = UITheme.FontRegular,
                .ForeColor = UITheme.ColorTextSecondary,
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleRight
            }

            tlpTopBar.Controls.Add(lblUserSession, 0, 0)
            tlpTopBar.Controls.Add(lblCajaStatus, 1, 0)
            tlpTopBar.Controls.Add(btnQuickPOS, 2, 0)
            tlpTopBar.Controls.Add(lblClock, 3, 0)
            pnlTopBar.Controls.Add(tlpTopBar)

            ' Contenedor central de vistas
            pnlContentHost = New Panel() With {
                .Dock = DockStyle.Fill,
                .BackColor = UITheme.ColorBackground
            }

            BuildDashboardHome()
            pnlContentHost.Controls.Add(pnlDashboardHome)

            Me.Controls.Add(pnlContentHost)
            Me.Controls.Add(pnlTopBar)
            Me.Controls.Add(pnlSidebar)

            ' Atajos globales
            Me.KeyPreview = True
            AddHandler Me.KeyDown, Sub(s, e)
                                       If e.KeyCode = Keys.F1 AndAlso AuthService.IsVendor Then
                                           e.SuppressKeyPress = True
                                           Nav_POS(Nothing, Nothing)
                                       End If
                                   End Sub
        End Sub

        Private Function CreateNavButton(text As String, clickHandler As EventHandler) As Button
            Dim btn As New Button() With {
                .Text = text,
                .Height = 44,
                .Dock = DockStyle.Top
            }
            UITheme.StyleButton(btn, "SIDEBAR")
            AddHandler btn.Click, clickHandler

            AddHandler btn.Paint, Sub(s, ev)
                                      Dim b = CType(s, Button)
                                      If b Is activeNavBtn Then
                                          ' Indicador activo: franja izquierda + fondo gradiente sutil
                                          ev.Graphics.FillRectangle(New SolidBrush(UITheme.ColorPrimary), 0, 0, 4, b.Height)
                                          b.ForeColor = Color.White
                                          b.BackColor = UITheme.ColorSidebarActive
                                          b.Font = UITheme.FontBold
                                      Else
                                           b.ForeColor = UITheme.ColorTextOnDark
                                          b.BackColor = Color.Transparent
                                          b.Font = UITheme.FontNavItem
                                      End If
                                  End Sub
            Return btn
        End Function

        Private Sub SetActiveNavButton(btn As Button)
            activeNavBtn = btn
            For Each ctrl As Control In pnlSidebar.Controls
                If TypeOf ctrl Is Panel Then
                    For Each inner As Control In ctrl.Controls
                        If TypeOf inner Is Button Then inner.Invalidate()
                    Next
                ElseIf TypeOf ctrl Is Button Then
                    ctrl.Invalidate()
                End If
            Next
        End Sub

        Private Sub BuildDashboardHome()
            pnlDashboardHome = New Panel() With {
                .Dock = DockStyle.Fill,
                .AutoScroll = True,
                .Padding = New Padding(30)
            }

            Dim lblDashTitle As New Label() With {
                .Text = "Panel de Control",
                .Font = UITheme.FontHeadingLarge,
                .ForeColor = UITheme.ColorTextPrimary,
                .Location = New Point(30, 20),
                .AutoSize = True
            }

            Dim lblDashSub As New Label() With {
                .Text = "Métricas en tiempo real de ventas, inventario y movimientos del local.",
                .Font = UITheme.FontRegular,
                .ForeColor = UITheme.ColorTextSecondary,
                .Location = New Point(32, 58),
                .AutoSize = True
            }

            pnlDashboardHome.Controls.AddRange({lblDashTitle, lblDashSub})

            ' Fila de 4 KPI Cards
            Dim pnlCardsRow As New FlowLayoutPanel() With {
                .Location = New Point(22, 100),
                .Size = New Size(980, 140),
                .AutoSize = True
            }

            Dim card1 = UITheme.CreateKpiCard("VENTAS DE HOY", "$ 0,00", "Total facturado en el día", UITheme.ColorSuccess)
            lblKpiVentasHoy = CType(card1.Controls(2), Label)

            Dim card2 = UITheme.CreateKpiCard("TICKETS EMITIDOS", "0", "Ventas completadas", UITheme.ColorPrimary)
            lblKpiTicketsHoy = CType(card2.Controls(2), Label)

            Dim card3 = UITheme.CreateKpiCard("PRENDAS STOCK CRÍTICO", "0", "Artículos a reponer", UITheme.ColorWarning)
            lblKpiStockCritico = CType(card3.Controls(2), Label)

            ' KPI Card 4: Adaptado por rol (Administrador ve 'USUARIOS REGISTRADOS', Gerente y Vendedor ven 'CAJA DEL TURNO')
            Dim card4 As Panel
            If AuthService.IsAdmin Then
                card4 = UITheme.CreateKpiCard("USUARIOS REGISTRADOS", "0", "Cuentas en el sistema", UITheme.ColorInfo)
            Else
                card4 = UITheme.CreateKpiCard("CAJA DEL TURNO", "$ 0,00", "Efectivo disponible", UITheme.ColorInfo)
            End If
            lblKpiCajaTurno = CType(card4.Controls(2), Label)

            pnlCardsRow.Controls.AddRange({card1, card2, card3, card4})
            pnlDashboardHome.Controls.Add(pnlCardsRow)

            ' ─── Tarjetas de Acceso Rápido filtradas estrictamente por rol ───
            Dim pnlAcciones As New Panel() With {
                .Location = New Point(30, 260),
                .Size = New Size(960, 115),
                .BackColor = UITheme.ColorSurface,
                .Padding = New Padding(20, 14, 20, 14)
            }
            AddHandler pnlAcciones.Paint, Sub(s, ev)
                                              Using pen As New Pen(UITheme.ColorBorder, 1)
                                                  ev.Graphics.DrawRectangle(pen, 0, 0, pnlAcciones.Width - 1, pnlAcciones.Height - 1)
                                              End Using
                                          End Sub

            Dim lblAccionesTitle As New Label() With {
                .Text = "ACCIONES RÁPIDAS",
                .Font = UITheme.FontSectionTitle,
                .ForeColor = UITheme.ColorTextSecondary,
                .Dock = DockStyle.Top,
                .Height = 24
            }

            Dim flpAcciones As New FlowLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(0, 4, 0, 0),
                .WrapContents = False,
                .AutoScroll = False
            }

            ' Operaciones para Vendedor (POS y Caja)
            If AuthService.IsVendor Then
                Dim btnAccionPOS As New Button() With {
                    .Text = "Realizar Nueva Venta",
                    .Height = 44,
                    .AutoSize = True,
                    .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    .Padding = New Padding(16, 0, 16, 0),
                    .Margin = New Padding(0, 0, 10, 8)
                }
                UITheme.StyleButton(btnAccionPOS, "Primary")
                AddHandler btnAccionPOS.Click, AddressOf Nav_POS
                flpAcciones.Controls.Add(btnAccionPOS)

                Dim btnAccionCaja As New Button() With {
                    .Text = "Arqueo / Cierre de Caja",
                    .Height = 44,
                    .AutoSize = True,
                    .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    .Padding = New Padding(16, 0, 16, 0),
                    .Margin = New Padding(0, 0, 10, 8)
                }
                UITheme.StyleButton(btnAccionCaja, "Secondary")
                AddHandler btnAccionCaja.Click, AddressOf Nav_Caja
                flpAcciones.Controls.Add(btnAccionCaja)
            End If

            ' Operaciones de Gestión (Admin y Gerente)
            If AuthService.IsAdminOrManager Then
                If AuthService.IsManager Then
                    Dim btnAccionPrenda As New Button() With {
                        .Text = "Cargar Nueva Prenda",
                        .Height = 44,
                        .AutoSize = True,
                        .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                        .Padding = New Padding(16, 0, 16, 0),
                        .Margin = New Padding(0, 0, 10, 8)
                    }
                    UITheme.StyleButton(btnAccionPrenda, "Success")
                    AddHandler btnAccionPrenda.Click, AddressOf Nav_Productos
                    flpAcciones.Controls.Add(btnAccionPrenda)
                End If

                If AuthService.IsManager Then
                    Dim btnAccionStock As New Button() With {
                        .Text = "Ingreso de Mercadería",
                        .Height = 44,
                        .AutoSize = True,
                        .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                        .Padding = New Padding(16, 0, 16, 0),
                        .Margin = New Padding(0, 0, 10, 8)
                    }
                    UITheme.StyleButton(btnAccionStock, "Secondary")
                    AddHandler btnAccionStock.Click, AddressOf Nav_Stock
                    flpAcciones.Controls.Add(btnAccionStock)
                End If
            End If

            ' Operaciones específicas para Vendedor
            If AuthService.IsVendor Then
                Dim btnAccionClientes As New Button() With {
                    .Text = "Directorio de Clientes",
                    .Height = 44,
                    .AutoSize = True,
                    .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    .Padding = New Padding(16, 0, 16, 0),
                    .Margin = New Padding(0, 0, 10, 8)
                }
                UITheme.StyleButton(btnAccionClientes, "Secondary")
                AddHandler btnAccionClientes.Click, AddressOf Nav_Clientes
                flpAcciones.Controls.Add(btnAccionClientes)
            End If

            ' Operaciones específicas para Administrador
            If AuthService.IsAdmin Then
                Dim btnAccionUsuarios As New Button() With {
                    .Text = "Gestión de Usuarios",
                    .Height = 44,
                    .AutoSize = True,
                    .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    .Padding = New Padding(16, 0, 16, 0),
                    .Margin = New Padding(0, 0, 10, 8)
                }
                UITheme.StyleButton(btnAccionUsuarios, "Secondary")
                AddHandler btnAccionUsuarios.Click, AddressOf Nav_Usuarios
                flpAcciones.Controls.Add(btnAccionUsuarios)

                Dim btnAccionBackup As New Button() With {
                    .Text = "Copias de Seguridad",
                    .Height = 44,
                    .AutoSize = True,
                    .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    .Padding = New Padding(16, 0, 16, 0),
                    .Margin = New Padding(0, 0, 10, 8)
                }
                UITheme.StyleButton(btnAccionBackup, "Secondary")
                AddHandler btnAccionBackup.Click, AddressOf Nav_Backup
                flpAcciones.Controls.Add(btnAccionBackup)

                Dim btnAccionConfig As New Button() With {
                    .Text = "Configuración del Sistema",
                    .Height = 44,
                    .AutoSize = True,
                    .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    .Padding = New Padding(16, 0, 16, 0),
                    .Margin = New Padding(0, 0, 10, 8)
                }
                UITheme.StyleButton(btnAccionConfig, "Secondary")
                AddHandler btnAccionConfig.Click, AddressOf Nav_Config
                flpAcciones.Controls.Add(btnAccionConfig)
            End If

            pnlAcciones.Controls.Add(flpAcciones)
            pnlAcciones.Controls.Add(lblAccionesTitle)
            lblAccionesTitle.BringToFront()
            flpAcciones.BringToFront()
            pnlDashboardHome.Controls.Add(pnlAcciones)
        End Sub

        Private Sub StartClock()
            timerClock = New Timer() With {.Interval = 1000}
            AddHandler timerClock.Tick, Sub()
                                            lblClock.Text = DateTime.Now.ToString("dddd, dd MMMM yyyy - HH:mm:ss")
                                        End Sub
            timerClock.Start()
        End Sub

        Public Sub LoadDashboardData()
            Try
                Dim resumen = reporteService.GetResumenHoy()
                lblKpiVentasHoy.Text = resumen.TotalVendido.ToString("C2")
                lblKpiTicketsHoy.Text = resumen.CantidadTickets.ToString()
                lblKpiStockCritico.Text = resumen.ArticulosStockBajo.ToString()

                If AuthService.IsAdmin Then
                    lblCajaStatus.Visible = False
                    If lblKpiCajaTurno IsNot Nothing Then
                        Dim usrList = usuarioService.GetUsuarios()
                        lblKpiCajaTurno.Text = usrList.Count.ToString()
                    End If
                Else
                    lblCajaStatus.Visible = True
                    Dim userId = If(AuthService.CurrentUser IsNot Nothing, AuthService.CurrentUser.Id, 1)
                    Dim caja = cajaService.GetCajaAbierta(userId)
                    If caja IsNot Nothing Then
                        lblCajaStatus.Text = $"Caja Abierta  ·  Turno #{caja.Id}"
                        lblCajaStatus.ForeColor = UITheme.ColorSuccess
                        Dim esp = caja.MontoInicial + caja.TotalVentasEfectivo + caja.TotalIngresos - caja.TotalEgresos
                        If lblKpiCajaTurno IsNot Nothing Then
                            lblKpiCajaTurno.Text = If(AuthService.IsAdminOrManager, esp.ToString("C2"), "Arqueo Ciego")
                        End If
                    Else
                        lblCajaStatus.Text = "Caja Cerrada"
                        lblCajaStatus.ForeColor = UITheme.ColorDanger
                        If lblKpiCajaTurno IsNot Nothing Then
                            lblKpiCajaTurno.Text = "$ 0,00"
                        End If
                    End If
                End If
            Catch ex As Exception
                ' Si no hay base de datos conectada
            End Try
        End Sub

        Private Sub OpenChildForm(childForm As Form)
            If activeChildForm IsNot Nothing Then
                activeChildForm.Close()
            End If

            activeChildForm = childForm
            childForm.TopLevel = False
            childForm.FormBorderStyle = FormBorderStyle.None
            childForm.Dock = DockStyle.Fill
            childForm.AutoScaleMode = AutoScaleMode.None

            pnlContentHost.Controls.Clear()
            pnlContentHost.Controls.Add(childForm)
            pnlContentHost.Tag = childForm
            childForm.BringToFront()
            childForm.Show()
        End Sub

        Private Sub Nav_Home(sender As Object, e As EventArgs)
            SetActiveNavButton(Nothing)
            If activeChildForm IsNot Nothing Then
                activeChildForm.Close()
                activeChildForm = Nothing
            End If
            pnlContentHost.Controls.Clear()
            pnlContentHost.Controls.Add(pnlDashboardHome)
            LoadDashboardData()
        End Sub

        Private Sub Nav_POS(sender As Object, e As EventArgs)
            If Not AuthService.IsVendor Then Return
            If sender IsNot Nothing AndAlso TypeOf sender Is Button Then SetActiveNavButton(CType(sender, Button))
            OpenChildForm(New FrmVentasPOS())
        End Sub

        Private Sub Nav_Productos(sender As Object, e As EventArgs)
            If AuthService.IsAdmin Then Return
            If sender IsNot Nothing AndAlso TypeOf sender Is Button Then SetActiveNavButton(CType(sender, Button))
            OpenChildForm(New FrmProductos())
        End Sub

        Private Sub Nav_Stock(sender As Object, e As EventArgs)
            If Not AuthService.IsManager Then Return
            If sender IsNot Nothing AndAlso TypeOf sender Is Button Then SetActiveNavButton(CType(sender, Button))
            OpenChildForm(New FrmStock())
        End Sub

        Private Sub Nav_Clientes(sender As Object, e As EventArgs)
            If sender IsNot Nothing AndAlso TypeOf sender Is Button Then SetActiveNavButton(CType(sender, Button))
            OpenChildForm(New FrmClientes())
        End Sub

        Private Sub Nav_Caja(sender As Object, e As EventArgs)
            If Not AuthService.IsVendor Then Return
            If sender IsNot Nothing AndAlso TypeOf sender Is Button Then SetActiveNavButton(CType(sender, Button))
            OpenChildForm(New FrmCaja())
        End Sub

        Private Sub Nav_Reportes(sender As Object, e As EventArgs)
            If AuthService.IsAdmin OrElse Not (AuthService.IsManager OrElse AuthService.IsVendor) Then
                Return
            End If
            If sender IsNot Nothing AndAlso TypeOf sender Is Button Then SetActiveNavButton(CType(sender, Button))
            OpenChildForm(New FrmReportes())
        End Sub

        Private Sub Nav_Usuarios(sender As Object, e As EventArgs)
            If Not AuthService.IsAdmin Then
                MessageBox.Show("El acceso al módulo de usuarios y equipo está reservado exclusivamente para Administradores.", "Acceso Restringido", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            If sender IsNot Nothing AndAlso TypeOf sender Is Button Then SetActiveNavButton(CType(sender, Button))
            OpenChildForm(New FrmUsuarios())
        End Sub

        Private Sub Nav_Backup(sender As Object, e As EventArgs)
            If Not AuthService.IsAdmin Then
                MessageBox.Show("El acceso a las copias de seguridad está reservado exclusivamente para Administradores.", "Acceso Restringido", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            If sender IsNot Nothing AndAlso TypeOf sender Is Button Then SetActiveNavButton(CType(sender, Button))
            OpenChildForm(New FrmBackup())
        End Sub


        Private Sub Nav_Config(sender As Object, e As EventArgs)
            If Not AuthService.SolicitarAutorizacionAdmin(Me, "La configuración general del sistema está reservada para Administradores.") Then
                Return
            End If
            Dim frm As New FrmConfiguracion()
            If frm.ShowDialog() = DialogResult.OK Then
                LoadDashboardData()
            End If
        End Sub

        Private Sub BtnLogout_Click(sender As Object, e As EventArgs)
            If MessageBox.Show("¿Deseas cerrar tu sesión actual?", "Cerrar Sesión", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                AuthService.Logout()
                Me.Hide()
                Dim frmLog As New FrmLogin()
                frmLog.Show()
            End If
        End Sub

    End Class
End Namespace
