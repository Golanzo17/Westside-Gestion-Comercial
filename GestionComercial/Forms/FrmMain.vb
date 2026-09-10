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

            ' ==================== SIDEBAR (IZQUIERDA) ====================
            pnlSidebar = New Panel() With {
                .Dock = DockStyle.Left,
                .Width = 320,
                .BackColor = UITheme.ColorSidebar,
                .Padding = New Padding(0, 0, 0, 10)
            }

            ' Logo y Marca en Sidebar
            Dim pnlBrand As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 120,
                .BackColor = Color.FromArgb(15, 23, 42),
                .Padding = New Padding(15, 18, 15, 10)
            }

            Dim lblBrandIcon As New Label() With {
                .Text = "👕",
                .Font = New Font("Segoe UI Emoji", 20.0F),
                .ForeColor = Color.White,
                .Location = New Point(15, 22),
                .AutoSize = True
            }

            Dim lblBrandTitle As New Label() With {
                .Text = "GESTIÓN RETAIL",
                .Font = New Font("Segoe UI", 12.0F, FontStyle.Bold),
                .ForeColor = Color.White,
                .Location = New Point(120, 22),
                .AutoSize = True
            }

            Dim lblBrandSubtitle As New Label() With {
                .Text = "Local de Indumentaria",
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular),
                .ForeColor = Color.FromArgb(148, 163, 184),
                .Location = New Point(122, 55),
                .AutoSize = True
            }

            pnlBrand.Controls.AddRange({lblBrandIcon, lblBrandTitle, lblBrandSubtitle})

            ' Botones de Navegación del Menú
            Dim pnlNavMenu As New Panel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(0, 5, 0, 5),
                .AutoScroll = True
            }

            Dim btnHome = CreateNavButton("📊 Panel Principal", AddressOf Nav_Home)
            Dim btnPos = CreateNavButton("🛒 Punto de Venta (POS)", AddressOf Nav_POS)
            Dim btnProductos = CreateNavButton("👕 Catálogo y Talles", AddressOf Nav_Productos)
            Dim btnStock = CreateNavButton("📦 Stock y Reposición", AddressOf Nav_Stock)
            Dim btnClientes = CreateNavButton("👥 Clientes", AddressOf Nav_Clientes)
            Dim btnCaja = CreateNavButton("💵 Caja Diaria y Arqueo", AddressOf Nav_Caja)
            Dim btnReportes = CreateNavButton("📈 Reportes y Ventas", AddressOf Nav_Reportes)
            Dim btnConfig = CreateNavButton("⚙ Configuración", AddressOf Nav_Config)

            Dim navButtons As Button() = {btnConfig, btnReportes, btnCaja, btnClientes, btnStock, btnProductos, btnPos, btnHome}
            For Each b In navButtons
                b.Dock = DockStyle.Top
                pnlNavMenu.Controls.Add(b)
            Next

            ' Botón Cerrar Sesión abajo
            Dim btnLogout = New Button() With {
                .Text = "🚪 Cerrar Sesión",
                .Dock = DockStyle.Bottom,
                .Height = 45
            }
            UITheme.StyleButton(btnLogout, "SIDEBAR")
            btnLogout.ForeColor = UITheme.ColorDanger
            AddHandler btnLogout.Click, AddressOf BtnLogout_Click

            ' Orden exacto de Docking en Sidebar:
            ' Fill primero, Bottom segundo, Top último para evitar que Brand tape la primera opción
            pnlSidebar.Controls.Add(pnlNavMenu)
            pnlSidebar.Controls.Add(btnLogout)
            pnlSidebar.Controls.Add(pnlBrand)

            ' ==================== TOP BAR (SUPERIOR) ====================
            pnlTopBar = New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 60,
                .BackColor = UITheme.ColorSurface,
                .Padding = New Padding(20, 12, 20, 12)
            }

            Dim pnlBorderBottom As New Panel() With {
                .Dock = DockStyle.Bottom,
                .Height = 1,
                .BackColor = UITheme.ColorBorder
            }
            pnlTopBar.Controls.Add(pnlBorderBottom)

            ' Usuario Conectado
            lblUserSession = New Label() With {
                .Text = "👤 " & If(AuthService.CurrentUser IsNot Nothing, $"{AuthService.CurrentUser.NombreCompleto} ({AuthService.CurrentUser.Rol})", "Usuario: Invitado"),
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorTextPrimary,
                .Location = New Point(20, 18),
                .AutoSize = True
            }
            pnlTopBar.Controls.Add(lblUserSession)

            ' Estado de Caja en TopBar
            lblCajaStatus = New Label() With {
                .Text = "● Verificando caja...",
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorWarning,
                .Location = New Point(320, 18),
                .AutoSize = True
            }
            pnlTopBar.Controls.Add(lblCajaStatus)

            ' Botón Venta Rápida
            Dim btnQuickPOS As New Button() With {
                .Text = "+ NUEVA VENTA (F1)",
                .Location = New Point(550, 12),
                .Size = New Size(160, 36)
            }
            UITheme.StyleButton(btnQuickPOS, "Primary")
            AddHandler btnQuickPOS.Click, AddressOf Nav_POS
            pnlTopBar.Controls.Add(btnQuickPOS)

            ' Reloj en vivo
            lblClock = New Label() With {
                .Text = DateTime.Now.ToString("dddd, dd MMMM yyyy - HH:mm:ss"),
                .Font = UITheme.FontRegular,
                .ForeColor = UITheme.ColorTextSecondary,
                .Dock = DockStyle.Right,
                .TextAlign = ContentAlignment.MiddleRight,
                .Width = 320
            }
            pnlTopBar.Controls.Add(lblClock)

            ' ==================== CONTENT HOST (CENTRAL) ====================
            pnlContentHost = New Panel() With {
                .Dock = DockStyle.Fill,
                .BackColor = UITheme.ColorBackground
            }

            BuildDashboardHome()
            pnlContentHost.Controls.Add(pnlDashboardHome)

            ' Orden exacto de Docking en el Form principal:
            ' Fill (pnlContentHost) primero, Top (pnlTopBar) segundo, Left (pnlSidebar) último
            ' De este modo el menú lateral se extiende a la izquierda y las páginas se ubican
            ' a su derecha sin quedar tapadas debajo del menú
            Me.Controls.Add(pnlContentHost)
            Me.Controls.Add(pnlTopBar)
            Me.Controls.Add(pnlSidebar)

            ' Atajos globales
            Me.KeyPreview = True
            AddHandler Me.KeyDown, Sub(s, e)
                                       If e.KeyCode = Keys.F1 Then
                                           e.SuppressKeyPress = True
                                           Nav_POS(Nothing, Nothing)
                                       End If
                                   End Sub
        End Sub

        Private Function CreateNavButton(text As String, clickHandler As EventHandler) As Button
            Dim btn As New Button() With {
                .Text = text,
                .Height = 52
            }
            UITheme.StyleButton(btn, "SIDEBAR")
            AddHandler btn.Click, clickHandler

            ' Pintado del indicador activo (franja de color en borde izquierdo)
            AddHandler btn.Paint, Sub(s, ev)
                                      Dim b = CType(s, Button)
                                      If b Is activeNavBtn Then
                                          ' Franja accent de 4px al borde izquierdo
                                          ev.Graphics.FillRectangle(New SolidBrush(UITheme.ColorPrimary), 0, 0, 4, b.Height)
                                          ' Texto más claro cuando está activo
                                          b.ForeColor = Color.White
                                          b.BackColor = UITheme.ColorSidebarActive
                                      Else
                                          b.ForeColor = Color.FromArgb(203, 213, 225)
                                          b.BackColor = Color.Transparent
                                      End If
                                  End Sub
            Return btn
        End Function

        ''' <summary>Marca el botón nav como activo y refresca todos los demás.</summary>
        Private Sub SetActiveNavButton(btn As Button)
            activeNavBtn = btn
            ' Forzar repintado de todos los botones del nav
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
                .Padding = New Padding(25)
            }

            Dim lblDashTitle As New Label() With {
                .Text = "Panel de Control y Resumen de Hoy",
                .Font = UITheme.FontHeading,
                .ForeColor = UITheme.ColorTextPrimary,
                .Location = New Point(25, 20),
                .AutoSize = True
            }

            Dim lblDashSub As New Label() With {
                .Text = "Métricas en tiempo real de ventas, inventario y movimientos del local.",
                .Font = UITheme.FontRegular,
                .ForeColor = UITheme.ColorTextSecondary,
                .Location = New Point(27, 65),
                .AutoSize = True
            }

            pnlDashboardHome.Controls.AddRange({lblDashTitle, lblDashSub})

            ' Fila de 4 KPI Cards
            Dim pnlCardsRow As New FlowLayoutPanel() With {
                .Location = New Point(20, 110),
                .Size = New Size(980, 115),
                .AutoSize = True
            }

            Dim card1 = UITheme.CreateKpiCard("VENTAS DE HOY", "$ 0,00", "Total facturado en el día", UITheme.ColorSuccess)
            lblKpiVentasHoy = CType(card1.Controls(2), Label)

            Dim card2 = UITheme.CreateKpiCard("TICKETS EMITIDOS", "0", "Ventas completadas", UITheme.ColorPrimary)
            lblKpiTicketsHoy = CType(card2.Controls(2), Label)

            Dim card3 = UITheme.CreateKpiCard("PRENDAS STOCK CRÍTICO", "0", "Artículos a reponer", UITheme.ColorWarning)
            lblKpiStockCritico = CType(card3.Controls(2), Label)

            Dim card4 = UITheme.CreateKpiCard("CAJA DEL TURNO", "$ 0,00", "Efectivo disponible", UITheme.ColorInfo)
            lblKpiCajaTurno = CType(card4.Controls(2), Label)

            pnlCardsRow.Controls.AddRange({card1, card2, card3, card4})
            pnlDashboardHome.Controls.Add(pnlCardsRow)

            ' Tarjetas de Acceso Rápido / Módulos
            Dim pnlAcciones As New GroupBox() With {
                .Text = "Acceso Rápido a Operaciones",
                .Location = New Point(25, 245),
                .Size = New Size(950, 160),
                .BackColor = UITheme.ColorSurface,
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorPrimaryDark,
                .Padding = New Padding(20)
            }

            Dim btnAccionPOS As New Button() With {.Text = "🛒 Realizar Nueva Venta", .Location = New Point(25, 40), .Size = New Size(210, 48)}
            UITheme.StyleButton(btnAccionPOS, "Primary")
            AddHandler btnAccionPOS.Click, AddressOf Nav_POS

            Dim btnAccionPrenda As New Button() With {.Text = "+ Cargar Nueva Prenda", .Location = New Point(255, 40), .Size = New Size(210, 48)}
            UITheme.StyleButton(btnAccionPrenda, "Success")
            AddHandler btnAccionPrenda.Click, AddressOf Nav_Productos

            Dim btnAccionStock As New Button() With {.Text = "📦 Ingreso de Mercadería", .Location = New Point(485, 40), .Size = New Size(210, 48)}
            UITheme.StyleButton(btnAccionStock, "Secondary")
            AddHandler btnAccionStock.Click, AddressOf Nav_Stock

            Dim btnAccionCaja As New Button() With {.Text = "💵 Arqueo / Cierre de Caja", .Location = New Point(715, 40), .Size = New Size(210, 48)}
            UITheme.StyleButton(btnAccionCaja, "Secondary")
            AddHandler btnAccionCaja.Click, AddressOf Nav_Caja

            pnlAcciones.Controls.AddRange({btnAccionPOS, btnAccionPrenda, btnAccionStock, btnAccionCaja})
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

                Dim userId = If(AuthService.CurrentUser IsNot Nothing, AuthService.CurrentUser.Id, 1)
                Dim caja = cajaService.GetCajaAbierta(userId)
                If caja IsNot Nothing Then
                    lblCajaStatus.Text = $"🟢 Caja Abierta (Turno #{caja.Id})"
                    lblCajaStatus.ForeColor = UITheme.ColorSuccess
                    Dim esp = caja.MontoInicial + caja.TotalVentasEfectivo + caja.TotalIngresos - caja.TotalEgresos
                    lblKpiCajaTurno.Text = esp.ToString("C2")
                Else
                    lblCajaStatus.Text = "🔴 Caja Cerrada"
                    lblCajaStatus.ForeColor = UITheme.ColorDanger
                    lblKpiCajaTurno.Text = "$ 0,00"
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
            If sender IsNot Nothing AndAlso TypeOf sender Is Button Then SetActiveNavButton(CType(sender, Button))
            OpenChildForm(New FrmVentasPOS())
        End Sub

        Private Sub Nav_Productos(sender As Object, e As EventArgs)
            If sender IsNot Nothing AndAlso TypeOf sender Is Button Then SetActiveNavButton(CType(sender, Button))
            OpenChildForm(New FrmProductos())
        End Sub

        Private Sub Nav_Stock(sender As Object, e As EventArgs)
            If sender IsNot Nothing AndAlso TypeOf sender Is Button Then SetActiveNavButton(CType(sender, Button))
            OpenChildForm(New FrmStock())
        End Sub

        Private Sub Nav_Clientes(sender As Object, e As EventArgs)
            If sender IsNot Nothing AndAlso TypeOf sender Is Button Then SetActiveNavButton(CType(sender, Button))
            OpenChildForm(New FrmClientes())
        End Sub

        Private Sub Nav_Caja(sender As Object, e As EventArgs)
            If sender IsNot Nothing AndAlso TypeOf sender Is Button Then SetActiveNavButton(CType(sender, Button))
            OpenChildForm(New FrmCaja())
        End Sub

        Private Sub Nav_Reportes(sender As Object, e As EventArgs)
            If sender IsNot Nothing AndAlso TypeOf sender Is Button Then SetActiveNavButton(CType(sender, Button))
            OpenChildForm(New FrmReportes())
        End Sub

        Private Sub Nav_Sync(sender As Object, e As EventArgs)
            OpenChildForm(New FrmSincronizacionEcommerce())
        End Sub

        Private Sub Nav_Config(sender As Object, e As EventArgs)
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
