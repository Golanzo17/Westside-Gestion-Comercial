Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms
Imports GestionComercial.Models
Imports GestionComercial.Services
Imports GestionComercial.UI

' ARCHIVO: FrmReportes.vb
' PROPÓSITO: Inteligencia Comercial, Análisis de Ventas por Cliente, Ranking de Rentabilidad y Anulación
' - Métricas Financieras y Ventas por Cliente:
'   Permite analizar compras agrupadas por cliente en un rango de fechas flexible, evaluando volumen
'   de prendas llevadas (mayor salida), ticket acumulado y rentabilidad neta aportada.
' - Ranking Dual de Artículos:
'   Permite alternar entre "Mayor Salida" (rotación de unidades vendidas) y "Mayor Margen de Ganancia"
'   (utilidad bruta recaudada descontando el costo de reposición de la prenda).
' - Separación de Privilegios:
'   El rol Vendedor accede al ranking de tendencias sin visualizar costos ni márgenes financieros confidenciales.
' - Anulación de Ventas con Auditoría y Restitución de Stock:
'   Operación protegida para roles Admin/Manager con transacción atómica.

Namespace Forms
    Public Class FrmReportes
        Inherits Form

        Private ventaService As New VentaService()
        Private reporteService As New ReporteService()
        Private clienteService As New ClienteService()
        Private usuarioService As New UsuarioService()
        Private configService As New ConfiguracionService()

        Private cboRapidoFechas As ComboBox
        Private dtpDesde As DateTimePicker
        Private dtpHasta As DateTimePicker
        Private cboVendedorGeneral As ComboBox
        Private btnFiltrar As Button
        Private btnDescargarPdf As Button
        Private btnAnularVenta As Button

        ' Tarjetas KPI Financieras
        Private lblTotalVentas As Label
        Private lblTotalVentasSub As Label
        Private lblCostoMercaderia As Label
        Private lblCostoMercaderiaSub As Label
        Private lblGananciaNeta As Label
        Private lblGananciaNetaSub As Label
        Private lblTicketPromedio As Label
        Private lblTicketPromedioSub As Label

        Private tabs As TabControl
        Private tabVentas As TabPage
        Private cboClienteGeneral As ComboBox
        Private cboMetodoPagoGeneral As ComboBox
        Private cboEstadoGeneral As ComboBox
        Private dgvVentas As DataGridView

        ' Pestaña: Anulaciones y Notas de Crédito
        Private tabAnulaciones As TabPage
        Private lblResumenAnulaciones As Label
        Private dgvAnulaciones As DataGridView

        ' Pestaña: Ventas por Cliente
        Private tabClientes As TabPage
        Private cboClientesFiltro As ComboBox
        Private cboOrdenClientes As ComboBox
        Private lblResumenClientes As Label
        Private dgvClientesReporte As DataGridView
        Private dgvTicketsCliente As DataGridView
        Private lblTicketsClienteHeader As Label
        Private btnVerDetalleTicketCliente As Button
        Private btnDescargarPdfTicketCliente As Button

        ' Pestaña: Ventas por Vendedor
        Private tabVendedores As TabPage
        Private cboVendedoresFiltro As ComboBox
        Private cboOrdenVendedores As ComboBox
        Private lblResumenVendedores As Label
        Private dgvVendedoresReporte As DataGridView
        Private dgvTicketsVendedor As DataGridView
        Private lblTicketsVendedorHeader As Label
        Private btnVerDetalleTicketVendedor As Button
        Private btnDescargarPdfTicketVendedor As Button

        ' Pestaña: Ranking de Prendas y Stock Inmovilizado
        Private tabRanking As TabPage
        Private rbMayorSalida As RadioButton
        Private rbMayorMargen As RadioButton
        Private rbStockInmovilizado As RadioButton
        Private cboLimiteRanking As ComboBox
        Private lblResumenRanking As Label
        Private dgvRanking As DataGridView

        Private isLoading As Boolean = True

        ' Helper interno para los ComboBox de filtro (Clientes y Vendedores)
        Private Class FiltroComboItem
            Public Property Id As Integer
            Public Property Descripcion As String

            Public Sub New(id As Integer, descripcion As String)
                Me.Id = id
                Me.Descripcion = descripcion
            End Sub

            Public Overrides Function ToString() As String
                Return Descripcion
            End Function
        End Class

        Public Sub New()
            InitializeUI()
            CargarComboClientes()
            CargarComboVendedores()
            isLoading = False
            LoadReportes()
        End Sub

        Protected Overrides Sub OnLoad(e As EventArgs)
            MyBase.OnLoad(e)
            If Not AuthService.IsManager Then
                MessageBox.Show("El acceso a reportes comerciales y ventas está reservado exclusivamente para el rol de Gerente.", "Acceso Restringido", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Me.BeginInvoke(New Action(AddressOf Me.Close))
            End If
        End Sub

        Private Sub InitializeUI()
            Me.Text = "Reportes Comerciales, Ventas por Cliente y Estadísticas"
            Me.Size = New Size(1120, 720)
            Me.MinimumSize = New Size(980, 600)
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.BackColor = UITheme.ColorBackground
            Me.Font = UITheme.FontRegular

            ' Header superior institucional
            Dim pnlHeader As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 60,
                .BackColor = UITheme.ColorSecondary,
                .Padding = New Padding(20, 15, 20, 15)
            }
            Dim lblTitle As New Label() With {
                .Text = "Reportes de ventas, clientes y rendimiento comercial",
                .Font = UITheme.FontHeading,
                .ForeColor = Color.White,
                .AutoSize = True,
                .Location = New Point(15, 15)
            }
            pnlHeader.Controls.Add(lblTitle)

            ' Toolbar de Fechas general responsiva
            Dim pnlFilter As New Panel() With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .MinimumSize = New Size(0, 52),
                .BackColor = UITheme.ColorSurface
            }
            Dim flpFilter As New FlowLayoutPanel() With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Padding = New Padding(12, 10, 12, 6),
                .WrapContents = True
            }

            Dim lblPeriodo As New Label() With {.Text = "Período:", .Font = UITheme.FontBold, .AutoSize = True, .Margin = New Padding(0, 6, 4, 4)}
            cboRapidoFechas = New ComboBox() With {
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Width = 140,
                .Margin = New Padding(0, 3, 0, 4)
            }
            UITheme.StyleComboBox(cboRapidoFechas)
            cboRapidoFechas.Items.AddRange({"Hoy", "Ayer", "Esta semana", "Este mes", "Últimos 30 días", "Últimos 90 días", "Todo el historial"})
            cboRapidoFechas.SelectedIndex = 4 ' Últimos 30 días por defecto
            AddHandler cboRapidoFechas.SelectedIndexChanged, AddressOf CboRapidoFechas_SelectedIndexChanged

            Dim lblD As New Label() With {.Text = "Desde:", .Font = UITheme.FontBold, .AutoSize = True, .Margin = New Padding(10, 6, 4, 4)}
            dtpDesde = New DateTimePicker() With {.Size = New Size(115, 28), .Format = DateTimePickerFormat.Short, .Value = DateTime.Today.AddDays(-30), .Margin = New Padding(0, 3, 0, 4)}

            Dim lblH As New Label() With {.Text = "Hasta:", .Font = UITheme.FontBold, .AutoSize = True, .Margin = New Padding(10, 6, 4, 4)}
            dtpHasta = New DateTimePicker() With {.Size = New Size(115, 28), .Format = DateTimePickerFormat.Short, .Value = DateTime.Today, .Margin = New Padding(0, 3, 0, 4)}

            Dim lblVendGen As New Label() With {.Text = "Vendedor:", .Font = UITheme.FontBold, .AutoSize = True, .Margin = New Padding(10, 6, 4, 4)}
            cboVendedorGeneral = New ComboBox() With {
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Width = 200,
                .Margin = New Padding(0, 3, 0, 4)
            }
            UITheme.StyleComboBox(cboVendedorGeneral)
            AddHandler cboVendedorGeneral.SelectedIndexChanged, Sub()
                                                                    If Not isLoading Then
                                                                        LoadVentasGenerales()
                                                                        LoadAnulaciones()
                                                                    End If
                                                                End Sub

            btnFiltrar = New Button() With {.Text = "Generar Reportes", .Height = 32, .AutoSize = True, .AutoSizeMode = AutoSizeMode.GrowAndShrink, .Padding = New Padding(12, 0, 12, 0), .Margin = New Padding(10, 1, 0, 4)}
            UITheme.StyleButton(btnFiltrar, "Primary")
            AddHandler btnFiltrar.Click, Sub() LoadReportes()

            btnDescargarPdf = New Button() With {.Text = "📄 Descargar Ticket PDF", .Height = 32, .AutoSize = True, .AutoSizeMode = AutoSizeMode.GrowAndShrink, .Padding = New Padding(12, 0, 12, 0), .Margin = New Padding(8, 1, 0, 4)}
            UITheme.StyleButton(btnDescargarPdf, "Secondary")
            AddHandler btnDescargarPdf.Click, AddressOf BtnDescargarPdf_Click

            btnAnularVenta = New Button() With {.Text = "Anular Venta Seleccionada", .Height = 32, .AutoSize = True, .AutoSizeMode = AutoSizeMode.GrowAndShrink, .Padding = New Padding(12, 0, 12, 0), .Margin = New Padding(8, 1, 0, 4)}
            UITheme.StyleButton(btnAnularVenta, "Danger")
            btnAnularVenta.Visible = AuthService.IsAdminOrManager
            AddHandler btnAnularVenta.Click, AddressOf BtnAnularVenta_Click

            flpFilter.Controls.AddRange({lblPeriodo, cboRapidoFechas, lblD, dtpDesde, lblH, dtpHasta, lblVendGen, cboVendedorGeneral, btnFiltrar, btnDescargarPdf, btnAnularVenta})
            pnlFilter.Controls.Add(flpFilter)

            ' Tarjetas KPI Financieras del período (responsivas en 4 columnas iguales)
            Dim pnlKpis As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 125,
                .BackColor = UITheme.ColorSurfaceMuted,
                .Padding = New Padding(12, 5, 12, 5)
            }

            Dim tlpKpis As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 4,
                .RowCount = 1,
                .BackColor = Color.Transparent
            }
            tlpKpis.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))
            tlpKpis.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))
            tlpKpis.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))
            tlpKpis.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))
            tlpKpis.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            Dim cardTot = UITheme.CreateKpiCard("FACTURACIÓN TOTAL", "$ 0,00", "0 tickets emitidos", UITheme.ColorPrimary)
            cardTot.Dock = DockStyle.Fill
            cardTot.Margin = New Padding(3, 3, 6, 3)
            lblTotalVentas = CType(cardTot.Controls(2), Label)
            lblTotalVentasSub = CType(cardTot.Controls(3), Label)

            Dim cardCosto = UITheme.CreateKpiCard("COSTO MERCADERÍA", "$ 0,00", "Costo de prendas vendidas", UITheme.ColorWarning)
            cardCosto.Dock = DockStyle.Fill
            cardCosto.Margin = New Padding(6, 3, 6, 3)
            lblCostoMercaderia = CType(cardCosto.Controls(2), Label)
            lblCostoMercaderiaSub = CType(cardCosto.Controls(3), Label)

            Dim cardGanancia = UITheme.CreateKpiCard("GANANCIA NETA (MARGEN)", "$ 0,00", "Rentabilidad real (0,0 %)", UITheme.ColorSuccess)
            cardGanancia.Dock = DockStyle.Fill
            cardGanancia.Margin = New Padding(6, 3, 6, 3)
            lblGananciaNeta = CType(cardGanancia.Controls(2), Label)
            lblGananciaNetaSub = CType(cardGanancia.Controls(3), Label)

            Dim cardProm = UITheme.CreateKpiCard("TICKET PROMEDIO", "$ 0,00", "Gasto promedio por compra", UITheme.ColorInfo)
            cardProm.Dock = DockStyle.Fill
            cardProm.Margin = New Padding(6, 3, 3, 3)
            lblTicketPromedio = CType(cardProm.Controls(2), Label)
            lblTicketPromedioSub = CType(cardProm.Controls(3), Label)

            tlpKpis.Controls.Add(cardTot, 0, 0)
            tlpKpis.Controls.Add(cardCosto, 1, 0)
            tlpKpis.Controls.Add(cardGanancia, 2, 0)
            tlpKpis.Controls.Add(cardProm, 3, 0)
            pnlKpis.Controls.Add(tlpKpis)

            ' TabControl con 5 secciones analíticas
            tabs = New TabControl() With {
                .Dock = DockStyle.Fill,
                .Padding = New Point(14, 7),
                .Font = UITheme.FontBold
            }

            ' Tab 1: Comprobantes y Ventas Globales (con filtros dinámicos)
            tabVentas = New TabPage("Listado General de Ventas") With {.BackColor = UITheme.ColorBackground}
            ConstruirTabVentas()
            tabs.TabPages.Add(tabVentas)

            ' Tab 2: Gestión de Anulaciones y Notas de Crédito
            tabAnulaciones = New TabPage("Anulaciones y Devoluciones") With {.BackColor = UITheme.ColorBackground}
            ConstruirTabAnulaciones()
            tabs.TabPages.Add(tabAnulaciones)

            ' Tab 3: Ventas por Cliente (con filtro y desglose)
            tabClientes = New TabPage("Ventas por Cliente y Fecha") With {.BackColor = UITheme.ColorBackground}
            ConstruirTabClientes()
            tabs.TabPages.Add(tabClientes)

            ' Tab 4: Ventas por Vendedor (con filtro y desglose)
            tabVendedores = New TabPage("Ventas por Vendedor y Fecha") With {.BackColor = UITheme.ColorBackground}
            ConstruirTabVendedores()
            tabs.TabPages.Add(tabVendedores)

            ' Tab 5: Ranking de Prendas (Mayor Salida, Mayor Margen y Stock Inmovilizado)
            tabRanking = New TabPage("Ranking de Prendas y Rentabilidad") With {.BackColor = UITheme.ColorBackground}
            ConstruirTabRanking()
            tabs.TabPages.Add(tabRanking)

            ' Jerarquía de docking
            Me.Controls.Add(tabs)
            Me.Controls.Add(pnlKpis)
            Me.Controls.Add(pnlFilter)
            Me.Controls.Add(pnlHeader)
        End Sub

        Private Sub ConstruirTabVentas()
            Dim pnlVentasToolbar As New Panel() With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .MinimumSize = New Size(0, 48),
                .BackColor = UITheme.ColorSurface
            }
            Dim flpVentas As New FlowLayoutPanel() With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Padding = New Padding(10, 10, 10, 6),
                .WrapContents = True
            }

            Dim lblCliGen As New Label() With {.Text = "Cliente:", .Font = UITheme.FontBold, .AutoSize = True, .Margin = New Padding(0, 6, 4, 4)}
            cboClienteGeneral = New ComboBox() With {
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Width = 230,
                .Margin = New Padding(0, 3, 0, 4)
            }
            UITheme.StyleComboBox(cboClienteGeneral)
            AddHandler cboClienteGeneral.SelectedIndexChanged, Sub()
                                                                   If Not isLoading Then LoadVentasGenerales()
                                                               End Sub

            Dim lblMetodoGen As New Label() With {.Text = "Medio de Pago:", .Font = UITheme.FontBold, .AutoSize = True, .Margin = New Padding(12, 6, 4, 4)}
            cboMetodoPagoGeneral = New ComboBox() With {
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Width = 175,
                .Margin = New Padding(0, 3, 0, 4)
            }
            UITheme.StyleComboBox(cboMetodoPagoGeneral)
            cboMetodoPagoGeneral.Items.AddRange({"[ Todos los medios ]", "Efectivo", "Tarjeta Débito", "Tarjeta Crédito", "Transferencia / QR", "Múltiple"})
            cboMetodoPagoGeneral.SelectedIndex = 0
            AddHandler cboMetodoPagoGeneral.SelectedIndexChanged, Sub()
                                                                      If Not isLoading Then LoadVentasGenerales()
                                                                  End Sub

            Dim lblEstGen As New Label() With {.Text = "Estado:", .Font = UITheme.FontBold, .AutoSize = True, .Margin = New Padding(12, 6, 4, 4)}
            cboEstadoGeneral = New ComboBox() With {
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Width = 160,
                .Margin = New Padding(0, 3, 0, 4)
            }
            UITheme.StyleComboBox(cboEstadoGeneral)
            cboEstadoGeneral.Items.AddRange({"[ Todos los estados ]", "Completada", "Anulada"})
            cboEstadoGeneral.SelectedIndex = 0
            AddHandler cboEstadoGeneral.SelectedIndexChanged, Sub()
                                                                  If Not isLoading Then LoadVentasGenerales()
                                                              End Sub

            Dim lblHint As New Label() With {
                .Text = "💡 Doble clic en un ticket para ver prendas o reimprimir en PDF",
                .Font = UITheme.FontSmall,
                .ForeColor = UITheme.ColorTextSecondary,
                .AutoSize = True,
                .Margin = New Padding(16, 7, 0, 4)
            }

            flpVentas.Controls.AddRange({lblCliGen, cboClienteGeneral, lblMetodoGen, cboMetodoPagoGeneral, lblEstGen, cboEstadoGeneral, lblHint})
            pnlVentasToolbar.Controls.Add(flpVentas)

            dgvVentas = New DataGridView() With {.Dock = DockStyle.Fill}
            UITheme.StyleDataGrid(dgvVentas)
            ConfigurarColumnasVentas()
            AddHandler dgvVentas.CellDoubleClick, AddressOf DgvVentas_CellDoubleClick

            tabVentas.Controls.Add(dgvVentas)
            tabVentas.Controls.Add(pnlVentasToolbar)
        End Sub

        Private Sub ConstruirTabAnulaciones()
            Dim pnlAnulacionesToolbar As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 44,
                .BackColor = UITheme.ColorSurface,
                .Padding = New Padding(12, 8, 12, 8)
            }

            lblResumenAnulaciones = New Label() With {
                .Dock = DockStyle.Fill,
                .AutoSize = False,
                .AutoEllipsis = True,
                .TextAlign = ContentAlignment.MiddleLeft,
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorDanger,
                .Text = "Cargando registro de anulaciones y devoluciones..."
            }
            pnlAnulacionesToolbar.Controls.Add(lblResumenAnulaciones)

            dgvAnulaciones = New DataGridView() With {.Dock = DockStyle.Fill}
            UITheme.StyleDataGrid(dgvAnulaciones)
            ConfigurarColumnasAnulaciones()
            AddHandler dgvAnulaciones.CellDoubleClick, AddressOf DgvAnulaciones_CellDoubleClick

            tabAnulaciones.Controls.Add(dgvAnulaciones)
            tabAnulaciones.Controls.Add(pnlAnulacionesToolbar)
        End Sub

        Private Sub ConstruirTabClientes()
            ' Panel superior de filtros para clientes (responsivo)
            Dim pnlClientesToolbar As New Panel() With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .MinimumSize = New Size(0, 48),
                .BackColor = UITheme.ColorSurface
            }
            Dim flpClientes As New FlowLayoutPanel() With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Padding = New Padding(10, 10, 10, 6),
                .WrapContents = True
            }

            Dim lblCli As New Label() With {.Text = "Cliente:", .Font = UITheme.FontBold, .AutoSize = True, .Margin = New Padding(0, 6, 4, 4)}
            cboClientesFiltro = New ComboBox() With {
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Width = 260,
                .Margin = New Padding(0, 3, 0, 4)
            }
            UITheme.StyleComboBox(cboClientesFiltro)
            AddHandler cboClientesFiltro.SelectedIndexChanged, Sub()
                                                                   If Not isLoading Then LoadVentasPorCliente()
                                                               End Sub

            Dim lblOrd As New Label() With {.Text = "Ordenar por:", .Font = UITheme.FontBold, .AutoSize = True, .Margin = New Padding(14, 6, 4, 4)}
            cboOrdenClientes = New ComboBox() With {
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Width = 220,
                .Margin = New Padding(0, 3, 0, 4)
            }
            UITheme.StyleComboBox(cboOrdenClientes)
            cboOrdenClientes.Items.AddRange({"Mayor Facturación ($)", "Mayor Salida (Prendas)", "Mayor Margen de Ganancia ($)"})
            cboOrdenClientes.SelectedIndex = 0
            AddHandler cboOrdenClientes.SelectedIndexChanged, Sub()
                                                                  If Not isLoading Then LoadVentasPorCliente()
                                                              End Sub

            lblResumenClientes = New Label() With {
                .AutoSize = True,
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorPrimaryDark,
                .Margin = New Padding(20, 7, 0, 4),
                .Text = "Cargando métricas..."
            }

            flpClientes.Controls.AddRange({lblCli, cboClientesFiltro, lblOrd, cboOrdenClientes, lblResumenClientes})
            pnlClientesToolbar.Controls.Add(flpClientes)

            ' SplitContainer: Clientes arriba, comprobantes abajo (adaptable en altura)
            Dim scClientes As New SplitContainer() With {
                .Dock = DockStyle.Fill,
                .Orientation = Orientation.Horizontal,
                .SplitterDistance = 230,
                .SplitterWidth = 6,
                .Panel1MinSize = 110,
                .Panel2MinSize = 120
            }
            AddHandler scClientes.Resize, Sub(s, ev)
                                              If scClientes.Height > 260 Then
                                                  Try
                                                      scClientes.SplitterDistance = Math.Max(120, CInt(scClientes.Height * 0.48))
                                                  Catch
                                                  End Try
                                              End If
                                          End Sub

            ' Grilla principal de Clientes
            dgvClientesReporte = New DataGridView() With {.Dock = DockStyle.Fill}
            UITheme.StyleDataGrid(dgvClientesReporte)
            ConfigurarColumnasClientesReporte()
            AddHandler dgvClientesReporte.SelectionChanged, AddressOf DgvClientesReporte_SelectionChanged
            scClientes.Panel1.Controls.Add(dgvClientesReporte)

            ' Panel inferior: detalle de tickets del cliente
            Dim pnlTicketsHeader As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 46,
                .BackColor = UITheme.ColorSurfaceMuted,
                .Padding = New Padding(12, 7, 12, 7)
            }
            lblTicketsClienteHeader = New Label() With {
                .Text = "Comprobantes emitidos del cliente seleccionado",
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorTextPrimary,
                .AutoSize = False,
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleLeft,
                .AutoEllipsis = True
            }
            btnVerDetalleTicketCliente = New Button() With {
                .Text = "Ver Detalle de Prendas del Ticket",
                .Dock = DockStyle.Right,
                .Width = 235,
                .Height = 32
            }
            UITheme.StyleButton(btnVerDetalleTicketCliente, "Secondary")
            AddHandler btnVerDetalleTicketCliente.Click, AddressOf BtnVerDetalleTicketCliente_Click

            Dim pnlSpacerTickets As New Panel() With {
                .Dock = DockStyle.Right,
                .Width = 8,
                .BackColor = Color.Transparent
            }

            btnDescargarPdfTicketCliente = New Button() With {
                .Text = "📄 Descargar Ticket PDF",
                .Dock = DockStyle.Right,
                .Width = 195,
                .Height = 32
            }
            UITheme.StyleButton(btnDescargarPdfTicketCliente, "Primary")
            AddHandler btnDescargarPdfTicketCliente.Click, AddressOf BtnDescargarPdfTicketCliente_Click

            pnlTicketsHeader.Controls.Add(lblTicketsClienteHeader)
            pnlTicketsHeader.Controls.Add(btnDescargarPdfTicketCliente)
            pnlTicketsHeader.Controls.Add(pnlSpacerTickets)
            pnlTicketsHeader.Controls.Add(btnVerDetalleTicketCliente)

            dgvTicketsCliente = New DataGridView() With {.Dock = DockStyle.Fill}
            UITheme.StyleDataGrid(dgvTicketsCliente)
            ConfigurarColumnasTicketsCliente()
            AddHandler dgvTicketsCliente.CellDoubleClick, AddressOf DgvTicketsCliente_CellDoubleClick

            scClientes.Panel2.Controls.Add(dgvTicketsCliente)
            scClientes.Panel2.Controls.Add(pnlTicketsHeader)

            tabClientes.Controls.Add(scClientes)
            tabClientes.Controls.Add(pnlClientesToolbar)
        End Sub

        Private Sub ConstruirTabVendedores()
            ' Panel superior de filtros para vendedores (responsivo)
            Dim pnlVendedoresToolbar As New Panel() With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .MinimumSize = New Size(0, 48),
                .BackColor = UITheme.ColorSurface
            }
            Dim flpVendedores As New FlowLayoutPanel() With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Padding = New Padding(10, 10, 10, 6),
                .WrapContents = True
            }

            Dim lblVend As New Label() With {.Text = "Vendedor:", .Font = UITheme.FontBold, .AutoSize = True, .Margin = New Padding(0, 6, 4, 4)}
            cboVendedoresFiltro = New ComboBox() With {
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Width = 260,
                .Margin = New Padding(0, 3, 0, 4)
            }
            UITheme.StyleComboBox(cboVendedoresFiltro)
            AddHandler cboVendedoresFiltro.SelectedIndexChanged, Sub()
                                                                     If Not isLoading Then LoadVentasPorVendedor()
                                                                 End Sub

            Dim lblOrdVend As New Label() With {.Text = "Ordenar por:", .Font = UITheme.FontBold, .AutoSize = True, .Margin = New Padding(14, 6, 4, 4)}
            cboOrdenVendedores = New ComboBox() With {
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Width = 220,
                .Margin = New Padding(0, 3, 0, 4)
            }
            UITheme.StyleComboBox(cboOrdenVendedores)
            cboOrdenVendedores.Items.AddRange({"Mayor Facturación ($)", "Mayor Salida (Prendas)", "Mayor Margen de Ganancia ($)"})
            cboOrdenVendedores.SelectedIndex = 0
            AddHandler cboOrdenVendedores.SelectedIndexChanged, Sub()
                                                                    If Not isLoading Then LoadVentasPorVendedor()
                                                                End Sub

            lblResumenVendedores = New Label() With {
                .AutoSize = True,
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorPrimaryDark,
                .Margin = New Padding(20, 7, 0, 4),
                .Text = "Cargando métricas..."
            }

            flpVendedores.Controls.AddRange({lblVend, cboVendedoresFiltro, lblOrdVend, cboOrdenVendedores, lblResumenVendedores})
            pnlVendedoresToolbar.Controls.Add(flpVendedores)

            ' SplitContainer: Vendedores arriba, comprobantes abajo (adaptable en altura)
            Dim scVendedores As New SplitContainer() With {
                .Dock = DockStyle.Fill,
                .Orientation = Orientation.Horizontal,
                .SplitterDistance = 230,
                .SplitterWidth = 6,
                .Panel1MinSize = 110,
                .Panel2MinSize = 120
            }
            AddHandler scVendedores.Resize, Sub(s, ev)
                                                If scVendedores.Height > 260 Then
                                                    Try
                                                        scVendedores.SplitterDistance = Math.Max(120, CInt(scVendedores.Height * 0.48))
                                                    Catch
                                                    End Try
                                                End If
                                            End Sub

            ' Grilla principal de Vendedores
            dgvVendedoresReporte = New DataGridView() With {.Dock = DockStyle.Fill}
            UITheme.StyleDataGrid(dgvVendedoresReporte)
            ConfigurarColumnasVendedoresReporte()
            AddHandler dgvVendedoresReporte.SelectionChanged, AddressOf DgvVendedoresReporte_SelectionChanged
            scVendedores.Panel1.Controls.Add(dgvVendedoresReporte)

            ' Panel inferior: detalle de tickets del vendedor
            Dim pnlTicketsVendHeader As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 46,
                .BackColor = UITheme.ColorSurfaceMuted,
                .Padding = New Padding(12, 7, 12, 7)
            }
            lblTicketsVendedorHeader = New Label() With {
                .Text = "Comprobantes emitidos por el vendedor seleccionado",
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorTextPrimary,
                .AutoSize = False,
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleLeft,
                .AutoEllipsis = True
            }
            btnVerDetalleTicketVendedor = New Button() With {
                .Text = "Ver Detalle de Prendas del Ticket",
                .Dock = DockStyle.Right,
                .Width = 235,
                .Height = 32
            }
            UITheme.StyleButton(btnVerDetalleTicketVendedor, "Secondary")
            AddHandler btnVerDetalleTicketVendedor.Click, AddressOf BtnVerDetalleTicketVendedor_Click

            Dim pnlSpacerTicketsVend As New Panel() With {
                .Dock = DockStyle.Right,
                .Width = 8,
                .BackColor = Color.Transparent
            }

            btnDescargarPdfTicketVendedor = New Button() With {
                .Text = "📄 Descargar Ticket PDF",
                .Dock = DockStyle.Right,
                .Width = 195,
                .Height = 32
            }
            UITheme.StyleButton(btnDescargarPdfTicketVendedor, "Primary")
            AddHandler btnDescargarPdfTicketVendedor.Click, AddressOf BtnDescargarPdfTicketVendedor_Click

            pnlTicketsVendHeader.Controls.Add(lblTicketsVendedorHeader)
            pnlTicketsVendHeader.Controls.Add(btnDescargarPdfTicketVendedor)
            pnlTicketsVendHeader.Controls.Add(pnlSpacerTicketsVend)
            pnlTicketsVendHeader.Controls.Add(btnVerDetalleTicketVendedor)

            dgvTicketsVendedor = New DataGridView() With {.Dock = DockStyle.Fill}
            UITheme.StyleDataGrid(dgvTicketsVendedor)
            ConfigurarColumnasTicketsVendedor()
            AddHandler dgvTicketsVendedor.CellDoubleClick, AddressOf DgvTicketsVendedor_CellDoubleClick

            scVendedores.Panel2.Controls.Add(dgvTicketsVendedor)
            scVendedores.Panel2.Controls.Add(pnlTicketsVendHeader)

            tabVendedores.Controls.Add(scVendedores)
            tabVendedores.Controls.Add(pnlVendedoresToolbar)
        End Sub

        Private Sub ConstruirTabRanking()
            Dim pnlRankingToolbar As New Panel() With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .MinimumSize = New Size(0, 48),
                .BackColor = UITheme.ColorSurface
            }
            Dim flpRanking As New FlowLayoutPanel() With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Padding = New Padding(12, 10, 10, 6),
                .WrapContents = True
            }

            Dim lblCrit As New Label() With {.Text = "Criterio:", .Font = UITheme.FontBold, .AutoSize = True, .Margin = New Padding(0, 5, 8, 0)}

            rbMayorSalida = New RadioButton() With {
                .Text = "Mayor Salida (Unidades)",
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorPrimaryDark,
                .AutoSize = True,
                .Checked = True,
                .Margin = New Padding(0, 3, 12, 0)
            }
            AddHandler rbMayorSalida.CheckedChanged, Sub()
                                                         If Not isLoading AndAlso rbMayorSalida.Checked Then LoadRanking()
                                                     End Sub

            rbMayorMargen = New RadioButton() With {
                .Text = "Mayor Margen (Rentabilidad $)",
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorSuccessDark,
                .AutoSize = True,
                .Margin = New Padding(0, 3, 12, 0)
            }
            AddHandler rbMayorMargen.CheckedChanged, Sub()
                                                         If Not isLoading AndAlso rbMayorMargen.Checked Then LoadRanking()
                                                     End Sub

            rbStockInmovilizado = New RadioButton() With {
                .Text = "Stock Inmovilizado (Sin rotación)",
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorWarning,
                .AutoSize = True,
                .Margin = New Padding(0, 3, 12, 0)
            }
            AddHandler rbStockInmovilizado.CheckedChanged, Sub()
                                                               If Not isLoading AndAlso rbStockInmovilizado.Checked Then LoadRanking()
                                                           End Sub

            Dim lblLim As New Label() With {.Text = "Mostrar:", .Font = UITheme.FontBold, .AutoSize = True, .Margin = New Padding(8, 5, 4, 0)}
            cboLimiteRanking = New ComboBox() With {
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Width = 110,
                .Margin = New Padding(0, 2, 0, 0)
            }
            UITheme.StyleComboBox(cboLimiteRanking)
            cboLimiteRanking.Items.AddRange({"Top 10", "Top 20", "Top 50", "Todos"})
            cboLimiteRanking.SelectedIndex = 1
            AddHandler cboLimiteRanking.SelectedIndexChanged, Sub()
                                                                  If Not isLoading Then LoadRanking()
                                                              End Sub

            lblResumenRanking = New Label() With {
                .AutoSize = True,
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorTextPrimary,
                .Margin = New Padding(14, 5, 0, 0),
                .Text = ""
            }

            flpRanking.Controls.AddRange({lblCrit, rbMayorSalida, rbMayorMargen, rbStockInmovilizado, lblLim, cboLimiteRanking, lblResumenRanking})
            pnlRankingToolbar.Controls.Add(flpRanking)

            dgvRanking = New DataGridView() With {.Dock = DockStyle.Fill}
            UITheme.StyleDataGrid(dgvRanking)
            ConfigurarColumnasRanking()

            tabRanking.Controls.Add(dgvRanking)
            tabRanking.Controls.Add(pnlRankingToolbar)
        End Sub

        Private Sub CargarComboClientes()
            cboClienteGeneral.Items.Clear()
            cboClienteGeneral.Items.Add(New FiltroComboItem(0, "[ Todos los clientes ]"))

            cboClientesFiltro.Items.Clear()
            cboClientesFiltro.Items.Add(New FiltroComboItem(0, "[ Todos los clientes ]"))

            Dim clientes = clienteService.GetClientes("", True)
            For Each c In clientes
                Dim desc = $"{c.Apellido}, {c.Nombre} (DNI: {c.DniCuit})"
                cboClienteGeneral.Items.Add(New FiltroComboItem(c.Id, desc))
                cboClientesFiltro.Items.Add(New FiltroComboItem(c.Id, desc))
            Next

            If cboClienteGeneral.Items.Count > 0 Then
                cboClienteGeneral.SelectedIndex = 0
            End If
            If cboClientesFiltro.Items.Count > 0 Then
                cboClientesFiltro.SelectedIndex = 0
            End If
        End Sub

        Private Sub CargarComboVendedores()
            cboVendedorGeneral.Items.Clear()
            cboVendedorGeneral.Items.Add(New FiltroComboItem(0, "[ Todos los vendedores ]"))

            cboVendedoresFiltro.Items.Clear()
            cboVendedoresFiltro.Items.Add(New FiltroComboItem(0, "[ Todos los vendedores ]"))

            Dim usuarios = usuarioService.GetUsuarios(soloActivos:=False)
            For Each u In usuarios
                If u.Rol.Equals("Vendedor", StringComparison.OrdinalIgnoreCase) OrElse u.Rol.Equals("Gerente", StringComparison.OrdinalIgnoreCase) Then
                    Dim desc = $"{u.NombreCompleto} ({u.Username})"
                    cboVendedorGeneral.Items.Add(New FiltroComboItem(u.Id, desc))
                    cboVendedoresFiltro.Items.Add(New FiltroComboItem(u.Id, desc))
                End If
            Next

            If cboVendedorGeneral.Items.Count > 0 Then
                cboVendedorGeneral.SelectedIndex = 0
            End If
            If cboVendedoresFiltro.Items.Count > 0 Then
                cboVendedoresFiltro.SelectedIndex = 0
            End If
        End Sub

        Private Sub ConfigurarColumnasVentas()
            dgvVentas.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            dgvVentas.Columns.Clear()
            dgvVentas.Columns.Add("Id", "ID")
            dgvVentas.Columns("Id").Visible = False

            dgvVentas.Columns.Add("Ticket", "Nº Ticket")
            dgvVentas.Columns("Ticket").FillWeight = 13
            dgvVentas.Columns("Ticket").MinimumWidth = 110

            dgvVentas.Columns.Add("Fecha", "Fecha y Hora")
            dgvVentas.Columns("Fecha").FillWeight = 13
            dgvVentas.Columns("Fecha").MinimumWidth = 115

            dgvVentas.Columns.Add("Vendedor", "Vendedor")
            dgvVentas.Columns("Vendedor").FillWeight = 14
            dgvVentas.Columns("Vendedor").MinimumWidth = 100

            dgvVentas.Columns.Add("Cliente", "Cliente")
            dgvVentas.Columns("Cliente").FillWeight = 18
            dgvVentas.Columns("Cliente").MinimumWidth = 120

            dgvVentas.Columns.Add("Metodo", "Medio Pago")
            dgvVentas.Columns("Metodo").FillWeight = 12
            dgvVentas.Columns("Metodo").MinimumWidth = 95

            dgvVentas.Columns.Add("Subtotal", "Subtotal")
            dgvVentas.Columns("Subtotal").FillWeight = 10
            dgvVentas.Columns("Subtotal").MinimumWidth = 85
            dgvVentas.Columns("Subtotal").DefaultCellStyle.Format = "C2"

            dgvVentas.Columns.Add("Descuento", "Desc. ($)")
            dgvVentas.Columns("Descuento").FillWeight = 9
            dgvVentas.Columns("Descuento").MinimumWidth = 75
            dgvVentas.Columns("Descuento").DefaultCellStyle.Format = "C2"

            dgvVentas.Columns.Add("Total", "Total Facturado")
            dgvVentas.Columns("Total").FillWeight = 12
            dgvVentas.Columns("Total").MinimumWidth = 100
            dgvVentas.Columns("Total").DefaultCellStyle.Format = "C2"
            dgvVentas.Columns("Total").DefaultCellStyle.Font = UITheme.FontBold

            dgvVentas.Columns.Add("Estado", "Estado")
            dgvVentas.Columns("Estado").FillWeight = 9
            dgvVentas.Columns("Estado").MinimumWidth = 80
        End Sub

        Private Sub ConfigurarColumnasClientesReporte()
            dgvClientesReporte.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            dgvClientesReporte.Columns.Clear()
            dgvClientesReporte.Columns.Add("Id", "ID")
            dgvClientesReporte.Columns("Id").Visible = False

            dgvClientesReporte.Columns.Add("Cliente", "Cliente")
            dgvClientesReporte.Columns("Cliente").FillWeight = 20
            dgvClientesReporte.Columns("Cliente").MinimumWidth = 130

            dgvClientesReporte.Columns.Add("Dni", "DNI / CUIT")
            dgvClientesReporte.Columns("Dni").FillWeight = 11
            dgvClientesReporte.Columns("Dni").MinimumWidth = 90

            dgvClientesReporte.Columns.Add("Telefono", "Teléfono")
            dgvClientesReporte.Columns("Telefono").FillWeight = 11
            dgvClientesReporte.Columns("Telefono").MinimumWidth = 90

            dgvClientesReporte.Columns.Add("Compras", "Tickets")
            dgvClientesReporte.Columns("Compras").FillWeight = 8
            dgvClientesReporte.Columns("Compras").MinimumWidth = 65
            dgvClientesReporte.Columns("Compras").DefaultCellStyle.Font = UITheme.FontBold
            dgvClientesReporte.Columns("Compras").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            dgvClientesReporte.Columns.Add("Unidades", "Prendas Adquiridas")
            dgvClientesReporte.Columns("Unidades").FillWeight = 12
            dgvClientesReporte.Columns("Unidades").MinimumWidth = 95
            dgvClientesReporte.Columns("Unidades").DefaultCellStyle.Font = UITheme.FontBold
            dgvClientesReporte.Columns("Unidades").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            dgvClientesReporte.Columns.Add("Total", "Total Facturado")
            dgvClientesReporte.Columns("Total").FillWeight = 13
            dgvClientesReporte.Columns("Total").MinimumWidth = 105
            dgvClientesReporte.Columns("Total").DefaultCellStyle.Format = "C2"
            dgvClientesReporte.Columns("Total").DefaultCellStyle.Font = UITheme.FontBold

            dgvClientesReporte.Columns.Add("Margen", "Margen Ganancia")
            dgvClientesReporte.Columns("Margen").FillWeight = 13
            dgvClientesReporte.Columns("Margen").MinimumWidth = 105
            dgvClientesReporte.Columns("Margen").DefaultCellStyle.Format = "C2"
            dgvClientesReporte.Columns("Margen").DefaultCellStyle.Font = UITheme.FontBold
            dgvClientesReporte.Columns("Margen").DefaultCellStyle.ForeColor = UITheme.ColorSuccessDark

            dgvClientesReporte.Columns.Add("Rentabilidad", "% Margen")
            dgvClientesReporte.Columns("Rentabilidad").FillWeight = 9
            dgvClientesReporte.Columns("Rentabilidad").MinimumWidth = 75
            dgvClientesReporte.Columns("Rentabilidad").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            dgvClientesReporte.Columns.Add("UltimaCompra", "Última Compra")
            dgvClientesReporte.Columns("UltimaCompra").FillWeight = 12
            dgvClientesReporte.Columns("UltimaCompra").MinimumWidth = 105
        End Sub

        Private Sub ConfigurarColumnasTicketsCliente()
            dgvTicketsCliente.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            dgvTicketsCliente.Columns.Clear()
            dgvTicketsCliente.Columns.Add("Id", "ID")
            dgvTicketsCliente.Columns("Id").Visible = False

            dgvTicketsCliente.Columns.Add("Ticket", "Nº Ticket")
            dgvTicketsCliente.Columns("Ticket").FillWeight = 15
            dgvTicketsCliente.Columns("Ticket").MinimumWidth = 110

            dgvTicketsCliente.Columns.Add("Fecha", "Fecha y Hora")
            dgvTicketsCliente.Columns("Fecha").FillWeight = 15
            dgvTicketsCliente.Columns("Fecha").MinimumWidth = 115

            dgvTicketsCliente.Columns.Add("Vendedor", "Vendedor")
            dgvTicketsCliente.Columns("Vendedor").FillWeight = 18
            dgvTicketsCliente.Columns("Vendedor").MinimumWidth = 110

            dgvTicketsCliente.Columns.Add("Metodo", "Medio Pago")
            dgvTicketsCliente.Columns("Metodo").FillWeight = 13
            dgvTicketsCliente.Columns("Metodo").MinimumWidth = 95

            dgvTicketsCliente.Columns.Add("Subtotal", "Subtotal")
            dgvTicketsCliente.Columns("Subtotal").FillWeight = 12
            dgvTicketsCliente.Columns("Subtotal").MinimumWidth = 85
            dgvTicketsCliente.Columns("Subtotal").DefaultCellStyle.Format = "C2"

            dgvTicketsCliente.Columns.Add("Descuento", "Desc. ($)")
            dgvTicketsCliente.Columns("Descuento").FillWeight = 10
            dgvTicketsCliente.Columns("Descuento").MinimumWidth = 75
            dgvTicketsCliente.Columns("Descuento").DefaultCellStyle.Format = "C2"

            dgvTicketsCliente.Columns.Add("Total", "Total Facturado")
            dgvTicketsCliente.Columns("Total").FillWeight = 14
            dgvTicketsCliente.Columns("Total").MinimumWidth = 100
            dgvTicketsCliente.Columns("Total").DefaultCellStyle.Format = "C2"
            dgvTicketsCliente.Columns("Total").DefaultCellStyle.Font = UITheme.FontBold

            dgvTicketsCliente.Columns.Add("Estado", "Estado")
            dgvTicketsCliente.Columns("Estado").FillWeight = 10
            dgvTicketsCliente.Columns("Estado").MinimumWidth = 80
        End Sub

        Private Sub ConfigurarColumnasVendedoresReporte()
            dgvVendedoresReporte.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            dgvVendedoresReporte.Columns.Clear()
            dgvVendedoresReporte.Columns.Add("Id", "ID")
            dgvVendedoresReporte.Columns("Id").Visible = False

            dgvVendedoresReporte.Columns.Add("Vendedor", "Vendedor")
            dgvVendedoresReporte.Columns("Vendedor").FillWeight = 20
            dgvVendedoresReporte.Columns("Vendedor").MinimumWidth = 130

            dgvVendedoresReporte.Columns.Add("Username", "Usuario")
            dgvVendedoresReporte.Columns("Username").FillWeight = 11
            dgvVendedoresReporte.Columns("Username").MinimumWidth = 85

            dgvVendedoresReporte.Columns.Add("Dni", "DNI")
            dgvVendedoresReporte.Columns("Dni").FillWeight = 11
            dgvVendedoresReporte.Columns("Dni").MinimumWidth = 85

            dgvVendedoresReporte.Columns.Add("Ventas", "Tickets")
            dgvVendedoresReporte.Columns("Ventas").FillWeight = 8
            dgvVendedoresReporte.Columns("Ventas").MinimumWidth = 65
            dgvVendedoresReporte.Columns("Ventas").DefaultCellStyle.Font = UITheme.FontBold
            dgvVendedoresReporte.Columns("Ventas").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            dgvVendedoresReporte.Columns.Add("Unidades", "Prendas Vendidas")
            dgvVendedoresReporte.Columns("Unidades").FillWeight = 12
            dgvVendedoresReporte.Columns("Unidades").MinimumWidth = 95
            dgvVendedoresReporte.Columns("Unidades").DefaultCellStyle.Font = UITheme.FontBold
            dgvVendedoresReporte.Columns("Unidades").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            dgvVendedoresReporte.Columns.Add("Total", "Total Facturado")
            dgvVendedoresReporte.Columns("Total").FillWeight = 13
            dgvVendedoresReporte.Columns("Total").MinimumWidth = 105
            dgvVendedoresReporte.Columns("Total").DefaultCellStyle.Format = "C2"
            dgvVendedoresReporte.Columns("Total").DefaultCellStyle.Font = UITheme.FontBold

            dgvVendedoresReporte.Columns.Add("Margen", "Margen Ganancia")
            dgvVendedoresReporte.Columns("Margen").FillWeight = 13
            dgvVendedoresReporte.Columns("Margen").MinimumWidth = 105
            dgvVendedoresReporte.Columns("Margen").DefaultCellStyle.Format = "C2"
            dgvVendedoresReporte.Columns("Margen").DefaultCellStyle.Font = UITheme.FontBold
            dgvVendedoresReporte.Columns("Margen").DefaultCellStyle.ForeColor = UITheme.ColorSuccessDark

            dgvVendedoresReporte.Columns.Add("Rentabilidad", "% Margen")
            dgvVendedoresReporte.Columns("Rentabilidad").FillWeight = 9
            dgvVendedoresReporte.Columns("Rentabilidad").MinimumWidth = 75
            dgvVendedoresReporte.Columns("Rentabilidad").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            dgvVendedoresReporte.Columns.Add("UltimaVenta", "Última Venta")
            dgvVendedoresReporte.Columns("UltimaVenta").FillWeight = 12
            dgvVendedoresReporte.Columns("UltimaVenta").MinimumWidth = 105
        End Sub

        Private Sub ConfigurarColumnasTicketsVendedor()
            dgvTicketsVendedor.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            dgvTicketsVendedor.Columns.Clear()
            dgvTicketsVendedor.Columns.Add("Id", "ID")
            dgvTicketsVendedor.Columns("Id").Visible = False

            dgvTicketsVendedor.Columns.Add("Ticket", "Nº Ticket")
            dgvTicketsVendedor.Columns("Ticket").FillWeight = 15
            dgvTicketsVendedor.Columns("Ticket").MinimumWidth = 110

            dgvTicketsVendedor.Columns.Add("Fecha", "Fecha y Hora")
            dgvTicketsVendedor.Columns("Fecha").FillWeight = 15
            dgvTicketsVendedor.Columns("Fecha").MinimumWidth = 115

            dgvTicketsVendedor.Columns.Add("Cliente", "Cliente")
            dgvTicketsVendedor.Columns("Cliente").FillWeight = 18
            dgvTicketsVendedor.Columns("Cliente").MinimumWidth = 115

            dgvTicketsVendedor.Columns.Add("Metodo", "Medio Pago")
            dgvTicketsVendedor.Columns("Metodo").FillWeight = 13
            dgvTicketsVendedor.Columns("Metodo").MinimumWidth = 95

            dgvTicketsVendedor.Columns.Add("Subtotal", "Subtotal")
            dgvTicketsVendedor.Columns("Subtotal").FillWeight = 12
            dgvTicketsVendedor.Columns("Subtotal").MinimumWidth = 85
            dgvTicketsVendedor.Columns("Subtotal").DefaultCellStyle.Format = "C2"

            dgvTicketsVendedor.Columns.Add("Descuento", "Desc. ($)")
            dgvTicketsVendedor.Columns("Descuento").FillWeight = 10
            dgvTicketsVendedor.Columns("Descuento").MinimumWidth = 75
            dgvTicketsVendedor.Columns("Descuento").DefaultCellStyle.Format = "C2"

            dgvTicketsVendedor.Columns.Add("Total", "Total Facturado")
            dgvTicketsVendedor.Columns("Total").FillWeight = 14
            dgvTicketsVendedor.Columns("Total").MinimumWidth = 100
            dgvTicketsVendedor.Columns("Total").DefaultCellStyle.Format = "C2"
            dgvTicketsVendedor.Columns("Total").DefaultCellStyle.Font = UITheme.FontBold

            dgvTicketsVendedor.Columns.Add("Estado", "Estado")
            dgvTicketsVendedor.Columns("Estado").FillWeight = 10
            dgvTicketsVendedor.Columns("Estado").MinimumWidth = 80
        End Sub

        Private Sub ConfigurarColumnasRanking()
            dgvRanking.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            dgvRanking.Columns.Clear()
            dgvRanking.Columns.Add("Prenda", "Prenda / Modelo")
            dgvRanking.Columns("Prenda").FillWeight = 32
            dgvRanking.Columns("Prenda").MinimumWidth = 160

            dgvRanking.Columns.Add("Talle", "Talle")
            dgvRanking.Columns("Talle").FillWeight = 9
            dgvRanking.Columns("Talle").MinimumWidth = 65
            dgvRanking.Columns("Talle").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            dgvRanking.Columns.Add("Cantidad", "Unidades Vendidas")
            dgvRanking.Columns("Cantidad").FillWeight = 13
            dgvRanking.Columns("Cantidad").MinimumWidth = 100
            dgvRanking.Columns("Cantidad").DefaultCellStyle.Font = UITheme.FontBold
            dgvRanking.Columns("Cantidad").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            dgvRanking.Columns.Add("Recaudado", "Recaudación Total")
            dgvRanking.Columns("Recaudado").FillWeight = 14
            dgvRanking.Columns("Recaudado").MinimumWidth = 110
            dgvRanking.Columns("Recaudado").DefaultCellStyle.Format = "C2"
            dgvRanking.Columns("Recaudado").DefaultCellStyle.Font = UITheme.FontBold

            dgvRanking.Columns.Add("Costo", "Costo Total")
            dgvRanking.Columns("Costo").FillWeight = 12
            dgvRanking.Columns("Costo").MinimumWidth = 95
            dgvRanking.Columns("Costo").DefaultCellStyle.Format = "C2"

            dgvRanking.Columns.Add("Margen", "Margen Ganancia")
            dgvRanking.Columns("Margen").FillWeight = 14
            dgvRanking.Columns("Margen").MinimumWidth = 110
            dgvRanking.Columns("Margen").DefaultCellStyle.Format = "C2"
            dgvRanking.Columns("Margen").DefaultCellStyle.Font = UITheme.FontBold
            dgvRanking.Columns("Margen").DefaultCellStyle.ForeColor = UITheme.ColorSuccessDark

            dgvRanking.Columns.Add("PctMargen", "% Rentabilidad")
            dgvRanking.Columns("PctMargen").FillWeight = 10
            dgvRanking.Columns("PctMargen").MinimumWidth = 85
            dgvRanking.Columns("PctMargen").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            If AuthService.IsVendor Then
                dgvRanking.Columns("Costo").Visible = False
                dgvRanking.Columns("Margen").Visible = False
                dgvRanking.Columns("PctMargen").Visible = False
            End If
        End Sub

        Private Sub ConfigurarColumnasStockInmovilizado()
            dgvRanking.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            dgvRanking.Columns.Clear()

            dgvRanking.Columns.Add("Codigo", "Código")
            dgvRanking.Columns("Codigo").FillWeight = 12
            dgvRanking.Columns("Codigo").MinimumWidth = 90

            dgvRanking.Columns.Add("Prenda", "Prenda / Modelo")
            dgvRanking.Columns("Prenda").FillWeight = 26
            dgvRanking.Columns("Prenda").MinimumWidth = 140

            dgvRanking.Columns.Add("Talle", "Talle")
            dgvRanking.Columns("Talle").FillWeight = 8
            dgvRanking.Columns("Talle").MinimumWidth = 60
            dgvRanking.Columns("Talle").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            dgvRanking.Columns.Add("Color", "Color")
            dgvRanking.Columns("Color").FillWeight = 10
            dgvRanking.Columns("Color").MinimumWidth = 75
            dgvRanking.Columns("Color").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            dgvRanking.Columns.Add("StockActual", "Stock Inmovilizado")
            dgvRanking.Columns("StockActual").FillWeight = 11
            dgvRanking.Columns("StockActual").MinimumWidth = 90
            dgvRanking.Columns("StockActual").DefaultCellStyle.Font = UITheme.FontBold
            dgvRanking.Columns("StockActual").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            dgvRanking.Columns.Add("CostoInmovilizado", "Capital en Costo ($)")
            dgvRanking.Columns("CostoInmovilizado").FillWeight = 13
            dgvRanking.Columns("CostoInmovilizado").MinimumWidth = 105
            dgvRanking.Columns("CostoInmovilizado").DefaultCellStyle.Format = "C2"
            dgvRanking.Columns("CostoInmovilizado").DefaultCellStyle.Font = UITheme.FontBold
            dgvRanking.Columns("CostoInmovilizado").DefaultCellStyle.ForeColor = UITheme.ColorWarning

            dgvRanking.Columns.Add("ValorVenta", "Valor Venta ($)")
            dgvRanking.Columns("ValorVenta").FillWeight = 13
            dgvRanking.Columns("ValorVenta").MinimumWidth = 105
            dgvRanking.Columns("ValorVenta").DefaultCellStyle.Format = "C2"

            dgvRanking.Columns.Add("UltimaVenta", "Última Venta Histórica")
            dgvRanking.Columns("UltimaVenta").FillWeight = 17
            dgvRanking.Columns("UltimaVenta").MinimumWidth = 130
        End Sub

        Private Sub ConfigurarColumnasAnulaciones()
            dgvAnulaciones.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            dgvAnulaciones.Columns.Clear()
            dgvAnulaciones.Columns.Add("Id", "ID")
            dgvAnulaciones.Columns("Id").Visible = False

            dgvAnulaciones.Columns.Add("Ticket", "Nº Ticket")
            dgvAnulaciones.Columns("Ticket").FillWeight = 13
            dgvAnulaciones.Columns("Ticket").MinimumWidth = 105

            dgvAnulaciones.Columns.Add("Fecha", "Fecha y Hora")
            dgvAnulaciones.Columns("Fecha").FillWeight = 13
            dgvAnulaciones.Columns("Fecha").MinimumWidth = 110

            dgvAnulaciones.Columns.Add("Cliente", "Cliente")
            dgvAnulaciones.Columns("Cliente").FillWeight = 15
            dgvAnulaciones.Columns("Cliente").MinimumWidth = 110

            dgvAnulaciones.Columns.Add("Vendedor", "Vendedor Original")
            dgvAnulaciones.Columns("Vendedor").FillWeight = 14
            dgvAnulaciones.Columns("Vendedor").MinimumWidth = 105

            dgvAnulaciones.Columns.Add("AutorizadoPor", "Autorizado Por")
            dgvAnulaciones.Columns("AutorizadoPor").FillWeight = 14
            dgvAnulaciones.Columns("AutorizadoPor").MinimumWidth = 105
            dgvAnulaciones.Columns("AutorizadoPor").DefaultCellStyle.Font = UITheme.FontBold

            dgvAnulaciones.Columns.Add("Metodo", "Medio Pago")
            dgvAnulaciones.Columns("Metodo").FillWeight = 11
            dgvAnulaciones.Columns("Metodo").MinimumWidth = 90

            dgvAnulaciones.Columns.Add("Total", "Monto Revertido")
            dgvAnulaciones.Columns("Total").FillWeight = 12
            dgvAnulaciones.Columns("Total").MinimumWidth = 100
            dgvAnulaciones.Columns("Total").DefaultCellStyle.Format = "C2"
            dgvAnulaciones.Columns("Total").DefaultCellStyle.Font = UITheme.FontBold
            dgvAnulaciones.Columns("Total").DefaultCellStyle.ForeColor = UITheme.ColorDanger

            dgvAnulaciones.Columns.Add("Motivo", "Motivo de Anulación / Devolución")
            dgvAnulaciones.Columns("Motivo").FillWeight = 22
            dgvAnulaciones.Columns("Motivo").MinimumWidth = 150
        End Sub

        Private Sub CboRapidoFechas_SelectedIndexChanged(sender As Object, e As EventArgs)
            If isLoading Then Return
            Dim hoy = DateTime.Today

            Select Case cboRapidoFechas.SelectedIndex
                Case 0 ' Hoy
                    dtpDesde.Value = hoy
                    dtpHasta.Value = hoy
                Case 1 ' Ayer
                    dtpDesde.Value = hoy.AddDays(-1)
                    dtpHasta.Value = hoy.AddDays(-1)
                Case 2 ' Esta semana
                    Dim diaSemana = CInt(hoy.DayOfWeek)
                    Dim delta = If(diaSemana = 0, 6, diaSemana - 1)
                    dtpDesde.Value = hoy.AddDays(-delta)
                    dtpHasta.Value = hoy
                Case 3 ' Este mes
                    dtpDesde.Value = New DateTime(hoy.Year, hoy.Month, 1)
                    dtpHasta.Value = hoy
                Case 4 ' Últimos 30 días
                    dtpDesde.Value = hoy.AddDays(-30)
                    dtpHasta.Value = hoy
                Case 5 ' Últimos 90 días
                    dtpDesde.Value = hoy.AddDays(-90)
                    dtpHasta.Value = hoy
                Case 6 ' Todo el historial
                    dtpDesde.Value = New DateTime(2020, 1, 1)
                    dtpHasta.Value = hoy
            End Select
            LoadReportes()
        End Sub

        ''' <summary>
        ''' Carga sincronizada de todas las pestañas de reporte según el rango de fechas.
        ''' </summary>
        Private Sub LoadReportes()
            LoadVentasGenerales()
            LoadAnulaciones()
            LoadVentasPorCliente()
            LoadVentasPorVendedor()
            LoadRanking()
        End Sub

        Private Sub LoadVentasGenerales()
            Dim desde = dtpDesde.Value.Date
            Dim hasta = dtpHasta.Value.Date

            Dim vendedorId As Integer = 0
            If cboVendedorGeneral IsNot Nothing AndAlso cboVendedorGeneral.SelectedItem IsNot Nothing Then
                Dim item = CType(cboVendedorGeneral.SelectedItem, FiltroComboItem)
                vendedorId = item.Id
            End If

            Dim clienteId As Integer = 0
            If cboClienteGeneral IsNot Nothing AndAlso cboClienteGeneral.SelectedItem IsNot Nothing Then
                Dim itemCli = CType(cboClienteGeneral.SelectedItem, FiltroComboItem)
                clienteId = itemCli.Id
            End If

            Dim metodoFiltro As String = ""
            If cboMetodoPagoGeneral IsNot Nothing AndAlso cboMetodoPagoGeneral.SelectedIndex > 0 Then
                metodoFiltro = cboMetodoPagoGeneral.SelectedItem.ToString()
            End If

            Dim estadoFiltro As String = ""
            If cboEstadoGeneral IsNot Nothing AndAlso cboEstadoGeneral.SelectedIndex > 0 Then
                estadoFiltro = cboEstadoGeneral.SelectedItem.ToString()
            End If

            Dim ventas = ventaService.GetVentas(desde, hasta, "", clienteId, vendedorId, metodoFiltro, estadoFiltro)
            dgvVentas.Rows.Clear()

            For Each v In ventas
                Dim rIdx = dgvVentas.Rows.Add(v.Id, v.NumeroTicket, v.Fecha.ToString("dd/MM/yyyy HH:mm"), v.UsuarioNombre, v.ClienteNombre, v.MetodoPago, v.Subtotal, v.DescuentoMonto, v.Total, v.Estado)
                If v.Estado <> "Completada" Then
                    dgvVentas.Rows(rIdx).DefaultCellStyle.ForeColor = UITheme.ColorTextMuted
                End If
            Next

            ' Actualizar las 4 tarjetas KPI Financieras (inteligencia comercial)
            Dim metricas = reporteService.GetMetricasFinancierasPeriodo(desde, hasta, vendedorId, clienteId, metodoFiltro)
            lblTotalVentas.Text = metricas.TotalFacturado.ToString("C2")
            lblTotalVentasSub.Text = $"{metricas.CantidadTickets} tickets | Efec: {metricas.TotalEfectivo:C2} | Dig: {metricas.TotalDigital:C2}"

            lblCostoMercaderia.Text = metricas.CostoMercaderia.ToString("C2")
            lblCostoMercaderiaSub.Text = "Costo de prendas vendidas"

            lblGananciaNeta.Text = metricas.GananciaNeta.ToString("C2")
            lblGananciaNetaSub.Text = $"Rentabilidad real ({metricas.PorcentajeMargen:N1} %)"

            lblTicketPromedio.Text = metricas.TicketPromedio.ToString("C2")
            lblTicketPromedioSub.Text = $"Gasto promedio ({metricas.CantidadTickets} compras)"
        End Sub

        Private Sub LoadAnulaciones()
            Dim desde = dtpDesde.Value.Date
            Dim hasta = dtpHasta.Value.Date

            Dim vendedorId As Integer = 0
            If cboVendedorGeneral IsNot Nothing AndAlso cboVendedorGeneral.SelectedItem IsNot Nothing Then
                Dim item = CType(cboVendedorGeneral.SelectedItem, FiltroComboItem)
                vendedorId = item.Id
            End If

            Dim anuladas = ventaService.GetVentas(desde, hasta, "", 0, vendedorId, "", "Anulada")
            dgvAnulaciones.Rows.Clear()

            Dim totalRevertido As Decimal = 0D
            For Each v In anuladas
                totalRevertido += v.Total
                dgvAnulaciones.Rows.Add(
                    v.Id,
                    v.NumeroTicket,
                    v.Fecha.ToString("dd/MM/yyyy HH:mm"),
                    v.ClienteNombre,
                    v.UsuarioNombre,
                    v.AutorizadoPorNombre,
                    v.MetodoPago,
                    v.Total,
                    v.MotivoAnulacion
                )
            Next

            lblResumenAnulaciones.Text = $"Ventas anuladas / devueltas en el período: {anuladas.Count}  |  Total revertido: {totalRevertido:C2}  (Doble clic para ver prendas devueltas)"
        End Sub

        Private Sub LoadVentasPorCliente()
            Dim desde = dtpDesde.Value.Date
            Dim hasta = dtpHasta.Value.Date

            Dim clienteId As Integer = 0
            If cboClientesFiltro.SelectedItem IsNot Nothing Then
                Dim item = CType(cboClientesFiltro.SelectedItem, FiltroComboItem)
                clienteId = item.Id
            End If

            Dim criterio As CriterioOrdenCliente = CriterioOrdenCliente.MayorFacturacion
            If cboOrdenClientes.SelectedItem IsNot Nothing Then
                Dim sel = cboOrdenClientes.SelectedItem.ToString()
                If sel.Contains("Salida") Then
                    criterio = CriterioOrdenCliente.MayorSalida
                ElseIf sel.Contains("Margen") Then
                    criterio = CriterioOrdenCliente.MayorMargen
                End If
            End If

            Dim lista = reporteService.GetReporteVentasPorCliente(desde, hasta, clienteId, criterio)
            dgvClientesReporte.Rows.Clear()

            Dim totClientes = lista.Count
            Dim totPrendas = 0
            Dim totFacturado As Decimal = 0D
            Dim totMargen As Decimal = 0D

            For Each c In lista
                totPrendas += c.TotalUnidades
                totFacturado += c.TotalFacturado
                totMargen += c.MargenGanancia

                Dim fechaStr = If(c.UltimaCompra.HasValue, c.UltimaCompra.Value.ToString("dd/MM/yyyy HH:mm"), "-")
                Dim pctStr = $"{c.PorcentajeMargen:N1} %"

                dgvClientesReporte.Rows.Add(c.ClienteId, c.ClienteNombre, c.DniCuit, c.Telefono, c.CantidadCompras, c.TotalUnidades, c.TotalFacturado, c.MargenGanancia, pctStr, fechaStr)
            Next

            lblResumenClientes.Text = $"Clientes: {totClientes} | Prendas: {totPrendas} | Facturado: {totFacturado:C2} | Margen Neto: {totMargen:C2}"

            DgvClientesReporte_SelectionChanged(Nothing, EventArgs.Empty)
        End Sub

        Private Sub LoadVentasPorVendedor()
            Dim desde = dtpDesde.Value.Date
            Dim hasta = dtpHasta.Value.Date

            Dim vendedorId As Integer = 0
            If cboVendedoresFiltro.SelectedItem IsNot Nothing Then
                Dim item = CType(cboVendedoresFiltro.SelectedItem, FiltroComboItem)
                vendedorId = item.Id
            End If

            Dim criterio As CriterioOrdenCliente = CriterioOrdenCliente.MayorFacturacion
            If cboOrdenVendedores.SelectedItem IsNot Nothing Then
                Dim sel = cboOrdenVendedores.SelectedItem.ToString()
                If sel.Contains("Salida") Then
                    criterio = CriterioOrdenCliente.MayorSalida
                ElseIf sel.Contains("Margen") Then
                    criterio = CriterioOrdenCliente.MayorMargen
                End If
            End If

            Dim lista = reporteService.GetReporteVentasPorVendedor(desde, hasta, vendedorId, criterio)
            dgvVendedoresReporte.Rows.Clear()

            Dim totVendedores = lista.Count
            Dim totPrendas = 0
            Dim totFacturado As Decimal = 0D
            Dim totMargen As Decimal = 0D

            For Each v In lista
                totPrendas += v.TotalUnidades
                totFacturado += v.TotalFacturado
                totMargen += v.MargenGanancia

                Dim fechaStr = If(v.UltimaVenta.HasValue, v.UltimaVenta.Value.ToString("dd/MM/yyyy HH:mm"), "-")
                Dim pctStr = $"{v.PorcentajeMargen:N1} %"

                dgvVendedoresReporte.Rows.Add(v.UsuarioId, v.VendedorNombre, v.Username, If(String.IsNullOrWhiteSpace(v.Dni), "-", v.Dni), v.CantidadVentas, v.TotalUnidades, v.TotalFacturado, v.MargenGanancia, pctStr, fechaStr)
            Next

            lblResumenVendedores.Text = $"Vendedores: {totVendedores} | Prendas: {totPrendas} | Facturado: {totFacturado:C2} | Margen Neto: {totMargen:C2}"

            DgvVendedoresReporte_SelectionChanged(Nothing, EventArgs.Empty)
        End Sub

        Private Sub DgvVendedoresReporte_SelectionChanged(sender As Object, e As EventArgs)
            If dgvVendedoresReporte.CurrentRow IsNot Nothing AndAlso dgvVendedoresReporte.CurrentRow.Cells("Id").Value IsNot Nothing Then
                Dim vendedorId As Integer = Convert.ToInt32(dgvVendedoresReporte.CurrentRow.Cells("Id").Value)
                Dim vendedorNombre As String = dgvVendedoresReporte.CurrentRow.Cells("Vendedor").Value.ToString()
                LoadTicketsVendedor(vendedorId, vendedorNombre)
            Else
                dgvTicketsVendedor.Rows.Clear()
                lblTicketsVendedorHeader.Text = "Seleccione un vendedor para ver sus comprobantes emitidos"
            End If
        End Sub

        Private Sub LoadTicketsVendedor(vendedorId As Integer, vendedorNombre As String)
            Dim desde = dtpDesde.Value.Date
            Dim hasta = dtpHasta.Value.Date

            lblTicketsVendedorHeader.Text = $"Comprobantes emitidos por: {vendedorNombre} (período {desde:dd/MM} al {hasta:dd/MM})"

            Dim tickets = ventaService.GetVentas(desde, hasta, "", 0, vendedorId)
            dgvTicketsVendedor.Rows.Clear()

            For Each v In tickets
                Dim rIdx = dgvTicketsVendedor.Rows.Add(v.Id, v.NumeroTicket, v.Fecha.ToString("dd/MM/yyyy HH:mm"), v.ClienteNombre, v.MetodoPago, v.Subtotal, v.DescuentoMonto, v.Total, v.Estado)
                If v.Estado <> "Completada" Then
                    dgvTicketsVendedor.Rows(rIdx).DefaultCellStyle.ForeColor = UITheme.ColorTextMuted
                End If
            Next
        End Sub

        Private Sub DgvClientesReporte_SelectionChanged(sender As Object, e As EventArgs)
            If dgvClientesReporte.CurrentRow IsNot Nothing AndAlso dgvClientesReporte.CurrentRow.Cells("Id").Value IsNot Nothing Then
                Dim clienteId As Integer = Convert.ToInt32(dgvClientesReporte.CurrentRow.Cells("Id").Value)
                Dim clienteNombre As String = dgvClientesReporte.CurrentRow.Cells("Cliente").Value.ToString()
                LoadTicketsCliente(clienteId, clienteNombre)
            Else
                dgvTicketsCliente.Rows.Clear()
                lblTicketsClienteHeader.Text = "Seleccione un cliente para ver sus comprobantes emitidos"
            End If
        End Sub

        Private Sub LoadTicketsCliente(clienteId As Integer, clienteNombre As String)
            Dim desde = dtpDesde.Value.Date
            Dim hasta = dtpHasta.Value.Date

            lblTicketsClienteHeader.Text = $"Comprobantes emitidos de: {clienteNombre} (período {desde:dd/MM} al {hasta:dd/MM})"

            Dim tickets = ventaService.GetVentas(desde, hasta, "", clienteId)
            dgvTicketsCliente.Rows.Clear()

            For Each v In tickets
                Dim rIdx = dgvTicketsCliente.Rows.Add(v.Id, v.NumeroTicket, v.Fecha.ToString("dd/MM/yyyy HH:mm"), v.UsuarioNombre, v.MetodoPago, v.Subtotal, v.DescuentoMonto, v.Total, v.Estado)
                If v.Estado <> "Completada" Then
                    dgvTicketsCliente.Rows(rIdx).DefaultCellStyle.ForeColor = UITheme.ColorTextMuted
                End If
            Next
        End Sub

        Private Sub LoadRanking()
            Dim desde = dtpDesde.Value.Date
            Dim hasta = dtpHasta.Value.Date

            Dim limite = 20
            If cboLimiteRanking.SelectedItem IsNot Nothing Then
                Select Case cboLimiteRanking.SelectedItem.ToString()
                    Case "Top 10"
                        limite = 10
                    Case "Top 20"
                        limite = 20
                    Case "Top 50"
                        limite = 50
                    Case "Todos"
                        limite = 0
                End Select
            End If

            If rbStockInmovilizado IsNot Nothing AndAlso rbStockInmovilizado.Checked Then
                ConfigurarColumnasStockInmovilizado()
                Dim inmovilizados = reporteService.GetStockInmovilizado(desde, hasta, limite)
                dgvRanking.Rows.Clear()

                Dim totalUnidInmov = 0
                Dim totalCapitalInmov As Decimal = 0D
                For Each item In inmovilizados
                    totalUnidInmov += item.StockActual
                    totalCapitalInmov += item.CostoInmovilizado
                    Dim ultVentaStr As String
                    If item.UltimaVenta.HasValue Then
                        Dim dias = Math.Max(0, CInt((DateTime.Today - item.UltimaVenta.Value.Date).TotalDays))
                        ultVentaStr = $"{item.UltimaVenta.Value:dd/MM/yyyy} (hace {dias} días)"
                    Else
                        ultVentaStr = "Nunca vendida"
                    End If
                    dgvRanking.Rows.Add(item.CodigoBarra, item.Nombre, item.Talle, item.Color, item.StockActual, item.CostoInmovilizado, item.ValorVenta, ultVentaStr)
                Next

                If lblResumenRanking IsNot Nothing Then
                    lblResumenRanking.Text = $"Variantes sin rotación: {inmovilizados.Count} ({totalUnidInmov} prendas) | Capital inmovilizado: {totalCapitalInmov:C2}"
                    lblResumenRanking.ForeColor = UITheme.ColorWarning
                End If
                Return
            End If

            ConfigurarColumnasRanking()
            Dim criterio = If(rbMayorMargen.Checked, CriterioRankingProducto.MayorMargen, CriterioRankingProducto.MayorSalida)
            Dim top = reporteService.GetTopProductosVendidos(limite, desde, hasta, criterio)
            dgvRanking.Rows.Clear()

            Dim totalUnidadesRanking = 0
            Dim totalMargenRanking As Decimal = 0D
            For Each p In top
                totalUnidadesRanking += p.CantidadVendida
                totalMargenRanking += p.MargenGanancia
                Dim pctStr = $"{p.PorcentajeMargen:N1} %"
                dgvRanking.Rows.Add(p.Nombre, p.Talle, p.CantidadVendida, p.TotalRecaudado, p.CostoTotal, p.MargenGanancia, pctStr)
            Next

            If lblResumenRanking IsNot Nothing Then
                lblResumenRanking.Text = $"Prendas vendidas en ranking: {totalUnidadesRanking} | Margen acumulado: {totalMargenRanking:C2}"
                lblResumenRanking.ForeColor = UITheme.ColorSuccessDark
            End If
        End Sub

        Private Sub DgvVentas_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
            If e.RowIndex >= 0 Then
                Dim row = dgvVentas.Rows(e.RowIndex)
                Dim ventaId As Integer = Convert.ToInt32(row.Cells("Id").Value)
                Dim ticketNum As String = row.Cells("Ticket").Value.ToString()
                MostrarModalDetalleVenta(ventaId, ticketNum)
            End If
        End Sub

        Private Sub DgvAnulaciones_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
            If e.RowIndex >= 0 Then
                Dim row = dgvAnulaciones.Rows(e.RowIndex)
                Dim ventaId As Integer = Convert.ToInt32(row.Cells("Id").Value)
                Dim ticketNum As String = row.Cells("Ticket").Value.ToString()
                MostrarModalDetalleVenta(ventaId, ticketNum)
            End If
        End Sub

        Private Sub DgvTicketsCliente_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
            If e.RowIndex >= 0 Then
                Dim row = dgvTicketsCliente.Rows(e.RowIndex)
                Dim ventaId As Integer = Convert.ToInt32(row.Cells("Id").Value)
                Dim ticketNum As String = row.Cells("Ticket").Value.ToString()
                MostrarModalDetalleVenta(ventaId, ticketNum)
            End If
        End Sub

        Private Sub BtnVerDetalleTicketCliente_Click(sender As Object, e As EventArgs)
            If dgvTicketsCliente.CurrentRow IsNot Nothing Then
                Dim ventaId As Integer = Convert.ToInt32(dgvTicketsCliente.CurrentRow.Cells("Id").Value)
                Dim ticketNum As String = dgvTicketsCliente.CurrentRow.Cells("Ticket").Value.ToString()
                MostrarModalDetalleVenta(ventaId, ticketNum)
            Else
                MessageBox.Show("Por favor seleccione un comprobante en la lista para ver el detalle de artículos.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End Sub

        Private Sub DgvTicketsVendedor_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
            If e.RowIndex >= 0 Then
                Dim row = dgvTicketsVendedor.Rows(e.RowIndex)
                Dim ventaId As Integer = Convert.ToInt32(row.Cells("Id").Value)
                Dim ticketNum As String = row.Cells("Ticket").Value.ToString()
                MostrarModalDetalleVenta(ventaId, ticketNum)
            End If
        End Sub

        Private Sub BtnVerDetalleTicketVendedor_Click(sender As Object, e As EventArgs)
            If dgvTicketsVendedor.CurrentRow IsNot Nothing Then
                Dim ventaId As Integer = Convert.ToInt32(dgvTicketsVendedor.CurrentRow.Cells("Id").Value)
                Dim ticketNum As String = dgvTicketsVendedor.CurrentRow.Cells("Ticket").Value.ToString()
                MostrarModalDetalleVenta(ventaId, ticketNum)
            Else
                MessageBox.Show("Por favor seleccione un comprobante en la lista para ver el detalle de artículos.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End Sub

        Private Sub BtnDescargarPdfTicketVendedor_Click(sender As Object, e As EventArgs)
            If dgvTicketsVendedor.CurrentRow IsNot Nothing AndAlso dgvTicketsVendedor.CurrentRow.Cells("Id").Value IsNot Nothing Then
                Dim ventaId As Integer = Convert.ToInt32(dgvTicketsVendedor.CurrentRow.Cells("Id").Value)
                DescargarTicketPdfPorId(ventaId)
            Else
                MessageBox.Show("Por favor seleccione un comprobante del vendedor para descargar su ticket en PDF.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End Sub

        Private Sub BtnDescargarPdf_Click(sender As Object, e As EventArgs)
            If tabs.SelectedTab Is tabVentas Then
                If dgvVentas.CurrentRow IsNot Nothing AndAlso dgvVentas.CurrentRow.Cells("Id").Value IsNot Nothing Then
                    Dim ventaId As Integer = Convert.ToInt32(dgvVentas.CurrentRow.Cells("Id").Value)
                    DescargarTicketPdfPorId(ventaId)
                Else
                    MessageBox.Show("Seleccione una venta en el listado general para descargar su ticket en PDF.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            ElseIf tabs.SelectedTab Is tabAnulaciones Then
                If dgvAnulaciones.CurrentRow IsNot Nothing AndAlso dgvAnulaciones.CurrentRow.Cells("Id").Value IsNot Nothing Then
                    Dim ventaId As Integer = Convert.ToInt32(dgvAnulaciones.CurrentRow.Cells("Id").Value)
                    DescargarTicketPdfPorId(ventaId)
                Else
                    MessageBox.Show("Seleccione un comprobante anulado para descargar su ticket en PDF.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            ElseIf tabs.SelectedTab Is tabClientes Then
                BtnDescargarPdfTicketCliente_Click(sender, e)
            ElseIf tabs.SelectedTab Is tabVendedores Then
                BtnDescargarPdfTicketVendedor_Click(sender, e)
            Else
                MessageBox.Show("Seleccione un comprobante en las pestañas de Ventas, Anulaciones, Clientes o Vendedores para descargar su ticket en PDF.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End Sub

        Private Sub BtnDescargarPdfTicketCliente_Click(sender As Object, e As EventArgs)
            If dgvTicketsCliente.CurrentRow IsNot Nothing AndAlso dgvTicketsCliente.CurrentRow.Cells("Id").Value IsNot Nothing Then
                Dim ventaId As Integer = Convert.ToInt32(dgvTicketsCliente.CurrentRow.Cells("Id").Value)
                DescargarTicketPdfPorId(ventaId)
            Else
                MessageBox.Show("Por favor seleccione un comprobante del cliente para descargar su ticket en PDF.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End Sub

        Private Sub DescargarTicketPdfPorId(ventaId As Integer)
            Dim venta = ventaService.GetVentaPorId(ventaId)
            If venta Is Nothing Then
                MessageBox.Show("No se encontró la venta seleccionada.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim detalles = ventaService.GetDetalleVenta(ventaId)
            If detalles Is Nothing OrElse detalles.Count = 0 Then
                MessageBox.Show("No se encontraron renglones de artículos para este ticket.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            venta.Detalles = detalles
            GuardarTicketEnPdf(Me, venta)
        End Sub

        Private Sub MostrarModalDetalleVenta(ventaId As Integer, ticketNumero As String)
            Dim venta = ventaService.GetVentaPorId(ventaId)
            If venta Is Nothing Then
                MessageBox.Show("No se encontró la venta seleccionada.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim detalles = ventaService.GetDetalleVenta(ventaId)
            If detalles Is Nothing OrElse detalles.Count = 0 Then
                MessageBox.Show("No se encontraron renglones de artículos para este ticket.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If
            venta.Detalles = detalles

            Dim modal As New Form() With {
                .Text = $"Detalle del Comprobante #{ticketNumero}",
                .Size = New Size(760, 440),
                .MinimumSize = New Size(640, 360),
                .StartPosition = FormStartPosition.CenterParent,
                .FormBorderStyle = FormBorderStyle.Sizable,
                .MaximizeBox = True,
                .MinimizeBox = False,
                .BackColor = UITheme.ColorBackground
            }

            Dim pnlHead As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 52,
                .BackColor = UITheme.ColorSecondary,
                .Padding = New Padding(15, 12, 15, 12)
            }
            Dim estadoExtra = If(venta.Estado.Equals("Anulada", StringComparison.OrdinalIgnoreCase), $" [ANULADA - Autorizó: {venta.AutorizadoPorNombre}]", "")
            Dim lblHead As New Label() With {
                .Text = $"Prendas adquiridas en Ticket #{ticketNumero}{estadoExtra}",
                .Font = UITheme.FontSubheading,
                .ForeColor = Color.White,
                .AutoSize = True,
                .Location = New Point(15, 13)
            }
            pnlHead.Controls.Add(lblHead)

            Dim dgvModal As New DataGridView() With {.Dock = DockStyle.Fill}
            UITheme.StyleDataGrid(dgvModal)
            dgvModal.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            dgvModal.Columns.Add("Codigo", "Código")
            dgvModal.Columns("Codigo").FillWeight = 15
            dgvModal.Columns("Codigo").MinimumWidth = 90
            dgvModal.Columns.Add("Articulo", "Prenda / Modelo")
            dgvModal.Columns("Articulo").FillWeight = 32
            dgvModal.Columns("Articulo").MinimumWidth = 140
            dgvModal.Columns.Add("Talle", "Talle")
            dgvModal.Columns("Talle").FillWeight = 10
            dgvModal.Columns("Talle").MinimumWidth = 60
            dgvModal.Columns.Add("Color", "Color")
            dgvModal.Columns("Color").FillWeight = 11
            dgvModal.Columns("Color").MinimumWidth = 70
            dgvModal.Columns.Add("Precio", "Precio Unit.")
            dgvModal.Columns("Precio").FillWeight = 14
            dgvModal.Columns("Precio").MinimumWidth = 90
            dgvModal.Columns("Precio").DefaultCellStyle.Format = "C2"
            dgvModal.Columns.Add("Cantidad", "Cant.")
            dgvModal.Columns("Cantidad").FillWeight = 9
            dgvModal.Columns("Cantidad").MinimumWidth = 55
            dgvModal.Columns("Cantidad").DefaultCellStyle.Font = UITheme.FontBold
            dgvModal.Columns("Cantidad").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvModal.Columns.Add("Subtotal", "Subtotal")
            dgvModal.Columns("Subtotal").FillWeight = 14
            dgvModal.Columns("Subtotal").MinimumWidth = 90
            dgvModal.Columns("Subtotal").DefaultCellStyle.Format = "C2"
            dgvModal.Columns("Subtotal").DefaultCellStyle.Font = UITheme.FontBold

            For Each d In detalles
                dgvModal.Rows.Add(d.CodigoBarra, d.DescripcionArticulo, d.TalleNombre, d.Color, d.PrecioUnitario, d.Cantidad, d.Subtotal)
            Next

            Dim pnlBottom As New Panel() With {
                .Dock = DockStyle.Bottom,
                .Height = 52,
                .BackColor = UITheme.ColorSurface
            }

            Dim btnExportarPdfModal As New Button() With {
                .Text = "📄 Reimprimir / Guardar PDF",
                .Size = New Size(210, 34),
                .Location = New Point(modal.ClientSize.Width - 340, 9),
                .Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
            }
            UITheme.StyleButton(btnExportarPdfModal, "Primary")
            AddHandler btnExportarPdfModal.Click, Sub() GuardarTicketEnPdf(modal, venta)

            Dim btnCerrar As New Button() With {
                .Text = "Cerrar",
                .Size = New Size(100, 34),
                .Location = New Point(modal.ClientSize.Width - 118, 9),
                .Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
            }
            UITheme.StyleButton(btnCerrar, "Secondary")
            AddHandler btnCerrar.Click, Sub() modal.Close()
            pnlBottom.Controls.AddRange({btnExportarPdfModal, btnCerrar})

            modal.Controls.Add(dgvModal)
            modal.Controls.Add(pnlBottom)
            modal.Controls.Add(pnlHead)

            modal.ShowDialog(Me)
        End Sub

        Friend Shared Sub GuardarTicketEnPdf(owner As IWin32Window, venta As Venta)
            If venta Is Nothing OrElse venta.Detalles Is Nothing OrElse venta.Detalles.Count = 0 Then
                MessageBox.Show(owner, "No hay información suficiente para imprimir el ticket.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim nombreImpresoraPdf As String = String.Empty
            For Each nombreImpresora As String In PrinterSettings.InstalledPrinters
                If nombreImpresora.Equals("Microsoft Print to PDF", StringComparison.OrdinalIgnoreCase) Then
                    nombreImpresoraPdf = nombreImpresora
                    Exit For
                End If
            Next

            If String.IsNullOrWhiteSpace(nombreImpresoraPdf) Then
                MessageBox.Show(owner, "No está instalada la impresora «Microsoft Print to PDF». Habilítela en las características de Windows para guardar el ticket.", "Impresora PDF no disponible", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim nombreArchivo = "Ticket-" & venta.NumeroTicket & ".pdf"
            For Each caracter In IO.Path.GetInvalidFileNameChars()
                nombreArchivo = nombreArchivo.Replace(caracter, "-"c)
            Next

            Using dialogoGuardar As New SaveFileDialog()
                dialogoGuardar.Title = "Guardar ticket de compra en PDF"
                dialogoGuardar.Filter = "Documento PDF (*.pdf)|*.pdf"
                dialogoGuardar.FileName = nombreArchivo
                dialogoGuardar.DefaultExt = "pdf"
                dialogoGuardar.AddExtension = True
                dialogoGuardar.OverwritePrompt = True

                If dialogoGuardar.ShowDialog(owner) <> DialogResult.OK Then
                    Return
                End If

                Try
                    Dim cfgSvc As New ConfiguracionService()
                    Dim lineasTicket = ConstruirLineasTicket(venta, cfgSvc.GetConfiguracion())
                    Dim indiceLinea As Integer = 0
                    Using documento As New PrintDocument()
                        documento.DocumentName = nombreArchivo
                        documento.PrinterSettings.PrinterName = nombreImpresoraPdf
                        documento.PrinterSettings.PrintToFile = True
                        documento.PrinterSettings.PrintFileName = dialogoGuardar.FileName
                        documento.OriginAtMargins = False
                        documento.DefaultPageSettings.Margins = New Margins(45, 45, 45, 45)

                        AddHandler documento.PrintPage,
                            Sub(sender, args)
                                Using fuente As New Font("Courier New", 9.0F, FontStyle.Regular)
                                    Using pincel As New SolidBrush(Color.Black)
                                        Dim altoLinea = fuente.GetHeight(args.Graphics) + 2.0F
                                        Dim posicionY = CSng(args.MarginBounds.Top)
                                        Dim margenIzquierdo = CSng(args.MarginBounds.Left)
                                        Dim limiteInferior = CSng(args.MarginBounds.Bottom)

                                        While indiceLinea < lineasTicket.Count AndAlso posicionY + altoLinea <= limiteInferior
                                            args.Graphics.DrawString(lineasTicket(indiceLinea), fuente, pincel, margenIzquierdo, posicionY)
                                            posicionY += altoLinea
                                            indiceLinea += 1
                                        End While

                                        args.HasMorePages = indiceLinea < lineasTicket.Count
                                    End Using
                                End Using
                            End Sub

                        documento.Print()
                    End Using

                    MessageBox.Show(owner, "El ticket se guardó correctamente en PDF.", "Ticket guardado", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    AppLogger.LogError("FrmReportes.GuardarTicketEnPdf", ex)
                    MessageBox.Show(owner, "No se pudo guardar el ticket en PDF: " & ex.Message, "Error de impresión", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End Using
        End Sub

        Friend Shared Function ConstruirLineasTicket(venta As Venta, configuracion As ConfiguracionComercio) As List(Of String)
            Dim vendedorNombre As String = If(String.IsNullOrWhiteSpace(venta.UsuarioNombre),
                                              If(AuthService.CurrentUser IsNot Nothing, AuthService.CurrentUser.NombreCompleto, "Caja"),
                                              venta.UsuarioNombre)

            Dim contenido As New System.Text.StringBuilder()
            contenido.AppendLine("==========================================")
            contenido.AppendLine(configuracion.NombreComercio.ToUpperInvariant())
            contenido.AppendLine($"CUIT: {configuracion.Cuit} - {configuracion.CondicionIva}")
            contenido.AppendLine($"Dirección: {configuracion.Direccion}")
            contenido.AppendLine($"Teléfono: {configuracion.Telefono}")
            contenido.AppendLine("==========================================")
            contenido.AppendLine($"TICKET Nº: {venta.NumeroTicket}")
            contenido.AppendLine($"Fecha: {venta.Fecha:dd/MM/yyyy HH:mm:ss}")
            contenido.AppendLine($"Vendedor: {vendedorNombre}")
            contenido.AppendLine($"Cliente: {venta.ClienteNombre}")
            contenido.AppendLine($"Medio de pago: {venta.MetodoPago}")
            If Not String.IsNullOrWhiteSpace(venta.Estado) AndAlso venta.Estado.Equals("Anulada", StringComparison.OrdinalIgnoreCase) Then
                contenido.AppendLine("ESTADO: ANULADA")
            End If
            contenido.AppendLine("------------------------------------------")
            contenido.AppendLine("CANT.  DESCRIPCIÓN / TALLE       SUBTOTAL")
            contenido.AppendLine("------------------------------------------")

            For Each detalle In venta.Detalles
                contenido.AppendLine($"{detalle.Cantidad,4}   {detalle.DescripcionArticulo} - {detalle.TalleNombre}")
                contenido.AppendLine($"      {detalle.PrecioUnitario:C2} c/u{detalle.Subtotal,25:C2}")
            Next

            contenido.AppendLine("------------------------------------------")
            contenido.AppendLine($"SUBTOTAL: {venta.Subtotal:C2}")
            If venta.DescuentoMonto > 0D Then
                contenido.AppendLine($"DESCUENTO ({venta.DescuentoPorcentaje:N2}%): -{venta.DescuentoMonto:C2}")
            End If
            If venta.RecargoMonto > 0D Then
                contenido.AppendLine($"RECARGO: {venta.RecargoMonto:C2}")
            End If
            contenido.AppendLine($"TOTAL A PAGAR: {venta.Total:C2}")
            contenido.AppendLine($"PAGADO: {venta.MontoAbonado:C2}")
            contenido.AppendLine($"VUELTO: {venta.Vuelto:C2}")
            contenido.AppendLine("==========================================")
            contenido.AppendLine(configuracion.MensajeTicket)
            contenido.AppendLine("==========================================")

            Dim lineasOriginales = contenido.ToString().Split({vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
            Dim lineasTicket As New List(Of String)()
            Const anchoMaximo As Integer = 85

            For Each linea In lineasOriginales
                Dim lineaActual As String = String.Empty
                For Each palabra In linea.Split(" "c)
                    Dim fragmento = palabra
                    While fragmento.Length > anchoMaximo
                        If lineaActual.Length > 0 Then
                            lineasTicket.Add(lineaActual)
                            lineaActual = String.Empty
                        End If
                        lineasTicket.Add(fragmento.Substring(0, anchoMaximo))
                        fragmento = fragmento.Substring(anchoMaximo)
                    End While

                    If lineaActual.Length = 0 Then
                        lineaActual = fragmento
                    ElseIf lineaActual.Length + fragmento.Length + 1 <= anchoMaximo Then
                        lineaActual &= " " & fragmento
                    Else
                        lineasTicket.Add(lineaActual)
                        lineaActual = fragmento
                    End If
                Next

                If lineaActual.Length > 0 Then
                    lineasTicket.Add(lineaActual)
                End If
            Next

            Return lineasTicket
        End Function

        ''' <summary>
        ''' Anulación de venta con reversión de inventario y auditoría.
        ''' Requiere autorización obligatoria de Gerente y justificación de motivo (cambio de prenda, falla, etc.).
        ''' </summary>
        Private Sub BtnAnularVenta_Click(sender As Object, e As EventArgs)
            If AuthService.IsAdmin AndAlso Not AuthService.SolicitarAutorizacionManager(Me, "Anular comprobantes de venta emitidos requiere autorización de un Gerente.") Then
                Return
            End If
            If Not AuthService.IsAdminOrManager Then
                Return
            End If

            If dgvVentas.CurrentRow IsNot Nothing Then
                Dim ventaId As Integer = Convert.ToInt32(dgvVentas.CurrentRow.Cells("Id").Value)
                Dim ticket As String = dgvVentas.CurrentRow.Cells("Ticket").Value.ToString()
                Dim estado As String = dgvVentas.CurrentRow.Cells("Estado").Value.ToString()

                If estado.Equals("Anulada", StringComparison.OrdinalIgnoreCase) Then
                    MessageBox.Show("Esta venta ya fue anulada.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                Dim prompt As New Form() With {
                    .Text = "Anular Venta y Devolver Stock",
                    .Size = New Size(420, 240),
                    .StartPosition = FormStartPosition.CenterParent,
                    .FormBorderStyle = FormBorderStyle.FixedDialog,
                    .MaximizeBox = False,
                    .MinimizeBox = False,
                    .BackColor = Color.White
                }

                Dim lblInfo As New Label() With {.Text = $"¿Confirma la anulación del Ticket #{ticket}?" & vbCrLf & "Las prendas se reintegrarán automáticamente al stock.", .Font = UITheme.FontBold, .Location = New Point(25, 20), .Size = New Size(360, 40)}
                Dim lblMot As New Label() With {.Text = "Motivo de la anulación / cambio:", .Font = UITheme.FontRegular, .Location = New Point(25, 65), .AutoSize = True}
                Dim txtMot As New TextBox() With {.Location = New Point(25, 90), .Size = New Size(350, 26), .Text = "Devolución de cliente"}
                UITheme.StyleTextBox(txtMot)

                Dim btnOk As New Button() With {.Text = "Confirmar Anulación", .Location = New Point(25, 135), .Size = New Size(350, 38)}
                UITheme.StyleButton(btnOk, "Danger")
                AddHandler btnOk.Click, Sub()
                                            prompt.DialogResult = DialogResult.OK
                                            prompt.Close()
                                        End Sub

                prompt.Controls.AddRange({lblInfo, lblMot, txtMot, btnOk})

                If prompt.ShowDialog() = DialogResult.OK Then
                    Dim userId As Integer = If(AuthService.CurrentUser IsNot Nothing, AuthService.CurrentUser.Id, 1)
                    Dim autorizadorId As Integer = If(AuthService.LastAuthorizingSupervisor IsNot Nothing, AuthService.LastAuthorizingSupervisor.Id, userId)
                    Dim errMsg As String = ""
                    If ventaService.AnularVenta(ventaId, autorizadorId, txtMot.Text.Trim(), errMsg) Then
                        MessageBox.Show("Venta anulada y stock devuelto exitosamente.", "Anulación Exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        LoadReportes()
                    Else
                        MessageBox.Show("Error: " & errMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End If
            End If
        End Sub

    End Class
End Namespace
