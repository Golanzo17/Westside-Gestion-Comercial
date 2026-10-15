Imports System.Drawing
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

        Private dtpDesde As DateTimePicker
        Private dtpHasta As DateTimePicker
        Private btnFiltrar As Button
        Private btnAnularVenta As Button

        Private lblTotalVentas As Label
        Private lblCantTickets As Label
        Private lblEfectivo As Label
        Private lblDigital As Label

        Private tabs As TabControl
        Private tabVentas As TabPage
        Private dgvVentas As DataGridView

        ' Pestaña: Ventas por Cliente
        Private tabClientes As TabPage
        Private cboClientesFiltro As ComboBox
        Private cboOrdenClientes As ComboBox
        Private lblResumenClientes As Label
        Private dgvClientesReporte As DataGridView
        Private dgvTicketsCliente As DataGridView
        Private lblTicketsClienteHeader As Label
        Private btnVerDetalleTicketCliente As Button

        ' Pestaña: Ranking de Prendas
        Private tabRanking As TabPage
        Private rbMayorSalida As RadioButton
        Private rbMayorMargen As RadioButton
        Private cboLimiteRanking As ComboBox
        Private dgvRanking As DataGridView

        Private isLoading As Boolean = True

        ' Helper interno para el ComboBox de clientes
        Private Class ClienteComboItem
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
            isLoading = False
            LoadReportes()
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

            ' Toolbar de Fechas general
            Dim pnlFilter As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 55,
                .BackColor = UITheme.ColorSurface
            }
            Dim flpFilter As New FlowLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(10, 12, 10, 10),
                .WrapContents = False
            }

            Dim lblD As New Label() With {.Text = "Desde:", .Font = UITheme.FontBold, .AutoSize = True, .Margin = New Padding(0, 6, 4, 0)}
            dtpDesde = New DateTimePicker() With {.Size = New Size(125, 26), .Format = DateTimePickerFormat.Short, .Value = DateTime.Today.AddDays(-7), .Margin = New Padding(0, 3, 0, 0)}

            Dim lblH As New Label() With {.Text = "Hasta:", .Font = UITheme.FontBold, .AutoSize = True, .Margin = New Padding(10, 6, 4, 0)}
            dtpHasta = New DateTimePicker() With {.Size = New Size(125, 26), .Format = DateTimePickerFormat.Short, .Value = DateTime.Today, .Margin = New Padding(0, 3, 0, 0)}

            btnFiltrar = New Button() With {.Text = "Generar Reportes", .Height = 30, .AutoSize = True, .AutoSizeMode = AutoSizeMode.GrowAndShrink, .Padding = New Padding(10, 0, 10, 0), .Margin = New Padding(12, 2, 0, 0)}
            UITheme.StyleButton(btnFiltrar, "Primary")
            AddHandler btnFiltrar.Click, Sub() LoadReportes()

            btnAnularVenta = New Button() With {.Text = "Anular Venta Seleccionada", .Height = 30, .AutoSize = True, .AutoSizeMode = AutoSizeMode.GrowAndShrink, .Padding = New Padding(10, 0, 10, 0), .Margin = New Padding(10, 2, 0, 0)}
            UITheme.StyleButton(btnAnularVenta, "Danger")
            btnAnularVenta.Visible = AuthService.IsAdminOrManager
            AddHandler btnAnularVenta.Click, AddressOf BtnAnularVenta_Click

            flpFilter.Controls.AddRange({lblD, dtpDesde, lblH, dtpHasta, btnFiltrar, btnAnularVenta})
            pnlFilter.Controls.Add(flpFilter)

            ' Tarjetas KPI para balance global del período
            Dim pnlKpis As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 125,
                .BackColor = UITheme.ColorSurfaceMuted,
                .Padding = New Padding(15, 5, 15, 5)
            }

            Dim cardTot = UITheme.CreateKpiCard("TOTAL FACTURADO", "$ 0,00", "En el período seleccionado", UITheme.ColorSuccess)
            cardTot.Location = New Point(15, 5)
            lblTotalVentas = CType(cardTot.Controls(2), Label)

            Dim cardCant = UITheme.CreateKpiCard("CANTIDAD VENTAS", "0", "Tickets emitidos", UITheme.ColorPrimary)
            cardCant.Location = New Point(265, 5)
            lblCantTickets = CType(cardCant.Controls(2), Label)

            Dim cardEf = UITheme.CreateKpiCard("COBRADO EN EFECTIVO", "$ 0,00", "Total en billetes", UITheme.ColorWarning)
            cardEf.Location = New Point(515, 5)
            lblEfectivo = CType(cardEf.Controls(2), Label)

            Dim cardDig = UITheme.CreateKpiCard("COBRADO DIGITAL", "$ 0,00", "Tarjetas y transferencias", UITheme.ColorInfo)
            cardDig.Location = New Point(765, 5)
            lblDigital = CType(cardDig.Controls(2), Label)

            pnlKpis.Controls.AddRange({cardTot, cardCant, cardEf, cardDig})

            ' TabControl con 3 secciones analíticas
            tabs = New TabControl() With {
                .Dock = DockStyle.Fill,
                .Padding = New Point(14, 7),
                .Font = UITheme.FontBold
            }

            ' Tab 1: Comprobantes y Ventas Globales
            tabVentas = New TabPage("Listado General de Ventas") With {.BackColor = UITheme.ColorBackground}
            dgvVentas = New DataGridView() With {.Dock = DockStyle.Fill}
            UITheme.StyleDataGrid(dgvVentas)
            ConfigurarColumnasVentas()
            AddHandler dgvVentas.CellDoubleClick, AddressOf DgvVentas_CellDoubleClick
            tabVentas.Controls.Add(dgvVentas)
            tabs.TabPages.Add(tabVentas)

            ' Tab 2: Ventas por Cliente (con filtro y desglose)
            tabClientes = New TabPage("Ventas por Cliente y Fecha") With {.BackColor = UITheme.ColorBackground}
            ConstruirTabClientes()
            tabs.TabPages.Add(tabClientes)

            ' Tab 3: Ranking de Prendas (Mayor Salida y Mayor Margen)
            tabRanking = New TabPage("Ranking de Prendas y Rentabilidad") With {.BackColor = UITheme.ColorBackground}
            ConstruirTabRanking()
            tabs.TabPages.Add(tabRanking)

            ' Restricción de privilegios para el perfil Vendedor
            If AuthService.IsVendor Then
                tabs.TabPages.Remove(tabVentas)
                pnlKpis.Visible = False
                btnAnularVenta.Visible = False

                ' Ocultamos métricas confidenciales de costo/margen
                rbMayorMargen.Visible = False
                cboOrdenClientes.Items.Remove("Mayor Margen de Ganancia ($)")
                dgvRanking.Columns("Costo").Visible = False
                dgvRanking.Columns("Margen").Visible = False
                dgvRanking.Columns("PctMargen").Visible = False
                dgvClientesReporte.Columns("Margen").Visible = False
                dgvClientesReporte.Columns("Rentabilidad").Visible = False
            End If

            ' Jerarquía de docking
            Me.Controls.Add(tabs)
            Me.Controls.Add(pnlKpis)
            Me.Controls.Add(pnlFilter)
            Me.Controls.Add(pnlHeader)
        End Sub

        Private Sub ConstruirTabClientes()
            ' Panel superior de filtros para clientes
            Dim pnlClientesToolbar As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 50,
                .BackColor = UITheme.ColorSurface
            }
            Dim flpClientes As New FlowLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(10, 10, 10, 8),
                .WrapContents = False
            }

            Dim lblCli As New Label() With {.Text = "Cliente:", .Font = UITheme.FontBold, .AutoSize = True, .Margin = New Padding(0, 6, 4, 0)}
            cboClientesFiltro = New ComboBox() With {
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Width = 260,
                .Margin = New Padding(0, 3, 0, 0)
            }
            UITheme.StyleComboBox(cboClientesFiltro)
            AddHandler cboClientesFiltro.SelectedIndexChanged, Sub()
                                                                   If Not isLoading Then LoadVentasPorCliente()
                                                               End Sub

            Dim lblOrd As New Label() With {.Text = "Ordenar por:", .Font = UITheme.FontBold, .AutoSize = True, .Margin = New Padding(14, 6, 4, 0)}
            cboOrdenClientes = New ComboBox() With {
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Width = 220,
                .Margin = New Padding(0, 3, 0, 0)
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
                .Margin = New Padding(20, 7, 0, 0),
                .Text = "Cargando métricas..."
            }

            flpClientes.Controls.AddRange({lblCli, cboClientesFiltro, lblOrd, cboOrdenClientes, lblResumenClientes})
            pnlClientesToolbar.Controls.Add(flpClientes)

            ' SplitContainer: Clientes arriba, comprobantes abajo
            Dim scClientes As New SplitContainer() With {
                .Dock = DockStyle.Fill,
                .Orientation = Orientation.Horizontal,
                .SplitterDistance = 270,
                .SplitterWidth = 6
            }

            ' Grilla principal de Clientes
            dgvClientesReporte = New DataGridView() With {.Dock = DockStyle.Fill}
            UITheme.StyleDataGrid(dgvClientesReporte)
            ConfigurarColumnasClientesReporte()
            AddHandler dgvClientesReporte.SelectionChanged, AddressOf DgvClientesReporte_SelectionChanged
            scClientes.Panel1.Controls.Add(dgvClientesReporte)

            ' Panel inferior: detalle de tickets del cliente
            Dim pnlTicketsHeader As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 36,
                .BackColor = UITheme.ColorSurfaceMuted,
                .Padding = New Padding(12, 6, 12, 6)
            }
            lblTicketsClienteHeader = New Label() With {
                .Text = "Comprobantes emitidos del cliente seleccionado",
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorTextPrimary,
                .AutoSize = True,
                .Dock = DockStyle.Left
            }
            btnVerDetalleTicketCliente = New Button() With {
                .Text = "Ver Detalle de Prendas del Ticket",
                .Dock = DockStyle.Right,
                .Width = 230,
                .Height = 26
            }
            UITheme.StyleButton(btnVerDetalleTicketCliente, "Secondary")
            AddHandler btnVerDetalleTicketCliente.Click, AddressOf BtnVerDetalleTicketCliente_Click

            pnlTicketsHeader.Controls.Add(lblTicketsClienteHeader)
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

        Private Sub ConstruirTabRanking()
            Dim pnlRankingToolbar As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 48,
                .BackColor = UITheme.ColorSurface
            }
            Dim flpRanking As New FlowLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(12, 10, 10, 8),
                .WrapContents = False
            }

            Dim lblCrit As New Label() With {.Text = "Criterio de Ranking:", .Font = UITheme.FontBold, .AutoSize = True, .Margin = New Padding(0, 5, 8, 0)}

            rbMayorSalida = New RadioButton() With {
                .Text = "Mayor Salida (Unidades Vendidas)",
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorPrimaryDark,
                .AutoSize = True,
                .Checked = True,
                .Margin = New Padding(0, 3, 14, 0)
            }
            AddHandler rbMayorSalida.CheckedChanged, Sub()
                                                         If Not isLoading AndAlso rbMayorSalida.Checked Then LoadRanking()
                                                     End Sub

            rbMayorMargen = New RadioButton() With {
                .Text = "Mayor Margen de Ganancia (Rentabilidad $)",
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorSuccessDark,
                .AutoSize = True,
                .Margin = New Padding(0, 3, 14, 0)
            }
            AddHandler rbMayorMargen.CheckedChanged, Sub()
                                                         If Not isLoading AndAlso rbMayorMargen.Checked Then LoadRanking()
                                                     End Sub

            Dim lblLim As New Label() With {.Text = "Mostrar:", .Font = UITheme.FontBold, .AutoSize = True, .Margin = New Padding(10, 5, 4, 0)}
            cboLimiteRanking = New ComboBox() With {
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Width = 120,
                .Margin = New Padding(0, 2, 0, 0)
            }
            UITheme.StyleComboBox(cboLimiteRanking)
            cboLimiteRanking.Items.AddRange({"Top 10", "Top 20", "Top 50", "Todos"})
            cboLimiteRanking.SelectedIndex = 1
            AddHandler cboLimiteRanking.SelectedIndexChanged, Sub()
                                                                  If Not isLoading Then LoadRanking()
                                                              End Sub

            flpRanking.Controls.AddRange({lblCrit, rbMayorSalida, rbMayorMargen, lblLim, cboLimiteRanking})
            pnlRankingToolbar.Controls.Add(flpRanking)

            dgvRanking = New DataGridView() With {.Dock = DockStyle.Fill}
            UITheme.StyleDataGrid(dgvRanking)
            ConfigurarColumnasRanking()

            tabRanking.Controls.Add(dgvRanking)
            tabRanking.Controls.Add(pnlRankingToolbar)
        End Sub

        Private Sub CargarComboClientes()
            cboClientesFiltro.Items.Clear()
            cboClientesFiltro.Items.Add(New ClienteComboItem(0, "[ Todos los clientes ]"))

            Dim clientes = clienteService.GetClientes("", True)
            For Each c In clientes
                Dim desc = $"{c.Apellido}, {c.Nombre} (DNI: {c.DniCuit})"
                cboClientesFiltro.Items.Add(New ClienteComboItem(c.Id, desc))
            Next

            If cboClientesFiltro.Items.Count > 0 Then
                cboClientesFiltro.SelectedIndex = 0
            End If
        End Sub

        Private Sub ConfigurarColumnasVentas()
            dgvVentas.Columns.Clear()
            dgvVentas.Columns.Add("Id", "ID")
            dgvVentas.Columns("Id").Visible = False

            dgvVentas.Columns.Add("Ticket", "Nº Ticket")
            dgvVentas.Columns("Ticket").Width = 140

            dgvVentas.Columns.Add("Fecha", "Fecha y Hora")
            dgvVentas.Columns("Fecha").Width = 135

            dgvVentas.Columns.Add("Vendedor", "Vendedor")
            dgvVentas.Columns("Vendedor").Width = 140

            dgvVentas.Columns.Add("Cliente", "Cliente")
            dgvVentas.Columns("Cliente").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

            dgvVentas.Columns.Add("Metodo", "Medio Pago")
            dgvVentas.Columns("Metodo").Width = 130

            dgvVentas.Columns.Add("Subtotal", "Subtotal")
            dgvVentas.Columns("Subtotal").Width = 110
            dgvVentas.Columns("Subtotal").DefaultCellStyle.Format = "C2"

            dgvVentas.Columns.Add("Descuento", "Desc. ($)")
            dgvVentas.Columns("Descuento").Width = 90
            dgvVentas.Columns("Descuento").DefaultCellStyle.Format = "C2"

            dgvVentas.Columns.Add("Total", "Total Facturado")
            dgvVentas.Columns("Total").Width = 130
            dgvVentas.Columns("Total").DefaultCellStyle.Format = "C2"
            dgvVentas.Columns("Total").DefaultCellStyle.Font = UITheme.FontBold

            dgvVentas.Columns.Add("Estado", "Estado")
            dgvVentas.Columns("Estado").Width = 110
        End Sub

        Private Sub ConfigurarColumnasClientesReporte()
            dgvClientesReporte.Columns.Clear()
            dgvClientesReporte.Columns.Add("Id", "ID")
            dgvClientesReporte.Columns("Id").Visible = False

            dgvClientesReporte.Columns.Add("Cliente", "Cliente")
            dgvClientesReporte.Columns("Cliente").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

            dgvClientesReporte.Columns.Add("Dni", "DNI / CUIT")
            dgvClientesReporte.Columns("Dni").Width = 115

            dgvClientesReporte.Columns.Add("Telefono", "Teléfono")
            dgvClientesReporte.Columns("Telefono").Width = 110

            dgvClientesReporte.Columns.Add("Compras", "Tickets")
            dgvClientesReporte.Columns("Compras").Width = 85
            dgvClientesReporte.Columns("Compras").DefaultCellStyle.Font = UITheme.FontBold
            dgvClientesReporte.Columns("Compras").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            dgvClientesReporte.Columns.Add("Unidades", "Prendas Adquiridas")
            dgvClientesReporte.Columns("Unidades").Width = 130
            dgvClientesReporte.Columns("Unidades").DefaultCellStyle.Font = UITheme.FontBold
            dgvClientesReporte.Columns("Unidades").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            dgvClientesReporte.Columns.Add("Total", "Total Facturado")
            dgvClientesReporte.Columns("Total").Width = 130
            dgvClientesReporte.Columns("Total").DefaultCellStyle.Format = "C2"
            dgvClientesReporte.Columns("Total").DefaultCellStyle.Font = UITheme.FontBold

            dgvClientesReporte.Columns.Add("Margen", "Margen Ganancia")
            dgvClientesReporte.Columns("Margen").Width = 130
            dgvClientesReporte.Columns("Margen").DefaultCellStyle.Format = "C2"
            dgvClientesReporte.Columns("Margen").DefaultCellStyle.Font = UITheme.FontBold
            dgvClientesReporte.Columns("Margen").DefaultCellStyle.ForeColor = UITheme.ColorSuccessDark

            dgvClientesReporte.Columns.Add("Rentabilidad", "% Margen")
            dgvClientesReporte.Columns("Rentabilidad").Width = 90
            dgvClientesReporte.Columns("Rentabilidad").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            dgvClientesReporte.Columns.Add("UltimaCompra", "Última Compra")
            dgvClientesReporte.Columns("UltimaCompra").Width = 130
        End Sub

        Private Sub ConfigurarColumnasTicketsCliente()
            dgvTicketsCliente.Columns.Clear()
            dgvTicketsCliente.Columns.Add("Id", "ID")
            dgvTicketsCliente.Columns("Id").Visible = False

            dgvTicketsCliente.Columns.Add("Ticket", "Nº Ticket")
            dgvTicketsCliente.Columns("Ticket").Width = 140

            dgvTicketsCliente.Columns.Add("Fecha", "Fecha y Hora")
            dgvTicketsCliente.Columns("Fecha").Width = 135

            dgvTicketsCliente.Columns.Add("Vendedor", "Vendedor")
            dgvTicketsCliente.Columns("Vendedor").Width = 140

            dgvTicketsCliente.Columns.Add("Metodo", "Medio Pago")
            dgvTicketsCliente.Columns("Metodo").Width = 120

            dgvTicketsCliente.Columns.Add("Subtotal", "Subtotal")
            dgvTicketsCliente.Columns("Subtotal").Width = 100
            dgvTicketsCliente.Columns("Subtotal").DefaultCellStyle.Format = "C2"

            dgvTicketsCliente.Columns.Add("Descuento", "Desc. ($)")
            dgvTicketsCliente.Columns("Descuento").Width = 90
            dgvTicketsCliente.Columns("Descuento").DefaultCellStyle.Format = "C2"

            dgvTicketsCliente.Columns.Add("Total", "Total Facturado")
            dgvTicketsCliente.Columns("Total").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            dgvTicketsCliente.Columns("Total").DefaultCellStyle.Format = "C2"
            dgvTicketsCliente.Columns("Total").DefaultCellStyle.Font = UITheme.FontBold

            dgvTicketsCliente.Columns.Add("Estado", "Estado")
            dgvTicketsCliente.Columns("Estado").Width = 110
        End Sub

        Private Sub ConfigurarColumnasRanking()
            dgvRanking.Columns.Clear()
            dgvRanking.Columns.Add("Prenda", "Prenda / Modelo")
            dgvRanking.Columns("Prenda").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

            dgvRanking.Columns.Add("Talle", "Talle")
            dgvRanking.Columns("Talle").Width = 90
            dgvRanking.Columns("Talle").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            dgvRanking.Columns.Add("Cantidad", "Unidades Vendidas")
            dgvRanking.Columns("Cantidad").Width = 140
            dgvRanking.Columns("Cantidad").DefaultCellStyle.Font = UITheme.FontBold
            dgvRanking.Columns("Cantidad").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            dgvRanking.Columns.Add("Recaudado", "Recaudación Total")
            dgvRanking.Columns("Recaudado").Width = 140
            dgvRanking.Columns("Recaudado").DefaultCellStyle.Format = "C2"
            dgvRanking.Columns("Recaudado").DefaultCellStyle.Font = UITheme.FontBold

            dgvRanking.Columns.Add("Costo", "Costo Total")
            dgvRanking.Columns("Costo").Width = 125
            dgvRanking.Columns("Costo").DefaultCellStyle.Format = "C2"

            dgvRanking.Columns.Add("Margen", "Margen Ganancia")
            dgvRanking.Columns("Margen").Width = 140
            dgvRanking.Columns("Margen").DefaultCellStyle.Format = "C2"
            dgvRanking.Columns("Margen").DefaultCellStyle.Font = UITheme.FontBold
            dgvRanking.Columns("Margen").DefaultCellStyle.ForeColor = UITheme.ColorSuccessDark

            dgvRanking.Columns.Add("PctMargen", "% Rentabilidad")
            dgvRanking.Columns("PctMargen").Width = 105
            dgvRanking.Columns("PctMargen").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            If AuthService.IsVendor Then
                dgvRanking.Columns("Costo").Visible = False
                dgvRanking.Columns("Margen").Visible = False
                dgvRanking.Columns("PctMargen").Visible = False
            End If
        End Sub

        ''' <summary>
        ''' Carga sincronizada de todas las pestañas de reporte según el rango de fechas.
        ''' </summary>
        Private Sub LoadReportes()
            LoadVentasGenerales()
            LoadVentasPorCliente()
            LoadRanking()
        End Sub

        Private Sub LoadVentasGenerales()
            Dim desde = dtpDesde.Value.Date
            Dim hasta = dtpHasta.Value.Date

            Dim ventas = ventaService.GetVentas(desde, hasta)
            dgvVentas.Rows.Clear()

            Dim totalPeriodo As Decimal = 0D
            Dim totalEfec As Decimal = 0D
            Dim totalDig As Decimal = 0D
            Dim validTickets As Integer = 0

            For Each v In ventas
                Dim rIdx = dgvVentas.Rows.Add(v.Id, v.NumeroTicket, v.Fecha.ToString("dd/MM/yyyy HH:mm"), v.UsuarioNombre, v.ClienteNombre, v.MetodoPago, v.Subtotal, v.DescuentoMonto, v.Total, v.Estado)
                If v.Estado = "Completada" Then
                    totalPeriodo += v.Total
                    validTickets += 1
                    If v.MetodoPago = "Efectivo" Then
                        totalEfec += v.Total
                    Else
                        totalDig += v.Total
                    End If
                Else
                    dgvVentas.Rows(rIdx).DefaultCellStyle.ForeColor = UITheme.ColorTextMuted
                End If
            Next

            lblTotalVentas.Text = totalPeriodo.ToString("C2")
            lblCantTickets.Text = validTickets.ToString()
            lblEfectivo.Text = totalEfec.ToString("C2")
            lblDigital.Text = totalDig.ToString("C2")
        End Sub

        Private Sub LoadVentasPorCliente()
            Dim desde = dtpDesde.Value.Date
            Dim hasta = dtpHasta.Value.Date

            Dim clienteId As Integer = 0
            If cboClientesFiltro.SelectedItem IsNot Nothing Then
                Dim item = CType(cboClientesFiltro.SelectedItem, ClienteComboItem)
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

            If AuthService.IsVendor Then
                lblResumenClientes.Text = $"Clientes compradores: {totClientes} | Total prendas llevadas: {totPrendas}"
            Else
                lblResumenClientes.Text = $"Clientes: {totClientes} | Prendas: {totPrendas} | Facturado: {totFacturado:C2} | Margen Neto: {totMargen:C2}"
            End If

            DgvClientesReporte_SelectionChanged(Nothing, EventArgs.Empty)
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

            Dim criterio = If(rbMayorMargen.Checked, CriterioRankingProducto.MayorMargen, CriterioRankingProducto.MayorSalida)

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

            Dim top = reporteService.GetTopProductosVendidos(limite, desde, hasta, criterio)
            dgvRanking.Rows.Clear()

            For Each p In top
                Dim pctStr = $"{p.PorcentajeMargen:N1} %"
                dgvRanking.Rows.Add(p.Nombre, p.Talle, p.CantidadVendida, p.TotalRecaudado, p.CostoTotal, p.MargenGanancia, pctStr)
            Next
        End Sub

        Private Sub DgvVentas_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
            If e.RowIndex >= 0 Then
                Dim row = dgvVentas.Rows(e.RowIndex)
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

        Private Sub MostrarModalDetalleVenta(ventaId As Integer, ticketNumero As String)
            Dim detalles = ventaService.GetDetalleVenta(ventaId)
            If detalles Is Nothing OrElse detalles.Count = 0 Then
                MessageBox.Show("No se encontraron renglones de artículos para este ticket.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim modal As New Form() With {
                .Text = $"Detalle del Comprobante #{ticketNumero}",
                .Size = New Size(720, 420),
                .StartPosition = FormStartPosition.CenterParent,
                .FormBorderStyle = FormBorderStyle.FixedDialog,
                .MaximizeBox = False,
                .MinimizeBox = False,
                .BackColor = UITheme.ColorBackground
            }

            Dim pnlHead As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 50,
                .BackColor = UITheme.ColorSecondary,
                .Padding = New Padding(15, 12, 15, 12)
            }
            Dim lblHead As New Label() With {
                .Text = $"Prendas adquiridas en Ticket #{ticketNumero}",
                .Font = UITheme.FontSubheading,
                .ForeColor = Color.White,
                .AutoSize = True,
                .Location = New Point(15, 12)
            }
            pnlHead.Controls.Add(lblHead)

            Dim dgvModal As New DataGridView() With {.Dock = DockStyle.Fill}
            UITheme.StyleDataGrid(dgvModal)
            dgvModal.Columns.Add("Codigo", "Código")
            dgvModal.Columns("Codigo").Width = 110
            dgvModal.Columns.Add("Articulo", "Prenda / Modelo")
            dgvModal.Columns("Articulo").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            dgvModal.Columns.Add("Talle", "Talle")
            dgvModal.Columns("Talle").Width = 80
            dgvModal.Columns.Add("Color", "Color")
            dgvModal.Columns("Color").Width = 80
            dgvModal.Columns.Add("Precio", "Precio Unit.")
            dgvModal.Columns("Precio").Width = 110
            dgvModal.Columns("Precio").DefaultCellStyle.Format = "C2"
            dgvModal.Columns.Add("Cantidad", "Cant.")
            dgvModal.Columns("Cantidad").Width = 70
            dgvModal.Columns("Cantidad").DefaultCellStyle.Font = UITheme.FontBold
            dgvModal.Columns("Cantidad").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvModal.Columns.Add("Subtotal", "Subtotal")
            dgvModal.Columns("Subtotal").Width = 110
            dgvModal.Columns("Subtotal").DefaultCellStyle.Format = "C2"
            dgvModal.Columns("Subtotal").DefaultCellStyle.Font = UITheme.FontBold

            For Each d In detalles
                dgvModal.Rows.Add(d.CodigoBarra, d.DescripcionArticulo, d.TalleNombre, d.Color, d.PrecioUnitario, d.Cantidad, d.Subtotal)
            Next

            Dim pnlBottom As New Panel() With {
                .Dock = DockStyle.Bottom,
                .Height = 50,
                .BackColor = UITheme.ColorSurface
            }
            Dim btnCerrar As New Button() With {
                .Text = "Cerrar",
                .Size = New Size(100, 32),
                .Location = New Point(modal.ClientSize.Width - 120, 9),
                .Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
            }
            UITheme.StyleButton(btnCerrar, "Secondary")
            AddHandler btnCerrar.Click, Sub() modal.Close()
            pnlBottom.Controls.Add(btnCerrar)

            modal.Controls.Add(dgvModal)
            modal.Controls.Add(pnlBottom)
            modal.Controls.Add(pnlHead)

            modal.ShowDialog(Me)
        End Sub

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
                    Dim errMsg As String = ""
                    If ventaService.AnularVenta(ventaId, userId, txtMot.Text.Trim(), errMsg) Then
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
