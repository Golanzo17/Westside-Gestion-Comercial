Imports System.Drawing
Imports System.Windows.Forms
Imports GestionComercial.Models
Imports GestionComercial.Services
Imports GestionComercial.UI

Namespace Forms
    Public Class FrmReportes
        Inherits Form

        Private ventaService As New VentaService()
        Private reporteService As New ReporteService()

        Private dtpDesde As DateTimePicker
        Private dtpHasta As DateTimePicker
        Private btnFiltrar As Button
        Private btnAnularVenta As Button

        Private lblTotalVentas As Label
        Private lblCantTickets As Label
        Private lblEfectivo As Label
        Private lblDigital As Label

        Private dgvVentas As DataGridView
        Private dgvRanking As DataGridView

        Public Sub New()
            InitializeUI()
            LoadReportes()
        End Sub

        Private Sub InitializeUI()
            Me.Text = "Reportes Comerciales y Estadísticas de Ventas"
            Me.Size = New Size(1060, 680)
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.BackColor = UITheme.ColorBackground
            Me.Font = UITheme.FontRegular

            ' Header
            Dim pnlHeader As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 60,
                .BackColor = UITheme.ColorSecondary,
                .Padding = New Padding(20, 15, 20, 15)
            }
            Dim lblTitle As New Label() With {
                .Text = "Reportes de ventas y rendimiento comercial",
                .Font = UITheme.FontHeading,
                .ForeColor = Color.White,
                .AutoSize = True,
                .Location = New Point(15, 15)
            }
            pnlHeader.Controls.Add(lblTitle)

            ' Toolbar de Fechas
            Dim pnlFilter As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 55,
                .BackColor = UITheme.ColorSurface,
                .Padding = New Padding(15, 12, 15, 10)
            }

            Dim lblD As New Label() With {.Text = "Desde:", .Font = UITheme.FontBold, .Location = New Point(15, 16), .AutoSize = True}
            dtpDesde = New DateTimePicker() With {.Location = New Point(70, 14), .Size = New Size(130, 26), .Format = DateTimePickerFormat.Short, .Value = DateTime.Today.AddDays(-7)}

            Dim lblH As New Label() With {.Text = "Hasta:", .Font = UITheme.FontBold, .Location = New Point(220, 16), .AutoSize = True}
            dtpHasta = New DateTimePicker() With {.Location = New Point(275, 14), .Size = New Size(130, 26), .Format = DateTimePickerFormat.Short, .Value = DateTime.Today}

            btnFiltrar = New Button() With {.Text = "Generar reporte", .Location = New Point(430, 12), .Size = New Size(160, 30)}
            UITheme.StyleButton(btnFiltrar, "Primary")
            AddHandler btnFiltrar.Click, Sub() LoadReportes()

            btnAnularVenta = New Button() With {.Text = "Anular venta seleccionada", .Location = New Point(610, 12), .Size = New Size(220, 30)}
            UITheme.StyleButton(btnAnularVenta, "Danger")
            btnAnularVenta.Visible = AuthService.IsAdminOrManager
            AddHandler btnAnularVenta.Click, AddressOf BtnAnularVenta_Click

            pnlFilter.Controls.AddRange({lblD, dtpDesde, lblH, dtpHasta, btnFiltrar, btnAnularVenta})

            ' Tarjetas KPI
            Dim pnlKpis As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 105,
                .BackColor = UITheme.ColorSurfaceMuted,
                .Padding = New Padding(15, 5, 15, 5)
            }

            Dim cardTot = UITheme.CreateKpiCard("TOTAL FACTURADO", "$ 0,00", "En el período seleccionado", UITheme.ColorSuccess)
            cardTot.Location = New Point(15, 5)
            lblTotalVentas = CType(cardTot.Controls(2), Label)

            Dim cardCant = UITheme.CreateKpiCard("CANTIDAD VENTAS", "0", "Tickets emitidos", UITheme.ColorPrimary)
            cardCant.Location = New Point(255, 5)
            lblCantTickets = CType(cardCant.Controls(2), Label)

            Dim cardEf = UITheme.CreateKpiCard("COBRADO EN EFECTIVO", "$ 0,00", "Total en billetes", UITheme.ColorWarning)
            cardEf.Location = New Point(495, 5)
            lblEfectivo = CType(cardEf.Controls(2), Label)

            Dim cardDig = UITheme.CreateKpiCard("COBRADO DIGITAL", "$ 0,00", "Tarjetas y transferencias", UITheme.ColorInfo)
            cardDig.Location = New Point(735, 5)
            lblDigital = CType(cardDig.Controls(2), Label)

            pnlKpis.Controls.AddRange({cardTot, cardCant, cardEf, cardDig})

            ' TabControl con Historial de Ventas y Ranking de Ropa
            Dim tabs As New TabControl() With {
                .Dock = DockStyle.Fill,
                .Padding = New Point(12, 6)
            }

            ' Tab 1: Ventas
            Dim tabVentas As New TabPage("Listado de ventas y tickets emitidos") With {.BackColor = UITheme.ColorBackground}
            dgvVentas = New DataGridView() With {.Dock = DockStyle.Fill}
            UITheme.StyleDataGrid(dgvVentas)
            ConfigurarColumnasVentas()
            tabVentas.Controls.Add(dgvVentas)
            tabs.TabPages.Add(tabVentas)

            ' Tab 2: Ranking Prendas
            Dim tabRanking As New TabPage("Ranking de prendas más vendidas") With {.BackColor = UITheme.ColorBackground}
            dgvRanking = New DataGridView() With {.Dock = DockStyle.Fill}
            UITheme.StyleDataGrid(dgvRanking)
            ConfigurarColumnasRanking()
            tabRanking.Controls.Add(dgvRanking)
            tabs.TabPages.Add(tabRanking)

            If AuthService.IsVendor Then
                tabs.TabPages.Remove(tabVentas)
                pnlKpis.Visible = False
                btnAnularVenta.Visible = False
            End If

            ' Orden exacto de Docking: Fill primero, KPIs segundo, Filter tercero, Header último
            Me.Controls.Add(tabs)
            Me.Controls.Add(pnlKpis)
            Me.Controls.Add(pnlFilter)
            Me.Controls.Add(pnlHeader)
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

        Private Sub ConfigurarColumnasRanking()
            dgvRanking.Columns.Clear()
            dgvRanking.Columns.Add("Prenda", "Prenda / Modelo")
            dgvRanking.Columns("Prenda").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

            dgvRanking.Columns.Add("Talle", "Talle")
            dgvRanking.Columns("Talle").Width = 100

            dgvRanking.Columns.Add("Cantidad", "Unidades Vendidas")
            dgvRanking.Columns("Cantidad").Width = 150
            dgvRanking.Columns("Cantidad").DefaultCellStyle.Font = UITheme.FontBold

            dgvRanking.Columns.Add("Recaudado", "Recaudación Total")
            dgvRanking.Columns("Recaudado").Width = 160
            dgvRanking.Columns("Recaudado").DefaultCellStyle.Format = "C2"
            dgvRanking.Columns("Recaudado").DefaultCellStyle.Font = UITheme.FontBold
        End Sub

        Private Sub LoadReportes()
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

            ' Cargar Ranking
            Dim top = reporteService.GetTopProductosVendidos(20, desde, hasta)
            dgvRanking.Rows.Clear()
            For Each p In top
                dgvRanking.Rows.Add(p.Nombre, p.Talle, p.CantidadVendida, p.TotalRecaudado)
            Next
        End Sub

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
