' ARCHIVO: FrmMisVentasTurno.vb
' PROPÓSITO: Historial de ventas del turno actual exclusivo para el Vendedor.
' Permite consultar únicamente los tickets emitidos por el vendedor en su caja abierta,
' reimprimir comprobantes en PDF y gestionar cambios de prenda por talle o color,
' ocultando costos de mercadería, márgenes de ganancia, rankings y ventas de otros turnos/vendedores.

Imports System.Drawing
Imports System.Windows.Forms
Imports GestionComercial.Models
Imports GestionComercial.Services
Imports GestionComercial.UI

Namespace Forms
    Public Class FrmMisVentasTurno
        Inherits Form

        Private ReadOnly cajaService As New CajaService()
        Private ReadOnly ventaService As New VentaService()
        Private ReadOnly catalogService As New CatalogService()
        Private cajaActual As Caja = Nothing

        Private txtBuscarTicket As TextBox
        Private dgvMisVentas As DataGridView
        Private lblResumenMisVentas As Label
        Private listaVentasTurno As List(Of Venta) = New List(Of Venta)()

        ' Constructor: inicializa la vista y carga las ventas del turno o rango seleccionado
        Public Sub New()
            InitializeUI()
            CargarVentasDelTurno()
        End Sub

        ' Construye filtros de búsqueda, grilla de ventas del turno y botones de ticket/cambio
        Private Sub InitializeUI()
            Me.Text = "Mis Ventas del Turno Actual"
            Me.Size = New Size(1020, 640)
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.BackColor = UITheme.ColorBackground
            Me.Font = UITheme.FontRegular

            ' Header superior
            Dim pnlHeader As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 60,
                .BackColor = UITheme.ColorSecondary,
                .Padding = New Padding(20, 15, 20, 15)
            }
            Dim lblTitle As New Label() With {
                .Text = "Mis ventas del turno — Consulta de tickets, reimpresión y cambios",
                .Font = UITheme.FontHeading,
                .ForeColor = Color.White,
                .AutoSize = True,
                .Location = New Point(15, 15)
            }
            pnlHeader.Controls.Add(lblTitle)

            ' Barra de búsqueda y acciones responsiva
            Dim pnlTopVentas As New Panel() With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .MinimumSize = New Size(0, 60),
                .BackColor = UITheme.ColorSurface,
                .Padding = New Padding(15, 12, 15, 8)
            }

            Dim flpVentasActions As New FlowLayoutPanel() With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .WrapContents = True,
                .BackColor = Color.Transparent
            }

            Dim lblBusq As New Label() With {
                .Text = "Buscar ticket / cliente / DNI:",
                .Font = UITheme.FontBold,
                .AutoSize = True,
                .Margin = New Padding(0, 8, 8, 4)
            }

            txtBuscarTicket = New TextBox() With {
                .Width = 230,
                .Margin = New Padding(0, 4, 12, 4),
                .PlaceholderText = "Nº Ticket, Nombre o DNI/CUIT..."
            }
            UITheme.StyleTextBox(txtBuscarTicket)
            AddHandler txtBuscarTicket.TextChanged, Sub() CargarVentasDelTurno()

            Dim btnVerDetalle As New Button() With {
                .Text = "🔍 Ver Detalle",
                .Size = New Size(135, 34),
                .Margin = New Padding(0, 2, 8, 4)
            }
            UITheme.StyleButton(btnVerDetalle, "Secondary")
            AddHandler btnVerDetalle.Click, AddressOf BtnVerDetalleTicket_Click

            Dim btnReimprimirPdf As New Button() With {
                .Text = "📄 Reimprimir / PDF",
                .Size = New Size(165, 34),
                .Margin = New Padding(0, 2, 8, 4)
            }
            UITheme.StyleButton(btnReimprimirPdf, "Primary")
            AddHandler btnReimprimirPdf.Click, AddressOf BtnReimprimirPdfTicket_Click

            Dim btnCambioPrenda As New Button() With {
                .Text = "🔄 Cambio por Talle / Color",
                .Size = New Size(215, 34),
                .Margin = New Padding(0, 2, 8, 4)
            }
            UITheme.StyleButton(btnCambioPrenda, "Success")
            AddHandler btnCambioPrenda.Click, AddressOf BtnCambioPrenda_Click

            Dim btnRefrescar As New Button() With {
                .Text = "Actualizar",
                .Size = New Size(110, 34),
                .Margin = New Padding(0, 2, 12, 4)
            }
            UITheme.StyleButton(btnRefrescar, "Secondary")
            AddHandler btnRefrescar.Click, Sub() CargarVentasDelTurno()

            lblResumenMisVentas = New Label() With {
                .Text = "Tickets de tu turno actual: 0",
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorTextSecondary,
                .AutoSize = True,
                .Margin = New Padding(4, 9, 0, 4)
            }

            flpVentasActions.Controls.AddRange({lblBusq, txtBuscarTicket, btnVerDetalle, btnReimprimirPdf, btnCambioPrenda, btnRefrescar, lblResumenMisVentas})
            pnlTopVentas.Controls.Add(flpVentasActions)

            ' Grilla de tickets del turno actual
            Dim pnlGridVentas As New Panel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(15)
            }
            dgvMisVentas = New DataGridView() With {.Dock = DockStyle.Fill}
            UITheme.StyleDataGrid(dgvMisVentas)
            ConfigurarColumnasMisVentas()
            AddHandler dgvMisVentas.CellDoubleClick, Sub(s, ev)
                                                         If ev.RowIndex >= 0 Then
                                                             BtnVerDetalleTicket_Click(s, EventArgs.Empty)
                                                         End If
                                                     End Sub
            pnlGridVentas.Controls.Add(dgvMisVentas)

            ' Orden de Docking
            Me.Controls.Add(pnlGridVentas)
            Me.Controls.Add(pnlTopVentas)
            Me.Controls.Add(pnlHeader)
        End Sub

        ' Configura las columnas y formatos de la grilla de ventas realizadas
        Private Sub ConfigurarColumnasMisVentas()
            dgvMisVentas.Columns.Clear()
            dgvMisVentas.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            dgvMisVentas.Columns.Add("Id", "ID")
            dgvMisVentas.Columns("Id").Visible = False

            dgvMisVentas.Columns.Add("Ticket", "Nº Ticket")
            dgvMisVentas.Columns("Ticket").FillWeight = 13
            dgvMisVentas.Columns("Ticket").MinimumWidth = 90
            dgvMisVentas.Columns("Ticket").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvMisVentas.Columns("Ticket").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            dgvMisVentas.Columns.Add("Fecha", "Hora Emisión")
            dgvMisVentas.Columns("Fecha").FillWeight = 13
            dgvMisVentas.Columns("Fecha").MinimumWidth = 95
            dgvMisVentas.Columns("Fecha").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvMisVentas.Columns("Fecha").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            dgvMisVentas.Columns.Add("Cliente", "Cliente")
            dgvMisVentas.Columns("Cliente").FillWeight = 22
            dgvMisVentas.Columns("Cliente").MinimumWidth = 130

            dgvMisVentas.Columns.Add("Comprobante", "Comprobante")
            dgvMisVentas.Columns("Comprobante").FillWeight = 12
            dgvMisVentas.Columns("Comprobante").MinimumWidth = 90
            dgvMisVentas.Columns("Comprobante").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvMisVentas.Columns("Comprobante").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            dgvMisVentas.Columns.Add("MetodoPago", "Medio de Pago")
            dgvMisVentas.Columns("MetodoPago").FillWeight = 14
            dgvMisVentas.Columns("MetodoPago").MinimumWidth = 100

            dgvMisVentas.Columns.Add("Total", "Total Cobrado ($)")
            dgvMisVentas.Columns("Total").FillWeight = 13
            dgvMisVentas.Columns("Total").MinimumWidth = 95
            dgvMisVentas.Columns("Total").DefaultCellStyle.Format = "C2"
            dgvMisVentas.Columns("Total").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvMisVentas.Columns("Total").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight

            dgvMisVentas.Columns.Add("Estado", "Estado / Cambios")
            dgvMisVentas.Columns("Estado").FillWeight = 18
            dgvMisVentas.Columns("Estado").MinimumWidth = 110
        End Sub

        ' Carga y filtra las ventas de la caja actual (o del vendedor) y actualiza los totales
        Private Sub CargarVentasDelTurno()
            If dgvMisVentas Is Nothing Then Return
            dgvMisVentas.Rows.Clear()

            Dim userId As Integer = If(AuthService.CurrentUser IsNot Nothing, AuthService.CurrentUser.Id, 1)
            cajaActual = cajaService.GetCajaAbierta(userId)

            If cajaActual Is Nothing Then
                listaVentasTurno.Clear()
                If lblResumenMisVentas IsNot Nothing Then
                    lblResumenMisVentas.Text = "Sin caja abierta — Abra su caja en 'Caja diaria y arqueo' para ver y operar sus tickets del turno."
                    lblResumenMisVentas.ForeColor = UITheme.ColorWarning
                End If
                Return
            End If

            Dim filtro As String = If(txtBuscarTicket IsNot Nothing, txtBuscarTicket.Text.Trim(), "")
            listaVentasTurno = ventaService.GetVentasPorTurnoVendedor(cajaActual.Id, userId, filtro)

            Dim cantActivas As Integer = 0
            For Each v In listaVentasTurno
                Dim estadoTexto As String = v.Estado
                If Not String.IsNullOrWhiteSpace(v.Observaciones) AndAlso v.Observaciones.IndexOf("CAMBIO", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    estadoTexto &= " · Con Cambio"
                End If
                Dim idx = dgvMisVentas.Rows.Add(v.Id, v.NumeroTicket, v.Fecha.ToString("dd/MM HH:mm"), v.ClienteNombre, v.TipoComprobante, v.MetodoPago, v.Total, estadoTexto)
                If v.Estado.Equals("Anulada", StringComparison.OrdinalIgnoreCase) Then
                    dgvMisVentas.Rows(idx).DefaultCellStyle.ForeColor = UITheme.ColorDanger
                Else
                    cantActivas += 1
                End If
            Next

            If lblResumenMisVentas IsNot Nothing Then
                lblResumenMisVentas.Text = $"Turno #{cajaActual.Id}  ·  Tickets emitidos en tu turno: {cantActivas}"
                lblResumenMisVentas.ForeColor = UITheme.ColorTextSecondary
            End If
        End Sub

        ' Devuelve el objeto Venta correspondiente a la fila seleccionada en la grilla
        Private Function ObtenerVentaSeleccionada() As Venta
            If dgvMisVentas Is Nothing OrElse dgvMisVentas.SelectedRows.Count = 0 Then
                MessageBox.Show("Seleccione un ticket de la grilla.", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return Nothing
            End If
            Dim idSel As Integer = Convert.ToInt32(dgvMisVentas.SelectedRows(0).Cells("Id").Value)
            Return listaVentasTurno.Find(Function(x) x.Id = idSel)
        End Function

        ' Abre el detalle del ticket seleccionado si hay una venta elegida
        Private Sub BtnVerDetalleTicket_Click(sender As Object, e As EventArgs)
            Dim v = ObtenerVentaSeleccionada()
            If v Is Nothing Then Return
            MostrarDetalleTicketVendedor(v)
        End Sub

        ' Reimprime o exporta a PDF el comprobante de la venta seleccionada
        Private Sub BtnReimprimirPdfTicket_Click(sender As Object, e As EventArgs)
            Dim v = ObtenerVentaSeleccionada()
            If v Is Nothing Then Return
            v.Detalles = ventaService.GetDetalleVenta(v.Id)
            FrmReportes.GuardarTicketEnPdf(Me, v)
        End Sub

        ' Abre el modal de cambio de prenda o devolución para la venta seleccionada
        Private Sub BtnCambioPrenda_Click(sender As Object, e As EventArgs)
            Dim v = ObtenerVentaSeleccionada()
            If v Is Nothing Then Return
            If v.Estado.Equals("Anulada", StringComparison.OrdinalIgnoreCase) Then
                MessageBox.Show("No se puede gestionar un cambio sobre un ticket anulado.", "Operación no disponible", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            MostrarModalCambioPrenda(v)
        End Sub

        ' Muestra un modal con el detalle completo de ítems, totales y datos del ticket
        Private Sub MostrarDetalleTicketVendedor(ventaSel As Venta)
            Dim detalles = ventaService.GetDetalleVenta(ventaSel.Id)
            ventaSel.Detalles = detalles

            Dim frmDet As New Form() With {
                .Text = $"Detalle de Ticket #{ventaSel.NumeroTicket}",
                .Size = New Size(760, 500),
                .MinimumSize = New Size(640, 420),
                .StartPosition = FormStartPosition.CenterParent,
                .FormBorderStyle = FormBorderStyle.Sizable,
                .MaximizeBox = True,
                .MinimizeBox = False,
                .BackColor = UITheme.ColorBackground,
                .Font = UITheme.FontRegular
            }

            Dim pnlTop As New Panel() With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .MinimumSize = New Size(0, 78),
                .BackColor = UITheme.ColorSurface,
                .Padding = New Padding(18, 12, 18, 12)
            }

            Dim lblInfo As New Label() With {
                .Text = $"Ticket #{ventaSel.NumeroTicket}   |   Fecha: {ventaSel.Fecha:dd/MM/yyyy HH:mm}   |   Cliente: {ventaSel.ClienteNombre}" & Environment.NewLine &
                        $"Comprobante: {ventaSel.TipoComprobante}   |   Medio de Pago: {ventaSel.MetodoPago}   |   Total Cobrado: {ventaSel.Total:C2}",
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorTextPrimary,
                .Dock = DockStyle.Top,
                .AutoSize = True
            }
            pnlTop.Controls.Add(lblInfo)

            If Not String.IsNullOrWhiteSpace(ventaSel.Observaciones) Then
                Dim lblNotas As New Label() With {
                    .Text = $"Notas / Cambios registrados: {ventaSel.Observaciones}",
                    .Font = UITheme.FontSmall,
                    .ForeColor = UITheme.ColorPrimaryDark,
                    .Dock = DockStyle.Bottom,
                    .AutoSize = True,
                    .Padding = New Padding(0, 6, 0, 0)
                }
                pnlTop.Controls.Add(lblNotas)
            End If

            Dim pnlGrid As New Panel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(18, 12, 18, 12)
            }
            Dim dgvDet As New DataGridView() With {.Dock = DockStyle.Fill}
            UITheme.StyleDataGrid(dgvDet)
            dgvDet.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            dgvDet.Columns.Add("Codigo", "Código")
            dgvDet.Columns("Codigo").FillWeight = 16
            dgvDet.Columns("Codigo").MinimumWidth = 95

            dgvDet.Columns.Add("Prenda", "Prenda / Artículo")
            dgvDet.Columns("Prenda").FillWeight = 36
            dgvDet.Columns("Prenda").MinimumWidth = 160

            dgvDet.Columns.Add("Talle", "Talle")
            dgvDet.Columns("Talle").FillWeight = 10
            dgvDet.Columns("Talle").MinimumWidth = 65
            dgvDet.Columns("Talle").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvDet.Columns("Talle").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            dgvDet.Columns.Add("Color", "Color")
            dgvDet.Columns("Color").FillWeight = 12
            dgvDet.Columns("Color").MinimumWidth = 75
            dgvDet.Columns("Color").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvDet.Columns("Color").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            dgvDet.Columns.Add("Cant", "Cant.")
            dgvDet.Columns("Cant").FillWeight = 8
            dgvDet.Columns("Cant").MinimumWidth = 55
            dgvDet.Columns("Cant").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvDet.Columns("Cant").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            dgvDet.Columns.Add("PrecioUnit", "Precio Unit.")
            dgvDet.Columns("PrecioUnit").FillWeight = 14
            dgvDet.Columns("PrecioUnit").MinimumWidth = 90
            dgvDet.Columns("PrecioUnit").DefaultCellStyle.Format = "C2"
            dgvDet.Columns("PrecioUnit").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvDet.Columns("PrecioUnit").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight

            dgvDet.Columns.Add("Subtotal", "Subtotal")
            dgvDet.Columns("Subtotal").FillWeight = 14
            dgvDet.Columns("Subtotal").MinimumWidth = 90
            dgvDet.Columns("Subtotal").DefaultCellStyle.Format = "C2"
            dgvDet.Columns("Subtotal").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvDet.Columns("Subtotal").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight

            For Each d In detalles
                dgvDet.Rows.Add(d.CodigoBarra, d.DescripcionArticulo, d.TalleNombre, d.Color, d.Cantidad, d.PrecioUnitario, d.Subtotal)
            Next
            pnlGrid.Controls.Add(dgvDet)

            Dim pnlBottom As New Panel() With {
                .Dock = DockStyle.Bottom,
                .Height = 56,
                .BackColor = UITheme.ColorSurface,
                .Padding = New Padding(18, 10, 18, 10)
            }

            Dim btnPdf As New Button() With {.Text = "📄 Reimprimir / Guardar PDF", .Size = New Size(210, 36), .Location = New Point(18, 10)}
            UITheme.StyleButton(btnPdf, "Primary")
            AddHandler btnPdf.Click, Sub() FrmReportes.GuardarTicketEnPdf(frmDet, ventaSel)

            Dim btnCambio As New Button() With {
                .Text = "🔄 Cambio por Talle / Color",
                .Size = New Size(210, 36),
                .Location = New Point(238, 10),
                .Enabled = Not ventaSel.Estado.Equals("Anulada", StringComparison.OrdinalIgnoreCase)
            }
            UITheme.StyleButton(btnCambio, "Success")
            AddHandler btnCambio.Click, Sub()
                                            frmDet.Close()
                                            MostrarModalCambioPrenda(ventaSel)
                                        End Sub

            Dim btnCerrar As New Button() With {.Text = "Cerrar", .Size = New Size(110, 36), .Anchor = AnchorStyles.Top Or AnchorStyles.Right, .Location = New Point(frmDet.ClientSize.Width - 128, 10)}
            UITheme.StyleButton(btnCerrar, "Secondary")
            AddHandler btnCerrar.Click, Sub() frmDet.Close()

            pnlBottom.Controls.AddRange({btnPdf, btnCambio, btnCerrar})

            frmDet.Controls.Add(pnlGrid)
            frmDet.Controls.Add(pnlBottom)
            frmDet.Controls.Add(pnlTop)
            frmDet.ShowDialog(Me)
        End Sub

        ' Presenta un modal para gestionar el cambio de talle/prenda y registrar diferencias en caja
        Private Sub MostrarModalCambioPrenda(ventaSel As Venta)
            Dim detalles = ventaService.GetDetalleVenta(ventaSel.Id)
            If detalles.Count = 0 Then
                MessageBox.Show("El ticket seleccionado no posee prendas asociadas.", "Sin ítems", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim frmCambio As New Form() With {
                .Text = $"Gestión de Cambio por Talle / Color — Ticket #{ventaSel.NumeroTicket}",
                .Size = New Size(560, 410),
                .StartPosition = FormStartPosition.CenterParent,
                .FormBorderStyle = FormBorderStyle.FixedDialog,
                .MaximizeBox = False,
                .MinimizeBox = False,
                .BackColor = Color.White,
                .Font = UITheme.FontRegular
            }

            Dim lblItem As New Label() With {
                .Text = "1. Seleccione la prenda del ticket que el cliente devuelve:",
                .Font = UITheme.FontBold,
                .Location = New Point(25, 20),
                .AutoSize = True
            }

            Dim cboItemTicket As New ComboBox() With {
                .Location = New Point(25, 45),
                .Size = New Size(495, 28),
                .DropDownStyle = ComboBoxStyle.DropDownList
            }
            UITheme.StyleComboBox(cboItemTicket)

            For Each d In detalles
                cboItemTicket.Items.Add(New With {
                    .Detalle = d,
                    .Display = $"{d.DescripcionArticulo} | Cant: {d.Cantidad} | Actual: Talle {d.TalleNombre} - Color {d.Color}"
                })
            Next
            cboItemTicket.DisplayMember = "Display"

            Dim lblVariante As New Label() With {
                .Text = "2. Seleccione el nuevo Talle / Color a entregar (con stock disponible):",
                .Font = UITheme.FontBold,
                .Location = New Point(25, 90),
                .AutoSize = True
            }

            Dim cboNuevaVariante As New ComboBox() With {
                .Location = New Point(25, 115),
                .Size = New Size(495, 28),
                .DropDownStyle = ComboBoxStyle.DropDownList
            }
            UITheme.StyleComboBox(cboNuevaVariante)

            Dim lblInfoStock As New Label() With {
                .Text = "",
                .Font = UITheme.FontSmall,
                .ForeColor = UITheme.ColorTextSecondary,
                .Location = New Point(25, 148),
                .AutoSize = True
            }

            AddHandler cboItemTicket.SelectedIndexChanged, Sub()
                                                               cboNuevaVariante.Items.Clear()
                                                               If cboItemTicket.SelectedItem Is Nothing Then Return
                                                               Dim detSel As DetalleVenta = CType(CallByName(cboItemTicket.SelectedItem, "Detalle", CallType.Get), DetalleVenta)
                                                               Dim variantes = catalogService.GetStockPorTalles(detSel.ProductoId)
                                                               For Each pt In variantes
                                                                   Dim esMisma As Boolean = (pt.TalleId = detSel.TalleId AndAlso String.Equals(pt.Color, detSel.Color, StringComparison.OrdinalIgnoreCase))
                                                                   If Not esMisma AndAlso pt.StockActual >= detSel.Cantidad Then
                                                                       cboNuevaVariante.Items.Add(New With {
                                                                           .Variante = pt,
                                                                           .Display = $"Talle: {pt.TalleNombre}   |   Color: {pt.Color}   (Stock disponible: {pt.StockActual} u.)"
                                                                       })
                                                                   End If
                                                               Next
                                                               cboNuevaVariante.DisplayMember = "Display"
                                                               If cboNuevaVariante.Items.Count > 0 Then
                                                                   cboNuevaVariante.SelectedIndex = 0
                                                                   lblInfoStock.Text = $"Se reingresará {detSel.Cantidad} u. de [T.{detSel.TalleNombre} {detSel.Color}] y se descontará del nuevo talle/color."
                                                                   lblInfoStock.ForeColor = UITheme.ColorSuccess
                                                               Else
                                                                   lblInfoStock.Text = "No hay otros talles/colores de este artículo con stock suficiente para el cambio."
                                                                   lblInfoStock.ForeColor = UITheme.ColorDanger
                                                               End If
                                                           End Sub

            If cboItemTicket.Items.Count > 0 Then
                cboItemTicket.SelectedIndex = 0
            End If

            Dim lblMotivo As New Label() With {
                .Text = "3. Motivo / Observación del cambio:",
                .Font = UITheme.FontBold,
                .Location = New Point(25, 185),
                .AutoSize = True
            }

            Dim txtMotivo As New TextBox() With {
                .Location = New Point(25, 210),
                .Size = New Size(495, 60),
                .Multiline = True,
                .Text = "Cambio de prenda por talle / color solicitado por el cliente"
            }
            UITheme.StyleTextBox(txtMotivo)

            Dim btnConfirmar As New Button() With {
                .Text = "Confirmar Cambio de Prenda",
                .Location = New Point(25, 295),
                .Size = New Size(340, 42)
            }
            UITheme.StyleButton(btnConfirmar, "Success")

            Dim btnCancelar As New Button() With {
                .Text = "Cancelar",
                .Location = New Point(380, 295),
                .Size = New Size(140, 42)
            }
            UITheme.StyleButton(btnCancelar, "Secondary")
            AddHandler btnCancelar.Click, Sub() frmCambio.Close()

            AddHandler btnConfirmar.Click, Sub()
                                               If cboItemTicket.SelectedItem Is Nothing Then
                                                   MessageBox.Show("Seleccione la prenda a cambiar.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                                   Return
                                               End If
                                               If cboNuevaVariante.SelectedItem Is Nothing Then
                                                   MessageBox.Show("Seleccione un nuevo talle/color con stock disponible.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                                   Return
                                               End If

                                               Dim detSel As DetalleVenta = CType(CallByName(cboItemTicket.SelectedItem, "Detalle", CallType.Get), DetalleVenta)
                                               Dim varNueva As ProductoTalle = CType(CallByName(cboNuevaVariante.SelectedItem, "Variante", CallType.Get), ProductoTalle)
                                               Dim userId As Integer = If(AuthService.CurrentUser IsNot Nothing, AuthService.CurrentUser.Id, 1)
                                               Dim errMsg As String = ""

                                               If ventaService.ProcesarCambioPrenda(ventaSel.Id, detSel.Id, varNueva.TalleId, varNueva.Color, userId, txtMotivo.Text.Trim(), errMsg) Then
                                                   MessageBox.Show($"Cambio registrado exitosamente.{Environment.NewLine}Se reingresó Talle {detSel.TalleNombre} ({detSel.Color}) y se entregó Talle {varNueva.TalleNombre} ({varNueva.Color}).", "Cambio Completado", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                                   frmCambio.DialogResult = DialogResult.OK
                                                   frmCambio.Close()
                                               Else
                                                   MessageBox.Show("No se pudo procesar el cambio: " & errMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                               End If
                                           End Sub

            frmCambio.Controls.AddRange({lblItem, cboItemTicket, lblVariante, cboNuevaVariante, lblInfoStock, lblMotivo, txtMotivo, btnConfirmar, btnCancelar})
            If frmCambio.ShowDialog(Me) = DialogResult.OK Then
                CargarVentasDelTurno()
            End If
        End Sub

    End Class
End Namespace

