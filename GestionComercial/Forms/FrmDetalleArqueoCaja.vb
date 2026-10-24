' ARCHIVO: FrmDetalleArqueoCaja.vb
' PROPÓSITO: Ventana de auditoría profunda e inspección de arqueo de un turno de caja específico.
' Permite al Gerente supervisar:
' 1. La fórmula matemática exacta de conciliación de efectivo (Inicial + Ventas - Gastos = Esperado vs Declarado).
' 2. El desglose de cobros discriminados por cada medio de pago (Efectivo, Débito, Crédito, Transferencia, QR).
' 3. La lista exhaustiva de tickets y comprobantes de venta emitidos durante el turno.
' 4. Los movimientos manuales de efectivo (egresos por gastos autorizados, retiros, ingresos extraordinarios).
' 5. Emisión de Acta Oficial de Arqueo y Cierre para firma física del Vendedor y Gerente.

Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports GestionComercial.Models
Imports GestionComercial.Services
Imports GestionComercial.UI

Namespace Forms
    Public Class FrmDetalleArqueoCaja
        Inherits Form

        Private ReadOnly _cajaId As Integer
        Private ReadOnly _cajaService As New CajaService()
        Private ReadOnly _configService As New ConfiguracionService()
        Private _caja As Caja

        ' Controles de UI
        Private lblTurnoHeader As Label
        Private lblEstadoBadge As Label
        Private lblSubtitulo As Label

        Private tabControl As TabControl
        Private tabConciliacion As TabPage
        Private tabVentas As TabPage
        Private tabMovimientos As TabPage
        Private tabActa As TabPage

        Private dgvMediosPago As DataGridView
        Private dgvVentas As DataGridView
        Private dgvMovimientos As DataGridView
        Private txtActa As TextBox

        Private lblFormulaInicial As Label
        Private lblFormulaVentasEfec As Label
        Private lblFormulaIngresos As Label
        Private lblFormulaEgresos As Label
        Private lblFormulaEsperado As Label
        Private lblFormulaReal As Label
        Private lblFormulaDiferencia As Label
        Private txtObservaciones As TextBox

        Public Sub New(cajaId As Integer)
            _cajaId = cajaId
            InitializeComponent()
            CargarDatosTurno()
        End Sub

        Private Sub InitializeComponent()
            Me.Text = $"Auditoría de Turno de Caja #{_cajaId}"
            Me.Size = New Size(950, 680)
            Me.MinimumSize = New Size(880, 600)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.BackColor = UITheme.ColorBackground
            Me.Font = UITheme.FontRegular

            ' ─── Panel Superior (Header) ───
            Dim pnlHeader As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 70,
                .BackColor = UITheme.ColorSecondary,
                .Padding = New Padding(20, 12, 20, 12)
            }

            lblTurnoHeader = New Label() With {
                .Text = $"AUDITORÍA DE CAJA #{_cajaId}",
                .Font = UITheme.FontHeading,
                .ForeColor = Color.White,
                .AutoSize = True,
                .Location = New Point(20, 12)
            }

            lblSubtitulo = New Label() With {
                .Text = "Cargando información del turno...",
                .Font = UITheme.FontSmall,
                .ForeColor = UITheme.ColorTextOnDark,
                .AutoSize = True,
                .Location = New Point(22, 40)
            }

            lblEstadoBadge = New Label() With {
                .Text = "CERRADA",
                .Font = UITheme.FontCaption,
                .ForeColor = Color.White,
                .BackColor = UITheme.ColorSuccess,
                .TextAlign = ContentAlignment.MiddleCenter,
                .Size = New Size(95, 26),
                .Location = New Point(820, 20),
                .Anchor = AnchorStyles.Top Or AnchorStyles.Right
            }
            UITheme.ApplyRoundedRegion(lblEstadoBadge, 6)

            pnlHeader.Controls.AddRange({lblTurnoHeader, lblSubtitulo, lblEstadoBadge})

            ' ─── Panel Inferior (Footer con botón de cierre) ───
            Dim pnlFooter As New Panel() With {
                .Dock = DockStyle.Bottom,
                .Height = 55,
                .BackColor = UITheme.ColorSurfaceMuted,
                .Padding = New Padding(15, 10, 20, 10)
            }

            Dim btnCerrar As New Button() With {
                .Text = "Cerrar Auditoría",
                .Size = New Size(140, 36),
                .Anchor = AnchorStyles.Right Or AnchorStyles.Top,
                .Location = New Point(780, 9)
            }
            UITheme.StyleButton(btnCerrar, "Secondary")
            AddHandler btnCerrar.Click, Sub() Me.Close()

            Dim btnImprimirActa As New Button() With {
                .Text = "🖨 Imprimir / Guardar Acta",
                .Size = New Size(190, 36),
                .Anchor = AnchorStyles.Right Or AnchorStyles.Top,
                .Location = New Point(580, 9)
            }
            UITheme.StyleButton(btnImprimirActa, "Primary")
            AddHandler btnImprimirActa.Click, AddressOf BtnImprimirActa_Click

            pnlFooter.Controls.AddRange({btnImprimirActa, btnCerrar})

            ' ─── TabControl Principal ───
            tabControl = New TabControl() With {
                .Dock = DockStyle.Fill,
                .Padding = New Point(14, 8),
                .Font = UITheme.FontBold
            }

            ' Pestaña 1: Conciliación y Medios de Pago
            tabConciliacion = New TabPage("1. Conciliación y Arqueo") With {.BackColor = UITheme.ColorBackground}
            ConstruirTabConciliacion()

            ' Pestaña 2: Comprobantes de Venta
            tabVentas = New TabPage("2. Tickets de Venta") With {.BackColor = UITheme.ColorBackground}
            ConstruirTabVentas()

            ' Pestaña 3: Movimientos Manuales
            tabMovimientos = New TabPage("3. Retiros e Ingresos") With {.BackColor = UITheme.ColorBackground}
            ConstruirTabMovimientos()

            ' Pestaña 4: Acta Oficial de Cierre
            tabActa = New TabPage("4. Acta de Cierre") With {.BackColor = UITheme.ColorBackground}
            ConstruirTabActa()

            tabControl.TabPages.AddRange({tabConciliacion, tabVentas, tabMovimientos, tabActa})

            Me.Controls.Add(tabControl)
            Me.Controls.Add(pnlFooter)
            Me.Controls.Add(pnlHeader)
        End Sub

        Private Sub ConstruirTabConciliacion()
            Dim pnlContent As New Panel() With {
                .Dock = DockStyle.Fill,
                .AutoScroll = True,
                .Padding = New Padding(20)
            }

            ' Tarjeta Izquierda: Fórmula y Conciliación Matemática
            Dim pnlFormulaCard As New Panel() With {
                .Location = New Point(20, 20),
                .Size = New Size(440, 480),
                .BackColor = UITheme.ColorSurface,
                .Padding = New Padding(18)
            }
            AddHandler pnlFormulaCard.Paint, Sub(s, ev)
                                                 Using pen As New Pen(UITheme.ColorBorder, 1)
                                                     ev.Graphics.DrawRectangle(pen, 0, 0, pnlFormulaCard.Width - 1, pnlFormulaCard.Height - 1)
                                                 End Using
                                             End Sub

            Dim lblTitleFormula As New Label() With {
                .Text = "DESGLOSE DEL ARQUEO DE EFECTIVO",
                .Font = UITheme.FontSectionTitle,
                .ForeColor = UITheme.ColorTextPrimary,
                .Location = New Point(18, 16),
                .AutoSize = True
            }

            Dim lblSubFormula As New Label() With {
                .Text = "Conciliación aritmética entre lo registrado y el dinero físico.",
                .Font = UITheme.FontSmall,
                .ForeColor = UITheme.ColorTextSecondary,
                .Location = New Point(18, 38),
                .AutoSize = True
            }

            pnlFormulaCard.Controls.AddRange({lblTitleFormula, lblSubFormula})

            Dim yPos = 70
            lblFormulaInicial = AgregarFilaFormula(pnlFormulaCard, "(+) Fondo Inicial de Cambio:", "$ 0,00", yPos, UITheme.ColorTextPrimary, False)
            yPos += 34
            lblFormulaVentasEfec = AgregarFilaFormula(pnlFormulaCard, "(+) Ventas Cobradas en Efectivo:", "$ 0,00", yPos, UITheme.ColorSuccess, False)
            yPos += 34
            lblFormulaIngresos = AgregarFilaFormula(pnlFormulaCard, "(+) Ingresos Extraordinarios:", "$ 0,00", yPos, UITheme.ColorInfo, False)
            yPos += 34
            lblFormulaEgresos = AgregarFilaFormula(pnlFormulaCard, "(-) Retiros / Gastos Operativos:", "$ 0,00", yPos, UITheme.ColorDanger, False)

            ' Divisor
            yPos += 38
            Dim div1 As New Panel() With {.Location = New Point(18, yPos), .Size = New Size(404, 1), .BackColor = UITheme.ColorBorder}
            pnlFormulaCard.Controls.Add(div1)

            yPos += 10
            lblFormulaEsperado = AgregarFilaFormula(pnlFormulaCard, "(=) TOTAL EFECTIVO ESPERADO:", "$ 0,00", yPos, UITheme.ColorPrimaryDark, True)
            yPos += 38
            lblFormulaReal = AgregarFilaFormula(pnlFormulaCard, "Conteo Real Declarado por Vendedor:", "$ 0,00", yPos, UITheme.ColorTextPrimary, True)

            yPos += 42
            Dim div2 As New Panel() With {.Location = New Point(18, yPos), .Size = New Size(404, 2), .BackColor = UITheme.ColorBorderStrong}
            pnlFormulaCard.Controls.Add(div2)

            yPos += 14
            lblFormulaDiferencia = AgregarFilaFormula(pnlFormulaCard, "DIFERENCIA DE ARQUEO:", "$ 0,00", yPos, UITheme.ColorSuccess, True)
            lblFormulaDiferencia.Font = UITheme.FontHeading

            ' Observaciones del cajero
            yPos += 48
            Dim lblObsTitle As New Label() With {
                .Text = "Observaciones ingresadas por el cajero al cerrar:",
                .Font = UITheme.FontSmallBold,
                .ForeColor = UITheme.ColorTextSecondary,
                .Location = New Point(18, yPos),
                .AutoSize = True
            }
            pnlFormulaCard.Controls.Add(lblObsTitle)

            yPos += 22
            txtObservaciones = New TextBox() With {
                .Location = New Point(18, yPos),
                .Size = New Size(404, 60),
                .Multiline = True,
                .ReadOnly = True,
                .BackColor = UITheme.ColorSurfaceMuted,
                .ScrollBars = ScrollBars.Vertical,
                .Font = UITheme.FontRegular
            }
            pnlFormulaCard.Controls.Add(txtObservaciones)

            ' Tarjeta Derecha: Conciliación de Medios de Pago
            Dim pnlMediosCard As New Panel() With {
                .Location = New Point(475, 20),
                .Size = New Size(420, 480),
                .BackColor = UITheme.ColorSurface,
                .Padding = New Padding(18)
            }
            AddHandler pnlMediosCard.Paint, Sub(s, ev)
                                                Using pen As New Pen(UITheme.ColorBorder, 1)
                                                    ev.Graphics.DrawRectangle(pen, 0, 0, pnlMediosCard.Width - 1, pnlMediosCard.Height - 1)
                                                End Using
                                            End Sub

            Dim lblTitleMedios As New Label() With {
                .Text = "VENTAS POR MEDIO DE PAGO",
                .Font = UITheme.FontSectionTitle,
                .ForeColor = UITheme.ColorTextPrimary,
                .Location = New Point(18, 16),
                .AutoSize = True
            }

            Dim lblSubMedios As New Label() With {
                .Text = "Control de terminales de tarjeta (Posnet), transferencias y QR.",
                .Font = UITheme.FontSmall,
                .ForeColor = UITheme.ColorTextSecondary,
                .Location = New Point(18, 38),
                .AutoSize = True
            }

            dgvMediosPago = New DataGridView() With {
                .Location = New Point(18, 70),
                .Size = New Size(384, 385),
                .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
            }
            UITheme.StyleDataGrid(dgvMediosPago)
            ConfigurarColumnasMediosPago()

            pnlMediosCard.Controls.AddRange({lblTitleMedios, lblSubMedios, dgvMediosPago})

            pnlContent.Controls.AddRange({pnlFormulaCard, pnlMediosCard})
            tabConciliacion.Controls.Add(pnlContent)
        End Sub

        Private Function AgregarFilaFormula(parent As Panel, labelText As String, valText As String, y As Integer, valColor As Color, isBold As Boolean) As Label
            Dim lblDesc As New Label() With {
                .Text = labelText,
                .Font = If(isBold, UITheme.FontBold, UITheme.FontRegular),
                .ForeColor = UITheme.ColorTextPrimary,
                .Location = New Point(18, y),
                .Size = New Size(240, 24),
                .TextAlign = ContentAlignment.MiddleLeft
            }

            Dim lblVal As New Label() With {
                .Text = valText,
                .Font = If(isBold, UITheme.FontBold, UITheme.FontRegular),
                .ForeColor = valColor,
                .Location = New Point(265, y),
                .Size = New Size(157, 24),
                .TextAlign = ContentAlignment.MiddleRight
            }

            parent.Controls.AddRange({lblDesc, lblVal})
            Return lblVal
        End Function

        Private Sub ConfigurarColumnasMediosPago()
            dgvMediosPago.Columns.Clear()
            dgvMediosPago.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            ' ── Medio de Pago (izquierda, relleno) ──
            dgvMediosPago.Columns.Add("MetodoPago", "Medio de Pago")
            dgvMediosPago.Columns("MetodoPago").FillWeight = 60
            dgvMediosPago.Columns("MetodoPago").MinimumWidth = 90
            dgvMediosPago.Columns("MetodoPago").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvMediosPago.Columns("MetodoPago").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft

            ' ── Tickets (centrado) ──
            dgvMediosPago.Columns.Add("CantidadTickets", "Tickets")
            dgvMediosPago.Columns("CantidadTickets").FillWeight = 20
            dgvMediosPago.Columns("CantidadTickets").MinimumWidth = 60
            dgvMediosPago.Columns("CantidadTickets").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvMediosPago.Columns("CantidadTickets").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            ' ── Total Cobrado (derecha, negrita) ──
            dgvMediosPago.Columns.Add("TotalCobrado", "Total Cobrado")
            dgvMediosPago.Columns("TotalCobrado").FillWeight = 30
            dgvMediosPago.Columns("TotalCobrado").MinimumWidth = 100
            dgvMediosPago.Columns("TotalCobrado").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvMediosPago.Columns("TotalCobrado").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvMediosPago.Columns("TotalCobrado").DefaultCellStyle.Format = "C2"
            dgvMediosPago.Columns("TotalCobrado").DefaultCellStyle.Font = UITheme.FontBold
        End Sub

        Private Sub ConstruirTabVentas()
            Dim pnlGrid As New Panel() With {.Dock = DockStyle.Fill, .Padding = New Padding(20)}
            dgvVentas = New DataGridView() With {.Dock = DockStyle.Fill}
            UITheme.StyleDataGrid(dgvVentas)
            dgvVentas.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            ' ── Nº Ticket (centrado) ──
            dgvVentas.Columns.Add("Ticket", "Nº Ticket")
            dgvVentas.Columns("Ticket").FillWeight = 12
            dgvVentas.Columns("Ticket").MinimumWidth = 100
            dgvVentas.Columns("Ticket").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvVentas.Columns("Ticket").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            ' ── Hora (centrado) ──
            dgvVentas.Columns.Add("Fecha", "Hora")
            dgvVentas.Columns("Fecha").FillWeight = 8
            dgvVentas.Columns("Fecha").MinimumWidth = 70
            dgvVentas.Columns("Fecha").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvVentas.Columns("Fecha").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            ' ── Cliente (izquierda, relleno) ──
            dgvVentas.Columns.Add("Cliente", "Cliente")
            dgvVentas.Columns("Cliente").FillWeight = 20
            dgvVentas.Columns("Cliente").MinimumWidth = 100
            dgvVentas.Columns("Cliente").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvVentas.Columns("Cliente").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft

            ' ── Medio de Pago (centrado) ──
            dgvVentas.Columns.Add("Metodo", "Medio de Pago")
            dgvVentas.Columns("Metodo").FillWeight = 11
            dgvVentas.Columns("Metodo").MinimumWidth = 90
            dgvVentas.Columns("Metodo").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvVentas.Columns("Metodo").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            ' ── Subtotal (derecha) ──
            dgvVentas.Columns.Add("Subtotal", "Subtotal")
            dgvVentas.Columns("Subtotal").FillWeight = 9
            dgvVentas.Columns("Subtotal").MinimumWidth = 80
            dgvVentas.Columns("Subtotal").DefaultCellStyle.Format = "C2"
            dgvVentas.Columns("Subtotal").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvVentas.Columns("Subtotal").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight

            ' ── Descuento (derecha) ──
            dgvVentas.Columns.Add("Descuento", "Descuento")
            dgvVentas.Columns("Descuento").FillWeight = 8
            dgvVentas.Columns("Descuento").MinimumWidth = 75
            dgvVentas.Columns("Descuento").DefaultCellStyle.Format = "C2"
            dgvVentas.Columns("Descuento").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvVentas.Columns("Descuento").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight

            ' ── Total Facturado (derecha, negrita) ──
            dgvVentas.Columns.Add("Total", "Total Facturado")
            dgvVentas.Columns("Total").FillWeight = 11
            dgvVentas.Columns("Total").MinimumWidth = 95
            dgvVentas.Columns("Total").DefaultCellStyle.Format = "C2"
            dgvVentas.Columns("Total").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvVentas.Columns("Total").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvVentas.Columns("Total").DefaultCellStyle.Font = UITheme.FontBold

            ' ── Estado (centrado) ──
            dgvVentas.Columns.Add("Estado", "Estado")
            dgvVentas.Columns("Estado").FillWeight = 9
            dgvVentas.Columns("Estado").MinimumWidth = 80
            dgvVentas.Columns("Estado").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvVentas.Columns("Estado").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            pnlGrid.Controls.Add(dgvVentas)
            tabVentas.Controls.Add(pnlGrid)
        End Sub

        Private Sub ConstruirTabMovimientos()
            Dim pnlGrid As New Panel() With {.Dock = DockStyle.Fill, .Padding = New Padding(20)}
            dgvMovimientos = New DataGridView() With {.Dock = DockStyle.Fill}
            UITheme.StyleDataGrid(dgvMovimientos)

            dgvMovimientos.Columns.Add("Hora", "Fecha y Hora")
            dgvMovimientos.Columns("Hora").Width = 130

            dgvMovimientos.Columns.Add("Tipo", "Tipo")
            dgvMovimientos.Columns("Tipo").Width = 110

            dgvMovimientos.Columns.Add("Concepto", "Concepto / Justificación")
            dgvMovimientos.Columns("Concepto").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

            dgvMovimientos.Columns.Add("Monto", "Monto ($)")
            dgvMovimientos.Columns("Monto").Width = 130
            dgvMovimientos.Columns("Monto").DefaultCellStyle.Format = "C2"
            dgvMovimientos.Columns("Monto").DefaultCellStyle.Font = UITheme.FontBold

            dgvMovimientos.Columns.Add("Usuario", "Registrado Por")
            dgvMovimientos.Columns("Usuario").Width = 160

            dgvMovimientos.Columns.Add("Referencia", "Comprobante Ref.")
            dgvMovimientos.Columns("Referencia").Width = 140

            AddHandler dgvMovimientos.CellFormatting, Sub(s, e)
                                                          If e.RowIndex >= 0 AndAlso dgvMovimientos.Columns(e.ColumnIndex).Name = "Tipo" Then
                                                              Dim val = dgvMovimientos.Rows(e.RowIndex).Cells("Tipo").Value?.ToString()
                                                              If val = "Egreso" Then
                                                                  e.CellStyle.ForeColor = UITheme.ColorDanger
                                                                  e.CellStyle.Font = UITheme.FontBold
                                                              ElseIf val = "Ingreso" Then
                                                                  e.CellStyle.ForeColor = UITheme.ColorSuccess
                                                                  e.CellStyle.Font = UITheme.FontBold
                                                              End If
                                                          End If
                                                      End Sub

            pnlGrid.Controls.Add(dgvMovimientos)
            tabMovimientos.Controls.Add(pnlGrid)
        End Sub

        Private Sub ConstruirTabActa()
            Dim pnlActa As New Panel() With {.Dock = DockStyle.Fill, .Padding = New Padding(20)}

            txtActa = New TextBox() With {
                .Multiline = True,
                .ReadOnly = True,
                .Dock = DockStyle.Fill,
                .Font = New Font("Courier New", 9.5F, FontStyle.Regular),
                .BackColor = Color.White,
                .ForeColor = Color.Black,
                .ScrollBars = ScrollBars.Vertical
            }

            pnlActa.Controls.Add(txtActa)
            tabActa.Controls.Add(pnlActa)
        End Sub

        Private Sub CargarDatosTurno()
            _caja = _cajaService.GetCajaById(_cajaId)
            If _caja Is Nothing Then
                MessageBox.Show("No se encontró la caja solicitada.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.Close()
                Return
            End If

            ' Header Info
            lblTurnoHeader.Text = $"AUDITORÍA DE TURNO DE CAJA #{_caja.Id} - VENDEDOR: {_caja.UsuarioNombre.ToUpper()}"
            Dim cierreInfo = If(_caja.FechaCierre.HasValue, _caja.FechaCierre.Value.ToString("dd/MM/yyyy HH:mm"), "En curso")
            lblSubtitulo.Text = $"Apertura: {_caja.FechaApertura:dd/MM/yyyy HH:mm}  |  Cierre: {cierreInfo}  |  Duración: {_caja.DuracionTexto}"

            If _caja.Estado = "Abierta" Then
                lblEstadoBadge.Text = "EN CURSO"
                lblEstadoBadge.BackColor = UITheme.ColorWarning
            Else
                lblEstadoBadge.Text = "CERRADA"
                lblEstadoBadge.BackColor = UITheme.ColorSuccess
            End If

            ' Desglose Matemático
            lblFormulaInicial.Text = _caja.MontoInicial.ToString("C2")
            lblFormulaVentasEfec.Text = _caja.TotalVentasEfectivo.ToString("C2")
            lblFormulaIngresos.Text = _caja.TotalIngresos.ToString("C2")
            lblFormulaEgresos.Text = _caja.TotalEgresos.ToString("C2")
            lblFormulaEsperado.Text = _caja.MontoEsperado.ToString("C2")

            If _caja.MontoReal.HasValue Then
                lblFormulaReal.Text = _caja.MontoReal.Value.ToString("C2")
            Else
                lblFormulaReal.Text = "(Pendiente de cierre)"
            End If

            If _caja.Diferencia.HasValue Then
                Dim dif = _caja.Diferencia.Value
                If dif = 0 Then
                    lblFormulaDiferencia.Text = "$ 0,00 (Exacto)"
                    lblFormulaDiferencia.ForeColor = UITheme.ColorSuccess
                ElseIf dif > 0 Then
                    lblFormulaDiferencia.Text = $"+{dif:C2} (Sobrante)"
                    lblFormulaDiferencia.ForeColor = UITheme.ColorInfo
                Else
                    lblFormulaDiferencia.Text = $"{dif:C2} (Faltante)"
                    lblFormulaDiferencia.ForeColor = UITheme.ColorDanger
                End If
            Else
                lblFormulaDiferencia.Text = "--"
                lblFormulaDiferencia.ForeColor = UITheme.ColorTextSecondary
            End If

            txtObservaciones.Text = If(String.IsNullOrWhiteSpace(_caja.Observaciones), "(Sin observaciones registradas por el cajero)", _caja.Observaciones)

            ' Cargar Medios de Pago
            Dim dtMedios = _cajaService.GetResumenMediosPagoCaja(_cajaId)
            dgvMediosPago.Rows.Clear()
            For Each row As DataRow In dtMedios.Rows
                dgvMediosPago.Rows.Add(
                    row("MetodoPago").ToString(),
                    row("CantidadTickets"),
                    Convert.ToDecimal(row("TotalCobrado"))
                )
            Next

            ' Cargar Tickets de Venta
            Dim ventas = _cajaService.GetVentasPorCaja(_cajaId)
            dgvVentas.Rows.Clear()
            For Each v In ventas
                dgvVentas.Rows.Add(
                    v.NumeroTicket,
                    v.Fecha.ToString("HH:mm:ss"),
                    v.ClienteNombre,
                    v.MetodoPago,
                    v.Subtotal,
                    v.DescuentoMonto,
                    v.Total,
                    v.Estado
                )
            Next
            tabVentas.Text = $"2. Tickets de Venta ({ventas.Count})"

            ' Cargar Movimientos de Efectivo
            Dim movs = _cajaService.GetMovimientosCaja(_cajaId)
            dgvMovimientos.Rows.Clear()
            For Each m In movs
                dgvMovimientos.Rows.Add(
                    m.Fecha.ToString("dd/MM/yyyy HH:mm"),
                    m.Tipo,
                    m.Concepto,
                    m.Monto,
                    m.UsuarioNombre,
                    m.Referencia
                )
            Next
            tabMovimientos.Text = $"3. Retiros e Ingresos ({movs.Count})"

            ' Generar Texto del Acta
            GenerarActaOficial()
        End Sub

        Private Sub GenerarActaOficial()
            If _caja Is Nothing Then Return
            Dim cfg = _configService.GetConfiguracion()
            Dim sb As New StringBuilder()

            sb.AppendLine("======================================================================")
            sb.AppendLine($"           {cfg.NombreComercio.ToUpper()} - CONTROL Y AUDITORÍA          ")
            sb.AppendLine($"CUIT: {cfg.Cuit}  |  {cfg.CondicionIva}  |  Tel: {cfg.Telefono}")
            sb.AppendLine($"Dirección: {cfg.Direccion}")
            sb.AppendLine("======================================================================")
            sb.AppendLine("                 ACTA DE ARQUEO Y CIERRE DE TURNO                    ")
            sb.AppendLine("======================================================================")
            sb.AppendLine($"Nº DE TURNO / CAJA : #{_caja.Id}")
            sb.AppendLine($"FECHA Y HORA APERTURA: {_caja.FechaApertura:dd/MM/yyyy HH:mm:ss}")
            Dim cierreStr = If(_caja.FechaCierre.HasValue, _caja.FechaCierre.Value.ToString("dd/MM/yyyy HH:mm:ss"), "EN CURSO (Sin cierre final)")
            sb.AppendLine($"FECHA Y HORA CIERRE  : {cierreStr}")
            sb.AppendLine($"DURACIÓN DEL TURNO   : {_caja.DuracionTexto}")
            sb.AppendLine($"CAJERO / RESPONSABLE : {_caja.UsuarioNombre}")
            sb.AppendLine($"ESTADO ACTUAL        : {_caja.Estado.ToUpper()}")
            sb.AppendLine("----------------------------------------------------------------------")
            sb.AppendLine("                  CONCILIACIÓN DEL DINERO FÍSICO                     ")
            sb.AppendLine("----------------------------------------------------------------------")
            sb.AppendLine($" (+) Fondo Inicial de Cambio     : {_caja.MontoInicial.ToString("C2").PadLeft(20)}")
            sb.AppendLine($" (+) Ventas Cobradas en Efectivo : {_caja.TotalVentasEfectivo.ToString("C2").PadLeft(20)}")
            sb.AppendLine($" (+) Ingresos Extraordinarios    : {_caja.TotalIngresos.ToString("C2").PadLeft(20)}")
            sb.AppendLine($" (-) Retiros / Gastos de Turno   : {_caja.TotalEgresos.ToString("C2").PadLeft(20)}")
            sb.AppendLine(" --------------------------------------------------------------------")
            sb.AppendLine($" (=) EFECTIVO TOTAL ESPERADO     : {_caja.MontoEsperado.ToString("C2").PadLeft(20)}")
            Dim realStr = If(_caja.MontoReal.HasValue, _caja.MontoReal.Value.ToString("C2"), "NO DECLARADO")
            sb.AppendLine($" ( ) EFECTIVO REAL DECLARADO     : {realStr.PadLeft(20)}")
            sb.AppendLine(" ====================================================================")

            If _caja.Diferencia.HasValue Then
                Dim dif = _caja.Diferencia.Value
                Dim etiquetaDif = If(dif = 0, "ARQUEO EXACTO", If(dif > 0, "SOBRANTE DE CAJA", "FALTANTE DE CAJA"))
                sb.AppendLine($" RESULTADO FINAL: {etiquetaDif} : {dif.ToString("C2").PadLeft(20)}")
            Else
                sb.AppendLine(" RESULTADO FINAL: TURNO EN CURSO (Sin balance definitivo)")
            End If

            sb.AppendLine("======================================================================")
            sb.AppendLine("                     VENTAS POR MEDIO DE PAGO                         ")
            sb.AppendLine("----------------------------------------------------------------------")
            sb.AppendLine($" Ventas en Efectivo              : {_caja.TotalVentasEfectivo.ToString("C2").PadLeft(20)}")
            sb.AppendLine($" Ventas Digitales (Tarjetas/QR)  : {_caja.TotalVentasDigital.ToString("C2").PadLeft(20)}")
            sb.AppendLine($" TOTAL FACTURADO EN EL TURNO     : {_caja.TotalVentasTotal.ToString("C2").PadLeft(20)}")
            sb.AppendLine("----------------------------------------------------------------------")
            sb.AppendLine("OBSERVACIONES DEL CAJERO:")
            sb.AppendLine(If(String.IsNullOrWhiteSpace(_caja.Observaciones), "Ninguna.", _caja.Observaciones))
            sb.AppendLine("======================================================================")
            sb.AppendLine("CONSTANCIA DE CONFORMIDAD Y AUDITORÍA:")
            sb.AppendLine()
            sb.AppendLine()
            sb.AppendLine("   _____________________________         _____________________________")
            sb.AppendLine("         Firma del Cajero                      Firma del Gerente      ")
            sb.AppendLine($"     {_caja.UsuarioNombre}                {If(AuthService.CurrentUser IsNot Nothing, AuthService.CurrentUser.NombreCompleto, "Auditor")}")
            sb.AppendLine("======================================================================")
            sb.AppendLine($"Impreso el {DateTime.Now:dd/MM/yyyy HH:mm:ss} por {If(AuthService.CurrentUser IsNot Nothing, AuthService.CurrentUser.NombreCompleto, "Sistema")}")

            txtActa.Text = sb.ToString()
        End Sub

        Private Sub BtnImprimirActa_Click(sender As Object, e As EventArgs)
            If String.IsNullOrWhiteSpace(txtActa.Text) Then Return

            Using sfd As New SaveFileDialog()
                sfd.Filter = "Documento de Texto (*.txt)|*.txt|Todos los archivos (*.*)|*.*"
                sfd.FileName = $"Acta_Arqueo_Caja_{_caja.Id}_{_caja.FechaApertura:yyyyMMdd}.txt"
                sfd.Title = "Guardar Acta de Arqueo y Cierre"

                If sfd.ShowDialog(Me) = DialogResult.OK Then
                    Try
                        System.IO.File.WriteAllText(sfd.FileName, txtActa.Text, Encoding.UTF8)
                        MessageBox.Show("Acta guardada correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Catch ex As Exception
                        MessageBox.Show("Error al guardar archivo: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End If
            End Using
        End Sub

    End Class
End Namespace
