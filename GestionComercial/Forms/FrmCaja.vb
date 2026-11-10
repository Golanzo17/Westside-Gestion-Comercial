' ARCHIVO: FrmCaja.vb
' PROPÓSITO: Control de turnos de caja, registro de gastos y arqueo ciego de efectivo.
' Esta ventana garantiza el control sobre el dinero físico en el salón:
' 1. Visualización de Balances: Muestra en tarjetas el estado de la caja y los totales recaudados
'    en el turno separados por Efectivo, Tarjetas (Posnet) y Transferencias / QR.
' 2. Seguridad en Retiros: Si el cajero registra un retiro o egreso de efectivo, el sistema
'    exige la autorización presencial de un Administrador o Gerente.
' 3. Arqueo y Cierre: Al finalizar el turno, pide contar los billetes en mano (montoReal),
'    calcula automáticamente si sobró o faltó plata y archiva el cierre para auditoría gerencial.


Imports System.Drawing
Imports System.Windows.Forms
Imports GestionComercial.Models
Imports GestionComercial.Services
Imports GestionComercial.UI

Namespace Forms
    Public Class FrmCaja
        Inherits Form

        Private cajaService As New CajaService()
        Private cajaActual As Caja = Nothing

        Private lblEstadoCaja As Label
        Private lblSubEstadoCaja As Label
        Private lblVentasEfectivo As Label
        Private lblVentasTarjetas As Label
        Private lblVentasTransferencia As Label

        Private btnAbrirCaja As Button
        Private btnNuevoMovimiento As Button
        Private btnCerrarCaja As Button
        Private btnRefrescar As Button
        Private dgvMovimientos As DataGridView

        Public Sub New()
            InitializeUI()
            LoadCajaData()
        End Sub

        Private Sub InitializeUI()
            Me.Text = "Caja Diaria y Arqueo de Turno"
            Me.Size = New Size(1000, 640)
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
                .Text = "Caja diaria, movimientos manuales y arqueo de turno",
                .Font = UITheme.FontHeading,
                .ForeColor = Color.White,
                .AutoSize = True,
                .Location = New Point(15, 15)
            }
            pnlHeader.Controls.Add(lblTitle)

            ' Tarjetas KPI de Estado de Caja (responsivas en 4 columnas iguales)
            Dim pnlKpis As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 125,
                .BackColor = UITheme.ColorSurface,
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

            ' Tarjeta 1: Estado de Caja y Fondo Inicial
            Dim cardEstado = UITheme.CreateKpiCard("ESTADO DE CAJA", "CERRADA", "Sin turno activo", UITheme.ColorDanger)
            cardEstado.Dock = DockStyle.Fill
            cardEstado.Margin = New Padding(3, 3, 6, 3)
            lblEstadoCaja = CType(cardEstado.Controls(2), Label)
            lblSubEstadoCaja = CType(cardEstado.Controls(3), Label)
            tlpKpis.Controls.Add(cardEstado, 0, 0)

            ' Tarjeta 2: Ventas Efectivo
            Dim cardEfec = UITheme.CreateKpiCard("VENTAS EFECTIVO", "$ 0,00", "Cobrado en billetes", UITheme.ColorSuccess)
            cardEfec.Dock = DockStyle.Fill
            cardEfec.Margin = New Padding(6, 3, 6, 3)
            lblVentasEfectivo = CType(cardEfec.Controls(2), Label)
            tlpKpis.Controls.Add(cardEfec, 1, 0)

            ' Tarjeta 3: Ventas con Tarjetas (Posnet)
            Dim cardTarj = UITheme.CreateKpiCard("TARJETAS (POSNET)", "$ 0,00", "Débito y Crédito", UITheme.ColorInfo)
            cardTarj.Dock = DockStyle.Fill
            cardTarj.Margin = New Padding(6, 3, 6, 3)
            lblVentasTarjetas = CType(cardTarj.Controls(2), Label)
            tlpKpis.Controls.Add(cardTarj, 2, 0)

            ' Tarjeta 4: Ventas por Transferencia / QR
            Dim cardTransf = UITheme.CreateKpiCard("TRANSFERENCIA / QR", "$ 0,00", "MercadoPago / Bancos", UITheme.ColorPrimary)
            cardTransf.Dock = DockStyle.Fill
            cardTransf.Margin = New Padding(6, 3, 3, 3)
            lblVentasTransferencia = CType(cardTransf.Controls(2), Label)
            tlpKpis.Controls.Add(cardTransf, 3, 0)

            pnlKpis.Controls.Add(tlpKpis)

            ' Barra de Acciones responsiva
            Dim pnlToolbar As New Panel() With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .MinimumSize = New Size(0, 54),
                .BackColor = UITheme.ColorSurfaceMuted
            }
            Dim flpToolbar As New FlowLayoutPanel() With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .WrapContents = True,
                .Padding = New Padding(15, 10, 15, 6),
                .BackColor = Color.Transparent
            }

            btnAbrirCaja = New Button() With {.Text = "Abrir caja de turno", .Size = New Size(180, 34), .Margin = New Padding(0, 0, 10, 4)}
            UITheme.StyleButton(btnAbrirCaja, "Primary")
            AddHandler btnAbrirCaja.Click, AddressOf BtnAbrirCaja_Click

            btnNuevoMovimiento = New Button() With {.Text = "Ingreso / retiro de efectivo", .Size = New Size(230, 34), .Margin = New Padding(0, 0, 10, 4)}
            UITheme.StyleButton(btnNuevoMovimiento, "Secondary")
            AddHandler btnNuevoMovimiento.Click, AddressOf BtnNuevoMovimiento_Click

            btnCerrarCaja = New Button() With {.Text = "Arqueo y cierre de caja", .Size = New Size(210, 34), .Margin = New Padding(0, 0, 10, 4)}
            UITheme.StyleButton(btnCerrarCaja, "Danger")
            AddHandler btnCerrarCaja.Click, AddressOf BtnCerrarCaja_Click

            btnRefrescar = New Button() With {.Text = "Actualizar", .Size = New Size(120, 34), .Margin = New Padding(0, 0, 0, 4)}
            UITheme.StyleButton(btnRefrescar, "Secondary")
            AddHandler btnRefrescar.Click, Sub() LoadCajaData()

            flpToolbar.Controls.AddRange({btnAbrirCaja, btnNuevoMovimiento, btnCerrarCaja, btnRefrescar})
            pnlToolbar.Controls.Add(flpToolbar)

            ' Grilla de Movimientos Manuales del Turno Actual
            Dim pnlGrid As New Panel() With {.Dock = DockStyle.Fill, .Padding = New Padding(15)}
            dgvMovimientos = New DataGridView() With {.Dock = DockStyle.Fill}
            UITheme.StyleDataGrid(dgvMovimientos)
            ConfigurarColumnas()
            pnlGrid.Controls.Add(dgvMovimientos)

            ' Orden de Docking: Fill primero, Top de abajo hacia arriba
            Me.Controls.Add(pnlGrid)
            Me.Controls.Add(pnlToolbar)
            Me.Controls.Add(pnlKpis)
            Me.Controls.Add(pnlHeader)
        End Sub

        Private Sub ConfigurarColumnas()
            dgvMovimientos.Columns.Clear()
            dgvMovimientos.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            dgvMovimientos.Columns.Add("Hora", "Hora")
            dgvMovimientos.Columns("Hora").FillWeight = 10
            dgvMovimientos.Columns("Hora").MinimumWidth = 75
            dgvMovimientos.Columns("Hora").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvMovimientos.Columns("Hora").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            dgvMovimientos.Columns.Add("Tipo", "Tipo")
            dgvMovimientos.Columns("Tipo").FillWeight = 11
            dgvMovimientos.Columns("Tipo").MinimumWidth = 80
            dgvMovimientos.Columns("Tipo").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvMovimientos.Columns("Tipo").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            dgvMovimientos.Columns.Add("Concepto", "Concepto / Detalle")
            dgvMovimientos.Columns("Concepto").FillWeight = 28
            dgvMovimientos.Columns("Concepto").MinimumWidth = 130

            dgvMovimientos.Columns.Add("Monto", "Monto ($)")
            dgvMovimientos.Columns("Monto").FillWeight = 13
            dgvMovimientos.Columns("Monto").MinimumWidth = 95
            dgvMovimientos.Columns("Monto").DefaultCellStyle.Format = "C2"
            dgvMovimientos.Columns("Monto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvMovimientos.Columns("Monto").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight

            dgvMovimientos.Columns.Add("Usuario", "Registrado Por")
            dgvMovimientos.Columns("Usuario").FillWeight = 14
            dgvMovimientos.Columns("Usuario").MinimumWidth = 100

            dgvMovimientos.Columns.Add("AutorizadoPor", "Autorizado Por")
            dgvMovimientos.Columns("AutorizadoPor").FillWeight = 14
            dgvMovimientos.Columns("AutorizadoPor").MinimumWidth = 100

            dgvMovimientos.Columns.Add("Referencia", "Comprobante Ref.")
            dgvMovimientos.Columns("Referencia").FillWeight = 12
            dgvMovimientos.Columns("Referencia").MinimumWidth = 90
        End Sub

        Private Sub LoadCajaData()
            Dim userId As Integer = If(AuthService.CurrentUser IsNot Nothing, AuthService.CurrentUser.Id, 1)
            cajaActual = cajaService.GetCajaAbierta(userId)

            If cajaActual IsNot Nothing Then
                lblEstadoCaja.Text = "ABIERTA"
                lblEstadoCaja.ForeColor = UITheme.ColorSuccess
                btnAbrirCaja.Enabled = False
                btnNuevoMovimiento.Enabled = True
                btnCerrarCaja.Enabled = True

                Dim desglose = cajaService.GetDesgloseMediosTurno(cajaActual.Id)
                lblVentasEfectivo.Text = desglose.TotalEfectivo.ToString("C2")
                lblVentasTarjetas.Text = desglose.TotalTarjetas.ToString("C2")
                lblVentasTransferencia.Text = desglose.TotalTransferencia.ToString("C2")

                Dim esperado As Decimal = cajaActual.MontoInicial + cajaActual.TotalVentasEfectivo + cajaActual.TotalIngresos - cajaActual.TotalEgresos
                If lblSubEstadoCaja IsNot Nothing Then
                    lblSubEstadoCaja.Text = If(AuthService.IsAdminOrManager, $"En cajón: {esperado:C2}", $"Fondo inic.: {cajaActual.MontoInicial:C2} · Arqueo ciego")
                End If

                ' Cargar movimientos
                Dim movs = cajaService.GetMovimientosCaja(cajaActual.Id)
                dgvMovimientos.Rows.Clear()
                For Each m In movs
                    dgvMovimientos.Rows.Add(m.Fecha.ToString("HH:mm:ss"), m.Tipo, m.Concepto, m.Monto, m.UsuarioNombre, m.AutorizadoPorNombre, m.Referencia)
                Next
            Else
                lblEstadoCaja.Text = "CERRADA"
                lblEstadoCaja.ForeColor = UITheme.ColorDanger
                If lblSubEstadoCaja IsNot Nothing Then
                    lblSubEstadoCaja.Text = "Sin turno activo"
                End If
                btnAbrirCaja.Enabled = True
                btnNuevoMovimiento.Enabled = False
                btnCerrarCaja.Enabled = False

                lblVentasEfectivo.Text = "$ 0,00"
                lblVentasTarjetas.Text = "$ 0,00"
                lblVentasTransferencia.Text = "$ 0,00"
                dgvMovimientos.Rows.Clear()
            End If
        End Sub

        Private Sub BtnAbrirCaja_Click(sender As Object, e As EventArgs)
            Dim prompt As New Form() With {
                .Text = "Apertura de Caja",
                .Size = New Size(400, 220),
                .StartPosition = FormStartPosition.CenterParent,
                .FormBorderStyle = FormBorderStyle.FixedDialog,
                .MaximizeBox = False,
                .MinimizeBox = False,
                .BackColor = Color.White
            }

            Dim lblM As New Label() With {.Text = "Monto inicial en efectivo (cambio):", .Font = UITheme.FontBold, .Location = New Point(30, 25), .AutoSize = True}
            Dim numIni As New NumericUpDown() With {.Location = New Point(30, 55), .Size = New Size(320, 28), .Maximum = 10000000, .DecimalPlaces = 2, .Value = 10000}
            Dim btnOk As New Button() With {.Text = "Confirmar Apertura", .Location = New Point(30, 110), .Size = New Size(320, 40)}
            UITheme.StyleButton(btnOk, "Success")

            AddHandler btnOk.Click, Sub()
                                        prompt.DialogResult = DialogResult.OK
                                        prompt.Close()
                                    End Sub

            prompt.Controls.AddRange({lblM, numIni, btnOk})
            If prompt.ShowDialog() = DialogResult.OK Then
                Dim userId As Integer = If(AuthService.CurrentUser IsNot Nothing, AuthService.CurrentUser.Id, 1)
                Dim errMsg As String = ""
                Dim nuevaCaja = cajaService.AbrirCaja(userId, numIni.Value, errMsg)
                If nuevaCaja IsNot Nothing Then
                    MessageBox.Show("Caja abierta exitosamente con $" & numIni.Value.ToString("N2"), "Caja Abierta", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    LoadCajaData()
                Else
                    MessageBox.Show("Error al abrir caja: " & errMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        End Sub

        Private Sub BtnNuevoMovimiento_Click(sender As Object, e As EventArgs)
            If cajaActual Is Nothing Then Return

            Dim frmMov As New Form() With {
                .Text = "Ingreso / Retiro de Efectivo de Caja",
                .Size = New Size(450, 300),
                .StartPosition = FormStartPosition.CenterParent,
                .FormBorderStyle = FormBorderStyle.FixedDialog,
                .MaximizeBox = False,
                .MinimizeBox = False,
                .BackColor = Color.White
            }

            Dim lblT As New Label() With {.Text = "Tipo de Movimiento:", .Font = UITheme.FontBold, .Location = New Point(30, 20), .AutoSize = True}
            Dim cboTipo As New ComboBox() With {.Location = New Point(30, 45), .Size = New Size(370, 26), .DropDownStyle = ComboBoxStyle.DropDownList}
            UITheme.StyleComboBox(cboTipo)
            cboTipo.Items.AddRange({"Egreso", "Ingreso"})
            cboTipo.SelectedIndex = 0

            Dim lblC As New Label() With {.Text = "Concepto (Ej: Pago flete, Retiro dueño, Proveedor):", .Font = UITheme.FontBold, .Location = New Point(30, 85), .AutoSize = True}
            Dim txtConcepto As New TextBox() With {.Location = New Point(30, 110), .Size = New Size(370, 26)}
            UITheme.StyleTextBox(txtConcepto)

            Dim lblM As New Label() With {.Text = "Monto ($):", .Font = UITheme.FontBold, .Location = New Point(30, 145), .AutoSize = True}
            Dim numMonto As New NumericUpDown() With {.Location = New Point(30, 170), .Size = New Size(180, 26), .Maximum = 10000000, .DecimalPlaces = 2}

            Dim btnOk As New Button() With {.Text = "Registrar", .Location = Point.Add(New Point(220, 168), New Size(0, 0)), .Size = New Size(180, 30)}
            UITheme.StyleButton(btnOk, "Primary")

            AddHandler btnOk.Click, Sub()
                                        If String.IsNullOrWhiteSpace(txtConcepto.Text) OrElse numMonto.Value <= 0 Then
                                            MessageBox.Show("Ingrese un concepto y monto válido.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                            Return
                                        End If
                                        frmMov.DialogResult = DialogResult.OK
                                        frmMov.Close()
                                    End Sub

            frmMov.Controls.AddRange({lblT, cboTipo, lblC, txtConcepto, lblM, numMonto, btnOk})
            If frmMov.ShowDialog() = DialogResult.OK Then
                Dim tipoMov As String = cboTipo.SelectedItem.ToString()
                ' Los egresos / retiros de dinero requieren autorización si el usuario es Vendedor
                If tipoMov.Equals("Egreso", StringComparison.OrdinalIgnoreCase) Then
                    If Not AuthService.SolicitarAutorizacionAdminOManager(Me, "Los retiros o egresos de efectivo de caja requieren autorización de un Administrador o Gerente.") Then
                        Return
                    End If
                End If

                Dim userId As Integer = If(AuthService.CurrentUser IsNot Nothing, AuthService.CurrentUser.Id, 1)
                Dim errMsg As String = ""
                Dim autorizadoPorId As Nullable(Of Integer) = If(tipoMov.Equals("Egreso", StringComparison.OrdinalIgnoreCase) AndAlso AuthService.LastAuthorizingSupervisor IsNot Nothing, New Nullable(Of Integer)(AuthService.LastAuthorizingSupervisor.Id), Nothing)
                If cajaService.RegistrarMovimiento(cajaActual.Id, userId, cboTipo.SelectedItem.ToString(), txtConcepto.Text.Trim(), numMonto.Value, "", errMsg, autorizadoPorId) Then
                    MessageBox.Show("Movimiento de caja registrado.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    LoadCajaData()
                Else
                    MessageBox.Show("Error: " & errMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        End Sub

        Private Sub BtnCerrarCaja_Click(sender As Object, e As EventArgs)
            If cajaActual Is Nothing Then Return

            Dim esperado As Decimal = cajaActual.MontoInicial + cajaActual.TotalVentasEfectivo + cajaActual.TotalIngresos - cajaActual.TotalEgresos
            Dim esAdmin As Boolean = AuthService.IsAdminOrManager

            Dim frmCierre As New Form() With {
                .Text = If(esAdmin, "Arqueo y Cierre de Caja (Auditoría)", "Cierre de Turno y Arqueo Ciego de Caja"),
                .Size = New Size(500, If(esAdmin, 370, 350)),
                .StartPosition = FormStartPosition.CenterParent,
                .FormBorderStyle = FormBorderStyle.FixedDialog,
                .MaximizeBox = False,
                .MinimizeBox = False,
                .BackColor = Color.White
            }

            Dim lblEsp As New Label() With {
                .Text = $"Efectivo que debería haber en el cajón: {esperado:C2}",
                .Font = UITheme.FontSubheading,
                .ForeColor = UITheme.ColorPrimary,
                .Location = New Point(30, 20),
                .AutoSize = True,
                .Visible = esAdmin
            }

            Dim lblInstruccion As New Label() With {
                .Text = "Por favor cuente el dinero físico del cajón e ingrese el total:",
                .Font = UITheme.FontRegular,
                .ForeColor = UITheme.ColorTextSecondary,
                .Location = New Point(30, 20),
                .AutoSize = True,
                .Visible = Not esAdmin
            }

            Dim lblReal As New Label() With {
                .Text = "Efectivo Real Contado en Billetes ($):",
                .Font = UITheme.FontBold,
                .Location = New Point(30, If(esAdmin, 65, 55)),
                .AutoSize = True
            }

            Dim numReal As New NumericUpDown() With {
                .Location = New Point(30, If(esAdmin, 90, 80)),
                .Size = New Size(200, 28),
                .Maximum = 10000000,
                .DecimalPlaces = 2,
                .Value = If(esAdmin, esperado, 0D)
            }

            Dim lblDif As New Label() With {
                .Text = "Diferencia: $ 0,00",
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorSuccess,
                .Location = New Point(250, If(esAdmin, 93, 83)),
                .AutoSize = True,
                .Visible = esAdmin
            }

            If esAdmin Then
                AddHandler numReal.ValueChanged, Sub()
                    Dim dif As Decimal = numReal.Value - esperado
                    If dif = 0 Then
                        lblDif.Text = "Exacto ($ 0,00)"
                        lblDif.ForeColor = UITheme.ColorSuccess
                    ElseIf dif > 0 Then
                        lblDif.Text = $"Sobrante: +{dif:C2}"
                        lblDif.ForeColor = UITheme.ColorInfo
                    Else
                        lblDif.Text = $"Faltante: {dif:C2}"
                        lblDif.ForeColor = UITheme.ColorDanger
                    End If
                End Sub
            End If

            Dim lblObs As New Label() With {.Text = "Observaciones de Cierre:", .Font = UITheme.FontBold, .Location = New Point(30, If(esAdmin, 135, 125)), .AutoSize = True}
            Dim txtObs As New TextBox() With {.Location = New Point(30, If(esAdmin, 160, 150)), .Size = New Size(420, 60), .Multiline = True}
            UITheme.StyleTextBox(txtObs)

            Dim btnConfirmar As New Button() With {.Text = "Confirmar cierre definitivo", .Location = New Point(30, If(esAdmin, 245, 230)), .Size = New Size(420, 42)}
            UITheme.StyleButton(btnConfirmar, "Danger")

            AddHandler btnConfirmar.Click, Sub()
                frmCierre.DialogResult = DialogResult.OK
                frmCierre.Close()
            End Sub

            frmCierre.Controls.AddRange({lblEsp, lblInstruccion, lblReal, numReal, lblDif, lblObs, txtObs, btnConfirmar})

            If frmCierre.ShowDialog() = DialogResult.OK Then
                Dim errMsg As String = ""
                If cajaService.CerrarCaja(cajaActual.Id, numReal.Value, txtObs.Text.Trim(), errMsg) Then
                    MessageBox.Show("Caja cerrada y arqueo archivado correctamente.", "Caja Cerrada", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    LoadCajaData()
                Else
                    MessageBox.Show("Error al cerrar caja: " & errMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        End Sub

    End Class
End Namespace
