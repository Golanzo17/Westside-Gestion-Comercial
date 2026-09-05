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
        Private lblMontoInicial As Label
        Private lblVentasEfectivo As Label
        Private lblVentasDigital As Label
        Private lblTotalIngresos As Label
        Private lblTotalEgresos As Label
        Private lblEfectivoEsperado As Label

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
                .Text = "💵 CAJA DIARIA, MOVIMIENTOS Y ARQUEO",
                .Font = UITheme.FontHeading,
                .ForeColor = Color.White,
                .AutoSize = True,
                .Location = New Point(15, 15)
            }
            pnlHeader.Controls.Add(lblTitle)

            ' Tarjetas KPI de Estado de Caja
            Dim pnlKpis As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 110,
                .BackColor = UITheme.ColorSurface,
                .Padding = New Padding(15, 10, 15, 10)
            }

            ' Tarjeta 1: Estado
            Dim cardEstado = UITheme.CreateKpiCard("ESTADO DE CAJA", "CERRADA", "Turno actual", UITheme.ColorDanger)
            cardEstado.Location = New Point(15, 5)
            lblEstadoCaja = CType(cardEstado.Controls(2), Label)
            pnlKpis.Controls.Add(cardEstado)

            ' Tarjeta 2: Ventas Efectivo
            Dim cardEfec = UITheme.CreateKpiCard("VENTAS EFECTIVO", "$ 0,00", "+ Monto inicial", UITheme.ColorSuccess)
            cardEfec.Location = New Point(250, 5)
            lblVentasEfectivo = CType(cardEfec.Controls(2), Label)
            pnlKpis.Controls.Add(cardEfec)

            ' Tarjeta 3: Ventas Digitales
            Dim cardDig = UITheme.CreateKpiCard("TARJETAS / QR", "$ 0,00", "Débito, crédito, transf.", UITheme.ColorInfo)
            cardDig.Location = New Point(485, 5)
            lblVentasDigital = CType(cardDig.Controls(2), Label)
            pnlKpis.Controls.Add(cardDig)

            ' Tarjeta 4: Efectivo Esperado
            Dim cardTotal = UITheme.CreateKpiCard("EFECTIVO EN CAJA", "$ 0,00", "Total esperado a rendir", UITheme.ColorPrimary)
            cardTotal.Location = New Point(720, 5)
            lblEfectivoEsperado = CType(cardTotal.Controls(2), Label)
            pnlKpis.Controls.Add(cardTotal)

            ' Barra de Acciones
            Dim pnlToolbar As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 55,
                .BackColor = Color.FromArgb(241, 245, 249),
                .Padding = New Padding(15, 10, 15, 10)
            }

            btnAbrirCaja = New Button() With {.Text = "🔓 Abrir Caja de Turno", .Location = New Point(15, 10), .Size = New Size(180, 34)}
            UITheme.StyleButton(btnAbrirCaja, "Primary")
            AddHandler btnAbrirCaja.Click, AddressOf BtnAbrirCaja_Click

            btnNuevoMovimiento = New Button() With {.Text = "💰 Ingreso / Retiro de Efectivo", .Location = New Point(205, 10), .Size = New Size(230, 34)}
            UITheme.StyleButton(btnNuevoMovimiento, "Secondary")
            AddHandler btnNuevoMovimiento.Click, AddressOf BtnNuevoMovimiento_Click

            btnCerrarCaja = New Button() With {.Text = "🔒 Arqueo y Cierre de Caja", .Location = New Point(445, 10), .Size = New Size(210, 34)}
            UITheme.StyleButton(btnCerrarCaja, "Danger")
            AddHandler btnCerrarCaja.Click, AddressOf BtnCerrarCaja_Click

            btnRefrescar = New Button() With {.Text = "🔄 Actualizar", .Location = New Point(665, 10), .Size = New Size(120, 34)}
            UITheme.StyleButton(btnRefrescar, "Secondary")
            AddHandler btnRefrescar.Click, Sub() LoadCajaData()

            pnlToolbar.Controls.AddRange({btnAbrirCaja, btnNuevoMovimiento, btnCerrarCaja, btnRefrescar})

            ' Grilla de Movimientos
            Dim pnlGrid As New Panel() With {.Dock = DockStyle.Fill, .Padding = New Padding(15)}
            dgvMovimientos = New DataGridView() With {.Dock = DockStyle.Fill}
            UITheme.StyleDataGrid(dgvMovimientos)
            ConfigurarColumnas()
            pnlGrid.Controls.Add(dgvMovimientos)

            ' Orden exacto de Docking: Fill primero, Toolbar segundo, KPIs tercero, Header último
            Me.Controls.Add(pnlGrid)
            Me.Controls.Add(pnlToolbar)
            Me.Controls.Add(pnlKpis)
            Me.Controls.Add(pnlHeader)
        End Sub

        Private Sub ConfigurarColumnas()
            dgvMovimientos.Columns.Clear()
            dgvMovimientos.Columns.Add("Hora", "Hora")
            dgvMovimientos.Columns("Hora").Width = 100

            dgvMovimientos.Columns.Add("Tipo", "Tipo")
            dgvMovimientos.Columns("Tipo").Width = 110

            dgvMovimientos.Columns.Add("Concepto", "Concepto / Detalle")
            dgvMovimientos.Columns("Concepto").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

            dgvMovimientos.Columns.Add("Monto", "Monto ($)")
            dgvMovimientos.Columns("Monto").Width = 130
            dgvMovimientos.Columns("Monto").DefaultCellStyle.Format = "C2"

            dgvMovimientos.Columns.Add("Usuario", "Registrado Por")
            dgvMovimientos.Columns("Usuario").Width = 160

            dgvMovimientos.Columns.Add("Referencia", "Comprobante Ref.")
            dgvMovimientos.Columns("Referencia").Width = 150
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

                lblVentasEfectivo.Text = cajaActual.TotalVentasEfectivo.ToString("C2")
                lblVentasDigital.Text = cajaActual.TotalVentasDigital.ToString("C2")

                Dim esperado As Decimal = cajaActual.MontoInicial + cajaActual.TotalVentasEfectivo + cajaActual.TotalIngresos - cajaActual.TotalEgresos
                lblEfectivoEsperado.Text = esperado.ToString("C2")

                ' Cargar movimientos
                Dim movs = cajaService.GetMovimientosCaja(cajaActual.Id)
                dgvMovimientos.Rows.Clear()
                For Each m In movs
                    dgvMovimientos.Rows.Add(m.Fecha.ToString("HH:mm:ss"), m.Tipo, m.Concepto, m.Monto, m.UsuarioNombre, m.Referencia)
                Next
            Else
                lblEstadoCaja.Text = "CERRADA"
                lblEstadoCaja.ForeColor = UITheme.ColorDanger
                btnAbrirCaja.Enabled = True
                btnNuevoMovimiento.Enabled = False
                btnCerrarCaja.Enabled = False

                lblVentasEfectivo.Text = "$ 0,00"
                lblVentasDigital.Text = "$ 0,00"
                lblEfectivoEsperado.Text = "$ 0,00"
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
                Dim userId As Integer = If(AuthService.CurrentUser IsNot Nothing, AuthService.CurrentUser.Id, 1)
                Dim errMsg As String = ""
                If cajaService.RegistrarMovimiento(cajaActual.Id, userId, cboTipo.SelectedItem.ToString(), txtConcepto.Text.Trim(), numMonto.Value, "", errMsg) Then
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

            Dim frmCierre As New Form() With {
                .Text = "Arqueo y Cierre de Caja",
                .Size = New Size(500, 360),
                .StartPosition = FormStartPosition.CenterParent,
                .FormBorderStyle = FormBorderStyle.FixedDialog,
                .MaximizeBox = False,
                .MinimizeBox = False,
                .BackColor = Color.White
            }

            Dim lblEsp As New Label() With {.Text = $"Efectivo que debería haber en el cajón: {esperado:C2}", .Font = UITheme.FontSubheading, .ForeColor = UITheme.ColorPrimary, .Location = New Point(30, 20), .AutoSize = True}
            Dim lblReal As New Label() With {.Text = "Efectivo Real Contado en Billetes ($):", .Font = UITheme.FontBold, .Location = New Point(30, 65), .AutoSize = True}
            Dim numReal As New NumericUpDown() With {.Location = New Point(30, 90), .Size = New Size(200, 28), .Maximum = 10000000, .DecimalPlaces = 2, .Value = esperado}

            Dim lblDif As New Label() With {.Text = "Diferencia: $ 0,00", .Font = UITheme.FontBold, .ForeColor = UITheme.ColorSuccess, .Location = New Point(250, 93), .AutoSize = True}

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

            Dim lblObs As New Label() With {.Text = "Observaciones de Cierre:", .Font = UITheme.FontBold, .Location = Point.Add(New Point(30, 135), New Size(0, 0)), .AutoSize = True}
            Dim txtObs As New TextBox() With {.Location = New Point(30, 160), .Size = New Size(420, 60), .Multiline = True}
            UITheme.StyleTextBox(txtObs)

            Dim btnConfirmar As New Button() With {.Text = "🔒 Confirmar Cierre Definitivo", .Location = New Point(30, 245), .Size = New Size(420, 42)}
            UITheme.StyleButton(btnConfirmar, "Danger")

            AddHandler btnConfirmar.Click, Sub()
                                               frmCierre.DialogResult = DialogResult.OK
                                               frmCierre.Close()
                                           End Sub

            frmCierre.Controls.AddRange({lblEsp, lblReal, numReal, lblDif, lblObs, txtObs, btnConfirmar})

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
