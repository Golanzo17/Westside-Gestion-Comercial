' ARCHIVO: FrmCaja.vb
' PROPÓSITO: Control de turnos de caja, registro de gastos y arqueo ciego de efectivo.
' Esta ventana garantiza el control sobre el dinero físico en el salón:
' 1. Visualización de Balances: Muestra en tarjetas el fondo inicial, las ventas en efectivo,
'    las ventas digitales (tarjetas), los egresos y el monto que debe haber en el cajón.
' 2. Seguridad en Retiros: Si el cajero registra un retiro o egreso de efectivo, el sistema
'    exige la autorización presencial de un Administrador o Gerente.
' 3. Arqueo y Cierre: Al finalizar el turno, pide contar los billetes en mano (montoReal),
'    calcula automáticamente si sobró o faltó plata y emite el balance impreso.
' 4. Pestaña de Historial: Permite a los supervisores consultar los cierres de cajas
'    anteriores para auditorías contables.


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

        ' Controles de Historial de Cajas
        Private tabControlCaja As TabControl
        Private dtpDesdeHistorial As DateTimePicker
        Private dtpHastaHistorial As DateTimePicker
        Private btnFiltrarHistorial As Button
        Private btnVerDetalleCaja As Button
        Private dgvHistorial As DataGridView
        Private lblResumenHistorial As Label

        Public Sub New()
            InitializeUI()
            LoadCajaData()
            LoadHistorialData()
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
                .Text = "Caja diaria, movimientos y arqueo",
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
                .BackColor = UITheme.ColorSurfaceMuted,
                .Padding = New Padding(15, 10, 15, 10)
            }

            btnAbrirCaja = New Button() With {.Text = "Abrir caja de turno", .Location = New Point(15, 10), .Size = New Size(180, 34)}
            UITheme.StyleButton(btnAbrirCaja, "Primary")
            AddHandler btnAbrirCaja.Click, AddressOf BtnAbrirCaja_Click

            btnNuevoMovimiento = New Button() With {.Text = "Ingreso / retiro de efectivo", .Location = New Point(205, 10), .Size = New Size(230, 34)}
            UITheme.StyleButton(btnNuevoMovimiento, "Secondary")
            AddHandler btnNuevoMovimiento.Click, AddressOf BtnNuevoMovimiento_Click

            btnCerrarCaja = New Button() With {.Text = "Arqueo y cierre de caja", .Location = New Point(445, 10), .Size = New Size(210, 34)}
            UITheme.StyleButton(btnCerrarCaja, "Danger")
            AddHandler btnCerrarCaja.Click, AddressOf BtnCerrarCaja_Click

            btnRefrescar = New Button() With {.Text = "Actualizar", .Location = New Point(665, 10), .Size = New Size(120, 34)}
            UITheme.StyleButton(btnRefrescar, "Secondary")
            AddHandler btnRefrescar.Click, Sub() LoadCajaData()

            pnlToolbar.Controls.AddRange({btnAbrirCaja, btnNuevoMovimiento, btnCerrarCaja, btnRefrescar})

            ' Grilla de Movimientos de Turno Actual
            Dim pnlGrid As New Panel() With {.Dock = DockStyle.Fill, .Padding = New Padding(15)}
            dgvMovimientos = New DataGridView() With {.Dock = DockStyle.Fill}
            UITheme.StyleDataGrid(dgvMovimientos)
            ConfigurarColumnas()
            pnlGrid.Controls.Add(dgvMovimientos)

            ' Pestañas de control de caja
            tabControlCaja = New TabControl() With {
                .Dock = DockStyle.Fill,
                .Padding = New Point(12, 6)
            }

            ' Pestaña 1: Turno Actual
            Dim tabTurnoActual As New TabPage("Caja diaria / Turno en curso") With {.BackColor = UITheme.ColorBackground}
            tabTurnoActual.Controls.Add(pnlGrid)
            tabTurnoActual.Controls.Add(pnlToolbar)
            tabTurnoActual.Controls.Add(pnlKpis)
            tabControlCaja.TabPages.Add(tabTurnoActual)

            ' Pestaña 2: Historial de Cajas Cerradas
            Dim tabHistorial As New TabPage("Historial de cajas cerradas y arqueos") With {.BackColor = UITheme.ColorBackground}

            Dim pnlFilterHistorial As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 55,
                .BackColor = UITheme.ColorSurface,
                .Padding = New Padding(15, 10, 15, 10)
            }

            Dim lblD As New Label() With {.Text = "Desde:", .Font = UITheme.FontBold, .Location = New Point(15, 16), .AutoSize = True}
            dtpDesdeHistorial = New DateTimePicker() With {.Location = New Point(70, 14), .Size = New Size(125, 26), .Format = DateTimePickerFormat.Short, .Value = DateTime.Today.AddDays(-30)}

            Dim lblH As New Label() With {.Text = "Hasta:", .Font = UITheme.FontBold, .Location = New Point(210, 16), .AutoSize = True}
            dtpHastaHistorial = New DateTimePicker() With {.Location = New Point(265, 14), .Size = New Size(125, 26), .Format = DateTimePickerFormat.Short, .Value = DateTime.Today}

            btnFiltrarHistorial = New Button() With {.Text = "Filtrar cajas", .Location = New Point(410, 12), .Size = New Size(130, 30)}
            UITheme.StyleButton(btnFiltrarHistorial, "Primary")
            AddHandler btnFiltrarHistorial.Click, Sub() LoadHistorialData()

            btnVerDetalleCaja = New Button() With {.Text = "Ver movimientos del turno", .Location = New Point(555, 12), .Size = New Size(220, 30)}
            UITheme.StyleButton(btnVerDetalleCaja, "Secondary")
            AddHandler btnVerDetalleCaja.Click, AddressOf BtnVerDetalleCaja_Click

            pnlFilterHistorial.Controls.AddRange({lblD, dtpDesdeHistorial, lblH, dtpHastaHistorial, btnFiltrarHistorial, btnVerDetalleCaja})

            Dim pnlGridHistorial As New Panel() With {.Dock = DockStyle.Fill, .Padding = New Padding(15)}
            dgvHistorial = New DataGridView() With {.Dock = DockStyle.Fill}
            UITheme.StyleDataGrid(dgvHistorial)
            ConfigurarColumnasHistorial()
            AddHandler dgvHistorial.CellDoubleClick, AddressOf DgvHistorial_CellDoubleClick
            AddHandler dgvHistorial.CellFormatting, AddressOf DgvHistorial_CellFormatting
            pnlGridHistorial.Controls.Add(dgvHistorial)

            Dim pnlFooterHistorial As New Panel() With {
                .Dock = DockStyle.Bottom,
                .Height = 35,
                .BackColor = UITheme.ColorSurfaceMuted,
                .Padding = New Padding(15, 8, 15, 8)
            }
            lblResumenHistorial = New Label() With {
                .Text = "Cargando historial de cajas...",
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorTextSecondary,
                .AutoSize = True,
                .Location = New Point(15, 8)
            }
            pnlFooterHistorial.Controls.Add(lblResumenHistorial)

            tabHistorial.Controls.Add(pnlGridHistorial)
            tabHistorial.Controls.Add(pnlFooterHistorial)
            tabHistorial.Controls.Add(pnlFilterHistorial)
            tabControlCaja.TabPages.Add(tabHistorial)
            If AuthService.IsVendor Then
                tabControlCaja.TabPages.Remove(tabHistorial)
            End If

            ' Orden exacto de Docking en FrmCaja: TabControl primero, Header último
            Me.Controls.Add(tabControlCaja)
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
                lblEfectivoEsperado.Text = If(AuthService.IsAdminOrManager, esperado.ToString("C2"), "Oculto (Arqueo Ciego)")

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
                    LoadHistorialData()
                Else
                    MessageBox.Show("Error al cerrar caja: " & errMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        End Sub

        Private Sub ConfigurarColumnasHistorial()
            dgvHistorial.Columns.Clear()
            dgvHistorial.Columns.Add("Id", "# Turno")
            dgvHistorial.Columns("Id").Width = 80

            dgvHistorial.Columns.Add("Usuario", "Vendedor Responsable")
            dgvHistorial.Columns("Usuario").Width = 160

            dgvHistorial.Columns.Add("Apertura", "Fecha Apertura")
            dgvHistorial.Columns("Apertura").Width = 135

            dgvHistorial.Columns.Add("Cierre", "Fecha Cierre")
            dgvHistorial.Columns("Cierre").Width = 135

            dgvHistorial.Columns.Add("Inicial", "Monto Inicial")
            dgvHistorial.Columns("Inicial").Width = 110
            dgvHistorial.Columns("Inicial").DefaultCellStyle.Format = "C2"

            dgvHistorial.Columns.Add("VentasEfec", "Ventas Efec.")
            dgvHistorial.Columns("VentasEfec").Width = 110
            dgvHistorial.Columns("VentasEfec").DefaultCellStyle.Format = "C2"

            dgvHistorial.Columns.Add("VentasDig", "Tarjetas/QR")
            dgvHistorial.Columns("VentasDig").Width = 110
            dgvHistorial.Columns("VentasDig").DefaultCellStyle.Format = "C2"

            dgvHistorial.Columns.Add("Ingresos", "Ingresos")
            dgvHistorial.Columns("Ingresos").Width = 100
            dgvHistorial.Columns("Ingresos").DefaultCellStyle.Format = "C2"

            dgvHistorial.Columns.Add("Egresos", "Egresos")
            dgvHistorial.Columns("Egresos").Width = 100
            dgvHistorial.Columns("Egresos").DefaultCellStyle.Format = "C2"

            dgvHistorial.Columns.Add("Esperado", "Monto Esperado")
            dgvHistorial.Columns("Esperado").Width = 120
            dgvHistorial.Columns("Esperado").DefaultCellStyle.Format = "C2"
            dgvHistorial.Columns("Esperado").Visible = AuthService.IsAdminOrManager

            dgvHistorial.Columns.Add("Real", "Efectivo Real")
            dgvHistorial.Columns("Real").Width = 110
            dgvHistorial.Columns("Real").DefaultCellStyle.Format = "C2"
            dgvHistorial.Columns("Real").DefaultCellStyle.Font = UITheme.FontBold

            dgvHistorial.Columns.Add("Diferencia", "Diferencia Arqueo")
            dgvHistorial.Columns("Diferencia").Width = 120
            dgvHistorial.Columns("Diferencia").DefaultCellStyle.Format = "C2"
            dgvHistorial.Columns("Diferencia").Visible = AuthService.IsAdminOrManager

            dgvHistorial.Columns.Add("Observaciones", "Observaciones")
            dgvHistorial.Columns("Observaciones").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        End Sub

        Public Sub LoadHistorialData()
            Dim desde = dtpDesdeHistorial.Value.Date
            Dim hasta = dtpHastaHistorial.Value.Date
            Dim lista = cajaService.GetHistorialCajas(desde, hasta)

            dgvHistorial.Rows.Clear()
            Dim totalRecaudado As Decimal = 0D

            For Each c In lista
                Dim cierreStr = If(c.FechaCierre.HasValue, c.FechaCierre.Value.ToString("dd/MM/yyyy HH:mm"), "Sin cerrar")
                Dim difVal As Object = If(c.Diferencia.HasValue, c.Diferencia.Value, Nothing)
                Dim realVal As Object = If(c.MontoReal.HasValue, c.MontoReal.Value, Nothing)

                dgvHistorial.Rows.Add(
                    c.Id,
                    c.UsuarioNombre,
                    c.FechaApertura.ToString("dd/MM/yyyy HH:mm"),
                    cierreStr,
                    c.MontoInicial,
                    c.TotalVentasEfectivo,
                    c.TotalVentasDigital,
                    c.TotalIngresos,
                    c.TotalEgresos,
                    c.MontoEsperado,
                    realVal,
                    difVal,
                    c.Observaciones
                )

                totalRecaudado += (c.TotalVentasEfectivo + c.TotalVentasDigital)
            Next

            lblResumenHistorial.Text = $"Total de turnos auditados: {lista.Count} | Recaudación del período: {totalRecaudado:C2}"
        End Sub

        Private Sub DgvHistorial_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)
            If e.RowIndex >= 0 Then
                If dgvHistorial.Columns(e.ColumnIndex).Name = "Diferencia" Then
                    If e.Value IsNot Nothing AndAlso Not IsDBNull(e.Value) Then
                        Dim dif As Decimal = 0D
                        If Decimal.TryParse(e.Value.ToString(), dif) Then
                            If dif = 0 Then
                                e.CellStyle.ForeColor = UITheme.ColorSuccess
                                e.CellStyle.Font = UITheme.FontBold
                            ElseIf dif > 0 Then
                                e.CellStyle.ForeColor = UITheme.ColorInfo
                                e.CellStyle.Font = UITheme.FontBold
                            Else
                                e.CellStyle.ForeColor = UITheme.ColorDanger
                                e.CellStyle.Font = UITheme.FontBold
                            End If
                        End If
                    End If
                End If
            End If
        End Sub

        Private Sub BtnVerDetalleCaja_Click(sender As Object, e As EventArgs)
            VerDetalleCajaSeleccionada()
        End Sub

        Private Sub DgvHistorial_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
            If e.RowIndex >= 0 Then
                VerDetalleCajaSeleccionada()
            End If
        End Sub

        Private Sub VerDetalleCajaSeleccionada()
            If dgvHistorial.CurrentRow IsNot Nothing Then
                Dim cajaId As Integer = Convert.ToInt32(dgvHistorial.CurrentRow.Cells("Id").Value)
                Dim vendedor As String = dgvHistorial.CurrentRow.Cells("Usuario").Value.ToString()
                Dim movs = cajaService.GetMovimientosCaja(cajaId)

                Using dlg As New Form()
                    dlg.Text = $"Movimientos de Efectivo - Turno #{cajaId} ({vendedor})"
                    dlg.Size = New Size(750, 480)
                    dlg.StartPosition = FormStartPosition.CenterParent
                    dlg.FormBorderStyle = FormBorderStyle.FixedDialog
                    dlg.MaximizeBox = False
                    dlg.MinimizeBox = False
                    dlg.BackColor = UITheme.ColorBackground

                    Dim pnlTop As New Panel() With {
                        .Dock = DockStyle.Top,
                        .Height = 50,
                        .BackColor = UITheme.ColorSecondary,
                        .Padding = New Padding(15, 12, 15, 10)
                    }
                    Dim lblT As New Label() With {
                        .Text = $"DETALLE DE MOVIMIENTOS - TURNO #{cajaId} ({vendedor})",
                        .Font = UITheme.FontSubheading,
                        .ForeColor = Color.White,
                        .AutoSize = True,
                        .Location = New Point(15, 14)
                    }
                    pnlTop.Controls.Add(lblT)

                    Dim dgvDet As New DataGridView() With {.Dock = DockStyle.Fill}
                    UITheme.StyleDataGrid(dgvDet)
                    dgvDet.Columns.Add("Hora", "Hora")
                    dgvDet.Columns("Hora").Width = 120
                    dgvDet.Columns.Add("Tipo", "Tipo")
                    dgvDet.Columns("Tipo").Width = 100
                    dgvDet.Columns.Add("Concepto", "Concepto")
                    dgvDet.Columns("Concepto").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                    dgvDet.Columns.Add("Monto", "Monto")
                    dgvDet.Columns("Monto").Width = 120
                    dgvDet.Columns("Monto").DefaultCellStyle.Format = "C2"
                    dgvDet.Columns.Add("Usuario", "Registrado Por")
                    dgvDet.Columns("Usuario").Width = 140

                    For Each m In movs
                        dgvDet.Rows.Add(m.Fecha.ToString("dd/MM/yyyy HH:mm"), m.Tipo, m.Concepto, m.Monto, m.UsuarioNombre)
                    Next

                    Dim pnlBot As New Panel() With {.Dock = DockStyle.Bottom, .Height = 50, .BackColor = UITheme.ColorSurfaceMuted}
                    Dim btnCerrar As New Button() With {.Text = "Cerrar", .Location = New Point(620, 10), .Size = New Size(100, 32)}
                    UITheme.StyleButton(btnCerrar, "Secondary")
                    AddHandler btnCerrar.Click, Sub() dlg.Close()
                    pnlBot.Controls.Add(btnCerrar)

                    dlg.Controls.AddRange({dgvDet, pnlBot, pnlTop})
                    dlg.ShowDialog(Me)
                End Using
            Else
                MessageBox.Show("Por favor seleccione un turno del historial.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End Sub

    End Class
End Namespace
