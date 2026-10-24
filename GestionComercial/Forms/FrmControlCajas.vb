' ARCHIVO: FrmControlCajas.vb
' PROPÓSITO: Tablero Gerencial de Auditoría y Control de Cierres y Arqueos de Caja.
' Diseñado para que el Gerente supervise el desempeño de todos los vendedores:
' 1. Filtros Multidimensionales: Por rango de fechas (con accesos rápidos), vendedor, estado y discrepancia.
' 2. Indicadores Clave de Control (KPI Cards): Cantidad de turnos, facturación, esperado vs real y balance neto de arqueo.
' 3. Grilla de Auditoría: Comparación visual de dinero esperado vs puesto por el vendedor, marcando en rojo faltantes,
'    en azul sobrantes y en verde arqueos exactos.
' 4. Acciones de Auditoría: Acceso al desglose de ventas, movimientos manuales de caja y emisión de actas firmadas.
' 5. Exportación a CSV/Excel: Descarga rápida de la auditoría para conciliación contable externa.

Imports System.Drawing
Imports System.IO
Imports System.Text
Imports System.Windows.Forms
Imports GestionComercial.Models
Imports GestionComercial.Services
Imports GestionComercial.UI

Namespace Forms
    Public Class FrmControlCajas
        Inherits Form

        Private ReadOnly cajaService As New CajaService()
        Private ReadOnly usuarioService As New UsuarioService()

        ' Controles de Filtros
        Private dtpDesde As DateTimePicker
        Private dtpHasta As DateTimePicker
        Private cboRapidoFechas As ComboBox
        Private cboVendedor As ComboBox
        Private cboEstado As ComboBox
        Private cboDiscrepancia As ComboBox
        Private btnFiltrar As Button
        Private btnLimpiar As Button
        Private btnExportarCSV As Button

        ' Tarjetas KPI
        Private lblKpiTurnos As Label
        Private lblKpiTurnosSub As Label
        Private lblKpiRecaudacion As Label
        Private lblKpiRecaudacionSub As Label
        Private lblKpiEsperado As Label
        Private lblKpiEsperadoSub As Label
        Private lblKpiBalance As Label
        Private lblKpiBalanceSub As Label
        Private cardBalance As Panel

        ' Grilla de Cajas y Turnos
        Private dgvCajas As DataGridView
        Private lblStatusFooter As Label

        ' Botones de Acción
        Private btnAuditar As Button
        Private btnVerActa As Button
        Private btnVerMovimientos As Button
        Private btnVerTickets As Button
        Private btnRefrescar As Button

        Private listaCajasActual As New List(Of Caja)()
        Private isLoading As Boolean = True

        Public Sub New()
            InitializeUI()
            CargarFiltroVendedores()
            isLoading = False
            CargarDatosAuditoria()
        End Sub

        Private Sub InitializeUI()
            Me.Text = "Control de Cierres y Arqueos de Caja - Auditoría Gerencial"
            Me.Size = New Size(1100, 720)
            Me.MinimumSize = New Size(980, 620)
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.BackColor = UITheme.ColorBackground
            Me.Font = UITheme.FontRegular

            ' ─── Panel Encabezado Institucional ───
            Dim pnlHeader As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 65,
                .BackColor = UITheme.ColorSecondary,
                .Padding = New Padding(20, 12, 20, 12)
            }

            Dim lblTitle As New Label() With {
                .Text = "Control y Auditoría de Cierres y Arqueos de Caja",
                .Font = UITheme.FontHeading,
                .ForeColor = Color.White,
                .AutoSize = True,
                .Location = New Point(18, 10)
            }

            Dim lblSubtitle As New Label() With {
                .Text = "Supervisión de turnos de todos los vendedores, conciliación de dinero esperado vs. real y control de discrepancias.",
                .Font = UITheme.FontSmall,
                .ForeColor = UITheme.ColorTextOnDark,
                .AutoSize = True,
                .Location = New Point(20, 38)
            }

            pnlHeader.Controls.AddRange({lblTitle, lblSubtitle})

            ' ─── Panel de Filtros ───
            Dim pnlFiltros As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 105,
                .BackColor = UITheme.ColorSurface,
                .Padding = New Padding(15, 10, 15, 10)
            }
            AddHandler pnlFiltros.Paint, Sub(s, ev)
                                             Using pen As New Pen(UITheme.ColorBorder, 1)
                                                 ev.Graphics.DrawLine(pen, 0, pnlFiltros.Height - 1, pnlFiltros.Width, pnlFiltros.Height - 1)
                                             End Using
                                         End Sub

            ' Fila 1 de Filtros: Fechas y accesos rápidos
            Dim lblF1 As New Label() With {.Text = "Período:", .Font = UITheme.FontSmallBold, .ForeColor = UITheme.ColorTextSecondary, .Location = New Point(15, 12), .AutoSize = True}
            cboRapidoFechas = New ComboBox() With {.Location = New Point(15, 30), .Size = New Size(140, 26), .DropDownStyle = ComboBoxStyle.DropDownList}
            cboRapidoFechas.Items.AddRange({"Hoy", "Ayer", "Esta semana", "Este mes", "Últimos 30 días", "Últimos 90 días", "Todo el historial"})
            cboRapidoFechas.SelectedIndex = 5 ' Últimos 90 días por defecto
            AddHandler cboRapidoFechas.SelectedIndexChanged, AddressOf CboRapidoFechas_SelectedIndexChanged

            Dim lblF2 As New Label() With {.Text = "Desde:", .Font = UITheme.FontSmallBold, .ForeColor = UITheme.ColorTextSecondary, .Location = New Point(165, 12), .AutoSize = True}
            dtpDesde = New DateTimePicker() With {.Location = New Point(165, 30), .Size = New Size(115, 26), .Format = DateTimePickerFormat.Short, .Value = DateTime.Today.AddDays(-90)}

            Dim lblF3 As New Label() With {.Text = "Hasta:", .Font = UITheme.FontSmallBold, .ForeColor = UITheme.ColorTextSecondary, .Location = New Point(290, 12), .AutoSize = True}
            dtpHasta = New DateTimePicker() With {.Location = New Point(290, 30), .Size = New Size(115, 26), .Format = DateTimePickerFormat.Short, .Value = DateTime.Today}

            ' Fila 1 de Filtros (cont): Vendedor
            Dim lblF4 As New Label() With {.Text = "Vendedor / Cajero:", .Font = UITheme.FontSmallBold, .ForeColor = UITheme.ColorTextSecondary, .Location = New Point(415, 12), .AutoSize = True}
            cboVendedor = New ComboBox() With {.Location = New Point(415, 30), .Size = New Size(180, 26), .DropDownStyle = ComboBoxStyle.DropDownList}

            ' Fila 1 de Filtros (cont): Estado de caja
            Dim lblF5 As New Label() With {.Text = "Estado de Turno:", .Font = UITheme.FontSmallBold, .ForeColor = UITheme.ColorTextSecondary, .Location = New Point(605, 12), .AutoSize = True}
            cboEstado = New ComboBox() With {.Location = New Point(605, 30), .Size = New Size(140, 26), .DropDownStyle = ComboBoxStyle.DropDownList}
            cboEstado.Items.AddRange({"Todos los estados", "Solo cerradas", "Solo abiertas (en curso)"})
            cboEstado.SelectedIndex = 0

            ' Fila 1 de Filtros (cont): Tipo de discrepancia
            Dim lblF6 As New Label() With {.Text = "Resultado Arqueo:", .Font = UITheme.FontSmallBold, .ForeColor = UITheme.ColorTextSecondary, .Location = New Point(755, 12), .AutoSize = True}
            cboDiscrepancia = New ComboBox() With {.Location = New Point(755, 30), .Size = New Size(190, 26), .DropDownStyle = ComboBoxStyle.DropDownList}
            cboDiscrepancia.Items.AddRange({
                "Todos los resultados",
                "Con descuadre (Diferencia <> 0)",
                "Solo con faltante (Diferencia < 0)",
                "Solo con sobrante (Diferencia > 0)",
                "Arqueo exacto (Diferencia = 0)"
            })
            cboDiscrepancia.SelectedIndex = 0

            ' Botones de Filtrado
            btnFiltrar = New Button() With {.Text = "🔍 Filtrar", .Location = New Point(15, 66), .Size = New Size(110, 30)}
            UITheme.StyleButton(btnFiltrar, "Primary")
            AddHandler btnFiltrar.Click, Sub() CargarDatosAuditoria()

            btnLimpiar = New Button() With {.Text = "Restablecer", .Location = New Point(135, 66), .Size = New Size(110, 30)}
            UITheme.StyleButton(btnLimpiar, "Secondary")
            AddHandler btnLimpiar.Click, AddressOf BtnLimpiar_Click

            btnExportarCSV = New Button() With {.Text = "📊 Exportar a Excel / CSV", .Location = New Point(255, 66), .Size = New Size(190, 30)}
            UITheme.StyleButton(btnExportarCSV, "Secondary")
            AddHandler btnExportarCSV.Click, AddressOf BtnExportarCSV_Click

            btnRefrescar = New Button() With {.Text = "🔄 Actualizar", .Location = New Point(455, 66), .Size = New Size(120, 30)}
            UITheme.StyleButton(btnRefrescar, "Ghost")
            AddHandler btnRefrescar.Click, Sub() CargarDatosAuditoria()

            pnlFiltros.Controls.AddRange({
                lblF1, cboRapidoFechas, lblF2, dtpDesde, lblF3, dtpHasta,
                lblF4, cboVendedor, lblF5, cboEstado, lblF6, cboDiscrepancia,
                btnFiltrar, btnLimpiar, btnExportarCSV, btnRefrescar
            })

            ' ─── Fila Superior de Tarjetas KPI (TableLayoutPanel responsivo) ───
            Dim pnlKpis As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 125,
                .BackColor = UITheme.ColorSurfaceMuted,
                .Padding = New Padding(10, 5, 10, 5)
            }

            ' TableLayoutPanel de 4 columnas iguales → se adapta a cualquier resolución
            Dim tlpKpis As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 4,
                .RowCount = 1,
                .BackColor = Color.Transparent
            }
            tlpKpis.ColumnStyles.Clear()
            For i = 0 To 3
                tlpKpis.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))
            Next
            tlpKpis.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            Dim card1 = UITheme.CreateKpiCard("TURNOS AUDITADOS", "0", "0 cerrados, 0 en curso", UITheme.ColorPrimary)
            card1.Dock = DockStyle.Fill
            card1.Margin = New Padding(5, 5, 5, 5)
            lblKpiTurnos = CType(card1.Controls(2), Label)
            lblKpiTurnosSub = CType(card1.Controls(3), Label)

            Dim card2 = UITheme.CreateKpiCard("RECAUDACIÓN TOTAL", "$ 0,00", "Efec: $ 0,00 | Digital: $ 0,00", UITheme.ColorSuccess)
            card2.Dock = DockStyle.Fill
            card2.Margin = New Padding(5, 5, 5, 5)
            lblKpiRecaudacion = CType(card2.Controls(2), Label)
            lblKpiRecaudacionSub = CType(card2.Controls(3), Label)

            Dim card3 = UITheme.CreateKpiCard("EFECTIVO ESPERADO", "$ 0,00", "Declarado: $ 0,00", UITheme.ColorInfo)
            card3.Dock = DockStyle.Fill
            card3.Margin = New Padding(5, 5, 5, 5)
            lblKpiEsperado = CType(card3.Controls(2), Label)
            lblKpiEsperadoSub = CType(card3.Controls(3), Label)

            cardBalance = UITheme.CreateKpiCard("BALANCE DE ARQUEO", "$ 0,00", "Faltantes: $ 0 | Sobrantes: $ 0", UITheme.ColorSuccess)
            cardBalance.Dock = DockStyle.Fill
            cardBalance.Margin = New Padding(5, 5, 5, 5)
            lblKpiBalance = CType(cardBalance.Controls(2), Label)
            lblKpiBalanceSub = CType(cardBalance.Controls(3), Label)

            tlpKpis.Controls.Add(card1, 0, 0)
            tlpKpis.Controls.Add(card2, 1, 0)
            tlpKpis.Controls.Add(card3, 2, 0)
            tlpKpis.Controls.Add(cardBalance, 3, 0)
            pnlKpis.Controls.Add(tlpKpis)

            ' ─── Panel Inferior de Acciones y Footer ───
            Dim pnlFooter As New Panel() With {
                .Dock = DockStyle.Bottom,
                .Height = 55,
                .BackColor = UITheme.ColorSurfaceMuted,
                .Padding = New Padding(15, 10, 15, 10)
            }

            btnAuditar = New Button() With {.Text = "🔍 Auditar Turno Seleccionado", .Location = New Point(15, 10), .Size = New Size(230, 35)}
            UITheme.StyleButton(btnAuditar, "Primary")
            AddHandler btnAuditar.Click, AddressOf BtnAuditar_Click

            btnVerActa = New Button() With {.Text = "📄 Ver Acta Oficial", .Location = New Point(255, 10), .Size = New Size(160, 35)}
            UITheme.StyleButton(btnVerActa, "Secondary")
            AddHandler btnVerActa.Click, AddressOf BtnVerActa_Click

            btnVerMovimientos = New Button() With {.Text = "💵 Gastos y Retiros", .Location = New Point(425, 10), .Size = New Size(160, 35)}
            UITheme.StyleButton(btnVerMovimientos, "Secondary")
            AddHandler btnVerMovimientos.Click, AddressOf BtnVerMovimientos_Click

            btnVerTickets = New Button() With {.Text = "🧾 Tickets de Venta", .Location = New Point(595, 10), .Size = New Size(160, 35)}
            UITheme.StyleButton(btnVerTickets, "Secondary")
            AddHandler btnVerTickets.Click, AddressOf BtnVerTickets_Click

            lblStatusFooter = New Label() With {
                .Text = "Haga doble clic en cualquier turno para abrir la auditoría completa.",
                .Font = UITheme.FontSmall,
                .ForeColor = UITheme.ColorTextSecondary,
                .AutoSize = True,
                .Location = New Point(770, 18),
                .Anchor = AnchorStyles.Top Or AnchorStyles.Right
            }

            pnlFooter.Controls.AddRange({btnAuditar, btnVerActa, btnVerMovimientos, btnVerTickets, lblStatusFooter})

            ' ─── Grilla Principal de Cajas ───
            Dim pnlGrid As New Panel() With {.Dock = DockStyle.Fill, .Padding = New Padding(15)}
            dgvCajas = New DataGridView() With {.Dock = DockStyle.Fill}
            UITheme.StyleDataGrid(dgvCajas)
            ConfigurarColumnasGrilla()
            AddHandler dgvCajas.CellDoubleClick, AddressOf DgvCajas_CellDoubleClick
            AddHandler dgvCajas.CellFormatting, AddressOf DgvCajas_CellFormatting
            pnlGrid.Controls.Add(dgvCajas)

            ' Orden de controles
            Me.Controls.Add(pnlGrid)
            Me.Controls.Add(pnlFooter)
            Me.Controls.Add(pnlKpis)
            Me.Controls.Add(pnlFiltros)
            Me.Controls.Add(pnlHeader)
        End Sub

        Private Sub ConfigurarColumnasGrilla()
            dgvCajas.Columns.Clear()
            ' Modo Fill: las columnas se distribuyen proporcionalmente al ancho disponible
            dgvCajas.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            ' ── Columna: Nº (centrada) ──
            dgvCajas.Columns.Add("Id", "Nº")
            dgvCajas.Columns("Id").FillWeight = 4
            dgvCajas.Columns("Id").MinimumWidth = 45
            dgvCajas.Columns("Id").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvCajas.Columns("Id").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            ' ── Columna: Vendedor / Cajero (izquierda) ──
            dgvCajas.Columns.Add("Usuario", "Vendedor / Cajero")
            dgvCajas.Columns("Usuario").FillWeight = 11
            dgvCajas.Columns("Usuario").MinimumWidth = 110
            dgvCajas.Columns("Usuario").DefaultCellStyle.Font = UITheme.FontBold
            dgvCajas.Columns("Usuario").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvCajas.Columns("Usuario").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft

            ' ── Columna: Estado (centrada) ──
            dgvCajas.Columns.Add("Estado", "Estado")
            dgvCajas.Columns("Estado").FillWeight = 7
            dgvCajas.Columns("Estado").MinimumWidth = 70
            dgvCajas.Columns("Estado").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvCajas.Columns("Estado").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            ' ── Columna: Fecha Apertura (centrada) ──
            dgvCajas.Columns.Add("Apertura", "Fecha Apertura")
            dgvCajas.Columns("Apertura").FillWeight = 9
            dgvCajas.Columns("Apertura").MinimumWidth = 100
            dgvCajas.Columns("Apertura").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvCajas.Columns("Apertura").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            ' ── Columna: Fecha Cierre (centrada) ──
            dgvCajas.Columns.Add("Cierre", "Fecha Cierre")
            dgvCajas.Columns("Cierre").FillWeight = 9
            dgvCajas.Columns("Cierre").MinimumWidth = 100
            dgvCajas.Columns("Cierre").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvCajas.Columns("Cierre").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            ' ── Columna: Duración (centrada) ──
            dgvCajas.Columns.Add("Duracion", "Duración")
            dgvCajas.Columns("Duracion").FillWeight = 6
            dgvCajas.Columns("Duracion").MinimumWidth = 65
            dgvCajas.Columns("Duracion").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvCajas.Columns("Duracion").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            ' ── Columna: Fondo Inicial (derecha) ──
            dgvCajas.Columns.Add("Inicial", "Fondo Inicial")
            dgvCajas.Columns("Inicial").FillWeight = 8
            dgvCajas.Columns("Inicial").MinimumWidth = 85
            dgvCajas.Columns("Inicial").DefaultCellStyle.Format = "C2"
            dgvCajas.Columns("Inicial").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvCajas.Columns("Inicial").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight

            ' ── Columna: Vtas. Efectivo (derecha) ──
            dgvCajas.Columns.Add("VentasEfec", "Vtas. Efectivo")
            dgvCajas.Columns("VentasEfec").FillWeight = 8
            dgvCajas.Columns("VentasEfec").MinimumWidth = 85
            dgvCajas.Columns("VentasEfec").DefaultCellStyle.Format = "C2"
            dgvCajas.Columns("VentasEfec").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvCajas.Columns("VentasEfec").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight

            ' ── Columna: Tarjetas / QR (derecha) ──
            dgvCajas.Columns.Add("VentasDig", "Tarjetas / QR")
            dgvCajas.Columns("VentasDig").FillWeight = 8
            dgvCajas.Columns("VentasDig").MinimumWidth = 85
            dgvCajas.Columns("VentasDig").DefaultCellStyle.Format = "C2"
            dgvCajas.Columns("VentasDig").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvCajas.Columns("VentasDig").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight

            ' ── Columna: Ingresos Extra (derecha) ──
            dgvCajas.Columns.Add("Ingresos", "Ingresos Extra")
            dgvCajas.Columns("Ingresos").FillWeight = 7
            dgvCajas.Columns("Ingresos").MinimumWidth = 80
            dgvCajas.Columns("Ingresos").DefaultCellStyle.Format = "C2"
            dgvCajas.Columns("Ingresos").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvCajas.Columns("Ingresos").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight

            ' ── Columna: Retiros / Gastos (derecha) ──
            dgvCajas.Columns.Add("Egresos", "Retiros / Gastos")
            dgvCajas.Columns("Egresos").FillWeight = 8
            dgvCajas.Columns("Egresos").MinimumWidth = 85
            dgvCajas.Columns("Egresos").DefaultCellStyle.Format = "C2"
            dgvCajas.Columns("Egresos").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvCajas.Columns("Egresos").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight

            ' ── Columna: Monto Esperado (derecha) ──
            dgvCajas.Columns.Add("Esperado", "Monto Esperado")
            dgvCajas.Columns("Esperado").FillWeight = 9
            dgvCajas.Columns("Esperado").MinimumWidth = 90
            dgvCajas.Columns("Esperado").DefaultCellStyle.Format = "C2"
            dgvCajas.Columns("Esperado").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvCajas.Columns("Esperado").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight

            ' ── Columna: Físico Declarado (derecha, negrita) ──
            dgvCajas.Columns.Add("Real", "Físico Declarado")
            dgvCajas.Columns("Real").FillWeight = 9
            dgvCajas.Columns("Real").MinimumWidth = 90
            dgvCajas.Columns("Real").DefaultCellStyle.Format = "C2"
            dgvCajas.Columns("Real").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvCajas.Columns("Real").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvCajas.Columns("Real").DefaultCellStyle.Font = UITheme.FontBold

            ' ── Columna: Diferencia Arqueo (derecha, negrita) ──
            dgvCajas.Columns.Add("Diferencia", "Diferencia Arqueo")
            dgvCajas.Columns("Diferencia").FillWeight = 9
            dgvCajas.Columns("Diferencia").MinimumWidth = 90
            dgvCajas.Columns("Diferencia").DefaultCellStyle.Format = "C2"
            dgvCajas.Columns("Diferencia").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvCajas.Columns("Diferencia").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvCajas.Columns("Diferencia").DefaultCellStyle.Font = UITheme.FontBold

            ' ── Columna: Observaciones (relleno restante, izquierda) ──
            dgvCajas.Columns.Add("Observaciones", "Observaciones del Cierre")
            dgvCajas.Columns("Observaciones").FillWeight = 10
            dgvCajas.Columns("Observaciones").MinimumWidth = 80
            dgvCajas.Columns("Observaciones").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvCajas.Columns("Observaciones").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
        End Sub

        Private Class ComboItem
            Public Property Id As Integer
            Public Property Texto As String

            Public Sub New(id As Integer, texto As String)
                Me.Id = id
                Me.Texto = texto
            End Sub

            Public Overrides Function ToString() As String
                Return Texto
            End Function
        End Class

        Private Sub CargarFiltroVendedores()
            cboVendedor.Items.Clear()
            cboVendedor.Items.Add(New ComboItem(0, "Todos los vendedores"))

            Dim usuarios = usuarioService.GetUsuarios(soloActivos:=False)
            For Each u In usuarios
                cboVendedor.Items.Add(New ComboItem(u.Id, $"{u.NombreCompleto} ({u.Rol})"))
            Next

            cboVendedor.SelectedIndex = 0
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
            CargarDatosAuditoria()
        End Sub

        Private Sub BtnLimpiar_Click(sender As Object, e As EventArgs)
            isLoading = True
            cboRapidoFechas.SelectedIndex = 5
            dtpDesde.Value = DateTime.Today.AddDays(-90)
            dtpHasta.Value = DateTime.Today
            cboVendedor.SelectedIndex = 0
            cboEstado.SelectedIndex = 0
            cboDiscrepancia.SelectedIndex = 0
            isLoading = False
            CargarDatosAuditoria()
        End Sub

        Public Sub CargarDatosAuditoria()
            Dim desde = dtpDesde.Value.Date
            Dim hasta = dtpHasta.Value.Date

            Dim itemVend = TryCast(cboVendedor.SelectedItem, ComboItem)
            Dim userFiltro As Nullable(Of Integer) = If(itemVend IsNot Nothing AndAlso itemVend.Id > 0, itemVend.Id, Nothing)

            Dim estadoFiltro As String = Nothing
            If cboEstado.SelectedIndex = 1 Then
                estadoFiltro = "Cerrada"
            ElseIf cboEstado.SelectedIndex = 2 Then
                estadoFiltro = "Abierta"
            End If

            Dim tipoDifFiltro As String = Nothing
            Select Case cboDiscrepancia.SelectedIndex
                Case 1 : tipoDifFiltro = "CON_DESCUADRE"
                Case 2 : tipoDifFiltro = "FALTANTE"
                Case 3 : tipoDifFiltro = "SOBRANTE"
                Case 4 : tipoDifFiltro = "EXACTO"
            End Select

            listaCajasActual = cajaService.GetHistorialCajasAvanzado(desde, hasta, userFiltro, estadoFiltro, tipoDifFiltro)

            ' Cargar grilla
            dgvCajas.Rows.Clear()

            Dim cantCerradas = 0
            Dim cantAbiertas = 0
            Dim totalVtasEfec = 0D
            Dim totalVtasDig = 0D
            Dim totalEsperado = 0D
            Dim totalReal = 0D
            Dim totalFaltantes = 0D
            Dim totalSobrantes = 0D

            For Each c In listaCajasActual
                If c.Estado = "Cerrada" Then
                    cantCerradas += 1
                Else
                    cantAbiertas += 1
                End If

                totalVtasEfec += c.TotalVentasEfectivo
                totalVtasDig += c.TotalVentasDigital
                totalEsperado += c.MontoEsperado

                If c.MontoReal.HasValue Then
                    totalReal += c.MontoReal.Value
                End If

                If c.Diferencia.HasValue Then
                    Dim dif = c.Diferencia.Value
                    If dif < 0 Then
                        totalFaltantes += Math.Abs(dif)
                    ElseIf dif > 0 Then
                        totalSobrantes += dif
                    End If
                End If

                Dim cierreStr = If(c.FechaCierre.HasValue, c.FechaCierre.Value.ToString("dd/MM/yyyy HH:mm"), "En curso")
                Dim realObj As Object = If(c.MontoReal.HasValue, c.MontoReal.Value, "--")
                Dim difObj As Object = If(c.Diferencia.HasValue, c.Diferencia.Value, "--")

                dgvCajas.Rows.Add(
                    c.Id,
                    c.UsuarioNombre,
                    c.Estado,
                    c.FechaApertura.ToString("dd/MM/yyyy HH:mm"),
                    cierreStr,
                    c.DuracionTexto,
                    c.MontoInicial,
                    c.TotalVentasEfectivo,
                    c.TotalVentasDigital,
                    c.TotalIngresos,
                    c.TotalEgresos,
                    c.MontoEsperado,
                    realObj,
                    difObj,
                    c.Observaciones
                )
            Next

            ' Actualizar KPIs
            lblKpiTurnos.Text = listaCajasActual.Count.ToString()
            lblKpiTurnosSub.Text = $"{cantCerradas} cerrados, {cantAbiertas} en curso"

            Dim granTotal = totalVtasEfec + totalVtasDig
            lblKpiRecaudacion.Text = granTotal.ToString("C2")
            lblKpiRecaudacionSub.Text = $"Efec: {totalVtasEfec:C2} | Digital: {totalVtasDig:C2}"

            lblKpiEsperado.Text = totalEsperado.ToString("C2")
            lblKpiEsperadoSub.Text = $"Declarado: {totalReal:C2}"

            Dim balanceNeto = totalSobrantes - totalFaltantes
            If balanceNeto = 0 Then
                lblKpiBalance.Text = "$ 0,00 (Exacto)"
                lblKpiBalance.ForeColor = UITheme.ColorSuccess
            ElseIf balanceNeto > 0 Then
                lblKpiBalance.Text = $"+{balanceNeto:C2}"
                lblKpiBalance.ForeColor = UITheme.ColorInfo
            Else
                lblKpiBalance.Text = $"-{Math.Abs(balanceNeto):C2}"
                lblKpiBalance.ForeColor = UITheme.ColorDanger
            End If
            lblKpiBalanceSub.Text = $"Faltantes: -{totalFaltantes:C2} | Sobrantes: +{totalSobrantes:C2}"

            lblStatusFooter.Text = $"Mostrando {listaCajasActual.Count} turnos auditados  |  Recaudación total: {granTotal:C2}"
        End Sub

        Private Sub DgvCajas_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)
            If e.RowIndex < 0 Then Return
            Dim colName = dgvCajas.Columns(e.ColumnIndex).Name

            If colName = "Estado" Then
                Dim estado = dgvCajas.Rows(e.RowIndex).Cells("Estado").Value?.ToString()
                If estado = "Abierta" Then
                    e.CellStyle.ForeColor = UITheme.ColorWarning
                    e.CellStyle.Font = UITheme.FontBold
                Else
                    e.CellStyle.ForeColor = UITheme.ColorSuccess
                    e.CellStyle.Font = UITheme.FontBold
                End If
            ElseIf colName = "Diferencia" Then
                Dim val = dgvCajas.Rows(e.RowIndex).Cells("Diferencia").Value
                If val IsNot Nothing AndAlso Not IsDBNull(val) Then
                    Dim dif As Decimal
                    If Decimal.TryParse(val.ToString(), dif) Then
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
                    Else
                        e.CellStyle.ForeColor = UITheme.ColorTextSecondary
                    End If
                End If
            End If
        End Sub

        Private Function GetSelectedCajaId() As Integer
            If dgvCajas.CurrentRow IsNot Nothing Then
                Dim val = dgvCajas.CurrentRow.Cells("Id").Value
                Dim id As Integer
                If Integer.TryParse(val?.ToString(), id) Then
                    Return id
                End If
            End If
            Return 0
        End Function

        Private Sub DgvCajas_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
            If e.RowIndex >= 0 Then
                AbrirAuditoriaSeleccionada()
            End If
        End Sub

        Private Sub BtnAuditar_Click(sender As Object, e As EventArgs)
            AbrirAuditoriaSeleccionada()
        End Sub

        Private Sub AbrirAuditoriaSeleccionada()
            Dim id = GetSelectedCajaId()
            If id <= 0 Then
                MessageBox.Show("Por favor seleccione un turno de la grilla para auditar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Using frm As New FrmDetalleArqueoCaja(id)
                frm.ShowDialog(Me)
            End Using
        End Sub

        Private Sub BtnVerActa_Click(sender As Object, e As EventArgs)
            Dim id = GetSelectedCajaId()
            If id <= 0 Then
                MessageBox.Show("Por favor seleccione un turno de la grilla para ver su acta de cierre.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Using frm As New FrmDetalleArqueoCaja(id)
                ' Enfocar pestaña de acta si está disponible
                frm.ShowDialog(Me)
            End Using
        End Sub

        Private Sub BtnVerMovimientos_Click(sender As Object, e As EventArgs)
            Dim id = GetSelectedCajaId()
            If id <= 0 Then
                MessageBox.Show("Por favor seleccione un turno para consultar sus retiros e ingresos.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim movs = cajaService.GetMovimientosCaja(id)
            Using dlg As New Form()
                dlg.Text = $"Movimientos de Efectivo - Turno #{id}"
                dlg.Size = New Size(760, 480)
                dlg.StartPosition = FormStartPosition.CenterParent
                dlg.BackColor = UITheme.ColorBackground

                Dim pnlTop As New Panel() With {.Dock = DockStyle.Top, .Height = 50, .BackColor = UITheme.ColorSecondary, .Padding = New Padding(15, 12, 15, 12)}
                Dim lblT As New Label() With {.Text = $"MOVIMIENTOS DE CAJA #{id}", .Font = UITheme.FontSubheading, .ForeColor = Color.White, .AutoSize = True}
                pnlTop.Controls.Add(lblT)

                Dim dgv As New DataGridView() With {.Dock = DockStyle.Fill}
                UITheme.StyleDataGrid(dgv)
                dgv.Columns.Add("Fecha", "Fecha y Hora")
                dgv.Columns("Fecha").Width = 130
                dgv.Columns.Add("Tipo", "Tipo")
                dgv.Columns("Tipo").Width = 100
                dgv.Columns.Add("Concepto", "Concepto")
                dgv.Columns("Concepto").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                dgv.Columns.Add("Monto", "Monto")
                dgv.Columns("Monto").Width = 120
                dgv.Columns("Monto").DefaultCellStyle.Format = "C2"
                dgv.Columns("Monto").DefaultCellStyle.Font = UITheme.FontBold
                dgv.Columns.Add("Usuario", "Registrado Por")
                dgv.Columns("Usuario").Width = 150
                dgv.Columns.Add("Ref", "Comprobante Ref.")
                dgv.Columns("Ref").Width = 120

                For Each m In movs
                    dgv.Rows.Add(m.Fecha.ToString("dd/MM/yyyy HH:mm"), m.Tipo, m.Concepto, m.Monto, m.UsuarioNombre, m.Referencia)
                Next

                Dim pnlBot As New Panel() With {.Dock = DockStyle.Bottom, .Height = 50, .BackColor = UITheme.ColorSurfaceMuted}
                Dim btnC = New Button() With {.Text = "Cerrar", .Location = New Point(630, 9), .Size = New Size(100, 32)}
                UITheme.StyleButton(btnC, "Secondary")
                AddHandler btnC.Click, Sub() dlg.Close()
                pnlBot.Controls.Add(btnC)

                dlg.Controls.AddRange({dgv, pnlBot, pnlTop})
                dlg.ShowDialog(Me)
            End Using
        End Sub

        Private Sub BtnVerTickets_Click(sender As Object, e As EventArgs)
            Dim id = GetSelectedCajaId()
            If id <= 0 Then
                MessageBox.Show("Por favor seleccione un turno para consultar sus tickets.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim ventas = cajaService.GetVentasPorCaja(id)
            Using dlg As New Form()
                dlg.Text = $"Comprobantes de Venta - Turno #{id}"
                dlg.Size = New Size(850, 500)
                dlg.StartPosition = FormStartPosition.CenterParent
                dlg.BackColor = UITheme.ColorBackground

                Dim pnlTop As New Panel() With {.Dock = DockStyle.Top, .Height = 50, .BackColor = UITheme.ColorSecondary, .Padding = New Padding(15, 12, 15, 12)}
                Dim lblT As New Label() With {.Text = $"VENTAS REGISTRADAS EN TURNO #{id} ({ventas.Count} tickets)", .Font = UITheme.FontSubheading, .ForeColor = Color.White, .AutoSize = True}
                pnlTop.Controls.Add(lblT)

                Dim dgv As New DataGridView() With {.Dock = DockStyle.Fill}
                UITheme.StyleDataGrid(dgv)
                dgv.Columns.Add("Ticket", "Nº Ticket")
                dgv.Columns("Ticket").Width = 140
                dgv.Columns.Add("Hora", "Hora")
                dgv.Columns("Hora").Width = 100
                dgv.Columns.Add("Cliente", "Cliente")
                dgv.Columns("Cliente").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                dgv.Columns.Add("Metodo", "Medio de Pago")
                dgv.Columns("Metodo").Width = 130
                dgv.Columns.Add("Total", "Total")
                dgv.Columns("Total").Width = 120
                dgv.Columns("Total").DefaultCellStyle.Format = "C2"
                dgv.Columns("Total").DefaultCellStyle.Font = UITheme.FontBold
                dgv.Columns.Add("Estado", "Estado")
                dgv.Columns("Estado").Width = 110

                For Each v In ventas
                    dgv.Rows.Add(v.NumeroTicket, v.Fecha.ToString("HH:mm:ss"), v.ClienteNombre, v.MetodoPago, v.Total, v.Estado)
                Next

                Dim pnlBot As New Panel() With {.Dock = DockStyle.Bottom, .Height = 50, .BackColor = UITheme.ColorSurfaceMuted}
                Dim btnC = New Button() With {.Text = "Cerrar", .Location = New Point(720, 9), .Size = New Size(100, 32)}
                UITheme.StyleButton(btnC, "Secondary")
                AddHandler btnC.Click, Sub() dlg.Close()
                pnlBot.Controls.Add(btnC)

                dlg.Controls.AddRange({dgv, pnlBot, pnlTop})
                dlg.ShowDialog(Me)
            End Using
        End Sub

        Private Sub BtnExportarCSV_Click(sender As Object, e As EventArgs)
            If listaCajasActual.Count = 0 Then
                MessageBox.Show("No hay registros en la grilla para exportar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Using sfd As New SaveFileDialog()
                sfd.Filter = "Archivo CSV delimitado por comas (*.csv)|*.csv"
                sfd.FileName = $"Auditoria_Cajas_{DateTime.Today:yyyyMMdd}.csv"
                sfd.Title = "Exportar Historial de Arqueos a CSV"

                If sfd.ShowDialog(Me) = DialogResult.OK Then
                    Try
                        Dim sb As New StringBuilder()
                        sb.AppendLine("ID_Turno;Vendedor;Estado;Fecha_Apertura;Fecha_Cierre;Duracion;Fondo_Inicial;Ventas_Efectivo;Ventas_Digital;Ingresos_Extra;Egresos_Retiros;Monto_Esperado;Monto_Real_Declarado;Diferencia_Arqueo;Observaciones")

                        For Each c In listaCajasActual
                            Dim cierre = If(c.FechaCierre.HasValue, c.FechaCierre.Value.ToString("yyyy-MM-dd HH:mm:ss"), "")
                            Dim realStr = If(c.MontoReal.HasValue, c.MontoReal.Value.ToString("F2", System.Globalization.CultureInfo.InvariantCulture), "")
                            Dim difStr = If(c.Diferencia.HasValue, c.Diferencia.Value.ToString("F2", System.Globalization.CultureInfo.InvariantCulture), "")
                            Dim obs = c.Observaciones.Replace(";", " ").Replace(vbCrLf, " ")

                            sb.AppendLine($"{c.Id};{c.UsuarioNombre};{c.Estado};{c.FechaApertura:yyyy-MM-dd HH:mm:ss};{cierre};{c.DuracionTexto};{c.MontoInicial.ToString("F2", System.Globalization.CultureInfo.InvariantCulture)};{c.TotalVentasEfectivo.ToString("F2", System.Globalization.CultureInfo.InvariantCulture)};{c.TotalVentasDigital.ToString("F2", System.Globalization.CultureInfo.InvariantCulture)};{c.TotalIngresos.ToString("F2", System.Globalization.CultureInfo.InvariantCulture)};{c.TotalEgresos.ToString("F2", System.Globalization.CultureInfo.InvariantCulture)};{c.MontoEsperado.ToString("F2", System.Globalization.CultureInfo.InvariantCulture)};{realStr};{difStr};{obs}")
                        Next

                        File.WriteAllText(sfd.FileName, sb.ToString(), Encoding.UTF8)
                        MessageBox.Show("Auditoría exportada exitosamente a CSV.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Catch ex As Exception
                        MessageBox.Show("Error al exportar archivo: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End If
            End Using
        End Sub

    End Class
End Namespace
