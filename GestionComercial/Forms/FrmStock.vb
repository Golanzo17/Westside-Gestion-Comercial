Imports System.Drawing
Imports System.Windows.Forms
Imports GestionComercial.Models
Imports GestionComercial.Services
Imports GestionComercial.UI

' ARCHIVO: FrmStock.vb
' PROPÓSITO: Panel de Control de Inventario, Reposición y Auditoría de Movimientos
' - Estructura de la pantalla:
'   1. 'Prendas con Stock Crítico': Alerta visual inmediata sobre artículos que alcanzaron o quebraron el stock mínimo configurado.
'   2. 'Historial de Movimientos': Trazabilidad completa (kardex) que muestra cada entrada, salida, ajuste, venta o anulación registrada.
' - Seguridad y Privilegios:
'   Un vendedor solo puede visualizar las alertas para consultar disponibilidad.
'   Registrar ingresos de mercadería o ajustes manuales requiere confirmación de credenciales de Administrador o Gerente.


Namespace Forms
    Public Class FrmStock
        Inherits Form

        ' Servicios de negocio reutilizados
        Private catalogService As New CatalogService()
        Private reporteService As New ReporteService()

        ' Controles de interfaz gráfica
        Private dgvAlertas As DataGridView
        Private dgvMovimientos As DataGridView
        Private btnIngresoMercaderia As Button
        Private btnRefrescar As Button
        Private lblResumenAlertas As Label

        Public Sub New()
            InitializeUI()
            LoadStockData()
        End Sub

        ''' <summary>
        ''' Construcción programática del formulario.
        ''' Implementamos un TabControl para organizar la información sin saturar al usuario.
        ''' </summary>
        Private Sub InitializeUI()
            Me.Text = "Control de Stock y Reposición de Mercadería"
            Me.Size = New Size(1020, 640)
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.BackColor = UITheme.ColorBackground
            Me.Font = UITheme.FontRegular

            ' Header institucional superior
            Dim pnlHeader As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 60,
                .BackColor = UITheme.ColorSecondary,
                .Padding = New Padding(20, 15, 20, 15)
            }
            Dim lblTitle As New Label() With {
                .Text = "Control de stock y alertas de reposición",
                .Font = UITheme.FontHeading,
                .ForeColor = Color.White,
                .AutoSize = True,
                .Location = New Point(15, 15)
            }
            pnlHeader.Controls.Add(lblTitle)

            ' Toolbar con FlowLayoutPanel para garantizar que los botones no se recorten en pantallas de baja resolución
            Dim pnlToolbar As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 55,
                .BackColor = UITheme.ColorSurface
            }
            Dim flpToolbar As New FlowLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(10, 10, 10, 10),
                .WrapContents = False
            }

            ' Botón de ingreso/ajuste: Solo visible inicialmente si el usuario actual es Admin/Gerente,
            ' pero si un vendedor lo presiona se invoca el diálogo de autorización.
            btnIngresoMercaderia = New Button() With {.Text = "+ Ingreso de Mercadería / Ajuste", .Height = 34, .AutoSize = True, .AutoSizeMode = AutoSizeMode.GrowAndShrink, .Padding = New Padding(10, 0, 10, 0)}
            UITheme.StyleButton(btnIngresoMercaderia, "Success")
            btnIngresoMercaderia.Visible = AuthService.IsAdminOrManager
            AddHandler btnIngresoMercaderia.Click, AddressOf BtnIngresoMercaderia_Click

            btnRefrescar = New Button() With {.Text = "Actualizar", .Height = 34, .AutoSize = True, .AutoSizeMode = AutoSizeMode.GrowAndShrink, .Padding = New Padding(10, 0, 10, 0), .Margin = New Padding(8, 0, 0, 0)}
            UITheme.StyleButton(btnRefrescar, "Secondary")
            AddHandler btnRefrescar.Click, Sub() LoadStockData()

            lblResumenAlertas = New Label() With {
                .Text = "Verificando niveles de inventario...",
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorWarning,
                .AutoSize = True,
                .Margin = New Padding(20, 8, 0, 0)
            }

            flpToolbar.Controls.AddRange({btnIngresoMercaderia, btnRefrescar, lblResumenAlertas})
            pnlToolbar.Controls.Add(flpToolbar)

            ' TabControl para separar la vista operativa de compras (alertas) de la auditoría histórica
            Dim tabControl As New TabControl() With {
                .Dock = DockStyle.Fill,
                .Padding = New Point(12, 6)
            }

            ' Pestaña 1: Prendas con Stock Crítico o Agotado
            Dim tabAlertas As New TabPage("Prendas con stock crítico / reposición") With {.BackColor = UITheme.ColorBackground}
            dgvAlertas = New DataGridView() With {.Dock = DockStyle.Fill}
            UITheme.StyleDataGrid(dgvAlertas)
            ConfigurarColumnasAlertas()
            tabAlertas.Controls.Add(dgvAlertas)
            tabControl.TabPages.Add(tabAlertas)

            ' Pestaña 2: Historial de Auditoría de Movimientos
            Dim tabMovs As New TabPage("Historial de movimientos de inventario") With {.BackColor = UITheme.ColorBackground}
            dgvMovimientos = New DataGridView() With {.Dock = DockStyle.Fill}
            UITheme.StyleDataGrid(dgvMovimientos)
            ConfigurarColumnasMovimientos()
            tabMovs.Controls.Add(dgvMovimientos)
            tabControl.TabPages.Add(tabMovs)

            ' Orden de Docking de WinForms (Fill primero, paneles acoplados después)
            Me.Controls.Add(tabControl)
            Me.Controls.Add(pnlToolbar)
            Me.Controls.Add(pnlHeader)
        End Sub

        Private Sub ConfigurarColumnasAlertas()
            dgvAlertas.Columns.Clear()
            dgvAlertas.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            ' ── Código Barra (centrado) ──
            dgvAlertas.Columns.Add("Codigo", "Código Barra")
            dgvAlertas.Columns("Codigo").FillWeight = 10
            dgvAlertas.Columns("Codigo").MinimumWidth = 85
            dgvAlertas.Columns("Codigo").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvAlertas.Columns("Codigo").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            ' ── Prenda / Artículo (relleno, izquierda) ──
            dgvAlertas.Columns.Add("Nombre", "Prenda / Artículo")
            dgvAlertas.Columns("Nombre").FillWeight = 30
            dgvAlertas.Columns("Nombre").MinimumWidth = 130
            dgvAlertas.Columns("Nombre").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvAlertas.Columns("Nombre").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft

            ' ── Talle (centrado) ──
            dgvAlertas.Columns.Add("Talle", "Talle")
            dgvAlertas.Columns("Talle").FillWeight = 7
            dgvAlertas.Columns("Talle").MinimumWidth = 60
            dgvAlertas.Columns("Talle").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvAlertas.Columns("Talle").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            ' ── Color (centrado) ──
            dgvAlertas.Columns.Add("Color", "Color")
            dgvAlertas.Columns("Color").FillWeight = 10
            dgvAlertas.Columns("Color").MinimumWidth = 80
            dgvAlertas.Columns("Color").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvAlertas.Columns("Color").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            ' ── Stock Actual (derecha) ──
            dgvAlertas.Columns.Add("StockActual", "Stock Actual")
            dgvAlertas.Columns("StockActual").FillWeight = 9
            dgvAlertas.Columns("StockActual").MinimumWidth = 75
            dgvAlertas.Columns("StockActual").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvAlertas.Columns("StockActual").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight

            ' ── Stock Mínimo (derecha) ──
            dgvAlertas.Columns.Add("StockMinimo", "Stock Mínimo")
            dgvAlertas.Columns("StockMinimo").FillWeight = 9
            dgvAlertas.Columns("StockMinimo").MinimumWidth = 75
            dgvAlertas.Columns("StockMinimo").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvAlertas.Columns("StockMinimo").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight

            ' ── Estado (centrado) ──
            dgvAlertas.Columns.Add("Estado", "Estado")
            dgvAlertas.Columns("Estado").FillWeight = 13
            dgvAlertas.Columns("Estado").MinimumWidth = 100
            dgvAlertas.Columns("Estado").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvAlertas.Columns("Estado").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        End Sub

        Private Sub ConfigurarColumnasMovimientos()
            dgvMovimientos.Columns.Clear()
            dgvMovimientos.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            ' ── Fecha y Hora (centrado) ──
            dgvMovimientos.Columns.Add("Fecha", "Fecha y Hora")
            dgvMovimientos.Columns("Fecha").FillWeight = 12
            dgvMovimientos.Columns("Fecha").MinimumWidth = 110
            dgvMovimientos.Columns("Fecha").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvMovimientos.Columns("Fecha").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            ' ── Tipo Movimiento (centrado) ──
            dgvMovimientos.Columns.Add("Tipo", "Tipo Movimiento")
            dgvMovimientos.Columns("Tipo").FillWeight = 11
            dgvMovimientos.Columns("Tipo").MinimumWidth = 100
            dgvMovimientos.Columns("Tipo").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvMovimientos.Columns("Tipo").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            ' ── Prenda (relleno, izquierda) ──
            dgvMovimientos.Columns.Add("Prenda", "Prenda")
            dgvMovimientos.Columns("Prenda").FillWeight = 22
            dgvMovimientos.Columns("Prenda").MinimumWidth = 120
            dgvMovimientos.Columns("Prenda").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvMovimientos.Columns("Prenda").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft

            ' ── Talle (centrado) ──
            dgvMovimientos.Columns.Add("Talle", "Talle")
            dgvMovimientos.Columns("Talle").FillWeight = 6
            dgvMovimientos.Columns("Talle").MinimumWidth = 55
            dgvMovimientos.Columns("Talle").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvMovimientos.Columns("Talle").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            ' ── Cantidad (derecha) ──
            dgvMovimientos.Columns.Add("Cantidad", "Cantidad")
            dgvMovimientos.Columns("Cantidad").FillWeight = 7
            dgvMovimientos.Columns("Cantidad").MinimumWidth = 65
            dgvMovimientos.Columns("Cantidad").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvMovimientos.Columns("Cantidad").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight

            ' ── Stock Anterior (derecha) ──
            dgvMovimientos.Columns.Add("StockAnt", "Stock Ant.")
            dgvMovimientos.Columns("StockAnt").FillWeight = 7
            dgvMovimientos.Columns("StockAnt").MinimumWidth = 65
            dgvMovimientos.Columns("StockAnt").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvMovimientos.Columns("StockAnt").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight

            ' ── Stock Nuevo (derecha) ──
            dgvMovimientos.Columns.Add("StockPost", "Stock Nuevo")
            dgvMovimientos.Columns("StockPost").FillWeight = 7
            dgvMovimientos.Columns("StockPost").MinimumWidth = 65
            dgvMovimientos.Columns("StockPost").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvMovimientos.Columns("StockPost").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight

            ' ── Motivo / Comprobante (relleno, izquierda) ──
            dgvMovimientos.Columns.Add("Motivo", "Motivo / Comprobante")
            dgvMovimientos.Columns("Motivo").FillWeight = 16
            dgvMovimientos.Columns("Motivo").MinimumWidth = 130
            dgvMovimientos.Columns("Motivo").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvMovimientos.Columns("Motivo").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
        End Sub

        ''' <summary>
        ''' Carga en paralelo los datos de alertas de stock crítico y la auditoría de movimientos.
        ''' Aplicamos formato condicional (rojo/amarillo) según la gravedad de la falta de stock.
        ''' </summary>
        Private Sub LoadStockData()
            ' 1. Consultar artículos bajo stock mínimo
            Dim alertas = reporteService.GetAlertasStockBajo()
            dgvAlertas.Rows.Clear()
            For Each a In alertas
                Dim estadoStr As String = If(a.StockActual = 0, "🔴 AGOTADO", "🟡 REPONER URGENTE")
                Dim rowIdx = dgvAlertas.Rows.Add(a.CodigoBarra, a.Nombre, a.Talle, a.Color, a.StockActual, a.StockMinimo, estadoStr)
                If a.StockActual = 0 Then
                    dgvAlertas.Rows(rowIdx).Cells("StockActual").Style.ForeColor = UITheme.ColorDanger
                Else
                    dgvAlertas.Rows(rowIdx).Cells("StockActual").Style.ForeColor = UITheme.ColorWarning
                End If
            Next

            lblResumenAlertas.Text = $"Se encontraron {alertas.Count} prendas en nivel crítico o sin stock."

            ' 2. Consultar los últimos 100 movimientos de inventario con datos de producto y talle
            Dim dtMovs = Data.DatabaseHelper.ExecuteQuery("SELECT ms.*, p.nombre AS prenda_nombre, t.nombre AS talle_nombre " &
                                                         "FROM `movimientos_stock` ms " &
                                                         "INNER JOIN `productos` p ON ms.producto_id = p.id " &
                                                         "INNER JOIN `talles` t ON ms.talle_id = t.id " &
                                                         "ORDER BY ms.fecha DESC LIMIT 100;")

            dgvMovimientos.Rows.Clear()
            For Each row As System.Data.DataRow In dtMovs.Rows
                dgvMovimientos.Rows.Add(
                    Convert.ToDateTime(row("fecha")).ToString("dd/MM/yyyy HH:mm"),
                    row("tipo_movimiento").ToString(),
                    row("prenda_nombre").ToString(),
                    row("talle_nombre").ToString(),
                    row("cantidad"),
                    row("stock_anterior"),
                    row("stock_posterior"),
                    If(IsDBNull(row("motivo")), "", row("motivo").ToString())
                )
            Next
        End Sub

        ''' <summary>
        ''' Valida permisos de supervisor antes de abrir el diálogo de ajuste de inventario.
        ''' </summary>
        Private Sub BtnIngresoMercaderia_Click(sender As Object, e As EventArgs)
            If Not AuthService.SolicitarAutorizacionAdminOManager(Me, "Los ingresos de mercadería y ajustes de inventario requieren permisos de Administrador o Gerente.") Then
                Return
            End If
            Dim frmAjuste As New FrmAjusteStock()
            If frmAjuste.ShowDialog() = DialogResult.OK Then
                LoadStockData()
            End If
        End Sub

    End Class

    ''' <summary>
    ''' Diálogo modal para registrar ingresos de mercadería (compras a proveedores)
    ''' o ajustes manuales de stock (roturas, devoluciones extraordinarias, mermas).
    ''' Cuenta con búsqueda por teclado y sugerencias en tiempo real, más selección controlada de color.
    ''' </summary>
    Public Class FrmAjusteStock
        Inherits Form

        Private catalogService As New CatalogService()
        Private todosLosProductos As New List(Of Producto)()
        Private productoSeleccionado As Producto = Nothing

        ' Controles de Búsqueda
        Private txtBuscarPrenda As TextBox
        Private lstSugerencias As ListBox
        Private lblPrendaSeleccionada As Label

        ' Variantes y Datos de Stock
        Private cboTalles As ComboBox
        Private cboColor As ComboBox
        Private btnNuevoColor As Button
        Private numCantidad As NumericUpDown
        Private txtMotivo As TextBox
        Private btnGuardar As Button
        Private btnCancelar As Button

        Public Sub New()
            InitializeUI()
            CargarCatalogo()
            CargarColores()
        End Sub

        Private Sub InitializeUI()
            Me.Text = "Ingreso de Mercadería / Reposición de Stock"
            Me.Size = New Size(560, 560)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.BackColor = UITheme.ColorBackground
            Me.Font = UITheme.FontRegular

            ' ─── Fila 1: Búsqueda por teclado en tiempo real ───
            Dim lblP As New Label() With {
                .Text = "Buscar Prenda por Teclado (Nombre o Código):",
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorPrimaryDark,
                .Location = New Point(30, 20),
                .AutoSize = True
            }

            txtBuscarPrenda = New TextBox() With {
                .Location = New Point(30, 46),
                .Size = New Size(480, 28),
                .Font = New Font("Segoe UI", 10.0F, FontStyle.Regular)
            }
            UITheme.StyleTextBox(txtBuscarPrenda)
            AddHandler txtBuscarPrenda.TextChanged, AddressOf TxtBuscarPrenda_TextChanged
            AddHandler txtBuscarPrenda.KeyDown, AddressOf TxtBuscarPrenda_KeyDown

            ' Tarjeta destacada de la prenda elegida
            lblPrendaSeleccionada = New Label() With {
                .Text = "Escriba arriba para buscar y seleccionar una prenda.",
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorTextSecondary,
                .BackColor = UITheme.ColorSurfaceAccent,
                .Location = New Point(30, 80),
                .Size = New Size(480, 42),
                .TextAlign = ContentAlignment.MiddleLeft,
                .Padding = New Padding(10, 0, 10, 0)
            }
            UITheme.ApplyRoundedRegion(lblPrendaSeleccionada, 6)

            ' Lista flotante de sugerencias en tiempo real
            lstSugerencias = New ListBox() With {
                .Location = New Point(30, 76),
                .Size = New Size(480, 150),
                .Font = UITheme.FontRegular,
                .Visible = False,
                .IntegralHeight = False
            }
            AddHandler lstSugerencias.Click, AddressOf LstSugerencias_Click
            AddHandler lstSugerencias.KeyDown, AddressOf LstSugerencias_KeyDown

            ' ─── Fila 2: Talle y Color ───
            Dim lblT As New Label() With {.Text = "Talle disponible:", .Font = UITheme.FontBold, .Location = New Point(30, 135), .AutoSize = True}
            cboTalles = New ComboBox() With {.Location = New Point(30, 160), .Size = New Size(220, 26), .DropDownStyle = ComboBoxStyle.DropDownList}
            UITheme.StyleComboBox(cboTalles)

            Dim lblC As New Label() With {.Text = "Color de la Prenda:", .Font = UITheme.FontBold, .Location = New Point(280, 135), .AutoSize = True}
            cboColor = New ComboBox() With {.Location = New Point(280, 160), .Size = New Size(140, 26), .DropDownStyle = ComboBoxStyle.DropDownList}
            UITheme.StyleComboBox(cboColor)
            AddHandler cboColor.SelectedIndexChanged, AddressOf CboColor_SelectedIndexChanged

            btnNuevoColor = New Button() With {
                .Text = "+ Nuevo",
                .Location = New Point(425, 158),
                .Size = New Size(85, 29),
                .Font = UITheme.FontSmallBold
            }
            UITheme.StyleButton(btnNuevoColor, "Secondary")
            AddHandler btnNuevoColor.Click, Sub() CrearNuevoColor()

            ' ─── Fila 3: Cantidad a Ingresar ───
            Dim lblCant As New Label() With {.Text = "Cantidad de Unidades a Ingresar:", .Font = UITheme.FontBold, .Location = New Point(30, 205), .AutoSize = True}
            numCantidad = New NumericUpDown() With {
                .Location = New Point(30, 230),
                .Size = New Size(220, 26),
                .Minimum = -9999,
                .Maximum = 9999,
                .Value = 1,
                .Font = UITheme.FontBold
            }

            ' ─── Fila 4: Motivo / Justificación ───
            Dim lblM As New Label() With {.Text = "Motivo (Ej: Factura Proveedor, Reposición de Taller):", .Font = UITheme.FontBold, .Location = New Point(30, 275), .AutoSize = True}
            txtMotivo = New TextBox() With {.Location = New Point(30, 300), .Size = New Size(480, 26), .Text = "Ingreso de mercadería / Proveedor"}
            UITheme.StyleTextBox(txtMotivo)

            ' ─── Panel Inferior de Botones ───
            Dim pnlBottom As New Panel() With {.Dock = DockStyle.Bottom, .Height = 55, .BackColor = UITheme.ColorSurfaceMuted, .Padding = New Padding(20, 10, 20, 10)}
            btnGuardar = New Button() With {.Text = "✔ Registrar Ingreso", .Dock = DockStyle.Right, .Width = 180}
            UITheme.StyleButton(btnGuardar, "Success")
            AddHandler btnGuardar.Click, AddressOf BtnGuardar_Click

            btnCancelar = New Button() With {.Text = "Cancelar", .Dock = DockStyle.Left, .Width = 100}
            UITheme.StyleButton(btnCancelar, "Secondary")
            AddHandler btnCancelar.Click, Sub() Me.Close()

            pnlBottom.Controls.AddRange({btnGuardar, btnCancelar})

            Me.Controls.AddRange({
                lblP, txtBuscarPrenda, lblPrendaSeleccionada, lstSugerencias,
                lblT, cboTalles, lblC, cboColor, btnNuevoColor,
                lblCant, numCantidad, lblM, txtMotivo, pnlBottom
            })
            lstSugerencias.BringToFront()
        End Sub

        Private Sub CargarCatalogo()
            todosLosProductos = catalogService.GetProductos("", 0, True)
            If todosLosProductos.Count > 0 Then
                SeleccionarPrenda(todosLosProductos(0))
            End If
        End Sub

        Private Sub CargarColores()
            cboColor.Items.Clear()
            Dim colores = catalogService.GetColoresDisponibles()
            For Each col In colores
                cboColor.Items.Add(col)
            Next
            cboColor.Items.Add("✨ [+ Crear nuevo color...]")

            If cboColor.Items.Count > 0 Then
                cboColor.SelectedIndex = 0
            End If
        End Sub

        Private Sub TxtBuscarPrenda_TextChanged(sender As Object, e As EventArgs)
            Dim term = txtBuscarPrenda.Text.Trim()
            If String.IsNullOrWhiteSpace(term) Then
                lstSugerencias.Visible = False
                Return
            End If

            Dim coincidentes = todosLosProductos.Where(Function(p) _
                p.Nombre.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0 OrElse _
                p.CodigoBarra.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0).ToList()

            lstSugerencias.Items.Clear()
            If coincidentes.Count > 0 Then
                For Each p In coincidentes
                    lstSugerencias.Items.Add(New PrendaItem(p))
                Next
                lstSugerencias.Visible = True
                lstSugerencias.BringToFront()
            Else
                lstSugerencias.Visible = False
            End If
        End Sub

        Private Sub TxtBuscarPrenda_KeyDown(sender As Object, e As KeyEventArgs)
            If e.KeyCode = Keys.Down AndAlso lstSugerencias.Visible Then
                lstSugerencias.Focus()
                If lstSugerencias.Items.Count > 0 Then lstSugerencias.SelectedIndex = 0
                e.Handled = True
            ElseIf e.KeyCode = Keys.Enter Then
                e.SuppressKeyPress = True
                If lstSugerencias.Visible AndAlso lstSugerencias.SelectedIndex >= 0 Then
                    ConfirmarSeleccionSugerencia()
                ElseIf lstSugerencias.Visible AndAlso lstSugerencias.Items.Count > 0 Then
                    lstSugerencias.SelectedIndex = 0
                    ConfirmarSeleccionSugerencia()
                End If
            ElseIf e.KeyCode = Keys.Escape Then
                lstSugerencias.Visible = False
            End If
        End Sub

        Private Sub LstSugerencias_Click(sender As Object, e As EventArgs)
            ConfirmarSeleccionSugerencia()
        End Sub

        Private Sub LstSugerencias_KeyDown(sender As Object, e As KeyEventArgs)
            If e.KeyCode = Keys.Enter Then
                e.SuppressKeyPress = True
                ConfirmarSeleccionSugerencia()
            ElseIf e.KeyCode = Keys.Escape Then
                lstSugerencias.Visible = False
                txtBuscarPrenda.Focus()
            End If
        End Sub

        Private Sub ConfirmarSeleccionSugerencia()
            Dim item = TryCast(lstSugerencias.SelectedItem, PrendaItem)
            If item IsNot Nothing Then
                SeleccionarPrenda(item.Producto)
                lstSugerencias.Visible = False
                txtBuscarPrenda.Text = item.Producto.Nombre
                txtBuscarPrenda.SelectAll()
            End If
        End Sub

        Private Sub SeleccionarPrenda(prod As Producto)
            productoSeleccionado = prod
            lblPrendaSeleccionada.Text = $"✔ Seleccionada: {prod.Nombre} (Cód: {prod.CodigoBarra})"
            lblPrendaSeleccionada.ForeColor = UITheme.ColorPrimaryDark

            ' Cargar talles disponibles
            cboTalles.Items.Clear()
            Dim talles = catalogService.GetTalles()
            For Each t In talles
                cboTalles.Items.Add(t)
            Next
            cboTalles.DisplayMember = "Nombre"
            If cboTalles.Items.Count > 0 Then cboTalles.SelectedIndex = 0
        End Sub

        Private Sub CboColor_SelectedIndexChanged(sender As Object, e As EventArgs)
            If cboColor.SelectedItem IsNot Nothing AndAlso cboColor.SelectedItem.ToString().Contains("[+ Crear nuevo color...]") Then
                CrearNuevoColor()
            End If
        End Sub

        Private Sub CrearNuevoColor()
            Using dlgPrompt As New Form()
                dlgPrompt.Text = "Registrar Nuevo Color de Prenda"
                dlgPrompt.Size = New Size(400, 180)
                dlgPrompt.StartPosition = FormStartPosition.CenterParent
                dlgPrompt.FormBorderStyle = FormBorderStyle.FixedDialog
                dlgPrompt.MaximizeBox = False
                dlgPrompt.MinimizeBox = False
                dlgPrompt.BackColor = UITheme.ColorBackground

                Dim lblPrompt As New Label() With {
                    .Text = "Escriba el nombre del nuevo color (ej: Terracota, Bordo):",
                    .Font = UITheme.FontBold,
                    .Location = New Point(20, 15),
                    .AutoSize = True
                }

                Dim txtNuevoCol As New TextBox() With {
                    .Location = New Point(20, 45),
                    .Size = New Size(345, 26),
                    .Font = UITheme.FontRegular
                }
                UITheme.StyleTextBox(txtNuevoCol)

                Dim btnOk As New Button() With {.Text = "Guardar", .Location = New Point(180, 90), .Size = New Size(95, 32)}
                UITheme.StyleButton(btnOk, "Success")
                Dim btnCancel As New Button() With {.Text = "Cancelar", .Location = New Point(285, 90), .Size = New Size(80, 32)}
                UITheme.StyleButton(btnCancel, "Secondary")

                AddHandler btnOk.Click, Sub()
                                            Dim colName = txtNuevoCol.Text.Trim()
                                            If String.IsNullOrWhiteSpace(colName) Then
                                                MessageBox.Show("Por favor escriba el nombre del color.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                                Return
                                            End If

                                            ' Formatear tipo título (primera letra mayúscula)
                                            colName = Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(colName.ToLower())

                                            ' Verificar si ya existe en el combo
                                            Dim idxExistente = -1
                                            For i = 0 To cboColor.Items.Count - 1
                                                If cboColor.Items(i).ToString().Equals(colName, StringComparison.OrdinalIgnoreCase) Then
                                                    idxExistente = i
                                                    Exit For
                                                End If
                                            Next

                                            If idxExistente >= 0 Then
                                                cboColor.SelectedIndex = idxExistente
                                            Else
                                                ' Insertar antes de la opción "+ Crear nuevo color..."
                                                Dim insertIdx = Math.Max(0, cboColor.Items.Count - 1)
                                                cboColor.Items.Insert(insertIdx, colName)
                                                cboColor.SelectedIndex = insertIdx
                                            End If

                                            dlgPrompt.DialogResult = DialogResult.OK
                                            dlgPrompt.Close()
                                        End Sub

                AddHandler btnCancel.Click, Sub() dlgPrompt.Close()

                dlgPrompt.Controls.AddRange({lblPrompt, txtNuevoCol, btnOk, btnCancel})
                dlgPrompt.ShowDialog(Me)
            End Using

            ' Si canceló o sigue seleccionada la opción de crear, volver a la primera opción
            If cboColor.SelectedItem IsNot Nothing AndAlso cboColor.SelectedItem.ToString().Contains("[+ Crear nuevo color...]") Then
                If cboColor.Items.Count > 1 Then cboColor.SelectedIndex = 0
            End If
        End Sub

        Private Class PrendaItem
            Public Property Producto As Producto

            Public Sub New(p As Producto)
                Me.Producto = p
            End Sub

            Public Overrides Function ToString() As String
                Return $"{Producto.Nombre}  |  Cód: {Producto.CodigoBarra}  |  Cat: {Producto.CategoriaNombre}  |  Stock: {Producto.TotalStock}"
            End Function
        End Class

        Private Sub BtnGuardar_Click(sender As Object, e As EventArgs)
            If Not AuthService.SolicitarAutorizacionAdminOManager(Me, "Guardar movimientos de stock manuales requiere permisos de Administrador o Gerente.") Then
                Return
            End If

            If productoSeleccionado Is Nothing Then
                MessageBox.Show("Por favor busque y seleccione una prenda del catálogo.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtBuscarPrenda.Focus()
                Return
            End If

            Dim talle = TryCast(cboTalles.SelectedItem, Talle)
            If talle Is Nothing Then
                MessageBox.Show("Seleccione el talle correspondiente.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim cant As Integer = Convert.ToInt32(numCantidad.Value)
            If cant = 0 Then
                MessageBox.Show("La cantidad no puede ser 0.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim colorFinal = If(cboColor.SelectedItem IsNot Nothing, cboColor.SelectedItem.ToString().Trim(), "Único")
            If colorFinal.Contains("[+ Crear nuevo color...]") Then colorFinal = "Único"

            Dim motivo As String = txtMotivo.Text.Trim()
            Dim userId As Integer = If(AuthService.CurrentUser IsNot Nothing, AuthService.CurrentUser.Id, 1)

            Dim errMsg As String = ""
            If catalogService.AjustarStockManual(productoSeleccionado.Id, talle.Id, colorFinal, cant, motivo, userId, errMsg) Then
                UITheme.ShowToast(Me, "Stock ingresado y registrado correctamente.", "Success")
                Me.DialogResult = DialogResult.OK
                Me.Close()
            Else
                UITheme.ShowToast(Me, "Error al registrar ajuste: " & errMsg, "Error")
            End If
        End Sub

    End Class
End Namespace
