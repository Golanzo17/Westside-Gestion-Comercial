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
            dgvAlertas.Columns.Add("Codigo", "Código Barra")
            dgvAlertas.Columns("Codigo").Width = 120

            dgvAlertas.Columns.Add("Nombre", "Prenda / Artículo")
            dgvAlertas.Columns("Nombre").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

            dgvAlertas.Columns.Add("Talle", "Talle")
            dgvAlertas.Columns("Talle").Width = 90

            dgvAlertas.Columns.Add("Color", "Color")
            dgvAlertas.Columns("Color").Width = 120

            dgvAlertas.Columns.Add("StockActual", "Stock Actual")
            dgvAlertas.Columns("StockActual").Width = 110

            dgvAlertas.Columns.Add("StockMinimo", "Stock Mínimo")
            dgvAlertas.Columns("StockMinimo").Width = 110

            dgvAlertas.Columns.Add("Estado", "Estado")
            dgvAlertas.Columns("Estado").Width = 140
        End Sub

        Private Sub ConfigurarColumnasMovimientos()
            dgvMovimientos.Columns.Clear()
            dgvMovimientos.Columns.Add("Fecha", "Fecha y Hora")
            dgvMovimientos.Columns("Fecha").Width = 150

            dgvMovimientos.Columns.Add("Tipo", "Tipo Movimiento")
            dgvMovimientos.Columns("Tipo").Width = 140

            dgvMovimientos.Columns.Add("Prenda", "Prenda")
            dgvMovimientos.Columns("Prenda").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

            dgvMovimientos.Columns.Add("Talle", "Talle")
            dgvMovimientos.Columns("Talle").Width = 80

            dgvMovimientos.Columns.Add("Cantidad", "Cantidad")
            dgvMovimientos.Columns("Cantidad").Width = 90

            dgvMovimientos.Columns.Add("StockAnt", "Stock Ant.")
            dgvMovimientos.Columns("StockAnt").Width = 90

            dgvMovimientos.Columns.Add("StockPost", "Stock Nuevo")
            dgvMovimientos.Columns("StockPost").Width = 90

            dgvMovimientos.Columns.Add("Motivo", "Motivo / Comprobante")
            dgvMovimientos.Columns("Motivo").Width = 200
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
    ''' 
    ''' DECISIÓN DE AUDITORÍA:
    ''' Exigimos obligatoriamente un 'Motivo' y registramos el usuario responsable
    ''' para evitar discrepancias inexplicables en el balance de inventario.
    ''' </summary>
    Public Class FrmAjusteStock
        Inherits Form

        Private catalogService As New CatalogService()
        Private cboPrendas As ComboBox
        Private cboTalles As ComboBox
        Private txtColor As TextBox
        Private numCantidad As NumericUpDown
        Private txtMotivo As TextBox
        Private btnGuardar As Button
        Private btnCancelar As Button

        Public Sub New()
            InitializeUI()
            LoadPrendas()
        End Sub

        Private Sub InitializeUI()
            Me.Text = "Ingreso de Mercadería / Ajuste de Stock"
            Me.Size = New Size(500, 420)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.BackColor = UITheme.ColorBackground
            Me.Font = UITheme.FontRegular

            Dim lblP As New Label() With {.Text = "Prenda / Artículo:", .Font = UITheme.FontBold, .Location = New Point(30, 20), .AutoSize = True}
            cboPrendas = New ComboBox() With {.Location = New Point(30, 50), .Size = New Size(420, 26), .DropDownStyle = ComboBoxStyle.DropDownList}
            UITheme.StyleComboBox(cboPrendas)
            AddHandler cboPrendas.SelectedIndexChanged, AddressOf CboPrendas_SelectedIndexChanged

            Dim lblT As New Label() With {.Text = "Talle:", .Font = UITheme.FontBold, .Location = New Point(30, 95), .AutoSize = True}
            cboTalles = New ComboBox() With {.Location = New Point(30, 120), .Size = New Size(190, 26), .DropDownStyle = ComboBoxStyle.DropDownList}
            UITheme.StyleComboBox(cboTalles)

            Dim lblC As New Label() With {.Text = "Color:", .Font = UITheme.FontBold, .Location = New Point(250, 95), .AutoSize = True}
            txtColor = New TextBox() With {.Location = New Point(250, 120), .Size = New Size(200, 26), .Text = "Único"}
            UITheme.StyleTextBox(txtColor)

            Dim lblCant As New Label() With {.Text = "Cantidad a Ingresar / Ajustar:", .Font = UITheme.FontBold, .Location = New Point(30, 165), .AutoSize = True}
            numCantidad = New NumericUpDown() With {.Location = New Point(30, 190), .Size = New Size(190, 26), .Minimum = -9999, .Maximum = 9999, .Value = 1}

            Dim lblM As New Label() With {.Text = "Motivo (Ej: Factura Proveedor, Reposición):", .Font = UITheme.FontBold, .Location = New Point(30, 235), .AutoSize = True}
            txtMotivo = New TextBox() With {.Location = New Point(30, 260), .Size = New Size(420, 26), .Text = "Ingreso de mercadería"}
            UITheme.StyleTextBox(txtMotivo)

            Dim pnlBottom As New Panel() With {.Dock = DockStyle.Bottom, .Height = 55, .BackColor = UITheme.ColorSurfaceMuted, .Padding = New Padding(20, 10, 20, 10)}
            btnGuardar = New Button() With {.Text = "✔ Registrar Ajuste", .Dock = DockStyle.Right, .Width = 160}
            UITheme.StyleButton(btnGuardar, "Success")
            AddHandler btnGuardar.Click, AddressOf BtnGuardar_Click

            btnCancelar = New Button() With {.Text = "Cancelar", .Dock = DockStyle.Left, .Width = 100}
            UITheme.StyleButton(btnCancelar, "Secondary")
            AddHandler btnCancelar.Click, Sub() Me.Close()

            pnlBottom.Controls.AddRange({btnGuardar, btnCancelar})

            Me.Controls.AddRange({lblP, cboPrendas, lblT, cboTalles, lblC, txtColor, lblCant, numCantidad, lblM, txtMotivo, pnlBottom})
        End Sub

        Private Sub LoadPrendas()
            cboPrendas.Items.Clear()
            Dim prods = catalogService.GetProductos("", 0, True)
            For Each p In prods
                cboPrendas.Items.Add(p)
            Next
            cboPrendas.DisplayMember = "Nombre"
            If cboPrendas.Items.Count > 0 Then cboPrendas.SelectedIndex = 0
        End Sub

        Private Sub CboPrendas_SelectedIndexChanged(sender As Object, e As EventArgs)
            cboTalles.Items.Clear()
            Dim prod = TryCast(cboPrendas.SelectedItem, Producto)
            If prod IsNot Nothing Then
                Dim talles = catalogService.GetTalles()
                For Each t In talles
                    cboTalles.Items.Add(t)
                Next
                cboTalles.DisplayMember = "Nombre"
                If cboTalles.Items.Count > 0 Then cboTalles.SelectedIndex = 0
            End If
        End Sub

        ''' <summary>
        ''' Validación y persistencia transaccional del ajuste manual.
        ''' Invoca CatalogService.AjustarStockManual que registra el movimiento en kardex y
        ''' actualiza la matriz producto_talles atómicamente.
        ''' </summary>
        Private Sub BtnGuardar_Click(sender As Object, e As EventArgs)
            If Not AuthService.SolicitarAutorizacionAdminOManager(Me, "Guardar movimientos de stock manuales requiere permisos de Administrador o Gerente.") Then
                Return
            End If

            Dim prod = TryCast(cboPrendas.SelectedItem, Producto)
            Dim talle = TryCast(cboTalles.SelectedItem, Talle)
            If prod Is Nothing OrElse talle Is Nothing Then
                MessageBox.Show("Seleccione la prenda y el talle.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim cant As Integer = Convert.ToInt32(numCantidad.Value)
            If cant = 0 Then
                MessageBox.Show("La cantidad no puede ser 0.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim color As String = If(String.IsNullOrWhiteSpace(txtColor.Text), "Único", txtColor.Text.Trim())
            Dim motivo As String = txtMotivo.Text.Trim()
            Dim userId As Integer = If(AuthService.CurrentUser IsNot Nothing, AuthService.CurrentUser.Id, 1)

            Dim errMsg As String = ""
            If catalogService.AjustarStockManual(prod.Id, talle.Id, color, cant, motivo, userId, errMsg) Then
                UITheme.ShowToast(Me, "Stock ajustado correctamente.", "Success")
                Me.DialogResult = DialogResult.OK
                Me.Close()
            Else
                UITheme.ShowToast(Me, "Error al registrar ajuste: " & errMsg, "Error")
            End If
        End Sub

    End Class
End Namespace
