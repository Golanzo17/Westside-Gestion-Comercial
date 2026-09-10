Imports System.Drawing
Imports System.Windows.Forms
Imports GestionComercial.Models
Imports GestionComercial.Services
Imports GestionComercial.UI

Namespace Forms
    Public Class FrmProductos
        Inherits Form

        Private catalogService As New CatalogService()

        Private txtBuscar As TextBox
        Private cboFiltroCategoria As ComboBox
        Private btnBuscar As Button
        Private btnNuevo As Button
        Private btnEditar As Button
        Private btnEliminar As Button
        Private btnRefrescar As Button
        Private dgvProductos As DataGridView
        Private lblTotalArticulos As Label

        Public Sub New()
            InitializeUI()
            LoadCategorias()
            LoadProductos()
        End Sub

        Private Sub InitializeUI()
            Me.Text = "Catálogo de Prendas y Matriz de Talles"
            Me.Size = New Size(1050, 650)
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
                .Text = "👕 CATÁLOGO DE PRODUCTOS E INDUMENTARIA",
                .Font = UITheme.FontHeading,
                .ForeColor = Color.White,
                .AutoSize = True,
                .Location = New Point(15, 15)
            }
            pnlHeader.Controls.Add(lblTitle)

            ' Barra de Filtros y Acciones
            Dim pnlToolbar As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 95,
                .BackColor = UITheme.ColorSurface,
                .Padding = New Padding(15, 15, 15, 10)
            }

            Dim lblB As New Label() With {.Text = "Buscar:", .Font = UITheme.FontBold, .Location = New Point(15, 20), .AutoSize = True}
            txtBuscar = New TextBox() With {.Location = New Point(70, 18), .Size = New Size(220, 26)}
            UITheme.StyleTextBox(txtBuscar)
            AddHandler txtBuscar.KeyDown, Sub(s, e)
                                              If e.KeyCode = Keys.Enter Then LoadProductos()
                                          End Sub

            Dim lblC As New Label() With {.Text = "Categoría:", .Font = UITheme.FontBold, .Location = New Point(305, 20), .AutoSize = True}
            cboFiltroCategoria = New ComboBox() With {.Location = New Point(380, 18), .Size = New Size(180, 26), .DropDownStyle = ComboBoxStyle.DropDownList}
            AddHandler cboFiltroCategoria.SelectedIndexChanged, Sub() LoadProductos()

            btnBuscar = New Button() With {.Text = "Filtrar", .Location = Point.Add(New Point(570, 16), New Size(0, 0)), .Size = New Size(80, 30)}
            UITheme.StyleButton(btnBuscar, "Primary")
            AddHandler btnBuscar.Click, Sub() LoadProductos()

            btnNuevo = New Button() With {.Text = "+ Nueva Prenda", .Location = New Point(680, 16), .Size = New Size(130, 30)}
            UITheme.StyleButton(btnNuevo, "Success")
            AddHandler btnNuevo.Click, AddressOf BtnNuevo_Click

            btnEditar = New Button() With {.Text = "✏ Editar", .Location = New Point(820, 16), .Size = New Size(95, 30)}
            UITheme.StyleButton(btnEditar, "Secondary")
            AddHandler btnEditar.Click, AddressOf BtnEditar_Click

            btnEliminar = New Button() With {.Text = "Desactivar", .Location = New Point(925, 16), .Size = New Size(95, 30)}
            UITheme.StyleButton(btnEliminar, "Danger")
            AddHandler btnEliminar.Click, AddressOf BtnEliminar_Click

            pnlToolbar.Controls.AddRange({lblB, txtBuscar, lblC, cboFiltroCategoria, btnBuscar, btnNuevo, btnEditar, btnEliminar})

            ' Grilla
            Dim pnlGrid As New Panel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(15)
            }

            dgvProductos = New DataGridView() With {
                .Dock = DockStyle.Fill
            }
            UITheme.StyleDataGrid(dgvProductos)
            ConfigurarColumnas()
            AddHandler dgvProductos.CellDoubleClick, AddressOf DgvProductos_CellDoubleClick

            pnlGrid.Controls.Add(dgvProductos)

            ' Footer con totales
            Dim pnlFooter As New Panel() With {
                .Dock = DockStyle.Bottom,
                .Height = 40,
                .BackColor = Color.FromArgb(241, 245, 249),
                .Padding = New Padding(15, 10, 15, 10)
            }
            lblTotalArticulos = New Label() With {
                .Text = "Cargando artículos...",
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorTextSecondary,
                .AutoSize = True,
                .Location = New Point(15, 10)
            }
            pnlFooter.Controls.Add(lblTotalArticulos)

            ' Orden exacto de Docking: Fill primero, Bottom segundo, Toolbar tercero, Header último
            Me.Controls.Add(pnlGrid)
            Me.Controls.Add(pnlFooter)
            Me.Controls.Add(pnlToolbar)
            Me.Controls.Add(pnlHeader)
        End Sub

        Private Sub ConfigurarColumnas()
            dgvProductos.Columns.Clear()
            dgvProductos.Columns.Add("Id", "ID")
            dgvProductos.Columns("Id").Width = 60

            dgvProductos.Columns.Add("Codigo", "Código Barra")
            dgvProductos.Columns("Codigo").Width = 120

            dgvProductos.Columns.Add("Nombre", "Nombre de Prenda")
            dgvProductos.Columns("Nombre").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

            dgvProductos.Columns.Add("Categoria", "Categoría")
            dgvProductos.Columns("Categoria").Width = 150

            dgvProductos.Columns.Add("PrecioCosto", "Costo")
            dgvProductos.Columns("PrecioCosto").Width = 110
            dgvProductos.Columns("PrecioCosto").DefaultCellStyle.Format = "C2"

            dgvProductos.Columns.Add("PrecioVenta", "Precio Venta")
            dgvProductos.Columns("PrecioVenta").Width = 120
            dgvProductos.Columns("PrecioVenta").DefaultCellStyle.Format = "C2"

            dgvProductos.Columns.Add("Margen", "Margen %")
            dgvProductos.Columns("Margen").Width = 90
            dgvProductos.Columns("Margen").DefaultCellStyle.Format = "0.0'%'"

            dgvProductos.Columns.Add("TotalStock", "Stock Total")
            dgvProductos.Columns("TotalStock").Width = 100
        End Sub

        Private Sub LoadCategorias()
            cboFiltroCategoria.Items.Clear()
            cboFiltroCategoria.Items.Add(New Categoria() With {.Id = 0, .Nombre = "-- Todas las Categorías --"})
            Dim cats = catalogService.GetCategorias(True)
            For Each c In cats
                cboFiltroCategoria.Items.Add(c)
            Next
            cboFiltroCategoria.SelectedIndex = 0
        End Sub

        Private Sub LoadProductos()
            Dim filtro As String = txtBuscar.Text.Trim()
            Dim catId As Integer = 0
            Dim selCat = TryCast(cboFiltroCategoria.SelectedItem, Categoria)
            If selCat IsNot Nothing Then catId = selCat.Id

            Dim lista = catalogService.GetProductos(filtro, catId, True)
            dgvProductos.Rows.Clear()

            For Each p In lista
                dgvProductos.Rows.Add(p.Id, p.CodigoBarra, p.Nombre, p.CategoriaNombre, p.PrecioCosto, p.PrecioVenta, p.PorcentajeGanancia, p.TotalStock)
            Next

            lblTotalArticulos.Text = $"Total de artículos listados: {lista.Count} prendas"
        End Sub

        Private Sub BtnNuevo_Click(sender As Object, e As EventArgs)
            Dim frmEditor As New FrmProductoEditor(0)
            If frmEditor.ShowDialog() = DialogResult.OK Then
                LoadProductos()
            End If
        End Sub

        Private Sub BtnEditar_Click(sender As Object, e As EventArgs)
            EditarSeleccionado()
        End Sub

        Private Sub DgvProductos_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
            If e.RowIndex >= 0 Then
                EditarSeleccionado()
            End If
        End Sub

        Private Sub EditarSeleccionado()
            If dgvProductos.CurrentRow IsNot Nothing Then
                Dim prodId As Integer = Convert.ToInt32(dgvProductos.CurrentRow.Cells("Id").Value)
                Dim frmEditor As New FrmProductoEditor(prodId)
                If frmEditor.ShowDialog() = DialogResult.OK Then
                    LoadProductos()
                End If
            Else
                MessageBox.Show("Por favor seleccione un producto para editar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End Sub

        Private Sub BtnEliminar_Click(sender As Object, e As EventArgs)
            If dgvProductos.CurrentRow IsNot Nothing Then
                Dim prodId As Integer = Convert.ToInt32(dgvProductos.CurrentRow.Cells("Id").Value)
                Dim nombre As String = dgvProductos.CurrentRow.Cells("Nombre").Value.ToString()

                Dim resp = MessageBox.Show($"¿Deseas dar de baja la prenda '{nombre}'?", "Confirmar Baja", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If resp = DialogResult.Yes Then
                    Dim errMsg As String = ""
                    If catalogService.EliminarProducto(prodId, errMsg) Then
                        UITheme.ShowToast(Me, "Prenda desactivada correctamente.", "Success")
                        LoadProductos()
                    Else
                        UITheme.ShowToast(Me, "Error al eliminar: " & errMsg, "Error")
                    End If
                End If
            End If
        End Sub

    End Class

    ''' <summary>
    ''' Formulario Modal de Edición de Producto con Matriz de Talles y Stock
    ''' </summary>
    Public Class FrmProductoEditor
        Inherits Form

        Private _productoId As Integer
        Private catalogService As New CatalogService()
        Private productoActual As Producto

        Private txtCodigo As TextBox
        Private txtNombre As TextBox
        Private txtDescripcion As TextBox
        Private cboCategoria As ComboBox
        Private numCosto As NumericUpDown
        Private numVenta As NumericUpDown
        Private numGanancia As NumericUpDown
        Private dgvTalles As DataGridView
        Private btnGuardar As Button
        Private btnCancelar As Button

        Public Sub New(productoId As Integer)
            _productoId = productoId
            InitializeUI()
            LoadData()
        End Sub

        Private Sub InitializeUI()
            Me.Text = If(_productoId = 0, "Nueva Prenda de Ropa", "Editar Prenda de Ropa")
            Me.Size = New Size(720, 640)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.BackColor = UITheme.ColorBackground
            Me.Font = UITheme.FontRegular

            ' Datos Principales
            Dim grpGeneral As New GroupBox() With {
                .Text = "Información de la Prenda",
                .Location = New Point(20, 15),
                .Size = New Size(665, 230),
                .BackColor = UITheme.ColorSurface,
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorPrimaryDark
            }

            Dim lblCod As New Label() With {.Text = "Código de Barra / SKU:", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextPrimary, .Location = New Point(20, 25), .AutoSize = True}
            txtCodigo = New TextBox() With {.Location = New Point(20, 45), .Size = New Size(200, 26)}
            UITheme.StyleTextBox(txtCodigo)

            Dim lblNom As New Label() With {.Text = "Nombre de la Prenda:", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextPrimary, .Location = New Point(240, 25), .AutoSize = True}
            txtNombre = New TextBox() With {.Location = New Point(240, 45), .Size = New Size(405, 26)}
            UITheme.StyleTextBox(txtNombre)

            Dim lblCat As New Label() With {.Text = "Categoría:", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextPrimary, .Location = New Point(20, 80), .AutoSize = True}
            cboCategoria = New ComboBox() With {.Location = New Point(20, 100), .Size = New Size(200, 26), .DropDownStyle = ComboBoxStyle.DropDownList}

            Dim lblDesc As New Label() With {.Text = "Descripción / Detalles:", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextPrimary, .Location = New Point(240, 80), .AutoSize = True}
            txtDescripcion = New TextBox() With {.Location = New Point(240, 100), .Size = New Size(405, 26)}
            UITheme.StyleTextBox(txtDescripcion)

            ' Precios
            Dim lblCos As New Label() With {.Text = "Precio Costo ($):", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextPrimary, .Location = New Point(20, 140), .AutoSize = True}
            numCosto = New NumericUpDown() With {.Location = New Point(20, 160), .Size = New Size(140, 26), .Maximum = 10000000, .DecimalPlaces = 2}
            AddHandler numCosto.ValueChanged, AddressOf CalcularPrecios

            Dim lblGan As New Label() With {.Text = "Ganancia (%):", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextPrimary, .Location = New Point(180, 140), .AutoSize = True}
            numGanancia = New NumericUpDown() With {.Location = New Point(180, 160), .Size = New Size(110, 26), .Maximum = 1000, .DecimalPlaces = 2, .Value = 50}
            AddHandler numGanancia.ValueChanged, AddressOf CalcularPrecios

            Dim lblVen As New Label() With {.Text = "Precio Venta Final ($):", .Font = UITheme.FontBold, .ForeColor = UITheme.ColorSuccess, .Location = New Point(310, 140), .AutoSize = True}
            numVenta = New NumericUpDown() With {.Location = New Point(310, 160), .Size = New Size(150, 26), .Maximum = 10000000, .DecimalPlaces = 2}

            grpGeneral.Controls.AddRange({lblCod, txtCodigo, lblNom, txtNombre, lblCat, cboCategoria, lblDesc, txtDescripcion, lblCos, numCosto, lblGan, numGanancia, lblVen, numVenta})
            Me.Controls.Add(grpGeneral)

            ' Matriz de Talles
            Dim grpTalles As New GroupBox() With {
                .Text = "Matriz de Stock por Talle y Color",
                .Location = New Point(20, 255),
                .Size = New Size(665, 275),
                .BackColor = UITheme.ColorSurface,
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorPrimaryDark
            }

            dgvTalles = New DataGridView() With {
                .Location = New Point(15, 25),
                .Size = New Size(635, 235),
                .AllowUserToAddRows = False,
                .AllowUserToDeleteRows = False
            }
            UITheme.StyleDataGrid(dgvTalles)
            ConfigurarColumnasTalles()
            grpTalles.Controls.Add(dgvTalles)
            Me.Controls.Add(grpTalles)

            ' Botones guardar / cancelar
            Dim pnlBottom As New Panel() With {
                .Dock = DockStyle.Bottom,
                .Height = 60,
                .BackColor = Color.FromArgb(241, 245, 249),
                .Padding = New Padding(20, 12, 20, 12)
            }

            btnGuardar = New Button() With {.Text = "💾 Guardar Prenda y Stock", .Dock = DockStyle.Right, .Width = 220}
            UITheme.StyleButton(btnGuardar, "Success")
            AddHandler btnGuardar.Click, AddressOf BtnGuardar_Click

            btnCancelar = New Button() With {.Text = "Cancelar", .Dock = DockStyle.Left, .Width = 100}
            UITheme.StyleButton(btnCancelar, "Secondary")
            AddHandler btnCancelar.Click, Sub() Me.Close()

            pnlBottom.Controls.AddRange({btnGuardar, btnCancelar})
            Me.Controls.Add(pnlBottom)
        End Sub

        Private Sub ConfigurarColumnasTalles()
            dgvTalles.Columns.Clear()
            dgvTalles.Columns.Add("TalleId", "ID")
            dgvTalles.Columns("TalleId").Visible = False

            dgvTalles.Columns.Add("Talle", "Talle")
            dgvTalles.Columns("Talle").ReadOnly = True
            dgvTalles.Columns("Talle").Width = 90

            dgvTalles.Columns.Add("Color", "Color / Variante")
            dgvTalles.Columns("Color").Width = 160

            dgvTalles.Columns.Add("Stock", "Stock Actual")
            dgvTalles.Columns("Stock").Width = 120

            dgvTalles.Columns.Add("StockMinimo", "Stock Mínimo")
            dgvTalles.Columns("StockMinimo").Width = 120

            dgvTalles.Columns.Add("Sku", "SKU Específico")
            dgvTalles.Columns("Sku").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        End Sub

        Private Sub LoadData()
            ' Cargar categorías
            cboCategoria.Items.Clear()
            Dim cats = catalogService.GetCategorias(True)
            For Each c In cats
                cboCategoria.Items.Add(c)
            Next
            If cboCategoria.Items.Count > 0 Then cboCategoria.SelectedIndex = 0

            ' Obtener talles disponibles
            Dim listaTalles = catalogService.GetTalles()

            If _productoId = 0 Then
                productoActual = New Producto()
                txtCodigo.Text = "779" & (New Random().Next(100000, 999999)).ToString()

                ' Llenar matriz vacía para los talles
                For Each t In listaTalles
                    dgvTalles.Rows.Add(t.Id, t.Nombre, "Único", 0, 2, $"{txtCodigo.Text}-{t.Nombre}")
                Next
            Else
                productoActual = catalogService.GetProductoPorCodigo(_productoId.ToString())
                If productoActual IsNot Nothing Then
                    txtCodigo.Text = productoActual.CodigoBarra
                    txtNombre.Text = productoActual.Nombre
                    txtDescripcion.Text = productoActual.Descripcion
                    numCosto.Value = productoActual.PrecioCosto
                    numGanancia.Value = productoActual.PorcentajeGanancia
                    numVenta.Value = productoActual.PrecioVenta

                    For i = 0 To cboCategoria.Items.Count - 1
                        Dim c = CType(cboCategoria.Items(i), Categoria)
                        If c.Id = productoActual.CategoriaId Then
                            cboCategoria.SelectedIndex = i
                            Exit For
                        End If
                    Next

                    Dim existentes = productoActual.TallesStock
                    For Each t In listaTalles
                        Dim match = existentes.FirstOrDefault(Function(x) x.TalleId = t.Id)
                        If match IsNot Nothing Then
                            dgvTalles.Rows.Add(t.Id, t.Nombre, match.Color, match.StockActual, match.StockMinimo, match.SkuEspecifico)
                        Else
                            dgvTalles.Rows.Add(t.Id, t.Nombre, "Único", 0, 2, $"{txtCodigo.Text}-{t.Nombre}")
                        End If
                    Next
                End If
            End If
        End Sub

        Private Sub CalcularPrecios(sender As Object, e As EventArgs)
            Dim costo As Decimal = numCosto.Value
            Dim ganancia As Decimal = numGanancia.Value
            numVenta.Value = Math.Round(costo * (1D + (ganancia / 100D)), 2)
        End Sub

        Private Sub BtnGuardar_Click(sender As Object, e As EventArgs)
            If String.IsNullOrWhiteSpace(txtCodigo.Text) Then
                MessageBox.Show("El código de barra o SKU es obligatorio.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If String.IsNullOrWhiteSpace(txtNombre.Text) Then
                MessageBox.Show("El nombre de la prenda es obligatorio.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim catSel = TryCast(cboCategoria.SelectedItem, Categoria)
            If catSel Is Nothing Then
                MessageBox.Show("Seleccione una categoría válida.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            productoActual.CodigoBarra = txtCodigo.Text.Trim()
            productoActual.Nombre = txtNombre.Text.Trim()
            productoActual.Descripcion = txtDescripcion.Text.Trim()
            productoActual.CategoriaId = catSel.Id
            productoActual.PrecioCosto = numCosto.Value
            productoActual.PrecioVenta = numVenta.Value
            productoActual.PorcentajeGanancia = numGanancia.Value
            productoActual.Activo = True

            ' Recoger matriz de talles y stock
            productoActual.TallesStock.Clear()
            For Each row As DataGridViewRow In dgvTalles.Rows
                Dim tId As Integer = Convert.ToInt32(row.Cells("TalleId").Value)
                Dim tNombre As String = row.Cells("Talle").Value.ToString()
                Dim color As String = If(row.Cells("Color").Value IsNot Nothing, row.Cells("Color").Value.ToString(), "Único")
                Dim stock As Integer = 0
                Integer.TryParse(If(row.Cells("Stock").Value IsNot Nothing, row.Cells("Stock").Value.ToString(), "0"), stock)
                Dim stockMin As Integer = 2
                Integer.TryParse(If(row.Cells("StockMinimo").Value IsNot Nothing, row.Cells("StockMinimo").Value.ToString(), "2"), stockMin)
                Dim sku As String = If(row.Cells("Sku").Value IsNot Nothing, row.Cells("Sku").Value.ToString(), "")

                productoActual.TallesStock.Add(New ProductoTalle() With {
                    .TalleId = tId,
                    .TalleNombre = tNombre,
                    .Color = color,
                    .StockActual = Math.Max(0, stock),
                    .StockMinimo = Math.Max(0, stockMin),
                    .SkuEspecifico = sku
                })
            Next

            Dim errMsg As String = ""
            If catalogService.GuardarProducto(productoActual, errMsg) Then
                MessageBox.Show("Prenda y matriz de stock guardadas correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.DialogResult = DialogResult.OK
                Me.Close()
            Else
                MessageBox.Show("Error al guardar: " & errMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End Sub

    End Class
End Namespace
