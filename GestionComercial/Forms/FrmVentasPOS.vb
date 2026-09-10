Imports System.Drawing
Imports System.Windows.Forms
Imports GestionComercial.Models
Imports GestionComercial.Services
Imports GestionComercial.UI

Namespace Forms
    Public Class FrmVentasPOS
        Inherits Form

        ' Servicios
        Private catalogService As New CatalogService()
        Private ventaService As New VentaService()
        Private clienteService As New ClienteService()
        Private cajaService As New CajaService()
        Private configService As New ConfiguracionService()

        ' Controles de Búsqueda y Selección de Prenda
        Private txtBuscarArticulo As TextBox
        Private btnBuscarArticulo As Button
        Private lblProductoSeleccionado As Label
        Private lblPrecioUnitario As Label
        Private pnlTallesContainer As FlowLayoutPanel
        Private cboColor As ComboBox
        Private numCantidad As NumericUpDown
        Private btnAgregarAlCarrito As Button

        ' Carrito de compras
        Private dgvCarrito As DataGridView
        Private btnQuitarItem As Button
        Private btnVaciarCarrito As Button

        ' Cliente
        Private cboClientes As ComboBox
        Private btnNuevoCliente As Button

        ' Métodos de Pago y Totales
        Private cboMetodoPago As ComboBox
        Private txtSubtotal As Label
        Private numDescuentoPorc As NumericUpDown
        Private lblTotalPagar As Label
        Private txtMontoAbonado As TextBox
        Private lblVuelto As Label
        Private btnFinalizarVenta As Button

        ' Estado actual
        Private productoActual As Producto = Nothing
        Private talleSeleccionado As ProductoTalle = Nothing
        Private listaClientes As List(Of Cliente) = New List(Of Cliente)()
        Private ventaActual As Venta = New Venta()

        Public Sub New()
            InitializeUI()
            LoadInitialData()
        End Sub

        Private Sub InitializeUI()
            Me.Text = "Punto de Venta (POS) - Local de Ropa"
            Me.Size = New Size(1100, 720)
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.BackColor = UITheme.ColorBackground
            Me.Font = UITheme.FontRegular

            ' Header POS
            Dim pnlHeader As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 55,
                .BackColor = UITheme.ColorSecondary,
                .Padding = New Padding(15, 10, 15, 10)
            }
            Dim lblTitle As New Label() With {
                .Text = "🛒 PUNTO DE VENTA Y FACTURACIÓN",
                .Font = UITheme.FontHeading,
                .ForeColor = Color.White,
                .AutoSize = True,
                .Location = New Point(15, 12)
            }
            pnlHeader.Controls.Add(lblTitle)

            ' Panel Principal Dividido: Izquierda (Catálogo y Carrito), Derecha (Cobro y Totales)
            Dim pnlMainContainer As New Panel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(15)
            }

            ' ==================== COLUMNA DERECHA: TOTALES Y COBRO ====================
            Dim pnlCobro As New Panel() With {
                .Dock = DockStyle.Right,
                .Width = 360,
                .BackColor = UITheme.ColorSurface,
                .Padding = New Padding(15)
            }

            Dim lblCobroHeader As New Label() With {
                .Text = "DATOS DE COBRO",
                .Font = UITheme.FontSubheading,
                .ForeColor = UITheme.ColorPrimaryDark,
                .Location = New Point(15, 15),
                .AutoSize = True
            }

            ' Selección de Cliente
            Dim lblCli As New Label() With {.Text = "Cliente:", .Font = UITheme.FontBold, .Location = New Point(15, 50), .AutoSize = True}
            cboClientes = New ComboBox() With {.Location = New Point(15, 70), .Size = New Size(260, 26), .DropDownStyle = ComboBoxStyle.DropDownList}
            btnNuevoCliente = New Button() With {.Text = "+", .Location = New Point(280, 69), .Size = New Size(35, 27)}
            UITheme.StyleButton(btnNuevoCliente, "Secondary")
            AddHandler btnNuevoCliente.Click, AddressOf BtnNuevoCliente_Click

            ' Medio de Pago
            Dim lblMetodo As New Label() With {.Text = "Medio de Pago:", .Font = UITheme.FontBold, .Location = New Point(15, 110), .AutoSize = True}
            cboMetodoPago = New ComboBox() With {.Location = New Point(15, 130), .Size = New Size(300, 26), .DropDownStyle = ComboBoxStyle.DropDownList}
            cboMetodoPago.Items.AddRange({"Efectivo", "Tarjeta Débito", "Tarjeta Crédito", "Transferencia / QR", "Múltiple"})
            cboMetodoPago.SelectedIndex = 0
            AddHandler cboMetodoPago.SelectedIndexChanged, AddressOf RecalcularTotales

            ' Descuento %
            Dim lblDesc As New Label() With {.Text = "Descuento (%):", .Font = UITheme.FontRegular, .Location = New Point(15, 170), .AutoSize = True}
            numDescuentoPorc = New NumericUpDown() With {.Location = New Point(15, 190), .Size = New Size(120, 26), .Minimum = 0, .Maximum = 100, .Value = 0}
            AddHandler numDescuentoPorc.ValueChanged, AddressOf RecalcularTotales

            ' Subtotal
            Dim lblSubt As New Label() With {.Text = "Subtotal:", .Font = UITheme.FontRegular, .Location = New Point(170, 170), .AutoSize = True}
            txtSubtotal = New Label() With {.Text = "$ 0,00", .Font = UITheme.FontBold, .Location = New Point(170, 192), .AutoSize = True}

            ' Tarjeta Destacada TOTAL
            Dim pnlTotalCard As New Panel() With {
                .Location = New Point(15, 235),
                .Size = New Size(300, 90),
                .BackColor = Color.FromArgb(238, 242, 255),
                .Padding = New Padding(12)
            }
            Dim lblTotalTitle As New Label() With {.Text = "TOTAL A PAGAR", .Font = UITheme.FontSmall, .ForeColor = UITheme.ColorPrimaryDark, .Location = New Point(10, 8), .AutoSize = True}
            lblTotalPagar = New Label() With {.Text = "$ 0,00", .Font = UITheme.FontPriceBig, .ForeColor = UITheme.ColorPrimary, .Location = New Point(10, 30), .AutoSize = True}
            pnlTotalCard.Controls.AddRange({lblTotalTitle, lblTotalPagar})

            ' Efectivo y Vuelto
            Dim lblAbonado As New Label() With {.Text = "Monto Abonado ($):", .Font = UITheme.FontBold, .Location = New Point(15, 340), .AutoSize = True}
            txtMontoAbonado = New TextBox() With {.Location = New Point(15, 360), .Size = New Size(140, 28), .Font = UITheme.FontBold}
            UITheme.StyleTextBox(txtMontoAbonado)
            AddHandler txtMontoAbonado.TextChanged, AddressOf TxtMontoAbonado_TextChanged

            Dim lblVueltoTitle As New Label() With {.Text = "Vuelto:", .Font = UITheme.FontBold, .Location = New Point(170, 340), .AutoSize = True}
            lblVuelto = New Label() With {.Text = "$ 0,00", .Font = New Font("Segoe UI", 13.0F, FontStyle.Bold), .ForeColor = UITheme.ColorSuccess, .Location = New Point(170, 360), .AutoSize = True}

            ' Botón Finalizar Venta
            btnFinalizarVenta = New Button() With {
                .Text = "✔ COBRAR / REGISTRAR VENTA (F5)",
                .Location = New Point(15, 415),
                .Size = New Size(300, 50),
                .Font = New Font("Segoe UI", 11.0F, FontStyle.Bold)
            }
            UITheme.StyleButton(btnFinalizarVenta, "Success")
            AddHandler btnFinalizarVenta.Click, AddressOf BtnFinalizarVenta_Click

            pnlCobro.Controls.AddRange({
                lblCobroHeader, lblCli, cboClientes, btnNuevoCliente,
                lblMetodo, cboMetodoPago, lblDesc, numDescuentoPorc,
                lblSubt, txtSubtotal, pnlTotalCard, lblAbonado, txtMontoAbonado,
                lblVueltoTitle, lblVuelto, btnFinalizarVenta
            })

            ' ==================== COLUMNA IZQUIERDA: BÚSQUEDA Y CARRITO ====================
            Dim pnlIzquierda As New Panel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(0, 0, 15, 0)
            }

            ' Panel de búsqueda y selección de prenda
            Dim grpArticulo As New GroupBox() With {
                .Text = "Búsqueda de Prendas y Talles",
                .Dock = DockStyle.Top,
                .Height = 255,
                .BackColor = UITheme.ColorSurface,
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorPrimaryDark,
                .Padding = New Padding(12)
            }

            ' Fila 1: Caja de búsqueda
            Dim lblBuscar As New Label() With {.Text = "Escanear Código de Barras o Nombre de Prenda:", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextPrimary, .Location = New Point(15, 25), .AutoSize = True}
            txtBuscarArticulo = New TextBox() With {.Location = New Point(15, 52), .Size = New Size(380, 28)}
            UITheme.StyleTextBox(txtBuscarArticulo)
            AddHandler txtBuscarArticulo.KeyDown, AddressOf TxtBuscarArticulo_KeyDown

            btnBuscarArticulo = New Button() With {.Text = "Buscar", .Location = New Point(405, 51), .Size = New Size(95, 30)}
            UITheme.StyleButton(btnBuscarArticulo, "Primary")
            AddHandler btnBuscarArticulo.Click, AddressOf BtnBuscarArticulo_Click

            ' Fila 2: Resumen del artículo encontrado y precio
            lblProductoSeleccionado = New Label() With {
                .Text = "Seleccione una prenda para ver talles y stock disponible.",
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorTextPrimary,
                .Location = New Point(15, 92),
                .Size = New Size(400, 30),
                .AutoEllipsis = True
            }
            lblPrecioUnitario = New Label() With {
                .Text = "",
                .Font = New Font("Segoe UI", 13.0F, FontStyle.Bold),
                .ForeColor = UITheme.ColorSuccess,
                .Location = New Point(425, 90),
                .AutoSize = True
            }

            ' Fila 3: Selector dinámico de talles (Botones interactivos)
            Dim lblTalles As New Label() With {.Text = "Talles y Stock Disponible:", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextSecondary, .Location = New Point(15, 125), .AutoSize = True}
            pnlTallesContainer = New FlowLayoutPanel() With {
                .Location = New Point(15, 148),
                .Size = New Size(560, 42),
                .AutoScroll = True,
                .WrapContents = False,
                .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            }

            ' Fila 4: Color, Cantidad y Botón Agregar al Carrito (bien espaciado, sin cortes)
            Dim lblCol As New Label() With {.Text = "Color:", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextSecondary, .Location = New Point(15, 203), .AutoSize = True}
            cboColor = New ComboBox() With {.Location = New Point(60, 200), .Size = New Size(110, 26), .DropDownStyle = ComboBoxStyle.DropDownList}

            Dim lblCant As New Label() With {.Text = "Cant:", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextSecondary, .Location = New Point(185, 203), .AutoSize = True}
            numCantidad = New NumericUpDown() With {.Location = New Point(225, 200), .Size = New Size(65, 26), .Minimum = 1, .Maximum = 999, .Value = 1}

            ' Botón agregar al carrito: amplio, destacado y con margen amplio respecto al borde derecho
            btnAgregarAlCarrito = New Button() With {
                .Text = "+ AGREGAR AL CARRITO",
                .Location = New Point(310, 196),
                .Size = New Size(210, 36),
                .Font = New Font("Segoe UI", 10.0F, FontStyle.Bold)
            }
            UITheme.StyleButton(btnAgregarAlCarrito, "Success")
            btnAgregarAlCarrito.Enabled = False
            AddHandler btnAgregarAlCarrito.Click, AddressOf BtnAgregarAlCarrito_Click

            grpArticulo.Controls.AddRange({
                lblBuscar, txtBuscarArticulo, btnBuscarArticulo,
                lblProductoSeleccionado, lblPrecioUnitario,
                lblTalles, pnlTallesContainer, lblCol, cboColor, lblCant, numCantidad, btnAgregarAlCarrito
            })

            ' Grilla del Carrito de Ventas
            Dim pnlGridContainer As New Panel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(0, 10, 0, 0)
            }

            dgvCarrito = New DataGridView() With {
                .Dock = DockStyle.Fill
            }
            UITheme.StyleDataGrid(dgvCarrito)
            ConfigurarColumnasCarrito()
            pnlGridContainer.Controls.Add(dgvCarrito)

            ' Barra de acciones del carrito
            Dim pnlAccionesCarrito As New Panel() With {
                .Dock = DockStyle.Bottom,
                .Height = 45,
                .Padding = New Padding(0, 8, 0, 0)
            }
            btnQuitarItem = New Button() With {.Text = "✖ Quitar Ítem Seleccionado", .Dock = DockStyle.Left, .Width = 200}
            UITheme.StyleButton(btnQuitarItem, "Danger")
            AddHandler btnQuitarItem.Click, AddressOf BtnQuitarItem_Click

            btnVaciarCarrito = New Button() With {.Text = "Vaciar Carrito", .Dock = DockStyle.Right, .Width = 130}
            UITheme.StyleButton(btnVaciarCarrito, "Secondary")
            AddHandler btnVaciarCarrito.Click, AddressOf BtnVaciarCarrito_Click

            pnlAccionesCarrito.Controls.AddRange({btnQuitarItem, btnVaciarCarrito})
            pnlGridContainer.Controls.Add(pnlAccionesCarrito)

            ' Orden en pnlIzquierda: Fill primero, luego Top
            pnlIzquierda.Controls.Add(pnlGridContainer)
            pnlIzquierda.Controls.Add(grpArticulo)

            ' Orden en pnlMainContainer: Fill (pnlIzquierda) primero, luego Right (pnlCobro)
            pnlMainContainer.Controls.Add(pnlIzquierda)
            pnlMainContainer.Controls.Add(pnlCobro)

            ' Orden en Form: Fill (pnlMainContainer) primero, luego Top (pnlHeader)
            Me.Controls.Add(pnlMainContainer)
            Me.Controls.Add(pnlHeader)

            ' Atajos
            Me.KeyPreview = True
            AddHandler Me.KeyDown, AddressOf FrmVentasPOS_KeyDown
        End Sub

        Private Sub ConfigurarColumnasCarrito()
            dgvCarrito.Columns.Clear()
            dgvCarrito.Columns.Add("ProductoId", "ID")
            dgvCarrito.Columns("ProductoId").Visible = False

            dgvCarrito.Columns.Add("Codigo", "Código")
            dgvCarrito.Columns("Codigo").Width = 100

            dgvCarrito.Columns.Add("Descripcion", "Prenda / Artículo")
            dgvCarrito.Columns("Descripcion").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

            dgvCarrito.Columns.Add("Talle", "Talle")
            dgvCarrito.Columns("Talle").Width = 80

            dgvCarrito.Columns.Add("Color", "Color")
            dgvCarrito.Columns("Color").Width = 90

            dgvCarrito.Columns.Add("Cantidad", "Cant.")
            dgvCarrito.Columns("Cantidad").Width = 70

            dgvCarrito.Columns.Add("PrecioUnitario", "Precio Unit.")
            dgvCarrito.Columns("PrecioUnitario").Width = 140
            dgvCarrito.Columns("PrecioUnitario").DefaultCellStyle.Format = "C2"

            dgvCarrito.Columns.Add("Subtotal", "Subtotal")
            dgvCarrito.Columns("Subtotal").Width = 140
            dgvCarrito.Columns("Subtotal").DefaultCellStyle.Format = "C2"
        End Sub

        Private Sub LoadInitialData()
            CargarClientes()
        End Sub

        Private Sub CargarClientes()
            cboClientes.Items.Clear()
            listaClientes = clienteService.GetClientes("", True)
            For Each c In listaClientes
                cboClientes.Items.Add(c)
            Next
            If cboClientes.Items.Count > 0 Then cboClientes.SelectedIndex = 0
        End Sub

        Private Sub TxtBuscarArticulo_KeyDown(sender As Object, e As KeyEventArgs)
            If e.KeyCode = Keys.Enter Then
                e.SuppressKeyPress = True
                BuscarArticulo()
            End If
        End Sub

        Private Sub BtnBuscarArticulo_Click(sender As Object, e As EventArgs)
            BuscarArticulo()
        End Sub

        Private Sub BuscarArticulo()
            Dim term As String = txtBuscarArticulo.Text.Trim()
            If String.IsNullOrWhiteSpace(term) Then Return

            ' Búsqueda exacta por código de barra o parcial por nombre
            Dim prod = catalogService.GetProductoPorCodigo(term)
            If prod Is Nothing Then
                Dim lista = catalogService.GetProductos(term, 0, True)
                If lista.Count > 0 Then
                    prod = catalogService.GetProductoPorCodigo(lista(0).CodigoBarra)
                End If
            End If

            If prod IsNot Nothing Then
                MostrarPrendaSeleccionada(prod)
            Else
                MessageBox.Show("No se encontró ningún artículo con el código o nombre ingresado.", "Artículo No Encontrado", MessageBoxButtons.OK, MessageBoxIcon.Information)
                txtBuscarArticulo.SelectAll()
            End If
        End Sub

        Private Sub MostrarPrendaSeleccionada(prod As Producto)
            productoActual = prod
            talleSeleccionado = Nothing
            lblProductoSeleccionado.Text = $"{prod.Nombre} (Cód: {prod.CodigoBarra})"
            lblPrecioUnitario.Text = prod.PrecioVenta.ToString("C2")

            ' Renderizar botones de talles interactivos
            pnlTallesContainer.Controls.Clear()
            cboColor.Items.Clear()

            If prod.TallesStock IsNot Nothing AndAlso prod.TallesStock.Count > 0 Then
                Dim colores As New HashSet(Of String)()

                For Each pt In prod.TallesStock
                    colores.Add(pt.Color)

                    Dim btnTalle As New Button() With {
                        .Text = $"{pt.TalleNombre} ({pt.StockActual})",
                        .Tag = pt,
                        .Height = 34,
                        .Width = 85,
                        .Margin = New Padding(3)
                    }

                    If pt.StockActual > 0 Then
                        UITheme.StyleButton(btnTalle, "Secondary")
                    Else
                        btnTalle.FlatStyle = FlatStyle.Flat
                        btnTalle.BackColor = Color.FromArgb(241, 245, 249)
                        btnTalle.ForeColor = Color.FromArgb(148, 163, 184)
                        btnTalle.Enabled = False
                    End If

                    AddHandler btnTalle.Click, Sub(s, ev)
                                                   SeleccionarTalle(CType(CType(s, Button).Tag, ProductoTalle))
                                               End Sub

                    pnlTallesContainer.Controls.Add(btnTalle)
                Next

                For Each col In colores
                    cboColor.Items.Add(col)
                Next
                If cboColor.Items.Count > 0 Then cboColor.SelectedIndex = 0

                ' Auto-seleccionar primer talle disponible con stock
                Dim primerConStock = prod.TallesStock.FirstOrDefault(Function(x) x.StockActual > 0)
                If primerConStock IsNot Nothing Then
                    SeleccionarTalle(primerConStock)
                End If
            Else
                Dim lblSinTalles As New Label() With {.Text = "Sin matriz de stock asignada", .ForeColor = UITheme.ColorDanger, .AutoSize = True}
                pnlTallesContainer.Controls.Add(lblSinTalles)
                btnAgregarAlCarrito.Enabled = False
            End If
        End Sub

        Private Sub SeleccionarTalle(pt As ProductoTalle)
            talleSeleccionado = pt
            btnAgregarAlCarrito.Enabled = (pt IsNot Nothing AndAlso pt.StockActual > 0)

            ' Destacar botón seleccionado
            For Each ctrl As Control In pnlTallesContainer.Controls
                If TypeOf ctrl Is Button Then
                    Dim b = CType(ctrl, Button)
                    Dim tagPt = CType(b.Tag, ProductoTalle)
                    If tagPt IsNot Nothing AndAlso tagPt.Id = pt.Id Then
                        b.BackColor = UITheme.ColorPrimary
                        b.ForeColor = Color.White
                    ElseIf tagPt IsNot Nothing AndAlso tagPt.StockActual > 0 Then
                        b.BackColor = Color.FromArgb(241, 245, 249)
                        b.ForeColor = UITheme.ColorTextPrimary
                    End If
                End If
            Next

            Dim cIdx = cboColor.FindStringExact(pt.Color)
            If cIdx >= 0 Then cboColor.SelectedIndex = cIdx
            numCantidad.Maximum = Math.Max(1, pt.StockActual)
            numCantidad.Value = 1
        End Sub

        Private Sub BtnAgregarAlCarrito_Click(sender As Object, e As EventArgs)
            If productoActual Is Nothing OrElse talleSeleccionado Is Nothing Then
                MessageBox.Show("Por favor seleccione un producto y un talle disponible.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim cant As Integer = Convert.ToInt32(numCantidad.Value)
            Dim colorSel As String = If(cboColor.SelectedItem IsNot Nothing, cboColor.SelectedItem.ToString(), talleSeleccionado.Color)

            ' Verificar si ya está en el carrito
            Dim itemExistente = ventaActual.Detalles.FirstOrDefault(Function(d) d.ProductoId = productoActual.Id AndAlso d.TalleId = talleSeleccionado.TalleId AndAlso d.Color = colorSel)
            If itemExistente IsNot Nothing Then
                If itemExistente.Cantidad + cant > talleSeleccionado.StockActual Then
                    MessageBox.Show($"No puedes agregar más unidades. El stock máximo para este talle es {talleSeleccionado.StockActual}.", "Stock Máximo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If
                itemExistente.Cantidad += cant
                itemExistente.Subtotal = itemExistente.Cantidad * itemExistente.PrecioUnitario
            Else
                Dim subtotal As Decimal = cant * productoActual.PrecioVenta
                Dim detalle As New DetalleVenta() With {
                    .ProductoId = productoActual.Id,
                    .TalleId = talleSeleccionado.TalleId,
                    .TalleNombre = talleSeleccionado.TalleNombre,
                    .Color = colorSel,
                    .CodigoBarra = productoActual.CodigoBarra,
                    .DescripcionArticulo = productoActual.Nombre,
                    .PrecioUnitario = productoActual.PrecioVenta,
                    .CostoUnitario = productoActual.PrecioCosto,
                    .Cantidad = cant,
                    .Subtotal = subtotal
                }
                ventaActual.Detalles.Add(detalle)
            End If

            RefrescarGrillaCarrito()
            RecalcularTotales()

            ' Limpiar búsqueda y hacer foco para siguiente código de barra
            txtBuscarArticulo.Clear()
            txtBuscarArticulo.Focus()
        End Sub

        Private Sub RefrescarGrillaCarrito()
            dgvCarrito.Rows.Clear()
            For Each item In ventaActual.Detalles
                dgvCarrito.Rows.Add(item.ProductoId, item.CodigoBarra, item.DescripcionArticulo, item.TalleNombre, item.Color, item.Cantidad, item.PrecioUnitario, item.Subtotal)
            Next
        End Sub

        Private Sub BtnQuitarItem_Click(sender As Object, e As EventArgs)
            If dgvCarrito.CurrentRow IsNot Nothing Then
                Dim idx = dgvCarrito.CurrentRow.Index
                If idx >= 0 AndAlso idx < ventaActual.Detalles.Count Then
                    ventaActual.Detalles.RemoveAt(idx)
                    RefrescarGrillaCarrito()
                    RecalcularTotales()
                End If
            End If
        End Sub

        Private Sub BtnVaciarCarrito_Click(sender As Object, e As EventArgs)
            If ventaActual.Detalles.Count > 0 Then
                If MessageBox.Show("¿Deseas vaciar todos los artículos del carrito?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    ventaActual.Detalles.Clear()
                    RefrescarGrillaCarrito()
                    RecalcularTotales()
                End If
            End If
        End Sub

        Private Sub RecalcularTotales()
            Dim subtotal As Decimal = ventaActual.Detalles.Sum(Function(d) d.Subtotal)
            Dim descPorc As Decimal = numDescuentoPorc.Value
            Dim descMonto As Decimal = Math.Round(subtotal * (descPorc / 100D), 2)
            Dim total As Decimal = subtotal - descMonto

            ventaActual.Subtotal = subtotal
            ventaActual.DescuentoPorcentaje = descPorc
            ventaActual.DescuentoMonto = descMonto
            ventaActual.Total = total

            txtSubtotal.Text = subtotal.ToString("C2")
            lblTotalPagar.Text = total.ToString("C2")

            CalcularVuelto()
        End Sub

        Private Sub TxtMontoAbonado_TextChanged(sender As Object, e As EventArgs)
            CalcularVuelto()
        End Sub

        Private Sub CalcularVuelto()
            Dim abonado As Decimal = 0D
            Decimal.TryParse(txtMontoAbonado.Text.Trim(), abonado)
            Dim vuelto As Decimal = abonado - ventaActual.Total
            If vuelto >= 0 Then
                lblVuelto.Text = vuelto.ToString("C2")
                lblVuelto.ForeColor = UITheme.ColorSuccess
            Else
                lblVuelto.Text = "$ 0,00"
                lblVuelto.ForeColor = UITheme.ColorTextSecondary
            End If
        End Sub

        Private Sub BtnNuevoCliente_Click(sender As Object, e As EventArgs)
            Dim frm As New FrmClientes()
            frm.ShowDialog()
            CargarClientes()
        End Sub

        Private Sub BtnFinalizarVenta_Click(sender As Object, e As EventArgs)
            FinalizarVenta()
        End Sub

        Private Sub FrmVentasPOS_KeyDown(sender As Object, e As KeyEventArgs)
            If e.KeyCode = Keys.F5 Then
                e.SuppressKeyPress = True
                FinalizarVenta()
            End If
        End Sub

        Private Sub FinalizarVenta()
            If ventaActual.Detalles.Count = 0 Then
                MessageBox.Show("El carrito de ventas está vacío. Agregue prendas para cobrar.", "Carrito Vacío", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Validar caja activa
            Dim cajaAbierta = cajaService.GetCajaAbierta(If(AuthService.CurrentUser IsNot Nothing, AuthService.CurrentUser.Id, 1))
            If cajaAbierta Is Nothing Then
                Dim resp = MessageBox.Show("No hay una caja abierta actualmente para registrar los cobros." & vbCrLf & vbCrLf & "¿Deseas abrir la caja del día ahora?", "Caja Cerrada", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If resp = DialogResult.Yes Then
                    Dim frmCaja As New FrmCaja()
                    frmCaja.ShowDialog()
                    cajaAbierta = cajaService.GetCajaAbierta(If(AuthService.CurrentUser IsNot Nothing, AuthService.CurrentUser.Id, 1))
                    If cajaAbierta Is Nothing Then Return
                Else
                    Return
                End If
            End If

            ' Obtener cliente seleccionado
            Dim clienteSel = TryCast(cboClientes.SelectedItem, Cliente)
            If clienteSel IsNot Nothing Then
                ventaActual.ClienteId = clienteSel.Id
                ventaActual.ClienteNombre = clienteSel.NombreCompleto
            Else
                ventaActual.ClienteId = 1
                ventaActual.ClienteNombre = "Consumidor Final"
            End If

            ventaActual.UsuarioId = If(AuthService.CurrentUser IsNot Nothing, AuthService.CurrentUser.Id, 1)
            ventaActual.CajaId = cajaAbierta.Id
            ventaActual.MetodoPago = cboMetodoPago.SelectedItem.ToString()
            ventaActual.TipoComprobante = "Ticket"
            ventaActual.NumeroTicket = ventaService.GenerarNumeroTicket()

            Decimal.TryParse(txtMontoAbonado.Text.Trim(), ventaActual.MontoAbonado)
            If ventaActual.MontoAbonado < ventaActual.Total AndAlso ventaActual.MetodoPago = "Efectivo" Then
                ventaActual.MontoAbonado = ventaActual.Total
            End If
            ventaActual.Vuelto = Math.Max(0D, ventaActual.MontoAbonado - ventaActual.Total)

            ' Procesar la venta en la base de datos
            Dim errMsg As String = ""
            If ventaService.ProcesarVenta(ventaActual, errMsg) Then
                ' Mostrar Ticket de Venta
                MostrarTicketImpresion(ventaActual)

                ' Resetear carrito
                ventaActual = New Venta()
                RefrescarGrillaCarrito()
                RecalcularTotales()
                txtMontoAbonado.Clear()
                lblVuelto.Text = "$ 0,00"
                txtBuscarArticulo.Focus()
            Else
                MessageBox.Show("No se pudo procesar la venta: " & errMsg, "Error en Venta", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End Sub

        Private Sub MostrarTicketImpresion(v As Venta)
            Dim cfg = configService.GetConfiguracion()
            Dim sb As New System.Text.StringBuilder()

            sb.AppendLine("==========================================")
            sb.AppendLine($"           {cfg.NombreComercio.ToUpper()}           ")
            sb.AppendLine($"CUIT: {cfg.Cuit} - {cfg.CondicionIva}")
            sb.AppendLine($"Dirección: {cfg.Direccion}")
            sb.AppendLine($"Tel: {cfg.Telefono}")
            sb.AppendLine("==========================================")
            sb.AppendLine($"TICKET Nº: {v.NumeroTicket}")
            sb.AppendLine($"Fecha: {v.Fecha.ToString("dd/MM/yyyy HH:mm:ss")}")
            sb.AppendLine($"Vendedor: {If(AuthService.CurrentUser IsNot Nothing, AuthService.CurrentUser.NombreCompleto, "Caja")}")
            sb.AppendLine($"Cliente: {v.ClienteNombre}")
            sb.AppendLine($"Medio de Pago: {v.MetodoPago}")
            sb.AppendLine("------------------------------------------")
            sb.AppendLine("CANT  DESCRIPCIÓN      TALLE       TOTAL  ")
            sb.AppendLine("------------------------------------------")

            For Each item In v.Detalles
                Dim desc = If(item.DescripcionArticulo.Length > 16, item.DescripcionArticulo.Substring(0, 16), item.DescripcionArticulo.PadRight(16))
                Dim talle = item.TalleNombre.PadRight(8)
                sb.AppendLine($"{item.Cantidad.ToString().PadRight(5)} {desc} {talle} {item.Subtotal.ToString("C2").PadLeft(9)}")
            Next

            sb.AppendLine("------------------------------------------")
            sb.AppendLine($"SUBTOTAL:        {v.Subtotal.ToString("C2").PadLeft(24)}")
            If v.DescuentoMonto > 0 Then
                sb.AppendLine($"DESCUENTO ({v.DescuentoPorcentaje}%):  -{v.DescuentoMonto.ToString("C2").PadLeft(20)}")
            End If
            sb.AppendLine($"TOTAL A PAGAR:   {v.Total.ToString("C2").PadLeft(24)}")
            sb.AppendLine($"PAGADO:          {v.MontoAbonado.ToString("C2").PadLeft(24)}")
            sb.AppendLine($"VUELTO:          {v.Vuelto.ToString("C2").PadLeft(24)}")
            sb.AppendLine("==========================================")
            sb.AppendLine($" {cfg.MensajeTicket} ")
            sb.AppendLine("==========================================")

            ' Modal con diseño térmico
            Dim frmTicket As New Form() With {
                .Text = "Comprobante de Venta - Ticket",
                .Size = New Size(420, 560),
                .StartPosition = FormStartPosition.CenterParent,
                .FormBorderStyle = FormBorderStyle.FixedDialog,
                .MaximizeBox = False,
                .MinimizeBox = False,
                .BackColor = Color.White
            }

            Dim txtRecibo As New TextBox() With {
                .Multiline = True,
                .ReadOnly = True,
                .Dock = DockStyle.Fill,
                .Font = New Font("Courier New", 9.5F, FontStyle.Regular),
                .BackColor = Color.White,
                .ForeColor = Color.Black,
                .ScrollBars = ScrollBars.Vertical,
                .Text = sb.ToString()
            }

            Dim btnImprimir As New Button() With {
                .Text = "🖨 Imprimir / Cerrar",
                .Dock = DockStyle.Bottom,
                .Height = 45
            }
            UITheme.StyleButton(btnImprimir, "Primary")
            AddHandler btnImprimir.Click, Sub() frmTicket.Close()

            frmTicket.Controls.AddRange({txtRecibo, btnImprimir})
            frmTicket.ShowDialog()
        End Sub

    End Class
End Namespace
