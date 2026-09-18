Imports System.Drawing
Imports System.Windows.Forms
Imports GestionComercial.Models
Imports GestionComercial.Services
Imports GestionComercial.UI

Namespace Forms
    Public Class FrmClientes
        Inherits Form

        Private clienteService As New ClienteService()

        Private txtBuscar As TextBox
        Private btnBuscar As Button
        Private btnNuevo As Button
        Private btnEditar As Button
        Private btnEliminar As Button
        Private dgvClientes As DataGridView
        Private lblTotal As Label

        Public Sub New()
            InitializeUI()
            LoadClientes()
        End Sub

        Private Sub InitializeUI()
            Me.Text = "Directorio de Clientes"
            Me.Size = New Size(950, 580)
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
                .Text = "Gestión de clientes",
                .Font = UITheme.FontHeading,
                .ForeColor = Color.White,
                .AutoSize = True,
                .Location = New Point(15, 15)
            }
            pnlHeader.Controls.Add(lblTitle)

            ' Toolbar
            Dim pnlToolbar As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 95,
                .BackColor = UITheme.ColorSurface,
                .Padding = New Padding(15, 15, 15, 10)
            }

            Dim lblB As New Label() With {.Text = "Buscar por DNI o nombre:", .Font = UITheme.FontBold, .Location = New Point(15, 20), .AutoSize = True}
            txtBuscar = New TextBox() With {.Location = New Point(215, 18), .Size = New Size(260, 26)}
            UITheme.StyleTextBox(txtBuscar)
            AddHandler txtBuscar.KeyDown, Sub(s, e)
                                              If e.KeyCode = Keys.Enter Then LoadClientes()
                                          End Sub

            btnBuscar = New Button() With {.Text = "Buscar", .Location = New Point(485, 16), .Size = New Size(80, 30)}
            UITheme.StyleButton(btnBuscar, "Primary")
            AddHandler btnBuscar.Click, Sub() LoadClientes()

            btnNuevo = New Button() With {.Text = "+ Nuevo Cliente", .Location = New Point(575, 16), .Size = New Size(130, 30)}
            UITheme.StyleButton(btnNuevo, "Success")
            AddHandler btnNuevo.Click, AddressOf BtnNuevo_Click

            btnEditar = New Button() With {.Text = "Editar", .Location = New Point(715, 16), .Size = New Size(95, 30)}
            UITheme.StyleButton(btnEditar, "Secondary")
            AddHandler btnEditar.Click, AddressOf BtnEditar_Click

            btnEliminar = New Button() With {.Text = "Eliminar", .Location = New Point(820, 16), .Size = New Size(95, 30)}
            UITheme.StyleButton(btnEliminar, "Danger")
            AddHandler btnEliminar.Click, AddressOf BtnEliminar_Click

            btnEditar.Visible = AuthService.IsAdminOrManager
            btnEliminar.Visible = AuthService.IsAdminOrManager

            pnlToolbar.Controls.AddRange({lblB, txtBuscar, btnBuscar, btnNuevo, btnEditar, btnEliminar})

            ' Grilla
            Dim pnlGrid As New Panel() With {.Dock = DockStyle.Fill, .Padding = New Padding(15)}
            dgvClientes = New DataGridView() With {.Dock = DockStyle.Fill}
            UITheme.StyleDataGrid(dgvClientes)
            ConfigurarColumnas()
            AddHandler dgvClientes.CellDoubleClick, Sub()
                                                    If AuthService.IsAdminOrManager Then EditarSeleccionado()
                                                End Sub
            pnlGrid.Controls.Add(dgvClientes)

            ' Footer
            Dim pnlFooter As New Panel() With {.Dock = DockStyle.Bottom, .Height = 35, .BackColor = UITheme.ColorSurfaceMuted, .Padding = New Padding(15, 8, 15, 8)}
            lblTotal = New Label() With {.Text = "Clientes: 0", .Font = UITheme.FontBold, .ForeColor = UITheme.ColorTextSecondary, .AutoSize = True}
            pnlFooter.Controls.Add(lblTotal)

            ' Orden exacto de Docking: Fill primero, Bottom segundo, Toolbar tercero, Header último
            Me.Controls.Add(pnlGrid)
            Me.Controls.Add(pnlFooter)
            Me.Controls.Add(pnlToolbar)
            Me.Controls.Add(pnlHeader)
        End Sub

        Private Sub ConfigurarColumnas()
            dgvClientes.Columns.Clear()
            dgvClientes.Columns.Add("Id", "ID")
            dgvClientes.Columns("Id").Width = 50

            dgvClientes.Columns.Add("Dni", "DNI / CUIT")
            dgvClientes.Columns("Dni").Width = 120

            dgvClientes.Columns.Add("NombreCompleto", "Apellido y Nombre")
            dgvClientes.Columns("NombreCompleto").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

            dgvClientes.Columns.Add("Telefono", "Teléfono / WhatsApp")
            dgvClientes.Columns("Telefono").Width = 140

            dgvClientes.Columns.Add("Email", "Email")
            dgvClientes.Columns("Email").Width = 180

            dgvClientes.Columns.Add("Ciudad", "Ciudad")
            dgvClientes.Columns("Ciudad").Width = 120
        End Sub

        Private Sub LoadClientes()
            Dim lista = clienteService.GetClientes(txtBuscar.Text.Trim(), True)
            dgvClientes.Rows.Clear()
            For Each c In lista
                dgvClientes.Rows.Add(c.Id, c.DniCuit, c.NombreCompleto, c.Telefono, c.Email, c.Ciudad)
            Next
            lblTotal.Text = $"Total clientes registrados: {lista.Count}"
        End Sub

        Private Sub BtnNuevo_Click(sender As Object, e As EventArgs)
            Dim frmEdit As New FrmClienteEditor(0)
            If frmEdit.ShowDialog() = DialogResult.OK Then
                LoadClientes()
            End If
        End Sub

        Private Sub BtnEditar_Click(sender As Object, e As EventArgs)
            EditarSeleccionado()
        End Sub

        Private Sub EditarSeleccionado()
            If dgvClientes.CurrentRow IsNot Nothing Then
                Dim id As Integer = Convert.ToInt32(dgvClientes.CurrentRow.Cells("Id").Value)
                Dim frmEdit As New FrmClienteEditor(id)
                If frmEdit.ShowDialog() = DialogResult.OK Then
                    LoadClientes()
                End If
            End If
        End Sub

        Private Sub BtnEliminar_Click(sender As Object, e As EventArgs)
            If dgvClientes.CurrentRow IsNot Nothing Then
                Dim id As Integer = Convert.ToInt32(dgvClientes.CurrentRow.Cells("Id").Value)
                If id = 1 Then
                    MessageBox.Show("El cliente 'Consumidor Final' es del sistema y no puede ser eliminado.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                Dim nombre As String = dgvClientes.CurrentRow.Cells("NombreCompleto").Value.ToString()
                If MessageBox.Show($"¿Deseas dar de baja a {nombre}?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Dim errMsg As String = ""
                    If clienteService.EliminarCliente(id, errMsg) Then
                        UITheme.ShowToast(Me, "Cliente dado de baja correctamente.", "Success")
                        LoadClientes()
                    Else
                        UITheme.ShowToast(Me, "Error: " & errMsg, "Error")
                    End If
                End If
            End If
        End Sub

    End Class

    Public Class FrmClienteEditor
        Inherits Form

        Private _clienteId As Integer
        Private clienteService As New ClienteService()
        Private clienteActual As Cliente

        Private txtDni As TextBox
        Private txtNombre As TextBox
        Private txtApellido As TextBox
        Private txtTelefono As TextBox
        Private txtEmail As TextBox
        Private txtDireccion As TextBox
        Private txtCiudad As TextBox
        Private txtNotas As TextBox
        Private btnGuardar As Button
        Private btnCancelar As Button

        Public Sub New(clienteId As Integer)
            _clienteId = clienteId
            InitializeUI()
            LoadData()
        End Sub

        Private Sub InitializeUI()
            Me.Text = If(_clienteId = 0, "Nuevo Cliente", "Editar Cliente")
            Me.Size = New Size(540, 530)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.BackColor = UITheme.ColorBackground
            Me.Font = UITheme.FontRegular

            Dim lblDni As New Label() With {.Text = "DNI / CUIT:", .Font = UITheme.FontBold, .Location = New Point(30, 20), .AutoSize = True}
            txtDni = New TextBox() With {.Location = New Point(30, 50), .Size = New Size(220, 26)}
            UITheme.StyleTextBox(txtDni)

            Dim lblTel As New Label() With {.Text = "Teléfono / Celular:", .Font = UITheme.FontBold, .Location = New Point(270, 20), .AutoSize = True}
            txtTelefono = New TextBox() With {.Location = New Point(270, 50), .Size = New Size(220, 26)}
            UITheme.StyleTextBox(txtTelefono)

            Dim lblNom As New Label() With {.Text = "Nombre:", .Font = UITheme.FontBold, .Location = New Point(30, 95), .AutoSize = True}
            txtNombre = New TextBox() With {.Location = New Point(30, 125), .Size = New Size(220, 26)}
            UITheme.StyleTextBox(txtNombre)

            Dim lblApe As New Label() With {.Text = "Apellido:", .Font = UITheme.FontBold, .Location = New Point(270, 95), .AutoSize = True}
            txtApellido = New TextBox() With {.Location = New Point(270, 125), .Size = New Size(220, 26)}
            UITheme.StyleTextBox(txtApellido)

            Dim lblEmail As New Label() With {.Text = "Email:", .Font = UITheme.FontBold, .Location = New Point(30, 170), .AutoSize = True}
            txtEmail = New TextBox() With {.Location = New Point(30, 200), .Size = New Size(460, 26)}
            UITheme.StyleTextBox(txtEmail)

            Dim lblDir As New Label() With {.Text = "Dirección:", .Font = UITheme.FontBold, .Location = New Point(30, 245), .AutoSize = True}
            txtDireccion = New TextBox() With {.Location = New Point(30, 275), .Size = New Size(290, 26)}
            UITheme.StyleTextBox(txtDireccion)

            Dim lblCiu As New Label() With {.Text = "Ciudad:", .Font = UITheme.FontBold, .Location = New Point(330, 245), .AutoSize = True}
            txtCiudad = New TextBox() With {.Location = New Point(330, 275), .Size = New Size(160, 26)}
            UITheme.StyleTextBox(txtCiudad)

            Dim lblNotas As New Label() With {.Text = "Notas / Preferencias (Talles preferidos, etc.):", .Font = UITheme.FontBold, .Location = New Point(30, 320), .AutoSize = True}
            txtNotas = New TextBox() With {.Location = New Point(30, 350), .Size = New Size(460, 60), .Multiline = True}
            UITheme.StyleTextBox(txtNotas)

            Dim pnlBottom As New Panel() With {.Dock = DockStyle.Bottom, .Height = 55, .BackColor = UITheme.ColorSurfaceMuted, .Padding = New Padding(20, 10, 20, 10)}
            btnGuardar = New Button() With {.Text = "💾 Guardar Cliente", .Dock = DockStyle.Right, .Width = 160}
            UITheme.StyleButton(btnGuardar, "Success")
            AddHandler btnGuardar.Click, AddressOf BtnGuardar_Click

            btnCancelar = New Button() With {.Text = "Cancelar", .Dock = DockStyle.Left, .Width = 100}
            UITheme.StyleButton(btnCancelar, "Secondary")
            AddHandler btnCancelar.Click, Sub() Me.Close()

            pnlBottom.Controls.AddRange({btnGuardar, btnCancelar})

            Me.Controls.AddRange({lblDni, txtDni, lblTel, txtTelefono, lblNom, txtNombre, lblApe, txtApellido, lblEmail, txtEmail, lblDir, txtDireccion, lblCiu, txtCiudad, lblNotas, txtNotas, pnlBottom})
        End Sub

        Private Sub LoadData()
            If _clienteId = 0 Then
                clienteActual = New Cliente()
            Else
                clienteActual = clienteService.GetClienteById(_clienteId)
                If clienteActual IsNot Nothing Then
                    txtDni.Text = clienteActual.DniCuit
                    txtNombre.Text = clienteActual.Nombre
                    txtApellido.Text = clienteActual.Apellido
                    txtTelefono.Text = clienteActual.Telefono
                    txtEmail.Text = clienteActual.Email
                    txtDireccion.Text = clienteActual.Direccion
                    txtCiudad.Text = clienteActual.Ciudad
                    txtNotas.Text = clienteActual.Notas
                End If
            End If
        End Sub

        Private Sub BtnGuardar_Click(sender As Object, e As EventArgs)
            If String.IsNullOrWhiteSpace(txtDni.Text) OrElse String.IsNullOrWhiteSpace(txtNombre.Text) Then
                MessageBox.Show("El DNI y el Nombre son obligatorios.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            clienteActual.DniCuit = txtDni.Text.Trim()
            clienteActual.Nombre = txtNombre.Text.Trim()
            clienteActual.Apellido = txtApellido.Text.Trim()
            clienteActual.Telefono = txtTelefono.Text.Trim()
            clienteActual.Email = txtEmail.Text.Trim()
            clienteActual.Direccion = txtDireccion.Text.Trim()
            clienteActual.Ciudad = txtCiudad.Text.Trim()
            clienteActual.Notas = txtNotas.Text.Trim()
            clienteActual.Activo = True

            Dim errMsg As String = ""
            If clienteService.GuardarCliente(clienteActual, errMsg) Then
                MessageBox.Show("Cliente guardado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.DialogResult = DialogResult.OK
                Me.Close()
            Else
                MessageBox.Show("Error al guardar: " & errMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End Sub

    End Class
End Namespace
