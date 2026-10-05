Imports System.Drawing
Imports System.Windows.Forms
Imports GestionComercial.Models
Imports GestionComercial.Services
Imports GestionComercial.UI

' ARCHIVO: FrmClientes.vb
' PROPÓSITO: Directorio Comercial, Fidelización y Gestión de Datos de Clientes
' - Registro de Clientes en Indumentaria:
'   Permite almacenar historial de compras, talles preferidos en el campo 'Notas',
'   y fecha de nacimiento para promociones de fidelización.
' - Integridad del Sistema ('Consumidor Final'):
'   El cliente con ID = 1 corresponde a 'Consumidor Final' (ventas de mostrador anónimas).
'   Se bloquea su eliminación para preservar la integridad referencial de la tabla 'ventas'.
' - Borrado Lógico y Reactivación:
'   Al igual que en usuarios y productos, los clientes se desactivan (Activo = 0)
'   para conservar la auditoría de ventas históricas asociadas a su DNI/CUIT.
'   Permite visualizar e alternar el estado (activo/inactivo) desde el listado mediante un filtro.

Namespace Forms
    Public Class FrmClientes
        Inherits Form

        Private clienteService As New ClienteService()

        Private txtBuscar As TextBox
        Private btnBuscar As Button
        Private btnNuevo As Button
        Private btnEditar As Button
        Private btnToggleActivo As Button
        Private chkMostrarInactivos As CheckBox
        Private dgvClientes As DataGridView
        Private lblTotal As Label

        Public Sub New()
            InitializeUI()
            LoadClientes()
        End Sub

        Private Sub InitializeUI()
            Me.Text = "Directorio de Clientes"
            Me.Size = New Size(1020, 580)
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.BackColor = UITheme.ColorBackground
            Me.Font = UITheme.FontRegular

            ' Header superior institucional
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

            ' Toolbar con FlowLayoutPanel para adaptar botones en resoluciones variadas
            Dim pnlToolbar As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 55,
                .BackColor = UITheme.ColorSurface
            }
            Dim flpToolbar As New FlowLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(10, 12, 10, 10),
                .WrapContents = False
            }

            Dim lblB As New Label() With {.Text = "Buscar por DNI o Nombre:", .Font = UITheme.FontBold, .AutoSize = True, .Margin = New Padding(0, 5, 4, 0)}
            txtBuscar = New TextBox() With {.Size = New Size(200, 26), .Margin = New Padding(0, 2, 0, 0)}
            UITheme.StyleTextBox(txtBuscar)
            AddHandler txtBuscar.KeyDown, Sub(s, e)
                                              If e.KeyCode = Keys.Enter Then LoadClientes()
                                          End Sub

            btnBuscar = New Button() With {.Text = "Buscar", .Height = 30, .AutoSize = True, .AutoSizeMode = AutoSizeMode.GrowAndShrink, .Padding = New Padding(10, 0, 10, 0), .Margin = New Padding(6, 2, 0, 0)}
            UITheme.StyleButton(btnBuscar, "Primary")
            AddHandler btnBuscar.Click, Sub() LoadClientes()

            btnNuevo = New Button() With {.Text = "+ Nuevo Cliente", .Height = 30, .AutoSize = True, .AutoSizeMode = AutoSizeMode.GrowAndShrink, .Padding = New Padding(10, 0, 10, 0), .Margin = New Padding(6, 2, 0, 0)}
            UITheme.StyleButton(btnNuevo, "Success")
            AddHandler btnNuevo.Click, AddressOf BtnNuevo_Click

            btnEditar = New Button() With {.Text = "Editar", .Height = 30, .AutoSize = True, .AutoSizeMode = AutoSizeMode.GrowAndShrink, .Padding = New Padding(10, 0, 10, 0), .Margin = New Padding(6, 2, 0, 0)}
            UITheme.StyleButton(btnEditar, "Secondary")
            AddHandler btnEditar.Click, AddressOf BtnEditar_Click

            btnToggleActivo = New Button() With {.Text = "🚫 Activar / Desactivar", .Height = 30, .AutoSize = True, .AutoSizeMode = AutoSizeMode.GrowAndShrink, .Padding = New Padding(10, 0, 10, 0), .Margin = New Padding(6, 2, 0, 0)}
            UITheme.StyleButton(btnToggleActivo, "Danger")
            AddHandler btnToggleActivo.Click, AddressOf BtnToggleActivo_Click

            chkMostrarInactivos = New CheckBox() With {.Text = "Mostrar inactivos", .Font = UITheme.FontBold, .AutoSize = True, .Margin = New Padding(10, 6, 0, 0)}
            AddHandler chkMostrarInactivos.CheckedChanged, Sub() LoadClientes()

            flpToolbar.Controls.AddRange({lblB, txtBuscar, btnBuscar, btnNuevo, btnEditar, btnToggleActivo, chkMostrarInactivos})
            pnlToolbar.Controls.Add(flpToolbar)

            ' Grilla de clientes registrados
            Dim pnlGrid As New Panel() With {.Dock = DockStyle.Fill, .Padding = New Padding(15)}
            dgvClientes = New DataGridView() With {.Dock = DockStyle.Fill}
            UITheme.StyleDataGrid(dgvClientes)
            ConfigurarColumnas()
            AddHandler dgvClientes.CellDoubleClick, Sub()
                                                    If AuthService.IsAdminOrManager Then EditarSeleccionado()
                                                End Sub
            AddHandler dgvClientes.CellFormatting, AddressOf DgvClientes_CellFormatting
            pnlGrid.Controls.Add(dgvClientes)

            ' Footer con conteo dinámico
            Dim pnlFooter As New Panel() With {.Dock = DockStyle.Bottom, .Height = 35, .BackColor = UITheme.ColorSurfaceMuted, .Padding = New Padding(15, 8, 15, 8)}
            lblTotal = New Label() With {.Text = "Clientes: 0", .Font = UITheme.FontBold, .ForeColor = UITheme.ColorTextSecondary, .AutoSize = True}
            pnlFooter.Controls.Add(lblTotal)

            ' Orden exacto de Docking de WinForms
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

            dgvClientes.Columns.Add("Apellido", "Apellido")
            dgvClientes.Columns("Apellido").Width = 130

            dgvClientes.Columns.Add("NombreCompleto", "Nombre")
            dgvClientes.Columns("NombreCompleto").Width = 130

            dgvClientes.Columns.Add("Telefono", "Teléfono")
            dgvClientes.Columns("Telefono").Width = 120

            dgvClientes.Columns.Add("Email", "Email")
            dgvClientes.Columns("Email").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

            dgvClientes.Columns.Add("Ciudad", "Ciudad")
            dgvClientes.Columns("Ciudad").Width = 110

            dgvClientes.Columns.Add("FechaNac", "Fecha Nac.")
            dgvClientes.Columns("FechaNac").Width = 95

            dgvClientes.Columns.Add("Activo", "Estado")
            dgvClientes.Columns("Activo").Width = 85
        End Sub

        ''' <summary>
        ''' Carga la lista de clientes filtrando por nombre o DNI y respetando el estado del filtro de inactivos.
        ''' </summary>
        Private Sub LoadClientes()
            Dim soloActivos As Boolean = Not chkMostrarInactivos.Checked
            Dim lista = clienteService.GetClientes(txtBuscar.Text.Trim(), soloActivos)
            dgvClientes.Rows.Clear()
            Dim activosCount As Integer = 0
            For Each c In lista
                Dim fnacStr = If(c.FechaNacimiento.HasValue, c.FechaNacimiento.Value.ToString("dd/MM/yyyy"), "-")
                Dim estadoStr = If(c.Activo, "Activo", "Inactivo")
                If c.Activo Then activosCount += 1
                dgvClientes.Rows.Add(c.Id, c.DniCuit, c.Apellido, c.Nombre, c.Telefono, c.Email, c.Ciudad, fnacStr, estadoStr)
            Next
            If chkMostrarInactivos.Checked Then
                lblTotal.Text = $"Total clientes: {lista.Count} ({activosCount} activos, {lista.Count - activosCount} inactivos)"
            Else
                lblTotal.Text = $"Total clientes registrados: {lista.Count}"
            End If
        End Sub

        ''' <summary>
        ''' Formato visual condicional para resaltar el estado activo / inactivo en verde o rojo.
        ''' </summary>
        Private Sub DgvClientes_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)
            If e.RowIndex >= 0 Then
                If dgvClientes.Columns(e.ColumnIndex).Name = "Activo" Then
                    Dim val = e.Value?.ToString()
                    If val = "Activo" Then
                        e.CellStyle.ForeColor = UITheme.ColorSuccess
                        e.CellStyle.Font = UITheme.FontBold
                    Else
                        e.CellStyle.ForeColor = UITheme.ColorDanger
                        e.CellStyle.Font = UITheme.FontBold
                    End If
                End If
            End If
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

        ''' <summary>
        ''' Control de activación/desactivación (baja lógica / reactivación) de clientes.
        ''' Protege al cliente ID=1 (Consumidor Final) impidiendo su desactivación.
        ''' </summary>
        Private Sub BtnToggleActivo_Click(sender As Object, e As EventArgs)
            If dgvClientes.CurrentRow IsNot Nothing Then
                Dim id As Integer = Convert.ToInt32(dgvClientes.CurrentRow.Cells("Id").Value)
                If id = 1 Then
                    MessageBox.Show("El cliente 'Consumidor Final' es del sistema y no puede ser desactivado.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                Dim nombre As String = dgvClientes.CurrentRow.Cells("NombreCompleto").Value.ToString()
                Dim estadoActual As String = dgvClientes.CurrentRow.Cells("Activo").Value?.ToString()
                Dim accion As String = If(estadoActual = "Activo", "desactivar", "reactivar")

                If MessageBox.Show($"¿Deseas {accion} al cliente '{nombre}'?", "Confirmar Acción", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Dim errMsg As String = ""
                    If clienteService.ToggleActivo(id, errMsg) Then
                        UITheme.ShowToast(Me, $"Cliente {accion}do correctamente.", "Success")
                        LoadClientes()
                    Else
                        UITheme.ShowToast(Me, "Error: " & errMsg, "Error")
                    End If
                End If
            End If
        End Sub

    End Class

    ''' <summary>
    ''' Diálogo modal para dar de alta un nuevo cliente o modificar datos de contacto y fidelización.
    ''' </summary>
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
        Private dtpFechaNac As DateTimePicker
        Private btnGuardar As Button
        Private btnCancelar As Button

        Public Sub New(clienteId As Integer)
            _clienteId = clienteId
            InitializeUI()
            LoadData()
        End Sub

        Private Sub InitializeUI()
            Me.Text = If(_clienteId = 0, "Nuevo Cliente", "Editar Cliente")
            Me.Size = New Size(540, 570)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.BackColor = UITheme.ColorBackground
            Me.Font = UITheme.FontRegular

            Dim lblDni As New Label() With {.Text = "DNI / CUIT *:", .Font = UITheme.FontBold, .Location = New Point(30, 20), .AutoSize = True}
            txtDni = New TextBox() With {.Location = New Point(30, 45), .Size = New Size(220, 26)}
            UITheme.StyleTextBox(txtDni)

            Dim lblTel As New Label() With {.Text = "Teléfono / Celular *:", .Font = UITheme.FontBold, .Location = New Point(270, 20), .AutoSize = True}
            txtTelefono = New TextBox() With {.Location = New Point(270, 45), .Size = New Size(220, 26)}
            UITheme.StyleTextBox(txtTelefono)

            Dim lblNom As New Label() With {.Text = "Nombre *:", .Font = UITheme.FontBold, .Location = New Point(30, 85), .AutoSize = True}
            txtNombre = New TextBox() With {.Location = New Point(30, 110), .Size = New Size(220, 26)}
            UITheme.StyleTextBox(txtNombre)

            Dim lblApe As New Label() With {.Text = "Apellido *:", .Font = UITheme.FontBold, .Location = New Point(270, 85), .AutoSize = True}
            txtApellido = New TextBox() With {.Location = New Point(270, 110), .Size = New Size(220, 26)}
            UITheme.StyleTextBox(txtApellido)

            Dim lblEmail As New Label() With {.Text = "Email *:", .Font = UITheme.FontBold, .Location = New Point(30, 150), .AutoSize = True}
            txtEmail = New TextBox() With {.Location = New Point(30, 175), .Size = New Size(460, 26)}
            UITheme.StyleTextBox(txtEmail)

            Dim lblDir As New Label() With {.Text = "Dirección *:", .Font = UITheme.FontBold, .Location = New Point(30, 215), .AutoSize = True}
            txtDireccion = New TextBox() With {.Location = New Point(30, 240), .Size = New Size(290, 26)}
            UITheme.StyleTextBox(txtDireccion)

            Dim lblCiu As New Label() With {.Text = "Ciudad *:", .Font = UITheme.FontBold, .Location = New Point(330, 215), .AutoSize = True}
            txtCiudad = New TextBox() With {.Location = New Point(330, 240), .Size = New Size(160, 26)}
            UITheme.StyleTextBox(txtCiudad)

            Dim lblFnac As New Label() With {.Text = "Fecha de Nacimiento *:", .Font = UITheme.FontBold, .Location = New Point(30, 280), .AutoSize = True}
            dtpFechaNac = New DateTimePicker() With {.Location = New Point(30, 305), .Size = New Size(460, 26), .Format = DateTimePickerFormat.Short, .Value = DateTime.Today.AddYears(-25)}

            Dim lblNotas As New Label() With {.Text = "Notas / Preferencias (Opcional):", .Font = UITheme.FontBold, .Location = New Point(30, 345), .AutoSize = True}
            txtNotas = New TextBox() With {.Location = New Point(30, 370), .Size = New Size(460, 55), .Multiline = True}
            UITheme.StyleTextBox(txtNotas)

            Dim pnlBottom As New Panel() With {.Dock = DockStyle.Bottom, .Height = 55, .BackColor = Color.FromArgb(241, 245, 249), .Padding = New Padding(20, 10, 20, 10)}
            btnGuardar = New Button() With {.Text = "💾 Guardar Cliente", .Dock = DockStyle.Right, .Width = 160}
            UITheme.StyleButton(btnGuardar, "Success")
            AddHandler btnGuardar.Click, AddressOf BtnGuardar_Click

            btnCancelar = New Button() With {.Text = "Cancelar", .Dock = DockStyle.Left, .Width = 100}
            UITheme.StyleButton(btnCancelar, "Secondary")
            AddHandler btnCancelar.Click, Sub() Me.Close()

            pnlBottom.Controls.AddRange({btnGuardar, btnCancelar})

            Me.Controls.AddRange({lblDni, txtDni, lblTel, txtTelefono, lblNom, txtNombre, lblApe, txtApellido, lblEmail, txtEmail, lblDir, txtDireccion, lblCiu, txtCiudad, lblFnac, dtpFechaNac, lblNotas, txtNotas, pnlBottom})
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
                    If clienteActual.FechaNacimiento.HasValue Then
                        dtpFechaNac.Value = clienteActual.FechaNacimiento.Value
                    End If
                End If
            End If
        End Sub

        ''' <summary>
        ''' Valida campos obligatorios y delega en ClienteService.GuardarCliente
        ''' que valida unicidad de DNI/CUIT antes del INSERT o UPDATE.
        ''' </summary>
        Private Sub BtnGuardar_Click(sender As Object, e As EventArgs)
            If String.IsNullOrWhiteSpace(txtDni.Text) OrElse
               String.IsNullOrWhiteSpace(txtNombre.Text) OrElse
               String.IsNullOrWhiteSpace(txtApellido.Text) OrElse
               String.IsNullOrWhiteSpace(txtTelefono.Text) OrElse
               String.IsNullOrWhiteSpace(txtEmail.Text) OrElse
               String.IsNullOrWhiteSpace(txtDireccion.Text) OrElse
               String.IsNullOrWhiteSpace(txtCiudad.Text) Then
                MessageBox.Show("Por favor complete todos los campos obligatorios (DNI, Nombre, Apellido, Teléfono, Email, Dirección, Ciudad).", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
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
            clienteActual.FechaNacimiento = dtpFechaNac.Value.Date
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
