Imports System.Drawing
Imports System.Windows.Forms
Imports GestionComercial.Models
Imports GestionComercial.Services
Imports GestionComercial.UI

Namespace Forms
    Public Class FrmUsuarios
        Inherits Form

        Private usuarioService As New UsuarioService()

        Private dgvUsuarios As DataGridView
        Private btnNuevo As Button
        Private btnEditar As Button
        Private btnCambiarPass As Button
        Private btnToggleActivo As Button
        Private btnRefrescar As Button
        Private lblTotalUsuarios As Label

        Public Sub New()
            InitializeUI()
            LoadUsuarios()
        End Sub

        Private Sub InitializeUI()
            Me.Text = "Gestión de Usuarios y Vendedores"
            Me.Size = New Size(1020, 640)
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
                .Text = "👥 GESTIÓN DE USUARIOS Y VENDEDORES",
                .Font = UITheme.FontHeading,
                .ForeColor = Color.White,
                .AutoSize = True,
                .Location = New Point(15, 15)
            }
            pnlHeader.Controls.Add(lblTitle)

            ' Toolbar con FlowLayoutPanel para evitar recorte de texto en botones
            Dim pnlToolbar As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 60,
                .BackColor = UITheme.ColorSurface
            }
            Dim flpToolbar As New FlowLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(10, 12, 10, 10),
                .WrapContents = False
            }

            btnNuevo = New Button() With {.Text = "+ Nuevo Usuario", .Height = 34, .AutoSize = True, .AutoSizeMode = AutoSizeMode.GrowAndShrink, .Padding = New Padding(10, 0, 10, 0)}
            UITheme.StyleButton(btnNuevo, "Success")
            AddHandler btnNuevo.Click, AddressOf BtnNuevo_Click

            btnEditar = New Button() With {.Text = "✏ Editar Datos", .Height = 34, .AutoSize = True, .AutoSizeMode = AutoSizeMode.GrowAndShrink, .Padding = New Padding(10, 0, 10, 0), .Margin = New Padding(6, 0, 0, 0)}
            UITheme.StyleButton(btnEditar, "Secondary")
            AddHandler btnEditar.Click, AddressOf BtnEditar_Click

            btnCambiarPass = New Button() With {.Text = "🔑 Cambiar Clave", .Height = 34, .AutoSize = True, .AutoSizeMode = AutoSizeMode.GrowAndShrink, .Padding = New Padding(10, 0, 10, 0), .Margin = New Padding(6, 0, 0, 0)}
            UITheme.StyleButton(btnCambiarPass, "Secondary")
            AddHandler btnCambiarPass.Click, AddressOf BtnCambiarPass_Click

            btnToggleActivo = New Button() With {.Text = "🚫 Activar / Desactivar", .Height = 34, .AutoSize = True, .AutoSizeMode = AutoSizeMode.GrowAndShrink, .Padding = New Padding(10, 0, 10, 0), .Margin = New Padding(6, 0, 0, 0)}
            UITheme.StyleButton(btnToggleActivo, "Danger")
            AddHandler btnToggleActivo.Click, AddressOf BtnToggleActivo_Click

            btnRefrescar = New Button() With {.Text = "🔄 Actualizar", .Height = 34, .AutoSize = True, .AutoSizeMode = AutoSizeMode.GrowAndShrink, .Padding = New Padding(10, 0, 10, 0), .Margin = New Padding(6, 0, 0, 0)}
            UITheme.StyleButton(btnRefrescar, "Secondary")
            AddHandler btnRefrescar.Click, Sub() LoadUsuarios()

            flpToolbar.Controls.AddRange({btnNuevo, btnEditar, btnCambiarPass, btnToggleActivo, btnRefrescar})
            pnlToolbar.Controls.Add(flpToolbar)

            ' Grilla
            Dim pnlGrid As New Panel() With {.Dock = DockStyle.Fill, .Padding = New Padding(15)}
            dgvUsuarios = New DataGridView() With {.Dock = DockStyle.Fill}
            UITheme.StyleDataGrid(dgvUsuarios)
            ConfigurarColumnas()
            AddHandler dgvUsuarios.CellDoubleClick, AddressOf DgvUsuarios_CellDoubleClick
            AddHandler dgvUsuarios.CellFormatting, AddressOf DgvUsuarios_CellFormatting
            pnlGrid.Controls.Add(dgvUsuarios)

            ' Footer
            Dim pnlFooter As New Panel() With {
                .Dock = DockStyle.Bottom,
                .Height = 40,
                .BackColor = Color.FromArgb(241, 245, 249),
                .Padding = New Padding(15, 10, 15, 10)
            }
            lblTotalUsuarios = New Label() With {
                .Text = "Cargando usuarios...",
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorTextSecondary,
                .AutoSize = True,
                .Location = New Point(15, 10)
            }
            pnlFooter.Controls.Add(lblTotalUsuarios)

            ' Orden exacto de docking
            Me.Controls.Add(pnlGrid)
            Me.Controls.Add(pnlFooter)
            Me.Controls.Add(pnlToolbar)
            Me.Controls.Add(pnlHeader)
        End Sub

        Private Sub ConfigurarColumnas()
            dgvUsuarios.Columns.Clear()
            dgvUsuarios.Columns.Add("Id", "ID")
            dgvUsuarios.Columns("Id").Width = 50

            dgvUsuarios.Columns.Add("Username", "Usuario")
            dgvUsuarios.Columns("Username").Width = 120

            dgvUsuarios.Columns.Add("Apellido", "Apellido")
            dgvUsuarios.Columns("Apellido").Width = 140

            dgvUsuarios.Columns.Add("Nombre", "Nombre")
            dgvUsuarios.Columns("Nombre").Width = 140

            dgvUsuarios.Columns.Add("Rol", "Rol")
            dgvUsuarios.Columns("Rol").Width = 130

            dgvUsuarios.Columns.Add("Telefono", "Teléfono")
            dgvUsuarios.Columns("Telefono").Width = 120

            dgvUsuarios.Columns.Add("Email", "Email")
            dgvUsuarios.Columns("Email").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

            dgvUsuarios.Columns.Add("FechaNac", "Fecha Nac.")
            dgvUsuarios.Columns("FechaNac").Width = 100

            dgvUsuarios.Columns.Add("Activo", "Estado")
            dgvUsuarios.Columns("Activo").Width = 90

            dgvUsuarios.Columns.Add("UltimoLogin", "Último Acceso")
            dgvUsuarios.Columns("UltimoLogin").Width = 150
        End Sub

        Public Sub LoadUsuarios()
            Dim lista = usuarioService.GetUsuarios()
            dgvUsuarios.Rows.Clear()

            Dim activosCount As Integer = 0
            For Each u In lista
                Dim ultAcceso = If(u.UltimoLogin.HasValue, u.UltimoLogin.Value.ToString("dd/MM/yyyy HH:mm"), "Nunca")
                Dim fnacStr = If(u.FechaNacimiento.HasValue, u.FechaNacimiento.Value.ToString("dd/MM/yyyy"), "-")
                Dim estadoStr = If(u.Activo, "Activo", "Inactivo")
                If u.Activo Then activosCount += 1
                dgvUsuarios.Rows.Add(u.Id, u.Username, u.Apellido, u.Nombre, u.Rol, u.Telefono, u.Email, fnacStr, estadoStr, ultAcceso)
            Next
            lblTotalUsuarios.Text = $"Total de usuarios: {lista.Count} ({activosCount} activos, {lista.Count - activosCount} inactivos)"
        End Sub

        Private Sub DgvUsuarios_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)
            If e.RowIndex >= 0 Then
                If dgvUsuarios.Columns(e.ColumnIndex).Name = "Activo" Then
                    Dim val = e.Value?.ToString()
                    If val = "Activo" Then
                        e.CellStyle.ForeColor = UITheme.ColorSuccess
                        e.CellStyle.Font = UITheme.FontBold
                    Else
                        e.CellStyle.ForeColor = UITheme.ColorDanger
                        e.CellStyle.Font = UITheme.FontBold
                    End If
                ElseIf dgvUsuarios.Columns(e.ColumnIndex).Name = "Rol" Then
                    Dim val = e.Value?.ToString()
                    If val = "Administrador" Then
                        e.CellStyle.ForeColor = UITheme.ColorPrimaryDark
                        e.CellStyle.Font = UITheme.FontBold
                    End If
                End If
            End If
        End Sub

        Private Sub DgvUsuarios_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
            If e.RowIndex >= 0 Then
                EditarSeleccionado()
            End If
        End Sub

        Private Sub BtnNuevo_Click(sender As Object, e As EventArgs)
            Dim frmEdit As New FrmUsuarioEditor(0)
            If frmEdit.ShowDialog() = DialogResult.OK Then
                LoadUsuarios()
            End If
        End Sub

        Private Sub BtnEditar_Click(sender As Object, e As EventArgs)
            EditarSeleccionado()
        End Sub

        Private Sub EditarSeleccionado()
            If dgvUsuarios.CurrentRow IsNot Nothing Then
                Dim uId As Integer = Convert.ToInt32(dgvUsuarios.CurrentRow.Cells("Id").Value)
                Dim frmEdit As New FrmUsuarioEditor(uId)
                If frmEdit.ShowDialog() = DialogResult.OK Then
                    LoadUsuarios()
                End If
            Else
                MessageBox.Show("Por favor seleccione un usuario de la lista.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End Sub

        Private Sub BtnCambiarPass_Click(sender As Object, e As EventArgs)
            If dgvUsuarios.CurrentRow IsNot Nothing Then
                Dim uId As Integer = Convert.ToInt32(dgvUsuarios.CurrentRow.Cells("Id").Value)
                Dim username As String = dgvUsuarios.CurrentRow.Cells("Username").Value.ToString()
                Dim frmPass As New FrmCambiarPassword(uId, username)
                If frmPass.ShowDialog() = DialogResult.OK Then
                    UITheme.ShowToast(Me, "Contraseña actualizada exitosamente.", "Success")
                End If
            Else
                MessageBox.Show("Por favor seleccione un usuario de la lista.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End Sub

        Private Sub BtnToggleActivo_Click(sender As Object, e As EventArgs)
            If dgvUsuarios.CurrentRow IsNot Nothing Then
                Dim uId As Integer = Convert.ToInt32(dgvUsuarios.CurrentRow.Cells("Id").Value)
                Dim username As String = dgvUsuarios.CurrentRow.Cells("Username").Value.ToString()
                Dim estadoActual As String = dgvUsuarios.CurrentRow.Cells("Activo").Value.ToString()
                Dim accion = If(estadoActual = "Activo", "desactivar", "reactivar")

                Dim resp = MessageBox.Show($"¿Deseas {accion} el acceso al usuario '{username}'?", "Confirmar Acción", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If resp = DialogResult.Yes Then
                    Dim errMsg As String = ""
                    If usuarioService.ToggleActivo(uId, errMsg) Then
                        UITheme.ShowToast(Me, $"Usuario {accion}do correctamente.", "Success")
                        LoadUsuarios()
                    Else
                        MessageBox.Show(errMsg, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End If
                End If
            End If
        End Sub

    End Class

    ''' <summary>
    ''' Modal para Crear o Editar Datos de Usuario
    ''' </summary>
    Public Class FrmUsuarioEditor
        Inherits Form

        Private _usuarioId As Integer
        Private usuarioService As New UsuarioService()

        Private txtNombre As TextBox
        Private txtApellido As TextBox
        Private txtUsername As TextBox
        Private cboRol As ComboBox
        Private txtPassword As TextBox
        Private txtTelefono As TextBox
        Private txtEmail As TextBox
        Private txtDireccion As TextBox
        Private txtCiudad As TextBox
        Private txtNotas As TextBox
        Private dtpFechaNac As DateTimePicker
        Private chkTieneFechaNac As CheckBox
        Private btnGuardar As Button
        Private btnCancelar As Button

        Public Sub New(usuarioId As Integer)
            _usuarioId = usuarioId
            InitializeUI()
            If _usuarioId > 0 Then
                LoadUserData()
            End If
        End Sub

        Private Sub InitializeUI()
            Me.Text = If(_usuarioId = 0, "Crear Nuevo Usuario / Empleado", "Editar Datos de Usuario / Empleado")
            Me.Size = New Size(560, If(_usuarioId = 0, 620, 570))
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.BackColor = UITheme.ColorBackground
            Me.Font = UITheme.FontRegular

            ' --- Fila 1: Nombre y Apellido ---
            Dim lblNom As New Label() With {.Text = "Nombre:", .Font = UITheme.FontBold, .Location = New Point(25, 20), .AutoSize = True}
            txtNombre = New TextBox() With {.Location = New Point(25, 45), .Size = New Size(230, 26)}
            UITheme.StyleTextBox(txtNombre)

            Dim lblApe As New Label() With {.Text = "Apellido:", .Font = UITheme.FontBold, .Location = New Point(275, 20), .AutoSize = True}
            txtApellido = New TextBox() With {.Location = New Point(275, 45), .Size = New Size(230, 26)}
            UITheme.StyleTextBox(txtApellido)

            ' --- Fila 2: Username y Rol ---
            Dim lblUser As New Label() With {.Text = "Usuario (Login):", .Font = UITheme.FontBold, .Location = New Point(25, 90), .AutoSize = True}
            txtUsername = New TextBox() With {.Location = New Point(25, 115), .Size = New Size(230, 26)}
            UITheme.StyleTextBox(txtUsername)

            Dim lblRol As New Label() With {.Text = "Rol en el Sistema:", .Font = UITheme.FontBold, .Location = New Point(275, 90), .AutoSize = True}
            cboRol = New ComboBox() With {.Location = New Point(275, 115), .Size = New Size(230, 26), .DropDownStyle = ComboBoxStyle.DropDownList}
            cboRol.Items.AddRange({"Vendedor", "Administrador"})
            cboRol.SelectedIndex = 0

            ' --- Fila 3: Teléfono y Email ---
            Dim lblTel As New Label() With {.Text = "Teléfono / WhatsApp:", .Font = UITheme.FontBold, .Location = New Point(25, 160), .AutoSize = True}
            txtTelefono = New TextBox() With {.Location = New Point(25, 185), .Size = New Size(230, 26)}
            UITheme.StyleTextBox(txtTelefono)

            Dim lblEmail As New Label() With {.Text = "Email:", .Font = UITheme.FontBold, .Location = New Point(275, 160), .AutoSize = True}
            txtEmail = New TextBox() With {.Location = New Point(275, 185), .Size = New Size(230, 26)}
            UITheme.StyleTextBox(txtEmail)

            ' --- Fila 4: Dirección y Ciudad ---
            Dim lblDir As New Label() With {.Text = "Dirección:", .Font = UITheme.FontBold, .Location = New Point(25, 230), .AutoSize = True}
            txtDireccion = New TextBox() With {.Location = New Point(25, 255), .Size = New Size(300, 26)}
            UITheme.StyleTextBox(txtDireccion)

            Dim lblCiu As New Label() With {.Text = "Ciudad:", .Font = UITheme.FontBold, .Location = New Point(340, 230), .AutoSize = True}
            txtCiudad = New TextBox() With {.Location = New Point(340, 255), .Size = New Size(165, 26)}
            UITheme.StyleTextBox(txtCiudad)

            ' --- Fila 5: Fecha de Nacimiento ---
            Dim lblFnac As New Label() With {.Text = "Fecha de Nacimiento:", .Font = UITheme.FontBold, .Location = New Point(25, 300), .AutoSize = True}
            dtpFechaNac = New DateTimePicker() With {.Location = New Point(25, 325), .Size = New Size(200, 26), .Format = DateTimePickerFormat.Short, .Value = DateTime.Today.AddYears(-25)}
            chkTieneFechaNac = New CheckBox() With {.Text = "Sin fecha", .Location = New Point(240, 328), .AutoSize = True, .Checked = True}
            dtpFechaNac.Enabled = Not chkTieneFechaNac.Checked
            AddHandler chkTieneFechaNac.CheckedChanged, Sub()
                                                             dtpFechaNac.Enabled = Not chkTieneFechaNac.Checked
                                                         End Sub

            ' --- Fila 6: Notas ---
            Dim lblNotas As New Label() With {.Text = "Notas / Observaciones:", .Font = UITheme.FontBold, .Location = New Point(25, 370), .AutoSize = True}
            txtNotas = New TextBox() With {.Location = New Point(25, 395), .Size = New Size(480, 55), .Multiline = True}
            UITheme.StyleTextBox(txtNotas)

            ' --- Contraseña (solo en modo creación) ---
            Dim nextY As Integer = 465
            If _usuarioId = 0 Then
                Dim lblPass As New Label() With {.Text = "Contraseña Inicial (mín. 4 caracteres):", .Font = UITheme.FontBold, .Location = New Point(25, nextY), .AutoSize = True}
                txtPassword = New TextBox() With {.Location = New Point(25, nextY + 25), .Size = New Size(480, 26), .UseSystemPasswordChar = True}
                UITheme.StyleTextBox(txtPassword)
                Me.Controls.AddRange({lblPass, txtPassword})
                nextY += 65
            End If

            ' --- Botones ---
            Dim pnlBottom As New Panel() With {.Dock = DockStyle.Bottom, .Height = 55, .BackColor = Color.FromArgb(241, 245, 249), .Padding = New Padding(20, 10, 20, 10)}
            btnGuardar = New Button() With {.Text = "💾 Guardar Empleado", .Dock = DockStyle.Right, .Width = 170}
            UITheme.StyleButton(btnGuardar, "Success")
            AddHandler btnGuardar.Click, AddressOf BtnGuardar_Click

            btnCancelar = New Button() With {.Text = "Cancelar", .Dock = DockStyle.Left, .Width = 100}
            UITheme.StyleButton(btnCancelar, "Secondary")
            AddHandler btnCancelar.Click, Sub() Me.Close()

            pnlBottom.Controls.AddRange({btnGuardar, btnCancelar})

            Me.Controls.AddRange({lblNom, txtNombre, lblApe, txtApellido,
                                   lblUser, txtUsername, lblRol, cboRol,
                                   lblTel, txtTelefono, lblEmail, txtEmail,
                                   lblDir, txtDireccion, lblCiu, txtCiudad,
                                   lblFnac, dtpFechaNac, chkTieneFechaNac,
                                   lblNotas, txtNotas, pnlBottom})
            Me.AcceptButton = btnGuardar
            Me.CancelButton = btnCancelar
        End Sub

        Private Sub LoadUserData()
            Dim u = usuarioService.GetUsuarioById(_usuarioId)
            If u IsNot Nothing Then
                txtNombre.Text = u.Nombre
                txtApellido.Text = u.Apellido
                txtUsername.Text = u.Username
                txtUsername.ReadOnly = True
                txtUsername.BackColor = Color.FromArgb(241, 245, 249)
                Dim idx = cboRol.FindStringExact(u.Rol)
                If idx >= 0 Then cboRol.SelectedIndex = idx
                txtTelefono.Text = u.Telefono
                txtEmail.Text = u.Email
                txtDireccion.Text = u.Direccion
                txtCiudad.Text = u.Ciudad
                txtNotas.Text = u.Notas
                If u.FechaNacimiento.HasValue Then
                    dtpFechaNac.Value = u.FechaNacimiento.Value
                    dtpFechaNac.Enabled = True
                    chkTieneFechaNac.Checked = False
                Else
                    chkTieneFechaNac.Checked = True
                End If
            End If
        End Sub

        Private Sub BtnGuardar_Click(sender As Object, e As EventArgs)
            Dim errMsg As String = ""
            Dim nom = txtNombre.Text.Trim()
            Dim ape = txtApellido.Text.Trim()
            Dim usr = txtUsername.Text.Trim()
            Dim rol = cboRol.SelectedItem.ToString()
            Dim fnac As Nullable(Of DateTime) = If(chkTieneFechaNac.Checked, Nothing, CType(dtpFechaNac.Value.Date, Nullable(Of DateTime)))

            If _usuarioId = 0 Then
                Dim pass = txtPassword.Text
                If usuarioService.CrearUsuario(usr, pass, nom, ape, rol, errMsg,
                                               txtTelefono.Text.Trim(), txtEmail.Text.Trim(),
                                               txtDireccion.Text.Trim(), txtCiudad.Text.Trim(),
                                               txtNotas.Text.Trim(), fnac) Then
                    MessageBox.Show($"Usuario '{usr}' creado exitosamente con rol {rol}.", "Usuario Creado", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Me.DialogResult = DialogResult.OK
                    Me.Close()
                Else
                    MessageBox.Show(errMsg, "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            Else
                If usuarioService.ActualizarUsuario(_usuarioId, nom, ape, rol, errMsg,
                                                    txtTelefono.Text.Trim(), txtEmail.Text.Trim(),
                                                    txtDireccion.Text.Trim(), txtCiudad.Text.Trim(),
                                                    txtNotas.Text.Trim(), fnac) Then
                    MessageBox.Show("Datos de usuario actualizados correctamente.", "Actualizado", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Me.DialogResult = DialogResult.OK
                    Me.Close()
                Else
                    MessageBox.Show(errMsg, "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            End If
        End Sub

    End Class

    ''' <summary>
    ''' Modal para Restablecer Contraseña de un Usuario
    ''' </summary>
    Public Class FrmCambiarPassword
        Inherits Form

        Private _usuarioId As Integer
        Private _username As String
        Private usuarioService As New UsuarioService()

        Private txtNueva As TextBox
        Private txtConfirmar As TextBox
        Private btnGuardar As Button
        Private btnCancelar As Button

        Public Sub New(usuarioId As Integer, username As String)
            _usuarioId = usuarioId
            _username = username
            InitializeUI()
        End Sub

        Private Sub InitializeUI()
            Me.Text = "Cambiar Contraseña de Usuario"
            Me.Size = New Size(400, 260)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.BackColor = Color.White
            Me.Font = UITheme.FontRegular

            Dim lblUser As New Label() With {
                .Text = $"Usuario: {_username}",
                .Font = UITheme.FontSubheading,
                .ForeColor = UITheme.ColorPrimaryDark,
                .Location = New Point(25, 18),
                .AutoSize = True
            }

            Dim lblN As New Label() With {.Text = "Nueva Contraseña (mín. 4 car.):", .Font = UITheme.FontBold, .Location = New Point(25, 55), .AutoSize = True}
            txtNueva = New TextBox() With {.Location = New Point(25, 78), .Size = New Size(335, 26), .UseSystemPasswordChar = True}
            UITheme.StyleTextBox(txtNueva)

            Dim lblC As New Label() With {.Text = "Confirmar Contraseña:", .Font = UITheme.FontBold, .Location = New Point(25, 115), .AutoSize = True}
            txtConfirmar = New TextBox() With {.Location = New Point(25, 138), .Size = New Size(335, 26), .UseSystemPasswordChar = True}
            UITheme.StyleTextBox(txtConfirmar)

            btnGuardar = New Button() With {.Text = "Actualizar Clave", .Location = New Point(155, 175), .Size = New Size(130, 34)}
            UITheme.StyleButton(btnGuardar, "Primary")
            AddHandler btnGuardar.Click, AddressOf BtnGuardar_Click

            btnCancelar = New Button() With {.Text = "Cancelar", .Location = New Point(295, 175), .Size = New Size(80, 34)}
            UITheme.StyleButton(btnCancelar, "Secondary")
            AddHandler btnCancelar.Click, Sub() Me.Close()

            Me.Controls.AddRange({lblUser, lblN, txtNueva, lblC, txtConfirmar, btnGuardar, btnCancelar})
            Me.AcceptButton = btnGuardar
            Me.CancelButton = btnCancelar
        End Sub

        Private Sub BtnGuardar_Click(sender As Object, e As EventArgs)
            If txtNueva.Text <> txtConfirmar.Text Then
                MessageBox.Show("Las contraseñas ingresadas no coinciden.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim errMsg As String = ""
            If usuarioService.CambiarPassword(_usuarioId, txtNueva.Text, errMsg) Then
                Me.DialogResult = DialogResult.OK
                Me.Close()
            Else
                MessageBox.Show(errMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End Sub

    End Class
End Namespace
