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
                .Text = "Gestión de usuarios y vendedores",
                .Font = UITheme.FontHeading,
                .ForeColor = Color.White,
                .AutoSize = True,
                .Location = New Point(15, 15)
            }
            pnlHeader.Controls.Add(lblTitle)

            ' Toolbar
            Dim pnlToolbar As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 60,
                .BackColor = UITheme.ColorSurface,
                .Padding = New Padding(15, 12, 15, 10)
            }

            btnNuevo = New Button() With {.Text = "+ Nuevo Usuario", .Location = New Point(15, 12), .Size = New Size(150, 34)}
            UITheme.StyleButton(btnNuevo, "Success")
            AddHandler btnNuevo.Click, AddressOf BtnNuevo_Click

            btnEditar = New Button() With {.Text = "Editar datos", .Location = New Point(175, 12), .Size = New Size(130, 34)}
            UITheme.StyleButton(btnEditar, "Secondary")
            AddHandler btnEditar.Click, AddressOf BtnEditar_Click

            btnCambiarPass = New Button() With {.Text = "Cambiar clave", .Location = New Point(315, 12), .Size = New Size(150, 34)}
            UITheme.StyleButton(btnCambiarPass, "Secondary")
            AddHandler btnCambiarPass.Click, AddressOf BtnCambiarPass_Click

            btnToggleActivo = New Button() With {.Text = "Activar / desactivar", .Location = New Point(475, 12), .Size = New Size(180, 34)}
            UITheme.StyleButton(btnToggleActivo, "Danger")
            AddHandler btnToggleActivo.Click, AddressOf BtnToggleActivo_Click

            btnRefrescar = New Button() With {.Text = "Actualizar", .Location = New Point(665, 12), .Size = New Size(120, 34)}
            UITheme.StyleButton(btnRefrescar, "Secondary")
            AddHandler btnRefrescar.Click, Sub() LoadUsuarios()

            Dim puedeModificar As Boolean = AuthService.IsAdmin
            btnNuevo.Visible = puedeModificar
            btnEditar.Visible = puedeModificar
            btnCambiarPass.Visible = puedeModificar
            btnToggleActivo.Visible = puedeModificar

            pnlToolbar.Controls.AddRange({btnNuevo, btnEditar, btnCambiarPass, btnToggleActivo, btnRefrescar})

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
                .BackColor = UITheme.ColorSurfaceMuted,
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
            dgvUsuarios.Columns("Id").Width = 60

            dgvUsuarios.Columns.Add("Username", "Usuario")
            dgvUsuarios.Columns("Username").Width = 140

            dgvUsuarios.Columns.Add("NombreCompleto", "Nombre y Apellido")
            dgvUsuarios.Columns("NombreCompleto").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

            dgvUsuarios.Columns.Add("Rol", "Rol")
            dgvUsuarios.Columns("Rol").Width = 140

            dgvUsuarios.Columns.Add("Activo", "Estado")
            dgvUsuarios.Columns("Activo").Width = 120

            dgvUsuarios.Columns.Add("UltimoLogin", "Último Acceso")
            dgvUsuarios.Columns("UltimoLogin").Width = 160

            dgvUsuarios.Columns.Add("CreatedAt", "Fecha de Alta")
            dgvUsuarios.Columns("CreatedAt").Width = 140
        End Sub

        Public Sub LoadUsuarios()
            Dim lista = usuarioService.GetUsuarios()
            dgvUsuarios.Rows.Clear()

            Dim activosCount As Integer = 0
            For Each u In lista
                Dim ultAcceso = If(u.UltimoLogin.HasValue, u.UltimoLogin.Value.ToString("dd/MM/yyyy HH:mm"), "Nunca")
                Dim fechaAlta = If(u.CreatedAt > DateTime.MinValue, u.CreatedAt.ToString("dd/MM/yyyy"), "-")
                Dim estadoStr = If(u.Activo, "Activo", "Inactivo")
                If u.Activo Then activosCount += 1

                dgvUsuarios.Rows.Add(u.Id, u.Username, u.NombreCompleto, u.Rol, estadoStr, ultAcceso, fechaAlta)
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
            If AuthService.IsAdmin AndAlso e.RowIndex >= 0 Then
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

        Private txtNombreCompleto As TextBox
        Private txtUsername As TextBox
        Private cboRol As ComboBox
        Private txtPassword As TextBox
        Private lblPassTitle As Label
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
            Me.Text = If(_usuarioId = 0, "Crear Nuevo Usuario", "Editar Datos de Usuario")
            Me.Size = New Size(440, If(_usuarioId = 0, 360, 300))
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.BackColor = Color.White
            Me.Font = UITheme.FontRegular

            Dim lblNom As New Label() With {.Text = "Nombre Completo del Empleado:", .Font = UITheme.FontBold, .Location = New Point(25, 20), .AutoSize = True}
            txtNombreCompleto = New TextBox() With {.Location = New Point(25, 45), .Size = New Size(375, 26)}
            UITheme.StyleTextBox(txtNombreCompleto)

            Dim lblUser As New Label() With {.Text = "Nombre de Usuario (Login):", .Font = UITheme.FontBold, .Location = New Point(25, 85), .AutoSize = True}
            txtUsername = New TextBox() With {.Location = New Point(25, 110), .Size = New Size(375, 26)}
            UITheme.StyleTextBox(txtUsername)

            Dim lblRol As New Label() With {.Text = "Rol en el Sistema:", .Font = UITheme.FontBold, .Location = New Point(25, 150), .AutoSize = True}
            cboRol = New ComboBox() With {.Location = New Point(25, 175), .Size = New Size(375, 26), .DropDownStyle = ComboBoxStyle.DropDownList}
            UITheme.StyleComboBox(cboRol)
            cboRol.Items.AddRange({"Vendedor", "Gerente", "Administrador"})
            cboRol.SelectedIndex = 0

            Dim nextY As Integer = 215
            If _usuarioId = 0 Then
                lblPassTitle = New Label() With {.Text = "Contraseña Inicial (mín. 4 caracteres):", .Font = UITheme.FontBold, .Location = New Point(25, nextY), .AutoSize = True}
                txtPassword = New TextBox() With {.Location = New Point(25, nextY + 25), .Size = New Size(375, 26), .UseSystemPasswordChar = True}
                UITheme.StyleTextBox(txtPassword)
                Me.Controls.AddRange({lblPassTitle, txtPassword})
                nextY += 65
            End If

            btnGuardar = New Button() With {.Text = "Guardar Usuario", .Location = New Point(180, nextY), .Size = New Size(130, 36)}
            UITheme.StyleButton(btnGuardar, "Success")
            AddHandler btnGuardar.Click, AddressOf BtnGuardar_Click

            btnCancelar = New Button() With {.Text = "Cancelar", .Location = New Point(320, nextY), .Size = New Size(80, 36)}
            UITheme.StyleButton(btnCancelar, "Secondary")
            AddHandler btnCancelar.Click, Sub() Me.Close()

            Me.Controls.AddRange({lblNom, txtNombreCompleto, lblUser, txtUsername, lblRol, cboRol, btnGuardar, btnCancelar})
            Me.AcceptButton = btnGuardar
            Me.CancelButton = btnCancelar
        End Sub

        Private Sub LoadUserData()
            Dim u = usuarioService.GetUsuarioById(_usuarioId)
            If u IsNot Nothing Then
                txtNombreCompleto.Text = u.NombreCompleto
                txtUsername.Text = u.Username
                txtUsername.ReadOnly = True
                txtUsername.BackColor = UITheme.ColorSurfaceMuted
                Dim idx = cboRol.FindStringExact(u.Rol)
                If idx >= 0 Then cboRol.SelectedIndex = idx
            End If
        End Sub

        Private Sub BtnGuardar_Click(sender As Object, e As EventArgs)
            Dim errMsg As String = ""
            Dim nom = txtNombreCompleto.Text.Trim()
            Dim usr = txtUsername.Text.Trim()
            Dim rol = cboRol.SelectedItem.ToString()

            If _usuarioId = 0 Then
                Dim pass = txtPassword.Text
                If usuarioService.CrearUsuario(usr, pass, nom, rol, errMsg) Then
                    MessageBox.Show($"Usuario '{usr}' creado exitosamente con rol {rol}.", "Usuario Creado", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Me.DialogResult = DialogResult.OK
                    Me.Close()
                Else
                    MessageBox.Show(errMsg, "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            Else
                If usuarioService.ActualizarUsuario(_usuarioId, nom, rol, errMsg) Then
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

            btnGuardar = New Button() With {.Text = "Actualizar Clave", .Location = New Point(170, 175), .Size = New Size(115, 34)}
            UITheme.StyleButton(btnGuardar, "Primary")
            AddHandler btnGuardar.Click, AddressOf BtnGuardar_Click

            btnCancelar = New Button() With {.Text = "Cancelar", .Location = New Point(290, 175), .Size = New Size(70, 34)}
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
