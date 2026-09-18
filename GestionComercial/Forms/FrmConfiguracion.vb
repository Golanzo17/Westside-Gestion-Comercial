Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports GestionComercial.Config
Imports GestionComercial.Data
Imports GestionComercial.Models
Imports GestionComercial.Services
Imports GestionComercial.UI

Namespace Forms
    Public Class FrmConfiguracion
        Inherits Form

        Private cboProvider As ComboBox
        Private txtSqlitePath As TextBox
        Private txtHost As TextBox
        Private numPort As NumericUpDown
        Private txtDatabase As TextBox
        Private txtUser As TextBox
        Private txtPassword As TextBox
        Private btnProbar As Button
        Private btnInicializarDb As Button

        Private txtNombreLocal As TextBox
        Private txtCuit As TextBox
        Private txtDireccion As TextBox
        Private txtTelefono As TextBox
        Private txtEmail As TextBox
        Private txtTicketMsg As TextBox
        Private cboIva As ComboBox

        Private btnGuardarTodo As Button
        Private btnCerrar As Button
        Private lblEstadoConexion As Label

        Private configService As New ConfiguracionService()

        Public Sub New()
            InitializeUI()
            LoadConfigData()
        End Sub

        Private Sub InitializeUI()
            Me.Text = "Configuración del Sistema y Base de Datos"
            Me.Size = New Size(760, 720)
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.BackColor = UITheme.ColorBackground
            Me.Font = UITheme.FontRegular

            Dim pnlHeader As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 70,
                .BackColor = UITheme.ColorSecondary,
                .Padding = New Padding(20, 15, 20, 15)
            }
            Dim lblTitle As New Label() With {
                .Text = "Configuración general y base de datos",
                .Font = UITheme.FontHeading,
                .ForeColor = Color.White,
                .AutoSize = True,
                .Location = New Point(20, 15)
            }
            Dim lblSub As New Label() With {
                .Text = "Motor de persistencia (SQLite / MySQL) y datos comerciales para tickets",
                .Font = UITheme.FontSmall,
                .ForeColor = UITheme.ColorTextMuted,
                .AutoSize = True,
                .Location = New Point(22, 45)
            }
            pnlHeader.Controls.AddRange({lblTitle, lblSub})
            Me.Controls.Add(pnlHeader)

            ' Grupo Base de Datos
            Dim grpDb As New GroupBox() With {
                .Text = "Motor y Conexión de Base de Datos",
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorPrimaryDark,
                .Location = New Point(20, 80),
                .Size = New Size(705, 245),
                .BackColor = UITheme.ColorSurface
            }

            Dim lblProv As New Label() With {.Text = "Motor de Base de Datos:", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextPrimary, .Location = New Point(20, 26), .AutoSize = True}
            cboProvider = New ComboBox() With {.Location = New Point(20, 48), .Size = New Size(320, 26), .DropDownStyle = ComboBoxStyle.DropDownList}
            UITheme.StyleComboBox(cboProvider)
            cboProvider.Items.AddRange({"SQLite (Local sin servidor - Recomendado)", "MySQL (Servidor externo)"})
            AddHandler cboProvider.SelectedIndexChanged, AddressOf CboProvider_SelectedIndexChanged

            Dim lblSqPath As New Label() With {.Text = "Ubicación del Archivo SQLite:", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextPrimary, .Location = New Point(360, 26), .AutoSize = True}
            txtSqlitePath = New TextBox() With {.Location = New Point(360, 48), .Size = New Size(325, 26), .ReadOnly = True}
            UITheme.StyleTextBox(txtSqlitePath)

            ' Servidor / Host
            Dim lblH As New Label() With {.Text = "Servidor / Host (MySQL):", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextPrimary, .Location = New Point(20, 85), .AutoSize = True}
            txtHost = New TextBox() With {.Location = New Point(20, 107), .Size = New Size(220, 26)}
            UITheme.StyleTextBox(txtHost)

            ' Puerto
            Dim lblP As New Label() With {.Text = "Puerto (default 3306):", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextPrimary, .Location = New Point(260, 85), .AutoSize = True}
            numPort = New NumericUpDown() With {.Location = New Point(260, 107), .Size = New Size(100, 26), .Minimum = 1, .Maximum = 65535, .Value = 3306}

            ' Base de datos
            Dim lblD As New Label() With {.Text = "Nombre Base de Datos:", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextPrimary, .Location = New Point(380, 85), .AutoSize = True}
            txtDatabase = New TextBox() With {.Location = New Point(380, 107), .Size = New Size(305, 26)}
            UITheme.StyleTextBox(txtDatabase)

            ' Usuario
            Dim lblU As New Label() With {.Text = "Usuario MySQL:", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextPrimary, .Location = New Point(20, 142), .AutoSize = True}
            txtUser = New TextBox() With {.Location = New Point(20, 164), .Size = New Size(220, 26)}
            UITheme.StyleTextBox(txtUser)

            ' Contraseña
            Dim lblPw As New Label() With {.Text = "Contraseña:", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextPrimary, .Location = New Point(260, 142), .AutoSize = True}
            txtPassword = New TextBox() With {.Location = New Point(260, 164), .Size = New Size(220, 26), .PasswordChar = "*"c}
            UITheme.StyleTextBox(txtPassword)

            ' Botones de prueba y creación
            btnProbar = New Button() With {.Text = "Probar conexión", .Location = New Point(20, 200), .Size = New Size(170, 34)}
            UITheme.StyleButton(btnProbar, "Secondary")
            AddHandler btnProbar.Click, AddressOf BtnProbar_Click

            btnInicializarDb = New Button() With {.Text = "Inicializar tablas y datos", .Location = New Point(200, 200), .Size = New Size(220, 34)}
            UITheme.StyleButton(btnInicializarDb, "Primary")
            AddHandler btnInicializarDb.Click, AddressOf BtnInicializarDb_Click

            lblEstadoConexion = New Label() With {
                .Text = "Estado: Verificando...",
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorTextSecondary,
                .Location = New Point(430, 207),
                .AutoSize = True
            }

            grpDb.Controls.AddRange({lblProv, cboProvider, lblSqPath, txtSqlitePath, lblH, txtHost, lblP, numPort, lblD, txtDatabase, lblU, txtUser, lblPw, txtPassword, btnProbar, btnInicializarDb, lblEstadoConexion})
            Me.Controls.Add(grpDb)

            ' Grupo Comercio
            Dim grpLocal As New GroupBox() With {
                .Text = "Datos del Comercio / Local de Ropa (Para Tickets y Comprobantes)",
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorPrimaryDark,
                .Location = New Point(20, 335),
                .Size = New Size(705, 260),
                .BackColor = UITheme.ColorSurface
            }

            Dim lblNL As New Label() With {.Text = "Nombre del Local / Marca:", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextPrimary, .Location = New Point(20, 30), .AutoSize = True}
            txtNombreLocal = New TextBox() With {.Location = New Point(20, 52), .Size = New Size(320, 26)}
            UITheme.StyleTextBox(txtNombreLocal)

            Dim lblCuit As New Label() With {.Text = "CUIT / CUIL:", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextPrimary, .Location = New Point(360, 30), .AutoSize = True}
            txtCuit = New TextBox() With {.Location = New Point(360, 52), .Size = New Size(160, 26)}
            UITheme.StyleTextBox(txtCuit)

            Dim lblIva As New Label() With {.Text = "Condición IVA:", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextPrimary, .Location = New Point(535, 30), .AutoSize = True}
            cboIva = New ComboBox() With {.Location = New Point(535, 52), .Size = New Size(150, 26), .DropDownStyle = ComboBoxStyle.DropDownList}
            UITheme.StyleComboBox(cboIva)
            cboIva.Items.AddRange({"Responsable Inscripto", "Monotributo", "Exento", "Consumidor Final"})
            cboIva.SelectedIndex = 0

            Dim lblDir As New Label() With {.Text = "Dirección Comercial:", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextPrimary, .Location = New Point(20, 90), .AutoSize = True}
            txtDireccion = New TextBox() With {.Location = New Point(20, 112), .Size = New Size(320, 26)}
            UITheme.StyleTextBox(txtDireccion)

            Dim lblTel As New Label() With {.Text = "Teléfono / WhatsApp:", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextPrimary, .Location = New Point(360, 90), .AutoSize = True}
            txtTelefono = New TextBox() With {.Location = New Point(360, 112), .Size = New Size(160, 26)}
            UITheme.StyleTextBox(txtTelefono)

            Dim lblEmail As New Label() With {.Text = "Email de Contacto:", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextPrimary, .Location = New Point(535, 90), .AutoSize = True}
            txtEmail = New TextBox() With {.Location = New Point(535, 112), .Size = New Size(150, 26)}
            UITheme.StyleTextBox(txtEmail)

            Dim lblMsg As New Label() With {.Text = "Leyenda / Política de Cambios en Pie de Ticket:", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextPrimary, .Location = New Point(20, 150), .AutoSize = True}
            txtTicketMsg = New TextBox() With {.Location = New Point(20, 172), .Size = New Size(665, 26)}
            UITheme.StyleTextBox(txtTicketMsg)

            grpLocal.Controls.AddRange({lblNL, txtNombreLocal, lblCuit, txtCuit, lblIva, cboIva, lblDir, txtDireccion, lblTel, txtTelefono, lblEmail, txtEmail, lblMsg, txtTicketMsg})
            Me.Controls.Add(grpLocal)

            ' Botones inferiores
            Dim pnlBottom As New Panel() With {
                .Dock = DockStyle.Bottom,
                .Height = 65,
                .BackColor = UITheme.ColorSurfaceMuted,
                .Padding = New Padding(20, 12, 20, 12)
            }

            btnGuardarTodo = New Button() With {
                .Text = "Guardar configuración",
                .Dock = DockStyle.Right,
                .Width = 200
            }
            UITheme.StyleButton(btnGuardarTodo, "Success")
            AddHandler btnGuardarTodo.Click, AddressOf BtnGuardarTodo_Click

            Dim btnIrUsuarios As New Button() With {
                .Text = "Usuarios y empleados",
                .Dock = DockStyle.Right,
                .Width = 190
            }
            UITheme.StyleButton(btnIrUsuarios, "Primary")
            AddHandler btnIrUsuarios.Click, Sub()
                                                Dim frmU As New FrmUsuarios()
                                                frmU.ShowDialog(Me)
                                            End Sub

            btnCerrar = New Button() With {
                .Text = "Cerrar",
                .Dock = DockStyle.Left,
                .Width = 100
            }
            UITheme.StyleButton(btnCerrar, "Secondary")
            AddHandler btnCerrar.Click, Sub() Me.Close()

            pnlBottom.Controls.AddRange({btnGuardarTodo, btnIrUsuarios, btnCerrar})
            Me.Controls.Add(pnlBottom)
        End Sub

        Private Sub CboProvider_SelectedIndexChanged(sender As Object, e As EventArgs)
            Dim isSqlite = (cboProvider.SelectedIndex = 0)
            txtHost.Enabled = Not isSqlite
            numPort.Enabled = Not isSqlite
            txtDatabase.Enabled = Not isSqlite
            txtUser.Enabled = Not isSqlite
            txtPassword.Enabled = Not isSqlite
            txtSqlitePath.Enabled = isSqlite
        End Sub

        Private Sub LoadConfigData()
            ' Cargar proveedor
            If AppConfig.Settings.Provider.Equals("MySQL", StringComparison.OrdinalIgnoreCase) Then
                cboProvider.SelectedIndex = 1
            Else
                cboProvider.SelectedIndex = 0
            End If

            txtSqlitePath.Text = AppConfig.Settings.GetSqlitePath()

            ' Cargar MySQL settings
            txtHost.Text = AppConfig.Settings.Host
            numPort.Value = AppConfig.Settings.Port
            txtDatabase.Text = AppConfig.Settings.Database
            txtUser.Text = AppConfig.Settings.Username
            txtPassword.Text = AppConfig.Settings.Password

            ' Intentar cargar datos del local si hay conexión
            Try
                Dim cfg = configService.GetConfiguracion()
                txtNombreLocal.Text = cfg.NombreComercio
                txtCuit.Text = cfg.Cuit
                txtDireccion.Text = cfg.Direccion
                txtTelefono.Text = cfg.Telefono
                txtEmail.Text = cfg.Email
                txtTicketMsg.Text = cfg.MensajeTicket

                Dim ivaIdx = cboIva.FindStringExact(cfg.CondicionIva)
                If ivaIdx >= 0 Then cboIva.SelectedIndex = ivaIdx
            Catch
                ' Si no hay base de datos aún, dejamos defaults
            End Try

            ' Comprobar estado actual de conexión
            Dim errMsg As String = ""
            If DatabaseHelper.TestConnection(errMsg) Then
                lblEstadoConexion.Text = "✔ Conectado a " & If(DatabaseHelper.IsSQLite, "SQLite", "MySQL")
                lblEstadoConexion.ForeColor = UITheme.ColorSuccess
            Else
                lblEstadoConexion.Text = "⚠ Sin conexión inicial"
                lblEstadoConexion.ForeColor = UITheme.ColorDanger
            End If
        End Sub

        Private Sub BtnProbar_Click(sender As Object, e As EventArgs)
            ApplyDbFormToSettings()

            Dim errMsg As String = ""
            Dim ok As Boolean = DatabaseHelper.TestConnection(errMsg)
            Dim motor = If(cboProvider.SelectedIndex = 0, "SQLite", "MySQL")

            If ok Then
                lblEstadoConexion.Text = "✔ Conexión Exitosa (" & motor & ")"
                lblEstadoConexion.ForeColor = UITheme.ColorSuccess
                MessageBox.Show($"¡Conexión establecida exitosamente con el motor {motor}!", "Conexión Correcta", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                lblEstadoConexion.Text = "✖ Error de conexión"
                lblEstadoConexion.ForeColor = UITheme.ColorDanger
                MessageBox.Show($"No se pudo conectar con {motor}." & vbCrLf & vbCrLf & "Detalle: " & errMsg, "Fallo de Conexión", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End Sub

        Private Sub BtnInicializarDb_Click(sender As Object, e As EventArgs)
            ApplyDbFormToSettings()
            AppConfig.SaveSettings()

            Dim outMsg As String = ""
            Dim ok As Boolean = DatabaseHelper.InitializeDatabaseAndTables(outMsg)
            If ok Then
                lblEstadoConexion.Text = "✔ Tablas y Datos Listos"
                lblEstadoConexion.ForeColor = UITheme.ColorSuccess
                MessageBox.Show(outMsg, "Estructura Lista", MessageBoxButtons.OK, MessageBoxIcon.Information)
                LoadConfigData()
            Else
                MessageBox.Show(outMsg, "Error de Inicialización", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End Sub

        Private Sub BtnGuardarTodo_Click(sender As Object, e As EventArgs)
            ApplyDbFormToSettings()
            AppConfig.SaveSettings()

            ' Guardar info del comercio
            Dim cfg As New ConfiguracionComercio() With {
                .NombreComercio = txtNombreLocal.Text.Trim(),
                .Cuit = txtCuit.Text.Trim(),
                .CondicionIva = If(cboIva.SelectedItem IsNot Nothing, cboIva.SelectedItem.ToString(), "Responsable Inscripto"),
                .Direccion = txtDireccion.Text.Trim(),
                .Telefono = txtTelefono.Text.Trim(),
                .Email = txtEmail.Text.Trim(),
                .MensajeTicket = txtTicketMsg.Text.Trim(),
                .MonedaSimbolo = "$"
            }

            Dim errMsg As String = ""
            If configService.GuardarConfiguracion(cfg, errMsg) Then
                MessageBox.Show("Configuración guardada correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.DialogResult = DialogResult.OK
                Me.Close()
            Else
                MessageBox.Show("Se guardaron los parámetros de base de datos, pero ocurrió un aviso con los datos del comercio: " & errMsg, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End Sub

        Private Sub ApplyDbFormToSettings()
            AppConfig.Settings.Provider = If(cboProvider.SelectedIndex = 0, "SQLite", "MySQL")
            AppConfig.Settings.Host = txtHost.Text.Trim()
            AppConfig.Settings.Port = Convert.ToInt32(numPort.Value)
            AppConfig.Settings.Database = txtDatabase.Text.Trim()
            AppConfig.Settings.Username = txtUser.Text.Trim()
            AppConfig.Settings.Password = txtPassword.Text
        End Sub

    End Class
End Namespace
