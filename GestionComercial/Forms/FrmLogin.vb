Imports System.Drawing
Imports System.Windows.Forms
Imports GestionComercial.Config
Imports GestionComercial.Data
Imports GestionComercial.Services
Imports GestionComercial.UI

Namespace Forms
    Public Class FrmLogin
        Inherits Form

        Private txtUser As TextBox
        Private txtPass As TextBox
        Private btnLogin As Button
        Private btnConfig As Button
        Private lblStatus As Label
        Private lblInfoCuentas As Label

        Public Sub New()
            InitializeUI()
            CheckInitialConnection()
        End Sub

        Private Sub InitializeUI()
            Me.Text = "Acceso al Sistema - Gestión Comercial de Indumentaria"
            Me.Size = New Size(450, 520)
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.BackColor = UITheme.ColorBackground
            Me.Font = UITheme.FontRegular
            Me.AutoScaleMode = AutoScaleMode.Dpi

            ' Banner superior
            Dim pnlBanner As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 130,
                .BackColor = UITheme.ColorSecondary,
                .Padding = New Padding(20)
            }

            Dim lblIcon As New Label() With {
                .Text = "👕",
                .Font = New Font("Segoe UI Emoji", 26.0F),
                .ForeColor = Color.White,
                .Location = New Point(25, 20),
                .AutoSize = True
            }

            Dim lblAppName As New Label() With {
                .Text = "SISTEMA DE GESTIÓN",
                .Font = New Font("Segoe UI", 14.0F, FontStyle.Bold),
                .ForeColor = Color.White,
                .Location = New Point(140, 25),
                .AutoSize = True
            }

            Dim lblAppSub As New Label() With {
                .Text = "Local de Ropa & Punto de Venta",
                .Font = New Font("Segoe UI", 9.5F, FontStyle.Regular),
                .ForeColor = Color.FromArgb(148, 163, 184),
                .Location = New Point(142, 55),
                .AutoSize = True
            }

            pnlBanner.Controls.AddRange({lblIcon, lblAppName, lblAppSub})
            Me.Controls.Add(pnlBanner)

            ' Tarjeta central de login
            Dim pnlCard As New Panel() With {
                .Location = New Point(35, 150),
                .Size = New Size(365, 300),
                .BackColor = UITheme.ColorSurface,
                .Padding = New Padding(20)
            }

            Dim lblUser As New Label() With {
                .Text = "Usuario:",
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorTextPrimary,
                .Location = New Point(20, 20),
                .AutoSize = True
            }
            txtUser = New TextBox() With {
                .Location = New Point(20, 45),
                .Size = New Size(325, 28),
                .Text = "admin"
            }
            UITheme.StyleTextBox(txtUser)

            Dim lblPass As New Label() With {
                .Text = "Contraseña:",
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorTextPrimary,
                .Location = New Point(20, 85),
                .AutoSize = True
            }
            txtPass = New TextBox() With {
                .Location = New Point(20, 110),
                .Size = New Size(325, 28),
                .PasswordChar = "●"c,
                .Text = "admin123"
            }
            UITheme.StyleTextBox(txtPass)

            btnLogin = New Button() With {
                .Text = "INGRESAR AL SISTEMA",
                .Location = New Point(20, 155),
                .Size = New Size(325, 42)
            }
            UITheme.StyleButton(btnLogin, "Primary")
            AddHandler btnLogin.Click, AddressOf BtnLogin_Click

            lblStatus = New Label() With {
                .Text = "Verificando base de datos...",
                .Font = UITheme.FontSmall,
                .ForeColor = UITheme.ColorTextSecondary,
                .Location = New Point(20, 205),
                .Size = New Size(325, 20),
                .TextAlign = ContentAlignment.MiddleCenter
            }

            lblInfoCuentas = New Label() With {
                .Text = "Accesos por defecto: admin / admin123 | vendedor / 1234",
                .Font = UITheme.FontSmall,
                .ForeColor = UITheme.ColorTextSecondary,
                .Location = New Point(20, 230),
                .Size = New Size(325, 30),
                .TextAlign = ContentAlignment.MiddleCenter
            }

            btnConfig = New Button() With {
                .Text = "⚙ Configuración de Base de Datos",
                .Location = New Point(20, 265),
                .Size = New Size(325, 26)
            }
            UITheme.StyleButton(btnConfig, "Secondary")
            btnConfig.Font = UITheme.FontSmall
            AddHandler btnConfig.Click, AddressOf BtnConfig_Click

            pnlCard.Controls.AddRange({lblUser, txtUser, lblPass, txtPass, btnLogin, lblStatus, lblInfoCuentas, btnConfig})
            Me.Controls.Add(pnlCard)

            ' Atajos de teclado
            Me.AcceptButton = btnLogin
        End Sub

        Private Sub CheckInitialConnection()
            Dim errMsg As String = ""
            Dim ok As Boolean = DatabaseHelper.TestConnection(errMsg)
            If ok Then
                If DatabaseHelper.IsSQLite Then
                    lblStatus.Text = "● Base de Datos SQLite Lista"
                Else
                    lblStatus.Text = "● MySQL Conectado (" & AppConfig.Settings.Database & ")"
                End If
                lblStatus.ForeColor = UITheme.ColorSuccess
            Else
                If DatabaseHelper.IsSQLite Then
                    lblStatus.Text = "⚠ Error SQLite: " & errMsg
                Else
                    lblStatus.Text = "⚠ Sin conexión a MySQL. Revisa Configuración"
                End If
                lblStatus.ForeColor = UITheme.ColorDanger
            End If
        End Sub

        Private Sub BtnLogin_Click(sender As Object, e As EventArgs)
            Dim user As String = txtUser.Text.Trim()
            Dim pass As String = txtPass.Text

            Dim errMsg As String = ""
            If AuthService.Login(user, pass, errMsg) Then
                ' Acceso concedido
                Me.Hide()
                Dim frm As New FrmMain()
                AddHandler frm.FormClosed, Sub() Me.Close()
                frm.Show()
            Else
                ' Si falló porque no existe la base de datos o conexión, ofrecer inicializar
                If errMsg.Contains("Unable to connect", StringComparison.OrdinalIgnoreCase) OrElse errMsg.Contains("Unknown database", StringComparison.OrdinalIgnoreCase) OrElse errMsg.Contains("no such table", StringComparison.OrdinalIgnoreCase) Then
                    Dim resp = MessageBox.Show("No se pudo conectar con la base de datos o faltan tablas." & vbCrLf & vbCrLf & "¿Deseas abrir la configuración para inicializar las tablas automáticamente?", "Base de Datos Requerida", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    If resp = DialogResult.Yes Then
                        BtnConfig_Click(Nothing, Nothing)
                    End If
                Else
                    MessageBox.Show(errMsg, "Error de Inicio de Sesión", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            End If
        End Sub

        Private Sub BtnConfig_Click(sender As Object, e As EventArgs)
            Dim frm As New FrmConfiguracion()
            frm.ShowDialog()
            CheckInitialConnection()
        End Sub

    End Class
End Namespace
