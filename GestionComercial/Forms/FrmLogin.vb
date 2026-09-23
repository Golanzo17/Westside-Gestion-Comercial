' ARCHIVO: FrmLogin.vb
' PROPÓSITO: Formulario de inicio de sesión y autenticación de usuarios.
' Es la primera ventana que ve el usuario al ejecutar el sistema:
' 1. Construcción 100% por Código: Paneles, etiquetas, cajas de texto y botones en InitializeUI() aplicando los tokens de UITheme.vb.
' 2. Chequeo de Conexión en Tiempo Real: En el evento de carga, comprueba si la base de datos responde correctamente y muestra un indicador verde/rojo.
' 3. Transición Segura a FrmMain: Si el login es exitoso, oculta el formulario de login y abre FrmMain.
'    Vincula el evento FormClosed de FrmMain para que al salir se cierre el proceso en Windows.


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
        Private lblStatus As Label
        Private lblInfoCuentas As Label

        Public Sub New()
            InitializeUI()
            ' Verificamos la base de datos antes de que el usuario intente ingresar
            CheckInitialConnection()
        End Sub

        Private Sub InitializeUI()
            Me.Text = "Acceso al Sistema - Gestión Comercial de Indumentaria"
            Me.Size = New Size(450, 475)
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
                .Text = "GC",
                .Font = New Font("Segoe UI", 20.0F, FontStyle.Bold),
                .ForeColor = Color.White,
                .Location = New Point(25, 20),
                .Size = New Size(78, 42),
                .TextAlign = ContentAlignment.MiddleCenter,
                .BackColor = UITheme.ColorPrimaryDark
            }

            Dim lblAppName As New Label() With {
                .Text = "SISTEMA DE GESTIÓN",
                .Font = New Font("Segoe UI", 14.0F, FontStyle.Bold),
                .ForeColor = Color.White,
                .Location = New Point(125, 25),
                .AutoSize = True
            }

            Dim lblAppSub As New Label() With {
                .Text = "Local de Ropa & Punto de Venta",
                .Font = New Font("Segoe UI", 9.5F, FontStyle.Regular),
                .ForeColor = UITheme.ColorTextMuted,
                .Location = New Point(127, 55),
                .AutoSize = True
            }

            pnlBanner.Controls.AddRange({lblIcon, lblAppName, lblAppSub})
            Me.Controls.Add(pnlBanner)

            ' Tarjeta central de login
            Dim pnlCard As New Panel() With {
                .Location = New Point(35, 150),
                .Size = New Size(365, 260),
                .BackColor = UITheme.ColorSurface,
                .Padding = New Padding(20)
            }
            UITheme.ApplyRoundedRegion(pnlCard, 10)

            Dim lblUser As New Label() With {
                .Text = "Usuario:",
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorTextPrimary,
                .Location = New Point(20, 20),
                .AutoSize = True
            }
            txtUser = New TextBox() With {
                .Location = New Point(20, 45),
                .Size = New Size(325, 28)
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
                .PasswordChar = "●"c
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
                .Text = "Ingrese sus credenciales para continuar",
                .Font = UITheme.FontSmall,
                .ForeColor = UITheme.ColorTextSecondary,
                .Location = New Point(20, 228),
                .Size = New Size(325, 24),
                .TextAlign = ContentAlignment.MiddleCenter
            }

            pnlCard.Controls.AddRange({lblUser, txtUser, lblPass, txtPass, btnLogin, lblStatus, lblInfoCuentas})
            Me.Controls.Add(pnlCard)

            ' Atajos de teclado
            Me.AcceptButton = btnLogin
        End Sub

        Private Sub CheckInitialConnection()
            Dim errMsg As String = ""
            Dim ok As Boolean = DatabaseHelper.TestConnection(errMsg)
            If ok Then
                lblStatus.Text = "● Base de Datos SQLite Lista"
                lblStatus.ForeColor = UITheme.ColorSuccess
            Else
                lblStatus.Text = "⚠ Error Base de Datos: " & errMsg
                lblStatus.ForeColor = UITheme.ColorDanger
            End If
        End Sub

        Private Sub BtnLogin_Click(sender As Object, e As EventArgs)
            Dim user As String = txtUser.Text.Trim()
            Dim pass As String = txtPass.Text

            If String.IsNullOrWhiteSpace(user) OrElse String.IsNullOrWhiteSpace(pass) Then
                MessageBox.Show("Por favor ingrese su usuario y contraseña.", "Campos Requeridos", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim errMsg As String = ""
            If AuthService.Login(user, pass, errMsg) Then
                ' Acceso concedido
                Me.Hide()
                Dim frm As New FrmMain()
                AddHandler frm.FormClosed, Sub() Me.Close()
                frm.Show()
            Else
                If errMsg.Contains("Unable to connect", StringComparison.OrdinalIgnoreCase) OrElse errMsg.Contains("Unknown database", StringComparison.OrdinalIgnoreCase) OrElse errMsg.Contains("no such table", StringComparison.OrdinalIgnoreCase) Then
                    MessageBox.Show("No se pudo conectar con la base de datos o el servicio no está disponible." & vbCrLf & vbCrLf & "Por favor verifique que la base de datos esté activa o contacte al administrador del sistema.", "Error de Conexión", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    MessageBox.Show(errMsg, "Error de Inicio de Sesión", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            End If
        End Sub

    End Class
End Namespace
