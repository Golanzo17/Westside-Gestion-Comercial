Imports System.Drawing
Imports System.Windows.Forms
Imports GestionComercial.Services
Imports GestionComercial.UI

Namespace Forms
    Public Class FrmSincronizacionEcommerce
        Inherits Form

        Private syncService As New EcommerceSyncService()

        Private txtHost As TextBox
        Private numPort As NumericUpDown
        Private txtDbName As TextBox
        Private txtUser As TextBox
        Private txtPassword As TextBox
        Private btnProbar As Button
        Private btnImportar As Button

        Private lblCatCount As Label
        Private lblTallesCount As Label
        Private lblProdCount As Label
        Private lblStockCount As Label
        Private txtLog As TextBox

        Public Sub New()
            InitializeUI()
        End Sub

        Private Sub InitializeUI()
            Me.Text = "Migración y Sincronización con E-commerce Web"
            Me.Size = New Size(780, 620)
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.BackColor = UITheme.ColorBackground
            Me.Font = UITheme.FontRegular

            ' Header
            Dim pnlHeader As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 65,
                .BackColor = UITheme.ColorSecondary,
                .Padding = New Padding(20, 15, 20, 15)
            }
            Dim lblTitle As New Label() With {
                .Text = "Migración y vinculación con la tienda web",
                .Font = UITheme.FontHeading,
                .ForeColor = Color.White,
                .AutoSize = True,
                .Location = New Point(15, 12)
            }
            Dim lblSub As New Label() With {
                .Text = "Importe el catálogo, categorías, talles y stock creados previamente en su página web",
                .Font = UITheme.FontSmall,
                .ForeColor = UITheme.ColorTextMuted,
                .AutoSize = True,
                .Location = New Point(17, 38)
            }
            pnlHeader.Controls.AddRange({lblTitle, lblSub})
            Me.Controls.Add(pnlHeader)

            ' Parámetros de conexión a la BD del Ecommerce
            Dim grpWebDb As New GroupBox() With {
                .Text = "Base de Datos de la Tienda Web (E-commerce)",
                .Location = New Point(20, 80),
                .Size = New Size(725, 165),
                .BackColor = UITheme.ColorSurface,
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorPrimaryDark
            }

            Dim lblH As New Label() With {.Text = "Servidor / Host:", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextPrimary, .Location = New Point(20, 25), .AutoSize = True}
            txtHost = New TextBox() With {.Location = New Point(20, 48), .Size = New Size(180, 26), .Text = "localhost"}
            UITheme.StyleTextBox(txtHost)

            Dim lblP As New Label() With {.Text = "Puerto:", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextPrimary, .Location = New Point(215, 25), .AutoSize = True}
            numPort = New NumericUpDown() With {.Location = New Point(215, 48), .Size = New Size(80, 26), .Minimum = 1, .Maximum = 65535, .Value = 3306}

            Dim lblD As New Label() With {.Text = "Nombre Base de Datos Web:", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextPrimary, .Location = New Point(310, 25), .AutoSize = True}
            txtDbName = New TextBox() With {.Location = New Point(310, 48), .Size = New Size(190, 26), .Text = "grupo5"}
            UITheme.StyleTextBox(txtDbName)

            Dim lblU As New Label() With {.Text = "Usuario:", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextPrimary, .Location = New Point(20, 85), .AutoSize = True}
            txtUser = New TextBox() With {.Location = New Point(20, 108), .Size = New Size(180, 26), .Text = "root"}
            UITheme.StyleTextBox(txtUser)

            Dim lblPw As New Label() With {.Text = "Contraseña:", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextPrimary, .Location = New Point(215, 85), .AutoSize = True}
            txtPassword = New TextBox() With {.Location = New Point(215, 108), .Size = New Size(180, 26), .PasswordChar = "*"c}
            UITheme.StyleTextBox(txtPassword)

            btnProbar = New Button() With {.Text = "Conectar y analizar web", .Location = New Point(420, 95), .Size = New Size(280, 42)}
            UITheme.StyleButton(btnProbar, "Primary")
            AddHandler btnProbar.Click, AddressOf BtnProbar_Click

            grpWebDb.Controls.AddRange({lblH, txtHost, lblP, numPort, lblD, txtDbName, lblU, txtUser, lblPw, txtPassword, btnProbar})
            Me.Controls.Add(grpWebDb)

            ' Resumen de datos detectados en la web
            Dim grpStats As New GroupBox() With {
                .Text = "Elementos Detectados en el E-commerce",
                .Location = New Point(20, 255),
                .Size = New Size(725, 95),
                .BackColor = UITheme.ColorSurface,
                .Font = UITheme.FontBold,
                .ForeColor = UITheme.ColorPrimaryDark
            }

            lblCatCount = New Label() With {.Text = "Categorías: -", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextPrimary, .Location = New Point(30, 40), .AutoSize = True}
            lblTallesCount = New Label() With {.Text = "Talles: -", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextPrimary, .Location = New Point(200, 40), .AutoSize = True}
            lblProdCount = New Label() With {.Text = "Productos: -", .Font = UITheme.FontRegular, .ForeColor = UITheme.ColorTextPrimary, .Location = New Point(350, 40), .AutoSize = True}
            lblStockCount = New Label() With {.Text = "Stock Total: -", .Font = UITheme.FontBold, .ForeColor = UITheme.ColorSuccess, .Location = Point.Add(New Point(520, 40), New Size(0, 0)), .AutoSize = True}

            grpStats.Controls.AddRange({lblCatCount, lblTallesCount, lblProdCount, lblStockCount})
            Me.Controls.Add(grpStats)

            ' Botón de acción principal
            btnImportar = New Button() With {
                .Text = "IMPORTAR CATÁLOGO AL SOFTWARE",
                .Location = New Point(20, 360),
                .Size = New Size(725, 45),
                .Font = New Font("Segoe UI", 11.0F, FontStyle.Bold)
            }
            UITheme.StyleButton(btnImportar, "Success")
            btnImportar.Enabled = False
            AddHandler btnImportar.Click, AddressOf BtnImportar_Click
            Me.Controls.Add(btnImportar)

            ' Log de salida
            Dim lblLogTitle As New Label() With {.Text = "Registro del Proceso de Migración:", .Font = UITheme.FontBold, .Location = New Point(20, 415), .AutoSize = True}
            txtLog = New TextBox() With {
                .Location = New Point(20, 440),
                .Size = New Size(725, 120),
                .Multiline = True,
                .ReadOnly = True,
                .ScrollBars = ScrollBars.Vertical,
                .Font = New Font("Consolas", 9.0F, FontStyle.Regular),
                .BackColor = UITheme.ColorSecondary,
                .ForeColor = UITheme.ColorTextOnDark,
                .Text = "Presione 'Conectar y Analizar Web' para detectar las tablas y prendas de su tienda web."
            }
            Me.Controls.AddRange({lblLogTitle, txtLog})
        End Sub

        Private Sub BtnProbar_Click(sender As Object, e As EventArgs)
            Dim host = txtHost.Text.Trim()
            Dim port = Convert.ToInt32(numPort.Value)
            Dim db = txtDbName.Text.Trim()
            Dim user = txtUser.Text.Trim()
            Dim pass = txtPassword.Text

            txtLog.Text = $"Conectando a {host}:{port}/{db}..." & vbCrLf

            Dim errMsg As String = ""
            If syncService.TestEcommerceConnection(host, port, user, pass, db, errMsg) Then
                txtLog.AppendText("✔ Conexión establecida con la base de datos de la web." & vbCrLf)
                txtLog.AppendText("Consultando registros de indumentaria..." & vbCrLf)

                Dim stats = syncService.GetEcommerceStats(host, port, user, pass, db)
                lblCatCount.Text = $"Categorías: {stats("categorias")}"
                lblTallesCount.Text = $"Talles: {stats("talles")}"
                lblProdCount.Text = $"Productos: {stats("productos")}"
                lblStockCount.Text = $"Stock Total: {stats("stock")} prendas"

                txtLog.AppendText($"Se encontraron {stats("productos")} prendas registradas con {stats("stock")} unidades de stock total distribuidas en sus talles." & vbCrLf)
                txtLog.AppendText("Listo para importar hacia el nuevo software de gestión." & vbCrLf)

                btnImportar.Enabled = True
            Else
                txtLog.AppendText("✖ No se pudo conectar a la base de datos del e-commerce: " & errMsg & vbCrLf)
                txtLog.AppendText("Verifique que MySQL esté iniciado y que el nombre de la base de datos sea correcto." & vbCrLf)
                btnImportar.Enabled = False
            End If
        End Sub

        Private Sub BtnImportar_Click(sender As Object, e As EventArgs)
            Dim host = txtHost.Text.Trim()
            Dim port = Convert.ToInt32(numPort.Value)
            Dim db = txtDbName.Text.Trim()
            Dim user = txtUser.Text.Trim()
            Dim pass = txtPassword.Text

            Dim resp = MessageBox.Show("¿Deseas iniciar la importación del catálogo y stock desde la web hacia este software?" & vbCrLf & vbCrLf & "Los productos existentes se actualizarán y los nuevos se darán de alta.", "Confirmar Migración", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If resp = DialogResult.Yes Then
                txtLog.AppendText("Iniciando migración de datos..." & vbCrLf)
                Dim report As String = ""
                If syncService.ImportCatalogFromEcommerce(host, port, user, pass, db, report) Then
                    txtLog.AppendText("✔ " & report & vbCrLf)
                    MessageBox.Show(report, "Migración Completada", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    txtLog.AppendText("✖ " & report & vbCrLf)
                    MessageBox.Show(report, "Error en Migración", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        End Sub

    End Class
End Namespace
