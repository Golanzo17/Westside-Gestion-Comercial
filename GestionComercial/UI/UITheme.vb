Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Namespace UI
    Public Module UITheme
        ' Paleta de colores moderna y elegante para indumentaria / retail
        Public ReadOnly ColorPrimary As Color = Color.FromArgb(79, 70, 229)       ' Indigo #4F46E5
        Public ReadOnly ColorPrimaryDark As Color = Color.FromArgb(67, 56, 202)   ' Indigo Dark #4338CA
        Public ReadOnly ColorSecondary As Color = Color.FromArgb(15, 23, 42)      ' Slate 900 #0F172A
        Public ReadOnly ColorSidebar As Color = Color.FromArgb(30, 41, 59)        ' Slate 800 #1E293B
        Public ReadOnly ColorSidebarHover As Color = Color.FromArgb(51, 65, 85)   ' Slate 700 #334155
        Public ReadOnly ColorBackground As Color = Color.FromArgb(248, 250, 252)  ' Slate 50 #F8FAFC
        Public ReadOnly ColorSurface As Color = Color.White
        Public ReadOnly ColorBorder As Color = Color.FromArgb(226, 232, 240)      ' Slate 200 #E2E8F0
        Public ReadOnly ColorTextPrimary As Color = Color.FromArgb(15, 23, 42)    ' Slate 900
        Public ReadOnly ColorTextSecondary As Color = Color.FromArgb(100, 116, 139) ' Slate 500
        Public ReadOnly ColorSuccess As Color = Color.FromArgb(16, 185, 129)     ' Emerald 500 #10B981
        Public ReadOnly ColorDanger As Color = Color.FromArgb(239, 68, 68)       ' Rose 500 #EF4444
        Public ReadOnly ColorWarning As Color = Color.FromArgb(245, 158, 11)     ' Amber 500 #F59E0B
        Public ReadOnly ColorInfo As Color = Color.FromArgb(59, 130, 246)        ' Blue 500 #3B82F6

        ' Tipografías
        Public ReadOnly FontHeading As New Font("Segoe UI", 16.0F, FontStyle.Bold)
        Public ReadOnly FontSubheading As New Font("Segoe UI", 12.0F, FontStyle.Bold)
        Public ReadOnly FontRegular As New Font("Segoe UI", 9.5F, FontStyle.Regular)
        Public ReadOnly FontBold As New Font("Segoe UI", 9.5F, FontStyle.Bold)
        Public ReadOnly FontSmall As New Font("Segoe UI", 8.5F, FontStyle.Regular)
        Public ReadOnly FontPriceBig As New Font("Segoe UI", 22.0F, FontStyle.Bold)

        ''' <summary>
        ''' Aplica estilo moderno a un botón con bordes planos y colores vivos
        ''' </summary>
        Public Sub StyleButton(btn As Button, Optional btnType As String = "Primary")
            btn.FlatStyle = FlatStyle.Flat
            btn.FlatAppearance.BorderSize = 0
            btn.Cursor = Cursors.Hand
            btn.Font = FontBold
            btn.Height = Math.Max(btn.Height, 38)

            Select Case btnType.ToUpper()
                Case "PRIMARY"
                    btn.BackColor = ColorPrimary
                    btn.ForeColor = Color.White
                    btn.FlatAppearance.MouseOverBackColor = ColorPrimaryDark
                Case "SUCCESS"
                    btn.BackColor = ColorSuccess
                    btn.ForeColor = Color.White
                    btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(5, 150, 105)
                Case "DANGER"
                    btn.BackColor = ColorDanger
                    btn.ForeColor = Color.White
                    btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(220, 38, 38)
                Case "WARNING"
                    btn.BackColor = ColorWarning
                    btn.ForeColor = Color.White
                    btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(217, 119, 6)
                Case "SECONDARY"
                    btn.BackColor = Color.FromArgb(241, 245, 249)
                    btn.ForeColor = ColorTextPrimary
                    btn.FlatAppearance.BorderSize = 1
                    btn.FlatAppearance.BorderColor = ColorBorder
                    btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(226, 232, 240)
                Case "SIDEBAR"
                    btn.BackColor = Color.Transparent
                    btn.ForeColor = Color.FromArgb(226, 232, 240)
                    btn.TextAlign = ContentAlignment.MiddleLeft
                    btn.Padding = New Padding(15, 0, 0, 0)
                    btn.Font = New Font("Segoe UI", 10.0F, FontStyle.Regular)
                    btn.FlatAppearance.MouseOverBackColor = ColorSidebarHover
                Case Else
                    btn.BackColor = ColorPrimary
                    btn.ForeColor = Color.White
            End Select
        End Sub

        ''' <summary>
        ''' Aplica estilo elegante y limpio a un DataGridView
        ''' </summary>
        Public Sub StyleDataGrid(dgv As DataGridView)
            dgv.BackgroundColor = ColorSurface
            dgv.BorderStyle = BorderStyle.None
            dgv.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
            dgv.GridColor = Color.FromArgb(241, 245, 249)
            dgv.EnableHeadersVisualStyles = False
            dgv.RowHeadersVisible = False
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            dgv.MultiSelect = False
            dgv.AllowUserToResizeRows = False
            dgv.RowTemplate.Height = 36

            ' Cabecera de columnas
            dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(241, 245, 249)
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = ColorTextSecondary
            dgv.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 9.0F, FontStyle.Bold)
            dgv.ColumnHeadersDefaultCellStyle.Padding = New Padding(8, 0, 8, 0)
            dgv.ColumnHeadersHeight = 40

            ' Filas y celdas
            dgv.DefaultCellStyle.BackColor = ColorSurface
            dgv.DefaultCellStyle.ForeColor = ColorTextPrimary
            dgv.DefaultCellStyle.Font = FontRegular
            dgv.DefaultCellStyle.Padding = New Padding(8, 0, 8, 0)
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(238, 242, 255) ' Indigo light
            dgv.DefaultCellStyle.SelectionForeColor = ColorPrimaryDark

            ' Filas alternas
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 250, 252)
            dgv.AlternatingRowsDefaultCellStyle.SelectionBackColor = Color.FromArgb(238, 242, 255)
            dgv.AlternatingRowsDefaultCellStyle.SelectionForeColor = ColorPrimaryDark
        End Sub

        ''' <summary>
        ''' Estiliza una caja de texto
        ''' </summary>
        Public Sub StyleTextBox(txt As TextBox)
            txt.BorderStyle = BorderStyle.FixedSingle
            txt.Font = FontRegular
            txt.BackColor = Color.White
            txt.ForeColor = ColorTextPrimary
        End Sub

        ''' <summary>
        ''' Crea una tarjeta de indicador numérico (KPI Card)
        ''' </summary>
        Public Function CreateKpiCard(title As String, value As String, subtitle As String, accentColor As Color) As Panel
            Dim pnl As New Panel() With {
                .BackColor = ColorSurface,
                .Size = New Size(220, 100),
                .Margin = New Padding(8),
                .Padding = New Padding(12)
            }

            ' Borde izquierdo decorativo
            Dim strip As New Panel() With {
                .BackColor = accentColor,
                .Dock = DockStyle.Left,
                .Width = 5
            }
            pnl.Controls.Add(strip)

            Dim lblTitle As New Label() With {
                .Text = title.ToUpper(),
                .Font = FontSmall,
                .ForeColor = ColorTextSecondary,
                .Location = New Point(15, 10),
                .AutoSize = True
            }
            pnl.Controls.Add(lblTitle)

            Dim lblValue As New Label() With {
                .Text = value,
                .Font = New Font("Segoe UI", 16.0F, FontStyle.Bold),
                .ForeColor = ColorTextPrimary,
                .Location = New Point(15, 30),
                .AutoSize = True
            }
            pnl.Controls.Add(lblValue)

            Dim lblSub As New Label() With {
                .Text = subtitle,
                .Font = FontSmall,
                .ForeColor = ColorTextSecondary,
                .Location = New Point(15, 68),
                .AutoSize = True
            }
            pnl.Controls.Add(lblSub)

            Return pnl
        End Function

    End Module
End Namespace
