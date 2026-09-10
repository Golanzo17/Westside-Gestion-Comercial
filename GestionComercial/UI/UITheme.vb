Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Namespace UI
    Public Module UITheme
        ' ═══════════════════════════════════════════════════════════
        ' PALETA PREMIUM – Colores más vibrantes y modernos
        ' ═══════════════════════════════════════════════════════════
        Public ReadOnly ColorPrimary As Color = Color.FromArgb(99, 102, 241)        ' Indigo 500 más vivo
        Public ReadOnly ColorPrimaryDark As Color = Color.FromArgb(79, 70, 229)     ' Indigo 600
        Public ReadOnly ColorPrimaryLight As Color = Color.FromArgb(224, 231, 255)  ' Indigo 100
        Public ReadOnly ColorSecondary As Color = Color.FromArgb(15, 23, 42)        ' Slate 900
        Public ReadOnly ColorSidebar As Color = Color.FromArgb(17, 24, 39)          ' Gray 900 profundo
        Public ReadOnly ColorSidebarHover As Color = Color.FromArgb(31, 41, 55)     ' Gray 800
        Public ReadOnly ColorSidebarActive As Color = Color.FromArgb(49, 46, 129)   ' Indigo 900
        Public ReadOnly ColorBackground As Color = Color.FromArgb(248, 250, 252)    ' Slate 50
        Public ReadOnly ColorSurface As Color = Color.White
        Public ReadOnly ColorBorder As Color = Color.FromArgb(226, 232, 240)        ' Slate 200
        Public ReadOnly ColorTextPrimary As Color = Color.FromArgb(15, 23, 42)      ' Slate 900
        Public ReadOnly ColorTextSecondary As Color = Color.FromArgb(100, 116, 139) ' Slate 500
        Public ReadOnly ColorSuccess As Color = Color.FromArgb(16, 185, 129)        ' Emerald 500
        Public ReadOnly ColorSuccessDark As Color = Color.FromArgb(5, 150, 105)     ' Emerald 600
        Public ReadOnly ColorDanger As Color = Color.FromArgb(239, 68, 68)          ' Rose 500
        Public ReadOnly ColorDangerDark As Color = Color.FromArgb(220, 38, 38)      ' Rose 600
        Public ReadOnly ColorWarning As Color = Color.FromArgb(245, 158, 11)        ' Amber 500
        Public ReadOnly ColorInfo As Color = Color.FromArgb(59, 130, 246)           ' Blue 500
        Public ReadOnly ColorRowHover As Color = Color.FromArgb(238, 242, 255)      ' Hover fila tabla

        ' Tipografías
        Public ReadOnly FontHeading As New Font("Segoe UI", 16.0F, FontStyle.Bold)
        Public ReadOnly FontSubheading As New Font("Segoe UI", 12.0F, FontStyle.Bold)
        Public ReadOnly FontRegular As New Font("Segoe UI", 9.5F, FontStyle.Regular)
        Public ReadOnly FontBold As New Font("Segoe UI", 9.5F, FontStyle.Bold)
        Public ReadOnly FontSmall As New Font("Segoe UI", 8.5F, FontStyle.Regular)
        Public ReadOnly FontPriceBig As New Font("Segoe UI", 22.0F, FontStyle.Bold)

        ' ═══════════════════════════════════════════════════════════
        ' BOTONES – Estilo plano con bordes redondeados
        ' ═══════════════════════════════════════════════════════════
        Public Sub StyleButton(btn As Button, Optional btnType As String = "Primary", Optional rounded As Boolean = True)
            btn.FlatStyle = FlatStyle.Flat
            btn.FlatAppearance.BorderSize = 0
            btn.Cursor = Cursors.Hand
            btn.Font = FontBold
            btn.Height = Math.Max(btn.Height, 38)
            btn.TextAlign = ContentAlignment.MiddleCenter

            Select Case btnType.ToUpper()
                Case "PRIMARY"
                    btn.BackColor = ColorPrimary
                    btn.ForeColor = Color.White
                    btn.FlatAppearance.MouseOverBackColor = ColorPrimaryDark
                Case "SUCCESS"
                    btn.BackColor = ColorSuccess
                    btn.ForeColor = Color.White
                    btn.FlatAppearance.MouseOverBackColor = ColorSuccessDark
                Case "DANGER"
                    btn.BackColor = ColorDanger
                    btn.ForeColor = Color.White
                    btn.FlatAppearance.MouseOverBackColor = ColorDangerDark
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
                    rounded = False
                Case "SIDEBAR"
                    btn.BackColor = Color.Transparent
                    btn.ForeColor = Color.FromArgb(203, 213, 225)
                    btn.TextAlign = ContentAlignment.MiddleLeft
                    btn.Padding = New Padding(20, 0, 0, 0)
                    btn.Font = New Font("Segoe UI", 10.0F, FontStyle.Regular)
                    btn.FlatAppearance.MouseOverBackColor = ColorSidebarHover
                    rounded = False
                Case Else
                    btn.BackColor = ColorPrimary
                    btn.ForeColor = Color.White
            End Select

            If rounded AndAlso btn.Width > 0 AndAlso btn.Height > 0 Then
                ApplyRoundedRegion(btn, 8)
                AddHandler btn.Resize, Sub(s, ev)
                                           Dim b = CType(s, Button)
                                           If b.Width > 0 AndAlso b.Height > 0 Then ApplyRoundedRegion(b, 8)
                                       End Sub
            End If
        End Sub

        ''' <summary>Aplica esquinas redondeadas a cualquier control.</summary>
        Public Sub ApplyRoundedRegion(ctrl As Control, radius As Integer)
            Try
                Dim path As New GraphicsPath()
                Dim r = radius * 2
                path.AddArc(0, 0, r, r, 180, 90)
                path.AddArc(ctrl.Width - r, 0, r, r, 270, 90)
                path.AddArc(ctrl.Width - r, ctrl.Height - r, r, r, 0, 90)
                path.AddArc(0, ctrl.Height - r, r, r, 90, 90)
                path.CloseFigure()
                ctrl.Region = New Region(path)
            Catch
                ' Si el control aún no tiene tamaño, ignorar
            End Try
        End Sub

        ' ═══════════════════════════════════════════════════════════
        ' DATAGRIDVIEW – Con hover en filas
        ' ═══════════════════════════════════════════════════════════
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
            dgv.RowTemplate.Height = 45
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            ' Cabecera de columnas – no cambia de color al recibir foco
            dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(241, 245, 249)
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = ColorTextSecondary
            dgv.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(241, 245, 249)
            dgv.ColumnHeadersDefaultCellStyle.SelectionForeColor = ColorTextSecondary
            dgv.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 9.0F, FontStyle.Bold)
            dgv.ColumnHeadersDefaultCellStyle.Padding = New Padding(8, 0, 8, 0)
            dgv.ColumnHeadersHeight = 50
            dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing

            ' Celdas normales
            dgv.DefaultCellStyle.BackColor = ColorSurface
            dgv.DefaultCellStyle.ForeColor = ColorTextPrimary
            dgv.DefaultCellStyle.Font = FontRegular
            dgv.DefaultCellStyle.Padding = New Padding(8, 0, 8, 0)
            dgv.DefaultCellStyle.SelectionBackColor = ColorPrimary
            dgv.DefaultCellStyle.SelectionForeColor = Color.White

            ' Filas alternas
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 250, 252)
            dgv.AlternatingRowsDefaultCellStyle.SelectionBackColor = ColorPrimary
            dgv.AlternatingRowsDefaultCellStyle.SelectionForeColor = Color.White

            ' ─── Efecto hover en filas ───
            AddHandler dgv.CellMouseEnter, Sub(s, ev)
                                               Dim g = CType(s, DataGridView)
                                               If ev.RowIndex >= 0 AndAlso Not g.Rows(ev.RowIndex).Selected Then
                                                   g.Rows(ev.RowIndex).DefaultCellStyle.BackColor = ColorRowHover
                                               End If
                                           End Sub
            AddHandler dgv.CellMouseLeave, Sub(s, ev)
                                               Dim g = CType(s, DataGridView)
                                               If ev.RowIndex >= 0 AndAlso Not g.Rows(ev.RowIndex).Selected Then
                                                   g.Rows(ev.RowIndex).DefaultCellStyle.BackColor = Color.Empty
                                               End If
                                           End Sub
        End Sub

        ' ═══════════════════════════════════════════════════════════
        ' TEXTBOX
        ' ═══════════════════════════════════════════════════════════
        Public Sub StyleTextBox(txt As TextBox)
            txt.BorderStyle = BorderStyle.FixedSingle
            txt.Font = FontRegular
            txt.BackColor = Color.White
            txt.ForeColor = ColorTextPrimary
        End Sub

        ' ═══════════════════════════════════════════════════════════
        ' TOAST / NOTIFICACIÓN INLINE – Aparece arriba y se va solo
        ' ═══════════════════════════════════════════════════════════
        Public Sub ShowToast(parentControl As Control, message As String, Optional toastType As String = "Success", Optional durationMs As Integer = 3000)
            Dim backColor As Color
            Dim icon As String
            Select Case toastType.ToUpper()
                Case "SUCCESS"
                    backColor = ColorSuccess
                    icon = "✔  "
                Case "ERROR"
                    backColor = ColorDanger
                    icon = "✘  "
                Case "WARNING"
                    backColor = ColorWarning
                    icon = "⚠  "
                Case Else
                    backColor = ColorInfo
                    icon = "ℹ  "
            End Select

            Dim toastWidth = Math.Min(parentControl.Width - 40, 520)
            Dim toast As New Panel() With {
                .Size = New Size(toastWidth, 50),
                .BackColor = backColor,
                .Location = New Point((parentControl.Width - toastWidth) \ 2, 14),
                .Padding = New Padding(16, 0, 16, 0)
            }

            Dim lblMsg As New Label() With {
                .Text = icon & message,
                .Font = FontBold,
                .ForeColor = Color.White,
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleLeft
            }
            toast.Controls.Add(lblMsg)

            ' Botón cerrar
            Dim btnClose As New Label() With {
                .Text = "×",
                .Font = New Font("Segoe UI", 14.0F, FontStyle.Bold),
                .ForeColor = Color.White,
                .Dock = DockStyle.Right,
                .Width = 30,
                .TextAlign = ContentAlignment.MiddleCenter,
                .Cursor = Cursors.Hand
            }
            AddHandler btnClose.Click, Sub() parentControl.Controls.Remove(toast)
            toast.Controls.Add(btnClose)

            parentControl.Controls.Add(toast)
            toast.BringToFront()
            ApplyRoundedRegion(toast, 10)

            Dim t As New Timer() With {.Interval = durationMs}
            AddHandler t.Tick, Sub()
                                   t.Stop()
                                   Try
                                       If parentControl.Controls.Contains(toast) Then
                                           parentControl.Controls.Remove(toast)
                                       End If
                                       toast.Dispose()
                                   Catch
                                   End Try
                               End Sub
            t.Start()
        End Sub

        ' ═══════════════════════════════════════════════════════════
        ' KPI CARDS – Tarjetas del dashboard
        ' ═══════════════════════════════════════════════════════════
        Public Function CreateKpiCard(title As String, value As String, subtitle As String, accentColor As Color) As Panel
            Dim pnl As New Panel() With {
                .BackColor = ColorSurface,
                .Size = New Size(220, 105),
                .Margin = New Padding(8),
                .Padding = New Padding(12)
            }

            ' Borde izquierdo de acento
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
                .Location = New Point(15, 72),
                .AutoSize = True
            }
            pnl.Controls.Add(lblSub)

            Return pnl
        End Function

    End Module
End Namespace
