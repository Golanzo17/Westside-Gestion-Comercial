' ARCHIVO: UITheme.vb
' PROPÓSITO: Sistema de Diseño Centralizado v2 (Design System, Paleta, Tipografías y Helpers UI)
' - Construcción de la UI 100% por Código:
'   1. En lugar de coordenadas fijas que se descalibran con resoluciones High-DPI,
'      instanciar y estilizar por código mediante 'UITheme' da un diseño modular y con control de cada píxel.
'   2. Cambios de tema o colores corporativos se propagan instantáneamente modificando las constantes aquí.
' - Características Incorporadas:
'   - Paleta cromática Slate e Indigo limpia y descansada para la vista.
'   - Bordes redondeados dinámicos mediante regiones GDI+ (GraphicsPath).
'   - Tablas DataGridView con hover interactivo de filas y cabeceras contrastadas.
'   - Sistema de Toasts flotantes no invasivos para notificaciones inmediatas.
'   - Sombras sutiles en tarjetas para crear profundidad visual moderna.
'   - Badges de estado para etiquetar información contextual.


Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Namespace UI
    Public Module UITheme
        ' Paleta de colores institucionales

        Public ReadOnly ColorPrimary As Color = Color.FromArgb(99, 102, 241)        ' Indigo 500
        Public ReadOnly ColorPrimaryDark As Color = Color.FromArgb(79, 70, 229)     ' Indigo 600
        Public ReadOnly ColorPrimaryLight As Color = Color.FromArgb(224, 231, 255)  ' Indigo 100
        Public ReadOnly ColorSecondary As Color = Color.FromArgb(15, 23, 42)        ' Slate 900
        Public ReadOnly ColorSidebar As Color = Color.FromArgb(17, 24, 39)          ' Gray 900 profundo
        Public ReadOnly ColorSidebarHover As Color = Color.FromArgb(31, 41, 55)     ' Gray 800
        Public ReadOnly ColorSidebarActive As Color = Color.FromArgb(49, 46, 129)   ' Indigo 900
        Public ReadOnly ColorBackground As Color = Color.FromArgb(248, 250, 252)    ' Slate 50
        Public ReadOnly ColorSurface As Color = Color.White
        Public ReadOnly ColorSurfaceMuted As Color = Color.FromArgb(241, 245, 249) ' Slate 100
        Public ReadOnly ColorSurfaceAccent As Color = Color.FromArgb(238, 242, 255) ' Indigo 50
        Public ReadOnly ColorBorder As Color = Color.FromArgb(226, 232, 240)        ' Slate 200
        Public ReadOnly ColorBorderStrong As Color = Color.FromArgb(203, 213, 225) ' Slate 300
        Public ReadOnly ColorTextPrimary As Color = Color.FromArgb(15, 23, 42)      ' Slate 900
        Public ReadOnly ColorTextSecondary As Color = Color.FromArgb(100, 116, 139) ' Slate 500
        Public ReadOnly ColorTextMuted As Color = Color.FromArgb(148, 163, 184)     ' Slate 400
        Public ReadOnly ColorTextOnDark As Color = Color.FromArgb(203, 213, 225)    ' Slate 300
        Public ReadOnly ColorSuccess As Color = Color.FromArgb(16, 185, 129)        ' Emerald 500
        Public ReadOnly ColorSuccessDark As Color = Color.FromArgb(5, 150, 105)     ' Emerald 600
        Public ReadOnly ColorSuccessLight As Color = Color.FromArgb(209, 250, 229)  ' Emerald 100
        Public ReadOnly ColorDanger As Color = Color.FromArgb(239, 68, 68)          ' Rose 500
        Public ReadOnly ColorDangerDark As Color = Color.FromArgb(220, 38, 38)      ' Rose 600
        Public ReadOnly ColorDangerLight As Color = Color.FromArgb(254, 226, 226)   ' Rose 100
        Public ReadOnly ColorWarning As Color = Color.FromArgb(245, 158, 11)        ' Amber 500
        Public ReadOnly ColorWarningLight As Color = Color.FromArgb(254, 243, 199)  ' Amber 100
        Public ReadOnly ColorInfo As Color = Color.FromArgb(59, 130, 246)           ' Blue 500
        Public ReadOnly ColorInfoLight As Color = Color.FromArgb(219, 234, 254)     ' Blue 100
        Public ReadOnly ColorRowHover As Color = Color.FromArgb(238, 242, 255)      ' Hover fila tabla
        Public ReadOnly ColorShadow As Color = Color.FromArgb(25, 15, 23, 42)       ' Sombra sutil (Slate 900 con 10% opacidad)
        Public ReadOnly ColorShadowMedium As Color = Color.FromArgb(40, 15, 23, 42) ' Sombra media
        Public ReadOnly ColorSidebarSection As Color = Color.FromArgb(55, 65, 81)   ' Gray 700 (separadores sidebar)

        ' Tipografías institucionales con jerarquía visual clara (Segoe UI)
        ' v2: Escala expandida para crear contraste entre secciones
        Public ReadOnly FontHeadingLarge As New Font("Segoe UI", 20.0F, FontStyle.Bold)
        Public ReadOnly FontHeading As New Font("Segoe UI", 16.0F, FontStyle.Bold)
        Public ReadOnly FontSubheading As New Font("Segoe UI", 12.0F, FontStyle.Bold)
        Public ReadOnly FontSectionTitle As New Font("Segoe UI", 10.5F, FontStyle.Bold)
        Public ReadOnly FontRegular As New Font("Segoe UI", 9.5F, FontStyle.Regular)
        Public ReadOnly FontBold As New Font("Segoe UI", 9.5F, FontStyle.Bold)
        Public ReadOnly FontSmall As New Font("Segoe UI", 8.5F, FontStyle.Regular)
        Public ReadOnly FontSmallBold As New Font("Segoe UI", 8.5F, FontStyle.Bold)
        Public ReadOnly FontCaption As New Font("Segoe UI", 7.5F, FontStyle.Bold)
        Public ReadOnly FontPriceBig As New Font("Segoe UI", 22.0F, FontStyle.Bold)
        Public ReadOnly FontKpiValue As New Font("Segoe UI", 18.0F, FontStyle.Bold)
        Public ReadOnly FontButton As New Font("Segoe UI", 9.5F, FontStyle.Bold)
        Public ReadOnly FontNavItem As New Font("Segoe UI", 10.0F, FontStyle.Regular)
        Public ReadOnly FontNavSection As New Font("Segoe UI", 7.5F, FontStyle.Bold)

        ' Botones con estilo plano y bordes redondeados

        Public Sub StyleButton(btn As Button, Optional btnType As String = "Primary", Optional rounded As Boolean = True)
            btn.FlatStyle = FlatStyle.Flat
            btn.FlatAppearance.BorderSize = 0
            btn.Cursor = Cursors.Hand
            btn.Font = FontButton
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
                    btn.BackColor = ColorSurface
                    btn.ForeColor = ColorTextPrimary
                    btn.FlatAppearance.BorderSize = 1
                    btn.FlatAppearance.BorderColor = ColorBorder
                    btn.FlatAppearance.MouseOverBackColor = ColorSurfaceMuted
                Case "GHOST"
                    ' Botón transparente que sólo muestra hover, ideal para acciones secundarias
                    btn.BackColor = Color.Transparent
                    btn.ForeColor = ColorPrimary
                    btn.FlatAppearance.BorderSize = 0
                    btn.FlatAppearance.MouseOverBackColor = ColorSurfaceAccent
                    rounded = True
                Case "SIDEBAR"
                    btn.BackColor = Color.Transparent
                    btn.ForeColor = ColorTextOnDark
                    btn.TextAlign = ContentAlignment.MiddleLeft
                    btn.Padding = New Padding(22, 0, 0, 0)
                    btn.Font = FontNavItem
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

        ''' <summary>
        ''' Aplica esquinas redondeadas modernas a cualquier control WinForms mediante GraphicsPath.
        ''' </summary>
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
                ' Si el control aún no tiene tamaño renderizado, se ignora de forma segura
            End Try
        End Sub

        ' DataGridView: formato de grillas con hover interactivo en filas

        Public Sub StyleDataGrid(dgv As DataGridView, Optional readOnlyGrid As Boolean = True)
            dgv.ReadOnly = readOnlyGrid
            dgv.AllowUserToAddRows = False
            dgv.AllowUserToDeleteRows = False
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

        ' Cajas de texto y selectores

        Public Sub StyleTextBox(txt As TextBox)
            txt.BorderStyle = BorderStyle.FixedSingle
            txt.Font = FontRegular
            txt.BackColor = Color.White
            txt.ForeColor = ColorTextPrimary
        End Sub

        Public Sub StyleComboBox(cbo As ComboBox)
            cbo.Font = FontRegular
            cbo.BackColor = Color.White
            cbo.ForeColor = ColorTextPrimary
            cbo.FlatStyle = FlatStyle.Flat
        End Sub

        ' Mensajes flotantes temporales (toasts no bloqueantes)

        ''' <summary>
        ''' Despliega un mensaje flotante temporal (toast) en la parte superior del formulario.
        ''' 
        ''' DECISIÓN DE UX:
        ''' En un punto de venta (POS), el cajero necesita velocidad. Interrumpir el flujo
        ''' con ventanas emergentes modales (MessageBox.Show) ralentiza el cobro en horas pico.
        ''' Los toasts confirman operaciones en 3 segundos sin bloquear el teclado ni el lector.
        ''' </summary>
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

            ' Botón para cerrar manualmente antes de que expire el temporizador
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

            ' Temporizador para autodestrucción suave
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

        ' ─── Componentes v2: Tarjetas, Badges, Separadores ───

        ''' <summary>
        ''' Genera una tarjeta métrica con borde sutil, sombra visual y barra lateral de acento de color.
        ''' Usado tanto en el Dashboard principal como en los reportes comerciales.
        ''' v2: Tamaño mayor, sombra simulada con doble panel, mejor espaciado.
        ''' </summary>
        Public Function CreateKpiCard(title As String, value As String, subtitle As String, accentColor As Color) As Panel
            Dim pnl As New Panel() With {
                .BackColor = ColorSurface,
                .Size = New Size(230, 115),
                .Margin = New Padding(8),
                .Padding = New Padding(14)
            }

            ' Franja vertical de acento visual (Verde, Azul, Ámbar, etc.)
            Dim strip As New Panel() With {
                .BackColor = accentColor,
                .Dock = DockStyle.Left,
                .Width = 4
            }
            pnl.Controls.Add(strip)

            Dim lblTitle As New Label() With {
                .Text = title.ToUpper(),
                .Font = FontCaption,
                .ForeColor = ColorTextSecondary,
                .Location = New Point(16, 12),
                .AutoSize = True
            }
            pnl.Controls.Add(lblTitle)

            Dim lblValue As New Label() With {
                .Text = value,
                .Font = FontKpiValue,
                .ForeColor = ColorTextPrimary,
                .Location = New Point(16, 32),
                .AutoSize = True
            }
            pnl.Controls.Add(lblValue)

            Dim lblSub As New Label() With {
                .Text = subtitle,
                .Font = FontSmall,
                .ForeColor = ColorTextMuted,
                .Location = New Point(16, 76),
                .AutoSize = True
            }
            pnl.Controls.Add(lblSub)

            ' Borde sutil y elegante de 1px usando Paint
            AddHandler pnl.Paint, Sub(s, ev)
                                      Using pen As New Pen(ColorBorder, 1)
                                          ev.Graphics.DrawRectangle(pen, 0, 0, pnl.Width - 1, pnl.Height - 1)
                                      End Using
                                  End Sub

            Return pnl
        End Function

        ''' <summary>
        ''' Crea un badge de estado con texto, color de fondo y color de texto.
        ''' Ideal para etiquetar estados en grillas o headers (ej: "Activa", "Anulada", "Bajo Stock").
        ''' </summary>
        Public Function CreateBadge(text As String, bgColor As Color, fgColor As Color) As Label
            Dim lbl As New Label() With {
                .Text = $"  {text}  ",
                .Font = FontSmallBold,
                .BackColor = bgColor,
                .ForeColor = fgColor,
                .AutoSize = True,
                .TextAlign = ContentAlignment.MiddleCenter,
                .Padding = New Padding(4, 2, 4, 2)
            }
            AddHandler lbl.Resize, Sub(s, ev)
                                       Dim l = CType(s, Label)
                                       If l.Width > 0 AndAlso l.Height > 0 Then ApplyRoundedRegion(l, 6)
                                   End Sub
            Return lbl
        End Function

        ''' <summary>
        ''' Crea un separador visual para el sidebar con etiqueta de sección.
        ''' Agrupa los botones de navegación en bloques lógicos (Operaciones, Gestión, Sistema).
        ''' </summary>
        Public Function CreateSidebarSection(sectionTitle As String) As Panel
            Dim pnl As New Panel() With {
                .Height = 32,
                .Dock = DockStyle.Top,
                .BackColor = Color.Transparent,
                .Padding = New Padding(22, 12, 10, 0)
            }

            Dim lbl As New Label() With {
                .Text = sectionTitle.ToUpper(),
                .Font = FontNavSection,
                .ForeColor = ColorSidebarSection,
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.BottomLeft
            }
            pnl.Controls.Add(lbl)
            Return pnl
        End Function

        ''' <summary>
        ''' Crea una línea divisoria sutil para separar visualmente secciones dentro de un panel.
        ''' </summary>
        Public Function CreateDivider(Optional height As Integer = 1, Optional topMargin As Integer = 8, Optional bottomMargin As Integer = 8) As Panel
            Dim pnl As New Panel() With {
                .Height = height + topMargin + bottomMargin,
                .Dock = DockStyle.Top,
                .BackColor = Color.Transparent,
                .Padding = New Padding(22, topMargin, 22, bottomMargin)
            }
            Dim line As New Panel() With {
                .Dock = DockStyle.Fill,
                .BackColor = Color.FromArgb(40, 255, 255, 255)
            }
            pnl.Controls.Add(line)
            Return pnl
        End Function

    End Module
End Namespace
