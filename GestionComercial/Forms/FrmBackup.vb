Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports GestionComercial.Services
Imports GestionComercial.UI

' ARCHIVO: FrmBackup.vb
' PROPÓSITO: Módulo Exclusivo de Administrador para Copias de Seguridad (Backups) y Restauración
' - Respaldos en Vivo Seguros (Online Backup):
'   Ejecuta copias atómicas consistentes mediante BackupService sin necesidad de detener la aplicación.
' - Historial de Puntos de Recuperación:
'   Muestra todas las copias disponibles en la carpeta de resguardo con fecha, tamaño e integridad.
' - Restauración Asistida:
'   Permite restablecer un punto anterior con confirmación estricta de seguridad y reinicio de conexiones.

Namespace Forms
    Public Class FrmBackup
        Inherits Form

        Private backupService As New BackupService()

        Private lblDbNombre As Label
        Private lblCantBackups As Label
        Private lblUltimoBackup As Label
        Private lblIntegridad As Label

        Private btnCrearBackup As Button
        Private btnCrearEnRuta As Button
        Private btnRestaurar As Button
        Private btnAbrirCarpeta As Button
        Private btnRefrescar As Button

        Private dgvBackups As DataGridView

        Public Sub New()
            InitializeUI()
            LoadDashboardBackup()
        End Sub

        Protected Overrides Sub OnLoad(e As EventArgs)
            MyBase.OnLoad(e)
            If Not AuthService.IsAdmin Then
                MessageBox.Show("El acceso a las copias de seguridad está reservado exclusivamente para Administradores.", "Acceso Restringido", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Me.BeginInvoke(New Action(AddressOf Me.Close))
            End If
        End Sub

        Private Sub InitializeUI()
            Me.Text = "Copias de Seguridad (Backup) y Restauración de Base de Datos"
            Me.Size = New Size(1060, 680)
            Me.MinimumSize = New Size(920, 560)
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.BackColor = UITheme.ColorBackground
            Me.Font = UITheme.FontRegular

            ' Header superior institucional
            Dim pnlHeader As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 65,
                .BackColor = UITheme.ColorSecondary,
                .Padding = New Padding(20, 12, 20, 12)
            }
            Dim lblTitle As New Label() With {
                .Text = "Copias de Seguridad (Backup) y Puntos de Restauración",
                .Font = UITheme.FontHeading,
                .ForeColor = Color.White,
                .AutoSize = True,
                .Location = New Point(18, 10)
            }
            Dim lblSub As New Label() With {
                .Text = "Gestión exclusiva de Administrador: resguardo transaccional del archivo relacional SQLite",
                .Font = UITheme.FontSmall,
                .ForeColor = UITheme.ColorTextMuted,
                .AutoSize = True,
                .Location = New Point(20, 38)
            }
            pnlHeader.Controls.AddRange({lblTitle, lblSub})

            ' Panel de Tarjetas Informativas (KPIs)
            Dim pnlCards As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 125,
                .BackColor = UITheme.ColorSurfaceMuted,
                .Padding = New Padding(15, 5, 15, 5)
            }

            Dim cardDb = UITheme.CreateKpiCard("BASE ACTIVA", "SQLite (.db)", "En uso por el sistema", UITheme.ColorPrimary)
            cardDb.Location = New Point(15, 5)
            lblDbNombre = CType(cardDb.Controls(2), Label)

            Dim cardCant = UITheme.CreateKpiCard("TOTAL BACKUPS", "0", "Copias disponibles", UITheme.ColorSuccess)
            cardCant.Location = New Point(265, 5)
            lblCantBackups = CType(cardCant.Controls(2), Label)

            Dim cardUlt = UITheme.CreateKpiCard("ÚLTIMO RESPALDO", "-", "Punto más reciente", UITheme.ColorInfo)
            cardUlt.Location = New Point(515, 5)
            lblUltimoBackup = CType(cardUlt.Controls(2), Label)

            Dim cardEst = UITheme.CreateKpiCard("INTEGRIDAD", "100%", "Conformidad ACID activa", UITheme.ColorWarning)
            cardEst.Location = New Point(765, 5)
            lblIntegridad = CType(cardEst.Controls(2), Label)

            pnlCards.Controls.AddRange({cardDb, cardCant, cardUlt, cardEst})

            ' Toolbar de Operaciones
            Dim pnlToolbar As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 55,
                .BackColor = UITheme.ColorSurface
            }
            Dim flpToolbar As New FlowLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(12, 10, 10, 8),
                .WrapContents = False
            }

            btnCrearBackup = New Button() With {
                .Text = "Crear Backup Inmediato",
                .Height = 34,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Padding = New Padding(12, 0, 12, 0),
                .Margin = New Padding(0, 0, 8, 0)
            }
            UITheme.StyleButton(btnCrearBackup, "Success")
            AddHandler btnCrearBackup.Click, AddressOf BtnCrearBackup_Click

            btnCrearEnRuta = New Button() With {
                .Text = "Guardar en Otra Carpeta...",
                .Height = 34,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Padding = New Padding(12, 0, 12, 0),
                .Margin = New Padding(0, 0, 8, 0)
            }
            UITheme.StyleButton(btnCrearEnRuta, "Primary")
            AddHandler btnCrearEnRuta.Click, AddressOf BtnCrearEnRuta_Click

            btnRestaurar = New Button() With {
                .Text = "Restaurar Copia Seleccionada",
                .Height = 34,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Padding = New Padding(12, 0, 12, 0),
                .Margin = New Padding(0, 0, 8, 0)
            }
            UITheme.StyleButton(btnRestaurar, "Danger")
            AddHandler btnRestaurar.Click, AddressOf BtnRestaurar_Click

            btnAbrirCarpeta = New Button() With {
                .Text = "Abrir Carpeta de Backups",
                .Height = 34,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Padding = New Padding(10, 0, 10, 0),
                .Margin = New Padding(0, 0, 8, 0)
            }
            UITheme.StyleButton(btnAbrirCarpeta, "Secondary")
            AddHandler btnAbrirCarpeta.Click, AddressOf BtnAbrirCarpeta_Click

            btnRefrescar = New Button() With {
                .Text = "Refrescar",
                .Height = 34,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Padding = New Padding(10, 0, 10, 0)
            }
            UITheme.StyleButton(btnRefrescar, "Secondary")
            AddHandler btnRefrescar.Click, Sub() LoadDashboardBackup()

            flpToolbar.Controls.AddRange({btnCrearBackup, btnCrearEnRuta, btnRestaurar, btnAbrirCarpeta, btnRefrescar})
            pnlToolbar.Controls.Add(flpToolbar)

            ' Grilla principal de respaldos
            Dim pnlGridContainer As New Panel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(15, 12, 15, 15),
                .BackColor = UITheme.ColorBackground
            }

            dgvBackups = New DataGridView() With {.Dock = DockStyle.Fill}
            UITheme.StyleDataGrid(dgvBackups)
            ConfigurarColumnasGrid()
            AddHandler dgvBackups.CellDoubleClick, Sub(s, e)
                                                       If e.RowIndex >= 0 Then BtnRestaurar_Click(s, e)
                                                   End Sub

            pnlGridContainer.Controls.Add(dgvBackups)

            ' Jerarquía de controles
            Me.Controls.Add(pnlGridContainer)
            Me.Controls.Add(pnlToolbar)
            Me.Controls.Add(pnlCards)
            Me.Controls.Add(pnlHeader)
        End Sub

        Private Sub ConfigurarColumnasGrid()
            dgvBackups.Columns.Clear()

            dgvBackups.Columns.Add("Nombre", "Archivo de Respaldo")
            dgvBackups.Columns("Nombre").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

            dgvBackups.Columns.Add("Fecha", "Fecha y Hora de Creación")
            dgvBackups.Columns("Fecha").Width = 170

            dgvBackups.Columns.Add("Tamano", "Tamaño")
            dgvBackups.Columns("Tamano").Width = 100
            dgvBackups.Columns("Tamano").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvBackups.Columns("Tamano").DefaultCellStyle.Font = UITheme.FontBold

            dgvBackups.Columns.Add("Ruta", "Ruta de Archivo")
            dgvBackups.Columns("Ruta").Width = 320
        End Sub

        Private Sub LoadDashboardBackup()
            ' Cargar información de la base de datos en vivo
            Dim info = backupService.ObtenerInfoBaseDatosActual()
            lblDbNombre.Text = info.TamanoFormateado
            CType(lblDbNombre.Parent.Controls(3), Label).Text = $"{info.CantidadTablas} tablas | {info.CantidadVentas} tickets"

            ' Cargar listado de copias disponibles
            Dim backups = backupService.GetBackups()
            dgvBackups.Rows.Clear()

            lblCantBackups.Text = backups.Count.ToString()

            If backups.Count > 0 Then
                lblUltimoBackup.Text = backups(0).FechaCreacion.ToString("dd/MM HH:mm")
                CType(lblUltimoBackup.Parent.Controls(3), Label).Text = backups(0).NombreArchivo
            Else
                lblUltimoBackup.Text = "Sin copias"
                CType(lblUltimoBackup.Parent.Controls(3), Label).Text = "Cree su primer respaldo"
            End If

            For Each b In backups
                dgvBackups.Rows.Add(b.NombreArchivo, b.FechaCreacion.ToString("dd/MM/yyyy HH:mm:ss"), b.TamanoFormateado, b.RutaCompleta)
            Next
        End Sub

        Private Sub BtnCrearBackup_Click(sender As Object, e As EventArgs)
            Dim errMsg As String = ""
            Dim rutaCreada = backupService.CrearBackup(errMsg)

            If Not String.IsNullOrEmpty(rutaCreada) Then
                Dim nombre = Path.GetFileName(rutaCreada)
                UITheme.ShowToast(Me, $"Respaldo creado exitosamente: {nombre}", "Success", 3500)
                LoadDashboardBackup()
            Else
                MessageBox.Show("Error al generar el respaldo: " & errMsg, "Error de Backup", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End Sub

        Private Sub BtnCrearEnRuta_Click(sender As Object, e As EventArgs)
            Using sfd As New SaveFileDialog()
                sfd.Title = "Guardar Copia de Seguridad de Base de Datos"
                sfd.Filter = "Base de Datos SQLite (*.db)|*.db|Todos los archivos (*.*)|*.*"
                sfd.FileName = $"backup_gestion_comercial_{DateTime.Now:yyyyMMdd_HHmmss}.db"

                If sfd.ShowDialog(Me) = DialogResult.OK Then
                    Dim errMsg As String = ""
                    Dim rutaCreada = backupService.CrearBackup(sfd.FileName, errMsg)

                    If Not String.IsNullOrEmpty(rutaCreada) Then
                        MessageBox.Show($"Copia de seguridad guardada exitosamente en:" & vbCrLf & sfd.FileName, "Backup Completado", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        LoadDashboardBackup()
                    Else
                        MessageBox.Show("Error al guardar el respaldo: " & errMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End If
            End Using
        End Sub

        Private Sub BtnRestaurar_Click(sender As Object, e As EventArgs)
            If dgvBackups.CurrentRow Is Nothing Then
                MessageBox.Show("Por favor seleccione un archivo de respaldo de la lista para restaurar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim rutaArchivo = dgvBackups.CurrentRow.Cells("Ruta").Value.ToString()
            Dim nombreArchivo = dgvBackups.CurrentRow.Cells("Nombre").Value.ToString()

            Dim advertencia = $"⚠️ ADVERTENCIA CRÍTICA:" & vbCrLf & vbCrLf &
                              $"Está a punto de restaurar la base de datos activa desde:" & vbCrLf &
                              $"'{nombreArchivo}'" & vbCrLf & vbCrLf &
                              $"Esta operación reemplazará todas las ventas, clientes y catálogo actuales por el estado del respaldo." & vbCrLf &
                              $"¿Confirma que desea proceder con la restauración?"

            Dim confirm = MessageBox.Show(advertencia, "Confirmar Restauración de Base de Datos", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)
            If confirm = DialogResult.Yes Then
                Dim errMsg As String = ""
                If backupService.RestaurarBackup(rutaArchivo, errMsg) Then
                    MessageBox.Show("La base de datos fue restaurada con éxito desde el respaldo seleccionado." & vbCrLf & vbCrLf &
                                    "Se actualizaron todas las tablas e información contable.", "Restauración Exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    LoadDashboardBackup()
                Else
                    MessageBox.Show("Error al restaurar la base de datos: " & errMsg, "Error de Restauración", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        End Sub

        Private Sub BtnAbrirCarpeta_Click(sender As Object, e As EventArgs)
            Try
                Dim dir = backupService.GetDirectorioBackups()
                If Not Directory.Exists(dir) Then
                    Directory.CreateDirectory(dir)
                End If
                Process.Start(New ProcessStartInfo("explorer.exe", dir) With {.UseShellExecute = True})
            Catch ex As Exception
                MessageBox.Show("No se pudo abrir la carpeta: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

    End Class
End Namespace
