' ARCHIVO: FrmAutorizacionSupervisor.vb
' PROPÓSITO: Diálogo modal reutilizable para autorizaciones de supervisor (Gerente / Administrador).
' Desacopla la interfaz gráfica de AuthService y aplica control anti fuerza bruta y auditoría del autorizante.

Imports System.Drawing
Imports System.Windows.Forms
Imports GestionComercial.Models
Imports GestionComercial.Services
Imports GestionComercial.UI

Namespace Forms
    Public Class FrmAutorizacionSupervisor
        Inherits Form

        Private ReadOnly _rolesPermitidos As String()
        Private ReadOnly _motivo As String
        Private txtPass As TextBox

        <System.ComponentModel.Browsable(False)>
        <System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)>
        Public Property SupervisorAutorizante As Usuario = Nothing

        Public Sub New(tituloHeader As String, etiquetaRol As String, motivo As String, rolesPermitidos As String())
            _rolesPermitidos = rolesPermitidos
            _motivo = motivo
            InitializeUI(tituloHeader, etiquetaRol)
        End Sub

        Private Sub InitializeUI(tituloHeader As String, etiquetaRol As String)
            Me.Text = "Autorización de Supervisor"
            Me.Size = New Size(430, 275)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.BackColor = Color.White

            Dim pnlTop As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 55,
                .BackColor = UITheme.ColorSecondary,
                .Padding = New Padding(15, 12, 15, 10)
            }
            Dim lblTitle As New Label() With {
                .Text = $"🔒 {tituloHeader.ToUpper()}",
                .Font = UITheme.FontSubheading,
                .ForeColor = Color.White,
                .AutoSize = True,
                .Location = New Point(12, 15)
            }
            pnlTop.Controls.Add(lblTitle)

            Dim lblMotivo As New Label() With {
                .Text = _motivo,
                .Font = UITheme.FontRegular,
                .ForeColor = UITheme.ColorDanger,
                .Location = New Point(20, 68),
                .Size = New Size(375, 42)
            }

            Dim lblPass As New Label() With {
                .Text = $"Contraseña de {etiquetaRol}:",
                .Font = UITheme.FontBold,
                .Location = New Point(20, 115),
                .AutoSize = True
            }

            txtPass = New TextBox() With {
                .Location = New Point(20, 138),
                .Size = New Size(375, 26),
                .UseSystemPasswordChar = True
            }
            UITheme.StyleTextBox(txtPass)

            Dim btnConfirmar As New Button() With {
                .Text = "Autorizar",
                .Location = New Point(185, 180),
                .Size = New Size(100, 34)
            }
            UITheme.StyleButton(btnConfirmar, "Primary")

            Dim btnCancelar As New Button() With {
                .Text = "Cancelar",
                .Location = New Point(295, 180),
                .Size = New Size(100, 34)
            }
            UITheme.StyleButton(btnCancelar, "Secondary")

            AddHandler btnConfirmar.Click, Sub() ProcesarConfirmacion(etiquetaRol)
            AddHandler btnCancelar.Click, Sub()
                                              Me.DialogResult = DialogResult.Cancel
                                              Me.Close()
                                          End Sub

            Me.Controls.AddRange({pnlTop, lblMotivo, lblPass, txtPass, btnConfirmar, btnCancelar})
            Me.AcceptButton = btnConfirmar
            Me.CancelButton = btnCancelar
        End Sub

        Private Sub ProcesarConfirmacion(etiquetaRol As String)
            If String.IsNullOrWhiteSpace(txtPass.Text) Then
                MessageBox.Show($"Ingrese la contraseña de {etiquetaRol}.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtPass.Focus()
                Return
            End If

            Dim errMsg As String = String.Empty
            Dim supervisor As Usuario = Nothing
            If AuthService.AutenticarSupervisorPorRoles(txtPass.Text, _rolesPermitidos, _motivo, supervisor, errMsg) Then
                SupervisorAutorizante = supervisor
                Me.DialogResult = DialogResult.OK
                Me.Close()
                Return
            End If

            MessageBox.Show(errMsg, "Acceso Denegado", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtPass.Clear()
            txtPass.Focus()
        End Sub

    End Class
End Namespace
