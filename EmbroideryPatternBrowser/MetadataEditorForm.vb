Imports System.Drawing
Imports System.Windows.Forms

Public Class MetadataEditorForm
    Inherits Form

    Private ReadOnly _txt As New TextBox()
    Private ReadOnly _btnOk As New Button()
    Private ReadOnly _btnCancel As New Button()

    Public Property MetadataText As String
        Get
            Return _txt.Text
        End Get
        Set(value As String)
            _txt.Text = value
        End Set
    End Property

    Public Sub New()
        Me.Text = "Edit Metadata"
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.StartPosition = FormStartPosition.CenterParent
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        Me.ShowInTaskbar = False
        Me.ClientSize = New Size(480, 260)
        Me.Font = New Font("Segoe UI", 9.0F)

        _txt.Multiline = True
        _txt.ScrollBars = ScrollBars.Vertical
        _txt.SetBounds(12, 12, 456, 200)

        _btnOk.Text = "OK"
        _btnOk.DialogResult = DialogResult.OK
        _btnOk.SetBounds(282, 220, 90, 28)

        _btnCancel.Text = "Cancel"
        _btnCancel.DialogResult = DialogResult.Cancel
        _btnCancel.SetBounds(378, 220, 90, 28)

        Me.Controls.AddRange(New Control() {_txt, _btnOk, _btnCancel})
        Me.AcceptButton = _btnOk
        Me.CancelButton = _btnCancel
    End Sub
End Class
