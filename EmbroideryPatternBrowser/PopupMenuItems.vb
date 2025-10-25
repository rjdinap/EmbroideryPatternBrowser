' Simple wrapper for the grid's right-click context menu.
Public Class PopupMenuItems
    Implements IDisposable

    Private ReadOnly _cms As ContextMenuStrip
    Private ReadOnly _openDirAction As Action
    Private ReadOnly _editMetadataAction As Action
    Private ReadOnly _copyPathAction As Action
    Private _disposed As Boolean

    Public Sub New(openDirAction As Action, editMetadataAction As Action, copyPathAction As Action)
        _openDirAction = openDirAction
        _editMetadataAction = editMetadataAction
        _copyPathAction = copyPathAction

        _cms = New ContextMenuStrip() With {
            .ShowCheckMargin = False,
            .ShowImageMargin = False,
            .AllowDrop = False
        }

        Dim miOpen As New ToolStripMenuItem("Open directory location")
        AddHandler miOpen.Click, AddressOf OnOpenDirClick

        Dim miMeta As New ToolStripMenuItem("Add / Edit Metadata")
        AddHandler miMeta.Click, AddressOf OnEditMetadataClick

        Dim miCopy As New ToolStripMenuItem("Copy file path to clipboard")
        AddHandler miCopy.Click, AddressOf OnCopyPathClick

        _cms.Items.Add(miOpen)
        _cms.Items.Add(New ToolStripSeparator())
        _cms.Items.Add(miMeta)
        _cms.Items.Add(New ToolStripSeparator())
        _cms.Items.Add(miCopy)
    End Sub

    Private Sub OnOpenDirClick(sender As Object, e As EventArgs)
        If _openDirAction Is Nothing Then Return
        Try
            _openDirAction()
        Catch ex As Exception
            Logger.Error(ex.Message, ex.StackTrace.ToString)
        End Try
    End Sub

    Private Sub OnEditMetadataClick(sender As Object, e As EventArgs)
        If _editMetadataAction Is Nothing Then Return
        Try
            _editMetadataAction()
        Catch ex As Exception
            Logger.Error(ex.Message, ex.StackTrace.ToString)
        End Try
    End Sub

    Private Sub OnCopyPathClick(sender As Object, e As EventArgs)
        If _copyPathAction Is Nothing Then Return
        Try
            _copyPathAction()
        Catch ex As Exception
            Logger.Error(ex.Message, ex.StackTrace.ToString)
        End Try
    End Sub

    Public Sub Show(owner As Control, location As Point)
        If owner Is Nothing OrElse _disposed Then Return
        Try
            _cms.Show(owner, location)
        Catch ex As Exception
            Logger.Error(ex.Message, ex.StackTrace.ToString)
        End Try
    End Sub

    ' Clean up menu resources explicitly
    Public Sub Dispose() Implements IDisposable.Dispose
        If _disposed Then Return
        _disposed = True
        Try
            RemoveHandler _cms.Closing, Nothing ' no-op; placeholder if handlers are added later
            _cms.Dispose()
        Catch ex As Exception
            Logger.Error(ex.Message, ex.StackTrace.ToString)
        End Try
    End Sub
End Class
