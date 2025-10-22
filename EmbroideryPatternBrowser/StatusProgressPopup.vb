Imports System.ComponentModel


'StatusProgress.ShowPopup(status:="Loading big file…", indeterminate:=True)
'StatusProgress.SetIndeterminate(False)
'StatusProgress.SetProgress(0, maximum:=100)
'StatusProgress.SetStatus($"Processing item {i} of 100")
'StatusProgress.SetProgress(i)
'StatusProgress.ClosePopup()

' A tiny, modeless popup window for status + progress that the user cannot close.
Public Class StatusProgressPopup
    Inherits Form

    Private ReadOnly lblStatus As Label
    Private ReadOnly bar As ProgressBar
    Private allowProgrammaticClose As Boolean = False

    Public Sub New()
        ' ---- Form chrome ----
        Me.Text = "Working…"
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.ControlBox = False              ' Hides Close/Minimize/Maximize buttons.
        Me.ShowInTaskbar = False
        Me.TopMost = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Size = New Drawing.Size(360, 120)
        Me.AutoScaleMode = AutoScaleMode.Dpi
        Me.DoubleBuffered = True

        ' ---- Layout panel ----
        Dim root As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 2,
            .Padding = New Padding(12),
            .AutoSize = False
        }
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 55.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 45.0F))
        Me.Controls.Add(root)

        ' ---- Status label ----
        lblStatus = New Label() With {
            .Dock = DockStyle.Fill,
            .AutoEllipsis = True,
            .Text = "Starting…",
            .TextAlign = Drawing.ContentAlignment.MiddleLeft
        }
        root.Controls.Add(lblStatus, 0, 0)

        ' ---- Progress bar ----
        bar = New ProgressBar() With {
            .Dock = DockStyle.Fill,
            .Minimum = 0,
            .Maximum = 100,
            .Style = ProgressBarStyle.Marquee, ' Default to indeterminate.
            .MarqueeAnimationSpeed = 25
        }
        root.Controls.Add(bar, 0, 1)
    End Sub

    ' --- Prevent user from closing the window manually ---
    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        If Not allowProgrammaticClose AndAlso e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
            Return
        End If
        MyBase.OnFormClosing(e)
    End Sub

    ' --- Safe programmatic close (from any thread) ---
    Public Sub CloseByCode()
        If Me.IsDisposed Then Return
        If Me.InvokeRequired Then
            Me.BeginInvoke(CType(Sub() CloseByCode(), MethodInvoker))
            Return
        End If
        allowProgrammaticClose = True
        Me.Close()
        Me.Dispose()
    End Sub

    ' --- Open (modeless). Optionally specify an owner window. ---
    Public Sub Open(Optional owner As IWin32Window = Nothing)
        If Me.InvokeRequired Then
            Me.BeginInvoke(CType(Sub() Open(owner), MethodInvoker))
            Return
        End If
        If owner Is Nothing Then
            Me.Show()      ' modeless
        Else
            Me.Show(owner) ' modeless with owner
        End If
        Me.BringToFront()
    End Sub

    ' --- Set status text (thread-safe) ---
    Public Sub UpdateStatus(message As String)
        If Me.IsDisposed Then Return
        If Me.InvokeRequired Then
            Me.BeginInvoke(CType(Sub() UpdateStatus(message), MethodInvoker))
            Return
        End If
        lblStatus.Text = If(String.IsNullOrEmpty(message), "", message)
        lblStatus.Refresh()
    End Sub

    ' --- Switch between indeterminate (marquee) and determinate modes (thread-safe) ---
    Public Sub SetIndeterminate(isIndeterminate As Boolean)
        If Me.IsDisposed Then Return
        If Me.InvokeRequired Then
            Me.BeginInvoke(CType(Sub() SetIndeterminate(isIndeterminate), MethodInvoker))
            Return
        End If
        If isIndeterminate Then
            bar.Style = ProgressBarStyle.Marquee
            bar.MarqueeAnimationSpeed = 25
        Else
            bar.MarqueeAnimationSpeed = 0
            bar.Style = ProgressBarStyle.Blocks
        End If
    End Sub

    ' --- Set progress value (and optionally maximum) in determinate mode (thread-safe) ---
    Public Sub SetProgress(current As Integer, Optional maximum As Integer? = Nothing)
        If Me.IsDisposed Then Return
        If Me.InvokeRequired Then
            Me.BeginInvoke(CType(Sub() SetProgress(current, maximum), MethodInvoker))
            Return
        End If
        If bar.Style <> ProgressBarStyle.Blocks Then
            bar.MarqueeAnimationSpeed = 0
            bar.Style = ProgressBarStyle.Blocks
        End If
        If maximum.HasValue AndAlso maximum.Value > 0 Then
            bar.Maximum = maximum.Value
        End If
        If current < bar.Minimum Then current = bar.Minimum
        If current > bar.Maximum Then current = bar.Maximum
        bar.Value = current
    End Sub

    ' --- Optional convenience: bump progress by delta ---
    Public Sub Increment(delta As Integer)
        If Me.IsDisposed Then Return
        If Me.InvokeRequired Then
            Me.BeginInvoke(CType(Sub() Increment(delta), MethodInvoker))
            Return
        End If
        If bar.Style <> ProgressBarStyle.Blocks Then
            bar.MarqueeAnimationSpeed = 0
            bar.Style = ProgressBarStyle.Blocks
        End If
        Dim nextVal = Math.Max(bar.Minimum, Math.Min(bar.Maximum, bar.Value + delta))
        bar.Value = nextVal
    End Sub
End Class

' --- Simple helper for one-liner use from anywhere ---
Public Module StatusProgress
    <EditorBrowsable(EditorBrowsableState.Never)>
    Private popupRef As StatusProgressPopup

    ' Show (or reuse) the popup.
    Public Sub ShowPopup(Optional owner As IWin32Window = Nothing, Optional status As String = "Working…", Optional indeterminate As Boolean = True)
        If popupRef Is Nothing OrElse popupRef.IsDisposed Then
            popupRef = New StatusProgressPopup()
        End If
        popupRef.UpdateStatus(status)
        popupRef.SetIndeterminate(indeterminate)
        popupRef.Open(owner)
    End Sub

    ' Update status text.
    Public Sub SetStatus(text As String)
        If popupRef Is Nothing OrElse popupRef.IsDisposed Then Return
        popupRef.UpdateStatus(text)
    End Sub

    ' Switch to determinate & set progress.
    Public Sub SetProgress(current As Integer, Optional maximum As Integer? = Nothing)
        If popupRef Is Nothing OrElse popupRef.IsDisposed Then Return
        popupRef.SetProgress(current, maximum)
    End Sub

    ' Switch between indeterminate/determinate.
    Public Sub SetIndeterminate(isIndeterminate As Boolean)
        If popupRef Is Nothing OrElse popupRef.IsDisposed Then Return
        popupRef.SetIndeterminate(isIndeterminate)
    End Sub

    ' Close programmatically.
    Public Sub ClosePopup()
        If popupRef Is Nothing OrElse popupRef.IsDisposed Then Return
        popupRef.CloseByCode()
        popupRef = Nothing
    End Sub
End Module
