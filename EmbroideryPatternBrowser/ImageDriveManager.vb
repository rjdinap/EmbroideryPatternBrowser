' ImageDriveComboManager.vb
Imports System.IO
Imports System.Management
Imports System.ComponentModel

''' <summary>
''' Keeps a ComboBox populated with "Default" + all Fixed/Removable drives,
''' auto-refreshes on plug/unplug, and exposes a Shared resolver that swaps drive roots
''' for file paths when a non-Default drive is selected.
''' </summary>
Public Class ImageDriveManager
    Implements IDisposable

    ' ========= Shared (global) path-resolver state & API =========

    ''' <summary>
    ''' Nothing/empty = use stored paths as-is ("Default"). Otherwise like "E:\".
    ''' </summary>
    Public Shared Property DriveRootOverride As String
        Get
            Return _driveRootOverride
        End Get
        Private Set(value As String)
            Dim norm As String = NormalizeRoot(value)
            If String.Equals(_driveRootOverride, norm, StringComparison.OrdinalIgnoreCase) Then Exit Property
            _driveRootOverride = norm
            RaiseEvent SelectedDriveChanged(_driveRootOverride)
        End Set
    End Property
    Private Shared _driveRootOverride As String = Nothing

    ''' <summary>
    ''' Raised whenever the effective selected drive root changes (Nothing = Default).
    ''' </summary>
    Public Shared Event SelectedDriveChanged(root As String)

    ''' <summary>
    ''' Returns the effective path to use for reads/copies.
    ''' If a drive override is set, swaps the drive root of the stored path.
    ''' Optionally falls back to the original path if the swapped path does not exist.
    ''' </summary>
    Public Shared Function Resolve(storedFullPath As String,
                                   Optional fallbackToOriginalWhenMissing As Boolean = True) As String
        If String.IsNullOrWhiteSpace(storedFullPath) Then Return storedFullPath
        If String.IsNullOrEmpty(DriveRootOverride) Then Return storedFullPath

        ' Don't touch UNC or non-rooted paths
        Dim oldRoot = Path.GetPathRoot(storedFullPath)
        If String.IsNullOrEmpty(oldRoot) Then Return storedFullPath
        If oldRoot.StartsWith("\\") Then Return storedFullPath

        Dim candidate = ReplaceDriveRoot(storedFullPath, DriveRootOverride)

        If fallbackToOriginalWhenMissing Then
            Try
                If Not File.Exists(candidate) AndAlso File.Exists(storedFullPath) Then
                    Return storedFullPath
                End If
            Catch
                ' If IO throws (permissions/race), just return the candidate below
            End Try
        End If
        Return candidate
    End Function

    ''' <summary>
    ''' Replaces the drive root (e.g., "C:\" -> "E:\") and preserves the remainder.
    ''' </summary>
    Public Shared Function ReplaceDriveRoot(path As String, newRoot As String) As String
        If String.IsNullOrWhiteSpace(path) Then Return path
        If String.IsNullOrWhiteSpace(newRoot) Then Return path

        Dim oldRoot = System.IO.Path.GetPathRoot(path)
        If String.IsNullOrEmpty(oldRoot) Then Return path
        If oldRoot.StartsWith("\\") Then Return path ' leave UNC alone

        Dim normNew = NormalizeRoot(newRoot)
        Return normNew & path.Substring(oldRoot.Length)
    End Function

    Private Shared Function NormalizeRoot(root As String) As String
        If String.IsNullOrWhiteSpace(root) Then Return Nothing
        Dim r = root.Trim()
        If r.StartsWith("\\") Then Return Nothing ' UNC not supported as override
        If r.Length = 2 AndAlso r(1) = ":"c Then r &= "\"
        If Not r.EndsWith("\") Then r &= "\"
        Return r
    End Function

    ' ========= Instance: combo population + WMI watcher =========

    Private ReadOnly _uiHost As ISynchronizeInvoke   ' typically your Form
    Private ReadOnly _combo As ComboBox
    Private _watcher As ManagementEventWatcher

    ' Represent items in the combo (shows Label, holds Root path; Root = Nothing for "Default")
    Private Class DriveItem
        Public Property Root As String
        Public Property Label As String
        Public Overrides Function ToString() As String
            Return Label
        End Function
    End Class

    Public Sub New(uiHost As ISynchronizeInvoke, combo As ComboBox)
        If uiHost Is Nothing Then Throw New ArgumentNullException(NameOf(uiHost))
        If combo Is Nothing Then Throw New ArgumentNullException(NameOf(combo))
        _uiHost = uiHost
        _combo = combo
    End Sub

    ''' <summary>Initial fill and start listening for volume changes.</summary>
    Public Sub Start()
        RefreshCombo(preserveSelection:=False)
        AddHandler _combo.SelectedIndexChanged, AddressOf Combo_SelectedIndexChanged
        StartWatcher()
        ' Ensure Shared override matches the current selection on startup
        UpdateSharedOverrideFromCombo()
    End Sub

    ''' <summary>Stop listening (e.g., when closing the form).</summary>
    Public Sub [Stop]()
        RemoveHandler _combo.SelectedIndexChanged, AddressOf Combo_SelectedIndexChanged
        StopWatcher()
    End Sub

    ''' <summary>Manually force a refresh (optional).</summary>
    Public Sub RefreshCombo(Optional preserveSelection As Boolean = True)
        Dim previousRoot As String = Nothing
        Dim prevItem = TryCast(_combo.SelectedItem, DriveItem)
        If preserveSelection AndAlso prevItem IsNot Nothing Then previousRoot = prevItem.Root

        _combo.BeginUpdate()
        Try
            _combo.Items.Clear()

            ' Top "Default" item (selected by default)
            Dim defItem As New DriveItem With {.Root = Nothing, .Label = "Default"}
            _combo.Items.Add(defItem)

            ' Add Fixed + Removable drives
            For Each di As DriveInfo In DriveInfo.GetDrives()
                If di.DriveType = DriveType.Fixed OrElse di.DriveType = DriveType.Removable Then
                    _combo.Items.Add(New DriveItem With {
                        .Root = di.Name,
                        .Label = BuildDriveLabel(di)
                    })
                End If
            Next

            ' Re-select previous if it still exists; else select "Default"
            Dim selectIndex As Integer = 0
            If preserveSelection AndAlso previousRoot IsNot Nothing Then
                For i As Integer = 0 To _combo.Items.Count - 1
                    Dim it = TryCast(_combo.Items(i), DriveItem)
                    If it IsNot Nothing AndAlso String.Equals(it.Root, previousRoot, StringComparison.OrdinalIgnoreCase) Then
                        selectIndex = i
                        Exit For
                    End If
                Next
            End If
            _combo.SelectedIndex = selectIndex

        Finally
            _combo.EndUpdate()
        End Try

        ' Ensure Shared override reflects the (possibly changed) selection
        UpdateSharedOverrideFromCombo()
    End Sub

    Private Sub Combo_SelectedIndexChanged(sender As Object, e As EventArgs)
        UpdateSharedOverrideFromCombo()
    End Sub

    Private Sub UpdateSharedOverrideFromCombo()
        Dim root As String = Nothing
        Dim it = TryCast(_combo.SelectedItem, DriveItem)
        If it IsNot Nothing Then root = it.Root
        DriveRootOverride = root  ' updates Shared state + raises SelectedDriveChanged if changed
    End Sub

    Private Function BuildDriveLabel(di As DriveInfo) As String
        Dim kind As String = If(di.DriveType = DriveType.Removable, "Removable", "Fixed")
        Dim vol As String = Nothing
        Try
            If di.IsReady Then vol = di.VolumeLabel
        Catch
        End Try
        If Not String.IsNullOrWhiteSpace(vol) Then
            Return $"{di.Name}  ({vol})"
        Else
            Return $"{di.Name}"
        End If
    End Function

    ' --- WMI watcher (arrival/removal) ---
    Private Sub StartWatcher()
        Try
            Dim q As New WqlEventQuery("SELECT * FROM Win32_VolumeChangeEvent WHERE EventType = 2 OR EventType = 3")
            _watcher = New ManagementEventWatcher(q)
            AddHandler _watcher.EventArrived, AddressOf OnVolumeEvent
            _watcher.Start()
        Catch
            ' WMI may be restricted; manual RefreshCombo still works.
        End Try
    End Sub

    Private Sub StopWatcher()
        If _watcher Is Nothing Then Return
        Try
            RemoveHandler _watcher.EventArrived, AddressOf OnVolumeEvent
            _watcher.Stop()
        Catch
        Finally
            _watcher.Dispose()
            _watcher = Nothing
        End Try
    End Sub

    Private Sub OnVolumeEvent(sender As Object, e As EventArrivedEventArgs)
        Dim refresher As MethodInvoker = Sub() RefreshCombo(True)
        If _uiHost.InvokeRequired Then
            Try : _uiHost.BeginInvoke(refresher, Nothing) : Catch : End Try
        Else
            refresher()
        End If
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        [Stop]()
    End Sub
End Class
