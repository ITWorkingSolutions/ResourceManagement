Imports System.Windows.Forms
Namespace ResourceManagement.CustomControls
  Friend Class ExtendedListView
    Inherits ListView

    Private _columnPercents As Integer()

    Public Property ColumnPercents As Integer()
      Get
        Return _columnPercents
      End Get
      Set(value As Integer())
        _columnPercents = value
        ResizeColumns()
      End Set
    End Property

    Protected Overrides Sub OnResize(e As EventArgs)
      MyBase.OnResize(e)
      ResizeColumns()
    End Sub

    Protected Overrides Sub OnLayout(levent As LayoutEventArgs)
      MyBase.OnLayout(levent)
      ResizeColumns()
    End Sub

    Private Sub ResizeColumns()
      If _columnPercents Is Nothing OrElse Columns.Count = 0 Then Exit Sub
      If _columnPercents.Length <> Columns.Count Then Exit Sub

      Dim available As Integer = ClientSize.Width
      'If VerticalScrollBarVisible() Then
      '  available -= SystemInformation.VerticalScrollBarWidth
      'End If

      For i = 0 To Columns.Count - 1
        Columns(i).Width = CInt(available * _columnPercents(i) / 100)
      Next
    End Sub

    Private Function VerticalScrollBarVisible() As Boolean
      Const GWL_STYLE As Integer = -16
      Const WS_VSCROLL As Integer = &H200000
      Return (GetWindowLong(Me.Handle, GWL_STYLE) And WS_VSCROLL) <> 0
    End Function

    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function GetWindowLong(hWnd As IntPtr, nIndex As Integer) As Integer
    End Function

  End Class
End Namespace
