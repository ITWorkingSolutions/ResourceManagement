Imports System.Windows.Forms
Imports System.Drawing

Namespace ResourceManagement.CustomControls
  Friend Class TabControlNoTabs
    Inherits TabControl

    ' Prevent Windows from drawing the tab headers
    Protected Overrides Sub WndProc(ByRef m As Message)
      Const TCM_ADJUSTRECT As Integer = &H1328

      If m.Msg = TCM_ADJUSTRECT Then
        ' Skip default processing so the tab strip is never drawn
        Return
      End If

      MyBase.WndProc(m)
    End Sub

    ' Fix the client area so pages fill the control correctly
    Public Overrides ReadOnly Property DisplayRectangle As Rectangle
      Get
        Dim rect = MyBase.DisplayRectangle
        ' Shift the client area upward to cover where tabs would be
        Return New Rectangle(rect.X, rect.Y - 20, rect.Width, rect.Height + 20)
      End Get
    End Property
  End Class
End Namespace
