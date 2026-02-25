Imports System.Windows.Forms

Public Class ExcelWindowHook
  Inherits NativeWindow

  Public Sub New(hwnd As IntPtr)
    Me.AssignHandle(hwnd)
  End Sub

  Protected Overrides Sub WndProc(ByRef m As Message)
    Select Case m.Msg

      Case WM_ENTERSIZEMOVE
        ClearActiveOverlays()
        m.Result = IntPtr.Zero
        Return
    End Select

    MyBase.WndProc(m)
  End Sub

End Class
