Imports System.Windows.Forms

Friend Class WindowWrapper
  Implements IWin32Window

  Private ReadOnly _hwnd As IntPtr

  Friend Sub New(handle As IntPtr)
    _hwnd = handle
  End Sub

  Friend ReadOnly Property Handle As IntPtr Implements IWin32Window.Handle
    Get
      Return _hwnd
    End Get
  End Property
End Class
