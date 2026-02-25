Imports System.Drawing
Imports System.Windows.Forms
Imports ExcelDna.Integration
Imports Microsoft.Office.Interop

Friend Class DropButtonOverlay
  Inherits NativeWindow
  Implements IDisposable

  Private ReadOnly _onClick As Action
  Private _hWnd As IntPtr

  Public Sub New(x As Integer, y As Integer, w As Integer, h As Integer, onClick As Action)
    _onClick = onClick

    Dim exStyle = WS_EX_TOOLWINDOW Or WS_EX_NOACTIVATE ' Or WS_EX_TOPMOST
    Dim style = WS_POPUP Or WS_VISIBLE

    ' Register class no use of STATIC
    Dim className = RegisterOverlayClass("DropButtonOverlayClass")
    ' Get excel parent window handle
    Dim parentHwnd As IntPtr = CType(ExcelDnaUtil.Application, Excel.Application).Hwnd

    _hWnd = CreateWindowEx(
        exStyle,
        className,   ' Correct window class
        "",
        style,
        x, y, w, h,
        parentHwnd, 'IntPtr.Zero, '
        IntPtr.Zero,
        IntPtr.Zero,
        IntPtr.Zero)

    If _hWnd = IntPtr.Zero Then
      Throw New InvalidOperationException("Failed to create overlay button window.")
    End If

    Me.AssignHandle(_hWnd)
  End Sub

  Protected Overrides Sub WndProc(ByRef m As Message)
    Select Case m.Msg
      Case WM_PAINT
        DrawButton()
        m.Result = IntPtr.Zero
        Return
      Case WM_LBUTTONDOWN
        m.Result = IntPtr.Zero
        Return
      Case WM_LBUTTONUP
        ExcelAsyncUtil.QueueAsMacro(
        Sub()
          _onClick?.Invoke()
        End Sub)
        m.Result = IntPtr.Zero
        Return
      Case Else
        m.Result = DefWindowProc(m.HWnd, m.Msg, m.WParam, m.LParam)
        Return
    End Select
  End Sub
  Private Function ExcelIsEditing() As Boolean
    Try
      Dim xl = CType(ExcelDnaUtil.Application, Excel.Application)
      Dim editMenu = xl.CommandBars("Worksheet Menu Bar").Controls("Edit")
      Return Not editMenu.Enabled
    Catch
      Return False
    End Try
  End Function

  Private Sub DrawButton()
    ' use BeginPaint / EndPaint (prevents flicker)
    Dim ps As New PAINTSTRUCT()
    Dim hdc = BeginPaint(_hWnd, ps)
    If hdc = IntPtr.Zero Then Exit Sub

    Dim rc As RECT
    GetClientRect(_hWnd, rc)
    Dim size = Math.Min(rc.Right - rc.Left, rc.Bottom - rc.Top)

    Using g = Graphics.FromHdc(hdc)
      g.Clear(Color.White)

      ' Border
      Using pen As New Pen(Color.Black)
        g.DrawRectangle(pen, 0, 0, size - 1, size - 1)
      End Using

      ' Font size = 30–35% of button height (matches Excel)
      Dim fontSize As Single = CSng(size * 0.33F)

      Using f As New Font("Segoe UI", fontSize)
        TextRenderer.DrawText(
            g,
            "▼",
            f,
            New Rectangle(0, 0, size, size),
            Color.Black,
            TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
      End Using
    End Using


    EndPaint(_hWnd, ps)
  End Sub

  Public Sub Dispose() Implements IDisposable.Dispose
    If _hWnd <> IntPtr.Zero Then
      DestroyWindow(_hWnd)
      _hWnd = IntPtr.Zero
    End If
    Me.ReleaseHandle()
  End Sub

End Class
