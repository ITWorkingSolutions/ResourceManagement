Imports System.Runtime.InteropServices
Imports System.Drawing
Imports System.Windows.Forms

Friend Module NativeForm
  ' ============================
  ' Win32 constants
  ' ============================
  Friend Const WS_POPUP As Integer = &H80000000
  Friend Const WS_VISIBLE As Integer = &H10000000
  Friend Const WS_CHILD As Integer = &H40000000

  Friend Const WS_EX_TOOLWINDOW As Integer = &H80
  Friend Const WS_EX_TOPMOST As Integer = &H8
  Friend Const WS_EX_NOACTIVATE As Integer = &H8000000

  Friend Const WM_LBUTTONDOWN As Integer = &H201
  Friend Const WM_LBUTTONUP As Integer = &H202
  Friend Const WM_PAINT As Integer = &HF

  Friend Const WM_KILLFOCUS As Integer = &H8
  Friend Const WM_MOUSEWHEEL As Integer = &H20A
  Friend Const WM_VSCROLL As Integer = &H115
  Friend Const WM_HSCROLL As Integer = &H114
  Friend Const WS_VSCROLL As Integer = &H200000
  Friend Const WM_WINDOWPOSCHANGED As Integer = &H47
  Friend Const WM_MOUSEACTIVATE As Integer = &H21
  Friend Const MA_IGNORE As Integer = 1


  Friend Const WM_ENTERSIZEMOVE As Integer = &H231
  Friend Const WM_EXITSIZEMOVE As Integer = &H232

  <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
  Friend Function CreateWindowEx(
      exStyle As Integer,
      lpClassName As String,
      lpWindowName As String,
      style As Integer,
      x As Integer,
      y As Integer,
      nWidth As Integer,
      nHeight As Integer,
      hWndParent As IntPtr,
      hMenu As IntPtr,
      hInstance As IntPtr,
      lpParam As IntPtr) As IntPtr
  End Function

  ' ============================
  ' Win32 API declarations
  ' ============================

  <DllImport("user32.dll", SetLastError:=True)>
  Friend Function DestroyWindow(hWnd As IntPtr) As Boolean
  End Function

  <DllImport("user32.dll")>
  Friend Function GetDC(hWnd As IntPtr) As IntPtr
  End Function

  <DllImport("user32.dll")>
  Friend Function ReleaseDC(hWnd As IntPtr, hdc As IntPtr) As Integer
  End Function

  <DllImport("user32.dll")>
  Friend Function DefWindowProc(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
  End Function

#Region "Paint Functions"
  <DllImport("user32.dll")>
  Friend Function BeginPaint(hWnd As IntPtr, ByRef lpPaint As PAINTSTRUCT) As IntPtr
  End Function

  <DllImport("user32.dll")>
  Friend Function EndPaint(hWnd As IntPtr, ByRef lpPaint As PAINTSTRUCT) As Boolean
  End Function

  <StructLayout(LayoutKind.Sequential)>
  Friend Structure PAINTSTRUCT
    Public hdc As IntPtr
    Public fErase As Boolean
    Public rcPaint As Rectangle
    Public fRestore As Boolean
    Public fIncUpdate As Boolean
    <MarshalAs(UnmanagedType.ByValArray, SizeConst:=32)>
    Public rgbReserved As Byte()
  End Structure

#End Region
#Region "Window Class Registration"
  ' ============================
  ' Window class registration
  ' ============================
  <DllImport("user32.dll", CharSet:=CharSet.Unicode)>
  Friend Function RegisterClassEx(ByRef wc As WNDCLASSEX) As UShort
  End Function

  <DllImport("kernel32.dll")>
  Friend Function GetModuleHandle(lpModuleName As String) As IntPtr
  End Function

  <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
  Friend Structure WNDCLASSEX
    Public cbSize As UInteger
    Public style As UInteger
    Public lpfnWndProc As IntPtr
    Public cbClsExtra As Integer
    Public cbWndExtra As Integer
    Public hInstance As IntPtr
    Public hIcon As IntPtr
    Public hCursor As IntPtr
    Public hbrBackground As IntPtr
    Public lpszMenuName As String
    Public lpszClassName As String
    Public hIconSm As IntPtr
  End Structure

  ' ============================
  ' Delegate instance
  ' ============================

  Friend Delegate Function WndProcDelegate(
    hWnd As IntPtr,
    msg As UInteger,
    wParam As IntPtr,
    lParam As IntPtr
) As IntPtr

  ' ============================
  ' Generic WndProc (no drawing)
  ' ============================
  Private Function OverlayWndProc(
    hWnd As IntPtr,
    msg As UInteger,
    wParam As IntPtr,
    lParam As IntPtr
) As IntPtr

    Return DefWindowProc(hWnd, msg, wParam, lParam)
  End Function

  ' ============================
  ' Register overlay class
  ' ============================
  Friend _registeredClasses As New HashSet(Of String)
  Friend ReadOnly _wndProcDelegate As New WndProcDelegate(AddressOf OverlayWndProc)

  Friend Function RegisterOverlayClass(className As String) As String
    If _registeredClasses.Contains(className) Then
      Return className
    End If

    Dim wc As New WNDCLASSEX()
    wc.cbSize = CUInt(Marshal.SizeOf(wc))
    wc.lpfnWndProc = Marshal.GetFunctionPointerForDelegate(_wndProcDelegate)
    wc.hInstance = GetModuleHandle(Nothing)
    wc.lpszClassName = className

    Dim atom = RegisterClassEx(wc)
    If atom = 0 Then
      ' If the class already exists, Windows sets ERROR_CLASS_ALREADY_EXISTS
      Dim err = Marshal.GetLastWin32Error()
      If err = 1410 Then ' ERROR_CLASS_ALREADY_EXISTS
        _registeredClasses.Add(className)
        Return className
      End If

      Throw New Exception("Failed to register class. Win32 error: " & err)
    End If

    _registeredClasses.Add(className)
    Return className
  End Function

#End Region

End Module
