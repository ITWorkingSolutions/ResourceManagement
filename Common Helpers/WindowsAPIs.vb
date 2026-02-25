Imports System.Runtime.InteropServices

Module WindowsAPIs
  ' ==========================================================================================
  ' Win32 API: SetWindowPos
  ' Purpose:
  '   Moves, resizes, and/or reorders a window. Used here to reposition the overlay button
  '   without activating it or changing Z‑order.
  ' ==========================================================================================
  <DllImport("user32.dll", SetLastError:=True)>
  Friend Function SetWindowPos(hWnd As IntPtr, hWndInsertAfter As IntPtr, X As Integer,
                               Y As Integer, cx As Integer, cy As Integer, uFlags As UInteger) As Boolean
  End Function

  ' --- SetWindowPos flags ---
  Friend Const SWP_NOSIZE As UInteger = &H1UI
  Friend Const SWP_NOMOVE As UInteger = &H2UI
  Friend Const SWP_NOZORDER As UInteger = &H4UI
  Friend Const SWP_NOACTIVATE As UInteger = &H10UI
  Friend Const SWP_SHOWWINDOW As UInteger = &H40UI

  Friend Const SB_LINEUP As Integer = 0
  Friend Const SB_LINEDOWN As Integer = 1
  Friend Const SB_PAGEUP As Integer = 2
  Friend Const SB_PAGEDOWN As Integer = 3
  Friend Const SB_THUMBTRACK As Integer = 5
  <StructLayout(LayoutKind.Sequential)>
  Friend Structure RECT
    Public Left As Integer
    Public Top As Integer
    Public Right As Integer
    Public Bottom As Integer
  End Structure

  'Friend Declare Function SetWindowPos Lib "user32.dll" (
  '  hWnd As IntPtr,
  '  hWndInsertAfter As IntPtr,
  '  X As Integer,
  '  Y As Integer,
  '  cx As Integer,
  '  cy As Integer,
  '  uFlags As UInteger) As Boolean

  Friend Declare Function GetClientRect Lib "user32.dll" (hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
  Friend Declare Function InvalidateRect Lib "user32.dll" (hWnd As IntPtr, lpRect As IntPtr, bErase As Boolean) As Boolean
  Friend Declare Function SetScrollInfo Lib "user32.dll" (hWnd As IntPtr, nBar As Integer, ByRef lpsi As SCROLLINFO, redraw As Boolean) As Integer
  Friend Declare Function GetScrollInfo Lib "user32.dll" (hWnd As IntPtr, nBar As Integer, ByRef lpsi As SCROLLINFO) As Integer
  Friend Structure POINT
    Public X As Integer
    Public Y As Integer
  End Structure

  Friend Declare Function ScreenToClient Lib "user32.dll" (hWnd As IntPtr, ByRef lpPoint As POINT) As Boolean
  Friend Const SB_VERT As Integer = 1

  Friend Declare Function ClientToScreen Lib "user32.dll" (hWnd As IntPtr, ByRef lpPoint As POINT) As Boolean

  <Runtime.InteropServices.StructLayout(Runtime.InteropServices.LayoutKind.Sequential)>
  Friend Structure SCROLLINFO
    Public cbSize As Integer
    Public fMask As Integer
    Public nMin As Integer
    Public nMax As Integer
    Public nPage As Integer
    Public nPos As Integer
    Public nTrackPos As Integer
  End Structure

  Friend Const SIF_RANGE As Integer = &H1
  Friend Const SIF_PAGE As Integer = &H2
  Friend Const SIF_POS As Integer = &H4
  Friend Const SIF_TRACKPOS As Integer = &H10
  Friend Const SIF_ALL As Integer = SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS

  <DllImport("user32.dll")>
  Friend Function GetWindowRect(hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
  End Function
End Module
