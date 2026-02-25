Imports System.Windows.Forms
Friend Module InputBoxHelper

  Private hook As IntPtr = IntPtr.Zero
  Private ownerHandle As IntPtr

  Private Const WH_CBT As Integer = 5
  Private Const HCBT_ACTIVATE As Integer = 5

  Private Delegate Function HookProc(nCode As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr

  Friend Function ShowExcelInputBox(owner As IWin32Window,
                                    app As Microsoft.Office.Interop.Excel.Application,
                                    prompt As String,
                                    title As String,
                                    defaultValue As String,
                                    inputType As Integer) As Object

    ' store owner for centering
    ownerHandle = owner.Handle

    ' install CBT hook for this thread
    hook = SetWindowsHookEx(WH_CBT, AddressOf CBTProc, IntPtr.Zero, GetCurrentThreadId())

    ' call Excel InputBox (returns Object: Range or Boolean False on cancel)
    Return app.InputBox(Prompt:=prompt, Title:=title, Default:=defaultValue, Type:=inputType)
  End Function

  Private Function CBTProc(nCode As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
    If nCode = HCBT_ACTIVATE Then
      Try
        CenterWindow(wParam)
      Catch
      End Try
      ' one-shot: remove hook
      UnhookWindowsHookEx(hook)
    End If
    Return IntPtr.Zero
  End Function

  Private Sub CenterWindow(hWnd As IntPtr)
    Dim rectOwner As RECT
    Dim rectDialog As RECT

    GetWindowRect(ownerHandle, rectOwner)
    GetWindowRect(hWnd, rectDialog)

    Dim x = rectOwner.Left + (rectOwner.Width - rectDialog.Width) \ 2
    Dim y = rectOwner.Top + (rectOwner.Height - rectDialog.Height) \ 2

    MoveWindow(hWnd, x, y, rectDialog.Width, rectDialog.Height, True)
  End Sub

  <System.Runtime.InteropServices.DllImport("user32.dll")>
  Private Function SetWindowsHookEx(idHook As Integer, lpfn As HookProc, hMod As IntPtr, dwThreadId As UInteger) As IntPtr
  End Function

  <System.Runtime.InteropServices.DllImport("user32.dll")>
  Private Function UnhookWindowsHookEx(hhk As IntPtr) As Boolean
  End Function

  <System.Runtime.InteropServices.DllImport("kernel32.dll")>
  Private Function GetCurrentThreadId() As UInteger
  End Function

  <System.Runtime.InteropServices.DllImport("user32.dll")>
  Private Function GetWindowRect(hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
  End Function

  <System.Runtime.InteropServices.DllImport("user32.dll")>
  Private Function MoveWindow(hWnd As IntPtr, X As Integer, Y As Integer, nWidth As Integer, nHeight As Integer, bRepaint As Boolean) As Boolean
  End Function

  Private Structure RECT
    Friend Left As Integer
    Friend Top As Integer
    Friend Right As Integer
    Friend Bottom As Integer
    Friend ReadOnly Property Width As Integer
      Get
        Return Right - Left
      End Get
    End Property
    Friend ReadOnly Property Height As Integer
      Get
        Return Bottom - Top
      End Get
    End Property
  End Structure

End Module