Imports System.Runtime.InteropServices
Imports System.Windows.Forms

' ---------------------------------------------------------------------------
'  MessageBoxHelper
'  Ensures MessageBox dialogs are centred on the Excel window.
'
'  Why this exists:
'    - MessageBox.Show(owner) does NOT centre on the owner window.
'    - Win32 centres MessageBox on the desktop by default.
'    - Excel add-ins often show dialogs on the wrong monitor or off-centre.
'
'  How it works:
'    - Installs a CBT (Computer-Based Training) hook for the current thread.
'    - The hook fires when the MessageBox is ACTIVATED.
'    - At that moment, we reposition the MessageBox relative to the owner.
'    - The hook is immediately removed (no drift, no leaks).
'
'  This is the standard, safe pattern used in Excel-DNA and VSTO add-ins.
' ---------------------------------------------------------------------------

Friend Module MessageBoxHelper

  ' Handle to the installed hook (so we can unhook it)
  Private hook As IntPtr = IntPtr.Zero

  ' Handle to the owner window (Excel or a form)
  Private ownerHandle As IntPtr

  ' -----------------------------------------------------------------------
  '  Friend entry point
  '  Use this instead of MessageBox.Show(...)
  ' -----------------------------------------------------------------------
  Friend Function Show(owner As IWin32Window,
                       text As String,
                       caption As String,
                       buttons As MessageBoxButtons,
                       icon As MessageBoxIcon) As DialogResult

    ' Store the owner handle for use inside the hook callback
    ownerHandle = owner.Handle

    ' Install a CBT hook for the current thread
    ' This hook will fire when the MessageBox is activated
    hook = SetWindowsHookEx(WH_CBT, AddressOf CBTProc, IntPtr.Zero, GetCurrentThreadId())

    ' Show the MessageBox normally
    ' The hook will reposition it once it appears
    Return MessageBox.Show(owner, text, caption, buttons, icon)
  End Function

  ' -----------------------------------------------------------------------
  '  Win32 constants
  ' -----------------------------------------------------------------------
  Private Const WH_CBT As Integer = 5
  Private Const HCBT_ACTIVATE As Integer = 5

  ' -----------------------------------------------------------------------
  '  Hook callback
  '  Called by Windows when certain window events occur.
  '  We only care about ACTIVATE for the MessageBox.
  ' -----------------------------------------------------------------------
  Private Function CBTProc(nCode As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr

    ' When the MessageBox is activated, reposition it
    If nCode = HCBT_ACTIVATE Then

      CenterWindow(wParam)

      ' Remove the hook immediately (one-shot behaviour)
      UnhookWindowsHookEx(hook)
    End If

    Return IntPtr.Zero
  End Function

  ' -----------------------------------------------------------------------
  '  Centers the MessageBox window (hWnd) relative to the owner window
  ' -----------------------------------------------------------------------
  Private Sub CenterWindow(hWnd As IntPtr)

    Dim rectOwner As RECT
    Dim rectDialog As RECT

    ' Get the bounding rectangles for owner and dialog
    GetWindowRect(ownerHandle, rectOwner)
    GetWindowRect(hWnd, rectDialog)

    ' Compute centered coordinates
    Dim x = rectOwner.Left + (rectOwner.Width - rectDialog.Width) \ 2
    Dim y = rectOwner.Top + (rectOwner.Height - rectDialog.Height) \ 2

    ' Move the MessageBox to the new position
    MoveWindow(hWnd, x, y, rectDialog.Width, rectDialog.Height, True)
  End Sub

  ' -----------------------------------------------------------------------
  '  Win32 API declarations
  ' -----------------------------------------------------------------------

  <DllImport("user32.dll")>
  Private Function SetWindowsHookEx(idHook As Integer,
                                    lpfn As HookProc,
                                    hMod As IntPtr,
                                    dwThreadId As UInteger) As IntPtr
  End Function

  <DllImport("user32.dll")>
  Private Function UnhookWindowsHookEx(hhk As IntPtr) As Boolean
  End Function

  <DllImport("kernel32.dll")>
  Private Function GetCurrentThreadId() As UInteger
  End Function

  <DllImport("user32.dll")>
  Private Function GetWindowRect(hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
  End Function

  <DllImport("user32.dll")>
  Private Function MoveWindow(hWnd As IntPtr,
                              X As Integer,
                              Y As Integer,
                              nWidth As Integer,
                              nHeight As Integer,
                              bRepaint As Boolean) As Boolean
  End Function

  ' Delegate type for the hook callback
  Private Delegate Function HookProc(nCode As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr

  ' RECT structure used by Win32
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
