Imports System.Drawing
Imports System.Windows.Forms
Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Friend Class DropListOverlay
  Inherits NativeWindow
  Implements IDisposable


  Private ReadOnly _options As List(Of String)
  Private ReadOnly _onSelect As Action(Of String)
  Private ReadOnly _itemHeight As Integer = 18
  Private _hWnd As IntPtr
  Private _scrollOffset As Integer = 0
  Private ReadOnly _listSelectType As String
  Private ReadOnly _selected As Boolean()

  ' ==========================================================================================
  ' Routine: New
  ' Purpose: Creates the overlay dropdown window with a native vertical scrollbar.
  ' Parameters:
  '   x, y, w - screen coordinates and width for the overlay window
  '   options - list of strings to display
  '   onSelect - callback invoked when an item is selected
  ' Returns:
  '   None
  ' Notes:
  '   Adds WS_VSCROLL to enable native scrollbars.
  ' ==========================================================================================
  Public Sub New(x As Integer, y As Integer, w As Integer, options As List(Of String),
                 listSelectType As String, preSelectedValues As List(Of String),
                 onSelect As Action(Of String))
    Try
      _options = options
      _onSelect = onSelect

      _listSelectType = listSelectType
      _selected = New Boolean(_options.Count - 1) {}

      If listSelectType = ExcelListSelectType.MultiSelect.ToString() Then
        For i = 0 To _options.Count - 1
          If preSelectedValues.Contains(_options(i)) Then
            _selected(i) = True
          End If
        Next
      End If

      ' Get excel parent window handle
      Dim parentHwnd As IntPtr = CType(ExcelDnaUtil.Application, Excel.Application).Hwnd
      Dim dpi = GetDpiForWindow(parentHwnd)
      Dim scale As Double = dpi / 96.0

      _itemHeight = CInt(20 * scale)   ' or whatever your base row height is
      Dim h = Math.Min(120 * scale, _options.Count * _itemHeight)

      Dim exStyle = WS_EX_TOOLWINDOW Or WS_EX_NOACTIVATE ' Or WS_EX_TOPMOST
      Dim style = WS_POPUP Or WS_VISIBLE Or WS_VSCROLL 'WS_POPUP WS_CHILD
      Dim className = RegisterOverlayClass("DropListOverlayClass")

      _hWnd = CreateWindowEx(
          exStyle,
          className,
          "",
          style,
          x, y, w, h,
          parentHwnd,  'IntPtr.Zero,
          IntPtr.Zero,
          IntPtr.Zero,
          IntPtr.Zero)

      If _hWnd = IntPtr.Zero Then
        Throw New InvalidOperationException("Failed to create overlay dropdown window.")
      End If

      Me.AssignHandle(_hWnd)
      UpdateScrollBar()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' No cleanup required
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: WndProc
  ' Purpose: Handles all window messages including paint, mouse, and scrollbar events.
  ' Parameters:
  '   m - Windows message structure
  ' Returns:
  '   None
  ' Notes:
  '   WM_VSCROLL is now handled for native scrollbar support.
  ' ==========================================================================================
  Protected Overrides Sub WndProc(ByRef m As Message)
    Try
      Select Case m.Msg
        Case WM_PAINT
          DrawList()
          m.Result = IntPtr.Zero
          Return
        Case WM_LBUTTONDOWN
          m.Result = IntPtr.Zero
          Return
        Case WM_LBUTTONUP
          Dim lParamCopy As IntPtr = m.LParam
          ExcelAsyncUtil.QueueAsMacro(Sub() HandleClick(lParamCopy))
          m.Result = IntPtr.Zero
          Return
        Case WM_MOUSEWHEEL
          HandleMouseWheel(ExtractWheelDelta(m.WParam))
          m.Result = IntPtr.Zero
          Return
        Case WM_VSCROLL
          HandleVScroll(m.WParam.ToInt32())
          m.Result = IntPtr.Zero
          Return
        Case WM_KILLFOCUS ' commit multi-select on focus loss
          If _listSelectType = ExcelListSelectType.MultiSelect.ToString() Then
            CommitSelection()
          End If
          Dispose()
          m.Result = IntPtr.Zero
          Return

        Case Else
          m.Result = DefWindowProc(m.HWnd, m.Msg, m.WParam, m.LParam)
          Return
      End Select

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' No cleanup required
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: CommitSelection
  ' Purpose: For a multi-select list, gathers all selected items and invokes the callback
  '          with a comma-separated string.
  ' Parameters:
  '   none
  ' Returns:
  '   none
  ' Notes:
  ' ==========================================================================================
  Private Sub CommitSelection()
    Dim values As New List(Of String)
    For i = 0 To _options.Count - 1
      If _selected(i) Then values.Add(_options(i))
    Next

    Dim finalValue = String.Join(", ", values)
    _onSelect?.Invoke(finalValue)
  End Sub


  ' ==========================================================================================
  ' Routine: ExtractWheelDelta
  ' Purpose: Safely extracts the wheel delta from WPARAM without overflow.
  ' Parameters:
  '   wParam - raw WPARAM value
  ' Returns:
  '   Integer - wheel delta (-120, +120, etc.)
  ' Notes:
  '   Avoids CShort overflow by manually sign-extending the 16-bit value.
  ' ==========================================================================================
  Private Function ExtractWheelDelta(wParam As IntPtr) As Integer
    Try
      Dim raw As Long = wParam.ToInt64()
      Dim masked As Integer = CInt((raw >> 16) And &HFFFF)

      ' Manually interpret masked as signed 16-bit
      If masked >= &H8000 Then
        masked -= &H10000
      End If

      Return masked

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return 0

    Finally
    End Try
  End Function

  ' ==========================================================================================
  ' Routine: HandleVScroll
  ' Purpose: Handles native scrollbar events and updates scroll offset.
  ' Parameters:
  '   wParam - scroll command
  ' Returns:
  '   None
  ' Notes:
  '   Uses SCROLLINFO to track thumb position and range.
  ' ==========================================================================================
  Private Sub HandleVScroll(wParam As Integer)
    Try
      Dim si As New SCROLLINFO With {.cbSize = Runtime.InteropServices.Marshal.SizeOf(GetType(SCROLLINFO)), .fMask = SIF_ALL}
      GetScrollInfo(_hWnd, SB_VERT, si)

      Select Case (wParam And &HFFFF)
        Case SB_LINEUP
          _scrollOffset -= _itemHeight

        Case SB_LINEDOWN
          _scrollOffset += _itemHeight

        Case SB_PAGEUP
          _scrollOffset -= si.nPage * _itemHeight

        Case SB_PAGEDOWN
          _scrollOffset += si.nPage * _itemHeight

        Case SB_THUMBTRACK
          _scrollOffset = si.nTrackPos * _itemHeight
      End Select

      ClampScrollOffset()
      UpdateScrollBar()
      InvalidateRect(_hWnd, IntPtr.Zero, True)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: UpdateScrollBar
  ' Purpose: Updates the native scrollbar range and position.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' Notes:
  '   Must be called whenever scroll offset or window size changes.
  ' ==========================================================================================
  Private Sub UpdateScrollBar()
    Try
      Dim si As New SCROLLINFO With {.cbSize = Runtime.InteropServices.Marshal.SizeOf(GetType(SCROLLINFO)), .fMask = SIF_ALL}

      si.nMin = 0
      si.nMax = _options.Count - 1
      si.nPage = Math.Max(1, GetWindowHeight() \ _itemHeight)
      si.nPos = _scrollOffset \ _itemHeight

      SetScrollInfo(_hWnd, SB_VERT, si, True)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: ClampScrollOffset
  ' Purpose: Ensures scroll offset stays within valid bounds.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' Notes:
  '   Prevents negative or excessive scroll offsets.
  ' ==========================================================================================
  Private Sub ClampScrollOffset()
    Try
      Dim maxOffset As Integer = Math.Max(0, (_options.Count * _itemHeight) - GetWindowHeight())

      If _scrollOffset < 0 Then _scrollOffset = 0
      If _scrollOffset > maxOffset Then _scrollOffset = maxOffset

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: DrawList
  ' Purpose: Renders the dropdown list items using GDI.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' Notes:
  '   Uses BeginPaint/EndPaint and must be wrapped due to GDI handle operations.
  ' ==========================================================================================
  Private Sub DrawList()
    Try
      ' use BeginPaint / EndPaint (prevents flicker)
      Dim ps As New PAINTSTRUCT()
      Dim hdc = BeginPaint(_hWnd, ps)
      If hdc = IntPtr.Zero Then Exit Sub

      Using g = Graphics.FromHdc(hdc)
        g.Clear(Color.White)

        Using pen As New Pen(Color.Black)
          g.DrawRectangle(pen, 0, 0, g.VisibleClipBounds.Width - 1, g.VisibleClipBounds.Height - 1)
        End Using

        Using f As New System.Drawing.Font("Segoe UI", 9.0F)

          For i = 0 To _options.Count - 1
            Dim yPos As Integer = (i * _itemHeight) - _scrollOffset

            If yPos + _itemHeight >= 0 AndAlso yPos < g.VisibleClipBounds.Height Then

              If _listSelectType = ExcelListSelectType.MultiSelect.ToString() Then
                ' --- Draw checkbox ---
                Dim boxRect As New System.Drawing.Rectangle(4, yPos + 3, 12, 12)
                g.DrawRectangle(Pens.Black, boxRect)

                If _selected(i) Then
                  g.FillRectangle(Brushes.Black, boxRect)
                End If

                ' --- Draw text shifted right ---
                Dim textRect As New System.Drawing.Rectangle(20, yPos, CInt(g.VisibleClipBounds.Width) - 24, _itemHeight)
                TextRenderer.DrawText(g, _options(i), f, textRect, Color.Black, TextFormatFlags.Left Or TextFormatFlags.VerticalCenter)

              Else
                ' --- Single-select (existing behaviour) ---
                Dim r = New System.Drawing.Rectangle(2, yPos, CInt(g.VisibleClipBounds.Width) - 4, _itemHeight)
                TextRenderer.DrawText(g, _options(i), f, r, Color.Black, TextFormatFlags.Left Or TextFormatFlags.VerticalCenter)
              End If

            End If
          Next

        End Using
      End Using

      EndPaint(_hWnd, ps)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: HandleClick
  ' Purpose: Maps click coordinates to an item index and invokes the callback.
  ' Parameters:
  '   lParam - raw mouse coordinate data
  ' Returns:
  '   None
  ' Notes:
  '   Must account for scroll offset; unmanaged coordinate extraction can fail.
  ' ==========================================================================================
  Private Sub HandleClick(lParam As IntPtr)
    Try
      Dim x = CInt(lParam.ToInt32() And &HFFFF)
      Dim y = CInt((lParam.ToInt32() >> 16) And &HFFFF)

      Dim index = (y + _scrollOffset) \ _itemHeight

      If index >= 0 AndAlso index < _options.Count Then

        If _listSelectType = ExcelListSelectType.MultiSelect.ToString() Then
          ' --- MULTI-SELECT: toggle state, redraw, do NOT close ---
          _selected(index) = Not _selected(index)
          InvalidateRect(_hWnd, IntPtr.Zero, True)
          Return
        Else
          ' --- SINGLE-SELECT: behaviour ---
          Dim value = _options(index)
          _onSelect?.Invoke(value)
          Dispose()
          Return
        End If

      End If

      'If index >= 0 AndAlso index < _options.Count Then
      '  Dim value = _options(index)
      '  _onSelect?.Invoke(value)
      'End If

      'Dispose()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: HandleMouseWheel
  ' Purpose: Adjusts scroll offset based on wheel delta and invalidates the window.
  ' Parameters:
  '   delta - wheel movement
  ' Returns:
  '   None
  ' Notes:
  '   Uses Win32 InvalidateRect; must be wrapped.
  ' ==========================================================================================
  Private Sub HandleMouseWheel(delta As Integer)
    Try
      _scrollOffset -= delta

      ClampScrollOffset()
      UpdateScrollBar()
      InvalidateRect(_hWnd, IntPtr.Zero, True)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: GetWindowHeight
  ' Purpose: Retrieves the client height of the overlay window.
  ' Parameters:
  '   None
  ' Returns:
  '   Integer - height of the client area
  ' Notes:
  '   Calls GetClientRect; ANY Win32 API can fail.
  ' ==========================================================================================
  Private Function GetWindowHeight() As Integer
    Try
      Dim rect As RECT
      GetClientRect(_hWnd, rect)
      Return rect.Bottom - rect.Top

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return 0

    Finally
    End Try
  End Function

  ' ==========================================================================================
  ' Routine: Dispose
  ' Purpose: Destroys the window and releases the native handle.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' Notes:
  '   DestroyWindow and ReleaseHandle must be wrapped.
  ' ==========================================================================================
  Public Sub Dispose() Implements IDisposable.Dispose
    Try
      If _hWnd <> IntPtr.Zero Then
        DestroyWindow(_hWnd)
        _hWnd = IntPtr.Zero
      End If

      Me.ReleaseHandle()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
    End Try
  End Sub

End Class