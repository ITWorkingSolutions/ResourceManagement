Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.Office.Interop

Module FormHelpers

  ' === Win32 API constants ===
  Private Const GWL_STYLE As Integer = -16
  Private Const LVS_NOCOLUMNHEADER As Integer = &H4000

  <DllImport("user32.dll", CharSet:=CharSet.Auto)>
  Private Function GetWindowLong(hWnd As IntPtr, nIndex As Integer) As Integer
  End Function

  <DllImport("user32.dll", CharSet:=CharSet.Auto)>
  Private Function SetWindowLong(hWnd As IntPtr, nIndex As Integer, dwNewLong As Integer) As Integer
  End Function

  ' Declare GetDpiForWindow (may not exist on older systems)
  <DllImport("user32.dll")>
  Friend Function GetDpiForWindow(hWnd As IntPtr) As UInteger
  End Function

  ' === Module-level state ===
  Private _excelRect As RECT
  Private _excelRectLoaded As Boolean

  ' ***********************************************************************************************
  '  Routine: InitializeExcelWindowBounds
  '  Purpose: Retrieves the Excel application window rectangle once and caches it for reuse.
  ' ***********************************************************************************************
  Friend Sub InitializeExcelWindowBounds()

    ' === Variable declarations ===
    Dim hwnd As IntPtr
    Dim ok As Boolean

    ' === Retrieve Excel window handle ===
    hwnd = ExcelDna.Integration.ExcelDnaUtil.WindowHandle

    ' === Load rectangle only once ===
    If Not _excelRectLoaded Then
      ok = GetWindowRect(hwnd, _excelRect)
      If ok Then
        _excelRectLoaded = True
      End If
    End If

  End Sub

  ' ***********************************************************************************************
  '  Routine: CenterFormOnExcel
  '  Purpose:
  '    Centers the supplied form relative to the Excel application window, while ensuring the form
  '    remains fully visible on the screen. Prevents the title bar from being positioned off-screen
  '    when Excel is near the top edge of the monitor.
  '
  '  Parameters:
  '    frm - The form to be positioned.
  '
  '  Returns:
  '    None
  '
  '  Notes:
  '    - Uses previously captured Excel window bounds (_excelRect).
  '    - Ensures the form's top edge is always >= 0 (visible title bar).
  '    - Optional clamping for left/right/bottom edges included for safety.
  ' ***********************************************************************************************
  Friend Sub CenterFormOnExcel(ByRef frm As Form)
    Try
      ' --- Normal execution ---

      ' === Ensure Excel bounds are loaded ===
      If Not _excelRectLoaded Then
        InitializeExcelWindowBounds()
      End If

      ' === Compute Excel window dimensions ===
      Dim excelWidth As Integer = _excelRect.Right - _excelRect.Left
      Dim excelHeight As Integer = _excelRect.Bottom - _excelRect.Top

      ' === Compute centered position ===
      Dim x As Integer = _excelRect.Left + (excelWidth - frm.Width) \ 2
      Dim y As Integer = _excelRect.Top + (excelHeight - frm.Height) \ 2

      ' === Clamp TOP so the title bar is always visible ===
      If y < 0 Then y = 0

      ' === Optional: Clamp LEFT so form never goes off-screen ===
      If x < 0 Then x = 0

      ' === Optional: Clamp RIGHT ===
      Dim screenWidth As Integer = Screen.PrimaryScreen.WorkingArea.Width
      If x + frm.Width > screenWidth Then
        x = screenWidth - frm.Width
      End If

      ' === Optional: Clamp BOTTOM ===
      Dim screenHeight As Integer = Screen.PrimaryScreen.WorkingArea.Height
      If y + frm.Height > screenHeight Then
        y = screenHeight - frm.Height
      End If

      ' === Apply manual positioning ===
      frm.StartPosition = FormStartPosition.Manual
      frm.Location = New System.Drawing.Point(x, y)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' --- Cleanup ---
      ' No disposable resources here.
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine:    GetScaleFactor
  ' Purpose:
  '   Determine the real DPI scaling factor of the monitor hosting the supplied control.
  '   Works for both top-level Forms and Task Pane UserControls.
  '
  ' Parameters:
  '   ctrl - Any WinForms control (Form or UserControl).
  '
  ' Returns:
  '   Single - The scale factor relative to 96 DPI (e.g., 1.0, 1.25, 1.5, 2.0).
  '
  ' Notes:
  '   - Excel is System-DPI aware and always starts at 96 DPI.
  '   - This routine retrieves the *actual* monitor DPI even when Excel reports 96.
  '   - Safe fallback ensures compatibility with older Windows versions.
  ' ==========================================================================================
  Public Function GetScaleFactor(ctrl As Control) As Single
    Try
      ' --- Normal execution ---
      Dim dpi As Single = 96.0F

      Try
        ' Try Windows 10+ API
        dpi = CSng(GetDpiForWindow(ctrl.Handle))
      Catch
        ' Fallback for older systems
        Using g = ctrl.CreateGraphics()
          dpi = g.DpiX
        End Using
      End Try

      Return dpi / 96.0F

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return 1.0F

    Finally
      ' --- Cleanup ---
    End Try
  End Function

  ' ==========================================================================================
  ' Routine:    ApplyDpiScaling
  ' Purpose:
  '   Entry point for DPI scaling. Retrieves scale factor and applies scaling to the form.
  '
  ' Parameters:
  '   form - The form to scale.
  '
  ' Returns:
  '   None
  '
  ' Notes:
  '   - Call this AFTER InitializeComponent and BEFORE showing the form.
  '   - Ensures consistent scaling even when Excel starts at high DPI.
  ' ==========================================================================================
  Public Sub ApplyDpiScaling(form As Form)
    Try
      ' --- Normal execution ---
      Dim scale As Single = GetScaleFactor(form)
      ScaleForm(form, scale)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' --- Cleanup ---
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine:    ApplyDpiScalingToTaskPane
  ' Purpose:
  '   Applies manual DPI scaling to a Task Pane UserControl hosted inside Excel.
  '   Unlike top-level WinForms forms, Task Pane controls must NOT scale the root
  '   container because Excel already bitmap-scales the pane. Only child controls
  '   are scaled to correct proportional shrinkage of multi-row controls.
  '
  ' Parameters:
  '   root - The UserControl that is hosted inside the Excel Task Pane.
  '
  ' Returns:
  '   None
  '
  ' Notes:
  '   - Excel is System-DPI aware and always starts at 96 DPI.
  '   - The Task Pane container is already scaled by Excel; scaling the root
  '     UserControl would cause double-scaling and oversized UI.
  '   - Fonts must NOT be scaled (Excel bitmap-scales them already).
  '   - Only child controls are scaled to restore correct proportions for
  '     TreeView, ListView, ListBox, TextBox (multiline), etc.
  '   - Call this AFTER InitializeComponent and BEFORE the Task Pane is shown.
  ' ==========================================================================================
  Public Sub ApplyDpiScalingToTaskPane(root As Control)
    Try
      ' --- Normal execution ---

      ' Determine the actual DPI scaling factor of the monitor Excel is on
      Dim scale As Single = GetScaleFactor(root)

      ' Do NOT scale the root control itself — Excel already scaled the pane container.
      ' Only scale the children to correct proportional shrinkage.
      For Each child As Control In root.Controls
        ScaleControl(child, scale)
      Next

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' --- Cleanup ---
      ' No disposable resources here.
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine:    ScaleForm
  ' Purpose:
  '   Apply proportional manual DPI scaling to a form and all child controls.
  '   This ensures correct sizing when Excel starts at high DPI (System-DPI mode).
  '
  ' Parameters:
  '   form  - The form to scale.
  '   scale - The scale factor (e.g., 1.0, 1.25, 1.5, 2.0).
  '
  ' Returns:
  '   None
  '
  ' Notes:
  '   - Must be called AFTER InitializeComponent and BEFORE the form is shown.
  '   - Avoids WinForms autoscaling, which is unreliable inside Excel.
  ' ==========================================================================================
  Private Sub ScaleForm(form As Form, scale As Single)
    Try
      ' --- Normal execution ---
      If Math.Abs(scale - 1.0F) < 0.01F Then Exit Sub

      form.SuspendLayout()
      ScaleControl(form, scale)
      form.ResumeLayout(True)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' --- Cleanup ---
      ' No disposable resources here.
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine:    ScaleControl
  ' Purpose:
  '   Recursively scale a control's font, size, and position based on the DPI factor.
  '
  ' Parameters:
  '   ctrl  - The control to scale.
  '   scale - The scale factor.
  '
  ' Returns:
  '   None
  '
  ' Notes:
  '   - Called internally by ScaleForm.
  '   - Scales fonts, bounds, and child controls.
  ' ==========================================================================================
  Private Sub ScaleControl(ctrl As Control, scale As Single)
    Try
      ' --- Normal execution ---

      ' Scale font
      ' DO NOT SCALE FONT — Excel already bitmap-scales it
      'ctrl.Font = New System.Drawing.Font(ctrl.Font.FontFamily,
      '                       ctrl.Font.Size * scale,
      '                       ctrl.Font.Style)

      ' Scale position and size
      ctrl.Left = CInt(ctrl.Left * scale)
      ctrl.Top = CInt(ctrl.Top * scale)
      ctrl.Width = CInt(ctrl.Width * scale)
      ctrl.Height = CInt(ctrl.Height * scale)

      ' Recurse into children
      For Each child As Control In ctrl.Controls
        ScaleControl(child, scale)
      Next

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' --- Cleanup ---
      ' No disposable resources here.
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: ValidateComboBoxSelection
  ' Purpose: Ensure ComboBox text matches an item in the list; prevent exit if invalid.
  ' Parameters:
  '   cb - ComboBox to validate
  '   e  - CancelEventArgs from Validating event
  ' Returns:
  '   (None)
  ' Notes:
  '   - Only applies when DropDownStyle = DropDown.
  '   - Prevents leaving control if typed text is not a valid list entry.
  ' ==========================================================================================
  Friend Sub ValidateComboBoxSelection(cb As ComboBox, e As System.ComponentModel.CancelEventArgs)
    Try
      '=== Ignore validation if empty ===
      If cb.Text.Trim() = "" Then Exit Sub

      '=== Check for exact match ===
      Dim index As Integer = cb.FindStringExact(cb.Text)

      If index = -1 Then
        MessageBox.Show("Please select a valid value from the list.", "Invalid Entry")
        e.Cancel = True   '=== Prevent leaving the control ===
      End If

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: DisableListViewHeader
  ' Purpose:
  '   Suppress the ListView column header by setting the LVS_NOCOLUMNHEADER style.
  ' Parameters:
  '   lv - the ListView whose header should be hidden
  ' Returns:
  '   None
  ' Notes:
  '   - Should be called after the ListView handle is created and View = Details.
  '   - Style-based: more reliable than hiding the header window.
  ' ==========================================================================================
  Friend Sub DisableListViewHeader(lv As ListView)
    If lv Is Nothing OrElse lv.Handle = IntPtr.Zero Then Exit Sub

    Dim style = GetWindowLong(lv.Handle, GWL_STYLE)
    style = style Or LVS_NOCOLUMNHEADER
    SetWindowLong(lv.Handle, GWL_STYLE, style)
  End Sub

  ' ==========================================================================================
  ' Routine:    GetScaleFactorFromHwnd
  ' Purpose:
  '   Determine the real DPI scaling factor for a window given its HWND.
  '
  ' Parameters:
  '   hwnd - Handle to any top-level window (Excel, overlay, etc.)
  '
  ' Returns:
  '   Single - DPI scale factor relative to 96 DPI.
  '
  ' Notes:
  '   - Required for worksheet overlays because they are not WinForms controls.
  ' ==========================================================================================
  Public Function GetScaleFactorFromHwnd(hwnd As IntPtr) As Single
    Try
      ' --- Normal execution ---
      Dim dpi As UInteger = GetDpiForWindow(hwnd)
      Return CSng(dpi / 96.0F)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return 1.0F

    Finally
      ' --- Cleanup ---
    End Try
  End Function

  ' ==========================================================================================
  ' Routine:    ExcelCoordinatesAreScaled
  ' Purpose:
  '   Determines whether Excel's PointsToScreenPixelsX/Y are already returning DPI-scaled
  '   pixel coordinates for the active window, or if they are still in raw 96-DPI units.
  '   This is used to decide whether manual scaling of worksheet overlay coordinates is
  '   required for correct alignment on high-DPI displays.
  '
  ' Parameters:
  '   xl        - The Excel.Application instance hosting the worksheet.
  '   hwndExcel - The window handle (HWND) of the Excel main window.
  '
  ' Returns:
  '   Boolean - True if Excel's pixel coordinates are already DPI-scaled and match the
  '             monitor DPI; False if they are raw 96-DPI coordinates and must be scaled.
  '
  ' Notes:
  '   - Excel is System-DPI aware and internally assumes 96 DPI.
  '   - On some configurations, PointsToScreenPixelsX/Y return bitmap-scaled coordinates;
  '     on others, they return raw coordinates that must be manually scaled.
  '   - This routine compares the expected DPI-scaled width of a column against the
  '     reported pixel width to infer which mode Excel is using.
  ' ==========================================================================================
  Private Function ExcelCoordinatesAreScaled(xl As Excel.Application,
                                             hwndExcel As IntPtr) As Boolean
    Try
      ' --- Normal execution ---

      Dim dpiScale As Single = GetScaleFactorFromHwnd(hwndExcel)

      ' Use the width of column A in points as a reference.
      Dim colWidthPoints As Double = xl.ActiveSheet.Columns(1).Width

      ' Get pixel coordinates for 0 and column width.
      Dim leftPx As Integer = xl.ActiveWindow.ActivePane.PointsToScreenPixelsX(0)
      Dim rightPx As Integer = xl.ActiveWindow.ActivePane.PointsToScreenPixelsX(colWidthPoints)

      Dim excelWidthPx As Integer = rightPx - leftPx
      Dim expectedWidthPx As Integer = CInt(colWidthPoints * dpiScale)

      ' If Excel's width is close to the expected DPI-scaled width, treat it as scaled.
      Return Math.Abs(excelWidthPx - expectedWidthPx) < 3

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return True ' Safe default: assume coordinates are already scaled.

    Finally
      ' --- Cleanup ---
      ' No disposable resources here.
    End Try
  End Function
End Module
