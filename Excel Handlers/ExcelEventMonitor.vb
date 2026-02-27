Imports System.Runtime.InteropServices
Imports ExcelDna.Integration
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel

Public Class ExcelEventMonitor
  ' Shared instance accessible from anywhere
  Friend Shared Instance As New ExcelEventMonitor()

  Private WithEvents App As Excel.Application
  ' ==========================================================================================
  ' Class: CellSnapshot
  ' Purpose:
  '   Immutable container for the minimal state required to detect paste operations by
  '   comparing BEFORE and AFTER snapshots.
  '
  ' Properties:
  '   Value   - Raw cell value (Value2)
  '   Formula - Cell formula string
  '   Guid    - GUID identity extracted from hidden RM_* name
  '
  ' Notes:
  '   - Does NOT store Range.ID because Range.ID is volatile and session-only.
  '   - Guid is the authoritative identity for paste detection.
  ' ==========================================================================================
  Friend Class CellSnapshot
    Public Property Value As Object
    Public Property Formula As String
    Public Property Guid As String
  End Class

  Friend Class RangeSnapshot
    Public Property Cells As Dictionary(Of String, CellSnapshot)
  End Class

  Private pasteInProgress As Boolean = False
  Private beforeSnapshot As RangeSnapshot = Nothing
  Private lastSnapshot As RangeSnapshot = Nothing

  ' --- Following are scroll detection and cell resize members 
  Private WithEvents scrollTimer As System.Windows.Forms.Timer
  Private lastRow As Long = -1
  Private lastCol As Long = -1
  ' Geometry tracking
  Friend lastCellTop As Double
  Friend lastCellLeft As Double
  Friend lastCellWidth As Double
  Friend lastCellHeight As Double

  '--- Following are window hooks to track scroll/move events ------------------------------
  Private workbookHooks As New Dictionary(Of IntPtr, ExcelWindowHook)

  Public Sub New()
    Dim xl As Excel.Application = CType(ExcelDnaUtil.Application, Excel.Application)
    App = xl
    Instance = Me
  End Sub

  Private Sub App_WorkbookBeforeClose(Wb As Workbook, ByRef Cancel As Boolean) _
        Handles App.WorkbookBeforeClose

    Dim hwnd As IntPtr = New IntPtr(Convert.ToInt64(Wb.Windows(1).Hwnd))
    TaskPaneManager.RemovePaneForWindow(hwnd)
  End Sub

  ' ==========================================================================================
  ' Routine: App_SheetDeactivate
  ' Purpose:
  '   Global handler for sheet deactivation. Stop the timmer and clear active overlays to
  '   prevent misalignment and stale state when
  ' Parameters:
  '   
  ' Returns:
  '   None
  ' Notes:
  ' ==========================================================================================
  Private Sub App_SheetDeactivate(Sh As Object) Handles App.SheetDeactivate
    ClearActiveOverlays()
  End Sub

  ' ==========================================================================================
  ' Routine: App_SheetSelectionChange
  ' Purpose:
  '   Global handler for all sheet selections changes. Delegates to ExcelSelectionChangeHandler to
  '   handle the processing.
  ' Parameters:
  '   
  ' Returns:
  '   None
  ' Notes:
  ' ==========================================================================================
  Private Sub App_SheetSelectionChange(sh As Object, target As Excel.Range) _
    Handles App.SheetSelectionChange

    ' Handle selection change
    ExcelSelectionChangeHandler(App.Selection)
  End Sub

  ' ==========================================================================================
  ' Routine: App_WorkbookOpen and App_WorkbookActivate
  ' Purpose:
  '   Handles WorkbookOpen and WorkbookActivate events to initialize GUID identities in the
  '   cell rang.id for the opened/activated workbook.
  ' Parameters:
  '   wb       - Workbook opened
  ' Returns:
  '
  ' Notes:
  ' ==========================================================================================
  Private Sub App_WorkbookOpen(wb As Excel.Workbook) Handles App.WorkbookOpen
    ExcelCellRuleStore.InitializeGuidIdentityForWorkbook(wb)
    HookWorkbookWindow(wb)
    StartScrollPolling()
  End Sub

  Private Sub App_WorkbookActivate(wb As Excel.Workbook) Handles App.WorkbookActivate
    ExcelCellRuleStore.InitializeGuidIdentityForWorkbook(wb)
    HookWorkbookWindow(wb)
    StartScrollPolling()
  End Sub


#Region "Helpers"
  ' ==========================================================================================
  ' Routine:    HookWorkbookWindow
  ' Purpose:    Hooks the workbook window to monitor for size/move events.
  ' Parameters:
  '   wb       - Workbook to hook
  ' Returns:
  '   None
  ' Notes:
  ' ==========================================================================================
  Private Sub HookWorkbookWindow(wb As Excel.Workbook)
    ' Hook the window to monitor for size/move events
    Dim hwndWorkbook As IntPtr = New IntPtr(Convert.ToInt64(wb.Windows(1).Hwnd))
    Dim xl = CType(ExcelDnaUtil.Application, Excel.Application)

    If Not workbookHooks.ContainsKey(hwndWorkbook) Then
      workbookHooks(hwndWorkbook) = New ExcelWindowHook(hwndWorkbook)
    End If

  End Sub

  ' ==========================================================================================
  ' Routine:    StartScrollPolling
  ' Purpose:    Starts a lightweight timer to detect worksheet scrolling by monitoring
  '             ActiveWindow.ScrollRow and ScrollColumn.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' Notes:
  '   - Excel does NOT expose scroll events through Excel-DNA.
  '   - Polling is the industry-standard solution used by major add-ins.
  ' ==========================================================================================
  Friend Sub StartScrollPolling()
    Try
      If scrollTimer Is Nothing Then
        scrollTimer = New System.Windows.Forms.Timer()
        scrollTimer.Interval = 50 ' ms
      End If

      scrollTimer.Start()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine:    ScrollTimer_Tick
  ' Purpose:    Detects viewport movement by comparing ScrollRow/ScrollColumn values.
  ' Parameters:
  '   sender, e - Timer event parameters.
  ' Returns:
  '   None
  ' Notes:
  '   - Fires when the user scrolls via mouse wheel, scrollbar drag, keyboard, or selection.
  '   - Calls ClearActiveOverlays to remove overlays that are now misaligned.
  ' ==========================================================================================
  Private Sub ScrollTimer_Tick(sender As Object, e As EventArgs) _
    Handles scrollTimer.Tick

    Try
      Dim xl = CType(ExcelDnaUtil.Application, Excel.Application)
      Dim wnd = xl.ActiveWindow
      If wnd Is Nothing Then Exit Sub

      Dim r = wnd.ScrollRow
      Dim c = wnd.ScrollColumn
      ' --- SCROLL DETECTED: keep button, close list, reposition based on lastOverlayCellAddress ---
      If r <> lastRow OrElse c <> lastCol Then

        lastRow = r
        lastCol = c

        If Not String.IsNullOrEmpty(lastOverlayCellAddress) Then
          Try
            Dim cell = xl.Range(lastOverlayCellAddress)

            If cell IsNot Nothing Then
              If CellIsVisible(xl, cell) Then
                ' Cell still visible → reposition button
                ExcelEventHandler.RepositionDropButton(cell)
              Else
                ' Cell off-screen → hide button
                If ExcelEventHandler.activeButtonOverlay IsNot Nothing Then
                  ExcelEventHandler.activeButtonOverlay.Dispose()
                  ExcelEventHandler.activeButtonOverlay = Nothing
                End If
              End If
            End If

          Catch comEx As System.Runtime.InteropServices.COMException
            Exit Sub
          End Try
        End If

        ' Always close list on scroll
        If ExcelEventHandler.activeListOverlay IsNot Nothing Then
          ExcelEventHandler.activeListOverlay.Dispose()
          ExcelEventHandler.activeListOverlay = Nothing
        End If

        Return
      End If

      ' --- NO SCROLL CHANGE: check for cell geometry change (resize) ---
      Try
        Dim cell = xl.Range(lastOverlayCellAddress)
        If cell Is Nothing Then Exit Sub

        If cell.Top <> lastCellTop OrElse
           cell.Left <> lastCellLeft OrElse
           cell.Width <> lastCellWidth OrElse
           cell.Height <> lastCellHeight Then

          lastCellTop = cell.Top
          lastCellLeft = cell.Left
          lastCellWidth = cell.Width
          lastCellHeight = cell.Height
          ExcelEventHandler.RepositionDropButton(cell)

          If ExcelEventHandler.activeListOverlay IsNot Nothing Then
            ExcelEventHandler.activeListOverlay.Dispose()
            ExcelEventHandler.activeListOverlay = Nothing
          End If
        End If
      Catch comEx As System.Runtime.InteropServices.COMException
        ' Excel is in a transient state (resize/edit/scroll) → skip this tick
        Exit Sub

      Catch ex As Exception
        ExcelAsyncUtil.QueueAsMacro(Sub()
                                      ErrorHandler.UnHandleError(ex)
                                    End Sub)
      End Try

    Catch ex As Exception
      ExcelAsyncUtil.QueueAsMacro(Sub()
                                    ErrorHandler.UnHandleError(ex)
                                  End Sub)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine:    CellIsVisible
  ' Purpose:    Checks if a given cell is currently visible in the active window viewport.
  ' Parameters:
  '   xl    - Excel application instance
  '   cell  - Cell to check for visibility
  ' Returns:
  '   True if the cell is visible in the current viewport, False otherwise.
  ' Notes:
  ' ==========================================================================================
  Private Function CellIsVisible(xl As Excel.Application, cell As Excel.Range) As Boolean
    Dim wnd = xl.ActiveWindow
    If wnd Is Nothing Then Return False

    Dim firstRow As Long = wnd.ScrollRow
    Dim firstCol As Long = wnd.ScrollColumn

    Dim vis = wnd.VisibleRange
    Dim rowCount As Long = vis.Rows.Count
    Dim colCount As Long = vis.Columns.Count

    Dim lastRow As Long = firstRow + rowCount - 1
    Dim lastCol As Long = firstCol + colCount - 1

    Return (cell.Row >= firstRow AndAlso cell.Row <= lastRow AndAlso
            cell.Column >= firstCol AndAlso cell.Column <= lastCol)
  End Function

  ' ==========================================================================================
  ' Routine:    StopAllTimers
  ' Purpose:    Stops all active timers in the ExcelEventMonitor.
  ' Parameters:
  '   None
  '   None
  ' Notes:
  ' ==========================================================================================
  Friend Sub StopAllTimers()
    If scrollTimer IsNot Nothing Then
      scrollTimer.Stop()
      scrollTimer.Dispose()
      scrollTimer = Nothing
    End If
  End Sub


#End Region

End Class


