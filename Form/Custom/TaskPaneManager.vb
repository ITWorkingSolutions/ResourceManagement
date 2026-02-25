Imports System.Windows.Forms
Imports ExcelDna.Integration
Imports ExcelDna.Integration.CustomUI
Imports Microsoft.Office.Interop.Excel

Friend Module TaskPaneManager

  '==========================================================================================
  ' MODULE: TaskPaneManager
  '
  ' PURPOSE:
  '   Manages one CustomTaskPane instance per Excel window. Ensures panes are created,
  '   shown, hidden, and removed deterministically based on window lifecycle events.
  '
  ' DESIGN:
  '   - Keyed by Window.Hwnd (IntPtr) to guarantee one pane per workbook window.
  '   - No workbook or model logic here; this module only manages pane instances.
  '   - Pane content (ExcelRuleDesigner) handles its own model reload/reset.
  '
  ' INVARIANTS:
  '   - _panes contains at most one pane per Excel window.
  '   - Keys must always be IntPtr constructed from Window.Hwnd.
  '   - Pane visibility is controlled explicitly; Excel handles window switching.
  '
  '==========================================================================================

  Private _panes As New Dictionary(Of IntPtr, CustomTaskPane)

  '------------------------------------------------------------------------------------------
  ' ROUTINE: ShowPaneForActiveWindow
  '
  ' PURPOSE:
  '   Ensures the active Excel window has a task pane instance. Creates it if missing,
  '   then makes it visible. Called from the Ribbon button.
  '
  ' PRECONDITIONS:
  '   - ExcelDnaUtil.Application.ActiveWindow is not Nothing.
  '
  ' POSTCONDITIONS:
  '   - A CustomTaskPane exists for the active window.
  '   - The pane is visible.
  '
  '------------------------------------------------------------------------------------------
  Public Sub ShowPaneForActiveWindow()
    Dim app = ExcelDnaUtil.Application
    Dim wn As Window = app.ActiveWindow
    If wn Is Nothing Then Exit Sub
    Dim hwnd As IntPtr = New IntPtr(Convert.ToInt64(wn.Hwnd))

    If Not _panes.ContainsKey(hwnd) Then

      Dim pane = CustomTaskPaneFactory.CreateCustomTaskPane(
                    GetType(ExcelRuleDesigner),
                    "Rule Designer",
                    app.ActiveWindow
                 )

      AddHandler Pane.VisibleStateChange, AddressOf Pane_VisibleStateChange
      Pane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
      Pane.Width = 300

      _panes(hwnd) = Pane

    End If

    _panes(hwnd).Visible = True

  End Sub

  '------------------------------------------------------------------------------------------
  ' ROUTINE: ActivatePaneForActiveWindow
  '
  ' PURPOSE:
  '   Makes the pane for the active window visible *if it exists*. Does not create a pane.
  '   Used only when Excel switches windows and the user previously opened a pane.
  '
  ' POSTCONDITIONS:
  '   - If a pane exists for the active window, it becomes visible.
  '   - If no pane exists, nothing happens.
  '
  '------------------------------------------------------------------------------------------
  Public Sub ActivatePaneForActiveWindow()
    Dim app = ExcelDnaUtil.Application
    Dim hwnd As IntPtr = New IntPtr(Convert.ToInt64(app.Hwnd))

    If _panes.ContainsKey(hwnd) Then
      _panes(hwnd).Visible = True
    End If
  End Sub

  '------------------------------------------------------------------------------------------
  ' ROUTINE: HidePaneForWindow
  '
  ' PURPOSE:
  '   Hides (but does not delete) the pane associated with the specified window.
  '   Called when a window is deactivated.
  '
  ' POSTCONDITIONS:
  '   - Pane remains in dictionary but is not visible.
  '
  '------------------------------------------------------------------------------------------
  Public Sub HidePaneForWindow(hwnd As IntPtr)
    If _panes.ContainsKey(hwnd) Then
      _panes(hwnd).Visible = False
    End If
  End Sub

  '------------------------------------------------------------------------------------------
  ' ROUTINE: RemovePaneForWindow
  '
  ' PURPOSE:
  '   Deletes and removes the pane associated with the specified window.
  '   Called when a workbook window is closing.
  '
  ' POSTCONDITIONS:
  '   - Pane is deleted and removed from dictionary.
  '   - No dangling references remain.
  '
  '------------------------------------------------------------------------------------------
  Public Sub RemovePaneForWindow(hwnd As IntPtr)
    If _panes.ContainsKey(hwnd) Then
      _panes(hwnd).Delete()
      _panes.Remove(hwnd)
    End If
  End Sub

  '------------------------------------------------------------------------------------------
  ' ROUTINE: Pane_VisibleStateChange
  '
  ' PURPOSE:
  '   Handles lifecycle events for each pane. When shown, reloads the model to ensure
  '   cross-window consistency. When hidden, clears transient UI state.
  '
  ' POSTCONDITIONS:
  '   - Visible: ExcelRuleDesigner.ReloadModel() is executed.
  '   - Hidden: ExcelRuleDesigner.ResetPane() is executed.
  '
  '------------------------------------------------------------------------------------------
  Private Sub Pane_VisibleStateChange(pane As CustomTaskPane)
    Dim ctrl = DirectCast(pane.ContentControl, ExcelRuleDesigner)

    If pane.Visible Then
      ' Excel needs one UI cycle to finish showing the pane
      System.Windows.Forms.Application.DoEvents()

      Dim root = TryCast(pane.ContentControl, System.Windows.Forms.Control)
      If root IsNot Nothing Then
        root.Select()
        root.Focus()
      End If

      ctrl.ReloadModel()   ' Another window may have changed DB state
    Else
      ctrl.ResetPane()
    End If
  End Sub

End Module