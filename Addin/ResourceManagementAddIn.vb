Imports System.Windows.Forms
Imports ExcelDna.Integration
Imports Microsoft.Office.Interop

' ============================================================================================
'  Class: ResourceManagementAddIn
'  Purpose:
'       Excel-DNA add-in entry point. Loads configuration and initializes application context.
'
'  Notes:
'       - Called automatically by Excel-DNA at add-in load/unload.
'       - Ensures DatabaseConfig is loaded once and available globally via AppContext.
' ============================================================================================
Public Class ResourceManagementAddIn
  Implements IExcelAddIn
  Private monitor As ExcelEventMonitor

  ' ========================================================================================
  '  Routine: AutoOpen
  '  Purpose:
  '       Executed when the Excel-DNA add-in loads. Initializes configuration and context.
  '
  '  Returns:
  '       None
  ' ========================================================================================
  Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen

    ' === Variable declarations ===
    Dim config As DatabaseConfig = Nothing
    Dim hwnd = Nothing

    Try
      ' --- Set Excel main window as owner for UI dialogs ---
      hwnd = ExcelDna.Integration.ExcelDnaUtil.WindowHandle
      UIMessageOwner.ExcelOwner = New WindowWrapper(hwnd)

      ' --- Initialize event monitor ---
      monitor = New ExcelEventMonitor()

      Dim app As Excel.Application = CType(ExcelDnaUtil.Application, Excel.Application)

      ' --- Add custom copy handler ---
      app.OnKey("^c", "RMCopyHandler")            ' Ctrl+C copy
      app.OnKey("+{INSERT}", "RMCopyHandler")     ' Shift+Insert copy
      ' --- Add custom cut handler ---
      app.OnKey("^x", "RMCutHandler")             ' Ctrl+X cut
      app.OnKey("+{DELETE}", "RMCutHandler")      ' Shift+Delete cut
      ' --- Add custom paste handler ---
      app.OnKey("^v", "RMPasteHandler")           ' Ctrl+V paste
      app.OnKey("+{INSERT}", "RMPasteHandler")    ' Shift+Insert paste
      app.OnKey("^+{V}", "RMPasteHandler")        ' Ctrl+Shift+V paste

      ' --- Add custom context menu items ---
      RM_AddContextMenuItems()

      ' --- Load configuration from AppData ---
      config = DatabaseConfigManager.Load()

      ' --- Store in global application context ---
      AddInContext.Current = New AddInContext With {
            .Config = config
        }

      ' Initialize all currently open workbooks
      For Each wb As Excel.Workbook In app.Workbooks
        ExcelCellRuleStore.InitializeGuidIdentityForWorkbook(wb)
      Next

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      config = Nothing
    End Try

  End Sub

  ' ========================================================================================
  '  Routine: AutoClose
  '  Purpose:
  '       Executed when the Excel-DNA add-in unloads. Performs cleanup if required.
  '
  '  Returns:
  '       None
  ' ========================================================================================
  Public Sub AutoClose() Implements IExcelAddIn.AutoClose

    Try
      ' --- Dispose event monitor ---
      monitor = Nothing

      Dim app As Excel.Application = CType(ExcelDnaUtil.Application, Excel.Application)

      ' Restore COPY
      app.OnKey("^c", "")
      app.OnKey("+{INSERT}", "")
      ' Restore CUT
      app.OnKey("^x", "")
      app.OnKey("+{DELETE}", "")
      ' Restore PASTE
      app.OnKey("^v", "")
      app.OnKey("+{INSERT}", "")
      app.OnKey("^+v", "")

      ' Remove custom context menu items
      RM_RemoveContextMenuItems()

      ExcelEventMonitor.Instance.StopAllTimers()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' No cleanup required
    End Try

  End Sub

End Class
