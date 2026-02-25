Option Explicit On
Imports System.Drawing
Imports System.IO
Imports System.Reflection
Imports System.Resources
Imports System.Runtime.InteropServices
Imports ExcelDna.Integration
Imports ExcelDna.Integration.CustomUI
Imports SQLitePCL

' Class needs to be public and ComVisible for Excel-DNA to find it
<ComVisible(True)>
Public Class Ribbon
  Inherits ExcelRibbon
  Friend Shared Instance As Ribbon
  Friend Shared RibbonUI As IRibbonUI

  Public Overrides Function GetCustomUI(ribbonID As String) As String
    Return LoadEmbeddedText("Ribbon.xml")
  End Function

  Private Function LoadEmbeddedText(resourceSuffix As String) As String
    Dim asm = Assembly.GetExecutingAssembly()
    Dim name = asm.GetManifestResourceNames().
        FirstOrDefault(Function(n) n.EndsWith(resourceSuffix, StringComparison.OrdinalIgnoreCase))

    If name Is Nothing Then
      Throw New Exception("Embedded resource not found: " & resourceSuffix)
    End If

    Using s = asm.GetManifestResourceStream(name)
      Using r As New StreamReader(s)
        Return r.ReadToEnd()
      End Using
    End Using
  End Function

  Public Function OnGetImage(Control As IRibbonControl) As Object
    Dim asmName As String = Assembly.GetExecutingAssembly().GetName().Name
    Dim pngPath As String = asmName & ".64"
    Select Case Control.Id
      Case "btnRulePane"
        Return LoadPng(pngPath & "report.png")
      Case "btnResourceManager"
        Return LoadPng(pngPath & "resourcemanager.png")
      Case "btnRefreshAll"
        Return LoadPng(pngPath & "refreshall.png")
      Case "btnRMCut"
        Return LoadPng(pngPath & "cut.png")
      Case "btnRMCopy"
        Return LoadPng(pngPath & "copy.png")
      Case "btnRMPaste"
        Return LoadPng(pngPath & "paste.png")
      Case "btnDatabaseManagement"
        Return LoadPng(pngPath & "databasemanager.png")
      Case "btnComapnyLogo"
        Return LoadPng(pngPath & "logo.png")
      Case "btnListItemTypeAndItem"
        Return LoadPng(pngPath & "listitem.png")
      Case "btnResourceListItem"
        Return LoadPng(pngPath & "tag.png")
      Case "btnClosure"
        Return LoadPng(pngPath & "closure.png")
      Case "btnAbout"
        Return LoadPng(pngPath & "question.png")
    End Select
    Return Nothing
  End Function

  ' ==========================================================================================
  ' Function: LoadPng
  ' Purpose:  Loads an embedded PNG (or any image) from the assembly's manifest resources.
  ' Notes:
  '   - The PNG must be added to the project and set to:
  '         Build Action = Embedded Resource
  '   - resourceName must match the full manifest name, e.g.:
  '         "YourAssemblyName.Images.report.png"
  '   - Returns a Bitmap which Excel-DNA accepts directly for Ribbon images.
  ' ==========================================================================================

  Private Function LoadPng(resourceName As String) As Bitmap
    Dim asm As Assembly = GetType(Ribbon).Assembly

    Using stream As Stream = asm.GetManifestResourceStream(resourceName)
      If stream Is Nothing Then
        Throw New Exception("Resource not found: " & resourceName)
      End If

      Return New Bitmap(stream)
    End Using
  End Function

  ' ==========================================================================================
  ' Routine: Ribbon_OnLoad
  ' Purpose: Save the handle to the ribbon when loaded to use in other routines.
  ' ==========================================================================================
  Public Sub Ribbon_OnLoad(ribbon As IRibbonUI)
    Const ROUTINE As String = "Ribbon_OnLoad"
    On Error GoTo Error_Handler
    Batteries_V2.Init()   ' <-- REQUIRED for Microsoft.Data.Sqlite in Excel-DNA
    Instance = Me
    RibbonUI = ribbon
    ' === Refresh the ribbon as it loads the report list first ===
    ResourceManagement_RefreshRibbon()

Exit_Routine:
    Exit Sub
Error_Handler:
    MsgBox("Error #" & Err.Number & " in " & ROUTINE & vbCrLf & Err.Description, vbCritical, "Unhandled Error")
    Resume Exit_Routine
  End Sub

  ' ==========================================================================================
  ' Routine: ResourceManagement_RefreshRibbon
  ' Purpose: Refresh the ribbon by invalidate. Causes ribbon to reload and therefore set button
  '          enabled state.
  ' ==========================================================================================
  '
  Public Sub ResourceManagement_RefreshRibbon()
    On Error Resume Next

    If Ribbon.RibbonUI IsNot Nothing Then
      Ribbon.RibbonUI.Invalidate()

    End If
  End Sub

  ' ==========================================================================================
  ' Routine: ResourceManagement_OnAction
  ' Purpose: This callback routine from the custom ribbon calls the routine based on custom ribbon
  '          control.id calling it.
  ' ==========================================================================================
  Public Sub ResourceManagement_OnAction(control As IRibbonControl)
    Dim app = ExcelDna.Integration.ExcelDnaUtil.Application
    Select Case control.Id
      Case "btnResourceManager"
        Dim oForm As New ResourceManager
        oForm = New ResourceManager()
        oForm.ShowDialog()
        oForm.Close()
        Return
      Case "btnRefreshAll"
        ' Allows the user to refresh all data connections and recalculate the workbook
        If app Is Nothing Then Exit Sub
        If app.ActiveWorkbook Is Nothing Then Exit Sub
        app.CalculateFull()
        app.ActiveWorkbook.RefreshAll()
      Case "btnRulePane"
        TaskPaneManager.ShowPaneForActiveWindow()
        'TaskPaneManager.ShowPane()
      Case "btnRMCut"
        RMCutHandler()
      Case "btnRMCopy"
        RMCopyHandler()
      Case "btnRMPaste"
        RMPasteHandler()
      Case "btnDatabaseManagement"
        Dim oForm As New DatabaseManager
        oForm = New DatabaseManager()
        oForm.ShowDialog()
        oForm.Close()
        Return
      Case "btnComapnyLogo"
        SaveCompanyLogo()
      Case "btnListItemTypeAndItem"
        Dim oForm As New ListItemTypeAndItem
        oForm = New ListItemTypeAndItem()
        oForm.ShowDialog()
        oForm.Close()
        Return
      Case "btnResourceListItem"
        Dim oForm As New ResourceListItem
        oForm = New ResourceListItem()
        oForm.ShowDialog()
        oForm.Close()
        Return
      Case "btnClosure"
        Dim oForm As New Closure
        oForm = New Closure()
        oForm.ShowDialog()
        oForm.Close()
        Return
      Case "btnAbout"
        AboutDisplay()
    End Select
  End Sub

  ' ==========================================================================================
  ' Routine: ResourceManagement_getEnabled
  ' Purpose: Enables or Disable custom ribbon buttons. Currenetly all buttons are enabled.
  ' ==========================================================================================
  Public Function ResourceManagement_getEnabled(control As IRibbonControl) As Boolean
    Return True
  End Function

  ' ==========================================================================================
  ' Routine:     ActivateResourceManagementTab
  ' Purpose:     Reasserts focus on the Resource Management Ribbon tab after context shift.
  '
  ' Notes:
  '   - Used to restore tab focus after workbook activation or modal form closure.
  '   - Ribbon tab ID must match XML definition: "ResourceManagement_Tab"
  ' ==========================================================================================
  Public Sub ActivateResourceManagementTab()
    Const ROUTINE As String = "ActivateResourceManagementTab"
    On Error GoTo Error_Handler

    ' === Defensive guard: ensure g_ribbonUI is initialized ===
    If Ribbon.RibbonUI IsNot Nothing Then
      Ribbon.RibbonUI.ActivateTab("ResourceManagement_Tab")
    End If

Exit_Routine:
    Exit Sub
Error_Handler:
    MsgBox("Error #" & Err.Number & " in " & ROUTINE & vbCrLf & Err.Description, vbCritical, "Unhandled Error")
    Resume Exit_Routine
  End Sub

End Class
