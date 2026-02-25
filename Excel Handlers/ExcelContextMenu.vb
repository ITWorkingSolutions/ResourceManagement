Imports System.Windows.Forms
Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports stdole
Imports Office = Microsoft.Office.Core


Friend Module ExcelContextMenu
  Private Const RM_COPY_ID As String = "RM_Context_Copy"
  Private Const RM_CUT_ID As String = "RM_Context_Cut"
  Private Const RM_PASTE_ID As String = "RM_Context_Paste"

  ' ********************************************************************************************
  ' Routine:    RM_AddContextMenuItems
  ' Purpose:    Inject RM-specific clipboard commands (Copy, Cut, Paste) into Excel’s
  '             right-click Cell context menu. These commands route through the RM identity
  '             pipeline instead of Excel’s native clipboard pipeline.
  '
  ' Contract:
  '   - Must be called once during add-in startup (AutoOpen/OnAddInLoad).
  '   - Adds only TEMPORARY controls (Excel removes them automatically on shutdown).
  '   - Uses unique Tag identifiers so controls can be removed deterministically.
  '   - Does NOT remove or override Excel’s built-in Cut/Copy/Paste commands.
  '
  ' Behaviour:
  '   - Locates the "Cell" CommandBar (Excel’s right-click menu for ranges).
  '   - Removes any existing RM controls to avoid duplicates.
  '   - Adds three RM commands: RM Copy, RM Cut, RM Paste.
  '   - Each command is bound to the corresponding RM handler:
  '         RMCopyHandler, RMCutHandler, RMPasteHandler.
  '
  ' Notes:
  '   - Safe to call multiple times; idempotent due to pre-removal step.
  '   - Does NOT mutate Selection or interfere with Excel’s paste pipeline.
  '   - Only affects the context menu for cell/range selections.
  ' ********************************************************************************************
  Public Sub RM_AddContextMenuItems()

    Dim app As Excel.Application = CType(ExcelDnaUtil.Application, Excel.Application)
    Dim cellMenu As Office.CommandBar = Nothing

    Try
      cellMenu = CType(app.CommandBars("Cell"), Office.CommandBar)
    Catch
      cellMenu = Nothing
    End Try

    If cellMenu Is Nothing Then Exit Sub

    ' Avoid duplicates
    RM_RemoveContextMenuItems()

    Dim bmp As System.Drawing.Image
    Dim pic As IPictureDisp

    ' --- RM Copy ---
    Dim btnCopy As Office.CommandBarButton =
        CType(cellMenu.Controls.Add(Office.MsoControlType.msoControlButton, Temporary:=True),
              Office.CommandBarButton)

    btnCopy.Caption = "RM Copy"
    btnCopy.Tag = RM_COPY_ID
    btnCopy.OnAction = "RMCopyHandler"
    bmp = My.Resources._24copywhite
    pic = PictureDispConverter.ToPictureDisp(bmp)
    btnCopy.Picture = pic

    ' --- RM Cut ---
    Dim btnCut As Office.CommandBarButton =
        CType(cellMenu.Controls.Add(Office.MsoControlType.msoControlButton, Temporary:=True),
              Office.CommandBarButton)

    btnCut.Caption = "RM Cut"
    btnCut.Tag = RM_CUT_ID
    btnCut.OnAction = "RMCutHandler"
    bmp = My.Resources._24cutwhite
    pic = PictureDispConverter.ToPictureDisp(bmp)
    btnCut.Picture = pic

    ' --- RM Paste ---
    Dim btnPaste As Office.CommandBarButton =
        CType(cellMenu.Controls.Add(Office.MsoControlType.msoControlButton, Temporary:=True),
              Office.CommandBarButton)

    btnPaste.Caption = "RM Paste"
    btnPaste.Tag = RM_PASTE_ID
    btnPaste.OnAction = "RMPasteHandler"
    bmp = My.Resources._24pastewhite
    pic = PictureDispConverter.ToPictureDisp(bmp)
    btnPaste.Picture = pic

  End Sub

  ' ********************************************************************************************
  ' Routine:    RM_RemoveContextMenuItems
  ' Purpose:    Cleanly remove all RM-specific clipboard commands from Excel’s Cell context
  '             menu. Ensures no orphaned controls remain after add-in unload.
  '
  ' Contract:
  '   - Must be called during add-in shutdown (AutoClose/OnAddInUnload).
  '   - Removes ONLY controls tagged with RM_COPY_ID, RM_CUT_ID, RM_PASTE_ID.
  '   - Safe to call even if controls were never added or were already removed.
  '
  ' Behaviour:
  '   - Locates the "Cell" CommandBar.
  '   - Iterates through all controls and deletes those matching RM Tag identifiers.
  '
  ' Notes:
  '   - Does NOT affect Excel’s built-in Cut/Copy/Paste commands.
  '   - Does NOT throw if the menu or controls are missing.
  '   - Ensures a clean shutdown and prevents duplicate controls on next load.
  ' ********************************************************************************************
  Public Sub RM_RemoveContextMenuItems()

    Dim app As Excel.Application = CType(ExcelDnaUtil.Application, Excel.Application)
    Dim cellMenu As Office.CommandBar = Nothing

    Try
      cellMenu = CType(app.CommandBars("Cell"), Office.CommandBar)
    Catch
      cellMenu = Nothing
    End Try

    If cellMenu Is Nothing Then Exit Sub

    Dim i As Integer = cellMenu.Controls.Count
    While i >= 1
      Dim ctrl As Office.CommandBarControl = cellMenu.Controls(i)
      Dim tag As String = Nothing

      Try
        tag = ctrl.Tag
      Catch
        tag = Nothing
      End Try

      If tag = RM_COPY_ID OrElse tag = RM_CUT_ID OrElse tag = RM_PASTE_ID Then
        Try
          ctrl.Delete()
        Catch
          ' swallow and continue
        End Try
      End If

      i -= 1
    End While

  End Sub

  ' ========================================================================================
  '  Class: PictureDispConverter
  '  Purpose:
  '       Provides a deterministic, single‑responsibility conversion from a .NET
  '       System.Drawing.Image instance to a COM stdole.IPictureDisp object suitable
  '       for assignment to Office CommandBarButton.Picture.
  '
  '  Rationale:
  '       Excel's CommandBar API requires a COM IPictureDisp (HBITMAP‑backed) object.
  '       .NET Image/Bitmap types are GDI+ objects and cannot be assigned directly.
  '       AxHost.GetIPictureDispFromPicture provides the only stable, framework‑native
  '       conversion path without invoking OleCreatePictureIndirect or manual P/Invoke.
  '
  '  Behaviour:
  '       - Accepts any System.Drawing.Image.
  '       - Returns a COM stdole.IPictureDisp wrapper.
  '       - Does not modify or dispose the source image.
  '       - Does not perform masking or transparency manipulation.
  '
  '  Limitations:
  '       - Output is a raw HBITMAP; alpha transparency is not preserved.
  '       - Caller is responsible for supplying a valid monochrome mask if required.
  '
  '  Notes:
  '       - Constructor is private because AxHost requires a host class instance,
  '         but no external instantiation is permitted or meaningful.
  '       - This class is intentionally minimal and side‑effect free.
  '
  '  ========================================================================================

  Friend Class PictureDispConverter
    Inherits AxHost

    Private Sub New()
      MyBase.New(String.Empty)
    End Sub

    Public Shared Function ToPictureDisp(img As System.Drawing.Image) As IPictureDisp
      Return CType(GetIPictureDispFromPicture(img), IPictureDisp)
    End Function
  End Class

End Module
