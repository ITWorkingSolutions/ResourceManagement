Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports ResourceManagement.ExcelEventMonitor

Public Module ExcelEventHandler
  Friend lastOverlayCellAddress As String
  ' Following used for copy / cut / paste handling
  Friend Enum RMPasteOperationType
    None
    Copy
    Cut
  End Enum

  Friend LastPasteOperation As RMPasteOperationType = RMPasteOperationType.None
  Private LastPasteSource As Excel.Range ' used to store the last paste source range
  Private buttonClickArmed As Boolean

  Public Sub ExcelSelectionChangeHandler(target As Excel.Range)
    Try
      If target Is Nothing Then Exit Sub

      Dim xl = CType(ExcelDnaUtil.Application, Excel.Application)
      If xl.CutCopyMode <> 0 Then
        ' Do NOT clean overlays
        ' Do NOT create overlays
        lastOverlayCellAddress = Nothing
        Exit Sub
      End If

      If lastOverlayCellAddress IsNot Nothing AndAlso
       String.Equals(target.Address, lastOverlayCellAddress, StringComparison.Ordinal) Then
        ' Same cell → do NOT destroy overlays
        Exit Sub
      End If

      ' Otherwise, cleanup
      If activeButtonOverlay IsNot Nothing Then
        activeButtonOverlay.Dispose()
        activeButtonOverlay = Nothing
      End If

      If activeListOverlay IsNot Nothing Then
        activeListOverlay.Dispose()
        activeListOverlay = Nothing
      End If

      ' Store address
      lastOverlayCellAddress = target.Address

      ' 1. Only act on single-cell selections
      If target.CountLarge <> 1 Then Exit Sub

      Dim ruleId As String = Nothing
      Dim ruleName As String = Nothing
      Dim listSelectType As String = Nothing
      Dim parameters As IList(Of ExcelCellRuleStore.RuleParameter) = Nothing
      Dim regionState As String = Nothing

      ' 2. Ask backend if this cell has a rule
      If Not ExcelCellRuleStore.TryGetRuleForCell(target, ruleId, listSelectType, parameters, regionState) Then
        Exit Sub
      End If

      ' 3. If region is broken, do not show dropdown
      If regionState = "NeedsRepair" Then
        Exit Sub
      End If

      ' 4. Resolve parameter values for this cell
      Dim resolvedArgs As Object() = ResolveParameterValues(target, parameters)

      ' 5. Get values from rule engine
      Dim result As Object = ExcelRuleEngine.GetRuleValues(ruleId, ruleName, resolvedArgs)

      Dim options As List(Of String) = Nothing

      ' We only support list-returning rules here
      If TypeOf result Is Object() Then
        ' Safe to cast now
        Dim arr = DirectCast(result, Object())
        options = arr.
        Select(Function(o) If(o Is Nothing, String.Empty, o.ToString())).
        ToList()
      ElseIf TypeOf result Is String Then
        ' Treat a single scalar as a single-option list (optional, but user-friendly)
        options = New List(Of String) From {CStr(result)}
      Else
        ' Any other shape is a user error: wrong rule type for dropdown
        MessageBox.Show(
        $"The rule '{ruleName}' does not return a list of values and cannot be used for a dropdown.",
        "Invalid Rule",
        MessageBoxButtons.OK,
        MessageBoxIcon.Warning
      )
        Exit Sub
      End If

      ' If we somehow ended up with no options, don't show the button
      If options Is Nothing OrElse options.Count = 0 Then
        Exit Sub
      End If


      ' 6. Build ExcelReference for the target cell
      Dim sheetName As String = CType(target.Worksheet, Excel.Worksheet).Name
      Dim callerRef As New ExcelReference(target.Row - 1, target.Row - 1,
                                        target.Column - 1, target.Column - 1,
                                        sheetName)

      ' 7. Show dropdown button over cell

      ShowDropButtonOverCell(target, options, callerRef, listSelectType)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try

  End Sub

  ' ==========================================================================================
  ' Routine:    RMCopyHandler
  ' Purpose:    Intercept Excel copy operations to capture the source range for later use in
  '             paste reconstruction and identity propagation.
  '
  ' Contract:
  '   - Must be bound to Ctrl+C (or equivalent) via OnKey.
  '   - Stores the user’s current Selection into LastPasteSource.
  '   - Must call Excel’s native Copy command to preserve clipboard behaviour.
  '   - Sets LastPasteOperation to Copy.
  '
  ' Behaviour:
  '   - Captures the exact source geometry used for paste.
  '   - Does NOT modify Selection.
  '   - Does NOT create GUIDs or modify identities.
  '
  ' Notes:
  '   - LastPasteSource is consumed by RMPasteHandler.
  '   - Safe to call repeatedly; always overwrites previous source.
  ' ==========================================================================================

  <ExcelCommand>
  Public Sub RMCopyHandler()
    Dim app As Excel.Application = CType(ExcelDnaUtil.Application, Excel.Application)
    ' Store the source range
    LastPasteSource = app.Selection
    ' Set the last operation
    LastPasteOperation = RMPasteOperationType.Copy
    ' Let Excel do the real copy
    app.CommandBars.ExecuteMso("Copy")
  End Sub

  ' ==========================================================================================
  ' Routine:    RMCutHandler
  ' Purpose:    Intercept Excel cut operations to capture the source range for later use in
  '             paste reconstruction and identity propagation.
  '
  ' Contract:
  '   - Must be bound to Ctrl+C (or equivalent) via OnKey.
  '   - Stores the user’s current Selection into LastPasteSource.
  '   - Must call Excel’s native Cut command to preserve clipboard behaviour.
  '   - Sets LastPasteOperation to Cut.
  '
  ' Behaviour:
  '   - Captures the exact source geometry used for paste.
  '   - Does NOT modify Selection.
  '   - Does NOT create GUIDs or modify identities.
  '
  ' Notes:
  '   - LastPasteSource is consumed by RMPasteHandler.
  '   - Safe to call repeatedly; always overwrites previous source.
  ' ==========================================================================================
  <ExcelCommand>
  Public Sub RMCutHandler()
    Dim app As Excel.Application = CType(ExcelDnaUtil.Application, Excel.Application)
    ' Store the source range
    LastPasteSource = app.Selection
    ' Set the last operation
    LastPasteOperation = RMPasteOperationType.Cut
    app.CommandBars.ExecuteMso("Cut")
  End Sub

  ' ==========================================================================================
  ' Routine:    RMPasteHandler
  ' Purpose:    Intercept Excel paste operations to ensure identity propagation is deterministic,
  '             audit-safe, and consistent with the resource model. Excel collapses Selection
  '             during paste, so this routine reconstructs the intended target range and applies
  '             identity logic over the correct geometry.
  '
  ' Contract:
  '   - LastPasteSource must contain the range that was copied (set by RMCopyHandler/RMCutHandler).
  '   - The user’s current Selection is treated as the paste anchor.
  '   - The intended target range is reconstructed via SetTargetRange(src).
  '   - BEFORE and AFTER snapshots are taken over the reconstructed target range, not Selection.
  '   - If IsPaste(beforeSnap, afterSnap) is true, identity propagation is applied.
  '
  ' Behaviour:
  '   - Performs a normal Excel paste via ExecuteMso("Paste").
  '   - Reconstructs the target range after paste (Selection may have collapsed).
  '   - Optionally fills missing IDs in tiled pastes (vertical tiling gaps).
  '   - Delegates full identity logic to PasteIdentityHandler.
  '
  ' Notes:
  '   - Does NOT modify Selection at any point (Selection is “hot” during paste).
  '   - Does NOT resize or mutate live Excel objects beyond the intended target range.
  '   - Safe to call repeatedly; no state leakage beyond LastPasteSource and LastPasteOperation.
  ' ==========================================================================================

  <ExcelCommand>
  Public Sub RMPasteHandler()
    Try
      Dim app As Excel.Application = CType(ExcelDnaUtil.Application, Excel.Application)
      Dim ws As Excel.Worksheet = app.ActiveSheet
      Dim src As Excel.Range = LastPasteSource
      Dim tgt As Excel.Range = Nothing

      ' Exist if no src range set
      If src Is Nothing Then Exit Sub

      ' Make sure the target is at least the source range size
      tgt = SetTargetRange(src)

      ' BEFORE snapshot
      Dim beforeSnap As RangeSnapshot = SnapshotRange(tgt)
      ' Let Excel do a normal paste
      app.CommandBars.ExecuteMso("Paste")

      ' Make sure the target is at least the source range size
      tgt = SetTargetRange(src)

      ' AFTER snapshot
      Dim afterSnap As RangeSnapshot = SnapshotRange(tgt)

      If IsPaste(beforeSnap, afterSnap) Then
        ' Fix Excel's missing ID propagation in tiles
        If LastPasteOperation = RMPasteOperationType.Copy Then
          ' COPY semantics
          If LastPasteSource IsNot Nothing AndAlso app.CutCopyMode = Excel.XlCutCopyMode.xlCopy Then
            PropagateMissingIdsFromSourceTile(LastPasteSource, tgt)
          End If
        ElseIf LastPasteOperation = RMPasteOperationType.Cut Then
          ' CUT semantics
          ' No tiling fix needed — identity moves, not duplicates
        End If
        PasteIdentityHandler(ws, tgt)
        LastPasteOperation = RMPasteOperationType.None ' Reset operation
      End If
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try

  End Sub

  ' ==========================================================================================
  ' Routine:    SetTargetRange
  ' Purpose:    Reconstruct the intended paste target range without mutating Selection.
  '             Excel collapses Selection during paste, so this routine creates a detached
  '             Range object representing the true target geometry.
  '
  ' Contract:
  '   - src: The source range captured at copy time.
  '   - The user’s Selection is treated as the paste anchor.
  '   - Returns a detached Range object (never the live Selection).
  '
  ' Behaviour:
  '   - Creates a new Range object using anchor.Address.
  '   - If the anchor is a single cell and the source is larger, expands the target to match
  '     the source shape (spill/tiling scenario).
  '   - If the anchor is already multi-cell, returns it unchanged.
  '
  ' Notes:
  '   - NEVER resizes Selection directly (would corrupt Excel’s paste pipeline).
  '   - ALWAYS returns a safe, detached Range object.
  '   - Does NOT validate shape compatibility; that is handled by downstream logic.
  ' ==========================================================================================

  Private Function SetTargetRange(src As Excel.Range) As Excel.Range
    Try
      ' Capture anchor so we can detach range from actual selection
      Dim anchor As Excel.Range = ExcelDnaUtil.Application.Selection
      ' Create a detached copy of the target range
      Dim tgt As Excel.Range = anchor.Worksheet.Range(anchor.Address)
      ' If the target selection is a single cell we need to check the source isn't larger 
      ' and we need to resize target for a spill
      If tgt.Count = 1 Then
        If src.Rows.Count > 1 Or src.Columns.Count > 1 Then
          tgt = tgt.Resize(src.Rows.Count, src.Columns.Count)
        End If
      End If
      Return tgt
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return Nothing
    End Try

  End Function
#Region "Dropdown Overlays"
  Friend activeButtonOverlay As DropButtonOverlay
  Friend activeListOverlay As DropListOverlay

  ' ==========================================================================================
  ' Routine:    ShowDropButtonOverCell
  ' Purpose:
  '   Displays a fixed‑size drop‑button overlay aligned to the bottom‑right edge of the target
  '   Excel cell. This matches native Excel dropdown button behaviour.
  '
  ' Parameters:
  '   target     - Excel.Range representing the cell to anchor the drop button to.
  '   options    - List(Of String) containing the dropdown options to display when clicked.
  '   callerRef  - ExcelReference identifying the calling cell for downstream processing.
  '
  ' Returns:
  '   None
  '
  ' Notes:
  '   - Button is positioned OUTSIDE the cell, not inside it.
  '   - Button remains visible while the cell is active.
  '   - Button click toggles the dropdown list overlay.
  ' ==========================================================================================
  Friend Sub ShowDropButtonOverCell(target As Excel.Range,
                                  options As List(Of String),
                                  callerRef As ExcelReference, listSelectType As String)

    Try
      If ExcelIsEditing() Then Return

      Dim xl = CType(ExcelDnaUtil.Application, Excel.Application)

      If xl.CutCopyMode <> 0 Then Return

      ' Cleanup existing overlay
      If activeButtonOverlay IsNot Nothing Then
        activeButtonOverlay.Dispose()
        activeButtonOverlay = Nothing
      End If

      ' Excel gives bitmap-scaled pixel coordinates (correct for positioning)
      Dim bottomPx As Integer =
            xl.ActiveWindow.ActivePane.PointsToScreenPixelsY(target.Top + target.Height)

      Dim rightPx As Integer =
            xl.ActiveWindow.ActivePane.PointsToScreenPixelsX(target.Left + target.Width)

      Dim dpi As Integer = GetDpiForWindow(ExcelDnaUtil.WindowHandle)
      Dim scale As Double = dpi / 96.0

      Dim btnH As Integer = CInt(20 * scale)
      Dim btnW As Integer = btnH   ' square

      ' Correct alignment
      Dim initialX As Integer = rightPx
      Dim initialY As Integer = bottomPx - btnH

      ExcelAsyncUtil.QueueAsMacro(
            Sub()
              ' --- Timer delay to ensure this runs after Excel's own selection handling and potential repaint ---
              Dim t As New Timer()
              t.Interval = 1
              AddHandler t.Tick,
                  Sub()
                    t.Stop()
                    t.Dispose()
                    ' Make sure no other overlay has been created in the meantime (e.g., by another SelectionChange event)
                    If activeButtonOverlay IsNot Nothing Then
                      activeButtonOverlay.Dispose()
                      activeButtonOverlay = Nothing
                    End If

                    activeButtonOverlay =
                      New DropButtonOverlay(initialX, initialY, btnW, btnH,
                      Sub()
                        If activeListOverlay IsNot Nothing Then
                          activeListOverlay.Dispose()
                          activeListOverlay = Nothing
                          Return
                        End If

                        ShowListDropdownOverlay(target, options, callerRef, listSelectType)
                      End Sub)
                  End Sub
              t.Start()

            End Sub)

      ' Update geometry tracking
      Dim mon = ExcelEventMonitor.Instance
      mon.lastCellTop = target.Top
      mon.lastCellLeft = target.Left
      mon.lastCellWidth = target.Width
      mon.lastCellHeight = target.Height

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try

  End Sub

  ' ==========================================================================================
  ' Routine: ShowListDropdownOverCell
  ' Purpose:
  '   Displays a fixed‑size dropdown list form directly below the target Excel cell. This uses
  '   a native excel form rather than winform so it works native form on the worksheet.
  '
  ' Parameters:
  '   target     - Excel.Range representing the cell whose dropdown list is being shown.
  '   options    - List(Of String) containing the selectable values to display in the list.
  '   callerRef  - ExcelReference identifying the cell to receive the selected value.
  '
  ' Returns:
  '   None. Displays a modeless dropdown form and registers cleanup handlers.
  '
  ' Notes:
  '   - The dropdown width is based on Excel's cell width.
  '   - Any existing dropdown form is closed before creating a new one.
  ' ==========================================================================================
  Private Sub ShowListDropdownOverlay(target As Excel.Range, options As List(Of String),
                                      callerRef As ExcelReference, listSelectType As String)

    Try
      If activeListOverlay IsNot Nothing Then
        activeListOverlay.Dispose()
        activeListOverlay = Nothing
      End If

      Dim xl = CType(ExcelDnaUtil.Application, Excel.Application)

      ' Excel pixel coordinates (already bitmap-scaled)
      Dim leftPx As Integer = xl.ActiveWindow.ActivePane.PointsToScreenPixelsX(target.Left)
      Dim rightPx As Integer = xl.ActiveWindow.ActivePane.PointsToScreenPixelsX(target.Left + target.Width)
      Dim bottomPx As Integer = xl.ActiveWindow.ActivePane.PointsToScreenPixelsY(target.Top + target.Height)

      ' Raw width from Excel (bitmap-scaled)
      Dim widthPx As Integer = rightPx - leftPx

      ' --- DPI scaling for overlay window size ---
      Dim hwndExcel As IntPtr = New IntPtr(xl.Hwnd)
      Dim scaledWidth = widthPx

      ' --- For multi-select lists, we need to parse the existing cell value into a list of pre-selected values ---
      Dim preSelectedValues As List(Of String)

      If listSelectType = ExcelListSelectType.MultiSelect.ToString() Then
        Dim raw As Object = target.Value

        If raw Is Nothing Then
          preSelectedValues = New List(Of String)()
        Else
          Dim text As String = CStr(raw)

          If String.IsNullOrWhiteSpace(text) Then
            preSelectedValues = New List(Of String)()
          Else
            preSelectedValues =
                text.Split(","c).
                     Select(Function(s) s.Trim()).
                     Where(Function(s) s.Length > 0).
                     ToList()
          End If
        End If
      Else
        preSelectedValues = New List(Of String)()
      End If

      activeListOverlay =
            New DropListOverlay(leftPx,
                                bottomPx,
                                scaledWidth,
                                options,
                                listSelectType,
                                preSelectedValues,
                Sub(selected As String)
                  ExcelAsyncUtil.QueueAsMacro(
                        Sub()
                          callerRef.SetValue(selected)
                        End Sub)
                End Sub)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try

  End Sub

  ' ==========================================================================================
  ' Routine:    ClearActiveOverlays
  ' Purpose:    Disposes and clears any active list or button overlays when the worksheet
  '             viewport changes (scroll, selection-driven scroll, etc.).
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' Notes:
  '   - Called from Excel scroll-related events.
  '   - Ensures overlays never drift out of alignment with the grid.
  ' ==========================================================================================
  Public Sub ClearActiveOverlays()
    Try
      ' --- Normal execution ---
      If ExcelEventHandler.activeListOverlay IsNot Nothing Then
        ExcelEventHandler.activeListOverlay.Dispose()
        ExcelEventHandler.activeListOverlay = Nothing
      End If

      If ExcelEventHandler.activeButtonOverlay IsNot Nothing Then
        ExcelEventHandler.activeButtonOverlay.Dispose()
        ExcelEventHandler.activeButtonOverlay = Nothing
      End If

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' --- Cleanup ---
      ' Nothing additional required
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine:    RepositionDropButton
  ' Purpose:
  '   Repositions the active drop‑button overlay so that it remains anchored to the
  '   bottom‑right corner of the target Excel cell. This matches native Excel behaviour
  '   during row/column resizing.
  '
  ' Parameters:
  '   cell  - Excel.Range representing the cell the button should anchor to.
  '
  ' Returns:
  '   None
  '
  ' Notes:
  '   - Called from ExcelEventMonitor when cell geometry changes (row height or column width).
  '   - Button is positioned using screen pixel coordinates derived from Excel's pane.
  '   - The dropdown list overlay is cleared separately when resizing occurs.
  ' ==========================================================================================
  Friend Sub RepositionDropButton(cell As Excel.Range)
    Try
      If activeButtonOverlay Is Nothing Then Exit Sub

      Dim xl = CType(ExcelDnaUtil.Application, Excel.Application)

      Dim bottomPx As Integer =
            xl.ActiveWindow.ActivePane.PointsToScreenPixelsY(cell.Top + cell.Height)

      Dim rightPx As Integer =
            xl.ActiveWindow.ActivePane.PointsToScreenPixelsX(cell.Left + cell.Width)

      Dim rc As RECT
      GetClientRect(activeButtonOverlay.Handle, rc)
      Dim btnW As Integer = rc.Right - rc.Left
      Dim btnH As Integer = rc.Bottom - rc.Top

      Dim newX As Integer = rightPx
      Dim newY As Integer = bottomPx - btnH

      SetWindowPos(activeButtonOverlay.Handle,
             IntPtr.Zero,
             newX,
             newY,
             0,
             0,
             SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_NOSIZE)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

#End Region
#Region "Helpers"
  ' ==========================================================================================
  ' Routine: ResolveParameterValues
  ' Purpose:
  '   Resolve rule parameters (Address, Name, Offset) into concrete runtime values for the
  '   specific cell being evaluated. Uses worksheet-scoped resolution only.
  '
  ' Parameters:
  '   cell       - Excel.Range representing the cell where the rule is being evaluated.
  '   parameters - IList(Of RuleParameter) defining the ordered parameter list.
  '
  ' Returns:
  '   Object() - Array of resolved parameter values to be passed to the rule engine.
  '
  ' Notes:
  '   - Address parameters are resolved on the same worksheet as the target cell.
  '   - Name parameters are resolved using worksheet-scoped names only.
  '   - Offset parameters are resolved relative to the target cell.
  ' ==========================================================================================
  Private Function ResolveParameterValues(cell As Excel.Range,
                                        parameters As IList(Of ExcelCellRuleStore.RuleParameter)) As Object()

    Dim results As New List(Of Object)
    Dim ws As Excel.Worksheet = cell.Worksheet
    Dim wb As Excel.Workbook = ws.Parent

    Try
      For Each p In parameters

        Select Case p.RefType

          Case ExcelRefType.Address.ToString()
            ' Resolve address on the SAME worksheet as the target cell
            Dim rng As Excel.Range = ws.Range(p.RefValue)
            results.Add(rng.Value2)

          Case ExcelRefType.Name.ToString()
            ' Resolve workbook-scoped name ONLY
            Dim nm As Excel.Name = Nothing
            Dim rng As Excel.Range = Nothing

            Try
              nm = wb.Names.Item(CObj(p.RefValue))

            Catch
              nm = Nothing
            End Try

            If nm IsNot Nothing Then
              rng = nm.RefersToRange
              results.Add(rng.Value2)
            Else
              results.Add(Nothing)
            End If

          Case ExcelRefType.Offset.ToString()
            Dim offset As String = p.RefValue
            Dim rowOffset As Integer = ExtractRowOffset(offset)
            Dim colOffset As Integer = ExtractColOffset(offset)
            Dim rng As Excel.Range = cell.Offset(rowOffset, colOffset)
            results.Add(rng.Value2)

          Case ExcelRefType.Literal.ToString()
            Dim literal As String = p.LiteralValue
            results.Add(literal)
        End Select

      Next

      Return results.ToArray()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "ResolveParameterValues")
      Return results.ToArray()

    Finally
      ws = Nothing
      wb = Nothing
    End Try

  End Function

  ' ==========================================================================================
  ' Routine: ExtractRowOffset
  ' Purpose: Parse the row offset component from an offset string in the form "R{row}C{col}".
  ' Parameters:
  '   offset - String containing the offset expression (e.g. R[-1]C[2], RC[-1], R[3]C)
  ' Returns:
  '   Integer - The parsed row offset.
  ' Notes:
  '   - Assumes the offset string is well-formed.
  '   - Used internally by ResolveParameterValues.
  ' ==========================================================================================
  Private Function ExtractRowOffset(offset As String) As Integer

    Dim rPart As String = offset.Split("C"c)(0).Substring(1) ' after "R"

    If String.IsNullOrEmpty(rPart) Then
      Return 0
    End If

    ' Expecting something like [-1] or [3]
    rPart = rPart.Trim()

    If rPart.StartsWith("[") AndAlso rPart.EndsWith("]") Then
      Dim inner = rPart.Substring(1, rPart.Length - 2)
      Return Integer.Parse(inner)
    End If

    ' If someone ever passes R5C2 style (absolute), handle it:
    Return Integer.Parse(rPart)
  End Function


  ' ==========================================================================================
  ' Routine: ExtractColOffset
  ' Purpose: Parse the column offset component from an offset string in the form "R{row}C{col}".
  ' Parameters:
  '   offset - String containing the offset expression (e.g. R[-1]C[2], RC[-1], R[3]C)
  ' Returns:
  '   Integer - The parsed column offset.
  ' Notes:
  '   - Assumes the offset string is well-formed.
  '   - Used internally by ResolveParameterValues.
  ' ==========================================================================================
  Private Function ExtractColOffset(offset As String) As Integer
    Dim cPart As String = offset.Split("C"c)(1)

    If String.IsNullOrEmpty(cPart) Then
      Return 0
    End If

    cPart = cPart.Trim()

    If cPart.StartsWith("[") AndAlso cPart.EndsWith("]") Then
      Dim inner = cPart.Substring(1, cPart.Length - 2)
      Return Integer.Parse(inner)
    End If

    Return Integer.Parse(cPart)
  End Function

  ' ==========================================================================================
  ' Routine: IsPaste
  ' Purpose:
  '   Determine whether a SheetChange event represents a TRUE paste operation by comparing
  '   BEFORE and AFTER snapshots across the entire Target range.
  '
  '   Paste is the ONLY Excel operation that changes BOTH:
  '       - Cell content (value or formula)
  '       - Cell GUID identity (hidden RM_* name)
  '
  ' Parameters:
  '   beforeSnap - RangeSnapshot captured during the previous SelectionChange.
  '   afterSnap  - RangeSnapshot captured after the SheetChange via QueueAsMacro.
  '
  ' Returns:
  '   Boolean - True if ANY cell in the range shows the paste signature.
  '
  ' Contract:
  '   - Caller must supply non-null snapshots for the same range shape.
  '   - Paste is defined as: (Value OR Formula changed) AND (Guid changed).
  '   - Excel fires ONE SheetChange for the entire paste transaction.
  ' ==========================================================================================
  Private Function IsPaste(beforeSnap As RangeSnapshot,
                         afterSnap As RangeSnapshot) As Boolean
    Try
      If beforeSnap Is Nothing OrElse afterSnap Is Nothing Then
        Return False
      End If

      ' Iterate through each cell in the BEFORE snapshot.
      For Each addr In beforeSnap.Cells.Keys
        Dim b = beforeSnap.Cells(addr)
        ' Try to get the AFTER cell safely
        Dim a As CellSnapshot = Nothing
        If Not afterSnap.Cells.TryGetValue(addr, a) Then
          ' AFTER snapshot doesn't contain this address — shape changed.
          ' For paste detection, we can safely ignore this address.
          Continue For
        End If
        ' Content changed?
        Dim contentChanged As Boolean =
        Not Object.Equals(b.Value, a.Value) _
        OrElse b.Formula <> a.Formula
        If Not contentChanged AndAlso b.Guid = a.Guid Then Continue For
        ' Identity changed?
        If b.Guid <> a.Guid Then
          ' This change matches our paste signature.
          Return True
        End If
      Next

      ' No cell matched the paste signature.
      Return False
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return False
    End Try

  End Function

  ' ==========================================================================================
  ' Routine: SnapshotRange
  ' Purpose:
  '   Capture the minimal identity-relevant state of each cell in the range so that paste operations
  '   can be detected deterministically by comparing BEFORE and AFTER snapshots.
  '
  '   Snapshot includes:
  '       - Value2      (raw value)
  '       - Formula     (string formula)
  '       - Guid        (cell identity derived from hidden RM_* name)
  '
  ' Parameters:
  '   rng - Excel.Range (can be one or many cells).
  '
  ' Returns:
  '   RangeSnapshot - A lightweight dictionary of object containing the captured state.
  '
  ' Contract:
  '   - SnapshotRange NEVER mutates workbook state.
  ' ==========================================================================================
  Private Function SnapshotRange(rng As Excel.Range) As RangeSnapshot
    Dim snap As New RangeSnapshot With {
        .Cells = New Dictionary(Of String, CellSnapshot)
    }

    For Each cell As Excel.Range In rng.Cells
      Dim cs As New CellSnapshot
      Try : cs.Value = cell.Value2 : Catch : cs.Value = Nothing : End Try
      Try : cs.Formula = CStr(cell.Formula) : Catch : cs.Formula = Nothing : End Try
      cs.Guid = cell.ID

      snap.Cells(cell.Address(False, False)) = cs
    Next cell

    Return snap
  End Function

  ' ==========================================================================================
  ' Routine: SameAddressSet
  ' Purpose:
  '   Determine whether two RangeSnapshot instances represent the SAME logical set of cells.
  '
  '   This guard is required because CUT operations trigger TWO SheetChange events:
  '       1. Paste target  (range shape matches BEFORE snapshot)
  '       2. Source clear  (range shape DOES NOT match BEFORE snapshot)
  '
  '   Only the first event is a candidate for paste detection. The second must be ignored.
  '
  ' Parameters:
  '   b - BEFORE RangeSnapshot captured during the last SelectionChange.
  '   a - AFTER  RangeSnapshot captured after the SheetChange via QueueAsMacro.
  '
  ' Returns:
  '   Boolean - True if BOTH snapshots contain identical address sets; False otherwise.
  '
  ' Contract:
  '   - Caller must supply non-null snapshots.
  '   - Address comparison is case-insensitive and uses A1-style addresses.
  '   - Mismatched shapes MUST NOT be passed to IsPaste; they represent non-paste events
  '     (typically the CUT source being cleared).
  ' ==========================================================================================
  Private Function SameAddressSet(b As RangeSnapshot,
                                a As RangeSnapshot) As Boolean

    If b Is Nothing OrElse a Is Nothing Then Return False
    If b.Cells.Count <> a.Cells.Count Then Return False

    For Each addr In b.Cells.Keys
      If Not a.Cells.ContainsKey(addr) Then Return False
    Next

    Return True
  End Function

  ' ********************************************************************************************
  ' Routine:    PropagateMissingIdsFromSourceTile
  ' Purpose:    Excel tiles values/formats when the target range is an exact multiple of the
  '             source range, but does NOT fully propagate Range.ID metadata (especially in
  '             vertically repeated tiles).
  '
  '             This routine reconstructs the tiling pattern and, for each target cell that
  '             does NOT have an ID, copies the corresponding source cell's ID into it.
  '
  ' Contract:
  '   - srcRange:  The range that was copied (captured at Ctrl+C time).
  '   - tgtRange:  The range selected at paste time (before and after paste).
  '
  ' Behaviour:
  '   - If tgtRange dimensions are NOT exact multiples of srcRange, routine exits safely.
  '   - For each cell in tgtRange:
  '       - Compute which source cell it corresponds to (tiling).
  '       - If target ID is empty and source ID is non-empty, assign target.ID = source.ID.
  '
  ' Notes:
  '   - This routine does NOT create new GUIDs.
  '   - This routine does NOT touch XML or names.
  '   - This routine does NOT overwrite existing IDs.
  '   - It exists purely to "finish" Excel's metadata tiling so that later logic
  '     (e.g. PasteIdentityHandler) can run over a consistent identity surface.
  '
  ' ********************************************************************************************
  Friend Sub PropagateMissingIdsFromSourceTile(
          ByVal srcRange As Excel.Range,
          ByVal tgtRange As Excel.Range)

    Try
      ' -----------------------------
      ' 1. Source dimensions
      ' -----------------------------
      Dim srcRows As Integer = srcRange.Rows.Count
      Dim srcCols As Integer = srcRange.Columns.Count

      ' -----------------------------
      ' 2. Target dimensions
      ' -----------------------------
      Dim tgtRows As Integer = tgtRange.Rows.Count
      Dim tgtCols As Integer = tgtRange.Columns.Count

      ' ------------------------------------------------------------------------
      ' 3. Only act if target is an exact multiple of source shape and not equal
      ' ------------------------------------------------------------------------
      If srcRows = tgtRows AndAlso srcCols = tgtCols Then Exit Sub ' If target shape equals source shape, Excel already propagated IDs correctly.
      If (tgtRows Mod srcRows) <> 0 Then Exit Sub
      If (tgtCols Mod srcCols) <> 0 Then Exit Sub

      Dim r As Integer, c As Integer

      ' ----------------------------------------------------------
      ' 4. Walk target and copy IDs from tiled source where missing
      ' ----------------------------------------------------------
      For r = 1 To tgtRows
        For c = 1 To tgtCols

          Dim tgtCell As Excel.Range = tgtRange.Cells(r, c)

          ' 4A. Read target ID
          Dim tgtId As String = ""
          Try
            Dim objT = tgtCell.GetType().InvokeMember("ID",
                            Reflection.BindingFlags.GetProperty,
                            Nothing, tgtCell, Nothing)
            If objT IsNot Nothing Then tgtId = CStr(objT)
          Catch
            tgtId = ""
          End Try

          ' If target already has an ID, do not touch it
          If tgtId <> "" Then Continue For

          ' 4B. Compute corresponding source cell via tiling
          Dim srcRow As Integer = ((r - 1) Mod srcRows) + 1
          Dim srcCol As Integer = ((c - 1) Mod srcCols) + 1
          Dim srcCell As Excel.Range = srcRange.Cells(srcRow, srcCol)

          ' 4C. Read source ID
          Dim srcId As String = ""
          Try
            Dim objS = srcCell.GetType().InvokeMember("ID",
                            Reflection.BindingFlags.GetProperty,
                            Nothing, srcCell, Nothing)
            If objS IsNot Nothing Then srcId = CStr(objS)
          Catch
            srcId = ""
          End Try

          ' If source has no ID, nothing to propagate
          If srcId = "" Then Continue For

          ' 4D. Assign target.ID = source.ID
          SetRangeIdValue(tgtCell, srcId)

        Next c
      Next r

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "PropagateMissingIdsFromSourceTile")
    End Try

  End Sub

  ' ==========================================================================================
  ' Routine: ExcelIsEditing
  ' Purpose: Returns True if Excel is currently in edit mode (cell being edited).
  ' Parameters:
  '   None
  ' Returns:
  '   True if Excel is in edit mode; False otherwise.
  ' Notes:
  '   When Excel enters edit mode all edit CommandBars become disabled
  ' ==========================================================================================
  Private Function ExcelIsEditing() As Boolean
    Try
      Dim xl = CType(ExcelDnaUtil.Application, Excel.Application)
      Dim editMenu = xl.CommandBars("Worksheet Menu Bar").Controls("Edit")
      Return Not editMenu.Enabled
    Catch
      Return False
    End Try
  End Function
#End Region
End Module
