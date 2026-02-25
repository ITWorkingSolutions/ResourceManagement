Imports System.Xml.Linq
Imports ExcelDna.Integration
Imports Microsoft
Imports Microsoft.Office.Core
Imports Excel = Microsoft.Office.Interop.Excel

' ==========================================================================================
' Module: ExcelCellRuleStore
' Purpose: Provides GUID-based per-cell identity using hidden Defined Names and stores
'          rule application metadata in a workbook-level Custom XML Part, grouped by
'          RuleRegion. Each RuleRegion represents a single rule application signature
'          (rule name, list type, ordered parameter list) and tracks all cells (by GUID)
'          that share that signature.
' Notes:
'   - Replaces all Range.ID-based logic.
'   - Each participating cell gets a GUID stored as a hidden workbook-level Defined Name.
'   - RuleRegions are identified by a stable GUID-based ID (RR_<guid>).
'   - RuleRegions have a "state" attribute to support future "NeedsRepair" semantics when
'     rule definitions change.
' ==========================================================================================
Friend Module ExcelCellRuleStore

  ' XML namespace and element/attribute constants.
  Private Const XmlNamespace As String = "http://schemas.resourcemanager/rules"
  Private Const RootElementName As String = "ResourceManagement"
  Private Const RuleRegionElementName As String = "RuleRegion"
  Private Const RuleElementName As String = "Rule"
  Private Const ParametersElementName As String = "Parameters"
  Private Const ParameterElementName As String = "Parameter"
  Private Const CellsElementName As String = "Cells"
  Private Const CellElementName As String = "Cell"

  ' Reverse map XML constants.
  Private Const ReverseMapRootElementName As String = "ReverseMap"
  Private Const ReverseMapEntryElementName As String = "Entry"
  Private Const ReverseMapCellAttrName As String = "cell"
  Private Const ReverseMapRegionAttrName As String = "region"

  ' RuleRegion ID prefix and state values.
  Private Const RuleRegionIdPrefix As String = "RR_"
  Private Const RuleRegionStateValid As String = "Valid"
  Private Const RuleRegionStateNeedsRepair As String = "NeedsRepair"

  ' Prefix for hidden Defined Names used to store per-cell GUID identity.
  Private Const NamePrefix As String = "RM_"

  ' ==========================================================================================
  ' Class: RuleParameter
  ' Purpose: Represents a single rule parameter for a rule application. The combination of
  '          all parameters (in order) forms part of the rule signature used for grouping.
  ' Properties:
  '   Name     - Logical name of the parameter (e.g., "Month", "Year", "Role").
  '   RefType  - Reference type enumerated by ExcelRefType
  '   RefValue - Reference value (e.g., "Sheet1!$B$2", "RosterYear", "R0C-2").
  '   LiteralValue - Literal value entered
  ' Notes:
  '   - Mode (Fixed/Relative) is not stored here; it is implied by RefType and RefValue.
  ' ==========================================================================================
  Friend Class RuleParameter

    Public Property Name As String
    Public Property RefType As String
    Public Property RefValue As String
    Public Property LiteralValue As String

  End Class

  ' ==========================================================================================
  ' Class: ExcelApplyInstance
  ' Purpose:
  '   Storage-layer representation of a RuleRegion (Apply instance).
  '   Returned to UILoaderSaver so it can map to UIExcelApplyInstance.
  ' Properties:
  '   ApplyID        - RuleRegion ID (e.g., "RR_<guid>")
  '   RuleID         - Raw GUID of the rule definition
  '   ListSelectType - List selection type stored in XML
  '   Parameters     - Ordered list of RuleParameter objects
  '   CellGuids      - List of bare GUIDs for all cells in this region
  ' ==========================================================================================
  Friend Class ExcelApplyInstance
    Public Property ApplyID As String
    Public Property ApplyName As String
    Public Property RuleID As String
    Public Property ListSelectType As String
    Public Property Parameters As List(Of RuleParameter)
    Public Property CellGuids As List(Of String)
    Public Property IsNew As Boolean
    Public Property IsDirty As Boolean
    Public Property IsDeleted As Boolean

  End Class

  Friend Sub InitializeGuidIdentityForWorkbook(wb As Excel.Workbook)

    Try
      ' -------------------------------------------------------------
      ' Pass 0: Load ReverseMap and build set of valid GUIDs.
      ' -------------------------------------------------------------
      Dim part As Object = Nothing
      Dim doc As XDocument = LoadReverseMap(wb, part)

      Dim xmlGuids As New HashSet(Of String)(StringComparer.Ordinal)
      For Each e In doc.Root.Elements()
        xmlGuids.Add(CStr(e.Attribute("cell")))
      Next

      ' -------------------------------------------------------------
      ' Pass 1: Group identity names by the cell they refer to.
      '         Also delete broken names (#REF!, invalid RefersTo).
      ' -------------------------------------------------------------
      Dim groups As New Dictionary(Of String, List(Of Excel.Name))(StringComparer.Ordinal)

      For Each nm As Excel.Name In wb.Names
        If Not nm.Visible AndAlso nm.Name.StartsWith(NamePrefix, StringComparison.Ordinal) Then

          ' Delete broken names immediately
          Dim refersTo As String = nm.RefersTo
          If String.IsNullOrEmpty(refersTo) OrElse refersTo.IndexOf("#REF!", StringComparison.OrdinalIgnoreCase) >= 0 Then
            RemoveGuidFromCustomXml(wb, nm.Name)
            nm.Delete()
            Continue For
          End If

          ' Resolve RefersToRange safely
          Dim rng As Excel.Range = Nothing
          Try
            rng = nm.RefersToRange
          Catch
            rng = Nothing
          End Try

          If rng Is Nothing OrElse rng.CountLarge <> 1L Then
            RemoveGuidFromCustomXml(wb, nm.Name)
            nm.Delete()
            Continue For
          End If

          ' Group by cell address
          Dim addr As String = rng.Worksheet.Name & "!" & rng.Address(False, False)
          If Not groups.ContainsKey(addr) Then
            groups(addr) = New List(Of Excel.Name)
          End If
          groups(addr).Add(nm)

        End If
      Next

      ' -------------------------------------------------------------
      ' Pass 2: Resolve duplicates and stale names.
      '         Keep exactly ONE identity name per cell.
      ' -------------------------------------------------------------
      For Each kvp In groups
        Dim nameList = kvp.Value

        If nameList.Count <= 1 Then
          ' Single identity name: ensure it exists in XML
          Dim nm = nameList(0)
          If Not xmlGuids.Contains(nm.Name) Then
            AddGuidToCustomXml(wb, nm.Name, nm.RefersToRange)
          End If
          Continue For
        End If

        ' Multiple identity names → choose the best one
        Dim best As Excel.Name = Nothing

        ' Prefer names that exist in XML
        Dim xmlMatches = nameList.Where(Function(n) xmlGuids.Contains(n.Name)).ToList()

        If xmlMatches.Count = 1 Then
          best = xmlMatches(0)
        ElseIf xmlMatches.Count > 1 Then
          ' Keep lexically newest GUID
          best = xmlMatches.OrderBy(Function(n) n.Name).Last()
        Else
          ' No XML match: keep lexically newest
          best = nameList.OrderBy(Function(n) n.Name).Last()
        End If

        ' Delete all others
        For Each nm In nameList
          If nm IsNot best Then
            RemoveGuidFromCustomXml(wb, nm.Name)
            nm.Delete()
          End If
        Next

        ' Ensure XML contains the surviving GUID
        If Not xmlGuids.Contains(best.Name) Then
          AddGuidToCustomXml(wb, best.Name, best.RefersToRange)
        End If
      Next

      ' -------------------------------------------------------------
      ' Pass 3: Rehydrate Range.ID for all surviving identity names.
      ' -------------------------------------------------------------
      For Each nm As Excel.Name In wb.Names
        If Not nm.Visible AndAlso nm.Name.StartsWith(NamePrefix, StringComparison.Ordinal) Then
          Dim targetRange As Excel.Range = Nothing
          Try
            targetRange = nm.RefersToRange
          Catch
            targetRange = Nothing
          End Try
          If targetRange Is Nothing Then Continue For
          If targetRange.CountLarge <> 1L Then Continue For

          SetRangeIdValue(targetRange, nm.Name)
        End If
      Next

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "InitializeGuidIdentityForWorkbook")
    End Try

  End Sub

  ' ==========================================================================================
  ' Routine: PasteIdentityHandler
  ' Purpose:
  '   Handle identity propagation after a paste operation. Excel does not provide a dedicated
  '   paste event, but SheetChange fires reliably. This routine determines whether each cell
  '   in the Target range represents:
  '       - A NEW identity (no Range.ID)
  '       - A CUT (Range.ID exists and the associated name's RefersTo could be broken or changed to target by excel)
  '       - A COPY (Range.ID exists and the associated name's RefersTo is still valid)
  '
  '   Based on the classification, the routine:
  '       - Creates new GUID identities for NEW or COPY cases
  '       - Updates existing workbook-level names for CUT cases
  '       - Updates Range.ID for the session
  '       - Duplicates XML RuleRegion membership for COPY cases
  '
  ' Parameters:
  '   ws     - Worksheet containing the changed cells.
  '   target - Range representing the cells modified by the paste.
  '
  ' Contract:
  '   - Identity names are workbook-scoped and hidden.
  '   - Range.ID stores the full identity name (NamePrefix + GUID).
  '   - CUT is detected by invalid range in the name's RefersTo or a match to the target cell.
  '   - COPY is detected when RefersTo is still valid but doesn't point to the target cell.
  ' ==========================================================================================
  Friend Sub PasteIdentityHandler(ws As Excel.Worksheet, target As Excel.Range)

    Dim wb As Excel.Workbook = CType(ws.Parent, Excel.Workbook)

    Try
      For Each cell As Excel.Range In target.Cells
        ' Read Range.ID safely (this is our identity name, e.g. RM_<guid>)
        Dim idValue As String = String.Empty
        Try
          Dim obj = cell.GetType().InvokeMember("ID",
                                              Reflection.BindingFlags.GetProperty,
                                              Nothing,
                                              cell,
                                              Nothing)
          If obj IsNot Nothing Then idValue = CStr(obj)
        Catch
          idValue = String.Empty
        End Try

        ' ---------------------------------------------------------
        ' CASE 1: No ID → cell has no identity yet → new identity
        ' ---------------------------------------------------------
        If String.IsNullOrEmpty(idValue) Then
          EnsureCellHasGuidName(cell)
          Continue For
        End If

        ' ---------------------------------------------------------
        ' CASE 2: ID exists → resolve the identity name
        ' ---------------------------------------------------------
        Dim nm As Excel.Name = Nothing

        For Each candidate As Excel.Name In wb.Names
          If String.Equals(candidate.Name, idValue, StringComparison.OrdinalIgnoreCase) Then
            nm = candidate
            Exit For
          End If
        Next

        ' If name does not exist → treat as new identity (defensive)
        If nm Is Nothing Then
          EnsureCellHasGuidName(cell)
          Continue For
        End If

        ' ---------------------------------------------------------
        ' Resolve RefersToRange safely
        ' ---------------------------------------------------------
        Dim refersRange As Excel.Range = Nothing
        Try
          refersRange = nm.RefersToRange
        Catch
          refersRange = Nothing
        End Try

        ' ---------------------------------------------------------
        ' CASE 2A: Orphaned / broken name (e.g. #REF!) → repair as CUT
        ' ---------------------------------------------------------
        If refersRange Is Nothing Then

          Dim newRef As String = "=" & cell.Address(RowAbsolute:=True, ColumnAbsolute:=True)
          nm.RefersTo = newRef

          ExcelAsyncUtil.QueueAsMacro(
          Sub()
            Dim xl = CType(ExcelDnaUtil.Application, Excel.Application)
            Dim refreshed = xl.Range(cell.Address)
            lastOverlayCellAddress = Nothing
            ExcelSelectionChangeHandler(refreshed)
          End Sub)

          Continue For
        End If

        ' ---------------------------------------------------------
        ' CASE 2B: CUT — name already points to this cell
        ' ---------------------------------------------------------
        If refersRange.Worksheet.Name = cell.Worksheet.Name AndAlso
         refersRange.Address = cell.Address Then

          ' Normal cut-paste: identity has moved here.
          ' Ensure XML/overlay are correct, but DO NOT create a new GUID.
          ExcelAsyncUtil.QueueAsMacro(
          Sub()
            lastOverlayCellAddress = Nothing
            ExcelSelectionChangeHandler(cell)
          End Sub)

          Continue For
        End If

        ' ---------------------------------------------------------
        ' CASE 2C: COPY — name still points to source cell
        ' ---------------------------------------------------------
        If refersRange.Worksheet.Name <> cell.Worksheet.Name _
          OrElse refersRange.Address <> cell.Address Then

          ExcelAsyncUtil.QueueAsMacro(
          Sub()
            Try
              ' Clean up any existing identities for this cell
              RemoveAllIdentitiesForCell(wb, cell)
              ' Create new GUID + workbook-level name
              Dim newGuid As String = System.Guid.NewGuid().ToString("N")
              Dim newName As String = NamePrefix & newGuid
              Dim newRefersTo As String = "=" & cell.Address(RowAbsolute:=True, ColumnAbsolute:=True)

              Dim newNm As Excel.Name = wb.Names.Add(Name:=newName, RefersTo:=newRefersTo)
              newNm.Visible = False
              ' Update Range.ID to the new identity name
              SetRangeIdValue(cell, newName)
              ' Duplicate XML membership from old identity to new
              DuplicateGuidXmlEntry(wb, oldName:=nm.Name, newName:=newName)
              ' Force overlay refresh for this cell
              lastOverlayCellAddress = Nothing
              ExcelSelectionChangeHandler(cell)

            Catch ex As Exception
              ErrorHandler.UnHandleError(ex, "PasteIdentityHandler-COPY")
            End Try
            ' ---- Final ID after paste ----
            Dim finalId As String = ""
            Try
              Dim obj2 = cell.GetType().InvokeMember("ID",
                                               Reflection.BindingFlags.GetProperty,
                                               Nothing,
                                               cell,
                                               Nothing)
              If obj2 IsNot Nothing Then finalId = CStr(obj2)
            Catch
              finalId = ""
            End Try

          End Sub)
        End If
      Next

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "PasteIdentityHandler")
    End Try

  End Sub

  ' ==========================================================================================
  ' Routine: EnsureCellHasGuidName
  ' Purpose: Ensure the specified cell has a hidden Defined Name whose name encodes a GUID.
  ' Parameters:
  '   cell - Excel.Range representing the target cell.
  ' Returns:
  '   String - The GUID string associated with the cell.
  ' Notes:
  '   - If a matching hidden name already exists, its GUID is returned.
  '   - If not, a new GUID is generated and a hidden worksheet-level name is created.
  '   - The Defined Name uses the pattern "__RM_{GUID}" and RefersTo the cell.
  ' ==========================================================================================

  Friend Function EnsureCellHasGuidName(cell As Excel.Range) As String

    Dim existingGuid As String = GetGuidFromCell(cell)
    If Not String.IsNullOrEmpty(existingGuid) Then
      Return existingGuid
    End If

    Dim wb As Excel.Workbook = cell.Worksheet.Parent
    Dim guid As String = System.Guid.NewGuid().ToString("N")   ' e.g. "8c87d665-d5f5-40cc-bb54-2313f93cda7d" to "8c87d665d5f540ccbb542313f93cda7d"
    Dim nameText As String = NamePrefix & guid                 ' e.g. "RM_8c87d665d5f540ccbb542313f93cda7d"

    ' Absolute reference for stable identity
    Dim refersTo As String = "=" & cell.Address(RowAbsolute:=True, ColumnAbsolute:=True)

    ExcelAsyncUtil.QueueAsMacro(
        Sub()
          Try
            ' Create the workbook-level identity name
            Dim nm As Excel.Name = wb.Names.Add(Name:=nameText, RefersTo:=refersTo)
            nm.Visible = False

            ' IMPORTANT: Set Range.ID immediately for this session
            SetRangeIdValue(cell, nameText)

          Catch ex As Exception
            ErrorHandler.UnHandleError(ex, "EnsureCellHasGuidName")
          End Try

        End Sub)

    Return guid

  End Function

  ' ==========================================================================================
  ' Routine: GetGuidFromCell
  ' Purpose:
  '   Resolve the GUID associated with a given cell by inspecting hidden, workbook-scoped
  '   identity names that start with NamePrefix.
  '
  ' Parameters:
  '   cell - Excel.Range representing the target cell.
  '
  ' Returns:
  '   String - GUID in "N" format if found; otherwise an empty string.
  ' ==========================================================================================
  Friend Function GetGuidFromCell(cell As Excel.Range) As String

    Try
      Dim wb As Excel.Workbook = CType(cell.Worksheet.Parent, Excel.Workbook)

      For Each nm As Excel.Name In wb.Names

        If Not nm.Visible AndAlso nm.Name.StartsWith(NamePrefix, StringComparison.Ordinal) Then

          Dim targetRange As Excel.Range = Nothing

          Try
            targetRange = nm.RefersToRange
          Catch
            targetRange = Nothing
          End Try

          If targetRange Is Nothing Then Continue For
          If targetRange.CountLarge <> 1L Then Continue For

          If targetRange.Worksheet Is cell.Worksheet _
           AndAlso targetRange.Row = cell.Row _
           AndAlso targetRange.Column = cell.Column Then

            Dim rawName As String = nm.Name
            Dim start As Integer = NamePrefix.Length

            If rawName.Length >= start + 32 Then
              Return rawName.Substring(start, 32)
            End If

          End If

        End If
      Next

      Return String.Empty

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "GetGuidFromCell")
      Return String.Empty
    End Try

  End Function

  ' ==========================================================================================
  ' Routine: SetCellRule
  ' Purpose:
  '   Assign a rule to a cell using stable GUID-based RegionID and update reverse map.
  ' Parameters:
  '   cell           - Excel.Range target cell.
  '   ruleId       - Rule Id.
  '   listSelectType - List selection type.
  '   parameters     - Ordered parameter list.
  ' Returns:
  '   None.
  ' Notes:
  '   - RegionID is RuleRegionIdPrefix + GUID (e.g., "RR_8c87...").
  '   - Reverse map stores RM_<guid> → RR_<guid>.
  ' ==========================================================================================
  Friend Sub SetCellRule(
    cell As Excel.Range,
    applyName As String,
    ruleId As String,
    listSelectType As String,
    parameters As IList(Of RuleParameter))

    Dim wb As Excel.Workbook = Nothing
    Dim part As Object = Nothing
    Dim doc As XDocument = Nothing
    Dim guid As String = Nothing
    Dim cellName As String = Nothing
    Dim regionId As String = Nothing
    Dim rr As XElement = Nothing

    Try
      ' --- Normal execution ---
      wb = CType(cell.Worksheet.Parent, Excel.Workbook)

      ' Ensure cell GUID (bare) and build full identity name (RM_<guid>)
      guid = EnsureCellHasGuidName(cell)
      cellName = NamePrefix & guid

      ' Load XML
      doc = LoadXmlDocument(wb, part)

      ' Try reverse map lookup first: RM_<guid> → RR_<guid>
      regionId = GetRegionIdForCellName(wb, cellName)
      If Not String.IsNullOrEmpty(regionId) Then
        rr = FindRuleRegionById(doc, regionId)
      End If

      ' If no region found, create a new one
      If rr Is Nothing Then
        rr = CreateRuleRegion(doc, Nothing, applyName, ruleId, listSelectType, parameters)
        regionId = CStr(rr.Attribute("id"))
      End If

      ' Add GUID (bare) to region's <Cells>
      AddGuidToRuleRegion(rr, guid)

      ' Update reverse map: RM_<guid> → RR_<guid>
      AddReverseMapEntry(wb, cellName, regionId)

      ' Save XML
      SaveXmlDocument(wb, part, doc)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "SetCellRule")

    Finally
      ' --- Cleanup ---
      wb = Nothing
      part = Nothing
      doc = Nothing
      rr = Nothing

    End Try

  End Sub

  ' ==========================================================================================
  ' Routine: ClearCellRule
  ' Purpose:
  '   Remove rule metadata association for a specific cell using reverse map lookup.
  '   Removes the GUID from its RuleRegion and removes the reverse map entry.
  ' Parameters:
  '   cell - Excel.Range representing the target cell.
  ' Returns:
  '   None.
  ' Notes:
  '   Deletes the RuleRegion if it becomes empty.
  ' ==========================================================================================
  Friend Sub ClearCellRule(cell As Excel.Range)

    Dim wb As Excel.Workbook = Nothing
    Dim part As Object = Nothing
    Dim doc As XDocument = Nothing
    Dim cellGuid As String = Nothing
    Dim cellName As String = Nothing
    Dim regionId As String = Nothing
    Dim rr As XElement = Nothing
    Dim cellsElem As XElement = Nothing

    Try
      ' --- Normal execution ---
      wb = CType(cell.Worksheet.Parent, Excel.Workbook)

      ' Get the cell's GUID (bare)
      cellGuid = GetGuidFromCell(cell)
      If String.IsNullOrEmpty(cellGuid) Then Exit Sub

      ' Full identity name
      cellName = NamePrefix & cellGuid

      ' Lookup region ID via reverse map
      regionId = GetRegionIdForCellName(wb, cellName)
      If String.IsNullOrEmpty(regionId) Then Exit Sub

      ' Load main XML
      doc = LoadXmlDocument(wb, part)

      ' Find the region
      rr = FindRuleRegionById(doc, regionId)
      If rr Is Nothing Then
        SaveXmlDocument(wb, part, doc)
        Exit Sub
      End If

      ' Remove GUID from region
      If Not RemoveGuidFromRuleRegion(rr, cellGuid) Then
        SaveXmlDocument(wb, part, doc)
        Exit Sub
      End If

      ' Remove reverse map entry
      RemoveReverseMapEntry(wb, cellName)

      ' Delete region if empty
      cellsElem = rr.Element(XName.Get(CellsElementName, XmlNamespace))
      If cellsElem Is Nothing OrElse
       Not cellsElem.Elements(XName.Get(CellElementName, XmlNamespace)).Any() Then
        rr.Remove()
      End If

      ' Save XML
      SaveXmlDocument(wb, part, doc)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "ClearCellRule")

    Finally
      ' --- Cleanup ---
      wb = Nothing
      part = Nothing
      doc = Nothing
      rr = Nothing
      cellsElem = Nothing

    End Try

  End Sub

  ' ==========================================================================================
  ' Routine: TryGetRuleForCell
  ' Purpose:
  '   Retrieve the rule metadata associated with a specific cell using the reverse map for
  '   fast lookup. Extracts rule name, list type, parameters, and region state.
  ' Parameters:
  '   cell           - Excel.Range representing the target cell.
  '   ruleId         - [ByRef] Output: rule id.
  '   listSelectType - [ByRef] Output: list select type.
  '   parameters     - [ByRef] Output: ordered parameter list.
  '   regionState    - [ByRef] Output: region state ("Valid" or "NeedsRepair").
  ' Returns:
  '   Boolean - True if rule metadata exists for the cell; otherwise False.
  ' Notes:
  '   Uses RM_guid → RR_guid reverse map for O(1) lookup.
  ' ==========================================================================================
  Friend Function TryGetRuleForCell(cell As Excel.Range,
                                  ByRef ruleId As String,
                                  ByRef listSelectType As String,
                                  ByRef parameters As IList(Of RuleParameter),
                                  ByRef regionState As String) As Boolean

    Dim wb As Excel.Workbook = Nothing
    Dim part As Object = Nothing
    Dim doc As XDocument = Nothing
    Dim cellGuid As String = Nothing
    Dim cellName As String = Nothing
    Dim regionId As String = Nothing
    Dim rr As XElement = Nothing

    Try

      ' --- Normal execution ---
      wb = CType(cell.Worksheet.Parent, Excel.Workbook)

      ' --- PRIMARY: use Range.ID as the identity name ---
      Dim idValue As String = String.Empty
      Try
        Dim obj = cell.GetType().InvokeMember("ID",
                                            Reflection.BindingFlags.GetProperty,
                                            Nothing,
                                            cell,
                                            Nothing)
        If obj IsNot Nothing Then idValue = CStr(obj)
      Catch
        idValue = String.Empty
      End Try

      If Not String.IsNullOrEmpty(idValue) AndAlso idValue.StartsWith(NamePrefix, StringComparison.Ordinal) Then
        cellName = idValue
        cellGuid = cellName.Substring(NamePrefix.Length)
      Else
        ' --- FALLBACK: legacy scan by names (GetGuidFromCell) ---
        cellGuid = GetGuidFromCell(cell) ' Get the cell's GUID (bare 32 hex chars)
        If String.IsNullOrEmpty(cellGuid) Then Return False
        cellName = NamePrefix & cellGuid  ' Convert to full identity name (RM_<guid>)
      End If

      ' Lookup region ID using reverse map
      regionId = GetRegionIdForCellName(wb, cellName)
      If String.IsNullOrEmpty(regionId) Then Return False


      ' Load main XML
      doc = LoadXmlDocument(wb, part)

      ' Find the RuleRegion by ID
      rr = FindRuleRegionById(doc, regionId)
      If rr Is Nothing Then Return False

      ' Extract state
      regionState = CStr(rr.Attribute("state"))

      ' Extract rule metadata
      Dim ruleElem As XElement = rr.Element(XName.Get(RuleElementName, XmlNamespace))
      ruleId = CStr(ruleElem.Attribute("ruleId"))
      listSelectType = CStr(ruleElem.Attribute("listSelectType"))

      ' Extract parameters
      parameters = New List(Of RuleParameter)
      Dim paramsElem As XElement = ruleElem.Element(XName.Get(ParametersElementName, XmlNamespace))

      For Each pElem In paramsElem.Elements(XName.Get(ParameterElementName, XmlNamespace))
        Dim p As New RuleParameter With {
        .Name = CStr(pElem.Attribute("name")),
        .RefType = CStr(pElem.Attribute("refType")),
        .RefValue = CStr(pElem.Attribute("refValue")),
        .LiteralValue = CStr(pElem.Attribute("literalValue"))
      }
        parameters.Add(p)
      Next

      Return True

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "TryGetRuleForCell")
      Return False

    Finally
      ' --- Cleanup ---
      wb = Nothing
      part = Nothing
      doc = Nothing
      rr = Nothing

    End Try

  End Function

  ' ==========================================================================================
  ' Routine: RemoveAllIdentitiesForCell
  ' Purpose: Removes all identity names and associated XML metadata for a given cell. Used when a cell is deleted
  ' Parameters:
  '   wb   - Excel.Workbook containing the identity names and XML metadata.
  '   cell - Excel.Range representing the target cell.  
  ' Returns:
  '   none
  ' Notes:
  '   {notes}
  ' ==========================================================================================
  Friend Sub RemoveAllIdentitiesForCell(wb As Excel.Workbook, cell As Excel.Range)
    For Each nm As Excel.Name In wb.Names
      If Not nm.Visible AndAlso nm.Name.StartsWith(NamePrefix, StringComparison.Ordinal) Then
        Dim r As Excel.Range = Nothing
        Try
          r = nm.RefersToRange
        Catch
          r = Nothing
        End Try
        If r Is Nothing Then Continue For
        If r.Worksheet Is cell.Worksheet AndAlso r.Address = cell.Address Then
          ' Remove from XML + delete name
          RemoveGuidFromCustomXml(wb, nm.Name)
          nm.Delete()
        End If
      End If
    Next
  End Sub


#Region "XML Helpers"
  ' ==========================================================================================
  ' Routine: GetOrCreateXmlPart
  ' Purpose: Retrieve the Custom XML Part used by ResourceManagement, or create it if missing.
  ' Parameters:
  '   wb - Excel.Workbook containing the XML parts.
  ' Returns:
  '   Object - The XML part COM object whose root element is <ResourceManagement>.
  ' Notes:
  '   Ensures a single authoritative XML part exists for all ResourceManagement metadata.
  ' ==========================================================================================
  Private Function GetOrCreateXmlPart(wb As Excel.Workbook) As Object

    Dim part As Object
    Dim xml As String
    Dim doc As XDocument
    Try

      ' Iterate existing CustomXMLParts and find one with the correct root element.
      For Each part In wb.CustomXMLParts
        xml = CStr(part.XML)
        If Not String.IsNullOrEmpty(xml) Then
          doc = XDocument.Parse(xml)
          If doc.Root IsNot Nothing AndAlso doc.Root.Name.LocalName = RootElementName Then
            Return part
          End If
        End If
      Next

      ' Create a new XML part if none exists.
      xml = "<" & RootElementName & " xmlns='" & XmlNamespace & "'></" & RootElementName & ">"
      Return wb.CustomXMLParts.Add(xml)
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "{RoutineName}")
      Return Nothing
    Finally
      ' --- Cleanup ---
      part = Nothing
      doc = Nothing
    End Try

  End Function

  ' ==========================================================================================
  ' Routine: LoadXmlDocument
  ' Purpose: Load the XDocument for the ResourceManagement XML part.
  ' Parameters:
  '   wb   - Excel.Workbook containing the XML parts.
  '   part - [ByRef] The Custom XML Part COM object returned.
  ' Returns:
  '   XDocument - Parsed XML document for the ResourceManagement part.
  ' Notes:
  '   Always returns a valid document with a <ResourceManagement> root, even if the existing
  '   part is empty or malformed.
  ' ==========================================================================================
  Private Function LoadXmlDocument(wb As Excel.Workbook, ByRef part As Object) As XDocument

    Dim xml As String
    Dim doc As XDocument
    Try
      ' Get or create the underlying CustomXMLPart.
      part = GetOrCreateXmlPart(wb)
      xml = CStr(part.XML)

      ' If the part is empty or invalid, create a new document with the correct root.
      If String.IsNullOrEmpty(xml) Then
        doc = New XDocument(New XElement(XName.Get(RootElementName, XmlNamespace)))
      Else
        doc = XDocument.Parse(xml)
        If doc.Root Is Nothing OrElse doc.Root.Name.LocalName <> RootElementName Then
          doc = New XDocument(New XElement(XName.Get(RootElementName, XmlNamespace)))
        End If
      End If

      Return doc
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "{RoutineName}")
      Return Nothing
    Finally
      ' --- Cleanup ---
      doc = Nothing
    End Try

  End Function

  ' ==========================================================================================
  ' Routine: SaveXmlDocument
  ' Purpose: Persist the updated XDocument back into the workbook's Custom XML Parts.
  ' Parameters:
  '   wb   - Excel.Workbook containing the XML parts.
  '   part - The existing Custom XML Part to be replaced.
  '   doc  - XDocument containing the updated XML.
  ' Returns:
  '   None.
  ' Notes:
  '   Deletes the old part and adds a new one with the updated XML content. This ensures the
  '   XML part remains consistent and avoids partial updates.
  ' ==========================================================================================
  Private Sub SaveXmlDocument(wb As Excel.Workbook, part As Object, doc As XDocument)

    Dim xml As String
    Try
      xml = doc.ToString()
      part.Delete()
      wb.CustomXMLParts.Add(xml)
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "{RoutineName}")
    Finally
      ' --- Cleanup ---
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: DuplicateGuidXmlEntry
  ' Purpose:
  '   When a cell identity is copied, the new GUID must be added to the same RuleRegion as the
  '   original GUID. Also updates the reverse map (RM_newGuid → RR_regionId).
  ' Parameters:
  '   wb      - Excel.Workbook containing the XML parts.
  '   oldName - Full identity name of the source cell (RM_<guid>).
  '   newName - Full identity name of the copied cell (RM_<guid>).
  ' Returns:
  '   None.
  ' Notes:
  '   - Preserves RuleRegion membership across copy operations.
  ' ==========================================================================================
  Private Sub DuplicateGuidXmlEntry(wb As Excel.Workbook, oldName As String, newName As String)

    Dim part As Object = Nothing
    Dim doc As XDocument = Nothing
    Dim oldGuid As String = Nothing
    Dim newGuid As String = Nothing
    Dim regionId As String = Nothing
    Dim rr As XElement = Nothing

    Try
      ' --- Normal execution ---
      doc = LoadXmlDocument(wb, part)

      oldGuid = oldName.Substring(NamePrefix.Length, 32)
      newGuid = newName.Substring(NamePrefix.Length, 32)

      ' Use reverse map: RM_oldGuid → RR_regionId
      regionId = GetRegionIdForCellName(wb, oldName)
      If String.IsNullOrEmpty(regionId) Then
        SaveXmlDocument(wb, part, doc)
        Exit Sub
      End If

      ' Find the region by ID
      rr = FindRuleRegionById(doc, regionId)
      If rr Is Nothing Then
        SaveXmlDocument(wb, part, doc)
        Exit Sub
      End If

      ' Add new GUID to region
      AddGuidToRuleRegion(rr, newGuid)

      ' Update reverse map for new identity
      AddReverseMapEntry(wb, newName, regionId)

      ' Save XML
      SaveXmlDocument(wb, part, doc)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "DuplicateGuidXmlEntry")

    Finally
      ' --- Cleanup ---
      part = Nothing
      doc = Nothing
      rr = Nothing

    End Try

  End Sub
#End Region

#Region "RuleRegion Helpers"
  ' ==========================================================================================
  ' Routine: FindRuleRegionById
  ' Purpose: Find an existing <RuleRegion> element by its ID.
  ' Parameters:
  '   doc      - XDocument representing the ResourceManagement XML.
  '   regionId - String RuleRegion ID to search for.
  ' Returns:
  '   XElement - The matching <RuleRegion> element if found; otherwise Nothing.
  ' Notes:
  '   - This is the primary lookup mechanism once the ID is known.
  ' ==========================================================================================
  Private Function FindRuleRegionById(doc As XDocument, regionId As String) As XElement

    Dim rr As XElement
    Try
      For Each rr In doc.Root.Elements(XName.Get(RuleRegionElementName, XmlNamespace))
        If CStr(rr.Attribute("id")) = regionId Then
          Return rr
        End If
      Next

      Return Nothing
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "FindRuleRegionById")
      Return Nothing

    Finally
      rr = Nothing

    End Try

  End Function

  ' ==========================================================================================
  ' Routine: CreateRuleRegion
  ' Purpose:
  '   Create a new RuleRegion element with a stable GUID-based ID and initial rule metadata.
  ' Parameters:
  '   doc            - XDocument containing the ResourceManagement root.
  '   regionId       - Legacy parameter (ignored; RegionID is derived from a new GUID).
  '   regionName     - Logical name of the region (for UI purposes).
  '   ruleId         - Rule id.
  '   listSelectType - List selection type.
  '   parameters     - Ordered parameter list.
  ' Returns:
  '   XElement - The newly created RuleRegion.
  ' Notes:
  '   - RegionID is RuleRegionIdPrefix + GUID (e.g., "RR_8c87...").
  ' ==========================================================================================
  Private Function CreateRuleRegion(doc As XDocument,
                                  regionId As String,
                                  regionName As String,
                                  ruleId As String,
                                  listSelectType As String,
                                  parameters As IList(Of RuleParameter)) As XElement

    Dim rr As XElement = Nothing
    Dim guid As String = Nothing
    Dim fullRegionId As String = Nothing
    Dim ruleElem As XElement = Nothing
    Dim paramsElem As XElement = Nothing
    Dim cellsElem As XElement = Nothing

    Try
      ' --- Normal execution ---
      guid = System.Guid.NewGuid().ToString("N")
      fullRegionId = RuleRegionIdPrefix & guid   ' e.g. "RR_8c87..."

      ' Build <Parameters>
      paramsElem = New XElement(XName.Get(ParametersElementName, XmlNamespace))
      If parameters IsNot Nothing Then
        For Each p In parameters
          Dim pElem As New XElement(XName.Get(ParameterElementName, XmlNamespace),
                                  New XAttribute("name", p.Name),
                                  New XAttribute("refType", p.RefType),
                                  New XAttribute("refValue", p.RefValue),
                                  New XAttribute("literalValue", p.LiteralValue))
          paramsElem.Add(pElem)
        Next
      End If

      ' Build <Rule>
      ruleElem = New XElement(XName.Get(RuleElementName, XmlNamespace),
                            New XAttribute("ruleId", ruleId),
                            New XAttribute("listSelectType", listSelectType),
                            paramsElem)

      ' Build <Cells>
      cellsElem = New XElement(XName.Get(CellsElementName, XmlNamespace))

      ' Build <RuleRegion>
      rr = New XElement(XName.Get(RuleRegionElementName, XmlNamespace),
                      New XAttribute("id", fullRegionId),
                      New XAttribute("name", regionName),
                      New XAttribute("state", RuleRegionStateValid),
                      ruleElem,
                      cellsElem)

      doc.Root.Add(rr)
      Return rr

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "CreateRuleRegion")
      Return Nothing

    Finally
      ' --- Cleanup ---
      rr = Nothing
      guid = Nothing
      fullRegionId = Nothing
      ruleElem = Nothing
      paramsElem = Nothing
      cellsElem = Nothing

    End Try

  End Function

  ' ==========================================================================================
  ' Routine: AddGuidToRuleRegion
  ' Purpose: Ensure the specified GUID is listed under the <Cells> element of a RuleRegion.
  ' Parameters:
  '   rr   - XElement representing the <RuleRegion>.
  '   guid - String GUID to add.
  ' Returns:
  '   None.
  ' Notes:
  '   - Does nothing if the GUID is already present.
  ' ==========================================================================================
  Private Sub AddGuidToRuleRegion(rr As XElement, guid As String)

    Dim cellsElem As XElement
    Dim cellElem As XElement
    Dim existing As Boolean
    Try
      ' Ensure the <Cells> container exists.
      cellsElem = rr.Element(XName.Get(CellsElementName, XmlNamespace))
      If cellsElem Is Nothing Then
        cellsElem = New XElement(XName.Get(CellsElementName, XmlNamespace))
        rr.Add(cellsElem)
      End If

      ' Check if the GUID is already present.
      existing = False
      For Each cellElem In cellsElem.Elements(XName.Get(CellElementName, XmlNamespace))
        If CStr(cellElem.Attribute("guid")) = guid Then
          existing = True
          Exit For
        End If
      Next

      ' Add a new <Cell> element if the GUID is not already listed.
      If Not existing Then
        cellElem = New XElement(XName.Get(CellElementName, XmlNamespace),
                              New XAttribute("guid", guid))
        cellsElem.Add(cellElem)
      End If
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "AddGuidToRuleRegion")

    Finally
      cellsElem = Nothing
      cellElem = Nothing

    End Try

  End Sub

  ' ==========================================================================================
  ' Routine: RemoveGuidFromRuleRegion
  ' Purpose: Remove the specified GUID from the <Cells> element of a RuleRegion.
  ' Parameters:
  '   rr   - XElement representing the <RuleRegion>.
  '   guid - String GUID to remove.
  ' Returns:
  '   Boolean - True if a GUID was removed; False if not found.
  ' Notes:
  '   - Does not remove the RuleRegion itself; caller is responsible for cleanup.
  ' ==========================================================================================
  Private Function RemoveGuidFromRuleRegion(rr As XElement, guid As String) As Boolean

    Dim cellsElem As XElement = Nothing
    Dim cellElem As XElement = Nothing
    Dim removed As Boolean = False

    Try
      cellsElem = rr.Element(XName.Get(CellsElementName, XmlNamespace))
      If cellsElem Is Nothing Then
        Return False
      End If

      For Each cellElem In cellsElem.Elements(XName.Get(CellElementName, XmlNamespace)).ToList()
        If CStr(cellElem.Attribute("guid")) = guid Then
          cellElem.Remove()
          removed = True
        End If
      Next

      Return removed

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "RemoveGuidFromRuleRegion")
      Return False

    Finally
      cellsElem = Nothing
      cellElem = Nothing

    End Try

  End Function

#End Region

#Region "Identity Helpers"
  ' ==========================================================================================
  ' Routine: AddGuidToCustomXml
  ' Purpose:
  '   Restore a cell identity (RM_<guid>) into BOTH:
  '     - ReverseMap (RM_<guid> → RR_<...>)
  '     - RuleRegion <Cells><Cell guid="..."/>
  '
  '   This is the logical inverse of RemoveGuidFromCustomXml:
  '     - If the GUID already belongs to a RuleRegion, it is reinserted there.
  '     - If the GUID exists in ReverseMap, its region is reused.
  '     - If the GUID has no region yet, only the ReverseMap entry is created
  '       with region="". Rule assignment will occur later via SetCellRule.
  '
  ' Parameters:
  '   wb       - Excel.Workbook containing the CustomXMLPart.
  '   nameText - Full identity name (e.g., "RM_<guid>").
  '   target   - Excel.Range the identity name refers to.
  '
  ' Notes:
  '   - Idempotent: calling it multiple times is safe.
  '   - Never invents a RuleRegion; only restores existing ones.
  ' ==========================================================================================
  Friend Sub AddGuidToCustomXml(wb As Excel.Workbook, nameText As String, target As Excel.Range)

    Dim part As Object = Nothing
    Dim doc As XDocument = Nothing
    Dim guid As String = String.Empty
    Dim regionId As String = String.Empty
    Dim rr As XElement = Nothing

    Try
      ' --- Normal execution ---

      ' Extract bare GUID from RM_<guid>
      If String.IsNullOrEmpty(nameText) OrElse Not nameText.StartsWith(NamePrefix, StringComparison.Ordinal) Then Exit Sub
      guid = nameText.Substring(NamePrefix.Length)

      ' Load main XML (ResourceManagement + RuleRegions)
      doc = LoadXmlDocument(wb, part)
      If doc Is Nothing OrElse part Is Nothing Then Exit Sub
      ' ---------------------------------------------------------
      ' 1. Try to find an existing RuleRegion containing this GUID
      ' ---------------------------------------------------------
      Dim cellElem As XElement =
            doc.
            Descendants(XName.Get(CellElementName, XmlNamespace)).
            FirstOrDefault(Function(c)
                             Dim a = c.Attribute("guid")
                             Return a IsNot Nothing AndAlso String.Equals(CStr(a.Value), guid, StringComparison.Ordinal)
                           End Function)

      If cellElem IsNot Nothing Then
        rr = cellElem.Ancestors(XName.Get(RuleRegionElementName, XmlNamespace)).FirstOrDefault()
        If rr IsNot Nothing Then
          Dim idAttr = rr.Attribute("id")
          If idAttr IsNot Nothing Then
            regionId = CStr(idAttr.Value)
          End If
        End If
      End If
      ' ---------------------------------------------------------
      ' 2. If no RuleRegion found, check ReverseMap
      ' ---------------------------------------------------------
      If String.IsNullOrEmpty(regionId) Then
        regionId = GetRegionIdForCellName(wb, nameText)
      End If
      ' ---------------------------------------------------------
      ' 3. Ensure ReverseMap entry exists (may have empty region)
      ' ---------------------------------------------------------
      AddReverseMapEntry(wb, nameText, regionId)
      ' ---------------------------------------------------------
      ' 4. If region known → ensure RuleRegion contains the GUID
      ' ---------------------------------------------------------
      If Not String.IsNullOrEmpty(regionId) Then
        If rr Is Nothing Then
          rr = FindRuleRegionById(doc, regionId)
        End If

        If rr IsNot Nothing Then
          AddGuidToRuleRegion(rr, guid)
        End If
      End If
      ' ---------------------------------------------------------
      ' 5. Save XML
      ' ---------------------------------------------------------
      SaveXmlDocument(wb, part, doc)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "AddGuidToCustomXml")

    Finally
      part = Nothing
      doc = Nothing
      rr = Nothing
    End Try

  End Sub

  ' ==========================================================================================
  ' Routine: RemoveGuidFromCustomXml
  ' Purpose:
  '   Remove any CustomXMLParts entry associated with a given identity name.
  '
  ' Parameters:
  '   wb       - Excel.Workbook containing the XML parts.
  '   nameText - Full workbook-level identity name (e.g., "RM_<GUID>").
  '
  ' Notes:
  '   - Uses simple substring matching; replace with schema-based lookup if needed.
  ' ==========================================================================================
  Private Sub RemoveGuidFromCustomXml(wb As Excel.Workbook, nameText As String)

    Dim part As Object = Nothing
    Dim doc As XDocument = Nothing
    Dim regionId As String = Nothing
    Dim rr As XElement = Nothing
    Dim cellsElem As XElement = Nothing
    Dim guid As String = Nothing

    Try
      ' --- Normal execution ---
      ' nameText is "RM_<guid>"
      guid = nameText.Substring(NamePrefix.Length, 32)

      ' Lookup region via reverse map
      regionId = GetRegionIdForCellName(wb, nameText)
      If String.IsNullOrEmpty(regionId) Then Exit Sub

      ' Load main ResourceManagement XML
      doc = LoadXmlDocument(wb, part)
      If doc Is Nothing OrElse doc.Root Is Nothing Then Exit Sub

      ' Find the region
      rr = FindRuleRegionById(doc, regionId)
      If rr Is Nothing Then
        SaveXmlDocument(wb, part, doc)
        Exit Sub
      End If

      ' Remove GUID from region
      If Not RemoveGuidFromRuleRegion(rr, guid) Then
        SaveXmlDocument(wb, part, doc)
        Exit Sub
      End If

      ' Delete region if empty
      cellsElem = rr.Element(XName.Get(CellsElementName, XmlNamespace))
      If cellsElem Is Nothing OrElse
       Not cellsElem.Elements(XName.Get(CellElementName, XmlNamespace)).Any() Then
        rr.Remove()
      End If

      ' Remove reverse map entry
      RemoveReverseMapEntry(wb, nameText)

      ' Save updated XML
      SaveXmlDocument(wb, part, doc)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "RemoveGuidFromCustomXml")

    Finally
      part = Nothing
      doc = Nothing
      rr = Nothing
      cellsElem = Nothing
      guid = Nothing

    End Try

  End Sub

  ' ==========================================================================================
  ' Routine:    SetRangeIdValue
  ' Purpose:    Safely assign the Excel Range.ID property using late binding.
  '
  ' Contract:
  '   - cell:      Excel.Range whose ID property should be set.
  '   - idValue:   Workbook-level name (e.g. "RM_<guid>") to assign.
  '
  ' Behaviour:
  '   - Uses InvokeMember to set the COM property "ID".
  '   - Silently ignores Excel versions that do not support Range.ID.
  '   - Never throws to the caller; all errors are handled internally.
  '   - Friend as called from other modules in the assembly.
  ' ==========================================================================================
  Friend Sub SetRangeIdValue(target As Excel.Range, idValue As String)

    Try
      target.GetType().InvokeMember("ID",
                                    Reflection.BindingFlags.SetProperty,
                                    Nothing,
                                    target,
                                    New Object() {idValue})
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "SetRangeIdValue")
    End Try

  End Sub
#End Region

#Region "Reverse MAP XML"
  ' ==========================================================================================
  ' Routine: GetOrCreateReverseMapPart
  ' Purpose: Retrieve the Custom XML Part used for the reverse cell→region map, or create it.
  ' Parameters:
  '   wb - Excel.Workbook containing the XML parts.
  ' Returns:
  '   Object - The XML part COM object whose root element is <ReverseMap>.
  ' Notes:
  '   Ensures a single authoritative XML part exists for reverse lookup metadata.
  ' ==========================================================================================
  Private Function GetOrCreateReverseMapPart(wb As Excel.Workbook) As Object

    Dim part As Object = Nothing
    Dim xml As String = Nothing
    Dim doc As XDocument = Nothing

    Try
      ' --- Normal execution ---
      For Each part In wb.CustomXMLParts
        xml = CStr(part.XML)
        If Not String.IsNullOrEmpty(xml) Then
          doc = XDocument.Parse(xml)
          If doc.Root IsNot Nothing AndAlso
           doc.Root.Name.LocalName = ReverseMapRootElementName Then
            Return part
          End If
        End If
      Next

      xml = "<" & ReverseMapRootElementName & " xmlns='" & XmlNamespace & "'></" & ReverseMapRootElementName & ">"
      Return wb.CustomXMLParts.Add(xml)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "GetOrCreateReverseMapPart")
      Return Nothing

    Finally
      ' --- Cleanup ---
      doc = Nothing
      xml = Nothing
      part = Nothing

    End Try

  End Function

  ' ==========================================================================================
  ' Routine: LoadReverseMap
  ' Purpose: Load the XDocument for the reverse cell→region map XML part.
  ' Parameters:
  '   wb   - Excel.Workbook containing the XML parts.
  '   part - [ByRef] The Custom XML Part COM object returned.
  ' Returns:
  '   XDocument - Parsed XML document for the reverse map part.
  ' Notes:
  '   Always returns a valid document with a <ReverseMap> root.
  ' ==========================================================================================
  Private Function LoadReverseMap(wb As Excel.Workbook, ByRef part As Object) As XDocument

    Dim xml As String = Nothing
    Dim doc As XDocument = Nothing

    Try
      ' --- Normal execution ---
      part = GetOrCreateReverseMapPart(wb)
      If part Is Nothing Then
        doc = New XDocument(New XElement(XName.Get(ReverseMapRootElementName, XmlNamespace)))
        Return doc
      End If

      xml = CStr(part.XML)

      If String.IsNullOrEmpty(xml) Then
        doc = New XDocument(New XElement(XName.Get(ReverseMapRootElementName, XmlNamespace)))
      Else
        doc = XDocument.Parse(xml)
        If doc.Root Is Nothing OrElse doc.Root.Name.LocalName <> ReverseMapRootElementName Then
          doc = New XDocument(New XElement(XName.Get(ReverseMapRootElementName, XmlNamespace)))
        End If
      End If

      Return doc

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "LoadReverseMap")
      Return New XDocument(New XElement(XName.Get(ReverseMapRootElementName, XmlNamespace)))

    Finally
      ' --- Cleanup ---
      xml = Nothing

    End Try

  End Function

  ' ==========================================================================================
  ' Routine: SaveReverseMap
  ' Purpose: Persist the updated reverse map XDocument back into the workbook's Custom XML Parts.
  ' Parameters:
  '   wb   - Excel.Workbook containing the XML parts.
  '   part - The existing Custom XML Part to be replaced.
  '   doc  - XDocument containing the updated reverse map XML.
  ' Returns:
  '   None.
  ' Notes:
  '   Deletes the old part and adds a new one with the updated XML content.
  ' ==========================================================================================
  Private Sub SaveReverseMap(wb As Excel.Workbook, part As Object, doc As XDocument)

    Dim xml As String = Nothing

    Try
      ' --- Normal execution ---
      If part IsNot Nothing Then
        xml = doc.ToString()
        part.Delete()
        wb.CustomXMLParts.Add(xml)
      End If

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "SaveReverseMap")

    Finally
      ' --- Cleanup ---
      xml = Nothing
      part = Nothing
      doc = Nothing

    End Try

  End Sub

  ' ==========================================================================================
  ' Routine: AddReverseMapEntry
  ' Purpose: Add or replace a reverse map entry mapping a cell identity name to a region ID.
  ' Parameters:
  '   wb       - Excel.Workbook containing the XML parts.
  '   cellName - String full identity name (e.g., "RM_<guid>").
  '   regionId - String region ID (e.g., "RR_<guid>").
  ' Returns:
  '   None.
  ' Notes:
  '   Any existing entry for the same cellName is removed before adding the new one.
  ' ==========================================================================================
  Private Sub AddReverseMapEntry(wb As Excel.Workbook, cellName As String, regionId As String)

    Dim part As Object = Nothing
    Dim doc As XDocument = Nothing

    Try
      ' --- Normal execution ---
      doc = LoadReverseMap(wb, part)
      Dim root = doc.Root
      Dim entryName = XName.Get(ReverseMapEntryElementName, XmlNamespace)

      Dim existing = root.Elements(entryName).
        FirstOrDefault(Function(e) String.Equals(CStr(e.Attribute(ReverseMapCellAttrName)),
                                                 cellName,
                                                 StringComparison.Ordinal))

      If existing IsNot Nothing Then existing.Remove()

      Dim newEntry As New XElement(entryName,
        New XAttribute(ReverseMapCellAttrName, cellName),
        New XAttribute(ReverseMapRegionAttrName, regionId))

      root.Add(newEntry)
      SaveReverseMap(wb, part, doc)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "AddReverseMapEntry")

    Finally
      ' --- Cleanup ---
      doc = Nothing
      part = Nothing

    End Try

  End Sub

  ' ==========================================================================================
  ' Routine: RemoveReverseMapEntry
  ' Purpose: Remove a reverse map entry for a given cell identity name, if present.
  ' Parameters:
  '   wb       - Excel.Workbook containing the XML parts.
  '   cellName - String full identity name (e.g., "RM_<guid>").
  ' Returns:
  '   None.
  ' Notes:
  '   Silent if no entry exists.
  ' ==========================================================================================
  Private Sub RemoveReverseMapEntry(wb As Excel.Workbook, cellName As String)

    Dim part As Object = Nothing
    Dim doc As XDocument = Nothing

    Try
      ' --- Normal execution ---
      doc = LoadReverseMap(wb, part)
      Dim root = doc.Root
      Dim entryName = XName.Get(ReverseMapEntryElementName, XmlNamespace)

      Dim existing = root.Elements(entryName).
        FirstOrDefault(Function(e) String.Equals(CStr(e.Attribute(ReverseMapCellAttrName)),
                                                 cellName,
                                                 StringComparison.Ordinal))

      If existing IsNot Nothing Then
        existing.Remove()
        SaveReverseMap(wb, part, doc)
      End If

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "RemoveReverseMapEntry")

    Finally
      ' --- Cleanup ---
      doc = Nothing
      part = Nothing

    End Try

  End Sub

  ' ==========================================================================================
  ' Routine: GetRegionIdForCellName
  ' Purpose: Retrieve the region ID associated with a given cell identity name.
  ' Parameters:
  '   wb       - Excel.Workbook containing the XML parts.
  '   cellName - String full identity name (e.g., "RM_<guid>").
  ' Returns:
  '   String - Region ID (e.g., "RR_<guid>") if found; otherwise empty string.
  ' Notes:
  '   Uses the reverse map XML part for O(1)-style lookup without scanning RuleRegions.
  ' ==========================================================================================
  Private Function GetRegionIdForCellName(wb As Excel.Workbook, cellName As String) As String

    Dim part As Object = Nothing
    Dim doc As XDocument = Nothing

    Try
      ' --- Normal execution ---
      doc = LoadReverseMap(wb, part)

      Dim root = doc.Root
      Dim entryName = XName.Get(ReverseMapEntryElementName, XmlNamespace)

      Dim existing = root.Elements(entryName).
        FirstOrDefault(Function(e) String.Equals(CStr(e.Attribute(ReverseMapCellAttrName)),
                                                 cellName,
                                                 StringComparison.Ordinal))

      If existing Is Nothing Then Return String.Empty
      Return CStr(existing.Attribute(ReverseMapRegionAttrName))

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "GetRegionIdForCellName")
      Return String.Empty

    Finally
      ' --- Cleanup ---
      doc = Nothing
      part = Nothing

    End Try

  End Function
#End Region

#Region "Apply Routines"
  ' ==========================================================================================
  ' Routine: GetAllApplyInstances
  ' Purpose:
  '   Load all <RuleRegion> elements from the ResourceManagement XML and convert them into
  '   storage-layer ExcelApplyInstance objects for the UI loader.
  '
  ' Parameters:
  '   wb - Excel.Workbook containing the XML parts.
  '
  ' Returns:
  '   List(Of ExcelApplyInstance) - One instance per RuleRegion.
  '
  ' Notes:
  '   - UILoaderSaver must never see XML; this routine hides all XML details.
  '   - RuleRegion.id becomes ApplyID.
  '   - RuleRegion.name becomes ApplyName.
  '   - <Rule ruleId="..."> is the RuleID (raw GUID).
  '   - listSelectType is stored on the <Rule> element.
  '   - Parameters are mapped to RuleParameter objects.
  '   - CellGuids contains all bare GUIDs under <Cells>.
  ' ==========================================================================================
  Friend Function GetAllApplyInstances(wb As Excel.Workbook) As List(Of ExcelApplyInstance)

    Dim list As New List(Of ExcelApplyInstance)
    Dim part As Object = Nothing
    Dim doc As XDocument = Nothing
    Dim rr As XElement = Nothing

    Try
      ' --- Normal execution ---
      doc = LoadXmlDocument(wb, part)
      If doc Is Nothing OrElse doc.Root Is Nothing Then Return list

      ' Enumerate all RuleRegions
      For Each rr In doc.Root.Elements(XName.Get(RuleRegionElementName, XmlNamespace))

        Dim inst As New ExcelApplyInstance
        inst.ApplyID = CStr(rr.Attribute("id"))
        inst.ApplyName = CStr(rr.Attribute("name"))
        inst.Parameters = New List(Of RuleParameter)
        inst.CellGuids = New List(Of String)

        ' Extract <Rule>
        Dim ruleElem As XElement = rr.Element(XName.Get(RuleElementName, XmlNamespace))
        If ruleElem IsNot Nothing Then
          inst.RuleID = CStr(ruleElem.Attribute("ruleId"))
          inst.ListSelectType = CStr(ruleElem.Attribute("listSelectType"))

          ' Extract parameters
          Dim paramsElem As XElement = ruleElem.Element(XName.Get(ParametersElementName, XmlNamespace))
          If paramsElem IsNot Nothing Then
            For Each pElem In paramsElem.Elements(XName.Get(ParameterElementName, XmlNamespace))
              Dim p As New RuleParameter With {
              .Name = CStr(pElem.Attribute("name")),
              .RefType = CStr(pElem.Attribute("refType")),
              .RefValue = CStr(pElem.Attribute("refValue")),
              .LiteralValue = CStr(pElem.Attribute("literalValue"))
            }
              inst.Parameters.Add(p)
            Next
          End If
        End If

        ' Extract cell GUIDs
        Dim cellsElem As XElement = rr.Element(XName.Get(CellsElementName, XmlNamespace))
        If cellsElem IsNot Nothing Then
          For Each cElem In cellsElem.Elements(XName.Get(CellElementName, XmlNamespace))
            inst.CellGuids.Add(CStr(cElem.Attribute("guid")))
          Next
        End If

        list.Add(inst)
      Next

      Return list

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return list

    Finally
      ' --- Cleanup ---
      part = Nothing
      doc = Nothing
      rr = Nothing

    End Try

  End Function

  ' ==========================================================================================
  ' Routine: SaveApplyInstance
  '
  ' Purpose:
  '   Persist changes to a RuleRegion (Apply instance) by:
  '     - Creating the region when new.
  '     - Updating ruleId, listSelectType, and parameters when dirty.
  '     - Deleting the region, its membership, reverse-map entries, and (unused) GUID names
  '       when marked as deleted.
  '     - Additively attaching the selected cells to the region, moving them from any
  '       other region if necessary.
  '
  ' Parameters:
  '   wb       - Excel.Workbook containing the XML parts and defined names.
  '   instance - ExcelApplyInstance describing the Apply operation:
  '                - ApplyID
  '                - RuleID
  '                - ListSelectType
  '                - Parameters
  '                - IsNew / IsDirty / IsDeleted flags
  '   target   - Excel.Range representing the user-selected cells to attach to this region
  '              (ignored when deleting).
  '
  ' Behaviour:
  '   - NEW:
  '       - Creates a new <RuleRegion> with the supplied RuleID, ListSelectType, and Parameters.
  '       - Assigns the generated region id back to instance.ApplyID.
  '       - Additively attaches all cells in target to this new region.
  '
  '   - DIRTY (update):
  '       - Ensures the <RuleRegion> exists (creates if missing).
  '       - Updates RuleID, ListSelectType, and Parameters on the region.
  '       - Additively attaches all cells in target:
  '           * If a cell is already in this region → no-op.
  '           * If a cell is attached to another region → removes it from that region and
  '             its reverse-map entry, then attaches it to this region.
  '
  '   - DELETED:
  '       - Locates the <RuleRegion> by ApplyID.
  '       - For each <Cell guid="..."> in that region:
  '           * Removes the reverse-map entry for the corresponding GUID name.
  '           * Deletes the GUID defined name if it is no longer referenced by any region.
  '           * Removes the <Cell> entry.
  '       - Removes the <RuleRegion> element itself.
  '
  ' Notes:
  '   - Cell identity is managed via EnsureCellHasGuidName, which guarantees a hidden
  '     workbook-level defined name "__RM_{GUID}" referring to the cell.
  '   - Reverse-map helpers (GetRegionIdForCellName / AddReverseMapEntry / RemoveReverseMapEntry)
  '     are used to enforce the invariant that a cell can belong to at most one region.
  '   - Membership is additive on update: existing cells in the region are preserved unless
  '     explicitly moved to another region.
  ' ==========================================================================================
  Friend Sub SaveApplyInstance(wb As Excel.Workbook,
                             instance As ExcelApplyInstance,
                             target As Excel.Range)

    Const RoutineName As String = "SaveApplyInstance"

    Dim part As Object = Nothing
    Dim doc As XDocument = Nothing
    Dim rr As XElement = Nothing
    Dim cellsElem As XElement = Nothing

    Try
      If wb Is Nothing Then Exit Sub
      If instance Is Nothing Then Exit Sub
      If target Is Nothing Then Exit Sub

      ' --- Load XML once ---
      doc = LoadXmlDocument(wb, part)
      If doc Is Nothing OrElse doc.Root Is Nothing Then Exit Sub

      ' =========================================================
      ' DELETE: remove region + reverse map + membership + names
      ' =========================================================
      If instance.IsDeleted Then

        rr = FindRuleRegionById(doc, instance.ApplyID)
        If rr IsNot Nothing Then

          cellsElem = rr.Element(XName.Get(CellsElementName, XmlNamespace))
          If cellsElem IsNot Nothing Then

            For Each cElem In cellsElem.Elements(XName.Get(CellElementName, XmlNamespace)).ToList()

              Dim guid As String = CStr(cElem.Attribute("guid"))
              Dim cellName As String = NamePrefix & guid

              ' Remove reverse map entry
              RemoveReverseMapEntry(wb, cellName)

              ' Delete defined name ONLY if unused elsewhere
              DeleteGuidNameIfUnused(wb, doc, guid)

              ' Remove membership entry
              cElem.Remove()

            Next

          End If

          ' Remove the region itself
          rr.Remove()

        End If

        SaveXmlDocument(wb, part, doc)
        Exit Sub

      End If

      ' =========================================================
      ' ADD / UPDATE: ensure region exists and metadata is correct
      ' =========================================================
      If instance.IsNew Then
        ' Create new RuleRegion
        rr = CreateRuleRegion(doc,
                            Nothing,
                            instance.ApplyName,
                            instance.RuleID,
                            instance.ListSelectType,
                            instance.Parameters)
        instance.ApplyID = CStr(rr.Attribute("id"))

      ElseIf instance.IsDirty Then
        ' Find existing region; if missing, create it
        rr = FindRuleRegionById(doc, instance.ApplyID)
        If rr Is Nothing Then
          rr = CreateRuleRegion(doc,
                              Nothing,
                              instance.ApplyName,
                              instance.RuleID,
                              instance.ListSelectType,
                              instance.Parameters)
          instance.ApplyID = CStr(rr.Attribute("id"))
        Else
          ' --- Update <RuleRegion> name ---
          rr.SetAttributeValue("name", instance.ApplyName)
          ' --- Update <Rule> metadata + parameters ---
          Dim ruleElem As XElement = rr.Element(XName.Get(RuleElementName, XmlNamespace))
          If ruleElem IsNot Nothing Then
            ruleElem.SetAttributeValue("ruleId", instance.RuleID)
            ruleElem.SetAttributeValue("listSelectType", instance.ListSelectType)

            Dim paramsElem As XElement = ruleElem.Element(XName.Get(ParametersElementName, XmlNamespace))
            If paramsElem IsNot Nothing Then paramsElem.RemoveAll()

            For Each p In instance.Parameters
              Dim pElem As New XElement(XName.Get(ParameterElementName, XmlNamespace),
                                      New XAttribute("name", p.Name),
                                      New XAttribute("refType", p.RefType),
                                      New XAttribute("refValue", p.RefValue),
                                      New XAttribute("literalValue", p.LiteralValue))
              paramsElem.Add(pElem)
            Next
          End If
        End If
      Else
        ' Nothing to do
        Exit Sub
      End If

      ' =========================================================
      ' Membership: additive only, with move-from-other-region logic
      ' =========================================================

      ' Ensure <Cells> element exists
      cellsElem = rr.Element(XName.Get(CellsElementName, XmlNamespace))
      If cellsElem Is Nothing Then
        cellsElem = New XElement(XName.Get(CellsElementName, XmlNamespace))
        rr.Add(cellsElem)
      End If

      ' Build a HashSet of existing GUIDs for fast lookup
      Dim existingGuids As New HashSet(Of String)(
        cellsElem.Elements(XName.Get(CellElementName, XmlNamespace)).
                  Select(Function(x) CStr(x.Attribute("guid")))
      )

      ' For each cell in the selected range
      For Each area As Excel.Range In target.Areas
        For Each cell As Excel.Range In area.Cells

          ' Ensure GUID + defined name
          Dim guid As String = EnsureCellHasGuidName(cell)
          Dim cellName As String = NamePrefix & guid

          ' If already in THIS region → skip
          If existingGuids.Contains(guid) Then Continue For

          ' Check if cell belongs to ANOTHER region
          Dim oldRegionId As String = GetRegionIdForCellName(wb, cellName)
          If Not String.IsNullOrEmpty(oldRegionId) AndAlso oldRegionId <> instance.ApplyID Then

            ' Remove from old region
            Dim oldRR As XElement = FindRuleRegionById(doc, oldRegionId)
            If oldRR IsNot Nothing Then
              Dim oldCells = oldRR.Element(XName.Get(CellsElementName, XmlNamespace))
              If oldCells IsNot Nothing Then
                Dim oldCellElem = oldCells.Elements(XName.Get(CellElementName, XmlNamespace)).
                                     FirstOrDefault(Function(x) CStr(x.Attribute("guid")) = guid)
                If oldCellElem IsNot Nothing Then oldCellElem.Remove()
              End If
            End If

            ' Remove old reverse map
            RemoveReverseMapEntry(wb, cellName)
          End If

          ' Add to THIS region
          AddGuidToRuleRegion(rr, guid)
          AddReverseMapEntry(wb, cellName, instance.ApplyID)
          existingGuids.Add(guid)

        Next
      Next


      '' =========================================================
      '' Membership: disconnected ranges → clear then re‑add
      '' =========================================================
      '' Clear existing membership for this region
      'cellsElem = rr.Element(XName.Get(CellsElementName, XmlNamespace))
      'If cellsElem Is Nothing Then
      '  cellsElem = New XElement(XName.Get(CellsElementName, XmlNamespace))
      '  rr.Add(cellsElem)
      'Else
      '  For Each cElem In cellsElem.Elements(XName.Get(CellElementName, XmlNamespace)).ToList()
      '    Dim guid As String = CStr(cElem.Attribute("guid"))
      '    Dim cellName As String = NamePrefix & guid
      '    RemoveReverseMapEntry(wb, cellName)
      '    cElem.Remove()
      '  Next
      'End If

      '' Re‑add membership from the (possibly disconnected) target range
      'For Each area As Excel.Range In target.Areas
      '  For Each cell As Excel.Range In area.Cells

      '    ' Ensure GUID identity + defined name
      '    Dim guid As String = EnsureCellHasGuidName(cell)
      '    Dim cellName As String = NamePrefix & guid

      '    ' Add GUID to region
      '    AddGuidToRuleRegion(rr, guid)

      '    ' Update reverse map: RM_<guid> → this ApplyID
      '    AddReverseMapEntry(wb, cellName, instance.ApplyID)

      '  Next
      'Next

      ' Persist XML
      SaveXmlDocument(wb, part, doc)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, RoutineName)

    Finally
      part = Nothing
      doc = Nothing
      rr = Nothing
      cellsElem = Nothing

    End Try

  End Sub

  ' ==========================================================================================
  ' Routine: DeleteGuidNameIfUnused
  ' Purpose:
  '   Delete the defined name for a GUID ONLY if no other RuleRegion references it.
  ' ==========================================================================================
  Private Sub DeleteGuidNameIfUnused(wb As Excel.Workbook,
                                     doc As XDocument,
                                     guid As String)

    Dim nameText As String = NamePrefix & guid

    ' Check if GUID appears in ANY other region
    Dim stillUsed As Boolean =
      doc.Root.
         Descendants(XName.Get(CellElementName, XmlNamespace)).
         Any(Function(x) CStr(x.Attribute("guid")) = guid)

    If stillUsed Then Exit Sub

    ' Safe to delete the defined name
    Try
      Dim nm As Excel.Name = Nothing
      Try
        nm = wb.Names(nameText)
      Catch
        nm = Nothing
      End Try

      If nm IsNot Nothing Then nm.Delete()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "DeleteGuidNameIfUnused")
    End Try

  End Sub
#End Region
End Module