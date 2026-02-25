Imports System.Text.Json.Serialization
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ToolTip

Friend Enum ExcelRuleDesignerAction
  None
  Add
  Update
  Delete
End Enum
Friend Class ListTypeDescriptor
  Public Property Id As String
  Public Property Name As String
  Public Property IsSystem As Boolean
End Class

Friend Class ListItemDescriptor
  Public Property Id As String
  Public Property Name As String
End Class

Friend Class UIModelExcelRuleDesigner

  ' --- List of rules shown in the left panel ---
  Friend Property Rules As New SortableBindingList(Of UIExcelRuleDesignerRuleRow)

  ' --- The rule currently selected in the list ---
  Friend Property SelectedRule As UIExcelRuleDesignerRuleRow

  ' --- The full rule definition for editing (JSON-backed) ---
  Friend Property RuleDetail As UIExcelRuleDesignerRuleRowDetail
  Friend Property ApplyRuleDetail As UIExcelRuleDesignerRuleRowDetail

  ' --- ViewMap Operational API for (views, fields, relationships) ---
  Friend Property ViewMapHelper As ExcelRuleViewMapHelper

  ' --- Apply instances (CustomXML-backed) ---
  Friend Property ApplyInstances As New List(Of UIExcelRuleDesignerApplyInstance)

  ' --- The apply currently being edited ---
  Friend Property SelectedApply As UIExcelRuleDesignerApplyInstance

  ' --- List types and items ---
  Friend Property ListTypes As List(Of ListTypeDescriptor)
  ''Friend Property ListItemsByType As Dictionary(Of String, List(Of ListItemDescriptor))

  ' --- Lookup lists for UI controls ---
  Friend Property AvailableViews As List(Of String)
  Friend Property AvailableFields As List(Of String)
  Friend Property AvailableOperators As List(Of String)
  Friend Property AvailableOpenParentheses As List(Of String)
  Friend Property AvailableCloseParentheses As List(Of String)
  Friend Property AvailableBooleanOperators As List(Of String)

  ' --- Pending action (Add/Update/Delete) ---
  Friend Property PendingAction As ExcelRuleDesignerAction = ExcelRuleDesignerAction.None

  ' --- Action objects for commit ---
  Friend Property ActionRule As UIExcelRuleDesignerRuleRow
  Friend Property ActionRuleDetail As UIExcelRuleDesignerRuleRowDetail
  Friend Property ActionApply As UIExcelRuleDesignerApplyInstance

  Friend Sub New()
    ListTypes = New List(Of ListTypeDescriptor)
    ''ListItemsByType = New Dictionary(Of String, List(Of ListItemDescriptor))(StringComparer.OrdinalIgnoreCase)
    AvailableViews = New List(Of String)
    AvailableFields = New List(Of String)
    AvailableOperators = New List(Of String)
    AvailableOpenParentheses = New List(Of String)
    AvailableCloseParentheses = New List(Of String)
    AvailableBooleanOperators = New List(Of String)
  End Sub

End Class

Friend Class UIExcelRuleDesignerRuleRow
  Public Property RuleID As String
  Public Property RuleName As String
  Public Property RuleType As String   ' see ExcelRuleType for values
End Class

Public Class UIExcelRuleDesignerRuleRowDetail

  ' --- Identity ---
  Public Property RuleID As String
  Public Property RuleName As String
  Public Property RuleType As String

  ' --- Core definition ---
  Public Property PrimaryView As String
  Public Property SelectedValues As New List(Of UIExcelRuleDesignerRuleSelectedValue)
  Public Property Filters As New List(Of UIExcelRuleDesignerRuleFilter)
  Public Property UsedViews As New List(Of String)

End Class

' --- Needs to be Public for JSON serialization ---
Public Class UIExcelRuleDesignerRuleSelectedValue
  Public Property View As String ' Holds the canonical view name, which may differ from the actual SourceView for display purposes
  Public Property Field As String ' Holds the canonical field name, which may differ from the actual SourceField for display purposes
  Public Property SourceView As String ' Holds the view from which the field is from, which may differ from the canonical view name
  Public Property SourceField As String ' Holds the field from which the Field is from, which may differ from the canonical field name
  Public Property FieldID As String ' Internal identifier for custom fields (null for physical fields)
  Public Property ListTypeID As String
End Class

' ==========================================================================================
' Class: UIExcelRuleDesignerRuleFilter
' Purpose:
'   Represents a single filter row in the Rule Designer. This is a DESIGN-TIME ONLY model
'   used to capture the user's intent for:
'     - Which field to filter on
'     - Which operator to use
'     - Where the value comes from (literal vs parameter)
'     - How this filter participates in the overall boolean expression:
'         - BooleanOperator: AND / OR (glue to the previous filter)
'         - OpenParenCount:  number of "(" inserted before this filter
'         - CloseParenCount: number of ")" inserted after this filter
'
' Notes:
'   - This class needs to be publis as JSON-serialized as part of UIExcelRuleDesignerRuleRowDetail.
'   - No runtime-resolved values are stored here; only instructions for later evaluation.
' ==========================================================================================
Public Class UIExcelRuleDesignerRuleFilter
  Public Property FilterID As String          ' Unique identifier for this filter
  Public Property View As String ' Holds the canonical view name, which may differ from the actual SourceView for display purposes
  Public Property Field As String ' Holds the canonical field name, which may differ from the actual SourceField for display purposes
  Public Property SourceView As String ' Holds the view from which the field is from, which may differ from the canonical view name
  Public Property SourceField As String ' Holds the field from which the Field is from, which may differ from the canonical field name
  Public Property FieldID As String ' Internal identifier for user defined fields (null for physical fields)
  Public Property FieldOperator As String

  ' Glue to the PREVIOUS filter:
  '   - ""   : first filter in expression (no operator)
  '   - "AND": logical AND with previous filter
  '   - "OR" : logical OR with previous filter
  Public Property BooleanOperator As String

  ' Number of opening parentheses to insert BEFORE this filter.
  ' Example: 1 → "("; 2 → "(("; 0 → none.
  Public Property OpenParenCount As Integer

  ' Number of closing parentheses to insert AFTER this filter.
  ' Example: 1 → ")"; 2 → "))"; 0 → none.
  Public Property CloseParenCount As Integer
  Public Property ListTypeID As String
  Public Property SlicingMode As String ' Used for time based fields that support time slicing
  ' Denotes when the value for the filter is bound
  Public Property ValueBinding As String
  ' Denotes the bound value for depending on the ValueBinding mode
  Public Property LiteralValue As String

End Class

Public Class UIExcelRuleDesignerApplyInstance
  Public Property ApplyID As String          ' GUID
  Public Property ApplyName As String        ' User-visible name
  Public Property RuleID As String
  Public Property ListSelectType As String
  Public Property Parameters As New List(Of UIExcelRuleDesignerApplyParameter)
End Class

Public Class UIExcelRuleDesignerApplyParameter
  Public Property FilterID As String
  Public Property RefType As String
  Public Property RefValue As String
  Public Property LiteralValue As String
End Class
Friend Class UIExcelRuleDesignerViewMapBuilder

  Private ReadOnly _baseMap As ExcelRuleViewMap

  Public Sub New(baseMap As ExcelRuleViewMap)
    _baseMap = baseMap
  End Sub

  ' ==========================================================================================
  ' BuildUiViewMap
  '
  ' Takes the raw ViewMap (from JSON) and produces a UI-safe version:
  '   - Removes metadata views
  '   - Removes internal ID/key fields
  '   - Adds custom fields from tblResourceListItem into vwDim_Resource
  '   - Adds concept fields, i.e. fields which are calcuated runtime into vwDim_Resource
  ' ==========================================================================================
  Friend Function BuildUiViewMap(customFields As IEnumerable(Of RecordResourceListItem)) _
      As ExcelRuleViewMap

    ' --- Clone the base map shallowly so we don't mutate the original ---
    Dim uiMap As New ExcelRuleViewMap With {
      .Views = New List(Of ExcelRuleViewMapView)()
    }

    Dim customView As ExcelRuleViewMapView = Nothing
    Dim conceptView As ExcelRuleViewMapView = Nothing

    ' --- Copy only user-facing views ---
    For Each v In _baseMap.Views
      If Not IsMetadataView(v.Name) And Not v.IsConceptual Then
        uiMap.Views.Add(CloneViewWithoutInternalFields(v))
      End If
      If String.Equals(v.Name, "vwFact_ResourceCustomField", StringComparison.OrdinalIgnoreCase) Then
        CustomView = v
      End If
      If String.Equals(v.Name, "vwConcept_Availability", StringComparison.OrdinalIgnoreCase) Then
        ConceptView = v
      End If
    Next

    ' --- Merge custom fields into vwDim_Resource ---
    MergeCustomFieldsIntoResourceView(uiMap, customView, customFields)

    ' --- Merge conceptual availability fields into vwDim_Resource ---
    MergeConceptualAvailabilityFields(uiMap, conceptView)

    Return uiMap
  End Function


  ' ==========================================================================================
  ' Metadata view detection
  ' ==========================================================================================
  Private Function IsMetadataView(viewName As String) As Boolean
    Return viewName = "vwDim_CustomField" _
        OrElse viewName = "vwDim_CustomFieldValue" _
        OrElse viewName = "vwFact_ResourceCustomField"
  End Function


  ' ==========================================================================================
  ' Clone a view but remove internal fields (keys, foreign keys, lookup IDs)
  ' ==========================================================================================
  Private Function CloneViewWithoutInternalFields(source As ExcelRuleViewMapView) _
      As ExcelRuleViewMapView

    Dim cloned As New ExcelRuleViewMapView With {
      .Name = source.Name,
      .DisplayName = source.DisplayName,
      .Fields = New List(Of ExcelRuleViewMapField)(),
      .Relations = source.Relations ' shallow copy is fine
    }

    For Each f In source.Fields
      If Not IsInternalField(f) Then
        cloned.Fields.Add(New ExcelRuleViewMapField With {
          .Name = f.Name,
          .DisplayName = f.DisplayName,
          .DataType = f.DataType,
          .SourceView = source.Name,
          .SourceField = f.Name,
          .FieldID = f.FieldID,
          .Role = f.Role,
          .AllowedOperators = f.AllowedOperators?.ToList(),
          .AllowedValues = f.AllowedValues?.ToList(),
          .IsTimeBounded = f.IsTimeBounded,
          .PartnerField = f.PartnerField,
          .RequiresPartner = f.RequiresPartner,
          .SupportsSlicing = f.SupportsSlicing,
          .SlicingOptions = f.SlicingOptions?.ToList(),
          .DefaultSlicing = f.DefaultSlicing,
          .TimeBoundHandler = f.TimeBoundHandler,
          .MergedIntoView = f.MergedIntoView
        })
      End If
    Next

    Return cloned
  End Function


  ' ==========================================================================================
  ' Internal field detection
  ' ==========================================================================================
  Private Function IsInternalField(f As ExcelRuleViewMapField) As Boolean
    Return f.Role = "Key" OrElse
           f.Role = "ForeignKey" OrElse
           f.Role = "LookupType"
  End Function


  ' ==========================================================================================
  ' Merge custom fields into vwDim_Resource
  '
  ' Each RecordResourceListItem becomes a synthetic field:
  '
  ' Cardinality is NOT stored in the ViewMap; it is derived later from ValueType.
  ' ==========================================================================================
  Private Sub MergeCustomFieldsIntoResourceView(uiMap As ExcelRuleViewMap, CustomView As ExcelRuleViewMapView,
                                                customFields As IEnumerable(Of RecordResourceListItem))


    Dim resourceView = uiMap.Views.FirstOrDefault(Function(v) v.Name = "vwDim_Resource")
    If resourceView Is Nothing Then Exit Sub
    For Each cvf In CustomView.Fields

      If String.Equals(cvf.Name, "ValueName", StringComparison.OrdinalIgnoreCase) Then
        For Each cf In customFields
          If cf.IsDeleted Then Continue For

          Dim synthetic As New ExcelRuleViewMapField With {
            .Name = cf.ResourceListItemName,
            .DisplayName = cf.ResourceListItemName,
            .FieldID = cf.ResourceListItemID,
            .SourceField = "ValueName",
            .SourceView = CustomView.Name,
            .DataType = cvf.DataType,
            .Role = cvf.Role,
            .AllowedOperators = cvf.AllowedOperators?.ToList(),
            .AllowedValues = cvf.AllowedValues?.ToList(),
            .IsTimeBounded = cvf.IsTimeBounded,
            .PartnerField = cvf.PartnerField,
            .RequiresPartner = cvf.RequiresPartner,
            .SupportsSlicing = cvf.SupportsSlicing,
            .SlicingOptions = cvf.SlicingOptions?.ToList(),
            .SlicingDescriptions = If(cvf.SlicingDescriptions IsNot Nothing,
              New Dictionary(Of String, String)(cvf.SlicingDescriptions),
              Nothing),
            .DefaultSlicing = cvf.DefaultSlicing,
            .TimeBoundHandler = cvf.TimeBoundHandler,
            .MergedIntoView = cvf.MergedIntoView
          }


          '.SourceField = "ValueName", ' As we are merging the custom fields from here, this is the correct source field
          '.SourceView = "vwFact_ResourceCustomField", ' As we are merging the custom fields from here, this is the correct source view
          '.DataType = "Text",
          '.Role = "Attribute"

          ' FULL metadata copy
          resourceView.Fields.Add(synthetic)
        Next
      End If
    Next
  End Sub

  ' ==========================================================================================
  ' MergeConceptualAvailabilityFields
  '
  ' Injects the three conceptual availability fields into vwDim_Resource for UI use.
  '
  ' These fields do NOT exist in the physical JSON ViewMap and are NOT backed by SQL.
  ' They are conceptual-only and routed at runtime to the availability evaluator via:
  '     SourceView = "vwConcept_Availability"
  '
  ' Conceptual fields added:
  '   - Availability            (Text: "Available" / "Unavailable")
  '   - AvailabilityFromDate    (Date: start of evaluation window)
  '   - AvailabilityToDate      (Date: end of evaluation window; must be >= FromDate)
  '
  ' Notes:
  '   - These fields are fully filterable and selectable by the user.
  '   - The rule engine must NOT assume any combination of these fields.
  '   - Validation enforces only one structural rule:
  '         If either AvailabilityFromDate or AvailabilityToDate is used,
  '         the other must also be present.
  '   - Runtime SQL generation must ignore these fields entirely.
  '   - Post-SQL evaluation handles availability logic based on the conceptual fields.
  '
  ' This routine must be called from BuildUiViewMap after merging custom fields.
  ' ==========================================================================================
  Private Sub MergeConceptualAvailabilityFields(uiMap As ExcelRuleViewMap, ConceptView As ExcelRuleViewMapView)

    Dim resourceView = uiMap.Views.FirstOrDefault(Function(v) v.Name = "vwDim_Resource")
    If resourceView Is Nothing Then Exit Sub

    For Each cf In ConceptView.Fields

      Dim synthetic As New ExcelRuleViewMapField With {
            .Name = cf.Name,
            .DisplayName = cf.DisplayName,
            .FieldID = cf.FieldID,
            .SourceField = cf.Name,
            .SourceView = ConceptView.Name,
            .DataType = cf.DataType,
            .Role = cf.Role,
            .AllowedOperators = cf.AllowedOperators?.ToList(),
            .AllowedValues = cf.AllowedValues?.ToList(),
            .IsTimeBounded = cf.IsTimeBounded,
            .PartnerField = cf.PartnerField,
            .RequiresPartner = cf.RequiresPartner,
            .SupportsSlicing = cf.SupportsSlicing,
            .SlicingOptions = cf.SlicingOptions?.ToList(),
            .SlicingDescriptions = If(cf.SlicingDescriptions IsNot Nothing,
                          New Dictionary(Of String, String)(cf.SlicingDescriptions),
                          Nothing),
            .DefaultSlicing = cf.DefaultSlicing,
            .TimeBoundHandler = cf.TimeBoundHandler,
            .MergedIntoView = cf.MergedIntoView
          }

      resourceView.Fields.Add(synthetic)
    Next

  End Sub

End Class