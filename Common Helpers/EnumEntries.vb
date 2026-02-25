Imports System.Drawing
Imports System.Reflection

Public Module EnumEntries
  Friend Enum FieldCardinality
    One
    Many
  End Enum
  ' ==========================================================================================
  ' Routine: GetCardinality
  ' Purpose:
  '   Mapping the ResourceListItemValueType to FieldCardinality.
  ' Parameters:
  '   valueType - The ResourceListItemValueType to map.
  ' Returns:
  '   FieldCardinality corresponding to the given ResourceListItemValueType.
  ' Notes:
  '   - Used for determining how many values a field can hold based on its type.
  ' ==========================================================================================
  Friend Function GetCardinality(valueType As ResourceListItemValueType) As FieldCardinality
    Select Case valueType
      Case ResourceListItemValueType.Text
        Return FieldCardinality.One

      Case ResourceListItemValueType.SingleSelectList
        Return FieldCardinality.One

      Case ResourceListItemValueType.MultiSelectList
        Return FieldCardinality.Many

      Case Else
        Throw New InvalidOperationException("Unknown ValueType: " & valueType.ToString())
    End Select
  End Function
  Public Enum ExcelRefType
    <EnumDisplay("Literal")>
    Literal
    <EnumDisplay("Absolute Range")>
    Address
    <EnumDisplay("Relative Range")>
    Offset
    <EnumDisplay("Named Range")>
    Name
  End Enum
  Friend ReadOnly ExcelRefTypeMap As New EnumMap(Of ExcelRefType)()

  Public Enum ExcelRuleType
    <EnumDisplay("Single Value")>
    SingleValue
    <EnumDisplay("List of Values")>
    ListOfValues
    <EnumDisplay("Range of Values")>
    RangeOfValues
  End Enum
  Friend ReadOnly ExcelRuleTypeMap As New EnumMap(Of ExcelRuleType)()

  Public Enum ResourceListItemValueType
    <EnumDisplay("Text")>
    Text
    <EnumDisplay("Single Select List")>
    SingleSelectList
    <EnumDisplay("Multi Select List")>
    MultiSelectList
  End Enum
  Friend ReadOnly ResourceListItemValueTypeMap As New EnumMap(Of ResourceListItemValueType)()
  Public Enum ExcelListSelectType
    <EnumDisplay("Single Select")>
    SingleSelect

    <EnumDisplay("Multi Select")>
    MultiSelect
  End Enum
  Friend ReadOnly ExcelListSelectTypeMap As New EnumMap(Of ExcelListSelectType)()

  Public Enum ValueBinding
    <EnumDisplay("Parameter")>
    Parameter

    <EnumDisplay("Rule")>
    Rule
  End Enum
  Friend ReadOnly ValueBindingMap As New EnumMap(Of ValueBinding)()

End Module

<AttributeUsage(AttributeTargets.Field)>
Public Class EnumDisplayAttribute
  Inherits Attribute
  Public ReadOnly Property Display As String
  Public Sub New(display As String)
    Me.Display = display
  End Sub
End Class

Public Class BindingItem(Of TEnum As Structure)
  Public ReadOnly Property EnumValue As TEnum
  Public ReadOnly Property Display As String
  Public Sub New(enumValue As TEnum, display As String)
    Me.EnumValue = enumValue
    Me.Display = display
  End Sub
  Public Overrides Function ToString() As String
    Return Display
  End Function
End Class

Public Class BindingItemString
  Public ReadOnly Property Value As String
  Public ReadOnly Property Display As String

  Public Sub New(value As String, display As String)
    Me.Value = value
    Me.Display = display
  End Sub

  Public Overrides Function ToString() As String
    Return Display
  End Function
End Class


Public NotInheritable Class EnumMap(Of TEnum As Structure)

  Private ReadOnly _enumToDisplay As Dictionary(Of TEnum, String)
  Private ReadOnly _displayToEnum As Dictionary(Of String, TEnum)

  Public Sub New()
    Dim enumType = GetType(TEnum)
    If Not enumType.IsEnum Then
      Throw New ArgumentException($"{enumType.Name} is not an Enum.")
    End If

    Dim pairs As New List(Of (EnumValue As TEnum, Display As String))

    For Each field In enumType.GetFields(BindingFlags.Public Or BindingFlags.Static)
      Dim enumValue = CType([Enum].Parse(enumType, field.Name), TEnum)
      Dim displayAttr = CType(field.GetCustomAttributes(GetType(EnumDisplayAttribute), False).FirstOrDefault(), EnumDisplayAttribute)

      If displayAttr Is Nothing Then
        Throw New InvalidOperationException($"Enum member '{field.Name}' is missing EnumDisplay attribute.")
      End If

      pairs.Add((enumValue, displayAttr.Display))
    Next

    _enumToDisplay = pairs.ToDictionary(Function(p) p.EnumValue, Function(p) p.Display)
    _displayToEnum = pairs.ToDictionary(Function(p) p.Display, Function(p) p.EnumValue, StringComparer.OrdinalIgnoreCase)
  End Sub

  Public Function Display(e As TEnum) As String
    Return _enumToDisplay(e)
  End Function

  Public Function FromDisplay(display As String) As TEnum
    Return _displayToEnum(display)
  End Function

  Public Function AllEnums() As List(Of TEnum)
    Return _enumToDisplay.Keys.ToList()
  End Function

  Public Function AllDisplays() As List(Of String)
    Return _enumToDisplay.Values.ToList()
  End Function

  Public Function BindingList() As List(Of BindingItem(Of TEnum))
    Return _enumToDisplay.Keys.Select(Function(e) New BindingItem(Of TEnum)(e, _enumToDisplay(e))).ToList()
  End Function
  Public Function DisplayFromString(value As String) As String
    If String.IsNullOrWhiteSpace(value) Then Return Nothing

    Dim enumValue As TEnum

    ' Try to parse the string into the enum
    If [Enum].IsDefined(GetType(TEnum), value) Then
      enumValue = CType([Enum].Parse(GetType(TEnum), value), TEnum)
      Return _enumToDisplay(enumValue)
    End If

    Return Nothing
  End Function


  Public Function BindingListOfStrings() As List(Of BindingItemString)
    Return _enumToDisplay.Keys.
        Select(Function(e) New BindingItemString(e.ToString(), _enumToDisplay(e))).
        ToList()
  End Function

End Class
