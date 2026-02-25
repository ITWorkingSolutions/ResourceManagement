Option Explicit On
Imports System.Text.Json.Serialization

Friend Class SchemaField

  ' ------------------------------------------------------------
  '  Properties (auto-implemented)
  ' ------------------------------------------------------------
  Friend Property Name As String
  Friend Property Type As String
  Friend Property DefaultValue As String

  <JsonPropertyName("primaryKey")>
  Friend Property IsPrimaryKey As Boolean

  <JsonPropertyName("nullable")>
  Public Property IsNullable As Boolean = True ' to handle JSON deserialization not using the constructor with parameters

  ' ------------------------------------------------------------
  '  Constructors
  ' ------------------------------------------------------------
  Friend Sub New()
    ' Parameterless constructor required for JSON deserialization
  End Sub

  Friend Sub New(name As String, type As String, Optional defaultValue As String = Nothing, Optional isPrimaryKey As Boolean = False, Optional IsNullable As Boolean = True)
    Me.Name = name
    Me.Type = type
    Me.DefaultValue = defaultValue
    Me.IsPrimaryKey = isPrimaryKey
    Me.IsNullable = IsNullable

    ' Primary keys must never be nullable
    If isPrimaryKey Then
      Me.IsNullable = False
    Else
      Me.IsNullable = IsNullable
    End If

  End Sub

End Class
