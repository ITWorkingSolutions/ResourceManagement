Option Explicit On

Public Class SchemaForeignKey

  ' ------------------------------------------------------------
  '  Properties (auto-implemented)
  ' ------------------------------------------------------------
  Public Property Field As List(Of String)
  Public Property ReferencesTable As String
  Public Property ReferencesField As List(Of String)
  Public Property OnDelete As String

  ' ------------------------------------------------------------
  '  Constructors
  ' ------------------------------------------------------------
  Public Sub New()
    ' Required for JSON deserialization
  End Sub

  Public Sub New(field As List(Of String),
                 referencesTable As String,
                 referencesField As List(Of String),
                 Optional onDelete As String = Nothing)

    Me.Field = field
    Me.ReferencesTable = referencesTable
    Me.ReferencesField = referencesField
    Me.OnDelete = onDelete
  End Sub


End Class
