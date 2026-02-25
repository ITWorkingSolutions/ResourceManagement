Option Explicit On

Friend Class SchemaForeignKey

  ' ------------------------------------------------------------
  '  Properties (auto-implemented)
  ' ------------------------------------------------------------
  Friend Property Field As List(Of String)
  Friend Property ReferencesTable As String
  Friend Property ReferencesField As List(Of String)
  Friend Property OnDelete As String

  ' ------------------------------------------------------------
  '  Constructors
  ' ------------------------------------------------------------
  Friend Sub New()
    ' Required for JSON deserialization
  End Sub

  Friend Sub New(field As List(Of String),
                 referencesTable As String,
                 referencesField As List(Of String),
                 Optional onDelete As String = Nothing)

    Me.Field = field
    Me.ReferencesTable = referencesTable
    Me.ReferencesField = referencesField
    Me.OnDelete = onDelete
  End Sub


End Class
