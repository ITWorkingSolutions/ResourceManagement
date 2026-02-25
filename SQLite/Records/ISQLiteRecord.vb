Option Explicit On
Friend Interface ISQLiteRecord
  ' ------------------------------------------------------------
  '  Table and schema metadata
  ' ------------------------------------------------------------
  Function TableName() As String
  Function FieldNames() As String()
  Function PrimaryKeyNames() As String()

  ' ------------------------------------------------------------
  '  Lifecycle flags
  ' ------------------------------------------------------------
  Property IsNew As Boolean
  Property IsDirty As Boolean
  Property IsDeleted As Boolean


End Interface
