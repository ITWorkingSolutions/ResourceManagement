Friend Class ListItemSystemReference
  Friend ReadOnly Property ListItemTypeID As String
  Friend ReadOnly Property TableName As String
  Friend ReadOnly Property FieldName As String

  Friend Sub New(listItemTypeID As String, tableName As String, fieldName As String)
    Me.ListItemTypeID = listItemTypeID
    Me.TableName = tableName
    Me.FieldName = fieldName
  End Sub
End Class

