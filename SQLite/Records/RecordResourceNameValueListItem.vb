Friend Class RecordResourceNameValueListItem
  Implements ISQLiteRecord

  ' ------------------------------------------------------------
  '  Lifecycle flags
  ' ------------------------------------------------------------
  Friend Property IsNew As Boolean Implements ISQLiteRecord.IsNew
  Friend Property IsDirty As Boolean Implements ISQLiteRecord.IsDirty
  Friend Property IsDeleted As Boolean Implements ISQLiteRecord.IsDeleted

  ' ------------------------------------------------------------
  '  Private backing fields
  ' ------------------------------------------------------------
  Private mResourceID As String
  Private mResourceListItemID As String
  Private mListItemID As String

  ' ------------------------------------------------------------
  '  Properties
  ' ------------------------------------------------------------
  Friend Property ResourceID As String
    Get
      Return mResourceID
    End Get
    Set(value As String)
      If mResourceID <> value Then
        mResourceID = value
        IsDirty = True
      End If
    End Set
  End Property

  Friend Property ResourceListItemID As String
    Get
      Return mResourceListItemID
    End Get
    Set(value As String)
      If mResourceListItemID <> value Then
        mResourceListItemID = value
        IsDirty = True
      End If
    End Set
  End Property

  ' Selected ListItemID when the value is lookup-based
  Friend Property ListItemID As String
    Get
      Return mListItemID
    End Get
    Set(value As String)
      If mListItemID <> value Then
        mListItemID = value
        IsDirty = True
      End If
    End Set
  End Property

  ' The value mode of the resource value TEXT, SINGLE_LIST, MULTI_LIST

  ' ------------------------------------------------------------
  '  Logical delete
  ' ------------------------------------------------------------
  Friend Sub Delete()
    IsDeleted = True
  End Sub

  ' ------------------------------------------------------------
  '  ISQLiteRecord implementation
  ' ------------------------------------------------------------
  Friend Function FieldNames() As String() Implements ISQLiteRecord.FieldNames
    Return {
      "ResourceID",
      "ResourceListItemID",
      "ListItemID"
    }
  End Function

  Friend Function TableName() As String Implements ISQLiteRecord.TableName
    Return "tblResourceNameValueListItem"
  End Function

  Friend Function PrimaryKeyNames() As String() Implements ISQLiteRecord.PrimaryKeyNames
    Return {
      "ResourceID",
      "ResourceListItemID",
      "ListItemID"
    }
  End Function
End Class
