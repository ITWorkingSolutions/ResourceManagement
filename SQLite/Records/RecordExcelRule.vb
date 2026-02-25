Friend Class RecordExcelRule
  Implements ISQLiteRecord

  ' ------------------------------------------------------------
  '  Lifecycle flags (runtime only)
  ' ------------------------------------------------------------
  Friend Property IsNew As Boolean Implements ISQLiteRecord.IsNew
  Friend Property IsDirty As Boolean Implements ISQLiteRecord.IsDirty
  Friend Property IsDeleted As Boolean Implements ISQLiteRecord.IsDeleted

  ' ------------------------------------------------------------
  '  Private backing fields
  ' ------------------------------------------------------------
  Private mRuleID As String
  Private mRuleName As String
  Private mRuleType As String
  Private mDefinitionJson As String

  ' ------------------------------------------------------------
  '  Properties with change tracking
  ' ------------------------------------------------------------
  Friend Property RuleID As String
    Get
      Return mRuleID
    End Get
    Set(value As String)
      If mRuleID <> "" AndAlso mRuleID <> value Then
        Throw New InvalidOperationException("RuleID cannot be modified after creation.")
      End If

      If mRuleID <> value Then
        mRuleID = value
        IsDirty = True
      End If
    End Set
  End Property
  Friend Property RuleName As String
    Get
      Return mRuleName
    End Get
    Set(value As String)
      If mRuleName <> value Then
        mRuleName = value
        IsDirty = True
      End If
    End Set
  End Property
  Friend Property RuleType As String
    Get
      Return mRuleType
    End Get
    Set(value As String)
      If mRuleType <> value Then
        mRuleType = value
        IsDirty = True
      End If
    End Set
  End Property
  Friend Property DefinitionJson As String
    Get
      Return mDefinitionJson
    End Get
    Set(value As String)
      If mDefinitionJson <> value Then
        mDefinitionJson = value
        IsDirty = True
      End If
    End Set
  End Property

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
      "RuleID",
      "RuleName",
      "RuleType",
      "DefinitionJson"
    }
  End Function

  Friend Function TableName() As String Implements ISQLiteRecord.TableName
    Return "tblExcelRule"
  End Function

  Friend Function PrimaryKeyNames() As String() Implements ISQLiteRecord.PrimaryKeyNames
    Return {"RuleID"}
  End Function
End Class

' ==========================================================================================
' Runtime DTO: RecordExcelRuleRowDetail
' Purpose:
'   Pure runtime representation of DefinitionJson stored in tblExcelRule.
'   Mirrors the JSON schema exactly, but contains NO UI-only fields or logic.
' ==========================================================================================
Friend Class RecordExcelRuleRowDetail
  Public Property RuleID As String
  Public Property RuleName As String
  Public Property RuleType As String

  Public Property PrimaryView As String
  Public Property SelectedValues As New List(Of RecordRuleSelectedValue)
  Public Property Filters As New List(Of RecordRuleFilter)
  Public Property UsedViews As New List(Of String)
End Class

Friend Class RecordRuleSelectedValue
  Public Property View As String
  Public Property Field As String
  Public Property ListTypeID As String
End Class

Friend Class RecordRuleFilter
  Public Property FilterID As String
  Public Property View As String
  Public Property Field As String
  Public Property FieldOperator As String
  Public Property BooleanOperator As String
  Public Property OpenParenCount As Integer
  Public Property CloseParenCount As Integer
  Public Property ListTypeID As String
End Class
