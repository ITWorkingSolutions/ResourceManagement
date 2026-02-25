' ==========================================================================================
' Module: ListItemTypeCatalog
' Purpose:
'   Defines the canonical list of all valid system ListItemTypes.
'   These values drive UI dropdowns and validation used for system fields but not the user
'   created ones.
' ==========================================================================================
Friend Module ListItemTypeSystemCatalog

  'Friend ReadOnly Property RoleFunctionTypeID As String = "A1F1E3C2-7C3A-4F8B-9E1A-001122334455"
  'Friend ReadOnly Property DepartmentTeamTypeID As String = "B2A2D4E3-8D4B-5A9C-A2B3-112233445566"
  'Friend ReadOnly Property LocationRegionTypeID As String = "C3B3F5D4-9E5C-6BAD-B3C4-223344556677"
  Friend ReadOnly Property SalutationTypeID As String = "D4C4A6E5-AF6D-7CBE-C4D5-334455667788"
  Friend ReadOnly Property GenderTypeID As String = "E5D5B7F6-B07E-8DCF-D5E6-445566778899"

  Friend Function AllSystemTypes() As List(Of RecordListItemType)
    Return New List(Of RecordListItemType) From {
      New RecordListItemType With {
        .ListItemTypeID = SalutationTypeID,
        .ListItemTypeName = "Salutation",
        .IsSystemType = 1,
        .IsNew = True
      },
      New RecordListItemType With {
        .ListItemTypeID = GenderTypeID,
        .ListItemTypeName = "Gender",
        .IsSystemType = 1,
        .IsNew = True
      }
    }
    'New RecordListItemType With {
    '  .ListItemTypeID = RoleFunctionTypeID,
    '  .ListItemTypeName = "Role / Function",
    '  .IsSystemType = 1,
    '  .IsNew = True
    '},
    'New RecordListItemType With {
    '  .ListItemTypeID = DepartmentTeamTypeID,
    '  .ListItemTypeName = "Department / Team",
    '  .IsSystemType = 1,
    '  .IsNew = True
    '},
    'New RecordListItemType With {
    '  .ListItemTypeID = LocationRegionTypeID,
    '  .ListItemTypeName = "Location / Region",
    '  .IsSystemType = 1,
    '  .IsNew = True
    '},
  End Function

End Module
