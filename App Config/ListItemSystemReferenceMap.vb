Friend Module ListItemSystemReferenceMap

  Friend ReadOnly ListItemSystemReferences As New List(Of ListItemSystemReference) From {
    New ListItemSystemReference(ListItemTypeSystemCatalog.SalutationTypeID, "tblResource", "SalutationID"),
    New ListItemSystemReference(ListItemTypeSystemCatalog.GenderTypeID, "tblResource", "GenderID")
  }

End Module
