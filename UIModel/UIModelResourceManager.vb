' ==========================================================================================
' Enum: ResourceManagerFilterMode
' Purpose:
'   Indicates which field the ResourceManager should filter on.
' ==========================================================================================
Friend Enum ResourceManagerFilterMode
  PreferredName = 0
  FullName = 1
  EmployeeID = 2
End Enum
' ==========================================================================================
' Class: UIModelResourceManager
' Purpose:
'   UI model for the ResourceManager form.
'   Holds the resource summary list, availability summary list, and filter state.
' ==========================================================================================
Friend Class UIModelResourceManager

  ' --- Primary resource list (summary rows) ---
  Friend Property ResourceSummaries As List(Of UIResourceManagerResourceSummaryRow)

  ' --- Availability list for selected resource ---
  Friend Property AvailabilitySummaries As List(Of UIResourceManagerAvailabilitySummaryRow)
  ' --- Currently selected resource ---
  Friend Property SelectedResourceID As String
  ' --- Currently selected resource ---
  Friend Property SelectedAvailabilityID As String
  ' --- Filtering state ---
  Friend Property ShowInactive As Boolean
  Friend Property FilterText As String
  Friend Property FilterMode As ResourceManagerFilterMode

  Friend Sub New()
    ResourceSummaries = New List(Of UIResourceManagerResourceSummaryRow)()
    AvailabilitySummaries = New List(Of UIResourceManagerAvailabilitySummaryRow)()
    ShowInactive = False
    FilterText = ""
    FilterMode = ResourceManagerFilterMode.PreferredName
  End Sub

End Class

' ==========================================================================================
' Class: UIResourceManagerResourceSummaryRow
' Purpose:
'   Represents a single row in the ResourceManager's primary resource list.
'   Contains only the fields required for list display and filtering.
' ==========================================================================================
Friend Class UIResourceManagerResourceSummaryRow

  ' --- Identity ---
  Public Property ResourceID As String

  ' --- Display fields ---
  Public Property PreferredName As String
  Public Property FullName As String        ' FirstName + " " + LastName
  'Friend Property EmployeeID As String

  ' --- Status ---
  Public Property IsInactive As Boolean     ' Derived from EndDate < Today

End Class

' ==========================================================================================
' Class: UIResourceManagerAvailabilitySummaryRow
' Purpose:
'   Represents a single availability summary row for the selected resource.
'   Contains only the AvailabilityID and the preformatted description string.
' ==========================================================================================
Friend Class UIResourceManagerAvailabilitySummaryRow

  ' --- Identity ---
  Public Property AvailabilityID As String

  ' --- Display ---
  Public Property Description As String     ' Built using BuildAvailabilityDescription()

  Public Overrides Function ToString() As String
    Return Description
  End Function

End Class