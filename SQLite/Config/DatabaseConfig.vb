' ============================================================================================
'  Class: DatabaseConfig
'  Purpose:
'       Holds the active database path and the list of known database paths.
'
'  Notes:
'       - Serialized to JSON in the user's AppData folder.
'       - Need to be Public for JSON serialization.
' ============================================================================================
Public Class DatabaseConfig

  Public Property ActiveDbPath As String
  Public Property KnownDbPaths As List(Of String)

  Public Sub New()
    KnownDbPaths = New List(Of String)
  End Sub

End Class
