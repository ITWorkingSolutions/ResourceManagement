' ============================================================================================
'  Class: DatabaseConfigManager
'  Purpose:
'       Loads and saves DatabaseConfig to the user's AppData folder.
'
'  Notes:
'       - Config file: %APPDATA%\ITWorkingSolutions\ResourceManagement\config.json
'       - Ensures folder exists.
'       - Ensures config file exists.
' ============================================================================================
Imports System.IO
Imports System.Text.Json

Friend Class DatabaseConfigManager

  Private Shared ReadOnly CompanyFolder As String =
      Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                   "ITWorkingSolutions")

  Private Shared ReadOnly AddinFolder As String =
      Path.Combine(CompanyFolder, "ResourceManagement")

  Private Shared ReadOnly ConfigPath As String =
      Path.Combine(AddinFolder, "config.json")

  ' ========================================================================================
  '  Routine: Load
  '  Purpose:
  '       Loads the DatabaseConfig from disk. Creates a default config if missing or invalid.
  '
  '  Returns:
  '       DatabaseConfig - fully populated configuration object.
  ' ========================================================================================
  Friend Shared Function Load() As DatabaseConfig

    ' === Variable declarations ===
    Dim json As String = Nothing
    Dim config As DatabaseConfig = Nothing

    Try
      ' --- Ensure folder exists ---
      If Not Directory.Exists(AddinFolder) Then
        Directory.CreateDirectory(AddinFolder)
      End If

      ' --- If config file does not exist, create default ---
      If Not File.Exists(ConfigPath) Then
        config = New DatabaseConfig()
        Save(config)
        Return config
      End If

      ' --- Read file ---
      json = File.ReadAllText(ConfigPath)

      ' --- Deserialize ---
      config = JsonSerializer.Deserialize(Of DatabaseConfig)(json)

      ' --- Safety: ensure object and list are not null ---
      If config Is Nothing Then
        config = New DatabaseConfig()
      End If

      If config.KnownDbPaths Is Nothing Then
        config.KnownDbPaths = New List(Of String)
      End If

      Return config

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "DatabaseConfigManager.Load")
      Return New DatabaseConfig()

    Finally
      ' --- Cleanup ---
      json = Nothing
      config = Nothing
    End Try

  End Function

  ' ========================================================================================
  '  Routine: Save
  '  Purpose:
  '       Saves the DatabaseConfig to disk as JSON.
  '
  '  Parameters:
  '       config - DatabaseConfig instance to save.
  '
  '  Returns:
  '       None
  ' ========================================================================================
  Friend Shared Sub Save(ByVal config As DatabaseConfig)

    ' === Variable declarations ===
    Dim json As String = Nothing

    Try
      ' --- Ensure folder exists ---
      If Not Directory.Exists(AddinFolder) Then
        Directory.CreateDirectory(AddinFolder)
      End If

      ' --- Serialize ---
      json = JsonSerializer.Serialize(config, New JsonSerializerOptions With {
          .WriteIndented = True
      })

      ' --- Write file ---
      File.WriteAllText(ConfigPath, json)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "DatabaseConfigManager.Save")

    Finally
      ' --- Cleanup ---
      json = Nothing
    End Try

  End Sub

End Class
