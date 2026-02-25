Option Explicit On
Imports System.IO
Imports System.Reflection
Imports System.Text.Json

Module ExcelRuleViewMapLoader

  ' ============================================================
  '  LoadExcelRuleViewMap
  '
  '  Loads ExcelRuleViewMap.json embedded in the DLL and
  '  deserializes it into a strongly typed ExcelRuleViewMap.
  '
  '  This is the ONLY source of truth for view metadata.
  ' ============================================================
  Friend Function LoadExcelRuleViewMap() As ExcelRuleViewMap
    Dim jsonText As String = LoadEmbeddedViewMapJson()
    Return DeserializeViewMap(jsonText)
  End Function


  ' ------------------------------------------------------------
  '  LoadEmbeddedViewMapJson
  '
  '  Reads ExcelRuleViewMap.json from embedded resources.
  ' ------------------------------------------------------------
  Private Function LoadEmbeddedViewMapJson() As String
    Dim asm = Assembly.GetExecutingAssembly()

    Dim resourceName As String =
        asm.GetManifestResourceNames().
            First(Function(n) n.EndsWith("ExcelRuleViewMap.json",
                    StringComparison.OrdinalIgnoreCase))

    Using stream = asm.GetManifestResourceStream(resourceName)
      Using reader As New StreamReader(stream)
        Return reader.ReadToEnd()
      End Using
    End Using
  End Function


  ' ------------------------------------------------------------
  '  DeserializeViewMap
  '
  '  Converts JSON text into ExcelRuleViewMap.
  ' ------------------------------------------------------------
  Private Function DeserializeViewMap(jsonText As String) As ExcelRuleViewMap

    Dim options As JsonSerializerOptions = Nothing
    Dim map As ExcelRuleViewMap = Nothing

    Try
      options = New JsonSerializerOptions With {
          .PropertyNameCaseInsensitive = True
      }

      map = JsonSerializer.Deserialize(Of ExcelRuleViewMap)(jsonText, options)

      If map Is Nothing OrElse map.Views Is Nothing Then
        Throw New InvalidDataException("ExcelRuleViewMap contains no views.")
      End If

      ' --------------------------------------------------------
      '  Ensure lists are initialized
      ' --------------------------------------------------------
      For Each v In map.Views
        If v.Fields Is Nothing Then v.Fields = New List(Of ExcelRuleViewMapField)
        If v.Relations Is Nothing Then v.Relations = New List(Of ExcelRuleViewMapRelation)
      Next

      Return map

    Catch ex As Exception
      Throw New InvalidDataException("ExcelRuleViewMap could not be deserialized.", ex)

    Finally
      options = Nothing
      map = Nothing
    End Try

  End Function

End Module
