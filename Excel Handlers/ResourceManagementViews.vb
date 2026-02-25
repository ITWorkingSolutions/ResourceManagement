Imports System.Runtime.InteropServices.ComTypes

Friend Module ResourceManagementViews

  Friend Function GetActiveDatabasePath() As String
    Return AddInContext.Current.Config.ActiveDbPath
  End Function

  Friend Function GetViewAsArray(viewName As String) As Object(,)
    Dim conn As SQLiteConnectionWrapper = Nothing
    Dim cmd As SQLiteCommandWrapper
    Dim ds As SQLiteDataRowReader
    Dim sql As String
    Try
      conn = OpenDatabase(AddInContext.Current.Config.ActiveDbPath)
      ' ------------------------------------------------------------
      '  Execute query
      ' ------------------------------------------------------------
      sql = "SELECT * FROM " & viewName
      cmd = conn.CreateCommand(sql)
      ds = cmd.OpenDataSet()
      Return DataTableTo2DArray(ds)
    Finally
      If conn IsNot Nothing Then conn.Close()
    End Try
  End Function

  Friend Function GetViewAsArray(viewName As String, includeHeaders As Boolean) As Object(,)
    Dim raw = GetViewAsArray(viewName) ' existing function
    If includeHeaders Then Return raw

    ' Strip header row
    Dim rowCount = raw.GetLength(0)
    Dim colCount = raw.GetLength(1)

    ' If only header exists, return empty 1x1
    If rowCount <= 1 Then
      Dim empty(0, 0) As Object
      empty(0, 0) = ""
      Return empty
    End If

    Dim result(rowCount - 2, colCount - 1) As Object
    For r = 1 To rowCount - 1
      For c = 0 To colCount - 1
        result(r - 1, c) = raw(r, c)
      Next
    Next

    Return result
  End Function

  Private Function DataTableTo2DArray(reader As SQLiteDataRowReader) As Object(,)
    Dim rows As New List(Of SQLiteDataRow)

    ' ------------------------------------------------------------
    '  Read all rows into memory
    ' ------------------------------------------------------------
    While reader.Read()
      rows.Add(reader.Row)
    End While

    ' ------------------------------------------------------------
    '  If no rows, return a 1x1 empty array
    ' ------------------------------------------------------------
    If rows.Count = 0 Then
      Dim empty(0, 0) As Object
      empty(0, 0) = ""
      Return empty
    End If

    ' ------------------------------------------------------------
    '  Extract column names from the first row
    ' ------------------------------------------------------------
    Dim firstRow = rows(0)
    Dim colNames = firstRow.FieldNames
    Dim colCount = colNames.Count
    Dim rowCount = rows.Count

    ' ------------------------------------------------------------
    '  Create result array (include header row)
    ' ------------------------------------------------------------
    Dim result(rowCount, colCount - 1) As Object

    ' Header row
    For c = 0 To colCount - 1
      result(0, c) = colNames(c)
    Next

    ' Data rows
    For r = 0 To rowCount - 1
      Dim row = rows(r)
      For c = 0 To colCount - 1
        result(r + 1, c) = row(colNames(c))
      Next
    Next

    Return result
  End Function


End Module
