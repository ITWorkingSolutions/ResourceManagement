Imports System.Globalization

Module DateTimeHelpers
  ' ==========================================================================================
  ' Routine: GetShortDateMask
  ' Purpose: Returns the system short date mask string (numeric only, e.g. "dd/mm/yyyy").
  '
  ' Notes:
  '   - Uses Format$ with "Short Date" on a dummy date.
  '   - Detects and normalizes month names (mmm, mmmm) back to "mm".
  ' ==========================================================================================
  Friend Function GetShortDateMask() As String
    Dim s As String

    ' Use a dummy date with unique day/month/year
    s = Format$(DateSerial(2025, 11, 14), "Short Date")

    ' Normalize month names to "mm"
    If InStr(1, s, "Nov", vbTextCompare) > 0 Then
      s = Replace(s, "Nov", "mm", , , vbTextCompare)
    End If
    If InStr(1, s, "November", vbTextCompare) > 0 Then
      s = Replace(s, "November", "mm", , , vbTextCompare)
    End If

    ' Replace numeric parts with tokens
    s = Replace(s, "14", "dd")
    s = Replace(s, "11", "mm")
    s = Replace(s, "2025", "yyyy")
    s = Replace(s, "25", "yy")   ' covers two-digit year cases

    GetShortDateMask = s
  End Function

  ' Convert TEXT → DateTime?
  Friend Function ConvertTextToDate(value As String) As DateTime?
    If String.IsNullOrWhiteSpace(value) Then Return Nothing
    Dim dt As DateTime
    If DateTime.TryParse(value, dt) Then Return dt
    Return Nothing
  End Function

  ' Convert TEXT → TimeSpan?
  Friend Function ConvertTextToTimeSpan(value As String) As TimeSpan?
    If String.IsNullOrWhiteSpace(value) Then Return Nothing
    Dim dt As DateTime
    If DateTime.TryParse(value, dt) Then Return dt.TimeOfDay
    Return Nothing
  End Function

  ' ==========================================================================================
  ' Routine: SafeParseDate
  ' Purpose: Converts a string to Date, trying multiple formats and cultures.
  ' Parameters:
  '   - value: Input date string.
  ' Returns:
  '   - Parsed Date.
  ' Notes:
  ' ==========================================================================================
  Friend Function SafeParseDate(value As String) As Date
    ' ------------------------------------------------------------
    ' 1. Null or empty → error
    ' ------------------------------------------------------------
    If String.IsNullOrWhiteSpace(value) Then
      Throw New ArgumentException("Date cannot be empty.")
    End If

    ' ------------------------------------------------------------
    ' 2. Try ISO formats first (yyyy-MM-dd, yyyy/MM/dd, yyyyMMdd)
    ' ------------------------------------------------------------
    Dim isoFormats As String() = {
        "yyyy-MM-dd",
        "yyyy/MM/dd",
        "yyyyMMdd",
        "yyyy-MM-ddTHH:mm:ss",
        "yyyy-MM-dd HH:mm:ss"
    }

    Dim dt As Date
    For Each fmt In isoFormats
      If Date.TryParseExact(value, fmt, CultureInfo.InvariantCulture,
                            DateTimeStyles.None, dt) Then
        Return dt.Date
      End If
    Next

    ' ------------------------------------------------------------
    ' 3. Try invariant culture general parse
    ' ------------------------------------------------------------
    If Date.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, dt) Then
      Return dt.Date
    End If

    ' ------------------------------------------------------------
    ' 4. Try current culture (AU, US, UK, etc.)
    ' ------------------------------------------------------------
    If Date.TryParse(value, CultureInfo.CurrentCulture, DateTimeStyles.None, dt) Then
      Return dt.Date
    End If

    ' ------------------------------------------------------------
    ' 5. Try Excel serial number (e.g., "45231")
    ' ------------------------------------------------------------
    Dim dbl As Double
    If Double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, dbl) Then
      ' Excel serial date: 1 = 1 Jan 1900
      ' .NET: 1 Jan 1900 = serial 1
      Try
        dt = Date.FromOADate(dbl)
        Return dt.Date
      Catch
        ' ignore and fall through
      End Try
    End If

    ' ------------------------------------------------------------
    ' 6. If all parsing fails → throw meaningful error
    ' ------------------------------------------------------------
    Throw New ArgumentException($"Invalid date format: '{value}'")
  End Function

  ' ==========================================================================================
  ' Routine: GenerateDateList
  ' Purpose: Generates a list of dates between startDt and endDt (inclusive).
  ' Parameters:
  '   - startDt: Start date.
  '   - endDt:   End date.
  ' Returns:
  '   - List(Of Date): List of dates in the range.
  ' Notes:
  ' ==========================================================================================
  Friend Function GenerateDateList(startDt As Date, endDt As Date) As List(Of Date)
    Dim list As New List(Of Date)

    ' Guard: empty if range is inverted
    If endDt < startDt Then Return list

    Dim d As Date = startDt
    While d <= endDt
      list.Add(d)
      d = d.AddDays(1)
    End While

    Return list
  End Function
End Module
