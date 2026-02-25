Option Explicit On
Module GuidFactory

  ' ============================================================
  '  Description:
  '     Provides a single, authoritative GUID generator for all
  '     record creation and inserts.
  '
  '  Returns:
  '     32-character uppercase GUID with no braces or hyphens.
  '
  '  Example:
  '     "A3F9C2E1B4D84F0A9C7E12D3F8A4B6C1"
  ' ============================================================

  Friend Function NewGuid() As String
    ' Generate a standard .NET GUID
    Dim g As String = Guid.NewGuid().ToString("N").ToUpperInvariant()
    Return g
  End Function

End Module
