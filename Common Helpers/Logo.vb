Option Explicit On
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms

Module Logo
  Friend Sub SaveCompanyLogo()

    Dim dlg As OpenFileDialog
    Dim img As Image
    Dim resized As Image
    Dim bytes As Byte()
    Dim rec As RecordBinaryAsset
    Dim conn As SQLiteConnectionWrapper = Nothing

    ' ------------------------------------------------------------
    '  Select image
    ' ------------------------------------------------------------
    dlg = New OpenFileDialog()
    dlg.Title = "Select Company Logo"
    dlg.Filter = "Image Files (*.png;*.jpg;*.jpeg;*.bmp)|*.png;*.jpg;*.jpeg;*.bmp"

    If dlg.ShowDialog() <> DialogResult.OK Then Exit Sub

    img = Image.FromFile(dlg.FileName)

    ' ------------------------------------------------------------
    '  Resize for report use (UI-side only)
    ' ------------------------------------------------------------
    resized = ResizeImageForReport(img)

    bytes = ImageToPngBytes(resized)

    ' ------------------------------------------------------------
    '  Create record (standard pattern)
    ' ------------------------------------------------------------
    rec = New RecordBinaryAsset()
    rec.Key = "CompanyLogo"
    rec.MimeType = "image/png"
    rec.Data = bytes

    ' ------------------------------------------------------------
    '  Save using standard DB routines
    '  Uses the existing OpenDatabase(...) and RecordSaver.SaveRecord(...)
    ' ------------------------------------------------------------
    Try
      ' Ensure an active DB is configured and open a validated connection
      conn = OpenDatabase(AddInContext.Current.Config.ActiveDbPath)

      ' Check if the record already exists to set IsNew flag
      Dim pkValues(0) As String
      pkValues(0) = "CompanyLogo"

      Dim existing As RecordBinaryAsset
      existing = RecordLoader.LoadRecord(Of RecordBinaryAsset)(conn, pkValues)

      If existing Is Nothing Then
        rec.IsNew = True
      Else
        rec.IsDirty = True
      End If

      RecordSaver.SaveRecord(conn, rec)
      ' Persist the binary asset using the existing persistence routine
      RecordSaver.SaveRecord(conn, rec)

      MessageBoxHelper.Show(UIMessageOwner.ExcelOwner, "Company logo saved successfully.",
                      "Saved",
                      MessageBoxButtons.OK,
                      MessageBoxIcon.Information)

    Catch ex As Exception
      MessageBoxHelper.Show(UIMessageOwner.ExcelOwner, "Failed to save company logo." & vbCrLf & ex.Message,
                      "Save Error",
                      MessageBoxButtons.OK,
                      MessageBoxIcon.Error)
    Finally
      If conn IsNot Nothing Then conn.Close()
    End Try


  End Sub

  Private Function ResizeImageForReport(src As Image) As Image

    Dim maxWidth As Integer
    Dim ratio As Double
    Dim newWidth As Integer
    Dim newHeight As Integer
    Dim bmp As Bitmap
    Dim g As Graphics

    maxWidth = 256 ' or whatever your report standard is

    ratio = src.Width / CDbl(src.Height)
    newWidth = maxWidth
    newHeight = CInt(maxWidth / ratio)

    bmp = New Bitmap(newWidth, newHeight)

    g = Graphics.FromImage(bmp)
    g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
    g.DrawImage(src, 0, 0, newWidth, newHeight)
    g.Dispose()

    Return bmp

  End Function

  Private Function ImageToPngBytes(img As Image) As Byte()
    Dim ms As New MemoryStream()
    img.Save(ms, Imaging.ImageFormat.Png)
    Return ms.ToArray()
  End Function
End Module
