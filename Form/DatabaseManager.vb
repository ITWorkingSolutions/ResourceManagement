'------------------------------------------------------------------------------
' DatabaseManagement.vb
' UI dialog for managing known SQLite database paths used by the add-in.
' Responsibilities:
'   - Present known database paths
'   - Allow Add / Update / Activate / Delete operations
'   - Persist changes to AddInContext.Current.Config via DatabaseConfigManager
' Notes:
'   - Uses owner-draw list to highlight the active path
'   - Relies on OpenDatabase(...) to validate/create DBs and on ExcelWindowHelper
'------------------------------------------------------------------------------
Imports System.Drawing
Imports System.Windows.Forms

Friend Class DatabaseManager
  ' UI model backing the list of database paths and active path
  Private _model As UIModelDatabaseManagement

  Friend Sub New()
    Try
      ' Disable WinForms autoscaling completely
      Me.AutoScaleMode = AutoScaleMode.None
      InitializeComponent()

      ' Apply DPI scaling AFTER controls exist
      ApplyDpiScaling(Me)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' Cleanup
    End Try
  End Sub

  '---------------------------------------------------------------------------
  ' Routine: DatabaseManagement_Load
  ' Purpose: Initialize dialog on load.
  '   - Center the form relative to Excel
  '   - Populate the UI model from AddInContext.Current.Config
  '   - Bind model paths to the list control and enable owner-draw
  ' Parameters:
  '   sender - event sender
  '   e      - Load event args
  '---------------------------------------------------------------------------
  Private Sub DatabaseManagement_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    Try
      ' === Center form relative to Excel ===
      FormHelpers.CenterFormOnExcel(Me)
      ' === Add ToolTip to controls ===
      ToolTip1.SetToolTip(btnAddNew, "Create a new database to add to the list.")
      ToolTip1.SetToolTip(btnAddExisting, "Select an existing database to add to the list.")
      ToolTip1.SetToolTip(btnUpdate, "Update the selected database path in the list.")
      ToolTip1.SetToolTip(btnDelete, "Remove the selected database path from the list.")
      ToolTip1.SetToolTip(btnActivate, "Set the selected database path as the active database.")

      ' === Load model using UILoaderSaver ===
      _model = New UIModelDatabaseManagement
      _model = LoadDatabaseManagementModel()

      ' === UI model to lstDatabasePaths  ===
      lstDatabasePaths.DataSource = _model.Paths

      ' === Set the drawing mode allow us to highlight the active path in the list ===
      lstDatabasePaths.DrawMode = DrawMode.OwnerDrawFixed
    Catch ex As UserFriendlyException
      MessageBoxHelper.Show(UIMessageOwner.ExcelOwner, ex.Message,
              "Error",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error
        )
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  '---------------------------------------------------------------------------
  ' Routine: lstDatabasePaths_MouseMove
  ' Purpose: Show tooltip with full path for the item under the mouse.
  ' Parameters:
  '   sender - event sender
  '   e      - Mouse event args (location used to find index)
  '---------------------------------------------------------------------------
  Private Sub lstDatabasePaths_MouseMove(sender As Object, e As MouseEventArgs) Handles lstDatabasePaths.MouseMove
    ' === Determine item under mouse ===
    Dim index As Integer = lstDatabasePaths.IndexFromPoint(e.Location)
    If index >= 0 AndAlso index < lstDatabasePaths.Items.Count Then
      '=== Get full path and set tooltip ===
      Dim fullPath As String = CType(lstDatabasePaths.Items(index), String)
      '=== Determine if active  ===
      Dim isActive As Boolean =
        String.Equals(fullPath, _model.ActivePath, StringComparison.OrdinalIgnoreCase)
      If isActive Then
        fullPath = "Active Database - " & fullPath
      End If
      ToolTip1.SetToolTip(lstDatabasePaths, fullPath)
      End If
  End Sub

  '---------------------------------------------------------------------------
  ' Routine: lstDatabasePaths_Format
  ' Purpose: Convert the bound full path value to a display value (file name).
  ' Parameters:
  '   sender - event sender
  '   e      - ListControlConvertEventArgs (contains ListItem and Value)
  '---------------------------------------------------------------------------
  Private Sub lstDatabasePaths_Format(sender As Object, e As ListControlConvertEventArgs) Handles lstDatabasePaths.Format
    Dim fullPath As String = CType(e.ListItem, String)
    ' === displays only the file name ===
    e.Value = IO.Path.GetFileName(fullPath)
  End Sub

  '---------------------------------------------------------------------------
  ' Routine: lstDatabasePaths_DrawItem
  ' Purpose: Owner-draw list item. Bold and color the active database entry.
  ' Parameters:
  '   sender - event sender
  '   e      - DrawItemEventArgs (graphics, bounds, font, etc.)
  '---------------------------------------------------------------------------
  Private Sub lstDatabasePaths_DrawItem(sender As Object, e As DrawItemEventArgs) Handles lstDatabasePaths.DrawItem
    ' === Routine to highlight the active database path ===

    If e.Index < 0 Then Return

    Dim fullPath As String = CType(lstDatabasePaths.Items(e.Index), String)
    Dim isActive As Boolean =
        String.Equals(fullPath, _model.ActivePath, StringComparison.OrdinalIgnoreCase)

    ' Always let Windows draw the selection background
    e.DrawBackground()

    ' Choose font
    Dim fontToUse As Font =
        If(isActive, New Font(e.Font, FontStyle.Regular), e.Font)

    ' Choose colour
    Dim textColor As Color =
        If(isActive, Color.Red, e.ForeColor)

    ' Draw text
    Dim text As String = IO.Path.GetFileName(fullPath)
    Using brush As New SolidBrush(textColor)
      e.Graphics.DrawString(text, fontToUse, brush, e.Bounds.X + 2, e.Bounds.Y + 1)
    End Using

    e.DrawFocusRectangle()
  End Sub

  '---------------------------------------------------------------------------
  '  Routine: btnAddNew_Click
  '  Purpose: Allows the user to create a new database file.
  '           - Ensures the selected filename does NOT already exist.
  '           - Calls OpenDatabase(path) to create/validate lifecycle.
  '           - Adds path to model/config and persists configuration.
  '
  '  Parameters:
  '    sender - event sender
  '    e      - Click event args
  '---------------------------------------------------------------------------
  Private Sub btnAddNew_Click(sender As Object, e As EventArgs) Handles btnAddNew.Click
    ' === Variable declarations ===
    Dim dlg As SaveFileDialog
    Dim selectedPath As String
    Try
      ' === Configure dialog ===
      dlg = New SaveFileDialog()
      dlg.Title = "Create New Database"
      dlg.Filter = "SQLite Database (*.db)|*.db|All Files (*.*)|*.*"
      dlg.DefaultExt = "db"
      dlg.AddExtension = True
      dlg.FileName = "NewDatabase.db"

      ' === Show dialog ===
      If dlg.ShowDialog() <> DialogResult.OK Then Exit Sub

      selectedPath = dlg.FileName

      ' ***********************************************************************************************
      '  Enforce: File MUST NOT already exist
      ' ***********************************************************************************************
      If IO.File.Exists(selectedPath) Then
        MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
      "The selected file already exists. Please choose a new filename that does not already exist.",
      "File Already Exists",
      MessageBoxButtons.OK,
      MessageBoxIcon.Warning)
        Exit Sub
      End If

      '  === Update model to add new database ===
      _model.NewPath = selectedPath
      _model.PendingAction = DatabaseManagementAction.AddNew
      SaveDatabaseManagementModel(_model)

      lstDatabasePaths.Invalidate()
      MessageBoxHelper.Show(UIMessageOwner.ExcelOwner, "New database path added successfully.",
                "Added",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)
    Catch ex As UserFriendlyException
      MessageBoxHelper.Show(UIMessageOwner.ExcelOwner, ex.Message,
              "Error",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error
        )
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try

  End Sub

  '---------------------------------------------------------------------------
  ' Routine: btnAddExisting_Click
  ' Purpose: Allows the user to select an existing SQLite database file without overwrite prompts.
  '   - Shows a OpenFileDialog (select an existing)
  '   - Calls OpenDatabase(path) to validate exsiting database
  '   - Adds path to model/config and persists configuration
  ' Parameters:
  '   sender - event sender
  '   e      - Click event args
  '---------------------------------------------------------------------------
  Private Sub btnAddExisting_Click(sender As Object, e As EventArgs) Handles btnAddExisting.Click

    Try
      ' === Variable declarations ===
      Dim dlg As OpenFileDialog
      Dim selectedPath As String
      Dim conn As SQLiteConnectionWrapper = Nothing

      ' === Configure dialog for existing files ===
      dlg = New OpenFileDialog()
      dlg.Title = "Select Existing Database"
      dlg.Filter = "SQLite Database (*.db)|*.db|All Files (*.*)|*.*"
      dlg.DefaultExt = "db"
      dlg.CheckFileExists = True

      If dlg.ShowDialog() <> DialogResult.OK Then Exit Sub

      selectedPath = dlg.FileName

      ' === Update model to add existing database ===
      _model.NewPath = selectedPath
      _model.PendingAction = DatabaseManagementAction.AddExisting
      SaveDatabaseManagementModel(_model)

      lstDatabasePaths.Invalidate()
      MessageBoxHelper.Show(UIMessageOwner.ExcelOwner, "Existing database path added successfully.",
                      "Added",
                      MessageBoxButtons.OK,
                      MessageBoxIcon.Information)
    Catch ex As UserFriendlyException
      MessageBoxHelper.Show(UIMessageOwner.ExcelOwner, ex.Message,
              "Error",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error
        )
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub
  '---------------------------------------------------------------------------
  ' Routine: btnActivate_Click
  ' Purpose: Set the selected path as the active database and persist config.
  ' Parameters:
  '   sender - event sender
  '   e      - Click event args
  '---------------------------------------------------------------------------
  Private Sub btnActivate_Click(sender As Object, e As EventArgs) Handles btnActivate.Click
    Try
      Dim index As Integer = lstDatabasePaths.SelectedIndex
      If index = -1 Then Exit Sub
      ' === Get selected path ===
      Dim selectedPath As String = CType(lstDatabasePaths.Items(index), String)
      ' === Update model to remove the selected path ===
      _model.SelectedPath = selectedPath
      _model.PendingAction = DatabaseManagementAction.Activate
      SaveDatabaseManagementModel(_model)

      lstDatabasePaths.Invalidate()   ' If owner-drawing active item
      MessageBoxHelper.Show(UIMessageOwner.ExcelOwner, "Database path activated successfully.",
                  "Updated",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Information)
    Catch ex As UserFriendlyException
      MessageBoxHelper.Show(UIMessageOwner.ExcelOwner, ex.Message,
              "Error",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error
        )
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  '---------------------------------------------------------------------------
  ' Routine: btnDelete_Click
  ' Purpose: Remove the selected path from the model/config. If it was active,
  '          clear the active selection.
  ' Parameters:
  '   sender - event sender
  '   e      - Click event args
  '---------------------------------------------------------------------------
  Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
    Try
      Dim index As Integer = lstDatabasePaths.SelectedIndex
      If index = -1 Then Exit Sub
      ' === Get selected path ===
      Dim selectedPath As String = CType(lstDatabasePaths.Items(index), String)
      ' === Update model to remove the selected path ===
      _model.SelectedPath = selectedPath
      _model.PendingAction = DatabaseManagementAction.Delete
      SaveDatabaseManagementModel(_model)

      lstDatabasePaths.Invalidate()   ' If owner-drawing active item
      MessageBoxHelper.Show(UIMessageOwner.ExcelOwner, "Database path deleted successfully.",
                  "Deleted",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Information)
    Catch ex As UserFriendlyException
      MessageBoxHelper.Show(UIMessageOwner.ExcelOwner, ex.Message,
              "Error",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error
        )
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  '---------------------------------------------------------------------------
  ' Routine: btnUpdate_Click
  ' Purpose: Allow the user to change the path for an existing entry. Validation:
  '   - Requires an existing selection
  '   - Update cannot create a new DB: new path must exist and be a valid SQLite DB
  '   - Updates model and active path if necessary and persists config
  ' Parameters:
  '   sender - event sender
  '   e      - Click event args
  '---------------------------------------------------------------------------
  Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
    Try
      ' === Ensure a path is selected ===
      Dim oldPath As String = TryCast(lstDatabasePaths.SelectedItem, String)
      If String.IsNullOrWhiteSpace(oldPath) Then
        MessageBoxHelper.Show(UIMessageOwner.ExcelOwner, "Please select a database path to update.",
                          "No Selection",
                          MessageBoxButtons.OK,
                          MessageBoxIcon.Information)
        Exit Sub
      End If

      ' === Configure dialog with existing filename and folder ===
      Dim dlg As New OpenFileDialog()
      dlg.Title = "Update Database Path"
      dlg.Filter = "SQLite Database (*.db)|*.db|All Files (*.*)|*.*"
      dlg.DefaultExt = "db"
      dlg.CheckFileExists = True
      dlg.AddExtension = True
      dlg.FileName = IO.Path.GetFileName(oldPath)

      Dim oldFolder As String = IO.Path.GetDirectoryName(oldPath)
      If Not String.IsNullOrWhiteSpace(oldFolder) AndAlso IO.Directory.Exists(oldFolder) Then
        dlg.InitialDirectory = oldFolder
      End If

      ' === Show dialog ===
      If dlg.ShowDialog() <> DialogResult.OK Then Exit Sub

      Dim newPath As String = dlg.FileName

      ' === Update model to replace the database path ===
      _model.SelectedPath = oldPath
      _model.NewPath = newPath
      _model.PendingAction = DatabaseManagementAction.Update
      SaveDatabaseManagementModel(_model)

      ' === Refresh owner-draw ===
      lstDatabasePaths.Invalidate()

      MessageBoxHelper.Show(UIMessageOwner.ExcelOwner, "Database path updated successfully.",
                      "Updated",
                      MessageBoxButtons.OK,
                      MessageBoxIcon.Information)
    Catch ex As UserFriendlyException
      MessageBoxHelper.Show(UIMessageOwner.ExcelOwner, ex.Message,
              "Error",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error
        )
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  '---------------------------------------------------------------------------
  ' Routine: btnClose_Click
  ' Purpose: Close the dialog.
  ' Parameters:
  '   sender - event sender
  '   e      - Click event args
  '---------------------------------------------------------------------------
  Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
    Me.Close()
  End Sub
End Class