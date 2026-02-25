Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.Json
Imports System.Windows.Documents
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports ResourceManagement.EnumEntries

Public Class ExcelRuleDesigner
  Inherits UserControl

#Region "Pane - Initialisation Events"
  ' ==========================================================================================
  ' Routine:    OnEnter
  ' Purpose:    Reasserts keyboard focus to the task pane after Excel attempts to reclaim it
  '             during text entry or control activation. Ensures the pane remains the active
  '             keyboard root so that Tab navigation continues to function correctly.
  '
  ' Parameters:
  '   e        - Standard event arguments for focus‑enter events.
  '
  ' Returns:
  '   None
  '
  ' Notes:
  '   - Excel aggressively steals focus back to the grid after keystrokes in task panes.
  '   - This override uses BeginInvoke to run after Excel's internal focus logic completes.
  '   - Only affects focus when the task pane or its child controls are active; does not
  '     interfere with Excel grid navigation.
  ' ==========================================================================================
  Protected Overrides Sub OnEnter(e As EventArgs)
    MyBase.OnEnter(e)
    Me.BeginInvoke(Sub()
                     Me.Select()
                     Me.Focus()
                   End Sub)
  End Sub
  ' ==========================================================================================
  ' Routine:    OnGotFocus
  ' Purpose:    Reinforces focus ownership for the task pane when any child control receives
  '             focus. Prevents Excel from overriding the pane's keyboard loop and ensures
  '             consistent Tab and Shift+Tab navigation within the pane.
  '
  ' Parameters:
  '   e        - Standard event arguments for focus‑acquired events.
  '
  ' Returns:
  '   None
  '
  ' Notes:
  '   - Complements OnEnter by handling cases where Excel redirects focus after keystrokes.
  '   - BeginInvoke ensures the pane reclaims focus after Excel's message pump completes.
  '   - Does not affect Excel grid behaviour; only applies when the pane is active.
  ' ==========================================================================================
  Protected Overrides Sub OnGotFocus(e As EventArgs)
    MyBase.OnGotFocus(e)

    Me.BeginInvoke(Sub()
                     Me.Select()
                   End Sub)
  End Sub

  Private Class FieldTag
    ' Structural metadata (never null)
    Friend Property FilterID As String
    Friend Property ListTypeID As String
    Friend Property ViewName As String  ' Holds the canonical view name, which may differ from the actual SourceView for display purposes
    Public Property SourceView As String ' Holds the view name from which the source field is from, which may differ from the canonical view name
    Public Property SourceField As String ' Holds the field name from which the field is from, which may differ from the canonical field name
    Friend Property FieldName As String ' Holds the canonical field name, which may differ from the actual SourceField for display purposes
    Friend Property FieldID As String
    Friend Property DisplayName As String   ' <- UI text, set once
    ' Operator metadata (nullable)
    Friend Property FieldOperator As String
    Friend Property BooleanOperator As String
    Friend Property SlicingMode As String ' Used for time based fields that support time slicing
    ' Parentheses (never null)
    Friend Property OpenParenCount As Integer
    Friend Property CloseParenCount As Integer
    Friend Property ValueBinding As String
    ' Binding metadata (nullable)
    Friend Property RefType As String
    Friend Property LiteralValue As String
    Friend Property RefValue As String

  End Class

  Private Enum FilterExpressionRenderMode
    Rules
    Apply
  End Enum

  Private Const DEFAULT_FIELD_OPERATOR = "="
  Private Const DEFAULT_OPEN_PARENTHESES_COUNT = 0
  Private Const DEFAULT_CLOSE_PARENTHESES_COUNT = 0
  Private Const DEFAULT_BOOLEAN_OPERATOR = "AND"
  Private DEFAULT_VALUE_BINDING As String = ValueBinding.Parameter.ToString()
  Private Const UI_NAME = "Rule Designer"

  Private Enum Condition
    FieldOperater
    OpenParentheses
    CloseParentheses
    BooleanOperater
  End Enum

  Private _model As UIModelExcelRuleDesigner
  Private Const NEW_RULE_SENTINEL As String = "<New Rule…>"
  Private Const NEW_APPLY_SENTINEL As String = "<New Apply…>"
  Private _isInitialising As Boolean = False
  Private _isBinding As Boolean = False
  Private _suppressTreeAfterCheck As Boolean = False
  Private _fieldIndex As Dictionary(Of (String, String), TreeNode)

  Private Const DropButtonWidth As Integer = 18 ' Drop button width drawn for ListView

  Friend Class ComboRuleItem
    Public Property Display As String
    Public Property Value As String

    Public Overrides Function ToString() As String
      Return Display
    End Function
  End Class

  Friend Class ComboApplyItem
    Public Property Display As String
    Public Property Value As String

    Public Overrides Function ToString() As String
      Return Display
    End Function
  End Class

  Public Sub New()
    Try
      Me.AutoScaleMode = AutoScaleMode.None
      InitializeComponent()

      ' Apply manual DPI scaling to the UserControl
      ApplyDpiScalingToTaskPane(Me)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: ExcelRuleDesigner_Load
  ' Purpose:
  '   Initialise the Rule Designer by loading the full model and preparing both tabs.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' Notes:
  '   - Calls loader in first-load mode (model = Nothing).
  '   - Initialises Rules and Apply tabs for immediate user interaction.
  ' ==========================================================================================
  Private Sub ExcelRuleDesigner_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    Try

      Me.PerformLayout()

      ReloadModel()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' --- No cleanup required ---
    End Try
  End Sub

  Public Sub ReloadModel()
    Try
      ' 1. Reload model from DB into _model
      UILoaderSaverExcelRuleDesigner.LoadExcelRuleDesignerModel(_model)

      ' 2. Rebind UI from _model
      InitialiseRulesTab()
      InitialiseApplyTab()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: ResetPane
  ' Purpose:
  '   Restores the Rule Designer task pane to a clean, initial state. This routine clears the
  '   current model, removes all UI selections, resets TreeView check states, and reinitialises
  '   internal flags. It ensures that reopening the pane always starts from a deterministic,
  '   known-good baseline without residual state from previous rule edits.
  '
  ' Parameters:
  '   (None)
  '
  ' Returns:
  '   (Nothing)
  '
  ' Notes:
  '   - Must be called whenever the CustomTaskPane becomes hidden (VisibleStateChange event).
  '   - Temporarily detaches tvRuleFields.AfterCheck to prevent unintended event firing while
  '     programmatically clearing node check states.
  '   - Does not dispose the pane; Excel-DNA owns the CustomTaskPane lifecycle.
  '   - Safe to call multiple times; all operations are idempotent.
  ' ==========================================================================================
  Public Sub ResetPane()
    ' Clear your model
    _model = Nothing

    ' Clear UI
    txtRuleName.Text = ""
    lvRuleValues.Items.Clear()
    lvRuleFilters.Items.Clear()
    optSingle.Checked = False
    optList.Checked = False
    optRange.Checked = False
    txtRuleFilterExpression.Text = ""

    ' Reset TreeView
    RemoveHandler tvRuleFields.AfterCheck, AddressOf tvRuleFields_AfterCheck
    For Each n As TreeNode In tvRuleFields.Nodes
      n.Checked = False
    Next
    AddHandler tvRuleFields.AfterCheck, AddressOf tvRuleFields_AfterCheck

  End Sub

  ' ==========================================================================================
  ' Routine: BuildContextMenuForListView
  ' Purpose:
  '   Construct the context menu for a specific ListViewItem, including movement,
  '   operator selection (Filters only), and removal.
  ' Parameters:
  '   lv   - the ListView containing the item
  '   item - the ListViewItem for which the menu is being built
  ' Returns:
  '   ContextMenuStrip - the fully constructed context menu
  ' Notes:
  '   - Operator submenu is only added for the Filters ListView.
  '   - Movement options are always included.
  ' ==========================================================================================
  Private Function BuildContextMenuForListView(lv As System.Windows.Forms.ListView, item As ListViewItem) As ContextMenuStrip
    Dim menu As New ContextMenuStrip()
    Dim tag As FieldTag = TryCast(item.Tag, FieldTag)

    If lv Is lvRuleValues Or lv Is lvRuleFilters Then
      ' ------------------------------
      ' Section 1: Move operations
      ' ------------------------------

      menu.Items.Add("Move Up", Nothing, Sub() MoveItem(lv, item, -1))
      menu.Items.Add("Move Down", Nothing, Sub() MoveItem(lv, item, +1))
      menu.Items.Add("Move to Beginning", Nothing, Sub() MoveItemTo(lv, item, 0))
      menu.Items.Add("Move to End", Nothing, Sub() MoveItemTo(lv, item, lv.Items.Count - 1))

      menu.Items.Add(New ToolStripSeparator())

      ' ------------------------------
      ' Section 2: Move between lists
      ' ------------------------------
      If lv Is lvRuleValues Then
        menu.Items.Add("Move to Filters", Nothing,
                                          Sub()
                                            MoveItemBetweenLists(item, lvRuleValues, lvRuleFilters)
                                            UpdateRuleTypeUI()   ' values / Filters list changed → recompute rule type options
                                          End Sub)
      ElseIf lv Is lvRuleFilters Then
        menu.Items.Add("Move to Values", Nothing,
                                          Sub()
                                            MoveItemBetweenLists(item, lvRuleFilters, lvRuleValues)
                                            UpdateRuleTypeUI()   ' values / Filters list changed → recompute rule type options
                                          End Sub)
      End If

      If lv Is lvRuleFilters Then
        Dim fieldInfo As ExcelRuleViewMapField = _model.ViewMapHelper.GetField(tag.ViewName, tag.FieldName)
        ' ------------------------------
        ' Section 3: Operator selection (Filters only)
        ' ------------------------------
        menu.Items.Add(New ToolStripSeparator())

        Dim allowedOps As List(Of String) = fieldInfo?.AllowedOperators

        Dim opMenu As New ToolStripMenuItem("Select Operator...")

        For Each op In _model.AvailableOperators
          Dim enabled As Boolean = (allowedOps Is Nothing OrElse allowedOps.Contains(op))

          Dim opItem As New ToolStripMenuItem(op) With {
              .Checked = (tag.FieldOperator = op),
              .Enabled = enabled
          }

          AddHandler opItem.Click, Sub()
                                     If enabled Then
                                       SetFilterCondition(Condition.FieldOperater, item, op)
                                     End If
                                   End Sub

          opMenu.DropDownItems.Add(opItem)
        Next

        menu.Items.Add(opMenu)

        ' ------------------------------
        ' Section 4: Open parentheses selection (Filters only)
        ' ------------------------------
        Dim opaMenu As New ToolStripMenuItem("Select Open Parentheses...")
        For Each opa In _model.AvailableOpenParentheses
          Dim opaItem As New ToolStripMenuItem(opa) With {
              .Checked = (tag.OpenParenCount = opa.Length)
          }
          AddHandler opaItem.Click, Sub()
                                      SetFilterCondition(Condition.OpenParentheses, item, opa)
                                    End Sub
          opaMenu.DropDownItems.Add(opaItem)
        Next
        menu.Items.Add(opaMenu)

        ' ------------------------------
        ' Section 5: Close parentheses selection (Filters only)
        ' ------------------------------
        Dim cpaMenu As New ToolStripMenuItem("Select Close Parentheses...")
        For Each cpa In _model.AvailableCloseParentheses
          Dim cpaItem As New ToolStripMenuItem(cpa) With {
              .Checked = (tag.CloseParenCount = cpa.Length)
          }
          AddHandler cpaItem.Click, Sub()
                                      SetFilterCondition(Condition.CloseParentheses, item, cpa)
                                    End Sub
          cpaMenu.DropDownItems.Add(cpaItem)
        Next
        menu.Items.Add(cpaMenu)

        ' ------------------------------
        ' Section 6: Boolean Operator selection (Filters only)
        ' ------------------------------

        Dim bopMenu As New ToolStripMenuItem("Select Boolean  Operator...")
        For Each bop In _model.AvailableBooleanOperators
          Dim bopItem As New ToolStripMenuItem(bop) With {
              .Checked = (tag.BooleanOperator = bop)
          }
          AddHandler bopItem.Click, Sub()
                                      SetFilterCondition(Condition.BooleanOperater, item, bop)
                                    End Sub
          bopMenu.DropDownItems.Add(bopItem)
        Next
        menu.Items.Add(bopMenu)
        ' ------------------------------
        ' Section 7: If field supports slicing then Slicing Mode
        ' ---------------------
        If fieldInfo.SupportsSlicing Then
          menu.Items.Add(New ToolStripSeparator())
          Dim sliceMenu As New ToolStripMenuItem("Select Slicing Mode...")
          Dim tooltip As New StringBuilder()
          tooltip.AppendLine("Slicing mode determines how time-based fields are evaluated.")

          For Each mode In fieldInfo.SlicingOptions
            Dim desc As String = Nothing
            If fieldInfo.SlicingDescriptions IsNot Nothing AndAlso fieldInfo.SlicingDescriptions.TryGetValue(mode, desc) Then
              tooltip.AppendLine($"{mode} - {desc}")
            Else
              tooltip.AppendLine(mode)
            End If

            Dim modeItem As New ToolStripMenuItem(mode) With {
                .Checked = (tag.SlicingMode = mode)
            }

            AddHandler modeItem.Click, Sub()
                                         tag.SlicingMode = mode
                                         ' Update checkmarks
                                         For Each mi As ToolStripMenuItem In sliceMenu.DropDownItems
                                           mi.Checked = (mi Is modeItem)
                                         Next
                                       End Sub

            sliceMenu.DropDownItems.Add(modeItem)
          Next
          sliceMenu.ToolTipText = tooltip.ToString().Trim()
          menu.Items.Add(sliceMenu)
        End If
        ' ------------------------------
        ' Section 8: Value Binding Mode (Filters only)
        ' ------------------------------
        menu.Items.Add(New ToolStripSeparator())
        Dim vbMenu As New ToolStripMenuItem("Value Binding...")

        For Each entry As BindingItem(Of ValueBinding) In ValueBindingMap.BindingList()
          Dim enumValue = entry.EnumValue
          Dim display = entry.Display

          Dim vbItem As New ToolStripMenuItem(display) With {
              .Tag = enumValue,
              .Checked = (tag.ValueBinding = enumValue.ToString())
          }

          AddHandler vbItem.Click,
              Sub()
                tag.ValueBinding = enumValue.ToString()
                tag.LiteralValue = Nothing   ' reset literal when switching modes
                ' Update checkmarks
                For Each mi As ToolStripMenuItem In vbMenu.DropDownItems
                  mi.Checked = (mi Is vbItem)
                Next
                UpdateRuleExpressionDisplay()
              End Sub

          vbMenu.DropDownItems.Add(vbItem)
        Next

        menu.Items.Add(vbMenu)

        ' ------------------------------
        ' Section 9: Value Selection (enabled only when Rule-bound)
        ' ------------------------------
        Dim valMenu As New ToolStripMenuItem("Select Value...") With {
            .Enabled = (tag.ValueBinding = ValueBinding.Rule.ToString())
        }

        If tag.ValueBinding = ValueBinding.Rule.ToString() Then
          Dim allowed = fieldInfo.AllowedValues

          If allowed IsNot Nothing AndAlso allowed.Count > 0 Then
            ' Allowed values list
            For Each v In allowed
              Dim vItem As New ToolStripMenuItem(v) With {
                  .Checked = (tag.LiteralValue = v)
              }

              AddHandler vItem.Click,
                  Sub()
                    tag.LiteralValue = v
                    For Each mi As ToolStripMenuItem In valMenu.DropDownItems
                      mi.Checked = (mi Is vItem)
                    Next
                    UpdateRuleExpressionDisplay()
                  End Sub

              valMenu.DropDownItems.Add(vItem)
            Next

          Else
            ' No allowedValues → literal dialog
            AddHandler valMenu.Click,
                Sub()
                  SelectRuleLiteralValue(item)
                  UpdateRuleExpressionDisplay()
                End Sub
          End If
        End If

        menu.Items.Add(valMenu)
      End If

      ' ------------------------------
      ' Section 10: Remove
      ' ------------------------------
      menu.Items.Add(New ToolStripSeparator())
      menu.Items.Add("Remove Field", Nothing,
                   Sub()
                     Try
                       RemoveFieldEverywhere(tag)
                       UpdateRuleTypeUI()   ' values list changed → recompute rule type options

                     Catch ex As Exception
                       ErrorHandler.UnHandleError(ex)
                     End Try
                   End Sub)
    ElseIf lv Is lvApplyFilters Then
      ' Do not prompt for binding and value on Appy if Rule-bound 
      If tag.ValueBinding = ValueBinding.Rule.ToString() Then
        Dim valMenu As New ToolStripMenuItem("View Rule-bound Value...")
        AddHandler valMenu.Click,
            Sub()
              MessageBox.Show(Me, $"This field is bound to a rule literal with value: '{tag.LiteralValue}'", "Rule-bound Value", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Sub
        menu.Items.Add(valMenu)
      Else
        ' ------------------------------
        ' Section 1: Binding Type
        ' ------------------------------
        Dim binMenu As New ToolStripMenuItem("Select Binding Type...")
        For Each entry As BindingItemString In ExcelRefTypeMap.BindingListOfStrings()
          Dim value As String = entry.Value        ' "Literal", "Address", etc.
          Dim display As String = entry.Display    ' Friendly text

          Dim binItem As New ToolStripMenuItem(display) With {
              .Tag = value,
              .Checked = (Not String.IsNullOrWhiteSpace(tag.RefType) AndAlso
                          String.Equals(tag.RefType, value, StringComparison.OrdinalIgnoreCase))
          }

          AddHandler binItem.Click,
        Sub()
          ' Store the string value
          tag.RefType = value
          ' Update checkmarks
          For Each mi As ToolStripMenuItem In binMenu.DropDownItems
            mi.Checked = (mi Is binItem)
          Next
        End Sub
          binMenu.DropDownItems.Add(binItem)
        Next

        menu.Items.Add(binMenu)
        ' ------------------------------
        ' Section 2: Parameter based on Binding Type
        ' ------------------------------
        Dim hasBindingType As Boolean = Not String.IsNullOrWhiteSpace(tag.RefType)
        Dim selectedBindingType As String = ""

        If hasBindingType Then
          selectedBindingType = tag.RefType
        End If
        Dim paramMenu As New ToolStripMenuItem("Select Parameter...") With {
            .Enabled = hasBindingType
        }

        'AddHandler paramMenu.Click,
        '  Sub()
        If hasBindingType Then

          Select Case selectedBindingType
            Case ExcelRefType.Literal.ToString()
              AddHandler paramMenu.Click, Sub()
                                            SelectLiteralParameter(item)
                                            UpdateApplyExpressionDisplay()
                                          End Sub
            Case ExcelRefType.Address.ToString()
              AddHandler paramMenu.Click, Sub()
                                            SelectAbsoluteRangeParameter(item)
                                            UpdateApplyExpressionDisplay()
                                          End Sub
            Case ExcelRefType.Offset.ToString()
              AddHandler paramMenu.Click, Sub()
                                            SelectRelativeRangeParameter(item)
                                            UpdateApplyExpressionDisplay()
                                          End Sub
            Case ExcelRefType.Name.ToString()
              BuildNamedRangeSubmenu(paramMenu, item)
          End Select

          'End Sub
        End If
        menu.Items.Add(paramMenu)
      End If
    End If

    Return menu
  End Function

  ' ==========================================================================================
  ' Routine: SelectRuleLiteralValue
  ' Purpose: Small wrapper routine for invoking the existing literal selection dialog in the
  '          specific context of rule-bound. 
  ' Parameters:
  '   item  - The ListViewItem whose parameter column will be updated.
  ' Returns:
  '   none
  ' Notes:
  '   - This routine temporarily forces the FieldTag.RefType to Literal to reuse the existing
  '   dialog
  ' ==========================================================================================
  Private Sub SelectRuleLiteralValue(item As ListViewItem)
    Dim tag As FieldTag = CType(item.Tag, FieldTag)

    ' Save original RefType (parameter binding)
    Dim originalRefType As String = tag.RefType

    Try
      ' Force literal mode for validation
      tag.RefType = ExcelRefType.Literal.ToString()

      ' Reuse the existing literal dialog
      SelectLiteralParameter(item)

      ' The dialog stores the literal into tag.LiteralValue
      ' That’s exactly what we want for rule-bound literals

    Finally
      ' Restore original binding mode
      tag.RefType = originalRefType
    End Try
  End Sub

#End Region

#Region "Pane - Control Events"
  ' ==========================================================================================
  ' Routine:    ProcessCmdKey
  ' Purpose:    Intercepts Tab and Shift+Tab before Excel can steal the keyboard loop,
  '             enabling deterministic keyboard navigation between controls inside the
  '             task pane. Prevents Excel from reclaiming focus after text entry.
  '
  ' Parameters:
  '   msg      - The Windows message being processed.
  '   keyData  - The key combination pressed by the user.
  '
  ' Returns:
  '   Boolean  - True if the key was handled by the task pane (e.g., Tab navigation);
  '              otherwise defers to the base class for normal processing.
  '
  ' Notes:
  '   - Excel aggressively reclaims keyboard focus after keystrokes in task panes.
  '   - This override ensures Tab navigation remains internal to the pane.
  '   - Only fires when the task pane or its child controls have focus; does NOT
  '     affect Excel grid navigation or global keyboard behaviour.
  ' ==========================================================================================
  Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
    ' Intercept Tab before Excel gets it
    If keyData = Keys.Tab OrElse keyData = (Keys.Shift Or Keys.Tab) Then
      Dim forward As Boolean = (keyData And Keys.Shift) = Keys.None

      ' Move focus to next/previous control inside the pane
      Me.SelectNextControl(Me.ActiveControl, forward, True, True, True)
      Return True ' swallow Tab so Excel never sees it
    End If

    Return MyBase.ProcessCmdKey(msg, keyData)
  End Function

  ' ==========================================================================================
  ' Routine: lv_DrawSubItem
  ' Purpose:
  '   Custom‑draw the ListView subitems to render the left‑aligned drop button and field text.
  ' Parameters:
  '   sender - the ListView raising the event (lvRuleValues, lvRuleFilters, lvApplyFilters)
  '   e      - drawing arguments for the specific subitem
  ' Returns:
  '   None
  ' Notes:
  '   - Requires OwnerDraw = True on ListViews.
  '   - Draws the drop button inside column 0 for each row.
  '   - Must account for scrollbar presence when calculating button bounds.
  ' ==========================================================================================
  Private Sub lv_DrawSubItem(sender As Object, e As DrawListViewSubItemEventArgs) Handles lvRuleValues.DrawSubItem, lvRuleFilters.DrawSubItem, lvApplyFilters.DrawSubItem

    Dim lv As System.Windows.Forms.ListView = DirectCast(sender, System.Windows.Forms.ListView)

    ' --- Draw background ---
    e.DrawBackground()

    ' --- Draw field text, leaving space for button ---
    Dim textBounds As System.Drawing.Rectangle = e.Bounds
    textBounds.Width = lv.ClientSize.Width - DropButtonWidth - 6

    TextRenderer.DrawText(
        e.Graphics,
        e.SubItem.Text,
        lv.Font,
        textBounds,
        lv.ForeColor,
        TextFormatFlags.Left
    )

    ' --- Draw drop button on far right of ListView ---
    If e.ColumnIndex = 0 Then
      Dim btnRect As New System.Drawing.Rectangle(
            lv.ClientSize.Width - DropButtonWidth - 2,
            e.Bounds.Top + 2,
            DropButtonWidth,
            e.Bounds.Height - 4
        )
      ControlPaint.DrawComboButton(e.Graphics, btnRect, ButtonState.Normal)
    End If
  End Sub

  ' ==========================================================================================
  ' Routine: lv_Resize
  ' Purpose:
  '   Intercepts the resize events for the listviews to ensure they have a minimum size incase
  '   ClientSize width is zero when this fires.
  ' Parameters:
  '   sender - the ListView receiving the resize event
  '   e      - resize event arguments
  ' Returns:
  '   None
  ' Notes:
  '   - Supports both Values and Filters ListViews.
  ' ==========================================================================================
  Private Sub lv_Resize(sender As Object, e As EventArgs) Handles lvRuleValues.Resize, lvRuleFilters.Resize, lvApplyFilters.Resize

    Dim lv = DirectCast(sender, System.Windows.Forms.ListView)

    ' Ignore resize until columns exist
    If lv.Columns.Count = 0 Then Exit Sub

    ' Safe guard resize to zero
    Dim w = lv.ClientSize.Width
    If w < 200 Then w = 200
    lv.Columns(0).Width = w

    'lv.Columns(0).Width = lv.ClientSize.Width
  End Sub

  ' ==========================================================================================
  ' Routine: lv_MouseDown
  ' Purpose:
  '   Detect clicks on the per‑row drop button and open the context menu for that item.
  ' Parameters:
  '   sender - the ListView receiving the mouse event
  '   e      - mouse event arguments including click location
  ' Returns:
  '   None
  ' Notes:
  '   - Hit‑tests the drop‑button rectangle only; does not rely on item selection.
  '   - Supports both Values and Filters ListViews.
  ' ==========================================================================================
  Private Sub lv_MouseDown(sender As Object, e As MouseEventArgs) Handles lvRuleValues.MouseDown, lvRuleFilters.MouseDown, lvApplyFilters.MouseDown

    Dim lv As System.Windows.Forms.ListView = DirectCast(sender, System.Windows.Forms.ListView)
    Dim DropButtonWidth As Integer = 18

    Dim hit = lv.HitTest(e.Location)
    If hit.Item Is Nothing Then Exit Sub

    ' --- Compute drop button rectangle on far right ---
    Dim dropRect As New System.Drawing.Rectangle(
        lv.ClientSize.Width - DropButtonWidth,
        hit.Item.Bounds.Top,
        DropButtonWidth,
        hit.Item.Bounds.Height
    )

    ' --- If click is inside drop button → show menu ---
    If dropRect.Contains(e.Location) Then
      Dim menu = BuildContextMenuForListView(lv, hit.Item)

      ' Show menu to the right of the button
      Dim menuX = dropRect.Right
      Dim menuY = dropRect.Bottom
      menu.Show(lv, New System.Drawing.Point(menuX, menuY))
    End If
  End Sub

#End Region

#Region "Pane - Helpers"

  ' ==========================================================================================
  ' Routine:      BuildFilterExpression
  '
  ' Purpose:
  '   Constructs a deterministic, linear string representation of filter conditions from
  '   either lvRuleFilters (Rules tab) or lvApplyFilters (Apply tab).
  '
  '   - Handles parentheses, boolean operators, and mismatch detection.
  '   - For Rules tab: inserts "{param}" placeholder.
  '   - For Apply tab: inserts binding description from tag
  '
  ' Parameters:
  '   lv        - ListView containing filter rows.
  '   isApply   - Boolean. True = Apply tab (use binding info). False = Rules tab.
  '
  ' Returns:
  '   String - fully assembled expression, or empty string on error.
  '
  ' Notes:
  '   - No UI mutation.
  '   - No model mutation.
  '   - No side effects beyond string construction and error logging.
  ' ==========================================================================================
  Private Function BuildFilterExpression(lv As System.Windows.Forms.ListView, mode As FilterExpressionRenderMode) As String
    Dim sb As New System.Text.StringBuilder()

    Try
      If lv Is Nothing OrElse lv.Items.Count = 0 Then
        Return String.Empty
      End If

      Dim depth As Integer = 0
      Dim lastIndex As Integer = lv.Items.Count - 1

      For i As Integer = 0 To lastIndex
        Dim item As ListViewItem = lv.Items(i)
        Dim tag As FieldTag = CType(item.Tag, FieldTag)

        Dim displayName As String = item.SubItems(0).Text
        Dim fieldOperator As String = tag.FieldOperator

        Dim openParenCount As Integer
        Integer.TryParse(tag.OpenParenCount, openParenCount)

        Dim closeParenCount As Integer
        Integer.TryParse(tag.CloseParenCount, closeParenCount)

        Dim booleanOperator As String = tag.BooleanOperator

        ' -------------------------------
        ' Boolean operator rules
        ' -------------------------------
        Dim opMarker As String = Nothing

        If i = 0 Then
          If Not String.IsNullOrEmpty(booleanOperator) Then
            opMarker = $"<<NO_OP_ALLOWED:{booleanOperator}>>"
          End If
        Else
          If String.IsNullOrEmpty(booleanOperator) Then
            opMarker = "<<MISSING_OP>>"
          End If
        End If

        If opMarker IsNot Nothing Then
          sb.Append(opMarker)
        Else
          sb.Append(booleanOperator)
        End If

        sb.Append(" ")

        ' -------------------------------
        ' Opening parentheses
        ' -------------------------------
        If openParenCount > 0 Then
          sb.Append(New String("("c, openParenCount))
          sb.Append(" ")
        End If

        depth += openParenCount

        ' -------------------------------
        ' Field + operator + parameter/binding
        ' -------------------------------
        sb.Append(displayName)
        sb.Append(" ")
        sb.Append(fieldOperator)
        sb.Append(" ")

        Select Case mode
          Case FilterExpressionRenderMode.Rules
            Dim vb As String = If(tag.ValueBinding, "").Trim()

            ' 1. No binding selected → placeholder
            If vb.Length = 0 Then
              sb.Append("<<NO_BINDING_MODE>>")
              Exit Select
            End If

            ' 2. Rule-bound → show literal or placeholder
            If vb = ValueBinding.Rule.ToString() Then
              Dim lit As String = If(tag.LiteralValue, "").Trim()

              If lit.Length = 0 Then
                sb.Append("<<MISSING_RULE_VALUE>>")
              Else
                sb.Append(FormatLiteralForDisplay(lit))
              End If

              Exit Select
            End If

            ' 3. Parameter-bound → show generic param placeholder
            sb.Append("{param}")

          Case FilterExpressionRenderMode.Apply
            If tag.ValueBinding = ValueBinding.Rule.ToString() Then
              Dim lit As String = If(tag.LiteralValue, "").Trim()
              ' Valid rule based binding → render
              sb.Append("[")
              sb.Append("Rule Literal")
              sb.Append(": ")
              sb.Append(FormatLiteralForDisplay(lit))
              sb.Append("]")
              Exit Select
            End If
            If String.IsNullOrWhiteSpace(tag.RefType) Then
              sb.Append("<<NO_BINDING_DEFINED>>")
              Exit Select
            End If

            Dim refType As String = tag.RefType
            Dim bindType As String = ExcelRefTypeMap.DisplayFromString(refType)

            Dim bindValue As String
            Select Case refType
              Case ExcelRefType.Literal.ToString
                bindValue = tag.LiteralValue?.Trim()
              Case ExcelRefType.Address.ToString, ExcelRefType.Offset.ToString, ExcelRefType.Name.ToString
                bindValue = tag.RefValue?.Trim()
              Case Else
                bindValue = ""
            End Select

            ' Missing binding type AND value
            If String.IsNullOrEmpty(bindType) AndAlso String.IsNullOrEmpty(bindValue) Then
              sb.Append("<<NO_BINDING_DEFINED>>")
              Exit Select
            End If
            ' Missing binding type
            If String.IsNullOrEmpty(bindType) Then
              sb.Append("<<MISSING_BINDING_TYPE>>")
              Exit Select
            End If
            ' Missing binding value
            If String.IsNullOrEmpty(bindValue) Then
              sb.Append($"<<MISSING_BINDING_VALUE:{bindType}>>")
              Exit Select
            End If
            ' Valid binding → render normally
            sb.Append("[")
            sb.Append(bindType)
            sb.Append(": ")
            sb.Append(bindValue)
            sb.Append("]")

        End Select

        ' -------------------------------
        ' Closing parentheses
        ' -------------------------------
        If closeParenCount > 0 Then
          sb.Append(" ")
          sb.Append(New String(")"c, closeParenCount))
        End If

        depth -= closeParenCount

        If depth < 0 Then
          sb.Append(" <<PAREN_MISMATCH>> ")
          depth = 0
        End If

        sb.Append(" ")
      Next

      ' -------------------------------
      ' Final balance check
      ' -------------------------------
      If depth <> 0 Then
        sb.Append($"<<UNBALANCED:{depth}>>")
      End If

      Return sb.ToString()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return String.Empty
    End Try
  End Function

  ' ==========================================================================================
  ' Routine:      HighlightMarkers
  '
  ' Purpose:
  '   Scans the RichTextBox text for known mismatch markers and applies colour and font
  '   formatting to visually indicate issues in the rule filter expression.
  '
  ' Parameters:
  '   rtb  - RichTextBox
  '          The control whose text should be scanned and highlighted.
  '
  ' Returns:
  '   None
  '
  ' Notes:
  '   - Does not modify the underlying text content, only formatting.
  '   - Assumes the RichTextBox already contains the full expression.
  '   - Markers include: <<MISSING_OP>>, <<NO_OP_ON_LAST>>, <<NO_OP_ALLOWED>>,
  '                      <<PAREN_MISMATCH>>, <<UNBALANCED:n>>
  ' ==========================================================================================
  Private Sub HighlightMarkers(rtb As RichTextBox)
    Dim markers As New Dictionary(Of String, Color) From {
        {"<<MISSING_OP>>", Color.Red},
        {"<<NO_OP_ALLOWED", Color.OrangeRed},
        {"<<PAREN_MISMATCH>>", Color.Red},
        {"<<UNBALANCED", Color.Red},
        {"<<NO_BINDING_DEFINED>>", Color.Red},
        {"<<MISSING_BINDING_TYPE>>", Color.Red},
        {"<<MISSING_BINDING_VALUE", Color.Red},
        {"{param}", Color.Gray},
        {"<<MISSING_RULE_VALUE>>", Color.Red},
        {"<<NO_BINDING_MODE>>", Color.Red}
    }

    For Each kvp In markers
      Dim key = kvp.Key
      Dim colour = kvp.Value


      Dim start As Integer = 0

      While start < rtb.TextLength
        Dim idx = rtb.Text.IndexOf(key, start, StringComparison.OrdinalIgnoreCase)
        If idx = -1 Then Exit While
        Dim startChar As Char = rtb.Text(idx)
        Dim endIdx As Integer = -1

        If startChar = "<"c Then
          ' Marker like <<MISSING_OP>>
          endIdx = rtb.Text.IndexOf(">>", idx)
        ElseIf startChar = "{"c Then
          ' Placeholder like {param}
          endIdx = rtb.Text.IndexOf("}", idx)
        Else
          ' Not a marker we care about
          start = idx + 1
          Continue While
        End If

        If endIdx = -1 Then Exit While

        Dim length = (endIdx - idx) + 2

        rtb.[Select](idx, length)
        rtb.SelectionColor = colour
        rtb.SelectionFont = New System.Drawing.Font(rtb.Font, System.Drawing.FontStyle.Bold)

        start = idx + length
      End While
    Next

    ' Reset selection
    rtb.[Select](0, 0)
  End Sub

  Private Function FormatLiteralForDisplay(literal As String) As String
    If String.IsNullOrEmpty(literal) Then Return ""
    Dim b As Boolean
    Dim n As Double
    Dim d As Date
    If Boolean.TryParse(literal, b) Then
      ' Boolean → no quotes
      Return literal
    ElseIf Double.TryParse(literal, n) Then
      ' Numeric → no quotes
      Return literal
    ElseIf Date.TryParse(literal, d) Then
      ' Date → no quotes (screen only)
      Return literal
    Else
      ' Everything else → treat as text → quote it
      Return $"'{literal}'"
    End If
  End Function
#End Region

#Region "Rule - Initialisaton Events"
  ' ==========================================================================================
  ' Routine: InitialiseRulesTab
  ' Purpose:
  '   Prepare the Rules tab UI by binding rule list, clearing detail, and populating fields tree.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' Notes:
  '   - Called on form load and after save/delete operations.
  ' ==========================================================================================
  Private Sub InitialiseRulesTab()
    Try

      ' --- Ensure TreeView behaves as field selector with checkboxes ---
      tvRuleFields.CheckBoxes = True
      tvRuleFields.ShowLines = True
      tvRuleFields.ShowPlusMinus = True
      tvRuleFields.AllowDrop = True ' allow drag and drop

      ' --- Populate rule ComboBox ---
      cmbRuleNames.Items.Clear()
      cmbRuleNames.Items.Add(NEW_RULE_SENTINEL)

      'For Each r In _model.Rules
      '  cmbRuleNames.Items.Add(r.RuleName)
      'Next
      For Each r In _model.Rules
        Dim item As New ComboRuleItem With {
          .Display = r.RuleName,
          .Value = r.RuleID
      }
        cmbRuleNames.Items.Add(item)
      Next

      cmbRuleNames.DropDownStyle = ComboBoxStyle.DropDownList
      If cmbRuleNames.Items.Count > 0 Then cmbRuleNames.SelectedIndex = 0

      ' --- Clear detail UI ---
      txtRuleName.Text = ""
      lvRuleValues.Items.Clear()
      lvRuleFilters.Items.Clear()
      lvRuleValues.OwnerDraw = True ' allow custom context menu drawing later
      lvRuleFilters.OwnerDraw = True ' allow custom context menu drawing later
      lvRuleValues.AllowDrop = True ' allow drag and drop
      lvRuleFilters.AllowDrop = True ' allow drag and drop
      optSingle.Checked = False
      optList.Checked = False
      optRange.Checked = False
      txtRuleFilterExpression.Text = ""

      ' --- Populate available fields tree ---
      PopulateAvailableFieldsTree()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' --- No cleanup required ---
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: PopulateAvailableFieldsTree
  ' Purpose:
  '   Populate the TreeView with views and their selectable fields for the Rule Designer.
  '   - Normal views: display user-facing field DisplayName and hide internal ID/lookup fields.
  '   - Special-case "vwDim_List": render ListType nodes (ListTypeName) with child item nodes
  '     (ItemName). Tree nodes store internal identifiers in Tag/Name for runtime mapping.
  '
  ' Parameters:
  '   None
  '
  ' Returns:
  '   None
  '
  ' Notes:
  '   - Uses model.ViewMapHelper plus model.ListTypes and model.ListItemsByType (populated by loader).
  '   - No database access here; all DB reads happen in the loader (RecordLoader).
  '   - Internal fields (Role = Key / ForeignKey / LookupType) are never shown to the user.
  '   - Builds _fieldIndex for fast lookup of (viewName, field Identifier) → TreeNode.
  '   - TreeNode.Tag holds the internal identifier (View name, field name and optional fieldID
  '     and optional item id); TreeNode.Name is set to allow UncheckFieldInTree matching.
  ' ==========================================================================================  
  Private Sub PopulateAvailableFieldsTree()
    Try
      tvRuleFields.Nodes.Clear()

      ' --- Build tree: View → Fields (special-case vwDim_List to show list types → items) ---
      For Each viewName In _model.AvailableViews

        ' Get view metadata
        Dim view = _model.ViewMapHelper.GetView(viewName)
        If view Is Nothing Then Continue For

        ' --- Add view node (DisplayName shown, Name stored) ---
        Dim viewNode As TreeNode = tvRuleFields.Nodes.Add(view.DisplayName)
        viewNode.Tag = view.Name   ' store internal view name

        If String.Equals(view.Name, "vwDim_List", StringComparison.OrdinalIgnoreCase) Then
          ' --- One node per ListType ---
          If _model.ListTypes IsNot Nothing Then
            For Each lt In _model.ListTypes
              ' Internal field is ALWAYS ListItemName
              Dim node As TreeNode = viewNode.Nodes.Add(lt.Name)
              node.Tag = New FieldTag With {
                  .ListTypeID = lt.Id,                     ' ListTypeID
                  .ViewName = "vwDim_List",
                  .FieldName = "ListItemName",      ' physical field name
                  .FieldID = Nothing,
                  .SourceView = "vwDim_List",
                  .SourceField = "ListItemName",
                  .FilterID = Nothing,
                  .DisplayName = lt.Name,
                  .SlicingMode = Nothing,
                  .RefType = Nothing, ' APPLY-TIME MUST ALWAYS BE NOTHING HERE
                  .LiteralValue = Nothing,
                  .RefValue = Nothing
              }
              node.Name = lt.Id                   ' for UncheckFieldInTree
            Next
          End If
          Continue For
        Else
          ' --- Normal view → show non-internal fields as before ---
          For Each f In view.Fields
            If String.Equals(f.Role, "Key", StringComparison.OrdinalIgnoreCase) _
             OrElse String.Equals(f.Role, "ForeignKey", StringComparison.OrdinalIgnoreCase) _
             OrElse String.Equals(f.Role, "LookupType", StringComparison.OrdinalIgnoreCase) Then
              Continue For
            End If

            Dim fieldDisplay As String = If(String.IsNullOrEmpty(f.DisplayName), f.Name, f.DisplayName)

            ' DisplayName shown, Name stored
            Dim fieldNode As TreeNode = viewNode.Nodes.Add(fieldDisplay)
            fieldNode.Tag = New FieldTag With {
              .ViewName = view.Name,
              .FieldName = f.Name,  'store internal field name
              .FieldID = f.FieldID,
              .SourceView = f.SourceView,
              .SourceField = f.SourceField,
              .ListTypeID = Nothing,
              .FilterID = Nothing,
              .DisplayName = fieldDisplay,
              .SlicingMode = Nothing,
              .RefType = Nothing, ' APPLY-TIME MUST ALWAYS BE NOTHING HERE
              .LiteralValue = Nothing,
              .RefValue = Nothing
            }
            fieldNode.Name = f.Name
          Next
        End If
      Next

      ' --- Build a dictionary to find fields to check quickly ---
      _fieldIndex = New Dictionary(Of (String, String), TreeNode)(New TupleStringComparer())

      For Each viewNode As TreeNode In tvRuleFields.Nodes
        For Each childNode As TreeNode In viewNode.Nodes
          Dim ft As FieldTag = TryCast(childNode.Tag, FieldTag)
          If ft IsNot Nothing Then
            Dim key = (ft.ViewName, ft.FieldName & "|" & If(ft.FieldID, "") & "|" & If(ft.ListTypeID, ""))
            _fieldIndex(key) = childNode
          End If
        Next
      Next

      tvRuleFields.CollapseAll()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' --- No cleanup required ---
    End Try
  End Sub

  Private Class TupleStringComparer
    Implements IEqualityComparer(Of (String, String))

    Public Overloads Function Equals(x As (String, String), y As (String, String)) As Boolean _
        Implements IEqualityComparer(Of (String, String)).Equals

      Return StringComparer.OrdinalIgnoreCase.Equals(x.Item1, y.Item1) AndAlso
               StringComparer.OrdinalIgnoreCase.Equals(x.Item2, y.Item2)
    End Function

    Public Overloads Function GetHashCode(obj As (String, String)) As Integer _
        Implements IEqualityComparer(Of (String, String)).GetHashCode

      Dim h1 = StringComparer.OrdinalIgnoreCase.GetHashCode(obj.Item1)
      Dim h2 = StringComparer.OrdinalIgnoreCase.GetHashCode(obj.Item2)
      Return h1 Xor h2
    End Function
  End Class


  ' ==========================================================================================
  ' Routine: BindRuleDetail
  ' Purpose:
  '   Bind the selected rule's detail (name, values, filters) to the UI controls.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' Notes:
  '   - Called after selecting an existing rule.
  ' ==========================================================================================
  Private Sub BindRuleDetail()
    Try
      _isBinding = True

      ' --- Rule name ---
      txtRuleName.Text = _model.RuleDetail.RuleName

      ' --- Reset tree and check selected fields ---
      ResetTree()

      ' --- Values ListView and TreeView ---
      lvRuleValues.Items.Clear()
      For Each v In _model.RuleDetail.SelectedValues
        ' --- Check in tree ---
        Dim node As TreeNode = Nothing
        Dim key = (v.View, v.Field & "|" & If(v.FieldID, "") & "|" & v.ListTypeID)
        If _fieldIndex.TryGetValue(key, node) Then node.Checked = True
        ' --- Add to Filters ListView ---
        Dim tag As New FieldTag With {
          .ListTypeID = v.ListTypeID,
          .ViewName = v.View,
          .SourceView = v.SourceView,
          .SourceField = v.SourceField,
          .FieldName = v.Field,
          .FieldID = v.FieldID
        }
        ' --- Resolve display text for ListViewItem (minimal, localised fix) ---
        Dim displayText As String
        If v.View = "vwDim_List" AndAlso Not String.IsNullOrEmpty(v.ListTypeID) Then
          ' List-type value: use ListType name from model.ListTypes
          Dim lt = _model.ListTypes?.FirstOrDefault(Function(x) x.Id = v.ListTypeID)
          displayText = If(lt IsNot Nothing AndAlso Not String.IsNullOrEmpty(lt.Name),
                         lt.Name,
                         v.Field)   ' fallback
        Else
          ' Normal field: use field name as before (or later ViewMapHelper if you choose)
          displayText = v.Field
        End If
        tag.DisplayName = displayText
        Dim item As New ListViewItem(displayText)
        item.Tag = tag
        lvRuleValues.Items.Add(item)
      Next

      ' --- Rule Type ---
      Select Case _model.RuleDetail.RuleType
        Case ExcelRuleType.SingleValue.ToString()
          optSingle.Checked = True
        Case ExcelRuleType.ListOfValues.ToString()
          optList.Checked = True
        Case ExcelRuleType.RangeOfValues.ToString()
          optRange.Checked = True
      End Select

      ' --- Filters ListView and TreeView---
      lvRuleFilters.Items.Clear()
      For Each f In _model.RuleDetail.Filters
        ' --- Check in tree ---
        Dim node As TreeNode = Nothing
        Dim key = (f.View, f.Field)
        If _fieldIndex.TryGetValue(key, node) Then node.Checked = True
        ' --- Add to Filters ListView ---
        Dim tag As New FieldTag With {
          .FilterID = f.FilterID,
          .ListTypeID = f.ListTypeID,
          .ViewName = f.View,
          .SourceView = f.SourceView,
          .SourceField = f.SourceField,
          .FieldName = f.Field,
          .FieldID = f.FieldID,
          .FieldOperator = f.FieldOperator,
          .BooleanOperator = f.BooleanOperator,
          .SlicingMode = f.SlicingMode,
          .OpenParenCount = f.OpenParenCount,
          .CloseParenCount = f.CloseParenCount,
          .ValueBinding = f.ValueBinding,
          .LiteralValue = If(String.IsNullOrWhiteSpace(f.LiteralValue), Nothing, f.LiteralValue.Trim())
        }
        ' --- Resolve display text for ListViewItem (same minimal rule as values) ---
        Dim displayText As String

        If f.View = "vwDim_List" AndAlso Not String.IsNullOrEmpty(f.ListTypeID) Then
          Dim lt = _model.ListTypes?.FirstOrDefault(Function(x) x.Id = f.ListTypeID)
          displayText = If(lt IsNot Nothing AndAlso Not String.IsNullOrEmpty(lt.Name),
                         lt.Name,
                         f.Field)
        Else
          displayText = f.Field
        End If
        tag.DisplayName = displayText
        Dim item As New ListViewItem(displayText)
        item.Tag = tag

        lvRuleFilters.Items.Add(item)
      Next

      ' --- Set the Rule Filter Expression for display ---
      UpdateRuleExpressionDisplay()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    Finally
      _isBinding = False
    End Try
  End Sub


  ' ==========================================================================================
  ' Routine: ResetTree
  ' Purpose:
  '   Clear all the checked items in the TreeView (used when loading a new rule) and collapse the
  '   tree.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' Notes:
  '   - Must handle nested nodes if the TreeView groups fields by view.
  ' ==========================================================================================
  Private Sub ResetTree()
    _isInitialising = True   ' suppress AfterCheck events

    For Each viewNode As TreeNode In tvRuleFields.Nodes
      For Each fieldNode As TreeNode In viewNode.Nodes
        fieldNode.Checked = False
      Next
      viewNode.Collapse()
    Next

    _isInitialising = False
  End Sub

#End Region

#Region "Rule - Control Events"

  ' ==========================================================================================
  ' Routine: tvRuleFields_AfterCheck
  ' Purpose:
  '   Respond to user checking/unchecking fields in the TreeView.
  '   Checked fields are added to Values; unchecked are removed.
  ' Parameters:
  '   sender - event sender
  '   e      - TreeViewEventArgs for the changed node
  ' Returns:
  '   None
  ' Notes:
  '   - Ignores view-level nodes; only acts on field-level nodes.
  '   - Uses internal names (Tag) but shows display names in ListView.
  ' ==========================================================================================
  Private Sub tvRuleFields_AfterCheck(sender As Object, e As TreeViewEventArgs) Handles tvRuleFields.AfterCheck

    Try
      If _isInitialising Then Exit Sub
      If _isBinding Then Exit Sub
      If _suppressTreeAfterCheck Then Exit Sub

      ' --- Safe guard the model hasn't been cleared ---
      If _model Is Nothing Then
        UILoaderSaverExcelRuleDesigner.LoadExcelRuleDesignerModel(_model)
        If _model Is Nothing Then Exit Sub
      End If

      ' Ignore view nodes (no parent means top-level view)
      If e.Node.Parent Is Nothing Then Exit Sub

      ' --- ALWAYS use FieldTag for selectable nodes ---
      Dim ft As FieldTag = TryCast(e.Node.Tag, FieldTag)
      If ft Is Nothing Then Exit Sub


      If e.Node.Checked Then
        AddFieldEverywhere(ft, lvRuleValues)
        UpdateRuleTypeUI()
      Else
        ' --- Suppress tree uncheck because the tree is the source of truth here ---
        _suppressTreeAfterCheck = True
        RemoveFieldEverywhere(ft)
        _suppressTreeAfterCheck = False
        UpdateRuleTypeUI()
      End If

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' --- No cleanup required ---
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: cmbRuleNames_SelectedIndexChanged
  ' Purpose:
  '   Handle user selection of an existing rule or the <New Rule…> sentinel.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' Notes:
  '   - Selecting <New Rule…> clears detail and sets PendingAction = Add.
  '   - Selecting an existing rule loads detail via loader refresh path.
  ' ==========================================================================================
  Private Sub cmbRuleNames_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbRuleNames.SelectedIndexChanged

    Try
      ' --- Safe guard the model hasn't been cleared ---
      If _model Is Nothing Then
        UILoaderSaverExcelRuleDesigner.LoadExcelRuleDesignerModel(_model)
        If _model Is Nothing Then Exit Sub
      End If

      ' --- Clear detail UI ---
      txtRuleName.Text = ""
      ResetTree()
      lvRuleValues.Items.Clear()
      lvRuleFilters.Items.Clear()
      optSingle.Checked = False
      optList.Checked = False
      optRange.Checked = False
      txtRuleFilterExpression.Text = ""

      Dim selectedObj As Object = cmbRuleNames.SelectedItem

      ' --- Sentinel path (new rule) ---
      If TypeOf selectedObj Is String AndAlso
           CStr(selectedObj) = NEW_RULE_SENTINEL Then
        _model.SelectedRule = Nothing
        _model.PendingAction = ExcelRuleDesignerAction.Add
        Return
      End If

      ' --- Existing rule path ---
      Dim item As ComboRuleItem = TryCast(selectedObj, ComboRuleItem)
      If item Is Nothing Then Exit Sub   ' Defensive
      Dim ruleId As String = item.Value
      ' Find the rule by RuleID (NOT by name anymore)
      Dim selected = _model.Rules.FirstOrDefault(Function(r) r.RuleID = ruleId)

      _model.SelectedRule = selected
      _model.PendingAction = ExcelRuleDesignerAction.Update

      ' --- Reload model (refresh path) ---
      UILoaderSaverExcelRuleDesigner.LoadExcelRuleDesignerModel(_model)

      ' --- Bind detail to UI ---
      BindRuleDetail()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' --- No cleanup required ---
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: UpdateRuleTypeUI
  ' Purpose:
  '   Enforces valid RuleType selections based on the number of value columns defined in
  '   lvRuleValues. Ensures the UI always reflects the only legal combinations:
  '     - 0 items  → only "List" enabled (rule incomplete)
  '     - 1 item   → "Single" and "List" enabled, "Range" disabled
  '     - 2+ items → "Range" forced and enabled, others disabled
  '
  ' Parameters:
  '   None
  '
  ' Returns:
  '   None
  '
  ' Notes:
  '   - This routine must be called after any modification to lvRuleValues:
  '       AddFieldEveryWhere, RemoveFieldEverywhere, MoveItemBetweenLists, ReloadModel, ResetPane, etc.
  '   - This routine does NOT select a RuleType for the user except when logically forced
  '     (2+ values → Range). In all other cases the user must explicitly choose.
  '   - Saving a rule MUST call GetRuleTypeFromUI(), which will throw if no valid option
  '     is selected.
  ' ==========================================================================================
  Private Sub UpdateRuleTypeUI()

    Dim count As Integer = lvRuleValues.Items.Count

    ' ----------------------------------------------------------------------
    ' CASE 1: Two or more value columns → must be a Range rule
    ' ----------------------------------------------------------------------
    If count >= 2 Then

      ' Force Range
      optRange.Checked = True
      optRange.Enabled = True

      ' Disable invalid options
      optSingle.Enabled = False
      optList.Enabled = False

      Exit Sub
    End If

    ' ----------------------------------------------------------------------
    ' CASE 2: Exactly one value column → Single or List allowed
    ' ----------------------------------------------------------------------
    If count = 1 Then

      ' Range is not valid with only one column
      optRange.Checked = False
      optRange.Enabled = False

      ' User must choose between Single or List
      optSingle.Enabled = True
      optList.Enabled = True

      Exit Sub
    End If

    ' ----------------------------------------------------------------------
    ' CASE 3: No value columns → rule incomplete
    ' ----------------------------------------------------------------------
    ' Only "List" is logically possible (but rule cannot be saved yet)
    optRange.Enabled = False
    optSingle.Enabled = False

    optList.Enabled = True
    optList.Checked = True   ' UI needs a visible state, but save will still block

  End Sub

  ' ==========================================================================================
  ' Routine: GetRuleTypeFromUI
  ' Purpose:
  '   Determines the rule's return type based on the user's explicit UI selection.
  '   This routine enforces that a rule cannot be saved unless the user has chosen
  '   exactly one of: Single value, List of values, or Range of values.
  '
  ' Parameters:
  '   None
  '
  ' Returns:
  '   String - One of: "single", "list", "range".
  '
  ' Notes:
  '   - No fallback is permitted. If no option is selected, this routine throws
  '     an InvalidOperationException so the caller can block the save and prompt
  '     the user.
  '   - UI state constraints (enabled/disabled options) must already be enforced
  '     by UpdateRuleTypeUI().
  ' ==========================================================================================
  Private Function GetRuleTypeFromUI() As String

    ' --- Single value selected ---
    If optSingle.Checked Then
      ' --- Make sure only one value selected in lvRuleValues ---
      If lvRuleValues.Items.Count <> 1 Then
        Throw New InvalidOperationException(
            "Please select one value for Single value rule."
        )
      End If
      Return ExcelRuleType.SingleValue.ToString() ' ExcelRuleTypeMap.Display(ExcelRuleType.SingleValue)

    End If

    ' --- List of values selected ---
    If optList.Checked Then
      ' --- Make sure only one value selected in lvRuleValues ---
      If lvRuleValues.Items.Count <> 1 Then
        Throw New InvalidOperationException(
            "Please select one value for List value rule."
        )
      End If
      Return ExcelRuleType.ListOfValues.ToString()  ' ExcelRuleTypeMap.Display(ExcelRuleType.ListOfValues)
    End If

    ' --- Range of values selected ---
    If optRange.Checked Then
      ' --- Make sure at least one value selected in lvRuleValues ---
      If lvRuleValues.Items.Count < 1 Then
        Throw New InvalidOperationException(
            "Please select at least one value for Range value rule."
        )
      End If
      Return ExcelRuleType.RangeOfValues.ToString() 'ExcelRuleTypeMap.Display(ExcelRuleType.RangeOfValues)
    End If

    ' --- No selection: this is an error and must block saving ---
    Throw New InvalidOperationException(
        "No rule type selected. Please choose Single value, List of values, or Range of values."
    )

  End Function

  ' ==========================================================================================
  ' Routine: btnRuleSave_Click
  ' Purpose:
  '   Save the current rule (Add or Update) using the loader/saver.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' Notes:
  '   - Maps UI → model.ActionRule and ActionRuleDetail.
  '   - Reloads model and reinitialises UI after save.
  ' ==========================================================================================
  Private Sub btnRuleSave_Click(sender As Object, e As EventArgs) Handles btnRuleSave.Click
    Try

      ' --- Validate rule structure ---
      If Not ValidateRuleBeforeSave() Then Exit Sub

      ' --- Safeguard: ensure model exists ---
      If _model Is Nothing Then
        UILoaderSaverExcelRuleDesigner.LoadExcelRuleDesignerModel(_model)
        If _model Is Nothing Then Exit Sub
      End If

      ' --- Always rebuild ActionRule from scratch ---
      _model.ActionRule = New UIExcelRuleDesignerRuleRow()

      ' --- Preserve RuleID on Update ---
      If _model.PendingAction = ExcelRuleDesignerAction.Update AndAlso
           _model.SelectedRule IsNot Nothing Then

        _model.ActionRule.RuleID = _model.SelectedRule.RuleID
      End If

      ' --- Validate and assign RuleName ---
      Dim name As String = txtRuleName.Text.Trim()
      If name.Length = 0 Then
        MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                              "Rule name cannot be empty.",
                            UI_NAME,
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning)
        Exit Sub
      End If
      _model.ActionRule.RuleName = name

      ' --- Determine RuleType ---
      Try
        _model.ActionRule.RuleType = GetRuleTypeFromUI()
      Catch ex As InvalidOperationException
        MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                                  ex.Message,
                                  UI_NAME,
                                  MessageBoxButtons.OK,
                                  MessageBoxIcon.Warning)
        Exit Sub
      End Try

      ' --- Rebuild ActionRuleDetail from scratch ---
      MapRuleDetailUIToModel()

      ' --- Commit ---
      UILoaderSaverExcelRuleDesigner.SavePendingRuleAction(_model)

      ' --- Reload model and rebind ---
      ReloadModel()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' --- No cleanup required ---
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: btnRuleDelete_Click
  ' Purpose:
  '   Delete the currently selected rule.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' Notes:
  '   - Performs soft delete via loader/saver.
  ' ==========================================================================================
  Private Sub btnRuleDelete_Click(sender As Object, e As EventArgs) Handles btnRuleDelete.Click
    Try
      ' --- Safe guard the model hasn't been cleared ---
      If _model Is Nothing Then
        UILoaderSaverExcelRuleDesigner.LoadExcelRuleDesignerModel(_model)
        If _model Is Nothing Then Exit Sub
      End If

      If _model.SelectedRule Is Nothing Then Exit Sub
      ' --- Always rebuild ActionRule from scratch ---
      _model.ActionRule = New UIExcelRuleDesignerRuleRow()
      ' --- Preserve RuleID on Delete ---
      _model.PendingAction = ExcelRuleDesignerAction.Delete
      _model.ActionRule.RuleID = _model.SelectedRule.RuleID

      UILoaderSaverExcelRuleDesigner.SavePendingRuleAction(_model)
      UILoaderSaverExcelRuleDesigner.LoadExcelRuleDesignerModel(_model)

      InitialiseRulesTab()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' --- No cleanup required ---
    End Try
  End Sub

#End Region

#Region "Rule - Item Move Events"
  ' ==========================================================================================
  ' Routine: lvRuleValues_ItemDrag
  ' Purpose:
  '   Initiates a drag operation when the user drags an item from the Rule Values ListView.
  ' Parameters:
  '   sender - The ListView raising the event.
  '   e      - Contains the dragged ListViewItem.
  ' Returns:
  '   None
  ' Notes:
  '   - Always initiates a Move drag effect.
  ' ==========================================================================================
  Private Sub lvRuleValues_ItemDrag(sender As Object, e As ItemDragEventArgs) Handles lvRuleValues.ItemDrag

    Dim item = DirectCast(e.Item, ListViewItem)
    DoDragDrop(item, DragDropEffects.Move)
  End Sub

  ' ==========================================================================================
  ' Routine: lvRuleValues_DragEnter
  ' Purpose:
  '   Determines whether the incoming drag data is valid for the Rule Values ListView.
  ' Parameters:
  '   sender - The ListView receiving the drag.
  '   e      - Drag event data including allowed effects and payload.
  ' Returns:
  '   None
  ' Notes:
  '   - Accepts FieldTag (from TreeView) or ListViewItem (from Filters).
  '   - Ensures a Move effect is applied when valid.
  ' ==========================================================================================
  Private Sub lvRuleValues_DragEnter(sender As Object, e As DragEventArgs) Handles lvRuleValues.DragEnter

    If e.Data.GetDataPresent(GetType(FieldTag)) OrElse
     e.Data.GetDataPresent(GetType(ListViewItem)) Then

      e.Effect = e.AllowedEffect And DragDropEffects.Move
      If e.Effect = DragDropEffects.None Then
        e.Effect = e.AllowedEffect
      End If
    Else
      e.Effect = DragDropEffects.None
    End If
  End Sub

  ' ==========================================================================================
  ' Routine: lvRuleValues_DragOver
  ' Purpose:
  '   Continuously validates drag data while hovering over the Rule Values ListView.
  ' Parameters:
  '   sender - The ListView receiving the drag.
  '   e      - Drag event data including allowed effects and payload.
  ' Returns:
  '   None
  ' Notes:
  '   - Mirrors DragEnter logic to maintain consistent drag feedback.
  ' ==========================================================================================
  Private Sub lvRuleValues_DragOver(sender As Object, e As DragEventArgs) Handles lvRuleValues.DragOver

    If e.Data.GetDataPresent(GetType(FieldTag)) OrElse
     e.Data.GetDataPresent(GetType(ListViewItem)) Then

      e.Effect = e.AllowedEffect And DragDropEffects.Move
      If e.Effect = DragDropEffects.None Then
        e.Effect = e.AllowedEffect
      End If
    Else
      e.Effect = DragDropEffects.None
    End If
  End Sub

  ' ==========================================================================================
  ' Routine: lvRuleValues_DragDrop
  ' Purpose:
  '   Handles drop operations onto the Rule Values ListView.
  ' Parameters:
  '   sender - The ListView receiving the drop.
  '   e      - Drag event data including payload.
  ' Returns:
  '   None
  ' Notes:
  '   - FieldTag → adds a new value.
  '   - ListViewItem from Filters → moves item between lists.
  ' ==========================================================================================
  Private Sub lvRuleValues_DragDrop(sender As Object, e As DragEventArgs) Handles lvRuleValues.DragDrop

    ' From TreeView → add new value
    If e.Data.GetDataPresent(GetType(FieldTag)) Then
      Dim tag = DirectCast(e.Data.GetData(GetType(FieldTag)), FieldTag)
      AddFieldEverywhere(tag, lvRuleValues)
      UpdateRuleTypeUI()
      Exit Sub
    End If

    ' From Filters ListView → move between lists
    If e.Data.GetDataPresent(GetType(ListViewItem)) Then
      Dim item = DirectCast(e.Data.GetData(GetType(ListViewItem)), ListViewItem)
      If item.ListView Is lvRuleFilters Then
        MoveItemBetweenLists(item, lvRuleFilters, lvRuleValues)
        UpdateRuleTypeUI()
      End If
    End If
  End Sub

  ' ==========================================================================================
  ' Routine: lvRuleFilters_ItemDrag
  ' Purpose:
  '   Initiates a drag operation when the user drags an item from the Rule Filters ListView.
  ' Parameters:
  '   sender - The ListView raising the event.
  '   e      - Contains the dragged ListViewItem.
  ' Returns:
  '   None
  ' Notes:
  '   - Always initiates a Move drag effect.
  ' ==========================================================================================
  Private Sub lvRuleFilters_ItemDrag(sender As Object, e As ItemDragEventArgs) Handles lvRuleFilters.ItemDrag

    Dim item = DirectCast(e.Item, ListViewItem)
    DoDragDrop(item, DragDropEffects.Move)
  End Sub

  ' ==========================================================================================
  ' Routine: lvRuleFilters_DragEnter
  ' Purpose:
  '   Determines whether the incoming drag data is valid for the Rule Filters ListView.
  ' Parameters:
  '   sender - The ListView receiving the drag.
  '   e      - Drag event data including allowed effects and payload.
  ' Returns:
  '   None
  ' Notes:
  '   - Accepts FieldTag (from TreeView) or ListViewItem (from Values).
  ' ==========================================================================================
  Private Sub lvRuleFilters_DragEnter(sender As Object, e As DragEventArgs) Handles lvRuleFilters.DragEnter

    If e.Data.GetDataPresent(GetType(FieldTag)) OrElse
     e.Data.GetDataPresent(GetType(ListViewItem)) Then

      e.Effect = e.AllowedEffect And DragDropEffects.Move
      If e.Effect = DragDropEffects.None Then
        e.Effect = e.AllowedEffect
      End If
    Else
      e.Effect = DragDropEffects.None
    End If
  End Sub

  ' ==========================================================================================
  ' Routine: lvRuleFilters_DragOver
  ' Purpose:
  '   Continuously validates drag data while hovering over the Rule Filters ListView.
  ' Parameters:
  '   sender - The ListView receiving the drag.
  '   e      - Drag event data including allowed effects and payload.
  ' Returns:
  '   None
  ' Notes:
  '   - Mirrors DragEnter logic to maintain consistent drag feedback.
  ' ==========================================================================================
  Private Sub lvRuleFilters_DragOver(sender As Object, e As DragEventArgs) Handles lvRuleFilters.DragOver

    If e.Data.GetDataPresent(GetType(FieldTag)) OrElse
     e.Data.GetDataPresent(GetType(ListViewItem)) Then

      e.Effect = e.AllowedEffect And DragDropEffects.Move
      If e.Effect = DragDropEffects.None Then
        e.Effect = e.AllowedEffect
      End If
    Else
      e.Effect = DragDropEffects.None
    End If
  End Sub

  ' ==========================================================================================
  ' Routine: lvRuleFilters_DragDrop
  ' Purpose:
  '   Handles drop operations onto the Rule Filters ListView.
  ' Parameters:
  '   sender - The ListView receiving the drop.
  '   e      - Drag event data including payload.
  ' Returns:
  '   None
  ' Notes:
  '   - FieldTag → adds a new filter.
  '   - ListViewItem from Values → moves item between lists.
  ' ==========================================================================================
  Private Sub lvRuleFilters_DragDrop(sender As Object, e As DragEventArgs) Handles lvRuleFilters.DragDrop

    ' From TreeView → add new filter
    If e.Data.GetDataPresent(GetType(FieldTag)) Then
      Dim tag = DirectCast(e.Data.GetData(GetType(FieldTag)), FieldTag)
      AddFieldEverywhere(tag, lvRuleFilters)
      Exit Sub
    End If

    ' From Values ListView → move between lists
    If e.Data.GetDataPresent(GetType(ListViewItem)) Then
      Dim item = DirectCast(e.Data.GetData(GetType(ListViewItem)), ListViewItem)
      If item.ListView Is lvRuleValues Then
        MoveItemBetweenLists(item, lvRuleValues, lvRuleFilters)
        UpdateRuleTypeUI()
      End If
    End If
  End Sub

  ' ==========================================================================================
  ' Routine: tvRuleFields_ItemDrag
  ' Purpose:
  '   Initiates a drag operation when the user drags a field node from the TreeView.
  ' Parameters:
  '   sender - The TreeView raising the event.
  '   e      - Contains the dragged TreeNode.
  ' Returns:
  '   None
  ' Notes:
  '   - Ignores root/view nodes.
  '   - Wraps field metadata into a typed FieldTag for downstream consumers.
  ' ==========================================================================================
  Private Sub tvRuleFields_ItemDrag(sender As Object, e As ItemDragEventArgs) Handles tvRuleFields.ItemDrag

    Dim node = DirectCast(e.Item, TreeNode)
    ' Ignore root/view nodes
    If node.Parent Is Nothing Then Exit Sub
    Dim tag = TryCast(node.Tag, FieldTag)
    If tag Is Nothing Then Exit Sub
    Dim data As New DataObject()
    data.SetData(GetType(FieldTag), tag)
    DoDragDrop(data, DragDropEffects.Move)

  End Sub

  ' ==========================================================================================
  ' Routine: tvRuleFields_DragEnter
  ' Purpose:
  '   Determines whether the TreeView accepts the incoming drag data.
  ' Parameters:
  '   sender - The TreeView receiving the drag.
  '   e      - Drag event data including payload.
  ' Returns:
  '   None
  ' Notes:
  '   - Only accepts ListViewItem (from Values or Filters) for removal.
  ' ==========================================================================================
  Private Sub tvRuleFields_DragEnter(sender As Object, e As DragEventArgs) Handles tvRuleFields.DragEnter

    If e.Data.GetDataPresent(GetType(ListViewItem)) Then
      e.Effect = DragDropEffects.Move
    End If
  End Sub

  ' ==========================================================================================
  ' Routine: tvRuleFields_DragDrop
  ' Purpose:
  '   Handles drop operations onto the TreeView.
  ' Parameters:
  '   sender - The TreeView receiving the drop.
  '   e      - Drag event data including payload.
  ' Returns:
  '   None
  ' Notes:
  '   - Only supports removing items from Values/Filters when dropped onto the TreeView.
  ' ==========================================================================================
  Private Sub tvRuleFields_DragDrop(sender As Object, e As DragEventArgs) Handles tvRuleFields.DragDrop

    If Not e.Data.GetDataPresent(GetType(ListViewItem)) Then Exit Sub

    Dim item = DirectCast(e.Data.GetData(GetType(ListViewItem)), ListViewItem)
    Dim sourceLv = item.ListView

    If sourceLv Is lvRuleValues OrElse sourceLv Is lvRuleFilters Then
      Dim tag = TryCast(item.Tag, FieldTag)
      RemoveFieldEverywhere(tag)
      UpdateRuleTypeUI()   ' values list changed → recompute rule type options
    End If
  End Sub

  ' ==========================================================================================
  ' Routine: MoveItem
  ' Purpose:
  '   Move a ListViewItem up or down by a relative offset.
  ' Parameters:
  '   lv    - the ListView containing the item
  '   item  - the ListViewItem to move
  '   delta - relative movement (+1 down, -1 up)
  ' Returns:
  '   None
  ' Notes:
  '   - Ensures the new index stays within valid bounds.
  '   - Preserves the item's Tag and selection state.
  ' ==========================================================================================
  Private Sub MoveItem(lv As System.Windows.Forms.ListView, item As ListViewItem, delta As Integer)
    Dim oldIndex = item.Index
    Dim newIndex = Math.Max(0, Math.Min(lv.Items.Count - 1, oldIndex + delta))
    If newIndex = oldIndex Then Exit Sub

    lv.Items.RemoveAt(oldIndex)
    lv.Items.Insert(newIndex, item)
    item.Selected = True
    ' --- Set the Rule Filter Expression for display ---
    UpdateRuleExpressionDisplay()
  End Sub

  ' ==========================================================================================
  ' Routine: MoveItemTo
  ' Purpose:
  '   Move a ListViewItem to an absolute index (beginning or end).
  ' Parameters:
  '   lv       - the ListView containing the item
  '   item     - the ListViewItem to move
  '   newIndex - the target index within the ListView
  ' Returns:
  '   None
  ' Notes:
  '   - Used for "Move to Beginning" and "Move to End".
  '   - Preserves the item's Tag and selection state.
  ' ==========================================================================================
  Private Sub MoveItemTo(lv As System.Windows.Forms.ListView, item As ListViewItem, newIndex As Integer)
    Dim oldIndex = item.Index
    If newIndex = oldIndex Then Exit Sub

    lv.Items.RemoveAt(oldIndex)
    lv.Items.Insert(newIndex, item)
    item.Selected = True
    ' --- Set the Rule Filter Expression for display ---
    UpdateRuleExpressionDisplay()
  End Sub

  ' ==========================================================================================
  ' Routine: MoveItemBetweenLists
  ' Purpose:
  '   Transfer a ListViewItem between the Values and Filters lists, recreating the item
  '   with appropriate tag while preserving metadata.
  ' Parameters:
  '   item   - the ListViewItem being moved
  '   fromLv - the source ListView
  '   toLv   - the destination ListView
  ' Returns:
  '   None
  ' Notes:
  '   - Preserves the FieldTag stored in item.Tag.
  ' ==========================================================================================
  Private Sub MoveItemBetweenLists(item As ListViewItem, fromLv As System.Windows.Forms.ListView, toLv As System.Windows.Forms.ListView)
    fromLv.Items.Remove(item)

    Dim newItem As New ListViewItem(item.Text)
    Dim tag As FieldTag = CType(item.Tag, FieldTag)


    ' If moving from Filters → Values, drop conditions
    If toLv Is lvRuleValues Then
      tag.FieldOperator = "" ' no operator 
      tag.OpenParenCount = "" ' no open parentheses 
      tag.CloseParenCount = "" ' no close parentheses 
      tag.BooleanOperator = "" ' no boolean operator 
      tag.ValueBinding = ""
      tag.SlicingMode = "" ' no slicing mode
    Else
      ' If moving from Values → Filters
      tag.FieldOperator = GetDefaultOperator(tag) ' If there is allowed list then make sure we default to one in list
      tag.OpenParenCount = DEFAULT_OPEN_PARENTHESES_COUNT
      tag.CloseParenCount = DEFAULT_CLOSE_PARENTHESES_COUNT
      If toLv.Items.Count = 0 Then
        tag.BooleanOperator = "" ' no boolean operator for the first filter
      Else
        tag.BooleanOperator = DEFAULT_BOOLEAN_OPERATOR
      End If
      tag.ValueBinding = DEFAULT_VALUE_BINDING
      tag.SlicingMode = "" ' no slicing mode
    End If
    newItem.Tag = tag
    toLv.Items.Add(newItem)
    newItem.Selected = True
    ' --- Set the Rule Filter Expression for display ---
    UpdateRuleExpressionDisplay()
  End Sub

  ' ==========================================================================================
  ' Routine: UncheckFieldInTree
  ' Purpose:
  '   Locate the TreeView node representing a field and uncheck it when the field is removed
  '   from Values or Filters.
  '
  ' Parameters:
  '   tag - the FieldTag identifying the field to uncheck
  ' Notes:
  '   - TreeNode.Tag ALWAYS holds a FieldTag for selectable nodes.
  '   - For list-typed fields, identity is (viewName, fieldName, fieldID, listTypeId).
  ' ==========================================================================================
  Private Sub UncheckFieldInTree(tag As FieldTag)

    For Each root As TreeNode In tvRuleFields.Nodes
      For Each node As TreeNode In root.Nodes
        Dim ft As FieldTag = TryCast(node.Tag, FieldTag)
        If ft Is Nothing Then Continue For

        If String.Equals(ft.ViewName, tag.ViewName, StringComparison.OrdinalIgnoreCase) AndAlso
         String.Equals(ft.FieldName, tag.FieldName, StringComparison.OrdinalIgnoreCase) AndAlso
         String.Equals(If(ft.FieldID, String.Empty), If(tag.FieldID, String.Empty),
                       StringComparison.OrdinalIgnoreCase) AndAlso
          String.Equals(If(ft.ListTypeID, String.Empty), If(tag.ListTypeID, String.Empty),
                       StringComparison.OrdinalIgnoreCase) Then

          node.Checked = False
          Exit Sub
        End If
      Next
    Next
  End Sub
#End Region

#Region "Rule - Helpers"
  ' ==========================================================================================
  ' Routine: SetFilterCondition
  ' Purpose:
  '   Update the condition displayed for a filter ListViewItem.
  ' Parameters:
  '   item - the ListViewItem whose condition is being changed
  '   value   - the operator symbol selected by the user
  '             or the open parantheses symbol selected by the user
  '             or the close parantheses symbol selected by the user
  '             or the boolean operator symbol selected by the user
  ' Returns:
  '   None
  ' Notes:
  '   - Updates only the visual condition; the save routine reads this value directly.
  '   - Applies only to items in the Filters ListView.
  ' ==========================================================================================
  Private Sub SetFilterCondition(Condition As Condition, item As ListViewItem, value As String)
    Dim tag As FieldTag = CType(item.Tag, FieldTag)
    Select Case Condition
      Case Condition.FieldOperater
        ' Update visual operator in the listview
        tag.FieldOperator = value
      Case Condition.OpenParentheses
        ' Update visual open parentheses in the listview
        tag.OpenParenCount = value.Length
      Case Condition.CloseParentheses
        ' Update visual close parentheses in the listview
        tag.CloseParenCount = value.Length
      Case Condition.BooleanOperater
        ' Update visual boolean operator in the listview
        tag.BooleanOperator = value
    End Select
    ' --- Set the Rule Filter Expression for display ---
    UpdateRuleExpressionDisplay()
  End Sub

  ' ==========================================================================================
  ' Routine:      UpdateRuleExpressionDisplay
  '
  ' Purpose:
  '   Rebuilds the rule filter expression string and updates the RichTextBox display.
  '   Applies syntax‑highlighting to any mismatch markers (e.g., <<MISSING_OP>>, <<UNBALANCED>>).
  '   This routine has no side effects beyond updating the RichTextBox UI.
  '
  ' Parameters:
  '   None
  '
  ' Returns:
  '   None
  '
  ' Notes:
  '   - Must be called after any change to lvRuleFilters.
  '   - Does not mutate model state.
  '   - Uses HighlightMarkers to apply colour formatting.
  ' ==========================================================================================
  Private Sub UpdateRuleExpressionDisplay()
    Dim text = BuildFilterExpression(lvRuleFilters, FilterExpressionRenderMode.Rules)

    txtRuleFilterExpression.SuspendLayout()
    txtRuleFilterExpression.Clear()
    txtRuleFilterExpression.Text = text

    HighlightMarkers(txtRuleFilterExpression)

    txtRuleFilterExpression.SelectionStart = 0
    txtRuleFilterExpression.SelectionLength = 0
    txtRuleFilterExpression.ResumeLayout()
  End Sub

  ' ==========================================================================================
  ' Routine: MapRuleDetailUIToModel
  ' Purpose:
  '   Convert the UI panels (Values, Filters) into ActionRuleDetail.
  '
  ' Notes:
  '   - ALWAYS rebuilds ActionRuleDetail from scratch.
  '   - PRESERVES FilterID when editing.
  '   - Generates new FilterID when missing.
  ' ==========================================================================================
  Private Sub MapRuleDetailUIToModel()
    Try
      ' --- Always rebuild detail from scratch ---
      _model.ActionRuleDetail = New UIExcelRuleDesignerRuleRowDetail()

      ' --- Identity ---
      _model.ActionRuleDetail.RuleID = _model.ActionRule.RuleID
      _model.ActionRuleDetail.RuleName = _model.ActionRule.RuleName
      _model.ActionRuleDetail.RuleType = _model.ActionRule.RuleType

      ' ============================================================
      ' VALUES LIST → SelectedValues
      ' ============================================================
      For Each item As ListViewItem In lvRuleValues.Items
        Dim tag = DirectCast(item.Tag, FieldTag)

        Dim sv As New UIExcelRuleDesignerRuleSelectedValue With {
                .View = tag.ViewName,
                .Field = tag.FieldName,
                .FieldID = tag.FieldID,
                .SourceView = tag.SourceView,
                .SourceField = tag.SourceField,
                .ListTypeID = tag.ListTypeID
            }

        _model.ActionRuleDetail.SelectedValues.Add(sv)
      Next

      ' ============================================================
      ' FILTERS LIST → Filters (with stable FilterID)
      ' ============================================================
      For Each item As ListViewItem In lvRuleFilters.Items
        Dim tag = DirectCast(item.Tag, FieldTag)

        Dim f As New UIExcelRuleDesignerRuleFilter

        ' --- Preserve FilterID if it exists ---
        If String.IsNullOrEmpty(tag.FilterID) Then
          f.FilterID = Guid.NewGuid().ToString()
        Else
          f.FilterID = tag.FilterID
        End If

        f.View = tag.ViewName
        f.Field = tag.FieldName
        f.FieldID = tag.FieldID
        f.SourceView = tag.SourceView
        f.SourceField = tag.SourceField
        f.ListTypeID = tag.ListTypeID
        f.FieldOperator = tag.FieldOperator
        f.OpenParenCount = CInt(tag.OpenParenCount)
        f.CloseParenCount = CInt(tag.CloseParenCount)
        f.BooleanOperator = tag.BooleanOperator
        f.SlicingMode = tag.SlicingMode
        f.ValueBinding = tag.ValueBinding
        f.LiteralValue = If(String.IsNullOrWhiteSpace(tag.LiteralValue), Nothing, tag.LiteralValue.Trim())
        _model.ActionRuleDetail.Filters.Add(f)
      Next

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: AddFieldEverywhere
  ' Purpose:
  '   Add a field (view + field) to the Values or Filters list in  UI
  ' Parameters:
  '   viewName  - internal view name (e.g. "vwDim_Resource")
  '   fieldName - internal field name (e.g. "PreferredName")
  '   ID        - optional ID to assign to the field (can be used for fieldID's) 
  '   target - listview to add to
  ' Returns:
  '   None
  ' Notes:
  ' ==========================================================================================
  Private Sub AddFieldEverywhere(ft As FieldTag,
                               target As System.Windows.Forms.ListView)


    Try
      If target Is Nothing Then Exit Sub
      If _model Is Nothing Then Exit Sub

      ' --- Prevent duplicates in the target ListView ---
      For Each item As ListViewItem In target.Items
        Dim existing = TryCast(item.Tag, FieldTag)
        If existing IsNot Nothing AndAlso
           String.Equals(existing.ViewName, ft.ViewName, StringComparison.OrdinalIgnoreCase) AndAlso
           String.Equals(existing.FieldName, ft.FieldName, StringComparison.OrdinalIgnoreCase) AndAlso
           String.Equals(If(existing.FieldID, ""), If(ft.FieldID, ""), StringComparison.OrdinalIgnoreCase) AndAlso
           String.Equals(If(existing.ListTypeID, ""), If(ft.ListTypeID, ""), StringComparison.OrdinalIgnoreCase) Then
          Exit Sub
        End If

      Next

      ' --- Build ListViewItem from FieldTag.DisplayName ---
      Dim lvi As New ListViewItem(ft.DisplayName)

      If target Is lvRuleFilters Then
        ' Filters list has operator + parentheses + boolean operator
        ft.FieldOperator = GetDefaultOperator(ft) ' If there is allowed list then make sure we default to one in list
        ft.OpenParenCount = DEFAULT_OPEN_PARENTHESES_COUNT
        ft.CloseParenCount = DEFAULT_CLOSE_PARENTHESES_COUNT
        If target.Items.Count = 0 Then
          ft.BooleanOperator = "" ' no boolean operator for the first filter
        Else
          ft.BooleanOperator = DEFAULT_BOOLEAN_OPERATOR
        End If
        ft.ValueBinding = DEFAULT_VALUE_BINDING
        ft.SlicingMode = "" ' no slicing mode
      End If
      lvi.Tag = ft

      target.Items.Add(lvi)

      ' --- Ensure TreeView node is checked ---
      Dim node As TreeNode = Nothing
      Dim key = (ft.ViewName, ft.FieldName & "|" & If(ft.FieldID, "") & "|" & ft.ListTypeID)

      If _fieldIndex.TryGetValue(key, node) Then
        If Not node.Checked Then
          _suppressTreeAfterCheck = True
          node.Checked = True
          _suppressTreeAfterCheck = False
        End If
      End If

      ' --- Update filter expression if Filters list changed ---
      If target Is lvRuleFilters Then
        UpdateRuleExpressionDisplay()
      End If

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' No cleanup required
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: RemoveFieldEverywhere
  ' Purpose:
  '   Remove a field (view + field) from the Values list, unchecks tree and rebuild filter
  '   expression  in UI
  ' Parameters:
  '   viewName  - internal view name (e.g. "vwDim_Resource")
  '   fieldName - internal field name (e.g. "PreferredName")
  ' Returns:
  '   None
  ' Notes:
  '   - Removes matching entries from ListViews unchecks tree if not supressed i.e. called from 
  '     tvRuleFields_AfterCheck, rebuild filter expression
  ' ==========================================================================================
  Private Sub RemoveFieldEverywhere(ft As FieldTag)

    Try
      ' --- Remove from Filters UI ---
      For i As Integer = lvRuleFilters.Items.Count - 1 To 0 Step -1
        Dim tag = TryCast(lvRuleFilters.Items(i).Tag, FieldTag)
        If tag IsNot Nothing AndAlso
          String.Equals(tag.ViewName, ft.ViewName, StringComparison.OrdinalIgnoreCase) AndAlso
          String.Equals(tag.FieldName, ft.FieldName, StringComparison.OrdinalIgnoreCase) AndAlso
          String.Equals(If(tag.FieldID, ""), If(ft.FieldID, ""), StringComparison.OrdinalIgnoreCase) AndAlso
          String.Equals(If(tag.ListTypeID, ""), If(ft.ListTypeID, ""), StringComparison.OrdinalIgnoreCase) Then
          lvRuleFilters.Items.RemoveAt(i)
        End If
      Next

      ' --- Remove from Values UI ---
      For i As Integer = lvRuleValues.Items.Count - 1 To 0 Step -1
        Dim tag = TryCast(lvRuleValues.Items(i).Tag, FieldTag)
        If tag IsNot Nothing AndAlso
          String.Equals(tag.ViewName, ft.ViewName, StringComparison.OrdinalIgnoreCase) AndAlso
          String.Equals(tag.FieldName, ft.FieldName, StringComparison.OrdinalIgnoreCase) AndAlso
          String.Equals(If(tag.FieldID, ""), If(ft.FieldID, ""), StringComparison.OrdinalIgnoreCase) AndAlso
          String.Equals(If(tag.ListTypeID, ""), If(ft.ListTypeID, ""), StringComparison.OrdinalIgnoreCase) Then
          lvRuleValues.Items.RemoveAt(i)
        End If
      Next

      ' --- Uncheck in tree unless suppressed ---
      If Not _suppressTreeAfterCheck Then
        UncheckFieldInTree(ft)
      End If

      ' --- Rebuild filter expression ---
      UpdateRuleExpressionDisplay()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: GetDefaultOperator
  ' Purpose:
  '   Determine the correct default operator for a filter field based on the field’s metadata.
  '   If the field defines AllowedOperators, the first allowed operator is returned.
  '   If no operator constraints exist, the global default operator (first in AvailableOperators)
  '   is returned instead.
  '
  ' Parameters:
  '   tag (FieldTag)
  '       The UI metadata for the filter row, used to identify the field and retrieve its
  '       corresponding ExcelRuleViewMapField metadata.
  '
  ' Returns:
  '   String
  '       The operator that should be used as the default for this field. Guaranteed to be valid
  '       for the field’s AllowedOperators constraint if one exists.
  '
  ' Notes:
  '   - Prevents invalid defaults such as "=" for fields that only allow ">=" or "<=".
  '   - Must be used whenever a field is added to lvRuleFilters or moved from Values → Filters.
  '   - Never returns Nothing; always returns a valid operator string.
  ' ==========================================================================================
  Private Function GetDefaultOperator(tag As FieldTag) As String
    Try
      Dim fieldInfo = _model.ViewMapHelper.GetField(tag.ViewName, tag.FieldName)

      If fieldInfo IsNot Nothing AndAlso
       fieldInfo.AllowedOperators IsNot Nothing AndAlso
       fieldInfo.AllowedOperators.Count > 0 Then

        Return fieldInfo.AllowedOperators(0)   ' first allowed operator
      End If

      ' fallback to global default only if no constraints exist
      Return DEFAULT_FIELD_OPERATOR
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return DEFAULT_FIELD_OPERATOR
    End Try
  End Function

  ' ==========================================================================================
  ' Routine: ValidateRuleBeforeSave
  ' Purpose:
  '   Perform structural validation of the rule definition before saving, including:
  '     - Rule name
  '     - Value column count and rule type consistency
  '     - Filter parameter validation
  '     - Availability date pairing
  '
  ' Parameters: None (reads UI directly).
  ' Returns:
  '   Boolean -
  '       True  if the rule definition is valid.
  '       False if validation fails and the save must be cancelled.
  ' Notes:
  '   - Must be called BEFORE the model is updated.
  '   - Displays user-facing messages instead of throwing exceptions.
  ' ==========================================================================================
  Private Function ValidateRuleBeforeSave() As Boolean
    ' ------------------------------------------------------------
    ' 1. Rule name must be provided
    ' ------------------------------------------------------------
    Dim ruleName As String = txtRuleName.Text.Trim()
    If ruleName.Length = 0 Then
      MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                              "Rule name cannot be empty.",
                              UI_NAME,
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Warning)
      Return False
    End If
    ' ------------------------------------------------------------
    ' 2. Must have at least one value column
    ' ------------------------------------------------------------
    Dim valueCount As Integer = lvRuleValues.Items.Count
    If valueCount = 0 Then
      MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                              "At least one value column must be defined before saving the rule.",
                              UI_NAME,
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Warning)
      Return False
    End If
    ' ------------------------------------------------------------
    ' 3. Rule type must be valid according to UpdateRuleTypeUI logic
    ' ------------------------------------------------------------
    If valueCount >= 2 Then
      If Not optRange.Checked Then
        MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                                  "Rules with two or more value columns must use the Range rule type.",
                                  UI_NAME,
                                  MessageBoxButtons.OK,
                                  MessageBoxIcon.Warning)
        Return False
      End If

    ElseIf valueCount = 1 Then
      If Not optSingle.Checked AndAlso Not optList.Checked Then
        MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                                  "Rules with one value column must be Single or List.",
                                  UI_NAME,
                                  MessageBoxButtons.OK,
                                  MessageBoxIcon.Warning)
        Return False
      End If
    End If
    ' ------------------------------------------------------------
    ' 4. Validate each filter row (parameter types, missing values, etc.)
    ' ------------------------------------------------------------
    If Not ValidateFilterExpressionFromUi() Then
      ' ValidateFilterExpressionFromUi already shows the message
      Return False
    End If
    ' ------------------------------------------------------------
    ' 5. Validate that ValueBinding exists and is consistent
    ' ------------------------------------------------------------
    For Each item As ListViewItem In lvRuleFilters.Items
      Dim tag = DirectCast(item.Tag, FieldTag)
      Dim fm = _model.ViewMapHelper.GetField(tag.ViewName, tag.FieldName)
      If fm Is Nothing Then Continue For
      ' 1. ValueBinding must be set (string must not be empty)
      Dim vb As String = If(tag.ValueBinding, "").Trim()
      If vb.Length = 0 Then
        MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                              $"Filter on '{tag.DisplayName}' is missing a value binding mode.",
                              UI_NAME,
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Warning)
        Return False
      End If
      ' 2. Rule-bound → literal must exist (but we do NOT validate membership)
      If tag.ValueBinding = ValueBinding.Rule.ToString() Then
        Dim lit As String = If(tag.LiteralValue, "").Trim()
        If lit.Length = 0 Then
          MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                                  $"Filter on '{tag.DisplayName}' requires a literal value because it is rule-bound.",
                                  UI_NAME,
                                  MessageBoxButtons.OK,
                                  MessageBoxIcon.Warning)
          Return False
        End If
        ' Validate literal type only
        Dim dataType As String = fm.DataType?.Trim().ToLowerInvariant()
        If Not ValidateRuleParameterType(dataType, ExcelRefType.Literal.ToString(), lit, "") Then
          MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                                  $"The literal value for '{tag.DisplayName}' is not valid for its data type.",
                                  UI_NAME,
                                  MessageBoxButtons.OK,
                                  MessageBoxIcon.Warning)
          Return False
        End If
      End If
    Next

    ' ------------------------------------------------------------
    ' 6. Validate AvailabilityFromDate / AvailabilityToDate pairing
    ' ------------------------------------------------------------
    Dim hasFrom As Boolean = False
    Dim hasTo As Boolean = False

    For Each item As ListViewItem In lvRuleFilters.Items
      Dim tag = DirectCast(item.Tag, FieldTag)
      If tag.FieldName = "AvailabilityFromDate" Then hasFrom = True
      If tag.FieldName = "AvailabilityToDate" Then hasTo = True
    Next

    If hasFrom Xor hasTo Then
      MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                              "Availability From Date and To Date must both be included when filtering by availability.",
                              UI_NAME,
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Warning)
      Return False
    End If
    ' ------------------------------------------------------------
    ' 7. Validate slicing mode for fields that support slicing
    ' ------------------------------------------------------------
    For Each item As ListViewItem In lvRuleFilters.Items
      Dim tag = DirectCast(item.Tag, FieldTag)

      ' Look up the field metadata
      Dim fm = _model.ViewMapHelper.GetField(tag.ViewName, tag.FieldName)
      If fm Is Nothing Then Continue For

      If fm.SupportsSlicing Then
        Dim slicing As String = If(tag.SlicingMode, "").Trim()

        If slicing.Length = 0 Then
          MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                                  $"Filter on '{tag.DisplayName}' requires a slicing mode.",
                                  UI_NAME,
                                  MessageBoxButtons.OK,
                                  MessageBoxIcon.Warning)
          Return False
        End If
      End If
    Next
    ' ------------------------------------------------------------
    ' 8. If both availability fields exist, their slicing modes must match
    ' ------------------------------------------------------------
    If hasFrom AndAlso hasTo Then
      Dim fromMode As String = Nothing
      Dim toMode As String = Nothing

      For Each item As ListViewItem In lvRuleFilters.Items
        Dim tag = DirectCast(item.Tag, FieldTag)

        If tag.FieldName = "AvailabilityFromDate" Then
          fromMode = If(tag.SlicingMode, "").Trim()
        ElseIf tag.FieldName = "AvailabilityToDate" Then
          toMode = If(tag.SlicingMode, "").Trim()
        End If
      Next

      If Not String.Equals(fromMode, toMode, StringComparison.OrdinalIgnoreCase) Then
        MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                              "Availability From Date and To Date must use the same slicing mode.",
                              UI_NAME,
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Warning)
        Return False
      End If
    End If
    Return True

  End Function

  ' ==========================================================================================
  ' Routine: ValidateFilterExpressionFromUi
  ' Purpose:
  '   Validates the boolean expression defined by lvRuleFilters, including:
  '     - First filter must not have a BooleanOperator.
  '     - Subsequent filters must have AND or OR.
  '     - Parentheses must be balanced.
  '     - OpenParenCount and CloseParenCount must never drive nesting below zero.
  ' Parameters:
  '   None (reads lvRuleFilters directly).
  ' Returns:
  '   Boolean -
  '       True  if the expression is valid.
  '       False if invalid (and displays a user-facing message).
  ' Notes:
  '   - This replaces RuleFilterExpressionBuilder.ValidateFilterExpression(detail),
  '     which cannot be used because the model is not populated yet.
  ' ==========================================================================================
  Private Function ValidateFilterExpressionFromUi() As Boolean

    Dim depth As Integer = 0

    For i As Integer = 0 To lvRuleFilters.Items.Count - 1

      Dim item As ListViewItem = lvRuleFilters.Items(i)
      Dim tag As FieldTag = CType(item.Tag, FieldTag)

      Dim openCount As Integer = CInt(tag.OpenParenCount)
      Dim closeCount As Integer = CInt(tag.CloseParenCount)
      Dim boolOp As String = If(tag.BooleanOperator, "").Trim()

      ' --- First filter must not have a BooleanOperator ---
      If i = 0 Then
        If boolOp.Length > 0 Then
          MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                                      "The first filter cannot have a Boolean operator.",
                                      UI_NAME,
                                      MessageBoxButtons.OK,
                                      MessageBoxIcon.Warning)
          Return False
        End If
      Else
        ' Subsequent filters must have AND or OR
        If Not (boolOp.Equals("AND", StringComparison.OrdinalIgnoreCase) OrElse
                    boolOp.Equals("OR", StringComparison.OrdinalIgnoreCase)) Then

          MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                                      $"Filter {i + 1} must specify AND or OR as the Boolean operator.",
                                      UI_NAME,
                                      MessageBoxButtons.OK,
                                      MessageBoxIcon.Warning)
          Return False
        End If
      End If

      ' --- Validate parentheses counts ---
      If openCount < 0 Then
        MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                                  $"Filter {i + 1} has a negative OpenParenCount.",
                                  UI_NAME,
                                  MessageBoxButtons.OK,
                                  MessageBoxIcon.Warning)
        Return False
      End If

      If closeCount < 0 Then
        MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                                  $"Filter {i + 1} has a negative CloseParenCount.",
                                  UI_NAME,
                                  MessageBoxButtons.OK,
                                  MessageBoxIcon.Warning)
        Return False
      End If

      depth += openCount
      depth -= closeCount

      If depth < 0 Then
        MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                                  $"Filter {i + 1} closes more parentheses than have been opened.",
                                  UI_NAME,
                                  MessageBoxButtons.OK,
                                  MessageBoxIcon.Warning)
        Return False
      End If

    Next

    ' --- Final depth must be zero ---
    If depth <> 0 Then
      MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                              "Filter expression has unbalanced parentheses.",
                              UI_NAME,
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Warning)
      Return False
    End If

    Return True

  End Function
#End Region

#Region "Apply - Initialisation Events"
  ' ==========================================================================================
  ' Routine: InitialiseApplyTab
  ' Purpose:
  '   Populate the Apply tab with rule list and list types.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' Notes:
  '   - Called on form load.
  ' ==========================================================================================
  Private Sub InitialiseApplyTab()
    Try
      ' --- Ensure combo boxes are not data-bound so we can use Items.Clear/Add safely ---
      If cmbApplyNames.DataSource IsNot Nothing Then cmbApplyNames.DataSource = Nothing
      If cmbApplyRules.DataSource IsNot Nothing Then cmbApplyRules.DataSource = Nothing
      If cmbApplyListSelectType.DataSource IsNot Nothing Then cmbApplyListSelectType.DataSource = Nothing

      ' --- Populate apply ComboBox ---
      cmbApplyNames.Items.Clear()
      cmbApplyNames.Items.Add(NEW_APPLY_SENTINEL)

      For Each a In _model.ApplyInstances
        Dim item As New ComboApplyItem With {
            .Display = a.ApplyName,
            .Value = a.ApplyID
        }
        cmbApplyNames.Items.Add(item)
      Next
      cmbApplyNames.DropDownStyle = ComboBoxStyle.DropDownList
      If cmbApplyNames.Items.Count > 0 Then cmbApplyNames.SelectedIndex = 0
      ' How to call 
      'If TypeOf cmbApplyNames.SelectedItem Is String Then
      '  ' NEW_APPLY_SENTINEL selected
      'Else
      '  Dim item As ComboApplyItem = CType(cmbApplyNames.SelectedItem, ComboApplyItem)
      '  Dim applyId As String = item.Value
      'End If

      ' --- Populate rule ComboBox ---
      cmbApplyRules.Items.Clear()

      For Each r In _model.Rules
        If r.RuleType = ExcelRuleType.ListOfValues.ToString() Then
          Dim item As New ComboRuleItem With {
            .Display = r.RuleName,
            .Value = r.RuleID
          }
          cmbApplyRules.Items.Add(item)
        End If
      Next
      cmbApplyRules.DropDownStyle = ComboBoxStyle.DropDownList
      ' How to call
      'Dim selected As ComboRuleItem = TryCast(cmbApplyRules.SelectedItem, ComboRuleItem)
      'If selected IsNot Nothing Then
      '  Dim ruleId As String = selected.Value
      'End If
      ' --- Populate list types (views) ---
      cmbApplyListSelectType.Items.Clear()
      cmbApplyListSelectType.DataSource = ExcelListSelectTypeMap.BindingListOfStrings()
      cmbApplyListSelectType.DisplayMember = "Display"
      cmbApplyListSelectType.ValueMember = "Value"
      cmbApplyListSelectType.DropDownStyle = ComboBoxStyle.DropDownList
      cmbApplyListSelectType.SelectedIndex = -1

      ' --- Clear detail UI ---
      lvApplyFilters.Items.Clear()
      lvApplyFilters.OwnerDraw = True ' allow custom context menu drawing later
      txtApplyFilterExpression.Text = ""

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' --- No cleanup required ---
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: BindApplyHeaderFromSelectedApply
  ' Purpose:
  '   Bind only the Apply header (name, rule combo, list-select type)
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' Notes:
  '   - Called after selecting an apply.
  ' ==========================================================================================
  Private Sub BindApplyHeaderFromSelectedApply()
    If _model Is Nothing OrElse _model.SelectedApply Is Nothing Then Exit Sub

    _isBinding = True
    Try
      ' Apply name
      txtApplyName.Text = _model.SelectedApply.ApplyName

      ' Select rule in cmbApplyRules based on RuleID
      Dim ruleId As String = _model.SelectedApply.RuleID
      Dim idx As Integer = -1
      For i As Integer = 0 To cmbApplyRules.Items.Count - 1
        Dim item = TryCast(cmbApplyRules.Items(i), ComboRuleItem)
        If item IsNot Nothing AndAlso item.Value = ruleId Then
          idx = i
          Exit For
        End If
      Next
      cmbApplyRules.SelectedIndex = idx

      ' ListSelectType binding
      BindListSelectTypeCombo(_model.SelectedApply.ListSelectType)

    Finally
      _isBinding = False
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: BindApplyFiltersFromCurrentRuleAndApply
  ' Purpose:
  '   Bind filters + parameter bindings from _model.RuleDetail and _model.SelectedApply
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' Notes:
  '   - Called after selecting an apply or an existing rule.
  ' ==========================================================================================
  Private Sub BindApplyFiltersFromCurrentRuleAndApply()
    Try
      _isBinding = True

      lvApplyFilters.Items.Clear()

      If _model Is Nothing OrElse _model.RuleDetail Is Nothing Then
        UpdateApplyExpressionDisplay()
        Exit Sub
      End If

      For Each f In _model.RuleDetail.Filters
        Dim tag As New FieldTag With {
        .ViewName = f.View,
        .FieldName = f.Field,
        .FieldID = f.FieldID,
        .SourceView = f.SourceView,
        .SourceField = f.SourceField,
        .ListTypeID = f.ListTypeID,
        .FilterID = f.FilterID,
        .FieldOperator = f.FieldOperator,
        .OpenParenCount = f.OpenParenCount,
        .CloseParenCount = f.CloseParenCount,
        .BooleanOperator = f.BooleanOperator,
        .ValueBinding = f.ValueBinding,
        .LiteralValue = f.LiteralValue
      }

        Dim item As New ListViewItem(f.Field)

        ' Match parameter by FilterID (if SelectedApply exists)
        Dim p As UIExcelRuleDesignerApplyParameter =
        _model.SelectedApply?.Parameters?.
          FirstOrDefault(Function(x) x.FilterID = f.FilterID)

        If p IsNot Nothing Then
          tag.RefType = p.RefType
          If p.RefType = ExcelRefType.Literal.ToString() Then
            tag.LiteralValue = If(p.LiteralValue, "").Trim()
            tag.RefValue = Nothing
          Else
            tag.LiteralValue = Nothing
            tag.RefValue = If(p.RefValue, "").Trim()
          End If
        End If

        item.Tag = tag
        lvApplyFilters.Items.Add(item)
      Next

      UpdateApplyExpressionDisplay()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    Finally
      _isBinding = False
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: BindListSelectTypeCombo
  ' Purpose:
  '   Bind the Apply List Select Type combo box (cmbApplyListSelectType) using the value
  '   stored on the selected Apply instance. Handles both string and enum values and ensures
  '   the combo is correctly matched to its ValueMember ("EnumValue").
  '
  ' Parameters:
  '   value - Object
  '           The ListSelectType stored on UIExcelRuleDesignerApplyInstance. May be:
  '             - ExcelListSelectType enum
  '             - String representation of the enum
  '             - Nothing (no selection)
  '
  ' Returns:
  '   None.
  '
  ' Notes:
  '   - cmbApplyListSelectType is populated with BindingItem(Of ExcelListSelectType)
  '     where ValueMember = "EnumValue".
  '   - SelectedValue must be assigned an actual ExcelListSelectType enum instance.
  '   - Silent failures occur if a string is assigned directly; this routine prevents that.
  ' ==========================================================================================
  Private Sub BindListSelectTypeCombo(value As Object)

    If value Is Nothing Then
      cmbApplyListSelectType.SelectedIndex = -1
      Exit Sub
    End If

    ' Always treat stored value as a string
    Dim s As String = CStr(value).Trim()

    ' Check if the combo contains this string in its ValueMember
    Dim found As Boolean =
        cmbApplyListSelectType.Items.
            Cast(Of BindingItemString)().
            Any(Function(bi) String.Equals(bi.Value, s, StringComparison.OrdinalIgnoreCase))

    If found Then
      cmbApplyListSelectType.SelectedValue = s
    Else
      cmbApplyListSelectType.SelectedIndex = -1
    End If

    'Dim enumVal As ExcelListSelectType

    '' --- Case 1: Stored as string (most common) ---
    'If TypeOf value Is String Then
    '  If [Enum].TryParse(value.ToString(), True, enumVal) Then
    '    cmbApplyListSelectType.SelectedValue = enumVal
    '  Else
    '    cmbApplyListSelectType.SelectedIndex = -1
    '  End If
    '  Exit Sub
    'End If

    '' --- Case 2: Stored as enum already ---
    'If TypeOf value Is ExcelListSelectType Then
    '  cmbApplyListSelectType.SelectedValue = CType(value, ExcelListSelectType)
    '  Exit Sub
    'End If

    '' --- Fallback ---
    'cmbApplyListSelectType.SelectedIndex = -1

  End Sub

  '  Private Sub btnClearAll_Click(sender As Object, e As EventArgs) Handles btnClearAll.Click
  '    Dim xl = CType(ExcelDnaUtil.Application, Excel.Application)
  '    Dim rng As Excel.Range = xl.Selection

  '    For Each cell As Excel.Range In rng.Cells
  '      ExcelCellRuleStore.ClearCellRule(cell)
  '    Next
  '  End Sub
#End Region

#Region "Apply - Control Events"
  ' ==========================================================================================
  ' Routine: cmbApplyNames_SelectedIndexChanged
  ' Purpose:
  '   Handle user selection of an existing apply or the <New Apply…> sentinel.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' Notes:
  '   - Selecting <New Apply…> clears detail and sets PendingAction = Add.
  '   - Selecting an existing rule loads detail via loader refresh path.
  ' ==========================================================================================
  Private Sub cmbApplyNames_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbApplyNames.SelectedIndexChanged

    Try
      ' --- Safe guard the model hasn't been cleared ---
      If _model Is Nothing Then
        UILoaderSaverExcelRuleDesigner.LoadExcelRuleDesignerModel(_model)
        If _model Is Nothing Then Exit Sub
      End If

      ' --- Clear detail UI ---
      txtApplyName.Text = ""
      cmbApplyListSelectType.SelectedIndex = -1
      cmbApplyRules.SelectedIndex = -1
      lvApplyFilters.Items.Clear()
      txtApplyFilterExpression.Text = ""

      Dim selectedObj As Object = cmbApplyNames.SelectedItem

      ' --- Sentinel path (new apply) ---
      If TypeOf selectedObj Is String AndAlso
           CStr(selectedObj) = NEW_APPLY_SENTINEL Then
        _model.SelectedApply = Nothing
        _model.PendingAction = ExcelRuleDesignerAction.Add
        Return
      End If

      ' --- Existing apply path ---
      Dim item As ComboApplyItem = TryCast(selectedObj, ComboApplyItem)
      If item Is Nothing Then Exit Sub   ' Defensive
      Dim applyId As String = item.Value

      ' Locate the Apply instance
      Dim apply = _model.ApplyInstances.FirstOrDefault(Function(a) a.ApplyID = applyId)
      If apply Is Nothing Then Exit Sub

      _model.SelectedApply = apply
      _model.PendingAction = ExcelRuleDesignerAction.Update

      ' --- CRITICAL: set SelectedRule for this Apply before reload ---
      Dim rule = _model.Rules.FirstOrDefault(Function(r) r.RuleID = apply.RuleID)
      _model.SelectedRule = rule

      ' --- Reload model (refresh path) so RuleDetail matches SelectedRule ---
      UILoaderSaverExcelRuleDesigner.LoadExcelRuleDesignerModel(_model)

      ' Now bind header + filters
      BindApplyHeaderFromSelectedApply()
      BindApplyFiltersFromCurrentRuleAndApply()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' --- No cleanup required ---
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: cmbApplyRules_SelectedIndexChanged
  ' Purpose:
  '   Handle user selection of an existing rule
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' Notes:
  '   - Selecting an existing rule loads detail via loader refresh path.
  ' ==========================================================================================
  Private Sub cmbApplyRules_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbApplyRules.SelectedIndexChanged

    Try
      If _isBinding Then Exit Sub   ' if binding then exit

      ' --- Safe guard the model hasn't been cleared ---
      If _model Is Nothing Then
        UILoaderSaverExcelRuleDesigner.LoadExcelRuleDesignerModel(_model)
        If _model Is Nothing Then Exit Sub
      End If

      ' --- Clear detail UI ---
      lvApplyFilters.Items.Clear()
      txtApplyFilterExpression.Text = ""
      cmbApplyListSelectType.SelectedIndex = -1

      ' --- Get selected rule (ComboRuleItem) from cmbApplyRules ---
      Dim selectedObj As Object = cmbApplyRules.SelectedItem
      Dim item As ComboRuleItem = TryCast(selectedObj, ComboRuleItem)
      If item Is Nothing Then Exit Sub   ' Defensive
      Dim ruleId As String = item.Value
      ' Find the rule by RuleID (NOT by name anymore)
      Dim selected = _model.Rules.FirstOrDefault(Function(r) r.RuleID = ruleId)
      If selected Is Nothing Then Exit Sub

      _model.SelectedRule = selected

      ' --- Reload model (refresh path) ---
      UILoaderSaverExcelRuleDesigner.LoadExcelRuleDesignerModel(_model)

      ' For a NEW Apply: SelectedApply is Nothing → filters with no parameters
      ' For an existing Apply: parameters will only match if FilterID lines up
      BindApplyFiltersFromCurrentRuleAndApply()


    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' --- No cleanup required ---
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: btnApplySave_Click
  ' Purpose:
  '   Handle the Apply → Save action. Validates the Apply section, rebuilds the
  '   UIExcelRuleDesignerApplyInstance from the current UI state, preserves IDs for updates,
  '   captures the selected Excel range, and delegates persistence to the loader/saver.
  '
  ' Parameters:
  '   sender  - Event source (Button)
  '   e       - Event arguments
  '
  ' Returns:
  '   None. Updates _model.ActionApply and triggers SavePendingApplyAction.
  '
  ' Notes:
  '   - Always rebuilds ActionApply from scratch to avoid stale state.
  '   - Preserves ApplyID when editing an existing Apply instance.
  '   - Extracts parameters from lvApplyFilters, including FilterID, RefType,
  '     LiteralValue, and RefValue.
  '   - Delegates all persistence to UILoaderSaverExcelRuleDesigner.
  '   - Reloads the model/UI after saving to ensure consistency.
  ' ==========================================================================================
  Private Sub btnApplySave_Click(sender As Object, e As EventArgs) Handles btnApplySave.Click
    Try
      ' --- Validate apply values ---
      If Not ValidateApplyBeforeSave() Then Exit Sub
      ' Capture the selected range (may be disconnected)
      Dim xl = CType(ExcelDnaUtil.Application, Excel.Application)
      Dim selectedRange As Excel.Range = xl.Selection

      ' --- Safeguard: ensure model exists ---
      If _model Is Nothing Then
        UILoaderSaverExcelRuleDesigner.LoadExcelRuleDesignerModel(_model)
        If _model Is Nothing Then Exit Sub
      End If
      ' --- Always rebuild ActionApply from scratch ---
      _model.ActionApply = New UIExcelRuleDesignerApplyInstance()

      ' --- Preserve ApplyID on Update ---
      If _model.PendingAction = ExcelRuleDesignerAction.Update AndAlso
             _model.SelectedApply IsNot Nothing Then

        _model.ActionApply.ApplyID = _model.SelectedApply.ApplyID
      End If
      ' --- Assign ApplyName ---
      _model.ActionApply.ApplyName = txtApplyName.Text.Trim()
      ' --- Assign RuleID ---
      _model.ActionApply.RuleID = _model.SelectedRule.RuleID
      ' --- Assign List Select Type ---
      _model.ActionApply.ListSelectType = cmbApplyListSelectType.SelectedValue
      ' ---- Assign parameters ---
      For Each item As ListViewItem In lvApplyFilters.Items
        Dim tag = DirectCast(item.Tag, FieldTag)
        ' Only save parameters for filters that are not rule-bound (i.e. where the user can select a value)
        If tag.ValueBinding <> ValueBinding.Rule.ToString() Then
          Dim p As New UIExcelRuleDesignerApplyParameter
          ' --- Preserve FilterID  ---
          p.FilterID = tag.FilterID
          If Not String.IsNullOrWhiteSpace(tag.RefType) Then
            p.RefType = tag.RefType
          Else
            ' No binding type selected yet → choose a default
            p.RefType = ExcelRefType.Literal.ToString()
          End If
          p.LiteralValue = If(tag.LiteralValue, "").Trim()
          p.RefValue = If(tag.RefValue, "").Trim()
          _model.ActionApply.Parameters.Add(p)
        End If

      Next

      ' Save + apply
      UILoaderSaverExcelRuleDesigner.SavePendingApplyAction(_model, selectedRange)

      ' Reload model/UI
      ReloadModel()
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' --- No cleanup required ---
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: btnApplyDelete_Click
  '
  ' Purpose:
  '   Delete the currently selected Apply instance (RuleRegion) by:
  '     - Building the ActionApply payload on the model
  '     - Setting PendingAction = Delete
  '     - Calling SavePendingApplyAction with a null target (delete does not need a range)
  '     - Reloading the model/UI
  '
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  '
  ' Returns:
  '   None.
  '
  ' Notes:
  '   - Requires a valid Apply selection in cmbApplyNames.
  '   - SavePendingApplyAction is the single entry point for Apply commits; UI must set
  '     model.ActionApply and model.PendingAction before calling it.
  ' ==========================================================================================
  Private Sub btnApplyDelete_Click(sender As Object, e As EventArgs) Handles btnApplyDelete.Click
    Try
      ' --- Validate selection ---
      Dim selectedItem = TryCast(cmbApplyNames.SelectedItem, ComboApplyItem)
      If selectedItem Is Nothing Then Exit Sub

      ' --- Ensure model present ---
      If _model Is Nothing Then
        UILoaderSaverExcelRuleDesigner.LoadExcelRuleDesignerModel(_model)
        If _model Is Nothing Then Exit Sub
      End If

      ' --- Locate the Apply instance in the model ---
      Dim apply = _model.ApplyInstances.
                  FirstOrDefault(Function(a) a.ApplyID = selectedItem.Value)

      If apply Is Nothing Then Exit Sub

      ' --- Build ActionApply payload (minimum required for delete) ---
      _model.ActionApply = New UIExcelRuleDesignerApplyInstance()
      _model.ActionApply.ApplyID = apply.ApplyID
      _model.ActionApply.RuleID = apply.RuleID
      _model.ActionApply.ListSelectType = apply.ListSelectType
      ' Parameters are not required for deletion, leave empty

      ' --- Set pending action to Delete and persist via loader/saver ---
      _model.PendingAction = ExcelRuleDesignerAction.Delete

      ' Pass Nothing for target range when deleting (SavePendingApplyAction accepts a range)
      UILoaderSaverExcelRuleDesigner.SavePendingApplyAction(_model, Nothing)

      ' --- Refresh UI/model to reflect deletion ---
      ReloadModel()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  '  Private Sub btnSet_Click(sender As Object, e As EventArgs) Handles btnSet.Click
  '    SetRangeID()
  '  End Sub

  '  Private Sub btnGet_Click(sender As Object, e As EventArgs) Handles btnGet.Click
  '    GetRangeID()
  '  End Sub
#End Region

#Region "Apply - Helpers"
  ' ==========================================================================================
  ' Routine:      UpdateApplyExpressionDisplay
  '
  ' Purpose:
  '   Rebuilds the apply filter expression string and updates the RichTextBox display.
  '   Applies syntax‑highlighting to any mismatch markers 
  '   This routine has no side effects beyond updating the RichTextBox UI.
  '
  ' Parameters:
  '   None
  '
  ' Returns:
  '   None
  '
  ' Notes:
  '   - Must be called after any change to lvApplyFilters.
  '   - Does not mutate model state.
  '   - Uses HighlightMarkers to apply colour formatting.
  ' ==========================================================================================
  Private Sub UpdateApplyExpressionDisplay()
    Dim text = BuildFilterExpression(lvApplyFilters, FilterExpressionRenderMode.Apply)

    txtApplyFilterExpression.SuspendLayout()
    txtApplyFilterExpression.Clear()
    txtApplyFilterExpression.Text = text

    HighlightMarkers(txtApplyFilterExpression)

    txtApplyFilterExpression.SelectionStart = 0
    txtApplyFilterExpression.SelectionLength = 0
    txtApplyFilterExpression.ResumeLayout()
  End Sub

  ' ==========================================================================================
  ' Routine: SelectLiteralParameter
  ' Purpose:
  '   Prompt the user to enter a literal value (centred on the Excel window) and store it
  '   into the ListViewItem parameter column (Tag).
  ' Parameters:
  '   item  - The ListViewItem whose parameter column will be updated.
  ' Returns:
  '   None
  ' Notes:
  '   - Uses InputBoxHelper.ShowExcelInputBox to centre the InputBox on the Excel window.
  '   - Treats the InputBox return as Object and checks for Boolean False (cancel).
  ' ==========================================================================================
  Private Sub SelectLiteralParameter(item As ListViewItem)
    Dim app As Excel.Application = Nothing
    Dim result As Object = Nothing

    Try
      app = CType(ExcelDnaUtil.Application, Excel.Application)
      Dim tag As FieldTag = CType(item.Tag, FieldTag)
      ' Use Excel InputBox (Type:=2 => string) via helper to centre the dialog
      result = InputBoxHelper.ShowExcelInputBox(Me, app,
                "Enter literal value:",
                "Literal Value",
                If(tag.LiteralValue, "").Trim(),
                2)

      ' Cancel pressed -> result is Boolean False
      If TypeOf result Is Boolean AndAlso CType(result, Boolean) = False Then Exit Sub

      Dim value As String = CStr(result)

      If String.IsNullOrWhiteSpace(value) Then Exit Sub
      ' Must be correct type for filter field
      If String.IsNullOrWhiteSpace(tag.RefType) Then
        MessageBox.Show("Please select a Reference Type for this filter.")
        Exit Sub
      End If
      Dim refType As String = tag.RefType
      Dim field As ExcelRuleViewMapField = _model.ViewMapHelper.GetField(tag.ViewName, tag.FieldName)
      Dim dataType As String = field.DataType?.Trim().ToLowerInvariant()
      If Not ValidateRuleParameterType(dataType, refType, value, "") Then
        MessageBox.Show("Parameter must be the correct type for the filter field.")
        Exit Sub
      End If
      ' Store value
      tag.LiteralValue = value

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    Finally
      app = Nothing
      result = Nothing
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: SelectAbsoluteRangeParameter
  ' Purpose: Prompts the user to select a 1‑dimensional absolute Excel range and stores it
  '          into the ListViewItem fieldTag.
  ' Parameters:
  '   item  - The ListViewItem whose parameter column will be updated.
  ' Returns:
  '   None
  ' Notes:
  '   - Uses Excel's InputBox(Type:=8) to capture a Range object.
  '   - Enforces 1D constraint (single row OR single column).
  '   - Stores A1 address (relative:=False) in tag.
  ' ==========================================================================================
  Private Sub SelectAbsoluteRangeParameter(item As ListViewItem)

    Dim app As Excel.Application = Nothing
    Dim rng As Excel.Range = Nothing
    Dim existingRange As Excel.Range = Nothing
    Dim result As Object = Nothing

    Try
      ' --- Normal execution ---
      app = CType(ExcelDnaUtil.Application, Excel.Application)
      Dim tag As FieldTag = CType(item.Tag, FieldTag)
      ' --- Convert existing range back into a real range (if present) ---
      Dim existingOffset As String = If(tag.RefValue, "").Trim()
      If Not String.IsNullOrEmpty(existingOffset) Then
        Try
          existingRange = app.Range(existingOffset)
        Catch
          existingRange = Nothing
        End Try
      End If
      ' Use InputBoxHelper to center the dialog on Excel; returns Object (Range) or Boolean False
      result = InputBoxHelper.ShowExcelInputBox(Me, app,
            "Select a 1D range:",
            "Absolute Range",
            If(existingRange IsNot Nothing, existingRange.Address(True, True), ""),
            8)

      'Cancel pressed
      If TypeOf result Is Boolean AndAlso result = False Then Exit Sub

      rng = CType(result, Excel.Range)
      ' Must be a single contiguous area
      If rng.Areas.Count <> 1 Then
        MessageBox.Show("Please select a single contiguous range.")
        Exit Sub
      End If
      ' Must be 1D
      If Not (rng.Rows.Count = 1 OrElse rng.Columns.Count = 1) Then
        MessageBox.Show("Range must be one-dimensional (single row or single column).")
        Exit Sub
      End If
      ' Must be correct type for filter field
      If String.IsNullOrWhiteSpace(tag.RefType) Then
        MessageBox.Show("Please select a Reference Type for this filter.")
        Exit Sub
      End If
      Dim refType As String = tag.RefType
      Dim field As ExcelRuleViewMapField = _model.ViewMapHelper.GetField(tag.ViewName, tag.FieldName)
      Dim dataType As String = field.DataType?.Trim().ToLowerInvariant()
      If Not ValidateRuleParameterType(dataType, refType, "", rng.Address) Then
        MessageBox.Show("Parameter must be the correct type for the filter field.")
        Exit Sub
      End If
      ' Store A1 address (no $)
      Dim addr As String =
            rng.Address(RowAbsolute:=False,
                        ColumnAbsolute:=False,
                        ReferenceStyle:=Excel.XlReferenceStyle.xlA1,
                        External:=False)
      tag.RefValue = addr

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    Finally
      ' --- Cleanup ---
      If rng IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(rng)
      rng = Nothing
      app = Nothing
      result = Nothing

    End Try

  End Sub

  ' ==========================================================================================
  ' Routine: SelectRelativeRangeParameter
  ' Purpose: Prompts the user to select a 1‑dimensional range and converts it into a relative
  '          R1C1 reference based on an anchor cell.
  ' Parameters:
  '   item  - The ListViewItem whose parameter column will be updated.
  ' Returns:
  '   None
  ' Notes:
  '   - Anchor must be an Excel.Range stored in item.Tag or another known location.
  '   - Enforces 1D constraint (single row OR single column).
  '   - Stores R1C1 relative address in fieldTag.
  ' ==========================================================================================
  Private Sub SelectRelativeRangeParameter(item As ListViewItem)

    Dim app As Excel.Application = Nothing
    Dim anchor As Excel.Range = Nothing
    Dim rng As Excel.Range = Nothing
    Dim existingRange As Excel.Range = Nothing
    Dim result As Object = Nothing
    Dim result2 As Object = Nothing

    Try
      ' --- Normal execution ---
      app = CType(ExcelDnaUtil.Application, Excel.Application)
      Dim tag As FieldTag = CType(item.Tag, FieldTag)
      ' ============================================================
      ' STEP 1: Prompt user for anchor cell (must be a single cell)
      ' ============================================================
      Do
        result = InputBoxHelper.ShowExcelInputBox(Me, app,
                "Select the ANCHOR cell (single cell only):",
                "Relative Range Anchor",
                "",
                8)
        'User cancelled → result is Boolean False
        If TypeOf result Is Boolean AndAlso CType(result, Boolean) = False Then Exit Sub
        anchor = CType(result, Excel.Range)
        ' Must be exactly one cell
        If anchor.Cells.Count = 1 Then Exit Do
        MessageBox.Show("Anchor must be a single cell. Please try again.")
        Marshal.ReleaseComObject(anchor)
        anchor = Nothing
      Loop

      ' ============================================================
      ' STEP 2: Highlight anchor (red dashed border)
      ' ============================================================
      anchor.BorderAround(
            LineStyle:=Excel.XlLineStyle.xlDash,
            Color:=RGB(255, 0, 0),
            Weight:=Excel.XlBorderWeight.xlMedium)

      ' ============================================================
      ' STEP 3: Convert existing offset back into a real range
      ' ============================================================
      Dim existingOffset As String = If(tag.RefValue, "").Trim()

      If Not String.IsNullOrEmpty(existingOffset) Then
        Try
          existingRange = anchor.Parent.Range(existingOffset)
        Catch
          existingRange = Nothing
        End Try
      End If
      ' ============================================================
      ' STEP 4: Prompt user for the actual relative range
      ' ============================================================
      result2 = InputBoxHelper.ShowExcelInputBox(Me, app,
                "Select a 1D range relative to the anchor:",
                "Relative Range",
                If(existingRange IsNot Nothing, existingRange.Address(True, True), ""),
                8)

      'User cancelled → result2 is Boolean False
      If TypeOf result2 Is Boolean AndAlso CType(result2, Boolean) = False Then Exit Sub
      rng = CType(result2, Excel.Range)
      ' Must be a single contiguous area
      If rng.Areas.Count <> 1 Then
        MessageBox.Show("Please select a single contiguous range.")
        Exit Sub
      End If
      ' Must be 1D
      If Not (rng.Rows.Count = 1 OrElse rng.Columns.Count = 1) Then
        MessageBox.Show("Range must be one-dimensional (single row or single column).")
        Exit Sub
      End If
      ' ============================================================
      ' STEP 5: Must be correct type for filter field
      ' ============================================================
      If String.IsNullOrWhiteSpace(tag.RefType) Then
        MessageBox.Show("Please select a Reference Type for this filter.")
        Exit Sub
      End If
      Dim refType As String = tag.RefType
      Dim field As ExcelRuleViewMapField = _model.ViewMapHelper.GetField(tag.ViewName, tag.FieldName)
      Dim dataType As String = field.DataType?.Trim().ToLowerInvariant()
      If Not ValidateRuleParameterType(dataType, refType, "", rng.Address) Then
        MessageBox.Show("Parameter must be the correct type for the filter field.")
        Exit Sub
      End If
      ' ============================================================
      ' STEP 6: Compute relative R1C1 offset
      ' ============================================================
      Dim relAddr As String =
            rng.Address(RowAbsolute:=False,
                        ColumnAbsolute:=False,
                        ReferenceStyle:=Excel.XlReferenceStyle.xlR1C1,
                        External:=False,
                        RelativeTo:=anchor)

      tag.RefValue = relAddr

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' ============================================================
      ' STEP 7: Remove anchor highlight
      ' ============================================================
      If anchor IsNot Nothing Then
        Try
          anchor.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
        Catch
        End Try
      End If

      ' Cleanup COM objects
      If existingRange IsNot Nothing Then Marshal.ReleaseComObject(existingRange)
      If rng IsNot Nothing Then Marshal.ReleaseComObject(rng)
      If anchor IsNot Nothing Then Marshal.ReleaseComObject(anchor)

      existingRange = Nothing
      rng = Nothing
      anchor = Nothing
      app = Nothing
      result = Nothing
      result2 = Nothing
    End Try

  End Sub

  ' ==========================================================================================
  ' Routine: BuildNamedRangeSubmenu
  ' Purpose: Dynamically populates the "Select Parameter..." context-menu item with a list of
  '          valid 1‑dimensional named ranges from the workbook. Each menu item updates the
  '          ListViewItem parameter cell (Text + Tag) when clicked.
  ' Parameters:
  '   parentMenu - The ToolStripMenuItem representing "Select Parameter..."
  '   item       - The ListViewItem whose parameter column will be updated.
  ' Returns:
  '   None
  ' Notes:
  '   - Only named ranges that refer to a single contiguous 1D range (row or column) are shown.
  '   - Uses Excel‑DNA: CType(ExcelDnaUtil.Application, Excel.Application)
  '   - Ensures COM cleanup and exception safety.
  ' ==========================================================================================
  Private Sub BuildNamedRangeSubmenu(parentMenu As ToolStripMenuItem, item As ListViewItem)

    Dim app As Excel.Application = Nothing
    Dim names As Excel.Names = Nothing
    Dim refers As Excel.Range = Nothing

    Try
      ' --- Normal execution ---
      parentMenu.DropDownItems.Clear()
      Dim tag As FieldTag = CType(item.Tag, FieldTag)
      app = CType(ExcelDnaUtil.Application, Excel.Application)
      names = app.Names
      For Each n As Excel.Name In names
        refers = Nothing
        ' Skip hidden names that start with "RM_" (Resource Management names
        If Not n.Visible AndAlso n.Name.StartsWith("RM_", StringComparison.OrdinalIgnoreCase) Then
          Continue For
        End If
        Try
          refers = n.RefersToRange
        Catch
          ' Skip invalid or non-range names
          Continue For
        End Try
        If refers Is Nothing Then Continue For
        ' Must be a single contiguous area
        If refers.Areas.Count <> 1 Then Continue For
        ' Must be 1D (single row OR single column)
        If Not (refers.Rows.Count = 1 OrElse refers.Columns.Count = 1) Then Continue For
        ' Build menu item
        Dim mi As New ToolStripMenuItem(n.Name) With {
                .Tag = n.Name,
                .Checked = (tag.RefValue IsNot Nothing AndAlso
                            CStr(tag.RefValue) = n.Name)
            }

        AddHandler mi.Click,
                Sub()
                  Try
                    ' Must be correct type for filter field
                    If String.IsNullOrWhiteSpace(tag.RefType) Then
                      MessageBox.Show("Please select a Reference Type for this filter.")
                      Exit Sub
                    End If
                    Dim refType As String = tag.RefType
                    Dim field As ExcelRuleViewMapField = _model.ViewMapHelper.GetField(tag.ViewName, tag.FieldName)
                    Dim dataType As String = field.DataType?.Trim().ToLowerInvariant()
                    If Not ValidateRuleParameterType(dataType, refType, "", n.Name) Then
                      MessageBox.Show("Parameter must be the correct type for the filter field.")
                      Exit Sub
                    End If
                    tag.RefValue = n.Name

                    ' Update checkmarks
                    For Each sibling As ToolStripMenuItem In parentMenu.DropDownItems
                      sibling.Checked = (sibling Is mi)
                    Next
                    UpdateApplyExpressionDisplay()
                  Catch ex As Exception
                    ErrorHandler.UnHandleError(ex)
                  End Try
                End Sub

        parentMenu.DropDownItems.Add(mi)

        ' Release COM for this iteration
        If refers IsNot Nothing Then
          System.Runtime.InteropServices.Marshal.ReleaseComObject(refers)
          refers = Nothing
        End If

      Next

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' --- Cleanup ---
      If refers IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(refers)
      If names IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(names)

      refers = Nothing
      names = Nothing
      app = Nothing

    End Try

  End Sub

  ' ==========================================================================================
  ' Routine: ValidateApplyBeforeSave
  ' Purpose:
  '   Perform data validation of the rule definition before saving
  ' Parameters:
  '   None
  ' Returns:
  '   True if validation passes
  ' Notes:
  ' ==========================================================================================
  Private Function ValidateApplyBeforeSave() As Boolean
    ' --- Validate ApplyName ---
    If cmbApplyNames.SelectedItem Is Nothing Then
      MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                            "Please select an Apply.",
                            UI_NAME,
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning)
      Return False
    End If
    Dim name As String = txtApplyName.Text.Trim()
    If name.Length = 0 Then
      MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                            "Apply name cannot be empty.",
                            UI_NAME,
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning)
      Return False
    End If

    '--- Validate Rule and List Select Type ---
    If cmbApplyListSelectType.SelectedItem Is Nothing Or cmbApplyRules.SelectedItem Is Nothing Then
      MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                            "Please select a Rule and List Select Type.",
                            UI_NAME,
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning)
      Return False
    End If
    '--- Validate parameters in lvApplyFilters ---
    For Each item As ListViewItem In lvApplyFilters.Items
      Dim tag As FieldTag = CType(item.Tag, FieldTag)
      ' If rule-bound, skip all apply-time binding validation
      If tag.ValueBinding = ValueBinding.Rule.ToString() Then
        ' Literal is already defined in the rule
        Continue For
      End If

      ' --- Validate RefType selected ---
      If String.IsNullOrWhiteSpace(tag.RefType) Then
        MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                              $"Please select a Reference Type for filter '{item.Text}'.",
                              UI_NAME,
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Warning)
        Return False
      End If
      Dim refType As String = tag.RefType.Trim()
      ' --- Validate LiteralValue or RefValue based on RefType ---
      If refType = ExcelRefType.Literal.ToString() Then
        If String.IsNullOrEmpty(tag.LiteralValue.Trim()) Then
          MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                                $"Please enter a Literal Value for filter '{item.Text}'.",
                                UI_NAME,
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning)
          Return False
        End If
      ElseIf refType = ExcelRefType.Address.ToString() Or refType = ExcelRefType.Offset.ToString() Or refType = ExcelRefType.Name.ToString() Then
        If String.IsNullOrEmpty(tag.RefValue.Trim()) Then
          MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                                $"Please enter a Reference Value for filter '{item.Text}'.",
                                UI_NAME,
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning)
          Return False
        End If
      End If
    Next

    '--- Validate Pending action already set to Add or Update ---
    If _model.PendingAction <> ExcelRuleDesignerAction.Add AndAlso
       _model.PendingAction <> ExcelRuleDesignerAction.Update Then
      MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                            "Internal error: Pending action is not set correctly.",
                            UI_NAME,
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error)
      Return False
    End If
    Return True
  End Function


#End Region
End Class
