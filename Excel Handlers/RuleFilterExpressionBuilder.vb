
'' ==========================================================================================
'' Module: RuleFilterExpressionBuilder
'' Purpose:
''   Provides validation and expression-tree construction for the boolean filter expression
''   defined by UIExcelRuleDesignerRuleRowDetail.Filters.
''
''   - ValidateFilterExpression:
''       Ensures parentheses are balanced and BooleanOperator usage is legal.
''   - BuildFilterExpressionTree:
''       Converts the linear filter list into a simple expression tree that can be
''       consumed by the rule engine (e.g. GetListOptions).
''
'' Notes:
''   - DESIGN-TIME ONLY: operates on UIExcelRuleDesignerRuleRowDetail.
''   - Does NOT resolve parameter values or touch runtime data.
'' ==========================================================================================
'Friend Module RuleFilterExpressionBuilder

'  ' ========================================================================================
'  ' Class: FilterExpressionNode
'  ' Purpose:
'  '   Represents a node in the boolean filter expression tree.
'  '   - Leaf nodes contain a single UIExcelRuleDesignerRuleFilter.
'  '   - Internal nodes contain an Operator ("AND"/"OR") and two child nodes.
'  ' ========================================================================================
'  Friend Class FilterExpressionNode
'    Public Property [Operator] As String          ' "AND", "OR", or "" for leaf
'    Public Property Left As FilterExpressionNode  ' Left child (for AND/OR)
'    Public Property Right As FilterExpressionNode ' Right child (for AND/OR)
'    Public Property Filter As UIExcelRuleDesignerRuleFilter ' Non-Nothing only for leaf nodes
'  End Class

'  ' ========================================================================================
'  ' Routine: ValidateFilterExpression
'  ' Purpose:
'  '   Validate the boolean expression defined by the Filters collection:
'  '     - Parentheses must be balanced.
'  '     - BooleanOperator must be "" for the first filter, "AND" or "OR" otherwise.
'  '     - OpenParenCount and CloseParenCount must never drive nesting below zero.
'  '
'  ' Parameters:
'  '   detail - UIExcelRuleDesignerRuleRowDetail containing Filters.
'  '
'  ' Returns:
'  '   None.
'  '
'  ' Notes:
'  '   - Throws InvalidOperationException if the expression is invalid.
'  '   - Intended to be called before saving a rule.
'  ' ========================================================================================
'  Friend Sub ValidateFilterExpression(detail As UIExcelRuleDesignerRuleRowDetail)

'    If detail Is Nothing Then
'      Throw New ArgumentNullException(NameOf(detail))
'    End If

'    Dim depth As Integer = 0

'    For i As Integer = 0 To detail.Filters.Count - 1
'      Dim f = detail.Filters(i)

'      ' --- First filter must not have a BooleanOperator ---
'      If i = 0 Then
'        If Not String.IsNullOrEmpty(f.BooleanOperator) Then
'          Throw New InvalidOperationException(
'            "The first filter in the list cannot have a Boolean operator.")
'        End If
'      Else
'        ' Subsequent filters must have AND or OR
'        If Not String.Equals(f.BooleanOperator, "AND", StringComparison.OrdinalIgnoreCase) AndAlso
'           Not String.Equals(f.BooleanOperator, "OR", StringComparison.OrdinalIgnoreCase) Then
'          Throw New InvalidOperationException(
'            $"Filter {i + 1} must specify AND or OR as the Boolean operator.")
'        End If
'      End If

'      ' --- Apply opening parentheses ---
'      If f.OpenParenCount < 0 Then
'        Throw New InvalidOperationException(
'          $"Filter {i + 1} has a negative OpenParenCount.")
'      End If
'      depth += f.OpenParenCount

'      ' --- Apply closing parentheses ---
'      If f.CloseParenCount < 0 Then
'        Throw New InvalidOperationException(
'          $"Filter {i + 1} has a negative CloseParenCount.")
'      End If
'      depth -= f.CloseParenCount

'      If depth < 0 Then
'        Throw New InvalidOperationException(
'          $"Filter {i + 1} closes more parentheses than have been opened.")
'      End If
'    Next

'    ' --- Final depth must be zero (balanced parentheses) ---
'    If depth <> 0 Then
'      Throw New InvalidOperationException(
'        "Filter expression has unbalanced parentheses.")
'    End If

'  End Sub

'  ' ========================================================================================
'  ' Routine: BuildFilterExpressionTree
'  ' Purpose:
'  '   Convert the linear Filters list (with BooleanOperator, OpenParenCount, CloseParenCount)
'  '   into a binary expression tree of FilterExpressionNode objects.
'  '
'  ' Parameters:
'  '   detail - UIExcelRuleDesignerRuleRowDetail containing Filters.
'  '
'  ' Returns:
'  '   FilterExpressionNode - Root of the expression tree, or Nothing if no filters.
'  '
'  ' Notes:
'  '   - Assumes ValidateFilterExpression has already been called successfully.
'  '   - Uses a simple operator-precedence model where AND/OR are left-associative and
'  '     parentheses control grouping.
'  '   - This is a DESIGN-TIME structure; runtime evaluation will walk this tree and
'  '     apply it to actual data rows.
'  ' ========================================================================================
'  Friend Function BuildFilterExpressionTree(detail As UIExcelRuleDesignerRuleRowDetail) _
'      As FilterExpressionNode

'    If detail Is Nothing OrElse detail.Filters Is Nothing OrElse detail.Filters.Count = 0 Then
'      Return Nothing
'    End If

'    ' Shunting-yard style: output stack for nodes, operator stack for AND/OR with parens.
'    Dim valueStack As New Stack(Of FilterExpressionNode)
'    Dim opStack As New Stack(Of String)
'    Dim parenStack As New Stack(Of Integer) ' track how many operators belong to each paren level

'    ' Helper to create a leaf node
'    Dim makeLeaf As Func(Of UIExcelRuleDesignerRuleFilter, FilterExpressionNode) =
'      Function(f As UIExcelRuleDesignerRuleFilter) As FilterExpressionNode
'        Return New FilterExpressionNode With {
'          .[Operator] = "",
'          .Left = Nothing,
'          .Right = Nothing,
'          .Filter = f
'        }
'      End Function

'    ' Helper to apply a single operator from opStack to valueStack
'    Dim applyOp As Action =
'      Sub()
'        If opStack.Count = 0 OrElse valueStack.Count < 2 Then Exit Sub

'        Dim op As String = opStack.Pop()
'        Dim rightNode As FilterExpressionNode = valueStack.Pop()
'        Dim leftNode As FilterExpressionNode = valueStack.Pop()

'        Dim parent As New FilterExpressionNode With {
'          .[Operator] = op,
'          .Left = leftNode,
'          .Right = rightNode,
'          .Filter = Nothing
'        }
'        valueStack.Push(parent)
'      End Sub

'    ' Track how many operators are associated with each parenthesis level
'    parenStack.Push(0)

'    For i As Integer = 0 To detail.Filters.Count - 1
'      Dim f = detail.Filters(i)

'      ' --- Handle opening parentheses BEFORE this filter ---
'      For p As Integer = 1 To f.OpenParenCount
'        parenStack.Push(0)
'      Next

'      ' --- Push leaf node for this filter ---
'      valueStack.Push(makeLeaf(f))

'      ' --- Handle BooleanOperator (glue to previous filter) ---
'      If i > 0 AndAlso Not String.IsNullOrEmpty(f.BooleanOperator) Then
'        ' For now AND/OR have same precedence; left-associative.
'        ' Apply the last operator at this level, then push the new one.
'        applyOp()
'        opStack.Push(f.BooleanOperator.ToUpperInvariant())

'        ' Increment operator count for current paren level
'        Dim topCount As Integer = parenStack.Pop()
'        parenStack.Push(topCount + 1)
'      End If

'      ' --- Handle closing parentheses AFTER this filter ---
'      For p As Integer = 1 To f.CloseParenCount
'        ' When closing a group, apply all operators that belong to this level.
'        Dim opsAtThisLevel As Integer = parenStack.Pop()
'        For k As Integer = 1 To opsAtThisLevel
'          applyOp()
'        Next
'      Next
'    Next

'    ' --- Apply any remaining operators (top-level) ---
'    While opStack.Count > 0
'      applyOp()
'    End While

'    If valueStack.Count <> 1 Then
'      Throw New InvalidOperationException("Filter expression could not be reduced to a single root node.")
'    End If

'    Return valueStack.Pop()

'  End Function

'End Module