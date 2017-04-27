Public Class Table
    Public numRange1 As Long
    Public numRange2 As Long
    Public equation1 As String
    Public equation2 As String
    Dim xLength As Integer = 0
    Private i As Integer

    Private Function FormatExpression(ByVal expression As String) As String
        Dim tableFormat1 As String = expression.Replace(" ", String.Empty)

        Dim xEvaluator As System.Text.RegularExpressions.MatchEvaluator = New System.Text.RegularExpressions.MatchEvaluator(AddressOf MultiplyX)
        tableFormat1 = System.Text.RegularExpressions.Regex.Replace(tableFormat1, "x+", xEvaluator)

        Return tableFormat1
    End Function

    Private Function MultiplyX(ByVal m As System.Text.RegularExpressions.Match) As String
        Return "*" & m.Value
    End Function
    Private Function ReplaceX(ByVal expression As String) As String
        Dim tableFormat2 As String = expression.Replace("x", numRange1)
        Return tableFormat2
    End Function
    Private Sub BtnApply_Click(sender As Object, e As EventArgs) Handles BtnApply.Click
        yValue.Clear()
        numRange1 = Range1.Text
        numRange2 = Range2.Text
        i = numRange1
        While i < numRange2
            xLength = xLength + 1
            i = i + 1
        End While
        Dim numY(xLength) As Long
        equation1 = TableFunction.Text
        equation1 = FormatExpression(equation1)
        i = 0
        While i < xLength
            equation2 = ReplaceX(equation1)
            getValue4()
            i = i + 1
            numRange1 = numRange1 + 1
        End While
        xLength = 0
    End Sub
End Class