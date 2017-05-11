Public Class Functions
    Public numRange1
    Public numRange2
    Public equation1 As String
    Public equation2 As String
    Dim xLength As Integer = 0
    Private i As Integer
    Dim lst1 = New Integer() {1, 2, 3}
    Dim lst2 = New Integer() {4, 5, 6}

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
        i = 0
        While i < xLength
            equation2 = ReplaceX(equation1)
            getValue4()
            i = i + 1
            numRange1 = numRange1 + 1
        End While
        xLength = 0
    End Sub

    Private Sub StandardToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StandardToolStripMenuItem.Click
        Form1.Show()
        Me.Hide()
    End Sub

    Private Sub ScientificToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ScientificToolStripMenuItem.Click
        Scientific.Show()
        Me.Hide()
    End Sub

    Private Sub ProgrammerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProgrammerToolStripMenuItem.Click
        Programmer.Show()
        Me.Hide()
    End Sub

    Private Sub GraphingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GraphingToolStripMenuItem.Click
        Graphing.Show()
        Me.Hide()
    End Sub

    Private Sub FunctionsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FunctionsToolStripMenuItem.Click
        Me.Show()
    End Sub

    Private Sub ConverterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConverterToolStripMenuItem.Click
        Converter.Show()
        Me.Hide()
    End Sub
    '---------------------------------
    'Function standardize(lst)
    'If lst(Len(lst) - 1) = 0 Then
    '     lst.pop()
    'End If
    'Return lst
    'End Function
    ' Function isZeroPoly(lst)
    'If lst = Nothing Then
    'Return True
    '  Else
    '         Return False
    '  End If
    '   End Function
    'Function addPoly(lst1, lst2)
    '   Dim result As Integer()
    'If Len(lst1) > Len(lst2) Then
    'For x As Integer = 0 To Len(lst1)
    'If Len(lst2) - 1 < x Then
    'ReDim Preserve result(result.Length)
    'result(result.Length - 1) = lst1(x)
    'Else
    'ReDim Preserve result(result.Length)
    'result(result.Length - 1) = lst1(x) + lst2(x)
    'End If
    'Next
    'Else
    'For x = 0 To Len(lst2)
    'If Len(lst2) - 1 < x Then
    'ReDim Preserve result(result.Length)
    '            result(result.Length - 1) = lst2(x)
    '          Else
    'ReDim Preserve result(result.Length)
    '         result(result.Length - 1) = lst1(x) + lst2(x)
    '        End If
    '      Next
    '  End If
    '  result = standardize(result)
    '  Return result
    ' End Function
    'Function scalePoly(scale, lst)
    '    For x = 0 To Len(lst)
    '       lst(x) = (lst(x) * scale)
    '   Next
    '  lst = standardize(lst)
    '   Return lst
    'End Function
    'Function constCoef(lst)
    '    Dim constant As Integer = lst(0)
    '  Return constant
    'End Function
    'Function shiftLeft(lst)
    'Dim newlst As Integer()
    'ReDim Preserve newlst(newlst.Length)
    'newlst(newlst.Length - 1) = 0
    'For x = 0 To Len(lst)
    'ReDim Preserve newlst(newlst.Length)
    'newlst(newlst.Length - 1) = lst(x)
    'Next
    'Return newlst
    'End Function
    'Function shiftRight(lst)
    'Return lst(1)
    'End Function
    'Function mulPoly(lst1, lst2)
    'If isZeroPoly(lst1) Then
    'Return Nothing
    'Else
    'Return addPoly(mulPoly(shiftRight(lst1), shiftLeft(lst2)),
    '    scalePoly(constCoef(lst1), lst2))
    'End If
    ' End Function
    '-----------------------------------
    Dim a
    Dim b
    Dim c
    Dim root1
    Dim root2
    Dim root3
    Private Sub Compute1_Click(sender As Object, e As EventArgs) Handles Compute1.Click
        Try
            a = ValA.Text
            b = ValB.Text
            c = ValC.Text
            root1 = (b ^ 2 - 4.0 * a * c)
            root2 = (-b + (Math.Sqrt(root1)) / (2 * a))
            root3 = (-b - (Math.Sqrt(root1)) / (2 * a))
            X1.Text = root2
            X2.Text = root3
        Catch ex As Exception
            MsgBox("There is a negative square root")
        End Try
    End Sub

    Private Sub fractionButton_Click(sender As Object, e As EventArgs) Handles fractionButton.Click
        fractionBox2.Clear()
        getValue5()
    End Sub

    Public Function NumberOfDecimalPlaces(ByVal number As Double) As Integer
        Dim numberAsString As String = number.ToString(System.Globalization.CultureInfo.InvariantCulture)
        Dim indexOfDecimalPoint As Integer = numberAsString.IndexOf(".")

        If (indexOfDecimalPoint = -1) Then
            Return 0
        Else
            Return numberAsString.Substring(indexOfDecimalPoint + 1).Length
        End If

    End Function
    Dim numDec
    Dim decToFrac
    Dim bot
    Private Sub decimalButton_Click(sender As Object, e As EventArgs) Handles decimalButton.Click
        decimalBox2.Clear()
        decToFrac = decimalBox1.Text
        numDec = NumberOfDecimalPlaces(decimalBox1.Text)
        bot = 10 ^ numDec
        decToFrac = decToFrac * bot

        Dim Numerator, Denominator, GCF As Integer
        Dim answer As String

        Numerator = Reduce_Numerator(Numerator, Denominator, GCF)
        Denominator = Reduce_Denominator(Numerator, Denominator, GCF)
        answer = String.Concat(Val(decToFrac) & "/" & Val(bot) & " Reduces To: " & Numerator & "/" & Denominator)
        If Val(decToFrac) = Numerator And Val(bot) = Denominator Then
            answer = "Fraction Already Reduced"
        End If
        decimalBox2.Text = answer
    End Sub
    Function Reduce_Numerator(ByVal Numerator As Integer, ByVal Denominator As Integer, ByVal GCF As Integer) As Integer
        Numerator = Val(decToFrac)
        Denominator = Val(bot)
        GCF = Factor(Numerator, Denominator)
        Numerator = Numerator / GCF
        Return Numerator

    End Function

    Function Reduce_Denominator(ByVal Numerator As Integer, ByVal denominator As Integer, ByVal GCF As Integer) As Integer
        Numerator = Val(decToFrac)
        denominator = Val(bot)
        GCF = Factor(Numerator, denominator)
        denominator = denominator / GCF
        Return denominator
    End Function
    Function Factor(ByVal numerator As Integer, ByVal denominator As Integer) As Integer
        Dim temp
        Dim GCF
        While denominator <> 0
            temp = numerator Mod denominator
            numerator = denominator
            denominator = temp
        End While
        GCF = numerator
        Return GCF
    End Function
    Dim answer
    Function PythagTheorem(ByVal a, ByVal b, ByVal c)

        If aSide.Text = "" And bSide.Text <> "" And cSide.Text <> "" Then
            If c < b Then
                answer = "C must be greater than B"
            Else
                answer = Math.Sqrt(c ^ 2 - b ^ 2)
            End If
        ElseIf bSide.Text = "" And aSide.Text <> "" And cSide.Text <> "" Then
            If c < a Then
                answer = "C must be greater than A"
            Else
                answer = Math.Sqrt(c ^ 2 - a ^ 2)
            End If
        ElseIf cSide.Text = "" And bSide.Text <> "" And aSide.Text <> "" Then
            answer = Math.Sqrt(a ^ 2 + b ^ 2)
        Else
            answer = "Enter two values"
        End If
        Return answer
    End Function

    Private Sub calcButton_Click(sender As Object, e As EventArgs) Handles calcButton.Click
        Result.Clear()
        Result.Text = PythagTheorem(aSide.Text, bSide.Text, cSide.Text)
    End Sub
End Class