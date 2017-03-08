Module Module1
    Public lastWasArith = False 'This boolean is used so the first number 
    'still displays when an arithmetic is used


    Public lastWasArith2 = False

    Sub ShowValue(ByVal Butt As Button)
        If lastWasArith = True Then
            Form1.TextBox1.Text = ""
            Scientific.TextBox1.Text = ""
            lastWasArith = False
        End If
        If Form1.lastWasEqual Then
            Form1.TextBox2.Text = ""
            Form1.TextBox2.Text = Form1.TextBox1.Text
        End If
        If Form1.TextBox1.Text = "0" Or Form1.contNum = True Or Form1.lastWasEqual = True Then
            Form1.TextBox1.Text = Butt.Text
            If Form1.TextBox1.Text = "0" Then
                Form1.TextBox2.Text = Butt.Text
            Else
                Form1.TextBox2.Text += Butt.Text
            End If
            Form1.contNum = False
        Else
            Form1.TextBox1.Text = (Form1.TextBox1.Text & Butt.Text)
            Form1.TextBox2.Text = (Form1.TextBox2.Text & Butt.Text)
        End If
    End Sub

    Sub ShowValue2(ByVal Butt As Button)
        If lastWasArith2 = True Then
            Scientific.TextBox1.Text = ""
            lastWasArith2 = False
        End If
        If Scientific.lastWasEqual2 Then
            Scientific.TextBox2.Text = ""
            Scientific.TextBox2.Text = Scientific.TextBox1.Text
        End If
        If Scientific.TextBox1.Text = "0" Or Scientific.contNum = True Or Scientific.lastWasEqual2 = True Then
            Scientific.TextBox1.Text = Butt.Text
            If Scientific.TextBox2.Text = "0" Then
                Scientific.TextBox2.Text = Butt.Text
            Else
                Scientific.TextBox2.Text += Butt.Text
            End If
            Scientific.contNum = False
        Else
            Scientific.TextBox1.Text = (Scientific.TextBox1.Text & Butt.Text)
            Scientific.TextBox2.Text = (Scientific.TextBox2.Text & Butt.Text)
        End If

    End Sub

    Sub Arithematic(ByVal butt As Button)
        If Form1.lastWasEqual Then
            Form1.TextBox2.Text = ""
            Form1.TextBox2.Text = Form1.TextBox1.Text
        End If
        Form1.lastWasEqual = False
        Form1.Value1 = Val(Form1.TextBox1.Text)
        Form1.Oper = butt.Text
        Form1.TextBox1.Text = ""
        Form1.TextBox1.Text = Form1.Value1
        Form1.TextBox2.Text += Form1.Oper
        lastWasArith = True
    End Sub

    Sub Arithematic2(ByVal butt As Button)
        If Scientific.lastWasEqual2 Then
            Scientific.TextBox2.Text = ""
            Scientific.TextBox2.Text = Scientific.TextBox1.Text
        End If
        Scientific.lastWasEqual2 = False
        Scientific.Value1 = Val(Scientific.TextBox1.Text)
        Scientific.Oper = butt.Text
        Scientific.TextBox1.Text = ""
        Scientific.TextBox1.Text = Scientific.Value1
        Scientific.TextBox2.Text += Scientific.Oper

        lastWasArith2 = True
    End Sub

    Sub Calculate()
        If Form1.contOper = 2 Then
            Select Case Form1.Oper
                Case "+"
                    Form1.TextBox1.Text = Form1.Value3 + Val(Form1.TextBox1.Text)
                Case "-"
                    Form1.TextBox1.Text = Form1.Value3 - Val(Form1.TextBox1.Text)
                Case "×"
                    Form1.TextBox1.Text = Form1.Value3 * Val(Form1.TextBox1.Text)
                Case "÷"
                    Form1.TextBox1.Text = Form1.Value3 / Val(Form1.TextBox1.Text)
            End Select
        Else
            Select Case Form1.Oper
                Case "+"
                    Form1.TextBox1.Text = Form1.Value1 + Val(Form1.TextBox1.Text)
                Case "-"
                    Form1.TextBox1.Text = Form1.Value1 - Val(Form1.TextBox1.Text)
                Case "×"
                    Form1.TextBox1.Text = Form1.Value1 * Val(Form1.TextBox1.Text)
                Case "÷"
                    Form1.TextBox1.Text = Form1.Value1 / Val(Form1.TextBox1.Text)
            End Select
        End If
    End Sub
    Sub Calculate2()
        If Scientific.contOper = 2 Then
            Select Case Scientific.Oper
                Case "+"
                    Scientific.TextBox1.Text = Scientific.Value3 + Val(Scientific.TextBox1.Text)
                Case "-"
                    Scientific.TextBox1.Text = Scientific.Value3 - Val(Scientific.TextBox1.Text)
                Case "×"
                    Scientific.TextBox1.Text = Scientific.Value3 * Val(Scientific.TextBox1.Text)
                Case "÷"
                    Scientific.TextBox1.Text = Scientific.Value3 / Val(Scientific.TextBox1.Text)
                Case "Mod"
                    Scientific.TextBox1.Text = Scientific.Value3 Mod Val(Scientific.TextBox1.Text)
                Case "xʸ"
                    Scientific.TextBox1.Text = Scientific.Value3 ^ Val(Scientific.TextBox1.Text)
                Case "Exp"
                    Scientific.TextBox1.Text = Scientific.Value3 * (10 ^ Val(Scientific.TextBox1.Text))
                Case "ʸ√ x"
                    Scientific.TextBox1.Text = Scientific.Value3 ^ (1 / Val(Scientific.TextBox1.Text))
            End Select
        Else
            Select Case Scientific.Oper
                Case "+"
                    Scientific.TextBox1.Text = Scientific.Value1 + Val(Scientific.TextBox1.Text)
                Case "-"
                    Scientific.TextBox1.Text = Scientific.Value1 - Val(Scientific.TextBox1.Text)
                Case "×"
                    Scientific.TextBox1.Text = Scientific.Value1 * Val(Scientific.TextBox1.Text)
                Case "÷"
                    Scientific.TextBox1.Text = Scientific.Value1 / Val(Scientific.TextBox1.Text)
                Case "Mod"
                    Scientific.TextBox1.Text = Scientific.Value1 Mod Val(Scientific.TextBox1.Text)
                Case "xʸ"
                    Scientific.TextBox1.Text = Scientific.Value1 ^ Val(Scientific.TextBox1.Text)
                Case "Exp"
                    Scientific.TextBox1.Text = Scientific.Value1 * (10 ^ Val(Scientific.TextBox1.Text))
                Case "ʸ√ x"
                    Scientific.TextBox1.Text = Scientific.Value1 ^ (1 / Val(Scientific.TextBox1.Text))
            End Select
        End If
    End Sub
    Sub ShowHistory()
        Form1.History.Text =
            Form1.History.Text &
            Form1.TextBox2.Text & vbCrLf &
            "---------------------------------------------------" & vbCrLf &
            Form1.TextBox1.Text & vbCrLf &
            "---------------------------------------------------" & vbCrLf
    End Sub
    Sub ShowHistory2()
        Scientific.History.Text =
            Scientific.History.Text &
            Scientific.TextBox2.Text & vbCrLf &
            "---------------------------------------------------" & vbCrLf &
            Scientific.TextBox1.Text & vbCrLf &
            "---------------------------------------------------" & vbCrLf & vbCrLf
    End Sub
    Sub SayIt()
        Dim Say
        Say = CreateObject("sapi.spvoice")
        Say.speak(Form1.TextBox1.Text)
    End Sub
    Sub SayIt2()
        Dim Say
        Say = CreateObject("sapi.spvoice")
        Say.speak(Scientific.TextBox1.Text)
    End Sub

    'Reverse Polish Notation + Shunting yard - to calculate
    '---------------------------------------------------------------------------------
    Sub getValue()
        Dim source As String = Scientific.TextBox2.Text
        Dim format As String = FormatExpression(source)
        Dim rpn() As Token = ShuntingYardAlgorithm(Scan(format))
        Scientific.TextBox1.Text = Evaluate(rpn)
    End Sub
    Private Function FormatExpression(ByVal expression As String) As String
        Dim format As String = expression.Replace(" ", String.Empty)

        format = format.Replace("-", "- ")
        format = format.Replace("(", " ( ").Replace(")", " ) ")


        Dim unaryEvaluator As System.Text.RegularExpressions.MatchEvaluator = New System.Text.RegularExpressions.MatchEvaluator(AddressOf ReplaceUnary)
        format = System.Text.RegularExpressions.Regex.Replace(format, "(\+|-|\*|\\|\^)-", unaryEvaluator)

        Dim digitEvaluator As System.Text.RegularExpressions.MatchEvaluator = New System.Text.RegularExpressions.MatchEvaluator(AddressOf ReplaceDigits)
        format = System.Text.RegularExpressions.Regex.Replace(format, "(-?[0-9]+(?:\.[0-9]*)?)", digitEvaluator)

        format = System.Text.RegularExpressions.Regex.Replace(format, " {2,}", " ")

        format = format.Trim()
        format = format.Replace("- - ", "- -")

        Return format
    End Function

    Private Function ReplaceUnary(ByVal m As System.Text.RegularExpressions.Match) As String
        Return " " & m.Value
    End Function

    Private Function ReplaceDigits(ByVal m As System.Text.RegularExpressions.Match) As String
        Return " " & m.Value & " "
    End Function


    Private Function AddNumbers(ByVal operand1 As Double, ByVal operand2 As Double) As Double
        Return operand1 + operand2
    End Function
    Private Function SubtractNumbers(ByVal operand1 As Double, ByVal operand2 As Double) As Double
        Return operand1 - operand2
    End Function
    Private Function MultiplyNumbers(ByVal operand1 As Double, ByVal operand2 As Double) As Double
        Return operand1 * operand2
    End Function
    Private Function DivideNumbers(ByVal operand1 As Double, ByVal operand2 As Double) As Double
        Return operand1 / operand2
    End Function
    Private Function ModuloNumbers(ByVal operand1 As Double, ByVal operand2 As Double) As Double
        Return operand1 Mod operand2
    End Function
    Private Function RaiseNumbers(ByVal operand1 As Double, ByVal operand2 As Double) As Double
        Return operand1 ^ operand2
    End Function

    Private Function Scan(ByVal source As String) As Token()
        'Rule 1: Left Parenthesis
        'Rule 2: Right Parenthesis
        'Rule 3: Doubles
        'Rules 4 - 5: Addition and Subtraction, precedence: 1
        'Rules 6 - 8: Multiplication, Division, and MOD, precedence: 2
        'Rule 9: Exponent, precedence: 3
        Dim definitions() As Token = {
            New Token() With {.Pattern = "^\($", .Type = Token.TokenType.LeftParenthesis, .Value = String.Empty},
            New Token() With {.Pattern = "^\)$", .Type = Token.TokenType.RightParenthesis, .Value = String.Empty},
            New Token() With {.Pattern = "^([-+]?(\d*[.])?\d+)$", .Type = Token.TokenType.Digit, .Value = String.Empty},
            New Token() With {.Operation = New Token.MathOperation(AddressOf AddNumbers), .Pattern = "^\+$", .Precedence = 1, .Type = Token.TokenType.Operator, .Value = String.Empty},
            New Token() With {.Operation = New Token.MathOperation(AddressOf SubtractNumbers), .Pattern = "^\-$", .Precedence = 1, .Type = Token.TokenType.Operator, .Value = String.Empty},
            New Token() With {.Operation = New Token.MathOperation(AddressOf MultiplyNumbers), .Pattern = "^×$", .Precedence = 2, .Type = Token.TokenType.Operator, .Value = String.Empty},
            New Token() With {.Operation = New Token.MathOperation(AddressOf DivideNumbers), .Pattern = "^\÷$", .Precedence = 2, .Type = Token.TokenType.Operator, .Value = String.Empty},
            New Token() With {.Operation = New Token.MathOperation(AddressOf ModuloNumbers), .Pattern = "^Mod$", .Precedence = 2, .Type = Token.TokenType.Operator, .Value = String.Empty},
            New Token() With {.Operation = New Token.MathOperation(AddressOf RaiseNumbers), .Pattern = "^\^$", .Precedence = 3, .Type = Token.TokenType.Operator, .Value = String.Empty}
        }

        Dim tokens As List(Of Token) = New List(Of Token)
        For Each item As String In source.Split({" "}, StringSplitOptions.RemoveEmptyEntries)
            Dim currentToken As Token = Nothing
            Dim regex As Text.RegularExpressions.Regex
            For Each definition As Token In definitions
                regex = New Text.RegularExpressions.Regex(definition.Pattern)
                If regex.IsMatch(item) Then
                    currentToken = New Token With {.Operation = definition.Operation, .Pattern = definition.Pattern, .Precedence = definition.Precedence, .Type = definition.Type, .Value = item}
                    tokens.Add(currentToken)
                    Exit For
                End If
            Next
        Next
        Return tokens.ToArray
    End Function
    Private Function ShuntingYardAlgorithm(ByVal tokens() As Token) As Token()
        Dim output As List(Of Token) = New List(Of Token)
        Dim operatorStack As Stack(Of Token) = New Stack(Of Token)
        For Each item As Token In tokens
            If item.Type = Token.TokenType.Digit Then
                output.Add(item)
            ElseIf item.Type = Token.TokenType.Operator Then
                While operatorStack.Count > 0 AndAlso item.Precedence <= operatorStack.Peek.Precedence
                    output.Add(operatorStack.Pop)
                End While
                operatorStack.Push(item)
            ElseIf item.Type = Token.TokenType.LeftParenthesis Then
                operatorStack.Push(item)
            ElseIf item.Type = Token.TokenType.RightParenthesis Then
                Dim flag As Boolean = False
                While operatorStack.Count > 0 AndAlso flag = False
                    If operatorStack.Peek.Type = Token.TokenType.Operator Then
                        output.Add(operatorStack.Pop)
                    ElseIf operatorStack.Peek.Type = Token.TokenType.LeftParenthesis Then
                        operatorStack.Pop()
                        flag = True
                    End If
                End While
            End If
        Next
        If operatorStack.Count > 0 Then
            Do
                operatorStack.Peek.Type = Token.TokenType.Operator
                output.Add(operatorStack.Pop)

            Loop Until operatorStack.Count = 0
        End If
        Return output.ToArray()
    End Function
    Private Function Evaluate(ByVal tokens() As Token) As Double
        Dim valueStack As Stack(Of Token) = New Stack(Of Token)
        For Each item As Token In tokens
            If item.Type = Token.TokenType.Operator Then
                Dim operand2 As Token = valueStack.Pop
                Dim operand1 As Token = valueStack.Pop
                valueStack.Push(New Token With {.Type = Token.TokenType.Digit, .Value = item.Operation(Double.Parse(operand1.Value), Double.Parse(operand2.Value)).ToString})
            Else
                valueStack.Push(item)
            End If
        Next
        Return Double.Parse(valueStack.Pop.Value)
    End Function
End Module

Public Class Token
    Public Enum TokenType
        Digit
        LeftParenthesis
        [Operator]
        RightParenthesis
    End Enum
    Public Delegate Function MathOperation(ByVal value1 As Double, ByVal value2 As Double) As Double
    Private operat As MathOperation
    Public Property Operation() As MathOperation
        Get
            Return operat
        End Get
        Set(ByVal value As MathOperation)
            operat = value
        End Set
    End Property
    Private pat As String
    Public Property Pattern() As String
        Get
            Return pat
        End Get
        Set(ByVal value As String)
            pat = value
        End Set
    End Property
    Private preced As Integer
    Public Property Precedence() As Integer
        Get
            Return preced
        End Get
        Set(ByVal value As Integer)
            preced = value
        End Set
    End Property
    Private t As TokenType
    Public Property Type() As TokenType
        Get
            Return t
        End Get
        Set(ByVal value As TokenType)
            t = value
        End Set
    End Property
    Private val As String
    Public Property Value() As String
        Get
            Return val
        End Get
        Set(ByVal value As String)
            val = value
        End Set
    End Property
End Class
