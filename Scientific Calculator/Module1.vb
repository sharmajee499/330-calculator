Module Module1
    Public lastWasArith = False 'This boolean is used so the first number 
    'still displays when an arithmetic is used
    Public lastWasEqual = False

    Public lastWasArith2 = False
    Public lastWasEqual2 = False

    Sub ShowValue(ByVal Butt As Button)
        If lastWasArith = True Then
            Form1.TextBox1.Text = ""
            Scientific.TextBox1.Text = ""
            lastWasArith = False
        End If
        If lastWasEqual Then
            Form1.TextBox2.Text = ""
            Form1.TextBox2.Text = Form1.TextBox1.Text
        End If
        If Form1.TextBox1.Text = "0" Then
            Form1.TextBox1.Text = Butt.Text
            Form1.TextBox2.Text = Butt.Text
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
        If lastWasEqual2 Then
            Scientific.TextBox2.Text = ""
            Scientific.TextBox2.Text = Scientific.TextBox1.Text
        End If
        If Scientific.TextBox1.Text = "0" Then
            Scientific.TextBox1.Text = Butt.Text
            Scientific.TextBox2.Text = Butt.Text
        Else
            Scientific.TextBox1.Text = (Scientific.TextBox1.Text & Butt.Text)
            Scientific.TextBox2.Text = (Scientific.TextBox2.Text & Butt.Text)
        End If

    End Sub

    Sub Arithematic(ByVal butt As Button)
        If lastWasEqual Then
            Form1.TextBox2.Text = ""
            Form1.TextBox2.Text = Form1.TextBox1.Text
        End If
        lastWasEqual = False
        Form1.Value1 = Val(Form1.TextBox1.Text)
        Form1.Oper = butt.Text
        Form1.TextBox1.Text = ""
        Form1.TextBox1.Text = Form1.Value1
        Form1.TextBox2.Text += Form1.Oper
        lastWasArith = True
    End Sub

    Sub Arithematic2(ByVal butt As Button)
        If lastWasEqual2 Then
            Scientific.TextBox2.Text = ""
            Scientific.TextBox2.Text = Scientific.TextBox1.Text
        End If
        lastWasEqual2 = False
        Scientific.Value1 = Val(Scientific.TextBox1.Text)
        Scientific.Oper = butt.Text
        Scientific.TextBox1.Text = ""
        Scientific.TextBox1.Text = Scientific.Value1
        Scientific.TextBox2.Text += Scientific.Oper

        lastWasArith2 = True
    End Sub

    Sub Calculate()
        lastWasEqual = True
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
    End Sub
    Sub Calculate2()
        lastWasEqual2 = True
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
    End Sub
    Sub ShowHistory()
        Form1.History.Text =
            Form1.History.Text &
            Form1.Value1 & vbCrLf &
            Form1.Oper & vbCrLf &
            Form1.Value2 & vbCrLf &
            "---------------------------------------------------" & vbCrLf &
            Form1.TextBox1.Text & vbCrLf &
            "---------------------------------------------------" & vbCrLf
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
End Module
