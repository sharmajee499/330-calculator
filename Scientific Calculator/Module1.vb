Module Module1
    Sub ShowValue(ByVal Butt As Button)
        Form1.TextBox1.Text = (Form1.TextBox1.Text & Butt.Text)
    End Sub

    Sub Arithematic(ByVal butt As Button)
        Form1.Value1 = Val(Form1.TextBox1.Text)
        Form1.Oper = butt.Text
        Form1.TextBox1.Text = ""
    End Sub
    Sub Calculate()
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
End Module
