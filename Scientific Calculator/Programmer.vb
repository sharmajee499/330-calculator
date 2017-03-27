Public Class Programmer
    Public Value1, Value2, Value3, Mem As Double
    Public Oper As Char
    Public contOper As Integer = 0
    Public halt As Boolean = False
    Public contNum As Boolean = False
    Public lastWasEqual3 = False
    Public format As String = "dec" 'decimal
    Public BinValue3 As Integer = 0 'Can't convert to binary with a double
    Public DecString As String 'Store equations in decimal form


    Private Sub Btn7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        ShowValue3(Button7)
        halt = False
    End Sub

    Private Sub Btn8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        ShowValue3(Button8)
        halt = False
    End Sub

    Private Sub Btn0_Click(sender As Object, e As EventArgs) Handles Button10.Click
        ShowValue3(Button10) 'Tenth button was 0
        halt = False
    End Sub

    Private Sub Btn1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ShowValue3(Button1)
        halt = False
    End Sub

    Private Sub Btn2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ShowValue3(Button2)
        halt = False
    End Sub

    Private Sub Btn3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ShowValue3(Button3)
        halt = False
    End Sub

    Private Sub Btn4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ShowValue3(Button4)
        halt = False
    End Sub

    Private Sub Btn5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ShowValue3(Button5)
        halt = False
    End Sub

    Private Sub Btn6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ShowValue3(Button6)
        halt = False
    End Sub

    Private Sub Btn9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        ShowValue3(Button9)
        halt = False
    End Sub
    Private Sub BtnDiv_Click(sender As Object, e As EventArgs) Handles BtnDiv.Click
        If contOper = 0 Then
            Arithematic3(BtnDiv)
            contOper = 1
            contNum = True
            halt = True
        ElseIf contOper = 1 Then
            If halt = False Then
                Calculate3()
                Arithematic3(BtnDiv)
                Value3 = TextBox1.Text
                contOper = 2
                contNum = True
                halt = True
            End If
        Else
            If halt = False Then
                Calculate3()
                Arithematic3(BtnDiv)
                Value3 = TextBox1.Text
                contNum = True
                halt = True
            End If
        End If
    End Sub

    Private Sub BtnMultiply_Click(sender As Object, e As EventArgs) Handles BtnMultiply.Click
        If contOper = 0 Then
            Arithematic3(BtnMultiply)
            contOper = 1
            contNum = True
            halt = True
        ElseIf contOper = 1 Then
            If halt = False Then
                Value2 = Val(TextBox1.Text)
                Calculate3()
                Arithematic3(BtnMultiply)
                Value3 = TextBox1.Text
                contOper = 2
                contNum = True
                halt = True
            End If
        Else
            If halt = False Then
                Calculate3()
                Arithematic3(BtnMultiply)
                Value3 = TextBox1.Text
                contNum = True
                halt = True
            End If
        End If
    End Sub

    Private Sub BtnMinus_Click(sender As Object, e As EventArgs) Handles BtnMinus.Click
        If contOper = 0 Then
            Arithematic3(BtnMinus)
            contOper = 1
            contNum = True
            halt = True
        ElseIf contOper = 1 Then
            If halt = False Then
                Value2 = Val(TextBox1.Text)
                Calculate3()
                Arithematic3(BtnMinus)
                Value3 = TextBox1.Text
                contOper = 2
                contNum = True
                halt = True
            End If
        Else
            If halt = False Then
                Calculate3()
                Arithematic3(BtnMinus)
                Value3 = TextBox1.Text
                contNum = True
                halt = True
            End If
        End If
    End Sub

    Private Sub BtnPlus_Click(sender As Object, e As EventArgs) Handles BtnPlus.Click
        If contOper = 0 Then
            Arithematic3(BtnPlus)
            contOper = 1
            contNum = True
            halt = True
        ElseIf contOper = 1 Then
            If halt = False Then
                Value2 = Val(TextBox1.Text)

                Calculate3()
                Arithematic3(BtnPlus)
                Value3 = TextBox1.Text
                contOper = 2
                contNum = True
                halt = True
            End If
        Else
            If halt = False Then
                Calculate3()
                Arithematic3(BtnPlus)
                Value3 = TextBox1.Text
                contNum = True
                halt = True
            End If
        End If
    End Sub
    Private Sub BtnEqual_Click(sender As Object, e As EventArgs) Handles BtnEqual.Click
        Try
            Value2 = Val(TextBox1.Text)
            If format.Equals("bin") Then
                Value1 = Convert.ToInt32(Value1, 2)
                Value2 = Convert.ToInt32(Value2, 2)
            End If
            DecString = Value1 & Oper & Value2
            'MsgBox(DecString)
            getValue3()
            'MsgBox(Value1)
            'MsgBox(Value2)
            'MsgBox(Value3)
            If format.Equals("bin") Then
                BinValue3 = Int(TextBox1.Text)
                BinValue3 = Convert.ToString(BinValue3, 2)
                BinValue3 = Convert.ToInt32(BinValue3)
                TextBox1.Text = BinValue3
            End If
            'MsgBox(BinValue3)
            ShowHistory()
            lastWasEqual3 = True
            Value2 = 0
            contOper = 0
        Catch ex As Exception
            MsgBox("Invalid Operation")
        End Try
    End Sub
    Private Sub StandardToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles StandardToolStripMenuItem.Click
        Form1.Show()
        Me.Hide()
    End Sub

    Private Sub ScientificToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ScientificToolStripMenuItem.Click
        Scientific.Show()
        Me.Hide()
    End Sub
    Private Sub ProgrammerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProgrammerToolStripMenuItem.Click
        Me.Show()
        Me.Hide()
    End Sub

    Private Sub Decimal_Click(sender As Object, e As EventArgs) Handles Dec.Click

    End Sub
    Private Sub Octal_Click(sender As Object, e As EventArgs) Handles Octal.Click

    End Sub
    Private Sub Hex_Click(sender As Object, e As EventArgs) Handles Hex.Click

    End Sub
    Private Sub Binary_Click(sender As Object, e As EventArgs) Handles Binary.Click
        format = "bin" 'binary
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
        Button7.Enabled = False
        Button8.Enabled = False
        Button9.Enabled = False
    End Sub

    Private Sub GraphingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GraphingToolStripMenuItem.Click
        Graphing.Show()
        Me.Hide()
    End Sub

    Private Sub FunctionsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FunctionsToolStripMenuItem.Click
        Functions.Show()
        Me.Hide()
    End Sub

    Private Sub ConverterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConverterToolStripMenuItem.Click
        Converter.Show()
        Me.Hide()
    End Sub
End Class