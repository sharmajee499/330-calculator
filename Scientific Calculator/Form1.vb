Public Class Form1
    Public Value1, Value2, Value3, Mem As Double
    Public Oper As Char
    Public contOper As Integer = 0
    Public halt As Boolean = False
    Public contNum As Boolean = False
    Public lastWasEqual = False

    Private Sub StandardToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles StandardToolStripMenuItem.Click
        Me.Show()
    End Sub

    Private Sub ScientificToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ScientificToolStripMenuItem.Click
        Scientific.Show()
        Me.Hide()
    End Sub

    Private Sub Btn7_Click(sender As Object, e As EventArgs) Handles Btn7.Click
        ShowValue(Btn7)
        halt = False
    End Sub

    Private Sub Btn8_Click(sender As Object, e As EventArgs) Handles Btn8.Click
        ShowValue(Btn8)
        halt = False
    End Sub

    Private Sub Btn0_Click(sender As Object, e As EventArgs) Handles Btn0.Click
        ShowValue(Btn0)
        halt = False
    End Sub

    Private Sub Btn1_Click(sender As Object, e As EventArgs) Handles Btn1.Click
        ShowValue(Btn1)
        halt = False
    End Sub

    Private Sub Btn2_Click(sender As Object, e As EventArgs) Handles Btn2.Click
        ShowValue(Btn2)
        halt = False
    End Sub

    Private Sub Btn3_Click(sender As Object, e As EventArgs) Handles Btn3.Click
        ShowValue(Btn3)
        halt = False
    End Sub

    Private Sub Btn4_Click(sender As Object, e As EventArgs) Handles Btn4.Click
        ShowValue(Btn4)
        halt = False
    End Sub

    Private Sub Btn5_Click(sender As Object, e As EventArgs) Handles Btn5.Click
        ShowValue(Btn5)
        halt = False
    End Sub

    Private Sub Btn6_Click(sender As Object, e As EventArgs) Handles Btn6.Click
        ShowValue(Btn6)
        halt = False
    End Sub

    Private Sub Btn9_Click(sender As Object, e As EventArgs) Handles Btn9.Click
        ShowValue(Btn9)
        halt = False
    End Sub

    Private Sub BtnDiv_Click(sender As Object, e As EventArgs) Handles BtnDiv.Click
        If contOper = 0 Then
            Arithematic(BtnDiv)
            contOper = 1
            contNum = True
            halt = True
        ElseIf contOper = 1 Then
            If halt = False Then
                Calculate()
                Arithematic(BtnDiv)
                Value3 = TextBox1.Text
                contOper = 2
                contNum = True
                halt = True
            End If
        Else
            If halt = False Then
                Calculate()
                Arithematic(BtnDiv)
                Value3 = TextBox1.Text
                contNum = True
                halt = True
            End If
        End If
    End Sub

    Private Sub BtnMultiply_Click(sender As Object, e As EventArgs) Handles BtnMultiply.Click
        If contOper = 0 Then
            Arithematic(BtnMultiply)
            contOper = 1
            contNum = True
            halt = True
        ElseIf contOper = 1 Then
            If halt = False Then
                Value2 = Val(TextBox1.Text)
                Calculate()
                Arithematic(BtnMultiply)
                Value3 = TextBox1.Text
                contOper = 2
                contNum = True
                halt = True
            End If
        Else
            If halt = False Then
                Calculate()
                Arithematic(BtnMultiply)
                Value3 = TextBox1.Text
                contNum = True
                halt = True
            End If
        End If
    End Sub

    Private Sub BtnMinus_Click(sender As Object, e As EventArgs) Handles BtnMinus.Click
        If contOper = 0 Then
            Arithematic(BtnMinus)
            contOper = 1
            contNum = True
            halt = True
        ElseIf contOper = 1 Then
            If halt = False Then
                Value2 = Val(TextBox1.Text)
                Calculate()
                Arithematic(BtnMinus)
                Value3 = TextBox1.Text
                contOper = 2
                contNum = True
                halt = True
            End If
        Else
            If halt = False Then
                Calculate()
                Arithematic(BtnMinus)
                Value3 = TextBox1.Text
                contNum = True
                halt = True
            End If
        End If
    End Sub

    Private Sub BtnPlus_Click(sender As Object, e As EventArgs) Handles BtnPlus.Click
        If contOper = 0 Then
            Arithematic(BtnPlus)
            contOper = 1
            contNum = True
            halt = True
        ElseIf contOper = 1 Then
            If halt = False Then
                Value2 = Val(TextBox1.Text)

                Calculate()
                Arithematic(BtnPlus)
                Value3 = TextBox1.Text
                contOper = 2
                contNum = True
                halt = True
            End If
        Else
            If halt = False Then
                Calculate()
                Arithematic(BtnPlus)
                Value3 = TextBox1.Text
                contNum = True
                halt = True
            End If
        End If
    End Sub

    Private Sub BtnEqual_Click(sender As Object, e As EventArgs) Handles BtnEqual.Click
        Value2 = Val(TextBox1.Text)
        getValue2()
        ShowHistory()
        lastWasEqual = True
        Value2 = 0
        contOper = 0
    End Sub

    Private Sub BtnDot_Click(sender As Object, e As EventArgs) Handles BtnDot.Click
        If Not TextBox1.Text.Contains(".") Then
            TextBox1.Text += "."
        End If
    End Sub

    Private Sub BtnHistory_Click(sender As Object, e As EventArgs) Handles BtnHistory.Click
        If Me.Width = 318 Then
            Me.Width = 601
        Else
            Me.Width = 318
        End If
    End Sub

    Private Sub BtnSqrRoot_Click(sender As Object, e As EventArgs) Handles BtnSqrRoot.Click
        TextBox1.Text = Math.Sqrt(Val(TextBox1.Text))
    End Sub

    Private Sub BtnFraction_Click(sender As Object, e As EventArgs) Handles BtnFraction.Click
        TextBox1.Text = 1 / Val(TextBox1.Text)
    End Sub

    Private Sub BtnPlusMinus_Click(sender As Object, e As EventArgs) Handles BtnPlusMinus.Click
        TextBox1.Text = Val(TextBox1.Text) * -1
    End Sub

    Private Sub BtnDel_Click(sender As Object, e As EventArgs) Handles BtnDel.Click
        If Len(TextBox1.Text) > 0 Then
            TextBox1.Text = Mid(TextBox1.Text, 1, Len(TextBox1.Text) - 1)
            TextBox2.Text = Mid(TextBox2.Text, 1, Len(TextBox2.Text) - 1)
        ElseIf Len(TextBox2.Text) > 0 Then
            TextBox2.Text = Mid(TextBox2.Text, 1, Len(TextBox2.Text) - 1)
        End If
    End Sub

    Private Sub BtnMPlus_Click(sender As Object, e As EventArgs) Handles BtnMPlus.Click
        Mem = Mem + Val(TextBox1.Text)
    End Sub

    Private Sub btnC_Click(sender As Object, e As EventArgs) Handles btnC.Click
        TextBox1.Text = 0
        TextBox2.Text = 0
        contOper = 0
        Value1 = 0
        Value2 = 0
        Value3 = 0
        halt = False
        contNum = False
        lastWasEqual = False
    End Sub

    Private Sub BtnMC_Click(sender As Object, e As EventArgs) Handles BtnMC.Click
        Mem = 0
    End Sub

    Private Sub BtnMR_Click(sender As Object, e As EventArgs) Handles BtnMR.Click
        TextBox1.Text = Mem
    End Sub

    Private Sub BtnMMinus_Click(sender As Object, e As EventArgs) Handles BtnMMinus.Click
        Mem = Mem - Val(TextBox1.Text)
    End Sub

    Private Sub btnCE_Click(sender As Object, e As EventArgs) Handles btnCE.Click
        TextBox1.Text = 0
        TextBox2.Text = 0
        Value1 = 0
        Value2 = 0
        Value3 = 0
        halt = False
        contNum = False
        lastWasEqual = False
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
        Functions.Show()
        Me.Hide()
    End Sub

    Private Sub ConverterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConverterToolStripMenuItem.Click
        Converter.Show()
        Me.Hide()
    End Sub

    Private Sub BtnXSqr_Click(sender As Object, e As EventArgs) Handles BtnXSqr.Click
        TextBox1.Text = Val(TextBox1.Text) ^ 2
        TextBox2.Text += "^2"
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Mem = 0
    End Sub

    Private Sub btnPercent_Click(sender As Object, e As EventArgs) Handles btnPercent.Click
        TextBox1.Text = Value1 * (Val(TextBox1.Text) / 100)
        TextBox2.Text += "%"
    End Sub
End Class
