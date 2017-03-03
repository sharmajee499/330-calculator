Public Class Form1
    Public Value1, Value2, Mem As Double
    Public Oper As Char
    Private Sub StandardToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles StandardToolStripMenuItem.Click
        Me.Show()
    End Sub

    Private Sub ScientificToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ScientificToolStripMenuItem.Click
        Scientific.Show()
        Me.Hide()
    End Sub

    Private Sub Btn7_Click(sender As Object, e As EventArgs) Handles Btn7.Click
        ShowValue(Btn7)
    End Sub

    Private Sub Btn8_Click(sender As Object, e As EventArgs) Handles Btn8.Click
        ShowValue(Btn8)
    End Sub

    Private Sub Btn0_Click(sender As Object, e As EventArgs) Handles Btn0.Click
        ShowValue(Btn0)
    End Sub

    Private Sub Btn1_Click(sender As Object, e As EventArgs) Handles Btn1.Click
        ShowValue(Btn1)
    End Sub

    Private Sub Btn2_Click(sender As Object, e As EventArgs) Handles Btn2.Click
        ShowValue(Btn2)
    End Sub

    Private Sub Btn3_Click(sender As Object, e As EventArgs) Handles Btn3.Click
        ShowValue(Btn3)
    End Sub

    Private Sub Btn4_Click(sender As Object, e As EventArgs) Handles Btn4.Click
        ShowValue(Btn4)
    End Sub

    Private Sub Btn5_Click(sender As Object, e As EventArgs) Handles Btn5.Click
        ShowValue(Btn5)
    End Sub

    Private Sub Btn6_Click(sender As Object, e As EventArgs) Handles Btn6.Click
        ShowValue(Btn6)
    End Sub

    Private Sub Btn9_Click(sender As Object, e As EventArgs) Handles Btn9.Click
        ShowValue(Btn9)
    End Sub

    Private Sub BtnDiv_Click(sender As Object, e As EventArgs) Handles BtnDiv.Click
        Arithematic(BtnDiv)
    End Sub

    Private Sub BtnMultiply_Click(sender As Object, e As EventArgs) Handles BtnMultiply.Click
        Arithematic(BtnMultiply)
    End Sub

    Private Sub BtnMinus_Click(sender As Object, e As EventArgs) Handles BtnMinus.Click
        Arithematic(BtnMinus)
    End Sub

    Private Sub BtnPlus_Click(sender As Object, e As EventArgs) Handles BtnPlus.Click
        Arithematic(BtnPlus)
    End Sub

    Private Sub BtnEqual_Click(sender As Object, e As EventArgs) Handles BtnEqual.Click
        Value2 = Val(TextBox1.Text)
        Calculate()
        ShowHistory()
        Value2 = 0
        SayIt()
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
        TextBox1.Text = Val(TextBox1.Text) \ 10
    End Sub

    Private Sub BtnMPlus_Click(sender As Object, e As EventArgs) Handles BtnMPlus.Click
        Mem = Mem + Val(TextBox1.Text)
    End Sub

    Private Sub btnC_Click(sender As Object, e As EventArgs) Handles btnC.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
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
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Mem = 0
    End Sub

    Private Sub btnPercent_Click(sender As Object, e As EventArgs) Handles btnPercent.Click
        TextBox1.Text = Value1 * (Val(TextBox1.Text) / 100)
    End Sub
End Class
