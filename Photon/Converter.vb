Public Class Converter
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
        Functions.Show()
        Me.Hide()
    End Sub

    Private Sub ConverterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConverterToolStripMenuItem.Click
        Me.Show()
    End Sub

    'Volume Converter'
    Private Sub BtnVolumeConvert_Click(sender As Object, e As EventArgs) Handles BtnVolumeConvert.Click
        'Milliliters'
        If LstVolumeFrom.SelectedIndex = 0 And LstVolumeTo.SelectedIndex = 0 Then
            TxtVolumeTo.Text = TxtVolumeFrom.Text * 1

        ElseIf LstVolumeFrom.SelectedIndex = 0 And LstVolumeTo.SelectedIndex = 1 Then
            TxtVolumeTo.Text = TxtVolumeFrom.Text * 1

        ElseIf LstVolumeFrom.SelectedIndex = 0 And LstVolumeTo.SelectedIndex = 2 Then
            TxtVolumeTo.Text = TxtVolumeFrom.Text * (1 / 1000)

        ElseIf LstVolumeFrom.SelectedIndex = 0 And LstVolumeTo.SelectedIndex = 3 Then
            TxtVolumeTo.Text = TxtVolumeFrom.Text * (1 / 1000000)

        ElseIf LstVolumeFrom.SelectedIndex = 0 And LstVolumeTo.SelectedIndex = 4 Then
            TxtVolumeTo.Text = TxtVolumeFrom.Text * (1 / 4.92892159) 'Teaspoon / ml'

        ElseIf LstVolumeFrom.SelectedIndex = 0 And LstVolumeTo.SelectedIndex = 5 Then
            TxtVolumeTo.Text = TxtVolumeFrom.Text * (1 / 14.78676478) 'tbsp / ml'

        ElseIf LstVolumeFrom.SelectedIndex = 0 And LstVolumeTo.SelectedIndex = 6 Then
            TxtVolumeTo.Text = TxtVolumeFrom.Text * (1 / 29.57352956) 'fl. oz. / ml'

        ElseIf LstVolumeFrom.SelectedIndex = 0 And LstVolumeTo.SelectedIndex = 7 Then
            TxtVolumeTo.Text = TxtVolumeFrom.Text * (1 / 236.5882365) 'cups / ml'

        ElseIf LstVolumeFrom.SelectedIndex = 0 And LstVolumeTo.SelectedIndex = 8 Then
            TxtVolumeTo.Text = TxtVolumeFrom.Text * (1 / 473.176473) 'pint / ml')

        ElseIf LstVolumeFrom.SelectedIndex = 0 And LstVolumeTo.SelectedIndex = 9 Then
            TxtVolumeTo.Text = TxtVolumeFrom.Text * (1 / 946.352946) 'quarts / ml'

        ElseIf LstVolumeFrom.SelectedIndex = 0 And LstVolumeTo.SelectedIndex = 10 Then
            TxtVolumeTo.Text = TxtVolumeFrom.Text * (1 / 3785.411784) 'gal/ ml'

        ElseIf LstVolumeFrom.SelectedIndex = 0 And LstVolumeTo.SelectedIndex = 11 Then
            TxtVolumeTo.Text = TxtVolumeFrom.Text * (1 / 16.387064) 'in^3 / ml'

        ElseIf LstVolumeFrom.SelectedIndex = 0 And LstVolumeTo.SelectedIndex = 12 Then
            TxtVolumeTo.Text = TxtVolumeFrom.Text * (1 / 28316.846592) 'ft^3 / ml' 

        ElseIf LstVolumeFrom.SelectedIndex = 0 And LstVolumeTo.SelectedIndex = 13 Then
            TxtVolumeTo.Text = TxtVolumeFrom.Text * (1 / 764554.85798) 'yd^3 / ml

            'Cubic centimeters'
        ElseIf LstVolumeFrom.SelectedIndex = 1 And LstVolumeTo.SelectedIndex = 0 Then
            TxtVolumeTo.Text = TxtVolumeFrom.Text * 1

        ElseIf LstVolumeFrom.SelectedIndex = 1 And LstVolumeTo.SelectedIndex = 1 Then
            TxtVolumeTo.Text = TxtVolumeFrom.Text

        ElseIf LstVolumeFrom.SelectedIndex = 1 And LstVolumeTo.SelectedIndex = 2 Then
            TxtVolumeTo.Text = TxtVolumeFrom.Text * (1 / 1000) 'L / cm^3'

        ElseIf LstVolumeFrom.SelectedIndex = 1 And LstVolumeTo.SelectedIndex = 3 Then
            TxtVolumeTo.Text = TxtVolumeFrom.Text * (1 / 1000000) 'm^3 / cm^3'

            'Liters'
        ElseIf LstVolumeFrom.SelectedIndex = 2 And LstVolumeTo.SelectedIndex = 0 Then
            TxtVolumeTo.Text = TxtVolumeFrom.Text * (1000 / 1) 'ml / L'

        ElseIf LstVolumeFrom.SelectedIndex = 2 And LstVolumeTo.SelectedIndex = 1 Then
            TxtVolumeTo.Text = TxtVolumeFrom.Text * (1000 / 1) 'cm^3 / L'

        ElseIf LstVolumeFrom.SelectedIndex = 2 And LstVolumeTo.SelectedIndex = 2 Then
            TxtVolumeTo.Text = TxtVolumeFrom.Text

        ElseIf LstVolumeFrom.SelectedIndex = 2 And LstVolumeTo.SelectedIndex = 3 Then
            TxtVolumeTo.Text = TxtVolumeFrom.Text * (1 / 1000) 'm^3 / L'
        End If
    End Sub

    'Length Converter'
    Private Sub BtnLengthConvert_Click(sender As Object, e As EventArgs) Handles BtnLengthConvert.Click
        If LstLengthFrom.SelectedIndex = 0 And LstLengthTo.SelectedIndex = 0 Then
            TxtLengthTo.Text = TxtLengthFrom.Text

        ElseIf LstLengthFrom.SelectedIndex = 0 And LstLengthTo.SelectedIndex = 1 Then
            TxtLengthTo.Text = TxtLengthFrom.Text * (1 / 1000) 'km / m'

        ElseIf LstLengthFrom.SelectedIndex = 0 And LstLengthTo.SelectedIndex = 2 Then
            TxtLengthTo.Text = TxtLengthFrom.Text * (100 / 1) 'cm / m'

        ElseIf LstLengthFrom.SelectedIndex = 0 And LstLengthTo.SelectedIndex = 3 Then
            TxtLengthTo.Text = TxtLengthFrom.Text * (1000 / 1) 'mm / m'

        ElseIf LstLengthFrom.SelectedIndex = 1 And LstLengthTo.SelectedIndex = 0 Then
            TxtLengthTo.Text = TxtLengthFrom.Text * (1000 / 1) 'm / km'

        ElseIf LstLengthFrom.SelectedIndex = 1 And LstLengthTo.SelectedIndex = 1 Then
            TxtLengthTo.Text = TxtLengthFrom.Text

        ElseIf LstLengthFrom.SelectedIndex = 1 And LstLengthTo.SelectedIndex = 2 Then
            TxtLengthTo.Text = TxtLengthFrom.Text * (100000 / 1) 'cm / km'

        ElseIf LstLengthFrom.SelectedIndex = 1 And LstLengthTo.SelectedIndex = 3 Then
            TxtLengthTo.Text = TxtLengthFrom.Text * (1000000 / 1) 'mm / km'

        ElseIf LstLengthFrom.SelectedIndex = 2 And LstLengthTo.SelectedIndex = 0 Then
            TxtLengthTo.Text = TxtLengthFrom.Text * (1 / 100) 'm / cm'

        ElseIf LstLengthFrom.SelectedIndex = 2 And LstLengthTo.SelectedIndex = 1 Then
            TxtLengthTo.Text = TxtLengthFrom.Text * (1 / 100000) 'km / cm'

        ElseIf LstLengthFrom.SelectedIndex = 2 And LstLengthTo.SelectedIndex = 2 Then
            TxtLengthTo.Text = TxtLengthFrom.Text

        ElseIf LstLengthFrom.SelectedIndex = 2 And LstLengthTo.SelectedIndex = 3 Then
            TxtLengthTo.Text = TxtLengthFrom.Text * (10 / 1) 'mm / cm'
        End If
    End Sub

    'Weight & Mass Converter'
    Private Sub BtnMassConvert_Click(sender As Object, e As EventArgs) Handles BtnMassConvert.Click
        If LstMassFrom.SelectedIndex = 0 And LstMassTo.SelectedIndex = 0 Then
            TxtMassTo.Text = TxtMassFrom.Text

        ElseIf LstMassFrom.SelectedIndex = 0 And LstMassTo.SelectedIndex = 1 Then
            TxtMassTo.Text = TxtMassFrom.Text * (1000 / 1) 'g / kg'

        ElseIf LstMassFrom.SelectedIndex = 0 And LstMassTo.SelectedIndex = 2 Then
            TxtMassTo.Text = TxtMassFrom.Text * (1000000 / 1) 'mg / kg'

        ElseIf LstMassFrom.SelectedIndex = 0 And LstMassTo.SelectedIndex = 3 Then
            TxtMassTo.Text = TxtMassFrom.Text * (1 / 907.18474) 'T / kg'

        ElseIf LstMassFrom.SelectedIndex = 1 And LstMassTo.SelectedIndex = 0 Then
            TxtMassTo.Text = TxtMassFrom.Text * (1 / 1000) 'kg / g'

        ElseIf LstMassFrom.SelectedIndex = 1 And LstMassTo.SelectedIndex = 1 Then
            TxtMassTo.Text = TxtMassFrom.Text

        ElseIf LstMassFrom.SelectedIndex = 1 And LstMassTo.SelectedIndex = 2 Then
            TxtMassTo.Text = TxtMassFrom.Text * (1000 / 1) 'mg / g'

        ElseIf LstMassFrom.SelectedIndex = 1 And LstMassTo.SelectedIndex = 3 Then
            TxtMassTo.Text = TxtMassFrom.Text * (1 / 907184.74) 'T / g'

        ElseIf LstMassFrom.SelectedIndex = 2 And LstMassTo.SelectedIndex = 0 Then
            TxtMassTo.Text = TxtMassFrom.Text * (1 / 1000000) 'kg / mg'

        ElseIf LstMassFrom.SelectedIndex = 2 And LstMassTo.SelectedIndex = 1 Then
            TxtMassTo.Text = TxtMassFrom.Text * (1 / 1000) 'g / mg'

        ElseIf LstMassFrom.SelectedIndex = 2 And LstMassTo.SelectedIndex = 2 Then
            TxtMassTo.Text = TxtMassFrom.Text

        ElseIf LstMassFrom.SelectedIndex = 2 And LstMassTo.SelectedIndex = 3 Then
            TxtMassTo.Text = TxtMassFrom.Text * (1 / 907184740) 'T / mg'
        End If
    End Sub

    'Energy Converter'
    Private Sub BtnEnergyConvert_Click(sender As Object, e As EventArgs) Handles BtnEnergyConvert.Click
        If LstEnergyFrom.SelectedIndex = 0 And LstEnergyTo.SelectedIndex = 0 Then
            TxtEnergyTo.Text = TxtEnergyFrom.Text

        ElseIf LstEnergyFrom.SelectedIndex = 0 And LstEnergyTo.SelectedIndex = 1 Then
            TxtEnergyTo.Text = TxtEnergyFrom.Text * 251.9957963122

        ElseIf LstEnergyFrom.SelectedIndex = 0 And LstEnergyTo.SelectedIndex = 2 Then
            TxtEnergyTo.Text = TxtEnergyFrom.Text * 778.1693709679

        ElseIf LstEnergyFrom.SelectedIndex = 0 And LstEnergyTo.SelectedIndex = 3 Then
            TxtEnergyTo.Text = TxtEnergyFrom.Text * 0.000001055056

        ElseIf LstEnergyFrom.SelectedIndex = 1 And LstEnergyTo.SelectedIndex = 0 Then
            TxtEnergyTo.Text = TxtEnergyFrom.Text * 0.003968320164996

        ElseIf LstEnergyFrom.SelectedIndex = 1 And LstEnergyTo.SelectedIndex = 1 Then
            TxtEnergyTo.Text = TxtEnergyFrom.Text

        ElseIf LstEnergyFrom.SelectedIndex = 1 And LstEnergyTo.SelectedIndex = 2 Then
            TxtEnergyTo.Text = TxtEnergyFrom.Text * 3.088025206594

        ElseIf LstEnergyFrom.SelectedIndex = 1 And LstEnergyTo.SelectedIndex = 3 Then
            TxtEnergyTo.Text = TxtEnergyFrom.Text * 0.0000000041868

        ElseIf LstEnergyFrom.SelectedIndex = 2 And LstEnergyTo.SelectedIndex = 0 Then
            TxtEnergyTo.Text = TxtEnergyFrom.Text * 0.001285067283946

        ElseIf LstEnergyFrom.SelectedIndex = 2 And LstEnergyTo.SelectedIndex = 1 Then
            TxtEnergyTo.Text = TxtEnergyFrom.Text * 0.3238315535329

        ElseIf LstEnergyFrom.SelectedIndex = 2 And LstEnergyTo.SelectedIndex = 2 Then
            TxtEnergyTo.Text = TxtEnergyFrom.Text

        ElseIf LstEnergyFrom.SelectedIndex = 2 And LstEnergyTo.SelectedIndex = 3 Then
            TxtEnergyTo.Text = TxtEnergyFrom.Text * 0.000000001355817948331
        End If
    End Sub

    'Temperature Converter'
    Private Sub BtnTempConvert_Click(sender As Object, e As EventArgs) Handles BtnTempConvert.Click
        If LstTempFrom.SelectedIndex = 0 And LstTempTo.SelectedIndex = 0 Then
            TxtTempTo.Text = TxtTempFrom.Text

        ElseIf LstTempFrom.SelectedIndex = 0 And LstTempTo.SelectedIndex = 1 Then
            TxtTempTo.Text = TxtTempFrom.Text + 273.15

        ElseIf LstTempFrom.SelectedIndex = 0 And LstTempTo.SelectedIndex = 2 Then
            TxtTempTo.Text = (TxtTempFrom.Text * (9 / 5) + 32)

        ElseIf LstTempFrom.SelectedIndex = 1 And LstTempTo.SelectedIndex = 0 Then
            TxtTempTo.Text = TxtTempFrom.Text - 273.15

        ElseIf LstTempFrom.SelectedIndex = 1 And LstTempTo.SelectedIndex = 1 Then
            TxtTempTo.Text = TxtTempFrom.Text

        ElseIf LstTempFrom.SelectedIndex = 1 And LstTempTo.SelectedIndex = 2 Then
            TxtTempTo.Text = (TxtTempFrom.Text - 273.15) * (9 / 5) + 32

        ElseIf LstTempFrom.SelectedIndex = 2 And LstTempTo.SelectedIndex = 0 Then
            TxtTempTo.Text = (TxtTempFrom.Text - 32) * (5 / 9)

        ElseIf LstTempFrom.SelectedIndex = 2 And LstTempTo.SelectedIndex = 1 Then
            TxtTempTo.Text = (TxtTempFrom.Text + 459.67) * (5 / 9)

        ElseIf LstTempFrom.SelectedIndex = 2 And LstTempTo.SelectedIndex = 2 Then
            TxtTempTo.Text = TxtTempFrom.Text
        End If
    End Sub
End Class