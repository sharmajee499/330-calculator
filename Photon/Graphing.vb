Public Class Graphing
    Dim xmin As Double, xmax As Double, ymin As Double, ymax As Double, xtick As Double, ytick As Double
    Dim ans = 0
    Dim angleX As Double, angleY As Double, zFactor As Double = 1, viewdistance As Double = 20, fov As Double = 800, last3DGraph As String = "", facesAreColored As Boolean = False
    Dim changesOccured As Boolean, fileCreated As String = ""
    Dim eps As Double = 0.0000000000001
    Dim deltaX As Double, deltaY As Double
    Dim lowA As Double, highA As Double, midA As Double
    Dim traceNumber As Double = 0, tracePoint As PointF = New PointF(0, 0), traceCoord As PointF = New PointF(0, 0), shouldTrace As Boolean = False, shouldDrawTrace As Boolean = False, traceCoord3D As New Point3D(0, 0, 0), traceCoord3DOriginal As New Point3D(0, 0, 0)
    Dim allThePoints As New ArrayList, all3DFaces As New List(Of Face3D), firstDerivativePoints As New ArrayList, secondDerivativePoints As New ArrayList, dataPoints As New ArrayList, integrals As New ArrayList, integralFlag As New ArrayList, all3DPoints As New ArrayList
    Dim firstDerivative As Boolean = False, secondDerivative As Boolean = False, shouldIntegrate As Boolean = False, all3DFunctionNames As New List(Of String)
    Dim findType As Integer = -1, startingGuess As Double = 0, foundPoint As PointF = New PointF(0, 0)
    Dim shouldReset As Boolean = False
    Dim functionPen As Pen, axisPen As Pen, gridPen As Pen, firstDPen As Pen, secondDPen As Pen, slopePen As Pen
    Dim shouldScroll As Boolean, isSlopeFieldSolve As Boolean = False
    Dim rndSeed As Double
    Dim sFieldLocationArray(0, 0) As PointF, sFieldSlopeArray(0, 0) As Double, slopeFieldPoint As PointF
    Dim currentX As Double, currentY As Double, isFindingSlopeField As Boolean
    Dim WithEvents reco As New System.Speech.Recognition.SpeechRecognitionEngine(New System.Globalization.CultureInfo("en-US"))
    Dim openParenth As Integer = 0, xtickMultiplier As Integer = 2, ytickMultiplier = 2
    Dim variableCreate As Boolean = False
    Dim lastDrawPoint As PointF, allDrawPoints As New ArrayList
    Dim transformX As String = "u", transformY As String = "v", transformZ As String = "w"
    'matrices
    Dim A(,) As Double
    Dim B(,) As Double
    Dim C(,) As Double
    'units
    Dim prefixDictionary As New Dictionary(Of String, Double)
    Dim lengthDictionary As New Dictionary(Of String, Double)
    Dim areaDictionary As New Dictionary(Of String, Double)
    Dim volumeDictionary As New Dictionary(Of String, Double)
    Dim pressureDictionary As New Dictionary(Of String, Double)
    Dim timeDictionary As New Dictionary(Of String, Double)
    Dim energyDictionary As New Dictionary(Of String, Double)
    Dim massDictionary As New Dictionary(Of String, Double)
    Dim fuelDictionary As New Dictionary(Of String, Double)
    Dim angleDictionary As New Dictionary(Of String, Double)
    Dim speedDictionary As New Dictionary(Of String, Double)
    Dim densityDictionary As New Dictionary(Of String, Double)
    Dim accelerationDictionary As New Dictionary(Of String, Double)
    Dim powerDictionary As New Dictionary(Of String, Double)
    Dim elementDictionary As New Dictionary(Of String, Double)

    ' click the mouse down - tells photon that the viewing range is about to be dragged, or that it's about to be drawn on
    Private Sub PictureBox1_MouseDown(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseDown
        If e.Button = MouseButtons.Left Then
            PictureBox1.Tag = "MOUSE DOWN"
            deltaX = e.X
            deltaY = e.Y
            If RichTextBox1.Text.Contains("y' = ") Then
                isSlopeFieldSolve = True
            End If
        ElseIf e.Button = MouseButtons.Right Then
            allDrawPoints.Clear()
            PictureBox1.Tag = "RIGHT MOUSE DOWN"
            lastDrawPoint = e.Location
            allDrawPoints.Add(New PointF(e.X / PictureBox1.Width * (xmax - xmin) + xmin, (PictureBox1.Height - e.Y) / PictureBox1.Height * (ymax - ymin) + ymin))
        End If
        PictureBox1.ContextMenuStrip = ContextMenuStrip1
    End Sub
    ' drag the graph to adjust viewing range
    Private Sub PictureBox1_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseMove
        If PictureBox1.Tag = "MOUSE DOWN" Then
            isSlopeFieldSolve = False
            Dim dX = (deltaX - e.X) / PictureBox1.Width * (xmax - xmin)
            Dim dY = (e.Y - deltaY) / PictureBox1.Height * (ymax - ymin)
            If all3DFaces.Count = 0 And all3DPoints.Count <= 1 Then
                xmin = xmin + dX
                xmax = xmax + dX
                ymin = ymin + dY
                ymax = ymax + dY
            Else
                angleX = (e.Y - deltaY)
                angleY = (deltaX - e.X)
            End If
            deltaX = e.X
            deltaY = e.Y
            PictureBox1.Refresh()
        ElseIf PictureBox1.Tag = "RIGHT MOUSE DOWN" Then
            PictureBox1.ContextMenuStrip = Nothing
            allDrawPoints.Add(New PointF(e.X / PictureBox1.Width * (xmax - xmin) + xmin, (PictureBox1.Height - e.Y) / PictureBox1.Height * (ymax - ymin) + ymin))
            Dim g As System.Drawing.Graphics = PictureBox1.CreateGraphics
            g.DrawLine(New Pen(Brushes.Blue, 1), lastDrawPoint, e.Location)
            lastDrawPoint = e.Location
            PictureBox1.Cursor = Cursors.Cross
        ElseIf all3DFaces.Count <> 0 Then
            If e.X < 40 Or e.X > PictureBox1.Width - 40 Then
                PictureBox1.Cursor = Cursors.NoMoveVert
            Else
                PictureBox1.Cursor = Cursors.Cross
            End If
        ElseIf all3DPoints.Count <= 1 Then 'for x-zoom or y-zoom, change cursor
            If e.X < 40 Or e.X > PictureBox1.Width - 40 Then
                PictureBox1.Cursor = Cursors.NoMoveVert
            ElseIf e.Y < 40 Or e.Y > PictureBox1.Height - 40 Then
                PictureBox1.Cursor = Cursors.NoMoveHoriz
            Else
                PictureBox1.Cursor = Cursors.Cross
            End If
        Else
            PictureBox1.Cursor = Cursors.Cross
        End If
    End Sub
    ' release the mouse.  If the user was dragging the right mouse button, find the best fit equation
    Private Sub PictureBox1_MouseUp(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseUp
        On Error GoTo theEnd
        If PictureBox1.Tag = "RIGHT MOUSE DOWN" And allDrawPoints.Count > 3 Then
            Dim functionString As String = ""
            If SimpleLinearToolStripMenuItem.Checked Then
                Dim x1 = Math.Round(allDrawPoints(0).x / xtick) * xtick
                Dim y1 = Math.Round(allDrawPoints(0).y / ytick) * ytick
                Dim x2 = Math.Round(allDrawPoints(allDrawPoints.Count - 1).x / xtick) * xtick
                Dim y2 = Math.Round(allDrawPoints(allDrawPoints.Count - 1).y / ytick) * ytick
                Dim rise = y2 - y1
                Dim run = x2 - x1
                Dim b = y2 - (rise / run) * x2
                Dim slope = (rise / run).ToString
                Dim y_int = b.ToString
                If My.Settings.fraction Then
                    slope = decToFraction(rise / run)
                    y_int = decToFraction(b)
                End If
                functionString = slope + "*x + " + y_int
                functionString = Replace(functionString, "+ -", "- ")
            ElseIf LinearToolStripMenuItem.Checked Then
                functionString = LinearRegression(allDrawPoints)
            ElseIf SimpleQuadraticToolStripMenuItem.Checked Then
                Dim p1 As PointF = New PointF(Math.Round(allDrawPoints(0).x / xtick) * xtick, Math.Round(allDrawPoints(0).y / ytick) * ytick)
                Dim p2 As PointF = New PointF(Math.Round(allDrawPoints(Math.Round(allDrawPoints.Count / 2) - 1).x / xtick) * xtick, Math.Round(allDrawPoints(Math.Round(allDrawPoints.Count / 2) - 1).y / ytick) * ytick)
                Dim p3 As PointF = New PointF(Math.Round(allDrawPoints(allDrawPoints.Count - 1).x / xtick) * xtick, Math.Round(allDrawPoints(allDrawPoints.Count - 1).y / ytick) * ytick)
                allDrawPoints.Clear()
                allDrawPoints.Add(p1)
                allDrawPoints.Add(p2)
                allDrawPoints.Add(p3)
                functionString = QuadraticRegression(allDrawPoints, True)
            ElseIf QuadraticToolStripMenuItem.Checked Then
                functionString = QuadraticRegression(allDrawPoints)
            ElseIf SimpleCubicToolStripMenuItem.Checked Then
                Dim p1 As PointF = New PointF(Math.Round(allDrawPoints(0).x / xtick) * xtick, Math.Round(allDrawPoints(0).y / ytick) * ytick)
                Dim p2 As PointF = New PointF(Math.Round(allDrawPoints(Math.Round(allDrawPoints.Count / 4) - 1).x / xtick) * xtick, Math.Round(allDrawPoints(Math.Round(allDrawPoints.Count / 4) - 1).y / ytick) * ytick)
                Dim p3 As PointF = New PointF(Math.Round(allDrawPoints(Math.Round(allDrawPoints.Count * 3 / 4) - 1).x / xtick) * xtick, Math.Round(allDrawPoints(Math.Round(allDrawPoints.Count * 3 / 4) - 1).y / ytick) * ytick)
                Dim p4 As PointF = New PointF(Math.Round(allDrawPoints(allDrawPoints.Count - 1).x / xtick) * xtick, Math.Round(allDrawPoints(allDrawPoints.Count - 1).y / ytick) * ytick)
                allDrawPoints.Clear()
                allDrawPoints.Add(p1)
                allDrawPoints.Add(p2)
                allDrawPoints.Add(p3)
                allDrawPoints.Add(p4)
                functionString = CubicRegression(allDrawPoints, True)
            Else
                functionString = CubicRegression(allDrawPoints)
            End If
            RichTextBox2.Text = functionString
            Button1.PerformClick()
        ElseIf e.Button = MouseButtons.Left Then 'solve slope field
            If isSlopeFieldSolve = True Then
                Dim xCoord As Double = e.X / PictureBox1.Width * (xmax - xmin) + xmin
                Dim yCoord As Double = (PictureBox1.Height - e.Y) / PictureBox1.Height * (ymax - ymin) + ymin
                slopeFieldPoint = New PointF(xCoord, yCoord)
            End If
        End If
theEnd:
        PictureBox1.Tag = ""
        PictureBox1.Refresh()
    End Sub

    Private Function RoundToSignificance(number As Double, significance As Double) As Integer
        RoundToSignificance = Math.Round(number / xtick) * xtick
    End Function

    Private Function det44(a, b, c, d, e, f, g, h, i, j, k, l, m, n, o, p)
        det44 = a * det33(f, g, h, j, k, l, n, o, p) - b * det33(e, g, h, i, k, l, m, o, p) + c * det33(e, f, h, i, j, l, m, n, p) - d * det33(e, f, g, i, j, k, m, n, o)
    End Function

    Private Function det33(a, b, c, d, e, f, g, h, i)
        det33 = a * det22(e, f, h, i) - b * det22(d, f, g, i) + c * det22(d, e, g, h)
    End Function

    Private Function det22(a, b, c, d)
        det22 = a * d - b * c
    End Function

    Private Sub calculateThingsBeforeDrawing()
        On Error GoTo ending
        'set up table if needed
        If Me.Tag = "table" Then
            While DataFunction.Columns.Count > 1
                DataFunction.Columns.RemoveAt(1)
            End While
            While DataFunction.Rows.Count > 0
                DataFunction.Rows.RemoveAt(0)
            End While
            For i = 0 To (DataFunction.Height - 23) / 23
                DataFunction.Rows.Add(Math.Round(Val(txtStart.Text) + Val(txtDelta.Text) * i, 12).ToString)
            Next i
        End If
        If My.Settings.autofix Then 'automatically scale the graph to get a square aspect ratio
            Dim tempXMin = xmin
            Dim tempXMax = xmax
            Dim delta = ((ymax - ymin) * PictureBox1.Width) / (PictureBox1.Height) - (xmax - xmin)
            xmin -= delta / 2
            xmax += delta / 2
            If xmin = xmax Then
                xmin = tempXMin
                xmax = tempXMax
                delta = ((xmax - xmin) * PictureBox1.Height) / (PictureBox1.Width) - (ymax - ymin)
                ymin -= delta / 2
                ymax += delta / 2
            End If
            If xmin < -1.0E+307 Then xmin = -1.0E+307
            If xmax > 1.0E+307 Then xmax = 1.0E+307
            If ymin < -1.0E+307 Then ymin = -1.0E+307
            If ymax > 1.0E+307 Then ymax = 1.0E+307
        End If
        If PictureBox1.Tag = "" Then 'skip this if we're just dragging
            If RichTextBox1.Text.StartsWith("trace") Then
                shouldTrace = True
                Dim s = RichTextBox1.Lines(0)
                traceNumber = eval(heal(Split(s, "trace")(1)), False, False)
            Else
                shouldDrawTrace = False
                shouldTrace = False
            End If
            allThePoints.Clear()
            firstDerivativePoints.Clear()
            secondDerivativePoints.Clear()
            dataPoints.Clear()
            integrals.Clear()
            integralFlag.Clear()
            ReDim sFieldLocationArray(0, 0)
            ReDim sFieldSlopeArray(0, 0)
            Dim aLines = RichTextBox1.Lines
            For i = 0 To RichTextBox1.Lines.Length - 1 'goes through the lines forwards for calculations - last number is ans
                Dim text = heal(aLines(i))
                aLines(i) = ""
                If text.Contains("->") Or text.StartsWith("'") Or text.Contains(" = ") Then
                    aLines(i) = text
                Else
                    aLines(i) = calculate(text)
                End If
            Next i
            RichTextBox1.Lines = aLines
            aLines = RichTextBox1.Lines
            For i = RichTextBox1.Lines.Length - 1 To 0 Step -1 'goes through the lines backwards for graphs - first graph accessed is trace
                Dim text = aLines(i)
                aLines(i) = ""
                If text.Contains(" = ") And text.StartsWith("'") = False Then
                    aLines(i) = calculate(text)
                Else
                    aLines(i) = text
                End If
            Next i
            RichTextBox1.Lines = aLines
            RichTextBox1.Text = Replace(RichTextBox1.Text, vbNewLine + vbNewLine, vbNewLine)
            If RichTextBox1.Text.EndsWith(vbNewLine) Then RichTextBox1.Text = RichTextBox1.Text.Substring(0, RichTextBox1.Text.Length - 1)
            'Application.DoEvents()
            For i = 0 To DataXY.Rows.Count - 1
                If IsNumeric(DataXY.Rows(i).Cells(0).Value) And IsNumeric(DataXY.Rows(i).Cells(1).Value) Then
                    Dim x = DataXY.Rows(i).Cells(0).Value
                    Dim y = DataXY.Rows(i).Cells(1).Value
                    Dim shouldDraw As Boolean = True
                    If LogXAxisToolStripMenuItem.Checked = True Then
                        If x <= 0 Then
                            shouldDraw = False
                        Else
                            x = Math.Log10(x)
                        End If
                    End If
                    If LogYAxisToolStripMenuItem.Checked = True Then
                        If y <= 0 Then
                            shouldDraw = False
                        Else
                            y = Math.Log10(y)
                        End If
                    End If
                    If shouldDraw = True Then dataPoints.Add(New PointF((x - xmin) / (xmax - xmin) * PictureBox1.Width, (ymax - y) / (ymax - ymin) * PictureBox1.Height))
                End If
            Next i
            If My.Settings.colorText = True Then
                ColorTheText()
            End If
        End If
ending:
    End Sub

    Private Function adjustXTickSpacing(ByVal inc As Double)
        While inc < PictureBox1.Width / 23 'adjust x ticks if there are too many or too few of them
            xtick = xtick * xtickMultiplier
            inc = xtick / (xmax - xmin) * PictureBox1.Width
            xtickMultiplier = 10 / xtickMultiplier
        End While
        While inc > PictureBox1.Width / 3 And PictureBox1.Width > 130
            xtick = xtick / (10 / xtickMultiplier)
            inc = xtick / (xmax - xmin) * PictureBox1.Width
            xtickMultiplier = 10 / xtickMultiplier
        End While
        adjustXTickSpacing = inc
    End Function

    Private Function adjustYTickSpacing(ByVal inc As Double)
        While inc < PictureBox1.Height / 23 'adjust y ticks if there are too many or too few of them
            ytick = ytick * ytickMultiplier
            inc = ytick / (ymax - ymin) * PictureBox1.Height
            ytickMultiplier = 10 / ytickMultiplier
        End While
        While inc > PictureBox1.Height / 3
            ytick = ytick / (10 / ytickMultiplier)
            inc = ytick / (ymax - ymin) * PictureBox1.Height
            ytickMultiplier = 10 / ytickMultiplier
        End While
        adjustYTickSpacing = inc
    End Function

    ' All graphics methods go here '''''''''''''''''''''''''''''
    Private Sub PictureBox1_Paint(sender As Object, e As PaintEventArgs) Handles PictureBox1.Paint
        On Error Resume Next
        tracePoint = New PointF(-1, -1)
        calculateThingsBeforeDrawing()
        'drawStuff:  'if the math stuff above crashes, we should at least draw axes and whatnot
        If all3DFaces.Count = 0 And all3DPoints.Count <= 1 Then                          '2D graph
            functionPen = New Pen(Brushes.Blue, My.Settings.thickness)
            firstDPen = New Pen(Brushes.Purple, My.Settings.thickness)
            secondDPen = New Pen(Brushes.Red, My.Settings.thickness)
            slopePen = New Pen(Brushes.DarkGreen, My.Settings.thickness)
            axisPen = New Pen(Brushes.Black, My.Settings.thickness)
            gridPen = New Pen(Brushes.LightGray, My.Settings.thickness)
            Dim xAxis = ymax / (ymax - ymin) * PictureBox1.Height           ' y value of the x axis
            Dim yAxis = -xmin / (xmax - xmin) * PictureBox1.Width           ' x value of the y axis
            ''''''''''''''''''''' draw ticks '''''''''''''''''''''''
            Dim increment = xtick / (xmax - xmin) * PictureBox1.Width
            increment = adjustXTickSpacing(increment)
            Dim t = (Math.Round(xmin / xtick) * xtick - xmin) / (xmax - xmin) * PictureBox1.Width
            Dim tNumber = Math.Round(xmin / xtick) * xtick
            Dim tickWidth = Math.Max(Math.Min(Math.Pow(My.Settings.thickness, 1 / 4) * (PictureBox1.Height + PictureBox1.Width) / 150 / Math.Pow(((xmax - xmin) / xtick / 6), 1 / 4), 30), 3)
            Dim fontSize = Math.Sqrt(tickWidth) * 3.8
            Dim nChars As Integer = Math.Max(Math.Min(Math.Round((PictureBox1.Width / (fontSize * 1.2)) / ((xmax - xmin) / xtick) - 1), 15), 2)
            If xmax - xmin < 0.000001 Then nChars = 15
            While t <= PictureBox1.Width 'x ticks, x tick labels, x grid, and y axis
                If Math.Abs(t - yAxis) > eps * 10 Then
                    e.Graphics.DrawLine(gridPen, New PointF(t, 0), New PointF(t, PictureBox1.Height)) 'draw grid line
                    e.Graphics.DrawLine(axisPen, New PointF(t, xAxis + tickWidth), New PointF(t, xAxis - tickWidth)) 'draw tick mark
                End If
                If Math.Abs(tNumber) > eps Then
                    Dim tempString = Math.Round(tNumber, nChars).ToString
                    If xAxis > PictureBox1.Height - fontSize * 3 Then 'need to stick numbers to the bottom edge
                        e.Graphics.DrawString(tempString, New Font("Verdana", fontSize), Brushes.Black, New PointF(t - fontSize * tempString.Length / 2.1, PictureBox1.Height - fontSize * 2))
                        If LogXAxisToolStripMenuItem.Checked Then e.Graphics.DrawString("10", New Font("Verdana", fontSize), Brushes.Black, New PointF(t - fontSize * tempString.Length / 2.1 - fontSize * 1.5, PictureBox1.Height - fontSize * 0.7))
                    ElseIf xAxis < fontSize * -0.75 Then 'need to stick numbers to the top edge
                        e.Graphics.DrawString(tempString, New Font("Verdana", fontSize), Brushes.Black, New PointF(t - fontSize * tempString.Length / 2.1, fontSize * 0.25))
                        If LogXAxisToolStripMenuItem.Checked Then e.Graphics.DrawString("10", New Font("Verdana", fontSize), Brushes.Black, New PointF(t - fontSize * tempString.Length / 2.1 - fontSize * 1.5, fontSize * 1.55))
                    Else 'no sticking required
                        e.Graphics.DrawString(tempString, New Font("Verdana", fontSize), Brushes.Black, New PointF(t - fontSize * tempString.Length / 2.1, xAxis + fontSize * 1.2))
                        If LogXAxisToolStripMenuItem.Checked Then e.Graphics.DrawString("10", New Font("Verdana", fontSize), Brushes.Black, New PointF(t - fontSize * tempString.Length / 2.1 - fontSize * 1.5, xAxis + fontSize * 2.5))
                    End If
                End If
                t += increment
                tNumber += xtick
            End While
            increment = ytick / (ymax - ymin) * PictureBox1.Height
            increment = adjustYTickSpacing(increment)
            t = (ymax - Math.Round(ymax / ytick) * ytick) / (ymax - ymin) * PictureBox1.Height
            tNumber = Math.Round(ymax / ytick) * ytick
            While t <= PictureBox1.Height 'y ticks, y tick lables, y grid, and x axis
                If Math.Abs(t - xAxis) > eps * 10 Then
                    e.Graphics.DrawLine(gridPen, New PointF(0, t), New PointF(PictureBox1.Width, t)) 'draw grid
                    e.Graphics.DrawLine(axisPen, New PointF(yAxis + tickWidth, t), New PointF(yAxis - tickWidth, t)) 'draw ticks
                End If
                If Math.Abs(tNumber) > eps Then
                    Dim tempString = Math.Round(tNumber, Math.Min(nChars + 1, 15)).ToString
                    If yAxis > PictureBox1.Width + 7 Then 'need to stick numbers to the right edge
                        e.Graphics.DrawString(tempString, New Font("Verdana", fontSize), Brushes.Black, New PointF(PictureBox1.Width - fontSize * 1.6 - tempString.Length * fontSize / 1.25 + 7, t - fontSize / 1.3))
                        If LogYAxisToolStripMenuItem.Checked Then e.Graphics.DrawString("10", New Font("Verdana", fontSize), Brushes.Black, New PointF(PictureBox1.Width - fontSize * 3.1 - tempString.Length * fontSize / 1.25 + 7, t - fontSize / 1.3 + fontSize * 1.3))
                    ElseIf yAxis < fontSize * 1.5 + tempString.Length * fontSize / 1.25 + 3 Then 'need to stick numbers to the left side
                        e.Graphics.DrawString(tempString, New Font("Verdana", fontSize), Brushes.Black, New PointF(2, t - fontSize / 1.3))
                        If LogYAxisToolStripMenuItem.Checked Then e.Graphics.DrawString("10", New Font("Verdana", fontSize), Brushes.Black, New PointF(2 - fontSize * 1.5, t - fontSize / 1.3 + fontSize * 1.3))
                    Else 'no sticking required
                        e.Graphics.DrawString(tempString, New Font("Verdana", fontSize), Brushes.Black, New PointF(yAxis - fontSize * 1.6 - tempString.Length * fontSize / 1.25, t - fontSize / 1.3))
                        If LogYAxisToolStripMenuItem.Checked Then e.Graphics.DrawString("10", New Font("Verdana", fontSize), Brushes.Black, New PointF(yAxis - fontSize * 3.1 - tempString.Length * fontSize / 1.25, t - fontSize / 1.3 + fontSize * 1.3))
                    End If
                End If
                t += increment
                tNumber -= ytick
            End While
            ''''''''''''''''''''' draw axes ''''''''''''''''''''''''
            e.Graphics.DrawLine(axisPen, New PointF(0, xAxis), New PointF(PictureBox1.Width, xAxis))
            e.Graphics.DrawLine(axisPen, New PointF(yAxis, 0), New PointF(yAxis, PictureBox1.Height))
            If PictureBox1.Tag <> "" Then Exit Sub
            If My.Settings.antialias Then e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            ''''''''''''''''''''' draw data ''''''''''''''''''''''''
            Dim size As Single = My.Settings.thickness + 7
            For Each p In dataPoints
                e.Graphics.FillEllipse(Brushes.Green, p.x - size / 2, p.y - size / 2, size, size)
                e.Graphics.DrawEllipse(Pens.LightGray, p.x - size / 2, p.y - size / 2, size, size)
            Next p
            '''''''''''''''''''' draw slope field ''''''''''''''''''
            If sFieldLocationArray.Length > 1 Then
                For i = 0 To sFieldLocationArray.GetLength(0) - 1
                    For j = 0 To sFieldLocationArray.GetLength(1) - 1
                        Dim dX = Math.Cos(Math.Atan(sFieldSlopeArray(i, j)))
                        Dim dY = Math.Sin(Math.Atan(sFieldSlopeArray(i, j)))
                        'Dim magnitude = Math.Sqrt(dX ^ 2 + dY ^ 2)
                        'dX = dX / magnitude
                        'dY = dY / magnitude
                        Dim tempPoint = New PointF((sFieldLocationArray(i, j).X - xmin) / (xmax - xmin) * PictureBox1.Width, (ymax - (sFieldLocationArray(i, j).Y)) / (ymax - ymin) * PictureBox1.Height)
                        Dim tempPoint1a = New PointF((sFieldLocationArray(i, j).X + dX - xmin) / (xmax - xmin) * PictureBox1.Width, (ymax - (sFieldLocationArray(i, j).Y + dY)) / (ymax - ymin) * PictureBox1.Height)
                        Dim magTempPoint1 = Math.Sqrt((tempPoint1a.X - tempPoint.X) ^ 2 + (tempPoint1a.Y - tempPoint.Y) ^ 2)
                        Dim vectorLength As Integer = 8
                        Dim tempPoint1 = New PointF(tempPoint.X + vectorLength * (tempPoint1a.X - tempPoint.X) / magTempPoint1, tempPoint.Y + vectorLength * (tempPoint1a.Y - tempPoint.Y) / magTempPoint1)
                        Dim tempPoint2 = New PointF(tempPoint.X - vectorLength * (tempPoint1a.X - tempPoint.X) / magTempPoint1, tempPoint.Y - vectorLength * (tempPoint1a.Y - tempPoint.Y) / magTempPoint1)
                        'tempPoint2 = tempPoint
                        If sFieldSlopeArray(i, j) <> 9999995 Then e.Graphics.DrawLine(slopePen, tempPoint1, tempPoint2)
                    Next j
                Next i
            End If
            ''''''''''''''''''''' draw functions '''''''''''''''''''
            Dim functionCounter = 0 'all regular functions
            For Each p In allThePoints
                If integralFlag(functionCounter) Then 'an arraylist stores which functions need to be drawn as integrals
                    e.Graphics.FillPolygon(Brushes.Blue, p) 'this fills in area for a visual display of integration
                Else
                    e.Graphics.DrawLines(functionPen, p) 'this just draws the function as a line
                End If
                functionCounter += 1
            Next p
            For Each p In firstDerivativePoints 'all derivatives
                e.Graphics.DrawLines(firstDPen, p)
            Next p
            For Each p In secondDerivativePoints 'all second derivatives
                e.Graphics.DrawLines(secondDPen, p)
            Next p
            ''''''''''''''''''''' draw trace '''''''''''''''''''''''
            If shouldDrawTrace And tracePoint <> New PointF(-1, -1) Then
                e.Graphics.FillEllipse(Brushes.Green, tracePoint.X - size / 2, tracePoint.Y - size / 2, size, size)
                e.Graphics.DrawEllipse(Pens.White, tracePoint.X - size / 2, tracePoint.Y - size / 2, size, size)
                Dim numberString As String = "t = " + traceNumber.ToString + vbNewLine + "x = " + traceCoord.X.ToString + vbNewLine + "y = " + traceCoord.Y.ToString
                Dim box = e.Graphics.MeasureString(numberString, New Font("verdana", fontSize))
                Dim boxX, boxY, px, py 'determine dimensions and location of box around trace info
                If tracePoint.X < PictureBox1.Width / 2 Then
                    boxX = 30
                    px = 30 + box.Width / 2
                Else
                    boxX = PictureBox1.Width - box.Width - 30
                    px = PictureBox1.Width - box.Width / 2 - 30
                End If
                If tracePoint.Y > PictureBox1.Height / 2 Then
                    boxY = 30
                    py = 35 + box.Height
                Else
                    boxY = PictureBox1.Height - box.Height - 60
                    py = PictureBox1.Height - box.Height - 65
                End If
                e.Graphics.FillRectangle(Brushes.White, boxX - 5, boxY - 5, box.Width + 10, box.Height + 10)
                e.Graphics.DrawString(numberString, New Font("verdana", fontSize), Brushes.Black, New Point(boxX, boxY))
                e.Graphics.DrawRectangle(Pens.Black, boxX - 5, boxY - 5, box.Width + 10, box.Height + 10)
                e.Graphics.DrawLine(Pens.Black, New PointF(px, py), tracePoint)
            End If
            ''''''''''''''''''''' draw numeric integrals '''''''''''
            Dim integralString As String = ""
            For i = integrals.Count - 1 To 0 Step -1
                integralString = integralString + "Area: " + eval(integrals(i), My.Settings.fraction, False).ToString + vbNewLine
            Next i
            If integralString <> "" Then
                integralString = integralString.Substring(0, integralString.Length - 1)
                Dim box = e.Graphics.MeasureString(integralString, New Font("verdana", fontSize))
                Dim boxX = 0, boxY = 20
                If tracePoint.X < PictureBox1.Width / 2 And tracePoint.Y > PictureBox1.Height / 2 And RichTextBox1.Text.StartsWith("trace") Then
                    boxX = PictureBox1.Width - box.Width - 20
                Else
                    boxX = 20
                End If
                e.Graphics.FillRectangle(Brushes.White, boxX - 5, boxY - 5, box.Width + 10, box.Height + 10)
                e.Graphics.DrawString(integralString, New Font("verdana", fontSize), Brushes.Black, New Point(boxX, boxY))
                e.Graphics.DrawRectangle(Pens.Black, boxX - 5, boxY - 5, box.Width + 10, box.Height + 10)
            End If
        ElseIf all3DPoints.Count <= 1 Then                        '3D graphics
beginningFaces:
            functionPen = New Pen(Brushes.Blue, My.Settings.thickness)
            axisPen = New Pen(Brushes.Black, My.Settings.thickness)
            Dim avgZ(all3DFaces.Count - 1) As Double
            Dim order(all3DFaces.Count - 1) As Integer
            Dim t As New List(Of Face3D)
            Dim tmp As Double, iMax As Integer
            For i = 0 To all3DFaces.Count - 1 'rotate with mouse, then project to the screen
                all3DFaces(i).p1 = all3DFaces(i).p1.RotateX(angleX).RotateY(angleY)
                all3DFaces(i).p2 = all3DFaces(i).p2.RotateX(angleX).RotateY(angleY)
                all3DFaces(i).p3 = all3DFaces(i).p3.RotateX(angleX).RotateY(angleY)
                all3DFaces(i).p4 = all3DFaces(i).p4.RotateX(angleX).RotateY(angleY)
                t.Add(New Face3D(all3DFaces(i).p1.Project(PictureBox1.Width, PictureBox1.Height, fov, viewdistance), all3DFaces(i).p2.Project(PictureBox1.Width, PictureBox1.Height, fov, viewdistance), all3DFaces(i).p3.Project(PictureBox1.Width, PictureBox1.Height, fov, viewdistance), all3DFaces(i).p4.Project(PictureBox1.Width, PictureBox1.Height, fov, viewdistance), all3DFaces(i).colors, all3DFaces(i).avgC))
                avgZ(i) = (t(i).p1.Z + t(i).p2.Z + t(i).p3.Z + t(i).p4.Z) / 4.0 'used for determining which order to draw faces
                order(i) = i 'used for ordering so we draw far faces first
            Next i
            For i = 0 To all3DFaces.Count - 2 'sort the faces so we can draw them furthest first and closest last,
                iMax = i
                For j = i + 1 To all3DFaces.Count - 1
                    If avgZ(j) > avgZ(iMax) Then
                        iMax = j
                    End If
                Next
                If iMax <> i Then
                    tmp = avgZ(i)
                    avgZ(i) = avgZ(iMax)
                    avgZ(iMax) = tmp
                    tmp = order(i)
                    order(i) = order(iMax)
                    order(iMax) = tmp
                End If
            Next i
            'the first time we graph something, assign colors to faces, rotate viewpoint
            If facesAreColored = False Then
                facesAreColored = True
                Dim zMax As Double = -(t(order(0)).p1.Z + t(order(0)).p2.Z + t(order(0)).p3.Z + t(order(0)).p4.Z) / 4
                Dim zMin As Double = -(t(order(avgZ.Count - 1)).p1.Z + t(order(avgZ.Count - 1)).p2.Z + t(order(avgZ.Count - 1)).p3.Z + t(order(avgZ.Count - 1)).p4.Z) / 4
                For i = 0 To all3DFaces.Count - 1
                    Dim centerScale As Double = ((all3DFaces(i).p1.Z + all3DFaces(i).p2.Z + all3DFaces(i).p3.Z + all3DFaces(i).p4.Z) / 4 - zMin) / (zMax - zMin)
                    Dim p1Scale As Double = (all3DFaces(i).p1.Z - zMin) / (zMax - zMin)
                    Dim p2Scale As Double = (all3DFaces(i).p2.Z - zMin) / (zMax - zMin)
                    Dim p3Scale As Double = (all3DFaces(i).p3.Z - zMin) / (zMax - zMin)
                    Dim p4Scale As Double = (all3DFaces(i).p4.Z - zMin) / (zMax - zMin)
                    all3DFaces(i).colors = {HSVtoColor(p1Scale * 260, 1, 1), HSVtoColor(p2Scale * 260, 1, 1), HSVtoColor(p3Scale * 260, 1, 1), HSVtoColor(p4Scale * 260, 1, 1)}
                    all3DFaces(i).avgC = HSVtoColor(centerScale * 260, 1, 1)
                    all3DFaces(i).p1 = all3DFaces(i).p1.RotateX(-90).RotateY(135).RotateX(20)
                    all3DFaces(i).p2 = all3DFaces(i).p2.RotateX(-90).RotateY(135).RotateX(20)
                    all3DFaces(i).p3 = all3DFaces(i).p3.RotateX(-90).RotateY(135).RotateX(20)
                    all3DFaces(i).p4 = all3DFaces(i).p4.RotateX(-90).RotateY(135).RotateX(20)
                Next i
                Dim theColorIsBlack As Color() = {Color.Black, Color.Black, Color.Black, Color.Black}
                all3DFaces.Add(New Face3D(New Point3D(0, -0.04, 0), New Point3D(0, 0.04, 0), New Point3D(5, 0.04, 0), New Point3D(5, -0.04, 0), theColorIsBlack, Color.Black))
                all3DFaces.Add(New Face3D(New Point3D(0, 0, -0.04), New Point3D(0, 0, 0.04), New Point3D(5, 0, 0.04), New Point3D(5, 0, -0.04), theColorIsBlack, Color.Black))
                all3DFaces.Add(New Face3D(New Point3D(-0.04, 0, 0), New Point3D(0.04, 0, 0), New Point3D(0.04, 5, 0), New Point3D(-0.04, 5, 0), theColorIsBlack, Color.Black))
                all3DFaces.Add(New Face3D(New Point3D(0, 0, -0.04), New Point3D(0, 0, 0.04), New Point3D(0, 5, 0.04), New Point3D(0, 5, -0.04), theColorIsBlack, Color.Black))
                all3DFaces.Add(New Face3D(New Point3D(0.03, -0.03, 0), New Point3D(-0.03, 0.03, 0), New Point3D(-0.03, 0.03, 5 * zFactor), New Point3D(0.03, -0.03, 5 * zFactor), theColorIsBlack, Color.Black))
                all3DFaces.Add(New Face3D(New Point3D(-0.03, -0.03, 0), New Point3D(0.03, 0.03, 0), New Point3D(0.03, 0.03, 5 * zFactor), New Point3D(-0.03, -0.03, 5 * zFactor), theColorIsBlack, Color.Black))
                For i = all3DFaces.Count - 6 To all3DFaces.Count - 1
                    all3DFaces(i).p1 = all3DFaces(i).p1.RotateX(-90).RotateY(135).RotateX(20)
                    all3DFaces(i).p2 = all3DFaces(i).p2.RotateX(-90).RotateY(135).RotateX(20)
                    all3DFaces(i).p3 = all3DFaces(i).p3.RotateX(-90).RotateY(135).RotateX(20)
                    all3DFaces(i).p4 = all3DFaces(i).p4.RotateX(-90).RotateY(135).RotateX(20)
                Next i
                GoTo beginningFaces
            End If
            'draw all the faces
            If My.Settings.antialias And CheckBox9Op.Checked Then e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            For i = 0 To all3DFaces.Count - 1
                Dim index As Integer = order(i)
                Dim factor = fov / (viewdistance + t(index).p1.Z)
                If factor > 0 And t(index).p1.X > -300 And t(index).p1.X < PictureBox1.Width + 300 And t(index).p1.Y > -300 And t(index).p1.Y < PictureBox1.Height + 300 Then
                    Dim points() As PointF = New PointF() {
                        New PointF(t(index).p1.X, t(index).p1.Y),
                        New PointF(t(index).p2.X, t(index).p2.Y),
                        New PointF(t(index).p3.X, t(index).p3.Y),
                        New PointF(t(index).p4.X, t(index).p4.Y)}
                    If My.Settings.antialias Then
                        Dim pBrush As New System.Drawing.Drawing2D.PathGradientBrush(points)
                        pBrush.SurroundColors = t(index).colors
                        pBrush.CenterColor = t(index).avgC
                        e.Graphics.FillPolygon(pBrush, points)
                    Else
                        e.Graphics.FillPolygon(New SolidBrush(t(index).avgC), points)
                    End If
                    If CheckBox9Op.Checked Then e.Graphics.DrawPolygon(Pens.Black, points)
                End If
            Next i
            If My.Settings.antialias Then e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            For i = 0 To all3DPoints(0).Length - 1 Step 2 'apply mouse rotation
                all3DPoints(0)(i) = all3DPoints(0)(i).RotateX(angleX).RotateY(angleY)
                all3DPoints(0)(i + 1) = all3DPoints(0)(i + 1).RotateX(angleX).RotateY(angleY)
                Dim projectedPoint1 As Point3D = all3DPoints(0)(i).project(PictureBox1.Width, PictureBox1.Height, fov, viewdistance)
                Dim projectedPoint2 As Point3D = all3DPoints(0)(i + 1).project(PictureBox1.Width, PictureBox1.Height, fov, viewdistance)
                Dim newAxisPen = New Pen(Brushes.Black, My.Settings.thickness)
                If i < 6 Then newAxisPen.DashPattern = New Single() {2.0F, 5.0F, 2.0F, 5.0F}
                e.Graphics.DrawLine(newAxisPen, New PointF(projectedPoint1.X, projectedPoint1.Y), New PointF(projectedPoint2.X, projectedPoint2.Y))
                If i = 0 Then e.Graphics.DrawString("x", New Font("Arial", 12), Brushes.Black, New PointF(projectedPoint2.X, projectedPoint2.Y))
                If i = 2 Then e.Graphics.DrawString("y", New Font("Arial", 12), Brushes.Black, New PointF(projectedPoint2.X, projectedPoint2.Y))
                If i = 4 Then e.Graphics.DrawString("z", New Font("Arial", 12), Brushes.Black, New PointF(projectedPoint2.X, projectedPoint2.Y))
            Next i
            angleX = 0
            angleY = 0
        Else '3D parametric
            If My.Settings.antialias Then e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            For j = 0 To all3DPoints.Count - 1 'loop through all lists of 3D points
                Dim projectedPoints(all3DPoints(j).length - 1) As PointF
                For i = 0 To all3DPoints(j).Length - 1 'apply mouse rotation
                    all3DPoints(j)(i) = all3DPoints(j)(i).RotateX(angleX).RotateY(angleY)
                    Dim projectedPoint As Point3D = all3DPoints(j)(i).project(PictureBox1.Width, PictureBox1.Height, fov, viewdistance)
                    projectedPoints(i) = New PointF(projectedPoint.X, projectedPoint.Y)
                Next i
                If j = 0 Then
                    For k = 0 To projectedPoints.Length - 1 Step 2
                        e.Graphics.DrawLine(axisPen, projectedPoints(k), projectedPoints(k + 1))
                    Next k
                    e.Graphics.DrawString("x", New Font("Arial", 12), Brushes.Black, projectedPoints(1))
                    e.Graphics.DrawString("y", New Font("Arial", 12), Brushes.Black, projectedPoints(3))
                    e.Graphics.DrawString("z", New Font("Arial", 12), Brushes.Black, projectedPoints(5))
                Else
                    e.Graphics.DrawLines(functionPen, projectedPoints)
                End If
            Next j
            If shouldDrawTrace Then
                Dim size As Single = My.Settings.thickness + 7
                traceCoord3D = traceCoord3D.RotateX(angleX).RotateY(angleY)
                Dim projectedPoint As Point3D = traceCoord3D.Project(PictureBox1.Width, PictureBox1.Height, fov, viewdistance)
                Dim tracePoint3D As PointF = New PointF(projectedPoint.X, projectedPoint.Y)
                e.Graphics.FillEllipse(Brushes.Green, tracePoint3D.X - size / 2, tracePoint3D.Y - size / 2, size, size)
                e.Graphics.DrawEllipse(Pens.White, tracePoint3D.X - size / 2, tracePoint3D.Y - size / 2, size, size)
                Dim tickWidth = Math.Max(Math.Min(Math.Pow(My.Settings.thickness, 1 / 4) * (PictureBox1.Height + PictureBox1.Width) / 150 / Math.Pow(((xmax - xmin) / xtick / 6), 1 / 4), 30), 3)
                Dim fontSize = Math.Sqrt(tickWidth) * 3.8
                Dim numberString As String = "t = " + traceNumber.ToString + vbNewLine + "x = " + traceCoord3DOriginal.X.ToString + vbNewLine + "y = " + traceCoord3DOriginal.Y.ToString + vbNewLine + "z = " + traceCoord3DOriginal.Z.ToString
                Dim box = e.Graphics.MeasureString(numberString, New Font("verdana", fontSize))
                Dim boxX, boxY, px, py 'determine dimensions and location of box around trace info
                If tracePoint3D.X < PictureBox1.Width / 2 Then
                    boxX = 30
                    px = 30 + box.Width / 2
                Else
                    boxX = PictureBox1.Width - box.Width - 30
                    px = PictureBox1.Width - box.Width / 2 - 30
                End If
                If tracePoint3D.Y > PictureBox1.Height / 2 Then
                    boxY = 30
                    py = 35 + box.Height
                Else
                    boxY = PictureBox1.Height - box.Height - 60
                    py = PictureBox1.Height - box.Height - 65
                End If
                e.Graphics.FillRectangle(Brushes.White, boxX - 5, boxY - 5, box.Width + 10, box.Height + 10)
                e.Graphics.DrawString(numberString, New Font("verdana", fontSize), Brushes.Black, New Point(boxX, boxY))
                e.Graphics.DrawRectangle(Pens.Black, boxX - 5, boxY - 5, box.Width + 10, box.Height + 10)
                e.Graphics.DrawLine(Pens.Black, New PointF(px, py), tracePoint3D)
            End If
            angleX = 0
            angleY = 0
        End If
    End Sub

    'converts hue, saturation, and brightness to a color
    Private Function HSVtoColor(hue As Double, saturation As Double, brightness As Double) As Color
        Dim r As Double, g As Double, b As Double
        Dim i As Single
        'validate hue value
        If hue > 360 Then
            Do Until hue <= 360
                hue = hue - 360
            Loop
        ElseIf hue < 0 Then
            Do Until hue >= 0
                hue = hue + 360
            Loop
        End If
        'Select hue
        Select Case hue
            Case 0 To 60
                r = 1 : b = 0
                g = hue / 60
            Case 60 To 120
                g = 1 : b = 0
                r = 1 - ((hue - 60) / 60)
            Case 120 To 180
                g = 1 : r = 0
                b = (hue - 120) / 60
            Case 180 To 240
                b = 1 : r = 0
                g = 1 - ((hue - 180) / 60)
            Case 240 To 300
                b = 1 : g = 0
                r = (hue - 240) / 60
            Case 300 To 360
                r = 1 : g = 0
                b = 1 - ((hue - 300) / 60)
        End Select
        'get intensity of saturation based on luminosity
        i = 1 - saturation
        'get colour channel values
        r = ((r * saturation) + i) * brightness
        g = ((g * saturation) + i) * brightness
        b = ((b * saturation) + i) * brightness
        HSVtoColor = Color.FromArgb(r * 255, g * 255, b * 255)
    End Function

    ' Form closing event - save things as needed
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        My.Settings.splitterDistance = SplitContainer1.SplitterDistance
        My.Settings.degreeMode = DegreeModeToolStripMenuItem.Checked
        My.Settings.Amin = TextBox2.Text
        My.Settings.Amax = TextBox3.Text
        If Me.WindowState = FormWindowState.Maximized Then
            My.Settings.isMaximized = True
        ElseIf Me.WindowState = FormWindowState.Minimized Then
            My.Settings.isMaximized = False
        Else
            My.Settings.isMaximized = False
            My.Settings.left = Me.Location.X
            My.Settings.top = Me.Location.Y
            My.Settings.width = Me.Width
            My.Settings.height = Me.Height
        End If
        ' save stuff to disk
        If My.Settings.saveOption = 1 Then
            If changesOccured Then
                Dim answer = MsgBox("Would you like to save your work before closing Photon?", vbYesNoCancel)
                If answer = vbYes Then
                    SaveToolStripMenuItem_Click(sender, e)
                ElseIf answer = vbNo Then
                    'close the program
                Else
                    e.Cancel = True
                End If
            End If
        ElseIf My.Settings.saveOption = 2 Then
            My.Settings.calcText = RichTextBox1.Text + "//" + xmin.ToString + "//" + xmax.ToString + "//" + ymin.ToString + "//" + ymax.ToString + "//" + xtick.ToString + "//" + ytick.ToString + "//" + LogXAxisToolStripMenuItem.Checked.ToString + "//" + LogYAxisToolStripMenuItem.Checked.ToString + "//" + transformX + "//" + transformY + "//" + transformZ
        Else
            'close the program
        End If
    End Sub
    ' Initializes things
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        System.Threading.Thread.CurrentThread.CurrentCulture = New Globalization.CultureInfo("en-US")
        Randomize()
        rndSeed = Rnd()
        Dim Thread1 As New System.Threading.Thread(AddressOf BuildUnitDictionary)
        Thread1.Start()
        xtick = 1
        ytick = 1
        MenuStrip1.ForeColor = Color.White
        xmin = My.Settings.x_min
        xmax = My.Settings.x_max
        ymin = My.Settings.y_min
        ymax = My.Settings.y_max
        TextBox2.Text = My.Settings.Amin
        TextBox3.Text = My.Settings.Amax
        ListBox1.SelectedIndex = 0
        XDataStats.Rows.Add("N")
        YDataStats.Rows.Add("N")
        XDataStats.Rows.Add("Mean")
        YDataStats.Rows.Add("Mean")
        XDataStats.Rows.Add("Sum")
        YDataStats.Rows.Add("Sum")
        XDataStats.Rows.Add("Sum of ^2")
        YDataStats.Rows.Add("Sum of ^2")
        XDataStats.Rows.Add("S stdev.")
        YDataStats.Rows.Add("S stdev.")
        XDataStats.Rows.Add("σ stdev.")
        YDataStats.Rows.Add("σ stdev.")
        XDataStats.Rows.Add("Min")
        YDataStats.Rows.Add("Min")
        XDataStats.Rows.Add("Q1")
        YDataStats.Rows.Add("Q1")
        XDataStats.Rows.Add("Med")
        YDataStats.Rows.Add("Med")
        XDataStats.Rows.Add("Q3")
        YDataStats.Rows.Add("Q3")
        XDataStats.Rows.Add("Max")
        YDataStats.Rows.Add("Max")
        XDataStats.ClearSelection()
        YDataStats.ClearSelection()
        XDataStats.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
        XDataStats.Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
        YDataStats.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
        YDataStats.Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
        DataXY.Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
        If My.Settings.isMaximized = True Then
            Me.WindowState = FormWindowState.Maximized
        Else
            Me.Width = My.Settings.width
            Me.Height = My.Settings.height
        End If
        RichTextBox1.Font = My.Settings.fontText
        Me.Location = New Point(My.Settings.left, My.Settings.top)
        SplitContainer1.Orientation = My.Settings.splitterOrientation
        SplitContainer1.SplitterDistance = My.Settings.splitterDistance
        DegreeModeToolStripMenuItem.Checked = My.Settings.degreeMode
        RadianModeToolStripMenuItem.Checked = My.Settings.degreeMode Xor True
        If My.Settings.projector Then
            RichTextBox1.BackColor = Color.White
            RichTextBox1.ForeColor = Color.Black
        Else
            RichTextBox1.BackColor = Color.Black
            RichTextBox1.ForeColor = Color.White
        End If
        If My.Settings.saveOption = 2 Then
            Dim s = Split(My.Settings.calcText, "//")
            RichTextBox1.Text = s(0)
            xmin = s(1)
            xmax = s(2)
            ymin = s(3)
            ymax = s(4)
            xtick = s(5)
            ytick = s(6)
            LogXAxisToolStripMenuItem.Checked = s(7)
            LogYAxisToolStripMenuItem.Checked = s(8)
            transformX = s(9)
            transformY = s(10)
            transformZ = s(11)
        Else
            My.Settings.calcText = ""
        End If
        FractionModeToolStripMenuItem.Checked = My.Settings.fraction
        If My.Settings.colorText = True Then
            ColorTheText()
        End If
        SimpleLinearToolStripMenuItem.Checked = False
        LinearToolStripMenuItem.Checked = False
        QuadraticToolStripMenuItem.Checked = False
        CubicToolStripMenuItem.Checked = False
        If My.Settings.drawMode = 0 Then
            SimpleLinearToolStripMenuItem.Checked = True
        ElseIf My.Settings.drawMode = 1 Then
            LinearToolStripMenuItem.Checked = True
        ElseIf My.Settings.drawMode = 2 Then
            SimpleQuadraticToolStripMenuItem.Checked = True
        ElseIf My.Settings.drawMode = 3 Then
            QuadraticToolStripMenuItem.Checked = True
        ElseIf My.Settings.drawMode = 4 Then
            SimpleCubicToolStripMenuItem.Checked = True
        Else
            CubicToolStripMenuItem.Checked = True
        End If
        Dim CommandLineArgs As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Application.CommandLineArgs
        If CommandLineArgs.Count <> 0 Then
            openFile(CommandLineArgs(0))
        End If
        changesOccured = False
        Button1.PerformClick()
        isFindingSlopeField = False
        'Speech recognition initialization
        Try
            reco.SetInputToDefaultAudioDevice()
            If System.IO.File.Exists(Application.StartupPath + "\grammar.xml") Then
                reco.LoadGrammarAsync(New System.Speech.Recognition.Grammar(Application.StartupPath + "\grammar.xml"))
            Else
                Dim objWriter As New System.IO.StreamWriter(Application.StartupPath + "\grammar.xml")
                objWriter.Write(My.Resources.grammar)
                objWriter.Close()
                reco.LoadGrammarAsync(New System.Speech.Recognition.Grammar(Application.StartupPath + "\grammar.xml"))
            End If
            If My.Settings.speechEnabled Then
                RichTextBox2.BackColor = Color.PaleGreen
                reco.RecognizeAsync(System.Speech.Recognition.RecognizeMode.Multiple)
            End If
        Catch
            Button2.Enabled = False
            My.Settings.speechEnabled = False
        End Try
        TextBox1Op.Text = My.Settings.x_min
        TextBox2Op.Text = My.Settings.x_max
        TextBox3Op.Text = My.Settings.y_min
        TextBox4Op.Text = My.Settings.y_max
        CheckBox1Op.Checked = My.Settings.autofix
        CheckBox2Op.Checked = My.Settings.fraction
        CheckBox6Op.Checked = My.Settings.grouping
        Button3Op.Font = RichTextBox1.Font
        CheckBox3Op.Checked = My.Settings.colorText
        CheckBox4Op.Checked = My.Settings.projector
        CheckBox10Op.Checked = My.Settings.highlighting
        If My.Settings.splitterOrientation = 1 Then
            RadioButton1Op.Checked = True
        Else
            RadioButton2Op.Checked = True
        End If
        If My.Settings.saveOption = 1 Then
            RadioButton3Op.Checked = True
        ElseIf My.Settings.saveOption = 2 Then
            RadioButton4Op.Checked = True
        Else
            RadioButton5Op.Checked = True
        End If
        TrackBar1Op.Value = My.Settings.thickness
        Label9Op.Text = My.Settings.thickness
        CheckBox5Op.Checked = My.Settings.antialias
        CheckBox7Op.Checked = My.Settings.speechEnabled
        CheckBox9Op.Checked = My.Settings.draw3DLines
        RadioButton6Op.Checked = My.Settings.autoCalculate
        RadioButton7Op.Checked = Not My.Settings.autoCalculate
        CheckBox8Op.Checked = My.Settings.autoDelete
        TextBox10.Text = My.Settings.userFunctions
        RichTextBox2.SelectionAlignment = HorizontalAlignment.Center
        If Button2.Enabled = False Then
            CheckBox7Op.Enabled = False
            CheckBox7Op.Text = "Speech is enabled (unavailable; no mic detected)"
        End If
        TabControl1Op.Tag = "yes"
    End Sub

    ' What happens when the user clicks "Go" or hits the ENTER key.
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        variableCreate = False
        If openParenth > 1 And My.Settings.speechEnabled Then
            Exit Sub
        End If
        While openParenth > 0
            RichTextBox2.Text += ")"
            openParenth -= 1
        End While
        If My.Settings.autoDelete Then
            RichTextBox1.Text = ""
        End If
        If RichTextBox2.Text = "Photon" Then
            PictureBox1.Image = PictureBox2.Image
            RichTextBox2.Text = ""
            Exit Sub
        End If
        If RichTextBox2.Text.StartsWith("A = ") Then Exit Sub
        Dim traceText = ""
        RichTextBox2.Tag = ""
        Dim chPos = RichTextBox1.SelectionStart
        If RichTextBox2.Text.StartsWith("trace") Then
            traceText = RichTextBox2.Text
            shouldTrace = True
            RichTextBox2.Text = ""
        ElseIf RichTextBox1.Text.StartsWith("trace") Then
            shouldTrace = True
            traceText = RichTextBox1.Lines(0)
        End If
        If RichTextBox1.Text = "" And RichTextBox2.Text <> "" Then
            RichTextBox1.Text = heal(RichTextBox2.Text)
            changesOccured = True
        ElseIf RichTextBox2.Text <> "" Then
            RichTextBox1.Text = RichTextBox1.Text + vbNewLine + heal(RichTextBox2.Text)
            chPos = RichTextBox1.Text.Length - 1
            changesOccured = True
        End If
        ' need to search for double variable creations and more than one differential slope field
        Dim bad = -1
        For j = 0 To RichTextBox1.Lines.Length - 1
            If RichTextBox1.Lines(j).StartsWith("trace") Then 'loop through and find another "trace"; delete it if it's found
                bad = j
                Exit For
            End If
            If RichTextBox1.Lines(j).Contains(":=") Then
                For i = 0 To RichTextBox1.Lines.Length - 1
                    Dim sArray = Split(RichTextBox1.Lines(i), ":=")
                    If sArray.Length > 1 And RichTextBox1.Lines(i).StartsWith(RichTextBox1.Lines(j).Substring(0, 1), vbTextCompare) And i <> j Then
                        bad = j
                        Exit For
                    End If
                Next i
                If bad <> -1 Then
                    Exit For
                End If
            ElseIf RichTextBox1.Lines(j).StartsWith("y' = ") Then
                For i = 0 To RichTextBox1.Lines.Length - 1
                    If RichTextBox1.Lines(i).StartsWith("y' = ") And i <> j Then
                        bad = j
                        Exit For
                    End If
                Next i
                If bad <> -1 Then
                    Exit For
                End If
            ElseIf RichTextBox1.Lines(j).StartsWith("z = ") Then
                For i = 0 To RichTextBox1.Lines.Length - 1
                    If RichTextBox1.Lines(i).StartsWith("z = ") And i <> j Then
                        bad = j
                        Exit For
                    End If
                Next i
                If bad <> -1 Then
                    Exit For
                End If
            End If
        Next j
        If bad <> -1 Then
            Dim outArr(RichTextBox1.Lines.Length - 2) As String
            Array.Copy(RichTextBox1.Lines, 0, outArr, 0, bad)
            Array.Copy(RichTextBox1.Lines, bad + 1, outArr, bad, RichTextBox1.Lines.Length - bad - 1)
            RichTextBox1.Lines = outArr
        End If
        RichTextBox2.Text = ""
        If traceText = "trace" Then traceText = "trace A"
        If shouldTrace Then RichTextBox1.Text = traceText + vbNewLine + RichTextBox1.Text
        Dim text1 = RichTextBox1.Text
        If text1 = "" Then text1 = "b"
        If Replace(Replace(text1, "iAdd", ""), "toMagAng", "").Contains("A") Then
            TrackBar1.Visible = True
            TextBox2.Visible = True
            TextBox3.Visible = True
        Else
            TrackBar1.Visible = False
            TextBox2.Visible = False
            TextBox3.Visible = False
        End If
        RichTextBox1.Lines = RichTextBox1.Text.Split(New Char() {ControlChars.Lf}, StringSplitOptions.RemoveEmptyEntries)
        zFactor = 1
        facesAreColored = False
        all3DFaces.Clear()
        last3DGraph = ""
        clear3DPathArrays()
        If Me.Tag = "graph" Then
            PictureBox1.Refresh() 'call graphics
        Else
            calculateThingsBeforeDrawing()
        End If
        RichTextBox1.SelectionStart = chPos
        RichTextBox1.ScrollToCaret()
    End Sub
    ' Erase text from textbox if the user has been tracing
    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles RichTextBox2.GotFocus, RichTextBox1.GotFocus
        If RichTextBox2.Text.StartsWith("A = ") Then RichTextBox2.Text = ""
    End Sub
    ' Resets tag to the character before the cursor if the user changes the location with the mouse
    Private Sub TextBox1_MouseUp(sender As Object, e As MouseEventArgs) Handles RichTextBox2.MouseUp, ToolStripTextBox7.MouseUp, ToolStripTextBox8.MouseUp
        Dim selection = sender.SelectionStart
        If selection <= 0 Or selection > sender.Text.Length Then
            sender.Tag = ""
            Exit Sub
        End If
        Dim myChar = sender.Text.Chars(selection - 1)
        If Char.IsDigit(myChar) Then 'number
            sender.tag = "*"
        ElseIf myChar.IsLetter(myChar) Then 'letter
            If myChar <> "E" Then
                sender.Tag = "^"
            Else
                sender.Tag = "("
            End If
        ElseIf myChar = ")" Then ')
            sender.Tag = ")"
        ElseIf myChar = "^" Or myChar = "," Or myChar = "(" Then
            sender.Tag = "("
        ElseIf myChar = "/" Then '/
            sender.Tag = "/"
        Else
            sender.tag = ""
        End If
    End Sub
    ' Automatically adds spaces, multiplication signs, and power signs to allow ease of equation input
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox2.KeyDown, ToolStripTextBox7.KeyDown, ToolStripTextBox8.KeyDown, ToolStripTextBox9.KeyDown
        If (e.Control And e.KeyCode = Keys.V) Then
            sender.Paste(DataFormats.GetFormat("Text"))
            e.Handled = True
        End If
        If sender.Text.Contains("&H") Or sender.Text.Contains("&O") Then Exit Sub
        Application.DoEvents()
        Dim selection = sender.SelectionStart
        If e.KeyCode = 37 Then selection -= 1
        If e.KeyCode = 39 Then selection += 1
        If selection <= 0 Or selection > sender.Text.Length Then
            sender.Tag = ""
            Exit Sub
        End If
        Dim myChar = sender.Text.Chars(selection - 1)
        If Char.IsDigit(myChar) Then 'number
            If (sender.Tag = "^" Or sender.Tag = ")" Or sender.tag = "A") And Char.IsDigit(Chr(e.KeyCode)) Then
                sender.Text = sender.Text.Insert(sender.SelectionStart - 1, "^")
                sender.SelectionStart = selection + 1
            End If
            sender.Tag = "*"
        ElseIf myChar = "A" Then ' "A" is a letter / number hybrid
            If (sender.Tag = "*" Or sender.Tag = ")") And Chr(e.KeyCode) = "A" Then 'previous is number
                sender.Text = sender.Text.Insert(sender.SelectionStart - 1, "*")
                sender.SelectionStart = selection + 1
            ElseIf (sender.Tag = "^" Or sender.Tag = ")") And Chr(e.KeyCode) = "A" Then 'previous is letter
                sender.Text = sender.Text.Insert(sender.SelectionStart - 1, "^")
                sender.SelectionStart = selection + 1
            End If
            sender.tag = "A"
        ElseIf Char.IsLetter(myChar) Then 'letter
            If myChar <> "E" Then
                If (sender.Tag = "*" Or sender.Tag = ")" Or sender.tag = "A") And Char.IsLetter(Chr(e.KeyCode)) Then
                    sender.Text = sender.Text.Insert(sender.SelectionStart - 1, "*")
                    sender.SelectionStart = selection + 1
                End If
                sender.Tag = "^"
            Else
                sender.Tag = "("
            End If
        ElseIf myChar = ")" Then ')
            sender.Tag = ")"
        ElseIf myChar = "^" Or myChar = "," Then '^
            sender.Tag = "("
        ElseIf myChar = "(" Then '(
            If (sender.Tag = "*" Or sender.Tag = ")") And ChrW(e.KeyCode) = "9" Then
                sender.Text = sender.Text.Insert(sender.SelectionStart - 1, "*")
                sender.SelectionStart = selection + 1
            End If
            sender.Tag = "("
        ElseIf myChar = "/" Then '/
            sender.Tag = "/"
        ElseIf (myChar = "=" Or myChar = "-" Or myChar = "+") And sender.Tag <> "(" And sender.Tag <> "/" Then 'put spaces around them
            If (e.KeyCode = 189 Or e.KeyCode = 187 Or e.KeyCode = 109 Or e.KeyCode = 107) And sender.SelectionStart > 1 Then '-, + or numpad -, +, or =
                If sender.Text.Substring(sender.SelectionStart - 2, 1) <> " " Then
                    sender.Text = sender.Text.Insert(sender.SelectionStart - 1, " ")
                    sender.SelectionStart = selection + 1
                    sender.Text = sender.Text.Insert(sender.SelectionStart, " ")
                    sender.SelectionStart = selection + 2
                End If
            End If
            sender.Tag = ""
        ElseIf myChar = ":" Then 'change ":" to " := "
            If e.KeyCode = 186 And sender.SelectionStart > 1 Then
                If sender.Text.Substring(sender.SelectionStart - 2, 1) <> " " Then
                    sender.Text = sender.Text.Insert(sender.SelectionStart - 1, " ")
                    sender.SelectionStart = selection + 1
                    sender.Text = sender.Text.Insert(sender.SelectionStart, "= ")
                    sender.SelectionStart = selection + 3
                Else
                    sender.Text = sender.Text.Insert(sender.SelectionStart, "=")
                    sender.SelectionStart = selection + 1
                End If
            End If
            sender.Tag = ""
        ElseIf myChar = ";" Then 'change ";" to "; "
            If e.KeyCode = 186 Then
                sender.Text = sender.Text.Insert(sender.SelectionStart, " ")
                sender.SelectionStart = selection + 1
            End If
            sender.Tag = ""
        ElseIf myChar = "?" Then
            If e.KeyCode = 191 Then
                If sender.Text.Contains("= ") = False Then
                    sender.Text = "(" + sender.Text.Substring(0, sender.Text.Length - 1) + ")/()"
                Else
                    sender.Text = sender.Text.Substring(0, sender.Text.Length - 1).Insert(sender.Text.IndexOf("= ") + 2, "(") + ")/()"
                End If
                If sender.Text.EndsWith("()/()") Then
                    sender.SelectionStart = selection
                Else
                    sender.SelectionStart = selection + 3
                End If
            End If
            sender.Tag = ""
        Else
            sender.Tag = ""
        End If
    End Sub
    ' Sends focus to the textbox if there is nothing in the rich text box
    Private Sub RichTextBox1_GotFocus(sender As Object, e As EventArgs) Handles RichTextBox1.GotFocus
        If RichTextBox1.Text = "" Then RichTextBox2.Focus()
    End Sub
    ' Checks to see if more than one character should be erased when the backspace key is pressed (ex: "2 - |3"  -->  "2|3") AND checks to see if "ans" should be inserted
    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles RichTextBox2.KeyPress, ToolStripTextBox7.KeyPress, ToolStripTextBox8.KeyPress, ToolStripTextBox9.KeyPress
        If sender.SelectionStart <= 2 Then
            If ((e.KeyChar = "+" Or e.KeyChar = "/" Or e.KeyChar = "*" Or e.KeyChar = "^") And sender.Text = "") Or (e.KeyChar = "-" And RichTextBox2.Text = "-") Then
                sender.Text = "ans"
                sender.SelectionStart = 3
            End If
            Exit Sub
        End If
        Dim sTemp1 = sender.Text.Substring(sender.SelectionStart - 1, 1)
        Dim sTemp2 = sender.Text.Substring(sender.SelectionStart - 2, 1)
        Dim sTemp3 = sender.Text.Substring(sender.SelectionStart - 3, 1)
        If e.KeyChar = vbBack Then
            If sTemp1 = " " And sTemp3 = " " And (sTemp2 = "-" Or sTemp2 = "+" Or sTemp2 = "=") Then
                Dim temp = sender.SelectionStart
                sender.Text = sender.Text.Remove(sender.SelectionStart - 2, 2)
                sender.SelectionStart = temp - 2
            ElseIf sTemp1 = " " And sTemp2 = "=" And sTemp3 = ":" Then
                If sender.SelectionStart <= 3 Then Exit Sub
                If sender.Text.Substring(sender.SelectionStart - 4, 1) = " " Then
                    Dim temp = sender.SelectionStart
                    sender.Text = sender.Text.Remove(sender.SelectionStart - 3, 3)
                    sender.SelectionStart = temp - 3
                End If
            ElseIf (sTemp1 = "=" And sTemp2 = ":") Or (sTemp1 = " " And sTemp2 = ";") Then
                Dim temp = sender.SelectionStart
                sender.Text = sender.Text.Remove(sender.SelectionStart - 1, 1)
                sender.SelectionStart = temp - 1
            End If
        End If
    End Sub

    'highlights maching parentheses
    Private Sub RichTextBox2_HighlightParentheses(sender As Object)
        Dim loc As Integer = sender.SelectionStart
        Dim selLength As Integer = sender.SelectionLength
        sender.SelectAll()
        sender.SelectionBackColor = sender.BackColor
        If loc > 0 Then 'check for ) before cursor
            If sender.Text(loc - 1) = ")" Then
                sender.SelectionStart = loc - 1
                sender.SelectionLength = 1
                sender.SelectionBackColor = Color.Yellow
                Dim pCount As Integer = 1
                Dim n As Integer = loc - 2
                While pCount > 0 And n >= 0
                    If sender.text(n) = vbLf Then Exit While
                    If sender.Text(n) = ")" Then
                        pCount += 1
                    ElseIf sender.Text(n) = "(" Then
                        pCount -= 1
                    End If
                    n -= 1
                End While
                If pCount = 0 Then
                    sender.SelectionStart = n + 1
                    sender.SelectionLength = 1
                    sender.SelectionBackColor = Color.Yellow
                Else
                    sender.SelectionStart = loc - 1
                    sender.SelectionLength = 1
                    sender.SelectionBackColor = Color.Red
                End If
            End If
        End If
        If loc < sender.Text.Length Then 'check for ( after cursor
            If sender.Text(loc) = "(" Then
                sender.SelectionStart = loc
                sender.SelectionLength = 1
                sender.SelectionBackColor = Color.Yellow
                Dim pCount As Integer = 1
                Dim n As Integer = loc + 1
                While pCount > 0 And n < sender.Text.Length
                    If sender.text(n) = vbLf Then Exit While
                    If sender.Text(n) = ")" Then
                        pCount -= 1
                    ElseIf sender.Text(n) = "(" Then
                        pCount += 1
                    End If
                    n += 1
                End While
                If pCount = 0 Then
                    sender.SelectionStart = n - 1
                    sender.SelectionLength = 1
                    sender.SelectionBackColor = Color.Yellow
                Else
                    sender.SelectionStart = loc
                    sender.SelectionLength = 1
                    sender.SelectionBackColor = Color.Red
                End If
            End If
        End If
        sender.SelectionStart = loc
        sender.SelectionLength = selLength
    End Sub

    Private Sub RichTextBox2_KeyUp(sender As Object, e As KeyEventArgs) Handles RichTextBox2.KeyUp, RichTextBox1.KeyUp
        If My.Settings.highlighting And e.Shift = False Then RichTextBox2_HighlightParentheses(sender)
    End Sub

    Private Sub RichTextBox2_Click(sender As Object, e As EventArgs) Handles RichTextBox2.Click, RichTextBox1.Click
        If My.Settings.highlighting Then RichTextBox2_HighlightParentheses(sender)
    End Sub

    ' "Heals" the text by removing excess spaces, adding appropriate spaces, and adding * and ^.
    Function heal(s As String, Optional auto As Boolean = True) As String
        If s.Contains("&H") Or s.Contains("&O") Or s.Contains("trace") Then
            heal = s
            Exit Function
        End If
        Dim text = s
        text = Replace(text, " ", "")
        text = Replace(text, "–", "-")
        If text = "" Then text = "0"
        If text.Contains(":=") = False And text.Contains("=") = True Then
            If text.StartsWith("dd") Then
                text = text.Insert(2, " ")
            ElseIf text.StartsWith("d") Then
                text = text.Insert(1, " ")
            ElseIf text.StartsWith("i") Then
                text = text.Insert(1, " ")
            End If
        End If
        Dim tempString = Split(Replace(text, "√", "sqrt"), "->")
        If tempString.Length > 1 And shouldReset = False And My.Settings.fraction = True Then
            Dim firstAnswer = eval(tempString(0), False, False)
            If Math.Abs(Val(firstAnswer) - Val(eval(tempString(1), False, False))) > eps Or IsNumeric(firstAnswer) = False Then
                text = tempString(0)
            Else
                text = Replace(text, "->", "<√>")
                ans = tempString(1)
            End If
        Else
            text = tempString(0)
        End If
        If text = "" Then text = "0"
        text = Replace(text, "pi", "π")
        text = Replace(text, "th", "θ")
        text = Replace(text, "coθ(", "coth(")
        Dim firstLetter As Char
        Dim secondLetter As Char
        Dim i = 0
        If auto = True And text.Contains("=") = False Then
            Dim text1 = text
            text1 = Replace(text1, "exp", "", 1, -1, vbTextCompare)   ' all these functions throw off auto-insert
            text1 = Replace(text1, "msgbox", "", 1, -1, vbTextCompare)
            text1 = Replace(text1, "fix", "", 1, -1, vbTextCompare)
            text1 = Replace(text1, "hex", "", 1, -1, vbTextCompare)
            text1 = Replace(text1, "rnd", "", 1, -1, vbTextCompare)
            text1 = Replace(text1, "round", "", 1, -1, vbTextCompare)
            text1 = Replace(text1, "torad", "", 1, -1, vbTextCompare)
            text1 = Replace(text1, "sqr", "", 1, -1, vbTextCompare)
            text1 = Replace(text1, "frac", "", 1, -1, vbTextCompare)
            text1 = Replace(text1, "prime", "", 1, -1, vbTextCompare)
            text1 = Replace(text1, "cross", "", 1, -1, vbTextCompare)
            text1 = Replace(text1, "atanxy", "", 1, -1, vbTextCompare)
            If text1 <> "" Then
                If text1.Contains("x") And text1.Contains("y") And RichTextBox1.Text.Contains("z :=") = False Then
                    text = "z = " + text
                ElseIf text1.Contains("x") And RichTextBox1.Text.Contains("x :=") = False Then
                    text = "y = " + text
                ElseIf text1.Contains("θ") And RichTextBox1.Text.Contains("θ :=") = False Then
                    text = "r = " + text
                ElseIf text1.Contains("y") And RichTextBox1.Text.Contains("y :=") = False Then
                    text = "x = " + text
                ElseIf text1.Contains("r") And RichTextBox1.Text.Contains("r :=") = False Then
                    text = "θ = " + text
                End If
            End If
        End If
        While (i < text.Length - 1)
            firstLetter = text.Substring(i, 1)
            secondLetter = text.Substring(i + 1, 1)
            If ((Char.IsLetter(firstLetter) And firstLetter <> "E") Or firstLetter = ")") And Char.IsNumber(secondLetter) Then
                text = text.Insert(i + 1, "^")
            ElseIf (Char.IsNumber(firstLetter) Or firstLetter = ")") And ((Char.IsLetter(secondLetter) And secondLetter <> "E") Or secondLetter = "(") Then
                text = text.Insert(i + 1, "*")
            ElseIf (firstLetter <> " " And firstLetter <> ":") And (secondLetter = "+" Or secondLetter = "-" Or secondLetter = "=" Or secondLetter = ":") Then
                text = text.Insert(i + 1, " ")
            ElseIf (secondLetter <> " ") And (firstLetter = "+" Or firstLetter = "-" Or firstLetter = "=" Or firstLetter = ";") Then
                text = text.Insert(i + 1, " ")
            End If
            i += 1
        End While
        If text.StartsWith("-") And text.Length > 1 Then
            text = text.Remove(1, 1)
        End If
        text = Replace(text, "+ - ", "+ -")
        text = Replace(text, "= - ", "= -")
        text = Replace(text, "; - ", "; -")
        text = Replace(text, ", - ", ",-")
        text = Replace(text, "( - ", "(-")
        text = Replace(text, "( + ", "(+")
        text = Replace(text, "^ - ", "^-")
        text = Replace(text, "E + ", "E+")
        text = Replace(text, "E - ", "E-")
        text = Replace(text, "* - ", "*-")
        text = Replace(text, "/ - ", "/-")
        text = Replace(text, "<√>", " -> ")
        text = Replace(text, " ->  - ", " -> -")
        text = Replace(text, "xy", "x*y")
        text = Replace(text, "atanx*y", "atanxy")
        heal = text
    End Function
    ' Main calculation area - distinguishes between a variable, function, and calculation, also differential
    Private Function calculate(text As String)
        If text.EndsWith(";") Then
            text = text.Substring(0, text.Length - 1)
        End If
beginning:
        Dim s = Replace(text, " ", "")
        Dim result = ""
        If (s.Contains(":=")) Then
            Dim sArray = Split(s, ":=")
            If (Char.IsLetter(sArray(0)) And sArray(0).Length = 1 And sArray(0) <> "π" And String.Equals(sArray(0), "e", vbTextCompare) = False) Then
                If sArray(1) = "" Then
                    sArray(1) = "0"
                End If
                'result = sArray(0) + " := " + eval(sArray(1)).ToString  ' <- does not allow for persisting variables
                result = sArray(0) + " := " + heal(sArray(1), False)
            Else
                If sArray(0).Length >= 4 Then
                    If (sArray(0).Substring(1, 1) = "(" And sArray(0).EndsWith(")") And Char.IsLetter(sArray(0).Substring(0, 1))) Then
                        If sArray(1) = "" Then
                            sArray(1) = "0"
                        End If
                        'result = sArray(0) + " := " + heal(Split(text, ":=")(1))
                        result = sArray(0) + " := " + heal(sArray(1), False)
                    Else
                        result = text + " -> Invalid variable name."
                    End If
                Else
                    If sArray(0).StartsWith("'") Then
                        result = text
                    Else
                        result = text + " -> Invalid variable name."
                    End If
                End If
            End If
        ElseIf s.StartsWith("y'=") Then
            Dim tempstring = Split(s, "=")
            Dim sField As String = tempstring(1)
            Dim increment = xtick / (xmax - xmin) * PictureBox1.Width
            adjustXTickSpacing(increment) 'we need to do this here because the tick spacing determines where we put the slope field elements
            increment = ytick / (ymax - ymin) * PictureBox1.Height
            adjustYTickSpacing(increment)
            Dim sWidth As Integer = Math.Ceiling((xmax - xmin) / xtick) * 2
            Dim sHeight As Integer = Math.Ceiling((ymax - ymin) / ytick) * 2
            ReDim sFieldLocationArray(sHeight, sWidth)
            ReDim sFieldSlopeArray(sHeight, sWidth)
            isFindingSlopeField = True
            eval(sField, False, False)
            isFindingSlopeField = False
            If isSlopeFieldSolve = True Then
                multiEval(sField, xmin, xmax, 11)
                isSlopeFieldSolve = False
            End If
            result = text
        ElseIf (s.Contains("=")) Then 'user wishes to graph a function
            result = text
            If s.StartsWith("dd") Then
                secondDerivative = True
                s = s.Substring(2, s.Length - 2)
            ElseIf s.StartsWith("d") Then
                firstDerivative = True
                s = s.Substring(1, s.Length - 1)
            ElseIf s.StartsWith("i") Then
                shouldIntegrate = True
                s = s.Substring(1, s.Length - 1)
            End If
            Dim tempString = Split(s, ";")
            Dim zMin = 0.0
            Dim zMax = 0.0
            Dim defaultBounds = False
            If tempString.Length <= 2 Or (s.Contains("x=") And s.Contains("y=") And s.Contains("z=") And tempString.Length = 3) Then
                defaultBounds = True
            Else
                If tempString(tempString.Length - 1).Contains("=") = False And tempString(tempString.Length - 2).Contains("=") = False Then
                    zMin = Math.Min(Val(eval(tempString(tempString.Length - 1), False, False)), Val(eval(tempString(tempString.Length - 2), False, False)))
                    zMax = Math.Max(Val(eval(tempString(tempString.Length - 1), False, False)), Val(eval(tempString(tempString.Length - 2), False, False)))
                    If zMin = zMax Then
                        zMax = zMin + (xmax - xmin) / PictureBox1.Width
                    End If
                Else
                    calculate = text + " -> error"
                    Exit Function
                End If
            End If
            'If tempString.Length = 1 Or (tempString.Length = 3 And defaultBounds = False) Then ' function like y = f(x), x = f(y), z = f(x, y), r = f(th), th = f(r), or f(x,y) = g(x,y)
            If tempString.Length = 1 Or tempString.Length = 3 Or tempString.Length = 5 Then ' function like y = f(x), x = f(y), z = f(x, y), r = f(th), th = f(r), or f(x,y) = g(x,y)
                Dim brokenDown = Split(tempString(0), "=")
                If brokenDown(0).Contains("(") And brokenDown(0).Length = 4 Then 'we're actually creating a function variable so we don't belong here
                    text = brokenDown(0) + ":=" + brokenDown(1)
                    GoTo beginning
                End If
                If String.Equals(brokenDown(0), "y", vbTextCompare) Then 'y = f(x)
                    If defaultBounds = True Then
                        zMin = xmin
                        zMax = xmax
                    End If
                    multiEval(brokenDown(1), zMin, zMax, 1)
                ElseIf String.Equals(brokenDown(0), "x", vbTextCompare) Then 'x = f(y) OR 3D parametric: x = f(t); y = g(t); z = h(t)
                    If defaultBounds = True Then
                        zMin = ymin
                        zMax = ymax
                    End If
                    If s.Contains("y=") And s.Contains("z=") And tempString.Length >= 3 Then '3D parametric
                        multiEval(brokenDown(1) + ";" + tempString(1).Split("=")(1) + ";" + tempString(2).Split("=")(1), zMin, zMax, 12)
                    Else
                        multiEval(brokenDown(1), zMin, zMax, 2) 'x = f(y)
                    End If
                ElseIf String.Equals(brokenDown(0), "z", vbTextCompare) Then 'z = f(x, y) - 3D graph
                    If defaultBounds = True Then
                        zMin = ymin
                        zMax = ymax
                    End If
                    If last3DGraph <> brokenDown(1) Then
                        all3DFaces.Clear()
                        facesAreColored = False
                        multiEval(brokenDown(1), zMin, zMax, 10)
                    End If
                ElseIf String.Equals(brokenDown(0), "r", vbTextCompare) Then 'r = f(th)
                    If defaultBounds = True Then
                        zMin = 0
                        If My.Settings.degreeMode = False Then
                            zMax = 2 * Math.PI
                        Else
                            zMax = 360
                        End If
                    End If
                    multiEval(brokenDown(1), zMin, zMax, 3)
                ElseIf String.Equals(brokenDown(0), "θ", vbTextCompare) Then 'th = f(r)
                    If defaultBounds = True Then
                        zMin = 0
                        If My.Settings.degreeMode = False Then
                            zMax = 2 * Math.PI
                        Else
                            zMax = 360
                        End If
                    End If
                    multiEval(brokenDown(1), zMin, zMax, 4)
                Else                                                       'f(x,y) = g(x,y) or f(th,r) = g(th,r)
                    Dim implicitFunction As String = brokenDown(0) + "-(" + brokenDown(1) + ")"
                    If implicitFunction.Contains("y") Then
                        If defaultBounds = True Then
                            zMin = xmin
                            zMax = xmax
                        End If
                        multiEval(implicitFunction, zMin, zMax, 7)
                        multiEval(implicitFunction, zMin, zMax, 9)
                    ElseIf implicitFunction.Contains("r") Then
                        If defaultBounds = True Then
                            zMin = 0
                            If My.Settings.degreeMode = False Then
                                zMax = 2 * Math.PI
                            Else
                                zMax = 360
                            End If
                        End If
                        multiEval(implicitFunction, zMin, zMax, 8)
                    End If
                End If
            ElseIf tempString.Length = 2 Or (tempString.Length = 4 And defaultBounds = False) Then 'parametric
                Dim brokenDown1 = Split(tempString(0), "=")
                Dim brokenDown2 = Split(tempString(1), "=")
                If String.Equals(brokenDown1(0), "x", vbTextCompare) And String.Equals(brokenDown2(0), "y", vbTextCompare) Then 'x = f(t), y = h(t)
                    If defaultBounds = True Then
                        zMin = xmin
                        zMax = xmax
                    End If
                    multiEval(brokenDown1(1), zMin, zMax, 5, brokenDown2(1))
                ElseIf String.Equals(brokenDown1(0), "y", vbTextCompare) And String.Equals(brokenDown2(0), "x", vbTextCompare) Then 'y = f(t), x = h(t)
                    If defaultBounds = True Then
                        zMin = xmin
                        zMax = xmax
                    End If
                    multiEval(brokenDown2(1), zMin, zMax, 5, brokenDown1(1))
                ElseIf String.Equals(brokenDown1(0), "r", vbTextCompare) And String.Equals(brokenDown2(0), "θ", vbTextCompare) Then 'r = f(t), th = h(t)
                    If defaultBounds = True Then
                        zMin = 0
                        If My.Settings.degreeMode = False Then
                            zMax = 2 * Math.PI
                        Else
                            zMax = 360
                        End If
                    End If
                    multiEval(brokenDown2(1), zMin, zMax, 6, brokenDown1(1))
                ElseIf String.Equals(brokenDown1(0), "θ", vbTextCompare) And String.Equals(brokenDown2(0), "r", vbTextCompare) Then 'th = f(t), r = h(t)
                    If defaultBounds = True Then
                        zMin = 0
                        If My.Settings.degreeMode = False Then
                            zMax = 2 * Math.PI
                        Else
                            zMax = 360
                        End If
                    End If
                    multiEval(brokenDown1(1), zMin, zMax, 6, brokenDown2(1))
                End If
            End If
        Else 'user wishes to perform a calculation or trace
            If s.StartsWith("trace") Then
                result = text
            Else
                Dim st As String = eval(s, My.Settings.fraction, True)
                If st = "" Then st = "0"
                result = text + " -> " + st
            End If
        End If
        calculate = result
    End Function

    Private Function evalFix(s As String) As String
        Dim expr = s
        expr = Replace(expr, "sin(", "f1(", 1, -1, vbTextCompare)
        expr = Replace(expr, "cos(", "f2(", 1, -1, vbTextCompare)
        expr = Replace(expr, "tan(", "f3(", 1, -1, vbTextCompare)
        expr = Replace(expr, "atn(", "af3(", 1, -1, vbTextCompare)
        expr = Replace(expr, "atanxy(", "af4(", 1, -1, vbTextCompare)
        expr = Replace(expr, "log(", "log10(", 1, -1, vbTextCompare)
        expr = Replace(expr, "mod(", "modulo(", 1, -1, vbTextCompare)
        expr = Replace(expr, "sqr(", "sqrt(", 1, -1, vbTextCompare)
        expr = Replace(expr, "E-", "<E>", 1, -1, vbTextCompare)
        expr = Replace(expr, "^-", "<^>", 1, -1, vbTextCompare)
        expr = Replace(expr, "/-", "</>", 1, -1, vbTextCompare)
        expr = Replace(expr, "-", "-1*", 1, -1, vbTextCompare)
        expr = Replace(expr, "<E>", "E-", 1, -1, vbTextCompare)
        expr = Replace(expr, "<^>", "^-", 1, -1, vbTextCompare)
        expr = Replace(expr, "</>", "/-", 1, -1, vbTextCompare)
        evalFix = expr
    End Function

    ' evaluates multiple values of a function and stores the points in an array to be plotted
    Private Sub multiEval(expr As String, zMin As Double, zMax As Double, type As Integer, Optional expr2 As String = "0")
        On Error GoTo errorStop
        If type = 12 Then
            For Each s In all3DFunctionNames
                If expr = s Then Exit Sub
            Next s
            all3DFunctionNames.Add(expr)
        End If
        Dim originalExpr As String = expr
        Dim code As String = ""
        Dim res3D = 6 'graphing 3D resolution; higher is finer detail
        If CheckBox9Op.Checked Then res3D = 3
        Dim halfSqWidth = res3D * 2
        Dim sc = New MSScriptControl.ScriptControl
        sc.Language = "VBScript"
        sc.AllowUI = True
        Dim illegalVariable As String = "'"
        expr = evalFix(expr)
        expr2 = evalFix(expr2)
        Dim transformX_Final = evalFix(transformX)
        Dim transformY_Final = evalFix(transformY)
        Dim transformZ_Final = evalFix(transformZ)
        Dim counter = 0
        Dim outString = ""
        For Each letter As String In expr
            If counter <> expr.Length - 1 Then
                If letter = "A" And Char.IsLetter(expr(counter + 1)) = False Then letter = "AA"
            Else
                If letter = "A" Then letter = "AA"
            End If
            outString += letter
            counter += 1
        Next letter
        expr = outString
        counter = 0
        outString = ""
        For Each letter As String In expr2
            If counter <> expr2.Length - 1 Then
                If letter = "A" And Char.IsLetter(expr2(counter + 1)) = False Then letter = "AA"
            Else
                If letter = "A" Then letter = "AA"
            End If
            outString += letter
            counter += 1
        Next letter
        expr2 = outString
        Dim toDegrees As String
        Dim numberToDegrees As Double
        If DegreeModeToolStripMenuItem.Checked = True Then
            toDegrees = "PI / 180"
            numberToDegrees = Math.PI / 180
        Else
            toDegrees = "1"
            numberToDegrees = 1
        End If
        Dim variables As String = ""
        Dim functions As String = ""
        variables += "ans = " + ans.ToString + vbNewLine
        If TrackBar1.Visible = True Then variables += "AA = " + midA.ToString + vbNewLine
        For i = 0 To RichTextBox1.Lines.Length - 1
            If RichTextBox1.Lines(i).Contains(":=") And RichTextBox1.Lines(i).StartsWith(illegalVariable) = False And RichTextBox1.Lines(i).StartsWith("e", vbTextCompare) = False And RichTextBox1.Lines(i).Contains("->") = False Then
                Dim info = Split(RichTextBox1.Lines(i).Replace(" ", ""), ":=")
                Dim myVar = Replace(info(0), "θ", "THET")
                If myVar.Length = 1 Or myVar = "THET" Then
                    variables = variables + myVar + " = " + evalFix(heal(info(1).Replace("A", "AA"), False)) + vbNewLine
                Else
                    functions = functions + "Public Function " + myVar + vbNewLine + myVar.First + " = " + evalFix(heal(info(1), False)) + vbNewLine + "End Function" + vbNewLine
                End If
            End If
        Next i
        code = "PI = 3.141592653589793" & vbNewLine &
               "e = " & Math.E.ToString & vbNewLine &
               "deg = " & toDegrees & vbNewLine
        code += functions
        code += vbNewLine + variables + My.Resources.Functions + vbNewLine
        code += TextBox10.Text + vbNewLine
        Dim inc As Double
        If transformX <> "u" Or transformY <> "v" Or transformZ <> "w" Then
            code += "Public Function transformX(u, v, w)" & vbNewLine &
                "On Error Resume Next" & vbNewLine &
                "transformX = " & transformX_Final & vbNewLine &
                "End Function" & vbNewLine &
                "Public Function transformY(u, v, w)" & vbNewLine &
                "On Error Resume Next" & vbNewLine &
                "transformY = " & transformY_Final & vbNewLine &
                "End Function" & vbNewLine &
                "Public Function transformZ(u, v, w)" & vbNewLine &
                "On Error Resume Next" & vbNewLine &
                "transformZ = " & transformZ_Final & vbNewLine &
                "End Function" & vbNewLine
        End If
        If type = 1 Then
            code += "Public Function result1(x)" & vbNewLine &
                    "On Error Resume Next" & vbNewLine &
                    "result1 = " & expr & vbNewLine &
                    "If result1 = ""error"" then result1 = 0" & vbNewLine &
                    "End Function" & vbNewLine &
                    "Public Function result2(t)" & vbNewLine &
                    "result2 = t" & vbNewLine &
                    "End Function" & vbNewLine
            inc = My.Settings.thickness * (xmax - xmin) / PictureBox1.Width
        ElseIf type = 2 Then
            code += "Public Function result1(t)" & vbNewLine &
                    "result1 = t" & vbNewLine &
                    "End Function" & vbNewLine &
                    "Public Function result2(y)" & vbNewLine &
                    "On Error Resume Next" & vbNewLine &
                    "result2 = " & expr & vbNewLine &
                    "If result2 = ""error"" then result2 = 0" & vbNewLine &
                    "End Function" & vbNewLine
            inc = My.Settings.thickness * (ymax - ymin) / PictureBox1.Height
        ElseIf type = 3 Then
            code += "Public Function result1(THET)" & vbNewLine &
                    "On Error Resume Next" & vbNewLine &
                    "result1 = " & expr & vbNewLine &
                    "If result1 = ""error"" then result1 = 0" & vbNewLine &
                    "End Function" & vbNewLine &
                    "Public Function result2(t)" & vbNewLine &
                    "result2 = t" & vbNewLine &
                    "End Function" & vbNewLine
            inc = My.Settings.thickness * (zMax - zMin) / PictureBox1.Width
        ElseIf type = 4 Then
            code += "Public Function result1(t)" & vbNewLine &
                    "result1 = t" & vbNewLine &
                    "End Function" & vbNewLine &
                    "Public Function result2(r)" & vbNewLine &
                    "On Error Resume Next" & vbNewLine &
                    "result2 = " & expr & vbNewLine &
                    "If result2 = ""error"" then result2 = 0" & vbNewLine &
                    "End Function" & vbNewLine
            inc = My.Settings.thickness * (zMax - zMin) / PictureBox1.Width
        ElseIf type = 7 Or type = 9 Then
            code += "Public Function result1(x,y)" & vbNewLine &
                    "On Error Resume Next" & vbNewLine &
                    "result1 = " & expr & vbNewLine &
                    "If result1 = ""error"" then result1 = 0" & vbNewLine &
                    "End Function" & vbNewLine &
                    "Public Function result2(t)" & vbNewLine &
                    "result2 = t" & vbNewLine &
                    "End Function" & vbNewLine
            inc = My.Settings.thickness * (zMax - zMin) / PictureBox1.Width
        ElseIf type = 8 Then
            code += "Public Function result1(THET, r)" & vbNewLine &
                    "On Error Resume Next" & vbNewLine &
                    "result1 = " & expr & vbNewLine &
                    "If result1 = ""error"" then result1 = 0" & vbNewLine &
                    "End Function" & vbNewLine &
                    "Public Function result2(t)" & vbNewLine &
                    "result2 = t" & vbNewLine &
                    "End Function" & vbNewLine
            inc = My.Settings.thickness * (zMax - zMin) / PictureBox1.Width
        ElseIf type = 5 Or type = 6 Then
            code += "Public Function result1(t)" & vbNewLine &
                    "On Error Resume Next" & vbNewLine &
                    "result1 = " & expr2 & vbNewLine &
                    "If result1 = ""error"" then result1 = 0" & vbNewLine &
                    "End Function" & vbNewLine &
                    "Public Function result2(t)" & vbNewLine &
                    "On Error Resume Next" & vbNewLine &
                    "result2 = " & expr & vbNewLine &
                    "If result2 = ""error"" then result2 = 0" & vbNewLine &
                    "End Function" & vbNewLine
            inc = My.Settings.thickness * (zMax - zMin) / PictureBox1.Width
        ElseIf type = 10 Then
            code += "Public Function result1(x,y)" & vbNewLine &
                    "On Error Resume Next" & vbNewLine &
                    "result1 = " & expr & vbNewLine &
                    "If result1 = ""error"" then result1 = 0" & vbNewLine &
                    "End Function" & vbNewLine
            inc = xtick / res3D
            last3DGraph = originalExpr
            shouldTrace = False
        ElseIf type = 11 Then
            code += "Public Function result1(x,y)" & vbNewLine &
                    "On Error Resume Next" & vbNewLine &
                    "result1 = " & expr & vbNewLine &
                    "If result1 = ""error"" then result1 = 0" & vbNewLine &
                    "End Function" & vbNewLine
            inc = (zMax - zMin)
        ElseIf type = 12 Then
            Dim splitString = expr.Split(";")
            code += "Public Function result1(t)" & vbNewLine &
                    "On Error Resume Next" & vbNewLine &
                    "result1 = " & splitString(0) & vbNewLine &
                    "If result1 = ""error"" then result1 = 0" & vbNewLine &
                    "End Function" & vbNewLine &
                    "Public Function result2(t)" & vbNewLine &
                    "On Error Resume Next" & vbNewLine &
                    "result2 = " & splitString(1) & vbNewLine &
                    "If result2 = ""error"" then result2 = 0" & vbNewLine &
                    "End Function" & vbNewLine &
                    "Public Function result3(t)" & vbNewLine &
                    "On Error Resume Next" & vbNewLine &
                    "result3 = " & splitString(2) & vbNewLine &
                    "If result3 = ""error"" then result3 = 0" & vbNewLine &
                    "End Function" & vbNewLine
            inc = My.Settings.thickness * (zMax - zMin) / PictureBox1.Width
        End If
        code = Replace(code, "π", "PI")
        code = Replace(code, "θ", "THET")
        sc.AddCode(code)
        If expr.Contains("rnd") Then
            sc.ExecuteStatement("Randomize(" + rndSeed.ToString + ")")
            rndSeed = Rnd()
        End If
        If Me.Tag = "graph" Then
            If type = 1 Then
                If findType = 1 Then 'find zero
                    Const cutoff = 0.000000001
                    Dim epsilon
                    Dim iterations As Integer = 0
                    If Math.Abs(deriveF(sc, startingGuess)) < 0.00001 Then
                        startingGuess += 3 * (xmax - xmin) / PictureBox1.Height
                    End If
                    Do
                        iterations += 1
                        epsilon = -1 * sc.Run("result1", startingGuess) / deriveF(sc, startingGuess)
                        startingGuess += epsilon
                    Loop While Math.Abs(epsilon) > cutoff And iterations < 1000
                    findType = -1
                    'traceNumber = startingGuess
                    traceNumber = Math.Round(startingGuess, 7)
                    If traceNumber.ToString = "NaN" Or iterations >= 1000 Then
                        shouldTrace = False
                        MsgBox("Photon could not find a zero near the value specified.")
                    Else
                        shouldTrace = True
                    End If
                ElseIf findType = 2 Then 'find min / max
                    Const cutoff = 0.000000001
                    Dim epsilon
                    Dim iterations As Integer = 0
                    Do
                        iterations += 1
                        epsilon = -1 * deriveF(sc, startingGuess) / derive2F(sc, startingGuess)
                        startingGuess += epsilon
                    Loop While Math.Abs(epsilon) > cutoff And iterations < 1000
                    findType = -1
                    traceNumber = Math.Round(startingGuess, 7)
                    If traceNumber.ToString = "NaN" Or iterations >= 1000 Then
                        shouldTrace = False
                        MsgBox("Photon could not find a minimum or maximum near the value specified.")
                    Else
                        shouldTrace = True
                    End If
                ElseIf findType = 3 Then 'find inflection point
                    Const cutoff = 0.00000001
                    Dim epsilon
                    Dim iterations As Integer = 0
                    Do
                        iterations += 1
                        epsilon = -1 * derive2F(sc, startingGuess) / derive3F(sc, startingGuess)
                        startingGuess += epsilon
                    Loop While Math.Abs(epsilon) > cutoff And iterations < 1000
                    findType = -1
                    traceNumber = Math.Round(startingGuess, 7)
                    If traceNumber.ToString = "NaN" Or iterations >= 1000 Then
                        shouldTrace = False
                        MsgBox("Photon could not find an inflection point near the value specified.")
                    Else
                        shouldTrace = True
                    End If
                ElseIf findType = 4 Then 'find intersection (not currently used)

                End If
            Else
                findType = -1
            End If
            Dim myPath As New List(Of PointF)
            Dim my3DPath As New List(Of Point3D)
            Dim t As Double = zMin
            Dim x1 As Double, y1 As Double = (ymax + ymin) / 2 + (ymax - ymin) / 2.5, z1 As Double
            Dim slopeStage As Integer = 0, previousSlope = 0, slopeMultiplier As Integer = 1
            Dim loopBackCount = 0 'used for implicit plotting
            If type = 8 Then y1 = 0
            Dim h As Double = 0.0001
            Dim cutoffImplicit As Double = (xmax - xmin) / PictureBox1.Width
            Dim doNotAdd As Boolean = True 'used for implicit plotting
            Dim checkedForRedraw As Boolean = False
            If type <> 7 And type <> 8 And type <> 9 Then doNotAdd = False
            If secondDerivative Then h = 0.001
            inc = Math.Max(inc, (zMax - zMin) / 1000)
            If type = 9 Then 'round 2 for implicit plotting
                inc = -inc
                t = zMax + inc
                y1 = (ymax + ymin) / 2 - (ymax - ymin) / 2.5
                type = 7
            ElseIf type = 10 Then '3D graph setup for t
                t = Math.Round(t / xtick) * xtick
                zMax = Math.Round(zMax / xtick) * xtick + inc / 2
            End If
            While t <= zMax 'we do two loops to ensure that t = zmax when the loop is done; outer loop really just goes once
                While t <= zMax 'the loop handles function values as well as first and second derivatives for all graphing modes
                    If shouldTrace Then
                        t = traceNumber 'first check for trace, then start the loop over at t = zMin
                        shouldDrawTrace = True
                    End If
                    Dim implicitR As Double = 0
                    If type = 1 Then 'y = f(x)
                        x1 = t
                        If firstDerivative Then
                            y1 = (sc.Run("result1", t + h) - sc.Run("result1", t - h)) / (2 * h)
                        ElseIf secondDerivative Then
                            y1 = (sc.Run("result1", t + h) - 2 * sc.Run("result1", t) + sc.Run("result1", t - h)) / (h * h)
                        Else
                            y1 = sc.Run("result1", t)
                        End If
                    ElseIf type = 2 Then 'x = f(y)
                        x1 = sc.Run("result2", t)
                        If firstDerivative Then
                            y1 = (2 * h) / (sc.Run("result2", t + h) - sc.Run("result2", t - h))
                        ElseIf secondDerivative Then
                            Dim x0 = sc.Run("result2", t - h)
                            Dim x2 = sc.Run("result2", t + h)
                            Dim x3 = sc.Run("result2", t + 2 * h)
                            Dim dx1 = (2 * h) / (x2 - x0)
                            Dim dx2 = (2 * h) / (x3 - x1)
                            y1 = (dx2 - dx1) / (x2 - x1)
                        Else
                            y1 = t
                        End If
                    ElseIf type = 3 Then 'r = f(th)
                        Dim r1
                        If firstDerivative Then
                            r1 = (sc.Run("result1", t + h) - sc.Run("result1", t - h)) / (2 * h)
                        ElseIf secondDerivative Then
                            r1 = (sc.Run("result1", t + h) - 2 * sc.Run("result1", t) + sc.Run("result1", t - h)) / (h * h)
                        Else
                            r1 = sc.Run("result1", t)
                        End If
                        x1 = r1 * Math.Cos(t * numberToDegrees)
                        y1 = r1 * Math.Sin(t * numberToDegrees)
                    ElseIf type = 4 Then 'th = f(r)
                        Dim th1 = sc.Run("result2", t)
                        Dim r1
                        If firstDerivative Then
                            r1 = (2 * h) / (sc.Run("result2", t + h) - sc.Run("result2", t - h))
                        ElseIf secondDerivative Then
                            Dim th0 = sc.Run("result2", t - h)
                            Dim th2 = sc.Run("result2", t + h)
                            Dim th3 = sc.Run("result2", t + 2 * h)
                            Dim dth1 = (2 * h) / (th2 - th0)
                            Dim dth2 = (2 * h) / (th3 - th1)
                            r1 = (dth2 - dth1) / (th2 - th1)
                        Else
                            r1 = t
                        End If
                        x1 = r1 * Math.Cos(th1 * numberToDegrees)
                        y1 = r1 * Math.Sin(th1 * numberToDegrees)
                    ElseIf type = 5 Then 'x = f(t); y = f(t)
                        x1 = sc.Run("result2", t)
                        y1 = sc.Run("result1", t)
                        If firstDerivative Then
                            Dim x2 = sc.Run("result2", t + h)
                            Dim y2 = sc.Run("result1", t + h)
                            y1 = (y2 - y1) / (x2 - x1)
                        ElseIf secondDerivative Then
                            Dim x0 = sc.Run("result2", t - h)
                            Dim y0 = sc.Run("result1", t - h)
                            Dim x2 = sc.Run("result2", t + h)
                            Dim y2 = sc.Run("result1", t + h)
                            Dim dx1 = (y1 - y0) / (x1 - x0)
                            Dim dx2 = (y2 - y1) / (x2 - x1)
                            y1 = (dx2 - dx1) / (x1 - x0)
                        End If
                    ElseIf type = 6 Then 'r = f(t), th = f(t)
                        Dim th1 = sc.Run("result2", t)
                        Dim r1 = sc.Run("result1", t)
                        If firstDerivative Then
                            Dim th2 = sc.Run("result2", t + h)
                            Dim r2 = sc.Run("result1", t + h)
                            r1 = (r2 - r1) / (th2 - th1)
                        ElseIf secondDerivative Then
                            Dim th0 = sc.Run("result2", t - h)
                            Dim r0 = sc.Run("result1", t - h)
                            Dim th2 = sc.Run("result2", t + h)
                            Dim r2 = sc.Run("result1", t + h)
                            Dim dth1 = (r1 - r0) / (th1 - th0)
                            Dim dth2 = (r2 - r1) / (th2 - th1)
                            r1 = (dth2 - dth1) / (th1 - th0)
                        End If
                        x1 = r1 * Math.Cos(th1 * numberToDegrees)
                        y1 = r1 * Math.Sin(th1 * numberToDegrees)
                    ElseIf type = 7 Or type = 8 Then 'f(x,y) = g(x,y) (implicit) or f(r,th) = g(r,th) (polar implicit)
                        If (t < zMin And type = 7) Or (t < -zMax And type = 8) Then
                            t = zMax + 1
                            Exit While
                        End If
                        x1 = t
                        Dim epsilon
                        Dim iterations As Integer = 0
                        If Math.Abs(derive2VarF(sc, x1, y1)) < 0.00001 Then
                            y1 += 3 * (xmax - xmin) / PictureBox1.Height
                        End If
                        Do
                            iterations += 1
                            epsilon = -1 * sc.Run("result1", x1, y1) / derive2VarF(sc, x1, y1)
                            y1 += epsilon
                        Loop While Math.Abs(epsilon) > cutoffImplicit And iterations < 8
                        If y1.ToString = "NaN" Or iterations >= 8 Then
                            If doNotAdd = False And loopBackCount < 4 Then 'loop back around
                                If y1 > ymin And y1 < ymax Then
                                    y1 = (ymax - myPath(myPath.Count - 1).Y * (ymax - ymin) / PictureBox1.Height) * 6 - (ymax - myPath(myPath.Count - 2).Y * (ymax - ymin) / PictureBox1.Height) * 5
                                    inc = -inc
                                    loopBackCount += 1
                                Else
                                    y1 = (ymax + ymin) / 2
                                    If type = 8 Then y1 = 0
                                End If
                            Else
                                y1 = (ymax + ymin) / 2 + ((ymax - ymin) / 2.5) * Math.Sign(inc)
                            End If
                            doNotAdd = True
                        Else
                            doNotAdd = False
                        End If
                        If type = 8 Then
                            implicitR = y1
                            Dim implicitTheta = x1
                            x1 = implicitR * Math.Cos(implicitTheta * numberToDegrees)
                            y1 = implicitR * Math.Sin(implicitTheta * numberToDegrees)
                        End If
                    ElseIf type = 10 Then                                                   '3D graphing!!!
                        'Dim u As Double = Math.Round(ymin / ytick) * ytick
                        Dim u = Math.Round(zMin / xtick) * xtick
                        While u < zMax
                            If transformX <> "u" Or transformY <> "v" Or transformZ <> "w" Then 'UV transform
                                Dim tLess = t - xtick / halfSqWidth
                                Dim tMore = t + xtick / halfSqWidth
                                Dim uLess = u - ytick / halfSqWidth
                                Dim uMore = u + ytick / halfSqWidth
                                Dim z_1 = sc.Run("result1", tLess, uLess)
                                Dim z_2 = sc.Run("result1", tLess, uMore)
                                Dim z_3 = sc.Run("result1", tMore, uMore)
                                Dim z_4 = sc.Run("result1", tMore, uLess)
                                Dim p1X = sc.Run("transformX", tLess, uLess, z_1)
                                Dim p1Y = sc.Run("transformY", tLess, uLess, z_1)
                                Dim p1Z = sc.Run("transformZ", tLess, uLess, z_1)
                                Dim p2X = sc.Run("transformX", tLess, uMore, z_2)
                                Dim p2Y = sc.Run("transformY", tLess, uMore, z_2)
                                Dim p2Z = sc.Run("transformZ", tLess, uMore, z_2)
                                Dim p3X = sc.Run("transformX", tMore, uMore, z_3)
                                Dim p3Y = sc.Run("transformY", tMore, uMore, z_3)
                                Dim p3Z = sc.Run("transformZ", tMore, uMore, z_3)
                                Dim p4X = sc.Run("transformX", tMore, uLess, z_4)
                                Dim p4Y = sc.Run("transformY", tMore, uLess, z_4)
                                Dim p4Z = sc.Run("transformZ", tMore, uLess, z_4)
                                Dim p1 As New Point3D(p1X, p1Y, p1Z * zFactor)
                                Dim p2 As New Point3D(p2X, p2Y, p2Z * zFactor)
                                Dim p3 As New Point3D(p3X, p3Y, p3Z * zFactor)
                                Dim p4 As New Point3D(p4X, p4Y, p4Z * zFactor)
                                all3DFaces.Add(New Face3D(p1, p2, p3, p4))
                            Else
                                Dim p1 As New Point3D(t - xtick / halfSqWidth, u - ytick / halfSqWidth, sc.Run("result1", t - xtick / halfSqWidth, u - ytick / halfSqWidth) * zFactor)
                                Dim p2 As New Point3D(t - xtick / halfSqWidth, u + ytick / halfSqWidth, sc.Run("result1", t - xtick / halfSqWidth, u + ytick / halfSqWidth) * zFactor)
                                Dim p3 As New Point3D(t + xtick / halfSqWidth, u + ytick / halfSqWidth, sc.Run("result1", t + xtick / halfSqWidth, u + ytick / halfSqWidth) * zFactor)
                                Dim p4 As New Point3D(t + xtick / halfSqWidth, u - ytick / halfSqWidth, sc.Run("result1", t + xtick / halfSqWidth, u - ytick / halfSqWidth) * zFactor)
                                all3DFaces.Add(New Face3D(p1, p2, p3, p4))
                            End If
                            u += ytick / res3D
                        End While
                    ElseIf type = 11 Then                     'slope field solver
                        Dim multiplier = Math.Min(xtick, ytick) / 20
                        If slopeStage = 0 Then
                            x1 = slopeFieldPoint.X
                            y1 = slopeFieldPoint.Y
                            Dim slope = sc.Run("result1", x1, y1)
                            If slope > 1 And previousSlope < -1 Or slope < -1 And previousSlope > 1 Then
                                slopeMultiplier *= -1
                            End If
                            Dim dx = Math.Cos(Math.Atan(slope)) * slopeMultiplier
                            Dim dy = Math.Sin(Math.Atan(slope)) * slopeMultiplier
                            Dim magnitude = Math.Sqrt(dx ^ 2 + dy ^ 2)
                            dx = dx * multiplier / magnitude
                            dy = dy * multiplier / magnitude
                            x1 += dx
                            y1 += dy
                            slopeFieldPoint.X = x1
                            slopeFieldPoint.Y = y1
                            If x1 < xmax And x1 > xmin And y1 < ymax And y1 > ymin Then
                                t = zMin - inc
                            Else
                                slopeStage = 1
                            End If
                            previousSlope = slope
                        Else
                            If slopeStage = 1 Then
                                slopeFieldPoint.X = myPath(0).X / PictureBox1.Width * (xmax - xmin) + xmin
                                slopeFieldPoint.Y = (PictureBox1.Height - myPath(0).Y) / PictureBox1.Height * (ymax - ymin) + ymin
                                slopeMultiplier = 1
                                previousSlope = 0
                                slopeStage = 2
                            End If
                            x1 = slopeFieldPoint.X
                            y1 = slopeFieldPoint.Y
                            Dim slope = sc.Run("result1", x1, y1)
                            If slope > 1 And previousSlope < -1 Or slope < -1 And previousSlope > 1 Then
                                slopeMultiplier *= -1
                            End If
                            Dim dx = Math.Cos(Math.Atan(slope)) * slopeMultiplier
                            Dim dy = Math.Sin(Math.Atan(slope)) * slopeMultiplier
                            Dim magnitude = Math.Sqrt(dx ^ 2 + dy ^ 2)
                            dx = dx * multiplier / magnitude
                            dy = dy * multiplier / magnitude
                            x1 -= dx
                            y1 -= dy
                            slopeFieldPoint.X = x1
                            slopeFieldPoint.Y = y1
                            If x1 < xmax And x1 > xmin And y1 < ymax And y1 > ymin Then
                                t = zMax - inc * 1.5
                            End If
                            previousSlope = slope
                        End If
                        If myPath.Count > 10000 Then t = zMax + inc * 2 'make sure the solution doesn't loop forever
                    ElseIf type = 12 Then                 '3D parametric equations
                        x1 = sc.Run("result1", t)
                        y1 = sc.Run("result2", t)
                        z1 = sc.Run("result3", t)
                    End If
                    If type <> 10 And transformX <> "u" Or transformY <> "v" Or transformZ <> "w" Then 'UV transform
                        Dim finalX = sc.Run("transformX", x1, y1, z1)
                        Dim finalY = sc.Run("transformY", x1, y1, z1)
                        Dim finalZ = sc.Run("transformZ", x1, y1, z1)
                        x1 = finalX
                        y1 = finalY
                        z1 = finalZ
                    End If
                    If type <> 10 And type <> 12 Then 'we don't need this for 3D graphing
                        If ((LogXAxisToolStripMenuItem.Checked And x1 <= 0) Or (LogYAxisToolStripMenuItem.Checked And y1 <= 0)) = False Then
                            If LogXAxisToolStripMenuItem.Checked Then x1 = Math.Log10(x1)
                            If LogYAxisToolStripMenuItem.Checked Then y1 = Math.Log10(y1)
                            If shouldTrace Then
                                t = zMin
                                traceCoord = New PointF(x1, y1)
                                tracePoint = New PointF(Math.Min(Math.Max((x1 - xmin) / (xmax - xmin) * PictureBox1.Width, -100), PictureBox1.Width + 100), Math.Min(Math.Max((ymax - y1) / (ymax - ymin) * PictureBox1.Height, -100), PictureBox1.Height + 100))
                            ElseIf doNotAdd = False Then
                                If type = 11 Then
                                    If slopeStage <> 2 Then
                                        myPath.Add(New PointF(Math.Min(Math.Max((x1 - xmin) / (xmax - xmin) * PictureBox1.Width, -100), PictureBox1.Width + 100), Math.Min(Math.Max((ymax - y1) / (ymax - ymin) * PictureBox1.Height, -100), PictureBox1.Height + 100)))
                                    Else
                                        myPath.Insert(0, New PointF(Math.Min(Math.Max((x1 - xmin) / (xmax - xmin) * PictureBox1.Width, -100), PictureBox1.Width + 100), Math.Min(Math.Max((ymax - y1) / (ymax - ymin) * PictureBox1.Height, -100), PictureBox1.Height + 100)))
                                    End If
                                Else
                                    If myPath.Count > 1 Then 'check for asymptotes
                                        If myPath(myPath.Count - 1).Y < 0 And y1 < ymin Then
                                            myPath.Add(New PointF(-100, -100))
                                            myPath.Add(New PointF(-100, PictureBox1.Height + 100))
                                        ElseIf myPath(myPath.Count - 1).Y > PictureBox1.Height And y1 > ymax Then
                                            myPath.Add(New PointF(-100, PictureBox1.Height + 100))
                                            myPath.Add(New PointF(-100, -100))
                                        End If
                                    End If
                                    myPath.Add(New PointF(Math.Min(Math.Max((x1 - xmin) / (xmax - xmin) * PictureBox1.Width, -100), PictureBox1.Width + 100), Math.Min(Math.Max((ymax - y1) / (ymax - ymin) * PictureBox1.Height, -100), PictureBox1.Height + 100)))
                                End If
                            End If
                        Else
                            If shouldTrace Then
                                tracePoint = New PointF(-1, -1)
                            End If
                        End If
                        shouldTrace = False
errorStop:
                        If type = 8 Then y1 = implicitR
                    ElseIf type = 12 Then
                        If shouldTrace Then
                            t = zMin
                            traceCoord3DOriginal = New Point3D(x1, y1, z1)
                            traceCoord3D = traceCoord3DOriginal.RotateX(-90).RotateY(135).RotateX(20)
                            shouldTrace = False
                        Else
                            my3DPath.Add(New Point3D(x1, y1, z1).RotateX(-90).RotateY(135).RotateX(20))
                        End If
                    End If
                    t += inc
                End While
                If (t - inc / 1.00001 < zMax) And type <> 10 Then t = zMax 'set t equal to zmax and run the loop one more time through
            End While
            If shouldIntegrate Then 'Implements Simpson's Rule for numerical integration
                If sc.Run("result2", zMin) > sc.Run("result2", zMax) Then ' fixes getting a negative answer when it should be positive
                    Dim tempZ = zMin
                    zMin = zMax
                    zMax = tempZ
                End If
                Dim n = 1000
                Dim h1 As Double = (zMax - zMin) / (n)
                Dim s As Double = 0
                If type = 1 Or type = 2 Or type = 5 Then ' Regular Integration I = a to b S f(x)dx
                    If transformX <> "u" Or transformY <> "v" Then 'UV transform
                        Dim y_1 = sc.Run("transformY", zMin, sc.Run("result1", zMin), 0)
                        Dim y_2 = sc.Run("transformY", zMax, sc.Run("result1", zMax), 0)
                        s = y_1 * deriveXTransformed(sc, zMin, sc.Run("result1", zMin)) + y_2 * deriveXTransformed(sc, zMax, sc.Run("result1", zMax))
                        For i = 1 To n Step 2
                            Dim x_1 = zMin + i * h1
                            y_1 = sc.Run("result1", x_1)
                            Dim y_1_Final = sc.Run("transformY", x_1, y_1, 0)
                            s += 4 * y_1_Final * deriveXTransformed(sc, x_1, y_1)
                        Next i
                        For i = 2 To n - 1 Step 2
                            Dim x_1 = zMin + i * h1
                            y_1 = sc.Run("result1", x_1)
                            Dim y_1_Final = sc.Run("transformY", x_1, y_1, 0)
                            s += 2 * y_1_Final * deriveXTransformed(sc, x_1, y_1)
                        Next i
                    Else
                        s = sc.Run("result1", zMin) * deriveX(sc, zMin) + sc.Run("result1", zMax) * deriveX(sc, zMax)
                        For i = 1 To n Step 2
                            s += 4 * sc.Run("result1", zMin + i * h1) * deriveX(sc, zMin + i * h1)
                        Next i
                        For i = 2 To n - 1 Step 2
                            s += 2 * sc.Run("result1", zMin + i * h1) * deriveX(sc, zMin + i * h1)
                        Next i
                    End If
                    If My.Settings.fraction Then
                        integrals.Add(Math.Round(s * h1 / 3.0, 15))
                    Else
                        integrals.Add(Math.Round(s * h1 / 3.0, 8))
                    End If
                Else                                     ' Polar Integration I = 1/2 * (a to b S r^2 d theta)
                    If transformX <> "u" Or transformY <> "v" Then 'UV transform
                        MsgBox("Polar integration is not allowed on a UV-transformed coordinate system because polar coordinates already transform the coordinate system.")
                    Else
                        s = sc.Run("result1", zMin) ^ 2 * deriveX(sc, zMin) + sc.Run("result1", zMax) ^ 2 * deriveX(sc, zMax)
                        For i = 1 To n Step 2
                            s += 4 * sc.Run("result1", zMin + i * h1) ^ 2 * deriveX(sc, zMin + i * h1)
                        Next i
                        For i = 2 To n - 1 Step 2
                            s += 2 * sc.Run("result1", zMin + i * h1) ^ 2 * deriveX(sc, zMin + i * h1)
                        Next i
                    End If
                    integrals.Add(Math.Round(s * h1 / 6.0, 8))
                End If
                integralFlag.Add(True)
            Else
                integralFlag.Add(False)
            End If
            If myPath.Count > 0 Then
                If firstDerivative Then
                    firstDerivativePoints.Add(myPath.ToArray())
                    firstDerivative = False
                ElseIf secondDerivative Then
                    secondDerivativePoints.Add(myPath.ToArray())
                    secondDerivative = False
                Else
                    If shouldIntegrate Then
                        If type = 1 Or type = 2 Or type = 5 Then 'draw point at (X2, 0) and (X1, 0) for regular graphs, (X2, -inf) and (X1, -inf) for log y graphs
                            If LogYAxisToolStripMenuItem.Checked = False Then
                                myPath.Add(New PointF(myPath.Item(myPath.Count - 1).X, Math.Min(Math.Max((ymax) / (ymax - ymin) * PictureBox1.Height, -100), PictureBox1.Height + 100)))
                                myPath.Add(New PointF(myPath.Item(0).X, Math.Min(Math.Max((ymax) / (ymax - ymin) * PictureBox1.Height, -100), PictureBox1.Height + 100)))
                            Else
                                myPath.Add(New PointF(myPath.Item(myPath.Count - 1).X, PictureBox1.Height))
                                myPath.Add(New PointF(myPath.Item(0).X, PictureBox1.Height))
                            End If
                        Else 'draw point at (0,0) for polar graphs
                            If LogXAxisToolStripMenuItem.Checked = False And LogYAxisToolStripMenuItem.Checked = False Then '(0,0) will work for no logs
                                myPath.Add(New PointF(Math.Min(Math.Max((-xmin) / (xmax - xmin) * PictureBox1.Width, -100), PictureBox1.Width + 100), Math.Min(Math.Max((ymax) / (ymax - ymin) * PictureBox1.Height, -100), PictureBox1.Height + 100)))
                            ElseIf LogXAxisToolStripMenuItem.Checked = True And LogYAxisToolStripMenuItem.Checked = False Then 'log(0) = -infinity x
                                myPath.Add(New PointF(0, myPath.Item(myPath.Count - 1).Y))
                                myPath.Add(New PointF(0, myPath.Item(0).Y))
                            ElseIf LogXAxisToolStripMenuItem.Checked = False And LogYAxisToolStripMenuItem.Checked = True Then 'log(0) = -infinity y
                                myPath.Add(New PointF(myPath.Item(myPath.Count - 1).X, PictureBox1.Height))
                                myPath.Add(New PointF(myPath.Item(0).X, PictureBox1.Height))
                            Else                                                                                               'log(0, 0) = -infinity both x and y
                                myPath.Add(New PointF(myPath.Item(myPath.Count - 1).X - PictureBox1.Width, myPath.Item(myPath.Count - 1).Y + PictureBox1.Height))
                                myPath.Add(New PointF(myPath.Item(0).X - PictureBox1.Width, myPath.Item(0).Y + PictureBox1.Height))
                            End If
                        End If
                    End If
                    allThePoints.Add(myPath.ToArray())
                End If
            ElseIf my3DPath.Count > 0 Then
                all3DPoints.Add(my3DPath.ToArray())
            End If
            shouldIntegrate = False
        ElseIf Me.Tag = "table" And type = 1 Then
            DataFunction.Columns.Insert(1, New DataGridViewColumn(DataFunction.Rows(0).Cells(0)))
            DataFunction.Columns(1).HeaderText = "y = " + originalExpr
            For i = 0 To DataFunction.Rows.Count - 1
                DataFunction.Rows(i).Cells(1).Value = sc.Run("result1", Math.Round(Val(txtStart.Text) + Val(txtDelta.Text) * i, 12))
            Next i
        End If
    End Sub

    'Takes the x(t) derivative of a function; ironically, this is only used for taking integrals
    Function deriveX(f As MSScriptControl.ScriptControl, x As Double)
        Dim v1 As Double, v2 As Double, v3 As Double, v4 As Double
        Const h As Double = 0.000456
        v1 = f.Run("result2", x - 2 * h)
        v2 = f.Run("result2", x - h)
        v3 = f.Run("result2", x + h)
        v4 = f.Run("result2", x + 2 * h)
        deriveX = (v1 - 8 * v2 + 8 * v3 - v4) / (12 * h)
    End Function

    Function deriveXTransformed(f As MSScriptControl.ScriptControl, x As Double, y As Double)
        Dim v1 As Double, v2 As Double, v3 As Double, v4 As Double
        Const h As Double = 0.000456
        v1 = f.Run("transformX", f.Run("result2", x - 2 * h), y, 0)
        v2 = f.Run("transformX", f.Run("result2", x - h), y, 0)
        v3 = f.Run("transformX", f.Run("result2", x + h), y, 0)
        v4 = f.Run("transformX", f.Run("result2", x + 2 * h), y, 0)
        deriveXTransformed = Math.Abs((v1 - 8 * v2 + 8 * v3 - v4) / (12 * h))
    End Function

    'Evaluates dy of a function defined implicitly; x is replaced with a number so we're just trying to solve for y
    Function derive2VarF(f As MSScriptControl.ScriptControl, x As Double, y As Double)
        Dim v1 As Double, v2 As Double, v3 As Double, v4 As Double
        Const h As Double = 0.000456
        v1 = f.Run("result1", x, y - 2 * h)
        v2 = f.Run("result1", x, y - h)
        v3 = f.Run("result1", x, y + h)
        v4 = f.Run("result1", x, y + 2 * h)
        derive2VarF = (v1 - 8 * v2 + 8 * v3 - v4) / (12 * h)
    End Function

    'Evaluates dy/dx of a function; this is used for finding roots, min / max, inflection points, and points of intersection
    Function deriveF(f As MSScriptControl.ScriptControl, x As Double)
        Dim v1 As Double, v2 As Double, v3 As Double, v4 As Double
        Const h As Double = 0.000456
        v1 = f.Run("result2", x - 2 * h)
        v2 = f.Run("result2", x - h)
        v3 = f.Run("result2", x + h)
        v4 = f.Run("result2", x + 2 * h)
        Dim deriveX = (v1 - 8 * v2 + 8 * v3 - v4) / (12 * h)
        v1 = f.Run("result1", x - 2 * h)
        v2 = f.Run("result1", x - h)
        v3 = f.Run("result1", x + h)
        v4 = f.Run("result1", x + 2 * h)
        Dim deriveY = (v1 - 8 * v2 + 8 * v3 - v4) / (12 * h)
        deriveF = deriveY / deriveX
    End Function

    ' Evaluates dy^2/d^2x of a function; used for finding min / max and inflection points
    Function derive2F(f As MSScriptControl.ScriptControl, x As Double)
        Dim v1 As Double, v2 As Double, v3 As Double, v4 As Double
        Const h As Double = 0.000456
        v1 = deriveF(f, x - 2 * h)
        v2 = deriveF(f, x - h)
        v3 = deriveF(f, x + h)
        v4 = deriveF(f, x + 2 * h)
        Dim deriveY = (v1 - 8 * v2 + 8 * v3 - v4) / (12 * h)
        v1 = f.Run("result2", x - 2 * h)
        v2 = f.Run("result2", x - h)
        v3 = f.Run("result2", x + h)
        v4 = f.Run("result2", x + 2 * h)
        Dim deriveX = (v1 - 8 * v2 + 8 * v3 - v4) / (12 * h)
        derive2F = deriveY / deriveX
    End Function

    ' Evaluates dy^3 / d^3x of a function; used for finding inflection points
    Function derive3F(f As MSScriptControl.ScriptControl, x As Double)
        Dim v1 As Double, v2 As Double, v3 As Double, v4 As Double
        Const h As Double = 0.000456
        v1 = derive2F(f, x - 2 * h)
        v2 = derive2F(f, x - h)
        v3 = derive2F(f, x + h)
        v4 = derive2F(f, x + 2 * h)
        Dim deriveY = (v1 - 8 * v2 + 8 * v3 - v4) / (12 * h)
        v1 = f.Run("result2", x - 2 * h)
        v2 = f.Run("result2", x - h)
        v3 = f.Run("result2", x + h)
        v4 = f.Run("result2", x + 2 * h)
        Dim deriveX = (v1 - 8 * v2 + 8 * v3 - v4) / (12 * h)
        derive3F = deriveY / deriveX
    End Function

    ' Evaluates text as a number - example: "3 + 4" becomes 7..."fractionMode" indicates fractions, and "isCalc" indicates whether we should store the answer in the "ans" variable
    Function eval(expr As String, fractionMode As Boolean, isCalc As Boolean) As Object
        On Error GoTo err
        Dim code As String = ""
        Dim sc = New MSScriptControl.ScriptControl
        Dim illegalVariable As String = "'"
        expr = evalFix(expr)
        Dim counter = 0
        Dim outString = ""
        For Each letter As String In expr
            If counter <> expr.Length - 1 Then
                If letter = "A" And Char.IsLetter(expr(counter + 1)) = False Then letter = "AA"
            Else
                If letter = "A" Then letter = "AA"
            End If
            outString += letter
            counter += 1
        Next letter
        expr = outString
        Dim toDegrees
        If DegreeModeToolStripMenuItem.Checked = True Then
            toDegrees = "PI / 180"
        Else
            toDegrees = "1"
        End If
        Dim variables As String = ""
        Dim functions As String = ""
        variables += "ans = " + ans.ToString + vbNewLine
        If TrackBar1.Visible = True Then variables += "AA = " + (TrackBar1.Value / TrackBar1.Maximum * (highA - lowA) + lowA).ToString + vbNewLine
        For i = 0 To RichTextBox1.Lines.Length - 1
            If RichTextBox1.Lines(i).Contains(":=") And RichTextBox1.Lines(i).StartsWith(illegalVariable) = False And RichTextBox1.Lines(i).StartsWith("e", vbTextCompare) = False And RichTextBox1.Lines(i).Contains("->") = False Then
                Dim info = Split(RichTextBox1.Lines(i).Replace(" ", ""), ":=")
                Dim myVar = Replace(info(0), "θ", "THET")
                If myVar.Length = 1 Or myVar = "THET" Then
                    variables = variables + myVar + " = " + evalFix(heal(info(1).Replace("A", "AA"))) + vbNewLine
                Else
                    functions = functions + "Public Function " + myVar + vbNewLine + myVar.First + " = " + evalFix(heal(info(1), False)) + vbNewLine + "End Function" + vbNewLine
                End If
            End If
        Next i
        code = "PI = 3.141592653589793" & vbNewLine &
               "e = " & Math.E.ToString & vbNewLine &
               "deg = " & toDegrees & vbNewLine
        code += functions
        code += vbNewLine & variables
        If isFindingSlopeField Then 'mimic multiEval for fast slope calculations
            code += "Public Function result(x, y)" & vbNewLine &
                "On Error Resume Next" & vbNewLine &
                "result = " & expr & vbNewLine &
                "If Err or result = ""error"" Then result = 9999995" & vbNewLine &
                "End Function" & vbNewLine
            code += My.Resources.Functions + vbNewLine
            code += TextBox10.Text + vbNewLine
            sc.Language = "VBScript"
            sc.AllowUI = True
            code = Replace(code, "π", "PI")
            code = Replace(code, "θ", "THET")
            sc.AddCode(code)
            If expr.Contains("rnd") Then
                sc.ExecuteStatement("Randomize(" + rndSeed.ToString + ")")
                rndSeed = Rnd()
            End If
            Dim startingX = Math.Ceiling(xmin / xtick * 2) * xtick / 2
            currentY = Math.Ceiling(ymin / ytick * 2) * ytick / 2
            currentX = startingX
            For i = 0 To sFieldSlopeArray.GetLength(0) - 1
                For j = 0 To sFieldSlopeArray.GetLength(1) - 1
                    sFieldLocationArray(i, j) = New PointF(currentX, currentY)
                    sFieldSlopeArray(i, j) = sc.Run("result", currentX, currentY)
                    currentX += xtick / 2
                Next j
                currentX = startingX
                currentY += ytick / 2
            Next i
            eval = 1
        Else 'the actual use of this function is here
            code += "Public Function result()" & vbNewLine &
                    "result = " & expr & vbNewLine &
                    "End Function" + vbNewLine
            code += My.Resources.Functions + vbNewLine
            code += TextBox10.Text + vbNewLine
            sc.Language = "VBScript"
            sc.AllowUI = True
            code = Replace(code, "π", "PI")
            code = Replace(code, "θ", "THET")
            sc.AddCode(code)
            If expr.Contains("rnd") Then
                sc.ExecuteStatement("Randomize(" + rndSeed.ToString + ")")
                rndSeed = Rnd()
            End If
            Dim calc = sc.Run("result")
            If isCalc And (IsNumeric(calc) Or calc.ToString.Contains("/")) Then ans = calc
            Dim result
            If fractionMode = True And IsNumeric(calc) Then
                Dim strings As New ArrayList
                'simple fraction
                strings.Add(decToFraction(calc))
                'square root
                Dim sString = Split(decToFraction(calc * calc, False), "/")
                Dim numerator, denominator
                numerator = sString(0)
                denominator = sString(1)
                numerator = numerator * denominator
                Dim commonSquare = ExtractSquare(numerator)
                numerator = numerator / (commonSquare * commonSquare)
                Dim coefString = Split(decToFraction(commonSquare / denominator, False), "/")
                commonSquare = coefString(0)
                denominator = coefString(1)
                Dim sqrtString As String
                sqrtString = commonSquare.ToString + "*√(" + (numerator).ToString + ")/" + denominator.ToString
                If sqrtString.EndsWith("/1") Then sqrtString = sqrtString.Substring(0, sqrtString.Length - 2)
                If sqrtString.StartsWith("1*") Then sqrtString = sqrtString.Substring(2, sqrtString.Length - 2)
                If calc < 0 Then sqrtString = "-" + sqrtString
                strings.Add(sqrtString)
                'pi
                Dim piPartString As String = decToFraction(calc / Math.PI)
                Dim piString
                If piPartString.Contains("/") = False Then
                    piString = piPartString + "π"
                Else
                    piString = Split(piPartString, "/")(0) + "π/" + Split(piPartString, "/")(1)
                End If
                piString = "`" + piString
                piString = Replace(piString, "`1π", "π")
                piString = Replace(piString, "`-1π", "-π")
                piString = Replace(piString, "`", "")
                strings.Add(piString)
                'e
                Dim ePartString As String = decToFraction(calc / Math.E)
                Dim eString
                If ePartString.Contains("/") = False Then
                    eString = ePartString + "e"
                Else
                    eString = Split(ePartString, "/")(0) + "e/" + Split(ePartString, "/")(1)
                End If
                eString = "`" + eString
                eString = Replace(eString, "`1e", "e")
                eString = Replace(eString, "`-1e", "-e")
                eString = Replace(eString, "`", "")
                strings.Add(eString)
                'finished with math simplification
                result = strings.Item(0)
                For i = 1 To strings.Count - 1 'pick which string is the shortest - that one is the best one
                    If result.Length > strings(i).length Then result = strings(i)
                Next i
            Else
                result = calc
                If ans.ToString = "error" Then ans = 0
            End If
            If IsNumeric(result) And isCalc = True And My.Settings.grouping = True And result.ToString.Contains("E") = False Then 'group digits so they are easy to read
                result = FormatNumber(calc.ToString, 15, , , TriState.True)
                While result.ToString.EndsWith("0")
                    result = result.ToString.Substring(0, result.ToString.Length - 1)
                End While
                If result.endswith(".") Then result = result.ToString.Substring(0, result.ToString.Length - 1)
            End If
            eval = result
        End If
        Exit Function
err:
        If isCalc Then ans = 0
        eval = "error"
    End Function

    'returns the square root of the largest square of a number
    Private Function ExtractSquare(x)
        If x > Integer.MaxValue Then GoTo err
        Dim n = Math.Floor(Math.Sqrt(x))
        While (x Mod (n * n) <> 0)
            n -= 1
        End While
        ExtractSquare = n
        Exit Function
err:
        ExtractSquare = 1
    End Function

    ' Save changes to plotting window range
    Private Sub ContextMenuStrip1_Closing(sender As Object, e As ToolStripDropDownClosingEventArgs) Handles ContextMenuStrip1.Closing
        On Error Resume Next
        xmin = eval(heal(ToolStripTextBox1.Text), False, False)
        xmax = eval(heal(ToolStripTextBox2.Text), False, False)
        ymin = eval(heal(ToolStripTextBox3.Text), False, False)
        ymax = eval(heal(ToolStripTextBox4.Text), False, False)
        xtick = eval(heal(ToolStripTextBox5.Text), False, False)
        ytick = eval(heal(ToolStripTextBox6.Text), False, False)
        transformX = heal(ToolStripTextBox7.Text).Replace("uv", "u*v")
        transformY = heal(ToolStripTextBox8.Text).Replace("uv", "u*v")
        transformZ = heal(ToolStripTextBox9.Text).Replace("uv", "u*v")
        If xtick <= 0 Then xtick = 1
        If ytick <= 0 Then ytick = 1
        If xmin >= xmax Then
            xmin = xmax - 1
            e.Cancel = True
            MsgBox("Invalid x window range; x min is too large.")
            Exit Sub
        End If
        If ymin >= ymax Then
            ymin = ymax - 1
            e.Cancel = True
            MsgBox("Invalid y window range; y min is too large.")
            Exit Sub
        End If
        PictureBox1.Refresh()
    End Sub

    ' Update plotting window range
    Private Sub ContextMenuStrip1_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening
        ToolStripTextBox1.Text = xmin
        ToolStripTextBox2.Text = xmax
        ToolStripTextBox3.Text = ymin
        ToolStripTextBox4.Text = ymax
        ToolStripTextBox5.Text = xtick
        ToolStripTextBox6.Text = ytick
        ToolStripTextBox7.Text = transformX
        ToolStripTextBox8.Text = transformY
        ToolStripTextBox9.Text = transformZ
        If My.Settings.autofix = True Then
            ToolStripMenuItem2.Text = "Turn autoscale off"
        Else
            ToolStripMenuItem2.Text = "Turn autoscale on"
        End If
    End Sub

    ' Toggle to degree mode
    Private Sub DegreeModeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DegreeModeToolStripMenuItem.Click
        RadianModeToolStripMenuItem.Checked = False
        DegreeModeToolStripMenuItem.Checked = True
        Button1_Click(sender, e)
    End Sub

    ' Toggle to radian mode
    Private Sub RadianModeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RadianModeToolStripMenuItem.Click
        DegreeModeToolStripMenuItem.Checked = False
        RadianModeToolStripMenuItem.Checked = True
        Button1_Click(sender, e)
    End Sub

    'turn on/off speech recognition
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If RichTextBox2.BackColor = Color.White Then
            RichTextBox2.BackColor = Color.PaleGreen
            My.Settings.speechEnabled = True
            reco.RecognizeAsync(System.Speech.Recognition.RecognizeMode.Multiple)
        Else
            RichTextBox2.BackColor = Color.White
            My.Settings.speechEnabled = False
            reco.RecognizeAsyncStop()
        End If
        RichTextBox2.Focus()
    End Sub

    ' Reset to the default viewing range
    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        xmin = My.Settings.x_min
        xmax = My.Settings.x_max
        ymin = My.Settings.y_min
        ymax = My.Settings.y_max
        xtick = 1
        ytick = 1
        xtickMultiplier = 2
        ytickMultiplier = 2
        PictureBox1.Refresh()
    End Sub

    'presses the speech button
    Private Sub ToggleSpeechToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ToggleSpeechToolStripMenuItem.Click
        If Button2.Enabled = True Then Button2.PerformClick()
    End Sub

    ' Close the context menu if the user presses the enter key
    Private Sub ToolStripTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ToolStripTextBox1.KeyDown, ToolStripTextBox2.KeyDown, ToolStripTextBox3.KeyDown, ToolStripTextBox4.KeyDown, ToolStripTextBox5.KeyDown, ToolStripTextBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            ContextMenuStrip1.Close()
        End If
    End Sub

    ' Select all text upon mouse enter
    Private Sub ToolStripTextBox1_MouseEnter(sender As Object, e As EventArgs) Handles ToolStripTextBox1.MouseEnter, ToolStripTextBox2.MouseEnter, ToolStripTextBox3.MouseEnter, ToolStripTextBox4.MouseEnter, ToolStripTextBox5.MouseEnter, ToolStripTextBox6.MouseEnter, ToolStripTextBox7.MouseEnter, ToolStripTextBox8.MouseEnter, ToolStripTextBox9.MouseEnter
        sender.focus()
        sender.selectall()
        sender.tag = ""
    End Sub

    ' Clear all the text from the rich text box
    Private Sub ClearAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearAllToolStripMenuItem.Click
        RichTextBox1.Text = ""
        TrackBar1.Visible = False
        TextBox2.Visible = False
        TextBox3.Visible = False
        all3DFaces.Clear()
        facesAreColored = False
        last3DGraph = ""
        all3DPoints.Clear()
        all3DFunctionNames.Clear()
        PictureBox1.Refresh()
        RichTextBox2.Focus()
    End Sub

    'choose what draw type to have have for the right mouse button
    Private Sub LinearToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LinearToolStripMenuItem.Click, QuadraticToolStripMenuItem.Click, CubicToolStripMenuItem.Click, SimpleLinearToolStripMenuItem.Click, SimpleQuadraticToolStripMenuItem.Click, SimpleCubicToolStripMenuItem.Click
        SimpleLinearToolStripMenuItem.Checked = False
        LinearToolStripMenuItem.Checked = False
        SimpleQuadraticToolStripMenuItem.Checked = False
        QuadraticToolStripMenuItem.Checked = False
        SimpleCubicToolStripMenuItem.Checked = False
        CubicToolStripMenuItem.Checked = False
        sender.checked = True
        My.Settings.drawMode = sender.tag
    End Sub

    'click button to view graphing window
    Private Sub pbxGraph_Click(sender As Object, e As EventArgs) Handles pbxGraph.Click
        PictureBox1.BringToFront()
        TrackBar1.BringToFront()
        TextBox2.BringToFront()
        TextBox3.BringToFront()
        Me.Tag = "graph"
        Dim tempText = RichTextBox2.Text
        Dim tempCursor = RichTextBox2.SelectionStart
        RichTextBox2.Text = ""
        Button1.PerformClick()
        RichTextBox2.Text = tempText
        RichTextBox2.Focus()
        RichTextBox2.Select(tempCursor, 0)
    End Sub

    'click button to view matrix calculator
    Private Sub pbxMatrix_Click(sender As Object, e As EventArgs) Handles pbxMatrix.Click
        palMatrix.BringToFront()
        Me.Tag = "matrix"
        Call Me.Form1_Resize(sender, e)
        TextBox4.Focus()
    End Sub

    'automatically perform matrix calculations
    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged, TextBox5.TextChanged, ListBox1.SelectedIndexChanged
        On Error GoTo onErr
        If ListBox1.SelectedIndex = 0 Or ListBox1.SelectedIndex = 5 Or ListBox1.SelectedIndex = 6 Or ListBox1.SelectedIndex = 8 Then
            Label2.Text = "Matrix B"
        Else
            Label2.Text = "Matrix B (not used for this operation)"
        End If
        'TextBox4.Text = Replace(Replace(TextBox4.Text, ",", vbTab), ";", vbnewline) 'replace commas and semicolons with tabs and new lines
        'TextBox5.Text = Replace(Replace(TextBox5.Text, ",", vbTab), ";", vbnewline)
        Dim throwError As Boolean = False
        Dim columns = Split(TextBox4.Lines(0), vbTab).Count
        Dim resultString = ""
        ReDim A(TextBox4.Lines.Count - 1, columns - 1) 'build the A matrix
        For i = 0 To TextBox4.Lines.Length - 1
            Dim s = Split(TextBox4.Lines(i), vbTab)
            For j = 0 To s.Length - 1
                A(i, j) = eval(s(j), False, False)
            Next j
        Next i
        If ListBox1.SelectedIndex = 0 Or ListBox1.SelectedIndex = 5 Or ListBox1.SelectedIndex = 6 Or ListBox1.SelectedIndex = 8 Then 'build the B matrix
            columns = Split(TextBox5.Lines(0), vbTab).Count
            ReDim B(TextBox5.Lines.Count - 1, columns - 1)
            For i = 0 To TextBox5.Lines.Length - 1
                Dim s = Split(TextBox5.Lines(i), vbTab)
                For j = 0 To s.Length - 1
                    B(i, j) = eval(s(j), False, False)
                Next j
            Next i
        End If
        If ListBox1.SelectedIndex = 0 Then 'Add
            If A.GetLength(0) = B.GetLength(0) And A.GetLength(1) = B.GetLength(1) Then
                ReDim C(A.GetLength(0) - 1, A.GetLength(1) - 1)
                For i = 0 To A.GetLength(0) - 1
                    For j = 0 To A.GetLength(1) - 1
                        C(i, j) = A(i, j) + B(i, j)
                    Next j
                Next i
                resultString = "A (" + A.GetLength(0).ToString + "x" + A.GetLength(1).ToString + ") + B (" + B.GetLength(0).ToString + "x" + B.GetLength(1).ToString + "):" + vbNewLine
            Else
                TextBox6.Text = "Addition error:" + vbNewLine + vbNewLine + "Attempted dimensions:" + vbNewLine + vbNewLine + "A:  " + A.GetLength(0).ToString + "x" + A.GetLength(1).ToString + vbNewLine + "B:  " + B.GetLength(0).ToString + "x" + B.GetLength(1).ToString + vbNewLine + vbNewLine + "Dimensions must be equal."
                throwError = True
            End If
        ElseIf ListBox1.SelectedIndex = 1 Then 'Adjugate
            If A.GetLength(0) = A.GetLength(1) Then
                ReDim C(A.GetLength(0) - 1, A.GetLength(1) - 1)
                For i = 0 To A.GetLength(0) - 1
                    For j = 0 To A.GetLength(0) - 1
                        Dim M = minor(A, i, j)
                        Dim detM = det(M)
                        C(j, i) = (-1) ^ (i + j) * detM
                    Next j
                Next i
                resultString = "Adjugate of A (" + A.GetLength(0).ToString + "x" + A.GetLength(1).ToString + "):" + vbNewLine
            Else
                TextBox6.Text = "Adjugate error:" + vbNewLine + vbNewLine + "Attempted dimensions:" + vbNewLine + vbNewLine + "A:  " + A.GetLength(0).ToString + "x" + A.GetLength(1).ToString + vbNewLine + vbNewLine + "Matrix A's rows and columns must be identical in size."
                throwError = True
            End If
        ElseIf ListBox1.SelectedIndex = 2 Then 'Cofactor
            If A.GetLength(0) = A.GetLength(1) Then
                ReDim C(A.GetLength(0) - 1, A.GetLength(1) - 1)
                For i = 0 To A.GetLength(0) - 1
                    For j = 0 To A.GetLength(0) - 1
                        Dim M = minor(A, i, j)
                        Dim detM = det(M)
                        C(i, j) = (-1) ^ (i + j) * detM
                    Next j
                Next i
                resultString = "Cofactor of A (" + A.GetLength(0).ToString + "x" + A.GetLength(1).ToString + "):" + vbNewLine
            Else
                TextBox6.Text = "Cofactor error:" + vbNewLine + vbNewLine + "Attempted dimensions:" + vbNewLine + vbNewLine + "A:  " + A.GetLength(0).ToString + "x" + A.GetLength(1).ToString + vbNewLine + vbNewLine + "Matrix A's rows and columns must be identical in size."
                throwError = True
            End If
        ElseIf ListBox1.SelectedIndex = 3 Then 'Determinant
            If A.GetLength(0) = A.GetLength(1) Then
                ReDim C(0, 0)
                C(0, 0) = det(A)
                resultString = "Determinant of A (" + A.GetLength(0).ToString + "x" + A.GetLength(1).ToString + "):" + vbNewLine
            Else
                TextBox6.Text = "Determinant error:" + vbNewLine + vbNewLine + "Attempted dimensions:" + vbNewLine + vbNewLine + "A:  " + A.GetLength(0).ToString + "x" + A.GetLength(1).ToString + vbNewLine + vbNewLine + "Matrix A's rows and columns must be identical in size."
                throwError = True
            End If
        ElseIf ListBox1.SelectedIndex = 4 Then 'Inverse
            If A.GetLength(0) = A.GetLength(1) Then
                If det(A) = 0 Then
                    TextBox6.Text = "Inverse error:" + vbNewLine + vbNewLine + "Could not find an inverse of Matrix A." + vbNewLine + vbNewLine + "Cause: Linear dependence."
                    throwError = True
                Else
                    ReDim C(A.GetLength(0) - 1, A.GetLength(1) - 1)
                    For i = 0 To A.GetLength(0) - 1
                        C(i, i) = 1
                    Next i
                    For j = 0 To A.GetLength(0) - 1
                        For i = j To A.GetLength(0) - 1
                            If A(i, j) <> 0 Then
                                For k = 0 To A.GetLength(0) - 1
                                    Dim s As Double
                                    s = A(j, k)
                                    A(j, k) = A(i, k)
                                    A(i, k) = s
                                    s = C(j, k)
                                    C(j, k) = C(i, k)
                                    C(i, k) = s
                                Next k
                                Dim t As Double = 1 / A(j, j)
                                For k = 0 To A.GetLength(0) - 1
                                    A(j, k) = t * A(j, k)
                                    C(j, k) = t * C(j, k)
                                Next k
                                For L = 0 To A.GetLength(0) - 1
                                    If L <> j Then
                                        t = -A(L, j)
                                        For k = 0 To A.GetLength(0) - 1
                                            A(L, k) = A(L, k) + t * A(j, k)
                                            C(L, k) = C(L, k) + t * C(j, k)
                                        Next k
                                    End If
                                Next L
                            End If
                        Next i
                    Next j
                    resultString = "Inverse of A (" + A.GetLength(0).ToString + "x" + A.GetLength(1).ToString + "):" + vbNewLine
                End If
            Else
                TextBox6.Text = "Inverse error:" + vbNewLine + vbNewLine + "Attempted dimensions:" + vbNewLine + vbNewLine + "A:  " + A.GetLength(0).ToString + "x" + A.GetLength(1).ToString + vbNewLine + vbNewLine + "Matrix A's rows and columns must be identical in size."
                throwError = True
            End If
        ElseIf ListBox1.SelectedIndex = 5 Then 'Multiply
            If A.GetLength(1) = B.GetLength(0) Then
                ReDim C(A.GetLength(0) - 1, B.GetLength(1) - 1)
                For i = 0 To A.GetLength(0) - 1
                    For j = 0 To B.GetLength(1) - 1
                        Dim tempSum = 0
                        For k = 0 To A.GetLength(1) - 1
                            tempSum += A(i, k) * B(k, j)
                        Next k
                        C(i, j) = tempSum
                    Next j
                Next i
                resultString = "A (" + A.GetLength(0).ToString + "x" + A.GetLength(1).ToString + ")" + " * B (" + B.GetLength(0).ToString + "x" + B.GetLength(1).ToString + "):" + vbNewLine
            Else
                TextBox6.Text = "Multiplication error:" + vbNewLine + vbNewLine + "Attempted dimensions:" + vbNewLine + vbNewLine + "A:  " + A.GetLength(0).ToString + "x" + A.GetLength(1).ToString + vbNewLine + "B:  " + B.GetLength(0).ToString + "x" + B.GetLength(1).ToString + vbNewLine + vbNewLine + "Matrix A's columns must equal the size of Matrix B's rows."
                throwError = True
            End If
        ElseIf ListBox1.SelectedIndex = 6 Then 'Multiply (Scalar)
            If A.Length > 1 And B.Length > 1 Then
                TextBox6.Text = "Multiplication error:" + vbNewLine + vbNewLine + "Cannot compute scalar multiplication on two matrices." + vbNewLine + "One argument must be a scalar value."
                throwError = True
            Else
                If B.Length = 1 Then
                    Dim temp = A
                    A = B
                    B = temp
                End If
                ReDim C(B.GetLength(0) - 1, B.GetLength(1) - 1)
                For i = 0 To B.GetLength(0) - 1
                    For j = 0 To B.GetLength(1) - 1
                        C(i, j) = A(0, 0) * B(i, j)
                    Next j
                Next i
            End If
            resultString = A(0, 0).ToString + " * (" + B.GetLength(0).ToString + "x" + B.GetLength(1).ToString + "):" + vbNewLine
        ElseIf ListBox1.SelectedIndex = 7 Then 'Row Reduce
            For j = 0 To Math.Min(A.GetLength(0) - 1, A.GetLength(1) - 1)
                For i = j To A.GetLength(0) - 1
                    If A(i, j) <> 0 Then
                        For k = 0 To A.GetLength(1) - 1
                            Dim s As Double
                            s = A(j, k)
                            A(j, k) = A(i, k)
                            A(i, k) = s
                        Next k
                        Dim t As Double = 1 / A(j, j)
                        For k = 0 To A.GetLength(1) - 1
                            A(j, k) = t * A(j, k)
                        Next k
                        For L = 0 To A.GetLength(0) - 1
                            If L <> j Then
                                t = -A(L, j)
                                For k = 0 To A.GetLength(1) - 1
                                    A(L, k) = A(L, k) + t * A(j, k)
                                Next k
                            End If
                        Next L
                    End If
                Next i
            Next j
            resultString = "Matrix A (" + A.GetLength(0).ToString + "x" + A.GetLength(1).ToString + ") reduced to Echelon form:" + vbNewLine
            C = A
        ElseIf ListBox1.SelectedIndex = 8 Then 'Subtract
            If A.GetLength(0) = B.GetLength(0) And A.GetLength(1) = B.GetLength(1) Then
                ReDim C(A.GetLength(0) - 1, A.GetLength(1) - 1)
                For i = 0 To A.GetLength(0) - 1
                    For j = 0 To A.GetLength(1) - 1
                        C(i, j) = A(i, j) - B(i, j)
                    Next j
                Next i
                resultString = "A (" + A.GetLength(0).ToString + "x" + A.GetLength(1).ToString + ") - B (" + B.GetLength(0).ToString + "x" + B.GetLength(1).ToString + "):" + vbNewLine
            Else
                TextBox6.Text = "Subtraction error:" + vbNewLine + vbNewLine + "Attempted dimensions:" + vbNewLine + vbNewLine + "A:  " + A.GetLength(0).ToString + "x" + A.GetLength(1).ToString + vbNewLine + "B:  " + B.GetLength(0).ToString + "x" + B.GetLength(1).ToString + vbNewLine + vbNewLine + "Dimensions must be equal."
                throwError = True
            End If
        ElseIf ListBox1.SelectedIndex = 9 Then 'trace
            If A.GetLength(0) = A.GetLength(1) Then
                ReDim C(0, 0)
                For i = 0 To A.GetLength(0) - 1
                    C(0, 0) += A(i, i)
                Next i
                resultString = "Trace of A (" + A.GetLength(0).ToString + "x" + A.GetLength(1).ToString + "):" + vbNewLine
            Else
                TextBox6.Text = "Trace error:" + vbNewLine + vbNewLine + "Attempted dimensions:" + vbNewLine + vbNewLine + "A:  " + A.GetLength(0).ToString + "x" + A.GetLength(1).ToString + vbNewLine + vbNewLine + "Matrix A's rows and columns must be identical in size."
                throwError = True
            End If
        ElseIf ListBox1.SelectedIndex = 10 Then 'Transpose
            ReDim C(A.GetLength(1) - 1, A.GetLength(0) - 1)
            For i = 0 To A.GetLength(0) - 1
                For j = 0 To A.GetLength(1) - 1
                    C(j, i) = A(i, j)
                Next j
            Next i
            resultString = "Transpose of A (" + A.GetLength(0).ToString + "x" + A.GetLength(1).ToString + "):" + vbNewLine
        End If
        'Print out the results
        If throwError = False Then
            For i = 0 To C.GetLength(0) - 1
                resultString = resultString + vbNewLine
                For j = 0 To C.GetLength(1) - 1
                    If My.Settings.fraction Then
                        resultString = resultString + eval(C(i, j), True, False).ToString + "    " + vbTab
                    Else
                        resultString = resultString + Math.Round(C(i, j), 5).ToString + " " + vbTab
                    End If
                Next j
                resultString = resultString.Substring(0, resultString.Length - 1)
                resultString = RTrim(resultString)
            Next i
            TextBox6.Text = resultString
        End If
        Exit Sub
onErr:
        TextBox6.Text = "Input error:" + vbNewLine + vbNewLine + "Photon could not interpret your text as numbers."
    End Sub

    'calculates the determinant of a matrix
    Private Function det(D(,) As Double) As Double
        If D.GetLength(0) = 1 Then
            det = D(0, 0)
        Else
            Dim total As Double = 0
            For i = 0 To D.GetLength(0) - 1
                total += ((-1) ^ i) * D(0, i) * det(minor(D, 0, i))
            Next i
            det = total
        End If
    End Function

    'deletes all data from the dataXY grid
    Private Sub ClearAllDataToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearAllDataToolStripMenuItem.Click
        DataXY.Rows.Clear()
    End Sub

    'eliminates a row and column from a matrix
    Private Function minor(P(,) As Double, row As Integer, column As Integer) As Double(,)
        Dim D(P.GetLength(0) - 2, P.GetLength(1) - 2) As Double
        Dim activeRow = 0
        Dim activeColumn = 0
        For i = 0 To P.GetLength(0) - 1
            If i <> row Then
                activeColumn = 0
                For j = 0 To P.GetLength(1) - 1
                    If j <> column Then
                        D(activeRow, activeColumn) = P(i, j)
                        activeColumn += 1
                    End If
                Next j
                activeRow += 1
            End If
        Next i
        minor = D
    End Function

    ' Colors the text
    Public Sub ColorTheText()
        Dim locality = 0
        Dim calcColor = Color.DeepSkyBlue
        Dim variableColor = Color.Red
        Dim symbolColor = Color.White
        Dim functioncolor = Color.LimeGreen
        Dim commentColor = Color.LightGray
        If My.Settings.projector Then
            calcColor = Color.Blue
            variableColor = Color.Red
            symbolColor = Color.Black
            functioncolor = Color.Green
            commentColor = Color.DimGray
        End If
        For i = 0 To RichTextBox1.Lines.Length - 1
            RichTextBox1.SelectionStart = locality
            If RichTextBox1.Lines(i).StartsWith("'") Or RichTextBox1.Lines(i).StartsWith("trace") Then
                RichTextBox1.SelectionLength = RichTextBox1.Lines(i).Length
                RichTextBox1.SelectionColor = commentColor
                GoTo theEnd
            End If
            Dim s = Split(RichTextBox1.Lines(i), "->")
            If s.Length = 2 Then
                RichTextBox1.SelectionLength = s(0).Length
                RichTextBox1.SelectionColor = calcColor
                RichTextBox1.SelectionStart = locality + s(0).Length
                RichTextBox1.SelectionLength = 2
                RichTextBox1.SelectionColor = symbolColor
                RichTextBox1.SelectionStart = locality + s(0).Length + 2
                RichTextBox1.SelectionLength = s(1).Length
                RichTextBox1.SelectionColor = calcColor
                GoTo theEnd
            End If
            s = Split(RichTextBox1.Lines(i), ":=")
            If s.Length = 2 Then
                RichTextBox1.SelectionLength = s(0).Length
                RichTextBox1.SelectionColor = variableColor
                RichTextBox1.SelectionStart = locality + s(0).Length
                RichTextBox1.SelectionLength = 2
                RichTextBox1.SelectionColor = symbolColor
                RichTextBox1.SelectionStart = locality + s(0).Length + 2
                RichTextBox1.SelectionLength = s(1).Length
                RichTextBox1.SelectionColor = variableColor
                GoTo theEnd
            End If
            If RichTextBox1.Lines(i).Contains("=") Or (RichTextBox1.Lines(i).StartsWith("[") And RichTextBox1.Lines(i).EndsWith("]")) Then
                RichTextBox1.SelectionLength = RichTextBox1.Lines(i).Length
                RichTextBox1.SelectionColor = functioncolor
            End If
theEnd:
            locality = locality + RichTextBox1.Lines(i).Length + 1
            RichTextBox1.SelectionLength = 0
        Next i
    End Sub

    'click button for data window
    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        palData.BringToFront()
        Me.Tag = "data"
        DataXY.Focus()
    End Sub

    ' Saves to the active file.  If there is no active file, the save dialog box appears
    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        If fileCreated = "" Then
            SaveAsToolStripMenuItem_Click(sender, e)
        Else
            Dim fOut As New System.IO.StreamWriter(fileCreated)
            fOut.Write(Replace(RichTextBox1.Text, vbNewLine, vbNewLine))
            fOut.Write("<split>" + xmin.ToString + "<split>" + xmax.ToString + "<split>" + ymin.ToString + "<split>" + ymax.ToString + "<split>" + xtick.ToString + "<split>" + ytick.ToString + "<split>" + LogXAxisToolStripMenuItem.Checked.ToString + "<split>" + LogYAxisToolStripMenuItem.Checked.ToString + "<split>" + transformX + "<split>" + transformY + "<split>" + transformZ)
            fOut.Write("<split>")
            For i = 0 To DataXY.Rows.Count - 2
                Dim firstNumber As Double = Val(DataXY.Rows(i).Cells(0).Value)
                Dim secondNumber As Double = Val(DataXY.Rows(i).Cells(1).Value)
                fOut.Write(firstNumber.ToString + "," + secondNumber.ToString + "/")
            Next i
            fOut.Close()
            changesOccured = False
            Me.Text = Split(Split(fileCreated, "\")(Split(fileCreated, "\").Length - 1), ".")(0) + " - Photon"
        End If
    End Sub
    ' Shows the save dialog box to create a file to save to
    Private Sub SaveAsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        SaveFileDialog1.FileName = Split(Me.Text, " - ")(0)
        Dim result = SaveFileDialog1.ShowDialog
        If result = vbOK Then
            Dim newBitmap As Bitmap = New Bitmap(PictureBox1.Width, PictureBox1.Height)
            PictureBox1.DrawToBitmap(newBitmap, New Rectangle(0, 0, PictureBox1.Width, PictureBox1.Height))
            PictureBox1.Image = newBitmap
            If SaveFileDialog1.FileName.EndsWith(".photon") Then
                fileCreated = SaveFileDialog1.FileName
                SaveToolStripMenuItem_Click(sender, e)
            ElseIf SaveFileDialog1.FileName.EndsWith(".jpg") Then
                PictureBox1.Image.Save(SaveFileDialog1.FileName, System.Drawing.Imaging.ImageFormat.Jpeg)
            ElseIf SaveFileDialog1.FileName.EndsWith(".png") Then
                PictureBox1.Image.Save(SaveFileDialog1.FileName, System.Drawing.Imaging.ImageFormat.Png)
            ElseIf SaveFileDialog1.FileName.EndsWith(".bmp") Then
                PictureBox1.Image.Save(SaveFileDialog1.FileName, System.Drawing.Imaging.ImageFormat.Bmp)
            ElseIf SaveFileDialog1.FileName.EndsWith(".gif") Then
                PictureBox1.Image.Save(SaveFileDialog1.FileName, System.Drawing.Imaging.ImageFormat.Gif)
            ElseIf SaveFileDialog1.FileName.EndsWith(".txt") Then
                Dim fOut As New System.IO.StreamWriter(SaveFileDialog1.FileName)
                fOut.Write(Replace(RichTextBox1.Text, vbNewLine, vbNewLine))
                fOut.Close()
            End If
            PictureBox1.Image = Nothing
        End If
        SaveFileDialog1.Dispose()
    End Sub

    'gets rid of the empty rows in the xy data sheet
    Private Sub RemoveEmptyRowsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RemoveEmptyRowsToolStripMenuItem.Click
        DataXY.CurrentCell = Nothing
        For i = DataXY.Rows.Count - 2 To 0 Step -1
            If DataXY.Rows(i).Cells(0).Value Is Nothing And DataXY.Rows(i).Cells(1).Value Is Nothing Then
                DataXY.Rows.RemoveAt(i)
            End If
        Next i
    End Sub

    ' Opens a file
    Private Sub openFile(fileName As String)
        Try
            Dim fIn As New System.IO.StreamReader(fileName)
            Dim tempString = Split(fIn.ReadToEnd(), "<split>")
            fIn.Close()
            If tempString.Length >= 1 Then
                RichTextBox1.Text = tempString(0)
                If fileName.EndsWith(".txt") Then
                    shouldReset = True
                    Dim fOut As New System.IO.StreamWriter(fileName)
                    Dim outString As String = ""
                    For Each l In RichTextBox1.Lines
                        outString = outString + calculate(heal(l)) + vbNewLine
                    Next l
                    outString = outString.Substring(0, outString.Length - 1)
                    fOut.Write(outString)
                    fOut.Close()
                    End
                End If
            End If
            If tempString.Length >= 12 Then
                xmin = Val(tempString(1))
                xmax = Val(tempString(2))
                ymin = Val(tempString(3))
                ymax = Val(tempString(4))
                xtick = Val(tempString(5))
                ytick = Val(tempString(6))
                LogXAxisToolStripMenuItem.Checked = Val(tempString(7))
                LogYAxisToolStripMenuItem.Checked = Val(tempString(8))
                transformX = tempString(9)
                transformY = tempString(10)
                transformZ = tempString(11)
            End If
            If tempString.Length >= 13 Then
                Dim dataPoints() As String = Split(tempString(9), "/")
                For i = 0 To dataPoints.Length - 2
                    Dim pt = Split(dataPoints(i), ",")
                    Dim x = pt(0)
                    Dim y = pt(1)
                    DataXY.Rows.Add(x, y)
                Next i
                UpdateData()
            End If
            If My.Settings.colorText Then ColorTheText()
            changesOccured = False
            fileCreated = fileName
            Me.Text = Split(Split(fileCreated, "\")(Split(fileCreated, "\").Length - 1), ".")(0) + " - Photon"
        Catch e As Exception
            MsgBox("Oops!  An error occured while attempting to read this file.")
        End Try
    End Sub

    ' Opens the open dialog box to open a file
    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        If changesOccured And My.Settings.saveOption = 1 Then
            Dim r = MsgBox("Would you like to save your work before opening another file?", vbYesNoCancel)
            If r = vbYes Then
                SaveToolStripMenuItem_Click(sender, e)
            ElseIf r = vbCancel Then
                Exit Sub
            Else
                'continue on our merry way
            End If
        End If
        Dim result = OpenFileDialog1.ShowDialog
        If result = vbOK Then
            If OpenFileDialog1.FileName.EndsWith(".photon") Then
                openFile(OpenFileDialog1.FileName)
                Button1_Click(sender, e)
            ElseIf OpenFileDialog1.FileName.EndsWith(".txt") Then
                Dim fIn As New System.IO.StreamReader(OpenFileDialog1.FileName)
                RichTextBox1.Text = fIn.ReadToEnd()
                shouldReset = True
                PictureBox1.Refresh()
                shouldReset = False
            End If
        End If
        OpenFileDialog1.Dispose()
    End Sub

    'go to function table window
    Private Sub pbxTable_Click(sender As Object, e As EventArgs) Handles pbxTable.Click
        palTable.BringToFront()
        Me.Tag = "table"
        Button1.PerformClick()
    End Sub

    'automatically refresh the function table
    Private Sub txtStart_TextChanged(sender As Object, e As EventArgs) Handles txtStart.TextChanged, txtDelta.TextChanged
        Button1.PerformClick()
    End Sub

    ' forces the graph to be redrawn when the window is resized; sometimes the graph got grumpy and didn't redraw
    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize, SplitContainer1.SplitterMoved
        If Me.Tag = "graph" Then
            RichTextBox2.Focus()
        Else
            TextBox4.Width = palMatrix.Width / 2 - 15
            TextBox5.Left = palMatrix.Width / 2 + 5
            TextBox5.Width = TextBox4.Width
            Label2.Left = TextBox5.Left - 3
            TextBox4.Height = palMatrix.Height / 2 - 67
            TextBox5.Height = TextBox4.Height
            TextBox6.Top = palMatrix.Height / 2 + 20
            Label3.Top = palMatrix.Height / 2 - 32
            Label4.Top = palMatrix.Height / 2
            TextBox6.Height = TextBox4.Height + 35
            ListBox1.Top = TextBox6.Top
            ListBox1.Height = TextBox6.Height + 4
            PictureBox1.Refresh()
        End If
    End Sub
    ' toggles between automatically adjusting the aspect ratio of the graph
    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        My.Settings.autofix = My.Settings.autofix Xor True
        PictureBox1.Refresh()
    End Sub
    ' returns a string that is a simplified fraction of the given decimal number
    Private Function decToFraction(x As Decimal, Optional shouldTrim As Boolean = True)
        Dim numerator As Long, denominator As Long
        Dim n As Long = Int(Math.Floor(x))
        Dim errMarg = eps
        x -= n
        If x < errMarg Then
            numerator = n
            denominator = 1
            GoTo theEnd
        ElseIf x >= 1 - errMarg Then
            numerator = n + 1
            denominator = 1
            GoTo theEnd
        End If
        Dim lowerN As Integer = 0
        Dim lowerD As Integer = 1
        Dim upperN As Integer = 1
        Dim upperD As Integer = 1
        Dim middleN As Decimal, middleD As Decimal
        Dim count = 0
        While count < 1000000
            middleN = lowerN + upperN
            middleD = lowerD + upperD
            If middleD * (x + errMarg) < middleN Then
                upperN = middleN
                upperD = middleD
            ElseIf middleN < (x - errMarg) * middleD Then
                lowerN = middleN
                lowerD = middleD
            Else
                numerator = n * middleD + middleN
                denominator = middleD
                GoTo theEnd
            End If
            count += 1
        End While
        numerator = n * middleD + middleN
        denominator = middleD
theEnd:
        Dim result = numerator.ToString + "/" + denominator.ToString
        If shouldTrim = True Then
            If result.EndsWith("/1") Then result = numerator.ToString
            If result.StartsWith("0/") Then result = "0"
        End If
        decToFraction = result
    End Function

    'displays the units window
    Private Sub bpxScale_Click(sender As Object, e As EventArgs) Handles bpxScale.Click
        Me.Tag = "units"
        palUnits.BringToFront()
        TextBox7.Focus()
    End Sub

    'automatically calculate unit conversions
    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        Dim totalString As String = Replace(TextBox7.Text, "  ", " ")
        totalString = Replace(totalString, "per", "/")
        totalString = Replace(totalString, " / ", "/")
        Dim twoParts() As String = Split(totalString, " ")
        If twoParts.Length = 2 Then
            If twoParts(0).Length > 0 And twoParts(1).Length > 0 And IsNumeric(twoParts(0)) Then
                Dim number As Double = twoParts(0)
                Dim unit As String = twoParts(1)
                unit = Replace(unit, "^", "")
                unit = Replace(unit, ".", "")
                unit = Replace(unit, "-", "")
                unit = Replace(unit, "s/", "/")
                If unit.EndsWith("s") And unit.Length > 3 Then unit = unit.Substring(0, unit.Length - 1)
                If prefixDictionary.ContainsKey(unit) Then
                    TextBox9.Text = "Metric Prefix" + vbNewLine + vbNewLine
                    convertUnits(prefixDictionary, unit, number)
                ElseIf lengthDictionary.ContainsKey(unit) Then
                    TextBox9.Text = "Distance" + vbNewLine + vbNewLine
                    convertUnits(lengthDictionary, unit, number)
                ElseIf areaDictionary.ContainsKey(unit) Then
                    TextBox9.Text = "Area" + vbNewLine + vbNewLine
                    convertUnits(areaDictionary, unit, number)
                ElseIf energyDictionary.ContainsKey(unit) Then
                    TextBox9.Text = "Energy" + vbNewLine + vbNewLine
                    convertUnits(energyDictionary, unit, number)
                ElseIf fuelDictionary.ContainsKey(unit) Then
                    TextBox9.Text = "Fuel" + vbNewLine + vbNewLine
                    convertUnits(fuelDictionary, unit, number)
                ElseIf massDictionary.ContainsKey(unit) Then
                    TextBox9.Text = "Mass/Force" + vbNewLine + vbNewLine
                    convertUnits(massDictionary, unit, number)
                ElseIf angleDictionary.ContainsKey(unit) Then
                    TextBox9.Text = "Angle" + vbNewLine + vbNewLine
                    convertUnits(angleDictionary, unit, number)
                ElseIf pressureDictionary.ContainsKey(unit) Then
                    TextBox9.Text = "Pressure" + vbNewLine + vbNewLine
                    convertUnits(pressureDictionary, unit, number)
                ElseIf speedDictionary.ContainsKey(unit) Then
                    TextBox9.Text = "Speed" + vbNewLine + vbNewLine
                    convertUnits(speedDictionary, unit, number)
                ElseIf timeDictionary.ContainsKey(unit) Then
                    TextBox9.Text = "Time" + vbNewLine + vbNewLine
                    convertUnits(timeDictionary, unit, number)
                ElseIf volumeDictionary.ContainsKey(unit) Then
                    TextBox9.Text = "Volume" + vbNewLine + vbNewLine
                    convertUnits(volumeDictionary, unit, number)
                ElseIf densityDictionary.ContainsKey(unit) Then
                    TextBox9.Text = "Density" + vbNewLine + vbNewLine
                    convertUnits(densityDictionary, unit, number)
                ElseIf accelerationDictionary.ContainsKey(unit) Then
                    TextBox9.Text = "Acceleration" + vbNewLine + vbNewLine
                    convertUnits(accelerationDictionary, unit, number)
                ElseIf powerDictionary.ContainsKey(unit) Then
                    TextBox9.Text = "Power" + vbNewLine + vbNewLine
                    convertUnits(powerDictionary, unit, number)
                ElseIf unit = "Kelvin" Or unit = "K" Or unit = "Celcius" Or unit = "C" Or unit = "Fahrenheit" Or unit = "F" Then
                    Dim theText = "Conversion" + vbNewLine + vbNewLine
                    TextBox9.Text = "Temperature" + vbNewLine + vbNewLine + "Kelvin" + vbNewLine + "Celsius" + vbNewLine + "Fahrenheit"
                    If unit = "Kelvin" Or unit = "K" Then
                        theText += number.ToString + vbNewLine
                        theText += (number - 273.15).ToString + vbNewLine
                        theText += ((number - 273.15) * 1.8 + 32).ToString + vbNewLine
                    ElseIf unit = "Celcius" Or unit = "C" Then
                        theText += (number + 273.15).ToString + vbNewLine
                        theText += number.ToString + vbNewLine
                        theText += (number * 1.8 + 32).ToString + vbNewLine
                    ElseIf unit = "Fahrenheit" Or unit = "F" Then
                        theText += (((number - 32) * 5 / 9) + 273.15).ToString + vbNewLine
                        theText += ((number - 32) * 5 / 9).ToString + vbNewLine
                        theText += number.ToString + vbNewLine
                    End If
                    TextBox8.Text = theText
                Else
                    TextBox8.Text = "No recognized units."
                    TextBox9.Text = ""
                End If
            Else
                TextBox8.Text = ""
                TextBox9.Text = ""
            End If
        ElseIf twoParts.Length = 3 Then
            If twoParts(0).Length > 0 And twoParts(1).Length > 0 And twoParts(2).Length > 0 And IsNumeric(twoParts(0)) Then
                Try
                    Dim stringToEvaluate As String = ""
                    Dim currentElement As String = ""
                    For Each letter As Char In twoParts(2)
                        If Char.IsLetter(letter) Then
                            If Char.IsUpper(letter) And currentElement <> "" Then
                                stringToEvaluate += "+" + elementDictionary(currentElement).ToString + "+"
                                currentElement = letter
                            Else
                                currentElement += letter
                            End If
                        ElseIf Char.IsNumber(letter) Then
                            If currentElement <> "" Then
                                stringToEvaluate += "+" + elementDictionary(currentElement).ToString + "*" + letter
                                currentElement = ""
                            Else
                                If stringToEvaluate <> "" Then stringToEvaluate += letter
                            End If
                        Else
                            If currentElement <> "" Then
                                stringToEvaluate += "+" + elementDictionary(currentElement).ToString + letter
                                currentElement = ""
                            Else
                                stringToEvaluate += letter
                            End If
                        End If
                    Next letter
                    If currentElement <> "" Then stringToEvaluate += "+" + elementDictionary(currentElement).ToString
                    stringToEvaluate = Replace(stringToEvaluate, ")", ")*")
                    stringToEvaluate = Replace(stringToEvaluate, "(", "+(")
                    Dim molarMass As Double = eval(stringToEvaluate, False, False)
                    TextBox8.Text = "Conversion" + vbNewLine + vbNewLine
                    TextBox9.Text = "Chemical" + vbNewLine + vbNewLine + "gram" + vbNewLine + "mole" + vbNewLine + "part"
                    If twoParts(1).StartsWith("g") Then
                        TextBox8.Text += twoParts(0).ToString + vbNewLine
                        TextBox8.Text += (twoParts(0) / molarMass).ToString + vbNewLine
                        TextBox8.Text += (twoParts(0) * 6.0221409E+23 / molarMass).ToString + vbNewLine
                    ElseIf twoParts(1).StartsWith("mol") Then
                        TextBox8.Text += (twoParts(0) * molarMass).ToString + vbNewLine
                        TextBox8.Text += twoParts(0).ToString + vbNewLine
                        TextBox8.Text += (twoParts(0) * 6.0221409E+23).ToString + vbNewLine
                    ElseIf twoParts(1).StartsWith("part") Then
                        TextBox8.Text += (twoParts(0) * molarMass / 6.0221409E+23).ToString + vbNewLine
                        TextBox8.Text += (twoParts(0) / 6.0221409E+23).ToString + vbNewLine
                        TextBox8.Text += twoParts(0).ToString
                    Else
                        TextBox8.Text = ""
                        TextBox9.Text = ""
                    End If
                Catch
                    TextBox8.Text = ""
                    TextBox9.Text = ""
                End Try
            Else
                TextBox8.Text = ""
                TextBox9.Text = ""
            End If
        Else
            TextBox8.Text = ""
            TextBox9.Text = ""
        End If
    End Sub

    Private Sub convertUnits(theDictionary As Dictionary(Of String, Double), unit As String, number As Double)
        Dim masterNumber = number * theDictionary(unit)
        Dim thetext = ""
        Dim previousValue As Double = 0
        Dim numb As Double = 0
        TextBox8.Text = "Conversion" + vbNewLine + vbNewLine
        For Each kvp As KeyValuePair(Of String, Double) In theDictionary
            Dim theKey As String = kvp.Key
            numb = kvp.Value
            If previousValue <> numb Then
                Dim result As String = (masterNumber / numb).ToString
                If result.ToString.Contains("E") = False Then
                    result = FormatNumber(masterNumber / numb, 15, , , TriState.True)
                    While result.ToString.EndsWith("0")
                        result = result.ToString.Substring(0, result.ToString.Length - 1)
                    End While
                    If result.EndsWith(".") Then result = result.ToString.Substring(0, result.ToString.Length - 1)
                End If
                thetext += result + vbNewLine
                TextBox9.Text += theKey + vbNewLine
            End If
            previousValue = numb
        Next kvp
        TextBox8.Text += thetext
    End Sub

    'builds a dictionary for unit conversions
    Private Sub BuildUnitDictionary()
        'prefix
        prefixDictionary.Add("yotta", 1.0E+24)
        prefixDictionary.Add("zetta", 1.0E+21)
        prefixDictionary.Add("exa", 1.0E+18)
        prefixDictionary.Add("peta", 1.0E+15)
        prefixDictionary.Add("tera", 1000000000000.0)
        prefixDictionary.Add("giga", 1000000000.0)
        prefixDictionary.Add("mega", 1000000.0)
        prefixDictionary.Add("kilo", 1000.0)
        prefixDictionary.Add("hecto", 100.0)
        prefixDictionary.Add("deca", 10.0)
        prefixDictionary.Add("u", 1)
        prefixDictionary.Add("deci", 0.1)
        prefixDictionary.Add("centi", 0.01)
        prefixDictionary.Add("milli", 0.001)
        prefixDictionary.Add("micro", 0.000001)
        prefixDictionary.Add("nano", 0.000000001)
        prefixDictionary.Add("pico", 0.000000000001)
        prefixDictionary.Add("femto", 0.000000000000001)
        prefixDictionary.Add("atto", 1.0E-18)
        prefixDictionary.Add("zepto", 1.0E-21)
        prefixDictionary.Add("yocto", 1.0E-24)
        'length
        lengthDictionary.Add("meter", 1)
        lengthDictionary.Add("m", 1)
        lengthDictionary.Add("kilometer", 1000)
        lengthDictionary.Add("km", 1000)
        lengthDictionary.Add("centimeter", 0.01)
        lengthDictionary.Add("cm", 0.01)
        lengthDictionary.Add("millimeter", 0.001)
        lengthDictionary.Add("mm", 0.001)
        lengthDictionary.Add("micrometer", 0.000001)
        lengthDictionary.Add("nanometer", 0.000000001)
        lengthDictionary.Add("nm", 0.000000001)
        lengthDictionary.Add("mile", 1609.344)
        lengthDictionary.Add("mi", 1609.344)
        lengthDictionary.Add("yard", 0.9144)
        lengthDictionary.Add("yd", 0.9144)
        lengthDictionary.Add("foot", 0.3048)
        lengthDictionary.Add("ft", 0.3048)
        lengthDictionary.Add("feet", 0.3048)
        lengthDictionary.Add("inch", 0.0254)
        lengthDictionary.Add("in", 0.0254)
        lengthDictionary.Add("inche", 0.0254)
        lengthDictionary.Add("nautical-mile", 1852)
        lengthDictionary.Add("AU", 149600000000.0)
        lengthDictionary.Add("au", 149600000000.0)
        lengthDictionary.Add("lightyear", 9.461E+15)
        lengthDictionary.Add("ly", 9.461E+15)
        lengthDictionary.Add("parsec", 3.086E+16)
        lengthDictionary.Add("pc", 3.086E+16)
        'area
        areaDictionary.Add("m2", 1)
        areaDictionary.Add("km2", 1000000.0)
        areaDictionary.Add("in2", 0.00064516)
        areaDictionary.Add("ft2", 0.092903)
        areaDictionary.Add("yd2", 0.836127)
        areaDictionary.Add("mi2", 2590000.0)
        areaDictionary.Add("hectare", 10000)
        areaDictionary.Add("ha", 10000)
        areaDictionary.Add("acre", 4046.86)
        'energy
        energyDictionary.Add("Joule", 1)
        energyDictionary.Add("J", 1)
        energyDictionary.Add("kiloJoule", 1000)
        energyDictionary.Add("kJ", 1000)
        energyDictionary.Add("calorie", 4.184)
        energyDictionary.Add("cal", 4.184)
        energyDictionary.Add("food-calorie", 4184)
        energyDictionary.Add("Cal", 4184)
        energyDictionary.Add("kilocalorie", 4184)
        energyDictionary.Add("kcal", 4184)
        energyDictionary.Add("watthour", 3600)
        energyDictionary.Add("Wh", 3600)
        energyDictionary.Add("kilowatthour", 3600000)
        energyDictionary.Add("kWh", 3600000)
        energyDictionary.Add("electronvolt", 1.6022E-19)
        energyDictionary.Add("ev", 1.6022E-19)
        energyDictionary.Add("BTU", 1055.06)
        energyDictionary.Add("btu", 1055.06)
        energyDictionary.Add("therm", 105500000.0)
        energyDictionary.Add("thm", 105500000.0)
        energyDictionary.Add("footpound", 1.35582)
        energyDictionary.Add("ftlb", 1.35582)
        'fuel consumption
        fuelDictionary.Add("kilometer/liter", 1)
        fuelDictionary.Add("km/l", 1)
        fuelDictionary.Add("mile/gallon", 0.425144)
        fuelDictionary.Add("mi/gal", 0.425144)
        fuelDictionary.Add("mpg", 0.425144)
        'mass/force
        massDictionary.Add("Newton", 0.101972)
        massDictionary.Add("N", 0.101972)
        massDictionary.Add("pound", 0.453592)
        massDictionary.Add("lb", 0.453592)
        massDictionary.Add("dyne", 0.00000101972)
        massDictionary.Add("kilogram", 1)
        massDictionary.Add("kg", 1)
        massDictionary.Add("ounce", 0.0283495)
        massDictionary.Add("oz", 0.0283495)
        massDictionary.Add("gram", 0.001)
        massDictionary.Add("g", 0.001)
        massDictionary.Add("milligram", 0.000001)
        massDictionary.Add("mg", 0.000001)
        massDictionary.Add("microgram", 0.000000001)
        massDictionary.Add("ton", 907.185)
        'angles
        angleDictionary.Add("revolution", 6.28318530717959)
        angleDictionary.Add("rotation", 6.28318530717959)
        angleDictionary.Add("rev", 6.28318530717959)
        angleDictionary.Add("radian", 1)
        angleDictionary.Add("rad", 1)
        angleDictionary.Add("degree", 0.0174533)
        angleDictionary.Add("deg", 0.0174533)
        angleDictionary.Add("gradian", 0.015708)
        angleDictionary.Add("grad", 0.015708)
        angleDictionary.Add("arcminute", 0.000290888)
        angleDictionary.Add("arcsecond", 0.0000048481)
        'pressure
        pressureDictionary.Add("Pascal", 1)
        pressureDictionary.Add("pascal", 1)
        pressureDictionary.Add("N/m2", 1)
        pressureDictionary.Add("pa", 1)
        pressureDictionary.Add("Pa", 1)
        pressureDictionary.Add("bar", 100000)
        pressureDictionary.Add("lb/in2", 6894.76)
        pressureDictionary.Add("torr", 133.322)
        'density
        densityDictionary.Add("kilogram/liter", 1000)
        densityDictionary.Add("kg/l", 1000)
        densityDictionary.Add("kilogram/meter3", 1)
        densityDictionary.Add("kg/m3", 1)
        densityDictionary.Add("gram/centimeter3", 1000)
        densityDictionary.Add("g/ml", 1000)
        densityDictionary.Add("g/cm3", 1000)
        densityDictionary.Add("ounce/foot3", 1.001154)
        densityDictionary.Add("ounce/feet3", 1.001154)
        densityDictionary.Add("oz/m3", 1.001154)
        densityDictionary.Add("pound/foot3", 16.018463)
        densityDictionary.Add("pound/feet3", 16.018463)
        densityDictionary.Add("lb/ft3", 16.018463)
        'speed
        speedDictionary.Add("meter/second", 1)
        speedDictionary.Add("m/s", 1)
        speedDictionary.Add("centimeter/second", 0.01)
        speedDictionary.Add("cm/", 0.01)
        speedDictionary.Add("mile/hour", 0.44704)
        speedDictionary.Add("mph", 0.44704)
        speedDictionary.Add("mi/h", 0.44704)
        speedDictionary.Add("kilometer/hour", 0.277778)
        speedDictionary.Add("km/h", 0.277778)
        speedDictionary.Add("feet/second", 0.3048)
        speedDictionary.Add("ft/", 0.3048)
        speedDictionary.Add("speed-of-light", 299792458)
        speedDictionary.Add("c", 299792458)
        speedDictionary.Add("knot", 0.514444)
        'time
        timeDictionary.Add("nanosecond", 0.000000001)
        timeDictionary.Add("ns", 0.000000001)
        timeDictionary.Add("microsecond", 0.000001)
        timeDictionary.Add("millisecond", 0.001)
        timeDictionary.Add("ms", 0.001)
        timeDictionary.Add("second", 1)
        timeDictionary.Add("sec", 1)
        timeDictionary.Add("s", 1)
        timeDictionary.Add("minute", 60)
        timeDictionary.Add("min", 60)
        timeDictionary.Add("hour", 3600)
        timeDictionary.Add("hr", 3600)
        timeDictionary.Add("h", 3600)
        timeDictionary.Add("day", 86400)
        timeDictionary.Add("week", 604800)
        timeDictionary.Add("month", 2628000.0)
        timeDictionary.Add("year", 31540000.0)
        timeDictionary.Add("decade", 315400000.0)
        timeDictionary.Add("century", 3154000000.0)
        timeDictionary.Add("millennium", 31540000000.0)
        'volume
        volumeDictionary.Add("m3", 1)
        volumeDictionary.Add("gallon", 0.00378541)
        volumeDictionary.Add("gal", 0.00378541)
        volumeDictionary.Add("quart", 0.000946353)
        volumeDictionary.Add("qt", 0.000946353)
        volumeDictionary.Add("pint", 0.000473176)
        volumeDictionary.Add("pt", 0.000473176)
        volumeDictionary.Add("cup", 0.000236588)
        volumeDictionary.Add("fluid-ounce", 0.00002957353)
        volumeDictionary.Add("floz", 0.00002957353)
        volumeDictionary.Add("tablespoon", 0.0000147867648)
        volumeDictionary.Add("tbsp", 0.0000147867648)
        volumeDictionary.Add("T", 0.0000147867648)
        volumeDictionary.Add("teaspoon", 0.00000492892)
        volumeDictionary.Add("tsp", 0.00000492892)
        volumeDictionary.Add("liter", 0.001)
        volumeDictionary.Add("l", 0.001)
        volumeDictionary.Add("milliliter", 0.000001)
        volumeDictionary.Add("ml", 0.000001)
        volumeDictionary.Add("ft3", 0.0283168)
        volumeDictionary.Add("in3", 0.000016387)
        'acceleration
        accelerationDictionary.Add("meter/second2", 1)
        accelerationDictionary.Add("m/s2", 1)
        accelerationDictionary.Add("foot/second2", 0.3048)
        accelerationDictionary.Add("feet/second2", 0.3048)
        accelerationDictionary.Add("ft/s2", 0.3048)
        accelerationDictionary.Add("mile/hour/second", 0.44704)
        accelerationDictionary.Add("mi/hr/", 0.44704)
        accelerationDictionary.Add("gravity", 9.80665)
        accelerationDictionary.Add("g's", 9.80665)
        'power
        powerDictionary.Add("Watt", 1)
        powerDictionary.Add("watt", 1)
        powerDictionary.Add("W", 1)
        powerDictionary.Add("Joule/second", 1)
        powerDictionary.Add("J/s", 1)
        powerDictionary.Add("calorie/second", 4.1868)
        powerDictionary.Add("cal/", 4.1868)
        powerDictionary.Add("Calorie/second", 4186.8)
        powerDictionary.Add("Cal/", 4186.8)
        powerDictionary.Add("BTU/minute", 17.584264)
        powerDictionary.Add("BTU/min", 17.584264)
        powerDictionary.Add("horsepower", 745.699892)
        powerDictionary.Add("hp", 745.699892)
        'elements
        elementDictionary.Add("H", 1.00794)
        elementDictionary.Add("He", 4.002602)
        elementDictionary.Add("Li", 6.941)
        elementDictionary.Add("Be", 9.012182)
        elementDictionary.Add("B", 10.811)
        elementDictionary.Add("C", 12.0107)
        elementDictionary.Add("N", 14.0067)
        elementDictionary.Add("O", 15.9994)
        elementDictionary.Add("F", 18.9984032)
        elementDictionary.Add("Ne", 20.1797)
        elementDictionary.Add("Na", 22.98976928)
        elementDictionary.Add("Mg", 24.305)
        elementDictionary.Add("Al", 26.9815386)
        elementDictionary.Add("Si", 28.0855)
        elementDictionary.Add("P", 30.973762)
        elementDictionary.Add("S", 32.065)
        elementDictionary.Add("Cl", 35.453)
        elementDictionary.Add("Ar", 39.948)
        elementDictionary.Add("K", 39.0983)
        elementDictionary.Add("Ca", 40.078)
        elementDictionary.Add("Sc", 44.955912)
        elementDictionary.Add("Ti", 47.867)
        elementDictionary.Add("V", 50.9415)
        elementDictionary.Add("Cr", 51.9961)
        elementDictionary.Add("Mn", 54.938045)
        elementDictionary.Add("Fe", 55.845)
        elementDictionary.Add("Co", 58.933195)
        elementDictionary.Add("Ni", 58.6934)
        elementDictionary.Add("Cu", 63.546)
        elementDictionary.Add("Zn", 65.38)
        elementDictionary.Add("Ga", 69.723)
        elementDictionary.Add("Ge", 72.64)
        elementDictionary.Add("As", 74.9216)
        elementDictionary.Add("Se", 78.96)
        elementDictionary.Add("Br", 79.904)
        elementDictionary.Add("Kr", 83.798)
        elementDictionary.Add("Rb", 85.4678)
        elementDictionary.Add("Sr", 87.62)
        elementDictionary.Add("Y", 88.90585)
        elementDictionary.Add("Zr", 91.224)
        elementDictionary.Add("Nb", 92.90638)
        elementDictionary.Add("Mo", 95.96)
        elementDictionary.Add("Tc", 97.9072)
        elementDictionary.Add("Ru", 101.07)
        elementDictionary.Add("Rh", 102.9055)
        elementDictionary.Add("Pd", 106.42)
        elementDictionary.Add("Ag", 107.8682)
        elementDictionary.Add("Cd", 112.411)
        elementDictionary.Add("In", 114.818)
        elementDictionary.Add("Sn", 118.71)
        elementDictionary.Add("Sb", 121.76)
        elementDictionary.Add("Te", 127.6)
        elementDictionary.Add("I", 126.90447)
        elementDictionary.Add("Xe", 131.293)
        elementDictionary.Add("Cs", 132.9054519)
        elementDictionary.Add("Ba", 137.327)
        elementDictionary.Add("La", 138.90547)
        elementDictionary.Add("Ce", 140.116)
        elementDictionary.Add("Pr", 140.90765)
        elementDictionary.Add("Nd", 144.242)
        elementDictionary.Add("Pm", 145)
        elementDictionary.Add("Sm", 150.36)
        elementDictionary.Add("Eu", 151.964)
        elementDictionary.Add("Gd", 157.25)
        elementDictionary.Add("Tb", 158.92535)
        elementDictionary.Add("Dy", 162.5)
        elementDictionary.Add("Ho", 164.93032)
        elementDictionary.Add("Er", 167.259)
        elementDictionary.Add("Tm", 168.93421)
        elementDictionary.Add("Yb", 173.054)
        elementDictionary.Add("Lu", 174.9668)
        elementDictionary.Add("Hf", 178.49)
        elementDictionary.Add("Ta", 180.94788)
        elementDictionary.Add("W", 183.84)
        elementDictionary.Add("Re", 186.207)
        elementDictionary.Add("Os", 190.23)
        elementDictionary.Add("Ir", 192.217)
        elementDictionary.Add("Pt", 195.084)
        elementDictionary.Add("Au", 196.966569)
        elementDictionary.Add("Hg", 200.59)
        elementDictionary.Add("Tl", 204.3833)
        elementDictionary.Add("Pb", 207.2)
        elementDictionary.Add("Bi", 208.9804)
        elementDictionary.Add("Po", 208.9824)
        elementDictionary.Add("At", 209.9871)
        elementDictionary.Add("Rn", 222.0176)
        elementDictionary.Add("Fr", 223)
        elementDictionary.Add("Ra", 226)
        elementDictionary.Add("Ac", 227)
        elementDictionary.Add("Th", 232.03806)
        elementDictionary.Add("Pa", 231.03588)
        elementDictionary.Add("U", 238.02891)
        elementDictionary.Add("Np", 237)
        elementDictionary.Add("Pu", 244)
        elementDictionary.Add("Am", 243)
        elementDictionary.Add("Cm", 247)
        elementDictionary.Add("Bk", 247)
        elementDictionary.Add("Cf", 251)
        elementDictionary.Add("Es", 252)
        elementDictionary.Add("Fm", 257)
        elementDictionary.Add("Md", 258)
        elementDictionary.Add("No", 259)
        elementDictionary.Add("Lr", 262)
        elementDictionary.Add("Rf", 261)
        elementDictionary.Add("Db", 262)
        elementDictionary.Add("Sg", 266)
        elementDictionary.Add("Bh", 264)
        elementDictionary.Add("Hs", 277)
        elementDictionary.Add("Mt", 268)
        elementDictionary.Add("Ds", 271)
        elementDictionary.Add("Rg", 272)
        elementDictionary.Add("Cn", 285)
        elementDictionary.Add("Nh", 284)
        elementDictionary.Add("Fl", 289)
        elementDictionary.Add("Mc", 288)
        elementDictionary.Add("Lv", 292)
        elementDictionary.Add("Ts", 294)
        elementDictionary.Add("Og", 294)
    End Sub

    'converts everything to SI
    Function ConvertToMasterUnit(n As Double, unit As String)
        ConvertToMasterUnit = 0
    End Function

    'brings settings window to front
    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        palSettings.BringToFront()
        Me.Tag = "settings"
    End Sub

    'update default graph settings
    Private Sub TextBox1Op_TextChanged(sender As Object, e As EventArgs) Handles TextBox1Op.TextChanged
        If TabControl1Op.Tag = "yes" Then
            Dim checkNumber = eval(heal(TextBox1Op.Text), False, False)
            If IsNumeric(checkNumber) Then
                My.Settings.x_min = checkNumber
            Else
                My.Settings.x_min = 0
            End If
        End If
    End Sub

    'update default graph settings
    Private Sub TextBox2Op_TextChanged(sender As Object, e As EventArgs) Handles TextBox2Op.TextChanged
        If TabControl1Op.Tag = "yes" Then
            Dim checkNumber = eval(heal(TextBox2Op.Text), False, False)
            If IsNumeric(checkNumber) Then
                My.Settings.x_max = checkNumber
            Else
                My.Settings.x_max = 0
            End If
        End If
    End Sub

    'update default graph settings
    Private Sub TextBox3Op_TextChanged(sender As Object, e As EventArgs) Handles TextBox3Op.TextChanged
        If TabControl1Op.Tag = "yes" Then
            Dim checkNumber = eval(heal(TextBox3Op.Text), False, False)
            If IsNumeric(checkNumber) Then
                My.Settings.y_min = checkNumber
            Else
                My.Settings.y_min = 0
            End If
        End If
    End Sub

    'update default graph settings
    Private Sub TextBox4Op_TextChanged(sender As Object, e As EventArgs) Handles TextBox4Op.TextChanged
        If TabControl1Op.Tag = "yes" Then
            Dim checkNumber = eval(heal(TextBox4Op.Text), False, False)
            If IsNumeric(checkNumber) Then
                My.Settings.y_max = checkNumber
            Else
                My.Settings.y_max = 0
            End If
            My.Settings.y_max = eval(heal(TextBox4Op.Text), False, False)
        End If
    End Sub

    'update settings for autoscale
    Private Sub CheckBox1Op_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1Op.CheckedChanged
        If TabControl1Op.Tag = "yes" Then My.Settings.autofix = CheckBox1Op.Checked
    End Sub

    'update settings for fraction mode
    Private Sub CheckBox2Op_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2Op.CheckedChanged
        If TabControl1Op.Tag = "yes" Then
            My.Settings.fraction = CheckBox2Op.Checked
            FractionModeToolStripMenuItem.Checked = My.Settings.fraction
            shouldReset = True
            Button1.PerformClick()
            shouldReset = False
        End If
    End Sub

    'update settings for grouping numbers
    Private Sub CheckBox6Op_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6Op.CheckedChanged
        If TabControl1Op.Tag = "yes" Then
            My.Settings.grouping = CheckBox6Op.Checked
            Button1.PerformClick()
        End If
    End Sub

    'update settings for font
    Private Sub Button3Op_Click(sender As Object, e As EventArgs) Handles Button3Op.Click
        FontDialog1.Font = Button3Op.Font
        Dim result = FontDialog1.ShowDialog
        If result = Windows.Forms.DialogResult.OK Then
            Button3Op.Font = FontDialog1.Font
            RichTextBox1.Font = Button3Op.Font
            My.Settings.fontText = Button3Op.Font
            If My.Settings.colorText Then ColorTheText()
        End If
    End Sub

    'update settings for coloring the text
    Private Sub CheckBox3Op_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3Op.CheckedChanged
        If TabControl1Op.Tag = "yes" Then
            My.Settings.colorText = CheckBox3Op.Checked
            If My.Settings.colorText = True Then
                ColorTheText()
            Else
                RichTextBox1.SelectAll()
                If CheckBox4Op.Checked Then
                    RichTextBox1.SelectionColor = Color.Black
                Else
                    RichTextBox1.SelectionColor = Color.White
                End If
                RichTextBox1.SelectionStart = 0
            End If
        End If
    End Sub

    'update settings for projector mode
    Private Sub CheckBox4Op_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4Op.CheckedChanged
        If TabControl1Op.Tag = "yes" Then
            My.Settings.projector = CheckBox4Op.Checked
            If CheckBox4Op.Checked = True Then
                RichTextBox1.BackColor = Color.White
                RichTextBox1.ForeColor = Color.Black
            Else
                RichTextBox1.BackColor = Color.Black
                RichTextBox1.ForeColor = Color.White
            End If
            If CheckBox3Op.Checked = True Then
                ColorTheText()
            End If
        End If
    End Sub

    'update settings for window orientation
    Private Sub RadioButton1Op_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1Op.CheckedChanged, RadioButton2Op.CheckedChanged
        If TabControl1Op.Tag = "yes" Then
            If RadioButton1Op.Checked = True Then
                My.Settings.splitterOrientation = 1
                SplitContainer1.Orientation = 1
            Else
                My.Settings.splitterOrientation = 0
                SplitContainer1.Orientation = 0
            End If
        End If
    End Sub
    'display user defined functions panel
    Private Sub bpxBook_Click(sender As Object, e As EventArgs) Handles bpxBook.Click
        palFunctions.BringToFront()
        Me.Tag = "functions"
        TextBox10.Focus()
        TextBox10.SelectionLength = 0
    End Sub
    'save user defined functions to settings
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        My.Settings.userFunctions = TextBox10.Text
        Button3.Text = "Text saved!"
    End Sub

    Private Sub CheckBox10Op_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox10Op.CheckedChanged
        If TabControl1Op.Tag = "yes" Then
            My.Settings.highlighting = CheckBox10Op.Checked
        End If
    End Sub
    'change text back
    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged
        Button3.Text = "Save"
    End Sub
    'Draw 3D Lines
    Private Sub CheckBox9Op_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox9Op.CheckedChanged
        If TabControl1Op.Tag = "yes" Then My.Settings.draw3DLines = CheckBox9Op.Checked
    End Sub

    'update settings for save/autosave options
    Private Sub RadioButton3Op_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3Op.CheckedChanged, RadioButton4Op.CheckedChanged, RadioButton5Op.CheckedChanged
        If TabControl1Op.Tag = "yes" Then
            If RadioButton3Op.Checked = True Then
                My.Settings.saveOption = 1
            ElseIf RadioButton4Op.Checked = True Then
                My.Settings.saveOption = 2
            Else
                My.Settings.saveOption = 3
            End If
        End If
    End Sub

    'update settings for line thickness
    Private Sub TrackBar1Op_Scroll(sender As Object, e As EventArgs) Handles TrackBar1Op.Scroll
        If TabControl1Op.Tag = "yes" Then
            My.Settings.thickness = TrackBar1Op.Value
            functionPen = New Pen(Brushes.Blue, My.Settings.thickness)
            axisPen = New Pen(Brushes.Black, My.Settings.thickness)
            gridPen = New Pen(Brushes.Gray, My.Settings.thickness)
        End If
        Label9Op.Text = TrackBar1Op.Value
        PictureBox1Op.Refresh()
    End Sub

    Private Sub CopyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem.Click
        Clipboard.SetDataObject(DataXY.GetClipboardContent())
    End Sub

    ' toggles between fraction mode and regular mode
    Private Sub FractionModeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FractionModeToolStripMenuItem.Click
        CheckBox2Op.Checked = My.Settings.fraction Xor True
        If RichTextBox2.Text.StartsWith("A = ") Then RichTextBox2.Text = ""
        shouldReset = True
        Button1.PerformClick()
        shouldReset = False
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
        Me.Show()
    End Sub

    Private Sub FunctionsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FunctionsToolStripMenuItem.Click
        Functions.Show()
        Me.Hide()
    End Sub

    Private Sub ConverterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConverterToolStripMenuItem.Click
        Converter.Show()
        Me.Hide()
    End Sub

    'update settings for antialiased graphing
    Private Sub CheckBox5Op_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5Op.CheckedChanged
        If TabControl1Op.Tag = "yes" Then My.Settings.antialias = CheckBox5Op.Checked
    End Sub

    'update settings for enabling speech
    Private Sub CheckBox7Op_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7Op.CheckedChanged
        If TabControl1Op.Tag = "yes" Then
            My.Settings.speechEnabled = CheckBox7Op.Checked
            If My.Settings.speechEnabled Then
                RichTextBox2.BackColor = Color.PaleGreen
                reco.RecognizeAsync(System.Speech.Recognition.RecognizeMode.Multiple)
            Else
                RichTextBox2.BackColor = Color.White
                reco.RecognizeAsyncStop()
            End If
        End If
    End Sub

    'update settings for automatically calculating text
    Private Sub RadioButton6Op_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton6Op.CheckedChanged
        If TabControl1Op.Tag = "yes" Then My.Settings.autoCalculate = RadioButton6Op.Checked
    End Sub

    'update settings for automatically deleting the history box
    Private Sub CheckBox8Op_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8Op.CheckedChanged
        If TabControl1Op.Tag = "yes" Then My.Settings.autoDelete = CheckBox8Op.Checked
    End Sub

    'resets all the settings to 'factory default'
    Private Sub Button4Op_Click(sender As Object, e As EventArgs) Handles Button4Op.Click
        Dim answer = MsgBox("Choosing 'Yes' will reset Photon's settings to their original state.  Would you like to proceed with this operation?", vbYesNo)
        If answer = vbYes Then
            My.Settings.x_min = -4.1
            My.Settings.x_max = 4.1
            My.Settings.y_min = -4.1
            My.Settings.y_max = 4.1
            My.Settings.fraction = False
            My.Settings.autofix = True
            My.Settings.splitterDistance = 474
            My.Settings.degreeMode = False
            My.Settings.width = 688
            My.Settings.height = 466
            My.Settings.splitterOrientation = 1
            My.Settings.isMaximized = False
            My.Settings.left = 0
            My.Settings.top = 0
            My.Settings.saveOption = 3
            My.Settings.calcText = ""
            My.Settings.colorText = True
            My.Settings.fontText = New Font("Consolas", 12, FontStyle.Bold)
            My.Settings.thickness = 1
            My.Settings.projector = False
            My.Settings.antialias = True
            My.Settings.Amin = "-2.5"
            My.Settings.Amax = "2.5"
            My.Settings.grouping = True
            My.Settings.speechEnabled = False
            My.Settings.autoCalculate = True
            My.Settings.autoDelete = False
            My.Settings.drawMode = 0
            My.Settings.draw3DLines = False
            My.Settings.highlighting = True
            Application.Restart()
        End If
    End Sub

    ' closes the program
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub
    ' clears all the calculations, functions, or equations made in the history box
    Private Sub ClearCalculationsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearCalculationsToolStripMenuItem.Click, ClearFunctionsToolStripMenuItem.Click, ClearDataToolStripMenuItem.Click, ClearCommentsToolStripMenuItem.Click
        ClearText(sender.tag)
        If RichTextBox1.Text.Contains("A") = False Then
            TextBox2.Visible = False
            TextBox3.Visible = False
            TrackBar1.Visible = False
        End If
    End Sub
    ' clears all lines containing the given string
    Private Sub ClearText(s As String)
        Dim lines = RichTextBox1.Lines
        For i = 0 To lines.Count - 1
            If lines(i).Contains(s) Then lines(i) = ""
            If s = " = " Then
                If lines(i).StartsWith("trace") Then lines(i) = ""
            End If
        Next
        RichTextBox1.Lines = lines
        RichTextBox1.Lines = RichTextBox1.Text.Split(New Char() {ControlChars.Lf}, StringSplitOptions.RemoveEmptyEntries)
        If My.Settings.colorText Then ColorTheText()
        If s = " = " Then
            all3DFaces.Clear()
            facesAreColored = False
            last3DGraph = ""
            all3DPoints.Clear()
            all3DFunctionNames.Clear()
        End If
        PictureBox1.Refresh()
        RichTextBox2.Focus()
    End Sub
    ' opens another photon window
    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click
        Process.Start(Application.ExecutablePath)
    End Sub
    ' sets graph to draw a logarithmic x axis
    Private Sub LogXAxisToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LogXAxisToolStripMenuItem.Click
        PictureBox1.Refresh()
    End Sub
    ' sets graph to draw a logarithmic y axis
    Private Sub LogYAxisToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LogYAxisToolStripMenuItem.Click
        PictureBox1.Refresh()
    End Sub
    ' evaluate the text in the text box
    Private Sub TextBox2_LostFocus(sender As Object, e As EventArgs) Handles TextBox2.LostFocus, TextBox3.LostFocus
        sender.text = Replace(heal(sender.text), "A", "1")
    End Sub

    Private Sub TrackBar1_MouseLeave(sender As Object, e As EventArgs) Handles TrackBar1.MouseLeave
        RichTextBox2.Focus()
    End Sub
    ' plot animation
    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        Dim hA = eval(TextBox3.Text, False, False)
        Dim lA = eval(TextBox2.Text, False, False)
        If IsNumeric(hA) Then
            highA = hA
        Else
            highA = 0
        End If
        If IsNumeric(lA) Then
            lowA = lA
        Else
            lowA = 0
        End If
        midA = (TrackBar1.Value / TrackBar1.Maximum * (highA - lowA) + lowA)
        RichTextBox2.Text = "A = " + midA.ToString
        facesAreColored = False
        all3DFaces.Clear()
        last3DGraph = ""
        clear3DPathArrays()
        PictureBox1.Refresh()
    End Sub

    Private Sub clear3DPathArrays()
        all3DPoints.Clear()
        all3DFunctionNames.Clear()
        Dim point3DArray(35) As Point3D
        point3DArray(0) = New Point3D(0, 0, 0).RotateX(-90).RotateY(135).RotateX(20)
        point3DArray(1) = New Point3D(5, 0, 0).RotateX(-90).RotateY(135).RotateX(20)
        point3DArray(2) = New Point3D(0, 0, 0).RotateX(-90).RotateY(135).RotateX(20)
        point3DArray(3) = New Point3D(0, 5, 0).RotateX(-90).RotateY(135).RotateX(20)
        point3DArray(4) = New Point3D(0, 0, 0).RotateX(-90).RotateY(135).RotateX(20)
        point3DArray(5) = New Point3D(0, 0, 5 * zFactor).RotateX(-90).RotateY(135).RotateX(20)
        Dim spacing As Double = 0.15
        For i = 1 To 10 Step 2
            point3DArray(i + 5) = New Point3D(Int(i / 2) + 1, -spacing, 0).RotateX(-90).RotateY(135).RotateX(20)
            point3DArray(i + 6) = New Point3D(Int(i / 2) + 1, spacing, 0).RotateX(-90).RotateY(135).RotateX(20)
        Next i
        For i = 1 To 10 Step 2
            point3DArray(i + 15) = New Point3D(-spacing, Int(i / 2) + 1, 0).RotateX(-90).RotateY(135).RotateX(20)
            point3DArray(i + 16) = New Point3D(spacing, Int(i / 2) + 1, 0).RotateX(-90).RotateY(135).RotateX(20)
        Next i
        For i = 1 To 10 Step 2
            point3DArray(i + 25) = New Point3D(0, -spacing, (Int(i / 2) + 1) * zFactor).RotateX(-90).RotateY(135).RotateX(20)
            point3DArray(i + 26) = New Point3D(0, spacing, (Int(i / 2) + 1) * zFactor).RotateX(-90).RotateY(135).RotateX(20)
        Next i
        all3DPoints.Add(point3DArray)
    End Sub

    ' Opens the internet browswer to the Photon help webpage
    Private Sub WebHelpToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Process.Start("http://vbgadgets.weebly.com/photon-help.html")
    End Sub
    ' Adds a function to the textbox
    Private Sub SinesinToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SinesinToolStripMenuItem.Click, CosinecosToolStripMenuItem.Click, TangenttanToolStripMenuItem.Click, CosecantcscToolStripMenuItem.Click, SecantsecToolStripMenuItem.Click, CotangentcotToolStripMenuItem.Click, SquareRootSqrtToolStripMenuItem.Click, RoundRoundToolStripMenuItem.Click, RandomDecimalRndToolStripMenuItem.Click, NaturalLogLnToolStripMenuItem.Click, ModulusModToolStripMenuItem.Click, LogBaseTenLogToolStripMenuItem.Click, LogBaseBLogBaseToolStripMenuItem.Click, InverseTangentAtanToolStripMenuItem.Click, InverseSineAsinToolStripMenuItem.Click, InverseSecantAsecToolStripMenuItem.Click, InverseHyperbolicTangentAtanhToolStripMenuItem.Click, InverseHyperbolicSineAsinhToolStripMenuItem.Click, InverseHyperbolicSecantAsecToolStripMenuItem.Click, InverseHyperbolicCotangentAcothToolStripMenuItem.Click, InverseHyperbolicCosineAcoshToolStripMenuItem.Click, InverseHyperbolicCosecantAcschToolStripMenuItem.Click, InverseCotangentAcotToolStripMenuItem.Click, InverseCosineAcosToolStripMenuItem.Click, InverseCosecantAcscToolStripMenuItem.Click, IntegerPartIntToolStripMenuItem.Click, HyperbolicTangentTanhToolStripMenuItem.Click, HyperbolicSinesinhToolStripMenuItem.Click, HyperbolicSecantSechToolStripMenuItem.Click, HyperbolicCotangentCothToolStripMenuItem.Click, HyperbolicCosineCoshToolStripMenuItem.Click, HyperbolicCosecantCschToolStripMenuItem.Click, FixFixToolStripMenuItem.Click, EXExpToolStripMenuItem.Click, DotProductDotToolStripMenuItem.Click, CrossProductCrossToolStripMenuItem.Click, ConvertToVectorComponentFormToCompToolStripMenuItem.Click, ConvertToMagnitudeAngleToolStripMenuItem.Click, ConvertToBase8OctToolStripMenuItem.Click, ConvertToBase2BinToolStripMenuItem.Click, ConvertToBase16HexToolStripMenuItem.Click, ConvertTo3DVectorComponentFormToCompZToolStripMenuItem.Click, ConvertTo3DMagnitudeAngleFormToMagAngZToolStripMenuItem.Click, AbsoluteValueAbsToolStripMenuItem.Click, RadiansToDegreesToDegToolStripMenuItem.Click, DegreesToRadiansToRadToolStripMenuItem.Click, SubtractISubToolStripMenuItem.Click, MultiplyIMultToolStripMenuItem.Click, DivideIDivToolStripMenuItem.Click, AddISumToolStripMenuItem.Click, NthPrimePrimeToolStripMenuItem.Click, NthFibbonaciFibToolStripMenuItem.Click, FractionalPartFracToolStripMenuItem.Click, FactorialFactToolStripMenuItem.Click, PermutationsNPkToolStripMenuItem.Click, CombinationsNCkToolStripMenuItem.Click, SignSgnToolStripMenuItem.Click, SquareRootISqrtToolStripMenuItem.Click, InverseTangentXYAtanxyToolStripMenuItem.Click
        RichTextBox2.Focus()
        Dim position = RichTextBox2.SelectionStart
        RichTextBox2.Text = RichTextBox2.Text.Insert(position, sender.tag + ")")
        RichTextBox2.SelectionStart = position + sender.tag.length()
    End Sub
    ' Adds a help example to the text box
    Private Sub PlotAFunctionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PlotAFunctionToolStripMenuItem.Click, TakeSecondDerivativeToolStripMenuItem.Click, TakeFirstDerivativeToolStripMenuItem.Click, SpecifyTraceToolStripMenuItem.Click, SpecifyPlotDomainToolStripMenuItem.Click, PlotToolStripMenuItem.Click, PlotAPolarFunctionToolStripMenuItem.Click, PlotAParametricFunctionToolStripMenuItem.Click, EvaluateIntegralToolStripMenuItem.Click, AnimateTraceToolStripMenuItem.Click, AnimateAPlotToolStripMenuItem.Click, AnimateIntegralToolStripMenuItem.Click, CreateAVariableToolStripMenuItem.Click, PlotSlopeFieldToolStripMenuItem.Click, PlotImplicitEquationToolStripMenuItem.Click, Plot3DGraphToolStripMenuItem.Click, PlotA3DParametricFunctionToolStripMenuItem.Click
        RichTextBox2.Text = sender.tag
        RichTextBox2.SelectionStart = RichTextBox2.Text.Length
    End Sub
    ' Tells Photon to find the root, extrema, or inflection point of a function
    Private Sub RootToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RootToolStripMenuItem.Click, MinMaxToolStripMenuItem.Click, InflectionPointToolStripMenuItem.Click
        If RichTextBox1.Text.Contains(" = ") = False Then
            MsgBox("You must be plotting a function to find this.")
        Else
            Dim guess As String = InputBox("Enter a starting t - value as a guess:")
            If guess <> "" Then
                Dim res = eval(heal(guess), False, False)
                If res.ToString = "error" Then
                    MsgBox("You must enter a valid number as a starting guess.  Mathematical expressions are also valid.")
                    Exit Sub
                End If
                startingGuess = res
                findType = sender.tag
                PictureBox1.Refresh()
            End If
        End If
    End Sub
    ' Tells Photon that the user is permitted to use the mouse scroll wheel
    Private Sub PictureBox1_MouseEnter(sender As Object, e As EventArgs) Handles PictureBox1.MouseEnter
        shouldScroll = True
    End Sub
    ' Tells Photon that the user is not permitted to use the mouse scroll wheel
    Private Sub PictureBox1_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox1.MouseLeave
        shouldScroll = False
    End Sub
    ' zoom in and out
    Private Sub PictureBox1_MouseWheel(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseWheel, RichTextBox1.MouseWheel, RichTextBox2.MouseWheel, TextBox2.MouseWheel, TextBox3.MouseWheel, TrackBar1.MouseWheel, Button1.MouseWheel, Me.MouseWheel
        If all3DFaces.Count = 0 And all3DPoints.Count <= 1 Then
            If shouldScroll Then
                Dim tempXMin = xmin
                Dim tempXMax = xmax
                Dim tempYMin = ymin
                Dim tempYMax = ymax
                If PictureBox1.Cursor = Cursors.Cross Then 'zoom all
                    If e.Delta < 0 Then 'zoom out
                        xmin = tempXMin - (tempXMax - tempXMin) / 10
                        xmax = tempXMax + (tempXMax - tempXMin) / 10
                        ymin = tempYMin - (ymax - tempYMin) / 10
                        ymax = tempYMax + (ymax - tempYMin) / 10
                    Else 'zoom in
                        xmin = tempXMin + (tempXMax - tempXMin) / 12
                        xmax = tempXMax - (tempXMax - tempXMin) / 12
                        ymin = tempYMin + (tempYMax - tempYMin) / 12
                        ymax = tempYMax - (tempYMax - tempYMin) / 12
                    End If
                ElseIf PictureBox1.Cursor = Cursors.NoMoveHoriz Then 'zoom x
                    My.Settings.autofix = False
                    If e.Delta < 0 Then
                        xmin = tempXMin - (tempXMax - tempXMin) / 10
                        xmax = tempXMax + (tempXMax - tempXMin) / 10
                    Else
                        xmin = tempXMin + (tempXMax - tempXMin) / 12
                        xmax = tempXMax - (tempXMax - tempXMin) / 12
                    End If
                ElseIf PictureBox1.Cursor = Cursors.NoMoveVert Then 'zoom y
                    My.Settings.autofix = False
                    If e.Delta < 0 Then
                        ymin = tempYMin - (ymax - tempYMin) / 10
                        ymax = tempYMax + (ymax - tempYMin) / 10
                    Else
                        ymin = tempYMin + (tempYMax - tempYMin) / 12
                        ymax = tempYMax - (tempYMax - tempYMin) / 12
                    End If
                End If
                If xmin < -1.0E+307 Then xmin = -1.0E+307
                If xmax > 1.0E+307 Then xmax = 1.0E+307
                If ymin < -1.0E+307 Then ymin = -1.0E+307
                If ymax > 1.0E+307 Then ymax = 1.0E+307
                If xmax - xmin < 0.00000000000003 Then
                    xmin = tempXMin - (tempXMax - tempXMin) / 10
                    xmax = tempXMax + (tempXMax - tempXMin) / 10
                End If
                If tempYMax - tempYMin < 0.00000000000003 Then
                    ymin = tempYMin - (tempYMax - tempYMin) / 10
                    ymax = tempYMax + (tempYMax - tempYMin) / 10
                End If
                PictureBox1.Refresh()
            End If
        Else 'for 3d graphing
            If PictureBox1.Cursor = Cursors.NoMoveVert And all3DFaces.Count > 0 Then
                If e.Delta < 0 Then
                    zFactor /= 1.5
                Else
                    zFactor *= 1.5
                End If
                facesAreColored = False
                all3DFaces.Clear()
                last3DGraph = ""
                clear3DPathArrays()
            Else
                If e.Delta < 0 Then
                    viewdistance *= 1.15
                Else
                    viewdistance /= 1.15
                End If
            End If
            PictureBox1.Refresh()
        End If
    End Sub

    'provides the interpretation of speech recognition
    Private Sub reco_SpeechRecognized(ByVal sender As Object, ByVal e As System.Speech.Recognition.RecognitionEventArgs) Handles reco.SpeechRecognized
        If e.Result.Text = "clear" Then
            ClearAllToolStripMenuItem.PerformClick()
            Exit Sub
        ElseIf e.Result.Text = "go" Then
            Button1.PerformClick()
            Exit Sub
        ElseIf e.Result.Text = "enable fraction mode" Or e.Result.Text = "turn on fraction mode" Then
            If My.Settings.fraction = False Then FractionModeToolStripMenuItem.PerformClick()
            Exit Sub
        ElseIf e.Result.Text = "disable fraction mode" Or e.Result.Text = "turn off fraction mode" Then
            If My.Settings.fraction = True Then FractionModeToolStripMenuItem.PerformClick()
            Exit Sub
        ElseIf e.Result.Text.Contains("radian mode") Then
            RadianModeToolStripMenuItem.PerformClick()
            Exit Sub
        ElseIf e.Result.Text.Contains("degree mode") Then
            DegreeModeToolStripMenuItem.PerformClick()
            Exit Sub
        ElseIf e.Result.Text = "cancel" Then
            RichTextBox2.Text = ""
            variableCreate = False
            openParenth = 0
            Exit Sub
        ElseIf e.Result.Text = "deactivate" Then
            Button2.PerformClick()
        End If
        Dim allWords = e.Result.Text.Split()
        Dim txtResult As String = ""
        Dim numD As Double = 0
        Dim numS As String = ""
        Dim nAddMode As Boolean = True
        For i = 0 To allWords.Length - 1
            If i = 0 Then
                If openParenth > 0 Then
                    txtResult += ")"
                    openParenth -= 1
                End If
            End If
            If allWords(i) = "oh" Or allWords(i) = "zero" Or allWords(i) = "zeroeth" Then
                If nAddMode = False Then
                    numS += "0"
                End If
                If numS = "" Then numS = "0"
            ElseIf allWords(i) = "one" Or allWords(i) = "first" Then
                If nAddMode = True Then
                    numD += 1
                    numS = numD.ToString
                Else
                    numS += "1"
                End If
                nAddMode = False
            ElseIf allWords(i) = "two" Or allWords(i) = "second" Then
                If nAddMode = True Then
                    numD += 2
                    numS = numD.ToString
                Else
                    numS += "2"
                End If
                nAddMode = False
            ElseIf allWords(i) = "three" Then
                If nAddMode = True Then
                    numD += 3
                    numS = numD.ToString
                Else
                    numS += "3"
                End If
                nAddMode = False
            ElseIf allWords(i) = "four" Then
                If nAddMode = True Then
                    numD += 4
                    numS = numD.ToString
                Else
                    numS += "4"
                End If
                nAddMode = False
            ElseIf allWords(i) = "five" Then
                If nAddMode = True Then
                    numD += 5
                    numS = numD.ToString
                Else
                    numS += "5"
                End If
                nAddMode = False
            ElseIf allWords(i) = "six" Then
                If nAddMode = True Then
                    numD += 6
                    numS = numD.ToString
                Else
                    numS += "6"
                End If
                nAddMode = False
            ElseIf allWords(i) = "seven" Then
                If nAddMode = True Then
                    numD += 7
                    numS = numD.ToString
                Else
                    numS += "7"
                End If
                nAddMode = False
            ElseIf allWords(i) = "eight" Then
                If nAddMode = True Then
                    numD += 8
                    numS = numD.ToString
                Else
                    numS += "8"
                End If
                nAddMode = False
            ElseIf allWords(i) = "nine" Then
                If nAddMode = True Then
                    numD += 9
                    numS = numD.ToString
                Else
                    numS += "9"
                End If
                nAddMode = False
            ElseIf allWords(i) = "ten" Then
                If nAddMode = True Then
                    numD += 10
                    numS = numD.ToString
                Else
                    numS += "10"
                End If
                nAddMode = False
            ElseIf allWords(i) = "eleven" Then
                If nAddMode = True Then
                    numD += 11
                    numS = numD.ToString
                Else
                    numS += "11"
                End If
                nAddMode = False
            ElseIf allWords(i) = "twelve" Then
                If nAddMode = True Then
                    numD += 12
                    numS = numD.ToString
                Else
                    numS += "12"
                End If
                nAddMode = False
            ElseIf allWords(i) = "thirteen" Then
                If nAddMode = True Then
                    numD += 13
                    numS = numD.ToString
                Else
                    numS += "13"
                End If
                nAddMode = False
            ElseIf allWords(i) = "fourteen" Then
                If nAddMode = True Then
                    numD += 14
                    numS = numD.ToString
                Else
                    numS += "14"
                End If
                nAddMode = False
            ElseIf allWords(i) = "fifteen" Then
                If nAddMode = True Then
                    numD += 15
                    numS = numD.ToString
                Else
                    numS += "15"
                End If
                nAddMode = False
            ElseIf allWords(i) = "sixteen" Then
                If nAddMode = True Then
                    numD += 16
                    numS = numD.ToString
                Else
                    numS += "16"
                End If
                nAddMode = False
            ElseIf allWords(i) = "seventeen" Then
                If nAddMode = True Then
                    numD += 17
                    numS = numD.ToString
                Else
                    numS += "17"
                End If
                nAddMode = False
            ElseIf allWords(i) = "eighteen" Then
                If nAddMode = True Then
                    numD += 18
                    numS = numD.ToString
                Else
                    numS += "18"
                End If
                nAddMode = False
            ElseIf allWords(i) = "nineteen" Then
                If nAddMode = True Then
                    numD += 19
                    numS = numD.ToString
                Else
                    numS += "19"
                End If
                nAddMode = False
            ElseIf allWords(i) = "twenty" Then
                If nAddMode = True Then
                    numD += 20
                    numS = numD.ToString
                Else
                    numS += "20"
                    numD = Val(numS)
                End If
                nAddMode = True
            ElseIf allWords(i) = "thirty" Then
                If nAddMode = True Then
                    numD += 30
                    numS = numD.ToString
                Else
                    numS += "30"
                    numD = Val(numS)
                End If
                nAddMode = True
            ElseIf allWords(i) = "forty" Then
                If nAddMode = True Then
                    numD += 40
                    numS = numD.ToString
                Else
                    numS += "40"
                    numD = Val(numS)
                End If
                nAddMode = True
            ElseIf allWords(i) = "fifty" Then
                If nAddMode = True Then
                    numD += 50
                    numS = numD.ToString
                Else
                    numS += "50"
                    numD = Val(numS)
                End If
                nAddMode = True
            ElseIf allWords(i) = "sixty" Then
                If nAddMode = True Then
                    numD += 60
                    numS = numD.ToString
                Else
                    numS += "60"
                    numD = Val(numS)
                End If
                nAddMode = True
            ElseIf allWords(i) = "seventy" Then
                If nAddMode = True Then
                    numD += 70
                    numS = numD.ToString
                Else
                    numS += "70"
                    numD = Val(numS)
                End If
                nAddMode = True
            ElseIf allWords(i) = "eighty" Then
                If nAddMode = True Then
                    numD += 80
                    numS = numD.ToString
                Else
                    numS += "80"
                    numD = Val(numS)
                End If
                nAddMode = True
            ElseIf allWords(i) = "ninety" Then
                If nAddMode = True Then
                    numD += 90
                    numS = numD.ToString
                Else
                    numS += "90"
                    numD = Val(numS)
                End If
                nAddMode = True
            ElseIf allWords(i) = "hundred" Then
                If numD < 1000 Then
                    numD = Val(numS) * 100
                    numS = numD.ToString
                Else
                    numS = numD.ToString
                    numS = numS.Insert(numS.Length - 3, numS.Substring(numS.Length - 1))
                    numS = numS.Substring(0, numS.Length - 1)
                    numD = Val(numS)
                End If
                nAddMode = True
            ElseIf allWords(i) = "thousand" Then
                If numD < 1000 Then
                    numD = Val(numS) * 1000
                    numS = numD.ToString
                Else
                    numS = numD.ToString
                    numS = numS.Insert(numS.Length - 6, numS.Substring(numS.Length - 3))
                    numS = numS.Substring(0, numS.Length - 3)
                    numD = Val(numS)
                End If
                nAddMode = True
            ElseIf allWords(i) = "million" Then
                If numD < 1000 Then
                    numD = Val(numS) * 1000000
                    numS = numD.ToString
                Else
                    numS = numD.ToString
                    numS = numS.Insert(numS.Length - 9, numS.Substring(numS.Length - 3))
                    numS = numS.Substring(0, numS.Length - 3)
                    numD = Val(numS)
                End If
                nAddMode = True
            ElseIf allWords(i) = "billion" Then
                If numD < 1000 Then
                    numD = Val(numS) * 1000000000
                    numS = numD.ToString
                Else
                    numS = numD.ToString
                    numS = numS.Insert(numS.Length - 12, numS.Substring(numS.Length - 3))
                    numS = numS.Substring(0, numS.Length - 3)
                    numD = Val(numS)
                End If
                nAddMode = True
            ElseIf allWords(i) = "trillion" Then
                numD = Val(numS) * 1000000000000
                numS = numD.ToString
                nAddMode = True
            ElseIf allWords(i) = "point" Then
                numS = numS + "."
                nAddMode = False
            Else
                If numS <> "" Then
                    txtResult += numS
                    numS = ""
                    numD = 0
                End If
                If allWords(i) = "a" Then
                    txtResult += "a"
                ElseIf allWords(i) = "bee" Then
                    txtResult += "b"
                ElseIf allWords(i) = "see" Then
                    txtResult += "c"
                ElseIf allWords(i) = "dee" Then
                    txtResult += "d"
                ElseIf allWords(i) = "ee" Then
                    txtResult += "e"
                ElseIf allWords(i) = "eff" Then
                    txtResult += "f"
                ElseIf allWords(i) = "gee" Then
                    txtResult += "g"
                ElseIf allWords(i) = "aitch" Then
                    txtResult += "h"
                ElseIf allWords(i) = "i" Then
                    txtResult += "i"
                ElseIf allWords(i) = "jay" Then
                    txtResult += "j"
                ElseIf allWords(i) = "kay" Then
                    txtResult += "k"
                ElseIf allWords(i) = "el" Then
                    txtResult += "l"
                ElseIf allWords(i) = "em" Then
                    txtResult += "m"
                ElseIf allWords(i) = "en" Then
                    txtResult += "n"
                ElseIf allWords(i) = "pee" Then
                    txtResult += "p"
                ElseIf allWords(i) = "cue" Then
                    txtResult += "q"
                ElseIf allWords(i) = "ar" Then
                    txtResult += "r"
                ElseIf allWords(i) = "tee" Then
                    txtResult += "t"
                ElseIf allWords(i) = "yu" Then
                    txtResult += "u"
                ElseIf allWords(i) = "vee" Then
                    txtResult += "v"
                ElseIf allWords(i) = "double-yu" Then
                    txtResult += "w"
                ElseIf allWords(i) = "ex" Then
                    txtResult += "x"
                ElseIf allWords(i) = "wye" Then
                    txtResult += "y"
                ElseIf allWords(i) = "zee" Then
                    txtResult += "z"
                ElseIf allWords(i) = "pie" Then
                    txtResult += "pi"
                ElseIf allWords(i) = "plus" Then
                    If TextBox2.Text = "" And i = 0 Then
                        txtResult += "ans"
                    End If
                    txtResult += "+"
                ElseIf allWords(i) = "minus" Or allWords(i) = "take" Then
                    If TextBox2.Text = "" And i = 0 Then
                        txtResult += "ans"
                    End If
                    txtResult += "-"
                ElseIf allWords(i) = "times" Then
                    If TextBox2.Text = "" And i = 0 Then
                        txtResult += "ans"
                    End If
                    txtResult += "*"
                ElseIf allWords(i) = "divided" Or allWords(i) = "over" Then
                    If TextBox2.Text = "" And i = 0 Then
                        txtResult += "ans"
                    End If
                    txtResult += "/"
                ElseIf allWords(i) = "half" Or allWords(i) = "halves" Then
                    txtResult += "/2"
                ElseIf allWords(i) = "third" Then
                    txtResult += "/3"
                ElseIf allWords(i) = "fourth" Then
                    txtResult += "/4"
                ElseIf allWords(i) = "fifth" Then
                    txtResult += "/5"
                ElseIf allWords(i) = "sixth" Then
                    txtResult += "/6"
                ElseIf allWords(i) = "seventh" Then
                    txtResult += "/7"
                ElseIf allWords(i) = "eighth" Then
                    txtResult += "/8"
                ElseIf allWords(i) = "ninth" Then
                    txtResult += "/9"
                ElseIf allWords(i) = "tenth" Then
                    txtResult += "/10"
                ElseIf allWords(i) = "equals" Or allWords(i) = "equal" Or allWords(i) = "be" Then
                    If variableCreate = False Then
                        While openParenth > 0
                            txtResult += ")"
                            openParenth -= 1
                        End While
                        txtResult += "="
                    Else
                        txtResult += ":="
                    End If
                ElseIf allWords(i) = "to" Then
                    If TextBox2.Text = "" And i = 0 Then
                        txtResult += "ans"
                    End If
                    txtResult += "^"
                ElseIf allWords(i) = "squared" Then
                    If TextBox2.Text = "" And i = 0 Then
                        txtResult += "ans"
                    End If
                    txtResult += "^2"
                ElseIf allWords(i) = "cubed" Then
                    If TextBox2.Text = "" And i = 0 Then
                        txtResult += "ans"
                    End If
                    txtResult += "^3"
                ElseIf allWords(i) = "sine" Then
                    txtResult += "sin"
                ElseIf allWords(i) = "cosign" Then
                    txtResult += "cos"
                ElseIf allWords(i) = "tangent" Then
                    txtResult += "tan"
                ElseIf allWords(i) = "cosecant" Then
                    txtResult += "csc"
                ElseIf allWords(i) = "secant" Then
                    txtResult += "sec"
                ElseIf allWords(i) = "cotangent" Then
                    txtResult += "cot"
                ElseIf allWords(i) = "ay-sine" Then
                    txtResult += "asin"
                ElseIf allWords(i) = "ay-cosign" Then
                    txtResult += "acos"
                ElseIf allWords(i) = "ay-tangent" Then
                    txtResult += "atan"
                ElseIf allWords(i) = "ay-cosecant" Then
                    txtResult += "acsc"
                ElseIf allWords(i) = "ay-secant" Then
                    txtResult += "asec"
                ElseIf allWords(i) = "ay-cotangent" Then
                    txtResult += "acot"
                ElseIf allWords(i) = "absolutevalue" Then
                    txtResult += "abs"
                ElseIf allWords(i) = "of" Or allWords(i) = "quantity" Or allWords(i) = "open" Then
                    txtResult += "("
                    openParenth += 1
                ElseIf allWords(i) = "cloze" Then
                    txtResult += ")"
                    openParenth -= 1
                ElseIf allWords(i) = "all" Then
                    While openParenth > 0
                        txtResult += ")"
                        openParenth -= 1
                    End While
                    If RichTextBox1.Tag = "" Then
                        RichTextBox2.Text = "(" + RichTextBox2.Text + txtResult + ")"
                        txtResult = ""
                    Else
                        txtResult = "(" + txtResult + ")"
                    End If

                ElseIf allWords(i) = "root" Then
                    txtResult += "sqrt"
                ElseIf allWords(i) = "data" Then
                    txtResult += allWords(i)
                ElseIf allWords(i) = "let" Then
                    variableCreate = True
                ElseIf allWords(i) = "prime" Then
                    txtResult += "'"
                ElseIf allWords(i) = "deewhy" Then
                    txtResult += "dy"
                ElseIf allWords(i) = "deedeewhy" Then
                    txtResult += "ddy"
                ElseIf allWords(i) = "negative" Then
                    txtResult += "-"
                ElseIf allWords(i) = "log" Then
                    txtResult += "log"
                ElseIf allWords(i) = "naturalog" Then
                    txtResult += "ln"
                Else

                End If
            End If
        Next i
        If numS <> "" Then
            txtResult += numS
        End If
        txtResult = txtResult.Replace("^/", "^")
        txtResult = txtResult.Replace("theta", "th")
        txtResult = txtResult.Replace("//", "/")
        variableCreate = False
        If RichTextBox1.SelectionLength = 0 Then
            RichTextBox2.Text += txtResult
            RichTextBox2.SelectionStart = RichTextBox2.Text.Length
        Else
            While openParenth > 0
                txtResult += ")"
                openParenth -= 1
            End While
            RichTextBox1.SelectedText = txtResult
        End If
        If My.Settings.autoCalculate Then Button1.PerformClick()
    End Sub

    'allows editing for richtextbox1's text components, delineated with spaces
    Private Sub RichTextBox1_MouseMove(sender As Object, e As MouseEventArgs) Handles RichTextBox1.MouseMove
        If RichTextBox1.Text = "" Or My.Settings.speechEnabled = False Or RichTextBox1.Tag = "" Then Exit Sub
        Dim pos As Integer = RichTextBox1.GetCharIndexFromPosition(e.Location)
        If RichTextBox1.Text.Chars(pos) = " " Then Exit Sub
        Dim startPos As Integer
        For startPos = pos To 0 Step -1
            If RichTextBox1.Text.Chars(startPos) = " " Or RichTextBox1.Text.Chars(startPos) = vbNewLine Then Exit For
        Next startPos
        startPos += 1
        Dim endPos As Integer
        For endPos = pos To RichTextBox1.Text.Length - 1 Step 1
            If RichTextBox1.Text.Chars(endPos) = " " Or RichTextBox1.Text.Chars(endPos) = vbNewLine Then Exit For
        Next endPos
        RichTextBox1.Focus()
        RichTextBox1.SelectionStart = startPos
        If endPos - startPos >= 0 Then RichTextBox1.SelectionLength = endPos - startPos
    End Sub

    'sets focus to textbox1 is speech mode is on
    Private Sub TextBox1_MouseEnter(sender As Object, e As EventArgs) Handles RichTextBox2.MouseEnter
        If My.Settings.speechEnabled Then
            RichTextBox1.SelectionLength = 0
            RichTextBox2.Focus()
            RichTextBox2.SelectionStart = RichTextBox2.Text.Length
        End If
    End Sub

    'sets the tag for richtextbox1 so speech editing is allowed to be performed
    Private Sub RichTextBox1_MouseDown(sender As Object, e As MouseEventArgs) Handles RichTextBox1.MouseDown
        RichTextBox1.Tag = "1"
    End Sub

    'resets richtextbox1's tag so editing won't happen until the box is clicked again
    Private Sub RichTextBox1_LostFocus(sender As Object, e As EventArgs) Handles RichTextBox1.LostFocus
        Dim position As Integer = RichTextBox1.SelectionStart
        Dim selLength As Integer = RichTextBox1.SelectionLength
        RichTextBox1.Tag = ""
        RichTextBox1.SelectAll()
        RichTextBox1.SelectionBackColor = RichTextBox1.BackColor
        RichTextBox1.SelectionStart = position
        RichTextBox1.SelectionLength = selLength
    End Sub
    'remove yellow highlighting when losing focus of the main text entry box
    Private Sub RichTextBox2_LostFocus(sender As Object, e As EventArgs) Handles RichTextBox2.LostFocus
        Dim position As Integer = RichTextBox2.SelectionStart
        Dim selLength As Integer = RichTextBox2.SelectionLength
        RichTextBox2.SelectAll()
        RichTextBox2.SelectionBackColor = RichTextBox2.BackColor
        RichTextBox2.SelectionStart = position
        RichTextBox2.SelectionLength = selLength
    End Sub

    'updates x and y data
    Private Sub UpdateData()
        Dim dataX(DataXY.Rows.Count - 2) As Double
        Dim dataY(DataXY.Rows.Count - 2) As Double
        If DataXY.Rows.Count <> 1 Then
            Dim i = 0
            For i = 0 To DataXY.Rows.Count - 2
                Try
                    dataX(i) = DataXY.Rows(i).Cells(0).Value
                Catch ex As Exception
                    dataX(i) = 0
                End Try
                Try
                    dataY(i) = DataXY.Rows(i).Cells(1).Value
                Catch ex As Exception
                    dataY(i) = 0
                End Try
            Next i
            Array.Sort(dataX)
            Array.Sort(dataY)
            Dim n = dataX.Count
            Dim sumX As Double = 0
            Dim sumY As Double = 0
            Dim sumX2 As Double = 0
            Dim sumY2 As Double = 0
            Dim largestX As Double = Double.MinValue
            Dim smallestX As Double = Double.MaxValue
            Dim largestY As Double = Double.MinValue
            Dim smallestY As Double = Double.MaxValue
            For j = 0 To dataX.Count - 1
                sumX += dataX(j)
                sumY += dataY(j)
                sumX2 += dataX(j) ^ 2
                sumY2 += dataY(j) ^ 2
                largestX = Math.Max(largestX, dataX(j))
                smallestX = Math.Min(smallestX, dataX(j))
                largestY = Math.Max(largestY, dataY(j))
                smallestY = Math.Min(smallestY, dataY(j))
            Next j
            Dim meanX As Double = sumX / n
            Dim meanY As Double = sumY / n
            Dim sumDifSqX As Double = 0
            Dim sumDifSqY As Double = 0
            For j = 0 To dataX.Count - 1
                sumDifSqX += (dataX(j) - meanX) ^ 2
                sumDifSqY += (dataY(j) - meanY) ^ 2
            Next j
            Dim SstdevX As Double = Math.Sqrt(sumDifSqX / (n - 1))
            Dim SstdevY As Double = Math.Sqrt(sumDifSqY / (n - 1))
            Dim SigmaStdevX As Double = Math.Sqrt(sumDifSqX / n)
            Dim SigmaStdevY As Double = Math.Sqrt(sumDifSqY / n)
            XDataStats.Rows(0).Cells(1).Value = n
            YDataStats.Rows(0).Cells(1).Value = n
            XDataStats.Rows(1).Cells(1).Value = meanX
            YDataStats.Rows(1).Cells(1).Value = meanY
            XDataStats.Rows(2).Cells(1).Value = sumX
            YDataStats.Rows(2).Cells(1).Value = sumY
            XDataStats.Rows(3).Cells(1).Value = sumX2
            YDataStats.Rows(3).Cells(1).Value = sumY2
            XDataStats.Rows(4).Cells(1).Value = SstdevX
            YDataStats.Rows(4).Cells(1).Value = SstdevY
            XDataStats.Rows(5).Cells(1).Value = SigmaStdevX
            YDataStats.Rows(5).Cells(1).Value = SigmaStdevY
            XDataStats.Rows(6).Cells(1).Value = smallestX
            YDataStats.Rows(6).Cells(1).Value = smallestY
            If dataX.Count > 1 Then
                XDataStats.Rows(7).Cells(1).Value = getQuartile(dataX, 1)
                YDataStats.Rows(7).Cells(1).Value = getQuartile(dataY, 1)
                XDataStats.Rows(8).Cells(1).Value = getQuartile(dataX, 2)
                YDataStats.Rows(8).Cells(1).Value = getQuartile(dataY, 2)
                XDataStats.Rows(9).Cells(1).Value = getQuartile(dataX, 3)
                YDataStats.Rows(9).Cells(1).Value = getQuartile(dataY, 3)
            End If
            XDataStats.Rows(10).Cells(1).Value = largestX
            YDataStats.Rows(10).Cells(1).Value = largestY
        End If
    End Sub

    'updates x and y data
    Private Sub DataXY_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataXY.CellValueChanged
        changesOccured = True
        UpdateData()
    End Sub

    'make sure the sorting that takes place is numeric
    Private Sub DataXY_SortCompare(sender As Object, e As DataGridViewSortCompareEventArgs) Handles DataXY.SortCompare
        If e.Column.Index <> 0 Then
            Return
        End If
        Try
            e.SortResult = If(Val(e.CellValue1) < Val(e.CellValue2), -1, 1)
            e.Handled = True
        Catch
        End Try
    End Sub

    'for pasting data into the datagridview
    Private Sub DataXY_EditingControlShowing(sender As System.Object, e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles DataXY.EditingControlShowing
        e.Control.ContextMenuStrip = ContextMenuStrip2
        AddHandler e.Control.KeyDown, AddressOf cell_KeyDown
    End Sub

    Private Sub cell_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.V And Keys.ControlKey Then
            PasteData(DataXY)
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub PasteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PasteToolStripMenuItem.Click
        PasteData(DataXY)
    End Sub

    Public Sub PasteData(ByRef dgv As DataGridView)
        Dim tArr() As String
        Dim arT() As String
        Dim i, ii As Integer
        Dim c, cc, r As Integer
        tArr = Clipboard.GetText().Trim().Split(Environment.NewLine)
        r = 0 'dgv.CurrentCellAddress.Y()
        c = Math.Max(dgv.CurrentCellAddress.X(), 0)
        If (tArr.Length > dgv.Rows.Count - 1) Then dgv.Rows.Add(tArr.Length - dgv.Rows.Count + 1)
        For i = 0 To tArr.Length - 1
            If tArr(i) <> "" Then
                arT = tArr(i).Split(vbTab)
                cc = c
                For ii = 0 To arT.Length - 1
                    If cc > dgv.Columns.Count - 1 Then Exit For
                    If r > dgv.Rows.Count - 1 Then Exit Sub
                    dgv.Item(cc, r).Value = arT(ii).TrimStart
                    cc = cc + 1
                Next
                r = r + 1
            End If
        Next
        dgv.CurrentCell = dgv(c, r)
    End Sub

    Function getQuartile(theArray() As Double, quartileOption As Integer) As Double
        Dim q As Double = theArray.Count * 0.25 * quartileOption - 0.5
        If q Mod 1 = 0 Then
            getQuartile = theArray(q)
        Else
            getQuartile = (theArray(q - 0.5) + theArray(q + 0.5)) / 2
        End If
    End Function

    Private Function LinearRegression(datapoints As ArrayList, Optional isDrawn As Boolean = False) As String
        Dim result = ""
        Dim Sx As Double = 0, Sy As Double = 0
        Dim Stt As Double = 0, Sts As Double = 0
        For i = 0 To datapoints.Count - 1
            Sx += datapoints(i).X
            Sy += datapoints(i).Y
        Next i
        For i = 0 To datapoints.Count - 1
            Dim t As Double = datapoints(i).X - Sx / datapoints.Count
            Stt += t * t
            Sts += t * datapoints(i).Y
        Next i
        Dim m As Double = Sts / Stt
        Dim b As Double = (Sy - Sx * m) / datapoints.Count
        If Math.Abs(m) > 0.01 Then m = Math.Round(m, 6)
        If Math.Abs(b) > 0.01 Then b = Math.Round(b, 6)
        result = "y = " + m.ToString + "*x + " + b.ToString
        result = Replace(result, "+ -", "- ")
        LinearRegression = result
    End Function

    Function QuadraticRegression(datapoints As ArrayList, Optional isDrawn As Boolean = False)
        Dim result = ""
        Dim n = datapoints.Count
        Dim S01 As Double = 0, S11 As Double = 0, S21 As Double = 0
        Dim S10 As Double = 0, S20 As Double = 0, S30 As Double = 0, S40 As Double = 0
        For i = 0 To datapoints.Count - 1
            S01 += datapoints(i).Y
            S10 += datapoints(i).X
            S11 += datapoints(i).X * datapoints(i).Y
            S20 += datapoints(i).X ^ 2
            S21 += datapoints(i).X ^ 2 * datapoints(i).Y
            S30 += datapoints(i).X ^ 3
            S40 += datapoints(i).X ^ 4
        Next
        Dim a As Double = (S01 * S10 * S30 - S11 * n * S30 - S01 * S20 ^ 2 + S11 * S10 * S20 + S21 * n * S20 - S21 * S10 ^ 2) / (n * S20 * S40 - S10 ^ 2 * S40 - n * S30 ^ 2 + 2 * S10 * S20 * S30 - S20 ^ 3)
        Dim b As Double = (S11 * n * S40 - S01 * S10 * S40 + S01 * S20 * S30 - S21 * n * S30 - S11 * S20 ^ 2 + S21 * S10 * S20) / (n * S20 * S40 - S10 ^ 2 * S40 - n * S30 ^ 2 + 2 * S10 * S20 * S30 - S20 ^ 3)
        Dim c As Double = (S01 * S20 * S40 - S11 * S10 * S40 - S01 * S30 ^ 2 + S11 * S20 * S30 + S21 * S10 * S30 - S21 * S20 ^ 2) / (n * S20 * S40 - S10 ^ 2 * S40 - n * S30 ^ 2 + 2 * S10 * S20 * S30 - S20 ^ 3)
        If Math.Abs(a) > 0.01 Then a = Math.Round(a, 6)
        If Math.Abs(b) > 0.01 Then b = Math.Round(b, 6)
        If Math.Abs(c) > 0.01 Then c = Math.Round(c, 6)
        If My.Settings.fraction And isDrawn = True Then
            result = "y = " + decToFraction(a) + "*x^2 + " + decToFraction(b) + "*x + " + decToFraction(c)
        Else
            If Math.Abs(a) > 0.01 Then a = Math.Round(a, 6)
            If Math.Abs(b) > 0.01 Then b = Math.Round(b, 6)
            If Math.Abs(c) > 0.01 Then c = Math.Round(c, 6)
            result = "y = " + a.ToString + "*x^2 + " + b.ToString + "*x + " + c.ToString
        End If
        result = Replace(result, "+ -", "- ")
        QuadraticRegression = result
    End Function

    Function CubicRegression(datapoints As ArrayList, Optional isDrawn As Boolean = False)
        Dim result = ""
        Dim n = datapoints.Count
        Dim S01 As Double = 0, S11 As Double = 0, S21 As Double = 0, S31 As Double = 0
        Dim S10 As Double = 0, S20 As Double = 0, S30 As Double = 0, S40 As Double = 0, S50 As Double = 0, S60 As Double = 0
        For i = 0 To datapoints.Count - 1
            S01 += datapoints(i).Y
            S11 += datapoints(i).X * datapoints(i).Y
            S21 += datapoints(i).X ^ 2 * datapoints(i).Y
            S31 += datapoints(i).X ^ 3 * datapoints(i).Y
            S10 += datapoints(i).X
            S20 += datapoints(i).X ^ 2
            S30 += datapoints(i).X ^ 3
            S40 += datapoints(i).X ^ 4
            S50 += datapoints(i).X ^ 5
            S60 += datapoints(i).X ^ 6
        Next i
        Dim det As Double = det44(S60, S50, S40, S30, S50, S40, S30, S20, S40, S30, S20, S10, S30, S20, S10, n)
        Dim detA As Double = det44(S31, S50, S40, S30, S21, S40, S30, S20, S11, S30, S20, S10, S01, S20, S10, n)
        Dim detB As Double = det44(S60, S31, S40, S30, S50, S21, S30, S20, S40, S11, S20, S10, S30, S01, S10, n)
        Dim detC As Double = det44(S60, S50, S31, S30, S50, S40, S21, S20, S40, S30, S11, S10, S30, S20, S01, n)
        Dim detD As Double = det44(S60, S50, S40, S31, S50, S40, S30, S21, S40, S30, S20, S11, S30, S20, S10, S01)
        Dim a As Double = detA / det
        Dim b As Double = detB / det
        Dim c As Double = detC / det
        Dim d As Double = detD / det
        If Math.Abs(a) > 0.01 Then a = Math.Round(a, 6)
        If Math.Abs(b) > 0.01 Then b = Math.Round(b, 6)
        If Math.Abs(c) > 0.01 Then c = Math.Round(c, 6)
        If Math.Abs(d) > 0.01 Then d = Math.Round(d, 6)
        If My.Settings.fraction And isDrawn = True Then
            result = "y = " + decToFraction(a) + "*x^3 + " + decToFraction(b) + "*x^2 + " + decToFraction(c) + "*x + " + decToFraction(d)
        Else
            If Math.Abs(a) > 0.01 Then a = Math.Round(a, 6)
            If Math.Abs(b) > 0.01 Then b = Math.Round(b, 6)
            If Math.Abs(c) > 0.01 Then c = Math.Round(c, 6)
            If Math.Abs(d) > 0.01 Then d = Math.Round(d, 6)
            result = "y = " + a.ToString + "*x^3 + " + b.ToString + "*x^2 + " + c.ToString + "*x + " + d.ToString
        End If
        result = Replace(result, "+ -", "- ")
        CubicRegression = result
    End Function

    'Data fit
    Private Sub Button1Data_Click(sender As Object, e As EventArgs) Handles Button1Data.Click
        DataXY.CurrentCell = Nothing
        For i = DataXY.Rows.Count - 2 To 0 Step -1
            If DataXY.Rows(i).Cells(0).Value Is Nothing And DataXY.Rows(i).Cells(1).Value Is Nothing Then
                DataXY.Rows.RemoveAt(i)
            End If
        Next i
        Dim dataPoints As New ArrayList
        If DataXY.Rows.Count <> 1 Then
            Dim functionString = ""
            For i = 0 To DataXY.Rows.Count - 2
                If IsNumeric(DataXY.Rows(i).Cells(0).Value) And IsNumeric(DataXY.Rows(i).Cells(1).Value) Then
                    Dim pointX = DataXY.Rows(i).Cells(0).Value
                    Dim pointY = DataXY.Rows(i).Cells(1).Value
                    If IsNumeric(pointX) And IsNumeric(pointY) Then
                        dataPoints.Add(New PointF(pointX, pointY))
                    End If
                End If
            Next i
            Dim n As Integer = dataPoints.Count
            If rbtLinear.Checked And dataPoints.Count >= 2 Then
                functionString = LinearRegression(dataPoints)
            ElseIf rbtQuadratic.Checked And dataPoints.Count >= 3 Then
                functionString = QuadraticRegression(dataPoints)
            ElseIf rbtCubic.Checked And dataPoints.Count >= 4 Then
                functionString = CubicRegression(dataPoints)
            ElseIf rbtExponential.Checked And dataPoints.Count >= 2 Then
                Dim Sx As Double = 0, Sy As Double = 0
                Dim Stt As Double = 0, Sts As Double = 0
                For i = 0 To dataPoints.Count - 1
                    If dataPoints(i).Y <= 0 Then
                        MsgBox("You cannot have any y values less than or equal to 0 for exponential fits.")
                        Exit Sub
                    End If
                    Sx += dataPoints(i).X
                    Sy += Math.Log(dataPoints(i).Y)
                Next i
                For i = 0 To dataPoints.Count - 1
                    Dim t As Double = dataPoints(i).X - Sx / n
                    Stt += t * t
                    Sts += t * Math.Log(dataPoints(i).Y)
                Next i
                Dim m As Double = Sts / Stt
                Dim b As Double = Math.Exp((Sy - Sx * m) / n)
                If Math.Abs(m) > 0.01 Then m = Math.Round(m, 6)
                If Math.Abs(b) > 0.01 Then b = Math.Round(b, 6)
                functionString = "y = " + b.ToString + "*e^(" + m.ToString + "*x)"
            ElseIf rbtLogarithmic.Checked And dataPoints.Count >= 2 Then
                Dim Sx As Double = 0, Sy As Double = 0
                Dim Stt As Double = 0, Sts As Double = 0
                For i = 0 To dataPoints.Count - 1
                    If dataPoints(i).X <= 0 Then
                        MsgBox("You cannot have any x values less than or equal to 0 for logarithmic fits.")
                        Exit Sub
                    End If
                    Sx += Math.Log(dataPoints(i).X)
                    Sy += dataPoints(i).Y
                Next i
                For i = 0 To dataPoints.Count - 1
                    Dim t As Double = Math.Log(dataPoints(i).X) - Sx / n
                    Stt += t * t
                    Sts += t * dataPoints(i).Y
                Next i
                Dim m As Double = Sts / Stt
                Dim b As Double = (Sy - Sx * m) / n
                If Math.Abs(m) > 0.01 Then m = Math.Round(m, 6)
                If Math.Abs(b) > 0.01 Then b = Math.Round(b, 6)
                functionString = "y = " + m.ToString + "*ln(x) + " + b.ToString
                functionString = Replace(functionString, "+ -", "- ")
            ElseIf rbtGaussian.Checked And dataPoints.Count >= 2 Then
                Dim weightedSum As Double = 0
                Dim weightedN As Double = 0
                For i = 0 To dataPoints.Count - 1
                    weightedSum += dataPoints(i).Y ^ 2 * dataPoints(i).X
                    weightedN += dataPoints(i).Y ^ 2
                Next i
                Dim mu As Double = weightedSum / weightedN
                Dim secondSum As Double = 0
                For i = 0 To dataPoints.Count - 1
                    secondSum += dataPoints(i).Y ^ 2 * (dataPoints(i).X - mu) ^ 2
                Next i
                Dim sigma As Double = Math.Sqrt(secondSum / (weightedN))
                If Math.Abs(mu) > 0.01 Then mu = Math.Round(mu, 6)
                If Math.Abs(sigma) > 0.01 Then sigma = Math.Round(sigma, 6)
                Dim sumXY As Double = 0
                Dim sumXS As Double = 0
                For i = 0 To dataPoints.Count - 1
                    Dim tempX As Double = Math.Exp(-1 / 2 * ((dataPoints(i).X - mu) / sigma) ^ 2)
                    'Dim tempx As Double = Math.Exp(-1 * ((dataX(i) - mu) ^ 2 / (2 * sigma ^ 2)))
                    sumXY += dataPoints(i).Y * tempX
                    sumXS += tempX * tempX
                Next i
                Dim m As Double = sumXY / sumXS
                If Math.Abs(m) > 0.01 Then m = Math.Round(m, 6)
                'functionString = "y = " + m.ToString + "*e^(-1/2*((x - (" + mu.ToString + "))/" + sigma.ToString + ")^2)"
                functionString = "y = " + m.ToString + "*e^(-(x - " + mu.ToString + ")^2/(2*" + sigma.ToString + "^2))"
                functionString = Replace(functionString, "+ -", "- ")
            End If
            If RichTextBox1.Text = "" Then
                RichTextBox1.Text += functionString
            Else
                RichTextBox1.Text += vbNewLine + functionString
            End If
            changesOccured = True
            Button1.PerformClick()
        End If
    End Sub

    'Focus on this control when the user moves the mouse over it
    Private Sub XDataStats_MouseEnter(sender As Object, e As EventArgs) Handles XDataStats.MouseEnter, YDataStats.MouseEnter
        sender.Focus()
    End Sub

    'put focus back to xy data when the mouse leaves
    Private Sub YDataStats_MouseLeave(sender As Object, e As EventArgs) Handles YDataStats.MouseLeave, XDataStats.MouseLeave
        DataXY.Focus()
    End Sub

    'force numeric entries in the function table textboxes
    Private Sub txtStart_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtStart.KeyPress, txtDelta.KeyPress
        If Asc(e.KeyChar) <> 8 And e.KeyChar <> "." And e.KeyChar <> "-" Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    'use dragging to navigate table
    Private Sub DataFunction_MouseMove(sender As Object, e As MouseEventArgs) Handles DataFunction.MouseMove
        If e.Button = MouseButtons.Left Then
            While e.Y - DataFunction.Tag > 21
                txtStart.Text -= Val(txtDelta.Text)
                DataFunction.Tag += 21
            End While
            While e.Y - DataFunction.Tag < -21
                txtStart.Text += Val(txtDelta.Text)
                DataFunction.Tag -= 21
            End While
        Else
            DataFunction.Tag = e.Y
        End If
    End Sub

    'zoom in and out on table
    Private Sub DataFunction_MouseWheel(sender As Object, e As MouseEventArgs) Handles DataFunction.MouseWheel
        If e.Delta > 0 And Math.Abs(Val(txtDelta.Text)) > 0.000000000001 Then
            txtDelta.Text /= 10
            txtStart.Text /= 10
        ElseIf e.Delta < 0 Then
            txtDelta.Text *= 10
            txtStart.Text *= 10
        End If
    End Sub

    'focus the control
    Private Sub DataFunction_MouseEnter(sender As Object, e As EventArgs) Handles DataFunction.MouseEnter
        DataFunction.Focus()
    End Sub

    'refresh graphics for line thickness preview
    Private Sub PictureBox1Op_Paint(sender As Object, e As PaintEventArgs) Handles PictureBox1Op.Paint
        e.Graphics.DrawLine(New Pen(Brushes.Black, Val(Label9Op.Text)), New Point(0, PictureBox1Op.Height / 2), New Point(PictureBox1Op.Width, PictureBox1Op.Height / 2))
    End Sub

    ' Displays the about box
    Private Sub AboutPhotonToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutPhotonToolStripMenuItem.Click
        MsgBox("Scientific Calculator:" + vbTab + vbTab + vbTab + "Version " + My.Application.Info.Version.ToString + vbNewLine +
               "Last build date:" + vbTab + vbTab + "5/8/2017" + vbNewLine + vbNewLine + "Created for you" + vbTab +
               vbTab + "Andriy Ilnytskyy")
    End Sub
End Class

Public Class Point3D
    Protected m_x As Double, m_y As Double, m_z As Double

    Public Sub New(ByVal x As Double, ByVal y As Double, ByVal z As Double)
        Me.X = x
        Me.Y = y
        Me.Z = z
    End Sub

    Public Property X() As Double
        Get
            Return m_x
        End Get
        Set(ByVal value As Double)
            m_x = value
        End Set
    End Property

    Public Property Y() As Double
        Get
            Return m_y
        End Get
        Set(ByVal value As Double)
            m_y = value
        End Set
    End Property

    Public Property Z() As Double
        Get
            Return m_z
        End Get
        Set(ByVal value As Double)
            m_z = value
        End Set
    End Property

    Public Function RotateX(ByVal angle As Integer) As Point3D
        Dim rad As Double, cosa As Double, sina As Double, yn As Double, zn As Double

        rad = angle * Math.PI / 180
        cosa = Math.Cos(rad)
        sina = Math.Sin(rad)
        yn = Me.Y * cosa - Me.Z * sina
        zn = Me.Y * sina + Me.Z * cosa
        Return New Point3D(Me.X, yn, zn)
    End Function

    Public Function RotateY(ByVal angle As Integer) As Point3D
        Dim rad As Double, cosa As Double, sina As Double, Xn As Double, Zn As Double

        rad = angle * Math.PI / 180
        cosa = Math.Cos(rad)
        sina = Math.Sin(rad)
        Zn = Me.Z * cosa + Me.X * sina
        Xn = -Me.Z * sina + Me.X * cosa

        Return New Point3D(Xn, Me.Y, Zn)
    End Function

    Public Function Project(ByVal viewWidth, ByVal viewHeight, ByVal fov, ByVal viewDistance)
        Dim factor As Double, Xn As Double, Yn As Double
        factor = fov / (viewDistance - Me.Z)
        Xn = Me.X * factor + viewWidth / 2
        Yn = -Me.Y * factor + viewHeight / 2
        Return New Point3D(Xn, Yn, -Me.Z)
    End Function
End Class

Public Class Face3D
    Public p1 As Point3D
    Public p2 As Point3D
    Public p3 As Point3D
    Public p4 As Point3D
    Public avgC As Color
    Public colors() As Color

    Public Sub New(ByVal p_1 As Point3D, ByVal p_2 As Point3D, ByVal p_3 As Point3D, ByVal p_4 As Point3D)
        p1 = p_1
        p2 = p_2
        p3 = p_3
        p4 = p_4
    End Sub

    Public Sub New(ByVal p_1 As Point3D, ByVal p_2 As Point3D, ByVal p_3 As Point3D, ByVal p_4 As Point3D, ByRef color_s() As Color, ByVal avg_C As Color)
        p1 = p_1
        p2 = p_2
        p3 = p_3
        p4 = p_4
        avgC = avg_C
        colors = color_s
    End Sub
End Class