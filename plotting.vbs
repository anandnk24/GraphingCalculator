'Plot the function entered by the user'
            Range("B1:B" & Counter1).Select
            Application.CutCopyMode = False
            Charts.Add
            ActiveChart.ChartType = xlXYScatterSmoothNoMarkers
            ActiveChart.SetSourceData Source:=Sheets("Sheet1").Range("B1:B" & Counter1 + 1)
            ActiveChart.SeriesCollection(1).XValues = Sheets("Sheet1").Range("A1:A" & Counter1 + 1)
            ActiveChart.Location Where:=xlLocationAsObject, Name:="Sheet1"
                        
            
            With ActiveChart
                .HasTitle = True
                .ChartTitle.Characters.Text = "Original Function"
                End With
                If (discontinuityMarker) Then
                     ActiveChart.SeriesCollection.NewSeries
                     ActiveChart.SeriesCollection(2).XValues = Sheets("Sheet1").Range("O1:O" & Counter1 + 1)
                     ActiveChart.SeriesCollection(2).Values = Sheets("Sheet1").Range("T1:T" & Counter1 + 1)
                End If
                'Marking if graph is convex or concave'
                defalutr = 255
                defalutg = 0
                defaultb = 0
                'r = InputBox("Enter how much red is to be used on the line of the convex sections of the graph using an integer from 0 to 255", "Input red", defaultr)
                'g = InputBox("Enter how much green is to be used on the line of the convex sections of the graph using an integer from 0 to 255", "Input green", defaultg)
                'b = InputBox("Enter how much blue is to be used on the line of the convex sections of the graph using an integer from 0 to 255", "Input blue", defaultb)
                r = TextBox3.Value
                g = TextBox4.Value
                b = TextBox5.Value
                
                Counter1 = 0
                For Counter2 = Lowerbound To Upperbound Step 1
                    Counter1 = Counter1 + 1
                    If (ActiveSheet.Cells(Counter1, 2) > 0) Then
                        ActiveChart.SeriesCollection(1).Select
                        ActiveChart.SeriesCollection(1).Points(Counter1).Select
                        With Selection
                            .MarkerStyle = xlMarkerStyleDiamond
                            .MarkerSize = 5
                            .MarkerBackgroundColor = RGB(r, g, b)
                        End With
                    Else
                        If (ActiveSheet.Cells(Counter1, 2) < 0 And ActiveSheet.Cells(Counter1, 2) <> "") Then
                            defaultr = 0
                            defaultg = 255
                            defaultb = 0
                            'r = InputBox("Enter how much red is to be used on the line of the concave sections of the graph using an integer from 0 to 255", "Input red", defaultr)
                            'g = InputBox("Enter how much green is to be used on the line of the concave sections of the graph using an integer from 0 to 255", "Input green", defaultg)
                            'b = InputBox("Enter how much blue is to be used on the line of the concave sections of the graph using an integer from 0 to 255", "Input blue", defaultb)
                            
                            r = TextBox23.Value
                            g = TextBox22.Value
                            b = TextBox21.Value
                
                            ActiveChart.SeriesCollection(1).Select
                            ActiveChart.SeriesCollection(1).Points(Counter1).Select
                                 With Selection
                                     .MarkerStyle = xlMarkerStyleDiamond
                                     .MarkerSize = 4
                                     .MarkerBackgroundColor = RGB(r, g, b)
                                End With
                         
                         Else
                             If (ActiveSheet.Cells(Counter1, 2) = 0 And ActiveSheet.Cells(Counter1, 2) <> "") Then
                                defaultr = 100
                                defaultg = 100
                                defaultb = 0
                                'r = InputBox("Enter how much red is to be used on the line of the straight sections of the graph using an integer from 0 to 255", "Input red", defaultr)
                                'g = InputBox("Enter how much green is to be used on the line of the straight sections of the graph using an integer from 0 to 255", "Input green", defaultg)
                                'b = InputBox("Enter how much blue is to be used on the line of the straight sections of the graph using an integer from 0 to 255", "Input blue", defaultb)
                                r = TextBox18.Value
                                g = TextBox19.Value
                                b = TextBox20.Value
                                ActiveChart.SeriesCollection(1).Select
                                ActiveChart.SeriesCollection(1).Points(Counter1).Select
                                    With Selection
                                        .MarkerStyle = xlMarkerStyleDiamond
                                        .MarkerSize = 4
                                        .MarkerBackgroundColor = RGB(r, g, b)
                                    End With
                            End If
                        End If
                    End If
                    
                Next
                'finding the maximum'
                Counter1 = 0
                defaultr = 0
                defaultg = 0
                defaultb = 255
                'r = InputBox("Enter how much red is to be used on the maimum of the graph using an integer from 0 to 255", "Input red", defaultr)
                'g = InputBox("Enter how much green is to be used on the maximum of the graph using an integer from 0 to 255", "Input green", defaultg)
                'b = InputBox("Enter how much blue is to be used on the maximum of the graph using an integer from 0 to 255", "Input blue", defaultb)
                r = TextBox14.Value
                g = TextBox13.Value
                b = TextBox12.Value
                               
                
                For Counter2 = Lowerbound To Upperbound Step 1
                    Counter1 = Counter1 + 1
                    If (ActiveSheet.Cells(Counter1, 2) = ActiveSheet.Cells(1, 5)) Then
                        ActiveChart.SeriesCollection(1).Select
                        ActiveChart.SeriesCollection(1).Points(Counter1).Select
                        With Selection
                            .MarkerStyle = xlMarkerStyleSquare
                            .MarkerSize = 8
                            .MarkerBackgroundColor = RGB(r, g, b)
                        End With
                    End If
                Next
                
                'Finding the minimum'
                Counter1 = 0
                defaultr = 250
                defaultg = 250
                defaultb = 0
               ' r = InputBox("Enter how much red is to be used on the minimum of the graph using an integer from 0 to 255", "Input red", defaultr)
               ' g = InputBox("Enter how much green is to be used on the minimum of the graph using an integer from 0 to 255", "Input green", defaultg)
               ' b = InputBox("Enter how much blue is to be used on the minimum of the graph using an integer from 0 to 255", "Input blue", defaultb)
                r = TextBox17.Value
                g = TextBox16.Value
                b = TextBox15.Value
                
                For Counter2 = Lowerbound To Upperbound Step 1
                    Counter1 = Counter1 + 1
                    If (ActiveSheet.Cells(Counter1, 2) = ActiveSheet.Cells(2, 5)) Then
                        ActiveChart.SeriesCollection(1).Select
                        ActiveChart.SeriesCollection(1).Points(Counter1).Select
                        With Selection
                            .MarkerStyle = xlMarkerStyleCircle
                            .MarkerSize = 8
                            .MarkerBackgroundColor = RGB(r, g, b)
                        End With
                    End If
                Next
                
                'Mark Point of inflection'
                defaultr = 100
                defaultg = 100
                defaultb = 100
               ' r = InputBox("Enter how much red is to be used on the inflection point(s) of the graph using an integer from 0 to 255", "Input red", defaultr)
               ' g = InputBox("Enter how much green is to be used on the inflection point(s) of the graph using an integer from 0 to 255", "Input green", defaultg)
               ' b = InputBox("Enter how much blue is to be used on the inflection point(s) of the graph using an integer from 0 to 255", "Input blue", defaultb)
              r = TextBox11.Value
                g = TextBox10.Value
                b = TextBox9.Value
                If (ActiveSheet.Cells(1, 6) >= 0 And inflectionMarker) Then
                    ActiveChart.SeriesCollection(1).Select
                    ActiveChart.SeriesCollection(1).Points(ActiveSheet.Cells(1, 6)).Select
                    With Selection
                        .MarkerStyle = xlMarkerStyleDiamond
                        .MarkerSize = 10
                        .MarkerBackgroundColor = RGB(r, g, b)
                        
                    End With
                End If
                
                'Plot the First derivative'
                Range("C1:C" & Counter1).Select
                Charts.Add
                ActiveChart.ChartType = xlXYScatterSmoothNoMarkers
                ActiveChart.SetSourceData Source:=Sheets("Sheet1").Range("C1:C" & Counter1)
                ActiveChart.SeriesCollection(1).XValues = Sheets("Sheet1").Range("A1:A" & Counter1)
                ActiveChart.Location Where:=xlLocationAsObject, Name:="Sheet1"
                
                With ActiveChart
                    .HasTitle = True
                    .ChartTitle.Characters.Text = "First Derivative"
                    .Parent.Left = 325
                    .Parent.Width = 300
                    .Parent.Top = 265
                    .Parent.Height = 355
                End With
                ' Second derivative
                
                Range("D1:D" & Counter1).Select
                Charts.Add
                ActiveChart.ChartType = xlXYScatterSmoothNoMarkers
                ActiveChart.SetSourceData Source:=Sheets("Sheet1").Range("D1:D" & Counter1)
                ActiveChart.SeriesCollection(1).XValues = Sheets("Sheet1").Range("A2:A" & Counter1)
                ActiveChart.Location Where:=xlLocationAsObject, Name:="Sheet1"
                
                With ActiveChart
                    .HasTitle = True
                    .ChartTitle.Characters.Text = "Second Derivative"
                    .Parent.Left = 250
                    .Parent.Width = 375
                    .Parent.Top = 700
                    .Parent.Height = 225
                End With
                    ActiveWindow.Visible = False
    End If
                
End Sub


Private Sub TextBox10_Change()

End Sub

Private Sub TextBox24_Change()

End Sub

Private Sub UserForm_Click()

End Sub