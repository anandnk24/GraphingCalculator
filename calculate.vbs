Private Sub CommandButton3_Click()
    Dim Counter1 As Integer 'Counter for loops'
    Dim Counter2 As Integer 'Counter for loops'
    Dim Counter3 As Integer 'Counter for loops'
    Dim Counter4 As Integer 'Counter for loops'
    Dim yvalue As Integer 'The y-value when The function has a discontinuity'
    Dim yvalue1 As Integer 'The y-value when The function has a discontinuity'
    Dim Lowerbound As Integer 'Leftbound of the graph as determined by user'
    Dim Upperbound As Integer 'Upperbound of the graph as determined by user'
    Dim r As Integer 'red color of line
    Dim g As Integer 'green color of line'
    Dim b As Integer 'blue Color of line'
    Dim defaultr As Integer 'Default amount of red to be used in the line'
    Dim defaultg As Integer 'Default amount of green to be used in the line'
    Dim defaultb As Integer 'Default amount of blue to be used in the line'
    Dim discontinuityMarker As Boolean
    Dim inflectionMarker As Boolean
    Lowerbound = TextBox1.Value
    Upperbound = TextBox2.Value

    'Formula = Application.InputBox("Enter the formula")
    Formula = TextBox24.Value
    
    If (Formula = "" Or Len(Formula) < 1) Then
        MsgBox ("Please enter a formula or expression to graph")
    
    Else
        Formula = "=" & Formula
    
    
        Counter1 = 0               'Starting the first counter'
        yvalue = 15
          yvalue1 = 20
        discontinuityMarker = False
        inflectionMarker = False
        Sheets("Sheet1").Delete
        Sheets.Add
        ActiveSheet.Name = "Sheet1"
        
        'Get the y-values'
        For Counter2 = Lowerbound To Upperbound Step 1
            Counter1 = Counter1 + 1
            ActiveSheet.Cells(Counter1, 1) = Counter2 'Set the x-values in column 1'
            ActiveSheet.Cells(Counter1, 2) = Replace(Formula, "x", Counter2) 'Find the corresponding y values for the x values in column1'
            ActiveSheet.Cells(Counter1, yvalue1) = ActiveSheet.Cells(Counter1, 2)
            
            If (IsError(ActiveSheet.Cells(Counter1, 2)) = True) Then
                ActiveSheet.Cells(Counter1, 2) = ""
                 ActiveSheet.Cells(Counter1, yvalue1) = 0
                discontinuityMarker = True
                Counter3 = 0
                
                For Counter4 = Lowerbound To Upperbound Step 1
                    Counter3 = Counter3 + 1
                    ActiveSheet.Cells(Counter3, yvalue) = Counter2
                   
                Next
                yvalue = yvalue + 1
            Else
             ActiveSheet.Cells(Counter1, 2) = ActiveSheet.Cells(Counter1, 2)
            End If
            
        Next
            'Calculate the first derivative'
            Counter1 = 1
            For Counter2 = Lowerbound To Upperbound - 1 Step 1
            Counter1 = Counter1 + 1
            
            ActiveSheet.Cells(Counter1 - 1, 3) = (ActiveSheet.Cells(Counter1, 2) - ActiveSheet.Cells(Counter1 - 1, 2)) / (ActiveSheet.Cells(Counter1, 1) - ActiveSheet.Cells(Counter1 - 1, 1))
            If (ActiveSheet.Cells(Counter1, 2) = "") Then
               'ActiveSheet.Cells(Counter1 - 1, 3) = ActiveSheet.Cells(Counter1 - 1, 2) / (ActiveSheet.Cells(Counter1, 1) - ActiveSheet.Cells(Counter1 - 1, 1))
               ActiveSheet.Cells(Counter1 - 1, 3) = ""
               ActiveSheet.Cells(Counter1 - 2, 3) = ""
            End If
            
            
            If (Counter2 = ActiveSheet.Cells(1, yvalue - 1) And discontinuityMarker) Then
                ActiveSheet.Cells(Counter1 - 1, 3) = ""
                
                End If
            Next
            'Calculate the second derivative'
                     
            Counter1 = 1
            For Counter2 = Lowerbound To Upperbound - 2 Step 1
            Counter1 = Counter1 + 1
            ActiveSheet.Cells(Counter1 - 1, 4) = (ActiveSheet.Cells(Counter1, 3) - ActiveSheet.Cells(Counter1 - 1, 3)) / (ActiveSheet.Cells(Counter1, 1) - ActiveSheet.Cells(Counter1 - 1, 1))
            If (ActiveSheet.Cells(Counter1, 3) = "") Then
                ActiveSheet.Cells(Counter1 - 1, 4) = ActiveSheet.Cells(Counter1 - 1, 3) / (ActiveSheet.Cells(Counter1, 1) - ActiveSheet.Cells(Counter1 - 1, 1))
                ActiveSheet.Cells(Counter1 - 1, 4) = ""
                ActiveSheet.Cells(Counter1 - 2, 4) = ""
            End If
          
           
        If (Counter2 = ActiveSheet.Cells(1, yvalue - 1) And discontinuityMarker) Then
                ActiveSheet.Cells(Counter1 - 1, 4) = ""
            Else
                If (ActiveSheet.Cells(Counter1 - 1, 4) = 0) Then
                    ActiveSheet.Cells(1, 6) = Counter1
                    inflectionMarker = True
                End If
            End If
        Next
        ActiveSheet.Cells(1, 5) = "=Max(B1:B" & Counter1 + 1 & ")"
        ActiveSheet.Cells(2, 5) = "=Min(B1:B" & Counter1 + 1 & ")"