Attribute VB_Name = "Module1"
Sub challenge_2()
    
Dim ticker As String
Dim start As Double
Dim final As Double
Dim yearly As Double
Dim row As Long
Dim row2 As Integer
Dim finish As Boolean
Dim increase As Integer
Dim decrease As Integer
Dim volume As Integer
Dim x As Integer
    
For x = 1 To 3

Worksheets(x).Activate

    row = 2
    row2 = 2
    ticker = Cells(row, 1)
    start = Cells(row, 3)
    increase = 2
    decrease = 2
    volume = 2
    Cells(row2, 12) = 0
    finish = False
    
    While finish = False
        
        If Cells(row, 1) = 0 Then
            finish = True
            Cells(row2, 9) = ticker
                'Yearly Change
            final = Cells(row - 1, 6)
            yearly = final - start
            Cells(row2, 10) = yearly
                
            If yearly > 0 Then
                Range("J" & row2).Interior.ColorIndex = 4
            ElseIf yearly < 0 Then
                Range("J" & row2).Interior.ColorIndex = 3
            Else
                Range("J" & row2).Interior.ColorIndex = 2
            End If
                
                'Percent Chance
            Cells(row2, 11) = yearly / start
            Range("K" & row2).NumberFormat = "0.00%"
            start = Cells(row, 3)
        ElseIf Cells(row, 1) = ticker Then
            Cells(row2, 12) = Cells(row2, 12) + Cells(row, 7)
            row = row + 1
        Else
                'ticker
            Cells(row2, 9) = ticker
            ticker = Cells(row, 1)
                'Yearly Change
            final = Cells(row - 1, 6)
            yearly = final - start
            Cells(row2, 10) = yearly
                
            If yearly > 0 Then
                Range("J" & row2).Interior.ColorIndex = 4
            ElseIf yearly < 0 Then
                Range("J" & row2).Interior.ColorIndex = 3
            Else
                Range("J" & row2).Interior.ColorIndex = 2
            End If
                
                'Percent Chance
            Cells(row2, 11) = yearly / start
            Range("K" & row2).NumberFormat = "0.00%"
            start = Cells(row, 3)
            
            
             'increase
            If Cells(increase, 11) < Cells(row2, 11) Then
                increase = row2
            End If
        
            'decrease
            If Cells(decrease, 11) > Cells(row2, 11) Then
                decrease = row2
            End If
        
            'total volume
            If Cells(volume, 12) < Cells(row2, 12) Then
            volume = row2
            End If
                
            row2 = row2 + 1
            'volume
            Cells(row2, 12) = Cells(row, 7)
            row = row + 1
            
        End If

    Wend

        
        Cells(2, 16) = Cells(increase, 9)
        Cells(2, 17) = Cells(increase, 11)
        Cells(3, 16) = Cells(decrease, 9)
        Cells(3, 17) = Cells(decrease, 11)
        Cells(4, 16) = Cells(decrease, 9)
        Cells(4, 17) = Cells(volume, 12)
        Range("Q2").NumberFormat = "0.00%"
        Range("Q3").NumberFormat = "0.00%"
        Range("Q4").NumberFormat = "##0.0E+0"
        
Next x

End Sub
