Sub Ticker()

'Start For loop for all worksheets in the file
For Each ws In Worksheets
        Dim WorksheetName As String
        WorksheetName = ws.Name
        Sheets(ws.Name).Select

Dim DateMinOpen As Variant
Dim DateMaxClose As Variant
Dim i As Double
Dim x As Double

Columns("I:Q").Select
Selection.Clear

'Column headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Value"
    Cells(1, 15).Value = "Ticker"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"

'Initialize the ticker column
x = 2
i = 2

Cells(x, 9).Value = Cells(x, 1).Value

DateMinOpen = Cells(i, 3).Value


'Total Percent Change
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow


    If Cells(i, 1).Value = Cells(x, 9).Value Then


        TotalV = TotalV + Cells(i, 7).Value
        DateMaxClose = Cells(i, 6).Value

        Else
     
            Cells(x, 10).Value = DateMaxClose - DateMinOpen

                If DateMaxClose <= 0 Then
            
                    Cells(x, 11).Value = 0
                    
                Else
                    
                    Cells(x, 11).Value = (DateMaxClose / (DateMinOpen + 1E-28)) - 1
                        
                    

                End If
                
Cells(x, 11).Style = "Percent"
                       
'Formatting the Yearly Change
    If Cells(x, 10).Value >= 0 Then
                                
            Cells(x, 10).Interior.ColorIndex = 4
                                    
            Else
                                
                Cells(x, 10).Interior.ColorIndex = 3
                    
    End If
                

    Cells(x, 12).Value = TotalV

    DateMinOpen = Cells(i, 3).Value

    TotalV = Cells(i, 7).Value


    x = x + 1
    Cells(x, 9).Value = Cells(i, 1).Value

    End If

Next i


' Greatest Increase/Decrease/Total
Volume_Greatest_Decrease = 100000
        Ticker_Greatest_Decrease = 100000
        
        LastRow = Cells(Rows.Count, 9).End(xlUp).Row
        
        For x = 2 To LastRow
        
        
            If Cells(x, 11).Value > Volume_Greatest_Increase Then
            
                Ticker_Greatest_Increase = Cells(x, 9).Value
                Volume_Greatest_Increase = Cells(x, 11).Value
        
            End If
        
        
            If Cells(x, 11).Value < Volume_Greatest_Decrease Then
            
                Ticker_Greatest_Decrease = Cells(x, 9).Value
                Volume_Greatest_Decrease = Cells(x, 11).Value
        
            End If
        
        
            If Cells(x, 12).Value > Volume_Greatest_Total_Volume Then
            
                Ticker_Greatest_Total_Volume = Cells(x, 9).Value
                Volume_Greatest_Total_Volume = Cells(x, 12).Value
        
            End If
        
        Next x
        



Next ws

Columns("A:Q").EntireColumn.AutoFit

End Sub
