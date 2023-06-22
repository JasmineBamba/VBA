Attribute VB_Name = "Module1"
Sub stockAnalysis()

'Declaring variables
Dim ws As Worksheet

Dim ticker As String
Dim yearlyChange As Double
Dim percentageChange As Double
Dim stockVolume As Double

Dim openingPrice As Double
Dim closingPrice As Double

Dim output As Long

Dim greatestIncreaseTicker As String
Dim greatestDecreaseTicker As String
Dim greatestStockVolumeTicker As String

Dim greatestIncreaseValue As Double
Dim greatestDecreaseValue As Double
Dim greatestStockVolumeValue As Double



'Loop through worksheets
For Each ws In ThisWorkbook.Worksheets

greatestIncreaseValue = 0
greatestDecreaseValue = 0
greatestStockVolumeValue = 0

greatestIncreaseTicker = ""
greatestDecreaseTicker = ""
greatestStockVolumeTicker = ""

'Worksheet where the data is stored
'Set ws = ThisWorkbook.Worksheets("A")


'Find the last used row in worksheet
lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row


'Assigning value to variable for output data that will start from row 2
output = 2


'Headers for output
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percentage Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"


'Auto-fit the width of columns to its content
ws.Columns("I:L").AutoFit
ws.Columns("O:Q").AutoFit


    'Starting from 2, as header being first row and looping till the last used row
    For i = 2 To lastRow

    
         If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
                 ticker = ws.Cells(i, 1).Value
            
                 openingPrice = ws.Cells(i, 3).Value
                 closingPrice = ws.Cells(i, 6).Value
            
                 'Calculate yearly change, percent change and total stock volume
                 yearlyChange = closingPrice - openingPrice
                 percentageChange = yearlyChange / openingPrice
                 stockVolume = stockVolume + ws.Cells(i, 7).Value
            
                 'Storing the output
                 ws.Cells(output, 9).Value = ticker
                 ws.Cells(output, 10).Value = yearlyChange
                 ws.Cells(output, 11).Value = percentageChange
                 ws.Cells(output, 11).NumberFormat = "0.00%"
                 ws.Cells(output, 12).Value = stockVolume
                 
                 
                                  
                 'Conditional formatting for yearly change
                 If yearlyChange >= 0 Then
                    ws.Cells(output, 10).Interior.Color = RGB(0, 255, 0) 'Green cell for positive values
            
                 Elseif yearlyChange < 0 Then
                    ws.Cells(output, 10).Interior.Color = RGB(255, 0, 0) 'Red cell for negative values
                 
                 End If
       
                        
                 'Conditional formatting for percentage change
                 If percentageChange >= 0 Then
                    ws.Cells(output, 11).Interior.Color = RGB(0, 255, 0) 'Green cell for positive values
                 
                 Elseif percentageChange < 0 Then
                    ws.Cells(output, 11).Interior.Color = RGB(255, 0, 0) 'Red cell for negative values
                 
                 End If
                 
               
                'Greastest % Increase
                If percentageChange > greatestIncreaseValue Then
                    
                    greatestIncreaseTicker = ticker
                    greatestIncreaseValue = percentageChange
                    
              

                'Greastest % Decrease
                ElseIf percentageChange < greatestDecreaseValue Then
                    
                    greatestDecreaseTicker = ticker
                    greatestDecreaseValue = percentageChange
                    
         
                'Greatest Total Volume
                ElseIf stockVolume > greatestStockVolumeValue Then
                
                    greatestStockVolumeTicker = ticker
                    greatestStockVolumeValue = stockVolume
                    
                End If
               
              output = output + 1
              
          End If
       
    
    Next i
               
               ws.Cells(2, 16).Value = greatestIncreaseTicker
               ws.Cells(3, 16).Value = greatestDecreaseTicker
               ws.Cells(4, 16).Value = greatestStockVolumeTicker
            
               ws.Cells(2, 17).Value = greatestIncreaseValue
               ws.Cells(2, 17).NumberFormat = "0.00%"
               ws.Cells(3, 17).Value = greatestDecreaseValue
               ws.Cells(3, 17).NumberFormat = "0.00%"
               ws.Cells(4, 17).Value = greatestStockVolumeValue
             
               
    Next ws

End Sub

