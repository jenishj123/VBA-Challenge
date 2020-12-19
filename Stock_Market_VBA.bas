Attribute VB_Name = "Module1"
Sub Stock_Market_VBA():

'Declaring Variables
Dim ticker As String
Dim numtickers As Integer
Dim yearlychange As Double
Dim openingprice As Double
Dim closingprice As Double
Dim numbertickers As Integer
Dim totalstockvolume As Double
Dim lastRow As Long

'Loop Function to loop over each worksheet
For Each ws In Worksheets

    'Activation of worksheets
    ws.Activate
    
    'Find the last row of each worksheet
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    'Header result columns for the worksheets
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    'Once declared, initiazling variables for each statement
    numbertickers = 0
    ticker = ""
    yearlychange = 0
    openingprice = 0
    perctangechange = 0
    totalstockvolume = 0
    
    'Loop Function to loop through the tickers, this will skip the header of ticker column
    For i = 2 To lastRow
    
    'Add the value of the ticker
    ticker = Cells(i, 1).Value
    
    'Opening price of the start at the start of the year will be added
    If openingprice = 0 Then
        openingprice = Cells(i, 3).Value
    End If
    
    'Run this function if a different function appears in the list. Then apply incrementation.
    If Cells(i + 1, 1).Value <> ticker Then
        numtickers = numtickers + 1
        Cells(numtickers + 1, 9) = ticker
        
    'Calculate the yearly change value for yearly change function
    yearlychange = closingprice - openingprice
    
    'Calculate and insert yearly change value in the result output in each worksheet.
    Cells(numtickers + 1, 10).Value = yearlychange
    
    'Change the shade color of each cell based on positive, negative, and 0 value for yearly change variable
    
   'If value is greater than 0 which refers to positive outcome, shade the cell color in green shade.
    If yearlychange > 0 Then
        Cells(numtickers + 1, 10).Interior.ColorIndex = 4
    
    'If value is less than 0 which refers to negative outcome, shade the cell color in red shade.
   ElseIf yearlychange < 0 Then
         Cells(numtickers + 1, 10).Interior.ColorIndex = 3
   
   'If the value is set to 0, then shade the cell color in yellow.
  'Cells(numtickers + 1, 11).Value = Formate(percentchange, "Percent")
   
   End If
   
   'Set the percent_change value as percent
   Cells(numtickers + 1, 11).Value = Format(percentchange, "Percent")
   
    'Calculate total stock volume in selected cells in worksheet
    Cells(numbertickers + 1, 12).Value = totalstockvolume
   
   'Calculate percent change value for ticker in worksheet
   If openingprice = 0 Then
         percentchange = 0
   Else
      
      percentchange = (yearlychange / openingprice)
   
   End If
   
    
    'Total stock volume will reset to 0 when the ticker is in different list
    totalstockvolume = 0
    
End If

Next i

Next ws

End Sub

    
        
