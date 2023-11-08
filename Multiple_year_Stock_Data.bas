Attribute VB_Name = "Module1"

Sub TickerChange()

'==========
'Declarations
'==========


Dim YearDifference As Double
Dim TickerName As String
Dim LastRow As Long
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim TickerVolume As Double
Dim WorksheetName As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim Summary_Table_Row As Long
Dim i As Long
Dim ws As Worksheet
Dim LastRowPercent As Long


For Each ws In Worksheets


'===========
' Headers
'===========

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

'===========
' Set initial variables
'===========

'Last Row Calculation
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' Set an initial variable for holding the total Ticker Volume
 TickerVolume = 0

' Keep track of the location for each Ticker Name in the summary table
  Summary_Table_Row = 2

'Establish Opening Price
OpenPrice = ws.Cells(2, 3).Value

'Establish Opening Volume
TickerVolume = 0

'===========
' Loop Through All the worksheets
'===========

  ' Loop through all Ticker Names
  For i = 2 To LastRow

    ' Check if we are still within the Ticker Name, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker name
      TickerName = ws.Cells(i, 1).Value
      
      ' Print the Ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = TickerName

      ' Add to the Ticker Volume Total
      TickerVolume = TickerVolume + ws.Cells(i, 7).Value

      ' Find Ticker Closing Value
      ClosePrice = ws.Cells(i, 6).Value
      
      ' Calculate Yearly Change
      YearlyChange = ClosePrice - OpenPrice
          
      ' Calculate Yearly Percentage
      PercentChange = (YearlyChange / OpenPrice)
      
      'Print YearlyChange
      ws.Range("J" & Summary_Table_Row).Value = YearlyChange
            
      ' Print Yearly Percentage
     ws.Range("K" & Summary_Table_Row).Value = PercentChange
     
      'Conditional Formating Red Negative Green Positive
          If ws.Range("J" & Summary_Table_Row).Value >= 0 Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            ElseIf ws.Range("J" & Summary_Table_Row).Value < 0 Then
            
         ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
         End If
             
     'Print Yearly Percentage as Percentage
      ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            
      ' Print the Ticker Volume Total to the Summary Table
     ws.Range("L" & Summary_Table_Row).Value = TickerVolume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Ticker Total
     TickerVolume = 0
         
     'Rest CloseValue
     CloseValue = 0
     
    'Reset Yearly Change = 0
    YearlyChange = 0
    
    'Reset Open Price
    OpenPrice = ws.Cells(i + 1, 3)

    ' If the cell immediately following a row if the ticker is the same
    Else

      ' Add to the Ticker Tickertotal
      TickerVolume = TickerVolume + ws.Cells(i, 7).Value

    End If

  Next i
  
  '===========
  ' Summary Table
  '===========
 
' Last Row Calculation
LastRowPercent = ws.Cells(Rows.Count, 11).End(xlUp).Row

'find max and min and print in table

For i = 2 To LastRowPercent
     
     'Calculate and print the maximum ticker percentage
       If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & LastRowPercent)) Then
             ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
             ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                        
    ' Calculate and print the minimum ticker percentage
             ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & LastRowPercent)) Then
                     ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                     ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                     
      ' Change format to percentage
         ws.Range("Q2:Q3").NumberFormat = "0.00%"
                     
    'Calculate and print the maximum Volume
             ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & LastRowPercent)) Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                 ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
                 
        End If
             
    Next i
  
  
Next ws

End Sub



