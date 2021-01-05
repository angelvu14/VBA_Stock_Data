# VBA_Stock_Data
Sub stock_data()
 'Define Variables
 Dim ws As Worksheet
 Dim Ticker As String
 Dim Volume As Double
 
 Dim Year_Open As Double
 Dim Year_Close As Double
 Dim Yearly_Change As Double
 Dim Percent_Change As Double
 
 Dim Counter As LongLong
 Dim Change As Double
 Dim i As LongLong
 Dim j As LongLong
 Dim c As LongLong
 Dim b As LongLong
 
 For Each ws In Worksheets
 
 'Set row count for the worksheet
 Dim Lastrow As Variant
 'Loop through all row
  Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
 Dim Summary_Table_Row As Integer
 Summary_Table_Row = 2
 
 
 'set headers
 ws.Cells(1, 9).Value = "Ticker"
 ws.Cells(1, 10).Value = "Yearly Change"
 ws.Cells(1, 11).Value = "Percent Change"
 ws.Cells(1, 12).Value = "Total Stock Volume"
 
 ws.Cells(1, 16).Value = "Ticker"
 ws.Cells(1, 17).Value = "Value"
 
 'Set value variables
 j = 2
 Ticker = 2
 
For i = 2 To Lastrow
 
 'Time to loop
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
    
'State the Values
        Ticker = ws.Cells(i, 1).Value
        Volume = Volume + ws.Cells(i, 7).Value
        
        ws.Range("I" & Summary_Table_Row).Value = Ticker
        ws.Range("L" & Summary_Table_Row).Value = Volume
        
         'Variable Set
        Volume = 0
        
        Year_Close = ws.Cells(i, 6)
        
        If Year_Open = 0 Then
        Yearly_Change = 0
        Percent_Change = 0
        
        Else
        Yearly_Change = Year_Close - Year_Open
        Percent_Change = (Year_Close - Year_Open) / Year_Open
        
        
    End If
        
        
     ws.Cells(Summary_Table_Row, 10).Value = Yearly_Change
     ws.Cells(Summary_Table_Row, 11).Value = Percent_Change
     ws.Cells(Summary_Table_Row, 11).Style = "Percent"
     ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
     
     Summary_Table_Row = Summary_Table_Row + 1
    
    ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
    Year_Open = ws.Cells(i, 3)
    
    Else
       Volume = Volume + ws.Cells(i, 7).Value
    
End If

'End of the four loop above
Next i

'Format Columns Color
For j = 2 To Lastrow

If ws.Cells(j, 10).Value > 0 Then
    '4 is green
    ws.Cells(j, 10).Interior.ColorIndex = 4
Else
    '3 is red
    ws.Cells(j, 10).Interior.ColorIndex = 3
        
End If

Next j

'Challenge:

'Set Header
 ws.Cells(2, 15).Value = "Greatest % Increase"
 ws.Cells(3, 15).Value = "Greatest % Decrease"
 ws.Cells(4, 15).Value = "Greatest Total Volume"
 ws.Cells(1, 16).Value = "Ticker"
 ws.Cells(1, 17).Value = "Value"
 
 'Declare Variables
 Dim Greatest_Increase As Double
 Dim Greatest_Decrease As Double
 Dim Greatest_Volume As Double
 
 Greatest_Increase = 0
 Greatest_Decrease = 0
 Greatest_Volume = 0
 Max_Range = ws.Range("K2:K" & Summary_Table_Row)
 Greatest_Increase = Application.WorksheetFunction.Max(Max_Range)
 Debug.Print (Greatest_Increase)
 
 For b = 2 To Summary_Table_Row
 
    If ws.Cells(b, 11).Value = Greatest_Increase Then
    ws.Range("P" & r).Value = ws.Range("I" & r).Value
    ws.Cells("Q2").Style = "Percent"
    ws.Cells("Q2").NumberFormat = "0.00%"
    ws.Cells("P2").Value = ws.Cells(b, 9).Value
    
    End If
    
Next b

 For c = 2 To Lastrow
 
    If ws.Cells(c, 11).Value < Greatest_Decrease Then
    Greatest_Decrease = ws.Cells(c, 11).Value
    ws.Cells("Q3").Value = Greatest_Decrease
    ws.Cells("Q3").Style = "Percent"
    ws.Cells("Q3").NumberFormat = "0.00%"
    ws.Cells("P3").Value = ws.Cells(c, 9).Value
    
    End If

Next c

 For a = 2 To Lastrow
 
    If ws.Cells(a, 12).Value > Greastest_Volume Then
    Greatest_Volume = ws.Cells(a, 12).Value
    ws.Cells("Q4").Value -Greatest_Volume
    ws.Cells("P4").Value -ws.Cells(a, 9).Value
    
    End If

'Autofit Table Columns
ws.Columns("A:Q").AutoFit

Next a

Next ws

End Sub



