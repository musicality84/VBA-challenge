VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
 
 
 Sub Stocks()
    
    'List of all needed variables
    Dim Ticker As String
    Dim Total_Volume As Long
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_Volume As Long
    Dim Current_Highest As Double
    Dim Current_Lowest As Double
    Dim Current_Volume As Double
    
    'Sets up ability to loop through each worksheet
    For Each ws In Worksheets
    
         'Set up summary tables
         Dim Summary_Table_Row As Integer
         Summary_Table_Row = 2
         ws.Range("I1").Value = "Ticker"
         ws.Range("J1").Value = "Yearly Change"
         ws.Range("K1").Value = "Percent Change"
         ws.Range("L1").Value = "Total Volume"
         ws.Range("N2").Value = "Greatest Percent Increase"
         ws.Range("N3").Value = "Greatest Percent Decrease"
         ws.Range("N4").Value = "Greatest Total Volume"
         ws.Range("O1").Value = "Ticker"
         ws.Range("P1").Value = "Value"
         
         'Determine Last Row
         Dim LastRow As Long
         LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
         
         'Runs through the raw data
         For i = 2 To LastRow
             'Look ahead for Ticker change, then add values to the summary table
             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                 Ticker = ws.Cells(i, 1).Value
                 Total_Value = Total_Value + ws.Cells(i, 7).Value
                 ws.Range("I" & Summary_Table_Row).Value = Ticker
                 ws.Range("L" & Summary_Table_Row).Value = Total_Value
                 Total_Value = 0 'Resets Total Value for next i
                 Close_Price = ws.Cells(i, 6).Value
                 Yearly_Change = Close_Price - Open_Price
                 ws.Range("J" & Summary_Table_Row).Value = Yearly_Change 'Set Yearly Change
                 Percent_Change = (Yearly_Change / Open_Price)
                 ws.Range("K" & Summary_Table_Row).Value = Percent_Change 'Set Percent Change
                 Summary_Table_Row = Summary_Table_Row + 1 'Add new row to summary table
                     
             Else 'If Ticker Value remains the same
                 Total_Value = Total_Value + ws.Cells(i, 7).Value 'Adds current row to total volume
                     If (ws.Cells(i, 2) = ws.Cells(2, 2)) Then 'Records opening price for current Ticker for January 2
                         Open_Price = ws.Cells(i, 3).Value
                     Else
                     End If
                 
             End If
             
         Next i
         
         'Formatting
         Set Rng = ws.Range("J2", "J3001")
         Rng.FormatConditions.Delete
         Set Condition1 = Rng.FormatConditions.Add(xlCellValue, xlGreater, "0")
         Set Condition2 = Rng.FormatConditions.Add(xlCellValue, xlLess, "0")
         
         With Condition1
             .Interior.Color = vbGreen
         End With
         
         With Condition2
             .Interior.Color = vbRed
         End With
         
         ws.Range("K2", "K3001").FormatConditions.Delete
         ws.Range("K2", "K3001").NumberFormat = "0.00%"
         
         'Set variables at 0 for the following for loop
        Current_Highest = 0
        Current_Lowest = 0
        Current_Volume = 0
         
         For i = 2 To 3001
             If ws.Cells(i, 11).Value > Current_Highest Then 'Checks for highest percent
                 Current_Highest = ws.Cells(i, 11).Value
                 ws.Range("O2").Value = ws.Cells(i, 9).Value
                 ws.Range("P2").Value = Current_Highest
             End If
             If ws.Cells(i, 11).Value < Current_Lowest Then 'Checks for lowest percent
                 Current_Lowest = ws.Cells(i, 11).Value
                 ws.Range("O3").Value = ws.Cells(i, 9).Value
                 ws.Range("P3").Value = Current_Lowest
             End If
             If ws.Cells(i, 12).Value > Current_Volume Then 'Checks for highest total volume
                 Current_Volume = ws.Cells(i, 12).Value
                 ws.Range("O4").Value = ws.Cells(i, 9).Value
                 ws.Range("P4").Value = Current_Volume
             End If
         Next i
         
         'More formatting
         ws.Range("P2", "P3").NumberFormat = "0.00%"
    Next ws
    
End Sub
