Attribute VB_Name = "Module1"
Sub Stock()
Dim Ticker As String
Dim Stock_Total As Double
Dim Summary_Table_Row As Integer
Dim ws As Worksheet
Dim Opening As Double
Dim Closing As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Price_Index As Long

'Print header rows
For Each ws In Worksheets
 ws.Range("I1").Value = "Ticker"
 ws.Range("J1").Value = "Yearly Change"
 ws.Range("K1").Value = "Percent Change"
 ws.Range("L1").Value = "Total Stock Volume”"


Opening = ws.Cells(2, 3).Value
Yearly_Change = 0
Summary_Table_Row = 2
Stock_Total = 0
Price_Index = 2
Greatest_Percent_Increase = 0
Greatest_Percent_Increase_Ticker = " "
Greatest_Percent_Decrease = 0
Greatest_Percent_Decrease_Ticker = " "
Greatest_Stock_Total = 0
Greatest_Stock_Toal_Ticker = " "
'setting headers for summary table



lastLine = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastLine
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            Ticker = ws.Cells(i, 1).Value
            Stock_Total = Stock_Total + ws.Cells(i, 7).Value
            
            'Print ouput
            ws.Range("I" & Summary_Table_Row).Value = Ticker
            ws.Range("L" & Summary_Table_Row).Value = Stock_Total
          
                If Stock_Total > Greatest_Stock_Total Then
                Greatest_Stock_Total = Stock_Total
                Greatest_Stock_Total_Ticker = Ticker
                End If
            
              Opening = ws.Range("C" & Price_Index).Value
              
                Closing = ws.Range("F" & i).Value
               Yearly_Change = Closing - Opening

    
            
            If (Opening = 0) Then
                Percent_Change = Yearly_Change
            Else
                Percent_Change = Yearly_Change / Opening
                
                
            End If
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                If Percent_Change > Greatest_Percent_Increase Then
                Greatest_Percent_Increase = Percent_Change
                Greatest_Percent_Increase_Ticker = Ticker
            End If
                If Percent_Change < Greatest_Percent_Decrease Then
                Greatest_Percent_Decrease = Percent_Change
                Greatest_Percent_Decrease_Ticker = Ticker
            End If
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                
            If Yearly_Change < 0 Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            
            Else
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            
            End If
            
            'Reset variables
            Summary_Table_Row = Summary_Table_Row + 1
            Stock_Total = 0
            'Opening = ws.Cells(i + 1, 3).Value
            Price_Index = i + 1
        Else
            Stock_Total = Stock_Total + ws.Cells(i, 7).Value
        
        
        End If
    Next i
    ws.Range("Q2").Value = Greatest_Percent_Increase_Ticker
    ws.Range("R2").Value = Greatest_Percent_Increase
    ws.Range("R2").NumberFormat = "0.00%"
    ws.Range("Q3").Value = Greatest_Percent_Decrease_Ticker
    ws.Range("R3").Value = Greatest_Percent_Decrease
    ws.Range("R3").NumberFormat = "0.00%"
    ws.Range("Q4").Value = Greatest_Stock_Total_Ticker
    ws.Range("R4").Value = Greatest_Stock_Toal

Next ws
End Sub

