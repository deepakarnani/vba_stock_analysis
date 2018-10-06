Sub ticker_sums()


Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet

For Each ws In ThisWorkbook.Worksheets
    ws.Activate
  ' Set an initial variable

  Dim ticker As String



  ' Set an initial variable for holding the total

  Dim ticker_total, open_stock, close_stock, yearly_change, percent_change As Double
  Dim greatperinc, greatperdec, greattotvol, greatper_range, greattolvol_range As Double
  ticker_total = 0



  ' Keep track of the location in the summary table

  Dim Summary_Table_Row As Integer
  Dim initial As Collection
  

  ' Loop through all worksheets
  

        ws.Cells(1, 9).Value = "Initial"
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Cells(1, 13).Value = "Total Stock Volume"
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(4, 16).Value = "Greatest Total Volume"
        
        Summary_Table_Row = 2
 ' Determine the Last Row

        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 ' Percent Max and min
 

        
  For i = 2 To LastRow

        Set initial = New Collection
        If open_stock = 0 Then
        open_stock = ws.Cells(i, 3).Value
        initial.Add open_stock
        End If
    ' Check if we are still within the same ticker label, if it is not...

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        close_stock = Cells(i, 6).Value
        
        
         yearly_change = close_stock - open_stock
         
        
        
       

      ' Set the ticker labels

      ticker = Cells(i, 1).Value

        

      ' Add to the ticker Total

      ticker_total = ticker_total + Cells(i, 7).Value



      ' Print the ticker label in the Summary Table

      Range("J" & Summary_Table_Row).Value = ticker



      ' Print the ticker volume to the Summary Table

      Range("M" & Summary_Table_Row).Value = ticker_total

    ' Print the yearly change to the Summary Table
    Range("I" & Summary_Table_Row).Value = open_stock
    Range("K" & Summary_Table_Row).Value = yearly_change

 If Range("I" & Summary_Table_Row).Value = 0 Then ' if denominator equals 0 then division by 0 occurs
        Cells(i, 12).Value = Null
Else
    percent_change = Abs(yearly_change) / Abs(open_stock)
    Cells(i, 12).Value = percent_change
    ' Range("L" & Summary_Table_Row).Style = "percent"
End If



' Print the percent change to the Summary Table
    
    Range("L" & Summary_Table_Row).Value = percent_change
    

      ' Add one to the summary table row

      Summary_Table_Row = Summary_Table_Row + 1

      Set initial = New Collection
      
      




      ' Reset the ticker Total

      ticker_total = 0
      open_stock = 0
      close_stock = 0

    ' If the cell immediately following a row is the same ticker...

    Else
         
         ' Add to the ticker Total
     ticker_total = ticker_total + Cells(i, 7).Value

    End If
    Next i
    
    ws.Range("L:L").NumberFormat = "0.00%"
    ws.Range("R2:R3").NumberFormat = "0.00%"

    greattotvol = Application.WorksheetFunction.Max(Range("M:M"))
    greatperinc = Application.WorksheetFunction.Max(Range("L:L"))
    greatperdec = Application.WorksheetFunction.Min(Range("L:L"))

    

     LastColor = ws.Range("J" & Rows.Count).End(xlUp).Row

    
For i = 2 To LastColor
    If ws.Cells(i, 13).Value = greattotvol Then
        ws.Range("Q4").Value = ws.Cells(i, 10).Value
        ws.Range("R4").Value = ws.Cells(i, 13).Value
    End If
    
     If ws.Cells(i, 12).Value = greatperinc Then
        ws.Range("Q2").Value = ws.Cells(i, 10).Value
        ws.Range("R2").Value = ws.Cells(i, 12).Value
       
    End If
      

    If ws.Cells(i, 12).Value = greatperdec Then
        ws.Range("Q3").Value = ws.Cells(i, 10).Value
        ws.Range("R3").Value = ws.Cells(i, 12).Value
        
    End If



    If ws.Cells(i, 11) <= 0 Then
        ws.Cells(i, 11).Interior.ColorIndex = 3
    ElseIf Cells(i, 11) > 0 Then
        ws.Cells(i, 11).Interior.ColorIndex = 4
    End If
Next i





    MsgBox ("Fixes Complete")
'Next starting_ws.Activate 'activate the worksheet that was originally active
Next

End Sub



