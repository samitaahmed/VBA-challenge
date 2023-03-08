Attribute VB_Name = "Module1"
Sub uniqueticker():
    
    
    
    Dim mainws As Worksheet
   
'to openup all the worksheet in the workbook through a loop
   
For Each mainws In ThisWorkbook.Worksheets
      
      
    Dim Rng As Range
    Dim datarange As Range
    Dim LastRow As Long
    
    Set datarange = mainws.Range("A1")
    LastRow = datarange.End(xlDown).Row
    
       
    Dim Totaltradevolume As Double
           
    Set Rng = mainws.Range("I1") 'Starting point to populate the unique ticker values
    
    mainws.Range("I:R").Clear 'Clear the columns
    
    
    mainws.Range("A:A").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Rng, Unique:=True ' unique ticker list from Coliumn A
    mainws.Cells(1, 9) = "Ticker"
    mainws.Cells(1, 10) = "Yearly Change"
    mainws.Cells(1, 11) = "Percent change"
    mainws.Cells(1, 12) = "Total stock volume"
       
    
   
    Dim maxdate As Long
    Dim mindate As Long
    
    'to get the minimum and maximum date of the year we have to convert the values in column B from text to general
    
    With mainws.Range("B2:B" & LastRow)
        .NumberFormat = "General"
        .Value = .Value
    
    End With
    maxdate = WorksheetFunction.max(mainws.Range("B2:B" & LastRow))
    mindate = WorksheetFunction.Min(mainws.Range("B2:B" & LastRow))
    
   ' Dim lastrowresult As Integer
 
    
    lastrowresult = Rng.End(xlDown).Row 'get the last row number of the unique ticker list to populate other corresponding information
    mainws.Range("J2:K" & lastrowresult).FormatConditions.Delete
    Dim resultarray() As Variant
    ReDim resultarray(1 To lastrowresult - 1, 1 To 4)
    
    resultarray = mainws.Range("I2:L" & lastrowresult)
   
    
   
      
    
    For i = 1 To lastrowresult - 1
        resultarray(i, 4) = WorksheetFunction.SumIf(mainws.Range("A2:A" & LastRow), mainws.Cells(i + 1, 9), mainws.Range("G2:G" & LastRow)) 'Total Trade volume of a ticker int he given year
        resultarray(i, 2) = WorksheetFunction.MaxIfs(mainws.Range("F2:F" & LastRow), mainws.Range("A2:A" & LastRow), mainws.Cells(i + 1, 9), mainws.Range("B2:B" & LastRow), "=" & maxdate&) - WorksheetFunction.MaxIfs(mainws.Range("C2:C" & LastRow), mainws.Range("A2:A" & LastRow), mainws.Cells(i + 1, 9), mainws.Range("B2:B" & LastRow), "=" & mindate&)
        resultarray(i, 3) = resultarray(i, 2) / WorksheetFunction.MaxIfs(mainws.Range("C2:C" & LastRow), mainws.Range("A2:A" & LastRow), mainws.Cells(i + 1, 9), mainws.Range("B2:B" & LastRow), "=" & mindate&)
                
    Next i
    
    
    'Changing the number format of the percentage change column to percentage with 2 decimel
    mainws.Range("I2:L" & lastrowresult).Value = resultarray
    mainws.Range("K2:K" & lastrowresult).NumberFormat = "0.00%"
 
 'Finding the greatest increase, decrease and greatest trade volume row number and corresponding ticker code
    Dim greatestincreaserow As Long
    Dim greatestdecreaserow As Long
    Dim greatesttradevolumerow As Long
    
    mainws.Cells(1, 17) = "Ticker"
    mainws.Cells(1, 18) = "Value"
    mainws.Cells(2, 16) = "Greatest % Increase"
    mainws.Cells(3, 16) = "Greatest % Decrease"
    mainws.Cells(4, 16) = "Greatest Total Volume"
    
  
    greatestincreaserow = WorksheetFunction.Match(WorksheetFunction.max(mainws.Range("K:K")), mainws.Range("K:K"), 0)
    
    mainws.Cells(2, 17) = resultarray(greatestincreaserow - 1, 1)
    mainws.Cells(2, 18) = resultarray(greatestincreaserow - 1, 3)
    
    greatestdecreaserow = WorksheetFunction.Match(WorksheetFunction.Min(mainws.Range("K:K")), mainws.Range("K:K"), 0)
    
    mainws.Cells(3, 17) = resultarray(greatestdecreaserow - 1, 1)
    mainws.Cells(3, 18) = resultarray(greatestdecreaserow - 1, 3)
    
    mainws.Range("R2:R3").NumberFormat = "0.00%"
    greatesttradevolumerow = WorksheetFunction.Match(WorksheetFunction.max(mainws.Range("L:L")), mainws.Range("L:L"), 0)
    
    mainws.Cells(4, 17) = resultarray(greatesttradevolumerow - 1, 1)
    mainws.Cells(4, 18) = resultarray(greatesttradevolumerow - 1, 4)
    
    ' Autofit the column to display the full text
    Columns("I:R").AutoFit
    
   mainws.Range("J2:K" & lastrowresult).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
   mainws.Range("J2:K" & lastrowresult).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
   mainws.Range("J2:K" & lastrowresult).FormatConditions(1).Interior.Color = vbGreen
   mainws.Range("J2:K" & lastrowresult).FormatConditions(2).Interior.Color = vbRed
   
   
   
   Erase resultarray()
   
    
    
   Next mainws

   
   
    
End Sub
