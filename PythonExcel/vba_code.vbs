Sub wall_street():

    For Each ws In Worksheets
    
        Dim tName, percent As String
        Dim vTotal As LongLong
        Dim oPrice, ePrice, yChange As Double
        
        vTotal = 0

        ' Determine the Last Row
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        
        'Adding new columns name
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'autofit columns
        ws.Range("J1").EntireColumn.AutoFit
        ws.Range("K1").EntireColumn.AutoFit
        ws.Range("L1").EntireColumn.AutoFit
        

        Dim tRow As Integer
        tRow = 2
        For i = 2 To lastRow
           
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            ' Ticker column
            tName = ws.Cells(i, 1).Value
            
            'Total Stock Volume column
            vTotal = vTotal + ws.Cells(i, 7).Value

           'getting closing price at the end of year
            ePrice = ws.Cells(i, 6).Value
           
           'Yearly change
           yChange = ePrice - oPrice
           
           'Change in percent
           If yChange = 0 Or oPrice = 0 Then
                percent = 0
            Else
                percent = yChange / oPrice
            End If
           
           'converting value to percent format
           percent = FormatPercent(percent)
            
           
            ws.Range("I" & tRow).Value = tName
            
             'adding value to yearly change column
            ws.Range("J" & tRow).Value = yChange
            
            ws.Range("K" & tRow).Value = percent
            
            
           'adding value to Total Stock Volume column
            ws.Range("L" & tRow).Value = vTotal

        
            tRow = tRow + 1
            
            vTotal = 0
            
            Else
              
              vTotal = vTotal + ws.Cells(i, 7).Value
              
                If ws.Cells(i, 2).Value Like "????0101" Then
                 'getting opening price at the begning of year
                   oPrice = ws.Cells(i, 3).Value
                End If
          
            End If
         
         Next i

         
        'conditional formating
        Dim MyRange, MyRange1 As Range
        Set MyRange = ws.Range("J2:J" & lastRow)
        Set MyRange1 = ws.Range("K2:K" & lastRow)
        
        'Delete previous conditional formats
        MyRange.FormatConditions.Delete
        MyRange1.FormatConditions.Delete
        'Add first rule
        MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
                Formula1:="=0"
        MyRange.FormatConditions(1).Interior.Color = vbGreen
        
        MyRange1.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
                Formula1:="=0"
        MyRange1.FormatConditions(1).Interior.Color = vbBlue
        
        'Add second rule
        MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                Formula1:="=0"
        MyRange.FormatConditions(2).Interior.Color = RGB(255, 0, 0)
        
        MyRange1.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                Formula1:="=0"
        MyRange1.FormatConditions(2).Interior.Color = RGB(255, 0, 0)
        
        
        'Answer of bonus question
        
        Dim tRange, ycRange, pRange, tvRange As Range
         Set tRange = ws.Range("I2:I" & lastRow)
         Set ycRange = ws.Range("J2:J" & lastRow)
         Set pRange = ws.Range("K2:K" & lastRow)
         Set tvRange = ws.Range("L2:L" & lastRow)
         
        'Adding new columns and rows name
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Value"
        ws.Cells(2, 16).Value = "Greatest % increase"
        ws.Cells(3, 16).Value = "Greatest % decrease"
        ws.Cells(4, 16).Value = "Greatest  Total Volume"
        
        'autofit columns
        ws.Range("P1").EntireColumn.AutoFit
        ws.Range("R1").EntireColumn.AutoFit

        
        'extracting max and min value
        mTvolume = WorksheetFunction.Max(tvRange)
        pIncrease = WorksheetFunction.Max(pRange)
        pDecrease = WorksheetFunction.Min(pRange)
        
        ws.Cells(2, 18).Value = pIncrease
        ws.Cells(2, 18).NumberFormat = "0.00%"
         
        ws.Cells(3, 18).Value = pDecrease
        ws.Cells(3, 18).NumberFormat = "0.00%"
        
        ws.Cells(4, 18).Value = mTvolume
        
        'ticker name with max percent increase
        Dim pitName, pdtName, mtName As String
        For Each p In pRange
         If p.Value = pIncrease Then
         pitName = p.Offset(, -2).Value
         ws.Cells(2, 17).Value = pitName
         End If
         
         'ticker name with max percent decrease
         If p.Value = pDecrease Then
         pdtName = p.Offset(, -2).Value
         ws.Cells(3, 17).Value = pdtName
         End If
         Next p
        
        'ticker name with greatest total volume
         For Each v In tvRange
         If v.Value = mTvolume Then
         mtName = v.Offset(, -3).Value
         ws.Cells(4, 17).Value = mtName
         End If
         Next v
 
        Next ws
    
End Sub