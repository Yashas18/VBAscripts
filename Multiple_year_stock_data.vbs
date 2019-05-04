Sub getStockVolume()

Application.DisplayAlerts = False
Application.ScreenUpdating = False

Set a = Application
Set w = WorksheetFunction

VolCol = "G"
startPxCol = "C"
closePxCol = "F"

nSheets = a.Sheets.Count

resultSheetFound = False
For i = nSheets To 1 Step -1
    If Sheets(i).Name = "Result" Then
        Sheets(i).Cells.Clear
        Set ws = Sheets(i)
        resultSheetFound = True
        Exit For
    End If
Next i

If resultSheetFound = False Then
    Set ws = Sheets.Add(After:=Sheets(nSheets))
    ws.Name = "Result"
End If

ws.Range("A1") = "Year"
ws.Range("B1") = "Ticker"
ws.Range("C1") = "SheetNo"




nSheets = a.Sheets.Count
For i = 1 To nSheets - 1
    
    Sheets(i).Columns("J:K").ClearContents
    
    Sheets(i).Range("J1:J" & w.CountA(Sheets(i).Columns("A"))).Value2 = _
    Sheets(i).Range("B1:B" & w.CountA(Sheets(i).Columns("A"))).Value2
    
    Sheets(i).Columns("J").TextToColumns Destination:=Range("J1"), DataType:=xlFixedWidth, _
            FieldInfo:=Array(Array(0, 1), Array(4, 1)), TrailingMinusNumbers:=True
    
    Sheets(i).Range("J1") = "Year"
    Sheets(i).Range("K1") = "Month/Day"
    
    'copy the tickers to result sheet
    numOfOccupiedRowsinResultSheet = w.CountA(ws.Columns("A"))
    
    ws.Range("B" & (numOfOccupiedRowsinResultSheet + 1) & ":B" & (numOfOccupiedRowsinResultSheet + w.CountA(Sheets(i).Columns("A")) - 1)).Value2 = _
    Sheets(i).Range("A2:A" & w.CountA(Sheets(i).Columns("A"))).Value2
    
    ws.Range("A" & (numOfOccupiedRowsinResultSheet + 1) & ":A" & (numOfOccupiedRowsinResultSheet + w.CountA(Sheets(i).Columns("A")) - 1)).Value2 = _
    Sheets(i).Range("J2:J" & w.CountA(Sheets(i).Columns("A"))).Value2
    
    ws.Range("C" & (numOfOccupiedRowsinResultSheet + 1) & ":C" & (numOfOccupiedRowsinResultSheet + w.CountA(Sheets(i).Columns("A")) - 1)).Value2 = _
    i
    
    numOfOccupiedRowsinResultSheet = w.CountA(ws.Columns("A"))
    
    ws.Range("A1:C" & numOfOccupiedRowsinResultSheet).RemoveDuplicates Columns:=Array(1, 2, 3), _
        Header:=xlYes
Next i




        
numOfOccupiedRowsinResultSheet = w.CountA(ws.Columns("A"))

year_arr = ws.Range("A1:A" & numOfOccupiedRowsinResultSheet).Value2
ticker_arr = ws.Range("B1:B" & numOfOccupiedRowsinResultSheet).Value2
sheetno_arr = ws.Range("C1:C" & numOfOccupiedRowsinResultSheet).Value2


ReDim volume_arr(1 To numOfOccupiedRowsinResultSheet, 1 To 1)

volume_arr(1, 1) = "Total Stock Volume"
For r = 2 To numOfOccupiedRowsinResultSheet
    volume_arr(r, 1) = a.SumIfs(Sheets(sheetno_arr(r, 1)).Columns(VolCol), _
                                    Sheets(sheetno_arr(r, 1)).Columns("A"), ticker_arr(r, 1), _
                                    Sheets(sheetno_arr(r, 1)).Columns("J"), year_arr(r, 1))
                                    
                                    
Next r

'sort all the sheets to get the min dates first
For i = 1 To nSheets - 1
    totalRowCount = w.CountA(Sheets(i).Columns("A"))
    Sheets(i).Range("A1:K" & totalRowCount).Sort _
    Key1:=Sheets(i).Range("A1"), Order1:=xlAscending, _
    Key2:=Sheets(i).Range("K1"), Order2:=xlAscending, _
    Header:=xlYes
Next i


ReDim startPx_arr(1 To numOfOccupiedRowsinResultSheet, 1 To 1)
startPx_arr(1, 1) = "OpenPx"
For r = 2 To numOfOccupiedRowsinResultSheet
    startPx_arr(r, 1) = a.Index(Sheets(sheetno_arr(r, 1)).Columns(startPxCol), a.Match(ticker_arr(r, 1), Sheets(sheetno_arr(r, 1)).Columns("A"), 0))
Next r


For i = 1 To nSheets - 1
    totalRowCount = w.CountA(Sheets(i).Columns("A"))
    Sheets(i).Range("A1:K" & totalRowCount).Sort _
    Key1:=Sheets(i).Range("A1"), Order1:=xlAscending, _
    Key2:=Sheets(i).Range("K1"), Order2:=xlDescending, _
    Header:=xlYes
Next i


ReDim closePx_arr(1 To numOfOccupiedRowsinResultSheet, 1 To 1)
ReDim yrChangePx_arr(1 To numOfOccupiedRowsinResultSheet, 1 To 1)
ReDim percChangePx_arr(1 To numOfOccupiedRowsinResultSheet, 1 To 1)

closePx_arr(1, 1) = "ClosePx"
yrChangePx_arr(1, 1) = "Yearly Change"
percChangePx_arr(1, 1) = "Percentage Change"

For r = 2 To numOfOccupiedRowsinResultSheet
    closePx_arr(r, 1) = a.Index(Sheets(sheetno_arr(r, 1)).Columns(closePxCol), a.Match(ticker_arr(r, 1), Sheets(sheetno_arr(r, 1)).Columns("A"), 0))
    yrChangePx_arr(r, 1) = closePx_arr(r, 1) - startPx_arr(r, 1)
    
    If startPx_arr(r, 1) <> 0 Then
        percChangePx_arr(r, 1) = closePx_arr(r, 1) / startPx_arr(r, 1) - 1
    Else
        percChangePx_arr(r, 1) = 0
    End If
Next r


Erase year_arr
Erase ticker_arr
Erase sheetno_arr
Erase startPx_arr
Erase closePx_arr

ws.Range("D1:D" & numOfOccupiedRowsinResultSheet).Value2 = yrChangePx_arr
ws.Range("E1:E" & numOfOccupiedRowsinResultSheet).Value2 = percChangePx_arr
ws.Range("F1:F" & numOfOccupiedRowsinResultSheet).Value2 = volume_arr


Erase volume_arr
Erase percChangePx_arr
Erase yrChangePx_arr

'clear temp data from each data sheet
For i = 1 To nSheets - 1
    Sheets(i).Columns("J:K").ClearContents
Next i

ws.Columns("C").Delete
ws.Columns("C").NumberFormat = "0.00000000"
ws.Columns("D").NumberFormat = "0.00%"
ws.Columns("E").NumberFormat = "#,##0"


'formatting data
ws.Range("C2:C" & numOfOccupiedRowsinResultSheet).Select
ws.Range("C2:C" & numOfOccupiedRowsinResultSheet).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 65280
        .TintAndShade = 0
    End With
    
ws.Range("C2:C" & numOfOccupiedRowsinResultSheet).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With

ws.Range("A1").Select



'finding max/min
'sort the output to ensure years are in order
totalRowCount = w.CountA(ws.Columns("A"))
ws.Range("A1:E" & totalRowCount).Sort _
    Key1:=Sheets(i).Range("A1"), Order1:=xlDescending, _
    Key2:=Sheets(i).Range("B1"), Order2:=xlAscending, _
    Header:=xlYes


start2016Row = 2
end2016Row = a.Match(2015, ws.Columns("A"), 0) - 1
start2015Row = end2016Row + 1
end2015Row = a.Match(2014, ws.Columns("A"), 0) - 1
start2014Row = end2015Row + 1
end2014Row = w.CountA(ws.Columns("A"))


ws.Range("H1") = "2016"
ws.Range("H2") = "Greatest % Increase"
ws.Range("H3") = "Greatest % Decrease"
ws.Range("H4") = "Greatest Total Volume"
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Value"
ws.Range("J2") = a.Max(ws.Range("D" & start2016Row & ":D" & end2016Row))
ws.Range("J3") = a.Min(ws.Range("D" & start2016Row & ":D" & end2016Row))
ws.Range("J4") = a.Max(ws.Range("E" & start2016Row & ":E" & end2016Row))
ws.Range("I2") = a.Index(ws.Range("B" & start2016Row & ":B" & end2016Row), _
                                a.Match(a.Max(ws.Range("D" & start2016Row & ":D" & end2016Row)), ws.Range("D" & start2016Row & ":D" & end2016Row), 0))
ws.Range("I3") = a.Index(ws.Range("B" & start2016Row & ":B" & end2016Row), _
                                a.Match(a.Min(ws.Range("D" & start2016Row & ":D" & end2016Row)), ws.Range("D" & start2016Row & ":D" & end2016Row), 0))
ws.Range("I4") = a.Index(ws.Range("B" & start2016Row & ":B" & end2016Row), _
                                a.Match(a.Max(ws.Range("E" & start2016Row & ":E" & end2016Row)), ws.Range("E" & start2016Row & ":E" & end2016Row), 0))

ws.Range("J2:J3").NumberFormat = "0.00%"
ws.Range("J4").NumberFormat = "#,##0"



ws.Range("H7") = "2015"
ws.Range("H8") = "Greatest % Increase"
ws.Range("H9") = "Greatest % Decrease"
ws.Range("H10") = "Greatest Total Volume"
ws.Range("I7") = "Ticker"
ws.Range("J7") = "Value"
ws.Range("J8") = a.Max(ws.Range("D" & start2015Row & ":D" & end2015Row))
ws.Range("J9") = a.Min(ws.Range("D" & start2015Row & ":D" & end2015Row))
ws.Range("J10") = a.Max(ws.Range("E" & start2015Row & ":E" & end2015Row))
ws.Range("I8") = a.Index(ws.Range("B" & start2015Row & ":B" & end2015Row), _
                                a.Match(a.Max(ws.Range("D" & start2015Row & ":D" & end2015Row)), ws.Range("D" & start2015Row & ":D" & end2015Row), 0))
ws.Range("I9") = a.Index(ws.Range("B" & start2015Row & ":B" & end2015Row), _
                                a.Match(a.Min(ws.Range("D" & start2015Row & ":D" & end2015Row)), ws.Range("D" & start2015Row & ":D" & end2015Row), 0))
ws.Range("I10") = a.Index(ws.Range("B" & start2015Row & ":B" & end2015Row), _
                                a.Match(a.Max(ws.Range("E" & start2015Row & ":E" & end2015Row)), ws.Range("E" & start2015Row & ":E" & end2015Row), 0))

ws.Range("J8:J9").NumberFormat = "0.00%"
ws.Range("J10").NumberFormat = "#,##0"



ws.Range("H13") = "2014"
ws.Range("H14") = "Greatest % Increase"
ws.Range("H15") = "Greatest % Decrease"
ws.Range("H16") = "Greatest Total Volume"
ws.Range("I13") = "Ticker"
ws.Range("J13") = "Value"
ws.Range("J14") = a.Max(ws.Range("D" & start2014Row & ":D" & end2014Row))
ws.Range("J15") = a.Min(ws.Range("D" & start2014Row & ":D" & end2014Row))
ws.Range("J16") = a.Max(ws.Range("E" & start2014Row & ":E" & end2014Row))
ws.Range("I14") = a.Index(ws.Range("B" & start2014Row & ":B" & end2014Row), _
                                a.Match(a.Max(ws.Range("D" & start2014Row & ":D" & end2014Row)), ws.Range("D" & start2014Row & ":D" & end2014Row), 0))
ws.Range("I15") = a.Index(ws.Range("B" & start2014Row & ":B" & end2014Row), _
                                a.Match(a.Min(ws.Range("D" & start2014Row & ":D" & end2014Row)), ws.Range("D" & start2014Row & ":D" & end2014Row), 0))
ws.Range("I16") = a.Index(ws.Range("B" & start2014Row & ":B" & end2014Row), _
                                a.Match(a.Max(ws.Range("E" & start2014Row & ":E" & end2014Row)), ws.Range("E" & start2014Row & ":E" & end2014Row), 0))


ws.Range("J14:J15").NumberFormat = "0.00%"
ws.Range("J16").NumberFormat = "#,##0"







ws.Cells.EntireColumn.AutoFit

MsgBox "Successfully Completed!"

Application.DisplayAlerts = True
Application.ScreenUpdating = True


End Sub


