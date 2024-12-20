 I’m @HUGO-ROCHA-MONDRAGON, student at Université Paris Dauphine-PSL (Financial Engineering).

 I will upload my projects here.

Sub AverageByTicker()
    Dim ws As Worksheet, wsResult As Worksheet
    Dim tickerDict As Object
    Dim lastRow As Long, i As Long
    Dim ticker As String
    Dim sumDict As Object, countDict As Object

    ' Initialize
    Set ws = ThisWorkbook.Sheets("ABS-AUTO") ' Update with your sheet name
    Set wsResult = ThisWorkbook.Sheets.Add
    wsResult.Name = "AverageByDeal"
    Set tickerDict = CreateObject("Scripting.Dictionary")
    Set sumDict = CreateObject("Scripting.Dictionary")
    Set countDict = CreateObject("Scripting.Dictionary")
    
    ' Find the last row in the sheet
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Loop through the data to sum AW by Ticker
    For i = 2 To lastRow ' Assuming headers are in row 1
        ticker = ws.Cells(i, 2).Value ' Ticker in column B
        If Not tickerDict.Exists(ticker) Then
            tickerDict.Add ticker, True
            sumDict.Add ticker, ws.Cells(i, "AW").Value
            countDict.Add ticker, 1
        Else
            sumDict(ticker) = sumDict(ticker) + ws.Cells(i, "AW").Value
            countDict(ticker) = countDict(ticker) + 1
        End If
    Next i
    
    ' Output the results
    wsResult.Cells(1, 1).Value = "Ticker"
    wsResult.Cells(1, 2).Value = "Average AW"
    
    Dim rowIndex As Long: rowIndex = 2
    For Each ticker In tickerDict.Keys
        wsResult.Cells(rowIndex, 1).Value = ticker
        wsResult.Cells(rowIndex, 2).Value = sumDict(ticker) / countDict(ticker)
        rowIndex = rowIndex + 1
    Next ticker
    
    MsgBox "Averages calculated and displayed in a new sheet!", vbInformation
End Sub
<!---
HUGO-ROCHA-MONDRAGON/HUGO-ROCHA-MONDRAGON is a ✨ special ✨ repository because its `README.md` (this file) appears on your GitHub profile.
You can click the Preview link to take a look at your changes.
--->
