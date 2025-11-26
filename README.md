 ' --------------------------------------------------------------------
'   4) TABLE RECAP : Tickers en A mais PAS en C
' --------------------------------------------------------------------

Dim wsRECAP As Worksheet
Dim dictC As Object
Dim lastA As Long, lastC As Long
Dim i As Long, tickerA As String

Set wsRECAP = ThisWorkbook.Worksheets.Add
wsRECAP.Name = "RECAP"

wsRECAP.Range("A1").Value = "Missing Ticker"
wsRECAP.Range("B1").Value = "Description"

' Charger les tickers de la colonne C dans un dictionnaire pour lookup rapide
Set dictC = CreateObject("Scripting.Dictionary")

lastC = wsOUT.Cells(wsOUT.Rows.Count, 3).End(xlUp).Row
For i = 2 To lastC
    If wsOUT.Cells(i, 3).Value <> "" Then
        dictC(wsOUT.Cells(i, 3).Value) = 1
    End If
Next i

' Comparer colonne A avec dictionnaire C
lastA = wsOUT.Cells(wsOUT.Rows.Count, 1).End(xlUp).Row

Dim recapRow As Long: recapRow = 2

For i = 2 To lastA
    tickerA = wsOUT.Cells(i, 1).Value
    
    If tickerA <> "" Then
        If Not dictC.Exists(tickerA) Then
            wsRECAP.Cells(recapRow, 1).Value = tickerA
            wsRECAP.Cells(recapRow, 2).Value = wsOUT.Cells(i, 2).Value   ' description
            recapRow = recapRow + 1
        End If
    End If
Next i