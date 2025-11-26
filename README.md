Sub Build_AB_C_FULL()

    Dim wsOUT As Worksheet, wsRECAP As Worksheet
    Dim wsAuto As Worksheet, wsRMBS As Worksheet
    Dim dictOut As Object, dictC As Object, dictC_lookup As Object
    Dim arrSource As Variant, arrSheets As Variant
    Dim sheetName As Variant, sheetName2 As Variant
    Dim sh As Worksheet, wsSrc As Worksheet
    Dim lastrow As Long, lastA As Long, lastC As Long
    Dim rowOUT As Long, recapRow As Long
    Dim rvis As Range, cell As Range
    Dim key As Variant
    Dim ticker As String, desc As String
    
    ' === Dictionaries ===
    Set dictOut = CreateObject("Scripting.Dictionary")
    dictOut.CompareMode = 1 ' case-insensitive
    
    Set dictC = CreateObject("Scripting.Dictionary")
    dictC.CompareMode = 1
    
    ' === Source sheets ===
    arrSource = Array("AAA_SSA", "BBB_SSA", "CCC_SSA")   ' <-- change names if needed
    
    Set wsAuto = ThisWorkbook.Worksheets("Auto")
    Set wsRMBS = ThisWorkbook.Worksheets("RMBS")
    
    ' === Output sheet RESULT ===
    Set wsOUT = ThisWorkbook.Worksheets.Add
    wsOUT.Name = "RESULT"
    
    wsOUT.Range("A1").Value = "Ticker"
    wsOUT.Range("B1").Value = "Description"
    wsOUT.Range("C1").Value = "Imported Tickers (AUTO + RMBS)"
    
    rowOUT = 2
    
    ' --------------------------------------------------------------------
    ' 1) LOOP AAA_SSA, BBB_SSA, CCC_SSA
    '    FILTER → TICKER (B) + DESC (F) → DEDUP → RESULT A/B
    ' --------------------------------------------------------------------
    
    For Each sheetName In arrSource
        
        Set sh = ThisWorkbook.Worksheets(CStr(sheetName))
        
        If sh.AutoFilterMode Then sh.AutoFilterMode = False
        
        lastrow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
        
        ' 4 critères sur la colonne 6
        sh.Range("A1").AutoFilter Field:=6, _
            Criteria1:=Array("Auto Loans", "Auto Leases", "Prime RMBS", "NC RMBS"), _
            Operator:=xlFilterValues
        
        ' Colonne B visible = tickers
        On Error Resume Next
        Set rvis = sh.Range("B2:B" & lastrow).SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        
        If Not rvis Is Nothing Then
            For Each cell In rvis
                ticker = Trim(cell.Value)
                If ticker <> "" Then
                    desc = Trim(sh.Cells(cell.Row, 6).Value)   ' colonne F
                    key = ticker & "|" & desc
                    If Not dictOut.Exists(key) Then
                        dictOut.Add key, desc
                    End If
                End If
            Next cell
        End If
        
        If sh.AutoFilterMode Then sh.AutoFilterMode = False
        
        Set rvis = Nothing
        
    Next sheetName
    
    ' Dump dictionnaire -> RESULT A/B
    For Each key In dictOut.Keys
        wsOUT.Cells(rowOUT, 1).Value = Split(key, "|")(0)
        wsOUT.Cells(rowOUT, 2).Value = dictOut(key)
        rowOUT = rowOUT + 1
    Next key
    
    ' --------------------------------------------------------------------
    ' 3) BUILD COLUMN C = DISTINCT TICKERS FROM AUTO + RMBS
    ' --------------------------------------------------------------------
    
    arrSheets = Array("Auto", "RMBS")      ' on boucle sur les NOMS
    
    For Each sheetName2 In arrSheets
        
        Set wsSrc = ThisWorkbook.Worksheets(CStr(sheetName2))
        
        If wsSrc.AutoFilterMode Then wsSrc.AutoFilterMode = False
        
        lastrow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
        
        For lastA = 2 To lastrow
            ticker = Trim(wsSrc.Cells(lastA, 1).Value)
            If ticker <> "" Then
                dictC(ticker) = 1
            End If
        Next lastA
        
    Next sheetName2
    
    ' Écrire C
    Dim writeRow As Long
    writeRow = 2
    For Each key In dictC.Keys
        wsOUT.Cells(writeRow, 3).Value = key
        writeRow = writeRow + 1
    Next key
    
    ' --------------------------------------------------------------------
    ' 4) RECAP = Tickers en A mais PAS en C
    ' --------------------------------------------------------------------
    
    Set wsRECAP = ThisWorkbook.Worksheets.Add
    wsRECAP.Name = "RECAP"
    
    wsRECAP.Range("A1").Value = "Missing Ticker"
    wsRECAP.Range("B1").Value = "Description"
    
    Set dictC_lookup = CreateObject("Scripting.Dictionary")
    dictC_lookup.CompareMode = 1
    
    ' Charger C dans dictC_lookup
    lastC = wsOUT.Cells(wsOUT.Rows.Count, 3).End(xlUp).Row
    For lastA = 2 To lastC
        ticker = Trim(wsOUT.Cells(lastA, 3).Value)
        If ticker <> "" Then
            dictC_lookup(ticker) = 1
        End If
    Next lastA
    
    ' Comparer A vs C
    recapRow = 2
    lastA = wsOUT.Cells(wsOUT.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastA
        ticker = Trim(wsOUT.Cells(i, 1).Value)
        If ticker <> "" Then
            If Not dictC_lookup.Exists(ticker) Then
                wsRECAP.Cells(recapRow, 1).Value = ticker
                wsRECAP.Cells(recapRow, 2).Value = wsOUT.Cells(i, 2).Value
                recapRow = recapRow + 1
            End If
        End If
    Next i
    
    MsgBox "Full process completed (AAA + BBB + CCC → RESULT + RECAP).", vbInformation

End Sub