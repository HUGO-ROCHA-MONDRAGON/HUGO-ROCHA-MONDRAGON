Sub Build_AB_C_FULL()

    Dim wsOUT As Worksheet, wsRECAP As Worksheet
    Dim wsAuto As Worksheet, wsRMBS As Worksheet
    Dim dictOut As Object, dictC As Object
    Dim arrSource As Variant, sh As Worksheet
    Dim lastrow As Long, lastOUT As Long
    Dim rng As Range, rvis As Range
    Dim i As Long, rowOUT As Long
    Dim key As Variant, ticker As String
    
    ' === Dictionary to store unique (Ticker|Description) pairs ===
    Set dictOut = CreateObject("Scripting.Dictionary")
    dictOut.CompareMode = 1 ' case-insensitive
    
    ' === Source sheets ===
    arrSource = Array("AAA_SSA", "BBB_SSA", "CCC_SSA") ' EDIT NAMES IF NEEDED
    
    Set wsAuto = ThisWorkbook.Worksheets("Auto")
    Set wsRMBS = ThisWorkbook.Worksheets("RMBS")
    
    ' === Output sheet ===
    Set wsOUT = ThisWorkbook.Worksheets.Add
    wsOUT.Name = "RESULT"
    
    wsOUT.Range("A1").Value = "Ticker"
    wsOUT.Range("B1").Value = "Description"
    wsOUT.Range("C1").Value = "Imported Tickers (AUTO + RMBS)"
    
    rowOUT = 2
    
    ' --------------------------------------------------------------------
    ' 1) LOOP THROUGH AAA_SSA, BBB_SSA, CCC_SSA
    '    FILTER → GET TICKER + DESCRIPTION → DEDUP → WRITE TO RESULT
    ' --------------------------------------------------------------------
    
    For Each shName In arrSource
        
        Set sh = ThisWorkbook.Worksheets(shName)
        
        ' Remove filters
        If sh.AutoFilterMode Then sh.AutoFilterMode = False
        
        lastrow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
        
        ' Apply 4-criteria filter
        sh.Range("A1").AutoFilter Field:=6, _
            Criteria1:=Array("Auto Loans", "Auto Leases", "Prime RMBS", "NC RMBS"), _
            Operator:=xlFilterValues
        
        ' Ticker column (B)
        On Error Resume Next
        Set rvis = sh.Range("B2:B" & lastrow).SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        
        If Not rvis Is Nothing Then
            
            For Each cell In rvis
                ticker = Trim(cell.Value)
                If ticker <> "" Then
                    desc = Trim(sh.Cells(cell.Row, 6).Value) ' Column F
                    key = ticker & "|" & desc
                    
                    If Not dictOut.Exists(key) Then
                        dictOut.Add key, desc
                    End If
                End If
            Next cell
            
        End If
        
        ' remove filter
        If sh.AutoFilterMode Then sh.AutoFilterMode = False
        
    Next shName
    
    ' --------------------------------------------------------------------
    ' 2) Dump deduplicated A + B into RESULT
    ' --------------------------------------------------------------------
    
    For Each key In dictOut.Keys
        wsOUT.Cells(rowOUT, 1).Value = Split(key, "|")(0)
        wsOUT.Cells(rowOUT, 2).Value = dictOut(key)
        rowOUT = rowOUT + 1
    Next key
    
    ' --------------------------------------------------------------------
    ' 3) BUILD COLUMN C = DISTINCT TICKERS FROM AUTO & RMBS
    ' --------------------------------------------------------------------
    
    Set dictC = CreateObject("Scripting.Dictionary")
    dictC.CompareMode = 1
    
    Dim arrSheets As Variant
    arrSheets = Array(wsAuto, wsRMBS)
    
    For Each sh In arrSheets
        
        If sh.AutoFilterMode Then sh.AutoFilterMode = False
        
        lastrow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastrow
            ticker = Trim(sh.Cells(i, 1).Value)
            If ticker <> "" Then dictC(ticker) = 1
        Next i
        
    Next sh
    
    ' Dump into column C
    Dim writeRow As Long: writeRow = 2
    
    For Each key In dictC.Keys
        wsOUT.Cells(writeRow, 3).Value = key
        writeRow = writeRow + 1
    Next key
    
    ' --------------------------------------------------------------------
    ' 4) CREATE RECAP = Tickers in A but NOT in C
    ' --------------------------------------------------------------------
    
    Set wsRECAP = ThisWorkbook.Worksheets.Add
    wsRECAP.Name = "RECAP"
    
    wsRECAP.Range("A1").Value = "Missing Ticker"
    wsRECAP.Range("B1").Value = "Description"
    
    Dim dictC_lookup As Object
    Set dictC_lookup = CreateObject("Scripting.Dictionary")
    dictC_lookup.CompareMode = 1
    
    ' load C into lookup
    lastC = wsOUT.Cells(wsOUT.Rows.Count, 3).End(xlUp).Row
    For i = 2 To lastC
        If wsOUT.Cells(i, 3).Value <> "" Then
            dictC_lookup(wsOUT.Cells(i, 3).Value) = 1
        End If
    Next i
    
    ' compare A with C
    Dim recapRow As Long: recapRow = 2
    lastA = wsOUT.Cells(wsOUT.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastA
        ticker = wsOUT.Cells(i, 1).Value
        
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